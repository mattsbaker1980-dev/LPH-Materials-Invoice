// LPH Materials vs Invoice — Cloudflare Worker relay.
//
// Holds a GitHub fine-grained PAT server-side (as the GITHUB_TOKEN secret) and
// commits writes to mattsbaker1980-dev/LPH-Materials-Invoice on the browser's
// behalf. The browser (index.html) and the daily sync script both POST JSON
// to these routes; this worker never runs any parsing/flagging logic itself —
// it only merges the records it's handed into the two data files and commits.
//
// Routes (all POST, JSON body):
//   /api/materials-jobs           { jobs: [...] }              upsert by jobNumber into data/jobs.json
//   /api/materials-invoice-checks { records: [...] }           upsert by id into data/invoice_checks.json
//   /api/materials-invoice-status { id, manualOverride }       patch one record's manualOverride
//   /api/materials-invoice-delete { id }                       remove one record from data/invoice_checks.json
//   /api/materials-jobs-clear     {}                           reset data/jobs.json to []
//
//   /api/state, /api/state/status, /api/state/note, /api/state/log   Job Review marks + notes (Cloudflare D1, see below)
//
// Required binding: DB — a Cloudflare D1 database (Settings -> Bindings -> D1 database, variable name DB).
// Required secret: GITHUB_TOKEN — a fine-grained PAT scoped to just this repo,
// Contents: Read and write, no expiration. Set with:
//   wrangler secret put GITHUB_TOKEN
// or via the Cloudflare dashboard: Worker -> Settings -> Variables and Secrets.
 
const OWNER = 'mattsbaker1980-dev';
const REPO = 'LPH-Materials-Invoice';
const BRANCH = 'main';
const API = `https://api.github.com/repos/${OWNER}/${REPO}/contents/`;
 
function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
  };
}
 
function json(body, status) {
  return new Response(JSON.stringify(body), {
    status: status || 200,
    headers: Object.assign({ 'Content-Type': 'application/json' }, corsHeaders()),
  });
}
 
function ghHeaders(token, extra) {
  return Object.assign(
    { Authorization: 'Bearer ' + token, 'User-Agent': 'lph-materials-invoice-worker', Accept: 'application/vnd.github+json' },
    extra || {}
  );
}
 
async function ghGet(token, p) {
  const res = await fetch(API + encodeURIComponent(p).replace(/%2F/g, '/') + '?ref=' + BRANCH + '&_cb=' + Date.now(), {
    headers: ghHeaders(token),
  });
  if (res.status === 404) return { data: null, sha: null };
  if (!res.ok) throw new Error(`GET ${p} failed: ${res.status} ${await res.text()}`);
  const body = await res.json();
  let text;
  if (body.content) {
    text = new TextDecoder().decode(Uint8Array.from(atob(body.content.replace(/\n/g, '')), (c) => c.charCodeAt(0)));
  } else {
    // Files over 1 MB come back without content. Read the exact blob for this sha
    // (never raw.githubusercontent.com: that copy is CDN-cached for minutes, and
    // writing a stale copy back silently erased other recent saves).
    const blobRes = await fetch(`https://api.github.com/repos/${OWNER}/${REPO}/git/blobs/${body.sha}`, {
      headers: ghHeaders(token, { Accept: 'application/vnd.github.raw' }),
    });
    if (!blobRes.ok) throw new Error(`GET ${p} blob failed: ${blobRes.status}`);
    text = await blobRes.text();
  }
  return { data: JSON.parse(text), sha: body.sha };
}
 
function b64EncodeUtf8(str) {
  return btoa(unescape(encodeURIComponent(str)));
}
 
// Read-modify-write with retry: on a 409 (someone else committed first) re-read the
// latest file and re-apply the change, instead of re-sending a stale copy.
async function ghUpdate(token, p, mutate, message) {
  let lastErrText = '';
  for (let attempt = 0; attempt < 4; attempt++) {
    const { data, sha } = await ghGet(token, p);
    const out = mutate(data);
    const res = await fetch(API + encodeURIComponent(p).replace(/%2F/g, '/'), {
      method: 'PUT',
      headers: ghHeaders(token, { 'Content-Type': 'application/json' }),
      body: JSON.stringify({ message: typeof message === 'function' ? message(out) : message, content: b64EncodeUtf8(JSON.stringify(out.data)), sha: sha || undefined, branch: BRANCH }),
    });
    if (res.ok) return out;
    lastErrText = await res.text();
    if (res.status === 409 || res.status === 422) continue;
    throw new Error(`PUT ${p} failed: ${res.status} ${lastErrText}`);
  }
  throw new Error(`PUT ${p} failed after retries: ${lastErrText}`);
}
 
const asArray = (d) => (Array.isArray(d) ? d : []);
 
function upsertBy(list, incoming, keyField, prep) {
  const idx = {};
  list.forEach((x, i) => { idx[String(x[keyField])] = i; });
  let added = 0, updated = 0;
  incoming.forEach((x) => {
    if (!x || x[keyField] === undefined || x[keyField] === null || x[keyField] === '') return;
    const key = String(x[keyField]);
    x[keyField] = key;
    if (prep) prep(x);
    if (idx[key] !== undefined) { list[idx[key]] = x; updated++; }
    else { list.push(x); idx[key] = list.length - 1; added++; }
  });
  return { added, updated };
}
 
async function handleJobsUpsert(token, body) {
  const incoming = Array.isArray(body.jobs) ? body.jobs : [];
  if (!incoming.length) return json({ ok: true, upserted: 0 });
  const out = await ghUpdate(token, 'data/jobs.json', (d) => {
    const jobs = asArray(d); const r = upsertBy(jobs, incoming, 'jobNumber');
    return { data: jobs, added: r.added, updated: r.updated, total: jobs.length };
  }, (o) => `Materials sync: upsert ${incoming.length} job(s) (+${o.added}/~${o.updated})`);
  return json({ ok: true, added: out.added, updated: out.updated, total: out.total });
}
 
async function handleInvoiceChecksUpsert(token, body) {
  const incoming = Array.isArray(body.records) ? body.records : [];
  if (!incoming.length) return json({ ok: true, upserted: 0 });
  const out = await ghUpdate(token, 'data/invoice_checks.json', (d) => {
    const records = asArray(d);
    const r = upsertBy(records, incoming, 'id', (x) => { if (!Array.isArray(x.flags)) x.flags = []; });
    return { data: records, added: r.added, updated: r.updated, total: records.length };
  }, (o) => `Materials sync: upsert ${incoming.length} invoice check(s) (+${o.added}/~${o.updated})`);
  return json({ ok: true, added: out.added, updated: out.updated, total: out.total });
}
 
async function handleInvoiceStatus(token, body) {
  const id = body.id !== undefined && body.id !== null ? String(body.id) : '';
  if (!id) return json({ ok: false, error: 'id is required' }, 400);
  const override = body.manualOverride === 'pass' || body.manualOverride === 'fail' ? body.manualOverride : null;
  let found = true;
  await ghUpdate(token, 'data/invoice_checks.json', (d) => {
    const records = asArray(d); const rec = records.find((r) => String(r.id) === id);
    if (!rec) { found = false; throw new Error('record not found'); }
    rec.manualOverride = override; return { data: records };
  }, `Materials sync: set status ${id} -> ${override}`).catch((e) => { if (found) throw e; });
  if (!found) return json({ ok: false, error: 'record not found' }, 404);
  return json({ ok: true });
}
 
async function handleInvoiceDelete(token, body) {
  const id = body.id !== undefined && body.id !== null ? String(body.id) : '';
  if (!id) return json({ ok: false, error: 'id is required' }, 400);
  const { data } = await ghGet(token, 'data/invoice_checks.json');
  if (!asArray(data).some((r) => String(r.id) === id)) return json({ ok: true, removed: false });
  await ghUpdate(token, 'data/invoice_checks.json', (d) => ({ data: asArray(d).filter((r) => String(r.id) !== id) }), `Materials sync: delete invoice check ${id}`);
  return json({ ok: true, removed: true });
}
 
async function handleJobsClear(token) {
  await ghUpdate(token, 'data/jobs.json', () => ({ data: [] }), 'Materials sync: clear all jobs');
  return json({ ok: true });
}
 
// ---------------------------------------------------------------------------
// Job Review state (Reviewed marks + notes) lives in Cloudflare D1, bound to this
// Worker as "DB". One row per job, so saves never overwrite each other, and every
// change is also written to a change log. A copy is backed up to GitHub
// (data/review/state_backup.json) at most every 10 minutes while people are working.
//
// Routes (POST, JSON):
//   /api/state         {}                                   -> { statuses, notes, imported }
//   /api/state/status  { job, status: 'pass'|'fail'|null }  -> { ok, row }
//   /api/state/note    { job, concern?, positive?, excuse?:{code,reason}, unexcuse?:code } -> { ok, row }
//   /api/state/log     { job? }                             -> { rows } (latest 500 changes)
// ---------------------------------------------------------------------------
const BACKUP_PATH = 'data/review/state_backup.json';
const BACKUP_EVERY_MS = 10 * 60 * 1000;
let schemaReady = false;

async function ensureSchema(db) {
  if (schemaReady) return;
  await db.batch([
    db.prepare('CREATE TABLE IF NOT EXISTS job_state (job TEXT PRIMARY KEY, status TEXT, status_at INTEGER, concern TEXT, positive TEXT, excused TEXT, note_at INTEGER)'),
    db.prepare('CREATE TABLE IF NOT EXISTS change_log (id INTEGER PRIMARY KEY AUTOINCREMENT, at INTEGER, job TEXT, field TEXT, value TEXT)'),
    db.prepare('CREATE TABLE IF NOT EXISTS meta (k TEXT PRIMARY KEY, v TEXT)'),
  ]);
  schemaReady = true;
}

async function getMeta(db, k) {
  const r = await db.prepare('SELECT v FROM meta WHERE k = ?').bind(k).first();
  return r ? r.v : null;
}
async function setMeta(db, k, v) {
  await db.prepare('INSERT INTO meta (k, v) VALUES (?, ?) ON CONFLICT(k) DO UPDATE SET v = excluded.v').bind(k, String(v)).run();
}

// First run only: copy the marks and notes that were stored in data/invoice_checks.json.
// Uses DO NOTHING on conflict, so it can never overwrite a newer save.
async function importLegacy(env) {
  const db = env.DB;
  const done = await getMeta(db, 'imported');
  if (done) return JSON.parse(done);
  const { data } = await ghGet(env.GITHUB_TOKEN, 'data/invoice_checks.json');
  const rows = {};
  asArray(data).forEach((x) => {
    if (!x || x.id === undefined || x.id === null) return;
    const id = String(x.id);
    if (id.indexOf('note:') === 0) {
      const job = id.slice(5);
      const ex = x.excused && typeof x.excused === 'object' ? x.excused : {};
      if (!x.concern && !x.positive && !Object.keys(ex).length) return;
      const r = rows[job] = rows[job] || { job };
      r.concern = x.concern || ''; r.positive = x.positive || ''; r.excused = JSON.stringify(ex); r.note_at = x.noteAt || x.savedAt || Date.now();
    } else if (x.manualOverride === 'pass' || x.manualOverride === 'fail') {
      const r = rows[id] = rows[id] || { job: id };
      r.status = x.manualOverride; r.status_at = x.savedAt || Date.now();
    }
  });
  const list = Object.values(rows);
  const stmt = db.prepare('INSERT INTO job_state (job, status, status_at, concern, positive, excused, note_at) VALUES (?, ?, ?, ?, ?, ?, ?) ON CONFLICT(job) DO NOTHING');
  for (let i = 0; i < list.length; i += 50) {
    await db.batch(list.slice(i, i + 50).map((r) => stmt.bind(r.job, r.status || null, r.status_at || null, r.concern || null, r.positive || null, r.excused || null, r.note_at || null)));
  }
  const summary = { at: Date.now(), statuses: list.filter((r) => r.status).length, notes: list.filter((r) => r.note_at).length };
  await setMeta(db, 'imported', JSON.stringify(summary));
  await db.prepare('INSERT INTO change_log (at, job, field, value) VALUES (?, ?, ?, ?)').bind(Date.now(), '*', 'import', JSON.stringify(summary)).run();
  return summary;
}

function rowOut(r) {
  let excused = {};
  try { excused = r.excused ? JSON.parse(r.excused) : {}; } catch (e) { excused = {}; }
  return { job: r.job, status: r.status || null, statusAt: r.status_at || null, concern: r.concern || '', positive: r.positive || '', excused, noteAt: r.note_at || null };
}

async function readAll(db) {
  const { results } = await db.prepare('SELECT * FROM job_state').all();
  const statuses = {}, notes = {};
  (results || []).forEach((r) => {
    const o = rowOut(r);
    if (o.status === 'pass' || o.status === 'fail') statuses[o.job] = o.status;
    if (o.concern || o.positive || Object.keys(o.excused).length) notes[o.job] = { concern: o.concern, positive: o.positive, excused: o.excused, at: o.noteAt };
  });
  return { statuses, notes };
}

async function maybeBackup(env, force) {
  try {
    const db = env.DB;
    const last = +(await getMeta(db, 'last_backup')) || 0;
    const dirty = +(await getMeta(db, 'last_change')) || 0;
    if (!force && (dirty <= last || Date.now() - last < BACKUP_EVERY_MS)) return;
    // Claim the backup slot atomically so simultaneous saves don't all back up at once.
    const now = Date.now();
    await db.prepare("INSERT INTO meta (k, v) VALUES ('last_backup', '0') ON CONFLICT(k) DO NOTHING").run();
    const claim = await db.prepare("UPDATE meta SET v = ? WHERE k = 'last_backup' AND CAST(v AS INTEGER) = ?").bind(String(now), last).run();
    if (!claim || !claim.meta || claim.meta.changes !== 1) return;
    const all = await readAll(db);
    const { results } = await db.prepare('SELECT * FROM job_state ORDER BY job').all();
    const payload = { backedUpAt: new Date().toISOString(), counts: { statuses: Object.keys(all.statuses).length, notes: Object.keys(all.notes).length }, rows: (results || []).map(rowOut) };
    await ghUpdate(env.GITHUB_TOKEN, BACKUP_PATH, () => ({ data: payload }), `Job Review state backup: ${payload.counts.statuses} reviewed, ${payload.counts.notes} notes`);
  } catch (e) { /* best effort; next save or page load retries */ }
}

async function logChange(db, job, field, value) {
  await db.batch([
    db.prepare('INSERT INTO change_log (at, job, field, value) VALUES (?, ?, ?, ?)').bind(Date.now(), job, field, value === null || value === undefined ? null : String(value)),
    db.prepare("INSERT INTO meta (k, v) VALUES ('last_change', ?) ON CONFLICT(k) DO UPDATE SET v = excluded.v").bind(String(Date.now())),
  ]);
}

function jobOf(body) {
  const j = body && body.job !== undefined && body.job !== null ? String(body.job).trim() : '';
  return /^[A-Za-z0-9_-]{1,40}$/.test(j) ? j : '';
}

async function handleState(env, body, ctx) {
  const imported = await importLegacy(env);
  const all = await readAll(env.DB);
  ctx.waitUntil(maybeBackup(env, false));
  return json({ ok: true, statuses: all.statuses, notes: all.notes, imported });
}

async function handleStateStatus(env, body, ctx) {
  const job = jobOf(body);
  if (!job) return json({ ok: false, error: 'job is required' }, 400);
  const status = body.status === 'pass' || body.status === 'fail' ? body.status : null;
  await importLegacy(env);
  const now = Date.now();
  await env.DB.prepare('INSERT INTO job_state (job, status, status_at) VALUES (?, ?, ?) ON CONFLICT(job) DO UPDATE SET status = excluded.status, status_at = excluded.status_at').bind(job, status, now).run();
  await logChange(env.DB, job, 'status', status);
  const row = await env.DB.prepare('SELECT * FROM job_state WHERE job = ?').bind(job).first();
  ctx.waitUntil(maybeBackup(env, false));
  return json({ ok: true, row: row ? rowOut(row) : null });
}

async function handleStateNote(env, body, ctx) {
  const job = jobOf(body);
  if (!job) return json({ ok: false, error: 'job is required' }, 400);
  await importLegacy(env);
  const cur = await env.DB.prepare('SELECT * FROM job_state WHERE job = ?').bind(job).first();
  const c = cur ? rowOut(cur) : { concern: '', positive: '', excused: {} };
  const next = { concern: c.concern, positive: c.positive, excused: Object.assign({}, c.excused) };
  if (typeof body.concern === 'string') next.concern = body.concern.slice(0, 5000);
  if (typeof body.positive === 'string') next.positive = body.positive.slice(0, 5000);
  if (body.excuse && body.excuse.code) next.excused[String(body.excuse.code)] = String(body.excuse.reason || 'Excused').slice(0, 500);
  if (body.unexcuse) delete next.excused[String(body.unexcuse)];
  const now = Date.now();
  await env.DB.prepare('INSERT INTO job_state (job, concern, positive, excused, note_at) VALUES (?, ?, ?, ?, ?) ON CONFLICT(job) DO UPDATE SET concern = excluded.concern, positive = excluded.positive, excused = excluded.excused, note_at = excluded.note_at')
    .bind(job, next.concern, next.positive, JSON.stringify(next.excused), now).run();
  await logChange(env.DB, job, 'note', JSON.stringify(next));
  const row = await env.DB.prepare('SELECT * FROM job_state WHERE job = ?').bind(job).first();
  ctx.waitUntil(maybeBackup(env, false));
  return json({ ok: true, row: row ? rowOut(row) : null });
}

async function handleStateLog(env, body) {
  const job = jobOf(body);
  const q = job
    ? env.DB.prepare('SELECT * FROM change_log WHERE job = ? ORDER BY id DESC LIMIT 500').bind(job)
    : env.DB.prepare('SELECT * FROM change_log ORDER BY id DESC LIMIT 500');
  const { results } = await q.all();
  return json({ ok: true, rows: results || [] });
}

const STATE_ROUTES = {
  '/api/state': handleState,
  '/api/state/status': handleStateStatus,
  '/api/state/note': handleStateNote,
  '/api/state/log': handleStateLog,
};
 
const ROUTES = {
  '/api/materials-jobs': handleJobsUpsert,
  '/api/materials-invoice-checks': handleInvoiceChecksUpsert,
  '/api/materials-invoice-status': handleInvoiceStatus,
  '/api/materials-invoice-delete': handleInvoiceDelete,
};
 
export default {
  async fetch(request, env, ctx) {
    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: corsHeaders() });
    }
    if (request.method !== 'POST') {
      return json({ ok: false, error: 'POST only' }, 405);
    }
    if (!env.GITHUB_TOKEN) {
      return json({ ok: false, error: 'Worker is not configured: GITHUB_TOKEN secret is missing' }, 500);
    }
 
    const url = new URL(request.url);
 
    try {
      if (STATE_ROUTES[url.pathname]) {
        if (!env.DB) return json({ ok: false, error: 'no_db', message: 'The D1 database is not connected to this Worker yet (Settings -> Bindings -> D1 database, variable name DB).' }, 503);
        await ensureSchema(env.DB);
        let body = {};
        try { body = await request.json(); } catch (e) { body = {}; }
        return await STATE_ROUTES[url.pathname](env, body, ctx);
      }
      if (url.pathname === '/api/materials-jobs-clear') {
        return await handleJobsClear(env.GITHUB_TOKEN);
      }
      const handler = ROUTES[url.pathname];
      if (!handler) return json({ ok: false, error: 'Not found' }, 404);
      let body = {};
      try { body = await request.json(); } catch (e) { body = {}; }
      return await handler(env.GITHUB_TOKEN, body);
    } catch (e) {
      return json({ ok: false, error: String((e && e.message) || e) }, 500);
    }
  },
};
 

