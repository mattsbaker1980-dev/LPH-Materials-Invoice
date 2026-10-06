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
 
const ROUTES = {
  '/api/materials-jobs': handleJobsUpsert,
  '/api/materials-invoice-checks': handleInvoiceChecksUpsert,
  '/api/materials-invoice-status': handleInvoiceStatus,
  '/api/materials-invoice-delete': handleInvoiceDelete,
};
 
export default {
  async fetch(request, env) {
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
 

