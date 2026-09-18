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
    text = atob(body.content.replace(/\n/g, ''));
  } else {
    const rawRes = await fetch(`https://raw.githubusercontent.com/${OWNER}/${REPO}/${BRANCH}/${p}?_cb=${Date.now()}`, {
      headers: { 'User-Agent': 'lph-materials-invoice-worker' },
    });
    if (!rawRes.ok) throw new Error(`GET ${p} fallback via raw.githubusercontent.com failed: ${rawRes.status}`);
    text = await rawRes.text();
  }
  return { data: JSON.parse(text), sha: body.sha };
}
 
function b64EncodeUtf8(str) {
  return btoa(unescape(encodeURIComponent(str)));
}
 
async function ghPut(token, p, obj, sha, message) {
  const content = b64EncodeUtf8(JSON.stringify(obj));
  let lastErrText = '';
  for (let attempt = 0; attempt < 3; attempt++) {
    const res = await fetch(API + encodeURIComponent(p).replace(/%2F/g, '/'), {
      method: 'PUT',
      headers: ghHeaders(token, { 'Content-Type': 'application/json' }),
      body: JSON.stringify({ message, content, sha: sha || undefined, branch: BRANCH }),
    });
    if (res.ok) return res.json();
    lastErrText = await res.text();
    if (res.status === 409 && attempt < 2) {
      const fresh = await ghGet(token, p);
      sha = fresh.sha;
      continue;
    }
    throw new Error(`PUT ${p} failed: ${res.status} ${lastErrText}`);
  }
  throw new Error(`PUT ${p} failed after retries: ${lastErrText}`);
}
 
async function readJobs(token) {
  const { data, sha } = await ghGet(token, 'data/jobs.json');
  return { jobs: Array.isArray(data) ? data : [], sha };
}
async function readInvoiceChecks(token) {
  const { data, sha } = await ghGet(token, 'data/invoice_checks.json');
  return { records: Array.isArray(data) ? data : [], sha };
}
 
async function handleJobsUpsert(token, body) {
  const incoming = Array.isArray(body.jobs) ? body.jobs : [];
  if (!incoming.length) return json({ ok: true, upserted: 0 });
  const { jobs, sha } = await readJobs(token);
  const idx = {};
  jobs.forEach((j, i) => { idx[j.jobNumber] = i; });
  let added = 0, updated = 0;
  incoming.forEach((j) => {
    if (!j || j.jobNumber === undefined || j.jobNumber === null || j.jobNumber === '') return;
    const key = String(j.jobNumber);
    j.jobNumber = key;
    if (idx[key] !== undefined) { jobs[idx[key]] = j; updated++; }
    else { jobs.push(j); idx[key] = jobs.length - 1; added++; }
  });
  await ghPut(token, 'data/jobs.json', jobs, sha, `Materials sync: upsert ${incoming.length} job(s) (+${added}/~${updated})`);
  return json({ ok: true, added, updated, total: jobs.length });
}
 
async function handleInvoiceChecksUpsert(token, body) {
  const incoming = Array.isArray(body.records) ? body.records : [];
  if (!incoming.length) return json({ ok: true, upserted: 0 });
  const { records, sha } = await readInvoiceChecks(token);
  const idx = {};
  records.forEach((r, i) => { idx[r.id] = i; });
  let added = 0, updated = 0;
  incoming.forEach((r) => {
    if (!r || r.id === undefined || r.id === null || r.id === '') return;
    const key = String(r.id);
    r.id = key;
    if (!Array.isArray(r.flags)) r.flags = [];
    if (idx[key] !== undefined) { records[idx[key]] = r; updated++; }
    else { records.push(r); idx[key] = records.length - 1; added++; }
  });
  await ghPut(token, 'data/invoice_checks.json', records, sha, `Materials sync: upsert ${incoming.length} invoice check(s) (+${added}/~${updated})`);
  return json({ ok: true, added, updated, total: records.length });
}
 
async function handleInvoiceStatus(token, body) {
  const id = body.id !== undefined && body.id !== null ? String(body.id) : '';
  if (!id) return json({ ok: false, error: 'id is required' }, 400);
  const override = body.manualOverride === 'pass' || body.manualOverride === 'fail' ? body.manualOverride : null;
  const { records, sha } = await readInvoiceChecks(token);
  const rec = records.find((r) => String(r.id) === id);
  if (!rec) return json({ ok: false, error: 'record not found' }, 404);
  rec.manualOverride = override;
  await ghPut(token, 'data/invoice_checks.json', records, sha, `Materials sync: set status ${id} -> ${override}`);
  return json({ ok: true });
}
 
async function handleInvoiceDelete(token, body) {
  const id = body.id !== undefined && body.id !== null ? String(body.id) : '';
  if (!id) return json({ ok: false, error: 'id is required' }, 400);
  const { records, sha } = await readInvoiceChecks(token);
  const next = records.filter((r) => String(r.id) !== id);
  if (next.length === records.length) return json({ ok: true, removed: false });
  await ghPut(token, 'data/invoice_checks.json', next, sha, `Materials sync: delete invoice check ${id}`);
  return json({ ok: true, removed: true });
}
 
async function handleJobsClear(token) {
  const { sha } = await readJobs(token);
  await ghPut(token, 'data/jobs.json', [], sha, 'Materials sync: clear all jobs');
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
 

