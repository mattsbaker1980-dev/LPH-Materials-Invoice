// Claude review step for the Job Review tool.
//
// The daily scheduled Claude run uses this in two steps:
//
//   1) node review_ai.js queue [--local <repoDir>] [--max 80] [--days 10] > queue.json
//      Prints the jobs whose write-up / invoice / estimates changed since their
//      last review (or were never reviewed), in a compact form for Claude to read.
//
//   2) Claude reads queue.json and writes results.json:
//      [{ "jobNumber": "123", "fingerprint": "<copied from queue>",
//         "verdict": "ok" | "issues",
//         "issues": [{ "type": "not_billed" | "billed_not_described" | "options" | "writeup" | "other",
//                      "severity": "high" | "low",
//                      "text": "Write-up says capacitor replaced; no capacitor on invoice" }],
//         "note": "optional one-line summary",
//         // only for jobs that have a "materialsGap" block in the queue:
//         "materialsGapFor": <purchasedOnPO copied from queue>, "likelyUnrecorded": ["SDR 35 pipe", ...], "gapNote": "..." }]
//
//   3) GH_TOKEN=... node review_ai.js apply results.json [--local <repoDir>]
//      Saves the results onto each job and recomputes flags for the month.

const fs = require('fs');
const L = require('./review_sync.js');

function parseArgs() {
  const a = process.argv.slice(2);
  const o = { cmd: a[0], local: null, max: 80, days: 10, gapDays: 45, file: null, onlyGap: false };
  for (let i = 1; i < a.length; i++) {
    if (a[i] === '--local') o.local = a[++i];
    else if (a[i] === '--max') o.max = parseInt(a[++i], 10);
    else if (a[i] === '--days') o.days = parseInt(a[++i], 10);
    else if (a[i] === '--only-gap') o.onlyGap = true;
    else o.file = a[i];
  }
  return o;
}

const clip = (s, n) => { s = String(s || '').replace(/\s+\n/g, '\n').trim(); return s.length > n ? s.slice(0, n) + ' …[cut]' : s; };

let SINCE = null; // only review recent jobs (older backlog is skipped)
let GAP_SINCE = null;
let ONLY_GAP = false;
const poOf = (j) => Math.round(j.invoices.reduce((s, i) => s + (i.poCost || 0) + (i.billCost || 0), 0) * 100) / 100;
// Jobs flagged "PO not recorded as materials" get a materials-gap read even if older than the normal window,
// and again whenever their PO total changes.
function needsGapReview(job) {
  if (!(job.flags || []).some((f) => f.code === 'po_not_recorded')) return false;
  if (GAP_SINCE && (!job.completionDate || job.completionDate < GAP_SINCE)) return false;
  return !job.aiReview || job.aiReview.fingerprint !== L.reviewFingerprint(job) || job.aiReview.gapFor !== poOf(job);
}
function needsReview(job) {
  if (needsGapReview(job)) return true;
  if (ONLY_GAP) return false;
  if (SINCE && (!job.completionDate || job.completionDate < SINCE)) return false;
  const hasWriteUp = job.invoices.some((i) => (i.items || []).length || i.summary);
  if (!hasWriteUp) return false;
  return !job.aiReview || job.aiReview.fingerprint !== L.reviewFingerprint(job);
}

async function queue(o) {
  SINCE = new Date(Date.now() - o.days * 86400000).toISOString().slice(0, 10);
  GAP_SINCE = new Date(Date.now() - o.gapDays * 86400000).toISOString().slice(0, 10);
  ONLY_GAP = o.onlyGap;
  const store = L.makeStore({ local: o.local });
  const index = (await store.read(`${L.DIR}/index.json`)).data;
  if (!index) { console.log('[]'); return; }
  const out = [];
  // newest month first, newest jobs first
  for (const mo of index.months.slice(0, 2)) {
    const data = (await store.read(`${L.DIR}/${mo}.json`)).data;
    const jobs = Object.values(data.jobs).filter(needsReview)
      .sort((a, b) => (needsGapReview(b) - needsGapReview(a)) || String(b.completionDate).localeCompare(String(a.completionDate)));
    for (const j of jobs) {
      if (out.length >= o.max) break;
      out.push({
        jobNumber: j.jobNumber,
        fingerprint: L.reviewFingerprint(j),
        jobType: j.jobType, tech: j.primaryTech, completed: j.completionDate,
        bookingNotes: clip(j.materials && j.materials.bookingNotes, 400),
        invoices: j.invoices.map((i) => ({
          total: i.total, writeUp: clip(i.summary, 2500),
          items: (i.items || []).map((x) => `${x.type === 'Material' ? '[M] ' : ''}${x.name} $${x.price}`),
        })),
        estimates: j.estimates.map((e) => `${e.status}: ${e.name} $${e.subtotal}${e.summary ? ' — ' + clip(e.summary, 300) : ''}`),
        costs: { material: j.materials ? j.materials.materialCost : null, po: poOf(j) },
      });
      if (needsGapReview(j)) {
        const last = out[out.length - 1];
        last.materialsGap = {
          purchasedOnPO: poOf(j),
          recordedMaterialCost: Math.round(j.invoices.reduce((s, i) => s + (i.materialCost || 0), 0) * 100) / 100,
          recordedMaterials: j.invoices.flatMap((i) => (i.items || []).filter((x) => /material|equipment/i.test(x.type)).map((x) => x.name)),
        };
      }
    }
  }
  const remaining = await countRemaining(store, index, out.length);
  process.stdout.write(JSON.stringify({ count: out.length, remainingAfterThis: remaining, jobs: out }, null, 1));
}

async function countRemaining(store, index, taken) {
  let n = 0;
  for (const mo of index.months.slice(0, 2)) n += Object.values((await store.read(`${L.DIR}/${mo}.json`)).data.jobs).filter(needsReview).length;
  return Math.max(0, n - taken);
}

async function apply(o) {
  const results = JSON.parse(fs.readFileSync(o.file, 'utf8'));
  const list = Array.isArray(results) ? results : results.results || results.jobs;
  const store = L.makeStore({ local: o.local });
  const idx = await store.read(`${L.DIR}/index.json`);
  const months = {};
  let saved = 0, stale = 0, missing = 0;
  for (const r of list) {
    const mo = idx.data.jobMonth[r.jobNumber];
    if (!mo) { missing++; continue; }
    if (!months[mo]) months[mo] = await store.read(`${L.DIR}/${mo}.json`);
    const job = months[mo].data.jobs[r.jobNumber];
    if (!job) { missing++; continue; }
    if (r.fingerprint !== L.reviewFingerprint(job)) { stale++; continue; } // data changed after it was queued
    job.aiReview = {
      fingerprint: r.fingerprint, reviewedAt: new Date().toISOString(),
      verdict: r.verdict === 'issues' && (r.issues || []).length ? 'issues' : 'ok',
      issues: (r.issues || []).slice(0, 6).map((x) => ({ type: String(x.type || 'other').replace(/[^a-z_]/gi, ''), severity: x.severity === 'high' ? 'high' : 'low', text: String(x.text || '').slice(0, 300) })),
      note: r.note ? String(r.note).slice(0, 300) : undefined,
    };
    if (r.materialsGapFor != null) {
      job.aiReview.gapFor = Number(r.materialsGapFor);
      job.aiReview.likelyUnrecorded = (Array.isArray(r.likelyUnrecorded) ? r.likelyUnrecorded : []).slice(0, 15).map((x) => String(x).slice(0, 120));
      job.aiReview.gapNote = r.gapNote ? String(r.gapNote).slice(0, 400) : undefined;
    }
    saved++;
  }
  for (const [mo, m] of Object.entries(months)) {
    L.recomputeMonth(m.data);
    await store.write(`${L.DIR}/${mo}.json`, m.sha, m.data, `Claude review: ${saved} job(s) [${mo}]`);
  }
  console.log(`Claude review saved for ${saved} job(s)${stale ? `, ${stale} skipped (data changed since queued)` : ''}${missing ? `, ${missing} not found` : ''}.`);
}

const o = parseArgs();
if (o.cmd === 'queue') queue(o).catch((e) => { console.error(e); process.exit(1); });
else if (o.cmd === 'apply' && o.file) apply(o).catch((e) => { console.error(e); process.exit(1); });
else { console.error('Usage: review_ai.js queue [--local dir] [--max N]  |  review_ai.js apply results.json [--local dir]'); process.exit(1); }
