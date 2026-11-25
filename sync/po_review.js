#!/usr/bin/env node
// Purchase Orders page (purchases.html): supplier invoices from Mindy's OneDrive "Invoices" folder,
// each matched to the ServiceTitan job it was bought for and sorted into a department tab.
//
//   data/po/invoices.json  every supplier invoice read so far, with its match (written here)
//   data/po/refresh.json   current PO/bill cost per job, all departments (the daily "PO Cost Refresh" report,
//                          saved by review_sync.js)
//   data/review/*.json     Job Review jobs (tech, customer, what was billed, write-up)
//
// How a match is made (an invoice's "amount due" equals the job's PO/Bill cost in ServiceTitan, tax included):
//   1. amount  - one job completed near the order date whose PO/bill cost equals the invoice total (or subtotal)
//   2. customer- the invoice's job name / job number field names that job's customer
//   3. tech    - "ordered by" is a tech who had exactly one job that day that could hold this purchase
//   otherwise  - a short list of possible jobs to pick from, or "not matched yet" (retried for 21 days)
//
// Commands (GH_TOKEN writes to GitHub; --local <repoRoot> works on a checkout):
//   known                      list the invoice files already recorded (so the task only reads new ones)
//   add <results.json>         save invoices read from the PDFs (see PO_INVOICES.md), then re-match
//   rematch                    re-match everything still open (new costs may have arrived)
'use strict';
const fs = require('fs');
const L = require('./review_sync.js');

const INV = 'data/po/invoices.json';
const WINDOW_BEFORE = 3, WINDOW_AFTER = 21, RETRY_DAYS = 21, KEEP_DAYS = 120;
const r2 = (n) => Math.round(n * 100) / 100;
const num = (v) => { const n = parseFloat(String(v == null ? '' : v).replace(/[$,]/g, '')); return isFinite(n) ? n : 0; };
const dayN = (s) => Math.floor(Date.parse(String(s).slice(0, 10) + 'T12:00:00Z') / 864e5);
const today = () => new Date().toLocaleDateString('en-CA', { timeZone: 'America/New_York' });
function ymd(v) {
  if (!v) return '';
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/); if (m) return m[0];
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})\b/);
  if (m) return `${m[3].length === 2 ? '20' + m[3] : m[3]}-${m[1].padStart(2, '0')}-${m[2].padStart(2, '0')}`;
  return '';
}
// "Hajoca 39831117  95279.pdf" -> { vendor, invoiceNo, po }, "APR 13393570  Consumables.pdf" -> { tail: 'Consumables' }
function parseName(name) {
  const base = String(name).replace(/\.pdf$/i, '').trim();
  const m = base.match(/^(.*?)\s+(\S+)\s{1,}(\S+)$/);
  if (!m) return { vendor: base, invoiceNo: '', po: '', tail: '' };
  const tail = m[3];
  return { vendor: m[1].trim(), invoiceNo: m[2], po: /^\d{4,6}$/.test(tail) && tail !== '7777' ? tail : '', tail: /^\d{4,6}$/.test(tail) && tail !== '7777' ? '' : tail };
}
// suppliers that send scanned (image) PDFs; all three are HVAC suppliers
const SCAN_VENDOR_DEPT = /^(re michel|johnstone|sid h)/i;
const STOCK_TAIL = /^(stock|consumables?|shop|trk\d*|truck\d*|tools?|7777)$/i;
const idOf = (uri) => String(uri || '').split('/').pop();

function group(bu) {
  const d = L.deptOf(bu);
  if (/^HVAC/.test(d)) return 'HVAC';
  if (/^Sewer/.test(d)) return 'Sewer';
  if (d === 'Plumbing') return 'Plumbing';
  if (d === 'Electrical') return 'Electrical';
  return 'General';
}
const firstName = (s) => String(s || '').trim().split(/[\s.,]+/)[0].toLowerCase();
// names on supplier counters that are spelled differently in ServiceTitan
const NAME_ALIAS = { tyler: 'tylor' };
const custWords = (s) => String(s || '').toLowerCase().replace(/[^a-z& ]+/g, ' ').split(/\s+/).filter((w) => w.length >= 4 && !/^(lancaster|plumbing|heating|plbg|htg|company|inc|job|the|and)$/.test(w));

// ---------------------------------------------------------------- job pool
async function loadPool(store) {
  const refresh = ((await store.read('data/po/refresh.json')).data || { jobs: {} }).jobs;
  const idx = (await store.read(`${L.DIR}/index.json`)).data || { months: [] };
  const since = new Date(Date.now() - (KEEP_DAYS + 30) * 864e5).toISOString().slice(0, 7);
  const jr = {};
  for (const mo of (idx.months || []).filter((m) => m >= since)) {
    const d = (await store.read(`${L.DIR}/${mo}.json`)).data;
    if (d && d.jobs) Object.assign(jr, d.jobs);
  }
  const pool = {};
  const add = (k) => pool[k] || (pool[k] = { job: k });
  for (const [k, j] of Object.entries(jr)) {
    const p = add(k);
    p.bu = j.businessUnit; p.done = j.completionDate; p.cust = j.customer;
    p.techs = String((j.detail && j.detail.assignedTechs) || j.primaryTech || '').split(/\s*,\s*/).filter(Boolean);
    p.tech = j.primaryTech;
    p.cost = r2((j.invoices || []).reduce((s, i) => s + (i.poCost || 0) + (i.billCost || 0), 0));
    p.jr = j;
  }
  for (const [k, x] of Object.entries(refresh)) {
    const p = add(k);
    p.bu = p.bu || x.bu; p.done = p.done || x.done; p.cust = p.cust || x.cust; p.tech = p.tech || x.tech;
    if (!p.techs || !p.techs.length) p.techs = String(x.techs || x.tech || '').split(/\s*,\s*/).filter(Boolean);
    p.cost = x.po; // the refresh report is newer than the Job Review snapshot
    p.matCost = x.mat; p.eqCost = x.eq;
  }
  const list = Object.values(pool).filter((p) => p.done);
  // each tech's usual department (for invoices that name a tech but match no job)
  const techDept = {};
  list.forEach((p) => p.techs.forEach((t) => { const f = firstName(t); const g = group(p.bu); (techDept[f] = techDept[f] || {})[g] = ((techDept[f] || {})[g] || 0) + 1; }));
  const usual = {};
  for (const [f, c] of Object.entries(techDept)) usual[f] = Object.entries(c).sort((a, b) => b[1] - a[1])[0][0];
  return { list, byJob: pool, usual };
}

function snapshot(p) {
  const j = p.jr;
  const out = { job: p.job, customer: p.cust || '', tech: (p.techs || []).join(', ') || p.tech || '', bu: p.bu || '', dept: group(p.bu), date: p.done || '', poCost: p.cost || 0 };
  if (p.matCost != null) out.matCost = p.matCost;
  if (j) {
    out.jobType = j.jobType || '';
    out.revenue = (j.materials && j.materials.revenue) || (j.invoices || []).reduce((s, i) => s + (i.total || 0), 0);
    out.items = (j.invoices || []).flatMap((i) => (i.items || [])).filter((x) => x && x.name).slice(0, 30)
      .map((x) => ({ name: x.name, type: x.type || '', price: x.price || 0 }));
    out.writeUp = (j.invoices || []).map((i) => i.summary || '').filter(Boolean).join('\n\n').slice(0, 900);
  }
  return out;
}

// ---------------------------------------------------------------- matching
function matchOne(inv, pool) {
  const name = parseName(inv.file);
  if (!inv.po && name.po) inv.po = name.po;
  const tail = name.tail || '';
  if (inv.unreadable) return { status: 'unreadable', dept: SCAN_VENDOR_DEPT.test(inv.vendor || inv.file) ? 'HVAC' : 'General' };
  if (STOCK_TAIL.test(tail) || /\btools\b/i.test(tail)) {
    const f = NAME_ALIAS[firstName(inv.orderedBy)] || firstName(inv.orderedBy);
    return { status: 'stock', dept: pool.usual[f] || 'General', note: tail };
  }
  const d0 = dayN(inv.orderDate || inv.invoiceDate);
  if (!isFinite(d0)) return { status: 'waiting', dept: 'General' };
  const win = pool.list.filter((p) => { const d = dayN(p.done) - d0; return d >= -WINDOW_BEFORE && d <= WINDOW_AFTER; });
  const total = num(inv.total), sub = num(inv.subtotal);
  const eq = (a, b) => b > 0 && Math.abs(a - b) <= 0.011;
  const ob = NAME_ALIAS[firstName(inv.orderedBy)] || firstName(inv.orderedBy);
  const refWords = custWords(inv.jobRef);
  const hasTech = (p) => ob && p.techs.some((t) => firstName(t) === ob);
  const hasCust = (p) => refWords.length && custWords(p.cust).some((w) => refWords.includes(w));
  const closer = (a, b) => Math.abs(dayN(a.done) - d0) - Math.abs(dayN(b.done) - d0);
  const result = (p, how) => ({ status: 'matched', how, job: p.job, dept: group(p.bu), jobInfo: snapshot(p) });

  // 1. exact amount
  let hits = win.filter((p) => eq(p.cost || 0, total) || eq(p.cost || 0, sub));
  if (hits.length === 1) return result(hits[0], 'amount');
  if (hits.length > 1) {
    const narrowed = hits.filter((p) => hasCust(p) || hasTech(p));
    if (narrowed.length === 1) return result(narrowed[0], 'amount');
  }
  const exact = hits;
  // 2. customer named on the invoice
  if (refWords.length) {
    const c = win.filter((p) => hasCust(p) && (p.cost || 0) >= Math.min(total, sub || total) - 0.011);
    if (c.length === 1) return result(c[0], 'customer');
    if (c.length > 1) hits = c;
  }
  // 3. ordered by a tech with one possible job that day or the next
  if (ob) {
    const t = win.filter((p) => hasTech(p) && dayN(p.done) - d0 >= 0 && dayN(p.done) - d0 <= 1 && (p.cost || 0) >= Math.min(total, sub || total) - 0.011);
    if (t.length === 1) return result(t[0], 'tech');
    if (t.length > 1 && t.length <= 5 && !exact.length && !(refWords.length && hits.length)) hits = t;
  }
  if (hits.length > 1 && hits.length <= 6) {
    const groups = [...new Set(hits.map((p) => group(p.bu)))];
    return { status: 'choose', dept: groups.length === 1 ? groups[0] : (pool.usual[ob] || 'General'), candidates: hits.sort(closer).map(snapshot) };
  }
  return { status: 'waiting', dept: pool.usual[ob] || 'General' };
}

function rematch(db, pool) {
  const t = dayN(today());
  const n = { matched: 0, choose: 0, waiting: 0, stock: 0, unreadable: 0, nomatch: 0 };
  for (const inv of Object.values(db.invoices)) {
    const age = t - dayN(inv.orderDate || inv.invoiceDate || inv.readAt);
    const open = !inv.match || inv.match.status === 'waiting' || inv.match.status === 'choose';
    // matched invoices get a fresh job snapshot for a week (costs and write-ups settle), open ones are retried for 21 days
    if (open && age <= RETRY_DAYS) inv.match = matchOne(inv, pool);
    else if (inv.match && inv.match.status === 'matched' && age <= 7 && pool.byJob[inv.match.job]) inv.match.jobInfo = snapshot(pool.byJob[inv.match.job]);
    else if (open && inv.match && inv.match.status === 'waiting') inv.match.status = 'nomatch';
    n[inv.match.status] = (n[inv.match.status] || 0) + 1;
  }
  return n;
}

function validate(x) {
  const errs = [];
  if (!x.file || !x.uri) errs.push('file and uri are required');
  if (!x.unreadable) {
    if (!(num(x.total) || num(x.subtotal))) errs.push('no total');
    const lines = Array.isArray(x.lines) ? x.lines : [];
    const sum = r2(lines.reduce((s, l) => s + num(l.ext), 0));
    if (lines.length && num(x.subtotal) && Math.abs(sum - num(x.subtotal)) > 0.05) errs.push(`lines add up to ${sum}, subtotal is ${x.subtotal}`);
  }
  return errs;
}

async function main() {
  const args = process.argv.slice(2);
  const opts = { local: null };
  const rest = [];
  for (let i = 0; i < args.length; i++) { if (args[i] === '--local') opts.local = args[++i]; else rest.push(args[i]); }
  const [cmd, file] = rest;
  if (!opts.local && !process.env.GH_TOKEN) { console.error('GH_TOKEN (or --local <repoRoot>) is required'); process.exit(1); }
  const store = L.makeStore(opts);
  const res = await store.read(INV);
  const db = res.data || { invoices: {} };

  if (cmd === 'known') {
    const names = Object.values(db.invoices).map((x) => x.file).sort();
    console.log(JSON.stringify({ count: names.length, highestPO: Math.max(0, ...Object.values(db.invoices).map((x) => +x.po || 0)), files: names }));
    return;
  }
  if (cmd !== 'add' && cmd !== 'rematch') { console.error('Usage: po_review.js known | add <results.json> | rematch'); process.exit(1); }

  let added = 0, rejected = [];
  if (cmd === 'add') {
    const items = JSON.parse(fs.readFileSync(file, 'utf8'));
    for (const x of (Array.isArray(items) ? items : [items])) {
      const errs = validate(x);
      if (errs.length) { rejected.push(`${x.file || '?'}: ${errs.join('; ')}`); continue; }
      const id = idOf(x.uri);
      const name = parseName(x.file);
      db.invoices[id] = {
        id, file: x.file, uri: x.uri, webUrl: x.webUrl || '', vendor: x.vendor || name.vendor, invoiceNo: x.invoiceNo || name.invoiceNo,
        po: String(x.po || name.po || ''), invoiceDate: ymd(x.invoiceDate), orderDate: ymd(x.orderDate) || ymd(x.invoiceDate),
        orderedBy: x.orderedBy || '', jobRef: x.jobRef || '', unreadable: !!x.unreadable,
        lines: (x.lines || []).map((l) => ({ desc: String(l.desc || ''), qty: num(l.qty), unit: num(l.unit), ext: num(l.ext) })),
        subtotal: num(x.subtotal), tax: num(x.tax), total: num(x.total) || num(x.subtotal),
        readAt: today(), match: null,
      };
      added++;
    }
  }
  // keep 120 days
  const cutoff = dayN(today()) - KEEP_DAYS;
  for (const [k, v] of Object.entries(db.invoices)) if (dayN(v.orderDate || v.invoiceDate || v.readAt) < cutoff) delete db.invoices[k];

  const pool = await loadPool(store);
  const n = rematch(db, pool);
  db.updatedAt = new Date().toISOString();
  await store.write(INV, res.sha, db, `Purchase Orders: ${cmd === 'add' ? `+${added} invoice(s), ` : ''}re-matched`);
  if (rejected.length) console.log('REJECTED ' + rejected.length + ':\n  ' + rejected.join('\n  '));
  console.log(`DONE ${cmd === 'add' ? `added ${added}, ` : ''}${Object.keys(db.invoices).length} invoices on file: ` +
    `${n.matched} matched, ${n.choose} pick-a-job, ${n.waiting} not matched yet, ${n.stock} stock, ${n.unreadable} unreadable scans, ${n.nomatch || 0} gave up`);
}

if (require.main === module) main().catch((e) => { console.error('ERROR', e.stack || e.message); process.exit(1); });
module.exports = { parseName, matchOne, group };
