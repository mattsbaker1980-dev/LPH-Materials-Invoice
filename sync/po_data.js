#!/usr/bin/env node
// Purchase orders + supplier invoices for the Job Review.
//
//   data/po/pos.json       ServiceTitan PO downloads (Purchasing > Purchase Orders > Download, "Detail" sheet),
//                          keyed by PO number. Only job POs matter for the Job Review; truck-restock POs are kept too.
//   data/po/supplier.json  What the supplier actually billed for each job PO, read from the PDFs in Mindy's
//                          OneDrive "Invoices" folder (file names look like "Hajoca 39831117  95279.pdf": vendor,
//                          supplier invoice number, our PO number).
//
// Commands (all accept --local <repoRoot>; otherwise GH_TOKEN writes to GitHub like review_sync.js):
//   import-xlsx <file.xlsx>...                 merge PO downloads (local .xlsx files)
//   (import-* and add-invoices then re-run the Job Review checks on the months those jobs are in)
//   import-text <savedPath::fileName>...       merge PO downloads read through the Microsoft 365 connector
//                                              (read_resource text of the Detail sheet: tab-separated rows)
//   todo [limit] [sinceDays]                   print the job POs still missing a supplier invoice, newest first (JSON)
//   add-invoices <results.json>                save supplier invoices read from the PDFs (see SUPPLIER_INVOICES.md)
'use strict';
const fs = require('fs');
const L = require('./review_sync.js');

const POS = 'data/po/pos.json';
const SUP = 'data/po/supplier.json';
const num = (v) => { const n = parseFloat(String(v == null ? '' : v).replace(/[$,]/g, '')); return isFinite(n) ? n : 0; };
const r2 = (n) => Math.round(n * 100) / 100;
function ymd(v) {
  if (!v) return '';
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  const m = String(v).match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})\b/);
  if (m) return `${m[3].length === 2 ? '20' + m[3] : m[3]}-${m[1].padStart(2, '0')}-${m[2].padStart(2, '0')}`;
  return String(v).slice(0, 10);
}

function mergeRows(db, rows, source) {
  let added = 0, updated = 0;
  const seen = new Set();
  for (const x of rows) {
    const po = String(x['Number'] || '').trim();
    if (!/^\d{3,7}$/.test(po)) continue;
    const line = {
      code: String(x['Material Code'] || ''), name: String(x['Material Name'] || ''), desc: String(x['Item Description'] || ''),
      vendorPart: String(x['Vendor Part'] || ''), qty: num(x['Quantity']), qtyReceived: num(x['Quantity Received']),
      cost: num(x['Cost']), total: num(x['Total']),
    };
    let rec = db.pos[po];
    if (!seen.has(po)) {
      // a fresh download replaces that PO's lines (status and costs change as the PO moves along)
      seen.add(po);
      if (rec) updated++; else added++;
      rec = db.pos[po] = {
        po, status: String(x['Status'] || ''), vendor: String(x['Vendor Name'] || ''), job: String(x['Job Number'] || '').trim(),
        bu: String(x['Business Unit'] || ''), tech: String(x['Technician'] || ''), location: String(x['Inventory Location'] || ''),
        shipTo: String(x['Ship To Name'] || ''), created: ymd(x['PO Created On Date']), sent: ymd(x['PO Sent On Date']),
        createdBy: String(x['PO Created By'] || ''), memo: String(x['Vendor Memo'] || ''), lines: [], total: 0, source,
      };
    }
    rec.lines.push(line);
    rec.total = r2(rec.total + line.total);
  }
  return { added, updated };
}

function rowsFromXlsx(file) {
  const XLSX = require('xlsx');
  const wb = XLSX.readFile(file, { cellDates: true });
  const ws = wb.Sheets['Detail'] || wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { defval: '', raw: false });
}
// read_resource text: "## Sheet: Detail ..." then a tab-separated header row starting "Number\tStatus"
function rowsFromText(text) {
  const lines = String(text).split(/\r?\n/);
  const h = lines.findIndex((l) => /^Number\tStatus\t/.test(l));
  if (h < 0) throw new Error('no "Number<TAB>Status" header row found (is this the Detail sheet of a PO download?)');
  const head = lines[h].split('\t');
  const out = [];
  for (let i = h + 1; i < lines.length; i++) {
    const l = lines[i];
    if (!l.trim() || /^(##|Formulas|This read covers|\[)/.test(l)) { if (/^##/.test(l)) break; continue; }
    const c = l.split('\t'); const o = {};
    head.forEach((k, j) => { o[k] = c[j] == null ? '' : c[j]; });
    out.push(o);
  }
  return out;
}

async function loadAll(store) {
  const p = (await store.read(POS)); const s = (await store.read(SUP));
  return {
    pos: { sha: p.sha, data: p.data || { updatedAt: null, imported: [], pos: {} } },
    sup: { sha: s.sha, data: s.data || { updatedAt: null, invoices: {}, misses: {} } },
  };
}

// supplier lines should add up to the subtotal; a big mismatch means the read went wrong
function checkInvoice(x) {
  const lines = (x.lines || []).map((l) => ({ desc: String(l.desc || '').slice(0, 160), qty: num(l.qty), unit: num(l.unit), ext: num(l.ext) }));
  const sum = r2(lines.reduce((a, l) => a + l.ext, 0));
  const subtotal = num(x.subtotal) || sum;
  const ok = !lines.length || Math.abs(sum - subtotal) <= Math.max(1, subtotal * 0.02);
  return { lines, subtotal: r2(subtotal), ok, sum };
}

// Re-run the Job Review checks on the months holding these jobs, so new PO / supplier data shows up on old jobs too.
async function recomputeJobs(store, jobNumbers) {
  const idx = (await store.read(L.DIR + '/index.json')).data;
  if (!idx) return 0;
  const months = Array.from(new Set(jobNumbers.map((j) => idx.jobMonth[j]).filter(Boolean))).sort();
  if (!months.length) return 0;
  const pb = (await store.read('data/pricebook.json')).data || {};
  pb._matRules = ((await store.read('data/baseline/material_rules.json')).data || {}).rules || null;
  pb._taskHours = L.taskHoursMap((await store.read('data/baseline/task_times.json')).data, (await store.read('data/baseline/task_live.json')).data);
  pb._po = { pos: (await store.read(POS)).data, sup: ((await store.read(SUP)).data || {}) };
  for (const mo of months) {
    const r = await store.read(`${L.DIR}/${mo}.json`);
    if (!r.data) continue;
    L.recomputeMonth(r.data, pb);
    await L.writeMonth(store, mo, r.sha, r.data, `Purchases: refresh ${mo}`);
  }
  return months.length;
}

async function main() {
  const args = process.argv.slice(2);
  const li = args.indexOf('--local'); let local = null;
  if (li >= 0) { local = args[li + 1]; args.splice(li, 2); }
  const cmd = args.shift();
  const store = L.makeStore(local ? { local } : { token: process.env.GH_TOKEN });
  const db = await loadAll(store);

  if (cmd === 'import-xlsx' || cmd === 'import-text') {
    let added = 0, updated = 0, names = [];
    for (const a of args) {
      const [file, name] = a.split('::');
      const rows = cmd === 'import-xlsx' ? rowsFromXlsx(file) : rowsFromText(fs.readFileSync(file, 'utf8'));
      const r = mergeRows(db.pos.data, rows, name || file.split('/').pop());
      added += r.added; updated += r.updated; names.push(name || file.split('/').pop());
      console.log(`READ ${name || file}: ${rows.length} line(s)`);
    }
    db.pos.data.imported = Array.from(new Set((db.pos.data.imported || []).concat(names))).slice(-500);
    db.pos.data.updatedAt = new Date().toISOString();
    await store.write(POS, db.pos.sha, db.pos.data, `PO import: ${added} new, ${updated} updated PO(s)`);
    const all = Object.values(db.pos.data.pos);
    const n = await recomputeJobs(store, Array.from(new Set(all.filter((p) => p.job).map((p) => p.job))));
    console.log(`REFRESHED ${n} month(s) of the Job Review`);
    console.log(`DONE PO import: ${added} new, ${updated} updated; ${all.length} POs on file, ${all.filter((p) => p.job).length} with a job`);
    return;
  }
  if (cmd === 'imported') { console.log(JSON.stringify(db.pos.data.imported || [])); return; }

  if (cmd === 'todo') {
    const limit = +(args[0] || 40), since = +(args[1] || 90);
    const cut = new Date(Date.now() - since * 864e5).toISOString().slice(0, 10);
    const today = new Date().toISOString().slice(0, 10);
    const inv = db.sup.data.invoices, miss = db.sup.data.misses || {};
    const list = Object.values(db.pos.data.pos)
      .filter((p) => p.job && p.created >= cut && !inv[p.po] && p.status !== 'Canceled')
      .filter((p) => { const m = miss[p.po]; if (!m) return true; if (m.tries >= 4) return false; return m.last < today; })
      .sort((a, b) => b.created.localeCompare(a.created) || b.po.localeCompare(a.po))
      .slice(0, limit)
      .map((p) => ({ po: p.po, vendor: p.vendor, job: p.job, created: p.created, poTotal: p.total }));
    console.log(JSON.stringify(list));
    return;
  }

  if (cmd === 'add-invoices') {
    const res = JSON.parse(fs.readFileSync(args[0], 'utf8'));
    if (!Array.isArray(res)) throw new Error('results must be a JSON array');
    const today = new Date().toISOString().slice(0, 10);
    let saved = 0, missed = 0, bad = 0;
    db.sup.data.misses = db.sup.data.misses || {};
    for (const x of res) {
      const po = String(x.po || '').trim(); if (!po) continue;
      if (x.notFound) { const m = db.sup.data.misses[po] || { tries: 0 }; m.tries++; m.last = today; db.sup.data.misses[po] = m; missed++; continue; }
      const c = checkInvoice(x);
      if (!c.ok) { console.log(`WARN PO ${po} ${x.file || ''}: lines add to ${c.sum} but subtotal is ${c.subtotal}; saved, marked unchecked`); bad++; }
      const rec = {
        file: String(x.file || ''), vendor: String(x.vendor || ''), invoiceNo: String(x.invoiceNo || ''), date: ymd(x.invoiceDate || x.date),
        orderedBy: x.orderedBy ? String(x.orderedBy) : undefined, lines: c.lines, subtotal: c.subtotal, tax: r2(num(x.tax)),
        total: r2(num(x.total) || c.subtotal + num(x.tax)), checked: c.ok, readAt: today,
      };
      const list = db.sup.data.invoices[po] = db.sup.data.invoices[po] || [];
      const i = list.findIndex((y) => y.file === rec.file || (rec.invoiceNo && y.invoiceNo === rec.invoiceNo));
      if (i >= 0) list[i] = rec; else list.push(rec);
      delete db.sup.data.misses[po];
      saved++;
    }
    db.sup.data.updatedAt = new Date().toISOString();
    await store.write(SUP, db.sup.sha, db.sup.data, `Supplier invoices: ${saved} read, ${missed} not found`);
    const jobs = res.filter((x) => !x.notFound && db.pos.data.pos[String(x.po)]).map((x) => db.pos.data.pos[String(x.po)].job).filter(Boolean);
    if (jobs.length) console.log(`REFRESHED ${await recomputeJobs(store, jobs)} month(s) of the Job Review`);
    console.log(`DONE supplier invoices: ${saved} saved (${bad} with totals that didn't add up), ${missed} not found; ${Object.keys(db.sup.data.invoices).length} POs have a supplier invoice`);
    return;
  }
  console.log('usage: po_data.js [--local root] import-xlsx|import-text|todo|imported|add-invoices ...');
  process.exit(1);
}
if (require.main === module) main().catch((e) => { console.log('ERROR ' + e.message); process.exit(1); });
module.exports = { rowsFromText, mergeRows, checkInvoice };
