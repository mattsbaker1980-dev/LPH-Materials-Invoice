// Materials vs Invoice — daily sync script.
//
// Fetched fresh each run from GitHub (raw.githubusercontent.com) by the
// scheduled task, so it always reflects whatever is committed to
// sync/materials_sync.js in the repo — no logic is re-derived by the LLM
// at run time.
//
// Usage:
//   node materials_sync.js <report-file-1> [report-file-2] [run_context]
//
// Each report file is auto-detected as either:
//   - the ServiceTitan "Jobs" custom report ("Daily Materials Vs. Invoice
//     Report" email) -> POSTed to /api/materials-jobs
//   - the joined Invoices + Invoice Items export ("Daily Invoice Line
//     Items with Description Report" email) -> POSTed to
//     /api/materials-invoice-checks
//
// Files may be passed in either order, and either one may be omitted if
// only one report showed up that day. Report files are expected as
// tab-separated text (the shape the Microsoft 365 connector's
// read_resource tool produces when it converts an .xlsx attachment), one
// header row followed by data rows.
//
// This mirrors, field-for-field, the parsing/flagging logic in
// index.html's client-side script, so behavior stays identical whether
// data arrives via manual upload on the site or via this daily sync.
 
const fs = require('fs');
const https = require('https');
 
const WORKER_URL = 'https://lph-materials-invoice.mattsbaker1980.workers.dev';
const OWNER = 'mattsbaker1980-dev';
const REPO = 'LPH-Materials-Invoice';
const BRANCH = 'main';
 
// ---------- shared helpers (ported from index.html) ----------
 
function normalize(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/[^a-z0-9%+/.?]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}
 
function buildColumnIndex(headers, map) {
  const idx = {};
  const normHeaders = headers.map(normalize);
  map.forEach((col) => {
    let foundIdx = -1;
    col.match.forEach((m) => {
      if (foundIdx !== -1) return;
      const nm = normalize(m);
      for (let i = 0; i < normHeaders.length; i++) {
        if (normHeaders[i] === nm) { foundIdx = i; break; }
      }
    });
    if (foundIdx === -1) {
      col.match.forEach((m) => {
        if (foundIdx !== -1) return;
        const nm = normalize(m);
        for (let i = 0; i < normHeaders.length; i++) {
          if (normHeaders[i].indexOf(nm) !== -1) { foundIdx = i; break; }
        }
      });
    }
    idx[col.key] = foundIdx;
  });
  return idx;
}
 
function toNum(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v).replace(/[,%$]/g, ''));
  return isNaN(n) ? 0 : n;
}
 
// Excel serial date (days since 1899-12-30) -> ISO string, else try Date parse.
function parseDateValue(v) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : v.toISOString();
  const s = String(v).trim();
  if (/^\d+(\.\d+)?$/.test(s)) {
    const serial = parseFloat(s);
    if (serial > 20000 && serial < 90000) {
      const ms = Math.round((serial - 25569) * 86400 * 1000);
      const d = new Date(ms);
      if (!isNaN(d.getTime())) return d.toISOString();
    }
  }
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d.toISOString();
  return '';
}
 
// Split tab-separated text into rows of cells. Handles \r\n and \n.
function splitRows(text) {
  return text
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .split('\n')
    .filter((line) => line.length > 0)
    .map((line) => line.split('\t').map((c) => c.trim()));
}
 
// ---------- Jobs report (index.html: COLUMN_MAP / parseWorkbook) ----------
 
const COLUMN_MAP = [
  { key: 'jobNumber', match: ['job #', 'job number', 'job id'] },
  { key: 'jobType', match: ['job type'] },
  { key: 'businessUnit', match: ['business unit'] },
  { key: 'status', match: ['status'] },
  { key: 'summary', match: ['summary'] },
  { key: 'revenue', match: ['jobs total revenue', 'job total revenue', 'total revenue'] },
  { key: 'materialCost', match: ['material costs as % of sales'] },
  { key: 'materialCostRaw', match: ['material costs'] },
  { key: 'materialEquipPOPct', match: ['materials + equip. + po/bill costs as % of sales', 'materials + equip + po/bill costs as % of sales'] },
  { key: 'totalCosts', match: ['jobs total costs', 'job total costs'] },
  { key: 'grossMargin', match: ['jobs gross margin', 'job gross margin'] },
  { key: 'grossMarginPct', match: ['jobs gross margin %', 'job gross margin %'] },
  { key: 'technician', match: ['primary technician'] },
  { key: 'customer', match: ['customer name'] },
  { key: 'completionDate', match: ['completion date'] },
  { key: 'invoiceDate', match: ['invoice date'] },
];
 
function looksLikeJobsReport(headerIdx) {
  return headerIdx.jobNumber !== -1 && (headerIdx.materialCostRaw !== -1 || headerIdx.grossMarginPct !== -1);
}
 
function parseJobsReport(text) {
  const rows = splitRows(text);
  if (rows.length < 2) return [];
  const headers = rows[0];
  const idx = buildColumnIndex(headers, COLUMN_MAP);
  if (!looksLikeJobsReport(idx)) return null;
  const jobs = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.length === 0) continue;
    const get = (key) => { const i = idx[key]; return (i === -1 || i === undefined) ? '' : row[i]; };
    const jobNumber = get('jobNumber');
    if (jobNumber === '' || jobNumber === undefined) continue;
    const revenue = toNum(get('revenue'));
    const materialCostRaw = toNum(get('materialCostRaw'));
    const materialPct = toNum(get('materialCost'));
    const grossMargin = toNum(get('grossMargin'));
    const grossMarginPct = toNum(get('grossMarginPct'));
    const totalCosts = toNum(get('totalCosts'));
    jobs.push({
      jobNumber: String(jobNumber),
      jobType: get('jobType') || '',
      businessUnit: get('businessUnit') || 'Unassigned',
      status: get('status') || '',
      summary: get('summary') || '',
      revenue: revenue,
      materialCostRaw: materialCostRaw,
      materialPct: materialPct * (Math.abs(materialPct) <= 1 && materialPct !== 0 ? 100 : 1),
      grossMargin: grossMargin,
      grossMarginPct: grossMarginPct * (Math.abs(grossMarginPct) <= 1 && grossMarginPct !== 0 ? 100 : 1),
      totalCosts: totalCosts,
      technician: get('technician') || '',
      customer: get('customer') || '',
      completionDate: parseDateValue(get('completionDate')),
      invoiceDate: parseDateValue(get('invoiceDate')),
    });
  }
  return jobs;
}
 
// ---------- Invoice line items report (index.html: BULK_COLUMN_MAP / parseBulkInvoiceWorkbook) ----------
 
const BULK_COLUMN_MAP = [
  { key: 'jobNumber', match: ['job #'] },
  { key: 'invoiceNumber', match: ['invoice #'] },
  { key: 'customer', match: ['customer name'] },
  { key: 'invoiceSummary', match: ['invoice summary'] },
  { key: 'itemName', match: ['item name'] },
  { key: 'itemType', match: ['item type'] },
  { key: 'invoiceTotal', match: ['invoice item totals'] },
];
 
function looksLikeLineItemsReport(headerIdx) {
  return headerIdx.jobNumber !== -1 && headerIdx.itemName !== -1;
}
 
function parseLineItemsReport(text) {
  const rows = splitRows(text);
  if (rows.length < 2) return null;
  const headers = rows[0];
  const idx = buildColumnIndex(headers, BULK_COLUMN_MAP);
  if (!looksLikeLineItemsReport(idx)) return null;
  const groups = {};
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.length === 0) continue;
    const get = (key) => { const i = idx[key]; return (i === -1 || i === undefined) ? '' : row[i]; };
    let jobNum = get('jobNumber');
    if (!jobNum && jobNum !== 0) continue;
    jobNum = String(jobNum);
    if (!groups[jobNum]) {
      groups[jobNum] = {
        jobNumber: jobNum,
        invoiceNumber: String(get('invoiceNumber') || jobNum),
        customer: get('customer') || '',
        invoiceSummary: String(get('invoiceSummary') || ''),
        invoiceTotal: toNum(get('invoiceTotal')),
        serviceItems: [],
        materialItems: [],
      };
    }
    const itemName = get('itemName');
    const itemType = String(get('itemType') || '').toLowerCase();
    if (itemName) {
      if (itemType.indexOf('material') !== -1) { groups[jobNum].materialItems.push(itemName); }
      else { groups[jobNum].serviceItems.push(itemName); }
    }
  }
  return Object.keys(groups).map((k) => groups[k]);
}
 
// ---------- Discrepancy flagging (index.html: KEYWORDS / getFlags) ----------
 
const KEYWORDS = [
  'faucet', 'valve', 'toilet', 'sink', 'drain', 'pipe', 'piping', 'water heater',
  'disposal', 'sump pump', 'sump', 'pump', 'snake', 'auger', 'flange', 'wax ring',
  'supply line', 'shut-off', 'shutoff', 'trap', 'vent', 'backflow',
  'pressure reducing valve', 'expansion tank', 'anode',
  'thermostat', 'condenser', 'evaporator', 'coil', 'compressor', 'capacitor',
  'contactor', 'furnace', 'blower', 'filter', 'ductwork', 'refrigerant',
  'breaker', 'panel', 'outlet', 'light switch', 'wall switch', 'wiring', 'gfci', 'disconnect', 'motor',
  'igniter', 'ignitor', 'thermocouple', 'gas valve', 'heat exchanger',
  'humidifier', 'uv light', 'zone valve', 'circulator', 'boiler', 'radiator',
  'pop up assembly', 'pop-up assembly', 'garbage disposal', 'water softener',
  'pressure tank', 'well pump', 'ejector pump',
];
 
function escRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
 
function getFlags(invoiceSummary, servicesText, materialsText) {
  const invoiced = ((servicesText || '') + ' ' + (materialsText || '')).toLowerCase();
  const narrative = (invoiceSummary || '').toLowerCase();
  const flags = [];
  KEYWORDS.forEach((k) => {
    const re = new RegExp('\\b' + escRegex(k) + 's?\\b', 'i');
    if (re.test(narrative) && !re.test(invoiced)) { flags.push(k); }
  });
  return flags;
}
 
// ---------- GitHub read (to preserve manualOverride / jobSummary / savedAt on merge) ----------
 
function ghRawGet(path) {
  return new Promise((resolve, reject) => {
    const url = `https://raw.githubusercontent.com/${OWNER}/${REPO}/${BRANCH}/${path}?t=${Date.now()}`;
    https.get(url, (res) => {
      if (res.statusCode !== 200) { resolve([]); res.resume(); return; }
      let data = '';
      res.on('data', (chunk) => { data += chunk; });
      res.on('end', () => {
        try { resolve(JSON.parse(data)); } catch (e) { resolve([]); }
      });
    }).on('error', () => resolve([]));
  });
}
 
// ---------- POST to Worker ----------
 
function postToWorker(path, body) {
  return new Promise((resolve, reject) => {
    const data = JSON.stringify(body);
    const url = new URL(WORKER_URL + path);
    const req = https.request(
      {
        hostname: url.hostname,
        path: url.pathname,
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Content-Length': Buffer.byteLength(data),
        },
      },
      (res) => {
        let resBody = '';
        res.on('data', (chunk) => { resBody += chunk; });
        res.on('end', () => {
          if (res.statusCode < 200 || res.statusCode >= 300) {
            reject(new Error(`Worker request to ${path} failed: ${res.statusCode} ${resBody}`));
          } else {
            resolve(resBody);
          }
        });
      }
    );
    req.on('error', reject);
    req.write(data);
    req.end();
  });
}
 
// ---------- Report processing ----------
 
async function processJobsReport(text, label) {
  const jobs = parseJobsReport(text);
  if (!jobs) { console.log(`${label}: does not look like a Jobs report, skipping.`); return; }
  if (jobs.length === 0) { console.log(`${label}: Jobs report parsed but found 0 job rows, skipping upload.`); return; }
  console.log(`${label}: parsed ${jobs.length} job(s) from Jobs report, uploading...`);
  await postToWorker('/api/materials-jobs', { jobs });
  console.log(`${label}: uploaded ${jobs.length} job(s).`);
}
 
async function processLineItemsReport(text, label) {
  const groups = parseLineItemsReport(text);
  if (!groups) { console.log(`${label}: does not look like an Invoice Line Items report, skipping.`); return; }
  if (groups.length === 0) { console.log(`${label}: Line items report parsed but found 0 invoice(s), skipping upload.`); return; }
 
  const existingRecords = await ghRawGet('data/invoice_checks.json');
  const existingById = {};
  (existingRecords || []).forEach((r) => { existingById[r.id] = r; });
 
  let flaggedCount = 0;
  const records = groups.map((g) => {
    const servicesText = g.serviceItems.join(', ');
    const materialsText = g.materialItems.join(', ');
    const flags = getFlags(g.invoiceSummary, servicesText, materialsText);
    if (flags.length) flaggedCount++;
    const existing = existingById[g.jobNumber];
    return {
      id: g.jobNumber,
      invoiceNumber: g.jobNumber,
      jobSummary: existing ? existing.jobSummary : '',
      invoiceSummary: g.invoiceSummary,
      servicesText,
      materialsText,
      total: g.invoiceTotal ? g.invoiceTotal.toFixed(2) : (existing ? existing.total : ''),
      flags,
      manualOverride: existing ? existing.manualOverride : null,
      savedAt: existing ? existing.savedAt : Date.now(),
    };
  });
 
  console.log(`${label}: parsed ${records.length} invoice(s) from Line Items report (${flaggedCount} flagged), uploading...`);
  await postToWorker('/api/materials-invoice-checks', { records });
  console.log(`${label}: uploaded ${records.length} invoice(s).`);
}
 
async function processFile(filePath) {
  const label = filePath.split('/').pop();
  const text = fs.readFileSync(filePath, 'utf8');
  const rows = splitRows(text);
  if (rows.length < 2) { console.log(`${label}: could not read a header + data row, skipping.`); return; }
  const headers = rows[0];
  const jobsIdx = buildColumnIndex(headers, COLUMN_MAP);
  const bulkIdx = buildColumnIndex(headers, BULK_COLUMN_MAP);
  if (looksLikeJobsReport(jobsIdx)) {
    await processJobsReport(text, label);
  } else if (looksLikeLineItemsReport(bulkIdx)) {
    await processLineItemsReport(text, label);
  } else {
    console.log(`${label}: header row did not match either known report shape, skipping.`);
  }
}
 
async function main() {
  const args = process.argv.slice(2);
  let runContext = 'scheduled';
  if (args.length && !fs.existsSync(args[args.length - 1])) {
    runContext = args.pop();
  }
  const paths = args.filter((p) => fs.existsSync(p));
  if (paths.length === 0) {
    console.log(`No report files found on disk (run context: ${runContext}). Nothing to do.`);
    return;
  }
  console.log(`Materials vs Invoice sync starting (run context: ${runContext}) with ${paths.length} file(s).`);
  for (const p of paths) {
    try {
      await processFile(p);
    } catch (err) {
      console.error(`Error processing ${p}: ${err && err.message ? err.message : err}`);
    }
  }
  console.log('Materials vs Invoice sync complete.');
}
 
main();
 
