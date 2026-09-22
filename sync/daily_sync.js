// Materials vs Invoice — daily sync script (direct-to-GitHub variant).
//
// Fetched fresh each run from GitHub (raw.githubusercontent.com) by the
// scheduled task, so it always reflects whatever is committed to
// sync/daily_sync.js in the repo — no logic is re-derived by the LLM at
// run time.
//
// Unlike sync/materials_sync.js (which POSTs to the site's Cloudflare
// Worker, for use from a real browser), this script writes straight to
// GitHub via the Contents API using a fine-grained PAT passed as GH_TOKEN.
// It exists because the Worker's own domain is not reachable from the
// scheduled-task sandbox's network egress allowlist, while api.github.com
// is.
//
// Usage:
//   GH_TOKEN=<fine-grained PAT, Contents:write on this repo only> \
//     node daily_sync.js <report-file-1> [report-file-2] [run_context]
//
// Each report file is auto-detected as either:
//   - the ServiceTitan "Jobs" custom report ("Daily Materials Vs. Invoice
//     Report" email) -> merged into data/jobs.json
//   - the joined Invoices + Invoice Items export ("Daily Invoice Line
//     Items with Description Report" email) -> merged into
//     data/invoice_checks.json
//
// Files may be passed in either order, and either one may be omitted if
// only one report showed up that day. Report files are expected as
// tab-separated text (the shape the Microsoft 365 connector's
// read_resource tool produces when it converts an .xlsx attachment), one
// header row followed by data rows. A leading "=== Sheet: ... ===" line
// (also produced by that connector) is stripped if present.
//
// This mirrors, field-for-field, the parsing/flagging logic in
// index.html's client-side script and in sync/materials_sync.js, so
// behavior stays identical no matter which path data arrives through.
// The one addition here: rows whose key id/job-number field doesn't look
// numeric are dropped as corrupted (see KNOWN ISSUE below), rather than
// silently polluting the data with garbage records.
//
// KNOWN ISSUE: the Microsoft 365 connector's xlsx->text conversion does
// not preserve embedded newlines inside a single cell (e.g. a long
// tech write-up) — those cells get split across multiple lines, which
// shifts column alignment for the fragment lines that follow until the
// next real row starts. Symptom: a handful of "phantom" records whose
// id ends up being a technician's name instead of a job/invoice number.
// This script filters those out by requiring the id to be all digits.
// A real xlsx (not text-converted) would not have this problem; fixing
// it properly would require a raw-binary attachment source, which the
// connector does not currently expose.

const fs = require('fs');

const OWNER = 'mattsbaker1980-dev';
const REPO = 'LPH-Materials-Invoice';
const BRANCH = 'main';
const GH_TOKEN = process.env.GH_TOKEN;

// ---------- shared helpers (ported from index.html / materials_sync.js) ----------

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

function stripSheetHeader(text) {
  return text.replace(/^=== Sheet:.*?===\r?\n?/, '');
}

function splitRows(text) {
  return stripSheetHeader(text)
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .split('\n')
    .filter((line) => line.length > 0)
    .map((line) => line.split('\t').map((c) => c.trim()));
}

// ---------- Jobs report ----------

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
    if (!/^\d+$/.test(String(jobNumber).trim())) continue; // guard against corrupted/shifted rows
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

// ---------- Invoice line items report ----------

const BULK_COLUMN_MAP = [
  { key: 'jobNumber', match: ['job #'] },
  { key: 'invoiceNumber', match: ['invoice #'] },
  { key: 'customer', match: ['customer name'] },
  { key: 'invoiceSummary', match: ['invoice summary'] },
  { key: 'itemName', match: ['item name'] },
  { key: 'itemType', match: ['item type'] },
  { key: 'itemPrice', match: ['item price'] },
  { key: 'pricebookPrice', match: ['pricebook price'] },
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
    jobNum = String(jobNum).trim();
    if (!/^\d+$/.test(jobNum)) continue; // guard against corrupted/shifted rows (see KNOWN ISSUE above)
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
    const itemPrice = toNum(get('itemPrice'));
    const pricebookPrice = toNum(get('pricebookPrice'));
    if (itemName) {
      const entry = { name: String(itemName), price: itemPrice };
      if (!itemPrice && pricebookPrice) { entry.listPrice = pricebookPrice; }
      if (itemType.indexOf('material') !== -1) { groups[jobNum].materialItems.push(entry); }
      else { groups[jobNum].serviceItems.push(entry); }
    }
  }
  return Object.keys(groups).map((k) => groups[k]);
}

// ---------- Discrepancy flagging ----------

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

// ---------- GitHub Contents API ----------

async function ghGet(path) {
  const url = `https://api.github.com/repos/${OWNER}/${REPO}/contents/${path}?ref=${BRANCH}`;
  const res = await fetch(url, {
    headers: {
      Authorization: `Bearer ${GH_TOKEN}`,
      'User-Agent': 'materials-invoice-daily-sync',
      Accept: 'application/vnd.github+json',
    },
  });
  if (res.status === 404) return { sha: null, data: [] };
  if (!res.ok) throw new Error(`GET ${path} failed: ${res.status} ${await res.text()}`);
  const json = await res.json();
  const data = JSON.parse(Buffer.from(json.content, 'base64').toString('utf8'));
  return { sha: json.sha, data };
}

async function ghPut(path, sha, dataObj, message) {
  const url = `https://api.github.com/repos/${OWNER}/${REPO}/contents/${path}`;
  const body = {
    message,
    content: Buffer.from(JSON.stringify(dataObj, null, 2)).toString('base64'),
    branch: BRANCH,
  };
  if (sha) body.sha = sha;
  const res = await fetch(url, {
    method: 'PUT',
    headers: {
      Authorization: `Bearer ${GH_TOKEN}`,
      'User-Agent': 'materials-invoice-daily-sync',
      Accept: 'application/vnd.github+json',
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`PUT ${path} failed: ${res.status} ${await res.text()}`);
  return res.json();
}

async function writeSyncLog(entry) {
  try {
    const { sha } = await ghGet('data/_sync_log.json').catch(() => ({ sha: null }));
    await ghPut('data/_sync_log.json', sha, entry, `Sync log: ${entry.stage} (${entry.run_context})`);
  } catch (err) {
    console.error('Failed to write sync log:', err.message);
  }
}

// ---------- Report processing ----------

async function processJobsReport(text, label, runContext) {
  const jobs = parseJobsReport(text);
  if (!jobs) { console.log(`${label}: does not look like a Jobs report, skipping.`); return null; }
  if (jobs.length === 0) { console.log(`${label}: Jobs report parsed but found 0 job rows, skipping upload.`); return null; }

  const { sha, data: existing } = await ghGet('data/jobs.json');
  const existingList = Array.isArray(existing) ? existing : [];
  const byNumber = {};
  existingList.forEach((j) => { byNumber[j.jobNumber] = j; });
  let added = 0, updated = 0;
  jobs.forEach((j) => {
    if (byNumber[j.jobNumber]) updated++; else added++;
    byNumber[j.jobNumber] = j;
  });
  const merged = Object.keys(byNumber).map((k) => byNumber[k]);
  await ghPut('data/jobs.json', sha, merged, `Materials sync (${runContext}): upsert ${jobs.length} job(s) (+${added}/~${updated})`);
  console.log(`${label}: uploaded ${jobs.length} job(s) (+${added} new / ~${updated} updated). Total now ${merged.length}.`);
  return { parsed: jobs.length, added, updated, total: merged.length };
}

async function processLineItemsReport(text, label, runContext) {
  const groups = parseLineItemsReport(text);
  if (!groups) { console.log(`${label}: does not look like an Invoice Line Items report, skipping.`); return null; }
  if (groups.length === 0) { console.log(`${label}: Line items report parsed but found 0 invoice(s), skipping upload.`); return null; }

  const { sha, data: existing } = await ghGet('data/invoice_checks.json');
  const existingList = Array.isArray(existing) ? existing : [];
  const existingById = {};
  existingList.forEach((r) => { existingById[r.id] = r; });

  let flaggedCount = 0, added = 0, updated = 0;
  const records = groups.map((g) => {
    const servicesText = g.serviceItems.map((it) => it.name).join(', ');
    const materialsText = g.materialItems.map((it) => it.name).join(', ');
    const flags = getFlags(g.invoiceSummary, servicesText, materialsText);
    if (flags.length) flaggedCount++;
    const existingRec = existingById[g.jobNumber];
    if (existingRec) updated++; else added++;
    return {
      id: g.jobNumber,
      invoiceNumber: g.jobNumber,
      jobSummary: existingRec ? existingRec.jobSummary : '',
      invoiceSummary: g.invoiceSummary,
      servicesText,
      materialsText,
      serviceItems: g.serviceItems,
      materialItems: g.materialItems,
      total: g.invoiceTotal ? g.invoiceTotal.toFixed(2) : (existingRec ? existingRec.total : ''),
      flags,
      manualOverride: existingRec ? existingRec.manualOverride : null,
      savedAt: existingRec ? existingRec.savedAt : Date.now(),
    };
  });
  records.forEach((r) => { existingById[r.id] = r; });
  const merged = Object.keys(existingById).map((k) => existingById[k]);

  await ghPut('data/invoice_checks.json', sha, merged, `Materials sync (${runContext}): upsert ${records.length} invoice check(s) (+${added}/~${updated})`);
  console.log(`${label}: uploaded ${records.length} invoice(s) (${flaggedCount} flagged, +${added} new / ~${updated} updated). Total now ${merged.length}.`);
  return { parsed: records.length, flagged: flaggedCount, added, updated, total: merged.length };
}

async function processFile(filePath, runContext) {
  const label = filePath.split('/').pop();
  const text = fs.readFileSync(filePath, 'utf8');
  const rows = splitRows(text);
  if (rows.length < 2) { console.log(`${label}: could not read a header + data row, skipping.`); return { label, result: null }; }
  const headers = rows[0];
  const jobsIdx = buildColumnIndex(headers, COLUMN_MAP);
  const bulkIdx = buildColumnIndex(headers, BULK_COLUMN_MAP);
  if (looksLikeJobsReport(jobsIdx)) {
    return { label, kind: 'jobs', result: await processJobsReport(text, label, runContext) };
  } else if (looksLikeLineItemsReport(bulkIdx)) {
    return { label, kind: 'lineItems', result: await processLineItemsReport(text, label, runContext) };
  } else {
    console.log(`${label}: header row did not match either known report shape, skipping.`);
    return { label, result: null };
  }
}

async function main() {
  if (!GH_TOKEN) {
    console.error('GH_TOKEN environment variable is required.');
    process.exit(1);
  }
  const args = process.argv.slice(2);
  let runContext = 'scheduled';
  if (args.length && !fs.existsSync(args[args.length - 1])) {
    runContext = args.pop();
  }
  const paths = args.filter((p) => fs.existsSync(p));
  if (paths.length === 0) {
    console.log(`No report files found on disk (run context: ${runContext}). Nothing to do.`);
    await writeSyncLog({ time: new Date().toISOString(), stage: 'no_files', detail: '', run_context: runContext });
    return;
  }
  console.log(`Materials vs Invoice sync starting (run context: ${runContext}) with ${paths.length} file(s).`);
  const results = [];
  let hadError = false;
  for (const p of paths) {
    try {
      results.push(await processFile(p, runContext));
    } catch (err) {
      hadError = true;
      console.error(`Error processing ${p}: ${err && err.message ? err.message : err}`);
      await writeSyncLog({ time: new Date().toISOString(), stage: 'failed', detail: `${p}: ${err && err.message ? err.message : err}`, run_context: runContext });
    }
  }
  console.log('Materials vs Invoice sync complete.');
  if (!hadError) {
    await writeSyncLog({
      time: new Date().toISOString(),
      stage: 'done',
      detail: JSON.stringify(results.map((r) => ({ label: r.label, kind: r.kind, result: r.result }))),
      run_context: runContext,
    });
  }
}

main();
