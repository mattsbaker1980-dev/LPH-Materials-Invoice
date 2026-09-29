// Job Review sync: reads the four ServiceTitan .xlsx reports (real workbook
// files, extracted from Gmail by gmail_extract.js), merges them into one
// record per job, computes review flags, and writes month files under
// data/review/ in this repo.
//
// Reports (auto-detected from their header row, any order, any subset):
//   - Materials vs Invoice      -> costs, margin, job booking notes
//   - Daily Job Detail          -> hours, estimate count, campaign, techs
//   - Invoice Line Items w/ Desc-> tech write-up (invoice summary), billed
//                                  items, PO + material costs per invoice
//   - All Estimates – Daily     -> every estimate (options) with summary
//
// Usage:
//   npm i xlsx@0.18.5   (once per sandbox)
//   GH_TOKEN=... node review_sync.js [--local <repoDir>] [--context scheduled] file1.xlsx [file2.xlsx ...]
//
// --local writes into a local clone instead of the GitHub Contents API.

const fs = require('fs');
const path = require('path');
let XLSX;
try { XLSX = require('xlsx'); } catch (e) {
  console.error('The "xlsx" package is missing. Run: npm i xlsx@0.18.5');
  process.exit(1);
}

const OWNER = 'mattsbaker1980-dev';
const REPO = 'LPH-Materials-Invoice';
const BRANCH = 'main';
const DIR = 'data/review';
const PENDING_TTL_DAYS = 90;

// ---------------------------------------------------------------- helpers

const norm = (s) => String(s == null ? '' : s).toLowerCase().replace(/\s+/g, ' ').trim();
const str = (v) => (v == null ? '' : String(v)).trim();
const num = (v) => {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  const n = parseFloat(String(v).replace(/[$,%\s]/g, ''));
  return isFinite(n) ? n : 0;
};
const bool = (v) => v === true || /^(true|yes|y|1)$/i.test(str(v));
const round = (n, d = 2) => Math.round(n * 10 ** d) / 10 ** d;
const isJobNo = (v) => /^\d{6,}$/.test(str(v));

function isoDate(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date && !isNaN(v)) {
    // SheetJS returns dates at local midnight; use UTC parts after nudging by 12h to avoid TZ drift
    const d = new Date(v.getTime() + 12 * 3600 * 1000);
    return d.toISOString().slice(0, 10);
  }
  if (typeof v === 'number' && v > 20000 && v < 80000) {
    const d = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
    return d.toISOString().slice(0, 10);
  }
  const s = str(v);
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (m) { const y = m[3].length === 2 ? '20' + m[3] : m[3]; return `${y}-${m[1].padStart(2, '0')}-${m[2].padStart(2, '0')}`; }
  return null;
}

// Header resolver. Headers may be plain ("Invoice ID") or suffixed by
// template ("Invoice ID (Invoice Items)"); duplicates are allowed.
function makeCols(headerRow) {
  const hs = headerRow.map((h) => norm(h));
  return {
    has(name) { return this.find(name) !== -1; },
    // find(name, template?, occurrence?)
    find(name, template, occurrence = 0) {
      const n = norm(name);
      const hits = [];
      hs.forEach((h, i) => {
        const base = h.replace(/\s*\((invoices|invoice items)\)\s*$/, '');
        const tmpl = (h.match(/\((invoices|invoice items)\)\s*$/) || [])[1] || null;
        if (base === n && (!template || !tmpl || tmpl === template)) hits.push({ i, tmpl });
      });
      if (template) {
        const exact = hits.filter((h) => h.tmpl === template);
        if (exact.length) return exact[occurrence] ? exact[occurrence].i : -1;
      }
      return hits[occurrence] ? hits[occurrence].i : -1;
    },
  };
}

function readWorkbook(file) {
  const wb = XLSX.readFile(file, { cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true, blankrows: false });
  let filters = {};
  const fname = wb.SheetNames.find((n) => /filter/i.test(n));
  if (fname) {
    XLSX.utils.sheet_to_json(wb.Sheets[fname], { header: 1, defval: null }).forEach((r) => {
      if (r && r[0] != null && r[1] != null) filters[str(r[0])] = str(r[1]);
    });
  }
  return { rows, filters };
}

function detect(cols) {
  if (cols.has('estimate id')) return 'estimates';
  if (cols.has('item name')) return 'lineItems';
  if (cols.has('total hours worked') || cols.has('sold hours')) return 'jobDetail';
  if (cols.has('material costs as % of sales') || cols.has('jobs gross margin')) return 'materials';
  return null;
}

// ---------------------------------------------------------------- parsers
// Each parser returns an array of partial updates: { jobNumber, section, data }

function parseMaterials(rows, cols) {
  const c = (n) => cols.find(n);
  const I = {
    job: c('job #'), type: c('job type'), bu: c('business unit'), status: c('status'), summary: c('summary'),
    rev: c('jobs total revenue'), matPct: c('material costs as % of sales'), mat: c('material costs'),
    mepo: c('materials + equip. + po/bill costs'), mepoPct: c('materials + equip. + po/bill costs as % of sales'),
    costs: c('jobs total costs'), gm: c('jobs gross margin'), gmPct: c('jobs gross margin %'),
    tech: c('primary technician'), cust: c('customer name'), done: c('completion date'), invDate: c('invoice date'),
  };
  const out = [];
  for (const r of rows) {
    const job = str(r[I.job]);
    if (!isJobNo(job) || (!str(r[I.cust]) && !str(r[I.type]))) continue; // skip totals rows
    const pct = (v) => { const n = num(v); return Math.abs(n) <= 1.5 && n !== 0 ? n * 100 : n; };
    out.push({
      jobNumber: job,
      base: {
        customer: str(r[I.cust]), jobType: str(r[I.type]), businessUnit: str(r[I.bu]), status: str(r[I.status]),
        primaryTech: str(r[I.tech]), completionDate: isoDate(r[I.done]), invoiceDate: isoDate(r[I.invDate]),
      },
      section: 'materials',
      data: {
        bookingNotes: str(r[I.summary]),
        revenue: round(num(r[I.rev])),
        materialCost: round(num(r[I.mat])),
        matEquipPoCost: round(num(r[I.mepo])),
        totalCosts: round(num(r[I.costs])),
        grossMargin: round(num(r[I.gm])),
        grossMarginPct: round(pct(r[I.gmPct]), 1),
      },
    });
  }
  return out;
}

function parseJobDetail(rows, cols) {
  const c = (n) => cols.find(n);
  const I = {
    id: c('job id'), job: c('job #'), type: c('job type'), campaign: c('job campaign'), bu: c('business unit'),
    status: c('status'), sched: c('scheduled date'), done: c('completion date'), techs: c('assigned technicians'),
    warranty: c('warranty'), recall: c('recall'), converted: c('converted'), zero: c('zero dollar job'),
    tech: c('primary technician'), sold: c('sold hours'), paid: c('total technician paid time'),
    worked: c('total hours worked'), nonbill: c('non-billable hours'), soldBy: c('sold by'), est: c('estimates'),
    estSales: c('jobs estimate sales subtotal'), rev: c('jobs total revenue'), member: c('member status'),
    jobClass: c('job class'),
  };
  const out = [];
  for (const r of rows) {
    const job = str(r[I.job]);
    if (!isJobNo(job) || (!str(r[I.type]) && !str(r[I.tech]))) continue;
    out.push({
      jobNumber: job,
      base: {
        jobId: str(r[I.id]) || job, jobType: str(r[I.type]), businessUnit: str(r[I.bu]), status: str(r[I.status]),
        primaryTech: str(r[I.tech]), completionDate: isoDate(r[I.done]), scheduledDate: isoDate(r[I.sched]),
      },
      section: 'detail',
      data: {
        assignedTechs: str(r[I.techs]), campaign: str(r[I.campaign]), jobClass: str(r[I.jobClass]),
        warranty: bool(r[I.warranty]), recall: bool(r[I.recall]), converted: bool(r[I.converted]),
        zeroDollar: bool(r[I.zero]), memberStatus: str(r[I.member]), soldBy: str(r[I.soldBy]),
        soldHours: round(num(r[I.sold])), paidHours: round(num(r[I.paid])), hoursWorked: round(num(r[I.worked])),
        nonBillableHours: round(num(r[I.nonbill])), estimatesCount: num(r[I.est]),
        estimateSales: round(num(r[I.estSales])), revenue: round(num(r[I.rev])),
      },
    });
  }
  return out;
}

function parseLineItems(rows, cols) {
  const I = {
    summary: cols.find('invoice summary', 'invoices'),
    job: cols.find('job #', 'invoices'),
    inv: cols.find('invoice #', 'invoices'),
    cust: cols.find('customer name', 'invoices'),
    invTotal: cols.find('invoice item totals', 'invoices'),
    invId: cols.find('invoice id', 'invoices', 0),
    itemJob: cols.find('job number', 'invoice items'),
    itemName: cols.find('item name', 'invoice items'),
    itemCode: cols.find('item code', 'invoice items'),
    itemType: cols.find('item type', 'invoice items'),
    price: cols.find('item price', 'invoice items'),
    pb: cols.find('pricebook price', 'invoice items'),
    tech: cols.find('primary technician', 'invoice items'),
    soldBy: cols.find('sold by technician', 'invoice items'),
    mat: cols.find('material costs', 'invoices'),
    po: cols.find('purchase order costs', 'invoices'),
    equip: cols.find('equipment costs', 'invoices'),
    bill: cols.find('bill costs', 'invoices'),
  };
  const invoices = {};
  for (const r of rows) {
    let job = str(r[I.job]);
    if (!isJobNo(job)) job = str(r[I.itemJob]);
    if (!isJobNo(job)) continue; // membership/recurring invoices with no job, and totals rows
    const invNo = str(r[I.inv]) || job;
    if (!invoices[invNo]) {
      invoices[invNo] = {
        jobNumber: job, customer: str(r[I.cust]),
        invoice: {
          invoiceNumber: invNo, invoiceId: str(r[I.invId]), summary: str(r[I.summary]),
          total: round(num(r[I.invTotal])), materialCost: round(num(r[I.mat])),
          poCost: round(num(r[I.po])), equipmentCost: round(num(r[I.equip])),
          billCost: I.bill === -1 ? 0 : round(num(r[I.bill])), items: [],
        },
        tech: '',
      };
    }
    const g = invoices[invNo];
    if (!g.invoice.summary && str(r[I.summary])) g.invoice.summary = str(r[I.summary]);
    if (!g.tech && str(r[I.tech])) g.tech = str(r[I.tech]);
    const name = str(r[I.itemName]);
    if (name) {
      const it = { name, code: str(r[I.itemCode]), type: str(r[I.itemType]) || 'Service', price: round(num(r[I.price])) };
      const pb = round(num(r[I.pb]));
      if (pb && pb !== it.price) it.pricebook = pb;
      if (str(r[I.soldBy])) it.soldBy = str(r[I.soldBy]);
      g.invoice.items.push(it);
    }
  }
  return Object.values(invoices).map((g) => ({
    jobNumber: g.jobNumber,
    base: { customer: g.customer, primaryTech: g.tech },
    section: 'invoice',
    data: g.invoice,
  }));
}

function parseEstimates(rows, cols) {
  const c = (n) => cols.find(n);
  const I = {
    id: c('estimate id'), name: c('estimate name'), jobId: c('parent job id'), job: c('parent job number'),
    bu: c('business unit'), oppStatus: c('opportunity status'), status: c('estimate status'),
    rec: c('recommended'), sub: c('estimates subtotal'), soldOn: c('sold on'), soldBy: c('sold by'),
    cust: c('customer name'), summary: c('estimate summary') !== -1 ? c('estimate summary') : c('summary'),
    created: c('creation date') !== -1 ? c('creation date') : c('created on'),
    createdBy: c('estimate created by') !== -1 ? c('estimate created by') : c('created by'),
  };
  const out = [];
  for (const r of rows) {
    const id = str(r[I.id]);
    let job = str(r[I.job]);
    if (!isJobNo(job)) job = str(r[I.jobId]);
    if (!isJobNo(id) || !isJobNo(job)) continue;
    const e = {
      id, name: str(r[I.name]), status: str(r[I.status]), opportunityStatus: str(r[I.oppStatus]),
      recommended: bool(r[I.rec]), subtotal: round(num(r[I.sub])), soldOn: isoDate(r[I.soldOn]),
      soldBy: str(r[I.soldBy]),
    };
    if (I.summary !== -1) e.summary = str(r[I.summary]);
    if (I.created !== -1) e.created = isoDate(r[I.created]);
    if (I.createdBy !== -1) e.createdBy = str(r[I.createdBy]);
    out.push({ jobNumber: job, base: { customer: str(r[I.cust]) }, section: 'estimate', data: e });
  }
  return out;
}

const PARSERS = { materials: parseMaterials, jobDetail: parseJobDetail, lineItems: parseLineItems, estimates: parseEstimates };

// ---------------------------------------------------------------- merge

function emptyJob(jobNumber) {
  return { jobNumber, jobId: '', customer: '', jobType: '', businessUnit: '', status: '', primaryTech: '',
    completionDate: null, scheduledDate: null, invoiceDate: null,
    detail: null, materials: null, invoices: [], estimates: [], flags: [], updated: {} };
}

function applyUpdate(job, u, today) {
  for (const [k, v] of Object.entries(u.base || {})) {
    if (v !== '' && v != null) {
      // Job Detail / Materials are authoritative for job-level fields; don't let later reports blank them
      if (!job[k] || u.section === 'detail' || u.section === 'materials') job[k] = v;
    }
  }
  if (u.section === 'detail') { job.detail = u.data; job.updated.detail = today; }
  else if (u.section === 'materials') { job.materials = u.data; job.updated.materials = today; }
  else if (u.section === 'invoice') {
    const i = job.invoices.findIndex((x) => x.invoiceNumber === u.data.invoiceNumber);
    if (i === -1) job.invoices.push(u.data); else job.invoices[i] = u.data;
    job.updated.invoices = today;
  } else if (u.section === 'estimate') {
    const i = job.estimates.findIndex((x) => x.id === u.data.id);
    if (i === -1) job.estimates.push(u.data);
    else job.estimates[i] = Object.assign({}, job.estimates[i], u.data);
    job.updated.estimates = today;
  }
}

// ---------------------------------------------------------------- flags

const KEYWORDS = [
  'faucet', 'valve', 'toilet', 'sink', 'drain', 'pipe', 'piping', 'water heater',
  'disposal', 'sump pump', 'sump', 'pump', 'flange', 'wax ring',
  'supply line', 'shut-off', 'shutoff', 'trap', 'backflow',
  'pressure reducing valve', 'prv', 'expansion tank', 'anode',
  'thermostat', 'condenser', 'evaporator', 'compressor', 'capacitor',
  'contactor', 'blower motor', 'refrigerant',
  'breaker', 'outlet', 'light switch', 'wall switch', 'gfci', 'disconnect',
  'igniter', 'ignitor', 'thermocouple', 'flame sensor', 'gas valve',
  'humidifier', 'uv light', 'zone valve', 'circulator',
  'pop up assembly', 'pop-up assembly', 'garbage disposal', 'water softener',
  'pressure tank', 'well pump', 'ejector pump', 'surge protector', 'smoke detector',
];
const REPLACE_WORDS = /\b(replac(ed|ing)|install(ed|ing)|swapp?ed|new|changed out|put in)\b/i;
const esc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

// Parts the write-up says the tech replaced/installed that don't appear on any invoice line.
// Deliberately conservative: the daily Claude review does the nuanced reading.
const NOT_DONE = /\b(recommend|quote|quoted|estimate|option|suggest|advis|will need|needs? to be|should be|would|could|consider|homeowner|customer|previous(ly)?|already|last year|prior)/i;
function summaryNotBilled(summary, itemNames) {
  const billed = itemNames.join(' | ').toLowerCase();
  const hits = [];
  const sentences = String(summary || '').split(/(?<=[.!?\n])\s+/);
  for (const k of KEYWORDS) {
    const kw = esc(k) + 's?';
    if (new RegExp('\\b' + kw + '\\b', 'i').test(billed)) continue;
    const before = new RegExp('\\b(replac(ed|ing)|install(ed|ing)|swapp?ed|changed out|put in|new)\\W+(\\w+\\W+){0,5}?' + kw + '\\b', 'i');
    const after = new RegExp('\\b' + kw + '\\W+(\\w+\\W+){0,3}?(was |were )?(replaced|installed|swapped)\\b', 'i');
    if (sentences.some((x) => (before.test(x) || after.test(x)) && !NOT_DONE.test(x))) hits.push(k);
  }
  return hits;
}

const NO_OPTION_EXEMPT = /(sold work|non-repair|install(?!.*estimate)|warranty|recall|callback|punch|permit|inspection only|follow up)/i;

function computeFlags(job, ctx) {
  const f = [];
  const add = (code, level, text) => f.push({ code, level, text });
  const d = job.detail || {};
  const m = job.materials || {};
  const invTotal = job.invoices.reduce((s, i) => s + (i.total || 0), 0);
  const items = job.invoices.flatMap((i) => i.items || []);
  const materialLines = items.filter((i) => /material|equipment/i.test(i.type));
  const poCost = job.invoices.reduce((s, i) => s + (i.poCost || 0) + (i.billCost || 0), 0);
  const matCost = Math.max(m.materialCost || 0, job.invoices.reduce((s, i) => s + (i.materialCost || 0), 0));
  const revenue = m.revenue != null ? m.revenue : (d.revenue != null ? d.revenue : invTotal);
  const type = job.jobType || '';
  const estCount = Math.max(job.estimates.length, d.estimatesCount || 0);
  const isEstimateJob = /^estimate/i.test(type);

  // --- options presented
  if (job.detail || job.estimates.length) {
    if (estCount === 0 && !NO_OPTION_EXEMPT.test(type) && !d.warranty && !d.recall) {
      const isCheck = /system check|maintenance|tune.?up/i.test(type);
      add('no_options', isEstimateJob || isCheck || d.zeroDollar ? 'warn' : 'info',
        isEstimateJob ? 'Estimate job with no estimates built' : isCheck ? 'System check with no estimate' : 'No estimate options presented');
    } else if (estCount === 1 && !NO_OPTION_EXEMPT.test(type)) {
      add('one_option', 'info', 'Only one option presented');
    }
  }

  // --- write-up vs invoice
  for (const inv of job.invoices) {
    if ((inv.total || 0) > 0 && !inv.summary) add('no_writeup', 'warn', `Invoice ${inv.invoiceNumber} billed with no tech write-up`);
  }
  // Write-up vs invoice matching is done by the daily Claude review (review_ai.js), not by keywords:
  // keyword matching could not tell "replaced the capacitor" from "recommend replacing the panel".
  const ai = job.aiReview;
  if (ai && ai.fingerprint === reviewFingerprint(job) && Array.isArray(ai.issues)) {
    ai.issues.forEach((x) => add('ai_' + (x.type || 'issue'), x.severity === 'high' ? 'warn' : 'info', x.text));
  }

  // --- costs vs billing
  if (poCost >= 25 && invTotal <= 0 && job.invoices.length) {
    const wr = d.warranty || d.recall || /warranty|recall/i.test(type);
    add('po_not_billed', wr ? 'info' : 'warn', `$${Math.round(poCost).toLocaleString()} in PO/bill costs, $0 invoiced${wr ? ' (warranty/recall)' : ''}`);
  }
  if (matCost >= 25 && job.invoices.length && materialLines.length === 0 && invTotal <= 0) {
    add('materials_not_billed', 'warn', `$${Math.round(matCost).toLocaleString()} in material cost, nothing billed`);
  }
  if (revenue >= 100 && m.grossMarginPct != null && m.grossMarginPct < 0) add('negative_margin', 'info', `Negative margin (${m.grossMarginPct}%)`);

  // --- time on site
  if (job.detail) {
    const hw = d.hoursWorked || 0;
    const med = ctx.medianHoursByType[type];
    if (hw > 0 && hw < 0.25 && revenue > 0) add('short_visit', 'info', `Only ${Math.round(hw * 60)} min on site for a billed job`);
    if (med && hw >= 3 && hw > med * 3) add('long_visit', 'info', `${hw.toFixed(1)} hrs on site — ${(hw / med).toFixed(1)}× typical for ${type}`);
  }
  return f;
}

// Changes whenever the write-up, billed items or estimates change, so the Claude review re-runs.
function reviewFingerprint(job) {
  const s = JSON.stringify([
    job.invoices.map((i) => [i.invoiceNumber, i.summary, i.total, (i.items || []).map((x) => [x.name, x.price])]),
    job.estimates.map((e) => [e.id, e.status, e.subtotal, e.summary || '']),
  ]);
  let h = 5381;
  for (let i = 0; i < s.length; i++) h = ((h * 33) ^ s.charCodeAt(i)) >>> 0;
  return h.toString(36);
}

function recomputeMonth(monthData) {
  const jobs = Object.values(monthData.jobs);
  const byType = {};
  jobs.forEach((j) => { if (j.detail && j.detail.hoursWorked > 0) (byType[j.jobType] = byType[j.jobType] || []).push(j.detail.hoursWorked); });
  const ctx = { medianHoursByType: {} };
  for (const [t, a] of Object.entries(byType)) if (a.length >= 5) ctx.medianHoursByType[t] = median(a);
  jobs.forEach((j) => { j.flags = computeFlags(j, ctx); j.medianHours = ctx.medianHoursByType[j.jobType] || null; });
  monthData.updatedAt = new Date().toISOString();
  monthData.jobCount = jobs.length;
  monthData.flaggedCount = jobs.filter((j) => j.flags.some((x) => x.level === 'warn')).length;
}

function median(a) { if (!a.length) return 0; const s = [...a].sort((x, y) => x - y); return s[Math.floor(s.length / 2)]; }

// ---------------------------------------------------------------- storage

function makeStore(opts) {
  if (opts.local) {
    const root = opts.local;
    return {
      async read(p) {
        const fp = path.join(root, p);
        if (!fs.existsSync(fp)) return { sha: null, data: null };
        return { sha: 'local', data: JSON.parse(fs.readFileSync(fp, 'utf8')) };
      },
      async write(p, sha, data) {
        const fp = path.join(root, p);
        fs.mkdirSync(path.dirname(fp), { recursive: true });
        fs.writeFileSync(fp, JSON.stringify(data));
      },
    };
  }
  const token = process.env.GH_TOKEN;
  if (!token) { console.error('GH_TOKEN is required (or use --local <repoDir>).'); process.exit(1); }
  const H = { Authorization: `Bearer ${token}`, 'User-Agent': 'lph-review-sync', Accept: 'application/vnd.github+json' };
  const api = `https://api.github.com/repos/${OWNER}/${REPO}/contents/`;
  return {
    async read(p) {
      const meta = await fetch(api + p + `?ref=${BRANCH}`, { headers: H });
      if (meta.status === 404) return { sha: null, data: null };
      if (!meta.ok) throw new Error(`GET ${p}: ${meta.status} ${await meta.text()}`);
      const j = await meta.json();
      let text;
      if (j.content && j.encoding === 'base64') text = Buffer.from(j.content, 'base64').toString('utf8');
      else {
        const raw = await fetch(api + p + `?ref=${BRANCH}`, { headers: { ...H, Accept: 'application/vnd.github.raw' } });
        if (!raw.ok) throw new Error(`GET raw ${p}: ${raw.status}`);
        text = await raw.text();
      }
      return { sha: j.sha, data: JSON.parse(text) };
    },
    async write(p, sha, data, message) {
      const body = { message, branch: BRANCH, content: Buffer.from(JSON.stringify(data)).toString('base64') };
      if (sha) body.sha = sha;
      const res = await fetch(api + p, { method: 'PUT', headers: { ...H, 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
      if (!res.ok) throw new Error(`PUT ${p}: ${res.status} ${await res.text()}`);
    },
  };
}

// ---------------------------------------------------------------- main

async function main() {
  const args = process.argv.slice(2);
  const opts = { local: null, context: 'manual', files: [] };
  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--local') opts.local = args[++i];
    else if (args[i] === '--context') opts.context = args[++i];
    else opts.files.push(args[i]);
  }
  const today = new Date().toISOString().slice(0, 10);
  const store = makeStore(opts);
  const log = { stage: 'started', run_context: opts.context, at: new Date().toISOString(), reports: [], errors: [] };

  if (opts.files.length === 0) {
    log.stage = 'no_files';
    console.log('No report files given — nothing to sync.');
    await writeLog(store, log);
    return;
  }

  // 1. parse
  const updates = [];
  for (const f of opts.files) {
    const label = path.basename(f);
    try {
      const { rows, filters } = readWorkbook(f);
      if (!rows.length) throw new Error('empty sheet');
      const cols = makeCols(rows[0]);
      const kind = detect(cols);
      if (!kind) { log.reports.push({ file: label, kind: null, note: 'unrecognized header row' }); console.log(`${label}: unrecognized report, skipped.`); continue; }
      const u = PARSERS[kind](rows.slice(1), cols);
      const note = [];
      if (kind === 'estimates' && !cols.has('estimate summary') && !cols.has('summary')) note.push('no Estimate Summary column');
      if (kind === 'estimates' && /sold/i.test(filters['Date Type'] || '')) note.push('filtered by Sold On (only sold estimates)');
      if (kind === 'lineItems' && cols.find('purchase order costs') === -1) note.push('no Purchase Order Costs column');
      updates.push(...u);
      const jobs = new Set(u.map((x) => x.jobNumber)).size;
      log.reports.push({ file: label, kind, rows: rows.length - 1, records: u.length, jobs, range: filters['Date Range'] || [filters['Invoices From'], filters['Invoices To']].filter(Boolean).join(' - '), note: note.join('; ') || undefined });
      console.log(`${label}: ${kind} — ${u.length} record(s) across ${jobs} job(s)${note.length ? ' [' + note.join('; ') + ']' : ''}`);
    } catch (err) {
      log.errors.push(`${label}: ${err.message}`);
      console.log(`${label}: ERROR ${err.message}`);
    }
  }
  if (!updates.length) {
    log.stage = log.errors.length ? 'failed' : 'no_data';
    await writeLog(store, log);
    if (log.errors.length) process.exit(1);
    return;
  }

  // 2. load index, month files touched, pending
  const idxRes = await store.read(`${DIR}/index.json`);
  const index = idxRes.data || { months: [], jobMonth: {}, updatedAt: null };
  const pendRes = await store.read(`${DIR}/pending.json`);
  const pending = (pendRes.data && pendRes.data.jobs) || {};
  const monthCache = {};
  async function loadMonth(mo) {
    if (!monthCache[mo]) {
      const r = await store.read(`${DIR}/${mo}.json`);
      monthCache[mo] = { sha: r.sha, data: r.data || { month: mo, jobs: {} }, dirty: false };
    }
    return monthCache[mo];
  }

  // 3. group updates by job and apply
  const byJob = {};
  for (const u of updates) (byJob[u.jobNumber] = byJob[u.jobNumber] || []).push(u);
  // detail/materials first so the completion month is known before invoices/estimates land
  const order = { detail: 0, materials: 1, invoice: 2, estimate: 3 };
  let added = 0, updated = 0;
  for (const [jobNo, us] of Object.entries(byJob)) {
    us.sort((a, b) => order[a.section] - order[b.section]);
    let mo = index.jobMonth[jobNo] || null;
    let job;
    if (mo) job = (await loadMonth(mo)).data.jobs[jobNo];
    if (!job && pending[jobNo]) job = pending[jobNo];
    if (!job) { job = emptyJob(jobNo); added++; } else updated++;
    for (const u of us) applyUpdate(job, u, today);
    const newMo = job.completionDate ? job.completionDate.slice(0, 7) : null;
    if (mo && newMo && newMo !== mo) { const old = await loadMonth(mo); delete old.data.jobs[jobNo]; old.dirty = true; mo = newMo; }
    if (!mo && newMo) mo = newMo;
    if (mo) {
      const bucket = await loadMonth(mo);
      bucket.data.jobs[jobNo] = job; bucket.dirty = true;
      index.jobMonth[jobNo] = mo;
      delete pending[jobNo];
    } else {
      job.pendingSince = job.pendingSince || today;
      pending[jobNo] = job;
    }
  }

  // 4. recompute flags for every touched month (medians are per month)
  for (const [mo, b] of Object.entries(monthCache)) {
    if (!b.dirty) continue;
    recomputeMonth(b.data);
    if (!index.months.includes(mo)) index.months.push(mo);
  }
  index.months.sort().reverse();

  // expire old pending entries
  const cutoff = new Date(Date.now() - PENDING_TTL_DAYS * 86400000).toISOString().slice(0, 10);
  for (const [k, v] of Object.entries(pending)) if (v.pendingSince && v.pendingSince < cutoff) delete pending[k];

  // 5. write
  const msg = `Review sync (${opts.context}): ${Object.keys(byJob).length} job(s) (+${added}/~${updated})`;
  for (const [mo, b] of Object.entries(monthCache)) {
    if (b.dirty) await store.write(`${DIR}/${mo}.json`, b.sha, b.data, `${msg} [${mo}]`);
  }
  await store.write(`${DIR}/pending.json`, pendRes.sha, { jobs: pending, updatedAt: new Date().toISOString() }, `${msg} [pending]`);
  index.updatedAt = new Date().toISOString();
  index.lastSync = { at: index.updatedAt, context: opts.context, reports: log.reports };
  await store.write(`${DIR}/index.json`, idxRes.sha, index, `${msg} [index]`);

  log.stage = log.errors.length ? 'partial' : 'done';
  log.jobs = { touched: Object.keys(byJob).length, added, updated, pending: Object.keys(pending).length };
  await writeLog(store, log);
  console.log(`Review sync ${log.stage}: ${Object.keys(byJob).length} job(s) touched (+${added} new / ~${updated} updated), ${Object.keys(pending).length} waiting for a completion date.`);
  if (log.errors.length) process.exit(1);
}

async function writeLog(store, log) {
  try {
    const r = await store.read(`${DIR}/_sync_log.json`);
    const hist = Array.isArray(r.data) ? r.data : [];
    hist.unshift(log);
    await store.write(`${DIR}/_sync_log.json`, r.sha, hist.slice(0, 60), `Review sync log: ${log.stage} (${log.run_context})`);
  } catch (e) { console.error('Could not write sync log:', e.message); }
}

if (require.main === module) main().catch((e) => { console.error('FATAL', e.stack || e.message); process.exit(1); });

module.exports = { readWorkbook, makeCols, detect, PARSERS, computeFlags, summaryNotBilled, reviewFingerprint, recomputeMonth, median, makeStore, DIR };
