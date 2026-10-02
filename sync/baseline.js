// Builds "knowledge" from historical ServiceTitan exports (Invoice Line Items + Job Detail).
// Historical jobs are NOT added to the review data -- only aggregate statistics are written:
//   data/baseline/task_times.json   per pricebook task: samples [tech, dept, hours, yyyy-mm] from single-task jobs
//   data/baseline/material_rules.json  per task: materials that are normally recorded with it
// Usage: node baseline.js build <outDir> <lineItems.xlsx...> -- <jobDetail.xlsx...>
const fs = require('fs'), path = require('path');
const L = require('./review_sync.js');

const { workTasks, repairTasks, deptOf: dept, NOT_TASK, WRAPPER } = L;
const isMat = (x) => /material/i.test(x.type || '');
const JUNK_MAT = /^(misc|materials? misc|misc spiff|spiff)/i;

function load(files) {
  const out = {};
  for (const f of files) {
    const { rows } = L.readWorkbook(f); const cols = L.makeCols(rows[0]); const k = L.detect(cols);
    for (const r of L.PARSERS[k](rows.slice(1), cols)) {
      const j = out[r.jobNumber] = out[r.jobNumber] || { jobNumber: r.jobNumber, invoices: [] };
      Object.assign(j, Object.fromEntries(Object.entries(r.base || {}).filter(([, v]) => v != null && v !== '')));
      if (r.section === 'invoice') j.invoices.push(r.data); else if (r.section === 'detail') j.detail = r.data;
    }
  }
  return out;
}

function build(jobs, pb) {
  const pbSvc = new Map(pb.services.map((s) => [s.code, s]));
  const pbMat = new Map(pb.materials.map((m) => [m.code, m]));
  const links = new Map(Array.isArray(pb.links) ? pb.links : Object.entries(pb.links || {}));
  const times = {}, withTask = {}, matWith = {};
  for (const j of Object.values(jobs)) {
    if (!j.invoices.length) continue;
    const tasks = workTasks(j);
    const mats = new Set(j.invoices.flatMap((i) => i.items || []).filter(isMat).map((x) => x.code).filter(Boolean));
    for (const t of tasks) {
      withTask[t.code] = (withTask[t.code] || 0) + 1;
      const m = matWith[t.code] = matWith[t.code] || {};
      mats.forEach((c) => { m[c] = (m[c] || 0) + 1; });
    }
    const tt = L.taskTimeOf(j);
    if (tt) {
      const hrs = tt.hrs, t = { code: tt.code, name: tt.name };
      const e = times[t.code] = times[t.code] || { code: t.code, name: (pbSvc.get(t.code) || {}).name || t.name, pbHours: (pbSvc.get(t.code) || {}).hours || null, samples: [] };
      e.samples.push([j.primaryTech, dept(j.businessUnit), hrs, (j.completionDate || '').slice(0, 7), tt.multi]);
    }
  }
  const counts = {};
  for (const [code, n] of Object.entries(withTask)) counts[code] = { n, m: matWith[code] || {} };
  const rules = L.deriveRules(counts, pb);
  return { times, rules, counts };
}

if (require.main === module) {
  const [cmd, outDir, ...rest] = process.argv.slice(2);
  if (cmd !== 'build') { console.error('usage: node baseline.js build <outDir> <files...>'); process.exit(1); }
  const pb = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'data', 'pricebook.json'), 'utf8'));
  const jobs = load(rest);
  const { times, rules, counts } = build(jobs, pb);
  fs.mkdirSync(outDir, { recursive: true });
  const meta = { builtAt: new Date().toISOString(), jobs: Object.keys(jobs).length, note: 'Aggregates from historical exports; no individual jobs stored.' };
  fs.writeFileSync(path.join(outDir, 'task_times.json'), JSON.stringify({ meta, tasks: times }));
  fs.writeFileSync(path.join(outDir, 'material_rules.json'), JSON.stringify({ meta, rules, counts }));
  console.log('jobs', meta.jobs, 'tasks with times', Object.keys(times).length, 'samples', Object.values(times).reduce((s, t) => s + t.samples.length, 0), 'material rules', Object.keys(rules).length);
}
module.exports = { load, build };
