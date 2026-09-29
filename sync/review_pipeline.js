// One-command wrapper for the daily Job Review run (used by the scheduled Claude task).
//
//   GH_TOKEN=... node review_pipeline.js prepare <gmail RAW message files...>
//       -> extracts the .xlsx attachments, runs review_sync.js (merge + rule flags, writes to GitHub),
//          then writes the Claude review queue to ./queue.json and prints a summary.
//   GH_TOKEN=... node review_pipeline.js finish <results.json>
//       -> saves Claude's review results (review_ai.js apply).
//
// Run from a folder containing gmail_extract.js, review_sync.js, review_ai.js, pricebook.js and
// node_modules/xlsx (npm i xlsx@0.18.5).

const { spawnSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const here = __dirname;
const run = (args, opts = {}) => {
  const r = spawnSync(process.execPath, args, { cwd: here, encoding: 'utf8', env: process.env, maxBuffer: 64 * 1024 * 1024, ...opts });
  return { code: r.status, out: (r.stdout || '').trim(), err: (r.stderr || '').trim() };
};

const [cmd, ...rest] = process.argv.slice(2);
if (!process.env.GH_TOKEN) { console.error('GH_TOKEN is required'); process.exit(1); }

if (cmd === 'prepare') {
  const outDir = path.join(here, 'xlsx');
  fs.rmSync(outDir, { recursive: true, force: true });
  let files = [];
  if (rest.length) {
    const ex = run(['gmail_extract.js', outDir, ...rest]);
    console.log(ex.out);
    if (ex.err) console.log(ex.err);
    files = ex.out.split('\n').filter((l) => l.startsWith('EXTRACTED ')).map((l) => l.slice(10).replace(/ \(\d+ bytes.*$/, ''));
  }
  const sync = run(['review_sync.js', '--context', 'scheduled', ...files]);
  console.log(sync.out); if (sync.err) console.log(sync.err);
  if (sync.code !== 0) { console.log('REVIEW_SYNC_FAILED'); process.exit(1); }
  const q = run(['review_ai.js', 'queue', '--max', '80']);
  if (q.code !== 0) { console.log(q.err); console.log('QUEUE_FAILED'); process.exit(1); }
  fs.writeFileSync(path.join(here, 'queue.json'), q.out);
  const qj = JSON.parse(q.out || '{"count":0,"jobs":[]}');
  const gt = qj.jobs.filter((j) => j.generalTime).length, gap = qj.jobs.filter((j) => j.materialsGap).length;
  console.log(`QUEUE ${qj.count} job(s) to review (${gt} with General Time, ${gap} with materials gap), ${qj.remainingAfterThis || 0} more waiting -> ${path.join(here, 'queue.json')}`);
} else if (cmd === 'finish' && rest[0]) {
  const a = run(['review_ai.js', 'apply', path.resolve(rest[0])]);
  console.log(a.out); if (a.err) console.log(a.err);
  process.exit(a.code || 0);
} else {
  console.error('Usage: review_pipeline.js prepare <raw files...> | finish <results.json>');
  process.exit(1);
}
