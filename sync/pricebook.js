// Pricebook helpers for the Job Review tool.
//
// 1) Build a compact data/pricebook.json from a ServiceTitan Pricebook export:
//      node pricebook.js build <Pricebook_export.xlsx> <out.json>
//    (Re-run whenever the pricebook changes meaningfully.)
//
// 2) Used by review_sync.js / review_ai.js:
//      parseGT(code)          -> { hours, materials } for "General Time" (Grid Task) codes GT-HHHH-MMMM
//      candidates(pb, text)   -> the pricebook tasks whose name/description best match a tech's description
//
// General Time codes: GT-HHHH-MMMM = HHHH/100 labor hours + $MMMM material allowance, chosen by the tech.
// Their price follows the same formula as regular tasks (~$770 per labor hour + ~1.67x material cost),
// so the difference vs. the "true" task comes from the hours/materials the tech picked.

const fs = require('fs');

function parseGT(code) {
  const m = String(code || '').match(/^GT-(\d{4})-(\d{4})$/i);
  if (!m) return null;
  return { hours: +m[1] / 100, materials: +m[2] };
}

const STOP = new Set(('a an and the of on in at to for with w/ w by or from into is are was were be been this that these those it its as per new old ' +
  'install installed installing replace replaced replacing replacement repair repaired repairing remove removed removing lph technician tech ' +
  'customer homeowner home house would will need needs needed provide provided existing unit system one two 1 2 3 4 5 6 7 8 9 0 test tested ' +
  'check checked proper properly ensure operation working work job all up out off then also inch').split(/\s+/));
const tok = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').split(' ').filter((w) => w.length > 1 && !STOP.has(w))
  .map((w) => w.replace(/(ies)$/, 'y').replace(/(es|s)$/, ''));

function build(xlsxPath, outPath) {
  const XLSX = require('xlsx');
  const wb = XLSX.readFile(xlsxPath);
  const sheet = (n) => (wb.Sheets[n] ? XLSX.utils.sheet_to_json(wb.Sheets[n], { defval: null }) : []);
  const active = (r) => r.Active === 1 || r.Active === true || r.Active === '1' || r.Active === 'True';
  const num = (v) => Math.round((+v || 0) * 100) / 100;
  const clip = (s, n) => String(s || '').replace(/\s+/g, ' ').trim().slice(0, n);
  const services = sheet('Services').filter(active).map((r) => ({
    code: String(r.Code), name: clip(r.Name, 120), desc: clip(r.Description, 260), cat: r['Category.Name'] || '',
    price: num(r.StaticPrice), member: num(r.StaticMemberPrice), addOn: num(r.StaticAddOnPrice), addOnMember: num(r.StaticAddOnMemberPrice),
    hours: num(r.Hours), materialCost: num(r.MaterialCost),
  }));
  const materials = sheet('Materials & Other Costs').filter(active).map((r) => ({
    code: String(r.Code), name: clip(r.Name, 120), cost: num(r.Cost), price: num(r.Price), member: num(r.MemberPrice),
  }));
  const equipment = sheet('Equipment').filter(active).map((r) => ({
    code: String(r.Code), name: clip(r.Name, 120), cost: num(r.Cost), price: num(r.Price), member: num(r.MemberPrice),
  }));
  const links = {};
  sheet('ServiceMaterialLinks').filter(active).forEach((r) => { (links[r['Service.Code']] = links[r['Service.Code']] || []).push([String(r['Material.Code']), +r.Quantity || 1]); });
  const out = { builtAt: new Date().toISOString(), source: require('path').basename(xlsxPath), services, materials, equipment, links };
  fs.writeFileSync(outPath, JSON.stringify(out));
  console.log(`pricebook: ${services.length} services, ${materials.length} materials, ${equipment.length} equipment, ${Object.keys(links).length} tasks with material links -> ${outPath}`);
}

// Rank real (non-GT) tasks by word overlap with the tech's description.
let _index = null;
function candidates(pb, text, n = 12) {
  if (!_index || _index.pb !== pb) {
    const df = {};
    const docs = pb.services.filter((s) => !/^GT-/i.test(s.code) && s.price > 0 && !/^(invoice|coupons)$/i.test(s.cat)).map((s) => {
      const nameT = new Set(tok(s.name)), descT = new Set(tok(s.desc + ' ' + s.cat));
      new Set([...nameT, ...descT]).forEach((w) => { df[w] = (df[w] || 0) + 1; });
      return { s, nameT, descT };
    });
    _index = { pb, docs, df, N: docs.length };
  }
  const q = new Set(tok(text));
  if (!q.size) return [];
  const idf = (w) => Math.log((_index.N + 1) / ((_index.df[w] || 0) + 1));
  return _index.docs.map((d) => {
    let sc = 0;
    q.forEach((w) => { if (d.nameT.has(w)) sc += 3 * idf(w); else if (d.descT.has(w)) sc += idf(w); });
    return { d, sc };
  }).filter((x) => x.sc > 0).sort((a, b) => b.sc - a.sc).slice(0, n).map((x) => x.d.s);
}

if (require.main === module) {
  const [cmd, a, b] = process.argv.slice(2);
  if (cmd === 'build' && a && b) build(a, b);
  else { console.error('Usage: node pricebook.js build <Pricebook_export.xlsx> <out.json>'); process.exit(1); }
}

module.exports = { parseGT, candidates, build };
