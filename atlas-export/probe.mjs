#!/usr/bin/env node
// Probe: run this FIRST, as soon as you have an API token.
//
// It answers two questions in under a minute:
//   1. Does the token work at all?
//   2. Does /unit/{id} return the actual unit planner content (essential
//      questions, knowledge, skills, activities…) or only the metadata
//      documented in the spec? This is the make-or-break question for the
//      whole export — the official docs only promise metadata.
//
// Raw responses are saved to output/probe/ so we can inspect exactly what
// the API returns and tune the renderer to it.
//
// Usage: npm run probe

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { apiGet, apiGetAll, requireToken } from './lib/api.mjs';

const ROOT = path.dirname(fileURLToPath(import.meta.url));
const PROBE_DIR = path.join(ROOT, 'output', 'probe');

// Fields the official spec documents on a Unit (including their typo,
// "curriclummap_id"). Anything beyond these is undocumented planner content.
const DOCUMENTED_UNIT_FIELDS = new Set([
  'id',
  'curriclummap_id',
  'curriculummap_id',
  'title',
  'time_frame',
  'last_updated',
  'start_date',
  'end_date',
  'url',
]);

function save(name, data) {
  fs.mkdirSync(PROBE_DIR, { recursive: true });
  const file = path.join(PROBE_DIR, name);
  fs.writeFileSync(file, JSON.stringify(data, null, 2));
  console.log(`  saved ${path.relative(ROOT, file)}`);
}

const token = requireToken();

console.log('1/4  GET /teacher — token check…');
const teachers = await apiGetAll('/teacher', {}, token);
save('teachers.json', teachers);
console.log(`  ✅ token works — ${teachers.length} teachers visible\n`);

console.log('2/4  GET /curriculummap — first page…');
const mapsEnvelope = await apiGet('/curriculummap', {}, token);
save('curriculummap-page1.json', mapsEnvelope);
const maps = Array.isArray(mapsEnvelope?.data) ? mapsEnvelope.data : [];
console.log(`  ${maps.length} maps on page 1` + (mapsEnvelope?.page?.total ? ` (of ${mapsEnvelope.page.total} total)` : ''));

const mapWithUnits = maps.find((m) => Array.isArray(m.units) && m.units.length > 0);
let unitId = mapWithUnits?.units?.[0]?.id;
if (!unitId) {
  console.log('  no embedded units on page-1 maps — falling back to GET /unit…');
  const unitList = await apiGet('/unit', {}, token);
  save('unit-list-page1.json', unitList);
  unitId = unitList?.data?.[0]?.id;
}
if (!unitId) {
  console.error('  ❌ could not find any unit to probe. Inspect output/probe/*.json.');
  process.exit(1);
}

console.log(`\n3/4  GET /unit/${unitId} — THE key question…`);
const unitDetail = await apiGetAll(`/unit/${unitId}`, {}, token);
const unit = unitDetail.length === 1 ? unitDetail[0] : unitDetail;
save(`unit-${unitId}.json`, unit);

const keys = Object.keys(Array.isArray(unit) ? unit[0] ?? {} : unit);
const extras = keys.filter((k) => !DOCUMENTED_UNIT_FIELDS.has(k));
console.log(`  fields returned: ${keys.join(', ')}`);
if (extras.length > 0) {
  console.log(`\n  ✅ VERDICT: /unit/{id} returns ${extras.length} undocumented field(s):`);
  console.log(`     ${extras.join(', ')}`);
  console.log('     These are likely the planner content — the API route is viable.');
  console.log('     Send these field names + output/probe/unit-*.json to Claude to tune the renderer.');
} else {
  console.log('\n  ⚠️  VERDICT: only documented metadata came back — NO planner content.');
  console.log('     The API alone cannot archive the unit prose. Fallbacks: bulk PDF');
  console.log('     export, a Faria-provided dump, or a browser-session scrape.');
}

console.log(`\n4/4  GET /assessment?unit_id=${unitId} — sample assessment…`);
const assessments = await apiGetAll('/assessment', { unit_id: unitId }, token);
save(`assessments-unit-${unitId}.json`, assessments);
if (assessments.length > 0) {
  console.log(`  ${assessments.length} assessment(s); fields: ${Object.keys(assessments[0]).join(', ')}`);
} else {
  console.log('  no assessments on this unit (fine — others may have them)');
}

console.log('\nProbe complete. Review output/probe/ before running the full extract.');
