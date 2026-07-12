#!/usr/bin/env node
// Phase 1: pull everything the Atlas API exposes into output/raw/ as JSON.
//
// This raw archive is the canonical, lossless copy of the curriculum data —
// keep it even after the Markdown render looks good. The renderer (render.mjs)
// only ever reads from output/raw/, so the API can be hit once and then
// switched off forever.
//
// Resumable: every file that already exists in output/raw/ is skipped, so an
// interrupted run continues where it left off.
//
// Usage:
//   npm run extract                full pull
//   node extract.mjs --limit 10    only fetch detail for the first 10 units (testing)
//   node extract.mjs --force       refetch everything, ignoring cached files

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { apiGetAll, requireToken, sleep } from './lib/api.mjs';

const ROOT = path.dirname(fileURLToPath(import.meta.url));
const RAW = path.join(ROOT, 'output', 'raw');

const args = process.argv.slice(2);
const FORCE = args.includes('--force');
const limitIx = args.indexOf('--limit');
const LIMIT = limitIx >= 0 ? Number(args[limitIx + 1]) : Infinity;
if (Number.isNaN(LIMIT)) {
  console.error('--limit needs a number, e.g. --limit 10');
  process.exit(1);
}

const token = requireToken();

// Fetch via `fetcher` unless the file is already on disk (resumability).
async function cached(rel, fetcher) {
  const file = path.join(RAW, rel);
  if (!FORCE && fs.existsSync(file)) {
    return JSON.parse(fs.readFileSync(file, 'utf8'));
  }
  const data = await fetcher();
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, JSON.stringify(data, null, 2));
  return data;
}

console.log(`Extracting to ${path.relative(process.cwd(), RAW)}${FORCE ? ' (--force)' : ''}\n`);

// --- Site taxonomy (cheap, one page each) ----------------------------------
const schools = await cached('schools.json', () => apiGetAll('/schools', {}, token));
console.log(`schools:      ${schools.length}`);
const subjects = await cached('subjects.json', () => apiGetAll('/subjects', {}, token));
console.log(`subjects:     ${subjects.length}`);
const grades = await cached('grades.json', () => apiGetAll('/grades', {}, token));
console.log(`grades:       ${grades.length}`);
const teachers = await cached('teachers.json', () => apiGetAll('/teacher', {}, token));
console.log(`teachers:     ${teachers.length}`);
const courses = await cached('courses.json', () => apiGetAll('/courses', {}, token));
console.log(`courses:      ${courses.length}`);

// --- Curriculum maps (each embeds its units + teachers) ---------------------
const maps = await cached('maps.json', () => apiGetAll('/curriculummap', {}, token));
console.log(`maps:         ${maps.length}`);

// --- Unit index + per-unit detail -------------------------------------------
// The site-wide /unit list and the units embedded in maps should agree, but we
// union them to be safe — detail fetches are where planner content lives.
const unitIndex = await cached('units.json', () => apiGetAll('/unit', {}, token));
console.log(`units (list): ${unitIndex.length}`);

const unitIds = [
  ...new Set(
    [
      ...unitIndex.map((u) => u.id),
      ...maps.flatMap((m) => (Array.isArray(m.units) ? m.units : [])).map((u) => u.id),
    ].filter((id) => id !== undefined && id !== null)
  ),
];

const toFetch = unitIds.slice(0, LIMIT === Infinity ? unitIds.length : LIMIT);
console.log(`unit details: fetching ${toFetch.length} of ${unitIds.length}…`);
let done = 0;
for (const id of toFetch) {
  const file = `units/${id}.json`;
  const existed = !FORCE && fs.existsSync(path.join(RAW, file));
  await cached(file, async () => {
    const detail = await apiGetAll(`/unit/${id}`, {}, token);
    return detail.length === 1 ? detail[0] : detail;
  });
  done++;
  if (!existed) await sleep(150); // politeness gap only when we actually hit the API
  if (done % 25 === 0 || done === toFetch.length) {
    console.log(`  ${done}/${toFetch.length}`);
  }
}

// --- Assessments + attachments (site-wide, paginated) ------------------------
const assessments = await cached('assessments.json', () => apiGetAll('/assessment', {}, token));
console.log(`assessments:  ${assessments.length}`);
const attachments = await cached('attachments.json', () => apiGetAll('/attachment', {}, token));
console.log(`attachments:  ${attachments.length}`);

console.log('\nExtract complete. Next: npm run render');
if (LIMIT !== Infinity) {
  console.log(`(note: --limit ${LIMIT} was set — run without it for the full unit pull)`);
}
