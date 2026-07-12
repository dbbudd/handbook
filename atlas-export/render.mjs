#!/usr/bin/env node
// Phase 2: render the raw JSON archive into a browsable Markdown hierarchy.
//
//   output/markdown/<Division>/<Course>/<Map>.md   one file per curriculum map
//   output/markdown/INDEX.md                       linked table of contents
//   output/links.csv                               every hyperlink found, with context
//
// Reads ONLY from output/raw/ (run extract.mjs first) — never touches the API,
// so it can be re-run freely while iterating on the template.
//
// Usage: npm run render

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.dirname(fileURLToPath(import.meta.url));
const RAW = path.join(ROOT, 'output', 'raw');
const OUT = path.join(ROOT, 'output', 'markdown');

let TurndownService;
try {
  ({ default: TurndownService } = await import('turndown'));
} catch {
  console.error('turndown is not installed — run `npm install` in atlas-export/ first.');
  process.exit(1);
}
const turndown = new TurndownService({ headingStyle: 'atx', bulletListMarker: '-' });

// ---------------------------------------------------------------------------
// Config
// ---------------------------------------------------------------------------

// Atlas "schools" become the top-level division folders. If the Atlas names
// aren't the folder names you want, rename them here.
const DIVISION_NAMES = {
  // 'HKIS High School': 'High School',
  // 'HKIS Middle School': 'Middle School',
};

// Unit fields treated as metadata (shown on the unit's header line). Every
// OTHER field on a unit is rendered as its own content section — so if the
// probe reveals undocumented planner fields (essential questions, skills…),
// they appear in the output automatically without code changes.
// "curriclummap_id" is a typo in Atlas's own spec; kept in case the live API uses it.
const UNIT_META_FIELDS = new Set([
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

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function loadJson(rel, fallback) {
  const file = path.join(RAW, rel);
  if (!fs.existsSync(file)) return fallback;
  return JSON.parse(fs.readFileSync(file, 'utf8'));
}

// Filesystem-safe but human-readable name (keeps spaces and case).
function sanitizeName(name, max = 110) {
  const clean = String(name ?? '')
    .replace(/[/\\:*?"<>|#%{}$!@`=+]/g, '-')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/\.+$/, '');
  return (clean || 'Untitled').slice(0, max).trim();
}

function prettyFieldName(key) {
  return String(key)
    .replace(/[_-]+/g, ' ')
    .replace(/([a-z])([A-Z])/g, '$1 $2')
    .replace(/\b\w/g, (c) => c.toUpperCase())
    .trim();
}

const looksLikeHtml = (s) => /<\/?[a-z][^>]*>/i.test(s);

function toMd(value) {
  if (value === null || value === undefined) return '';
  const s = String(value);
  return looksLikeHtml(s) ? turndown.turndown(s).trim() : s.trim();
}

// Render any JSON value as Markdown — strings (HTML-aware), arrays, objects.
function valueToMd(value) {
  if (value === null || value === undefined || value === '') return '';
  if (Array.isArray(value)) {
    if (value.length === 0) return '';
    if (value.every((v) => typeof v !== 'object' || v === null)) {
      return value.map((v) => `- ${toMd(v)}`).join('\n');
    }
    return '```json\n' + JSON.stringify(value, null, 2) + '\n```';
  }
  if (typeof value === 'object') {
    return '```json\n' + JSON.stringify(value, null, 2) + '\n```';
  }
  return toMd(value);
}

// --- Hyperlink collection ----------------------------------------------------

const allLinks = []; // { division, course, map, unit, text, url }

// Scans `markdown` for links, records each into the global inventory AND the
// per-document `sink` (used to build that document's hyperlink appendix).
function collectLinks(markdown, ctx, sink) {
  const found = [];
  const mdLink = /\[([^\]]*)\]\(<?(https?:\/\/[^)\s>]+)>?\)/g;
  let m;
  while ((m = mdLink.exec(markdown)) !== null) {
    found.push({ text: m[1].trim(), url: m[2] });
  }
  // Bare URLs not already captured as markdown links
  const captured = new Set(found.map((l) => l.url));
  const bareUrl = /https?:\/\/[^\s<>"')\]]+/g;
  while ((m = bareUrl.exec(markdown)) !== null) {
    const url = m[0].replace(/[.,;:]+$/, '');
    if (!captured.has(url)) {
      captured.add(url);
      found.push({ text: '', url });
    }
  }
  for (const l of found) {
    const record = { ...ctx, ...l };
    allLinks.push(record);
    if (sink) sink.push(record);
  }
  return found;
}

function csvEscape(v) {
  const s = String(v ?? '');
  return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

// ---------------------------------------------------------------------------
// Load the raw archive
// ---------------------------------------------------------------------------

const maps = loadJson('maps.json', null);
if (!maps) {
  console.error('output/raw/maps.json not found — run `npm run extract` first.');
  process.exit(1);
}
const courses = loadJson('courses.json', []);
const assessments = loadJson('assessments.json', []);
const attachments = loadJson('attachments.json', []);

const coursesById = new Map(courses.map((c) => [c.id, c]));

// Per-unit detail files (the ones that may carry planner content)
const unitDetailById = new Map();
const unitsDir = path.join(RAW, 'units');
if (fs.existsSync(unitsDir)) {
  for (const f of fs.readdirSync(unitsDir)) {
    if (!f.endsWith('.json')) continue;
    const unit = JSON.parse(fs.readFileSync(path.join(unitsDir, f), 'utf8'));
    if (unit && unit.id !== undefined) unitDetailById.set(unit.id, unit);
  }
}

// Index assessments / attachments by unit and by map
const assessmentsByUnit = new Map();
const assessmentsByMap = new Map();
for (const a of assessments) {
  if (a.unit_id) {
    if (!assessmentsByUnit.has(a.unit_id)) assessmentsByUnit.set(a.unit_id, []);
    assessmentsByUnit.get(a.unit_id).push(a);
  } else if (a.curriculummap_id) {
    if (!assessmentsByMap.has(a.curriculummap_id)) assessmentsByMap.set(a.curriculummap_id, []);
    assessmentsByMap.get(a.curriculummap_id).push(a);
  }
}

const attachmentsByUnit = new Map();
const looseAttachments = [];
for (const att of attachments) {
  // Spec gives attachments ItemType/ItemID rather than unit_id; handle both.
  const unitId =
    att.unit_id ??
    (String(att.ItemType ?? '').toLowerCase().includes('unit') ? Number(att.ItemID) : null);
  if (unitId) {
    if (!attachmentsByUnit.has(unitId)) attachmentsByUnit.set(unitId, []);
    attachmentsByUnit.get(unitId).push(att);
  } else {
    looseAttachments.push(att);
  }
}

// ---------------------------------------------------------------------------
// ███ MARKDOWN TEMPLATE ███
//
// renderMap() below defines the layout of every per-map .md file. It is a
// sensible default — REPLACE its body with the agreed format once Daniel
// provides the example document. Everything above/below this block (hierarchy,
// link collection, INDEX) is independent of the template.
// ---------------------------------------------------------------------------

function renderUnit(unit, index, ctx, sink) {
  const lines = [];
  const title = unit.title || `Unit ${unit.id}`;
  lines.push(`### ${index + 1}. ${title}`);

  const meta = [];
  if (unit.time_frame) meta.push(`**Timeframe:** ${unit.time_frame}`);
  if (unit.start_date || unit.end_date) {
    meta.push(`**Dates:** ${unit.start_date ?? '?'} – ${unit.end_date ?? '?'}`);
  }
  if (unit.last_updated) meta.push(`**Last updated:** ${unit.last_updated}`);
  if (unit.url) meta.push(`[View in Atlas](${unit.url})`);
  if (meta.length) lines.push(meta.join(' · '));

  // Planner content: every non-meta field, rendered generically so that
  // whatever the live API returns shows up without code changes.
  for (const [key, value] of Object.entries(unit)) {
    if (UNIT_META_FIELDS.has(key)) continue;
    const md = valueToMd(value);
    if (!md) continue;
    lines.push(`#### ${prettyFieldName(key)}`);
    lines.push(md);
  }

  // Assessments for this unit
  const unitAssessments = assessmentsByUnit.get(unit.id) ?? [];
  if (unitAssessments.length) {
    lines.push('#### Assessments');
    for (const a of unitAssessments) {
      lines.push(`**${a.title ?? 'Untitled assessment'}**` +
        [a.type && ` _(${a.type}${a.method ? ` · ${a.method}` : ''})_`].filter(Boolean).join(''));
      if (Array.isArray(a.standards) && a.standards.length) {
        lines.push('Standards:');
        lines.push(a.standards.map((s) => `- ${toMd(s)}`).join('\n'));
      }
      const content = toMd(a.content);
      if (content) lines.push(content);
    }
  }

  // Attachments for this unit
  const unitAttachments = attachmentsByUnit.get(unit.id) ?? [];
  if (unitAttachments.length) {
    lines.push('#### Attachments');
    lines.push(
      unitAttachments
        .map((att) => `- [${att.title || att.url}](${att.url})${att.AttachmentType ? ` _(${att.AttachmentType})_` : ''}`)
        .join('\n')
    );
  }

  const md = lines.join('\n\n');
  collectLinks(md, { ...ctx, unit: title }, sink);
  return md;
}

function renderMap(map, courseName, divisionName) {
  const ctx = { division: divisionName, course: courseName, map: map.title ?? `Map ${map.id}`, unit: '' };
  const sink = []; // links found anywhere in this document

  const frontMatter = [
    '---',
    `title: ${JSON.stringify(map.title ?? '')}`,
    `course: ${JSON.stringify(courseName)}`,
    `division: ${JSON.stringify(divisionName)}`,
    `subject: ${JSON.stringify(map.subject ?? '')}`,
    `grade: ${JSON.stringify(map.grade ?? '')}`,
    `map_type: ${JSON.stringify(map.map_type ?? '')}`,
    `atlas_map_id: ${map.id}`,
    `last_updated: ${JSON.stringify(map.last_updated ?? '')}`,
    `teachers: ${JSON.stringify((map.teachers ?? []).map((t) => `${t.first_name ?? ''} ${t.last_name ?? ''}`.trim()))}`,
    `atlas_url: ${JSON.stringify(map.url ?? '')}`,
    '---',
  ].join('\n');

  const headerLines = [frontMatter, `# ${map.title ?? courseName}`];
  const overview = [
    map.subject && `**Subject:** ${map.subject}`,
    map.grade && `**Grade:** ${map.grade}`,
    map.school && `**Division:** ${map.school}`,
    (map.teachers ?? []).length &&
      `**Teachers:** ${map.teachers.map((t) => `${t.first_name ?? ''} ${t.last_name ?? ''}`.trim()).join(', ')}`,
    map.url && `[View in Atlas](${map.url})`,
  ].filter(Boolean);
  if (overview.length) headerLines.push(overview.join(' · '));

  // Units: prefer the detailed per-unit JSON; fall back to the embedded summary.
  const embedded = Array.isArray(map.units) ? map.units : [];
  const units = embedded
    .map((u) => unitDetailById.get(u.id) ?? u)
    .sort((a, b) => {
      const da = Date.parse(a.start_date ?? '') || Infinity;
      const db = Date.parse(b.start_date ?? '') || Infinity;
      return da - db;
    });

  const unitChunks = units.map((u, i) => renderUnit(u, i, ctx, sink));

  // Map-level assessments (not tied to a specific unit)
  const mapLevelLines = [];
  const mapAssessments = assessmentsByMap.get(map.id) ?? [];
  if (mapAssessments.length) {
    mapLevelLines.push('## Course-level Assessments');
    for (const a of mapAssessments) {
      mapLevelLines.push(`**${a.title ?? 'Untitled assessment'}**`);
      const content = toMd(a.content);
      if (content) mapLevelLines.push(content);
    }
  }

  // Collect links from the non-unit chunks (unit links were collected in
  // renderUnit with unit attribution).
  const headerMd = headerLines.join('\n\n');
  const mapLevelMd = mapLevelLines.join('\n\n');
  collectLinks(headerMd, ctx, sink);
  if (mapLevelMd) collectLinks(mapLevelMd, ctx, sink);

  let body = [headerMd, ...(unitChunks.length ? ['## Units', ...unitChunks] : []), mapLevelMd]
    .filter(Boolean)
    .join('\n\n');

  // Hyperlink appendix: everything linked anywhere in this document.
  const unique = [...new Map(sink.map((l) => [l.url, l])).values()];
  if (unique.length) {
    body +=
      '\n\n## Hyperlinks\n\n' +
      unique.map((l) => `- ${l.text ? `[${l.text}](${l.url})` : l.url}`).join('\n');
  }
  return body + '\n';
}

// ---------------------------------------------------------------------------
// Build the hierarchy and write files
// ---------------------------------------------------------------------------

fs.rmSync(OUT, { recursive: true, force: true });
fs.mkdirSync(OUT, { recursive: true });

// Course folder names, de-duplicated across grades (e.g. two "Mathematics"
// courses in different grades get the grade appended).
const courseFolderById = new Map();
{
  const claimed = new Map(); // sanitized name -> course id
  for (const c of courses) {
    let name = sanitizeName(c.name);
    if (claimed.has(name) && claimed.get(name) !== c.id) {
      name = sanitizeName(`${c.name} (${c.grade_name || `course ${c.id}`})`);
    }
    claimed.set(name, c.id);
    courseFolderById.set(c.id, name);
  }
}

const index = new Map(); // division -> Map(course -> [{title, relPath}])
let filesWritten = 0;

for (const map of maps) {
  const divisionRaw = map.school || 'Unsorted';
  const division = sanitizeName(DIVISION_NAMES[divisionRaw] ?? divisionRaw);

  const course = coursesById.get(map.course_id);
  const courseFolder = course
    ? courseFolderById.get(course.id)
    : sanitizeName(map.subject ? `${map.subject} (no course)` : 'Other Maps');
  const courseLabel = course?.name ?? courseFolder;

  const dir = path.join(OUT, division, courseFolder);
  fs.mkdirSync(dir, { recursive: true });

  let fileName = sanitizeName(map.title || `Map ${map.id}`);
  if (fs.existsSync(path.join(dir, `${fileName}.md`))) fileName = `${fileName} (${map.id})`;
  const filePath = path.join(dir, `${fileName}.md`);

  fs.writeFileSync(filePath, renderMap(map, courseLabel, division));
  filesWritten++;

  if (!index.has(division)) index.set(division, new Map());
  const divIndex = index.get(division);
  if (!divIndex.has(courseFolder)) divIndex.set(courseFolder, []);
  divIndex.get(courseFolder).push({
    title: map.title || `Map ${map.id}`,
    relPath: path.join(division, courseFolder, `${fileName}.md`),
  });
}

// --- INDEX.md ----------------------------------------------------------------
{
  const lines = ['# Curriculum Archive — Index', ''];
  for (const division of [...index.keys()].sort()) {
    lines.push(`## ${division}`, '');
    const divIndex = index.get(division);
    for (const courseFolder of [...divIndex.keys()].sort()) {
      const entries = divIndex.get(courseFolder);
      lines.push(`- **${courseFolder}**`);
      for (const e of entries.sort((a, b) => a.title.localeCompare(b.title))) {
        lines.push(`  - [${e.title}](<${e.relPath}>)`);
      }
    }
    lines.push('');
  }
  fs.writeFileSync(path.join(OUT, 'INDEX.md'), lines.join('\n'));
}

// --- Loose attachments (not tied to a unit) ----------------------------------
if (looseAttachments.length) {
  const lines = ['# Attachments not tied to a unit', ''];
  for (const att of looseAttachments) {
    lines.push(`- [${att.title || att.url}](${att.url}) — ItemType: ${att.ItemType ?? '?'}, ItemID: ${att.ItemID ?? '?'}`);
    if (att.url) allLinks.push({ division: '', course: '', map: '', unit: '', text: att.title ?? '', url: att.url });
  }
  fs.writeFileSync(path.join(OUT, 'UNFILED-ATTACHMENTS.md'), lines.join('\n'));
}

// --- links.csv ----------------------------------------------------------------
{
  const header = 'division,course,map,unit,text,url';
  const rows = allLinks.map((l) =>
    [l.division, l.course, l.map, l.unit ?? '', l.text, l.url].map(csvEscape).join(',')
  );
  fs.writeFileSync(path.join(ROOT, 'output', 'links.csv'), [header, ...rows].join('\n'));
}

console.log(`Rendered ${filesWritten} map document(s) across ${index.size} division(s).`);
console.log(`Hyperlinks collected: ${allLinks.length} (see output/links.csv)`);
console.log(`Browse from: ${path.relative(process.cwd(), path.join(OUT, 'INDEX.md'))}`);
