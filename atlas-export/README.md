# Atlas Curriculum Export

Pulls all of HKIS's curriculum data out of Rubicon Atlas (now OnAtlas) through
the official read-only API, archives it as raw JSON, and renders it into a
browsable Markdown hierarchy organised by **division → course → curriculum map**,
with every hyperlink inventoried.

API docs: <https://api.onatlas.com/docs/> (spec: `/docs/ra_api.yaml`)

## One-time setup

```bash
cd atlas-export
nvm use 22          # any Node ≥ 20 works
npm install
cp .env.example .env
# …then paste your Atlas API access key into .env
```

The access key comes from your Atlas site admin account or Faria/Atlas support
([help.onatlas.com](https://help.onatlas.com/)).

## Run order

| Step | Command | What it does |
|---|---|---|
| 1 | `npm run probe` | Verifies the token, then answers the make-or-break question: does `/unit/{id}` return real planner content or only metadata? Saves raw samples to `output/probe/`. |
| 2 | `npm run extract` | Full pull → `output/raw/` (JSON). Resumable — re-running skips files already downloaded. `--limit 10` for a test run, `--force` to refetch. |
| 3 | `npm run render` | `output/raw/` → `output/markdown/` hierarchy + `INDEX.md` + `output/links.csv`. Never touches the API; re-run freely. |

**Stop after the probe and review `output/probe/unit-*.json`.** If it shows only
the documented metadata fields (title, timeframe, dates), the API does not expose
the unit planner prose and we need a fallback (bulk PDF export / Faria dump /
browser scrape) before the full extract is worth running.

## Output layout

```
output/
├── raw/                      ← canonical lossless archive — KEEP THIS
│   ├── schools.json  subjects.json  grades.json  teachers.json
│   ├── courses.json  maps.json  units.json
│   ├── units/{id}.json       ← per-unit detail (planner content, if exposed)
│   ├── assessments.json  attachments.json
├── markdown/
│   ├── INDEX.md              ← linked table of contents
│   └── <Division>/<Course>/<Map>.md
└── links.csv                 ← every hyperlink: division, course, map, unit, text, url
```

## The Markdown template

The per-map document layout lives in `renderMap()` / `renderUnit()` in
[render.mjs](render.mjs), inside the block marked **`███ MARKDOWN TEMPLATE ███`**.
The current layout is a placeholder default: front-matter, overview line, one
section per unit (any planner fields the API returns are rendered automatically),
assessments, attachments, and a hyperlink appendix.

**To adopt the agreed format:** replace the body of `renderMap()`/`renderUnit()`
once the example document is available. Nothing else needs to change — the
hierarchy, link inventory, and INDEX generation are independent of the template.

Division folder names default to the Atlas "school" names; rename them via the
`DIVISION_NAMES` map at the top of `render.mjs`.

## Notes

- The Atlas API is **read-only** (GET only) — these scripts cannot modify
  anything in Atlas.
- HTML in content fields is converted to Markdown (via turndown); links survive
  the conversion and are also collected into `links.csv`.
- Attachments are **linked, not downloaded**. If the probe shows attachment URLs
  are fetchable with the API token, a download step can be added to `extract.mjs`.
- `output/` and `.env` are git-ignored.
