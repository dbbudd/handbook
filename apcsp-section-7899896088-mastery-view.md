# AP Computer Science Principles — Mastery View

**Section:** 7(A-B) (`7899896088`) · **Course:** AP Computer Science Principles (`488287437`)
**Grading period:** 2025–2026 (`1123041`, 08/14/2025–06/17/2026) — current & active, verified · **As of:** 2026-05-30
**Students:** 14 · **Graded records:** 324 (290 summative, 34 formative)

> Names are redacted to opaque IDs by the Schoology MCP (PII guard). Pair each ID
> with the live roster to recover the real student.

## Proficiency scale (HKIS General Academic Scale)

Confirmed from section grading scale `21337256` / `23495360`. Numeric grade → level:

| Level | Abbr. | Band (0–100) |
|---|---|---|
| Exhibiting Depth | ED | ≥ 87.5 |
| Exhibiting | EX | 62.5 – 87.4 |
| Developing | D | 37.5 – 62.4 |
| Emerging | EM | 12.5 – 37.4 |
| Insufficient Evidence | IE | < 12.5 |

## Mastery grid (by grading category)

Each cell = the student's mean score in that category, mapped to the scale above.
`(n)` = number of graded items.

| Student | Formative | Summative | Final |
|---|---|---|---|
| Student_7f0e0c01 | ED (4) | ED (19) | **ED** · 97.9 |
| Student_ad29eb45 | ED (5) | ED (19) | **ED** · 96.1 |
| Student_5395ab6d | ED (2) | ED (19) | **ED** · 93.2 |
| Student_bd41343d | EX (5) | ED (20) | **ED** · 90.4 |
| Student_a5bae26c | ED (1) | ED (21) | **ED** · 90.3 |
| Student_2ee70066 | IE (1)* | ED (20) | **ED** · 90.0 |
| Student_3b4157cc | IE (1)* | ED (19) | **ED** · 89.8 |
| Student_02d98260 | ED (1) | ED (19) | **ED** · 89.6 |
| Student_eef0ff8c | ED (3) | EX (20) | **EX** · 86.3 |
| Student_149f1e4d | ED (3) | EX (22) | **EX** · 85.4 |
| Student_d8aa70d7 | D (3) | EX (21) | **EX** · 83.6 |
| Student_1247d64c | IE (1)* | EX (22) | **EX** · 74.2 |
| Student_cecf1a53 | ED (3) | EX (23) | **EX** · 73.3 |
| Student_a3530e31 | ED (1) | EX (22) | **EX** · 72.4 |

## Distribution (by Final)

| Level | Students |
|---|---|
| Exhibiting Depth (ED) | 8 |
| Exhibiting (EX) | 6 |
| Developing (D) | 0 |
| Emerging (EM) | 0 |
| Insufficient Evidence (IE) | 0 |

## What this is — and what it can't be (important)

**This grid is built entirely from the Schoology MCP (REST), not scraped.** But REST
imposes a real limit on the columns:

- **Columns are grading categories (Formative / Summative), not measurement topics
  or reporting categories.** Schoology's REST API exposes only the two grading
  categories on this section (`category_id` 86984135 / 90141393). It does **not**
  expose measurement topics or reporting categories — those live solely in
  Schoology's **Mastery Gradebook**, which the grading-scale labels literally point
  to ("See Mastery Gradebook"). No MCP tool can reach that data; every analytics
  tool (`gradebook_mastery_overview`, `gradebook_summary`, `predict_proficiency_dates`,
  `get_student_grades`) composes the same REST grades fetch and inherits the same
  blind spot. `get_course_rubrics` for this course is empty (0 rubrics).
- **The Final column is authoritative** — Schoology itself rolls the final onto this
  exact scale. The **Formative / Summative** cells are this view's mapping of each
  category mean through the same scale; the official Mastery Gradebook may compute
  category/topic proficiency with different rules (e.g. most-recent or decaying
  average rather than a simple mean).
- **`*` formative cells** (`IE (1)`) are a single zero-scored formative item — almost
  certainly missing/incomplete work, not a mastery signal — on students whose
  summative mastery is strong. Treat as a data artifact.

**To get true per-measurement-topic / per-reporting-category proficiency**, the data
must come from the Mastery Gradebook UI (a browser read), which is outside the MCP.
This is the same gap captured in `schoology-mcp-section-teacher-gap.md` territory —
the MCP's own `gradebook_mastery_overview` documents it as requiring a future
"Phase 3 Playwright scrape."
