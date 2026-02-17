# Deterministic Data Workflow

This repository contains a deterministic workflow that ranks MAcc courses from the anonymized 2024 exit survey workbook.

## Question answered

**Rank order the programs or courses based on student ratings or preferences in this dataset.**

The workflow produces two rankings:
- **Core course preference ranking** using the drag-and-drop rank question (lower average rank is better).
- **Elective course rating ranking** using the 1-5 rating question (higher average rating is better).

## Deterministic approach

`scripts/rank_courses.py` parses the XLSX directly using only Python standard library modules:
- Reads `xl/sharedStrings.xml` and `xl/worksheets/sheet1.xml`.
- Uses Qualtrics row 2 as human-readable headers.
- Uses rows 4+ as responses (skipping metadata rows).
- Aggregates averages by course.
- Uses fixed tie-breaking by course name for deterministic ordering.

## Run locally

```bash
python scripts/rank_courses.py
```

Outputs are written to `output/`:
- `course_ranking_report.md`
- `core_course_ranking.csv`
- `elective_course_ranking.csv`
- `course_ranking_summary.json`

## GitHub Actions

Workflow: `.github/workflows/deterministic-course-ranking.yml`

It runs on push / pull request / manual dispatch and uploads the generated ranking files as build artifacts.

## Current ranking (from included 2024 dataset)

See `output/course_ranking_report.md` for the full ranked tables.
