# COM6005 Midterm Project — Code

Xin Xiang · Katz School of Science and Health, Yeshiva University

This snapshot reproduces every figure, table and number in the midterm progress report.
Nothing here is illustrative: each quantity was measured on 2026-11-01 from the project's
canonical (L1) and model-shaped (L2) parquet stores.

## Contents

| File | What it does |
|---|---|
| `make_figures.py` | Builds the three report figures, each with a same-named `.csv` of the plotted values |
| `build_report.py` | Builds `MIDTERM_REPORT.docx` (9 pages) with python-docx |
| `build_slides.py` | Builds `midterm_slides.pptx` (7 slides) with python-pptx |
| `MEASUREMENTS.md` | The exact queries behind every number, so each claim can be re-derived |
| `figures/` | Generated figures and their data |

## Run

```bash
pip install python-docx python-pptx matplotlib pandas numpy pyarrow
python make_figures.py
python build_report.py
python build_slides.py
```

## Note on data

The underlying cohort is the iPAT harmonized national database, which is governed by a data
use agreement and is not redistributable. `MEASUREMENTS.md` gives the query for every number
so the results can be checked by anyone with authorised access. The figure `.csv` files carry
the aggregate values actually plotted, which are not disclosive.
