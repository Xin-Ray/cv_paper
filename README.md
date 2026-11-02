# COM6005 Midterm Project — Code

Xin Xiang · Katz School of Science and Health, Yeshiva University

This snapshot reproduces every figure, table and number in the midterm progress report.
Nothing here is illustrative: each quantity was measured on 2026-11-01 from the project's
canonical (L1) and model-shaped (L2) parquet stores.

## Contents

| File | What it does |
|---|---|
| `make_figures.py` | Builds the three data-audit figures, each with a same-named `.csv` of the plotted values |
| `make_pipeline_figure.py` | Builds the DCPTL pipeline figure used in Section 3.4 |
| `build_report.py` | Builds `MIDTERM_REPORT.docx` (8 full pages) inside the course template `_template.docx`, using the template's own styles, page setup and fonts |
| `build_slides.py` | Builds `midterm_slides.pptx` (10 slides) with python-pptx |
| `_template.docx` | The course report template, converted once from the supplied `.doc`; the report is written into it rather than re-creating its layout |
| `MEASUREMENTS.md` | The exact queries behind every number, so each claim can be re-derived |
| `figures/` | Generated figures and their data |

## Run

```bash
pip install python-docx python-pptx matplotlib pandas numpy pyarrow
python make_figures.py
python make_pipeline_figure.py
python build_report.py
python build_slides.py
```

## Note on data

The underlying cohort is the iPAT harmonized national database, which is governed by a data
use agreement and is not redistributable. `MEASUREMENTS.md` gives the query for every number
so the results can be checked by anyone with authorised access. The figure `.csv` files carry
the aggregate values actually plotted, which are not disclosive.
