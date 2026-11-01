# Measurements behind every number in the report

Taken 2026-11-01 from the project's canonical (L1) and model-shaped (L2) parquet stores.
Cohort A = state-level source (5,114 participants). Cohort B = region-level source (47,297).
Pooled = 52,244. Paths are relative to the repository root.

---

## §2.1 — the observation grid

```python
d = pd.read_parquet('data/L1_canonical/cohort_a/L1_canonical.parquet')
d.groupby(['visit_idx','hm_year']).agg(
    rows=('hm_id','size'),
    diet=('HM_AHEI2010_TOTAL_SCORE', lambda s: s.notna().sum()),
    clinical=('hm_bmi', lambda s: s.notna().sum()),
    score=('hm_CES_D60', lambda s: s.notna().sum()),
    count=('hm_cvd_repeated_count', lambda s: s.notna().sum()),
    diet_mask=('mask_HM_AHEI2010_TOTAL_SCORE','mean'),
    clinical_mask=('mask_hm_bmi','mean'))
```

Result (cohort A, 19 steps, 5,114 participants each):

| step | year | diet | clinical | score | count | diet mask | clinical mask |
|---|---|---|---|---|---|---|---|
| 0 | 1985 | 5110 | 5097 | 0 | 5114 | 0.999 | 0.997 |
| 1 | 1986 | 0 | 0 | 0 | 5114 | 0.000 | 0.000 |
| 2 | 1987 | 0 | 0 | 0 | 5114 | 0.000 | 0.000 |
| 3 | 1990 | 0 | 0 | 4290 | 5114 | 0.000 | 0.000 |
| 4 | 1993 | 3942 | 0 | 0 | 5114 | 0.771 | 0.000 |
| 5 | 1994 | 0 | 0 | 0 | 5114 | 0.000 | 0.000 |
| 6 | 1995 | 0 | 0 | 3875 | 5114 | 0.000 | 0.000 |
| 7–10 | 1996–1999 | 0 | 0 | 0 | 5114 | 0.000 | 0.000 |
| 11 | 2000 | 0 | 0 | 3622 | 5114 | 0.000 | 0.000 |
| 12–15 | 2001–2004 | 0 | 0 | 0 | 5114 | 0.000 | 0.000 |
| 16 | 2006 | 3141 | 0 | 3448 | 5114 | 0.614 | 0.000 |
| 17 | 2011 | 0 | 0 | 3454 | 5114 | 0.000 | 0.000 |
| 18 | 2016 | 0 | 0 | 3268 | 5114 | 0.000 | 0.000 |

**11 of 19 steps carry no measurement of any kind.** The empty years (1986, 1994,
1996–1999, 2001–2004) are follow-up years of the other source cohort; 1987 is a cohort A
examination year whose clinical content was not carried through harmonisation.

The decay-variant result quoted in §2.1 (0.7552 against 0.7909 for a plain recurrent encoder)
is from the project's existing leaderboard, not measured here.

---

## §2.2 — sequence length against realised observations

```python
for coh in ('cohort_a','cohort_b'):
    d = pd.read_parquet(f'data/L1_canonical/{coh}/L1_canonical.parquet')
    g = d.groupby('hm_id')[col].apply(lambda s: s.notna().sum())
    g[g>0].mean(), g.max(), (g>0).sum()
```

| block | steps | obs/participant | max | participants |
|---|---|---|---|---|
| clinical panel, cohort A | 19 | 1.00 | 1 | 3,918–5,112 (76.6–99.96%) |
| diet, cohort A (24 cols) | 19 | 2.38 | 3 | 5,114 (100%) |
| diet, cohort B (12 cols) | 15 | 2.35 | 6 | 47,297 (100%) |
| symptom score, cohort B | 15 | 2.83 | 8 | 47,122 (99.6%) |
| symptom score, cohort A | 19 | 4.63 | 6 | 4,746 |
| event count, cohort B | 15 | 15.00 | 15 | 47,297 (100%) |
| event count, cohort A | 19 | 19.00 | 19 | 5,114 (100%) |

Diet calendar years — cohort A: 1985 (5,110), 1993 (3,942), 2006 (3,143). Cohort B spans 11
calendar years, but that is staggered enrolment: the two modes are 1993 (47,219) and 1996
(26,362), i.e. baseline and year 3.

Columns present in L1 and absent from L2 (40 in total, removed as outcome-source):
`hm_CES_D18`, `hm_CES_D60`, `hm_cvd_repeated_count`, `hm_cvd_survival`,
`hm_diab_repeated_binary`, `hm_dep_repeated_binary`, 13 cardiovascular subtype indicators and
their day columns, and the event-day columns.

---

## §2.3 — two-phase covariate structure

```python
d = pd.read_parquet('data/L2_model_shape/cohort_b/L2_sequence_wide.parquet')
g = d.groupby('hm_id')[col].apply(lambda s: s.notna().any())
```

Cohort B, participants with at least one non-null value:

| column | participants | share |
|---|---|---|
| age, sex, race, systolic | 1,670 | 3.53% |
| glucose | 1,668 | 3.53% |
| BMI | 1,660 | 3.51% |
| education | 1,659 | 3.51% |
| marital status | 1,658 | 3.51% |
| diet score | 47,297 | 100% |
| region | 47,297 | 100% |

Set relations: BMI ⊂ age (participants with BMI but no age = 0); Jaccard(age, BMI) = 0.9940.

Selection by enrolment arm:

| arm | participants | with panel | rate |
|---|---|---|---|
| observational | 27,361 | 308 | 1.13% |
| trial arm 1 | 7,002 | 403 | 5.76% |
| trial arm 2 | 2,331 | 225 | 9.65% |
| trial arm 3 | 10,603 | 734 | 6.92% |

Case enrichment of the measured subsample (from the project's exploratory analysis of the
source file): +4.1, +5.9 and +9.4 percentage points for the cardiovascular, metabolic and
psychiatric outcomes, p = 7.1e-44 for the psychiatric comparison.

Cohort A arithmetic inconsistency: a derived insulin-resistance index is present at three
calendar years (1985, 1993, 2006; 11,046 rows, 2.55 per participant) while glucose (5,033
rows) and insulin (5,044 rows) are each present at one. The index is a function of the two,
so the repeated inputs exist upstream.

---

## §4.3 — evaluation reference points

From the project's existing result stores, not measured here:

- visit-index-only predictor, pooled psychiatric outcome: 0.6918 ± 0.0027; strongest sequence
  model 0.7379; prevalence floor 0.2899.
- paired minimum detectable effect across the nine cohort-by-outcome cells: 0.0011 to 0.0107.
- cross-device drift on identical code: up to 0.0056 in the five-fold mean, 0.0244 in a single
  fold.
- gradient-boosted trees match or exceed every sequence model in all nine cells; 55–91% of
  their gain is attributed to visit-schedule features.
- feature-axis attention: 0.7815 against 0.7934 for trees, paired difference −0.0120, p = 0.038.
- per-group model selection: positive on validation folds in 18 of 18 cells, negative
  out-of-fold in 16 of 18.
