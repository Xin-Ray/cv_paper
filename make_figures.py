"""Figures for the COM6005 midterm progress report.

Every number drawn here was measured on 2026-11-01 from the project's L1/L2 parquet files
(see cv/MEASUREMENTS.md for the exact queries). Nothing is illustrative; nothing is invented.

Out: cv/figures/*.png  +  a same-named .csv for each figure (project rule).
Run: python cv/make_figures.py
"""
from __future__ import annotations

from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np               # noqa: E402
import pandas as pd              # noqa: E402

HERE = Path(__file__).resolve().parent
FIG = HERE / "figures"
FIG.mkdir(parents=True, exist_ok=True)

INK, MUTED, RULE = "#0C182B", "#566577", "#E1E7EE"
TEAL, AMBER, CRIMSON, BLUE = "#007D74", "#DC7700", "#A61F49", "#21507B"


def _style():
    plt.rcParams.update({
        "figure.facecolor": "white", "axes.facecolor": "white",
        "font.family": "DejaVu Sans", "axes.edgecolor": RULE,
        "axes.labelcolor": INK, "text.color": INK,
        "xtick.color": MUTED, "ytick.color": MUTED, "axes.titlecolor": INK,
    })


# ----------------------------------------------------------------- Figure 1
# Cohort A's 19-step national grid: what is actually recorded at each step.
GRID = pd.DataFrame({
    "step": list(range(19)),
    "year": [1985, 1986, 1987, 1990, 1993, 1994, 1995, 1996, 1997, 1998,
             1999, 2000, 2001, 2002, 2003, 2004, 2006, 2011, 2016],
    "diet":       [5110, 0, 0, 0, 3942, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 3141, 0, 0],
    "clinical":   [5097, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    "cesd":       [0, 0, 0, 4290, 0, 0, 3875, 0, 0, 0, 0, 3622, 0, 0, 0, 0, 3448, 3454, 3268],
    "eventcount": [5114] * 19,
})
N_A = 5114


def figure1():
    _style()
    rows = ["eventcount", "cesd", "diet", "clinical"]
    label = {"eventcount": "Event count\n(outcome-source)", "cesd": "Symptom score\n(continuous)",
             "diet": "Diet quality\n(24 columns)", "clinical": "Clinical panel\n(13 columns)"}
    M = np.array([GRID[r].values / N_A for r in rows])

    fig, ax = plt.subplots(figsize=(13.2, 3.9))
    ax.imshow(M, aspect="auto", cmap="Blues", vmin=0, vmax=1)
    for i in range(len(rows)):
        for j in range(19):
            v = M[i, j]
            if v > 0:
                ax.text(j, i, f"{v*100:.0f}", ha="center", va="center", fontsize=8.0,
                        color="white" if v > 0.55 else INK)
    # mark the structurally empty steps
    empty = [j for j in range(19) if GRID.loc[j, ["diet", "clinical", "cesd"]].sum() == 0]
    for j in empty:
        ax.add_patch(plt.Rectangle((j - .5, -.5), 1, len(rows), fill=False,
                                   edgecolor=CRIMSON, lw=2.0, zorder=5))
    ax.set_xticks(range(19))
    ax.set_xticklabels([f"{y}" for y in GRID.year], rotation=60, fontsize=8.5)
    ax.set_yticks(range(len(rows)))
    ax.set_yticklabels([label[r] for r in rows], fontsize=9.5)
    ax.set_xlabel("national harmonisation time step  (cohort A, 19 steps)", fontsize=10.5)
    ax.set_title("Percent of participants with a recorded value at each modelled time step",
                 fontsize=12.5, pad=10)
    fig.text(0.5, -0.17,
             f"Outlined columns are the {len(empty)} of 19 steps that carry no measurement for this "
             "cohort: they are calendar positions contributed by the other cohort. Separating them "
             "from an attended-but-unmeasured visit, and from a scheduled visit the participant "
             "missed, is what the three-state mask of Stage 1 encodes.",
             ha="center", va="top", fontsize=9.3, color=MUTED, wrap=True)
    fig.tight_layout()
    fig.savefig(FIG / "fig1_grid_audit.png", dpi=200, bbox_inches="tight")
    plt.close(fig)
    GRID.to_csv(FIG / "fig1_grid_audit.csv", index=False)
    print(f"  [fig1] {len(empty)} structurally empty steps of 19")


# ----------------------------------------------------------------- Figure 2
# Modelled sequence length versus the number of real observations it carries.
SHAPE = pd.DataFrame({
    "block": ["Clinical panel\n(cohort A)", "Diet quality\n(cohort A)", "Diet quality\n(cohort B)",
              "Symptom score\n(cohort B)", "Symptom score\n(cohort A)",
              "Event count\n(cohort B)", "Event count\n(cohort A)"],
    "steps":    [19, 19, 15, 15, 19, 15, 19],
    "observed": [1.00, 2.38, 2.35, 2.83, 4.63, 15.00, 19.00],
})


def figure2():
    _style()
    fig, ax = plt.subplots(figsize=(11.0, 4.6))
    y = np.arange(len(SHAPE))
    ax.barh(y, SHAPE.steps, height=.62, color=RULE, label="time steps the encoder unrolls")
    ax.barh(y, SHAPE.observed, height=.34, color=TEAL, label="real observations per participant (mean)")
    for i, r in SHAPE.iterrows():
        ax.text(r.steps + .25, i, f"{r.steps}", va="center", fontsize=9.5, color=MUTED)
        ax.text(r.observed + .25, i - .02, f"{r.observed:.2f}", va="center", fontsize=9.5,
                color=TEAL, fontweight="bold")
    ax.set_yticks(y)
    ax.set_yticklabels(SHAPE.block, fontsize=10)
    ax.invert_yaxis()
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.set_xlabel("count", fontsize=10.5)
    ax.set_xlim(0, 21.5)
    ax.set_title("Sequence length is set by the grid, not by the data", fontsize=13, pad=10)
    ax.legend(frameon=False, fontsize=10, loc="lower right")
    fig.text(0.5, -0.07,
             "The two blocks that fill every step are outcome-source columns, which the current "
             "feature build removes in their entirety. The blocks the model is actually given "
             "carry between 1.00 and 2.38 observations across 19 unrolled steps.",
             ha="center", va="top", fontsize=9.5, color=MUTED, wrap=True)
    fig.tight_layout()
    fig.savefig(FIG / "fig2_shape_mismatch.png", dpi=200, bbox_inches="tight")
    plt.close(fig)
    SHAPE.to_csv(FIG / "fig2_shape_mismatch.csv", index=False)
    print("  [fig2] shape mismatch")


# ----------------------------------------------------------------- Figure 3
# Cohort B: who carries the clinical panel, and how that subsample was selected.
PHASE = pd.DataFrame({
    "arm": ["Observational", "Trial arm 1", "Trial arm 2", "Trial arm 3"],
    "n":        [27361, 7002, 2331, 10603],
    "measured": [308, 403, 225, 734],
})
PHASE["rate"] = 100 * PHASE.measured / PHASE.n
ENRICH = pd.DataFrame({
    "outcome": ["Cardiovascular", "Diabetes", "Depression"],
    "pp": [4.1, 5.9, 9.4],
})


def figure3():
    _style()
    fig, axes = plt.subplots(1, 2, figsize=(12.6, 4.2),
                             gridspec_kw={"width_ratios": [1.5, 1.0]})
    ax = axes[0]
    x = np.arange(len(PHASE))
    ax.bar(x, PHASE.rate, color=BLUE, width=.58)
    for i, r in PHASE.iterrows():
        ax.text(i, r.rate + .25, f"{r.rate:.2f}%\n{r.measured:,} / {r.n:,}", ha="center",
                fontsize=9.3, color=INK)
    ax.set_xticks(x)
    ax.set_xticklabels(PHASE.arm, fontsize=10)
    ax.set_ylabel("share with a clinical panel", fontsize=10.5)
    ax.set_ylim(0, 12.5)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.set_title("Selection into the measured subsample is not uniform", fontsize=12, pad=8)

    ax = axes[1]
    y = np.arange(len(ENRICH))
    ax.barh(y, ENRICH.pp, color=CRIMSON, height=.5)
    for i, r in ENRICH.iterrows():
        ax.text(r.pp + .15, i, f"+{r.pp:.1f} pp", va="center", fontsize=10, color=CRIMSON,
                fontweight="bold")
    ax.set_yticks(y)
    ax.set_yticklabels(ENRICH.outcome, fontsize=10)
    ax.invert_yaxis()
    ax.set_xlim(0, 12)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.set_xlabel("excess event rate in the measured subsample", fontsize=10.5)
    ax.set_title("and it over-samples cases  (p = 7.1e-44)", fontsize=12, pad=8)

    fig.text(0.5, -0.06,
             "Only 1,670 of 47,297 cohort B participants (3.53%) carry any clinical covariate, and "
             "the same 1,670 carry all of them: age, sex, race, education, blood pressure and "
             "glucose coincide to a Jaccard index of 0.9940. The design is therefore a two-phase "
             "sample with a known selection rate, which Stage 4 uses as an inverse-probability "
             "weight over the full cohort.",
             ha="center", va="top", fontsize=9.5, color=MUTED, wrap=True)
    fig.tight_layout()
    fig.savefig(FIG / "fig3_two_phase.png", dpi=200, bbox_inches="tight")
    plt.close(fig)
    PHASE.to_csv(FIG / "fig3_two_phase.csv", index=False)
    print("  [fig3] two-phase sampling")


def main():
    print("[cv] building midterm figures")
    figure1()
    figure2()
    figure3()
    print(f"[cv] -> {FIG}")


if __name__ == "__main__":
    main()
