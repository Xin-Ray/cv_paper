"""Pipeline figure for the Proposed Method section of the midterm report.

The course template asks for a pipeline figure in "Our proposed method". This draws the four
stages of DCPTL as a left-to-right pipeline, with the inputs on the left and the three
outcome-matched heads on the right.

Run: python cv/make_pipeline_figure.py
"""
from __future__ import annotations

from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch

OUT = Path(__file__).resolve().parent / "figures" / "fig4_pipeline.png"

INK = "#0c182b"
TEAL = "#006e66"
CRIMSON = "#9e1c44"
GREY = "#4e5b6b"
FILL = "#ffffff"
EDGE = "#c8d2dd"


def box(ax, x, y, w, h, title, lines, accent=TEAL, fill=FILL, tsize=8.4, bsize=7.0):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.012,rounding_size=0.02",
                                linewidth=0.9, edgecolor=EDGE, facecolor=fill, zorder=2))
    nl = chr(10)
    ax.text(x + w / 2, y + h - 0.045, title, ha="center", va="top", fontsize=tsize,
            color=accent, weight="bold", zorder=3, linespacing=1.35)
    drop = 0.055 + 0.058 * (title.count(nl) + 1)
    ax.text(x + w / 2, y + h - drop, nl.join(lines), ha="center", va="top", fontsize=bsize,
            color=INK, zorder=3, linespacing=1.45)


def arrow(ax, x0, x1, y, label=None):
    ax.add_patch(FancyArrowPatch((x0, y), (x1, y), arrowstyle="-|>", mutation_scale=9,
                                 linewidth=1.0, color=GREY, zorder=3))
    if label:
        ax.text((x0 + x1) / 2, y + 0.018, label, ha="center", va="bottom", fontsize=6.2,
                color=GREY, style="italic")


def main() -> None:
    fig, ax = plt.subplots(figsize=(6.6, 2.55), dpi=300)
    ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")

    # ---------------------------------------------------------------- input
    box(ax, 0.005, 0.30, 0.135, 0.46, "Patient record",
        ["visits at", "irregular years", "", "most values", "missing", "", "3-state", "mask"],
        accent=GREY, fill="#f3f6f9", tsize=7.8, bsize=6.4)
    arrow(ax, 0.145, 0.185, 0.53)

    # --------------------------------------------------------------- stages
    box(ax, 0.19, 0.28, 0.175, 0.50, "Stage 1\nShared encoder",
        ["time-aware attention", "keyed on elapsed time", "(or a neural ODE)", "",
         "-> latent health", "state z_i shared by", "all three diseases"])
    arrow(ax, 0.370, 0.410, 0.53, "z_i")

    box(ax, 0.415, 0.28, 0.175, 0.50, "Stage 2\nDisease-conditioned\ndynamics",
        ["one learnable f_d(z_i,t)", "per disease, chosen by", "a disease code d", "",
         "-> three evolutions", "of the same state"], accent=CRIMSON)
    arrow(ax, 0.595, 0.635, 0.53)

    box(ax, 0.640, 0.28, 0.175, 0.50, "Stage 3\nHeterogeneity",
        ["mixture of experts /", "trajectory prototypes", "", "-> several modes",
         "per disease, with", "per-patient weights"])
    arrow(ax, 0.820, 0.840, 0.53)

    # ---------------------------------------------------------------- heads
    box(ax, 0.843, 0.620, 0.150, 0.160, "Cardiovascular",
        ["survival + count"], accent=CRIMSON, tsize=7.4, bsize=6.4)
    box(ax, 0.843, 0.435, 0.150, 0.160, "Diabetes",
        ["longitudinal logistic"], accent=TEAL, tsize=7.4, bsize=6.4)
    box(ax, 0.843, 0.250, 0.150, 0.160, "Depression",
        ["two-part continuous"], accent=TEAL, tsize=7.4, bsize=6.4)
    ax.text(0.918, 0.800, "Stage 4  matched heads", ha="center", va="bottom", fontsize=7.4,
            color=GREY, weight="bold")

    # ------------------------------------------------------- audit callouts
    for x, t in ((0.2775, "Finding 1\n3-state mask"),
                 (0.5025, "Finding 2\nper-head deletion"),
                 (0.9180, "Finding 3\nsampling weights")):
        ax.text(x, 0.195, t, ha="center", va="top", fontsize=6.2, color=CRIMSON,
                style="italic", linespacing=1.3)

    ax.text(0.5, 0.020, "Italic labels mark where each finding of Section 4.4 enters the pipeline.", ha="center", va="bottom", fontsize=6.0, color=GREY,
            style="italic")

    fig.subplots_adjust(left=0, right=1, top=1, bottom=0)
    OUT.parent.mkdir(exist_ok=True)
    fig.savefig(OUT, bbox_inches="tight", facecolor="white")
    print(f"[cv] wrote {OUT}")


if __name__ == "__main__":
    main()
