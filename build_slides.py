"""Build the COM6005 midterm presentation (5 minutes, 7 slides).

The course notice asks the midterm talk to focus on progress/update and the updated timeline,
and explicitly not to revisit the problem and significance description. The deck is built to
that instruction: one slide of context, three of findings, one of the revised method, one of
the timeline.

Cohort names never appear; the database is the iPAT harmonized national database with sources
cohort A and cohort B.

Run: python cv/build_slides.py
"""
from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Inches, Pt

HERE = Path(__file__).resolve().parent
FIG = HERE / "figures"
OUT = HERE / "midterm_slides.pptx"

BG = RGBColor(0xFB, 0xFA, 0xF7)
INK = RGBColor(0x0C, 0x18, 0x2B)
MUTED = RGBColor(0x56, 0x65, 0x77)
TEAL = RGBColor(0x00, 0x7D, 0x74)
CRIMSON = RGBColor(0xA6, 0x1F, 0x49)
RULE = RGBColor(0xE1, 0xE7, 0xEE)

W, H = Inches(13.333), Inches(7.5)
M = Inches(0.62)
CW = W - 2 * M

_page = {"n": 0}


def slide(prs, eyebrow, title):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = BG
    _page["n"] += 1

    tb = s.shapes.add_textbox(M, Inches(0.42), CW, Inches(0.3)).text_frame
    tb.word_wrap = True
    p = tb.paragraphs[0]
    r = p.add_run(); r.text = eyebrow.upper()
    r.font.size, r.font.bold, r.font.color.rgb, r.font.name = Pt(11), True, TEAL, "Aptos"

    tb = s.shapes.add_textbox(M, Inches(0.78), CW, Inches(0.75)).text_frame
    tb.word_wrap = True
    p = tb.paragraphs[0]
    r = p.add_run(); r.text = title
    r.font.size, r.font.bold, r.font.color.rgb, r.font.name = Pt(29), True, INK, "Aptos Display"

    ln = s.shapes.add_shape(1, M, Inches(1.62), CW, Emu(9525))
    ln.fill.solid(); ln.fill.fore_color.rgb = RULE; ln.line.fill.background()

    ft = s.shapes.add_textbox(M, H - Inches(0.52), CW, Inches(0.3)).text_frame
    p = ft.paragraphs[0]; p.alignment = PP_ALIGN.RIGHT
    r = p.add_run(); r.text = f"{_page['n']} / 7"
    r.font.size, r.font.color.rgb, r.font.name = Pt(10), MUTED, "Aptos"
    return s


def text(s, x, y, w, h, lines, size=15, color=INK, bold_first=False):
    tf = s.shapes.add_textbox(x, y, w, h).text_frame
    tf.word_wrap = True
    for i, t in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(9)
        r = p.add_run(); r.text = t
        r.font.size = Pt(size)
        r.font.color.rgb = color
        r.font.name = "Aptos"
        r.font.bold = bold_first and i == 0
    return tf


def stat(s, x, y, value, label, color=TEAL, vsize=36):
    tf = s.shapes.add_textbox(x, y, Inches(3.0), Inches(1.2)).text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = value
    r.font.size, r.font.bold, r.font.color.rgb, r.font.name = Pt(vsize), True, color, "Aptos Display"
    p2 = tf.add_paragraph()
    r = p2.add_run(); r.text = label
    r.font.size, r.font.color.rgb, r.font.name = Pt(12), MUTED, "Aptos"


def table(s, x, y, w, header, rows, colw, size=12):
    n = len(rows) + 1
    shape = s.shapes.add_table(n, len(header), x, y, w, Inches(0.36 * n))
    t = shape.table
    for i, cw in enumerate(colw):
        t.columns[i].width = Inches(cw)
    for j, htxt in enumerate(header):
        c = t.cell(0, j); c.text = ""
        p = c.text_frame.paragraphs[0]
        r = p.add_run(); r.text = htxt
        r.font.size, r.font.bold, r.font.color.rgb, r.font.name = Pt(size), True, INK, "Aptos"
    for i, row in enumerate(rows, 1):
        for j, v in enumerate(row):
            c = t.cell(i, j); c.text = ""
            p = c.text_frame.paragraphs[0]
            r = p.add_run(); r.text = str(v)
            r.font.size, r.font.color.rgb, r.font.name = Pt(size), INK, "Aptos"
    return t


def build():
    prs = Presentation()
    prs.slide_width, prs.slide_height = W, H

    # 1 ------------------------------------------------------------- title
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.background.fill.solid(); s.background.fill.fore_color.rgb = BG
    _page["n"] += 1
    tf = s.shapes.add_textbox(M, Inches(2.1), CW, Inches(1.8)).text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "Learning Disease-Conditioned Dynamics from Sparse,\nIrregularly Observed Longitudinal Records"
    r.font.size, r.font.bold, r.font.color.rgb, r.font.name = Pt(34), True, INK, "Aptos Display"
    text(s, M, Inches(4.0), CW, Inches(1.6),
         ["COM6005 Computer Vision  ·  Midterm Project Progress",
          "Xin Xiang  ·  Katz School of Science and Health, Yeshiva University",
          "November 2, 2026"], size=15, color=MUTED)
    text(s, M, Inches(6.5), CW, Inches(0.5),
         ["Data access and domain supervision: H. Fang, NIDDK R01DK129432"],
         size=11, color=MUTED)

    # 2 ------------------------------------------------------- where we were
    s = slide(prs, "where the proposal left off", "The plan was an architecture comparison")
    text(s, M, Inches(2.0), Inches(7.4), Inches(4.2),
         ["Proposed: encode one axis at a time, with an operator matched to each axis, and "
          "identify every coordinate by what it is rather than where it sits.",
          "Planned first experiment: factorised encoder vs joint attention vs single-axis "
          "variants, capacity matched.",
          "Six weeks of data work did not overturn that argument — but it changed which "
          "experiment should come first."],
         size=16)
    stat(s, Inches(8.4), Inches(2.1), "3", "findings that forced a revision")
    stat(s, Inches(8.4), Inches(3.5), "0", "of them required new data")
    stat(s, Inches(8.4), Inches(4.9), "1 week", "added to the timeline")

    # 3 --------------------------------------------------------- finding 1
    s = slide(prs, "finding 1", "The sequence length is set by the calendar, not the data")
    s.shapes.add_picture(str(FIG / "fig1_grid_audit.png"), M, Inches(1.95), width=Inches(9.3))
    text(s, Inches(10.2), Inches(2.1), Inches(2.6), Inches(4.4),
         ["11 of 19 time steps carry no measurement of any kind.",
          "They exist because the grid is the union of two source calendars.",
          "The mask cannot tell them apart from “attended but not measured” — or from "
          "“due and did not attend.”",
          "That third case is the signal we now want to model."],
         size=13, color=MUTED)

    # 4 --------------------------------------------------------- finding 2
    s = slide(prs, "finding 2", "The two densest signals are the two we delete")
    table(s, M, Inches(2.05), Inches(8.6),
          ["Variable block", "Steps", "Observations", "Status"],
          [["Clinical panel (13 cols)", "19", "1.00", "used"],
           ["Diet quality (24 cols)", "19", "2.38", "used"],
           ["Continuous symptom score", "19", "4.63", "removed"],
           ["Cumulative event count", "19", "19.00", "removed"]],
          [3.9, 1.0, 1.9, 1.8], size=14)
    text(s, M, Inches(4.5), Inches(8.6), Inches(2.4),
         ["Both removed blocks are outcome-source columns, dropped per column to prevent leakage.",
          "Per head rather than per column recovers them without weakening the guarantee: a "
          "symptom score is a label for one head and a covariate for the other two."],
         size=14)
    stat(s, Inches(10.1), Inches(2.2), "99.6%", "of cohort B have the symptom score", TEAL, 32)
    stat(s, Inches(10.1), Inches(3.7), "every\nstep", "the event count is populated", CRIMSON, 26)

    # 5 --------------------------------------------------------- finding 3
    s = slide(prs, "finding 3", "The covariate panel is a two-phase sample, not a gap")
    s.shapes.add_picture(str(FIG / "fig3_two_phase.png"), M, Inches(2.0), width=Inches(9.2))
    text(s, Inches(10.1), Inches(2.2), Inches(2.7), Inches(4.2),
         ["3.53% of cohort B carry any clinical covariate — and the same people carry all of "
          "them (Jaccard 0.9940).",
          "Selection rate varies 8-fold across arms and over-samples cases by 4–9 points.",
          "So imputation would import that enrichment. A two-phase estimator is required."],
         size=13, color=MUTED)

    # 6 ------------------------------------------------------- revised plan
    s = slide(prs, "revised approach", "Four stages, each with its own falsifiable claim")
    table(s, M, Inches(2.05), CW,
          ["Stage", "Mechanism", "Claim under test", "Control"],
          [["1  Shared encoder", "time-aware attention on elapsed time",
            "structural steps are padding", "remove vs retain vs 3-state mask"],
           ["2  Disease conditioning", "per-outcome dynamics, shared patient state",
            "conditioning helps without negative transfer", "capacity-matched shared dynamics"],
           ["3  Heterogeneity", "mixture of experts / trajectory prototypes",
            "modes are stable and identifiable", "synthetic data with planted modes"],
           ["4  Outcome-matched heads", "continuous / binary / zero-inflated count",
            "outcome representation changes which model wins", "same targets, same horizons"]],
          [2.3, 3.0, 3.4, 3.4], size=12)
    text(s, M, Inches(5.4), CW, Inches(1.2),
         ["Hypotheses, not results. The tree baseline currently matches or beats every sequence "
          "model on this database across all nine cells, and 55–91% of its gain comes from "
          "visit-schedule features. Every claim above is measured against it, paired, on one "
          "machine."],
         size=13, color=MUTED)

    # 7 ---------------------------------------------------------- timeline
    s = slide(prs, "updated timeline", "Week 1 is new; it follows directly from finding 2")
    table(s, M, Inches(2.0), CW,
          ["Week", "Dates", "Deliverable"],
          [["1", "Nov 2 – 8", "Rebuild features per head; recover the symptom score and the "
                                   "event count; report zero inflation and true-zero vs missing"],
           ["2", "Nov 9 – 15", "Synthetic generator; test whether planted dynamics are recovered"],
           ["3", "Nov 16 – 22", "Baselines on identical folds; paired comparison against trees"],
           ["4", "Nov 23 – 29", "Grid ablation and outcome-representation ablation"],
           ["5", "Nov 30 – Dec 6", "Public benchmark; capacity and injection controls"],
           ["6", "Dec 7 – 14", "Write the paper; all tables generated from result files"],
           ["7", "Dec 15 – 20", "Final package: slides, code, documentation"],
           ["—", "Dec 21", "Submit"]],
          [0.9, 2.0, 9.2], size=12)

    assert _page["n"] == 7, f"expected 7 slides, built {_page['n']}"
    prs.save(OUT)
    print(f"[cv] wrote {OUT}  ({_page['n']} slides)")


if __name__ == "__main__":
    build()
