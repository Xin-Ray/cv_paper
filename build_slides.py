"""Build the COM6005 midterm presentation (5 minutes, 10 slides).

Written to three rules the second version got wrong:

1. Give the audience context. Slides 2 and 3 say what the problem is and what the
   proposed method is, in full, before any result is shown. A classmate who has not
   read the proposal can follow the talk.
2. Plain English. Short declarative sentences, concrete nouns, no metaphors and no
   inverted "not X but Y" constructions. Every slide headline states a fact.
3. The proposed framework (DCPTL) is the spine of the talk, not a footnote. Slide 3
   gives its four stages, slide 8 says how the data findings changed it, slide 9 gives
   the four experiments and the baselines.

Cohort names never appear. The database is the iPAT harmonized national database with
sources cohort A and cohort B.

Run: python cv/build_slides.py
"""
from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Emu, Inches, Pt

HERE = Path(__file__).resolve().parent
FIG = HERE / "figures"
OUT = HERE / "midterm_slides.pptx"

N_SLIDES = 10

BG = RGBColor(0xFB, 0xFA, 0xF7)
INK = RGBColor(0x0C, 0x18, 0x2B)
MUTED = RGBColor(0x4E, 0x5B, 0x6B)
TEAL = RGBColor(0x00, 0x6E, 0x66)
TEAL_BG = RGBColor(0xE4, 0xF1, 0xEF)
CRIMSON = RGBColor(0x9E, 0x1C, 0x44)
AMBER_BG = RGBColor(0xFB, 0xEF, 0xDC)
RULE = RGBColor(0xDF, 0xE5, 0xEC)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

W, H = Inches(13.333), Inches(7.5)
M = Inches(0.62)
CW = W - 2 * M

_page = {"n": 0}


# ----------------------------------------------------------------- primitives
def slide(prs, eyebrow, title, subtitle=None):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = BG
    _page["n"] += 1

    tb = s.shapes.add_textbox(M, Inches(0.40), CW, Inches(0.3)).text_frame
    tb.word_wrap = True
    r = tb.paragraphs[0].add_run()
    r.text = eyebrow.upper()
    r.font.size, r.font.bold, r.font.color.rgb, r.font.name = Pt(12), True, TEAL, "Aptos"

    tb = s.shapes.add_textbox(M, Inches(0.74), CW, Inches(0.8)).text_frame
    tb.word_wrap = True
    r = tb.paragraphs[0].add_run()
    r.text = title
    r.font.size, r.font.bold, r.font.color.rgb, r.font.name = Pt(30), True, INK, "Aptos Display"

    y = Inches(1.52)
    if subtitle:
        tb = s.shapes.add_textbox(M, y, CW, Inches(0.40)).text_frame
        tb.word_wrap = True
        r = tb.paragraphs[0].add_run()
        r.text = subtitle
        r.font.size, r.font.color.rgb, r.font.name = Pt(15), MUTED, "Aptos"
        y = Inches(1.92)

    ln = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, M, y, CW, Emu(9525))
    ln.fill.solid(); ln.fill.fore_color.rgb = RULE; ln.line.fill.background()
    ln.shadow.inherit = False

    ft = s.shapes.add_textbox(M, H - Inches(0.50), CW, Inches(0.3)).text_frame
    p = ft.paragraphs[0]; p.alignment = PP_ALIGN.RIGHT
    r = p.add_run(); r.text = f"{_page['n']} / {N_SLIDES}"
    r.font.size, r.font.color.rgb, r.font.name = Pt(10), MUTED, "Aptos"
    return s


def text(s, x, y, w, h, lines, size=16, color=INK, lead=10, bold=()):
    tf = s.shapes.add_textbox(x, y, w, h).text_frame
    tf.word_wrap = True
    for i, t in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(lead)
        r = p.add_run(); r.text = t
        r.font.size = Pt(size)
        r.font.color.rgb = color
        r.font.name = "Aptos"
        r.font.bold = i in bold
    return tf


def bullets(s, x, y, w, h, items, size=15, lead=11):
    """items: list of (lead-in, rest). The lead-in is bold, the rest is plain."""
    tf = s.shapes.add_textbox(x, y, w, h).text_frame
    tf.word_wrap = True
    for i, (head, rest) in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(lead)
        if head:
            r = p.add_run(); r.text = head
            r.font.size, r.font.bold, r.font.color.rgb, r.font.name = Pt(size), True, INK, "Aptos"
        if rest:
            r = p.add_run(); r.text = (" " if head else "") + rest
            r.font.size, r.font.color.rgb, r.font.name = Pt(size), MUTED, "Aptos"
    return tf


def card(s, x, y, w, h, title, lines, fill=WHITE, accent=TEAL, tsize=15, bsize=13):
    box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    box.fill.solid(); box.fill.fore_color.rgb = fill
    box.line.color.rgb = RULE; box.line.width = Pt(1)
    box.shadow.inherit = False
    box.adjustments[0] = 0.06
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = tf.margin_right = Inches(0.17)
    tf.margin_top = Inches(0.13)
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = title
    r.font.size, r.font.bold, r.font.color.rgb, r.font.name = Pt(tsize), True, accent, "Aptos"
    for t in lines:
        p = tf.add_paragraph(); p.space_before = Pt(5)
        r = p.add_run(); r.text = t
        r.font.size, r.font.color.rgb, r.font.name = Pt(bsize), INK, "Aptos"
    return box


def arrow(s, x, y, w=Inches(0.26), h=Inches(0.26)):
    a = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, x, y, w, h)
    a.fill.solid(); a.fill.fore_color.rgb = RULE
    a.line.fill.background(); a.shadow.inherit = False
    return a


def stat(s, x, y, value, label, color=TEAL, vsize=34, w=3.0):
    tf = s.shapes.add_textbox(x, y, Inches(w), Inches(1.2)).text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = value
    r.font.size, r.font.bold, r.font.color.rgb, r.font.name = Pt(vsize), True, color, "Aptos Display"
    p = tf.add_paragraph()
    r = p.add_run(); r.text = label
    r.font.size, r.font.color.rgb, r.font.name = Pt(12.5), MUTED, "Aptos"


def table(s, x, y, w, header, rows, colw, size=12, rowh=0.36):
    n = len(rows) + 1
    t = s.shapes.add_table(n, len(header), x, y, w, Inches(rowh * n)).table
    for i, cw in enumerate(colw):
        t.columns[i].width = Inches(cw)
    for j, htxt in enumerate(header):
        c = t.cell(0, j); c.text = ""
        r = c.text_frame.paragraphs[0].add_run(); r.text = htxt
        r.font.size, r.font.bold, r.font.color.rgb, r.font.name = Pt(size), True, INK, "Aptos"
    for i, row in enumerate(rows, 1):
        for j, v in enumerate(row):
            c = t.cell(i, j); c.text = ""
            r = c.text_frame.paragraphs[0].add_run(); r.text = str(v)
            r.font.size, r.font.color.rgb, r.font.name = Pt(size), INK, "Aptos"
    return t


# --------------------------------------------------------------------- slides
def build():
    prs = Presentation()
    prs.slide_width, prs.slide_height = W, H

    # 1 -------------------------------------------------------------- title
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.background.fill.solid(); s.background.fill.fore_color.rgb = BG
    _page["n"] += 1
    tf = s.shapes.add_textbox(M, Inches(1.95), CW, Inches(1.9)).text_frame
    tf.word_wrap = True
    r = tf.paragraphs[0].add_run()
    r.text = "Disease-Conditioned Personalized\nTrajectory Learning"
    r.font.size, r.font.bold, r.font.color.rgb, r.font.name = Pt(42), True, INK, "Aptos Display"
    text(s, M, Inches(3.75), CW, Inches(0.5),
         ["One model that learns a different time behavior for each disease, "
          "and different patient subtypes inside each disease."],
         size=18, color=TEAL)
    text(s, M, Inches(4.65), CW, Inches(1.6),
         ["Midterm Project Progress  ·  COM6005 Computer Vision",
          "Xin Xiang  ·  Katz School of Science and Health, Yeshiva University",
          "November 2, 2026"], size=15, color=MUTED, lead=5)
    text(s, M, Inches(6.55), CW, Inches(0.5),
         ["Data access and domain supervision: H. Fang, NIDDK R01DK129432"],
         size=11, color=MUTED)

    # 2 ------------------------------------------------------------ context
    s = slide(prs, "1 · the problem", "Three diseases, three different kinds of time",
              "I predict three outcomes from the same patient records. Each one moves "
              "through time in a different way.")
    y = Inches(2.22)
    card(s, M, y, Inches(3.98), Inches(1.95), "Cardiovascular disease",
         ["A rare event.",
          "I predict whether the event happens and when.",
          "Risk accumulates over decades."])
    card(s, M + Inches(4.18), y, Inches(3.98), Inches(1.95), "Diabetes",
         ["A state that changes.",
          "I predict onset and progression.",
          "The metabolic state is observable before onset."])
    card(s, M + Inches(8.36), y, Inches(3.98), Inches(1.95), "Depression",
         ["A level that moves up and down.",
          "I predict the level, how long it lasts, and relapse.",
          "The same patient can cross the threshold many times."])
    text(s, M, Inches(4.55), CW, Inches(1.0),
         ["Standard risk models give every disease the same shape. A Cox model assumes one "
          "hazard form for all three diseases. A multi-task sequence model learns one shared "
          "dynamic and attaches three output heads to that one dynamic. Neither the Cox model "
          "nor the shared-dynamic model lets the disease change how time works."],
         size=16, lead=8)
    card(s, M, Inches(5.62), CW, Inches(1.12),
         "The question this project asks",
         ["Can one framework learn patient-specific and disease-specific latent dynamics, "
          "find the distinct trajectory subtypes inside each disease, and improve personalized "
          "prediction for diseases that behave differently in time?"],
         fill=TEAL_BG, tsize=14, bsize=15)

    # 3 ----------------------------------------------------------- the plan
    s = slide(prs, "2 · proposed method", "DCPTL: four stages",
              "Proposed framework. The patient state is shared across diseases; the dynamics "
              "are not.")
    bw, bh, gap = Inches(2.80), Inches(2.62), Inches(0.26)
    bx = M
    by = Inches(2.30)
    card(s, bx, by, bw, bh, "Stage 1\nShared encoder",
         ["Input: one patient's visits, at irregular years, with most values missing.",
          "Time-aware attention over elapsed time, or a Neural ODE.",
          "Output: a latent health state shared by all three diseases."], bsize=12.5)
    arrow(s, bx + bw + Inches(0.03), by + Inches(1.1))
    bx += bw + gap
    card(s, bx, by, bw, bh, "Stage 2\nDisease-conditioned dynamics",
         ["One learnable dynamics function per disease, conditioned on a disease code.",
          "The same latent state evolves under three different operators."],
         accent=CRIMSON, bsize=12.5)
    arrow(s, bx + bw + Inches(0.03), by + Inches(1.1))
    bx += bw + gap
    card(s, bx, by, bw, bh, "Stage 3\nHeterogeneity",
         ["A mixture of experts, or a set of learned trajectory prototypes.",
          "One disease holds several trajectory subtypes.",
          "A patient can match more than one, with weights."], bsize=12.5)
    arrow(s, bx + bw + Inches(0.03), by + Inches(1.1))
    bx += bw + gap
    card(s, bx, by, bw, bh, "Stage 4\nMatched prediction heads",
         ["Cardiovascular: event time and survival probability.",
          "Diabetes: binary onset over a fixed horizon.",
          "Depression: next symptom level, duration, and state change."], bsize=12.5)
    text(s, M, Inches(5.05), CW, Inches(1.4),
         ["Status: DCPTL is specified here and built from week 1. No stage is trained yet, "
          "so no number on the following slides comes from DCPTL.",
          "Two things are new. First, the dynamics are conditioned on the disease instead of "
          "shared — the same patient state evolves under three learned operators. Second, "
          "heterogeneity is modeled twice: between patients, and between trajectory subtypes "
          "inside one disease. Each is stated as a claim with its own control, and tested in "
          "E1 and E2.",
          "Target venue: ICML. The contribution is a general method for dynamics and "
          "heterogeneity, demonstrated on synthetic data, a national cohort and a public "
          "benchmark."],
         size=15, lead=8, color=MUTED)

    # 4 ------------------------------------------------------- what is built
    s = slide(prs, "3 · where the project stands", "Two tracks: one finished, one specified",
              "The system DCPTL will run inside is built and measured. DCPTL itself is "
              "specified and starts in week 1.")
    y = Inches(2.26)
    card(s, M, y, Inches(3.90), Inches(2.05), "Data platform — done",
         ["Four-layer pipeline, raw to model input.",
          "52,244 participants, 9 outcome cells.",
          "Subject-level 5-fold splits, leakage controlled."], bsize=13.5)
    card(s, M + Inches(4.10), y, Inches(3.90), Inches(2.05), "Evaluation protocol — done",
         ["Patient-weighted PR-AUC as primary metric.",
          "Minimum detectable effect per cell: 0.0011–0.0107.",
          "Visit-index reference row beside every result."], bsize=13.5)
    card(s, M + Inches(8.20), y, Inches(3.90), Inches(2.05), "Benchmark suite — done",
         ["Nine existing architectures trained, all 9 cells.",
          "Gradient-boosted trees as the bar to clear.",
          "All code public and reproducible."], bsize=13.5)
    text(s, M, Inches(4.62), Inches(7.6), Inches(2.0),
         ["The six weeks produced three measurements on the system already running. Each one "
          "fixes a DCPTL design choice that the proposal had left open:",
          "Measurement 1 — the visit schedule carries information, so Stage 1 will encode the "
          "schedule with a three-state mask.",
          "Measurement 2 — a per-head feature set is exact, so Stage 1 will read 4.63 "
          "observations per patient instead of 1.00.",
          "Measurement 3 — the covariate panel is a two-phase sample, so Stage 4 will use an "
          "inverse-probability weighted objective."],
         size=15, lead=8)
    stat(s, Inches(9.1), Inches(4.70), "3", "DCPTL design choices fixed by measurement")
    stat(s, Inches(9.1), Inches(5.95), "0", "requiring new data collection")

    # 5 ------------------------------------------------- result 1 · stage 1
    s = slide(prs, "4 · measurement 1 → stage 1", "The visit schedule, measured and modeled",
              "I mapped all 19 time steps, and the map turns the missingness mask into a "
              "three-state signal.")
    s.shapes.add_picture(str(FIG / "fig1_grid_audit.png"), M, Inches(2.25), width=Inches(9.2))
    bullets(s, Inches(10.05), Inches(2.35), Inches(2.75), Inches(4.3),
            [("The map.", "One row per variable block, one column per time step; dark cells "
                          "hold data. 8 of 19 steps carry measurement, 11 are calendar "
                          "positions from the other cohort."),
             ("The design rule it fixes.", "Stage 1 will take a three-state mask: not "
                                           "scheduled, scheduled and unmeasured, scheduled "
                                           "and missed."),
             ("Why this is a gain.", "A missed visit is information about the patient. Giving a "
                                     "missed visit its own mask state lets the encoder use the "
                                     "missed visit as a feature. Gradient-boosted trees on this "
                                     "data already draw 55–91% of their gain from "
                                     "visit-schedule features.")],
            size=12.5, lead=9)

    # 6 ------------------------------------------------- result 2 · stage 1
    s = slide(prs, "5 · measurement 2 → stage 1 input",
              "Each prediction head gets its own feature set",
              "A variable defines the label for at most one of the three heads.")
    table(s, M, Inches(2.45), Inches(8.6),
          ["Variable block", "Obs. per patient", "Defines the label for",
           "Used as a covariate by"],
          [["Clinical panel (13 columns)", "1.00", "no head", "all three heads"],
           ["Diet quality (24 columns)", "2.38", "no head", "all three heads"],
           ["Symptom score (CES-D)", "4.63", "the depression head",
            "the diabetes and cardiovascular heads"],
           ["Cumulative event count", "19.00", "the cardiovascular head",
            "the depression and diabetes heads"]],
          [3.3, 1.7, 1.8, 1.8], size=13.5, rowh=0.42)
    text(s, M, Inches(5.20), Inches(8.6), Inches(1.5),
         ["The leakage guarantee is exact: a prediction head never receives the variable that "
          "defines that head’s own label.",
          "The diabetes head therefore reads the symptom score and the event count as "
          "covariates, which takes Stage 1 from 1.00 observation per patient to 4.63."],
         size=14.5, lead=9)
    stat(s, Inches(9.7), Inches(2.50), "1.00 → 4.63", "observations per patient into Stage 1",
         TEAL, 27, 3.1)
    stat(s, Inches(9.7), Inches(3.95), "99.6%", "coverage of the symptom score", TEAL, 30, 3.1)
    stat(s, Inches(9.7), Inches(5.40), "19 / 19", "steps hold the event count", TEAL, 30, 3.1)

    # 7 ------------------------------------------------- result 3 · stage 4
    s = slide(prs, "6 · measurement 3 → stage 4",
              "The covariate panel is a two-phase design",
              "I recovered the sampling design: who was measured, and at what rate, is "
              "known exactly.")
    s.shapes.add_picture(str(FIG / "fig3_two_phase.png"), M, Inches(2.30), width=Inches(9.1))
    bullets(s, Inches(9.95), Inches(2.40), Inches(2.85), Inches(4.3),
            [("The design is clean.", "One phase-two subsample of 1,670 carries every "
                                      "covariate together — age, sex, race, education, blood "
                                      "pressure, glucose, Jaccard 0.9940."),
             ("The selection rate is known.", "1.1% to 9.7% by enrollment arm, measured "
                                              "exactly, so the inverse-probability weight is "
                                              "computable per participant."),
             ("The design rule it fixes.", "Stage 4 will weight the likelihood of each participant by the "
                                  "inverse of that participant’s selection rate — the standard "
                                  "two-phase estimator."),
             ("Why the weighted estimator.", "The weighted estimator trains on all 47,297 "
                                            "participants and stays unbiased for the whole "
                                            "cohort.")],
            size=12.5, lead=9)

    # 8 ----------------------------------------------- the evaluation setup
    s = slide(prs, "7 · evaluation protocol", "How every DCPTL claim will be measured",
              "The protocol is fixed before the first model is trained.")
    bullets(s, M, Inches(2.32), Inches(6.3), Inches(4.3),
            [("Splits.", "Subject-level five-fold cross-validation. No patient appears on "
                         "both sides of a fold, and all scaling is fitted inside the training "
                         "folds."),
             ("Primary metric.", "Patient-weighted PR-AUC — one vote per patient, so a "
                                 "frequently seen patient cannot dominate. Mean and standard "
                                 "deviation over the five folds."),
             ("Matched to the outcome.", "Concordance and zero-inflated deviance for the "
                                         "cardiovascular head, longitudinal logistic for "
                                         "diabetes, two-part continuous for depression."),
             ("Paired comparison.", "Every contrast runs on identical folds, seeds and patient "
                                    "weights, with Holm correction across the nine cells.")],
            size=14.5, lead=10)
    card(s, Inches(7.25), Inches(2.32), Inches(5.47), Inches(2.05),
         "Two reference rows printed beside every result",
         ["Visit-index predictor — a model that reads only the index of the visit scores "
          "0.6918 on pooled depression, where the prevalence floor is 0.2899. Every results "
          "table carries this row.",
          "Minimum detectable effect — 0.0011 to 0.0107 paired, per cell. A difference below "
          "it is reported as not detectable."], tsize=14, bsize=13.5)
    card(s, Inches(7.25), Inches(4.60), Inches(5.47), Inches(2.02),
         "The bar DCPTL has to clear",
         ["Gradient-boosted trees over trajectory summaries, trained on the same rows, folds, "
          "weights and early-stopping rule, across all nine cells. 55 to 91 percent of their "
          "gain comes from visit-schedule features — which is exactly the signal Stage 1's "
          "three-state mask is designed to capture."],
         fill=TEAL_BG, tsize=14, bsize=13.5)

    # 9 ------------------------------------------------------- experiments
    s = slide(prs, "8 · experiments", "Four experiments, each with its control fixed in advance",
              "None of these has run yet. Weeks 2 to 5 run them, synthetic data first.")
    table(s, M, Inches(2.32), Inches(8.9),
          ["Experiment", "What it establishes", "Control it is measured against"],
          [["E1  Disease dynamics",
            "that the three diseases need different dynamics",
            "capacity-matched shared dynamics, same folds"],
           ["E2  Heterogeneity",
            "that trajectory subtypes are stable and interpretable",
            "planted subtypes on synthetic data, cross-fold agreement"],
           ["E3  Predictive gain",
            "that disease conditioning improves prediction",
            "Cox, multi-state and gradient-boosted trees, paired"],
           ["E4  Generalization",
            "that the method transfers beyond this cohort",
            "a public benchmark of irregular clinical series"]],
          [2.4, 3.4, 3.1], size=12.5, rowh=0.56)
    bullets(s, Inches(9.75), Inches(2.40), Inches(3.05), Inches(4.2),
            [("Data.", "Synthetic trajectories with planted subtypes; the harmonized "
                       "national database, cohort A and cohort B; one public EHR benchmark "
                       "for reproducibility."),
             ("Baselines.", "Cox, multi-state models, XGBoost, Transformer, Neural ODE, "
                            "mixture of experts."),
             ("Metric.", "Patient-weighted PR-AUC, five folds, mean and standard deviation."),
             ("Guardrail.", "The paired minimum detectable effect on this data is 0.0011 to "
                            "0.0107. I report a difference as real only above it.")],
            size=12.5, lead=10)

    # 10 --------------------------------------------------------- timeline
    s = slide(prs, "9 · updated timeline", "Seven weeks to submission",
              "Week 1 turns Result 2 into the feature build; weeks 2 to 7 were already "
              "planned.")
    table(s, M, Inches(2.30), CW,
          ["Week", "Dates", "Deliverable", "Status"],
          [["1", "Nov 2 – 8",
            "Rebuild features per prediction head. Recover the symptom score and the event "
            "count. Report zero inflation, and separate true zeros from missing values.", "new"],
           ["2", "Nov 9 – 15",
            "Synthetic generator with planted dynamics and planted subtypes. Test whether "
            "Stage 2 and Stage 3 recover them.", "planned"],
           ["3", "Nov 16 – 22",
            "Baselines on identical folds: Cox, multi-state, XGBoost, Transformer, Neural ODE.",
            "planned"],
           ["4", "Nov 23 – 29",
            "E1 and E2 on real data. Mask ablation and outcome-representation ablation.",
            "planned"],
           ["5", "Nov 30 – Dec 6",
            "E4 on the public benchmark. Capacity controls and sampling-weight controls.",
            "planned"],
           ["6", "Dec 7 – 14",
            "Write the paper. Every table generated from a result file.", "planned"],
           ["7", "Dec 15 – 20",
            "Final package: report, slides, code, documentation.", "planned"],
           ["—", "Dec 21", "Submit.", "planned"]],
          [0.75, 1.65, 7.75, 1.0], size=11.5, rowh=0.455)
    text(s, M, Inches(6.62), CW, Inches(0.5),
         ["Code: github.com/Xin-Ray/cv_paper   ·   every number on these slides is "
          "reproduced by a script in that repository."],
         size=13, color=MUTED)

    assert _page["n"] == N_SLIDES, f"expected {N_SLIDES} slides, built {_page['n']}"
    prs.save(OUT)
    print(f"[cv] wrote {OUT}  ({_page['n']} slides)")


if __name__ == "__main__":
    build()
