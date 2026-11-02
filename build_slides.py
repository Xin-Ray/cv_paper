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
        tb = s.shapes.add_textbox(M, y, CW, Inches(0.34)).text_frame
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
         ["Standard risk models give every disease the same shape. Cox models assume one "
          "hazard form. Deep sequence models learn one shared dynamic and attach three "
          "output heads to it. Neither one lets the disease decide how time works."],
         size=16, lead=8)
    card(s, M, Inches(5.62), CW, Inches(1.12),
         "The question this project asks",
         ["Can one framework learn patient-specific and disease-specific latent dynamics, "
          "find the distinct trajectory subtypes inside each disease, and improve personalized "
          "prediction for diseases that behave differently in time?"],
         fill=TEAL_BG, tsize=14, bsize=15)

    # 3 ----------------------------------------------------------- the plan
    s = slide(prs, "2 · proposed method", "DCPTL: four stages",
              "Disease-Conditioned Personalized Trajectory Learning. The patient state is "
              "shared across diseases. The dynamics are not.")
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
          "The same latent state evolves under three different operators.",
          "This is the main claim of the project."], accent=CRIMSON, bsize=12.5)
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
         ["Two things are new here, and both are hypotheses I still have to test. "
          "First, the dynamics are conditioned on the disease instead of shared. "
          "Second, heterogeneity is modeled twice: between patients, and between "
          "trajectory subtypes inside one disease.",
          "Target venue: ICML. The contribution has to be a general method for dynamics "
          "and heterogeneity, not a better score on one cohort."],
         size=15, lead=8, color=MUTED)

    # 4 ------------------------------------------------------- what I did
    s = slide(prs, "3 · progress", "I audited the data before training anything",
              "The proposal put an architecture comparison first. I ran the data audit "
              "first instead, and it changed the plan.")
    bullets(s, M, Inches(2.28), Inches(7.6), Inches(3.4),
            [("Why.", "Stage 2 compares three dynamics. If the three outcomes do not actually "
                      "carry comparable observations, that comparison measures the data, not "
                      "the dynamics."),
             ("What I measured.", "Every column of the harmonized database: how many time "
                                  "steps exist, how many are really observed, which columns "
                                  "survive into the model input, and who has a covariate value."),
             ("What I found.", "Three properties that would have confounded the first "
                               "experiment. None of them needs new data to fix."),
             ("", ""),
             ("Finding 1.", "11 of 19 time steps carry no measurement."),
             ("Finding 2.", "The two densest variable blocks are deleted before the model "
                            "sees them."),
             ("Finding 3.", "The covariate panel is a two-phase sample, and it over-samples "
                            "cases.")],
            size=15.5, lead=11)
    stat(s, Inches(8.65), Inches(2.35), "3", "findings that changed the plan")
    stat(s, Inches(8.65), Inches(3.70), "0", "that need new data")
    stat(s, Inches(8.65), Inches(5.05), "+1 week", "added to the timeline", CRIMSON, 30)

    # 5 --------------------------------------------------------- finding 1
    s = slide(prs, "4 · finding 1", "11 of 19 time steps carry no measurement",
              "The time grid is the union of two source calendars, so most steps exist "
              "for the other cohort, not for this one.")
    s.shapes.add_picture(str(FIG / "fig1_grid_audit.png"), M, Inches(2.25), width=Inches(9.2))
    bullets(s, Inches(10.05), Inches(2.35), Inches(2.75), Inches(4.3),
            [("What the plot shows.", "One row per variable block, one column per time step. "
                                      "Only the dark cells hold data."),
             ("Why it matters.", "The missingness mask writes 0 in three different "
                                 "situations: no visit was scheduled, the patient came but "
                                 "this variable was not measured, and the patient was due "
                                 "and did not come."),
             ("The third case is informative.", "A missed visit is a fact about the patient. "
                                                "Right now the model cannot tell it from an "
                                                "empty calendar year.")],
            size=12.5, lead=9)

    # 6 --------------------------------------------------------- finding 2
    s = slide(prs, "5 · finding 2", "The two densest variable blocks are deleted",
              "They are deleted to prevent label leakage. The deletion is applied per "
              "column, so all three outcomes lose them.")
    table(s, M, Inches(2.30), Inches(8.5),
          ["Variable block", "Time steps", "Observations per patient", "Reaches the model"],
          [["Clinical panel (13 columns)", "19", "1.00", "yes"],
           ["Diet quality (24 columns)", "19", "2.38", "yes"],
           ["Continuous symptom score", "19", "4.63", "no"],
           ["Cumulative event count", "19", "19.00", "no"]],
          [3.5, 1.3, 2.3, 1.4], size=13.5, rowh=0.40)
    text(s, M, Inches(4.45), Inches(8.5), Inches(2.0),
         ["The symptom score is the label for depression. The event count is the label for "
          "cardiovascular disease. Neither is a label for diabetes.",
          "Fix: delete per prediction head instead of per column. A variable that is a label "
          "for one head stays a covariate for the other two. The leakage guarantee is "
          "unchanged, and Stage 1 gets 4.6 observations per patient instead of 1."],
         size=14.5, lead=9)
    stat(s, Inches(9.6), Inches(2.40), "99.6%", "of cohort B have the symptom score", TEAL, 30, 3.1)
    stat(s, Inches(9.6), Inches(3.85), "19 / 19", "steps hold the event count", CRIMSON, 30, 3.1)
    stat(s, Inches(9.6), Inches(5.30), "1.00", "observations in the block we do keep",
         MUTED, 30, 3.1)

    # 7 --------------------------------------------------------- finding 3
    s = slide(prs, "6 · finding 3", "The covariate panel is a two-phase sample",
              "Only 3.53% of cohort B has any clinical covariate, and the sampling rate "
              "depends on the enrollment arm.")
    s.shapes.add_picture(str(FIG / "fig3_two_phase.png"), M, Inches(2.25), width=Inches(9.1))
    bullets(s, Inches(9.95), Inches(2.35), Inches(2.85), Inches(4.3),
            [("It is not scattered missingness.", "The same 1,670 people carry all of the "
                                                  "covariates. Jaccard overlap between age "
                                                  "and BMI is 0.9940."),
             ("The rate varies 8-fold.", "1.1% in the observational arm, 9.7% in one trial "
                                         "arm. The plot shows all four arms."),
             ("The measured subset holds more cases.", "Case rate is 4.1 to 9.4 points higher "
                                                       "than in the unmeasured group."),
             ("So imputation is wrong here.", "Filling the values in would copy that case "
                                              "enrichment into the training set. A two-phase "
                                              "estimator with sampling weights is the correct "
                                              "treatment.")],
            size=12.5, lead=9)

    # 8 --------------------------------------------- how the plan changed
    s = slide(prs, "7 · revision", "What each finding changes in DCPTL",
              "The method is unchanged. What changed is the order of the experiments and "
              "one input decision in each stage.")
    table(s, M, Inches(2.32), CW,
          ["Finding", "Stage affected", "Change", "How I will check it"],
          [["11 of 19 steps are empty", "Stage 1 encoder",
            "Give the mask three states instead of two: not scheduled, scheduled and "
            "unmeasured, due and missed",
            "Compare two-state and three-state masks on the same folds"],
           ["Dense blocks are deleted", "Stage 1 input",
            "Delete per prediction head, not per column",
            "Rebuild the features and re-measure observations per patient"],
           ["Two-phase covariate sample", "Stage 4 objective",
            "Weight the likelihood by the inverse sampling rate instead of imputing",
            "Compare weighted and imputed fits; report both"],
           ["Trees currently win", "Whole model",
            "Treat the visit schedule as a modeled input, not a nuisance",
            "Paired comparison against gradient-boosted trees in every cell"]],
          [2.7, 1.9, 3.7, 3.8], size=12, rowh=0.52)
    card(s, M, Inches(5.48), CW, Inches(1.32), "The result that set this agenda",
         ["Gradient-boosted trees match or beat every sequence model I have trained, in all "
          "nine cohort-by-outcome cells. 55 to 91 percent of their gain comes from features "
          "built out of the visit schedule. So the schedule carries signal, and a model that "
          "treats it as noise is leaving that signal on the table."],
         fill=AMBER_BG, accent=CRIMSON, tsize=14, bsize=14.5)

    # 9 ------------------------------------------------------- experiments
    s = slide(prs, "8 · experiments", "Four experiments, and what would falsify each one",
              "Synthetic data first, so I can check that the method recovers dynamics I "
              "planted myself.")
    table(s, M, Inches(2.32), Inches(8.9),
          ["Experiment", "Question", "Falsified if"],
          [["E1  Disease dynamics",
            "Do the three diseases need different dynamics?",
            "Capacity-matched shared dynamics does as well"],
           ["E2  Heterogeneity",
            "Are the trajectory subtypes stable and interpretable?",
            "Subtypes do not repeat across folds or seeds"],
           ["E3  Predictive gain",
            "Does DCPTL beat conventional risk models?",
            "Trees or Cox match it within the detectable effect"],
           ["E4  Generalization",
            "Does it transfer to another dataset?",
            "The gain only exists on this cohort"]],
          [2.4, 3.6, 2.9], size=12.5, rowh=0.56)
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
    s = slide(prs, "9 · updated timeline", "Week 1 is new, and the rest shifts by one week",
              "Week 1 exists because of Finding 2. Everything after it was already planned.")
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
