"""Build the COM6005 midterm progress report (.docx).

Follows the course report template: Title / Authors / Abstract / 1 Introduction (1.1, 1.2) /
2 ... / 5 Project Timeline (table) / References.

Every quantitative claim is a measurement taken on 2026-11-01 from the project's L1/L2 stores.
The cohort is referred to only as the iPAT harmonized national database; its two sources are
cohort A (state-level geography) and cohort B (region-level geography).

Run: python cv/build_report.py
"""
from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor

HERE = Path(__file__).resolve().parent
FIG = HERE / "figures"
OUT = HERE / "MIDTERM_REPORT.docx"

INK = RGBColor(0x0C, 0x18, 0x2B)
MUTED = RGBColor(0x56, 0x65, 0x77)
FONT = "Times New Roman"


def setup(doc: Document) -> None:
    s = doc.sections[0]
    s.page_width, s.page_height = Inches(8.5), Inches(11)
    s.top_margin = s.bottom_margin = Inches(1.0)
    s.left_margin = s.right_margin = Inches(1.0)
    st = doc.styles["Normal"]
    st.font.name = FONT
    st.font.size = Pt(11)
    st.paragraph_format.space_after = Pt(6)
    st.paragraph_format.line_spacing = 1.15


def para(doc, text, *, size=11, bold=False, italic=False, align="just", after=6, color=INK):
    p = doc.add_paragraph()
    p.alignment = {"just": WD_ALIGN_PARAGRAPH.JUSTIFY, "c": WD_ALIGN_PARAGRAPH.CENTER,
                   "l": WD_ALIGN_PARAGRAPH.LEFT}[align]
    p.paragraph_format.space_after = Pt(after)
    r = p.add_run(text)
    r.font.name, r.font.size, r.bold, r.italic = FONT, Pt(size), bold, italic
    r.font.color.rgb = color
    return p


def head(doc, text, level=1):
    size = {1: 13, 2: 11.5}[level]
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12 if level == 1 else 9)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text)
    r.font.name, r.font.size, r.bold = FONT, Pt(size), True
    r.font.color.rgb = INK


def table(doc, header, rows, widths=None, caption=None, fontsize=9.0):
    t = doc.add_table(rows=1, cols=len(header))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(header):
        c = t.rows[0].cells[i]
        c.text = ""
        r = c.paragraphs[0].add_run(h)
        r.font.name, r.font.size, r.bold = FONT, Pt(fontsize), True
        r.font.color.rgb = INK
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            r = cells[i].paragraphs[0].add_run(str(v))
            r.font.name, r.font.size = FONT, Pt(fontsize)
            r.font.color.rgb = INK
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Inches(w)
    if caption:
        para(doc, caption, size=8.8, italic=True, align="l", after=10, color=MUTED)
    return t


def figure(doc, name, caption, width=6.4):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    p.add_run().add_picture(str(FIG / name), width=Inches(width))
    para(doc, caption, size=8.8, italic=True, align="l", after=10, color=MUTED)


# ------------------------------------------------------------------ content
def build() -> None:
    doc = Document()
    setup(doc)

    para(doc, "Learning Disease-Conditioned Dynamics from Sparse, Irregularly Observed "
              "Longitudinal Health Records", size=16, bold=True, align="c", after=4)
    para(doc, "COM6005 Computer Vision — Midterm Project Progress Report", size=11,
         italic=True, align="c", after=4, color=MUTED)
    para(doc, "Xin Xiang", size=11, align="c", after=0)
    para(doc, "Katz School of Science and Health, Yeshiva University, New York, NY", size=10,
         align="c", after=0, color=MUTED)
    para(doc, "xxiang@mail.yu.edu", size=10, align="c", after=2, color=MUTED)
    para(doc, "November 1, 2026", size=10, align="c", after=12, color=MUTED)

    head(doc, "Abstract")
    para(doc,
         "The proposal for this project argued that a longitudinal cohort is a three-way array over "
         "time, place and measured variable, and that the three axes deserve operators matched to "
         "their geometry rather than a single flattening. Six weeks of work on the data have not "
         "overturned that argument, but they have changed what the first experiment should be. An "
         "audit of the observation grid shows that the encoder unrolls nineteen time steps over a "
         "record that contains, on average, between one and 2.38 real measurements, and that eleven "
         "of those nineteen steps contain no measurement of any kind because the grid is the union "
         "of two source calendars. A second audit shows that the two densest longitudinal signals in "
         "the database, a continuous symptom score observed up to eight times per participant and a "
         "cumulative event count observed at every step, are removed in their entirety by the "
         "current feature build because they are outcome-source columns. A third shows that the "
         "clinical covariate panel is present for 3.53 percent of one source cohort and for a single "
         "time point in the other, and that the measured subsample over-samples cases by four to "
         "nine percentage points. We therefore revise the project from an architecture comparison "
         "toward learning disease-conditioned dynamics under an informative observation process, and "
         "we report the revised design, baselines, evaluation protocol and week-by-week timeline.",
         size=10)

    head(doc, "1  Introduction")
    head(doc, "1.1  Problem Specification", 2)
    para(doc,
         "A longitudinal cohort record can be written as a three-way array whose axes index time, "
         "place of residence, and the measured variables. The proposal observed that these axes are "
         "not interchangeable: time is ordered but unevenly spaced, place is a graph rather than a "
         "line, and the variable axis has no canonical order at all, so a convolution along it is "
         "not even well defined. It proposed an axis-faithful encoder that identifies each "
         "coordinate by a learned identity vector rather than by its position, and that attends "
         "along one axis at a time with an operator matched to that axis.")
    para(doc,
         "That framing is unchanged. What has changed is the empirical picture of the object the "
         "encoder is meant to consume. The work reported in Section 2 was undertaken to size the "
         "tensor before building operators for it, and it found three properties that make a "
         "straightforward architecture comparison uninformative as a first experiment. Section 3 "
         "states the revised approach, which keeps the per-axis identity commitment but moves the "
         "scientific question from “which operator is better” to “what dynamics are "
         "identifiable when the observation process is itself informative.”")

    head(doc, "1.2  Significance", 2)
    para(doc,
         "Three diseases with different temporal behaviour are modelled here: a cardiovascular "
         "outcome that accumulates as a count of events, a metabolic outcome that is naturally "
         "binary, and a psychiatric outcome that is measured on a continuous symptom scale and can "
         "remit and recur. Treating all three as the same binarised onset label, which is what the "
         "current pipeline does, discards the distinctions that make them interesting. The "
         "methodological question generalises beyond health: whenever several heterogeneous "
         "response types share one irregularly observed covariate history, the representation of the "
         "outcome and the representation of the observation process interact, and neither can be "
         "assessed without the other.")

    head(doc, "2  Progress Since the Proposal")
    para(doc,
         "All measurements below were taken on 2026-11-01 directly from the canonical and "
         "model-shaped stores of the iPAT harmonized national database. The database pools two "
         "source cohorts, denoted cohort A (5,114 participants, state-level geography) and cohort B "
         "(47,297 participants, region-level geography); pooled, 52,244 participants, of whom 90.2 "
         "percent come from cohort B.")

    head(doc, "2.1  The modelled grid is a calendar union, not an observation schedule", 2)
    para(doc,
         "Cohort A is represented by exactly nineteen rows per participant and cohort B by exactly "
         "fifteen, irrespective of how often that participant was actually seen. Tabulating what is "
         "recorded at each step shows that the diet battery of twenty-four columns is populated at "
         "three steps only, the thirteen-column clinical panel at one step only, and the continuous "
         "symptom score at six. Eleven of the nineteen steps carry no measurement at all. Those "
         "eleven correspond to follow-up years of the other source cohort: because the two "
         "calendars are aligned onto one national index, each cohort receives rows at the other's "
         "visit years.")
    figure(doc, "fig1_grid_audit.png",
           "Figure 1. Percent of cohort A participants with a recorded value at each of the "
           "nineteen modelled time steps. Red boxes mark the eleven steps that carry no measurement "
           "of any kind. Measured 2026-11-01 from the canonical store.")
    para(doc,
         "This has a consequence for the missingness encoding. The pipeline uses the standard "
         "mask-and-time-gap representation, in which a mask of zero means “not observed.” "
         "At step zero the mask behaves as intended: it is 0.999 for diet, marking four participants "
         "who attended the examination but have no diet record. At the eleven empty steps the mask "
         "is identically zero for every variable and every participant, which encodes something "
         "entirely different, namely that no collection was scheduled. A third situation, a "
         "participant who was due for an examination and did not attend, is also encoded as zero. "
         "The representation therefore collapses a structural property of the study calendar, an "
         "administrative gap, and a behaviour that is plausibly related to health into one symbol. "
         "The third of these is exactly the quantity the revised project proposes to model.")
    para(doc,
         "The time-gap channel inherits the same defect. It reports the elapsed time since the last "
         "observation, computed across intervals that were never observation opportunities, so a "
         "participant seen at two consecutive examinations can present a gap of thirteen years. "
         "This is a candidate explanation for a result already recorded in the project: the decay "
         "variant of the recurrent encoder is the weakest sequence model tested here (0.7552 "
         "patient-weighted PR-AUC) and is beaten by a plain recurrent encoder given the same "
         "channels as raw inputs (0.7909).")

    head(doc, "2.2  The densest longitudinal signals are removed before modelling", 2)
    para(doc,
         "Table 1 compares the number of time steps the encoder unrolls with the number of real "
         "observations each block contributes. Two blocks fill every step. Both are classified as "
         "outcome-source columns and are dropped in their entirety when the model-shaped store is "
         "built, a rule applied per column in order to prevent label leakage.")
    table(doc,
          ["Variable block", "Steps unrolled", "Observations per participant", "Status"],
          [["Clinical panel, cohort A (13 columns)", "19", "1.00 (all at baseline)", "used"],
           ["Diet quality, cohort A (24 columns)", "19", "2.38 (3 calendar years)", "used"],
           ["Diet quality, cohort B (12 columns)", "15", "2.35 (max 6)", "used"],
           ["Continuous symptom score, cohort B", "15", "2.83 (max 8), 99.6% of participants", "removed"],
           ["Continuous symptom score, cohort A", "19", "4.63 (6 calendar years)", "removed"],
           ["Cumulative event count, cohort B", "15", "15.00 (every step)", "removed"],
           ["Cumulative event count, cohort A", "19", "19.00 (every step)", "removed"]],
          widths=[2.9, 1.0, 1.9, 0.8],
          caption="Table 1. Modelled sequence length against realised observation count, measured "
                  "2026-11-01. The two blocks that populate every step are precisely the two that "
                  "the feature build removes.")
    figure(doc, "fig2_shape_mismatch.png",
           "Figure 2. The same comparison drawn to scale. Sequence length is determined by the "
           "harmonisation grid; information content is not.")
    para(doc,
         "The leakage rule is correct in intent and too coarse in execution. A continuous symptom "
         "score is the source of the psychiatric label and must not enter that head as a feature, "
         "but it is a legitimate and well-populated covariate for the metabolic and cardiovascular "
         "heads. The same applies in reverse to the event count and to thirteen cardiovascular "
         "subtype histories. Applying the rule per prediction head rather than per column recovers "
         "this information without weakening the leakage guarantee. No discussion of a per-head "
         "relaxation exists in the project record, so we treat it as an open design item rather "
         "than an oversight.")
    para(doc,
         "The removed columns are also the two response types that the project's domain advisor has "
         "asked be modelled on their own scale: the psychiatric outcome as a continuous trajectory "
         "and the cardiovascular outcome as a zero-inflated count, with the metabolic outcome "
         "remaining binary. Those three representations are available in the data as it stands; "
         "they are discarded by configuration, not absent.")

    head(doc, "2.3  The covariate panel follows a two-phase sampling design", 2)
    para(doc,
         "In cohort B, 1,670 of 47,297 participants (3.53 percent) carry any clinical covariate, and "
         "the same participants carry all of them: age, sex, race, education, systolic pressure and "
         "glucose agree on the same set to a Jaccard index of 0.9940, and no participant has a body "
         "mass index without an age. The missingness is therefore block-structured by data source "
         "rather than scattered, which rules out the imputation strategies used in the published "
         "literature on these cohorts, where covariate missingness runs between one and eight "
         "percent and is handled by complete-case deletion or by chained-equation imputation.")
    figure(doc, "fig3_two_phase.png",
           "Figure 3. Left: the share of each enrolment arm that carries a clinical panel, which "
           "ranges from 1.13 to 9.65 percent. Right: excess event rate in the measured subsample.")
    para(doc,
         "Selection into the measured subsample is neither uniform nor independent of outcome. The "
         "rate varies by a factor of eight across enrolment arms, and the measured subsample carries "
         "4.1, 5.9 and 9.4 additional percentage points of cardiovascular, metabolic and psychiatric "
         "events respectively. This is the signature of a case-cohort design, and it means that an "
         "imputation model fitted on the measured participants would import that enrichment into "
         "the 45,627 participants it is used to fill. The appropriate treatment is a two-phase "
         "estimator with sampling weights; how to combine such weights with multiple imputation is "
         "itself an open methodological question in the statistical literature.")
    para(doc,
         "Cohort A is complete in a different way and incomplete in a different way. Its "
         "demographics are present for 100 percent of participants and its clinical panel for "
         "between 76.6 and 99.96 percent, which matches the completeness reported in published "
         "analyses of this cohort. But every one of those thirteen measurements sits at a single "
         "time point. A derived insulin-resistance index is present at three time points while the "
         "two quantities from which it is computed are present at one, which is arithmetically "
         "impossible unless the repeated inputs exist upstream and were not carried through "
         "harmonisation. Recovering them is a data-provenance task, and we have deliberately "
         "designed the next phase so that it does not depend on the outcome.")

    head(doc, "2.4  What the three findings change", 2)
    para(doc,
         "None of the three findings is a reason to abandon the proposal's premise. Each is a "
         "reason to reorder the experiments, because each identifies a confound that would have "
         "contaminated the architecture comparison that was planned as the first study. Table 2 "
         "states the reordering explicitly.")
    table(doc,
          ["Finding", "Why it blocks the planned first experiment", "What it becomes instead"],
          [["Grid is a calendar union (2.1)",
            "Two operators compared on a sequence that is 58 percent structural padding differ "
            "mostly in how they tolerate padding, not in how they use the axes.",
            "An ablation in its own right: remove, retain, or name the empty steps with a third "
            "mask state."],
           ["Densest blocks removed (2.2)",
            "The comparison would be run on 1.00 to 2.38 observations per participant, which "
            "cannot identify the multi-step dynamics the method is meant to express.",
            "A feature-store rebuild with per-head rules, which is Week 1 of the revised plan."],
           ["Two-phase covariate sample (2.3)",
            "Any model given the clinical panel is fitted on 3.53 percent of one cohort, and that "
            "subsample over-samples cases, so apparent gains are partly selection.",
            "A design constraint: the main line uses only blocks that are present for all "
            "participants, and the panel enters through a weighted two-phase analysis."]],
          widths=[1.5, 2.8, 2.3],
          caption="Table 2. Each finding converts a planned comparison into a prerequisite "
                  "experiment. None of them requires data the project does not already hold.")
    para(doc,
         "It is worth being explicit about what has not changed. The per-axis identity commitment "
         "survives, because nothing measured here bears on it. The claim that the variable axis has "
         "no canonical order survives. The honesty constraint recorded in the proposal also "
         "survives: a feature-axis attention module of exactly the kind proposed has already been "
         "evaluated on this database and did not help, reaching 0.7815 against 0.7934 for gradient "
         "boosted trees, a paired difference of −0.0120 at p = 0.038. The revised project "
         "therefore does not claim that attending over the variable axis is novel. It claims that "
         "the joint structure, the outcome-matched heads, and the treatment of the observation "
         "process are worth more than any single axis operator, and it is designed so that this "
         "claim can fail visibly.")

    head(doc, "3  Revised Approach: Disease-Conditioned Trajectory Learning")
    para(doc,
         "The revision keeps the proposal's two design commitments, identity rather than position "
         "and factorisation rather than fusion, and adds a third: the observation process is part "
         "of the model, not a nuisance to be corrected away. The resulting framework has four "
         "stages.")
    para(doc,
         "Stage 1, a shared trajectory encoder. Irregular observations are encoded into a latent "
         "health state by a time-aware attention operator keyed on elapsed time rather than on step "
         "index, with the structurally empty steps of Section 2.1 either removed from the sequence "
         "or marked by a third mask state. Which of the two is preferable is an experiment, not an "
         "assumption.")
    para(doc,
         "Stage 2, disease-conditioned dynamics. Rather than one risk function shared by all three "
         "outcomes, a conditioning mechanism supplies a separate learnable dynamics function per "
         "outcome while the patient representation remains shared. The falsifiable claim is that "
         "this helps without negative transfer; the control is a capacity-matched shared-dynamics "
         "model, because a gain obtainable by adding parameters is not evidence for conditioning.")
    para(doc,
         "Stage 3, heterogeneity-aware pattern learning. A mixture of experts, or a set of learnable "
         "trajectory prototypes, allows one disease to be represented by several dynamic modes with "
         "per-participant weights. The claim to be tested is that the recovered modes are stable "
         "across folds and identifiable from the available sampling density, not merely that they "
         "improve a score.")
    para(doc,
         "Stage 4, outcome-appropriate prediction. Each head is given a likelihood matched to its "
         "response type: a continuous head with a two-part component for exact zeros, a longitudinal "
         "logistic head, and a zero-inflated count head. All three predict the same trajectory "
         "targets over the same horizons so that the comparison across outcomes is meaningful.")
    para(doc,
         "We state plainly which parts of this are hypotheses. That multiple dynamic modes exist and "
         "are recoverable is unverified. That disease conditioning avoids negative transfer is "
         "unverified. That the sampling density in this database supports identifying more than one "
         "mode is doubtful for the diet block, which has two to three observations per participant, "
         "and is the main reason the experimental design begins with synthetic data of known "
         "mechanism.")

    head(doc, "4  Experimental Design")
    head(doc, "4.1  Data", 2)
    table(doc,
          ["Source", "Role", "Scale", "Why it is needed"],
          [["Synthetic generator", "mechanism recovery",
            "configurable; reversible, irreversible, periodic and multi-modal dynamics",
            "the only setting with ground-truth dynamics, so a negative result is interpretable"],
           ["Public benchmark (MIMIC-style irregular clinical series)", "reproducibility",
            "public, widely used", "lets the method be compared without access to our cohort"],
           ["iPAT harmonized national database", "target application",
            "52,244 participants; 545,883 to 654,037 scored rows per outcome",
            "the long horizon and three heterogeneous outcomes this project is about"]],
          widths=[1.6, 1.1, 1.8, 2.1],
          caption="Table 3. Three data sources with distinct roles. A method that succeeds only on "
                  "the third is not a methodological contribution.")

    head(doc, "4.2  Baselines", 2)
    para(doc,
         "Proportional-hazards regression, a multi-state model, gradient-boosted trees over "
         "trajectory summaries, a transformer encoder, a neural ordinary differential equation, and "
         "a mixture-of-experts variant. The tree baseline is not a formality: on this database it "
         "currently matches or exceeds every sequence model tested across all nine cohort-by-outcome "
         "cells, and attribution shows that between 55 and 91 percent of its gain comes from "
         "features describing the visit schedule rather than from any physiological measurement. "
         "Any claim made by a sequence model must be made against it.")

    head(doc, "4.3  Evaluation protocol", 2)
    para(doc,
         "Subject-level five-fold cross-validation with mean and standard deviation; a "
         "patient-weighted ranking metric for binary heads, with outcome-appropriate metrics for the "
         "continuous and count heads, since binarising a continuous response to reuse one metric is "
         "not defensible; paired fold-level tests with Holm correction across cells; subject-level "
         "bootstrap intervals. Every difference is reported against the cell's minimum detectable "
         "effect, which in this project ranges from 0.0011 to 0.0107 for paired comparisons. A "
         "visit-index-only predictor is reported beside every result: it reaches 0.6918 on the "
         "pooled psychiatric outcome against 0.7379 for the strongest sequence model, so scoring "
         "against the prevalence floor of 0.2899 overstates the model roughly ninefold.")

    head(doc, "4.4  Ablations", 2)
    table(doc,
          ["Ablation", "Question it answers"],
          [["Shared vs disease-conditioned dynamics, capacity matched",
            "does conditioning help, or does capacity?"],
           ["Single vs multiple dynamic modes", "is within-disease heterogeneity identifiable?"],
           ["Structural steps removed vs retained vs third mask state",
            "how much of the sequence is padding, and does naming it help?"],
           ["Outcome-source columns per column vs per head",
            "what is the cost of the current leakage rule?"],
           ["Binary landmark vs continuous and count heads",
            "does the outcome representation change which model wins?"],
           ["Axis shuffle and synthetic effect injection",
            "is the detector sensitive, and is it sensitive to the right thing?"]],
          widths=[3.1, 3.5],
          caption="Table 4. Each row is a comparison on identical folds, seeds and patient weights, "
                  "so that every difference is paired.")

    head(doc, "5  Updated Project Timeline")
    table(doc,
          ["Week", "Dates", "Deliverable"],
          [["1", "Nov 2 – Nov 8",
            "Rebuild the feature store with per-head outcome-source rules; recover the continuous "
            "score and the event count. Report their distributions, zero inflation, and the "
            "separation of true zeros from missing values."],
           ["2", "Nov 9 – Nov 15",
            "Synthetic generator with known dynamics. Fit shared, conditioned and multi-mode "
            "variants; test whether the planted mechanism is recovered."],
           ["3", "Nov 16 – Nov 22",
            "Baselines on identical folds: hazards, multi-state, trees, transformer, neural ODE, "
            "mixture of experts. Paired comparison against the tree baseline."],
           ["4", "Nov 23 – Nov 29",
            "Grid ablation: structural steps removed, retained, or marked by a third mask state. "
            "Outcome-representation ablation across the three response types."],
           ["5", "Nov 30 – Dec 6",
            "Public benchmark replication; capacity-matched controls; axis shuffle and injection "
            "controls."],
           ["6", "Dec 7 – Dec 14",
            "Write the paper. All tables and figures generated directly from result files."],
           ["7", "Dec 15 – Dec 20",
            "Final submission package: slides, code, data documentation, user documentation."],
           ["—", "Dec 21", "Submit."]],
          widths=[0.5, 1.3, 4.8],
          caption="Table 5. Revised timeline. Week 1 is new relative to the proposal and is a direct "
                  "consequence of Section 2.2.")

    head(doc, "6  Risks and Honest Limitations")
    para(doc,
         "The most likely failure is that disease-conditioned dynamics do not beat a capacity-matched "
         "shared-dynamics model. This is planned for rather than hoped against: the deliverable in "
         "that case is a quantified negative result with a stated detectability bar, together with "
         "the grid audit and the outcome-representation comparison, neither of which depends on the "
         "conditioning mechanism winning.")
    para(doc,
         "Three further risks are already quantified from earlier work on this database. First, "
         "per-group model selection in a related experiment produced gains on validation folds in "
         "every one of eighteen cells and lost to a single global model in sixteen of them "
         "out-of-fold, which is a clear precedent for selection overfitting in this setting. "
         "Second, re-running identical code on a different accelerator moves results by up to 0.0056 "
         "in the five-fold mean and 0.0244 in a single fold, which is the same order as several "
         "effects we might wish to claim; comparisons are therefore run paired, on one machine. "
         "Third, the diet block offers two to three observations per participant, which is unlikely "
         "to identify multiple continuous dynamic modes; this is why the continuous symptom score, "
         "at up to eight observations, is the block on which Stage 3 will first be attempted.")
    para(doc,
         "A limitation of scope should also be recorded. The clinical panel is single-time-point in "
         "one source cohort and present for 3.53 percent of the other. Recovering repeated clinical "
         "measurements is being pursued through the data-provenance route described in Section 2.3, "
         "but no part of the design above depends on that recovery succeeding.")

    head(doc, "Acknowledgment")
    para(doc,
         "Data access and domain supervision are provided by H. Fang, supported by NIDDK "
         "R01DK129432. The views here are the author's own.", size=10)

    head(doc, "References")
    refs = [
        "A. Sharma, E. Vans, D. Shigemizu, K. A. Boroevich, and T. Tsunoda. DeepInsight: A "
        "methodology to transform a non-image data to an image for convolution neural network "
        "architecture. Scientific Reports, 9:11399, 2019.",
        "O. Bazgir, R. Zhang, S. R. Dhruba, R. Rahman, S. Ghosh, and R. Pal. Representation of "
        "features as images with neighborhood dependencies for compatibility with convolutional "
        "neural networks. Nature Communications, 11:4391, 2020.",
        "Y. Zhu, T. Brettin, F. Xia, A. Partin, M. Shukla, H. Yoo, Y. A. Evrard, J. H. Doroshow, and "
        "R. L. Stevens. Converting tabular data into images for deep learning with convolutional "
        "neural networks. Scientific Reports, 11:11325, 2021.",
        "J. Lee, Y. Lee, J. Kim, A. R. Kosiorek, S. Choi, and Y. W. Teh. Set Transformer: A "
        "framework for attention-based permutation-invariant neural networks. ICML, 2019.",
        "Y. Gorishniy, I. Rubachev, V. Khrulkov, and A. Babenko. Revisiting deep learning models for "
        "tabular data. NeurIPS 34, 2021.",
        "S. Tipirneni and C. K. Reddy. Self-supervised transformer for sparse and irregularly "
        "sampled multivariate clinical time-series. ACM TKDD, 16(6):1–17, 2022.",
        "J. Ho, N. Kalchbrenner, D. Weissenborn, and T. Salimans. Axial attention in "
        "multidimensional transformers. arXiv:1912.12180, 2019.",
        "A. Arnab, M. Dehghani, G. Heigold, C. Sun, M. Lučić, and C. Schmid. ViViT: A video "
        "vision transformer. ICCV, 2021.",
        "G. Bertasius, H. Wang, and L. Torresani. Is space-time attention all you need for video "
        "understanding? ICML, 2021.",
        "K. He, X. Chen, S. Xie, Y. Li, P. Dollár, and R. Girshick. Masked autoencoders are "
        "scalable vision learners. CVPR, 2022.",
        "C. D. Bahadir, A. Q. Wang, A. V. Dalca, and M. R. Sabuncu. Deep-learning-based optimization "
        "of the under-sampling pattern in MRI. IEEE Transactions on Computational Imaging, 6:1139–"
        "1152, 2020.",
        "R. T. Q. Chen, Y. Rubanova, J. Bettencourt, and D. Duvenaud. Neural ordinary differential "
        "equations. NeurIPS 31, 2018.",
        "Y. Rubanova, R. T. Q. Chen, and D. Duvenaud. Latent ODEs for irregularly-sampled time "
        "series. NeurIPS 32, 2019.",
        "Z. Che, S. Purushotham, K. Cho, D. Sontag, and Y. Liu. Recurrent neural networks for "
        "multivariate time series with missing values. Scientific Reports, 8:6085, 2018.",
        "C. Lee, W. R. Zame, J. Yoon, and M. van der Schaar. DeepHit: A deep learning approach to "
        "survival analysis with competing risks. AAAI, 2018.",
        "I. Perros, E. E. Papalexakis, R. Vuduc, E. Searles, and J. Sun. Temporal phenotyping of "
        "medically complex children via PARAFAC2 tensor factorization. Journal of Biomedical "
        "Informatics, 93:103125, 2019.",
        "T. G. Kolda and B. W. Bader. Tensor decompositions and applications. SIAM Review, "
        "51(3):455–500, 2009.",
        "Z. Yue, Y. Wang, J. Duan, T. Yang, C. Huang, Y. Tong, and B. Xu. TS2Vec: Towards universal "
        "representation of time series. AAAI, 2022.",
        "J. Dong, H. Wu, H. Zhang, L. Zhang, J. Wang, and M. Long. SimMTM: A simple pre-training "
        "framework for masked time-series modeling. NeurIPS 36, 2023.",
        "R. H. Keogh and I. R. White. Multiple imputation of missing data in nested case-control and "
        "case-cohort studies. Biometrics, 74(4):1438–1449, 2018.",
        "P. Buzkova and T. Lumley. Semiparametric modeling of repeated measurements under "
        "outcome-dependent follow-up. Statistics in Medicine, 28(6):987–1003, 2009.",
        "J. Adebayo, J. Gilmer, M. Muelly, I. Goodfellow, M. Hardt, and B. Kim. Sanity checks for "
        "saliency maps. NeurIPS 31, 2018.",
    ]
    for i, r in enumerate(refs, 1):
        para(doc, f"[{i}]  {r}", size=9.0, after=3)

    doc.save(OUT)
    print(f"[cv] wrote {OUT}")


if __name__ == "__main__":
    build()
