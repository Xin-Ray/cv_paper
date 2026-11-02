"""Build MIDTERM_REPORT.docx directly inside the course template.

The course supplies project_midterm_report_template.doc: a two-column CVPR-style layout,
Times New Roman 10 pt, 1.0 in top / 1.13 in bottom / 1.0 in left / 0.63 in right margins,
column width 3.279 in with a 0.315 in gutter. The template says in as many words: do not change
the font size, the line spacing or the margins, and the paper is no less than eight full pages
including text, figures, tables and references.

So this script does not re-create that layout. It opens the template (converted once to
cv/_template.docx), empties the body, and writes the report back using the template's own
styles - Abstract, Heading 1, Heading 2, Text, References. Page setup, fonts and spacing are
inherited untouched.

Section order follows the template exactly:

    Abstract | 1 Introduction | 2 Related Work | 3 Proposed Method (3.1-3.3 existing methods,
    3.4 our method with a pipeline figure) | 4 Experiments (4.1 Datasets, 4.2 Evaluation
    metrics, 4.3 Implementation details, 4.4 Current progress and results) | 5 Conclusion |
    References (at least 20 items)

Cohort names never appear. The database is the iPAT harmonized national database with sources
cohort A and cohort B.

Run: python cv/make_figures.py && python cv/make_pipeline_figure.py && python cv/build_report.py
"""
from __future__ import annotations

import copy
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

HERE = Path(__file__).resolve().parent
FIG = HERE / "figures"
TEMPLATE = HERE / "_template.docx"
OUT = HERE / "MIDTERM_REPORT.docx"

COL_W = Inches(3.20)          # the template's column is 3.279 in
FULL_W = Inches(6.80)         # 8.5 - 1.0 - 0.63 of margin


# --------------------------------------------------------------------- helpers
def clear_body(doc) -> None:
    """Remove every block of the template body, keeping the final sectPr."""
    body = doc.element.body
    sect = body.find(qn("w:sectPr"))
    for tag in ("w:headerReference", "w:footerReference"):
        for el in sect.findall(qn(tag)):
            sect.remove(el)
    for el in list(body):
        if el is not sect:
            body.remove(el)


def _sect_pr(doc, cols: int):
    """A copy of the body section properties with a given column count, continuous."""
    src = doc.element.body.find(qn("w:sectPr"))
    new = copy.deepcopy(src)
    for tag in ("w:headerReference", "w:footerReference"):
        for el in new.findall(qn(tag)):
            new.remove(el)
    t = new.find(qn("w:type"))
    if t is None:
        t = OxmlElement("w:type")
        new.insert(0, t)
    t.set(qn("w:val"), "continuous")
    c = new.find(qn("w:cols"))
    if c is None:
        c = OxmlElement("w:cols")
        new.append(c)
    for child in list(c):
        c.remove(child)
    c.set(qn("w:num"), str(cols))
    c.set(qn("w:space"), "454")
    c.set(qn("w:equalWidth"), "1")
    return new


def break_section(doc, cols: int) -> None:
    """Close the running section with the given column count."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p._p.get_or_add_pPr().append(_sect_pr(doc, cols))


def para(doc, text, style="Text", *, size=None, bold=False, italic=False, align=None,
         indent=None, after=None, before=None):
    p = doc.add_paragraph(style=style)
    r = p.add_run(text)
    if size is not None:
        r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    if align is not None:
        p.alignment = align
    if indent is not None:
        p.paragraph_format.first_line_indent = Pt(indent)
    if after is not None:
        p.paragraph_format.space_after = Pt(after)
    if before is not None:
        p.paragraph_format.space_before = Pt(before)
    return p


def head(doc, text, level=1):
    p = doc.add_paragraph(style=f"Heading {level}")
    r = p.add_run(text)
    r.font.bold = True
    r.font.size = Pt(11 if level == 1 else 10)
    return p


def caption(doc, text):
    p = doc.add_paragraph(style="Text")
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_after = Pt(9)
    r = p.add_run(text)
    r.font.size = Pt(8.5)
    r.font.italic = True
    return p


def figure(doc, name, cap, width=COL_W):
    p = doc.add_paragraph(style="Text")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(3)
    p.add_run().add_picture(str(FIG / name), width=width)
    caption(doc, cap)


def _grid_borders(t) -> None:
    """The course template carries no table style, so draw hairline borders by hand."""
    tblPr = t._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:color"), "808080")
        borders.append(el)
    tblPr.append(borders)


def table(doc, header, rows, widths, cap, fontsize=7.6):
    t = doc.add_table(rows=len(rows) + 1, cols=len(header))
    _grid_borders(t)
    t.autofit = False
    for j, w in enumerate(widths):
        for row in t.rows:
            row.cells[j].width = Inches(w)
    for j, h in enumerate(header):
        c = t.cell(0, j)
        c.text = ""
        r = c.paragraphs[0].add_run(h)
        r.font.size = Pt(fontsize)
        r.font.bold = True
        r.font.name = "Times New Roman"
    for i, row in enumerate(rows, 1):
        for j, v in enumerate(row):
            c = t.cell(i, j)
            c.text = ""
            r = c.paragraphs[0].add_run(str(v))
            r.font.size = Pt(fontsize)
            r.font.name = "Times New Roman"
    for row in t.rows:
        for c in row.cells:
            for p in c.paragraphs:
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(1)
                p.paragraph_format.first_line_indent = Pt(0)
    caption(doc, cap)
    return t


# --------------------------------------------------------------------- content
REFS = [
    "Z. Che, S. Purushotham, K. Cho, D. Sontag, and Y. Liu. Recurrent neural networks for "
    "multivariate time series with missing values. Scientific Reports, 8:6085, 2018.",
    "Y. Rubanova, R. T. Q. Chen, and D. Duvenaud. Latent ODEs for irregularly-sampled time "
    "series. NeurIPS 32, 2019.",
    "R. T. Q. Chen, Y. Rubanova, J. Bettencourt, and D. Duvenaud. Neural ordinary differential "
    "equations. NeurIPS 31, 2018.",
    "P. Kidger, J. Morrill, J. Foster, and T. Lyons. Neural controlled differential equations "
    "for irregular time series. NeurIPS 33, 2020.",
    "S. Tipirneni and C. K. Reddy. Self-supervised transformer for sparse and irregularly "
    "sampled multivariate clinical time-series. ACM TKDD, 16(6):1-17, 2022.",
    "S. N. Shukla and B. M. Marlin. Multi-time attention networks for irregularly sampled time "
    "series. ICLR, 2021.",
    "D. R. Cox. Regression models and life-tables. Journal of the Royal Statistical Society: "
    "Series B, 34(2):187-202, 1972.",
    "C. Lee, W. R. Zame, J. Yoon, and M. van der Schaar. DeepHit: A deep learning approach to "
    "survival analysis with competing risks. AAAI, 2018.",
    "C. Lee, J. Yoon, and M. van der Schaar. Dynamic-DeepHit: A deep learning approach for "
    "dynamic survival analysis with competing risks based on longitudinal data. IEEE "
    "Transactions on Biomedical Engineering, 67(1):122-133, 2020.",
    "H. Putter, M. Fiocco, and R. B. Geskus. Tutorial in biostatistics: Competing risks and "
    "multi-state models. Statistics in Medicine, 26(11):2389-2430, 2007.",
    "T. Chen and C. Guestrin. XGBoost: A scalable tree boosting system. KDD, 2016.",
    "G. Ke, Q. Meng, T. Finley, T. Wang, W. Chen, W. Ma, Q. Ye, and T.-Y. Liu. LightGBM: A "
    "highly efficient gradient boosting decision tree. NeurIPS 30, 2017.",
    "N. Shazeer, A. Mirhoseini, K. Maziarz, A. Davis, Q. Le, G. Hinton, and J. Dean. "
    "Outrageously large neural networks: The sparsely-gated mixture-of-experts layer. ICLR, "
    "2017.",
    "W. Fedus, B. Zoph, and N. Shazeer. Switch Transformers: Scaling to trillion parameter "
    "models with simple and efficient sparsity. JMLR, 23(120):1-39, 2022.",
    "C. Lee and M. van der Schaar. Temporal phenotyping using deep predictive clustering of "
    "disease progression. ICML, 2020.",
    "I. Perros, E. E. Papalexakis, R. Vuduc, E. Searles, and J. Sun. Temporal phenotyping of "
    "medically complex children via PARAFAC2 tensor factorization. Journal of Biomedical "
    "Informatics, 93:103125, 2019.",
    "T. G. Kolda and B. W. Bader. Tensor decompositions and applications. SIAM Review, "
    "51(3):455-500, 2009.",
    "D. Lambert. Zero-inflated Poisson regression, with an application to defects in "
    "manufacturing. Technometrics, 34(1):1-14, 1992.",
    "P. Buzkova and T. Lumley. Semiparametric modeling of repeated measurements under "
    "outcome-dependent follow-up. Statistics in Medicine, 28(6):987-1003, 2009.",
    "R. H. Keogh and I. R. White. Multiple imputation of missing data in nested case-control "
    "and case-cohort studies. Biometrics, 74(4):1438-1449, 2018.",
    "N. E. Breslow and N. Chatterjee. Design and analysis of two-phase studies with binary "
    "outcome applied to Wilms tumour prognosis. Journal of the Royal Statistical Society: "
    "Series C, 48(4):457-468, 1999.",
    "J. Lee, Y. Lee, J. Kim, A. R. Kosiorek, S. Choi, and Y. W. Teh. Set Transformer: A "
    "framework for attention-based permutation-invariant neural networks. ICML, 2019.",
    "Y. Gorishniy, I. Rubachev, V. Khrulkov, and A. Babenko. Revisiting deep learning models "
    "for tabular data. NeurIPS 34, 2021.",
    "A. Sharma, E. Vans, D. Shigemizu, K. A. Boroevich, and T. Tsunoda. DeepInsight: A "
    "methodology to transform a non-image data to an image for convolution neural network "
    "architecture. Scientific Reports, 9:11399, 2019.",
    "Y. Zhu, T. Brettin, F. Xia, A. Partin, M. Shukla, H. Yoo, Y. A. Evrard, J. H. Doroshow, "
    "and R. L. Stevens. Converting tabular data into images for deep learning with "
    "convolutional neural networks. Scientific Reports, 11:11325, 2021.",
    "J. Ho, N. Kalchbrenner, D. Weissenborn, and T. Salimans. Axial attention in "
    "multidimensional transformers. arXiv:1912.12180, 2019.",
    "K. He, X. Chen, S. Xie, Y. Li, P. Dollar, and R. Girshick. Masked autoencoders are "
    "scalable vision learners. CVPR, 2022.",
    "C. D. Bahadir, A. Q. Wang, A. V. Dalca, and M. R. Sabuncu. Deep-learning-based "
    "optimization of the under-sampling pattern in MRI. IEEE Transactions on Computational "
    "Imaging, 6:1139-1152, 2020.",
    "A. Johnson, L. Bulgarelli, L. Shen, A. Gayles, A. Shammout, S. Horng, T. J. Pollard, et "
    "al. MIMIC-IV, a freely accessible electronic health record dataset. Scientific Data, "
    "10:1, 2023.",
    "J. Adebayo, J. Gilmer, M. Muelly, I. Goodfellow, M. Hardt, and B. Kim. Sanity checks for "
    "saliency maps. NeurIPS 31, 2018.",
]


def build() -> None:
    doc = Document(str(TEMPLATE))
    clear_body(doc)

    # ------------------------------------------------- title block, one column
    para(doc, "Disease-Conditioned Personalized Trajectory Learning for Multi-Disease "
              "Prediction from Sparse, Irregularly Observed Longitudinal Records",
         style="Normal", size=16, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER,
         indent=0, before=12, after=10)
    para(doc, "Xin Xiang", style="Normal", size=11, align=WD_ALIGN_PARAGRAPH.CENTER,
         indent=0, after=0)
    para(doc, "Katz School of Science and Health, Yeshiva University, New York, NY",
         style="Normal", size=10, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, after=0)
    para(doc, "xxiang@mail.yu.edu", style="Normal", size=10,
         align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, after=0)
    para(doc, "COM6005 Computer Vision - Midterm Project Progress Report - November 2, 2026",
         style="Normal", size=9, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0,
         after=10)
    break_section(doc, 1)

    # ------------------------------------------------------------------ abstract
    para(doc, "Abstract", style="Normal", size=12, bold=True,
         align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, before=4, after=4)
    para(doc,
         "This project builds one framework, Disease-Conditioned Personalized Trajectory "
         "Learning (DCPTL), that predicts three diseases from the same longitudinal patient "
         "records. The three diseases move through time in different ways. A cardiovascular "
         "outcome is a rare event whose risk accumulates over decades. A metabolic outcome is a "
         "state that changes and is observable before onset. A psychiatric outcome is a level "
         "that rises, falls, remits and returns. Conventional risk models give all three the "
         "same shape, and multi-task sequence models learn one shared dynamic and attach three "
         "output heads to it. DCPTL instead learns a shared patient state, a separate dynamics "
         "function per disease, a set of trajectory subtypes inside each disease, and a "
         "prediction head matched to each response type. The proposal planned an architecture "
         "comparison as the first experiment; we ran a data audit first, and it changed the "
         "plan. Eleven of nineteen modelled time steps hold no measurement of any kind. The two "
         "densest longitudinal signals in the database are deleted in full by the current "
         "feature build. The clinical covariate panel is a two-phase sample holding four to nine "
         "percentage points more cases than the unmeasured group. None of the three findings "
         "needs new data, each changes one input decision in one stage, and the experiments are "
         "reordered so that mechanism recovery on synthetic data precedes comparison on real "
         "data.",
         style="Abstract")

    # -------------------------------------------------------------- introduction
    head(doc, "Introduction")
    para(doc,
         "We predict three diseases from one longitudinal record per patient. The record is "
         "sparse and irregular: visits are years apart, the spacing differs between patients, "
         "and most variables are missing at most visits. The three diseases do not behave the "
         "same way in time. The cardiovascular outcome is a rare event, and the question is "
         "whether it happens and when. The metabolic outcome is a state that changes, and the "
         "question is onset and progression. The psychiatric outcome is a level that moves up "
         "and down, and the question is the level, how long it lasts, and whether it returns.")
    para(doc,
         "Existing methods answer these questions one disease at a time, or answer them together "
         "by forcing one shape on all three. A proportional-hazards model assumes a single "
         "hazard form [7]. A multi-task sequence model learns one shared latent dynamic and "
         "attaches three output heads to it [9]. In both cases the disease is allowed to change "
         "the output but not the way time works. The question this project asks is therefore:")
    para(doc,
         "Can a unified learning framework discover patient-specific and disease-specific latent "
         "dynamics, capture heterogeneous disease trajectories, and improve personalized "
         "predictions across diseases with distinct temporal behaviors?", italic=True)
    para(doc,
         "The answer matters beyond this cohort. Whenever several response types of different "
         "kinds share one irregularly observed covariate history, the representation of the "
         "outcome and the representation of the observation process interact, and neither can be "
         "judged without the other. The pipeline this project inherited turns all three diseases "
         "into the same binarised onset label, which discards exactly what makes them different: "
         "the count structure of the cardiovascular outcome, the state structure of the "
         "metabolic one, and the continuous, recurring structure of the psychiatric one.")
    para(doc,
         "The project proposal approached this through the geometry of the data. It wrote a "
         "cohort record as a three-way array over time, place of residence and measured "
         "variable, observed that the three axes are not interchangeable, and proposed an "
         "encoder that identifies every coordinate by a learned identity vector rather than by "
         "its position. That commitment is unchanged and carries into Stage 1 of the framework "
         "described here. What changed is the order of the work. We ran an audit of the data "
         "before building any operator, and it found three properties that would have made an "
         "architecture comparison uninformative as a first experiment.")
    para(doc, "This report makes four contributions.")
    para(doc,
         "First, it states DCPTL, a four-stage framework in which the patient state is shared "
         "across diseases and the dynamics are not (Section 3.4).")
    para(doc,
         "Second, it reports three measurements on the target database that constrain any model "
         "built on it, each reproducible from a script in the public repository (Section 4.4).")
    para(doc,
         "Third, it states four experiments together with the result that would falsify each "
         "one, so that a negative outcome is interpretable rather than merely disappointing "
         "(Section 4).")
    para(doc,
         "Fourth, it reports a baseline result that sets the difficulty of the problem honestly. "
         "Gradient-boosted trees over trajectory summaries currently match or exceed every "
         "sequence model we have trained, in all nine cohort-by-outcome cells, and between 55 "
         "and 91 percent of their gain comes from features describing the visit schedule rather "
         "than from any physiological measurement.")

    # ------------------------------------------------------------- related work
    head(doc, "Related Work")
    para(doc,
         "Irregular longitudinal modelling. GRU-D [1] handles missing values in clinical time "
         "series by decaying the hidden state toward an empirical mean as the gap since the last "
         "observation grows. Latent ODE models [2, 3] and neural controlled differential "
         "equations [4] replace the discrete recurrence with a continuous-time dynamic, so the "
         "model is defined between observations rather than only at them. STraTS [5] and "
         "multi-time attention networks [6] drop the grid entirely and treat the record as a set "
         "of time, variable and value triplets with a learned embedding for each component. The "
         "last family is closest to the identity-rather-than-position commitment of our "
         "proposal, and Stage 1 of our framework follows it.")
    para(doc,
         "Survival and multi-state models. Cox regression [7] remains the conventional reference "
         "for time-to-event outcomes and is the baseline against which any cardiovascular claim "
         "must be made. DeepHit [8] replaces the proportional-hazards assumption with a discrete "
         "distribution over event times learned directly, and Dynamic-DeepHit [9] extends this "
         "to longitudinal covariates with a shared recurrent encoder and one output head per "
         "cause. Multi-state models [10] represent transitions between disease states explicitly "
         "and are the natural statistical form for the metabolic outcome. These methods "
         "condition the output on the cause; none of them conditions the dynamics on it, which "
         "is the gap Stage 2 addresses.")
    para(doc,
         "Gradient-boosted trees. XGBoost [11] and LightGBM [12] remain strong on tabular "
         "summaries of longitudinal data, and in our own measurements they are the hardest "
         "baseline on this database. Treating them as a formality is a common failure of deep "
         "longitudinal papers, and Section 4.4 reports what happens when they are given the same "
         "folds, the same patient weights and the same early-stopping rule as the sequence "
         "models.")
    para(doc,
         "Heterogeneity and phenotyping. Mixture-of-experts layers [13, 14] route each input to "
         "a subset of specialised sub-networks and are the generic mechanism for representing "
         "several modes within one task. In the clinical setting, deep predictive clustering "
         "[15] learns clusters of disease progression that are predictive rather than merely "
         "descriptive, and PARAFAC2 tensor factorisation [16, 17] recovers temporal phenotypes "
         "from irregular records without assuming a shared time grid. Stage 3 of our framework "
         "borrows from both lines, and a mixture-of-experts variant without disease conditioning "
         "serves as the control that isolates what conditioning adds.")
    para(doc,
         "Outcome representation and informative observation. Zero-inflated models [18] are the "
         "standard treatment for count responses with excess zeros, which is the correct form "
         "for the cardiovascular outcome here. Outcome-dependent follow-up [19] and two-phase or "
         "case-cohort designs [20, 21] address the situation in which who is measured depends on "
         "what happened to them. Section 4.4 shows that the covariate panel of our target "
         "database is exactly such a design, which makes these methods a requirement rather than "
         "a refinement.")
    para(doc,
         "Variable-axis representation. Because the variable axis of a cohort tensor has no "
         "canonical order, a convolution along it is not well defined. Set Transformer [22] and "
         "FT-Transformer [23] handle this by permutation-invariant attention over learned "
         "feature embeddings, while DeepInsight [24] and IGTD [25] instead impose an order by "
         "arranging features into an image so that convolutional architectures apply. Axial "
         "attention [26] factorises attention over the axes of a multidimensional array, which "
         "is the operator our proposal adopts. Masked autoencoding [27] and learned "
         "under-sampling patterns in MRI reconstruction [28] are the computer-vision precedents "
         "for treating the sampling pattern itself as something to model, which is the view this "
         "project takes of the visit schedule.")

    # --------------------------------------------------------- proposed method
    head(doc, "Proposed Method")
    para(doc,
         "This section introduces the three existing methods closest to our setting, states what "
         "each does well and where it breaks on our data, and then states our framework.")

    head(doc, "Existing method 1: GRU-D", 2)
    para(doc,
         "GRU-D [1] augments a gated recurrent unit with two mechanisms. A missingness mask "
         "tells the cell which inputs were observed, and a trainable decay term pulls both the "
         "input and the hidden state toward an empirical mean as the elapsed time since the last "
         "observation of that variable grows. The decay is an explicit prior: the longer a "
         "variable has gone unmeasured, the less its last value should be trusted.")
    para(doc,
         "The advantage is that the model uses the missingness pattern rather than imputing it "
         "away, and the original paper shows that this helps substantially on intensive-care "
         "records, where gaps are measured in hours.")
    para(doc,
         "The limitation appears when the gaps are much longer, and we measured it directly. On "
         "cohort A with five-fold subject-level cross-validation, a plain recurrent encoder that "
         "receives elapsed time as an ordinary input channel reaches 0.7909 patient-weighted "
         "PR-AUC, while GRU-D on the same folds reaches 0.7552. Our gaps are years, not hours, "
         "and missingness runs at 85 to 90 percent, so the decay term saturates almost "
         "immediately and the cell is effectively reset to the mean at every step. The prior "
         "that helps in the intensive-care unit destroys information here. This single result "
         "shaped our Stage 1 design: elapsed time is useful, but as a feature rather than as a "
         "decay mechanism.")

    head(doc, "Existing method 2: Latent ODE", 2)
    para(doc,
         "Latent ODE models [2, 3] encode an irregular sequence into an initial latent state and "
         "then evolve that state continuously by integrating a learned ordinary differential "
         "equation, reading out predictions at whatever times are required. Neural controlled "
         "differential equations [4] extend this so that the incoming data drives the dynamics "
         "rather than only setting the initial condition.")
    para(doc,
         "The advantage for us is direct. The model is defined at every real-valued time, so "
         "nothing forces the record onto a shared grid, and irregular spacing is handled by "
         "construction rather than by padding. This is the correct answer to the problem "
         "reported in Section 4.4, where we show that the grid currently in use is an artefact "
         "of harmonisation rather than a property of the data.")
    para(doc,
         "The limitation is that a latent ODE learns one dynamics function. When three outcomes "
         "share an encoder, the usual construction attaches three read-out heads to a single "
         "trajectory, so the three diseases may differ in what is read off the state but not in "
         "how the state evolves. A second limitation is practical: integration cost grows with "
         "the number of evaluation points, which matters over a thirty-year horizon. Stage 2 of "
         "our framework keeps the continuous-time formulation and makes the dynamics function "
         "itself disease-dependent.")

    head(doc, "Existing method 3: Dynamic-DeepHit", 2)
    para(doc,
         "Dynamic-DeepHit [9] is the closest existing method to our prediction stage. A shared "
         "recurrent encoder consumes the longitudinal covariate history, and a cause-specific "
         "sub-network per competing risk outputs a discrete distribution over event times. It "
         "makes no proportional-hazards assumption and it updates its prediction as new visits "
         "arrive, which is precisely the continually updated personalized prediction this "
         "project wants.")
    para(doc,
         "The advantage is that it already demonstrates several outcomes on one shared patient "
         "representation, with a loss combining a likelihood term and a ranking term.")
    para(doc,
         "The limitation is twofold. First, all competing risks are represented as times to a "
         "first event, which fits our cardiovascular outcome but not a recurring symptom level "
         "or a reversible metabolic state. Second, the cause-specific sub-networks sit after the "
         "shared dynamic, so the heterogeneity they express is between diseases, not between "
         "patient subtypes within one disease. Stages 3 and 4 of our framework are designed "
         "around those two gaps.")

    head(doc, "Our proposed method", 2)
    para(doc,
         "We call our framework Disease-Conditioned Personalized Trajectory Learning, "
         "abbreviated DCPTL. It does not assume that all diseases follow one law of risk. It "
         "learns three things from the longitudinal record: the dynamics of an individual "
         "patient, the dynamics specific to each disease, and the distinct trajectory subtypes "
         "that exist inside a single disease. It then predicts each future outcome from the "
         "patient's own history under that disease's own dynamics. Figure 1 gives the pipeline "
         "and Table 1 states what each stage claims.")
    break_section(doc, 2)
    figure(doc, "fig4_pipeline.png",
           "Figure 1. The DCPTL pipeline. A shared encoder maps one patient's irregular visit "
           "history to a latent health state; a disease-conditioned dynamics function evolves "
           "that state differently for each disease; a mixture component represents several "
           "trajectory subtypes inside one disease; and each outcome receives a prediction head "
           "matched to its response type. Italic labels mark where each of the three audit "
           "findings of Section 4.4 enters the pipeline.", width=FULL_W)
    break_section(doc, 1)
    para(doc,
         "Stage 1, a shared trajectory encoder. Irregular observations are encoded into a latent "
         "health state by a time-aware attention operator keyed on elapsed time rather than on "
         "step index, following [5, 6], with a neural ODE [2] as the continuous-time "
         "alternative. The structurally empty steps identified in Section 4.4 are either removed "
         "from the sequence or marked by a third mask state. Which of the two is preferable is "
         "an experiment, not an assumption.")
    para(doc,
         "Stage 2, disease-conditioned dynamics. Rather than one risk function shared by all "
         "three outcomes, a conditioning mechanism supplies a separate learnable dynamics "
         "function per outcome while the patient representation remains shared. The falsifiable "
         "claim is that this helps without negative transfer, and the control is a "
         "capacity-matched shared-dynamics model, because a gain obtainable by adding parameters "
         "is not evidence for conditioning.")
    para(doc,
         "Stage 3, heterogeneity-aware pattern learning. A mixture of experts [13], or a set of "
         "learnable trajectory prototypes in the spirit of [15, 16], allows one disease to be "
         "represented by several dynamic modes with per-participant weights. The claim to be "
         "tested is that the recovered modes are stable across folds and identifiable at the "
         "available sampling density, not merely that they improve a score.")
    para(doc,
         "Stage 4, matched prediction heads. Each head is given a likelihood matched to its "
         "response type. The cardiovascular head predicts event time and survival probability "
         "with a zero-inflated count component [18] for repeated events. The metabolic head is a "
         "longitudinal logistic model of onset over a fixed horizon. The psychiatric head is "
         "continuous, with a two-part component for exact zeros, and also predicts duration and "
         "the transition between visits. All three predict the same trajectory targets over the "
         "same horizons, so that comparison across diseases is meaningful. Binarising a "
         "continuous response in order to reuse one metric is not appropriate and is not done "
         "here.")
    table(doc,
          ["Stage", "Mechanism", "Claim under test", "Control"],
          [["1. Shared encoder", "time-aware attention on elapsed time, or a neural ODE",
            "the structurally empty steps are padding and can be named as such",
            "two-state against three-state mask, same folds"],
           ["2. Disease-conditioned dynamics",
            "one learnable dynamics function per disease, chosen by a disease code",
            "conditioning helps, and without negative transfer",
            "capacity-matched shared-dynamics model"],
           ["3. Heterogeneity", "mixture of experts, or learnable trajectory prototypes",
            "the modes are stable across folds and identifiable at this density",
            "synthetic data with planted modes"],
           ["4. Matched heads", "survival, longitudinal logistic and two-part likelihoods",
            "the outcome representation changes which model wins",
            "the same targets and horizons under a single binary head"]],
          [0.78, 0.82, 0.84, 0.76],
          "Table 1. The four stages of DCPTL. Every stage carries a claim that an experiment in "
          "Table 2 can falsify.")
    para(doc,
         "We state plainly which parts of DCPTL are hypotheses rather than results. Whether "
         "disease conditioning and double heterogeneity amount to a sufficient methodological "
         "contribution has to be settled by comparison with prior work, by mechanism-recovery "
         "experiments, by ablation and by independent validation, not by assertion. That "
         "multiple dynamic modes exist and are recoverable is unverified. That disease "
         "conditioning avoids negative transfer is unverified. That the sampling density in this "
         "database supports identifying more than one mode is doubtful for the diet block, which "
         "carries two to three observations per participant, and that doubt is the main reason "
         "the experimental design begins with synthetic data of known mechanism.")

    # ----------------------------------------------------------- experiments
    head(doc, "Experiments")
    para(doc,
         "Four experiments test DCPTL, and each is stated together with the result that would "
         "falsify it. They run in the order given: mechanism recovery comes before comparison on "
         "real data, because a method that cannot recover dynamics we planted ourselves cannot "
         "be trusted to have found dynamics we did not.")
    table(doc,
          ["Experiment", "Question", "Setting", "Falsified if"],
          [["E1 Disease dynamics", "do the three diseases need different dynamics?",
            "synthetic data with three planted dynamics, then the real cohort",
            "a capacity-matched shared-dynamics model does as well"],
           ["E2 Patient heterogeneity",
            "are the trajectory subtypes inside one disease stable and interpretable?",
            "synthetic data with planted subtypes, then the real cohort",
            "the recovered subtypes do not repeat across folds or seeds"],
           ["E3 Predictive gain", "does disease conditioning beat conventional risk models?",
            "the real cohort, identical folds, paired comparison",
            "trees or Cox match it within the minimum detectable effect"],
           ["E4 Generalization", "does the method transfer to another dataset?",
            "a public benchmark of irregular clinical series",
            "the gain exists only on our cohort"]],
          [0.80, 0.80, 0.84, 0.76],
          "Table 2. The four experiments with falsification conditions fixed in advance. E1 and "
          "E2 correspond to Stages 2 and 3 of Table 1.")

    head(doc, "Datasets", 2)
    para(doc,
         "Three sources are used, with distinct roles. A method that succeeds only on the third "
         "is not a methodological contribution.")
    para(doc,
         "A synthetic generator produces trajectories with known dynamics and known subtypes: "
         "reversible, irreversible, periodic and multi-modal. It is the only setting with a "
         "ground-truth mechanism, so a negative result there is interpretable, and it is the "
         "only setting in which sampling density can be varied deliberately, which is how we "
         "will establish the density at which Stage 3 stops being identifiable.")
    para(doc,
         "A public benchmark of irregularly sampled clinical series, MIMIC-IV [29], supports "
         "reproduction by other groups, who cannot access our cohort. It is distributed at "
         "https://physionet.org/content/mimiciv/ under a credentialed-access licence.")
    para(doc,
         "The target application is the iPAT harmonized national database, assembled under the "
         "supervision of H. Fang (NIDDK R01DK129432). It pools two source cohorts, denoted here "
         "cohort A (5,114 participants, state-level geography, 19 modelled time steps covering "
         "1985 to 2016) and cohort B (47,297 participants, region-level geography, 15 steps); "
         "pooled, 52,244 participants, of whom 90.2 percent come from cohort B. Three outcomes "
         "are defined on it, giving nine cohort-by-outcome cells with between 545,883 and "
         "654,037 scored participant-visit rows per outcome.")
    para(doc,
         "Preprocessing runs as a four-layer pipeline, and this report's measurements come from "
         "layers one and two. L0 is the raw source extract, which is never modified. L1 "
         "canonicalises column names, units and the time axis, and computes outcome labels. L2 "
         "reshapes the record into the model-facing sequence and removes outcome-source columns "
         "to prevent label leakage; the audit in Section 4.4 is largely about the cost of that "
         "removal rule. L3 builds subject-level five-fold splits, so no participant appears on "
         "both sides of any fold, and all feature scaling is fitted inside the training folds "
         "only.")

    head(doc, "Evaluation metrics", 2)
    para(doc,
         "The primary metric for the binary heads is patient-weighted PR-AUC, in which each "
         "participant contributes one unit of weight regardless of how many visits they "
         "contribute. Positives are rare, between 1.5 and 35 percent depending on the cell, so "
         "ROC-AUC is reported alongside but is not used for ranking. The psychiatric head is "
         "evaluated by root mean squared error and by two-part deviance, and the cardiovascular "
         "head by the concordance index and by zero-inflated Poisson deviance [18]. Calibration "
         "is reported as the Brier score and the expected calibration error, since a well-ranked "
         "but poorly calibrated risk is of limited clinical use. Subtype recovery in E2 is "
         "measured by adjusted Rand index against the planted assignment on synthetic data, and "
         "by cross-fold cluster agreement on real data.")
    para(doc,
         "Every result is reported as a mean and a standard deviation over five subject-level "
         "folds, with paired fold-level tests and Holm correction across cells. Two reference "
         "points accompany every number. The first is the cell's minimum detectable effect, "
         "which on this database ranges from 0.0011 to 0.0107 for paired comparisons; a smaller "
         "difference is reported as not detectable rather than as zero. The second is a "
         "predictor that uses only the index of the visit, learning an event rate per visit "
         "position and nothing else. It reaches 0.6918 on the pooled psychiatric outcome where "
         "the strongest sequence model reaches 0.7379 and the prevalence floor is 0.2899. "
         "Scoring against the prevalence floor therefore overstates the model roughly ninefold, "
         "and the visit-index row is printed beside every result table in this project.")

    head(doc, "Implementation details", 2)
    para(doc,
         "Everything is written in Python 3.12 and PyTorch 2.6.0 with CUDA 12.4, on Windows 11, "
         "with NumPy, pandas, scikit-learn and PyArrow for the data layers and LightGBM [12] for "
         "the tree baseline. The graph attention operator is written by hand rather than taken "
         "from PyTorch Geometric, so the repository carries no compiled geometric dependency. "
         "Training runs on two machines, one with a single GPU and one with four; splits, seeds "
         "and the early-stopping rule are identical on both, and we report a measured "
         "cross-device drift of up to 0.0056 in a five-fold mean, so that results produced on "
         "different machines are not compared naively.")
    para(doc,
         "Model code is shared and configuration-driven. Every architecture is registered by "
         "name in a single registry and selected from a configuration file, so a new model is a "
         "new file plus one registry entry plus one configuration, and no experiment folder "
         "contains a copy of a model class. Early stopping uses a validation fold and the "
         "primary metric. This is not a detail: the tree baseline scores 0.7539 with a fixed "
         "budget and 0.7934 with early stopping on the same cell, and a baseline given worse "
         "treatment than the proposed model is worse than no baseline at all. All figures and "
         "tables in this report are generated by scripts in the public repository at "
         "https://github.com/Xin-Ray/cv_paper.")

    head(doc, "Current progress and results", 2)
    para(doc,
         "Progress since the proposal consists of an audit of the data, carried out before any "
         "operator was built, plus a baseline result that sets the difficulty of the problem. "
         "All measurements below were taken on 2026-11-01 directly from the canonical and "
         "model-shaped stores, and each is reproducible from the measurement log in the "
         "repository.")
    para(doc,
         "Finding 1: eleven of nineteen time steps hold no measurement. The modelled time axis "
         "is the union of the visit calendars of the two source cohorts, so a step exists if "
         "either cohort was seen that year. For cohort A this produces nineteen steps of which "
         "eleven, namely 1986, 1994, 1996 to 1999 and 2001 to 2004, carry no measurement of any "
         "kind. Figure 2 shows the audit. The encoder nevertheless unrolls a recurrence over all "
         "nineteen, and the missingness mask writes a zero at each of them. That zero is "
         "overloaded: it means no visit was scheduled, it means the participant attended but "
         "this variable was not measured, and it means the participant was due and did not "
         "attend. The third case is a fact about the participant and is the one worth modelling; "
         "at present nothing distinguishes it from an empty calendar year. This is why Stage 1 "
         "is tested with a three-state mask.")
    figure(doc, "fig1_grid_audit.png",
           "Figure 2. Percentage of cohort A participants with a recorded value at each of the "
           "nineteen modelled time steps, by variable block. Only the dark cells carry data. The "
           "eleven outlined columns carry no measurement of any kind and exist because the grid "
           "is a union of two cohorts' visit calendars.")
    para(doc,
         "Finding 2: the two densest longitudinal signals are deleted. Table 3 compares the "
         "number of steps the encoder unrolls with the number of real observations behind them. "
         "The clinical panel reaches the model with a single observation per participant. The "
         "two blocks with the most observations, a continuous symptom score with 4.63 "
         "observations per participant in cohort A and 2.83 in cohort B, and a cumulative event "
         "count present at every step in both cohorts, are removed in full before the model sees "
         "them, because the leakage rule deletes outcome-source columns per column. The symptom "
         "score is the label for the psychiatric outcome and the event count is the label for "
         "the cardiovascular one, but neither is a label for the metabolic outcome. Deleting per "
         "prediction head instead of per column keeps the leakage guarantee exactly as strong "
         "while returning both blocks as covariates to the heads for which they are not labels. "
         "Figure 3 draws the same comparison to scale.")
    table(doc,
          ["Variable block", "Steps", "Obs. per participant", "Reaches model"],
          [["Clinical panel, cohort A (13 cols)", "19", "1.00", "yes"],
           ["Diet quality, cohort A (24 cols)", "19", "2.38", "yes"],
           ["Diet quality, cohort B (12 cols)", "15", "2.35", "yes"],
           ["Symptom score, cohort B", "15", "2.83 (max 8)", "no"],
           ["Symptom score, cohort A", "19", "4.63 (max 6)", "no"],
           ["Event count, cohort B", "15", "15.00", "no"],
           ["Event count, cohort A", "19", "19.00", "no"]],
          [1.30, 0.42, 0.90, 0.58],
          "Table 3. Modelled sequence length against realised observation count. The two densest "
          "blocks are the two the current feature build removes.")
    figure(doc, "fig2_shape_mismatch.png",
           "Figure 3. Sequence length against realised observations per participant, by block "
           "and cohort. Length is set by the harmonisation calendar; the observation count is "
           "set by the study protocol, and the two are unrelated.")
    para(doc,
         "Finding 3: the covariate panel is a two-phase sample. In cohort B, 1,670 of 47,297 "
         "participants, or 3.53 percent, carry any clinical covariate, and the same 1,670 carry "
         "all of them: age, sex, race, education, blood pressure and glucose coincide to a "
         "Jaccard index of 0.9940, and the body-mass set is a strict subset of the age set with "
         "no exceptions. This is not scattered missingness that imputation could repair. Figure "
         "4 shows that the selection rate varies roughly eightfold across enrolment arms, from "
         "1.13 percent in the observational arm to 9.65 percent in one trial arm, and that the "
         "measured subsample carries 4.1, 5.9 and 9.4 percentage points more cases than the "
         "unmeasured group for the cardiovascular, metabolic and psychiatric outcomes "
         "respectively, the last at p = 7.1e-44. Imputing the panel from the measured subsample "
         "would therefore import that case enrichment into the training set. The correct "
         "treatment is a two-phase or case-cohort estimator with sampling weights [20, 21], "
         "which is why the Stage 4 objective is weighted rather than imputed. Cohort A fails in "
         "the mirror-image way: its clinical panel covers 76.6 to 99.96 percent of participants "
         "but exists at a single calendar year, and a derived insulin-resistance index appears "
         "at three years while its two inputs appear at one, which proves that repeated "
         "measurements exist upstream and were not carried through harmonisation.")
    figure(doc, "fig3_two_phase.png",
           "Figure 4. Left: the share of each enrolment arm of cohort B that carries a clinical "
           "panel. Right: the excess case rate in the measured subsample relative to the "
           "unmeasured group, in percentage points, for the three outcomes.")
    para(doc,
         "Baseline result. Gradient-boosted trees over trajectory summaries, given the same "
         "rows, the same folds, the same patient weights and the same early-stopping rule as the "
         "sequence models, produce a higher point estimate in eight of the nine "
         "cohort-by-outcome cells, and in four of them the margin exceeds the unpaired minimum "
         "detectable effect. Attribution of their gain shows that features describing the visit "
         "schedule, namely the number of observations, the time since the last observation, the "
         "largest jump between adjacent observations and the visit index, account for 55 to 91 "
         "percent of total split gain, and the observation count alone is the leading single "
         "feature in five of the six larger cells. We do not read this as evidence that sequence "
         "models are useless. We read it as evidence that the visit schedule carries real "
         "signal, so a model that treats it as a nuisance to be corrected away is discarding the "
         "very thing the strongest baseline is using. That is the reasoning behind the "
         "three-state mask in Stage 1 and behind printing a schedule-only reference row in every "
         "results table.")
    para(doc,
         "Table 4 states what each finding changes. The framework itself is unchanged; what "
         "changed is one input decision per stage and the order of the experiments.")
    table(doc,
          ["Finding", "Stage affected", "Change", "Check"],
          [["Eleven of nineteen steps are empty", "Stage 1 encoder",
            "three mask states: not scheduled, scheduled and unmeasured, due and missed",
            "two-state against three-state on the same folds"],
           ["Dense blocks are deleted", "Stage 1 input",
            "delete outcome-source columns per head, not per column",
            "rebuild features and re-measure observations per participant"],
           ["Two-phase covariate sample", "Stage 4 objective",
            "weight the likelihood by the inverse sampling rate instead of imputing",
            "weighted against imputed fits, both reported"],
           ["Trees currently win", "whole model",
            "treat the visit schedule as a modelled input rather than a nuisance",
            "paired comparison against trees in every cell"]],
          [0.86, 0.60, 0.95, 0.79],
          "Table 4. Each audit finding, the stage it constrains, and the experiment that settles "
          "it.")

    head(doc, "Risks and limitations", 2)
    para(doc,
         "Three risks are worth naming now, because each one can make the project's central "
         "claim unverifiable rather than merely false.")
    para(doc,
         "Identifiability. Stage 3 asks for several dynamic modes per disease, but the diet "
         "block carries two to three observations per participant and the clinical panel one. "
         "A mixture is not identifiable from one observation per participant, so for some "
         "variable blocks the honest answer may be that the data cannot support the question. "
         "The synthetic study in week 2 exists to locate that boundary before we spend real "
         "folds on it, by sweeping the sampling density until recovery of the planted modes "
         "fails.")
    para(doc,
         "Model selection on the wrong signal. Per-group model selection in an earlier "
         "experiment in this project was positive on the validation fold in eighteen of "
         "eighteen cells and negative out of fold in sixteen of eighteen. That is a clean "
         "demonstration that a selection procedure can be tuned into an apparent gain, and it "
         "is the reason every comparison reported here is paired on identical folds, with the "
         "selection rule fixed before the result is read.")
    para(doc,
         "Comparability across machines. Identical code on two machines differs by up to "
         "0.0056 in a five-fold mean and by 0.0244 in a single fold, which is larger than "
         "several of the effects this project cares about. Results are therefore only compared "
         "within a machine, and any cross-machine comparison is re-run rather than quoted.")
    para(doc,
         "A fourth limitation is one of scope rather than of risk. The target database is "
         "observational, so every association reported in Section 4.4 is an association and not "
         "a causal effect, and nothing in DCPTL changes that. The methodological claim of the "
         "project is about representation and identifiability under an informative observation "
         "process, not about aetiology.")

    # ------------------------------------------------------------- conclusion
    head(doc, "Conclusion")
    para(doc,
         "Six weeks of work have produced a framework and an audit. The framework, DCPTL, learns "
         "a shared patient state, a separate dynamics function per disease, trajectory subtypes "
         "within each disease, and a prediction head matched to each response type; it is stated "
         "in Section 3.4 with its claims and its controls fixed in advance. The audit measured "
         "the data the framework will consume and found three properties that would have "
         "confounded the architecture comparison the proposal planned to run first: a time grid "
         "in which eleven of nineteen steps are empty, a feature build that deletes the two "
         "densest longitudinal blocks, and a covariate panel that is a two-phase sample with "
         "four to nine points of case enrichment. None of the three requires new data.")
    para(doc,
         "We are equally explicit about what has not been shown. No component of DCPTL has yet "
         "been trained. The strongest model on this database at present is a gradient-boosted "
         "tree baseline whose advantage comes mostly from the visit schedule, and until a "
         "disease-conditioned model beats that baseline on identical folds by more than the "
         "minimum detectable effect, the central claim of this project remains a hypothesis. "
         "Table 5 gives the remaining schedule. Week 1 is new relative to the proposal and is a "
         "direct consequence of Finding 2; the subsequent weeks were already planned and shift "
         "by one week.")
    table(doc,
          ["Week", "Dates", "Deliverable"],
          [["1", "Nov 2-8",
            "Rebuild features per prediction head; recover the symptom score and the event "
            "count; report zero inflation and separate true zeros from missing values"],
           ["2", "Nov 9-15",
            "Synthetic generator with planted dynamics and planted subtypes; test whether Stage "
            "2 and Stage 3 recover them"],
           ["3", "Nov 16-22",
            "Baselines on identical folds: Cox, multi-state, XGBoost, Transformer, neural ODE"],
           ["4", "Nov 23-29",
            "E1 and E2 on the real cohort; mask ablation and outcome-representation ablation"],
           ["5", "Nov 30-Dec 6",
            "E4 on the public benchmark; capacity controls and sampling-weight controls"],
           ["6", "Dec 7-14", "Write the paper; every table generated from a result file"],
           ["7", "Dec 15-20", "Final package: report, slides, code, documentation"],
           ["-", "Dec 21", "Submit"]],
          [0.42, 0.62, 2.16],
          "Table 5. Revised timeline to submission.")
    para(doc,
         "Acknowledgment. Data access and domain supervision were provided by H. Fang, supported "
         "by NIDDK R01DK129432. The views expressed are the author's own.")

    # ------------------------------------------------------------- references
    para(doc, "References", style="Normal", size=11, bold=True, indent=0,
         before=10, after=4)
    for r in REFS:
        p = doc.add_paragraph(style="References")
        run = p.add_run(r)
        run.font.size = Pt(9)
        p.paragraph_format.space_after = Pt(2)

    doc.save(OUT)
    print(f"[cv] wrote {OUT}")


if __name__ == "__main__":
    build()
