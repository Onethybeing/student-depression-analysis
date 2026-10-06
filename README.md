# Determinants of Depression among University Students

An econometric analysis of what is associated with depression among 300 university students
(aged 18–24), using **binary logit and probit models with marginal effects**.

**Team:** Atharv Lodhi, Sourav Choudhary, Piyush Bisht. Department of Humanities and Social Sciences, IIT Kharagpur.

| Deliverable | File |
|---|---|
| Research paper, IEEE two-column (PDF / Word) | [PDF](report/Student_Depression_Research_Paper.pdf) · [DOCX](report/Student_Depression_Research_Paper.docx) |
| Full report, single column (PDF / Word) | [PDF](report/Student_Depression_Report.pdf) · [DOCX](report/Student_Depression_Report.docx) |
| Presentation (10 slides, native charts) | [report/Student_Depression_Presentation.pptx](report/Student_Depression_Presentation.pptx) |
| Google Form (survey instrument) | [forms/](forms/) (Apps Script that creates the form and its response sheet) |
| Python notebook (executed, with all outputs) | [notebooks/student_depression_analysis.ipynb](notebooks/student_depression_analysis.ipynb) |
| Stata do-file | [stata/student_depression_analysis.do](stata/student_depression_analysis.do) |
| Excel workbook with every result table and chart | [results/Student_Depression_Analysis_Results.xlsx](results/Student_Depression_Analysis_Results.xlsx) |
| Data (300 students) | [data/student_sample_300.csv](data/student_sample_300.csv) · [.xlsx](data/student_sample_300.xlsx) |

## Key findings

Average marginal effects from the logit model (change in the probability of depression):

| Factor | Effect | p-value |
|---|---:|---:|
| Single (vs in a relationship) | **+36.4 pp** | < 0.001 |
| Rich family (vs middle class) | **−17.2 pp** | 0.011 |
| Ever married (vs unmarried) | **−16.4 pp** | 0.003 |
| Male (vs female) | **−10.7 pp** | 0.050 |
| Monthly living cost (per 1,000) | **+1.15 pp** | 0.045 |
| CGPA, age, semester, friend circle, living with family | not significant | > 0.10 |

* 89 of 300 students (29.7%) report depression.
* Logit model: LR χ²(11) = 69.57 (p < 0.001), pseudo-R² = 0.191, AUC = 0.781, Hosmer–Lemeshow p = 0.314.
* Probit and linear probability models give almost identical marginal effects.
* Self-reported health condition almost perfectly predicts depression (100% of "critical", 0% of "good"),
  so it is analysed descriptively and only added in a robustness model.

## Methods

1. Data checks and coding (dummy variables; very small categories merged to avoid perfect prediction)
2. Descriptive statistics and charts
3. Chi-square tests (with Cramér's V), Welch t-tests and Mann–Whitney tests
4. Multicollinearity check (VIF)
5. Binary logit: coefficients, odds ratios, average marginal effects and marginal effects at the means
6. Binary probit with marginal effects
7. Fit diagnostics: LR test, pseudo-R², AIC/BIC, classification tables, ROC/AUC, Hosmer–Lemeshow
8. Robustness: robust standard errors, linear probability model, model with health status, male/female subsamples

## Repository layout

```
data/                 analysis sample (300 students)
notebooks/            executed analysis notebook (source of every number)
stata/                Stata do-file replicating the notebook
results/
  figures/            11 charts used in the report
  tables/             every result table as CSV
  model_output/       full model summaries (used in the report appendix)
  results.json        key scalar results used in the report text
  Student_Depression_Analysis_Results.xlsx
report/               research paper, full report (DOCX + PDF) and presentation (PPTX)
forms/                Google Apps Script that creates the survey form
assets/               IIT Kharagpur logo used on the cover pages
scripts/
  build_paper.py      IEEE-style two-column paper from results/ (no hand-typed numbers)
  build_report.py     single-column full report from results/
  export_pdf.py       refreshes contents pages in MS Word and exports both PDFs
  export_slide_data.py + build_slides.js   10-slide deck with native charts
```

## Reproduce

```bash
pip install -r requirements.txt
jupyter nbconvert --to notebook --execute --inplace notebooks/student_depression_analysis.ipynb
python scripts/build_paper.py
python scripts/build_report.py
python scripts/export_pdf.py          # Windows + Microsoft Word
python scripts/export_slide_data.py
npm install pptxgenjs && node scripts/build_slides.js
```

Stata (run from the repository root): `do stata/student_depression_analysis.do`.
The do-file mirrors the notebook step by step; the numbers in the report come from the Python run.
