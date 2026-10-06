# CONTEXT — cross-agent working notes

Read this first. It records the goal, hard rules, decisions and current status so any agent
(Claude, Codex, Gemini, Cursor) can continue the work without the original conversation.

## Goal
Course research project (econometrics, same format as earlier team reports): analyse what is associated
with depression among university students and deliver
1. an executed Python notebook (source of every number),
2. a Stata do-file replicating it,
3. an Excel workbook with all results,
4. a research report (DOCX + PDF) in the course's report format (title page, abstract, TOC,
   5 chapters, appendix, references) with charts and tables,
5. all of it in a public GitHub repo.

## Hard rules
- **Never mention the city/region or the country** of the data anywhere in the repo, report, notebook,
  Excel, commit messages or PR text. The region column is dropped before the data enter the repo.
- The full source dataset (`dataset/`) and the sampling script (`private/make_sample.py`) are git-ignored
  and stay on the owner's machine only. Only `data/student_sample_300.*` is published.
- `examples/` holds other teams' reports used only as a format reference — git-ignored, never publish.
- Report numbers must come from `results/` (built by `scripts/build_report.py`), never typed by hand.
- Do not claim the data were collected by the team through a Google Form; describe it as survey data.
- Commits/PRs: no AI attribution lines. Work on a branch + PR; do not merge without the owner's approval.

## Data
- `data/student_sample_300.csv`: 300 students, ages 18–24, one home region, random sample (seed 42),
  region column removed, `Student_ID` added. 89 depressed (29.7%), 144 male / 156 female.
- Columns: Gender, Age, Family_background, Relationship_status, Marital_status, Living_style,
  Current_semester, CGPA, friend_circle, Health_condition, Cost_without_study, Depression.

## Decisions (with reasons)
- Main model = binary logit; probit and LPM as checks (outcome is Yes/No; the format reference used
  ordered logit/probit for an ordinal outcome).
- `Health_condition` excluded from the main model: it nearly determines the outcome (Critical 100% depressed,
  Good 0%) → quasi-complete separation, and it likely measures the same condition. Shown descriptively and in
  robustness model (5) only.
- Relationship "Complicated" (n=11, 0 depressed) merged into not-single → `single` dummy.
  Marital "Separated" (n=6) merged with Married → `ever_married`.
- Male/female subsample models are LPMs, not logits: among men, 0 of 29 non-single students are depressed,
  so `single` perfectly predicts the outcome in the male subsample.
- Bases: middle class, not single, unmarried, bachelor living, no friend circle, female.

## Measured results (logit, N = 300)
- LR χ²(11) = 69.57, p < 0.001; pseudo-R² 0.191; AUC 0.781; Hosmer–Lemeshow p = 0.314; max VIF 1.61.
- AMEs: single +36.4 pp (p<0.001); rich −17.2 pp (p=0.011); ever married −16.4 pp (p=0.003);
  male −10.7 pp (p=0.050); monthly cost +1.15 pp per 1,000 (p=0.045). CGPA, age, semester,
  friend circle, living with family not significant.

## Pipeline
1. (private) `python private/make_sample.py` → `data/student_sample_300.*`
2. `jupyter nbconvert --to notebook --execute --inplace notebooks/student_depression_analysis.ipynb`
   → `results/` (figures, tables, model_output, results.json, Excel)
3. `python scripts/build_report.py` → `report/Student_Depression_Report.docx`
4. `python scripts/export_pdf.py` (Windows + Word) → refreshes TOC, writes the PDF

## Status (2026-10-06)
- All deliverables built and checked. Stata is not installed on the build machine, so the do-file has not
  been executed; it mirrors the notebook command by command.
- Title page has placeholders for team member names and roll numbers — the owner will fill them in.
- Data source/provenance link not yet provided by the owner.

## Maintaining this file
Update this file in the same commit as any change to data, method, results, structure or decisions.
Keep it short and factual; record only verified numbers.
