# Result Analysis Toolkit

This project extracts, enriches, and analyses engineering result workbooks. It produces
machine-friendly outputs that cover current-semester performance, arrear tracking, and
community/gender based insights.

## Key Features

* **Automated sheet parsing** – Detects register numbers, subject columns, semester
  numbers, and arrear sheets using configurable patterns.
* **Flexible layout support** – Handles both consolidated sheets and per-student
  semester logs (semester/subject/grade columns).
* **Unified master dataset** – Consolidates every attended subject attempt into a single
  long-form table (`master_long`).
* **Biodata enrichment** – Dynamically merges gender and community information for each
  register number using `biodata.xlsx`.
* **Automated GPA calculation** – Maps grades to points via the credit catalogue,
  skipping uncleared subjects and arrears when computing current-semester GPAs.
* **Comprehensive analytics** – Computes subject-wise pass rates, per-semester student
  summaries, cumulative totals for current and backlog subjects, and “all clear” status.
* **Equity reporting** – Breaks down all-clear counts by gender and community to highlight
  cohort-level trends.
* **Data-quality feedback** – Captures common ingestion issues (missing columns, empty
  sheets, etc.) into `data_quality_report.csv`.

## Getting Started

### Fully automated run (recommended)

When the Excel workbooks and scripts live in the same directory, you can let
`automate_pipeline.py` perform every setup step for you:

```bash
python automate_pipeline.py
```

The helper script will verify that `new_result.xlsx` matches the supported
layout, create `.venv`, install `requirements.txt`, execute
`combine_results.py`, and generate the optional class performance snapshot.
Customise file names or skip the snapshot via the available CLI flags
(`python automate_pipeline.py --help`).

### Manual run

1. Install dependencies (preferably inside a virtual environment):

   ```bash
   pip install -r requirements.txt
   ```

2. Place the following input files in the repository root (default locations):

* `new_result.xlsx` – consolidated mark statements across semesters.
* `CGPA.xlsx` – subject-to-credit catalogue used for GPA computation.
* `biodata.xlsx` – register number to gender/community lookup.

3. Run the analysis:

   ```bash
   python combine_results.py --input "new_result.xlsx"
   ```

   Optional flags:

   * `--outdir` – directory to write outputs (defaults to `outputs`).
   * `--config` – alternate JSON config with column/regex overrides.
   * `--biodata` – path to the biodata workbook (defaults to `biodata.xlsx`).
   * `--subject-catalog` – credits/semester workbook for GPA mapping (defaults to `CGPA.xlsx`).
   * `--current-semester` – semester number considered the “current” term (defaults to 8).

   For a detailed, end-to-end checklist covering environment setup, validation,
   execution, troubleshooting, and automation options, see
   [`docs/combine_results_protocol.md`](docs/combine_results_protocol.md).

### One-click Windows launcher

Windows users can simply double-click `run_analysis.bat` once the input Excel files are in
place. The launcher will:

1. Ensure Python 3 and a local virtual environment are available.
2. Install/update the Python dependencies listed in `requirements.txt`.
3. Execute `combine_results.py` with:
   * `new_result.xlsx` as the input workbook.
   * `CGPA.xlsx` as the subject credit catalogue.
   * `biodata.xlsx` for community/gender enrichment.
   * `outputs` as the destination folder.
   * Semester 5 treated as the current semester by default.

If you need to analyse a different semester or point to renamed files, open the batch file
in a text editor and adjust the configuration block at the top.

4. Generate the class performance snapshot (gender-wise first/second class counts):

   ```bash
   python generate_class_performance_summary.py
   ```

   Use the `--output` flag to choose CSV or Excel output formats (e.g. `--output summary.xlsx`).

## Outputs

All outputs are written to the specified `--outdir`.

* `master_results.xlsx`
  * `master_long` – detailed subject-level attempts with arrear flags and biodata.
  * `subject_outcomes` – per-semester subject pass totals and percentages.
  * `student_summary` – per-student, per-semester totals with GPA/CGPA snapshots.
  * `student_overall` – cumulative attempts, passes, arrear counts, and all-clear flags.
  * `class_performance_summary` – gender-wise appeared/first-class/second-class/yet-to-pass counts.
  * `gender_community` – all-clear counts/pct by gender and community.
  * `semester_overview` – semester-level participation and pass-rate metrics.
  * `arrear_counts_by_semester` – per-student arrear subject counts split across semesters.
  * `all_clear_registers` – list of register numbers that are fully clear.
* CSV counterparts for each analytics sheet for easy integration with other tools.
* `analytics_student_arrear_counts.csv` – same arrear-by-semester table in CSV form.
* `class_performance_summary.csv` – gender-wise counts of appeared students and pass classifications (first/second class).
* `data_quality_report.csv` – ingestion warnings/errors that require manual review.

## Batch Usage

Windows users can continue to launch the pipeline via `run_analysis.bat`, which should be
updated to call the Python script with the desired arguments.

## Extending the Pipeline

* Update `config.json` if new subject header patterns, register formats, or arrear
  keywords are introduced.
* Integrate additional biodata attributes by extending `load_biodata` with new columns
  and normalisation rules.
* Use the generated CSVs as input for dashboards or BI tools to build student-wise,
  semester-wise, and cohort-wide reports.

