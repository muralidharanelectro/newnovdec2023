# Step-by-step protocol: `combine_results.py`

This protocol walks you through preparing the environment, validating the
inputs, executing the pipeline, and collecting the outputs produced by
`combine_results.py`.

## 1. Confirm prerequisites

1. Install **Python 3.9+** on your machine. The scripts rely on
   `pandas`, `numpy`, and `openpyxl`, which ship wheels for Python 3.9 and
   later on all major platforms.
2. Ensure that **Microsoft Excel workbooks** (`.xlsx`) can be read from the
   filesystem path where you intend to run the analysis.
3. Optional but recommended: install `git` so that you can track changes to the
   repository and configuration files.

## 2. Obtain the project files

1. Clone or download this repository to a convenient location.
2. Place the input workbooks alongside the scripts (repository root):
   * `new_result.xlsx` – consolidated or per-semester mark statements.
   * `CGPA.xlsx` – mapping of subject codes to credit values and nominal
     semesters.
   * `biodata.xlsx` – register number to gender/community mapping.
3. Verify that `config.json` exists in the repository root. Adjust the
   configuration only if your workbooks use different column headings or regex
   patterns.

## 3. Set up a virtual environment (recommended)

Creating an isolated environment avoids interfering with system-wide Python
packages.

```bash
python -m venv .venv
source .venv/bin/activate  # On Windows use: .venv\\Scripts\\activate
```

To exit the environment later, run `deactivate`.

## 4. Install dependencies

Install every required package using the provided requirements file.

```bash
pip install --upgrade pip
pip install -r requirements.txt
```

If you see build errors, confirm that you are using a modern Python (3.9+) and
have internet access to fetch wheels.

## 5. Validate the script entry point

Run the help command to ensure the script loads without syntax errors and to
inspect the available options.

```bash
python combine_results.py --help
```

This should print the CLI usage banner without raising exceptions.

## 6. Execute the analysis

Run the pipeline by pointing `--input` at the master results workbook. The
other files default to the names listed earlier; override them if your file
names differ.

```bash
python combine_results.py \
  --input "new_result.xlsx" \
  --outdir outputs \
  --subject-catalog "CGPA.xlsx" \
  --biodata "biodata.xlsx" \
  --current-semester 8
```

Key notes:

* `--outdir` is created automatically if it does not exist.
* `--current-semester` defaults to the value in `config.json` (8) when omitted.
  Supply a different number when analysing another term.
* To use an alternate config, supply `--config /path/to/config.json`.

The script prints progress and data-quality warnings to standard output. Review
those messages to catch missing sheets, unmatched subject codes, or malformed
rows.

## 7. Review generated artefacts

Upon success, the output directory contains:

* `master_results.xlsx` with multiple analysis sheets (`master_long`,
  `subject_outcomes`, `student_summary`, `student_overall`, etc.).
* CSV counterparts for each worksheet.
* `analytics_student_arrear_counts.csv`, `class_performance_summary.csv`, and
  `data_quality_report.csv` for downstream reporting.

Open `data_quality_report.csv` first to inspect any warnings flagged during the
run.

## 8. (Optional) Generate the class performance snapshot

If you want a standalone gender-wise classification table, run:

```bash
python generate_class_performance_summary.py --output class_performance_summary.xlsx
```

This script reads the CSV produced by the main pipeline and writes a summary in
Excel or CSV format depending on the chosen extension.

## 9. Troubleshooting tips

* **Missing columns or sheets** – Review `data_quality_report.csv` for detailed
  diagnostics. Update `config.json` to teach the parser about new header
  labels.
* **Unhandled subject codes** – Confirm that `CGPA.xlsx` contains every subject
  referenced in the semester workbooks. Add credit rows if necessary.
* **Biodata mismatches** – Ensure `biodata.xlsx` uses consistent register
  numbers. The script normalises whitespace but expects unique identifiers per
  row.
* **Re-running after fixes** – Delete the old `outputs` directory or choose a
  fresh `--outdir` to avoid mixing stale and updated artefacts.

## 10. Automate everything with a single command

Prefer not to run each step manually? Execute the bundled automation helper:

```bash
python automate_pipeline.py
```

The script orchestrates the entire checklist above: it verifies the required
Excel workbooks, provisions `.venv`, installs dependencies from
`requirements.txt`, runs `combine_results.py`, and finally invokes
`generate_class_performance_summary.py`. Use `python automate_pipeline.py --help`
to customise file names, change the output directory, skip the optional class
snapshot, or pass additional flags directly to the main pipeline.

Following these steps guarantees a clean run of `combine_results.py` and helps
identify data-quality issues quickly.
