# Finish-today delivery: Multi-brand sales reporting automation

## What this package adds

- `src/full_field_synthetic_writer.py`: saves **all 25 verified sales fields** with workbook-specific columns for five businesses; blocks unmarked inputs, existing output paths and conflicting transaction data; verifies output after saving.
- `src/full_field_synthetic_batch.py`: stages five workbooks, adds returns using the existing **six verified return columns**, publishes five outputs and a JSON financial summary only after staging succeeds, includes SHA-256 hashes.
- `scripts/run_final_synthetic_demo.py`: generates *fictional* source templates and transactions in a temporary folder, then produces a five-brand demonstration in `output/synthetic_final_demo/`.
- `tests/test_full_field_synthetic_delivery.py`: integration and safety tests.

**Compatibility:** Add these files to your existing repository. Do not replace any current source code or tests.

## Run on your Windows computer

Unzip into `C:\Users\Tharindu\Desktop\multi-brand-reporting-automation`, preserving folders. Open PowerShell in the project root.

```powershell
.\.venv\Scripts\python.exe -m pytest -q
.\.venv\Scripts\python.exe -m scripts.run_final_synthetic_demo
```

The first command should add **17 new tests** to the previously passing 212 tests, giving a **target of 229 passing tests**. Actual results depend on your local repo; do not commit if any test fails.

The second command generates five fictional Excel reports and `output/synthetic_final_demo/batch_summary.json`. It **refuses to overwrite existing demo outputs**. To rerun, manually inspect and remove only `output/synthetic_final_demo/` (never your original business workbooks).

## Final evidence check

```powershell
Get-Content output\synthetic_final_demo\batch_summary.json
.\.venv\Scripts\python.exe -m pytest -q
git status
```

Check that summary has exactly five businesses, 5 sales, 5 returns, and net amount 750.00. The output contains only example names and identifiers.

## Publish only safe code

```powershell
git add src/full_field_synthetic_writer.py src/full_field_synthetic_batch.py scripts/run_final_synthetic_demo.py tests/test_full_field_synthetic_delivery.py FINISH_TODAY_README.md
git diff --cached --check
git diff --cached --name-only
git commit -m "Complete synthetic five-brand full-field reporting demo"
git push origin main
git status
```

Never use `git add .` when local real company files may exist. Do **not** commit `output/`, `data/`, private workbooks, or real customer records. Check `.gitignore` carefully.

## Honest project status

**Demonstration complete (synthetic-only), pending local verification.** The engine validates 25 mapped sales fields and six return identity/amount fields across five Excel schemas. It detects full-field conflicts, validates dates, provides summaries and hashes, and safeguards source files. The original business templates are intentionally read-only.

**Not production-ready:** real export date formats, remaining return metadata, business-approval controls, Excel formula recalculation, true crash-atomic batch publishing and deployment permissions require independent approval and tests. The demo should be described to recruiters as a robust tested prototype, not a deployed live-company system.

## A final README summary for your repository

**Multi-Brand Retail Reporting Automation** is a Python-based prototype for five retail brands. It includes read-only source inspection, data-quality checks, business classification, financial reconciliation, Excel schema mapping, duplicate-aware synthetic Excel reporting, staging-based batch generation, saved-file verification, SHA-256 output hashes, and automated pytest coverage. Synthetic demonstration files illustrate the workflow without exposing real business data. Original customer exports and production Excel templates are excluded from writes and GitHub commits.
