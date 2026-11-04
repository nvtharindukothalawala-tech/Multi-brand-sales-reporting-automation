# Multi-Brand ChannelWise Reporting Automation

**A Python-based, local reporting dashboard for validating monthly sales and returns across five retail businesses.**

The project combines data ingestion, business mapping, reconciliation, Excel workbook integrity checks, and an employee-friendly **Streamlit** interface. Its current real-data workflow is **read-only**: it audits existing ChannelWise workbooks and exports verified copies rather than inserting or changing transactions.

> **Project status:** Tested reporting prototype with real-data read-only reconciliation. Automated updates to real ChannelWise workbooks, Excel formula recalculation, and pivot-table refresh are **not implemented or production-approved**.

## Business problem

Monthly ERP exports contain sales and returns from multiple locations. Reporting staff need to identify the records relevant to each business and verify them against existing Excel reports. The workbooks use different field positions and may contain historical or other-location transactions, making manual copy/paste vulnerable to duplicates, misplaced values, and incorrect totals.

This project covers five businesses:

- **Be Desi**
- **Protecto**
- **Giggles**
- **Safety Sure**
- **Be Dapper**

## What the system does

| Capability | Current support |
| --- | --- |
| Upload monthly sales and returns XLSX files through Streamlit | Yes |
| Select the reporting month and local template directory | Yes |
| Validate input dates and compare records with five workbooks | Yes |
| Identify missing, conflicting, and unmapped transactions | Yes |
| Review extra transactions already present in workbooks | Yes |
| Download the audit results | Yes |
| Download a separate additional-transactions report | Yes |
| Download five verified copies when mapped records reconcile | Yes |
| Run synthetic full-field sales/returns demonstrations | Yes, in isolated test workflows |
| Insert genuinely new transactions into actual business workbooks | **No — not yet implemented** |
| Refresh Excel formulas, pivots, or charts | **No — not independently verified** |

**Verified copy** means the exported workbook is an unchanged copy of the corresponding original, checked using SHA-256. It does **not** mean the report was recalculated or updated with new transactions.

## Application workflow

```text
Monthly Sales XLSX + Monthly Returns XLSX
                  |
                  v
        Validate reporting month
                  |
                  v
     Assign records to five businesses
       + flag unrecognized locations
                  |
                  v
 Read-only comparison with ChannelWise
       SalesRepItem and Ord.Rtn
                  |
          +-------+-------+
          |               |
          v               v
    Audit results    Additional records
          |          review report (ZIP)
          v
  All mapped records match?
          |
     +----+----+
     |         |
    Yes        No
     |         |
     v         v
 Verified     Block report-copy export;
 copies ZIP   review differences
```

Unmapped transactions are not silently discarded or assigned to a business. Additional workbook records are not automatically deleted.

## Technology

- **Python** — validation, reporting services, batch logic, and audit tools
- **Streamlit** — local browser interface for nontechnical users
- **pandas / openpyxl** — spreadsheet ingestion and workbook inspection
- **pytest** — automated verification and regression tests
- **SHA-256** — integrity checking for source and copied files
- **Git / GitHub** — version control for source code and documentation only

## Quick start (Windows)

### 1. Get the code and create an environment

```powershell
git clone https://github.com/nvtharindukothalawala-tech/Multi-brand-sales-reporting-automation.git
cd Multi-brand-sales-reporting-automation
py -m venv .venv
.\.venv\Scripts\Activate.ps1
```

Install the project's existing Python dependencies, then ensure Streamlit is available:

```powershell
python -m pip install -r requirements.txt
python -m pip install streamlit
```

If your local checkout has no `requirements.txt`, install dependencies listed by the Python imports or add a reviewed dependency file before distributing the app. Never install untrusted packages just because they appear in an uploaded spreadsheet.

### 2. Prepare local business files

Keep the original five **ChannelWise XLSX** workbooks in a private directory outside the public Git repository. For each reporting period, have the ERP sales and returns exports available locally.

The application expects the established workbook layouts for the five businesses; it is not a universal Excel importer.

### 3. Start the dashboard

From the repository root:

```powershell
.\.venv\Scripts\python.exe -m streamlit run streamlit_app.py --server.address 127.0.0.1
```

Visit **http://127.0.0.1:8501** in your browser. If provided in your checkout, `START_CHANNELWISE_DASHBOARD.bat` is an alternative launcher after the environment has been configured.

### 4. Run a monthly check

1. Enter the reporting month in `YYYY-MM` format.
2. Select the folder containing the original ChannelWise templates.
3. Upload the **sales** and **returns** exports.
4. Select **Validate monthly data (read-only)**.
5. Review the per-business counts, missing records, field differences, warnings, and unmapped totals.
6. Download the **audit JSON** and, if needed, the **additional transactions report**.
7. If the application allows it and the results have been reviewed, prepare and download the **five verified report copies**.

**When source transactions are new or have changed**, the current dashboard intentionally blocks verified-copy export. It does not write those transactions into Excel.

## Project structure

The main components include:

```text
multi-brand-reporting-automation/
|-- streamlit_app.py                  # Local user interface
|-- src/
|   |-- reporting_ui_service.py      # UI-to-audit integration
|   |-- real_business_dry_run.py     # Read-only workbook reconciliation
|   |-- additional_transactions_report.py
|   |-- business_schema.py           # Five-business worksheet mappings
|   |-- full_field_source_mapping.py # 25-field sales source mapping
|   |-- schema_aware_synthetic_writer.py
|   |-- synthetic_monthly_batch.py
|   `-- ...
|-- scripts/
|   |-- export_verified_channelwise_copies.py
|   `-- run_final_synthetic_demo.py
|-- tests/                         # Automated unit/integration tests
|-- config/                        # Business mapping/configuration
|-- docs/                          # Technical documentation
|-- output/                        # PRIVATE; generated files (Git-ignored)
`-- README.md
```

The exact modules in your checkout may include additional audit, financial-validation, publication, and recovery utilities.

## Run the tests

```powershell
.\.venv\Scripts\python.exe -m pytest -q
```

The repository has a substantial automated test suite, including synthetic five-business Excel tests, reconciliation safeguards, real-data read-only audits, and UI-service tests. Run the full suite after installing or modifying the dashboard; do not assume an old pass result covers new changes.

Important safety checks include rejecting unmarked workbooks from synthetic writers, preserving duplicate occurrence counts, refusing unsafe overwrites, and detecting changes to verified outputs.

## Results and limitations

In the October 2026 validation exercise, all **6,177 mapped sales rows** and **293 mapped return rows** were found in the corresponding business workbooks. No mapped insertions were needed for that dataset. The system also flagged unmapped source records and extra workbook rows for review. These are **period-specific observations**, not guarantees for future months.

The current design intentionally separates:

1. **Read-only real-data processing:** audits, comparisons, exception review, and verified workbook copies.
2. **Synthetic writing tests:** simulated insertion, mapping, idempotency, and financial checks on independently generated, explicitly marked test workbooks.

**Production write-back is a future milestone.** It requires an approved business mapping, immutable input backups, conflict approval, safe output-only writing, field-by-field reconciliation, Excel recalculation or refresh, and post-publication audit checks.

## Security and data handling

This is a **local-only** dashboard intended for authorized company use. Do not deploy it on a public Streamlit service with real company data.

- Never commit real XLSX/CSV exports, generated reports, private audit JSON, customer details, or credentials.
- Keep company workbooks and downloaded reports in approved private locations.
- Confirm that `.gitignore` excludes at least `.venv/`, `output/`, `data/raw/`, private template directories, `*.xlsx`, `*.xlsm`, `*.csv`, and local secrets. Review tracked files separately: `.gitignore` does not untrack existing commits.
- Never place the `SYNTHETIC_TEST_ONLY` marker in a real business workbook or a copy containing real data.
- Use independent backups and explicit authorization before implementing or executing future real-workbook write operations.

## Roadmap

- [x] Business-level sales/returns mapping and reconciliation
- [x] Five-business workbook schema validation
- [x] Synthetic writer and monthly batch tests
- [x] Read-only real-data comparison and audit
- [x] Local Streamlit upload and review interface
- [x] Verified workbook-copy and additional-transaction downloads
- [ ] Approved-copy real transaction update engine
- [ ] Excel calculated-summary and pivot-refresh verification
- [ ] Nontechnical-user acceptance testing and controlled rollout

## Why this project matters

It demonstrates practical data engineering and quality-assurance skills: understanding inconsistent enterprise spreadsheets, preserving financial integrity, building repeatable validations, separating test and production data, and turning a Python workflow into a usable internal application.

---

**Confidentiality note:** This repository is intended to contain source code and documentation only. Real customer, sales, return, and workbook data must remain private.
