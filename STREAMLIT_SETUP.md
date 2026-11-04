# Multi-Brand Reporting: local Streamlit interface

## 1. Install and launch

Place the three Python files in the paths shown in the ZIP. This adds a UI layer to your existing project; it **does not replace** current modules or change data files.

From your project root in VS Code PowerShell:

```powershell
.\.venv\Scripts\python.exe -m pip install "streamlit>=1.60,<2"
.\.venv\Scripts\python.exe -m pytest -q
.\.venv\Scripts\python.exe -m streamlit run streamlit_app.py --server.address 127.0.0.1
```

Streamlit generally opens a browser on `http://localhost:8501`. If not, manually open that address. Keep the application on localhost with confidential data.

## 2. How to use

1. Put the five original ChannelWise workbooks in `C:\Users\Tharindu\Desktop\Ass`, or supply your own private directory in the sidebar.
2. Select reporting month, e.g. `2026-10`.
3. Upload the **actual original** sales Excel export and returns Excel export for **that month**. Do not upload the per-business CSVs for this audit; it requires the original source headers.
4. Click **Validate monthly data**. It compares each source export against the real workbook, shows mapped/unmapped counts and discrepancies, and saves nothing into original workbooks.
5. Download the private audit JSON if needed.
6. **Only when the audit finds zero missing and zero conflicting mapped records**, you can confirm the checkbox and download **five byte-identical copies** of the existing reports.
7. If new transactions are missing (typical in a future month), the export button is disabled. An approved controlled update engine is required before generating updated reports.

## 3. Security and limitations

- Only use this app locally or on a company-approved secure machine. Do not deploy to a public Streamlit hosting platform with real customer/company data.
- Original input workbooks are never modified. Uploaded exports are temporarily stored for processing and deleted afterward.
- The UI is **not** a production update engine. It produces copies only when the real-data audit says existing workbook transactions already match the source.
- Real ChannelWise templates have known irregular `Paid By` / `Date` data column positions; the installed audit module handles the observed layout read-only.
- It does not recalculate Excel formulas, refresh pivots, or insert transactions.
- Unmapped sales/returns remain exceptions and need business owner approval.
- Never commit generated reports, uploaded sources or audit JSON to GitHub. Add `/output/` to `.gitignore` and keep original private files outside the repository.

## 4. Tests and commit

```powershell
.\.venv\Scripts\python.exe -m pytest -q

git add streamlit_app.py src/reporting_ui_service.py tests/test_reporting_ui_service.py STREAMLIT_SETUP.md
git diff --cached --check
git commit -m "Add local read-only Streamlit ChannelWise reporting dashboard"
git push origin main
```

If the existing baseline has 236 tests and no files changed, **247 tests** are expected (11 added). This is a target until verified in your own environment.

## 5. Nontechnical staff: double-click launcher

After Streamlit has been installed once, double-click `START_CHANNELWISE_DASHBOARD.bat` in the project folder. Use the browser interface, then close the terminal window to stop the local server.
