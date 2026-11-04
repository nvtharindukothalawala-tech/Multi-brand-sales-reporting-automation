# Real-data monthly reporting — safe operating runbook

## Scope and decision

This is an **actual October 2026 reconciliation** of the two supplied company exports against five supplied company reporting templates. It is **not** a permission to write to any company workbook. The program reads source workbooks with `read_only=True`, verifies input SHA-256 checksums before and after, and writes **only aggregate audit outputs** to a fresh private directory.

For October 2026, no new *mapped* sale or return needs insertion. The 6,177 source sale records and 293 source return records are already present, and the compared fields have zero differences after using the **observed data columns** (not two misleading late-stage sales header labels).

## Critical workbook-specific finding

All five real `SalesRepItem` worksheets have their row-4 **Paid By** and **Date** labels two columns before the transaction data values. The columns underneath the labels contain numeric fields rather than the displayed header values. This read-only script follows their actual data locations, avoiding a false anomaly. **Do not patch the original templates automatically; review the business meaning of those extra numeric columns before changing any workbook layout.**

## Real source files (private; outside GitHub)

Put these seven original input files in your private local folder, for example `C:\Users\Tharindu\Desktop\Ass`:

- `sales_reports_bill_wise_detail_xlsx_11_1_2026.xlsx`
- `return_reports_customer_bill_wise_2026-11-01_11_05.xlsx`
- `BeDapper_ChannelWise_Report.xlsx`
- `BeDesi_ChannelWise_Report.xlsx`
- `Giggles_ChannelWise_Report.xlsx`
- `Protecto_ChannelWise_Report.xlsx`
- `SafetySure_ChannelWise_Report.xlsx`

## Setup

Copy `src/real_business_dry_run.py` into your existing project's `src` folder. Copy `tests/test_real_business_dry_run.py` into the existing `tests` folder. The script uses the existing `openpyxl` dependency. **Do not copy any input workbooks into the GitHub repository.**

Run from your project's root in Windows PowerShell (adjust the private folder if different):

```powershell
.\.venv\Scripts\python.exe -m pytest -q

.\.venv\Scripts\python.exe -m src.real_business_dry_run `
  --sales "C:\Users\Tharindu\Desktop\Ass\sales_reports_bill_wise_detail_xlsx_11_1_2026.xlsx" `
  --returns "C:\Users\Tharindu\Desktop\Ass\return_reports_customer_bill_wise_2026-11-01_11_05.xlsx" `
  --templates-dir "C:\Users\Tharindu\Desktop\Ass" `
  --month "2026-10" `
  --output "C:\Users\Tharindu\Desktop\Ass\PRIVATE_audit_october_2026"
```

Expected with the supplied October inputs:

```text
REAL BUSINESS READ-ONLY AUDIT: NO_INSERTS_NEEDED_REVIEW_EXTRAS
Be Dapper: sales 644/644 (missing 0, extra 136, drift 0); returns 36/36 (missing 0, extra 3, drift 0)
Be Desi: sales 2456/2456 (missing 0, extra 215, drift 0); returns 62/62 (missing 0, extra 8, drift 0)
Giggles: sales 1036/1036 (missing 0, extra 0, drift 0); returns 47/47 (missing 0, extra 0, drift 0)
Protecto: sales 1175/1175 (missing 0, extra 0, drift 0); returns 65/65 (missing 0, extra 0, drift 0)
Safety Sure: sales 866/866 (missing 0, extra 0, drift 0); returns 83/83 (missing 0, extra 0, drift 0)
Unmapped sales: 353 | returns: 11
Original inputs SHA-256 unchanged: YES
```

Do **not** run the program a second time with the same output directory: it intentionally blocks overwriting. To rerun, select another new private directory name.

## Output (keep private)

- `real_data_audit.json` — per-business record matches, totals, aggregate location labels, field differences, SHA-256 fingerprints and audit status
- `business_summary.csv` — business-level counts, financial totals and discrepancies (no customer names)
- `field_differences.csv` — aggregate difference counts by field (empty apart from header on supplied October inputs)
- `READ_ME_FIRST.txt` — scope and confidentiality warning

These outputs are aggregate but still contain **potentially confidential company performance information**. Never commit them to a public GitHub repository or share them without approval.

## Interpretation and required approvals

- Source totals: 6,530 sales lines and 304 return lines.
- Mapped: 6,177 sales lines; 293 returns. The mapped sale amount is 28,238,844.06 and mapped return amount is 1,261,677.60 (source currency units).
- Unmapped: 353 sales lines (3,934,121.60) and 11 return lines (72,630.37). Preserve as exceptions, not silently discarded.
- Templates include 351 additional sales and 11 additional returns — **all other-location lines**, not missing mapped transactions. These belong to Sultan, Raja & Rani, Wedding Studio, LuxDesi and related warehouse streams. The 2 other unmapped sales lines are Mega Deals and were not found as extras in these five templates.
- All source files were SHA-256 checked unchanged before and after audit.
- This scope compares the 25 mapped sales fields and seven key return fields; remaining source return metadata and formula recalculation have not been production-validated.

**Operational rule:** Where zero mapped records are missing, do **not** insert duplicate transactions. Where records are missing, the audit status becomes `REVIEW_REQUIRED`; do not write until finance/business owners confirm the mapped scope, a backup strategy, a controlled destination copy, and a reversible publication procedure.

## Source-code publication

The new script uses real business and location names in configuration. Obtain employer permission before publishing any part of this private workflow. Do not commit input workbooks, actual transaction exports, or generated audit reports. The existing 229-test synthetic prototype is separately demo-ready.

## Current project boundary

**October reconciliation delivered and verified read-only; synthetic reporting prototype working; production workbook mutation intentionally not enabled.** This distinction is important for an accurate internship presentation.
