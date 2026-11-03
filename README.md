# Multi-Brand Sales & Returns Reporting Automation

## Project Overview

This project develops a Python-based ETL and business reporting automation pipeline for multiple retail business units.

The goal is to reduce manual Excel processing, improve data quality, and support reliable sales and returns reporting.

**Current status:** Development — Version 0.1.0

## Objectives

- Automate sales and returns data ingestion.
- Clean and standardize raw transaction data.
- Validate missing values and transaction structures.
- Classify transactions using configurable business rules.
- Reconcile sales, returns, and quantity totals.
- Inspect Excel report templates and their formulas.
- Detect potential duplicate transactions.
- Develop a safe Excel reporting automation system.

## Technologies

- Python
- Pandas
- OpenPyXL
- Git and GitHub
- Microsoft Excel

## Project Structure

```text
multi-brand-reporting-automation/
├── config/
├── data/
│   ├── raw/
│   ├── processed/
│   └── templates/
├── src/
├── tests/
├── output/
├── docs/
├── .gitignore
└── README.md
```

## Current Development Progress

- [x] Python environment setup
- [x] Data ingestion
- [x] Data profiling
- [x] Data quality validation
- [x] Data cleaning
- [x] Business classification
- [x] Financial reconciliation
- [x] Excel template inspection
- [x] Transaction overlap analysis
- [ ] Approved reporting-scope configuration
- [ ] Dry-run update planner
- [ ] Automated Excel workbook updates
- [ ] End-to-end automated testing

## Data Privacy

This repository is intended to contain source code and documentation only.

Original business transactions, customer information, financial reports, and confidential company workbooks must not be committed.

## Future Development

1. Build a dry-run update planner.
2. Implement duplicate-safe Excel updates.
3. Add automated quality tests.
4. Validate final workbook outputs.
5. Add reporting dashboards.

## Author

GitHub: [nvtharindukothalawala-tech](https://github.com/nvtharindukothalawala-tech)

## Project Status

This is an ongoing learning and development project. The reporting automation is not yet approved for production use.
