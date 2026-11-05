# Supervisor Setup — Multi-Brand ChannelWise Reporting Automation

**GitHub:** https://github.com/nvtharindukothalawala-tech/Multi-brand-sales-reporting-automation

This guide explains how to review the **local, read-only** Streamlit reporting dashboard on a Windows laptop. The system validates monthly sales and returns, shows exceptions, and provides verified copies of existing ChannelWise workbooks. **It does not insert new transactions into real Excel workbooks.**

## 1. Download the project

Open the GitHub link above and choose **Code → Download ZIP**, then extract the ZIP. Alternatively, if Git is installed:

```powershell
git clone https://github.com/nvtharindukothalawala-tech/Multi-brand-sales-reporting-automation.git
cd Multi-brand-sales-reporting-automation
```

If you downloaded the ZIP, open PowerShell **inside the extracted project folder** instead of running `cd`.

## 2. Install Python and dependencies

Install **Python 3.11 or newer** on Windows, then run from the project folder:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
```

If the repository includes `requirements.txt`, use:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

If there is no requirements file, install the main dependencies:

```powershell
.\.venv\Scripts\python.exe -m pip install streamlit pandas openpyxl pytest
```

Internet access is needed during installation. Installation issues involving extra dependencies should be resolved using the repository's latest documentation.

## 3. Prepare private Excel input files

**These files are not in public GitHub and must be obtained through an approved company channel:**

- `BeDesi_ChannelWise_Report.xlsx`
- `Protecto_ChannelWise_Report.xlsx`
- `Giggles_ChannelWise_Report.xlsx`
- `SafetySure_ChannelWise_Report.xlsx`
- `BeDapper_ChannelWise_Report.xlsx`
- The **original monthly sales XLSX export** and **original monthly returns XLSX export**.

Store the five templates in a private directory, for example `C:\ChannelWise_Private_Templates`.

## 4. Start the application

Double-click `START_CHANNELWISE_DASHBOARD.bat`, or run:

```powershell
.\.venv\Scripts\python.exe -m streamlit run streamlit_app.py --server.address 127.0.0.1
```

Open **http://127.0.0.1:8501** in your browser. Leave the terminal open while the app runs.

## 5. Evaluate one monthly reporting cycle

1. Choose the month in `YYYY-MM` format.
2. Enter the local directory holding the **five approved baseline workbooks**.
3. Upload the original **sales** and **returns** Excel exports.
4. Select **Validate monthly data (read-only)**.
5. Review mapped records, missing records, conflicts, unmapped counts and warnings.
6. Download the private audit JSON and **Additional Transactions Report** if needed.
7. If the application permits verified copying, confirm the disclosure and download the **five ChannelWise report copies (ZIP)**.

**Important:** If new or changed mapped transactions are present, the app blocks verified-copy export. Automated updates to real company workbooks, formula recalculation and pivot refresh are **not** implemented.

## 6. Run automated tests (optional)

```powershell
.\.venv\Scripts\python.exe -m pytest -q
```

The developer's most recently reported local baseline was **251 passing tests**. The number may change as code evolves; the supervisor should evaluate the result on their own machine.

## 7. Data security

Keep source data, templates, report ZIPs, and audit details **outside the public repository**. Do not upload customer information or financial workbook copies to GitHub. Keep Streamlit bound to `127.0.0.1` when handling private data. Stop the local application with **Ctrl+C**.

---

**Project status:** Working local reporting, reconciliation and verified-copy prototype; not a production-approved real-workbook update engine.
