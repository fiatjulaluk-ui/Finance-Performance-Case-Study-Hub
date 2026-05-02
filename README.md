# Etex Australia Finance Showcase App

This turns the Excel management-accounting showcase into a Streamlit portfolio app.

## Why Streamlit

For job applications, Streamlit is the best first version because it puts the finance model, variance logic, and executive story in front of the reviewer quickly. It is easier to explain in an interview than Power BI, and much faster to build than a full React/Next.js dashboard.

## App Content

- Role-fit page mapping the app to the management accountant responsibilities
- Executive dashboard from the workbook KPI and strategic-context tabs
- Interactive rolling forecast simulator driven by revenue and margin assumptions
- April variance bridge and EBITDA driver explanation
- BGC acquisition synergy tracker
- FX risk and ESG-financial bridge
- QA, source trail, and reconciliation controls
- Workbook explorer for the full Excel file

## Run Locally

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run app.py
```

The app reads:

```text
data/Etex_Australia_MA_Showcase_v5_bc.xlsx
```

## Recommended Portfolio Use

Use this app as the interactive front door and keep the Excel workbook available as backup evidence. The app shows communication and decision support; the workbook proves modelling detail.

For the Etex role, position the project as a management reporting and operational finance toolkit: monthly reporting, variance commentary, driver-based forecasting, manufacturing cost awareness, balance sheet controls, source discipline, and continuous improvement.
