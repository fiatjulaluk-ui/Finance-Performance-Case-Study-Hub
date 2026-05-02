
from __future__ import annotations

from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st


APP_DIR = Path(__file__).resolve().parent
WORKBOOK_PATH = APP_DIR / "data" / "Etex_Australia_MA_Showcase_v5_bc.xlsx"

BRAND_ORANGE = "#E8500A"
BRAND_CHARCOAL = "#1A1A1A"
MUTED_GREY = "#6B7A8D"


st.set_page_config(
    page_title="Etex Australia Finance Performance Case Study Hub",
    page_icon=":bar_chart:",
    layout="wide",
    initial_sidebar_state="expanded",
)


st.markdown(
    f"""
    <style>
    .stApp {{
        background: #f7f7f5;
        color: {BRAND_CHARCOAL};
    }}
    [data-testid="stSidebar"] {{
        background: #ffffff;
        border-right: 1px solid #e7e2de;
    }}
    h1, h2, h3 {{
        letter-spacing: 0 !important;
    }}
    div[data-testid="stMetric"] {{
        background: #ffffff;
        border: 1px solid #e7e2de;
        border-left: 4px solid {BRAND_ORANGE};
        border-radius: 8px;
        padding: 14px 16px;
        min-height: 118px;
    }}
    .small-note {{
        color: {MUTED_GREY};
        font-size: 0.9rem;
    }}
    .section-label {{
        color: {BRAND_ORANGE};
        font-weight: 700;
        text-transform: uppercase;
        font-size: 0.78rem;
        letter-spacing: 0.04em;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner=False)
def load_workbook(path: Path) -> dict[str, pd.DataFrame]:
    if not path.exists():
        st.error(f"Workbook not found: {path}")
        st.stop()

    xls = pd.ExcelFile(path)
    return {
        sheet: pd.read_excel(path, sheet_name=sheet, header=None, engine="openpyxl")
        for sheet in xls.sheet_names
    }


def clean_table(df: pd.DataFrame) -> pd.DataFrame:
    table = df.copy()
    table = table.dropna(how="all").dropna(axis=1, how="all")
    table.columns = [str(c).replace("\n", " ").strip() for c in table.columns]
    return table.reset_index(drop=True)


def slice_table(
    sheets: dict[str, pd.DataFrame],
    sheet_name: str,
    start_row: int,
    end_row: int,
    start_col: int,
    end_col: int,
    header: bool = True,
) -> pd.DataFrame:
    raw = sheets[sheet_name].iloc[start_row - 1 : end_row, start_col - 1 : end_col]
    raw = raw.dropna(how="all").dropna(axis=1, how="all")
    if raw.empty:
        return pd.DataFrame()
    if header:
        columns = raw.iloc[0].fillna("").astype(str).str.replace("\n", " ", regex=False)
        raw = raw.iloc[1:].copy()
        raw.columns = columns
    return clean_table(raw)


def as_number(value, default: float = 0.0) -> float:
    try:
        if pd.isna(value):
            return default
        return float(value)
    except (TypeError, ValueError):
        return default


def money(value: float) -> str:
    return f"${value:,.0f}k"


def pct(value: float) -> str:
    return f"{value:.1%}"


def parse_metric(text: str) -> str:
    return str(text).split(":", 1)[-1].strip() if ":" in str(text) else str(text)


sheets = load_workbook(WORKBOOK_PATH)

st.sidebar.title("Case Study Hub")
st.sidebar.caption("Finance performance app built from the Excel model.")
view = st.sidebar.radio(
    "Navigate",
    [
        "Role Fit",
        "Executive Dashboard",
        "Forecast Simulator",
        "Variance Bridge",
        "BGC Synergies",
        "Risk & ESG",
        "Controls & Audit Trail",
        "Workbook Explorer",
    ],
)

st.title("Etex Australia Finance Performance Case Study Hub")
st.caption(
    "April 2025 | AUD $000s unless stated | Built from the Excel workbook tabs, formulas, and source trail."
)

dashboard = sheets["Dashboard"]
drivers = sheets["Drivers"]
rolling = sheets["Rolling Forecast"]

if view == "Role Fit":
    st.markdown('<div class="section-label">Job application positioning</div>', unsafe_allow_html=True)
    st.subheader("Best App Positioning for This Management Accountant Role")
    st.write(
        "For this role, the app should be presented as a management reporting and business performance toolkit. "
        "The strongest message is that you can coordinate inputs, protect reporting integrity, explain performance, "
        "lead forecast cycles, and make finance easier for operations teams to act on."
    )

    role_map = pd.DataFrame(
        [
            {
                "Role requirement": "Monthly sales reporting and sales vs budget analysis",
                "App evidence": "Executive Dashboard, Monthly P&L, Variance Bridge",
                "Interview message": "I can turn monthly actuals into timely commentary that management can use.",
            },
            {
                "Role requirement": "Product costings, BOM reviews, inventory revaluations",
                "App evidence": "Product Costing, Inventory, Workbook Explorer",
                "Interview message": "I understand manufacturing finance drivers beyond the P&L summary.",
            },
            {
                "Role requirement": "True and fair view of business performance",
                "App evidence": "Variance Bridge and strategic context panels",
                "Interview message": "I separate volume, rate, mix, cost control, and integration impacts.",
            },
            {
                "Role requirement": "Forecasting and budgeting leadership",
                "App evidence": "Forecast Simulator built from Drivers and Rolling Forecast",
                "Interview message": "I can gather assumptions and translate them into a live forecast quickly.",
            },
            {
                "Role requirement": "Internal controls, reconciliations, audit support",
                "App evidence": "Controls & Audit Trail, BS Recs, Sources, QA Governance",
                "Interview message": "I design reporting with traceability, checks, and review evidence.",
            },
            {
                "Role requirement": "Continuous improvement and finance process training",
                "App evidence": "Interactive sliders, workbook explorer, simplified management views",
                "Interview message": "I can make finance processes easier for manufacturing and state teams to use.",
            },
            {
                "Role requirement": "Compliance culture and statutory/audit/GDPR diligence",
                "App evidence": "Source trail, QA checklist, reconciliation controls",
                "Interview message": "I treat accuracy, documentation, and governance as part of the deliverable.",
            },
        ]
    )
    st.dataframe(role_map, width="stretch", hide_index=True)

    st.subheader("Recommended Application Version")
    c1, c2, c3 = st.columns(3)
    c1.metric("Best first build", "Streamlit", "Fast, finance-led, interview friendly")
    c2.metric("Best later polish", "React/Next.js", "Use if you need a public portfolio site")
    c3.metric("Best corporate BI", "Power BI", "Use if the employer asks for BI reporting")

    st.info(
        "My recommendation: use Streamlit for this job application now. It lets the reviewer interact with the "
        "forecast and variance logic immediately, while the original Excel file remains the detailed audit backup."
    )

    st.subheader("How to Talk About It")
    st.markdown(
        """
        - "I built this from an Excel management reporting pack to show how I would explain Australian operations performance."
        - "The app connects driver assumptions to forecast outcomes, so managers can test scenarios during forecast reviews."
        - "I included source, QA, and reconciliation views because reporting integrity matters as much as the dashboard."
        - "The BGC synergy and ESG-financial pages show that I can connect operational initiatives to financial results."
        """
    )

elif view == "Executive Dashboard":
    st.markdown('<div class="section-label">YTD performance</div>', unsafe_allow_html=True)
    metric_cols = st.columns(5)
    metric_positions = [0, 3, 6, 9, 12]
    for col, pos in zip(metric_cols, metric_positions):
        label = dashboard.iat[3, pos]
        actual = parse_metric(dashboard.iat[5, pos])
        budget = parse_metric(dashboard.iat[6, pos])
        achievement = parse_metric(dashboard.iat[7, pos])
        col.metric(str(label), actual, f"Budget {budget}")
        col.caption(f"Achievement: {achievement}")

    st.divider()
    left, right = st.columns([1.35, 1])

    with left:
        st.subheader("Group 5-Year Performance")
        group = slice_table(sheets, "Dashboard", 11, 18, 1, 7)
        numeric_years = ["2021", "2022", "2023", "2024", "2025"]
        chart = group[["Metric", *numeric_years]].set_index("Metric").T
        st.line_chart(chart, height=320)
        st.dataframe(group, width="stretch", hide_index=True)

    with right:
        st.subheader("Australia Strategic Context")
        context = dashboard.iloc[10:19, [9, 11]].dropna(how="all")
        context.columns = ["Topic", "Management message"]
        st.dataframe(clean_table(context), width="stretch", hide_index=True)
        st.info(
            "The strongest job-application story is not just reporting: it shows variance analysis, BGC integration, driver-based forecasting, controls, and source discipline."
        )

elif view == "Forecast Simulator":
    st.markdown('<div class="section-label">Driver-based rolling forecast</div>', unsafe_allow_html=True)
    st.subheader("Live Forecast Simulator")
    st.write(
        "Adjust the commercial drivers and the May-Dec forecast recalculates immediately. "
        "This translates the workbook's Drivers and Rolling Forecast tabs into an interview-ready app workflow."
    )

    months = ["May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    base_revenue = pd.Series(drivers.iloc[5, 1:9].astype(float).values, index=months)
    budget_revenue = pd.Series(drivers.iloc[6, 1:9].astype(float).values, index=months)
    capex = pd.Series(drivers.iloc[24, 1:9].astype(float).values, index=months)

    base_rates = {
        "Materials": as_number(drivers.iat[12, 1]),
        "Labour": as_number(drivers.iat[13, 1]),
        "Manufacturing OH": as_number(drivers.iat[14, 1]),
        "SG&A": as_number(drivers.iat[17, 1]),
    }

    controls = st.columns(5)
    revenue_uplift = controls[0].slider("Revenue uplift", -10.0, 15.0, 0.0, 0.5) / 100
    materials_rate = controls[1].slider("Materials %", 45.0, 62.0, base_rates["Materials"] * 100, 0.1) / 100
    labour_rate = controls[2].slider("Labour %", 10.0, 20.0, base_rates["Labour"] * 100, 0.1) / 100
    overhead_rate = controls[3].slider("Mfg OH %", 8.0, 16.0, base_rates["Manufacturing OH"] * 100, 0.1) / 100
    sga_rate = controls[4].slider("SG&A %", 8.0, 14.0, base_rates["SG&A"] * 100, 0.1) / 100

    revenue = base_revenue * (1 + revenue_uplift)
    materials = revenue * materials_rate
    labour = revenue * labour_rate
    overhead = revenue * overhead_rate
    gross_profit = revenue - materials - labour - overhead
    sga = revenue * sga_rate
    ebitda = gross_profit - sga

    scenario = pd.DataFrame(
        {
            "Revenue": revenue,
            "Budget Revenue": budget_revenue,
            "Gross Profit": gross_profit,
            "EBITDA": ebitda,
            "CapEx": capex,
            "EBITDA Margin": ebitda / revenue,
        }
    )

    kpis = st.columns(4)
    kpis[0].metric("H2 revenue", money(scenario["Revenue"].sum()), f"{money((revenue - budget_revenue).sum())} vs budget")
    kpis[1].metric("H2 EBITDA", money(scenario["EBITDA"].sum()), pct(scenario["EBITDA"].sum() / scenario["Revenue"].sum()))
    kpis[2].metric("Gross margin", pct(scenario["Gross Profit"].sum() / scenario["Revenue"].sum()))
    kpis[3].metric("CapEx plan", money(scenario["CapEx"].sum()))

    chart_cols = st.columns(2)
    chart_cols[0].subheader("Revenue vs Budget")
    chart_cols[0].bar_chart(scenario[["Revenue", "Budget Revenue"]], height=300)
    chart_cols[1].subheader("EBITDA Margin")
    chart_cols[1].line_chart(scenario[["EBITDA Margin"]], height=300)

    display = scenario.copy()
    for col in ["Revenue", "Budget Revenue", "Gross Profit", "EBITDA", "CapEx"]:
        display[col] = display[col].map(lambda x: round(x, 1))
    display["EBITDA Margin"] = display["EBITDA Margin"].map(lambda x: f"{x:.1%}")
    st.dataframe(display, width="stretch")

elif view == "Variance Bridge":
    st.markdown('<div class="section-label">Management explanation</div>', unsafe_allow_html=True)
    summary = slice_table(sheets, "Variance Bridge", 5, 14, 1, 9)
    bridge = slice_table(sheets, "Variance Bridge", 17, 23, 1, 8)

    st.subheader("April Actual vs Budget")
    st.dataframe(summary, width="stretch", hide_index=True)

    st.subheader("EBITDA Waterfall Drivers")
    bridge_chart = bridge[["Bridge Step", "Impact $000s"]].copy()
    bridge_chart["Impact $000s"] = pd.to_numeric(bridge_chart["Impact $000s"], errors="coerce").fillna(0)
    st.bar_chart(bridge_chart.set_index("Bridge Step"), height=320)
    st.dataframe(bridge, width="stretch", hide_index=True)

elif view == "BGC Synergies":
    st.markdown('<div class="section-label">Acquisition integration</div>', unsafe_allow_html=True)
    synergies = slice_table(sheets, "BGC Synergy Tracker", 5, 13, 1, 11)
    st.subheader("BGC Synergy Delivery")

    total_target = pd.to_numeric(synergies["FY25 Target AUD $000s"], errors="coerce").sum()
    total_actual = pd.to_numeric(synergies["FY25 Actual AUD $000s"], errors="coerce").sum()
    on_track = int((synergies["Status"] == "ON TRACK").sum())

    c1, c2, c3 = st.columns(3)
    c1.metric("FY25 actual synergies", money(total_actual), f"{pct(total_actual / total_target)} delivered")
    c2.metric("FY26 full-run target", money(pd.to_numeric(synergies["FY26 Full-Run AUD $000s"], errors="coerce").sum()))
    c3.metric("On-track categories", f"{on_track} of {len(synergies)}")

    chart = synergies[
        [
            "Synergy Category",
            "FY25 Target AUD $000s",
            "FY25 Actual AUD $000s",
            "FY26 Full-Run AUD $000s",
        ]
    ].copy()
    chart = chart[~chart["Synergy Category"].astype(str).str.contains("TOTAL", na=False)]
    chart = chart.rename(
        columns={
            "FY25 Target AUD $000s": "FY25 Target",
            "FY25 Actual AUD $000s": "FY25 Actual",
            "FY26 Full-Run AUD $000s": "FY26 Full-Run",
        }
    )
    chart_long = chart.melt("Synergy Category", var_name="Measure", value_name="AUD $000s")
    chart_long["AUD $000s"] = pd.to_numeric(chart_long["AUD $000s"], errors="coerce")
    synergy_chart = (
        alt.Chart(chart_long)
        .mark_bar()
        .encode(
            y=alt.Y("Synergy Category:N", sort="-x", title=None),
            x=alt.X("AUD $000s:Q", title="AUD $000s"),
            color=alt.Color(
                "Measure:N",
                scale=alt.Scale(range=[BRAND_ORANGE, "#1F77B4", "#9E9E9E"]),
                legend=alt.Legend(orient="bottom"),
            ),
            tooltip=["Synergy Category", "Measure", alt.Tooltip("AUD $000s:Q", format=",.0f")],
        )
        .properties(height=340)
    )
    st.altair_chart(synergy_chart, width="stretch")
    st.dataframe(synergies, width="stretch", hide_index=True)

elif view == "Risk & ESG":
    st.markdown('<div class="section-label">Risk, sustainability, financial linkage</div>', unsafe_allow_html=True)
    risk_tab, esg_tab = st.tabs(["FX and rates risk", "ESG-financial bridge"])

    with risk_tab:
        risk = slice_table(sheets, "Risk & Sensitivity", 5, 11, 1, 10)
        st.subheader("Currency Sensitivity")
        st.dataframe(risk, width="stretch", hide_index=True)
        aud_row = risk[risk["Currency"].astype(str).str.contains("Australian", na=False)]
        if not aud_row.empty:
            st.info(str(aud_row.iloc[0]["Significance to Altona"]))

    with esg_tab:
        esg = slice_table(sheets, "ESG-Financial Bridge", 5, 13, 1, 10)
        st.subheader("Sustainability KPI to P&L Linkage")
        st.dataframe(esg, width="stretch", hide_index=True)
        progress = esg["Progress to Target"].value_counts()
        st.bar_chart(progress, height=260)

elif view == "Controls & Audit Trail":
    st.markdown('<div class="section-label">Governance proof</div>', unsafe_allow_html=True)
    qa = slice_table(sheets, "QA Governance", 5, 16, 1, 7)
    sources = slice_table(sheets, "Sources", 3, 20, 1, 10)
    bs_recs = slice_table(sheets, "BS Recs", 4, 18, 1, 10)

    st.subheader("Model QA Controls")
    pass_count = int((qa["Result"] == "PASS").sum()) if "Result" in qa else 0
    st.metric("QA checks passed", f"{pass_count} of {len(qa)}")
    st.dataframe(qa, width="stretch", hide_index=True)

    st.subheader("Source Trail")
    st.dataframe(sources, width="stretch", hide_index=True)

    st.subheader("Balance Sheet Reconciliation Snapshot")
    st.dataframe(bs_recs, width="stretch", hide_index=True)

else:
    st.markdown('<div class="section-label">Workbook content</div>', unsafe_allow_html=True)
    sheet_name = st.selectbox("Sheet", list(sheets.keys()))
    preview = clean_table(sheets[sheet_name].head(80))
    st.dataframe(preview, width="stretch")
    st.download_button(
        "Download source workbook",
        WORKBOOK_PATH.read_bytes(),
        file_name=WORKBOOK_PATH.name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.sidebar.divider()
st.sidebar.markdown(
    '<p class="small-note">Recommended use: include this as a portfolio link alongside the Excel file so reviewers can see both modelling depth and business presentation.</p>',
    unsafe_allow_html=True,
)
