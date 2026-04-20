import json
import sqlite3
from datetime import datetime
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

APP_TITLE = "IT Audit Compliance and Risk Assessment Tool"
DATA_DIR = Path("data")
DB_PATH = Path("audit_runs.db")

REQUIRED_SCHEMAS = {
    "framework": [
        "framework",
        "control_id",
        "control_name",
        "category",
        "description",
        "criticality",
    ],
    "company": [
        "company_control_id",
        "control_name",
        "mapped_control_id",
        "owner",
        "frequency",
        "implemented",
    ],
    "testing": [
        "company_control_id",
        "test_result",
        "evidence_available",
        "remarks",
    ],
    "mapping": ["company_control_id", "framework", "control_id"],
}

VALID_FRAMEWORKS = {"ISO27001", "SOX"}
VALID_CRITICALITY = {"High", "Medium", "Low"}
VALID_IMPLEMENTED = {"Yes", "No"}
VALID_TEST_RESULTS = {"Pass", "Fail", "Partial"}
VALID_EVIDENCE = {"Yes", "No"}

CRITICALITY_WEIGHTS = {"High": 3, "Medium": 2, "Low": 1}
RESULT_SCORES = {"Pass": 1.0, "Partial": 0.5, "Fail": 0.0}


st.set_page_config(page_title=APP_TITLE, layout="wide", page_icon=":clipboard:")

st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700&family=IBM+Plex+Serif:wght@500;600&display=swap');

:root {
    --bg-0: #070b12;
    --bg-1: #0b1320;
    --bg-2: #111b2b;
    --card: #121d2d;
    --card-2: #172438;
    --ink: #e8effa;
    --muted: #9bb0c8;
    --accent: #20b486;
    --accent-2: #3f8cff;
    --warn: #ff6b6b;
    --border: rgba(154, 180, 208, 0.26);
}

html, body, [class*="css"] {
    font-family: 'Manrope', sans-serif;
    color: var(--ink);
}

.stApp {
    background:
        radial-gradient(1100px 650px at 12% 8%, rgba(62, 140, 255, 0.12), transparent 62%),
        radial-gradient(900px 540px at 86% 12%, rgba(32, 180, 134, 0.14), transparent 58%),
        linear-gradient(170deg, var(--bg-0), var(--bg-1) 46%, var(--bg-2));
}

h1, h2, h3 {
    font-family: 'IBM Plex Serif', serif;
    letter-spacing: 0.15px;
    color: #f4f8ff;
}

[data-testid="stAppViewContainer"] .main .block-container {
    padding-top: 1.2rem;
    padding-bottom: 2rem;
}

[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0d1524, #101a2a 60%, #0e1827);
    border-right: 1px solid var(--border);
}

[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 {
    color: var(--ink);
}

.kpi-card {
    background: linear-gradient(155deg, rgba(23, 36, 56, 0.92), rgba(18, 29, 45, 0.94));
    border: 1px solid var(--border);
    border-radius: 14px;
    padding: 14px;
    min-height: 102px;
    box-shadow: 0 12px 28px rgba(3, 8, 15, 0.34);
}

.kpi-label {
    font-size: 0.85rem;
    color: var(--muted);
    text-transform: uppercase;
    letter-spacing: 0.55px;
}

.kpi-value {
    font-size: 1.5rem;
    font-weight: 700;
    color: #f7fbff;
}

.section-banner {
    background: linear-gradient(108deg, rgba(20, 34, 55, 0.95), rgba(18, 44, 53, 0.94) 45%, rgba(20, 47, 84, 0.95));
    color: white;
    padding: 18px 20px 16px 20px;
    border: 1px solid var(--border);
    border-radius: 14px;
    margin-bottom: 16px;
    box-shadow: 0 16px 34px rgba(3, 9, 18, 0.42);
}

.section-banner h2 {
    margin: 0;
    line-height: 1.2;
}

.section-banner p {
    margin: 0.42rem 0 0 0;
    color: #bfcee1;
    font-size: 0.93rem;
}

[data-testid="stHorizontalBlock"] [data-testid="stMetric"] {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 0.5rem;
}

[data-testid="stTabs"] [role="tablist"] {
    gap: 0.45rem;
}

[data-testid="stTabs"] [role="tab"] {
    background: rgba(18, 33, 52, 0.86);
    border: 1px solid rgba(149, 176, 204, 0.24);
    color: #d0ddee;
    border-radius: 9px;
    padding: 0.5rem 0.95rem;
}

[data-testid="stTabs"] [aria-selected="true"] {
    background: linear-gradient(135deg, rgba(26, 98, 78, 0.93), rgba(31, 88, 168, 0.93));
    border-color: rgba(155, 220, 199, 0.45);
    color: #ffffff;
}

[data-testid="stDataFrame"], div[data-baseweb="select"], .stTextInput input,
.stTextArea textarea, .stNumberInput input {
    background: rgba(15, 25, 40, 0.88) !important;
    color: var(--ink) !important;
    border-radius: 10px !important;
    border-color: rgba(146, 175, 204, 0.30) !important;
}

.stButton button, .stDownloadButton button, [data-testid="baseButton-secondary"] {
    background: linear-gradient(135deg, #1f6d5c, #2c6ec2) !important;
    color: #f7fcff !important;
    border: 1px solid rgba(161, 214, 199, 0.40) !important;
    border-radius: 10px !important;
    transition: all 0.2s ease;
}

.stButton button:hover, .stDownloadButton button:hover {
    transform: translateY(-1px);
    filter: brightness(1.05);
}

[data-testid="stMarkdownContainer"] hr {
    border-color: rgba(147, 177, 208, 0.26);
}

[data-testid="stAlert"] {
    border-radius: 12px;
    border: 1px solid rgba(147, 177, 208, 0.22);
}
</style>
""",
    unsafe_allow_html=True,
)


def normalize_value(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_title_case(value):
    text = normalize_value(value)
    if not text:
        return ""
    return text[:1].upper() + text[1:].lower()


def normalize_upper(value):
    return normalize_value(value).upper()


@st.cache_data(show_spinner=False)
def read_csv_cached(file_obj):
    return pd.read_csv(file_obj)


def validate_schema(df, schema_name):
    required = REQUIRED_SCHEMAS[schema_name]
    missing = [col for col in required if col not in df.columns]
    extra = [col for col in df.columns if col not in required]
    errors = []
    if missing:
        errors.append(f"Missing columns: {', '.join(missing)}")
    if extra:
        errors.append(f"Unexpected columns: {', '.join(extra)}")
    return errors


def validate_values(framework_df, company_df, testing_df, mapping_df=None):
    errors = []

    frame_values = set(framework_df["framework"].astype(str).str.upper().str.strip())
    if not frame_values.issubset(VALID_FRAMEWORKS):
        errors.append("Framework file has invalid values in framework. Allowed: ISO27001, SOX")

    criticality_values = set(framework_df["criticality"].astype(str).str.title().str.strip())
    if not criticality_values.issubset(VALID_CRITICALITY):
        errors.append("Framework file has invalid criticality. Allowed: High, Medium, Low")

    implemented_values = set(company_df["implemented"].astype(str).str.title().str.strip())
    if not implemented_values.issubset(VALID_IMPLEMENTED):
        errors.append("Company file has invalid implemented values. Allowed: Yes, No")

    test_values = set(testing_df["test_result"].astype(str).str.title().str.strip())
    if not test_values.issubset(VALID_TEST_RESULTS):
        errors.append("Testing file has invalid test_result values. Allowed: Pass, Fail, Partial")

    evidence_values = set(testing_df["evidence_available"].astype(str).str.title().str.strip())
    if not evidence_values.issubset(VALID_EVIDENCE):
        errors.append("Testing file has invalid evidence_available values. Allowed: Yes, No")

    if mapping_df is not None:
        map_framework_values = set(mapping_df["framework"].astype(str).str.upper().str.strip())
        if not map_framework_values.issubset(VALID_FRAMEWORKS):
            errors.append("Mapping file has invalid framework values. Allowed: ISO27001, SOX")

    return errors


def parse_mapped_controls(value):
    raw = normalize_value(value)
    if not raw:
        return []
    for delimiter in [";", "|"]:
        raw = raw.replace(delimiter, ",")
    controls = [item.strip() for item in raw.split(",") if item.strip()]
    return controls


def standardize_dataframes(framework_df, company_df, testing_df, mapping_df=None):
    framework = framework_df.copy()
    framework["framework"] = framework["framework"].map(normalize_upper)
    framework["control_id"] = framework["control_id"].map(normalize_value)
    framework["control_name"] = framework["control_name"].map(normalize_value)
    framework["category"] = framework["category"].map(normalize_value)
    framework["description"] = framework["description"].map(normalize_value)
    framework["criticality"] = framework["criticality"].map(normalize_title_case)

    company = company_df.copy()
    company["company_control_id"] = company["company_control_id"].map(normalize_value)
    company["control_name"] = company["control_name"].map(normalize_value)
    company["mapped_control_id"] = company["mapped_control_id"].map(normalize_value)
    company["owner"] = company["owner"].map(normalize_value)
    company["frequency"] = company["frequency"].map(normalize_value)
    company["implemented"] = company["implemented"].map(normalize_title_case)

    testing = testing_df.copy()
    testing["company_control_id"] = testing["company_control_id"].map(normalize_value)
    testing["test_result"] = testing["test_result"].map(normalize_title_case)
    testing["evidence_available"] = testing["evidence_available"].map(normalize_title_case)
    testing["remarks"] = testing["remarks"].map(normalize_value)

    mapping = None
    if mapping_df is not None:
        mapping = mapping_df.copy()
        mapping["company_control_id"] = mapping["company_control_id"].map(normalize_value)
        mapping["framework"] = mapping["framework"].map(normalize_upper)
        mapping["control_id"] = mapping["control_id"].map(normalize_value)

    return framework, company, testing, mapping


def build_mapping_table(framework_df, company_df, mapping_df=None):
    if mapping_df is not None and not mapping_df.empty:
        mapped = company_df.merge(mapping_df, on="company_control_id", how="left")
        mapped = mapped.merge(
            framework_df,
            on=["framework", "control_id"],
            how="left",
            suffixes=("_company", "_framework"),
        )
    else:
        expanded = company_df.copy()
        expanded["control_id"] = expanded["mapped_control_id"].apply(parse_mapped_controls)
        expanded = expanded.explode("control_id")
        expanded["control_id"] = expanded["control_id"].fillna("").map(normalize_value)
        mapped = expanded.merge(
            framework_df,
            on="control_id",
            how="left",
            suffixes=("_company", "_framework"),
        )

    mapped = mapped.rename(columns={"control_name_company": "company_control_name"})

    if "control_name_framework" in mapped.columns:
        mapped = mapped.rename(columns={"control_name_framework": "framework_control_name"})
    else:
        mapped["framework_control_name"] = mapped.get("control_name", "")

    for col in ["framework", "control_id", "framework_control_name", "category", "criticality"]:
        if col not in mapped.columns:
            mapped[col] = ""

    return mapped


def combine_with_testing(mapped_df, testing_df):
    merged = mapped_df.merge(testing_df, on="company_control_id", how="left")
    merged["test_result"] = merged["test_result"].fillna("Fail")
    merged["evidence_available"] = merged["evidence_available"].fillna("No")
    merged["remarks"] = merged["remarks"].fillna("No remarks")
    merged["implemented"] = merged["implemented"].fillna("No")
    merged["criticality"] = merged["criticality"].replace("", "Medium").fillna("Medium")
    merged["category"] = merged["category"].replace("", "Unassigned").fillna("Unassigned")
    merged["framework"] = merged["framework"].replace("", "UNMAPPED").fillna("UNMAPPED")

    merged["score"] = merged["test_result"].map(RESULT_SCORES).fillna(0.0)
    return merged


def apply_gap_logic(df):
    out = df.copy()

    def issue_list(row):
        issues = []
        if row["implemented"] == "No":
            issues.append("GAP")
        if row["test_result"] == "Fail":
            issues.append("DEFICIENCY")
        if row["test_result"] == "Partial":
            issues.append("WEAKNESS")
        if row["evidence_available"] == "No":
            issues.append("EVIDENCE GAP")
        return issues

    out["issues"] = out.apply(issue_list, axis=1)
    out["issue_summary"] = out["issues"].apply(lambda x: ", ".join(x) if x else "None")
    out["has_gap"] = out["issues"].apply(lambda x: len(x) > 0)
    return out


def calculate_risk(df, high_threshold=2.5, medium_threshold=1.5):
    out = df.copy()

    def severity(row):
        levels = [0.0]
        if row["implemented"] == "No":
            levels.append(1.0)
        if row["test_result"] == "Fail":
            levels.append(1.0)
        if row["test_result"] == "Partial":
            levels.append(0.6)
        if row["evidence_available"] == "No":
            levels.append(0.4)
        return max(levels)

    out["criticality_weight"] = out["criticality"].map(CRITICALITY_WEIGHTS).fillna(2)
    out["failure_severity"] = out.apply(severity, axis=1)
    out["risk_score"] = out["criticality_weight"] * out["failure_severity"]

    def classify(score):
        if score >= high_threshold:
            return "High"
        if score >= medium_threshold:
            return "Medium"
        return "Low"

    out["risk_classification"] = out["risk_score"].apply(classify)
    return out


def compute_compliance_metrics(df):
    if df.empty:
        return {
            "overall_score": 0.0,
            "framework_scores": pd.DataFrame(columns=["framework", "score_percent"]),
            "category_scores": pd.DataFrame(columns=["category", "score_percent"]),
            "status_counts": pd.DataFrame(columns=["test_result", "count"]),
        }

    overall = round(float(df["score"].mean() * 100), 2)

    framework_scores = (
        df.groupby("framework", as_index=False)["score"].mean().rename(columns={"score": "score_percent"})
    )
    framework_scores["score_percent"] = (framework_scores["score_percent"] * 100).round(2)

    category_scores = (
        df.groupby("category", as_index=False)["score"].mean().rename(columns={"score": "score_percent"})
    )
    category_scores["score_percent"] = (category_scores["score_percent"] * 100).round(2)

    status_counts = df["test_result"].value_counts(dropna=False).reset_index()
    status_counts.columns = ["test_result", "count"]

    return {
        "overall_score": overall,
        "framework_scores": framework_scores,
        "category_scores": category_scores,
        "status_counts": status_counts,
    }


def recommendation_for_row(row):
    issues = row["issues"]
    category = row["category"]
    if "DEFICIENCY" in issues:
        return f"Perform immediate remediation and re-test control effectiveness in {category}."
    if "WEAKNESS" in issues:
        return f"Strengthen control design and increase monitoring frequency for {category}."
    if "GAP" in issues:
        return f"Implement missing control procedures and assign an accountable control owner for {category}."
    if "EVIDENCE GAP" in issues:
        return f"Improve evidence retention and documentation standards for {category}."
    return "No action required."


def generate_findings(df):
    findings_source = df[df["has_gap"]].copy()
    if findings_source.empty:
        return pd.DataFrame(
            columns=[
                "company_control_id",
                "company_control_name",
                "framework",
                "framework_control_name",
                "issue",
                "risk",
                "recommendation",
            ]
        )

    findings_source["recommendation"] = findings_source.apply(recommendation_for_row, axis=1)
    findings_source["risk"] = findings_source["risk_classification"]

    findings = findings_source[
        [
            "company_control_id",
            "company_control_name",
            "framework",
            "framework_control_name",
            "issue_summary",
            "risk",
            "recommendation",
        ]
    ].rename(columns={"issue_summary": "issue"})

    return findings.sort_values(by=["risk", "company_control_id"], ascending=[True, True])


def to_excel_report(summary_df, framework_scores, category_scores, findings_df, risk_df, evidence_df, mapped_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        framework_scores.to_excel(writer, sheet_name="Framework Scores", index=False)
        category_scores.to_excel(writer, sheet_name="Category Scores", index=False)
        findings_df.to_excel(writer, sheet_name="Findings", index=False)
        risk_df.to_excel(writer, sheet_name="Risk Register", index=False)
        evidence_df.to_excel(writer, sheet_name="Evidence Gaps", index=False)
        mapped_df.to_excel(writer, sheet_name="Mapped Controls", index=False)
    output.seek(0)
    return output


def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS audit_runs (
            run_id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_timestamp TEXT NOT NULL,
            framework_filter TEXT,
            overall_score REAL,
            total_controls INTEGER,
            pass_count INTEGER,
            partial_count INTEGER,
            fail_count INTEGER,
            gap_count INTEGER,
            high_risk_count INTEGER,
            snapshot_json TEXT
        )
        """
    )
    conn.commit()
    conn.close()


def save_audit_run(summary):
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        """
        INSERT INTO audit_runs (
            run_timestamp,
            framework_filter,
            overall_score,
            total_controls,
            pass_count,
            partial_count,
            fail_count,
            gap_count,
            high_risk_count,
            snapshot_json
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            summary.get("framework_filter", "ALL"),
            summary.get("overall_score", 0.0),
            summary.get("total_controls", 0),
            summary.get("pass_count", 0),
            summary.get("partial_count", 0),
            summary.get("fail_count", 0),
            summary.get("gap_count", 0),
            summary.get("high_risk_count", 0),
            json.dumps(summary),
        ),
    )
    conn.commit()
    conn.close()


def load_audit_runs(limit=20):
    if not DB_PATH.exists():
        return pd.DataFrame()
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query(
        """
        SELECT run_id, run_timestamp, framework_filter, overall_score,
               total_controls, pass_count, partial_count, fail_count,
               gap_count, high_risk_count
        FROM audit_runs
        ORDER BY run_id DESC
        LIMIT ?
        """,
        conn,
        params=[limit],
    )
    conn.close()
    return df


def load_sample_data():
    framework = pd.read_csv(DATA_DIR / "sample_framework_controls.csv")
    company = pd.read_csv(DATA_DIR / "sample_company_controls.csv")
    testing = pd.read_csv(DATA_DIR / "sample_audit_testing.csv")
    mapping = pd.read_csv(DATA_DIR / "sample_mapping_table.csv")
    return framework, company, testing, mapping


def render_kpi(label, value):
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def apply_enterprise_chart_theme(fig, height=360):
    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(11, 19, 32, 0.84)",
        font={"family": "Manrope, sans-serif", "color": "#dce6f2", "size": 13},
        title_font={"family": "IBM Plex Serif, serif", "size": 18, "color": "#f4f8ff"},
        margin={"l": 14, "r": 14, "t": 56, "b": 20},
        height=height,
    )
    fig.update_xaxes(showgrid=True, gridcolor="rgba(144, 171, 199, 0.16)", zeroline=False)
    fig.update_yaxes(showgrid=True, gridcolor="rgba(144, 171, 199, 0.16)", zeroline=False)
    return fig


st.markdown(
    f"""
    <div class='section-banner'>
        <h2>{APP_TITLE}</h2>
        <p>Enterprise visibility for control health, compliance posture, and risk trends.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

init_db()

if "manual_controls" not in st.session_state:
    st.session_state.manual_controls = pd.DataFrame(columns=REQUIRED_SCHEMAS["company"])

with st.sidebar:
    st.header("Audit Inputs")
    use_sample = st.checkbox("Use sample data", value=True)

    framework_file = st.file_uploader("Framework Controls CSV", type=["csv"])
    company_file = st.file_uploader("Company Controls CSV", type=["csv"])
    testing_file = st.file_uploader("Audit Testing CSV", type=["csv"])
    mapping_file = st.file_uploader("Mapping Table CSV (Optional)", type=["csv"])

    framework_filter = st.multiselect("Framework Filter", options=["ISO27001", "SOX"], default=["ISO27001", "SOX"])

    st.markdown("---")
    st.subheader("Manual Control Input")
    with st.form("manual_control_form"):
        mc_id = st.text_input("Company Control ID")
        mc_name = st.text_input("Control Name")
        mc_map = st.text_input("Mapped Control IDs (comma-separated)")
        mc_owner = st.text_input("Owner")
        mc_freq = st.selectbox("Frequency", options=["Daily", "Weekly", "Monthly", "Quarterly", "Yearly"])
        mc_impl = st.selectbox("Implemented", options=["Yes", "No"])
        add_manual = st.form_submit_button("Add Control")

    if add_manual:
        if mc_id.strip() and mc_name.strip():
            new_row = pd.DataFrame(
                [
                    {
                        "company_control_id": mc_id.strip(),
                        "control_name": mc_name.strip(),
                        "mapped_control_id": mc_map.strip(),
                        "owner": mc_owner.strip(),
                        "frequency": mc_freq,
                        "implemented": mc_impl,
                    }
                ]
            )
            st.session_state.manual_controls = pd.concat(
                [st.session_state.manual_controls, new_row], ignore_index=True
            )
            st.success("Manual control added to this run.")
        else:
            st.error("Company Control ID and Control Name are required.")

    st.markdown("---")
    st.subheader("Previous Audit Runs")
    runs_df = load_audit_runs(limit=10)
    if runs_df.empty:
        st.caption("No previous runs saved yet.")
    else:
        st.dataframe(runs_df, use_container_width=True, hide_index=True)

try:
    if use_sample:
        framework_df, company_df, testing_df, mapping_df = load_sample_data()
    else:
        if not framework_file or not company_file or not testing_file:
            st.info("Upload required datasets or enable sample data to continue.")
            st.stop()

        framework_df = read_csv_cached(framework_file)
        company_df = read_csv_cached(company_file)
        testing_df = read_csv_cached(testing_file)
        mapping_df = read_csv_cached(mapping_file) if mapping_file else None

    schema_errors = []
    schema_errors.extend(validate_schema(framework_df, "framework"))
    schema_errors.extend(validate_schema(company_df, "company"))
    schema_errors.extend(validate_schema(testing_df, "testing"))
    if mapping_df is not None:
        schema_errors.extend(validate_schema(mapping_df, "mapping"))

    if schema_errors:
        st.error("Schema validation failed:")
        for err in schema_errors:
            st.write(f"- {err}")
        st.stop()

    framework_df, company_df, testing_df, mapping_df = standardize_dataframes(
        framework_df, company_df, testing_df, mapping_df
    )

    value_errors = validate_values(framework_df, company_df, testing_df, mapping_df)
    if value_errors:
        st.error("Value validation failed:")
        for err in value_errors:
            st.write(f"- {err}")
        st.stop()

    if not st.session_state.manual_controls.empty:
        company_df = pd.concat([company_df, st.session_state.manual_controls], ignore_index=True)

    mapped_df = build_mapping_table(framework_df, company_df, mapping_df)
    merged_df = combine_with_testing(mapped_df, testing_df)

    if framework_filter:
        merged_df = merged_df[merged_df["framework"].isin(framework_filter)]

    merged_df = apply_gap_logic(merged_df)

    high_threshold = 2.5
    medium_threshold = 1.5
    merged_df = calculate_risk(merged_df, high_threshold=high_threshold, medium_threshold=medium_threshold)

    metrics = compute_compliance_metrics(merged_df)
    findings_df = generate_findings(merged_df)

    evidence_gaps_df = merged_df[merged_df["evidence_available"] == "No"][
        [
            "company_control_id",
            "company_control_name",
            "framework",
            "framework_control_name",
            "owner",
            "remarks",
        ]
    ].drop_duplicates()

    failures_df = merged_df[
        (merged_df["test_result"].isin(["Fail", "Partial"]))
        | (merged_df["implemented"] == "No")
        | (merged_df["evidence_available"] == "No")
    ][
        [
            "company_control_id",
            "company_control_name",
            "framework",
            "framework_control_name",
            "category",
            "criticality",
            "test_result",
            "implemented",
            "evidence_available",
            "issue_summary",
            "risk_classification",
        ]
    ].drop_duplicates()

    tabs = st.tabs(["Dashboard", "Control Mapping", "Gap Analysis", "Risk Analysis", "Audit Report"])

    with tabs[0]:
        st.subheader("Compliance Dashboard")

        total_controls = int(len(merged_df))
        pass_count = int((merged_df["test_result"] == "Pass").sum())
        partial_count = int((merged_df["test_result"] == "Partial").sum())
        fail_count = int((merged_df["test_result"] == "Fail").sum())

        c1, c2, c3, c4, c5 = st.columns(5)
        with c1:
            render_kpi("Total Controls", f"{total_controls}")
        with c2:
            render_kpi("Passed", f"{pass_count}")
        with c3:
            render_kpi("Partial", f"{partial_count}")
        with c4:
            render_kpi("Failed", f"{fail_count}")
        with c5:
            render_kpi("Overall Score", f"{metrics['overall_score']:.2f}%")

        chart_col1, chart_col2 = st.columns(2)

        with chart_col1:
            if not metrics["category_scores"].empty:
                bar_fig = px.bar(
                    metrics["category_scores"],
                    x="category",
                    y="score_percent",
                    color="score_percent",
                    color_continuous_scale=["#2f6cad", "#26bf93"],
                    title="Compliance by Category (%)",
                )
                apply_enterprise_chart_theme(bar_fig, height=360)
                st.plotly_chart(bar_fig, use_container_width=True)

        with chart_col2:
            if not metrics["status_counts"].empty:
                pie_fig = px.pie(
                    metrics["status_counts"],
                    names="test_result",
                    values="count",
                    hole=0.45,
                    title="Testing Outcome Split",
                    color="test_result",
                    color_discrete_map={"Pass": "#20b486", "Partial": "#ffc247", "Fail": "#ff6b6b"},
                )
                apply_enterprise_chart_theme(pie_fig, height=360)
                st.plotly_chart(pie_fig, use_container_width=True)

        st.markdown("### Failed and Weak Controls")
        if failures_df.empty:
            st.success("No failed, weak, or missing evidence controls in selected framework filter.")
        else:
            st.dataframe(failures_df, use_container_width=True, hide_index=True)

        st.markdown("### Risk Heatmap")
        heat_data = (
            merged_df.groupby(["category", "framework"], as_index=False)["risk_score"].mean()
            .pivot(index="category", columns="framework", values="risk_score")
            .fillna(0)
        )
        if heat_data.empty:
            st.info("No risk data available.")
        else:
            heat_fig = px.imshow(
                heat_data,
                aspect="auto",
                color_continuous_scale=["#1d3562", "#2f89e6", "#f3b74f", "#f36e6e"],
                labels={"color": "Avg Risk Score"},
                title="Average Risk Score by Category and Framework",
            )
            apply_enterprise_chart_theme(heat_fig, height=420)
            st.plotly_chart(heat_fig, use_container_width=True)

    with tabs[1]:
        st.subheader("Control Mapping")
        display_mapping = merged_df[
            [
                "company_control_id",
                "company_control_name",
                "owner",
                "frequency",
                "implemented",
                "framework",
                "control_id",
                "framework_control_name",
                "category",
                "criticality",
            ]
        ].drop_duplicates()

        st.dataframe(display_mapping, use_container_width=True, hide_index=True)

        st.markdown("### Framework-wise Compliance Scores")
        st.dataframe(metrics["framework_scores"], use_container_width=True, hide_index=True)

    with tabs[2]:
        st.subheader("Gap Analysis")
        gap_rows = merged_df[merged_df["has_gap"]].copy()

        if gap_rows.empty:
            st.success("No gaps identified based on current data.")
        else:
            gap_counter = (
                gap_rows["issues"].explode().value_counts().rename_axis("issue_type").reset_index(name="count")
            )
            g1, g2 = st.columns([1, 2])
            with g1:
                st.markdown("#### Gap Type Counts")
                st.dataframe(gap_counter, use_container_width=True, hide_index=True)
            with g2:
                st.markdown("#### Detailed Gaps")
                st.dataframe(
                    gap_rows[
                        [
                            "company_control_id",
                            "company_control_name",
                            "framework",
                            "framework_control_name",
                            "implemented",
                            "test_result",
                            "evidence_available",
                            "issue_summary",
                            "remarks",
                        ]
                    ],
                    use_container_width=True,
                    hide_index=True,
                )

        st.markdown("### Evidence Tracker")
        if evidence_gaps_df.empty:
            st.success("Evidence is available for all controls.")
        else:
            st.dataframe(evidence_gaps_df, use_container_width=True, hide_index=True)

    with tabs[3]:
        st.subheader("Risk Analysis")

        t1, t2 = st.columns(2)
        with t1:
            adj_high = st.slider("High Risk Threshold", min_value=2.0, max_value=3.0, value=2.5, step=0.1)
        with t2:
            adj_medium = st.slider("Medium Risk Threshold", min_value=0.5, max_value=2.0, value=1.5, step=0.1)

        risk_df = calculate_risk(merged_df, high_threshold=adj_high, medium_threshold=adj_medium)

        risk_summary = (
            risk_df["risk_classification"].value_counts().rename_axis("risk_classification").reset_index(name="count")
        )

        r1, r2 = st.columns(2)
        with r1:
            st.markdown("#### Risk Classification Summary")
            st.dataframe(risk_summary, use_container_width=True, hide_index=True)
        with r2:
            risk_pie = px.pie(
                risk_summary,
                names="risk_classification",
                values="count",
                title="Risk Distribution",
                color="risk_classification",
                color_discrete_map={"High": "#ff6b6b", "Medium": "#ffc247", "Low": "#20b486"},
            )
            apply_enterprise_chart_theme(risk_pie, height=360)
            st.plotly_chart(risk_pie, use_container_width=True)

        st.markdown("#### Detailed Risk Register")
        st.dataframe(
            risk_df[
                [
                    "company_control_id",
                    "company_control_name",
                    "framework",
                    "framework_control_name",
                    "category",
                    "criticality",
                    "failure_severity",
                    "risk_score",
                    "risk_classification",
                    "issue_summary",
                ]
            ].sort_values(by=["risk_score", "criticality"], ascending=[False, True]),
            use_container_width=True,
            hide_index=True,
        )

    with tabs[4]:
        st.subheader("Audit Report")

        report_summary = pd.DataFrame(
            [
                {
                    "Generated At": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Framework Filter": ", ".join(framework_filter) if framework_filter else "ALL",
                    "Total Controls": total_controls,
                    "Passed": pass_count,
                    "Partial": partial_count,
                    "Failed": fail_count,
                    "Overall Compliance Score (%)": metrics["overall_score"],
                    "Total Gaps": int(merged_df["has_gap"].sum()),
                    "High Risk Controls": int((merged_df["risk_classification"] == "High").sum()),
                }
            ]
        )

        st.markdown("### Executive Summary")
        st.dataframe(report_summary, use_container_width=True, hide_index=True)

        st.markdown("### Audit Findings")
        if findings_df.empty:
            st.success("No findings to report.")
        else:
            st.dataframe(findings_df, use_container_width=True, hide_index=True)

        narrative_lines = []
        for _, row in findings_df.head(25).iterrows():
            line = (
                f"Control: {row['company_control_name']} | Issue: {row['issue']} | "
                f"Risk: {row['risk']} | Recommendation: {row['recommendation']}"
            )
            narrative_lines.append(line)

        st.markdown("### Finding Narrative")
        if narrative_lines:
            st.text_area("Auto-generated finding statements", value="\n".join(narrative_lines), height=220)
        else:
            st.caption("No narrative generated because there are no findings.")

        export_bytes = to_excel_report(
            summary_df=report_summary,
            framework_scores=metrics["framework_scores"],
            category_scores=metrics["category_scores"],
            findings_df=findings_df,
            risk_df=merged_df[
                [
                    "company_control_id",
                    "company_control_name",
                    "framework",
                    "framework_control_name",
                    "category",
                    "criticality",
                    "risk_score",
                    "risk_classification",
                    "issue_summary",
                ]
            ],
            evidence_df=evidence_gaps_df,
            mapped_df=merged_df,
        )

        st.download_button(
            label="Download Audit Report (Excel)",
            data=export_bytes,
            file_name=f"audit_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        if st.button("Save Current Audit Run"):
            save_audit_run(
                {
                    "framework_filter": ",".join(framework_filter) if framework_filter else "ALL",
                    "overall_score": metrics["overall_score"],
                    "total_controls": total_controls,
                    "pass_count": pass_count,
                    "partial_count": partial_count,
                    "fail_count": fail_count,
                    "gap_count": int(merged_df["has_gap"].sum()),
                    "high_risk_count": int((merged_df["risk_classification"] == "High").sum()),
                }
            )
            st.success("Audit run saved in SQLite.")

except Exception as exc:
    st.error("An unexpected error occurred while processing the data.")
    st.exception(exc)
