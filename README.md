# IT Audit Compliance Tool (Streamlit)

A web-based IT audit compliance tool that maps company controls to frameworks (ISO 27001 and SOX), evaluates control testing, calculates compliance scores, identifies gaps, and generates audit-ready reports.

## Features

- Upload and validate framework, company control, audit testing, and optional mapping CSV files
- Map controls across frameworks with one-to-many support
- Compliance scoring engine:
  - Pass = 1
  - Partial = 0.5
  - Fail = 0
- Gap analysis:
  - implemented = No -> GAP
  - test_result = Fail -> DEFICIENCY
  - test_result = Partial -> WEAKNESS
  - evidence_available = No -> EVIDENCE GAP
- Risk scoring and classification
- Dashboard with KPI cards, charts, failed controls, and risk heatmap
- Audit findings generator with recommendations
- Evidence tracker
- Downloadable Excel audit report
- Bonus:
  - Save previous audit runs (SQLite)
  - Manual control input form

## Expected CSV Schemas

### 1) Framework Controls
Required columns:
- framework
- control_id
- control_name
- category
- description
- criticality

### 2) Company Controls
Required columns:
- company_control_id
- control_name
- mapped_control_id
- owner
- frequency
- implemented

### 3) Audit Testing
Required columns:
- company_control_id
- test_result
- evidence_available
- remarks

### 4) Optional Mapping Table
Required columns:
- company_control_id
- framework
- control_id

## Quick Start

1. Create and activate a Python environment.
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the app:
   ```bash
   streamlit run app.py
   ```
4. Upload your CSVs from the sidebar, or enable sample data.

## Notes

- If no separate mapping table is uploaded, the app uses `mapped_control_id` from company controls. You can provide one or multiple control IDs separated by comma/semicolon/pipe.
- The tool is designed for audit simulation and operational insights.
