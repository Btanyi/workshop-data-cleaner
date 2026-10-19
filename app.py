#pip install pandas streamlit openpyxl
from io import BytesIO
import re

import pandas as pd
import streamlit as st
from openpyxl.styles import PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

_EXCEL_FORMULA_PREFIX = re.compile(r"^[\s\x00-\x1f]*[=+\-@]")


def _excel_safe_export_copy(dataframe):
    export_df = dataframe.copy(deep=True)
    export_df.columns = [
        _excel_safe_export_value(column) for column in export_df.columns
    ]
    for column_position in range(export_df.shape[1]):
        export_df.isetitem(
            column_position,
            export_df.iloc[:, column_position].map(_excel_safe_export_value),
        )
    return export_df


def _excel_safe_export_value(value):
    if isinstance(value, str) and _EXCEL_FORMULA_PREFIX.match(value):
        return "'" + value
    return value


def _render_progress(progress_slot, current_step):
    steps = [
        ("Upload", "&#10003;"),
        ("Preview", "&#9673;"),
        ("Clean", "&#9881;"),
        ("Download", "&#8681;"),
    ]
    current_index = {
        "upload": 0,
        "preview": 1,
        "clean": 2,
        "download": 3,
        "done": 4,
    }[current_step]
    step_markup = []
    for index, (label, icon) in enumerate(steps):
        status = "complete" if index < current_index else "pending"
        if index == current_index and current_index < len(steps):
            status = "current"
        current_attribute = ' aria-current="step"' if status == "current" else ""
        step_markup.append(
            f'<div class="audit-step {status}" role="listitem"{current_attribute}>'
            f'<span class="audit-step-icon" aria-hidden="true">{icon}</span>'
            f'<span class="audit-step-label">{label}</span>'
            "</div>"
        )

    progress_slot.markdown(
        """
        <div class="audit-progress-card">
          <div class="audit-progress" role="list" aria-label="Audit workflow progress">
            {steps}
          </div>
        </div>
        """.replace("{steps}", "".join(step_markup)),
        unsafe_allow_html=True,
    )


def _mark_downloaded():
    st.session_state["_incentive_audit_stage"] = "done"


def _table_column_config(dataframe, cleaned=False):
    widths = (
        {
            "Student_ID": 98,
            "Name": 119,
            "Incentive_Amount": 105,
            "Term": 105,
            "Audit_Status": 105,
            "Audit_Notes": 168,
        }
        if cleaned
        else {
            "Student_ID": 110,
            "Name": 170,
            "Incentive_Amount": 150,
            "Term": 150,
        }
    )
    config = {}
    for column, width in widths.items():
        if column not in dataframe.columns:
            continue
        if column == "Incentive_Amount":
            config[column] = st.column_config.NumberColumn(
                width=width, alignment="right", format="$%.2f"
            )
        elif column == "Audit_Status":
            config[column] = st.column_config.TextColumn(
                width=width, alignment="center"
            )
        else:
            config[column] = st.column_config.TextColumn(
                width=width, alignment="left"
            )
    return config


st.set_page_config(page_title="Student Data Cleaner", page_icon="💵", layout="wide")

st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600;700&family=Space+Grotesk:wght@400;500;600;700&display=swap');

        :root,
        html,
        body,
        .stApp {
            color-scheme: light;
        }
        html, body, [data-testid="stAppViewContainer"], .stApp {
            background-color: #F1F6FC !important;
            color: #1E293B;
        }
        input, textarea, select, button {
            color-scheme: light;
        }
        .stApp {
            background: #F1F6FC;
            color: #1E293B;
            font-family: "IBM Plex Mono", monospace;
        }
        .stApp input, .stApp textarea, .stApp select, .stApp button,
        .stApp [data-testid="stMarkdownContainer"],
        .stApp [data-testid="stWidgetLabel"],
        .stApp [data-testid="stMetric"] {
            font-family: "IBM Plex Mono", monospace;
        }
        .stApp h1, .stApp h2, .stApp h3, .stApp h4,
        .cleaner-hero h1 {
            font-family: "Space Grotesk", sans-serif;
        }
        .block-container {
            max-width: 1320px;
            padding: 2rem 2.5rem 3rem;
        }
        .cleaner-hero {
            position: relative;
            display: grid;
            grid-template-columns: minmax(0, 1fr) minmax(180px, 0.5fr);
            align-items: center;
            gap: 2rem;
            background: linear-gradient(120deg, #123b70 0%, #1769aa 58%, #2384c6 100%);
            color: #fff;
            border: 1px solid #0e3668;
            border-radius: 24px;
            padding: 2.5rem 3rem;
            min-height: 250px;
            box-shadow: 0 16px 36px rgba(21, 73, 124, 0.18);
        }
        .cleaner-hero::before {
            content: "";
            position: absolute;
            inset: 0;
            border-radius: inherit;
            background-image: radial-gradient(
                circle,
                rgba(255, 255, 255, 0.2) 1px,
                transparent 1.4px
            );
            background-size: 18px 18px;
            opacity: 0.32;
            pointer-events: none;
        }
        .cleaner-hero-copy {
            position: relative;
            z-index: 1;
        }
        .cleaner-hero h1 {
            margin: 0;
            color: #fff;
            font-size: clamp(2rem, 4vw, 3rem);
            line-height: 1.1;
            font-weight: 700;
            letter-spacing: -0.035em;
        }
        .cleaner-hero p {
            max-width: 470px;
            margin: 0.8rem 0 0;
            color: #fff;
            text-align: left;
            font-size: 1.05rem;
            line-height: 1.35;
        }
        .hero-globe {
            position: relative;
            z-index: 1;
            display: block;
            width: min(100%, 220px);
            justify-self: center;
            overflow: visible;
            filter: drop-shadow(0 12px 18px rgba(5, 30, 65, 0.24));
        }
        .hero-globe-surface {
            transform-box: fill-box;
            transform-origin: center;
            animation: globe-spin 18s linear infinite;
        }
        .hero-globe-latitude {
            transform-box: fill-box;
            transform-origin: center;
            animation: latitude-shift 5s ease-in-out infinite alternate;
        }
        .confetti-burst {
            position: fixed;
            z-index: 100000;
            top: 45%;
            left: 50%;
            width: 0;
            height: 0;
            pointer-events: none;
        }
        .confetti-burst span {
            position: absolute;
            width: 9px;
            height: 14px;
            border-radius: 2px;
            background: #1769aa;
            animation: confetti-explode 1.25s cubic-bezier(0.12, 0.7, 0.3, 1) both;
        }
        .confetti-burst span:nth-child(4n + 1) { background: #f4c95d; }
        .confetti-burst span:nth-child(4n + 2) { background: #66c7a3; }
        .confetti-burst span:nth-child(4n + 3) { background: #ef7891; }
        .confetti-burst span:nth-child(4n) { background: #78b9e8; }
        @keyframes confetti-explode {
            0% {
                opacity: 1;
                transform: translate(-50%, -50%) scale(0.4) rotate(0deg);
            }
            100% {
                opacity: 0;
                transform: translate(var(--dx), var(--dy)) scale(1) rotate(650deg);
            }
        }
        @keyframes globe-spin {
            from { transform: rotate(0deg); }
            to { transform: rotate(360deg); }
        }
        @keyframes latitude-shift {
            from { transform: scaleX(0.82); }
            to { transform: scaleX(1); }
        }
        .cleaner-features {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 1.25rem;
            padding: 1.5rem 0;
            color: #1E293B;
            font-size: 1rem;
            font-weight: 600;
        }
        .cleaner-features ul {
            list-style: none;
            margin: 0;
            padding: 1.1rem 1.4rem;
            background: #fff;
            border: 1px solid #E2E8F0;
            border-radius: 16px;
            box-shadow: 0 5px 16px rgba(33, 76, 124, 0.06);
        }
        .cleaner-features li {
            padding: 0.35rem 0;
        }
        .cleaner-features li::before {
            content: "✓";
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 1.35rem;
            height: 1.35rem;
            margin-right: 0.6rem;
            border-radius: 50%;
            background: #e5f1ff;
            color: #1769aa;
            font-size: 0.85rem;
            font-weight: 800;
        }
        .upload-section {
            margin-top: 0.25rem;
        }
        .audit-progress-card {
            background: #fff;
            border: 1px solid #E2E8F0;
            border-radius: 16px;
            margin: 0.25rem 0 1.5rem;
            padding: 18px 22px 14px;
        }
        .audit-progress {
            display: flex;
            width: 100%;
        }
        .audit-step {
            align-items: center;
            display: flex;
            flex: 1;
            flex-direction: column;
            min-width: 0;
            position: relative;
        }
        .audit-step:not(:last-child)::after {
            background: #e5e7eb;
            content: "";
            height: 3px;
            left: calc(50% + 22px);
            position: absolute;
            right: calc(-50% + 22px);
            top: 20px;
        }
        .audit-step.complete:not(:last-child)::after {
            background: #20bf68;
        }
        .audit-step-icon {
            align-items: center;
            background: #f1f2f4;
            border-radius: 50%;
            color: #8b919c;
            display: flex;
            font-size: 17px;
            height: 42px;
            justify-content: center;
            position: relative;
            width: 42px;
            z-index: 1;
        }
        .audit-step.complete .audit-step-icon {
            background: #20bf68;
            color: #fff;
        }
        .audit-step.current .audit-step-icon {
            background: #6558f5;
            box-shadow: 0 4px 14px #6558f54a;
            color: #fff;
        }
        .audit-step-label {
            color: #8b919c;
            font-size: 0.78rem;
            font-weight: 500;
            margin-top: 8px;
            text-align: center;
        }
        .audit-step.complete .audit-step-label {
            color: #20a95c;
        }
        .audit-step.current .audit-step-label {
            color: #6558f5;
            font-weight: 700;
        }
        div[data-testid="stFileUploader"] section {
            background: #FFFFFF !important;
            border: 1px dashed #9BB8D8 !important;
            border-radius: 16px;
            padding: 1rem;
            min-height: 125px;
            box-shadow: 0 5px 16px rgba(33, 76, 124, 0.05);
        }
        div[data-testid="stFileUploader"] section,
        div[data-testid="stFileUploader"] section * {
            color-scheme: light;
        }
        div[data-testid="stFileUploader"] label,
        [data-testid="stWidgetLabel"] p,
        label {
            color: #1E293B;
        }
        .stFileUploader label {
            color: #1E293B;
            font-size: 1.1rem;
            font-weight: 700;
        }
        input, textarea, [data-baseweb="select"] > div {
            background-color: #FFFFFF !important;
            border-color: #E2E8F0 !important;
            color: #1E293B !important;
        }
        input::placeholder, textarea::placeholder {
            color: #64748B !important;
            opacity: 1;
        }
        [data-testid="stBaseButton-primary"] {
            background-color: #1769AA;
            border-color: #1769AA;
            color: #FFFFFF;
        }
        [data-testid="stBaseButton-secondary"] {
            background-color: #eaf3ff;
            border-color: #c4d9f1;
            color: #164e83;
            border-radius: 10px;
            font-weight: 600;
        }
        [data-testid="stBaseButton-secondary"]:hover {
            background-color: #dcecff;
            border-color: #8eb6e2;
            color: #103c68;
        }
        [data-testid="stDataFrame"] {
            background: #fff;
            border-radius: 10px;
        }
        [data-testid="stDataFrame"] [data-testid="stDataFrameResizable"] {
            --gdg-bg-cell: #FFFFFF;
            --gdg-bg-cell-medium: #FFFFFF;
            --gdg-bg-header: #F8FAFC;
            --gdg-bg-header-hovered: #F8FAFC;
            --gdg-text-dark: #1E293B;
            --gdg-text-medium: #1E293B;
            --gdg-text-header: #1E293B;
            --gdg-border-color: #E2E8F0;
            --gdg-horizontal-border-color: #E2E8F0;
            --gdg-font-family: "IBM Plex Mono", monospace;
            --gdg-header-font-style: 600 13px;
            --gdg-base-font-style: 13px;
            --gdg-cell-horizontal-padding: 15px;
            --gdg-cell-vertical-padding: 11px;
            background: #FFFFFF;
            border: 1px solid #E2E8F0 !important;
            border-radius: 10px !important;
            font-family: "IBM Plex Mono", monospace;
        }
        [data-testid="stDataFrame"] [role="grid"] th,
        [data-testid="stDataFrame"] [role="grid"] td {
            font-family: "IBM Plex Mono", monospace;
        }
        [data-testid="stDataFrame"] [role="grid"] th {
            background: #F8FAFC;
            color: #1E293B;
            font-size: 13px;
            font-weight: 600;
            padding: 10px 15px !important;
            text-align: left;
            vertical-align: middle;
            white-space: nowrap;
        }
        [data-testid="stDataFrame"] [role="grid"] td {
            background: #FFFFFF;
            border-bottom: 1px solid #E2E8F0;
            color: #1E293B;
            font-size: 13px;
            padding: 10px 15px !important;
            vertical-align: middle;
            white-space: nowrap;
        }
        [data-testid="stDataFrame"] [role="grid"] tr:hover td {
            background: #FFFFFF;
        }
        [data-testid="stMetric"] {
            box-sizing: border-box;
            height: 112px;
            padding: 1.15rem 1rem;
            border: 0;
            border-radius: 14px;
            box-shadow: 0 7px 18px rgba(33, 76, 124, 0.12);
            overflow: hidden;
            position: relative;
        }
        .st-key-audit-summary [data-testid="stColumn"]:nth-child(1) [data-testid="stMetric"] {
            background: linear-gradient(135deg, #347ff0, #2362db);
        }
        .st-key-audit-summary [data-testid="stColumn"]:nth-child(2) [data-testid="stMetric"] {
            background: linear-gradient(135deg, #20c568, #11a957);
        }
        .st-key-audit-summary [data-testid="stColumn"]:nth-child(3) [data-testid="stMetric"] {
            background: linear-gradient(135deg, #f27616, #e9550a);
        }
        .st-key-audit-summary [data-testid="stColumn"]:nth-child(4) [data-testid="stMetric"] {
            background: linear-gradient(135deg, #9950ec, #7736dc);
        }
        .st-key-audit-summary [data-testid="stMetric"] > div:first-child {
            max-width: calc(100% - 2rem);
            min-width: 0;
        }
        .st-key-audit-summary [data-testid="stMetricLabel"] .stMarkdownEllipsis {
            overflow: visible;
            text-overflow: clip;
            white-space: normal;
        }
        .st-key-audit-summary [data-testid="stMetricLabel"] p {
            color: rgba(255, 255, 255, 0.9);
            font-size: 0.66rem;
            font-weight: 700;
            letter-spacing: 0.02em;
            line-height: 1.1;
            text-transform: uppercase;
            white-space: normal;
        }
        .st-key-audit-summary [data-testid="stMetricValue"] {
            color: #fff;
            font-size: clamp(1.35rem, 2vw, 1.8rem);
            font-weight: 750;
        }
        .st-key-audit-summary [data-testid="stColumn"]:nth-child(4) [data-testid="stMetricValue"] {
            font-size: clamp(1.35rem, 2vw, 1.8rem);
            line-height: 1.2;
            white-space: nowrap;
        }
        .st-key-audit-summary [data-testid="stColumn"]:nth-child(4) [data-testid="stMetricValue"] .stMarkdownEllipsis {
            overflow: visible;
            text-overflow: clip;
            white-space: nowrap;
        }
        .st-key-audit-summary [data-testid="stMetric"]::after {
            align-items: center;
            background: rgba(255, 255, 255, 0.2);
            border-radius: 9px;
            color: rgba(255, 255, 255, 0.95);
            display: flex;
            font-size: 1.35rem;
            height: 32px;
            justify-content: center;
            position: absolute;
            right: 0.5rem;
            top: 50%;
            transform: translateY(-50%);
            width: 32px;
            background-position: center;
            background-repeat: no-repeat;
            background-size: 22px 22px;
        }
        .st-key-audit-summary [data-testid="stColumn"]:nth-child(1) [data-testid="stMetric"]::after {
            background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%23fff' stroke-width='1.8' stroke-linecap='round' stroke-linejoin='round'%3E%3Cellipse cx='12' cy='5' rx='8' ry='3'/%3E%3Cpath d='M4 5v14c0 1.7 3.6 3 8 3s8-1.3 8-3V5M4 12c0 1.7 3.6 3 8 3s8-1.3 8-3'/%3E%3C/svg%3E");
            content: "";
        }
        .st-key-audit-summary [data-testid="stColumn"]:nth-child(2) [data-testid="stMetric"]::after {
            background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%23fff' stroke-width='1.8' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M4 20V10m5 10V4m6 16v-7m5 7V7'/%3E%3C/svg%3E");
            content: "";
        }
        .st-key-audit-summary [data-testid="stColumn"]:nth-child(3) [data-testid="stMetric"]::after {
            background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%23fff' stroke-width='1.8' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M10.3 3.9 2.8 17a2 2 0 0 0 1.7 3h15a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0Z'/%3E%3Cpath d='M12 9v4m0 3h.01'/%3E%3C/svg%3E");
            content: "";
        }
        .st-key-audit-summary [data-testid="stColumn"]:nth-child(4) [data-testid="stMetric"]::after {
            background-image: none;
            content: "$";
        }
        h2, h3, label {
            color: #1E293B;
        }
        @media (max-width: 700px) {
            .block-container {
                padding: 1rem 1rem 2rem;
            }
            .audit-progress-card {
                padding-left: 10px;
                padding-right: 10px;
            }
            .audit-step:not(:last-child)::after {
                left: calc(50% + 17px);
                right: calc(-50% + 17px);
                top: 16px;
            }
            .audit-step-icon {
                font-size: 15px;
                height: 34px;
                width: 34px;
            }
            .audit-step-label {
                font-size: 0.68rem;
            }
            .cleaner-hero, .cleaner-features {
                grid-template-columns: 1fr;
            }
            .hero-globe {
                position: absolute;
                width: 110px;
                right: 0.75rem;
                bottom: 0.5rem;
                opacity: 0.38;
            }
            .cleaner-hero { position: relative; }
            .cleaner-features {
                padding: 1rem 0;
            }
        }
        @media (prefers-reduced-motion: reduce) {
            .hero-globe-surface, .hero-globe-latitude {
                animation: none;
            }
            .confetti-burst {
                display: none;
            }
        }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <header class="cleaner-hero">
            <div class="cleaner-hero-copy">
                <h1>Workshop &amp;<br>Incentive Data<br>Cleaner</h1>
                <p>Transforming raw student data into a clean and structured format for analysis and reporting</p>
            </div>
            <svg class="hero-globe" viewBox="0 0 200 200" role="img" aria-label="Animated globe">
                <defs>
                    <clipPath id="globe-clip">
                        <circle cx="100" cy="100" r="76"/>
                    </clipPath>
                    <linearGradient id="globe-water" x1="0" y1="0" x2="1" y2="1">
                        <stop offset="0%" stop-color="#d9f0ff"/>
                        <stop offset="100%" stop-color="#78b9e8"/>
                    </linearGradient>
                </defs>
                <circle cx="100" cy="100" r="79" fill="url(#globe-water)" stroke="#eaf7ff" stroke-width="4"/>
                <g clip-path="url(#globe-clip)" class="hero-globe-surface">
                    <path d="M36 61l15-18 24-8 14 11-7 12 12 9-4 15-17 5-8 20-13-2-7-18-17-9z
                             M98 92l16-10 17 7 3 13-10 8 6 17-11 25-11-7-5-22-12-12z
                             M132 53l15-9 19 8 5 15-13 8-14-5-10 6-8-10z
                             M29 124l16-8 17 5 7 15-15 12-19-5z
                             M153 119l14-7 12 10-5 17-17 4-9-11z"
                        fill="#397e5b"/>
                    <ellipse cx="100" cy="100" rx="35" ry="76" fill="none" stroke="#edf8ff" stroke-width="2.5" opacity="0.9"/>
                    <ellipse class="hero-globe-latitude" cx="100" cy="100" rx="76" ry="28" fill="none" stroke="#edf8ff" stroke-width="2.5" opacity="0.85"/>
                    <ellipse cx="100" cy="100" rx="76" ry="54" fill="none" stroke="#edf8ff" stroke-width="2" opacity="0.55"/>
                </g>
                <circle cx="100" cy="100" r="76" fill="none" stroke="#f4fbff" stroke-width="2"/>
            </svg>
        </header>
    <section class="cleaner-features">
        <ul>
            <li>Auto-trim Whitespace</li>
            <li>Remove Duplicate Records</li>
        </ul>
        <ul>
            <li>Outlier &amp; Anomaly Detection</li>
            <li>Validate Data Types</li>
        </ul>
    </section>
    """,
    unsafe_allow_html=True,
)

uploaded_files = st.file_uploader(
    "Upload your file",
    type=["xlsx", "csv", "xls"],
    accept_multiple_files=True,
)

if uploaded_files:
    file_signature = tuple(
        (uploaded_file.name, uploaded_file.size) for uploaded_file in uploaded_files
    )
    if st.session_state.get("_incentive_audit_files") != file_signature:
        st.session_state["_incentive_audit_files"] = file_signature
        st.session_state["_incentive_audit_stage"] = "preview"
    elif "_incentive_audit_stage" not in st.session_state:
        st.session_state["_incentive_audit_stage"] = "preview"
else:
    st.session_state["_incentive_audit_files"] = None
    st.session_state["_incentive_audit_stage"] = "upload"

progress_slot = st.empty()
_render_progress(progress_slot, st.session_state["_incentive_audit_stage"])

if uploaded_files:
    for uploaded_file in uploaded_files:
        if uploaded_file.type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
            df = pd.read_excel(uploaded_file, dtype={"Student_ID": "string"})
        elif uploaded_file.type == "text/csv":
            df = pd.read_csv(uploaded_file, dtype={"Student_ID": "string"})
        elif uploaded_file.type == "application/vnd.ms-excel":
            df = pd.read_excel(uploaded_file, dtype={"Student_ID": "string"})
        else:
            st.error("Unsupported file type. Please upload an Excel or CSV file.")
            continue

        if len(df.columns) == 1 and "," in str(df.columns[0]):
            column_names = [name.strip() for name in str(df.columns[0]).split(",")]
            split_data = df.iloc[:, 0].astype("string").str.split(",", expand=True)
            if split_data.shape[1] == len(column_names):
                split_data.columns = column_names
                df = split_data.apply(lambda column: column.str.strip())
                if "Incentive_Amount" in df.columns:
                    df["Incentive_Amount"] = pd.to_numeric(
                        df["Incentive_Amount"], errors="coerce"
                    )

        if "Incentive_Amount" in df.columns:
            amount_values = df["Incentive_Amount"].astype("string").str.replace(
                r"[$,]", "", regex=True
            ).str.strip()
            df["Incentive_Amount"] = pd.to_numeric(amount_values, errors="coerce")

        st.write(f"Data from {uploaded_file.name}:")
        st.dataframe(
            df,
            width="stretch",
            row_height=46,
            hide_index=True,
            column_config=_table_column_config(df),
        )

        if st.button(f"Run Audit for {uploaded_file.name}"):
            st.session_state["_incentive_audit_stage"] = "clean"
            _render_progress(progress_slot, "clean")
            st.markdown(
                """
                <div class="confetti-burst" aria-hidden="true">
                    <span style="--dx:-210px;--dy:-150px"></span>
                    <span style="--dx:-165px;--dy:-210px"></span>
                    <span style="--dx:-115px;--dy:-120px"></span>
                    <span style="--dx:-65px;--dy:-245px"></span>
                    <span style="--dx:-20px;--dy:-170px"></span>
                    <span style="--dx:30px;--dy:-225px"></span>
                    <span style="--dx:80px;--dy:-130px"></span>
                    <span style="--dx:125px;--dy:-205px"></span>
                    <span style="--dx:185px;--dy:-155px"></span>
                    <span style="--dx:225px;--dy:-75px"></span>
                    <span style="--dx:-235px;--dy:-65px"></span>
                    <span style="--dx:-175px;--dy:5px"></span>
                    <span style="--dx:-120px;--dy:75px"></span>
                    <span style="--dx:-60px;--dy:35px"></span>
                    <span style="--dx:5px;--dy:95px"></span>
                    <span style="--dx:65px;--dy:45px"></span>
                    <span style="--dx:125px;--dy:90px"></span>
                    <span style="--dx:195px;--dy:25px"></span>
                    <span style="--dx:245px;--dy:105px"></span>
                    <span style="--dx:35px;--dy:-80px"></span>
                    <span style="--dx:-95px;--dy:-205px"></span>
                    <span style="--dx:155px;--dy:-235px"></span>
                    <span style="--dx:-205px;--dy:115px"></span>
                    <span style="--dx:215px;--dy:-5px"></span>
                </div>
                """,
                unsafe_allow_html=True,
            )
            st.success(f"Audit completed for {uploaded_file.name}.")
            if "Student_ID" in df.columns and "Name" in df.columns:
                missing_fields = df["Student_ID"].isnull() | df["Name"].isnull()
                df.loc[missing_fields, "Audit_Status"] = "Flagged"
                df.loc[missing_fields, "Audit_Notes"] = "Missing Student_ID or Name"

            if "Student_ID" in df.columns:
                duplicates = df.duplicated(subset=["Student_ID"], keep=False)
                df.loc[duplicates, "Audit_Status"] = "Flagged"
                df.loc[duplicates, "Audit_Notes"] = "Duplicate Student_ID found"

            if "Audit_Status" in df.columns and "Audit_Notes" in df.columns:
                no_status = df["Audit_Status"].isna() | df["Audit_Status"].astype(str).str.strip().eq("")
                df.loc[no_status, "Audit_Notes"] = ""

            total_records = len(df)
            flagged_records = len(df[df["Audit_Status"] == "Flagged"]) if "Audit_Status" in df.columns else 0
            approved_records = total_records - flagged_records
            total_money = df["Incentive_Amount"].sum() if "Incentive_Amount" in df.columns else 0

            with st.container(key="audit-summary"):
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Total Records", total_records)
                col2.metric("Approved Records", approved_records)
                col3.metric("Flagged Records", flagged_records)
                col4.metric("Total Incentive Amount", f"${total_money:,.2f}")

            st.write("Cleaned Audit Report:")
            st.dataframe(
                df,
                width="stretch",
                row_height=46,
                hide_index=True,
                column_config=_table_column_config(df, cleaned=True),
            )

            output_file = f"audit_results_{uploaded_file.name.split('.')[0]}.xlsx"
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                _excel_safe_export_copy(df).to_excel(
                    writer, sheet_name="Audit Results", index=False
                )
                worksheet = writer.sheets["Audit Results"]
                table = Table(displayName="AuditResults", ref=worksheet.dimensions)
                table.tableStyleInfo = TableStyleInfo(
                    name="TableStyleMedium2",
                    showFirstColumn=False,
                    showLastColumn=False,
                    showRowStripes=True,
                    showColumnStripes=False,
                )
                worksheet.add_table(table)
                if "Incentive_Amount" in df.columns:
                    amount_column = df.columns.get_loc("Incentive_Amount") + 1
                    for row_index in range(2, len(df) + 2):
                        worksheet.cell(row=row_index, column=amount_column).number_format = '"$"#,##0.00'
                if "Audit_Status" in df.columns:
                    flagged_fill = PatternFill(fill_type="solid", fgColor="FFFFC7CE")
                    for row_index in range(2, len(df) + 2):
                        if df.iloc[row_index - 2]["Audit_Status"] == "Flagged":
                            for cell in worksheet[row_index]:
                                cell.fill = flagged_fill

            st.session_state["_incentive_audit_stage"] = "download"
            _render_progress(progress_slot, "download")
            st.download_button(
                label="Download Audit Results",
                data=excel_buffer.getvalue(),
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"download_{uploaded_file.name.split('.')[0]}",
                on_click=_mark_downloaded,
            )
