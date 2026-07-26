"""
Clinical Pharmacy Report Analyzer
==================================
A Streamlit application for analyzing clinical pharmacy data from Excel
workbooks and generating professional Word reports.

Run with: streamlit run app.py
"""

import streamlit as st
import pandas as pd
from datetime import datetime

from utils.data_loader import (
    load_excel,
    auto_detect_sheets,
    auto_map_columns,
    get_required_fields,
    normalize_consultant_names,
)
from utils.analyzer import (
    analyze_interventions,
    interventions_by_ward,
    interventions_by_consultant,
    analyze_errors,
    errors_by_ward,
    analyze_ham_lasa,
    analyze_adrs,
)
from utils.charts import (
    acceptance_pie_chart,
    ward_interventions_chart,
    consultant_chart,
    error_type_chart,
    errors_by_ward_chart,
    ham_lasa_drug_chart,
    ham_lasa_type_chart,
    adr_summary_chart,
)
from utils.report_generator import generate_report


# ── Page Config ───────────────────────────────────────────────────────
st.set_page_config(
    page_title="Clinical Pharmacy Analyzer",
    page_icon="💊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Import Google Font */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    /* Global font */
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    /* Main header styling */
    .main-header {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
        padding: 2rem 2.5rem;
        border-radius: 16px;
        margin-bottom: 2rem;
        border: 1px solid rgba(79, 143, 234, 0.2);
        box-shadow: 0 8px 32px rgba(0, 0, 0, 0.3);
    }
    .main-header h1 {
        color: #ffffff;
        font-size: 2rem;
        font-weight: 700;
        margin: 0;
        letter-spacing: -0.02em;
    }
    .main-header p {
        color: #8899aa;
        font-size: 1rem;
        margin: 0.5rem 0 0 0;
    }

    /* KPI Metric Cards */
    .kpi-card {
        background: linear-gradient(135deg, #1e2530 0%, #2a3442 100%);
        border: 1px solid rgba(79, 143, 234, 0.15);
        border-radius: 12px;
        padding: 1.5rem;
        text-align: center;
        transition: all 0.3s ease;
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.2);
    }
    .kpi-card:hover {
        border-color: rgba(79, 143, 234, 0.4);
        transform: translateY(-2px);
        box-shadow: 0 8px 24px rgba(79, 143, 234, 0.15);
    }
    .kpi-value {
        font-size: 2.5rem;
        font-weight: 700;
        margin: 0;
        line-height: 1.2;
    }
    .kpi-label {
        font-size: 0.85rem;
        color: #8899aa;
        margin: 0.3rem 0 0 0;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        font-weight: 500;
    }

    /* Section headers */
    .section-header {
        background: linear-gradient(90deg, rgba(79, 143, 234, 0.1), transparent);
        border-left: 4px solid #4F8FEA;
        padding: 0.8rem 1.2rem;
        border-radius: 0 8px 8px 0;
        margin: 2rem 0 1rem 0;
    }
    .section-header h2 {
        color: #ffffff;
        font-size: 1.3rem;
        font-weight: 600;
        margin: 0;
    }

    /* Chart description */
    .chart-description {
        background: rgba(79, 143, 234, 0.08);
        border-radius: 8px;
        padding: 1rem 1.2rem;
        margin: 0.8rem 0;
        font-size: 0.92rem;
        color: #aabbcc;
        line-height: 1.6;
        border: 1px solid rgba(79, 143, 234, 0.1);
    }

    /* Sidebar styling */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0E1117 0%, #1a1a2e 100%);
    }
    [data-testid="stSidebar"] .stMarkdown h1,
    [data-testid="stSidebar"] .stMarkdown h2,
    [data-testid="stSidebar"] .stMarkdown h3 {
        color: #4F8FEA;
    }

    /* Upload area */
    [data-testid="stFileUploader"] {
        border-radius: 12px;
    }

    /* Success/warning/error boxes */
    .success-box {
        background: rgba(46, 204, 113, 0.1);
        border: 1px solid rgba(46, 204, 113, 0.3);
        border-radius: 8px;
        padding: 1rem;
        color: #2ECC71;
    }
    .warning-box {
        background: rgba(243, 156, 18, 0.1);
        border: 1px solid rgba(243, 156, 18, 0.3);
        border-radius: 8px;
        padding: 1rem;
        color: #F39C12;
    }

    /* Hide default Streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    /* Tabs styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px 8px 0 0;
        padding: 10px 20px;
    }
</style>
""", unsafe_allow_html=True)


# ── Helper: Render KPI cards ─────────────────────────────────────────
def render_kpi(label: str, value, color: str = "#4F8FEA"):
    """Render a single KPI metric card."""
    st.markdown(f"""
    <div class="kpi-card">
        <p class="kpi-value" style="color: {color};">{value}</p>
        <p class="kpi-label">{label}</p>
    </div>
    """, unsafe_allow_html=True)


def render_section_header(title: str):
    """Render a section header with accent bar."""
    st.markdown(f"""
    <div class="section-header">
        <h2>{title}</h2>
    </div>
    """, unsafe_allow_html=True)


def render_description(text: str):
    """Render a chart description box."""
    st.markdown(f'<div class="chart-description">{text}</div>', unsafe_allow_html=True)


# ── Main App Header ──────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>💊 Clinical Pharmacy Report Analyzer</h1>
    <p>Upload your monthly Excel report → Analyze → Generate Word Document</p>
</div>
""", unsafe_allow_html=True)


# ── Sidebar ──────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📁 Upload & Settings")

    uploaded_file = st.file_uploader(
        "Upload Excel Report",
        type=["xlsx", "xls"],
        help="Upload the clinical pharmacy monthly report Excel file.",
    )

    st.markdown("---")
    st.markdown("### 🏥 Report Settings")

    hospital_name = st.text_input(
        "Hospital Name",
        value="General Hospital",
        help="This will appear on the Word report title page.",
    )

    # Month/year for the report
    current_date = datetime.now()
    months = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December",
    ]
    default_month = current_date.month - 2 if current_date.month > 1 else 11
    report_month = st.selectbox("Report Month", months, index=default_month)
    report_year = st.number_input("Report Year", min_value=2020, max_value=2030, value=current_date.year)
    month_year = f"{report_month} {report_year}"

    st.markdown("---")
    st.markdown("### ⚙️ Advanced")
    fuzzy_threshold = st.slider(
        "Consultant Name Match Threshold",
        min_value=50, max_value=100, value=85,
        help="Higher = stricter matching (only very similar names are merged). Lower = more aggressive merging.",
    )


# ── Main Content ─────────────────────────────────────────────────────
if uploaded_file is None:
    # Landing state
    st.markdown("---")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
        #### 📊 Automated Analysis
        Upload your Excel file and get instant analysis of interventions,
        medication errors, HAM/LASA consumption, and ADRs.
        """)
    with col2:
        st.markdown("""
        #### 📈 Interactive Charts
        Explore your data through beautiful, interactive charts.
        Filter, zoom, and hover for detailed insights.
        """)
    with col3:
        st.markdown("""
        #### 📝 Word Report
        Generate a professional Word document with introduction,
        findings, discussion, and conclusion — ready to submit.
        """)

    st.info("👈 **Upload your Excel file in the sidebar to get started.**")
    st.stop()


# ── Load & Process ───────────────────────────────────────────────────
@st.cache_data
def process_file(file_bytes, fuzzy_thresh):
    """Load and process the uploaded file (cached)."""
    from io import BytesIO
    buf = BytesIO(file_bytes)

    class FakeFile:
        def read(self):
            return buf.read()

    sheets = load_excel(FakeFile())
    sheet_mapping = auto_detect_sheets(sheets)
    return sheets, sheet_mapping


file_bytes = uploaded_file.read()
uploaded_file.seek(0)

try:
    sheets, sheet_mapping = process_file(file_bytes, fuzzy_threshold)
except Exception as e:
    st.error(f"❌ Error reading file: {e}")
    st.stop()


# ── Sheet Mapping UI ─────────────────────────────────────────────────
with st.expander("🔧 Sheet & Column Mapping", expanded=False):
    st.markdown("The app auto-detected the following sheet assignments. Adjust if needed:")

    sheet_names = list(sheets.keys())
    sheet_types = ["interventions", "medication_errors", "ham_lasa", "adr"]
    sheet_labels = {
        "interventions": "📋 Interventions Sheet",
        "medication_errors": "⚠️ Medication Errors Sheet",
        "ham_lasa": "💊 HAM/LASA Sheet",
        "adr": "🔴 ADR Sheet",
    }

    cols = st.columns(2)
    final_mapping = {}
    for i, stype in enumerate(sheet_types):
        with cols[i % 2]:
            default_idx = 0
            options = ["(Not available)"] + sheet_names
            if stype in sheet_mapping:
                try:
                    default_idx = options.index(sheet_mapping[stype])
                except ValueError:
                    default_idx = 0

            selected = st.selectbox(
                sheet_labels[stype],
                options=options,
                index=default_idx,
                key=f"sheet_{stype}",
            )
            if selected != "(Not available)":
                final_mapping[stype] = selected

    # Column mapping per sheet
    st.markdown("---")
    st.markdown("**Column Mappings** (auto-detected, adjust if incorrect):")

    col_mappings = {}
    for stype, sname in final_mapping.items():
        with st.expander(f"Columns for: {sname} ({stype})"):
            df = sheets[sname]
            required = get_required_fields(stype)
            auto_mapped = auto_map_columns(df, required)
            col_options = ["(None)"] + [str(c) for c in df.columns]

            mapped = {}
            cols_ui = st.columns(2)
            for j, field in enumerate(required):
                with cols_ui[j % 2]:
                    default = 0
                    if auto_mapped.get(field):
                        try:
                            default = col_options.index(auto_mapped[field])
                        except ValueError:
                            default = 0
                    sel = st.selectbox(
                        field.replace("_", " ").title(),
                        col_options,
                        index=default,
                        key=f"col_{stype}_{field}",
                    )
                    mapped[field] = sel if sel != "(None)" else None
            col_mappings[stype] = mapped


# ── Run Analysis ─────────────────────────────────────────────────────
# Use auto-mapping if user hasn't expanded the mapping section
if not col_mappings:
    for stype, sname in final_mapping.items():
        df = sheets[sname]
        required = get_required_fields(stype)
        col_mappings[stype] = auto_map_columns(df, required)

# Initialize results
intervention_results = None
ward_df = pd.DataFrame()
consultant_df = pd.DataFrame()
error_results = None
error_ward_df = pd.DataFrame()
ham_lasa_results = None
adr_results = None
all_charts = {}

# ── Interventions Analysis ────────────────────────────────────────────
if "interventions" in final_mapping:
    df_int = sheets[final_mapping["interventions"]].copy()
    cmap = col_mappings.get("interventions", {})

    # Normalize consultant names
    cons_col = cmap.get("consultant")
    if cons_col and cons_col in df_int.columns:
        df_int[cons_col] = normalize_consultant_names(df_int[cons_col], threshold=fuzzy_threshold)

    intervention_results = analyze_interventions(df_int, cmap)
    ward_df = interventions_by_ward(df_int, cmap)
    consultant_df = interventions_by_consultant(df_int, cmap)

# ── Medication Errors Analysis ────────────────────────────────────────
if "medication_errors" in final_mapping:
    df_err = sheets[final_mapping["medication_errors"]].copy()
    cmap_err = col_mappings.get("medication_errors", {})
    error_results = analyze_errors(df_err, cmap_err)
    error_ward_df = errors_by_ward(df_err, cmap_err)

# ── HAM/LASA Analysis ────────────────────────────────────────────────
if "ham_lasa" in final_mapping:
    df_hl = sheets[final_mapping["ham_lasa"]].copy()
    cmap_hl = col_mappings.get("ham_lasa", {})
    ham_lasa_results = analyze_ham_lasa(df_hl, cmap_hl)

# ── ADR Analysis ─────────────────────────────────────────────────────
if "adr" in final_mapping:
    df_adr = sheets[final_mapping["adr"]].copy()
    cmap_adr = col_mappings.get("adr", {})
    adr_results = analyze_adrs(df_adr, cmap_adr)


# ══════════════════════════════════════════════════════════════════════
# DISPLAY RESULTS
# ══════════════════════════════════════════════════════════════════════

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📋 Interventions", "⚠️ Medication Errors", "💊 HAM/LASA",
    "🔴 ADRs", "📝 Generate Report"
])

# ── Tab 1: Interventions ─────────────────────────────────────────────
with tab1:
    render_section_header("Pharmacist Interventions Overview")

    if intervention_results:
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            render_kpi("Files Reviewed", intervention_results["total_files_reviewed"], "#4F8FEA")
        with k2:
            render_kpi("Total Interventions", intervention_results["total_interventions"], "#9B59B6")
        with k3:
            render_kpi("Accepted", intervention_results["accepted"], "#2ECC71")
        with k4:
            render_kpi("Rejected", intervention_results["rejected"], "#E74C3C")

        st.markdown("")

        # Acceptance Rate
        c1, c2 = st.columns(2)
        with c1:
            fig_pie = acceptance_pie_chart(intervention_results["accepted"], intervention_results["rejected"])
            all_charts["acceptance_pie"] = acceptance_pie_chart(
                intervention_results["accepted"], intervention_results["rejected"], for_report=True
            )
            st.plotly_chart(fig_pie, use_container_width=True)
            render_description(
                f"The acceptance rate stands at <b>{intervention_results['acceptance_rate']}%</b>. "
                f"This reflects the clinical relevance and quality of pharmacist interventions. "
                f"A high acceptance rate indicates strong collaboration between pharmacy and medical teams."
            )

        with c2:
            if not ward_df.empty:
                fig_ward = ward_interventions_chart(ward_df)
                all_charts["ward_interventions"] = ward_interventions_chart(ward_df, for_report=True)
                st.plotly_chart(fig_ward, use_container_width=True)
                top_ward = ward_df.iloc[0]
                render_description(
                    f"<b>{top_ward['Ward']}</b> had the highest number of interventions "
                    f"({top_ward['Count']}). This may correlate with higher patient acuity, "
                    f"prescription complexity, or pharmacist coverage in this ward."
                )

        # Consultant chart (full width)
        if not consultant_df.empty:
            st.markdown("")
            fig_cons = consultant_chart(consultant_df)
            all_charts["consultant"] = consultant_chart(consultant_df, for_report=True)
            st.plotly_chart(fig_cons, use_container_width=True)
            top_cons = consultant_df.iloc[0]
            render_description(
                f"<b>{top_cons['Consultant']}</b> received the most interventions ({top_cons['Count']}). "
                f"This chart helps identify prescribers who may benefit from additional pharmacy "
                f"support or educational outreach. Note: Similar consultant names have been "
                f"automatically merged using fuzzy matching (threshold: {fuzzy_threshold}%)."
            )
    else:
        st.warning("⚠️ No interventions sheet detected. Please check sheet mapping above.")

# ── Tab 2: Medication Errors ─────────────────────────────────────────
with tab2:
    render_section_header("Medication Error Analysis")

    if error_results and error_results["total_errors"] > 0:
        k1, k2, k3, k4, k5 = st.columns(5)
        with k1:
            render_kpi("Total Errors", error_results["total_errors"], "#E74C3C")
        with k2:
            render_kpi("Prescribing", error_results["prescribing"], "#4F8FEA")
        with k3:
            render_kpi("Dispensing", error_results["dispensing"], "#F39C12")
        with k4:
            render_kpi("Transcription", error_results["transcription"], "#9B59B6")
        with k5:
            render_kpi("Illegible / Abbrev.", 
                      error_results["illegible_handwriting"] + error_results["incorrect_abbreviation"],
                      "#E91E90")

        st.markdown("")

        c1, c2 = st.columns(2)
        with c1:
            fig_err = error_type_chart(error_results)
            all_charts["error_type"] = error_type_chart(error_results, for_report=True)
            st.plotly_chart(fig_err, use_container_width=True)
            render_description(
                "This chart shows the distribution of medication errors by type. "
                "Prescribing errors typically include dose, frequency, or drug selection issues. "
                "Dispensing errors relate to incorrect drug or quantity dispensed. "
                "Illegible handwriting and unapproved abbreviations represent documentation-related errors."
            )

        with c2:
            if not error_ward_df.empty:
                fig_ew = errors_by_ward_chart(error_ward_df)
                all_charts["errors_by_ward"] = errors_by_ward_chart(error_ward_df, for_report=True)
                st.plotly_chart(fig_ew, use_container_width=True)
                render_description(
                    "Ward-wise error distribution helps identify units where targeted "
                    "intervention programs or additional pharmacist support may be needed. "
                    "Wards with high error rates should be prioritized for educational sessions."
                )
    elif error_results:
        st.info("ℹ️ No medication errors were recorded in this dataset.")
    else:
        st.warning("⚠️ No medication errors sheet detected. Please check sheet mapping above.")

# ── Tab 3: HAM/LASA ─────────────────────────────────────────────────
with tab3:
    render_section_header("HAM/LASA Consumption Analysis")

    if ham_lasa_results and ham_lasa_results["total_records"] > 0:
        k1, k2 = st.columns(2)
        with k1:
            render_kpi("Patients Received HAM/LASA", ham_lasa_results["total_patients"], "#E74C3C")
        with k2:
            render_kpi("Total Consumption Records", ham_lasa_results["total_records"], "#F39C12")

        st.markdown("")

        c1, c2 = st.columns(2)
        with c1:
            if not ham_lasa_results["by_type"].empty:
                fig_hl_type = ham_lasa_type_chart(ham_lasa_results["by_type"])
                all_charts["ham_lasa_type"] = ham_lasa_type_chart(ham_lasa_results["by_type"], for_report=True)
                st.plotly_chart(fig_hl_type, use_container_width=True)
                render_description(
                    "The HAM vs LASA distribution shows the proportion of High Alert Medications "
                    "versus Look-Alike Sound-Alike drugs consumed during the month. HAM drugs "
                    "require extra safety protocols, while LASA drugs need careful labeling to "
                    "prevent mix-ups."
                )

        with c2:
            if not ham_lasa_results["by_drug"].empty:
                fig_hl_drug = ham_lasa_drug_chart(ham_lasa_results["by_drug"])
                all_charts["ham_lasa_drugs"] = ham_lasa_drug_chart(ham_lasa_results["by_drug"], for_report=True)
                st.plotly_chart(fig_hl_drug, use_container_width=True)
                top_drug = ham_lasa_results["by_drug"].iloc[0]
                render_description(
                    f"<b>{top_drug['Drug']}</b> was the most commonly consumed HAM/LASA medication "
                    f"({top_drug['Count']} instances). The pharmacy department should ensure that "
                    f"all safety protocols, including double-checking and tall-man lettering, are "
                    f"being followed for these medications."
                )
    elif ham_lasa_results:
        st.info("ℹ️ No HAM/LASA consumption records found in this dataset.")
    else:
        st.warning("⚠️ No HAM/LASA sheet detected. Please check sheet mapping above.")

# ── Tab 4: ADRs ──────────────────────────────────────────────────────
with tab4:
    render_section_header("Adverse Drug Reactions (ADRs)")

    if adr_results and adr_results["total_adrs"] > 0:
        k1, k2, k3 = st.columns(3)
        with k1:
            render_kpi("Total ADRs", adr_results["total_adrs"], "#E74C3C")
        with k2:
            render_kpi("From HAM/LASA", adr_results["adrs_from_ham_lasa"], "#F39C12")
        with k3:
            render_kpi("From Other Drugs", adr_results["adrs_other"], "#4F8FEA")

        st.markdown("")

        c1, c2 = st.columns(2)
        with c1:
            fig_adr = adr_summary_chart(adr_results["total_adrs"], adr_results["adrs_from_ham_lasa"])
            all_charts["adr_summary"] = adr_summary_chart(
                adr_results["total_adrs"], adr_results["adrs_from_ham_lasa"], for_report=True
            )
            st.plotly_chart(fig_adr, use_container_width=True)
            render_description(
                "This chart shows the proportion of ADRs caused by HAM/LASA medications "
                "versus other drugs. ADRs from high-alert medications are of particular "
                "concern and may require additional monitoring protocols."
            )

        with c2:
            if not adr_results["by_drug"].empty:
                st.markdown("##### 💊 ADRs by Causative Drug")
                st.dataframe(
                    adr_results["by_drug"].head(10),
                    use_container_width=True,
                    hide_index=True,
                )
                render_description(
                    "The table above shows the drugs most frequently associated with adverse "
                    "reactions. These drugs should be closely monitored, and prescribers should "
                    "be informed about the ADR patterns."
                )
    elif adr_results:
        st.success("✅ No ADRs were reported during this period — a positive patient safety indicator.")
    else:
        st.warning("⚠️ No ADR sheet detected. Please check sheet mapping above.")

# ── Tab 5: Generate Report ───────────────────────────────────────────
with tab5:
    render_section_header("Generate Word Report")

    st.markdown(f"""
    Generate a professional Word document report for **{month_year}** at **{hospital_name}**.

    The report will include:
    - 📄 Title page with hospital name and month
    - 📝 Introduction with context
    - 📊 Findings with all charts and tables
    - 💬 Auto-generated discussion of trends
    - ✅ Conclusion with recommendations
    """)

    st.markdown("")

    if st.button("📥 Generate Word Report", type="primary", use_container_width=True):
        with st.spinner("Generating report... This may take a moment while charts are rendered."):
            try:
                # Generate report-style (light theme) charts
                report_charts = {}
                if intervention_results:
                    report_charts["acceptance_pie"] = acceptance_pie_chart(
                        intervention_results["accepted"], intervention_results["rejected"], for_report=True
                    )
                if not ward_df.empty:
                    report_charts["ward_interventions"] = ward_interventions_chart(ward_df, for_report=True)
                if not consultant_df.empty:
                    report_charts["consultant"] = consultant_chart(consultant_df, for_report=True)
                if error_results and error_results["total_errors"] > 0:
                    report_charts["error_type"] = error_type_chart(error_results, for_report=True)
                if not error_ward_df.empty:
                    report_charts["errors_by_ward"] = errors_by_ward_chart(error_ward_df, for_report=True)
                if ham_lasa_results and not ham_lasa_results["by_type"].empty:
                    report_charts["ham_lasa_type"] = ham_lasa_type_chart(
                        ham_lasa_results["by_type"], for_report=True
                    )
                if ham_lasa_results and not ham_lasa_results["by_drug"].empty:
                    report_charts["ham_lasa_drugs"] = ham_lasa_drug_chart(
                        ham_lasa_results["by_drug"], for_report=True
                    )
                if adr_results and adr_results["total_adrs"] > 0:
                    report_charts["adr_summary"] = adr_summary_chart(
                        adr_results["total_adrs"], adr_results["adrs_from_ham_lasa"], for_report=True
                    )

                doc_buffer = generate_report(
                    month_year=month_year,
                    hospital_name=hospital_name,
                    intervention_results=intervention_results,
                    ward_df=ward_df if not ward_df.empty else None,
                    consultant_df=consultant_df if not consultant_df.empty else None,
                    error_results=error_results,
                    error_ward_df=error_ward_df if not error_ward_df.empty else None,
                    ham_lasa_results=ham_lasa_results,
                    adr_results=adr_results,
                    charts=report_charts,
                )

                st.success("✅ Report generated successfully!")
                filename = f"Clinical_Pharmacy_Report_{report_month}_{report_year}.docx"
                st.download_button(
                    label=f"⬇️ Download {filename}",
                    data=doc_buffer,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    use_container_width=True,
                )

            except Exception as e:
                st.error(f"❌ Error generating report: {e}")
                st.exception(e)


# ── Footer ───────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    '<p style="text-align: center; color: #556677; font-size: 0.8rem;">'
    'Clinical Pharmacy Report Analyzer • Built with Streamlit & Plotly'
    '</p>',
    unsafe_allow_html=True,
)
