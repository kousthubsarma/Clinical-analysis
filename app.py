import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from scipy import stats
import os
import re

# -----------------------------------------------------------------------------
# Page Configuration & Aesthetics
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Clinical Analytics & Polypharmacy Surveillance",
    page_icon="🩺",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for modern, healthcare enterprise aesthetics
st.markdown("""
<style>
    /* Global style adjustments */
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        padding-left: 2.5rem;
        padding-right: 2.5rem;
    }
    
    /* Headers & Typography */
    h1, h2, h3 {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        color: #0F172A;
    }
    
    /* Top Banner */
    .hero-banner {
        background: linear-gradient(135deg, #1E3A8A 0%, #0F172A 100%);
        color: white;
        padding: 24px 30px;
        border-radius: 14px;
        margin-bottom: 25px;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.15);
    }
    .hero-banner h1 {
        color: #FFFFFF;
        font-size: 2.0rem;
        font-weight: 700;
        margin-bottom: 6px;
    }
    .hero-banner p {
        color: #94A3B8;
        font-size: 0.95rem;
        margin-bottom: 0;
    }

    /* Metric Cards */
    .metric-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 16px 20px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
        transition: transform 0.15s ease, box-shadow 0.15s ease;
    }
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 12px -2px rgba(0,0,0,0.05);
    }
    .metric-label {
        font-size: 0.78rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: #64748B;
        margin-bottom: 4px;
    }
    .metric-value {
        font-size: 1.7rem;
        font-weight: 700;
        color: #0F172A;
    }
    .metric-sub {
        font-size: 0.8rem;
        color: #059669;
        font-weight: 500;
        margin-top: 2px;
    }
    .metric-sub.warning {
        color: #DC2626;
    }
    
    /* Polypharmacy Badges */
    .badge-poly {
        background-color: #FEE2E2;
        color: #991B1B;
        padding: 3px 10px;
        border-radius: 9999px;
        font-size: 0.75rem;
        font-weight: 700;
        display: inline-block;
    }
    .badge-standard {
        background-color: #DCFCE7;
        color: #166534;
        padding: 3px 10px;
        border-radius: 9999px;
        font-size: 0.75rem;
        font-weight: 700;
        display: inline-block;
    }
    
    /* Patient Dossier card */
    .dossier-card {
        background: #F8FAFC;
        border: 1px solid #CBD5E1;
        border-radius: 12px;
        padding: 20px;
        margin-top: 15px;
    }
</style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# Data Loading & Caching
# -----------------------------------------------------------------------------
@st.cache_data
def load_data():
    excel_path = os.path.join(os.path.dirname(__file__), "sample_patients_100.xlsx")
    if not os.path.exists(excel_path):
        st.error(f"Excel file not found at {excel_path}")
        return pd.DataFrame()
    
    df = pd.read_excel(excel_path, sheet_name="Patient Records")
    
    # Standardize column types
    df["Age"] = pd.to_numeric(df["Age"], errors="coerce")
    df["Med Count"] = pd.to_numeric(df["Med Count"], errors="coerce")
    df["Date of Visit"] = pd.to_datetime(df["Date of Visit"])
    df["Is_Polypharmacy"] = df["Polypharmacy?"].apply(lambda x: 1 if str(x).strip().lower() == "yes" else 0)
    
    # Extract City & State from Address
    def parse_city_state(addr):
        parts = [p.strip() for p in str(addr).split(",")]
        if len(parts) >= 3:
            city = parts[-2]
            state_zip = parts[-1].strip().split()
            state = state_zip[0] if state_zip else "Unknown"
            return city, state
        return "Unknown", "Unknown"
    
    df[["City", "State"]] = df["Residential Address"].apply(lambda x: pd.Series(parse_city_state(x)))
    return df

df_raw = load_data()

if df_raw.empty:
    st.stop()

# -----------------------------------------------------------------------------
# Sidebar Controls & Global Filters
# -----------------------------------------------------------------------------
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/medical-history.png", width=64)
    st.markdown("### **Clinical Filters**")
    st.markdown("Refine patient cohort for real-time statistical surveillance.")
    
    # Global search
    search_query = st.text_input("🔍 Global Keyword Search", "", placeholder="e.g. Asthma, MRI, PT-10042, Lisinopril")
    
    # Specialist multi-select
    specialties = sorted(df_raw["Specialist Seen"].dropna().unique().tolist())
    selected_specialties = st.multiselect(
        "Clinical Specialty",
        options=specialties,
        default=specialties,
        help="Select one or multiple medical specialties"
    )
    
    # Polypharmacy filter
    poly_filter = st.radio(
        "Polypharmacy Status",
        options=["All Patients", "Polypharmacy (≥ 5 Meds)", "Standard (< 5 Meds)"],
        index=0
    )
    
    # Age range slider
    min_age = int(df_raw["Age"].min())
    max_age = int(df_raw["Age"].max())
    age_range = st.slider("Age Range (Years)", min_value=min_age, max_value=max_age, value=(min_age, max_age))
    
    # Gender filter
    all_genders = sorted(df_raw["Gender"].dropna().unique().tolist())
    selected_genders = st.multiselect("Gender", options=all_genders, default=all_genders)
    
    # Medication Count Slider
    min_meds = int(df_raw["Med Count"].min())
    max_meds = int(df_raw["Med Count"].max())
    med_range = st.slider("Medication Count Range", min_value=min_meds, max_value=max_meds, value=(min_meds, max_meds))

    st.markdown("---")
    st.markdown("💡 *Clinical benchmark:* **Polypharmacy** is defined as routine concurrent prescription of $\ge 5$ pharmaceuticals.")
    if st.button("🔄 Reset All Filters", use_container_width=True):
        st.rerun()

# -----------------------------------------------------------------------------
# Filter Execution
# -----------------------------------------------------------------------------
df = df_raw.copy()

if selected_specialties:
    df = df[df["Specialist Seen"].isin(selected_specialties)]
else:
    df = df.iloc[0:0]

if poly_filter == "Polypharmacy (≥ 5 Meds)":
    df = df[df["Polypharmacy?"] == "Yes"]
elif poly_filter == "Standard (< 5 Meds)":
    df = df[df["Polypharmacy?"] == "No"]

df = df[(df["Age"] >= age_range[0]) & (df["Age"] <= age_range[1])]
df = df[df["Gender"].isin(selected_genders)]
df = df[(df["Med Count"] >= med_range[0]) & (df["Med Count"] <= med_range[1])]

if search_query.strip():
    q = search_query.strip().lower()
    match_mask = (
        df["Patient ID"].str.lower().str.contains(q, na=False) |
        df["Current Medications"].str.lower().str.contains(q, na=False) |
        df["Past Medical History (PMHx)"].str.lower().str.contains(q, na=False) |
        df["Presenting Symptoms"].str.lower().str.contains(q, na=False) |
        df["Scans & Tests Suggested"].str.lower().str.contains(q, na=False) |
        df["Doctor's Clinical Notes"].str.lower().str.contains(q, na=False) |
        df["Specialist Seen"].str.lower().str.contains(q, na=False)
    )
    df = df[match_mask]

# -----------------------------------------------------------------------------
# Top Hero Banner
# -----------------------------------------------------------------------------
st.markdown("""
<div class="hero-banner">
    <h1>Clinical Analytics & Polypharmacy Surveillance Platform</h1>
    <p>Comprehensive Statistical Analysis & Decision Dashboard • N=100 De-Identified Patient Cohort</p>
</div>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# Key Metric Row (Executive KPIs)
# -----------------------------------------------------------------------------
col1, col2, col3, col4, col5 = st.columns(5)

cohort_n = len(df)
total_raw = len(df_raw)
poly_n = (df["Polypharmacy?"] == "Yes").sum() if cohort_n > 0 else 0
poly_pct = (poly_n / cohort_n * 100) if cohort_n > 0 else 0
mean_age = df["Age"].mean() if cohort_n > 0 else 0
median_age = df["Age"].median() if cohort_n > 0 else 0
mean_meds = df["Med Count"].mean() if cohort_n > 0 else 0
high_risk_n = (df["Med Count"] >= 6).sum() if cohort_n > 0 else 0

with col1:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">Active Cohort</div>
        <div class="metric-value">{cohort_n}</div>
        <div class="metric-sub">{cohort_n/total_raw*100:.0f}% of total ({total_raw})</div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    badge_class = "warning" if poly_pct >= 50 else ""
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">Polypharmacy Rate</div>
        <div class="metric-value">{poly_pct:.1f}%</div>
        <div class="metric-sub {badge_class}">{poly_n} of {cohort_n} patients</div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">Average Patient Age</div>
        <div class="metric-value">{mean_age:.1f} <span style="font-size: 1rem; color: #64748B;">yrs</span></div>
        <div class="metric-sub">Median: {median_age:.0f} yrs</div>
    </div>
    """, unsafe_allow_html=True)

with col4:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">Avg Medication Burden</div>
        <div class="metric-value">{mean_meds:.2f}</div>
        <div class="metric-sub">Meds / Patient</div>
    </div>
    """, unsafe_allow_html=True)

with col5:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">Severe Burden (≥6 Meds)</div>
        <div class="metric-value">{high_risk_n}</div>
        <div class="metric-sub warning">{high_risk_n/cohort_n*100 if cohort_n else 0:.1f}% elevated risk</div>
    </div>
    """, unsafe_allow_html=True)

st.write("")

if cohort_n == 0:
    st.warning("⚠️ No patients match the current filter selection. Please broaden your filters in the left sidebar.")
    st.stop()

# -----------------------------------------------------------------------------
# Main Analysis Tabs
# -----------------------------------------------------------------------------
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Executive Overview & Demographics",
    "💊 Polypharmacy & Drug Burden",
    "🩺 Clinical Conditions & Diagnostics",
    "📈 Statistical Hypothesis Testing",
    "📋 Patient Registry & Clinical Dossier"
])

# =============================================================================
# TAB 1: EXECUTIVE OVERVIEW & DEMOGRAPHICS
# =============================================================================
with tab1:
    st.markdown("### Demographic & Operational Surveillance")
    r1_col1, r1_col2 = st.columns([3, 2])
    
    with r1_col1:
        # Age distribution histogram with Polypharmacy hue
        fig_age = px.histogram(
            df,
            x="Age",
            color="Polypharmacy?",
            color_discrete_map={"Yes": "#EF4444", "No": "#10B981"},
            nbins=15,
            barmode="overlay",
            opacity=0.75,
            title="Patient Age Distribution Stratified by Polypharmacy Status",
            labels={"Age": "Patient Age (Years)", "count": "Patient Count"}
        )
        fig_age.update_layout(
            font_family="Inter, sans-serif",
            legend_title_text="Polypharmacy?",
            plot_bgcolor="rgba(248, 250, 252, 0.5)",
            paper_bgcolor="rgba(0,0,0,0)",
            margin=dict(l=20, r=20, t=40, b=20)
        )
        st.plotly_chart(fig_age, use_container_width=True)
        
    with r1_col2:
        # Gender Donut Chart
        gender_counts = df["Gender"].value_counts().reset_index()
        gender_counts.columns = ["Gender", "Count"]
        fig_gender = px.pie(
            gender_counts,
            values="Count",
            names="Gender",
            hole=0.55,
            color="Gender",
            color_discrete_map={"Female": "#EC4899", "Male": "#3B82F6", "Non-Binary": "#8B5CF6"},
            title="Cohort Gender Representation"
        )
        fig_gender.update_traces(textposition='inside', textinfo='percent+label')
        fig_gender.update_layout(
            font_family="Inter, sans-serif",
            showlegend=False,
            margin=dict(l=20, r=20, t=40, b=20)
        )
        st.plotly_chart(fig_gender, use_container_width=True)
        
    r2_col1, r2_col2 = st.columns([3, 2])
    
    with r2_col1:
        # Specialty Case Load Bar Chart
        spec_df = df["Specialist Seen"].value_counts().reset_index()
        spec_df.columns = ["Specialist", "Count"]
        fig_spec = px.bar(
            spec_df,
            x="Count",
            y="Specialist",
            orientation="h",
            color="Count",
            color_continuous_scale="Blues",
            title="Patient Caseload by Medical Specialty",
            text="Count"
        )
        fig_spec.update_layout(
            yaxis={'categoryorder': 'total ascending'},
            font_family="Inter, sans-serif",
            plot_bgcolor="rgba(248, 250, 252, 0.5)",
            paper_bgcolor="rgba(0,0,0,0)",
            margin=dict(l=20, r=20, t=40, b=20),
            coloraxis_showscale=False
        )
        fig_spec.update_traces(textposition='outside')
        st.plotly_chart(fig_spec, use_container_width=True)
        
    with r2_col2:
        # Geographic distribution (Top States)
        state_counts = df["State"].value_counts().head(8).reset_index()
        state_counts.columns = ["State", "Patients"]
        fig_geo = px.bar(
            state_counts,
            x="State",
            y="Patients",
            title="Top Geographic Clusters (Patient Home States)",
            color="Patients",
            color_continuous_scale="Tealgrn"
        )
        fig_geo.update_layout(
            font_family="Inter, sans-serif",
            plot_bgcolor="rgba(248, 250, 252, 0.5)",
            paper_bgcolor="rgba(0,0,0,0)",
            coloraxis_showscale=False,
            margin=dict(l=20, r=20, t=40, b=20)
        )
        st.plotly_chart(fig_geo, use_container_width=True)

# =============================================================================
# TAB 2: POLYPHARMACY & DRUG BURDEN
# =============================================================================
with tab2:
    st.markdown("### Pharmacotherapeutic Burden & Risk Profiling")
    
    p_col1, p_col2 = st.columns(2)
    
    with p_col1:
        # Scatter: Age vs Med Count with OLS Trendline
        fig_scatter = px.scatter(
            df,
            x="Age",
            y="Med Count",
            color="Polypharmacy?",
            color_discrete_map={"Yes": "#DC2626", "No": "#059669"},
            trendline="ols",
            hover_data=["Patient ID", "Specialist Seen"],
            title="Correlation: Age vs Concurrent Prescription Count",
            labels={"Age": "Age (Years)", "Med Count": "Total Concurrent Medications"}
        )
        fig_scatter.add_hline(y=4.5, line_dash="dash", line_color="#E11D48", annotation_text="Polypharmacy Threshold (≥5)", annotation_position="bottom right")
        fig_scatter.update_layout(
            font_family="Inter, sans-serif",
            plot_bgcolor="rgba(248, 250, 252, 0.5)",
            paper_bgcolor="rgba(0,0,0,0)",
            margin=dict(l=20, r=20, t=40, b=20)
        )
        st.plotly_chart(fig_scatter, use_container_width=True)
        
    with p_col2:
        # Boxplot: Med Count by Specialty
        fig_box = px.box(
            df,
            x="Specialist Seen",
            y="Med Count",
            color="Specialist Seen",
            title="Medication Count Distribution by Clinical Specialty",
            labels={"Med Count": "Prescription Count", "Specialist Seen": "Specialty"}
        )
        fig_box.update_layout(
            font_family="Inter, sans-serif",
            xaxis_tickangle=-45,
            showlegend=False,
            plot_bgcolor="rgba(248, 250, 252, 0.5)",
            paper_bgcolor="rgba(0,0,0,0)",
            margin=dict(l=20, r=20, t=40, b=20)
        )
        st.plotly_chart(fig_box, use_container_width=True)
        
    st.markdown("#### Most Frequently Prescribed Active Pharmacotherapies")
    
    # Extract Individual Medications from list
    all_meds = []
    for m_str in df["Current Medications"].dropna():
        lines = str(m_str).split("\n")
        for line in lines:
            cleaned = line.replace("•", "").strip()
            if cleaned:
                # Extract primary drug name (first 1-2 words before dosage)
                match = re.match(r"^([A-Za-z\-]+(?:\s+[A-Za-z\-]+)?)", cleaned)
                if match:
                    drug_candidate = match.group(1).title()
                    # Filter out common dosage words if caught
                    if drug_candidate.lower() not in ["daily", "inhalation", "puff", "delayed-release"]:
                        all_meds.append(drug_candidate)

    if all_meds:
        med_counts = pd.Series(all_meds).value_counts().head(12).reset_index()
        med_counts.columns = ["Medication", "Prescription Frequency"]
        
        fig_drugs = px.bar(
            med_counts,
            x="Prescription Frequency",
            y="Medication",
            orientation="h",
            color="Prescription Frequency",
            color_continuous_scale="Viridis",
            text="Prescription Frequency",
            title="Top 12 Most Frequently Prescribed Medications Across Filtered Cohort"
        )
        fig_drugs.update_layout(
            yaxis={'categoryorder': 'total ascending'},
            font_family="Inter, sans-serif",
            plot_bgcolor="rgba(248, 250, 252, 0.5)",
            paper_bgcolor="rgba(0,0,0,0)",
            coloraxis_showscale=False,
            margin=dict(l=20, r=20, t=40, b=20)
        )
        fig_drugs.update_traces(textposition='outside')
        st.plotly_chart(fig_drugs, use_container_width=True)

# =============================================================================
# TAB 3: CLINICAL CONDITIONS & DIAGNOSTICS
# =============================================================================
with tab3:
    st.markdown("### Clinical Comorbidity & Diagnostic Workup Spectrum")
    
    c_col1, c_col2 = st.columns(2)
    
    with c_col1:
        # Prevalent Medical Conditions
        conditions = [
            "Hypertension", "Diabetes", "Neuropathy", "Dyslipidemia", "Heart Failure",
            "Atrial Fibrillation", "COPD", "Asthma", "Parkinson", "Migraine",
            "Rheumatoid Arthritis", "Lupus", "Kidney Disease", "Colitis", "Cancer",
            "Osteoarthritis", "Osteoporosis", "Psoriasis", "Thrombocytopenia"
        ]
        
        cond_counts = []
        for cond in conditions:
            c = df["Past Medical History (PMHx)"].str.contains(cond, case=False, na=False).sum()
            cond_counts.append({"Condition": cond, "Prevalence": c})
            
        cond_df = pd.DataFrame(cond_counts).sort_values(by="Prevalence", ascending=True)
        cond_df = cond_df[cond_df["Prevalence"] > 0]
        
        fig_cond = px.bar(
            cond_df,
            x="Prevalence",
            y="Condition",
            orientation="h",
            title="Prevalence of Key Chronic Conditions in Active Cohort",
            color="Prevalence",
            color_continuous_scale="Purples",
            text="Prevalence"
        )
        fig_cond.update_layout(
            font_family="Inter, sans-serif",
            plot_bgcolor="rgba(248, 250, 252, 0.5)",
            paper_bgcolor="rgba(0,0,0,0)",
            coloraxis_showscale=False,
            margin=dict(l=20, r=20, t=40, b=20)
        )
        fig_cond.update_traces(textposition='outside')
        st.plotly_chart(fig_cond, use_container_width=True)
        
    with c_col2:
        # Diagnostic Modalities Suggested
        test_modalities = [
            "Echocardiogram", "ECG", "MRI", "CT", "X-Ray", "Ultrasound",
            "HbA1c", "Lipid Panel", "PFT", "Troponin", "Creatinine", "CMP", "CBC",
            "Biopsy", "Colonoscopy", "DEXA", "ESR"
        ]
        test_counts = []
        for test in test_modalities:
            c = df["Scans & Tests Suggested"].str.contains(test, case=False, na=False).sum()
            test_counts.append({"Diagnostic Test": test, "Orders Suggested": c})
            
        test_df = pd.DataFrame(test_counts).sort_values(by="Orders Suggested", ascending=True)
        test_df = test_df[test_df["Orders Suggested"] > 0]
        
        fig_tests = px.bar(
            test_df,
            x="Orders Suggested",
            y="Diagnostic Test",
            orientation="h",
            title="Most Frequently Ordered Scans, Labs & Diagnostic Tests",
            color="Orders Suggested",
            color_continuous_scale="Teal",
            text="Orders Suggested"
        )
        fig_tests.update_layout(
            font_family="Inter, sans-serif",
            plot_bgcolor="rgba(248, 250, 252, 0.5)",
            paper_bgcolor="rgba(0,0,0,0)",
            coloraxis_showscale=False,
            margin=dict(l=20, r=20, t=40, b=20)
        )
        fig_tests.update_traces(textposition='outside')
        st.plotly_chart(fig_tests, use_container_width=True)

# =============================================================================
# TAB 4: STATISTICAL HYPOTHESIS TESTING
# =============================================================================
with tab4:
    st.markdown("### Formal Statistical Inference & Correlation Matrix")
    st.markdown("Automated biostatistical tests evaluating clinical differences and associations within the cohort.")
    
    stat1, stat2 = st.columns(2)
    
    # 1. Welch's Two-Sample t-test: Age vs Polypharmacy
    poly_ages = df[df["Polypharmacy?"] == "Yes"]["Age"]
    non_poly_ages = df[df["Polypharmacy?"] == "No"]["Age"]
    
    with stat1:
        st.markdown("#### 1. Welch's t-test: Age vs Polypharmacy Status")
        if len(poly_ages) >= 2 and len(non_poly_ages) >= 2:
            t_stat, p_val = stats.ttest_ind(poly_ages, non_poly_ages, equal_var=False)
            diff = poly_ages.mean() - non_poly_ages.mean()
            sig_text = "Statistically Significant ($p < 0.05$)" if p_val < 0.05 else "Not Statistically Significant ($p \ge 0.05$)"
            
            st.info(f"""
            **Hypothesis:** $H_0$: There is no difference in mean age between polypharmacy and non-polypharmacy cohorts.  
            - **Polypharmacy Cohort Mean Age:** `{poly_ages.mean():.1f}` years ($N={len(poly_ages)}$)  
            - **Non-Polypharmacy Cohort Mean Age:** `{non_poly_ages.mean():.1f}` years ($N={len(non_poly_ages)}$)  
            - **Mean Difference:** `{diff:+.1f}` years  
            - **Test Statistic ($t$):** `{t_stat:.3f}`  
            - **Two-Tailed $p$-value:** `{p_val:.4f}`  
            - **Conclusion:** **{sig_text}**
            """)
        else:
            st.warning("Insufficient samples in one of the cohorts to compute two-sample t-test.")
            
    with stat2:
        st.markdown("#### 2. Correlation: Age vs Medication Count")
        if len(df) >= 3:
            pearson_r, pearson_p = stats.pearsonr(df["Age"], df["Med Count"])
            spearman_rho, spearman_p = stats.spearmanr(df["Age"], df["Med Count"])
            
            st.info(f"""
            **Association Testing:**  
            - **Pearson Correlation ($r$):** `{pearson_r:.3f}` ($p = {pearson_p:.4f}$)  
            - **Spearman Rank Correlation ($\\rho$):** `{spearman_rho:.3f}` ($p = {spearman_p:.4f}$)  
            - **Coefficient of Determination ($R^2$):** `{pearson_r**2:.3f}`  
            - **Interpretation:** {"Positive linear tendency between patient age and total prescribed regimens." if pearson_r > 0 else "Weak or negative correlation in active selection."}
            """)
        else:
            st.warning("Insufficient observations for correlation testing.")

    st.markdown("---")
    st.markdown("#### 3. Summary Statistics by Polypharmacy Status")
    
    summary_cols = ["Age", "Med Count"]
    summary_table = df.groupby("Polypharmacy?")[summary_cols].agg([
        ("Count", "count"),
        ("Mean", "mean"),
        ("Std Dev", "std"),
        ("Median", "median"),
        ("IQR", lambda x: np.percentile(x, 75) - np.percentile(x, 25)),
        ("Min", "min"),
        ("Max", "max")
    ]).round(2)
    
    st.dataframe(summary_table, use_container_width=True)

# =============================================================================
# TAB 5: PATIENT REGISTRY & CLINICAL DOSSIER
# =============================================================================
with tab5:
    st.markdown("### Interactive Patient Registry & Detailed Clinical Dossier")
    
    # Table of current filtered cohort
    display_df = df[[
        "Patient ID", "Age", "Gender", "Specialist Seen", "Date of Visit",
        "Polypharmacy?", "Med Count", "City", "State"
    ]].copy()
    
    # Styled table
    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Patient ID": st.column_config.TextColumn("Patient ID", width="small"),
            "Age": st.column_config.NumberColumn("Age", format="%d yrs"),
            "Polypharmacy?": st.column_config.TextColumn("Polypharmacy"),
            "Med Count": st.column_config.ProgressColumn("Med Burden", min_value=1, max_value=8, format="%d meds"),
            "Date of Visit": st.column_config.DateColumn("Visit Date", format="YYYY-MM-DD")
        }
    )
    
    # Download button for filtered dataset
    csv_bytes = df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download Filtered Cohort Data (CSV)",
        data=csv_bytes,
        file_name="filtered_patient_cohort.csv",
        mime="text/csv",
        use_container_width=False
    )
    
    st.markdown("---")
    st.markdown("### 🔍 Individual Patient Clinical Dossier")
    st.markdown("Select an individual patient ID to view their de-identified comprehensive clinical dossier.")
    
    patient_ids = df["Patient ID"].tolist()
    if patient_ids:
        selected_pt_id = st.selectbox("Select Patient Identifier:", options=patient_ids, index=0)
        pt_data = df[df["Patient ID"] == selected_pt_id].iloc[0]
        
        # Dossier card
        d1, d2, d3 = st.columns([1, 1, 1.3])
        
        with d1:
            st.markdown(f"#### 👤 Demographics & Contact")
            st.markdown(f"**Patient ID:** `{pt_data['Patient ID']}`")
            st.markdown(f"**Age:** {pt_data['Age']} years")
            st.markdown(f"**Gender:** {pt_data['Gender']}")
            st.markdown(f"**Contact Phone:** `{pt_data['Contact Phone']}`")
            st.markdown(f"**Portal Email:** `{pt_data['Contact Email']}`")
            st.markdown(f"**Residential Address:**  \n{pt_data['Residential Address']}")
            
        with d2:
            st.markdown(f"#### 💊 Pharmacotherapy & Specialty")
            poly_tag = f"<span class='badge-poly'>POLYPHARMACY (≥5)</span>" if pt_data["Polypharmacy?"] == "Yes" else "<span class='badge-standard'>STANDARD REGIMEN</span>"
            st.markdown(f"**Status:** {poly_tag}", unsafe_allow_html=True)
            st.markdown(f"**Specialist Seen:** {pt_data['Specialist Seen']}")
            st.markdown(f"**Date of Visit:** {pt_data['Date of Visit'].strftime('%Y-%m-%d')}")
            st.markdown(f"**Active Regimens ({pt_data['Med Count']}):**")
            st.text(pt_data["Current Medications"])
            
        with d3:
            st.markdown(f"#### 🩺 Clinical Presentation & Plan")
            st.markdown(f"**Past Medical History (PMHx):**  \n{pt_data['Past Medical History (PMHx)']}")
            st.markdown(f"**Presenting Chief Complaints:**  \n*{pt_data['Presenting Symptoms']}*")
            st.markdown(f"**Scans & Diagnostic Tests Suggested:**  \n`{pt_data['Scans & Tests Suggested']}`")
            st.markdown(f"**Attending Physician Notes:**")
            st.info(pt_data["Doctor's Clinical Notes"])
    else:
        st.info("No patient matching current filters.")

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #94A3B8; font-size: 0.8rem;">
    Clinical Analytics & Polypharmacy Surveillance Platform • Synthetic De-Identified Dataset • Built with Streamlit & Plotly
</div>
""", unsafe_allow_html=True)
