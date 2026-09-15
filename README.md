# Clinical Analytics & Polypharmacy Surveillance Platform 🩺

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io/)

A modern, interactive clinical intelligence and polypharmacy surveillance dashboard built with Streamlit, Plotly, SciPy, and Pandas. Designed for healthcare researchers, epidemiologists, and clinicians to analyze drug burden, comorbid trajectories, and diagnostic testing patterns across a de-identified cohort of 100 sample patients.

---

## 🌟 Key Features

1. **📊 Executive Overview & Demographics**:
   - Dynamic cohort filtering by medical specialty, polypharmacy status ($\ge 5$ medications), age range, and gender.
   - Interactive age distribution histograms with KDE and polypharmacy stratification.
   - Demographic donut charts and specialty caseload analysis.
   - Geographic distribution of patient home states.

2. **💊 Polypharmacy & Drug Burden Analytics**:
   - Correlation scatter plot of Age vs. Concurrent Medication Count with OLS trendline, $R^2$, and polypharmacy risk threshold ($\ge 5$).
   - Medication count distribution by specialty with box plots and interquartile ranges.
   - Frequency ranking of the top prescribed active pharmacotherapies.
   - Identification of severe drug burden patients ($\ge 6$ concurrent medications).

3. **🩺 Clinical Conditions & Diagnostic Workup**:
   - Prevalence spectrum of chronic comorbidities (Hypertension, Type 2 Diabetes, CKD, CAD, Heart Failure, COPD, etc.).
   - Order frequency of diagnostic imaging (MRI, CT, Echocardiogram, Ultrasound, DEXA) and laboratory assays (A1c, CMP, BMP, Troponin, etc.).

4. **📈 Biostatistical Hypothesis Testing**:
   - **Two-Sample Welch's t-test**: Quantifies whether there is a statistically significant difference in mean age between polypharmacy and non-polypharmacy cohorts.
   - **Correlation Analysis**: Pearson ($r$) and Spearman ($\rho$) rank correlation testing between age and medication count with exact $p$-values.
   - **Descriptive Statistics Matrix**: Stratified by polypharmacy status (Mean, Std Dev, Median, IQR, Min, Max).

5. **📋 Patient Registry & Individual Clinical Dossier**:
   - Full tabular patient registry with interactive sorting, search, and progress bars.
   - Individual Patient Inspector displaying complete de-identified clinical dossier:
     - Demographics, Contact Phone & Portal Email, Residential Address
     - Active Prescription Regimens with dosages and frequencies
     - Past Medical History (PMHx) & Presenting Chief Complaints
     - Recommended Diagnostic Scans & Ordered Labs
     - Attending Physician Consultation Assessment & Plan Notes
   - Filtered dataset export to CSV.

---

## 🔒 De-Identification Compliance (HIPAA Safe Harbor)
All 100 patient profiles in `sample_patients_100.xlsx` adhere strictly to de-identification guidelines:
- **No Personal Names**: Unique alphanumeric identifiers (`PT-10001` through `PT-10100`).
- **Synthetic Contact Info**: Portal email handles (`pt.pt-10001@patientportal.health`) and standard test phone numbers.
- **Provider Confidentiality**: Physician notes reference findings and specialties without provider names.

---

## 🚀 Local Quickstart

### 1. Clone Repository & Install Dependencies
```bash
git clone https://github.com/kousthubsarma/chappal.git
cd chappal
pip install -r requirements.txt
```

### 2. Run Streamlit Application
```bash
streamlit run app.py
```
The app will be available locally at `http://localhost:8501`.

---

## ☁️ Deploy to Streamlit Community Cloud

1. Visit [Streamlit Community Cloud](https://share.streamlit.io/).
2. Sign in with your GitHub account.
3. Click **"New app"**.
4. Select repository: `kousthubsarma/chappal` (or your chosen repo name).
5. Branch: `main`
6. Main file path: `app.py`
7. Click **"Deploy!"**
