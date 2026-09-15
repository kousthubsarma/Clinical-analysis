#!/usr/bin/env python3
"""
Generate a synthetic dataset of 100 sample patients in an Excel (.xlsx) file.
Meets all requirements:
- Fields: ID, Contact (Phone, Email), Address, Age, Gender, Polypharmacy status & count,
  Medication List, Specialist, Date of Visit, Doctor's Notes, Symptoms, Scans/Tests Suggested, Medical History.
- NO NAMES (strict adherence: IDs only, neutral email handles, no provider names).
- High visual aesthetics: custom typography, color palette, zebra striping, conditional badges, auto-fit widths, and a KPI summary sheet.
"""

import random
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Set seed for consistent high-quality generation
random.seed(42)

# Clinical specialty profiles with coherent medical conditions, medications, symptoms, tests, and doctor notes
SPECIALTY_TEMPLATES = [
    {
        "specialist": "Cardiologist",
        "weight": 14,
        "scenarios": [
            {
                "pmhx": "Essential Hypertension, Hyperlipidemia, Coronary Artery Disease (s/p DES to LAD 2021)",
                "symptoms": "Exertional retrosternal chest tightness, dyspnea on stair climbing, mild bilateral ankle swelling",
                "meds": ["Atorvastatin 80mg daily", "Metoprolol Succinate 50mg daily", "Aspirin 81mg daily", "Lisinopril 20mg daily", "Nitroglycerin 0.4mg SL PRN"],
                "tests": "12-Lead Electrocardiogram (ECG), Transthoracic Echocardiogram (TTE), High-Sensitivity Troponin, Lipid Panel",
                "notes": "Patient evaluated for recurrent exertional chest tightness over last 3 weeks. Resting vitals: BP 136/84 mmHg, HR 68 bpm. Cardiac auscultation shows regular rhythm without murmurs or gallops. Trace pretibial edema noted. Adjusted beta-blocker dosage and ordered outpatient TTE to assess LV wall motion. Advised strict adherence to low-sodium diet and return in 6 weeks."
            },
            {
                "pmhx": "Heart Failure with Reduced Ejection Fraction (HFrEF, EF 35%), Non-ischemic Cardiomyopathy, Chronic Kidney Disease Stage 3a",
                "symptoms": "Orthopnea (2 pillows), paroxysmal nocturnal dyspnea, 4 lb weight gain over 5 days, fatigue",
                "meds": ["Sacubitril/Valsartan 24/26mg BID", "Carvedilol 12.5mg BID", "Spironolactone 25mg daily", "Dapagliflozin 10mg daily", "Furosemide 40mg BID"],
                "tests": "NT-proBNP, Serum Electrolytes & Creatinine, Chest X-Ray (AP/Lateral), 2D Echocardiogram",
                "notes": "Acute decompensated heart failure evaluation. JVD noted at 7 cm above sternal angle; bilateral basilar crackles on lung exam; 2+ pitting bilateral lower extremity edema. Up-titrated oral loop diuretic for 7 days with daily weight monitoring protocol. Recheck renal panel in 10 days."
            },
            {
                "pmhx": "Paroxysmal Atrial Fibrillation (CHA2DS2-VASc = 3), Hypertension, Mild Mitral Regurgitation",
                "symptoms": "Sudden onset rapid irregular heart flutter, lightheadedness during episodes, decreased exercise tolerance",
                "meds": ["Apixaban 5mg BID", "Diltiazem CD 180mg daily", "Losartan 50mg daily"],
                "tests": "14-day Holter / Zio Patch Cardiac Monitor, Transthoracic Echocardiogram, TSH & Free T4, Serum Magnesium",
                "notes": "Follow-up for episodic palpitations occurring twice weekly lasting 15-45 minutes. In-clinic ECG confirms normal sinus rhythm at 72 bpm. Anticoagulation therapy confirmed compliant. Ordered 14-day continuous ambulatory cardiac patch to quantify AF burden. Discussed prospective catheter ablation if symptoms remain refractory to rate control."
            }
        ]
    },
    {
        "specialist": "Endocrinologist",
        "weight": 13,
        "scenarios": [
            {
                "pmhx": "Type 2 Diabetes Mellitus (HbA1c 8.8%), Diabetic Peripheral Neuropathy, Dyslipidemia, Obesity (Class II)",
                "symptoms": "Polyuria, polydipsia, nocturnal bilateral burning foot pain, persistent daytime fatigue",
                "meds": ["Metformin 1000mg BID", "Semaglutide 1.0mg SQ weekly", "Empagliflozin 25mg daily", "Gabapentin 300mg TID", "Atorvastatin 40mg daily", "Glimepiride 2mg daily"],
                "tests": "Hemoglobin A1c, Urine Albumin-to-Creatinine Ratio (uACR), Comprehensive Metabolic Panel (CMP), Bilateral Lower Extremity Monofilament Sensory Exam",
                "notes": "Suboptimal glycemic control review. Glucometer logs demonstrate morning fasting blood glucose between 175-210 mg/dL. Neurological exam reveals reduced pinprick and vibratory sensation in stocking distribution up to mid-calf. Reinforced diabetic meal planning; transitioned to weekly GLP-1 receptor agonist titration; podiatry consult requested."
            },
            {
                "pmhx": "Primary Hypothyroidism (Hashimoto's Thyroiditis), Osteopenia, Pernicious Anemia",
                "symptoms": "Chronic fatigue, unexplained cold intolerance, dry coarse skin, constipation, sluggish cognition",
                "meds": ["Levothyroxine 112mcg daily on empty stomach", "Cyanocobalamin 1000mcg IM monthly", "Calcium Carbonate 500mg + Vitamin D3 daily"],
                "tests": "Thyroid Stimulating Hormone (TSH), Free T4, Anti-TPO Antibodies, Complete Blood Count (CBC), DEXA Bone Density Scan",
                "notes": "Follow-up for persistent hypothyroid symptoms despite current dosage. Recent lab tests indicate TSH elevated at 6.8 mIU/L. Reviewed proper administration protocol (avoiding calcium and iron supplements within 4 hours of levothyroxine). Increased levothyroxine to 125mcg daily with repeat TSH scheduled in 8 weeks."
            },
            {
                "pmhx": "Primary Hyperparathyroidism, Recurrent Nephrolithiasis, Osteoporosis (T-score -2.8)",
                "symptoms": "Mild diffuse musculoskeletal aches, bone pain in forearm and lumbar spine, episodic flank discomfort",
                "meds": ["Alendronate 70mg weekly", "Cholecalciferol 2000 IU daily", "Hydrochlorothiazide 25mg daily"],
                "tests": "Serum Calcium (ionized & total), Intact PTH, 24-hour Urine Calcium & Creatinine, Sestamibi Parathyroid SPECT/CT Scan",
                "notes": "Evaluation for persistent hypercalcemia (serum calcium 10.9 mg/dL) with elevated intact PTH (88 pg/mL). Sestamibi scan ordered to localize suspected solitary parathyroid adenoma. Patient instructed to maintain adequate hydration (>2.5L water daily) and avoid prolonged immobility. Endocrine surgery referral initiated."
            }
        ]
    },
    {
        "specialist": "Pulmonologist",
        "weight": 10,
        "scenarios": [
            {
                "pmhx": "Chronic Obstructive Pulmonary Disease (COPD, GOLD Group E), Chronic Bronchitis, Remote Tobacco Use (35 pack-years)",
                "symptoms": "Productive morning cough with mucoid sputum, progressive dyspnea on walking one block, audible expiratory wheeze",
                "meds": ["Fluticasone Furoate / Umeclidinium / Vilanterol (Trelegy Ellipta) 1 inhalation daily", "Albuterol HFA Inhaler 2 puffs q4h PRN", "Roflumilast 500mcg daily", "Montelukast 10mg daily", "Azithromycin 250mg 3x/week"],
                "tests": "Complete Pulmonary Function Test (PFT / Spirometry), High-Resolution Chest CT, Pulse Oximetry (Rest & Exertion), Arterial Blood Gas",
                "notes": "Follow-up for frequent moderate COPD exacerbations over the past 12 months. SpO2 92% on ambient air, dropping to 88% upon 6-minute walk test. Auscultation demonstrates distant breath sounds and scattered expiratory wheezes throughout lung fields. Enrolled patient in structured pulmonary rehabilitation; prescribed nocturnal supplemental oxygen (2 L/min via nasal cannula)."
            },
            {
                "pmhx": "Severe Persistent Asthma (Eosinophilic Phenotype), Allergic Rhinitis, Gastroesophageal Reflux Disease",
                "symptoms": "Nocturnal chest tightness, recurrent cough triggered by cold air, frequent reliance on rescue albuterol (>4x/week)",
                "meds": ["Budesonide/Formoterol 160/4.5mcg 2 puffs BID", "Benralizumab 30mg SQ q8w", "Fluticasone Propionate Nasal Spray 50mcg daily", "Cetirizine 10mg daily", "Omeprazole 20mg daily"],
                "tests": "Fractional Exhaled Nitric Oxide (FeNO), Absolute Eosinophil Count, Total Serum IgE, Spirometry with Bronchodilator Reversibility",
                "notes": "Assessment of eosinophilic asthma response to biologic agent. Pre-treatment eosinophil count was 620 cells/uL, down to 50 cells/uL on current regimen. Patient reports 60% reduction in rescue inhaler use over past 2 months. Inhaler technique verified with spacer. Sustained maintenance therapy approved."
            }
        ]
    },
    {
        "specialist": "Neurologist",
        "weight": 11,
        "scenarios": [
            {
                "pmhx": "Parkinson's Disease (Hoehn & Yahr Stage II), Essential Tremor, Orthostatic Hypotension",
                "symptoms": "Asymmetric resting tremor in right upper extremity, progressive bradykinesia, micrographia, morning stiffness",
                "meds": ["Carbidopa/Levodopa 25/100mg 1 tab TID", "Pramipexole 0.5mg TID", "Rasagiline 1mg daily", "Midodrine 5mg TID PRN", "Polyethylene Glycol 17g daily"],
                "tests": "DaTscan (Ioflupane I-123 SPECT), UPDRS Motor Assessment, Orthostatic Vital Signs, Formal Neuropsychological Evaluation",
                "notes": "Parkinson's disease motor symptom review. Observed right-sided resting 'pill-rolling' tremor and cogwheel rigidity at right wrist. Gait reveals reduced right arm swing with preserved postural stability. Adjusted carbidopa/levodopa timing to 30 minutes before meals to optimize bioavailability. Referred to physical therapy for LSVT BIG protocol."
            },
            {
                "pmhx": "Chronic Migraine with Aura, Cervicogenic Cephalea, Insomnia",
                "symptoms": "Hemicranial throbbing headache 12-14 days/month, scintillating scotoma aura, severe nausea and photophobia",
                "meds": ["Erenumab 140mg SQ monthly", "Rimegepant 75mg PRN at onset", "Topiramate 50mg BID", "Melatonin 5mg qHS"],
                "tests": "Brain MRI with and without IV Contrast, MR Angiography of Head and Neck, Headache Severity Diary Review",
                "notes": "Evaluation of refractory chronic migraine. Neurologic physical exam normal without focal deficits. Cranial nerves II-XII intact bilaterally. Brain MRI completed 2 months ago revealed no acute intracranial abnormalities or vascular malformations. Initiated CGRP monoclonal antibody therapy; recommended continued trigger diary and sleep hygiene maintenance."
            }
        ]
    },
    {
        "specialist": "Rheumatologist",
        "weight": 9,
        "scenarios": [
            {
                "pmhx": "Rheumatoid Arthritis (Seropositive, Anti-CCP >250 U/mL), Secondary Sjogren's Syndrome, Osteopenia",
                "symptoms": "Bilateral wrist and MCP joint swelling, morning joint stiffness lasting >90 minutes, severe xerostomia and dry eyes",
                "meds": ["Methotrexate 15mg weekly", "Folic Acid 1mg daily", "Adalimumab 40mg SQ bi-weekly", "Prednisone 5mg daily taper", "Hydroxychloroquine 200mg daily", "Pilocarpine 5mg TID"],
                "tests": "Erythrocyte Sedimentation Rate (ESR), High-Sensitivity C-Reactive Protein (hs-CRP), Bilateral Hand and Wrist Radiographs, Complete Blood Count, LFTs",
                "notes": "Patient reports moderate disease flare following recent viral infection. Exam reveals active synovitis with boggy tenderness in bilateral 2nd and 3rd MCP joints and right wrist. Methotrexate hepatic safety labs within acceptable limits. Plan: temporary low-dose prednisone pulse taper while maintaining TNF-inhibitor biologic."
            },
            {
                "pmhx": "Systemic Lupus Erythematosus (SLE), Lupus Nephritis (Class II, quiescent), Raynaud's Phenomenon",
                "symptoms": "Malar facial erythema flare, episodic triphasic color changes in digits exposed to cold, persistent joint arthralgias",
                "meds": ["Hydroxychloroquine 300mg daily", "Mycophenolate Mofetil 1000mg BID", "Amlodipine 5mg daily", "Belimumab 200mg SQ weekly", "Vitamin D3 2000 IU daily"],
                "tests": "Anti-dsDNA Antibody Titer, Serum Complement C3/C4, Spot Urine Protein-to-Creatinine Ratio, Baseline Retinal Exam / OCT",
                "notes": "Routine lupus monitoring. No active peripheral edema or active urinary sediment. Serum C3 and C4 remain within normal range, anti-dsDNA stable at 1:40. Annual dilated retinal exam and visual field testing confirmed no hydroxychloroquine maculopathy. Advised strict ultraviolet photoprotection and thermal gloves for Raynaud's."
            }
        ]
    },
    {
        "specialist": "Nephrologist",
        "weight": 8,
        "scenarios": [
            {
                "pmhx": "Chronic Kidney Disease Stage 4 (eGFR 22 mL/min/1.73m2), Diabetic Nephropathy, Hypertensive Arteriolosclerosis, Secondary Hyperparathyroidism",
                "symptoms": "Generalized pruritus, metallic taste, chronic low-grade fatigue, bilateral lower extremity edema (1+)",
                "meds": ["Losartan 50mg daily", "Torsemide 20mg daily", "Calcium Acetate 667mg TID with meals", "Sodium Bicarbonate 650mg BID", "Calcitriol 0.25mcg daily", "Darbepoetin Alfa 40mcg SQ monthly"],
                "tests": "Serum Creatinine & eGFR, Serum Potassium & Bicarbonate, Intact PTH, 24-Hour Urine Protein, Bilateral Renal Ultrasound",
                "notes": "CKD Stage 4 progression surveillance. Serum creatinine stable at 2.8 mg/dL with eGFR 22. Potassium 4.9 mEq/L. Initiated counseling regarding prospective renal replacement modalities including peritoneal dialysis, hemodialysis, and pre-emptive living donor kidney transplant evaluation. Low-potassium and low-phosphorus dietary guidance reinforced."
            }
        ]
    },
    {
        "specialist": "Gastroenterologist",
        "weight": 8,
        "scenarios": [
            {
                "pmhx": "Ulcerative Colitis (Pancolitis, Moderate-to-Severe), Iron Deficiency Anemia, Primary Sclerosing Cholangitis (early)",
                "symptoms": "Bloody diarrhea (6-8 loose stools per day), nocturnal fecal urgency, crampy left lower quadrant abdominal pain, 6 lb weight loss",
                "meds": ["Ustekinumab 90mg SQ q8w", "Mesalamine 1.2g delayed-release 4 tabs daily", "Ferrous Sulfate 325mg daily", "Budesonide Multi-Matrix 9mg daily taper"],
                "tests": "Diagnostic Colonoscopy with Targeted Mucosal Biopsies, Fecal Calprotectin, Serum Iron Profile & Ferritin, C-Reactive Protein",
                "notes": "Presents for evaluation of active colitis flare. Abdomen soft with localized tenderness in left iliac fossa; bowel sounds hyperactive. Fecal calprotectin significantly elevated at 840 mcg/g. Scheduled outpatient colonoscopy to assess Mayo endoscopic score and evaluate for biologic therapeutic drug monitoring."
            },
            {
                "pmhx": "Non-Alcoholic Steatohepatitis (NASH / MASH with Stage 3 Fibrosis), Type 2 Diabetes, Gastroesophageal Reflux Disease (Barrett's Esophagus)",
                "symptoms": "Right upper quadrant dull abdominal fullness, postprandial retrosternal burning pyrosis, chronic daytime lethargy",
                "meds": ["Resmetirom 80mg daily", "Pioglitazone 30mg daily", "Pantoprazole 40mg daily before breakfast", "Atorvastatin 20mg daily"],
                "tests": "Hepatic FibroScan (Transient Elastography), Abdominal Ultrasound with Doppler, Liver Function Panel (AST/ALT/ALP/Bilirubin), Surveillance EGD",
                "notes": "MASH fibrosis follow-up. ALT 64 U/L, AST 52 U/L. Recent FibroScan showed liver stiffness of 10.4 kPa (consistent with advanced F3 fibrosis). Initiated newly approved thyroid hormone receptor-beta agonist resmetirom. Upper endoscopy scheduled for surveillance of known short-segment non-dysplastic Barrett's esophagus."
            }
        ]
    },
    {
        "specialist": "Medical Oncologist",
        "weight": 7,
        "scenarios": [
            {
                "pmhx": "Invasive Ductal Carcinoma of Breast (Stage IIB, HR+/HER2-), s/p Lumpectomy and Adjuvant Radiation (2023), Osteopenia",
                "symptoms": "Bilateral joint stiffness and myalgias attributed to aromatase inhibitor, hot flashes, mild treatment-related fatigue",
                "meds": ["Letrozole 2.5mg daily", "Zoledronic Acid 4mg IV q6m", "Calcium Carbonate 600mg + D3 BID", "Venlafaxine ER 37.5mg daily for vasomotor symptoms"],
                "tests": "Annual Diagnostic Mammogram, Dual-Energy X-Ray Absorptiometry (DEXA), Comprehensive Metabolic Panel, CA 15-3 and CEA Biomarkers",
                "notes": "Routine post-treatment oncologic surveillance at 18 months post-surgery. Physical examination of bilateral breasts and regional nodal basins shows well-healed surgical site with no palpable masses, retraction, or lymphadenopathy. Aromatase inhibitor-associated arthralgias managed symptomatically. DEXA scan scheduled to monitor bone mineral density."
            }
        ]
    },
    {
        "specialist": "Psychiatrist",
        "weight": 7,
        "scenarios": [
            {
                "pmhx": "Bipolar I Disorder (Most Recent Episode Depressed, Moderate), Generalized Anxiety Disorder, Metabolic Syndrome",
                "symptoms": "Psychomotor slowing, persistent anhedonia, early morning awakening insomnia, pervasive cognitive fog",
                "meds": ["Lithium Carbonate 600mg BID", "Lurasidone 40mg daily with evening meal", "Lamotrigine 150mg daily", "Clonazepam 0.5mg PRN severe panic"],
                "tests": "Serum Lithium Level (12-hour trough), Comprehensive Metabolic Panel (BUN/Creatinine), Thyroid Panel (TSH), Fasting Lipid Profile & Glucose",
                "notes": "Psychiatric pharmacotherapy follow-up. Current serum lithium trough level at 0.72 mEq/L (therapeutic target 0.6-0.8). PHQ-9 depression score 15, GAD-7 anxiety score 11. Denies active suicidal ideation or psychotic phenomena. Sleep hygiene reinforced; cognitive behavioral therapy engagement validated."
            }
        ]
    },
    {
        "specialist": "Orthopedic Surgeon",
        "weight": 6,
        "scenarios": [
            {
                "pmhx": "Severe Primary Osteoarthritis of Right Knee (Kellgren-Lawrence Grade IV), Lumbar Spondylosis, Hypertension",
                "symptoms": "Severe right knee joint pain on weight bearing, mechanical catching, joint crepitus, inability to ambulate >150 yards",
                "meds": ["Meloxicam 15mg daily", "Acetaminophen 1000mg TID", "Tramadol 50mg PRN severe breakthrough pain", "Omeprazole 20mg daily", "Lisinopril 10mg daily"],
                "tests": "Standing Weight-Bearing Bilateral Knee Radiographs (AP, Lateral, Merchant views), Pre-operative ECG, Basic Coagulation Panel (PT/INR, aPTT)",
                "notes": "Pre-operative evaluation for elective right Total Knee Arthroplasty (TKA). Conservative management failed over 18 months including physical therapy, unloader bracing, and intra-articular corticosteroid and hyaluronic acid injections. Radiographs reveal tricompartmental joint space collapse with prominent subchondral sclerosis. Surgical consent obtained."
            }
        ]
    },
    {
        "specialist": "Geriatrician",
        "weight": 5,
        "scenarios": [
            {
                "pmhx": "Major Neurocognitive Disorder (Vascular / Alzheimer's Mixed, Moderate), Frailty Syndrome, Osteoporosis, Urge Incontinence, Chronic Constipation",
                "symptoms": "Short-term memory loss, episodic nighttime wandering, two mechanical falls in the past 6 months without fracture, poor appetite",
                "meds": ["Donepezil 10mg qHS", "Memantine 10mg BID", "Alendronate 70mg weekly", "Mirabegron 25mg daily", "Docusate Sodium 100mg BID", "Cholecalciferol 2000 IU daily", "Polyethylene Glycol 17g daily"],
                "tests": "Montreal Cognitive Assessment (MoCA), Timed Up and Go (TUG) Mobility Test, Vitamin B12 and Folate Levels, Post-Void Residual Bladder Ultrasound",
                "notes": "Comprehensive geriatric assessment conducted with family caregiver present. MoCA score 17/30. TUG test 16 seconds indicating elevated fall risk. Completed formal deprescribing medication review; eliminated sedating antihistamines and anticholinergics to decrease delirium risk. Home physical therapy safety evaluation ordered."
            }
        ]
    },
    {
        "specialist": "Dermatologist",
        "weight": 5,
        "scenarios": [
            {
                "pmhx": "Severe Plaque Psoriasis (PASI 14.2), Psoriatic Arthritis, Dyslipidemia",
                "symptoms": "Erythematous, indurated plaques with silvery mica-like scales covering bilateral extensor elbows, knees, and scalp; intense pruritus",
                "meds": ["Ixekizumab 80mg SQ monthly", "Clobetasol Propionate 0.05% Topical Foam PRN", "Calcipotriene 0.005% Topical Ointment daily", "Cetirizine 10mg daily"],
                "tests": "Psoriasis Area and Severity Index (PASI), QuantiFERON-TB Gold Screening, Hepatitis B & C Serology Panel, Complete Blood Count",
                "notes": "Follow-up for biologic therapy monitoring in severe plaque psoriasis. Body surface area involvement improved from 16% to 3.5% after 12 weeks of IL-17 antagonist therapy. No active mucocutaneous candidiasis or secondary skin infection observed. Patient reports dramatic improvement in pruritus and quality of life."
            }
        ]
    },
    {
        "specialist": "Hematologist",
        "weight": 5,
        "scenarios": [
            {
                "pmhx": "Chronic Immune Thrombocytopenic Purpura (ITP), Antiphospholipid Syndrome (Triple Positive), History of DVT",
                "symptoms": "Spontaneous petechiae on lower extremities, easy mucosal ecchymosis, intermittent epistaxis",
                "meds": ["Eltrombopag 50mg daily on empty stomach", "Warfarin 5mg daily (target INR 2.5-3.5)", "Folic Acid 1mg daily", "Omeprazole 20mg daily"],
                "tests": "Complete Blood Count with Peripheral Blood Smear, PT/INR, Lupus Anticoagulant & Anti-Cardiolipin Antibodies, Reticulocyte Count",
                "notes": "Monitoring for immune thrombocytopenia. Current platelet count responded to thrombopoietin receptor agonist, rising from 18,000/uL to 62,000/uL. Therapeutic INR maintained at 2.8. No gross active bleeding or hematuria. Counselled on food-drug interactions regarding dietary calcium and eltrombopag administration."
            }
        ]
    }
]

# Cities, States, and ZIP codes across the US
LOCATIONS = [
    ("Seattle", "WA", "98101"),
    ("Spokane", "WA", "99201"),
    ("Portland", "OR", "97201"),
    ("San Francisco", "CA", "94102"),
    ("Los Angeles", "CA", "90012"),
    ("San Diego", "CA", "92101"),
    ("Phoenix", "AZ", "85001"),
    ("Tucson", "AZ", "85701"),
    ("Denver", "CO", "80202"),
    ("Austin", "TX", "78701"),
    ("Dallas", "TX", "75201"),
    ("Houston", "TX", "77002"),
    ("San Antonio", "TX", "78205"),
    ("Minneapolis", "MN", "55401"),
    ("Chicago", "IL", "60601"),
    ("Indianapolis", "IN", "46204"),
    ("Columbus", "OH", "43215"),
    ("Cleveland", "OH", "44114"),
    ("Detroit", "MI", "48226"),
    ("Atlanta", "GA", "30303"),
    ("Miami", "FL", "33101"),
    ("Orlando", "FL", "32801"),
    ("Tampa", "FL", "33602"),
    ("Nashville", "TN", "37201"),
    ("Charlotte", "NC", "28202"),
    ("Raleigh", "NC", "27601"),
    ("Richmond", "VA", "23219"),
    ("Philadelphia", "PA", "19102"),
    ("Pittsburgh", "PA", "15222"),
    ("New York", "NY", "10001"),
    ("Buffalo", "NY", "14201"),
    ("Boston", "MA", "02108"),
    ("Baltimore", "MD", "21201"),
    ("Saint Louis", "MO", "63101"),
    ("Kansas City", "MO", "64106"),
    ("Salt Lake City", "UT", "84101")
]

STREET_NAMES = [
    "Maple Crest Way", "Highland View Dr", "Cedar Ridge Rd", "Oakwood Blvd", "Pine Valley Ln",
    "Meadowbrook Court", "Summit Ridge Ave", "Sycamore Grove Ter", "Willow Creek Path", "Brookside Parkway",
    "Riverdale Crossing", "Evergreen Terrace", "Chestnut Hill Dr", "Fox Run Court", "Stonegate Way",
    "Fairview Heights", "Lakeside Promenade", "Beacon Point Rd", "Sunset Ridge Dr", "Canyon Creek Ln",
    "Magnolia Garden Way", "Timberline Pass", "Bridlewood Circle", "Heron Lake Blvd", "Crestview Terrace"
]

AREA_CODES = [
    "206", "503", "415", "213", "619", "602", "303", "512", "214", "713",
    "612", "312", "317", "614", "216", "313", "404", "305", "407", "813",
    "615", "704", "919", "804", "215", "412", "212", "716", "617", "410"
]

def generate_patient_records(n=100):
    records = []
    
    # Pre-select specialty scenarios based on weights
    specialty_pool = []
    for spec in SPECIALTY_TEMPLATES:
        for _ in range(spec["weight"]):
            specialty_pool.append(spec)
    
    # Dates within the last 6 months (late 2025 to early 2026)
    base_date = datetime(2026, 3, 15)
    
    for i in range(1, n + 1):
        pt_id = f"PT-{10000 + i}"
        
        # Pick specialty and scenario
        spec_info = random.choice(specialty_pool)
        specialist = spec_info["specialist"]
        scenario = random.choice(spec_info["scenarios"])
        
        # Age distribution tailored realistically to specialty
        if specialist == "Geriatrician":
            age = random.randint(75, 94)
        elif specialist == "Medical Oncologist":
            age = random.randint(45, 82)
        elif specialist == "Psychiatrist":
            age = random.randint(21, 64)
        elif specialist == "Rheumatologist":
            age = random.randint(34, 76)
        elif specialist == "Orthopedic Surgeon":
            age = random.randint(52, 85)
        elif specialist == "Cardiologist":
            age = random.randint(48, 86)
        elif specialist == "Nephrologist":
            age = random.randint(50, 84)
        else:
            age = random.randint(24, 79)
            
        # Gender
        gender_weights = [48, 48, 4]
        gender = random.choices(["Female", "Male", "Non-Binary"], weights=gender_weights)[0]
        
        # Address (no names)
        city, state, zip_code = random.choice(LOCATIONS)
        street_num = random.randint(101, 9899)
        street_name = random.choice(STREET_NAMES)
        unit = f", Apt {random.randint(101, 808)}" if random.random() > 0.4 else ""
        full_address = f"{street_num} {street_name}{unit}, {city}, {state} {zip_code}"
        
        # Contact (NO personal names in email or phone)
        area_code = random.choice(AREA_CODES)
        phone = f"({area_code}) 555-{random.randint(1000, 9999)}"
        email_domains = ["patientportal.health", "carechart.org", "myclinicrecord.net", "medsecure.org"]
        email = f"pt.{pt_id.lower()}@{random.choice(email_domains)}"
        
        # Medications & Polypharmacy
        meds_list = list(scenario["meds"])
        
        # Introduce slight natural variation in medication count
        # In medicine, polypharmacy is formally defined as taking >= 5 concurrent medications
        if random.random() < 0.25 and len(meds_list) > 3:
            # Drop 1 med occasionally
            meds_list = meds_list[:-1]
        elif random.random() < 0.35:
            # Add an OTC or standard supplement
            supplements = [
                "Colecalciferol (Vitamin D3) 2000 IU daily",
                "Omega-3 Acid Ethyl Esters 1g daily",
                "Calcium Carbonate 500mg daily",
                "Acetaminophen 500mg PRN mild discomfort",
                "Multivitamin Silver 1 tab daily"
            ]
            cand = random.choice(supplements)
            if cand not in meds_list:
                meds_list.append(cand)
                
        med_count = len(meds_list)
        is_polypharmacy = "Yes" if med_count >= 5 else "No"
        polypharmacy_display = f"Yes ({med_count} meds)" if med_count >= 5 else f"No ({med_count} meds)"
        formatted_meds = "\n".join([f"• {m}" for m in meds_list])
        
        # Visit date (past 180 days)
        days_ago = random.randint(1, 180)
        visit_date = (base_date - timedelta(days=days_ago)).strftime("%Y-%m-%d")
        
        # Clinical fields from scenario
        pmhx = scenario["pmhx"]
        symptoms = scenario["symptoms"]
        tests = scenario["tests"]
        notes = scenario["notes"]
        
        records.append({
            "patient_id": pt_id,
            "age": age,
            "gender": gender,
            "contact_phone": phone,
            "contact_email": email,
            "full_address": full_address,
            "specialist": specialist,
            "date_of_visit": visit_date,
            "polypharmacy_status": is_polypharmacy,
            "med_count": med_count,
            "medication_list": formatted_meds,
            "medical_history": pmhx,
            "symptoms": symptoms,
            "tests_suggested": tests,
            "doctor_notes": notes
        })
        
    return records

def create_excel_workbook(records, output_filename="sample_patients_100.xlsx"):
    wb = openpyxl.Workbook()
    
    # ---------------------------------------------------------
    # Sheet 1: Patients Directory
    # ---------------------------------------------------------
    ws = wb.active
    ws.title = "Patient Records"
    ws.views.sheetView[0].showGridLines = True
    
    # Columns definition
    headers = [
        ("Patient ID", 14, "center"),
        ("Age", 8, "center"),
        ("Gender", 12, "center"),
        ("Contact Phone", 16, "center"),
        ("Contact Email", 27, "left"),
        ("Residential Address", 34, "left"),
        ("Specialist Seen", 18, "left"),
        ("Date of Visit", 14, "center"),
        ("Polypharmacy?", 15, "center"),
        ("Med Count", 12, "center"),
        ("Current Medications", 42, "left"),
        ("Past Medical History (PMHx)", 36, "left"),
        ("Presenting Symptoms", 36, "left"),
        ("Scans & Tests Suggested", 34, "left"),
        ("Doctor's Clinical Notes", 54, "left"),
    ]
    
    # Styling Palette
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid") # Navy Blue
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    data_font = Font(name="Calibri", size=10, color="0F172A")
    id_font = Font(name="Calibri", size=10, bold=True, color="1E3A8A")
    
    # Zebra striping
    row_fill_even = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid") # Soft slate
    row_fill_odd = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    
    # Polypharmacy badge fills
    poly_yes_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid") # Soft red/coral
    poly_yes_font = Font(name="Calibri", size=10, bold=True, color="991B1B")
    poly_no_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid") # Soft green
    poly_no_font = Font(name="Calibri", size=10, bold=True, color="166534")
    
    thin_border_color = "E2E8F0"
    border_side = Side(border_style="thin", color=thin_border_color)
    cell_border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)
    
    # Write Headers
    ws.row_dimensions[1].height = 28
    for col_idx, (col_title, col_width, alignment_h) in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_title)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = cell_border
        ws.column_dimensions[get_column_letter(col_idx)].width = col_width

    # Write Data Rows
    for row_idx, rec in enumerate(records, 2):
        ws.row_dimensions[row_idx].height = 68 # Comfortable height for wrapped text
        is_even = (row_idx % 2 == 0)
        base_fill = row_fill_even if is_even else row_fill_odd
        
        row_values = [
            rec["patient_id"],
            rec["age"],
            rec["gender"],
            rec["contact_phone"],
            rec["contact_email"],
            rec["full_address"],
            rec["specialist"],
            rec["date_of_visit"],
            rec["polypharmacy_status"],
            rec["med_count"],
            rec["medication_list"],
            rec["medical_history"],
            rec["symptoms"],
            rec["tests_suggested"],
            rec["doctor_notes"]
        ]
        
        for col_idx, val in enumerate(row_values, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            col_info = headers[col_idx - 1]
            align_h = col_info[2]
            
            # Wrap text for multiline columns
            wrap = col_idx in [10, 11, 12, 13, 14, 15]
            cell.alignment = Alignment(horizontal=align_h, vertical="top" if wrap else "center", wrap_text=wrap)
            cell.font = id_font if col_idx == 1 else data_font
            cell.border = cell_border
            cell.fill = base_fill
            
            # Highlight Polypharmacy column
            if col_idx == 9:
                if val == "Yes":
                    cell.fill = poly_yes_fill
                    cell.font = poly_yes_font
                else:
                    cell.fill = poly_no_fill
                    cell.font = poly_no_font
                    
    # Freeze Header Row
    ws.freeze_panes = "A2"
    
    # Add Auto-Filter
    last_col_letter = get_column_letter(len(headers))
    ws.auto_filter.ref = f"A1:{last_col_letter}{len(records) + 1}"

    # ---------------------------------------------------------
    # Sheet 2: Executive Cohort Summary & Statistics
    # ---------------------------------------------------------
    ws_sum = wb.create_sheet(title="Cohort Overview")
    ws_sum.views.sheetView[0].showGridLines = True
    
    # Title Block
    ws_sum.column_dimensions["A"].width = 4
    ws_sum.column_dimensions["B"].width = 30
    ws_sum.column_dimensions["C"].width = 18
    ws_sum.column_dimensions["D"].width = 8
    ws_sum.column_dimensions["E"].width = 32
    ws_sum.column_dimensions["F"].width = 18
    
    ws_sum.row_dimensions[2].height = 30
    title_cell = ws_sum.cell(row=2, column=2, value="Cohort Summary & Polypharmacy Analysis")
    title_cell.font = Font(name="Calibri", size=16, bold=True, color="1E3A8A")
    
    sub_cell = ws_sum.cell(row=3, column=2, value="De-identified sample patient demographic & clinical metrics (N=100)")
    sub_cell.font = Font(name="Calibri", size=10, italic=True, color="64748B")
    
    # KPI Cards / Metrics Table
    total_patients = len(records)
    poly_count = sum(1 for r in records if r["polypharmacy_status"] == "Yes")
    avg_age = round(sum(r["age"] for r in records) / total_patients, 1)
    avg_meds = round(sum(r["med_count"] for r in records) / total_patients, 1)
    
    kpi_header_fill = PatternFill(start_color="3B82F6", end_color="3B82F6", fill_type="solid")
    kpi_header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    kpi_lbl_font = Font(name="Calibri", size=10, bold=True, color="334155")
    kpi_val_font = Font(name="Calibri", size=11, bold=True, color="1E3A8A")
    
    # KPI Table 1: High Level Metrics
    ws_sum.cell(row=5, column=2, value="Core Clinical Metric").fill = kpi_header_fill
    ws_sum.cell(row=5, column=2).font = kpi_header_font
    ws_sum.cell(row=5, column=3, value="Value").fill = kpi_header_fill
    ws_sum.cell(row=5, column=3).font = kpi_header_font
    
    kpis = [
        ("Total De-Identified Patients", total_patients),
        ("Mean Patient Age", f"{avg_age} years"),
        ("Polypharmacy Rate (>= 5 concurrent meds)", f"{poly_count}% ({poly_count}/100)"),
        ("Non-Polypharmacy Rate (< 5 meds)", f"{total_patients - poly_count}%"),
        ("Average Medications per Patient", f"{avg_meds} meds"),
        ("Oldest Patient", f"{max(r['age'] for r in records)} years"),
        ("Youngest Patient", f"{min(r['age'] for r in records)} years")
    ]
    
    for idx, (label, val) in enumerate(kpis, 6):
        c1 = ws_sum.cell(row=idx, column=2, value=label)
        c2 = ws_sum.cell(row=idx, column=3, value=val)
        c1.font = kpi_lbl_font
        c2.font = kpi_val_font
        c1.border = cell_border
        c2.border = cell_border
        c2.alignment = Alignment(horizontal="center")
        if idx % 2 == 0:
            c1.fill = row_fill_even
            c2.fill = row_fill_even
            
    # Specialty Breakdown Table
    ws_sum.cell(row=5, column=5, value="Specialist Distribution").fill = kpi_header_fill
    ws_sum.cell(row=5, column=5).font = kpi_header_font
    ws_sum.cell(row=5, column=6, value="Patient Count").fill = kpi_header_fill
    ws_sum.cell(row=5, column=6).font = kpi_header_font
    
    # Count by specialty
    specialist_counts = {}
    for r in records:
        spec = r["specialist"]
        specialist_counts[spec] = specialist_counts.get(spec, 0) + 1
        
    sorted_specs = sorted(specialist_counts.items(), key=lambda x: x[1], reverse=True)
    
    for idx, (spec, count) in enumerate(sorted_specs, 6):
        c1 = ws_sum.cell(row=idx, column=5, value=spec)
        c2 = ws_sum.cell(row=idx, column=6, value=f"{count} patients")
        c1.font = kpi_lbl_font
        c2.font = kpi_val_font
        c1.border = cell_border
        c2.border = cell_border
        c2.alignment = Alignment(horizontal="center")
        if idx % 2 == 0:
            c1.fill = row_fill_even
            c2.fill = row_fill_even

    # Gender Breakdown Table below KPI 1
    g_start_row = 15
    ws_sum.cell(row=g_start_row, column=2, value="Gender Representation").fill = kpi_header_fill
    ws_sum.cell(row=g_start_row, column=2).font = kpi_header_font
    ws_sum.cell(row=g_start_row, column=3, value="Count").fill = kpi_header_fill
    ws_sum.cell(row=g_start_row, column=3).font = kpi_header_font
    
    gender_counts = {}
    for r in records:
        g = r["gender"]
        gender_counts[g] = gender_counts.get(g, 0) + 1
        
    for idx, (gender, count) in enumerate(gender_counts.items(), g_start_row + 1):
        c1 = ws_sum.cell(row=idx, column=2, value=gender)
        c2 = ws_sum.cell(row=idx, column=3, value=f"{count} ({count}%)")
        c1.font = kpi_lbl_font
        c2.font = kpi_val_font
        c1.border = cell_border
        c2.border = cell_border
        c2.alignment = Alignment(horizontal="center")
        if idx % 2 == 0:
            c1.fill = row_fill_even
            c2.fill = row_fill_even

    # Age Group Breakdown Table
    a_start_row = 20
    ws_sum.cell(row=a_start_row, column=2, value="Age Stratification").fill = kpi_header_fill
    ws_sum.cell(row=a_start_row, column=2).font = kpi_header_font
    ws_sum.cell(row=a_start_row, column=3, value="Count").fill = kpi_header_fill
    ws_sum.cell(row=a_start_row, column=3).font = kpi_header_font
    
    age_groups = {
        "Young Adults (18 - 39)": sum(1 for r in records if r["age"] < 40),
        "Middle-aged (40 - 64)": sum(1 for r in records if 40 <= r["age"] <= 64),
        "Older Adults (65 - 79)": sum(1 for r in records if 65 <= r["age"] <= 79),
        "Geriatric (80+)": sum(1 for r in records if r["age"] >= 80)
    }
    
    for idx, (group, count) in enumerate(age_groups.items(), a_start_row + 1):
        c1 = ws_sum.cell(row=idx, column=2, value=group)
        c2 = ws_sum.cell(row=idx, column=3, value=f"{count} ({count}%)")
        c1.font = kpi_lbl_font
        c2.font = kpi_val_font
        c1.border = cell_border
        c2.border = cell_border
        c2.alignment = Alignment(horizontal="center")
        if idx % 2 == 0:
            c1.fill = row_fill_even
            c2.fill = row_fill_even

    wb.save(output_filename)
    print(f"Successfully generated '{output_filename}' with {len(records)} records and summary sheet.")

if __name__ == "__main__":
    records = generate_patient_records(100)
    create_excel_workbook(records, "sample_patients_100.xlsx")
