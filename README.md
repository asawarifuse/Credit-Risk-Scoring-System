# Credit Risk Scoring System

**Rule-based credit risk decision engine with LAMBDA classification, VBA automation, audit trail, and logistic regression benchmark.**

---

## 📅 Project Timeline (14 Days)

### Day 1 — Project Setup & Dataset Design ✅
- Created professional folder structure
- Initialized private GitHub repository
- Set up Excel workbook with 5 sheets (RAW_DATA, SCORING_ENGINE, DECISION_OUTPUT, AUDIT_TRAIL, DASHBOARD)
- Enabled Developer Tab, Macros, Iterative Calculation
- Generated 500 synthetic loan applicant records via Mockaroo
- Built `ApplicantData` table with 12 KPI columns (Applicant_ID, Age, Annual_Income, Employment_Type, Employment_Years, Credit_Score, Existing_Loans, Loan_Amount_Requested, Loan_Tenure_Years, Monthly_EMI_Obligation, Collateral_Value, Defaulter_History)

---

## 🛠 Tech Stack
- Excel (LAMBDA, LET, SUMPRODUCT, Dynamic Arrays)
- VBA Macros
- Power Query
- Logistic Regression (Python benchmark)
- GitHub

---

## 📁 Folder Structure
Credit-Risk-Scoring-System/

├── data/ ← Raw CSV datasets

├── excel/ ← Main .xlsm workbook

├── screenshots/ ← Dashboard screenshots

├── docs/ ← Project documentation

└── README.md

### Day 2 — Eligibility Criteria Logic ✅
- Built 5 eligibility checks (Age 21–60, Income ≥3L, Credit Score ≥650, Employment ≥1yr, No Defaulter History)
- Nested IF/AND chains for pass/fail per criterion
- Combined eligibility using AND across all checks
- Rejection reason column using TEXTJOIN
- Result: 127 Eligible / 373 Ineligible

### Day 3 — Weighted Scoring Engine ✅
- 7 scoring criteria with IF-based scoring bands
- Criteria: Age, Income, Employment Years, Credit Score, Existing Loans, Collateral Value, Defaulter History
- Each criterion scored 0 to its max, total composite: 0–100
- Remarks: Strong (70+), Moderate (50–69), Weak (<50), Not Scored (Ineligible)
- Avg composite score, MAX/MIN verification