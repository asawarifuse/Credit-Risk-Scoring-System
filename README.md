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

### Day 4 — Decision Classification + LAMBDA + LET ✅
- Built custom LAMBDA function: CLASSIFY_CREDIT
- Uses LET to name intermediate logic (IsIneligible, IsStrong, IsModerate)
- Auto-classifies: Approve (70+), Review (50–69), Reject (<50 or Ineligible)
- LAMBDA deployed across all 500 rows
- Result: 106 Approve, 20 Review, 374 Reject

### Day 5 — Dynamic Arrays + Smart Lookups ✅
- FILTER + SORT to extract Approved / Review / Rejected dynamically
- UNIQUE score bands
- Summary counts above each section
- All lists update live with source data

### Day 6 — Audit Trail (Formula Layer) ✅
- Auto-log all 500 applicants: Timestamp, ID, Score, Decision, Reason
- Unique timestamps per row
- Reason strings explain each decision
- VBA static timestamp upgrade scheduled for Day 8

### Day 7 — Phase 1 Review + Stress Test ✅
- Verified all 500 rows for consistency
- Checked min/max edge cases (Age 21–65, Credit 302–900, Income 1.5L–25L)
- Manual spot check of 3 random applicants — logic verified
- Zero formula errors across all sheets
- Engine works end-to-end

---

## ✅ PHASE 1 COMPLETE — Intelligence Layer
The rule engine is operational: eligibility → scoring → classification → audit trail.

### Day 8 — VBA Audit Log Writer ✅
- Macro stamps static timestamps per record (no volatile formulas)
- Loops through all 500 applicants
- Reason strings for each decision
- Clears old log before refresh
- 500 records written in under 1 second

### Day 9 — One-Click Automation Macro ✅
- RunFullEngine macro: refresh → score → classify → audit log → save
- Form control button on DECISION_OUTPUT sheet
- Screen updating disabled for speed
- Error handler with user-friendly message box

### Day 10 — Error Handling + Validation Layer ✅
- ValidateRawData macro: checks Age, Credit Score, Income, Loan Amount, missing IDs
- RunFullEngine calls validation before processing
- On Error GoTo error handler with user-friendly messages
- Screen updating + calculation control for stability