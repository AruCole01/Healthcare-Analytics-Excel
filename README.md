🏥 Medivista Hospital — Healthcare Analytics Dashboard

Transforming raw hospital activity into measurable performance and actionable leadership decisions.# Healthcare-Analytics-Excel


📖 Project Overview
This is my first end-to-end healthcare analytics portfolio project, built entirely in Microsoft Excel using Power Pivot and DAX. It analyses operational, financial, and clinical performance data for a fictional hospital, Medivista Hospital, across a 3-year period (2023–2025).
The project goes beyond reporting. It is structured around the Insight → Decision → Action framework used by high-performing analytics teams to move from data to measurable business impact.


🎯 Project Goals

Analyse financial performance across revenue, claims, and collections
Understand patient population risk and clinical trends
Measure operational efficiency across departments and appointments
Deliver evidence-based strategic recommendations to hospital leadership
Demonstrate senior-level analytics thinking for portfolio purposes



🔍 The Business Problem
Hospital leadership needed answers to three strategic questions:
Strategic PillarCore Question💰 Growth & SustainabilityAre we financially and operationally sustainable?🛡️ Risk & QualityAre we managing patient risk effectively?⚙️ Efficiency & OptimisationAre we operating efficiently enough to grow?
Without structured analytics, leaders were:


Reacting blindly to financial pressure
Misallocating resources across departments
Missing revenue leakage in claims and billing
Unable to identify high-risk patient groups early




🗂️ Dataset Structure
The data model contains 13 interconnected tables covering the full hospital operations lifecycle:
TableKey ColumnsPurposeAppointmentsappointment_id, scheduled_date, statusTrack scheduling & no-showsClaimsclaim_id, payments, total_charges, submission_date, claim_statusFinancial performanceEncountersencounter_id, encounter_type, admit_dateDepartment activityPatientspatient_id, dob, gender, locationDemographicsVitalsvital_id, bmi, systolic_bp, diastolic_bp, heart_rateClinical health indicatorsDiagnosesdiagnosis_code, description, chronic_flagDisease burdenDepartmentsdepartment_id, department_nameOperational structureProvidersprovider_id, specialtyClinical workforceMedicationsmedication_id, encounter_idPrescription trackingLabslab_id, test_name, resultDiagnostic dataImagingimaging_id, modality, body_partRadiology activityInventoryitem_id, quantity, reorder_pointSupply chainSatisfactionsatisfaction_id, overall_scorePatient experience
Total Records: ~100,000+ rows across all tables





🛠️ Tools & Technologies
ToolUsageMicrosoft ExcelPrimary analysis and dashboard environmentPower PivotData modelling, table relationships, DAX engineDAX (Data Analysis Expressions)All KPI measures, YoY/MoM calculationsPower QueryData cleaning, type correction, column transformationPivotTables & PivotChartsData aggregation and visualisationConditional FormattingKPI colour coding (▲ green / ▼ red)SlicersInteractive dashboard filtering





🔗 Data Model & Relationships
The data model follows a star schema pattern with Claims and Encounters as the central fact tables:
Patients ──────────── Encounters ──────────── Departments
                           │
                    Appointments
                           │
              ┌────────────┼────────────┐
           Claims       Vitals      Diagnoses
              │
         ┌────┴────┐
      Imaging    Labs

      
Key Relationships
From TableTo TableJoin KeyTypeEncountersPatientspatient_idMany-to-OneEncountersDepartmentsdepartment_idMany-to-OneClaimsEncountersencounter_idMany-to-OneAppointmentsPatientspatient_idMany-to-OneVitalsPatientspatient_idMany-to-One



📐 DAX Measures Built
Financial KPIs
dax-- Total Payments
Total Payments = SUM(Claims[payments])

 Claims Paid Count
Claims Paid = CALCULATE(COUNT(Claims[claim_id]), Claims[claim_status] = "Paid")

 Collection Rate
Collection Rate % = DIVIDE(SUM(Claims[payments]), SUM(Claims[total_charges]), 0)

 Average Pay Per Claim
Avg Pay Per Claim = DIVIDE(SUM(Claims[payments]), COUNT(Claims[claim_id]), 0)



Charge-to-Payment Variance
Charge-to-Payment Variance % = 1 - [Collection Rate %]
Year-over-Year (YoY) Measures
dax-- YoY Total Payments
YoY TP =
VAR CurrentYear = YEAR(MAX(Claims[submission_date]))
VAR PrevYear = CurrentYear - 1
VAR CurrentPayments = CALCULATE(SUM(Claims[payments]))
VAR PrevPayments = CALCULATE(
    SUM(Claims[payments]),
    YEAR(Claims[submission_date]) = PrevYear
)
RETURN CurrentPayments - PrevPayments

-- YoY Total Payments %
YoY TP % =
VAR CurrentYear = YEAR(MAX(Claims[submission_date]))
VAR PrevYear = CurrentYear - 1
VAR CurrentPayments = CALCULATE(SUM(Claims[payments]))
VAR PrevPayments = CALCULATE(
    SUM(Claims[payments]),
    YEAR(Claims[submission_date]) = PrevYear
)
RETURN DIVIDE(CurrentPayments - PrevPayments, PrevPayments, 0)

Note: SAMEPERIODLASTYEAR() was not used due to version limitations in Excel Power Pivot.
The VAR + YEAR(MAX()) pattern was used instead — fully compatible with all Excel versions.

Month-over-Month (MoM) Measures
dax-- MoM Total Payments (January-safe — rolls back to Dec of prior year)
MoM TP =
VAR CurrentYear = YEAR(MAX(Claims[submission_date]))
VAR CurrentMonth = MONTH(MAX(Claims[submission_date]))
VAR PrevMonth = CurrentMonth - 1
VAR PrevYear = IF(PrevMonth = 0, CurrentYear - 1, CurrentYear)
VAR AdjPrevMonth = IF(PrevMonth = 0, 12, PrevMonth)
VAR CurrentPayments = CALCULATE(SUM(Claims[payments]))
VAR PrevPayments = CALCULATE(
    SUM(Claims[payments]),
    YEAR(Claims[submission_date]) = PrevYear,
    MONTH(Claims[submission_date]) = AdjPrevMonth
)
RETURN CurrentPayments - PrevPayments

The same VAR pattern was applied to Claims Paid, Avg Pay Per Claim, and Collection Rate % for both YoY and MoM — 16 measures in total.



⚠️ Technical Challenges & How I Solved Them
This section documents the real problems encountered during the build — and how they were resolved.


❌ Challenge 1: SAMEPERIODLASTYEAR Returning Duplicate Date Error
Error Message:
A date column containing duplicate dates was specified in the call to function 'SAMEPERIODLASTYEAR'. This is not supported.
Root Cause:
The Appointments table contained multiple rows per date (many appointments per day), making it impossible to use as a date table for time intelligence functions.
What I Tried:


Marking the Appointments table as a Date Table → Failed (duplicate dates)
Creating a separate Dates table in Power Query → Table was query-sourced and couldn't be deleted from Power Pivot directly
Attempting Design → Mark as Date Table on the Dates table → Still failed due to inherited duplicates from source query


Final Solution:
Abandoned SAMEPERIODLASTYEAR entirely. Switched to a self-contained VAR-based pattern using YEAR(MAX()) and MONTH(MAX()) which:



Requires no marked Date Table
Works on all Excel versions including older Power Pivot builds
Correctly handles January → December rollback for MoM calculations
Is fully slicer-compatible



❌ Challenge 2: SELECTEDVALUE Not Recognised
Error Message:
Failed to resolve name 'SELECTEDVALUE'. It is not a valid table, variable, or function name.
Root Cause:
SELECTEDVALUE() was introduced in later versions of DAX and is not available in older Excel Power Pivot builds.
Solution:
Replaced SELECTEDVALUE() with MAX() which returns the same result in a filtered slicer context and is universally supported:
dax-- Instead of this:
VAR CurrentYear = YEAR(SELECTEDVALUE(Claims[submission_date]))


Use this:
VAR CurrentYear = YEAR(MAX(Claims[submission_date]))
❌ Challenge 3: Measure Returning Blank (No Figures Showing)
Root Cause:
The MAX() function needs a filter context to know what "current period" means. Without a slicer connected to the PivotTable, there was no context — so all YoY/MoM measures returned blank.
Solution:



Added a Year slicer using submission_date (Year) column
Added a Month slicer using submission_date (Month) column
Connected both slicers to the Financial PivotTable via Report Connections
Once a year/month was selected, all measures populated correctly



❌ Challenge 4:  Wrong Table Assignment for Measures
Error Message:
Measure group 'Patients' must have at least one partition defined in Tabular mode.
Root Cause:
Measures referencing Claims columns were being created under the Appointments table, causing a cross-table conflict.
Solution:
Changed the Table name dropdown in the Manage Measures dialog from Appointments to Claims for all financial measures. DAX measures should always live in the same table whose columns they reference.




❌ Challenge 5: scheduled_time Column Showing 12/30/1899
Root Cause:
Excel's base date is December 30, 1899. When a column contains time-only values, Excel defaults to this base date, causing the column to appear as 12/30/1899 10:15 AM instead of just 10:15 AM.
Solution:
Fixed in Power Query Editor:



Located the scheduled_time column
Right-clicked → Change Type → Time
This stripped the false base date and retained only the time value
Applied Close & Load to push changes back to the data model




📊 Key Findings & Insights
💰 Financial Performance
KPIValueInsightTotal Revenue₦20,768,087Full-year collected revenueTotal Payments₦16,665,503Actual cash receivedOutstanding Balance₦12,995,055Uncollected revenue at riskCollection Rate56.2%⚠️ Well below 80% industry benchmarkCharge-to-Payment Gap43.8%Significant billing inefficiencyDenied Claims Value₦1,328,412Revenue lost to claim denialsAvg Pay Per Claim₦2,083Baseline for payer benchmarkingYoY Payment Growth+2%Modest growth, sustainability concern
Key Insight: While total payments grew 2% YoY, the collection rate of 56.2% means nearly half of all charges are not being collected. The ₦13M outstanding balance represents a critical cash flow risk.



🏥 Patient Population & Clinical Risk
KPIValueInsightTotal Patients2,000Active patient populationHypertension (Stage 1+2)65.2%4 in 5 patients at cardiovascular riskObese Patients21.0%High chronic disease burdenDiabetic Patients18.7%Significant comorbidity overlapSmokers18.0%Compounding risk factorElderly (75+)119 patientsAging population trendAverage Age45.96 yearsMiddle-aged dominant cohortAverage BMI25.92Borderline overweight average
Key Insight: Over 65% of patients present with hypertension, a chronic, manageable condition. Combined with 21% obesity and 18.7% diabetes, the hospital is managing a high-risk population that demands proactive preventive care investment.



⚙️ Operational Performance
KPIValueInsightTotal Appointments6,0003-year appointment volumeCompleted4,789 (79.8%)✅ Strong completion baselineNo-Show Rate12.1% (725)⚠️ Revenue leakage riskCancellation Rate8.1% (486)Combined 20% appointment loss2023 Appointments2,067Baseline year2024 Appointments2,056▼ 0.53% YoY decline2025 Appointments1,877▼ 8.71% YoY decline — alarming trendOutpatient Share42.0%Largest department by volumeEmergency/Inpatient24.8% eachEqual load,  staffing concernSurgical/OR8.3%Underutilised capacity
Key Insight: Appointment volume is declining year-on-year with a sharp 8.71% drop in 2025. Combined with a 20% appointment loss rate from no-shows and cancellations, the hospital faces a dual threat to both revenue and capacity utilisation.



🩺 Clinical Activity
KPIValueTotal Encounters8,000Total Procedures9,471Total Lab Tests13,461Total Medications6,802Medication per Encounter0.85Average Satisfaction Score2.97 / 5Deceased Patients44 (2.2%)Average SPO296.5%Average Systolic BP124.5 mmHg



💊 Revenue by Payer
PayerRevenueEmployer Plan A₦2,894,950Private Self-Pay₦2,850,045GreenCare HMO₦2,834,932Community Aid₦2,772,468National Health Insurance₦2,656,726Blue Shield Health₦2,656,382
Key Insight: Revenue is evenly distributed across payers, no single payer dominates. This reduces dependency risk but limits negotiating leverage with any individual payer.



📅 Revenue by Month
MonthRevenue (₦)MoM %Jan1,718,709—Feb1,499,196▼ 12.77%Mar1,866,867▲ 24.52%Apr1,749,847▼ 6.27%May1,843,224▲ 5.34%Jun1,773,148▼ 3.80%Jul2,047,744▲ 15.49%Aug1,924,682▼ 6.01%Sep1,681,019▼ 12.66%Oct1,949,513▲ 15.97%Nov1,409,082▼ 27.72%Dec1,305,055▼ 7.38%
Key Insight: Revenue peaks in July and October then drops sharply in November–December. This seasonal pattern suggests either reduced patient volume or delayed claim submissions in Q4, both worth investigating.



🎯 Strategic Recommendations
Structured using the Insight → Decision → Action framework across three executive decision domains:



1️⃣ Financial Performance — CFO Decision Support
LayerDetailInsightCollection rate is 56.2% vs 80% benchmark. ₦13M outstanding. Denied claims = ₦1.33M lost.DecisionPrioritise claims denial management. Renegotiate payer contracts. Target 80%+ collection rate.ActionAssign dedicated billing officer. Implement 30-day claim follow-up cycle. Build weekly collection rate KPI report. Set automated alerts for denied claims.



2️⃣ Patient Population & Clinical Risk — Medical Director Decision Support
LayerDetailInsight65% hypertensive. 21% obese. 18.7% diabetic. Chronic disease burden concentrated in middle-aged and elderly cohorts.DecisionLaunch chronic disease management programme. Prioritise preventive care budget. Create high-risk patient registry.ActionMonthly BP screening clinics. Diabetic patient follow-up call programme. BMI-targeted nutrition counselling. Risk-stratified care pathways. Quarterly population health report.



3️⃣ Operational Efficiency — COO Decision Support
LayerDetailInsight20% appointment loss rate. Volume declining 8.7% YoY in 2025. Emergency and Inpatient departments equally loaded. Surgical/OR underutilised at 8.3%.DecisionInvest in scheduling automation. Implement overbooking policy. Review staffing vs demand data. Expand Surgical capacity planning.ActionSMS/WhatsApp appointment reminders. Cancellation deposit policy. Weekly no-show dashboard review. Monthly department capacity report. Quarterly workflow audit.




📱 Dashboard Structure
The project includes 5 dashboard views:
DashboardAudienceFocusIntro DashboardAll stakeholdersHigh-level KPI overviewFinancials DashboardCFO / FinanceRevenue, claims, collection metricsKPIs Insights DashboardExecutive teamCross-functional performance summaryDashboard 2OperationsPatient volume, appointments, encountersDashboard VitalsMedical DirectorClinical health indicators and riskRecommendations SlideBoard / LeadershipStrategic Insight → Decision → Action



📚 What I Learned
Technical Skills


Built a multi-table data model in Power Pivot with 13 related tables
Wrote 16 custom DAX measures for YoY and MoM calculations across 4 KPIs
Debugged and resolved 5 distinct Power Pivot/DAX errors from scratch
Used the VAR + YEAR(MAX()) pattern as a version-safe alternative to SAMEPERIODLASTYEAR
Applied custom number formatting (▲0.00%;▼0.00%) for professional KPI display
Fixed data type issues in Power Query (time columns defaulting to 12/30/1899 base date)



Analytical Thinking

Structured analysis around What Happened → Why → So What stakeholder framework
Translated raw metrics into executive-level narratives across CFO, COO, and Medical Director domains
Applied Insight → Decision → Action framework to move beyond reporting into business impact
Identified seasonal revenue patterns, declining appointment trends, and chronic disease concentration



Lessons Learned

Always check which Excel version you are working with before choosing DAX functions
Time intelligence functions require a clean, deduplicated Date Table — build one from scratch rather than sourcing from operational tables
Measures must be assigned to the same table whose columns they reference to avoid partition errors
Dashboard design should follow the audience — what a CFO needs to see differs from what an operations manager needs




👤 Connect With Me
Aru Collins Agaji
Data Analyst | Port Harcourt, Nigeria
