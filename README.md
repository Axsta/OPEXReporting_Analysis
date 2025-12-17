# Operational Expense Variance Reporting & Dashboard (Excel + PowerPivot + VBA)

OPEXReporting_Analysis is a structured Excel solution designed to automate monthly Operational Expense (OPEX) reporting.
It streamlines financial data entry, variance calculation, and executive dashboarding using a combination of:

VBA automation | PowerPivot Data Model | DAX measures | PivotTables, charts & slicers

The goal is to provide a real-world, enterprise-grade financial reporting workflow for FP&A and MIS teams.

<img width="1128" height="734" alt="image" src="https://github.com/user-attachments/assets/c8f7ed41-0349-4225-986e-f70aeaec952a" />

___

### Key Features
1. Monthly OPEX Input Files

Standardized worksheets for each month (Jan–Dec 2025) for consistent expense data entry.

2. Automated Summary Generation (VBA)

A macro-driven engine produces a formatted, professional summary page :

* KPI Block (Budget, Actual, Net Variance, Variance %)
* Variance Table
* Top 5 Variances Chart
* Clean formatting, borders, colors, and alignment
* Error handling + automatic sheet selection

3. Integrated Data Model (PowerPivot)

PowerPivot is used to model all monthly files into a single unified dataset.
Includes DAX measures for:

* YTD Actuals
* YTD Budget
* Variance $ and Variance %
* Rolling 3-Month Actuals
* Current month KPIs
4. Interactive Dashboard

Excel dashboard powered by:

* PivotTables
* Slicers
* PivotCharts
* DAX-based metrics

5. Export to PDF

A button-driven VBA module exports the summary sheet as a polished PDF report.

----
### Ideal For
Finance teams, MIS reporting, budgeting, and monthly variance analysis.

----

### Repository Structure

    /Input/              → Monthly OPEX input files (Jan–Dec 2025)
    /Excel_Files/        → Main reporting model (.xlsm)
    /Scripts/            → VBA modules (summary, formatting, PDF export)
    /Output/             → Generated PDFs & dashboard screenshots


---
## Business Context
Operational expense (OPEX) reporting requires accurate monthly tracking, variance analysis, and executive-ready summaries. Manual reporting increases errors and delays.

---
## Outcome
- Automated monthly OPEX reporting
- Reduced manual effort and reporting time
- Improved visibility into cost variances for stakeholders
