# Corporate Finance Analytics Suite
**Excel · Power Query · PivotTables · Multi-Entity GL Consolidation · FY 2018–2020**

<p align="left">
  <img src="https://img.shields.io/badge/Excel-Advanced-217346?style=flat-square&logo=microsoft-excel&logoColor=white" />
  <img src="https://img.shields.io/badge/Power%20Query-ETL-F2C811?style=flat-square&logo=powerbi&logoColor=black" />
  <img src="https://img.shields.io/badge/GL%20Entries-27%2C909-FF6B35?style=flat-square" />
  <img src="https://img.shields.io/badge/Entities-7%20Countries-6C5CE7?style=flat-square" />
  <img src="https://img.shields.io/badge/Statements-IFRS%20Aligned-0078D4?style=flat-square" />
  <img src="https://img.shields.io/badge/Status-Live-00B894?style=flat-square" />
</p>

---

> Built a CFO-ready reporting suite from 27,909 raw General Ledger entries across 7 subsidiaries (USA, Canada, UK, Germany, France, Australia, New Zealand). Used **Power Query to merge 4 source tables** into a single consolidated dataset, then built live financial statements via **PivotTables and SUMIFS** — zero hardcoded output values.

---

## Snapshot

| | |
|---|---|
| **GL Volume** | 27,909 journal entries · FY2018–FY2020 |
| **Entities** | 7 countries · 4 regions (Americas, Europe, Oceania) |
| **ETL Layer** | Power Query — merged GL + Chart of Accounts + Territory + Calendar into one flat table |
| **Reporting Layer** | PivotTables (P&L, Balance Sheet) · SUMIFS (CFO Dashboard KPIs) |
| **Statements Produced** | P&L · Balance Sheet · Cash Flow (IAS 7) · SoCE |
| **Modelling Modules** | Capital Appraisal · 5-Method Depreciation · Loan Amortisation |
| **Revenue Tracked** | $3.6M → $7.8M (+119% over 3 years) |

---

## Results at a Glance

| Metric | FY 2018 | FY 2019 | FY 2020 |
|---|---|---|---|
| Revenue (USD) | 3,575,428 | 5,697,845 | 7,835,369 |
| Gross Profit | 2,383,246 | 3,968,546 | 5,341,360 |
| EBITDA | 740,801 | 1,475,688 | 1,522,166 |
| Net Profit | 623,856 | 1,303,147 | 1,289,945 |
| Gross Margin | 66.7% | 69.6% | 68.2% |
| Net Margin | 17.4% | 22.9% | 16.5% |
| Return on Equity | 20.6% | 16.7% | 12.5% |
| Current Ratio | 5.17× | 8.20× | 8.08× |
| Total Assets | 3,875,802 | 9,238,096 | 12,320,001 |

---

<!-- SCREENSHOT PLACEHOLDER — replace with actual image -->
<!-- ![CFO Dashboard](docs/dashboard-preview.png) -->
<!-- ![P&L Statement](docs/pl-preview.png) -->

---

## Problem Statement

A multi-entity business operating across three regions needed to consolidate subsidiary financials into a single reporting framework. Requirements:

- Aggregate raw GL data from 7 legal entities into group-level statements
- Produce IFRS-aligned P&L, Balance Sheet, and Cash Flow with a single formula refresh
- Enable investment decisions via NPV/IRR/MIRR scenario analysis
- Compare depreciation methods for asset-heavy entities
- Give the CFO a one-page view of group KPIs, margins, and year-on-year movement

---

## Architecture

```
LAYER 1 — SOURCE DATA (GL_SourceData.xlsx)  4 tables
│
├── GL                 27,909 entries · Date / Entity / Account / Amount
├── Chart of Accounts  54 accounts mapped to Report > Class > SubClass > Account
├── Territory          7 countries × 4 regions
├── Calendar           1,096-day date dimension (FY2018–FY2020)
├── CashFlow_St        Account-to-Cash-Flow-line mapping logic
└── SoCE_St            Account-to-Equity-movement mapping logic
          │
          │  POWER QUERY — Merge Queries
          │  GL joined to Chart of Accounts, Territory, Calendar
          │  Output: single flat table with 14 enriched columns
          ▼
LAYER 2 — CONSOLIDATED DATASET (GL RAW DATA sheet)
  27,909 rows · GL columns + Account hierarchy + Country + Region
          │
          ├─── PIVOTTABLES ──────────────────────────────────────────┐
          │    P&L Statement · Balance Sheet                         │
          │    Slice by Year / Entity / Account Class                │
          │                                                          │
          └─── SUMIFS (multi-condition) ────────────────────────────┐│
               CFO Dashboard KPIs                                   ││
               Account_key + Territory_key + Year filters           ││
                                                                    ▼▼
LAYER 3 — OUTPUT MODEL (CorporateFinanceAnalyticsSuite.xlsx)
│
├── Assumptions        Central input hub — all drivers in one place
├── CFO Dashboard      8 KPIs · margin analysis · YoY variance · multi-entity summary
├── P&L Statement      PivotTable · Consolidated Income Statement FY2018–FY2020
├── Balance Sheet      PivotTable · IFRS-style · Assets = Liabilities + Equity check
├── Cash Flow          Indirect method per IAS 7
├── SoCE               Share capital · premium · retained earnings movements
├── Capital Appraisal  NPV / IRR / MIRR / Payback · 3-scenario (Base/Best/Worst)
├── Depreciation       SLN · DB · DDB · SYD · VDB side-by-side comparison
└── Loan Amortisation  120-month schedule · PMT / PPMT / IPMT / CUMIPMT
```

---

## Module Detail

### CFO Executive Dashboard
Eight KPI panels covering Revenue, Gross Profit, Net Profit, Total Assets, Cash & Equivalents, and Operating Expenses — each showing FY2018/19/20 columns. Margin analysis block (Gross Margin %, Net Margin %, ROE, ROA, Current Ratio, D/E Ratio). YoY variance table with absolute and percentage movement. Multi-country consolidation block showing entity list and region classification. Entire sheet is formula-driven — no data entry.

### P&L Statement
Built as a PivotTable on the consolidated GL RAW DATA flat table. Rows: account line items (Revenue, Cost of Sales, Operating Expenses, EBITDA, D&A, Interest, Tax, Net Profit). Columns: FY2018 / FY2019 / FY2020. PivotTable field filters allow instant slicing by entity, region, or account class — no formula rewriting required.

### Balance Sheet
Built as a PivotTable on the same consolidated dataset, filtered to Balance Sheet accounts only. IFRS presentation: Current Assets → Non-Current Assets → Current Liabilities → Non-Current Liabilities → Equity. Assets = Liabilities + Equity check formula confirms balance at all times.

### Cash Flow Statement
Indirect method per IAS 7. Operating section starts from Profit Before Tax and adjusts for: depreciation, working capital movements (receivables, inventory, payables), interest, and tax paid. Investing section covers CAPEX and disposals. Financing section covers share issues, dividends, and loan movements. Closing cash reconciles to Balance Sheet cash balance.

### Capital Appraisal
Three scenarios (Base WACC 12%, Best 8%, Worst 16%). Outputs: NPV, IRR, MIRR (with 10% reinvestment rate), Payback Period, Discounted Payback Period. All inputs linked to Assumptions sheet — change WACC once and all three scenarios update.

| Excel Function | Purpose |
|---|---|
| `=NPV()` | Present value of future cash flows |
| `=IRR()` | Internal rate of return |
| `=MIRR()` | Modified IRR with explicit reinvestment rate |
| Manual formula | Payback and discounted payback periods |

### Depreciation Schedule

| Method | Function | Behaviour |
|---|---|---|
| Straight Line | `=SLN()` | Equal annual charge |
| Declining Balance | `=DB()` | Fixed % on reducing book value |
| Double Declining | `=DDB()` | Accelerated · 2× SLN rate |
| Sum of Years' Digits | `=SYD()` | Accelerated · fraction-based |
| Variable Declining | `=VDB()` | Switches to SLN when optimal |

Asset: $1,200,000 cost · $100,000 salvage · 10-year life. Cumulative depreciation and closing book value columns for each method.

### Loan Amortisation
$500,000 principal · 8.0% annual rate · 10-year term · monthly payments. Full 120-row schedule. Inputs on Assumptions sheet — change any parameter and the entire schedule rebuilds.

| Function | Output |
|---|---|
| `=PMT()` | Fixed monthly instalment: $6,066 |
| `=PPMT()` | Principal component per period |
| `=IPMT()` | Interest component per period |
| `=CUMIPMT()` | Cumulative interest at any period |

---

## Key Technical Decisions

**Why Power Query for the ETL layer?**
Four separate source tables (GL, Chart of Accounts, Territory, Calendar) needed to be joined before any reporting could happen. Doing this with VLOOKUP or INDEX-MATCH inside the output model would create fragile, formula-heavy sheets that break when source columns shift. Power Query merge queries produce a clean, refreshable flat table — one click updates the entire consolidated dataset.

**Why PivotTables for P&L and Balance Sheet?**
PivotTables sit directly on the Power Query output table. When source data is refreshed, P&L and Balance Sheet update in one step with no formula maintenance. They also enable instant entity-level or region-level slicing that would require rewriting SUMIFS arrays manually.

**Why SUMIFS for the CFO Dashboard?**
Dashboard KPIs need precise, labelled cells with custom layout — not the grid structure PivotTables produce. SUMIFS with Account_key + Territory_key + Year conditions give exact control over where each number lands on the dashboard.

**Why a central Assumptions sheet?**
Any model where inputs are scattered across sheets creates version-control problems. A single Assumptions hub means one change cascades correctly through Capital Appraisal, Depreciation, and Loan Amortisation simultaneously.

---

## File Structure

```
corporate-finance-analytics-suite/
│
├── README.md
├── data/
│   └── GL_SourceData.xlsx              ← Raw GL + all reference tables
├── model/
│   └── CorporateFinanceAnalyticsSuite.xlsx  ← Full 9-module output model
└── docs/
    ├── dashboard-preview.png           ← CFO Dashboard screenshot
    └── pl-preview.png                  ← P&L Statement screenshot
```

---

## How to Open

1. Download both files — keep them in the same folder
2. Open `CorporateFinanceAnalyticsSuite.xlsx` first
3. If prompted to update links, click **Enable Content** and **Update**
4. Navigate to `Assumptions` to change any scenario inputs
5. All output sheets update automatically on recalculation (F9)

> Regional note: If `=PMT()` shows `#NAME?`, your Excel language pack uses a localised function name. This does not affect any other sheet.

---

## Skills Demonstrated

**ETL & Data Preparation** — Power Query · multi-table merge (GL + Chart of Accounts + Territory + Calendar) · refreshable flat table output

**Accounting & Reporting** — P&L · Balance Sheet · Cash Flow (IAS 7) · SoCE · IFRS presentation

**Financial Modelling** — Capital appraisal · multi-scenario NPV/IRR/MIRR · sensitivity analysis

**Excel — Advanced** — PivotTables on Power Query · SUMIFS · PMT/PPMT/IPMT/CUMIPMT · SLN/DB/DDB/SYD/VDB · NPV · IRR · MIRR

**FP&A** — KPI design · margin analytics · YoY variance · multi-entity consolidation

---

## Author

**G. Suneel** ·
Finance Analyst | Excel | Power Query | FP&A | Reporting Automation  
Hyderabad, India
[github.com/sunilg9090](https://github.com/sunilg9090)

---

*27,909 GL entries · 7 countries · 9 output modules · all formulas live from source*# corporate-finance-analytics-suite
Multi-entity GL consolidation | Power Query + PivotTables + Excel | 27,909 GL entries | 7 countries | FY2018–2020
