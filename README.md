# Orchids International School — Parent Lead Analysis Dashboard

> **A data-driven Excel dashboard to diagnose why parents don't enroll after admissions calls, built for the Trainee School Manager program.**

---

## Table of Contents

1. [Business Problem](#business-problem)
2. [Objective](#objective)
3. [Project Structure](#project-structure)
4. [Data Schema](#data-schema)
5. [Dashboard Sections](#dashboard-sections)
6. [Formulas Used](#formulas-used)
7. [How to Use](#how-to-use)
8. [How to Refresh Data](#how-to-refresh-data)
9. [Requirements](#requirements)
10. [Known Limitations](#known-limitations)
11. [Author & Context](#author--context)

---

## Business Problem

Orchids International School runs a **Trainee School Manager program** that includes a field marketing campaign — phone calls to parents of prospective students. Two primary segments are called:

- **Kindergarten (New)** — parents seeking first-time admission (~60% of leads)
- **Transfer** — parents whose children are currently in another school (~40% of leads)

Despite high call volumes, **enrollment conversion is approximately 20%**. The L&D and admissions teams have no structured, data-driven way to understand why the remaining 80% do not enroll. Specific pain points:

- No visibility into which fee or policy objections are most common.
- No measurement of awareness gaps — how many parents have never heard of key initiatives (0% EMI, sibling discounts, SOCH counselling, 125+ branch transfer policy, etc.).
- No prioritised follow-up system — high-potential pending leads are not separated from low-potential ones.
- No tracking of transport or safety concerns as drop-off drivers.

---

## Objective

Build an Excel dashboard that:

- Aggregates 100 parent lead records with realistic attributes derived from Orchids' fee structure, special initiatives, safety protocols, and lab offerings.
- Identifies the **top reasons for non-enrollment** among lost leads.
- Quantifies **awareness gaps** for 16 key initiatives.
- Provides a **follow-up priority list** segmented by High / Medium / Low priority.
- Works in **all Excel versions** — no dynamic array formulas (FILTER, SORT, UNIQUE).

---

## Project Structure

```
Orchids_Parent_Lead_Analysis/
│
├── Orchids_Parent_Lead_Analysis.xlsx   # Main Excel workbook
│   ├── Raw_Data                        # 100 parent lead records, 33 columns
│   ├── Dashboard                       # Live metrics, awareness gap, reasons, charts
│   └── Instructions                    # Formula reference guide
│
├── Orchids_Dashboard_Guide.docx        # Full step-by-step build guide (this repo)
├── README.md                           # This file
└── build_dashboard.py                  # Python script to regenerate the workbook
```

---

## Data Schema

The `Raw_Data` sheet contains **100 rows × 33 columns**.

### Core Lead Columns (A–K)

| Column | Field | Values |
|--------|-------|--------|
| A | Lead_ID | LD001 – LD100 |
| B | Parent_Type | Kindergarten (New) / Transfer |
| C | Current_School | School name or N/A |
| D | Child_Grade | Nursery, LKG, UKG, Grade 1–8 |
| E | Call_Date | 01-Apr-2026 to 30-May-2026 |
| F | Call_Outcome | Interested / Not interested / Callback requested |
| G | Reason_Not_Switched | One of 15 predefined objections (see below) |
| H | Conversion_Status | Enrolled / Lost / Follow-up pending |
| I | Followup_Needed | Yes / No |
| J | Transport_Feasibility | Yes - within route / No - not available / Not sure |
| K | Safety_Mentioned | Yes / No |

### Awareness Flags (L–AA) — 16 columns, each Yes/No

| Column | Initiative | Target Awareness Rate |
|--------|-----------|----------------------|
| L | SeatBlocking_NonRefundable | ~30% |
| M | Tuition_Installments | ~25% |
| N | 0Percent_EMI | ~18% |
| O | OneShot_Discount | ~28% |
| P | Sibling_Concession | ~35% |
| Q | Corporate_TieUp | ~15% |
| R | MidYear_Concession | ~20% |
| S | Imprest_Refundable | ~12% |
| T | Internal_Transfer_125Branches | ~22% |
| U | SOCH_Counselling | ~17% |
| V | Beyond_The_Bell | ~15% |
| W | Robotics_Lab | ~25% |
| X | Coding_Lab | ~20% |
| Y | Dance_Music_Theatre | ~30% |
| Z | Active_Schooling | ~18% |
| AA | OCFP_Program | ~10% |

### Objection Columns (AB–AF)

| Column | Field |
|--------|-------|
| AB | SeatBlocking_Objection |
| AC | Imprest_Fees_Concern |
| AD | Transport_Not_Available |
| AE | Safety_Concern |
| AF | Workshop_Interest |

### Priority Column (AG)

| Value | Condition |
|-------|-----------|
| High | Follow-up pending + reason is a fee/discount objection |
| Medium | Follow-up pending + reason is an awareness gap |
| Low | Follow-up pending + any other reason |
| N/A | Lead is Enrolled or Lost |

### Predefined Reasons (column G)

1. Seat blocking non-refundable too high
2. No 0% EMI awareness
3. No sibling discount
4. No corporate tie-up
5. Imprest fees not refundable
6. Transport not available
7. No knowledge of 125+ branch transfer
8. Safety concerns (anti-ragging, POCSO)
9. No counselling program (SOCH)
10. Curriculum mismatch
11. Fees too high – unaware of discounts
12. No after-school program (Beyond the Bell)
13. No robotics/coding lab interest
14. Happy with current school
15. Other

---

## Dashboard Sections

| Section | Location | What it shows |
|---------|----------|---------------|
| Key Metrics | B5:B10 | Total leads, enrolled, lost, follow-up pending, conversion rate, avg awareness |
| Conversion Funnel | A15:C18 | Drop-off from total calls → interested → follow-up → enrolled |
| Awareness Gap | F6:H21 | Count and % of leads aware of each initiative |
| Top Reasons | F26:H40 | Count and % of lost leads per objection reason |
| Follow-up Priority | A22:C25 | Breakdown of High / Medium / Low priority follow-up leads |
| Transport Breakdown | A32:C34 | Count and % per transport feasibility option |
| Chart Instructions | A37+ | Step-by-step guide to insert all 4 charts |

---

## Formulas Used

All formulas use standard Excel functions compatible with **Excel 2010 and above**. No dynamic array functions are used.

```excel
-- Total leads
=COUNTA(Raw_Data!A2:A101)

-- Count by conversion status
=COUNTIF(Raw_Data!H2:H101,"Enrolled")
=COUNTIF(Raw_Data!H2:H101,"Lost")
=COUNTIF(Raw_Data!H2:H101,"Follow-up pending")

-- Conversion rate (with error guard)
=IFERROR(B6/B5,0)

-- Awareness count per initiative (example: column L)
=COUNTIF(Raw_Data!L2:L101,"Yes")

-- Awareness percentage
=IFERROR(G6/B5,0)

-- Lost count per reason (two-condition filter)
=COUNTIFS(Raw_Data!G2:G101,F26,Raw_Data!H2:H101,"Lost")

-- Reason as % of total lost
=IFERROR(G26/B7,0)

-- Follow-up priority count
=COUNTIF(Raw_Data!AG2:AG101,"High")
```

**Key design rule:** Every division is wrapped in `IFERROR(..., 0)` to prevent `#DIV/0!` errors if data is cleared or the sheet is empty.

---

## How to Use

### View the dashboard
1. Open `Orchids_Parent_Lead_Analysis.xlsx`.
2. Click the **Dashboard** tab.
3. All metrics update automatically from Raw_Data.

### Generate the follow-up call list
1. Go to the **Raw_Data** sheet.
2. Click any cell in row 1 → **Data → Filter** (enable if not on).
3. Click the **Followup_Priority** (column AG) dropdown → select **High** → OK.
4. Click the **Followup_Needed** (column I) dropdown → select **Yes** → OK.
5. Select all visible rows → Ctrl+C → paste into a new sheet named `High_Priority_Calls`.
6. Clear filters (Data → Clear) when done.

### Add the 4 charts
Follow the chart instructions in the dashboard (rows 37+) or refer to `Orchids_Dashboard_Guide.docx` Section 4.

---

## How to Refresh Data

### Adding new leads
1. Paste new rows into Raw_Data starting at **row 102**.
2. Update formula ranges: on the Dashboard sheet, press **Ctrl+H** → Find `101` → Replace with your new last row number → Replace All.
3. Press **Ctrl+Alt+F9** to force full recalculation.

### Refreshing PivotTables (if you created any)
- Right-click any PivotTable → **Refresh**, or
- Go to **Data → Refresh All**.

---

## Requirements

| Requirement | Minimum version |
|-------------|----------------|
| Microsoft Excel | 2010 (formulas); 2019 for native Funnel chart |
| Python (to regenerate workbook) | 3.8+ |
| openpyxl | 3.1+ |

### Python dependencies

```bash
pip install openpyxl
```

### Regenerate the workbook

```bash
python build_dashboard.py
```

Output: `Orchids_Parent_Lead_Analysis.xlsx` in the current directory.

---

## Known Limitations

- **Mock data only.** All 100 leads are synthetically generated. Awareness rates, reasons, and outcomes are randomly assigned from realistic probability distributions — they are not real parent responses.
- **No dynamic arrays.** To ensure compatibility with Excel 2010–2019, FILTER/SORT/UNIQUE are not used. The follow-up list must be extracted manually via AutoFilter.
- **Static date range.** Call dates are hardcoded between 01-Apr-2026 and 30-May-2026. Update `start_date` and `end_date` in `build_dashboard.py` before regenerating for a new campaign cycle.
- **Funnel chart availability.** The native Funnel chart type requires Excel 2019 or Microsoft 365. Older versions need a manual workaround using a Stacked Bar chart (see guide Section 4.3).
- **No real-time data connection.** This is a static workbook. It does not connect to a CRM or call centre platform. Data must be manually pasted into Raw_Data.

---

## Author & Context

| Field | Details |
|-------|---------|
| Program | Trainee School Manager Program |
| Team | L&D and Admissions, Orchids International School |
| Document type | Internal training and analysis tool |
| Campaign period | April–May 2026 |
| Data | 100 synthetic leads (seed: 42, fully reproducible) |
| Workbook generated | Python 3 + openpyxl |
| Guide authored | May 2026 |

---

*For questions about the dashboard formulas or chart setup, refer to `Orchids_Dashboard_Guide.docx` or the Instructions sheet inside the workbook.*
