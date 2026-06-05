# Payment Review Tracker

An Excel-native payment monitoring and reconciliation tool for tracking pension and allowance payments across 15 claimants over 12 months. Built with advanced Excel formulas, conditional formatting, data validation, and VBA automation macros.

---

## Overview

This tracker simulates the kind of payment review workflow used in government compliance and insurance organizations. Unlike a code-first pipeline, this project lives entirely in Excel — the way real analysts work day-to-day — using formulas to automate calculations, conditional formatting to surface issues visually, and VBA macros to handle repetitive review tasks.

---

## Tech Stack

| Tool | Usage |
|------|-------|
| Excel (Advanced) | Core workbook — formulas, pivot-style summaries, charts, data validation |
| Excel Formulas | SUMIF, COUNTIFS, IF/AND, IFERROR, conditional aggregations |
| Conditional Formatting | Color-coded risk levels, discrepancy heat maps, compliance rate scales |
| VBA | 5 automation macros for formatting, letter generation, export, and status updates |
| Python (openpyxl) | Workbook generation and data population |

---

## Workbook Structure (5 Sheets)

### 1. Instructions
User guide explaining the purpose of each sheet, key formulas used, and step-by-step usage instructions.

### 2. Payment Log
Master data entry sheet — 180 records (15 claimants × 12 months). Contains:
- Expected vs actual payment amounts
- Auto-calculated discrepancy column (`=Actual - Expected`)
- Status dropdown (Paid / Late / Missed / Under Review / Discrepancy)
- Color-coded rows by status using conditional formatting

### 3. Claimant Summary
Pivot-style summary auto-calculated from Payment Log using:
- `SUMIF` — total expected and actual per claimant
- `COUNTIFS` — flag count per claimant
- `IF` — risk level assignment (HIGH / MEDIUM / LOW)
- `IFERROR` — safe compliance rate calculation
- Color scale on discrepancy column (green → red gradient)

### 4. Monthly Dashboard
Month-by-month KPI breakdown with:
- Total disbursed vs expected per month
- Compliance rate per month (color scale formatting)
- Bar chart: Expected vs Actual payments by month
- Pie chart: Payment status distribution

### 5. Discrepancy Register
Filtered review queue of all non-Paid transactions (78 flagged records):
- AutoFilter enabled for sorting by month, status, claimant
- Resolution dropdown (Pending / In Progress / Resolved / Escalated)
- Color-coded by severity

---

## Key Excel Formulas Used

```excel
-- Total expected payments per claimant
=SUMIF('Payment Log'!A:A, A4, 'Payment Log'!F:F)

-- Count non-compliant payments per claimant
=COUNTIFS('Payment Log'!A:A, A4, 'Payment Log'!I:I, "<>Paid")

-- Risk level classification
=IF(G4>=3, "HIGH", IF(G4>=1, "MEDIUM", "LOW"))

-- Safe compliance rate (avoids division by zero)
=IFERROR(1 - G4/12, 0)

-- Monthly payment total
=SUMIF('Payment Log'!E:E, B4, 'Payment Log'!G:G)
```

---

## VBA Macros (Alt + F8)

| Macro | What it does |
|-------|-------------|
| `FormatAllSheets` | Auto-fits columns and standardizes fonts across all sheets |
| `HighlightHighRisk` | Highlights HIGH risk claimant rows in red in Claimant Summary |
| `GenerateMissedPaymentsList` | Creates a new sheet listing all missed payments for letter generation |
| `MarkResolved` | Marks selected Discrepancy Register row as Resolved (green highlight) |
| `ExportDiscrepancyRegister` | Exports discrepancy register to a dated .xlsx file on Desktop |

---

## Sample Output

- **180 payment records** tracked (15 claimants × 12 months)
- **78 flagged transactions** (43.3% of records had at least one anomaly)
- Risk levels auto-assigned: HIGH / MEDIUM / LOW per claimant
- Missed payments automatically queued for letter generation via VBA

---

## Skills Demonstrated

- Advanced Excel formula design (SUMIF, COUNTIFS, IFERROR, nested IF)
- Conditional formatting for compliance monitoring workflows
- Excel data validation (dropdown lists for status and resolution fields)
- VBA macro development for analyst workflow automation
- Excel chart creation (bar + pie charts linked to live formula data)
- Payment reconciliation and compliance review process design
