# Payment Monitoring & Reconciliation Tracker

An Excel-native payment monitoring and reconciliation tool for tracking pension and allowance payments across 15 claimants over 12 months. Built with advanced Excel formulas, conditional formatting, VBA automation macros, and a Power Automate Desktop workflow.

---

## Overview

This tracker simulates the kind of payment review workflow used in government compliance and insurance organizations. This project lives primarily in Excel — using formulas to automate calculations, conditional formatting to surface issues visually, VBA macros to handle repetitive review tasks, and Power Automate Desktop to automate daily file handling and report distribution.

---

## Tech Stack

| Tool | Usage |
|------|-------|
| Excel (Advanced) | Core workbook — formulas, pivot-style summaries, charts, data validation |
| Excel Formulas | SUMIF, COUNTIFS, IF/AND, IFERROR, conditional aggregations |
| Conditional Formatting | Color-coded risk levels, discrepancy heat maps, compliance rate scales |
| VBA | 5 automation macros for formatting, letter generation, export, and status updates |
| Power Automate Desktop | Automated daily file handling, macro execution, and flagged report distribution |
| Python (openpyxl) | Workbook generation and data population |

---

## Sample Output

- **180 payment records** tracked (15 claimants × 12 months)
- **78 flagged transactions** (43.3% of records had at least one anomaly)
- Risk levels auto-assigned: HIGH / MEDIUM / LOW per claimant
- Missed payments automatically queued for letter generation via VBA

---
