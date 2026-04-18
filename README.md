# VBA Forecasting Tool & SQL Platform Integration

> Excel VBA-based forecasting tool for Deloitte Technology Finance — syncing forecast inputs to QlikSense, SAP HANA, Anaplan, and Oracle simultaneously via SQL after each period close.

![VBA](https://img.shields.io/badge/VBA-217346?style=flat-square&logo=microsoftexcel&logoColor=white)
![SQL](https://img.shields.io/badge/SQL-4479A1?style=flat-square&logo=mysql&logoColor=white)
![Excel](https://img.shields.io/badge/Advanced%20Excel-217346?style=flat-square&logo=microsoftexcel&logoColor=white)
![SAP HANA](https://img.shields.io/badge/SAP%20HANA-0FAAFF?style=flat-square&logo=sap&logoColor=white)
![Anaplan](https://img.shields.io/badge/Anaplan-0070F3?style=flat-square&logoColor=white)
![Oracle](https://img.shields.io/badge/Oracle%20EBS-F80000?style=flat-square&logo=oracle&logoColor=white)
![Status](https://img.shields.io/badge/Status-Live%20at%20Deloitte-success?style=flat-square)

---

## What This Does

The VBA Forecasting Tool is a structured Excel-based forecasting platform for Deloitte Technology Finance. Analysts input forecast numbers directly in Excel — and after period close, the tool automatically pushes those numbers via SQL to all four of Deloitte's analytics platforms: QlikSense, SAP HANA, Anaplan, and Oracle.

This standardised a previously fragmented forecasting process — replacing a mix of SharePoint forms, server-based inputs, and manual entries with a single, user-friendly Excel interface.

---

## Key Features

- **Single source of truth:** All forecast inputs entered in one Excel tool — one set of numbers, four platforms updated
- **Period-close sync:** After period close, SQL pushes forecast data to all platforms simultaneously
- **Multi-sheet structure:** Separate sheets for each dimension of forecasting
- **Full hierarchy coverage:** Revenue → Product → WBS → GL account → Cost center
- **Rate card integration:** Internal labor rate cards and OPR inputs built into the tool
- **Actuals-to-date view:** Current actuals visible alongside forecast for in-period comparison
- **User-friendly:** Replaces server-based inputs with a familiar Excel interface — reducing submission errors

---

## Workbook Structure

```
Forecasting Tool (Excel Workbook)
├── Sheet 1: Account Level Summary
│   └── GL account level forecast inputs
│   └── Actuals to date vs forecast comparison
│
├── Sheet 2: Internal Labor
│   └── Headcount by grade and cost center
│   └── Rate card applied automatically
│
├── Sheet 3: OPR Inputs
│   └── Operational run rate forecast
│   └── Project-specific cost inputs
│
├── Sheet 4: Revenue & Product
│   └── Revenue forecast by product line
│   └── WBS-level breakdown
│
├── Sheet 5: Current Year Plan
│   └── AOP plan reference (read-only)
│   └── Plan vs forecast variance view
│
└── Sheet 6: SQL Sync Log
    └── Confirmation of last sync timestamp
    └── Platform sync status (QlikSense / SAP / Anaplan / Oracle)
```

---

## SQL Integration Flow

```
Analyst completes forecast inputs in Excel
        |
Period close date reached
        |
        v
VBA Macro triggers SQL push
        |
  ┌─────┴──────────────────────────────┐
  |     SQL writes to:                 |
  |     · QlikSense data layer         |
  |     · SAP HANA                     |
  |     · Anaplan                      |
  |     · Oracle EBS                   |
  └────────────────────────────────────┘
        |
        v
All platforms reflect same forecast numbers
        |
        v
SQL Sync Log updated in Excel (timestamp + status)
```

---

## VBA Macro Overview

```vba
' Core macro tasks on period close:
Sub SyncForecastToAllPlatforms()
    ' 1. Validate all forecast inputs (no blanks, within tolerance)
    Call ValidateInputs
    
    ' 2. Build SQL INSERT/UPDATE statements from Excel data
    Call BuildSQLStatements
    
    ' 3. Connect to each platform via ODBC
    Call ConnectQlikSense
    Call ConnectSAPHANA
    Call ConnectAnaplan
    Call ConnectOracle
    
    ' 4. Execute SQL push to all platforms
    Call ExecuteSyncAllPlatforms
    
    ' 5. Log sync result to Sheet 6
    Call LogSyncResult
    
    MsgBox "Forecast sync complete. All platforms updated."
End Sub
```

---

## Impact Metrics

| Metric | Before | After |
|---|---|---|
| Data entry points | Multiple (SharePoint, server, forms) | Single (Excel) |
| Platform sync | Manual re-entry in each system | Automated SQL push |
| Submission errors | High (manual re-entry) | Near zero |
| Process consistency | Varied by analyst | Standardised |
| Platforms updated | Per-platform manually | 4 simultaneously |

---

> Note: All data in this repo uses anonymised sample data. Deloitte-specific SQL connections, credentials, and configuration are confidential.

---

*Part of the [Sughosh Anney Finance × AI Portfolio](https://github.com/sughosh-anney)*
