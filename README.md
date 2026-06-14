# Pharmaceutical Stability Dashboard

## Overview
Excel-based dashboard developed for a pharmaceutical stability laboratory to automate 
status tracking, cross-sheet data integration, and weekly reporting.

## Features
- Cross-sheet formula automation using INDEX/MATCH between WEEKLY and INVOICE sheets
- Automatic Status Final population in Invoice tracker using Stability No. + Timepoint as key
- Auxiliary columns for dynamic filtering: Year, Week, Status Clean, OOS Classification
- Pivot tables with slicers for interactive analysis by client and status
- Protected dashboard sheet accessible only to authorized users

## Tools Used
- Microsoft Excel (advanced formulas, pivot tables, slicers)
- Python (openpyxl) for formula generation and file structure

## Key Formulas Applied
- `INDEX/MATCH` for cross-sheet lookups
- `IF/ISNUMBER/FIND` for OOS classification (Open/Closed/No OOS)
- `WEEKNUM` and `YEAR` for time-based filtering
- `TRIM/CLEAN` for data normalization in key columns

## Note
All data in this file is fictitious and generated for portfolio demonstration purposes only.
