# Fleet Cost Controller — Excel VBA Portfolio Project

A fully functional fleet controlling workbook built in Excel and VBA,
modelled on a vehicle rental fleet operation.

## What it does

- Tracks 300 simulated rental transactions across 6 vehicle classes,
  3 months, and 3 branches
- Calculates actual revenue, total costs, contribution margin,
  and rental days by vehicle class and month
- Compares actuals against budget with automated variance flagging
- Exports a dated management KPI dashboard as PDF in one click

## Structure

| Sheet | Purpose |
|---|---|
| RAW_DATA | Single source of truth — all transaction data |
| BUDGET | Manual planning assumptions by vehicle class and month |
| SUMMARY | SUMIFS-driven actuals by vehicle class and month |
| VARIANCE | Budget vs actual comparison with automated flags |
| KPI_DASHBOARD | Management-facing output with PDF export button |

## VBA Modules

| Module | Purpose |
|---|---|
| mod_ValidateData | Checks data quality and highlights errors |
| mod_RefreshData | Forces correct calculation order across all sheets |
| mod_FormatReports | Applies consistent formatting across all report sheets |
| mod_ExportDashboard | Exports KPI dashboard as dated PDF |
| mod_Master | Calls all modules in sequence from a single button |

## Tools Used

Excel, VBA, Power Query, SUMIFS, Conditional Formatting

## Author

Pragya Prasad
pragyaprasad2000@gmail.com