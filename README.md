# ASD with Live PivotTables & Email Reporting

An interactive sales dashboard built in Excel / Google Sheets, paired with a Google Apps Script
automation that turns it into a self-distributing reporting tool.
Automated Sales Dashboard (ASD) with Live PivotTables & Email Reporting

## Overview

Built an interactive sales dashboard using SUMIFS/SUMPRODUCT-driven pivot summaries and dropdown
filters that recalculate KPI cards and charts in real time. Layered on a Google Apps Script
automation that exports the dashboard as a formatted PDF, archives it to Drive, and emails it to
stakeholders, either on a weekly schedule or on demand via an in-sheet button.

## Features

- **Pivot-style cross-tabs** (Region x Category revenue, monthly trend, sales rep leaderboard)
  built with `SUMIFS` / `COUNTIFS` so they recalculate automatically as raw data changes
- **Dropdown filters** (Region, Category) driving every KPI card via `SUMPRODUCT` array formulas
- **Native charts** (bar, pie, line) wired to the pivot summaries
- **One-click PDF export**: a checkbox "button" on the Dashboard tab triggers an on-demand send
- **Scheduled reporting**: a weekly time-based trigger sends the dashboard automatically
- **Recipient management**: stakeholders are maintained as a simple table on their own tab,
  no hardcoded emails
- **Drive archiving**: every generated PDF is saved to a dedicated Drive folder for a running
  history of reports

## Tech Stack

- Google Sheets / Excel (formulas: `SUMIFS`, `SUMPRODUCT`, `COUNTIFS`, `IFERROR`)
- Google Apps Script (`GmailApp`, `DriveApp`, `UrlFetchApp`, `SpreadsheetApp`)
- Data validation & conditional formatting for the interactive filter UI

## Files

| File | Description |
|---|---|
| `sales_dashboard.xlsx` | The spreadsheet template with Raw Data, Pivot Summary, Dashboard, and Read Me tabs |
| `dashboard_pdf_dispatch.gs` | Apps Script for PDF export, Drive archiving, and email dispatch (scheduled and on-demand) |

## Setup

1. Open the spreadsheet in Google Sheets, then go to Extensions > Apps Script.
2. Paste in `dashboard_pdf_dispatch.gs`, replacing the default `Code.gs` contents.
3. Edit the `CONFIG` block at the top (folder name, schedule, sheet names) if needed.
4. Run `setup()` once and authorize Gmail/Drive/Sheets access. This creates:
   - a **Recipients** tab for stakeholder emails
   - a checkbox "send now" button on the **Dashboard** tab
   - a weekly scheduled trigger
5. Replace the sample rows on the Recipients tab with real stakeholder emails.

## How It Works

- **Filters to KPIs.** Each KPI card uses a `SUMPRODUCT` formula with an embedded `IF`, letting
  the Region/Category dropdowns mean "ignore this filter" when set to "All." This is
  Excel/Sheets' native equivalent of an array formula.
- **Pivot Summary tab.** Cross-tab and trend tables are pure `SUMIFS`/`COUNTIFS`, giving the
  same output as a native PivotTable while staying formula-driven and portable across Excel and
  Sheets.
- **Automation.** An installable `onEdit` trigger watches the Dashboard checkbox. Ticking it (or
  the weekly time trigger firing) calls `sendDashboardPDF()`, which exports the Dashboard tab as
  a landscape PDF, saves it to Drive, and emails it with `GmailApp`.

