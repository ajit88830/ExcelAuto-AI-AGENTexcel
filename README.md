# ExcelAuto

**AI Data Cleaning Agent for Microsoft Excel**

ExcelAuto is a Microsoft Excel Task Pane Add-in that automates data cleaning and quality analysis directly inside your spreadsheet. It runs as a side panel within Excel, powered by the Office.js API, and works on any selected data range without requiring any external tools, databases, or internet connection during processing.

---

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Tech Stack](#tech-stack)
- [Project Structure](#project-structure)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Running the Add-in](#running-the-add-in)
- [How to Use](#how-to-use)
- [Cleaning Options](#cleaning-options)
- [Smart Primary Key Detection](#smart-primary-key-detection)
- [Quality Report](#quality-report)
- [Data Quality Score](#data-quality-score)
- [Build for Production](#build-for-production)
- [Sideloading in Excel](#sideloading-in-excel)
- [Known Limitations](#known-limitations)
- [License](#license)

---

## Overview

Data quality is one of the most common and time-consuming problems faced by analysts, operations teams, and anyone who regularly works with spreadsheets. ExcelAuto solves this by bringing smart, automated cleaning directly into Excel - no copy-pasting into external tools, no manual formulas, no VBA scripting.

Select your data range, configure your cleaning preferences, and let ExcelAuto handle the rest. It detects issues, fixes what it can automatically, flags what needs manual attention, and generates a full before-and-after quality report.

---

## Features

### Analyze Data
- Scans the selected range and produces a full Data Quality Analysis
- Detects blank cells, missing values, duplicate rows, extra whitespace, capitalization inconsistencies, data type mismatches, date format issues, invalid dates, and statistical outliers
- Calculates a Data Quality Score (0 to 100 percent) using a weighted severity model
- Builds a per-column profile showing data type, completeness, unique value count, and extra-space count
- Auto-detects the Primary Key column using a heuristic scoring engine

### Clean Data
- Trims leading and trailing whitespace from all string cells
- Standardizes text capitalization for categorical columns (Title Case)
- Converts numeric text strings (e.g. "42") to actual Number type
- Normalizes date formats across the dataset to Excel serial format
- Replaces blank and missing cells using the user-selected strategy
- Removes duplicate rows based on the selected or auto-detected Primary Key column
- Reports a full Cleaning Summary after each run

### Quality Report
- Generates a before-vs-after quality report showing score improvement in percentage points
- Distinguishes clearly between Issues Detected, Issues Cleaned, and Issues Remaining
- Includes a Manual Review section for invalid dates and statistical outliers that cannot be auto-fixed
- Final status verdict: Clean, Partially Cleaned, or Review Required
- Download the full report as a Word (.doc) file using the download button in the report header

### Smart Primary Key Detection
- Heuristic engine automatically identifies the most likely primary key column
- Scores columns based on header keywords (ID, Code, Key, No, Index), uniqueness percentage, and blank cell count
- User can override the auto-detected selection via a dropdown in Cleaning Options
- Options: Auto-Detect (Heuristics), Entire Row (Exact Match), or any specific column header

### Loading Bar
- Animated progress bar shown during Analyze, Clean, and Quality Report operations
- Always hides cleanly after completion or on error, using try/catch/finally

### Download Report
- Downloads the Quality Report as a Word-compatible .doc file
- File is named: ExcelAuto_Quality_Report_YYYY-MM-DD.doc
- Opens correctly in Microsoft Word and LibreOffice Writer

---

## Tech Stack

| Layer           | Technology                              |
|-----------------|-----------------------------------------|
| Add-in Runtime  | Office.js (Microsoft Office JS API v1)  |
| Frontend        | HTML5, CSS3, Vanilla JavaScript (ES6+)  |
| Build Tool      | Webpack 5                               |
| Transpiler      | Babel 7                                 |
| Polyfills       | core-js + regenerator-runtime           |
| Dev Server      | webpack-dev-server                      |
| Office CLI      | office-addin-debugging                  |
| Package Manager | npm / Node.js                           |
| Target Host     | Microsoft Excel (Desktop and Online)    |

---

## Project Structure

```
ExcelAuto/
├── assets/                  Add-in icons and images
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-80.png
│   └── logo-filled.png
├── dist/                    Production build output (auto-generated)
├── src/
│   ├── commands/
│   │   ├── commands.html
│   │   └── commands.js
│   └── taskpane/
│       ├── taskpane.html    Add-in UI structure
│       ├── taskpane.css     Add-in styling and animations
│       └── taskpane.js      All add-in logic (2600+ lines)
├── .eslintrc.json
├── babel.config.json
├── manifest.xml             Office Add-in manifest
├── package.json
├── webpack.config.js
├── README.md
└── ExcelAuto_Implementation_Document.doc
```

---

## Prerequisites

Before running ExcelAuto, make sure you have the following installed:

- Node.js (version 16 or higher recommended)
- npm (comes with Node.js)
- Microsoft Excel Desktop (Windows or Mac) or access to Excel Online
- A modern web browser (for the dev server)

---

## Installation

1. Clone or download this repository:

```bash
git clone https://github.com/your-username/ExcelAuto.git
cd ExcelAuto
```

2. Install dependencies:

```bash
npm install
```

---

## Running the Add-in

### Recommended - Start with Excel (Development)

This is the correct way to run the project. It starts the server AND automatically sideloads the add-in into a new Excel window:

```bash
npm start
```

A new Excel window will open with ExcelAuto already loaded in the task pane.

### Server Only

If you only want to start the webpack dev server without launching Excel:

```bash
npm run dev-server
```

Note: Using this command alone will not show the add-in in Excel. You must manually sideload the manifest separately.

### Stop Debugging

```bash
npm run stop
```

---

## How to Use

1. Open Microsoft Excel with the ExcelAuto add-in loaded (via npm start).
2. Open or paste your dataset into a worksheet.
3. Select the full data range including the header row (e.g. A1:J500).
4. Use the three main action buttons in the task pane:

**Analyze Data**
Scans your selected range and shows a full data quality breakdown. No changes are made to your data. The Primary Key column is also auto-detected and shown in the dropdown.

**Clean Data**
Applies all cleaning operations to your selected range based on your chosen Cleaning Options. A Cleaning Summary is displayed after completion.

**Quality Report**
Generates a detailed before-and-after quality report. If cleaning was run beforehand, it shows the score improvement. Click the download button (top-right of the report header) to save the report as a Word file.

---

## Cleaning Options

Before running Clean Data, configure how ExcelAuto handles missing cells and duplicates:

### Missing Value Action

| Option           | Behavior                                                      |
|------------------|---------------------------------------------------------------|
| Smart Replace    | Fills numeric blanks with column mean, text blanks with N/A   |
| Replace with 0   | Fills all blank cells with 0                                  |
| Replace with N/A | Fills all blank cells with the text N/A                       |
| Replace with Mean| Fills numeric blanks with the column mean value               |
| Replace with Median | Fills numeric blanks with the column median value          |
| Leave Blank      | Does not modify blank cells                                   |

### Data Type Correction

A separate Convert button converts numeric text values (e.g. "42", "3.14") to proper Number type across the entire selected range.

---

## Smart Primary Key Detection

ExcelAuto uses a heuristic scoring engine to automatically identify the most likely Primary Key column in your dataset. This is used to detect duplicates intelligently - instead of comparing entire rows, it checks whether the primary key value repeats.

### How Auto-Detection Works

Each column is scored based on three criteria:

| Criterion                                              | Score |
|--------------------------------------------------------|-------|
| Header name contains "ID", "Code", "Key", "No", "Index" | +50   |
| Column has zero blank cells                            | +20   |
| 95 percent or more of values are unique               | +30   |

The column with the highest score above 60 is automatically selected as the Primary Key and labeled "(Auto-detected)" in the dropdown.

### Dropdown Options

| Option                  | Behavior                                                    |
|-------------------------|-------------------------------------------------------------|
| Auto-Detect (Heuristics)| Uses the heuristic engine to find the primary key           |
| Entire Row (Exact Match)| Matches all columns - only perfectly identical rows are duplicates |
| Any column header       | Uses that specific column as the unique identifier          |

---

## Quality Report

The Quality Report is generated fresh each time you click Quality Report. It always reflects the current state of the data.

### Report Sections

| Section          | Description                                                           |
|------------------|-----------------------------------------------------------------------|
| Dataset Summary  | Total rows, columns, and cells in the selected range                  |
| Quality Score    | Score before cleaning, after cleaning, and the improvement in points  |
| Issues Detected  | All issues found during the most recent analysis                      |
| Cleaning Summary | Every action taken during the most recent clean operation             |
| Issues Remaining | Issues still present after cleaning                                   |
| Manual Review    | Invalid dates and outliers that require human review                  |
| Final Status     | Overall verdict: Clean, Partially Cleaned, or Review Required         |

### Final Status Logic

- All remaining issues resolved: "Data is clean."
- Only outliers remain: "Cleaning completed. X outliers remain for review."
- Other issues remain: "Data cleaning completed with issues requiring review." with a list of remaining issues.

### Download Button

Click the download button at the top-right corner of the report header to save the full report as:

```
ExcelAuto_Quality_Report_YYYY-MM-DD.doc
```

---

## Data Quality Score

The score is calculated using a weighted severity model:

| Issue Type        | Weight per Issue | Severity     |
|-------------------|-----------------|--------------|
| Invalid Dates     | 3.0             | High         |
| Duplicate Rows    | 3.0             | High         |
| Missing Values    | 2.5             | High         |
| Blank Cells       | 2.5             | High         |
| Data Type Issues  | 2.0             | Medium-High  |
| Category Issues   | 1.5             | Medium       |
| Date Format Issues| 1.5             | Medium       |
| Outliers          | 0.5             | Low          |
| Extra Spaces      | 0.3             | Low          |

Score thresholds: 90 percent or above = Good (green), 70 to 89 percent = Fair (yellow), below 70 percent = Poor (red).

---

## Build for Production

To build the production-ready bundle:

```bash
npm run build
```

Output is placed in the dist/ directory. The manifest.xml points to https://localhost:3000 by default. Update the URLs in manifest.xml if deploying to a hosted server.

---

## Sideloading in Excel

### Windows Desktop

1. Make sure npm start or npm run dev-server is running.
2. In Excel, go to Insert > Add-ins > My Add-ins.
3. Click Upload My Add-in.
4. Browse to and select manifest.xml from the project root.
5. The ExcelAuto task pane will appear.

### Mac Desktop

1. Copy manifest.xml to: ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
2. Restart Excel.
3. Open Insert > Add-ins to find ExcelAuto.

### Excel Online

1. Go to Insert > Add-ins > Upload and Manage My Add-ins.
2. Upload manifest.xml.

Note: If you see an "Add-in is no longer available" error in Excel, it means the dev server was stopped while Excel still had the old session registered. Run npm start again to restart cleanly.

---

## Known Limitations

- The add-in operates only on the currently selected range. Always select the full dataset including the header row before running any operation.
- Outliers are detected and flagged but are never automatically removed. They require manual review before any modification.
- Invalid dates are detected and reported but are not modified by the cleaner. They require manual correction.
- Duplicate row detection is based on the Primary Key column (or entire row in fallback mode). It does not delete entire worksheet rows to avoid accidentally removing data outside the selected range.
- The downloaded Quality Report (.doc) is a Word-compatible HTML file. It opens correctly in Microsoft Word and LibreOffice Writer but is not a native DOCX binary file.
- All processing runs locally in the browser runtime inside Excel. There is no backend server, database, or cloud API involved.

---

## License

This project is licensed under the MIT License.

---

*ExcelAuto - Built for analysts, data teams, and anyone who lives in spreadsheets.*
