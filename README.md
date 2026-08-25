# ExcelAuto

**AI Data Cleaning Agent for Microsoft Excel**

ExcelAuto is a Microsoft Excel Add-in that automates data cleaning and quality analysis directly inside your spreadsheet. It runs as a task pane within Excel, powered by the Office.js API, and works on any selected data range without requiring any external tools, databases, or internet connection.

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
- [Quality Report](#quality-report)
- [Build for Production](#build-for-production)
- [Sideloading in Excel](#sideloading-in-excel)
- [Known Limitations](#known-limitations)
- [License](#license)

---

## Overview

Data quality is one of the most common and time-consuming problems faced by analysts, operations teams, and anyone who regularly works with spreadsheets. ExcelAuto solves this by bringing smart, automated cleaning directly into Excel - no copy-pasting into external tools, no manual formulas, no VBA scripting.

Select your data range, choose your cleaning preferences, and let ExcelAuto handle the rest. It detects issues, fixes what it can automatically, and clearly flags what needs manual attention.

---

## Features

### Analyze Data

- Scans the selected range and produces a full Data Quality Analysis
- Counts blank cells, duplicate rows, extra whitespace, capitalization inconsistencies, data type mismatches, date format issues, invalid dates, and statistical outliers
- Calculates a Data Quality Score (0-100%) based on the detected issues
- Builds a per-column profile showing data type, completeness, and unique value count

### Clean Data

- Trims leading and trailing whitespace from all string cells
- Standardizes text capitalization for categorical columns
- Converts numeric text (e.g. "42") to actual Number type
- Normalizes date formats across the dataset
- Replaces blank and missing cells using the user-selected strategy
- Reports a full Cleaning Summary after each run

### Quality Report

- Generates a before-vs-after report showing the Quality Score before and after cleaning
- Clearly distinguishes between Issues Detected, Issues Cleaned, and Issues Remaining
- Includes a Manual Review section for invalid dates and statistical outliers that cannot be auto-fixed
- Shows improvement in percentage points
- Can be downloaded as a Word (.doc) file for sharing or documentation

### Loading Bar

- A progress bar with a status label is shown during Analyze, Clean, and Report operations so users always know something is running

---

## Tech Stack

| Layer           | Technology                             |
| --------------- | -------------------------------------- |
| Add-in Runtime  | Office.js (Microsoft Office JS API v1) |
| Frontend        | HTML5, CSS3, Vanilla JavaScript (ES6+) |
| Build Tool      | Webpack 5                              |
| Transpiler      | Babel                                  |
| Package Manager | npm / Node.js                          |
| Target Host     | Microsoft Excel (Desktop and Online)   |

---

## Project Structure

```
ExcelAuto/
├── assets/                  # Add-in icons and images
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-80.png
│   └── logo-filled.png
├── dist/                    # Production build output (generated)
├── src/
│   ├── commands/
│   │   ├── commands.html
│   │   └── commands.js
│   └── taskpane/
│       ├── taskpane.html    # Add-in UI structure
│       ├── taskpane.css     # Add-in styling
│       └── taskpane.js      # All add-in logic
├── .eslintrc.json
├── babel.config.json
├── manifest.xml             # Office Add-in manifest (required for sideloading)
├── package.json
├── webpack.config.js
└── README.md
```

---

## Prerequisites

Before running ExcelAuto, make sure you have the following installed:

- [Node.js](https://nodejs.org/) (version 16 or higher recommended)
- npm (comes with Node.js)
- Microsoft Excel (Desktop - Windows or Mac) or access to Excel Online
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

### Development Mode

Start the local development server:

```bash
npm run dev-server
```

This starts a webpack dev server (typically on `https://localhost:3000`). Keep this running while using the add-in in Excel.

Then sideload the add-in into Excel (see [Sideloading in Excel](#sideloading-in-excel)).

### Alternative - Start with Excel

If you have the Office Add-in CLI tools configured:

```bash
npm start
```

This will attempt to start the dev server and open Excel automatically.

---

## How to Use

1. Open Microsoft Excel and load your spreadsheet.
2. Open the ExcelAuto task pane from the Insert > Add-ins menu (after sideloading).
3. Select the data range you want to analyze or clean (including the header row).
4. Use the three main buttons:

**Analyze Data**
Scans your selected range and shows a full data quality breakdown in the task pane. No changes are made to your data.

**Clean Data**
Applies all cleaning operations to your selected range based on the chosen Cleaning Options. A summary of changes is displayed after cleaning.

**Quality Report**
Generates a detailed before-and-after quality report. If you have run cleaning, it shows the score improvement. The report can be downloaded as a Word document using the download button in the report header.

---

## Cleaning Options

Before running Clean Data, you can configure how ExcelAuto handles missing and blank cells:

| Option              | Behavior                                                    |
| ------------------- | ----------------------------------------------------------- |
| Smart Replace       | Fills numeric blanks with column mean, text blanks with N/A |
| Replace with 0      | Fills all blank cells with 0                                |
| Replace with N/A    | Fills all blank cells with the text "N/A"                   |
| Replace with Mean   | Fills numeric blanks with the column mean value             |
| Replace with Median | Fills numeric blanks with the column median value           |
| Leave Blank         | Does not modify blank cells                                 |

There is also a separate Data Type Correction button that converts numeric text values to proper Number types across the entire selected range.

---

## Quality Report

The Quality Report is generated fresh each time you click the Quality Report button. It reflects the current state of the dataset.

### Report Sections

| Section          | Description                                                    |
| ---------------- | -------------------------------------------------------------- |
| Dataset Summary  | Total rows, columns, and cells in the selected range           |
| Quality Score    | Score before cleaning, after cleaning, and the improvement     |
| Issues Detected  | All issues found during the most recent analysis               |
| Cleaning Summary | Every action taken during the most recent clean operation      |
| Issues Remaining | Issues still present after cleaning                            |
| Manual Review    | Invalid dates and outliers that require human review           |
| Final Status     | Overall verdict - clean, partially cleaned, or review required |

### Downloading the Report

Click the download button (shown on the right side of the report header) to download the full report as a Word (.doc) file. The file is named:

```
ExcelAuto_Quality_Report_YYYY-MM-DD.doc
```

---

## Build for Production

To build the production-ready bundle:

```bash
npm run build
```

Output is placed in the `dist/` directory. The manifest.xml points to `https://localhost:3000` by default - update the URLs in `manifest.xml` if deploying to a hosted server.

---

## Sideloading in Excel

### Windows (Desktop)

1. Go to **Insert > Add-ins > My Add-ins**.
2. Click **Upload My Add-in**.
3. Browse to and select `manifest.xml` from the project root.
4. The ExcelAuto task pane will appear in the ribbon.

### Mac (Desktop)

1. Copy `manifest.xml` to: `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/`
2. Restart Excel.
3. Open **Insert > Add-ins** to find ExcelAuto.

### Excel Online

1. Go to **Insert > Add-ins > Upload and Manage My Add-ins**.
2. Upload `manifest.xml`.

> The development server (`npm run dev-server`) must be running whenever you use the add-in in development mode.

---

## Known Limitations

- The add-in operates only on the currently selected range. Always select the full dataset including the header row before running any operation.
- Outliers are detected and flagged but are never automatically removed. They must be reviewed and handled manually.
- Invalid dates (e.g. "Feb 30") are detected and reported but are not modified by the cleaner. They require manual correction.
- Duplicate row detection reports the count of duplicates but does not delete entire worksheet rows automatically (to avoid accidentally removing data outside the selected range).
- The downloaded Quality Report (.doc) is a Word-compatible HTML file. It opens correctly in Microsoft Word and LibreOffice Writer.

---

## License

This project is licensed under the MIT License.

---

_ExcelAuto - Built for analysts, data teams, and anyone who lives in spreadsheets._
