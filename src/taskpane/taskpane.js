/* ============================================================
   ExcelAuto - Excel AI Data Cleaning Agent
   Fully updated taskpane.js
   ============================================================ */

Office.onReady(function (info) {

    if (info.host !== Office.HostType.Excel) {
        return;
    }

    console.log("ExcelAuto loaded successfully.");

    // ---------------------------------------------------------
    // DOM ELEMENTS
    // ---------------------------------------------------------

    const analyzeButton = document.getElementById("analyzeButton");
    const cleanButton = document.getElementById("cleanButton");
    const reportButton = document.getElementById("reportButton");
    const convertNumericTextBtn = document.getElementById("convertNumericTextBtn");
    
    const primaryKeySelect = document.getElementById("primaryKeySelect");
    let detectedPrimaryKeyIndex = -1;

    const missingValueOption =
        document.getElementById("blankReplacementMode");

    const statusElement =
        document.getElementById("status");

    const resultElement =
        document.getElementById("result");

    const rangeAddressElement =
        document.getElementById("rangeAddress");

    const rowCountElement =
        document.getElementById("rowCount");

    const columnCountElement =
        document.getElementById("columnCount");

    const blankCountElement =
        document.getElementById("blankCount");

    const duplicateCountElement =
        document.getElementById("duplicateCount");

    const cleaningSummaryElement =
        document.getElementById("cleaningSummary");

    const issuesContainer =
        document.getElementById("issuesContainer");

    const outliersContainer =
        document.getElementById("outliersContainer");

    const columnProfileElement =
        document.getElementById("columnProfile");


    // Global State for Mixed Data Types
    let currentMixedDataCells = [];
    let currentAnalysisRangeAddress = "";
    // Global State for Invalid Dates (populated during analyzeData, used by quality report)
    let currentInvalidDateCells = [];
    // Global State for outlier count + category issues (populated during analyzeData, used by quality report)
    let currentOutlierCount = 0;
    let currentCategoryIssuesCount = 0;
    // Global State for Quality Report: before/after scores, issue snapshots, cleaning summary
    let currentQualityScore = null;          // Score from most recent analyzeData
    let currentAnalysisSnapshot = null;       // Issue counts from most recent analyzeData
    let beforeCleaningScore = null;           // Score captured just BEFORE cleanData runs
    let beforeCleaningSnapshot = null;        // Issue counts captured just BEFORE cleanData runs
    let lastCleaningSummary = null;           // Cleaning counters from most recent cleanData

    // ---------------------------------------------------------
    // IMPORTANT FIX
    // setStatus() was missing in your previous code.
    // ---------------------------------------------------------

    function setStatus(message, type = "normal") {

        if (!statusElement) {
            console.error("Status element not found.");
            return;
        }

        statusElement.textContent = message;

        statusElement.classList.remove(
            "success",
            "error",
            "warning"
        );

        if (type === "success") {
            statusElement.classList.add("success");
        }

        if (type === "error") {
            statusElement.classList.add("error");
        }

        if (type === "warning") {
            statusElement.classList.add("warning");
        }

        console.log("ExcelAuto Status:", message);
    }


    // ---------------------------------------------------------
    // LOADING BAR HELPERS
    // ---------------------------------------------------------

    const loadingBarWrap  = document.getElementById("loadingBarWrap");
    const loadingBarLabel = document.getElementById("loadingBarLabel");

    function showLoading(label) {
        if (loadingBarWrap) {
            loadingBarLabel.textContent = label || "";
            loadingBarWrap.style.display = "block";
        }
    }

    function hideLoading() {
        if (loadingBarWrap) {
            loadingBarWrap.style.display = "none";
            loadingBarLabel.textContent = "";
        }
    }


    // ---------------------------------------------------------
    // HELPER FUNCTIONS
    // ---------------------------------------------------------

    function isBlank(value) {

        return (
            value === null ||
            value === undefined ||
            (typeof value === "string" && value.trim() === "")
        );
    }


    function isMissingValue(value) {
        if (typeof value !== "string") {
            return false;
        }
        const text = value.trim().toLowerCase();
        return ["n/a", "na", "not available", "null", "none", "-"].includes(text);
    }


    function shouldConvertToNumber(value) {

        if (typeof value !== "string") {
            return false;
        }

        const text = value.trim();

        if (text === "") {
            return false;
        }

        // Skip IDs with leading zero (e.g., "0123") but not "0" or "0.5"
        if (/^-?0\d/.test(text)) {
            return false;
        }

        // Skip long strings of digits (e.g., Phone numbers, Account numbers)
        // Commas and decimals will prevent matching, which is intended.
        if (/^-?\d{10,}$/.test(text)) {
            return false;
        }

        // Remove currency symbols, commas, and quotes
        const cleaned = text.replace(/[$€£,"']/g, "");

        if (cleaned === "") {
            return false;
        }

        return Number.isFinite(Number(cleaned));
    }


    function calculateIQRBounds(numbers) {
        if (numbers.length < 4) return null; // Not enough data points
        const sorted = [...numbers].sort((a, b) => a - b);
        const q1Index = Math.floor(sorted.length * 0.25);
        const q3Index = Math.floor(sorted.length * 0.75);
        const q1 = sorted[q1Index];
        const q3 = sorted[q3Index];
        const iqr = q3 - q1;
        
        // Return lower and upper bounds
        return {
            lowerBound: q1 - 1.5 * iqr,
            upperBound: q3 + 1.5 * iqr
        };
    }


    function isNumeric(value) {

        if (typeof value === "number") {
            return Number.isFinite(value);
        }

        if (typeof value === "string") {
            return shouldConvertToNumber(value);
        }

        return false;
    }


    function numericValue(value) {

        if (typeof value === "number") {
            return value;
        }

        if (typeof value === "string") {
            return Number(value.trim().replace(/[$€£,]/g, ""));
        }

        return NaN;
    }


    function formatNumber(value) {

        if (!Number.isFinite(value)) {
            return "";
        }

        return Number.isInteger(value)
            ? String(value)
            : String(Number(value.toFixed(4)));
    }


    function calculateMean(values) {

        const numbers = values
            .filter(isNumeric)
            .map(numericValue)
            .filter(Number.isFinite);

        if (numbers.length === 0) {
            return null;
        }

        const total = numbers.reduce(
            (sum, value) => sum + value,
            0
        );

        return total / numbers.length;
    }


    function calculateMedian(values) {

        const numbers = values
            .filter(isNumeric)
            .map(numericValue)
            .filter(Number.isFinite)
            .sort((a, b) => a - b);

        if (numbers.length === 0) {
            return null;
        }

        const middle = Math.floor(numbers.length / 2);

        if (numbers.length % 2 === 0) {

            return (
                numbers[middle - 1] +
                numbers[middle]
            ) / 2;

        }

        return numbers[middle];
    }


    function normalizeText(value) {

        if (typeof value !== "string") {
            return value;
        }

        return value.trim();
    }


    function capitalizeText(value) {

        if (typeof value !== "string") {
            return value;
        }

        const trimmed = value.trim();

        if (trimmed === "") {
            return value;
        }

        return trimmed
            .toLowerCase()
            .replace(/\b\w/g, function (letter) {
                return letter.toUpperCase();
            });
    }


    function parseAndValidateDate(text) {
        let safeText = text.replace(/[-.]/g, "/");
        
        const months = {
            jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
            jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12
        };

        const daysInMonth = (y, m) => new Date(y, m, 0).getDate();

        let day = NaN, month = NaN, year = NaN;
        
        const parts = safeText.split(/[\/\s,]+/);
        
        if (parts.length >= 3) {
            const p0 = parts[0];
            const p1 = parts[1];
            const p2 = parts[2];

            const p0num = parseInt(p0, 10);
            const p1num = parseInt(p1, 10);
            const p2num = parseInt(p2, 10);

            if (isNaN(p0num) && months[p0.toLowerCase().substring(0,3)]) {
                month = months[p0.toLowerCase().substring(0,3)];
                day = p1num;
                year = p2num;
            } else if (isNaN(p1num) && months[p1.toLowerCase().substring(0,3)]) {
                day = p0num;
                month = months[p1.toLowerCase().substring(0,3)];
                year = p2num;
            } else {
                if (p0num > 1000) {
                    year = p0num;
                    month = p1num;
                    day = p2num;
                } else if (p2num > 1000 || (p2num >= 0 && p2num <= 99)) {
                    year = p2num;
                    if (year < 100) year += 2000;
                    if (p0num > 12) {
                        day = p0num;
                        month = p1num;
                    } else {
                        month = p0num;
                        day = p1num;
                    }
                }
            }
        }

        if (isNaN(year) || isNaN(month) || isNaN(day)) {
            return { isValid: false, reason: "Unrecognized format" };
        }

        if (month < 1 || month > 12) {
            return { isValid: false, reason: "Invalid month" };
        }
        
        if (day < 1 || day > daysInMonth(year, month)) {
            if (month === 2) {
                return { isValid: false, reason: "Invalid day for February" };
            }
            return { isValid: false, reason: "Invalid day/month" };
        }

        const dateObj = new Date(Date.UTC(year, month - 1, day));
        return { isValid: true, dateObj: dateObj };
    }

    function looksLikeDate(value) {
        if (typeof value !== "string") return false;

        const text = value.trim();
        if (text === "") return false;
        if (/^-?\d+(\.\d+)?$/.test(text)) return false;

        const res = parseAndValidateDate(text);
        if (!res.isValid) return false;

        return (
            /^\d{1,4}[-/\s.]\d{1,2}[-/\s.]\d{1,4}$/.test(text) ||
            /^\d{1,2}[-/\s.][A-Za-z]{3,9}[-/\s.]\d{2,4}$/.test(text) ||
            /^[A-Za-z]{3,9}\s\d{1,2},?\s\d{4}$/.test(text)
        );
    }

    function isInvalidDate(value) {
        if (typeof value !== "string") return false;
        
        const text = value.trim();
        if (text === "") return false;
        if (/^-?\d+(\.\d+)?$/.test(text)) return false;

        const matchesPattern = (
            /^\d{1,4}[-/\s.]\d{1,2}[-/\s.]\d{1,4}$/.test(text) ||
            /^\d{1,2}[-/\s.][A-Za-z]{3,9}[-/\s.]\d{2,4}$/.test(text) ||
            /^[A-Za-z]{3,9}\s\d{1,2},?\s\d{4}$/.test(text)
        );

        if (!matchesPattern) return false;

        const res = parseAndValidateDate(text);
        if (!res.isValid) {
            return res.reason;
        }
        return false;
    }

    function getDateFormatSignature(text) {
        return text.trim()
            .replace(/\b\d{4}\b/g, 'Y')
            .replace(/\b\d{1,2}\b/g, 'N')
            .replace(/[a-zA-Z]+/g, 'L');
    }

    function normalizeDate(value) {
        if (!looksLikeDate(value)) {
            return value;
        }

        const res = parseAndValidateDate(value);
        if (!res.isValid) {
            return value;
        }

        const dateObj = res.dateObj;
        const utcDate = Date.UTC(dateObj.getUTCFullYear(), dateObj.getUTCMonth(), dateObj.getUTCDate());
        const epoch = Date.UTC(1899, 11, 30);
        
        return (utcDate - epoch) / 86400000;
    }


    // ---------------------------------------------------------
    // DATA QUALITY SCORE (weighted)
    //
    // Severity weights per issue type (penalty points per issue):
    //   HIGH   = 3.0  — invalid dates, duplicate rows, missing values, data type errors
    //   MEDIUM = 1.5  — category inconsistencies, date format issues
    //   LOW-MED= 0.5  — outliers (may be legitimate business values)
    //   LOW    = 0.3  — extra spaces, capitalization
    //
    // The total weighted penalty is measured against totalRows
    // (not totalCells) so that even a handful of issues produce
    // a visible score drop.
    // ---------------------------------------------------------

    function calculateDataQualityScore(issues, totalRows) {
        const weights = {
            invalidDates:      3.0,
            duplicateRows:     3.0,
            blankCells:        2.5,
            missingValues:     2.5,
            formatIssues:      2.0,
            dateFormatIssues:  1.5,
            categoryIssues:    1.5,
            outliers:          0.5,
            trimIssues:        0.3,
            capitalization:    0.3
        };

        let totalWeightedPenalty = 0;
        const breakdown = {};

        for (const key in issues) {
            const count = issues[key] || 0;
            if (count <= 0) continue;
            const weight = weights[key] || 1.0;
            const penalty = count * weight;
            totalWeightedPenalty += penalty;
            breakdown[key] = { count, weight, penalty: Math.round(penalty * 100) / 100 };
        }

        // Scale against totalRows (not totalCells) for meaningful impact.
        // Cap deduction at 100 so score never goes below 0.
        const safeRows = Math.max(totalRows, 1);
        const rawDeduction = (totalWeightedPenalty / safeRows) * 100;
        const deduction = Math.min(rawDeduction, 100);
        const score = Math.round(Math.max(0, 100 - deduction));

        const totalIssues = Object.values(issues).reduce((a, b) => a + (b || 0), 0);

        return { score, totalIssues, totalWeightedPenalty, breakdown };
    }


    function valuesEqual(a, b) {

        if (isBlank(a) && isBlank(b)) {
            return true;
        }

        return String(a).trim() === String(b).trim();
    }

    function detectPrimaryKeyIndex(values, dataStartRow) {
        if (!values || values.length <= dataStartRow || values[0].length === 0) return -1;
        
        const headers = values[0].map(h => String(h).toLowerCase());
        const keywordRegex = /(id|code|key|no|index)\b/i;
        
        let bestIndex = -1;
        let bestScore = 0;
        
        for (let c = 0; c < headers.length; c++) {
            let score = 0;
            if (keywordRegex.test(headers[c])) {
                score += 50;
            }
            
            const uniqueVals = new Set();
            let blanks = 0;
            let total = 0;
            for (let r = dataStartRow; r < values.length; r++) {
                const val = values[r][c];
                total++;
                if (isBlank(val)) blanks++;
                else uniqueVals.add(String(val).trim().toLowerCase());
            }
            
            if (total > 0 && blanks === 0) {
                score += 20; 
                const uniquePercent = uniqueVals.size / total;
                if (uniquePercent > 0.95) score += 30;
            }
            
            if (score > bestScore && score > 60) {
                bestScore = score;
                bestIndex = c;
            }
        }
        return bestIndex;
    }


    function createRowKey(row, pkIndex = -1) {
        if (pkIndex >= 0 && pkIndex < row.length) {
            const val = row[pkIndex];
            if (isBlank(val)) return `BLANK_${Math.random()}`; 
            return String(val).replace(/\s+/g, " ").trim().toLowerCase();
        }

        return row
            .map(function (value) {
                if (isBlank(value)) return "";
                return String(value).replace(/\s+/g, " ").trim().toLowerCase();
            })
            .join("||");
    }


    // ---------------------------------------------------------
    // GET SELECTED RANGE
    // ---------------------------------------------------------

    async function getSelectedData() {
        return Excel.run(async function (context) {
            const range = context.workbook.getSelectedRange();
            range.load([
                "address",
                "values",
                "rowCount",
                "columnCount",
                "text",
                "numberFormat",
                "formulas",
                "rowIndex",
                "columnIndex",
                "worksheet"
            ]);
            await context.sync();

            let headerValues = [];
            let dataStartRow = 0;

            if (range.rowIndex > 0) {
                // User selected only the data body, header is just above
                const headerRange = range.worksheet.getRangeByIndexes(range.rowIndex - 1, range.columnIndex, 1, range.columnCount);
                headerRange.load("values");
                await context.sync();
                headerValues = headerRange.values[0];
                dataStartRow = 0; // entire selection is data
            } else {
                // User included the header in selection
                headerValues = range.values[0];
                dataStartRow = 1; // data starts on the second row of selection
            }

            return {
                range: range,
                address: range.address,
                values: range.values,
                text: range.text,
                numberFormat: range.numberFormat,
                formulas: range.formulas,
                rowCount: range.rowCount,
                columnCount: range.columnCount,
                headerValues: headerValues,
                dataStartRow: dataStartRow
            };
        });
    }


    // ---------------------------------------------------------
    // ANALYZE DATA
    // ---------------------------------------------------------

    async function analyzeData() {

        try {

            setStatus("Analyzing selected data...");

            const data = await getSelectedData();

            const values = data.values;

            if (!values || values.length === 0) {

                setStatus(
                    "Please select a dataset first.",
                    "warning"
                );

                return;
            }


            // -------------------------------------------------
            // BASIC COUNTS
            // -------------------------------------------------

            let blankCount = 0;
            let blankCols = new Set();
            let missingNaCount = 0;
            let missingNaCols = new Set();

            for (let r = data.dataStartRow; r < values.length; r++) {
                for (let c = 0; c < values[r].length; c++) {
                    if (isBlank(values[r][c])) {
                        blankCount++;
                        blankCols.add(c);
                    } else if (isMissingValue(values[r][c])) {
                        missingNaCount++;
                        missingNaCols.add(c);
                    }
                }
            }

            // -------------------------------------------------
            // PRIMARY KEY DETECTION & UI UPDATE
            // -------------------------------------------------
            
            detectedPrimaryKeyIndex = detectPrimaryKeyIndex(values, data.dataStartRow);
            
            if (primaryKeySelect && values[0]) {
                const currentVal = primaryKeySelect.value;
                primaryKeySelect.innerHTML = `
                    <option value="auto">Auto-Detect (Heuristics)</option>
                    <option value="-1">Entire Row (Exact Match)</option>
                `;
                
                values[0].forEach((header, index) => {
                    const isDetected = (index === detectedPrimaryKeyIndex);
                    const label = isDetected ? `${header} (Auto-detected)` : header;
                    primaryKeySelect.innerHTML += `<option value="${index}">${label}</option>`;
                });
                
                // Restore selection if valid
                if (Array.from(primaryKeySelect.options).some(opt => opt.value === currentVal)) {
                    primaryKeySelect.value = currentVal;
                }
            }
            
            // Determine the actual index to use for duplicate detection
            let currentPkIndex = -1;
            if (primaryKeySelect) {
                if (primaryKeySelect.value === "auto") currentPkIndex = detectedPrimaryKeyIndex;
                else currentPkIndex = parseInt(primaryKeySelect.value, 10);
            }

            // -------------------------------------------------
            // DUPLICATE ROWS
            // -------------------------------------------------

            const seenRows = new Set();
            let duplicateCount = 0;

            for (let r = data.dataStartRow; r < values.length; r++) {
                const key = createRowKey(values[r], currentPkIndex);
                if (seenRows.has(key)) {
                    duplicateCount++;
                } else {
                    seenRows.add(key);
                }
            }


            // -------------------------------------------------
            // UPDATE BASIC INFORMATION
            // -------------------------------------------------

            if (rangeAddressElement) {
                rangeAddressElement.textContent =
                    data.address;
            }

            if (rowCountElement) {

                rowCountElement.textContent =
                    Math.max(data.rowCount - 1, 0);
            }

            if (columnCountElement) {

                columnCountElement.textContent =
                    data.columnCount;
            }

            if (blankCountElement) {

                blankCountElement.textContent =
                    blankCount;
            }

            if (duplicateCountElement) {

                duplicateCountElement.textContent =
                    duplicateCount;
            }


            // -------------------------------------------------
            // DETECT ISSUES
            // -------------------------------------------------

            const issues = [];
            const headers = data.headerValues || [];

            if (blankCount > 0) {
                const colNames = Array.from(blankCols).map(c => headers[c] || `Col ${c+1}`).join(", ");
                issues.push(`
                    <strong>⚠ Blank Cells</strong><br><br>
                    Column: ${colNames}<br>
                    ${blankCount} truly empty cell(s) detected
                `);
            }

            if (missingNaCount > 0) {
                const colNames = Array.from(missingNaCols).map(c => headers[c] || `Col ${c+1}`).join(", ");
                issues.push(`
                    <strong>⚠ Missing/NA Values</strong><br><br>
                    Column: ${colNames}<br>
                    ${missingNaCount} value(s) like N/A, NA, etc. detected
                `);
            }

            if (duplicateCount > 0) {
                issues.push(`
                    <strong>⚠ Duplicate Rows</strong><br><br>
                    ${duplicateCount} completely duplicated record(s) detected
                `);
            }


            // Check unnecessary spaces

            let trimIssues = 0;
            let trimCols = new Set();

            for (let r = data.dataStartRow; r < values.length; r++) {
                for (let c = 0; c < values[r].length; c++) {
                    const value = values[r][c];
                    if (
                        typeof value === "string" &&
                        value !== value.replace(/\s+/g, " ").trim()
                    ) {
                        trimIssues++;
                        trimCols.add(c);
                    }
                }
            }


            if (trimIssues > 0) {
                const colNames = Array.from(trimCols).map(c => headers[c] || `Col ${c+1}`).join(", ");
                issues.push(`
                    <strong>⚠ Extra Spaces</strong><br><br>
                    Column: ${colNames}<br>
                    ${trimIssues} cell(s) with unnecessary spaces detected
                `);
            }

            let formatIssues = 0;
            let dateIssues = 0;
            let formatCols = new Set();
            let dateCols = new Set();
            
            const dateStatsByCol = {};
            const numericValuesByCol = {};
            const categoryVariationsByCol = {};
            const localMixedDataCells = [];
            const localInvalidDateCells = [];
            currentAnalysisRangeAddress = data.address;

            for (let r = data.dataStartRow; r < values.length; r++) {
                for (let c = 0; c < values[r].length; c++) {
                    const value = values[r][c];
                    const textValue = data.text && data.text[r] ? data.text[r][c] : String(value);
                    const formulaValue = data.formulas && data.formulas[r] ? String(data.formulas[r][c]) : "";
                    const isFormula = formulaValue.trim().startsWith("=");
                    
                    if (!dateStatsByCol[c]) {
                        dateStatsByCol[c] = { validCount: 0, invalidCount: 0, formats: new Set() };
                    }
                    if (!numericValuesByCol[c]) {
                        numericValuesByCol[c] = [];
                    }
                    const dStats = dateStatsByCol[c];

                    if (typeof value === "string") {
                        if (looksLikeDate(value)) {
                            dStats.validCount++;
                            dStats.formats.add(getDateFormatSignature(value));
                            dateIssues++;
                            dateCols.add(c);
                        } else if (isInvalidDate(value)) {
                            const reason = isInvalidDate(value);
                            dStats.invalidCount++;
                            dateIssues++;
                            dateCols.add(c);
                            localInvalidDateCells.push({ r, c, val: value, reason: typeof reason === 'string' ? reason : 'Invalid date' });
                        } else if (shouldConvertToNumber(value)) {
                            formatIssues++;
                            formatCols.add(c);
                            const num = parseFloat(value.replace(/[^0-9.-]/g, ""));
                            if (!isNaN(num)) numericValuesByCol[c].push({ r, num, text: textValue, isFormula });
                            localMixedDataCells.push({ r, c, val: value });
                        } else if (isNumeric(value) && !isBlank(value) && !isMissingValue(value)) {
                            const num = parseFloat(value.replace(/[^0-9.-]/g, ""));
                            if (!isNaN(num)) numericValuesByCol[c].push({ r, num, text: textValue, isFormula });
                            localMixedDataCells.push({ r, c, val: value });
                        } else if (!isBlank(value) && !isMissingValue(value)) {
                            const trimmed = value.trim();
                            if (trimmed !== "") {
                                const lower = trimmed.toLowerCase();
                                if (!categoryVariationsByCol[c]) {
                                    categoryVariationsByCol[c] = {};
                                }
                                if (!categoryVariationsByCol[c][lower]) {
                                    categoryVariationsByCol[c][lower] = new Set();
                                }
                                categoryVariationsByCol[c][lower].add(trimmed);
                            }
                        }
                    } else if (typeof value === "number") {
                        const numFormat = data.numberFormat && data.numberFormat[r] ? data.numberFormat[r][c] : "";
                        if (looksLikeDate(textValue)) {
                            dStats.validCount++;
                            dStats.formats.add(numFormat);
                            if (numFormat !== "m/d/yyyy") {
                                dateIssues++;
                                dateCols.add(c);
                            }
                        } else {
                            if (/[$€£,]/.test(textValue)) {
                                formatIssues++;
                                formatCols.add(c);
                            }
                            numericValuesByCol[c].push({ r, num: value, text: textValue, isFormula });
                        }
                    }
                }
            }

            // -------------------------------------------------
            // RESOLVE INVALID DATE CELL ADDRESSES & HIGHLIGHT
            // (must happen before issue cards are built below)
            // -------------------------------------------------
            if (localInvalidDateCells.length > 0) {
                await Excel.run(async (context) => {
                    const sheet = context.workbook.worksheets.getActiveWorksheet();
                    const range = sheet.getRange(data.address);
                    
                    localInvalidDateCells.forEach(cell => {
                        const cellRange = range.getCell(cell.r, cell.c);
                        cellRange.load("address");
                        cell.cellRange = cellRange;
                        cellRange.format.fill.color = "#FFD9B3";
                        cellRange.format.font.color = "#7D3C00";
                    });
                    
                    await context.sync();
                    
                    localInvalidDateCells.forEach(cell => {
                        cell.address = cell.cellRange.address.split('!').pop();
                    });
                });
                currentInvalidDateCells = localInvalidDateCells;
            } else {
                currentInvalidDateCells = [];
            }

            for (const c in dateStatsByCol) {
                const stats = dateStatsByCol[c];
                const colName = headers[c] || `Col ${parseInt(c)+1}`;

                if (stats.validCount > 0) {
                    if (stats.formats.size > 1) {
                        issues.push(`
                            <strong>⚠ Date Format Issue</strong><br><br>
                            Column: ${colName}<br>
                            ${stats.validCount} date values analyzed<br>
                            ${stats.formats.size} different date formats detected
                        `);
                    } else {
                        const formatArray = Array.from(stats.formats);
                        if (formatArray[0] !== "m/d/yyyy") {
                            issues.push(`
                                <strong>⚠ Date Format Issue</strong><br><br>
                                Column: ${colName}<br>
                                ${stats.validCount} date values analyzed<br>
                                1 date format used (needs normalization)
                            `);
                        }
                    }
                }

                if (stats.invalidCount > 0) {
                    const colInvalidCells = localInvalidDateCells.filter(cell => cell.c === parseInt(c));
                    let detailRows = "";
                    colInvalidCells.forEach(cell => {
                        detailRows += `<tr><td style="padding:3px 5px">${cell.address || '(resolving...)'}</td><td style="padding:3px 5px">${cell.val}</td><td style="padding:3px 5px">${cell.reason}</td></tr>`;
                    });
                    issues.push(`
                        <strong>❌ Invalid Date</strong><br><br>
                        Column: ${colName}<br>
                        ${stats.invalidCount} invalid date(s) detected<br><br>
                        <table style="width:100%;border-collapse:collapse;font-size:12px">
                            <tr style="background:#f8d7da"><th style="padding:3px 5px;text-align:left">Cell</th><th style="padding:3px 5px;text-align:left">Value</th><th style="padding:3px 5px;text-align:left">Reason</th></tr>
                            ${detailRows}
                        </table><br>
                        <em style="color:#856404">⚠️ Manual Review Required — these values cannot be automatically corrected.</em>
                    `);
                }
            }

            if (formatIssues > 0) {
                const colNames = Array.from(formatCols).map(c => headers[c] || `Col ${c+1}`).join(", ");
                // Optional legacy format reporting if no localMixedDataCells
                if (localMixedDataCells.length === 0) {
                    issues.push(`
                        <strong>⚠ Number Format Issues</strong><br><br>
                        Column: ${colNames}<br>
                        ${formatIssues} cell(s) have number formatting issues
                    `);
                }
            }

            if (localMixedDataCells.length > 0) {
                currentMixedDataCells = localMixedDataCells; // save to global state for the Convert button
                await Excel.run(async (context) => {
                    const sheet = context.workbook.worksheets.getActiveWorksheet();
                    const range = sheet.getRange(data.address);
                    
                    localMixedDataCells.forEach(cell => {
                        const cellRange = range.getCell(cell.r, cell.c);
                        cellRange.load("address");
                        cell.cellRange = cellRange;
                    });
                    
                    await context.sync();
                    
                    localMixedDataCells.forEach(cell => {
                        cell.address = cell.cellRange.address.split('!').pop();
                    });
                });

                const mixedByCol = {};
                localMixedDataCells.forEach(cell => {
                    if (!mixedByCol[cell.c]) mixedByCol[cell.c] = [];
                    mixedByCol[cell.c].push(cell);
                });

                for (const c in mixedByCol) {
                    const colName = headers[c] || `Col ${parseInt(c)+1}`;
                    const cells = mixedByCol[c];
                    const sample = cells[0];
                    issues.push(`
                        <strong>⚠️ Mixed Data Types</strong><br><br>
                        Column: ${colName}<br>
                        Expected type: Numeric<br>
                        Text-formatted numeric values: ${cells.length}<br>
                        Affected cell: ${sample.address}<br>
                        Value: "${sample.val}"
                    `);
                }
            } else {
                currentMixedDataCells = []; // reset
            }

            for (const c in categoryVariationsByCol) {
                const colName = headers[c] || `Col ${parseInt(c)+1}`;
                let variationHtml = "";
                
                for (const lowerKey in categoryVariationsByCol[c]) {
                    const variations = categoryVariationsByCol[c][lowerKey];
                    if (variations.size > 1) {
                        variationHtml += `
                            ${Array.from(variations).join('<br>')}
                            <br><br>
                        `;
                    }
                }
                
                if (variationHtml !== "") {
                    issues.push(`
                        <strong>⚠ Inconsistent Categories</strong><br><br>
                        Column: ${colName}<br><br>
                        Possible variations:<br><br>
                        ${variationHtml}
                    `);
                }
            }


            // -------------------------------------------------
            // OUTLIER ANALYSIS
            // -------------------------------------------------

            let totalOutliers = 0;
            if (outliersContainer) {
                const outlierIssues = [];
                const allOutlierCells = [];
                const outlierDataByCol = {};

                for (const c in numericValuesByCol) {
                    const items = numericValuesByCol[c];
                    if (items.length < 4) continue;
                    
                    const numbers = items.map(i => i.num);
                    const formulaCount = items.filter(i => i.isFormula).length;
                    const isDerived = formulaCount > (items.length / 2);
                    
                    const bounds = calculateIQRBounds(numbers);
                    if (bounds) {
                        const outliers = items.filter(i => i.num < bounds.lowerBound || i.num > bounds.upperBound);
                        if (outliers.length > 0) {
                            outlierDataByCol[c] = { bounds, outliers, isDerived };
                            outliers.forEach(o => {
                                allOutlierCells.push({ r: o.r, c: parseInt(c), obj: o });
                            });
                            totalOutliers += outliers.length;
                        }
                    }
                }

                if (allOutlierCells.length > 0) {
                    await Excel.run(async (context) => {
                        const sheet = context.workbook.worksheets.getActiveWorksheet();
                        const range = sheet.getRange(data.address);
                        
                        allOutlierCells.forEach(cell => {
                            const cellRange = range.getCell(cell.r, cell.c);
                            cellRange.load("address");
                            cell.cellRange = cellRange;
                            
                            // Apply red formatting to outliers
                            cellRange.format.fill.color = "#FFC7CE";
                            cellRange.format.font.color = "#9C0006";
                        });
                        
                        await context.sync();
                        
                        allOutlierCells.forEach(cell => {
                            // Extract just the cell reference (e.g. "B5") from "Sheet1!B5"
                            cell.obj.address = cell.cellRange.address.split('!').pop();
                        });
                    });

                    for (const c in outlierDataByCol) {
                        const colData = outlierDataByCol[c];
                        const colName = headers[c] || `Col ${parseInt(c)+1}`;
                        
                        let tableRows = colData.outliers.map(o => {
                            const reason = o.num < colData.bounds.lowerBound ? 'Below lower bound' : 'Above upper bound';
                            return `
                                <tr style="border-bottom: 1px solid #eee;">
                                    <td style="padding: 5px;">${o.address}</td>
                                    <td style="padding: 5px;">${o.text}</td>
                                    <td style="padding: 5px;">${reason}</td>
                                </tr>
                            `;
                        }).join("");

                        const headerTitle = colData.isDerived ? "⚠️ Derived-column outlier" : "🔴 Outliers Detected";
                        outlierIssues.push(`
                            <div class="outlier-card">
                                <strong>${headerTitle}</strong><br><br>
                                Column: ${colName}<br>
                                Method: IQR<br><br>
                                Outliers Found: ${colData.outliers.length}<br><br>
                                
                                <table style="width: 100%; text-align: left; border-collapse: collapse; margin-bottom: 15px; font-size: 13px;">
                                    <tr style="border-bottom: 1px solid #ccc;">
                                        <th style="padding: 5px; width: 25%;">Cell</th>
                                        <th style="padding: 5px; width: 30%;">Value</th>
                                        <th style="padding: 5px; width: 45%;">Reason</th>
                                    </tr>
                                    ${tableRows}
                                </table>

                                Lower Bound: ${colData.bounds.lowerBound.toFixed(2)}<br>
                                Upper Bound: ${colData.bounds.upperBound.toFixed(2)}<br><br>
                                <span style="color: #856404; background-color: #fff3cd; padding: 5px; border-radius: 3px; display: inline-block; font-size: 12px; line-height: 1.4;">⚠️ Review these values before removing or modifying them.</span>
                            </div>
                        `);
                    }
                }
                
                if (outlierIssues.length === 0) {
                    outliersContainer.innerHTML = `
                        <div class="no-issues">
                            ✓ No outliers detected.
                        </div>
                    `;
                } else {
                    outliersContainer.innerHTML = outlierIssues.join("");
                }
            }


            // -------------------------------------------------
            // DISPLAY ISSUES
            // -------------------------------------------------

            if (issuesContainer) {
                let categoryIssuesCount = 0;
                for (const c in categoryVariationsByCol) {
                    for (const lowerKey in categoryVariationsByCol[c]) {
                        if (categoryVariationsByCol[c][lowerKey].size > 1) categoryIssuesCount += categoryVariationsByCol[c][lowerKey].size;
                    }
                }
                let invalidDatesCount = 0;
                for (const c in dateStatsByCol) {
                    invalidDatesCount += dateStatsByCol[c].invalidCount;
                }

                // Persist to globals so Quality Report can read them
                currentOutlierCount = totalOutliers;
                currentCategoryIssuesCount = categoryIssuesCount;

                // dateIssues includes BOTH valid-date-format-issues AND invalid dates.
                // Separate the "correctable date format" count from truly invalid dates.
                const dateFormatOnlyCount = Math.max(dateIssues - invalidDatesCount, 0);

                const totalRows = Math.max(values.length - data.dataStartRow, 1);

                const scoreResult = calculateDataQualityScore({
                    blankCells:       blankCount,
                    missingValues:    missingNaCount,
                    duplicateRows:    duplicateCount,
                    trimIssues:       trimIssues,
                    formatIssues:     formatIssues,
                    invalidDates:     invalidDatesCount,
                    dateFormatIssues: dateFormatOnlyCount,
                    outliers:         totalOutliers,
                    categoryIssues:   categoryIssuesCount
                }, totalRows);

                const qualityScore = scoreResult.score;

                // Persist to globals so Quality Report & before-capture can use them
                currentQualityScore = qualityScore;
                currentAnalysisSnapshot = {
                    blankCells:       blankCount,
                    missingValues:    missingNaCount,
                    duplicateRows:    duplicateCount,
                    trimIssues:       trimIssues,
                    formatIssues:     formatIssues,
                    invalidDates:     invalidDatesCount,
                    dateFormatIssues: dateFormatOnlyCount,
                    outliers:         totalOutliers,
                    categoryIssues:   categoryIssuesCount,
                    totalRows:        totalRows,
                    totalColumns:     data.columnCount,
                    totalIssues:      scoreResult.totalIssues
                };

                let scoreHtml = `
                    <div class="score-card" style="margin-bottom: 20px; text-align: center;">
                        <h2 style="margin: 0; font-size: 16px; color: #555;">DATA QUALITY SCORE</h2>
                        <div style="font-size: 32px; font-weight: bold; color: ${qualityScore > 80 ? '#28a745' : (qualityScore > 60 ? '#ffc107' : '#dc3545')}; margin: 10px 0;">
                            ${qualityScore}%
                        </div>
                        ${scoreResult.totalIssues > 0 ? `<div style="font-size: 12px; color: #888;">Issues Requiring Attention: ${scoreResult.totalIssues}</div>` : ''}
                    </div>
                    
                    <div style="margin-bottom: 20px;">
                        <h3 style="margin: 0 0 5px 0; font-size: 14px;">Issues Found</h3>
                        <hr style="margin: 0 0 10px 0; border: none; border-top: 1px solid #ccc;">
                        <table style="width: 100%; text-align: left; font-size: 13px;">
                `;
                
                if (blankCount > 0) scoreHtml += `<tr><td style="padding: 2px 0;">Blank Cells</td><td style="text-align: right; padding: 2px 0;">${blankCount}</td></tr>`;
                if (missingNaCount > 0) scoreHtml += `<tr><td style="padding: 2px 0;">Missing Values</td><td style="text-align: right; padding: 2px 0;">${missingNaCount}</td></tr>`;
                if (duplicateCount > 0) scoreHtml += `<tr><td style="padding: 2px 0;">Duplicate Rows</td><td style="text-align: right; padding: 2px 0;">${duplicateCount}</td></tr>`;
                if (trimIssues > 0) scoreHtml += `<tr><td style="padding: 2px 0;">Extra Spaces</td><td style="text-align: right; padding: 2px 0;">${trimIssues}</td></tr>`;
                if (formatIssues > 0) scoreHtml += `<tr><td style="padding: 2px 0;">Number Format Issues</td><td style="text-align: right; padding: 2px 0;">${formatIssues}</td></tr>`;
                if (dateFormatOnlyCount > 0) scoreHtml += `<tr><td style="padding: 2px 0;">Date Format Issues</td><td style="text-align: right; padding: 2px 0;">${dateFormatOnlyCount}</td></tr>`;
                if (invalidDatesCount > 0) scoreHtml += `<tr><td style="padding: 2px 0; color: #dc3545;">Invalid Dates</td><td style="text-align: right; padding: 2px 0; color: #dc3545;">${invalidDatesCount}</td></tr>`;
                if (totalOutliers > 0) scoreHtml += `<tr><td style="padding: 2px 0;">Outliers</td><td style="text-align: right; padding: 2px 0;">${totalOutliers}</td></tr>`;
                if (categoryIssuesCount > 0) scoreHtml += `<tr><td style="padding: 2px 0;">Category Issues</td><td style="text-align: right; padding: 2px 0;">${categoryIssuesCount}</td></tr>`;
                if (scoreResult.totalIssues === 0) scoreHtml += `<tr><td style="padding: 2px 0; color: #28a745;">No issues found</td><td style="text-align: right; padding: 2px 0; color: #28a745;">✓</td></tr>`;

                scoreHtml += `
                        </table>
                    </div>
                `;

                if (issues.length === 0) {
                    issuesContainer.innerHTML = scoreHtml + `
                        <div class="no-issues" style="margin-top: 20px;">
                            ✓ No obvious data quality issues detected.
                        </div>
                    `;
                } else {
                    issuesContainer.innerHTML = scoreHtml + `
                        <h3 style="margin: 0 0 10px 0; font-size: 14px;">Issue Highlighting</h3>
                    ` + issues.map(function (issue) {
                        return `
                            <div class="issue-item">
                                ${issue}
                            </div>
                        `;
                    }).join("");
                }
            }


            // -------------------------------------------------
            // COLUMN PROFILE
            // -------------------------------------------------

            generateColumnProfile(values, data);


            // -------------------------------------------------
            // RESULT VISIBILITY
            // -------------------------------------------------

            if (resultElement) {
                resultElement.style.display = "block";
            }


            setStatus(
                "Analysis completed successfully.",
                "success"
            );

        } catch (error) {

            console.error(
                "Analyze error:",
                error
            );

            setStatus(
                "Error: " + (
                    error.message ||
                    error.code ||
                    "Unable to analyze data."
                ),
                "error"
            );
        }
    }


    // ---------------------------------------------------------
    // COLUMN PROFILE
    // ---------------------------------------------------------

    function generateColumnProfile(values, data) {

        if (!columnProfileElement) {
            return;
        }

        if (!values || values.length === 0) {
            columnProfileElement.innerHTML = "";
            return;
        }

        const headers = values[0];

        let html = "";

        for (
            let c = 0;
            c < headers.length;
            c++
        ) {

            const header =
                isBlank(headers[c])
                    ? `Column ${c + 1}`
                    : headers[c];

            let nonBlank = 0;
            let blank = 0;
            let missingNa = 0;
            let numeric = 0;
            let text = 0;
            let date = 0;
            let extraSpaces = 0;
            let invalidDates = 0;
            const uniqueValues = new Set();

            for (
                let r = 1;
                r < values.length;
                r++
            ) {

                const value = values[r][c];
                const textValue = data && data.text ? data.text[r][c] : String(value);

                if (isBlank(value)) {
                    blank++;
                } else {
                    nonBlank++;
                    
                    if (isMissingValue(value)) {
                        missingNa++;
                    }

                    if (typeof value === "string") {
                        uniqueValues.add(value.trim());
                        
                        if (value !== value.replace(/\s+/g, " ").trim()) {
                            extraSpaces++;
                        }

                        if (looksLikeDate(value)) {
                            date++;
                        } else if (isInvalidDate(value)) {
                            date++;
                            invalidDates++;
                        } else if (isNumeric(value)) {
                            numeric++;
                        } else {
                            text++;
                        }
                    } else if (typeof value === "number") {
                        uniqueValues.add(value);
                        if (looksLikeDate(textValue)) {
                            date++;
                        } else if (isInvalidDate(textValue)) {
                            date++;
                            invalidDates++;
                        } else {
                            numeric++;
                        }
                    } else {
                        uniqueValues.add(String(value));
                        text++;
                    }
                }
            }


            let dataType = "Mixed";

            if (date > 0 && date === nonBlank) {
                dataType = "Date";
            } else if (numeric > 0 && text === 0 && date === 0) {
                dataType = "Numeric";
            } else if (text > 0 && numeric === 0 && date === 0) {
                dataType = "Text";
            }

            const duplicates = nonBlank > 0 ? nonBlank - uniqueValues.size : 0;

            html += `
                <div class="profile-item">
                    <strong>${escapeHtml(String(header))}</strong>
                    <div>Data type: ${dataType}</div>
                    <div>Total values: ${values.length - 1}</div>
                    <div>Blank count: ${blank}</div>
                    <div>Missing/NA count: ${missingNa}</div>
                    <div>Unique values: ${uniqueValues.size}</div>
                    <div>Duplicate values: ${duplicates}</div>
                    <div>Extra-space count: ${extraSpaces}</div>
            `;
            
            if (dataType === "Date" || date > 0) {
                html += `<div>Invalid-date count: ${invalidDates}</div>`;
            }
            
            html += `</div>`;
        }

        columnProfileElement.innerHTML = html;
    }


    function escapeHtml(value) {

        return value
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
    }


    // ---------------------------------------------------------
    // CLEAN DATA
    // ---------------------------------------------------------

    async function cleanData() {

        try {

            setStatus("Cleaning selected data...");

            const selectedOption =
                missingValueOption
                    ? missingValueOption.value
                    : "leave";


            console.log(
                "Selected cleaning option:",
                selectedOption
            );


            // -------------------------------------------------
            // IMPORTANT:
            // Get the selected range again.
            // -------------------------------------------------

            const data = await getSelectedData();

            const values =
                data.values.map(function (row) {

                    return row.slice();

                });

            const formats =
                data.numberFormat.map(function (row) {

                    return row.slice();

                });


            if (
                !values ||
                values.length <= 1
            ) {

                setStatus(
                    "Please select a dataset containing data rows.",
                    "warning"
                );

                return;
            }


            // -------------------------------------------------
            // COUNTERS
            // -------------------------------------------------

            let cellsTrimmed = 0;
            let capitalizationFixed = 0;
            let numbersConverted = 0;
            let datesNormalized = 0;
            let blankCellsReplaced = 0;
            let duplicateRowsRemoved = 0;


            // -------------------------------------------------
            // CALCULATE COLUMN INFORMATION
            // BEFORE MODIFYING DATA
            // -------------------------------------------------

            const columnTypes = [];

            const columnMeans = [];

            const columnMedians = [];


            for (
                let c = 0;
                c < data.columnCount;
                c++
            ) {

                const columnValues = [];

                for (
                    let r = 1;
                    r < values.length;
                    r++
                ) {

                    const value = values[r][c];

                    if (!isBlank(value)) {
                        columnValues.push(value);
                    }
                }


                let numericCount = 0;
                let textCount = 0;

                columnValues.forEach(function (value) {

                    if (isNumeric(value)) {
                        numericCount++;
                    } else {
                        textCount++;
                    }
                });


                let type = "mixed";

                if (
                    numericCount > 0 &&
                    textCount === 0
                ) {

                    type = "numeric";

                } else if (
                    textCount > 0 &&
                    numericCount === 0
                ) {

                    type = "text";
                }


                columnTypes[c] = type;

                columnMeans[c] =
                    calculateMean(columnValues);

                columnMedians[c] =
                    calculateMedian(columnValues);
            }


            // =================================================
            // STEP 1
            // TRIM SPACES
            // =================================================

            for (
                let r = 1;
                r < values.length;
                r++
            ) {

                for (
                    let c = 0;
                    c < values[r].length;
                    c++
                ) {

                    const value = values[r][c];

                    if (typeof value === "string") {

                        const trimmed =
                            value.replace(/\s+/g, " ").trim();

                        if (trimmed !== value) {

                            values[r][c] =
                                trimmed;

                            cellsTrimmed++;
                        }
                    }
                }
            }


            // =================================================
            // STEP 2
            // CAPITALIZATION
            // =================================================

            for (
                let r = 1;
                r < values.length;
                r++
            ) {

                for (
                    let c = 0;
                    c < values[r].length;
                    c++
                ) {

                    const value = values[r][c];

                    // Only modify text columns.
                    if (
                        typeof value === "string" &&
                        columnTypes[c] === "text" &&
                        value.trim() !== ""
                    ) {

                        const capitalized =
                            capitalizeText(value);

                        if (
                            capitalized !== value
                        ) {

                            values[r][c] =
                                capitalized;

                            capitalizationFixed++;
                        }
                    }
                }
            }


            // =================================================
            // STEP 3
            // NUMBER CONVERSION
            // =================================================

            for (
                let r = 1;
                r < values.length;
                r++
            ) {

                for (
                    let c = 0;
                    c < values[r].length;
                    c++
                ) {

                    const value = values[r][c];

                    const textValue = data.text[r][c];

                    if (typeof value === "string") {
                        if (shouldConvertToNumber(value)) {
                            const number = numericValue(value);
                            if (value !== number) {
                                values[r][c] = number;
                                formats[r][c] = "General";
                                numbersConverted++;
                            }
                        }
                    } else if (typeof value === "number") {
                        if (/[$€£,]/.test(textValue)) {
                            formats[r][c] = "General";
                            numbersConverted++;
                        }
                    }
                }
            }


            // =================================================
            // STEP 4
            // DATE NORMALIZATION
            // =================================================

            for (
                let r = 1;
                r < values.length;
                r++
            ) {

                for (
                    let c = 0;
                    c < values[r].length;
                    c++
                ) {

                    const value = values[r][c];
                    const textValue = data.text[r][c];
                    const numFormat = formats[r][c] || "";

                    if (typeof value === "string") {
                        if (looksLikeDate(value)) {
                            const normalized = normalizeDate(value);
                            if (normalized !== value) {
                                values[r][c] = normalized;
                                formats[r][c] = "m/d/yyyy";
                                datesNormalized++;
                            }
                        }
                    } else if (typeof value === "number") {
                        if (looksLikeDate(textValue) && numFormat !== "m/d/yyyy") {
                            formats[r][c] = "m/d/yyyy";
                            datesNormalized++;
                        }
                    }
                }
            }


            // =================================================
            // STEP 5
            // MISSING VALUE REPLACEMENT
            //
            // THIS IS THE IMPORTANT FIX.
            //
            // The exact dropdown value controls the action.
            //
            // "na"   -> N/A
            // "zero" -> 0
            // "mean" -> column mean
            // "median" -> column median
            // "smart" -> numeric mean / text N/A
            // "leave" -> leave blank
            // =================================================

            for (
                let r = 1;
                r < values.length;
                r++
            ) {

                for (
                    let c = 0;
                    c < values[r].length;
                    c++
                ) {

                    if (!isBlank(values[r][c])) {
                        continue;
                    }


                    // -----------------------------------------
                    // LEAVE BLANK
                    // -----------------------------------------

                    if (
                        selectedOption === "leave"
                    ) {

                        continue;
                    }


                    // -----------------------------------------
                    // REPLACE WITH N/A
                    //
                    // ALWAYS N/A.
                    // It must NEVER calculate mean.
                    // -----------------------------------------

                    if (
                        selectedOption === "na"
                    ) {

                        values[r][c] = "N/A";

                        blankCellsReplaced++;

                        continue;
                    }


                    // -----------------------------------------
                    // REPLACE WITH 0
                    //
                    // ALWAYS numeric 0.
                    // It must NEVER calculate mean.
                    // -----------------------------------------

                    if (
                        selectedOption === "zero"
                    ) {

                        values[r][c] = 0;

                        blankCellsReplaced++;

                        continue;
                    }


                    // -----------------------------------------
                    // REPLACE WITH MEAN
                    // -----------------------------------------

                    if (
                        selectedOption === "mean"
                    ) {

                        if (
                            columnMeans[c] !== null &&
                            columnTypes[c] === "numeric"
                        ) {

                            values[r][c] =
                                columnMeans[c];

                            blankCellsReplaced++;
                        }

                        continue;
                    }


                    // -----------------------------------------
                    // REPLACE WITH MEDIAN
                    // -----------------------------------------

                    if (
                        selectedOption === "median"
                    ) {

                        if (
                            columnMedians[c] !== null &&
                            columnTypes[c] === "numeric"
                        ) {

                            values[r][c] =
                                columnMedians[c];

                            blankCellsReplaced++;
                        }

                        continue;
                    }


                    // -----------------------------------------
                    // SMART REPLACE
                    //
                    // Numeric column -> Mean
                    // Text column -> N/A
                    // Mixed column -> N/A
                    // -----------------------------------------

                    if (
                        selectedOption === "smart"
                    ) {

                        if (
                            columnTypes[c] === "numeric" &&
                            columnMeans[c] !== null
                        ) {

                            values[r][c] =
                                columnMeans[c];

                        } else {

                            values[r][c] =
                                "N/A";
                        }

                        blankCellsReplaced++;

                        continue;
                    }
                }
            }


            // =================================================
            // DUPLICATE REMOVAL
            //
            // We remove duplicate rows from the array and shift
            // the remaining data up, padding the bottom with
            // blank cells to maintain the exact range dimensions.
            // =================================================

            const uniqueValues = [];
            const uniqueFormats = [];
            
            // Keep the header
            uniqueValues.push(values[0]);
            uniqueFormats.push(formats[0]);
            
            const seenRows = new Set();
            
            let currentPkIndex = -1;
            if (primaryKeySelect) {
                if (primaryKeySelect.value === "auto") currentPkIndex = detectedPrimaryKeyIndex;
                else currentPkIndex = parseInt(primaryKeySelect.value, 10);
            }

            for (
                let r = 1;
                r < values.length;
                r++
            ) {

                const key = createRowKey(values[r], currentPkIndex);

                if (seenRows.has(key)) {

                    duplicateRowsRemoved++;

                } else {

                    seenRows.add(key);
                    uniqueValues.push(values[r]);
                    uniqueFormats.push(formats[r]);
                }
            }

            // Pad the bottom with empty cells
            const emptyRow = new Array(data.columnCount).fill("");
            const emptyFormat = new Array(data.columnCount).fill("General");
            while (uniqueValues.length < values.length) {
                uniqueValues.push([...emptyRow]);
                uniqueFormats.push([...emptyFormat]);
            }

            // Replace the original values array with the unique one
            for (let r = 0; r < values.length; r++) {
                values[r] = uniqueValues[r];
                formats[r] = uniqueFormats[r];
            }


            // =================================================
            // WRITE CLEANED DATA TO EXCEL
            // =================================================

            await Excel.run(async function (context) {

                const range =
                    context.workbook.getSelectedRange();

                range.values = values;
                range.numberFormat = formats;

                await context.sync();
            });


            // =================================================
            // UPDATE SUMMARY
            // =================================================

            displayCleaningSummary({

                cellsTrimmed:
                    cellsTrimmed,

                capitalizationFixed:
                    capitalizationFixed,

                numbersConverted:
                    numbersConverted,

                datesNormalized:
                    datesNormalized,

                blankCellsReplaced:
                    blankCellsReplaced,

                duplicateRowsRemoved:
                    duplicateRowsRemoved,

                selectedOption:
                    selectedOption
            });


            // =================================================
            // REFRESH ANALYSIS
            // Capture BEFORE state first, then re-analyze for AFTER state
            // =================================================

            // Snapshot BEFORE state (currentQualityScore/currentAnalysisSnapshot
            // still hold the pre-cleaning analysis values at this point)
            beforeCleaningScore = currentQualityScore;
            beforeCleaningSnapshot = currentAnalysisSnapshot ? { ...currentAnalysisSnapshot } : null;

            // Persist the cleaning counters so generateQualityReport can show them
            lastCleaningSummary = {
                cellsTrimmed,
                capitalizationFixed,
                numbersConverted,
                datesNormalized,
                blankCellsReplaced,
                duplicateRowsRemoved,
                selectedOption
            };

            await analyzeData();


            setStatus(
                "Cleaning completed successfully.",
                "success"
            );


        } catch (error) {

            console.error(
                "Clean error:",
                error
            );

            setStatus(
                "Error: " + (
                    error.message ||
                    error.code ||
                    "Unable to clean data."
                ),
                "error"
            );
        }
    }


    // ---------------------------------------------------------
    // CLEANING SUMMARY
    // ---------------------------------------------------------

    function displayCleaningSummary(summary) {

        if (!cleaningSummaryElement) {
            return;
        }


        const optionNames = {

            leave:
                "Leave Blank",

            na:
                "Replace with N/A",

            zero:
                "Replace with 0",

            mean:
                "Replace with Mean",

            median:
                "Replace with Median",

            smart:
                "Smart Replace"
        };


        const selectedName =
            optionNames[summary.selectedOption] ||
            "Unknown";


        // -----------------------------------------------------
        // If absolutely nothing changed
        // -----------------------------------------------------

        if (
            summary.cellsTrimmed === 0 &&
            summary.capitalizationFixed === 0 &&
            summary.numbersConverted === 0 &&
            summary.datesNormalized === 0 &&
            summary.blankCellsReplaced === 0 &&
            summary.duplicateRowsRemoved === 0
        ) {

            cleaningSummaryElement.innerHTML = `
                <div class="no-issues">
                    ✓ Data is already clean.
                    No changes were required.
                    <br><br>
                    Missing Value Action:
                    <strong>${selectedName}</strong>
                </div>
            `;

            return;
        }


        // -----------------------------------------------------
        // Normal summary
        // -----------------------------------------------------

        cleaningSummaryElement.innerHTML = `

            <div class="cleaning-success">

                <strong>
                    ✓ Cleaning Completed
                </strong>

                <div>
                    Cells trimmed
                    <strong>${summary.cellsTrimmed}</strong>
                </div>

                <div>
                    Capitalization fixed
                    <strong>${summary.capitalizationFixed}</strong>
                </div>

                <div>
                    Numbers converted
                    <strong>${summary.numbersConverted}</strong>
                </div>

                <div>
                    Dates normalized
                    <strong>${summary.datesNormalized}</strong>
                </div>

                <div>
                    Blank cells replaced
                    <strong>${summary.blankCellsReplaced}</strong>
                </div>

                <div>
                    Duplicate rows removed
                    <strong>${summary.duplicateRowsRemoved}</strong>
                </div>

                <br>

                <div>
                    Missing Value Action:
                    <strong>${selectedName}</strong>
                </div>

            </div>
        `;
    }


    // ---------------------------------------------------------
    // QUALITY REPORT
    // ---------------------------------------------------------

    async function generateQualityReport() {

        try {

            setStatus("Generating quality report...");

            // -------------------------------------------------
            // Use currentAnalysisSnapshot for CURRENT REMAINING issues
            // (it was updated by the most recent analyzeData call)
            // If no analysis has run yet, run one now.
            // -------------------------------------------------
            if (!currentAnalysisSnapshot) {
                await analyzeData();
            }

            const snap = currentAnalysisSnapshot;
            const afterScore = currentQualityScore;

            // -------------------------------------------------
            // BEFORE / AFTER determination
            // If cleanData was run, beforeCleaningSnapshot holds the pre-clean state.
            // If only analyzeData was run (no cleaning yet), before = after.
            // -------------------------------------------------
            const hadCleaning = (beforeCleaningSnapshot !== null && lastCleaningSummary !== null);
            const beforeSnap  = hadCleaning ? beforeCleaningSnapshot : snap;
            const beforeScore = hadCleaning ? beforeCleaningScore    : afterScore;
            const improvement = hadCleaning ? (afterScore - beforeScore) : null;
            const cl          = lastCleaningSummary || {};

            // -------------------------------------------------
            // REMAINING ISSUES (from current snapshot = after analysis)
            // -------------------------------------------------
            const remaining = {
                blankCells:       snap.blankCells,
                missingValues:    snap.missingValues,
                duplicateRows:    snap.duplicateRows,
                trimIssues:       snap.trimIssues,
                formatIssues:     snap.formatIssues,
                dateFormatIssues: snap.dateFormatIssues,
                invalidDates:     snap.invalidDates,
                outliers:         snap.outliers,       // outliers are never auto-removed
                categoryIssues:   snap.categoryIssues
            };
            const anyRemaining = Object.values(remaining).some(v => v > 0);
            const totalRemaining = snap.totalIssues;

            // -------------------------------------------------
            // SCORE COLOR HELPER
            // -------------------------------------------------
            const scoreColor = (s) => s >= 90 ? '#28a745' : s >= 70 ? '#ffc107' : '#dc3545';

            // -------------------------------------------------
            // CSS helpers (inline, no external dependencies)
            // -------------------------------------------------
            const S = {
                section:  'margin: 0 0 16px 0; padding: 12px; background: #f8f9fa; border-radius: 6px; border-left: 3px solid #dee2e6;',
                heading:  'margin: 0 0 8px 0; font-size: 13px; font-weight: 700; letter-spacing: 0.5px; color: #333; text-transform: uppercase;',
                divider:  'border: none; border-top: 1px solid #dee2e6; margin: 6px 0 10px 0;',
                row:      'display: flex; justify-content: space-between; font-size: 12px; padding: 2px 0;',
                label:    'color: #555;',
                value:    'font-weight: 600; color: #222;',
                good:     'color: #28a745; font-weight: 600;',
                warn:     'color: #856404; font-weight: 600;',
                danger:   'color: #c0392b; font-weight: 600;',
            };

            const row = (label, value, valueStyle = '') =>
                `<div style="${S.row}"><span style="${S.label}">${label}</span><span style="${valueStyle || S.value}">${value}</span></div>`;

            const sectionHtml = (title, content, borderColor = '#6c757d') =>
                `<div style="${S.section.replace('#dee2e6', borderColor)}">
                    <h4 style="${S.heading}">${title}</h4>
                    <hr style="${S.divider}">
                    ${content}
                </div>`;

            // -------------------------------------------------
            // SECTION 1: DATASET SUMMARY
            // -------------------------------------------------
            const totalCells = snap.totalRows * snap.totalColumns;
            const datasetHtml =
                row('Rows', snap.totalRows) +
                row('Columns', snap.totalColumns) +
                row('Total Cells', totalCells);

            // -------------------------------------------------
            // SECTION 2: QUALITY SCORE
            // -------------------------------------------------
            let scoreHtml = '';
            if (hadCleaning) {
                const imp = improvement >= 0 ? `+${improvement}` : `${improvement}`;
                const impColor = improvement > 0 ? '#28a745' : improvement < 0 ? '#dc3545' : '#555';
                scoreHtml =
                    row('Before Cleaning', `${beforeScore}%`, `font-weight:700;color:${scoreColor(beforeScore)}`) +
                    row('After Cleaning',  `${afterScore}%`,  `font-weight:700;color:${scoreColor(afterScore)}`) +
                    row('Improvement', `${imp} pts`, `font-weight:700;color:${impColor}`);
            } else {
                scoreHtml =
                    `<div style="text-align:center;padding:8px 0">` +
                    `<span style="font-size:28px;font-weight:700;color:${scoreColor(afterScore)}">${afterScore}%</span>` +
                    `<div style="font-size:11px;color:#888;margin-top:2px">${anyRemaining ? `${totalRemaining} issues detected` : 'No issues detected'}</div>` +
                    `</div>`;
            }

            // -------------------------------------------------
            // SECTION 3: ISSUES DETECTED (from before-snapshot if cleaned, else current)
            // -------------------------------------------------
            const detSnap = hadCleaning ? beforeSnap : snap;
            let detectedHtml = '';
            if (detSnap.blankCells > 0)       detectedHtml += row('Blank Cells',          detSnap.blankCells);
            if (detSnap.missingValues > 0)     detectedHtml += row('Missing Values',       detSnap.missingValues);
            if (detSnap.duplicateRows > 0)     detectedHtml += row('Duplicate Rows',       detSnap.duplicateRows);
            if (detSnap.trimIssues > 0)        detectedHtml += row('Extra Spaces',         detSnap.trimIssues);
            if (detSnap.formatIssues > 0)      detectedHtml += row('Data Type Issues',     detSnap.formatIssues);
            if (detSnap.dateFormatIssues > 0)  detectedHtml += row('Date Format Issues',   detSnap.dateFormatIssues);
            if (detSnap.invalidDates > 0)      detectedHtml += row('Invalid Dates',        detSnap.invalidDates, S.danger);
            if (detSnap.outliers > 0)          detectedHtml += row('Outliers',             detSnap.outliers, S.warn);
            if (detSnap.categoryIssues > 0)    detectedHtml += row('Category Issues',      detSnap.categoryIssues);
            if (!detectedHtml)                 detectedHtml  = `<div style="font-size:12px;color:#28a745">✓ No issues detected.</div>`;

            // -------------------------------------------------
            // SECTION 4: CLEANING SUMMARY (only if cleanData was run)
            // -------------------------------------------------
            let cleaningHtml = '';
            if (hadCleaning) {
                if (cl.blankCellsReplaced > 0)   cleaningHtml += row('Missing Values Replaced', cl.blankCellsReplaced);
                if (cl.duplicateRowsRemoved > 0)  cleaningHtml += row('Duplicate Rows Removed',  cl.duplicateRowsRemoved);
                if (cl.cellsTrimmed > 0)          cleaningHtml += row('Spaces Cleaned',           cl.cellsTrimmed);
                if (cl.capitalizationFixed > 0)   cleaningHtml += row('Categories Standardized', cl.capitalizationFixed);
                if (cl.numbersConverted > 0)      cleaningHtml += row('Numbers Converted',        cl.numbersConverted);
                if (cl.datesNormalized > 0)       cleaningHtml += row('Dates Normalized',         cl.datesNormalized);
                if (!cleaningHtml)                cleaningHtml  = `<div style="font-size:12px;color:#888">No changes were made.</div>`;
            }

            // -------------------------------------------------
            // SECTION 5: REMAINING ISSUES (always from current snapshot)
            // -------------------------------------------------
            let remainingHtml = '';
            if (remaining.blankCells > 0)       remainingHtml += row('Blank Cells',        remaining.blankCells,    S.danger);
            if (remaining.missingValues > 0)     remainingHtml += row('Missing Values',     remaining.missingValues, S.danger);
            if (remaining.duplicateRows > 0)     remainingHtml += row('Duplicate Rows',     remaining.duplicateRows, S.danger);
            if (remaining.trimIssues > 0)        remainingHtml += row('Extra Spaces',       remaining.trimIssues);
            if (remaining.formatIssues > 0)      remainingHtml += row('Data Type Issues',   remaining.formatIssues,  S.warn);
            if (remaining.dateFormatIssues > 0)  remainingHtml += row('Date Format Issues', remaining.dateFormatIssues, S.warn);
            if (remaining.invalidDates > 0)      remainingHtml += row('Invalid Dates',      remaining.invalidDates,  S.danger);
            if (remaining.outliers > 0)          remainingHtml += row('Outliers (Review)', remaining.outliers,       S.warn);
            if (remaining.categoryIssues > 0)    remainingHtml += row('Category Issues',   remaining.categoryIssues, S.warn);
            if (!remainingHtml)                  remainingHtml  = `<div style="font-size:12px;color:#28a745">✓ No remaining issues.</div>`;

            // -------------------------------------------------
            // SECTION 6: MANUAL REVIEW — Invalid Dates table
            // -------------------------------------------------
            let manualHtml = '';
            if (currentInvalidDateCells && currentInvalidDateCells.length > 0) {
                const dateRows = currentInvalidDateCells.map(cell => `
                    <tr>
                        <td style="padding:4px 8px;border-bottom:1px solid #eee;font-size:11px">${cell.address || '—'}</td>
                        <td style="padding:4px 8px;border-bottom:1px solid #eee;font-size:11px">${cell.val}</td>
                        <td style="padding:4px 8px;border-bottom:1px solid #eee;font-size:11px;color:#c0392b">${cell.reason}</td>
                        <td style="padding:4px 8px;border-bottom:1px solid #eee;font-size:11px;color:#856404">Manual Review</td>
                    </tr>`).join('');
                manualHtml +=
                    `<div style="font-size:12px;color:#c0392b;font-weight:600;margin-bottom:6px">` +
                    `Invalid Dates: ${currentInvalidDateCells.length} — These were NOT modified and require manual correction.</div>` +
                    `<table style="width:100%;border-collapse:collapse">` +
                    `<tr style="background:#f8d7da;font-weight:700;font-size:11px">` +
                    `<td style="padding:4px 8px">Cell</td><td style="padding:4px 8px">Value</td>` +
                    `<td style="padding:4px 8px">Reason</td><td style="padding:4px 8px">Action</td></tr>` +
                    dateRows + `</table><br>`;
            }
            if (remaining.outliers > 0) {
                manualHtml +=
                    `<div style="font-size:12px;color:#856404;font-weight:600;margin-top:4px">` +
                    `⚠️ ${remaining.outliers} statistical outlier${remaining.outliers > 1 ? 's' : ''} detected. ` +
                    `These may be legitimate values. Review recommended before removing or modifying them.</div>`;
            }
            if (!manualHtml) {
                manualHtml = `<div style="font-size:12px;color:#28a745">✓ No items requiring manual review.</div>`;
            }

            // -------------------------------------------------
            // SECTION 7: FINAL STATUS
            // -------------------------------------------------
            let finalHtml = '';
            const onlyOutliersRemain = anyRemaining &&
                remaining.outliers > 0 &&
                (remaining.blankCells + remaining.missingValues + remaining.duplicateRows +
                 remaining.trimIssues + remaining.formatIssues + remaining.dateFormatIssues +
                 remaining.invalidDates + remaining.categoryIssues) === 0;

            if (!anyRemaining) {
                finalHtml =
                    `<div style="font-size:12px;color:#28a745;line-height:1.8">` +
                    `✓ Data is clean.<br>✓ No unresolved data-quality issues detected.</div>`;
            } else if (onlyOutliersRemain) {
                finalHtml =
                    `<div style="font-size:12px;color:#856404;line-height:1.8">` +
                    `⚠️ Cleaning completed.<br>${remaining.outliers} statistical outlier${remaining.outliers > 1 ? 's' : ''} remain for review.<br>` +
                    `<em style="color:#888">Outliers may be legitimate business values — verify before removing.</em></div>`;
            } else {
                const lines = [];
                if (remaining.invalidDates > 0)   lines.push(`Invalid Dates: ${remaining.invalidDates}`);
                if (remaining.outliers > 0)        lines.push(`Outliers: ${remaining.outliers}`);
                if (remaining.missingValues > 0)   lines.push(`Missing Values: ${remaining.missingValues}`);
                if (remaining.blankCells > 0)      lines.push(`Blank Cells: ${remaining.blankCells}`);
                if (remaining.duplicateRows > 0)   lines.push(`Duplicate Rows: ${remaining.duplicateRows}`);
                if (remaining.categoryIssues > 0)  lines.push(`Category Issues: ${remaining.categoryIssues}`);
                if (remaining.formatIssues > 0)    lines.push(`Data Type Issues: ${remaining.formatIssues}`);
                if (remaining.trimIssues > 0)      lines.push(`Extra Spaces: ${remaining.trimIssues}`);
                finalHtml =
                    `<div style="font-size:12px;color:#c0392b;line-height:1.8">` +
                    `⚠️ Data cleaning completed with issues requiring review.<br>` +
                    `<strong>Remaining issues:</strong><br>` +
                    lines.map(l => `• ${l}`).join('<br>') + `</div>`;
            }

            // -------------------------------------------------
            // ASSEMBLE FULL REPORT
            // -------------------------------------------------

            // Store full report data so the download button can use it
            const reportTimestamp = new Date().toLocaleString();
            const reportSections = {
                datasetHtml,
                scoreHtml,
                detectedHtml,
                cleaningHtml: hadCleaning ? cleaningHtml : null,
                remainingHtml,
                manualHtml,
                finalHtml,
                afterScore,
                anyRemaining,
                hadCleaning,
                timestamp: reportTimestamp
            };

            if (issuesContainer) {
                let reportHtml = `<div style="font-family:inherit">`;

                // Header with download button (flex row: spacer | title+date | download btn)
                reportHtml += `<div style="display:flex;align-items:center;margin-bottom:16px;padding:10px 12px;background:linear-gradient(135deg,#667eea22,#764ba222);border-radius:8px;gap:8px">
                    <div style="flex:1"></div>
                    <div style="text-align:center;flex:0 1 auto">
                        <div style="font-size:13px;font-weight:700;letter-spacing:1px;color:#333">DATA QUALITY REPORT</div>
                        <div style="font-size:11px;color:#888;margin-top:2px">ExcelAuto &#8212; Generated ${reportTimestamp}</div>
                    </div>
                    <div style="flex:1;display:flex;justify-content:flex-end">
                        <button
                            id="downloadReportBtn"
                            title="Download Report as Word (.doc)"
                            style="background:none;border:1px solid #667eea;cursor:pointer;padding:5px 8px;
                                   border-radius:5px;line-height:1;font-size:15px;color:#667eea;
                                   transition:background 0.2s,color 0.2s;"
                            onmouseover="this.style.background='#667eea';this.style.color='#fff'"
                            onmouseout="this.style.background='none';this.style.color='#667eea'"
                        >&#11015; .doc</button>
                    </div>
                </div>`;

                reportHtml += sectionHtml('📊 Dataset Summary', datasetHtml, '#667eea');
                reportHtml += sectionHtml('⭐ Quality Score', scoreHtml, afterScore >= 90 ? '#28a745' : afterScore >= 70 ? '#ffc107' : '#dc3545');
                reportHtml += sectionHtml('🔍 Issues Detected', detectedHtml, '#6c757d');

                if (hadCleaning) {
                    reportHtml += sectionHtml('🧹 Cleaning Summary', cleaningHtml, '#17a2b8');
                }

                reportHtml += sectionHtml('⚠️ Issues Remaining', remainingHtml, anyRemaining ? '#dc3545' : '#28a745');
                reportHtml += sectionHtml('👁️ Manual Review', manualHtml, '#856404');
                reportHtml += sectionHtml('✅ Final Status', finalHtml, anyRemaining ? '#dc3545' : '#28a745');

                reportHtml += `</div>`;
                issuesContainer.innerHTML = reportHtml;

                // Wire up the download button
                const dlBtn = document.getElementById('downloadReportBtn');
                if (dlBtn) {
                    dlBtn.addEventListener('click', () => downloadQualityReport(reportSections, sectionHtml, row));
                }
            }

            setStatus("Quality report generated successfully.", "success");

        } catch (error) {

            console.error("Report error:", error);

            setStatus(
                "Error: " + (error.message || error.code || "Unable to generate report."),
                "error"
            );
        }
    }


    // ---------------------------------------------------------
    // DOWNLOAD QUALITY REPORT
    // Builds a self-contained HTML file and triggers download.
    // Does NOT touch any cleaning or analysis logic.
    // ---------------------------------------------------------

    function downloadQualityReport(sections, sectionHtmlFn, rowFn) {

        const scoreColor = (s) => s >= 90 ? '#28a745' : s >= 70 ? '#ffc107' : '#dc3545';

        // Rebuild section HTML for the standalone file
        let body = '';

        body += sectionHtmlFn('📊 Dataset Summary', sections.datasetHtml, '#667eea');
        body += sectionHtmlFn('⭐ Quality Score',   sections.scoreHtml,
                    sections.afterScore >= 90 ? '#28a745' : sections.afterScore >= 70 ? '#ffc107' : '#dc3545');
        body += sectionHtmlFn('🔍 Issues Detected', sections.detectedHtml, '#6c757d');

        if (sections.hadCleaning && sections.cleaningHtml) {
            body += sectionHtmlFn('🧹 Cleaning Summary', sections.cleaningHtml, '#17a2b8');
        }

        body += sectionHtmlFn('⚠️ Issues Remaining', sections.remainingHtml,
                    sections.anyRemaining ? '#dc3545' : '#28a745');
        body += sectionHtmlFn('👁️ Manual Review',   sections.manualHtml, '#856404');
        body += sectionHtmlFn('✅ Final Status',     sections.finalHtml,
                    sections.anyRemaining ? '#dc3545' : '#28a745');

        // Word-compatible HTML document (opens directly in Microsoft Word)
        const wordHtml = `<html xmlns:o='urn:schemas-microsoft-com:office:office'
              xmlns:w='urn:schemas-microsoft-com:office:word'
              xmlns='http://www.w3.org/TR/REC-html40'>
<head>
<meta charset='UTF-8'>
<title>ExcelAuto &mdash; Data Quality Report</title>
<!--[if gte mso 9]>
<xml>
  <w:WordDocument>
    <w:View>Print</w:View>
    <w:Zoom>100</w:Zoom>
    <w:DoNotOptimizeForBrowser/>
  </w:WordDocument>
</xml>
<![endif]-->
<style>
  body {
    font-family: Calibri, 'Segoe UI', Arial, sans-serif;
    font-size: 11pt;
    color: #222;
    margin: 2cm;
  }
  h1 {
    font-size: 16pt;
    font-weight: bold;
    letter-spacing: 1px;
    color: #333;
    text-align: center;
    margin-bottom: 4px;
  }
  .subtitle {
    font-size: 9pt;
    color: #888;
    text-align: center;
    margin-bottom: 20px;
  }
  .section {
    margin-bottom: 16px;
    padding: 10px 12px;
    background: #f8f9fa;
    border-left: 3px solid #dee2e6;
    border-radius: 4px;
  }
  .section h4 {
    font-size: 10pt;
    font-weight: bold;
    letter-spacing: 0.5px;
    color: #333;
    text-transform: uppercase;
    margin: 0 0 6px 0;
  }
  .row {
    display: flex;
    justify-content: space-between;
    font-size: 10pt;
    padding: 2px 0;
  }
  .label { color: #555; }
  .value { font-weight: 600; color: #222; }
  hr { border: none; border-top: 1px solid #dee2e6; margin: 4px 0 8px 0; }
  .footer {
    text-align: center;
    margin-top: 24px;
    font-size: 8pt;
    color: #aaa;
  }
  table { width: 100%; border-collapse: collapse; font-size: 9pt; }
  td { padding: 4px 8px; border-bottom: 1px solid #eee; }
  @page { margin: 2cm; }
</style>
</head>
<body>
  <h1>DATA QUALITY REPORT</h1>
  <p class="subtitle">ExcelAuto &mdash; Generated ${sections.timestamp}</p>
  ${body}
  <p class="footer">ExcelAuto &bull; Excel Data Quality Assistant</p>
</body>
</html>`;

        // Trigger .doc download (Word-compatible HTML)
        const blob = new Blob([wordHtml], { type: 'application/msword;charset=utf-8' });
        const url  = URL.createObjectURL(blob);
        const a    = document.createElement('a');
        a.href     = url;
        a.download = `ExcelAuto_Quality_Report_${new Date().toISOString().slice(0,10)}.doc`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }


    // ---------------------------------------------------------
    // BUTTON EVENTS
    // ---------------------------------------------------------

    async function convertNumericText() {
        if (!currentMixedDataCells || currentMixedDataCells.length === 0) {
            setStatus("No numeric text to convert.");
            return;
        }

        const convertedCount = currentMixedDataCells.length;

        try {
            setStatus("Converting numeric text...");
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const range = sheet.getRange(currentAnalysisRangeAddress);

                currentMixedDataCells.forEach(cell => {
                    const cellRange = range.getCell(cell.r, cell.c);
                    const num = parseFloat(cell.val.replace(/[^0-9.-]/g, ""));
                    cellRange.values = [[num]];
                    cellRange.numberFormat = [["General"]];
                });

                await context.sync();
            });
            
            setStatus("Converted numeric text. Re-analyzing...");
            currentMixedDataCells = []; // reset state

            if (cleaningSummaryElement) {
                // If there are existing cleaning actions, append to them, or replace the 'No cleaning performed yet' text
                const currentHtml = cleaningSummaryElement.innerHTML;
                const newRow = `
                    <div class="result-item" style="border-bottom: 1px solid #eaeaea; padding-bottom: 8px;">
                        <span>Numbers converted:</span>
                        <strong>${convertedCount}</strong>
                    </div>
                `;
                if (currentHtml.includes("No cleaning performed yet")) {
                    cleaningSummaryElement.innerHTML = newRow;
                } else {
                    cleaningSummaryElement.innerHTML = currentHtml + newRow;
                }
            }

            await analyzeData(); // re-trigger analysis to update UI
            
        } catch (error) {
            console.error(error);
            setStatus("Error converting numeric text.");
        }
    }


    if (analyzeButton) {
        analyzeButton.addEventListener("click", async () => {
            showLoading("Analyzing data...");
            try { await analyzeData(); } finally { hideLoading(); }
        });
    }

    if (cleanButton) {
        cleanButton.addEventListener("click", async () => {
            showLoading("Cleaning data...");
            try { await cleanData(); } finally { hideLoading(); }
        });
    }

    if (reportButton) {
        reportButton.addEventListener("click", async () => {
            showLoading("Generating quality report...");
            try { await generateQualityReport(); } finally { hideLoading(); }
        });
    }

    if (convertNumericTextBtn) {
        convertNumericTextBtn.addEventListener("click", async () => {
            showLoading("Converting numeric text...");
            try { await convertNumericText(); } finally { hideLoading(); }
        });
    }


    // ---------------------------------------------------------
    // INITIAL STATUS
    // ---------------------------------------------------------

    setStatus("Ready");

});