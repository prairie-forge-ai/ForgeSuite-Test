/**
 * Sheet Formatting - Shared utilities for consistent Excel sheet formatting
 * 
 * Usage:
 *   import { formatSheetHeaders, formatCurrencyColumn, formatDateColumn, formatNumberColumn } from "../../Common/sheet-formatting.js";
 */

// Standard Prairie Forge header style: black background, white bold text
export const HEADER_STYLE = {
    fillColor: "#000000",
    fontColor: "#FFFFFF",
    bold: true
};

// Number format patterns
export const NUMBER_FORMATS = {
    currency: "$#,##0.00",
    currencyWithNegative: "$#,##0.00;($#,##0.00)",
    number: "#,##0.00",
    integer: "#,##0",
    percent: "0.00%",
    date: "yyyy-mm-dd",
    dateTime: "yyyy-mm-dd hh:mm"
};

/**
 * Apply standard Prairie Forge header formatting to a range
 * @param {Excel.Range} headerRange - The header row range
 */
export function formatSheetHeaders(headerRange) {
    headerRange.format.fill.color = HEADER_STYLE.fillColor;
    headerRange.format.font.color = HEADER_STYLE.fontColor;
    headerRange.format.font.bold = HEADER_STYLE.bold;
}

/**
 * Write data to a sheet with formatted headers
 * Clears existing content, writes headers + data, applies formatting
 * 
 * @param {Excel.RequestContext} context - Excel context
 * @param {Excel.Worksheet} sheet - Target worksheet
 * @param {string[]} headers - Array of header strings
 * @param {any[][]} dataRows - 2D array of data rows
 * @param {Object} [options] - Optional formatting options
 * @param {Object.<number, string>} [options.columnFormats] - Map of column index to format string
 * @returns {Promise<Excel.Range>} The data range (excluding headers)
 */
export async function writeSheetWithHeaders(context, sheet, headers, dataRows, options = {}) {
    // Clear existing content
    const used = sheet.getUsedRangeOrNullObject();
    used.load("isNullObject");
    await context.sync();
    if (!used.isNullObject) {
        used.clear();
    }
    
    const columnCount = headers.length;
    const rowCount = dataRows.length + 1; // +1 for header row
    
    // Combine headers and data
    const output = [headers, ...dataRows];
    
    // Write all data
    const fullRange = sheet.getRangeByIndexes(0, 0, rowCount, columnCount);
    fullRange.values = output;
    
    // Format header row
    const headerRange = sheet.getRangeByIndexes(0, 0, 1, columnCount);
    formatSheetHeaders(headerRange);
    
    // Apply column formats if specified
    if (options.columnFormats && dataRows.length > 0) {
        for (const [colIndex, format] of Object.entries(options.columnFormats)) {
            const col = parseInt(colIndex, 10);
            if (col >= 0 && col < columnCount) {
                const colRange = sheet.getRangeByIndexes(1, col, dataRows.length, 1);
                colRange.numberFormat = [[format]];
            }
        }
    }
    
    // Auto-fit columns
    fullRange.format.autofitColumns();
    await context.sync();
    
    // Return the data range (for further operations if needed)
    return dataRows.length > 0 
        ? sheet.getRangeByIndexes(1, 0, dataRows.length, columnCount)
        : null;
}

/**
 * Format a column as currency
 * @param {Excel.Worksheet} sheet - Target worksheet
 * @param {number} colIndex - 0-based column index
 * @param {number} dataRowCount - Number of data rows (excluding header)
 * @param {boolean} [showNegativeInParens=false] - Show negative values in parentheses
 */
export function formatCurrencyColumn(sheet, colIndex, dataRowCount, showNegativeInParens = false) {
    if (dataRowCount <= 0) return;
    const colRange = sheet.getRangeByIndexes(1, colIndex, dataRowCount, 1);
    colRange.numberFormat = [[showNegativeInParens ? NUMBER_FORMATS.currencyWithNegative : NUMBER_FORMATS.currency]];
}

/**
 * Format a column as a number with commas and 2 decimal places
 * @param {Excel.Worksheet} sheet - Target worksheet
 * @param {number} colIndex - 0-based column index
 * @param {number} dataRowCount - Number of data rows (excluding header)
 */
export function formatNumberColumn(sheet, colIndex, dataRowCount) {
    if (dataRowCount <= 0) return;
    const colRange = sheet.getRangeByIndexes(1, colIndex, dataRowCount, 1);
    colRange.numberFormat = [[NUMBER_FORMATS.number]];
}

/**
 * Format a column as a date
 * @param {Excel.Worksheet} sheet - Target worksheet
 * @param {number} colIndex - 0-based column index
 * @param {number} dataRowCount - Number of data rows (excluding header)
 * @param {string} [format] - Custom date format (defaults to yyyy-mm-dd)
 */
export function formatDateColumn(sheet, colIndex, dataRowCount, format = NUMBER_FORMATS.date) {
    if (dataRowCount <= 0) return;
    const colRange = sheet.getRangeByIndexes(1, colIndex, dataRowCount, 1);
    colRange.numberFormat = [[format]];
}

/**
 * Format a column as integer (no decimals)
 * @param {Excel.Worksheet} sheet - Target worksheet
 * @param {number} colIndex - 0-based column index
 * @param {number} dataRowCount - Number of data rows (excluding header)
 */
export function formatIntegerColumn(sheet, colIndex, dataRowCount) {
    if (dataRowCount <= 0) return;
    const colRange = sheet.getRangeByIndexes(1, colIndex, dataRowCount, 1);
    colRange.numberFormat = [[NUMBER_FORMATS.integer]];
}

/**
 * Format a column as percentage
 * @param {Excel.Worksheet} sheet - Target worksheet
 * @param {number} colIndex - 0-based column index
 * @param {number} dataRowCount - Number of data rows (excluding header)
 */
export function formatPercentColumn(sheet, colIndex, dataRowCount) {
    if (dataRowCount <= 0) return;
    const colRange = sheet.getRangeByIndexes(1, colIndex, dataRowCount, 1);
    colRange.numberFormat = [[NUMBER_FORMATS.percent]];
}

