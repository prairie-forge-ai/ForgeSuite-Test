/**
 * Gateway - Shared Office/Excel initialization and helpers for all modules
 */

// Excel runtime check
export function hasExcelRuntime() {
    return typeof Excel !== "undefined" && typeof Excel.run === "function";
}

// Initialize Office and call the provided callback when ready
export function initializeOffice(onReady) {
    try {
        Office.onReady((info) => {
            console.log("Office.onReady fired:", info);
            if (info.host === Office.HostType.Excel) {
                onReady(info);
            } else {
                console.warn("Not running in Excel, host:", info.host);
                onReady(info); // Still init for testing
            }
        });
    } catch (error) {
        console.warn("Office.onReady failed:", error);
        onReady(null);
    }
}

// Shared config table name - single source of truth for all modules
export const SHARED_CONFIG_TABLE = "SS_PF_Config";

// Find a config table by name candidates (prioritizes SS_PF_Config)
export async function getConfigTable(context, tableCandidates = [SHARED_CONFIG_TABLE]) {
    const tables = context.workbook.tables;
    tables.load("items/name");
    await context.sync();

    const match = tables.items?.find((t) => tableCandidates.includes(t.name));
    if (!match) {
        console.warn("Config table not found. Looking for:", tableCandidates);
        return null;
    }

    return context.workbook.tables.getItem(match.name);
}

// Get column indices from table headers
export function getColumnIndices(headers) {
    const normalizedHeaders = headers.map(h => String(h || "").trim().toLowerCase());
    return {
        field: normalizedHeaders.findIndex(h => h === "field" || h === "field name" || h === "setting"),
        value: normalizedHeaders.findIndex(h => h === "value" || h === "setting value"),
        type: normalizedHeaders.findIndex(h => h === "type" || h === "category"),
        title: normalizedHeaders.findIndex(h => h === "title" || h === "display name"),
        permanent: normalizedHeaders.findIndex(h => h === "permanent" || h === "persist")
    };
}

// Load configuration from a config table
export async function loadConfigFromTable(tableCandidates = [SHARED_CONFIG_TABLE]) {
    if (!hasExcelRuntime()) {
        return {};
    }

    try {
        return await Excel.run(async (context) => {
            const table = await getConfigTable(context, tableCandidates);
            if (!table) {
                return {};
            }

            const body = table.getDataBodyRange();
            const headerRange = table.getHeaderRowRange();
            body.load("values");
            headerRange.load("values");
            await context.sync();

            const headers = headerRange.values[0] || [];
            const cols = getColumnIndices(headers);

            if (cols.field === -1 || cols.value === -1) {
                console.warn("Config table missing FIELD or VALUE columns. Headers:", headers);
                return {};
            }

            const values = {};
            const rows = body.values || [];
            rows.forEach((row) => {
                const field = String(row[cols.field] || "").trim();
                if (field) {
                    values[field] = row[cols.value] ?? "";
                }
            });

            console.log("Configuration loaded:", Object.keys(values).length, "fields");
            return values;
        });
    } catch (error) {
        console.error("Failed to load configuration:", error);
        return {};
    }
}

// Save a config value to a config table
export async function saveConfigValue(fieldName, value, tableCandidates = [SHARED_CONFIG_TABLE]) {
    if (!hasExcelRuntime()) return false;

    try {
        await Excel.run(async (context) => {
            const table = await getConfigTable(context, tableCandidates);
            if (!table) {
                console.warn("Config table not found for write");
                return;
            }

            const body = table.getDataBodyRange();
            const headerRange = table.getHeaderRowRange();
            body.load("values");
            headerRange.load("values");
            await context.sync();

            const headers = headerRange.values[0] || [];
            const cols = getColumnIndices(headers);

            if (cols.field === -1 || cols.value === -1) {
                console.error("Config table missing FIELD or VALUE columns");
                return;
            }

            const rows = body.values || [];
            const targetIndex = rows.findIndex(
                (row) => String(row[cols.field] || "").trim() === fieldName
            );

            if (targetIndex >= 0) {
                body.getCell(targetIndex, cols.value).values = [[value]];
            } else {
                // Add new row with Category, Field, Value, Permanent structure
                const newRow = new Array(headers.length).fill("");
                if (cols.type >= 0) newRow[cols.type] = "Run Settings";
                newRow[cols.field] = fieldName;
                newRow[cols.value] = value;
                if (cols.permanent >= 0) newRow[cols.permanent] = "N";
                if (cols.title >= 0) newRow[cols.title] = "";
                table.rows.add(null, [newRow]);
                console.log("Added new config row:", fieldName, "=", value);
            }

            await context.sync();
            console.log("Saved config:", fieldName, "=", value);
        });
        return true;
    } catch (error) {
        console.error("Failed to save config:", fieldName, error);
        return false;
    }
}

// Activate a worksheet by name
export async function activateWorksheet(sheetName) {
    if (!hasExcelRuntime()) return;

    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
            await context.sync();
            if (!sheet.isNullObject) {
                sheet.activate();
                await context.sync();
            }
        });
    } catch (error) {
        console.error("Failed to activate worksheet:", sheetName, error);
    }
}

// Get used range values from a worksheet
export async function getSheetData(sheetName) {
    if (!hasExcelRuntime()) return [];

    try {
        return await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(sheetName);
            const range = sheet.getUsedRangeOrNullObject();
            range.load("values");
            await context.sync();

            if (range.isNullObject) {
                return [];
            }
            return range.values || [];
        });
    } catch (error) {
        console.error("Failed to get sheet data:", sheetName, error);
        return [];
    }
}

// Write data to a worksheet (clears existing data first)
export async function writeSheetData(sheetName, data) {
    if (!hasExcelRuntime() || !data.length) return false;

    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(sheetName);
            const existingRange = sheet.getUsedRangeOrNullObject();
            await context.sync();

            if (!existingRange.isNullObject) {
                existingRange.clear();
            }

            const targetRange = sheet.getRangeByIndexes(0, 0, data.length, data[0].length);
            targetRange.values = data;
            await context.sync();
        });
        return true;
    } catch (error) {
        console.error("Failed to write sheet data:", sheetName, error);
        return false;
    }
}
