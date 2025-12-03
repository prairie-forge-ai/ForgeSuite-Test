import { SHEET_NAMES, METRIC_ELEMENT_IDS, JE_METRIC_ELEMENT_IDS } from "./constants.js";

const currencyFormatter = new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD"
});

let reconciliation = { systemTotal: null, rawTotal: null };
let journal = { sourceTotal: null, debitTotal: null, creditTotal: null };

export function getReconciliation() {
    return { ...reconciliation };
}

export function getJournal() {
    return { ...journal };
}

export async function refreshReconciliation() {
    reconciliation = { systemTotal: null, rawTotal: null };
    try {
        reconciliation = await Excel.run(async (context) => {
            const totals = { systemTotal: null, rawTotal: null };
            const dataSheet = context.workbook.worksheets.getItem(SHEET_NAMES.DATA).getUsedRange();
            dataSheet.load("values");
            let cleanSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.DATA_CLEAN);
            cleanSheet.load("name");
            await context.sync();

            const dataValues = dataSheet.values || [];
            totals.rawTotal = sumRawData(dataValues);

            if (!cleanSheet.isNullObject) {
                const cleanRange = context.workbook.worksheets.getItem(SHEET_NAMES.DATA_CLEAN).getUsedRange();
                cleanRange.load("values");
                await context.sync();
                totals.systemTotal = sumColumn(cleanRange.values || []);
            }
            return totals;
        });
    } catch (error) {
        console.error("Failed to load reconciliation totals:", error);
        throw error;
    }
    return getReconciliation();
}

export async function refreshJournal() {
    journal = { sourceTotal: null, debitTotal: null, creditTotal: null };
    try {
        journal = await Excel.run(async (context) => {
            const totals = { sourceTotal: null, debitTotal: null, creditTotal: null };
            const cleanRange = context.workbook.worksheets.getItem(SHEET_NAMES.DATA_CLEAN).getUsedRange();
            cleanRange.load("values");
            await context.sync();
            totals.sourceTotal = sumColumn(cleanRange.values || []);

            const jeRange = context.workbook.worksheets.getItem(SHEET_NAMES.JE_DRAFT).getUsedRange();
            jeRange.load("values");
            await context.sync();
            const rows = jeRange.values || [];
            if (rows.length > 1) {
                const headers = rows[0].map((value) => normalize(value));
                const debitIndex = headers.indexOf("debit");
                const creditIndex = headers.indexOf("credit");
                for (let i = 1; i < rows.length; i++) {
                    const row = rows[i];
                    if (debitIndex >= 0) totals.debitTotal += parseNumber(row[debitIndex]);
                    if (creditIndex >= 0) totals.creditTotal += parseNumber(row[creditIndex]);
                }
            }
            return totals;
        });
    } catch (error) {
        console.error("JE Review: failed to load totals:", error);
        throw error;
    }
    return getJournal();
}

export function bindReconciliationUI({ onError }) {
    const elements = {
        system: document.getElementById(METRIC_ELEMENT_IDS.systemTotal),
        raw: document.getElementById(METRIC_ELEMENT_IDS.rawTotal),
        variance: document.getElementById("recon-variance")
    };
    updateReconciliationUI(elements, getReconciliation());
    return {
        async reload() {
            try {
                await refreshReconciliation();
                updateReconciliationUI(elements, getReconciliation());
            } catch (error) {
                onError?.(error);
            }
        },
        setBankInput(value) {
            const varianceEl = document.getElementById("recon-variance");
            const bankValue = parseNumber(value);
            const { rawTotal } = getReconciliation();
            const variance = rawTotal != null && !Number.isNaN(bankValue) ? rawTotal - bankValue : null;
            if (varianceEl) {
                varianceEl.textContent = formatCurrency(variance);
            }
        }
    };
}

export function bindJournalUI({ onError }) {
    const map = JE_METRIC_ELEMENT_IDS;
    const nodes = {
        source: document.getElementById(map.sourceTotal),
        debit: document.getElementById(map.debitTotal),
        credit: document.getElementById(map.creditTotal),
        variance: document.getElementById(map.variance)
    };
    updateJournalUI(nodes, getJournal());
    return {
        async reload() {
            try {
                await refreshJournal();
                updateJournalUI(nodes, getJournal());
            } catch (error) {
                onError?.(error);
            }
        }
    };
}

function updateReconciliationUI(elements = {}, totals = {}) {
    if (!elements) return;
    const { system, raw, variance } = elements;
    system && (system.textContent = formatCurrency(totals.systemTotal));
    raw && (raw.textContent = formatCurrency(totals.rawTotal));
    if (variance) {
        const diff =
            totals.rawTotal != null && totals.systemTotal != null
                ? totals.rawTotal - totals.systemTotal
                : null;
        variance.textContent = formatCurrency(diff);
    }
}

function updateJournalUI(nodes = {}, totals = {}) {
    if (!nodes) return;
    nodes.source && (nodes.source.textContent = formatCurrency(totals.sourceTotal));
    nodes.debit && (nodes.debit.textContent = formatCurrency(totals.debitTotal));
    nodes.credit && (nodes.credit.textContent = formatCurrency(totals.creditTotal));
    if (nodes.variance) {
        const diff =
            totals.sourceTotal != null && totals.debitTotal != null && totals.creditTotal != null
                ? totals.sourceTotal - (totals.debitTotal - totals.creditTotal)
                : null;
        nodes.variance.textContent = formatCurrency(diff);
    }
}

function formatCurrency(value) {
    return value == null || Number.isNaN(value) ? "---" : currencyFormatter.format(value);
}

function parseNumber(value) {
    if (typeof value === "number") return value;
    if (value == null) return 0;
    let cleaned = String(value).trim();
    if (!cleaned) return 0;
    const isNegative = cleaned.startsWith("(") && cleaned.endsWith(")");
    cleaned = cleaned.replace(/[$,()\s]/g, "");
    const parsed = parseFloat(cleaned);
    return Number.isNaN(parsed) ? 0 : isNegative ? -parsed : parsed;
}

function normalize(value) {
    return value ? value.toString().trim().toLowerCase() : "";
}

function sumRawData(rows) {
    try {
        const { indexMap, dataStartRow } = locateHeader(rows);
        const headerKeys = (rows[dataStartRow - 1] || []).map((value) => normalize(value));
        const totalsColumns = new Set(["grosspay", "gross pay", "gross"]);
        const employeeIndex = indexMap["employee"] ?? indexMap["employee-name"] ?? indexMap["name"];
        let total = 0;
        for (let row = dataStartRow; row < rows.length; row++) {
            const dataRow = rows[row];
            if (!dataRow || !dataRow.length) continue;
            if (employeeIndex !== undefined && !normalize(dataRow[employeeIndex])) continue;
            if (isSummaryRow(dataRow)) continue;
            for (let col = 0; col < dataRow.length; col++) {
                if (col === employeeIndex) continue;
                const header = headerKeys[col];
                if (header && totalsColumns.has(header)) continue;
                total += parseNumber(dataRow[col]);
            }
        }
        return total;
    } catch (error) {
        console.warn("Reconciliation raw total skipped:", error.message);
        return null;
    }
}

function sumColumn(rows) {
    try {
        const { indexMap, dataStartRow } = locateHeader(rows);
        const amountIndex =
            indexMap["amount"] ??
            indexMap["payroll-amount"] ??
            indexMap["amount-(usd)"] ??
            indexMap["total-amount"];
        if (amountIndex === undefined) return null;
        let total = 0;
        for (let row = dataStartRow; row < rows.length; row++) {
            const dataRow = rows[row];
            if (!dataRow || amountIndex >= dataRow.length) continue;
            total += parseNumber(dataRow[amountIndex]);
        }
        return total;
    } catch (error) {
        console.warn("Unable to resolve headers for total column:", error.message);
        return null;
    }
}

function locateHeader(rows = []) {
    const candidates = [];
    if (rows.length > 1) candidates.push(1);
    candidates.push(0);
    for (const index of candidates) {
        const row = rows[index];
        if (!row) continue;
        const map = buildIndexMap(row);
        if (Object.keys(map).length) {
            return { indexMap: map, dataStartRow: index + 1 };
        }
    }
    throw new Error("Unable to locate a header row. Ensure row 2 contains column names.");
}

function buildIndexMap(row = []) {
    const map = {};
    row.forEach((value, idx) => {
        const key = normalize(value);
        if (key && map[key] === undefined) {
            map[key] = idx;
        }
    });
    return map;
}

function isSummaryRow(row = []) {
    return row.some((cell) => {
        if (!cell) return false;
        const text = cell.toString().toLowerCase().replace(/[^a-z]/g, "");
        return text.includes("total") || text.includes("totals");
    });
}

export const __testables = {
    parseNumber,
    normalize,
    sumColumn,
    sumRawData,
    locateHeader,
    buildIndexMap
};
