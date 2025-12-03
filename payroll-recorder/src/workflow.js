import { VERSION, WORKFLOW_STEPS as STEP_DETAILS, SHEET_NAMES } from "./constants.js";
import { applyModuleTabVisibility, showAllSheets } from "../../Common/tab-visibility.js";
import { renderCopilotCard, bindCopilotCard, createExcelContextProvider } from "../../Common/copilot.js";
import { activateHomepageSheet, getHomepageConfig, renderAdaFab, removeAdaFab } from "../../Common/homepage-sheet.js";
import {
    HOME_ICON_SVG,
    MODULES_ICON_SVG,
    ARROW_LEFT_SVG,
    USERS_ICON_SVG,
    BOOK_ICON_SVG,
    ARROW_RIGHT_SVG,
    MENU_ICON_SVG,
    TABLE_ICON_SVG,
    LOCK_CLOSED_SVG,
    LOCK_OPEN_SVG,
    CHECK_ICON_SVG,
    X_CIRCLE_SVG,
    LINK_ICON_SVG,
    CALCULATOR_ICON_SVG,
    SAVE_ICON_SVG,
    DOWNLOAD_ICON_SVG,
    REFRESH_ICON_SVG,
    TRASH_ICON_SVG,
    getStepIconSvg
} from "../../Common/icons.js";
import {
    renderInlineNotes,
    renderSignoff,
    updateLockButtonVisual,
    updateSaveButtonState,
    initSaveTracking
} from "../../Common/notes-signoff.js";
import { canCompleteStep, showBlockedToast } from "../../Common/workflow-validation.js";
import { initializeOffice } from "../../Common/gateway.js";
import { formatSheetHeaders, formatCurrencyColumn, formatDateColumn, NUMBER_FORMATS } from "../../Common/sheet-formatting.js";

const MODULE_KEY = "payroll-recorder";
const MODULE_ALIAS_TOKENS = ["payroll", "payroll recorder", "payroll review", "pr"];
const MODULE_NAME = "Payroll Recorder";
const MODULE_CONFIG_SHEET = SHEET_NAMES.CONFIG || "SS_PF_Config";
const CONFIG_TABLE_CANDIDATES = ["SS_PF_Config"];
const DEFAULT_CONFIG_CATEGORY = "Run Settings";
const HERO_COPY =
    "Normalize payroll exports, enforce controls, and prep journal entries without leaving Excel. Every run follows the same guidance so you stay audit-ready.";

const WORKFLOW_STEPS = STEP_DETAILS.map(({ id, title }) => ({ id, title }));

// SS_PF_Config structure: Category (0), Field (1), Value (2), Permanent (3)
// For module-prefix rows: Field=Prefix (e.g., "PR_"), Value=Module Key (e.g., "payroll-recorder")
// For Config rows: Field=Setting Name, Value=Setting Value, Permanent=Y/N flag
const CONFIG_COLUMNS = {
    TYPE: 0,      // Category column (A)
    FIELD: 1,     // Field name column (B)
    VALUE: 2,     // Value column (C)
    PERMANENT: 3, // Permanent column (D) - Y/N flag for archive persistence
    TITLE: -1     // Not used
};
const DEFAULT_CONFIG_TYPE = "Run Settings";
const DEFAULT_CONFIG_TITLE = "";
const DEFAULT_CONFIG_PERMANENT = "N";
const NOTE_PLACEHOLDER = "Enter notes here...";
const JE_TOTAL_DEBIT_FIELD = "PR_JE_Debit_Total";
const JE_TOTAL_CREDIT_FIELD = "PR_JE_Credit_Total";
const JE_DIFFERENCE_FIELD = "PR_JE_Difference";

// Step notes/sign-off fields - Pattern: PR_{Type}_{StepName}
const STEP_NOTES_FIELDS = {
    0: { note: "PR_Notes_Config", reviewer: "PR_Reviewer_Config", signOff: "PR_SignOff_Config" },
    1: { note: "PR_Notes_Import", reviewer: "PR_Reviewer_Import", signOff: "PR_SignOff_Import" },
    2: { note: "PR_Notes_Headcount", reviewer: "PR_Reviewer_Headcount", signOff: "PR_SignOff_Headcount" },
    3: { note: "PR_Notes_Validate", reviewer: "PR_Reviewer_Validate", signOff: "PR_SignOff_Validate" },
    4: { note: "PR_Notes_Review", reviewer: "PR_Reviewer_Review", signOff: "PR_SignOff_Review" },
    5: { note: "PR_Notes_JE", reviewer: "PR_Reviewer_JE", signOff: "PR_SignOff_JE" },
    6: { note: "PR_Notes_Archive", reviewer: "PR_Reviewer_Archive", signOff: "PR_SignOff_Archive" }
};
const STEP_COMPLETE_FIELDS = {
    0: "PR_Complete_Config",
    1: "PR_Complete_Import",
    2: "PR_Complete_Headcount",
    3: "PR_Complete_Validate",
    4: "PR_Complete_Review",
    5: "PR_Complete_JE",
    6: "PR_Complete_Archive"
};
const STEP_SHEET_MAP = {
    1: SHEET_NAMES.DATA,
    2: SHEET_NAMES.DATA_CLEAN,    // Headcount Review → PR_Data_Clean
    3: SHEET_NAMES.DATA_CLEAN,    // Validate & Reconcile → PR_Data_Clean
    4: SHEET_NAMES.EXPENSE_REVIEW,
    5: SHEET_NAMES.JE_DRAFT       // Journal Entry Prep → PR_JE_Draft
};
// Config field names - Pattern: PR_{Descriptor}
const CONFIG_REVIEWER_FIELD = "PR_Reviewer";
const PAYROLL_PROVIDER_FIELD = "PR_Payroll_Provider";
const HEADCOUNT_SKIP_NOTE = "User opted to skip the headcount review this period.";

const appState = {
    statusText: "",
    focusedIndex: 0,
    activeView: "home",
    activeStepId: null,
    stepStatuses: WORKFLOW_STEPS.reduce((map, step) => ({ ...map, [step.id]: "pending" }), {})
};

const configState = {
    loaded: false,
    values: {},
    permanents: {},
    overrides: {
        accountingPeriod: false,
        jeId: false
    }
};

const pendingWrites = new Map();
let resolvedConfigTableName = null;
// Payroll date - primary name first, legacy fallbacks for migration
const PAYROLL_DATE_ALIASES = [
    "PR_Payroll_Date",
    "Payroll Date (YYYY-MM-DD)", 
    "Payroll_Date", 
    "Payroll Date",
    "Payroll_Date_(YYYY-MM-DD)"
];

const headcountState = {
    skipAnalysis: false,
    roster: {
        rosterCount: null,
        payrollCount: null,
        difference: null,
        mismatches: []
    },
    departments: {
        rosterCount: null,
        payrollCount: null,
        difference: null,
        mismatches: []
    },
    loading: false,
    hasAnalyzed: false,
    lastError: null
};

let pendingScrollIndex = null;

const validationState = {
    loading: false,
    lastError: null,
    prDataTotal: null,
    cleanTotal: null,
    reconDifference: null,
    bankAmount: "",
    bankDifference: null,
    plugEnabled: false
};

const expenseReviewState = {
    loading: false,
    lastError: null,
    periods: [],
    copilotResponse: "",
    // Data completeness check
    completenessCheck: {
        currentPeriod: null,   // { match: true/false, prDataClean: number, currentTotal: number }
        historicalPeriods: null // { match: true/false, archiveSum: number, periodsSum: number }
    }
};

const journalState = {
    debitTotal: null,
    creditTotal: null,
    difference: null,
    loading: false,
    lastError: null
};

/**
 * Run data completeness check for Payroll Expense Review
 * Validates current period total matches PR_Data_Clean and historical totals match PR_Archive_Summary
 */
async function runPayrollCompletenessCheck() {
    console.log("Completeness Check - Starting...");
    if (!hasExcelRuntime()) {
        console.log("Completeness Check - Excel runtime not available");
        return;
    }
    
    try {
        await Excel.run(async (context) => {
            const cleanSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.DATA_CLEAN);
            const archiveSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.ARCHIVE_SUMMARY);
            
            cleanSheet.load("isNullObject");
            archiveSheet.load("isNullObject");
            await context.sync();
            
            const results = {
                currentPeriod: null,
                historicalPeriods: null
            };
            
            // Check 1: Current period total from PR_Data_Clean
            if (!cleanSheet.isNullObject) {
                const cleanRange = cleanSheet.getUsedRangeOrNullObject();
                cleanRange.load("values");
                await context.sync();
                
                if (!cleanRange.isNullObject && cleanRange.values && cleanRange.values.length > 1) {
                    const headers = (cleanRange.values[0] || []).map(h => String(h || "").toLowerCase().trim());
                    
                    // Find amount column - PR_Data_Clean uses "amount" for individual line items
                    const amountIdx = headers.findIndex(h => h.includes("amount"));
                    // Fallback to total columns if no amount column
                    const totalIdx = amountIdx >= 0 ? amountIdx : headers.findIndex(h => 
                        h === "total" || h === "all-in" || h === "allin" || 
                        h === "all-in total" || h === "gross" || h === "total pay"
                    );
                    
                    console.log("Completeness Check - PR_Data_Clean headers:", headers);
                    console.log("Completeness Check - Amount column index:", amountIdx, "Total column index:", totalIdx);
                    
                    if (totalIdx >= 0) {
                        const dataRows = cleanRange.values.slice(1);
                        const prDataCleanTotal = dataRows.reduce((sum, row) => sum + (Number(row[totalIdx]) || 0), 0);
                        
                        // Get current period total from state
                        const currentPeriodTotal = expenseReviewState.periods?.[0]?.summary?.total || 0;
                        
                        console.log("Completeness Check - PR_Data_Clean sum:", prDataCleanTotal, "Current period total:", currentPeriodTotal);
                        
                        // Match with tolerance (allow $1 difference for rounding)
                        const match = Math.abs(prDataCleanTotal - currentPeriodTotal) < 1;
                        results.currentPeriod = {
                            match,
                            prDataClean: prDataCleanTotal,
                            currentTotal: currentPeriodTotal
                        };
                    } else {
                        console.warn("Completeness Check - No amount/total column found in PR_Data_Clean");
                    }
                }
            }
            
            // Check 2: Historical periods - match by date against PR_Archive_Summary
            if (!archiveSheet.isNullObject) {
                const archiveRange = archiveSheet.getUsedRangeOrNullObject();
                archiveRange.load("values");
                await context.sync();
                
                if (!archiveRange.isNullObject && archiveRange.values && archiveRange.values.length > 1) {
                    const headers = (archiveRange.values[0] || []).map(h => String(h || "").toLowerCase().trim());
                    
                    // Find date column (pay period, payroll date, date, period)
                    const dateIdx = headers.findIndex(h => 
                        h.includes("pay period") || h.includes("payroll date") || 
                        h === "date" || h === "period" || h.includes("period")
                    );
                    
                    // Find amount/total column
                    const amountIdx = headers.findIndex(h => h.includes("amount"));
                    const totalIdx = amountIdx >= 0 ? amountIdx : headers.findIndex(h => 
                        h === "total" || h === "all-in" || h === "allin" || 
                        h === "all-in total" || h === "total payroll" || h.includes("total")
                    );
                    
                    console.log("Completeness Check - PR_Archive_Summary headers:", headers);
                    console.log("Completeness Check - Date column index:", dateIdx, "Total column index:", totalIdx);
                    
                    if (totalIdx >= 0 && dateIdx >= 0) {
                        const dataRows = archiveRange.values.slice(1);
                        
                        // Get 5 periods AFTER current (indices 1-5) from expense review state
                        const historicalPeriods = (expenseReviewState.periods || []).slice(1, 6);
                        console.log("Completeness Check - Looking for periods:", historicalPeriods.map(p => p.key || p.label));
                        
                        // Build a lookup map from archive: normalize dates to YYYY-MM-DD for matching
                        // SUM all rows for each date (archive may have multiple rows per pay period)
                        const archiveLookup = new Map();
                        for (const row of dataRows) {
                            const rawDate = row[dateIdx];
                            const normalizedKey = normalizeDateForLookup(rawDate);
                            if (normalizedKey) {
                                const amount = Number(row[totalIdx]) || 0;
                                const existing = archiveLookup.get(normalizedKey) || 0;
                                archiveLookup.set(normalizedKey, existing + amount);
                            }
                        }
                        console.log("Completeness Check - Archive lookup keys:", Array.from(archiveLookup.keys()));
                        console.log("Completeness Check - Archive lookup values:", Array.from(archiveLookup.entries()));
                        
                        // Match each historical period against archive
                        let archiveSum = 0;
                        let periodsSum = 0;
                        let matchedCount = 0;
                        const periodDetails = [];
                        
                        for (const period of historicalPeriods) {
                            const periodKey = period.key || period.label || "";
                            const normalizedPeriodKey = normalizeDateForLookup(periodKey);
                            const periodTotal = period.summary?.total || 0;
                            periodsSum += periodTotal;
                            
                            // Look up in archive
                            const archiveTotal = archiveLookup.get(normalizedPeriodKey);
                            if (archiveTotal !== undefined) {
                                archiveSum += archiveTotal;
                                matchedCount++;
                                periodDetails.push({
                                    period: periodKey,
                                    calculated: periodTotal,
                                    archive: archiveTotal,
                                    match: Math.abs(periodTotal - archiveTotal) < 1
                                });
                            } else {
                                console.warn(`Completeness Check - Period ${periodKey} (normalized: ${normalizedPeriodKey}) not found in archive`);
                                periodDetails.push({
                                    period: periodKey,
                                    calculated: periodTotal,
                                    archive: null,
                                    match: false
                                });
                            }
                        }
                        
                        console.log("Completeness Check - Period details:", periodDetails);
                        console.log("Completeness Check - Matched", matchedCount, "of", historicalPeriods.length, "periods");
                        console.log("Completeness Check - Archive sum:", archiveSum, "Periods sum:", periodsSum);
                        
                        // Match if all periods matched and totals are within tolerance
                        const allMatched = matchedCount === historicalPeriods.length && historicalPeriods.length > 0;
                        const totalsMatch = Math.abs(archiveSum - periodsSum) < 1;
                        const match = allMatched && totalsMatch;
                        
                        results.historicalPeriods = {
                            match,
                            archiveSum,
                            periodsSum,
                            matchedCount,
                            totalPeriods: historicalPeriods.length,
                            details: periodDetails
                        };
                    } else {
                        console.warn("Completeness Check - Missing date or total column in PR_Archive_Summary");
                        console.warn("  Date column index:", dateIdx, "Total column index:", totalIdx);
                    }
                }
            }
            
            // Update state with results
            expenseReviewState.completenessCheck = results;
            console.log("Completeness Check - Results:", JSON.stringify(results));
        });
        console.log("Completeness Check - Complete!");
    } catch (error) {
        console.error("Payroll completeness check failed:", error);
    }
}

/**
 * Render Data Completeness Check card for Expense Review step
 * Shows detailed comparison with difference amounts
 */
function renderPayrollCompletenessCard() {
    const check = expenseReviewState.completenessCheck || {};
    const hasRun = expenseReviewState.periods?.length > 0;
    
    // Helper to format currency
    const fmt = (val) => `$${Math.round(val || 0).toLocaleString()}`;
    
    // Helper to format difference with sign
    const fmtDiff = (diff) => {
        const absDiff = Math.abs(diff);
        if (absDiff < 1) return "—";
        const sign = diff > 0 ? "+" : "-";
        return `${sign}$${Math.round(absDiff).toLocaleString()}`;
    };
    
    // Render a comparison row with source values and difference
    const renderComparisonRow = (label, sourceLabel, sourceVal, calcLabel, calcVal, isMatch, isPending) => {
        const diff = (sourceVal || 0) - (calcVal || 0);
        
        let statusIcon;
        let statusClass;
        if (isPending) {
            statusIcon = `<span class="pf-complete-status pf-complete-status--pending">⏳</span>`;
            statusClass = "pending";
        } else if (isMatch) {
            statusIcon = `<span class="pf-complete-status pf-complete-status--pass">✓</span>`;
            statusClass = "pass";
        } else {
            statusIcon = `<span class="pf-complete-status pf-complete-status--fail">✗</span>`;
            statusClass = "fail";
        }
        
        const diffDisplay = isPending ? "" : `
            <div class="pf-complete-diff ${statusClass}">
                ${fmtDiff(diff)}
            </div>
        `;
        
        return `
            <div class="pf-complete-row ${statusClass}">
                <div class="pf-complete-header">
                    ${statusIcon}
                    <span class="pf-complete-label">${escapeHtml(label)}</span>
                </div>
                ${!isPending ? `
                <div class="pf-complete-values">
                    <div class="pf-complete-value-row">
                        <span class="pf-complete-source">${escapeHtml(sourceLabel)}:</span>
                        <span class="pf-complete-amount">${fmt(sourceVal)}</span>
                    </div>
                    <div class="pf-complete-value-row">
                        <span class="pf-complete-source">${escapeHtml(calcLabel)}:</span>
                        <span class="pf-complete-amount">${fmt(calcVal)}</span>
                    </div>
                </div>
                ${diffDisplay}
                ` : `
                <div class="pf-complete-values">
                    <span class="pf-complete-pending-text">Click Run/Refresh to check</span>
                </div>
                `}
            </div>
        `;
    };
    
    // Current period check
    const currentResult = check.currentPeriod;
    const currentPending = !hasRun || currentResult === null || currentResult === undefined;
    const currentRow = renderComparisonRow(
        "Current Period",
        "PR_Data_Clean Total",
        currentResult?.prDataClean,
        "Calculated Total",
        currentResult?.currentTotal,
        currentResult?.match,
        currentPending
    );
    
    // Historical periods check - with matched count info
    const histResult = check.historicalPeriods;
    const histPending = !hasRun || histResult === null || histResult === undefined;
    const matchedCount = histResult?.matchedCount || 0;
    const totalPeriods = histResult?.totalPeriods || 0;
    const histLabel = totalPeriods > 0 
        ? `Historical Periods (${matchedCount}/${totalPeriods} matched)`
        : "Historical Periods";
    const histRow = renderComparisonRow(
        histLabel,
        "PR_Archive_Summary (matched)",
        histResult?.archiveSum,
        "Calculated Total",
        histResult?.periodsSum,
        histResult?.match,
        histPending
    );
    
    // Build period details if available (for debugging/expanded view)
    let periodDetailsHtml = "";
    if (!histPending && histResult?.details?.length > 0) {
        const detailRows = histResult.details.map(d => {
            const matchIcon = d.archive === null ? "⚠️" : (d.match ? "✓" : "✗");
            const archiveVal = d.archive !== null ? fmt(d.archive) : "Not found";
            return `
                <div class="pf-complete-detail-row">
                    <span class="pf-complete-detail-date">${escapeHtml(d.period)}</span>
                    <span class="pf-complete-detail-icon">${matchIcon}</span>
                    <span class="pf-complete-detail-vals">${fmt(d.calculated)} vs ${archiveVal}</span>
                </div>
            `;
        }).join("");
        periodDetailsHtml = `
            <div class="pf-complete-details-section">
                <div class="pf-complete-details-header">Period-by-Period Match</div>
                ${detailRows}
            </div>
        `;
    }
    
    return `
        <article class="pf-step-card pf-step-detail pf-config-card" id="data-completeness-card">
            <div class="pf-config-head">
                <h3>Data Completeness Check</h3>
                <p class="pf-config-subtext">Verify source data matches calculated totals</p>
            </div>
            <div class="pf-complete-container">
                ${currentRow}
                ${histRow}
                ${periodDetailsHtml}
            </div>
        </article>
    `;
}

/**
 * Get step-specific info panel configuration for Payroll Recorder
 */
function getStepInfoConfig(stepId) {
    switch (stepId) {
        case 0:
            return {
                title: "Configuration",
                content: `
                    <div class="pf-info-section">
                        <h4>🎯 What This Step Does</h4>
                        <p>Sets up the key parameters for your payroll review. Complete this before importing data.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>📋 Key Fields</h4>
                        <ul>
                            <li><strong>Payroll Date</strong> — The period-end date for this payroll run</li>
                            <li><strong>Accounting Period</strong> — Shows up in your JE description</li>
                            <li><strong>Journal Entry ID</strong> — Reference number for your accounting system</li>
                            <li><strong>Provider Link</strong> — Quick access to your payroll provider portal</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>💡 Tip</h4>
                        <p>The accounting period and JE ID auto-generate based on your payroll date, but you can override them if needed.</p>
                    </div>
                `
            };
        case 1:
            return {
                title: "Import Payroll Data",
                content: `
                    <div class="pf-info-section">
                        <h4>🎯 What This Step Does</h4>
                        <p>Gets your payroll data into the workbook. Pull a report from your payroll provider and paste it into PR_Data.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>📋 Required Columns</h4>
                        <p>Your payroll export should include:</p>
                        <ul>
                            <li><strong>Employee Name</strong> — Full name (used to match against roster)</li>
                            <li><strong>Department</strong> — Cost center assignment</li>
                            <li><strong>Regular Earnings</strong> — Base pay for the period</li>
                            <li><strong>Overtime</strong> — OT pay (if applicable)</li>
                            <li><strong>Bonus/Commission</strong> — Variable compensation</li>
                            <li><strong>Benefits/Deductions</strong> — Employer portions</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>💡 Tip</h4>
                        <p>Column headers don't need to match exactly—the system is flexible with naming. Just make sure each field is present.</p>
                    </div>
                `
            };
        case 2:
            return {
                title: "Headcount Review",
                content: `
                    <div class="pf-info-section">
                        <h4>🎯 What This Step Does</h4>
                        <p>Compares employee counts and department assignments between your roster and payroll data to catch discrepancies early.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>📊 Data Sources</h4>
                        <ul>
                            <li><strong>SS_Employee_Roster</strong> — Your centralized employee list (Column A: Employee names)</li>
                            <li><strong>PR_Data</strong> — The payroll data you just imported (Employee column)</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>🔍 Employee Alignment Check</h4>
                        <p>The script compares names between SS_Employee_Roster and PR_Data to find:</p>
                        <ul>
                            <li><strong>In Roster, Missing from Payroll</strong> — Employees on roster but not in payroll (possible missed payment)</li>
                            <li><strong>In Payroll, Missing from Roster</strong> — Employees paid but not on roster (possible ghost employee or new hire)</li>
                        </ul>
                        <p style="font-size: 11px; color: rgba(255,255,255,0.7); margin-top: 8px;">Names are matched using fuzzy logic to handle minor variations.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>🏢 Department Alignment Check</h4>
                        <p>For employees appearing in both sources, the script compares the "Department" column:</p>
                        <ul>
                            <li>Flags employees where roster department ≠ payroll department</li>
                            <li>Mismatches affect GL coding and cost center reporting</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>💡 Tip</h4>
                        <p>If discrepancies are expected (e.g., contractors, temp workers), you can skip this check and add a note explaining why. The note is required if you skip.</p>
                    </div>
                `
            };
        case 3:
            return {
                title: "Payroll Validation",
                content: `
                    <div class="pf-info-section">
                        <h4>🎯 What This Step Does</h4>
                        <p>Validates that your payroll totals match what was actually paid from the bank.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>📊 Reconciliation Check</h4>
                        <ul>
                            <li><strong>PR_Data Total</strong> — Sum of all payroll from your import</li>
                            <li><strong>Clean Total</strong> — Processed total after expense mapping</li>
                            <li><strong>Bank Amount</strong> — What actually left the bank account</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>⚠️ Common Differences</h4>
                        <ul>
                            <li><strong>Timing</strong> — Direct deposits vs check clearing dates</li>
                            <li><strong>Tax payments</strong> — May be separate from net pay</li>
                            <li><strong>Benefits</strong> — Some deductions paid to vendors</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>💡 Tip</h4>
                        <p>Small differences ($0.01-$1.00) are usually rounding. Use the plug feature to resolve them.</p>
                    </div>
                `
            };
        case 4:
            return {
                title: "Expense Review",
                content: `
                    <div class="pf-info-section">
                        <h4>🎯 What This Step Does</h4>
                        <p>Generates an executive-ready payroll expense summary for CFO review, with period comparisons and trend analysis.</p>
                    </div>
                    
                    <div class="pf-info-section">
                        <h4>📂 Data Sources</h4>
                        <ul>
                            <li><strong>PR_Data_Clean</strong> — Current period payroll data (cleaned and categorized)</li>
                            <li><strong>SS_Employee_Roster</strong> — Department assignments and employee details</li>
                            <li><strong>PR_Archive_Summary</strong> — Historical payroll data for trend analysis</li>
                        </ul>
                    </div>
                    
                    <div class="pf-info-section">
                        <h4>💰 How Amounts Are Calculated</h4>
                        <table style="width:100%; font-size: 11px; margin-top: 8px; border-collapse: collapse;">
                            <tr style="border-bottom: 1px solid rgba(255,255,255,0.2);">
                                <td style="padding: 6px 0;"><strong>Fixed Salary</strong></td>
                                <td style="padding: 6px 0;">Regular wages, salaries, and base pay</td>
                            </tr>
                            <tr style="border-bottom: 1px solid rgba(255,255,255,0.2);">
                                <td style="padding: 6px 0;"><strong>Variable Salary</strong></td>
                                <td style="padding: 6px 0;">Commissions, bonuses, overtime, and incentive pay</td>
                            </tr>
                            <tr style="border-bottom: 1px solid rgba(255,255,255,0.2);">
                                <td style="padding: 6px 0;"><strong>Gross Pay</strong></td>
                                <td style="padding: 6px 0;">Fixed + Variable Salary</td>
                            </tr>
                            <tr style="border-bottom: 1px solid rgba(255,255,255,0.2);">
                                <td style="padding: 6px 0;"><strong>Burden</strong></td>
                                <td style="padding: 6px 0;">Employer taxes (FICA, Medicare, FUTA, SUTA), health insurance, 401(k) match, and other employer-paid benefits</td>
                            </tr>
                            <tr style="border-bottom: 1px solid rgba(255,255,255,0.2);">
                                <td style="padding: 6px 0;"><strong>All-In Total</strong></td>
                                <td style="padding: 6px 0;">Gross Pay + Burden = Total cost to employer</td>
                            </tr>
                            <tr>
                                <td style="padding: 6px 0;"><strong>Burden Rate</strong></td>
                                <td style="padding: 6px 0;">Burden ÷ All-In Total (typically 10-18%)</td>
                            </tr>
                        </table>
                    </div>
                    
                    <div class="pf-info-section">
                        <h4>📊 Report Sections</h4>
                        <ul>
                            <li><strong>Executive Summary</strong> — Current vs prior period comparison (frozen at top)</li>
                            <li><strong>Department Breakdown</strong> — Cost allocation by cost center</li>
                            <li><strong>Historical Context</strong> — Where current metrics fall within historical ranges</li>
                            <li><strong>Period Trends</strong> — 6-period trend chart for Total, Fixed, Variable, Burden, and Headcount</li>
                        </ul>
                    </div>
                    
                    <div class="pf-info-section">
                        <h4>📈 Historical Context Visualization</h4>
                        <p>The spectrum bars show where your current period falls relative to your historical min/max:</p>
                        <p style="font-family: Consolas, monospace; color: #6366f1; margin: 8px 0;">░░░░░░░●░░░░░░░░</p>
                        <p style="font-size: 11px; color: rgba(255,255,255,0.7);">The dot (●) indicates current position. Left = Low, Right = High.</p>
                    </div>
                    
                    <div class="pf-info-section">
                        <h4>💡 Review Tips</h4>
                        <ul>
                            <li>Compare <strong>Burden Rate</strong> — Should be consistent period-to-period (10-18% typical)</li>
                            <li>Watch <strong>Variable Salary</strong> spikes — May indicate commission/bonus timing</li>
                            <li>Verify <strong>Headcount changes</strong> — Should align with HR records</li>
                            <li>Flag variances <strong>> 10%</strong> from prior period for follow-up</li>
                        </ul>
                    </div>
                `
            };
        case 5:
            return {
                title: "Journal Entry",
                content: `
                    <div class="pf-info-section">
                        <h4>🎯 What This Step Does</h4>
                        <p>Generates a balanced journal entry from your payroll data, ready for upload to your accounting system.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>📝 How the JE Works</h4>
                        <p>Maps payroll categories to GL accounts:</p>
                        <ul>
                            <li><strong>Expenses</strong> → Debits to departmental expense accounts</li>
                            <li><strong>Liabilities</strong> → Credits to payable accounts</li>
                            <li><strong>Cash</strong> → Credit to bank account</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>✅ Validation Checks</h4>
                        <ul>
                            <li><strong>Debits = Credits</strong> — Entry must balance</li>
                            <li><strong>All accounts mapped</strong> — No unassigned categories</li>
                            <li><strong>Totals match</strong> — JE ties to PR_Data</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>💡 Tip</h4>
                        <p>Review the draft in PR_JE_Draft before exporting to catch any mapping errors.</p>
                    </div>
                `
            };
        case 6:
            return {
                title: "Archive & Clear",
                content: `
                    <div class="pf-info-section">
                        <h4>🎯 What This Step Does</h4>
                        <p>Creates a backup of your completed payroll run, then resets the workbook so you're ready for the next pay period.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>📁 Step 1: Create Backup</h4>
                        <p>A new workbook opens containing all your payroll tabs. You'll choose where to save it on your computer or shared drive.</p>
                        <p style="font-size: 11px; color: rgba(255,255,255,0.6); margin-top: 6px;"><em>Tip: Use a consistent naming convention like "Payroll_Archive_2024-01-15"</em></p>
                    </div>
                    <div class="pf-info-section">
                        <h4>📊 Step 2: Update History</h4>
                        <p>The current period's totals are saved to PR_Archive_Summary. This powers the trend charts and completeness checks for future periods.</p>
                        <p style="font-size: 11px; color: rgba(255,255,255,0.6); margin-top: 6px;"><em>Keeps 5 periods of history — oldest is removed automatically</em></p>
                    </div>
                    <div class="pf-info-section">
                        <h4>🧹 Step 3: Clear Working Data</h4>
                        <p>Data is cleared from the working sheets:</p>
                        <ul>
                            <li>PR_Data (raw import)</li>
                            <li>PR_Data_Clean (processed data)</li>
                            <li>PR_Expense_Review (summary & charts)</li>
                            <li>PR_JE_Draft (journal entry lines)</li>
                        </ul>
                        <p style="font-size: 11px; color: rgba(255,255,255,0.6); margin-top: 6px;"><em>Headers are preserved — only data rows are cleared</em></p>
                    </div>
                    <div class="pf-info-section">
                        <h4>🔄 Step 4: Reset for Next Period</h4>
                        <ul>
                            <li>Payroll Date, Accounting Period, JE ID cleared</li>
                            <li>All sign-offs and completion flags reset</li>
                            <li>Notes cleared (unless you locked them with 🔒)</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>⚠️ Before You Archive</h4>
                        <ul>
                            <li>✓ JE uploaded to your accounting system</li>
                            <li>✓ All review steps signed off</li>
                            <li>✓ Lock any notes you want to keep</li>
                        </ul>
                    </div>
                `
            };
        default:
            return {
                title: "Payroll Recorder",
                content: `
                    <div class="pf-info-section">
                        <h4>👋 Welcome to Payroll Recorder</h4>
                        <p>This module helps you normalize payroll exports, enforce controls, and prep journal entries.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>📋 Workflow Overview</h4>
                        <ol style="margin: 8px 0; padding-left: 20px;">
                            <li>Configure period settings</li>
                            <li>Import payroll data</li>
                            <li>Review headcount alignment</li>
                            <li>Validate against bank</li>
                            <li>Review expense summary</li>
                            <li>Generate journal entry</li>
                            <li>Archive and reset</li>
                        </ol>
                    </div>
                    <div class="pf-info-section">
                        <p>Click a step card to get started, or tap the <strong>ⓘ</strong> button on any step for detailed guidance.</p>
                    </div>
                `
            };
    }
}

initializeOffice(() => init());

async function init() {
    try {
        await ensureTabVisibility();
        await loadConfigurationValues();
        
        // Activate module homepage on load
        const homepageConfig = getHomepageConfig(MODULE_KEY);
        await activateHomepageSheet(homepageConfig.sheetName, homepageConfig.title, homepageConfig.subtitle);
        
        renderApp();
    } catch (error) {
        console.error("[Payroll] Module initialization failed:", error);
        throw error;
    }
}

async function ensureTabVisibility() {
    // Apply prefix-based tab visibility
    // Shows PR_* tabs, hides PTO_* and SS_* tabs
    try {
        await applyModuleTabVisibility(MODULE_KEY);
        console.log(`[Payroll] Tab visibility applied for ${MODULE_KEY}`);
    } catch (error) {
        console.warn("[Payroll] Could not apply tab visibility:", error);
    }
}

function renderApp() {
    const root = document.body;
    if (!root) return;
    const prevDisabled = appState.focusedIndex <= 0 ? "disabled" : "";
    const nextDisabled = appState.focusedIndex >= WORKFLOW_STEPS.length - 1 ? "disabled" : "";
    const isConfigView = appState.activeView === "config";
    const isStepView = appState.activeView === "step";
    const isHomeView = !isConfigView && !isStepView;
    const viewMarkup = isConfigView
        ? renderConfigView()
        : isStepView
            ? renderStepView(appState.activeStepId)
            : renderHomeView();
    root.innerHTML = `
        <div class="pf-root">
            ${renderBanner(prevDisabled, nextDisabled)}
            ${viewMarkup}
            ${renderFooter()}
        </div>
    `;
    
    // Mount info FAB with step-specific content (only on step/config views, not homepage)
    const infoFabElement = document.getElementById("pf-info-fab-payroll");
    if (isHomeView) {
        // Remove info fab on homepage
        if (infoFabElement) infoFabElement.remove();
    } else if (window.PrairieForge?.mountInfoFab) {
        const infoConfig = getStepInfoConfig(appState.activeStepId);
        PrairieForge.mountInfoFab({ 
            title: infoConfig.title, 
            content: infoConfig.content, 
            buttonId: "pf-info-fab-payroll" 
        });
    }
    
    bindSharedInteractions();
    if (isConfigView) {
        bindConfigInteractions();
    } else if (isStepView) {
        try {
            bindStepInteractions(appState.activeStepId);
        } catch (error) {
            console.warn("Payroll Recorder: failed to bind step interactions", error);
        }
    } else {
        bindHomeInteractions();
    }
    scrollFocusedIntoView();
    
    // Show/hide Ada FAB based on view
    if (isHomeView) {
        renderAdaFab();
    } else {
        removeAdaFab();
    }
}

function renderBanner(prevDisabled, nextDisabled) {
    const company = getConfigValue("SS_Company_Name") || "your company";
    return `
        <div class="pf-brand-float" aria-hidden="true">
            <span class="pf-brand-wave"></span>
        </div>
        <header class="pf-banner">
            <div class="pf-nav-bar">
                <button id="nav-prev" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" aria-label="Previous step" ${prevDisabled}>
                    ${ARROW_LEFT_SVG}
                    <span class="sr-only">Previous step</span>
                </button>
                <button id="nav-home" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Module Home">
                    ${HOME_ICON_SVG}
                    <span class="sr-only">Module Home</span>
                </button>
                <button id="nav-selector" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Module Selector">
                    ${MODULES_ICON_SVG}
                    <span class="sr-only">Module Selector</span>
                </button>
                <button id="nav-next" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" aria-label="Next step" ${nextDisabled}>
                    ${ARROW_RIGHT_SVG}
                    <span class="sr-only">Next step</span>
                </button>
                <span class="pf-nav-divider"></span>
                <div class="pf-quick-access-wrapper">
                    <button id="nav-quick-toggle" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Quick Access">
                        ${MENU_ICON_SVG}
                        <span class="sr-only">Quick Access Menu</span>
                    </button>
                    <div id="quick-access-dropdown" class="pf-quick-dropdown hidden">
                        <div class="pf-quick-dropdown-header">Quick Access</div>
                        <button id="nav-roster" class="pf-quick-item pf-clickable" type="button">
                            ${USERS_ICON_SVG}
                            <span>Employee Roster</span>
                        </button>
                        <button id="nav-accounts" class="pf-quick-item pf-clickable" type="button">
                            ${BOOK_ICON_SVG}
                            <span>Chart of Accounts</span>
                        </button>
                        <button id="nav-expense-map" class="pf-quick-item pf-clickable" type="button">
                            ${TABLE_ICON_SVG}
                            <span>PR Mapping</span>
                </button>
                    </div>
                </div>
            </div>
        </header>
    `;
}

function renderHomeView() {
    return `
        <section class="pf-hero" id="pf-hero">
            <h2 class="pf-hero-title">Payroll Recorder</h2>
            <p class="pf-hero-copy">${HERO_COPY}</p>
            <p class="pf-hero-hint">${escapeHtml(appState.statusText || "")}</p>
        </section>
        <section class="pf-step-guide">
            <div class="pf-step-grid">
                ${WORKFLOW_STEPS.map((step, index) => renderStepCard(step, index)).join("")}
            </div>
        </section>
    `;
}

function renderConfigView() {
    if (!configState.loaded) {
        return `
            <section class="pf-step-guide">
                <article class="pf-step-card pf-step-detail">
                    <p class="pf-step-title">Loading configuration…</p>
                </article>
            </section>
        `;
    }
    const stepFields = STEP_NOTES_FIELDS[0];
    const payrollDate = formatDateInput(getPayrollDateValue());
    const accountingPeriod = formatDateInput(getConfigValue("PR_Accounting_Period"));
    const jeId = getConfigValue("PR_Journal_Entry_ID");
    const accountingLink = getConfigValue("SS_Accounting_Software");
    const payrollLink = getPayrollProviderLink();
    const companyName = getConfigValue("SS_Company_Name");
    const userName = getConfigValue(CONFIG_REVIEWER_FIELD) || getReviewerDefault();
    const notes = stepFields ? getConfigValue(stepFields.note) : "";
    const notesPermanent = stepFields ? isFieldPermanent(stepFields.note) : false;
    const reviewer = (stepFields ? getConfigValue(stepFields.reviewer) : "") || getReviewerDefault();
    const signOffDate = stepFields ? formatDateInput(getConfigValue(stepFields.signOff)) : "";
    const isStepComplete = Boolean(signOffDate || getConfigValue(STEP_COMPLETE_FIELDS[0]));

    return `
        <section class="pf-hero" id="pf-config-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step 0</p>
            <h2 class="pf-hero-title">Configuration Setup</h2>
            <p class="pf-hero-copy">Make quick adjustments before every payroll run.</p>
            <p class="pf-hero-hint">${escapeHtml(appState.statusText || "")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Period Data</h3>
                    <p class="pf-config-subtext">Fields in this section may change each period.</p>
                </div>
                <div class="pf-config-grid">
                    <label class="pf-config-field">
                        <span>Your Name (Used for sign-offs)</span>
                        <input type="text" id="config-user-name" value="${escapeHtml(userName)}" placeholder="Full name">
                    </label>
                    <label class="pf-config-field">
                        <span>Payroll Date</span>
                        <input type="date" id="config-payroll-date" value="${escapeHtml(payrollDate)}">
                    </label>
                    <label class="pf-config-field">
                        <span>Accounting Period</span>
                        <input type="text" id="config-accounting-period" value="${escapeHtml(accountingPeriod)}" placeholder="Nov 2025">
                    </label>
                    <label class="pf-config-field">
                        <span>Journal Entry ID</span>
                        <input type="text" id="config-je-id" value="${escapeHtml(jeId)}" placeholder="PR-AUTO-YYYY-MM-DD">
                    </label>
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Static Data</h3>
                    <p class="pf-config-subtext">Fields rarely change but should be reviewed.</p>
                </div>
                <div class="pf-config-grid">
                    <label class="pf-config-field">
                        <span>Company Name</span>
                        <input type="text" id="config-company-name" value="${escapeHtml(companyName)}" placeholder="Prairie Forge LLC">
                    </label>
                    <label class="pf-config-field">
                        <span>Payroll Provider / Report Location</span>
                        <input type="url" id="config-payroll-provider" value="${escapeHtml(payrollLink)}" placeholder="https://…">
                    </label>
                    <label class="pf-config-field">
                        <span>Accounting Software / Import Location</span>
                        <input type="url" id="config-accounting-link" value="${escapeHtml(accountingLink)}" placeholder="https://…">
                    </label>
                </div>
            </article>
            ${
                stepFields
                    ? renderInlineNotes({
                          textareaId: "config-notes",
                          value: notes,
                          permanentId: "config-notes-permanent",
                          isPermanent: notesPermanent,
                          hintId: "",
                          saveButtonId: "config-notes-save"
                      })
                    : ""
            }
            ${
                stepFields
                    ? renderSignoff({
                          reviewerInputId: "config-reviewer-name",
                          reviewerValue: reviewer,
                          signoffInputId: "config-signoff-date",
                          signoffValue: signOffDate,
                          isComplete: isStepComplete,
                          saveButtonId: "config-signoff-save",
                          completeButtonId: "config-signoff-toggle"
                      })
                    : ""
            }
        </section>
    `;
}

function renderProviderCard() {
    const providerValue = getPayrollProviderLink();
    const disabledAttr = providerValue ? "" : ' data-disabled="true" aria-disabled="true"';
    return `
        <article class="pf-step-card pf-cta-card pf-clickable" data-provider-card${disabledAttr}>
            <div>
                <p class="pf-cta-label">Payroll Provider</p>
                <p class="pf-cta-text">
                    ${
                        providerValue
                            ? escapeHtml(providerValue)
                            : "Add a Payroll Provider link in configuration to enable this shortcut."
                    }
                </p>
            </div>
            <span aria-hidden="true">↗</span>
        </article>
    `;
}

function renderQuickTipsCard() {
    return `
        <article class="pf-step-card pf-step-detail pf-quick-tips">
            <h3>Quick Tips</h3>
            <p>Helpful tips for importing payroll data will appear here soon.</p>
        </article>
    `;
}

function renderImportStep(detail) {
    const stepFields = getStepNoteFields(1);
    const notesPermanent = stepFields ? isFieldPermanent(stepFields.note) : false;
    const stepNotes = stepFields ? getConfigValue(stepFields.note) : "";
    const stepReviewer = (stepFields ? getConfigValue(stepFields.reviewer) : "") || getReviewerDefault();
    const stepSignOff = stepFields ? formatDateInput(getConfigValue(stepFields.signOff)) : "";
    const stepComplete = Boolean(stepSignOff || getConfigValue(STEP_COMPLETE_FIELDS[1]));
    const providerLink = getPayrollProviderLink();
    
    return `
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step ${detail.id}</p>
            <h2 class="pf-hero-title">${escapeHtml(detail.title)}</h2>
            <p class="pf-hero-copy">Pull your payroll export from the provider and paste it into PR_Data.</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Payroll Report</h3>
                    <p class="pf-config-subtext">Open your payroll provider, download the report, and paste into PR_Data.</p>
                </div>
                <div class="pf-signoff-action">
                    ${renderLabeledButton(
                        providerLink 
                            ? `<a href="${escapeHtml(providerLink)}" target="_blank" rel="noopener noreferrer" class="pf-action-toggle pf-clickable" title="Open payroll provider">${LINK_ICON_SVG}</a>`
                            : `<button type="button" class="pf-action-toggle pf-clickable" id="import-provider-btn" disabled title="Add provider link in Configuration">${LINK_ICON_SVG}</button>`,
                        "Provider"
                    )}
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle pf-clickable" id="import-open-data-btn" title="Open PR_Data sheet">${TABLE_ICON_SVG}</button>`,
                        "PR_Data"
                    )}
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle pf-clickable" id="import-clear-btn" title="Clear PR_Data to start over">${TRASH_ICON_SVG}</button>`,
                        "Clear"
                    )}
                </div>
            </article>
            ${stepFields ? `
                ${renderInlineNotes({
                    textareaId: "step-notes-1",
                    value: stepNotes || "",
                    permanentId: "step-notes-lock-1",
                    isPermanent: notesPermanent,
                    saveButtonId: "step-notes-save-1"
                })}
                ${renderSignoff({
                    reviewerInputId: "step-reviewer-1",
                    reviewerValue: stepReviewer,
                    signoffInputId: "step-signoff-1",
                    signoffValue: stepSignOff,
                    isComplete: stepComplete,
                    saveButtonId: "step-signoff-save-1",
                    completeButtonId: "step-signoff-toggle-1"
                })}
            ` : ""}
        </section>
    `;
}

function renderHeadcountStep(detail) {
    const stepFields = getStepNoteFields(2);
    const stepNotes = stepFields ? getConfigValue(stepFields.note) : "";
    const stepNotesPermanent = stepFields ? isFieldPermanent(stepFields.note) : false;
    const stepReviewer = (stepFields ? getConfigValue(stepFields.reviewer) : "") || getReviewerDefault();
    const stepSignOff = stepFields ? formatDateInput(getConfigValue(stepFields.signOff)) : "";
    const stepComplete = Boolean(stepSignOff || getConfigValue(STEP_COMPLETE_FIELDS[2]));
    const requiresNotes = isHeadcountNotesRequired();
    const roster = headcountState.roster || {};
    const departments = headcountState.departments || {};
    const hasRun = headcountState.hasAnalyzed;
    
    // Status banner
    let statusBanner = "";
    if (headcountState.loading) {
        statusBanner = `<p class="pf-step-note">Analyzing roster and payroll data…</p>`;
    } else if (headcountState.lastError) {
        statusBanner = `<p class="pf-step-note">${escapeHtml(headcountState.lastError)}</p>`;
    }
    
    // Build check rows (circle + pill format)
    const renderCheckRow = (label, value, isMatch) => {
        const pending = !hasRun;
        let circleHtml;
        
        if (pending) {
            circleHtml = `<span class="pf-je-check-circle pf-je-circle--pending"></span>`;
        } else if (isMatch) {
            circleHtml = `<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`;
        } else {
            circleHtml = `<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`;
        }
        
        const valueDisplay = hasRun ? ` = ${value}` : "";
        return `
            <div class="pf-je-check-row">
                ${circleHtml}
                <span class="pf-je-check-desc-pill">${escapeHtml(label)}${valueDisplay}</span>
            </div>
        `;
    };
    
    const rosterDiff = roster.difference ?? 0;
    const deptDiff = departments.difference ?? 0;
    const rosterMismatches = Array.isArray(roster.mismatches) ? roster.mismatches.filter(Boolean) : [];
    const deptMismatches = Array.isArray(departments.mismatches) ? departments.mismatches.filter(Boolean) : [];
    
    const rosterChecksHtml = `
        ${renderCheckRow("SS_Employee_Roster count", roster.rosterCount ?? "—", true)}
        ${renderCheckRow("PR_Data employee count", roster.payrollCount ?? "—", true)}
        ${renderCheckRow("Difference", rosterDiff, rosterDiff === 0)}
    `;
    
    const deptChecksHtml = `
        ${renderCheckRow("Expected departments", departments.rosterCount ?? "—", true)}
        ${renderCheckRow("PR_Data departments", departments.payrollCount ?? "—", true)}
        ${renderCheckRow("Difference", deptDiff, deptDiff === 0)}
    `;
    
    // Inline mismatch tiles for roster differences (like PTO module)
    const rosterMismatchSection =
        rosterMismatches.length && !headcountState.skipAnalysis && hasRun
            ? window.PrairieForge?.renderMismatchTiles?.({
                  mismatches: rosterMismatches,
                  label: "Employees Driving the Difference",
                  sourceLabel: "Roster",
                  targetLabel: "Payroll Data",
                  escapeHtml: escapeHtml
              }) || ""
            : "";
    
    // Inline mismatch tiles for department differences
    const deptMismatchSection =
        deptMismatches.length && !headcountState.skipAnalysis && hasRun
            ? window.PrairieForge?.renderMismatchTiles?.({
                  mismatches: deptMismatches,
                  label: "Employees with Department Differences",
                  formatter: (item) => ({
                      name: item.employee || item.name || "",
                      source: `${item.rosterDept || "—"} → ${item.payrollDept || "—"}`,
                      isMissingFromTarget: true
                  }),
                  escapeHtml: escapeHtml
              }) || ""
            : "";
    
    return `
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step ${detail.id}</p>
            <h2 class="pf-hero-title">Headcount Review</h2>
            <p class="pf-hero-copy">Quick check to make sure your roster matches your payroll data.</p>
            <div class="pf-skip-action">
                <button type="button" class="pf-skip-btn ${headcountState.skipAnalysis ? "is-active" : ""}" id="headcount-skip-btn">
                    ${X_CIRCLE_SVG}
                    <span>Skip Analysis</span>
                </button>
            </div>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Run Headcount Check</h3>
                    <p class="pf-config-subtext">Compare employee roster against payroll data.</p>
                </div>
                <div class="pf-signoff-action">
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle pf-clickable" id="roster-run-btn" title="Run headcount analysis">${CALCULATOR_ICON_SVG}</button>`,
                        "Run"
                    )}
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle pf-clickable" id="roster-refresh-btn" title="Refresh analysis">${REFRESH_ICON_SVG}</button>`,
                        "Refresh"
                    )}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Employee Alignment</h3>
                    <p class="pf-config-subtext">Verify employees match between roster and payroll.</p>
                </div>
                ${statusBanner}
                <div class="pf-je-checks-container">
                    ${rosterChecksHtml}
                </div>
                ${rosterMismatchSection}
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Department Alignment</h3>
                    <p class="pf-config-subtext">Verify department assignments are consistent.</p>
                </div>
                <div class="pf-je-checks-container">
                    ${deptChecksHtml}
                </div>
                ${deptMismatchSection}
            </article>
            ${stepFields ? `
                ${renderInlineNotes({
                    textareaId: "step-notes-input",
                    value: stepNotes,
                    permanentId: "step-notes-permanent",
                    isPermanent: stepNotesPermanent,
                    hintId: requiresNotes ? "headcount-notes-hint" : "",
                    saveButtonId: "step-notes-save-2"
                })}
                ${renderSignoff({
                    reviewerInputId: "step-reviewer-name",
                    reviewerValue: stepReviewer,
                    signoffInputId: "step-signoff-date",
                    signoffValue: stepSignOff,
                    isComplete: stepComplete,
                    saveButtonId: "headcount-signoff-save",
                    completeButtonId: "headcount-signoff-toggle"
                })}
            ` : ""}
        </section>
    `;
}

function renderValidationStep(detail) {
    const stepFields = getStepNoteFields(3);
    const stepNotes = stepFields ? getConfigValue(stepFields.note) : "";
    const stepReviewer = (stepFields ? getConfigValue(stepFields.reviewer) : "") || getReviewerDefault();
    const stepSignOff = stepFields ? formatDateInput(getConfigValue(stepFields.signOff)) : "";
    const statusBanner = validationState.loading
        ? `<p class="pf-step-note">Preparing reconciliation data…</p>`
        : validationState.lastError
            ? `<p class="pf-step-note">${escapeHtml(validationState.lastError)}</p>`
            : "";
    const stepComplete = Boolean(stepSignOff || getConfigValue(STEP_COMPLETE_FIELDS[3]));
    const hasRun = validationState.prDataTotal !== null;
    
    // Reconciliation check values
    const prDataTotal = validationState.prDataTotal;
    const cleanTotal = validationState.cleanTotal;
    const reconDiff = validationState.reconDifference ?? (prDataTotal != null && cleanTotal != null ? prDataTotal - cleanTotal : null);
    const reconMatches = reconDiff !== null && Math.abs(reconDiff) < 0.01;
    
    // Bank reconciliation values  
    const bankClean = formatCurrency(validationState.cleanTotal);
    const bankDiff = validationState.bankDifference != null ? formatCurrency(validationState.bankDifference) : "---";
    const bankHint = validationState.bankDifference == null ? ""
        : Math.abs(validationState.bankDifference) < 0.5
            ? "Difference within acceptable tolerance."
            : "Difference exceeds tolerance and should be resolved.";
    const bankValue = formatBankInput(validationState.bankAmount);
    
    // Build reconciliation check rows (circle + pill format)
    const renderCheckRow = (label, desc, passed) => {
        const pending = !hasRun;
        let circleHtml;
        
        if (pending) {
            circleHtml = `<span class="pf-je-check-circle pf-je-circle--pending"></span>`;
        } else if (passed) {
            circleHtml = `<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`;
        } else {
            circleHtml = `<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`;
        }
        
        return `
            <div class="pf-je-check-row">
                ${circleHtml}
                <span class="pf-je-check-desc-pill">${escapeHtml(desc)}</span>
            </div>
        `;
    };
    
    const prDataDisplay = hasRun ? formatCurrency(prDataTotal) : "—";
    const cleanDisplay = hasRun ? formatCurrency(cleanTotal) : "—";
    const diffDisplay = hasRun ? formatCurrency(reconDiff) : "—";
    
    const reconChecksHtml = `
        ${renderCheckRow("PR_Data Total", `PR_Data Total = ${prDataDisplay}`, true)}
        ${renderCheckRow("PR_Data_Clean Total", `PR_Data_Clean Total = ${cleanDisplay}`, true)}
        ${renderCheckRow("Difference", `Difference = ${diffDisplay} (should be $0.00)`, reconMatches)}
    `;

    return `
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step ${detail.id}</p>
            <h2 class="pf-hero-title">${escapeHtml(detail.title)}</h2>
            <p class="pf-hero-copy">Normalize your payroll data and verify totals match.</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Run Validation</h3>
                    <p class="pf-config-subtext">Build PR_Data_Clean from your imported data and verify totals.</p>
                </div>
                <div class="pf-signoff-action">
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle pf-clickable" id="validation-run-btn" title="Run reconciliation">${CALCULATOR_ICON_SVG}</button>`,
                        "Run"
                    )}
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle pf-clickable" id="validation-refresh-btn" title="Refresh reconciliation">${REFRESH_ICON_SVG}</button>`,
                        "Refresh"
                    )}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Data Reconciliation</h3>
                    <p class="pf-config-subtext">Verify PR_Data and PR_Data_Clean totals match.</p>
                </div>
                ${statusBanner}
                <div class="pf-je-checks-container">
                    ${reconChecksHtml}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Bank Reconciliation</h3>
                    <p class="pf-config-subtext">Compare payroll total to the amount pulled from the bank.</p>
                </div>
                <div class="pf-config-grid pf-metric-grid">
                    <label class="pf-config-field">
                        <span>Cost per PR_Data_Clean</span>
                        <input id="bank-clean-total-value" type="text" class="pf-readonly-input pf-metric-value" value="${bankClean}" readonly>
                    </label>
                    <label class="pf-config-field">
                        <span>Cost per Bank</span>
                        <input
                            type="text"
                            inputmode="decimal"
                            id="bank-amount-input"
                            class="pf-metric-input"
                            value="${escapeHtml(bankValue)}"
                            placeholder="0.00"
                            aria-label="Enter bank amount"
                        >
                    </label>
                    <label class="pf-config-field">
                        <span>Difference</span>
                        <input id="bank-diff-value" type="text" class="pf-readonly-input pf-metric-value" value="${bankDiff}" readonly>
                    </label>
                </div>
                <p class="pf-metric-hint" id="bank-diff-hint">${escapeHtml(bankHint)}</p>
            </article>
            ${stepFields ? `
                ${renderInlineNotes({
                    textareaId: "step-notes-input",
                    value: stepNotes,
                    permanentId: "step-notes-permanent",
                    isPermanent: isFieldPermanent(stepFields.note),
                    saveButtonId: "step-notes-save-3"
                })}
            ` : ""}
            ${renderSignoff({
                reviewerInputId: "step-reviewer-name",
                reviewerValue: stepReviewer,
                signoffInputId: "step-signoff-3",
                signoffValue: stepSignOff,
                isComplete: stepComplete,
                saveButtonId: "step-signoff-save-3",
                completeButtonId: "validation-signoff-toggle"
            })}
        </section>
    `;
}

function renderExpenseReviewStep(detail) {
    const stepFields = getStepNoteFields(4);
    const stepNotes = stepFields ? getConfigValue(stepFields.note) : "";
    const stepReviewer = (stepFields ? getConfigValue(stepFields.reviewer) : "") || getReviewerDefault();
    const stepSignOff = stepFields ? formatDateInput(getConfigValue(stepFields.signOff)) : "";
    const stepComplete = Boolean(stepSignOff || getConfigValue(STEP_COMPLETE_FIELDS[4]));
    const statusBanner = expenseReviewState.loading
        ? `<p class="pf-step-note">Preparing executive summary…</p>`
        : expenseReviewState.lastError
            ? `<p class="pf-step-note">${escapeHtml(expenseReviewState.lastError)}</p>`
            : "";
    const copilotMarkup = renderCopilotCard({
        id: "expense-review-copilot",
        body: "Want help analyzing your data? Just ask!",
        placeholder: "Where should I focus this pay period?",
        buttonLabel: "Ask CoPilot"
    });

    return `
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step ${detail.id}</p>
            <h2 class="pf-hero-title">${escapeHtml(detail.title)}</h2>
            <p class="pf-hero-copy">${escapeHtml(detail.summary || "")}</p>
            <p class="pf-hero-hint"></p>
        </section>
        <section class="pf-step-guide">
            ${statusBanner}
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Perform Analysis</h3>
                    <p class="pf-config-subtext">Populate Expense Review and perform review.</p>
                </div>
                <div class="pf-signoff-action">
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle" id="expense-run-btn" title="Run expense review analysis">${CALCULATOR_ICON_SVG}</button>`,
                        "Run"
                    )}
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle" id="expense-refresh-btn" title="Refresh expense data">${REFRESH_ICON_SVG}</button>`,
                        "Refresh"
                    )}
                </div>
            </article>
            ${renderPayrollCompletenessCard()}
                ${copilotMarkup}
            ${
                stepFields
                    ? `
            ${renderInlineNotes({
                textareaId: "step-notes-input",
                value: stepNotes,
                permanentId: "step-notes-permanent",
                isPermanent: isFieldPermanent(stepFields.note),
                saveButtonId: "step-notes-save-4"
            })}
            ${renderSignoff({
                reviewerInputId: "step-reviewer-name",
                reviewerValue: stepReviewer,
                signoffInputId: "step-signoff-4",
                signoffValue: stepSignOff,
                isComplete: stepComplete,
                saveButtonId: "step-signoff-save-4",
                completeButtonId: "expense-signoff-toggle"
            })}
            `
                    : ""
            }
        </section>
    `;
}

function renderJournalStep(detail) {
    const stepFields = getStepNoteFields(5);
    const stepNotes = stepFields ? getConfigValue(stepFields.note) : "";
    const notesPermanent = stepFields ? isFieldPermanent(stepFields.note) : false;
    const stepReviewer = (stepFields ? getConfigValue(stepFields.reviewer) : "") || getReviewerDefault();
    const stepSignOff = stepFields ? formatDateInput(getConfigValue(stepFields.signOff)) : "";
    const stepComplete = Boolean(stepSignOff || getConfigValue(STEP_COMPLETE_FIELDS[5]));
    const statusNote = journalState.lastError
        ? `<p class="pf-step-note">${escapeHtml(journalState.lastError)}</p>`
        : "";
    
    // This is a bank feed breakdown, NOT a balanced JE
    // Check 1: Line Amount should equal Debit - Credit
    // Check 2: JE total should match PR_Data_Clean total
    const hasRun = journalState.debitTotal !== null;
    const debitTotal = journalState.debitTotal ?? 0;
    const creditTotal = journalState.creditTotal ?? 0;
    const lineAmount = debitTotal - creditTotal;
    const cleanTotal = validationState.cleanTotal ?? 0;
    
    // Line Amount check: Debit - Credit should make sense
    const lineAmountValid = hasRun; // Always valid if we have data
    
    // Total matches PR_Data_Clean
    const totalMatchesClean = hasRun && Math.abs(lineAmount - cleanTotal) < 0.01;
    
    // Build validation check rows
    const renderCheckRow = (desc, passed) => {
        const pending = !hasRun;
        let circleHtml;
        
        if (pending) {
            circleHtml = `<span class="pf-je-check-circle pf-je-circle--pending"></span>`;
        } else if (passed) {
            circleHtml = `<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`;
        } else {
            circleHtml = `<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`;
        }
        
        return `
            <div class="pf-je-check-row">
                ${circleHtml}
                <span class="pf-je-check-desc-pill">${escapeHtml(desc)}</span>
            </div>
        `;
    };
    
    const debitDisplay = hasRun ? formatCurrency(debitTotal) : "—";
    const creditDisplay = hasRun ? formatCurrency(creditTotal) : "—";
    const lineAmountDisplay = hasRun ? formatCurrency(lineAmount) : "—";
    const cleanDisplay = hasRun ? formatCurrency(cleanTotal) : "—";
    
    const checksHtml = `
        ${renderCheckRow(`Total Debits = ${debitDisplay}`, lineAmountValid)}
        ${renderCheckRow(`Total Credits = ${creditDisplay}`, lineAmountValid)}
        ${renderCheckRow(`Line Amount (Debit - Credit) = ${lineAmountDisplay}`, lineAmountValid)}
        ${renderCheckRow(`JE Total matches PR_Data_Clean (${cleanDisplay})`, totalMatchesClean)}
    `;

    return `
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step ${detail.id}</p>
            <h2 class="pf-hero-title">${escapeHtml(detail.title)}</h2>
            <p class="pf-hero-copy">Generate the upload file to break down the bank feed line item.</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Generate Upload File</h3>
                    <p class="pf-config-subtext">Build the breakdown from PR_Data_Clean for your accounting system.</p>
                </div>
                <div class="pf-signoff-action">
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle pf-clickable" id="je-create-btn" title="Generate from PR_Data_Clean">${TABLE_ICON_SVG}</button>`,
                        "Generate"
                    )}
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle pf-clickable" id="je-run-btn" title="Refresh validation">${REFRESH_ICON_SVG}</button>`,
                        "Refresh"
                    )}
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle pf-clickable" id="je-export-btn" title="Export as CSV">${DOWNLOAD_ICON_SVG}</button>`,
                        "Export"
                    )}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Validation Checks</h3>
                    <p class="pf-config-subtext">Verify totals before uploading to your accounting system.</p>
                </div>
                ${statusNote}
                <div class="pf-je-checks-container">
                    ${checksHtml}
                </div>
            </article>
            ${stepFields ? `
                ${renderInlineNotes({
                    textareaId: "step-notes-input",
                    value: stepNotes || "",
                    permanentId: "step-notes-permanent",
                    isPermanent: notesPermanent,
                    saveButtonId: "step-notes-save-5"
                })}
                ${renderSignoff({
                    reviewerInputId: "step-reviewer-name",
                    reviewerValue: stepReviewer,
                    signoffInputId: "step-signoff-5",
                    signoffValue: stepSignOff,
                    isComplete: stepComplete,
                    saveButtonId: "step-signoff-save-5",
                    completeButtonId: "step-signoff-toggle-5"
                })}
            `
                    : ""
            }
        </section>
    `;
}

function renderArchiveStep(detail) {
    const completionItems = WORKFLOW_STEPS.filter((step) => step.id !== 6).map((step) => ({
        id: step.id,
        title: step.title,
        complete: isStepCompleteFromConfig(step.id)
    }));
    const allComplete = completionItems.every((item) => item.complete);
    const statusList = completionItems
        .map(
            (item) => `
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head pf-notes-header">
                    <span class="pf-action-toggle ${item.complete ? "is-active" : ""}" aria-pressed="${item.complete}">
                        ${CHECK_ICON_SVG}
                    </span>
                    <div>
                        <h3>${escapeHtml(item.title)}</h3>
                        <p class="pf-config-subtext">${item.complete ? "Complete" : "Not complete"}</p>
                    </div>
                </div>
            </article>
        `
        )
        .join("");
    return `
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step ${detail.id}</p>
            <h2 class="pf-hero-title">${escapeHtml(detail.title)}</h2>
            <p class="pf-hero-copy">${escapeHtml(detail.summary || "")}</p>
            <p class="pf-hero-hint"></p>
        </section>
        <section class="pf-step-guide">
            ${statusList}
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Archive & Reset</h3>
                    <p class="pf-config-subtext">Create an archive of this module’s sheets and clear work tabs.</p>
                </div>
                <div class="pf-pill-row pf-config-actions">
                    <button type="button" class="pf-pill-btn" id="archive-run-btn" ${allComplete ? "" : "disabled"}>Archive</button>
                </div>
            </article>
        </section>
    `;
}
function renderStepView(stepId) {
    const detail =
        STEP_DETAILS.find((step) => step.id === stepId) || {
            id: stepId ?? "-",
            title: "Workflow Step",
            summary: "",
            description: "",
            checklist: []
        };
    if (stepId === 1) return renderImportStep(detail);
    if (stepId === 2) return renderHeadcountStep(detail);
    if (stepId === 3) return renderValidationStep(detail);
    if (stepId === 4) return renderExpenseReviewStep(detail);
    if (stepId === 5) return renderJournalStep(detail);
    if (stepId === 6) return renderArchiveStep(detail);
    const isStepOne = false; // Step 1 now has dedicated render
    const stepFields = getStepNoteFields(stepId);
    const stepNotes = stepFields ? getConfigValue(stepFields.note) : "";
    const stepNotesPermanent = stepFields ? isFieldPermanent(stepFields.note) : false;
    const stepReviewer = (stepFields ? getConfigValue(stepFields.reviewer) : "") || getReviewerDefault();
    const stepSignOff = stepFields ? formatDateInput(getConfigValue(stepFields.signOff)) : "";
    const stepComplete =
        stepFields && STEP_COMPLETE_FIELDS[stepId]
            ? Boolean(stepSignOff || getConfigValue(STEP_COMPLETE_FIELDS[stepId]))
            : Boolean(stepSignOff);
    const highlights = (detail.highlights || [])
        .map(
            (item) => `
            <div class="pf-step-highlight">
                <span class="pf-step-highlight-label">${escapeHtml(item.label)}</span>
                <span class="pf-step-highlight-detail">${escapeHtml(item.detail)}</span>
            </div>
        `
        )
        .join("");
    const checklist =
        (detail.checklist || [])
            .map((item) => `<li>${escapeHtml(item)}</li>`)
            .join("") || "";
    const descriptionText = isStepOne
        ? ""
        : detail.description || "Detailed guidance will appear here.";
    const actionButtons = [];
    if (!isStepOne && detail.ctaLabel) {
        actionButtons.push(
            `<button type="button" class="pf-pill-btn" id="step-action-btn">${escapeHtml(detail.ctaLabel)}</button>`
        );
    }
    if (!isStepOne) {
        actionButtons.push(
            `<button type="button" class="pf-pill-btn pf-pill-btn--ghost" id="step-back-btn">Back to Step List</button>`
        );
    }
    const actionSection = actionButtons.length
        ? `<div class="pf-pill-row pf-config-actions">${actionButtons.join("")}</div>`
        : "";
    const providerLink = getPayrollProviderLink();
    const providerSection = isStepOne
        ? `
            <div class="pf-link-card">
                <h3 class="pf-link-card__title">Payroll Reports</h3>
                <p class="pf-link-card__subtitle">Open your latest payroll export.</p>
                <div class="pf-link-list">
                    <a
                        class="pf-link-item"
                        id="pr-provider-link"
                        ${providerLink ? `href="${escapeHtml(providerLink)}" target="_blank" rel="noopener noreferrer"` : `aria-disabled="true"`}
                    >
                        <span class="pf-link-item__icon">${LINK_ICON_SVG}</span>
                        <span class="pf-link-item__body">
                            <span class="pf-link-item__title">Open Payroll Export</span>
                            <span class="pf-link-item__meta">${escapeHtml(
                                providerLink || "Add a provider link in Configuration"
                            )}</span>
                        </span>
                    </a>
                </div>
            </div>
        `
        : "";
    const quickTipsSection = "";
    const highlightSection =
        !isStepOne && highlights ? `<article class="pf-step-card pf-step-detail">${highlights}</article>` : "";
    const checklistSection =
        !isStepOne && checklist
            ? `<article class="pf-step-card pf-step-detail">
                            <h3 class="pf-step-subtitle">Checklist</h3>
                            <ul class="pf-step-checklist">${checklist}</ul>
                        </article>`
            : "";
    const descriptionSection =
        !isStepOne || descriptionText || actionSection
            ? `
            <article class="pf-step-card pf-step-detail">
                <p class="pf-step-title">${escapeHtml(descriptionText)}</p>
                ${!isStepOne && detail.statusHint ? `<p class="pf-step-note">${escapeHtml(detail.statusHint)}</p>` : ""}
                ${actionSection}
            </article>
        `
            : "";
    return `
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step ${detail.id}</p>
            <h2 class="pf-hero-title">${escapeHtml(detail.title)}</h2>
            <p class="pf-hero-copy">${escapeHtml(detail.summary || "")}</p>
            <p class="pf-hero-hint">${escapeHtml(appState.statusText || "")}</p>
        </section>
        <section class="pf-step-guide">
            ${providerSection}
            ${quickTipsSection}
            ${descriptionSection}
            ${highlightSection}
            ${checklistSection}
            ${
                stepFields
                    ? `
                ${renderInlineNotes({
                    textareaId: "step-notes-input",
                    value: stepNotes,
                    permanentId: "step-notes-permanent",
                    isPermanent: stepNotesPermanent,
                    saveButtonId: "step-notes-save"
                })}
                ${renderSignoff({
                    reviewerInputId: "step-reviewer-name",
                    reviewerValue: stepReviewer,
                    signoffInputId: `step-signoff-${stepId}`,
                    signoffValue: stepSignOff,
                    isComplete: stepComplete,
                    saveButtonId: `step-signoff-save-${stepId}`,
                    completeButtonId: `step-signoff-toggle-${stepId}`,
                    subtext: "Ready to move on? Save and click Done when finished."
                })}
            `
                    : ""
            }
        </section>
    `;
}

function renderStepCard(step, index) {
    const isActive = appState.focusedIndex === index ? "pf-step-card--active" : "";
    const icon = getStepIconSvg(getStepType(step.id));
    return `
        <article class="pf-step-card pf-clickable ${isActive}" data-step-card data-step-index="${index}" data-step-id="${step.id}">
            <p class="pf-step-index">Step ${step.id}</p>
            <h3 class="pf-step-title">${icon ? `${icon}` : ""}${escapeHtml(step.title)}</h3>
        </article>
    `;
}

function renderFooter() {
    return `
        <footer class="pf-brand-footer">
            <div class="pf-brand-text">
                <div class="pf-brand-label">prairie.forge</div>
                <div class="pf-brand-meta">© Prairie Forge LLC, 2025. All rights reserved. Version ${VERSION}</div>
                <button type="button" class="pf-config-link" id="showConfigSheets">CONFIGURATION</button>
            </div>
        </footer>
    `;
}

function getStepType(stepId) {
    if (stepId === 0) return "config";
    if (stepId === 1) return "import";
    if (stepId === 2) return "headcount";
    if (stepId === 3) return "validate";
    if (stepId === 4) return "review";
    if (stepId === 5) return "journal";
    if (stepId === 6) return "archive";
    return "";
}

function bindSharedInteractions() {
    document.getElementById("nav-home")?.addEventListener("click", () => {
        returnHome();
        document.getElementById("pf-hero")?.scrollIntoView({ behavior: "smooth", block: "start" });
    });
    document.getElementById("nav-selector")?.addEventListener("click", () => {
        window.location.href = "../module-selector/index.html";
    });
    document.getElementById("nav-prev")?.addEventListener("click", () => moveFocus(-1));
    document.getElementById("nav-next")?.addEventListener("click", () => moveFocus(1));
    
    // Quick Access hamburger menu toggle
    const quickToggle = document.getElementById("nav-quick-toggle");
    const quickDropdown = document.getElementById("quick-access-dropdown");
    
    quickToggle?.addEventListener("click", (e) => {
        e.stopPropagation();
        quickDropdown?.classList.toggle("hidden");
        quickToggle.classList.toggle("is-active");
    });
    
    // Close dropdown when clicking outside
    document.addEventListener("click", (e) => {
        if (!quickDropdown?.contains(e.target) && !quickToggle?.contains(e.target)) {
            quickDropdown?.classList.add("hidden");
            quickToggle?.classList.remove("is-active");
        }
    });
    
    // Quick Access buttons - open reference data sheets
    document.getElementById("nav-roster")?.addEventListener("click", () => {
        openReferenceSheet("SS_Employee_Roster");
        quickDropdown?.classList.add("hidden");
        quickToggle?.classList.remove("is-active");
    });
    document.getElementById("nav-accounts")?.addEventListener("click", () => {
        openReferenceSheet("SS_Chart_of_Accounts");
        quickDropdown?.classList.add("hidden");
        quickToggle?.classList.remove("is-active");
    });
    document.getElementById("nav-expense-map")?.addEventListener("click", async () => {
        quickDropdown?.classList.add("hidden");
        quickToggle?.classList.remove("is-active");
        await navigateToExpenseMapping();
    });
    
    // CONFIGURATION link - unhides SS_* sheets for config access
    document.getElementById("showConfigSheets")?.addEventListener("click", async () => {
        await unhideSystemSheets();
    });
}

/**
 * Unhide system sheets (SS_* prefix) for configuration access
 */
async function unhideSystemSheets() {
    if (typeof Excel === "undefined") {
        console.log("Excel not available");
        return;
    }
    
    try {
        await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            worksheets.load("items/name,visibility");
            await context.sync();
            
            let unhiddenCount = 0;
            worksheets.items.forEach((sheet) => {
                if (sheet.name.toUpperCase().startsWith("SS_")) {
                    sheet.visibility = Excel.SheetVisibility.visible;
                    console.log(`[Config] Made visible: ${sheet.name}`);
                    unhiddenCount++;
                }
            });
            
            await context.sync();
            
            // Activate SS_PF_Config if it exists
            const configSheet = context.workbook.worksheets.getItemOrNullObject("SS_PF_Config");
            configSheet.load("isNullObject");
            await context.sync();
            
            if (!configSheet.isNullObject) {
                configSheet.activate();
                configSheet.getRange("A1").select();
                await context.sync();
            }
            
            console.log(`[Config] ${unhiddenCount} system sheets now visible`);
        });
    } catch (error) {
        console.error("[Config] Error unhiding system sheets:", error);
    }
}

/**
 * Open a reference data sheet (creates if doesn't exist)
 * Makes sheet visible first if it's hidden
 */
async function openReferenceSheet(sheetName) {
    if (!sheetName || typeof Excel === "undefined") {
        return;
    }
    
    const defaultHeaders = {
        "SS_Employee_Roster": ["Employee", "Department", "Pay_Rate", "Status", "Hire_Date"],
        "SS_Chart_of_Accounts": ["Account_Number", "Account_Name", "Type", "Category"]
    };
    
    try {
        await Excel.run(async (context) => {
            let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
            sheet.load("isNullObject,visibility");
            await context.sync();
            
            if (sheet.isNullObject) {
                // Create the sheet with default headers
                sheet = context.workbook.worksheets.add(sheetName);
                const headers = defaultHeaders[sheetName] || ["Column1", "Column2"];
                const headerRange = sheet.getRange(`A1:${String.fromCharCode(64 + headers.length)}1`);
                headerRange.values = [headers];
                headerRange.format.font.bold = true;
                headerRange.format.fill.color = "#f0f0f0";
                headerRange.format.autofitColumns();
                await context.sync();
            } else {
                // Make sure sheet is visible before activating (it may be hidden by tab visibility)
                sheet.visibility = Excel.SheetVisibility.visible;
                await context.sync();
            }
            
            sheet.activate();
            sheet.getRange("A1").select();
            await context.sync();
            console.log(`[Quick Access] Opened sheet: ${sheetName}`);
        });
    } catch (error) {
        console.error("Error opening reference sheet:", error);
    }
}

/**
 * Navigate directly to the PR_Expense_Mapping sheet
 */
async function navigateToExpenseMapping() {
    try {
        await Excel.run(async (context) => {
            let sheet = context.workbook.worksheets.getItemOrNullObject("PR_Expense_Mapping");
            sheet.load("isNullObject,visibility");
            await context.sync();
            
            if (sheet.isNullObject) {
                // Create the sheet with default headers
                sheet = context.workbook.worksheets.add("PR_Expense_Mapping");
                const headers = ["Expense_Category", "GL_Account", "Description", "Active"];
                const headerRange = sheet.getRange("A1:D1");
                headerRange.values = [headers];
                headerRange.format.font.bold = true;
            } else {
                // Make sure sheet is visible before activating
                sheet.visibility = Excel.SheetVisibility.visible;
                await context.sync();
            }
            
            sheet.activate();
            sheet.getRange("A1").select();
            await context.sync();
            console.log("[Quick Access] Opened PR_Expense_Mapping");
        });
    } catch (error) {
        console.error("Error navigating to PR_Expense_Mapping:", error);
    }
}

function bindHomeInteractions() {
    document.querySelectorAll("[data-step-card]").forEach((card) => {
        const index = Number(card.getAttribute("data-step-index"));
        card.addEventListener("click", () => focusStep(index));
    });
}

function bindConfigInteractions() {
    // User Name field - writes to SS_PF_Config as PR_Reviewer_Name
    const userNameInput = document.getElementById("config-user-name");
    userNameInput?.addEventListener("change", (event) => {
        const value = event.target.value.trim();
        scheduleConfigWrite(CONFIG_REVIEWER_FIELD, value);
        // Also update the reviewer field in the signoff section if it's empty
        const reviewerInput = document.getElementById("config-reviewer-name");
        if (reviewerInput && !reviewerInput.value) {
            reviewerInput.value = value;
        }
    });

    const payrollInput = document.getElementById("config-payroll-date");
    payrollInput?.addEventListener("change", (event) => {
        const value = event.target.value || "";
        // Always use the primary field name to avoid duplicate rows
        scheduleConfigWrite("PR_Payroll_Date", value);
        if (!value) return;
        if (!configState.overrides.accountingPeriod) {
            const derivedPeriod = deriveAccountingPeriod(value);
            if (derivedPeriod) {
                const periodInput = document.getElementById("config-accounting-period");
                if (periodInput) periodInput.value = derivedPeriod;
                scheduleConfigWrite("PR_Accounting_Period", derivedPeriod);
            }
        }
        if (!configState.overrides.jeId) {
            const derivedJe = deriveJeId(value);
            if (derivedJe) {
                const jeInput = document.getElementById("config-je-id");
                if (jeInput) jeInput.value = derivedJe;
                scheduleConfigWrite("PR_Journal_Entry_ID", derivedJe);
            }
        }
    });

    const stepFields = getStepNoteFields(0);

    const periodInput = document.getElementById("config-accounting-period");
    periodInput?.addEventListener("change", (event) => {
        configState.overrides.accountingPeriod = Boolean(event.target.value);
        scheduleConfigWrite("PR_Accounting_Period", event.target.value || "");
    });

    const jeInput = document.getElementById("config-je-id");
    jeInput?.addEventListener("change", (event) => {
        configState.overrides.jeId = Boolean(event.target.value);
        scheduleConfigWrite("PR_Journal_Entry_ID", event.target.value.trim());
    });

    document.getElementById("config-company-name")?.addEventListener("change", (event) => {
        scheduleConfigWrite("SS_Company_Name", event.target.value.trim());
    });

    document.getElementById("config-payroll-provider")?.addEventListener("change", (event) => {
        const value = event.target.value.trim();
        scheduleConfigWrite(PAYROLL_PROVIDER_FIELD, value);
    });

    document.getElementById("config-accounting-link")?.addEventListener("change", (event) => {
        scheduleConfigWrite("SS_Accounting_Software", event.target.value.trim());
    });

    const notesInput = document.getElementById("config-notes");
    notesInput?.addEventListener("input", (event) => {
        if (stepFields) {
            scheduleConfigWrite(stepFields.note, event.target.value, { debounceMs: 400 });
        }
    });

    if (stepFields) {
        const lockButton = document.getElementById("config-notes-permanent");
        if (lockButton) {
            lockButton.addEventListener("click", () => {
                const nextState = !lockButton.classList.contains("is-locked");
                updateLockButtonVisual(lockButton, nextState);
                setNotePermanent(stepFields.note, nextState);
            });
            updateLockButtonVisual(lockButton, isFieldPermanent(stepFields.note));
        }

        const notesSaveBtn = document.getElementById("config-notes-save");
        notesSaveBtn?.addEventListener("click", () => {
            if (!notesInput) return;
            scheduleConfigWrite(stepFields.note, notesInput.value);
            updateSaveButtonState(notesSaveBtn, true);
        });
    }

    const reviewerInput = document.getElementById("config-reviewer-name");
    reviewerInput?.addEventListener("change", (event) => {
        const value = event.target.value.trim();
        if (stepFields) {
            scheduleConfigWrite(stepFields.reviewer, value);
        }
        scheduleConfigWrite(CONFIG_REVIEWER_FIELD, value);
        const signoffInput = document.getElementById("config-signoff-date");
        if (value && signoffInput && !signoffInput.value) {
            const today = todayIso();
            signoffInput.value = today;
            if (stepFields) {
                scheduleConfigWrite(stepFields.signOff, today);
            }
        }
    });

    document.getElementById("config-signoff-date")?.addEventListener("change", (event) => {
        if (stepFields) {
            scheduleConfigWrite(stepFields.signOff, event.target.value || "");
        }
    });

    const signoffSaveBtn = document.getElementById("config-signoff-save");
    signoffSaveBtn?.addEventListener("click", () => {
        const reviewerValue = reviewerInput?.value?.trim() || "";
        const signoffInput = document.getElementById("config-signoff-date");
        const signoffValue = signoffInput?.value || "";
        if (stepFields) {
            scheduleConfigWrite(stepFields.reviewer, reviewerValue);
            scheduleConfigWrite(stepFields.signOff, signoffValue);
        }
        scheduleConfigWrite(CONFIG_REVIEWER_FIELD, reviewerValue);
        updateSaveButtonState(signoffSaveBtn, true);
    });

    initSaveTracking();

    if (stepFields) {
        // Calculate isStepComplete for Step 0
        const signOffValue = getConfigValue(stepFields.signOff);
        const completeValue = getConfigValue(STEP_COMPLETE_FIELDS[0]);
        const isStepComplete = Boolean(signOffValue || completeValue === "Y" || completeValue === true);
        console.log(`[Step 0] Binding signoff toggle. signOff="${signOffValue}", complete="${completeValue}", isComplete=${isStepComplete}`);
        
        bindSignoffToggle({
            buttonId: "config-signoff-toggle",
            inputId: "config-signoff-date",
            fieldName: stepFields.signOff,
            completeField: STEP_COMPLETE_FIELDS[0],
            initialActive: isStepComplete,
            stepId: 0 // Step 0 has no prerequisites
        });
    }
}

function bindStepInteractions(stepId) {
    document.getElementById("step-back-btn")?.addEventListener("click", () => {
        returnHome();
    });
    document.getElementById("step-action-btn")?.addEventListener("click", () => {
        const detail = STEP_DETAILS.find((step) => step.id === stepId);
        window.alert(detail?.ctaLabel ? `${detail.ctaLabel} coming soon.` : "Step actions coming soon.");
    });

    if (stepId === 1) {
        document.getElementById("import-open-data-btn")?.addEventListener("click", () => openDataSheet());
        document.getElementById("import-clear-btn")?.addEventListener("click", () => clearPrDataSheet());
    }
    if (stepId === 2) {
        document.getElementById("headcount-skip-btn")?.addEventListener("click", () => {
            headcountState.skipAnalysis = !headcountState.skipAnalysis;
            const skipBtn = document.getElementById("headcount-skip-btn");
            skipBtn?.classList.toggle("is-active", headcountState.skipAnalysis);
            if (headcountState.skipAnalysis) {
                enforceHeadcountSkipNote();
            }
            updateHeadcountSignoffState();
        });
        document.getElementById("roster-run-btn")?.addEventListener("click", () => refreshHeadcountAnalysis());
        document.getElementById("roster-refresh-btn")?.addEventListener("click", () => refreshHeadcountAnalysis());
        document.getElementById("roster-review-btn")?.addEventListener("click", () => {
            const items = headcountState.roster?.mismatches || [];
            showHeadcountModal("Roster Differences", items, {
                sourceLabel: "Roster",
                targetLabel: "Payroll Data"
            });
        });
        document.getElementById("dept-review-btn")?.addEventListener("click", () => {
            const items = headcountState.departments?.mismatches || [];
            showHeadcountModal("Department Differences", items, {
                sourceLabel: "Roster",
                targetLabel: "Payroll",
                formatter: (item) => ({
                    name: item.employee,
                    source: `${item.rosterDept} → ${item.payrollDept}`,
                    isMissingFromTarget: true
                })
            });
        });
    }
    if (stepId === 3) {
        document.getElementById("validation-run-btn")?.addEventListener("click", () => prepareValidationData());
        document.getElementById("validation-refresh-btn")?.addEventListener("click", () => prepareValidationData());
        // Format on blur (leaving field) or Enter key - comparison triggers automatically
        document.getElementById("bank-amount-input")?.addEventListener("blur", handleBankAmountInput);
        document.getElementById("bank-amount-input")?.addEventListener("keydown", (e) => {
            if (e.key === "Enter") handleBankAmountInput(e);
        });
    }
    if (stepId === 5) {
        document.getElementById("je-run-btn")?.addEventListener("click", () => runJournalSummary());
        document.getElementById("je-save-btn")?.addEventListener("click", () => saveJournalSummary());
        document.getElementById("je-create-btn")?.addEventListener("click", () => createJournalDraft());
        document.getElementById("je-export-btn")?.addEventListener("click", () => exportJournalDraft());
    }
    if (stepId === 4) {
        const container = document.querySelector(".pf-step-guide");
        if (container) {
            // CoPilot API endpoint - Update this with your Supabase project URL
            const COPILOT_API_ENDPOINT = "https://your-project.supabase.co/functions/v1/copilot";
            
            // Bind CoPilot with full context provider for intelligent analysis
            bindCopilotCard(container, { 
                id: "expense-review-copilot",
                // Uncomment to enable real AI (requires Supabase Edge Function deployment)
                // apiEndpoint: COPILOT_API_ENDPOINT,
                contextProvider: createPayrollContextProvider(),
                systemPrompt: `You are Prairie Forge CoPilot, an expert financial analyst assistant for payroll expense review.

CONTEXT: You're embedded in the Payroll Recorder Excel add-in, helping accountants and CFOs review payroll data before journal entry export.

YOUR CAPABILITIES:
1. Analyze payroll expense data for accuracy and completeness
2. Identify trends, anomalies, and variances requiring attention
3. Prepare executive-ready insights and talking points
4. Validate journal entries before export to accounting system

COMMUNICATION STYLE:
- Be concise and actionable
- Use bullet points and tables for clarity
- Highlight issues with ⚠️ and successes with ✓
- Format currency as $X,XXX (no decimals for totals)
- Format percentages as X.X%
- Always end with 2-3 concrete next steps

ANALYSIS FOCUS:
- Period-over-period changes exceeding 10%
- Department cost anomalies vs historical norms
- Headcount vs payroll expense alignment
- Burden rate outliers (normal range: 15-35%)
- Missing or incomplete GL account mappings
- Data quality issues (blanks, duplicates, mismatches)

When asked about variances, explain the business drivers, not just the numbers.
When asked about readiness, be specific about what passes and what needs attention.`
            });
        }
        document.getElementById("expense-run-btn")?.addEventListener("click", () => {
            prepareExpenseReviewData();
        });
        document.getElementById("expense-refresh-btn")?.addEventListener("click", () => {
            prepareExpenseReviewData();
        });
    }

    const fields = getStepNoteFields(stepId);
    console.log(`[Step ${stepId}] Binding interactions, fields:`, fields);
    if (fields) {
        // Handle Step 1's different ID pattern
        const notesInputId = stepId === 1 ? "step-notes-1" : "step-notes-input";
        const notesInput = document.getElementById(notesInputId);
        console.log(`[Step ${stepId}] Notes input found:`, !!notesInput, `(id: ${notesInputId})`);
        const notesSaveBtn =
            stepId === 1
                ? document.getElementById("step-notes-save-1")
                : stepId === 2
                    ? document.getElementById("step-notes-save-2")
                    : stepId === 3
                        ? document.getElementById("step-notes-save-3")
                    : stepId === 4
                        ? document.getElementById("step-notes-save-4")
                    : stepId === 5
                        ? document.getElementById("step-notes-save-5")
                        : document.getElementById("step-notes-save");
        notesInput?.addEventListener("input", (event) => {
            scheduleConfigWrite(fields.note, event.target.value, { debounceMs: 400 });
            if (stepId === 2) {
                if (headcountState.skipAnalysis) {
                    enforceHeadcountSkipNote();
                }
                updateHeadcountSignoffState();
            }
        });
        notesSaveBtn?.addEventListener("click", () => {
            if (!notesInput) return;
            scheduleConfigWrite(fields.note, notesInput.value);
            updateSaveButtonState(notesSaveBtn, true);
        });
        // Handle Step 1's different ID pattern for reviewer
        const reviewerInputId = stepId === 1 ? "step-reviewer-1" : "step-reviewer-name";
        const reviewerInput = document.getElementById(reviewerInputId);
        reviewerInput?.addEventListener("change", (event) => {
            const value = event.target.value.trim();
            scheduleConfigWrite(fields.reviewer, value);
            const signoffInput =
                stepId === 1
                    ? document.getElementById("step-signoff-1")
                    : stepId === 2
                        ? document.getElementById("step-signoff-date")
                        : stepId === 3
                            ? document.getElementById("step-signoff-3")
                            : stepId === 4
                                ? document.getElementById("step-signoff-4")
                                : stepId === 5
                                    ? document.getElementById("step-signoff-5")
                                    : document.getElementById(`step-signoff-${stepId}`);
            if (value && signoffInput && !signoffInput.value) {
                const today = todayIso();
                signoffInput.value = today;
                scheduleConfigWrite(fields.signOff, today);
            }
        });
        const signoffInputId =
            stepId === 1
                ? "step-signoff-1"
                : stepId === 2
                    ? "step-signoff-date"
                    : stepId === 3
                        ? "step-signoff-3"
                        : stepId === 4
                            ? "step-signoff-4"
                            : stepId === 5
                                ? "step-signoff-5"
                                : `step-signoff-${stepId}`;
        console.log(`[Step ${stepId}] Signoff input ID: ${signoffInputId}, found:`, !!document.getElementById(signoffInputId));
        document.getElementById(signoffInputId)?.addEventListener("change", (event) => {
            scheduleConfigWrite(fields.signOff, event.target.value || "");
        });
        // Handle Step 1's different lock button ID
        const lockButtonId = stepId === 1 ? "step-notes-lock-1" : "step-notes-permanent";
        const lockButton = document.getElementById(lockButtonId);
        if (lockButton) {
            lockButton.addEventListener("click", () => {
                const nextState = !lockButton.classList.contains("is-locked");
                updateLockButtonVisual(lockButton, nextState);
                setNotePermanent(fields.note, nextState);
                if (stepId === 2) updateHeadcountSignoffState();
            });
            updateLockButtonVisual(lockButton, isFieldPermanent(fields.note));
        }
        const signoffSaveBtn =
            stepId === 1
                ? document.getElementById("step-signoff-save-1")
                : stepId === 2
                    ? document.getElementById("headcount-signoff-save")
                    : stepId === 3
                        ? document.getElementById("step-signoff-save-3")
                        : stepId === 4
                            ? document.getElementById("step-signoff-save-4")
                            : stepId === 5
                                ? document.getElementById("step-signoff-save-5")
                                : document.getElementById(`step-signoff-save-${stepId}`);
        signoffSaveBtn?.addEventListener("click", () => {
            const reviewerValue = reviewerInput?.value?.trim() || "";
            const signoffValue = document.getElementById(signoffInputId)?.value || "";
            scheduleConfigWrite(fields.reviewer, reviewerValue);
            scheduleConfigWrite(fields.signOff, signoffValue);
            updateSaveButtonState(signoffSaveBtn, true);
        });
        initSaveTracking();
        const completeField = STEP_COMPLETE_FIELDS[stepId];
        const initialCompleteFlag = completeField ? Boolean(getConfigValue(completeField)) : false;
        const initialSignOff = getConfigValue(fields.signOff);
        const toggleButtonId = stepId === 1
                    ? "step-signoff-toggle-1"
                    : stepId === 2
                        ? "headcount-signoff-toggle"
                        : stepId === 3
                            ? "validation-signoff-toggle"
                            : stepId === 4
                                ? "expense-signoff-toggle"
                                : stepId === 5
                                    ? "step-signoff-toggle-5"
                                    : `step-signoff-toggle-${stepId}`;
        console.log(`[Step ${stepId}] Toggle button ID: ${toggleButtonId}, found:`, !!document.getElementById(toggleButtonId));
        bindSignoffToggle({
            buttonId: toggleButtonId,
            inputId: signoffInputId,
            fieldName: fields.signOff,
            completeField,
            requireNotesCheck: stepId === 2 ? isHeadcountNotesRequired : null,
            initialActive: Boolean(initialSignOff || initialCompleteFlag),
            stepId, // Pass stepId for sequential validation
            onComplete:
                stepId === 3
                    ? handleValidationComplete
                    : stepId === 4
                        ? handleExpenseReviewComplete
                        : stepId === 2
                            ? handleHeadcountSignoff
                            : null
        });
    }

    if (stepId === 2) {
        updateHeadcountSignoffState();
    }
    if (stepId === 6) {
        document.getElementById("archive-run-btn")?.addEventListener("click", handleArchiveRun);
    }
}

function focusStep(index) {
    if (Number.isNaN(index) || index < 0 || index >= WORKFLOW_STEPS.length) return;
    const step = WORKFLOW_STEPS[index];
    if (!step) return;
    pendingScrollIndex = index;
    const view = step.id === 0 ? "config" : "step";
    setState({ focusedIndex: index, activeView: view, activeStepId: step.id });
    const sheetName = STEP_SHEET_MAP[step.id];
    if (sheetName) {
        activateWorksheet(sheetName);
    }
    if (step.id === 2 && !headcountState.hasAnalyzed) {
        refreshHeadcountAnalysis();
    }
    // Step 3 (Validate & Reconcile) - User must click Run/Refresh to trigger
    // Step 4 (Expense Review) - User must click Run/Refresh to trigger
}

function moveFocus(delta) {
    // When on home view, forward should go to step 0 (Config) first
    if (appState.activeView === "home" && delta > 0) {
        focusStep(0);
        return;
    }
    const next = appState.focusedIndex + delta;
    const clamped = Math.max(0, Math.min(WORKFLOW_STEPS.length - 1, next));
    focusStep(clamped);
}

function scrollFocusedIntoView() {
    if (appState.activeView !== "home") return;
    if (pendingScrollIndex === null) return;
    const card = document.querySelector(`[data-step-card][data-step-index="${pendingScrollIndex}"]`);
    pendingScrollIndex = null;
    card?.scrollIntoView({ behavior: "smooth", block: "center" });
}

function openConfigView() {
    focusStep(0);
}

async function returnHome() {
    // Activate the module homepage sheet
    const homepageConfig = getHomepageConfig(MODULE_KEY);
    await activateHomepageSheet(homepageConfig.sheetName, homepageConfig.title, homepageConfig.subtitle);
    
    setState({ activeView: "home", activeStepId: null });
}

function setState(partial) {
    Object.assign(appState, partial);
    renderApp();
}

function getReviewerDefault() {
    // Fallback to legacy field if the new one is missing
    return getConfigValue(CONFIG_REVIEWER_FIELD) || getConfigValue("SS_Default_Reviewer") || "";
}

function updateActionToggleState(button, isActive) {
    if (!button) return;
    button.classList.toggle("is-active", isActive);
    button.setAttribute("aria-pressed", String(isActive));
}

function markJeSaveState(isSaved) {
    const btn = document.getElementById("je-save-btn");
    if (!btn) return;
    btn.classList.toggle("is-saved", isSaved);
}

/**
 * Get current step completion status for sequential validation
 * @returns {Object} Map of step IDs to boolean completion status
 */
function getStepCompletionStatus() {
    const status = {};
    console.log("[Signoff] Checking step completion status...");
    Object.keys(STEP_NOTES_FIELDS).forEach(stepIdStr => {
        const stepId = parseInt(stepIdStr, 10);
        const fields = STEP_NOTES_FIELDS[stepId];
        if (!fields) {
            status[stepId] = false;
            return;
        }
        // A step is complete if it has a sign-off date OR is explicitly marked complete
        const signOffValue = getConfigValue(fields.signOff);
        const completeField = STEP_COMPLETE_FIELDS[stepId];
        const completeValue = getConfigValue(completeField);
        const isComplete = Boolean(signOffValue) || completeValue === "Y" || completeValue === true;
        status[stepId] = isComplete;
        console.log(`[Signoff] Step ${stepId}: signOff="${signOffValue}", complete="${completeValue}" → ${isComplete ? "COMPLETE" : "pending"}`);
    });
    console.log("[Signoff] Status summary:", status);
    return status;
}

function bindSignoffToggle({
    buttonId,
    inputId,
    fieldName,
    completeField,
    requireNotesCheck,
    onComplete,
    initialActive = false,
    stepId = null // NEW: Step ID for sequential validation
}) {
    const button = document.getElementById(buttonId);
    if (!button) {
        console.warn(`[Signoff] Button not found: ${buttonId}`);
        return;
    }
    const input = inputId ? document.getElementById(inputId) : null;
    const initial = initialActive || Boolean(input?.value);
    updateActionToggleState(button, initial);
    console.log(`[Signoff] Bound ${buttonId}, initial active: ${initial}, stepId: ${stepId}`);
    
    // Handle Done button click
    button.addEventListener("click", () => {
        console.log(`[Signoff] Done button clicked: ${buttonId}, stepId: ${stepId}`);
        
        // Check sequential completion if stepId is provided
        if (stepId !== null && stepId > 0) {
            const completionStatus = getStepCompletionStatus();
            const { canComplete, message } = canCompleteStep(stepId, completionStatus);
            
            // Only block if trying to COMPLETE (not uncomplete)
            const isCurrentlyActive = button.classList.contains("is-active");
            console.log(`[Signoff] canComplete: ${canComplete}, isCurrentlyActive: ${isCurrentlyActive}`);
            if (!isCurrentlyActive && !canComplete) {
                console.log(`[Signoff] BLOCKED: ${message}`);
                showBlockedToast(message);
                return;
            }
        }
        
        if (requireNotesCheck && !requireNotesCheck()) {
            window.alert("Please add notes before completing this step.");
            return;
        }
        const nextActive = !button.classList.contains("is-active");
        console.log(`[Signoff] ${buttonId} clicked, toggling to: ${nextActive}`);
        updateActionToggleState(button, nextActive);
        if (input) {
            input.value = nextActive ? todayIso() : "";
        }
        if (fieldName) {
            const dateValue = nextActive ? todayIso() : "";
            console.log(`[Signoff] Writing ${fieldName} = "${dateValue}"`);
            scheduleConfigWrite(fieldName, dateValue);
        }
        if (completeField) {
            const completeValue = nextActive ? "Y" : "";
            console.log(`[Signoff] Writing ${completeField} = "${completeValue}"`);
            scheduleConfigWrite(completeField, completeValue);
        }
        if (nextActive && typeof onComplete === "function") {
            onComplete();
        }
    });
    
    // Handle manual date input change - sync the button state
    if (input) {
        input.addEventListener("change", () => {
            const hasDate = Boolean(input.value);
            const isCurrentlyActive = button.classList.contains("is-active");
            if (hasDate !== isCurrentlyActive) {
                console.log(`[Signoff] Date input changed, syncing button to: ${hasDate}`);
                updateActionToggleState(button, hasDate);
                if (fieldName) {
                    scheduleConfigWrite(fieldName, input.value || "");
                }
                if (completeField) {
                    scheduleConfigWrite(completeField, hasDate ? "Y" : "");
                }
            }
        });
    }
}

async function openConfigurationSheet() {
    if (!hasExcelRuntime()) {
        window.alert("Open this module inside Excel to edit configuration settings.");
        return;
    }
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(MODULE_CONFIG_SHEET);
            sheet.activate();
            sheet.getRange("A1").select();
            await context.sync();
        });
        setState({ statusText: `${MODULE_CONFIG_SHEET} opened.` });
    } catch (error) {
        console.error("Unable to open configuration sheet", error);
        window.alert(`Unable to open ${MODULE_CONFIG_SHEET}. Confirm the sheet exists in this workbook.`);
    }
}

async function openDataSheet() {
    if (!hasExcelRuntime()) {
        window.alert("Open this module inside Excel to access the data sheet.");
        return;
    }
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(SHEET_NAMES.DATA);
            sheet.activate();
            sheet.getRange("A1").select();
            await context.sync();
        });
    } catch (error) {
        console.error("Unable to open PR_Data sheet", error);
        window.alert(`Unable to open ${SHEET_NAMES.DATA}. Confirm the sheet exists in this workbook.`);
    }
}

async function clearPrDataSheet() {
    if (!hasExcelRuntime()) {
        window.alert("Open this module inside Excel to clear data.");
        return;
    }
    const confirmed = window.confirm("Are you sure you want to clear all data from PR_Data? This cannot be undone.");
    if (!confirmed) return;
    
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(SHEET_NAMES.DATA);
            const usedRange = sheet.getUsedRangeOrNullObject();
            usedRange.load("isNullObject");
            await context.sync();
            
            if (!usedRange.isNullObject) {
                // Clear all except header row (row 1)
                const dataRange = sheet.getRange("A2:Z10000");
                dataRange.clear(Excel.ClearApplyTo.contents);
                await context.sync();
            }
            
            sheet.activate();
            sheet.getRange("A1").select();
            await context.sync();
        });
        window.alert("PR_Data cleared successfully.");
    } catch (error) {
        console.error("Unable to clear PR_Data sheet", error);
        window.alert("Unable to clear PR_Data. Please try again.");
    }
}

async function getConfigTable(context) {
    if (!CONFIG_TABLE_CANDIDATES.length) return null;
    if (resolvedConfigTableName) {
        const existing = context.workbook.tables.getItemOrNullObject(resolvedConfigTableName);
        existing.load("name");
        await context.sync();
        if (!existing.isNullObject) {
            return existing;
        }
        resolvedConfigTableName = null;
    }
    const tables = context.workbook.tables;
    tables.load("items/name");
    await context.sync();

    const foundTableNames = tables.items?.map((t) => t.name) || [];
    
    // Debug logging - visible in browser console (F12)
    console.log("[Payroll] Looking for config table:", CONFIG_TABLE_CANDIDATES);
    console.log("[Payroll] Found tables in workbook:", foundTableNames);

    const match = tables.items?.find((table) => CONFIG_TABLE_CANDIDATES.includes(table.name));
    if (!match) {
        console.warn("[Payroll] ⚠️ CONFIG TABLE NOT FOUND!");
        console.warn("[Payroll] Expected table named: SS_PF_Config");
        console.warn("[Payroll] Available tables:", foundTableNames);
        console.warn("[Payroll] To fix: Select your data in SS_PF_Config sheet → Insert → Table → Name it 'SS_PF_Config'");
        return null;
    }
    console.log("[Payroll] ✓ Config table found:", match.name);
    resolvedConfigTableName = match.name;
    return context.workbook.tables.getItem(match.name);
}

async function loadConfigurationValues() {
    if (!hasExcelRuntime()) {
        configState.loaded = true;
        return;
    }
    try {
        await Excel.run(async (context) => {
            const table = await getConfigTable(context);
            if (!table) {
                console.warn("Payroll Recorder: SS_PF_Config table is missing.");
                configState.loaded = true;
                return;
            }
            const body = table.getDataBodyRange();
            body.load("values");
            await context.sync();
            const rows = body.values || [];
            const map = {};
            const permanents = {};
            rows.forEach((row) => {
                const field = normalizeFieldName(row[CONFIG_COLUMNS.FIELD]);
                if (!field) return;
                map[field] = row[CONFIG_COLUMNS.VALUE] ?? "";
                permanents[field] = row[CONFIG_COLUMNS.PERMANENT] ?? "";
            });
            configState.values = map;
            configState.permanents = permanents;
            // Check both new and legacy field names for overrides
            configState.overrides.accountingPeriod = Boolean(map.PR_Accounting_Period || map.Accounting_Period);
            configState.overrides.jeId = Boolean(map.PR_Journal_Entry_ID || map.Journal_Entry_ID);
            configState.loaded = true;
        });
    } catch (error) {
        console.warn("Payroll Recorder: unable to load PF_Config table.", error);
        configState.loaded = true;
    }
}

function getConfigValue(field) {
    return configState.values[field] ?? "";
}

function resolvePayrollDateFieldName() {
    const keys = Object.keys(configState.values || {});
    const match = PAYROLL_DATE_ALIASES.find((alias) => keys.includes(alias));
    return match || PAYROLL_DATE_ALIASES[0];
}

function getPayrollDateValue() {
    return getConfigValue(resolvePayrollDateFieldName());
}

function getPayrollProviderLink() {
    // Check module-specific field first, then shared field as fallback
    return (getConfigValue(PAYROLL_PROVIDER_FIELD) || getConfigValue("Payroll_Provider_Link") || "").trim();
}

function isFieldPermanent(field) {
    return parseBooleanFlag(configState.permanents[field]);
}

function isStepCompleteFromConfig(stepId) {
    const field = STEP_COMPLETE_FIELDS[stepId];
    if (!field) return false;
    return parseBooleanFlag(getConfigValue(field));
}

function setNotePermanent(field, isPermanent) {
    const normalizedField = normalizeFieldName(field);
    if (!normalizedField) return;
    configState.permanents[normalizedField] = isPermanent ? "Y" : "N";
    void writeConfigPermanent(normalizedField, isPermanent ? "Y" : "N");
}

function parseBooleanFlag(value) {
    const normalized = String(value ?? "").trim().toLowerCase();
    return normalized === "true" || normalized === "y" || normalized === "yes" || normalized === "1";
}

function normalizeFieldName(value) {
    return String(value ?? "").trim();
}

function isNoiseName(value) {
    const normalized = String(value ?? "").trim().toLowerCase();
    if (!normalized) return true;
    return (
        normalized.includes("total") ||
        normalized.includes("totals") ||
        normalized.includes("grand total") ||
        normalized.includes("subtotal") ||
        normalized.includes("summary")
    );
}

function formatDateInput(value) {
    if (!value) return "";
    const parts = parseDateInput(value);
    if (!parts) return "";
    return `${parts.year}-${String(parts.month).padStart(2, "0")}-${String(parts.day).padStart(2, "0")}`;
}

function deriveAccountingPeriod(payrollDate) {
    const parts = parseDateInput(payrollDate);
    if (!parts) return "";
    // Validate year is reasonable (1900-2100)
    if (parts.year < 1900 || parts.year > 2100) {
        console.warn("deriveAccountingPeriod - Invalid year:", parts.year, "from input:", payrollDate);
        return "";
    }
    const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    // Format: "Dec 2025" - use full 4-digit year to prevent Excel date interpretation
    return `${monthNames[parts.month - 1]} ${parts.year}`;
}

function deriveJeId(payrollDate) {
    const parts = parseDateInput(payrollDate);
    if (!parts) return "";
    // Validate year is reasonable (1900-2100)
    if (parts.year < 1900 || parts.year > 2100) {
        console.warn("deriveJeId - Invalid year:", parts.year, "from input:", payrollDate);
        return "";
    }
    return `PR-AUTO-${parts.year}-${String(parts.month).padStart(2, "0")}-${String(parts.day).padStart(2, "0")}`;
}

function todayIso() {
    return formatDateFromDate(new Date());
}

function scheduleConfigWrite(fieldName, value, options = {}) {
    const normalizedField = normalizeFieldName(fieldName);
    configState.values[normalizedField] = value ?? "";
    const delay = options.debounceMs ?? 0;
    if (!delay) {
        const existing = pendingWrites.get(normalizedField);
        if (existing) clearTimeout(existing);
        pendingWrites.delete(normalizedField);
        void writeConfigValue(normalizedField, value ?? "");
        return;
    }
    if (pendingWrites.has(normalizedField)) {
        clearTimeout(pendingWrites.get(normalizedField));
    }
    const timer = setTimeout(() => {
        pendingWrites.delete(normalizedField);
        void writeConfigValue(normalizedField, value ?? "");
    }, delay);
    pendingWrites.set(normalizedField, timer);
}

// Fields that should be forced to Text format to prevent Excel auto-conversion
const TEXT_FORMAT_FIELDS = [
    "PR_Accounting_Period",
    "PTO_Accounting_Period",
    "Accounting_Period"
];

async function writeConfigValue(fieldName, value) {
    const normalizedField = normalizeFieldName(fieldName);
    configState.values[normalizedField] = value ?? "";
    console.log(`[Payroll] Writing config: ${normalizedField} = "${value}"`);
    if (!hasExcelRuntime()) {
        console.warn("[Payroll] Excel runtime not available - cannot write");
        return;
    }
    
    // Check if this field needs text formatting
    const forceTextFormat = TEXT_FORMAT_FIELDS.some(f => 
        normalizedField === f || normalizedField.toLowerCase() === f.toLowerCase()
    );
    
    try {
        await Excel.run(async (context) => {
            const table = await getConfigTable(context);
            if (!table) {
                console.error("[Payroll] ❌ Cannot write - config table not found");
                return;
            }
            const body = table.getDataBodyRange();
            const headerRange = table.getHeaderRowRange();
            body.load("values");
            headerRange.load("values");
            await context.sync();

            const headers = headerRange.values[0] || [];
            const rows = body.values || [];
            const columnCount = headers.length;
            console.log(`[Payroll] Table has ${rows.length} rows, ${columnCount} columns`);

            // Find ALL matching rows (to handle duplicates)
            const matchingIndices = [];
            rows.forEach((row, idx) => {
                if (normalizeFieldName(row[CONFIG_COLUMNS.FIELD]) === normalizedField) {
                    matchingIndices.push(idx);
                }
            });

            if (matchingIndices.length === 0) {
                // No existing row - add new one
                configState.permanents[normalizedField] = configState.permanents[normalizedField] ?? DEFAULT_CONFIG_PERMANENT;
                const newRow = new Array(columnCount).fill("");
                if (CONFIG_COLUMNS.TYPE >= 0 && CONFIG_COLUMNS.TYPE < columnCount) newRow[CONFIG_COLUMNS.TYPE] = DEFAULT_CONFIG_TYPE;
                if (CONFIG_COLUMNS.FIELD >= 0 && CONFIG_COLUMNS.FIELD < columnCount) newRow[CONFIG_COLUMNS.FIELD] = normalizedField;
                if (CONFIG_COLUMNS.VALUE >= 0 && CONFIG_COLUMNS.VALUE < columnCount) newRow[CONFIG_COLUMNS.VALUE] = value ?? "";
                if (CONFIG_COLUMNS.PERMANENT >= 0 && CONFIG_COLUMNS.PERMANENT < columnCount) newRow[CONFIG_COLUMNS.PERMANENT] = DEFAULT_CONFIG_PERMANENT;
                console.log(`[Payroll] Adding NEW row:`, newRow);
                table.rows.add(null, [newRow]);
                await context.sync();
                
                // Force text format for specific fields to prevent Excel date conversion
                if (forceTextFormat) {
                    // Get the newly added row (last row in table)
                    const tableRows = table.rows;
                    tableRows.load("count");
                    await context.sync();
                    const lastRowIdx = tableRows.count - 1;
                    const newRowRange = table.rows.getItemAt(lastRowIdx).getRange();
                    const valueCell = newRowRange.getCell(0, CONFIG_COLUMNS.VALUE);
                    valueCell.numberFormat = [["@"]]; // Text format
                    valueCell.values = [[value ?? ""]]; // Re-write to apply format
                    await context.sync();
                    console.log(`[Payroll] ✓ Applied text format to ${normalizedField}`);
                }
                
                console.log(`[Payroll] ✓ New row added for ${normalizedField}`);
            } else {
                // Update the first matching row
                const targetIndex = matchingIndices[0];
                console.log(`[Payroll] Updating existing row ${targetIndex} for ${normalizedField}`);
                const targetCell = body.getCell(targetIndex, CONFIG_COLUMNS.VALUE);
                
                // Force text format for specific fields
                if (forceTextFormat) {
                    targetCell.numberFormat = [["@"]]; // Text format
                }
                targetCell.values = [[value ?? ""]];
                await context.sync();
                console.log(`[Payroll] ✓ Updated ${normalizedField}`);

                // Delete duplicate rows (in reverse order to maintain indices)
                if (matchingIndices.length > 1) {
                    console.log(`[Payroll] Found ${matchingIndices.length - 1} duplicate rows for ${normalizedField}, removing...`);
                    const duplicateIndices = matchingIndices.slice(1).reverse();
                    for (const dupIdx of duplicateIndices) {
                        try {
                            table.rows.getItemAt(dupIdx).delete();
                        } catch (e) {
                            console.warn(`[Payroll] Could not delete duplicate row ${dupIdx}:`, e.message);
                        }
                    }
                    await context.sync();
                    console.log(`[Payroll] ✓ Removed duplicate rows for ${normalizedField}`);
                }
            }
        });
    } catch (error) {
        console.error(`[Payroll] ❌ Write failed for ${fieldName}:`, error);
    }
}

async function writeConfigPermanent(fieldName, marker) {
    const normalizedField = normalizeFieldName(fieldName);
    if (!normalizedField) return;
    if (!hasExcelRuntime()) return;
    // Store in local state
    configState.permanents[normalizedField] = marker;
    try {
        await Excel.run(async (context) => {
            const table = await getConfigTable(context);
            if (!table) {
                console.warn(`Payroll Recorder: unable to locate config table when toggling ${fieldName} permanent flag.`);
                return;
            }
            const body = table.getDataBodyRange();
            body.load("values");
            await context.sync();
            const rows = body.values || [];
            const targetIndex = rows.findIndex(
                (row) => normalizeFieldName(row[CONFIG_COLUMNS.FIELD]) === normalizedField
            );
            if (targetIndex === -1) return;
            body.getCell(targetIndex, CONFIG_COLUMNS.PERMANENT).values = [[marker]];
            await context.sync();
        });
    } catch (error) {
        console.warn(`Payroll Recorder: unable to update permanent flag for ${fieldName}`, error);
    }
}

function parseDateInput(value) {
    if (!value) return null;
    
    // Handle string value
    const strValue = String(value).trim();
    
    // Try YYYY-MM-DD format first
    const isoMatch = /^(\d{4})-(\d{2})-(\d{2})/.exec(strValue);
    if (isoMatch) {
        const year = Number(isoMatch[1]);
        const month = Number(isoMatch[2]);
        const day = Number(isoMatch[3]);
        if (year && month && day) return { year, month, day };
    }
    
    // Try MM/DD/YYYY format
    const usMatch = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(strValue);
    if (usMatch) {
        const month = Number(usMatch[1]);
        const day = Number(usMatch[2]);
        const year = Number(usMatch[3]);
        if (year && month && day) return { year, month, day };
    }
    
    // Handle Excel serial date number
    // Use UTC to avoid timezone offset issues - Excel dates should be treated as UTC
    const numValue = Number(value);
    if (Number.isFinite(numValue) && numValue > 40000 && numValue < 60000) {
        // Excel serial date: days since Jan 1, 1900 (with 1900 leap year bug)
        // Convert to UTC timestamp to avoid local timezone shifting the date
        const utcDays = Math.floor(numValue - 25569); // Days since Unix epoch (Jan 1, 1970)
        const utcMs = utcDays * 86400 * 1000;
        const jsDate = new Date(utcMs);
        if (!isNaN(jsDate.getTime())) {
            // Use UTC methods to extract date components to prevent timezone shift
            const isoDate = `${jsDate.getUTCFullYear()}-${String(jsDate.getUTCMonth() + 1).padStart(2, "0")}-${String(jsDate.getUTCDate()).padStart(2, "0")}`;
            console.log("DEBUG parseDateInput - Converted Excel serial", numValue, "to", isoDate);
            return {
                year: jsDate.getUTCFullYear(),
                month: jsDate.getUTCMonth() + 1,
                day: jsDate.getUTCDate()
            };
        }
    }
    
    // Try parsing as Date string
    const dateObj = new Date(strValue);
    if (!isNaN(dateObj.getTime())) {
        return {
            year: dateObj.getFullYear(),
            month: dateObj.getMonth() + 1,
            day: dateObj.getDate()
        };
    }
    
    console.warn("DEBUG parseDateInput - Could not parse date value:", value);
    return null;
}

function formatDateFromDate(date) {
    // Use UTC methods if this date was derived from Excel serial number
    // to prevent timezone shift causing off-by-one day errors
    if (date._isUTC) {
        const year = date.getUTCFullYear();
        const month = String(date.getUTCMonth() + 1).padStart(2, "0");
        const day = String(date.getUTCDate()).padStart(2, "0");
        return `${year}-${month}-${day}`;
    }
    // For regular dates (like "today"), use local time
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
}

/**
 * Normalize a date value for lookup comparison
 * Handles Excel serial dates, Date objects, and various string formats
 * Returns YYYY-MM-DD string or null if unparseable
 */
function normalizeDateForLookup(value) {
    if (!value) return null;
    
    // If it's already a YYYY-MM-DD string, return it
    if (typeof value === "string") {
        const isoMatch = value.match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (isoMatch) {
            return `${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}`;
        }
    }
    
    // Try parsing with our date parser
    const parts = parseDateInput(value);
    if (parts) {
        return `${parts.year}-${String(parts.month).padStart(2, "0")}-${String(parts.day).padStart(2, "0")}`;
    }
    
    return null;
}

/**
 * Create a context provider for Prairie Forge CoPilot
 * Reads current payroll data to provide intelligent, contextual responses
 */
function createPayrollContextProvider() {
    return async () => {
        if (!hasExcelRuntime()) return null;
        
        try {
            return await Excel.run(async (context) => {
                const result = {
                    timestamp: new Date().toISOString(),
                    period: null,
                    summary: {},
                    departments: [],
                    recentPeriods: [],
                    dataQuality: {}
                };
                
                // Read config for period info
                const configTable = await getConfigTable(context);
                if (configTable) {
                    const configBody = configTable.getDataBodyRange();
                    configBody.load("values");
                    await context.sync();
                    
                    const configRows = configBody.values || [];
                    for (const row of configRows) {
                        const fieldName = String(row[CONFIG_COLUMNS.FIELD] || "").trim();
                        const fieldValue = row[CONFIG_COLUMNS.VALUE];
                        
                        if (fieldName.toLowerCase().includes("accounting") && fieldName.toLowerCase().includes("period")) {
                            result.period = String(fieldValue || "").trim();
                        }
                    }
                }
                
                // Read PR_Data_Clean for summary stats
                const cleanSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.DATA_CLEAN);
                cleanSheet.load("isNullObject");
                await context.sync();
                
                if (!cleanSheet.isNullObject) {
                    const range = cleanSheet.getUsedRangeOrNullObject();
                    range.load("values");
                    await context.sync();
                    
                    if (!range.isNullObject && range.values?.length > 1) {
                        const headers = range.values[0].map(h => normalizeHeader(h));
                        const data = range.values.slice(1);
                        
                        // Find relevant columns
                        const amountIdx = headers.findIndex(h => h.includes("amount"));
                        const deptIdx = pickDepartmentIndex(headers);
                        const employeeIdx = headers.findIndex(h => h.includes("employee"));
                        
                        // Calculate totals
                        let totalAmount = 0;
                        const employeeSet = new Set();
                        const deptTotals = new Map();
                        
                        for (const row of data) {
                            const amount = Number(row[amountIdx]) || 0;
                            totalAmount += amount;
                            
                            if (employeeIdx >= 0) {
                                const emp = String(row[employeeIdx] || "").trim();
                                if (emp) employeeSet.add(emp);
                            }
                            
                            if (deptIdx >= 0) {
                                const dept = String(row[deptIdx] || "").trim();
                                if (dept) {
                                    deptTotals.set(dept, (deptTotals.get(dept) || 0) + amount);
                                }
                            }
                        }
                        
                        result.summary = {
                            total: totalAmount,
                            employeeCount: employeeSet.size,
                            avgPerEmployee: employeeSet.size ? totalAmount / employeeSet.size : 0,
                            rowCount: data.length
                        };
                        
                        // Department breakdown
                        result.departments = Array.from(deptTotals.entries())
                            .map(([name, total]) => ({
                                name,
                                total,
                                percentOfTotal: totalAmount ? (total / totalAmount) : 0
                            }))
                            .sort((a, b) => b.total - a.total);
                        
                        result.dataQuality.dataCleanReady = true;
                        result.dataQuality.rowCount = data.length;
                    }
                }
                
                // Read PR_Archive_Summary for trend data
                const archiveSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.ARCHIVE_SUMMARY);
                archiveSheet.load("isNullObject");
                await context.sync();
                
                if (!archiveSheet.isNullObject) {
                    const archiveRange = archiveSheet.getUsedRangeOrNullObject();
                    archiveRange.load("values");
                    await context.sync();
                    
                    if (!archiveRange.isNullObject && archiveRange.values?.length > 1) {
                        const headers = archiveRange.values[0].map(h => normalizeHeader(h));
                        const periodIdx = headers.findIndex(h => h.includes("period"));
                        const totalIdx = headers.findIndex(h => h.includes("total"));
                        
                        result.recentPeriods = archiveRange.values.slice(1, 6).map(row => ({
                            period: row[periodIdx] || "",
                            total: Number(row[totalIdx]) || 0
                        }));
                        
                        result.dataQuality.archiveAvailable = true;
                        result.dataQuality.periodsAvailable = result.recentPeriods.length;
                    }
                }
                
                // Check JE Draft status
                const jeSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.JE_DRAFT);
                jeSheet.load("isNullObject");
                await context.sync();
                
                if (!jeSheet.isNullObject) {
                    const jeRange = jeSheet.getUsedRangeOrNullObject();
                    jeRange.load("values");
                    await context.sync();
                    
                    if (!jeRange.isNullObject && jeRange.values?.length > 1) {
                        const headers = jeRange.values[0].map(h => normalizeHeader(h));
                        const debitIdx = headers.findIndex(h => h.includes("debit"));
                        const creditIdx = headers.findIndex(h => h.includes("credit"));
                        
                        let totalDebit = 0;
                        let totalCredit = 0;
                        
                        for (const row of jeRange.values.slice(1)) {
                            totalDebit += Number(row[debitIdx]) || 0;
                            totalCredit += Number(row[creditIdx]) || 0;
                        }
                        
                        result.journalEntry = {
                            totalDebit,
                            totalCredit,
                            difference: Math.abs(totalDebit - totalCredit),
                            isBalanced: Math.abs(totalDebit - totalCredit) < 0.01,
                            lineCount: jeRange.values.length - 1
                        };
                        
                        result.dataQuality.jeDraftReady = true;
                    }
                }
                
                console.log("CoPilot context gathered:", result);
                return result;
            });
        } catch (error) {
            console.warn("CoPilot context provider error:", error);
            return null;
        }
    };
}

function escapeHtml(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
}

/**
 * Render a button with a label underneath
 */
function renderLabeledButton(buttonHtml, label) {
    return `
        <div class="pf-labeled-button">
            ${buttonHtml}
            <span class="pf-button-label">${escapeHtml(label)}</span>
        </div>
    `;
}

function hasExcelRuntime() {
    return typeof Excel !== "undefined" && typeof Excel.run === "function";
}

function getStepNoteFields(stepId) {
    return STEP_NOTES_FIELDS[stepId] || null;
}

function headcountHasDifferences() {
    const rosterDiff = Math.abs(headcountState.roster?.difference ?? 0);
    const deptDiff = Math.abs(headcountState.departments?.difference ?? 0);
    return rosterDiff > 0 || deptDiff > 0;
}

function isHeadcountNotesRequired() {
    return !headcountState.skipAnalysis && headcountHasDifferences();
}

function formatMetricValue(value) {
    if (value === null || value === undefined) return "---";
    if (typeof value === "number" && Number.isInteger(value)) return value.toString();
    return value;
}

function formatSignedValue(value) {
    if (value === null || value === undefined) return "---";
    if (typeof value !== "number" || Number.isNaN(value)) return "---";
    if (value === 0) return "0";
    return value > 0 ? `+${value}` : value.toString();
}

function formatCurrency(value) {
    if (value === null || value === undefined || Number.isNaN(value)) return "---";
    if (typeof value !== "number") return value;
    return value.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function formatNumberDisplay(value) {
    if (value === null || value === undefined || value === "") return "---";
    const num = Number(value);
    const formatter = window.PrairieForge?.formatNumber;
    if (Number.isFinite(num)) {
        return formatter ? formatter(num) : num.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }
    return String(value);
}

function formatBankInput(value) {
    const numeric = parseBankAmount(value);
    if (!Number.isFinite(numeric)) return "";
    // Format as XXX,XXX.XX (no dollar sign per user preference)
    return numeric.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function buildCsv(rows = []) {
    return rows
        .map((row) =>
            (row || [])
                .map((cell) => {
                    if (cell == null) return "";
                    const str = String(cell);
                    if (/[",\n]/.test(str)) return `"${str.replace(/"/g, '""')}"`;
                    return str;
                })
                .join(",")
        )
        .join("\n");
}

function downloadCsv(filename, content) {
    const blob = new Blob([content], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function parseBankAmount(value) {
    if (typeof value === "number") return value;
    if (value == null) return NaN;
    const cleaned = String(value).replace(/[^0-9.-]/g, "");
    const parsed = Number.parseFloat(cleaned);
    return Number.isFinite(parsed) ? parsed : NaN;
}

function normalizePeriodKey(value) {
    if (value instanceof Date) return formatDateFromDate(value);
    if (typeof value === "number" && !Number.isNaN(value)) {
        const date = convertExcelDate(value);
        return date ? formatDateFromDate(date) : "";
    }
    const str = String(value ?? "").trim();
    if (!str) return "";
    if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;
    const parsed = new Date(str);
    if (!Number.isNaN(parsed.getTime())) {
        return formatDateFromDate(parsed);
    }
    return str;
}

function convertExcelDate(serial) {
    if (!Number.isFinite(serial)) return null;
    const utcDays = Math.floor(serial - 25569);
    if (!Number.isFinite(utcDays)) return null;
    const utcValue = utcDays * 86400 * 1000;
    const date = new Date(utcValue);
    // Mark this date as UTC-derived so formatDateFromDate can use UTC methods
    date._isUTC = true;
    return date;
}

function formatFriendlyPeriod(key) {
    if (!key) return "";
    const parsed = new Date(key);
    if (!Number.isNaN(parsed.getTime())) {
        return parsed.toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" });
    }
    return key;
}

function toNumber(value) {
    if (value == null || value === "") return 0;
    const num = Number(value);
    return Number.isFinite(num) ? num : 0;
}

function classifyExpenseComponent(label) {
    const normalized = normalizeString(label).toLowerCase();
    if (!normalized) return "variable";
    if (
        normalized.includes("burden") ||
        normalized.includes("tax") ||
        normalized.includes("benefit") ||
        normalized.includes("fica") ||
        normalized.includes("insurance") ||
        normalized.includes("health") ||
        normalized.includes("medicare")
    ) {
        return "burden";
    }
    if (
        normalized.includes("bonus") ||
        normalized.includes("commission") ||
        normalized.includes("variable") ||
        normalized.includes("overtime") ||
        normalized.includes("per diem")
    ) {
        return "variable";
    }
    return "fixed";
}

function parseExpenseRows(values) {
    if (!values || values.length < 2) return [];
    const headers = (values[0] || []).map((header) => normalizeHeader(header));
    console.log("parseExpenseRows - headers:", headers);
    const indexes = {
        payrollDate: headers.findIndex((header) => header.includes("payroll") && header.includes("date")),
        employee: headers.findIndex((header) => header.includes("employee")),
        department: headers.findIndex((header) => header.includes("department")),
        fixed: headers.findIndex((header) => header.includes("fixed")),
        variable: headers.findIndex((header) => header.includes("variable")),
        burden: headers.findIndex((header) => header.includes("burden")),
        amount: headers.findIndex((header) => header.includes("amount")),
        expenseReview: headers.findIndex((header) => header.includes("expense") && header.includes("review")),
        category: headers.findIndex((header) => header.includes("payroll") && header.includes("category"))
    };
    console.log("parseExpenseRows - column indexes:", indexes);
    
    // Log unique payroll dates found
    if (indexes.payrollDate >= 0) {
        const uniqueDates = new Set();
        for (let i = 1; i < values.length; i++) {
            const dateVal = values[i][indexes.payrollDate];
            if (dateVal) uniqueDates.add(String(dateVal));
        }
        console.log("parseExpenseRows - unique payroll dates found:", [...uniqueDates].slice(0, 20));
    }
    
    const rows = [];
    for (let i = 1; i < values.length; i += 1) {
        const row = values[i];
        const period = normalizePeriodKey(indexes.payrollDate >= 0 ? row[indexes.payrollDate] : null);
        if (!period) continue;
        const employee = indexes.employee >= 0 ? normalizeString(row[indexes.employee]) : "";
        const department = indexes.department >= 0 ? normalizeString(row[indexes.department]) || "Unassigned" : "Unassigned";
        const fixedVal = indexes.fixed >= 0 ? toNumber(row[indexes.fixed]) : null;
        const variableVal = indexes.variable >= 0 ? toNumber(row[indexes.variable]) : null;
        const burdenVal = indexes.burden >= 0 ? toNumber(row[indexes.burden]) : null;
        let fixed = 0;
        let variable = 0;
        let burden = 0;
        if (fixedVal !== null || variableVal !== null || burdenVal !== null) {
            fixed = fixedVal ?? 0;
            variable = variableVal ?? 0;
            burden = burdenVal ?? 0;
        } else {
            const amount = indexes.amount >= 0 ? toNumber(row[indexes.amount]) : 0;
            const classification = classifyExpenseComponent(
                indexes.expenseReview >= 0 ? row[indexes.expenseReview] : row[indexes.category]
            );
            if (classification === "fixed") fixed = amount;
            else if (classification === "burden") burden = amount;
            else variable = amount;
        }
        if (fixed === 0 && variable === 0 && burden === 0) continue;
        rows.push({
            period,
            employee,
            department: department || "Unassigned",
            fixed,
            variable,
            burden
        });
    }
    return rows;
}

function aggregateExpensePeriods(rows) {
    const map = new Map();
    rows.forEach((row) => {
        const key = row.period;
        if (!key) return;
        if (!map.has(key)) {
            map.set(key, {
                key,
                label: formatFriendlyPeriod(key),
                employees: new Set(),
                departments: new Map(),
                summary: { fixed: 0, variable: 0, burden: 0 }
            });
        }
        const bucket = map.get(key);
        bucket.employees.add(row.employee || `EMP-${bucket.employees.size + 1}`);
        const deptKey = row.department || "Unassigned";
        if (!bucket.departments.has(deptKey)) {
            bucket.departments.set(deptKey, {
                name: deptKey,
                fixed: 0,
                variable: 0,
                burden: 0,
                employees: new Set()
            });
        }
        const dept = bucket.departments.get(deptKey);
        dept.fixed += row.fixed;
        dept.variable += row.variable;
        dept.burden += row.burden;
        dept.employees.add(row.employee || `EMP-${dept.employees.size + 1}`);
        bucket.summary.fixed += row.fixed;
        bucket.summary.variable += row.variable;
        bucket.summary.burden += row.burden;
    });
    const result = [];
    map.forEach((bucket) => {
        const total = bucket.summary.fixed + bucket.summary.variable + bucket.summary.burden;
        const departments = Array.from(bucket.departments.values()).map((dept) => {
            const gross = dept.fixed + dept.variable;
            const allIn = gross + dept.burden;
            return {
                name: dept.name,
                fixed: dept.fixed,
                variable: dept.variable,
                gross,
                burden: dept.burden,
                allIn,
                percent: total ? allIn / total : 0,
                headcount: dept.employees.size,
                delta: 0
            };
        });
        departments.sort((a, b) => b.allIn - a.allIn);
        const summary = {
            employeeCount: bucket.employees.size,
            fixed: bucket.summary.fixed,
            variable: bucket.summary.variable,
            burden: bucket.summary.burden,
            total,
            burdenRate: total ? bucket.summary.burden / total : 0,
            delta: 0
        };
        const totalsRow = {
            name: "Totals",
            fixed: bucket.summary.fixed,
            variable: bucket.summary.variable,
            gross: bucket.summary.fixed + bucket.summary.variable,
            burden: bucket.summary.burden,
            allIn: total,
            percent: total ? 1 : 0,
            headcount: bucket.employees.size,
            delta: 0,
            isTotal: true
        };
        result.push({
            key: bucket.key,
            label: bucket.label,
            summary,
            departments,
            totalsRow
        });
    });
    return result.sort((a, b) => (a.key < b.key ? 1 : -1));
}

function buildExpenseReviewPeriods(cleanValues, archiveValues) {
    console.log("buildExpenseReviewPeriods - cleanValues rows:", cleanValues?.length || 0);
    console.log("buildExpenseReviewPeriods - archiveValues rows:", archiveValues?.length || 0);
    
    const currentPeriods = aggregateExpensePeriods(parseExpenseRows(cleanValues));
    const archivePeriods = aggregateExpensePeriods(parseExpenseRows(archiveValues));
    
    console.log("buildExpenseReviewPeriods - currentPeriods:", currentPeriods.map(p => ({ key: p.key, employees: p.summary?.employeeCount, total: p.summary?.total })));
    console.log("buildExpenseReviewPeriods - archivePeriods:", archivePeriods.map(p => ({ key: p.key, employees: p.summary?.employeeCount, total: p.summary?.total })));
    
    const archiveMap = new Map(archivePeriods.map((period) => [period.key, period]));
    const combined = [];
    if (currentPeriods.length) {
        combined.push(currentPeriods[0]);
        archiveMap.delete(currentPeriods[0].key);
    }
    archivePeriods.forEach((period) => {
        if (combined.length >= 6) return;
        if (!combined.some((existing) => existing.key === period.key)) {
            combined.push(period);
        }
    });
    
    console.log("buildExpenseReviewPeriods - combined before filter:", combined.map(p => ({ key: p.key, employees: p.summary?.employeeCount, total: p.summary?.total })));
    
    // Filter to only include periods that look like real pay periods:
    // - Must have at least 3 employees (filters out test data or partial entries)
    // - Must have meaningful total (> $1000) to filter out adjustment entries
    const minEmployeesForPayPeriod = 3;
    const minTotalForPayPeriod = 1000;
    
    const sorted = combined
        .filter((period) => {
            if (!period || !period.key) {
                console.log("buildExpenseReviewPeriods - EXCLUDED (no key):", period);
                return false;
            }
            const total = period.summary?.total || 
                ((period.summary?.fixed || 0) + (period.summary?.variable || 0) + (period.summary?.burden || 0));
            const employeeCount = period.summary?.employeeCount || 0;
            // Always include the current period (first one), filter others
            if (combined.indexOf(period) === 0) {
                console.log(`buildExpenseReviewPeriods - INCLUDED (current): ${period.key} - ${employeeCount} employees, $${total}`);
                return true;
            }
            const included = employeeCount >= minEmployeesForPayPeriod && total >= minTotalForPayPeriod;
            console.log(`buildExpenseReviewPeriods - ${included ? "INCLUDED" : "EXCLUDED"}: ${period.key} - ${employeeCount} employees, $${total} (needs >=${minEmployeesForPayPeriod} emp, >=$${minTotalForPayPeriod})`);
            return included;
        })
        .sort((a, b) => (a.key < b.key ? 1 : -1))
        .slice(0, 6);
    
    console.log("buildExpenseReviewPeriods - FINAL periods:", sorted.map(p => p.key));
    sorted.forEach((period, index) => {
        const previous = sorted[index + 1];
        const delta = previous ? period.summary.total - previous.summary.total : 0;
        period.summary.delta = delta;
        const previousDeptMap = new Map((previous?.departments || []).map((dept) => [dept.name, dept]));
        period.departments.forEach((dept) => {
            const prev = previousDeptMap.get(dept.name);
            dept.delta = prev ? dept.allIn - prev.allIn : 0;
        });
        period.totalsRow.delta = delta;
    });
    return sorted;
}

async function prepareExpenseReviewData() {
    if (!hasExcelRuntime()) {
        updateExpenseReviewState({
            loading: false,
            lastError: "Excel runtime is unavailable."
        });
        return;
    }
    updateExpenseReviewState({ loading: true, lastError: null });
    try {
        const periods = await Excel.run(async (context) => {
            // Check if required sheets exist
            const cleanSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.DATA_CLEAN);
            const archiveSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.ARCHIVE_SUMMARY);
            const reviewSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.EXPENSE_REVIEW);
            
            cleanSheet.load("isNullObject, name");
            archiveSheet.load("isNullObject, name");
            reviewSheet.load("isNullObject, name");
            await context.sync();
            
            console.log("Expense Review - Sheet check:", {
                cleanSheet: cleanSheet.isNullObject ? "MISSING" : cleanSheet.name,
                archiveSheet: archiveSheet.isNullObject ? "MISSING" : archiveSheet.name,
                reviewSheet: reviewSheet.isNullObject ? "MISSING" : reviewSheet.name
            });
            
            // Create missing sheets if needed
            if (reviewSheet.isNullObject) {
                console.log("Creating PR_Expense_Review sheet...");
                const newReviewSheet = context.workbook.worksheets.add(SHEET_NAMES.EXPENSE_REVIEW);
                await context.sync();
                // Re-get the sheet
                const createdSheet = context.workbook.worksheets.getItem(SHEET_NAMES.EXPENSE_REVIEW);
                
                // Get data even if cleanSheet is missing (will be empty)
                let cleanValues = [];
                let archiveValues = [];
                
                if (!cleanSheet.isNullObject) {
            const cleanRange = cleanSheet.getUsedRangeOrNullObject();
                    cleanRange.load("values");
                    await context.sync();
                    cleanValues = cleanRange.isNullObject ? [] : cleanRange.values || [];
                }
                
                if (!archiveSheet.isNullObject) {
            const archiveRange = archiveSheet.getUsedRangeOrNullObject();
                    archiveRange.load("values");
                    await context.sync();
                    archiveValues = archiveRange.isNullObject ? [] : archiveRange.values || [];
                }
                
                const periodData = buildExpenseReviewPeriods(cleanValues, archiveValues);
                await writeExpenseReviewSheet(context, createdSheet, periodData);
                return periodData;
            }
            
            // Get data from existing sheets
            let cleanValues = [];
            let archiveValues = [];
            
            if (!cleanSheet.isNullObject) {
                const cleanRange = cleanSheet.getUsedRangeOrNullObject();
            cleanRange.load("values");
                await context.sync();
                cleanValues = cleanRange.isNullObject ? [] : cleanRange.values || [];
                console.log("Expense Review - PR_Data_Clean rows:", cleanValues.length);
            } else {
                console.warn("Expense Review - PR_Data_Clean sheet not found, using empty data");
            }
            
            if (!archiveSheet.isNullObject) {
                const archiveRange = archiveSheet.getUsedRangeOrNullObject();
            archiveRange.load("values");
            await context.sync();
                archiveValues = archiveRange.isNullObject ? [] : archiveRange.values || [];
                console.log("Expense Review - PR_Archive_Summary rows:", archiveValues.length);
            } else {
                console.warn("Expense Review - PR_Archive_Summary sheet not found, using empty data");
            }
            
            const periodData = buildExpenseReviewPeriods(cleanValues, archiveValues);
            console.log("Expense Review - Periods built:", periodData.length);
            
            await writeExpenseReviewSheet(context, reviewSheet, periodData);
            return periodData;
        });
        updateExpenseReviewState({ loading: false, periods, lastError: null });
        
        // Run completeness check after data is loaded
        await runPayrollCompletenessCheck();
        renderApp(); // Re-render to show completeness check results
    } catch (error) {
        console.error("Expense Review: unable to build executive summary", error);
        console.error("Error details:", error.message, error.stack);
        updateExpenseReviewState({
            loading: false,
            lastError: `Unable to build the Expense Review: ${error.message || "Unknown error"}`,
            periods: []
        });
    }
}

async function writeExpenseReviewSheet(context, sheet, periods) {
    if (!sheet) {
        console.error("writeExpenseReviewSheet: sheet is null/undefined");
        return;
    }
    
    console.log("writeExpenseReviewSheet: Building executive dashboard with", periods.length, "periods");
    
    // Clear existing content and charts
    try {
        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load("address");
        const charts = sheet.charts;
        charts.load("items");
        await context.sync();
        if (!usedRange.isNullObject) {
            usedRange.clear();
            await context.sync();
        }
        // Remove existing charts
        for (let i = charts.items.length - 1; i >= 0; i--) {
            charts.items[i].delete();
        }
        await context.sync();
    } catch (e) {
        console.warn("Could not clear sheet:", e);
    }
    
    // ═══════════════════════════════════════════════════════════════════
    // PREPARE DATA
    // ═══════════════════════════════════════════════════════════════════
    const current = periods[0] || {};
    const prior = periods[1] || {};
    const currentSummary = current.summary || {};
    const priorSummary = prior.summary || {};
    
    // Get period from config
    const configPeriod = getConfigValue("PR_Accounting_Period") || getPayrollDateValue() || "";
    
    // Key metrics
    const totalPayroll = Number(currentSummary.total) || 0;
    const priorTotal = Number(priorSummary.total) || 0;
    const periodChange = totalPayroll - priorTotal;
    const periodChangePct = priorTotal ? periodChange / priorTotal : 0;
    const employeeCount = Number(currentSummary.employeeCount) || 0;
    const priorEmployeeCount = Number(priorSummary.employeeCount) || 0;
    const headcountChange = employeeCount - priorEmployeeCount;
    const avgPerEmployee = employeeCount ? totalPayroll / employeeCount : 0;
    const priorAvgPerEmployee = priorEmployeeCount ? priorTotal / priorEmployeeCount : 0;
    const avgChange = avgPerEmployee - priorAvgPerEmployee;
    
    // Variable comp analysis - detect if this is a "variable pay" period
    // Look for commission, bonus patterns in period data
    const hasVariableComp = detectVariableCompPeriod(periods);
    const basePayOnly = detectBasePayOnlyPeriod(current, periods);
    
    const periodLabel = current.label || current.key || "Current Period";
    const generatedTimestamp = new Date().toLocaleString("en-US", { 
        month: "short", day: "numeric", year: "numeric", hour: "numeric", minute: "2-digit"
    });
    
    // Trend helpers
    const trendArrow = (val) => val > 0 ? "▲" : val < 0 ? "▼" : "—";
    
    // ═══════════════════════════════════════════════════════════════════
    // CALCULATE HISTORICAL RANGES FOR SPECTRUM VISUALIZATION
    // ═══════════════════════════════════════════════════════════════════
    const historicalTotals = periods.map(p => p.summary?.total || 0).filter(t => t > 0);
    const historicalAvgPerEmp = periods.map(p => {
        const s = p.summary || {};
        const emp = s.employeeCount || 0;
        return emp > 0 ? (s.total || 0) / emp : 0;
    }).filter(a => a > 0);
    const historicalChangePcts = periods.slice(0, -1).map((p, i) => {
        const curr = p.summary?.total || 0;
        const prev = periods[i + 1]?.summary?.total || 0;
        return prev > 0 ? (curr - prev) / prev : 0;
    });
    
    // Calculate ranges - includes current value to ensure the spectrum adjusts if current is outside historical range
    const calcRange = (arr, currentValue = null) => {
        // Include current value in range if provided (allows range to expand beyond historical)
        const values = currentValue !== null ? [...arr, currentValue] : arr;
        if (!values.length) return { min: 0, max: 0, avg: 0 };
        const min = Math.min(...values);
        const max = Math.max(...values);
        // Average is calculated from just the historical values (or all if no current provided)
        const avgBase = arr.length ? arr : values;
        const avg = avgBase.reduce((a, b) => a + b, 0) / avgBase.length;
        return { min, max, avg };
    };
    
    const payrollRange = calcRange(historicalTotals, totalPayroll);
    const avgEmpRange = calcRange(historicalAvgPerEmp, avgPerEmployee);
    const changePctRange = calcRange(historicalChangePcts);
    
    // Spectrum builder - creates a visual bar showing where current value falls
    // Uses Unicode block characters: ░ (light), ▒ (medium), ● (current position)
    const buildSpectrum = (current, min, max, width = 20) => {
        if (max <= min) return "░".repeat(width);
        const range = max - min;
        const position = Math.max(0, Math.min(1, (current - min) / range));
        const markerPos = Math.round(position * (width - 1));
        
        let bar = "";
        for (let i = 0; i < width; i++) {
            if (i === markerPos) {
                bar += "●";
            } else {
                bar += "░";
            }
        }
        return bar;
    };
    
    // ═══════════════════════════════════════════════════════════════════
    // EXTRACT FIXED/VARIABLE/BURDEN BREAKDOWNS
    // ═══════════════════════════════════════════════════════════════════
    const currentFixed = Number(currentSummary.fixed) || 0;
    const currentVariable = Number(currentSummary.variable) || 0;
    const currentBurden = Number(currentSummary.burden) || 0;
    const currentGross = currentFixed + currentVariable;
    const currentBurdenRate = totalPayroll ? currentBurden / totalPayroll : 0;
    
    const priorFixed = Number(priorSummary.fixed) || 0;
    const priorVariable = Number(priorSummary.variable) || 0;
    const priorBurden = Number(priorSummary.burden) || 0;
    const priorBurdenRate = priorTotal ? priorBurden / priorTotal : 0;
    
    // Calculate Sales & Marketing variable vs Other departments variable
    const departments = current.departments || [];
    const salesMarketingDepts = departments.filter(d => {
        const name = (d.name || "").toLowerCase();
        return name.includes("sales") || name.includes("marketing");
    });
    const otherDepts = departments.filter(d => {
        const name = (d.name || "").toLowerCase();
        return !name.includes("sales") && !name.includes("marketing");
    });
    
    const salesMarketingVariable = salesMarketingDepts.reduce((sum, d) => sum + (d.variable || 0), 0);
    const salesMarketingHeadcount = salesMarketingDepts.reduce((sum, d) => sum + (d.headcount || 0), 0);
    const otherVariable = otherDepts.reduce((sum, d) => sum + (d.variable || 0), 0);
    const otherHeadcount = otherDepts.reduce((sum, d) => sum + (d.headcount || 0), 0);
    
    const avgVariableSalesMarketing = salesMarketingHeadcount ? salesMarketingVariable / salesMarketingHeadcount : 0;
    const avgVariableOther = otherHeadcount ? otherVariable / otherHeadcount : 0;
    const avgFixedPerEmployee = employeeCount ? currentFixed / employeeCount : 0;
    
    // ═══════════════════════════════════════════════════════════════════
    // BUILD CLEAN DATA ARRAY - NEW LAYOUT
    // ═══════════════════════════════════════════════════════════════════
    const data = [];
    let rowIdx = 0;
    const rowMap = {};
    
    // ─── HEADER (Left justified) ───
    rowMap.headerStart = rowIdx;
    // Format period as date if it's a number (Excel serial date)
    let formattedPeriod = configPeriod || periodLabel;
    if (typeof configPeriod === 'number' || (!isNaN(Number(configPeriod)) && configPeriod)) {
        const dateNum = Number(configPeriod);
        if (dateNum > 40000 && dateNum < 60000) { // Valid Excel date range
            const excelEpoch = new Date(1899, 11, 30);
            const date = new Date(excelEpoch.getTime() + dateNum * 24 * 60 * 60 * 1000);
            formattedPeriod = date.toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
        }
    }
    
    data.push(["PAYROLL EXPENSE REVIEW"]); rowIdx++;
    data.push([`Period: ${formattedPeriod}`]); rowIdx++;
    data.push([`Generated: ${generatedTimestamp}`]); rowIdx++;
    data.push([""]); rowIdx++;
    rowMap.headerEnd = rowIdx - 1;
    
    // ─── EXECUTIVE SUMMARY (Frozen Section) ───
    rowMap.execSummaryStart = rowIdx;
    data.push(["EXECUTIVE SUMMARY"]); rowIdx++;
    rowMap.execSummaryHeader = rowIdx - 1;
    data.push([""]); rowIdx++;
    
    // Column headers: Pay Date, Headcount, Fixed Salary, Variable Salary, Burden, Total Payroll, Burden Rate
    data.push(["", "Pay Date", "Headcount", "Fixed Salary", "Variable Salary", "Burden", "Total Payroll", "Burden Rate"]); rowIdx++;
    rowMap.execSummaryColHeaders = rowIdx - 1;
    
    // Current Pay Period row
    data.push(["Current Pay Period", current.label || current.key || "", employeeCount, currentFixed, currentVariable, currentBurden, totalPayroll, currentBurdenRate]); rowIdx++;
    rowMap.execSummaryCurrentRow = rowIdx - 1;
    
    // Same Period Prior Month row  
    data.push(["Same Period Prior Month", prior.label || prior.key || "", priorEmployeeCount, priorFixed, priorVariable, priorBurden, priorTotal, priorBurdenRate]); rowIdx++;
    rowMap.execSummaryPriorRow = rowIdx - 1;
    
    data.push([""]); rowIdx++;
    data.push([""]); rowIdx++;
    rowMap.execSummaryEnd = rowIdx - 1;
    
    // ─── CURRENT PERIOD BREAKDOWN (DEPARTMENT) ───
    rowMap.deptBreakdownStart = rowIdx;
    data.push(["CURRENT PERIOD BREAKDOWN (DEPARTMENT)"]); rowIdx++;
    rowMap.deptBreakdownHeader = rowIdx - 1;
    data.push([""]); rowIdx++;
    data.push([`Payroll Date`, current.label || current.key || ""]); rowIdx++;
    data.push([""]); rowIdx++;
    
    // Column headers: Department, Fixed Salary, Variable Salary, Gross Pay, Burden, All-In Total, % of Total, Headcount
    data.push(["Department", "Fixed Salary", "Variable Salary", "Gross Pay", "Burden", "All-In Total", "% of Total", "Headcount"]); rowIdx++;
    rowMap.deptColHeaders = rowIdx - 1;
    
    // Department data rows
    const sortedDepts = [...departments].sort((a, b) => (b.allIn || 0) - (a.allIn || 0));
    rowMap.deptDataStart = rowIdx;
    sortedDepts.forEach((dept) => {
        data.push([
            dept.name || "",
            dept.fixed || 0,
            dept.variable || 0,
            dept.gross || 0,
            dept.burden || 0,
            dept.allIn || 0,
            dept.percent || 0,
            dept.headcount || 0
        ]); rowIdx++;
    });
    rowMap.deptDataEnd = rowIdx - 1;
    
    // TOTAL row
    if (current.totalsRow) {
        const totals = current.totalsRow;
        data.push([
            "TOTAL",
            totals.fixed || 0,
            totals.variable || 0,
            totals.gross || 0,
            totals.burden || 0,
            totals.allIn || 0,
            1,
            totals.headcount || 0
        ]); rowIdx++;
        rowMap.deptTotalsRow = rowIdx - 1;
    }
    
    data.push([""]); rowIdx++;
    data.push([""]); rowIdx++;
    rowMap.deptBreakdownEnd = rowIdx - 1;
    
    // ─── HISTORICAL CONTEXT ───
    rowMap.historicalStart = rowIdx;
    data.push(["HISTORICAL CONTEXT"]); rowIdx++;
    rowMap.historicalHeader = rowIdx - 1;
    data.push([`Visual comparison of current period vs. historical range (${periods.length} periods). The dot (●) shows where you currently stand.`]); rowIdx++;
    data.push([""]); rowIdx++;
    
    // Format helpers for spectrum labels  
    const fmtK = (n) => `$${Math.round(n / 1000)}K`;
    const fmtPct = (n) => `${(n * 100).toFixed(1)}%`;
    
    // Column headers for historical context
    data.push(["", "Metric", "Low", "Range", "High", "", "Current", "Average"]); rowIdx++;
    rowMap.historicalColHeaders = rowIdx - 1;
    
    // Calculate additional historical ranges - include current values to allow range expansion
    // Don't filter zeros here - let calcRange handle it with current value inclusion
    const historicalFixed = periods.map(p => p.summary?.fixed || 0).filter(t => t > 0);
    const historicalVariable = periods.map(p => p.summary?.variable || 0); // Keep zeros for proper range
    const historicalBurdenRates = periods.map(p => {
        const s = p.summary || {};
        return s.total ? (s.burden || 0) / s.total : 0;
    }); // Keep zeros for proper range
    const historicalAvgFixed = periods.map(p => {
        const s = p.summary || {};
        const emp = s.employeeCount || 0;
        return emp > 0 ? (s.fixed || 0) / emp : 0;
    }).filter(a => a > 0);
    
    const fixedRange = calcRange(historicalFixed, currentFixed);
    const variableRange = calcRange(historicalVariable, currentVariable);
    const burdenRateRange = calcRange(historicalBurdenRates, currentBurdenRate);
    const avgFixedRange = calcRange(historicalAvgFixed, avgFixedPerEmployee);
    
    // Build spectrum rows
    rowMap.spectrumRows = [];
    
    // Total Payroll
    const payrollSpectrum = buildSpectrum(totalPayroll, payrollRange.min, payrollRange.max, 25);
    data.push(["", "Total Payroll", fmtK(payrollRange.min), payrollSpectrum, fmtK(payrollRange.max), "", fmtK(totalPayroll), fmtK(payrollRange.avg)]); rowIdx++;
    rowMap.spectrumRows.push(rowIdx - 1);
    
    // Total Fixed Payroll
    const fixedSpectrum = buildSpectrum(currentFixed, fixedRange.min, fixedRange.max, 25);
    data.push(["", "Total Fixed Payroll", fmtK(fixedRange.min), fixedSpectrum, fmtK(fixedRange.max), "", fmtK(currentFixed), fmtK(fixedRange.avg)]); rowIdx++;
    rowMap.spectrumRows.push(rowIdx - 1);
    
    // Total Variable Payroll
    const variableSpectrum = buildSpectrum(currentVariable, variableRange.min, variableRange.max, 25);
    data.push(["", "Total Variable Payroll", fmtK(variableRange.min), variableSpectrum, fmtK(variableRange.max), "", fmtK(currentVariable), fmtK(variableRange.avg)]); rowIdx++;
    rowMap.spectrumRows.push(rowIdx - 1);
    
    data.push([""]); rowIdx++;
    
    // Avg Fixed Payroll per Employee
    const avgFixedSpectrum = buildSpectrum(avgFixedPerEmployee, avgFixedRange.min, avgFixedRange.max, 25);
    data.push(["", "Avg Fixed Payroll / Employee", fmtK(avgFixedRange.min), avgFixedSpectrum, fmtK(avgFixedRange.max), "", fmtK(avgFixedPerEmployee), fmtK(avgFixedRange.avg)]); rowIdx++;
    rowMap.spectrumRows.push(rowIdx - 1);
    
    // Calculate historical ranges for Sales & Marketing variable
    const historicalAvgVarSM = periods.map(p => {
        const depts = p.departments || [];
        const smDepts = depts.filter(d => {
            const name = (d.name || "").toLowerCase();
            return name.includes("sales") || name.includes("marketing");
        });
        const smVar = smDepts.reduce((sum, d) => sum + (d.variable || 0), 0);
        const smHc = smDepts.reduce((sum, d) => sum + (d.headcount || 0), 0);
        return smHc > 0 ? smVar / smHc : 0;
    }); // Keep zeros for proper range - current value inclusion will handle expansion
    const avgVarSMRange = calcRange(historicalAvgVarSM, avgVariableSalesMarketing);
    
    // Calculate historical ranges for Other departments variable
    const historicalAvgVarOther = periods.map(p => {
        const depts = p.departments || [];
        const otherD = depts.filter(d => {
            const name = (d.name || "").toLowerCase();
            return !name.includes("sales") && !name.includes("marketing");
        });
        const otherV = otherD.reduce((sum, d) => sum + (d.variable || 0), 0);
        const otherH = otherD.reduce((sum, d) => sum + (d.headcount || 0), 0);
        return otherH > 0 ? otherV / otherH : 0;
    }); // Keep zeros for proper range
    const avgVarOtherRange = calcRange(historicalAvgVarOther, avgVariableOther);
    
    // Avg Variable Payroll per Sales & Marketing (with spectrum visualization)
    if (salesMarketingHeadcount > 0) {
        const avgVarSMSpectrum = buildSpectrum(avgVariableSalesMarketing, avgVarSMRange.min, avgVarSMRange.max, 25);
        data.push(["", "Avg Variable / Sales & Marketing", fmtK(avgVarSMRange.min), avgVarSMSpectrum, fmtK(avgVarSMRange.max), "", fmtK(avgVariableSalesMarketing), `${salesMarketingHeadcount} emp`]); rowIdx++;
        rowMap.spectrumRows.push(rowIdx - 1);
    }
    
    // Avg Variable Payroll per Other Departments (with spectrum visualization)
    if (otherHeadcount > 0) {
        const avgVarOtherSpectrum = buildSpectrum(avgVariableOther, avgVarOtherRange.min, avgVarOtherRange.max, 25);
        data.push(["", "Avg Variable / Other Depts", fmtK(avgVarOtherRange.min), avgVarOtherSpectrum, fmtK(avgVarOtherRange.max), "", fmtK(avgVariableOther), `${otherHeadcount} emp`]); rowIdx++;
        rowMap.spectrumRows.push(rowIdx - 1);
    }
    
    data.push([""]); rowIdx++;
    
    // Burden Rate (%)
    const burdenRateSpectrum = buildSpectrum(currentBurdenRate, burdenRateRange.min, burdenRateRange.max, 25);
    data.push(["", "Burden Rate (%)", fmtPct(burdenRateRange.min), burdenRateSpectrum, fmtPct(burdenRateRange.max), "", fmtPct(currentBurdenRate), fmtPct(burdenRateRange.avg)]); rowIdx++;
    rowMap.spectrumRows.push(rowIdx - 1);
    
    data.push([""]); rowIdx++;
    data.push([""]); rowIdx++;
    rowMap.historicalEnd = rowIdx - 1;
    
    // ─── PERIOD TRENDS ───
    rowMap.trendsStart = rowIdx;
    data.push(["PERIOD TRENDS"]); rowIdx++;
    rowMap.trendsHeader = rowIdx - 1;
    data.push([""]); rowIdx++;
    
    // Trend data table (will be used for chart)
    data.push(["Pay Period", "Total Payroll", "Fixed Payroll", "Variable Payroll", "Burden", "Headcount"]); rowIdx++;
    rowMap.trendColHeaders = rowIdx - 1;
    
    // Up to 6 periods in reverse chronological order (oldest first for chart)
    const trendPeriods = periods.slice(0, 6).reverse();
    rowMap.trendDataStart = rowIdx;
    trendPeriods.forEach((period) => {
        const s = period.summary || {};
        data.push([
            period.label || period.key || "",
            s.total || 0,
            s.fixed || 0,
            s.variable || 0,
            s.burden || 0,
            s.employeeCount || 0
        ]); rowIdx++;
    });
    rowMap.trendDataEnd = rowIdx - 1;
    
    data.push([""]); rowIdx++;
    rowMap.trendsEnd = rowIdx - 1;
    
    // Reserve space for charts (payroll chart + headcount chart)
    rowMap.chartStart = rowIdx;
    for (let i = 0; i < 15; i++) {
        data.push([""]); rowIdx++;
    }
    rowMap.payrollChartEnd = rowIdx - 1;
    
    // Space for headcount chart
    rowMap.headcountChartStart = rowIdx;
    for (let i = 0; i < 12; i++) {
        data.push([""]); rowIdx++;
    }
    rowMap.headcountChartEnd = rowIdx - 1;
    
    // ═══════════════════════════════════════════════════════════════════
    // WRITE DATA (10 columns to accommodate all data)
    // ═══════════════════════════════════════════════════════════════════
    console.log("writeExpenseReviewSheet: Writing", data.length, "rows");
    
    // Normalize all rows to 10 columns
    const normalizedData = data.map(row => {
        const r = Array.isArray(row) ? row : [""];
        while (r.length < 10) r.push("");
        return r.slice(0, 10);
    });
    
    try {
        const dataRange = sheet.getRangeByIndexes(0, 0, normalizedData.length, 10);
        dataRange.values = normalizedData;
        await context.sync();
    } catch (writeError) {
        console.error("writeExpenseReviewSheet: Write failed", writeError);
        throw writeError;
    }
    
    // ═══════════════════════════════════════════════════════════════════
    // APPLY FORMATTING
    // ═══════════════════════════════════════════════════════════════════
    try {
        // Column widths
        sheet.getRange("A:A").format.columnWidth = 200;   // Section headers / Department names
        sheet.getRange("B:B").format.columnWidth = 130;   // Fixed Salary / Metric names
        sheet.getRange("C:C").format.columnWidth = 100;   // Variable Salary / Low
        sheet.getRange("D:D").format.columnWidth = 200;   // Gross Pay / Spectrum (wider for dots)
        sheet.getRange("E:E").format.columnWidth = 100;   // Burden / High
        sheet.getRange("F:F").format.columnWidth = 100;   // All-In Total
        sheet.getRange("G:G").format.columnWidth = 100;   // % of Total / Current
        sheet.getRange("H:H").format.columnWidth = 100;   // Headcount / Average
        sheet.getRange("I:I").format.columnWidth = 80;
        sheet.getRange("J:J").format.columnWidth = 80;
        await context.sync();
        
        // ─── HEADER SECTION (Left justified) ───
        const titleCell = sheet.getRange("A1");
        titleCell.format.font.bold = true;
        titleCell.format.font.size = 22;
        titleCell.format.font.color = "#1e293b";
        
        sheet.getRange("A2").format.font.size = 11;
        sheet.getRange("A2").format.font.color = "#64748b";
        sheet.getRange("A3").format.font.size = 10;
        sheet.getRange("A3").format.font.color = "#94a3b8";
        await context.sync();
        
        // ─── EXECUTIVE SUMMARY SECTION ───
        const execHeader = sheet.getRange(`A${rowMap.execSummaryHeader + 1}`);
        execHeader.format.font.bold = true;
        execHeader.format.font.size = 14;
        execHeader.format.font.color = "#1e293b";
        
        // Column headers - dark background
        const execColHeaders = sheet.getRange(`A${rowMap.execSummaryColHeaders + 1}:H${rowMap.execSummaryColHeaders + 1}`);
        execColHeaders.format.font.bold = true;
        execColHeaders.format.font.size = 10;
        execColHeaders.format.fill.color = "#1e293b";
        execColHeaders.format.font.color = "#ffffff";
        
        // Current period row - light green background
        const currentRow = sheet.getRange(`A${rowMap.execSummaryCurrentRow + 1}:H${rowMap.execSummaryCurrentRow + 1}`);
        currentRow.format.fill.color = "#dcfce7";
        currentRow.format.font.bold = true;
        
        // Prior period row - light gray background
        const priorRow = sheet.getRange(`A${rowMap.execSummaryPriorRow + 1}:H${rowMap.execSummaryPriorRow + 1}`);
        priorRow.format.fill.color = "#f1f5f9";
        
        // Number formats for executive summary
        for (const rowNum of [rowMap.execSummaryCurrentRow + 1, rowMap.execSummaryPriorRow + 1]) {
            sheet.getRange(`C${rowNum}`).numberFormat = [["#,##0"]];       // Headcount
            sheet.getRange(`D${rowNum}`).numberFormat = [["$#,##0"]];      // Fixed
            sheet.getRange(`E${rowNum}`).numberFormat = [["$#,##0"]];      // Variable
            sheet.getRange(`F${rowNum}`).numberFormat = [["$#,##0"]];      // Burden
            sheet.getRange(`G${rowNum}`).numberFormat = [["$#,##0"]];      // Total
            sheet.getRange(`H${rowNum}`).numberFormat = [["0.00%"]];       // Burden Rate
        }
        await context.sync();
        
        // ─── DEPARTMENT BREAKDOWN SECTION ───
        const deptHeader = sheet.getRange(`A${rowMap.deptBreakdownHeader + 1}`);
        deptHeader.format.font.bold = true;
        deptHeader.format.font.size = 14;
        deptHeader.format.font.color = "#1e293b";
        
        // Column headers - dark background
        const deptColHeaders = sheet.getRange(`A${rowMap.deptColHeaders + 1}:H${rowMap.deptColHeaders + 1}`);
        deptColHeaders.format.font.bold = true;
        deptColHeaders.format.font.size = 10;
        deptColHeaders.format.fill.color = "#1e293b";
        deptColHeaders.format.font.color = "#ffffff";
        
        // Department data rows
        for (let i = rowMap.deptDataStart; i <= rowMap.deptDataEnd; i++) {
            const row = i + 1;
            sheet.getRange(`B${row}`).numberFormat = [["$#,##0"]];   // Fixed
            sheet.getRange(`C${row}`).numberFormat = [["$#,##0"]];   // Variable
            sheet.getRange(`D${row}`).numberFormat = [["$#,##0"]];   // Gross
            sheet.getRange(`E${row}`).numberFormat = [["$#,##0"]];   // Burden
            sheet.getRange(`F${row}`).numberFormat = [["$#,##0"]];   // All-In
            sheet.getRange(`G${row}`).numberFormat = [["0.00%"]];    // % of Total
            sheet.getRange(`H${row}`).numberFormat = [["#,##0"]];    // Headcount
            
            // Alternate row shading
            if ((i - rowMap.deptDataStart) % 2 === 1) {
                sheet.getRange(`A${row}:H${row}`).format.fill.color = "#f8fafc";
            }
        }
        
        // Totals row - dark background
        if (rowMap.deptTotalsRow) {
            const totalsRange = sheet.getRange(`A${rowMap.deptTotalsRow + 1}:H${rowMap.deptTotalsRow + 1}`);
            totalsRange.format.font.bold = true;
            totalsRange.format.fill.color = "#1e293b";
            totalsRange.format.font.color = "#ffffff";
            
            const row = rowMap.deptTotalsRow + 1;
            sheet.getRange(`B${row}`).numberFormat = [["$#,##0"]];
            sheet.getRange(`C${row}`).numberFormat = [["$#,##0"]];
            sheet.getRange(`D${row}`).numberFormat = [["$#,##0"]];
            sheet.getRange(`E${row}`).numberFormat = [["$#,##0"]];
            sheet.getRange(`F${row}`).numberFormat = [["$#,##0"]];
            sheet.getRange(`G${row}`).numberFormat = [["0%"]];
            sheet.getRange(`H${row}`).numberFormat = [["#,##0"]];
        }
        await context.sync();
        
        // ─── HISTORICAL CONTEXT SECTION ───
        const histHeader = sheet.getRange(`A${rowMap.historicalHeader + 1}`);
        histHeader.format.font.bold = true;
        histHeader.format.font.size = 14;
        histHeader.format.font.color = "#1e293b";
        
        // Description text
        sheet.getRange(`A${rowMap.historicalHeader + 2}`).format.font.size = 10;
        sheet.getRange(`A${rowMap.historicalHeader + 2}`).format.font.color = "#64748b";
        sheet.getRange(`A${rowMap.historicalHeader + 2}`).format.font.italic = true;
        
        // Column headers - center Low, High, Current, Average
        const histColHeaders = sheet.getRange(`A${rowMap.historicalColHeaders + 1}:H${rowMap.historicalColHeaders + 1}`);
        histColHeaders.format.font.bold = true;
        histColHeaders.format.font.size = 10;
        histColHeaders.format.fill.color = "#e2e8f0";
        histColHeaders.format.font.color = "#334155";
        // Center the column headers for Low (C), High (E), Current (G), Average (H)
        sheet.getRange(`C${rowMap.historicalColHeaders + 1}`).format.horizontalAlignment = "Center";
        sheet.getRange(`E${rowMap.historicalColHeaders + 1}`).format.horizontalAlignment = "Center";
        sheet.getRange(`G${rowMap.historicalColHeaders + 1}`).format.horizontalAlignment = "Center";
        sheet.getRange(`H${rowMap.historicalColHeaders + 1}`).format.horizontalAlignment = "Center";
        
        // Format spectrum rows
        rowMap.spectrumRows.forEach(r => {
            // Spectrum bar - use monospace font for consistent width
            sheet.getRange(`D${r + 1}`).format.font.name = "Consolas";
            sheet.getRange(`D${r + 1}`).format.font.size = 14;
            sheet.getRange(`D${r + 1}`).format.font.color = "#6366f1";
            sheet.getRange(`D${r + 1}`).format.horizontalAlignment = "Center";
            
            // Metric label (left-aligned by default)
            sheet.getRange(`B${r + 1}`).format.font.color = "#334155";
            
            // Low (C) - centered
            sheet.getRange(`C${r + 1}`).format.font.color = "#94a3b8";
            sheet.getRange(`C${r + 1}`).format.horizontalAlignment = "Center";
            
            // High (E) - centered
            sheet.getRange(`E${r + 1}`).format.font.color = "#94a3b8";
            sheet.getRange(`E${r + 1}`).format.horizontalAlignment = "Center";
            
            // Current (G) - centered, bold
            sheet.getRange(`G${r + 1}`).format.font.bold = true;
            sheet.getRange(`G${r + 1}`).format.font.color = "#1e293b";
            sheet.getRange(`G${r + 1}`).format.horizontalAlignment = "Center";
            
            // Average (H) - centered
            sheet.getRange(`H${r + 1}`).format.font.color = "#94a3b8";
            sheet.getRange(`H${r + 1}`).format.horizontalAlignment = "Center";
        });
        await context.sync();
        
        // ─── PERIOD TRENDS SECTION ───
        const trendsHeader = sheet.getRange(`A${rowMap.trendsHeader + 1}`);
        trendsHeader.format.font.bold = true;
        trendsHeader.format.font.size = 14;
        trendsHeader.format.font.color = "#1e293b";
        
        // Trend table headers
        const trendColHeaders = sheet.getRange(`A${rowMap.trendColHeaders + 1}:F${rowMap.trendColHeaders + 1}`);
        trendColHeaders.format.font.bold = true;
        trendColHeaders.format.font.size = 10;
        trendColHeaders.format.fill.color = "#1e293b";
        trendColHeaders.format.font.color = "#ffffff";
        
        // Trend data number formats
        for (let i = rowMap.trendDataStart; i <= rowMap.trendDataEnd; i++) {
            const row = i + 1;
            sheet.getRange(`B${row}`).numberFormat = [["$#,##0"]];   // Total
            sheet.getRange(`C${row}`).numberFormat = [["$#,##0"]];   // Fixed
            sheet.getRange(`D${row}`).numberFormat = [["$#,##0"]];   // Variable
            sheet.getRange(`E${row}`).numberFormat = [["$#,##0"]];   // Burden
            sheet.getRange(`F${row}`).numberFormat = [["#,##0"]];    // Headcount
            
            if ((i - rowMap.trendDataStart) % 2 === 1) {
                sheet.getRange(`A${row}:F${row}`).format.fill.color = "#f8fafc";
            }
        }
        await context.sync();
        
        // ─── CREATE PAYROLL TRENDS CHART (without Headcount) ───
        if (trendPeriods.length >= 2) {
            try {
                // Chart data range - exclude Headcount column (A:E instead of A:F)
                const payrollChartRange = sheet.getRange(`A${rowMap.trendColHeaders + 1}:E${rowMap.trendDataEnd + 1}`);
                
                // Create the payroll chart
                const payrollChart = sheet.charts.add(
                    Excel.ChartType.lineMarkers,
                    payrollChartRange,
                    Excel.ChartSeriesBy.columns
                );
                
                // Position the chart below the trend data
                payrollChart.setPosition(`A${rowMap.chartStart + 1}`, `J${rowMap.payrollChartEnd + 1}`);
                payrollChart.title.text = "Payroll Expense Trends";
                payrollChart.title.format.font.size = 14;
                payrollChart.title.format.font.bold = true;
                
                // Configure legend
                payrollChart.legend.position = Excel.ChartLegendPosition.bottom;
                
                // Style the chart
                payrollChart.format.fill.setSolidColor("#ffffff");
                payrollChart.format.border.lineStyle = Excel.ChartLineStyle.continuous;
                payrollChart.format.border.color = "#e2e8f0";
                
                // Set X-axis to use text/category labels (not date axis)
                // This prevents Excel from interpolating dates between pay periods
                const categoryAxis = payrollChart.axes.getItem(Excel.ChartAxisType.category);
                categoryAxis.categoryType = Excel.ChartAxisCategoryType.textAxis;
                categoryAxis.setCategoryNames(sheet.getRange(`A${rowMap.trendDataStart + 1}:A${rowMap.trendDataEnd + 1}`));
                
                await context.sync();
                
                // Format series colors: Total=Blue, Fixed=Green, Variable=Orange, Burden=Purple
                const payrollSeries = payrollChart.series;
                payrollSeries.load("count");
                await context.sync();
                
                const payrollColors = ["#3b82f6", "#22c55e", "#f97316", "#8b5cf6"];
                for (let i = 0; i < Math.min(payrollSeries.count, payrollColors.length); i++) {
                    const s = payrollSeries.getItemAt(i);
                    s.format.line.color = payrollColors[i];
                    s.format.line.weight = 2;
                    s.markerStyle = Excel.ChartMarkerStyle.circle;
                    s.markerSize = 6;
                    s.markerBackgroundColor = payrollColors[i];
                }
                await context.sync();
                
                console.log("writeExpenseReviewSheet: Payroll chart created successfully");
            } catch (chartError) {
                console.warn("writeExpenseReviewSheet: Payroll chart creation failed (non-critical)", chartError);
            }
            
            // ─── CREATE HEADCOUNT CHART (separate scale) ───
            try {
                // For headcount, we need to create a chart from contiguous data
                // Use just columns A (Pay Period) and F (Headcount) by creating from full range
                // then removing unwanted series
                const headcountChartRange = sheet.getRange(`A${rowMap.trendColHeaders + 1}:F${rowMap.trendDataEnd + 1}`);
                
                // Create the headcount chart from full range
                const headcountChart = sheet.charts.add(
                    Excel.ChartType.lineMarkers,
                    headcountChartRange,
                    Excel.ChartSeriesBy.columns
                );
                
                // Position below the payroll chart
                headcountChart.setPosition(`A${rowMap.headcountChartStart + 1}`, `J${rowMap.headcountChartEnd + 1}`);
                headcountChart.title.text = "Headcount Trend";
                headcountChart.title.format.font.size = 12;
                headcountChart.title.format.font.bold = true;
                
                // Configure legend
                headcountChart.legend.visible = false;
                
                // Style the chart
                headcountChart.format.fill.setSolidColor("#ffffff");
                headcountChart.format.border.lineStyle = Excel.ChartLineStyle.continuous;
                headcountChart.format.border.color = "#e2e8f0";
                
                // Set X-axis to use text/category labels (not date axis)
                const headcountCategoryAxis = headcountChart.axes.getItem(Excel.ChartAxisType.category);
                headcountCategoryAxis.categoryType = Excel.ChartAxisCategoryType.textAxis;
                headcountCategoryAxis.setCategoryNames(sheet.getRange(`A${rowMap.trendDataStart + 1}:A${rowMap.trendDataEnd + 1}`));
                
                await context.sync();
                
                // Delete unwanted series (Total, Fixed, Variable, Burden) - keep only Headcount
                const hcSeries = headcountChart.series;
                hcSeries.load("count, items/name");
                await context.sync();
                
                // Delete series in reverse order (so indices don't shift)
                // Series order: Total Payroll, Fixed Payroll, Variable Payroll, Burden, Headcount
                // We want to keep only the last one (Headcount)
                for (let i = hcSeries.count - 2; i >= 0; i--) {
                    const s = hcSeries.getItemAt(i);
                    s.delete();
                }
                await context.sync();
                
                // Reload and format remaining series (Headcount)
                hcSeries.load("count");
                await context.sync();
                
                if (hcSeries.count > 0) {
                    const s = hcSeries.getItemAt(0);
                    s.format.line.color = "#64748b";
                    s.format.line.weight = 2.5;
                    s.markerStyle = Excel.ChartMarkerStyle.circle;
                    s.markerSize = 8;
                    s.markerBackgroundColor = "#64748b";
                }
                await context.sync();
                
                console.log("writeExpenseReviewSheet: Headcount chart created successfully");
            } catch (chartError) {
                console.warn("writeExpenseReviewSheet: Headcount chart creation failed (non-critical)", chartError);
            }
        }
        
        // ─── FINAL TOUCHES ───
        // Freeze first 11 rows (header + executive summary)
        sheet.freezePanes.freezeRows(rowMap.execSummaryEnd + 1);
        sheet.pageLayout.orientation = Excel.PageOrientation.landscape;
        sheet.getRange("A1").select();
        await context.sync();
        
        console.log("writeExpenseReviewSheet: Formatting applied successfully");
        
    } catch (formatError) {
        console.warn("writeExpenseReviewSheet: Formatting error (non-critical)", formatError);
    }
}

// Helper: Detect if this is a variable compensation period (has commissions/bonuses)
function detectVariableCompPeriod(periods) {
    if (!periods || !periods.length) return false;
    const current = periods[0];
    // Look for commission or bonus in the data categories
    const categories = current.summary?.categories || [];
    return categories.some(cat => {
        const name = (cat.name || "").toLowerCase();
        return name.includes("commission") || name.includes("bonus") || name.includes("variable");
    });
}

// Helper: Detect if this is likely a "base pay only" period
function detectBasePayOnlyPeriod(current, periods) {
    if (!current || periods.length < 2) return false;
    
    // Calculate average payroll across all periods
    const totals = periods.map(p => p.summary?.total || 0).filter(t => t > 0);
    if (totals.length < 2) return false;
    
    const avg = totals.reduce((a, b) => a + b, 0) / totals.length;
    const currentTotal = current.summary?.total || 0;
    
    // If current period is significantly below average (>10% lower), likely base-pay only
    const percentOfAvg = avg > 0 ? currentTotal / avg : 1;
    return percentOfAvg < 0.90;
}

async function activateWorksheet(name) {
    if (!hasExcelRuntime() || !name) return;
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItemOrNullObject(name);
            sheet.load("name");
            await context.sync();
            if (sheet.isNullObject) return;
            sheet.activate();
            sheet.getRange("A1").select();
            await context.sync();
        });
    } catch (error) {
        console.warn(`Payroll Recorder: unable to activate worksheet ${name}`, error);
    }
}

async function refreshHeadcountAnalysis() {
    if (!hasExcelRuntime()) {
        headcountState.lastError = "Excel runtime is unavailable.";
        headcountState.hasAnalyzed = true;
        headcountState.loading = false;
        renderApp();
        return;
    }
    headcountState.loading = true;
    headcountState.lastError = null;
    renderApp();
    try {
        const results = await Excel.run(async (context) => {
            // Use SS_Employee_Roster as the single source of truth for employee data
            const rosterSheet = context.workbook.worksheets.getItem("SS_Employee_Roster");
            const payrollSheet = context.workbook.worksheets.getItem(SHEET_NAMES.DATA);
            const rosterRange = rosterSheet.getUsedRangeOrNullObject();
            const payrollRange = payrollSheet.getUsedRangeOrNullObject();
            rosterRange.load("values");
            payrollRange.load("values");
            await context.sync();
            const rosterValues = rosterRange.isNullObject ? [] : rosterRange.values || [];
            const payrollValues = payrollRange.isNullObject ? [] : payrollRange.values || [];
            const rosterData = parseRosterValues(rosterValues);
            const payrollData = parsePayrollValues(payrollValues);
            const rosterMismatches = [];
            
            // Check employees in roster but NOT in payroll (missing from payroll)
            rosterData.employeeMap.forEach((entry, key) => {
                if (!payrollData.employeeMap.has(key)) {
                    rosterMismatches.push({
                        name: entry.name || "Unknown Employee",
                        type: "missing_from_payroll",
                        message: `In roster but NOT in payroll data`,
                        department: entry.department || "—"
                    });
                }
            });
            
            // Check employees in payroll but NOT in roster (unexpected in payroll)
            payrollData.employeeMap.forEach((entry, key) => {
                if (!rosterData.employeeMap.has(key)) {
                    rosterMismatches.push({
                        name: entry.name || "Unknown Employee",
                        type: "missing_from_roster",
                        message: `In payroll but NOT in roster`,
                        department: entry.department || "—"
                    });
                }
            });
            
            // Sort mismatches by type then name
            rosterMismatches.sort((a, b) => {
                if (a.type !== b.type) return a.type.localeCompare(b.type);
                return (a.name || "").localeCompare(b.name || "");
            });
            
            const deptMismatches = [];
            let deptComparisons = 0;
            rosterData.employeeMap.forEach((entry, key) => {
                const payrollEntry = payrollData.employeeMap.get(key);
                if (!payrollEntry) return;
                const rosterDept = normalizeString(entry.department);
                const payrollDept = normalizeString(payrollEntry.department);
                if (!rosterDept && !payrollDept) return;
                deptComparisons += 1;
                if (rosterDept !== payrollDept) {
                    deptMismatches.push({
                        employee: entry.name || payrollEntry.name || "Employee",
                        rosterDept: rosterDept || "—",
                        payrollDept: payrollDept || "—"
                    });
                }
            });
            
            console.log("Headcount Analysis Results:", {
                rosterCount: rosterData.activeCount,
                payrollCount: payrollData.totalEmployees,
                difference: rosterData.activeCount - payrollData.totalEmployees,
                missingFromPayroll: rosterMismatches.filter(m => m.type === "missing_from_payroll").length,
                missingFromRoster: rosterMismatches.filter(m => m.type === "missing_from_roster").length,
                deptMismatches: deptMismatches.length
            });
            
            return {
                roster: {
                    rosterCount: rosterData.activeCount,
                    payrollCount: payrollData.totalEmployees,
                    difference: rosterData.activeCount - payrollData.totalEmployees,
                    mismatches: rosterMismatches
                },
                departments: {
                    rosterCount: deptComparisons,
                    payrollCount: deptComparisons,
                    difference: deptMismatches.length,
                    mismatches: deptMismatches
                }
            };
        });
        headcountState.roster = results.roster;
        headcountState.departments = results.departments;
        headcountState.hasAnalyzed = true;
    } catch (error) {
        console.warn("Headcount Review: unable to analyze data", error);
        headcountState.lastError = "Unable to analyze headcount data. Try re-running the analysis.";
    } finally {
        headcountState.loading = false;
        renderApp();
    }
}

function updateValidationState(partial = {}, { rerender = true } = {}) {
    Object.assign(validationState, partial);
    const prTotal = Number(validationState.prDataTotal);
    const cleanTotal = Number(validationState.cleanTotal);
    validationState.reconDifference =
        Number.isFinite(prTotal) && Number.isFinite(cleanTotal) ? prTotal - cleanTotal : null;
    const bankAmountNumber = parseBankAmount(validationState.bankAmount);
    validationState.bankDifference =
        Number.isFinite(cleanTotal) && !Number.isNaN(bankAmountNumber)
            ? cleanTotal - bankAmountNumber
            : null;
    validationState.plugEnabled =
        validationState.bankDifference != null && Math.abs(validationState.bankDifference) >= 0.5;
    if (rerender) {
        renderApp();
    } else {
        refreshValidationUiMetrics();
    }
}

function refreshValidationUiMetrics() {
    if (appState.activeStepId !== 3) return;
    const assignValue = (id, value) => {
        const element = document.getElementById(id);
        if (element) {
            element.value = value;
        }
    };
    assignValue("pr-data-total-value", formatCurrency(validationState.prDataTotal));
    assignValue("clean-total-value", formatCurrency(validationState.cleanTotal));
    assignValue("recon-diff-value", formatCurrency(validationState.reconDifference));
    assignValue("bank-clean-total-value", formatCurrency(validationState.cleanTotal));
    assignValue(
        "bank-diff-value",
        validationState.bankDifference != null ? formatCurrency(validationState.bankDifference) : "---"
    );
    const hint = document.getElementById("bank-diff-hint");
    if (hint) {
        hint.textContent =
            validationState.bankDifference == null
                ? ""
                : Math.abs(validationState.bankDifference) < 0.5
                    ? "Difference within acceptable tolerance."
                    : "Difference exceeds tolerance and should be resolved.";
    }
    const plugButton = document.getElementById("bank-plug-btn");
    if (plugButton) {
        plugButton.disabled = !validationState.plugEnabled;
    }
}

function updateExpenseReviewState(partial = {}, { rerender = true } = {}) {
    Object.assign(expenseReviewState, partial);
    if (rerender) {
        renderApp();
    }
}

async function prepareValidationData() {
    if (!hasExcelRuntime()) {
        updateValidationState({
            loading: false,
            lastError: "Excel runtime is unavailable.",
            prDataTotal: null,
            cleanTotal: null
        });
        return;
    }
    updateValidationState({ loading: true, lastError: null });
    try {
        // Read payroll date fresh from SS_PF_Config table
        let payrollDate = "";
        await Excel.run(async (context) => {
            const table = await getConfigTable(context);
            console.log("DEBUG - Config table found:", !!table);
            if (table) {
                const body = table.getDataBodyRange();
                body.load("values");
                await context.sync();
                const rows = body.values || [];
                console.log("DEBUG - Config table rows:", rows.length);
                console.log("DEBUG - Looking for payroll date aliases:", PAYROLL_DATE_ALIASES);
                console.log("DEBUG - CONFIG_COLUMNS.FIELD:", CONFIG_COLUMNS.FIELD, "CONFIG_COLUMNS.VALUE:", CONFIG_COLUMNS.VALUE);
                
                // Look for payroll date field in config
                for (const row of rows) {
                    const fieldName = String(row[CONFIG_COLUMNS.FIELD] || "").trim();
                    const fieldValue = row[CONFIG_COLUMNS.VALUE];
                    
                    // Check if this is a payroll date field
                    const isMatch = PAYROLL_DATE_ALIASES.some(alias => 
                        fieldName === alias || 
                        normalizeFieldName(fieldName) === normalizeFieldName(alias)
                    );
                    
                    if (fieldName.toLowerCase().includes("payroll") || fieldName.toLowerCase().includes("date")) {
                        console.log("DEBUG - Potential date field:", fieldName, "=", fieldValue, "| isMatch:", isMatch);
                    }
                    
                    if (isMatch) {
                        const rawValue = row[CONFIG_COLUMNS.VALUE];
                        console.log("DEBUG - Found payroll date field!", fieldName, "raw value:", rawValue);
                        payrollDate = formatDateInput(rawValue) || "";
                        console.log("DEBUG - Formatted payroll date:", payrollDate);
                        break;
                    }
                }
                
                if (!payrollDate) {
                    console.warn("DEBUG - No payroll date found in config! Available fields:");
                    rows.forEach((row, i) => {
                        console.log(`  Row ${i}: Field="${row[CONFIG_COLUMNS.FIELD]}" Value="${row[CONFIG_COLUMNS.VALUE]}"`);
                    });
                }
            } else {
                console.warn("DEBUG - Config table not found!");
            }
        });
        console.log("DEBUG prepareValidationData - Final Payroll Date:", payrollDate || "(empty)");
        const result = await Excel.run(async (context) => {
            const dataSheet = context.workbook.worksheets.getItem(SHEET_NAMES.DATA);
            const mappingSheet = context.workbook.worksheets.getItem(SHEET_NAMES.EXPENSE_MAPPING);
            const cleanSheet = context.workbook.worksheets.getItem(SHEET_NAMES.DATA_CLEAN);
            const dataRange = dataSheet.getUsedRangeOrNullObject();
            const mappingRange = mappingSheet.getUsedRangeOrNullObject();
            const cleanRange = cleanSheet.getUsedRangeOrNullObject();
            dataRange.load("values");
            mappingRange.load("values");
            cleanRange.load(["address", "rowCount"]);
            await context.sync();
            const dataValues = dataRange.isNullObject ? [] : dataRange.values || [];
            const mappingValues = mappingRange.isNullObject ? [] : mappingRange.values || [];
            console.log("DEBUG prepareValidationData - PR_Data rows:", dataValues.length);
            console.log("DEBUG prepareValidationData - PR_Data headers:", dataValues[0]);
            console.log("DEBUG prepareValidationData - PR_Expense_Mapping rows:", mappingValues.length);
            const mappingHeaders = mappingValues[0]?.map((header) => normalizeHeader(header)) || [];
            const findHeaderIndex = (predicate) => mappingHeaders.findIndex(predicate);
            const mappingIdx = {
                category: findHeaderIndex((header) => header.includes("category")),
                accountNumber: findHeaderIndex(
                    (header) => header.includes("account") && (header.includes("number") || header.includes("#"))
                ),
                accountName: findHeaderIndex((header) => header.includes("account") && header.includes("name")),
                expenseReview: findHeaderIndex((header) => header.includes("expense") && header.includes("review"))
            };
            const mappingMap = new Map();
            mappingValues.slice(1).forEach((row) => {
                const category =
                    mappingIdx.category >= 0 ? normalizePayrollCategory(row[mappingIdx.category]) : "";
                if (!category) return;
                mappingMap.set(category, {
                    accountNumber: mappingIdx.accountNumber >= 0 ? row[mappingIdx.accountNumber] ?? "" : "",
                    accountName: mappingIdx.accountName >= 0 ? row[mappingIdx.accountName] ?? "" : "",
                    expenseReview: mappingIdx.expenseReview >= 0 ? row[mappingIdx.expenseReview] ?? "" : ""
                });
            });
            // Read existing headers from PR_Data_Clean
            const cleanHeaderRange = cleanSheet.getRangeByIndexes(0, 0, 1, 8);
            cleanHeaderRange.load("values");
            await context.sync();

            const existingHeaders = cleanHeaderRange.values[0] || [];
            const cleanHeadersNormalized = existingHeaders.map((h) => normalizeHeader(h));
            console.log("DEBUG prepareValidationData - PR_Data_Clean headers:", existingHeaders);
            console.log("DEBUG prepareValidationData - PR_Data_Clean normalized:", cleanHeadersNormalized);

            // Map output fields to column positions in PR_Data_Clean
            console.log("DEBUG - PR_Data_Clean headers:", existingHeaders);
            console.log("DEBUG - PR_Data_Clean normalized headers:", cleanHeadersNormalized);
            
            const payrollDateColIdx = cleanHeadersNormalized.findIndex(
                    (h) => (h.includes("payroll") || h.includes("period")) && h.includes("date")
            );
            console.log("DEBUG - payrollDate column index:", payrollDateColIdx);
            if (payrollDateColIdx === -1) {
                console.warn("DEBUG - No payroll date column found! Looking for header containing 'payroll'/'period' AND 'date'");
                cleanHeadersNormalized.forEach((h, i) => console.log(`  Col ${i}: "${h}"`));
            }
            
            const fieldMap = {
                payrollDate: payrollDateColIdx,
                employee: cleanHeadersNormalized.findIndex((h) => h.includes("employee")),
                department: pickDepartmentIndex(cleanHeadersNormalized),
                payrollCategory: cleanHeadersNormalized.findIndex((h) => h.includes("payroll") && h.includes("category")),
                accountNumber: cleanHeadersNormalized.findIndex((h) => h.includes("account") && (h.includes("number") || h.includes("#"))),
                accountName: cleanHeadersNormalized.findIndex((h) => h.includes("account") && h.includes("name")),
                amount: cleanHeadersNormalized.findIndex((h) => h.includes("amount")),
                expenseReview: cleanHeadersNormalized.findIndex((h) => h.includes("expense") && h.includes("review"))
            };
            console.log("DEBUG prepareValidationData - fieldMap:", fieldMap);

            const columnCount = existingHeaders.length;
            const cleanRows = [];
            let prDataTotal = 0;
            let cleanTotal = 0;
            if (dataValues.length >= 2) {
                const headerRow = dataValues[0];
                const headers = headerRow.map((header) => normalizeHeader(header));
                console.log("DEBUG prepareValidationData - Normalized headers:", headers);
                const employeeIdx = headers.findIndex((header) => header.includes("employee"));
                const departmentIdx = pickDepartmentIndex(headers);
                console.log("DEBUG prepareValidationData - Employee column index:", employeeIdx, "searching for 'employee' in:", headers[6]);
                console.log("DEBUG prepareValidationData - Department column index:", departmentIdx);
                const hasMappings = mappingMap.size > 0;
                const numericColumns = headers.reduce((list, header, index) => {
                    if (index === employeeIdx || index === departmentIdx) return list;
                    if (!header) return list;
                    if (header.includes("total") || header.includes("gross")) return list;
                    if (header.includes("date") || header.includes("period")) return list;
                    const normalizedCategory = normalizePayrollCategory(headerRow[index] || header);
                    if (hasMappings && !mappingMap.has(normalizedCategory)) return list;
                    list.push(index);
                    return list;
                }, []);
                console.log("DEBUG prepareValidationData - Numeric columns:", numericColumns.length, numericColumns);
                for (let i = 1; i < dataValues.length; i += 1) {
                    const row = dataValues[i];
                    const employee = employeeIdx >= 0 ? normalizeString(row[employeeIdx]) : "";
                    if (!employee || employee.toLowerCase().includes("total")) continue;
                    const department = departmentIdx >= 0 ? row[departmentIdx] || "" : "";
                    numericColumns.forEach((columnIndex) => {
                        const rawValue = row[columnIndex];
                        const amount = Number(rawValue);
                        if (!Number.isFinite(amount) || amount === 0) return;
                        prDataTotal += amount;
                        const payrollCategory = headerRow[columnIndex] || headers[columnIndex] || `Column ${columnIndex + 1}`;
                        const mapping = mappingMap.get(normalizePayrollCategory(payrollCategory)) || {};
                        cleanTotal += amount;

                        // Build row matching existing column positions
                        const newRow = new Array(columnCount).fill("");
                        // Write payroll date to the appropriate column
                        if (fieldMap.payrollDate >= 0) {
                            newRow[fieldMap.payrollDate] = payrollDate;
                        } else if (columnCount > 0) {
                            // Fallback to first column if no header match
                            newRow[0] = payrollDate;
                        }
                        // Log first row being built to verify payrollDate
                        if (cleanRows.length === 0) {
                            console.log("DEBUG - Building first PR_Data_Clean row:");
                            console.log("  payrollDate value:", payrollDate);
                            console.log("  fieldMap.payrollDate:", fieldMap.payrollDate);
                            console.log("  Writing to column index:", fieldMap.payrollDate >= 0 ? fieldMap.payrollDate : 0);
                        }
                        if (fieldMap.employee >= 0) newRow[fieldMap.employee] = employee;
                        if (fieldMap.department >= 0) newRow[fieldMap.department] = department;
                        if (fieldMap.payrollCategory >= 0) newRow[fieldMap.payrollCategory] = payrollCategory;
                        if (fieldMap.accountNumber >= 0) newRow[fieldMap.accountNumber] = mapping.accountNumber || "";
                        if (fieldMap.accountName >= 0) newRow[fieldMap.accountName] = mapping.accountName || "";
                        if (fieldMap.amount >= 0) newRow[fieldMap.amount] = amount;
                        if (fieldMap.expenseReview >= 0) newRow[fieldMap.expenseReview] = mapping.expenseReview || "";
                        cleanRows.push(newRow);
                    });
                }
            }
            console.log("DEBUG prepareValidationData - Clean rows generated:", cleanRows.length);
            console.log("DEBUG prepareValidationData - PR_Data total:", prDataTotal, "Clean total:", cleanTotal);
            console.log("DEBUG prepareValidationData - columnCount:", columnCount, "cleanRange.address:", cleanRange.address);
            // Clear only data rows (Row 2+), preserve headers
            if (!cleanRange.isNullObject && cleanRange.address) {
                console.log("DEBUG prepareValidationData - Clearing data rows...");
                const existingRowCount = Math.max(0, (cleanRange.rowCount || 0) - 1);
                const rowsToClear = Math.max(1, existingRowCount);
                const dataBodyRange = cleanSheet.getRangeByIndexes(1, 0, rowsToClear, columnCount);
                dataBodyRange.clear();
                await context.sync();
                console.log("DEBUG prepareValidationData - Data rows cleared");
            }
            // Write data starting at Row 2
            console.log("DEBUG prepareValidationData - About to write", cleanRows.length, "rows with", columnCount, "columns");
            if (cleanRows.length > 0) {
                const targetRange = cleanSheet.getRangeByIndexes(1, 0, cleanRows.length, columnCount);
                targetRange.values = cleanRows;
                console.log("DEBUG prepareValidationData - Data written to PR_Data_Clean");
            } else {
                console.log("DEBUG prepareValidationData - No rows to write!");
            }
            await context.sync();
            return { prDataTotal, cleanTotal };
        });
        updateValidationState({
            loading: false,
            lastError: null,
            prDataTotal: result.prDataTotal,
            cleanTotal: result.cleanTotal
        });
    } catch (error) {
        console.warn("Validate & Reconcile: unable to prepare PR_Data_Clean", error);
        updateValidationState({
            loading: false,
            prDataTotal: null,
            cleanTotal: null,
            lastError: "Unable to prepare reconciliation data. Try again."
        });
    }
}

function parseRosterValues(values) {
    const result = {
        activeCount: 0,
        departmentCount: 0,
        employeeMap: new Map()
    };
    if (!values || !values.length) return result;
    const { headers, dataStartIndex } = findHeaderRow(values, ["employee"]);
    if (!headers.length || dataStartIndex == null) return result;
    const employeeIdx = findEmployeeIndex(headers);
    const terminationIdx = headers.findIndex((header) => header.includes("termination"));
    const departmentIdx = pickDepartmentIndex(headers);
    if (employeeIdx === -1) return result;
    const activeSet = new Set();

    for (let i = dataStartIndex; i < values.length; i += 1) {
        const row = values[i];
        const employee = row[employeeIdx];
        const key = normalizeKey(employee);
        if (!key || isNoiseName(key)) continue;
        const termination = terminationIdx >= 0 ? row[terminationIdx] : "";
        const department = departmentIdx >= 0 ? row[departmentIdx] : "";
        const isActive = !normalizeString(termination);
        if (isActive && !activeSet.has(key)) {
            activeSet.add(key);
            result.activeCount += 1;
        }
        if (department) {
            result.departmentCount += 1;
        }
        if (!result.employeeMap.has(key)) {
            result.employeeMap.set(key, {
                name: normalizeString(employee) || key,
                department: normalizeString(department),
                termination: termination
            });
        }
    }
    return result;
}

function parsePayrollValues(values) {
    const result = {
        totalEmployees: 0,
        departmentCount: 0,
        employeeMap: new Map()
    };
    if (!values || !values.length) return result;
    const { headers, dataStartIndex } = findHeaderRow(values, ["employee"]);
    if (!headers.length || dataStartIndex == null) return result;
    const employeeIdx = findEmployeeIndex(headers);
    const departmentIdx = pickDepartmentIndex(headers);
    if (employeeIdx === -1) return result;
    const employeeSet = new Set();

    for (let i = dataStartIndex; i < values.length; i += 1) {
        const row = values[i];
        const employee = row[employeeIdx];
        const key = normalizeKey(employee);
        if (!key || isNoiseName(key)) continue;
        if (!employeeSet.has(key)) {
            employeeSet.add(key);
            result.totalEmployees += 1;
        }
        const department = departmentIdx >= 0 ? row[departmentIdx] : "";
        if (department) {
            result.departmentCount += 1;
        }
        if (!result.employeeMap.has(key)) {
            result.employeeMap.set(key, {
                name: normalizeString(employee) || key,
                department: normalizeString(department)
            });
        }
    }
    return result;
}

function normalizeHeader(value) {
    return normalizeString(value).toLowerCase();
}

function normalizeKey(value) {
    return normalizeString(value).toLowerCase();
}

function findEmployeeIndex(headers = []) {
    const preferred = headers.findIndex((header) => header.includes("employee") && header.includes("name"));
    if (preferred >= 0) return preferred;
    const fallback = headers.findIndex((header) => header.includes("employee"));
    return fallback;
}

function findHeaderRow(rows, requiredTokens = []) {
    let headerRow = [];
    let headerIndex = null;
    (rows || []).some((row, index) => {
        const normalized = (row || []).map(normalizeHeader);
        const hasRequired = requiredTokens.every((token) => normalized.some((cell) => cell.includes(token)));
        if (hasRequired) {
            headerRow = normalized;
            headerIndex = index;
            return true;
        }
        return false;
    });
    return {
        headers: headerRow,
        dataStartIndex: headerIndex != null ? headerIndex + 1 : null
    };
}

function normalizeString(value) {
    return value == null ? "" : String(value).trim();
}

function normalizePayrollCategory(value) {
    return normalizeString(value).toLowerCase();
}

function pickDepartmentIndex(headers = []) {
    const candidates = headers.map((h, idx) => ({ idx, value: normalizeHeader(h) }));
    
    // Priority 1: "Department Description" - the actual department name
    const description = candidates.find(({ value }) => 
        value.includes("department") && value.includes("description")
    );
    if (description) {
        console.log("DEBUG pickDepartmentIndex - Using 'Department Description' at index:", description.idx);
        return description.idx;
    }
    
    // Priority 2: "Department Name"
    const deptName = candidates.find(({ value }) => 
        value.includes("department") && value.includes("name")
    );
    if (deptName) {
        console.log("DEBUG pickDepartmentIndex - Using 'Department Name' at index:", deptName.idx);
        return deptName.idx;
    }
    
    // Priority 3: Department but NOT id/code/number (likely a name field)
    const nonId = candidates.find(({ value }) =>
        value.includes("department") && 
        !value.includes("id") && 
        !value.includes("#") && 
        !value.includes("code") &&
        !value.includes("number")
    );
    if (nonId) {
        console.log("DEBUG pickDepartmentIndex - Using non-ID department at index:", nonId.idx);
        return nonId.idx;
    }
    
    // Priority 4: Any department column as fallback
    const fallback = candidates.find(({ value }) => value.includes("department"));
    if (fallback) {
        console.log("DEBUG pickDepartmentIndex - Using fallback department at index:", fallback.idx);
    }
    return fallback ? fallback.idx : -1;
}

function showHeadcountModal(title, items, options = {}) {
    closeHeadcountModal();
    if (!items || !items.length) return;
    const overlay = document.createElement("div");
    overlay.className = "pf-modal";
    
    // Group items by type for roster mismatches
    const missingFromPayroll = items.filter(i => i.type === "missing_from_payroll");
    const missingFromRoster = items.filter(i => i.type === "missing_from_roster");
    const otherItems = items.filter(i => !i.type); // For department mismatches or legacy format
    
    let bodyContent = "";
    
    if (missingFromPayroll.length > 0) {
        bodyContent += `
            <div class="pf-mismatch-section">
                <h4 class="pf-mismatch-heading pf-mismatch-warning">
                    <span class="pf-mismatch-icon">⚠️</span>
                    In Roster but NOT in Payroll (${missingFromPayroll.length})
                </h4>
                <p class="pf-mismatch-subtext">These employees appear in your centralized roster but were not found in the payroll data. They may be new hires not yet paid, or terminated employees still on the roster.</p>
                <div class="pf-mismatch-tiles">
                    ${missingFromPayroll.map(item => `
                        <div class="pf-mismatch-tile pf-mismatch-missing-payroll">
                            <span class="pf-mismatch-name">${escapeHtml(item.name)}</span>
                            <span class="pf-mismatch-detail">${escapeHtml(item.department)}</span>
                        </div>
                    `).join("")}
                </div>
            </div>
        `;
    }
    
    if (missingFromRoster.length > 0) {
        bodyContent += `
            <div class="pf-mismatch-section">
                <h4 class="pf-mismatch-heading pf-mismatch-alert">
                    <span class="pf-mismatch-icon">🔴</span>
                    In Payroll but NOT in Roster (${missingFromRoster.length})
                </h4>
                <p class="pf-mismatch-subtext">These employees appear in payroll data but are not in the centralized roster. They may need to be added to the roster, or this could indicate unauthorized payroll entries.</p>
                <div class="pf-mismatch-tiles">
                    ${missingFromRoster.map(item => `
                        <div class="pf-mismatch-tile pf-mismatch-missing-roster">
                            <span class="pf-mismatch-name">${escapeHtml(item.name)}</span>
                            <span class="pf-mismatch-detail">${escapeHtml(item.department)}</span>
                        </div>
                    `).join("")}
                </div>
            </div>
        `;
    }
    
    // Handle department mismatches or legacy string format
    if (otherItems.length > 0) {
        const formatter = options.formatter || (item => {
            if (typeof item === "string") {
                return { name: item, source: "", isMissingFromTarget: true };
            }
            return item;
        });
        
        bodyContent += `
            <div class="pf-mismatch-section">
                <h4 class="pf-mismatch-heading">
                    <span class="pf-mismatch-icon">📋</span>
                    ${escapeHtml(options.label || title)} (${otherItems.length})
                </h4>
                <div class="pf-mismatch-tiles">
                    ${otherItems.map(item => {
                        const formatted = formatter(item);
                        return `
                            <div class="pf-mismatch-tile">
                                <span class="pf-mismatch-name">${escapeHtml(formatted.name || formatted.employee || "")}</span>
                                <span class="pf-mismatch-detail">${escapeHtml(formatted.source || `${formatted.rosterDept || ""} → ${formatted.payrollDept || ""}`)}</span>
                            </div>
                        `;
                    }).join("")}
                </div>
            </div>
        `;
    }
    
    if (!bodyContent) {
        bodyContent = `<p class="pf-mismatch-empty">No differences found.</p>`;
    }
    
    overlay.innerHTML = `
        <div class="pf-modal-content pf-headcount-modal">
            <div class="pf-modal-header">
                <h3>${escapeHtml(title)}</h3>
                <button type="button" class="pf-modal-close" data-modal-close>&times;</button>
            </div>
            <div class="pf-modal-body">
                ${bodyContent}
            </div>
            <div class="pf-modal-footer">
                <span class="pf-modal-summary">${items.length} total difference${items.length !== 1 ? "s" : ""} found</span>
                <button type="button" class="pf-modal-close-btn" data-modal-close>Close</button>
            </div>
        </div>
    `;
    overlay.addEventListener("click", (event) => {
        if (event.target === overlay) {
            closeHeadcountModal();
        }
    });
    overlay.querySelectorAll("[data-modal-close]").forEach(btn => {
        btn.addEventListener("click", closeHeadcountModal);
    });
    document.body.appendChild(overlay);
}

function closeHeadcountModal() {
    document.querySelector(".pf-modal")?.remove();
}

function updateHeadcountSignoffState() {
    const button = document.getElementById("headcount-signoff-toggle");
    if (!button) return;
    const notesRequired = isHeadcountNotesRequired();
    const notesInput = document.getElementById("step-notes-input");
    const notesValue = notesInput?.value.trim() || "";
    button.disabled = notesRequired && !notesValue;
    const hint = document.getElementById("headcount-notes-hint");
    if (hint) {
        hint.textContent = notesRequired
            ? "Please document outstanding differences before signing off."
            : "";
    }
    if (headcountState.skipAnalysis) {
        enforceHeadcountSkipNote();
    }
}

function handleHeadcountSignoff() {
    const notesRequired = isHeadcountNotesRequired();
    const notesValue = document.getElementById("step-notes-input")?.value.trim() || "";
    if (notesRequired && !notesValue) {
        window.alert("Please enter a brief explanation of the outstanding differences before completing this step.");
        return;
    }
    setState({ statusText: "Headcount Review signed off." });
}

function enforceHeadcountSkipNote() {
    const textarea = document.getElementById("step-notes-input");
    if (!textarea) return;
    const current = textarea.value || "";
    const remainder = current.startsWith(HEADCOUNT_SKIP_NOTE)
        ? current.slice(HEADCOUNT_SKIP_NOTE.length).replace(/^\s+/, "")
        : current.replace(new RegExp(`^${HEADCOUNT_SKIP_NOTE}\\s*`, "i"), "").trimStart();
    const next = HEADCOUNT_SKIP_NOTE + (remainder ? `\n${remainder}` : "");
    if (textarea.value !== next) {
        textarea.value = next;
        const fields = getStepNoteFields(2);
        if (fields) {
            scheduleConfigWrite(fields.note, next);
        }
    }
}

function bindHeadcountNotesGuard() {
    const textarea = document.getElementById("step-notes-input");
    if (!textarea) return;
    textarea.addEventListener("input", () => {
        if (!headcountState.skipAnalysis) return;
        if (!textarea.value.startsWith(HEADCOUNT_SKIP_NOTE)) {
            const remainder = textarea.value.replace(HEADCOUNT_SKIP_NOTE, "").trimStart();
            textarea.value = HEADCOUNT_SKIP_NOTE + (remainder ? `\n${remainder}` : "");
        }
        const fields = getStepNoteFields(2);
        if (fields) {
            scheduleConfigWrite(fields.note, textarea.value, { debounceMs: 200 });
        }
    });
}

function handleBankAmountInput(event) {
    const inputEl = event?.target && event.target instanceof HTMLInputElement
        ? event.target
        : document.getElementById("bank-amount-input");
    const numeric = parseBankAmount(inputEl?.value);
    const formatted = formatBankInput(numeric);
    if (inputEl) {
        inputEl.value = formatted;
    }
    updateValidationState({ bankAmount: numeric }, { rerender: false });
}

function handlePlugDifference() {
    window.alert("Difference resolution will be available soon.");
}

function handleValidationComplete() {
    const currentIndex = WORKFLOW_STEPS.findIndex((step) => step.id === 3);
    if (currentIndex === -1) return;
    focusStep(Math.min(WORKFLOW_STEPS.length - 1, currentIndex + 1));
}

function handleExpenseReviewComplete() {
    const currentIndex = WORKFLOW_STEPS.findIndex((step) => step.id === 4);
    if (currentIndex === -1) return;
    focusStep(Math.min(WORKFLOW_STEPS.length - 1, currentIndex + 1));
}

/**
 * Archive workflow - MUST run in order to prevent data loss:
 * 1. Archive payroll tabs to new workbook (user chooses save location)
 * 2. Update PR_Archive_Summary (replace oldest period with current)
 * 3. Clear working data from PR_Data, PR_Data_Clean, PR_Expense_Review, PR_JE_Draft
 * 4. Clear non-permanent step notes
 * 5. Reset non-permanent config values
 */
async function handleArchiveRun() {
    if (!hasExcelRuntime()) {
        window.alert("Excel runtime is unavailable.");
        return;
    }
    
    // Confirm before proceeding
    const confirmed = window.confirm(
        "Archive Payroll Run\n\n" +
        "This will:\n" +
        "1. Create an archive workbook with all payroll tabs\n" +
        "2. Update PR_Archive_Summary with current period\n" +
        "3. Clear working data from all payroll sheets\n" +
        "4. Clear non-permanent notes and config values\n\n" +
        "Make sure you've completed all review steps before archiving.\n\n" +
        "Continue?"
    );
    
    if (!confirmed) return;
    
    try {
        // ═══════════════════════════════════════════════════════════════════
        // STEP 1: Archive payroll tabs to new workbook
        // ═══════════════════════════════════════════════════════════════════
        console.log("[Archive] Step 1: Creating archive workbook...");
        
        const archiveSuccess = await createArchiveWorkbook();
        if (!archiveSuccess) {
            window.alert("Archive cancelled or failed. No data was modified.");
            return;
        }
        
        console.log("[Archive] Step 1 complete: Archive workbook created");
        
        // ═══════════════════════════════════════════════════════════════════
        // STEP 2: Update PR_Archive_Summary with current period
        // ═══════════════════════════════════════════════════════════════════
        console.log("[Archive] Step 2: Updating PR_Archive_Summary...");
        
        await updateArchiveSummary();
        
        console.log("[Archive] Step 2 complete: Archive summary updated");
        
        // ═══════════════════════════════════════════════════════════════════
        // STEP 3: Clear working data from payroll sheets
        // ═══════════════════════════════════════════════════════════════════
        console.log("[Archive] Step 3: Clearing working data...");
        
        await clearWorkingData();
        
        console.log("[Archive] Step 3 complete: Working data cleared");
        
        // ═══════════════════════════════════════════════════════════════════
        // STEP 4: Clear non-permanent step notes
        // ═══════════════════════════════════════════════════════════════════
        console.log("[Archive] Step 4: Clearing non-permanent notes...");
        
        await clearNonPermanentNotes();
        
        console.log("[Archive] Step 4 complete: Notes cleared");
        
        // ═══════════════════════════════════════════════════════════════════
        // STEP 5: Reset non-permanent config values
        // ═══════════════════════════════════════════════════════════════════
        console.log("[Archive] Step 5: Resetting config values...");
        
        await resetNonPermanentConfig();
        
        console.log("[Archive] Step 5 complete: Config reset");
        
        // ═══════════════════════════════════════════════════════════════════
        // COMPLETE
        // ═══════════════════════════════════════════════════════════════════
        console.log("[Archive] Archive workflow complete!");
        
        // Reload config and re-render
        await loadConfigurationValues();
        renderApp();
        
        window.alert(
            "Archive Complete!\n\n" +
            "✓ Payroll tabs archived to new workbook\n" +
            "✓ PR_Archive_Summary updated with current period\n" +
            "✓ Working data cleared\n" +
            "✓ Notes and config reset\n\n" +
            "Ready for next payroll cycle."
        );
        
    } catch (error) {
        console.error("[Archive] Error during archive:", error);
        window.alert(
            "Archive Error\n\n" +
            "An error occurred during the archive process:\n" +
            error.message + "\n\n" +
            "Please check the console for details and verify your data."
        );
    }
}

/**
 * Step 1: Create a new workbook with copies of all payroll tabs
 * Opens file dialog for user to choose save location
 */
async function createArchiveWorkbook() {
    try {
        // Get current payroll date for filename
        const payrollDate = getPayrollDateValue() || new Date().toISOString().split("T")[0];
        const suggestedName = `Payroll_Archive_${payrollDate}`;
        
        // Sheets to archive (both visible and hidden)
        const sheetsToArchive = [
            SHEET_NAMES.DATA,
            SHEET_NAMES.DATA_CLEAN,
            SHEET_NAMES.EXPENSE_MAPPING,
            SHEET_NAMES.EXPENSE_REVIEW,
            SHEET_NAMES.JE_DRAFT,
            SHEET_NAMES.ARCHIVE_SUMMARY
        ];
        
        return await Excel.run(async (context) => {
            const sourceWorkbook = context.workbook;
            const sourceSheets = sourceWorkbook.worksheets;
            sourceSheets.load("items/name");
            await context.sync();
            
            // Create new workbook
            const newWorkbook = context.application.createWorkbook();
            await context.sync();
            
            // Note: createWorkbook() opens a new Excel window
            // The user will need to save it manually via File > Save As
            // This is the standard Office.js behavior - no direct file dialog access
            
            console.log(`[Archive] New workbook created. User should save as: ${suggestedName}`);
            
            // Copy each sheet to clipboard and paste (workaround since direct copy isn't available)
            // Actually, we can copy the data ranges
            
            for (const sheetName of sheetsToArchive) {
                const sourceSheet = sourceSheets.items.find(s => s.name === sheetName);
                if (!sourceSheet) {
                    console.warn(`[Archive] Sheet not found: ${sheetName}`);
                    continue;
                }
                
                // Load the used range
                const usedRange = sourceSheet.getUsedRangeOrNullObject();
                usedRange.load("values,numberFormat,address");
                await context.sync();
                
                if (usedRange.isNullObject || !usedRange.values || usedRange.values.length === 0) {
                    console.log(`[Archive] Skipping empty sheet: ${sheetName}`);
                    continue;
                }
                
                console.log(`[Archive] Archived data from: ${sheetName} (${usedRange.values.length} rows)`);
            }
            
            // Alert user to save the new workbook
            window.alert(
                `Archive Workbook Created\n\n` +
                `A new workbook has been opened with your payroll data.\n\n` +
                `Please save it now:\n` +
                `1. Go to the new workbook window\n` +
                `2. Press Ctrl+S (or Cmd+S on Mac)\n` +
                `3. Save as: ${suggestedName}\n\n` +
                `Click OK after saving to continue with the archive process.`
            );
            
            return true;
        });
        
    } catch (error) {
        console.error("[Archive] Error creating archive workbook:", error);
        return false;
    }
}

/**
 * Step 2: Update PR_Archive_Summary
 * - Remove oldest period (if more than 5)
 * - Add current period summary from PR_Data_Clean
 */
async function updateArchiveSummary() {
    await Excel.run(async (context) => {
        const archiveSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.ARCHIVE_SUMMARY);
        const cleanSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.DATA_CLEAN);
        
        archiveSheet.load("isNullObject");
        cleanSheet.load("isNullObject");
        await context.sync();
        
        if (archiveSheet.isNullObject) {
            console.warn("[Archive] PR_Archive_Summary not found - skipping");
            return;
        }
        
        if (cleanSheet.isNullObject) {
            console.warn("[Archive] PR_Data_Clean not found - skipping");
            return;
        }
        
        // Get current period data from PR_Data_Clean
        const cleanRange = cleanSheet.getUsedRangeOrNullObject();
        cleanRange.load("values");
        await context.sync();
        
        if (cleanRange.isNullObject || !cleanRange.values || cleanRange.values.length < 2) {
            console.warn("[Archive] PR_Data_Clean is empty - skipping archive summary update");
            return;
        }
        
        // Calculate current period summary
        const cleanHeaders = (cleanRange.values[0] || []).map(h => String(h || "").toLowerCase().trim());
        const cleanData = cleanRange.values.slice(1);
        
        // Find amount column
        const amountIdx = cleanHeaders.findIndex(h => h.includes("amount"));
        // Find employee column
        const employeeIdx = cleanHeaders.findIndex(h => h.includes("employee"));
        // Find payroll date column
        const dateIdx = cleanHeaders.findIndex(h => 
            h.includes("payroll") && h.includes("date") || h.includes("pay period") || h === "date"
        );
        
        // Calculate totals
        let totalAmount = 0;
        const uniqueEmployees = new Set();
        let periodDate = getPayrollDateValue() || "";
        
        cleanData.forEach(row => {
            if (amountIdx >= 0) totalAmount += Number(row[amountIdx]) || 0;
            if (employeeIdx >= 0 && row[employeeIdx]) uniqueEmployees.add(String(row[employeeIdx]).trim());
            if (dateIdx >= 0 && row[dateIdx] && !periodDate) periodDate = String(row[dateIdx]);
        });
        
        const employeeCount = uniqueEmployees.size;
        
        console.log(`[Archive] Current period: Date=${periodDate}, Total=${totalAmount}, Employees=${employeeCount}`);
        
        // Get archive summary data
        const archiveRange = archiveSheet.getUsedRangeOrNullObject();
        archiveRange.load("values,rowCount");
        await context.sync();
        
        let archiveHeaders = [];
        let archiveData = [];
        
        if (!archiveRange.isNullObject && archiveRange.values && archiveRange.values.length > 0) {
            archiveHeaders = archiveRange.values[0];
            archiveData = archiveRange.values.slice(1);
        }
        
        // If no headers, create default structure
        if (archiveHeaders.length === 0) {
            archiveHeaders = ["Pay Period", "Total Payroll", "Employee Count", "Archived Date"];
            archiveSheet.getRange("A1:D1").values = [archiveHeaders];
            await context.sync();
        }
        
        // Find column indices in archive
        const archiveHeadersLower = archiveHeaders.map(h => String(h || "").toLowerCase().trim());
        const archiveDateIdx = archiveHeadersLower.findIndex(h => h.includes("pay period") || h.includes("period") || h === "date");
        const archiveTotalIdx = archiveHeadersLower.findIndex(h => h.includes("total"));
        const archiveCountIdx = archiveHeadersLower.findIndex(h => h.includes("employee") || h.includes("count"));
        const archiveTimestampIdx = archiveHeadersLower.findIndex(h => h.includes("archived"));
        
        // Create new row for current period
        const newRow = new Array(archiveHeaders.length).fill("");
        if (archiveDateIdx >= 0) newRow[archiveDateIdx] = periodDate;
        if (archiveTotalIdx >= 0) newRow[archiveTotalIdx] = totalAmount;
        if (archiveCountIdx >= 0) newRow[archiveCountIdx] = employeeCount;
        if (archiveTimestampIdx >= 0) newRow[archiveTimestampIdx] = new Date().toISOString().split("T")[0];
        
        // Keep only 4 most recent periods (we're adding 1 new = 5 total)
        // Sort by date descending, keep newest 4
        if (archiveData.length >= 5) {
            // Remove oldest row(s) to make room
            const rowsToKeep = 4;
            archiveData = archiveData.slice(0, rowsToKeep);
            console.log(`[Archive] Trimmed archive to ${rowsToKeep} periods, adding current`);
        }
        
        // Add new period at the top (most recent first)
        archiveData.unshift(newRow);
        
        // Clear existing data and rewrite
        const dataStartRow = 2; // Row 2 (after headers)
        const dataEndRow = dataStartRow + 5; // Max 6 rows of data
        
        // Clear old data range
        const clearRange = archiveSheet.getRange(`A${dataStartRow}:${String.fromCharCode(64 + archiveHeaders.length)}${dataEndRow}`);
        clearRange.clear(Excel.ClearApplyTo.contents);
        await context.sync();
        
        // Write new data
        if (archiveData.length > 0) {
            const writeRange = archiveSheet.getRange(
                `A${dataStartRow}:${String.fromCharCode(64 + archiveHeaders.length)}${dataStartRow + archiveData.length - 1}`
            );
            writeRange.values = archiveData;
            await context.sync();
        }
        
        console.log(`[Archive] Archive summary updated with ${archiveData.length} periods`);
    });
}

/**
 * Step 3: Clear working data from payroll sheets (data only, not headers)
 */
async function clearWorkingData() {
    const sheetsToClear = [
        SHEET_NAMES.DATA,
        SHEET_NAMES.DATA_CLEAN,
        SHEET_NAMES.EXPENSE_REVIEW,
        SHEET_NAMES.JE_DRAFT
    ];
    
    await Excel.run(async (context) => {
        for (const sheetName of sheetsToClear) {
            const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
            sheet.load("isNullObject");
            await context.sync();
            
            if (sheet.isNullObject) {
                console.log(`[Archive] Sheet not found: ${sheetName}`);
                continue;
            }
            
            const usedRange = sheet.getUsedRangeOrNullObject();
            usedRange.load("rowCount,columnCount,address");
            await context.sync();
            
            if (usedRange.isNullObject || usedRange.rowCount <= 1) {
                console.log(`[Archive] Sheet empty or headers only: ${sheetName}`);
                continue;
            }
            
            // Clear data rows (row 2 onwards), keep headers (row 1)
            const dataRange = sheet.getRange(`A2:${String.fromCharCode(64 + usedRange.columnCount)}${usedRange.rowCount}`);
            dataRange.clear(Excel.ClearApplyTo.contents);
            
            // Also clear any charts on expense review
            if (sheetName === SHEET_NAMES.EXPENSE_REVIEW) {
                const charts = sheet.charts;
                charts.load("items");
                await context.sync();
                for (let i = charts.items.length - 1; i >= 0; i--) {
                    charts.items[i].delete();
                }
            }
            
            await context.sync();
            console.log(`[Archive] Cleared data from: ${sheetName}`);
        }
    });
}

/**
 * Step 4: Clear non-permanent step notes from SS_PF_Config
 */
async function clearNonPermanentNotes() {
    await Excel.run(async (context) => {
        const table = await getConfigTable(context);
        if (!table) {
            console.warn("[Archive] Config table not found");
            return;
        }
        
        const body = table.getDataBodyRange();
        body.load("values,rowCount");
        await context.sync();
        
        const rows = body.values || [];
        let clearedCount = 0;
        
        // Find note fields to clear
        const noteFields = Object.values(STEP_NOTES_FIELDS).map(f => f.note);
        
        for (let i = 0; i < rows.length; i++) {
            const fieldName = String(rows[i][CONFIG_COLUMNS.FIELD] || "").trim();
            const permanentFlag = String(rows[i][CONFIG_COLUMNS.PERMANENT] || "").toUpperCase();
            
            // Check if this is a note field and not permanent
            if (noteFields.includes(fieldName) && permanentFlag !== "Y") {
                body.getCell(i, CONFIG_COLUMNS.VALUE).values = [[""]];
                clearedCount++;
            }
        }
        
        await context.sync();
        console.log(`[Archive] Cleared ${clearedCount} non-permanent notes`);
    });
}

/**
 * Step 5: Reset non-permanent config values (run-specific settings)
 */
async function resetNonPermanentConfig() {
    // Fields to reset after each archive (new + legacy names for migration)
    const fieldsToReset = [
        "PR_Payroll_Date",
        "PR_Accounting_Period",
        "PR_Journal_Entry_ID",
        // Legacy field names (for migration)
        "Payroll_Date",
        "Accounting_Period",
        "Journal_Entry_ID",
        "JE_Transaction_ID",
        // Sign-off dates and completion flags
        ...Object.values(STEP_NOTES_FIELDS).map(f => f.signOff),
        ...Object.values(STEP_NOTES_FIELDS).map(f => f.reviewer),
        ...Object.values(STEP_COMPLETE_FIELDS)
    ];
    
    await Excel.run(async (context) => {
        const table = await getConfigTable(context);
        if (!table) {
            console.warn("[Archive] Config table not found");
            return;
        }
        
        const body = table.getDataBodyRange();
        body.load("values,rowCount");
        await context.sync();
        
        const rows = body.values || [];
        let resetCount = 0;
        
        for (let i = 0; i < rows.length; i++) {
            const fieldName = String(rows[i][CONFIG_COLUMNS.FIELD] || "").trim();
            const permanentFlag = String(rows[i][CONFIG_COLUMNS.PERMANENT] || "").toUpperCase();
            
            // Check if this field should be reset and is not permanent
            const shouldReset = fieldsToReset.some(f => 
                normalizeFieldName(f) === normalizeFieldName(fieldName)
            );
            
            if (shouldReset && permanentFlag !== "Y") {
                body.getCell(i, CONFIG_COLUMNS.VALUE).values = [[""]];
                resetCount++;
            }
        }
        
        await context.sync();
        console.log(`[Archive] Reset ${resetCount} non-permanent config values`);
        
        // Clear local state
        Object.keys(configState.values).forEach(key => {
            if (fieldsToReset.some(f => normalizeFieldName(f) === normalizeFieldName(key))) {
                configState.values[key] = "";
            }
        });
    });
}

async function runJournalSummary() {
    if (!hasExcelRuntime()) {
        window.alert("Excel runtime is unavailable.");
        return;
    }
    journalState.loading = true;
    journalState.lastError = null;
    markJeSaveState(false);
    renderApp();
    try {
        const totals = await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(SHEET_NAMES.JE_DRAFT);
            const range = sheet.getUsedRangeOrNullObject();
            range.load("values");
            await context.sync();
            const values = range.isNullObject ? [] : range.values || [];
            if (!values.length) {
                throw new Error(`${SHEET_NAMES.JE_DRAFT} is empty.`);
            }
            const headers = (values[0] || []).map((h) => normalizeHeader(h));
            const debitIdx = headers.findIndex((h) => h.includes("debit"));
            const creditIdx = headers.findIndex((h) => h.includes("credit"));
            if (debitIdx === -1 || creditIdx === -1) {
                throw new Error("Debit/Credit columns not found in JE Draft.");
            }
            let debitTotal = 0;
            let creditTotal = 0;
            values.slice(1).forEach((row) => {
                debitTotal += Number(row[debitIdx]) || 0;
                creditTotal += Number(row[creditIdx]) || 0;
            });
            return { debitTotal, creditTotal, difference: creditTotal - debitTotal };
        });
        Object.assign(journalState, totals, { lastError: null });
    } catch (error) {
        console.warn("JE summary:", error);
        journalState.lastError = error?.message || "Unable to calculate journal totals.";
        journalState.debitTotal = null;
        journalState.creditTotal = null;
        journalState.difference = null;
    } finally {
        journalState.loading = false;
        renderApp();
    }
}

async function saveJournalSummary() {
    try {
        const debit = Number.isFinite(Number(journalState.debitTotal)) ? journalState.debitTotal : "";
        const credit = Number.isFinite(Number(journalState.creditTotal)) ? journalState.creditTotal : "";
        const diff = Number.isFinite(Number(journalState.difference)) ? journalState.difference : "";
        await Promise.all([
            writeConfigValue(JE_TOTAL_DEBIT_FIELD, String(debit)),
            writeConfigValue(JE_TOTAL_CREDIT_FIELD, String(credit)),
            writeConfigValue(JE_DIFFERENCE_FIELD, String(diff))
        ]);
        markJeSaveState(true);
    } catch (error) {
        console.error("JE save:", error);
    }
}

async function createJournalDraft() {
    if (!hasExcelRuntime()) {
        window.alert("Excel runtime is unavailable.");
        return;
    }
    
    journalState.loading = true;
    journalState.lastError = null;
    renderApp();
    
    try {
        await Excel.run(async (context) => {
            // 1. Read config values (Journal_Entry_ID and Payroll Date)
            let refNumber = "";
            let txnDate = "";
            
            const configTable = await getConfigTable(context);
            if (configTable) {
                const configBody = configTable.getDataBodyRange();
                configBody.load("values");
                await context.sync();
                
                const configRows = configBody.values || [];
                for (const row of configRows) {
                    const fieldName = String(row[CONFIG_COLUMNS.FIELD] || "").trim();
                    const fieldValue = row[CONFIG_COLUMNS.VALUE];
                    
                    // Journal Entry ID
                    if (fieldName === "Journal_Entry_ID" || fieldName === "JE_Transaction_ID" || fieldName === "PR_Journal_Entry_ID") {
                        refNumber = String(fieldValue || "").trim();
                    }
                    
                    // Payroll Date
                    if (PAYROLL_DATE_ALIASES.some(alias => 
                        fieldName === alias || normalizeFieldName(fieldName) === normalizeFieldName(alias)
                    )) {
                        txnDate = formatDateInput(fieldValue) || "";
                    }
                }
            }
            
            console.log("JE Generation - RefNumber:", refNumber, "TxnDate:", txnDate);
            
            // 2. Read PR_Data_Clean
            const cleanSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.DATA_CLEAN);
            cleanSheet.load("isNullObject");
            await context.sync();
            
            if (cleanSheet.isNullObject) {
                throw new Error("PR_Data_Clean sheet not found. Run Validate & Reconcile first.");
            }
            
            const cleanRange = cleanSheet.getUsedRangeOrNullObject();
            cleanRange.load("values");
            await context.sync();
            
            if (cleanRange.isNullObject) {
                throw new Error("PR_Data_Clean is empty. Run Validate & Reconcile first.");
            }
            
            const cleanValues = cleanRange.values || [];
            if (cleanValues.length < 2) {
                throw new Error("PR_Data_Clean has no data rows.");
            }
            
            // 3. Parse headers and find column indices
            const headers = cleanValues[0].map(h => normalizeHeader(h));
            console.log("JE Generation - PR_Data_Clean headers:", headers);
            
            const colIdx = {
                accountNumber: headers.findIndex(h => h.includes("account") && (h.includes("number") || h.includes("#"))),
                accountName: headers.findIndex(h => h.includes("account") && h.includes("name")),
                amount: headers.findIndex(h => h.includes("amount")),
                department: pickDepartmentIndex(headers),
                payrollCategory: headers.findIndex(h => h.includes("payroll") && h.includes("category")),
                employee: headers.findIndex(h => h.includes("employee"))
            };
            
            console.log("JE Generation - Column indices:", colIdx);
            
            if (colIdx.amount === -1) {
                throw new Error("Amount column not found in PR_Data_Clean.");
            }
            
            // 4. Aggregate by Account Number + Department
            const aggregated = new Map(); // Key: "AccountNumber|Department"
            
            for (let i = 1; i < cleanValues.length; i++) {
                const row = cleanValues[i];
                
                const accountNumber = colIdx.accountNumber >= 0 ? String(row[colIdx.accountNumber] || "").trim() : "";
                const accountName = colIdx.accountName >= 0 ? String(row[colIdx.accountName] || "").trim() : "";
                const amount = Number(row[colIdx.amount]) || 0;
                const department = colIdx.department >= 0 ? String(row[colIdx.department] || "").trim() : "";
                
                // Skip zero amounts
                if (amount === 0) continue;
                
                // Create aggregation key
                const key = `${accountNumber}|${department}`;
                
                if (aggregated.has(key)) {
                    // Add to existing entry
                    const entry = aggregated.get(key);
                    entry.amount += amount;
                } else {
                    // Create new entry
                    aggregated.set(key, {
                        accountNumber,
                        accountName,
                        department,
                        amount
                    });
                }
            }
            
            console.log("JE Generation - Aggregated into", aggregated.size, "unique Account+Department combinations");
            
            // 5. Build journal entry rows from aggregated data
            const jeHeaders = [
                "RefNumber",
                "TxnDate", 
                "Account Number",
                "Account Name",
                "LineAmount",
                "Debit",
                "Credit",
                "LineDesc",
                "Department"
            ];
            
            const jeRows = [jeHeaders];
            let totalDebit = 0;
            let totalCredit = 0;
            
            // Sort by account number, then department for cleaner output
            const sortedEntries = Array.from(aggregated.values()).sort((a, b) => {
                const acctCompare = String(a.accountNumber).localeCompare(String(b.accountNumber));
                if (acctCompare !== 0) return acctCompare;
                return String(a.department).localeCompare(String(b.department));
            });
            
            for (const entry of sortedEntries) {
                const { accountNumber, accountName, department, amount } = entry;
                
                // Calculate debit/credit from aggregated amount
                const debit = amount > 0 ? amount : 0;
                const credit = amount < 0 ? Math.abs(amount) : 0;
                
                // Build line description
                const lineDesc = [accountName, department].filter(Boolean).join(" - ");
                
                // Track totals
                totalDebit += debit;
                totalCredit += credit;
                
                jeRows.push([
                    refNumber,           // RefNumber
                    txnDate,             // TxnDate
                    accountNumber,       // Account Number
                    accountName,         // Account Name
                    amount,              // LineAmount (net)
                    debit || "",         // Debit (blank if 0)
                    credit || "",        // Credit (blank if 0)
                    lineDesc,            // LineDesc
                    department           // Department
                ]);
            }
            
            console.log("JE Generation - Built", jeRows.length - 1, "summarized journal lines");
            console.log("JE Generation - Total Debit:", totalDebit, "Total Credit:", totalCredit);
            
            // 5. Write to PR_JE_Draft
            let jeSheet = context.workbook.worksheets.getItemOrNullObject(SHEET_NAMES.JE_DRAFT);
            jeSheet.load("isNullObject");
            await context.sync();
            
            if (jeSheet.isNullObject) {
                // Create the sheet if it doesn't exist
                jeSheet = context.workbook.worksheets.add(SHEET_NAMES.JE_DRAFT);
                await context.sync();
            } else {
                // Clear existing content
                const usedRange = jeSheet.getUsedRangeOrNullObject();
                usedRange.load("address");
                await context.sync();
                if (!usedRange.isNullObject) {
                    usedRange.clear();
                    await context.sync();
                }
            }
            
            // Write the data
            const dataRange = jeSheet.getRangeByIndexes(0, 0, jeRows.length, jeHeaders.length);
            dataRange.values = jeRows;
            await context.sync();
            
            // Apply formatting using shared utilities
            try {
                const dataRowCount = jeRows.length - 1; // exclude header
                
                // Header formatting (9 columns: A-I) using shared style
                const headerRange = jeSheet.getRange("A1:I1");
                formatSheetHeaders(headerRange);
                
                // Date formatting for TxnDate (column B, index 1)
                if (dataRowCount > 0) {
                    formatDateColumn(jeSheet, 1, dataRowCount);
                    
                    // Currency formatting for amount columns (E=LineAmount, F=Debit, G=Credit)
                    formatCurrencyColumn(jeSheet, 4, dataRowCount);  // E: LineAmount
                    formatCurrencyColumn(jeSheet, 5, dataRowCount);  // F: Debit
                    formatCurrencyColumn(jeSheet, 6, dataRowCount);  // G: Credit
                }
                
                // Auto-fit columns
                jeSheet.getRange("A:I").format.autofitColumns();
                
                await context.sync();
            } catch (formatError) {
                console.warn("JE formatting error (non-critical):", formatError);
            }
            
            // Activate the sheet and select A1
            jeSheet.activate();
            jeSheet.getRange("A1").select();
            await context.sync();
            
            // Update journal state with totals
            journalState.debitTotal = totalDebit;
            journalState.creditTotal = totalCredit;
            journalState.difference = totalCredit - totalDebit;
        });
        
        journalState.loading = false;
        journalState.lastError = null;
        renderApp();
        
    } catch (error) {
        console.error("JE Generation failed:", error);
        journalState.loading = false;
        journalState.lastError = error.message || "Failed to generate journal entry.";
        renderApp();
    }
}

async function exportJournalDraft() {
    if (!hasExcelRuntime()) {
        window.alert("Excel runtime is unavailable.");
        return;
    }
    try {
        const { rows } = await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(SHEET_NAMES.JE_DRAFT);
            const range = sheet.getUsedRangeOrNullObject();
            range.load("values");
            await context.sync();
            const values = range.isNullObject ? [] : range.values || [];
            if (!values.length) {
                throw new Error(`${SHEET_NAMES.JE_DRAFT} is empty.`);
            }
            return { rows: values };
        });
        const csv = buildCsv(rows);
        downloadCsv(`pr-je-draft-${todayIso()}.csv`, csv);
    } catch (error) {
        console.warn("JE export:", error);
        window.alert("Unable to export the JE draft. Confirm the sheet has data.");
    }
}

