import { applyModuleTabVisibility, showAllSheets } from "../../Common/tab-visibility.js";
import { bindInstructionsButton } from "../../Common/instructions.js";
import { activateHomepageSheet, getHomepageConfig, renderAdaFab, removeAdaFab } from "../../Common/homepage-sheet.js";
import {
    HOME_ICON_SVG,
    MODULES_ICON_SVG,
    ARROW_LEFT_SVG,
    USERS_ICON_SVG,
    BOOK_ICON_SVG,
    ARROW_RIGHT_SVG,
    MENU_ICON_SVG,
    LOCK_CLOSED_SVG,
    LOCK_OPEN_SVG,
    CHECK_ICON_SVG,
    X_CIRCLE_SVG,
    CALCULATOR_ICON_SVG,
    LINK_ICON_SVG,
    SAVE_ICON_SVG,
    TABLE_ICON_SVG,
    DOWNLOAD_ICON_SVG,
    REFRESH_ICON_SVG,
    TRASH_ICON_SVG,
    getStepIconSvg
} from "../../Common/icons.js";
import { renderInlineNotes, renderSignoff, renderLabeledButton, updateLockButtonVisual, updateSaveButtonState, initSaveTracking } from "../../Common/notes-signoff.js";
import { canCompleteStep, showBlockedToast } from "../../Common/workflow-validation.js";
import { loadConfigFromTable, saveConfigValue, hasExcelRuntime } from "../../Common/gateway.js";
import { formatSheetHeaders, formatCurrencyColumn, formatNumberColumn, formatDateColumn, NUMBER_FORMATS } from "../../Common/sheet-formatting.js";

const MODULE_VERSION = "1.1.0";
const MODULE_KEY = "pto-accrual";
const MODULE_ALIAS_TOKENS = ["pto", "pto-accrual", "pto review", "accrual"];
const MODULE_NAME = "PTO Accrual";
const HERO_COPY =
    "Calculate your PTO liability, compare against last period, and generate a balanced journal entry—all without leaving Excel.";
const SELECTOR_URL = "../module-selector/index.html";
const LOADER_ID = "pf-loader-overlay";
const CONFIG_TABLES = ["SS_PF_Config"];
// PTO Config Fields - Pattern: PTO_{Descriptor}
const PTO_CONFIG_FIELDS = {
    payrollProvider: "PTO_Payroll_Provider",
    payrollDate: "PTO_Analysis_Date",
    accountingPeriod: "PTO_Accounting_Period",
    journalEntryId: "PTO_Journal_Entry_ID",
    companyName: "SS_Company_Name",           // Shared field
    accountingSoftware: "SS_Accounting_Software", // Shared field
    reviewerName: "PTO_Reviewer",
    validationDataBalance: "PTO_Validation_Data_Balance",
    validationCleanBalance: "PTO_Validation_Clean_Balance",
    validationDifference: "PTO_Validation_Difference",
    headcountRosterCount: "PTO_Headcount_Roster_Count",
    headcountPayrollCount: "PTO_Headcount_Payroll_Count",
    headcountDifference: "PTO_Headcount_Difference",
    journalDebitTotal: "PTO_JE_Debit_Total",
    journalCreditTotal: "PTO_JE_Credit_Total",
    journalDifference: "PTO_JE_Difference"
};
const HEADCOUNT_SKIP_NOTE = "User opted to skip the headcount review this period.";
// Step notes/sign-off fields - Pattern: PTO_{Type}_{StepName}
const STEP_CONFIG_FIELDS = {
    0: { note: "PTO_Notes_Config", reviewer: "PTO_Reviewer_Config", signOff: "PTO_SignOff_Config" },
    1: { note: "PTO_Notes_Import", reviewer: "PTO_Reviewer_Import", signOff: "PTO_SignOff_Import" },
    2: { note: "PTO_Notes_Headcount", reviewer: "PTO_Reviewer_Headcount", signOff: "PTO_SignOff_Headcount" },
    3: { note: "PTO_Notes_Validate", reviewer: "PTO_Reviewer_Validate", signOff: "PTO_SignOff_Validate" },
    4: { note: "PTO_Notes_Review", reviewer: "PTO_Reviewer_Review", signOff: "PTO_SignOff_Review" },
    5: { note: "PTO_Notes_JE", reviewer: "PTO_Reviewer_JE", signOff: "PTO_SignOff_JE" },
    6: { note: "PTO_Notes_Archive", reviewer: "PTO_Reviewer_Archive", signOff: "PTO_SignOff_Archive" }
};
const STEP_COMPLETE_FIELDS = {
    0: "PTO_Complete_Config",
    1: "PTO_Complete_Import",
    2: "PTO_Complete_Headcount",
    3: "PTO_Complete_Validate",
    4: "PTO_Complete_Review",
    5: "PTO_Complete_JE",
    6: "PTO_Complete_Archive"
};

const PTO_ACTIVITY_COLUMNS = [
    { key: "employeeId", header: "Employee ID" },
    { key: "employeeName", header: "Employee Name" },
    { key: "actionDate", header: "Action Date" },
    { key: "actionType", header: "Action" },
    { key: "hours", header: "Hours" },
    { key: "notes", header: "Notes" },
    { key: "source", header: "Source" }
];

// PTO_Analysis columns - combines enriched data + analysis view
const PTO_ANALYSIS_COLUMNS = [
    { key: "analysisDate", header: "Analysis Date" },
    { key: "employeeName", header: "Employee Name" },
    { key: "department", header: "Department" },
    { key: "payRate", header: "Pay Rate" },
    { key: "accrualRate", header: "Accrual Rate" },
    { key: "carryOver", header: "Carry Over" },
    { key: "balance", header: "Balance" },
    { key: "liabilityAmount", header: "Liability Amount" }
];

const PTO_EXPENSE_COLUMNS = [
    { key: "department", header: "Department" },
    { key: "currentPeriod", header: "Current Period" },
    { key: "priorPeriod", header: "Prior Period" },
    { key: "variance", header: "Variance" },
    { key: "comment", header: "Comment" }
];

const PTO_JOURNAL_COLUMNS = [
    { key: "account", header: "Account" },
    { key: "description", header: "Description" },
    { key: "debit", header: "Debit" },
    { key: "credit", header: "Credit" },
    { key: "reference", header: "Reference" }
];

const WORKFLOW_STEPS = [
    {
        id: 0,
        title: "Configuration",
        summary: "Set the analysis date, accounting period, and review details for this run.",
        description: "Complete this step first to ensure all downstream calculations use the correct period settings.",
        actionLabel: "Configure Workbook",
        secondaryAction: { sheet: "SS_PF_Config", label: "Open Config Sheet" }
    },
    {
        id: 1,
        title: "Import PTO Data",
        summary: "Pull your latest PTO export from payroll and paste it into PTO_Data.",
        description: "Open your payroll provider, download the PTO report, and paste the data into the PTO_Data tab.",
        actionLabel: "Import Sample Data",
        secondaryAction: { sheet: "PTO_Data", label: "Open Data Sheet" }
    },
    {
        id: 2,
        title: "Headcount Review",
        summary: "Quick check to make sure your roster matches your PTO data.",
        description: "Compare employees in PTO_Data against your employee roster to catch any discrepancies.",
        actionLabel: "Open Headcount Review",
        secondaryAction: { sheet: "SS_Employee_Roster", label: "Open Sheet" }
    },
    {
        id: 3,
        title: "Data Quality Review",
        summary: "Scan your PTO data for potential errors before crunching numbers.",
        description: "Identify negative balances, overdrawn accounts, and other anomalies that might need attention.",
        actionLabel: "Click to Run Quality Check"
    },
    {
        id: 4,
        title: "PTO Accrual Review",
        summary: "Review the calculated liability for each employee and compare to last period.",
        description: "The analysis enriches your PTO data with pay rates and department info, then calculates the liability.",
        actionLabel: "Click to Perform Review"
    },
    {
        id: 5,
        title: "Journal Entry Prep",
        summary: "Generate a balanced journal entry, run validation checks, and export when ready.",
        description: "Build the JE from your PTO data, verify debits equal credits, and export for upload to your accounting system.",
        actionLabel: "Open Journal Draft",
        secondaryAction: { sheet: "PTO_JE_Draft", label: "Open Sheet" }
    },
    {
        id: 6,
        title: "Archive & Reset",
        summary: "Save this period's results and prepare for the next cycle.",
        description: "Archive the current analysis so it becomes the 'prior period' for your next review.",
        actionLabel: "Archive Run"
    }
];

const WORKBOOK_SHEETS = [
    {
        name: "PTO_Instructions",
        description: "Overview of the PTO workflow",
        position: "beginning",
        onCreate: async (sheet, context) => {
            const rows = [
                ["Prairie Forge PTO Accrual", `Version ${MODULE_VERSION}`],
                ["Step 0 — Configuration", "Set your analysis date, accounting period, and JE reference."],
                ["Step 1 — Import PTO Data", "Pull data from payroll and paste into PTO_Data."],
                ["Step 2 — Headcount Review", "Verify employees match between roster and PTO data."],
                ["Step 3 — Data Quality", "Scan for balance issues and anomalies."],
                ["Step 4 — Accrual Review", "Calculate liabilities and compare to prior period."],
                ["Step 5 — Journal Entry", "Generate JE, validate, and export for upload."],
                ["Step 6 — Archive", "Save results for next period comparison."]
            ];
            const target = sheet.getRangeByIndexes(0, 0, rows.length, 2);
            target.values = rows;
            target.format.autofitColumns();

            const header = sheet.getRange("A1:B1");
            header.merge();
            header.format.font.bold = true;
            header.format.font.size = 18;
            header.format.font.color = "#111827";

            const info = sheet.getRange("A2:B8");
            info.format.fill.color = "#f1f5f9";
            await context.sync();
        }
    },
    // PTO_Config removed - consolidated into SS_PF_Config
    { name: "SS_PF_Config", description: "Prairie Forge shared configuration (all modules)" },
    { name: "PTO_Rates", description: "Accrual rate definitions" },
    { name: "PTO_Data", description: "Raw PTO transactions" },
    { name: "PTO_Analysis", description: "Calculated balances" },
    { name: "PTO_ExpenseReview", description: "Expense review workspace" },
    { name: "PTO_JE_Draft", description: "Journal entry prep" },
    { name: "PTO_Archive", description: "Archive register" },
    { name: "SS_Employee_Roster", description: "Centralized employee roster" }
];

const TABLE_DEFINITIONS = [
    // PTO_Config table removed - all config consolidated into SS_PF_Config
    {
        sheetName: "PTO_Rates",
        tableName: "PTORates",
        headers: ["Tier", "Description", "Hours_Per_Period", "Max_Carryover", "Carryover_Reset"],
        sampleRows: [
            ["Standard", "0-4 years tenure", 6.67, 80, "Dec 31"],
            ["Senior", "5-9 years tenure", 10, 120, "Dec 31"],
            ["Executive", "10+ years", 13.33, 160, "Dec 31"]
        ]
    },
    {
        sheetName: "PTO_Archive",
        tableName: "PTOArchiveLog",
        headers: ["Timestamp", "Action", "Notes"],
        sampleRows: []
    }
];

const SAMPLE_PTO_ACTIVITY = [
    { employeeId: "AC1001", employeeName: "Kelly Mendez", actionDate: "2025-01-15", actionType: "Accrual", hours: 6.67, notes: "Monthly accrual", source: "Payroll" },
    { employeeId: "AC1001", employeeName: "Kelly Mendez", actionDate: "2025-01-22", actionType: "Usage", hours: 8, notes: "Vacation day", source: "HRIS" },
    { employeeId: "AC2044", employeeName: "Justin Reid", actionDate: "2025-01-15", actionType: "Accrual", hours: 10, notes: "Senior tier", source: "Payroll" },
    { employeeId: "AC2044", employeeName: "Justin Reid", actionDate: "2025-01-28", actionType: "Usage", hours: 4, notes: "Doctor appointment", source: "HRIS" },
    { employeeId: "AC3020", employeeName: "Amelia Yates", actionDate: "2025-01-15", actionType: "Accrual", hours: 13.33, notes: "Executive tier", source: "Payroll" },
    { employeeId: "AC3020", employeeName: "Amelia Yates", actionDate: "2025-01-18", actionType: "Carryover", hours: 20, notes: "Year-end carryover", source: "Finance" }
];

// Sample data will be generated dynamically from PTO_Data via syncPtoAnalysis
const SAMPLE_ANALYSIS_DATA = [];

const SAMPLE_EXPENSE_REVIEW = [
    { department: "Operations", currentPeriod: 5400, priorPeriod: 5200, variance: 200, comment: "Increased carryover true-up" },
    { department: "Sales", currentPeriod: 2300, priorPeriod: 2600, variance: -300, comment: "Usage down vs. forecast" },
    { department: "Executive", currentPeriod: 4100, priorPeriod: 4100, variance: 0, comment: "Steady quarter" }
];

const SAMPLE_JOURNAL_LINES = [
    { account: "2100.100", description: "PTO Accrual Expense", debit: 11800, credit: 0, reference: "PTO-EXP-2025-01" },
    { account: "2190.250", description: "PTO Accrual Liability", debit: 0, credit: 11800, reference: "PTO-EXP-2025-01" }
];

const stepStatuses = WORKFLOW_STEPS.reduce((acc, step) => {
    acc[step.id] = "pending";
    return acc;
}, {});

const appState = {
    activeView: "home",
    activeStepId: null,
    focusedIndex: 0,
    stepStatuses
};

const configState = {
    loaded: false,
    steps: {},
    permanents: {},
    completes: {},
    values: {},
    overrides: {
        accountingPeriod: false,
        journalId: false
    }
};

let rootEl = null;
let loadingEl = null;
let pendingScrollIndex = null;
const pendingConfigWrites = new Map();
const headcountState = {
    skipAnalysis: false,
    roster: {
        rosterCount: null,
        payrollCount: null,
        difference: null,
        mismatches: []
    },
    loading: false,
    hasAnalyzed: false,
    lastError: null
};
const journalState = {
    debitTotal: null,
    creditTotal: null,
    difference: null,
    lineAmountSum: null,        // Sum of all line amounts (should be 0)
    analysisChangeTotal: null,  // Total Change from PTO_Analysis
    jeChangeTotal: null,        // Total Change captured in JE (expense lines only)
    loading: false,
    lastError: null,
    // Validation results with details
    validationRun: false,
    issues: []                  // [{check: "name", passed: false, detail: "explanation"}]
};
const dataQualityState = {
    hasRun: false,
    loading: false,
    acknowledged: false,       // User acknowledged issues and wants to proceed
    // Quality check results
    balanceIssues: [],         // [{name, issue, rowIndex}] - Negative balance or used more than available
    zeroBalances: [],          // [{name, rowIndex}]
    accrualOutliers: [],       // [{name, accrualRate, rowIndex}] - rates > 8 hrs/period
    // Summary counts
    totalIssues: 0,
    totalEmployees: 0
};

const analysisState = {
    cleanDataReady: false,
    employeeCount: 0,
    lastRun: null,
    loading: false,
    lastError: null,
    // Data quality tracking
    missingPayRates: [],      // [{name: "John Doe", rowIndex: 2}, ...]
    missingDepartments: [],   // [{name: "Jane Smith", rowIndex: 3}, ...]
    ignoredMissingPayRates: new Set(), // Names user chose to ignore
    // Data completeness check (PTO_Data vs PTO_Analysis sums)
    completenessCheck: {
        accrualRate: null,    // { match: true/false, ptoData: number, ptoAnalysis: number }
        carryOver: null,
        ytdAccrued: null,
        ytdUsed: null,
        balance: null
    }
};

async function init() {
    try {
        rootEl = document.getElementById("app");
        loadingEl = document.getElementById("loading");
        
        await ensureTabVisibility();
        await loadStepConfig();
        // Load shared config for fallback values (Company Name, Default Reviewer, etc.)
        if (window.PrairieForge?.loadSharedConfig) {
            await window.PrairieForge.loadSharedConfig();
        }
        
        // Activate module homepage on load
        const homepageConfig = getHomepageConfig(MODULE_KEY);
        await activateHomepageSheet(homepageConfig.sheetName, homepageConfig.title, homepageConfig.subtitle);
        
        if (loadingEl) loadingEl.remove();
        if (rootEl) rootEl.hidden = false;
        renderApp();
    } catch (error) {
        console.error("[PTO] Module initialization failed:", error);
        throw error;
    }
}

async function ensureTabVisibility() {
    // Apply prefix-based tab visibility
    // Shows PTO_* tabs, hides PR_* and SS_* tabs
    try {
        await applyModuleTabVisibility(MODULE_KEY);
        console.log(`[PTO] Tab visibility applied for ${MODULE_KEY}`);
    } catch (error) {
        console.warn("[PTO] Could not apply tab visibility:", error);
    }
}

async function loadStepConfig() {
    if (!hasExcelRuntime()) {
        configState.loaded = true;
        return;
    }
    try {
        // Load from module-specific config table (backwards compatibility)
        const moduleValues = await loadConfigFromTable(CONFIG_TABLES);
        
        // Load from shared config (SS_PF_Config) - takes precedence
        let sharedValues = {};
        if (window.PrairieForge?.loadSharedConfig) {
            await window.PrairieForge.loadSharedConfig();
            // Convert cached Map to object
            if (window.PrairieForge._sharedConfigCache) {
                window.PrairieForge._sharedConfigCache.forEach((value, key) => {
                    sharedValues[key] = value;
                });
            }
        }
        
        // Merge: module values first, then shared values override
        // Also map new naming convention to old field names
        const values = { ...moduleValues };
        
        // Map shared config fields to PTO field names (new + legacy names)
        const fieldMappings = {
            "SS_Default_Reviewer": PTO_CONFIG_FIELDS.reviewerName,
            "Default_Reviewer": PTO_CONFIG_FIELDS.reviewerName,
            "PTO_Reviewer": PTO_CONFIG_FIELDS.reviewerName,
            "SS_Company_Name": PTO_CONFIG_FIELDS.companyName,
            "Company_Name": PTO_CONFIG_FIELDS.companyName,
            "SS_Payroll_Provider": PTO_CONFIG_FIELDS.payrollProvider,
            "Payroll_Provider_Link": PTO_CONFIG_FIELDS.payrollProvider,
            "SS_Accounting_Software": PTO_CONFIG_FIELDS.accountingSoftware,
            "Accounting_Software_Link": PTO_CONFIG_FIELDS.accountingSoftware
        };
        
        // Apply shared values using mapped field names
        Object.entries(fieldMappings).forEach(([sharedField, ptoField]) => {
            if (sharedValues[sharedField] && !values[ptoField]) {
                values[ptoField] = sharedValues[sharedField];
            }
        });
        
        // Also apply any exact-match PTO_* fields from shared config
        Object.entries(sharedValues).forEach(([key, value]) => {
            if (key.startsWith("PTO_") && value) {
                values[key] = value;
            }
        });
        
        configState.permanents = await loadPermanentFlags();
        configState.values = values || {};
        configState.overrides.accountingPeriod = Boolean(values?.[PTO_CONFIG_FIELDS.accountingPeriod]);
        configState.overrides.journalId = Boolean(values?.[PTO_CONFIG_FIELDS.journalEntryId]);
        Object.entries(STEP_CONFIG_FIELDS).forEach(([stepId, fields]) => {
            configState.steps[stepId] = {
                notes: values[fields.note] ?? "",
                reviewer: values[fields.reviewer] ?? "",
                signOffDate: values[fields.signOff] ?? ""
            };
        });
        configState.completes = Object.entries(STEP_COMPLETE_FIELDS).reduce((acc, [stepId, field]) => {
            acc[stepId] = values[field] ?? "";
            return acc;
        }, {});
        configState.loaded = true;
    } catch (error) {
        console.warn("PTO: unable to load configuration fields", error);
        configState.loaded = true;
    }
}

async function loadPermanentFlags() {
    const permanents = {};
    if (!hasExcelRuntime()) return permanents;
    const fieldToStep = new Map();
    Object.entries(STEP_CONFIG_FIELDS).forEach(([stepId, fields]) => {
        if (fields.note) {
            fieldToStep.set(fields.note.trim(), Number(stepId));
        }
    });
    try {
        await Excel.run(async (context) => {
            const table = context.workbook.tables.getItemOrNullObject(CONFIG_TABLES[0]);
            await context.sync();
            if (table.isNullObject) return;
            const body = table.getDataBodyRange();
            const header = table.getHeaderRowRange();
            body.load("values");
            header.load("values");
            await context.sync();

            const headers = header.values[0] || [];
            const normalizedHeaders = headers.map((h) => String(h || "").trim().toLowerCase());
            const idx = {
                field: normalizedHeaders.findIndex((h) => h === "field" || h === "field name" || h === "setting"),
                permanent: normalizedHeaders.findIndex((h) => h === "permanent" || h === "persist")
            };
            if (idx.field === -1 || idx.permanent === -1) return;
            (body.values || []).forEach((row) => {
                const fieldName = String(row[idx.field] || "").trim();
                const stepId = fieldToStep.get(fieldName);
                if (stepId == null) return;
                const flag = parsePermanentFlag(row[idx.permanent]);
                permanents[stepId] = flag;
            });
        });
    } catch (error) {
        console.warn("PTO: unable to load permanent flags", error);
    }
    return permanents;
}

function renderApp() {
    if (!rootEl) return;
    const prevDisabled = appState.focusedIndex <= 0 ? "disabled" : "";
    const nextDisabled = appState.focusedIndex >= WORKFLOW_STEPS.length - 1 ? "disabled" : "";
    const isStepView = appState.activeView === "step" && appState.activeStepId != null;
    const isConfigView = appState.activeView === "config";
    const content = isConfigView
        ? renderConfigView()
        : isStepView
            ? renderStepView(appState.activeStepId)
            : `${renderHero()}${renderWorkflow()}`;
    rootEl.innerHTML = `
        <div class="pf-root">
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
                        </div>
                    </div>
                </div>
            </header>
            ${content}
            <footer class="pf-brand-footer">
                <div class="pf-brand-text">
                    <div class="pf-brand-label">prairie.forge</div>
                    <div class="pf-brand-meta">© Prairie Forge LLC, 2025. All rights reserved. Version ${MODULE_VERSION}</div>
                    <button type="button" class="pf-config-link" id="showConfigSheets">CONFIGURATION</button>
                </div>
            </footer>
        </div>
    `;
    // Determine if on home view
    const isHomeView = appState.activeView === "home" || (appState.activeView !== "step" && appState.activeView !== "config");
    
    // Mount info FAB with step-specific content (only on step/config views, not homepage)
    const infoFabElement = document.getElementById("pf-info-fab-pto");
    if (isHomeView) {
        // Remove info fab on homepage
        if (infoFabElement) infoFabElement.remove();
    } else if (window.PrairieForge?.mountInfoFab) {
        const infoConfig = getStepInfoConfig(appState.activeStepId);
        PrairieForge.mountInfoFab({ 
            title: infoConfig.title, 
            content: infoConfig.content, 
            buttonId: "pf-info-fab-pto" 
        });
    }
    
    bindInteractions();
    scrollFocusedIntoView();
    
    // Show/hide Ada FAB based on view
    if (isHomeView) {
        renderAdaFab();
    } else {
        removeAdaFab();
    }
}

/**
 * Get step-specific info panel configuration
 */
function getStepInfoConfig(stepId) {
    switch (stepId) {
        case 0:
            return {
                title: "Configuration",
                content: `
                    <div class="pf-info-section">
                        <h4>🎯 What This Step Does</h4>
                        <p>Sets up the key parameters for your PTO review. Complete this before importing data.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>📋 Key Fields</h4>
                        <ul>
                            <li><strong>Analysis Date</strong> — The period-end date (e.g., 11/30/2024)</li>
                            <li><strong>Accounting Period</strong> — Shows up in your JE description</li>
                            <li><strong>Journal Entry ID</strong> — Reference number for your accounting system</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>💡 Tip</h4>
                        <p>The accounting period and JE ID auto-generate based on your analysis date, but you can override them if needed.</p>
                    </div>
                `
            };
        case 1:
            return {
                title: "Data Import",
                content: `
                    <div class="pf-info-section">
                        <h4>🎯 What This Step Does</h4>
                        <p>Gets your PTO data into the workbook. Pull a report from your payroll provider and paste it into PTO_Data.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>📋 Required Columns</h4>
                        <p>Your payroll export should include:</p>
                        <ul>
                            <li><strong>Employee Name</strong> — Full name (used to match against roster)</li>
                            <li><strong>Accrual Rate</strong> — Hours accrued per pay period</li>
                            <li><strong>Carry Over</strong> — Hours carried from prior year</li>
                            <li><strong>YTD Accrued</strong> — Total hours accrued this year</li>
                            <li><strong>YTD Used</strong> — Total hours used this year</li>
                            <li><strong>Balance</strong> — Current available hours</li>
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
                        <p>Compares employee counts between your roster and PTO data to catch discrepancies early.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>📊 Data Sources</h4>
                        <ul>
                            <li><strong>SS_Employee_Roster</strong> — Your centralized employee list</li>
                            <li><strong>PTO_Data</strong> — The data you just imported</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>🔍 What to Look For</h4>
                        <ul>
                            <li><strong>In Roster, Not in PTO</strong> — May need to add PTO records</li>
                            <li><strong>In PTO, Not in Roster</strong> — Could be terminated employees</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>💡 Tip</h4>
                        <p>If discrepancies are expected (e.g., contractors without PTO), you can skip this check.</p>
                    </div>
                `
            };
        case 3:
            return {
                title: "Data Quality Review",
                content: `
                    <div class="pf-info-section">
                        <h4>🎯 What This Step Does</h4>
                        <p>Scans your PTO data for anomalies that could cause problems later in the process.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>⚠️ Balance Issues (Critical)</h4>
                        <p>Flags when:</p>
                        <ul>
                            <li><strong>Negative Balance</strong> — Balance is less than zero</li>
                            <li><strong>Overdrawn</strong> — Used more than available (YTD Used > Carry Over + YTD Accrued)</li>
                        </ul>
                        <p class="pf-info-note">Usually indicates missing accrual entries or data errors in payroll.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>📊 High Accrual Rates (Warning)</h4>
                        <p>Employees with Accrual Rate > 8 hours/period may have data entry errors.</p>
                        <p class="pf-info-note">Most bi-weekly accruals are 3-6 hours.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>💡 Tip</h4>
                        <p>You can acknowledge issues and proceed, but it's best to fix them in your source system first.</p>
                    </div>
                `
            };
        case 4:
            return {
                title: "PTO Accrual Review",
                content: `
                    <div class="pf-info-section">
                        <h4>🎯 What This Step Does</h4>
                        <p>Calculates the PTO liability for each employee and compares it to last period.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>📊 Data Sources</h4>
                        <ul>
                            <li><strong>PTO_Data</strong> — Your imported PTO balances</li>
                            <li><strong>SS_Employee_Roster</strong> — Department assignments</li>
                            <li><strong>PR_Archive_Summary</strong> — Pay rates from payroll history</li>
                            <li><strong>PTO_Archive_Summary</strong> — Last period's liability (for comparison)</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>💰 How Liability is Calculated</h4>
                        <div class="pf-info-formula">
                            Liability = Balance (hours) × Hourly Rate
                        </div>
                        <p class="pf-info-note">Hourly rate comes from Regular Earnings ÷ 80 hours in your payroll history.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>📈 How Change is Calculated</h4>
                        <div class="pf-info-formula">
                            Change = Current Liability − Prior Period Liability
                        </div>
                        <ul>
                            <li><span style="color: #30d158;">Positive</span> = Liability went up (book expense)</li>
                            <li><span style="color: #ff453a;">Negative</span> = Liability went down (reverse expense)</li>
                        </ul>
                    </div>
                `
            };
        case 5:
            return {
                title: "Journal Entry Prep",
                content: `
                    <div class="pf-info-section">
                        <h4>🎯 What This Step Does</h4>
                        <p>Generates a balanced journal entry from your PTO analysis, ready for upload to your accounting system.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>📝 How the JE Works</h4>
                        <p>Groups the <strong>Change</strong> amounts by department:</p>
                        <ul>
                            <li><span style="color: #30d158;">Positive Change</span> → Debit expense account</li>
                            <li><span style="color: #ff453a;">Negative Change</span> → Credit expense account</li>
                        </ul>
                        <p>The offset always goes to <strong>21540</strong> (Accrued PTO liability).</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>🏢 Department → Account Mapping</h4>
                        <table style="width:100%; font-size: 12px; margin-top: 8px;">
                            <tr><td>General & Admin</td><td style="text-align:right">64110</td></tr>
                            <tr><td>R&D</td><td style="text-align:right">62110</td></tr>
                            <tr><td>Marketing</td><td style="text-align:right">61610</td></tr>
                            <tr><td>Sales & Marketing</td><td style="text-align:right">61110</td></tr>
                            <tr><td>COGS Onboarding</td><td style="text-align:right">53110</td></tr>
                            <tr><td>COGS Prof. Services</td><td style="text-align:right">56110</td></tr>
                            <tr><td>COGS Support</td><td style="text-align:right">52110</td></tr>
                            <tr><td>Client Success</td><td style="text-align:right">61811</td></tr>
                        </table>
                    </div>
                    <div class="pf-info-section">
                        <h4>✅ Validation Checks</h4>
                        <ul>
                            <li><strong>Debits = Credits</strong> — Entry must balance</li>
                            <li><strong>Line Amounts = $0</strong> — Net change must be zero</li>
                            <li><strong>JE Matches Analysis</strong> — Totals tie back to your data</li>
                        </ul>
                    </div>
                `
            };
        case 6:
            return {
                title: "Archive & Reset",
                content: `
                    <div class="pf-info-section">
                        <h4>🎯 What This Step Does</h4>
                        <p>Saves this period's results so they become the "prior period" for your next review.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>📦 What Gets Saved</h4>
                        <ul>
                            <li><strong>PTO_Archive_Summary</strong> — Employee name, liability amount, and analysis date</li>
                            <li>This data is used to calculate the "Change" column next period</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>⚠️ Important</h4>
                        <p>Only the <strong>most recent period</strong> is kept in the archive. Running archive again will overwrite the previous data.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>💡 Tip</h4>
                        <p>Make sure your JE has been uploaded to your accounting system before archiving.</p>
                    </div>
                `
            };
        default:
            return {
                title: "PTO Accrual",
                content: `
                    <div class="pf-info-section">
                        <h4>👋 Welcome to PTO Accrual</h4>
                        <p>This module helps you calculate PTO liabilities and generate journal entries each period.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>📋 Workflow Overview</h4>
                        <ol style="margin: 8px 0; padding-left: 20px;">
                            <li>Configure your period settings</li>
                            <li>Import PTO data from payroll</li>
                            <li>Review headcount alignment</li>
                            <li>Check data quality</li>
                            <li>Review calculated liabilities</li>
                            <li>Generate and export journal entry</li>
                            <li>Archive for next period</li>
                        </ol>
                    </div>
                    <div class="pf-info-section">
                        <p>Click a step card to get started, or tap the <strong>ⓘ</strong> button on any step for detailed guidance.</p>
                    </div>
                `
            };
    }
}

function renderHero() {
    return `
        <section class="pf-hero" id="pf-hero">
            <h2 class="pf-hero-title">PTO Accrual</h2>
            <p class="pf-hero-copy">${HERO_COPY}</p>
        </section>
    `;
}

function renderWorkflow() {
    return `
        <section class="pf-step-guide">
            <div class="pf-step-grid">
                ${WORKFLOW_STEPS.map((step, index) => renderStepCard(step, index)).join("")}
            </div>
        </section>
    `;
}

function renderStepCard(step, index) {
    const status = appState.stepStatuses[step.id] || "pending";
    const isActive =
        appState.activeView === "step" && appState.focusedIndex === index ? "pf-step-card--active" : "";
    const icon = getStepIconSvg(getStepType(step.id));
    return `
        <article class="pf-step-card pf-clickable ${isActive}" data-step-card data-step-index="${index}" data-step-id="${step.id}">
            <p class="pf-step-index">Step ${step.id}</p>
            <h3 class="pf-step-title">${icon ? `${icon}` : ""}${step.title}</h3>
        </article>
    `;
}

function renderArchiveStep(detail) {
    const completionItems = WORKFLOW_STEPS.filter((step) => step.id !== 6).map((step) => ({
        id: step.id,
        title: step.title,
        complete: isStepComplete(step.id)
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
        </section>
        <section class="pf-step-guide">
            ${statusList}
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Archive & Reset</h3>
                    <p class="pf-config-subtext">Only enabled when all steps above are complete.</p>
                </div>
                <div class="pf-pill-row pf-config-actions">
                    <button type="button" class="pf-pill-btn" id="archive-run-btn" ${allComplete ? "" : "disabled"}>Archive</button>
                </div>
            </article>
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
    const payrollDate = formatDateInput(getConfigValue(PTO_CONFIG_FIELDS.payrollDate));
    const accountingPeriod = formatDateInput(getConfigValue(PTO_CONFIG_FIELDS.accountingPeriod));
    const journalEntryId = getConfigValue(PTO_CONFIG_FIELDS.journalEntryId);
    const accountingLink = getConfigValue(PTO_CONFIG_FIELDS.accountingSoftware);
    const payrollLink = getConfigValue(PTO_CONFIG_FIELDS.payrollProvider);
    const companyName = getConfigValue(PTO_CONFIG_FIELDS.companyName);
    const userName = getConfigValue(PTO_CONFIG_FIELDS.reviewerName);
    const stepFields = getStepConfig(0);
    const notesPermanent = Boolean(configState.permanents[0]);
    const isStepComplete = Boolean(parseBooleanFlag(configState.completes[0]) || stepFields.signOffDate);
    const stepReviewer = getReviewerWithFallback(stepFields?.reviewer);
    const stepSignOff = stepFields?.signOffDate || "";

    return `
        <section class="pf-hero" id="pf-config-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step 0</p>
            <h2 class="pf-hero-title">Configuration Setup</h2>
            <p class="pf-hero-copy">Make quick adjustments before every PTO run.</p>
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
                        <span>PTO Analysis Date</span>
                        <input type="date" id="config-payroll-date" value="${escapeHtml(payrollDate)}">
                    </label>
                    <label class="pf-config-field">
                        <span>Accounting Period</span>
                        <input type="text" id="config-accounting-period" value="${escapeHtml(accountingPeriod)}" placeholder="Nov 2025">
                    </label>
                    <label class="pf-config-field">
                        <span>Journal Entry ID</span>
                        <input type="text" id="config-journal-id" value="${escapeHtml(journalEntryId)}" placeholder="PTO-AUTO-YYYY-MM-DD">
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
            ${renderInlineNotes({
                textareaId: "config-notes",
                value: stepFields.notes || "",
                permanentId: "config-notes-lock",
                isPermanent: notesPermanent,
                hintId: "",
                saveButtonId: "config-notes-save"
            })}
            ${renderSignoff({
                reviewerInputId: "config-reviewer",
                reviewerValue: stepReviewer,
                signoffInputId: "config-signoff-date",
                signoffValue: stepSignOff,
                isComplete: isStepComplete,
                saveButtonId: "config-signoff-save",
                completeButtonId: "config-signoff-toggle"
            })}
        </section>
    `;
}

function renderImportStep(detail) {
    const stepFields = getStepConfig(1);
    const notesPermanent = Boolean(configState.permanents[1]);
    const stepReviewer = getReviewerWithFallback(stepFields?.reviewer);
    const stepSignOff = stepFields?.signOffDate || "";
    const stepComplete = Boolean(parseBooleanStrict(configState.completes[1]) || stepSignOff);
    const providerLink = getConfigValue(PTO_CONFIG_FIELDS.payrollProvider);
    
    return `
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step ${detail.id}</p>
            <h2 class="pf-hero-title">${escapeHtml(detail.title)}</h2>
            <p class="pf-hero-copy">${escapeHtml(detail.summary || "")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Payroll Report</h3>
                    <p class="pf-config-subtext">Access your payroll provider to download the latest PTO export, then paste into PTO_Data.</p>
                </div>
                <div class="pf-signoff-action">
                    ${renderLabeledButton(
                        providerLink 
                            ? `<a href="${escapeHtml(providerLink)}" target="_blank" rel="noopener noreferrer" class="pf-action-toggle pf-clickable" title="Open payroll provider">${LINK_ICON_SVG}</a>`
                            : `<button type="button" class="pf-action-toggle pf-clickable" id="import-provider-btn" disabled title="Add provider link in Configuration">${LINK_ICON_SVG}</button>`,
                        "Provider"
                    )}
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle pf-clickable" id="import-open-data-btn" title="Open PTO_Data sheet">${TABLE_ICON_SVG}</button>`,
                        "PTO_Data"
                    )}
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle pf-clickable" id="import-clear-btn" title="Clear PTO_Data to start over">${TRASH_ICON_SVG}</button>`,
                        "Clear"
                    )}
                </div>
            </article>
            ${renderInlineNotes({
                textareaId: "step-notes-1",
                value: stepFields?.notes || "",
                permanentId: "step-notes-lock-1",
                isPermanent: notesPermanent,
                hintId: "",
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
        </section>
    `;
}

function renderStepView(stepId) {
    const detail = WORKFLOW_STEPS.find((step) => step.id === stepId);
    if (!detail) return "";
    if (stepId === 0) return renderConfigView();
    if (stepId === 1) return renderImportStep(detail);
    if (stepId === 2) return renderHeadcountStep(detail);
    if (stepId === 3) return renderDataQualityStep(detail);
    if (stepId === 4) return renderAccrualReviewStep(detail);
    if (stepId === 5) return renderJournalStep(detail);
    if (detail.id === 6) {
        return renderArchiveStep(detail);
    }
    // Generic step rendering (shouldn't be reached now)
    const stepFields = getStepConfig(stepId);
    const notesPermanent = Boolean(configState.permanents[stepId]);
    const stepReviewer = getReviewerWithFallback(stepFields?.reviewer);
    const stepSignOff = stepFields?.signOffDate || "";
    const stepComplete = Boolean(parseBooleanStrict(configState.completes[stepId]) || stepSignOff);
    return `
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step ${detail.id}</p>
            <h2 class="pf-hero-title">${escapeHtml(detail.title)}</h2>
            <p class="pf-hero-copy">${escapeHtml(detail.summary || "")}</p>
        </section>
        <section class="pf-step-guide">
            ${renderInlineNotes({
                textareaId: `step-notes-${stepId}`,
                value: stepFields?.notes || "",
                permanentId: `step-notes-lock-${stepId}`,
                isPermanent: notesPermanent,
                hintId: "",
                saveButtonId: `step-notes-save-${stepId}`
            })}
            ${renderSignoff({
                reviewerInputId: `step-reviewer-${stepId}`,
                reviewerValue: stepReviewer,
                signoffInputId: `step-signoff-${stepId}`,
                signoffValue: stepSignOff,
                isComplete: stepComplete,
                saveButtonId: `step-signoff-save-${stepId}`,
                completeButtonId: `step-signoff-toggle-${stepId}`
            })}
        </section>
    `;
}

function bindInteractions() {
    document.getElementById("nav-home")?.addEventListener("click", async () => {
        // Activate the module homepage sheet
        const homepageConfig = getHomepageConfig(MODULE_KEY);
        await activateHomepageSheet(homepageConfig.sheetName, homepageConfig.title, homepageConfig.subtitle);
        
        setState({ activeView: "home", activeStepId: null });
        document.getElementById("pf-hero")?.scrollIntoView({ behavior: "smooth", block: "start" });
    });
    document.getElementById("nav-selector")?.addEventListener("click", () => {
        window.location.href = SELECTOR_URL;
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
    
    // CONFIGURATION link - unhides SS_* sheets for config access
    document.getElementById("showConfigSheets")?.addEventListener("click", async () => {
        await unhideSystemSheets();
    });

    document.querySelectorAll("[data-step-card]").forEach((card) => {
        const index = Number(card.getAttribute("data-step-index"));
        const stepId = Number(card.getAttribute("data-step-id"));
        card.addEventListener("click", () => focusStep(index, stepId));
    });
    if (appState.activeView === "config") {
        bindConfigView();
    } else if (appState.activeView === "step" && appState.activeStepId != null) {
        bindStepView(appState.activeStepId);
    }
}

function bindStepView(stepId) {
    const notesInput =
        stepId === 2 ? document.getElementById("step-notes-input") : document.getElementById(`step-notes-${stepId}`);
    const reviewerInput =
        stepId === 2 ? document.getElementById("step-reviewer-name") : document.getElementById(`step-reviewer-${stepId}`);
    const signoffInput =
        stepId === 2 ? document.getElementById("step-signoff-date") : document.getElementById(`step-signoff-${stepId}`);
    const backBtn = document.getElementById("step-back-btn");
    const lockBtn =
        stepId === 2 ? document.getElementById("step-notes-lock-2") : document.getElementById(`step-notes-lock-${stepId}`);

    // Save button for notes
    const notesSaveBtn = stepId === 2 ? document.getElementById("step-notes-save-2") : document.getElementById(`step-notes-save-${stepId}`);
    notesSaveBtn?.addEventListener("click", async () => {
        const notes = notesInput?.value || "";
        await saveStepField(stepId, "notes", notes);
        updateSaveButtonState(notesSaveBtn, true);
    });

    // Save button for sign-off section (reviewer name)
    const signoffSaveBtn = stepId === 2 ? document.getElementById("headcount-signoff-save") : document.getElementById(`step-signoff-save-${stepId}`);
    signoffSaveBtn?.addEventListener("click", async () => {
        const reviewer = reviewerInput?.value || "";
        await saveStepField(stepId, "reviewer", reviewer);
        updateSaveButtonState(signoffSaveBtn, true);
    });

    // Auto-wire save state tracking for all save buttons (marks as unsaved on input change)
    initSaveTracking();

    const signoffButtonId = stepId === 2 ? "headcount-signoff-toggle" : `step-signoff-toggle-${stepId}`;
    const signoffInputId = stepId === 2 ? "step-signoff-date" : `step-signoff-${stepId}`;
    bindSignoffToggle(stepId, {
        buttonId: signoffButtonId,
        inputId: signoffInputId,
        canActivate:
            stepId === 2
                ? () => {
                      if (!isHeadcountNotesRequired()) return true;
                      const notes = document.getElementById("step-notes-input")?.value.trim() || "";
                      if (notes) return true;
                      window.alert("Please enter a brief explanation of the headcount differences before completing this step.");
                      return false;
                  }
                : null,
        onComplete: stepId === 2 ? handleHeadcountSignoff : null
    });
    backBtn?.addEventListener("click", async () => {
        const homepageConfig = getHomepageConfig(MODULE_KEY);
        await activateHomepageSheet(homepageConfig.sheetName, homepageConfig.title, homepageConfig.subtitle);
        setState({ activeView: "home", activeStepId: null });
    });
    lockBtn?.addEventListener("click", async () => {
        const nextLocked = !lockBtn.classList.contains("is-locked");
        updateLockButtonVisual(lockBtn, nextLocked);
        await toggleNotePermanent(stepId, nextLocked);
    });
    if (stepId === 6) {
        document.getElementById("archive-run-btn")?.addEventListener("click", () => {
            // Archive flow coming soon
        });
    }
    // Step 1: Import PTO Data
    if (stepId === 1) {
        document.getElementById("import-open-data-btn")?.addEventListener("click", () => openSheet("PTO_Data"));
        document.getElementById("import-clear-btn")?.addEventListener("click", () => clearPtoData());
    }
    // Step 2: Headcount Review
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
        document.getElementById("headcount-run-btn")?.addEventListener("click", () => refreshHeadcountAnalysis());
        document.getElementById("headcount-refresh-btn")?.addEventListener("click", () => refreshHeadcountAnalysis());
        bindHeadcountNotesGuard();
        if (headcountState.skipAnalysis) {
            enforceHeadcountSkipNote();
        }
        updateHeadcountSignoffState();
    }
    if (stepId === 3) {
        document.getElementById("quality-run-btn")?.addEventListener("click", () => runDataQualityCheck());
        document.getElementById("quality-refresh-btn")?.addEventListener("click", () => runDataQualityCheck());
        document.getElementById("quality-acknowledge-btn")?.addEventListener("click", () => acknowledgeQualityIssues());
    }
    if (stepId === 4) {
        // Both buttons trigger the full analysis with all checks
        document.getElementById("analysis-refresh-btn")?.addEventListener("click", () => runFullAnalysis());
        document.getElementById("analysis-run-btn")?.addEventListener("click", () => runFullAnalysis());
        
        // Missing pay rate card bindings
        document.getElementById("payrate-save-btn")?.addEventListener("click", handlePayRateSave);
        document.getElementById("payrate-ignore-btn")?.addEventListener("click", handlePayRateIgnore);
        document.getElementById("payrate-input")?.addEventListener("keydown", (e) => {
            if (e.key === "Enter") handlePayRateSave();
        });
    }
    if (stepId === 5) {
        document.getElementById("je-create-btn")?.addEventListener("click", () => createJournalDraft());
        document.getElementById("je-run-btn")?.addEventListener("click", () => runJournalSummary());
        document.getElementById("je-export-btn")?.addEventListener("click", () => exportJournalDraft());
    }
}

function bindConfigView() {
    const payrollInput = document.getElementById("config-payroll-date");
    payrollInput?.addEventListener("change", (event) => {
        const value = event.target.value || "";
        scheduleConfigWrite(PTO_CONFIG_FIELDS.payrollDate, value);
        if (!value) return;
        if (!configState.overrides.accountingPeriod) {
            const derivedPeriod = deriveAccountingPeriod(value);
            if (derivedPeriod) {
                const periodInput = document.getElementById("config-accounting-period");
                if (periodInput) periodInput.value = derivedPeriod;
                scheduleConfigWrite(PTO_CONFIG_FIELDS.accountingPeriod, derivedPeriod);
            }
        }
        if (!configState.overrides.journalId) {
            const derivedJe = deriveJournalId(value);
            if (derivedJe) {
                const jeInput = document.getElementById("config-journal-id");
                if (jeInput) jeInput.value = derivedJe;
                scheduleConfigWrite(PTO_CONFIG_FIELDS.journalEntryId, derivedJe);
            }
        }
    });

    const periodInput = document.getElementById("config-accounting-period");
    periodInput?.addEventListener("change", (event) => {
        configState.overrides.accountingPeriod = Boolean(event.target.value);
        scheduleConfigWrite(PTO_CONFIG_FIELDS.accountingPeriod, event.target.value || "");
    });

    const journalInput = document.getElementById("config-journal-id");
    journalInput?.addEventListener("change", (event) => {
        configState.overrides.journalId = Boolean(event.target.value);
        scheduleConfigWrite(PTO_CONFIG_FIELDS.journalEntryId, event.target.value.trim());
    });

    document.getElementById("config-company-name")?.addEventListener("change", (event) => {
        scheduleConfigWrite(PTO_CONFIG_FIELDS.companyName, event.target.value.trim());
    });

    document.getElementById("config-payroll-provider")?.addEventListener("change", (event) => {
        scheduleConfigWrite(PTO_CONFIG_FIELDS.payrollProvider, event.target.value.trim());
    });

    document.getElementById("config-accounting-link")?.addEventListener("change", (event) => {
        scheduleConfigWrite(PTO_CONFIG_FIELDS.accountingSoftware, event.target.value.trim());
    });

    document.getElementById("config-user-name")?.addEventListener("change", (event) => {
        scheduleConfigWrite(PTO_CONFIG_FIELDS.reviewerName, event.target.value.trim());
    });

    const notesInput = document.getElementById("config-notes");
    notesInput?.addEventListener("input", (event) => {
        saveStepField(0, "notes", event.target.value);
    });

    const lockButton = document.getElementById("config-notes-lock");
    lockButton?.addEventListener("click", async () => {
        const nextLocked = !lockButton.classList.contains("is-locked");
        updateLockButtonVisual(lockButton, nextLocked);
        await toggleNotePermanent(0, nextLocked);
    });

    const notesSaveBtn = document.getElementById("config-notes-save");
    notesSaveBtn?.addEventListener("click", async () => {
        if (!notesInput) return;
        await saveStepField(0, "notes", notesInput.value);
        updateSaveButtonState(notesSaveBtn, true);
    });

    const reviewerInput = document.getElementById("config-reviewer");
    reviewerInput?.addEventListener("change", (event) => {
        const value = event.target.value.trim();
        saveStepField(0, "reviewer", value);
        const signoffInput = document.getElementById("config-signoff-date");
        if (value && signoffInput && !signoffInput.value) {
            const today = todayIso();
            signoffInput.value = today;
            saveStepField(0, "signOffDate", today);
            saveCompletionFlag(0, true);
        }
    });

    document.getElementById("config-signoff-date")?.addEventListener("change", (event) => {
        saveStepField(0, "signOffDate", event.target.value || "");
    });
    const signoffSaveBtn = document.getElementById("config-signoff-save");
    signoffSaveBtn?.addEventListener("click", async () => {
        const reviewerValue = reviewerInput?.value?.trim() || "";
        const signoffValue = document.getElementById("config-signoff-date")?.value || "";
        await saveStepField(0, "reviewer", reviewerValue);
        await saveStepField(0, "signOffDate", signoffValue);
        updateSaveButtonState(signoffSaveBtn, true);
    });

    initSaveTracking();
    bindSignoffToggle(0, {
        buttonId: "config-signoff-toggle",
        inputId: "config-signoff-date",
        onComplete: persistConfigBasics
    });
}

function focusStep(index, stepId = null) {
    if (index < 0 || index >= WORKFLOW_STEPS.length) return;
    pendingScrollIndex = index;
    const resolvedStepId = stepId ?? WORKFLOW_STEPS[index].id;
    const nextView = resolvedStepId === 0 ? "config" : "step";
    setState({ focusedIndex: index, activeView: nextView, activeStepId: resolvedStepId });
    if (resolvedStepId === 1) {
        openSheet("PTO_Data");
    }
    if (resolvedStepId === 2 && !headcountState.hasAnalyzed) {
        void syncPtoAnalysis();
        refreshHeadcountAnalysis();
    }
    if (resolvedStepId === 3) {
        openSheet("PTO_Data");
    }
    // Step 4 no longer runs checks automatically - user clicks Run/Refresh
    if (resolvedStepId === 5) {
        openSheet("PTO_JE_Draft");
    }
}

function moveFocus(delta) {
    const next = appState.focusedIndex + delta;
    const target = Math.max(0, Math.min(WORKFLOW_STEPS.length - 1, next));
    focusStep(target, WORKFLOW_STEPS[target].id);
}

function scrollFocusedIntoView() {
    if (pendingScrollIndex === null) return;
    const card = document.querySelector(`[data-step-index="${pendingScrollIndex}"]`);
    pendingScrollIndex = null;
    card?.scrollIntoView({ behavior: "smooth", block: "center" });
}

function isStepComplete(stepId) {
    return parseBooleanFlag(configState.completes[stepId]);
}

function handleStepAction(stepId) {
    switch (stepId) {
        case 1:
            importSampleData();
            break;
        case 2:
            openSheet("SS_Employee_Roster");
            break;
        case 3:
            // Validation triggered via button only for explicit user control
            break;
        case 4:
            openSheet("PTO_ExpenseReview");
            break;
        case 5:
            openSheet("PTO_JE_Draft");
            break;
        case 6:
            archiveAndReset();
            break;
        default:
            break;
    }
}

function setState(partial) {
    if (partial.stepStatuses) {
        appState.stepStatuses = { ...appState.stepStatuses, ...partial.stepStatuses };
    }
    Object.assign(appState, { ...partial, stepStatuses: appState.stepStatuses });
    renderApp();
}

function hasExcel() {
    return typeof Excel !== "undefined" && typeof Excel.run === "function";
}

async function importSampleData() {
    if (!hasExcel()) {
        return;
    }
    toggleLoader(true, "Importing sample data...");
    try {
        await Excel.run(async (context) => {
            await writeDatasetToSheet(context, "PTO_Data", PTO_ACTIVITY_COLUMNS, SAMPLE_PTO_ACTIVITY);
            await writeDatasetToSheet(context, "PTO_ExpenseReview", PTO_EXPENSE_COLUMNS, SAMPLE_EXPENSE_REVIEW);
            await writeDatasetToSheet(context, "PTO_JE_Draft", PTO_JOURNAL_COLUMNS, SAMPLE_JOURNAL_LINES);
            // Sync PTO_Analysis from PTO_Data (enriches with department, pay rate, calculates liability)
            await syncPtoAnalysis(context);
            const dataSheet = context.workbook.worksheets.getItem("PTO_Data");
            dataSheet.activate();
            dataSheet.getRange("A1").select();
            await context.sync();
        });

        setState({ stepStatuses: { 1: "complete" } });
    } catch (error) {
        console.error(error);
    } finally {
        toggleLoader(false);
    }
}

async function refreshValidationData() {
    if (!hasExcel()) {
        window.alert("Excel is not available. Open this module inside Excel to refresh data.");
        return;
    }
    toggleLoader(true, "Refreshing PTO_Analysis...");
    try {
        await syncPtoAnalysis();
        // After refresh, automatically run validation and completeness check
        toggleLoader(false);
        await runBalanceValidation();
        await runCompletenessCheck();
        // Re-render step 4 if we're on it to update the completeness pills
        if (appState.activeStepId === 4) {
            renderApp();
        }
    } catch (error) {
        console.error("Refresh error:", error);
        window.alert(`Failed to refresh data: ${error.message}`);
        toggleLoader(false);
    }
}

/**
 * Handle save button click for missing pay rate card
 * Updates PTO_Analysis directly with the entered pay rate
 */
async function handlePayRateSave() {
    const input = document.getElementById("payrate-input");
    if (!input) return;
    
    const payRate = parseFloat(input.value);
    const employeeName = input.dataset.employee;
    const rowIndex = parseInt(input.dataset.row, 10);
    
    if (isNaN(payRate) || payRate <= 0) {
        window.alert("Please enter a valid pay rate greater than 0.");
        return;
    }
    
    if (!employeeName || isNaN(rowIndex)) {
        console.error("Missing employee data on input");
        return;
    }
    
    toggleLoader(true, "Updating pay rate...");
    
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("PTO_Analysis");
            
            // Pay Rate is column D (index 3)
            const payRateCell = sheet.getCell(rowIndex - 1, 3); // -1 because Excel rows are 1-based but getCell is 0-based
            payRateCell.values = [[payRate]];
            
            // Recalculate Liability Amount (column J, index 9) = Balance × Pay Rate
            const balanceCell = sheet.getCell(rowIndex - 1, 8); // Balance is column I (index 8)
            balanceCell.load("values");
            await context.sync();
            
            const balance = Number(balanceCell.values[0][0]) || 0;
            const liabilityAmount = balance * payRate;
            
            const liabilityCell = sheet.getCell(rowIndex - 1, 9);
            liabilityCell.values = [[liabilityAmount]];
            
            // Also recalculate Change column (column L, index 11) = Liability - Prior
            const priorCell = sheet.getCell(rowIndex - 1, 10); // Prior is column K (index 10)
            priorCell.load("values");
            await context.sync();
            
            const priorLiability = Number(priorCell.values[0][0]) || 0;
            const change = liabilityAmount - priorLiability;
            
            const changeCell = sheet.getCell(rowIndex - 1, 11);
            changeCell.values = [[change]];
            
            await context.sync();
        });
        
        // Remove this employee from missing list and re-render
        analysisState.missingPayRates = analysisState.missingPayRates.filter(e => e.name !== employeeName);
        toggleLoader(false);
        
        // Re-render to show next missing employee or remove card
        focusStep(3, 3);
        
    } catch (error) {
        console.error("Failed to save pay rate:", error);
        window.alert(`Failed to save pay rate: ${error.message}`);
        toggleLoader(false);
    }
}

/**
 * Handle ignore button click for missing pay rate card
 * Skips this employee and shows the next one
 */
function handlePayRateIgnore() {
    const input = document.getElementById("payrate-input");
    if (!input) return;
    
    const employeeName = input.dataset.employee;
    if (employeeName) {
        // Add to ignored set
        analysisState.ignoredMissingPayRates.add(employeeName);
        // Remove from current missing list
        analysisState.missingPayRates = analysisState.missingPayRates.filter(e => e.name !== employeeName);
    }
    
    // Re-render to show next missing employee or remove card
    focusStep(3, 3);
}

async function runDataQualityCheck() {
    if (!hasExcel()) {
        window.alert("Excel is not available. Open this module inside Excel to run quality check.");
        return;
    }
    
    dataQualityState.loading = true;
    toggleLoader(true, "Analyzing data quality...");
    updateSaveButtonState(document.getElementById("quality-save-btn"), false);
    
    try {
        await Excel.run(async (context) => {
            const dataSheet = context.workbook.worksheets.getItem("PTO_Data");
            const dataRange = dataSheet.getUsedRangeOrNullObject();
            dataRange.load("values");
            await context.sync();
            
            const dataValues = dataRange.isNullObject ? [] : dataRange.values || [];
            
            if (!dataValues.length || dataValues.length < 2) {
                throw new Error("PTO_Data is empty or has no data rows.");
            }
            
            // Parse headers
            const headers = (dataValues[0] || []).map(h => normalizeName(h));
            console.log("[Data Quality] PTO_Data headers:", dataValues[0]);
            
            // Find employee name column - be specific to avoid matching company name
            let nameIdx = headers.findIndex(h => h === "employee name" || h === "employeename");
            if (nameIdx === -1) {
                // Fallback: look for column containing "employee" AND "name"
                nameIdx = headers.findIndex(h => h.includes("employee") && h.includes("name"));
            }
            if (nameIdx === -1) {
                // Last resort: just "name" but not if it also contains "company" or "form"
                nameIdx = headers.findIndex(h => h === "name" || (h.includes("name") && !h.includes("company") && !h.includes("form")));
            }
            console.log("[Data Quality] Employee name column index:", nameIdx, "Header:", dataValues[0]?.[nameIdx]);
            const balanceIdx = findColumnIndex(headers, ["balance"]);
            const accrualRateIdx = findColumnIndex(headers, ["accrual rate", "accrualrate"]);
            const carryOverIdx = findColumnIndex(headers, ["carry over", "carryover"]);
            const ytdAccruedIdx = findColumnIndex(headers, ["ytd accrued", "ytdaccrued"]);
            const ytdUsedIdx = findColumnIndex(headers, ["ytd used", "ytdused"]);
            
            // Reset state
            const balanceIssues = [];
            const zeroBalances = [];
            const accrualOutliers = [];
            
            // Process each employee row
            const dataRows = dataValues.slice(1);
            dataRows.forEach((row, idx) => {
                const rowIndex = idx + 2; // 1-based, after header
                const name = nameIdx !== -1 ? String(row[nameIdx] || "").trim() : `Row ${rowIndex}`;
                if (!name) return;
                
                const balance = balanceIdx !== -1 ? Number(row[balanceIdx]) || 0 : 0;
                const accrualRate = accrualRateIdx !== -1 ? Number(row[accrualRateIdx]) || 0 : 0;
                const carryOver = carryOverIdx !== -1 ? Number(row[carryOverIdx]) || 0 : 0;
                const ytdAccrued = ytdAccruedIdx !== -1 ? Number(row[ytdAccruedIdx]) || 0 : 0;
                const ytdUsed = ytdUsedIdx !== -1 ? Number(row[ytdUsedIdx]) || 0 : 0;
                
                // Check 1: Balance issues - negative balance or used more than available
                const maxUsable = carryOver + ytdAccrued;
                if (balance < 0) {
                    balanceIssues.push({ 
                        name, 
                        issue: `Negative balance: ${balance.toFixed(2)} hrs`,
                        rowIndex 
                    });
                } else if (ytdUsed > maxUsable && maxUsable > 0) {
                    balanceIssues.push({ 
                        name, 
                        issue: `Used ${ytdUsed.toFixed(0)} hrs but only ${maxUsable.toFixed(0)} available`,
                        rowIndex 
                    });
                }
                
                // Check 2: Zero balances (informational)
                if (balance === 0 && (carryOver > 0 || ytdAccrued > 0)) {
                    zeroBalances.push({ name, rowIndex });
                }
                
                // Check 3: Accrual rate outliers (> 8 hrs per period is unusual)
                if (accrualRate > 8) {
                    accrualOutliers.push({ name, accrualRate, rowIndex });
                }
            });
            
            // Update state
            dataQualityState.balanceIssues = balanceIssues;
            dataQualityState.zeroBalances = zeroBalances;
            dataQualityState.accrualOutliers = accrualOutliers;
            dataQualityState.totalIssues = balanceIssues.length;
            dataQualityState.totalEmployees = dataRows.filter(r => r.some(c => c !== null && c !== "")).length;
            dataQualityState.hasRun = true;
        });
        
        // Update step status
        const hasBlockingIssues = dataQualityState.balanceIssues.length > 0;
        setState({ stepStatuses: { 3: hasBlockingIssues ? "blocked" : "complete" } });
        
    } catch (error) {
        console.error("Data quality check error:", error);
        window.alert(`Quality check failed: ${error.message}`);
        dataQualityState.hasRun = false;
    } finally {
        dataQualityState.loading = false;
        toggleLoader(false);
        renderApp();
    }
}

/**
 * User acknowledges quality issues and wants to proceed anyway
 */
function acknowledgeQualityIssues() {
    dataQualityState.acknowledged = true;
    // Allow sign-off even with issues
    setState({ stepStatuses: { 3: "complete" } });
    renderApp();
}

// Note: Save functions removed - auto-save happens via config writes during analysis

/**
 * Run data completeness check comparing PTO_Data to PTO_Analysis column sums
 */
async function runCompletenessCheck() {
    if (!hasExcel()) return;
    
    try {
        await Excel.run(async (context) => {
            const dataSheet = context.workbook.worksheets.getItem("PTO_Data");
            const analysisSheet = context.workbook.worksheets.getItemOrNullObject("PTO_Analysis");
            
            const dataRange = dataSheet.getUsedRangeOrNullObject();
            dataRange.load("values");
            analysisSheet.load("isNullObject");
            await context.sync();
            
            if (analysisSheet.isNullObject) {
                // Reset all checks to null if no analysis sheet
                analysisState.completenessCheck = {
                    accrualRate: null,
                    carryOver: null,
                    ytdAccrued: null,
                    ytdUsed: null,
                    balance: null
                };
                return;
            }
            
            const analysisRange = analysisSheet.getUsedRangeOrNullObject();
            analysisRange.load("values");
            await context.sync();
            
            const dataValues = dataRange.isNullObject ? [] : dataRange.values || [];
            const analysisValues = analysisRange.isNullObject ? [] : analysisRange.values || [];
            
            if (!dataValues.length || !analysisValues.length) {
                analysisState.completenessCheck = {
                    accrualRate: null,
                    carryOver: null,
                    ytdAccrued: null,
                    ytdUsed: null,
                    balance: null
                };
                return;
            }
            
            // Helper to find column and sum values
            const sumColumn = (rows, columnAliases, label) => {
                const headers = (rows[0] || []).map(h => normalizeName(h));
                const idx = findColumnIndex(headers, columnAliases);
                if (idx === -1) return null;
                const dataRows = rows.slice(1);
                return dataRows.reduce((sum, row) => sum + (Number(row[idx]) || 0), 0);
            };
            
            // Column mappings for each field
            const fields = [
                { key: "accrualRate", aliases: ["accrual rate", "accrualrate"] },
                { key: "carryOver", aliases: ["carry over", "carryover", "carry_over"] },
                { key: "ytdAccrued", aliases: ["ytd accrued", "ytdaccrued", "ytd_accrued"] },
                { key: "ytdUsed", aliases: ["ytd used", "ytdused", "ytd_used"] },
                { key: "balance", aliases: ["balance"] }
            ];
            
            const results = {};
            
            for (const field of fields) {
                const ptoDataSum = sumColumn(dataValues, field.aliases, "PTO_Data");
                const analysisSum = sumColumn(analysisValues, field.aliases, "PTO_Analysis");
                
                if (ptoDataSum === null || analysisSum === null) {
                    results[field.key] = null;
                } else {
                    // Use tolerance for floating point comparison
                    const match = Math.abs(ptoDataSum - analysisSum) < 0.01;
                    results[field.key] = {
                        match,
                        ptoData: ptoDataSum,
                        ptoAnalysis: analysisSum
                    };
                }
            }
            
            analysisState.completenessCheck = results;
        });
    } catch (error) {
        console.error("Completeness check failed:", error);
    }
}

/**
 * Run full analysis - syncs data and runs all verification checks
 */
async function runFullAnalysis() {
    if (!hasExcel()) {
        window.alert("Excel is not available. Open this module inside Excel to run analysis.");
        return;
    }
    
    toggleLoader(true, "Running analysis...");
    
    try {
        // 1. Sync PTO_Analysis from PTO_Data with enrichment
        await syncPtoAnalysis();
        
        // 2. Run completeness check (compare column sums)
        await runCompletenessCheck();
        
        // 3. Update state
        analysisState.cleanDataReady = true;
        
        // 4. Re-render to show results
        renderApp();
        
    } catch (error) {
        console.error("Full analysis error:", error);
        window.alert(`Analysis failed: ${error.message}`);
    } finally {
        toggleLoader(false);
    }
}

async function populatePtoAnalysis() {
    if (!hasExcel()) {
        analysisState.lastError = "Excel is not available. Open this module inside Excel to run analysis.";
        renderApp();
        return;
    }
    
    analysisState.loading = true;
    analysisState.lastError = null;
    renderApp();
    
    try {
        // Sync PTO_Analysis directly from PTO_Data (with enrichment from roster/archive)
        await syncPtoAnalysis();
        
        // Get the row count from PTO_Analysis
        const result = await Excel.run(async (context) => {
            const analysisSheet = context.workbook.worksheets.getItemOrNullObject("PTO_Analysis");
            analysisSheet.load("isNullObject");
            await context.sync();
            
            if (analysisSheet.isNullObject) {
                throw new Error("PTO_Analysis sheet was not created. Please check that PTO_Data has data.");
            }
            
            const analysisRange = analysisSheet.getUsedRangeOrNullObject();
            analysisRange.load("values");
            await context.sync();
            
            const values = analysisRange.isNullObject ? [] : analysisRange.values || [];
            const dataRows = values.length > 1 ? values.length - 1 : 0;
            
            if (dataRows === 0) {
                throw new Error("No employee data found. Please import PTO data in Step 1 first.");
            }
            
            return { employeeCount: dataRows };
        });
        
        // Update state with success
        analysisState.cleanDataReady = true;
        analysisState.employeeCount = result.employeeCount;
        analysisState.lastRun = new Date().toISOString();
        analysisState.lastError = null;
        
        setState({ stepStatuses: { 4: "complete" } });
    } catch (error) {
        console.error(error);
        analysisState.lastError = error?.message || "An unexpected error occurred while running the analysis.";
        analysisState.lastRun = null;
    } finally {
        analysisState.loading = false;
        renderApp();
    }
}

/**
 * Check PTO_Analysis status when entering Step 4
 */
async function checkAnalysisPrerequisites() {
    if (!hasExcel()) {
        analysisState.cleanDataReady = false;
        analysisState.employeeCount = 0;
        return;
    }
    
    try {
        const result = await Excel.run(async (context) => {
            const analysisSheet = context.workbook.worksheets.getItemOrNullObject("PTO_Analysis");
            analysisSheet.load("isNullObject");
            await context.sync();
            
            if (analysisSheet.isNullObject) {
                return { ready: false, count: 0 };
            }
            
            const analysisRange = analysisSheet.getUsedRangeOrNullObject();
            analysisRange.load("values");
            await context.sync();
            
            const values = analysisRange.isNullObject ? [] : analysisRange.values || [];
            const dataRows = values.length > 1 ? values.length - 1 : 0;
            
            return { ready: dataRows > 0, count: dataRows };
        });
        
        analysisState.cleanDataReady = result.ready;
        analysisState.employeeCount = result.count;
    } catch (error) {
        console.error("Error checking analysis prerequisites:", error);
        analysisState.cleanDataReady = false;
        analysisState.employeeCount = 0;
    }
}

async function runJournalSummary() {
    if (!hasExcel()) {
        window.alert("Excel is not available. Open this module inside Excel to run journal checks.");
        return;
    }
    journalState.loading = true;
    journalState.lastError = null;
    updateSaveButtonState(document.getElementById("je-save-btn"), false);
    renderApp();
    try {
        const totals = await Excel.run(async (context) => {
            // Read JE Draft
            const jeSheet = context.workbook.worksheets.getItem("PTO_JE_Draft");
            const jeRange = jeSheet.getUsedRangeOrNullObject();
            jeRange.load("values");
            
            // Read PTO_Analysis for comparison
            const analysisSheet = context.workbook.worksheets.getItemOrNullObject("PTO_Analysis");
            analysisSheet.load("isNullObject");
            await context.sync();
            
            const jeValues = jeRange.isNullObject ? [] : jeRange.values || [];
            if (!jeValues.length) {
                throw new Error("PTO_JE_Draft is empty. Generate the JE first.");
            }
            
            // Parse JE headers
            const jeHeaders = (jeValues[0] || []).map((h) => normalizeName(h));
            const debitIdx = findColumnIndex(jeHeaders, ["debit"]);
            const creditIdx = findColumnIndex(jeHeaders, ["credit"]);
            const lineAmountIdx = findColumnIndex(jeHeaders, ["lineamount", "line amount"]);
            const acctNumIdx = findColumnIndex(jeHeaders, ["account number", "accountnumber"]);
            
            if (debitIdx === -1 || creditIdx === -1) {
                throw new Error("Could not find Debit and Credit columns in PTO_JE_Draft.");
            }
            
            let debitTotal = 0;
            let creditTotal = 0;
            let lineAmountSum = 0;
            let jeChangeTotal = 0; // Sum of expense line amounts (not liability offset)
            
            jeValues.slice(1).forEach((row) => {
                const debit = Number(row[debitIdx]) || 0;
                const credit = Number(row[creditIdx]) || 0;
                const lineAmount = lineAmountIdx !== -1 ? Number(row[lineAmountIdx]) || 0 : 0;
                const acctNum = acctNumIdx !== -1 ? String(row[acctNumIdx] || "").trim() : "";
                
                debitTotal += debit;
                creditTotal += credit;
                lineAmountSum += lineAmount;
                
                // Sum expense lines only (not the liability offset 21540)
                if (acctNum && acctNum !== "21540") {
                    jeChangeTotal += lineAmount;
                }
            });
            
            // Get PTO_Analysis total change
            let analysisChangeTotal = 0;
            if (!analysisSheet.isNullObject) {
                const analysisRange = analysisSheet.getUsedRangeOrNullObject();
                analysisRange.load("values");
                await context.sync();
                
                const analysisValues = analysisRange.isNullObject ? [] : analysisRange.values || [];
                if (analysisValues.length > 1) {
                    const analysisHeaders = (analysisValues[0] || []).map(h => normalizeName(h));
                    const changeIdx = findColumnIndex(analysisHeaders, ["change"]);
                    
                    if (changeIdx !== -1) {
                        analysisValues.slice(1).forEach(row => {
                            analysisChangeTotal += Number(row[changeIdx]) || 0;
                        });
                    }
                }
            }
            
            // Build validation issues array
            const difference = debitTotal - creditTotal;
            const issues = [];
            
            // Check 1: Debits = Credits
            if (Math.abs(difference) >= 0.01) {
                issues.push({
                    check: "Debits = Credits",
                    passed: false,
                    detail: difference > 0 
                        ? `Debits exceed credits by $${Math.abs(difference).toLocaleString(undefined, {minimumFractionDigits: 2})}`
                        : `Credits exceed debits by $${Math.abs(difference).toLocaleString(undefined, {minimumFractionDigits: 2})}`
                });
            } else {
                issues.push({ check: "Debits = Credits", passed: true, detail: "" });
            }
            
            // Check 2: Line Amounts Sum to Zero
            if (Math.abs(lineAmountSum) >= 0.01) {
                issues.push({
                    check: "Line Amounts Sum to Zero",
                    passed: false,
                    detail: `Line amounts sum to $${lineAmountSum.toLocaleString(undefined, {minimumFractionDigits: 2})} (should be $0.00)`
                });
            } else {
                issues.push({ check: "Line Amounts Sum to Zero", passed: true, detail: "" });
            }
            
            // Check 3: JE Matches Analysis Total
            const changeDiff = Math.abs(jeChangeTotal - analysisChangeTotal);
            if (changeDiff >= 0.01) {
                issues.push({
                    check: "JE Matches Analysis Total",
                    passed: false,
                    detail: `JE expense total ($${jeChangeTotal.toLocaleString(undefined, {minimumFractionDigits: 2})}) differs from PTO_Analysis Change total ($${analysisChangeTotal.toLocaleString(undefined, {minimumFractionDigits: 2})}) by $${changeDiff.toLocaleString(undefined, {minimumFractionDigits: 2})}`
                });
            } else {
                issues.push({ check: "JE Matches Analysis Total", passed: true, detail: "" });
            }
            
            return { 
                debitTotal, 
                creditTotal, 
                difference,
                lineAmountSum,
                jeChangeTotal,
                analysisChangeTotal,
                issues,
                validationRun: true
            };
        });
        Object.assign(journalState, totals, { lastError: null });
    } catch (error) {
        console.warn("PTO JE summary:", error);
        journalState.lastError = error?.message || "Unable to calculate journal totals.";
        journalState.debitTotal = null;
        journalState.creditTotal = null;
        journalState.difference = null;
        journalState.lineAmountSum = null;
        journalState.jeChangeTotal = null;
        journalState.analysisChangeTotal = null;
        journalState.issues = [];
        journalState.validationRun = false;
    } finally {
        journalState.loading = false;
        renderApp();
    }
}

/**
 * Department to Expense Account mapping for PTO accrual entries
 */
const DEPARTMENT_EXPENSE_ACCOUNTS = {
    "general & administrative": "64110",
    "general and administrative": "64110",
    "g&a": "64110",
    "research & development": "62110",
    "research and development": "62110",
    "r&d": "62110",
    "marketing": "61610",
    "cogs onboarding": "53110",
    "cogs prof. services": "56110",
    "cogs professional services": "56110",
    "sales & marketing": "61110",
    "sales and marketing": "61110",
    "cogs support": "52110",
    "client success": "61811"
};

const LIABILITY_OFFSET_ACCOUNT = "21540";

/**
 * Create PTO Journal Entry Draft
 * Groups Change amounts by Department and creates proper debit/credit entries
 */
async function createJournalDraft() {
    if (!hasExcel()) {
        window.alert("Excel is not available. Open this module inside Excel to create the journal entry.");
        return;
    }
    
    toggleLoader(true, "Creating PTO Journal Entry...");
    
    try {
        await Excel.run(async (context) => {
            // 1. Read config values
            const configTable = context.workbook.tables.getItem(CONFIG_TABLES[0]);
            const configRange = configTable.getDataBodyRange();
            configRange.load("values");
            
            // 2. Read PTO_Analysis
            const analysisSheet = context.workbook.worksheets.getItem("PTO_Analysis");
            const analysisRange = analysisSheet.getUsedRangeOrNullObject();
            analysisRange.load("values");
            
            // 3. Read Chart of Accounts for account name lookups
            const coaSheet = context.workbook.worksheets.getItemOrNullObject("SS_Chart_of_Accounts");
            coaSheet.load("isNullObject");
            await context.sync();
            
            let coaData = [];
            if (!coaSheet.isNullObject) {
                const coaRange = coaSheet.getUsedRangeOrNullObject();
                coaRange.load("values");
                await context.sync();
                coaData = coaRange.isNullObject ? [] : coaRange.values || [];
            }
            
            const configValues = configRange.values || [];
            const analysisValues = analysisRange.isNullObject ? [] : analysisRange.values || [];
            
            if (!analysisValues.length) {
                throw new Error("PTO_Analysis is empty. Run the analysis first.");
            }
            
            // Parse config values - SS_PF_Config structure:
            // Column 0 = Category (e.g., "Tab Structure", "PTO")
            // Column 1 = Field name (e.g., "PTO_Journal_Entry_ID")
            // Column 2 = Value (the actual value we need)
            // Column 3 = Value2 (used for tab structure)
            const configMap = {};
            configValues.forEach(row => {
                const field = String(row[1] || "").trim(); // Column 1 = Field name
                const value = row[2]; // Column 2 = Value
                if (field) {
                    configMap[field] = value;
                }
            });
            
            // Log config lookup results (useful for troubleshooting)
            if (!configMap[PTO_CONFIG_FIELDS.journalEntryId] || !configMap[PTO_CONFIG_FIELDS.payrollDate]) {
                console.warn("[JE Draft] Missing config values - RefNumber:", configMap[PTO_CONFIG_FIELDS.journalEntryId], "TxnDate:", configMap[PTO_CONFIG_FIELDS.payrollDate]);
            }
            
            const refNumber = configMap[PTO_CONFIG_FIELDS.journalEntryId] || "";
            const rawTxnDate = configMap[PTO_CONFIG_FIELDS.payrollDate] || "";
            const accountingPeriod = configMap[PTO_CONFIG_FIELDS.accountingPeriod] || "";
            
            // Format TxnDate as MM/DD/YYYY
            // Handle Excel serial numbers, ISO strings, or already formatted dates
            let txnDate = "";
            if (rawTxnDate) {
                try {
                    let d;
                    // Check if it's an Excel serial number (typically 5 digits like 45626)
                    if (typeof rawTxnDate === "number" || /^\d{4,5}$/.test(String(rawTxnDate).trim())) {
                        // Excel serial date: days since Jan 1, 1900 (with leap year bug adjustment)
                        const serialNum = Number(rawTxnDate);
                        // Excel incorrectly treats 1900 as leap year, so subtract 1 for dates after Feb 28, 1900
                        const excelEpoch = new Date(1899, 11, 30); // Dec 30, 1899
                        d = new Date(excelEpoch.getTime() + serialNum * 24 * 60 * 60 * 1000);
                    } else {
                        d = new Date(rawTxnDate);
                    }
                    
                    if (!isNaN(d.getTime()) && d.getFullYear() > 1970) {
                        const mm = String(d.getMonth() + 1).padStart(2, "0");
                        const dd = String(d.getDate()).padStart(2, "0");
                        const yyyy = d.getFullYear();
                        txnDate = `${mm}/${dd}/${yyyy}`;
                    } else {
                        console.warn("[JE Draft] Date parsing resulted in invalid date:", rawTxnDate, "->", d);
                        txnDate = String(rawTxnDate); // Fall back to raw value
                    }
                } catch (e) {
                    console.warn("[JE Draft] Could not parse TxnDate:", rawTxnDate, e);
                    txnDate = String(rawTxnDate); // Fall back to raw value
                }
            }
            
            const lineDescBase = accountingPeriod ? `${accountingPeriod} PTO Accrual` : "PTO Accrual";
            
            // Build Chart of Accounts lookup map
            const coaMap = {};
            if (coaData.length > 1) {
                const coaHeaders = (coaData[0] || []).map(h => normalizeName(h));
                const acctNumIdx = findColumnIndex(coaHeaders, ["account number", "accountnumber", "account", "acct"]);
                const acctNameIdx = findColumnIndex(coaHeaders, ["account name", "accountname", "name", "description"]);
                
                if (acctNumIdx !== -1 && acctNameIdx !== -1) {
                    coaData.slice(1).forEach(row => {
                        const num = String(row[acctNumIdx] || "").trim();
                        const name = String(row[acctNameIdx] || "").trim();
                        if (num) coaMap[num] = name;
                    });
                }
            }
            
            // Parse PTO_Analysis headers and find columns
            const headers = (analysisValues[0] || []).map(h => normalizeName(h));
            const deptIdx = findColumnIndex(headers, ["department"]);
            const changeIdx = findColumnIndex(headers, ["change"]);
            
            if (deptIdx === -1 || changeIdx === -1) {
                throw new Error("Could not find Department or Change columns in PTO_Analysis.");
            }
            
            // Group Change amounts by Department
            const deptTotals = {};
            analysisValues.slice(1).forEach(row => {
                const dept = String(row[deptIdx] || "").trim();
                const change = Number(row[changeIdx]) || 0;
                
                if (dept && change !== 0) {
                    if (!deptTotals[dept]) {
                        deptTotals[dept] = 0;
                    }
                    deptTotals[dept] += change;
                }
            });
            
            // Build JE rows
            const jeHeaders = ["RefNumber", "TxnDate", "Account Number", "Account Name", "LineAmount", "Debit", "Credit", "LineDesc", "Department"];
            const jeRows = [jeHeaders];
            
            let totalDebits = 0;
            let totalCredits = 0;
            
            // Create expense lines by department
            Object.entries(deptTotals).forEach(([dept, change]) => {
                if (Math.abs(change) < 0.01) return; // Skip zero changes
                
                // Look up expense account for this department
                const deptKey = dept.toLowerCase().trim();
                const expenseAcct = DEPARTMENT_EXPENSE_ACCOUNTS[deptKey] || "";
                const expenseAcctName = coaMap[expenseAcct] || "";
                
                // Determine debit/credit
                // Positive change = liability increased = expense (debit)
                // Negative change = liability decreased = expense reversal (credit)
                const debit = change > 0 ? Math.abs(change) : 0;
                const credit = change < 0 ? Math.abs(change) : 0;
                
                totalDebits += debit;
                totalCredits += credit;
                
                jeRows.push([
                    refNumber,
                    txnDate,
                    expenseAcct,
                    expenseAcctName,
                    change, // Line amount (positive or negative)
                    debit,
                    credit,
                    lineDescBase,
                    dept
                ]);
            });
            
            // Add liability offset entry (account 21540)
            // Net offset = credits - debits (opposite of expense entries)
            const netChange = totalDebits - totalCredits;
            if (Math.abs(netChange) >= 0.01) {
                const liabilityDebit = netChange < 0 ? Math.abs(netChange) : 0;
                const liabilityCredit = netChange > 0 ? Math.abs(netChange) : 0;
                const liabilityAcctName = coaMap[LIABILITY_OFFSET_ACCOUNT] || "Accrued PTO";
                
                jeRows.push([
                    refNumber,
                    txnDate,
                    LIABILITY_OFFSET_ACCOUNT,
                    liabilityAcctName,
                    -netChange, // Opposite of net expense
                    liabilityDebit,
                    liabilityCredit,
                    lineDescBase,
                    "" // No department for liability
                ]);
            }
            
            // Write to PTO_JE_Draft sheet
            let jeSheet = context.workbook.worksheets.getItemOrNullObject("PTO_JE_Draft");
            jeSheet.load("isNullObject");
            await context.sync();
            
            if (jeSheet.isNullObject) {
                jeSheet = context.workbook.worksheets.add("PTO_JE_Draft");
            } else {
                // Clear existing content
                const usedRange = jeSheet.getUsedRangeOrNullObject();
                usedRange.load("isNullObject");
                await context.sync();
                if (!usedRange.isNullObject) {
                    usedRange.clear();
                }
            }
            
            // Write all rows
            if (jeRows.length > 0) {
                const writeRange = jeSheet.getRangeByIndexes(0, 0, jeRows.length, jeHeaders.length);
                writeRange.values = jeRows;
                
                // Format headers
                const headerRange = jeSheet.getRangeByIndexes(0, 0, 1, jeHeaders.length);
                formatSheetHeaders(headerRange);
                
                // Format currency columns (E: LineAmount, F: Debit, G: Credit)
                const dataRowCount = jeRows.length - 1;
                if (dataRowCount > 0) {
                    formatCurrencyColumn(jeSheet, 4, dataRowCount, true);  // E: LineAmount (with negative)
                    formatCurrencyColumn(jeSheet, 5, dataRowCount);        // F: Debit
                    formatCurrencyColumn(jeSheet, 6, dataRowCount);        // G: Credit
                }
                
                // Auto-fit columns
                writeRange.format.autofitColumns();
            }
            
            await context.sync();
            
            // Activate the sheet and select A1
            jeSheet.activate();
            jeSheet.getRange("A1").select();
            await context.sync();
        });
        
        // Run journal summary to update totals
        await runJournalSummary();
        
    } catch (error) {
        console.error("Create JE Draft error:", error);
        window.alert(`Unable to create Journal Entry: ${error.message}`);
    } finally {
        toggleLoader(false);
    }
}

async function exportJournalDraft() {
    if (!hasExcel()) {
        window.alert("Excel is not available. Open this module inside Excel to export.");
        return;
    }
    toggleLoader(true, "Preparing JE CSV...");
    try {
        const { rows } = await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("PTO_JE_Draft");
            const range = sheet.getUsedRangeOrNullObject();
            range.load("values");
            await context.sync();
            const values = range.isNullObject ? [] : range.values || [];
            if (!values.length) {
                throw new Error("PTO_JE_Draft is empty.");
            }
            return { rows: values };
        });
        const csv = buildCsv(rows);
        downloadCsv(`pto-je-draft-${todayIso()}.csv`, csv);
    } catch (error) {
        console.error("PTO JE export:", error);
        window.alert("Unable to export the JE draft. Confirm the sheet has data.");
    } finally {
        toggleLoader(false);
    }
}

async function archiveAndReset() {
    if (!appState.workbookReady || !hasExcel()) {
        return;
    }
    toggleLoader(true, "Archiving PTO outputs...");
    try {
        await Excel.run(async (context) => {
            // capture clean rows before any clears
            const analysisRowsBefore = await getAnalysisRows(context);

            // Add archive log entry
            const archiveTable = context.workbook.tables.getItemOrNullObject("PTOArchiveLog");
            archiveTable.load("isNullObject");
            await context.sync();
            if (!archiveTable.isNullObject) {
                archiveTable.rows.add(null, [[new Date().toISOString(), "Archived PTO run", "Completed via module"]]);
            }

            // Export all PTO-related sheets defined in SS_PF_Config
            await exportPtoSheets(context);

            // Snapshot PTO_Analysis into PTO_Archive_Summary (row 2+)
            await copyAnalysisToArchiveSummary(context, analysisRowsBefore);

            // Clear data rows (row 2+) from target sheets
            const sheetsToClear = ["PTO_JE_Draft", "PTO_Analysis", "PTO_Data"];
            for (const name of sheetsToClear) {
                await clearSheetBelowHeader(context, name);
            }

            // Clear non-permanent config values
            await clearNonPermanentConfig(context);

            await context.sync();
        });

        setState({
            stepStatuses: {
                1: "pending",
                2: "pending",
                3: "pending",
                4: "pending",
                5: "pending",
                6: "complete"
            }
        });
    } catch (error) {
        console.error(error);
    } finally {
        toggleLoader(false);
    }
}

async function openSheet(sheetName) {
    if (!sheetName || !hasExcel()) {
        return;
    }
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(sheetName);
            sheet.activate();
            // Select cell A1 so user always starts at a consistent location
            sheet.getRange("A1").select();
            await context.sync();
        });
    } catch (error) {
        console.error(error);
    }
}

/**
 * Clear PTO_Data sheet to start fresh
 */
async function clearPtoData() {
    if (!hasExcel()) return;
    
    const confirmed = window.confirm("This will clear all data in PTO_Data. Are you sure?");
    if (!confirmed) return;
    
    toggleLoader(true);
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("PTO_Data");
            const usedRange = sheet.getUsedRangeOrNullObject();
            usedRange.load("rowCount");
            await context.sync();
            
            if (!usedRange.isNullObject && usedRange.rowCount > 1) {
                // Keep header row, clear everything else
                const dataRange = sheet.getRangeByIndexes(1, 0, usedRange.rowCount - 1, 20);
                dataRange.clear(Excel.ClearApplyTo.contents);
                await context.sync();
            }
            
            sheet.activate();
            sheet.getRange("A1").select();
            await context.sync();
        });
        window.alert("PTO_Data cleared successfully. You can now paste new data.");
    } catch (error) {
        console.error("Clear PTO_Data error:", error);
        window.alert(`Failed to clear PTO_Data: ${error.message}`);
    } finally {
        toggleLoader(false);
    }
}

/**
 * Open a reference data sheet (creates if doesn't exist)
 */
async function openReferenceSheet(sheetName) {
    if (!sheetName || !hasExcel()) {
        return;
    }
    
    const defaultHeaders = {
        "SS_Employee_Roster": ["Employee", "Department", "Pay_Rate", "Status", "Hire_Date"],
        "SS_Chart_of_Accounts": ["Account_Number", "Account_Name", "Type", "Category"]
    };
    
    try {
        await Excel.run(async (context) => {
            let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
            sheet.load("isNullObject");
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
            }
            
            sheet.activate();
            sheet.getRange("A1").select();
            await context.sync();
        });
    } catch (error) {
        console.error("Error opening reference sheet:", error);
    }
}

/**
 * Unhide system sheets (SS_* prefix) for configuration access
 */
async function unhideSystemSheets() {
    if (!hasExcel()) {
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

async function writeDatasetToSheet(context, sheetName, columns, rows) {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const usedRange = sheet.getUsedRangeOrNullObject();
    usedRange.load("address");
    await context.sync();
    if (!usedRange.isNullObject) {
        usedRange.clear();
    }
    const data = [
        columns.map((col) => col.header),
        ...rows.map((row) => columns.map((col) => row[col.key]))
    ];
    const range = sheet.getRangeByIndexes(0, 0, data.length, data[0]?.length || 1);
    range.values = data;
    range.format.autofitColumns();
    await context.sync();
}

async function clearSheetBelowHeader(context, sheetName) {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const usedRange = sheet.getUsedRangeOrNullObject();
    usedRange.load("rowCount");
    await context.sync();
    if (usedRange.isNullObject || usedRange.rowCount <= 1) return;
    const dataRange = sheet.getRangeByIndexes(1, 0, usedRange.rowCount - 1, usedRange.columnCount);
    dataRange.clear();
}

async function getAnalysisRows(context) {
    const sheet = context.workbook.worksheets.getItemOrNullObject("PTO_Analysis");
    sheet.load("isNullObject");
    await context.sync();
    if (sheet.isNullObject) return [];
    const range = sheet.getUsedRangeOrNullObject();
    range.load("values");
    await context.sync();
    const values = range.isNullObject ? [] : range.values || [];
    if (values.length <= 1) return [];
    return values.slice(1);
}

async function writeRowsStartingAt(context, sheetName, rows) {
    if (!rows.length) return;
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRangeByIndexes(1, 0, rows.length, rows[0].length);
    range.values = rows;
    range.format.autofitColumns();
}

async function clearNonPermanentConfig(context) {
    const table = context.workbook.tables.getItemOrNullObject(CONFIG_TABLES[0]);
    await context.sync();
    if (table.isNullObject) return;
    const body = table.getDataBodyRange();
    const header = table.getHeaderRowRange();
    body.load("values");
    header.load("values");
    await context.sync();
    const headers = header.values[0] || [];
    const normalizedHeaders = headers.map((h) => normalizeName(h));
    const idx = {
        field: normalizedHeaders.findIndex((h) => h === "field" || h === "field name" || h === "setting"),
        permanent: normalizedHeaders.findIndex((h) => h === "permanent" || h === "persist"),
        value: normalizedHeaders.findIndex((h) => h === "value" || h === "setting value")
    };
    if (idx.field === -1 || idx.value === -1 || idx.permanent === -1) return;
    (body.values || []).forEach((row, rowIndex) => {
        const permanent = String(row[idx.permanent] ?? "").trim().toLowerCase();
        const shouldClear = permanent !== "y" && permanent !== "yes" && permanent !== "true" && permanent !== "t" && permanent !== "1";
        if (shouldClear) {
            body.getCell(rowIndex, idx.value).values = [[""]];
        }
    });
}

async function exportPtoSheets(context) {
    const sheetNames = await getPtoSheetNames(context);
    if (!sheetNames.length) return;
    const chunks = [];
    for (const name of sheetNames) {
        try {
            const sheet = context.workbook.worksheets.getItemOrNullObject(name);
            sheet.load("isNullObject");
            await context.sync();
            if (sheet.isNullObject) continue;
            const range = sheet.getUsedRangeOrNullObject();
            range.load("values");
            await context.sync();
            const values = range.isNullObject ? [] : range.values || [];
            const csv = values
                .map((row) => row.map((cell) => `"${String(cell ?? "").replace(/"/g, '""')}"`).join(","))
                .join("\n");
            chunks.push(`# Sheet: ${name}\n${csv}`);
        } catch (error) {
            console.warn("PTO: unable to export sheet", name, error);
        }
    }
    if (!chunks.length) return;
    const fileName = `${new Date().toISOString().slice(0, 10)} - Tai Software PTO Accrual.xlsx`;
    // rudimentary multi-tab export: CSV sections inside a .xlsx extension to break links
    const blob = new Blob([chunks.join("\n\n")], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

async function getPtoSheetNames(context) {
    const configSheet = context.workbook.worksheets.getItemOrNullObject("SS_PF_Config");
    configSheet.load("isNullObject");
    await context.sync();
    if (configSheet.isNullObject) return [];
    const usedRange = configSheet.getUsedRangeOrNullObject();
    usedRange.load("values");
    await context.sync();
    const values = usedRange.isNullObject ? [] : usedRange.values || [];
    if (!values.length) return [];
    const headers = (values[0] || []).map((h) => normalizeName(h));
    const idx = {
        category: headers.findIndex((h) => h === "category"),
        module: headers.findIndex((h) => h === "module"),
        field: headers.findIndex((h) => h === "field"),
        value: headers.findIndex((h) => h === "value")
    };
    if (idx.category === -1 || idx.field === -1) return [];
    const targetModule = normalizeName(MODULE_NAME);
    const names = values
        .slice(1)
        .filter((row) => {
            const category = normalizeName(row[idx.category]);
            const module = idx.module >= 0 ? normalizeName(row[idx.module]) : "";
            const moduleValue = idx.value >= 0 ? normalizeName(row[idx.value]) : "";
            return category === "tab-structure" && (module === targetModule || moduleValue === "pto-accrual");
        })
        .map((row) => String(row[idx.field] ?? "").trim())
        .filter(Boolean);
    return Array.from(new Set(names));
}

/**
 * Copy PTO_Analysis to PTO_Archive_Summary.
 * Only retains the MOST RECENT period - older data is replaced.
 * e.g., When archiving 11/30, this overwrites any existing 10/31 data.
 */
async function copyAnalysisToArchiveSummary(context, analysisRows = null) {
    const analysisSheet = context.workbook.worksheets.getItemOrNullObject("PTO_Analysis");
    const archiveSheet = context.workbook.worksheets.getItemOrNullObject("PTO_Archive_Summary");
    analysisSheet.load("isNullObject");
    archiveSheet.load("isNullObject");
    await context.sync();
    if (analysisSheet.isNullObject || archiveSheet.isNullObject) return;
    
    // Get analysis data (use passed rows or fetch fresh)
    let values = analysisRows;
    if (!values || !values.length) {
        const analysisRange = analysisSheet.getUsedRangeOrNullObject();
        analysisRange.load("values");
        await context.sync();
        values = analysisRange.isNullObject ? [] : analysisRange.values || [];
    }
    if (!values.length) return;
    
    // Clear existing archive data (only retain most recent period)
    const existingRange = archiveSheet.getUsedRangeOrNullObject();
    existingRange.load("isNullObject");
    await context.sync();
    if (!existingRange.isNullObject) {
        existingRange.clear();
    }
    
    // Write current period's analysis as the new archive
    const archiveRange = archiveSheet.getRangeByIndexes(0, 0, values.length, values[0].length);
    archiveRange.values = values;
    archiveRange.format.autofitColumns();
    
    // Reset selection to A1
    archiveSheet.getRange("A1").select();
    await context.sync();
}

function getConfigValue(fieldName) {
    const key = String(fieldName ?? "").trim();
    return configState.values?.[key] ?? "";
}

/**
 * Get reviewer name with fallback chain:
 * 1. Step-specific reviewer (from step config fields)
 * 2. Module default (PTO_Reviewer_Name in SS_PF_Config)
 * 3. Shared default (Default_Reviewer in SS_PF_Config)
 */
function getReviewerWithFallback(stepReviewer) {
    // Step-specific
    if (stepReviewer) return stepReviewer;
    
    // Module default
    const moduleDefault = getConfigValue(PTO_CONFIG_FIELDS.reviewerName);
    if (moduleDefault) return moduleDefault;
    
    // Shared default (from cached SS_PF_Config) - check new + legacy names
    if (window.PrairieForge?._sharedConfigCache) {
        const sharedDefault = window.PrairieForge._sharedConfigCache.get("SS_Default_Reviewer") 
            || window.PrairieForge._sharedConfigCache.get("Default_Reviewer");
        if (sharedDefault) return sharedDefault;
    }
    
    return "";
}

function scheduleConfigWrite(fieldName, value, options = {}) {
    const key = String(fieldName ?? "").trim();
    if (!key) return;
    configState.values[key] = value ?? "";
    const delay = options.debounceMs ?? 0;
    if (!delay) {
        const existing = pendingConfigWrites.get(key);
        if (existing) clearTimeout(existing);
        pendingConfigWrites.delete(key);
        void saveConfigValue(key, value ?? "", CONFIG_TABLES);
        return;
    }
    if (pendingConfigWrites.has(key)) {
        clearTimeout(pendingConfigWrites.get(key));
    }
    const timer = setTimeout(() => {
        pendingConfigWrites.delete(key);
        void saveConfigValue(key, value ?? "", CONFIG_TABLES);
    }, delay);
    pendingConfigWrites.set(key, timer);
}

function normalizeName(value) {
    return String(value ?? "").trim().toLowerCase();
}

function columnLetterFromIndex(index) {
    let dividend = index + 1;
    let columnName = "";
    while (dividend > 0) {
        const modulo = (dividend - 1) % 26;
        columnName = String.fromCharCode(65 + modulo) + columnName;
        dividend = Math.floor((dividend - modulo) / 26);
    }
    return columnName;
}

function toggleLoader(show, message = "Working...") {
    // Suppress loader overlay for smoother transitions between views.
    const overlay = document.getElementById(LOADER_ID);
    if (overlay) overlay.style.display = "none";
}

function bootstrapModule() {
    init();
}

if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(() => bootstrapModule()).catch(() => bootstrapModule());
} else {
    bootstrapModule();
}

function getStepConfig(stepId) {
    return configState.steps[stepId] || { notes: "", reviewer: "", signOffDate: "" };
}

function getFieldNames(stepId) {
    return STEP_CONFIG_FIELDS[stepId] || {};
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

async function saveStepField(stepId, key, value) {
    const current = configState.steps[stepId] || { notes: "", reviewer: "", signOffDate: "" };
    current[key] = value;
    configState.steps[stepId] = current;
    const fieldNames = getFieldNames(stepId);
    const targetField =
        key === "notes" ? fieldNames.note : key === "reviewer" ? fieldNames.reviewer : fieldNames.signOff;
    if (!targetField) return;
    if (!hasExcelRuntime()) return;
    try {
        await saveConfigValue(targetField, value, CONFIG_TABLES);
    } catch (error) {
        console.warn("PTO: unable to save field", targetField, error);
    }
}

async function toggleNotePermanent(stepId, isPermanent) {
    configState.permanents[stepId] = isPermanent;
    const fieldNames = getFieldNames(stepId);
    if (!fieldNames?.note) return;
    if (!hasExcelRuntime()) return;
    try {
        await Excel.run(async (context) => {
            const table = context.workbook.tables.getItemOrNullObject(CONFIG_TABLES[0]);
            await context.sync();
            if (table.isNullObject) return;
            const body = table.getDataBodyRange();
            const header = table.getHeaderRowRange();
            body.load("values");
            header.load("values");
            await context.sync();
            const headers = header.values[0] || [];
            const normalizedHeaders = headers.map((h) => String(h || "").trim().toLowerCase());
            const idx = {
                field: normalizedHeaders.findIndex((h) => h === "field" || h === "field name" || h === "setting"),
                permanent: normalizedHeaders.findIndex((h) => h === "permanent" || h === "persist"),
                value: normalizedHeaders.findIndex((h) => h === "value" || h === "setting value"),
                type: normalizedHeaders.findIndex((h) => h === "type" || h === "category"),
                title: normalizedHeaders.findIndex((h) => h === "title" || h === "display name")
            };
            if (idx.field === -1) return;
            const rows = body.values || [];
            const targetIndex = rows.findIndex(
                (row) => String(row[idx.field] || "").trim() === fieldNames.note
            );
            if (targetIndex >= 0) {
                if (idx.permanent >= 0) {
                    body.getCell(targetIndex, idx.permanent).values = [[isPermanent ? "Y" : "N"]];
                }
            } else {
                // create row with permanent flag if missing
                const newRow = new Array(headers.length).fill("");
                if (idx.type >= 0) newRow[idx.type] = "Other";
                if (idx.title >= 0) newRow[idx.title] = "";
                newRow[idx.field] = fieldNames.note;
                if (idx.permanent >= 0) newRow[idx.permanent] = isPermanent ? "Y" : "N";
                if (idx.value >= 0) newRow[idx.value] = configState.steps[stepId]?.notes || "";
                table.rows.add(null, [newRow]);
            }
            await context.sync();
        });
    } catch (error) {
        console.warn("PTO: unable to update permanent flag", error);
    }
}

async function saveCompletionFlag(stepId, isComplete) {
    const fieldName = STEP_COMPLETE_FIELDS[stepId];
    if (!fieldName) return;
    configState.completes[stepId] = isComplete ? "Y" : "";
    if (!hasExcelRuntime()) return;
    try {
        await saveConfigValue(fieldName, isComplete ? "Y" : "", CONFIG_TABLES);
    } catch (error) {
        console.warn("PTO: unable to save completion flag", fieldName, error);
    }
}

function updateActionToggleState(button, isActive) {
    if (!button) return;
    button.classList.toggle("is-active", isActive);
    button.setAttribute("aria-pressed", String(isActive));
}

/**
 * Get current step completion status for sequential validation
 * @returns {Object} Map of step IDs to boolean completion status
 */
function getStepCompletionStatus() {
    const status = {};
    Object.keys(STEP_CONFIG_FIELDS).forEach(stepIdStr => {
        const id = parseInt(stepIdStr, 10);
        // A step is complete if it has a sign-off date OR is explicitly marked complete
        const hasSignOff = Boolean(configState.steps[id]?.signOffDate);
        const isMarkedComplete = Boolean(configState.completes[id]);
        status[id] = hasSignOff || isMarkedComplete;
    });
    return status;
}

function bindSignoffToggle(stepId, { buttonId, inputId, canActivate = null, onComplete = null }) {
    const button = document.getElementById(buttonId);
    if (!button) return;
    const input = document.getElementById(inputId);
    const initial =
        Boolean(configState.steps[stepId]?.signOffDate) || Boolean(configState.completes[stepId]);
    updateActionToggleState(button, initial);
    button.addEventListener("click", () => {
        // Check sequential completion (only when trying to activate, not deactivate)
        const isCurrentlyActive = button.classList.contains("is-active");
        if (!isCurrentlyActive && stepId > 0) {
            const completionStatus = getStepCompletionStatus();
            const { canComplete, message } = canCompleteStep(stepId, completionStatus);
            if (!canComplete) {
                showBlockedToast(message);
                return;
            }
        }
        
        if (typeof canActivate === "function" && !canActivate()) return;
        const next = !button.classList.contains("is-active");
        updateActionToggleState(button, next);
        if (input) {
            input.value = next ? todayIso() : "";
            saveStepField(stepId, "signOffDate", input.value);
        }
        saveCompletionFlag(stepId, next);
        if (next && typeof onComplete === "function") {
            onComplete();
        }
    });
}

function escapeHtml(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
}

function escapeAttr(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

function parseBooleanFlag(value) {
    const normalized = String(value ?? "").trim().toLowerCase();
    return normalized === "true" || normalized === "y" || normalized === "yes" || normalized === "1";
}

function parseBooleanStrict(value) {
    const normalized = String(value ?? "").trim().toLowerCase();
    return normalized === "true" || normalized === "y" || normalized === "yes" || normalized === "1";
}

function parseDateInput(value) {
    if (!value) return null;
    const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(value));
    if (!match) return null;
    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);
    if (!year || !month || !day) return null;
    return { year, month, day };
}

function formatDateInput(value) {
    if (!value) return "";
    const parts = parseDateInput(value);
    if (!parts) return "";
    const { year, month, day } = parts;
    return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function formatDateFromDate(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
}

function deriveAccountingPeriod(payrollDate) {
    const parts = parseDateInput(payrollDate);
    if (!parts) return "";
    const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    return `${monthNames[parts.month - 1]} ${parts.year}`;
}

function deriveJournalId(payrollDate) {
    const parts = parseDateInput(payrollDate);
    if (!parts) return "";
    return `PTO-AUTO-${parts.year}-${String(parts.month).padStart(2, "0")}-${String(parts.day).padStart(2, "0")}`;
}

function todayIso() {
    const now = new Date();
    const y = now.getFullYear();
    const m = String(now.getMonth() + 1).padStart(2, "0");
    const d = String(now.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
}

function parsePermanentFlag(value) {
    const normalized = String(value ?? "").trim().toLowerCase();
    return normalized === "y" || normalized === "yes" || normalized === "true" || normalized === "t" || normalized === "1";
}

function coerceTimestamp(value) {
    if (value instanceof Date) return value.getTime();
    if (typeof value === "number") {
        const date = convertExcelSerialDate(value);
        return date ? date.getTime() : null;
    }
    const parsed = new Date(value);
    return Number.isNaN(parsed.getTime()) ? null : parsed.getTime();
}

function convertExcelSerialDate(serial) {
    if (!Number.isFinite(serial)) return null;
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    return new Date(excelEpoch.getTime() + serial * 24 * 60 * 60 * 1000);
}

function persistConfigBasics() {
    const getVal = (id) => document.getElementById(id)?.value?.trim() || "";
    const fields = [
        { id: "config-payroll-date", field: PTO_CONFIG_FIELDS.payrollDate },
        { id: "config-accounting-period", field: PTO_CONFIG_FIELDS.accountingPeriod },
        { id: "config-journal-id", field: PTO_CONFIG_FIELDS.journalEntryId },
        { id: "config-company-name", field: PTO_CONFIG_FIELDS.companyName },
        { id: "config-payroll-provider", field: PTO_CONFIG_FIELDS.payrollProvider },
        { id: "config-accounting-link", field: PTO_CONFIG_FIELDS.accountingSoftware },
        { id: "config-user-name", field: PTO_CONFIG_FIELDS.reviewerName }
    ];
    fields.forEach(({ id, field }) => {
        const value = getVal(id);
        if (!field) return;
        scheduleConfigWrite(field, value);
    });
}

function findColumnIndex(headers, keywords = []) {
    const normalizedKeywords = keywords.map((k) => normalizeName(k));
    return headers.findIndex((header) =>
        normalizedKeywords.some((keyword) => header.includes(keyword))
    );
}

function renderHeadcountStep(detail) {
    const stepFields = getStepConfig(2);
    const stepNotes = stepFields?.notes || "";
    const stepNotesPermanent = Boolean(configState.permanents[2]);
    const stepReviewer = getReviewerWithFallback(stepFields?.reviewer);
    const stepSignOff = stepFields?.signOffDate || "";
    const stepComplete = Boolean(parseBooleanStrict(configState.completes[2]) || stepSignOff);
    const roster = headcountState.roster || {};
    const hasRun = headcountState.hasAnalyzed;
    const rosterDiff = headcountState.roster?.difference ?? 0;
    const requiresNotes = !headcountState.skipAnalysis && Math.abs(rosterDiff) > 0;
    const rosterCount = roster.rosterCount ?? 0;
    const payrollCount = roster.payrollCount ?? 0;
    const diffValue = roster.difference ?? payrollCount - rosterCount;
    const mismatchList = Array.isArray(roster.mismatches) ? roster.mismatches.filter(Boolean) : [];
    
    // Status banner
    let statusBanner = "";
    if (headcountState.loading) {
        statusBanner = window.PrairieForge?.renderStatusBanner?.({
            type: "info",
            message: "Analyzing headcount…",
            escapeHtml
        }) || "";
    } else if (headcountState.lastError) {
        statusBanner = window.PrairieForge?.renderStatusBanner?.({
            type: "error",
            message: headcountState.lastError,
            escapeHtml
        }) || "";
    }
    
    // Build check rows (circle + pill format)
    const renderCheckRow = (label, desc, value, isMatch) => {
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
    
    const checkRowsHtml = `
        ${renderCheckRow("SS_Employee_Roster count", "Active employees in roster", rosterCount, true)}
        ${renderCheckRow("PTO_Data count", "Unique employees in PTO data", payrollCount, true)}
        ${renderCheckRow("Difference", "Should be zero", diffValue, diffValue === 0)}
    `;
    
    // Mismatch section
    const mismatchSection =
        mismatchList.length && !headcountState.skipAnalysis && hasRun
            ? window.PrairieForge.renderMismatchTiles({
                  mismatches: mismatchList,
                  label: "Employees Driving the Difference",
                  sourceLabel: "Roster",
                  targetLabel: "PTO Data",
                  escapeHtml: escapeHtml
              })
            : "";
    
    return `
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step ${detail.id}</p>
            <h2 class="pf-hero-title">${escapeHtml(detail.title)}</h2>
            <p class="pf-hero-copy">${escapeHtml(detail.summary || "")}</p>
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
                    <p class="pf-config-subtext">Compare employee roster against PTO data to identify discrepancies.</p>
                </div>
                <div class="pf-signoff-action">
                    ${renderLabeledButton(`<button type="button" class="pf-action-toggle pf-clickable" id="headcount-run-btn" title="Run headcount analysis">${CALCULATOR_ICON_SVG}</button>`, "Run")}
                    ${renderLabeledButton(`<button type="button" class="pf-action-toggle pf-clickable" id="headcount-refresh-btn" title="Refresh headcount analysis">${REFRESH_ICON_SVG}</button>`, "Refresh")}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Headcount Comparison</h3>
                    <p class="pf-config-subtext">Verify roster and payroll data align before proceeding.</p>
                </div>
                ${statusBanner}
                <div class="pf-je-checks-container">
                    ${checkRowsHtml}
                </div>
                ${mismatchSection}
            </article>
            ${renderInlineNotes({
                textareaId: "step-notes-input",
                value: stepNotes,
                permanentId: "step-notes-lock-2",
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
        </section>
    `;
}

/**
 * Render unified Apple-inspired Data Readiness card
 * Combines completeness check pills + missing pay rate resolution
 */
function renderDataReadinessCard() {
    const check = analysisState.completenessCheck || {};
    const missing = analysisState.missingPayRates || [];
    
    // Completeness check fields with descriptions
    const fields = [
        { key: "accrualRate", label: "Accrual Rate", desc: "∑ PTO_Data = ∑ PTO_Analysis" },
        { key: "carryOver", label: "Carry Over", desc: "∑ PTO_Data = ∑ PTO_Analysis" },
        { key: "ytdAccrued", label: "YTD Accrued", desc: "∑ PTO_Data = ∑ PTO_Analysis" },
        { key: "ytdUsed", label: "YTD Used", desc: "∑ PTO_Data = ∑ PTO_Analysis" },
        { key: "balance", label: "Balance", desc: "∑ PTO_Data = ∑ PTO_Analysis" }
    ];
    
    // Calculate overall status
    const allChecked = fields.every(f => check[f.key] !== null && check[f.key] !== undefined);
    const allPassed = allChecked && fields.every(f => check[f.key]?.match);
    const hasMissingRates = missing.length > 0;
    
    // Build check rows (circle + pill format like JE tab)
    const renderCheckRow = (field) => {
        const result = check[field.key];
        const pending = result === null || result === undefined;
        let circleHtml;
        
        if (pending) {
            circleHtml = `<span class="pf-je-check-circle pf-je-circle--pending"></span>`;
        } else if (result.match) {
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
                <span class="pf-je-check-desc-pill">${escapeHtml(field.label)}: ${escapeHtml(field.desc)}</span>
            </div>
        `;
    };
    
    const checkRowsHtml = fields.map(f => renderCheckRow(f)).join("");
    
    // Missing pay rate section
    let missingSection = "";
    if (hasMissingRates) {
        const employee = missing[0];
        const remainingCount = missing.length - 1;
        
        missingSection = `
            <div class="pf-readiness-divider"></div>
            <div class="pf-readiness-issue">
                <div class="pf-readiness-issue-header">
                    <span class="pf-readiness-issue-badge">Action Required</span>
                    <span class="pf-readiness-issue-title">Missing Pay Rate</span>
                </div>
                <p class="pf-readiness-issue-desc">
                    Enter hourly rate for <strong>${escapeHtml(employee.name)}</strong> to calculate liability
                </p>
                <div class="pf-readiness-input-row">
                    <div class="pf-readiness-input-field">
                        <span class="pf-readiness-input-prefix">$</span>
                        <input type="number" 
                               id="payrate-input" 
                               class="pf-readiness-input" 
                               placeholder="0.00" 
                               step="0.01"
                               min="0"
                               data-employee="${escapeAttr(employee.name)}"
                               data-row="${employee.rowIndex}">
                    </div>
                    <button type="button" class="pf-readiness-btn pf-readiness-btn--secondary" id="payrate-ignore-btn">
                        Skip
                    </button>
                    <button type="button" class="pf-readiness-btn pf-readiness-btn--primary" id="payrate-save-btn">
                        Save
                    </button>
                </div>
                ${remainingCount > 0 ? `<p class="pf-readiness-remaining">${remainingCount} more employee${remainingCount > 1 ? "s" : ""} need pay rates</p>` : ""}
            </div>
        `;
    }
    
    return `
        <article class="pf-step-card pf-step-detail pf-config-card" id="data-readiness-card">
            <div class="pf-config-head">
                <h3>Data Completeness</h3>
                <p class="pf-config-subtext">Quick check that all your data transferred correctly.</p>
            </div>
            <div class="pf-je-checks-container">
                ${checkRowsHtml}
            </div>
            ${missingSection}
        </article>
    `;
}

function renderDataQualityStep(detail) {
    const stepFields = getStepConfig(3);
    const notesPermanent = Boolean(configState.permanents[3]);
    const stepReviewer = getReviewerWithFallback(stepFields?.reviewer);
    const stepSignOff = stepFields?.signOffDate || "";
    const stepComplete = Boolean(parseBooleanStrict(configState.completes[3]) || stepSignOff);
    
    // Build quality check results display
    const hasRun = dataQualityState.hasRun;
    const { balanceIssues, zeroBalances, accrualOutliers, totalEmployees } = dataQualityState;
    
    // Status banner
    let statusBanner = "";
    if (dataQualityState.loading) {
        statusBanner = window.PrairieForge?.renderStatusBanner?.({
            type: "info",
            message: "Analyzing data quality...",
            escapeHtml
        }) || "";
    } else if (hasRun) {
        const criticalCount = balanceIssues.length;
        const warningCount = accrualOutliers.length + zeroBalances.length;
        
        if (criticalCount > 0) {
            statusBanner = window.PrairieForge?.renderStatusBanner?.({
                type: "error",
                title: `${criticalCount} Balance Issue${criticalCount > 1 ? "s" : ""} Found`,
                message: "Review the issues below. Fix in PTO_Data and re-run, or acknowledge to continue.",
                escapeHtml
            }) || "";
        } else if (warningCount > 0) {
            statusBanner = window.PrairieForge?.renderStatusBanner?.({
                type: "warning",
                title: "No Critical Issues",
                message: `${warningCount} informational item${warningCount > 1 ? "s" : ""} to review (see below).`,
                escapeHtml
            }) || "";
        } else {
            statusBanner = window.PrairieForge?.renderStatusBanner?.({
                type: "success",
                title: "Data Quality Passed",
                message: `${totalEmployees} employee${totalEmployees !== 1 ? "s" : ""} checked — no anomalies found.`,
                escapeHtml
            }) || "";
        }
    }
    
    // Build issue cards
    const issueCards = [];
    
    if (hasRun && balanceIssues.length > 0) {
        issueCards.push(`
            <div class="pf-quality-issue pf-quality-issue--critical">
                <div class="pf-quality-issue-header">
                    <span class="pf-quality-issue-icon">⚠️</span>
                    <span class="pf-quality-issue-title">Balance Issues (${balanceIssues.length})</span>
                </div>
                <ul class="pf-quality-issue-list">
                    ${balanceIssues.slice(0, 5).map(e => 
                        `<li><strong>${escapeHtml(e.name)}</strong>: ${escapeHtml(e.issue)}</li>`
                    ).join("")}
                    ${balanceIssues.length > 5 ? `<li class="pf-quality-more">+${balanceIssues.length - 5} more</li>` : ""}
                </ul>
            </div>
        `);
    }
    
    if (hasRun && accrualOutliers.length > 0) {
        issueCards.push(`
            <div class="pf-quality-issue pf-quality-issue--warning">
                <div class="pf-quality-issue-header">
                    <span class="pf-quality-issue-icon">📊</span>
                    <span class="pf-quality-issue-title">High Accrual Rates (${accrualOutliers.length})</span>
                </div>
                <ul class="pf-quality-issue-list">
                    ${accrualOutliers.slice(0, 5).map(e => 
                        `<li><strong>${escapeHtml(e.name)}</strong>: ${e.accrualRate.toFixed(2)} hrs/period</li>`
                    ).join("")}
                    ${accrualOutliers.length > 5 ? `<li class="pf-quality-more">+${accrualOutliers.length - 5} more</li>` : ""}
                </ul>
            </div>
        `);
    }
    
    if (hasRun && zeroBalances.length > 0) {
        issueCards.push(`
            <div class="pf-quality-issue pf-quality-issue--info">
                <div class="pf-quality-issue-header">
                    <span class="pf-quality-issue-icon">ℹ️</span>
                    <span class="pf-quality-issue-title">Zero Balances (${zeroBalances.length})</span>
                </div>
                <ul class="pf-quality-issue-list">
                    ${zeroBalances.slice(0, 5).map(e => 
                        `<li><strong>${escapeHtml(e.name)}</strong></li>`
                    ).join("")}
                    ${zeroBalances.length > 5 ? `<li class="pf-quality-more">+${zeroBalances.length - 5} more</li>` : ""}
                </ul>
            </div>
        `);
    }
    
    return `
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step ${detail.id}</p>
            <h2 class="pf-hero-title">${escapeHtml(detail.title)}</h2>
            <p class="pf-hero-copy">${escapeHtml(detail.summary || "")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Run Quality Check</h3>
                    <p class="pf-config-subtext">Scan your imported data for common errors before proceeding.</p>
                </div>
                ${statusBanner}
                <div class="pf-signoff-action">
                    ${renderLabeledButton(`<button type="button" class="pf-action-toggle pf-clickable" id="quality-run-btn" title="Run data quality checks">${CALCULATOR_ICON_SVG}</button>`, "Run")}
                </div>
            </article>
            ${issueCards.length > 0 ? `
                <article class="pf-step-card pf-step-detail">
                    <div class="pf-config-head">
                        <h3>Issues Found</h3>
                        <p class="pf-config-subtext">Fix issues in PTO_Data and re-run, or acknowledge to continue.</p>
                    </div>
                    <div class="pf-quality-issues-grid">
                        ${issueCards.join("")}
                    </div>
                    <div class="pf-quality-actions-bar">
                        ${dataQualityState.acknowledged 
                            ? `<p class="pf-quality-actions-hint"><span class="pf-acknowledged-badge">✓ Issues Acknowledged</span></p>` 
                            : ""}
                        <div class="pf-signoff-action">
                            ${renderLabeledButton(`<button type="button" class="pf-action-toggle pf-clickable" id="quality-refresh-btn" title="Re-run quality checks">${REFRESH_ICON_SVG}</button>`, "Refresh")}
                            ${!dataQualityState.acknowledged ? renderLabeledButton(`<button type="button" class="pf-action-toggle pf-clickable" id="quality-acknowledge-btn" title="Acknowledge issues and continue">${CHECK_ICON_SVG}</button>`, "Continue") : ""}
                        </div>
                    </div>
                </article>
            ` : ""}
            ${renderInlineNotes({
                textareaId: "step-notes-3",
                value: stepFields?.notes || "",
                permanentId: "step-notes-lock-3",
                isPermanent: notesPermanent,
                hintId: "",
                saveButtonId: "step-notes-save-3"
            })}
            ${renderSignoff({
                reviewerInputId: "step-reviewer-3",
                reviewerValue: stepReviewer,
                signoffInputId: "step-signoff-3",
                signoffValue: stepSignOff,
                isComplete: stepComplete,
                saveButtonId: "step-signoff-save-3",
                completeButtonId: "step-signoff-toggle-3"
            })}
        </section>
    `;
}

function renderAccrualReviewStep(detail) {
    const stepFields = getStepConfig(4);
    const notesPermanent = Boolean(configState.permanents[4]);
    const stepReviewer = getReviewerWithFallback(stepFields?.reviewer);
    const stepSignOff = stepFields?.signOffDate || "";
    const stepComplete = Boolean(parseBooleanStrict(configState.completes[4]) || stepSignOff);
    
    return `
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step ${detail.id}</p>
            <h2 class="pf-hero-title">${escapeHtml(detail.title)}</h2>
            <p class="pf-hero-copy">${escapeHtml(detail.summary || "")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Run Analysis</h3>
                    <p class="pf-config-subtext">Calculate liabilities and compare against last period.</p>
                </div>
                <div class="pf-signoff-action">
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle pf-clickable" id="analysis-run-btn" title="Run analysis and checks">${CALCULATOR_ICON_SVG}</button>`,
                        "Run"
                    )}
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle pf-clickable" id="analysis-refresh-btn" title="Refresh data from PTO_Data">${REFRESH_ICON_SVG}</button>`,
                        "Refresh"
                    )}
                </div>
            </article>
            ${renderDataReadinessCard()}
            ${renderInlineNotes({
                textareaId: "step-notes-4",
                value: stepFields?.notes || "",
                permanentId: "step-notes-lock-4",
                isPermanent: notesPermanent,
                hintId: "",
                saveButtonId: "step-notes-save-4"
            })}
            ${renderSignoff({
                reviewerInputId: "step-reviewer-4",
                reviewerValue: stepReviewer,
                signoffInputId: "step-signoff-4",
                signoffValue: stepSignOff,
                isComplete: stepComplete,
                saveButtonId: "step-signoff-save-4",
                completeButtonId: "step-signoff-toggle-4"
            })}
        </section>
    `;
}

function renderJournalStep(detail) {
    const stepFields = getStepConfig(5);
    const notesPermanent = Boolean(configState.permanents[5]);
    const stepReviewer = getReviewerWithFallback(stepFields?.reviewer);
    const stepSignOff = stepFields?.signOffDate || "";
    const stepComplete = Boolean(parseBooleanStrict(configState.completes[5]) || stepSignOff);
    const statusNote = journalState.lastError
        ? `<p class="pf-step-note">${escapeHtml(journalState.lastError)}</p>`
        : "";
    
    // Build validation check rows: circle icon on left, description pill on right
    const hasRun = journalState.validationRun;
    const issues = journalState.issues || [];
    
    // Define check descriptions (what's being calculated)
    const checkDefinitions = [
        { key: "Debits = Credits", desc: "∑ Debit column = ∑ Credit column" },
        { key: "Line Amounts Sum to Zero", desc: "∑ Line Amount = $0.00" },
        { key: "JE Matches Analysis Total", desc: "∑ Expense line amounts = ∑ PTO_Analysis Change" }
    ];
    
    const renderCheckRow = (def) => {
        const issue = issues.find(i => i.check === def.key);
        const pending = !hasRun;
        let circleHtml;
        
        if (pending) {
            // Empty circle - waiting for validation
            circleHtml = `<span class="pf-je-check-circle pf-je-circle--pending"></span>`;
        } else if (issue?.passed) {
            // Checkmark in circle
            circleHtml = `<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`;
        } else {
            // X in circle
            circleHtml = `<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`;
        }
        
        return `
            <div class="pf-je-check-row">
                ${circleHtml}
                <span class="pf-je-check-desc-pill">${escapeHtml(def.desc)}</span>
            </div>
        `;
    };
    
    const checkRows = checkDefinitions.map(def => renderCheckRow(def)).join("");
    
    // Build issues card if there are failures
    const failedIssues = issues.filter(i => !i.passed);
    let issuesCard = "";
    if (hasRun && failedIssues.length > 0) {
        issuesCard = `
            <article class="pf-step-card pf-step-detail pf-je-issues-card">
                <div class="pf-config-head">
                    <h3>⚠️ Issues Identified</h3>
                    <p class="pf-config-subtext">The following checks did not pass:</p>
                </div>
                <ul class="pf-je-issues-list">
                    ${failedIssues.map(i => `<li><strong>${escapeHtml(i.check)}:</strong> ${escapeHtml(i.detail)}</li>`).join("")}
                </ul>
            </article>
        `;
    }
    
    return `
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step ${detail.id}</p>
            <h2 class="pf-hero-title">${escapeHtml(detail.title)}</h2>
            <p class="pf-hero-copy">${escapeHtml(detail.summary || "")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Generate Journal Entry</h3>
                    <p class="pf-config-subtext">Create a balanced JE from your imported PTO data, grouped by department.</p>
                </div>
                <div class="pf-signoff-action">
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle pf-clickable" id="je-create-btn" title="Generate journal entry from PTO_Analysis">${TABLE_ICON_SVG}</button>`,
                        "Generate"
                    )}
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle pf-clickable" id="je-run-btn" title="Refresh validation checks">${REFRESH_ICON_SVG}</button>`,
                        "Refresh"
                    )}
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle pf-clickable" id="je-export-btn" title="Export journal draft as CSV">${DOWNLOAD_ICON_SVG}</button>`,
                        "Export"
                    )}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Validation Checks</h3>
                    <p class="pf-config-subtext">These checks run automatically after generating your JE.</p>
                </div>
                ${statusNote}
                <div class="pf-je-checks-container">
                    ${checkRows}
                </div>
            </article>
            ${issuesCard}
            ${renderInlineNotes({
                textareaId: "step-notes-5",
                value: stepFields?.notes || "",
                permanentId: "step-notes-lock-5",
                isPermanent: notesPermanent,
                hintId: "",
                saveButtonId: "step-notes-save-5"
            })}
            ${renderSignoff({
                reviewerInputId: "step-reviewer-5",
                reviewerValue: stepReviewer,
                signoffInputId: "step-signoff-5",
                signoffValue: stepSignOff,
                isComplete: stepComplete,
                saveButtonId: "step-signoff-save-5",
                completeButtonId: "step-signoff-toggle-5"
            })}
        </section>
    `;
}

function headcountHasDifferences() {
    const rosterDiff = Math.abs(headcountState.roster?.difference ?? 0);
    return rosterDiff > 0;
}

function isHeadcountNotesRequired() {
    return !headcountState.skipAnalysis && headcountHasDifferences();
}

function formatMetricValue(value) {
    if (value === null || value === undefined || value === "") return "---";
    const num = Number(value);
    const formatter = window.PrairieForge?.formatNumber;
    if (Number.isFinite(num)) {
        return formatter ? formatter(num) : num.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }
    return String(value);
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

function formatSignedValue(value) {
    if (value === null || value === undefined) return "---";
    if (typeof value !== "number" || Number.isNaN(value)) return "---";
    if (value === 0) return "0";
    return value > 0 ? `+${value}` : value.toString();
}

async function refreshHeadcountAnalysis() {
    if (!hasExcelRuntime()) {
        headcountState.loading = false;
        headcountState.lastError = "Excel runtime is unavailable.";
        renderApp();
        return;
    }
    headcountState.loading = true;
    headcountState.lastError = null;
    // Reset save button since data is being refreshed
    updateSaveButtonState(document.getElementById("headcount-save-btn"), false);
    renderApp();
    try {
        const results = await Excel.run(async (context) => {
            // Use SS_Employee_Roster as the single source of truth for employee data
            const rosterSheet = context.workbook.worksheets.getItem("SS_Employee_Roster");
            const payrollSheet = context.workbook.worksheets.getItem("PTO_Data");
            const analysisSheet = context.workbook.worksheets.getItemOrNullObject("PTO_Analysis");
            const rosterRange = rosterSheet.getUsedRangeOrNullObject();
            const payrollRange = payrollSheet.getUsedRangeOrNullObject();
            rosterRange.load("values");
            payrollRange.load("values");
            analysisSheet.load("isNullObject");
            await context.sync();
            let analysisRange = null;
            if (!analysisSheet.isNullObject) {
                analysisRange = analysisSheet.getUsedRangeOrNullObject();
                analysisRange.load("values");
            }
            await context.sync();
            const rosterValues = rosterRange.isNullObject ? [] : rosterRange.values || [];
            const payrollValues = payrollRange.isNullObject ? [] : payrollRange.values || [];
            const analysisValues = analysisRange && !analysisRange.isNullObject ? analysisRange.values || [] : [];
            // Prefer PTO_Analysis if it has data, otherwise fall back to PTO_Data
            const payrollSource = analysisValues.length ? analysisValues : payrollValues;
            return parseHeadcount(rosterValues, payrollSource);
        });
        headcountState.roster = results.roster;
        headcountState.hasAnalyzed = true;
        headcountState.lastError = null;
    } catch (error) {
        console.warn("PTO headcount: unable to analyze data", error);
        headcountState.lastError = "Unable to analyze headcount data. Try re-running the analysis.";
    } finally {
        headcountState.loading = false;
        renderApp();
    }
}

/**
 * Check if a value looks like a summary row (total, subtotal, etc.)
 * @param {string} value - Value to check
 * @returns {boolean} True if it should be excluded
 */
function isSummaryOrEmpty(value) {
    if (!value) return true;
    const lower = value.toLowerCase().trim();
    if (!lower) return true;
    const summaryPatterns = ["total", "subtotal", "sum", "count", "grand", "average", "avg"];
    return summaryPatterns.some((pattern) => lower.includes(pattern));
}

function parseHeadcount(rosterValues, payrollValues) {
    const roster = {
        rosterCount: 0,
        payrollCount: 0,
        difference: 0,
        mismatches: []
    };

    // Require at least header + 1 data row
    if ((rosterValues?.length || 0) < 2 || (payrollValues?.length || 0) < 2) {
        console.warn("Headcount: insufficient data rows", {
            rosterRows: rosterValues?.length || 0,
            payrollRows: payrollValues?.length || 0
        });
        return { roster };
    }

    const rosterHeaderInfo = findHeaderRow(rosterValues);
    const payrollHeaderInfo = findHeaderRow(payrollValues);

    const rosterHeaders = rosterHeaderInfo.headers;
    const payrollHeaders = payrollHeaderInfo.headers;

    const rosterIdx = {
        employee: getEmployeeColumnIndex(rosterHeaders),
        termination: rosterHeaders.findIndex((h) => h.includes("termination"))
    };
    const payrollIdx = {
        employee: getEmployeeColumnIndex(payrollHeaders)
    };

    // Log column detection for debugging
    console.log("Headcount column detection:", {
        rosterEmployeeCol: rosterIdx.employee,
        rosterTerminationCol: rosterIdx.termination,
        payrollEmployeeCol: payrollIdx.employee,
        rosterHeaders: rosterHeaders.slice(0, 5),
        payrollHeaders: payrollHeaders.slice(0, 5)
    });

    const rosterSet = new Set();
    const payrollSet = new Set();

    // Count unique employees from roster (excluding terminated and summary rows)
    for (let i = rosterHeaderInfo.startIndex; i < rosterValues.length; i += 1) {
        const row = rosterValues[i];
        const employee = rosterIdx.employee >= 0 ? normalizeString(row[rosterIdx.employee]) : "";
        if (isSummaryOrEmpty(employee)) continue;
        // Skip terminated employees
        const termination = rosterIdx.termination >= 0 ? normalizeString(row[rosterIdx.termination]) : "";
        if (termination) continue;
        rosterSet.add(employee.toLowerCase());
    }

    // Count unique employees from payroll (excluding summary rows)
    for (let i = payrollHeaderInfo.startIndex; i < payrollValues.length; i += 1) {
        const row = payrollValues[i];
        const employee = payrollIdx.employee >= 0 ? normalizeString(row[payrollIdx.employee]) : "";
        if (isSummaryOrEmpty(employee)) continue;
        payrollSet.add(employee.toLowerCase());
    }

    roster.rosterCount = rosterSet.size;
    roster.payrollCount = payrollSet.size;
    roster.difference = roster.payrollCount - roster.rosterCount;

    console.log("Headcount results:", {
        rosterCount: roster.rosterCount,
        payrollCount: roster.payrollCount,
        difference: roster.difference
    });

    const missingInPayroll = [...rosterSet].filter((name) => !payrollSet.has(name));
    const missingInRoster = [...payrollSet].filter((name) => !rosterSet.has(name));
    roster.mismatches = [
        ...missingInPayroll.map((name) => `In roster, missing in PTO_Data: ${name}`),
        ...missingInRoster.map((name) => `In PTO_Data, missing in roster: ${name}`)
    ];
    return { roster };
}

function findHeaderRow(values) {
    if (!Array.isArray(values) || !values.length) {
        return { headers: [], startIndex: 1 };
    }
    const headerRowIndex = values.findIndex((row = []) =>
        row.some((cell) => {
            const normalized = normalizeString(cell).toLowerCase();
            return normalized.includes("employee");
        })
    );
    const index = headerRowIndex === -1 ? 0 : headerRowIndex;
    const headers = (values[index] || []).map((h) => normalizeString(h).toLowerCase());
    return { headers, startIndex: index + 1 };
}

function getEmployeeColumnIndex(headers = []) {
    let bestIndex = -1;
    let bestScore = -1;
    headers.forEach((header, index) => {
        const value = header.toLowerCase();
        if (!value.includes("employee")) return;
        let score = 1; // baseline: contains "employee"
        if (value.includes("name")) {
            score = 4; // prefer explicit name column
        } else if (value.includes("id")) {
            score = 2; // lower priority than name
        } else {
            score = 3; // generic employee column without name/id hints
        }
        if (score > bestScore) {
            bestScore = score;
            bestIndex = index;
        }
    });
    return bestIndex;
}

function normalizeString(value) {
    return value == null ? "" : String(value).trim();
}

/**
 * Sync PTO_Analysis from PTO_Data with enrichment from roster, payroll archive, and prior period.
 * This replaces the old PTO_Data_Clean - PTO_Analysis now serves as the single
 * source of enriched employee PTO data.
 * 
 * Columns:
 * - Analysis Date, Employee Name, Department, Pay Rate, Accrual Rate, Carry Over, YTD Used, Balance
 * - Liability Amount (current period)
 * - Accrued PTO $ [Prior Period] (from PTO_Archive_Summary, 0 if not found)
 * - Change (current - prior)
 */
async function syncPtoAnalysis(contextArg = null) {
    const runner = async (context) => {
        const dataSheet = context.workbook.worksheets.getItem("PTO_Data");
        const analysisSheet = context.workbook.worksheets.getItemOrNullObject("PTO_Analysis");
        const rosterSheet = context.workbook.worksheets.getItemOrNullObject("SS_Employee_Roster");
        const prArchiveSheet = context.workbook.worksheets.getItemOrNullObject("PR_Archive_Summary");
        const ptoArchiveSheet = context.workbook.worksheets.getItemOrNullObject("PTO_Archive_Summary");
        const dataRange = dataSheet.getUsedRangeOrNullObject();
        dataRange.load("values");
        analysisSheet.load("isNullObject");
        rosterSheet.load("isNullObject");
        prArchiveSheet.load("isNullObject");
        ptoArchiveSheet.load("isNullObject");
        await context.sync();

        const dataValues = dataRange.isNullObject ? [] : dataRange.values || [];
        if (!dataValues.length) return;
        const dataHeaders = (dataValues[0] || []).map((h) => normalizeName(h));
        const empIdx = dataHeaders.findIndex((h) => h.includes("employee") && h.includes("name"));
        const targetEmpIdx = empIdx >= 0 ? empIdx : 0;
        const accrualRateIdx = findColumnIndex(dataHeaders, ["accrual rate"]);
        const carryOverIdx = findColumnIndex(dataHeaders, ["carry over", "carryover"]);
        // Be specific - "ytd" must be in the column name to avoid matching "Pay Period Accrued"
        const ytdAccruedIdx = dataHeaders.findIndex((h) => h.includes("ytd") && (h.includes("accrued") || h.includes("accrual")));
        // Be specific - "ytd" must be in the column name to avoid matching "Pay Period Used"
        const ytdUsedIdx = dataHeaders.findIndex((h) => h.includes("ytd") && h.includes("used"));
        const balanceIdx = findColumnIndex(dataHeaders, ["balance", "current balance", "pto balance"]);
        
        // Debug logging for column detection
        console.log("[PTO Analysis] PTO_Data headers:", dataHeaders);
        console.log("[PTO Analysis] Column indices found:", {
            employee: targetEmpIdx,
            accrualRate: accrualRateIdx,
            carryOver: carryOverIdx,
            ytdAccrued: ytdAccruedIdx,
            ytdUsed: ytdUsedIdx,
            balance: balanceIdx
        });
        
        // Log the actual header names at each index
        if (ytdUsedIdx >= 0) {
            console.log(`[PTO Analysis] YTD Used column: "${dataHeaders[ytdUsedIdx]}" at index ${ytdUsedIdx}`);
        } else {
            console.warn("[PTO Analysis] YTD Used column NOT FOUND. Headers:", dataHeaders);
        }
        const names = dataValues
            .slice(1)
            .map((row) => normalizeString(row[targetEmpIdx]))
            .filter((name) => name && !name.toLowerCase().includes("total"));
        const dataMap = new Map();
        dataValues.slice(1).forEach((row) => {
            const name = normalizeName(row[targetEmpIdx]);
            if (!name || name.includes("total")) return;
            dataMap.set(name, row);
        });

        // Build roster map for department lookup from SS_Employee_Roster
        let rosterMap = new Map();
        if (!rosterSheet.isNullObject) {
            const rosterRange = rosterSheet.getUsedRangeOrNullObject();
            rosterRange.load("values");
            await context.sync();
            const rosterValues = rosterRange.isNullObject ? [] : rosterRange.values || [];
            if (rosterValues.length) {
                const rosterHeaders = (rosterValues[0] || []).map((h) => normalizeName(h));
                console.log("[PTO Analysis] SS_Employee_Roster headers:", rosterHeaders);
                
                // More flexible column matching - try multiple patterns
                let rosterNameIdx = rosterHeaders.findIndex((h) => h.includes("employee") && h.includes("name"));
                if (rosterNameIdx < 0) {
                    rosterNameIdx = rosterHeaders.findIndex((h) => h === "employee" || h === "name" || h === "full name");
                }
                
                const rosterDeptIdx = rosterHeaders.findIndex((h) => h.includes("department"));
                
                console.log(`[PTO Analysis] Roster column indices - Name: ${rosterNameIdx}, Dept: ${rosterDeptIdx}`);
                
                if (rosterNameIdx >= 0 && rosterDeptIdx >= 0) {
                    rosterValues.slice(1).forEach((row) => {
                        const empName = normalizeName(row[rosterNameIdx]);
                        const dept = normalizeString(row[rosterDeptIdx]);
                        if (empName) {
                            rosterMap.set(empName, dept);
                        }
                    });
                    console.log(`[PTO Analysis] Built roster map with ${rosterMap.size} employees`);
                } else {
                    console.warn("[PTO Analysis] Could not find Name or Department columns in SS_Employee_Roster");
                }
            }
        } else {
            console.warn("[PTO Analysis] SS_Employee_Roster sheet not found");
        }

        // Build archive map for pay rate lookup (from PR_Archive_Summary)
        let payRateMap = new Map();
        if (!prArchiveSheet.isNullObject) {
            const archiveRange = prArchiveSheet.getUsedRangeOrNullObject();
            archiveRange.load("values");
            await context.sync();
            const archiveValues = archiveRange.isNullObject ? [] : archiveRange.values || [];
            if (archiveValues.length) {
                const archiveHeaders = (archiveValues[0] || []).map((h) => normalizeName(h));
                const idx = {
                    payrollDate: findColumnIndex(archiveHeaders, ["payroll date"]),
                    employee: findColumnIndex(archiveHeaders, ["employee"]),
                    category: findColumnIndex(archiveHeaders, ["payroll category", "category"]),
                    amount: findColumnIndex(archiveHeaders, ["amount", "gross salary", "gross_salary", "earnings"])
                };
                if (idx.employee >= 0 && idx.category >= 0 && idx.amount >= 0) {
                    archiveValues.slice(1).forEach((row) => {
                        const name = normalizeName(row[idx.employee]);
                        if (!name) return;
                        const category = normalizeName(row[idx.category]);
                        if (!category.includes("regular") || !category.includes("earn")) return;
                        const amount = Number(row[idx.amount]) || 0;
                        if (!amount) return;
                        const timestamp = coerceTimestamp(row[idx.payrollDate]);
                        const existing = payRateMap.get(name);
                        if (!existing || (timestamp != null && timestamp > existing.timestamp)) {
                            payRateMap.set(name, { payRate: amount / 80, timestamp });
                        }
                    });
                }
            }
        }

        // Build prior period liability map (from PTO_Archive_Summary)
        let priorLiabilityMap = new Map();
        if (!ptoArchiveSheet.isNullObject) {
            const ptoArchiveRange = ptoArchiveSheet.getUsedRangeOrNullObject();
            ptoArchiveRange.load("values");
            await context.sync();
            const ptoArchiveValues = ptoArchiveRange.isNullObject ? [] : ptoArchiveRange.values || [];
            if (ptoArchiveValues.length > 1) {
                const ptoArchiveHeaders = (ptoArchiveValues[0] || []).map((h) => normalizeName(h));
                const priorNameIdx = ptoArchiveHeaders.findIndex((h) => h.includes("employee") && h.includes("name"));
                const priorLiabilityIdx = findColumnIndex(ptoArchiveHeaders, ["liability amount", "liability", "accrued pto"]);
                if (priorNameIdx >= 0 && priorLiabilityIdx >= 0) {
                    ptoArchiveValues.slice(1).forEach((row) => {
                        const name = normalizeName(row[priorNameIdx]);
                        if (!name) return;
                        const liability = Number(row[priorLiabilityIdx]) || 0;
                        priorLiabilityMap.set(name, liability);
                    });
                }
            }
        }

        const analysisDate = getConfigValue(PTO_CONFIG_FIELDS.payrollDate) || "";
        
        // Track data quality issues
        const missingPayRates = [];
        const missingDepartments = [];
        
        const rows = names.map((name, idx) => {
            const normalized = normalizeName(name);
            const department = rosterMap.get(normalized) || "";
            const payRate = payRateMap.get(normalized)?.payRate ?? "";
            const dataRow = dataMap.get(normalized);
            const accrualRate = dataRow && accrualRateIdx >= 0 ? dataRow[accrualRateIdx] ?? "" : "";
            const carryOver = dataRow && carryOverIdx >= 0 ? dataRow[carryOverIdx] ?? "" : "";
            const ytdAccrued = dataRow && ytdAccruedIdx >= 0 ? dataRow[ytdAccruedIdx] ?? "" : "";
            const ytdUsed = dataRow && ytdUsedIdx >= 0 ? dataRow[ytdUsedIdx] ?? "" : "";
            
            // Debug logging for specific employees with issues
            if (normalized.includes("avalos") || normalized.includes("sarah")) {
                console.log(`[PTO Debug] ${name}:`, {
                    ytdUsedIdx,
                    rawValue: dataRow ? dataRow[ytdUsedIdx] : "no dataRow",
                    ytdUsed,
                    fullRow: dataRow
                });
            }
            const balance = dataRow && balanceIdx >= 0 ? Number(dataRow[balanceIdx]) || 0 : 0;
            
            // Track missing data (row index is +2 because of header and 0-based)
            const rowIndex = idx + 2;
            if (!payRate && typeof payRate !== "number") {
                missingPayRates.push({ name, rowIndex });
            }
            if (!department) {
                missingDepartments.push({ name, rowIndex });
            }
            
            // Calculate current liability amount: Balance (hours) × Pay Rate ($/hr)
            const liabilityAmount = (typeof payRate === "number" && balance) ? balance * payRate : 0;
            
            // Get prior period liability (0 if employee not found in archive)
            const priorLiability = priorLiabilityMap.get(normalized) ?? 0;
            
            // Calculate change (current - prior)
            const change = (typeof liabilityAmount === "number" ? liabilityAmount : 0) - priorLiability;
            
            return [analysisDate, name, department, payRate, accrualRate, carryOver, ytdAccrued, ytdUsed, balance, liabilityAmount, priorLiability, change];
        });
        
        // Update state with data quality info
        analysisState.missingPayRates = missingPayRates.filter(e => !analysisState.ignoredMissingPayRates.has(e.name));
        analysisState.missingDepartments = missingDepartments;
        
        console.log(`[PTO Analysis] Data quality: ${missingPayRates.length} missing pay rates, ${missingDepartments.length} missing departments`);
        const output = [
            ["Analysis Date", "Employee Name", "Department", "Pay Rate", "Accrual Rate", "Carry Over", "YTD Accrued", "YTD Used", "Balance", "Liability Amount", "Accrued PTO $ [Prior Period]", "Change"],
            ...rows
        ];

        // Write to PTO_Analysis (create if doesn't exist)
        const targetSheet = analysisSheet.isNullObject ? context.workbook.worksheets.add("PTO_Analysis") : analysisSheet;
        const used = targetSheet.getUsedRangeOrNullObject();
        used.load("address");
        await context.sync();
        if (!used.isNullObject) {
            used.clear();
        }
        
        const columnCount = output[0].length;
        const rowCount = output.length;
        const dataRowCount = rows.length;
        const range = targetSheet.getRangeByIndexes(0, 0, rowCount, columnCount);
        range.values = output;
        
        // Format header row using shared utility (black bg, white bold text)
        const headerRange = targetSheet.getRangeByIndexes(0, 0, 1, columnCount);
        formatSheetHeaders(headerRange);
        
        // Apply column formatting using shared utilities
        // Column indices (0-based):
        // A=0: Analysis Date, B=1: Employee Name, C=2: Department
        // D=3: Pay Rate, E=4: Accrual Rate, F=5: Carry Over, G=6: YTD Accrued, H=7: YTD Used
        // I=8: Balance, J=9: Liability Amount, K=10: Accrued PTO $ [Prior Period], L=11: Change
        
        if (dataRowCount > 0) {
            formatDateColumn(targetSheet, 0, dataRowCount);           // A: Analysis Date
            formatCurrencyColumn(targetSheet, 3, dataRowCount);       // D: Pay Rate
            formatNumberColumn(targetSheet, 4, dataRowCount);         // E: Accrual Rate
            formatNumberColumn(targetSheet, 5, dataRowCount);         // F: Carry Over
            formatNumberColumn(targetSheet, 6, dataRowCount);         // G: YTD Accrued
            formatNumberColumn(targetSheet, 7, dataRowCount);         // H: YTD Used
            formatNumberColumn(targetSheet, 8, dataRowCount);         // I: Balance
            formatCurrencyColumn(targetSheet, 9, dataRowCount);       // J: Liability Amount
            formatCurrencyColumn(targetSheet, 10, dataRowCount);      // K: Accrued PTO $ [Prior Period]
            formatCurrencyColumn(targetSheet, 11, dataRowCount, true); // L: Change (with negative parens)
        }
        
        range.format.autofitColumns();
        
        // Reset selection to A1 so tab switching doesn't land on a random cell
        targetSheet.getRange("A1").select();
        await context.sync();
    };

    if (!hasExcelRuntime()) return;
    if (contextArg) {
        await runner(contextArg);
    } else {
        await Excel.run(runner);
    }
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
    }
    saveStepField(2, "notes", textarea.value);
}

function bindHeadcountNotesGuard() {
    const textarea = document.getElementById("step-notes-input");
    if (!textarea) return;
    textarea.addEventListener("input", () => {
        if (!headcountState.skipAnalysis) return;
        const value = textarea.value || "";
        if (!value.startsWith(HEADCOUNT_SKIP_NOTE)) {
            const remainder = value.replace(HEADCOUNT_SKIP_NOTE, "").trimStart();
            textarea.value = HEADCOUNT_SKIP_NOTE + (remainder ? `\n${remainder}` : "");
        }
        saveStepField(2, "notes", textarea.value);
    });
}

function handleHeadcountSignoff() {
    const notesRequired = isHeadcountNotesRequired();
    const notesValue = document.getElementById("step-notes-input")?.value.trim() || "";
    if (notesRequired && !notesValue) {
        window.alert("Please enter a brief explanation of the outstanding differences before completing this step.");
        return;
    }
}
