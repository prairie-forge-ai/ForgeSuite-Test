import { applyModuleTabVisibility } from "../Common/tab-visibility.js";
import { activateHomepageSheet, getHomepageConfig, renderAdaFab } from "../Common/homepage-sheet.js";

/**
 * ACF ForgeSuite - Module Selector
 * © 2025 Prairie Forge LLC
 */

// Admin password (simple protection against accidental changes)
const ADMIN_PASSWORD = "Boise123";

// Step Templates - define standard steps with their associated tabs and fields
const STEP_TEMPLATES = [
    {
        id: "config",
        name: "Config",
        description: "Module configuration and setup (stored in SS_PF_Config)",
        icon: "⚙️",
        suggestedTabs: [], // No separate tab - config lives in SS_PF_Config
        fields: ["Notes", "Reviewer", "SignOff", "Complete"]
    },
    {
        id: "import",
        name: "Import",
        description: "Import data from external source",
        icon: "📥",
        suggestedTabs: [
            { suffix: "_Data", template: "data-import", description: "Raw imported data" },
            { suffix: "_Data_Clean", template: "data-clean", description: "Validated/cleaned data" }
        ],
        fields: ["Notes", "Reviewer", "SignOff", "Complete"]
    },
    {
        id: "headcount",
        name: "Headcount",
        description: "Employee headcount reconciliation",
        icon: "👥",
        suggestedTabs: [
            { suffix: "_Roster", template: "roster", description: "Employee roster comparison" }
        ],
        fields: ["Notes", "Reviewer", "SignOff", "Complete"]
    },
    {
        id: "validate",
        name: "Validate",
        description: "Data validation and verification",
        icon: "✓",
        suggestedTabs: [],
        fields: ["Notes", "Reviewer", "SignOff", "Complete"]
    },
    {
        id: "review",
        name: "Review",
        description: "Review and analysis of data",
        icon: "🔍",
        suggestedTabs: [
            { suffix: "_Analysis", template: "analysis", description: "Analysis and calculations" }
        ],
        fields: ["Notes", "Reviewer", "SignOff", "Complete"]
    },
    {
        id: "expense",
        name: "Expense",
        description: "Expense mapping and categorization",
        icon: "💰",
        suggestedTabs: [
            { suffix: "_Expense_Review", template: "expense", description: "Expense review and mapping" },
            { suffix: "_Expense_Mapping", template: "mapping", description: "Account mapping configuration" }
        ],
        fields: ["Notes", "Reviewer", "SignOff", "Complete"]
    },
    {
        id: "je",
        name: "JE",
        description: "Journal entry creation",
        icon: "📝",
        suggestedTabs: [
            { suffix: "_JE_Draft", template: "journal", description: "Draft journal entry" }
        ],
        fields: ["Notes", "Reviewer", "SignOff", "Complete"]
    },
    {
        id: "archive",
        name: "Archive",
        description: "Archive data for historical reference",
        icon: "📁",
        suggestedTabs: [
            { suffix: "_Archive_Summary", template: "archive", description: "Archived summary data" }
        ],
        fields: ["Notes", "Reviewer", "SignOff", "Complete"]
    }
];

// Tab Templates - define standard tab structures with headers
const TAB_TEMPLATES = {
    "config": {
        name: "Configuration",
        headers: ["Field", "Value"],
        description: "Key-value storage (note: module config uses SS_PF_Config instead)"
    },
    "data-import": {
        name: "Data Import",
        headers: ["Employee", "Department", "Value", "Rate", "Balance"],
        description: "Raw data from external source (headers vary by source)"
    },
    "data-clean": {
        name: "Cleaned Data",
        headers: ["Employee", "Department", "Value", "Rate", "Balance"],
        description: "Validated and standardized data"
    },
    "roster": {
        name: "Roster",
        headers: ["Employee", "Department", "Status", "Hire_Date"],
        description: "Employee roster for reconciliation"
    },
    "analysis": {
        name: "Analysis",
        headers: ["Date", "Employee", "Department", "Pay_Rate", "Accrual_Rate", "Balance", "Liability"],
        description: "Calculated analysis output"
    },
    "expense": {
        name: "Expense Review",
        headers: ["Account", "Description", "Amount", "Category", "Status"],
        description: "Expense line items for review"
    },
    "mapping": {
        name: "Mapping",
        headers: ["Source_Account", "Source_Description", "Target_Account", "Target_Description"],
        description: "Account mapping configuration"
    },
    "journal": {
        name: "Journal Entry",
        headers: ["Account", "Description", "Debit", "Credit"],
        description: "Journal entry draft"
    },
    "archive": {
        name: "Archive",
        headers: ["Period", "Employee", "Department", "Amount", "Archived_Date"],
        description: "Historical archived data"
    }
};

// Legacy: simple step names for backwards compatibility
const STANDARD_STEPS = STEP_TEMPLATES.map(s => s.name);

// Module registry (editable via admin)
let moduleRegistry = [
    { name: "PTO Accrual", prefix: "PTO", steps: ["Config", "Import", "Headcount", "Validate", "Review", "JE", "Archive"] },
    { name: "Payroll Recorder", prefix: "PR", steps: ["Config", "Import", "Headcount", "Validate", "Expense", "JE", "Archive"] }
];

// Shared config values
let sharedConfig = {
    companyName: "",
    defaultReviewer: "",
    accountingLink: "",
    payrollLink: ""
};

// Tab mappings (loaded from SS_PF_Config "Tab Structure" rows)
let tabMappings = [];

const MODULES = [
    {
        key: "payroll-recorder",
        name: "Payroll Recorder",
        available: true,
        url: "../payroll-recorder/index.html"
    },
    {
        key: "pto-accrual",
        name: "PTO Accrual",
        available: true,
        url: "../pto-accrual/index.html"
    },
    {
        key: "credit-card-expense",
        name: "Credit Card Expense Review",
        available: true,
        url: "#"
    },
    {
        key: "commission-calc",
        name: "Commission Calculation",
        available: true,
        url: "#"
    }
];

// Lucide icons (https://lucide.dev) used in place of emoji badges (ISC license).
const MODULE_ICON_MAP = {
    "payroll-recorder": `
        <svg
            class="module-icon-svg"
            aria-hidden="true"
            focusable="false"
            xmlns="http://www.w3.org/2000/svg"
            width="24"
            height="24"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            stroke-width="2"
            stroke-linecap="round"
            stroke-linejoin="round"
        >
            <path d="M3.85 8.62a4 4 0 0 1 4.78-4.77 4 4 0 0 1 6.74 0 4 4 0 0 1 4.78 4.78 4 4 0 0 1 0 6.74 4 4 0 0 1-4.77 4.78 4 4 0 0 1-6.75 0 4 4 0 0 1-4.78-4.77 4 4 0 0 1 0-6.76Z" />
            <path d="M16 8h-6a2 2 0 1 0 0 4h4a2 2 0 1 1 0 4H8" />
            <path d="M12 18V6" />
        </svg>
    `.trim(),
    "employee-roster": `
        <svg
            class="module-icon-svg"
            aria-hidden="true"
            focusable="false"
            xmlns="http://www.w3.org/2000/svg"
            width="24"
            height="24"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            stroke-width="2"
            stroke-linecap="round"
            stroke-linejoin="round"
        >
            <path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2" />
            <path d="M16 3.128a4 4 0 0 1 0 7.744" />
            <path d="M22 21v-2a4 4 0 0 0-3-3.87" />
            <circle cx="9" cy="7" r="4" />
        </svg>
    `.trim(),
    "pto-accrual": `
        <svg
            class="module-icon-svg"
            aria-hidden="true"
            focusable="false"
            xmlns="http://www.w3.org/2000/svg"
            width="24"
            height="24"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            stroke-width="2"
            stroke-linecap="round"
            stroke-linejoin="round"
        >
            <circle cx="12" cy="12" r="4" />
            <path d="M12 2v2" />
            <path d="M12 20v2" />
            <path d="m4.93 4.93 1.41 1.41" />
            <path d="m17.66 17.66 1.41 1.41" />
            <path d="M2 12h2" />
            <path d="M20 12h2" />
            <path d="m6.34 17.66-1.41 1.41" />
            <path d="m19.07 4.93-1.41 1.41" />
        </svg>
    `.trim(),
    "credit-card-expense": `
        <svg
            class="module-icon-svg"
            aria-hidden="true"
            focusable="false"
            xmlns="http://www.w3.org/2000/svg"
            width="24"
            height="24"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            stroke-width="2"
            stroke-linecap="round"
            stroke-linejoin="round"
        >
            <rect width="20" height="14" x="2" y="5" rx="2" />
            <line x1="2" x2="22" y1="10" y2="10" />
        </svg>
    `.trim(),
    "commission-calc": `
        <svg
            class="module-icon-svg"
            aria-hidden="true"
            focusable="false"
            xmlns="http://www.w3.org/2000/svg"
            width="24"
            height="24"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            stroke-width="2"
            stroke-linecap="round"
            stroke-linejoin="round"
        >
            <line x1="12" x2="12" y1="2" y2="22" />
            <path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6" />
        </svg>
    `.trim()
};

const DEFAULT_MODULE_ICON = `
    <svg
        class="module-icon-svg"
        aria-hidden="true"
        focusable="false"
        xmlns="http://www.w3.org/2000/svg"
        width="24"
        height="24"
        viewBox="0 0 24 24"
        fill="none"
        stroke="currentColor"
        stroke-width="2"
        stroke-linecap="round"
        stroke-linejoin="round"
    >
        <rect x="2" y="4" width="20" height="16" rx="2" />
        <path d="M10 4v4" />
        <path d="M2 8h20" />
        <path d="M6 4v4" />
    </svg>
`.trim();

const MODULE_ALIAS_TOKENS = {
    "payroll-recorder": ["payroll-recorder", "payroll", "payroll recorder", "payroll review", "pr"],
    "employee-roster": ["employee-roster", "employee roster", "headcount", "headcount review", "roster"],
    "pto-accrual": ["pto-accrual", "pto", "pto accrual", "pto review"],
    "credit-card-expense": ["credit-card", "credit card", "cc expense", "card expense"],
    "commission-calc": ["commission", "commission calc", "commissions", "sales commission"]
};

const heroGreetingEl = document.getElementById("heroGreeting");
const moduleCountEl = document.getElementById("moduleCount");
const moduleGridEl = document.getElementById("moduleGrid");

// Reference Data state
const refDataState = {
    type: null, // 'roster' or 'accounts'
    data: [],
    headers: []
};

// Reference Data configuration
const REF_DATA_CONFIG = {
    roster: {
        title: "Employee Roster",
        sheetName: "SS_Employee_Roster",
        defaultHeaders: ["Employee", "Department", "Pay_Rate", "Status", "Hire_Date", "Term_Date"]
    },
    accounts: {
        title: "Chart of Accounts",
        sheetName: "SS_Chart_of_Accounts",
        defaultHeaders: ["Account_Number", "Account_Name", "Type", "Category"]
    }
};

function escapeHtml(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

Office.onReady(() => init()).catch((err) => {
    console.error("Office.onReady error:", err);
    init();
});

async function init() {
    try {
        console.log("ForgeSuite init starting...");
        renderHero();
        console.log("renderHero complete");
        renderModules();
        console.log("renderModules complete");
        wireActions();
        console.log("wireActions complete");
        initQuickAccess();
        console.log("initQuickAccess complete");
        
        // Activate homepage sheet
        const homepageConfig = getHomepageConfig("module-selector");
        await activateHomepageSheet(homepageConfig.sheetName, homepageConfig.title, homepageConfig.subtitle);
        console.log("Homepage sheet activated");
        
        // Show Ada FAB on module selector
        renderAdaFab();
        console.log("Ada FAB rendered");
        
        console.log("ForgeSuite init complete");
    } catch (err) {
        console.error("ForgeSuite init error:", err);
        // Show error to user
        const heroEl = document.getElementById("heroGreeting");
        if (heroEl) {
            heroEl.innerHTML = `<span style="color: #ef4444;">Error loading: ${err.message}</span>`;
        }
    }
}

function renderHero() {
    heroGreetingEl.innerHTML = buildGreeting();
    moduleCountEl.textContent = MODULES.filter((mod) => mod.available).length;
}

function renderModules() {
    moduleGridEl.innerHTML = MODULES.map((module) => {
        const disabledClass = module.available ? "" : "disabled";
        return `
            <div class="module-card pf-clickable ${disabledClass}" data-module="${escapeHtml(module.key)}">
                <div class="module-name">
                    ${renderIcon(module.key)}
                    <span>${escapeHtml(module.name)}</span>
                </div>
            </div>
        `;
    }).join("");

    moduleGridEl.querySelectorAll(".module-card").forEach((card) => {
        if (card.classList.contains("disabled")) return;
        card.addEventListener("click", () => {
            const key = card.getAttribute("data-module");
            launchModule(key);
        });
    });
}

function renderIcon(key) {
    const icon = MODULE_ICON_MAP[key] || DEFAULT_MODULE_ICON;
    return `<span class="module-icon pf-icon" aria-hidden="true">${icon}</span>`;
}

function wireActions() {
    // Support card temporarily removed
}

function buildGreeting() {
    const hour = new Date().getHours();
    let timeGreeting;
    if (hour < 12) timeGreeting = "Good morning";
    else if (hour < 18) timeGreeting = "Good afternoon";
    else timeGreeting = "Good evening";
    
    // Rotating taglines - changes daily based on date
    const taglines = [
        "Let's get your work moving.",
        "Let's make today easier.",
        "Ready when you are.",
        "Time to simplify your workflow."
    ];
    
    // Use day of year to pick tagline (consistent for the whole day)
    const now = new Date();
    const start = new Date(now.getFullYear(), 0, 0);
    const diff = now - start;
    const oneDay = 1000 * 60 * 60 * 24;
    const dayOfYear = Math.floor(diff / oneDay);
    const tagline = taglines[dayOfYear % taglines.length];
    
    return `<strong>${timeGreeting}</strong> <span class="greeting-divider">|</span> <span class="greeting-tagline">${tagline}</span>`;
}

async function launchModule(moduleKey) {
    const module = MODULES.find((mod) => mod.key === moduleKey && mod.available);
    if (!module) {
        console.warn("Module not available yet.");
        return;
    }
    await applyModuleTabVisibility(moduleKey);
    window.location.href = module.url;
}

function getAliasTokens(moduleKey) {
    return MODULE_ALIAS_TOKENS[moduleKey] ?? [moduleKey];
}

// =============================================================================
// ADMIN PORTAL
// =============================================================================

const adminOverlay = document.getElementById("adminOverlay");
const adminGate = document.getElementById("adminGate");
const adminDashboard = document.getElementById("adminDashboard");
const adminPassword = document.getElementById("adminPassword");
const adminError = document.getElementById("adminError");
const adminLog = document.getElementById("adminLog");

let isAdminAuthenticated = false;

// Module Builder State
const builderState = {
    currentStep: 1,
    mode: 'new', // 'new' or 'edit'
    module: {
        name: '',
        prefix: '',
        description: '',
        templateModule: null
    },
    selectedSteps: [],
    tabAssignments: {}, // stepId -> [tabNames]
    functionAssignments: {}, // stepId -> [functionIds]
    fields: {} // tabName -> [fieldNames]
};

// Builder wizard step templates (simplified for UI)
const BUILDER_STEP_TEMPLATES = [
    { id: 'config', name: 'Configuration', icon: '⚙️', defaultTabs: ['_Config'], description: 'Module settings and setup' },
    { id: 'import', name: 'Data Import', icon: '📥', defaultTabs: ['_Raw_Data'], description: 'Import data from source systems' },
    { id: 'cleanup', name: 'Data Cleanup', icon: '🧹', defaultTabs: ['_Data_Clean'], description: 'Validate and clean imported data' },
    { id: 'review', name: 'Department Review', icon: '👥', defaultTabs: ['_Dept_Review'], description: 'Review by department/team' },
    { id: 'analysis', name: 'Analysis', icon: '📊', defaultTabs: ['_Analysis'], description: 'Calculations and analysis' },
    { id: 'journal', name: 'Journal Entry', icon: '📝', defaultTabs: ['_JE_Draft'], description: 'Create journal entries' },
    { id: 'signoff', name: 'Sign-off', icon: '✅', defaultTabs: [], description: 'Final review and approval' }
];

// Builder wizard function templates
const BUILDER_FUNCTION_TEMPLATES = [
    { id: 'paste', name: 'Paste from Clipboard', icon: '📋', steps: ['import', 'cleanup'] },
    { id: 'file_import', name: 'Import from File', icon: '📁', steps: ['import'] },
    { id: 'external_pull', name: 'Pull from External System', icon: '🔗', steps: ['import'] },
    { id: 'validate', name: 'Validate Data', icon: '✓', steps: ['cleanup', 'review'] },
    { id: 'compare_roster', name: 'Compare to Roster', icon: '👥', steps: ['review'] },
    { id: 'flag_discrepancies', name: 'Flag Discrepancies', icon: '🚩', steps: ['review', 'analysis'] },
    { id: 'approve_reject', name: 'Approve / Reject', icon: '✅', steps: ['review', 'signoff'] },
    { id: 'calculate', name: 'Run Calculations', icon: '🔢', steps: ['analysis'] },
    { id: 'generate_je', name: 'Generate Journal Entry', icon: '📝', steps: ['journal'] },
    { id: 'export', name: 'Export Data', icon: '📤', steps: ['journal', 'signoff'] },
    { id: 'copilot', name: 'Copilot Assistance', icon: '🤖', steps: ['import', 'cleanup', 'review', 'analysis', 'journal'] }
];

function initAdmin() {
    // Admin link click
    document.getElementById("adminLink")?.addEventListener("click", openAdminOverlay);
    
    // Close buttons
    document.getElementById("adminClose")?.addEventListener("click", closeAdminOverlay);
    document.getElementById("adminCloseAuth")?.addEventListener("click", closeAdminOverlay);
    
    // Password submit
    document.getElementById("adminSubmit")?.addEventListener("click", handleAdminLogin);
    adminPassword?.addEventListener("keypress", (e) => {
        if (e.key === "Enter") handleAdminLogin();
    });
    
    // Mode switching (Builder vs Settings)
    document.querySelectorAll(".admin-mode-btn").forEach(btn => {
        btn.addEventListener("click", () => switchAdminMode(btn.dataset.mode));
    });
    
    // Settings tabs (in Settings mode)
    document.querySelectorAll(".settings-tab").forEach(tab => {
        tab.addEventListener("click", () => switchSettingsTab(tab.dataset.settings));
    });
    
    // Tab Management (now in Settings mode)
    document.getElementById("tabModuleFilter")?.addEventListener("change", renderTabMappings);
    document.getElementById("addTabMappingBtn")?.addEventListener("click", addTabMapping);
    
    // Shared config save
    document.getElementById("saveSharedConfig")?.addEventListener("click", saveSharedConfig);
    
    // Utilities
    document.getElementById("validateConfigBtn")?.addEventListener("click", validateConfig);
    document.getElementById("generateFieldsBtn")?.addEventListener("click", writeAllFieldsToSheet);
    document.getElementById("exportConfigBtn")?.addEventListener("click", exportConfig);
    
    // Close on overlay click
    adminOverlay?.addEventListener("click", (e) => {
        if (e.target === adminOverlay) closeAdminOverlay();
    });
    
    // Module Builder wizard events
    initModuleBuilder();
}

function initModuleBuilder() {
    // Step 1: New/Edit choice
    document.getElementById("builderNewModule")?.addEventListener("click", () => {
        builderState.mode = 'new';
        document.getElementById("builderNewModule").classList.add("selected");
        document.getElementById("builderEditModule").classList.remove("selected");
        document.getElementById("builderNewModuleForm").hidden = false;
        document.getElementById("builderEditModuleForm").hidden = true;
    });
    
    document.getElementById("builderEditModule")?.addEventListener("click", () => {
        builderState.mode = 'edit';
        document.getElementById("builderEditModule").classList.add("selected");
        document.getElementById("builderNewModule").classList.remove("selected");
        document.getElementById("builderNewModuleForm").hidden = true;
        document.getElementById("builderEditModuleForm").hidden = false;
        populateEditModuleDropdown();
    });
    
    // Template checkbox toggle
    document.getElementById("builderUseTemplate")?.addEventListener("change", (e) => {
        document.getElementById("builderTemplateModule").hidden = !e.target.checked;
        if (e.target.checked) {
            populateTemplateModuleDropdown();
        }
    });
    
    // Edit module selection
    document.getElementById("builderEditModuleSelect")?.addEventListener("change", (e) => {
        const moduleName = e.target.value;
        if (moduleName) {
            loadModuleForEditing(moduleName);
        } else {
            document.getElementById("builderModuleSummary").hidden = true;
        }
    });
    
    // Navigation buttons
    document.getElementById("builderPrevBtn")?.addEventListener("click", builderPrevStep);
    document.getElementById("builderNextBtn")?.addEventListener("click", builderNextStep);
    document.getElementById("builderCreateBtn")?.addEventListener("click", createModule);
    
    // Custom step form
    document.getElementById("addCustomStep")?.addEventListener("click", () => {
        document.getElementById("customStepForm").hidden = false;
    });
    document.getElementById("cancelCustomStep")?.addEventListener("click", () => {
        document.getElementById("customStepForm").hidden = true;
        document.getElementById("customStepName").value = '';
    });
    document.getElementById("confirmCustomStep")?.addEventListener("click", addCustomStep);
    
    // Fields tab selection
    document.getElementById("builderFieldsTabSelect")?.addEventListener("change", (e) => {
        const tabName = e.target.value;
        document.getElementById("builderFieldsConfig").hidden = !tabName;
        if (tabName) {
            renderFieldsForTab(tabName);
        }
    });
    
    // Add field button - use inline input (prompt() not supported in Office Add-ins)
    document.getElementById("addFieldBtn")?.addEventListener("click", () => {
        const btn = document.getElementById("addFieldBtn");
        const existingInput = btn.parentElement.querySelector('.inline-field-input');
        if (existingInput) {
            existingInput.focus();
            return;
        }
        
        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'admin-input inline-field-input';
        input.placeholder = 'Column name';
        input.style.cssText = 'width: 120px; padding: 6px 10px; font-size: 12px; margin-right: 8px;';
        
        btn.parentElement.insertBefore(input, btn);
        input.focus();
        
        const addField = () => {
            const value = input.value.trim();
            if (value) {
                addFieldToCurrentTab(value);
                input.value = '';
                input.focus();
            }
        };
        
        input.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') addField();
        });
        input.addEventListener('blur', () => {
            // Keep input visible for multiple entries
        });
    });
    
    // Fields dropzone
    initBuilderFieldsDropzone();
    
    // Progress step clicks
    document.querySelectorAll(".builder-progress-step").forEach(step => {
        step.addEventListener("click", () => {
            const targetStep = parseInt(step.dataset.step);
            if (targetStep < builderState.currentStep) {
                goToBuilderStep(targetStep);
            }
        });
    });
}

function switchAdminMode(mode) {
    document.querySelectorAll(".admin-mode-btn").forEach(btn => {
        btn.classList.toggle("active", btn.dataset.mode === mode);
    });
    document.querySelectorAll(".admin-mode-content").forEach(content => {
        content.classList.toggle("active", content.id === `mode-${mode}`);
    });
}

function switchSettingsTab(tabId) {
    document.querySelectorAll(".settings-tab").forEach(tab => {
        tab.classList.toggle("active", tab.dataset.settings === tabId);
    });
    document.querySelectorAll(".settings-content").forEach(content => {
        content.classList.toggle("active", content.id === `settings-${tabId}`);
    });
}

function goToBuilderStep(step) {
    builderState.currentStep = step;
    
    // Update progress indicator
    document.querySelectorAll(".builder-progress-step").forEach(s => {
        const stepNum = parseInt(s.dataset.step);
        s.classList.remove("active", "completed");
        if (stepNum === step) {
            s.classList.add("active");
        } else if (stepNum < step) {
            s.classList.add("completed");
        }
    });
    
    // Show/hide steps
    for (let i = 1; i <= 6; i++) {
        const stepEl = document.getElementById(`builderStep${i}`);
        if (stepEl) {
            stepEl.hidden = i !== step;
            if (i === step) {
                stepEl.classList.add("active");
            } else {
                stepEl.classList.remove("active");
            }
        }
    }
    
    // Update navigation buttons
    const prevBtn = document.getElementById("builderPrevBtn");
    const nextBtn = document.getElementById("builderNextBtn");
    const createBtn = document.getElementById("builderCreateBtn");
    
    prevBtn.hidden = step === 1;
    nextBtn.hidden = step === 6;
    createBtn.hidden = step !== 6;
    
    // Render step-specific content
    if (step === 2) renderBuilderSteps();
    if (step === 3) renderBuilderTabAssignments();
    if (step === 4) renderBuilderFunctions();
    if (step === 5) renderBuilderFieldsTabSelect();
    if (step === 6) renderBuilderReview();
}

function builderPrevStep() {
    if (builderState.currentStep > 1) {
        goToBuilderStep(builderState.currentStep - 1);
    }
}

function builderNextStep() {
    // Validate current step
    if (!validateBuilderStep(builderState.currentStep)) {
        return;
    }
    
    // Save current step data
    saveBuilderStepData(builderState.currentStep);
    
    if (builderState.currentStep < 6) {
        goToBuilderStep(builderState.currentStep + 1);
    }
}

function validateBuilderStep(step) {
    switch (step) {
        case 1:
            if (builderState.mode === 'new') {
                const name = document.getElementById("builderModuleName")?.value.trim();
                const prefix = document.getElementById("builderModulePrefix")?.value.trim();
                if (!name || !prefix) {
                    alert("Please enter a module name and prefix.");
                    return false;
                }
            } else {
                const selected = document.getElementById("builderEditModuleSelect")?.value;
                if (!selected) {
                    alert("Please select a module to edit.");
                    return false;
                }
            }
            return true;
        case 2:
            if (builderState.selectedSteps.length === 0) {
                alert("Please select at least one step.");
                return false;
            }
            return true;
        default:
            return true;
    }
}

function saveBuilderStepData(step) {
    switch (step) {
        case 1:
            if (builderState.mode === 'new') {
                builderState.module.name = document.getElementById("builderModuleName")?.value.trim();
                builderState.module.prefix = document.getElementById("builderModulePrefix")?.value.trim().toUpperCase();
                builderState.module.description = document.getElementById("builderModuleDesc")?.value.trim();
                
                const useTemplate = document.getElementById("builderUseTemplate")?.checked;
                if (useTemplate) {
                    builderState.module.templateModule = document.getElementById("builderTemplateModule")?.value;
                }
            }
            break;
        case 2:
            // Steps are saved in real-time via checkbox changes
            break;
        case 3:
            // Tab assignments saved in real-time
            break;
        case 4:
            // Functions saved in real-time
            break;
        case 5:
            // Fields saved in real-time
            break;
    }
}

function renderBuilderSteps() {
    const container = document.getElementById("builderStepsList");
    if (!container) return;
    
    // Update module name in header
    const moduleName = builderState.mode === 'new' 
        ? builderState.module.name 
        : document.getElementById("builderEditModuleSelect")?.value;
    document.getElementById("step2ModuleName").textContent = moduleName || "this module";
    
    container.innerHTML = BUILDER_STEP_TEMPLATES.map((step, idx) => {
        const isSelected = builderState.selectedSteps.includes(step.id);
        return `
            <label class="builder-step-item ${isSelected ? 'selected' : ''}" data-step-id="${step.id}">
                <input type="checkbox" ${isSelected ? 'checked' : ''}>
                <span class="step-number">${idx + 1}</span>
                <span class="step-name">${step.icon} ${step.name}</span>
                <span class="step-tabs">${step.description}</span>
            </label>
        `;
    }).join('');
    
    // Add custom steps
    builderState.selectedSteps.forEach(stepId => {
        if (!BUILDER_STEP_TEMPLATES.find(t => t.id === stepId)) {
            container.innerHTML += `
                <label class="builder-step-item selected" data-step-id="${stepId}">
                    <input type="checkbox" checked>
                    <span class="step-number">+</span>
                    <span class="step-name">📌 ${stepId}</span>
                    <span class="step-tabs">Custom step</span>
                </label>
            `;
        }
    });
    
    // Wire up checkboxes
    container.querySelectorAll("input[type='checkbox']").forEach(cb => {
        cb.addEventListener("change", (e) => {
            const stepId = e.target.closest(".builder-step-item").dataset.stepId;
            if (e.target.checked) {
                if (!builderState.selectedSteps.includes(stepId)) {
                    builderState.selectedSteps.push(stepId);
                }
            } else {
                builderState.selectedSteps = builderState.selectedSteps.filter(s => s !== stepId);
            }
            e.target.closest(".builder-step-item").classList.toggle("selected", e.target.checked);
        });
    });
}

function addCustomStep() {
    const name = document.getElementById("customStepName")?.value.trim();
    if (!name) return;
    
    const stepId = name.toLowerCase().replace(/\s+/g, '_');
    if (!builderState.selectedSteps.includes(stepId)) {
        builderState.selectedSteps.push(stepId);
    }
    
    document.getElementById("customStepForm").hidden = true;
    document.getElementById("customStepName").value = '';
    renderBuilderSteps();
}

function renderBuilderTabAssignments() {
    const container = document.getElementById("builderTabsAssignment");
    if (!container) return;
    
    const prefix = builderState.module.prefix || 'XX';
    
    container.innerHTML = builderState.selectedSteps.map((stepId, idx) => {
        const template = BUILDER_STEP_TEMPLATES.find(t => t.id === stepId);
        const stepName = template ? template.name : stepId;
        const defaultTabs = template ? template.defaultTabs : [];
        
        // Initialize tab assignments if not set
        if (!builderState.tabAssignments[stepId]) {
            builderState.tabAssignments[stepId] = defaultTabs.map(t => `${prefix}${t}`);
        }
        
        const tabs = builderState.tabAssignments[stepId];
        
        return `
            <div class="tab-assignment-group" data-step-id="${stepId}">
                <div class="tab-assignment-header">
                    <span class="step-badge">${idx + 1}</span>
                    <h4>${stepName}</h4>
                </div>
                <div class="tab-assignment-tabs">
                    ${tabs.map(tabName => `
                        <span class="tab-pill">
                            <span class="tab-status pending"></span>
                            ${tabName}
                        </span>
                    `).join('')}
                    <button type="button" class="tab-add-btn" data-step="${stepId}">+ Add Tab</button>
                </div>
            </div>
        `;
    }).join('');
    
    // Wire up add buttons - show inline input instead of prompt()
    container.querySelectorAll(".tab-add-btn").forEach(btn => {
        btn.addEventListener("click", () => {
            try {
                console.log("Add Tab button clicked for step:", btn.dataset.step);
                const stepId = btn.dataset.step;
                
                // Create inline input form (prompt() not supported in Office Add-ins)
                const existingForm = btn.parentElement.querySelector('.inline-tab-form');
                if (existingForm) {
                    existingForm.remove();
                    return;
                }
                
                const form = document.createElement('div');
                form.className = 'inline-tab-form';
                form.innerHTML = `
                    <input type="text" class="admin-input inline-tab-input" placeholder="Tab name (e.g., Review)" style="width: 120px; padding: 6px 10px; font-size: 12px;">
                    <button type="button" class="admin-btn admin-btn--small inline-tab-confirm" style="padding: 6px 12px; font-size: 11px;">Add</button>
                `;
                form.style.cssText = 'display: inline-flex; gap: 6px; margin-left: 8px;';
                
                btn.parentElement.appendChild(form);
                
                const input = form.querySelector('.inline-tab-input');
                const confirmBtn = form.querySelector('.inline-tab-confirm');
                
                input.focus();
                
                const addTab = () => {
                    const tabName = input.value.trim();
                    if (tabName) {
                        const fullName = `${prefix}_${tabName}`;
                        console.log("Adding tab:", fullName);
                        if (!builderState.tabAssignments[stepId]) {
                            builderState.tabAssignments[stepId] = [];
                        }
                        builderState.tabAssignments[stepId].push(fullName);
                        renderBuilderTabAssignments();
                    } else {
                        form.remove();
                    }
                };
                
                confirmBtn.addEventListener('click', addTab);
                input.addEventListener('keypress', (e) => {
                    if (e.key === 'Enter') addTab();
                    if (e.key === 'Escape') form.remove();
                });
                
            } catch (err) {
                console.error("Error in Add Tab handler:", err);
            }
        });
    });
}

function renderBuilderFunctions() {
    const container = document.getElementById("builderFunctionsList");
    if (!container) return;
    
    container.innerHTML = builderState.selectedSteps.map((stepId, idx) => {
        const template = BUILDER_STEP_TEMPLATES.find(t => t.id === stepId);
        const stepName = template ? template.name : stepId;
        
        // Get relevant functions for this step
        const relevantFunctions = BUILDER_FUNCTION_TEMPLATES.filter(f => 
            f.steps.includes(stepId) || f.steps.includes('all')
        );
        
        // Initialize function assignments if not set
        if (!builderState.functionAssignments[stepId]) {
            builderState.functionAssignments[stepId] = relevantFunctions.map(f => f.id);
        }
        
        return `
            <div class="function-group" data-step-id="${stepId}">
                <div class="function-group-header">
                    <span class="step-badge">${idx + 1}</span>
                    <h4>${stepName}</h4>
                </div>
                <div class="function-checkboxes">
                    ${relevantFunctions.map(func => {
                        const isChecked = builderState.functionAssignments[stepId].includes(func.id);
                        return `
                            <label class="function-checkbox">
                                <input type="checkbox" data-step="${stepId}" data-func="${func.id}" ${isChecked ? 'checked' : ''}>
                                <span>${func.icon} ${func.name}</span>
                            </label>
                        `;
                    }).join('')}
                </div>
            </div>
        `;
    }).join('');
    
    // Wire up checkboxes
    container.querySelectorAll("input[type='checkbox']").forEach(cb => {
        cb.addEventListener("change", (e) => {
            const stepId = e.target.dataset.step;
            const funcId = e.target.dataset.func;
            
            if (!builderState.functionAssignments[stepId]) {
                builderState.functionAssignments[stepId] = [];
            }
            
            if (e.target.checked) {
                if (!builderState.functionAssignments[stepId].includes(funcId)) {
                    builderState.functionAssignments[stepId].push(funcId);
                }
            } else {
                builderState.functionAssignments[stepId] = 
                    builderState.functionAssignments[stepId].filter(f => f !== funcId);
            }
        });
    });
}

function renderBuilderFieldsTabSelect() {
    const select = document.getElementById("builderFieldsTabSelect");
    if (!select) return;
    
    // Gather all tabs from assignments
    const allTabs = [];
    Object.values(builderState.tabAssignments).forEach(tabs => {
        tabs.forEach(tab => {
            if (!allTabs.includes(tab)) {
                allTabs.push(tab);
            }
        });
    });
    
    select.innerHTML = '<option value="">Select a tab to configure...</option>' +
        allTabs.map(tab => `<option value="${tab}">${tab}</option>`).join('');
}

function renderFieldsForTab(tabName) {
    const container = document.getElementById("builderFieldsList");
    if (!container) return;
    
    if (!builderState.fields[tabName]) {
        builderState.fields[tabName] = [];
    }
    
    const fields = builderState.fields[tabName];
    
    container.innerHTML = fields.map(field => `
        <span class="field-tag" data-field="${field}">
            ${field}
            <button type="button" class="field-remove" data-field="${field}">×</button>
        </span>
    `).join('') || '<p style="color: var(--pf-text-secondary); font-size: 13px;">No columns defined yet. Upload a file or add manually.</p>';
    
    // Wire up remove buttons
    container.querySelectorAll(".field-remove").forEach(btn => {
        btn.addEventListener("click", () => {
            const field = btn.dataset.field;
            const currentTab = document.getElementById("builderFieldsTabSelect")?.value;
            if (currentTab && builderState.fields[currentTab]) {
                builderState.fields[currentTab] = builderState.fields[currentTab].filter(f => f !== field);
                renderFieldsForTab(currentTab);
            }
        });
    });
}

function addFieldToCurrentTab(fieldName) {
    const currentTab = document.getElementById("builderFieldsTabSelect")?.value;
    if (!currentTab) return;
    
    if (!builderState.fields[currentTab]) {
        builderState.fields[currentTab] = [];
    }
    
    if (!builderState.fields[currentTab].includes(fieldName)) {
        builderState.fields[currentTab].push(fieldName);
        renderFieldsForTab(currentTab);
    }
}

function initBuilderFieldsDropzone() {
    const dropzone = document.getElementById("builderFieldsDropzone");
    const fileInput = document.getElementById("builderFieldsFile");
    
    if (!dropzone || !fileInput) return;
    
    dropzone.addEventListener("click", () => fileInput.click());
    
    dropzone.addEventListener("dragover", (e) => {
        e.preventDefault();
        dropzone.classList.add("dragover");
    });
    
    dropzone.addEventListener("dragleave", () => {
        dropzone.classList.remove("dragover");
    });
    
    dropzone.addEventListener("drop", (e) => {
        e.preventDefault();
        dropzone.classList.remove("dragover");
        const file = e.dataTransfer?.files[0];
        if (file) handleBuilderFieldsFile(file);
    });
    
    fileInput.addEventListener("change", (e) => {
        const file = e.target.files?.[0];
        if (file) handleBuilderFieldsFile(file);
    });
}

function handleBuilderFieldsFile(file) {
    const currentTab = document.getElementById("builderFieldsTabSelect")?.value;
    if (!currentTab) {
        alert("Please select a tab first.");
        return;
    }
    
    const reader = new FileReader();
    reader.onload = (e) => {
        const text = e.target.result;
        const lines = text.split('\n');
        if (lines.length > 0) {
            const headers = lines[0].split(',').map(h => h.trim().replace(/"/g, ''));
            builderState.fields[currentTab] = headers.filter(h => h);
            renderFieldsForTab(currentTab);
        }
    };
    reader.readAsText(file);
}

function renderBuilderReview() {
    const container = document.getElementById("builderReview");
    if (!container) return;
    
    const totalSteps = builderState.selectedSteps.length;
    const totalTabs = Object.values(builderState.tabAssignments).flat().length;
    const totalFunctions = Object.values(builderState.functionAssignments).flat().length;
    const totalFields = Object.values(builderState.fields).flat().length;
    
    container.innerHTML = `
        <div class="review-stats">
            <div class="review-stat">
                <span class="review-stat-value">${totalSteps}</span>
                <span class="review-stat-label">Steps</span>
            </div>
            <div class="review-stat">
                <span class="review-stat-value">${totalTabs}</span>
                <span class="review-stat-label">Tabs</span>
            </div>
            <div class="review-stat">
                <span class="review-stat-value">${totalFunctions}</span>
                <span class="review-stat-label">Functions</span>
            </div>
            <div class="review-stat">
                <span class="review-stat-value">${totalFields}</span>
                <span class="review-stat-label">Fields</span>
            </div>
        </div>
        
        <div class="review-section">
            <h4>Module</h4>
            <div class="review-section-content">
                <div class="review-item">
                    <span class="review-check">✓</span>
                    <strong>${builderState.module.name}</strong> (${builderState.module.prefix})
                </div>
            </div>
        </div>
        
        <div class="review-section">
            <h4>Steps</h4>
            <div class="review-section-content">
                ${builderState.selectedSteps.map(stepId => {
                    const template = BUILDER_STEP_TEMPLATES.find(t => t.id === stepId);
                    const name = template ? template.name : stepId;
                    return `<div class="review-item"><span class="review-check">✓</span> ${name}</div>`;
                }).join('')}
            </div>
        </div>
        
        <div class="review-section">
            <h4>Tabs to Create</h4>
            <div class="review-section-content">
                ${Object.values(builderState.tabAssignments).flat().map(tab => 
                    `<div class="review-item"><span class="review-check">✓</span> ${tab}</div>`
                ).join('')}
            </div>
        </div>
    `;
}

function populateEditModuleDropdown() {
    const select = document.getElementById("builderEditModuleSelect");
    if (!select) return;
    
    select.innerHTML = '<option value="">Choose a module...</option>' +
        MODULES.map(m => `<option value="${m.name}">${m.name}</option>`).join('');
}

function populateTemplateModuleDropdown() {
    const select = document.getElementById("builderTemplateModule");
    if (!select) return;
    
    select.innerHTML = '<option value="">Select module...</option>' +
        MODULES.map(m => `<option value="${m.name}">${m.name}</option>`).join('');
}

function loadModuleForEditing(moduleName) {
    const module = MODULES.find(m => m.name === moduleName);
    if (!module) return;
    
    builderState.module.name = module.name;
    builderState.module.prefix = module.prefix || moduleName.substring(0, 3).toUpperCase();
    
    // Show summary
    const summary = document.getElementById("builderModuleSummary");
    if (summary) {
        summary.hidden = false;
        // These would be populated from actual config
        document.getElementById("summarySteps").textContent = "6";
        document.getElementById("summaryTabs").textContent = tabMappings.filter(t => t.module === moduleName).length.toString();
        document.getElementById("summaryFunctions").textContent = "12";
    }
}

async function createModule() {
    const progressContainer = document.getElementById("builderCreateProgress");
    const progressFill = document.getElementById("builderProgressFill");
    const progressText = document.getElementById("builderProgressText");
    
    if (!progressContainer) return;
    
    progressContainer.hidden = false;
    
    try {
        // Step 1: Register module
        progressText.textContent = "Registering module...";
        progressFill.style.width = "20%";
        await new Promise(r => setTimeout(r, 500));
        
        // Step 2: Create tabs
        progressText.textContent = "Creating Excel tabs...";
        progressFill.style.width = "40%";
        
        if (hasExcelRuntime()) {
            await Excel.run(async (context) => {
                const allTabs = Object.values(builderState.tabAssignments).flat();
                for (const tabName of allTabs) {
                    const sheet = context.workbook.worksheets.getItemOrNullObject(tabName);
                    sheet.load("isNullObject");
                    await context.sync();
                    
                    if (sheet.isNullObject) {
                        context.workbook.worksheets.add(tabName);
                    }
                }
                await context.sync();
            });
        }
        
        progressFill.style.width = "60%";
        await new Promise(r => setTimeout(r, 300));
        
        // Step 3: Add headers
        progressText.textContent = "Setting up column headers...";
        progressFill.style.width = "80%";
        
        if (hasExcelRuntime()) {
            await Excel.run(async (context) => {
                for (const [tabName, fields] of Object.entries(builderState.fields)) {
                    if (fields.length > 0) {
                        const sheet = context.workbook.worksheets.getItem(tabName);
                        const headerRange = sheet.getRange(`A1:${String.fromCharCode(64 + fields.length)}1`);
                        headerRange.values = [fields];
                        headerRange.format.font.bold = true;
                    }
                }
                await context.sync();
            });
        }
        
        // Step 4: Save config
        progressText.textContent = "Saving configuration...";
        progressFill.style.width = "100%";
        await new Promise(r => setTimeout(r, 500));
        
        progressText.textContent = "✓ Module created successfully!";
        progressFill.style.background = "linear-gradient(90deg, #10b981, #059669)";
        
        // Reload tab mappings
        await loadTabMappings();
        renderTabMappings();
        
    } catch (error) {
        console.error("Error creating module:", error);
        progressText.textContent = "Error: " + error.message;
        progressFill.style.background = "#ef4444";
    }
}

function resetBuilderState() {
    builderState.currentStep = 1;
    builderState.mode = 'new';
    builderState.module = { name: '', prefix: '', description: '', templateModule: null };
    builderState.selectedSteps = [];
    builderState.tabAssignments = {};
    builderState.functionAssignments = {};
    builderState.fields = {};
    
    // Reset UI
    goToBuilderStep(1);
    document.getElementById("builderNewModuleForm").hidden = true;
    document.getElementById("builderEditModuleForm").hidden = true;
    document.getElementById("builderNewModule")?.classList.remove("selected");
    document.getElementById("builderEditModule")?.classList.remove("selected");
    document.getElementById("builderCreateProgress").hidden = true;
}

function openAdminOverlay() {
    adminOverlay.hidden = false;
    if (isAdminAuthenticated) {
        showAdminDashboard();
    } else {
        showAdminGate();
    }
}

function closeAdminOverlay() {
    adminOverlay.hidden = true;
    adminPassword.value = "";
    adminError.hidden = true;
}

function showAdminGate() {
    adminGate.hidden = false;
    adminDashboard.hidden = true;
    setTimeout(() => adminPassword?.focus(), 100);
}

async function showAdminDashboard() {
    adminGate.hidden = true;
    adminDashboard.hidden = false;
    
    // Reset builder state when opening
    resetBuilderState();
    
    // Load data from Excel
    loadSharedConfigFromExcel();
    await loadTabMappings();
    renderTabMappings();
}

function handleAdminLogin() {
    const password = adminPassword.value;
    if (password === ADMIN_PASSWORD) {
        isAdminAuthenticated = true;
        adminError.hidden = true;
        showAdminDashboard();
    } else {
        adminError.hidden = false;
        adminPassword.value = "";
        adminPassword.focus();
    }
}

// switchAdminTab removed - replaced by switchAdminMode and switchSettingsTab

// =============================================================================
// STEP & TAB TEMPLATES
// =============================================================================

/**
 * Initialize collapsible section toggles
 */
function initCollapsibleSections() {
    document.querySelectorAll(".admin-section-toggle").forEach(toggle => {
        toggle.addEventListener("click", () => {
            const contentId = toggle.id.replace("Toggle", "Content");
            const content = document.getElementById(contentId);
            const isExpanded = toggle.classList.contains("expanded");
            
            if (isExpanded) {
                toggle.classList.remove("expanded");
                if (content) content.hidden = true;
            } else {
                toggle.classList.add("expanded");
                if (content) content.hidden = false;
            }
        });
    });
}

/**
 * Render step templates list
 */
function renderStepTemplates() {
    const container = document.getElementById("stepTemplatesList");
    const countBadge = document.getElementById("stepTemplateCount");
    
    if (!container) return;
    
    if (countBadge) {
        countBadge.textContent = STEP_TEMPLATES.length;
    }
    
    container.innerHTML = STEP_TEMPLATES.map(step => {
        const tabsHtml = step.suggestedTabs.length > 0
            ? step.suggestedTabs.map(tab => 
                `<span class="step-template-tab">{PREFIX}${tab.suffix}</span>`
            ).join("")
            : `<span class="step-template-tab" style="opacity: 0.5">No tabs</span>`;
        
        return `
            <div class="step-template-card" data-step-id="${step.id}">
                <div class="step-template-icon">${step.icon}</div>
                <div class="step-template-body">
                    <div class="step-template-name">${escapeHtml(step.name)}</div>
                    <div class="step-template-desc">${escapeHtml(step.description)}</div>
                    <div class="step-template-tabs">${tabsHtml}</div>
                </div>
                <div class="step-template-actions">
                    <button type="button" class="step-template-edit" data-step-id="${step.id}">Edit</button>
                </div>
            </div>
        `;
    }).join("");
    
    // Wire up edit buttons
    container.querySelectorAll(".step-template-edit").forEach(btn => {
        btn.addEventListener("click", () => {
            const stepId = btn.dataset.stepId;
            editStepTemplate(stepId);
        });
    });
}

/**
 * Render tab templates list
 */
function renderTabTemplates() {
    const container = document.getElementById("tabTemplatesList");
    const countBadge = document.getElementById("tabTemplateCount");
    
    if (!container) return;
    
    const templateEntries = Object.entries(TAB_TEMPLATES);
    
    if (countBadge) {
        countBadge.textContent = templateEntries.length;
    }
    
    container.innerHTML = templateEntries.map(([id, template]) => {
        const headersHtml = template.headers.map(h => 
            `<span class="tab-template-header-tag">${escapeHtml(h)}</span>`
        ).join("");
        
        return `
            <div class="tab-template-card" data-template-id="${id}">
                <div class="tab-template-header">
                    <span class="tab-template-name">${escapeHtml(template.name)}</span>
                    <span class="tab-template-id">${id}</span>
                </div>
                <div class="tab-template-desc">${escapeHtml(template.description)}</div>
                <div class="tab-template-headers">${headersHtml}</div>
            </div>
        `;
    }).join("");
}

/**
 * Edit a step template (placeholder - would open a modal)
 */
function editStepTemplate(stepId) {
    const step = STEP_TEMPLATES.find(s => s.id === stepId);
    if (!step) return;
    
    // For now, just log - in future, open edit modal
    console.log("Edit step template:", step);
    alert(`Edit Step: ${step.name}\n\nThis feature coming soon!\n\nSuggested tabs:\n${step.suggestedTabs.map(t => "• " + t.suffix).join("\n") || "None"}`);
}

/**
 * Get step template by name
 */
function getStepTemplate(name) {
    return STEP_TEMPLATES.find(s => s.name.toLowerCase() === name.toLowerCase());
}

/**
 * Generate suggested tabs for a module based on selected steps
 */
function generateSuggestedTabs(prefix, stepNames) {
    const tabs = [];
    
    stepNames.forEach(stepName => {
        const template = getStepTemplate(stepName);
        if (template && template.suggestedTabs) {
            template.suggestedTabs.forEach(tab => {
                tabs.push({
                    name: `${prefix}${tab.suffix}`,
                    template: tab.template,
                    step: stepName,
                    description: tab.description
                });
            });
        }
    });
    
    return tabs;
}

// =============================================================================
// MODULE REGISTRY
// =============================================================================

function renderModuleRegistry() {
    const container = document.getElementById("moduleRegistry");
    if (!container) return;
    
    container.innerHTML = moduleRegistry.map((module, index) => `
        <div class="module-registry-item" data-index="${index}">
            <input type="text" class="module-name-input" value="${escapeHtml(module.name)}" placeholder="Module Name">
            <input type="text" class="module-prefix-input" value="${escapeHtml(module.prefix)}" placeholder="Prefix" maxlength="4">
            <button type="button" class="remove-module" data-index="${index}">Remove</button>
        </div>
    `).join("");
    
    // Wire up inputs
    container.querySelectorAll(".module-name-input").forEach((input, index) => {
        input.addEventListener("change", () => {
            moduleRegistry[index].name = input.value;
        });
    });
    
    container.querySelectorAll(".module-prefix-input").forEach((input, index) => {
        input.addEventListener("change", () => {
            moduleRegistry[index].prefix = input.value.toUpperCase();
            input.value = input.value.toUpperCase();
        });
    });
    
    container.querySelectorAll(".remove-module").forEach(btn => {
        btn.addEventListener("click", () => {
            const index = parseInt(btn.dataset.index, 10);
            moduleRegistry.splice(index, 1);
            renderModuleRegistry();
        });
    });
}

function addModuleToRegistry() {
    moduleRegistry.push({
        name: "New Module",
        prefix: "NEW",
        steps: ["Config", "Import", "Review"]
    });
    renderModuleRegistry();
}

// =============================================================================
// TAB CREATION WIZARD
// =============================================================================

// Wizard state
const wizardState = {
    selectedModule: null,
    suggestedTabs: [],
    selectedTabs: new Set(),
    existingTabs: new Set()
};

/**
 * Initialize the Tab Creation Wizard
 */
function initWizard() {
    const moduleSelect = document.getElementById("wizardModuleSelect");
    const createBtn = document.getElementById("wizardCreateBtn");
    
    // Populate module dropdown
    if (moduleSelect) {
        moduleSelect.innerHTML = '<option value="">Choose a module...</option>' +
            moduleRegistry.map(m => 
                `<option value="${escapeHtml(m.prefix)}">${escapeHtml(m.name)} (${m.prefix})</option>`
            ).join("");
        
        moduleSelect.addEventListener("change", onWizardModuleChange);
    }
    
    // Create button handler
    if (createBtn) {
        createBtn.addEventListener("click", executeWizardCreate);
    }
}

/**
 * Handle module selection change
 */
async function onWizardModuleChange() {
    const moduleSelect = document.getElementById("wizardModuleSelect");
    const moduleInfo = document.getElementById("wizardModuleInfo");
    const prefixSpan = document.getElementById("wizardModulePrefix");
    const stepsSpan = document.getElementById("wizardModuleSteps");
    const tabsHint = document.getElementById("wizardTabsHint");
    
    const prefix = moduleSelect?.value;
    
    if (!prefix) {
        wizardState.selectedModule = null;
        wizardState.suggestedTabs = [];
        wizardState.selectedTabs.clear();
        if (moduleInfo) moduleInfo.hidden = true;
        if (tabsHint) tabsHint.textContent = "Select a module first to see suggested tabs.";
        renderWizardTabs();
        updateWizardSummary();
        return;
    }
    
    // Find the module
    const module = moduleRegistry.find(m => m.prefix === prefix);
    if (!module) return;
    
    wizardState.selectedModule = module;
    
    // Show module info
    if (moduleInfo) moduleInfo.hidden = false;
    if (prefixSpan) prefixSpan.textContent = module.prefix;
    if (stepsSpan) stepsSpan.textContent = module.steps.join(", ");
    
    // Generate suggested tabs
    wizardState.suggestedTabs = generateSuggestedTabs(module.prefix, module.steps);
    wizardState.selectedTabs.clear();
    
    // Check which tabs already exist
    await checkExistingTabs();
    
    // Pre-select tabs that don't exist yet
    wizardState.suggestedTabs.forEach(tab => {
        if (!wizardState.existingTabs.has(tab.name)) {
            wizardState.selectedTabs.add(tab.name);
        }
    });
    
    if (tabsHint) {
        tabsHint.textContent = `${wizardState.suggestedTabs.length} tabs suggested based on module steps. Select which to create.`;
    }
    
    renderWizardTabs();
    updateWizardSummary();
}

/**
 * Check which tabs already exist in the workbook
 */
async function checkExistingTabs() {
    wizardState.existingTabs.clear();
    
    if (!hasExcelRuntime()) return;
    
    try {
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();
            
            sheets.items.forEach(sheet => {
                wizardState.existingTabs.add(sheet.name);
            });
        });
    } catch (error) {
        console.error("Error checking existing tabs:", error);
    }
}

/**
 * Render the wizard tabs list
 */
function renderWizardTabs() {
    const container = document.getElementById("wizardTabsList");
    if (!container) return;
    
    if (wizardState.suggestedTabs.length === 0) {
        container.innerHTML = "";
        return;
    }
    
    container.innerHTML = wizardState.suggestedTabs.map(tab => {
        const exists = wizardState.existingTabs.has(tab.name);
        const selected = wizardState.selectedTabs.has(tab.name);
        const template = TAB_TEMPLATES[tab.template];
        
        return `
            <div class="wizard-tab-item ${selected ? 'selected' : ''} ${exists ? 'exists' : ''}" 
                 data-tab-name="${escapeHtml(tab.name)}"
                 ${exists ? 'title="This tab already exists"' : ''}>
                <div class="wizard-tab-checkbox"></div>
                <div class="wizard-tab-body">
                    <div class="wizard-tab-name">${escapeHtml(tab.name)}</div>
                    <div class="wizard-tab-desc">${escapeHtml(tab.description || template?.description || "")}</div>
                </div>
                ${exists 
                    ? '<span class="wizard-tab-exists">Exists</span>'
                    : `<span class="wizard-tab-template">${escapeHtml(tab.template)}</span>`
                }
            </div>
        `;
    }).join("");
    
    // Add custom tab input
    container.innerHTML += `
        <div class="wizard-custom-tab">
            <h5>+ Add Custom Tab</h5>
            <div class="wizard-custom-form">
                <input type="text" id="wizardCustomTabName" class="admin-input" 
                       placeholder="Tab name (e.g., ${wizardState.selectedModule?.prefix || 'XX'}_Custom)">
                <select id="wizardCustomTabTemplate" class="admin-select">
                    <option value="">Template...</option>
                    ${Object.entries(TAB_TEMPLATES).map(([id, t]) => 
                        `<option value="${id}">${t.name}</option>`
                    ).join("")}
                </select>
                <button type="button" class="admin-btn admin-btn--secondary" id="wizardAddCustomBtn">Add</button>
            </div>
        </div>
    `;
    
    // Wire up click handlers
    container.querySelectorAll(".wizard-tab-item").forEach(item => {
        item.addEventListener("click", () => {
            const tabName = item.dataset.tabName;
            if (wizardState.existingTabs.has(tabName)) return; // Can't toggle existing tabs
            
            if (wizardState.selectedTabs.has(tabName)) {
                wizardState.selectedTabs.delete(tabName);
            } else {
                wizardState.selectedTabs.add(tabName);
            }
            
            renderWizardTabs();
            updateWizardSummary();
        });
    });
    
    // Wire up custom tab button
    document.getElementById("wizardAddCustomBtn")?.addEventListener("click", addCustomWizardTab);
}

/**
 * Add a custom tab to the wizard
 */
function addCustomWizardTab() {
    const nameInput = document.getElementById("wizardCustomTabName");
    const templateSelect = document.getElementById("wizardCustomTabTemplate");
    
    const name = nameInput?.value?.trim();
    const template = templateSelect?.value;
    
    if (!name) {
        alert("Please enter a tab name.");
        return;
    }
    
    // Check if already in list
    if (wizardState.suggestedTabs.some(t => t.name === name)) {
        alert("This tab is already in the list.");
        return;
    }
    
    // Add to suggested tabs
    wizardState.suggestedTabs.push({
        name: name,
        template: template || "data-import",
        step: "Custom",
        description: "Custom tab"
    });
    
    // Auto-select if doesn't exist
    if (!wizardState.existingTabs.has(name)) {
        wizardState.selectedTabs.add(name);
    }
    
    // Clear inputs
    if (nameInput) nameInput.value = "";
    if (templateSelect) templateSelect.value = "";
    
    renderWizardTabs();
    updateWizardSummary();
}

/**
 * Update the wizard summary section
 */
function updateWizardSummary() {
    const summary = document.getElementById("wizardSummary");
    const createBtn = document.getElementById("wizardCreateBtn");
    
    const selectedCount = wizardState.selectedTabs.size;
    
    if (!summary) return;
    
    if (selectedCount === 0) {
        summary.innerHTML = '<p class="wizard-hint">Select tabs above to see summary.</p>';
        if (createBtn) createBtn.disabled = true;
        return;
    }
    
    const selectedTabs = wizardState.suggestedTabs.filter(t => wizardState.selectedTabs.has(t.name));
    
    summary.innerHTML = `
        <p class="wizard-summary-title">${selectedCount} tab${selectedCount > 1 ? 's' : ''} will be created:</p>
        <div class="wizard-summary-list">
            ${selectedTabs.map(t => `<span class="wizard-summary-item">${escapeHtml(t.name)}</span>`).join("")}
        </div>
    `;
    
    if (createBtn) createBtn.disabled = false;
}

/**
 * Execute the wizard - create selected tabs
 */
async function executeWizardCreate() {
    const createBtn = document.getElementById("wizardCreateBtn");
    const progress = document.getElementById("wizardProgress");
    const progressFill = document.getElementById("wizardProgressFill");
    const progressText = document.getElementById("wizardProgressText");
    
    if (wizardState.selectedTabs.size === 0) return;
    
    if (!hasExcelRuntime()) {
        alert("Excel not available.");
        return;
    }
    
    // Show progress
    if (createBtn) createBtn.disabled = true;
    if (progress) progress.hidden = false;
    
    const tabsToCreate = wizardState.suggestedTabs.filter(t => wizardState.selectedTabs.has(t.name));
    const total = tabsToCreate.length;
    let created = 0;
    let errors = [];
    
    try {
        for (const tab of tabsToCreate) {
            if (progressText) progressText.textContent = `Creating ${tab.name}...`;
            if (progressFill) progressFill.style.width = `${(created / total) * 100}%`;
            
            try {
                await createExcelTab(tab);
                created++;
            } catch (err) {
                errors.push({ tab: tab.name, error: err.message });
            }
        }
        
        // Complete
        if (progressFill) progressFill.style.width = "100%";
        if (progressText) {
            progressText.textContent = errors.length > 0
                ? `Created ${created} tabs. ${errors.length} failed.`
                : `Successfully created ${created} tabs!`;
        }
        
        // Refresh existing tabs and re-render
        await checkExistingTabs();
        wizardState.selectedTabs.clear();
        renderWizardTabs();
        updateWizardSummary();
        
        // Log results
        logAdmin(`Created ${created} Excel tabs`, "success");
        errors.forEach(e => logAdmin(`Failed: ${e.tab} - ${e.error}`, "error"));
        
    } catch (error) {
        console.error("Wizard create error:", error);
        if (progressText) progressText.textContent = `Error: ${error.message}`;
        logAdmin(`Error: ${error.message}`, "error");
    }
    
    // Re-enable button after delay
    setTimeout(() => {
        if (createBtn) createBtn.disabled = wizardState.selectedTabs.size === 0;
        if (progress) progress.hidden = true;
    }, 3000);
}

/**
 * Create a single Excel tab with headers
 */
async function createExcelTab(tab) {
    const template = TAB_TEMPLATES[tab.template];
    const headers = template?.headers || ["Column1", "Column2", "Column3"];
    const moduleName = wizardState.selectedModule?.name || "Unknown";
    const folder = moduleRegistry.find(m => m.name === moduleName)?.prefix?.toLowerCase() + "-module" || "module";
    
    await Excel.run(async (context) => {
        // Create the worksheet
        const sheet = context.workbook.worksheets.add(tab.name);
        
        // Add headers
        const headerRange = sheet.getRange(`A1:${String.fromCharCode(64 + headers.length)}1`);
        headerRange.values = [headers];
        headerRange.format.font.bold = true;
        headerRange.format.fill.color = "#f0f0f0";
        
        // Auto-fit columns
        headerRange.format.autofitColumns();
        
        await context.sync();
        
        // Register in SS_PF_Config
        await registerTabInConfig(tab.name, moduleName, folder);
    });
}

/**
 * Register a tab in SS_PF_Config
 */
async function registerTabInConfig(tabName, moduleName, folder) {
    // Check if already registered
    if (tabMappings.some(t => t.tabName === tabName && t.module === moduleName)) {
        return; // Already registered
    }
    
    await Excel.run(async (context) => {
        const configSheet = context.workbook.worksheets.getItemOrNullObject("SS_PF_Config");
        configSheet.load("isNullObject");
        await context.sync();
        
        if (configSheet.isNullObject) return;
        
        const usedRange = configSheet.getUsedRange();
        usedRange.load("rowCount");
        await context.sync();
        
        const nextRow = usedRange.rowCount + 1;
        const newRange = configSheet.getRange(`A${nextRow}:D${nextRow}`);
        newRange.values = [["Tab Structure", moduleName, tabName, folder]];
        await context.sync();
    });
    
    // Update local state
    tabMappings.push({
        rowIndex: -1, // Will be corrected on next load
        module: moduleName,
        tabName: tabName,
        folder: folder
    });
}

// =============================================================================
// REFERENCE DATA (Quick Access)
// =============================================================================

/**
 * Initialize Quick Access button handlers
 */
function initQuickAccess() {
    document.getElementById("quickRosterBtn")?.addEventListener("click", () => openRefData("roster"));
    document.getElementById("quickAccountsBtn")?.addEventListener("click", () => openRefData("accounts"));
    
    // Modal close handlers
    document.getElementById("refDataClose")?.addEventListener("click", closeRefData);
    document.getElementById("refDataOverlay")?.addEventListener("click", (e) => {
        if (e.target.id === "refDataOverlay") closeRefData();
    });
    
    // Toolbar buttons
    document.getElementById("refDataRefresh")?.addEventListener("click", refreshRefData);
    document.getElementById("refDataOpenSheet")?.addEventListener("click", openRefDataSheet);
    
    // Edit mode button
    document.getElementById("refDataEditBtn")?.addEventListener("click", showEditMode);
    
    // Edit type selection
    document.getElementById("editTypeQuick")?.addEventListener("click", showQuickEdit);
    document.getElementById("editTypeBulk")?.addEventListener("click", showBulkUpload);
    document.getElementById("editBackToView")?.addEventListener("click", showViewMode);
    
    // Quick Edit: Employee
    document.getElementById("quickEditBackEmployee")?.addEventListener("click", showEditTypeSelection);
    document.getElementById("quickEditCancelEmployee")?.addEventListener("click", resetQuickEditEmployee);
    document.getElementById("quickEditConfirmEmployee")?.addEventListener("click", confirmQuickEditEmployee);
    
    // Quick Edit: Accounts
    document.getElementById("quickEditBackAccounts")?.addEventListener("click", showEditTypeSelection);
    document.getElementById("quickEditCancelAccounts")?.addEventListener("click", resetQuickEditAccounts);
    document.getElementById("quickEditConfirmAccounts")?.addEventListener("click", confirmQuickEditAccounts);
    
    // Bulk Upload
    document.getElementById("bulkUploadBack")?.addEventListener("click", showEditTypeSelection);
    document.getElementById("bulkUploadCancel")?.addEventListener("click", resetBulkUpload);
    
    const confirmBtn = document.getElementById("bulkUploadConfirm");
    if (confirmBtn) {
        console.log("Attaching click handler to bulkUploadConfirm button");
        confirmBtn.addEventListener("click", () => {
            console.log("bulkUploadConfirm button clicked!");
            confirmBulkUpload();
        });
    } else {
        console.warn("bulkUploadConfirm button not found in DOM");
    }
    
    // Bulk upload dropzone
    initBulkUploadDropzone();
    console.log("initQuickAccess complete");
}

/**
 * Open Reference Data modal
 */
async function openRefData(type) {
    const config = REF_DATA_CONFIG[type];
    if (!config) return;
    
    refDataState.type = type;
    
    const overlay = document.getElementById("refDataOverlay");
    const title = document.getElementById("refDataTitle");
    
    if (title) title.textContent = config.title;
    if (overlay) overlay.hidden = false;
    
    // Reset all edit mode states when opening
    resetAllEditStates();
    
    // Show view mode (not edit mode)
    showViewMode();
    
    await loadRefData();
}

/**
 * Reset all edit states to initial values
 */
function resetAllEditStates() {
    // Reset edit state object
    editState.mode = null;
    editState.action = null;
    editState.selectedEmployee = null;
    editState.selectedAccount = null;
    editState.formData = {};
    editState.uploadedHeaders = [];
    editState.uploadedData = [];
    editState.pendingChanges = [];
    
    // Reset bulk upload UI
    resetBulkUpload();
    resetBulkUploadButton();
    
    // Hide all edit panels
    const editTypeSelection = document.getElementById("editTypeSelection");
    const quickEditEmployee = document.getElementById("quickEditEmployee");
    const quickEditAccounts = document.getElementById("quickEditAccounts");
    const bulkUploadPanel = document.getElementById("bulkUploadPanel");
    
    if (editTypeSelection) editTypeSelection.hidden = true;
    if (quickEditEmployee) quickEditEmployee.hidden = true;
    if (quickEditAccounts) quickEditAccounts.hidden = true;
    if (bulkUploadPanel) bulkUploadPanel.hidden = true;
}

/**
 * Close Reference Data modal
 */
function closeRefData() {
    const overlay = document.getElementById("refDataOverlay");
    if (overlay) overlay.hidden = true;
    refDataState.type = null;
    refDataState.data = [];
    refDataState.headers = [];
    
    // Reset all edit states when closing
    resetAllEditStates();
}

/**
 * Load Reference Data from Excel
 */
async function loadRefData() {
    const config = REF_DATA_CONFIG[refDataState.type];
    if (!config || !hasExcelRuntime()) return;
    
    const loading = document.getElementById("refDataLoading");
    const empty = document.getElementById("refDataEmpty");
    const rowCount = document.getElementById("refDataRowCount");
    const lastUpdated = document.getElementById("refDataLastUpdated");
    const sheetName = document.getElementById("refDataSheetName");
    const sheetBadge = document.getElementById("refDataSheetBadge");
    const actions = document.querySelector(".ref-data-actions");
    const summary = document.querySelector(".ref-data-summary");
    const sheetStatus = document.querySelector(".ref-data-sheet-status");
    
    if (loading) loading.hidden = false;
    if (empty) empty.hidden = true;
    if (sheetName) sheetName.textContent = config.sheetName;
    
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItemOrNullObject(config.sheetName);
            sheet.load("isNullObject");
            await context.sync();
            
            if (sheet.isNullObject) {
                // Sheet doesn't exist - show actions so user can create it
                refDataState.headers = config.defaultHeaders;
                refDataState.data = [];
                if (loading) loading.hidden = true;
                if (rowCount) rowCount.textContent = "0";
                if (lastUpdated) lastUpdated.textContent = "—";
                if (sheetBadge) {
                    sheetBadge.textContent = "Missing";
                    sheetBadge.classList.add("missing");
                }
                // Still show actions - Open in Excel will create the sheet
                if (actions) actions.style.display = '';
                if (summary) summary.style.display = '';
                if (sheetStatus) sheetStatus.style.display = '';
                return;
            }
            
            const usedRange = sheet.getUsedRangeOrNullObject();
            usedRange.load("values");
            await context.sync();
            
            if (usedRange.isNullObject || !usedRange.values || usedRange.values.length === 0) {
                refDataState.headers = config.defaultHeaders;
                refDataState.data = [];
            } else {
                refDataState.headers = usedRange.values[0].map(h => String(h || ""));
                refDataState.data = usedRange.values.slice(1);
            }
            
            if (loading) loading.hidden = true;
            if (actions) actions.style.display = '';
            if (summary) summary.style.display = '';
            if (sheetStatus) sheetStatus.style.display = '';
            if (rowCount) rowCount.textContent = refDataState.data.length.toString();
            if (sheetBadge) {
                sheetBadge.textContent = "Ready";
                sheetBadge.classList.remove("missing");
            }
            
            // Set last updated to now (in a real app, this would come from sheet metadata)
            if (lastUpdated) {
                const now = new Date();
                const options = { month: 'short', day: 'numeric' };
                lastUpdated.textContent = now.toLocaleDateString('en-US', options);
            }
        });
    } catch (error) {
        console.error("Error loading reference data:", error);
        if (loading) loading.hidden = true;
        if (actions) actions.style.display = '';
        if (summary) summary.style.display = '';
        if (sheetStatus) sheetStatus.style.display = '';
        if (rowCount) rowCount.textContent = "0";
        if (lastUpdated) lastUpdated.textContent = "—";
        if (sheetBadge) {
            sheetBadge.textContent = "Error";
            sheetBadge.classList.add("missing");
        }
    }
}

/**
 * Refresh Reference Data
 */
async function refreshRefData() {
    await loadRefData();
}

/**
 * Open Reference Data sheet in Excel
 * Makes sheet visible first if it's hidden by tab visibility
 */
async function openRefDataSheet() {
    const config = REF_DATA_CONFIG[refDataState.type];
    if (!config || !hasExcelRuntime()) return;
    
    try {
        await Excel.run(async (context) => {
            let sheet = context.workbook.worksheets.getItemOrNullObject(config.sheetName);
            sheet.load("isNullObject,visibility");
            await context.sync();
            
            if (sheet.isNullObject) {
                // Create the sheet if it doesn't exist
                sheet = context.workbook.worksheets.add(config.sheetName);
                const headerRange = sheet.getRange(`A1:${String.fromCharCode(64 + config.defaultHeaders.length)}1`);
                headerRange.values = [config.defaultHeaders];
                headerRange.format.font.bold = true;
                headerRange.format.fill.color = "#f0f0f0";
                await context.sync();
            } else {
                // Make sure sheet is visible before activating (may be hidden by tab visibility)
                sheet.visibility = Excel.SheetVisibility.visible;
                await context.sync();
            }
            
            sheet.activate();
            sheet.getRange("A1").select();
            await context.sync();
            console.log(`[Module Selector] Opened sheet: ${config.sheetName}`);
        });
        
        closeRefData();
    } catch (error) {
        console.error("Error opening sheet:", error);
        alert("Error opening sheet: " + error.message);
    }
}

/**
 * Create Reference Data sheet
 */
async function createRefDataSheet() {
    const config = REF_DATA_CONFIG[refDataState.type];
    if (!config || !hasExcelRuntime()) return;
    
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.add(config.sheetName);
            const headerRange = sheet.getRange(`A1:${String.fromCharCode(64 + config.defaultHeaders.length)}1`);
            headerRange.values = [config.defaultHeaders];
            headerRange.format.font.bold = true;
            headerRange.format.fill.color = "#f0f0f0";
            headerRange.format.autofitColumns();
            await context.sync();
            
            // Register in SS_PF_Config
            const configSheet = context.workbook.worksheets.getItemOrNullObject("SS_PF_Config");
            configSheet.load("isNullObject");
            await context.sync();
            
            if (!configSheet.isNullObject) {
                const usedRange = configSheet.getUsedRange();
                usedRange.load("rowCount");
                await context.sync();
                
                const nextRow = usedRange.rowCount + 1;
                configSheet.getRange(`A${nextRow}:D${nextRow}`).values = [
                    ["Tab Structure", "Shared", config.sheetName, "N/A"]
                ];
                await context.sync();
            }
        });
        
        // Reload the data
        await loadRefData();
        
    } catch (error) {
        console.error("Error creating sheet:", error);
        alert("Error creating sheet: " + error.message);
    }
}

// Make createRefDataSheet available globally for onclick
window.createRefDataSheet = createRefDataSheet;

// =============================================================================
// EDIT MODE - Reference Data Editing Workflow
// =============================================================================

// Edit mode state
const editState = {
    mode: null, // 'quick' or 'bulk'
    action: null, // 'add', 'terminate', 'transfer', 'edit', 'delete', 'rename', 'renumber'
    selectedRecord: null,
    pendingChanges: [],
    uploadedData: [],
    uploadedHeaders: []
};

/**
 * Show the edit mode selection screen
 */
function showEditMode() {
    const viewMode = document.getElementById("refDataViewMode");
    const editMode = document.getElementById("refDataEditMode");
    const editTypeSelection = document.getElementById("editTypeSelection");
    
    if (viewMode) viewMode.hidden = true;
    if (editMode) editMode.hidden = false;
    if (editTypeSelection) editTypeSelection.hidden = false;
    
    // Hide all edit panels
    document.getElementById("quickEditEmployee")?.setAttribute("hidden", "");
    document.getElementById("quickEditAccounts")?.setAttribute("hidden", "");
    document.getElementById("bulkUploadPanel")?.setAttribute("hidden", "");
    
    // Reset state
    editState.mode = null;
    editState.action = null;
    editState.selectedRecord = null;
    editState.pendingChanges = [];
}

/**
 * Show view mode (close edit mode)
 */
function showViewMode() {
    const viewMode = document.getElementById("refDataViewMode");
    const editMode = document.getElementById("refDataEditMode");
    
    if (viewMode) viewMode.hidden = false;
    if (editMode) editMode.hidden = true;
    
    // Refresh data
    loadRefData();
}

/**
 * Show edit type selection
 */
function showEditTypeSelection() {
    const editTypeSelection = document.getElementById("editTypeSelection");
    
    if (editTypeSelection) editTypeSelection.hidden = false;
    
    document.getElementById("quickEditEmployee")?.setAttribute("hidden", "");
    document.getElementById("quickEditAccounts")?.setAttribute("hidden", "");
    document.getElementById("bulkUploadPanel")?.setAttribute("hidden", "");
    
    editState.mode = null;
    editState.action = null;
}

/**
 * Show Quick Edit panel based on data type
 */
function showQuickEdit() {
    const editTypeSelection = document.getElementById("editTypeSelection");
    if (editTypeSelection) editTypeSelection.hidden = true;
    
    editState.mode = 'quick';
    
    if (refDataState.type === 'roster') {
        const panel = document.getElementById("quickEditEmployee");
        if (panel) panel.hidden = false;
        wireQuickEditEmployeeActions();
    } else if (refDataState.type === 'accounts') {
        const panel = document.getElementById("quickEditAccounts");
        if (panel) panel.hidden = false;
        wireQuickEditAccountsActions();
    }
}

/**
 * Show Bulk Upload panel
 */
function showBulkUpload() {
    const editTypeSelection = document.getElementById("editTypeSelection");
    if (editTypeSelection) editTypeSelection.hidden = true;
    
    const panel = document.getElementById("bulkUploadPanel");
    if (panel) panel.hidden = false;
    
    const title = document.getElementById("bulkUploadTitle");
    const config = REF_DATA_CONFIG[refDataState.type];
    if (title && config) {
        title.textContent = `Bulk Upload: ${config.title}`;
    }
    
    editState.mode = 'bulk';
    resetBulkUpload();
    
    // Re-initialize dropzone (in case it wasn't ready before)
    initBulkUploadDropzone();
}

// =============================================================================
// QUICK EDIT: EMPLOYEE ROSTER
// =============================================================================

/**
 * Wire up quick edit action buttons for employee roster
 */
function wireQuickEditEmployeeActions() {
    const actions = document.querySelectorAll("#quickEditEmployee .quick-action-btn");
    actions.forEach(btn => {
        btn.addEventListener("click", () => {
            const action = btn.dataset.action;
            selectQuickEditAction("employee", action);
            
            // Update button states
            actions.forEach(b => b.classList.remove("active"));
            btn.classList.add("active");
        });
    });
}

/**
 * Select a quick edit action and show appropriate form
 */
function selectQuickEditAction(type, action) {
    editState.action = action;
    editState.selectedRecord = null;
    
    const form = document.getElementById(`quickEditForm${type === 'employee' ? 'Employee' : 'Accounts'}`);
    const preview = document.getElementById(`quickEditPreview${type === 'employee' ? 'Employee' : 'Accounts'}`);
    const buttons = document.getElementById(`quickEditButtons${type === 'employee' ? 'Employee' : 'Accounts'}`);
    
    if (preview) preview.hidden = true;
    if (buttons) buttons.hidden = true;
    
    if (type === 'employee') {
        renderEmployeeForm(action, form);
    } else {
        renderAccountsForm(action, form);
    }
}

/**
 * Render the employee quick edit form based on action
 */
function renderEmployeeForm(action, form) {
    if (!form) return;
    
    const employees = refDataState.data.map((row, idx) => ({
        index: idx,
        name: row[0] || '',
        department: row[1] || '',
        payRate: row[2] || '',
        status: row[3] || 'Active',
        hireDate: row[4] || '',
        termDate: row[5] || ''
    }));
    
    switch (action) {
        case 'add':
            form.innerHTML = `
                <div class="form-group">
                    <label>Employee Name</label>
                    <input type="text" id="empAddName" placeholder="First Last">
                </div>
                <div class="form-group">
                    <label>Department</label>
                    <input type="text" id="empAddDept" placeholder="Department name">
                </div>
                <div class="form-group">
                    <label>Pay Rate</label>
                    <input type="number" id="empAddPayRate" placeholder="0.00" step="0.01">
                </div>
                <div class="form-group">
                    <label>Hire Date</label>
                    <input type="date" id="empAddHireDate">
                </div>
            `;
            wireEmployeeAddForm();
            break;
            
        case 'terminate':
            form.innerHTML = `
                <div class="form-group">
                    <label>Search Employee</label>
                    <input type="text" id="empTermSearch" placeholder="Start typing name...">
                    <div class="search-results" id="empTermResults"></div>
                </div>
                <div class="form-group" id="empTermDateGroup" hidden>
                    <label>Termination Date</label>
                    <input type="date" id="empTermDate" value="${new Date().toISOString().split('T')[0]}">
                </div>
            `;
            wireEmployeeSearchForm('terminate', employees);
            break;
            
        case 'transfer':
            form.innerHTML = `
                <div class="form-group">
                    <label>Search Employee</label>
                    <input type="text" id="empTransferSearch" placeholder="Start typing name...">
                    <div class="search-results" id="empTransferResults"></div>
                </div>
                <div class="form-group" id="empTransferDeptGroup" hidden>
                    <label>New Department</label>
                    <input type="text" id="empTransferDept" placeholder="New department name">
                </div>
            `;
            wireEmployeeSearchForm('transfer', employees);
            break;
            
        case 'edit':
            form.innerHTML = `
                <div class="form-group">
                    <label>Search Employee</label>
                    <input type="text" id="empEditSearch" placeholder="Start typing name...">
                    <div class="search-results" id="empEditResults"></div>
                </div>
                <div id="empEditFields" hidden>
                    <div class="form-group">
                        <label>Employee Name</label>
                        <input type="text" id="empEditName">
                    </div>
                    <div class="form-group">
                        <label>Department</label>
                        <input type="text" id="empEditDept">
                    </div>
                    <div class="form-group">
                        <label>Pay Rate</label>
                        <input type="number" id="empEditPayRate" step="0.01">
                    </div>
                    <div class="form-group">
                        <label>Status</label>
                        <select id="empEditStatus">
                            <option value="Active">Active</option>
                            <option value="Terminated">Terminated</option>
                            <option value="On Leave">On Leave</option>
                        </select>
                    </div>
                </div>
            `;
            wireEmployeeSearchForm('edit', employees);
            break;
    }
}

/**
 * Wire up the Add Employee form
 */
function wireEmployeeAddForm() {
    const inputs = ['empAddName', 'empAddDept', 'empAddPayRate', 'empAddHireDate'];
    inputs.forEach(id => {
        document.getElementById(id)?.addEventListener("input", updateEmployeeAddPreview);
    });
}

/**
 * Update preview for Add Employee
 */
function updateEmployeeAddPreview() {
    const name = document.getElementById("empAddName")?.value?.trim() || '';
    const dept = document.getElementById("empAddDept")?.value?.trim() || '';
    const payRate = document.getElementById("empAddPayRate")?.value || '';
    const hireDate = document.getElementById("empAddHireDate")?.value || '';
    
    const preview = document.getElementById("quickEditPreviewEmployee");
    const buttons = document.getElementById("quickEditButtonsEmployee");
    
    if (!name) {
        if (preview) preview.hidden = true;
        if (buttons) buttons.hidden = true;
        return;
    }
    
    editState.pendingChanges = [{
        type: 'add',
        data: { name, department: dept, payRate, status: 'Active', hireDate }
    }];
    
    if (preview) {
        preview.hidden = false;
        preview.innerHTML = `
            <h4>Preview</h4>
            <div class="preview-change">
                <span class="preview-icon">➕</span>
                <div class="preview-text">
                    <div class="preview-field">${escapeHtml(name)} will be added</div>
                    <div class="preview-values">
                        Department: ${escapeHtml(dept) || '—'}<br>
                        Pay Rate: ${payRate ? '$' + parseFloat(payRate).toFixed(2) : '—'}<br>
                        Hire Date: ${hireDate || '—'}
                    </div>
                </div>
            </div>
        `;
    }
    
    if (buttons) buttons.hidden = false;
}

/**
 * Wire up employee search form
 */
function wireEmployeeSearchForm(actionType, employees) {
    const searchId = `emp${actionType.charAt(0).toUpperCase() + actionType.slice(1)}Search`;
    const resultsId = `emp${actionType.charAt(0).toUpperCase() + actionType.slice(1)}Results`;
    
    const searchInput = document.getElementById(searchId);
    const resultsContainer = document.getElementById(resultsId);
    
    if (!searchInput || !resultsContainer) return;
    
    searchInput.addEventListener("input", () => {
        const query = searchInput.value.toLowerCase().trim();
        
        if (query.length < 2) {
            resultsContainer.innerHTML = '';
            return;
        }
        
        const matches = employees.filter(emp => 
            emp.name.toLowerCase().includes(query) ||
            emp.department.toLowerCase().includes(query)
        ).slice(0, 10);
        
        if (matches.length === 0) {
            resultsContainer.innerHTML = '<div class="search-result-item">No matches found</div>';
            return;
        }
        
        resultsContainer.innerHTML = matches.map(emp => `
            <div class="search-result-item" data-index="${emp.index}">
                <div class="search-result-name">${escapeHtml(emp.name)}</div>
                <div class="search-result-detail">${escapeHtml(emp.department)} • ${escapeHtml(emp.status)}</div>
            </div>
        `).join('');
        
        resultsContainer.querySelectorAll(".search-result-item").forEach(item => {
            item.addEventListener("click", () => {
                const idx = parseInt(item.dataset.index, 10);
                selectEmployee(actionType, employees[idx]);
                
                // Update visual
                resultsContainer.querySelectorAll(".search-result-item").forEach(i => i.classList.remove("selected"));
                item.classList.add("selected");
            });
        });
    });
}

/**
 * Select an employee for editing
 */
function selectEmployee(actionType, employee) {
    editState.selectedRecord = employee;
    
    const preview = document.getElementById("quickEditPreviewEmployee");
    const buttons = document.getElementById("quickEditButtonsEmployee");
    
    switch (actionType) {
        case 'terminate':
            document.getElementById("empTermDateGroup").hidden = false;
            updateTerminatePreview(employee);
            
            document.getElementById("empTermDate")?.addEventListener("change", () => {
                updateTerminatePreview(employee);
            });
            break;
            
        case 'transfer':
            document.getElementById("empTransferDeptGroup").hidden = false;
            const transferDeptInput = document.getElementById("empTransferDept");
            if (transferDeptInput) {
                transferDeptInput.addEventListener("input", () => {
                    updateTransferPreview(employee, transferDeptInput.value);
                });
            }
            break;
            
        case 'edit':
            document.getElementById("empEditFields").hidden = false;
            document.getElementById("empEditName").value = employee.name;
            document.getElementById("empEditDept").value = employee.department;
            document.getElementById("empEditPayRate").value = employee.payRate;
            document.getElementById("empEditStatus").value = employee.status;
            
            ['empEditName', 'empEditDept', 'empEditPayRate', 'empEditStatus'].forEach(id => {
                document.getElementById(id)?.addEventListener("input", () => updateEditEmployeePreview(employee));
                document.getElementById(id)?.addEventListener("change", () => updateEditEmployeePreview(employee));
            });
            break;
    }
}

/**
 * Update preview for Terminate action
 */
function updateTerminatePreview(employee) {
    const termDate = document.getElementById("empTermDate")?.value || '';
    
    editState.pendingChanges = [{
        type: 'terminate',
        index: employee.index,
        data: { ...employee, status: 'Terminated', termDate }
    }];
    
    const preview = document.getElementById("quickEditPreviewEmployee");
    const buttons = document.getElementById("quickEditButtonsEmployee");
    
    if (preview) {
        preview.hidden = false;
        preview.innerHTML = `
            <h4>Preview</h4>
            <div class="preview-change">
                <span class="preview-icon">❌</span>
                <div class="preview-text">
                    <div class="preview-field">${escapeHtml(employee.name)} will be marked as Terminated</div>
                    <div class="preview-values">
                        Status: <span class="preview-old">${escapeHtml(employee.status)}</span>
                        <span class="preview-arrow">→</span>
                        <span class="preview-new">Terminated</span><br>
                        Term Date: ${termDate || '—'}
                    </div>
                </div>
            </div>
        `;
    }
    
    if (buttons) buttons.hidden = false;
}

/**
 * Update preview for Transfer action
 */
function updateTransferPreview(employee, newDept) {
    if (!newDept?.trim()) {
        document.getElementById("quickEditPreviewEmployee").hidden = true;
        document.getElementById("quickEditButtonsEmployee").hidden = true;
        return;
    }
    
    editState.pendingChanges = [{
        type: 'transfer',
        index: employee.index,
        data: { name: employee.name, department: newDept.trim(), payRate: employee.payRate, status: employee.status, hireDate: employee.hireDate, termDate: employee.termDate }
    }];
    
    const preview = document.getElementById("quickEditPreviewEmployee");
    const buttons = document.getElementById("quickEditButtonsEmployee");
    
    if (preview) {
        preview.hidden = false;
        preview.innerHTML = `
            <h4>Preview</h4>
            <div class="preview-change">
                <span class="preview-icon">🔄</span>
                <div class="preview-text">
                    <div class="preview-field">${escapeHtml(employee.name)} will be transferred</div>
                    <div class="preview-values">
                        Department: <span class="preview-old">${escapeHtml(employee.department)}</span>
                        <span class="preview-arrow">→</span>
                        <span class="preview-new">${escapeHtml(newDept.trim())}</span>
                    </div>
                </div>
            </div>
        `;
    }
    
    if (buttons) buttons.hidden = false;
}

/**
 * Update preview for Edit Employee action
 */
function updateEditEmployeePreview(originalEmployee) {
    const newName = document.getElementById("empEditName")?.value?.trim() || '';
    const newDept = document.getElementById("empEditDept")?.value?.trim() || '';
    const newPayRate = document.getElementById("empEditPayRate")?.value || '';
    const newStatus = document.getElementById("empEditStatus")?.value || '';
    
    const changes = [];
    
    if (newName !== originalEmployee.name) {
        changes.push({ field: 'Name', old: originalEmployee.name, new: newName });
    }
    if (newDept !== originalEmployee.department) {
        changes.push({ field: 'Department', old: originalEmployee.department, new: newDept });
    }
    if (newPayRate !== originalEmployee.payRate) {
        changes.push({ field: 'Pay Rate', old: originalEmployee.payRate, new: newPayRate });
    }
    if (newStatus !== originalEmployee.status) {
        changes.push({ field: 'Status', old: originalEmployee.status, new: newStatus });
    }
    
    const preview = document.getElementById("quickEditPreviewEmployee");
    const buttons = document.getElementById("quickEditButtonsEmployee");
    
    if (changes.length === 0) {
        if (preview) preview.hidden = true;
        if (buttons) buttons.hidden = true;
        return;
    }
    
    editState.pendingChanges = [{
        type: 'edit',
        index: originalEmployee.index,
        data: { name: newName, department: newDept, payRate: newPayRate, status: newStatus, hireDate: originalEmployee.hireDate, termDate: originalEmployee.termDate }
    }];
    
    if (preview) {
        preview.hidden = false;
        preview.innerHTML = `
            <h4>Preview</h4>
            ${changes.map(c => `
                <div class="preview-change">
                    <span class="preview-icon">✏️</span>
                    <div class="preview-text">
                        <div class="preview-field">${escapeHtml(c.field)}</div>
                        <div class="preview-values">
                            <span class="preview-old">${escapeHtml(c.old || '—')}</span>
                            <span class="preview-arrow">→</span>
                            <span class="preview-new">${escapeHtml(c.new || '—')}</span>
                        </div>
                    </div>
                </div>
            `).join('')}
        `;
    }
    
    if (buttons) buttons.hidden = false;
}

/**
 * Reset Quick Edit Employee form
 */
function resetQuickEditEmployee() {
    const form = document.getElementById("quickEditFormEmployee");
    const preview = document.getElementById("quickEditPreviewEmployee");
    const buttons = document.getElementById("quickEditButtonsEmployee");
    
    if (form) form.innerHTML = '';
    if (preview) preview.hidden = true;
    if (buttons) buttons.hidden = true;
    
    // Reset action buttons
    document.querySelectorAll("#quickEditEmployee .quick-action-btn").forEach(btn => {
        btn.classList.remove("active");
    });
    
    editState.action = null;
    editState.selectedRecord = null;
    editState.pendingChanges = [];
}

/**
 * Confirm and apply quick edit changes for Employee
 */
async function confirmQuickEditEmployee() {
    if (editState.pendingChanges.length === 0) return;
    
    if (!hasExcelRuntime()) {
        alert("Excel not available");
        return;
    }
    
    const config = REF_DATA_CONFIG.roster;
    
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(config.sheetName);
            
            for (const change of editState.pendingChanges) {
                if (change.type === 'add') {
                    // Append new row
                    const usedRange = sheet.getUsedRange();
                    usedRange.load("rowCount");
                    await context.sync();
                    
                    const newRow = usedRange.rowCount + 1;
                    const data = change.data;
                    // Columns: Employee, Department, Pay_Rate, Status, Hire_Date, Term_Date
                    const newRange = sheet.getRange(`A${newRow}:F${newRow}`);
                    newRange.values = [[data.name, data.department, data.payRate, data.status, data.hireDate, ""]];
                    
                } else if (change.type === 'terminate') {
                    // Update existing row with termination
                    const rowNum = change.index + 2;
                    const data = change.data;
                    const updateRange = sheet.getRange(`A${rowNum}:F${rowNum}`);
                    updateRange.values = [[data.name, data.department, data.payRate, data.status, data.hireDate, data.termDate || ""]];
                    
                } else if (change.type === 'transfer' || change.type === 'edit') {
                    // Update existing row (preserve term date if exists)
                    const rowNum = change.index + 2;
                    const data = change.data;
                    const updateRange = sheet.getRange(`A${rowNum}:F${rowNum}`);
                    updateRange.values = [[data.name, data.department, data.payRate, data.status, data.hireDate, data.termDate || ""]];
                }
            }
            
            await context.sync();
        });
        
        // Success - refresh and close edit mode
        alert("Changes saved successfully!");
        showViewMode();
        
    } catch (error) {
        console.error("Error saving changes:", error);
        alert("Error saving changes: " + error.message);
    }
}

// =============================================================================
// QUICK EDIT: CHART OF ACCOUNTS
// =============================================================================

/**
 * Wire up quick edit action buttons for accounts
 */
function wireQuickEditAccountsActions() {
    const actions = document.querySelectorAll("#quickEditAccounts .quick-action-btn");
    actions.forEach(btn => {
        btn.addEventListener("click", () => {
            const action = btn.dataset.action;
            selectQuickEditAction("accounts", action);
            
            // Update button states
            actions.forEach(b => b.classList.remove("active"));
            btn.classList.add("active");
        });
    });
}

/**
 * Render the accounts quick edit form based on action
 */
function renderAccountsForm(action, form) {
    if (!form) return;
    
    const accounts = refDataState.data.map((row, idx) => ({
        index: idx,
        number: row[0] || '',
        name: row[1] || '',
        type: row[2] || '',
        category: row[3] || ''
    }));
    
    switch (action) {
        case 'add':
            form.innerHTML = `
                <div class="form-group">
                    <label>Account Number</label>
                    <input type="text" id="acctAddNumber" placeholder="e.g., 6200">
                </div>
                <div class="form-group">
                    <label>Account Name</label>
                    <input type="text" id="acctAddName" placeholder="e.g., Payroll Expense">
                </div>
                <div class="form-group">
                    <label>Type</label>
                    <select id="acctAddType">
                        <option value="">Select...</option>
                        <option value="Asset">Asset</option>
                        <option value="Liability">Liability</option>
                        <option value="Equity">Equity</option>
                        <option value="Revenue">Revenue</option>
                        <option value="Expense">Expense</option>
                    </select>
                </div>
                <div class="form-group">
                    <label>Category</label>
                    <input type="text" id="acctAddCategory" placeholder="e.g., Operating">
                </div>
            `;
            wireAccountsAddForm();
            break;
            
        case 'delete':
            form.innerHTML = `
                <div class="form-group">
                    <label>Search Account</label>
                    <input type="text" id="acctDeleteSearch" placeholder="Search by number or name...">
                    <div class="search-results" id="acctDeleteResults"></div>
                </div>
            `;
            wireAccountsSearchForm('delete', accounts);
            break;
            
        case 'rename':
            form.innerHTML = `
                <div class="form-group">
                    <label>Search Account</label>
                    <input type="text" id="acctRenameSearch" placeholder="Search by number or name...">
                    <div class="search-results" id="acctRenameResults"></div>
                </div>
                <div class="form-group" id="acctRenameNewGroup" hidden>
                    <label>New Account Name</label>
                    <input type="text" id="acctRenameNewName" placeholder="Enter new name">
                </div>
            `;
            wireAccountsSearchForm('rename', accounts);
            break;
            
        case 'renumber':
            form.innerHTML = `
                <div class="form-group">
                    <label>Search Account</label>
                    <input type="text" id="acctRenumberSearch" placeholder="Search by number or name...">
                    <div class="search-results" id="acctRenumberResults"></div>
                </div>
                <div class="form-group" id="acctRenumberNewGroup" hidden>
                    <label>New Account Number</label>
                    <input type="text" id="acctRenumberNewNumber" placeholder="Enter new number">
                </div>
            `;
            wireAccountsSearchForm('renumber', accounts);
            break;
    }
}

/**
 * Wire up the Add Account form
 */
function wireAccountsAddForm() {
    const inputs = ['acctAddNumber', 'acctAddName', 'acctAddType', 'acctAddCategory'];
    inputs.forEach(id => {
        document.getElementById(id)?.addEventListener("input", updateAccountsAddPreview);
        document.getElementById(id)?.addEventListener("change", updateAccountsAddPreview);
    });
}

/**
 * Update preview for Add Account
 */
function updateAccountsAddPreview() {
    const number = document.getElementById("acctAddNumber")?.value?.trim() || '';
    const name = document.getElementById("acctAddName")?.value?.trim() || '';
    const type = document.getElementById("acctAddType")?.value || '';
    const category = document.getElementById("acctAddCategory")?.value?.trim() || '';
    
    const preview = document.getElementById("quickEditPreviewAccounts");
    const buttons = document.getElementById("quickEditButtonsAccounts");
    
    if (!number || !name) {
        if (preview) preview.hidden = true;
        if (buttons) buttons.hidden = true;
        return;
    }
    
    editState.pendingChanges = [{
        type: 'add',
        data: { number, name, type, category }
    }];
    
    if (preview) {
        preview.hidden = false;
        preview.innerHTML = `
            <h4>Preview</h4>
            <div class="preview-change">
                <span class="preview-icon">➕</span>
                <div class="preview-text">
                    <div class="preview-field">New account will be added:</div>
                    <div class="preview-values">
                        ${escapeHtml(number)} - ${escapeHtml(name)}<br>
                        Type: ${escapeHtml(type) || '—'} | Category: ${escapeHtml(category) || '—'}
                    </div>
                </div>
            </div>
        `;
    }
    
    if (buttons) buttons.hidden = false;
}

/**
 * Wire up accounts search form
 */
function wireAccountsSearchForm(actionType, accounts) {
    const searchId = `acct${actionType.charAt(0).toUpperCase() + actionType.slice(1)}Search`;
    const resultsId = `acct${actionType.charAt(0).toUpperCase() + actionType.slice(1)}Results`;
    
    const searchInput = document.getElementById(searchId);
    const resultsContainer = document.getElementById(resultsId);
    
    if (!searchInput || !resultsContainer) return;
    
    searchInput.addEventListener("input", () => {
        const query = searchInput.value.toLowerCase().trim();
        
        if (query.length < 1) {
            resultsContainer.innerHTML = '';
            return;
        }
        
        const matches = accounts.filter(acct => 
            acct.number.toLowerCase().includes(query) ||
            acct.name.toLowerCase().includes(query)
        ).slice(0, 10);
        
        if (matches.length === 0) {
            resultsContainer.innerHTML = '<div class="search-result-item">No matches found</div>';
            return;
        }
        
        resultsContainer.innerHTML = matches.map(acct => `
            <div class="search-result-item" data-index="${acct.index}">
                <div class="search-result-name">${escapeHtml(acct.number)} - ${escapeHtml(acct.name)}</div>
                <div class="search-result-detail">${escapeHtml(acct.type)} | ${escapeHtml(acct.category)}</div>
            </div>
        `).join('');
        
        resultsContainer.querySelectorAll(".search-result-item").forEach(item => {
            item.addEventListener("click", () => {
                const idx = parseInt(item.dataset.index, 10);
                selectAccount(actionType, accounts[idx]);
                
                // Update visual
                resultsContainer.querySelectorAll(".search-result-item").forEach(i => i.classList.remove("selected"));
                item.classList.add("selected");
            });
        });
    });
}

/**
 * Select an account for editing
 */
function selectAccount(actionType, account) {
    editState.selectedRecord = account;
    
    const preview = document.getElementById("quickEditPreviewAccounts");
    const buttons = document.getElementById("quickEditButtonsAccounts");
    
    switch (actionType) {
        case 'delete':
            editState.pendingChanges = [{
                type: 'delete',
                index: account.index,
                data: account
            }];
            
            if (preview) {
                preview.hidden = false;
                preview.innerHTML = `
                    <h4>Preview</h4>
                    <div class="preview-change">
                        <span class="preview-icon">❌</span>
                        <div class="preview-text">
                            <div class="preview-field">Account will be deleted:</div>
                            <div class="preview-values">
                                ${escapeHtml(account.number)} - ${escapeHtml(account.name)}
                            </div>
                        </div>
                    </div>
                `;
            }
            if (buttons) buttons.hidden = false;
            break;
            
        case 'rename':
            document.getElementById("acctRenameNewGroup").hidden = false;
            const renameInput = document.getElementById("acctRenameNewName");
            if (renameInput) {
                renameInput.addEventListener("input", () => {
                    updateRenameAccountPreview(account, renameInput.value);
                });
            }
            break;
            
        case 'renumber':
            document.getElementById("acctRenumberNewGroup").hidden = false;
            const renumberInput = document.getElementById("acctRenumberNewNumber");
            if (renumberInput) {
                renumberInput.addEventListener("input", () => {
                    updateRenumberAccountPreview(account, renumberInput.value);
                });
            }
            break;
    }
}

/**
 * Update preview for Rename Account
 */
function updateRenameAccountPreview(account, newName) {
    const preview = document.getElementById("quickEditPreviewAccounts");
    const buttons = document.getElementById("quickEditButtonsAccounts");
    
    if (!newName?.trim()) {
        if (preview) preview.hidden = true;
        if (buttons) buttons.hidden = true;
        return;
    }
    
    editState.pendingChanges = [{
        type: 'rename',
        index: account.index,
        data: { ...account, name: newName.trim() }
    }];
    
    if (preview) {
        preview.hidden = false;
        preview.innerHTML = `
            <h4>Preview</h4>
            <div class="preview-change">
                <span class="preview-icon">✏️</span>
                <div class="preview-text">
                    <div class="preview-field">Account ${escapeHtml(account.number)} will be renamed</div>
                    <div class="preview-values">
                        <span class="preview-old">${escapeHtml(account.name)}</span>
                        <span class="preview-arrow">→</span>
                        <span class="preview-new">${escapeHtml(newName.trim())}</span>
                    </div>
                </div>
            </div>
        `;
    }
    
    if (buttons) buttons.hidden = false;
}

/**
 * Update preview for Renumber Account
 */
function updateRenumberAccountPreview(account, newNumber) {
    const preview = document.getElementById("quickEditPreviewAccounts");
    const buttons = document.getElementById("quickEditButtonsAccounts");
    
    if (!newNumber?.trim()) {
        if (preview) preview.hidden = true;
        if (buttons) buttons.hidden = true;
        return;
    }
    
    editState.pendingChanges = [{
        type: 'renumber',
        index: account.index,
        data: { ...account, number: newNumber.trim() }
    }];
    
    if (preview) {
        preview.hidden = false;
        preview.innerHTML = `
            <h4>Preview</h4>
            <div class="preview-change">
                <span class="preview-icon">#️⃣</span>
                <div class="preview-text">
                    <div class="preview-field">${escapeHtml(account.name)} will be renumbered</div>
                    <div class="preview-values">
                        <span class="preview-old">${escapeHtml(account.number)}</span>
                        <span class="preview-arrow">→</span>
                        <span class="preview-new">${escapeHtml(newNumber.trim())}</span>
                    </div>
                </div>
            </div>
        `;
    }
    
    if (buttons) buttons.hidden = false;
}

/**
 * Reset Quick Edit Accounts form
 */
function resetQuickEditAccounts() {
    const form = document.getElementById("quickEditFormAccounts");
    const preview = document.getElementById("quickEditPreviewAccounts");
    const buttons = document.getElementById("quickEditButtonsAccounts");
    
    if (form) form.innerHTML = '';
    if (preview) preview.hidden = true;
    if (buttons) buttons.hidden = true;
    
    // Reset action buttons
    document.querySelectorAll("#quickEditAccounts .quick-action-btn").forEach(btn => {
        btn.classList.remove("active");
    });
    
    editState.action = null;
    editState.selectedRecord = null;
    editState.pendingChanges = [];
}

/**
 * Confirm and apply quick edit changes for Accounts
 */
async function confirmQuickEditAccounts() {
    if (editState.pendingChanges.length === 0) return;
    
    if (!hasExcelRuntime()) {
        alert("Excel not available");
        return;
    }
    
    const config = REF_DATA_CONFIG.accounts;
    
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(config.sheetName);
            
            for (const change of editState.pendingChanges) {
                if (change.type === 'add') {
                    // Append new row
                    const usedRange = sheet.getUsedRange();
                    usedRange.load("rowCount");
                    await context.sync();
                    
                    const newRow = usedRange.rowCount + 1;
                    const data = change.data;
                    const newRange = sheet.getRange(`A${newRow}:D${newRow}`);
                    newRange.values = [[data.number, data.name, data.type, data.category]];
                    
                } else if (change.type === 'delete') {
                    // Delete row
                    const rowNum = change.index + 2;
                    const deleteRange = sheet.getRange(`A${rowNum}:D${rowNum}`);
                    deleteRange.delete(Excel.DeleteShiftDirection.up);
                    
                } else if (change.type === 'rename' || change.type === 'renumber') {
                    // Update existing row
                    const rowNum = change.index + 2;
                    const data = change.data;
                    const updateRange = sheet.getRange(`A${rowNum}:D${rowNum}`);
                    updateRange.values = [[data.number, data.name, data.type, data.category]];
                }
            }
            
            await context.sync();
        });
        
        // Success - refresh and close edit mode
        alert("Changes saved successfully!");
        showViewMode();
        
    } catch (error) {
        console.error("Error saving changes:", error);
        alert("Error saving changes: " + error.message);
    }
}

// =============================================================================
// BULK UPLOAD
// =============================================================================

/**
 * Initialize bulk upload dropzone
 */
function initBulkUploadDropzone() {
    const dropzone = document.getElementById("bulkUploadDropzone");
    const fileInput = document.getElementById("bulkUploadFile");
    
    if (!dropzone || !fileInput) return;
    
    // Prevent duplicate listeners by checking for a flag
    if (dropzone.dataset.initialized) return;
    dropzone.dataset.initialized = "true";
    
    dropzone.addEventListener("click", (e) => {
        e.stopPropagation();
        fileInput.click();
    });
    
    dropzone.addEventListener("dragover", (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropzone.classList.add("dragover");
    });
    
    dropzone.addEventListener("dragleave", (e) => {
        e.preventDefault();
        dropzone.classList.remove("dragover");
    });
    
    dropzone.addEventListener("drop", (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropzone.classList.remove("dragover");
        if (e.dataTransfer.files.length > 0) {
            handleBulkUploadFile(e.dataTransfer.files[0]);
        }
    });
    
    fileInput.addEventListener("change", () => {
        if (fileInput.files.length > 0) {
            handleBulkUploadFile(fileInput.files[0]);
        }
    });
    
    document.getElementById("bulkUploadFileRemove")?.addEventListener("click", resetBulkUpload);
}

/**
 * Handle bulk upload file
 */
async function handleBulkUploadFile(file) {
    const fileInfo = document.getElementById("bulkUploadFileInfo");
    const fileName = document.getElementById("bulkUploadFileName");
    const fileRows = document.getElementById("bulkUploadFileRows");
    const dropzone = document.getElementById("bulkUploadDropzone");
    const warningEl = document.getElementById("bulkUploadWarning");
    const warningText = document.getElementById("bulkUploadWarningText");
    
    if (!file) return;
    
    const ext = file.name.split('.').pop().toLowerCase();
    
    try {
        let headers = [];
        let data = [];
        let headerRowIndex = 0;
        let headerQuality = 'good';
        
        if (ext === 'csv') {
            const result = await parseBulkCSV(file);
            headers = result.headers;
            data = result.data;
            headerRowIndex = result.headerRowIndex;
            headerQuality = result.headerQuality;
        } else if (ext === 'xlsx' || ext === 'xls') {
            // For now, show a message about CSV being preferred
            if (warningEl && warningText) {
                warningEl.hidden = false;
                warningText.textContent = "Excel files work, but CSV provides more reliable results. Consider saving as CSV if you encounter issues.";
            }
            return;
        } else {
            if (warningEl && warningText) {
                warningEl.hidden = false;
                warningText.textContent = "Please upload a CSV or Excel file.";
            }
            return;
        }
        
        editState.uploadedHeaders = headers;
        editState.uploadedData = data;
        
        // Show file info
        if (dropzone) dropzone.style.display = 'none';
        if (fileInfo) fileInfo.hidden = false;
        if (fileName) fileName.textContent = file.name;
        if (fileRows) fileRows.textContent = `${data.length} rows`;
        
        // Show warnings based on header detection
        if (warningEl && warningText) {
            if (headerRowIndex > 0) {
                warningEl.hidden = false;
                warningText.textContent = `Headers detected in row ${headerRowIndex + 1} (skipped ${headerRowIndex} title row${headerRowIndex > 1 ? 's' : ''}). Verify the column mapping below is correct.`;
            } else if (headerQuality === 'poor') {
                warningEl.hidden = false;
                warningText.textContent = "⚠️ Headers may not be recognized correctly. For best results, ensure column headers are in row 1 with clear names like 'Account Number', 'Name', etc.";
            } else {
                warningEl.hidden = true;
            }
        }
        
        // Generate comparison
        await generateBulkComparison();
        
    } catch (error) {
        console.error("Error parsing file:", error);
        if (warningEl && warningText) {
            warningEl.hidden = false;
            warningText.textContent = "Error reading file: " + error.message;
        }
    }
}

/**
 * Parse CSV file for bulk upload
 */
function parseBulkCSV(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const text = e.target.result;
                const lines = text.split(/\r?\n/).filter(line => line.trim());
                
                if (lines.length === 0) {
                    reject(new Error("File is empty"));
                    return;
                }
                
                // Parse rows
                const parseRow = (row) => {
                    const result = [];
                    let current = '';
                    let inQuotes = false;
                    
                    for (let i = 0; i < row.length; i++) {
                        const char = row[i];
                        if (char === '"') {
                            inQuotes = !inQuotes;
                        } else if (char === ',' && !inQuotes) {
                            result.push(current.trim());
                            current = '';
                        } else {
                            current += char;
                        }
                    }
                    result.push(current.trim());
                    return result;
                };
                
                // Parse all rows first
                const allRows = lines.map(line => parseRow(line));
                
                // Try to detect the actual header row
                // Look for a row that contains likely header keywords
                const headerKeywords = [
                    'account', 'number', 'name', 'type', 'description', 'category',
                    'employee', 'department', 'status', 'date', 'rate', 'id',
                    'code', 'amount', 'debit', 'credit', 'balance'
                ];
                
                let headerRowIndex = 0;
                let bestScore = 0;
                
                // Check first 10 rows for potential headers
                const rowsToCheck = Math.min(10, allRows.length);
                for (let i = 0; i < rowsToCheck; i++) {
                    const row = allRows[i];
                    let score = 0;
                    
                    for (const cell of row) {
                        const cellLower = String(cell || '').toLowerCase();
                        // Check if cell contains any header keyword
                        for (const keyword of headerKeywords) {
                            if (cellLower.includes(keyword)) {
                                score++;
                                break;
                            }
                        }
                        // Bonus for cells that look like column headers (short, no numbers only)
                        if (cellLower.length > 0 && cellLower.length < 30 && !/^\d+$/.test(cellLower)) {
                            score += 0.5;
                        }
                    }
                    
                    if (score > bestScore) {
                        bestScore = score;
                        headerRowIndex = i;
                    }
                }
                
                const headers = allRows[headerRowIndex];
                const data = allRows.slice(headerRowIndex + 1);
                
                // Determine if we should warn about header quality
                const headerQuality = bestScore >= 2 ? 'good' : bestScore >= 1 ? 'fair' : 'poor';
                const headerRowFound = headerRowIndex;
                
                resolve({ 
                    headers, 
                    data,
                    headerRowIndex,
                    headerQuality,
                    originalRowCount: allRows.length
                });
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = () => reject(new Error("Failed to read file"));
        reader.readAsText(file);
    });
}

/**
 * Generate comparison between uploaded data and current data
 */
async function generateBulkComparison() {
    const comparison = document.getElementById("bulkUploadComparison");
    const results = document.getElementById("comparisonResults");
    const warning = document.getElementById("bulkUploadWarning");
    const warningText = document.getElementById("bulkUploadWarningText");
    const buttons = document.getElementById("bulkUploadButtons");
    
    if (!comparison || !results) return;
    
    const config = REF_DATA_CONFIG[refDataState.type];
    const targetHeaders = config?.defaultHeaders || [];
    const currentData = refDataState.data;
    const uploadedData = editState.uploadedData;
    const uploadedHeaders = editState.uploadedHeaders;
    
    // Build column mapping to show user what was detected
    const columnMapping = buildColumnMapping(uploadedHeaders, targetHeaders);
    const mappedColumns = Object.keys(columnMapping).length;
    const unmappedColumns = uploadedHeaders.filter((h, i) => columnMapping[i] === undefined);
    
    // Determine the key column (first mapped column, usually Employee or Account_Number)
    const keySourceIdx = Object.keys(columnMapping).find(i => columnMapping[i] === 0);
    
    let currentKeys = new Set();
    let uploadedKeys = new Set();
    
    if (keySourceIdx !== undefined) {
        currentKeys = new Set(currentData.map(row => String(row[0] || '').toLowerCase().trim()));
        uploadedKeys = new Set(uploadedData.map(row => String(row[keySourceIdx] || '').toLowerCase().trim()));
    }
    
    // Find new, removed records
    const newRecords = keySourceIdx !== undefined 
        ? uploadedData.filter(row => !currentKeys.has(String(row[keySourceIdx] || '').toLowerCase().trim()))
        : uploadedData;
    const removedRecords = keySourceIdx !== undefined
        ? currentData.filter(row => !uploadedKeys.has(String(row[0] || '').toLowerCase().trim()))
        : [];
    
    // Render comparison results
    let html = '';
    
    // Filter out empty headers for display
    const nonEmptyHeaders = uploadedHeaders.filter(h => h && String(h).trim());
    const mappedNonEmpty = nonEmptyHeaders.filter((h, i) => {
        const origIdx = uploadedHeaders.indexOf(h);
        return columnMapping[origIdx] !== undefined;
    });
    
    // Show column mapping info
    html += `
        <div class="comparison-section">
            <div class="comparison-header">
                <span class="icon-add">📋</span>
                Column Mapping
            </div>
            <div class="comparison-items">
                <div class="comparison-item">✓ ${mappedNonEmpty.length} of ${nonEmptyHeaders.length} columns matched</div>
                ${uploadedHeaders.map((h, i) => {
                    const targetIdx = columnMapping[i];
                    const sourceLabel = h && String(h).trim() ? escapeHtml(h) : `(Column ${i + 1})`;
                    if (targetIdx !== undefined) {
                        return `<div class="comparison-item" style="color: #10b981;">• ${sourceLabel} → ${escapeHtml(targetHeaders[targetIdx])}</div>`;
                    }
                    return '';
                }).join('')}
                ${unmappedColumns.filter(h => h && String(h).trim()).length > 0 
                    ? `<div class="comparison-item" style="color: #f59e0b;">⚠ Unmapped: ${unmappedColumns.filter(h => h && String(h).trim()).map(h => escapeHtml(h)).join(', ')}</div>` 
                    : ''}
            </div>
        </div>
    `;
    
    if (newRecords.length > 0 && keySourceIdx !== undefined) {
        html += `
            <div class="comparison-section">
                <div class="comparison-header">
                    <span class="icon-add">➕</span>
                    ${newRecords.length} New Record${newRecords.length > 1 ? 's' : ''}
                </div>
                <div class="comparison-items">
                    ${newRecords.slice(0, 5).map(r => `<div class="comparison-item">• ${escapeHtml(r[keySourceIdx])}</div>`).join('')}
                    ${newRecords.length > 5 ? `<div class="comparison-item">• ... and ${newRecords.length - 5} more</div>` : ''}
                </div>
            </div>
        `;
    }
    
    if (removedRecords.length > 0) {
        html += `
            <div class="comparison-section">
                <div class="comparison-header">
                    <span class="icon-remove">➖</span>
                    ${removedRecords.length} Will Be Removed
                </div>
                <div class="comparison-items">
                    ${removedRecords.slice(0, 5).map(r => `<div class="comparison-item">• ${escapeHtml(r[0])}</div>`).join('')}
                    ${removedRecords.length > 5 ? `<div class="comparison-item">• ... and ${removedRecords.length - 5} more</div>` : ''}
                </div>
            </div>
        `;
    }
    
    results.innerHTML = html;
    comparison.hidden = false;
    
    // Show warning
    if (warning && warningText) {
        warning.hidden = false;
        warningText.textContent = `This will replace all ${currentData.length} current records with ${uploadedData.length} records from the uploaded file.`;
    }
    
    if (buttons) buttons.hidden = false;
}

/**
 * Reset bulk upload panel
 */
function resetBulkUpload() {
    const fileInfo = document.getElementById("bulkUploadFileInfo");
    const dropzone = document.getElementById("bulkUploadDropzone");
    const comparison = document.getElementById("bulkUploadComparison");
    const warning = document.getElementById("bulkUploadWarning");
    const buttons = document.getElementById("bulkUploadButtons");
    const fileInput = document.getElementById("bulkUploadFile");
    
    if (fileInfo) fileInfo.hidden = true;
    if (dropzone) dropzone.style.display = '';
    if (comparison) comparison.hidden = true;
    if (warning) warning.hidden = true;
    if (buttons) buttons.hidden = true;
    if (fileInput) fileInput.value = '';
    
    editState.uploadedHeaders = [];
    editState.uploadedData = [];
}

/**
 * Confirm bulk upload - replace all data
 */
async function confirmBulkUpload() {
    console.log("confirmBulkUpload called");
    console.log("uploadedData length:", editState.uploadedData.length);
    console.log("refDataState.type:", refDataState.type);
    
    const confirmBtn = document.getElementById("bulkUploadConfirm");
    const warningEl = document.getElementById("bulkUploadWarning");
    const warningText = document.getElementById("bulkUploadWarningText");
    
    if (editState.uploadedData.length === 0) {
        showBulkUploadStatus("No data to upload", "error");
        return;
    }
    
    if (!hasExcelRuntime()) {
        showBulkUploadStatus("Excel not available", "error");
        return;
    }
    
    const config = REF_DATA_CONFIG[refDataState.type];
    console.log("Config:", config);
    if (!config) {
        showBulkUploadStatus("Configuration error", "error");
        return;
    }
    
    // Show loading state
    if (confirmBtn) {
        confirmBtn.disabled = true;
        confirmBtn.innerHTML = `<span class="loading-spinner"></span> Uploading...`;
    }
    
    console.log("User confirmed, starting Excel operation...");
    
    try {
        await Excel.run(async (context) => {
            console.log("Inside Excel.run");
            
            // Get or create the sheet
            let sheet = context.workbook.worksheets.getItemOrNullObject(config.sheetName);
            sheet.load("isNullObject");
            await context.sync();
            console.log("Sheet isNullObject:", sheet.isNullObject);
            
            if (sheet.isNullObject) {
                console.log("Creating new sheet:", config.sheetName);
                // Create the sheet with headers
                sheet = context.workbook.worksheets.add(config.sheetName);
                const headerRange = sheet.getRange(`A1:${String.fromCharCode(64 + config.defaultHeaders.length)}1`);
                headerRange.values = [config.defaultHeaders];
                headerRange.format.font.bold = true;
                headerRange.format.fill.color = "#f0f0f0";
                await context.sync();
                console.log("Sheet created with headers");
            }
            
            // Get the target headers from config
            const targetHeaders = config.defaultHeaders;
            const numTargetCols = targetHeaders.length;
            
            // Build column mapping: source column index -> target column index
            const columnMapping = buildColumnMapping(editState.uploadedHeaders, targetHeaders);
            
            console.log("Column mapping:", columnMapping);
            console.log("Source headers:", editState.uploadedHeaders);
            console.log("Target headers:", targetHeaders);
            
            // Clear existing data (keep headers)
            console.log("Clearing existing data...");
            const usedRange = sheet.getUsedRangeOrNullObject();
            usedRange.load("rowCount");
            await context.sync();
            
            if (!usedRange.isNullObject && usedRange.rowCount > 1) {
                console.log("Clearing rows 2 to", usedRange.rowCount);
                const dataRange = sheet.getRange(`A2:Z${usedRange.rowCount}`);
                dataRange.clear();
                await context.sync();
            }
            
            // Write new data with column mapping
            console.log("Writing", editState.uploadedData.length, "rows...");
            if (editState.uploadedData.length > 0) {
                const endCol = String.fromCharCode(64 + numTargetCols);
                const endRow = editState.uploadedData.length + 1;
                const rangeAddress = `A2:${endCol}${endRow}`;
                console.log("Target range:", rangeAddress);
                
                const dataRange = sheet.getRange(rangeAddress);
                
                // Map each row to target columns
                const mappedData = editState.uploadedData.map(sourceRow => {
                    const targetRow = new Array(numTargetCols).fill('');
                    
                    // Map each source column to target
                    for (let srcIdx = 0; srcIdx < sourceRow.length; srcIdx++) {
                        const targetIdx = columnMapping[srcIdx];
                        if (targetIdx !== undefined && targetIdx >= 0) {
                            targetRow[targetIdx] = sourceRow[srcIdx];
                        }
                    }
                    
                    return targetRow;
                });
                
                console.log("Mapped data rows:", mappedData.length);
                console.log("Mapped data sample:", mappedData.slice(0, 3));
                
                dataRange.values = mappedData;
                console.log("Calling final context.sync()...");
                await context.sync();
                console.log("Data written successfully!");
            }
        });
        
        console.log("Excel.run completed successfully");
        
        const rowsUploaded = editState.uploadedData.length;
        
        // Reset states
        resetBulkUploadButton();
        resetBulkUpload();
        
        // Go back to view mode and refresh data
        showViewMode();
        await loadRefData();
        
        // Show success message in the view mode
        showRefDataSuccessMessage(`✓ ${rowsUploaded} records uploaded successfully`);
        
    } catch (error) {
        console.error("Error uploading data:", error);
        console.error("Error stack:", error.stack);
        showBulkUploadStatus("Error: " + error.message, "error");
        resetBulkUploadButton();
    }
}

/**
 * Show a temporary success message in the ref data view mode
 */
function showRefDataSuccessMessage(message) {
    // Find or create success message element
    let successEl = document.getElementById("refDataSuccess");
    const body = document.querySelector(".ref-data-body");
    
    if (!successEl && body) {
        successEl = document.createElement("div");
        successEl.id = "refDataSuccess";
        successEl.style.cssText = `
            padding: 12px 16px;
            background: rgba(16, 185, 129, 0.15);
            border: 1px solid #10b981;
            border-radius: 10px;
            color: #10b981;
            font-size: 13px;
            font-weight: 500;
            margin-bottom: 16px;
            text-align: center;
            animation: fadeIn 0.3s ease;
        `;
        body.insertBefore(successEl, body.firstChild);
    }
    
    if (successEl) {
        successEl.textContent = message;
        successEl.hidden = false;
        
        // Auto-hide after 4 seconds
        setTimeout(() => {
            successEl.style.opacity = '0';
            successEl.style.transition = 'opacity 0.3s ease';
            setTimeout(() => {
                successEl.hidden = true;
                successEl.style.opacity = '1';
            }, 300);
        }, 4000);
    }
}

/**
 * Show status message in bulk upload panel
 */
function showBulkUploadStatus(message, type = "info") {
    const warningEl = document.getElementById("bulkUploadWarning");
    const warningText = document.getElementById("bulkUploadWarningText");
    
    if (warningEl && warningText) {
        warningEl.hidden = false;
        warningText.textContent = message;
        
        // Update styling based on type
        warningEl.className = 'bulk-upload-warning';
        if (type === 'success') {
            warningEl.style.background = 'rgba(16, 185, 129, 0.15)';
            warningEl.style.borderColor = '#10b981';
            warningText.style.color = '#10b981';
        } else if (type === 'error') {
            warningEl.style.background = 'rgba(239, 68, 68, 0.15)';
            warningEl.style.borderColor = '#ef4444';
            warningText.style.color = '#ef4444';
        } else {
            warningEl.style.background = '';
            warningEl.style.borderColor = '';
            warningText.style.color = '';
        }
    }
}

/**
 * Reset the bulk upload confirm button to default state
 */
function resetBulkUploadButton() {
    const confirmBtn = document.getElementById("bulkUploadConfirm");
    if (confirmBtn) {
        confirmBtn.disabled = false;
        confirmBtn.innerHTML = 'Confirm & Replace All';
    }
}

/**
 * Build column mapping from source headers to target headers
 * Uses fuzzy matching for flexibility
 */
function buildColumnMapping(sourceHeaders, targetHeaders) {
    const mapping = {}; // sourceIndex -> targetIndex
    
    // Normalize header for comparison - strip special chars, spaces, underscores
    const normalize = (str) => String(str || '').toLowerCase()
        .replace(/[_\s\-#@$%&*()]/g, '')  // Remove special chars including #
        .trim();
    
    // Map of source patterns to target field names
    // Keys are normalized source header patterns
    // Values are normalized target field names
    const aliasGroups = {
        // Account Number variations
        'accountnumber': ['accountnumber', 'accountnum', 'acctnum', 'acctno', 'account', 'acct', 'accountno', 'glcode', 'glaccount'],
        // Account Name variations  
        'accountname': ['accountname', 'acctname', 'name', 'fullname', 'glname', 'accountdescription'],
        // Type variations
        'type': ['type', 'accounttype', 'accttype', 'gltype'],
        // Category variations
        'category': ['category', 'detailtype', 'subcategory', 'subtype', 'classification', 'class'],
        // Employee variations
        'employee': ['employee', 'employeename', 'name', 'fullname', 'empname', 'worker', 'staffname', 'associatename'],
        // Department variations
        'department': ['department', 'dept', 'division', 'team', 'unit', 'group', 'costcenter'],
        // Pay Rate variations
        'payrate': ['payrate', 'rate', 'salary', 'hourlyrate', 'wage', 'compensation', 'pay'],
        // Status variations
        'status': ['status', 'active', 'employeestatus', 'empstatus', 'state'],
        // Hire Date variations
        'hiredate': ['hiredate', 'startdate', 'dateofhire', 'joindate', 'begindate', 'employmentdate'],
        // Term Date variations
        'termdate': ['termdate', 'terminationdate', 'enddate', 'leavedate', 'exitdate', 'separationdate'],
        // Description variations
        'description': ['description', 'desc', 'memo', 'notes', 'comment', 'details']
    };
    
    // Build reverse lookup: normalized source header -> target field
    const aliasToTarget = {};
    for (const [targetField, aliases] of Object.entries(aliasGroups)) {
        for (const alias of aliases) {
            aliasToTarget[alias] = targetField;
        }
    }
    
    sourceHeaders.forEach((sourceHeader, srcIdx) => {
        const srcNorm = normalize(sourceHeader);
        if (!srcNorm) return; // Skip empty headers
        
        // Try exact match first
        let targetIdx = targetHeaders.findIndex(t => normalize(t) === srcNorm);
        
        // Try alias lookup
        if (targetIdx === -1) {
            const targetField = aliasToTarget[srcNorm];
            if (targetField) {
                targetIdx = targetHeaders.findIndex(t => normalize(t) === targetField);
            }
        }
        
        // Try partial match (source contains target or vice versa)
        if (targetIdx === -1) {
            targetIdx = targetHeaders.findIndex(t => {
                const tNorm = normalize(t);
                return srcNorm.includes(tNorm) || tNorm.includes(srcNorm);
            });
        }
        
        // Try finding a target that shares an alias group with the source
        if (targetIdx === -1) {
            for (const [targetField, aliases] of Object.entries(aliasGroups)) {
                if (aliases.some(a => srcNorm.includes(a) || a.includes(srcNorm))) {
                    targetIdx = targetHeaders.findIndex(t => normalize(t) === targetField);
                    if (targetIdx !== -1) break;
                }
            }
        }
        
        if (targetIdx !== -1 && !Object.values(mapping).includes(targetIdx)) {
            // Only map if target column isn't already mapped
            mapping[srcIdx] = targetIdx;
        }
    });
    
    return mapping;
}

// =============================================================================
// COLUMN MAPPING
// =============================================================================

// Mapping state
const mappingState = {
    targetTab: null,
    sourceHeaders: [],
    sampleData: [],
    mappings: {} // sourceColumn -> targetField
};

/**
 * Initialize Column Mapping UI
 */
function initColumnMapping() {
    // Sub-tab switching
    document.querySelectorAll(".create-subtab").forEach(tab => {
        tab.addEventListener("click", () => {
            const subtabId = tab.dataset.subtab;
            
            // Update tab buttons
            document.querySelectorAll(".create-subtab").forEach(t => 
                t.classList.toggle("active", t.dataset.subtab === subtabId)
            );
            
            // Update content
            document.querySelectorAll(".create-subtab-content").forEach(content => 
                content.classList.toggle("active", content.id === `subtab-${subtabId}`)
            );
        });
    });
    
    // Target tab dropdown
    const targetSelect = document.getElementById("mappingTargetTab");
    if (targetSelect) {
        populateMappingTabDropdown();
        targetSelect.addEventListener("change", onMappingTabChange);
    }
    
    // File upload
    const dropzone = document.getElementById("uploadDropzone");
    const fileInput = document.getElementById("uploadFileInput");
    
    if (dropzone && fileInput) {
        dropzone.addEventListener("click", () => fileInput.click());
        dropzone.addEventListener("dragover", (e) => {
            e.preventDefault();
            dropzone.classList.add("dragover");
        });
        dropzone.addEventListener("dragleave", () => {
            dropzone.classList.remove("dragover");
        });
        dropzone.addEventListener("drop", (e) => {
            e.preventDefault();
            dropzone.classList.remove("dragover");
            if (e.dataTransfer.files.length > 0) {
                handleFileUpload(e.dataTransfer.files[0]);
            }
        });
        fileInput.addEventListener("change", () => {
            if (fileInput.files.length > 0) {
                handleFileUpload(fileInput.files[0]);
            }
        });
    }
    
    // Remove file button
    document.getElementById("mappingFileRemove")?.addEventListener("click", clearMappingFile);
    
    // Save/Clear buttons
    document.getElementById("saveMappingBtn")?.addEventListener("click", saveColumnMapping);
    document.getElementById("clearMappingBtn")?.addEventListener("click", clearAllMappings);
}

/**
 * Populate the target tab dropdown with tabs that have templates
 */
function populateMappingTabDropdown() {
    const select = document.getElementById("mappingTargetTab");
    if (!select) return;
    
    // Get tabs from tab mappings that might need mapping
    const dataTabs = tabMappings.filter(t => 
        t.tabName.includes("_Data") || 
        t.tabName.includes("_Import") ||
        t.tabName.includes("_JE")
    );
    
    select.innerHTML = '<option value="">Select a tab...</option>' +
        dataTabs.map(t => 
            `<option value="${escapeHtml(t.tabName)}">${escapeHtml(t.tabName)} (${t.module})</option>`
        ).join("");
}

/**
 * Handle target tab selection change
 */
function onMappingTabChange() {
    const select = document.getElementById("mappingTargetTab");
    mappingState.targetTab = select?.value || null;
    
    // Load existing mapping if available
    loadExistingMapping();
    renderMappingGrid();
}

/**
 * Handle file upload
 */
async function handleFileUpload(file) {
    const fileInfo = document.getElementById("mappingFileInfo");
    const fileName = document.getElementById("mappingFileName");
    const fileCols = document.getElementById("mappingFileCols");
    const hint = document.getElementById("mappingHint");
    
    if (!file) return;
    
    const ext = file.name.split('.').pop().toLowerCase();
    
    try {
        let headers = [];
        let sampleRow = [];
        
        if (ext === 'csv') {
            const result = await parseCSV(file);
            headers = result.headers;
            sampleRow = result.sampleRow;
        } else if (ext === 'xlsx' || ext === 'xls') {
            const result = await parseExcel(file);
            headers = result.headers;
            sampleRow = result.sampleRow;
        } else {
            alert("Please upload a CSV or Excel file.");
            return;
        }
        
        mappingState.sourceHeaders = headers;
        mappingState.sampleData = sampleRow;
        
        // Show file info
        if (fileInfo) fileInfo.hidden = false;
        if (fileName) fileName.textContent = file.name;
        if (fileCols) fileCols.textContent = `${headers.length} columns`;
        if (hint) hint.textContent = `Detected ${headers.length} columns. Map each to a target field or ignore.`;
        
        // Initialize mappings with auto-match
        autoMatchColumns();
        renderMappingGrid();
        
    } catch (error) {
        console.error("Error parsing file:", error);
        alert("Error reading file: " + error.message);
    }
}

/**
 * Parse CSV file
 */
function parseCSV(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const text = e.target.result;
                const lines = text.split(/\r?\n/).filter(line => line.trim());
                
                if (lines.length === 0) {
                    reject(new Error("File is empty"));
                    return;
                }
                
                // Simple CSV parsing (handles basic cases)
                const parseRow = (row) => {
                    const result = [];
                    let current = '';
                    let inQuotes = false;
                    
                    for (let i = 0; i < row.length; i++) {
                        const char = row[i];
                        if (char === '"') {
                            inQuotes = !inQuotes;
                        } else if (char === ',' && !inQuotes) {
                            result.push(current.trim());
                            current = '';
                        } else {
                            current += char;
                        }
                    }
                    result.push(current.trim());
                    return result;
                };
                
                const headers = parseRow(lines[0]);
                const sampleRow = lines.length > 1 ? parseRow(lines[1]) : [];
                
                resolve({ headers, sampleRow });
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = () => reject(new Error("Failed to read file"));
        reader.readAsText(file);
    });
}

/**
 * Parse Excel file (requires reading from Excel if available, or basic parsing)
 */
async function parseExcel(file) {
    // For Excel files, we'll try to use the browser's FileReader
    // Note: Full Excel parsing would require a library like SheetJS
    // For now, we'll show a message to use CSV
    
    return new Promise((resolve, reject) => {
        // Basic XLSX parsing attempt
        const reader = new FileReader();
        reader.onload = async (e) => {
            try {
                // Try to extract headers from Excel
                // This is a simplified approach - recommend using CSV for now
                const arrayBuffer = e.target.result;
                
                // Look for sheet data in the file
                // This is a very basic check - real XLSX parsing needs a library
                const uint8Array = new Uint8Array(arrayBuffer);
                const text = new TextDecoder().decode(uint8Array);
                
                // If we can't parse it, suggest CSV
                if (!text.includes('worksheet')) {
                    // Try to find any recognizable text headers
                    alert("For best results, please export your data as CSV. Excel file parsing is limited.");
                    resolve({ headers: [], sampleRow: [] });
                    return;
                }
                
                // Fallback: suggest CSV
                alert("Excel parsing is limited. For best results, please save as CSV and re-upload.");
                resolve({ headers: [], sampleRow: [] });
                
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = () => reject(new Error("Failed to read file"));
        reader.readAsArrayBuffer(file);
    });
}

/**
 * Auto-match columns based on name similarity
 */
function autoMatchColumns() {
    if (!mappingState.targetTab) return;
    
    // Get target fields from tab template
    const tabMapping = tabMappings.find(t => t.tabName === mappingState.targetTab);
    const templateId = getTemplateIdForTab(mappingState.targetTab);
    const template = TAB_TEMPLATES[templateId];
    const targetFields = template?.headers || [];
    
    mappingState.mappings = {};
    
    // Try to auto-match based on name similarity
    mappingState.sourceHeaders.forEach(sourceCol => {
        const sourceLower = sourceCol.toLowerCase().replace(/[_\s-]/g, '');
        
        for (const targetField of targetFields) {
            const targetLower = targetField.toLowerCase().replace(/[_\s-]/g, '');
            
            // Exact match or partial match
            if (sourceLower === targetLower || 
                sourceLower.includes(targetLower) || 
                targetLower.includes(sourceLower)) {
                mappingState.mappings[sourceCol] = targetField;
                break;
            }
        }
    });
}

/**
 * Get template ID for a tab name
 */
function getTemplateIdForTab(tabName) {
    if (tabName.includes("_Data_Clean")) return "data-clean";
    if (tabName.includes("_Data")) return "data-import";
    if (tabName.includes("_JE")) return "journal";
    if (tabName.includes("_Analysis")) return "analysis";
    if (tabName.includes("_Roster")) return "roster";
    if (tabName.includes("_Expense")) return "expense";
    if (tabName.includes("_Archive")) return "archive";
    return "data-import"; // Default
}

/**
 * Render the mapping grid
 */
function renderMappingGrid() {
    const grid = document.getElementById("mappingGrid");
    if (!grid) return;
    
    if (mappingState.sourceHeaders.length === 0) {
        grid.innerHTML = "";
        return;
    }
    
    // Get target fields
    const templateId = getTemplateIdForTab(mappingState.targetTab || "");
    const template = TAB_TEMPLATES[templateId];
    const targetFields = template?.headers || ["Column1", "Column2", "Column3"];
    
    grid.innerHTML = mappingState.sourceHeaders.map((sourceCol, idx) => {
        const mappedTo = mappingState.mappings[sourceCol] || "";
        const sampleValue = mappingState.sampleData[idx] || "";
        const isMapped = mappedTo && mappedTo !== "__ignore__";
        const isIgnored = mappedTo === "__ignore__";
        
        return `
            <div class="mapping-row-item ${isMapped ? 'mapped' : ''} ${isIgnored ? 'ignored' : ''}" data-source="${escapeHtml(sourceCol)}">
                <div class="mapping-source">
                    <span class="mapping-source-name">${escapeHtml(sourceCol)}</span>
                    ${sampleValue ? `<span class="mapping-source-sample">e.g., "${escapeHtml(String(sampleValue).substring(0, 30))}"</span>` : ''}
                </div>
                <div class="mapping-arrow">→</div>
                <select class="mapping-target-select" data-source="${escapeHtml(sourceCol)}">
                    <option value="">— Select target —</option>
                    <option value="__ignore__" ${isIgnored ? 'selected' : ''}>🚫 Ignore</option>
                    ${targetFields.map(f => 
                        `<option value="${escapeHtml(f)}" ${mappedTo === f ? 'selected' : ''}>${escapeHtml(f)}</option>`
                    ).join("")}
                </select>
            </div>
        `;
    }).join("");
    
    // Wire up change handlers
    grid.querySelectorAll(".mapping-target-select").forEach(select => {
        select.addEventListener("change", () => {
            const sourceCol = select.dataset.source;
            const targetField = select.value;
            
            if (targetField) {
                mappingState.mappings[sourceCol] = targetField;
            } else {
                delete mappingState.mappings[sourceCol];
            }
            
            renderMappingGrid();
            updateMappingSummary();
        });
    });
    
    updateMappingSummary();
}

/**
 * Update the mapping summary
 */
function updateMappingSummary() {
    const summary = document.getElementById("mappingSummary");
    const saveBtn = document.getElementById("saveMappingBtn");
    
    if (!summary) return;
    
    const mappedEntries = Object.entries(mappingState.mappings)
        .filter(([source, target]) => target && target !== "__ignore__");
    
    if (mappedEntries.length === 0) {
        summary.innerHTML = '<p class="mapping-hint">Configure mappings above to see summary.</p>';
        if (saveBtn) saveBtn.disabled = true;
        return;
    }
    
    summary.innerHTML = `
        <p style="margin: 0 0 12px 0; font-weight: 600; color: var(--pf-text-primary);">
            ${mappedEntries.length} column${mappedEntries.length > 1 ? 's' : ''} mapped:
        </p>
        <div class="mapping-summary-grid">
            ${mappedEntries.map(([source, target]) => `
                <div class="mapping-summary-item">
                    <span class="mapping-summary-source">${escapeHtml(source)}</span>
                    <span class="mapping-summary-arrow">→</span>
                    <span class="mapping-summary-target">${escapeHtml(target)}</span>
                </div>
            `).join("")}
        </div>
    `;
    
    if (saveBtn) saveBtn.disabled = !mappingState.targetTab;
}

/**
 * Clear the uploaded file
 */
function clearMappingFile() {
    const fileInfo = document.getElementById("mappingFileInfo");
    const fileInput = document.getElementById("uploadFileInput");
    const hint = document.getElementById("mappingHint");
    
    mappingState.sourceHeaders = [];
    mappingState.sampleData = [];
    mappingState.mappings = {};
    
    if (fileInfo) fileInfo.hidden = true;
    if (fileInput) fileInput.value = "";
    if (hint) hint.textContent = "Upload a file to see detected columns.";
    
    renderMappingGrid();
    updateMappingSummary();
}

/**
 * Clear all mappings
 */
function clearAllMappings() {
    mappingState.mappings = {};
    renderMappingGrid();
    updateMappingSummary();
}

/**
 * Load existing mapping from SS_PF_Config
 */
async function loadExistingMapping() {
    if (!mappingState.targetTab || !hasExcelRuntime()) return;
    
    // Look for existing mapping rows in SS_PF_Config
    // Format: Category="Column Mapping", Field="{tab}|{source}", Value="{target}"
    try {
        await Excel.run(async (context) => {
            const configSheet = context.workbook.worksheets.getItemOrNullObject("SS_PF_Config");
            configSheet.load("isNullObject");
            await context.sync();
            
            if (configSheet.isNullObject) return;
            
            const usedRange = configSheet.getUsedRangeOrNullObject();
            usedRange.load("values");
            await context.sync();
            
            if (usedRange.isNullObject) return;
            
            const rows = usedRange.values;
            const prefix = `${mappingState.targetTab}|`;
            
            for (let i = 1; i < rows.length; i++) {
                const category = String(rows[i][0] || "").trim();
                const field = String(rows[i][1] || "").trim();
                const value = String(rows[i][2] || "").trim();
                
                if (category === "Column Mapping" && field.startsWith(prefix)) {
                    const sourceCol = field.substring(prefix.length);
                    if (value) {
                        mappingState.mappings[sourceCol] = value;
                    }
                }
            }
        });
    } catch (error) {
        console.error("Error loading existing mapping:", error);
    }
}

/**
 * Save column mapping to SS_PF_Config
 */
async function saveColumnMapping() {
    if (!mappingState.targetTab || !hasExcelRuntime()) {
        alert("Please select a target tab first.");
        return;
    }
    
    const mappedEntries = Object.entries(mappingState.mappings)
        .filter(([source, target]) => target);
    
    if (mappedEntries.length === 0) {
        alert("No mappings to save.");
        return;
    }
    
    try {
        await Excel.run(async (context) => {
            const configSheet = context.workbook.worksheets.getItemOrNullObject("SS_PF_Config");
            configSheet.load("isNullObject");
            await context.sync();
            
            if (configSheet.isNullObject) {
                alert("SS_PF_Config sheet not found.");
                return;
            }
            
            // Get current data to find/update existing mappings
            const usedRange = configSheet.getUsedRangeOrNullObject();
            usedRange.load("values, rowCount");
            await context.sync();
            
            const rows = usedRange.isNullObject ? [] : usedRange.values;
            const prefix = `${mappingState.targetTab}|`;
            
            // Find existing mapping rows for this tab
            const existingRows = new Map(); // field -> rowIndex
            for (let i = 1; i < rows.length; i++) {
                const category = String(rows[i][0] || "").trim();
                const field = String(rows[i][1] || "").trim();
                
                if (category === "Column Mapping" && field.startsWith(prefix)) {
                    existingRows.set(field, i + 1); // 1-based row number
                }
            }
            
            let nextRow = rows.length + 1;
            
            // Update or add each mapping
            for (const [source, target] of mappedEntries) {
                const field = `${prefix}${source}`;
                
                if (existingRows.has(field)) {
                    // Update existing row
                    const rowNum = existingRows.get(field);
                    configSheet.getRange(`C${rowNum}`).values = [[target]];
                } else {
                    // Add new row
                    configSheet.getRange(`A${nextRow}:C${nextRow}`).values = [
                        ["Column Mapping", field, target]
                    ];
                    nextRow++;
                }
            }
            
            await context.sync();
        });
        
        logAdmin(`Saved ${mappedEntries.length} column mappings for ${mappingState.targetTab}`, "success");
        alert("Column mappings saved successfully!");
        
    } catch (error) {
        console.error("Error saving mapping:", error);
        logAdmin(`Error saving mappings: ${error.message}`, "error");
        alert("Error saving mappings: " + error.message);
    }
}

// =============================================================================
// TAB MANAGEMENT
// =============================================================================

/**
 * Load tab mappings from SS_PF_Config (rows with Category = "Tab Structure")
 */
async function loadTabMappings() {
    tabMappings = [];
    
    if (!hasExcelRuntime()) {
        console.warn("Excel not available - cannot load tab mappings");
        return;
    }
    
    try {
        await Excel.run(async (context) => {
            const configSheet = context.workbook.worksheets.getItemOrNullObject("SS_PF_Config");
            configSheet.load("isNullObject");
            await context.sync();
            
            if (configSheet.isNullObject) {
                console.warn("SS_PF_Config sheet not found");
                return;
            }
            
            const usedRange = configSheet.getUsedRangeOrNullObject();
            usedRange.load("values");
            await context.sync();
            
            if (usedRange.isNullObject || !usedRange.values) return;
            
            const rows = usedRange.values;
            // Skip header row, find "Tab Structure" rows
            for (let i = 1; i < rows.length; i++) {
                const category = String(rows[i][0] || "").trim();
                if (category === "Tab Structure") {
                    tabMappings.push({
                        rowIndex: i + 1, // 1-based row number
                        module: String(rows[i][1] || "").trim(),
                        tabName: String(rows[i][2] || "").trim(),
                        folder: String(rows[i][3] || "").trim()
                    });
                }
            }
            
            console.log("Loaded tab mappings:", tabMappings.length);
        });
    } catch (error) {
        console.error("Error loading tab mappings:", error);
    }
}

/**
 * Render tab mappings list in the admin panel
 */
function renderTabMappings() {
    const container = document.getElementById("tabMgmtList");
    const filterSelect = document.getElementById("tabModuleFilter");
    const newTabModuleSelect = document.getElementById("newTabModule");
    
    if (!container) return;
    
    // Get unique modules from tab mappings
    const modules = [...new Set(tabMappings.map(t => t.module))].sort();
    
    // Populate filter dropdown
    if (filterSelect) {
        const currentFilter = filterSelect.value;
        filterSelect.innerHTML = '<option value="">All Modules</option>' +
            modules.map(m => `<option value="${escapeHtml(m)}">${escapeHtml(m)}</option>`).join("");
        filterSelect.value = currentFilter;
    }
    
    // Populate new tab module dropdown
    if (newTabModuleSelect) {
        newTabModuleSelect.innerHTML = '<option value="">Select Module...</option>' +
            modules.map(m => `<option value="${escapeHtml(m)}">${escapeHtml(m)}</option>`).join("");
    }
    
    // Filter tabs
    const filter = filterSelect?.value || "";
    const filteredTabs = filter 
        ? tabMappings.filter(t => t.module === filter)
        : tabMappings;
    
    // Render count
    const countEl = container.previousElementSibling;
    if (countEl && countEl.classList.contains("tab-mgmt-count")) {
        countEl.textContent = `Showing ${filteredTabs.length} of ${tabMappings.length} tab mappings`;
    }
    
    // Render tabs
    if (filteredTabs.length === 0) {
        container.innerHTML = `<div class="tab-mgmt-empty">No tab mappings found. Add one below.</div>`;
        return;
    }
    
    container.innerHTML = filteredTabs.map((tab, idx) => `
        <div class="tab-mgmt-item" data-row="${tab.rowIndex}">
            <span class="tab-mgmt-module">${escapeHtml(tab.module)}</span>
            <span class="tab-mgmt-name">${escapeHtml(tab.tabName)}</span>
            <span class="tab-mgmt-folder">${escapeHtml(tab.folder || "—")}</span>
            <button type="button" class="tab-mgmt-remove" data-row="${tab.rowIndex}" title="Remove mapping">×</button>
        </div>
    `).join("");
    
    // Wire up remove buttons
    container.querySelectorAll(".tab-mgmt-remove").forEach(btn => {
        btn.addEventListener("click", () => removeTabMapping(parseInt(btn.dataset.row, 10)));
    });
}

// Valid module keys for module-prefix category
const VALID_MODULE_KEYS = ["payroll-recorder", "pto-accrual", "credit-card-expense", "commission-calc", "system"];

/**
 * Add a new tab mapping with validation
 */
async function addTabMapping() {
    const moduleSelect = document.getElementById("newTabModule");
    const tabNameInput = document.getElementById("newTabName");
    const folderInput = document.getElementById("newTabFolder");
    
    const module = moduleSelect?.value?.trim();
    const tabName = tabNameInput?.value?.trim();
    let folder = folderInput?.value?.trim() || "";
    
    // ===== VALIDATION =====
    
    if (!module) {
        showValidationError("Please select a module.");
        return;
    }
    
    if (!tabName) {
        showValidationError("Please enter a tab name.");
        return;
    }
    
    // Validate and normalize the module key (Value2)
    if (!folder) {
        // Auto-populate folder based on module name
        const moduleKeyMap = {
            "Payroll Recorder": "payroll-recorder",
            "PTO Accrual": "pto-accrual",
            "Employee Roster": "employee-roster",
            "Headcount Review": "employee-roster"
        };
        folder = moduleKeyMap[module] || module.toLowerCase().replace(/\s+/g, "-");
        if (folderInput) folderInput.value = folder;
    }
    
    // Normalize the folder value
    const normalizedFolder = folder.toLowerCase().trim().replace(/[\s_]+/g, "-");
    
    // Check if folder is a valid module key
    if (!VALID_MODULE_KEYS.includes(normalizedFolder)) {
        showValidationError(
            `"${folder}" is not a valid module key.\n\n` +
            `Valid options:\n• ${VALID_MODULE_KEYS.join("\n• ")}\n\n` +
            `Please update the Module field.`
        );
        return;
    }
    
    // Check for duplicate
    if (tabMappings.some(t => t.tabName === tabName && t.module === module)) {
        showValidationError("This tab is already mapped to this module.");
        return;
    }
    
    // Check tab name doesn't have invalid characters
    if (/[\\/:*?"<>|]/.test(tabName)) {
        showValidationError("Tab name contains invalid characters: \\ / : * ? \" < > |");
        return;
    }
    
    if (!hasExcelRuntime()) {
        logAdmin("Excel not available", "error");
        return;
    }
    
    try {
        await Excel.run(async (context) => {
            const configSheet = context.workbook.worksheets.getItem("SS_PF_Config");
            const usedRange = configSheet.getUsedRange();
            usedRange.load("rowCount");
            await context.sync();
            
            const nextRow = usedRange.rowCount + 1;
            const newRange = configSheet.getRange(`A${nextRow}:D${nextRow}`);
            newRange.values = [["Tab Structure", module, tabName, normalizedFolder]];
            await context.sync();
            
            // Add to local state
            tabMappings.push({
                rowIndex: nextRow,
                module: module,
                tabName: tabName,
                folder: normalizedFolder
            });
            
            // Clear inputs
            if (tabNameInput) tabNameInput.value = "";
            if (folderInput) folderInput.value = "";
            
            // Re-render
            renderTabMappings();
            
            logAdmin(`✓ Added tab mapping: ${tabName} → ${module} (${normalizedFolder})`, "success");
        });
    } catch (error) {
        console.error("Error adding tab mapping:", error);
        logAdmin(`Error: ${error.message}`, "error");
    }
}

/**
 * Show validation error in admin log (Office Add-in compatible)
 */
function showValidationError(message) {
    logAdmin(`❌ Validation Error: ${message}`, "error");
    // Also highlight the log area
    const logArea = document.getElementById("adminLog");
    if (logArea) {
        logArea.scrollTop = logArea.scrollHeight;
    }
}

/**
 * Remove a tab mapping
 */
async function removeTabMapping(rowIndex) {
    const tab = tabMappings.find(t => t.rowIndex === rowIndex);
    if (!tab) return;
    
    // Note: window.confirm() is not supported in Office Add-ins
    // Proceeding directly - the remove button click is intentional
    
    if (!hasExcelRuntime()) {
        logAdmin("Excel not available", "error");
        return;
    }
    
    try {
        await Excel.run(async (context) => {
            const configSheet = context.workbook.worksheets.getItem("SS_PF_Config");
            const row = configSheet.getRange(`A${rowIndex}:D${rowIndex}`);
            row.delete(Excel.DeleteShiftDirection.up);
            await context.sync();
            
            // Remove from local state and reload (row indices changed)
            await loadTabMappings();
            renderTabMappings();
            
            logAdmin(`Removed tab mapping: ${tab.tabName}`, "success");
        });
    } catch (error) {
        console.error("Error removing tab mapping:", error);
        logAdmin(`Error: ${error.message}`, "error");
    }
}

// =============================================================================
// SHARED CONFIG
// =============================================================================

async function loadSharedConfigFromExcel() {
    if (!hasExcelRuntime()) {
        logAdmin("Excel not available - using defaults", "info");
        return;
    }
    
    try {
        await Excel.run(async (context) => {
            const configSheet = context.workbook.worksheets.getItemOrNullObject("SS_PF_Config");
            configSheet.load("isNullObject");
            await context.sync();
            
            if (configSheet.isNullObject) {
                logAdmin("SS_PF_Config sheet not found", "info");
                return;
            }
            
            const range = configSheet.getUsedRangeOrNullObject();
            range.load("values");
            await context.sync();
            
            if (range.isNullObject) return;
            
            const values = range.values || [];
            const configMap = buildConfigMap(values);
            
            // Populate shared config fields (check new + legacy names)
            sharedConfig.companyName = configMap.get("SS_Company_Name") || configMap.get("Company_Name") || "";
            sharedConfig.defaultReviewer = configMap.get("SS_Default_Reviewer") || configMap.get("Default_Reviewer") || "";
            sharedConfig.accountingLink = configMap.get("SS_Accounting_Software") || configMap.get("Accounting_Software_Link") || "";
            sharedConfig.payrollLink = configMap.get("SS_Payroll_Provider") || configMap.get("Payroll_Provider_Link") || "";
            
            // Update UI
            document.getElementById("sharedCompanyName").value = sharedConfig.companyName;
            document.getElementById("sharedReviewerName").value = sharedConfig.defaultReviewer;
            document.getElementById("sharedAccountingLink").value = sharedConfig.accountingLink;
            document.getElementById("sharedPayrollLink").value = sharedConfig.payrollLink;
            
            logAdmin("Loaded shared config from SS_PF_Config", "success");
        });
    } catch (error) {
        logAdmin(`Error loading config: ${error.message}`, "error");
    }
}

async function saveSharedConfig() {
    // Update local state from UI
    sharedConfig.companyName = document.getElementById("sharedCompanyName")?.value || "";
    sharedConfig.defaultReviewer = document.getElementById("sharedReviewerName")?.value || "";
    sharedConfig.accountingLink = document.getElementById("sharedAccountingLink")?.value || "";
    sharedConfig.payrollLink = document.getElementById("sharedPayrollLink")?.value || "";
    
    if (!hasExcelRuntime()) {
        logAdmin("Excel not available - config saved locally only", "info");
        return;
    }
    
    try {
        await Excel.run(async (context) => {
            const configSheet = context.workbook.worksheets.getItemOrNullObject("SS_PF_Config");
            configSheet.load("isNullObject");
            await context.sync();
            
            if (configSheet.isNullObject) {
                logAdmin("SS_PF_Config sheet not found - creating...", "info");
                // Would create sheet here in production
                return;
            }
            
            // Write shared config values (using new SS_ prefix)
            const fieldsToWrite = [
                { field: "SS_Company_Name", value: sharedConfig.companyName, category: "shared" },
                { field: "SS_Default_Reviewer", value: sharedConfig.defaultReviewer, category: "shared" },
                { field: "SS_Accounting_Software", value: sharedConfig.accountingLink, category: "shared" },
                { field: "SS_Payroll_Provider", value: sharedConfig.payrollLink, category: "shared" }
            ];
            
            await writeConfigFields(context, configSheet, fieldsToWrite);
            await context.sync();
            
            logAdmin("Shared config saved to SS_PF_Config", "success");
        });
    } catch (error) {
        logAdmin(`Error saving config: ${error.message}`, "error");
    }
}

// =============================================================================
// UTILITIES
// =============================================================================

async function validateConfig() {
    logAdmin("Validating SS_PF_Config sheet...", "info");
    
    if (!hasExcelRuntime()) {
        logAdmin("Excel not available - cannot validate", "error");
        return;
    }
    
    // Valid categories for reference
    const VALID_CATEGORIES = {
        "module-prefix": { label: "Module Prefix", description: "Maps tab prefixes to modules" },
        "run-settings": { label: "Run Settings", description: "Per-period configuration values" },
        "step-notes": { label: "Step Notes", description: "Notes and sign-off data" },
        "shared": { label: "Shared", description: "Global settings" },
        "column-mapping": { label: "Column Mapping", description: "Maps source to target columns" }
    };
    
    const validModuleKeys = ["payroll-recorder", "pto-accrual", "credit-card-expense", "commission-calc", "system"];
    
    try {
        await Excel.run(async (context) => {
            const configSheet = context.workbook.worksheets.getItemOrNullObject("SS_PF_Config");
            await context.sync();
            
            if (configSheet.isNullObject) {
                logAdmin("❌ SS_PF_Config sheet not found!", "error");
                return;
            }
            
            const usedRange = configSheet.getUsedRangeOrNullObject();
            usedRange.load("values");
            await context.sync();
            
            if (usedRange.isNullObject || !usedRange.values || usedRange.values.length < 2) {
                logAdmin("⚠️ SS_PF_Config appears empty", "error");
                return;
            }
            
            const rows = usedRange.values;
            const headers = rows[0].map(h => String(h || "").toLowerCase().trim());
            logAdmin(`Found ${rows.length - 1} data rows`, "info");
            logAdmin(`Headers: ${rows[0].join(", ")}`, "info");
            
            let errors = 0;
            let warnings = 0;
            
            // Validate each row
            for (let i = 1; i < rows.length; i++) {
                const row = rows[i];
                const rowNum = i + 1;
                const category = String(row[0] || "").toLowerCase().trim().replace(/[\s_]+/g, "-");
                const field = String(row[1] || "").trim();
                const value = String(row[2] || "").trim();
                const value2 = String(row[3] || "").trim();
                
                // Check category is valid
                if (!category) {
                    logAdmin(`Row ${rowNum}: Missing Category`, "error");
                    errors++;
                    continue;
                }
                
                if (!VALID_CATEGORIES[category]) {
                    logAdmin(`Row ${rowNum}: Invalid Category "${row[0]}". Valid: Tab Structure, Run Settings, Step Notes, Shared, Column Mapping`, "error");
                    errors++;
                    continue;
                }
                
                // Check Field is not empty
                if (!field) {
                    logAdmin(`Row ${rowNum}: Missing Field name`, "error");
                    errors++;
                }
                
                // Category-specific validation
                if (category === "module-prefix") {
                    // Validate prefix format (should end with underscore)
                    if (!field.endsWith("_")) {
                        logAdmin(`Row ${rowNum}: Prefix "${field}" should end with underscore (e.g., "PR_")`, "error");
                        warnings++;
                    }
                    // Validate module key
                    const normalizedValue = value.toLowerCase().trim().replace(/[\s_]+/g, "-");
                    if (!validModuleKeys.includes(normalizedValue)) {
                        logAdmin(`Row ${rowNum}: Module "${value}" not recognized. Expected: ${validModuleKeys.join(", ")}`, "error");
                        warnings++;
                    }
                }
            }
            
            // Summary
            logAdmin("─".repeat(40), "info");
            if (errors === 0 && warnings === 0) {
                logAdmin("✓ SS_PF_Config is valid! No issues found.", "success");
            } else {
                if (errors > 0) logAdmin(`❌ ${errors} error(s) found`, "error");
                if (warnings > 0) logAdmin(`⚠️ ${warnings} warning(s) found`, "error");
            }
        });
    } catch (error) {
        logAdmin(`Validation error: ${error.message}`, "error");
    }
}

function generateAllFields() {
    logAdmin("Generating field names for all modules...", "info");
    
    const allFields = [];
    
    // Shared fields (SS_ prefix)
    allFields.push(
        { category: "shared", field: "SS_Company_Name", value: "" },
        { category: "shared", field: "SS_Default_Reviewer", value: "" },
        { category: "shared", field: "SS_Accounting_Software", value: "" },
        { category: "shared", field: "SS_Payroll_Provider", value: "" }
    );
    
    // Module-specific fields
    moduleRegistry.forEach(module => {
        const prefix = module.prefix;
        
        // Module default reviewer
        allFields.push({ category: "module", field: `${prefix}_Reviewer`, value: "" });
        
        // Step-specific fields
        module.steps.forEach(step => {
            allFields.push(
                { category: "step", field: `${prefix}_${step}_Notes`, value: "" },
                { category: "step", field: `${prefix}_${step}_Reviewer`, value: "" },
                { category: "step", field: `${prefix}_${step}_SignOff`, value: "" },
                { category: "step", field: `${prefix}_${step}_Complete`, value: "" }
            );
        });
    });
    
    logAdmin(`Generated ${allFields.length} field names:`, "success");
    allFields.forEach(f => logAdmin(`  ${f.field}`, "info"));
    
    return allFields;
}

async function writeAllFieldsToSheet() {
    logAdmin("Writing fields to SS_PF_Config...", "info");
    console.log("writeAllFieldsToSheet called");
    
    const allFields = generateAllFields();
    console.log("Generated fields:", allFields.length);
    
    // Check if PrairieForge is available
    console.log("PrairieForge available:", !!window.PrairieForge);
    console.log("writeAllSharedConfig available:", !!(window.PrairieForge && window.PrairieForge.writeAllSharedConfig));
    
    if (window.PrairieForge && window.PrairieForge.writeAllSharedConfig) {
        try {
            logAdmin(`Attempting to write ${allFields.length} fields...`, "info");
            const result = await window.PrairieForge.writeAllSharedConfig(allFields);
            console.log("Write result:", result);
            
            if (result.error) {
                logAdmin(`Error: ${result.error}`, "error");
            } else {
                logAdmin(`Done! Added ${result.added} new fields, skipped ${result.skipped} existing fields.`, "success");
            }
            
            return result;
        } catch (err) {
            console.error("Write error:", err);
            logAdmin(`Exception: ${err.message}`, "error");
            return { added: 0, skipped: 0, error: err.message };
        }
    } else {
        logAdmin("PrairieForge.writeAllSharedConfig not available - is Common/common.js loaded?", "error");
        return { added: 0, skipped: 0, error: "Function not available" };
    }
}

async function exportConfig() {
    logAdmin("Exporting configuration...", "info");
    
    const exportData = {
        timestamp: new Date().toISOString(),
        sharedConfig,
        moduleRegistry
    };
    
    const blob = new Blob([JSON.stringify(exportData, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `pf-config-export-${new Date().toISOString().slice(0, 10)}.json`;
    link.click();
    URL.revokeObjectURL(url);
    
    logAdmin("Config exported to JSON file", "success");
}

// =============================================================================
// HELPERS
// =============================================================================

function hasExcelRuntime() {
    return typeof Excel !== "undefined" && typeof Excel.run === "function";
}

function buildConfigMap(values) {
    const map = new Map();
    if (!values || !values.length) return map;
    
    const headers = (values[0] || []).map(h => String(h || "").toLowerCase().trim());
    const fieldIdx = headers.findIndex(h => h === "field");
    const valueIdx = headers.findIndex(h => h === "value");
    
    if (fieldIdx === -1 || valueIdx === -1) return map;
    
    values.slice(1).forEach(row => {
        const field = String(row[fieldIdx] || "").trim();
        const value = row[valueIdx];
        if (field) map.set(field, value);
    });
    
    return map;
}

async function writeConfigFields(context, sheet, fields) {
    const range = sheet.getUsedRangeOrNullObject();
    range.load("values");
    await context.sync();
    
    if (range.isNullObject) return;
    
    const values = range.values || [];
    const headers = (values[0] || []).map(h => String(h || "").toLowerCase().trim());
    const categoryIdx = headers.findIndex(h => h === "category");
    const fieldIdx = headers.findIndex(h => h === "field");
    const valueIdx = headers.findIndex(h => h === "value");
    
    if (fieldIdx === -1 || valueIdx === -1) {
        logAdmin("Config sheet missing Field/Value columns", "error");
        return;
    }
    
    // Find and update each field
    for (const { field, value, category } of fields) {
        let rowIndex = values.slice(1).findIndex(row => 
            String(row[fieldIdx] || "").trim() === field
        );
        
        if (rowIndex >= 0) {
            // Update existing row
            const cellRange = sheet.getCell(rowIndex + 1, valueIdx);
            cellRange.values = [[value]];
        } else {
            // Add new row (append at end)
            const newRow = new Array(headers.length).fill("");
            if (categoryIdx >= 0) newRow[categoryIdx] = category || "";
            newRow[fieldIdx] = field;
            newRow[valueIdx] = value;
            
            const lastRow = values.length;
            const appendRange = sheet.getRangeByIndexes(lastRow, 0, 1, headers.length);
            appendRange.values = [newRow];
        }
    }
}

function logAdmin(message, type = "info") {
    const log = document.getElementById("adminLog");
    if (!log) return;
    
    // Remove placeholder
    const placeholder = log.querySelector(".admin-log-placeholder");
    if (placeholder) placeholder.remove();
    
    const entry = document.createElement("p");
    entry.className = `admin-log-entry admin-log-${type}`;
    entry.textContent = message;
    log.appendChild(entry);
    
    // Scroll to bottom
    log.scrollTop = log.scrollHeight;
}

// Initialize admin on load
initAdmin();
