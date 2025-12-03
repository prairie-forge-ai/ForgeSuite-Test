export const VERSION = "1.0.0.7";

export const SHEET_NAMES = {
    CONFIG: "SS_PF_Config",
    DATA: "PR_Data",
    DATA_CLEAN: "PR_Data_Clean",
    EXPENSE_MAPPING: "PR_Expense_Mapping",
    EXPENSE_REVIEW: "PR_Expense_Review",
    JE_DRAFT: "PR_JE_Draft",
    ARCHIVE_SUMMARY: "PR_Archive_Summary"
};

export const PAY_CATEGORY_HEADERS = [
    "Regular Hours",
    "Overtime Hours",
    "Holiday Hours",
    "Vacation Hours",
    "Sick Hours",
    "Bonus",
    "Commission",
    "Reimbursement",
    "Other Pay",
    "Federal Taxes",
    "State Taxes",
    "Social Security Tax",
    "Medicare Tax",
    "State Disability",
    "Unemployment Tax",
    "Workers Comp",
    "Health Insurance",
    "Dental Insurance",
    "Vision Insurance",
    "Life Insurance",
    "401k Employee",
    "401k Employer",
    "Other Deductions"
];

export const SHEET_BLUEPRINTS = [
    {
        name: "Instructions",
        description: "How to use the Prairie Forge payroll template"
    },
    {
        name: "Data_Input",
        description: "Paste WellsOne export data here"
    },
    {
        name: SHEET_NAMES.CONFIG,
        description: "Prairie Forge shared configuration storage (all modules)"
    },
    // PR_Config removed - consolidated into SS_PF_Config
    {
        name: "Config_Keywords",
        description: "Keyword-based account mapping rules"
    },
    {
        name: "Config_Accounts",
        description: "Account rewrite rules"
    },
    {
        name: "Config_Locations",
        description: "Location normalization rules"
    },
    {
        name: "Config_Vendors",
        description: "Vendor-specific overrides"
    },
    {
        name: "Config_Settings",
        description: "Prairie Forge system settings"
    },
    {
        name: SHEET_NAMES.EXPENSE_MAPPING,
        description: "Expense category mappings"
    },
    {
        name: SHEET_NAMES.DATA,
        description: "Processed payroll data staging"
    },
    {
        name: SHEET_NAMES.DATA_CLEAN,
        description: "Cleaned and validated payroll data"
    },
    {
        name: SHEET_NAMES.EXPENSE_REVIEW,
        description: "Expense review workspace"
    },
    {
        name: SHEET_NAMES.JE_DRAFT,
        description: "Journal entry preparation area"
    }
];

export const TABLE_BLUEPRINTS = [
    {
        sheetName: "Config_Keywords",
        tableName: "KeywordMappings",
        description: "Keyword-based account mapping rules",
        headers: ["Keyword", "New_Account", "New_Description", "Priority"],
        sampleRows: [
            ["meal", "4980", "Meals & Entertainment", 60],
            ["food", "4980", "Meals & Entertainment", 60],
            ["restaurant", "4980", "Meals & Entertainment", 60],
            ["software", "4250", "Software & Subscriptions", 30],
            ["office supplies", "4800", "Office Supplies", 40],
            ["fuel", "4700", "Vehicle Fuel", 50],
            ["gas", "4700", "Vehicle Fuel", 50]
        ]
    },
    {
        sheetName: "Config_Accounts",
        tableName: "AccountRewrites",
        description: "Account rewrite rules",
        headers: ["Old_Account", "New_Account", "Condition"],
        sampleRows: [
            ["4620", "4980", "meal_detected"],
            ["4500", "4800", "office_supplies"]
        ]
    },
    {
        sheetName: "Config_Locations",
        tableName: "LocationCorrections",
        description: "Location/dept normalization",
        headers: ["Old_Location", "New_Location", "Department"],
        sampleRows: [
            ["pf", "PF", "Operations"],
            ["prairie forge", "PF", "Operations"],
            ["admin", "ADMIN", "Administration"]
        ]
    },
    {
        sheetName: "Config_Vendors",
        tableName: "VendorRules",
        description: "Vendor-specific overrides",
        headers: ["Vendor_Pattern", "Account", "Description", "Location"],
        sampleRows: [
            ["amazon", "4800", "Office Supplies", "PF"],
            ["staples", "4800", "Office Supplies", "PF"],
            ["shell", "4700", "Vehicle Fuel", "PF"]
        ]
    },
    {
        sheetName: "Config_Settings",
        tableName: "Settings",
        description: "General settings",
        headers: ["Setting_Name", "Value", "Type"],
        sampleRows: [
            ["capitalizationThreshold", 3000, "number"],
            ["version", VERSION, "string"],
            ["defaultLocation", "PF", "string"]
        ]
    }
];

export const CONFIG_FIELDS = [
    {
        category: "Company Profile",
        field: "Company Name",
        description: "Displayed on instructions, exports, and journal entries",
        required: true
    },
    {
        category: "Company Profile",
        field: "Logo URL",
        description: "Public image URL for the instruction sheet header",
        required: false
    },
    {
        category: "Branding",
        field: "Brand Primary Color",
        description: "Hex value for table headers",
        defaultValue: "#0078d4",
        required: false
    },
    {
        category: "Branding",
        field: "Brand Accent Color",
        description: "Accent color used in instructions & highlights",
        defaultValue: "#106ebe",
        required: false
    },
    {
        category: "System Links",
        field: "Employee Mapping URL",
        description: "Link to the Employee Mapping workbook",
        required: true
    },
    {
        category: "System Links",
        field: "Payroll Provider URL",
        description: "Source report link or dashboard for payroll",
        required: true
    },
    {
        category: "System Links",
        field: "Accounting Import URL",
        description: "Destination folder or system import link",
        required: false
    },
    {
        category: "System Links",
        field: "Archive Folder URL",
        description: "Location for storing processed payroll archives",
        required: false
    },
    {
        category: "Run Settings",
        field: "Payroll Date (YYYY-MM-DD)",
        description: "Use ISO format; update each payroll run",
        required: true
    },
    {
        category: "Run Settings",
        field: "Reporting Period",
        description: "Readable label for the payroll period (e.g., Jan 2025)",
        required: true
    },
    {
        category: "Run Settings",
        field: "JE Transaction ID",
        description: "Unique identifier for exported journal entries",
        required: false
    },
    {
        category: "Run Settings",
        field: "Builder Mode",
        description: "TRUE keeps consultant tools visible; set to FALSE before handing the workbook to customers.",
        defaultValue: "TRUE",
        required: true
    }
];

export const WORKFLOW_STEPS = [
    {
        id: 0,
        title: "Configuration Setup",
        summary: "Company profile, branding, and run settings",
        description: "Keep the SS_PF_Config table current before every payroll run so downstream sheets inherit the right colors, links, and identifiers.",
        icon: "🧭",
        ctaLabel: "Open Configuration Form",
        statusHint: "Configuration edits happen inside the PF_Config table and are available to every step instantly.",
        highlights: [
            {
                label: "Company Profile",
                detail: "Company name, logos, payroll date, reporting period."
            },
            {
                label: "Brand Identity",
                detail: "Primary + accent colors carry through dashboards and exports."
            },
            {
                label: "System Links",
                detail: "Quick jumps to HRIS, payroll provider, accounting import, and archive folders."
            }
        ],
        checklist: [
            "Review profile, branding, links, and run settings each payroll cycle.",
            "Click Save to write updates back to the SS_PF_Config sheet."
        ]
    },
    {
        id: 1,
        title: "Import Payroll Data",
        summary: "Paste the payroll provider export into the Data sheet",
        description: "Pull your payroll data from your provider’s portal and paste it into the Data tab. If the columns match, just paste the rows; if they don’t, paste your headers and data right over the top. Formatting is fully automated.",
        icon: "📥",
        ctaLabel: "Prepare Import Sheet",
        statusHint: "The Data worksheet is activated so you can paste the latest provider export.",
        highlights: [
            {
                label: "Source File",
                detail: "Use WellsOne/ADP export with every pay category column visible."
            },
            {
                label: "Structure",
                detail: "Row 2 headers, row 3+ data, no blank columns, totals removed."
            },
            {
                label: "Quality",
                detail: "Spot-check employee counts and pay period filters before moving on."
            }
        ],
        checklist: [
            "Download the payroll detail export covering this pay period.",
            "Paste values into the Data sheet starting at cell A3.",
            "Confirm all pay category headers remain intact and spelled consistently."
        ]
    },
    {
        id: 2,
        title: "Headcount Review",
        summary: "Ensure roster and payroll rows agree",
        description: "This step is optional, but strongly recommended. A centralized employee roster keeps every payroll-related workbook aligned while ensuring key attributes such as department and location stay consistent each pay period.",
        icon: "👥",
        ctaLabel: "Launch Headcount Review",
        statusHint: "Data and mapping sheets are surfaced so you can reconcile roster counts before validation.",
        highlights: [
            {
                label: "Roster Alignment",
                detail: "Compare active roster to the employees present in the Data sheet."
            },
            {
                label: "Variance Tracking",
                detail: "Note missing departments or unexpected hires before the validation run."
            },
            {
                label: "Approvals",
                detail: "Capture reviewer initials and date for audit coverage."
            }
        ],
        checklist: [
            "Filter the Data sheet by Department to ensure every team appears.",
            "Look for duplicate or out-of-period employees and resolve upstream.",
            "Document findings on the Headcount Review tab or your tracker of choice."
        ]
    },
    {
        id: 3,
        title: "Validate & Reconcile",
        summary: "Normalize payroll data and reconcile totals",
        description: "Automatically rebuild the PR_Data_Clean sheet, confirm payroll totals match, and prep the bank reconciliation before moving to Expense Review.",
        icon: "✅",
        statusHint: "Run completes automatically when you enter this step.",
        highlights: [
            {
                label: "Normalized Data",
                detail: "Creates one row per employee and payroll category."
            },
            {
                label: "Outputs",
                detail: "Data_Clean rebuilt with payroll category + mapping details."
            },
            {
                label: "Reconciliation",
                detail: "Displays PR_Data vs PR_Data_Clean totals plus bank comparison."
            }
        ]
    },
    {
        id: 4,
        title: "Expense Review",
        summary: "Generate an executive-ready payroll summary",
        description: "Build a six-period payroll dashboard (current + five prior), including department-level breakouts and variance indicators, plus notes and CoPilot guidance.",
        icon: "📊",
        statusHint: "Selecting this step rebuilds PR_Expense_Review automatically.",
        highlights: [
            {
                label: "Time Series",
                detail: "Shows six consecutive payroll periods."
            },
            {
                label: "Departments",
                detail: "All-in totals, burden rates, and headcount by department."
            },
            {
                label: "Guidance",
                detail: "Use CoPilot to summarize trends and capture review notes."
            }
        ],
        checklist: []
    },
    {
        id: 5,
        title: "Journal Entry Prep",
        summary: "Generate a QuickBooks-ready journal draft",
        description: "Create the JE Draft sheet with the headers QuickBooks Online/Desktop expect so you only need to paste balanced lines.",
        icon: "🧾",
        ctaLabel: "Generate JE Draft",
        statusHint: "JE Draft contains headers for RefNumber, TxnDate, account columns, and line descriptions.",
        highlights: [
            {
                label: "Structure",
                detail: "Debit/Credit columns prepared with standard import headers."
            },
            {
                label: "Context",
                detail: "JE Transaction ID from configuration is referenced for traceability."
            },
            {
                label: "Next Step",
                detail: "Populate amounts from Expense Review to finalize the journal."
            }
        ],
        checklist: [
            "Ensure validation + expense review steps are complete.",
            "Run the generator to rebuild the JE Draft sheet.",
            "Paste balanced lines and export to QuickBooks / ERP import format."
        ]
    },
    {
        id: 6,
        title: "Archive & Clear",
        summary: "Snapshot workpapers and reset working tabs",
        description: "Capture a log of each payroll run, note the archive destination, and optionally clear staging sheets for the next cycle.",
        icon: "🗂️",
        ctaLabel: "Create Archive Summary",
        statusHint: "Archive summary headers help you log when data was exported and where the files live.",
        highlights: [
            {
                label: "Run Log",
                detail: "Payroll date, reporting period, JE ID, and who processed the run."
            },
            {
                label: "Storage",
                detail: "Link to the Archive folder defined in configuration."
            },
            {
                label: "Reset",
                detail: "Reminder to clear Data/Data_Clean once files are safely archived."
            }
        ],
        checklist: [
            "Record archive destination links and reviewer approvals.",
            "Copy Data/Data_Clean/JE Draft tabs to the archive workbook if needed.",
            "Clear sensitive data so the template is ready for the next payroll."
        ]
    }
];

const globalBuilderAllowlist =
    (typeof window !== "undefined" && Array.isArray(window.PF_BUILDER_ALLOWLIST)
        ? window.PF_BUILDER_ALLOWLIST
        : []
    ).map((entry) => String(entry || "").trim().toLowerCase());

export const BUILDER_ALLOWED_USERS = globalBuilderAllowlist;

export const BUILDER_VISIBILITY_SHEETS = [
    "Instructions",
    "Config_Keywords",
    "Config_Accounts",
    "Config_Locations",
    "Config_Vendors",
    "Config_Settings"
];

export const METRIC_ELEMENT_IDS = {
    dataEmployees: "metric-data-employees",
    rosterEmployees: "metric-roster-employees",
    difference: "metric-difference",
    nameDifferences: "metric-name-diffs",
    departmentMismatches: "metric-dept-mismatches"
};

export const JE_METRIC_ELEMENT_IDS = {
    sourceTotal: "je-metric-source",
    debitTotal: "je-metric-debit",
    creditTotal: "je-metric-credit",
    variance: "je-metric-variance"
};

export const LOGO_URL = "https://assets.prairieforge.ai/storage/v1/object/public/Other%20Public%20Material/Smartsheet_noBR.png";
