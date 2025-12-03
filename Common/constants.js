// Shared constants and utilities for Prairie Forge Payroll Recorder
// Used by both Google Sheets (Apps Script) and Excel (Office.js) versions

// Common worksheet names
export const WORKSHEET_NAMES = {
  EMPLOYEES: 'Employees',
  PAYROLL: 'PR_Data',
  VALIDATION: 'PR_Data_Clean',
  CONFIGURATION: 'SS_PF_Config',
  EXPENSE_REVIEW: 'PR_Expense_Review',
  JOURNAL_ENTRY: 'PR_JE_Draft',
  ARCHIVE: 'PR_Archive_Summary',
  HEADCOUNT: 'SS_Headcount_Review',
  RECONCILIATION: 'Reconciliation',
  IMPORTS: 'Imports',
  REPORTS: 'Reports',
  SUMMARY: 'Summary',
  DASHBOARD: 'Dashboard'
};

// Common table names for Excel
export const TABLE_NAMES = {
  EMPLOYEE_LIST: 'EmployeeTable',
  PAYROLL_DATA: 'PayrollTable', 
  CONFIGURATION_SETTINGS: 'ConfigTable',
  EXPENSE_ITEMS: 'ExpenseTable',
  JOURNAL_ENTRIES: 'JournalTable'
};

// Branding constants
export const BRANDING = {
  COMPANY_NAME: 'Prairie Forge LLC',
  PRODUCT_NAME: 'Prairie Forge Tools',
  SUPPORT_URL: 'https://prairieforge.ai/support',
  ADA_IMAGE_URL: 'https://assets.prairieforge.ai/storage/v1/object/public/Other%20Public%20Material/Prairie%20Forge/Ada%20Image.png'
};

// GitHub repository information
export const REPOSITORY = {
  OWNER: 'prairie-forge-ai',
  NAME: 'prairie-forge-templates',
  GITHUB_PAGES_BASE: 'https://prairie-forge-ai.github.io/prairie-forge-templates'
};
