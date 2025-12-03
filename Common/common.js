/**
 * Prairie Forge Common Utilities
 * Shared functions used across all modules
 */

// Create global PrairieForge namespace
window.PrairieForge = window.PrairieForge || {};

const PF_MODULE_LOADER_URL = 'https://prairie-forge-ai.github.io/business-tools/module-loader.html';

/**
 * Configuration
 */
PrairieForge.config = {
  supabaseUrl: null, // Will be set by each module
  manifestToken: null, // Customer's unique token
  currentModule: null,
  moduleLoaderUrl: PF_MODULE_LOADER_URL
};

/**
 * Initialize Prairie Forge with customer configuration
 */
PrairieForge.initialize = async function(manifestToken) {
  PrairieForge.config.manifestToken = manifestToken;
};

/**
 * Normalize date from various formats to YYYY-MM-DD
 * @param {string|Date} dateValue - Date in various formats
 * @returns {string|null} - Normalized date or null if invalid
 */
PrairieForge.normalizeDate = function(dateValue) {
  if (!dateValue) return null;

  try {
    let date;

    // If already a Date object
    if (dateValue instanceof Date) {
      date = dateValue;
    }
    // If Excel serial date (number)
    else if (typeof dateValue === 'number') {
      // Excel dates are days since 1900-01-01
      const excelEpoch = new Date(1900, 0, 1);
      date = new Date(excelEpoch.getTime() + (dateValue - 2) * 86400000);
    }
    // If string
    else {
      date = new Date(dateValue);
    }

    if (isNaN(date.getTime())) {
      return null;
    }

    // Return in YYYY-MM-DD format
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');

    return `${year}-${month}-${day}`;
  } catch (error) {
    console.error('Error normalizing date:', error);
    return null;
  }
};

/**
 * Normalize currency value to number
 * @param {string|number} value - Currency value (e.g., "$1,234.56")
 * @returns {number|null} - Numeric value or null if invalid
 */
PrairieForge.normalizeCurrency = function(value) {
  if (value === null || value === undefined || value === '') return null;

  // If already a number
  if (typeof value === 'number') return value;

  // Remove currency symbols, commas, spaces
  const cleaned = String(value).replace(/[$,\s]/g, '');

  const number = parseFloat(cleaned);

  return isNaN(number) ? null : number;
};

/**
 * Format number as currency
 * @param {number} value - Numeric value
 * @param {string} currency - Currency code (default: USD)
 * @returns {string} - Formatted currency string
 */
PrairieForge.formatCurrency = function(value, currency = 'USD') {
  if (value === null || value === undefined) return '';

  return new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: currency
  }).format(value);
};

/**
 * Format a numeric value with commas and fixed decimals.
 * Returns an empty string if the value is not a finite number.
 */
PrairieForge.formatNumber = function(value, decimals = 2) {
  if (value === null || value === undefined) return '';
  const num = Number(value);
  if (!Number.isFinite(num)) return '';
  return num.toLocaleString('en-US', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals
  });
};

/**
 * Apply number formatting to inputs that opt in via the pf-input-number class.
 * Formats on load and on blur; does not affect dates because date inputs do not use that class.
 */
PrairieForge.applyNumberFormatting = function(root = document) {
  const inputs = root.querySelectorAll('input.pf-input-number');
  inputs.forEach((input) => {
    const formatValue = () => {
      const formatted = PrairieForge.formatNumber(input.value);
      if (formatted) input.value = formatted;
    };
    // Initial format if a value exists
    if (input.value) formatValue();
    // Re-format when leaving the field
    input.removeEventListener('blur', formatValue);
    input.addEventListener('blur', formatValue);
  });
};

document.addEventListener('DOMContentLoaded', () => {
  if (window.PrairieForge?.applyNumberFormatting) {
    window.PrairieForge.applyNumberFormatting(document);
  }
});

// SVG icons for status feedback (Lucide-style)
const STATUS_ICON_SUCCESS = `<svg class="pf-icon pf-status-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><path d="m9 12 2 2 4-4"/></svg>`;
const STATUS_ICON_ERROR = `<svg class="pf-icon pf-status-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><path d="m15 9-6 6"/><path d="m9 9 6 6"/></svg>`;
const STATUS_ICON_WARNING = `<svg class="pf-icon pf-status-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3"/><path d="M12 9v4"/><path d="M12 17h.01"/></svg>`;
const STATUS_ICON_INFO = `<svg class="pf-icon pf-status-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><path d="M12 16v-4"/><path d="M12 8h.01"/></svg>`;

const STATUS_ICONS = {
    success: STATUS_ICON_SUCCESS,
    error: STATUS_ICON_ERROR,
    warning: STATUS_ICON_WARNING,
    info: STATUS_ICON_INFO
};

/**
 * Show error message to user
 * @param {string} message - Error message to display
 * @param {number} duration - Duration in ms (default: 5000)
 */
PrairieForge.showError = function(message, duration = 5000) {
  console.error('Error:', message);

  const alert = document.createElement('div');
  alert.className = 'pf-alert pf-alert-error';
  alert.innerHTML = `
    <div class="pf-alert-content">
      <span class="pf-alert-icon">${STATUS_ICON_ERROR}</span>
      <span class="pf-alert-message">${message}</span>
    </div>
  `;

  document.body.appendChild(alert);

  // Auto-remove after duration
  setTimeout(() => {
    alert.style.opacity = '0';
    setTimeout(() => alert.remove(), 300);
  }, duration);
};

/**
 * Show success message to user
 * @param {string} message - Success message to display
 * @param {number} duration - Duration in ms (default: 3000)
 */
PrairieForge.showSuccess = function(message, duration = 3000) {
  console.log('Success:', message);

  const alert = document.createElement('div');
  alert.className = 'pf-alert pf-alert-success';
  alert.innerHTML = `
    <div class="pf-alert-content">
      <span class="pf-alert-icon">${STATUS_ICON_SUCCESS}</span>
      <span class="pf-alert-message">${message}</span>
    </div>
  `;

  document.body.appendChild(alert);

  // Auto-remove after duration
  setTimeout(() => {
    alert.style.opacity = '0';
    setTimeout(() => alert.remove(), 300);
  }, duration);
};

/**
 * Render an inline status banner for persistent feedback within a container.
 * 
 * @param {Object} options
 * @param {string} options.type - Banner type: 'success' | 'warning' | 'error' | 'info'
 * @param {string} [options.title] - Optional bold title
 * @param {string} options.message - Main message text
 * @param {Object} [options.action] - Optional action button
 * @param {string} options.action.label - Button label
 * @param {string} [options.action.id] - Button ID for event binding
 * @param {Function} [options.escapeHtml] - HTML escape function
 * @returns {string} - HTML string for the status banner
 * 
 * @example
 * PrairieForge.renderStatusBanner({
 *     type: 'warning',
 *     title: 'Data Not Ready',
 *     message: 'Complete the Headcount Review step first.',
 *     action: { label: 'Go to Step 2', id: 'go-step-2-btn' }
 * });
 */
PrairieForge.renderStatusBanner = function({
    type = 'info',
    title = '',
    message = '',
    action = null,
    escapeHtml = null
} = {}) {
    const escape = escapeHtml || function(str) {
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    };
    
    const icon = STATUS_ICONS[type] || STATUS_ICONS.info;
    const titleHtml = title ? `<h4 class="pf-status-banner__title">${escape(title)}</h4>` : '';
    const actionHtml = action ? `
        <div class="pf-status-banner__action">
            <button type="button" class="pf-status-banner__btn" ${action.id ? `id="${escape(action.id)}"` : ''}>
                ${escape(action.label)}
            </button>
        </div>
    ` : '';
    
    return `
        <div class="pf-status-banner pf-status-banner--${type}">
            <div class="pf-status-banner__icon">${icon}</div>
            <div class="pf-status-banner__content">
                ${titleHtml}
                <p class="pf-status-banner__message">${escape(message)}</p>
                ${actionHtml}
            </div>
        </div>
    `;
};

/**
 * Render a prerequisite checklist showing what's ready vs what's missing.
 * 
 * @param {Object} options
 * @param {Array<Object>} options.checks - Array of prerequisite checks
 * @param {string} options.checks[].label - Description of the requirement
 * @param {boolean} options.checks[].ready - Whether this requirement is met
 * @param {Object} [options.checks[].action] - Optional action for unmet requirements
 * @param {string} options.checks[].action.label - Button label
 * @param {string} [options.checks[].action.id] - Button ID for event binding
 * @param {string} [options.allReadyMessage] - Message when all checks pass
 * @param {string} [options.notReadyMessage] - Message when some checks fail
 * @param {boolean} [options.showSummary=true] - Whether to show the summary message
 * @param {Function} [options.escapeHtml] - HTML escape function
 * @returns {string} - HTML string for the prerequisite checklist
 * 
 * @example
 * PrairieForge.renderPrerequisiteCheck({
 *     checks: [
 *         { label: 'PTO Data imported', ready: true },
 *         { label: 'Employee roster synced', ready: false, action: { label: 'Sync Now', id: 'sync-btn' } }
 *     ],
 *     allReadyMessage: 'All prerequisites met!',
 *     notReadyMessage: 'Complete the items below to continue.'
 * });
 */
PrairieForge.renderPrerequisiteCheck = function({
    checks = [],
    allReadyMessage = 'All prerequisites met. Ready to proceed.',
    notReadyMessage = 'Complete the following items before continuing.',
    showSummary = true,
    escapeHtml = null
} = {}) {
    if (!checks.length) return '';
    
    const escape = escapeHtml || function(str) {
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    };
    
    const allReady = checks.every(c => c.ready);
    
    const items = checks.map(check => {
        const itemClass = check.ready ? 'pf-prereq-item--ready' : 'pf-prereq-item--missing';
        const icon = check.ready ? STATUS_ICON_SUCCESS : STATUS_ICON_WARNING;
        const actionHtml = !check.ready && check.action ? `
            <div class="pf-prereq-item__action">
                <button type="button" class="pf-prereq-item__btn" ${check.action.id ? `id="${escape(check.action.id)}"` : ''}>
                    ${escape(check.action.label)}
                </button>
            </div>
        ` : '';
        
        return `
            <div class="pf-prereq-item ${itemClass}">
                <div class="pf-prereq-item__icon">${icon}</div>
                <span class="pf-prereq-item__label">${escape(check.label)}</span>
                ${actionHtml}
            </div>
        `;
    }).join('');
    
    const summaryHtml = showSummary ? `
        <div class="pf-prereq-summary ${allReady ? 'pf-prereq-summary--ready' : 'pf-prereq-summary--blocked'}">
            ${escape(allReady ? allReadyMessage : notReadyMessage)}
        </div>
    ` : '';
    
    return `
        <div class="pf-prereq-container">
            <p class="pf-prereq-header">Prerequisites</p>
            <div class="pf-prereq-list">
                ${items}
            </div>
            ${summaryHtml}
        </div>
    `;
};

/**
 * Show/hide loading spinner
 * @param {boolean} show - True to show, false to hide
 * @param {string} message - Optional loading message
 */
PrairieForge.showLoading = function(show, message = 'Loading...') {
  let loader = document.getElementById('pf-loader');

  if (show) {
    if (!loader) {
      loader = document.createElement('div');
      loader.id = 'pf-loader';
      loader.className = 'pf-loader';
      loader.innerHTML = `
        <div class="pf-loader-content">
          <div class="pf-spinner"></div>
          <div class="pf-loader-message">${message}</div>
        </div>
      `;
      document.body.appendChild(loader);
    } else {
      loader.querySelector('.pf-loader-message').textContent = message;
      loader.style.display = 'flex';
    }
  } else {
    if (loader) {
      loader.style.display = 'none';
    }
  }
};

/**
 * Export data to CSV file
 * @param {Array} data - Array of objects to export
 * @param {string} filename - Name of file (default: export.csv)
 * @param {Array} columns - Optional array of column names to include
 */
PrairieForge.exportToCSV = function(data, filename = 'export.csv', columns = null) {
  if (!data || data.length === 0) {
    PrairieForge.showError('No data to export');
    return;
  }

  // Get column names
  const cols = columns || Object.keys(data[0]);

  // Build CSV content
  let csv = cols.join(',') + '\n';

  data.forEach(row => {
    const values = cols.map(col => {
      let value = row[col] || '';

      // Escape quotes and wrap in quotes if contains comma
      value = String(value).replace(/"/g, '""');
      if (value.includes(',') || value.includes('\n') || value.includes('"')) {
        value = `"${value}"`;
      }

      return value;
    });

    csv += values.join(',') + '\n';
  });

  // Create download link
  const blob = new Blob([csv], { type: 'text/csv' });
  const url = window.URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  link.click();

  window.URL.revokeObjectURL(url);

  PrairieForge.showSuccess(`Exported ${data.length} rows to ${filename}`);
};

/**
 * Resolve manifest token from config, PF_CONTEXT, or URL
 * @returns {string|null}
 */
PrairieForge.getManifestToken = function() {
  if (PrairieForge.config.manifestToken) {
    return PrairieForge.config.manifestToken;
  }
  if (window.PF_CONTEXT?.token) {
    PrairieForge.config.manifestToken = window.PF_CONTEXT.token;
    return PrairieForge.config.manifestToken;
  }
  const fromQuery = getQueryParamValue('token');
  if (fromQuery) {
    PrairieForge.config.manifestToken = fromQuery;
    return fromQuery;
  }
  return null;
};

/**
 * Build URL back to the hosted module selector
 * @returns {string}
 */
PrairieForge.buildModuleSelectorUrl = function() {
  const base = PrairieForge.config.moduleLoaderUrl || PF_MODULE_LOADER_URL;
  const params = new URLSearchParams();
  const token = PrairieForge.getManifestToken();
  if (token) {
    params.set('token', token);
  }
  params.set('v', Date.now().toString());
  return `${base}?${params.toString()}`;
};

/**
 * Attach a standard \"Back to modules\" button
 * @param {HTMLElement} element
 * @param {{label?: string, dark?: boolean}} options
 */
PrairieForge.attachModuleHomeButton = function(element, options = {}) {
  if (!element || element.dataset.pfModuleHome === 'true') return;
  const { label, dark = false } = options;
  if (label) {
    element.textContent = label;
  }
  if (!element.classList.contains('pf-apple-button')) {
    element.classList.add('pf-apple-button');
  }
  if (dark) {
    element.classList.add('pf-apple-button-dark');
  }
  if (element.tagName === 'BUTTON' && !element.hasAttribute('type')) {
    element.type = 'button';
  }
  element.addEventListener('click', (event) => {
    event.preventDefault();
    PrairieForge.navigateToModuleSelector();
  });
  element.dataset.pfModuleHome = 'true';
};

/**
 * Navigate to module selector (home page)
 */
PrairieForge.navigateToModuleSelector = function() {
  const target = PrairieForge.buildModuleSelectorUrl();
  window.location.href = target;
};

/**
 * Navigate to a specific module
 * @param {string} moduleKey - Module key to navigate to (e.g., 'payroll-recorder')
 */
PrairieForge.navigateToModule = function(moduleKey) {
  PrairieForge.showLoading(true, `Loading ${moduleKey}...`);

  // In production, this would load from Supabase Storage
  // For local testing, we'll navigate to the module folder
  window.location.href = `/${moduleKey}/`;
};

/**
 * Verify customer has access to a module
 * @param {string} moduleKey - Module key to check
 * @returns {Promise<boolean>} - True if has access
 */
PrairieForge.verifyAccess = async function(moduleKey) {
  if (!PrairieForge.config.manifestToken) {
    console.warn('No manifest token set');
    return false;
  }

  try {
    // This would call the verify-access Edge Function
    // For now, we'll just return true for local testing
    console.log(`Verifying access to ${moduleKey} for token ${PrairieForge.config.manifestToken}`);
    return true;
  } catch (error) {
    console.error('Error verifying access:', error);
    return false;
  }
};

/**
 * Get Excel data from a range
 * @param {string} sheetName - Name of worksheet (null for active sheet)
 * @param {string} range - Range to read (e.g., "A1:D10")
 * @returns {Promise<Array>} - Array of row arrays
 */
PrairieForge.getExcelData = async function(sheetName = null, range = null) {
  return Excel.run(async (context) => {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();

    const usedRange = range
      ? sheet.getRange(range)
      : sheet.getUsedRange();

    usedRange.load('values');
    await context.sync();

    return usedRange.values;
  });
};

/**
 * Write data to Excel
 * @param {Array} data - 2D array of data to write
 * @param {string} startCell - Starting cell (default: A1)
 * @param {string} sheetName - Sheet name (null for active sheet)
 */
PrairieForge.writeExcelData = async function(data, startCell = 'A1', sheetName = null) {
  return Excel.run(async (context) => {
    const sheet = sheetName
      ? context.workbook.worksheets.getItem(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();

    const range = sheet.getRange(startCell).getResizedRange(data.length - 1, data[0].length - 1);
    range.values = data;

    await context.sync();
  });
};

/**
 * Trim and normalize whitespace in string
 * @param {string} value - String to normalize
 * @returns {string} - Normalized string
 */
PrairieForge.normalizeString = function(value) {
  if (!value) return '';
  return String(value).trim().replace(/\s+/g, ' ');
};

/**
 * Validate email address
 * @param {string} email - Email to validate
 * @returns {boolean} - True if valid
 */
PrairieForge.validateEmail = function(email) {
  const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return re.test(String(email).toLowerCase());
};

/**
 * Mount a floating info button that opens a modal with page-specific content.
 * @param {Object} options
 * @param {string} options.title - Modal title
 * @param {string|HTMLElement} options.content - Modal body content
 * @param {string} [options.buttonId="pf-info-fab"] - Optional id for the button wrapper
 */
PrairieForge.mountInfoFab = function({ title = "Info", content = "", buttonId = "pf-info-fab" } = {}) {
  // Remove existing FAB with same ID to update content on step change
  const existing = document.getElementById(buttonId);
  if (existing) {
    existing.remove();
  }
  
  const fab = document.createElement("div");
  fab.className = "pf-info-fab";
  fab.id = buttonId;
  const btn = document.createElement("button");
  btn.type = "button";
  btn.setAttribute("aria-label", "Show info");
  btn.textContent = "i";
  btn.addEventListener("click", () => openInfoModal(title, content));
  fab.appendChild(btn);
  document.body.appendChild(fab);
};

function openInfoModal(title, content) {
  document.querySelector(".pf-info-overlay")?.remove();
  const overlay = document.createElement("div");
  overlay.className = "pf-info-overlay";
  const modal = document.createElement("div");
  modal.className = "pf-info-modal";
  modal.innerHTML = `
    <button type="button" class="pf-info-close" aria-label="Close info">&times;</button>
    <h3>${title}</h3>
    <div class="pf-info-body"></div>
  `;
  const body = modal.querySelector(".pf-info-body");
  if (typeof content === "string") {
    const p = document.createElement("p");
    p.innerHTML = content;
    body.appendChild(p);
  } else if (content instanceof HTMLElement) {
    body.appendChild(content);
  }
  modal.querySelector(".pf-info-close").addEventListener("click", () => overlay.remove());
  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) overlay.remove();
  });
  overlay.appendChild(modal);
  document.body.appendChild(overlay);
}

function getQueryParamValue(name) {
  try {
    const params = new URLSearchParams(window.location.search);
    return params.get(name);
  } catch (_error) {
    return null;
  }
}

// SVG icons for mismatch tiles (Lucide-style)
const MISMATCH_ICON_UP = `
    <svg class="pf-icon pf-mismatch-icon-svg" aria-hidden="true" focusable="false" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="m18 9-6-6-6 6"/><path d="M12 3v14"/><path d="M5 21h14"/>
    </svg>
`;

const MISMATCH_ICON_DOWN = `
    <svg class="pf-icon pf-mismatch-icon-svg" aria-hidden="true" focusable="false" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="m6 15 6 6 6-6"/><path d="M12 21V7"/><path d="M5 3h14"/>
    </svg>
`;

/**
 * Render mismatch tiles for headcount or roster reconciliation views.
 * 
 * @param {Object} options
 * @param {Array<string|Object>} options.mismatches - Array of mismatch items (strings or objects with employee/name)
 * @param {string} [options.label] - Label above the list (default: "Employees Driving the Difference")
 * @param {string} [options.sourceLabel] - Label for "missing from source" items (default: "Roster")
 * @param {string} [options.targetLabel] - Label for "missing from target" items (default: "PTO Data")
 * @param {Function} [options.escapeHtml] - HTML escape function (uses basic fallback if not provided)
 * @param {boolean} [options.showLabel] - Whether to show the label (default: true)
 * @returns {string} - HTML string for the mismatch section
 * 
 * @example
 * // Basic usage with string array
 * PrairieForge.renderMismatchTiles({
 *   mismatches: ["In roster, missing in PTO_Data: John Smith"],
 *   sourceLabel: "Roster",
 *   targetLabel: "PTO Data"
 * });
 * 
 * @example
 * // Usage with custom formatter (for department mismatches)
 * PrairieForge.renderMismatchTiles({
 *   mismatches: items,
 *   label: "Department Differences",
 *   formatter: (item) => ({ name: item.employee, source: `${item.rosterDept} → ${item.payrollDept}` })
 * });
 */
PrairieForge.renderMismatchTiles = function({
    mismatches = [],
    label = "Employees Driving the Difference",
    sourceLabel = "Roster",
    targetLabel = "PTO Data",
    escapeHtml = null,
    showLabel = true,
    formatter = null
} = {}) {
    // Filter out empty values
    const items = Array.isArray(mismatches) ? mismatches.filter(Boolean) : [];
    
    if (!items.length) {
        return '';
    }
    
    // Basic HTML escape fallback
    const escape = escapeHtml || function(str) {
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    };
    
    // Parse mismatch items and render tiles
    const tiles = items.map(item => {
        let name, sourceText, isMissingFromTarget;
        
        // Handle custom formatter (for complex objects like department mismatches)
        if (typeof formatter === 'function') {
            const formatted = formatter(item);
            name = formatted.name || '';
            sourceText = formatted.source || '';
            isMissingFromTarget = formatted.isMissingFromTarget ?? true;
        }
        // Handle string items (roster mismatches)
        else if (typeof item === 'string') {
            // Extract the name (everything after the colon)
            const colonIndex = item.indexOf(':');
            name = colonIndex > -1 ? item.substring(colonIndex + 1).trim() : item;
            
            // Detect direction based on common patterns
            const lowerItem = item.toLowerCase();
            isMissingFromTarget = lowerItem.includes('missing in pto') ||
                                   lowerItem.includes('missing in payroll') ||
                                   lowerItem.includes(`missing in ${targetLabel.toLowerCase()}`);
            
            sourceText = isMissingFromTarget ? `Missing from ${targetLabel}` : `Missing from ${sourceLabel}`;
        }
        // Handle object items
        else if (typeof item === 'object') {
            name = item.name || item.employee || '';
            sourceText = item.source || item.sourceText || '';
            isMissingFromTarget = item.isMissingFromTarget ?? true;
        }
        else {
            return '';
        }
        
        const tileClass = isMissingFromTarget ? 'pf-mismatch-tile--missing-source' : 'pf-mismatch-tile--missing-target';
        const icon = isMissingFromTarget ? MISMATCH_ICON_UP : MISMATCH_ICON_DOWN;
        
        return `
            <div class="pf-mismatch-tile ${tileClass}">
                <span class="pf-mismatch-icon">${icon}</span>
                <div class="pf-mismatch-body">
                    <span class="pf-mismatch-name">${escape(name)}</span>
                    <span class="pf-mismatch-source">${escape(sourceText)}</span>
                </div>
            </div>
        `;
    }).join('');
    
    const labelHtml = showLabel ? `<p class="pf-mismatch-label">${escape(label)}</p>` : '';
    
    return `
        <div class="pf-mismatch-container">
            ${labelHtml}
            <div class="pf-mismatch-grid">
                ${tiles}
            </div>
        </div>
    `;
};

// =============================================================================
// SHARED CONFIG READER
// =============================================================================

/**
 * Shared configuration field names
 */
// Shared config field names - SS_ prefix for shared settings
PrairieForge.SHARED_CONFIG_FIELDS = {
    companyName: "SS_Company_Name",
    defaultReviewer: "SS_Default_Reviewer",
    accountingLink: "SS_Accounting_Software",
    payrollLink: "SS_Payroll_Provider"
};

/**
 * Cache for shared config values (populated once per session)
 */
PrairieForge._sharedConfigCache = null;

/**
 * Read a value from SS_PF_Config (shared config sheet).
 * Results are cached for performance.
 * 
 * @param {string} fieldName - The field name to look up
 * @returns {Promise<string>} - The config value or empty string
 */
PrairieForge.getSharedConfig = async function(fieldName) {
    // Return from cache if available
    if (PrairieForge._sharedConfigCache && PrairieForge._sharedConfigCache.has(fieldName)) {
        return PrairieForge._sharedConfigCache.get(fieldName) || "";
    }
    
    // Load cache if not yet loaded
    if (!PrairieForge._sharedConfigCache) {
        await PrairieForge.loadSharedConfig();
    }
    
    return PrairieForge._sharedConfigCache?.get(fieldName) || "";
};

/**
 * Load all shared config values from SS_PF_Config into cache.
 * Called automatically by getSharedConfig if cache is empty.
 * 
 * @returns {Promise<void>}
 */
PrairieForge.loadSharedConfig = async function() {
    PrairieForge._sharedConfigCache = new Map();
    
    if (typeof Excel === "undefined" || typeof Excel.run !== "function") {
        console.warn("Excel not available - shared config will use defaults");
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
            
            const range = configSheet.getUsedRangeOrNullObject();
            range.load("values");
            await context.sync();
            
            if (range.isNullObject) return;
            
            const values = range.values || [];
            if (!values.length) return;
            
            // Find Field and Value columns
            const headers = (values[0] || []).map(h => String(h || "").toLowerCase().trim());
            const fieldIdx = headers.findIndex(h => h === "field");
            const valueIdx = headers.findIndex(h => h === "value");
            
            if (fieldIdx === -1 || valueIdx === -1) {
                console.warn("SS_PF_Config missing Field/Value columns");
                return;
            }
            
            // Build cache
            values.slice(1).forEach(row => {
                const field = String(row[fieldIdx] || "").trim();
                const value = row[valueIdx];
                if (field) {
                    PrairieForge._sharedConfigCache.set(field, value != null ? String(value) : "");
                }
            });
            
            console.log(`Loaded ${PrairieForge._sharedConfigCache.size} shared config values`);
        });
    } catch (error) {
        console.error("Error loading shared config:", error);
    }
};

/**
 * Clear the shared config cache (forces reload on next access).
 */
PrairieForge.clearSharedConfigCache = function() {
    PrairieForge._sharedConfigCache = null;
};

/**
 * Get a config value with fallback logic.
 * Checks step-specific → module-specific → shared default.
 * 
 * @param {Object} options
 * @param {Function} options.getModuleConfig - Function to get module config value (sync)
 * @param {string} options.stepField - Step-specific field name
 * @param {string} options.moduleField - Module-level field name  
 * @param {string} options.sharedField - Shared config field name
 * @returns {Promise<string>} - The resolved value
 * 
 * @example
 * const reviewer = await PrairieForge.getConfigWithFallback({
 *     getModuleConfig: (field) => getConfigValue(field),
 *     stepField: "Reviewer_PTO_Headcount",
 *     moduleField: "PTO_Reviewer",
 *     sharedField: "Default_Reviewer"
 * });
 */
PrairieForge.getConfigWithFallback = async function({
    getModuleConfig,
    stepField,
    moduleField,
    sharedField
} = {}) {
    // Try step-specific first
    if (stepField && typeof getModuleConfig === "function") {
        const stepValue = getModuleConfig(stepField);
        if (stepValue) return stepValue;
    }
    
    // Try module-level next
    if (moduleField && typeof getModuleConfig === "function") {
        const moduleValue = getModuleConfig(moduleField);
        if (moduleValue) return moduleValue;
    }
    
    // Fall back to shared config
    if (sharedField) {
        return await PrairieForge.getSharedConfig(sharedField);
    }
    
    return "";
};

/**
 * Generate standard field names for a module based on naming convention.
 * 
 * @param {string} prefix - Module prefix (e.g., "PTO", "PR")
 * @param {Array<string>} steps - Step names (e.g., ["Config", "Import", "Headcount"])
 * @returns {Array<Object>} - Array of field definitions
 */
PrairieForge.generateModuleFields = function(prefix, steps = []) {
    const fields = [];
    
    // Module-level reviewer
    fields.push({
        category: "module",
        field: `${prefix}_Reviewer`,
        description: `Default reviewer for ${prefix} module`
    });
    
    // Step-specific fields
    steps.forEach(step => {
        fields.push(
            { category: "step", field: `${prefix}_${step}_Notes`, description: `${step} step notes` },
            { category: "step", field: `${prefix}_${step}_Reviewer`, description: `${step} step reviewer override` },
            { category: "step", field: `${prefix}_${step}_SignOff`, description: `${step} step sign-off date` },
            { category: "step", field: `${prefix}_${step}_Complete`, description: `${step} step completion flag` }
        );
    });
    
    return fields;
};

/**
 * Set a single config value in SS_PF_Config.
 * Creates the row if it doesn't exist, updates if it does.
 * 
 * @param {string} fieldName - The field name (e.g., "PTO_Reviewer")
 * @param {string} value - The value to set
 * @param {string} [category="module"] - Category for new rows
 * @returns {Promise<boolean>} - Success flag
 */
PrairieForge.setSharedConfig = async function(fieldName, value, category = "module") {
    if (typeof Excel === "undefined" || typeof Excel.run !== "function") {
        console.warn("Excel not available - cannot set shared config");
        return false;
    }
    
    try {
        await Excel.run(async (context) => {
            const configSheet = context.workbook.worksheets.getItemOrNullObject("SS_PF_Config");
            configSheet.load("isNullObject");
            await context.sync();
            
            if (configSheet.isNullObject) {
                // Create the sheet if it doesn't exist
                const newSheet = context.workbook.worksheets.add("SS_PF_Config");
                newSheet.getRange("A1:C1").values = [["Category", "Field", "Value"]];
                newSheet.getRange("A2:C2").values = [[category, fieldName, value]];
                await context.sync();
                
                // Update cache
                if (PrairieForge._sharedConfigCache) {
                    PrairieForge._sharedConfigCache.set(fieldName, value);
                }
                return;
            }
            
            // Sheet exists - find or add the row
            const usedRange = configSheet.getUsedRangeOrNullObject();
            usedRange.load("values, rowCount");
            await context.sync();
            
            if (usedRange.isNullObject) {
                // Empty sheet - add headers and first row
                configSheet.getRange("A1:C1").values = [["Category", "Field", "Value"]];
                configSheet.getRange("A2:C2").values = [[category, fieldName, value]];
                await context.sync();
            } else {
                const values = usedRange.values;
                let rowIndex = -1;
                
                // Find existing row
                for (let i = 1; i < values.length; i++) {
                    if (values[i][1] === fieldName) {
                        rowIndex = i;
                        break;
                    }
                }
                
                if (rowIndex >= 0) {
                    // Update existing row
                    configSheet.getRange(`C${rowIndex + 1}`).values = [[value]];
                } else {
                    // Add new row at the end
                    const newRow = values.length + 1;
                    configSheet.getRange(`A${newRow}:C${newRow}`).values = [[category, fieldName, value]];
                }
                await context.sync();
            }
            
            // Update cache
            if (PrairieForge._sharedConfigCache) {
                PrairieForge._sharedConfigCache.set(fieldName, value);
            }
        });
        
        return true;
    } catch (error) {
        console.error("Error setting shared config:", error);
        return false;
    }
};

/**
 * Write multiple config fields to SS_PF_Config at once.
 * Creates the sheet and headers if needed.
 * Only adds fields that don't already exist (won't overwrite existing values).
 * 
 * @param {Array<Object>} fields - Array of {category, field, value?} objects
 * @returns {Promise<{added: number, skipped: number}>} - Count of added/skipped fields
 */
PrairieForge.writeAllSharedConfig = async function(fields) {
    if (typeof Excel === "undefined" || typeof Excel.run !== "function") {
        console.warn("Excel not available - cannot write shared config");
        return { added: 0, skipped: 0 };
    }
    
    let added = 0;
    let skipped = 0;
    
    try {
        await Excel.run(async (context) => {
            let configSheet = context.workbook.worksheets.getItemOrNullObject("SS_PF_Config");
            configSheet.load("isNullObject");
            await context.sync();
            
            // Create sheet if it doesn't exist
            if (configSheet.isNullObject) {
                configSheet = context.workbook.worksheets.add("SS_PF_Config");
                configSheet.getRange("A1:C1").values = [["Category", "Field", "Value"]];
                await context.sync();
            }
            
            // Get existing fields
            const usedRange = configSheet.getUsedRangeOrNullObject();
            usedRange.load("values, rowCount");
            await context.sync();
            
            const existingFields = new Set();
            let nextRow = 2;
            
            if (!usedRange.isNullObject && usedRange.values) {
                const values = usedRange.values;
                nextRow = values.length + 1;
                
                // Build set of existing field names
                for (let i = 1; i < values.length; i++) {
                    if (values[i][1]) {
                        existingFields.add(values[i][1]);
                    }
                }
            }
            
            // Add new fields
            const newRows = [];
            for (const field of fields) {
                if (existingFields.has(field.field)) {
                    skipped++;
                } else {
                    newRows.push([field.category || "module", field.field, field.value || ""]);
                    added++;
                }
            }
            
            // Write all new rows at once
            if (newRows.length > 0) {
                const range = configSheet.getRange(`A${nextRow}:C${nextRow + newRows.length - 1}`);
                range.values = newRows;
                await context.sync();
            }
            
            // Refresh cache
            await PrairieForge.loadSharedConfig();
        });
        
        return { added, skipped };
    } catch (error) {
        console.error("Error writing shared config:", error);
        return { added: 0, skipped: 0, error: error.message };
    }
};

/**
 * Show ALL sheets in the workbook (emergency recovery / debugging)
 * Makes every sheet visible, including veryHidden ones
 */
PrairieForge.showAllSheets = async function() {
    if (typeof Excel === 'undefined') {
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
                if (sheet.visibility !== Excel.SheetVisibility.visible) {
                    sheet.visibility = Excel.SheetVisibility.visible;
                    console.log(`[ShowAll] Made visible: ${sheet.name}`);
                    unhiddenCount++;
                }
            });
            
            await context.sync();
            console.log(`[ShowAll] Done! Made ${unhiddenCount} sheets visible. Total: ${worksheets.items.length}`);
        });
    } catch (error) {
        console.error("[ShowAll] Error:", error);
    }
};

/**
 * Unhide system sheets (SS_* prefix)
 */
PrairieForge.unhideSystemSheets = async function() {
    if (typeof Excel === 'undefined') {
        console.log("Excel not available");
        return;
    }
    
    try {
        await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            worksheets.load("items/name,visibility");
            await context.sync();
            
            worksheets.items.forEach((sheet) => {
                if (sheet.name.toUpperCase().startsWith("SS_")) {
                    sheet.visibility = Excel.SheetVisibility.visible;
                    console.log(`[Unhide] Made visible: ${sheet.name}`);
                }
            });
            
            await context.sync();
            console.log("[Unhide] System sheets are now visible!");
        });
    } catch (error) {
        console.error("[Unhide] Error:", error);
    }
};

console.log('Prairie Forge Common Utilities loaded');
