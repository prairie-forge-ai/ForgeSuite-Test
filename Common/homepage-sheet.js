/**
 * Homepage Sheet Utilities
 * Creates and manages module landing pages with clean black backgrounds
 */

import { BRANDING } from "./constants.js";

// Ada Assistant Configuration - imported from constants
const ADA_IMAGE_URL = BRANDING.ADA_IMAGE_URL;

/**
 * Creates or activates a module homepage sheet with formatted styling
 * @param {string} sheetName - Name of the homepage sheet (e.g., "PR_Homepage")
 * @param {string} title - Module title to display (e.g., "Payroll Recorder")
 * @param {string} subtitle - Description/subtext for the module
 */
export async function activateHomepageSheet(sheetName, title, subtitle) {
    if (typeof Excel === "undefined") {
        console.warn("Excel runtime not available for homepage sheet");
        return;
    }
    
    try {
        await Excel.run(async (context) => {
            // Check if sheet exists
            const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
            sheet.load("isNullObject, name");
            await context.sync();
            
            let targetSheet;
            
            if (sheet.isNullObject) {
                // Create the sheet
                targetSheet = context.workbook.worksheets.add(sheetName);
                await context.sync();
                
                // Set up the homepage content
                await setupHomepageContent(context, targetSheet, title, subtitle);
            } else {
                targetSheet = sheet;
                // Update content in case title/subtitle changed
                await setupHomepageContent(context, targetSheet, title, subtitle);
            }
            
            // Activate the sheet
            targetSheet.activate();
            targetSheet.getRange("A1").select();
            await context.sync();
        });
    } catch (error) {
        console.error(`Error activating homepage sheet ${sheetName}:`, error);
    }
}

/**
 * Sets up the homepage sheet content and formatting
 */
async function setupHomepageContent(context, sheet, title, subtitle) {
    // Clear existing content
    try {
        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load("isNullObject");
        await context.sync();
        if (!usedRange.isNullObject) {
            usedRange.clear();
            await context.sync();
        }
    } catch (e) {
        // Ignore clear errors
    }
    
    // Hide gridlines
    sheet.showGridlines = false;
    
    // Set up column widths for a clean look
    sheet.getRange("A:A").format.columnWidth = 400; // Content area
    sheet.getRange("B:B").format.columnWidth = 50;  // Right padding
    
    // Set row heights
    sheet.getRange("1:1").format.rowHeight = 60;  // Title row
    sheet.getRange("2:2").format.rowHeight = 30;  // Subtitle row
    
    // Write content starting at A1
    const data = [
        [title, ""],
        [subtitle, ""],
        ["", ""],
        ["", ""]
    ];
    
    const dataRange = sheet.getRangeByIndexes(0, 0, 4, 2);
    dataRange.values = data;
    
    // Apply black background to entire visible area
    const backgroundRange = sheet.getRange("A1:Z100");
    backgroundRange.format.fill.color = "#0f0f0f";
    
    // Format title (A1)
    const titleCell = sheet.getRange("A1");
    titleCell.format.font.bold = true;
    titleCell.format.font.size = 36;
    titleCell.format.font.color = "#ffffff";
    titleCell.format.font.name = "Segoe UI Light";
    titleCell.format.verticalAlignment = "Center";
    
    // Format subtitle (A2)
    const subtitleCell = sheet.getRange("A2");
    subtitleCell.format.font.size = 14;
    subtitleCell.format.font.color = "#a0a0a0";
    subtitleCell.format.font.name = "Segoe UI";
    subtitleCell.format.verticalAlignment = "Top";
    
    // Freeze panes to prevent scrolling issues
    sheet.freezePanes.freezeRows(0);
    sheet.freezePanes.freezeColumns(0);
    
    await context.sync();
}

/**
 * Homepage configurations for each module
 */
export const HOMEPAGE_CONFIG = {
    "module-selector": {
        sheetName: "SS_Homepage",
        title: "ForgeSuite",
        subtitle: "Select a module from the side panel to get started."
    },
    "payroll-recorder": {
        sheetName: "PR_Homepage",
        title: "Payroll Recorder",
        subtitle: "Normalize payroll exports, enforce controls, and prep journal entries without leaving Excel."
    },
    "pto-accrual": {
        sheetName: "PTO_Homepage",
        title: "PTO Accrual",
        subtitle: "Calculate employee PTO liabilities, compare period-over-period changes, and prepare accrual journal entries."
    }
};

/**
 * Get homepage config by module key
 */
export function getHomepageConfig(moduleKey) {
    return HOMEPAGE_CONFIG[moduleKey] || HOMEPAGE_CONFIG["module-selector"];
}

/**
 * Renders the floating Ada assistant button
 * Call this when showing the home view
 */
export function renderAdaFab() {
    // Remove existing FAB if present
    removeAdaFab();
    
    const fab = document.createElement("button");
    fab.className = "pf-ada-fab";
    fab.id = "pf-ada-fab";
    fab.setAttribute("aria-label", "Ask Ada");
    fab.setAttribute("title", "Ask Ada");
    fab.innerHTML = `
        <span class="pf-ada-fab__ring"></span>
        <img 
            class="pf-ada-fab__image" 
            src="${ADA_IMAGE_URL}" 
            alt="Ada - Your AI Assistant"
            onerror="this.style.display='none'"
        />
    `;
    
    document.body.appendChild(fab);
    
    // Bind click event
    fab.addEventListener("click", showAdaModal);
    
    return fab;
}

/**
 * Removes the Ada FAB from the DOM
 * Call this when navigating away from home view
 */
export function removeAdaFab() {
    const existingFab = document.getElementById("pf-ada-fab");
    if (existingFab) {
        existingFab.remove();
    }
    
    // Also remove modal if open
    const existingModal = document.getElementById("pf-ada-modal-overlay");
    if (existingModal) {
        existingModal.remove();
    }
}

/**
 * Shows the Ada assistant modal
 */
export function showAdaModal() {
    // Remove existing modal if present
    const existingModal = document.getElementById("pf-ada-modal-overlay");
    if (existingModal) {
        existingModal.remove();
    }
    
    const overlay = document.createElement("div");
    overlay.className = "pf-ada-modal-overlay";
    overlay.id = "pf-ada-modal-overlay";
    
    overlay.innerHTML = `
        <div class="pf-ada-modal">
            <div class="pf-ada-modal__header">
                <button class="pf-ada-modal__close" id="ada-modal-close" aria-label="Close">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <line x1="18" y1="6" x2="6" y2="18"></line>
                        <line x1="6" y1="6" x2="18" y2="18"></line>
                    </svg>
                </button>
                <img class="pf-ada-modal__avatar" src="${ADA_IMAGE_URL}" alt="Ada" />
                <h2 class="pf-ada-modal__title"><span style="font-weight:400;">ask</span><span style="font-weight:700;">ADA</span></h2>
                <p class="pf-ada-modal__subtitle">Your AI-powered assistant</p>
            </div>
            <div class="pf-ada-modal__body">
                <div class="pf-ada-modal__features">
                    <div class="pf-ada-modal__feature">
                        <div class="pf-ada-modal__feature-icon">💬</div>
                        <span class="pf-ada-modal__feature-text">Ask questions about your data</span>
                    </div>
                    <div class="pf-ada-modal__feature">
                        <div class="pf-ada-modal__feature-icon">📊</div>
                        <span class="pf-ada-modal__feature-text">Get insights and trend analysis</span>
                    </div>
                    <div class="pf-ada-modal__feature">
                        <div class="pf-ada-modal__feature-icon">🔍</div>
                        <span class="pf-ada-modal__feature-text">Troubleshoot issues quickly</span>
                    </div>
                </div>
            </div>
            <div class="pf-ada-modal__footer">
                <span class="pf-ada-modal__powered-by">Powered by ChatGPT</span>
            </div>
        </div>
    `;
    
    document.body.appendChild(overlay);
    
    // Trigger animation
    requestAnimationFrame(() => {
        overlay.classList.add("is-visible");
    });
    
    // Bind close events
    const closeBtn = document.getElementById("ada-modal-close");
    closeBtn?.addEventListener("click", hideAdaModal);
    
    // Close on overlay click
    overlay.addEventListener("click", (e) => {
        if (e.target === overlay) {
            hideAdaModal();
        }
    });
    
    // Close on Escape key
    const handleEscape = (e) => {
        if (e.key === "Escape") {
            hideAdaModal();
            document.removeEventListener("keydown", handleEscape);
        }
    };
    document.addEventListener("keydown", handleEscape);
}

/**
 * Hides the Ada assistant modal
 */
export function hideAdaModal() {
    const overlay = document.getElementById("pf-ada-modal-overlay");
    if (overlay) {
        overlay.classList.remove("is-visible");
        setTimeout(() => {
            overlay.remove();
        }, 300);
    }
}

