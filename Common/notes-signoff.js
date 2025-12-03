import { LOCK_CLOSED_SVG, LOCK_OPEN_SVG, CHECK_ICON_SVG, SAVE_ICON_SVG } from "./icons.js";

/**
 * Escape HTML special characters
 * @param {string} str - String to escape
 * @returns {string} Escaped string
 */
function escapeHtml(str) {
    if (str == null) return "";
    return String(str)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

/**
 * Render a labeled button (icon + label below)
 * @param {string} buttonHtml - The button HTML
 * @param {string} label - Label text to display below the button
 * @returns {string} HTML string wrapped with label
 */
export function renderLabeledButton(buttonHtml, label) {
    return `
        <div class="pf-labeled-btn">
            ${buttonHtml}
            <span class="pf-btn-label">${label}</span>
        </div>
    `;
}

/**
 * Render the notes section with save and lock buttons
 * @param {Object} options - Rendering options
 * @param {string} options.textareaId - ID for the textarea element
 * @param {string} options.value - Current notes value
 * @param {string} options.permanentId - ID for the lock button (optional)
 * @param {boolean} options.isPermanent - Whether notes are locked
 * @param {string} options.hintId - ID for hint element (optional)
 * @param {string} options.saveButtonId - ID for save button (optional)
 * @param {boolean} options.isSaved - Whether notes are currently saved (optional)
 * @param {string} options.placeholder - Placeholder text (optional)
 * @returns {string} HTML string
 */
export function renderInlineNotes({
    textareaId,
    value,
    permanentId,
    isPermanent,
    hintId,
    saveButtonId,
    isSaved = false,
    placeholder = "Enter notes here..."
}) {
    const lockIcon = isPermanent ? LOCK_CLOSED_SVG : LOCK_OPEN_SVG;

    const saveButton = saveButtonId
        ? `<button type="button" class="pf-action-toggle pf-save-btn ${isSaved ? "is-saved" : ""}" id="${saveButtonId}" data-save-input="${textareaId}" title="Save notes">${SAVE_ICON_SVG}</button>`
        : "";

    const lockButton = permanentId
        ? `<button type="button" class="pf-action-toggle pf-notes-lock ${isPermanent ? "is-locked" : ""}" id="${permanentId}" aria-pressed="${isPermanent}" title="Lock notes (retain after archive)">${lockIcon}</button>`
        : "";

    return `
        <article class="pf-step-card pf-step-detail pf-notes-card">
            <div class="pf-notes-header">
                <div>
                    <h3 class="pf-notes-title">Notes</h3>
                    <p class="pf-notes-subtext">Leave notes your future self will appreciate. Notes clear after archiving. Click lock to retain permanently.</p>
                </div>
            </div>
            <div class="pf-notes-body">
                <textarea id="${textareaId}" rows="6" placeholder="${escapeHtml(placeholder)}">${escapeHtml(value || "")}</textarea>
                ${hintId ? `<p class="pf-signoff-hint" id="${hintId}"></p>` : ""}
            </div>
            <div class="pf-notes-action">
                ${permanentId ? renderLabeledButton(lockButton, "Lock") : ""}
                ${saveButtonId ? renderLabeledButton(saveButton, "Save") : ""}
            </div>
        </article>
    `;
}

/**
 * Render the sign-off section with save and complete buttons
 * @param {Object} options - Rendering options
 * @param {string} options.reviewerInputId - ID for reviewer name input
 * @param {string} options.reviewerValue - Current reviewer name
 * @param {string} options.signoffInputId - ID for sign-off date input
 * @param {string} options.signoffValue - Current sign-off date
 * @param {boolean} options.isComplete - Whether step is marked complete
 * @param {string} options.saveButtonId - ID for save button
 * @param {boolean} options.isSaved - Whether sign-off is currently saved (optional)
 * @param {string} options.completeButtonId - ID for complete toggle button
 * @param {string} options.subtext - Subtext description (optional)
 * @returns {string} HTML string
 */
export function renderSignoff({
    reviewerInputId,
    reviewerValue,
    signoffInputId,
    signoffValue,
    isComplete,
    saveButtonId, // retained for compatibility; no separate save button rendered
    isSaved = false, // retained for compatibility
    completeButtonId,
    subtext = "Sign-off below. Click checkmark icon. Done."
}) {
    const completeButton = `<button type="button" class="pf-action-toggle ${isComplete ? "is-active" : ""}" id="${completeButtonId}" aria-pressed="${Boolean(isComplete)}" title="Mark step complete">${CHECK_ICON_SVG}</button>`;

    return `
        <article class="pf-step-card pf-step-detail pf-config-card">
            <div class="pf-config-head pf-notes-header">
                <div>
                    <h3>Sign-off</h3>
                    <p class="pf-config-subtext">${escapeHtml(subtext)}</p>
                </div>
            </div>
            <div class="pf-config-grid">
                <label class="pf-config-field">
                    <span>Reviewer Name</span>
                    <input type="text" id="${reviewerInputId}" value="${escapeHtml(reviewerValue)}" placeholder="Full name">
                </label>
                <label class="pf-config-field">
                    <span>Sign-off Date</span>
                    <input type="date" id="${signoffInputId}" value="${escapeHtml(signoffValue)}" readonly>
                </label>
            </div>
            <div class="pf-signoff-action">
                ${renderLabeledButton(completeButton, "Done")}
            </div>
        </article>
    `;
}

/**
 * Render combined notes and sign-off sections
 * @param {Object} options - Combined options for both sections
 * @returns {string} HTML string
 */
export function renderNotesAndSignoff({
    // Notes options
    textareaId,
    notesValue,
    permanentId,
    isPermanent,
    hintId,
    notesSaveButtonId,
    placeholder,
    // Sign-off options
    reviewerInputId,
    reviewerValue,
    signoffInputId,
    signoffValue,
    isComplete,
    signoffSaveButtonId,
    completeButtonId,
    subtext
}) {
    return `
        ${renderInlineNotes({
            textareaId,
            value: notesValue,
            permanentId,
            isPermanent,
            hintId,
            saveButtonId: notesSaveButtonId,
            placeholder
        })}
        ${renderSignoff({
            reviewerInputId,
            reviewerValue,
            signoffInputId,
            signoffValue,
            isComplete,
            saveButtonId: signoffSaveButtonId,
            completeButtonId,
            subtext
        })}
    `;
}

/**
 * Update lock button visual state
 * @param {HTMLElement} button - Lock button element
 * @param {boolean} isLocked - Whether locked
 */
export function updateLockButtonVisual(button, isLocked) {
    if (!button) return;
    button.classList.toggle("is-locked", isLocked);
    button.setAttribute("aria-pressed", String(isLocked));
    button.innerHTML = isLocked ? LOCK_CLOSED_SVG : LOCK_OPEN_SVG;
}

/**
 * Update action toggle button state
 * @param {HTMLElement} button - Toggle button element
 * @param {boolean} isActive - Whether active
 */
export function updateActionToggleState(button, isActive) {
    if (!button) return;
    button.classList.toggle("is-active", isActive);
    button.setAttribute("aria-pressed", String(isActive));
}

/**
 * Update save button visual state
 * @param {HTMLElement} button - Save button element
 * @param {boolean} isSaved - Whether content is saved
 */
export function updateSaveButtonState(button, isSaved) {
    if (!button) return;
    button.classList.toggle("is-saved", isSaved);
}

/**
 * Initialize save state tracking for all save buttons in a container.
 * Finds all buttons with class "pf-save-btn" and data-save-input attribute,
 * then auto-wires input change tracking to toggle the saved state.
 *
 * @param {HTMLElement|Document} container - Container to search for save buttons (defaults to document)
 * @returns {Function} Cleanup function to remove all event listeners
 */
export function initSaveTracking(container = document) {
    const saveButtons = container.querySelectorAll(".pf-save-btn[data-save-input]");
    const cleanupFns = [];

    saveButtons.forEach((saveButton) => {
        const inputId = saveButton.getAttribute("data-save-input");
        const input = document.getElementById(inputId);
        if (!input) return;

        const handleChange = () => {
            updateSaveButtonState(saveButton, false);
        };

        input.addEventListener("input", handleChange);
        cleanupFns.push(() => input.removeEventListener("input", handleChange));
    });

    return () => cleanupFns.forEach((fn) => fn());
}
