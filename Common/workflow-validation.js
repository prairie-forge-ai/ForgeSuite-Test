/**
 * Workflow Validation Utilities
 * 
 * Ensures steps are completed sequentially.
 * Users can VIEW any step, but cannot COMPLETE until previous steps are signed off.
 */

/**
 * Check if a step can be completed based on previous step completion status
 * @param {number} stepId - The step ID to check (0-based)
 * @param {Object} completionStatus - Object mapping step IDs to completion status
 * @returns {{ canComplete: boolean, blockedBy: number|null, message: string }}
 */
export function canCompleteStep(stepId, completionStatus) {
    // Step 0 can always be completed (no prerequisites)
    if (stepId === 0) {
        return { canComplete: true, blockedBy: null, message: "" };
    }
    
    // Check all previous steps
    for (let i = 0; i < stepId; i++) {
        if (!completionStatus[i]) {
            return {
                canComplete: false,
                blockedBy: i,
                message: `Complete Step ${i} before signing off on this step.`
            };
        }
    }
    
    return { canComplete: true, blockedBy: null, message: "" };
}

/**
 * Get completion status for all steps from config state
 * @param {Object} configState - The module's config state object
 * @param {Object} stepFields - Mapping of step IDs to their config field names
 * @returns {Object} Map of step IDs to boolean completion status
 */
export function getStepCompletionStatus(configState, stepFields) {
    const status = {};
    
    Object.keys(stepFields).forEach(stepId => {
        const id = parseInt(stepId, 10);
        // A step is complete if it has a sign-off date OR is explicitly marked complete
        const hasSignOff = Boolean(configState.signoffs?.[id] || configState.values?.[stepFields[id]?.signOff]);
        const isMarkedComplete = Boolean(configState.completes?.[id]);
        status[id] = hasSignOff || isMarkedComplete;
    });
    
    return status;
}

/**
 * Handle sign-off attempt with sequential validation
 * @param {number} stepId - The step being signed off
 * @param {Object} completionStatus - Current completion status of all steps
 * @param {Function} onAllowed - Callback if sign-off is allowed
 * @param {Function} onBlocked - Callback if sign-off is blocked (receives blockedBy step ID and message)
 */
export function handleSignoffAttempt(stepId, completionStatus, onAllowed, onBlocked) {
    const { canComplete, blockedBy, message } = canCompleteStep(stepId, completionStatus);
    
    if (canComplete) {
        onAllowed();
    } else {
        onBlocked(blockedBy, message);
    }
}

/**
 * Create a visual blocker overlay for sign-off section
 * @param {string} message - Message to display
 * @param {number} blockedByStep - The step number that needs completion
 * @returns {string} HTML string for the blocker overlay
 */
export function renderSignoffBlocker(message, blockedByStep) {
    return `
        <div class="pf-signoff-blocker">
            <div class="pf-signoff-blocker-content">
                <span class="pf-signoff-blocker-icon">🔒</span>
                <p class="pf-signoff-blocker-message">${message}</p>
                <button type="button" class="pf-signoff-blocker-btn" data-goto-step="${blockedByStep}">
                    Go to Step ${blockedByStep}
                </button>
            </div>
        </div>
    `;
}

/**
 * Show a toast notification when sign-off is blocked
 * @param {string} message - Message to display
 */
export function showBlockedToast(message) {
    // Remove any existing toast
    const existing = document.querySelector('.pf-workflow-toast');
    if (existing) existing.remove();
    
    const toast = document.createElement('div');
    toast.className = 'pf-workflow-toast pf-workflow-toast--warning';
    toast.innerHTML = `
        <span class="pf-workflow-toast-icon">⚠️</span>
        <span class="pf-workflow-toast-message">${message}</span>
    `;
    
    document.body.appendChild(toast);
    
    // Trigger animation
    requestAnimationFrame(() => {
        toast.classList.add('pf-workflow-toast--visible');
    });
    
    // Auto-remove after 4 seconds
    setTimeout(() => {
        toast.classList.remove('pf-workflow-toast--visible');
        setTimeout(() => toast.remove(), 300);
    }, 4000);
}

/**
 * Bind workflow validation to sign-off buttons
 * @param {HTMLElement} container - Container element
 * @param {number} stepId - Current step ID
 * @param {Function} getCompletionStatus - Function that returns current completion status
 * @param {Function} onSignoff - Callback when sign-off is allowed and clicked
 * @param {Function} navigateToStep - Function to navigate to a specific step
 */
export function bindWorkflowValidation(container, stepId, getCompletionStatus, onSignoff, navigateToStep) {
    // Find sign-off toggle buttons (they have IDs like "step-signoff-toggle-X" or similar patterns)
    const signoffToggles = container.querySelectorAll('[id*="signoff-toggle"], [id*="step-signoff-toggle"]');
    
    signoffToggles.forEach(toggle => {
        const originalClickHandler = toggle.onclick;
        
        toggle.addEventListener('click', (e) => {
            const completionStatus = getCompletionStatus();
            const { canComplete, blockedBy, message } = canCompleteStep(stepId, completionStatus);
            
            if (!canComplete) {
                e.preventDefault();
                e.stopPropagation();
                showBlockedToast(message);
                
                // Optionally highlight the blocked step
                if (navigateToStep && blockedBy !== null) {
                    // Could add a "Go to Step X" button in the toast
                }
                return false;
            }
            
            // If allowed, proceed with sign-off
            if (onSignoff) {
                onSignoff();
            }
        });
    });
}

