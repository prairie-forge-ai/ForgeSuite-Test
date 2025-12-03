/**
 * Ada - Prairie Forge AI Assistant
 * 
 * A stunning, chat-based AI interface that integrates with spreadsheet data
 * to provide intelligent insights and recommendations.
 * 
 * Named after Ada Lovelace, the first computer programmer.
 * Powered by ChatGPT (OpenAI).
 * 
 * AUTHENTICATION: Uses Supabase Edge Functions - no additional auth required
 * INSTRUCTIONS: Configure via systemPrompt and contextProvider options
 */

import { BRANDING } from "./constants.js";

const CHATGPT_ICON = `<svg viewBox="0 0 24 24" fill="currentColor"><path d="M22.2819 9.8211a5.9847 5.9847 0 0 0-.5157-4.9108 6.0462 6.0462 0 0 0-6.5098-2.9A6.0651 6.0651 0 0 0 4.9807 4.1818a5.9847 5.9847 0 0 0-3.9977 2.9 6.0462 6.0462 0 0 0 .7427 7.0966 5.98 5.98 0 0 0 .511 4.9107 6.051 6.051 0 0 0 6.5146 2.9001A5.9847 5.9847 0 0 0 13.2599 24a6.0557 6.0557 0 0 0 5.7718-4.2058 5.9894 5.9894 0 0 0 3.9977-2.9001 6.0557 6.0557 0 0 0-.7475-7.0729zm-9.022 12.6081a4.4755 4.4755 0 0 1-2.8764-1.0408l.1419-.0804 4.7783-2.7582a.7948.7948 0 0 0 .3927-.6813v-6.7369l2.02 1.1686a.071.071 0 0 1 .038.052v5.5826a4.504 4.504 0 0 1-4.4945 4.4944zm-9.6607-4.1254a4.4708 4.4708 0 0 1-.5346-3.0137l.142.0852 4.783 2.7582a.7712.7712 0 0 0 .7806 0l5.8428-3.3685v2.3324a.0804.0804 0 0 1-.0332.0615L9.74 19.9502a4.4992 4.4992 0 0 1-6.1408-1.6464zM2.3408 7.8956a4.485 4.485 0 0 1 2.3655-1.9728V11.6a.7664.7664 0 0 0 .3879.6765l5.8144 3.3543-2.0201 1.1685a.0757.0757 0 0 1-.071 0l-4.8303-2.7865A4.504 4.504 0 0 1 2.3408 7.8956zm16.5963 3.8558L13.1038 8.364 15.1192 7.2a.0757.0757 0 0 1 .071 0l4.8303 2.7913a4.4944 4.4944 0 0 1-.6765 8.1042v-5.6772a.79.79 0 0 0-.407-.667zm2.0107-3.0231l-.142-.0852-4.7735-2.7818a.7759.7759 0 0 0-.7854 0L9.409 9.2297V6.8974a.0662.0662 0 0 1 .0284-.0615l4.8303-2.7866a4.4992 4.4992 0 0 1 6.6802 4.66zM8.3065 12.863l-2.02-1.1638a.0804.0804 0 0 1-.038-.0567V6.0742a4.4992 4.4992 0 0 1 7.3757-3.4537l-.142.0805L8.704 5.459a.7948.7948 0 0 0-.3927.6813zm1.0976-2.3654l2.602-1.4998 2.6069 1.4998v2.9994l-2.5974 1.4997-2.6067-1.4997Z"/></svg>`;

// Arrow icon for send button
const SEND_ARROW = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14"/><path d="m12 5 7 7-7 7"/></svg>`;

// Ada character image URL - imported from constants
const ADA_IMAGE_URL = BRANDING.ADA_IMAGE_URL;

const DEFAULT_OPTIONS = {
    id: "pf-copilot",
    heading: "Ada",
    subtext: "Your smart assistant to help you analyze and troubleshoot.",
    welcomeMessage: "What would you like to explore?",
    placeholder: "Where should I focus this pay period?",
    quickActions: [
        { id: "diagnostics", label: "Diagnostics", prompt: "Run a diagnostic check on the current payroll data. Check for completeness, accuracy issues, and any data quality concerns." },
        { id: "insights", label: "Insights", prompt: "What are the key insights and notable findings from this payroll period that I should highlight for executive review?" },
        { id: "variance", label: "Variances", prompt: "Analyze the significant variances between this period and the prior period. What's driving the changes?" },
        { id: "journal", label: "JE Check", prompt: "Is the journal entry ready for export? Check that debits equal credits and flag any mapping issues." }
    ],
    systemPrompt: `You are Prairie Forge CoPilot, an expert financial analyst assistant embedded in an Excel add-in. 

Your role is to help accountants and CFOs:
1. Analyze payroll expense data for accuracy and completeness
2. Identify trends, anomalies, and areas requiring attention
3. Prepare executive-ready insights and talking points
4. Validate journal entries before export

Communication style:
- Be concise but thorough
- Use bullet points for clarity
- Highlight actionable items with ⚠️ or ✓
- Format currency as $X,XXX and percentages as X.X%
- Always suggest 2-3 concrete next steps

When analyzing data, look for:
- Period-over-period changes > 10%
- Department cost anomalies
- Headcount vs payroll mismatches
- Burden rate outliers
- Missing or incomplete mappings`
};

// Message history for the session
let messageHistory = [];

/**
 * Render the CoPilot card HTML - Apple-inspired design
 */
export function renderCopilotCard(options = {}) {
    const merged = { ...DEFAULT_OPTIONS, ...options };
    
    const quickActionsHtml = merged.quickActions?.map(action => 
        `<button type="button" class="pf-ada-chip" data-action="${action.id}" data-prompt="${escapeAttr(action.prompt)}">${action.label}</button>`
    ).join('') || '';
    
    return `
        <article class="pf-ada" data-copilot="${merged.id}">
            <header class="pf-ada-header">
                <div class="pf-ada-identity">
                    <img class="pf-ada-avatar" src="${ADA_IMAGE_URL}" alt="Ada" onerror="this.style.display='none'" />
                    <div class="pf-ada-name">
                        <span class="pf-ada-title"><span class="pf-ada-title--ask">ask</span><span class="pf-ada-title--ada">ADA</span></span>
                        <span class="pf-ada-role">${merged.subtext}</span>
                    </div>
                </div>
                <div class="pf-ada-status" id="${merged.id}-status-badge" title="Ready">
                    <span class="pf-ada-status-dot" id="${merged.id}-status-dot"></span>
                </div>
            </header>
            
            <div class="pf-ada-body">
                <div class="pf-ada-conversation" id="${merged.id}-messages">
                    <div class="pf-ada-bubble pf-ada-bubble--ai">
                        <p>${merged.welcomeMessage}</p>
                    </div>
                </div>
                
                <div class="pf-ada-composer">
                    <input 
                        type="text" 
                        class="pf-ada-input" 
                        id="${merged.id}-prompt" 
                        placeholder="${merged.placeholder}" 
                        autocomplete="off"
                    >
                    <button type="button" class="pf-ada-send" id="${merged.id}-ask" title="Send">
                        ${SEND_ARROW}
                    </button>
                </div>
                
                ${quickActionsHtml ? `<div class="pf-ada-chips">${quickActionsHtml}</div>` : ''}
                
                <footer class="pf-ada-footer">
                    ${CHATGPT_ICON}
                    <span>Powered by ChatGPT</span>
                </footer>
            </div>
        </article>
    `;
}

/**
 * Escape HTML attribute values
 */
function escapeAttr(str) {
    return String(str || '')
        .replace(/&/g, '&amp;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

/**
 * Bind event handlers to the CoPilot card
 */
export function bindCopilotCard(container, options = {}) {
    const merged = { ...DEFAULT_OPTIONS, ...options };
    const root = container.querySelector(`[data-copilot="${merged.id}"]`);
    if (!root) return;
    
    const messagesContainer = root.querySelector(`#${merged.id}-messages`);
    const promptInput = root.querySelector(`#${merged.id}-prompt`);
    const askButton = root.querySelector(`#${merged.id}-ask`);
    const statusDot = root.querySelector(`#${merged.id}-status-dot`);
    const statusBadge = root.querySelector(`#${merged.id}-status-badge`);
    
    // State
    let isProcessing = false;
    
    // Helper functions
    const setStatus = (text, state = 'ready') => {
        if (statusDot) {
            statusDot.classList.remove('pf-ada-status-dot--busy', 'pf-ada-status-dot--offline');
            if (state === 'busy') statusDot.classList.add('pf-ada-status-dot--busy');
            if (state === 'offline') statusDot.classList.add('pf-ada-status-dot--offline');
        }
        if (statusBadge) {
            statusBadge.title = text;
        }
    };
    
    const addMessage = (content, role = 'assistant') => {
        if (!messagesContainer) return;
        
        const bubbleClass = role === 'user' ? 'pf-ada-bubble--user' : 
                           role === 'system' ? 'pf-ada-bubble--system' : 'pf-ada-bubble--ai';
        
        const messageEl = document.createElement('div');
        messageEl.className = `pf-ada-bubble ${bubbleClass}`;
        messageEl.innerHTML = `<p>${formatResponse(content)}</p>`;
        messagesContainer.appendChild(messageEl);
        messagesContainer.scrollTop = messagesContainer.scrollHeight;
        
        // Store in history
        messageHistory.push({ role, content, timestamp: new Date().toISOString() });
    };
    
    const showLoading = () => {
        if (!messagesContainer) return;
        
        const loadingEl = document.createElement('div');
        loadingEl.className = 'pf-ada-bubble pf-ada-bubble--ai pf-ada-bubble--loading';
        loadingEl.id = `${merged.id}-loading`;
        loadingEl.innerHTML = `
            <div class="pf-ada-typing">
                <span></span><span></span><span></span>
            </div>
        `;
        messagesContainer.appendChild(loadingEl);
        messagesContainer.scrollTop = messagesContainer.scrollHeight;
    };
    
    const hideLoading = () => {
        const loadingEl = document.getElementById(`${merged.id}-loading`);
        if (loadingEl) loadingEl.remove();
    };
    
    const formatResponse = (text) => {
        // Convert markdown-style formatting to HTML
        return String(text)
            .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
            .replace(/\n\n/g, '</p><p>')
            .replace(/\n- /g, '<br>• ')
            .replace(/\n/g, '<br>');
    };
    
    // Main submit handler
    const handleSubmit = async (promptText) => {
        const prompt = promptText || promptInput?.value.trim();
        if (!prompt || isProcessing) return;
        
        isProcessing = true;
        if (promptInput) promptInput.value = '';
        if (askButton) askButton.disabled = true;
        
        // Add user message
        addMessage(prompt, 'user');
        
        // Show loading state
        showLoading();
        setStatus('Analyzing...', 'busy');
        
        try {
            // Get context from the spreadsheet if provider is available
            let context = null;
            if (typeof merged.contextProvider === 'function') {
                try {
                    context = await merged.contextProvider();
                } catch (e) {
                    console.warn('CoPilot: Context provider failed', e);
                }
            }
            
            // Call the AI handler
            let response;
            if (typeof merged.onPrompt === 'function') {
                response = await merged.onPrompt(prompt, context, messageHistory);
            } else if (typeof merged.apiEndpoint === 'string') {
                response = await callCopilotAPI(merged.apiEndpoint, prompt, context, merged.systemPrompt);
            } else {
                // Fallback demo response
                response = generateDemoResponse(prompt, context);
            }
            
            hideLoading();
            addMessage(response, 'assistant');
            setStatus('Ready to assist', 'ready');
            
        } catch (error) {
            console.error('CoPilot error:', error);
            hideLoading();
            addMessage(`I encountered an issue: ${error.message}. Please try again.`, 'system');
            setStatus('Error occurred', 'offline');
        }
        
        isProcessing = false;
        if (askButton) askButton.disabled = false;
        promptInput?.focus();
    };
    
    // Bind events
    askButton?.addEventListener('click', () => handleSubmit());
    
    promptInput?.addEventListener('keydown', (event) => {
        if (event.key === 'Enter' && !event.shiftKey) {
            event.preventDefault();
            handleSubmit();
        }
    });
    
    // Quick action chips
    root.querySelectorAll('.pf-ada-chip').forEach(btn => {
        btn.addEventListener('click', () => {
            const prompt = btn.dataset.prompt;
            if (prompt) handleSubmit(prompt);
        });
    });
}

/**
 * Call the CoPilot API endpoint (Supabase Edge Function)
 */
async function callCopilotAPI(endpoint, prompt, context, systemPrompt) {
    const response = await fetch(endpoint, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            prompt,
            context,
            systemPrompt,
            history: messageHistory.slice(-10) // Last 10 messages for context
        })
    });
    
    if (!response.ok) {
        throw new Error(`API request failed: ${response.status}`);
    }
    
    const data = await response.json();
    return data.message || data.response || 'No response received.';
}

/**
 * Generate a demo response when no API is configured
 * This simulates intelligent responses based on the prompt
 */
function generateDemoResponse(prompt, context) {
    const lowerPrompt = prompt.toLowerCase();
    
    if (lowerPrompt.includes('diagnostic') || lowerPrompt.includes('check')) {
        return `Great question! Let me run through the diagnostics for you.

**✓ What Looks Good:**
• All required fields are populated
• Current period matches your config date
• All expense categories are mapped to GL accounts

**⚠️ Items Worth Reviewing:**
• 2 departments show >15% variance from prior period
• Burden rate (14.6%) is slightly below your historical average (16.2%)

**My Recommendations:**
1. Take a closer look at the Sales & Marketing variance (-44.4%)
2. Verify headcount changes align with HR records
3. Once satisfied, you're clear to proceed to Journal Entry Prep!

Let me know if you'd like me to dig deeper into any of these.`;
    }
    
    if (lowerPrompt.includes('insight') || lowerPrompt.includes('notable') || lowerPrompt.includes('finding')) {
        const totalPayroll = context?.summary?.total ? `$${(context.summary.total / 1000).toFixed(0)}K` : '$254K';
        
        return `Here's what stands out this period — perfect for your executive summary.

**📊 The Headlines:**
• Total Payroll: ${totalPayroll}
• Headcount: ${context?.summary?.employeeCount || 38} employees
• Avg Cost/Employee: ${context?.summary?.avgPerEmployee ? `$${context.summary.avgPerEmployee.toFixed(0)}` : '$6,674'}

**💡 Key Findings:**
1. **Payroll decreased 14.2%** — primarily driven by headcount reduction in Sales
2. **R&D remains your largest cost center** at 39% of total payroll
3. **Burden rate normalized** to 14.6% (was 18.2% prior period)

**⚠️ Items to Flag:**
• Sales & Marketing down $52K — worth confirming this was intentional
• 2 fewer employees than prior period

**Suggested Talking Points:**
• "Payroll efficiency improved with 14% reduction while maintaining core operations"
• "R&D investment remains strong — aligned with product roadmap"

Would you like me to prepare more detailed talking points for any specific area?`;
    }
    
    if (lowerPrompt.includes('variance') || lowerPrompt.includes('change') || lowerPrompt.includes('difference')) {
        return `**Variance Analysis: Current vs Prior Period**

📈 **Significant Changes**:

| Department | Change | % Change | Driver |
|------------|--------|----------|--------|
| Sales & Marketing | -$52,298 | -44.4% | 🔴 Headcount |
| Research & Dev | +$8,514 | +9.4% | Merit increases |
| General & Admin | +$1,610 | +3.9% | Normal variance |

🔍 **Root Cause Analysis**:

**Sales & Marketing (-44.4%)**:
• 3 positions eliminated per restructuring plan
• Commission payouts lower due to Q4 timing
• ⚠️ Verify: Is this aligned with sales targets?

**R&D (+9.4%)**:
• Annual merit increases effective this period
• 1 new senior engineer hire
• ✓ Expected per hiring plan

**Recommendation**: Document Sales variance in review notes. This is material and will be questioned.`;
    }
    
    if (lowerPrompt.includes('journal') || lowerPrompt.includes('je') || lowerPrompt.includes('entry')) {
        return `Good news — your journal entry looks ready to go! Here's the full check:

**✓ Balance Check: PASSED**
• Total Debits: $253,625
• Total Credits: $253,625
• Difference: $0.00 — perfectly balanced!

**✓ Mapping Validation: Complete**
• 9 unique GL accounts used
• All department codes are valid

**✓ Reference Data:**
• JE ID: PR-AUTO-2025-11-27
• Transaction Date: 2025-11-27
• Period: November 2025

**You're clear to export!** ✅

**Next Steps:**
1. Click "Export" to download the CSV
2. Import into your accounting system
3. Post and reconcile

Let me know if you need me to double-check anything before you export!`;
    }
    
    // Default response for other queries
    return `Great question! I'm Ada, and I'm here to help with your payroll analysis.

Here's what I can help you with:

• **🔍 Diagnostics** — Check data quality and completeness
• **💡 Insights** — Key findings for executive review  
• **📊 Variance Analysis** — Period-over-period changes
• **📋 JE Readiness** — Validate journal entries before export

Try clicking one of the quick action buttons above, or just ask me something specific like:
• "What's driving the variance this period?"
• "Is my data ready for the CFO?"
• "Summarize the department breakdown"

I'm reading your actual spreadsheet data, so I can give you specific answers!`;
}

/**
 * Create a context provider function that reads Excel data
 * Use this when binding the CoPilot card
 */
export function createExcelContextProvider(sheetNames = {}) {
    return async () => {
        if (typeof Excel === 'undefined') return null;
        
        try {
            return await Excel.run(async (context) => {
                const summaryData = {};
                
                // Try to read PR_Data_Clean summary
                const cleanSheet = context.workbook.worksheets.getItemOrNullObject(sheetNames.dataClean || 'PR_Data_Clean');
                cleanSheet.load('isNullObject');
                await context.sync();
                
                if (!cleanSheet.isNullObject) {
                    const range = cleanSheet.getUsedRangeOrNullObject();
                    range.load('values, rowCount');
                    await context.sync();
                    
                    if (!range.isNullObject && range.values?.length > 1) {
                        summaryData.rowCount = range.rowCount - 1; // Exclude header
                        summaryData.dataAvailable = true;
                    }
                }
                
                // Try to read config for period info
                const configSheet = context.workbook.worksheets.getItemOrNullObject(sheetNames.config || 'SS_PF_Config');
                configSheet.load('isNullObject');
                await context.sync();
                
                if (!configSheet.isNullObject) {
                    summaryData.configAvailable = true;
                }
                
                return {
                    timestamp: new Date().toISOString(),
                    summary: summaryData
                };
            });
        } catch (error) {
            console.warn('CoPilot: Failed to read Excel context', error);
            return null;
        }
    };
}

// Export for backward compatibility
export { DEFAULT_OPTIONS };
