export const DEFAULT_SYSTEM_PROMPT = `You are Prairie Forge Payroll Copilot. Stay within Microsoft 365 boundaries and only perform the actions that Prairie Forge exposes to you. Focus on payroll diagnostics, expense review insights, and journal entry readiness. Explain findings plainly and list the next two actions the accountant should take.`;

export const COPILOT_ACTIONS = [
    {
        id: "PF.Diagnostics",
        label: "Diagnostics",
        description: "Summarize workbook health, totals, and blocking issues.",
        payloadHint: "Provide Data_Clean totals, reconciliation status, and any alerts."
    },
    {
        id: "PF.GenerateExpenseInsights",
        label: "Expense Insights",
        description: "Describe notable pay variances by department + category.",
        payloadHint: "Attach Expense Review pivots or totals for the current run."
    },
    {
        id: "PF.JournalPrep",
        label: "Journal Prep QA",
        description: "Validate draft journal entries and highlight imbalances.",
        payloadHint: "Send JE Draft totals, JE_ID, and mapping anomalies."
    },
    {
        id: "PF.CustomPrompt",
        label: "Ad-hoc Prompt",
        description: "Use the text box below for a bespoke request.",
        payloadHint: "Describe the question you want Copilot to answer.",
        showInPanel: false
    }
];

function getOfficeCopilot() {
    if (typeof Office === "undefined") return null;
    return Office?.context?.copilot ?? null;
}

function buildMessageId() {
    return `msg-${Date.now().toString(36)}-${Math.random().toString(16).slice(2)}`;
}

export class CopilotManager {
    constructor(actions = COPILOT_ACTIONS, systemPrompt = DEFAULT_SYSTEM_PROMPT) {
        this.actions = actions;
        this.state = {
            isAvailable: false,
            lastAvailabilityCheck: null,
            isBusy: false,
            statusText: "Microsoft 365 Copilot is idle.",
            isPanelCollapsed: false,
            messages: [],
            systemPrompt,
            lastError: null
        };
        this.listeners = new Set();
        this.customRequestHandler = null;
        this.contextProvider = null;
    }

    configure(options = {}) {
        if (typeof options.requestHandler === "function") {
            this.customRequestHandler = options.requestHandler;
        }
        if (typeof options.contextProvider === "function") {
            this.contextProvider = options.contextProvider;
        }
        if (typeof options.systemPrompt === "string" && options.systemPrompt.trim()) {
            this.state.systemPrompt = options.systemPrompt.trim();
        }
    }

    subscribe(listener) {
        if (typeof listener !== "function") return () => {};
        this.listeners.add(listener);
        listener(this.getState());
        return () => this.listeners.delete(listener);
    }

    notify() {
        const snapshot = this.getState();
        this.listeners.forEach((listener) => listener(snapshot));
    }

    getState() {
        return {
            ...this.state,
            messages: [...this.state.messages]
        };
    }

    setStatus(message) {
        this.state.statusText = message;
        this.notify();
    }

    setPanelCollapsed(collapsed) {
        this.state.isPanelCollapsed = collapsed;
        this.notify();
    }

    logMessage(role, content, metadata = {}) {
        this.state.messages = [
            ...this.state.messages,
            {
                id: buildMessageId(),
                role,
                content,
                timestamp: new Date().toISOString(),
                metadata
            }
        ].slice(-20);
        this.notify();
    }

    async detectAvailability() {
        const copilot = getOfficeCopilot();
        const available = Boolean(copilot);
        this.state.isAvailable = available;
        this.state.lastAvailabilityCheck = new Date().toISOString();
        this.setStatus(
            available
                ? "Microsoft 365 Copilot is available in this workbook."
                : "Copilot is not yet provisioned in this environment."
        );
        return available;
    }

    async requestAction(actionId, options = {}) {
        const action = this.actions.find((entry) => entry.id === actionId);
        if (!action) {
            const error = new Error("Unknown Copilot action.");
            this.state.lastError = error;
            this.logMessage("system", error.message);
            throw error;
        }

        if (!this.state.isAvailable && !(await this.detectAvailability())) {
            const error = new Error("Copilot is unavailable. Please contact your admin.");
            this.state.lastError = error;
            this.logMessage("system", error.message);
            throw error;
        }

        const payload = {
            actionId,
            prompt: options.prompt || action.description,
            contextSummary: options.contextSummary || null,
            data: options.data || null,
            systemPrompt: this.state.systemPrompt
        };

        const handler = this.customRequestHandler || this.simulateResponse.bind(this);

        try {
            this.state.isBusy = true;
            this.state.lastError = null;
            this.setStatus(`Sending ${action.label} to Copilot...`);

            const context = this.contextProvider ? await this.contextProvider(action) : null;
            const outboundPayload = context ? { ...payload, context } : payload;

            this.logMessage("user", payload.prompt, { actionId });
            const response = await handler(actionId, outboundPayload);

            this.state.isBusy = false;
            if (response?.message) {
                this.logMessage("assistant", response.message, {
                    actionId,
                    response
                });
            } else {
                this.logMessage("assistant", "Copilot responded without a message payload.", {
                    actionId,
                    response
                });
            }
            this.setStatus("Copilot responded.");
            return response;
        } catch (error) {
            this.state.isBusy = false;
            this.state.lastError = error;
            const message = error?.message || "Copilot request failed.";
            this.logMessage("system", message, { actionId });
            this.setStatus(message);
            throw error;
        }
    }

    async simulateResponse(actionId, payload) {
        const action = this.actions.find((entry) => entry.id === actionId);
        return new Promise((resolve) => {
            setTimeout(() => {
                const guidance = [
                    "Document any blockers directly in the taskpane.",
                    "Capture reviewer initials once Copilot guidance is applied.",
                    "Schedule the archive step after journal export is approved."
                ];
                resolve({
                    message: [
                        `**${action?.label || actionId}** request acknowledged.`,
                        payload?.contextSummary
                            ? `Context received: ${payload.contextSummary}`
                            : "No runtime context was attached, so Copilot will rely on your prompt.",
                        `Next steps: ${guidance[Math.floor(Math.random() * guidance.length)]}`
                    ].join("\n\n"),
                    echo: payload
                });
            }, 600);
        });
    }
}
