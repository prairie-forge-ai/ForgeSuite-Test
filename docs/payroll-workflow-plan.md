# Payroll Recorder: Workflow Reattachment Plan

## Decision
- The deeper workflow + configuration experience **must be reattached**. The current `payroll-recorder/src/index.js` home shell (lines 1‑140) only renders the ForgeSuite banner, hero text, and static step cards; it never imports `createWorkflowController`, `CopilotManager`, or the Excel automation helpers that were written under `payroll-recorder/src/`. As a result, none of the run buttons, reconciliation metrics, or notes panel described in `workflow.js` (entire file) can be reached from the task pane even though those modules are still maintained in source control.

## Architecture Overview
1. **Dual-view shell**
   - Keep the existing template as the “Module Home” view (`div[data-view="home"]` inside the root).
   - Add a sibling container (e.g., `div[data-view="workflow"]`) that will host the legacy workflow markup: header, progress sidebar, copilot shell, `workflow-view`, `config-view`, and `step-view`.
   - Home view stays visible until a step card is activated; at that point we hide the home container and show the workflow container. Provide a “Back to Module Home” pill near the workflow header so users can exit.
2. **Controller wiring**
   - Import `createWorkflowController` plus builder/metric helpers.
   - Instantiate it after Office is ready with callbacks that wrap the existing Excel runners (placeholder functions in this sprint) so `controller.showConfig()` and `controller.showStep(n)` can drive the view.
   - Preserve the lightweight state that powers arrow navigation + hero hint text so the ForgeSuite hero stays informative even while the deeper view is open.
3. **Shared constants**
   - Stop duplicating the step metadata in `index.js`; consume `WORKFLOW_STEPS` from `constants.js` so card copy stays in sync with the detailed descriptions/checklists.
4. **Event routing**
   - Keep `nav-prev`/`nav-next` behavior when the home view is focused.
   - When the workflow view is active, repurpose the same buttons to call `controller.goRelative(±1)` instead of scrolling cards.

## Step Card → Screen Mapping
| Card | Controller call | Notes |
| --- | --- | --- |
| Step 0 – Configuration Setup | `controller.showConfig()` | Ensure builder tooling runs first so `SS_PF_Config` is visible before switching views. |
| Step 1 – Import Payroll Data | `controller.showStep(1)` | Trigger the `runStep(1)` callback from the CTA inside the detailed view; home card only handles the transition. |
| Step 2 – Headcount Review | `controller.showStep(2)` | Maintain the optional skip toggle + metric refresh logic already present in `workflow.js:360‑430`. |
| Step 3 – Validate & Reconcile | `controller.showStep(3)` | Wire the reconciliation refresher + bank input persistence via the controller options. |
| Step 4 – Expense Review | `controller.showStep(4)` | Leave placeholder Copilot callouts in the detailed view; they will be wired after Copilot returns. |
| Step 5 – Journal Entry Prep | `controller.showStep(5)` | Keep JE metric reload + export interactions intact. |
| Step 6 – Archive & Reset | `controller.showStep(6)` | Use the general-purpose CTA handler inside `workflow.js` for now. |

## Implementation Steps
1. **Reshape `src/index.js`**
   - Import `createWorkflowController`, `WORKFLOW_STEPS`, metric/builder helpers, and tab-visibility.
   - Split rendering into two helpers: `renderHomeView()` (existing template) and `renderWorkflowHost()` (new container with empty child nodes `pf-header-content`, `workflow-view`, `config-view`, `step-view`, and `pf-copilot-shell` so the controller can hydrate them).
   - Add state + DOM classes so clicking a card calls either `openWorkflow(stepId)` or `returnHome()`.
2. **Style adjustments**
   - Extend `payroll-recorder/styles.css` with simple layout rules that hide/show `[data-view]` containers and keep the workflow view scrollable while preserving the prairie gradient background.
3. **Placeholder runners**
   - Until the Excel automation is reinstated, stub the `runStep`, `openHeadcountReview`, etc., callbacks with `console.warn` + toast notifications so the UI feels interactive without touching workbook state. Replace these stubs with real implementations as soon as we port the original `taskpane.js` logic back in.
4. **Testing**
   - Run `npm run build:payroll` and smoke-test in Excel Online to verify: (a) cards open the correct detailed screen, (b) `Back to home` works, (c) nav buttons shift steps in both contexts. No workbook mutations are expected yet.

## Icon + Copilot Follow-ups
1. **Icons** – Once design ships the new single-color silhouettes, drop them in `assets/` and update the nav buttons + card badges (currently emoji in `src/index.js:55‑63`). Keep sizes consistent with the `.pf-nav-btn--icon` style.
2. **Copilot return** – After the shell stabilizes, re-enable the panel by instantiating `CopilotManager` from `payroll-recorder/src/copilot.js`. Mount its template inside `pf-copilot-shell` and reuse the shared badge styles from `Common/copilot.js` to avoid regressions.
