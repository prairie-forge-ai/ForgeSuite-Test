import { BUILDER_ALLOWED_USERS, BUILDER_VISIBILITY_SHEETS } from "./constants.js";

export function createBuilderController({ readConfig }) {
    let isBuilder = true;

    async function readBuilderFlag() {
        if (typeof readConfig !== "function") return true;
        try {
            const config = await readConfig();
            const flag = config?.["Builder Mode"];
            if (flag == null) return true;
            return String(flag).trim().toLowerCase() === "true";
        } catch (error) {
            console.warn("Unable to read Builder Mode flag:", error);
            return true;
        }
    }

    async function determineBuilderAccess() {
        try {
            if (OfficeRuntime?.auth?.getAccessToken) {
                const token = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: false });
                if (token) {
                    const response = await fetch("https://graph.microsoft.com/v1.0/me?$select=userPrincipalName", {
                        headers: { Authorization: `Bearer ${token}` }
                    });
                    if (response.ok) {
                        const payload = await response.json();
                        const upn = (payload?.userPrincipalName || "").toLowerCase();
                        const allowed = BUILDER_ALLOWED_USERS.some((entry) => entry.toLowerCase() === upn);
                        if (allowed) {
                            return await readBuilderFlag();
                        }
                        return false;
                    }
                    console.warn("Graph request for builder mode failed:", await response.text());
                }
            }
        } catch (error) {
            console.warn("Builder mode user lookup failed:", error);
        }
        return await readBuilderFlag();
    }

    async function setSheetVisibility(visible) {
        try {
            await Excel.run(async (context) => {
                const targets = BUILDER_VISIBILITY_SHEETS.map((sheetName) => {
                    const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
                    sheet.load("name");
                    return sheet;
                });
                await context.sync();
                targets.forEach((sheet) => {
                    if (!sheet.isNullObject) {
                        sheet.visibility = visible ? Excel.SheetVisibility.visible : Excel.SheetVisibility.hidden;
                    }
                });
                await context.sync();
            });
        } catch (error) {
            console.warn("Unable to toggle builder sheet visibility:", error);
        }
    }

    return {
        async refresh() {
            isBuilder = await determineBuilderAccess();
            await setSheetVisibility(isBuilder);
            return isBuilder;
        },
        isBuilder: () => isBuilder
    };
}
