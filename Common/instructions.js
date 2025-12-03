const DEFAULT_MESSAGE =
    "Instructions for this module are on the way. Prairie Forge is publishing the new guidance shortly.";

function createHandler({ moduleName, message, url }) {
    const finalMessage =
        message || (moduleName ? `Instructions for ${moduleName} are coming soon.` : DEFAULT_MESSAGE);
    return () => {
        if (url) {
            window.open(url, "_blank", "noopener,noreferrer");
            return;
        }
        window.alert(finalMessage);
    };
}

/**
 * Attach a shared placeholder handler to an Instructions button until deep
 * documentation is ready. Pass an element or selector so we can rebind after rerenders.
 */
export function bindInstructionsButton(target, options = {}) {
    if (!target) return;
    const handler = createHandler(options);
    if (typeof target === "string") {
        document.querySelectorAll(target).forEach((element) => {
            element?.addEventListener("click", handler);
        });
        return;
    }
    if (target instanceof Element) {
        target.addEventListener("click", handler);
    }
}
