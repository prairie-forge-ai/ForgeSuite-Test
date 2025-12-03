#!/usr/bin/env node

/**
 * Bundles the PTO Accrual task pane.
 */

const esbuild = require("esbuild");
const path = require("path");
const fs = require("fs");

const PROJECT_ROOT = path.resolve(__dirname, "..");
const ENTRY_POINT = path.resolve(PROJECT_ROOT, "pto-accrual", "src", "index.js");
const OUT_FILE = path.resolve(PROJECT_ROOT, "pto-accrual", "app.bundle.js");
const INDEX_HTML = path.resolve(PROJECT_ROOT, "pto-accrual", "index.html");

// Generate build timestamp for cache busting
const BUILD_VERSION = Date.now().toString(36);

async function main() {
    try {
        await esbuild.build({
            entryPoints: [ENTRY_POINT],
            bundle: true,
            outfile: OUT_FILE,
            format: "iife",
            platform: "browser",
            target: ["es2019"],
            sourcemap: true,
            minify: true,
            banner: {
                js: "/* Prairie Forge PTO Accrual */"
            },
            logLevel: "info"
        });
        console.log(`Generated ${path.relative(PROJECT_ROOT, OUT_FILE)} with esbuild.`);

        // Update index.html with new build version
        updateHtmlVersion(INDEX_HTML, BUILD_VERSION);
    } catch (error) {
        console.error("PTO Accrual build failed:", error);
        process.exit(1);
    }
}

function updateHtmlVersion(filePath, version) {
    try {
        let content = fs.readFileSync(filePath, "utf8");
        // Replace version query parameters in script/style tags
        content = content.replace(
            /app\.bundle\.js\?v=[^"]+/g,
            `app.bundle.js?v=${version}`
        );
        fs.writeFileSync(filePath, content, "utf8");
        console.log(`Updated index.html with build version: ${version}`);
    } catch (error) {
        console.warn("Could not update index.html version:", error.message);
    }
}

main();
