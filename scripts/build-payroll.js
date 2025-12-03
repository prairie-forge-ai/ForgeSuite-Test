#!/usr/bin/env node

/**
 * Bundles the Payroll Recorder task pane using esbuild so we can split the
 * source into modules and still ship a single Office-compatible artifact.
 */

const esbuild = require('esbuild');
const path = require('path');
const fs = require('fs');

const PROJECT_ROOT = path.resolve(__dirname, '..');
const ENTRY_POINT = path.resolve(PROJECT_ROOT, 'payroll-recorder', 'src', 'workflow.js');
const OUT_FILE = path.resolve(PROJECT_ROOT, 'payroll-recorder', 'app.bundle.js');
const APP_JS_FILE = path.resolve(PROJECT_ROOT, 'payroll-recorder', 'app.js');

// Generate build timestamp for cache busting
const BUILD_VERSION = Date.now().toString(36);

async function main() {
    try {
        await esbuild.build({
            entryPoints: [ENTRY_POINT],
            bundle: true,
            outfile: OUT_FILE,
            format: 'iife',
            platform: 'browser',
            target: ['es2019'],
            sourcemap: true,
            minify: true,
            banner: {
                js: '/* Prairie Forge Payroll Recorder */'
            },
            logLevel: 'info'
        });
        console.log(`Generated ${path.relative(PROJECT_ROOT, OUT_FILE)} with esbuild.`);

        // Update app.js with new build version
        updateAppJsVersion(APP_JS_FILE, BUILD_VERSION);
    } catch (error) {
        console.error('Payroll Recorder build failed:', error);
        process.exit(1);
    }
}

function updateAppJsVersion(filePath, version) {
    try {
        let content = fs.readFileSync(filePath, 'utf8');
        // Replace the version query parameter in bundleSrc
        content = content.replace(
            /const bundleSrc = 'app\.bundle\.js\?v=[^']+';/,
            `const bundleSrc = 'app.bundle.js?v=${version}';`
        );
        fs.writeFileSync(filePath, content, 'utf8');
        console.log(`Updated app.js with build version: ${version}`);
    } catch (error) {
        console.warn('Could not update app.js version:', error.message);
    }
}

main();
