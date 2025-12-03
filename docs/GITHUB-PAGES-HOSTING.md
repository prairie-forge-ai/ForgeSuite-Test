# GitHub Pages Hosting (Manual & Simple)

The Module Selector taskpane loads from GitHub Pages at:

```
https://prairie-forge-ai.github.io/Customer-ArchCollins-Foundry/module-selector/
```

Individual modules (e.g., Payroll Recorder) are also served from their folders within the same site, so there is no build pipeline—just edit the static files and push to `main`.

## 1. Edit locally

- `module-selector/index.html`
- `module-selector/selector.js`
- `module-selector/selector.css`
- `payroll-recorder/index.html`
- `payroll-recorder/app.js`
- `payroll-recorder/styles.css`

Keep assets self-contained (relative paths only) so they resolve correctly on GitHub Pages.

## 2. Deploy

1. Commit your changes.
2. Push to `main`.
3. GitHub Pages automatically publishes the latest commit to the URL above (allow ~60 seconds).

## 3. Manifest URL

The production manifest (`ACPTools_manifest.xml`) already points at `https://prairie-forge-ai.github.io/Customer-ArchCollins-Foundry/module-selector/index.html`, so no edits are required unless you move the site. If Excel caches the add-in, bump `<Version>` in the manifest and re-upload it.

## 4. Test in Excel

1. Insert → Office Add-ins → Upload My Add-in → select `ACPTools_manifest.xml`.
2. Use the ribbon button (ForgeSuite → Launch Modules) to load the Module Selector.
3. Launch a module (e.g., Payroll Recorder) and confirm it loads from GitHub Pages.

Repeat steps 1–2 whenever you modify the static files.
