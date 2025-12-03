# Prairie Forge Modules

Modern Prairie Forge task panes now rely on **ES module source files** that are bundled with [esbuild](https://esbuild.github.io/). Every customer-facing module ships a readable `src/` directory plus a compiled `app.bundle.js` that gets loaded by Office.

## Why bundle?

We chose the esbuild-per-module approach because it gives us:

- Source maps for easier debugging inside Excel/PowerPoint task panes
- Dead-code elimination and tree shaking so legacy helpers don’t leak into the shipping bundle
- Real `import`/`export` semantics which play nicely with TypeScript, tests, and modern lint tooling

All *shared* utilities (the `Common/` directory) are authored as ES modules and imported into each bundle. Most helpers come through the module bundles; a few light UI hooks (e.g., the floating info FAB) can also be loaded directly via `../Common/common.js` in the HTML when needed. Because of that:

- Any changes to `Common/` require rebuilding every module that imports those files.
- `npm run build:payroll`, `npm run build:pto`, `npm run build:roster`, or `npm run build:all` are the commands to regenerate bundles.
- During development you can run the same commands with `--watch` appended (e.g. `npm run build:payroll -- --watch`) to keep esbuild rebuilding as you edit.

### Shared styling
- Core colors/typography live in `Common/styles.css`. Avoid module-level `:root` overrides so the shared theme (black background, PF gradients, soft-white text) applies everywhere.
- Step cards, hero text, nav buttons, and the footer all pick up gradients and soft-white text from the shared CSS. If you see mismatches, remove module-specific background/border overrides.
- Shared layout tokens: `--pf-card-gap` controls vertical spacing between cards (20px). Card spacing and gradients are defined in common; module CSS should not reset them.

## Repository layout

```
Common/                 Shared ES modules (tab visibility, branding, copilot helpers, etc.)
module-selector/        Standalone selector HTML/CSS/JS (no bundle yet)
payroll-recorder/       Compiled task pane + readable source in src/
pto-accrual/
scripts/                Small build helpers that wrap esbuild per module
```

Each module folder contains:

- `index.html` – task pane host page
- `app.js` – tiny bootstrapper the manifest points to
- `app.bundle.js` – generated with esbuild (do not hand edit)
- `src/` – the real source files you should modify

## Building everything

```bash
# install dependencies first
npm install

# rebuild every module bundle
npm run build:all
```

Or rebuild a single module:

```bash
npm run build:payroll
npm run build:pto
```

If you change files under `Common/`, run at least the module builds that import those helpers so the task pane picks up the new code. When debugging an Office runtime issue, open the `app.bundle.js.map` emitted beside each bundle so stack traces point back to the ES module source lines.

## Adding new modules

1. Create `<module-name>/src/…` with ES modules.
2. Add a `scripts/build-<module-name>.js` that calls esbuild (see existing scripts for boilerplate).
3. Update `package.json` scripts if you want dedicated build commands.
4. Import any shared helpers from `Common/` inside your source files; esbuild will pull them into the bundle automatically.
5. Run the build script and load the generated `app.bundle.js` via your task pane.

This keeps everything on the same modern module graph and avoids the old global `window.PrairieForge` namespace altogether.
