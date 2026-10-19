# CONINCO Excel Auditor

Offline desktop application for auditing contractor Excel workbooks. The
desktop UI is React, TypeScript, Tailwind CSS, and Tauri 2. The calculation
engine is Python and runs locally; SQLite stores the archive.

## Run the application

```bash
cd desktop-app
npm install
npm run tauri:dev
```

For a production bundle:

```bash
cd desktop-app
npm run tauri:build
```

The generated macOS `.app` and `.dmg` are placed under
`desktop-app/src-tauri/target/release/bundle/`. Windows installers must be
built on Windows because Tauri packages native platform binaries.

## Python engine

The engine supports `.xlsx` and `.xls`, calculates K/S/S-inner, preserves the
source workbook layout, and writes comparison results. See
[`python-engine/README.md`](python-engine/README.md) for engine tests and the
JSON process contract.

## Project layout

- `desktop-app/`: React/Tauri desktop application.
- `python-engine/`: offline Excel reader, classifier, formulas, and report writer.
- `configs/`, `contracts/`, `docs/`: configuration, contracts, and design notes.
