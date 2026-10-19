# CONINCO Excel Auditor Desktop

The new desktop client uses React, TypeScript, Tailwind CSS and Tauri 2. The
existing Python engine remains the calculation boundary; Tauri launches it as
a local process during development and as a bundled sidecar in release builds.

## Run the UI preview

```bash
npm install
npm run dev
```

The Vite preview includes safe demo data so the Process, Archive and
Comparison screens can be reviewed without opening a workbook.

## Run the desktop shell

Install Rust and the Tauri prerequisites for the target operating system, then:

```bash
npm run tauri:dev
npm run tauri:build
```

The release bundle expects the packaged Python executable at
`src-tauri/engine/excel-auditor-engine`. The development path uses
`../python-engine/.venv/bin/python` automatically.

## Build Windows

Build Windows packages on a Windows machine:

```powershell
cd desktop-app
& .\scripts\build-engine.ps1
npm install
npm run tauri:build
```

The result is a Windows installer under
`src-tauri/target/release/bundle/`. Python, pandas, openpyxl, xlrd, Node.js,
and the Rust runtime are bundled; the target machine does not need these
dependencies installed separately.
