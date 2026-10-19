# Development guide

## Requirements

- Node.js and npm.
- Rust toolchain and Tauri prerequisites for the target operating system.
- Python is only needed for development; release builds bundle the Python
  engine as a sidecar.

## Build and run

From `desktop-app/`:

```bash
npm install
npm run dev
npm run tauri:dev
npm run tauri:build
```

`npm run dev` opens a browser preview with demo data. `npm run tauri:dev`
runs the native desktop application and connects it to the local Python
engine. `npm run tauri:build` creates the platform installer or application
bundle.

## Python engine setup

```bash
python -m venv python-engine/.venv
python-engine/.venv/bin/python -m pip install -r python-engine/requirements.txt
PYTHONPATH=python-engine/src python-engine/.venv/bin/python \
  -m unittest discover -s python-engine/tests -p 'test_*.py'
```

The Tauri development build uses `python-engine/.venv/bin/python`. Release
builds use the bundled executable at `desktop-app/src-tauri/engine/`.

