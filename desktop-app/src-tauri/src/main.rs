#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use chrono::Utc;
use rfd::FileDialog;
use rusqlite::{params, Connection};
use serde::{Deserialize, Serialize};
use serde_json::Value;
use std::{fs, path::{Path, PathBuf}, process::Command};
use tauri::{AppHandle, Manager};
use uuid::Uuid;

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ArchiveEntry {
    id: String,
    submitted_name: String,
    processed_name: String,
    sheet: String,
    status: String,
    rows: i64,
    mismatches: i64,
    created_at: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct AuditResult {
    id: String,
    submitted_name: String,
    processed_name: String,
    sheet: String,
    rows: i64,
    mismatches: i64,
    reference_formulas: i64,
}

// Return the per-user application directory used for archive and reference data.
fn app_root(app: &AppHandle) -> Result<PathBuf, String> {
    app.path().app_data_dir().map_err(|error| error.to_string())
}

// Open the local SQLite archive and create its table on first use.
fn database(app: &AppHandle) -> Result<Connection, String> {
    let root = app_root(app)?;
    fs::create_dir_all(&root).map_err(|error| error.to_string())?;
    let connection = Connection::open(root.join("archive.sqlite3")).map_err(|error| error.to_string())?;
    connection.execute_batch("CREATE TABLE IF NOT EXISTS audits (id TEXT PRIMARY KEY, submitted_path TEXT NOT NULL, processed_path TEXT NOT NULL, submitted_name TEXT NOT NULL, processed_name TEXT NOT NULL, sheet TEXT NOT NULL, status TEXT NOT NULL, rows INTEGER NOT NULL, mismatches INTEGER NOT NULL, created_at TEXT NOT NULL)").map_err(|error| error.to_string())?;
    Ok(connection)
}

#[tauri::command]
// Show the contractor workbook picker.
fn choose_workbook() -> Option<String> {
    FileDialog::new().add_filter("Excel workbook", &["xlsx", "xls"]).pick_file().map(|path| path.to_string_lossy().to_string())
}

#[tauri::command]
// Show the standard formula workbook picker, including Apple Numbers files.
fn choose_reference_workbook() -> Option<String> {
    FileDialog::new().add_filter("Formula reference workbook", &["xlsx", "xls", "numbers"]).pick_file().map(|path| path.to_string_lossy().to_string())
}

// Return the private directory where a copied standard workbook is stored.
fn reference_storage_dir(app: &AppHandle) -> Result<PathBuf, String> {
    Ok(app_root(app)?.join("formula-reference"))
}

// Return the text manifest that remembers the saved standard workbook path.
fn reference_manifest_path(app: &AppHandle) -> Result<PathBuf, String> {
    Ok(app_root(app)?.join("formula-reference-path.txt"))
}

// Return the JSON cache path for the parsed standard formula catalog.
fn formula_catalog_cache_path(app: &AppHandle) -> Result<PathBuf, String> {
    Ok(reference_storage_dir(app)?.join("formula-catalog.json"))
}

// Remove either a file or directory when it exists.
fn remove_path(path: &Path) -> Result<(), String> {
    if path.is_dir() {
        fs::remove_dir_all(path).map_err(|error| error.to_string())?;
    } else if path.exists() {
        fs::remove_file(path).map_err(|error| error.to_string())?;
    }
    Ok(())
}

// Copy a file or recursively copy an Apple Numbers package directory.
fn copy_path(source: &Path, destination: &Path) -> Result<(), String> {
    if source.is_dir() {
        fs::create_dir_all(destination).map_err(|error| error.to_string())?;
        for entry in fs::read_dir(source).map_err(|error| error.to_string())? {
            let entry = entry.map_err(|error| error.to_string())?;
            copy_path(&entry.path(), &destination.join(entry.file_name()))?;
        }
    } else {
        if let Some(parent) = destination.parent() {
            fs::create_dir_all(parent).map_err(|error| error.to_string())?;
        }
        fs::copy(source, destination).map_err(|error| error.to_string())?;
    }
    Ok(())
}

fn cached_formula_catalog(
    app: &AppHandle,
    reference_path: &Path,
    reference_sheet: &str,
) -> Result<Option<Value>, String> {
    // Reuse only a cache tied to the same reference path and worksheet.
    let cache_path = formula_catalog_cache_path(app)?;
    if !cache_path.is_file() {
        return Ok(None);
    }
    let text = fs::read_to_string(cache_path).map_err(|error| error.to_string())?;
    let cached: Value = match serde_json::from_str(&text) {
        Ok(value) => value,
        Err(_) => return Ok(None),
    };
    let expected_path = reference_path.to_string_lossy();
    let same_source = cached.get("reference_path").and_then(Value::as_str) == Some(expected_path.as_ref());
    let same_sheet = cached.get("reference_sheet").and_then(Value::as_str) == Some(reference_sheet);
    if !same_source || !same_sheet {
        return Ok(None);
    }
    Ok(cached.get("catalog").cloned().filter(|catalog| catalog.get("entries").and_then(Value::as_array).is_some()))
}

fn load_or_create_formula_catalog(
    app: &AppHandle,
    reference_path: &Path,
    reference_sheet: &str,
) -> Result<Value, String> {
    // Formula workflow: reuse the cached catalog or build it from the selected sheet.
    if let Some(catalog) = cached_formula_catalog(app, reference_path, reference_sheet)? {
        return Ok(catalog);
    }
    if !reference_path.is_file() {
        return Err("Formula reference workbook does not exist and its JSON catalog is unavailable.".into());
    }
    let response = run_engine(app, serde_json::json!({
        "action": "load_reference_catalog",
        "path": reference_path,
        "sheet": reference_sheet,
    }))?;
    let catalog = response.get("result").cloned().ok_or("Python did not return the formula catalog.")?;
    let cache_path = formula_catalog_cache_path(app)?;
    let cache = serde_json::json!({
        "reference_path": reference_path.to_string_lossy(),
        "reference_sheet": reference_sheet,
        "catalog": catalog,
    });
    fs::write(cache_path, serde_json::to_vec_pretty(&cache).map_err(|error| error.to_string())?)
        .map_err(|error| error.to_string())?;
    Ok(cache.get("catalog").cloned().unwrap_or(Value::Null))
}

#[tauri::command]
// Return the saved reference path only when the copied source still exists.
fn get_reference_workbook(app: AppHandle) -> Result<Option<String>, String> {
    let manifest = reference_manifest_path(&app)?;
    if !manifest.is_file() {
        return Ok(None);
    }
    let saved = fs::read_to_string(manifest).map_err(|error| error.to_string())?;
    let path = PathBuf::from(saved.trim());
    Ok(path.exists().then(|| path.to_string_lossy().to_string()))
}

#[tauri::command]
// Replace the stored reference atomically and invalidate its old catalog cache.
fn save_reference_workbook(app: AppHandle, path: String) -> Result<String, String> {
    // Formula workflow: copy the selected reference file into app storage.
    let source = PathBuf::from(path);
    if !source.is_file() && !source.is_dir() {
        return Err("Formula reference workbook does not exist.".into());
    }
    let root = app_root(&app)?;
    fs::create_dir_all(&root).map_err(|error| error.to_string())?;
    let staging = root.join("formula-reference-new");
    let storage = reference_storage_dir(&app)?;
    remove_path(&staging)?;
    fs::create_dir_all(&staging).map_err(|error| error.to_string())?;
    let stored_file = staging.join(source.file_name().unwrap_or_default());
    copy_path(&source, &stored_file)?;
    remove_path(&storage)?;
    fs::rename(&staging, &storage).map_err(|error| error.to_string())?;
    let stored_path = storage.join(source.file_name().unwrap_or_default());
    fs::write(reference_manifest_path(&app)?, stored_path.to_string_lossy().as_bytes()).map_err(|error| error.to_string())?;
    remove_path(&formula_catalog_cache_path(&app)?)?;
    Ok(stored_path.to_string_lossy().to_string())
}

#[tauri::command]
// Remove the saved reference, manifest, and cached catalog.
fn clear_reference_workbook(app: AppHandle) -> Result<(), String> {
    remove_path(&reference_storage_dir(&app)?)?;
    remove_path(&reference_manifest_path(&app)?)
}

#[tauri::command]
// Ask Python to inspect the contractor workbook and list its worksheets.
fn list_sheets(app: AppHandle, path: String) -> Result<Vec<String>, String> {
    let response = run_engine(&app, serde_json::json!({"action": "list_sheets", "path": path}))?;
    response.get("result").and_then(|result| result.get("sheets")).and_then(Value::as_array).map(|items| items.iter().filter_map(|item| item.get("name").and_then(Value::as_str)).map(String::from).collect()).ok_or_else(|| "Python did not return worksheets.".into())
}

#[tauri::command]
// Ask Python to inspect the standard workbook and list its worksheets.
fn list_reference_sheets(app: AppHandle, path: String) -> Result<Vec<String>, String> {
    let response = run_engine(&app, serde_json::json!({"action": "list_reference_sheets", "path": path}))?;
    response.get("result").and_then(|result| result.get("sheets")).and_then(Value::as_array).map(|items| items.iter().filter_map(|item| item.get("name").and_then(Value::as_str)).map(String::from).collect()).ok_or_else(|| "Python did not return reference worksheets.".into())
}

#[tauri::command]
// Return parsed catalog entries for the selected standard worksheet.
fn load_reference_catalog(app: AppHandle, path: String, sheet: String) -> Result<Vec<Value>, String> {
    // Formula workflow: expose standard entries to the catalog page.
    let catalog = load_or_create_formula_catalog(&app, &PathBuf::from(path), &sheet)?;
    catalog
        .get("entries")
        .and_then(Value::as_array)
        .cloned()
        .ok_or_else(|| "Python did not return the formula catalog.".into())
}

#[tauri::command]
// Save the parsed catalog to a user-selected JSON path.
fn export_formula_catalog(app: AppHandle, path: String, sheet: String) -> Result<Option<String>, String> {
    let reference = PathBuf::from(path);
    let catalog = load_or_create_formula_catalog(&app, &reference, &sheet)?;
    let Some(target) = FileDialog::new()
        .add_filter("JSON file", &["json"])
        .set_file_name("formula-catalog.json")
        .save_file()
    else {
        return Ok(None);
    };
    let export = serde_json::json!({
        "reference_path": reference.to_string_lossy(),
        "reference_sheet": sheet,
        "catalog": catalog,
    });
    fs::write(&target, serde_json::to_vec_pretty(&export).map_err(|error| error.to_string())?)
        .map_err(|error| format!("Could not save formula catalog JSON: {error}"))?;
    Ok(Some(target.to_string_lossy().to_string()))
}

#[tauri::command]
// Run the full audit, preserve both workbooks, and register the result in SQLite.
fn run_audit(app: AppHandle, path: String, sheet: String, reference_path: String, reference_sheet: String) -> Result<AuditResult, String> {
    // Run Audit step 1: resolve the submitted workbook and standard reference.
    let source = PathBuf::from(path);
    let reference = PathBuf::from(reference_path);
    // Run Audit step 2: load the authoritative formula catalog.
    let catalog = load_or_create_formula_catalog(&app, &reference, &reference_sheet)?;
    // Run Audit step 3: create archive folders and unique output paths.
    let root = app_root(&app)?;
    let id = Uuid::new_v4().to_string();
    let submitted_dir = root.join("submitted");
    let processed_dir = root.join("processed");
    fs::create_dir_all(&submitted_dir).map_err(|error| error.to_string())?;
    fs::create_dir_all(&processed_dir).map_err(|error| error.to_string())?;
    let submitted = submitted_dir.join(format!("{}_{}", &id[..8], source.file_name().unwrap_or_default().to_string_lossy()));
    let processed = processed_dir.join(format!("{}_calculated.xlsx", id));
    // Run Audit step 4: preserve the original submitted workbook in the archive.
    fs::copy(&source, &submitted).map_err(|error| error.to_string())?;
    // Run Audit step 5: ask Python to map, calculate, compare, and export.
    let response = run_engine(&app, serde_json::json!({"action":"export_calculated_workbook", "path": submitted, "output": processed, "sheet": sheet, "reference_path": reference, "reference_sheet": reference_sheet, "reference_catalog": catalog}))?;
    // Run Audit step 6: read Python's summary for the UI and archive record.
    let result = response.get("result").ok_or("Python did not return an audit result")?;
    let rows = result.get("calculated_count").and_then(Value::as_i64).unwrap_or(0);
    let mismatches = result.get("mismatch_count").and_then(Value::as_i64).unwrap_or(0);
    let reference_formulas = result.get("reference_formula_count").and_then(Value::as_i64).unwrap_or(0);
    // Run Audit step 7: store the processed workbook and mismatch summary.
    let connection = database(&app)?;
    connection.execute("INSERT INTO audits VALUES (?1, ?2, ?3, ?4, ?5, ?6, 'COMPLETED', ?7, ?8, ?9)", params![id, submitted.to_string_lossy(), processed.to_string_lossy(), source.file_name().unwrap_or_default().to_string_lossy(), processed.file_name().unwrap_or_default().to_string_lossy(), sheet, rows, mismatches, Utc::now().to_rfc3339()]).map_err(|error| error.to_string())?;
    Ok(AuditResult { id, submitted_name: source.file_name().unwrap_or_default().to_string_lossy().to_string(), processed_name: processed.file_name().unwrap_or_default().to_string_lossy().to_string(), sheet, rows, mismatches, reference_formulas })
}

fn run_engine(app: &AppHandle, request: Value) -> Result<Value, String> {
    // Bridge one JSON request from Tauri to the Python engine and return its response.
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    let python = root.join("python-engine/.venv/bin/python");
    let bundled_engine = app.path().resource_dir().map_err(|error| error.to_string())?.join("engine/excel-auditor-engine");
    #[cfg(target_os = "windows")]
    let bundled_engine = bundled_engine.with_extension("exe");
    let (program, args, cwd): (String, Vec<String>, Option<PathBuf>) = if python.is_file() { (python.to_string_lossy().to_string(), vec!["-m".to_string(), "excel_engine.main".to_string()], Some(root.join("python-engine"))) } else { (bundled_engine.to_string_lossy().to_string(), vec![], None) };
    let mut command = Command::new(program);
    command.args(args).stdin(std::process::Stdio::piped()).stdout(std::process::Stdio::piped()).stderr(std::process::Stdio::piped());
    if let Some(directory) = cwd { command.current_dir(directory).env("PYTHONPATH", root.join("python-engine/src")); }
    let mut child = command.spawn().map_err(|error| error.to_string())?;
    { use std::io::Write; child.stdin.take().ok_or("Python stdin unavailable")?.write_all(request.to_string().as_bytes()).map_err(|error| error.to_string())?; }
    let output = child.wait_with_output().map_err(|error| error.to_string())?;
    if !output.status.success() { return Err(String::from_utf8_lossy(&output.stderr).to_string()); }
    serde_json::from_slice(&output.stdout).map_err(|error| format!("Invalid Python response: {error}"))
}

#[tauri::command]
// Read archived audit metadata for the Archive page.
fn list_archive(app: AppHandle) -> Result<Vec<ArchiveEntry>, String> {
    let connection = database(&app)?;
    let mut query = connection.prepare("SELECT id, submitted_name, processed_name, sheet, status, rows, mismatches, created_at FROM audits ORDER BY created_at DESC").map_err(|error| error.to_string())?;
    let entries = query.query_map([], |row| Ok(ArchiveEntry { id: row.get(0)?, submitted_name: row.get(1)?, processed_name: row.get(2)?, sheet: row.get(3)?, status: row.get(4)?, rows: row.get(5)?, mismatches: row.get(6)?, created_at: row.get(7)? })).map_err(|error| error.to_string())?.filter_map(Result::ok).collect();
    Ok(entries)
}

#[tauri::command]
// Delete one audit's stored files, preview, and database row.
fn delete_archive(app: AppHandle, id: String) -> Result<(), String> {
    let connection = database(&app)?;
    let paths: (String, String) = connection.query_row("SELECT submitted_path, processed_path FROM audits WHERE id = ?1", [&id], |row| Ok((row.get(0)?, row.get(1)?))).map_err(|error| error.to_string())?;
    for path in [paths.0, paths.1] {
        if Path::new(&path).exists() {
            fs::remove_file(&path).map_err(|error| format!("Could not delete {path}: {error}"))?;
        }
    }
    let processed_path: String = connection.query_row("SELECT processed_path FROM audits WHERE id = ?1", [&id], |row| row.get(0)).map_err(|error| error.to_string())?;
    let preview = Path::new(&processed_path).with_file_name(format!("{}_preview.tsv", Path::new(&processed_path).file_stem().unwrap_or_default().to_string_lossy()));
    if preview.exists() {
        fs::remove_file(preview).map_err(|error| format!("Could not delete comparison preview: {error}"))?;
    }
    connection.execute("DELETE FROM audits WHERE id = ?1", [&id]).map_err(|error| error.to_string())?;
    Ok(())
}

#[tauri::command]
// Copy the submitted or processed workbook to a user-selected destination.
fn download_archive(app: AppHandle, id: String, kind: String) -> Result<String, String> {
    let connection = database(&app)?;
    let column = if kind == "processed" { "processed_path" } else { "submitted_path" };
    let path: String = connection.query_row(&format!("SELECT {column} FROM audits WHERE id = ?1"), [&id], |row| row.get(0)).map_err(|error| error.to_string())?;
    if let Some(target) = FileDialog::new().set_file_name(Path::new(&path).file_name().unwrap_or_default().to_string_lossy()).save_file() {
        fs::copy(&path, &target).map_err(|error| format!("Could not save file: {error}"))?;
        return Ok(target.to_string_lossy().to_string());
    }
    Ok(String::new())
}

#[tauri::command]
// Open the submitted or processed workbook with the platform default app.
fn open_archive_file(app: AppHandle, id: String, kind: String) -> Result<(), String> {
    let connection = database(&app)?;
    let column = if kind == "processed" { "processed_path" } else { "submitted_path" };
    let path: String = connection.query_row(&format!("SELECT {column} FROM audits WHERE id = ?1"), [&id], |row| row.get(0)).map_err(|error| error.to_string())?;
    #[cfg(target_os = "macos")] { Command::new("open").arg(path).spawn().map_err(|error| error.to_string())?; }
    #[cfg(target_os = "windows")] { Command::new("cmd").args(["/C", "start", "", &path]).spawn().map_err(|error| error.to_string())?; }
    Ok(())
}

#[tauri::command]
// Load the TSV comparison preview generated alongside the processed workbook.
fn load_comparison(app: AppHandle, id: String) -> Result<Vec<Value>, String> {
    let connection = database(&app)?;
    let path: String = connection.query_row("SELECT processed_path FROM audits WHERE id = ?1", [&id], |row| row.get(0)).map_err(|error| error.to_string())?;
    let preview = Path::new(&path).with_file_name(format!("{}_preview.tsv", Path::new(&path).file_stem().unwrap_or_default().to_string_lossy()));
    let text = fs::read_to_string(preview).map_err(|error| error.to_string())?;
    let mut reader = csv::ReaderBuilder::new()
        .delimiter(b'\t')
        .has_headers(true)
        .from_reader(text.as_bytes());
    let headers = reader
        .headers()
        .map_err(|error| format!("Could not read comparison headers: {error}"))?
        .iter()
        .map(str::to_owned)
        .collect::<Vec<_>>();
    reader
        .records()
        .map(|record| {
            let record = record.map_err(|error| format!("Could not read comparison row: {error}"))?;
            Ok(headers
                .iter()
                .zip(record.iter())
                .map(|(key, value)| (key.clone(), Value::String(value.to_owned())))
                .collect::<serde_json::Map<_, _>>()
                .into())
        })
        .collect()
}

// Build the Tauri application and register every frontend command.
fn main() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![choose_workbook, choose_reference_workbook, get_reference_workbook, save_reference_workbook, clear_reference_workbook, list_sheets, list_reference_sheets, load_reference_catalog, export_formula_catalog, run_audit, list_archive, delete_archive, download_archive, open_archive_file, load_comparison])
        .run(tauri::generate_context!())
        .expect("error while running CONINCO Excel Auditor");
}
