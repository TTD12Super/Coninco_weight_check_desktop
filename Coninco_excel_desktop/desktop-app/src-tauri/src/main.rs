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
}

fn app_root(app: &AppHandle) -> Result<PathBuf, String> {
    app.path().app_data_dir().map_err(|error| error.to_string())
}

fn database(app: &AppHandle) -> Result<Connection, String> {
    let root = app_root(app)?;
    fs::create_dir_all(&root).map_err(|error| error.to_string())?;
    let connection = Connection::open(root.join("archive.sqlite3")).map_err(|error| error.to_string())?;
    connection.execute_batch("CREATE TABLE IF NOT EXISTS audits (id TEXT PRIMARY KEY, submitted_path TEXT NOT NULL, processed_path TEXT NOT NULL, submitted_name TEXT NOT NULL, processed_name TEXT NOT NULL, sheet TEXT NOT NULL, status TEXT NOT NULL, rows INTEGER NOT NULL, mismatches INTEGER NOT NULL, created_at TEXT NOT NULL)").map_err(|error| error.to_string())?;
    Ok(connection)
}

#[tauri::command]
fn choose_workbook() -> Option<String> {
    FileDialog::new().add_filter("Excel workbook", &["xlsx", "xls"]).pick_file().map(|path| path.to_string_lossy().to_string())
}

#[tauri::command]
fn list_sheets(app: AppHandle, path: String) -> Result<Vec<String>, String> {
    let response = run_engine(&app, serde_json::json!({"action": "list_sheets", "path": path}))?;
    response.get("result").and_then(|result| result.get("sheets")).and_then(Value::as_array).map(|items| items.iter().filter_map(|item| item.get("name").and_then(Value::as_str)).map(String::from).collect()).ok_or_else(|| "Python did not return worksheets.".into())
}

#[tauri::command]
fn run_audit(app: AppHandle, path: String, sheet: String) -> Result<AuditResult, String> {
    let source = PathBuf::from(path);
    let root = app_root(&app)?;
    let id = Uuid::new_v4().to_string();
    let submitted_dir = root.join("submitted");
    let processed_dir = root.join("processed");
    fs::create_dir_all(&submitted_dir).map_err(|error| error.to_string())?;
    fs::create_dir_all(&processed_dir).map_err(|error| error.to_string())?;
    let submitted = submitted_dir.join(format!("{}_{}", &id[..8], source.file_name().unwrap_or_default().to_string_lossy()));
    let processed = processed_dir.join(format!("{}_calculated.xlsx", id));
    fs::copy(&source, &submitted).map_err(|error| error.to_string())?;
    let response = run_engine(&app, serde_json::json!({"action":"export_calculated_workbook", "path": submitted, "output": processed, "sheet": sheet}))?;
    let result = response.get("result").ok_or("Python did not return an audit result")?;
    let rows = result.get("calculated_count").and_then(Value::as_i64).unwrap_or(0);
    let mismatches = result.get("mismatch_count").and_then(Value::as_i64).unwrap_or(0);
    let connection = database(&app)?;
    connection.execute("INSERT INTO audits VALUES (?1, ?2, ?3, ?4, ?5, ?6, 'COMPLETED', ?7, ?8, ?9)", params![id, submitted.to_string_lossy(), processed.to_string_lossy(), source.file_name().unwrap_or_default().to_string_lossy(), processed.file_name().unwrap_or_default().to_string_lossy(), sheet, rows, mismatches, Utc::now().to_rfc3339()]).map_err(|error| error.to_string())?;
    Ok(AuditResult { id, submitted_name: source.file_name().unwrap_or_default().to_string_lossy().to_string(), processed_name: processed.file_name().unwrap_or_default().to_string_lossy().to_string(), sheet, rows, mismatches })
}

fn run_engine(app: &AppHandle, request: Value) -> Result<Value, String> {
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
fn list_archive(app: AppHandle) -> Result<Vec<ArchiveEntry>, String> {
    let connection = database(&app)?;
    let mut query = connection.prepare("SELECT id, submitted_name, processed_name, sheet, status, rows, mismatches, created_at FROM audits ORDER BY created_at DESC").map_err(|error| error.to_string())?;
    let entries = query.query_map([], |row| Ok(ArchiveEntry { id: row.get(0)?, submitted_name: row.get(1)?, processed_name: row.get(2)?, sheet: row.get(3)?, status: row.get(4)?, rows: row.get(5)?, mismatches: row.get(6)?, created_at: row.get(7)? })).map_err(|error| error.to_string())?.filter_map(Result::ok).collect();
    Ok(entries)
}

#[tauri::command]
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
fn open_archive_file(app: AppHandle, id: String, kind: String) -> Result<(), String> {
    let connection = database(&app)?;
    let column = if kind == "processed" { "processed_path" } else { "submitted_path" };
    let path: String = connection.query_row(&format!("SELECT {column} FROM audits WHERE id = ?1"), [&id], |row| row.get(0)).map_err(|error| error.to_string())?;
    #[cfg(target_os = "macos")] { Command::new("open").arg(path).spawn().map_err(|error| error.to_string())?; }
    #[cfg(target_os = "windows")] { Command::new("cmd").args(["/C", "start", "", &path]).spawn().map_err(|error| error.to_string())?; }
    Ok(())
}

#[tauri::command]
fn load_comparison(app: AppHandle, id: String) -> Result<Vec<Value>, String> {
    let connection = database(&app)?;
    let path: String = connection.query_row("SELECT processed_path FROM audits WHERE id = ?1", [&id], |row| row.get(0)).map_err(|error| error.to_string())?;
    let preview = Path::new(&path).with_file_name(format!("{}_preview.tsv", Path::new(&path).file_stem().unwrap_or_default().to_string_lossy()));
    let text = fs::read_to_string(preview).map_err(|error| error.to_string())?;
    let mut lines = text.lines(); let headers: Vec<&str> = lines.next().unwrap_or_default().split('\t').collect();
    Ok(lines.map(|line| headers.iter().zip(line.split('\t')).map(|(key, value)| ((*key).to_string(), Value::String((*value).to_string()))).collect::<serde_json::Map<_, _>>().into()).collect())
}

fn main() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![choose_workbook, list_sheets, run_audit, list_archive, delete_archive, download_archive, open_archive_file, load_comparison])
        .run(tauri::generate_context!())
        .expect("error while running CONINCO Excel Auditor");
}
