# Architecture

CONINCO Excel Auditor is an offline desktop application. React renders the
interface, Tauri 2 provides the native desktop shell and SQLite archive, and
the Python engine performs workbook reading, classification, calculation, and
report generation.

```text
React + TypeScript + Tailwind
            |
            v
       Tauri 2 + SQLite
            |
            | local JSON over stdin/stdout
            v
       Python Excel engine
            |
            v
   calculated workbook + comparison data
```

The Process screen selects a workbook and sheet, then invokes the Python
engine. The result is stored in the local archive and displayed in the
Comparison screen. The Archive screen keeps submitted and processed files
together and allows the user to open, download, or delete them.

