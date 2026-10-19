# Hướng dẫn đóng gói CONINCO Excel Auditor trên Windows

Tài liệu này dành cho người có máy Windows để build bộ cài `.exe` cho ứng dụng CONINCO Excel Auditor.

## 1. Yêu cầu máy build

Nên dùng Windows 10 hoặc Windows 11 64-bit. Máy build cần có Internet trong lần đầu cài thư viện.

Cần cài:

1. Python 3.12 hoặc 3.13.
2. Node.js bản LTS.
3. Rust với toolchain `stable-msvc`.
4. Visual Studio Build Tools với workload `Desktop development with C++`.
5. Microsoft Edge WebView2 Runtime.

### Link tải

- Python: https://www.python.org/downloads/windows/
- Node.js LTS: https://nodejs.org/en/download
- Rust: https://www.rust-lang.org/tools/install
- Visual Studio Build Tools: https://visualstudio.microsoft.com/visual-cpp-build-tools/
- WebView2 Runtime: https://developer.microsoft.com/en-us/microsoft-edge/webview2/

## 2. Cài Python

Trong trình cài Python, chọn:

```text
Add python.exe to PATH
```

Sau khi cài, mở PowerShell mới:

```powershell
python --version
pip --version
```

## 3. Cài Node.js

Cài Node.js bản LTS theo tùy chọn mặc định. Sau đó mở PowerShell mới:

```powershell
node --version
npm --version
```

Không cần cài React hoặc Tailwind riêng. `npm install` sẽ tự cài chúng từ `package.json`.

## 4. Cài Rust

Mở PowerShell Administrator:

```powershell
winget install --id Rustlang.Rustup
```

Đóng PowerShell, mở lại:

```powershell
rustup default stable-msvc
rustc --version
cargo --version
```

## 5. Cài Visual Studio C++ Build Tools

Mở trình cài Visual Studio Build Tools và chọn workload:

```text
Desktop development with C++
```

Đảm bảo Windows SDK và MSVC C++ build tools được chọn.

## 6. Cài WebView2

Windows 10 mới và Windows 11 thường đã có WebView2. Nếu chưa có, tải Evergreen Bootstrapper từ:

https://developer.microsoft.com/en-us/microsoft-edge/webview2/

## 7. Chuẩn bị source code

Giải nén thư mục `Coninco_excel_desktop` vào ví dụ:

```text
C:\Projects\Coninco_excel_desktop
```

Mở PowerShell:

```powershell
cd C:\Projects\Coninco_excel_desktop\desktop-app
```

Kiểm tra có các file:

```text
package.json
package-lock.json
scripts\build-engine.ps1
src-tauri\Cargo.toml
src-tauri\tauri.conf.json
```

## 8. Cài thư viện React/Tauri

```powershell
npm install
```

Nếu PowerShell chặn script:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
npm install
```

## 9. Đóng gói Python engine thành .exe

Trong thư mục `desktop-app`:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\scripts\build-engine.ps1
```

Script sẽ tự động tạo virtual environment, cài pandas, openpyxl, xlrd, PyInstaller và build Python engine.

Kiểm tra kết quả:

```powershell
Test-Path .\src-tauri\engine\excel-auditor-engine.exe
```

Kết quả phải là:

```text
True
```

## 10. Chạy thử desktop

```powershell
npm run tauri:dev
```

Kiểm tra:

1. Chọn file `.xls` hoặc `.xlsx`.
2. Chọn sheet.
3. Bấm `RUN AUDIT`.
4. Vào `Archive`.
5. Mở `VIEW COMPARISON`.
6. Tải file gốc và file đã xử lý.
7. Xóa archive và kiểm tra hộp xác nhận.

## 11. Tạo bộ cài Windows

```powershell
npm run tauri:build -- --bundles nsis
```

File cài đặt nằm tại:

```text
src-tauri\target\release\bundle\nsis\
```

Tên thường có dạng:

```text
CONINCO Excel Auditor_0.2.0_x64-setup.exe
```

Đây là file duy nhất cần gửi cho người dùng cuối.

## 12. Máy người dùng cuối

Installer đã đóng gói sẵn:

- Giao diện React.
- Tauri runtime.
- Python engine.
- pandas.
- openpyxl.
- xlrd.
- SQLite.

Máy người dùng cuối không cần cài Python, Node.js, Rust hoặc npm.

WebView2 thường có sẵn trên Windows 10 mới và Windows 11. Nếu chưa có, installer có thể cần Internet để tải runtime.

## 13. Lỗi thường gặp

### Python không được nhận diện

Cài lại Python và chọn `Add python.exe to PATH`, sau đó mở PowerShell mới.

### npm không được nhận diện

Đóng PowerShell, mở lại sau khi cài Node.js.

### Thiếu linker hoặc MSVC

Mở Visual Studio Installer, chọn `Modify`, sau đó cài:

```text
Desktop development with C++
```

### Không tạo được Python engine

Chạy lại:

```powershell
cd C:\Projects\Coninco_excel_desktop\desktop-app
.\scripts\build-engine.ps1
```

### App không chạy sau khi cài

Kiểm tra:

1. Có file `src-tauri\engine\excel-auditor-engine.exe` trước khi build.
2. Đã cài WebView2.
3. Đã build lại sau khi thay đổi source code.

