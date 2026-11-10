# Universal Office Protection & VBA Password Remover (Rust)

A high-performance, full-featured Rust equivalent of `VBA Project Password Remover.py`.
Supports both CLI mode and a built-in Web Interface at `http://localhost:5000`.

---

## Supported Protections

- **VBA Projects (Excel, Word, PowerPoint)**:
  - `DPB=` &rarr; `DPx=` (Password Hash unlock)
  - `CMG=` &rarr; `CMx=` (Project Protection State - view locked projects)
  - `GC=` &rarr; `GCx=` (Project Visibility State - module/form tree visibility)
- **Microsoft Excel (`.xlsm`, `.xlsx`, `.xlam`, `.xlsb`)**:
  - Worksheet Protection (`<sheetProtection>`) on worksheets, chartsheets, dialogsheets
  - Protected Cell Ranges (`<protectedRanges>`)
  - Workbook Structure & Window Protection (`<workbookProtection>`)
  - Password to Modify / Read-Only prompt (`<fileSharing>`)
  - Optional unhide hidden and veryHidden sheets (`state="veryHidden"|"hidden"` &rarr; `state="visible"`)
- **Microsoft Word (`.docx`, `.docm`)**:
  - Restrict Editing / Document Protection (`<w:documentProtection>`)
  - Write Reservation Password (`<w:writeProtection>`)
- **Microsoft PowerPoint (`.pptx`, `.pptm`)**:
  - Presentation Modify Password (`<p:modifyVerifier>`)
- **Legacy Binary Office Files (`.xls`, `.doc`, `.ppt`)**:
  - Byte-level VBA protection replacement (`DPB`, `CMG`, `GC`)

---

## Build

```bash
cargo build --release
```

The optimized standalone binary will be available at:
`./target/release/vba_project_password_remover.exe`

---

## Usage

### 1. Web UI Mode
Launch without file arguments to start the web interface:

```bash
vba_project_password_remover
# Or with cargo:
cargo run --release
```

Navigate to **`http://localhost:5000`** in your browser. Upload any Office file, select options (all checked by default), and download the unlocked file directly.

### 2. CLI Mode

```bash
# Remove ALL protections by default (creates an automated safe backup)
vba_project_password_remover file.xlsm

# Word document editing restrictions
vba_project_password_remover document.docx

# PowerPoint modify password
vba_project_password_remover presentation.pptx

# Remove ONLY sheet protections
vba_project_password_remover file.xlsm --sheets

# Unhide all hidden / veryHidden sheets as well
vba_project_password_remover file.xlsm --unhide-sheets

# Remove all protections EXCEPT VBA
vba_project_password_remover file.xlsm --no-vba
```

### CLI Flags

| Positive Flag | Description |
|---|---|
| `--vba` | Remove all VBA project protections |
| `--dpb` | Remove only DPB (VBA Password hash) |
| `--cmg` | Remove only CMG (VBA Protection state) |
| `--gc` | Remove only GC (VBA Visibility state) |
| `--sheets` | Remove sheet protection from all worksheets |
| `--workbook` | Remove workbook structure protection |
| `--file-sharing` | Remove password to modify / read-only prompt |
| `--unhide-sheets` | Unhide all hidden and veryHidden sheets |
| `--word` | Remove Word restrict editing and write protection |
| `--ppt` | Remove PowerPoint modify password |

| Negative Flag | Description |
|---|---|
| `--no-vba` | Do not remove VBA protection |
| `--no-sheets` | Do not remove worksheet protection |
| `--no-workbook` | Do not remove workbook structure protection |
| `--no-file-sharing` | Do not remove file sharing / modify password |
| `--no-unhide` | Do not unhide hidden/veryHidden sheets |
| `--no-word` | Do not remove Word editing restrictions |
| `--no-ppt` | Do not remove PowerPoint modify protections |

---

## Testing

Run the full integration test suite:

```bash
cargo test
```
