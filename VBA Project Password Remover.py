#!/usr/bin/env python3
"""
Office Protection & Password Remover (All-in-One)
Python / Flask Web UI + CLI Tool

Supported Protections:
- VBA Projects (Excel, Word, PowerPoint):
    * DPB= -> DPx= (Password Hash)
    * CMG= -> CMx= (Project Protection State)
    * GC=  -> GCx= (Project Visibility State)
- Microsoft Excel (.xlsm, .xlsx, .xlam, .xlsb):
    * Worksheet Protection (<sheetProtection>) on all worksheets, chartsheets, dialogsheets
    * Protected Ranges (<protectedRanges>)
    * Workbook Structure & Window Protection (<workbookProtection>)
    * Modify / Read-Only Password (<fileSharing>)
    * Unhide Hidden & VeryHidden Sheets (state="veryHidden" -> state="visible")
- Microsoft Word (.docx, .docm):
    * Restrict Editing / Document Protection (<w:documentProtection>)
    * Write Reservation Password (<w:writeProtection>)
- Microsoft PowerPoint (.pptx, .pptm):
    * Presentation Modify Password (<p:modifyVerifier>)
- Legacy Binary Files (.xls, .doc, .ppt):
    * Byte-level VBA protection replacement (DPB, CMG, GC)

Usage:
  CLI Mode:
    python "VBA Project Password Remover.py" file.xlsm                  (Removes ALL protections by default)
    python "VBA Project Password Remover.py" doc.docx                   (Removes Word editing restrictions)
    python "VBA Project Password Remover.py" file.xlsm --sheets         (Only unprotect worksheets)
    python "VBA Project Password Remover.py" file.xlsm --no-unhide      (Keep hidden sheets hidden)
  Web UI:
    python "VBA Project Password Remover.py"                           (Launches web interface at http://localhost:5000)
"""

import argparse
import io
import os
import re
import shutil
import sys
import zipfile
from dataclasses import dataclass
from flask import Flask, request, render_template_string, send_file

app = Flask(__name__)

# Common vbaProject paths in OpenXML packages
VBA_NAMES = ('xl/vbaProject.bin', 'word/vbaProject.bin', 'ppt/vbaProject.bin')

# Regex patterns for XML-based protections
SHEET_PROT_RE = re.compile(rb'<sheetProtection\b[^>]*?(?:\/>|>.*?<\/sheetProtection>)', re.DOTALL)
SHEET_RANGES_RE = re.compile(rb'<protectedRanges\b[^>]*?(?:\/>|>.*?<\/protectedRanges>)', re.DOTALL)
WB_PROT_RE = re.compile(rb'<workbookProtection\b[^>]*?(?:\/>|>.*?<\/workbookProtection>)', re.DOTALL)
FILE_SHARING_RE = re.compile(rb'<fileSharing\b[^>]*?(?:\/>|>.*?<\/fileSharing>)', re.DOTALL)
SHEET_HIDDEN_RE = re.compile(rb'\bstate="(?:hidden|veryHidden)"', re.IGNORECASE)

WORD_DOC_PROT_RE = re.compile(rb'<w:documentProtection\b[^>]*?(?:\/>|>.*?<\/w:documentProtection>)', re.DOTALL)
WORD_WRITE_PROT_RE = re.compile(rb'<w:writeProtection\b[^>]*?(?:\/>|>.*?<\/w:writeProtection>)', re.DOTALL)

PPT_MODIFY_RE = re.compile(rb'<p:modifyVerifier\b[^>]*?(?:\/>|>.*?<\/p:modifyVerifier>)', re.DOTALL)


@dataclass
class ProtectionOptions:
    # VBA
    vba_cmg: bool = True
    vba_dpb: bool = True
    vba_gc: bool = True

    # Excel
    excel_sheets: bool = True
    excel_workbook: bool = True
    excel_filesharing: bool = True
    excel_unhide_sheets: bool = False

    # Word
    word_protection: bool = True

    # PowerPoint
    ppt_protection: bool = True

    def any_vba(self) -> bool:
        return self.vba_cmg or self.vba_dpb or self.vba_gc

    def any_selected(self) -> bool:
        return (
            self.any_vba()
            or self.excel_sheets
            or self.excel_workbook
            or self.excel_filesharing
            or self.excel_unhide_sheets
            or self.word_protection
            or self.ppt_protection
        )


INDEX_HTML = """<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Universal Office Protection Remover</title>
  <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
  <style>
    body { background-color: #f8f9fa; }
    .card-header { font-weight: 600; }
    .tag-code { background: #e9ecef; padding: 2px 5px; border-radius: 3px; font-family: monospace; }
  </style>
</head>
<body>
<div class="container mt-4 mb-5" style="max-width: 860px;">
  <div class="card shadow-sm border-0 mb-4">
    <div class="card-body">
      <h2 class="card-title text-primary">Office Protection & Password Remover</h2>
      <p class="text-muted">Instantly remove VBA project passwords, sheet locks, workbook restrictions, Word edit protection, and PowerPoint modify passwords.</p>

      {% if error %}
        <div class="alert alert-danger">{{ error }}</div>
      {% endif %}

      <form method="post" action="/" enctype="multipart/form-data">
        <div class="form-group bg-light p-3 rounded border">
          <label class="font-weight-bold">Select Office File (Excel, Word, PowerPoint):</label>
          <input type="file" name="office" class="form-control-file" accept=".doc,.docm,.xls,.xlsm,.ppt,.pptm,.xlsb,.xlam,.xlsx,.docx" required>
        </div>

        <div class="d-flex justify-content-between align-items-center mt-4 mb-2">
          <h5 class="mb-0">Protections to Remove:</h5>
          <div>
            <button type="button" class="btn btn-sm btn-outline-primary" onclick="toggleAll(true)">Select All</button>
            <button type="button" class="btn btn-sm btn-outline-secondary" onclick="toggleAll(false)">Deselect All</button>
          </div>
        </div>
        <p class="text-muted small">All options are selected by default. Uncheck any you wish to keep.</p>

        <!-- VBA Section -->
        <div class="card mb-3">
          <div class="card-header bg-white text-dark py-2">
            <strong>1. VBA Project Protections (Excel, Word, PowerPoint)</strong>
          </div>
          <div class="card-body py-2">
            <div class="custom-control custom-checkbox mb-2">
              <input type="checkbox" class="custom-control-input prot-chk" id="vba_dpb" name="vba_dpb" value="1" checked>
              <label class="custom-control-label" for="vba_dpb">
                <strong>Password Hash (<span class="tag-code">DPB</span>)</strong> &mdash; <span class="text-muted">Unlocks password prompt in VBA editor</span>
              </label>
            </div>
            <div class="custom-control custom-checkbox mb-2">
              <input type="checkbox" class="custom-control-input prot-chk" id="vba_cmg" name="vba_cmg" value="1" checked>
              <label class="custom-control-label" for="vba_cmg">
                <strong>Protection State (<span class="tag-code">CMG</span>)</strong> &mdash; <span class="text-muted">Removes "Lock project for viewing"</span>
              </label>
            </div>
            <div class="custom-control custom-checkbox">
              <input type="checkbox" class="custom-control-input prot-chk" id="vba_gc" name="vba_gc" value="1" checked>
              <label class="custom-control-label" for="vba_gc">
                <strong>Visibility State (<span class="tag-code">GC</span>)</strong> &mdash; <span class="text-muted">Restores module/form tree visibility</span>
              </label>
            </div>
          </div>
        </div>

        <!-- Excel Section -->
        <div class="card mb-3">
          <div class="card-header bg-white text-dark py-2">
            <strong>2. Excel Specific Protections</strong>
          </div>
          <div class="card-body py-2">
            <div class="custom-control custom-checkbox mb-2">
              <input type="checkbox" class="custom-control-input prot-chk" id="excel_sheets" name="excel_sheets" value="1" checked>
              <label class="custom-control-label" for="excel_sheets">
                <strong>Worksheet & Range Protection</strong> &mdash; <span class="text-muted">Removes <span class="tag-code">&lt;sheetProtection&gt;</span> and <span class="tag-code">&lt;protectedRanges&gt;</span> from all sheets</span>
              </label>
            </div>
            <div class="custom-control custom-checkbox mb-2">
              <input type="checkbox" class="custom-control-input prot-chk" id="excel_workbook" name="excel_workbook" value="1" checked>
              <label class="custom-control-label" for="excel_workbook">
                <strong>Workbook Structure Protection</strong> &mdash; <span class="text-muted">Removes <span class="tag-code">&lt;workbookProtection&gt;</span> (add/rename/delete sheet locks)</span>
              </label>
            </div>
            <div class="custom-control custom-checkbox mb-2">
              <input type="checkbox" class="custom-control-input prot-chk" id="excel_filesharing" name="excel_filesharing" value="1" checked>
              <label class="custom-control-label" for="excel_filesharing">
                <strong>Password to Modify / Read-Only Prompt</strong> &mdash; <span class="text-muted">Removes <span class="tag-code">&lt;fileSharing&gt;</span></span>
              </label>
            </div>
            <div class="custom-control custom-checkbox">
              <input type="checkbox" class="custom-control-input prot-chk" id="excel_unhide_sheets" name="excel_unhide_sheets" value="1">
              <label class="custom-control-label" for="excel_unhide_sheets">
                <strong>Unhide All Sheets</strong> &mdash; <span class="text-muted">Converts all <em>hidden</em> and <em>veryHidden</em> sheets to visible (Optional)</span>
              </label>
            </div>
          </div>
        </div>

        <!-- Word Section -->
        <div class="card mb-3">
          <div class="card-header bg-white text-dark py-2">
            <strong>3. Word Specific Protections</strong>
          </div>
          <div class="card-body py-2">
            <div class="custom-control custom-checkbox">
              <input type="checkbox" class="custom-control-input prot-chk" id="word_protection" name="word_protection" value="1" checked>
              <label class="custom-control-label" for="word_protection">
                <strong>Word Restrict Editing & Write Protection</strong> &mdash; <span class="text-muted">Removes <span class="tag-code">&lt;w:documentProtection&gt;</span> (forms, read-only, comments) and <span class="tag-code">&lt;w:writeProtection&gt;</span></span>
              </label>
            </div>
          </div>
        </div>

        <!-- PowerPoint Section -->
        <div class="card mb-3">
          <div class="card-header bg-white text-dark py-2">
            <strong>4. PowerPoint Specific Protections</strong>
          </div>
          <div class="card-body py-2">
            <div class="custom-control custom-checkbox">
              <input type="checkbox" class="custom-control-input prot-chk" id="ppt_protection" name="ppt_protection" value="1" checked>
              <label class="custom-control-label" for="ppt_protection">
                <strong>Presentation Modify Password</strong> &mdash; <span class="text-muted">Removes <span class="tag-code">&lt;p:modifyVerifier&gt;</span></span>
              </label>
            </div>
          </div>
        </div>

        <button class="btn btn-primary btn-block btn-lg mt-3" type="submit">Unlock Document Now</button>
      </form>
    </div>
  </div>

  <div class="card border-0 bg-light p-3">
    <h5>Quick Instructions After Download:</h5>
    <ul class="mb-0">
      <li><strong>Sheets, Word &amp; PowerPoint:</strong> Immediately open and edit with no passwords requested.</li>
      <li><strong>VBA Project:</strong> Open the file, click <em>Yes/OK</em> on key warnings, press <kbd>Alt + F11</kbd> &gt; <em>Tools &gt; VBAProject Properties &gt; Protection</em>, set a temporary password, save &amp; close, then reopen and uncheck <em>"Lock project for viewing"</em>.</li>
    </ul>
  </div>
</div>

<script>
function toggleAll(state) {
  document.querySelectorAll('.prot-chk').forEach(el => el.checked = state);
}
</script>
</body>
</html>
"""

_SAFE_STEM_RE = re.compile(r'[^A-Za-z0-9._-]')
_SAFE_EXT_RE = re.compile(r'[^A-Za-z0-9]')


def safe_name(filename: str) -> str:
    """Return a filesystem-friendly filename while keeping original extension."""
    base = filename.split('/')[-1].split('\\')[-1]
    stem, ext = os.path.splitext(base)

    stem = _SAFE_STEM_RE.sub('_', stem).lstrip('.')
    if not stem:
        stem = 'file'

    cleaned_ext = _SAFE_EXT_RE.sub('', ext.lstrip('.'))
    ext = f'.{cleaned_ext}' if cleaned_ext else ''
    return f'{stem}{ext}'


def patch_vba_bytes(data: bytes, opts: ProtectionOptions) -> tuple[bytes, dict[str, int]]:
    """Replace selected CMG=, DPB=, GC= with CMx=, DPx=, GCx= preserving exact byte length."""
    counts = {}
    new_data = data
    tags = []
    if opts.vba_cmg:
        tags.append((b'CMG=', b'CMx='))
    if opts.vba_dpb:
        tags.append((b'DPB=', b'DPx='))
    if opts.vba_gc:
        tags.append((b'GC=', b'GCx='))

    for old, new in tags:
        cnt = new_data.count(old)
        counts[old.decode('ascii')] = cnt
        if cnt > 0:
            new_data = new_data.replace(old, new)
    return new_data, counts


def process_modern_office(in_bytes: bytes, opts: ProtectionOptions) -> tuple[bytes | None, str | None, dict]:
    """Patch OOXML archive (zip format) for all requested Office protections."""
    bio_in = io.BytesIO(in_bytes)
    stats = {
        'vba': {},
        'sheets_unprotected': 0,
        'ranges_unprotected': 0,
        'workbook_unprotected': 0,
        'filesharing_removed': 0,
        'sheets_unhidden': 0,
        'word_protection_removed': 0,
        'ppt_protection_removed': 0,
    }

    try:
        with zipfile.ZipFile(bio_in, 'r') as zin:
            namelist = zin.namelist()

            # Identify vbaProject file if requested
            vba_name = None
            if opts.any_vba():
                for name in namelist:
                    if name in VBA_NAMES or name.lower().endswith('vbaproject.bin'):
                        vba_name = name
                        break

            bio_out = io.BytesIO()
            with zipfile.ZipFile(bio_out, 'w') as zout:
                for item in zin.infolist():
                    content = zin.read(item.filename)

                    # 1. VBA Project (Excel, Word, PowerPoint)
                    if vba_name and item.filename == vba_name:
                        content, counts = patch_vba_bytes(content, opts)
                        stats['vba'] = counts

                    # 2. Excel Worksheet Protections
                    elif opts.excel_sheets and re.match(r'^xl/(?:worksheets|chartsheets|dialogsheets)/.*\.xml$', item.filename):
                        content, cnt_prot = SHEET_PROT_RE.subn(b'', content)
                        content, cnt_ranges = SHEET_RANGES_RE.subn(b'', content)
                        stats['sheets_unprotected'] += cnt_prot
                        stats['ranges_unprotected'] += cnt_ranges

                    # 3. Excel Workbook Structure / File Sharing / Unhide Sheets
                    elif item.filename == 'xl/workbook.xml':
                        if opts.excel_workbook:
                            content, cnt_wb = WB_PROT_RE.subn(b'', content)
                            stats['workbook_unprotected'] += cnt_wb
                        if opts.excel_filesharing:
                            content, cnt_fs = FILE_SHARING_RE.subn(b'', content)
                            stats['filesharing_removed'] += cnt_fs
                        if opts.excel_unhide_sheets:
                            content, cnt_unhide = SHEET_HIDDEN_RE.subn(rb'state="visible"', content)
                            stats['sheets_unhidden'] += cnt_unhide

                    # 4. Word Document Protection & Write Protection
                    elif opts.word_protection and item.filename == 'word/settings.xml':
                        content, cnt_w1 = WORD_DOC_PROT_RE.subn(b'', content)
                        content, cnt_w2 = WORD_WRITE_PROT_RE.subn(b'', content)
                        stats['word_protection_removed'] += (cnt_w1 + cnt_w2)

                    # 5. PowerPoint Modify Protection
                    elif opts.ppt_protection and item.filename == 'ppt/presentation.xml':
                        content, cnt_ppt = PPT_MODIFY_RE.subn(b'', content)
                        stats['ppt_protection_removed'] += cnt_ppt

                    zout.writestr(item, content)

            total_changes = (
                sum(stats['vba'].values())
                + stats['sheets_unprotected']
                + stats['ranges_unprotected']
                + stats['workbook_unprotected']
                + stats['filesharing_removed']
                + stats['sheets_unhidden']
                + stats['word_protection_removed']
                + stats['ppt_protection_removed']
            )

            if total_changes == 0:
                return None, "No matching protection locks were found in this file.", stats

            return bio_out.getvalue(), None, stats

    except zipfile.BadZipFile:
        return None, "Not a valid OOXML zip archive.", stats


def process_legacy_office(in_bytes: bytes, opts: ProtectionOptions) -> tuple[bytes | None, str | None, dict]:
    """Patch binary legacy Office file (.xls, .doc, .ppt)."""
    stats = {'vba': {}}
    if not opts.any_vba():
        return None, "Legacy binary files (.xls, .doc, .ppt) currently only support VBA protection removal.", stats

    new_data, counts = patch_vba_bytes(in_bytes, opts)
    stats['vba'] = counts
    total = sum(counts.values())
    if total == 0:
        return None, "No VBA protection keys found in binary file.", stats
    return new_data, None, stats


def patch_file(file_path: str, opts: ProtectionOptions) -> bool:
    """CLI function to patch a file in place with automated backup."""
    if not os.path.exists(file_path):
        print(f"Error: File not found: '{file_path}'")
        return False

    abs_path = os.path.abspath(file_path)
    stem, ext = os.path.splitext(abs_path)

    # Determine safe backup filename
    backup_path = f"{stem}_backup{ext}"
    counter = 1
    while os.path.exists(backup_path):
        backup_path = f"{stem}_backup_{counter}{ext}"
        counter += 1

    try:
        shutil.copy2(abs_path, backup_path)
        print(f"[+] Backup created at: {backup_path}")
    except Exception as e:
        print(f"[-] Error creating backup: {e}")
        return False

    with open(abs_path, 'rb') as f:
        data = f.read()

    if data[:2] == b'PK':
        modified, err, stats = process_modern_office(data, opts)
    else:
        modified, err, stats = process_legacy_office(data, opts)

    if err:
        print(f"[-] Notice: {err}")
        return False

    # Write patched file
    temp_path = f"{abs_path}.tmp"
    with open(temp_path, 'wb') as f:
        f.write(modified)
    os.replace(temp_path, abs_path)

    print(f"\n[+] Successfully unlocked: {abs_path}")
    print("    Summary of removals:")
    if stats.get('vba'):
        vba_summary = ', '.join(f"{k}: {v}" for k, v in stats['vba'].items() if v > 0)
        if vba_summary:
            print(f"    - VBA keys replaced: {vba_summary}")
    if stats.get('sheets_unprotected', 0) > 0:
        print(f"    - Worksheets unprotected: {stats['sheets_unprotected']}")
    if stats.get('ranges_unprotected', 0) > 0:
        print(f"    - Cell range restrictions removed: {stats['ranges_unprotected']}")
    if stats.get('workbook_unprotected', 0) > 0:
        print(f"    - Workbook structure protection removed: {stats['workbook_unprotected']}")
    if stats.get('filesharing_removed', 0) > 0:
        print(f"    - Password to Modify / Read-Only prompt removed: {stats['filesharing_removed']}")
    if stats.get('sheets_unhidden', 0) > 0:
        print(f"    - Hidden/VeryHidden sheets made visible: {stats['sheets_unhidden']}")
    if stats.get('word_protection_removed', 0) > 0:
        print(f"    - Word Restrict Editing / Write protections removed: {stats['word_protection_removed']}")
    if stats.get('ppt_protection_removed', 0) > 0:
        print(f"    - PowerPoint Modify protections removed: {stats['ppt_protection_removed']}")

    print("\n--- Next Steps ---")
    if stats.get('sheets_unprotected', 0) > 0 or stats.get('word_protection_removed', 0) > 0 or stats.get('ppt_protection_removed', 0) > 0:
        print("* Sheets / Documents / Presentations are immediately unlocked and fully editable.")
    if sum(stats.get('vba', {}).values()) > 0:
        print("* For VBA:")
        print("  1. Open in Office and click OK through warning popups.")
        print("  2. Alt + F11 > Tools > VBAProject Properties > Protection.")
        print("  3. Set a temporary password (e.g. 1234) and save/close.")
        print("  4. Reopen, uncheck 'Lock project for viewing', clear password, and save.\n")
    return True


@app.route('/', methods=['GET', 'POST'])
def index():
    error = ''
    if request.method == 'POST':
        uploaded = request.files.get('office')
        if not uploaded or uploaded.filename == '':
            error = "No file received. Please select an Office file."
            return render_template_string(INDEX_HTML, error=error)

        opts = ProtectionOptions(
            vba_cmg='vba_cmg' in request.form,
            vba_dpb='vba_dpb' in request.form,
            vba_gc='vba_gc' in request.form,
            excel_sheets='excel_sheets' in request.form,
            excel_workbook='excel_workbook' in request.form,
            excel_filesharing='excel_filesharing' in request.form,
            excel_unhide_sheets='excel_unhide_sheets' in request.form,
            word_protection='word_protection' in request.form,
            ppt_protection='ppt_protection' in request.form,
        )

        if not opts.any_selected():
            error = "Please select at least one protection option to remove."
            return render_template_string(INDEX_HTML, error=error)

        filename = safe_name(uploaded.filename)
        data = uploaded.read()
        if len(data) == 0:
            error = "Uploaded file seems empty."
            return render_template_string(INDEX_HTML, error=error)

        if data[:2] == b'PK':
            modified, err, stats = process_modern_office(data, opts)
        else:
            modified, err, stats = process_legacy_office(data, opts)

        if err:
            return render_template_string(INDEX_HTML, error=err)

        return send_file(
            io.BytesIO(modified),
            as_attachment=True,
            download_name=filename,
            mimetype='application/octet-stream'
        )

    return render_template_string(INDEX_HTML, error=error)


def parse_cli_args():
    parser = argparse.ArgumentParser(
        description="Universal Office Protection Remover (Excel, Word, PowerPoint, VBA)",
        epilog="Examples:\n"
               "  python \"VBA Project Password Remover.py\" file.xlsm                  (Removes ALL protections)\n"
               "  python \"VBA Project Password Remover.py\" document.docx             (Removes Word editing restrictions)\n"
               "  python \"VBA Project Password Remover.py\" file.xlsm --sheets         (Removes ONLY sheet protection)\n"
               "  python \"VBA Project Password Remover.py\" file.xlsm --no-unhide      (Keep hidden sheets hidden)\n"
               "  python \"VBA Project Password Remover.py\"                           (Starts Web UI at http://localhost:5000)",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument('file', nargs='?', help='Path to Office file (.xlsm, .xlsx, .docx, .pptx, etc.)')

    # Positive filters
    parser.add_argument('--vba', action='store_true', help='Remove all VBA project protections')
    parser.add_argument('--dpb', action='store_true', help='Remove only DPB (VBA Password hash)')
    parser.add_argument('--cmg', action='store_true', help='Remove only CMG (VBA Protection state)')
    parser.add_argument('--gc', action='store_true', help='Remove only GC (VBA Visibility state)')
    parser.add_argument('--sheets', action='store_true', help='Remove sheet protection from all worksheets')
    parser.add_argument('--workbook', action='store_true', help='Remove workbook structure protection')
    parser.add_argument('--file-sharing', action='store_true', help='Remove password to modify / read-only prompt')
    parser.add_argument('--unhide-sheets', action='store_true', help='Unhide all hidden and veryHidden sheets')
    parser.add_argument('--word', action='store_true', help='Remove Word restrict editing and write protection')
    parser.add_argument('--ppt', action='store_true', help='Remove PowerPoint modify password')

    # Negative filters
    parser.add_argument('--no-vba', action='store_true', help='Do not remove VBA protection')
    parser.add_argument('--no-sheets', action='store_true', help='Do not remove worksheet protection')
    parser.add_argument('--no-workbook', action='store_true', help='Do not remove workbook structure protection')
    parser.add_argument('--no-file-sharing', action='store_true', help='Do not remove file sharing / modify password')
    parser.add_argument('--no-unhide', action='store_true', help='Do not unhide hidden/veryHidden sheets')
    parser.add_argument('--no-word', action='store_true', help='Do not remove Word editing restrictions')
    parser.add_argument('--no-ppt', action='store_true', help='Do not remove PowerPoint modify protections')

    return parser.parse_args()


if __name__ == '__main__':
    args = parse_cli_args()

    if args.file:
        positive_filters = [
            args.vba, args.dpb, args.cmg, args.gc,
            args.sheets, args.workbook, args.file_sharing,
            args.unhide_sheets, args.word, args.ppt
        ]

        if any(positive_filters):
            # Explicit selection mode
            opts = ProtectionOptions(
                vba_cmg=args.vba or args.cmg,
                vba_dpb=args.vba or args.dpb,
                vba_gc=args.vba or args.gc,
                excel_sheets=args.sheets,
                excel_workbook=args.workbook,
                excel_filesharing=args.file_sharing,
                excel_unhide_sheets=args.unhide_sheets,
                word_protection=args.word,
                ppt_protection=args.ppt,
            )
        else:
            # Default: ALL protections selected (except unhide_sheets which is opt-in)
            opts = ProtectionOptions(
                vba_cmg=not args.no_vba,
                vba_dpb=not args.no_vba,
                vba_gc=not args.no_vba,
                excel_sheets=not args.no_sheets,
                excel_workbook=not args.no_workbook,
                excel_filesharing=not args.no_file_sharing,
                excel_unhide_sheets=args.unhide_sheets,
                word_protection=not args.no_word,
                ppt_protection=not args.no_ppt,
            )

        success = patch_file(args.file, opts)
        sys.exit(0 if success else 1)
    else:
        print("No file argument provided. Starting web interface...")
        print("Navigate to http://localhost:5000 in your browser.")
        print("CLI Usage: python \"VBA Project Password Remover.py\" <filename>")
        app.run(host='0.0.0.0', port=5000, debug=True)
