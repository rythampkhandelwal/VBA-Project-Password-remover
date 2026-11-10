use std::collections::BTreeMap;
use std::io::{Cursor, Read, Write};
use std::path::{Path, PathBuf};
use std::sync::LazyLock;

use regex::bytes::{Captures, Regex, RegexBuilder};
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

pub const VBA_NAMES: &[&str] = &["xl/vbaProject.bin", "word/vbaProject.bin", "ppt/vbaProject.bin"];

// Regex patterns for XML-based protections
static SHEET_PROT_RE: LazyLock<Regex> = LazyLock::new(|| {
    RegexBuilder::new(r"<sheetProtection\b[^>]*?(?:/>|>.*?<\/sheetProtection>)")
        .dot_matches_new_line(true)
        .build()
        .unwrap()
});

static SHEET_RANGES_RE: LazyLock<Regex> = LazyLock::new(|| {
    RegexBuilder::new(r"<protectedRanges\b[^>]*?(?:/>|>.*?<\/protectedRanges>)")
        .dot_matches_new_line(true)
        .build()
        .unwrap()
});

static WB_PROT_RE: LazyLock<Regex> = LazyLock::new(|| {
    RegexBuilder::new(r"<workbookProtection\b[^>]*?(?:/>|>.*?<\/workbookProtection>)")
        .dot_matches_new_line(true)
        .build()
        .unwrap()
});

static FILE_SHARING_RE: LazyLock<Regex> = LazyLock::new(|| {
    RegexBuilder::new(r"<fileSharing\b[^>]*?(?:/>|>.*?<\/fileSharing>)")
        .dot_matches_new_line(true)
        .build()
        .unwrap()
});

static SHEET_HIDDEN_RE: LazyLock<Regex> = LazyLock::new(|| {
    RegexBuilder::new(r#"\bstate="(?:hidden|veryHidden)""#)
        .case_insensitive(true)
        .build()
        .unwrap()
});

static WORD_DOC_PROT_RE: LazyLock<Regex> = LazyLock::new(|| {
    RegexBuilder::new(r"<w:documentProtection\b[^>]*?(?:/>|>.*?<\/w:documentProtection>)")
        .dot_matches_new_line(true)
        .build()
        .unwrap()
});

static WORD_WRITE_PROT_RE: LazyLock<Regex> = LazyLock::new(|| {
    RegexBuilder::new(r"<w:writeProtection\b[^>]*?(?:/>|>.*?<\/w:writeProtection>)")
        .dot_matches_new_line(true)
        .build()
        .unwrap()
});

static PPT_MODIFY_RE: LazyLock<Regex> = LazyLock::new(|| {
    RegexBuilder::new(r"<p:modifyVerifier\b[^>]*?(?:/>|>.*?<\/p:modifyVerifier>)")
        .dot_matches_new_line(true)
        .build()
        .unwrap()
});

static EXCEL_SHEET_FILE_RE: LazyLock<regex::Regex> = LazyLock::new(|| {
    regex::Regex::new(r"^xl/(?:worksheets|chartsheets|dialogsheets)/.*\.xml$").unwrap()
});

static SAFE_STEM_RE: LazyLock<regex::Regex> =
    LazyLock::new(|| regex::Regex::new(r"[^A-Za-z0-9._-]").unwrap());

static SAFE_EXT_RE: LazyLock<regex::Regex> =
    LazyLock::new(|| regex::Regex::new(r"[^A-Za-z0-9]").unwrap());

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProtectionOptions {
    // VBA
    pub vba_cmg: bool,
    pub vba_dpb: bool,
    pub vba_gc: bool,

    // Excel
    pub excel_sheets: bool,
    pub excel_workbook: bool,
    pub excel_filesharing: bool,
    pub excel_unhide_sheets: bool,

    // Word
    pub word_protection: bool,

    // PowerPoint
    pub ppt_protection: bool,
}

impl Default for ProtectionOptions {
    fn default() -> Self {
        Self {
            vba_cmg: true,
            vba_dpb: true,
            vba_gc: true,
            excel_sheets: true,
            excel_workbook: true,
            excel_filesharing: true,
            excel_unhide_sheets: false,
            word_protection: true,
            ppt_protection: true,
        }
    }
}

impl ProtectionOptions {
    pub fn any_vba(&self) -> bool {
        self.vba_cmg || self.vba_dpb || self.vba_gc
    }

    pub fn any_selected(&self) -> bool {
        self.any_vba()
            || self.excel_sheets
            || self.excel_workbook
            || self.excel_filesharing
            || self.excel_unhide_sheets
            || self.word_protection
            || self.ppt_protection
    }
}

#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub struct RemovalStats {
    pub vba: BTreeMap<String, usize>,
    pub sheets_unprotected: usize,
    pub ranges_unprotected: usize,
    pub workbook_unprotected: usize,
    pub filesharing_removed: usize,
    pub sheets_unhidden: usize,
    pub word_protection_removed: usize,
    pub ppt_protection_removed: usize,
}

impl RemovalStats {
    pub fn total_changes(&self) -> usize {
        let vba_total: usize = self.vba.values().sum();
        vba_total
            + self.sheets_unprotected
            + self.ranges_unprotected
            + self.workbook_unprotected
            + self.filesharing_removed
            + self.sheets_unhidden
            + self.word_protection_removed
            + self.ppt_protection_removed
    }
}

/// Return a filesystem-friendly filename while keeping original extension.
pub fn safe_name(filename: &str) -> String {
    let base = filename.split(&['/', '\\'][..]).last().unwrap_or(filename);
    let (stem, ext) = match base.rfind('.') {
        Some(idx) => (&base[..idx], &base[idx..]),
        None => (base, ""),
    };

    let mut cleaned_stem = SAFE_STEM_RE.replace_all(stem, "_").into_owned();
    cleaned_stem = cleaned_stem.trim_start_matches('.').to_string();
    if cleaned_stem.is_empty() {
        cleaned_stem = "file".to_string();
    }

    let cleaned_ext = SAFE_EXT_RE
        .replace_all(ext.trim_start_matches('.'), "")
        .into_owned();
    let ext_str = if !cleaned_ext.is_empty() {
        format!(".{}", cleaned_ext)
    } else {
        String::new()
    };

    format!("{}{}", cleaned_stem, ext_str)
}

/// Replace selected CMG=, DPB=, GC= with CMx=, DPx=, GCx= preserving exact byte length.
pub fn patch_vba_bytes(data: &[u8], opts: &ProtectionOptions) -> (Vec<u8>, BTreeMap<String, usize>) {
    let mut counts = BTreeMap::new();
    let mut new_data = data.to_vec();
    let mut tags: Vec<(&'static [u8], &'static [u8])> = Vec::new();
    if opts.vba_cmg {
        tags.push((b"CMG=", b"CMx="));
    }
    if opts.vba_dpb {
        tags.push((b"DPB=", b"DPx="));
    }
    if opts.vba_gc {
        tags.push((b"GC=", b"GCx="));
    }

    for (old, new) in tags {
        let tag_name = std::str::from_utf8(old).unwrap_or("").to_string();
        let cnt = new_data.windows(old.len()).filter(|w| *w == old).count();
        counts.insert(tag_name, cnt);
        if cnt > 0 {
            let mut replaced = Vec::with_capacity(new_data.len());
            let mut i = 0;
            while i < new_data.len() {
                if i + old.len() <= new_data.len() && &new_data[i..i + old.len()] == old {
                    replaced.extend_from_slice(new);
                    i += old.len();
                } else {
                    replaced.push(new_data[i]);
                    i += 1;
                }
            }
            new_data = replaced;
        }
    }
    (new_data, counts)
}

fn replace_all_count(re: &Regex, input: &[u8], repl: &[u8]) -> (Vec<u8>, usize) {
    let mut count = 0;
    let new_data = re.replace_all(input, |_: &Captures| {
        count += 1;
        repl
    });
    (new_data.into_owned(), count)
}

/// Patch OOXML archive (zip format) for all requested Office protections.
pub fn process_modern_office(
    in_bytes: &[u8],
    opts: &ProtectionOptions,
) -> Result<(Vec<u8>, RemovalStats), String> {
    let mut stats = RemovalStats::default();
    let reader = Cursor::new(in_bytes);
    let mut zin = match ZipArchive::new(reader) {
        Ok(z) => z,
        Err(_) => return Err("Not a valid OOXML zip archive.".to_string()),
    };

    // Identify vbaProject file if requested
    let mut vba_name: Option<String> = None;
    if opts.any_vba() {
        for i in 0..zin.len() {
            if let Ok(file) = zin.by_index(i) {
                let name = file.name();
                if VBA_NAMES.iter().any(|v| *v == name)
                    || name.to_lowercase().ends_with("vbaproject.bin")
                {
                    vba_name = Some(name.to_string());
                    break;
                }
            }
        }
    }

    let mut bio_out = Vec::with_capacity(in_bytes.len());
    {
        let mut zout = ZipWriter::new(Cursor::new(&mut bio_out));

        for i in 0..zin.len() {
            let mut file = zin
                .by_index(i)
                .map_err(|e| format!("Failed to read zip entry: {}", e))?;
            let name = file.name().to_string();
            let is_dir = file.is_dir();
            let compression = file.compression();
            let last_modified = file.last_modified();
            let unix_permissions = file.unix_mode();

            let mut options = SimpleFileOptions::default().compression_method(compression);
            if let Some(lm) = last_modified {
                options = options.last_modified_time(lm);
            }
            if let Some(mode) = unix_permissions {
                options = options.unix_permissions(mode);
            }

            if is_dir {
                zout.add_directory(&name, options)
                    .map_err(|e| format!("Failed to add directory entry: {}", e))?;
                continue;
            }

            let mut content = Vec::with_capacity(file.size() as usize);
            file.read_to_end(&mut content)
                .map_err(|e| format!("Failed to read entry content: {}", e))?;

            // 1. VBA Project (Excel, Word, PowerPoint)
            if vba_name.as_deref() == Some(&name) {
                let (patched, counts) = patch_vba_bytes(&content, opts);
                content = patched;
                stats.vba = counts;
            }
            // 2. Excel Worksheet Protections
            else if opts.excel_sheets && EXCEL_SHEET_FILE_RE.is_match(&name) {
                let (c1, cnt_prot) = replace_all_count(&SHEET_PROT_RE, &content, b"");
                let (c2, cnt_ranges) = replace_all_count(&SHEET_RANGES_RE, &c1, b"");
                content = c2;
                stats.sheets_unprotected += cnt_prot;
                stats.ranges_unprotected += cnt_ranges;
            }
            // 3. Excel Workbook Structure / File Sharing / Unhide Sheets
            else if name == "xl/workbook.xml" {
                if opts.excel_workbook {
                    let (c, cnt_wb) = replace_all_count(&WB_PROT_RE, &content, b"");
                    content = c;
                    stats.workbook_unprotected += cnt_wb;
                }
                if opts.excel_filesharing {
                    let (c, cnt_fs) = replace_all_count(&FILE_SHARING_RE, &content, b"");
                    content = c;
                    stats.filesharing_removed += cnt_fs;
                }
                if opts.excel_unhide_sheets {
                    let (c, cnt_unhide) =
                        replace_all_count(&SHEET_HIDDEN_RE, &content, br#"state="visible""#);
                    content = c;
                    stats.sheets_unhidden += cnt_unhide;
                }
            }
            // 4. Word Document Protection & Write Protection
            else if opts.word_protection && name == "word/settings.xml" {
                let (c1, cnt_w1) = replace_all_count(&WORD_DOC_PROT_RE, &content, b"");
                let (c2, cnt_w2) = replace_all_count(&WORD_WRITE_PROT_RE, &c1, b"");
                content = c2;
                stats.word_protection_removed += cnt_w1 + cnt_w2;
            }
            // 5. PowerPoint Modify Protection
            else if opts.ppt_protection && name == "ppt/presentation.xml" {
                let (c, cnt_ppt) = replace_all_count(&PPT_MODIFY_RE, &content, b"");
                content = c;
                stats.ppt_protection_removed += cnt_ppt;
            }

            zout.start_file(&name, options)
                .map_err(|e| format!("Failed to start entry in zip: {}", e))?;
            zout.write_all(&content)
                .map_err(|e| format!("Failed to write entry to zip: {}", e))?;
        }

        zout.finish()
            .map_err(|e| format!("Failed to finalize zip archive: {}", e))?;
    }

    if stats.total_changes() == 0 {
        return Err("No matching protection locks were found in this file.".to_string());
    }

    Ok((bio_out, stats))
}

/// Patch binary legacy Office file (.xls, .doc, .ppt).
pub fn process_legacy_office(
    in_bytes: &[u8],
    opts: &ProtectionOptions,
) -> Result<(Vec<u8>, RemovalStats), String> {
    let mut stats = RemovalStats::default();
    if !opts.any_vba() {
        return Err(
            "Legacy binary files (.xls, .doc, .ppt) currently only support VBA protection removal."
                .to_string(),
        );
    }

    let (new_data, counts) = patch_vba_bytes(in_bytes, opts);
    let total: usize = counts.values().sum();
    stats.vba = counts;
    if total == 0 {
        return Err("No VBA protection keys found in binary file.".to_string());
    }
    Ok((new_data, stats))
}

/// Process office file bytes (detecting modern OOXML zip vs legacy binary).
pub fn process_office_bytes(
    data: &[u8],
    opts: &ProtectionOptions,
) -> Result<(Vec<u8>, RemovalStats), String> {
    if data.starts_with(b"PK") {
        process_modern_office(data, opts)
    } else {
        process_legacy_office(data, opts)
    }
}

/// CLI function to patch a file in place with automated backup.
pub fn patch_file(file_path: &str, opts: &ProtectionOptions) -> bool {
    let p = Path::new(file_path);
    if !p.exists() {
        println!("Error: File not found: '{}'", file_path);
        return false;
    }

    let abs_path_buf = match std::fs::canonicalize(p) {
        Ok(canon) => {
            let s = canon.to_string_lossy().to_string();
            PathBuf::from(s.strip_prefix(r"\\?\").unwrap_or(&s))
        }
        Err(_) => match std::env::current_dir() {
            Ok(cwd) => cwd.join(p),
            Err(_) => PathBuf::from(file_path),
        },
    };
    let abs_path = abs_path_buf.to_string_lossy().to_string();

    let ext = abs_path_buf
        .extension()
        .and_then(|s| s.to_str())
        .map(|s| format!(".{}", s))
        .unwrap_or_default();
    let stem = abs_path_buf
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("file");
    let parent = abs_path_buf.parent().unwrap_or(Path::new(""));

    // Determine safe backup filename
    let mut backup_path = parent.join(format!("{}_backup{}", stem, ext));
    let mut counter = 1;
    while backup_path.exists() {
        backup_path = parent.join(format!("{}_backup_{}{}", stem, counter, ext));
        counter += 1;
    }

    if let Err(e) = std::fs::copy(&abs_path_buf, &backup_path) {
        println!("[-] Error creating backup: {}", e);
        return false;
    }
    println!("[+] Backup created at: {}", backup_path.display());

    let data = match std::fs::read(&abs_path_buf) {
        Ok(d) => d,
        Err(e) => {
            println!("[-] Error reading file: {}", e);
            return false;
        }
    };

    let (modified, stats) = match process_office_bytes(&data, opts) {
        Ok(res) => res,
        Err(err) => {
            println!("[-] Notice: {}", err);
            return false;
        }
    };

    // Write patched file atomically via temp file
    let temp_path = PathBuf::from(format!("{}.tmp", abs_path));
    if let Err(e) = std::fs::write(&temp_path, &modified) {
        println!("[-] Error writing modified file: {}", e);
        return false;
    }

    if let Err(_e) = std::fs::rename(&temp_path, &abs_path_buf) {
        // Fallback for cross-device rename
        if let Err(e_copy) = std::fs::copy(&temp_path, &abs_path_buf) {
            println!("[-] Error replacing original file: {}", e_copy);
            let _ = std::fs::remove_file(&temp_path);
            return false;
        }
        let _ = std::fs::remove_file(&temp_path);
    }

    println!("\n[+] Successfully unlocked: {}", abs_path);
    println!("    Summary of removals:");
    if !stats.vba.is_empty() {
        let vba_items: Vec<String> = stats
            .vba
            .iter()
            .filter(|&(_, v)| *v > 0)
            .map(|(k, v)| format!("{}: {}", k, v))
            .collect();
        if !vba_items.is_empty() {
            println!("    - VBA keys replaced: {}", vba_items.join(", "));
        }
    }
    if stats.sheets_unprotected > 0 {
        println!(
            "    - Worksheets unprotected: {}",
            stats.sheets_unprotected
        );
    }
    if stats.ranges_unprotected > 0 {
        println!(
            "    - Cell range restrictions removed: {}",
            stats.ranges_unprotected
        );
    }
    if stats.workbook_unprotected > 0 {
        println!(
            "    - Workbook structure protection removed: {}",
            stats.workbook_unprotected
        );
    }
    if stats.filesharing_removed > 0 {
        println!(
            "    - Password to Modify / Read-Only prompt removed: {}",
            stats.filesharing_removed
        );
    }
    if stats.sheets_unhidden > 0 {
        println!(
            "    - Hidden/VeryHidden sheets made visible: {}",
            stats.sheets_unhidden
        );
    }
    if stats.word_protection_removed > 0 {
        println!(
            "    - Word Restrict Editing / Write protections removed: {}",
            stats.word_protection_removed
        );
    }
    if stats.ppt_protection_removed > 0 {
        println!(
            "    - PowerPoint Modify protections removed: {}",
            stats.ppt_protection_removed
        );
    }

    println!("\n--- Next Steps ---");
    if stats.sheets_unprotected > 0
        || stats.word_protection_removed > 0
        || stats.ppt_protection_removed > 0
    {
        println!("* Sheets / Documents / Presentations are immediately unlocked and fully editable.");
    }
    let total_vba: usize = stats.vba.values().sum();
    if total_vba > 0 {
        println!("* For VBA:");
        println!("  1. Open in Office and click OK through warning popups.");
        println!("  2. Alt + F11 > Tools > VBAProject Properties > Protection.");
        println!("  3. Set a temporary password (e.g. 1234) and save/close.");
        println!("  4. Reopen, uncheck 'Lock project for viewing', clear password, and save.\n");
    }

    true
}

pub mod web {
    use super::*;
    use axum::{
        extract::{DefaultBodyLimit, Multipart},
        http::{header, HeaderValue, StatusCode},
        response::{Html, IntoResponse, Response},
        routing::{get, post},
        Router,
    };
    use std::collections::HashSet;

    const INDEX_HTML_TEMPLATE: &str = r#"<!doctype html>
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

      {{ ERROR_PLACEHOLDER }}

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
"#;

    fn escape_html(input: &str) -> String {
        input
            .replace('&', "&amp;")
            .replace('<', "&lt;")
            .replace('>', "&gt;")
            .replace('"', "&quot;")
            .replace('\'', "&#39;")
    }

    pub fn render_html(error: Option<&str>) -> Html<String> {
        let err_block = match error {
            Some(err) if !err.trim().is_empty() => {
                format!(
                    r#"<div class="alert alert-danger">{}</div>"#,
                    escape_html(err)
                )
            }
            _ => String::new(),
        };

        let html = INDEX_HTML_TEMPLATE.replace("{{ ERROR_PLACEHOLDER }}", &err_block);
        Html(html)
    }

    async fn index_get() -> impl IntoResponse {
        render_html(None)
    }

    async fn index_post(mut multipart: Multipart) -> Response {
        let mut form_fields = HashSet::new();
        let mut file_data: Option<Vec<u8>> = None;
        let mut raw_filename: Option<String> = None;

        while let Ok(Some(field)) = multipart.next_field().await {
            let field_name = field.name().unwrap_or("").to_string();
            if field_name == "office" {
                raw_filename = field.file_name().map(|s| s.to_string());
                match field.bytes().await {
                    Ok(bytes) => {
                        file_data = Some(bytes.to_vec());
                    }
                    Err(e) => {
                        return render_html(Some(&format!("Error reading upload: {}", e)))
                            .into_response();
                    }
                }
            } else {
                form_fields.insert(field_name);
            }
        }

        let filename = match raw_filename {
            Some(ref name) if !name.trim().is_empty() => name.clone(),
            _ => {
                return render_html(Some("No file received. Please select an Office file."))
                    .into_response();
            }
        };

        let opts = ProtectionOptions {
            vba_cmg: form_fields.contains("vba_cmg"),
            vba_dpb: form_fields.contains("vba_dpb"),
            vba_gc: form_fields.contains("vba_gc"),
            excel_sheets: form_fields.contains("excel_sheets"),
            excel_workbook: form_fields.contains("excel_workbook"),
            excel_filesharing: form_fields.contains("excel_filesharing"),
            excel_unhide_sheets: form_fields.contains("excel_unhide_sheets"),
            word_protection: form_fields.contains("word_protection"),
            ppt_protection: form_fields.contains("ppt_protection"),
        };

        if !opts.any_selected() {
            return render_html(Some(
                "Please select at least one protection option to remove.",
            ))
            .into_response();
        }

        let data = match file_data {
            Some(d) if !d.is_empty() => d,
            _ => {
                return render_html(Some("Uploaded file seems empty.")).into_response();
            }
        };

        let cleaned_name = safe_name(&filename);

        match process_office_bytes(&data, &opts) {
            Ok((modified, _stats)) => {
                let content_disp = format!(r#"attachment; filename="{}""#, cleaned_name);
                let mut resp = (StatusCode::OK, modified).into_response();
                resp.headers_mut().insert(
                    header::CONTENT_TYPE,
                    HeaderValue::from_static("application/octet-stream"),
                );
                if let Ok(val) = HeaderValue::from_str(&content_disp) {
                    resp.headers_mut().insert(header::CONTENT_DISPOSITION, val);
                }
                resp
            }
            Err(err) => render_html(Some(&err)).into_response(),
        }
    }

    pub fn app() -> Router {
        Router::new()
            .route("/", get(index_get))
            .route("/", post(index_post))
            .layer(DefaultBodyLimit::max(250 * 1024 * 1024))
    }

    pub async fn run_server(port: u16) -> Result<(), Box<dyn std::error::Error>> {
        let addr = format!("0.0.0.0:{}", port);
        println!("No file argument provided. Starting web interface...");
        println!("Navigate to http://localhost:{} in your browser.", port);
        println!("CLI Usage: vba_project_password_remover <filename>");

        let listener = tokio::net::TcpListener::bind(&addr).await?;
        axum::serve(listener, app()).await?;
        Ok(())
    }
}
