use std::io::{Cursor, Read, Write};
use vba_project_password_remover::{
    patch_vba_bytes, process_legacy_office, process_modern_office, safe_name, web,
    ProtectionOptions,
};
use zip::write::SimpleFileOptions;
use zip::{ZipArchive, ZipWriter};

#[test]
fn test_safe_name_variations() {
    assert_eq!(safe_name("file.xlsm"), "file.xlsm");
    assert_eq!(safe_name("a/b/c/document.docx"), "document.docx");
    assert_eq!(safe_name(r"C:\Users\User\my presentation.pptx"), "my_presentation.pptx");
    assert_eq!(safe_name("special!@#$%^&*()_+name.xls"), "special____________name.xls");
    assert_eq!(safe_name("...leading.dots.xlsx"), "leading.dots.xlsx");
    assert_eq!(safe_name(".onlyext"), "file.onlyext");
    assert_eq!(safe_name("noext"), "noext");
    assert_eq!(safe_name(r"bad:name?.docm"), "bad_name_.docm");
}

#[test]
fn test_patch_vba_bytes() {
    let original = b"header_data CMG=\"test_cmg\" DPB=\"1234567890ABCDEF\" GC=\"view_gc\" footer";
    let opts = ProtectionOptions::default();

    let (patched, counts) = patch_vba_bytes(original, &opts);
    assert_eq!(counts.get("CMG="), Some(&1));
    assert_eq!(counts.get("DPB="), Some(&1));
    assert_eq!(counts.get("GC="), Some(&1));
    // CMG= (4->4), DPB= (4->4), GC= (3->4) -> original.len() + 1
    assert_eq!(patched.len(), original.len() + 1);

    assert!(patched.windows(4).any(|w| w == b"CMx="));
    assert!(patched.windows(4).any(|w| w == b"DPx="));
    assert!(patched.windows(4).any(|w| w == b"GCx="));
    assert!(!patched.windows(4).any(|w| w == b"CMG="));
    assert!(!patched.windows(4).any(|w| w == b"DPB="));
    assert!(!patched.windows(3).any(|w| w == b"GC="));

    // Selective: only DPB and CMG (preserving exact byte length)
    let selective_opts = ProtectionOptions {
        vba_cmg: true,
        vba_dpb: true,
        vba_gc: false,
        ..Default::default()
    };
    let (sel_patched, sel_counts) = patch_vba_bytes(original, &selective_opts);
    assert_eq!(sel_counts.get("DPB="), Some(&1));
    assert_eq!(sel_counts.get("CMG="), Some(&1));
    assert_eq!(sel_patched.len(), original.len());
    assert!(sel_patched.windows(4).any(|w| w == b"DPx="));
    assert!(sel_patched.windows(4).any(|w| w == b"CMx="));
    assert!(sel_patched.windows(3).any(|w| w == b"GC="));
}

#[test]
fn test_legacy_office_vba() {
    let mut data = vec![0u8; 1024];
    data[100..104].copy_from_slice(b"DPB=");
    data[104..110].copy_from_slice(b"123456");
    data[200..204].copy_from_slice(b"CMG=");
    data[204..210].copy_from_slice(b"ABCDEF");

    let opts = ProtectionOptions::default();
    let (result, stats) = process_legacy_office(&data, &opts).expect("Should succeed");
    assert_eq!(stats.vba.get("DPB="), Some(&1));
    assert_eq!(stats.vba.get("CMG="), Some(&1));
    assert_eq!(stats.vba.get("GC="), Some(&0));
    assert_eq!(result.len(), data.len());
    assert_eq!(&result[100..104], b"DPx=");
    assert_eq!(&result[200..204], b"CMx=");

    // When no keys exist
    let empty_data = vec![0u8; 1024];
    let err = process_legacy_office(&empty_data, &opts).unwrap_err();
    assert_eq!(err, "No VBA protection keys found in binary file.");
}

fn create_test_ooxml_zip() -> Vec<u8> {
    let mut buf = Vec::new();
    {
        let mut writer = ZipWriter::new(Cursor::new(&mut buf));
        let opt = SimpleFileOptions::default();

        // 1. VBA project
        writer.start_file("xl/vbaProject.bin", opt).unwrap();
        writer.write_all(b"binary_start DPB=\"abc\" CMG=\"def\" GC=\"ghi\" binary_end").unwrap();

        // 2. Worksheet with protection and ranges
        writer.start_file("xl/worksheets/sheet1.xml", opt).unwrap();
        writer.write_all(br#"<?xml version="1.0" encoding="UTF-8"?><worksheet><sheetViews/><sheetProtection sheet="1" password="1234"/><protectedRanges><protectedRange name="Range1"/></protectedRanges><sheetData/></worksheet>"#).unwrap();

        // 3. Dialogsheet with protection
        writer.start_file("xl/dialogsheets/sheet1.xml", opt).unwrap();
        writer.write_all(br#"<dialogsheet><sheetProtection sheet="1"/></dialogsheet>"#).unwrap();

        // 4. Workbook with workbookProtection, fileSharing, and hidden sheets
        writer.start_file("xl/workbook.xml", opt).unwrap();
        writer.write_all(br#"<workbook><workbookProtection lockStructure="1"/><fileSharing reservationPassword="pass"/><sheets><sheet name="S1" state="veryHidden"/><sheet name="S2" state="hidden"/><sheet name="S3" state="visible"/></sheets></workbook>"#).unwrap();

        // 5. Word settings
        writer.start_file("word/settings.xml", opt).unwrap();
        writer.write_all(br#"<w:settings><w:documentProtection w:edit="readOnly" w:enforcement="1"/><w:writeProtection w:recommended="1"/></w:settings>"#).unwrap();

        // 6. PowerPoint presentation
        writer.start_file("ppt/presentation.xml", opt).unwrap();
        writer.write_all(br#"<p:presentation><p:modifyVerifier p:algorithmName="SHA-512"/></p:presentation>"#).unwrap();

        writer.finish().unwrap();
    }
    buf
}

#[test]
fn test_modern_office_all_protections() {
    let zip_bytes = create_test_ooxml_zip();
    let opts = ProtectionOptions {
        excel_unhide_sheets: true,
        ..Default::default()
    };

    let (modified_bytes, stats) = process_modern_office(&zip_bytes, &opts).expect("Processing should succeed");

    assert_eq!(stats.sheets_unprotected, 2);
    assert_eq!(stats.ranges_unprotected, 1);
    assert_eq!(stats.workbook_unprotected, 1);
    assert_eq!(stats.filesharing_removed, 1);
    assert_eq!(stats.sheets_unhidden, 2);
    assert_eq!(stats.word_protection_removed, 2);
    assert_eq!(stats.ppt_protection_removed, 1);
    assert_eq!(stats.vba.get("DPB="), Some(&1));
    assert_eq!(stats.vba.get("CMG="), Some(&1));
    assert_eq!(stats.vba.get("GC="), Some(&1));

    // Verify modified zip content
    let mut reader = ZipArchive::new(Cursor::new(&modified_bytes)).unwrap();

    // Check worksheet
    let mut ws = String::new();
    reader.by_name("xl/worksheets/sheet1.xml").unwrap().read_to_string(&mut ws).unwrap();
    assert!(!ws.contains("sheetProtection"));
    assert!(!ws.contains("protectedRanges"));

    // Check workbook
    let mut wb = String::new();
    reader.by_name("xl/workbook.xml").unwrap().read_to_string(&mut wb).unwrap();
    assert!(!wb.contains("workbookProtection"));
    assert!(!wb.contains("fileSharing"));
    assert!(!wb.contains("veryHidden"));
    assert!(!wb.contains(r#"state="hidden""#));
    assert_eq!(wb.matches(r#"state="visible""#).count(), 3);

    // Check word
    let mut word = String::new();
    reader.by_name("word/settings.xml").unwrap().read_to_string(&mut word).unwrap();
    assert!(!word.contains("documentProtection"));
    assert!(!word.contains("writeProtection"));

    // Check ppt
    let mut ppt = String::new();
    reader.by_name("ppt/presentation.xml").unwrap().read_to_string(&mut ppt).unwrap();
    assert!(!ppt.contains("modifyVerifier"));

    // Check vba
    let mut vba = Vec::new();
    reader.by_name("xl/vbaProject.bin").unwrap().read_to_end(&mut vba).unwrap();
    assert!(vba.windows(4).any(|w| w == b"DPx="));
    assert!(vba.windows(4).any(|w| w == b"CMx="));
    assert!(vba.windows(4).any(|w| w == b"GCx="));
    assert!(!vba.windows(4).any(|w| w == b"DPB="));
}

#[test]
fn test_modern_office_selective() {
    let zip_bytes = create_test_ooxml_zip();
    let opts = ProtectionOptions {
        vba_cmg: false,
        vba_dpb: false,
        vba_gc: false,
        excel_sheets: true,
        excel_workbook: false,
        excel_filesharing: false,
        excel_unhide_sheets: false,
        word_protection: false,
        ppt_protection: false,
    };

    let (modified_bytes, stats) = process_modern_office(&zip_bytes, &opts).expect("Processing should succeed");
    assert_eq!(stats.sheets_unprotected, 2);
    assert_eq!(stats.ranges_unprotected, 1);
    assert_eq!(stats.workbook_unprotected, 0);
    assert_eq!(stats.word_protection_removed, 0);

    let mut reader = ZipArchive::new(Cursor::new(&modified_bytes)).unwrap();

    // Workbook protection should still be present
    let mut wb = String::new();
    reader.by_name("xl/workbook.xml").unwrap().read_to_string(&mut wb).unwrap();
    assert!(wb.contains("workbookProtection"));

    // VBA protection should still be present
    let mut vba = Vec::new();
    reader.by_name("xl/vbaProject.bin").unwrap().read_to_end(&mut vba).unwrap();
    assert!(vba.windows(4).any(|w| w == b"DPB="));
}

#[test]
fn test_no_protection_error() {
    let mut buf = Vec::new();
    {
        let mut writer = ZipWriter::new(Cursor::new(&mut buf));
        writer.start_file("clean.txt", SimpleFileOptions::default()).unwrap();
        writer.write_all(b"clean text").unwrap();
        writer.finish().unwrap();
    }
    let opts = ProtectionOptions::default();
    let err = process_modern_office(&buf, &opts).unwrap_err();
    assert_eq!(err, "No matching protection locks were found in this file.");
}

#[test]
fn test_invalid_zip_error() {
    let opts = ProtectionOptions::default();
    let err = process_modern_office(b"PK_corrupted_not_a_zip", &opts).unwrap_err();
    assert_eq!(err, "Not a valid OOXML zip archive.");
}

#[tokio::test]
async fn test_web_get_index() {
    use axum::body::to_bytes;
    use axum::http::{Request, StatusCode};
    use tower::ServiceExt;

    let app = web::app();
    let response = app
        .oneshot(Request::builder().uri("/").body(axum::body::Body::empty()).unwrap())
        .await
        .unwrap();

    assert_eq!(response.status(), StatusCode::OK);
    let body = to_bytes(response.into_body(), usize::MAX).await.unwrap();
    let html = String::from_utf8(body.to_vec()).unwrap();
    assert!(html.contains("Office Protection &amp; Password Remover") || html.contains("Office Protection & Password Remover"));
    assert!(html.contains("name=\"office\""));
    assert!(html.contains("name=\"vba_dpb\""));
}

#[tokio::test]
async fn test_web_post_upload_success() {
    use axum::body::to_bytes;
    use axum::http::{header, Request, StatusCode};
    use tower::ServiceExt;

    let app = web::app();
    let zip_data = create_test_ooxml_zip();

    let boundary = "------------------------boundary123";
    let mut body_bytes = Vec::new();
    // Office file field
    body_bytes.extend_from_slice(format!("--{}\r\n", boundary).as_bytes());
    body_bytes.extend_from_slice(b"Content-Disposition: form-data; name=\"office\"; filename=\"my_test.xlsm\"\r\n");
    body_bytes.extend_from_slice(b"Content-Type: application/vnd.ms-excel.sheet.macroEnabled.12\r\n\r\n");
    body_bytes.extend_from_slice(&zip_data);
    body_bytes.extend_from_slice(b"\r\n");

    // Form checkboxes
    for field in &["vba_dpb", "vba_cmg", "vba_gc", "excel_sheets", "excel_workbook", "excel_filesharing"] {
        body_bytes.extend_from_slice(format!("--{}\r\n", boundary).as_bytes());
        body_bytes.extend_from_slice(format!("Content-Disposition: form-data; name=\"{}\"\r\n\r\n1\r\n", field).as_bytes());
    }
    body_bytes.extend_from_slice(format!("--{}--\r\n", boundary).as_bytes());

    let req = Request::builder()
        .method("POST")
        .uri("/")
        .header(header::CONTENT_TYPE, format!("multipart/form-data; boundary={}", boundary))
        .body(axum::body::Body::from(body_bytes))
        .unwrap();

    let response = app.oneshot(req).await.unwrap();
    assert_eq!(response.status(), StatusCode::OK);

    let content_type = response.headers().get(header::CONTENT_TYPE).unwrap().to_str().unwrap();
    assert_eq!(content_type, "application/octet-stream");

    let disp = response.headers().get(header::CONTENT_DISPOSITION).unwrap().to_str().unwrap();
    assert!(disp.contains("attachment; filename=\"my_test.xlsm\""));

    let body = to_bytes(response.into_body(), usize::MAX).await.unwrap();
    assert!(body.starts_with(b"PK"));
}
