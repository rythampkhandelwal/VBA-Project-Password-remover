use clap::Parser;
use std::process;
use vba_project_password_remover::{patch_file, web, ProtectionOptions};

#[derive(Parser, Debug)]
#[command(
    name = "vba_project_password_remover",
    about = "Universal Office Protection Remover (Excel, Word, PowerPoint, VBA)",
    after_help = "Examples:\n  vba_project_password_remover file.xlsm                  (Removes ALL protections)\n  vba_project_password_remover document.docx             (Removes Word editing restrictions)\n  vba_project_password_remover file.xlsm --sheets         (Removes ONLY sheet protection)\n  vba_project_password_remover file.xlsm --no-unhide      (Keep hidden sheets hidden)\n  vba_project_password_remover                           (Starts Web UI at http://localhost:5000)"
)]
pub struct CliArgs {
    #[arg(help = "Path to Office file (.xlsm, .xlsx, .docx, .pptx, etc.)")]
    pub file: Option<String>,

    // Positive filters
    #[arg(long, help = "Remove all VBA project protections")]
    pub vba: bool,
    #[arg(long, help = "Remove only DPB (VBA Password hash)")]
    pub dpb: bool,
    #[arg(long, help = "Remove only CMG (VBA Protection state)")]
    pub cmg: bool,
    #[arg(long, help = "Remove only GC (VBA Visibility state)")]
    pub gc: bool,
    #[arg(long, help = "Remove sheet protection from all worksheets")]
    pub sheets: bool,
    #[arg(long, help = "Remove workbook structure protection")]
    pub workbook: bool,
    #[arg(long = "file-sharing", help = "Remove password to modify / read-only prompt")]
    pub file_sharing: bool,
    #[arg(long = "unhide-sheets", help = "Unhide all hidden and veryHidden sheets")]
    pub unhide_sheets: bool,
    #[arg(long, help = "Remove Word restrict editing and write protection")]
    pub word: bool,
    #[arg(long, help = "Remove PowerPoint modify password")]
    pub ppt: bool,

    // Negative filters
    #[arg(long = "no-vba", help = "Do not remove VBA protection")]
    pub no_vba: bool,
    #[arg(long = "no-sheets", help = "Do not remove worksheet protection")]
    pub no_sheets: bool,
    #[arg(long = "no-workbook", help = "Do not remove workbook structure protection")]
    pub no_workbook: bool,
    #[arg(long = "no-file-sharing", help = "Do not remove file sharing / modify password")]
    pub no_file_sharing: bool,
    #[arg(long = "no-unhide", help = "Do not unhide hidden/veryHidden sheets")]
    pub no_unhide: bool,
    #[arg(long = "no-word", help = "Do not remove Word editing restrictions")]
    pub no_word: bool,
    #[arg(long = "no-ppt", help = "Do not remove PowerPoint modify protections")]
    pub no_ppt: bool,
}

#[tokio::main]
async fn main() {
    let args = CliArgs::parse();

    if let Some(file_path) = args.file {
        let positive_filters = [
            args.vba,
            args.dpb,
            args.cmg,
            args.gc,
            args.sheets,
            args.workbook,
            args.file_sharing,
            args.unhide_sheets,
            args.word,
            args.ppt,
        ];

        let opts = if positive_filters.iter().any(|&x| x) {
            ProtectionOptions {
                vba_cmg: args.vba || args.cmg,
                vba_dpb: args.vba || args.dpb,
                vba_gc: args.vba || args.gc,
                excel_sheets: args.sheets,
                excel_workbook: args.workbook,
                excel_filesharing: args.file_sharing,
                excel_unhide_sheets: args.unhide_sheets,
                word_protection: args.word,
                ppt_protection: args.ppt,
            }
        } else {
            ProtectionOptions {
                vba_cmg: !args.no_vba,
                vba_dpb: !args.no_vba,
                vba_gc: !args.no_vba,
                excel_sheets: !args.no_sheets,
                excel_workbook: !args.no_workbook,
                excel_filesharing: !args.no_file_sharing,
                excel_unhide_sheets: args.unhide_sheets,
                word_protection: !args.no_word,
                ppt_protection: !args.no_ppt,
            }
        };

        let success = patch_file(&file_path, &opts);
        process::exit(if success { 0 } else { 1 });
    } else {
        if let Err(e) = web::run_server(5000).await {
            eprintln!("Error starting web server: {}", e);
            process::exit(1);
        }
    }
}
