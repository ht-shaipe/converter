//! Example: Basic file conversion using the file_converter library

use file_converter::{docx_to_md, md_to_docx, xlsx_to_md, md_to_xlsx, pdf_to_md, md_to_pdf};
use std::path::Path;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Example 1: Convert Word to Markdown
    println!("Example 1: Converting Word document to Markdown");
    // Note: You need an actual .docx file for this to work
    // docx_to_md(Path::new("document.docx"), Path::new("output.md"))?;

    // Example 2: Convert Markdown to Word
    println!("Example 2: Converting Markdown to Word document");
    // md_to_docx(Path::new("readme.md"), Path::new("document.docx"))?;

    // Example 3: Convert Excel to Markdown
    println!("Example 3: Converting Excel spreadsheet to Markdown");
    // xlsx_to_md(Path::new("data.xlsx"), Path::new("tables.md"))?;

    // Example 4: Convert Markdown to Excel
    println!("Example 4: Converting Markdown to Excel spreadsheet");
    // md_to_xlsx(Path::new("tables.md"), Path::new("data.xlsx"))?;

    // Example 5: Convert PDF to Markdown
    println!("Example 5: Converting PDF to Markdown");
    // pdf_to_md(Path::new("report.pdf"), Path::new("notes.md"))?;

    // Example 6: Convert Markdown to PDF
    println!("Example 6: Converting Markdown to PDF");
    // md_to_pdf(Path::new("notes.md"), Path::new("report.pdf"))?;

    println!("\nAll examples completed!");
    println!("Note: The actual conversions are commented out.");
    println!("Uncomment and provide valid input files to test conversions.");

    Ok(())
}
