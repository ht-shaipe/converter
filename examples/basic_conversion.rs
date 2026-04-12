//! Example: Basic file conversion using the file_converter library

use file_converter::{FileConverter, FileFormat};
use std::path::Path;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Initialize the converter
    let converter = FileConverter::new();

    // Example 1: Convert Word to Markdown
    println!("Example 1: Converting Word document to Markdown");
    // Note: You need an actual .docx file for this to work
    // converter.convert(Path::new("document.docx"), Path::new("output.md"))?;

    // Example 2: Convert Excel to Markdown
    println!("Example 2: Converting Excel spreadsheet to Markdown");
    // converter.convert(Path::new("data.xlsx"), Path::new("output.md"))?;

    // Example 3: Convert PDF to Markdown
    println!("Example 3: Converting PDF to Markdown");
    // converter.convert(Path::new("report.pdf"), Path::new("output.md"))?;

    // Example 4: Convert Markdown to Word
    println!("Example 4: Converting Markdown to Word document");
    // converter.convert(Path::new("readme.md"), Path::new("document.docx"))?;

    // Example 5: Convert Markdown to Excel
    println!("Example 5: Converting Markdown to Excel spreadsheet");
    // converter.convert(Path::new("tables.md"), Path::new("data.xlsx"))?;

    // Example 6: Convert Markdown to PDF
    println!("Example 6: Converting Markdown to PDF");
    // converter.convert(Path::new("notes.md"), Path::new("document.pdf"))?;

    // Example 7: Using automatic output path generation
    println!("Example 7: Using automatic output path generation");
    // let output_path = converter.convert_auto(
    //     Path::new("input.docx"),
    //     FileFormat::Markdown,
    // )?;
    // println!("Output file: {:?}", output_path);

    // Example 8: Using custom output directory
    println!("Example 8: Using custom output directory");
    let converter_with_dir = FileConverter::new()
        .with_output_dir(std::path::PathBuf::from("/tmp/converted"));
    // converter_with_dir.convert_auto(Path::new("input.docx"), FileFormat::Markdown)?;

    println!("\nAll examples completed!");
    println!("Note: The actual conversions are commented out.");
    println!("Uncomment and provide valid input files to test conversions.");

    Ok(())
}
