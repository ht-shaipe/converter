//! PDF converter (PDF <-> Markdown)

use crate::error::{ConverterError, Result};
use crate::formats::FileFormat;
use lopdf::Document;
use std::fs;
use std::path::Path;

/// Convert PDF to Markdown text using `lopdf` text extraction.
pub fn pdf_to_md(input_path: &Path, output_path: &Path) -> Result<()> {
    if !input_path.exists() {
        return Err(ConverterError::FileNotFound(input_path.display().to_string()));
    }
    if FileFormat::from_extension(input_path) != FileFormat::Pdf {
        return Err(ConverterError::UnsupportedFormat(
            input_path
                .extension()
                .and_then(|e| e.to_str())
                .unwrap_or("")
                .to_string(),
        ));
    }

    log::info!("Converting PDF to Markdown: {:?}", input_path);
    let doc = Document::load(input_path).map_err(|e| ConverterError::PdfError(e.to_string()))?;
    let pages: Vec<u32> = doc.get_pages().keys().copied().collect();
    let text = doc
        .extract_text(&pages)
        .map_err(|e| ConverterError::PdfError(e.to_string()))?;

    fs::write(output_path, text)?;
    Ok(())
}
