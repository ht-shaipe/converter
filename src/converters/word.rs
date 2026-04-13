//! Word document converter (DOCX <-> Markdown)

use crate::error::{ConverterError, Result};
use crate::formats::FileFormat;
use std::fs::{self, File};
use std::io::Write;
use std::path::Path;

fn xml_escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

/// Convert DOCX to Markdown（标题、粗斜体、表格、内嵌图片等，逻辑与 feishu `doc-converter` 对齐）。
pub fn docx_to_md(input_path: &Path, output_path: &Path) -> Result<()> {
    if !input_path.exists() {
        return Err(ConverterError::FileNotFound(input_path.display().to_string()));
    }
    if FileFormat::from_extension(input_path) != FileFormat::Word {
        return Err(ConverterError::UnsupportedFormat(
            input_path
                .extension()
                .and_then(|e| e.to_str())
                .unwrap_or("")
                .to_string(),
        ));
    }

    log::info!("Converting DOCX to Markdown: {:?}", input_path);
    let bytes = fs::read(input_path)?;
    let md = super::docx_markdown::convert_docx_bytes_to_markdown(&bytes)?;
    fs::write(output_path, md)?;
    Ok(())
}

/// Convert Markdown to a minimal valid DOCX (ZIP package).
pub fn md_to_docx(input_path: &Path, output_path: &Path) -> Result<()> {
    if !input_path.exists() {
        return Err(ConverterError::FileNotFound(input_path.display().to_string()));
    }
    let fmt = FileFormat::from_extension(input_path);
    if fmt != FileFormat::Markdown && fmt != FileFormat::Unknown {
        return Err(ConverterError::UnsupportedFormat(
            input_path
                .extension()
                .and_then(|e| e.to_str())
                .unwrap_or("")
                .to_string(),
        ));
    }

    log::info!("Converting Markdown to DOCX: {:?}", input_path);
    let markdown_content = fs::read_to_string(input_path)?;

    let mut paragraphs = String::new();
    for line in markdown_content.lines() {
        if line.trim().is_empty() {
            continue;
        }
        let esc = xml_escape(line);
        paragraphs.push_str(&format!(
            "<w:p><w:r><w:t xml:space=\"preserve\">{esc}</w:t></w:r></w:p>"
        ));
    }

    let document_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>{paragraphs}</w:body>
</w:document>"#
    );

    let file = File::create(output_path)?;
    let mut zip = zip::ZipWriter::new(file);
    let opts = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", opts)
        .map_err(|e| ConverterError::WordError(e.to_string()))?;
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#,
    )
    .map_err(|e| ConverterError::WordError(e.to_string()))?;

    zip.start_file("_rels/.rels", opts)
        .map_err(|e| ConverterError::WordError(e.to_string()))?;
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#,
    )
    .map_err(|e| ConverterError::WordError(e.to_string()))?;

    zip.start_file("word/_rels/document.xml.rels", opts)
        .map_err(|e| ConverterError::WordError(e.to_string()))?;
    zip.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
"#,
    )
    .map_err(|e| ConverterError::WordError(e.to_string()))?;

    zip.start_file("word/document.xml", opts)
        .map_err(|e| ConverterError::WordError(e.to_string()))?;
    zip.write_all(document_xml.as_bytes())
        .map_err(|e| ConverterError::WordError(e.to_string()))?;

    zip.finish().map_err(|e| ConverterError::WordError(e.to_string()))?;
    Ok(())
}
