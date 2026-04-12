//! Word document converter (DOCX <-> Markdown)

use crate::error::{ConverterError, Result};
use std::fs;
use std::io::Read;
use std::path::Path;

/// Convert DOCX to Markdown
pub fn docx_to_md(input_path: &Path, output_path: &Path) -> Result<()> {
    log::info!("Converting DOCX to Markdown: {:?} -> {:?}", input_path, output_path);

    // Read the DOCX file as a ZIP archive
    let file = fs::File::open(input_path)?;
    let mut archive = zip::ZipArchive::new(file)
        .map_err(|e| ConverterError::WordError(format!("Failed to open DOCX: {}", e)))?;

    let mut markdown_content = String::new();

    // Try to extract and parse document.xml
    if let Ok(mut document) = archive.by_name("word/document.xml") {
        let mut xml_content = String::new();
        document.read_to_string(&mut xml_content)?;
        
        // Parse XML to extract text
        markdown_content = parse_docx_xml(&xml_content);
    } else {
        return Err(ConverterError::WordError(
            "Document.xml not found in DOCX".to_string(),
        ));
    }

    // Write markdown content
    let mut output_file = fs::File::create(output_path)?;
    output_file.write_all(markdown_content.as_bytes())?;

    log::info!("Successfully converted DOCX to Markdown");
    Ok(())
}

/// Convert Markdown to DOCX
pub fn md_to_docx(input_path: &Path, output_path: &Path) -> Result<()> {
    log::info!("Converting Markdown to DOCX: {:?} -> {:?}", input_path, output_path);

    // Read markdown content
    let markdown_content = fs::read_to_string(input_path)?;

    // Parse markdown and create DOCX structure
    let docx_content = parse_md_to_docx(&markdown_content)?;

    // Write DOCX file
    let mut output_file = fs::File::create(output_path)?;
    output_file.write_all(&docx_content)?;

    log::info!("Successfully converted Markdown to DOCX");
    Ok(())
}

/// Parse DOCX XML content to Markdown
fn parse_docx_xml(xml_content: &str) -> String {
    let mut markdown = String::new();
    let mut in_paragraph = false;
    let mut current_text = String::new();
    let mut paragraph_stack: Vec<String> = Vec::new();

    for line in xml_content.lines() {
        let trimmed = line.trim();

        // Detect paragraph tags
        if trimmed.contains("<w:p>") {
            in_paragraph = true;
            current_text.clear();
        }

        // Check for heading styles
        let is_heading = trimmed.contains("<w:pStyle w:val=\"Heading1\"") 
            || trimmed.contains("<w:pStyle w:val=\"Heading2\"")
            || trimmed.contains("<w:pStyle w:val=\"Heading3\"");
        
        let heading_level = if trimmed.contains("Heading1") {
            1
        } else if trimmed.contains("Heading2") {
            2
        } else if trimmed.contains("Heading3") {
            3
        } else {
            0
        };

        // Extract text content
        if in_paragraph {
            // Extract text from w:t tags
            extract_text_from_xml(trimmed, &mut current_text);
        }

        // End of paragraph
        if trimmed.contains("</w:p>") && in_paragraph {
            if !current_text.is_empty() {
                if heading_level > 0 {
                    // Add heading markers
                    for _ in 0..heading_level {
                        markdown.push('#');
                    }
                    markdown.push(' ');
                    markdown.push_str(&current_text);
                } else {
                    markdown.push_str(&current_text);
                }
                markdown.push('\n');
                markdown.push('\n');
            }
            in_paragraph = false;
        }
    }

    markdown
}

/// Extract text from XML tags
fn extract_text_from_xml(xml_line: &str, text_buffer: &mut String) {
    let mut start = 0;
    while let Some(tag_start) = xml_line[start..].find("<w:t") {
        let abs_start = start + tag_start;
        if let Some(tag_end) = xml_line[abs_start..].find('>') {
            let content_start = abs_start + tag_end + 1;
            if let Some(content_end) = xml_line[content_start..].find("</w:t>") {
                let text = xml_line[content_start..content_start + content_end].trim();
                if !text.is_empty() {
                    if !text_buffer.is_empty() && !text_buffer.ends_with(' ') {
                        text_buffer.push(' ');
                    }
                    text_buffer.push_str(text);
                }
                start = content_start + content_end;
            } else {
                break;
            }
        } else {
            break;
        }
    }
}

/// Parse Markdown content to DOCX binary
fn parse_md_to_docx(markdown_content: &str) -> Result<Vec<u8>> {
    use docx_rs::*;

    let mut doc = Docx::new();
    
    // Parse markdown lines and convert to DOCX paragraphs
    for line in markdown_content.lines() {
        let trimmed = line.trim();
        
        if trimmed.starts_with("# ") {
            // Heading 1
            let text = trimmed.trim_start_matches("# ").trim();
            let run = Run::new().add_text(text);
            let paragraph = Paragraph::new().add_run(run);
            doc = doc.add_paragraph(paragraph);
        } else if trimmed.starts_with("## ") {
            // Heading 2
            let text = trimmed.trim_start_matches("## ").trim();
            let run = Run::new().add_text(text);
            let paragraph = Paragraph::new().add_run(run);
            doc = doc.add_paragraph(paragraph);
        } else if trimmed.starts_with("### ") {
            // Heading 3
            let text = trimmed.trim_start_matches("### ").trim();
            let run = Run::new().add_text(text);
            let paragraph = Paragraph::new().add_run(run);
            doc = doc.add_paragraph(paragraph);
        } else if trimmed.starts_with("#### ") {
            // Heading 4
            let text = trimmed.trim_start_matches("#### ").trim();
            let run = Run::new().add_text(text);
            let paragraph = Paragraph::new().add_run(run);
            doc = doc.add_paragraph(paragraph);
        } else if trimmed.starts_with("- ") || trimmed.starts_with("* ") {
            // List item
            let text = trimmed.trim_start_matches("- ").trim_start_matches("* ").trim();
            let run = Run::new().add_text(text);
            let paragraph = Paragraph::new()
                .add_run(run)
                .set_style(ParagraphStyle::ListParagraph);
            doc = doc.add_paragraph(paragraph);
        } else if !trimmed.is_empty() {
            // Regular paragraph
            let run = Run::new().add_text(trimmed);
            let paragraph = Paragraph::new().add_run(run);
            doc = doc.add_paragraph(paragraph);
        }
    }

    // Build the DOCX file
    let mut buffer = Vec::new();
    doc.build().write(&mut buffer)
        .map_err(|e| ConverterError::WordError(format!("Failed to build DOCX: {}", e)))?;

    Ok(buffer)
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::tempdir;

    #[test]
    fn test_extract_text_from_xml() {
        let xml = "<w:t>Hello</w:t> <w:t>World</w:t>";
        let mut text = String::new();
        extract_text_from_xml(xml, &mut text);
        assert_eq!(text, "Hello World");
    }

    #[test]
    fn test_parse_docx_xml_basic() {
        let xml = r#"
            <w:document>
                <w:p><w:r><w:t>Hello World</w:t></w:r></w:p>
                <w:p><w:r><w:t>Second paragraph</w:t></w:r></w:p>
            </w:document>
        "#;
        let result = parse_docx_xml(xml);
        assert!(result.contains("Hello World"));
        assert!(result.contains("Second paragraph"));
    }
}
