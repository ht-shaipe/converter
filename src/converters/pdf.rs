//! PDF converter (PDF <-> Markdown)

use crate::error::{ConverterError, Result};
use lopdf::Document;
use std::fs;
use std::path::Path;

/// Convert PDF to Markdown
pub fn pdf_to_md(input_path: &Path, output_path: &Path) -> Result<()> {
    log::info!("Converting PDF to Markdown: {:?} -> {:?}", input_path, output_path);

    // Load PDF document
    let doc = Document::load(input_path)
        .map_err(|e| ConverterError::PdfError(format!("Failed to load PDF: {}", e)))?;

    let mut markdown_content = String::new();

    // Extract text from each page
    for page_id in doc.get_pages() {
        let page_num = page_id.0;
        markdown_content.push_str(&format!("## Page {}\n\n", page_num));

        // Try to extract text content from the page
        let text = extract_text_from_page(&doc, page_id);
        if !text.is_empty() {
            markdown_content.push_str(&text);
            markdown_content.push_str("\n\n");
        }
    }

    // Write markdown content
    fs::write(output_path, markdown_content)?;

    log::info!("Successfully converted PDF to Markdown");
    Ok(())
}

/// Convert Markdown to PDF
pub fn md_to_pdf(input_path: &Path, output_path: &Path) -> Result<()> {
    log::info!("Converting Markdown to PDF: {:?} -> {:?}", input_path, output_path);

    let markdown_content = fs::read_to_string(input_path)?;

    // Convert markdown to HTML first
    let html_content = markdown_to_html(&markdown_content)?;

    // Create PDF from HTML using lopdf
    create_pdf_from_html(&html_content, output_path)?;

    log::info!("Successfully converted Markdown to PDF");
    Ok(())
}

/// Extract text from a PDF page
fn extract_text_from_page(doc: &Document, page_id: (u32, u32)) -> String {
    let mut text = String::new();
    
    // Get the page object
    if let Ok(page_obj) = doc.get_object(page_id) {
        // Try to get the Contents stream
        if let Ok(contents_obj) = get_page_contents(doc, page_id) {
            if let Ok(stream) = contents_obj.as_stream() {
                if let Ok(content_data) = stream.decompressed_data() {
                    // Parse content stream for text
                    text = parse_pdf_content_stream(&content_data);
                }
            }
        }
    }

    text
}

/// Get the contents object for a page
fn get_page_contents(doc: &Document, page_id: (u32, u32)) -> Result<&lopdf::Object> {
    let page_dict = doc.get_object(page_id)
        .map_err(|e| ConverterError::PdfError(format!("Failed to get page: {}", e)))?;
    
    if let Ok(page_dict) = page_dict.as_dict() {
        if let Ok(contents) = page_dict.get(b"Contents") {
            return Ok(contents);
        }
    }
    
    Err(ConverterError::PdfError("No contents found for page".to_string()))
}

/// Parse PDF content stream to extract text
fn parse_pdf_content_stream(content_data: &[u8]) -> String {
    let mut text = String::new();
    let content_str = String::from_utf8_lossy(content_data);
    
    // Simple text extraction - look for Tj and TJ operators
    // This is a simplified implementation; real PDF text extraction is complex
    let mut in_text_object = false;
    
    for line in content_str.lines() {
        if line.contains("BT") {
            in_text_object = true;
        } else if line.contains("ET") {
            in_text_object = false;
        }
        
        if in_text_object {
            // Extract text between parentheses for Tj operator
            if let Some(start) = line.find('(') {
                if let Some(end) = line.find(')') {
                    if start < end {
                        let extracted = &line[start + 1..end];
                        if !extracted.is_empty() {
                            text.push_str(extracted);
                            text.push(' ');
                        }
                    }
                }
            }
        }
    }

    text.trim().to_string()
}

/// Convert Markdown to HTML
fn markdown_to_html(markdown: &str) -> Result<String> {
    use pulldown_cmark::{Parser, Options, html};

    // Set up options for parsing
    let mut options = Options::empty();
    options.insert(Options::ENABLE_STRIKETHROUGH);
    options.insert(Options::ENABLE_TABLES);
    options.insert(Options::ENABLE_FOOTNOTES);

    // Parse markdown
    let parser = Parser::new_ext(markdown, options);

    // Render to HTML
    let mut html_output = String::new();
    html::push_html(&mut html_output, parser);

    Ok(html_output)
}

/// Create PDF from HTML content
fn create_pdf_from_html(html_content: &str, output_path: &Path) -> Result<()> {
    use lopdf::{Document, Object, Stream, Dictionary};
    use lopdf::Object::*;
    use lopdf::Dictionary as LoDictionary;

    // Create a new PDF document
    let mut doc = Document::with_version("1.7");

    // Create catalog
    let catalog_id = doc.new_object_id();
    
    // Create pages node
    let pages_id = doc.new_object_id();
    
    // Create font
    let font_id = doc.new_object_id();
    
    // Create content stream
    let content_id = doc.new_object_id();

    // Prepare text content - simple approach
    let mut content_bytes = Vec::new();
    
    // Add BT/ET blocks with text
    content_bytes.extend_from_slice(b"BT\n/F1 12 Tf\n50 750 Td\n");
    
    // Process HTML to extract plain text (simplified)
    let plain_text = strip_html_tags(html_content);
    
    // Split into lines and add to content
    for line in plain_text.lines().take(20) { // Limit to 20 lines for simplicity
        let escaped = escape_pdf_string(line);
        content_bytes.extend_from_slice(format!("({}) Tj\n", escaped).as_bytes());
        content_bytes.extend_from_slice(b"0 -15 Td\n");
    }
    
    content_bytes.extend_from_slice(b"ET\n");

    // Create content stream object
    let content_stream = Stream::new(
        LoDictionary::new(),
        content_bytes,
    );
    doc.set_object(content_id, content_stream);

    // Create font dictionary
    let mut font_dict = LoDictionary::new();
    font_dict.set("Type", "Font");
    font_dict.set("Subtype", "Type1");
    font_dict.set("BaseFont", "Helvetica");
    doc.set_object(font_id, font_dict);

    // Create page dictionary
    let mut page_dict = LoDictionary::new();
    page_dict.set("Type", "Page");
    page_dict.set("Parent", pages_id);
    page_dict.set("MediaBox", vec![0.into(), 0.into(), 595.into(), 842.into()]); // A4 size
    page_dict.set("Contents", content_id);
    
    let mut resources = LoDictionary::new();
    let mut fonts = LoDictionary::new();
    fonts.set("F1", font_id);
    resources.set("Font", fonts);
    page_dict.set("Resources", resources);
    
    let page_id = doc.new_object_id();
    doc.set_object(page_id, page_dict);

    // Update pages node
    let mut pages_dict = LoDictionary::new();
    pages_dict.set("Type", "Pages");
    pages_dict.set("Kids", vec![Reference(page_id)]);
    pages_dict.set("Count", 1);
    doc.set_object(pages_id, pages_dict);

    // Update catalog
    let mut catalog_dict = LoDictionary::new();
    catalog_dict.set("Type", "Catalog");
    catalog_dict.set("Pages", pages_id);
    doc.set_object(catalog_id, catalog_dict);

    // Set trailer
    doc.trailer.set("Root", catalog_id);

    // Save document
    doc.save(output_path)
        .map_err(|e| ConverterError::PdfError(format!("Failed to save PDF: {}", e)))?;

    Ok(())
}

/// Strip HTML tags from content
fn strip_html_tags(html: &str) -> String {
    let mut result = String::new();
    let mut skip = false;
    
    for c in html.chars() {
        if c == '<' {
            skip = true;
        } else if c == '>' {
            skip = false;
        } else if !skip {
            result.push(c);
        }
    }
    
    result
}

/// Escape special characters for PDF strings
fn escape_pdf_string(s: &str) -> String {
    s.replace('\\', "\\\\")
        .replace('(', "\\(")
        .replace(')', "\\)")
        .replace('\n', "\\n")
        .replace('\r', "\\r")
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::tempdir;

    #[test]
    fn test_strip_html_tags() {
        let html = "<p>Hello <strong>World</strong></p>";
        let result = strip_html_tags(html);
        assert_eq!(result, "Hello World");
    }

    #[test]
    fn test_escape_pdf_string() {
        let input = "Hello (World)\\Test";
        let result = escape_pdf_string(input);
        assert_eq!(result, "Hello \\(World\\)\\\\Test");
    }

    #[test]
    fn test_markdown_to_html() {
        let markdown = "# Hello\n\nThis is **bold** text.";
        let result = markdown_to_html(markdown).unwrap();
        assert!(result.contains("<h1>"));
        assert!(result.contains("<p>"));
    }
}
