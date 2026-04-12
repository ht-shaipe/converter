//! File conversion implementations

use crate::error::{ConverterError, Result};
use crate::formats::{ConversionType, FileFormat};
use std::fs;
use std::path::{Path, PathBuf};

/// Word document converter (DOCX <-> Markdown)
pub mod word;
/// Excel spreadsheet converter (XLSX <-> Markdown)
pub mod excel;
/// PDF converter (PDF <-> Markdown)
pub mod pdf;

/// Main file converter struct
#[derive(Debug, Default)]
pub struct FileConverter {
    /// Output directory (optional, defaults to same directory as input)
    output_dir: Option<PathBuf>,
}

impl FileConverter {
    /// Create a new file converter
    pub fn new() -> Self {
        Self { output_dir: None }
    }

    /// Set the output directory for converted files
    pub fn with_output_dir(mut self, output_dir: PathBuf) -> Self {
        self.output_dir = Some(output_dir);
        self
    }

    /// Convert a file from one format to another
    pub fn convert(&self, input_path: &Path, output_path: &Path) -> Result<()> {
        let source_format = FileFormat::from_path(input_path)?;
        let target_format = FileFormat::from_extension(output_path);

        if target_format == FileFormat::Unknown {
            return Err(ConverterError::UnsupportedFormat(
                output_path.display().to_string(),
            ));
        }

        let conversion_type = ConversionType::new(source_format.clone(), target_format.clone());

        if !conversion_type.is_supported() {
            return Err(ConverterError::UnsupportedConversion {
                from: source_format.to_string(),
                to: target_format.to_string(),
            });
        }

        // Ensure output directory exists
        if let Some(parent) = output_path.parent() {
            fs::create_dir_all(parent)?;
        }

        // Perform the conversion based on formats
        match (source_format, target_format) {
            (FileFormat::Word, FileFormat::Markdown) => {
                word::docx_to_md(input_path, output_path)
            }
            (FileFormat::Markdown, FileFormat::Word) => {
                word::md_to_docx(input_path, output_path)
            }
            (FileFormat::Excel, FileFormat::Markdown) => {
                excel::xlsx_to_md(input_path, output_path)
            }
            (FileFormat::Markdown, FileFormat::Excel) => {
                excel::md_to_xlsx(input_path, output_path)
            }
            (FileFormat::Pdf, FileFormat::Markdown) => {
                pdf::pdf_to_md(input_path, output_path)
            }
            (FileFormat::Markdown, FileFormat::Pdf) => {
                pdf::md_to_pdf(input_path, output_path)
            }
            _ => Err(ConverterError::UnsupportedConversion {
                from: source_format.to_string(),
                to: target_format.to_string(),
            }),
        }
    }

    /// Convert a file with automatic output path generation
    pub fn convert_auto(&self, input_path: &Path, target_format: FileFormat) -> Result<PathBuf> {
        let source_format = FileFormat::from_path(input_path)?;
        
        if !source_format.can_convert_to(&target_format) {
            return Err(ConverterError::UnsupportedConversion {
                from: source_format.to_string(),
                to: target_format.to_string(),
            });
        }

        // Generate output path
        let output_path = self.generate_output_path(input_path, &target_format)?;
        
        // Perform conversion
        self.convert(input_path, &output_path)?;
        
        Ok(output_path)
    }

    /// Generate output path based on input file and target format
    fn generate_output_path(&self, input_path: &Path, target_format: &FileFormat) -> Result<PathBuf> {
        let stem = input_path
            .file_stem()
            .and_then(|s| s.to_str())
            .ok_or_else(|| ConverterError::InvalidPath("Invalid filename".to_string()))?;

        let output_dir = self.output_dir.clone().unwrap_or_else(|| {
            input_path
                .parent()
                .unwrap_or_else(|| Path::new("."))
                .to_path_buf()
        });

        let filename = format!("{}.{}", stem, target_format.extension());
        Ok(output_dir.join(filename))
    }
}

/// Word document conversion functions
pub mod word {
    use super::*;
    use std::io::{Read, Write};

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
            
            // Simple XML parsing to extract text
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

        for line in xml_content.lines() {
            let trimmed = line.trim();

            // Detect paragraph tags
            if trimmed.starts_with("<w:p>") || trimmed.contains("<w:p>") {
                in_paragraph = true;
                current_text.clear();
            }

            // Extract text content
            if in_paragraph {
                // Simple text extraction (this is a simplified version)
                if let Some(start) = trimmed.find("<w:t>") {
                    if let Some(end) = trimmed.find("</w:t>") {
                        let text_start = start + 5;
                        let text = trimmed[text_start..end].trim();
                        if !text.is_empty() {
                            if !current_text.is_empty() {
                                current_text.push(' ');
                            }
                            current_text.push_str(text);
                        }
                    }
                }
            }

            // End of paragraph
            if trimmed.contains("</w:p>") && in_paragraph {
                if !current_text.is_empty() {
                    markdown.push_str(&current_text);
                    markdown.push('\n');
                    markdown.push('\n');
                }
                in_paragraph = false;
            }
        }

        // Handle heading detection (simplified)
        detect_headings(&mut markdown);

        markdown
    }

    /// Detect and convert headings in markdown
    fn detect_headings(markdown: &mut String) {
        // This is a placeholder for more sophisticated heading detection
        // In a real implementation, you would analyze the DOCX structure
        // to identify heading styles
    }

    /// Parse Markdown content to DOCX binary
    fn parse_md_to_docx(markdown_content: &str) -> Result<Vec<u8>> {
        // This is a simplified implementation
        // A full implementation would use docx-rs library properly
        
        use docx_rs::*;

        let mut doc = Docx::new();
        
        // Parse markdown lines and convert to DOCX paragraphs
        for line in markdown_content.lines() {
            if line.starts_with("# ") {
                // Heading 1
                let text = line.trim_start_matches("# ").trim();
                doc = doc.add_paragraph(Paragraph::new().add_run(Run::new().add_text(text)));
            } else if line.starts_with("## ") {
                // Heading 2
                let text = line.trim_start_matches("## ").trim();
                doc = doc.add_paragraph(Paragraph::new().add_run(Run::new().add_text(text)));
            } else if line.starts_with("### ") {
                // Heading 3
                let text = line.trim_start_matches("### ").trim();
                doc = doc.add_paragraph(Paragraph::new().add_run(Run::new().add_text(text)));
            } else if !line.is_empty() {
                // Regular paragraph
                doc = doc.add_paragraph(Paragraph::new().add_run(Run::new().add_text(line)));
            }
        }

        // Build the DOCX file
        let mut buffer = Vec::new();
        doc.build().write(&mut buffer)
            .map_err(|e| ConverterError::WordError(format!("Failed to build DOCX: {}", e)))?;

        Ok(buffer)
    }
}

/// Excel spreadsheet conversion functions
pub mod excel {
    use super::*;
    use calamine::{Reader, Xlsx};

    /// Convert XLSX to Markdown
    pub fn xlsx_to_md(input_path: &Path, output_path: &Path) -> Result<()> {
        log::info!("Converting XLSX to Markdown: {:?} -> {:?}", input_path, output_path);

        let file = fs::File::open(input_path)?;
        let mut workbook: Xlsx<_> = Xlsx::new(file)
            .map_err(|e| ConverterError::ExcelError(format!("Failed to open XLSX: {}", e)))?;

        let mut markdown_content = String::new();

        // Iterate through all sheets
        for sheet_name in workbook.sheet_names().to_owned() {
            markdown_content.push_str(&format!("# {}\n\n", sheet_name));

            if let Ok(range) = workbook.worksheet_range(&sheet_name) {
                if range.is_empty() {
                    continue;
                }

                // Convert to markdown table
                let mut table = String::new();
                let mut headers = true;

                for row in range.rows() {
                    let cells: Vec<String> = row
                        .iter()
                        .map(|cell| match cell {
                            calamine::Data::String(s) => s.clone(),
                            calamine::Data::Int(i) => i.to_string(),
                            calamine::Data::Float(f) => f.to_string(),
                            calamine::Data::Bool(b) => b.to_string(),
                            calamine::Data::DateTime(dt) => dt.to_string(),
                            _ => String::new(),
                        })
                        .collect();

                    if !cells.is_empty() {
                        table.push('|');
                        table.push_str(&cells.join("|"));
                        table.push_str("|\n");

                        // Add header separator after first row
                        if headers {
                            table.push('|');
                            for _ in &cells {
                                table.push_str("---|");
                            }
                            table.push('\n');
                            headers = false;
                        }
                    }
                }

                markdown_content.push_str(&table);
                markdown_content.push_str("\n\n");
            }
        }

        // Write markdown content
        fs::write(output_path, markdown_content)?;

        log::info!("Successfully converted XLSX to Markdown");
        Ok(())
    }

    /// Convert Markdown to XLSX
    pub fn md_to_xlsx(input_path: &Path, output_path: &Path) -> Result<()> {
        log::info!("Converting Markdown to XLSX: {:?} -> {:?}", input_path, output_path);

        let markdown_content = fs::read_to_string(input_path)?;
        
        // Parse markdown tables and create Excel workbook
        // This is a simplified implementation
        // A full implementation would need proper markdown table parsing
        
        use calamine::{Writer, Xlsx as CalamineXlsx};
        use std::io::BufWriter;

        let file = fs::File::create(output_path)?;
        let mut writer = Writer::new(BufWriter::new(file));

        // For now, create a simple single-sheet workbook
        // In a real implementation, you would parse markdown tables
        let mut data: Vec<Vec<String>> = Vec::new();
        
        // Parse markdown content for tables (simplified)
        let mut current_sheet = Vec::new();
        for line in markdown_content.lines() {
            if line.starts_with('#') {
                // Sheet name or header
                if !current_sheet.is_empty() {
                    writer.write_sheet(current_sheet.into_iter())?;
                    current_sheet = Vec::new();
                }
            } else if line.starts_with('|') {
                // Table row
                let cells: Vec<String> = line
                    .trim_matches('|')
                    .split('|')
                    .map(|s| s.trim().to_string())
                    .filter(|s| !s.is_empty() && !s.chars().all(|c| c == '-'))
                    .collect();
                
                if !cells.is_empty() {
                    current_sheet.push(cells);
                }
            }
        }

        // Write last sheet
        if !current_sheet.is_empty() {
            writer.write_sheet(current_sheet.into_iter())?;
        }

        writer.close()?;

        log::info!("Successfully converted Markdown to XLSX");
        Ok(())
    }
}

/// PDF conversion functions
pub mod pdf {
    use super::*;

    /// Convert PDF to Markdown
    pub fn pdf_to_md(input_path: &Path, output_path: &Path) -> Result<()> {
        log::info!("Converting PDF to Markdown: {:?} -> {:?}", input_path, output_path);

        let file = fs::File::open(input_path)?;
        
        // Use lopdf to read PDF
        let doc = lopdf::Document::load(input_path)
            .map_err(|e| ConverterError::PdfError(format!("Failed to load PDF: {}", e)))?;

        let mut markdown_content = String::new();

        // Extract text from each page
        for page_id in doc.get_pages() {
            let page_num = page_id.0;
            markdown_content.push_str(&format!("## Page {}\n\n", page_num));

            // Try to extract text content
            if let Ok(page_obj) = doc.get_object(page_id) {
                // This is a simplified text extraction
                // Real PDF text extraction is complex due to encoding and layout
                let text = extract_text_from_page(&doc, page_id);
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

        // Convert markdown to PDF using lopdf
        // This is a simplified implementation
        // A production implementation might use a different approach
        // such as converting MD -> HTML -> PDF
        
        let pdf_content = create_pdf_from_markdown(&markdown_content)?;

        // Write PDF file
        fs::write(output_path, pdf_content)?;

        log::info!("Successfully converted Markdown to PDF");
        Ok(())
    }

    /// Extract text from a PDF page (simplified)
    fn extract_text_from_page(doc: &lopdf::Document, page_id: (u32, u32)) -> String {
        let mut text = String::new();

        // Get page dictionary
        if let Ok(page_dict) = doc.get_dictionary(page_id) {
            // Try to get Contents stream
            if let Ok(contents) = page_dict.get(b"Contents") {
                if let Ok(stream) = doc.get_dictionary(contents.as_reference().unwrap()) {
                    if let Ok(content_data) = stream.get_stream() {
                        // Decode and parse content stream
                        text = String::from_utf8_lossy(content_data).to_string();
                        
                        // Simple cleanup - remove PDF operators
                        text = text
                            .lines()
                            .filter(|line| !line.trim().is_empty())
                            .collect::<Vec<_>>()
                            .join("\n");
                    }
                }
            }
        }

        // If no text extracted, provide a message
        if text.is_empty() {
            text = "[Text extraction not available for this page]".to_string();
        }

        text
    }

    /// Create a simple PDF from markdown content
    fn create_pdf_from_markdown(markdown_content: &str) -> Result<Vec<u8>> {
        use lopdf::{Document, Object, Stream, Dictionary};
        use lopdf::content::Content;

        // Create a new PDF document
        let mut doc = Document::with_version("1.7");

        // Create pages from markdown content
        let mut pages = Vec::new();
        let mut current_y = 700.0;
        let mut page_contents = Vec::new();

        for line in markdown_content.lines() {
            // Simple text placement
            if line.starts_with("# ") {
                // Heading - new page
                if !page_contents.is_empty() {
                    pages.push(create_page(&doc, page_contents));
                    page_contents = Vec::new();
                    current_y = 700.0;
                }
                let text = line.trim_start_matches("# ");
                page_contents.push(format!("BT /F1 24 Tf 50 {} Td ({}) Tj ET", current_y, escape_pdf_text(text)));
                current_y -= 40.0;
            } else if !line.is_empty() {
                // Regular text
                page_contents.push(format!("BT /F1 12 Tf 50 {} Td ({}) Tj ET", current_y, escape_pdf_text(line)));
                current_y -= 20.0;

                // New page if needed
                if current_y < 50.0 {
                    pages.push(create_page(&doc, page_contents));
                    page_contents = Vec::new();
                    current_y = 700.0;
                }
            }
        }

        // Add last page
        if !page_contents.is_empty() {
            pages.push(create_page(&doc, page_contents));
        }

        // Create catalog and pages tree
        let catalog_id = doc.new_object_id();
        let pages_id = doc.new_object_id();

        // Add font
        let font_id = doc.new_object_id();
        doc.objects.insert(font_id, Object::Dictionary(Dictionary::from_iter(vec![
            ("Type", Object::Name("Font".into())),
            ("Subtype", Object::Name("Type1".into())),
            ("BaseFont", Object::Name("Helvetica".into())),
        ])));

        // Create pages tree
        let page_refs: Vec<Object> = pages.iter().map(|&id| Object::Reference(id)).collect();
        doc.objects.insert(pages_id, Object::Dictionary(Dictionary::from_iter(vec![
            ("Type", Object::Name("Pages".into())),
            ("Kids", Object::Array(page_refs)),
            ("Count", Object::Integer(pages.len() as i64)),
        ])));

        // Create catalog
        doc.objects.insert(catalog_id, Object::Dictionary(Dictionary::from_iter(vec![
            ("Type", Object::Name("Catalog".into())),
            ("Pages", Object::Reference(pages_id)),
        ])));

        doc.trailer.set("Root", catalog_id);

        // Save to buffer
        let mut buffer = Vec::new();
        doc.save_to(&mut buffer)
            .map_err(|e| ConverterError::PdfError(format!("Failed to save PDF: {}", e)))?;

        Ok(buffer)
    }

    /// Create a single PDF page
    fn create_page(doc: &Document, contents: Vec<String>) -> lopdf::ObjectId {
        use lopdf::Object;
        
        let page_id = doc.new_object_id();
        let contents_id = doc.new_object_id();

        let content_stream = contents.join("\n");
        let content_obj = Stream::new(
            Dictionary::new(),
            content_stream.into_bytes(),
        );

        doc.objects.insert(contents_id, Object::Stream(content_obj));

        let media_box = Object::Array(vec![
            Object::Integer(0),
            Object::Integer(0),
            Object::Integer(595),
            Object::Integer(842),
        ]);

        let page_dict = Dictionary::from_iter(vec![
            ("Type", Object::Name("Page".into())),
            ("Parent", Object::Reference(doc.trailer.get("Pages").unwrap().as_reference().unwrap())),
            ("MediaBox", media_box),
            ("Contents", Object::Reference(contents_id)),
            ("Resources", Object::Dictionary(Dictionary::from_iter(vec![
                ("Font", Object::Dictionary(Dictionary::from_iter(vec![
                    ("F1", Object::Reference(doc.get_objects().keys().find(|&&id| {
                        if let Some(obj) = doc.get_object(id) {
                            if let Some(dict) = obj.as_dictionary() {
                                return dict.get(b"Type").and_then(|t| t.as_name()).ok() == Some(b"Font");
                            }
                        }
                        false
                    }).unwrap_or(lopdf::ObjectId(0, 0)))).into())),
                ]))),
            ]))),
        ]);

        doc.objects.insert(page_id, Object::Dictionary(page_dict));
        page_id
    }

    /// Escape special characters for PDF text
    fn escape_pdf_text(text: &str) -> String {
        text.replace('\\', "\\\\")
            .replace('(', "\\(")
            .replace(')', "\\)")
            .replace('\r', "\\r")
    }
}
