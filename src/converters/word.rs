//! Word document converter (DOCX <-> Markdown)
//! 
//! This module provides robust conversion between DOCX and Markdown formats.
//! 
//! # Features
//! - Extract text with formatting (bold, italic, strike-through)
//! - Convert headings (H1-H6)
//! - Handle lists (bulleted and numbered)
//! - Convert tables to Markdown format
//! - Embed images as Base64 data URIs or save to disk
//! - Detect code blocks from styled paragraphs
//! - Graceful fallback for non-conformant DOCX files
//! 
//! # Image embedding
//! When `embed_images` is enabled (default: **true**), images found inside
//! the `.docx` package are extracted and embedded as Base64 data URIs:
//! ```markdown
//! ![image](data:image/png;base64,iVBORw0...)
//! ```

use crate::error::{ConverterError, Result};
use base64::{engine::general_purpose::STANDARD as B64, Engine as _};
use regex::Regex;
use std::collections::HashMap;
use std::fs;
use std::io::Read;
use std::path::{Path, PathBuf};

// ─────────────────────────────────────────────────────────────────────────────
// Image map: rId → (data_uri_string or relative path)
// ─────────────────────────────────────────────────────────────────────────────

/// Maps relationship IDs (e.g. `"rId5"`) → base64 data URIs or file paths.
type ImageMap = HashMap<String, String>;

// ─────────────────────────────────────────────────────────────────────────────
// Converter configuration
// ─────────────────────────────────────────────────────────────────────────────

/// Word document converter with configurable options.
#[derive(Debug, Clone)]
pub struct WordConverter {
    /// When true (default), images are extracted and embedded as Base64 data URIs.
    pub embed_images: bool,
    /// When set, images are saved to this directory instead of embedding.
    pub output_images_dir: Option<PathBuf>,
}

impl Default for WordConverter {
    fn default() -> Self {
        Self { 
            embed_images: true, 
            output_images_dir: None 
        }
    }
}

impl WordConverter {
    /// Create a new converter with default settings.
    pub fn new() -> Self {
        Self::default()
    }

    /// Disable image embedding (images will be silently skipped).
    pub fn no_images(mut self) -> Self {
        self.embed_images = false;
        self
    }

    /// Save extracted images to the given directory instead of embedding as Base64.
    pub fn output_images_dir(mut self, dir: impl Into<PathBuf>) -> Self {
        self.output_images_dir = Some(dir.into());
        self
    }

    /// Convert a `.docx` file to a Markdown string.
    pub fn convert_to_md(&self, input_path: &Path) -> Result<String> {
        log::info!("Converting DOCX to Markdown: {:?}", input_path);
        
        // Validate input
        let ext = input_path.extension().and_then(|e| e.to_str()).unwrap_or("");
        if !ext.eq_ignore_ascii_case("docx") {
            return Err(ConverterError::WordError(
                format!("Unsupported extension: {} (only .docx is supported)", ext)
            ));
        }

        // Read the DOCX file
        let bytes = fs::read(input_path)
            .map_err(|e| ConverterError::WordError(format!("Failed to open DOCX: {}", e)))?;
        
        // Build image map
        let image_map = if self.embed_images || self.output_images_dir.is_some() {
            self.extract_image_map(&bytes).unwrap_or_default()
        } else {
            ImageMap::new()
        };

        // Try structured parsing first, then fallback to plain text extraction
        match self.parse_document(&bytes, &image_map) {
            Ok(md) => Ok(md),
            Err(_) => {
                // Fallback: extract plain text from XML
                let fallback = self.extract_plain_text(&bytes, &image_map)?;
                if fallback.trim().is_empty() {
                    Err(ConverterError::WordError("无法解析 docx 内容".to_string()))
                } else {
                    Ok(fallback)
                }
            }
        }
    }

    /// Convert a `.docx` file and write to an output Markdown file.
    pub fn convert_file(&self, input_path: &Path, output_path: &Path) -> Result<()> {
        let md = self.convert_to_md(input_path)?;
        fs::write(output_path, &md)
            .map_err(|e| ConverterError::IoError(std::io::Error::new(
                std::io::ErrorKind::Other,
                format!("Failed to write output: {}", e)
            )))?;
        log::info!("Successfully converted DOCX to Markdown: {:?}", output_path);
        Ok(())
    }

    /// Convert Markdown to DOCX
    pub fn convert_md_to_docx(&self, input_path: &Path, output_path: &Path) -> Result<()> {
        log::info!("Converting Markdown to DOCX: {:?} -> {:?}", input_path, output_path);

        // Read markdown content
        let markdown_content = fs::read_to_string(input_path)?;

        // Parse markdown and create DOCX structure
        let docx_content = parse_md_to_docx(&markdown_content)?;

        // Write DOCX file
        fs::write(output_path, &docx_content)?;

        log::info!("Successfully converted Markdown to DOCX");
        Ok(())
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// Public API functions (for backward compatibility)
// ─────────────────────────────────────────────────────────────────────────────

/// Convert DOCX to Markdown with default settings
pub fn docx_to_md(input_path: &Path, output_path: &Path) -> Result<()> {
    let converter = WordConverter::new();
    converter.convert_file(input_path, output_path)
}

/// Convert Markdown to DOCX
pub fn md_to_docx(input_path: &Path, output_path: &Path) -> Result<()> {
    let converter = WordConverter::new();
    converter.convert_md_to_docx(input_path, output_path)
}

// ─────────────────────────────────────────────────────────────────────────────
// Internal implementation
// ─────────────────────────────────────────────────────────────────────────────

impl WordConverter {
    /// Extract all images from the docx ZIP.
    fn extract_image_map(&self, bytes: &[u8]) -> Result<ImageMap> {
        let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes))
            .map_err(|e| ConverterError::WordError(format!("zip error: {:?}", e)))?;

        let rels = self.parse_relationships(&mut zip)?;
        let mut map = ImageMap::new();

        for (rid, target) in &rels {
            let zip_path = match normalise_media_path(target) {
                Some(p) => p,
                None => continue,
            };

            let mut img_bytes = Vec::new();
            match zip.by_name(&zip_path) {
                Ok(mut f) => {
                    f.read_to_end(&mut img_bytes)
                        .map_err(|e| ConverterError::WordError(format!("read image: {}", e)))?;
                }
                Err(_) => continue,
            }

            if let Some(ref dir) = self.output_images_dir {
                // Save to disk and return relative path
                let ext = target.rsplit('.').next().unwrap_or("png");
                let filename = format!("{}_{}.{}", rid.replace(':', "_"), sanitize_filename(target), ext);
                let img_path = dir.join(&filename);
                
                // Create directory if it doesn't exist
                if let Some(parent) = img_path.parent() {
                    fs::create_dir_all(parent)
                        .map_err(|e| ConverterError::WordError(format!("create dir: {}", e)))?;
                }
                
                fs::write(&img_path, &img_bytes)
                    .map_err(|e| ConverterError::WordError(format!("write image {}: {}", filename, e)))?;
                
                let rel = format!("./{}", filename);
                map.insert(rid.clone(), rel);
            } else {
                // Base64 data URI
                let mime = mime_from_path(&zip_path);
                let b64 = B64.encode(&img_bytes);
                let data_uri = format!("data:{};base64,{}", mime, b64);
                map.insert(rid.clone(), data_uri);
            }
        }

        Ok(map)
    }

    /// Parse `word/_rels/document.xml.rels` and return `rId → Target` map.
    fn parse_relationships(&self, zip: &mut zip::ZipArchive<std::io::Cursor<&[u8]>>) -> Result<HashMap<String, String>> {
        let mut xml = String::new();
        match zip.by_name("word/_rels/document.xml.rels") {
            Ok(mut f) => {
                f.read_to_string(&mut xml)
                    .map_err(|e| ConverterError::WordError(format!("read rels: {}", e)))?;
            }
            Err(_) => return Ok(HashMap::new()),
        }

        // <Relationship Id="rId5" Type="...image..." Target="../media/image1.png"/>
        let re = Regex::new(r#"(?i)<Relationship[^>]+Id="([^"]+)"[^>]+Target="([^"]+)"[^>]*/>"#)
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;

        let mut map = HashMap::new();
        for cap in re.captures_iter(&xml) {
            let id = cap[1].to_string();
            let target = cap[2].to_string();
            map.insert(id, target);
        }
        Ok(map)
    }

    /// Main document parsing using XML pass
    fn parse_document(&self, bytes: &[u8], image_map: &ImageMap) -> Result<String> {
        // Extract drawing map for inline images
        let drawing_order = self.extract_drawing_map(bytes).unwrap_or_default();
        self.document_to_md_xml_pass(bytes, image_map, &drawing_order)
    }

    /// Extract a map of paragraph index → image markdown from raw XML <w:drawing> elements.
    fn extract_drawing_map(&self, bytes: &[u8]) -> Result<Vec<String>> {
        let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes))
            .map_err(|e| ConverterError::WordError(format!("zip error: {:?}", e)))?;

        let mut xml = String::new();
        {
            let mut f = zip.by_name("word/document.xml")
                .map_err(|e| ConverterError::WordError(format!("missing document.xml: {:?}", e)))?;
            f.read_to_string(&mut xml)
                .map_err(|e| ConverterError::WordError(format!("read document.xml: {}", e)))?;
        }

        // Extract all r:embed values inside <w:drawing> blocks (order preserved)
        let drawing_re = Regex::new(r"(?s)<w:drawing\b.*?</w:drawing>")
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;
        let embed_re = Regex::new(r#"r:embed="([^"]+)""#)
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;
        let descr_re = Regex::new(r#"(?:descr|name)="([^"]+)""#)
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;

        let mut result = Vec::new();
        for drawing in drawing_re.find_iter(&xml) {
            let drawing_xml = drawing.as_str();
            if let Some(cap) = embed_re.captures(drawing_xml) {
                let rid = cap[1].to_string();
                // Try to get alt text from descr or name attribute
                let alt = descr_re.captures(drawing_xml)
                    .and_then(|c| c.get(1))
                    .map(|m| m.as_str().to_string())
                    .unwrap_or_else(|| "image".to_string());
                result.push(format!("{}|{}", rid, alt));
            }
        }
        Ok(result)
    }

    /// Extract plain text from docx XML, preserving paragraph boundaries.
    fn extract_plain_text(&self, bytes: &[u8], image_map: &ImageMap) -> Result<String> {
        let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes))
            .map_err(|e| ConverterError::WordError(format!("zip error: {:?}", e)))?;

        let mut xml = String::new();
        {
            let mut docx_xml = zip
                .by_name("word/document.xml")
                .map_err(|e| ConverterError::WordError(format!("missing document.xml: {:?}", e)))?;
            docx_xml.read_to_string(&mut xml)
                .map_err(|e| ConverterError::WordError(format!("read xml: {}", e)))?;
        }

        // Paragraph boundary markers
        let para_re = Regex::new(r#"<w:p[ >]"#)
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;
        let end_para_re = Regex::new(r#"</w:p>"#)
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;

        // Extract text content
        let text_re = Regex::new(r"<w:t[^>]*>([^<]*)</w:t>")
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;

        // Extract heading styles
        let heading1_re = Regex::new(r#"<w:pStyle w:val="Heading1"/>|<w:pStyle w:val="Title"/>"#)
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;
        let heading2_re = Regex::new(r#"<w:pStyle w:val="Heading2"/>"#)
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;
        let heading3_re = Regex::new(r#"<w:pStyle w:val="Heading3"/>"#)
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;

        // Formatting
        let bold_re = Regex::new(r"<w:b[^/]*/>")
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;
        let italic_re = Regex::new(r"<w:i[^/]*/>")
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;
        let strike_re = Regex::new(r"<w:(?:d?)strike[^/]*/>")
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;
        let run_re = Regex::new(r"<w:r\b[^>]*>(.*?)</w:r>")
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;

        // Drawing / image reference
        let drawing_re = Regex::new(r"(?s)<w:drawing\b.*?</w:drawing>")
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;
        let embed_re = Regex::new(r#"r:embed="([^"]+)""#)
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;
        let descr_re = Regex::new(r#"(?:descr|name)="([^"]+)""#)
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;

        let mut out = String::new();
        let mut seen_para_start = false;
        let xml_lower = xml.to_lowercase();

        for para_start in para_re.find_iter(&xml_lower) {
            let pos = para_start.start();
            let para_slice = &xml[pos..];

            // Detect heading
            let (is_heading, heading_level) = if heading1_re.is_match(para_slice) {
                (true, 1)
            } else if heading2_re.is_match(para_slice) {
                (true, 2)
            } else if heading3_re.is_match(para_slice) {
                (true, 3)
            } else {
                (false, 0)
            };

            // Find end of this paragraph
            let para_end = end_para_re.find(para_slice)
                .map(|m| m.start())
                .unwrap_or(para_slice.len());
            let para_xml = &para_slice[..para_end];

            // Images inside this paragraph
            let mut img_parts: Vec<String> = Vec::new();
            if self.embed_images {
                for drawing in drawing_re.find_iter(para_xml) {
                    let dxml = drawing.as_str();
                    if let Some(cap) = embed_re.captures(dxml) {
                        let rid = &cap[1];
                        let alt = descr_re.captures(dxml)
                            .and_then(|c| c.get(1))
                            .map(|m| m.as_str())
                            .unwrap_or("image");
                        if let Some(data_uri) = image_map.get(rid) {
                            img_parts.push(format!("![{}]({})\n\n", alt, data_uri));
                        }
                    }
                }
            }

            // Extract runs with per-run formatting
            let mut para_parts: Vec<String> = Vec::new();
            for run_cap in run_re.captures_iter(para_xml) {
                let run_xml = run_cap.get(1).map(|m| m.as_str()).unwrap_or("");

                let is_bold = bold_re.is_match(run_xml);
                let is_italic = italic_re.is_match(run_xml);
                let is_strike = strike_re.is_match(run_xml);

                let run_texts: Vec<&str> = text_re.captures_iter(run_xml)
                    .filter_map(|c| c.get(1).map(|m| m.as_str().trim()))
                    .filter(|s| !s.is_empty())
                    .collect();

                for raw in run_texts {
                    let mut t = raw.to_string();
                    if is_strike { t = format!("~~{}~~", t); }
                    if is_bold   { t = format!("**{}**", t); }
                    if is_italic { t = format!("_{}_", t); }
                    para_parts.push(t);
                }
            }

            let joined = para_parts.join("").trim().to_string();
            let has_text = !joined.is_empty();
            let has_img = !img_parts.is_empty();

            if !has_text && !has_img {
                continue;
            }

            if seen_para_start {
                out.push('\n');
            }
            seen_para_start = true;

            if has_text {
                if is_heading {
                    out.push_str(&"#".repeat(heading_level));
                    out.push(' ');
                }
                out.push_str(&joined);
                out.push_str("\n\n");
            }

            for img in img_parts {
                out.push_str(&img);
                out.push_str("\n\n");
            }
        }

        // Table fallback
        let table_re = Regex::new(r"<w:tbl>(.*?)</w:tbl>")
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;
        let row_re = Regex::new(r"<w:tr[ >](.*?)</w:tr>")
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;
        let cell_re = Regex::new(r"<w:tc[ >](.*?)</w:tc>")
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;
        let cell_text_re = Regex::new(r"<w:t[^>]*>([^<]*)</w:t>")
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;

        for tcap in table_re.captures_iter(&xml_lower) {
            let table_xml = tcap.get(1).map(|m| m.as_str()).unwrap_or("");

            if out.contains(&table_preview(table_xml)) {
                continue;
            }

            // Pseudo-table detection: single-column code/directory-tree content
            if let Some((code_lang, skip_first)) = detect_code_language(table_xml) {
                out.push_str(&format!("\n```{}\n", code_lang));
                let mut first = true;
                for rcap in row_re.captures_iter(table_xml) {
                    let row_xml = rcap.get(1).map(|m| m.as_str()).unwrap_or("");
                    if first && skip_first {
                        first = false;
                        continue;
                    }
                    first = false;
                    for ccap in cell_re.captures_iter(row_xml) {
                        if let Some(cell_xml) = ccap.get(1).map(|m| m.as_str()) {
                            let texts: String = cell_text_re.captures_iter(cell_xml)
                                .filter_map(|tc| tc.get(1).map(|m| m.as_str().trim()))
                                .collect();
                            if !texts.trim().is_empty() {
                                out.push_str(texts.trim());
                                out.push('\n');
                            }
                        }
                    }
                }
                out.push_str("```\n\n");
                continue;
            }

            out.push('\n');
            let mut header_written = false;

            for rcap in row_re.captures_iter(table_xml) {
                let row_xml = rcap.get(1).map(|m| m.as_str()).unwrap_or("");
                let cells: Vec<String> = cell_re.captures_iter(row_xml)
                    .filter_map(|ccap| {
                        let cell_xml = ccap.get(1)?.as_str();
                        let texts: Vec<&str> = cell_text_re.captures_iter(cell_xml)
                            .filter_map(|tc| tc.get(1).map(|m| m.as_str().trim()))
                            .filter(|s| !s.is_empty())
                            .collect();
                        if texts.is_empty() { None } else { Some(texts.join(" ")) }
                    })
                    .collect();

                if cells.is_empty() { continue; }

                if !header_written {
                    let col_count = cells.len();
                    out.push('|');
                    out.push_str(&cells.iter().map(|c| format!(" {} |", c)).collect::<String>());
                    out.push('\n');
                    out.push('|');
                    out.push_str(&"| --- |".repeat(col_count));
                    out.push('\n');
                    header_written = true;
                } else {
                    out.push('|');
                    out.push_str(&cells.iter().map(|c| format!(" {} |", c.replace('|', "\\|"))).collect::<String>());
                    out.push('\n');
                }
            }
            out.push('\n');
        }

        Ok(out.trim().to_string())
    }

    /// Process the raw XML paragraph list directly.
    fn document_to_md_xml_pass(&self, bytes: &[u8], image_map: &ImageMap, _drawing_order: &[String]) -> Result<String> {
        let mut zip = zip::ZipArchive::new(std::io::Cursor::new(bytes))
            .map_err(|e| ConverterError::WordError(format!("zip error: {:?}", e)))?;

        let mut xml = String::new();
        {
            let mut docx_xml = zip.by_name("word/document.xml")
                .map_err(|e| ConverterError::WordError(format!("missing document.xml: {:?}", e)))?;
            docx_xml.read_to_string(&mut xml)
                .map_err(|e| ConverterError::WordError(format!("read xml: {}", e)))?;
        }

        let para_re = Regex::new(r"(?s)<w:p[ >].*?</w:p>")
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;
        let text_re  = Regex::new(r"<w:t[^>]*>([^<]*)</w:t>")
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;

        let heading1_re = Regex::new(r#"<w:pStyle[^>]+w:val="(Heading1|Title)"/>"#)
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;
        let heading2_re = Regex::new(r#"<w:pStyle[^>]+w:val="Heading2"/>"#)
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;
        let heading3_re = Regex::new(r#"<w:pStyle[^>]+w:val="Heading3"/>"#)
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;

        let bold_re   = Regex::new(r"<w:b[^/]*/>")
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;
        let italic_re = Regex::new(r"<w:i[^/]*/>")
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;
        let strike_re = Regex::new(r"<w:(?:d?)strike[^/]*/>")
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;

        let drawing_re = Regex::new(r"(?s)<w:drawing\b.*?</w:drawing>")
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;
        let embed_re   = Regex::new(r#"r:embed="([^"]+)""#)
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;
        let descr_re   = Regex::new(r#"(?:descr|name)="([^"]+)""#)
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;

        let mut out = String::new();

        let table_re  = Regex::new(r"(?s)<w:tbl>.*?</w:tbl>")
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;
        let row_re    = Regex::new(r"(?s)<w:tr[ >].*?</w:tr>")
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;
        let cell_re   = Regex::new(r"(?s)<w:tc[ >].*?</w:tc>")
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;
        let cell_txt  = Regex::new(r"<w:t[^>]*>([^<]*)</w:t>")
            .map_err(|e| ConverterError::WordError(format!("regex error: {}", e)))?;

        #[derive(Debug, Clone, Copy, PartialEq)]
        enum ElementKind { Para, Table }

        #[derive(Debug)]
        struct Element {
            kind: ElementKind,
            start: usize,
            content: String,
        }

        let mut elements: Vec<Element> = Vec::new();

        // Collect all table ranges so we can filter out paragraphs inside them
        let mut table_ranges: Vec<(usize, usize)> = Vec::new();
        for tbl_match in table_re.find_iter(&xml) {
            table_ranges.push((tbl_match.start(), tbl_match.end()));
        }

        // Helper to check if a position is inside any table
        let is_inside_table = |pos: usize| -> bool {
            for (start, end) in &table_ranges {
                if *start <= pos && pos < *end {
                    return true;
                }
            }
            false
        };

        // Collect body paragraphs from original xml, filtering out those inside tables
        for para_match in para_re.find_iter(&xml) {
            if is_inside_table(para_match.start()) {
                continue;
            }
            elements.push(Element {
                kind: ElementKind::Para,
                start: para_match.start(),
                content: para_match.as_str().to_string(),
            });
        }

        // Collect tables
        for tbl_match in table_re.find_iter(&xml) {
            elements.push(Element {
                kind: ElementKind::Table,
                start: tbl_match.start(),
                content: tbl_match.as_str().to_string(),
            });
        }

        // Sort by original document position
        elements.sort_by_key(|e| e.start);

        // Process elements in order
        for elem in elements {
            match elem.kind {
                ElementKind::Para => {
                    let para_xml = &elem.content;
                    let text_parts: Vec<String> = text_re.captures_iter(para_xml)
                        .filter_map(|c| c.get(1).map(|m| m.as_str().trim().to_string()))
                        .filter(|s| !s.is_empty())
                        .collect();

                    // Extract drawings in this paragraph
                    let para_drawings: Vec<&str> = drawing_re.find_iter(para_xml).map(|m| m.as_str()).collect();

                    // Skip empty paragraphs with no drawings
                    if text_parts.is_empty() && para_drawings.is_empty() {
                        continue;
                    }

                    // Apply inline formatting across all text runs
                    let raw_text = text_parts.join("");
                    if !raw_text.trim().is_empty() {
                        let formatted = format_text_with_styles(para_xml, &text_re, &bold_re, &italic_re, &strike_re);

                        // Check for code style
                        let is_code_style = para_xml.contains(r#"w:val="Code""#) ||
                                            para_xml.contains(r#"w:val="InlineCode""#) ||
                                            para_xml.contains(r#"w:val="code""#) ||
                                            para_xml.contains(r#"w:val="inlinecode""#);

                        let line = if !formatted.trim().is_empty() {
                            if heading1_re.is_match(para_xml) {
                                format!("# {}\n\n", formatted.trim())
                            } else if heading2_re.is_match(para_xml) {
                                format!("## {}\n\n", formatted.trim())
                            } else if heading3_re.is_match(para_xml) {
                                format!("### {}\n\n", formatted.trim())
                            } else if is_code_style {
                                format!("```\n{}\n```\n\n", formatted.trim())
                            } else {
                                let mut para_out = formatted.trim().to_string();
                                para_out.push_str("\n\n");
                                para_out
                            }
                        } else {
                            String::new()
                        };
                        out.push_str(&line);
                    }

                    // Emit images from drawings in this paragraph
                    for drawing_xml in para_drawings {
                        if let Some(cap) = embed_re.captures(drawing_xml) {
                            let rid = cap[1].to_string();
                            let alt = descr_re.captures(drawing_xml)
                                .and_then(|c| c.get(1))
                                .map(|m| m.as_str())
                                .unwrap_or("image");
                            if let Some(data_uri) = image_map.get(&rid) {
                                out.push_str(&format!("![{}]({})\n\n", alt, data_uri));
                            }
                        }
                    }
                }

                ElementKind::Table => {
                    let table_xml = &elem.content;

                    // Pseudo-table detection: single-column code/directory-tree content
                    if let Some((code_lang, skip_first)) = detect_code_language(table_xml) {
                        out.push_str(&format!("```{}\n", code_lang));
                        let mut first = true;
                        for rcap in row_re.find_iter(table_xml) {
                            let row_xml = rcap.as_str();
                            if first && skip_first {
                                first = false;
                                continue;
                            }
                            let is_first_row = first;
                            first = false;

                            for ccap in cell_re.find_iter(row_xml) {
                                let cell_xml = ccap.as_str();
                                let texts: Vec<&str> = cell_txt.captures_iter(cell_xml)
                                    .filter_map(|tc| tc.get(1).map(|m| m.as_str().trim()))
                                    .collect();
                                let combined = texts.join(" ");
                                if combined.trim().is_empty() { continue; }

                                let final_text = if is_first_row && !code_lang.is_empty() {
                                    let lower = combined.to_ascii_lowercase();
                                    let label = code_lang.to_lowercase();
                                    if lower.starts_with(&label) {
                                        let after = combined[label.len()..].trim_start();
                                        if after.is_empty() { continue; }
                                        after.to_string()
                                    } else {
                                        combined
                                    }
                                } else {
                                    combined
                                };

                                if !final_text.trim().is_empty() {
                                    out.push_str(final_text.trim());
                                    out.push('\n');
                                }
                            }
                        }
                        out.push_str("```\n\n");
                        continue;
                    }

                    // Normal table: render as Markdown table
                    let mut header_written = false;
                    for rcap in row_re.captures_iter(table_xml) {
                        let row_xml = rcap.get(0).map(|m| m.as_str()).unwrap_or("");
                        let cells: Vec<String> = cell_re.captures_iter(row_xml)
                            .filter_map(|ccap| {
                                let cell_xml = ccap.get(0).map(|m| m.as_str()).unwrap_or("");
                                let texts: Vec<&str> = cell_txt.captures_iter(cell_xml)
                                    .filter_map(|tc| tc.get(1).map(|m| m.as_str().trim()))
                                    .filter(|s| !s.is_empty())
                                    .collect();
                                if texts.is_empty() { None } else { Some(texts.join(" ")) }
                            })
                            .collect();

                        if cells.is_empty() { continue; }
                        let col_count = cells.len();

                        out.push('|');
                        out.push_str(&cells.iter().map(|c| format!(" {} |", c.replace('|', "\\|"))).collect::<String>());
                        out.push('\n');

                        if !header_written {
                            out.push('|');
                            for _ in 0..col_count { out.push_str(" --- |"); }
                            out.push('\n');
                            header_written = true;
                        }
                    }
                    out.push('\n');
                }
            }
        }

        Ok(out.trim().to_string())
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// Helper functions
// ─────────────────────────────────────────────────────────────────────────────

/// Apply bold/italic/strike formatting to raw text segments within a paragraph XML block.
fn format_text_with_styles(
    para_xml: &str,
    text_re: &Regex,
    bold_re: &Regex,
    italic_re: &Regex,
    strike_re: &Regex,
) -> String {
    let run_re = Regex::new(r"(?s)<w:r\b[^>]*>.*?</w:r>").unwrap();
    let mut parts: Vec<String> = Vec::new();

    for run_cap in run_re.captures_iter(para_xml) {
        let run_xml = run_cap.get(0).map(|m| m.as_str()).unwrap_or("");
        let is_bold   = bold_re.is_match(run_xml);
        let is_italic = italic_re.is_match(run_xml);
        let is_strike = strike_re.is_match(run_xml);

        for tc in text_re.captures_iter(run_xml) {
            let raw = tc.get(1).map(|m| m.as_str().trim()).unwrap_or("");
            if raw.is_empty() { continue; }
            let mut s = escape_md(raw);
            if is_strike { s = format!("~~{}~~", s); }
            if is_bold   { s = format!("**{}**", s); }
            if is_italic { s = format!("_{}_", s); }
            parts.push(s);
        }
    }

    parts.join("")
}

/// Detect if a table is a "pseudo-table" that should be rendered as a code block.
fn detect_code_language(table_xml: &str) -> Option<(String, bool)> {
    let cell_re = Regex::new(r"(?s)<w:tc[ >].*?</w:tc>").unwrap();
    let cell_txt = Regex::new(r"<w:t[^>]*>([^<]*)</w:t>").unwrap();

    let row_re = Regex::new(r"(?s)<w:tr[ >].*?</w:tr>").unwrap();
    let rows: Vec<String> = row_re.find_iter(table_xml)
        .map(|rcap| {
            let row_xml = rcap.as_str();
            let mut all_texts = Vec::new();
            for ccap in cell_re.find_iter(row_xml) {
                let cell_xml = ccap.as_str();
                let texts: Vec<&str> = cell_txt.captures_iter(cell_xml)
                    .filter_map(|tc| tc.get(1).map(|m| m.as_str().trim()))
                    .collect();
                let combined = texts.join(" ");
                if !combined.trim().is_empty() {
                    all_texts.push(combined.trim().to_string());
                }
            }
            all_texts.join(" ")
        })
        .collect();

    if rows.is_empty() {
        return None;
    }

    let lang_labels: Vec<(&str, &str)> = vec![
        ("plain text", "text"),
        ("code block", "text"),
        ("json", "json"),
        ("xml", "xml"),
        ("yaml", "yaml"),
        ("toml", "toml"),
        ("html", "html"),
        ("css", "css"),
        ("javascript", "javascript"),
        ("typescript", "typescript"),
        ("python", "python"),
        ("rust", "rust"),
        ("sql", "sql"),
        ("bash", "bash"),
        ("shell", "bash"),
        ("cmd", "cmd"),
        ("powershell", "powershell"),
        ("go", "go"),
        ("java", "java"),
        ("c++", "cpp"),
        ("c#", "csharp"),
        ("php", "php"),
        ("ruby", "ruby"),
        ("swift", "swift"),
        ("kotlin", "kotlin"),
        ("markdown", "markdown"),
        ("log", "log"),
        ("output", ""),
        ("result", ""),
        ("error", ""),
        ("warning", ""),
        ("info", ""),
    ];

    let is_lang_label = |line: &str| -> Option<String> {
        let lower = line.to_ascii_lowercase();
        for (label, lang) in &lang_labels {
            if lower == *label {
                return Some(lang.to_string());
            }
        }
        None
    };

    let is_code_like = |line: &str| -> bool {
        if line.is_empty() { return false; }
        let lower = line.to_ascii_lowercase();

        if line.starts_with('@') { return true; }
        if lower.contains(":/") || line.contains('\\') || line.starts_with('/') { return true; }
        if lower.contains("├──") || lower.contains("└──") || lower.contains("│") { return true; }
        if line.starts_with(' ') && line.contains(|c: char| c == ':' || c == '#' || c == '$') { return true; }
        
        let cli_commands = [
            "sudo ", "npm ", "cargo ", "git ", "docker ", "kubectl ", "helm ",
            "curl ", "wget ", "openssl ", "keytool ", "javac ", "gradle ",
            "mvn ", "make ", "cmake ", "pip ", "pip3 ", "python ", "python3 ",
            "node ", "ruby ", "perl ", "go ", "rustc ", "gcc ", "clang ",
            "ls ", "cd ", "pwd ", "mkdir ", "rm ", "cp ", "mv ", "cat ",
            "grep ", "sed ", "awk ", "find ", "xargs ", "chmod ", "chown ",
            "tar ", "zip ", "unzip ", "ssh ", "scp ", "rsync ", "ping ",
        ];
        for cmd in &cli_commands {
            if lower.starts_with(cmd) { return true; }
        }
        
        if line.contains(|c: char| c.is_whitespace()) {
            let parts: Vec<&str> = line.split_whitespace().collect();
            if parts.len() >= 2 {
                let first = parts[0].to_lowercase();
                let second = parts[1];
                if second.starts_with('-') { return true; }
                if is_lang_label(&first).is_some() { return true; }
            }
        }
        
        if is_lang_label(line).is_some() { return true; }
        false
    };

    if rows.len() >= 2 {
        let first_lang = is_lang_label(&rows[0]);
        if first_lang.is_some() {
            let rest_all_code = rows[1..].iter().all(|r| is_code_like(r) || r.is_empty());
            if rest_all_code {
                return Some((first_lang.unwrap_or_default(), true));
            }
        }
    }

    let all_code = rows.iter().all(|r| is_code_like(r) || r.is_empty());
    if all_code {
        let combined = rows.join(" ");
        let lower = combined.to_ascii_lowercase();
        for (label, lang) in &lang_labels {
            if lower.starts_with(label) {
                let after_label = lower[label.len()..].trim_start();
                let is_pure_label_row = rows.len() == 1 && after_label.is_empty();
                let skip_first = rows.len() > 1 || is_pure_label_row;
                return Some((lang.to_string(), skip_first));
            }
            if lower.contains(&format!(" {} ", label)) {
                return Some((lang.to_string(), false));
            }
        }
        return Some((String::new(), false));
    }

    None
}

fn table_preview(table_xml: &str) -> String {
    let cell_re = Regex::new(r"<w:tc[ >](.*?)</w:tc>").unwrap();
    let cell_text_re = Regex::new(r"<w:t[^>]*>([^<]*)</w:t>").unwrap();
    let row_re = Regex::new(r"<w:tr[ >](.*?)</w:tr>").unwrap();

    let first_row = row_re.captures_iter(table_xml).next()
        .and_then(|r| r.get(1))
        .map(|m| m.as_str())
        .unwrap_or("");

    cell_re.captures_iter(first_row)
        .filter_map(|c| {
            let txt = c.get(1)?.as_str();
            let first_text = cell_text_re.captures(txt)?.get(1)?.as_str().trim();
            Some(first_text)
        })
        .collect::<Vec<_>>()
        .join(" ")
}

/// Normalise a relationship Target to a ZIP entry path under `word/`.
fn normalise_media_path(target: &str) -> Option<String> {
    let lower = target.to_ascii_lowercase();

    if !lower.contains("media/") {
        return None;
    }

    let known_ext = ["png", "jpg", "jpeg", "gif", "bmp", "webp", "tiff", "tif", "emf", "wmf", "svg"];
    let has_ext = known_ext.iter().any(|ext| lower.ends_with(ext));
    if !has_ext {
        return None;
    }

    if target.starts_with("../") {
        Some(format!("word/{}", &target[3..]))
    } else if target.starts_with('/') {
        Some(target.trim_start_matches('/').to_string())
    } else {
        Some(format!("word/{}", target))
    }
}

/// Guess MIME type from file path extension.
fn mime_from_path(path: &str) -> &'static str {
    let lower = path.to_ascii_lowercase();
    if lower.ends_with(".png")  { return "image/png"; }
    if lower.ends_with(".jpg") || lower.ends_with(".jpeg") { return "image/jpeg"; }
    if lower.ends_with(".gif")  { return "image/gif"; }
    if lower.ends_with(".bmp")  { return "image/bmp"; }
    if lower.ends_with(".webp") { return "image/webp"; }
    if lower.ends_with(".tiff") || lower.ends_with(".tif") { return "image/tiff"; }
    if lower.ends_with(".svg")  { return "image/svg+xml"; }
    "image/x-emf"
}

/// Strip path components and unsafe characters to produce a safe filename.
fn sanitize_filename(target: &str) -> String {
    let base = std::path::Path::new(target)
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("image");

    base.chars()
        .map(|c| {
            if c.is_alphanumeric() || c == '-' || c == '_' {
                c
            } else {
                '_'
            }
        })
        .collect::<String>()
        .trim_matches('_')
        .to_string()
}

/// Escape Markdown special characters
fn escape_md(text: &str) -> String {
    text.replace('\\', "\\\\")
        .replace('*', "\\*")
        .replace('_', "\\_")
        .replace('[', "\\[")
        .replace(']', "\\]")
        .replace('(', "\\(")
        .replace(')', "\\)")
        .replace('#', "\\#")
        .replace('+', "\\+")
        .replace('!', "\\!")
}

/// Parse Markdown content to DOCX binary
fn parse_md_to_docx(markdown_content: &str) -> Result<Vec<u8>> {
    use docx_rs::*;

    let mut doc = Docx::new();
    
    for line in markdown_content.lines() {
        let trimmed = line.trim();
        
        if trimmed.starts_with("# ") {
            let text = trimmed.trim_start_matches("# ").trim();
            let run = Run::new().add_text(text);
            let paragraph = Paragraph::new().add_run(run);
            doc = doc.add_paragraph(paragraph);
        } else if trimmed.starts_with("## ") {
            let text = trimmed.trim_start_matches("## ").trim();
            let run = Run::new().add_text(text);
            let paragraph = Paragraph::new().add_run(run);
            doc = doc.add_paragraph(paragraph);
        } else if trimmed.starts_with("### ") {
            let text = trimmed.trim_start_matches("### ").trim();
            let run = Run::new().add_text(text);
            let paragraph = Paragraph::new().add_run(run);
            doc = doc.add_paragraph(paragraph);
        } else if trimmed.starts_with("#### ") {
            let text = trimmed.trim_start_matches("#### ").trim();
            let run = Run::new().add_text(text);
            let paragraph = Paragraph::new().add_run(run);
            doc = doc.add_paragraph(paragraph);
        } else if trimmed.starts_with("- ") || trimmed.starts_with("* ") {
            let text = trimmed.trim_start_matches("- ").trim_start_matches("* ").trim();
            let run = Run::new().add_text(text);
            let paragraph = Paragraph::new()
                .add_run(run)
                .set_style(ParagraphStyle::ListParagraph);
            doc = doc.add_paragraph(paragraph);
        } else if !trimmed.is_empty() {
            let run = Run::new().add_text(trimmed);
            let paragraph = Paragraph::new().add_run(run);
            doc = doc.add_paragraph(paragraph);
        }
    }

    let mut buffer = Vec::new();
    doc.build().write(&mut buffer)
        .map_err(|e| ConverterError::WordError(format!("Failed to build DOCX: {}", e)))?;

    Ok(buffer)
}

// ─────────────────────────────────────────────────────────────────────────────
// Tests
// ─────────────────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_escape_md() {
        assert_eq!(escape_md("hello world"), "hello world");
        assert_eq!(escape_md("**bold**"), "\\*\\*bold\\*\\*");
    }

    #[test]
    fn test_normalise_media_path() {
        assert_eq!(normalise_media_path("../media/image1.png"), Some("word/media/image1.png".to_string()));
        assert_eq!(normalise_media_path("media/image1.jpg"), Some("word/media/image1.jpg".to_string()));
        assert_eq!(normalise_media_path("../embeddings/sheet.xlsx"), None);
        assert_eq!(normalise_media_path("hyperlink_target"), None);
    }

    #[test]
    fn test_mime_from_path() {
        assert_eq!(mime_from_path("word/media/img.png"), "image/png");
        assert_eq!(mime_from_path("word/media/photo.JPEG"), "image/jpeg");
        assert_eq!(mime_from_path("word/media/anim.gif"), "image/gif");
    }

    #[test]
    fn test_converter_no_images_flag() {
        let c = WordConverter::new().no_images();
        assert!(!c.embed_images);
    }
}
