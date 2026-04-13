//! File format definitions

use std::path::Path;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FileFormat {
    Word,
    Excel,
    Pdf,
    Markdown,
    Unknown,
}

impl FileFormat {
    pub fn from_extension(path: &Path) -> Self {
        match path.extension().and_then(|e| e.to_str()).map(|e| e.to_ascii_lowercase()) {
            Some(ext) if ext == "docx" => FileFormat::Word,
            Some(ext) if ext == "xlsx" || ext == "xls" || ext == "xlsm" => FileFormat::Excel,
            Some(ext) if ext == "pdf" => FileFormat::Pdf,
            Some(ext) if ext == "md" || ext == "markdown" => FileFormat::Markdown,
            _ => FileFormat::Unknown,
        }
    }

    pub fn extension(&self) -> &'static str {
        match self {
            FileFormat::Word => "docx",
            FileFormat::Excel => "xlsx",
            FileFormat::Pdf => "pdf",
            FileFormat::Markdown => "md",
            FileFormat::Unknown => "",
        }
    }

    pub fn can_convert_to(&self, to: &FileFormat) -> bool {
        matches!(
            (self, to),
            (FileFormat::Word, FileFormat::Markdown)
                | (FileFormat::Markdown, FileFormat::Word)
                | (FileFormat::Excel, FileFormat::Markdown)
                | (FileFormat::Markdown, FileFormat::Excel)
                | (FileFormat::Pdf, FileFormat::Markdown)
                | (FileFormat::Markdown, FileFormat::Pdf)
        )
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ConversionType {
    WordToMarkdown,
    MarkdownToWord,
    ExcelToMarkdown,
    MarkdownToExcel,
    PdfToMarkdown,
    MarkdownToPdf,
    ImageToIco,
}

impl ConversionType {
    pub fn description(&self) -> &'static str {
        match self {
            ConversionType::WordToMarkdown => "Word to Markdown",
            ConversionType::MarkdownToWord => "Markdown to Word",
            ConversionType::ExcelToMarkdown => "Excel to Markdown",
            ConversionType::MarkdownToExcel => "Markdown to Excel",
            ConversionType::PdfToMarkdown => "PDF to Markdown",
            ConversionType::MarkdownToPdf => "Markdown to PDF",
            ConversionType::ImageToIco => "Image to ICO",
        }
    }

    /// MCP / CLI 使用的短标识
    pub fn kind_id(&self) -> &'static str {
        match self {
            ConversionType::WordToMarkdown => "docx_md",
            ConversionType::MarkdownToWord => "md_docx",
            ConversionType::ExcelToMarkdown => "xlsx_md",
            ConversionType::MarkdownToExcel => "md_xlsx",
            ConversionType::PdfToMarkdown => "pdf_md",
            ConversionType::MarkdownToPdf => "md_pdf",
            ConversionType::ImageToIco => "img_ico",
        }
    }

    pub fn from_kind_id(s: &str) -> Option<Self> {
        match s {
            "docx_md" | "word_md" => Some(ConversionType::WordToMarkdown),
            "md_docx" | "md_word" => Some(ConversionType::MarkdownToWord),
            "xlsx_md" | "excel_md" => Some(ConversionType::ExcelToMarkdown),
            "md_xlsx" | "md_excel" => Some(ConversionType::MarkdownToExcel),
            "pdf_md" => Some(ConversionType::PdfToMarkdown),
            "md_pdf" => Some(ConversionType::MarkdownToPdf),
            "img_ico" | "image_ico" => Some(ConversionType::ImageToIco),
            _ => None,
        }
    }
}
