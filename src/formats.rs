//! File format definitions and detection

use crate::error::{ConverterError, Result};
use std::path::Path;

/// Supported file formats
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum FileFormat {
    /// Microsoft Word document (DOCX)
    Word,
    /// Microsoft Excel spreadsheet (XLSX)
    Excel,
    /// Portable Document Format (PDF)
    Pdf,
    /// Markdown document
    Markdown,
    /// Unknown or unsupported format
    Unknown,
}

impl FileFormat {
    /// Detect file format from file extension
    pub fn from_extension(path: &Path) -> Self {
        match path.extension().and_then(|ext| ext.to_str()) {
            Some("docx") => FileFormat::Word,
            Some("xlsx") | Some("xls") => FileFormat::Excel,
            Some("pdf") => FileFormat::Pdf,
            Some("md") | Some("markdown") => FileFormat::Markdown,
            _ => FileFormat::Unknown,
        }
    }

    /// Detect file format from file path
    pub fn from_path(path: &Path) -> Result<Self> {
        if !path.exists() {
            return Err(ConverterError::FileNotFound(
                path.display().to_string(),
            ));
        }

        let format = Self::from_extension(path);
        if format == FileFormat::Unknown {
            Err(ConverterError::UnsupportedFormat(
                path.display().to_string(),
            ))
        } else {
            Ok(format)
        }
    }

    /// Get the primary file extension for this format
    pub fn extension(&self) -> &'static str {
        match self {
            FileFormat::Word => "docx",
            FileFormat::Excel => "xlsx",
            FileFormat::Pdf => "pdf",
            FileFormat::Markdown => "md",
            FileFormat::Unknown => "",
        }
    }

    /// Check if conversion to another format is supported
    pub fn can_convert_to(&self, target: &FileFormat) -> bool {
        matches!(
            (self, target),
            (FileFormat::Word, FileFormat::Markdown)
                | (FileFormat::Markdown, FileFormat::Word)
                | (FileFormat::Excel, FileFormat::Markdown)
                | (FileFormat::Markdown, FileFormat::Excel)
                | (FileFormat::Pdf, FileFormat::Markdown)
                | (FileFormat::Markdown, FileFormat::Pdf)
        )
    }
}

/// Conversion type specifying source and target formats
#[derive(Debug, Clone)]
pub struct ConversionType {
    pub source: FileFormat,
    pub target: FileFormat,
}

impl ConversionType {
    /// Create a new conversion type
    pub fn new(source: FileFormat, target: FileFormat) -> Self {
        Self { source, target }
    }

    /// Check if this conversion is supported
    pub fn is_supported(&self) -> bool {
        self.source.can_convert_to(&self.target)
    }
}

impl std::fmt::Display for FileFormat {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            FileFormat::Word => write!(f, "Word (DOCX)"),
            FileFormat::Excel => write!(f, "Excel (XLSX)"),
            FileFormat::Pdf => write!(f, "PDF"),
            FileFormat::Markdown => write!(f, "Markdown"),
            FileFormat::Unknown => write!(f, "Unknown"),
        }
    }
}

impl std::fmt::Display for ConversionType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{} -> {}", self.source, self.target)
    }
}
