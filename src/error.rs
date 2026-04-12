//! Error types for file conversion operations

use thiserror::Error;

/// Result type alias for converter operations
pub type Result<T> = std::result::Result<T, ConverterError>;

/// Errors that can occur during file conversion
#[derive(Error, Debug)]
pub enum ConverterError {
    #[error("File not found: {0}")]
    FileNotFound(String),

    #[error("Unsupported file format: {0}")]
    UnsupportedFormat(String),

    #[error("Conversion not supported: {from} to {to}")]
    UnsupportedConversion { from: String, to: String },

    #[error("IO error: {0}")]
    IoError(#[from] std::io::Error),

    #[error("Word document error: {0}")]
    WordError(String),

    #[error("Excel document error: {0}")]
    ExcelError(String),

    #[error("PDF error: {0}")]
    PdfError(String),

    #[error("Markdown processing error: {0}")]
    MarkdownError(String),

    #[error("Invalid path: {0}")]
    InvalidPath(String),

    #[error("Conversion failed: {0}")]
    ConversionFailed(String),
}

impl From<ConverterError> for anyhow::Error {
    fn from(err: ConverterError) -> Self {
        anyhow::anyhow!(err.to_string())
    }
}
