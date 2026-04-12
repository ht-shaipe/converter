//! File Converter Library
//! 
//! This library provides functionality to convert between various file formats:
//! - Word documents (DOCX) ↔ Markdown
//! - Excel spreadsheets (XLSX) ↔ Markdown  
//! - PDF ↔ Markdown

pub mod error;
pub mod converter;
pub mod formats;

pub use error::{ConverterError, Result};
pub use converter::FileConverter;
pub use formats::{FileFormat, ConversionType};
