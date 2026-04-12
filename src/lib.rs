//! File Converter Library
//! 
//! This library provides functionality to convert between various file formats:
//! - Word documents (DOCX) ↔ Markdown
//! - Excel spreadsheets (XLSX) ↔ Markdown  
//! - PDF ↔ Markdown
//!
//! # Example
//!
//! ```rust,no_run
//! use file_converter::{docx_to_md, xlsx_to_md, pdf_to_md};
//! use std::path::Path;
//!
//! // Convert Word to Markdown
//! docx_to_md(Path::new("input.docx"), Path::new("output.md")).unwrap();
//!
//! // Convert Excel to Markdown
//! xlsx_to_md(Path::new("data.xlsx"), Path::new("tables.md")).unwrap();
//!
//! // Convert PDF to Markdown
//! pdf_to_md(Path::new("report.pdf"), Path::new("notes.md")).unwrap();
//! ```
//!
//! # Architecture
//!
//! Each file format converter is implemented in a separate module:
//! - [`converters::word`] - Word document (DOCX) conversion
//! - [`converters::excel`] - Excel spreadsheet (XLSX) conversion
//! - [`converters::pdf`] - PDF document conversion
//!
//! All converters support bidirectional conversion with Markdown as the intermediate format.

pub mod error;
pub mod formats;
pub mod converters;

// Re-export error types
pub use error::{ConverterError, Result};

// Re-export format types
pub use formats::{FileFormat, ConversionType};

// Re-export converter modules
pub use converters::{word, excel, pdf};

// Re-export convenience functions
pub use converters::{
    docx_to_md, md_to_docx,
    xlsx_to_md, md_to_xlsx,
    pdf_to_md, md_to_pdf,
};
