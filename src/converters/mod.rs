//! Converters module
//! 
//! This module contains individual converters for different file formats:
//! - Word (DOCX) ↔ Markdown
//! - Excel (XLSX) ↔ Markdown
//! - PDF ↔ Markdown

pub mod word;
pub mod excel;
pub mod pdf;

// Re-export commonly used items
pub use word::{docx_to_md, md_to_docx};
pub use excel::{xlsx_to_md, md_to_xlsx};
pub use pdf::{pdf_to_md, md_to_pdf};
