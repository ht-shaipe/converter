//! Converters module
//! 
//! This module contains individual converters for different file formats:
//! - Word (DOCX) ↔ Markdown
//! - Excel (XLSX) ↔ Markdown
//! - PDF ↔ Markdown
//! - Image → ICO

pub mod word;
pub mod excel;
pub mod pdf;
pub mod ico;

// Re-export commonly used items
pub use word::{docx_to_md, md_to_docx};
pub use excel::{xlsx_to_md, md_to_xlsx};
pub use pdf::{pdf_to_md, md_to_pdf};
pub use ico::{convert_image_to_ico, convert_image_to_ico_with_config, convert_image_to_ico_size, convert_image_to_ico_multi, IcoConfig};
