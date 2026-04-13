mod docx_markdown;
pub mod excel;
pub mod ico;
pub mod markdown;
pub mod pdf;
pub mod word;

// Re-exports from markdown module
pub use markdown::{
    convert_markdown_to_pdf, markdown_file_to_pdf, markdown_to_pdf_bytes,
    presets, MarkdownConfigSource, MarkdownToPdfOptions,
};

pub use excel::{md_to_xlsx, xlsx_to_md};
pub use ico::{
    convert_image_to_ico, convert_image_to_ico_multi, convert_image_to_ico_size,
    convert_image_to_ico_with_config, IcoConfig,
};
pub use pdf::pdf_to_md;

use std::path::Path;

/// Markdown → PDF：使用新的 markdown2pdf 库
pub fn md_to_pdf(input_path: &Path, output_path: &Path) -> crate::Result<()> {
    let options = crate::converters::MarkdownToPdfOptions::default();
    crate::converters::markdown_file_to_pdf(input_path, output_path, &options)
}
pub use word::{docx_to_md, md_to_docx};
