//! File Converter Library

pub mod converters;
pub mod error;
pub mod formats;
pub mod markdown2pdf; // 内联的 markdown2pdf 源码

pub use converters::{
    convert_image_to_ico, convert_image_to_ico_multi, convert_image_to_ico_size,
    convert_image_to_ico_with_config, docx_to_md, md_to_docx, md_to_pdf, md_to_xlsx, pdf_to_md,
    xlsx_to_md, IcoConfig,
};
pub use error::{ConverterError, Result};
pub use formats::{ConversionType, FileFormat};

use std::path::Path;

/// 按 [`ConversionType`] 执行一次转换。
pub fn run_conversion(ct: ConversionType, input: &Path, output: &Path) -> Result<()> {
    match ct {
        ConversionType::WordToMarkdown => docx_to_md(input, output),
        ConversionType::MarkdownToWord => md_to_docx(input, output),
        ConversionType::ExcelToMarkdown => xlsx_to_md(input, output),
        ConversionType::MarkdownToExcel => md_to_xlsx(input, output),
        ConversionType::PdfToMarkdown => pdf_to_md(input, output),
        ConversionType::MarkdownToPdf => md_to_pdf(input, output),
        ConversionType::ImageToIco => convert_image_to_ico(input, output),
    }
}
