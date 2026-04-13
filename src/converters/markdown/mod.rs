//! Markdown 转换模块
//!
//! 提供 Markdown 与其他格式之间的转换功能
//! 使用内联的 markdown2pdf 源码以便于直接修改和调试

pub mod markdown2pdf_converter;

pub use markdown2pdf_converter::{
    convert_markdown_to_pdf, markdown_file_to_pdf, markdown_to_pdf_bytes,
    presets, MarkdownConfigSource, MarkdownToPdfOptions,
};
