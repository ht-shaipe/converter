//! Markdown 转 PDF 转换器 - 基于 markdown2pdf 库
//!
//! 支持丰富的 Markdown 特性：
//! - 标题、段落、强调、链接
//! - 代码块（支持语法高亮标记）
//! - 表格（支持对齐）
//! - 有序/无序列表（支持嵌套）
//! - 引用块、水平线
//!
//! 配置方式：
//! - 默认样式：直接使用
//! - 文件配置：指定 TOML 配置文件路径
//! - 内嵌配置：直接传入 TOML 字符串

use crate::error::{ConverterError, Result};
use crate::markdown2pdf as mdp; // 使用内联的 markdown2pdf 源码
use regex::Regex;
use std::path::Path;

/// 预处理 Markdown 文本，修复可能导致解析错误的特殊字符
fn preprocess_markdown(markdown: &str) -> String {
    let mut result = markdown.to_string();

    // 移除 base64 内嵌图片（会导致 PDF 页面溢出，无法换行）
    result = Regex::new(r"!\[([^\]]*)\]\(data:image/[^)]+\)")
        .unwrap()
        .replace_all(&result, |caps: &regex::Captures| {
            format!("[图片: {}]", &caps[1])
        })
        .to_string();

    // 替换中文/法文引号为普通引号
    result = result
        .replace('\u{201C}', "\"")
        .replace('\u{201D}', "\"")
        .replace('\u{2018}', "'")
        .replace('\u{2019}', "'")
        .replace('\u{00AB}', "\"")
        .replace('\u{00BB}', "\"");

    // 替换非断行空格和其他特殊空白字符为普通空格
    result = result
        .replace('\u{00A0}', " ")
        .replace('\u{3000}', " ")
        .replace('\u{200B}', "");

    // 移除 BOM (Byte Order Mark) - U+FEFF
    if result.starts_with('\u{FEFF}') {
        result = result[3..].to_string();
    }

    // 激进地清理损坏的强调语法格式
    // 这个文档有类似 "注 ** ** 文 ** ** 档" 的格式，需要彻底清理

    // 首先，移除所有孤立的星号（不是配对的 ** 或 *）
    // 移除单独一行的任何星号组合
    result = Regex::new(r"(?m)^\s*\*+\s*$")
        .unwrap()
        .replace_all(&result, "")
        .to_string();

    // 移除行首的孤立星号
    result = Regex::new(r"^\s*\*+")
        .unwrap()
        .replace_all(&result, "")
        .to_string();

    // 移除行尾的孤立星号
    result = Regex::new(r"\*+\s*$")
        .unwrap()
        .replace_all(&result, "")
        .to_string();

    // 移除损坏格式中的星号（字符之间插入的星号，如 "注 ** ** 文"）
    // 移除 " ** " 模式
    result = result.replace(" ** ", "");
    result = result.replace(" **", "");
    result = result.replace("** ", "");

    // 移除孤立的星号（星号前后有空格的情况）
    result = result.replace(" * ", "");
    result = result.replace(" *", "");
    result = result.replace("* ", "");

    // 多次应用以处理嵌套情况
    loop {
        let new_result = result.replace(" ** ", "");
        if new_result == result {
            break;
        }
        result = new_result;
    }

    // 处理转义星号 \* （替换为全角星号，避免被当作强调语法）
    result = result.replace("\\*", "\u{FF0A}");

    // 最后清理剩余的星号
    result = result.replace("**", "");

    // 修复单独一行的下划线问题（如 uint32_t）
    result = Regex::new(r"_[a-zA-Z0-9]")
        .unwrap()
        .replace_all(&result, "\u{FF3F}")
        .to_string();
    result = Regex::new(r"__{2,}")
        .unwrap()
        .replace_all(&result, |caps: &regex::Captures| {
            "\u{FF3F}".repeat(caps[0].len())
        })
        .to_string();
    result = result.replace("\\_", "\u{FF3F}");
    result = result.replace("\\\\", "\\");

    result
}

/// markdown2pdf 配置源
#[derive(Debug, Clone)]
pub enum MarkdownConfigSource {
    /// 使用内置默认样式
    Default,
    /// 从 TOML 文件加载样式配置
    File(String),
    /// 使用内嵌的 TOML 配置字符串
    Embedded(String),
}

impl Default for MarkdownConfigSource {
    fn default() -> Self {
        Self::Default
    }
}

/// Markdown 转 PDF 转换选项
#[derive(Debug, Clone)]
pub struct MarkdownToPdfOptions {
    /// 配置源
    pub config: MarkdownConfigSource,
    /// 默认字体名称（如 "Helvetica", "Times", "Arial"）
    pub default_font: Option<String>,
    /// 代码字体名称（如 "Courier"）
    pub code_font: Option<String>,
}

impl Default for MarkdownToPdfOptions {
    fn default() -> Self {
        Self {
            config: MarkdownConfigSource::Default,
            default_font: None,
            code_font: None,
        }
    }
}

/// 将 Markdown 字符串转换为 PDF 字节数据
pub fn markdown_to_pdf_bytes(
    markdown: &str,
    options: &MarkdownToPdfOptions,
) -> Result<Vec<u8>> {
    // 预处理 Markdown，处理可能导致解析错误的特殊字符
    let processed = preprocess_markdown(markdown);

    // 调试：保存预处理后的内容
    let _ = std::fs::write("/tmp/preprocessed.md", &processed);

    eprintln!("DEBUG: First 100 chars of processed:");
    for (i, c) in processed.chars().take(100).enumerate() {
        if c == '*' || c == '_' {
            eprintln!("  {}: '*' (star)", i);
        } else {
            eprintln!("  {}: {:?}", i, c);
        }
    }

    let config = match &options.config {
        MarkdownConfigSource::Default => mdp::config::ConfigSource::Default,
        MarkdownConfigSource::File(path) => {
            mdp::config::ConfigSource::File(path.as_str())
        }
        MarkdownConfigSource::Embedded(toml) => {
            mdp::config::ConfigSource::Embedded(toml.as_str())
        }
    };

    // 构建字体配置
    let font_config = if options.default_font.is_some() || options.code_font.is_some() {
        let mut fc = mdp::fonts::FontConfig::new();
        if let Some(ref font) = options.default_font {
            fc = fc.with_default_font(font.as_str());
        }
        if let Some(ref font) = options.code_font {
            fc = fc.with_code_font(font.as_str());
        }
        Some(fc)
    } else {
        None
    };

    mdp::parse_into_bytes(processed, config, font_config.as_ref())
        .map_err(|e| ConverterError::PdfError(e.to_string()))
}

/// 将 Markdown 文件转换为 PDF
pub fn markdown_file_to_pdf(
    input_path: impl AsRef<Path>,
    output_path: impl AsRef<Path>,
    options: &MarkdownToPdfOptions,
) -> Result<()> {
    let markdown = std::fs::read_to_string(input_path.as_ref())
        .map_err(|e| ConverterError::IoError(e))?;

    // 预处理 Markdown，处理可能导致解析错误的特殊字符
    let processed = preprocess_markdown(&markdown);

    let config = match &options.config {
        MarkdownConfigSource::Default => mdp::config::ConfigSource::Default,
        MarkdownConfigSource::File(path) => {
            mdp::config::ConfigSource::File(path.as_str())
        }
        MarkdownConfigSource::Embedded(toml) => {
            mdp::config::ConfigSource::Embedded(toml.as_str())
        }
    };

    // 构建字体配置
    let font_config = if options.default_font.is_some() || options.code_font.is_some() {
        let mut fc = mdp::fonts::FontConfig::new();
        if let Some(ref font) = options.default_font {
            fc = fc.with_default_font(font.as_str());
        }
        if let Some(ref font) = options.code_font {
            fc = fc.with_code_font(font.as_str());
        }
        Some(fc)
    } else {
        None
    };

    mdp::parse_into_file(
        processed,
        output_path.as_ref().to_str().unwrap_or("output.pdf"),
        config,
        font_config.as_ref(),
    )
    .map_err(|e| ConverterError::PdfError(e.to_string()))
}

/// 从字符串直接转换（便捷函数）
pub fn convert_markdown_to_pdf(markdown: &str, output_path: &str) -> Result<()> {
    let options = MarkdownToPdfOptions::default();
    let bytes = markdown_to_pdf_bytes(markdown, &options)?;
    std::fs::write(output_path, bytes)
        .map_err(|e| ConverterError::IoError(e))
}

/// 常用样式配置模板
pub mod presets {
    /// 学术论文样式
    pub fn academic_paper_config() -> String {
        r#"[page]
margins = { top = 72, right = 72, bottom = 72, left = 72 }
size = "a4"

[heading.1]
size = 18
bold = true
beforespacing = 12
afterspacing = 6
textcolor = { r = 0, g = 0, b = 0.4 }

[heading.2]
size = 16
bold = true
beforespacing = 10
afterspacing = 4
textcolor = { r = 0, g = 0, b = 0.3 }

[heading.3]
size = 14
bold = true
beforespacing = 8
afterspacing = 3

[text]
size = 11
beforespacing = 0
afterspacing = 6
"#
        .to_string()
    }

    /// 现代简约样式
    pub fn modern_minimal_config() -> String {
        r#"[page]
margins = { top = 40, right = 40, bottom = 40, left = 40 }
size = "a4"

[heading.1]
size = 24
bold = true
beforespacing = 0
afterspacing = 8
"#
        .to_string()
    }

    /// 代码文档样式
    pub fn code_documentation_config() -> String {
        r#"[text]
size = 10
fontfamily = "Courier"

[code]
size = 9
fontfamily = "Courier"

[heading.1]
size = 16
bold = true
fontfamily = "Helvetica"
"#
        .to_string()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::tempdir;

    #[test]
    fn test_preprocess_underscore() {
        let input = "uint32_t";
        let result = preprocess_markdown(input);
        let has_fullwidth = result.contains('\u{FF3F}');
        assert!(has_fullwidth, "Expected escaped or fullwidth underscore in: {:?}", result);
    }

    #[test]
    fn test_preprocess_unicode_quotes() {
        let input = "他说\u{201C}你好\u{201D}";
        let result = preprocess_markdown(input);
        assert!(result.contains("\"你好\""), "Expected ASCII quotes in: {:?}", result);
    }

    #[test]
    fn test_preprocess_bom() {
        let input = "\u{FEFF}# Hello";
        let result = preprocess_markdown(input);
        assert!(!result.starts_with('\u{FEFF}'), "BOM should be removed");
    }

    #[test]
    fn test_markdown_to_pdf_bytes() {
        let markdown = r#"# Test Document

This is a **test** with some *emphasis*.

## Features

- Item 1
- Item 2
- Item 3

## Code

```rust
fn main() {
    println!("Hello!");
}
```

| Header 1 | Header 2 |
|:---------|:--------:|
| Cell 1   | Cell 2   |
"#;

        let options = MarkdownToPdfOptions::default();
        let result = markdown_to_pdf_bytes(markdown, &options);

        assert!(result.is_ok());
        let bytes = result.unwrap();
        assert!(bytes.starts_with(b"%PDF-"));
    }

    #[test]
    fn test_markdown_file_to_pdf() {
        let temp_dir = tempdir().unwrap();
        let md_path = temp_dir.path().join("test.md");
        let pdf_path = temp_dir.path().join("output.pdf");

        std::fs::write(&md_path, "# Test\n\nHello **World**!").unwrap();

        let options = MarkdownToPdfOptions::default();
        let result = markdown_file_to_pdf(&md_path, &pdf_path, &options);

        assert!(result.is_ok());
        assert!(pdf_path.exists());

        let content = std::fs::read(&pdf_path).unwrap();
        assert!(content.starts_with(b"%PDF-"));
    }

    #[test]
    fn test_with_custom_font() {
        let markdown = "# Custom Font Test\n\nUsing Helvetica font.";
        let options = MarkdownToPdfOptions {
            config: MarkdownConfigSource::Default,
            default_font: Some("Helvetica".to_string()),
            code_font: Some("Courier".to_string()),
        };

        let result = markdown_to_pdf_bytes(markdown, &options);
        assert!(result.is_ok());
    }

    #[test]
    fn test_with_embedded_config() {
        let markdown = "# Styled Document\n\nWith custom styling.";
        let options = MarkdownToPdfOptions {
            config: MarkdownConfigSource::Embedded(presets::modern_minimal_config()),
            default_font: None,
            code_font: None,
        };

        let result = markdown_to_pdf_bytes(markdown, &options);
        assert!(result.is_ok());
    }
}
