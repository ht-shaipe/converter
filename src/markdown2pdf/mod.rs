//! 内联 markdown2pdf 源码模块
//!
//! 从 https://github.com/theiskaa/markdown2pdf 整合而来
//! 直接集成以便自由修改和调试

pub mod config;
mod debug;
pub mod fonts;
pub mod markdown;
pub mod pdf;
pub mod styling;
pub mod validation;

use std::error::Error;
use std::fmt;
use markdown::Lexer;
use pdf::Pdf;

// Re-export public types
pub use config::{load_config_from_source, parse_config_string, ConfigSource};
pub use fonts::{load_font, FontConfig, FontSource};
pub use markdown::{Lexer as MdLexer, LexerError, Token};
pub use styling::{BasicTextStyle, Margins, StyleMatch, TextAlignment};
pub use validation::{ValidationWarning, WarningKind, validate_conversion};

/// Represents errors that can occur during the markdown-to-pdf conversion process.
/// This includes both parsing failures and PDF generation issues.
#[derive(Debug)]
pub enum MdpError {
    /// Indicates an error occurred while parsing the Markdown content
    ParseError {
        message: String,
        position: Option<usize>,
        suggestion: Option<String>,
    },
    /// Indicates an error occurred during PDF file generation
    PdfError {
        message: String,
        path: Option<String>,
        suggestion: Option<String>,
    },
    /// Indicates a font loading error
    FontError {
        font_name: String,
        message: String,
        suggestion: String,
    },
    /// Indicates an invalid configuration
    ConfigError { message: String, suggestion: String },
    /// Indicates an I/O error
    IoError {
        message: String,
        path: String,
        suggestion: String,
    },
}

impl Error for MdpError {}
impl fmt::Display for MdpError {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match self {
            MdpError::ParseError {
                message,
                position,
                suggestion,
            } => {
                write!(f, "Markdown Parsing Error: {}", message)?;
                if let Some(pos) = position {
                    write!(f, " (at position {})", pos)?;
                }
                if let Some(hint) = suggestion {
                    write!(f, "\nSuggestion: {}", hint)?;
                }
                Ok(())
            }
            MdpError::PdfError {
                message,
                path,
                suggestion,
            } => {
                write!(f, "PDF Generation Error: {}", message)?;
                if let Some(p) = path {
                    write!(f, "\nPath: {}", p)?;
                }
                if let Some(hint) = suggestion {
                    write!(f, "\nSuggestion: {}", hint)?;
                }
                Ok(())
            }
            MdpError::FontError {
                font_name,
                message,
                suggestion,
            } => {
                write!(f, "Font Error: Failed to load font '{}'", font_name)?;
                write!(f, "\nReason: {}", message)?;
                write!(f, "\nSuggestion: {}", suggestion)?;
                Ok(())
            }
            MdpError::ConfigError {
                message,
                suggestion,
            } => {
                write!(f, "Configuration Error: {}", message)?;
                write!(f, "\nSuggestion: {}", suggestion)?;
                Ok(())
            }
            MdpError::IoError {
                message,
                path,
                suggestion,
            } => {
                write!(f, "File Error: {}", message)?;
                write!(f, "\nPath: {}", path)?;
                write!(f, "\nSuggestion: {}", suggestion)?;
                Ok(())
            }
        }
    }
}

impl MdpError {
    /// Creates a simple parse error with just a message
    pub fn parse_error(message: impl Into<String>) -> Self {
        MdpError::ParseError {
            message: message.into(),
            position: None,
            suggestion: Some(
                "Check your Markdown syntax for unclosed brackets, quotes, or code blocks"
                    .to_string(),
            ),
        }
    }

    /// Creates a simple PDF error with just a message
    pub fn pdf_error(message: impl Into<String>) -> Self {
        MdpError::PdfError {
            message: message.into(),
            path: None,
            suggestion: Some(
                "Check that the output directory exists and you have write permissions".to_string(),
            ),
        }
    }
}

/// Transforms Markdown content into a styled PDF document and saves it to the specified path.
pub fn parse_into_file(
    markdown: String,
    path: &str,
    config: config::ConfigSource,
    font_config: Option<&fonts::FontConfig>,
) -> Result<(), MdpError> {
    // Validate output path exists
    if let Some(parent) = std::path::Path::new(path).parent() {
        if !parent.as_os_str().is_empty() && !parent.exists() {
            return Err(MdpError::IoError {
                message: format!("Output directory does not exist"),
                path: parent.display().to_string(),
                suggestion: format!("Create the directory first: mkdir -p {}", parent.display()),
            });
        }
    }

    let mut lexer = Lexer::new(markdown);
    let tokens = lexer.parse().map_err(|e| {
        let msg = format!("{:?}", e);
        MdpError::ParseError {
            message: msg.clone(),
            position: None,
            suggestion: Some(if msg.contains("UnexpectedEndOfInput") {
                "Check for unclosed code blocks (```), links, or image tags".to_string()
            } else {
                "Verify your Markdown syntax is valid. Try testing with a simpler document first."
                    .to_string()
            }),
        }
    })?;

    let style = config::load_config_from_source(config);
    let pdf = Pdf::new(tokens, style, font_config)?;
    let document = pdf.render_into_document();

    if let Some(err) = Pdf::render(document, path) {
        return Err(MdpError::PdfError {
            message: err.clone(),
            path: Some(path.to_string()),
            suggestion: Some(if err.contains("Permission") || err.contains("denied") {
                "Check that you have write permissions for this location".to_string()
            } else if err.contains("No such file") {
                "Make sure the output directory exists".to_string()
            } else {
                "Try a different output path or check available disk space".to_string()
            }),
        });
    }

    Ok(())
}

/// Transforms Markdown content into a styled PDF document and returns the PDF data as bytes.
pub fn parse_into_bytes(
    markdown: String,
    config: config::ConfigSource,
    font_config: Option<&fonts::FontConfig>,
) -> Result<Vec<u8>, MdpError> {
    let mut lexer = Lexer::new(markdown);
    let tokens = lexer.parse().map_err(|e| {
        let msg = format!("{:?}", e);
        MdpError::ParseError {
            message: msg.clone(),
            position: None,
            suggestion: Some(if msg.contains("UnexpectedEndOfInput") {
                "Check for unclosed code blocks (```), links, or image tags".to_string()
            } else {
                "Verify your Markdown syntax is valid. Try testing with a simpler document first."
                    .to_string()
            }),
        }
    })?;

    let style = config::load_config_from_source(config);
    let pdf = Pdf::new(tokens, style, font_config)?;
    let document = pdf.render_into_document();

    Pdf::render_to_bytes(document).map_err(|err| MdpError::PdfError {
        message: err,
        path: None,
        suggestion: Some("Check available memory and try with a smaller document".to_string()),
    })
}
