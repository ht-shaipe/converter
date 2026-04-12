//! Unit tests for file_converter

#[cfg(test)]
mod tests {
    use file_converter::{FileConverter, FileFormat};
    use std::path::Path;

    #[test]
    fn test_file_format_from_extension() {
        assert_eq!(FileFormat::from_extension(Path::new("test.docx")), FileFormat::Word);
        assert_eq!(FileFormat::from_extension(Path::new("test.xlsx")), FileFormat::Excel);
        assert_eq!(FileFormat::from_extension(Path::new("test.xls")), FileFormat::Excel);
        assert_eq!(FileFormat::from_extension(Path::new("test.pdf")), FileFormat::Pdf);
        assert_eq!(FileFormat::from_extension(Path::new("test.md")), FileFormat::Markdown);
        assert_eq!(FileFormat::from_extension(Path::new("test.markdown")), FileFormat::Markdown);
        assert_eq!(FileFormat::from_extension(Path::new("test.txt")), FileFormat::Unknown);
    }

    #[test]
    fn test_file_format_extension() {
        assert_eq!(FileFormat::Word.extension(), "docx");
        assert_eq!(FileFormat::Excel.extension(), "xlsx");
        assert_eq!(FileFormat::Pdf.extension(), "pdf");
        assert_eq!(FileFormat::Markdown.extension(), "md");
        assert_eq!(FileFormat::Unknown.extension(), "");
    }

    #[test]
    fn test_conversion_support() {
        // Word to Markdown should be supported
        assert!(FileFormat::Word.can_convert_to(&FileFormat::Markdown));
        
        // Markdown to Word should be supported
        assert!(FileFormat::Markdown.can_convert_to(&FileFormat::Word));
        
        // Excel to Markdown should be supported
        assert!(FileFormat::Excel.can_convert_to(&FileFormat::Markdown));
        
        // Markdown to Excel should be supported
        assert!(FileFormat::Markdown.can_convert_to(&FileFormat::Excel));
        
        // PDF to Markdown should be supported
        assert!(FileFormat::Pdf.can_convert_to(&FileFormat::Markdown));
        
        // Markdown to PDF should be supported
        assert!(FileFormat::Markdown.can_convert_to(&FileFormat::Pdf));
        
        // Direct Word to Excel should NOT be supported
        assert!(!FileFormat::Word.can_convert_to(&FileFormat::Excel));
        
        // Direct PDF to Excel should NOT be supported
        assert!(!FileFormat::Pdf.can_convert_to(&FileFormat::Excel));
    }

    #[test]
    fn test_converter_creation() {
        let converter = FileConverter::new();
        // Just test that we can create a converter
        assert!(true);
    }

    #[test]
    fn test_converter_with_output_dir() {
        let converter = FileConverter::new()
            .with_output_dir(std::path::PathBuf::from("/tmp/test"));
        // Just test that we can create a converter with output dir
        assert!(true);
    }

    #[test]
    fn test_file_not_found_error() {
        let converter = FileConverter::new();
        let result = converter.convert(
            Path::new("nonexistent.docx"),
            Path::new("output.md")
        );
        
        assert!(result.is_err());
        let err = result.unwrap_err();
        assert!(err.to_string().contains("File not found"));
    }

    #[test]
    fn test_unsupported_format_error() {
        let converter = FileConverter::new();
        let result = converter.convert(
            Path::new("test.txt"),
            Path::new("output.md")
        );
        
        assert!(result.is_err());
        let err = result.unwrap_err();
        assert!(err.to_string().contains("Unsupported") || err.to_string().contains("not found"));
    }

    #[test]
    fn test_unsupported_conversion_error() {
        let converter = FileConverter::new();
        // Create a temporary markdown file for testing
        let temp_md = tempfile::NamedTempFile::new().unwrap();
        std::fs::write(temp_md.path(), "# Test").unwrap();
        
        // Try to convert markdown to an unsupported format (e.g., directly to another markdown)
        let temp_out = tempfile::NamedTempFile::new().unwrap();
        let result = converter.convert(
            temp_md.path(),
            temp_out.path()
        );
        
        // This should fail because we're not specifying a valid target format
        assert!(result.is_err());
    }
}
