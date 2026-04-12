//! Unit tests for file_converter

#[cfg(test)]
mod tests {
    use file_converter::{FileFormat, docx_to_md, md_to_docx, xlsx_to_md, md_to_xlsx, pdf_to_md, md_to_pdf};
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
    fn test_word_converter_functions_exist() {
        // Just verify the functions are exported and callable
        let temp_in = tempfile::NamedTempFile::new().unwrap();
        let temp_out = tempfile::NamedTempFile::new().unwrap();
        
        // Write some dummy content
        std::fs::write(temp_in.path(), "# Test").unwrap();
        
        // These will fail but prove the functions exist
        let _ = docx_to_md(temp_in.path(), temp_out.path());
        let _ = md_to_docx(temp_in.path(), temp_out.path());
    }

    #[test]
    fn test_excel_converter_functions_exist() {
        let temp_in = tempfile::NamedTempFile::new().unwrap();
        let temp_out = tempfile::NamedTempFile::new().unwrap();
        
        std::fs::write(temp_in.path(), "# Test").unwrap();
        
        let _ = xlsx_to_md(temp_in.path(), temp_out.path());
        let _ = md_to_xlsx(temp_in.path(), temp_out.path());
    }

    #[test]
    fn test_pdf_converter_functions_exist() {
        let temp_in = tempfile::NamedTempFile::new().unwrap();
        let temp_out = tempfile::NamedTempFile::new().unwrap();
        
        std::fs::write(temp_in.path(), "# Test").unwrap();
        
        let _ = pdf_to_md(temp_in.path(), temp_out.path());
        let _ = md_to_pdf(temp_in.path(), temp_out.path());
    }

    #[test]
    fn test_file_not_found_error() {
        let result = docx_to_md(
            Path::new("nonexistent.docx"),
            Path::new("output.md")
        );
        
        assert!(result.is_err());
    }

    #[test]
    fn test_unsupported_format_error() {
        let temp_txt = tempfile::NamedTempFile::new().unwrap();
        std::fs::write(temp_txt.path(), "test").unwrap();
        
        let result = docx_to_md(
            temp_txt.path(),
            Path::new("output.md")
        );
        
        assert!(result.is_err());
    }
}
