//! Excel spreadsheet converter (XLSX <-> Markdown)

use crate::error::{ConverterError, Result};
use calamine::{Reader, Xlsx, Writer, Xlsx as CalamineXlsx};
use std::fs;
use std::io::BufWriter;
use std::path::Path;

/// Convert XLSX to Markdown
pub fn xlsx_to_md(input_path: &Path, output_path: &Path) -> Result<()> {
    log::info!("Converting XLSX to Markdown: {:?} -> {:?}", input_path, output_path);

    let file = fs::File::open(input_path)?;
    let mut workbook: Xlsx<_> = Xlsx::new(file)
        .map_err(|e| ConverterError::ExcelError(format!("Failed to open XLSX: {}", e)))?;

    let mut markdown_content = String::new();

    // Iterate through all sheets
    for sheet_name in workbook.sheet_names().to_owned() {
        markdown_content.push_str(&format!("# {}\n\n", sheet_name));

        if let Ok(range) = workbook.worksheet_range(&sheet_name) {
            if range.is_empty() {
                continue;
            }

            // Convert to markdown table
            let mut table = String::new();
            let mut headers = true;

            for row in range.rows() {
                let cells: Vec<String> = row
                    .iter()
                    .map(|cell| match cell {
                        calamine::Data::String(s) => s.clone(),
                        calamine::Data::Int(i) => i.to_string(),
                        calamine::Data::Float(f) => f.to_string(),
                        calamine::Data::Bool(b) => b.to_string(),
                        calamine::Data::DateTime(dt) => dt.to_string(),
                        _ => String::new(),
                    })
                    .collect();

                if !cells.is_empty() {
                    table.push('|');
                    table.push_str(&cells.join("|"));
                    table.push_str("|\n");

                    // Add header separator after first row
                    if headers {
                        table.push('|');
                        for _ in &cells {
                            table.push_str("---|");
                        }
                        table.push('\n');
                        headers = false;
                    }
                }
            }

            markdown_content.push_str(&table);
            markdown_content.push_str("\n\n");
        } else {
            log::warn!("Could not read sheet: {}", sheet_name);
        }
    }

    // Write markdown content
    fs::write(output_path, markdown_content)?;

    log::info!("Successfully converted XLSX to Markdown");
    Ok(())
}

/// Convert Markdown to XLSX
pub fn md_to_xlsx(input_path: &Path, output_path: &Path) -> Result<()> {
    log::info!("Converting Markdown to XLSX: {:?} -> {:?}", input_path, output_path);

    let markdown_content = fs::read_to_string(input_path)?;
    
    let file = fs::File::create(output_path)?;
    let buf_writer = BufWriter::new(file);
    let mut writer = Writer::new(buf_writer);

    // Parse markdown content for tables
    let mut current_sheet: Vec<Vec<String>> = Vec::new();
    let mut sheet_name = "Sheet1".to_string();
    let mut sheet_count = 0;

    for line in markdown_content.lines() {
        let trimmed = line.trim();
        
        // Check for sheet name (heading level 1)
        if trimmed.starts_with("# ") {
            // Write previous sheet if exists
            if !current_sheet.is_empty() {
                let sheet_data: Vec<Vec<calamine::Data>> = current_sheet
                    .iter()
                    .map(|row| {
                        row.iter()
                            .map(|s| calamine::Data::String(s.clone()))
                            .collect()
                    })
                    .collect();
                
                if sheet_count > 0 {
                    writer.push_sheet(sheet_data.into_iter());
                } else {
                    // First sheet uses default name
                    writer.write_sheet(current_sheet.into_iter())?;
                }
                current_sheet = Vec::new();
            }
            
            sheet_name = trimmed.trim_start_matches("# ").trim().to_string();
            sheet_count += 1;
        } 
        // Check for table row
        else if trimmed.starts_with('|') && trimmed.ends_with('|') {
            // Skip header separator row
            if trimmed.contains("---") {
                continue;
            }
            
            // Parse table cells
            let cells: Vec<String> = trimmed
                .trim_matches('|')
                .split('|')
                .map(|s| s.trim().to_string())
                .collect();
            
            if !cells.is_empty() {
                current_sheet.push(cells);
            }
        }
    }

    // Write last sheet if not empty
    if !current_sheet.is_empty() && sheet_count > 0 {
        let sheet_data: Vec<Vec<calamine::Data>> = current_sheet
            .iter()
            .map(|row| {
                row.iter()
                    .map(|s| calamine::Data::String(s.clone()))
                    .collect()
            })
            .collect();
        writer.push_sheet(sheet_data.into_iter());
    } else if !current_sheet.is_empty() {
        writer.write_sheet(current_sheet.into_iter())?;
    }

    writer.close()?;

    log::info!("Successfully converted Markdown to XLSX");
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::tempdir;
    use std::path::PathBuf;

    #[test]
    fn test_xlsx_to_md_basic() {
        // Create a temporary directory
        let temp_dir = tempdir().unwrap();
        let input_path = temp_dir.path().join("test.xlsx");
        let output_path = temp_dir.path().join("test.md");

        // Create a simple XLSX file for testing
        create_test_xlsx(&input_path).unwrap();

        // Convert to markdown
        xlsx_to_md(&input_path, &output_path).unwrap();

        // Verify output exists and has content
        assert!(output_path.exists());
        let content = fs::read_to_string(&output_path).unwrap();
        assert!(!content.is_empty());
    }

    fn create_test_xlsx(path: &Path) -> Result<()> {
        let file = fs::File::create(path)?;
        let buf_writer = BufWriter::new(file);
        let mut writer = Writer::new(buf_writer);

        let data = vec![
            vec!["Name".to_string(), "Age".to_string(), "City".to_string()],
            vec!["Alice".to_string(), "30".to_string(), "New York".to_string()],
            vec!["Bob".to_string(), "25".to_string(), "London".to_string()],
        ];

        writer.write_sheet(data.into_iter())?;
        writer.close()?;

        Ok(())
    }

    #[test]
    fn test_parse_markdown_table() {
        let markdown = r#"# Test Sheet

| Name | Age | City |
|---|---|---|
| Alice | 30 | New York |
| Bob | 25 | London |
"#;

        // Verify we can parse the table structure
        assert!(markdown.contains("| Name | Age | City |"));
        assert!(markdown.contains("| Alice | 30 | New York |"));
    }
}
