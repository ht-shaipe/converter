//! Excel converter

use crate::error::{ConverterError, Result};
use crate::formats::FileFormat;
use calamine::{open_workbook_auto, Data, Reader};
use rust_xlsxwriter::Workbook;
use std::fs;
use std::path::Path;

fn cell_to_string(c: &Data) -> String {
    match c {
        Data::Empty => String::new(),
        Data::String(s) => s.clone(),
        Data::Float(f) => f.to_string(),
        Data::Int(i) => i.to_string(),
        Data::Bool(b) => b.to_string(),
        Data::DateTime(dt) => format!("{dt:?}"),
        Data::DateTimeIso(s) => s.clone(),
        Data::DurationIso(s) => s.clone(),
        Data::Error(e) => format!("{e:?}"),
    }
}

pub fn xlsx_to_md(input: &Path, output: &Path) -> Result<()> {
    if !input.exists() {
        return Err(ConverterError::FileNotFound(input.display().to_string()));
    }
    let fmt = FileFormat::from_extension(input);
    if !matches!(fmt, FileFormat::Excel) {
        return Err(ConverterError::UnsupportedFormat(
            input.extension().and_then(|e| e.to_str()).unwrap_or("").to_string(),
        ));
    }

    log::info!("Converting spreadsheet to Markdown: {:?}", input);
    let mut workbook = open_workbook_auto(input).map_err(|e| ConverterError::ExcelError(e.to_string()))?;
    let mut out = String::new();

    for sheet in workbook.sheet_names().to_vec() {
        out.push_str(&format!("## {sheet}\n\n"));
        let range = workbook
            .worksheet_range(&sheet)
            .map_err(|e| ConverterError::ExcelError(e.to_string()))?;
        for row in range.rows() {
            let cells: Vec<String> = row.iter().map(cell_to_string).collect();
            if cells.iter().all(|c| c.is_empty()) {
                continue;
            }
            out.push('|');
            for c in &cells {
                out.push(' ');
                out.push_str(c);
                out.push_str(" |");
            }
            out.push('\n');
        }
        out.push('\n');
    }

    fs::write(output, out)?;
    Ok(())
}

pub fn md_to_xlsx(input: &Path, output: &Path) -> Result<()> {
    if !input.exists() {
        return Err(ConverterError::FileNotFound(input.display().to_string()));
    }
    let fmt = FileFormat::from_extension(input);
    if fmt != FileFormat::Markdown && fmt != FileFormat::Unknown {
        return Err(ConverterError::UnsupportedFormat(
            input.extension().and_then(|e| e.to_str()).unwrap_or("").to_string(),
        ));
    }

    log::info!("Converting Markdown to XLSX: {:?}", input);
    let md = fs::read_to_string(input)?;
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    for (row, line) in md.lines().enumerate() {
        let row_u32 = u32::try_from(row).map_err(|_| ConverterError::ExcelError("too many rows".into()))?;
        if line.contains('\t') {
            for (col, cell) in line.split('\t').enumerate() {
                let col_u16 = u16::try_from(col).map_err(|_| ConverterError::ExcelError("too many columns".into()))?;
                worksheet
                    .write_string(row_u32, col_u16, cell)
                    .map_err(|e| ConverterError::ExcelError(e.to_string()))?;
            }
        } else if line.contains('|') {
            let parts: Vec<&str> = line.split('|').map(str::trim).filter(|p| !p.is_empty()).collect();
            for (col, cell) in parts.iter().enumerate() {
                let col_u16 = u16::try_from(col).map_err(|_| ConverterError::ExcelError("too many columns".into()))?;
                worksheet
                    .write_string(row_u32, col_u16, *cell)
                    .map_err(|e| ConverterError::ExcelError(e.to_string()))?;
            }
        } else {
            worksheet
                .write_string(row_u32, 0, line)
                .map_err(|e| ConverterError::ExcelError(e.to_string()))?;
        }
    }

    workbook
        .save(output)
        .map_err(|e| ConverterError::ExcelError(e.to_string()))?;
    Ok(())
}
