# File Converter

A Rust-based file converter crate and CLI tool for converting between Word, Excel, PDF, and Markdown formats.

## Features

- **Word Documents (DOCX)** ↔ **Markdown**
- **Excel Spreadsheets (XLSX)** ↔ **Markdown**
- **PDF** ↔ **Markdown**

## Installation

### From Source

```bash
git clone https://github.com/yourusername/file_converter.git
cd file_converter
cargo build --release
```

The binary will be available at `target/release/file_converter`.

### Add as a Dependency

Add this to your `Cargo.toml`:

```toml
[dependencies]
file_converter = "0.1.0"
```

## Usage

### CLI Tool

#### List Supported Conversions

```bash
file_converter list
```

#### Convert Word to Markdown

```bash
file_converter word-to-md document.docx
file_converter word-to-md document.docx -o output.md
```

#### Convert Markdown to Word

```bash
file_converter md-to-word readme.md
file_converter md-to-word readme.md -o document.docx
```

#### Convert Excel to Markdown

```bash
file_converter excel-to-md data.xlsx
file_converter excel-to-md data.xlsx -o output.md
```

#### Convert Markdown to Excel

```bash
file_converter md-to-excel tables.md
file_converter md-to-excel tables.md -o data.xlsx
```

#### Convert PDF to Markdown

```bash
file_converter pdf-to-md report.pdf
file_converter pdf-to-md report.pdf -o output.md
```

#### Convert Markdown to PDF

```bash
file_converter md-to-pdf notes.md
file_converter md-to-pdf notes.md -o document.pdf
```

#### Generic Convert Command

```bash
# Convert using target format
file_converter convert input.docx --to markdown

# Convert with specific output path
file_converter convert input.md --output output.pdf
```

#### Options

- `-v, --verbose`: Enable verbose output
- `-o, --output-dir`: Set output directory for converted files

### Library Usage

```rust
use file_converter::{FileConverter, FileFormat};
use std::path::Path;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Create a new converter
    let converter = FileConverter::new();
    
    // Convert DOCX to Markdown
    converter.convert(
        Path::new("document.docx"),
        Path::new("output.md")
    )?;
    
    // Convert with automatic output path generation
    let output_path = converter.convert_auto(
        Path::new("data.xlsx"),
        FileFormat::Markdown
    )?;
    
    println!("Converted file: {:?}", output_path);
    
    Ok(())
}
```

```rust
// Custom output directory
let converter = FileConverter::new()
    .with_output_dir(PathBuf::from("/output/path"));
```

## Supported Formats

| Format | Extension | Description |
|--------|-----------|-------------|
| Word | `.docx` | Microsoft Word documents |
| Excel | `.xlsx`, `.xls` | Microsoft Excel spreadsheets |
| PDF | `.pdf` | Portable Document Format |
| Markdown | `.md`, `.markdown` | Markdown text files |

## Conversion Matrix

| From \ To | Word | Excel | PDF | Markdown |
|-----------|------|-------|-----|----------|
| Word      | -    | ❌    | ❌  | ✅       |
| Excel     | ❌   | -     | ❌  | ✅       |
| PDF       | ❌   | ❌    | -   | ✅       |
| Markdown  | ✅   | ✅    | ✅  | -        |

## Architecture

The project is organized into the following modules:

- **`error`**: Error types and result handling
- **`formats`**: File format detection and conversion type definitions
- **`converter`**: Core conversion implementations
  - `word`: Word document conversions
  - `excel`: Excel spreadsheet conversions
  - `pdf`: PDF conversions

## Dependencies

- **clap**: CLI argument parsing
- **anyhow/thisterror**: Error handling
- **docx-rs**: Word document creation
- **zip**: DOCX archive handling
- **calamine**: Excel file reading/writing
- **lopdf**: PDF processing
- **pulldown-cmark**: Markdown parsing

## Development

### Build

```bash
cargo build
```

### Run Tests

```bash
cargo test
```

### Run with Verbose Output

```bash
RUST_LOG=info cargo run -- word-to-md test.docx -v
```

## Limitations

- **PDF Text Extraction**: PDF to Markdown conversion uses basic text extraction. Complex PDFs with images, tables, or complex layouts may not convert perfectly.
- **Markdown to PDF**: The current implementation creates simple PDFs. For production use with complex formatting, consider using a dedicated MD→HTML→PDF pipeline.
- **DOCX Formatting**: Some advanced Word formatting may be lost during conversion.

## Roadmap

- [ ] Improved PDF text extraction with layout preservation
- [ ] Better handling of tables in all formats
- [ ] Image extraction and embedding
- [ ] Support for more file formats (PPTX, HTML, etc.)
- [ ] Batch conversion support
- [ ] Configuration file support

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## Acknowledgments

- Thanks to all the Rust community members who created the underlying libraries
- Inspired by the need for reliable file format conversions in document workflows
