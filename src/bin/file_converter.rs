//! File Converter CLI
//! 
//! A command-line tool for converting between Word, Excel, PDF, and Markdown formats.
//!
//! # Usage
//!
//! ```bash
//! # Convert Word to Markdown
//! file_converter word-to-md document.docx
//!
//! # Convert Markdown to Word
//! file_converter md-to-word readme.md
//!
//! # Convert Excel to Markdown
//! file_converter excel-to-md data.xlsx
//!
//! # Convert Markdown to Excel
//! file_converter md-to-excel tables.md
//!
//! # Convert PDF to Markdown
//! file_converter pdf-to-md report.pdf
//!
//! # Convert Markdown to PDF
//! file_converter md-to-pdf notes.md
//!
//! # List all supported conversions
//! file_converter list
//! ```

use anyhow::{Context, Result};
use clap::{Parser, Subcommand, ValueEnum};
use std::path::PathBuf;

#[derive(Parser)]
#[command(name = "file_converter")]
#[command(author = "Your Name")]
#[command(version = "0.1.0")]
#[command(about = "Convert files between Word, Excel, PDF, and Markdown formats", long_about = None)]
struct Cli {
    #[command(subcommand)]
    command: Commands,

    /// Enable verbose output
    #[arg(short, long, global = true)]
    verbose: bool,

    /// Output directory for converted files
    #[arg(short, long, global = true)]
    output_dir: Option<PathBuf>,
}

#[derive(Subcommand)]
enum Commands {
    /// Convert a file to another format
    Convert {
        /// Input file path
        input: PathBuf,

        /// Output file path (optional, auto-generated if not provided)
        #[arg(short, long)]
        output: Option<PathBuf>,

        /// Target format (optional, inferred from output extension if not provided)
        #[arg(short, long)]
        to: Option<FormatArg>,
    },

    /// Convert Word document to Markdown
    WordToMd {
        /// Input DOCX file
        input: PathBuf,

        /// Output MD file (optional)
        #[arg(short, long)]
        output: Option<PathBuf>,
    },

    /// Convert Markdown to Word document
    MdToWord {
        /// Input MD file
        input: PathBuf,

        /// Output DOCX file (optional)
        #[arg(short, long)]
        output: Option<PathBuf>,
    },

    /// Convert Excel spreadsheet to Markdown
    ExcelToMd {
        /// Input XLSX file
        input: PathBuf,

        /// Output MD file (optional)
        #[arg(short, long)]
        output: Option<PathBuf>,
    },

    /// Convert Markdown to Excel spreadsheet
    MdToExcel {
        /// Input MD file
        input: PathBuf,

        /// Output XLSX file (optional)
        #[arg(short, long)]
        output: Option<PathBuf>,
    },

    /// Convert PDF to Markdown
    PdfToMd {
        /// Input PDF file
        input: PathBuf,

        /// Output MD file (optional)
        #[arg(short, long)]
        output: Option<PathBuf>,
    },

    /// Convert Markdown to PDF
    MdToPdf {
        /// Input MD file
        input: PathBuf,

        /// Output PDF file (optional)
        #[arg(short, long)]
        output: Option<PathBuf>,
    },

    /// List supported conversions
    List,
}

#[derive(ValueEnum, Clone, Debug)]
enum FormatArg {
    Word,
    Excel,
    Pdf,
    Markdown,
}

fn main() -> Result<()> {
    let cli = Cli::parse();

    // Initialize logging
    if cli.verbose {
        env_logger::Builder::from_env(env_logger::Env::default().default_filter_or("info")).init();
    }

    match cli.command {
        Commands::Convert { input, output, to } => {
            handle_convert(&input, output, to)?;
        }
        Commands::WordToMd { input, output } => {
            handle_word_to_md(&input, output)?;
        }
        Commands::MdToWord { input, output } => {
            handle_md_to_word(&input, output)?;
        }
        Commands::ExcelToMd { input, output } => {
            handle_excel_to_md(&input, output)?;
        }
        Commands::MdToExcel { input, output } => {
            handle_md_to_excel(&input, output)?;
        }
        Commands::PdfToMd { input, output } => {
            handle_pdf_to_md(&input, output)?;
        }
        Commands::MdToPdf { input, output } => {
            handle_md_to_pdf(&input, output)?;
        }
        Commands::List => {
            handle_list();
        }
    }

    Ok(())
}

fn handle_convert(input: &PathBuf, output: Option<PathBuf>, to: Option<FormatArg>) -> Result<()> {
    use file_converter::{FileFormat, docx_to_md, md_to_docx, xlsx_to_md, md_to_xlsx, pdf_to_md, md_to_pdf};
    
    let source_format = FileFormat::from_path(input)?;
    
    if let Some(output_path) = output {
        let target_format = FileFormat::from_extension(&output_path);
        
        match (source_format, target_format) {
            (FileFormat::Word, FileFormat::Markdown) => docx_to_md(input, &output_path),
            (FileFormat::Markdown, FileFormat::Word) => md_to_docx(input, &output_path),
            (FileFormat::Excel, FileFormat::Markdown) => xlsx_to_md(input, &output_path),
            (FileFormat::Markdown, FileFormat::Excel) => md_to_xlsx(input, &output_path),
            (FileFormat::Pdf, FileFormat::Markdown) => pdf_to_md(input, &output_path),
            (FileFormat::Markdown, FileFormat::Pdf) => md_to_pdf(input, &output_path),
            _ => anyhow::bail!("Unsupported conversion: {:?} to {:?}", source_format, target_format),
        }?;
        
        println!("Successfully converted {:?} to {:?}", input, output_path);
    } else if let Some(target_format_arg) = to {
        let target_format = match target_format_arg {
            FormatArg::Word => FileFormat::Word,
            FormatArg::Excel => FileFormat::Excel,
            FormatArg::Pdf => FileFormat::Pdf,
            FormatArg::Markdown => FileFormat::Markdown,
        };
        
        let output_path = generate_output_path(input, &target_format);
        
        match (source_format, target_format) {
            (FileFormat::Word, FileFormat::Markdown) => docx_to_md(input, &output_path),
            (FileFormat::Markdown, FileFormat::Word) => md_to_docx(input, &output_path),
            (FileFormat::Excel, FileFormat::Markdown) => xlsx_to_md(input, &output_path),
            (FileFormat::Markdown, FileFormat::Excel) => md_to_xlsx(input, &output_path),
            (FileFormat::Pdf, FileFormat::Markdown) => pdf_to_md(input, &output_path),
            (FileFormat::Markdown, FileFormat::Pdf) => md_to_pdf(input, &output_path),
            _ => anyhow::bail!("Unsupported conversion: {:?} to {:?}", source_format, target_format),
        }?;
        
        println!("Successfully converted {:?} to {:?}", input, output_path);
    } else {
        anyhow::bail!("Either --output or --to must be specified");
    }

    Ok(())
}

fn handle_word_to_md(input: &PathBuf, output: Option<PathBuf>) -> Result<()> {
    use file_converter::docx_to_md;
    
    let output_path = output.unwrap_or_else(|| {
        input.with_extension("md")
    });

    docx_to_md(input, &output_path)
        .with_context(|| format!("Failed to convert Word to Markdown"))?;
    
    println!("Successfully converted {:?} to {:?}", input, output_path);
    Ok(())
}

fn handle_md_to_word(input: &PathBuf, output: Option<PathBuf>) -> Result<()> {
    use file_converter::md_to_docx;
    
    let output_path = output.unwrap_or_else(|| {
        input.with_extension("docx")
    });

    md_to_docx(input, &output_path)
        .with_context(|| format!("Failed to convert Markdown to Word"))?;
    
    println!("Successfully converted {:?} to {:?}", input, output_path);
    Ok(())
}

fn handle_excel_to_md(input: &PathBuf, output: Option<PathBuf>) -> Result<()> {
    use file_converter::xlsx_to_md;
    
    let output_path = output.unwrap_or_else(|| {
        input.with_extension("md")
    });

    xlsx_to_md(input, &output_path)
        .with_context(|| format!("Failed to convert Excel to Markdown"))?;
    
    println!("Successfully converted {:?} to {:?}", input, output_path);
    Ok(())
}

fn handle_md_to_excel(input: &PathBuf, output: Option<PathBuf>) -> Result<()> {
    use file_converter::md_to_xlsx;
    
    let output_path = output.unwrap_or_else(|| {
        input.with_extension("xlsx")
    });

    md_to_xlsx(input, &output_path)
        .with_context(|| format!("Failed to convert Markdown to Excel"))?;
    
    println!("Successfully converted {:?} to {:?}", input, output_path);
    Ok(())
}

fn handle_pdf_to_md(input: &PathBuf, output: Option<PathBuf>) -> Result<()> {
    use file_converter::pdf_to_md;
    
    let output_path = output.unwrap_or_else(|| {
        input.with_extension("md")
    });

    pdf_to_md(input, &output_path)
        .with_context(|| format!("Failed to convert PDF to Markdown"))?;
    
    println!("Successfully converted {:?} to {:?}", input, output_path);
    Ok(())
}

fn handle_md_to_pdf(input: &PathBuf, output: Option<PathBuf>) -> Result<()> {
    use file_converter::md_to_pdf;
    
    let output_path = output.unwrap_or_else(|| {
        input.with_extension("pdf")
    });

    md_to_pdf(input, &output_path)
        .with_context(|| format!("Failed to convert Markdown to PDF"))?;
    
    println!("Successfully converted {:?} to {:?}", input, output_path);
    Ok(())
}

fn generate_output_path(input: &PathBuf, target_format: &FileFormat) -> PathBuf {
    let stem = input.file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("output");
    
    let dir = input.parent().unwrap_or_else(|| std::path::Path::new("."));
    
    dir.join(format!("{}.{}", stem, target_format.extension()))
}

fn handle_list() {
    println!("Supported conversions:");
    println!();
    println!("  Word (DOCX) <-> Markdown");
    println!("  Excel (XLSX) <-> Markdown");
    println!("  PDF <-> Markdown");
    println!();
    println!("Examples:");
    println!("  file_converter word-to-md document.docx");
    println!("  file_converter md-to-word readme.md");
    println!("  file_converter excel-to-md data.xlsx");
    println!("  file_converter md-to-excel tables.md");
    println!("  file_converter pdf-to-md report.pdf");
    println!("  file_converter md-to-pdf notes.md");
    println!();
    println!("Or use the generic convert command:");
    println!("  file_converter convert input.docx --to markdown");
    println!("  file_converter convert input.md --output output.pdf");
}
