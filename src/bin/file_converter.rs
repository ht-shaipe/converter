//! File Converter CLI
//! 
//! A command-line tool for converting between Word, Excel, PDF, and Markdown formats.

use anyhow::{Context, Result};
use clap::{Parser, Subcommand, ValueEnum};
use file_converter::{FileConverter, FileFormat};
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

impl From<FormatArg> for FileFormat {
    fn from(format: FormatArg) -> Self {
        match format {
            FormatArg::Word => FileFormat::Word,
            FormatArg::Excel => FileFormat::Excel,
            FormatArg::Pdf => FileFormat::Pdf,
            FormatArg::Markdown => FileFormat::Markdown,
        }
    }
}

fn main() -> Result<()> {
    let cli = Cli::parse();

    // Initialize logging
    if cli.verbose {
        env_logger::Builder::from_env(env_logger::Env::default().default_filter_or("info")).init();
    }

    // Create converter with optional output directory
    let mut converter = FileConverter::new();
    if let Some(output_dir) = cli.output_dir {
        converter = converter.with_output_dir(output_dir);
    }

    match cli.command {
        Commands::Convert { input, output, to } => {
            handle_convert(&converter, &input, output, to)?;
        }
        Commands::WordToMd { input, output } => {
            handle_word_to_md(&converter, &input, output)?;
        }
        Commands::MdToWord { input, output } => {
            handle_md_to_word(&converter, &input, output)?;
        }
        Commands::ExcelToMd { input, output } => {
            handle_excel_to_md(&converter, &input, output)?;
        }
        Commands::MdToExcel { input, output } => {
            handle_md_to_excel(&converter, &input, output)?;
        }
        Commands::PdfToMd { input, output } => {
            handle_pdf_to_md(&converter, &input, output)?;
        }
        Commands::MdToPdf { input, output } => {
            handle_md_to_pdf(&converter, &input, output)?;
        }
        Commands::List => {
            handle_list();
        }
    }

    Ok(())
}

fn handle_convert(converter: &FileConverter, input: &PathBuf, output: Option<PathBuf>, to: Option<FormatArg>) -> Result<()> {
    if let Some(output_path) = output {
        // Use provided output path
        converter.convert(input, &output_path)
            .with_context(|| format!("Failed to convert {:?} to {:?}", input, output_path))?;
        println!("Successfully converted {:?} to {:?}", input, output_path);
    } else if let Some(target_format) = to {
        // Use target format to auto-generate output path
        let output_path = converter.convert_auto(input, target_format.into())
            .with_context(|| format!("Failed to convert {:?} to {:?}", input, target_format))?;
        println!("Successfully converted {:?} to {:?}", input, output_path);
    } else {
        anyhow::bail!("Either --output or --to must be specified");
    }

    Ok(())
}

fn handle_word_to_md(converter: &FileConverter, input: &PathBuf, output: Option<PathBuf>) -> Result<()> {
    let output_path = output.unwrap_or_else(|| {
        input.with_extension("md")
    });

    converter.convert(input, &output_path)
        .with_context(|| format!("Failed to convert Word to Markdown"))?;
    
    println!("Successfully converted {:?} to {:?}", input, output_path);
    Ok(())
}

fn handle_md_to_word(converter: &FileConverter, input: &PathBuf, output: Option<PathBuf>) -> Result<()> {
    let output_path = output.unwrap_or_else(|| {
        input.with_extension("docx")
    });

    converter.convert(input, &output_path)
        .with_context(|| format!("Failed to convert Markdown to Word"))?;
    
    println!("Successfully converted {:?} to {:?}", input, output_path);
    Ok(())
}

fn handle_excel_to_md(converter: &FileConverter, input: &PathBuf, output: Option<PathBuf>) -> Result<()> {
    let output_path = output.unwrap_or_else(|| {
        input.with_extension("md")
    });

    converter.convert(input, &output_path)
        .with_context(|| format!("Failed to convert Excel to Markdown"))?;
    
    println!("Successfully converted {:?} to {:?}", input, output_path);
    Ok(())
}

fn handle_md_to_excel(converter: &FileConverter, input: &PathBuf, output: Option<PathBuf>) -> Result<()> {
    let output_path = output.unwrap_or_else(|| {
        input.with_extension("xlsx")
    });

    converter.convert(input, &output_path)
        .with_context(|| format!("Failed to convert Markdown to Excel"))?;
    
    println!("Successfully converted {:?} to {:?}", input, output_path);
    Ok(())
}

fn handle_pdf_to_md(converter: &FileConverter, input: &PathBuf, output: Option<PathBuf>) -> Result<()> {
    let output_path = output.unwrap_or_else(|| {
        input.with_extension("md")
    });

    converter.convert(input, &output_path)
        .with_context(|| format!("Failed to convert PDF to Markdown"))?;
    
    println!("Successfully converted {:?} to {:?}", input, output_path);
    Ok(())
}

fn handle_md_to_pdf(converter: &FileConverter, input: &PathBuf, output: Option<PathBuf>) -> Result<()> {
    let output_path = output.unwrap_or_else(|| {
        input.with_extension("pdf")
    });

    converter.convert(input, &output_path)
        .with_context(|| format!("Failed to convert Markdown to PDF"))?;
    
    println!("Successfully converted {:?} to {:?}", input, output_path);
    Ok(())
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
