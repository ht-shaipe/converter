//! 命令行：按子命令调用库内转换实现。

use clap::{Parser, Subcommand};
use file_converter::{run_conversion, ConversionType};
use std::path::PathBuf;

#[derive(Parser)]
#[command(name = "file_converter")]
#[command(about = "DOCX/XLSX/PDF/Markdown/图片 与目标格式之间的转换", long_about = None)]
struct Cli {
    #[command(subcommand)]
    command: Commands,
}

#[derive(Subcommand)]
enum Commands {
    /// DOCX → Markdown
    DocxToMd {
        #[arg(short, long)]
        input: PathBuf,
        #[arg(short, long)]
        output: PathBuf,
    },
    /// Markdown → DOCX
    MdToDocx {
        #[arg(short, long)]
        input: PathBuf,
        #[arg(short, long)]
        output: PathBuf,
    },
    /// Excel（xlsx/xls/…）→ Markdown
    XlsxToMd {
        #[arg(short, long)]
        input: PathBuf,
        #[arg(short, long)]
        output: PathBuf,
    },
    /// Markdown → XLSX
    MdToXlsx {
        #[arg(short, long)]
        input: PathBuf,
        #[arg(short, long)]
        output: PathBuf,
    },
    /// PDF → Markdown（文本抽取）
    PdfToMd {
        #[arg(short, long)]
        input: PathBuf,
        #[arg(short, long)]
        output: PathBuf,
    },
    /// Markdown → PDF（简单文本页）
    MdToPdf {
        #[arg(short, long)]
        input: PathBuf,
        #[arg(short, long)]
        output: PathBuf,
    },
    /// 位图（png/jpg/…）→ ICO
    ImgToIco {
        #[arg(short, long)]
        input: PathBuf,
        #[arg(short, long)]
        output: PathBuf,
    },
}

fn main() -> anyhow::Result<()> {
    env_logger::Builder::from_default_env().format_timestamp(None).init();
    let cli = Cli::parse();
    let ct = match &cli.command {
        Commands::DocxToMd { .. } => ConversionType::WordToMarkdown,
        Commands::MdToDocx { .. } => ConversionType::MarkdownToWord,
        Commands::XlsxToMd { .. } => ConversionType::ExcelToMarkdown,
        Commands::MdToXlsx { .. } => ConversionType::MarkdownToExcel,
        Commands::PdfToMd { .. } => ConversionType::PdfToMarkdown,
        Commands::MdToPdf { .. } => ConversionType::MarkdownToPdf,
        Commands::ImgToIco { .. } => ConversionType::ImageToIco,
    };
    let (input, output) = match cli.command {
        Commands::DocxToMd { input, output }
        | Commands::MdToDocx { input, output }
        | Commands::XlsxToMd { input, output }
        | Commands::MdToXlsx { input, output }
        | Commands::PdfToMd { input, output }
        | Commands::MdToPdf { input, output }
        | Commands::ImgToIco { input, output } => (input, output),
    };
    run_conversion(ct, &input, &output)?;
    log::info!("完成: {:?} → {:?}", input, output);
    Ok(())
}
