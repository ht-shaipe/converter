//! MCP（stdio）：暴露 `list_kinds` 与 `convert` 工具，内部调用本库 `run_conversion`。

use file_converter::{run_conversion, ConversionType};
use rmcp::{
    handler::server::wrapper::Parameters, schemars, serde, tool, tool_router, transport::stdio,
    ServiceExt,
};
use std::path::Path;

#[derive(Debug, serde::Deserialize, schemars::JsonSchema)]
struct ConvertArgs {
    #[schemars(description = "转换类型 id，见 list_kinds：docx_md、md_docx、xlsx_md、md_xlsx、pdf_md、md_pdf、img_ico")]
    kind: String,
    #[schemars(description = "输入文件路径")]
    input: String,
    #[schemars(description = "输出文件路径")]
    output: String,
}

#[derive(Clone)]
struct ConverterMcp;

#[tool_router(server_handler)]
impl ConverterMcp {
    #[tool(description = "列出所有支持的 kind_id 及说明")]
    fn list_kinds(&self) -> String {
        use ConversionType::*;
        [
            WordToMarkdown,
            MarkdownToWord,
            ExcelToMarkdown,
            MarkdownToExcel,
            PdfToMarkdown,
            MarkdownToPdf,
            ImageToIco,
        ]
        .iter()
        .map(|c| format!("{} — {}", c.kind_id(), c.description()))
        .collect::<Vec<_>>()
        .join("\n")
    }

    #[tool(description = "执行一次文件转换")]
    fn convert(
        &self,
        Parameters(ConvertArgs { kind, input, output }): Parameters<ConvertArgs>,
    ) -> String {
        let kind_trim = kind.trim();
        let Some(ct) = ConversionType::from_kind_id(kind_trim) else {
            return format!("未知的 kind: {kind}。请先调用 list_kinds。");
        };
        match run_conversion(ct, Path::new(input.trim()), Path::new(output.trim())) {
            Ok(()) => format!("成功: {} {:?} → {:?}", kind_trim, input, output),
            Err(e) => format!("失败: {e}"),
        }
    }
}

#[tokio::main]
async fn main() -> anyhow::Result<()> {
    tracing_subscriber::fmt()
        .with_env_filter(tracing_subscriber::EnvFilter::from_default_env())
        .with_writer(std::io::stderr)
        .with_ansi(false)
        .init();

    let service = ConverterMcp.serve(stdio()).await?;
    service.waiting().await?;
    Ok(())
}
