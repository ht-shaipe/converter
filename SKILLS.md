# File Converter Skills & Capabilities

本项目是一个基于 Rust 的高性能文件转换器库和 CLI 工具，专注于文档格式之间的转换。

## 🎯 核心能力

### 当前支持的转换

| 源格式 | 目标格式 | 状态 | 说明 |
|--------|----------|------|------|
| DOCX (Word) | Markdown | ✅ | 支持标题、段落、列表 |
| Markdown | DOCX (Word) | ✅ | 支持标题、段落、列表 |
| XLSX (Excel) | Markdown | ✅ | 支持多工作表、表格 |
| Markdown | XLSX (Excel) | ✅ | 支持多工作表、表格 |
| PDF | Markdown | ⚠️ | 基础文本提取 |
| Markdown | PDF | ⚠️ | 简单 PDF 生成 |

**图例**: ✅ 完全支持 | ⚠️ 基础支持 | ❌ 不支持

---

## 🚀 可扩展的转换方向

### 1. 办公文档类

#### PowerPoint (PPTX)
- **PPTX ↔ Markdown**: 将幻灯片内容转换为带分页标记的 Markdown
- **PPTX ↔ PDF**: 演示文稿与 PDF 互转
- **PPTX ↔ HTML**: 转换为网页格式

#### OpenDocument 格式
- **ODT (OpenDocument Text) ↔ DOCX/Markdown**
- **ODS (OpenDocument Spreadsheet) ↔ XLSX/Markdown**
- **ODP (OpenDocument Presentation) ↔ PPTX**

### 2. 网页与标记语言类

#### HTML
- **HTML ↔ Markdown**: 双向转换，保留基本格式
- **HTML ↔ PDF**: 网页转 PDF
- **HTML ↔ DOCX**: 网页与 Word 互转

#### 其他标记语言
- **Markdown ↔ reStructuredText**: Python 文档常用格式
- **Markdown ↔ AsciiDoc**: 技术文档格式
- **Markdown ↔ Org-mode**: Emacs 组织模式
- **Markdown ↔ Textile**: 轻量级标记语言

#### 电子书格式
- **EPUB ↔ Markdown**: 电子书内容提取
- **EPUB ↔ PDF**: 电子书格式转换
- **MOBI/AZW3 ↔ EPUB/PDF**: Kindle 格式转换

### 3. 数据与表格类

#### CSV/TSV
- **CSV ↔ Markdown**: 表格数据转换
- **CSV ↔ XLSX**: Excel 与 CSV 互转
- **CSV ↔ JSON**: 结构化数据转换

#### JSON/YAML/TOML
- **JSON ↔ YAML**: 配置文件格式互转
- **JSON ↔ TOML**: Rust 生态常用配置格式
- **JSON/XML**: 数据交换格式转换

#### 数据库导出
- **SQL Dump ↔ CSV/JSON**: 数据库导出格式转换
- **SQLite ↔ CSV/JSON**: 轻量数据库导出

### 4. 图像与多媒体类

#### 图像格式
- **PNG ↔ JPG ↔ WEBP**: 常见图像格式转换
- **SVG ↔ PNG/JPG**: 矢量图与位图转换
- **ICO ↔ PNG**: 图标格式转换

#### 图像中的文本 (OCR)
- **Image (PNG/JPG) → Markdown**: OCR 文字识别
- **PDF → Markdown (增强版)**: 使用 OCR 处理扫描版 PDF

#### 图表与可视化
- **Mermaid → SVG/PNG**: 图表渲染
- **Graphviz DOT → SVG/PNG**: 流程图渲染
- **Markdown 表格 → 图表**: 数据可视化

### 5. 代码与文档类

#### API 文档
- **OpenAPI/Swagger ↔ Markdown**: API 文档转换
- **GraphQL Schema ↔ Markdown**: GraphQL 文档生成

#### 代码文档
- **Rust Doc → Markdown**: Rust 文档提取
- **JSDoc → Markdown**: JavaScript 文档提取
- **Docstring → Markdown**: Python 文档提取

#### Notebook 格式
- **Jupyter Notebook (.ipynb) ↔ Markdown**: 交互式笔记本转换
- **Jupyter Notebook ↔ HTML**: 网页展示

### 6. 归档与压缩类

- **ZIP ↔ TAR/GZ**: 压缩格式转换
- **RAR ↔ ZIP**: 压缩格式互转
- **7Z ↔ ZIP**: 压缩格式互转

---

## 🛠️ 技术实现建议

### 推荐依赖库

| 功能 | Crate | 说明 |
|------|-------|------|
| PPTX 处理 | `powerpoint-rs`, `zip` | PowerPoint 读写 |
| ODT/ODS | `odf-rs` | OpenDocument 格式 |
| HTML 处理 | `html5ever`, `scraper` | HTML 解析 |
| EPUB | `epub-builder`, `epub-tools` | EPUB 处理 |
| CSV | `csv` | CSV 读写 |
| JSON | `serde_json` | JSON 序列化 |
| YAML | `serde_yaml` | YAML 处理 |
| TOML | `toml` | TOML 处理 |
| XML | `quick-xml` | XML 解析 |
| 图像处理 | `image` | 图像格式转换 |
| OCR | `tesseract` | 文字识别 |
| Mermaid | 调用外部工具或 WASM | 图表渲染 |

### 架构设计原则

1. **模块化**: 每个转换器独立模块 (`src/converters/*.rs`)
2. **统一接口**: 所有转换器遵循相同的函数签名
   ```rust
   pub fn convert(input: &Path, output: &Path) -> Result<()>;
   ```
3. **错误处理**: 统一的错误类型系统
4. **异步支持**: 可选的异步运行时支持大文件处理
5. **流式处理**: 对大文件使用流式处理减少内存占用

---

## 📋 优先级建议

### 高优先级 (实用性强，实现简单)
1. **HTML ↔ Markdown**: 网页内容处理需求大
2. **CSV ↔ XLSX/Markdown**: 数据处理常用
3. **JSON ↔ YAML/TOML**: 配置文件转换
4. **EPUB ↔ Markdown/PDF**: 电子书处理

### 中优先级 (有一定复杂度)
1. **PPTX ↔ Markdown/PDF**: 演示文稿处理
2. **ODT/ODS ↔ DOCX/XLSX**: 开源办公格式
3. **Jupyter Notebook ↔ Markdown**: 数据科学场景
4. **图像格式转换**: 基础图像处理

### 低优先级 (复杂度高或需求较小)
1. **OCR 功能**: 需要外部依赖，准确率低
2. **MOBI/AZW3**: 专有格式，限制较多
3. **数据库导出**: 场景特定

---

## 🔧 CLI 扩展建议

### 新增命令示例

```bash
# HTML 转换
file_converter html-to-md webpage.html
file_converter md-to-html readme.md

# CSV 转换
file_converter csv-to-md data.csv
file_converter md-to-csv tables.md
file_converter csv-to-xlsx data.csv

# JSON/YAML 转换
file_converter json-to-yaml config.json
file_converter yaml-to-toml config.yaml

# EPUB 转换
file_converter epub-to-md book.epub
file_converter md-to-epub notes.md

# 批量转换
file_converter batch --from dir/input --to dir/output --format markdown
```

### 插件系统设计

可以考虑实现插件系统，允许用户自定义转换器：

```rust
// 插件 trait
pub trait ConverterPlugin {
    fn name(&self) -> &str;
    fn supported_formats(&self) -> Vec<(FileFormat, FileFormat)>;
    fn convert(&self, input: &Path, output: &Path) -> Result<()>;
}
```

---

## 📊 性能优化方向

1. **并行处理**: 使用 `rayon` 进行多文件并行转换
2. **增量转换**: 检测文件变化，只转换变更部分
3. **缓存机制**: 缓存中间结果，加速重复转换
4. **流式处理**: 大文件分块处理，降低内存占用
5. **SIMD 优化**: 对文本处理使用 SIMD 指令集

---

## 🧪 测试策略

1. **单元测试**: 每个转换器的核心逻辑
2. **集成测试**: 端到端转换验证
3. **快照测试**: 确保输出格式稳定性
4. **模糊测试**: 处理异常输入
5. **性能测试**: 基准测试和优化

---

## 📚 学习资源

- [Rust Book](https://doc.rust-lang.org/book/)
- [Awesome Rust](https://github.com/rust-lang/awesome-rust)
- [Crate Documentation](https://docs.rs/)
- [Rust by Example](https://doc.rust-lang.org/rust-by-example/)

---

## 🤝 贡献指南

欢迎贡献新的转换器！请参考以下步骤：

1. 在 `src/converters/` 创建新模块
2. 实现统一的转换接口
3. 添加单元测试
4. 更新 README 和本文档
5. 提交 PR

详见 [CONTRIBUTING.md](CONTRIBUTING.md)
