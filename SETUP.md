# 文档转换器设置指南

此文档转换器支持多种文件格式转换为 Markdown。为了获得最佳转换效果，建议安装以下工具：

## 支持的文件格式

### 基本支持（无需额外工具）
- **HTML/HTM**: 使用内置的 turndown 库
- **DOCX**: 使用 mammoth 库
- **PDF**: 使用 pdf-parse 库（文本提取）
- **TXT**: 基本文本处理

### 增强支持（需要外部工具）
- **DOC**: 推荐使用 Pandoc 或 Unstructured
- **RTF**: 需要 Pandoc
- **ODT**: 需要 Pandoc

## 安装指南

### 1. 安装 Pandoc（推荐）

Pandoc 是一个通用的文档转换工具，支持多种格式。

#### Windows
1. 访问 [Pandoc 官网](https://pandoc.org/installing.html) 
   - 或者运行`winget install --id JohnMacFarlane.Pandoc -e`
2. 下载 Windows 安装包
3. 运行安装程序
4. 验证安装：在命令行中运行 `pandoc --version`

#### macOS
```bash
brew install pandoc
```

#### Linux (Ubuntu/Debian)
```bash
sudo apt-get install pandoc
```

### 2. 安装 Python Unstructured 库（可选）

Unstructured 是一个 Python 库，专门用于处理非结构化文档。

```bash
# 安装 Python（如果尚未安装）
# 然后安装 unstructured
pip install unstructured[all-docs]
```

注意：unstructured 库需要额外的依赖项，安装可能需要一些时间。

如果需要DOC文件，需要下载：
https://www.libreoffice.org/


### 3. 验证安装

运行应用程序后，控制台会显示可用工具的状态：
- ✓ Pandoc 可用
- ✓ Unstructured 可用

## 转换策略

### DOC 文件转换优先级
1. **LibreOffice + Mammoth**（推荐，兼容性最佳）
2. **Pandoc**（备选方案）

### 推荐配置
- **最小配置**: 仅使用内置库（支持 DOCX, PDF, HTML, TXT）
- **推荐配置**: 安装 LibreOffice（支持所有常见格式，包括DOC）
- **完整配置**: 安装 LibreOffice + Pandoc（最佳转换质量和格式支持）

## 故障排除

### DOC 文件转换失败
如果 DOC 文件转换失败，建议：
1. 将 DOC 文件另存为 DOCX 格式
2. 检查 Pandoc 是否正确安装
3. 确认文件没有损坏

### PDF 转换质量不佳
PDF 转换基于文本提取，对于以下情况效果有限：
- 扫描版 PDF（图片格式）
- 复杂布局的 PDF
- 包含大量图表的 PDF

建议使用专门的 OCR 工具处理扫描版 PDF。

## 性能说明

- 大文件转换可能需要较长时间
- Pandoc 转换通常比 Mammoth 慢但质量更好
- Unstructured 对复杂文档的结构识别更准确

## 开发说明

如需添加新的文件格式支持，可以：
1. 在 `supportedFormats` 中添加新的扩展名
2. 实现对应的转换方法
3. 添加必要的依赖项

示例：
```javascript
this.supportedFormats['.新格式'] = this.convertNewFormat.bind(this);
```
