const mammoth = require('mammoth');
const pdfParse = require('pdf-parse');
const TurndownService = require('turndown');
const fs = require('fs-extra');
const path = require('path');
const { execSync } = require('child_process');
const os = require('os');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');

class DocumentConverter {
    constructor() {
        this.turndownService = new TurndownService({
            headingStyle: 'atx',
            codeBlockStyle: 'fenced'
        });
          this.supportedFormats = {
            '.docx': this.convertDocx.bind(this),
            '.doc': this.convertDoc.bind(this),
            '.pdf': this.convertPdf.bind(this),
            '.html': this.convertHtml.bind(this),
            '.htm': this.convertHtml.bind(this),
            '.rtf': this.convertWithPandoc.bind(this),
            '.odt': this.convertWithPandoc.bind(this),
            '.txt': this.convertText.bind(this),
            '.xlsx': this.convertExcel.bind(this),
            '.xls': this.convertExcel.bind(this),
            '.xlsm': this.convertExcel.bind(this),
            '.xlsb': this.convertExcel.bind(this),
            '.csv': this.convertCsv.bind(this)
        };
          // 检查外部工具可用性
        this.toolsAvailable = {
            pandoc: this.checkPandocAvailable(),
            unstructured: this.checkUnstructuredAvailable(),
            libreoffice: this.checkLibreOfficeAvailable()
        };
    }

    checkPandocAvailable() {
        try {
            execSync('pandoc --version', { stdio: 'ignore' });
            console.log('✓ Pandoc 可用');
            return true;
        } catch (error) {
            console.log('⚠ Pandoc 不可用，某些格式转换可能受限');
            return false;
        }
    }    
    checkLibreOfficeAvailable() {
        try {
            execSync('soffice --version', { stdio: 'ignore' });
            console.log('✓ LibreOffice 可用');
            return true;
        } catch (error) {
            console.log('⚠ LibreOffice 不可用，DOC 文件转换将受限');
            return false;
        }
    }    
    checkUnstructuredAvailable() {
        try {
            // 首先检查 unstructured 是否安装
            execSync('python -c "import unstructured"', { stdio: 'ignore' });
            
            // 然后检查 LibreOffice 是否可用（直接检查，避免循环调用）
            try {
                execSync('soffice --version', { stdio: 'ignore' });
                console.log('✓ Unstructured 和 LibreOffice 都可用');
                return true;
            } catch (libreOfficeError) {
                console.log('⚠ Unstructured 已安装但 LibreOffice 不可用，DOC 文件转换将受限');
                return false;
            }
        } catch (error) {
            console.log('⚠ Unstructured 不可用，将使用其他方法');
            return false;
        }
    }getSupportedFormats() {
        return Object.keys(this.supportedFormats);
    }

    getToolStatus() {
        return {
            pandoc: this.toolsAvailable.pandoc,
            unstructured: this.toolsAvailable.unstructured,
            libreoffice: this.toolsAvailable.libreoffice,
            recommendations: this.getRecommendations()
        };
    }    getRecommendations() {
        const recommendations = [];
        
        if (!this.toolsAvailable.libreoffice) {
            recommendations.push({
                tool: 'LibreOffice',
                reason: '支持 DOC 文件转换（推荐方案）',
                install: 'https://www.libreoffice.org/download/download/'
            });
        }
        
        if (!this.toolsAvailable.pandoc) {
            recommendations.push({
                tool: 'Pandoc',
                reason: '提高 DOC/DOCX/RTF/ODT 文件转换质量',
                install: 'https://pandoc.org/installing.html'
            });
        }
        
        if (!this.toolsAvailable.unstructured) {
            recommendations.push({
                tool: 'Unstructured',
                reason: '提高文档结构识别和转换质量（可选）',
                install: 'pip install unstructured[all-docs]'
            });
        }
        
        return recommendations;
    }

    isSupported(filePath) {
        const ext = path.extname(filePath).toLowerCase();
        return this.supportedFormats.hasOwnProperty(ext);
    }

    async convertFile(inputPath, outputPath) {
        const ext = path.extname(inputPath).toLowerCase();
        
        if (!this.isSupported(inputPath)) {
            throw new Error(`不支持的文件格式: ${ext}`);
        }

        console.log(`开始转换: ${inputPath} -> ${outputPath}`);
        
        try {
            const converter = this.supportedFormats[ext];
            const markdown = await converter(inputPath);
            
            // 确保输出目录存在
            await fs.ensureDir(path.dirname(outputPath));
            
            // 写入Markdown文件
            await fs.writeFile(outputPath, markdown, 'utf8');
            
            console.log(`转换完成: ${outputPath}`);
            return {
                success: true,
                inputPath,
                outputPath,
                size: markdown.length
            };
        } catch (error) {
            console.error(`转换失败 ${inputPath}:`, error.message);
            throw error;
        }
    }    async convertDoc(filePath) {
        console.log(`正在转换 DOC 文件: ${filePath}`);
        
        // 对于DOC文件，优先使用 LibreOffice + mammoth 方案（最可靠）
        if (this.toolsAvailable.libreoffice) {
            try {
                console.log('使用 LibreOffice + mammoth 转换 DOC 文件...');
                return await this.convertDocViaLibreOffice(filePath);
            } catch (error) {
                console.warn(`LibreOffice + mammoth 转换失败: ${error.message}`);
            }
        }
        
        // 备选方案：尝试使用 pandoc
        if (this.toolsAvailable.pandoc) {
            try {
                console.log('尝试使用 Pandoc 转换 DOC 文件...');
                return await this.convertWithPandoc(filePath);
            } catch (error) {
                console.warn(`Pandoc 转换失败: ${error.message}`);
            }
        }
        
        // 如果所有方法都失败，提供详细的错误信息和解决方案
        const suggestions = [
            '1. 将 DOC 文件手动转换为 DOCX 格式',
            '2. 确保 LibreOffice 已正确安装: https://www.libreoffice.org/download/download/',
            '3. 确保 Pandoc 已正确安装: https://pandoc.org/installing.html',
            '4. 使用在线转换工具将文件转换为支持的格式'
        ];
        
        throw new Error(`DOC文件转换失败。所有转换方法都不可用或失败。建议的解决方案：\n${suggestions.join('\n')}`);
    }

    async convertWithPandoc(filePath) {
        const ext = path.extname(filePath).toLowerCase();
        const tempDir = os.tmpdir();
        const tempMdFile = path.join(tempDir, `temp_${Date.now()}.md`);
        
        try {
            // 构建 pandoc 命令
            let inputFormat = '';
            switch (ext) {
                case '.doc':
                case '.docx':
                    inputFormat = 'docx';
                    break;
                case '.rtf':
                    inputFormat = 'rtf';
                    break;
                case '.odt':
                    inputFormat = 'odt';
                    break;
                case '.html':
                case '.htm':
                    inputFormat = 'html';
                    break;
                default:
                    inputFormat = 'docx'; // 默认尝试 docx
            }
            
            const command = `pandoc -f ${inputFormat} -t markdown "${filePath}" -o "${tempMdFile}"`;
            console.log(`执行命令: ${command}`);
            
            execSync(command, { 
                stdio: 'pipe',
                maxBuffer: 1024 * 1024 * 10 // 10MB buffer
            });
            
            // 读取生成的 Markdown 文件
            let markdown = await fs.readFile(tempMdFile, 'utf8');
            
            // 清理临时文件
            await fs.remove(tempMdFile);
            
            // 添加文档信息
            const fileName = path.basename(filePath);
            markdown = `# ${fileName}\n\n${markdown}`;
            
            console.log(`Pandoc 转换成功: ${filePath}`);
            return markdown;
            
        } catch (error) {
            // 清理临时文件
            try {
                await fs.remove(tempMdFile);
            } catch (cleanupError) {
                // 忽略清理错误
            }
            throw new Error(`Pandoc 转换失败: ${error.message}`);
        }
    }    async convertWithUnstructured(filePath) {
        const tempDir = os.tmpdir();
        const tempJsonFile = path.join(tempDir, `temp_${Date.now()}.json`);
        const tempPyFile = path.join(tempDir, `temp_${Date.now()}.py`);
        
        try {
            // 使用 unstructured 库提取文档内容
            const pythonScript = `
import json
import sys
from unstructured.partition.auto import partition

try:
    elements = partition("${filePath.replace(/\\/g, '\\\\')}")
    data = []
    for element in elements:
        data.append({
            "type": str(type(element).__name__),
            "text": element.text,
            "metadata": element.metadata.__dict__ if hasattr(element, 'metadata') else {}
        })

    with open("${tempJsonFile.replace(/\\/g, '\\\\')}", 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        
except Exception as e:
    # 将错误信息写入文件，以便 Node.js 可以读取
    error_info = {
        "error": str(e),
        "type": type(e).__name__
    }
    with open("${tempJsonFile.replace(/\\/g, '\\\\')}", 'w', encoding='utf-8') as f:
        json.dump({"error": error_info}, f, ensure_ascii=False, indent=2)
    sys.exit(1)
`;
            
            await fs.writeFile(tempPyFile, pythonScript, 'utf8');
            
            execSync(`python "${tempPyFile}"`, { 
                stdio: 'pipe',
                maxBuffer: 1024 * 1024 * 10
            });
            
            // 读取解析结果
            const result = JSON.parse(await fs.readFile(tempJsonFile, 'utf8'));
            
            // 检查是否有错误
            if (result.error) {
                throw new Error(`${result.error.type}: ${result.error.error}`);
            }
            
            const data = result;
            
            // 转换为 Markdown
            let markdown = '';
            const fileName = path.basename(filePath);
            markdown += `# ${fileName}\n\n`;
            
            for (const element of data) {
                const text = element.text.trim();
                if (!text) continue;
                
                // 根据元素类型格式化
                switch (element.type) {
                    case 'Title':
                        markdown += `## ${text}\n\n`;
                        break;
                    case 'Header':
                        markdown += `### ${text}\n\n`;
                        break;
                    case 'ListItem':
                        markdown += `- ${text}\n`;
                        break;
                    case 'Table':
                        markdown += `${text}\n\n`;
                        break;
                    default:
                        markdown += `${text}\n\n`;
                }
            }
            
            // 清理临时文件
            await fs.remove(tempPyFile);
            await fs.remove(tempJsonFile);
            
            console.log(`Unstructured 转换成功: ${filePath}`);
            return markdown;
            
        } catch (error) {
            // 清理临时文件
            try {
                await fs.remove(tempPyFile);
                await fs.remove(tempJsonFile);
            } catch (cleanupError) {
                // 忽略清理错误
            }
            
            // 检查是否是 LibreOffice 相关错误
            if (error.message.includes('soffice command was not found') || 
                error.message.includes('libreoffice') ||
                error.message.includes('FileNotFoundError')) {
                throw new Error(`soffice command was not found. Please install libreoffice`);
            }
            
            throw new Error(`Unstructured 转换失败: ${error.message}`);
        }
    }

    async convertText(filePath) {
        try {
            const text = await fs.readFile(filePath, 'utf8');
            const fileName = path.basename(filePath);
            
            // 简单的文本格式化
            let markdown = `# ${fileName}\n\n`;
            
            // 按段落分割
            const paragraphs = text.split(/\n\s*\n/);
            for (const paragraph of paragraphs) {
                const cleanParagraph = paragraph.trim().replace(/\n/g, ' ');
                if (cleanParagraph) {
                    markdown += `${cleanParagraph}\n\n`;
                }
            }
            
            return markdown;
        } catch (error) {
            throw new Error(`文本文件转换失败: ${error.message}`);
        }
    }

    async convertDocx(filePath) {
        try {
            const buffer = await fs.readFile(filePath);
            const result = await mammoth.convertToHtml({
                buffer: buffer,
                options: {
                    convertImage: mammoth.images.imgElement(function(image) {
                        return image.read("base64").then(function(imageBuffer) {
                            return {
                                src: "data:" + image.contentType + ";base64," + imageBuffer
                            };
                        });
                    })
                }
            });
            
            // 将HTML转换为Markdown
            let markdown = this.turndownService.turndown(result.value);
            
            // 添加文档信息
            const fileName = path.basename(filePath);
            markdown = `# ${fileName}\n\n${markdown}`;
            
            return markdown;
        } catch (error) {
            throw new Error(`DOCX转换失败: ${error.message}`);
        }
    }

    async convertPdf(filePath) {
        try {
            const buffer = await fs.readFile(filePath);
            const data = await pdfParse(buffer);
            
            // 清理文本并格式化为Markdown
            let text = data.text;
            
            // 基本的文本清理
            text = text.replace(/\r\n/g, '\n')
                      .replace(/\r/g, '\n')
                      .replace(/\n{3,}/g, '\n\n')
                      .trim();
            
            // 添加文档信息
            const fileName = path.basename(filePath);
            let markdown = `# ${fileName}\n\n`;
            
            // 尝试识别标题和段落
            const lines = text.split('\n');
            let inParagraph = false;
            
            for (let line of lines) {
                line = line.trim();
                if (!line) {
                    if (inParagraph) {
                        markdown += '\n\n';
                        inParagraph = false;
                    }
                    continue;
                }
                
                // 简单的标题检测（全大写或首字母大写且较短）
                if (line.length < 80 && (line === line.toUpperCase() || /^[A-Z][^.]*$/.test(line))) {
                    if (inParagraph) {
                        markdown += '\n\n';
                    }
                    markdown += `## ${line}\n\n`;
                    inParagraph = false;
                } else {
                    if (!inParagraph) {
                        inParagraph = true;
                    } else {
                        markdown += ' ';
                    }
                    markdown += line;
                }
            }
            
            return markdown;
        } catch (error) {
            throw new Error(`PDF转换失败: ${error.message}`);
        }
    }

    async convertHtml(filePath) {
        try {
            const html = await fs.readFile(filePath, 'utf8');
            console.log(`HTML文件大小: ${html.length} 字符`);
            
            let markdown = this.turndownService.turndown(html);
            console.log(`转换后Markdown大小: ${markdown.length} 字符`);
            
            // 添加文档信息
            const fileName = path.basename(filePath);
            markdown = `# ${fileName}\n\n${markdown}`;
            
            return markdown;
        } catch (error) {
            throw new Error(`HTML转换失败: ${error.message}`);
        }
    }

    generateOutputPath(inputPath, sourceDir, targetDir) {
        const relativePath = path.relative(sourceDir, inputPath);
        const parsedPath = path.parse(relativePath);
        const outputFileName = parsedPath.name + '.md';
        return path.join(targetDir, parsedPath.dir, outputFileName);
    }

    async convertDocViaLibreOffice(filePath) {
        const tempDir = os.tmpdir();
        const tempDocxFile = path.join(tempDir, `temp_${Date.now()}.docx`);
        
        try {
            console.log('使用 LibreOffice 将 DOC 转换为临时 DOCX 文件...');
            
            // 使用 LibreOffice 转换 DOC 到 DOCX
            const command = `soffice --headless --convert-to docx --outdir "${tempDir}" "${filePath}"`;
            console.log(`执行命令: ${command}`);
            
            execSync(command, { 
                stdio: 'pipe',
                maxBuffer: 1024 * 1024 * 10,
                timeout: 30000 // 30秒超时
            });
            
            // 构建预期的输出文件路径
            const originalName = path.basename(filePath, '.doc');
            const actualTempFile = path.join(tempDir, `${originalName}.docx`);
            
            // 检查文件是否生成
            if (!await fs.pathExists(actualTempFile)) {
                throw new Error('LibreOffice 转换未生成预期的 DOCX 文件');
            }
            
            // 使用 mammoth 处理生成的 DOCX 文件
            console.log('使用 mammoth 处理转换后的 DOCX 文件...');
            const markdown = await this.convertDocx(actualTempFile);
            
            // 清理临时文件
            await fs.remove(actualTempFile);
            
            console.log(`LibreOffice + mammoth 转换成功: ${filePath}`);
            return markdown;
            
        } catch (error) {
            // 清理临时文件
            try {
                const originalName = path.basename(filePath, '.doc');
                const actualTempFile = path.join(tempDir, `${originalName}.docx`);
                await fs.remove(actualTempFile);
            } catch (cleanupError) {
                // 忽略清理错误
            }
              throw new Error(`LibreOffice 转换失败: ${error.message}`);
        }
    }

    async convertExcel(filePath) {
        try {
            console.log(`正在转换 Excel 文件: ${filePath}`);
            
            const workbook = XLSX.readFile(filePath);
            const fileName = path.basename(filePath);
            let markdown = `# ${fileName}\n\n`;
            
            // 遍历所有工作表
            const sheetNames = workbook.SheetNames;
            
            for (let i = 0; i < sheetNames.length; i++) {
                const sheetName = sheetNames[i];
                const worksheet = workbook.Sheets[sheetName];
                
                // 如果有多个工作表，添加工作表标题
                if (sheetNames.length > 1) {
                    markdown += `## 工作表: ${sheetName}\n\n`;
                }
                
                // 将工作表转换为JSON数组
                const jsonData = XLSX.utils.sheet_to_json(worksheet, {
                    header: 1, // 使用数组而不是对象
                    defval: '', // 空单元格的默认值
                    blankrows: false // 跳过空行
                });
                
                if (jsonData.length === 0) {
                    markdown += '*该工作表为空*\n\n';
                    continue;
                }
                
                // 检查是否有数据
                const hasData = jsonData.some(row => row.some(cell => cell !== ''));
                if (!hasData) {
                    markdown += '*该工作表为空*\n\n';
                    continue;
                }
                
                // 找到最大列数
                const maxCols = Math.max(...jsonData.map(row => row.length));
                
                // 如果第一行看起来像标题行，使用表格格式
                if (jsonData.length > 1) {
                    const firstRow = jsonData[0];
                    const hasHeaders = firstRow.every(cell => 
                        typeof cell === 'string' && cell.trim() !== ''
                    );
                    
                    if (hasHeaders) {
                        // 表格格式
                        markdown += this.formatAsTable(jsonData, maxCols);
                    } else {
                        // 简单列表格式
                        markdown += this.formatAsSimpleList(jsonData);
                    }
                } else {
                    // 单行数据
                    markdown += this.formatAsSimpleList(jsonData);
                }
                
                markdown += '\n';
            }
            
            console.log(`Excel 转换成功: ${filePath}`);
            return markdown;
            
        } catch (error) {
            throw new Error(`Excel转换失败: ${error.message}`);
        }
    }

    async convertCsv(filePath) {
        try {
            console.log(`正在转换 CSV 文件: ${filePath}`);
            
            const csvContent = await fs.readFile(filePath, 'utf8');
            const fileName = path.basename(filePath);
            let markdown = `# ${fileName}\n\n`;
            
            // 简单的CSV解析（处理基本情况）
            const lines = csvContent.split(/\r?\n/).filter(line => line.trim());
            
            if (lines.length === 0) {
                markdown += '*CSV文件为空*\n\n';
                return markdown;
            }
            
            // 解析CSV行
            const rows = lines.map(line => this.parseCSVLine(line));
            
            // 找到最大列数
            const maxCols = Math.max(...rows.map(row => row.length));
            
            // 格式化为表格
            if (rows.length > 1) {
                markdown += this.formatAsTable(rows, maxCols);
            } else {
                markdown += this.formatAsSimpleList(rows);
            }
            
            console.log(`CSV 转换成功: ${filePath}`);
            return markdown;
            
        } catch (error) {
            throw new Error(`CSV转换失败: ${error.message}`);
        }
    }

    // 辅助方法：格式化为Markdown表格
    formatAsTable(data, maxCols) {
        if (data.length === 0) return '';
        
        let markdown = '';
        
        // 表头
        const headers = data[0];
        markdown += '| ';
        for (let i = 0; i < maxCols; i++) {
            const header = headers[i] || `列${i + 1}`;
            markdown += `${this.escapeMarkdown(String(header))} | `;
        }
        markdown += '\n';
        
        // 分隔符行
        markdown += '|';
        for (let i = 0; i < maxCols; i++) {
            markdown += ' --- |';
        }
        markdown += '\n';
        
        // 数据行
        for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
            const row = data[rowIndex];
            markdown += '| ';
            for (let i = 0; i < maxCols; i++) {
                const cell = row[i] || '';
                markdown += `${this.escapeMarkdown(String(cell))} | `;
            }
            markdown += '\n';
        }
        
        return markdown;
    }

    // 辅助方法：格式化为简单列表
    formatAsSimpleList(data) {
        let markdown = '';
        
        data.forEach((row, rowIndex) => {
            if (row.length === 0) return;
            
            // 如果行只有一个有效值，直接显示
            const validCells = row.filter(cell => cell !== '');
            if (validCells.length === 1) {
                markdown += `**第${rowIndex + 1}行**: ${this.escapeMarkdown(String(validCells[0]))}\n\n`;
            } else {
                // 多个值，使用列表格式
                markdown += `**第${rowIndex + 1}行**:\n`;
                row.forEach((cell, cellIndex) => {
                    if (cell !== '') {
                        markdown += `- 列${cellIndex + 1}: ${this.escapeMarkdown(String(cell))}\n`;
                    }
                });
                markdown += '\n';
            }
        });
        
        return markdown;
    }

    // 辅助方法：简单的CSV行解析
    parseCSVLine(line) {
        const result = [];
        let current = '';
        let inQuotes = false;
        
        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            const nextChar = line[i + 1];
            
            if (char === '"') {
                if (inQuotes && nextChar === '"') {
                    // 转义的引号
                    current += '"';
                    i++; // 跳过下一个引号
                } else {
                    // 切换引号状态
                    inQuotes = !inQuotes;
                }
            } else if (char === ',' && !inQuotes) {
                // 字段分隔符
                result.push(current.trim());
                current = '';
            } else {
                current += char;
            }
        }
        
        // 添加最后一个字段
        result.push(current.trim());
        
        return result;
    }

    // 辅助方法：转义Markdown特殊字符
    escapeMarkdown(text) {
        if (!text) return '';
        return text.replace(/[|\\`*_{}[\]()#+\-.!]/g, '\\$&');
    }
}

module.exports = DocumentConverter;
