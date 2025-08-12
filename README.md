# 文档转换器 - Document to Markdown Converter
- Author: LiXizhi
- Date: 2025.8.12

一个强大的文档转换工具，支持将DOC、DOCX、PDF、HTML、Excel等格式的文档自动转换为Markdown格式，并提供类似rsync的增量同步功能。用户批量同步RAG数据源。内部转化工具支持pandoc, unstructured, mammoth, xlsx。 

将./data 目录下的所有文件（含子目录）转为./data_markdown 目录下。 只转变化的文件，./data/.filehashes 保存文件hash值，若需要可以清除这个文件。 

## ✨ 主要功能

- 📄 **多格式支持**: 支持DOC、DOCX、PDF、HTML、Excel等常见文档格式
- 🔄 **增量同步**: 类似rsync的功能，只处理有变化的文件
- 👀 **实时监控**: 自动监控源目录的文件变化
- 🌐 **Web界面**: 美观的前端界面，实时显示转换进度
- ⚡ **实时通信**: WebSocket实时推送转换状态
- 📊 **详细日志**: 完整的转换日志和错误信息

## 🚀 快速开始

### 1. 启动服务器

```bash
npm start
```

### 2. 开发模式（自动重启）

```bash
npm run dev
```

### 3. 访问Web界面

打开浏览器访问：http://localhost:3000

## 📁 项目结构

```
├── server.js              # 主服务器文件
├── lib/
│   ├── documentConverter.js  # 文档转换器
│   └── fileWatcher.js        # 文件监控器
├── public/
│   ├── index.html           # 前端页面
│   ├── styles.css           # 样式文件
│   └── app.js              # 前端JavaScript
├── data/                   # 源文档目录
└── data_markdown/          # 转换后的Markdown目录
```

## 🎯 使用方法

### Web界面操作

1. **设置源目录**: 在输入框中输入要监控的文件夹路径
2. **转换文件**: 
   - 点击"转换所有文件"进行增量转换
   - 点击"强制转换所有"重新转换所有文件
   - 单击文件旁的"转换"按钮转换单个文件
3. **实时监控**: 系统会自动监控源目录，新增或修改的文件会自动转换

### API接口

#### 获取系统状态
```http
GET /api/status
```

#### 获取文件列表
```http
GET /api/files
```

#### 设置源目录
```http
POST /api/set-source-dir
Content-Type: application/json

{
  "directory": "C:\\Documents"
}
```

#### 转换文档
```http
POST /api/convert
Content-Type: application/json

{
  "filePath": "C:\\Documents\\example.docx",  // 可选，转换单个文件
  "force": false  // 是否强制转换
}
```

## 🔧 支持的文件格式

| 格式 | 扩展名 | 说明 |
|------|--------|------|
| Word文档 | .docx, .doc | 支持文本、表格、图片转换 |
| PDF文档 | .pdf | 提取文本内容并格式化 |
| HTML文档 | .html, .htm | 转换为标准Markdown |
| Excel表格 | .xlsx, .xls, .xlsm, .xlsb | 支持多工作表，自动识别表格结构 |
| CSV文件 | .csv | 转换为Markdown表格格式 |

## ⚙️ 配置说明

### 环境变量

- `PORT`: 服务器端口，默认3000

### 转换规则

1. **文件名处理**: 保持原文件名，扩展名改为`.md`
2. **目录结构**: 保持源目录的文件夹结构
3. **增量检测**: 基于文件修改时间和MD5哈希
4. **图片处理**: DOCX中的图片转换为base64内嵌

## 🐛 故障排除

### 常见问题

1. **转换失败**
   - 检查文件是否已被其他程序占用
   - 确认文件格式是否支持
   - 查看详细错误信息

2. **监控不工作**
   - 检查源目录路径是否正确
   - 确认目录访问权限
   - 重启服务器

3. **WebSocket连接失败**
   - 检查防火墙设置
   - 确认端口未被占用

## 📝 开发说明

### 技术栈

- **后端**: Node.js + Express
- **前端**: 原生JavaScript + WebSocket
- **文档转换**: mammoth (DOCX) + pdf-parse (PDF) + turndown (HTML→MD) + xlsx (Excel)
- **文件监控**: chokidar

### 扩展功能

要添加新的文档格式支持，需要：

1. 在`DocumentConverter`类中添加转换方法
2. 在`supportedFormats`对象中注册格式
3. 安装相应的解析库

## 📄 许可证

MIT License

## 🤝 贡献

欢迎提交问题和功能请求！

---

🚀 **享受高效的文档转换体验！**
