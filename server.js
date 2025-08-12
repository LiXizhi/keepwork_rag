const express = require('express');
const fs = require('fs-extra');
const path = require('path');
const chokidar = require('chokidar');
const WebSocket = require('ws');
const http = require('http');
const cors = require('cors');
const DocumentConverter = require('./lib/documentConverter');
const FileWatcher = require('./lib/fileWatcher');

class DocumentConverterServer {
    constructor() {
        this.app = express();
        this.server = http.createServer(this.app);
        this.wss = new WebSocket.Server({ server: this.server });
        this.port = process.env.PORT || 3000;
        
        this.sourceDir = path.join(__dirname, 'data');
        this.targetDir = path.join(__dirname, 'data_markdown');
        
        this.converter = new DocumentConverter();
        this.fileWatcher = new FileWatcher(this.sourceDir, this.targetDir, this.converter);
        
        this.setupMiddleware();
        this.setupRoutes();
        this.setupWebSocket();
        this.setupFileWatcher();
    }

    setupMiddleware() {
        this.app.use(cors());
        this.app.use(express.json());
        this.app.use(express.static(path.join(__dirname, 'public')));
        
        // 确保目标目录存在
        fs.ensureDirSync(this.targetDir);
    }

    setupRoutes() {
        // 首页
        this.app.get('/', (req, res) => {
            res.sendFile(path.join(__dirname, 'public', 'index.html'));
        });

        // 获取文件列表
        this.app.get('/api/files', async (req, res) => {
            try {
                const sourceFiles = await this.getFileList(this.sourceDir);
                const markdownFiles = await this.getFileList(this.targetDir);
                
                res.json({
                    sourceFiles,
                    markdownFiles,
                    status: 'success'
                });
            } catch (error) {
                res.status(500).json({
                    error: error.message,
                    status: 'error'
                });
            }
        });

        // 手动触发转换
        this.app.post('/api/convert', async (req, res) => {
            try {
                const { filePath, force = false } = req.body;
                
                if (filePath) {
                    // 转换单个文件
                    const result = await this.fileWatcher.processFile(filePath, force);
                    res.json({ result, status: 'success' });
                } else {
                    // 转换所有文件
                    const results = await this.fileWatcher.processAllFiles(force);
                    res.json({ results, status: 'success' });
                }
            } catch (error) {
                res.status(500).json({
                    error: error.message,
                    status: 'error'
                });
            }
        });

        // 获取转换状态
        this.app.get('/api/status', (req, res) => {
            res.json({
                watching: this.fileWatcher.isWatching,
                sourceDir: this.sourceDir,
                targetDir: this.targetDir,
                supportedFormats: this.converter.getSupportedFormats(),
                status: 'success'
            });
        });

        // 设置源目录
        this.app.post('/api/set-source-dir', async (req, res) => {
            try {
                const { directory } = req.body;
                if (!directory) {
                    return res.status(400).json({
                        error: 'Directory path is required',
                        status: 'error'
                    });
                }

                const fullPath = path.resolve(directory);
                const exists = await fs.pathExists(fullPath);
                
                if (!exists) {
                    return res.status(400).json({
                        error: 'Directory does not exist',
                        status: 'error'
                    });
                }

                this.sourceDir = fullPath;
                this.fileWatcher.updateSourceDir(fullPath);
                
                res.json({
                    message: 'Source directory updated',
                    sourceDir: this.sourceDir,
                    status: 'success'
                });
            } catch (error) {
                res.status(500).json({
                    error: error.message,
                    status: 'error'
                });
            }
        });
    }

    setupWebSocket() {
        this.wss.on('connection', (ws) => {
            console.log('客户端连接已建立');
            
            // 发送初始状态
            ws.send(JSON.stringify({
                type: 'status',
                data: {
                    watching: this.fileWatcher.isWatching,
                    sourceDir: this.sourceDir,
                    targetDir: this.targetDir
                }
            }));

            ws.on('close', () => {
                console.log('客户端连接已断开');
            });
        });

        // 设置文件观察器事件监听
        this.fileWatcher.on('fileProcessed', (data) => {
            this.broadcast({
                type: 'fileProcessed',
                data
            });
        });

        this.fileWatcher.on('error', (error) => {
            this.broadcast({
                type: 'error',
                data: { message: error.message }
            });
        });
    }

    setupFileWatcher() {
        this.fileWatcher.start();
    }

    broadcast(message) {
        this.wss.clients.forEach((client) => {
            if (client.readyState === WebSocket.OPEN) {
                client.send(JSON.stringify(message));
            }
        });
    }

    async getFileList(directory) {
        const files = [];
        
        if (!(await fs.pathExists(directory))) {
            return files;
        }

        const items = await fs.readdir(directory, { withFileTypes: true });
        
        for (const item of items) {
            if (item.isFile()) {
                const filePath = path.join(directory, item.name);
                const stats = await fs.stat(filePath);
                
                files.push({
                    name: item.name,
                    path: filePath,
                    size: stats.size,
                    mtime: stats.mtime,
                    ext: path.extname(item.name).toLowerCase()
                });
            }
        }

        return files;
    }

    start() {
        this.server.listen(this.port, () => {
            console.log(`文档转换服务器运行在 http://localhost:${this.port}`);
            console.log(`监控目录: ${this.sourceDir}`);
            console.log(`输出目录: ${this.targetDir}`);
        });
    }

    stop() {
        this.fileWatcher.stop();
        this.server.close();
    }
}

// 启动服务器
const server = new DocumentConverterServer();
server.start();

// 优雅关闭
process.on('SIGINT', () => {
    console.log('\n正在关闭服务器...');
    server.stop();
    process.exit(0);
});

module.exports = DocumentConverterServer;
