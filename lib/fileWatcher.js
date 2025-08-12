const chokidar = require('chokidar');
const fs = require('fs-extra');
const path = require('path');
const { EventEmitter } = require('events');

class FileWatcher extends EventEmitter {
    constructor(sourceDir, targetDir, converter) {
        super();
        this.sourceDir = sourceDir;
        this.targetDir = targetDir;
        this.converter = converter;
        this.watcher = null;
        this.isWatching = false;
        this.fileHashes = new Map(); // 存储文件哈希用于增量检测
        this.processing = new Set(); // 避免重复处理
        this.hashFilePath = path.join(process.cwd(), 'data', '.filehashes'); // 哈希文件存储路径
        this.loadFileHashes(); // 启动时加载已保存的哈希
    }

    // 加载保存的文件哈希
    async loadFileHashes() {
        try {
            // 确保数据目录存在
            await fs.ensureDir(path.dirname(this.hashFilePath));
            
            if (await fs.pathExists(this.hashFilePath)) {
                const hashData = await fs.readFile(this.hashFilePath, 'utf8');
                const hashObject = JSON.parse(hashData);
                
                // 将对象转换为 Map
                this.fileHashes = new Map(Object.entries(hashObject));
                console.log(`已加载 ${this.fileHashes.size} 个文件哈希记录`);
            } else {
                console.log('未找到哈希文件，使用空的哈希缓存');
            }
        } catch (error) {
            console.warn('加载文件哈希失败，使用空的哈希缓存:', error.message);
            this.fileHashes = new Map();
        }
    }
    
    // 保存文件哈希到磁盘
    async saveFileHashes() {
        try {
            // 确保数据目录存在
            await fs.ensureDir(path.dirname(this.hashFilePath));
            
            // 将 Map 转换为对象
            const hashObject = Object.fromEntries(this.fileHashes);
            
            // 写入文件
            await fs.writeFile(this.hashFilePath, JSON.stringify(hashObject, null, 2), 'utf8');
            console.log(`已保存 ${this.fileHashes.size} 个文件哈希记录`);
        } catch (error) {
            console.error('保存文件哈希失败:', error.message);
        }
    }

    updateSourceDir(newSourceDir) {
        this.stop();
        this.sourceDir = newSourceDir;
        this.start();
    }

    start() {
        if (this.isWatching) {
            return;
        }

        console.log(`开始监控目录: ${this.sourceDir}`);
        
        this.watcher = chokidar.watch(this.sourceDir, {
            ignored: /(^|[\/\\])\../, // 忽略隐藏文件
            persistent: true,
            ignoreInitial: false,
            followSymlinks: false,
            depth: 10
        });

        this.watcher
            .on('add', (filePath) => this.handleFileChange(filePath, 'added'))
            .on('change', (filePath) => this.handleFileChange(filePath, 'changed'))
            .on('unlink', (filePath) => this.handleFileDelete(filePath))
            .on('error', (error) => {
                console.error('文件监控错误:', error);
                this.emit('error', error);
            })
            .on('ready', () => {
                console.log('文件监控已就绪');
                this.isWatching = true;
            });
    }    stop() {
        if (this.watcher) {
            this.watcher.close();
            this.watcher = null;
        }
        this.isWatching = false;
        // 保存哈希到磁盘
        this.saveFileHashes().catch(error => {
            console.error('停止时保存文件哈希失败:', error.message);
        });
        console.log('文件监控已停止');
    }

    async handleFileChange(filePath, eventType) {
        if (!this.converter.isSupported(filePath)) {
            return;
        }

        if (this.processing.has(filePath)) {
            return;
        }

        try {
            this.processing.add(filePath);
            
            // 检查文件是否需要处理（基于修改时间和哈希）
            const needsProcessing = await this.needsProcessing(filePath);
            
            if (needsProcessing) {
                await this.processFile(filePath);
            }
        } catch (error) {
            console.error(`处理文件失败 ${filePath}:`, error.message);
            this.emit('error', error);
        } finally {
            this.processing.delete(filePath);
        }
    }

    async handleFileDelete(filePath) {
        if (!this.converter.isSupported(filePath)) {
            return;
        }

        try {
            const outputPath = this.converter.generateOutputPath(filePath, this.sourceDir, this.targetDir);
            
            if (await fs.pathExists(outputPath)) {
                await fs.unlink(outputPath);
                console.log(`删除对应的Markdown文件: ${outputPath}`);
                
                this.emit('fileProcessed', {
                    inputPath: filePath,
                    outputPath,
                    action: 'deleted',
                    success: true
                });
            }
              // 从哈希缓存中移除
            this.fileHashes.delete(filePath);
            // 保存哈希到磁盘
            await this.saveFileHashes();
        } catch (error) {
            console.error(`删除文件失败 ${filePath}:`, error.message);
            this.emit('error', error);
        }
    }

    async needsProcessing(filePath, force = false) {
        if (force) {
            return true;
        }

        try {
            const stats = await fs.stat(filePath);
            const outputPath = this.converter.generateOutputPath(filePath, this.sourceDir, this.targetDir);
            
            // 检查输出文件是否存在
            if (!(await fs.pathExists(outputPath))) {
                return true;
            }
            
            // 检查输出文件的修改时间
            const outputStats = await fs.stat(outputPath);
            if (stats.mtime > outputStats.mtime) {
                return true;
            }
              // 检查文件哈希（可选，用于更精确的检测）
            const currentHash = await this.getFileHash(filePath);
            const cachedHash = this.fileHashes.get(filePath);
            
            if (currentHash !== cachedHash) {
                // 注意：这里只检测变化，不更新缓存，缓存会在成功处理后更新
                return true;
            }
            
            return false;
        } catch (error) {
            // 出错时默认需要处理
            return true;
        }
    }

    async getFileHash(filePath) {
        try {
            const crypto = require('crypto');
            const buffer = await fs.readFile(filePath);
            return crypto.createHash('md5').update(buffer).digest('hex');
        } catch (error) {
            // 如果无法计算哈希，使用修改时间作为替代
            const stats = await fs.stat(filePath);
            return stats.mtime.getTime().toString();
        }
    }

    async processFile(filePath, force = false) {
        if (!force && !await this.needsProcessing(filePath, force)) {
            return {
                skipped: true,
                reason: 'No changes detected'
            };
        }

        try {
            const outputPath = this.converter.generateOutputPath(filePath, this.sourceDir, this.targetDir);
            const result = await this.converter.convertFile(filePath, outputPath);
              // 更新哈希缓存
            const hash = await this.getFileHash(filePath);
            this.fileHashes.set(filePath, hash);
            // 保存哈希到磁盘
            await this.saveFileHashes();
            
            this.emit('fileProcessed', {
                ...result,
                action: 'converted'
            });
            
            return result;
        } catch (error) {
            const errorResult = {
                success: false,
                inputPath: filePath,
                error: error.message,
                action: 'failed'
            };
            
            this.emit('fileProcessed', errorResult);
            throw error;
        }
    }

    async processAllFiles(force = false) {
        console.log('开始处理所有支持的文件...');
        const results = [];
        
        try {
            const files = await this.getAllSupportedFiles();
            
            for (const filePath of files) {
                try {
                    const result = await this.processFile(filePath, force);
                    results.push(result);
                } catch (error) {
                    results.push({
                        success: false,
                        inputPath: filePath,
                        error: error.message
                    });
                }
            }
            
            console.log(`批量处理完成，共处理 ${results.length} 个文件`);
            return results;
        } catch (error) {
            console.error('批量处理失败:', error);
            throw error;
        }
    }

    async getAllSupportedFiles() {
        const files = [];
        
        if (!(await fs.pathExists(this.sourceDir))) {
            return files;
        }

        const walk = async (dir) => {
            const items = await fs.readdir(dir, { withFileTypes: true });
            
            for (const item of items) {
                const fullPath = path.join(dir, item.name);
                
                if (item.isDirectory()) {
                    await walk(fullPath);
                } else if (item.isFile() && this.converter.isSupported(fullPath)) {
                    files.push(fullPath);
                }
            }
        };

        await walk(this.sourceDir);
        return files;
    }
}

module.exports = FileWatcher;
