import * as fs from 'fs';
import * as path from 'path';
import { ConnectionManager } from '../connection/ConnectionManager';
import { FileComparer } from '../compare/FileComparer';
import { NotificationManager } from '../notification/NotificationManager';
import { SyncDetails, VBAModule } from '../../types';

// 调试日志开关
const DEBUG = true;

// 统一日志函数
function debugLog(message: string, ...args: any[]) {
    if (DEBUG) {
        console.log('[VBA SYNC DEBUG] ' + message, ...args);
    }
}

export class IdeToDocumentSync {
    private connectionManager: ConnectionManager;
    private notificationManager: NotificationManager;

    constructor(connectionManager: ConnectionManager, notificationManager: NotificationManager) {
        this.connectionManager = connectionManager;
        this.notificationManager = notificationManager;
    }

    async syncFile(filePath: string, documentPath: string): Promise<SyncDetails> {
        return this.syncFiles([filePath], documentPath);
    }

    async syncFiles(filePaths: string[], documentPath: string): Promise<SyncDetails> {
        const docFileName = path.basename(documentPath);
        const syncDetails: SyncDetails = {
            addedModules: [],
            deletedModules: [],
            clearedModules: [],
            modifiedModules: []
        };
        
        // 显示同步开始通知
        this.notificationManager.showSyncStart(docFileName, 'ideToDoc');
        
        try {
            const vbaFolder = path.dirname(filePaths[0]);
            const tempFolder = path.join(vbaFolder, '.vba-sync-tmp-ide-to-doc');
            
            if (!fs.existsSync(tempFolder)) {
                fs.mkdirSync(tempFolder, { recursive: true });
            }
            
            try {
                // 获取文档中已有的模块列表
                const documentModules: VBAModule[] = await this.connectionManager.listModules(documentPath);
                
                // 导出所有模块到临时文件夹
                await this.connectionManager.exportAll(documentPath, tempFolder);
                
                const filesToSync: string[] = [];
                const deletedFiles: string[] = [];
                
                // 第一层：模块列表比较
                debugLog('=== 第一层：模块列表比较 ===');
                
                // 检查每个文件是否需要同步
                for (const filePath of filePaths) {
                    const moduleFileName = path.basename(filePath);
                    const moduleName = path.basename(filePath, path.extname(filePath));
                    const ext = path.extname(filePath).toLowerCase();
                    
                    if (ext === '.bas' || ext === '.cls') {
                        const tempFilePath = path.join(tempFolder, moduleFileName);
                        
                        // 根据扩展名确定期望的模块类型
                        let expectedTypes: number[];
                        if (ext === '.bas') {
                            expectedTypes = [1];  // 标准模块
                        } else {
                            expectedTypes = [2, 100];  // 类模块或文档模块
                        }
                        
                        // 检查文件是否存在（处理删除事件）
                        if (!fs.existsSync(filePath)) {
                            // 文件已被删除，需要从文档中删除对应模块
                            deletedFiles.push(filePath);
                        } else {
                            // 检查模块是否已存在于文档中（同时比较名称和类型）
                            const moduleExists = documentModules.some(module => 
                                module.name === moduleName && expectedTypes.includes(module.type)
                            );
                            
                            // 第二层：具体模块内容比较
                            debugLog('=== 第二层：具体模块内容比较 ===');
                            if (fs.existsSync(tempFilePath)) {
                                const diffType = FileComparer.compareFilesWithMultiLevel(filePath, tempFilePath);
                                
                                if (diffType === 'identical') {
                                    debugLog('File content is identical, skipping sync for: ' + moduleFileName);
                                    continue;
                                } else if (diffType === 'case-only' || diffType === 'substantial') {
                                    // IDE发起的同步，直接同步所有差异
                                    filesToSync.push(filePath);
                                }
                            } else if (!moduleExists) {
                                // 模块在文档中不存在，是新增
                                filesToSync.push(filePath);
                            } else {
                                // 模块在文档中存在，但临时文件导出失败，仍然视为修改
                                filesToSync.push(filePath);
                            }
                        }
                    }
                }
                
                // 处理删除的文件
                if (deletedFiles.length > 0) {
                    for (const filePath of deletedFiles) {
                        const moduleName = path.basename(filePath, path.extname(filePath));
                        const fileName = path.basename(filePath);
                        const tempFilePath = path.join(tempFolder, fileName);
                        
                        // 检查是否为文档模块
                        let isDocumentModule = this.isDocumentModule(filePath);
                        
                        // 如果文件不存在，从临时导出文件夹中检查（最高优先级）
                        if (!isDocumentModule && !fs.existsSync(filePath) && fs.existsSync(tempFilePath)) {
                            try {
                                const content = fs.readFileSync(tempFilePath).toString();
                                // 检查文档模块标记，支持不同的换行符格式
                                const hasDocumentModuleMarker = content.startsWith('@vba-sync-document-module') || 
                                                            content.startsWith('@vba-sync-document-module\r\n') ||
                                                            content.startsWith('@vba-sync-document-module\n');
                                isDocumentModule = hasDocumentModuleMarker;
                                debugLog('Checked temp file for document module: ' + tempFilePath + ', result: ' + isDocumentModule);
                                debugLog('Temp file content starts with: ' + content.substring(0, 30) + '...');
                            } catch (error) {
                                debugLog('Error checking temp file for document module:', error);
                            }
                        }
                        
                        // 基于模块名称识别文档模块（第二优先级）
                        if (!isDocumentModule) {
                            // Excel文档模块
                            if (moduleName === 'ThisWorkbook' || 
                                moduleName.startsWith('Sheet') || 
                                moduleName.startsWith('工作表')) {
                                isDocumentModule = true;
                                debugLog('Identified Excel document module by name: ' + moduleName);
                            }
                            // Word文档模块
                            else if (moduleName === 'ThisDocument') {
                                isDocumentModule = true;
                                debugLog('Identified Word document module by name: ' + moduleName);
                            }
                            // PowerPoint文档模块
                            else if (moduleName === 'ThisPresentation') {
                                isDocumentModule = true;
                                debugLog('Identified PowerPoint document module by name: ' + moduleName);
                            }
                        }
                        
                        debugLog('Final document module check result for ' + moduleName + ': ' + isDocumentModule);
                        
                        let userChoice;
                        if (isDocumentModule) {
                            // 文档模块：提示清空代码
                            userChoice = await this.notificationManager.showInformationWithOptions(
                                '检测到文档模块【' + moduleName + '】已从IDE中删除，是否清空文档中的对应模块代码？',
                                '清空代码', '取消'
                            );
                        } else {
                            // 其他模块：提示删除
                            userChoice = await this.notificationManager.showInformationWithOptions(
                                '检测到模块【' + moduleName + '】已从IDE中删除，是否从文档中删除该模块？',
                                '删除', '取消'
                            );
                        }
                        
                        if (isDocumentModule && userChoice === '清空代码') {
                            // 清空文档模块代码
                            try {
                                await this.connectionManager.clearModuleCode(documentPath, moduleName);
                                debugLog('Document module code cleared successfully: ' + moduleName);
                                syncDetails.clearedModules.push(moduleName);
                                
                                // 恢复IDE侧的文档模块文件
                                if (!fs.existsSync(filePath)) {
                                    try {
                                        fs.writeFileSync(filePath, '@vba-sync-document-module\n');
                                        debugLog('Document module file restored in IDE: ' + filePath);
                                        this.notificationManager.showInformation('文档模块【' + moduleName + '】文件已在IDE中恢复');
                                    } catch (error) {
                                        debugLog('Error restoring document module file: ' + filePath, error);
                                        this.notificationManager.showError('恢复文档模块文件失败: ' + (error as Error).message);
                                    }
                                }
                                
                                this.notificationManager.showInformation('文档模块【' + moduleName + '】代码已清空');
                            } catch (error) {
                                debugLog('Error clearing document module code: ' + moduleName, error);
                                this.notificationManager.showError('清空文档模块代码失败: ' + (error as Error).message);
                            }
                        } else if (!isDocumentModule && userChoice === '删除') {
                            // 删除其他模块
                            try {
                                await this.connectionManager.removeModule(documentPath, moduleName);
                                debugLog('Module removed successfully: ' + moduleName);
                                syncDetails.deletedModules.push(moduleName);
                                this.notificationManager.showInformation('模块【' + moduleName + '】已从文档中删除');
                            } catch (error: any) {
                                debugLog('Error removing module: ' + moduleName, error);
                                debugLog('Error message: ' + (error.message || 'No message'));
                                // 检查是否是文档模块删除错误
                                if (error.message && (error.message.includes('Cannot delete document modules') || moduleName === 'ThisDocument')) {
                                    debugLog('Cannot delete document module, falling back to clearing code');
                                    // 回退到清空代码操作
                                    try {
                                        await this.connectionManager.clearModuleCode(documentPath, moduleName);
                                        debugLog('Document module code cleared successfully after delete attempt: ' + moduleName);
                                        syncDetails.clearedModules.push(moduleName);
                                        
                                        // 恢复IDE侧的文档模块文件
                                        if (!fs.existsSync(filePath)) {
                                            try {
                                                fs.writeFileSync(filePath, '@vba-sync-document-module\n');
                                                debugLog('Document module file restored in IDE: ' + filePath);
                                                this.notificationManager.showInformation('文档模块【' + moduleName + '】文件已在IDE中恢复');
                                            } catch (fileError) {
                                                debugLog('Error restoring document module file: ' + filePath, fileError);
                                                this.notificationManager.showError('恢复文档模块文件失败: ' + (fileError as Error).message);
                                            }
                                        }
                                        
                                        this.notificationManager.showInformation('文档模块【' + moduleName + '】代码已清空（文档模块无法删除）');
                                    } catch (clearError) {
                                        debugLog('Error clearing document module code: ' + moduleName, clearError);
                                        this.notificationManager.showError('清空文档模块代码失败: ' + (clearError as Error).message);
                                    }
                                } else {
                                    this.notificationManager.showError('删除模块失败: ' + (error.message || '未知错误'));
                                }
                            }
                        }
                    }
                }
                
                // 处理需要同步的文件
                if (filesToSync.length === 0 && deletedFiles.length === 0) {
                    debugLog('No files need to be synced');
                    this.notificationManager.showSyncNotExecuted(docFileName, 'ideToDoc', '文件内容未变更');
                    return syncDetails;
                }
                
                // 执行同步
                if (filesToSync.length > 0) {
                    debugLog('Importing modules: ' + filesToSync.join(', ') + ' to document: ' + documentPath);
                    
                    for (const filePath of filesToSync) {
                        const moduleName = path.basename(filePath, path.extname(filePath));
                        const fileName = path.basename(filePath);
                        const tempFilePath = path.join(tempFolder, fileName);
                        const ext = path.extname(filePath).toLowerCase();
                        
                        await this.connectionManager.importModule(documentPath, filePath, moduleName);
                        debugLog('Module imported successfully: ' + moduleName);
                        
                        // 根据扩展名确定期望的模块类型
                        let expectedTypes: number[];
                        if (ext === '.bas') {
                            expectedTypes = [1];  // 标准模块
                        } else {
                            expectedTypes = [2, 100];  // 类模块或文档模块
                        }
                        
                        // 检查模块是否已存在于文档中（同时比较名称和类型）
                        const moduleExists = documentModules.some(module => 
                            module.name === moduleName && expectedTypes.includes(module.type)
                        );
                        
                        if (moduleExists) {
                            // 模块已存在，是修改
                            syncDetails.modifiedModules.push(moduleName);
                        } else {
                            // 模块不存在，是新增
                            syncDetails.addedModules.push(moduleName);
                        }
                    }
                }
                
                const hasSyncChanges = syncDetails.addedModules.length > 0 || 
                    syncDetails.deletedModules.length > 0 || 
                    syncDetails.clearedModules.length > 0 || 
                    syncDetails.modifiedModules.length > 0;
                
                if (hasSyncChanges) {
                    debugLog('Sync completed with changes');
                    debugLog(`Added: ${syncDetails.addedModules.join(', ')}`);
                    debugLog(`Deleted: ${syncDetails.deletedModules.join(', ')}`);
                    debugLog(`Cleared: ${syncDetails.clearedModules.join(', ')}`);
                    debugLog(`Modified: ${syncDetails.modifiedModules.join(', ')}`);
                }
                
            } finally {
                // 清理临时文件夹
                if (fs.existsSync(tempFolder)) {
                    fs.rmSync(tempFolder, { recursive: true, force: true });
                }
            }
        } catch (error: any) {
                debugLog('Error syncing files to document:', error);
                // 检查是否是文档访问异常错误
                if (error.message && error.message.includes('文档访问异常，请关闭已打开的')) {
                    this.notificationManager.showSyncFailure(docFileName, 'ideToDoc', error.message);
                } else {
                    this.notificationManager.showSyncFailure(docFileName, 'ideToDoc', error.message || '未知错误');
                }
                throw error;
            }
            
            return syncDetails;
        }

    compareFiles(ideFile: string, docFile: string): boolean {
        try {
            // 读取两个文件的内容
            const ideContent = fs.readFileSync(ideFile).toString().toLowerCase();
            const docContent = fs.readFileSync(docFile).toString().toLowerCase();
            
            // 比较内容（不考虑大小写）
            return ideContent !== docContent;
        } catch (error) {
            debugLog('Error comparing files:', error);
            return true; // 出错时默认认为文件不同
        }
    }

    private isDocumentModule(filePath: string): boolean {
        try {
            if (fs.existsSync(filePath)) {
                const content = fs.readFileSync(filePath).toString();
                return content.startsWith('@vba-sync-document-module\n');
            }
            return false;
        } catch (error) {
            debugLog('Error checking if document module:', error);
            return false;
        }
    }
}