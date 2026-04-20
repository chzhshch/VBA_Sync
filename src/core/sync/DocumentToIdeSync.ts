import * as fs from 'fs';
import * as path from 'path';
import * as vscode from 'vscode';
import { ConnectionManager } from '../connection/ConnectionManager';
import { FileComparer, FileDifference } from '../compare/FileComparer';
import { NotificationManager } from '../notification/NotificationManager';
import { SyncDetails } from '../../types';

// 调试日志开关
const DEBUG = true;

// 统一日志函数
function debugLog(message: string, ...args: any[]) {
    if (DEBUG) {
        console.log('[VBA SYNC DEBUG] ' + message, ...args);
    }
}

export class DocumentToIdeSync {
    private connectionManager: ConnectionManager;
    private notificationManager: NotificationManager;
    private openDiffViewCount: number = 0;
    private diffContext: {
        tempFolder: string;
        vbaFolder: string;
        documentName: string;
        changedFiles: string[];
    } | null = null;
    private diffDisposables: vscode.Disposable[] = [];
    private syncDetails: SyncDetails;
    private pendingCleanupTempFolder: string | null = null;

    constructor(connectionManager: ConnectionManager, notificationManager: NotificationManager) {
        this.connectionManager = connectionManager;
        this.notificationManager = notificationManager;
        this.syncDetails = {
            addedModules: [],
            deletedModules: [],
            clearedModules: [],
            modifiedModules: []
        };
    }

    async exportAndSync(documentPath: string, vbaFolder: string): Promise<SyncDetails> {
        const syncDocFileName = path.basename(documentPath);
        
        // 重置同步详细信息
        this.syncDetails = {
            addedModules: [],
            deletedModules: [],
            clearedModules: [],
            modifiedModules: []
        };
        
        // 显示同步开始通知
        this.notificationManager.showSyncStart(syncDocFileName, 'docToIde');
        
        // 创建临时文件夹
        const tempFolder = path.join(vbaFolder, '.vba-sync-tmp-doc-to-ide');
        debugLog('VBA folder: ' + vbaFolder + ', Temp folder: ' + tempFolder);

        // 确保临时文件夹存在
        if (!fs.existsSync(tempFolder)) {
            fs.mkdirSync(tempFolder, { recursive: true });
            debugLog('Created temp folder: ' + tempFolder);
        }

        try {
            // 导出所有模块到临时文件夹
            debugLog('Exporting all modules to temp folder');
            await this.connectionManager.exportAll(documentPath, tempFolder);
            debugLog('Export completed');
            
            // 第一层：模块列表比较
            debugLog('=== 第一层：模块列表比较 ===');
            const differences = FileComparer.compareFilesWithDetail(tempFolder, vbaFolder, false);
            debugLog('Found ' + differences.length + ' differences');

            // 分类文件
            const caseOnlyFiles = differences.filter(d => d.type === 'case-only').map(d => d.fileName);
            const newFiles = differences.filter(d => d.type === 'new').map(d => d.fileName);
            const deletedFiles = differences.filter(d => d.type === 'deleted').map(d => d.fileName);
            const substantialFiles = differences.filter(d => d.type === 'substantial').map(d => d.fileName);

            debugLog('Case-only files: ' + caseOnlyFiles.join(', '));
            debugLog('New files: ' + newFiles.join(', '));
            debugLog('Deleted files: ' + deletedFiles.join(', '));
            debugLog('Substantial files: ' + substantialFiles.join(', '));

            // 第二层：具体模块处理
            debugLog('=== 第二层：具体模块处理 ===');
            
            // 处理新增文件和仅大小写差异（自动同步）
            const autoSyncFiles = [...newFiles, ...caseOnlyFiles];
            if (autoSyncFiles.length > 0) {
                await this.autoSyncFiles(tempFolder, vbaFolder, autoSyncFiles, syncDocFileName);
            }

            // 处理删除文件（需要确认）
            if (deletedFiles.length > 0) {
                await this.handleDeletedFiles(deletedFiles, vbaFolder);
            }

            // 处理实质性差异（原流程）
            if (substantialFiles.length > 0) {
                await this.handleSubstantialFiles(tempFolder, vbaFolder, syncDocFileName, substantialFiles);
            }

            // 等待用户完成同步操作（如果选择了查看差异）
            if (this.diffContext) {
                await new Promise<void>((resolve) => {
                    const checkInterval = setInterval(() => {
                        if (!this.diffContext) {
                            clearInterval(checkInterval);
                            resolve();
                        }
                    }, 500);
                });
            }

            // 显示统一的同步结果通知
            const hasSyncChanges = this.syncDetails.addedModules.length > 0 || 
                this.syncDetails.deletedModules.length > 0 || 
                this.syncDetails.modifiedModules.length > 0;
            
            if (hasSyncChanges) {
                debugLog('Sync completed with changes');
                debugLog(`Added: ${this.syncDetails.addedModules.join(', ')}`);
                debugLog(`Deleted: ${this.syncDetails.deletedModules.join(', ')}`);
                debugLog(`Modified: ${this.syncDetails.modifiedModules.join(', ')}`);
            } else if (differences.length > 0) {
                // 有差异但未同步
                this.notificationManager.showSyncNotExecuted(syncDocFileName, 'docToIde', '用户取消同步');
            } else {
                // 没有变更
                this.notificationManager.showSyncNotExecuted(syncDocFileName, 'docToIde', '未检测到变更');
            }
        } catch (error: any) {
            debugLog('Error in exportAndSync:', error);
            // 检查是否是文档访问异常错误
            if (error.message && error.message.includes('文档访问异常，请关闭已打开的')) {
                this.notificationManager.showSyncFailure(syncDocFileName, 'docToIde', error.message);
            } else {
                this.notificationManager.showSyncFailure(syncDocFileName, 'docToIde', error.message || '未知错误');
            }
        } finally {
            // 清理临时文件夹
            this.cleanupTempFolder(tempFolder);
            this.pendingCleanupTempFolder = null;
        }
        
        // 返回同步详细信息
        return { ...this.syncDetails };
    }

    private async compareAndSync(tempFolder: string, vbaFolder: string, documentName: string, changedFiles: string[]) {
        debugLog('Syncing changed files to IDE');
        
        // 再次检查临时文件夹是否存在
        if (fs.existsSync(tempFolder)) {
            let syncFailed = false;
            
            for (const file of changedFiles) {
                const tempPath = path.join(tempFolder, file);
                const targetPath = path.join(vbaFolder, file);
                
                // 检查文件是否存在
                if (fs.existsSync(tempPath)) {
                    debugLog('Copying ' + tempPath + ' to ' + targetPath);
                    fs.copyFileSync(tempPath, targetPath);
                } else {
                    debugLog('Temp file not found: ' + tempPath);
                    syncFailed = true;
                    break;
                }
            }

            if (!syncFailed) {
                debugLog('Sync to IDE completed successfully');
                

            } else {
                this.notificationManager.showError('同步失败：临时文件不存在');
            }
        } else {
            debugLog('Temp folder not found: ' + tempFolder);
            this.notificationManager.showError('同步失败：临时文件夹不存在');
        }
    }

    private async autoSyncFiles(tempFolder: string, vbaFolder: string, files: string[], documentName: string) {
        debugLog('Auto syncing files: ' + files.join(', '));
        
        for (const file of files) {
            const tempPath = path.join(tempFolder, file);
            const targetPath = path.join(vbaFolder, file);
            const moduleName = path.basename(file, path.extname(file));
            
            if (fs.existsSync(tempPath)) {
                debugLog('Copying ' + tempPath + ' to ' + targetPath);
                fs.copyFileSync(tempPath, targetPath);
                this.syncDetails.addedModules.push(moduleName);
            } else {
                debugLog('Temp file not found: ' + tempPath);
            }
        }
        

    }

    private async handleDeletedFiles(deletedFiles: string[], vbaFolder: string) {
        debugLog('Handling deleted files: ' + deletedFiles.join(', '));
        
        if (deletedFiles.length === 0) {
            return;
        }
        
        const moduleNames = deletedFiles.map(file => path.basename(file)).join('、');
        const userChoice = await this.notificationManager.showInformationWithOptions(
            '检测到以下模块已从文档中删除：【' + moduleNames + '】，是否从 IDE 中删除这些模块？',
            '删除', '取消'
        );
        
        if (userChoice === '删除') {
            for (const file of deletedFiles) {
                const targetPath = path.join(vbaFolder, file);
                const moduleName = path.basename(file, path.extname(file));
                if (fs.existsSync(targetPath)) {
                    debugLog('Deleting file: ' + targetPath);
                    fs.unlinkSync(targetPath);
                    this.syncDetails.deletedModules.push(moduleName);
                }
            }
        } else {
            debugLog('Delete cancelled by user');
        }
    }

    private async handleSubstantialFiles(tempFolder: string, vbaFolder: string, documentName: string, substantialFiles: string[]) {
        debugLog('Handling substantial files: ' + substantialFiles.join(', '));
        
        if (substantialFiles.length === 0) {
            return;
        }
        
        const moduleNames = substantialFiles.map(file => path.basename(file)).join('、');
        const userChoice = await this.notificationManager.showInformationWithOptions(
            '文档【' + documentName + '】检测到 ' + substantialFiles.length + ' 个文件有变更：【' + moduleNames + '】，是否同步到 IDE？',
            '查看差异', '同步', '取消'
        );
        
        debugLog('User choice: ' + userChoice);

        if (userChoice === '查看差异') {
            this.diffContext = {
                tempFolder,
                vbaFolder,
                documentName,
                changedFiles: substantialFiles
            };
            await this.showDiff(tempFolder, vbaFolder, substantialFiles);
        } else if (userChoice === '同步') {
            await this.syncFiles(tempFolder, vbaFolder, substantialFiles);
        } else {
            debugLog('Sync cancelled by user or no choice made');
        }
    }

    private async syncFiles(tempFolder: string, vbaFolder: string, files: string[]) {
        debugLog('Syncing files: ' + files.join(', '));
        
        for (const file of files) {
            const tempPath = path.join(tempFolder, file);
            const targetPath = path.join(vbaFolder, file);
            const moduleName = path.basename(file, path.extname(file));
            
            if (fs.existsSync(tempPath)) {
                debugLog('Copying ' + tempPath + ' to ' + targetPath);
                fs.copyFileSync(tempPath, targetPath);
                this.syncDetails.modifiedModules.push(moduleName);
            } else {
                debugLog('Temp file not found: ' + tempPath);
            }
        }
        

    }

    private async showDiff(tempFolder: string, vbaFolder: string, changedFiles: string[]) {
        debugLog('Showing diff for changed files');
        
        this.openDiffViewCount = 0;
        this.disposeDiffDisposables();
        
        const tabGroups = vscode.window.tabGroups;
        
        for (let i = 0; i < changedFiles.length; i++) {
            const file = changedFiles[i];
            const tempPath = path.join(tempFolder, file);
            const targetPath = path.join(vbaFolder, file);
            
            const tempUri = vscode.Uri.file(tempPath);
            const targetUri = vscode.Uri.file(targetPath);
            
            const tempTitle = '文档中的 ' + file;
            const targetTitle = 'IDE 中的 ' + file;
            
            debugLog('Opening diff for: ' + file + ' in viewColumn: ViewColumn.Beside');
            
            // URI顺序：IDE在左，文档在右
            // 标题顺序：IDE↔文档，与实际展示顺序一致
            // 使用ViewColumn.Beside，不受9列限制
            await vscode.commands.executeCommand('vscode.diff', targetUri, tempUri, `${targetTitle} ↔ ${tempTitle}`, {
                viewColumn: vscode.ViewColumn.Beside
            });
            
            this.openDiffViewCount++;
        }
        
        debugLog('Opened ' + this.openDiffViewCount + ' diff views');
        
        const tabDisposable = tabGroups.onDidChangeTabs(async (e) => {
            // 检查所有打开的标签页，统计实际的diff窗口数
            const openDiffTabs = tabGroups.all.flatMap(group => 
                group.tabs.filter(tab => {
                    if (tab.input) {
                        const input = tab.input as any;
                        return input.viewType === 'vscode.diff' || (input.original && input.modified);
                    }
                    return false;
                })
            );
            
            const actualOpenDiffCount = openDiffTabs.length;
            debugLog('Actual open diff tabs: ' + actualOpenDiffCount);
            
            // 更新计数器为实际打开的diff窗口数
            this.openDiffViewCount = actualOpenDiffCount;
            
            if (actualOpenDiffCount <= 0) {
                debugLog('All diff views closed, showing sync confirmation');
                await this.onAllDiffViewsClosed();
            } else {
                debugLog('Remaining diff views: ' + actualOpenDiffCount);
            }
        });
        
        this.diffDisposables.push(tabDisposable);
    }
    
    private async onAllDiffViewsClosed() {
        this.disposeDiffDisposables();
        
        if (!this.diffContext) {
            return;
        }
        
        const { tempFolder, vbaFolder, documentName, changedFiles } = this.diffContext;
        const moduleNames = changedFiles.map(file => path.basename(file)).join('、');
        
        const userChoice = await this.notificationManager.showInformationWithOptions(
            '文档【' + documentName + '】检测到 ' + changedFiles.length + ' 个文件有变更：【' + moduleNames + '】，是否同步到 IDE？',
            '同步', '取消'
        );
        
        debugLog('User choice after diff: ' + userChoice);
        
        if (userChoice === '同步') {
            await this.syncFiles(tempFolder, vbaFolder, changedFiles);
        } else {
            debugLog('Sync cancelled by user after diff');
        }
        
        // 清理临时文件夹
        this.cleanupTempFolder(tempFolder);
        this.pendingCleanupTempFolder = null;
        
        this.diffContext = null;
    }
    
    private cleanupTempFolder(tempFolder: string) {
        if (fs.existsSync(tempFolder)) {
            debugLog('Cleaning up temp folder');
            fs.rmSync(tempFolder, { recursive: true, force: true });
            debugLog('Temp folder cleaned up');
        }
    }
    
    private disposeDiffDisposables() {
        debugLog('Disposing diff disposables');
        this.diffDisposables.forEach(d => d.dispose());
        this.diffDisposables = [];
    }
}