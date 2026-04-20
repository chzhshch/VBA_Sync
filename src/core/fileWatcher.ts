import * as vscode from 'vscode';
import * as path from 'path';
import { SyncEngine } from './sync/SyncEngine';
import { OFFICE_MACRO_EXTENSIONS, VBA_FILE_EXTENSIONS } from '../types';

// 调试日志开关
const DEBUG = true;

// 统一日志函数
function debugLog(message: string, ...args: any[]) {
    if (DEBUG) {
        console.log('[VBA SYNC DEBUG] ' + message, ...args);
    }
}

export class FileWatcher {
    private syncEngine: SyncEngine;
    private watchers: vscode.FileSystemWatcher[] = [];
    private debounceTimers = new Map<string, NodeJS.Timeout>();
    private pendingFileChanges = new Map<string, Set<string>>();
    private isProcessingSave = false;
    
    private isTempFolder(filePath: string): boolean {
        const fileName = path.basename(filePath);
        return fileName.startsWith('.vba-sync-tmp-');
    }

    private isInRefractoryPeriod(filePath: string): boolean {
        const docConfig = this.syncEngine.getConfigManager().getDocumentConfigForFile(filePath);
        if (!docConfig) {
            return false;
        }

        const workspaceRoot = this.syncEngine.getConfigManager().getWorkspaceRoot();
        if (!workspaceRoot) {
            return false;
        }

        const documentPath = path.join(workspaceRoot, docConfig.path);
        const state = this.syncEngine.getStateManager().getState(documentPath);
        
        if (state?.lastDocToIdeSyncTime) {
            const config = vscode.workspace.getConfiguration('vbaSync');
            const refractoryPeriod = config.get('syncRefractoryPeriod', 10000);
            const timeSinceLastSync = Date.now() - state.lastDocToIdeSyncTime.getTime();
            
            return timeSinceLastSync < refractoryPeriod;
        }

        return false;
    }

    constructor(syncEngine: SyncEngine) {
        this.syncEngine = syncEngine;
        this.setupWatchers();
    }

    private setupWatchers() {
        if (!vscode.workspace.workspaceFolders) return;

        const workspaceRoot = vscode.workspace.workspaceFolders[0].uri.fsPath;
        const config = this.syncEngine.getConfigManager().getConfig();

        // 监控 Office 文档文件
        for (const ext of OFFICE_MACRO_EXTENSIONS) {
            const pattern = new vscode.RelativePattern(workspaceRoot, `**/*${ext}`);
            const watcher = vscode.workspace.createFileSystemWatcher(pattern);
            
            watcher.onDidChange(this.handleDocumentChange.bind(this));
            this.watchers.push(watcher);
        }

        // 监控 VBA 文件 - 只监控已配置文档的 VBA 文件夹
        if (config) {
            for (const doc of config.documents) {
                if (doc.enabled) {
                    const vbaFolder = path.join(workspaceRoot, doc.vbaFolder);
                    
                    for (const ext of VBA_FILE_EXTENSIONS) {
                        const pattern = new vscode.RelativePattern(vbaFolder, `**/*${ext}`);
                        const watcher = vscode.workspace.createFileSystemWatcher(pattern);
                        
                        watcher.onDidChange(this.handleVbaFileChange.bind(this));
                        watcher.onDidCreate(this.handleVbaFileCreate.bind(this));
                        watcher.onDidDelete(this.handleVbaFileDelete.bind(this));
                        
                        this.watchers.push(watcher);
                    }
                }
            }
        }
    }

    private handleDocumentChange(uri: vscode.Uri) {
        const documentPath = uri.fsPath;
        this.debounce(documentPath, () => {
            this.syncEngine.syncDocumentToIde(documentPath);
        });
    }

    private handleVbaFileChange(uri: vscode.Uri) {
        const filePath = uri.fsPath;
        
        if (this.isTempFolder(filePath) || filePath.includes('.vba-sync-tmp-')) {
            return;
        }
        
        // 检查是否在不应期内
        if (this.isInRefractoryPeriod(filePath)) {
            debugLog(`Skipping file change due to refractory period: ${filePath}`);
            return;
        }
        
        // 如果正在处理保存操作，跳过以避免重复触发
        if (this.isProcessingSave) {
            return;
        }
        
        this.collectAndDebounceVbaFile(filePath);
    }

    private async handleVbaFileCreate(uri: vscode.Uri) {
        const filePath = uri.fsPath;
        
        if (this.isTempFolder(filePath) || filePath.includes('.vba-sync-tmp-')) {
            return;
        }
        
        // 检查是否在不应期内
        if (this.isInRefractoryPeriod(filePath)) {
            debugLog(`File create detected during refractory period: ${filePath}, checking module list consistency`);
            
            // 获取文档配置和完整路径
            const docConfig = this.syncEngine.getConfigManager().getDocumentConfigForFile(filePath);
            if (!docConfig) {
                debugLog('Document config not found, skipping');
                return;
            }

            const workspaceRoot = this.syncEngine.getConfigManager().getWorkspaceRoot();
            if (!workspaceRoot) {
                debugLog('No workspace root, skipping');
                return;
            }

            const documentPath = path.join(workspaceRoot, docConfig.path);
            const vbaFolder = path.join(workspaceRoot, docConfig.vbaFolder);
            
            // 检查模块列表一致性
            const isConsistent = await this.syncEngine.checkModuleListConsistency(documentPath, vbaFolder);
            
            if (isConsistent) {
                debugLog('Module list is consistent, skipping sync');
                return;
            } else {
                debugLog('Module list is inconsistent, proceeding with sync');
            }
        }
        
        this.collectAndDebounceVbaFile(filePath);
    }

    private collectAndDebounceVbaFile(filePath: string) {
        // 获取文件对应的文档配置
        const docConfig = this.syncEngine.getConfigManager().getDocumentConfigForFile(filePath);
        if (!docConfig) {
            return;
        }

        const workspaceRoot = this.syncEngine.getConfigManager().getWorkspaceRoot();
        if (!workspaceRoot) {
            return;
        }

        const documentPath = path.join(workspaceRoot, docConfig.path);
        
        // 检查同一文件夹内是否有未保存的VBA文件
        const unsavedFiles = this.getUnsavedVbaFilesInSameFolder(filePath);
        
        if (unsavedFiles.length > 0) {
            // 有未保存的文件，显示选择对话框
            this.debounce(documentPath, async () => {
                const choice = await this.showSyncChoiceDialog(unsavedFiles, filePath);
                
                if (choice === 'syncAll') {
                    // 同步所有文件，先保存未保存的文件
                    this.isProcessingSave = true;
                    try {
                        await this.saveAllFiles(unsavedFiles);
                        // 收集所有需要同步的文件（包括当前文件和未保存的文件）
                        const allFiles = new Set<string>();
                        allFiles.add(filePath);
                        unsavedFiles.forEach(f => allFiles.add(f));
                        
                        // 等待一小段时间，确保文件保存完成
                        await new Promise(resolve => setTimeout(resolve, 200));
                        
                        // 统一进行同步
                        await this.syncEngine.syncFilesToDocument(Array.from(allFiles), documentPath);
                    } finally {
                        this.isProcessingSave = false;
                    }
                } else {
                    // 仅同步当前文件
                    await this.syncEngine.syncFilesToDocument([filePath], documentPath);
                }
                
                this.pendingFileChanges.delete(documentPath);
            });
        } else {
            // 没有未保存的文件，正常处理
            if (!this.pendingFileChanges.has(documentPath)) {
                this.pendingFileChanges.set(documentPath, new Set());
            }
            this.pendingFileChanges.get(documentPath)!.add(filePath);
            
            this.debounce(documentPath, () => {
                const files = this.pendingFileChanges.get(documentPath);
                if (files && files.size > 0) {
                    this.syncEngine.syncFilesToDocument(Array.from(files), documentPath);
                }
                this.pendingFileChanges.delete(documentPath);
            });
        }
    }
    
    private getUnsavedVbaFilesInSameFolder(filePath: string): string[] {
        const folderPath = path.dirname(filePath);
        const unsavedFiles: string[] = [];
        
        // 获取所有未保存的文本编辑器
        const unsavedEditors = vscode.workspace.textDocuments.filter(doc => 
            doc.isDirty && 
            VBA_FILE_EXTENSIONS.some(ext => doc.fileName.endsWith(ext)) &&
            path.dirname(doc.fileName) === folderPath &&
            doc.fileName !== filePath
        );
        
        unsavedEditors.forEach(doc => {
            unsavedFiles.push(doc.fileName);
        });
        
        return unsavedFiles;
    }
    
    private async showSyncChoiceDialog(unsavedFiles: string[], currentFile: string): Promise<'syncCurrent' | 'syncAll'> {
        const fileNames = unsavedFiles.map(f => path.basename(f)).join(', ');
        const message = `检测到同一文件夹内有 ${unsavedFiles.length} 个未保存的文件：${fileNames}\n\n您想要如何处理？`;
        
        const choice = await vscode.window.showInformationMessage(
            message,
            { modal: true },
            '仅同步当前文件',
            '同步所有文件'
        );
        
        return choice === '同步所有文件' ? 'syncAll' : 'syncCurrent';
    }
    
    private async saveAllFiles(filePaths: string[]): Promise<void> {
        const savePromises: Promise<boolean>[] = [];
        
        filePaths.forEach(filePath => {
            const doc = vscode.workspace.textDocuments.find(d => d.fileName === filePath);
            if (doc && doc.isDirty) {
                savePromises.push(Promise.resolve(doc.save()));
            }
        });
        
        await Promise.all(savePromises);
    }

    private async handleVbaFileDelete(uri: vscode.Uri) {
        // 处理文件删除逻辑
        const filePath = uri.fsPath;
        
        if (this.isTempFolder(filePath) || filePath.includes('.vba-sync-tmp-')) {
            return;
        }
        
        // 检查是否在不应期内
        if (this.isInRefractoryPeriod(filePath)) {
            debugLog(`File delete detected during refractory period: ${filePath}, checking module list consistency`);
            
            // 获取文档配置和完整路径
            const docConfig = this.syncEngine.getConfigManager().getDocumentConfigForFile(filePath);
            if (!docConfig) {
                debugLog('Document config not found, skipping');
                return;
            }

            const workspaceRoot = this.syncEngine.getConfigManager().getWorkspaceRoot();
            if (!workspaceRoot) {
                debugLog('No workspace root, skipping');
                return;
            }

            const documentPath = path.join(workspaceRoot, docConfig.path);
            const vbaFolder = path.join(workspaceRoot, docConfig.vbaFolder);
            
            // 检查模块列表一致性
            const isConsistent = await this.syncEngine.checkModuleListConsistency(documentPath, vbaFolder);
            
            if (isConsistent) {
                debugLog('Module list is consistent, skipping sync');
                return;
            } else {
                debugLog('Module list is inconsistent, proceeding with sync');
            }
        }
        
        // 如果正在处理保存操作，跳过以避免重复触发
        if (this.isProcessingSave) {
            return;
        }
        
        // 检查未保存文件并处理
        debugLog('Handling VBA file delete: ' + filePath);
        this.collectAndDebounceVbaFile(filePath);
    }

    private debounce(key: string, callback: () => void) {
        const config = vscode.workspace.getConfiguration('vbaSync');
        const syncDelay = config.get('syncDelay', 500);

        if (this.debounceTimers.has(key)) {
            clearTimeout(this.debounceTimers.get(key)!);
        }

        const timer = setTimeout(() => {
            callback();
            this.debounceTimers.delete(key);
        }, syncDelay);

        this.debounceTimers.set(key, timer);
    }

    refresh() {
        // 清理现有 watcher
        this.dispose();
        // 重新创建 watcher
        this.setupWatchers();
    }

    dispose() {
        for (const watcher of this.watchers) {
            watcher.dispose();
        }
        this.watchers = [];

        for (const timer of this.debounceTimers.values()) {
            clearTimeout(timer);
        }
        this.debounceTimers.clear();
        
        this.pendingFileChanges.clear();
    }
}
