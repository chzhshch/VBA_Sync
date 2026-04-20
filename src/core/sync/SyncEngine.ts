import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import { ConnectionManager } from '../connection/ConnectionManager';
import { DocumentToIdeSync } from './DocumentToIdeSync';
import { IdeToDocumentSync } from './IdeToDocumentSync';
import { ConfigManager } from '../config/ConfigManager';
import { StateManager } from '../state/StateManager';
import { ConflictResolver } from '../conflict/ConflictResolver';
import { NotificationManager } from '../notification/NotificationManager';
import { ErrorHandler } from '../error/ErrorHandler';
import { FileComparer } from '../compare/FileComparer';

// 调试日志开关
const DEBUG = true;

// 统一日志函数
function debugLog(message: string, ...args: any[]) {
    if (DEBUG) {
        console.log('[VBA SYNC DEBUG] ' + message, ...args);
    }
}

export class SyncEngine {
    private connectionManager: ConnectionManager;
    private documentToIdeSync: DocumentToIdeSync;
    private ideToDocumentSync: IdeToDocumentSync;
    private configManager: ConfigManager;
    private stateManager: StateManager;
    private conflictResolver: ConflictResolver | null = null;
    private notificationManager: NotificationManager;
    
    private isSyncingDocToIde: boolean = false;
    private isSyncingIdeToDoc: boolean = false;

    constructor(
        connectionManager: ConnectionManager,
        configManager: ConfigManager,
        stateManager: StateManager,
        notificationManager: NotificationManager
    ) {
        this.connectionManager = connectionManager;
        this.configManager = configManager;
        this.stateManager = stateManager;
        this.notificationManager = notificationManager;
        
        this.documentToIdeSync = new DocumentToIdeSync(connectionManager, notificationManager);
        this.ideToDocumentSync = new IdeToDocumentSync(connectionManager, notificationManager);
        
        this.initConflictResolver();
    }

    private initConflictResolver() {
        const config = this.configManager.getConfig();
        if (config) {
            this.conflictResolver = new ConflictResolver(config.settings);
        }
    }

    async syncDocumentToIde(documentPath: string) {
        return ErrorHandler.wrapAsync(async () => {
            if (this.isSyncingIdeToDoc) {
                debugLog('Skipping syncDocumentToIde because isSyncingIdeToDoc is true');
                return;
            }
            
            this.isSyncingDocToIde = true;
            try {
                debugLog('Syncing document to IDE: ' + documentPath);
                const docConfig = this.configManager.getDocumentConfig(documentPath);
                
                if (!docConfig) {
                    debugLog('Document config not found');
                    this.notificationManager.showError('未找到文档配置');
                    return;
                }

                const workspaceRoot = this.configManager.getWorkspaceRoot();
                if (!workspaceRoot) {
                    debugLog('No workspace folder found');
                    this.notificationManager.showError('未找到工作区文件夹');
                    return;
                }

                const vbaFolder = path.join(workspaceRoot, docConfig.vbaFolder);
                
                // 确保VBA文件夹存在
                if (!fs.existsSync(vbaFolder)) {
                    fs.mkdirSync(vbaFolder, { recursive: true });
                }

                // 检查文档是否打开
                let state = this.stateManager.getState(documentPath);
                if (!state?.isOfficeOpen || !this.connectionManager.hasConnection(documentPath)) {
                    debugLog('Document not open or no valid connection, opening...');
                    const fileName = path.basename(documentPath);
                    this.notificationManager.showInformation('正在打开文档【' + fileName + '】...');
                    
                    const appType = this.connectionManager.getAppType(documentPath);
                    const success = await this.connectionManager.openDocument(documentPath, appType, false, 'doc-to-ide');
                    
                    if (success) {
                        state = this.stateManager.setState(documentPath, { isOfficeOpen: true });
                        this.connectionManager.addConnection(documentPath);
                        debugLog('Document opened successfully');
                    } else {
                        throw new Error('Failed to open document');
                    }
                }

                // 检查VBA访问权限
                debugLog('Checking VBA access for document: ' + documentPath);
                const vbaAccess = await this.connectionManager.checkVbaAccess(documentPath, 'doc-to-ide');
                debugLog('VBA access check result: ' + (vbaAccess.success ? 'success' : 'error'));
                debugLog('VBA access error: ' + vbaAccess.error);
                
                // 如果VBA访问失败，处理错误
                if (!vbaAccess.success) {
                    await this.handleVbaAccessError(documentPath, vbaAccess.error);
                }

                // 执行同步
                const syncDetails = await this.documentToIdeSync.exportAndSync(documentPath, vbaFolder);
                
                const hasSyncChanges = syncDetails.addedModules.length > 0 || 
                    syncDetails.deletedModules.length > 0 || 
                    syncDetails.modifiedModules.length > 0;
                
                if (hasSyncChanges) {
                    const docName = path.basename(documentPath);
                    this.notificationManager.showSyncSuccessDetails(docName, syncDetails, '文档');
                    
                    // 更新状态 - 记录文档到IDE同步时间用于不应期，时间精确到秒
                    const now = new Date(Math.floor(Date.now() / 1000) * 1000);
                    this.stateManager.setState(documentPath, {
                        lastSyncTime: now,
                        lastDocToIdeSyncTime: now,
                        ideCodeSynced: true
                    });
                    debugLog('State updated for document: ' + documentPath);
                }
                
                // 检查并配置文件关联
                await this.checkAndConfigureFileAssociations(vbaFolder);
            } finally {
                this.isSyncingDocToIde = false;
            }
        });
    }

    async syncFileToDocument(filePath: string) {
        return ErrorHandler.wrapAsync(async () => {
            await this.syncFilesToDocument([filePath], null);
        });
    }

    async syncFilesToDocument(filePaths: string[], documentPath: string | null) {
        return ErrorHandler.wrapAsync(async () => {
            if (this.isSyncingDocToIde) {
                debugLog('Skipping syncFilesToDocument because isSyncingDocToIde is true');
                return;
            }

            // 检查是否在不应期内
            let targetDocumentPath = documentPath;
            if (!targetDocumentPath && filePaths.length > 0) {
                const docConfig = this.configManager.getDocumentConfigForFile(filePaths[0]);
                if (docConfig) {
                    const workspaceRoot = this.configManager.getWorkspaceRoot();
                    if (workspaceRoot) {
                        targetDocumentPath = path.join(workspaceRoot, docConfig.path);
                    }
                }
            }

            if (targetDocumentPath) {
                const state = this.stateManager.getState(targetDocumentPath);
                if (state?.lastDocToIdeSyncTime) {
                    const config = vscode.workspace.getConfiguration('vbaSync');
                    const refractoryPeriod = config.get('syncRefractoryPeriod', 10000);
                    const timeSinceLastSync = Date.now() - state.lastDocToIdeSyncTime.getTime();
                    
                    if (timeSinceLastSync < refractoryPeriod) {
                        debugLog(`Skipping sync due to refractory period: ${timeSinceLastSync}ms < ${refractoryPeriod}ms`);
                        return;
                    }
                }
            }
            
            this.isSyncingIdeToDoc = true;
            try {
                debugLog('Syncing files to document: ' + filePaths.join(', '));
                
                let targetDocumentPath = documentPath;
                let docConfig = null;
                
                // 如果没有提供documentPath，从第一个文件获取
                if (!targetDocumentPath && filePaths.length > 0) {
                    docConfig = this.configManager.getDocumentConfigForFile(filePaths[0]);
                    if (!docConfig) {
                        debugLog('File not in any VBA sync folder');
                        this.notificationManager.showError('文件不在任何VBA同步文件夹中');
                        throw new Error('File not in any VBA sync folder');
                    }
                    
                    const workspaceRoot = this.configManager.getWorkspaceRoot();
                    if (!workspaceRoot) {
                        debugLog('No workspace folder found');
                        this.notificationManager.showError('未找到工作区文件夹');
                        throw new Error('No workspace folder found');
                    }
                    
                    targetDocumentPath = path.join(workspaceRoot, docConfig.path);
                } else if (targetDocumentPath) {
                    // 如果提供了documentPath，获取文档配置
                    docConfig = this.configManager.getDocumentConfig(targetDocumentPath);
                }
                
                if (!targetDocumentPath || !docConfig) {
                    debugLog('Invalid document configuration');
                    this.notificationManager.showError('无效的文档配置');
                    throw new Error('Invalid document configuration');
                }

                const workspaceRoot = this.configManager.getWorkspaceRoot();
                if (!workspaceRoot) {
                    debugLog('No workspace folder found');
                    this.notificationManager.showError('未找到工作区文件夹');
                    throw new Error('No workspace folder found');
                }

                const vbaFolder = path.join(workspaceRoot, docConfig.vbaFolder);
                
                // 检查文档是否打开
                let state = this.stateManager.getState(targetDocumentPath);
                if (!state?.isOfficeOpen || !this.connectionManager.hasConnection(targetDocumentPath)) {
                    debugLog('Document not open or no valid connection, opening...');
                    const fileName = path.basename(targetDocumentPath);
                    this.notificationManager.showInformation('正在打开文档【' + fileName + '】...');
                    
                    const appType = this.connectionManager.getAppType(targetDocumentPath);
                    const success = await this.connectionManager.openDocument(targetDocumentPath, appType, false, 'ide-to-doc');
                    
                    if (success) {
                        state = this.stateManager.setState(targetDocumentPath, { isOfficeOpen: true });
                        this.connectionManager.addConnection(targetDocumentPath);
                        debugLog('Document opened successfully');
                    } else {
                        throw new Error('Failed to open document');
                    }
                }

                // 检查VBA访问权限
                debugLog('Checking VBA access for document: ' + targetDocumentPath);
                const vbaAccess = await this.connectionManager.checkVbaAccess(targetDocumentPath, 'ide-to-doc');
                debugLog('VBA access check result: ' + (vbaAccess.success ? 'success' : 'error'));
                debugLog('VBA access error: ' + vbaAccess.error);
                
                // 如果VBA访问失败，处理错误
                if (!vbaAccess.success) {
                    await this.handleVbaAccessError(targetDocumentPath, vbaAccess.error);
                }

                // 检查冲突
                const shouldSync = await this.checkConflict(targetDocumentPath);
                if (!shouldSync) {
                    debugLog('Sync cancelled due to conflict');
                    return;
                }

                // 执行批量同步
                const syncDetails = await this.ideToDocumentSync.syncFiles(filePaths, targetDocumentPath);
                
                const hasSyncChanges = syncDetails.addedModules.length > 0 || 
                    syncDetails.deletedModules.length > 0 || 
                    syncDetails.clearedModules.length > 0 || 
                    syncDetails.modifiedModules.length > 0;
                
                if (hasSyncChanges) {
                    const docName = path.basename(targetDocumentPath);
                    this.notificationManager.showSyncSuccessDetails(docName, syncDetails, 'IDE');
                }
                
                // 更新状态
                state = this.stateManager.setState(targetDocumentPath, {
                    ...state,
                    isPasswordUnlocked: true
                });

                this.stateManager.setState(targetDocumentPath, {
                    lastSyncTime: new Date(),
                    ideCodeSynced: true
                });
                this.stateManager.incrementPendingChanges(targetDocumentPath);
                debugLog('State updated successfully');
                
                // 检查并配置文件关联
                await this.checkAndConfigureFileAssociations(vbaFolder);
                
                // 检查调试提示
                this.checkDebugPrompt(this.stateManager.getState(targetDocumentPath)!);
            } finally {
                this.isSyncingIdeToDoc = false;
            }
        });
    }

    private async handleVbaAccessError(documentPath: string, error: string) {
        const fileName = path.basename(documentPath);
        
        // 检查是否是密码保护错误
        const isPasswordProtected = error.includes('该工程已被保护') || 
                                   error.includes('工程已被保护') ||
                                   error.toLowerCase().includes('project is protected');
        
        // 检查是否是文档访问异常错误（需要用户关闭文档）
        const isDocumentAccessError = error.includes('文档访问异常') || error.includes('请关闭已打开的');
        
        if (isPasswordProtected) {
            // VBA项目被保护的情况
            debugLog('VBA access denied, showing Office window for user to input password');
            await this.connectionManager.setWindowVisible(documentPath, true);
            
            // 提示用户输入密码
            const choice = await this.notificationManager.showInformationWithOptions(
                '文档【' + fileName + '】的VBA项目已被保护，请在打开的Office窗口中输入密码，然后点击"继续"',
                '继续', '取消'
            );
            
            if (choice === '取消') {
                throw new Error('用户取消了操作');
            }
            
            // 重新检查VBA访问权限
            const syncDirection = this.isSyncingDocToIde ? 'doc-to-ide' : 'ide-to-doc';
            const retryAccess = await this.connectionManager.checkVbaAccess(documentPath, syncDirection);
            if (!retryAccess.success) {
                throw new Error('VBA访问仍然失败，请检查密码是否正确');
            }
            
            // 保持窗口可见，不再设置为隐藏
            const appType = this.connectionManager.getAppType(documentPath);
            try {
                debugLog('Keeping Office window visible for document: ' + documentPath);
                await this.connectionManager.setWindowVisible(documentPath, true);
                debugLog('Office window kept visible successfully');
            } catch (e) {
                debugLog('Failed to set Office window visible:', e);
            }
        } else if (isDocumentAccessError) {
            // 文档访问异常，需要用户关闭文档
            debugLog('Document access error, showing notification to user');
            // 显示错误通知给用户
            this.notificationManager.showError(error);
            throw new Error(error);
        } else {
            // 其他错误，尝试重新打开文档
            debugLog('Office instance closed or other error, reopening document');
            // 重新打开文档
            const appType = this.connectionManager.getAppType(documentPath);
            const syncDirection = this.isSyncingDocToIde ? 'doc-to-ide' : 'ide-to-doc';
            await this.connectionManager.closeDocument(documentPath);
            await this.connectionManager.openDocument(documentPath, appType, false, syncDirection);
            
            // 重新获取状态
            let state = this.stateManager.getState(documentPath);
            if (!state) {
                state = this.stateManager.initState(documentPath);
            }
            state.isOfficeOpen = true;
            
            // 重新检查VBA访问权限
            const retryAccess = await this.connectionManager.checkVbaAccess(documentPath, syncDirection);
            if (!retryAccess.success) {
                // 再次失败时，完整处理错误（包括密码保护）
                await this.handleVbaAccessError(documentPath, retryAccess.error!);
            }
        }
    }

    async checkConflict(documentPath: string): Promise<boolean> {
        if (!this.conflictResolver) {
            return true;
        }
        
        return this.conflictResolver.resolveConflict(documentPath);
    }

    async enableSync(documentPath: string) {
        return ErrorHandler.wrapAsync(async () => {
            const workspaceRoot = this.configManager.getWorkspaceRoot();
            if (!workspaceRoot) {
                debugLog('No workspace folder found');
                this.notificationManager.showError('未找到工作区文件夹，请先打开一个工作区');
                throw new Error('No workspace folder found');
            }

            const appType = this.connectionManager.getAppType(documentPath);
            const fileName = path.basename(documentPath);
            const baseName = path.basename(fileName, path.extname(fileName));
            const ext = path.extname(fileName).replace('.', '');
            const vbaFolder = path.join(workspaceRoot, 'vba', baseName + '_' + ext);
            debugLog('File naming debug info: { documentPath: ' + documentPath + ', fileName: ' + fileName + ', baseName: ' + baseName + ', ext: ' + ext + ', vbaFolder: ' + vbaFolder + ' }');

            const docConfig = this.configManager.addDocument(documentPath, vbaFolder);
            debugLog('Added document to config: ' + documentPath);
            
            this.notificationManager.showInformation('正在打开文档【' + fileName + '】并导出VBA代码...');
            
            try {
                debugLog('Opening document: ' + documentPath + ' with appType: ' + appType);
                await this.connectionManager.openDocument(documentPath, appType);
                debugLog('Document opened successfully: ' + documentPath);
                
                // 首先检查VBA访问权限
                debugLog('Checking VBA access for document: ' + documentPath);
                const vbaAccess = await this.connectionManager.checkVbaAccess(documentPath, 'doc-to-ide');
                
                // 检查VBA访问权限的实际结果
                debugLog('VBA access check result: ' + (vbaAccess.success ? 'success' : 'error'));
                debugLog('VBA access error: ' + vbaAccess.error);
                
                // 如果VBA访问失败，根据错误类型处理
                if (!vbaAccess.success) {
                    await this.handleVbaAccessError(documentPath, vbaAccess.error);
                }

                // 导出所有模块到VBA文件夹
                if (!fs.existsSync(vbaFolder)) {
                    fs.mkdirSync(vbaFolder, { recursive: true });
                }
                
                // 导出所有模块到VBA文件夹并获取导出的模块列表
                const exportedModules = await this.connectionManager.exportAll(documentPath, vbaFolder);
                
                // 提取模块名称（去掉文件扩展名）
                const moduleNames = Array.isArray(exportedModules) ? exportedModules.map((moduleFile: string) => {
                    return path.basename(moduleFile, path.extname(moduleFile));
                }) : [];
                
                // 更新状态
                const now = new Date(Math.floor(Date.now() / 1000) * 1000);
                this.stateManager.setState(documentPath, {
                    isOfficeOpen: true,
                    isPasswordUnlocked: true,
                    lastSyncTime: now,
                    lastDocToIdeSyncTime: now,
                    ideCodeSynced: true
                });
                
                // 构建并显示通知消息
                const moduleList = moduleNames.join('、');
                this.notificationManager.showInformation('【' + fileName + '】首次同步成功，模块列表：【' + moduleList + '】');
                
            } catch (error: any) {
                debugLog('Error enabling sync:', error);
                this.notificationManager.showError('启用同步失败: ' + (error.message || '未知错误'));
                throw error;
            }
        });
    }

    async disableSync(documentPath: string) {
        return ErrorHandler.wrapAsync(async () => {
            debugLog('Disabling sync for document: ' + documentPath);
            
            try {
                // 关闭文档
                await this.connectionManager.closeDocument(documentPath);
                debugLog('Document closed successfully');
            } catch (error) {
                debugLog('Error closing document:', error);
            }
            
            // 清理状态
            this.stateManager.deleteState(documentPath);
            this.configManager.disableDocument(documentPath);
            
            const fileName = path.basename(documentPath);
            this.notificationManager.showInformation('已禁用文档【' + fileName + '】的同步');
        });
    }

    getSyncState(documentPath: string) {
        return this.stateManager.getState(documentPath);
    }

    getConfig() {
        return this.configManager.getConfig();
    }

    getConfigManager(): ConfigManager {
        return this.configManager;
    }

    getStateManager(): StateManager {
        return this.stateManager;
    }

    getConnectionManager(): ConnectionManager {
        return this.connectionManager;
    }

    async showAllWindows() {
        await this.connectionManager.showAllWindows();
    }

    private checkDebugPrompt(state: any) {
        const config = vscode.workspace.getConfiguration('vbaSync');
        const pendingChangesThreshold = config.get('debugPrompt.pendingChangesThreshold', 5);

        if (state.pendingChanges >= pendingChangesThreshold) {
            vscode.window.showInformationMessage(
                '已累积 ' + state.pendingChanges + ' 次变更，是否打开调试窗口？',
                '打开', '取消'
            ).then(choice => {
                if (choice === '打开') {
                    // 打开调试窗口
                    vscode.commands.executeCommand('vba-sync.openDebug');
                }
            });
        }
    }

    private async checkAndConfigureFileAssociations(vbaFolder: string) {
        try {
            const workspaceRoot = this.configManager.getWorkspaceRoot();
            if (!workspaceRoot) return;

            // 在工作区根目录（第一级）创建 .vscode 文件夹
            const vscodeFolder = path.join(workspaceRoot, '.vscode');
            const settingsFile = path.join(vscodeFolder, 'settings.json');

            // 检查 .vscode 文件夹是否存在
            if (!fs.existsSync(vscodeFolder)) {
                fs.mkdirSync(vscodeFolder, { recursive: true });
            }

            let settings: any = {};
            let needsUpdate = false;

            // 检查 settings.json 文件是否存在
            if (fs.existsSync(settingsFile)) {
                const content = fs.readFileSync(settingsFile, 'utf8');
                try {
                    settings = JSON.parse(content);
                } catch (error) {
                    debugLog('Failed to parse settings.json:', error);
                    settings = {};
                }
            }

            // 检查文件关联配置
            if (!settings['files.associations']) {
                settings['files.associations'] = {};
                needsUpdate = true;
            }

            // 检查 .bas 和 .cls 文件的关联
            if (settings['files.associations']['*.bas'] !== 'vb') {
                settings['files.associations']['*.bas'] = 'vb';
                needsUpdate = true;
            }

            if (settings['files.associations']['*.cls'] !== 'vb') {
                settings['files.associations']['*.cls'] = 'vb';
                needsUpdate = true;
            }

            // 如果需要更新配置
            if (needsUpdate) {
                fs.writeFileSync(settingsFile, JSON.stringify(settings, null, 2));
                debugLog('Updated settings.json with VBA file associations');
            } else {
                debugLog('VBA file associations already configured correctly');
            }
        } catch (error) {
            debugLog('Failed to check and configure file associations:', error);
        }
    }

    async detectAndConnectToOpenDocuments() {
        debugLog('Starting detectAndConnectToOpenDocuments');
        
        // 检查是否启用了文档自动检测
        const config = vscode.workspace.getConfiguration('vbaSync');
        const enableAutoDetection = config.get('enableDocumentAutoDetection', true);
        
        if (!enableAutoDetection) {
            debugLog('Document auto-detection is disabled, skipping detectAndConnectToOpenDocuments');
            return;
        }
        
        const configData = this.configManager.getConfig();
        if (!configData) {
            debugLog('No config available, skipping detectAndConnectToOpenDocuments');
            return;
        }

        const workspaceRoot = this.configManager.getWorkspaceRoot();
        if (!workspaceRoot) {
            debugLog('No workspace root available, skipping detectAndConnectToOpenDocuments');
            return;
        }

        const enabledDocs = configData.documents.filter(doc => doc.enabled);
        debugLog(`Found ${enabledDocs.length} enabled documents`);

        if (enabledDocs.length === 0) {
            debugLog('No enabled documents found, skipping detectAndConnectToOpenDocuments');
            return;
        }

        // 并行处理文档连接，提高效率
        const connectPromises = enabledDocs.map(async (docConfig) => {
            try {
                const fullDocumentPath = path.join(workspaceRoot, docConfig.path);
                debugLog(`Processing document: ${fullDocumentPath}`);

                const appType = this.connectionManager.getAppType(fullDocumentPath);
                debugLog(`Document appType: ${appType}`);

                // 只连接已打开的文档，不打开未打开的文档
                const isOpened = await this.connectionManager.openDocument(fullDocumentPath, appType, true);
                if (isOpened) {
                    debugLog(`Document already opened, connected successfully: ${fullDocumentPath}`);
                    this.stateManager.setState(fullDocumentPath, { isOfficeOpen: true });
                    debugLog(`State updated for document: ${fullDocumentPath}, isOfficeOpen=true`);
                } else {
                    debugLog(`Document not open, skipped: ${fullDocumentPath}`);
                }
            } catch (error: any) {
                debugLog(`Error processing document ${docConfig.path}:`, error);
            }
        });

        // 等待所有连接尝试完成
        await Promise.all(connectPromises);

        debugLog('detectAndConnectToOpenDocuments completed');
    }

    async checkModuleListConsistency(documentPath: string, vbaFolder: string): Promise<boolean> {
        debugLog('Checking module list consistency');
        
        // 创建临时文件夹
        const tempFolder = path.join(vbaFolder, '.vba-sync-tmp-check-consistency');
        debugLog('Temp folder for consistency check: ' + tempFolder);

        try {
            // 确保临时文件夹存在
            if (!fs.existsSync(tempFolder)) {
                fs.mkdirSync(tempFolder, { recursive: true });
                debugLog('Created temp folder: ' + tempFolder);
            }

            // 导出所有模块到临时文件夹
            debugLog('Exporting all modules to temp folder for consistency check');
            await this.connectionManager.exportAll(documentPath, tempFolder);
            debugLog('Export completed');

            // 获取两个文件夹中的文件列表（忽略临时文件）
            const getFilteredFiles = (dir: string): string[] => {
                if (!fs.existsSync(dir)) {
                    return [];
                }
                return fs.readdirSync(dir).filter(file => {
                    const filePath = path.join(dir, file);
                    if (fs.statSync(filePath).isDirectory()) {
                        return false;
                    }
                    return !FileComparer.isTempFile(file);
                });
            };

            const tempFiles = getFilteredFiles(tempFolder);
            const vbaFiles = getFilteredFiles(vbaFolder);

            debugLog('Files in temp folder: ' + tempFiles.join(', '));
            debugLog('Files in VBA folder: ' + vbaFiles.join(', '));

            // 比较文件名列表（不区分大小写排序）
            const sortFiles = (files: string[]): string[] => 
                files.map(f => f.toLowerCase()).sort();

            const sortedTempFiles = sortFiles(tempFiles);
            const sortedVbaFiles = sortFiles(vbaFiles);

            const areEqual = sortedTempFiles.length === sortedVbaFiles.length &&
                sortedTempFiles.every((file, index) => file === sortedVbaFiles[index]);

            debugLog('Module list consistency check result: ' + (areEqual ? 'consistent' : 'inconsistent'));
            return areEqual;

        } catch (error: any) {
            debugLog('Error checking module list consistency:', error);
            return false; // 安全起见，返回 false 表示不一致，继续同步
        } finally {
            // 清理临时文件夹
            if (fs.existsSync(tempFolder)) {
                debugLog('Cleaning up temp folder for consistency check');
                try {
                    fs.rmSync(tempFolder, { recursive: true, force: true });
                    debugLog('Temp folder cleaned up');
                } catch (cleanupError: any) {
                    debugLog('Error cleaning up temp folder:', cleanupError);
                }
            }
        }
    }

    async dispose() {
        for (const documentPath of this.stateManager.getAllDocumentPaths()) {
            const state = this.stateManager.getState(documentPath);
            if (state?.isOfficeOpen) {
                try {
                    await this.connectionManager.closeDocument(documentPath);
                } catch (error) {
                    console.error('Failed to close document:', error);
                }
            }
        }
    }
}