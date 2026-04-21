import * as vscode from 'vscode';
import * as path from 'path';
import { PythonClient } from './PythonClient';
import { OfficeAppType, EXT_TO_APP, COM_ERROR_CODES, VBAModule } from '../../types';

// 调试日志开关
const DEBUG = true;

// 统一日志函数
function debugLog(message: string, ...args: any[]) {
    if (DEBUG) {
        console.log('[VBA SYNC DEBUG] ' + message, ...args);
    }
}

function debugError(message: string, ...args: any[]) {
    if (DEBUG) {
        console.error('[VBA SYNC DEBUG ERROR] ' + message, ...args);
    }
}

export class ConnectionManager {
    private pythonClient: PythonClient;
    private connections: Map<string, any> = new Map();

    constructor(pythonClient: PythonClient) {
        this.pythonClient = pythonClient;
    }

    async cleanupConnections() {
        // 并行检查连接有效性，提高效率
        const checkPromises = Array.from(this.connections.entries()).map(async ([documentPath, connection]) => {
            try {
                // 尝试访问 VBAProject 来验证连接是否有效
                const result = await this.checkVbaAccess(documentPath);
                // 即使有密码，也算是有效连接
                if (!result.success && !result.error.includes('保护')) {
                    return documentPath;
                }
            } catch (error) {
                return documentPath;
            }
            return null;
        });
        
        const results = await Promise.all(checkPromises);
        const connectionsToRemove = results.filter((docPath): docPath is string => docPath !== null);
        
        for (const documentPath of connectionsToRemove) {
            this.connections.delete(documentPath);
        }
    }

    async detectAndConnectToOpenDocuments() {
        // 获取所有同步文档
        const config = await this.getConfig();
        if (!config || !Array.isArray((config as any).documents)) {
            return;
        }
        
        const documents = (config as any).documents;
        
        // 并行处理文档连接，提高效率
        const connectPromises = documents.map(async (doc: any) => {
            if (doc.enabled) {
                const documentPath = doc.path;
                // 检查文档是否已经有有效连接
                if (this.connections.has(documentPath)) {
                    return;
                }
                
                // 尝试通过 GetObject 连接已打开的文档
                try {
                    const success = await this.openDocument(documentPath, this.getAppType(documentPath), true);
                    if (success) {
                        this.connections.set(documentPath, true);
                    }
                } catch (error) {
                    // 忽略错误，继续处理其他文档
                }
            }
        });
        
        await Promise.all(connectPromises);
    }

    private async getConfig() {
        try {
            const response = await this.pythonClient.sendRequestWithRetry({ action: 'get_config', documentPath: '' });
            return response.data?.config || null;
        } catch (error) {
            return null;
        }
    }

    addConnection(documentPath: string) {
        this.connections.set(documentPath, true);
    }

    removeConnection(documentPath: string) {
        this.connections.delete(documentPath);
    }

    hasConnection(documentPath: string): boolean {
        return this.connections.has(documentPath);
    }

    async openDocument(documentPath: string, appType: OfficeAppType, onlyIfOpen = false, syncDirection?: string) {
        const response = await this.pythonClient.sendRequestWithRetry({
            action: 'open_document',
            documentPath,
            appType,
            visible: false,
            onlyIfOpen,
            syncDirection,
        });
        // 返回是否成功，而不是抛出错误
        return response.success && response.data?.success;
    }

    async closeDocument(documentPath: string) {
        await this.pythonClient.sendRequestWithRetry({
            action: 'close_document',
            documentPath,
        });
    }

    async checkVbaAccess(documentPath: string, syncDirection?: string) {
        const response = await this.pythonClient.sendRequestWithRetry({
            action: 'check_vba_access',
            documentPath,
            syncDirection,
        });
        
        if (!response) {
            throw new Error('VBA access check returned no response');
        }
        
        return {
            success: response.data?.success ?? false,
            error: String(response.data?.error ?? '')
        };
    }

    async setWindowVisible(documentPath: string, visible: boolean) {
        await this.pythonClient.sendRequestWithRetry({
            action: 'set_window_visible',
            documentPath,
            visible,
        });
    }

    async importModule(documentPath: string, filePath: string, moduleName?: string) {
        await this.pythonClient.sendRequestWithRetry({
            action: 'import_module',
            documentPath,
            filePath,
            moduleName,
        });
    }

    async exportAll(documentPath: string, outputDir: string) {
        const response = await this.pythonClient.sendRequestWithRetry({
            action: 'export_all',
            documentPath,
            outputDir,
        });
        return response.exported || [];
    }

    async removeModule(documentPath: string, moduleName: string) {
        await this.pythonClient.sendRequestWithRetry({
            action: 'delete_module',
            documentPath,
            moduleName,
        });
    }

    async clearModuleCode(documentPath: string, moduleName: string) {
        await this.pythonClient.sendRequestWithRetry({
            action: 'clear_module_code',
            documentPath,
            moduleName,
        });
    }

    async getLastSyncTime(documentPath: string) {
        const response = await this.pythonClient.sendRequestWithRetry({
            action: 'get_last_sync_time',
            documentPath
        });
        
        return response.data?.lastSyncTime;
    }

    async setLastSyncTime(documentPath: string, time: Date) {
        await this.pythonClient.sendRequestWithRetry({
            action: 'set_last_sync_time',
            documentPath,
            time: time.toISOString()
        });
    }

    getAppType(documentPath: string): OfficeAppType {
        const ext = path.extname(documentPath).toLowerCase();
        return EXT_TO_APP[ext] || 'excel';
    }

    async showAllWindows() {
        await this.pythonClient.sendRequest({ action: 'showAllWindows', documentPath: '' });
    }

    async listModules(documentPath: string): Promise<VBAModule[]> {
        const response = await this.pythonClient.sendRequestWithRetry({
            action: 'list_modules',
            documentPath,
        });
        return Array.isArray(response.data?.modules) ? response.data.modules : [];
    }

    async runMacro(documentPath: string, macroName: string) {
        const response = await this.pythonClient.sendRequestWithRetry({
            action: 'run_macro',
            documentPath,
            macroName,
        });
        return response.data || {};
    }

    async listMacros(documentPath: string) {
        const response = await this.pythonClient.sendRequestWithRetry({
            action: 'list_macros',
            documentPath,
        });
        return response.data || {};
    }
}