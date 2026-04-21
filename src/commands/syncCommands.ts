import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import { SyncEngine } from '../core/sync/SyncEngine';

export async function syncAll(syncEngine: SyncEngine) {
    try {
        vscode.window.showInformationMessage('同步所有文档功能已简化');
    } catch (error) {
        vscode.window.showErrorMessage(`同步失败: ${(error as Error).message}`);
    }
}

export async function exportAll(syncEngine: SyncEngine) {
    try {
        vscode.window.showInformationMessage('导出所有模块功能已简化');
    } catch (error) {
        vscode.window.showErrorMessage(`导出失败: ${(error as Error).message}`);
    }
}

export async function viewStatus(syncEngine: SyncEngine) {
    try {
        const config = syncEngine.getConfig();
        if (!config || config.documents.length === 0) {
            vscode.window.showInformationMessage('没有启用同步的文档');
            return;
        }

        const workspaceRoot = vscode.workspace.workspaceFolders?.[0].uri.fsPath;
        if (!workspaceRoot) return;

        let statusText = 'VBA Sync 状态:\n\n';

        for (const doc of config.documents) {
            const absPath = path.join(workspaceRoot, doc.path);
            const state = syncEngine.getSyncState(absPath);

            statusText += `文档: ${doc.path}\n`;
            statusText += `状态: ${doc.enabled ? '启用' : '禁用'}\n`;
            statusText += `VBA 文件夹: ${doc.vbaFolder}\n`;

            if (state) {
                statusText += `Office 进程: ${state.isOfficeOpen ? '打开' : '关闭'}\n`;
                statusText += `窗口可见: ${state.isWindowVisible ? '是' : '否'}\n`;
                statusText += `待调试变更: ${state.pendingChanges}\n`;
                statusText += `上次同步: ${state.lastSyncTime?.toLocaleString() || '从未'}\n`;
            }

            statusText += '\n';
        }

        vscode.window.showInformationMessage(statusText, { modal: true });
    } catch (error) {
        vscode.window.showErrorMessage(`查看状态失败: ${(error as Error).message}`);
    }
}

export async function resetPendingChanges(syncEngine: SyncEngine) {
    try {
        vscode.window.showInformationMessage('重置同步计数功能已简化');
    } catch (error) {
        vscode.window.showErrorMessage(`重置同步计数失败: ${(error as Error).message}`);
    }
}

export async function openConfig() {
    try {
        if (!vscode.workspace.workspaceFolders) {
            vscode.window.showInformationMessage('没有打开的工作区');
            return;
        }

        const workspaceRoot = vscode.workspace.workspaceFolders[0].uri.fsPath;
        const configPath = path.join(workspaceRoot, '.vba-sync.json');

        if (!fs.existsSync(configPath)) {
            vscode.window.showInformationMessage('配置文件不存在');
            return;
        }

        const docUri = vscode.Uri.file(configPath);
        await vscode.window.showTextDocument(docUri);
    } catch (error) {
        vscode.window.showErrorMessage(`打开配置文件失败: ${(error as Error).message}`);
    }
}

export async function syncSingleFile(syncEngine: SyncEngine, uri?: vscode.Uri) {
    try {
        if (!uri) {
            vscode.window.showErrorMessage('请选择要同步的文件');
            return;
        }

        await syncEngine.syncFileToDocument(uri.fsPath);
    } catch (error) {
        vscode.window.showErrorMessage(`同步文件失败: ${(error as Error).message}`);
    }
}
