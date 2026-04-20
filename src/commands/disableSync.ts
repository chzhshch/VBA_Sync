import * as vscode from 'vscode';
import * as path from 'path';
import { SyncEngine } from '../core/sync/SyncEngine';
import { FileWatcher } from '../core/fileWatcher';
import { StatusBar } from '../ui/statusBar';

export async function disableSync(
    syncEngine: SyncEngine,
    fileWatcher: FileWatcher,
    statusBar: StatusBar
) {
    try {
        const workspaceRoot = vscode.workspace.workspaceFolders?.[0].uri.fsPath;
        if (!workspaceRoot) return;

        // 显示所有隐藏的 Office 窗口
        console.log('VBA Sync: 显示所有隐藏的 Office 窗口');
        await syncEngine.showAllWindows();
        
        // 等待一小段时间，确保窗口显示完成
        await new Promise(resolve => setTimeout(resolve, 500));

        // 刷新文件监控
        fileWatcher.refresh();

        // 更新状态栏
        statusBar.updateStatus();

        vscode.window.showInformationMessage('VBA 同步已禁用');
        console.log('VBA Sync: 同步禁用成功');
    } catch (error) {
        vscode.window.showErrorMessage(`禁用同步失败: ${(error as Error).message}`);
    }
}
