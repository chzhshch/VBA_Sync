import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import { SyncEngine } from '../core/sync/SyncEngine';
import { FileWatcher } from '../core/fileWatcher';
import { StatusBar } from '../ui/statusBar';

export async function enableSync(
    syncEngine: SyncEngine,
    fileWatcher: FileWatcher,
    statusBar: StatusBar,
    uri?: vscode.Uri
) {
    console.log('VBA Sync: 开始启用同步...');
    try {
        console.log('VBA Sync: 检查 syncEngine:', !!syncEngine);
        if (!syncEngine) {
            console.error('VBA Sync: syncEngine 未初始化');
            vscode.window.showErrorMessage('VBA Sync 未初始化，请检查 Python 环境');
            return;
        }

        let documentPath: string;

        if (uri) {
            // 从右键菜单触发
            documentPath = uri.fsPath;
            console.log('VBA Sync: 从右键菜单获取文件路径:', documentPath);
        } else {
            // 从命令面板触发，选择文件
            console.log('VBA Sync: 从命令面板选择文件');
            const fileUris = await vscode.window.showOpenDialog({
                filters: {
                    'Office 宏文件': ['xlsm', 'xlsb', 'xltm', 'pptm', 'potm', 'ppsm', 'docm', 'dotm'],
                    'All Files': ['*']
                },
                canSelectFiles: true,
                canSelectFolders: false,
                canSelectMany: false
            });

            if (!fileUris || fileUris.length === 0) {
                console.log('VBA Sync: 用户取消选择文件');
                return;
            }

            documentPath = fileUris[0].fsPath;
            console.log('VBA Sync: 从命令面板获取文件路径:', documentPath);
        }

        // 检查文件是否存在
        console.log('VBA Sync: 检查文件是否存在:', documentPath);
        if (!fs.existsSync(documentPath)) {
            console.error('VBA Sync: 文件不存在:', documentPath);
            vscode.window.showErrorMessage('指定的文件不存在');
            return;
        }

        // 显示同步中状态
        console.log('VBA Sync: 显示同步中状态');
        statusBar.showSyncing();

        vscode.window.showInformationMessage('正在准备启用VBA同步...');

        // 启用同步
        console.log('VBA Sync: 调用 syncEngine.enableSync');
        await syncEngine.enableSync(documentPath);
        console.log('VBA Sync: enableSync 完成');

        // 刷新文件监控
        console.log('VBA Sync: 刷新文件监控');
        fileWatcher.refresh();

        // 更新状态栏
        console.log('VBA Sync: 更新状态栏');
        const documentName = path.basename(documentPath);
        statusBar.updateStatus(undefined, documentName);

        console.log('VBA Sync: 同步启用成功');
    } catch (error) {
        const errorMessage = (error as Error).message || '未知错误';
        console.error('VBA Sync: 启用同步失败:', error);
        statusBar.showError(errorMessage);
        vscode.window.showErrorMessage(`启用同步失败: ${errorMessage}`);
    }
}
