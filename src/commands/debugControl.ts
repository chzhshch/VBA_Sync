import * as vscode from 'vscode';
import { SyncEngine } from '../core/sync/SyncEngine';
import { StatusBar } from '../ui/statusBar';

export async function openDebug(syncEngine: SyncEngine, statusBar: StatusBar) {
    try {
        vscode.window.showInformationMessage('打开调试窗口功能已简化');
    } catch (error) {
        vscode.window.showErrorMessage(`打开调试窗口失败: ${(error as Error).message}`);
    }
}

export async function minimizeToBackground(syncEngine: SyncEngine, statusBar: StatusBar) {
    try {
        vscode.window.showInformationMessage('最小化到后台功能已简化');
    } catch (error) {
        vscode.window.showErrorMessage(`最小化到后台失败: ${(error as Error).message}`);
    }
}
