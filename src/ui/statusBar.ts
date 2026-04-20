import * as vscode from 'vscode';
import { SyncEngine } from '../core/sync/SyncEngine';
import { DocumentSyncState } from '../types';

export class StatusBar {
    private statusBarItem: vscode.StatusBarItem;
    private syncEngine: SyncEngine | null;

    constructor(syncEngine: SyncEngine | null) {
        this.syncEngine = syncEngine;
        this.statusBarItem = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Left);
        this.statusBarItem.command = 'vba-sync.viewStatus';
        this.statusBarItem.show();
        this.updateStatus();
    }

    updateStatus(state?: DocumentSyncState, documentName?: string) {
        if (!state) {
            // 未连接状态
            this.statusBarItem.text = `$(circle-slash) VBA Sync`;
            this.statusBarItem.color = undefined;
            this.statusBarItem.tooltip = 'VBA Sync 未连接';
            return;
        }

        if (state.isDebugging) {
            // 调试中状态
            this.statusBarItem.text = `$(debug-alt) VBA Sync: ${documentName} 调试中`;
            this.statusBarItem.color = new vscode.ThemeColor('statusBar.debuggingBackground');
            this.statusBarItem.tooltip = `调试中\n上次同步: ${state.lastSyncTime?.toLocaleString()}`;
        } else if (state.pendingChanges > 0) {
            // 待调试状态
            this.statusBarItem.text = `$(warning) VBA Sync: ${documentName} (${state.pendingChanges}次变更待调试)`;
            this.statusBarItem.color = new vscode.ThemeColor('statusBar.warningBackground');
            this.statusBarItem.tooltip = `有 ${state.pendingChanges} 次变更待调试\n上次同步: ${state.lastSyncTime?.toLocaleString()}`;
        } else if (state.isOfficeOpen) {
            // 后台运行状态
            this.statusBarItem.text = `$(check) VBA Sync: ${documentName}`;
            this.statusBarItem.color = undefined;
            this.statusBarItem.tooltip = `后台运行中\n上次同步: ${state.lastSyncTime?.toLocaleString()}`;
        } else {
            // 未连接状态
            this.statusBarItem.text = `$(circle-slash) VBA Sync`;
            this.statusBarItem.color = undefined;
            this.statusBarItem.tooltip = 'VBA Sync 未连接';
        }
    }

    showSyncing() {
        this.statusBarItem.text = `$(loading~spin) VBA Sync: 同步中...`;
        this.statusBarItem.color = new vscode.ThemeColor('statusBar.noFolderBackground');
        this.statusBarItem.tooltip = '正在同步...';
    }

    showError(message: string) {
        this.statusBarItem.text = `$(error) VBA Sync: 同步失败`;
        this.statusBarItem.color = new vscode.ThemeColor('errorForeground');
        this.statusBarItem.tooltip = `同步失败: ${message}`;
    }

    dispose() {
        this.statusBarItem.dispose();
    }
}
