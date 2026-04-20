import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import { SyncEngine } from '../core/sync/SyncEngine';

export async function newModule(syncEngine: SyncEngine) {
    try {
        // 暂时简化实现
        vscode.window.showInformationMessage('创建新模块功能已简化');
    } catch (error) {
        vscode.window.showErrorMessage(`创建模块失败: ${(error as Error).message}`);
    }
}
