import * as vscode from 'vscode';
import { VbaSyncConfig } from '../../types';

// 调试日志开关
const DEBUG = true;

// 统一日志函数
function debugLog(message: string, ...args: any[]) {
    if (DEBUG) {
        console.log('[VBA SYNC DEBUG] ' + message, ...args);
    }
}

export class ConflictResolver {
    private settings: VbaSyncConfig['settings'];

    constructor(settings: VbaSyncConfig['settings']) {
        this.settings = settings;
    }

    async resolveConflict(documentPath: string): Promise<boolean> {
        // 这里可以实现更复杂的冲突检测逻辑
        // 目前简单返回 true，表示总是允许同步
        return true;
    }

    async checkConflict(documentPath: string, lastSyncTime: Date | null): Promise<boolean> {
        // 简单的冲突检测逻辑
        // 可以根据时间戳或其他方式判断是否有冲突
        return true;
    }
}
