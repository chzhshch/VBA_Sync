import * as vscode from 'vscode';
import { SyncDetails } from '../../types';

// 调试日志开关
const DEBUG = true;

// 统一日志函数
function debugLog(message: string, ...args: any[]) {
    if (DEBUG) {
        console.log('[VBA SYNC DEBUG] ' + message, ...args);
    }
}

export class NotificationManager {
    /**
     * 显示信息通知
     * @param message 通知消息
     */
    showInformation(message: string) {
        debugLog('Showing information: ' + message);
        vscode.window.showInformationMessage(message);
    }

    /**
     * 显示错误通知
     * @param message 错误消息
     */
    showError(message: string) {
        debugLog('Showing error: ' + message);
        vscode.window.showErrorMessage(message);
    }

    /**
     * 显示警告通知
     * @param message 警告消息
     */
    showWarning(message: string) {
        debugLog('Showing warning: ' + message);
        vscode.window.showWarningMessage(message);
    }

    /**
     * 显示带有选项的通知
     * @param message 通知消息
     * @param options 选项数组
     * @returns 用户选择的选项
     */
    async showInformationWithOptions(message: string, ...options: string[]): Promise<string | undefined> {
        debugLog('Showing information with options: ' + message, options);
        return vscode.window.showInformationMessage(message, ...options);
    }

    /**
     * 显示同步成功通知
     * @param documentName 文档名称
     * @param modules 同步的模块列表
     * @param direction 同步方向 (docToIde | ideToDoc)
     */
    showSyncSuccess(documentName: string, modules: string[], direction?: 'docToIde' | 'ideToDoc') {
        const moduleList = modules.join('、');
        let message = `文档【${documentName}】的模块【${moduleList}】同步成功`;
        
        if (direction) {
            const directionText = direction === 'docToIde' ? '从文档同步到IDE' : '从IDE同步到文档';
            message = `【${documentName}】${directionText}成功，同步的模块：【${moduleList}】`;
        }
        
        this.showInformation(message);
    }

    /**
     * 显示同步失败通知
     * @param documentName 文档名称
     * @param error 错误信息
     */
    showSyncError(documentName: string, error: string) {
        const message = `同步文档【${documentName}】失败: ${error}`;
        this.showError(message);
    }

    /**
     * 显示同步开始通知
     * @param documentName 文档名称
     * @param direction 同步方向 (docToIde | ideToDoc)
     */
    showSyncStart(documentName: string, direction: 'docToIde' | 'ideToDoc') {
        const directionText = direction === 'docToIde' ? '从文档同步到IDE' : '从IDE同步到文档';
        const message = `正在${directionText}【${documentName}】...`;
        this.showInformation(message);
    }

    /**
     * 显示同步失败通知
     * @param documentName 文档名称
     * @param direction 同步方向 (docToIde | ideToDoc)
     * @param error 错误信息
     */
    showSyncFailure(documentName: string, direction: 'docToIde' | 'ideToDoc', error: string) {
        const directionText = direction === 'docToIde' ? '从文档同步到IDE' : '从IDE同步到文档';
        const message = `同步【${documentName}】${directionText}失败: ${error}`;
        this.showError(message);
    }

    /**
     * 显示同步未执行通知
     * @param documentName 文档名称
     * @param direction 同步方向 (docToIde | ideToDoc)
     * @param reason 未执行原因
     */
    showSyncNotExecuted(documentName: string, direction: 'docToIde' | 'ideToDoc', reason: string) {
        const directionText = direction === 'docToIde' ? '从文档同步到IDE' : '从IDE同步到文档';
        const message = `【${documentName}】${directionText}未执行：${reason}`;
        this.showInformation(message);
    }

    /**
     * 显示详细的同步成功通知
     * @param documentName 文档名称
     * @param syncDetails 同步详细信息
     * @param direction 同步发起方 (文档 | IDE)
     */
    showSyncSuccessDetails(documentName: string, syncDetails: SyncDetails, direction: '文档' | 'IDE') {
        const parts: string[] = [];
        
        if (syncDetails.addedModules.length > 0) {
            parts.push(`增【${syncDetails.addedModules.join('、')}】`);
        }
        if (syncDetails.deletedModules.length > 0) {
            parts.push(`删【${syncDetails.deletedModules.join('、')}】`);
        }
        if (syncDetails.clearedModules.length > 0) {
            parts.push(`清【${syncDetails.clearedModules.join('、')}】`);
        }
        if (syncDetails.modifiedModules.length > 0) {
            parts.push(`改【${syncDetails.modifiedModules.join('、')}】`);
        }
        
        const message = `【${documentName}】自${direction}同步完成：${parts.join('，')}`;
        debugLog('Showing detailed sync success: ' + message);
        this.showInformation(message);
    }
}