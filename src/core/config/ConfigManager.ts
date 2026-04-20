import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import { z } from 'zod';
import { DocumentConfig, VbaSyncConfig } from '../../types';

// 调试日志开关
const DEBUG = true;

// 统一日志函数
function debugLog(message: string, ...args: any[]) {
    if (DEBUG) {
        console.log('[VBA SYNC DEBUG] ' + message, ...args);
    }
}

// Zod验证模式
const DocumentConfigSchema = z.object({
    path: z.string(),
    vbaFolder: z.string(),
    enabled: z.boolean(),
});

const VbaSyncConfigSchema = z.object({
    version: z.string(),
    documents: z.array(DocumentConfigSchema),
    settings: z.object({
        autoSync: z.boolean(),
        conflictStrategy: z.enum(['ask', 'ide-priority', 'document-priority']),
        showNotifications: z.boolean(),
    }),
});

export class ConfigManager {
    private config: VbaSyncConfig | null = null;

    constructor() {
        this.loadConfig();
    }

    private loadConfig() {
        const workspaceRoot = this.getWorkspaceRoot();
        if (!workspaceRoot) {
            return;
        }

        const configPath = path.join(workspaceRoot, '.vba-sync-config.json');
        if (fs.existsSync(configPath)) {
            try {
                const content = fs.readFileSync(configPath, 'utf8');
                const parsed = JSON.parse(content);
                
                // 使用Zod验证配置
                const validated = VbaSyncConfigSchema.safeParse(parsed);
                if (!validated.success) {
                    debugLog('Invalid config:', validated.error);
                    // 使用默认配置并通知用户
                    this.config = this.createDefaultConfig();
                    // 通知用户配置文件无效
                    vscode.window.showWarningMessage('VBA Sync 配置文件无效，已使用默认配置');
                } else {
                    this.config = validated.data;
                    debugLog('Config loaded and validated successfully');
                }
            } catch (error) {
                debugLog('Failed to load config:', error);
                this.config = this.createDefaultConfig();
                // 通知用户配置文件加载失败
                vscode.window.showWarningMessage('VBA Sync 配置文件加载失败，已使用默认配置');
            }
        } else {
            this.config = this.createDefaultConfig();
        }
    }

    private createDefaultConfig(): VbaSyncConfig {
        return {
            version: '2.0.0',
            documents: [],
            settings: {
                autoSync: false,
                conflictStrategy: 'ask',
                showNotifications: true
            }
        };
    }

    private saveConfig() {
        const workspaceRoot = this.getWorkspaceRoot();
        if (!workspaceRoot || !this.config) {
            return;
        }

        const configPath = path.join(workspaceRoot, '.vba-sync-config.json');
        try {
            fs.writeFileSync(configPath, JSON.stringify(this.config, null, 2));
            debugLog('Config saved successfully');
        } catch (error) {
            debugLog('Failed to save config:', error);
        }
    }

    getConfig(): VbaSyncConfig | null {
        return this.config;
    }

    getWorkspaceRoot(): string | null {
        if (vscode.workspace.workspaceFolders && vscode.workspace.workspaceFolders.length > 0) {
            return vscode.workspace.workspaceFolders[0].uri.fsPath;
        }
        return null;
    }

    addDocument(documentPath: string, vbaFolder: string): DocumentConfig {
        if (!this.config) {
            this.config = this.createDefaultConfig();
        }

        const docConfig: DocumentConfig = {
            path: path.relative(this.getWorkspaceRoot()!, documentPath),
            vbaFolder: path.relative(this.getWorkspaceRoot()!, vbaFolder),
            enabled: true
        };

        // 检查是否已存在
        const existingIndex = this.config.documents.findIndex(doc => doc.path === docConfig.path);
        if (existingIndex >= 0) {
            this.config.documents[existingIndex] = docConfig;
        } else {
            this.config.documents.push(docConfig);
        }

        this.saveConfig();
        return docConfig;
    }

    disableDocument(documentPath: string) {
        if (!this.config) {
            return;
        }

        const workspaceRoot = this.getWorkspaceRoot();
        if (!workspaceRoot) {
            return;
        }

        const relativePath = path.relative(workspaceRoot, documentPath);
        const docConfig = this.config.documents.find(doc => doc.path === relativePath);
        if (docConfig) {
            docConfig.enabled = false;
            this.saveConfig();
        }
    }

    getDocumentConfig(documentPath: string): DocumentConfig | null {
        if (!this.config) {
            return null;
        }

        const workspaceRoot = this.getWorkspaceRoot();
        if (!workspaceRoot) {
            return null;
        }

        const relativePath = path.relative(workspaceRoot, documentPath);
        return this.config.documents.find(doc => doc.path === relativePath) || null;
    }

    getDocumentConfigForFile(filePath: string): DocumentConfig | null {
        if (!this.config) {
            return null;
        }

        const workspaceRoot = this.getWorkspaceRoot();
        if (!workspaceRoot) {
            return null;
        }

        for (const doc of this.config.documents) {
            const vbaFolderPath = path.join(workspaceRoot, doc.vbaFolder);
            if (filePath.startsWith(vbaFolderPath)) {
                return doc;
            }
        }

        return null;
    }

    ensureFileInVbaFolder(filePath: string, vbaFolder: string): boolean {
        try {
            const relativePath = path.relative(vbaFolder, filePath);
            return !relativePath.startsWith('..');
        } catch (error) {
            return false;
        }
    }

    async checkAndConfigureFileAssociations(vbaFolder: string) {
        try {
            const workspaceRoot = this.getWorkspaceRoot();
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
}
