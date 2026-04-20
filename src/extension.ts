import * as vscode from 'vscode';
import { PythonClient } from './core/connection/PythonClient';
import { ConnectionManager } from './core/connection/ConnectionManager';
import { SyncEngine } from './core/sync/SyncEngine';
import { FileWatcher } from './core/fileWatcher';
import { StatusBar } from './ui/statusBar';
import { EnvironmentSetup } from './utils/environmentSetup';
import { ConfigManager } from './core/config/ConfigManager';
import { StateManager } from './core/state/StateManager';
import { NotificationManager } from './core/notification/NotificationManager';
import { enableSync } from './commands/enableSync';
import { disableSync } from './commands/disableSync';
import { newModule } from './commands/newModule';
import { openDebug, minimizeToBackground } from './commands/debugControl';
import { syncAll, exportAll, viewStatus, resetPendingChanges, openConfig, syncSingleFile } from './commands/syncCommands';

export class VbaSyncExtension {
    private pythonClient: PythonClient | null = null;
    private connectionManager: ConnectionManager | null = null;
    private syncEngine: SyncEngine | null = null;
    private fileWatcher: FileWatcher | null = null;
    private statusBar: StatusBar | null = null;
    private configManager: ConfigManager | null = null;
    private stateManager: StateManager | null = null;
    private notificationManager: NotificationManager | null = null;
    private disposables: vscode.Disposable[] = [];

    async activate(context: vscode.ExtensionContext) {
        console.log('VBA Sync: 开始激活插件...');
        try {
            // 初始化状态栏
            console.log('VBA Sync: 初始化状态栏');
            this.statusBar = new StatusBar(null as any);
            this.statusBar.updateStatus();

            // 环境检测
            console.log('VBA Sync: 开始环境检测');
            const envResult = await EnvironmentSetup.ensureEnvironment();
            console.log('VBA Sync: 环境检测结果:', envResult);

            if (!envResult.pythonReady) {
                console.log('VBA Sync: Python 环境未就绪');
                vscode.window.showWarningMessage('Python 环境未就绪，VBA Sync 功能将受限');
                // 仍然注册命令，只是功能受限
                console.log('VBA Sync: 注册命令');
                this.registerCommands(context);
                return;
            }

            // 初始化组件
            console.log('VBA Sync: 初始化 PythonClient');
            this.pythonClient = new PythonClient();
            
            console.log('VBA Sync: 初始化 ConnectionManager');
            this.connectionManager = new ConnectionManager(this.pythonClient);
            
            console.log('VBA Sync: 初始化 ConfigManager');
            this.configManager = new ConfigManager();
            
            console.log('VBA Sync: 初始化 StateManager');
            this.stateManager = new StateManager();
            
            console.log('VBA Sync: 初始化 NotificationManager');
            this.notificationManager = new NotificationManager();
            
            console.log('VBA Sync: 初始化 SyncEngine');
            this.syncEngine = new SyncEngine(
                this.connectionManager,
                this.configManager,
                this.stateManager,
                this.notificationManager
            );
            
            console.log('VBA Sync: 初始化 FileWatcher');
            this.fileWatcher = new FileWatcher(this.syncEngine);
            
            // 更新状态栏
            console.log('VBA Sync: 更新状态栏');
            this.statusBar.dispose();
            this.statusBar = new StatusBar(this.syncEngine);

            // 注册命令
            console.log('VBA Sync: 注册命令');
            this.registerCommands(context);

            // 清理无效连接
            console.log('VBA Sync: 清理无效连接');
            try {
                await this.connectionManager?.cleanupConnections();
                console.log('VBA Sync: 连接清理完成');
            } catch (error) {
                console.error('VBA Sync: 连接清理失败:', error);
            }

            // 检测并连接已打开的文档
            console.log('VBA Sync: 检测并连接已打开的文档');
            try {
                await this.connectionManager?.detectAndConnectToOpenDocuments();
                console.log('VBA Sync: 已成功连接已打开的文档');
            } catch (error) {
                console.error('VBA Sync: 检测并连接已打开的文档失败:', error);
            }

            console.log('VBA Sync: 激活完成');
            vscode.window.showInformationMessage('VBA Sync 已激活');
        } catch (error) {
            console.error('VBA Sync: 激活失败:', error);
            vscode.window.showErrorMessage(`VBA Sync 激活失败: ${error}`);
        }
    }

    private registerCommands(context: vscode.ExtensionContext) {
        const commands = [
            vscode.commands.registerCommand('vba-sync.enableSync', (uri) => 
                enableSync(this.syncEngine!, this.fileWatcher!, this.statusBar!, uri)
            ),
            vscode.commands.registerCommand('vba-sync.disableSync', () => 
                disableSync(this.syncEngine!, this.fileWatcher!, this.statusBar!)
            ),
            vscode.commands.registerCommand('vba-sync.newModule', () => 
                newModule(this.syncEngine!)
            ),
            vscode.commands.registerCommand('vba-sync.syncAll', () => 
                syncAll(this.syncEngine!)
            ),
            vscode.commands.registerCommand('vba-sync.openDebug', () => 
                openDebug(this.syncEngine!, this.statusBar!)
            ),
            vscode.commands.registerCommand('vba-sync.minimizeToBackground', () => 
                minimizeToBackground(this.syncEngine!, this.statusBar!)
            ),
            vscode.commands.registerCommand('vba-sync.exportAll', () => 
                exportAll(this.syncEngine!)
            ),
            vscode.commands.registerCommand('vba-sync.viewStatus', () => 
                viewStatus(this.syncEngine!)
            ),
            vscode.commands.registerCommand('vba-sync.resetPendingChanges', () => 
                resetPendingChanges(this.syncEngine!)
            ),
            vscode.commands.registerCommand('vba-sync.openConfig', () => 
                openConfig()
            ),
            vscode.commands.registerCommand('vba-sync.syncSingleFile', (uri) => 
                syncSingleFile(this.syncEngine!, uri)
            ),
        ];

        commands.forEach(cmd => {
            this.disposables.push(cmd);
            context.subscriptions.push(cmd);
        });
    }

    async deactivate() {
        // 清理资源
        if (this.fileWatcher) {
            this.fileWatcher.dispose();
        }

        if (this.connectionManager) {
            // 显示所有隐藏的 Office 窗口
            await this.connectionManager.showAllWindows();
        }

        if (this.pythonClient) {
            await this.pythonClient.dispose();
        }

        if (this.statusBar) {
            this.statusBar.dispose();
        }

        this.disposables.forEach(d => d.dispose());
    }
}

let extensionInstance: VbaSyncExtension | null = null;

export function activate(context: vscode.ExtensionContext) {
    extensionInstance = new VbaSyncExtension();
    context.subscriptions.push({
        dispose: () => {
            if (extensionInstance) {
                return extensionInstance.deactivate();
            }
        }
    });
    return extensionInstance.activate(context);
}

export async function deactivate() {
    if (extensionInstance) {
        await extensionInstance.deactivate();
    }
}
