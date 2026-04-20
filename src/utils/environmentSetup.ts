import * as vscode from 'vscode';
import { execSync, exec } from 'child_process';

export class EnvironmentSetup {
    static async ensureEnvironment() {
        const pythonReady = await this.checkPythonInstalled();
        const pywin32Ready = pythonReady ? await this.checkPywin32Installed() : false;
        
        if (pythonReady && !pywin32Ready) {
            await this.installPywin32();
        }

        return { pythonReady, pywin32Ready: pythonReady && await this.checkPywin32Installed() };
    }

    private static async checkPythonInstalled(): Promise<boolean> {
        try {
            execSync('python --version', { stdio: 'ignore' });
            return true;
        } catch (error) {
            const choice = await vscode.window.showInformationMessage(
                '未检测到 Python 安装，是否使用 winget 自动安装？',
                '使用 winget 安装',
                '手动安装',
                '取消'
            );

            if (choice === '使用 winget 安装') {
                const success = await this.installPythonViaWinget();
                return success;
            } else if (choice === '手动安装') {
                vscode.env.openExternal(vscode.Uri.parse('https://www.python.org/downloads/'));
                return false;
            } else {
                return false;
            }
        }
    }

    private static async installPythonViaWinget(): Promise<boolean> {
        return new Promise((resolve) => {
            exec('winget install --id Python.Python.3.12', (error, stdout, stderr) => {
                if (error) {
                    vscode.window.showErrorMessage(`安装 Python 失败: ${stderr}`);
                    resolve(false);
                } else {
                    vscode.window.showInformationMessage('Python 安装成功');
                    resolve(true);
                }
            });
        });
    }

    private static async checkPywin32Installed(): Promise<boolean> {
        try {
            execSync('python -c "import win32com.client"', { stdio: 'ignore' });
            return true;
        } catch (error) {
            return false;
        }
    }

    private static async installPywin32(): Promise<boolean> {
        return new Promise((resolve) => {
            const progressOptions: vscode.ProgressOptions = {
                location: vscode.ProgressLocation.Notification,
                title: '安装 pywin32',
                cancellable: false
            };

            vscode.window.withProgress(progressOptions, async (progress) => {
                progress.report({ increment: 0, message: '正在安装 pywin32...' });

                try {
                    execSync('pip install pywin32>=306', { stdio: 'ignore' });
                    progress.report({ increment: 100, message: '安装完成' });
                    vscode.window.showInformationMessage('pywin32 安装成功');
                    resolve(true);
                } catch (error) {
                    vscode.window.showErrorMessage('安装 pywin32 失败，请手动运行: pip install pywin32>=306');
                    resolve(false);
                }
            });
        });
    }
}
