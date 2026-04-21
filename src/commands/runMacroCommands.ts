import * as vscode from 'vscode';
import * as path from 'path';
import { SyncEngine } from '../core/sync/SyncEngine';

export async function runMacro(syncEngine: SyncEngine) {
    try {
        let documentPath: string | null = null;
        let allMacros: Array<{ label: string; description: string; documentPath: string; macroName: string }> = [];

        // 获取当前打开的文档
        const activeDocument = vscode.window.activeTextEditor?.document;
        if (activeDocument) {
            // 获取文档路径
            const vbaFilePath = activeDocument.fileName;
            const workspaceRoot = vscode.workspace.workspaceFolders?.[0].uri.fsPath;
            if (workspaceRoot) {
                // 查找对应的 Office 文档
                documentPath = syncEngine.getDocumentPathForVbaFile(vbaFilePath);
                if (documentPath) {
                    // 检查代码一致性
                    const isConsistent = await syncEngine.checkCodeConsistency(documentPath);
                    if (!isConsistent) {
                        const choice = await vscode.window.showInformationMessage(
                            '代码不一致，需要先同步。是否立即同步？',
                            '同步',
                            '取消'
                        );
                        
                        if (choice === '同步') {
                            await syncEngine.syncDocumentToIde(documentPath);
                            await syncEngine.syncFilesToDocument([], documentPath);
                        } else {
                            return;
                        }
                    }

                    // 列出可用的宏
                    const macrosResult = await syncEngine.getConnectionManager().listMacros(documentPath!);
                    if (macrosResult.success) {
                        const macros = macrosResult.macros as string[] || [];
                        const documentName = path.basename(documentPath!);
                        allMacros = macros.map((macro: string) => ({
                            label: `${documentName} - ${macro}`,
                            description: `运行 ${macro}`,
                            documentPath: documentPath!,
                            macroName: macro
                        }));
                    }
                }
            }
        }

        // 如果没有从活动文档获取到宏，获取所有已连接的文档的宏
        if (allMacros.length === 0) {
            const documentPaths = syncEngine.getStateManager().getAllDocumentPaths();
            if (documentPaths.length === 0) {
                vscode.window.showErrorMessage('没有可用的文档');
                return;
            }

            // 为每个文档列出可用的宏
            for (const docPath of documentPaths) {
                const macrosResult = await syncEngine.getConnectionManager().listMacros(docPath);
                if (macrosResult.success) {
                    const macros = macrosResult.macros as string[] || [];
                    const documentName = path.basename(docPath);
                    const docMacros = macros.map((macro: string) => ({
                        label: `${documentName} - ${macro}`,
                        description: `运行 ${macro}`,
                        documentPath: docPath,
                        macroName: macro
                    }));
                    allMacros = allMacros.concat(docMacros);
                }
            }
        }

        if (allMacros.length === 0) {
            vscode.window.showInformationMessage('没有可用的宏');
            return;
        }

        // 让用户选择宏
        const selectedItem = await vscode.window.showQuickPick(allMacros, {
            placeHolder: '选择要运行的宏'
        });

        if (!selectedItem) {
            return;
        }

        // 运行宏
        const result = await syncEngine.getConnectionManager().runMacro(selectedItem.documentPath, selectedItem.macroName);
        if (result.success) {
            vscode.window.showInformationMessage(`宏 ${selectedItem.macroName} 执行成功`);
        } else {
            vscode.window.showErrorMessage(`运行宏失败: ${result.error || '未知错误'}`);
        }
    } catch (error) {
        vscode.window.showErrorMessage(`运行宏失败: ${(error as Error).message}`);
    }
}

export async function listMacros(syncEngine: SyncEngine) {
    try {
        let allMacros: string[] = [];

        // 获取当前打开的文档
        const activeDocument = vscode.window.activeTextEditor?.document;
        if (activeDocument) {
            // 获取文档路径
            const vbaFilePath = activeDocument.fileName;
            const workspaceRoot = vscode.workspace.workspaceFolders?.[0].uri.fsPath;
            if (workspaceRoot) {
                // 查找对应的 Office 文档
                const documentPath = syncEngine.getDocumentPathForVbaFile(vbaFilePath);
                if (documentPath) {
                    // 列出可用的宏
                    const result = await syncEngine.getConnectionManager().listMacros(documentPath);
                    if (result.success) {
                        const macros = result.macros as string[] || [];
                        const documentName = path.basename(documentPath);
                        const docMacros = macros.map((macro: string) => `${documentName} - ${macro}`);
                        allMacros = allMacros.concat(docMacros);
                    }
                }
            }
        }

        // 如果没有从活动文档获取到宏，获取所有已连接的文档的宏
        if (allMacros.length === 0) {
            const documentPaths = syncEngine.getStateManager().getAllDocumentPaths();
            if (documentPaths.length === 0) {
                vscode.window.showErrorMessage('没有可用的文档');
                return;
            }

            // 为每个文档列出可用的宏
            for (const docPath of documentPaths) {
                const result = await syncEngine.getConnectionManager().listMacros(docPath);
                if (result.success) {
                    const macros = result.macros as string[] || [];
                    const documentName = path.basename(docPath);
                    const docMacros = macros.map((macro: string) => `${documentName} - ${macro}`);
                    allMacros = allMacros.concat(docMacros);
                }
            }
        }

        if (allMacros.length === 0) {
            vscode.window.showInformationMessage('没有可用的宏');
            return;
        }

        // 显示宏列表
        let macroList = '可用的宏:\n\n';
        allMacros.forEach((macro: string) => {
            macroList += `- ${macro}\n`;
        });

        vscode.window.showInformationMessage(macroList, { modal: true });
    } catch (error) {
        vscode.window.showErrorMessage(`获取宏列表失败: ${(error as Error).message}`);
    }
}
