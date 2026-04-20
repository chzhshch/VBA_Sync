import * as fs from 'fs';
import * as path from 'path';

// 调试日志开关
const DEBUG = true;

// 统一日志函数
function debugLog(message: string, ...args: any[]) {
    if (DEBUG) {
        console.log('[VBA SYNC DEBUG] ' + message, ...args);
    }
}

export interface FileDifference {
    fileName: string;
    type: 'case-only' | 'substantial' | 'new' | 'deleted';
}

export class FileComparer {
    /**
     * 比较两个文件夹中的文件
     * @param sourceDir 源文件夹
     * @param targetDir 目标文件夹
     * @param includeTemp 是否包含临时文件
     * @returns 不同的文件列表
     */
    static compareFiles(sourceDir: string, targetDir: string, includeTemp: boolean = false): string[] {
        const changedFiles: string[] = [];

        // 读取源文件夹中的所有文件
        if (!fs.existsSync(sourceDir)) {
            debugLog('Source directory does not exist: ' + sourceDir);
            return changedFiles;
        }

        const files = fs.readdirSync(sourceDir);

        for (const file of files) {
            const sourcePath = path.join(sourceDir, file);
            const targetPath = path.join(targetDir, file);

            // 跳过临时文件
            if (!includeTemp && FileComparer.isTempFile(file)) {
                debugLog('Skipping temp file: ' + file);
                continue;
            }

            // 跳过文件夹
            if (fs.statSync(sourcePath).isDirectory()) {
                continue;
            }

            // 检查文件是否存在于目标文件夹
            if (!fs.existsSync(targetPath)) {
                changedFiles.push(file);
                debugLog('File not in target: ' + file);
                continue;
            }

            // 比较文件内容
            const sourceContent = fs.readFileSync(sourcePath).toString();
            const targetContent = fs.readFileSync(targetPath).toString();

            if (sourceContent !== targetContent) {
                changedFiles.push(file);
                debugLog('File content different: ' + file);
            }
        }

        return changedFiles;
    }

    /**
     * 比较两个文件的差异，使用多层级比较策略
     * @param sourceFile 源文件路径
     * @param targetFile 目标文件路径
     * @returns 差异类型
     */
    static compareFilesWithMultiLevel(sourceFile: string, targetFile: string): 'identical' | 'case-only' | 'substantial' {
        try {
            // 1. 比较文件大小
            const sourceStats = fs.statSync(sourceFile);
            const targetStats = fs.statSync(targetFile);
            
            if (sourceStats.size !== targetStats.size) {
                debugLog('File size different: ' + path.basename(sourceFile));
                return 'substantial';
            }
            
            // 2. 读取文件内容
            const sourceContent = fs.readFileSync(sourceFile).toString();
            const targetContent = fs.readFileSync(targetFile).toString();
            
            // 3. 完全相同
            if (sourceContent === targetContent) {
                return 'identical';
            }
            
            // 4. 仅大小写差异
            if (sourceContent.toLowerCase() === targetContent.toLowerCase()) {
                debugLog('File has only case differences: ' + path.basename(sourceFile));
                return 'case-only';
            }
            
            // 5. 实质性差异
            debugLog('File has substantial differences: ' + path.basename(sourceFile));
            return 'substantial';
        } catch (error) {
            debugLog('Error comparing files:', error);
            return 'substantial';
        }
    }

    /**
     * 比较两个文件夹中的文件，返回详细的差异类型
     * @param sourceDir 源文件夹（文档导出）
     * @param targetDir 目标文件夹（IDE）
     * @param includeTemp 是否包含临时文件
     * @returns 差异文件列表，包含文件名和差异类型
     */
    static compareFilesWithDetail(sourceDir: string, targetDir: string, includeTemp: boolean = false): FileDifference[] {
        const differences: FileDifference[] = [];

        // 读取源文件夹中的所有文件
        if (!fs.existsSync(sourceDir)) {
            debugLog('Source directory does not exist: ' + sourceDir);
            return differences;
        }

        const sourceFiles = fs.readdirSync(sourceDir);
        const targetFiles = fs.existsSync(targetDir) ? fs.readdirSync(targetDir) : [];

        // 检查源文件夹中的文件
        for (const file of sourceFiles) {
            const sourcePath = path.join(sourceDir, file);
            const targetPath = path.join(targetDir, file);

            // 跳过临时文件
            if (!includeTemp && FileComparer.isTempFile(file)) {
                debugLog('Skipping temp file: ' + file);
                continue;
            }

            // 跳过文件夹
            if (fs.statSync(sourcePath).isDirectory()) {
                continue;
            }

            // 检查文件是否存在于目标文件夹
            if (!fs.existsSync(targetPath)) {
                // 新增文件
                differences.push({ fileName: file, type: 'new' });
                debugLog('File is new (not in target): ' + file);
                continue;
            }

            // 使用多层级比较
            const diffType = FileComparer.compareFilesWithMultiLevel(sourcePath, targetPath);
            
            if (diffType === 'case-only') {
                differences.push({ fileName: file, type: 'case-only' });
            } else if (diffType === 'substantial') {
                differences.push({ fileName: file, type: 'substantial' });
            }
        }

        // 检查目标文件夹中的文件（删除的文件）
        for (const file of targetFiles) {
            const sourcePath = path.join(sourceDir, file);
            const targetPath = path.join(targetDir, file);

            // 跳过临时文件
            if (!includeTemp && FileComparer.isTempFile(file)) {
                debugLog('Skipping temp file: ' + file);
                continue;
            }

            // 跳过文件夹
            if (fs.existsSync(targetPath) && fs.statSync(targetPath).isDirectory()) {
                continue;
            }

            // 检查文件是否存在于源文件夹
            if (!fs.existsSync(sourcePath)) {
                // 删除的文件
                differences.push({ fileName: file, type: 'deleted' });
                debugLog('File is deleted (not in source): ' + file);
            }
        }

        return differences;
    }

    /**
     * 检查是否为临时文件
     * @param fileName 文件名
     * @returns 是否为临时文件
     */
    static isTempFile(fileName: string): boolean {
        // 检查临时文件夹
        if (fileName.startsWith('.vba-sync-tmp-')) {
            return true;
        }
        return false;
    }

    /**
     * 检查是否为临时文件夹
     * @param folderPath 文件夹路径
     * @returns 是否为临时文件夹
     */
    static isTempFolder(folderPath: string): boolean {
        const folderName = path.basename(folderPath);
        return folderName.startsWith('.vba-sync-tmp-');
    }
}
