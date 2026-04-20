import { spawn, ChildProcess } from 'child_process';
import * as path from 'path';
import * as vscode from 'vscode';
import { PythonRequest, PythonResponse } from '../types';

// 调试日志开关
const DEBUG = true;

// 统一日志函数
function debugLog(message: string, ...args: any[]) {
    if (DEBUG) {
        console.log(`[VBA SYNC DEBUG] ${message}`, ...args);
    }
}

function debugError(message: string, ...args: any[]) {
    if (DEBUG) {
        console.error(`[VBA SYNC DEBUG ERROR] ${message}`, ...args);
    }
}

class ExponentialBackoff {
    private initialDelay: number;
    private maxDelay: number;
    private delay: number;

    constructor(options: { initialDelay: number; maxDelay: number }) {
        this.initialDelay = options.initialDelay;
        this.maxDelay = options.maxDelay;
        this.delay = options.initialDelay;
    }

    wait(): Promise<void> {
        return new Promise(resolve => setTimeout(resolve, this.delay));
    }

    reset() {
        this.delay = this.initialDelay;
    }

    increase() {
        this.delay = Math.min(this.delay * 2, this.maxDelay);
    }
}

export class PythonRunner {
    private process: ChildProcess | null = null;
    private requestId = 0;
    private pendingRequests = new Map<number, (response: PythonResponse) => void>();
    private stdoutBuffer = '';
    private restartCount = 0;
    private maxRestarts = 3;
    private _disposed = false;
    private healthCheckInterval: NodeJS.Timeout | null = null;

    async start() {
        if (this.process) {
            debugLog('Python process already running');
            return;
        }

        const pythonPath = this.getPythonPath();
        // 使用更可靠的路径计算方式
        const extensionPath = path.resolve(__dirname, '..');
        let pythonScript = path.join(extensionPath, 'python', 'vba_sync.py');

        debugLog(`Starting Python process with: ${pythonPath} ${pythonScript}`);
        debugLog(`Current directory: ${__dirname}`);
        debugLog(`Extension path: ${extensionPath}`);
        
        // 检查Python脚本是否存在
        const fs = require('fs');
        if (fs.existsSync(pythonScript)) {
            debugLog(`Python script exists: ${pythonScript}`);
        } else {
            debugError(`Python script does not exist: ${pythonScript}`);
            // 尝试其他可能的路径
            const alternativePath = path.join(__dirname, 'python', 'vba_sync.py');
            debugLog(`Trying alternative path: ${alternativePath}`);
            if (fs.existsSync(alternativePath)) {
                debugLog(`Alternative path exists: ${alternativePath}`);
                pythonScript = alternativePath;
            }
        }

        try {
            debugLog(`Spawning Python process with command: ${pythonPath} -u ${pythonScript}`);
            debugLog(`Environment variables: PYTHONUTF8=1`);
            
            this.process = spawn(pythonPath, ['-u', pythonScript], {
                env: {
                    ...(process.env as any),
                    PYTHONUTF8: '1',
                },
            });

            debugLog('Python process spawned successfully, PID:', this.process.pid);

            this.setupEventListeners();
            this.setupHealthCheck();

            debugLog('Waiting for Python process to be ready...');
            await this.waitForProcessReady();
            debugLog('Python process is ready');
        } catch (error) {
            debugError('Failed to start Python process:', error);
            throw new Error(`Failed to start Python process: ${error}`);
        }
    }

    private async waitForProcessReady(timeoutMs = 10000): Promise<void> {
        const startTime = Date.now();
        const backoff = new ExponentialBackoff({ initialDelay: 100, maxDelay: 2000 });
        
        debugLog(`Waiting for Python process to be ready... (timeout: ${timeoutMs}ms)`);
        
        while (Date.now() - startTime < timeoutMs) {
            try {
                debugLog(`Sending ping request to Python process... (attempt: ${Date.now() - startTime}ms)`);
                const response = await this.sendRequest({ action: 'ping', documentPath: '' });
                debugLog('Python process responded to ping:', JSON.stringify(response));
                return;
            } catch (error) {
                debugLog(`Ping failed, retrying... (${Date.now() - startTime}ms):`, error instanceof Error ? error.message : String(error));
                await backoff.wait();
                backoff.increase();
            }
        }
        debugError('Python process startup timeout');
        throw new Error('Python 进程启动超时');
    }

    private getPythonPath() {
        const config = vscode.workspace.getConfiguration('vbaSync');
        return config.get('pythonPath') as string || 'python';
    }

    private setupEventListeners() {
        const process = this.process;
        if (!process) return;

        if (process.stdout) {
            process.stdout.on('data', (data) => {
                debugLog('Python stdout:', data.toString());
                this.handleStdout(data);
            });
        }

        if (process.stderr) {
            process.stderr.on('data', (data) => {
                debugLog('Python stderr:', data.toString());
            });
        }

        process.on('exit', (code, signal) => {
            debugLog('Python process exited with code:', code, 'signal:', signal);
            this.handleProcessExit();
        });

        process.on('error', (error) => {
            debugError('Python process error:', error);
        });

        process.on('spawn', () => {
            debugLog('Python process spawned successfully');
        });
    }

    private handleStdout(data: Buffer) {
        debugLog('Handling stdout data:', data.toString());
        this.stdoutBuffer += data.toString();
        const lines = this.stdoutBuffer.split('\n');

        for (const line of lines) {
            if (line.trim() === '') continue;

            try {
                debugLog('Parsing JSON line:', line);
                const response: PythonResponse = JSON.parse(line);
                const id = response.id;

                if (id !== undefined && this.pendingRequests.has(id)) {
                    debugLog('Found matching request for id:', id);
                    const callback = this.pendingRequests.get(id)!;
                    this.pendingRequests.delete(id);
                    callback(response);
                } else {
                    debugLog('No matching request found for id:', id);
                }
            } catch (error) {
                debugError('Failed to parse Python response:', error);
            }
        }

        // 保留最后一个不完整的行
        if (this.stdoutBuffer.endsWith('\n')) {
            debugLog('Buffer ends with newline, clearing buffer');
            this.stdoutBuffer = '';
        } else {
            debugLog('Buffer does not end with newline, keeping last line:', lines[lines.length - 1]);
            this.stdoutBuffer = lines[lines.length - 1];
        }
    }

    private handleProcessExit() {
        if (this._disposed) {
            debugLog('Process already disposed, skipping exit handling');
            return;
        }

        if (this.restartCount < this.maxRestarts) {
            this.restartCount++;
            debugLog(`Restarting Python process (${this.restartCount}/${this.maxRestarts})`);
            
            setTimeout(async () => {
                try {
                    await this.start();
                } catch (error) {
                    debugError('Failed to restart Python process:', error);
                }
            }, 3000);
        } else {
            debugError('Maximum restart attempts reached');
        }
    }

    private setupHealthCheck() {
        debugLog('Setting up health check interval');
        this.healthCheckInterval = setInterval(async () => {
            try {
                debugLog('Performing health check ping');
                await this.sendRequest({ action: 'ping', documentPath: '' });
                debugLog('Health check successful');
            } catch (error) {
                debugError('Health check failed:', error);
                if (this.process && this.process.killed) {
                    debugLog('Process killed, restarting');
                    await this.start();
                }
            }
        }, 60000); // 每60秒检查一次
    }

    async sendRequest(request: PythonRequest): Promise<PythonResponse> {
        if (!this.process || this.process.killed) {
            debugLog('Python process not available, starting...');
            await this.start();
        }

        const id = this.requestId++;
        const requestWithId = { ...request, id };

        debugLog(`Sending request ${id}: ${request.action}`);

        return new Promise((resolve, reject) => {
            const timeout = setTimeout(() => {
                debugError(`Request ${id} timeout`);
                this.pendingRequests.delete(id);
                reject(new Error(`Request timeout: ${request.action}`));
            }, 10000);

            this.pendingRequests.set(id, (response) => {
                clearTimeout(timeout);
                debugLog(`Received response for request ${id}: ${response.success ? 'success' : 'error'}`);
                if (response.success) {
                    resolve(response);
                } else {
                    debugError(`Request ${id} failed: ${response.error}`);
                    reject(new Error(response.error || 'Unknown error'));
                }
            });

            if (this.process && this.process.stdin) {
                const requestString = JSON.stringify(requestWithId) + '\n';
                debugLog(`Writing to stdin: ${requestString.trim()}`);
                this.process.stdin.write(requestString);
            } else {
                clearTimeout(timeout);
                debugError('Python process stdin not available');
                reject(new Error('Python process not available'));
            }
        });
    }

    async dispose() {
        debugLog('Disposing PythonRunner');
        this._disposed = true;

        if (this.healthCheckInterval) {
            debugLog('Clearing health check interval');
            clearInterval(this.healthCheckInterval);
        }

        if (this.process) {
            debugLog('Sending cleanup request to Python process');
            try {
                await this.sendRequest({ action: 'cleanup', documentPath: '' });
                debugLog('Cleanup request sent successfully');
            } catch (error) {
                debugError('Error sending cleanup request:', error);
                // 忽略错误
            }

            debugLog('Killing Python process');
            this.process.kill();
            this.process = null;
        }

        debugLog('Clearing pending requests');
        this.pendingRequests.clear();
        debugLog('PythonRunner disposed');
    }
}
