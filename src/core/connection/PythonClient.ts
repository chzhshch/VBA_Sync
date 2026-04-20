import * as vscode from 'vscode';
import * as childProcess from 'child_process';
import * as path from 'path';
import { PythonRequest, PythonResponse } from '../../types';

// 调试日志开关
const DEBUG = true;

// 统一日志函数
function debugLog(message: string, ...args: any[]) {
    if (DEBUG) {
        console.log('[VBA SYNC DEBUG] ' + message, ...args);
    }
}

function debugError(message: string, ...args: any[]) {
    if (DEBUG) {
        console.error('[VBA SYNC DEBUG ERROR] ' + message, ...args);
    }
}

export class PythonClient {
    private pythonProcess: childProcess.ChildProcess | null = null;
    private requestId = 0;
    private pendingRequests = new Map<number, { resolve: (response: PythonResponse) => void; reject: (error: Error) => void }>();
    private stdoutBuffer = '';
    private healthCheckInterval: NodeJS.Timeout | null = null;
    private restartCount = 0;
    private maxRestarts = 3;
    private _disposed = false;

    constructor() {
        this.startPythonProcess();
    }

    private startPythonProcess() {
        const extensionPath = vscode.extensions.getExtension('vba-sync.vba-sync')?.extensionPath || __dirname;
        const pythonScriptPath = path.join(extensionPath, 'python', 'vba_sync.py');
        
        debugLog('Starting Python process with script: ' + pythonScriptPath);

        this.pythonProcess = childProcess.spawn('python', ['-u', pythonScriptPath], {
            stdio: ['pipe', 'pipe', 'pipe'],
            env: {
                ...(process.env as any),
                PYTHONUTF8: '1',
            },
        });

        this.pythonProcess.stdout?.on('data', (data) => {
            this.handleStdout(data);
        });

        this.pythonProcess.stderr?.on('data', (data) => {
            debugLog('Python stderr: ' + data.toString());
        });

        this.pythonProcess.on('error', (error) => {
            debugError('Python process error:', error);
        });

        this.pythonProcess.on('exit', (code) => {
            debugLog('Python process exited with code:', code);
            this.handleProcessExit();
        });
        
        this.setupHealthCheck();
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
                if (this.pythonProcess && !this._disposed) {
                    debugLog('Process error, restarting');
                    this.startPythonProcess();
                }
            }
        }, 60000); // 每60秒检查一次
    }
    
    private handleProcessExit() {
        if (this._disposed) {
            debugLog('Process already disposed, skipping exit handling');
            return;
        }

        if (this.restartCount < this.maxRestarts) {
            this.restartCount++;
            debugLog(`Restarting Python process (${this.restartCount}/${this.maxRestarts})`);
            
            setTimeout(() => {
                try {
                    this.startPythonProcess();
                } catch (error) {
                    debugError('Failed to restart Python process:', error);
                }
            }, 3000);
        } else {
            debugError('Maximum restart attempts reached');
        }
    }

    private handleStdout(data: Buffer) {
        this.stdoutBuffer += data.toString();
        
        const lines = this.stdoutBuffer.split('\n');
        for (let i = 0; i < lines.length - 1; i++) {
            const line = lines[i].trim();
            if (line) {
                this.parseResponse(line);
            }
        }
        
        this.stdoutBuffer = lines[lines.length - 1];
    }

    private parseResponse(line: string) {
        try {
            const response = JSON.parse(line) as PythonResponse;
            const requestId = response.id;
            
            if (requestId && this.pendingRequests.has(requestId)) {
                const { resolve, reject } = this.pendingRequests.get(requestId)!;
                this.pendingRequests.delete(requestId);
                
                if (response.success) {
                    resolve(response);
                } else {
                    reject(new Error(response.error || 'Unknown error'));
                }
            }
        } catch (error) {
            debugError('Failed to parse Python response:', error, 'Line:', line);
        }
    }

    async sendRequest(request: PythonRequest): Promise<PythonResponse> {
        return new Promise((resolve, reject) => {
            const id = ++this.requestId;
            const requestWithId = { ...request, id };
            
            this.pendingRequests.set(id, { resolve, reject });
            
            try {
                debugLog('Sending request ' + id + ': ' + request.action);
                const jsonString = JSON.stringify(requestWithId) + '\n';
                this.pythonProcess?.stdin?.write(jsonString);
            } catch (error) {
                debugError('Failed to send request:', error);
                this.pendingRequests.delete(id);
                reject(error);
            }
        });
    }

    async sendRequestWithRetry(request: PythonRequest, maxRetries = 3): Promise<PythonResponse> {
        let retries = 0;
        debugLog('Sending request with retry: ' + request.action + ', maxRetries: ' + maxRetries);

        while (retries <= maxRetries) {
            try {
                debugLog('Attempt ' + (retries + 1) + ': Sending request ' + request.action);
                const response = await this.sendRequest(request);
                debugLog('Received response for ' + request.action + ': ' + (response.success ? 'success' : 'error'));
                
                return response;
            } catch (error: any) {
                if (retries < maxRetries) {
                    debugLog('Request failed, retrying... (' + (retries + 1) + '/' + maxRetries + ')', error);
                    retries++;
                    // 指数退避
                    await new Promise(resolve => setTimeout(resolve, Math.pow(2, retries) * 1000));
                } else {
                    debugError('Request failed after ' + maxRetries + ' retries:', error);
                    throw error;
                }
            }
        }
        
        throw new Error('Request failed');
    }

    async dispose() {
        debugLog('Disposing PythonClient');
        this._disposed = true;

        if (this.healthCheckInterval) {
            debugLog('Clearing health check interval');
            clearInterval(this.healthCheckInterval);
        }

        if (this.pythonProcess) {
            debugLog('Showing all hidden Office windows before cleanup');
            try {
                await this.sendRequest({ action: 'showAllWindows', documentPath: '' });
                debugLog('Show all windows request sent and completed successfully');
                // 等待一小段时间，确保窗口显示完成
                await new Promise(resolve => setTimeout(resolve, 500));
            } catch (error) {
                debugError('Error sending show all windows request:', error);
                // 忽略错误，继续清理
            }

            debugLog('Killing Python process');
            this.pythonProcess.kill();
            this.pythonProcess = null;
        }

        debugLog('Clearing pending requests');
        this.pendingRequests.clear();
        debugLog('PythonClient disposed');
    }
}