// 调试日志开关
const DEBUG = true;

// 统一日志函数
function debugLog(message: string, ...args: any[]) {
    if (DEBUG) {
        console.log('[VBA SYNC DEBUG] ' + message, ...args);
    }
}

export class ErrorHandler {
    /**
     * 包装异步函数，捕获并处理错误
     * @param fn 异步函数
     * @returns 包装后的函数
     */
    static async wrapAsync<T>(fn: () => Promise<T>): Promise<T | undefined> {
        try {
            return await fn();
        } catch (error: any) {
            debugLog('Error caught:', error);
            return undefined;
        }
    }

    /**
     * 处理错误
     * @param error 错误对象
     * @param message 错误消息
     */
    static handleError(error: any, message: string) {
        debugLog(message, error);
        // 这里可以添加更多错误处理逻辑，如日志记录、用户通知等
    }
}
