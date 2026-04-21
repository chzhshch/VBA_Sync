// VBA 组件类型
export enum VBAModuleType {
    Standard = 1,    // .bas
    Class = 2,       // .cls
    Form = 3,        // .frm（不支持）
    Document = 100,  // 文档模块
}

// Office 应用类型
export type OfficeAppType = 'excel' | 'powerpoint' | 'word';

// 配置
export interface DocumentConfig {
    path: string;       // 相对路径
    vbaFolder: string;  // VBA 文件夹相对路径
    enabled: boolean;
}

export interface VbaSyncConfig {
    version: string;
    documents: DocumentConfig[];
    settings: {
        autoSync: boolean;
        conflictStrategy: 'ask' | 'ide-priority' | 'document-priority';
        showNotifications: boolean;
    };
}

// 状态
export interface DocumentSyncState {
    isOfficeOpen: boolean;
    isWindowVisible: boolean;
    isPasswordUnlocked: boolean;
    lastSyncTime: Date | null;
    lastDocToIdeSyncTime?: Date | null;
    lastDebugTime: Date | null;
    pendingChanges: number;
    isDebugging: boolean;
    ideCodeSynced: boolean;
    documentHasUnsavedChanges: boolean;
}

// 通信
export interface PythonRequest {
    action: string;
    documentPath: string;
    id?: number;
    [key: string]: unknown;
}

export interface PythonResponse {
    success: boolean;
    data?: Record<string, unknown>;
    error?: string;
    errorCode?: number;
    id?: number;
    [key: string]: unknown;
}

// 常量映射
export const EXT_TO_APP: Record<string, OfficeAppType> = {
    '.xlsm': 'excel',
    '.xlsb': 'excel',
    '.xltm': 'excel',
    '.pptm': 'powerpoint',
    '.potm': 'powerpoint',
    '.ppsm': 'powerpoint',
    '.docm': 'word',
    '.dotm': 'word',
};

export const APP_PROGIDS: Record<OfficeAppType, string> = {
    excel: 'Excel.Application',
    powerpoint: 'PowerPoint.Application',
    word: 'Word.Application',
};

export const OFFICE_MACRO_EXTENSIONS = [
    '.xlsm', '.xlsb', '.xltm',
    '.pptm', '.potm', '.ppsm',
    '.docm', '.dotm'
];

export const VBA_FILE_EXTENSIONS = ['.bas', '.cls'];

export const COM_ERROR_CODES = {
    VBA_PROJECT_LOCKED: -2147221491,
    VBA_ACCESS_DENIED: -2146827284,
    DOCUMENT_IN_USE: -2147417846,
    CALL_REJECTED: -2147418113,
};

// VBA 模块信息
export interface VBAModule {
    name: string;
    type: number;
    hasCode: boolean;
}

// 同步详细信息类型
export interface SyncDetails {
    addedModules: string[];
    deletedModules: string[];
    clearedModules: string[];
    modifiedModules: string[];
}
