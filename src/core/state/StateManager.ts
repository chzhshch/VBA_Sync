import { DocumentSyncState } from '../../types';

// 调试日志开关
const DEBUG = true;

// 统一日志函数
function debugLog(message: string, ...args: any[]) {
    if (DEBUG) {
        console.log('[VBA SYNC DEBUG] ' + message, ...args);
    }
}

export class StateManager {
    private states: Map<string, DocumentSyncState> = new Map();

    initState(documentPath: string): DocumentSyncState {
        const state = this.createDefaultState();
        this.states.set(documentPath, state);
        debugLog('Initialized state for document: ' + documentPath);
        return state;
    }

    private createDefaultState(): DocumentSyncState {
        return {
            isOfficeOpen: false,
            isWindowVisible: false,
            isPasswordUnlocked: false,
            lastSyncTime: null,
            lastDocToIdeSyncTime: null,
            lastDebugTime: null,
            pendingChanges: 0,
            isDebugging: false,
            ideCodeSynced: false,
            documentHasUnsavedChanges: false,
        };
    }

    getState(documentPath: string): DocumentSyncState | undefined {
        return this.states.get(documentPath);
    }

    setState(documentPath: string, state: Partial<DocumentSyncState>): DocumentSyncState {
        const currentState = this.getState(documentPath) || this.initState(documentPath);
        const newState = { ...currentState, ...state };
        this.states.set(documentPath, newState);
        debugLog('Updated state for document: ' + documentPath, newState);
        return newState;
    }

    deleteState(documentPath: string) {
        this.states.delete(documentPath);
        debugLog('Deleted state for document: ' + documentPath);
    }

    incrementPendingChanges(documentPath: string) {
        const state = this.getState(documentPath);
        if (state) {
            state.pendingChanges += 1;
            this.states.set(documentPath, state);
            debugLog('Incremented pending changes for document: ' + documentPath + ', now: ' + state.pendingChanges);
        }
    }

    getAllDocumentPaths(): string[] {
        return Array.from(this.states.keys());
    }
}
