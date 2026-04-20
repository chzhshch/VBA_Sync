#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
VBA Sync Python 桥接脚本
处理与 TypeScript 扩展的通信，以及 Office COM 操作
"""

import sys
import json
import traceback
import os

# 调试日志开关
DEBUG = True

# 统一日志函数
def debug_log(message):
    """调试日志"""
    if DEBUG:
        print(f'[VBA SYNC DEBUG] {message}', file=sys.stderr)
        sys.stderr.flush()

def debug_error(message, error=None):
    """调试错误日志"""
    if DEBUG:
        print(f'[VBA SYNC DEBUG ERROR] {message}', file=sys.stderr)
        if error:
            print(f'[VBA SYNC DEBUG ERROR] {error}', file=sys.stderr)
        sys.stderr.flush()

# 添加脚本所在目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from lib.utils import log_info, log_error

# 延迟导入Office相关模块，使其在win32com不可用时仍然能够响应ping请求
try:
    from lib.office_connector import OfficeConnector
    from lib.vba_exporter import VBAExporter
    from lib.vba_importer import VBAImporter
    from lib.properties import PropertiesManager
    OFFICE_MODULES_AVAILABLE = True
except ImportError as e:
    log_error(f'Office modules not available: {e}')
    OFFICE_MODULES_AVAILABLE = False
    OfficeConnector = None
    VBAExporter = None
    VBAImporter = None
    PropertiesManager = None

class VbaSyncBridge:
    def __init__(self):
        self.connectors = {}
        self.exporters = {}
        self.importers = {}
        self.properties = {}

    def handle_request(self, request):
        try:
            action = request.get('action')
            document_path = request.get('documentPath')
            debug_log(f'Handling request: {action} for document: {document_path}')

            if action == 'ping':
                debug_log('Handling ping request')
                return {'success': True, 'data': {'message': 'pong', 'officeModulesAvailable': OFFICE_MODULES_AVAILABLE}}

            if action == 'cleanup':
                debug_log('Handling cleanup request')
                return self.cleanup()

            if action == 'showAllWindows':
                debug_log('Handling show all windows request')
                return self.show_all_windows()

            # 检查Office模块是否可用
            if not OFFICE_MODULES_AVAILABLE:
                debug_error('Office modules not available')
                return {'success': False, 'error': 'Office modules not available. Please install pywin32.'}

            if not document_path:
                debug_error('Document path is required')
                return {'success': False, 'error': 'Document path is required'}

            # 确保连接器存在
            if document_path not in self.connectors:
                debug_log(f'Creating new connectors for document: {document_path}')
                self.connectors[document_path] = OfficeConnector()
                self.exporters[document_path] = VBAExporter(self.connectors[document_path])
                self.importers[document_path] = VBAImporter(self.connectors[document_path])
                self.properties[document_path] = PropertiesManager(self.connectors[document_path])
                debug_log(f'Connectors created successfully for document: {document_path}')

            connector = self.connectors[document_path]
            exporter = self.exporters[document_path]
            importer = self.importers[document_path]
            properties = self.properties[document_path]

            # 检查文档连接状态
            if hasattr(connector, 'is_document_connection') and connector.is_document_connection(document_path):
                debug_log(f'Document {document_path} is in document connection state')
                # 确定同步方向 - 根据操作类型设置默认值
                sync_direction = request.get('syncDirection')
                if sync_direction is None:
                    # 根据操作类型设置默认同步方向
                    if action in ['export_all', 'export_module_code', 'list_modules']:
                        # 导出操作默认是从文档到IDE
                        sync_direction = 'doc-to-ide'
                    else:
                        # 其他操作默认是从IDE到文档
                        sync_direction = 'ide-to-doc'
                debug_log(f'Determined sync direction: {sync_direction}')
                if sync_direction == 'doc-to-ide':
                    # 从文档侧发起的同步：关闭并重新打开文档
                    debug_log('Sync direction: doc-to-ide, closing and reopening document')
                    try:
                        # 关闭当前文档
                        close_result = connector.close_document(document_path)
                        if close_result['success']:
                            debug_log('Document closed successfully')
                            # 重新打开文档
                            app_type = request.get('appType', 'excel')
                            open_result = connector.open_document(document_path, app_type, visible=True)
                            if open_result['success']:
                                debug_log('Document reopened successfully')
                                # 移除文档连接状态标记
                                if hasattr(connector, 'remove_document_connection'):
                                    connector.remove_document_connection(document_path)
                            else:
                                debug_error(f'Failed to reopen document: {open_result.get("error", "Unknown error")}')
                                return {'success': False, 'error': f'无法重新打开文档: {open_result.get("error", "Unknown error")}'}
                        else:
                            debug_error(f'Failed to close document: {close_result.get("error", "Unknown error")}')
                            return {'success': False, 'error': f'无法关闭文档: {close_result.get("error", "Unknown error")}'}
                    except Exception as e:
                        debug_error(f'Error handling document connection: {e}')
                        return {'success': False, 'error': f'处理文档连接时出错: {str(e)}'}
                else:
                    # 从IDE侧发起的同步：通知用户关闭文档
                    debug_log('Sync direction: ide-to-doc, notifying user to close document')
                    document_name = os.path.basename(document_path)
                    return {'success': False, 'error': f'文档访问异常，请关闭已打开的【{document_name}】后重试。'}

            if action == 'open_document':
                app_type = request.get('appType', 'excel')
                visible = request.get('visible', False)
                only_if_open = request.get('onlyIfOpen', False)
                debug_log(f'Opening document: {document_path} with app_type: {app_type}, visible: {visible}, only_if_open: {only_if_open}')
                result = connector.open_document(document_path, app_type, visible, only_if_open)
                debug_log(f'Open document result: {result}')
                return {'success': True, 'data': result}

            elif action == 'check_vba_access':
                debug_log(f'Checking VBA access for document: {document_path}')
                result = connector.check_vba_access(document_path)
                debug_log(f'VBA access check result: {result}')
                return {'success': True, 'data': result}

            elif action == 'export_all':
                output_dir = request.get('outputDir')
                if not output_dir:
                    debug_error('Output directory is required')
                    return {'success': False, 'error': 'Output directory is required'}
                debug_log(f'Exporting all modules from document: {document_path} to: {output_dir}')
                result = exporter.export_all(document_path, output_dir)
                debug_log(f'Export all result: {result}')
                # 直接返回 exporter 的结果，不再包装在 data 中
                return result

            elif action == 'import_module':
                file_path = request.get('filePath')
                if not file_path:
                    debug_error('File path is required')
                    return {'success': False, 'error': 'File path is required'}
                module_name = request.get('moduleName')
                debug_log(f'Importing module: {module_name} from file: {file_path} to document: {document_path}')
                result = importer.import_module(document_path, file_path, module_name)
                debug_log(f'Import module result: {result}')
                # 直接返回 importer 的结果，不再包装在 data 中
                return result

            elif action == 'import_all':
                directory = request.get('directory')
                if not directory:
                    debug_error('Directory is required')
                    return {'success': False, 'error': 'Directory is required'}
                debug_log(f'Importing all modules from directory: {directory} to document: {document_path}')
                result = importer.import_all(document_path, directory)
                debug_log(f'Import all result: {result}')
                # 直接返回 importer 的结果，不再包装在 data 中
                return result

            elif action == 'delete_module':
                module_name = request.get('moduleName')
                if not module_name:
                    debug_error('Module name is required')
                    return {'success': False, 'error': 'Module name is required'}
                debug_log(f'Deleting module: {module_name} from document: {document_path}')
                result = importer.delete_module(document_path, module_name)
                debug_log(f'Delete module result: {result}')
                return {'success': True, 'data': result}

            elif action == 'clear_module_code':
                module_name = request.get('moduleName')
                if not module_name:
                    debug_error('Module name is required')
                    return {'success': False, 'error': 'Module name is required'}
                debug_log(f'Clearing module code: {module_name} from document: {document_path}')
                result = importer.clear_module_code(document_path, module_name)
                debug_log(f'Clear module code result: {result}')
                return {'success': True, 'data': result}

            elif action == 'export_module_code':
                module_name = request.get('moduleName')
                if not module_name:
                    debug_error('Module name is required')
                    return {'success': False, 'error': 'Module name is required'}
                debug_log(f'Exporting module code: {module_name} from document: {document_path}')
                result = exporter.export_module_code(document_path, module_name)
                debug_log(f'Export module code result: {result}')
                return {'success': True, 'data': result}

            elif action == 'list_modules':
                debug_log(f'Listing modules for document: {document_path}')
                result = exporter.list_modules(document_path)
                debug_log(f'List modules result: {result}')
                return {'success': True, 'data': result}

            elif action == 'get_last_sync_time':
                debug_log(f'Getting last sync time for document: {document_path}')
                result = properties.get_last_sync_time(document_path)
                debug_log(f'Last sync time: {result}')
                return {'success': True, 'data': {'lastSyncTime': result}}

            elif action == 'set_last_sync_time':
                sync_time = request.get('syncTime')
                debug_log(f'Setting last sync time for document: {document_path} to: {sync_time}')
                result = properties.set_last_sync_time(document_path, sync_time)
                debug_log(f'Set last sync time result: {result}')
                return {'success': True, 'data': result}

            elif action == 'set_window_visible':
                visible = request.get('visible', False)
                debug_log(f'Setting window visible for document: {document_path} to: {visible}')
                result = connector.set_window_visible(document_path, visible)
                debug_log(f'Set window visible result: {result}')
                return {'success': True, 'data': result}

            elif action == 'save_document':
                debug_log(f'Saving document: {document_path}')
                result = connector.save_document(document_path)
                debug_log(f'Save document result: {result}')
                return {'success': True, 'data': result}

            elif action == 'close_document':
                debug_log(f'Closing document: {document_path}')
                result = connector.close_document(document_path)
                debug_log(f'Close document result: {result}')
                if result['success']:
                    # 清理相关对象
                    debug_log(f'Cleaning up objects for document: {document_path}')
                    if document_path in self.connectors:
                        del self.connectors[document_path]
                    if document_path in self.exporters:
                        del self.exporters[document_path]
                    if document_path in self.importers:
                        del self.importers[document_path]
                    if document_path in self.properties:
                        del self.properties[document_path]
                    debug_log(f'Objects cleaned up for document: {document_path}')
                return {'success': True, 'data': result}

            else:
                debug_error(f'Unknown action: {action}')
                return {'success': False, 'error': f'Unknown action: {action}'}

        except Exception as e:
            debug_error(f'Error handling request: {e}')
            debug_error(traceback.format_exc())
            error_code = None
            # 提取 COM 错误码
            if hasattr(e, 'args') and len(e.args) > 0:
                import win32com
                if isinstance(e, win32com.client.pywintypes.com_error):
                    error_code = e.args[0]
            return {
                'success': False,
                'error': str(e),
                'errorCode': error_code
            }

    def show_all_windows(self):
        """显示所有隐藏的Office窗口"""
        try:
            debug_log('Showing all hidden Office windows')
            for document_path, connector in self.connectors.items():
                try:
                    debug_log(f'Setting window visible for document: {document_path}')
                    connector.set_window_visible(document_path, True)
                    debug_log(f'Window visible set for document: {document_path}')
                except Exception as e:
                    debug_error(f'Error showing window for document {document_path}: {e}')
            debug_log('All hidden windows shown')
            return {'success': True, 'data': {'message': 'All hidden windows shown'}}
        except Exception as e:
            debug_error(f'Error showing windows: {e}')
            return {'success': False, 'error': str(e)}

    def cleanup(self):
        """清理所有资源"""
        try:
            debug_log('Starting cleanup process')
            for document_path, connector in self.connectors.items():
                try:
                    debug_log(f'Closing document: {document_path}')
                    connector.close_document(document_path)
                    debug_log(f'Document closed: {document_path}')
                except Exception as e:
                    debug_error(f'Error closing document {document_path}: {e}')
            debug_log('Clearing connectors')
            self.connectors.clear()
            debug_log('Clearing exporters')
            self.exporters.clear()
            debug_log('Clearing importers')
            self.importers.clear()
            debug_log('Clearing properties')
            self.properties.clear()
            debug_log('Cleanup completed')
            return {'success': True, 'data': {'message': 'Cleanup completed'}}
        except Exception as e:
            debug_error(f'Error during cleanup: {e}')
            return {'success': False, 'error': str(e)}

def main():
    """主函数"""
    # 只输出JSON格式的响应，其他信息输出到stderr
    import sys
    debug_log('VBA Sync Python bridge starting...')
    debug_log(f'Python version: {sys.version}')
    debug_log(f'Office modules available: {OFFICE_MODULES_AVAILABLE}')
    
    bridge = VbaSyncBridge()
    debug_log('VbaSyncBridge initialized')

    # 读取标准输入
    debug_log('Waiting for input...')
    
    for line in sys.stdin:
        debug_log(f'Received line: {line.strip()}')
        
        line = line.strip()
        if not line:
            continue

        try:
            debug_log('Parsing JSON...')
            request = json.loads(line)
            debug_log(f'Parsed request: {json.dumps(request)}')
            
            debug_log('Handling request...')
            response = bridge.handle_request(request)
            debug_log(f'Generated response: {json.dumps(response)}')
            
            # 添加请求 ID
            if 'id' in request:
                response['id'] = request['id']
                debug_log(f'Added request ID: {request["id"]}')
            
            # 输出响应（只输出JSON格式）
            debug_log(f'Sending response: {json.dumps(response)}')
            print(json.dumps(response))
            sys.stdout.flush()
            debug_log('Response sent successfully')
        except json.JSONDecodeError as e:
            debug_error(f'JSON decode error: {e}')
            response = {'success': False, 'error': f'Invalid JSON: {e}'}
            print(json.dumps(response))
            sys.stdout.flush()
        except Exception as e:
            debug_error(f'Unexpected error: {e}')
            debug_error(traceback.format_exc())
            response = {'success': False, 'error': str(e)}
            print(json.dumps(response))
            sys.stdout.flush()

if __name__ == '__main__':
    main()
