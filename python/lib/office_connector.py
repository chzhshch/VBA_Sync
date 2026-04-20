#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Office COM 连接器
处理与 Office 应用的通信
"""

import win32com.client
import os
from .utils import log_info, log_error


def _paths_equal(path1: str, path2: str) -> bool:
    """比较两个路径是否相等，处理大小写和格式差异"""
    try:
        norm1 = os.path.normpath(os.path.normcase(os.path.abspath(path1)))
        norm2 = os.path.normpath(os.path.normcase(os.path.abspath(path2)))
        result = norm1 == norm2
        if result:
            log_info(f'Paths match: {norm1} == {norm2}')
        else:
            log_info(f'Paths DO NOT match: {norm1} != {norm2}')
        return result
    except Exception as e:
        log_error(f'Error comparing paths: {e}')
        # 降级处理：直接比较
        return path1.lower() == path2.lower()


class OfficeConnector:
    def __init__(self):
        self.apps = {}
        self.documents = {}
        self.document_connections = {}  # 标记文档连接状态

    def open_document(self, document_path, app_type, visible=False, only_if_open=False, syncDirection=None):
        """打开 Office 文档 - 重构版本"""
        try:
            log_info(f'Opening document: {document_path}, app_type={app_type}, only_if_open={only_if_open}, syncDirection={syncDirection}')

            # 快速检查：文档是否已经在我们的字典中
            if document_path in self.documents:
                # 验证连接是否有效
                if self._is_connection_valid(document_path):
                    log_info('Document already connected and valid')
                    return {'success': True, 'message': 'Document already open'}
                else:
                    log_info('Document connection invalid, removing and reconnecting')
                    self._remove_invalid_connection(document_path)

            # 对于文档侧同步，优先使用GetObject
            if syncDirection == 'doc-to-ide':
                log_info('Using GetObject for doc-to-ide sync')
                try:
                    import win32com.client
                    # 使用GetObject获取文档对象
                    doc = win32com.client.GetObject(document_path)
                    log_info(f'Successfully got document object using GetObject: {type(doc)}')
                    
                    # 获取Application对象
                    app = None
                    if hasattr(doc, 'Application'):
                        app = doc.Application
                        log_info(f'Successfully got Application from document object: {type(app)}')
                        
                        # 验证Application对象
                        if app is not None:
                            try:
                                app_name = app.Name
                                log_info(f'Application object validated: {app_name}')
                            except Exception as validate_err:
                                log_info(f'Application object validation failed: {validate_err}')
                                app = None
                    
                    # 保存连接
                    self.documents[document_path] = doc
                    
                    # 确保应用实例在我们的字典中（如果没有）
                    if app is not None:
                        if app_type not in self.apps:
                            self.apps[app_type] = app

                        # 始终保持窗口可见
                        app.Visible = True
                    else:
                        log_info('Document found but no application instance available')
                        # 标记为文档连接状态
                        if not hasattr(self, 'document_connections'):
                            self.document_connections = {}
                        self.document_connections[document_path] = True

                    return {'success': True, 'message': 'Connected to document using GetObject'}
                except Exception as getobject_err:
                    log_info(f'Error using GetObject: {getobject_err}, falling back to ROT search')
                    # GetObject失败，回退到常规查找

            # 对于IDE侧同步，根据应用类型决定是否使用GetObject
            if syncDirection == 'ide-to-doc' and app_type in ['excel', 'word']:
                log_info('Using GetObject for ide-to-doc sync (Excel/Word)')
                try:
                    import win32com.client
                    # 使用GetObject获取文档对象
                    doc = win32com.client.GetObject(document_path)
                    log_info(f'Successfully got document object using GetObject: {type(doc)}')
                    
                    # 获取Application对象
                    app = None
                    if hasattr(doc, 'Application'):
                        app = doc.Application
                        log_info(f'Successfully got Application from document object: {type(app)}')
                        
                        # 验证Application对象
                        if app is not None:
                            try:
                                app_name = app.Name
                                log_info(f'Application object validated: {app_name}')
                                # 始终保持窗口可见
                                app.Visible = True
                            except Exception as validate_err:
                                log_info(f'Application object validation failed: {validate_err}')
                                app = None
                    
                    # 保存连接
                    self.documents[document_path] = doc
                    
                    # 确保应用实例在我们的字典中（如果没有）
                    if app is not None:
                        if app_type not in self.apps:
                            self.apps[app_type] = app
                    else:
                        log_info('Document found but no application instance available')
                        # 标记为文档连接状态
                        if not hasattr(self, 'document_connections'):
                            self.document_connections = {}
                        self.document_connections[document_path] = True

                    return {'success': True, 'message': 'Connected to document using GetObject'}
                except Exception as getobject_err:
                    log_info(f'Error using GetObject: {getobject_err}, falling back to ROT search')
                    # GetObject失败，回退到常规查找

            # Step 1: 完整查找文档是否已在任何地方打开
            found, doc, app = self._try_find_open_document(document_path, app_type)

            if found and doc is not None:
                log_info('Found document already open! Connecting to it...')
                # 找到已打开的文档！
                self.documents[document_path] = doc

                # 确保应用实例在我们的字典中（如果没有）
                if app is not None:
                    if app_type not in self.apps:
                        self.apps[app_type] = app

                    # 始终保持窗口可见
                    app.Visible = True
                else:
                    log_info('Document found but no application instance available')
                    # 标记为文档连接状态
                    if not hasattr(self, 'document_connections'):
                        self.document_connections = {}
                    self.document_connections[document_path] = True

                return {'success': True, 'message': 'Connected to already open document'}

            # Step 2: 文档未找到，根据 only_if_open 参数决定下一步
            if only_if_open:
                log_info('Document not open and only_if_open=True, returning failure')
                return {'success': False, 'error': 'Document not open'}

            # Step 3: 需要打开文档
            log_info('Document not open, opening it now...')
            return self._open_new_document(document_path, app_type, visible)

        except Exception as e:
            log_error(f'Error in open_document: {e}')
            return {'success': False, 'error': str(e)}

    def _process_rot_moniker(self, moniker, app_type, instances, target_document_path=None):
        """
        处理 ROT 中的单个 moniker，尝试获取 Office 应用实例
        如果提供了 target_document_path，还会直接检查是否是目标文档
        返回: (found_target: bool, document: object|None, app: object|None)
        """
        log_info('=== STARTING ROT MONIKER PROCESSING ===')
        try:
            import pythoncom
            
            # 尝试获取 moniker 的显示名称（用于调试）
            display_name = None
            try:
                ctx = pythoncom.CreateBindCtx(0)
                display_name = moniker.GetDisplayName(ctx, None)
                log_info(f'ROT moniker display name: {display_name}')
            except Exception as name_err:
                log_info(f'Could not get moniker display name: {name_err}')
            
            # 首先检查display name是否直接匹配目标文档路径
            if target_document_path is not None and display_name is not None:
                log_info(f'Checking display name against target path: {target_document_path}')
                if _paths_equal(display_name, target_document_path):
                    log_info('Display name matches target document path!')
                    # 尝试获取文档对象
                    obj = None
                    try:
                        rot = pythoncom.GetRunningObjectTable()
                        obj = rot.GetObject(moniker)
                        log_info('Successfully got document object from ROT')
                        log_info(f'Object type before conversion: {type(obj)}')
                        
                        # 优先尝试使用GetObject获取对象
                        log_info('Trying to get object using GetObject first...')
                        try:
                            import win32com.client
                            # 使用GetObject获取对象
                            getobject_obj = win32com.client.GetObject(display_name)
                            log_info(f'Successfully got object using GetObject: {type(getobject_obj)}')
                            
                            # 尝试从GetObject返回的对象获取Application
                            if hasattr(getobject_obj, 'Application'):
                                log_info('GetObject returned object has Application attribute')
                                try:
                                    app = getobject_obj.Application
                                    log_info('Successfully got Application from GetObject returned object')
                                    # 验证Application对象的有效性
                                    if app is not None:
                                        try:
                                            app_name = app.Name
                                            log_info(f'Application object from GetObject validated: {app_name}')
                                            return True, getobject_obj, app
                                        except Exception as validate_err:
                                            log_info(f'Application object validation failed: {validate_err}')
                                except Exception as app_err:
                                    log_info(f'Error getting Application from GetObject returned object: {app_err}')
                            else:
                                log_info('GetObject returned object has no Application attribute')
                        except Exception as getobject_err:
                            log_info(f'Error using GetObject: {getobject_err}')
                        
                        # 如果GetObject失败，尝试转换PyIUnknown对象
                        try:
                            import win32com.client
                            # 尝试使用Dispatch来转换对象
                            if hasattr(obj, '__class__') and 'PyIUnknown' in str(type(obj)):
                                log_info('Attempting to convert PyIUnknown object')
                                # 尝试使用Dispatch转换
                                try:
                                    obj = win32com.client.Dispatch(obj)
                                    log_info('Successfully converted PyIUnknown object using Dispatch')
                                    log_info(f'Object type after conversion: {type(obj)}')
                                except Exception as conv_err:
                                    log_info(f'Error converting PyIUnknown object: {conv_err}')
                        except Exception as conv_err:
                            log_info(f'Error converting object: {conv_err}')
                        
                        # 尝试从文档对象获取Application
                        app = None
                        if hasattr(obj, 'Application'):
                            log_info('Document has Application attribute, trying to get it...')
                            # 尝试获取Application对象
                            try:
                                app = obj.Application
                                log_info('Successfully got Application object')
                                # 验证Application对象的有效性
                                if app is not None:
                                    try:
                                        app_name = app.Name
                                        log_info(f'Application object validated: {app_name}')
                                        return True, obj, app
                                    except Exception as validate_err:
                                        log_info(f'Application object validation failed: {validate_err}')
                                        app = None
                            except Exception as app_err:
                                log_info(f'Error getting Application from document object: {app_err}')
                        else:
                            log_info('Document object found but no Application attribute')
                        
                        # 即使获取Application失败，也返回文档对象
                        return True, obj, None
                    except Exception as e:
                        log_info(f'Error getting document object: {e}')
            
            # 使用正确的ROT API获取对象
            obj = None
            try:
                # 参考 PyXLL 示例，使用 BindToObject 方法
                log_info('Trying to get object using BindToObject()')
                context = pythoncom.CreateBindCtx(0)
                obj = moniker.BindToObject(context, None, pythoncom.IID_IDispatch)
                log_info('Successfully got object from ROT using BindToObject()')
            except Exception as bind_err:
                log_info(f'Error using BindToObject(): {bind_err}, trying GetObject()')
                try:
                    rot = pythoncom.GetRunningObjectTable()
                    obj = rot.GetObject(moniker)
                    log_info('Successfully got object from ROT using GetObject()')
                except Exception as rot_err:
                    log_info(f'Error using ROT.GetObject(): {rot_err}, trying CoGetObject()')
                    try:
                        obj = pythoncom.CoGetObject(moniker, None)
                        log_info('Successfully got object using CoGetObject()')
                    except Exception as coget_err:
                        log_info(f'Error using CoGetObject(): {coget_err}')
                        log_info('=== ROT MONIKER PROCESSING COMPLETED: ERROR ===')
                        return False, None, None
            
            # 记录获取到的对象信息
            if obj is not None:
                log_info(f'Got object from ROT, type: {type(obj)}')
                log_info(f'Object repr: {repr(obj)}')
                
                # 尝试转换PyIUnknown对象
                try:
                    import win32com.client
                    if hasattr(obj, '__class__') and 'PyIUnknown' in str(type(obj)):
                        log_info('Attempting to convert PyIUnknown object')
                        # 尝试使用Dispatch来转换对象
                        try:
                            obj = win32com.client.Dispatch(obj)
                            log_info('Successfully converted PyIUnknown object using Dispatch')
                            log_info(f'Object type after conversion: {type(obj)}')
                        except Exception as conv_err:
                            log_info(f'Error converting PyIUnknown object: {conv_err}')
                            # 尝试其他转换方法
                            try:
                                obj = pythoncom.CoCreateInstanceFromMoniker(obj)
                                log_info('Successfully converted PyIUnknown object using CoCreateInstanceFromMoniker')
                                log_info(f'Object type after conversion: {type(obj)}')
                            except Exception as conv_err2:
                                log_info(f'Error with CoCreateInstanceFromMoniker: {conv_err2}')
                                # 跳过无法转换的PyIUnknown对象
                                log_info('Skipping PyIUnknown object that cannot be converted')
                                log_info('=== ROT MONIKER PROCESSING COMPLETED: SKIPPED ===')
                                return False, None, None
                except Exception as conv_err:
                    log_info(f'Error in PyIUnknown conversion: {conv_err}')
                
                # 尝试获取对象的一些属性
                try:
                    if hasattr(obj, 'Name'):
                        try:
                            log_info(f'Object Name: {obj.Name}')
                        except Exception as name_err:
                            log_info(f'Error getting Name: {name_err}')
                    if hasattr(obj, 'FullName'):
                        try:
                            log_info(f'Object FullName: {obj.FullName}')
                        except Exception as fullname_err:
                            log_info(f'Error getting FullName: {fullname_err}')
                    if hasattr(obj, 'Application'):
                        log_info(f'Object has Application attribute')
                    # 检查对象的方法和属性
                    log_info(f'Object attributes: {[attr for attr in dir(obj) if not attr.startswith("_")]}')
                except Exception as attr_err:
                    log_info(f'Error getting object attributes: {attr_err}')
            
            # 首先检查是否是文档对象并且是我们要找的文档
            if target_document_path is not None and obj is not None:
                log_info('Checking if object is target document...')
                doc_result = self._check_if_target_document(obj, target_document_path, app_type)
                if doc_result[0]:
                    log_info('Found target document directly in ROT!')
                    log_info('=== ROT MONIKER PROCESSING COMPLETED: FOUND TARGET ===')
                    return doc_result
            
            # 检查是否是 Office 应用程序对象
            if obj is not None:
                log_info('Trying to add object as Office instance...')
                self._try_add_office_instance(obj, app_type, instances)
            
            # 也尝试获取 Application 属性（有些对象可能是文档对象）
            if obj is not None and hasattr(obj, 'Application'):
                try:
                    log_info('Getting Application property from object...')
                    app = obj.Application
                    # 验证Application对象的有效性
                    if app is not None:
                        try:
                            app_name = app.Name
                            log_info(f'Application object validated: {app_name}')
                            self._try_add_office_instance(app, app_type, instances)
                        except Exception as validate_err:
                            log_info(f'Application object validation failed: {validate_err}')
                except Exception as app_err:
                    log_info(f'Error getting Application from object: {app_err}')
                    
            log_info('=== ROT MONIKER PROCESSING COMPLETED: NO TARGET ===')
            return False, None, None
                    
        except Exception as e:
            log_info(f'Error processing ROT moniker: {e}')
            import traceback
            log_info(f'Traceback: {traceback.format_exc()}')
            log_info('=== ROT MONIKER PROCESSING COMPLETED: ERROR ===')
            return False, None, None
    
    def _check_if_target_document(self, obj, target_document_path, app_type):
        """
        检查对象是否是目标文档
        返回: (found: bool, document: object|None, app: object|None)
        """
        log_info('=== STARTING DOCUMENT OBJECT CHECK ===')
        try:
            log_info(f'Checking object type: {type(obj)}')
            log_info(f'Object repr: {repr(obj)}')
            
            # 检查是否有FullName属性（文档对象的特征）
            if hasattr(obj, 'FullName'):
                try:
                    full_name = obj.FullName
                    log_info(f'Checking document from ROT: {full_name}')
                    if _paths_equal(full_name, target_document_path):
                        log_info('FOUND TARGET DOCUMENT IN ROT!')
                        
                        # 优先尝试使用GetObject获取对象
                        log_info('Trying to get object using GetObject first...')
                        try:
                            import win32com.client
                            # 使用GetObject获取对象
                            getobject_obj = win32com.client.GetObject(full_name)
                            log_info(f'Successfully got object using GetObject: {type(getobject_obj)}')
                            
                            # 尝试从GetObject返回的对象获取Application
                            if hasattr(getobject_obj, 'Application'):
                                log_info('GetObject returned object has Application attribute')
                                try:
                                    app = getobject_obj.Application
                                    log_info('Successfully got Application from GetObject returned object')
                                    # 验证Application对象的有效性
                                    if app is not None:
                                        try:
                                            app_name = app.Name
                                            log_info(f'Application object from GetObject validated: {app_name}')
                                            return True, getobject_obj, app
                                        except Exception as validate_err:
                                            log_info(f'Application object validation failed: {validate_err}')
                                except Exception as app_err:
                                    log_info(f'Error getting Application from GetObject returned object: {app_err}')
                            else:
                                log_info('GetObject returned object has no Application attribute')
                        except Exception as getobject_err:
                            log_info(f'Error using GetObject: {getobject_err}')
                        
                        # 如果GetObject失败，尝试转换PyIUnknown对象
                        try:
                            import win32com.client
                            if hasattr(obj, '__class__') and 'PyIUnknown' in str(type(obj)):
                                log_info('Attempting to convert PyIUnknown object in _check_if_target_document')
                                # 尝试使用Dispatch转换
                                try:
                                    obj = win32com.client.Dispatch(obj)
                                    log_info('Successfully converted PyIUnknown object using Dispatch')
                                    log_info(f'Object type after conversion: {type(obj)}')
                                except Exception as conv_err:
                                    log_info(f'Error converting PyIUnknown object: {conv_err}')
                        except Exception as conv_err:
                            log_info(f'Error in PyIUnknown conversion: {conv_err}')
                        
                        # 尝试从文档对象获取Application
                        app = None
                        if hasattr(obj, 'Application'):
                            log_info('Document has Application attribute, trying to get it...')
                            # 尝试获取Application对象
                            try:
                                app = obj.Application
                                log_info('Successfully got Application object')
                                # 验证Application对象的有效性
                                if app is not None:
                                    try:
                                        app_name = app.Name
                                        log_info(f'Application object validated: {app_name}')
                                        return True, obj, app
                                    except Exception as validate_err:
                                        log_info(f'Application object validation failed: {validate_err}')
                                        app = None
                            except Exception as app_err:
                                log_info(f'Error getting Application from document object: {app_err}')
                        else:
                            log_info('Found target document but no Application attribute')
                        
                        # 即使获取Application失败，也返回文档对象
                        return True, obj, None
                except Exception as name_err:
                    log_info(f'Error getting FullName from object: {name_err}')
            
            # 尝试其他属性检查
            if hasattr(obj, 'Name'):
                try:
                    name = obj.Name
                    log_info(f'Checking object Name: {name}')
                    # 检查Name是否包含目标文档的文件名
                    target_filename = os.path.basename(target_document_path)
                    if target_filename.lower() in name.lower():
                        log_info('Object Name matches target filename!')
                        # 获取Application对象
                        app = None
                        if hasattr(obj, 'Application'):
                            log_info('Object has Application attribute, trying to get it...')
                            try:
                                app = obj.Application
                                log_info('Successfully got Application object')
                            except Exception as app_err:
                                log_info(f'Error getting Application from object: {app_err}')
                        return True, obj, app
                except Exception as name_err:
                    log_info(f'Error getting Name from object: {name_err}')
            
            # 尝试直接检查对象是否有文档特征属性
            has_document_features = False
            if app_type == 'excel':
                has_document_features = hasattr(obj, 'Worksheets') or hasattr(obj, 'Sheets')
            elif app_type == 'word':
                has_document_features = hasattr(obj, 'Paragraphs') or hasattr(obj, 'Sections')
            elif app_type == 'powerpoint':
                has_document_features = hasattr(obj, 'Slides') or hasattr(obj, 'Shapes')
            
            if has_document_features:
                log_info(f'Object has document features for {app_type}')
                # 尝试获取Application对象
                app = None
                if hasattr(obj, 'Application'):
                    log_info('Document has Application attribute, trying to get it...')
                    try:
                        app = obj.Application
                        log_info('Successfully got Application object')
                    except Exception as app_err:
                        log_info(f'Error getting Application from document object: {app_err}')
                return True, obj, app
            
            log_info('Object is not a target document')
            log_info('=== DOCUMENT OBJECT CHECK COMPLETED: NOT TARGET ===')
            return False, None, None
        except Exception as e:
            log_info(f'Error checking if object is target document: {e}')
            import traceback
            log_info(f'Traceback: {traceback.format_exc()}')
            log_info('=== DOCUMENT OBJECT CHECK COMPLETED: ERROR ===')
            return False, None, None
    
    def _try_add_office_instance(self, obj, app_type, instances):
        """
        尝试添加 Office 应用实例到列表中（如果还没有）
        """
        try:
            # 检查是否有 Name 属性
            if not hasattr(obj, 'Name'):
                return
            
            app_name = obj.Name.lower()
            
            # 检查是否是目标应用类型
            if app_type not in app_name:
                return
                
            # 避免重复添加同一个实例
            is_duplicate = False
            obj_hwnd = None
            
            # 尝试获取 HWND 用于去重
            try:
                if hasattr(obj, 'Hwnd'):
                    obj_hwnd = obj.Hwnd
            except Exception:
                pass
                
            for instance in instances:
                try:
                    if instance is obj:
                        is_duplicate = True
                        break
                    if obj_hwnd is not None and hasattr(instance, 'Hwnd'):
                        if instance.Hwnd == obj_hwnd:
                            is_duplicate = True
                            break
                except Exception:
                    pass
                    
            if not is_duplicate:
                # 记录找到的实例信息
                doc_count = self._get_document_count(obj, app_type)
                log_info(f'Found additional {app_type} app instance with {doc_count} open documents')
                
                # 列出该实例中打开的所有文档（用于调试）
                self._list_open_documents(obj, app_type)
                
                instances.append(obj)
                
        except Exception as e:
            log_info(f'Error trying to add Office instance: {e}')
    
    def _list_open_documents(self, app, app_type):
        """列出应用实例中打开的所有文档（用于调试）"""
        try:
            if app_type == 'excel':
                for workbook in app.Workbooks:
                    try:
                        log_info(f'  Open workbook: {workbook.FullName}')
                    except Exception:
                        pass
            elif app_type == 'powerpoint':
                for presentation in app.Presentations:
                    try:
                        log_info(f'  Open presentation: {presentation.FullName}')
                    except Exception:
                        pass
            elif app_type == 'word':
                for doc in app.Documents:
                    try:
                        log_info(f'  Open document: {doc.FullName}')
                    except Exception:
                        pass
        except Exception as e:
            log_info(f'Error listing open documents: {e}')
    
    def _try_find_open_document(self, document_path, app_type):
        """
        尝试查找已打开的文档
        返回: (found: bool, document: object|None, app: object|None)
        """
        log_info(f'=== STARTING DOCUMENT SEARCH: {document_path} (app_type: {app_type}) ===')
        try:
            # 方式 1: GetActiveObject + 在该实例中查找
            prog_id = {
                'excel': 'Excel.Application',
                'powerpoint': 'PowerPoint.Application',
                'word': 'Word.Application'
            }.get(app_type)

            if prog_id:
                # 尝试获取所有正在运行的 Office 实例
                instances = []
                
                # 首先尝试 GetActiveObject 获取当前激活的实例
                log_info('Step 1: Trying GetActiveObject for active instance')
                try:
                    app = win32com.client.GetActiveObject(prog_id)
                    log_info(f'✓ Found active {app_type} app instance')
                    instances.append(app)
                    
                    # 在激活实例中查找文档
                    log_info('Step 1.1: Searching document in active app instance')
                    doc = self._find_document_in_app(app, document_path, app_type)
                    if doc:
                        log_info('✓ Found document in active app instance!')
                        log_info('=== DOCUMENT SEARCH COMPLETED: FOUND IN ACTIVE INSTANCE ===')
                        return True, doc, app
                except Exception as e:
                    log_info(f'✗ No active {app_type} app: {e}')
                
                # 尝试使用 ROT (Running Object Table) 枚举所有实例
                log_info('Step 2: Enumerating ROT to collect all Office instances')
                try:
                    import pythoncom
                    from win32com.client import Dispatch
                    
                    # 初始化 COM
                    pythoncom.CoInitialize()
                    
                    # 获取 Running Object Table
                    rot = pythoncom.GetRunningObjectTable()
                    enum = rot.EnumRunning()
                    
                    # 遍历 ROT 中的所有对象，优先收集 Office 实例
                    log_info('Enumerating ROT to collect Office instances...')
                    
                    # 参考 PyXLL 示例的遍历方式
                    context = pythoncom.CreateBindCtx(0)
                    
                    # 尝试不同的遍历方式
                    try:
                        # 方式1：使用 Next() 方法
                        log_info('Method 1: Using enum.Next()')
                        while True:
                            try:
                                # 调用 Next(1)，返回的是 (moniker_array, num_fetched)
                                result = enum.Next(1)
                                log_info(f'ROT Next() result: {result}, type: {type(result)}')
                                
                                # 检查是否有获取到对象
                                if not result:
                                    break
                                    
                                if isinstance(result, tuple):
                                    # 处理不同格式的返回值
                                    if len(result) == 2:
                                        # 标准格式：(monikers, num_fetched)
                                        monikers, num_fetched = result
                                        if num_fetched == 0:
                                            break
                                        
                                        log_info(f'ROT Next() fetched {num_fetched} moniker(s)')
                                        
                                        for moniker in monikers:
                                            try:
                                                # 先尝试添加为 Office 实例
                                                self._process_rot_moniker(moniker, app_type, instances, None)
                                            except Exception as moniker_err:
                                                log_info(f'Error processing moniker: {moniker_err}')
                                    elif len(result) == 1:
                                        # 单个元素，可能是 moniker
                                        element = result[0]
                                        try:
                                            if hasattr(element, 'GetDisplayName'):
                                                # 单个 moniker
                                                log_info(f'Processing single moniker: {element}')
                                                self._process_rot_moniker(element, app_type, instances, None)
                                            elif hasattr(element, '__iter__') and not isinstance(element, (str, bytes)):
                                                # moniker 数组
                                                for moniker in element:
                                                    try:
                                                        self._process_rot_moniker(moniker, app_type, instances, None)
                                                    except Exception as moniker_err:
                                                        log_info(f'Error processing moniker: {moniker_err}')
                                        except Exception as element_err:
                                            log_info(f'Error processing element: {element_err}')
                                else:
                                    # 非元组结果，可能是单个 moniker
                                    log_info(f'ROT Next() returned non-tuple: {result}')
                                    try:
                                        if hasattr(result, 'GetDisplayName'):
                                            self._process_rot_moniker(result, app_type, instances, None)
                                    except Exception as single_err:
                                        log_info(f'Error processing single moniker: {single_err}')
                            except StopIteration:
                                break
                            except Exception as e:
                                log_info(f'Error in ROT iteration: {e}')
                                # 继续尝试其他方式
                                break
                    except Exception as next_err:
                        log_info(f'Error with Next() method: {next_err}')
                        
                    # 方式2：使用迭代器方式（如果支持）
                    try:
                        log_info('Method 2: Using iterator')
                        for moniker in rot:
                            try:
                                self._process_rot_moniker(moniker, app_type, instances, None)
                            except Exception as moniker_err:
                                log_info(f'Error processing moniker: {moniker_err}')
                    except Exception as iter_err:
                        log_info(f'Error with iterator: {iter_err}')
                except Exception as e:
                    log_info(f'Error enumerating ROT: {e}')
                
                # 优先检查所有收集到的 Office 实例
                log_info(f'Step 3: Checking {len(instances)} collected app instances (priority check)')
                for i, app in enumerate(instances):
                    try:
                        log_info(f'Checking instance {i+1}/{len(instances)}')
                        # 在这个实例中查找文档
                        doc = self._find_document_in_app(app, document_path, app_type)
                        if doc:
                            log_info('✓ Found document in collected app instance!')
                            log_info('=== DOCUMENT SEARCH COMPLETED: FOUND IN COLLECTED INSTANCE ===')
                            return True, doc, app
                    except Exception as e:
                        log_info(f'Error checking app instance {i+1}: {e}')
                
                # 只有在实例中找不到文档时，才尝试从 ROT 中直接查找文档
                log_info('Step 4: No document found in app instances, trying direct ROT document lookup...')
                try:
                    import pythoncom
                    from win32com.client import Dispatch
                    
                    # 初始化 COM
                    pythoncom.CoInitialize()
                    
                    # 获取 Running Object Table
                    rot = pythoncom.GetRunningObjectTable()
                    enum = rot.EnumRunning()
                    
                    # 遍历 ROT 中的所有对象，直接查找文档
                    log_info('Enumerating ROT for direct document lookup...')
                    while True:
                        try:
                            # 调用 Next(1)，返回的是 (moniker_array, num_fetched)
                            result = enum.Next(1)
                            
                            # 检查是否有获取到对象
                            if not result:
                                break
                                
                            if isinstance(result, tuple):
                                # 检查元组的第一个元素是否是 PyIMoniker 类型
                                if len(result) >= 1 and hasattr(result[0], 'GetDisplayName'):
                                    # 这是直接包含 moniker 的情况
                                    moniker = result[0]
                                    try:
                                        found, doc, app = self._process_rot_moniker(
                                            moniker, app_type, [], document_path
                                        )
                                        if found:
                                            log_info('✓ Found document directly in ROT!')
                                            log_info('=== DOCUMENT SEARCH COMPLETED: FOUND IN ROT ===')
                                            return True, doc, app
                                    except Exception as moniker_err:
                                        log_info(f'Error processing moniker: {moniker_err}')
                                elif len(result) >= 2:
                                    # 标准格式：(monikers, num_fetched)
                                    monikers, num_fetched = result
                                    if num_fetched == 0:
                                        break
                                    
                                    for moniker in monikers:
                                        try:
                                            found, doc, app = self._process_rot_moniker(
                                                moniker, app_type, [], document_path
                                            )
                                            if found:
                                                log_info('✓ Found document directly in ROT!')
                                                log_info('=== DOCUMENT SEARCH COMPLETED: FOUND IN ROT ===')
                                                return True, doc, app
                                        except Exception as moniker_err:
                                            log_info(f'Error processing moniker: {moniker_err}')
                                elif len(result) == 1:
                                    # 只有一个元素，可能是 moniker 数组或 num_fetched
                                    element = result[0]
                                    try:
                                        # 尝试作为 moniker 数组处理
                                        if hasattr(element, '__iter__') and not isinstance(element, (str, bytes)):
                                            for moniker in element:
                                                try:
                                                    found, doc, app = self._process_rot_moniker(
                                                        moniker, app_type, [], document_path
                                                    )
                                                    if found:
                                                        log_info('✓ Found document directly in ROT!')
                                                        log_info('=== DOCUMENT SEARCH COMPLETED: FOUND IN ROT ===')
                                                        return True, doc, app
                                                except Exception as moniker_err:
                                                    log_info(f'Error processing moniker: {moniker_err}')
                                        else:
                                            # 单个 moniker
                                            found, doc, app = self._process_rot_moniker(
                                                element, app_type, [], document_path
                                            )
                                            if found:
                                                log_info('✓ Found document directly in ROT!')
                                                log_info('=== DOCUMENT SEARCH COMPLETED: FOUND IN ROT ===')
                                                return True, doc, app
                                    except Exception as element_err:
                                        log_info(f'Error processing single element: {element_err}')
                                else:
                                    # 空元组，结束
                                    break
                            else:
                                # 非元组结果，可能是单个 moniker
                                try:
                                    found, doc, app = self._process_rot_moniker(
                                        result, app_type, [], document_path
                                    )
                                    if found:
                                        log_info('✓ Found document directly in ROT!')
                                        log_info('=== DOCUMENT SEARCH COMPLETED: FOUND IN ROT ===')
                                        return True, doc, app
                                except Exception as single_err:
                                    log_info(f'Error processing single moniker: {single_err}')
                                    break
                        except StopIteration:
                            break
                        except Exception as e:
                            log_info(f'Error in ROT document lookup: {e}')
                            break
                except Exception as e:
                    log_info(f'Error in direct ROT document lookup: {e}')

            # 未来可以添加其他查找方式

            log_info('=== DOCUMENT SEARCH COMPLETED: NOT FOUND ===')
            return False, None, None

        except Exception as e:
            log_info(f'Error in _try_find_open_document: {e}')
            import traceback
            log_info(f'Traceback: {traceback.format_exc()}')
            log_info('=== DOCUMENT SEARCH COMPLETED: ERROR ===')
            return False, None, None

    def _find_document_in_app(self, app, document_path, app_type):
        """在指定的应用实例中查找目标文档"""
        log_info(f'Searching for document {document_path} in {app_type} app instance...')
        log_info(f'App object type: {type(app)}')
        log_info(f'App object repr: {repr(app)}')

        try:
            if app_type == 'excel':
                try:
                    workbooks_count = app.Workbooks.Count
                    log_info(f'Excel Workbooks.Count: {workbooks_count}')
                    
                    # 尝试使用不同的遍历方式
                    log_info('Attempting to iterate through workbooks...')
                    
                    # 方法1：直接遍历
                    try:
                        log_info('Method 1: Direct iteration')
                        for i in range(1, workbooks_count + 1):
                            try:
                                workbook = app.Workbooks(i)
                                full_name = workbook.FullName
                                log_info(f'  Workbook {i}: {full_name}')
                                if _paths_equal(full_name, document_path):
                                    log_info('  MATCH FOUND!')
                                    return workbook
                            except Exception as wb_err:
                                log_info(f'  Error accessing workbook {i}: {wb_err}')
                    except Exception as iter_err:
                        log_info(f'Error with direct iteration: {iter_err}')
                        
                    # 方法2：for循环遍历
                    try:
                        log_info('Method 2: For loop iteration')
                        for workbook in app.Workbooks:
                            try:
                                full_name = workbook.FullName
                                log_info(f'  Found workbook: {full_name}')
                                if _paths_equal(full_name, document_path):
                                    log_info('  MATCH FOUND!')
                                    return workbook
                            except Exception as wb_err:
                                log_info(f'  Error checking workbook: {wb_err}')
                    except Exception as for_err:
                        log_info(f'Error with for loop iteration: {for_err}')
                except Exception as count_err:
                    log_info(f'Error getting Workbooks.Count: {count_err}')

            elif app_type == 'powerpoint':
                try:
                    presentations_count = app.Presentations.Count
                    log_info(f'PowerPoint Presentations.Count: {presentations_count}')
                    for presentation in app.Presentations:
                        try:
                            full_name = presentation.FullName
                            log_info(f'  Found presentation: {full_name}')
                            if _paths_equal(full_name, document_path):
                                log_info('  MATCH FOUND!')
                                return presentation
                        except Exception as ppt_err:
                            log_info(f'  Error checking presentation: {ppt_err}')
                except Exception as count_err:
                    log_info(f'Error getting Presentations.Count: {count_err}')

            elif app_type == 'word':
                try:
                    documents_count = app.Documents.Count
                    log_info(f'Word Documents.Count: {documents_count}')
                    for doc in app.Documents:
                        try:
                            full_name = doc.FullName
                            log_info(f'  Found document: {full_name}')
                            if _paths_equal(full_name, document_path):
                                log_info('  MATCH FOUND!')
                                return doc
                        except Exception as doc_err:
                            log_info(f'  Error checking document: {doc_err}')
                except Exception as count_err:
                    log_info(f'Error getting Documents.Count: {count_err}')

            log_info('Document NOT found in this app instance')
            return None

        except Exception as e:
            log_error(f'Error in _find_document_in_app: {e}')
            import traceback
            log_info(f'Traceback: {traceback.format_exc()}')
            return None

    def _is_connection_valid(self, document_path):
        """检查文档连接是否有效"""
        try:
            if document_path not in self.documents:
                return False
            
            doc = self.documents[document_path]
            # 尝试访问文档的基本属性来验证连接
            if hasattr(doc, 'FullName'):
                doc.FullName
            elif hasattr(doc, 'Name'):
                doc.Name
            else:
                # 对于没有这些属性的对象，尝试其他方式验证
                pass
            return True
        except Exception as e:
            log_info(f'Connection validation failed: {e}')
            return False

    def _remove_invalid_connection(self, document_path):
        """移除无效的连接"""
        if document_path in self.documents:
            del self.documents[document_path]
        if document_path in self.document_connections:
            del self.document_connections[document_path]
        log_info(f'Removed invalid connection for document: {document_path}')

    def _open_new_document(self, document_path, app_type, visible=False):
        """打开一个新的文档"""
        try:
            # 获取或创建应用实例
            app = self._get_or_create_app(app_type)

            # 处理可见性
            if app_type == 'powerpoint':
                log_info('PowerPoint: forcing Visible=True')
                app.Visible = True
            else:
                # 始终保持窗口可见
                app.Visible = True
            log_info(f'App visible set to: {app.Visible}')

            # 打开文档（以读写模式打开，支持从IDE同步代码到文档）
            log_info(f'Opening document from file: {document_path}')
            if app_type == 'excel':
                workbook = app.Workbooks.Open(document_path)
                self.documents[document_path] = workbook
            elif app_type == 'powerpoint':
                # PowerPoint 特殊处理
                try:
                    presentation = app.Presentations.Open(document_path, ReadOnly=False)
                except Exception:
                    presentation = app.Presentations.Open(document_path, WithWindow=True, ReadOnly=False)
                self.documents[document_path] = presentation
            elif app_type == 'word':
                # 始终保持窗口可见
                doc = app.Documents.Open(document_path, ReadOnly=False, AddToRecentFiles=False, Visible=True)
                self.documents[document_path] = doc
            else:
                return {'success': False, 'error': f'Unknown app type: {app_type}'}

            log_info(f'Opened document successfully: {document_path}')
            return {'success': True, 'message': 'Document opened successfully'}

        except Exception as e:
            log_error(f'Error opening new document: {e}')
            return {'success': False, 'error': str(e)}

    def check_vba_access(self, document_path, syncDirection=None):
        """检查 VBA 工程访问权限"""
        try:
            if document_path not in self.documents:
                return {'success': False, 'error': 'Document not open'}

            # 检查连接是否有效
            if not self._is_connection_valid(document_path):
                log_info('Invalid connection detected, removing and reconnecting')
                self._remove_invalid_connection(document_path)
                return {'success': False, 'error': 'Connection invalid, please reconnect'}

            document = self.documents[document_path]
            log_info(f'Document object type: {type(document)}')
            
            # 尝试转换PyIUnknown对象
            try:
                import win32com.client
                if hasattr(document, '__class__') and 'PyIUnknown' in str(type(document)):
                    log_info('Attempting to convert PyIUnknown document object')
                    # 尝试使用Dispatch来转换对象
                    document = win32com.client.Dispatch(document)
                    log_info(f'Document object type after conversion: {type(document)}')
                    # 更新文档对象
                    self.documents[document_path] = document
            except Exception as conv_err:
                log_info(f'Error converting document object: {conv_err}')

            vba_project = document.VBProject

            # 尝试访问 VBA 工程属性
            name = vba_project.Name

            # 尝试访问 VBComponents 来检测是否被保护
            try:
                components = vba_project.VBComponents
                return {'success': True, 'data': {'projectName': name}}
            except Exception as components_error:
                log_error(f'VBA project is protected: {components_error}')
                return {'success': False, 'error': str(components_error)}

        except Exception as e:
            log_error(f'Error checking VBA access: {e}')
            error_str = str(e)

            # 检查是否是连接失效错误
            if '-2147417848' in error_str or '被调用的对象已与其客户端断开连接' in error_str:
                log_info('Connection lost error detected, removing invalid connection')
                self._remove_invalid_connection(document_path)
                return {'success': False, 'error': 'Connection lost, please reconnect'}

            # 检查是否是PyIUnknown对象错误
            if 'PyIUnknown' in error_str:
                log_info('PyIUnknown object error detected')
                # 尝试重新获取文档
                if document_path in self.documents:
                    del self.documents[document_path]
                    log_info(f'Removed problematic document: {document_path}')
                return {'success': False, 'error': 'Document object needs to be reconnected'}

            # 检查是否是密码保护错误
            is_password_protected = ('该工程已被保护' in error_str or
                                    '工程已被保护' in error_str or
                                    'project is protected' in error_str.lower())

            # 如果不是密码保护错误，检查是否是Office实例已关闭的错误
            if not is_password_protected:
                is_office_closed = ('RPC' in error_str or
                                 '服务器不可用' in error_str or
                                 'Object does not exist' in error_str or
                                 '对象不存在' in error_str)

                if is_office_closed:
                    if document_path in self.documents:
                        self._remove_invalid_connection(document_path)
                        log_info(f'Removed closed document: {document_path}')

            return {'success': False, 'error': error_str}

    def set_window_visible(self, document_path, visible):
        """控制窗口可见性"""
        try:
            if document_path not in self.documents:
                return {'success': False, 'error': 'Document not open'}

            document = self.documents[document_path]
            app = self._get_app_from_document(document)
            if app:
                app.Visible = visible
                return {'success': True, 'message': f'Window visibility set to {visible}'}
            return {'success': False, 'error': 'Could not get application instance'}

        except Exception as e:
            log_error(f'Error setting window visibility: {e}')
            return {'success': False, 'error': str(e)}

    def save_document(self, document_path):
        """保存文档"""
        try:
            if document_path not in self.documents:
                return {'success': False, 'error': 'Document not open'}

            document = self.documents[document_path]

            try:
                document.Save()
            except Exception as e:
                if 'Read-only' in str(e):
                    document.SaveAs(document_path, ReadOnlyRecommended=False)
                else:
                    raise

            log_info(f'Saved document: {document_path}')
            return {'success': True, 'message': 'Document saved successfully'}

        except Exception as e:
            log_error(f'Error saving document: {e}')
            return {'success': False, 'error': str(e)}

    def close_document(self, document_path):
        """关闭文档"""
        try:
            if document_path not in self.documents:
                # 移除文档连接状态标记
                self.remove_document_connection(document_path)
                return {'success': True, 'message': 'Document already closed'}

            document = self.documents[document_path]
            document.Close()
            del self.documents[document_path]

            # 移除文档连接状态标记
            self.remove_document_connection(document_path)

            # 检查是否还有其他文档打开
            app_type = self._get_app_type_from_document(document)
            if app_type and app_type in self.apps:
                app = self.apps[app_type]
                if self._get_document_count(app, app_type) == 0:
                    app.Quit()
                    del self.apps[app_type]

            log_info(f'Closed document: {document_path}')
            return {'success': True, 'message': 'Document closed successfully'}

        except Exception as e:
            log_error(f'Error closing document: {e}')
            # 即使出错，也尝试移除文档连接状态标记
            self.remove_document_connection(document_path)
            return {'success': False, 'error': str(e)}

    def _get_or_create_app(self, app_type):
        """获取或创建 Office 应用实例"""
        if app_type not in self.apps:
            prog_id = {
                'excel': 'Excel.Application',
                'powerpoint': 'PowerPoint.Application',
                'word': 'Word.Application'
            }.get(app_type)
            if not prog_id:
                raise ValueError(f'Unknown app type: {app_type}')
            self.apps[app_type] = win32com.client.Dispatch(prog_id)
        return self.apps[app_type]

    def _get_app_from_document(self, document):
        """从文档对象获取应用实例"""
        try:
            if hasattr(document, 'Application'):
                return document.Application
            return None
        except Exception:
            return None

    def _get_app_type_from_document(self, document):
        """从文档对象获取应用类型"""
        try:
            app = self._get_app_from_document(document)
            if app:
                app_name = app.Name.lower()
                if 'excel' in app_name:
                    return 'excel'
                elif 'powerpoint' in app_name:
                    return 'powerpoint'
                elif 'word' in app_name:
                    return 'word'
            return None
        except Exception:
            return None

    def _get_document_count(self, app, app_type):
        """获取应用中打开的文档数量"""
        try:
            if app_type == 'excel':
                return app.Workbooks.Count
            elif app_type == 'powerpoint':
                return app.Presentations.Count
            elif app_type == 'word':
                return app.Documents.Count
            return 0
        except Exception:
            return 0

    def is_document_connection(self, document_path):
        """检查文档是否是文档连接状态"""
        return document_path in self.document_connections and self.document_connections[document_path]

    def remove_document_connection(self, document_path):
        """移除文档连接状态标记"""
        if document_path in self.document_connections:
            del self.document_connections[document_path]
