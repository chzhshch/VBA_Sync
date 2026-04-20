#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
VBA 模块导入器
处理 VBA 模块的导入功能
"""

import os
import tempfile
from .utils import log_info, log_error, extract_module_name, _filter_attributes

class VBAImporter:
    def __init__(self, connector):
        self.connector = connector

    def import_module(self, document_path, file_path, module_name=None):
        """导入单个模块"""
        try:
            if document_path not in self.connector.documents:
                return {'success': False, 'error': 'Document not open'}

            document = self.connector.documents[document_path]
            vba_project = document.VBProject
            components = vba_project.VBComponents

            # 读取文件内容 - 处理中文路径
            try:
                # 尝试直接打开文件
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
            except UnicodeDecodeError:
                # 尝试使用gbk编码
                with open(file_path, 'r', encoding='gbk') as f:
                    content = f.read()
            except Exception as e:
                # 如果打开文件失败，尝试使用临时文件策略
                log_error(f'Error opening file: {e}')
                return {'success': False, 'error': f'无法读取文件: {str(e)}'}

            # 检查是否为文档模块
            is_document_module = self._is_document_module(content)
            if is_document_module:
                # 移除文档模块标记
                content = content.replace('@vba-sync-document-module\n', '', 1)

            # 提取模块名
            if not module_name:
                try:
                    # 尝试从文件名提取模块名
                    file_basename = os.path.basename(file_path)
                    module_name = extract_module_name(content, file_basename)
                except Exception as e:
                    log_error(f'Error extracting module name: {e}')
                    return {'success': False, 'error': f'无法提取模块名称: {str(e)}'}

            # 选择导入策略
            if is_document_module:
                # 策略 D: 安全更新文档模块
                result = self._import_strategy_d(components, module_name, content)
            else:
                # 检查模块是否存在
                existing_component = None
                for comp in components:
                    if comp.Name == module_name:
                        existing_component = comp
                        break

                if existing_component:
                    # 检查文件类型
                    if file_path.endswith('.bas'):
                        # 策略 A: 直接更新标准模块
                        result = self._import_strategy_a(existing_component, content)
                    else:
                        # 策略 A: 直接更新类模块
                        # 原因：避免临时文件导入策略（策略B）中的编码问题
                        # 临时文件导入策略的优缺点：
                        # 优点：完整保留类模块元数据（VERSION头部、Attribute标记等）
                        # 缺点：临时文件写入/读取过程中存在编码转换问题，导致中文注释变成乱码
                        # 直接更新策略的优缺点：
                        # 优点：避免编码问题，操作更直接，实时性更好
                        # 缺点：可能丢失类模块元数据
                        # 权衡：对于中文用户，编码正确性比元数据完整性更重要
                        result = self._import_strategy_a(existing_component, content)
                else:
                    # 策略 C: 创建新模块
                    if file_path.endswith('.bas'):
                        # 标准模块
                        result = self._import_strategy_c(components, module_name, content, 1)
                    else:
                        # 类模块
                        result = self._import_strategy_c(components, module_name, content, 2)

            log_info(f'Imported module: {module_name} from {file_path}')
            return result

        except Exception as e:
            log_error(f'Error importing module: {e}')
            return {'success': False, 'error': str(e)}

    def import_all(self, document_path, directory):
        """导入目录中所有模块"""
        try:
            if document_path not in self.connector.documents:
                return {'success': False, 'error': 'Document not open'}

            imported_modules = []

            for file_name in os.listdir(directory):
                if file_name.endswith(('.bas', '.cls')):
                    file_path = os.path.join(directory, file_name)
                    result = self.import_module(document_path, file_path)
                    if result['success']:
                        imported_modules.append(file_name)

            return {'success': True, 'imported': imported_modules}

        except Exception as e:
            log_error(f'Error importing all modules: {e}')
            return {'success': False, 'error': str(e)}

    def delete_module(self, document_path, module_name):
        """删除模块"""
        try:
            if document_path not in self.connector.documents:
                return {'success': False, 'error': 'Document not open'}

            document = self.connector.documents[document_path]
            vba_project = document.VBProject
            components = vba_project.VBComponents

            component = None
            for comp in components:
                if comp.Name == module_name:
                    component = comp
                    break

            if not component:
                return {'success': False, 'error': f'Module {module_name} not found'}

            # 不能删除文档模块
            if component.Type == 100:
                return {'success': False, 'error': 'Cannot delete document modules'}

            components.Remove(component)
            log_info(f'Deleted module: {module_name}')
            return {'success': True, 'message': f'Module {module_name} deleted'}

        except Exception as e:
            log_error(f'Error deleting module: {e}')
            return {'success': False, 'error': str(e)}

    def clear_module_code(self, document_path, module_name):
        """清空模块代码"""
        try:
            if document_path not in self.connector.documents:
                return {'success': False, 'error': 'Document not open'}

            document = self.connector.documents[document_path]
            vba_project = document.VBProject
            components = vba_project.VBComponents

            component = None
            for comp in components:
                if comp.Name == module_name:
                    component = comp
                    break

            if not component:
                return {'success': False, 'error': f'Module {module_name} not found'}

            # 清空代码
            code_module = component.CodeModule
            line_count = code_module.CountOfLines
            if line_count > 0:
                code_module.DeleteLines(1, line_count)
            log_info(f'Cleared code for module: {module_name}')
            return {'success': True, 'message': f'Module {module_name} code cleared'}

        except Exception as e:
            log_error(f'Error clearing module code: {e}')
            return {'success': False, 'error': str(e)}

    def _import_strategy_a(self, component, content):
        """策略 A: 直接更新"""
        try:
            # 过滤属性
            filtered_content = _filter_attributes(content)
            
            code_module = component.CodeModule
            line_count = code_module.CountOfLines
            if line_count > 0:
                code_module.DeleteLines(1, line_count)
            code_module.AddFromString(filtered_content)
            return {'success': True, 'message': 'Module updated using strategy A'}
        except Exception as e:
            log_error(f'Error using strategy A: {e}')
            raise e

    def _import_strategy_b(self, components, module_name, content):
        """策略 B: 临时文件导入"""
        try:
            # 检查是否存在
            existing_component = None
            for comp in components:
                if comp.Name == module_name:
                    existing_component = comp
                    break

            # 删除现有模块
            if existing_component:
                components.Remove(existing_component)

            # 使用临时文件导入
            with tempfile.NamedTemporaryFile(suffix='.cls', delete=False) as temp_file:
                temp_path = temp_file.name

            with open(temp_path, 'w', encoding='utf-8') as f:
                f.write(content)

            # 导入模块
            new_component = components.Import(temp_path)
            # 确保模块名正确
            if new_component.Name != module_name:
                new_component.Name = module_name

            # 清理临时文件
            os.unlink(temp_path)

            return {'success': True, 'message': 'Module imported using strategy B'}
        except Exception as e:
            # 清理临时文件
            if 'temp_path' in locals() and os.path.exists(temp_path):
                try:
                    os.unlink(temp_path)
                except:
                    pass
            log_error(f'Error using strategy B: {e}')
            raise e

    def _import_strategy_c(self, components, module_name, content, module_type):
        """策略 C: 创建新模块"""
        try:
            # 创建新模块
            new_component = components.Add(module_type)
            new_component.Name = module_name

            # 过滤属性
            filtered_content = _filter_attributes(content)

            # 添加代码
            code_module = new_component.CodeModule
            code_module.AddFromString(filtered_content)

            return {'success': True, 'message': 'Module created using strategy C'}
        except Exception as e:
            log_error(f'Error using strategy C: {e}')
            raise e

    def _import_strategy_d(self, components, module_name, content):
        """策略 D: 安全更新文档模块"""
        try:
            # 找到文档模块
            component = None
            for comp in components:
                if comp.Name == module_name and comp.Type == 100:
                    component = comp
                    break

            if not component:
                return {'success': False, 'error': f'Document module {module_name} not found'}

            # 更新代码
            code_module = component.CodeModule
            line_count = code_module.CountOfLines
            if line_count > 0:
                code_module.DeleteLines(1, line_count)
            code_module.AddFromString(content)

            return {'success': True, 'message': 'Document module updated using strategy D'}
        except Exception as e:
            log_error(f'Error using strategy D: {e}')
            raise e

    def _is_document_module(self, content):
        """检查是否为文档模块"""
        return content.startswith('@vba-sync-document-module\n')
