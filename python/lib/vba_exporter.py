#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
VBA 模块导出器
处理 VBA 模块的导出功能
"""

import os
import tempfile
from .utils import log_info, log_error, extract_module_name

class VBAExporter:
    def __init__(self, connector):
        self.connector = connector

    def export_all(self, document_path, output_dir):
        """导出所有 VBA 模块"""
        try:
            if document_path not in self.connector.documents:
                return {'success': False, 'error': 'Document not open'}

            document = self.connector.documents[document_path]
            vba_project = document.VBProject
            
            # 首先尝试访问VBProject以检查是否被锁定
            try:
                project_name = vba_project.Name
                log_info(f'VBA Project name: {project_name}')
            except Exception as access_error:
                log_error(f'VBA project locked: {access_error}')
                return {'success': False, 'error': str(access_error)}
            
            components = vba_project.VBComponents
            component_count = components.Count
            log_info(f'Total VBComponents found: {component_count}')

            # 确保输出目录存在
            os.makedirs(output_dir, exist_ok=True)

            exported_modules = []

            for i in range(component_count):
                try:
                    component = components.Item(i + 1)  # COM集合从1开始索引
                    module_name = component.Name
                    module_type = component.Type
                    log_info(f'Processing component {i+1}/{component_count}: {module_name}, Type: {module_type}')

                    # 确定文件扩展名
                    if module_type == 1:  # 标准模块
                        ext = '.bas'
                    elif module_type == 2:  # 类模块
                        ext = '.cls'
                    elif module_type == 100:  # 文档模块
                        ext = '.cls'
                    else:
                        # 跳过不支持的类型（如窗体）
                        log_info(f'Skipping unsupported module type: {module_type}')
                        continue

                    file_name = f'{module_name}{ext}'
                    file_path = os.path.join(output_dir, file_name)
                    log_info(f'Exporting to: {file_path}')

                    if module_type == 100:  # 文档模块
                        # 特殊处理文档模块
                        code = self._export_document_module(component)
                        with open(file_path, 'w', encoding='utf-8') as f:
                            f.write(code)
                    else:
                        # 标准模块和类模块
                        self._export_component(component, file_path)

                    exported_modules.append(file_name)
                    log_info(f'Successfully exported module: {module_name} to {file_path}')

                except Exception as e:
                    log_error(f'Error exporting component {i+1}: {e}')
                    import traceback
                    log_error(f'Traceback: {traceback.format_exc()}')
                    continue

            log_info(f'Exported {len(exported_modules)} modules out of {component_count} components')
            if not exported_modules:
                return {'success': False, 'error': 'No modules exported. This may be due to VBA project password protection.'}
            return {'success': True, 'exported': exported_modules}

        except Exception as e:
            log_error(f'Error exporting all modules: {e}')
            import traceback
            log_error(f'Traceback: {traceback.format_exc()}')
            return {'success': False, 'error': str(e)}

    def _export_component(self, component, output_path):
        """导出单个组件"""
        try:
            # 使用临时文件导出
            with tempfile.NamedTemporaryFile(suffix='.bas', delete=False) as temp_file:
                temp_path = temp_file.name

            component.Export(temp_path)
            log_info(f'Exported to temp file: {temp_path}')

            # 尝试多种编码读取文件
            encodings = ['utf-8-sig', 'utf-8', 'gbk', 'gb2312', 'gb18030', 'cp1252']
            content = None
            used_encoding = None
            
            for encoding in encodings:
                try:
                    log_info(f'Trying to read temp file with encoding: {encoding}')
                    with open(temp_path, 'r', encoding=encoding) as f:
                        content = f.read()
                    used_encoding = encoding
                    log_info(f'Successfully read file with encoding: {encoding}')
                    break
                except UnicodeDecodeError as e:
                    log_info(f'Failed with encoding {encoding}: {e}')
                    continue
            
            if content is None:
                raise Exception(f'Failed to read file with any encoding. Tried: {encodings}')

            # 过滤掉VBE自动生成的属性信息
            # 对于类模块，过滤VERSION 1.0 CLASS部分和Attribute行
            # 对于标准模块，过滤Attribute VB_Name行
            lines = content.split('\n')
            filtered_lines = []
            skip_version_section = False
            
            for line in lines:
                line = line.rstrip()
                
                # 开始跳过VERSION部分
                if line.startswith('VERSION 1.0 CLASS'):
                    skip_version_section = True
                    continue
                
                # 结束跳过VERSION部分
                if skip_version_section and line == 'END':
                    skip_version_section = False
                    continue
                
                # 跳过Attribute行
                if line.startswith('Attribute '):
                    continue
                
                # 跳过空行（只在VERSION部分后）
                if skip_version_section:
                    continue
                
                # 添加非空行
                if line:
                    filtered_lines.append(line)
            
            # 重新组合内容
            filtered_content = '\n'.join(filtered_lines)
            
            # 以UTF-8编码写入目标文件
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(filtered_content)

            # 清理临时文件
            os.unlink(temp_path)
            log_info(f'Cleaned up temp file')

        except Exception as e:
            log_error(f'Error in _export_component: {e}')
            # 如果临时文件存在，尝试清理
            if 'temp_path' in locals() and os.path.exists(temp_path):
                try:
                    os.unlink(temp_path)
                except:
                    pass
            raise e

    def _export_document_module(self, component):
        """导出文档模块"""
        try:
            code_module = component.CodeModule
            line_count = code_module.CountOfLines
            if line_count > 0:
                # 逐行读取代码，避免空行问题
                code_lines = []
                for i in range(1, line_count + 1):
                    line = code_module.Lines(i, 1)
                    # 移除行尾的换行符，因为我们会在后面统一添加
                    line = line.rstrip('\r\n')
                    code_lines.append(line)
                # 使用单个换行符连接所有行
                code = '\n'.join(code_lines)
            else:
                code = ''

            # 添加文档模块标记
            return f'@vba-sync-document-module\n{code}'

        except Exception as e:
            log_error(f'Error exporting document module: {e}')
            raise e

    def export_module_code(self, document_path, module_name):
        """导出单个模块的代码"""
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

            if component.Type == 100:  # 文档模块
                code = self._export_document_module(component)
            else:
                code_module = component.CodeModule
                line_count = code_module.CountOfLines
                if line_count > 0:
                    code = code_module.Lines(1, line_count)
                else:
                    code = ''

            return {'success': True, 'code': code}

        except Exception as e:
            log_error(f'Error exporting module code: {e}')
            return {'success': False, 'error': str(e)}

    def list_modules(self, document_path):
        """列出所有模块"""
        try:
            if document_path not in self.connector.documents:
                return {'success': False, 'error': 'Document not open'}

            document = self.connector.documents[document_path]
            vba_project = document.VBProject
            components = vba_project.VBComponents

            modules = []
            for component in components:
                module_info = {
                    'name': component.Name,
                    'type': component.Type,
                    'hasCode': component.CodeModule.CountOfLines > 0
                }
                modules.append(module_info)

            return {'success': True, 'modules': modules}

        except Exception as e:
            log_error(f'Error listing modules: {e}')
            return {'success': False, 'error': str(e)}
