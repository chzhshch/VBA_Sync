#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CustomDocumentProperties 管理器
处理文档的自定义属性
"""

from .utils import log_info, log_error

class PropertiesManager:
    def __init__(self, connector):
        self.connector = connector

    def get_last_sync_time(self, document_path):
        """获取最后同步时间"""
        try:
            if document_path not in self.connector.documents:
                return None

            document = self.connector.documents[document_path]
            properties = self._get_custom_properties(document)

            if properties is None:
                return None

            try:
                # 尝试获取属性
                prop = properties('VBA_LastSyncTime')
                return prop.Value
            except Exception:
                # 属性不存在
                return None

        except Exception as e:
            log_error(f'Error getting last sync time: {e}')
            return None

    def set_last_sync_time(self, document_path, sync_time=None):
        """设置最后同步时间"""
        try:
            if document_path not in self.connector.documents:
                return {'success': False, 'error': 'Document not open'}

            document = self.connector.documents[document_path]
            properties = self._get_custom_properties(document)

            if properties is None:
                return {'success': False, 'error': 'Cannot access custom properties'}

            # 检查属性是否存在
            prop_exists = False
            try:
                prop = properties('VBA_LastSyncTime')
                prop_exists = True
            except Exception:
                pass

            if prop_exists:
                # 更新现有属性
                prop = properties('VBA_LastSyncTime')
                prop.Value = sync_time
            else:
                # 创建新属性
                properties.Add('VBA_LastSyncTime', False, 4, sync_time)

            # 保存文档以持久化属性
            document.Save()

            log_info(f'Set last sync time: {sync_time}')
            return {'success': True, 'message': 'Last sync time updated'}

        except Exception as e:
            log_error(f'Error setting last sync time: {e}')
            return {'success': False, 'error': str(e)}

    def _get_custom_properties(self, document):
        """获取文档的自定义属性"""
        try:
            # 不同 Office 应用的属性访问方式不同
            if hasattr(document, 'CustomDocumentProperties'):
                return document.CustomDocumentProperties
            elif hasattr(document, 'BuiltInDocumentProperties'):
                # 对于某些应用，可能需要通过不同方式访问
                return document.BuiltInDocumentProperties
            return None
        except Exception as e:
            log_error(f'Error getting custom properties: {e}')
            return None
