#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
导入管理模块
负责协调Excel读取和数据库导入的核心逻辑
"""

import logging
import pandas as pd

logger = logging.getLogger(__name__)

class ImportManager:
    """导入管理器，负责协调Excel读取和数据库导入"""
    
    def __init__(self, excel_reader, db_handler):
        """初始化导入管理器
        
        Args:
            excel_reader (ExcelReader): Excel读取器实例
            db_handler (DBHandler): 数据库处理器实例
        """
        self.excel_reader = excel_reader
        self.db_handler = db_handler
        self.mapping = []  # 改为列表形式，支持相同表头
        self.primary_key = None
        self.unique_keys = []  # 改为列表，支持多个唯一索引字段
    
    def set_mapping(self, mapping):
        """设置字段映射关系
        
        Args:
            mapping (list): 字段映射关系列表，每个元素为 (excel_col, db_col) 元组
        """
        self.mapping = mapping
        logger.info(f"设置字段映射: {mapping}")
    
    def set_primary_key(self, primary_key):
        """设置主键字段
        
        Args:
            primary_key (str): 主键字段名
        """
        self.primary_key = primary_key
        logger.info(f"设置主键: {primary_key}")
    
    def set_unique_key(self, unique_key):
        """设置唯一索引字段
        
        Args:
            unique_key (str): 唯一索引字段名
        """
        # 支持单个唯一键，兼容性处理
        if unique_key:
            self.unique_keys = [unique_key]
        else:
            self.unique_keys = []
        logger.info(f"设置唯一索引: {unique_key}")
    
    def set_unique_keys(self, unique_keys):
        """设置多个唯一索引字段
        
        Args:
            unique_keys (list): 唯一索引字段名列表
        """
        self.unique_keys = unique_keys if unique_keys else []
        logger.info(f"设置唯一索引: {unique_keys}")
    
    def generate_insert_sql(self, table_name):
        """生成插入SQL语句
        
        Args:
            table_name (str): 表名
            
        Returns:
            str: 插入SQL语句
        """
        if not self.mapping:
            raise ValueError("未设置字段映射")
        
        # 数据库字段列表
        db_columns = [db_col for _, db_col in self.mapping]
        columns_str = ", ".join(db_columns)
        
        # 占位符列表
        placeholders = ", ".join(["%s"] * len(db_columns))
        
        # 基础插入语句
        sql = f"INSERT INTO {table_name} ({columns_str}) VALUES ({placeholders})"
        
        # 生成更新部分（总是添加ON DUPLICATE KEY UPDATE，当唯一索引存在时自动更新）
        update_clause = []
        for col in db_columns:
            # 检查是否为主键或唯一索引字段
            is_key_field = col == self.primary_key or col in self.unique_keys
            if not is_key_field:
                update_clause.append(f"{col} = VALUES({col})")
        
        if update_clause:
            sql += f" ON DUPLICATE KEY UPDATE {', '.join(update_clause)}"
        
        logger.info(f"生成SQL语句: {sql}")
        return sql
    
    def prepare_data(self):
        """准备导入数据
        
        Returns:
            list: 数据列表，每个元素为一个元组
        """
        if not self.mapping:
            raise ValueError("未设置字段映射")
        
        # 获取数据
        df = self.excel_reader.df
        
        # 处理每个字段映射，直接生成数据列表
        data = []
        
        # 遍历每一行数据
        for _, row in df.iterrows():
            # 检查是否需要跳过当前行：如果设置了唯一索引且对应Excel列为空
            skip_row = False
            if self.unique_keys:
                # 检查所有唯一索引字段
                for unique_key in self.unique_keys:
                    # 找到唯一索引对应的Excel列
                    for excel_col, db_col in self.mapping:
                        if db_col == unique_key:
                            # 检查Excel列是否存在
                            if excel_col in df.columns:
                                unique_value = row[excel_col]
                                # 如果唯一索引值为空，则跳过当前行
                                if pd.isna(unique_value) or unique_value == '':
                                    skip_row = True
                                    break
                    if skip_row:
                        break
                
            if skip_row:
                continue
            
            # 为当前行创建一个数据元组
            row_data = []
            
            # 处理每个字段映射
            for excel_col, db_col in self.mapping:
                # 检查excel_col是否存在于原始数据中
                if excel_col in df.columns:
                    value = row[excel_col]
                else:
                    # 处理列名不存在的情况（例如：不同大小写）
                    # 尝试不区分大小写匹配
                    matched_col = None
                    for col in df.columns:
                        if col.lower() == excel_col.lower():
                            matched_col = col
                            break
                    
                    if matched_col:
                        value = row[matched_col]
                    else:
                        # 如果找不到匹配的列，使用空值
                        value = None
                
                # 详细处理空值：确保所有类型的空值都被转换为None
                if pd.isna(value) or value == '' or value == 'NULL':
                    value = None
                logger.debug(f"处理字段 {db_col}: 原值='{value}'，处理后='{value}'")
                row_data.append(value)
            
            # 将当前行数据添加到数据列表
            data.append(tuple(row_data))
        
        logger.info(f"准备了{len(data)}行导入数据")
        logger.info(f"数据列数: {len(self.mapping)}, 每行数据长度: {len(data[0]) if data else 0}")
        return data
    
    def import_data(self, table_name):
        """执行数据导入
        
        Args:
            table_name (str): 表名
            
        Returns:
            dict: 导入结果
        """
        try:
            # 生成SQL
            sql = self.generate_insert_sql(table_name)
            
            # 准备数据
            self.table_name = table_name  # 保存表名用于准备数据时的NOT NULL检查
            data = self.prepare_data()
            
            # 执行批量插入
            affected_rows = self.db_handler.execute_many(sql, data)
            
            result = {
                'success': True,
                'affected_rows': affected_rows,
                'message': f"成功导入{affected_rows}行数据"
            }
            
            logger.info(result['message'])
            return result
        except Exception as e:
            # 添加更详细的错误信息
            logger.error(f"导入失败，错误类型: {type(e)}, 错误内容: {e}")
            import traceback
            logger.error(f"错误堆栈: {traceback.format_exc()}")
            
            result = {
                'success': False,
                'affected_rows': 0,
                'message': f"导入失败: {str(e)}"
            }
            
            return result
    
    def validate_import_settings(self):
        """验证导入设置
        
        Returns:
            dict: 验证结果
        """
        validation_result = {
            'valid': True,
            'errors': []
        }
        
        # 检查字段映射
        if not self.mapping:
            validation_result['valid'] = False
            validation_result['errors'].append("未设置字段映射")
        else:
            # 提取所有数据库字段
            db_columns = [db_col for _, db_col in self.mapping]
            
            # 检查主键或唯一索引是否在映射中
            if self.primary_key and self.primary_key not in db_columns:
                validation_result['valid'] = False
                validation_result['errors'].append(f"主键'{self.primary_key}'不在映射字段中")
            
            # 检查所有唯一索引字段
            for unique_key in self.unique_keys:
                if unique_key not in db_columns:
                    validation_result['valid'] = False
                    validation_result['errors'].append(f"唯一索引'{unique_key}'不在映射字段中")
        
        logger.info(f"导入设置验证结果: {validation_result}")
        return validation_result
