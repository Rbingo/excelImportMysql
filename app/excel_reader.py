#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel读取模块
负责Excel文件的读取、解析和数据处理
"""

import pandas as pd
import logging

logger = logging.getLogger(__name__)

class ExcelReader:
    """Excel读取类"""
    
    def __init__(self, file_path):
        """初始化Excel读取器
        
        Args:
            file_path (str): Excel文件路径
        """
        self.file_path = file_path
        self.df = None
        self.sheet_names = []
        self.current_sheet = None
    
    def load_file(self):
        """加载Excel文件，获取所有sheet名称
        
        Returns:
            list: sheet名称列表
        """
        try:
            # 获取所有sheet名称
            excel_file = pd.ExcelFile(self.file_path)
            self.sheet_names = excel_file.sheet_names
            logger.info(f"成功加载Excel文件，包含{len(self.sheet_names)}个sheet: {self.sheet_names}")
            return self.sheet_names
        except Exception as e:
            logger.error(f"加载Excel文件失败: {str(e)}")
            raise
    
    def read_sheet(self, sheet_name):
        """读取指定sheet的数据
        
        Args:
            sheet_name (str): sheet名称
        
        Returns:
            pandas.DataFrame: 读取的数据
        """
        try:
            # 读取Excel文件，保留空值为NaN
            self.df = pd.read_excel(
                self.file_path, 
                sheet_name=sheet_name,
                keep_default_na=True,  # 保持默认的空值处理
                na_values=['']  # 将空字符串也视为NaN
            )
            
            # 进一步处理：确保所有空值都被识别为NaN
            self.df = self.df.replace('', pd.NA)
            
            self.current_sheet = sheet_name
            logger.info(f"成功读取sheet '{sheet_name}'，共{len(self.df)}行，{len(self.df.columns)}列")
            logger.info(f"Excel列名: {list(self.df.columns)}")
            return self.df
        except Exception as e:
            logger.error(f"读取sheet '{sheet_name}'失败: {str(e)}")
            raise
    
    def get_columns(self):
        """获取当前sheet的列名
        
        Returns:
            list: 列名列表
        """
        if self.df is None:
            raise ValueError("未读取任何sheet数据")
        
        return list(self.df.columns)
    
    def get_data(self, start_row=0, end_row=None):
        """获取数据，可指定行范围
        
        Args:
            start_row (int): 起始行索引
            end_row (int, optional): 结束行索引
        
        Returns:
            pandas.DataFrame: 数据
        """
        if self.df is None:
            raise ValueError("未读取任何sheet数据")
        
        return self.df.iloc[start_row:end_row]
    
    def get_row_count(self):
        """获取当前sheet的行数
        
        Returns:
            int: 行数
        """
        if self.df is None:
            return 0
        
        return len(self.df)
    
    def get_column_count(self):
        """获取当前sheet的列数
        
        Returns:
            int: 列数
        """
        if self.df is None:
            return 0
        
        return len(self.df.columns)
    
    def preview_data(self, rows=10):
        """预览数据
        
        Args:
            rows (int): 预览行数
        
        Returns:
            pandas.DataFrame: 预览数据
        """
        if self.df is None:
            raise ValueError("未读取任何sheet数据")
        
        return self.df.head(rows)
    
    def validate_data(self):
        """验证数据完整性
        
        Returns:
            dict: 验证结果
        """
        if self.df is None:
            raise ValueError("未读取任何sheet数据")
        
        validation_result = {
            'total_rows': len(self.df),
            'total_columns': len(self.df.columns),
            'null_values': self.df.isnull().sum().to_dict(),
            'duplicate_rows': self.df.duplicated().sum()
        }
        
        logger.info(f"数据验证结果: {validation_result}")
        return validation_result
    
    def convert_to_dict(self):
        """将数据转换为字典列表
        
        Returns:
            list: 字典列表
        """
        if self.df is None:
            raise ValueError("未读取任何sheet数据")
        
        return self.df.to_dict('records')
    
    def get_data_types(self):
        """获取各列的数据类型
        
        Returns:
            dict: 列名到数据类型的映射
        """
        if self.df is None:
            raise ValueError("未读取任何sheet数据")
        
        return self.df.dtypes.to_dict()
