#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
数据库处理模块
负责数据库连接、表查询、字段获取和数据操作
"""

import pymysql
import logging

logger = logging.getLogger(__name__)

class DBHandler:
    """数据库处理类"""
    
    def __init__(self, host, port, user, password, database):
        """初始化数据库处理器
        
        Args:
            host (str): 数据库主机名
            port (int): 数据库端口
            user (str): 用户名
            password (str): 密码
            database (str): 数据库名
        """
        self.db_config = {
            'host': host,
            'port': port,
            'user': user,
            'password': password,
            'database': database,
            'charset': 'utf8mb4',
            'cursorclass': pymysql.cursors.DictCursor
        }
        self.conn = None
    
    def connect(self):
        """连接到数据库
        
        Returns:
            bool: 连接是否成功
        """
        try:
            self.conn = pymysql.connect(**self.db_config)
            logger.info(f"成功连接到数据库 {self.db_config['database']}")
            return True
        except Exception as e:
            logger.error(f"数据库连接失败: {str(e)}")
            raise
    
    def close(self):
        """关闭数据库连接"""
        if self.conn:
            self.conn.close()
            logger.info("数据库连接已关闭")
    
    def get_tables(self):
        """获取数据库中所有表名
        
        Returns:
            list: 表名列表
        """
        if not self.conn:
            raise ConnectionError("数据库未连接")
        
        try:
            with self.conn.cursor() as cursor:
                cursor.execute("SHOW TABLES")
                tables = [list(table.values())[0] for table in cursor.fetchall()]
                logger.info(f"成功获取{len(tables)}个表")
                return tables
        except Exception as e:
            logger.error(f"获取表失败: {str(e)}")
            raise
    
    def get_table_columns(self, table_name):
        """获取指定表的字段信息
        
        Args:
            table_name (str): 表名
            
        Returns:
            list: 字段名列表
        """
        if not self.conn:
            raise ConnectionError("数据库未连接")
        
        try:
            with self.conn.cursor() as cursor:
                cursor.execute(f"DESCRIBE {table_name}")
                columns = [column['Field'] for column in cursor.fetchall()]
                logger.info(f"成功获取表{table_name}的{len(columns)}个字段")
                return columns
        except Exception as e:
            logger.error(f"获取表{table_name}字段失败: {str(e)}")
            raise
    
    def get_table_schema(self, table_name):
        """获取指定表的完整架构信息
        
        Args:
            table_name (str): 表名
            
        Returns:
            list: 字段详细信息列表
        """
        if not self.conn:
            raise ConnectionError("数据库未连接")
        
        try:
            with self.conn.cursor() as cursor:
                cursor.execute(f"DESCRIBE {table_name}")
                schema = cursor.fetchall()
                logger.info(f"成功获取表{table_name}的架构信息")
                return schema
        except Exception as e:
            logger.error(f"获取表{table_name}架构失败: {str(e)}")
            raise
    
    def execute_sql(self, sql, params=None):
        """执行SQL语句
        
        Args:
            sql (str): SQL语句
            params (tuple, optional): 参数
            
        Returns:
            int: 影响的行数
        """
        if not self.conn:
            raise ConnectionError("数据库未连接")
        
        try:
            with self.conn.cursor() as cursor:
                result = cursor.execute(sql, params)
                self.conn.commit()
                logger.info(f"成功执行SQL，影响{result}行")
                return result
        except Exception as e:
            self.conn.rollback()
            logger.error(f"执行SQL失败: {str(e)}")
            logger.error(f"SQL语句: {sql}")
            if params:
                logger.error(f"参数: {params}")
            raise
    
    def execute_many(self, sql, data):
        """批量执行SQL语句
        
        Args:
            sql (str): SQL语句
            data (list): 数据列表
            
        Returns:
            int: 影响的行数
        """
        if not self.conn:
            raise ConnectionError("数据库未连接")
        
        try:
            with self.conn.cursor() as cursor:
                result = cursor.executemany(sql, data)
                # 获取受影响的行数
                affected_rows = cursor.rowcount
                
                # 额外安全检查：如果result是元组，取第一个元素
                if isinstance(result, tuple):
                    affected_rows = result[0]
                    logger.info(f"执行结果是元组，使用第一个元素作为受影响行数: {affected_rows}")
                
                self.conn.commit()
                logger.info(f"成功批量执行SQL，影响{affected_rows}行")
                return affected_rows
        except Exception as e:
            self.conn.rollback()
            logger.error(f"批量执行SQL失败: {str(e)}")
            logger.error(f"SQL语句: {sql}")
            if data:
                logger.error(f"数据示例: {data[:1]}")
            raise
    
    def test_connection(self):
        """测试数据库连接
        
        Returns:
            bool: 连接是否成功
        """
        try:
            temp_conn = pymysql.connect(**self.db_config)
            temp_conn.close()
            return True
        except Exception as e:
            logger.error(f"数据库连接测试失败: {str(e)}")
            return False
