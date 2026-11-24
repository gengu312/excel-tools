#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
创建示例Excel文件用于演示重复数据检测工具
"""

import pandas as pd

def create_sample_excel():
    """创建包含空格问题的示例Excel文件"""
    
    # 创建示例数据，包含各种空格问题
    data = {
        '姓名': [
            '张三',      # 正常数据
            '张三 ',     # 尾部空格
            ' 张三',     # 头部空格
            '  张三  ',  # 头尾空格
            '李四',      # 正常数据
            '李四  ',    # 尾部多个空格
            ' 李四 ',    # 头尾空格
            '王五',      # 正常数据
            '王五',      # 完全相同
            '赵六 ',     # 尾部空格
            '赵六',      # 正常数据
            '钱七',      # 正常数据
            ' 钱七',     # 头部空格
            '孙八  ',    # 尾部空格
            '孙八',      # 正常数据
        ],
        '部门': [
            '销售部',
            '销售部 ',
            ' 销售部',
            '技术部',
            '技术部  ',
            '人事部',
            ' 人事部 ',
            '财务部',
            '财务部',
            '市场部 ',
            '市场部',
            '运营部',
            ' 运营部',
            '客服部  ',
            '客服部',
        ],
        '工号': [
            'E001',
            'E002',
            'E003',
            'E004',
            'E005',
            'E006',
            'E007',
            'E008',
            'E009',
            'E010',
            'E011',
            'E012',
            'E013',
            'E014',
            'E015',
        ]
    }
    
    # 创建DataFrame
    df = pd.DataFrame(data)
    
    # 保存为Excel文件
    output_file = 'sample_data.xlsx'
    df.to_excel(output_file, index=False)
    
    print(f"示例Excel文件已创建：{output_file}")
    print("\n文件内容预览：")
    print(df.to_string(index=False))
    
    print("\n数据说明：")
    print("- 姓名列包含多个重复但有空格差异的数据")
    print("- 部门列也包含类似问题")
    print("- 可以用这个文件测试重复数据检测工具")
    
    return output_file

if __name__ == "__main__":
    create_sample_excel()