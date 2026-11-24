#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
创建xls格式的示例文件用于测试
"""

import pandas as pd

def create_xls_sample():
    """创建包含空格问题的示例xls文件"""
    
    # 创建示例数据，包含各种空格问题
    data = {
        '员工姓名': [
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
        ],
        '所属部门': [
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
        ],
        '员工编号': [
            'EMP001',
            'EMP002',
            'EMP003',
            'EMP004',
            'EMP005',
            'EMP006',
            'EMP007',
            'EMP008',
            'EMP009',
            'EMP010',
            'EMP011',
        ]
    }
    
    # 创建DataFrame
    df = pd.DataFrame(data)
    
    # 保存为xls文件
    output_file = 'sample_data_test.xls'
    try:
        # 由于新版pandas可能不支持xlwt，我们先保存为xlsx然后重命名
        temp_file = 'temp_sample.xlsx'
        df.to_excel(temp_file, index=False, engine='openpyxl')
        
        # 使用xlwt直接创建xls文件
        import xlwt
        workbook = xlwt.Workbook()
        worksheet = workbook.add_sheet('Sheet1')
        
        # 写入列标题
        for col_idx, col_name in enumerate(df.columns):
            worksheet.write(0, col_idx, col_name)
        
        # 写入数据
        for row_idx, row in df.iterrows():
            for col_idx, value in enumerate(row):
                worksheet.write(row_idx + 1, col_idx, value)
        
        workbook.save(output_file)
        
        # 删除临时文件
        import os
        if os.path.exists(temp_file):
            os.remove(temp_file)
            
        print(f"XLS示例文件已创建：{output_file}")
        
    except Exception as e:
        print(f"创建XLS文件失败：{e}")
        # 如果失败，创建xlsx文件作为替代
        output_file = 'sample_data_test.xlsx'
        df.to_excel(output_file, index=False, engine='openpyxl')
        print(f"已创建XLSX文件作为替代：{output_file}")
    
    print("\n文件内容预览：")
    print(df.to_string(index=False))
    
    print("\n数据说明：")
    print("- 员工姓名列包含多个重复但有空格差异的数据")
    print("- 所属部门列也包含类似问题")
    print("- 这是xls格式文件，可以用来测试程序对老版本Excel文件的支持")
    
    return output_file

if __name__ == "__main__":
    create_xls_sample()