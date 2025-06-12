import pandas as pd
import os

# 读取文件
excel_file = pd.ExcelFile('data/哈工计算机25考研复试信息表_纠正后_合并.xlsx')

# 获取指定工作表中的数据
df = excel_file.parse('Sheet1')
df['跨考类别（必填）'] = df['跨考类别（必填）'].apply(lambda x: x.split('：')[0])

# 按本科学校类别（必填）和跨考类别（必填）列进行分组，统计每组的人数
grouped_data = df.groupby(['本科学校类别（必填）', '跨考类别（必填）'])[
    '本科学校类别（必填）'].count().reset_index(name='人数')

# 定义本科学校类别的顺序
category_order = ['C9', '985（非C9）', '211', '一本', '二本']

# # 使用Categorical类型对本科学校类别（必填）列进行排序
grouped_data['本科学校类别（必填）'] = pd.Categorical(grouped_data['本科学校类别（必填）'], categories=category_order, ordered=True)

# # 按照本科学校类别（必填）列进行排序
grouped_data = grouped_data.sort_values('本科学校类别（必填）').reset_index(drop=True)

# 将结果保存为 Excel 文件
output_folder = f"output/学校和跨考分析"
os.makedirs(output_folder, exist_ok=True)
save_path = f"{output_folder}/学校和跨考分析.xlsx"
grouped_data.to_excel(save_path, index=False)