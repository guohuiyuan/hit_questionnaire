import pandas as pd
from decimal import Decimal
import os

# 读取Excel数据（需确保文件路径正确）
df = pd.read_excel('data/哈工计算机25考研复试信息表_纠正后_合并.xlsx', sheet_name='Sheet1')
df = df.fillna(0)
# place = '校本部'
# place = '校本部'
# df = df[df['院系所码与名称'] == f'013-计算学部']
print(len(df))

# 定义科目与列名的映射关系（根据原表列名调整）
# subject_mapping = {
#     '政治': '政治',
#     '外语': '外语',
#     '数学': '业务课一',
#     '408': '业务课二',
#     '初试总分': '初试总分',
#     '机试(满分160)': '专业综合测试成绩',
#     '面试(满分150)': '面试成绩',
#     '总成绩': '总成绩'
# }
subject_mapping = {
    '政治': '初试政治成绩（必填）',
    '外语（包括日语和俄语）': '初试英语成绩（必填）',
    '数学': '初试数学成绩（必填）',
    '408': '初试专业课（408）成绩（必填）',
    '初试总分': '初试总分（必填）',
    '机试(满分160)': '复试机试成绩（总分160）（必填）',
    '面试(满分150)': '复试面试成绩（总分150）（必填）',
    '总成绩': '总成绩'
}

# 初始化结果字典
stats = {
    '科目': [],
    '最低分': [],
    '最高分': [],
    '平均分': [],
    '中位数': []
}

for subject, column in subject_mapping.items():
    # 提取列数据并转换为数值类型（处理可能的非数字字符）
    try:
        data = pd.to_numeric(df[column], errors='coerce').dropna()
    except Exception as e:
        print(f"处理{subject}时出错: {e}")
        continue

    # 计算统计值
    min_val = data.min()
    max_val = data.max()
    avg_val = data.mean().round(2)
    median_val = data.median()
    
    # 统一数值格式（保留两位小数，中位数取整）
    stats['科目'].append(subject)
    stats['最低分'].append(f"{min_val:.2f}")
    stats['最高分'].append(f"{max_val:.2f}")
    stats['平均分'].append(avg_val)
    stats['中位数'].append(f"{median_val:.0f}" if isinstance(median_val, (int, float)) else median_val)

# 转换为DataFrame并输出
result_df = pd.DataFrame(stats)
print("各科目成绩统计结果：")
# print(result_df.to_markdown(index=False))

# 保存为Excel文件（可选）
output_file = 'output/成绩表格'
os.makedirs(output_file, exist_ok=True)
result_df.to_excel(f'{output_file}/问卷成绩统计结果.xlsx', index=False)