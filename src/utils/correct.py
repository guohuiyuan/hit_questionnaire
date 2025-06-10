import pandas as pd

# 读取文件
excel_file1 = pd.ExcelFile('data/总复试成绩单.xlsx')
excel_file2 = pd.ExcelFile('data/哈工计算机25考研复试信息表_纠正后.xlsx')

# 获取指定工作表中的数据
df1 = excel_file1.parse('Sheet1')
df2 = excel_file2.parse('Sheet1')

# 定义关联字段
key_columns = ['政治', '外语', '业务课一', '业务课二', '初试总分']

# 确定 df1 中除了关联字段之外需要保留的列
cols_to_keep = [col for col in df1.columns if col not in key_columns]

# 根据关联字段合并两个 DataFrame
merged_df = pd.merge(df1, df2, left_on=key_columns, right_on=['初试政治成绩（必填）', '初试英语成绩（必填）', '初试数学成绩（必填）', '初试专业课（408）成绩（必填）', '初试总分（必填）'], how='inner')

# 用 df1 的机试原始分和面试成绩更新 df2 的对应列
for index, row in merged_df.iterrows():
    df2_index = df2[(df2['初试政治成绩（必填）'] == row['政治']) &
                    (df2['初试英语成绩（必填）'] == row['外语']) &
                    (df2['初试数学成绩（必填）'] == row['业务课一']) &
                    (df2['初试专业课（408）成绩（必填）'] == row['业务课二']) &
                    (df2['初试总分（必填）'] == row['初试总分'])].index[0]

    df2.at[df2_index, '复试机试成绩（总分160）（必填）'] = row['机试原始分']
    df2.at[df2_index, '复试面试成绩（总分150）（必填）'] = row['面试成绩']

    # 将 df1 中其他需要保留的列数据添加到 df2 中
    for col in cols_to_keep:
        df2.at[df2_index, col] = row[col]

# 将结果保存为 Excel 文件
df2.to_excel('data/哈工计算机25考研复试信息表_纠正后_合并.xlsx', index=False)