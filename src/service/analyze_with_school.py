import pandas as pd
import os

# 读取Excel文件
file_path = "data/哈工计算机25考研复试信息表_纠正后_合并.xlsx"
df = pd.read_excel(file_path)

# statistic = "本科学校类别（必填）"
# statistic = "跨考类别（必填）"
# statistic = "是否二战及以上（必填）"
# statistic = "项目经历（必填）"
# statistic = "论文发表情况（必填）"
# statistic = "数学建模获奖情况（必填）"
# statistic = "ICPC竞赛经历（必填）"
# statistic = "开始复习月份（必填）"
# statistic = "备考状态（必填）"
statistic = "OI竞赛经历（必填）"
# 定义分组映射字典
month_groups = {
    "没参加过OI": ["没参加过OI"],
    "参加过OI": [
        "NOIP提高组二等奖以下",
        "NOIP提高组二等奖",
    ],
}

# 初始化分组列，默认值为原始月份
df[statistic + "_分组"] = df[statistic]

# 对每个分组进行映射
for group_name, months in month_groups.items():
    df.loc[df[statistic].isin(months), statistic + "_分组"] = group_name
# 按本科学校类别（必填）列进行分组，计算各分组下初试总分（必填）、复试机试成绩（总分160）（必填）、复试面试成绩（总分150）（必填）和总成绩的最低分、最高分和平均分
grouped = (
    df.groupby(statistic + "_分组")[
        [
            "初试总分（必填）",
            "复试机试成绩（总分160）（必填）",
            "复试面试成绩（总分150）（必填）",
        ]
    ]
    .agg(["min", "max", "mean"])
    .reset_index()
)

# 重命名列名
grouped.columns = [
    statistic,
    "初试总分最低分",
    "初试总分最高分",
    "初试总分平均分",
    "复试机试成绩最低分",
    "复试机试成绩最高分",
    "复试机试成绩平均分",
    "复试面试成绩最低分",
    "复试面试成绩最高分",
    "复试面试成绩平均分",
]

# 将平均分结果保留两位小数
for col in grouped.columns:
    if "平均分" in col:
        grouped[col] = grouped[col].round(2)

# 定义本科学校类别的顺序
# category_order = ['C9', '985（非C9）', '211', '一本', '二本']

# category_order = [
#     "有高于课程设计水平的项目 (包括公司的实习项目)",
#     "有不高于课程设计水平的项目",
#     "没有任何计算机相关项目",
# ]
# # 使用Categorical类型对本科学校类别（必填）列进行排序
# grouped[statistic] = pd.Categorical(
#     grouped[statistic], categories=category_order, ordered=True
# )

# # 按照本科学校类别（必填）列进行排序
# grouped = grouped.sort_values(statistic).reset_index(drop=True)

# 将结果保存为 Excel 文件
output_folder = f"output/{statistic}分析"
os.makedirs(output_folder, exist_ok=True)
save_path = f"{output_folder}/{statistic}分析.xlsx"
grouped.to_excel(save_path, index=False)
