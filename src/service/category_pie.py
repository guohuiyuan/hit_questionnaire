import pandas as pd
import matplotlib.pyplot as plt
import os
from pathlib import Path
import sys

# 把项目根目录添加到系统路径
project_dir = str(Path(__file__).resolve().parents[2])
sys.path.append(project_dir)
from src.utils.watermark_generator import add_watermark

# 读取Excel文件
file_path = 'data/哈工计算机25考研复试信息表_纠正后_合并.xlsx'
df = pd.read_excel(file_path)

# statistic = "本科学校类别（必填）"
statistic = "跨考类别（必填）"
# 统计“本科学校类别（必填）”的分布
df[statistic] = df[statistic].apply(lambda x: x.split('：')[0])
category_counts = df[statistic].value_counts()
total = category_counts.sum()

# 计算百分比
category_percentages = (category_counts / total * 100).round(1)

# 创建结果DataFrame
result_df = pd.DataFrame({
    statistic: category_counts.index,
    '人数': category_counts.values,
    '百分比(%)': category_percentages.values
})

# 按人数降序排序
result_df = result_df.sort_values('人数', ascending=False)

# 设置中文字体和解决负号显示问题
plt.rcParams["font.sans-serif"] = ["SimHei"]
plt.rcParams["axes.unicode_minus"] = False

# 自定义格式化函数，用于显示人数和百分比
def make_autopct(values):
    def my_autopct(pct):
        total = sum(values)
        val = int(round(pct * total / 100.0))
        return f'{val} ({pct:.1f}%)'
    return my_autopct

# 创建图表
fig, ax = plt.subplots(figsize=(10, 8))

# 突出最大一块
explode = [0.05 if val == category_counts.max() else 0 for val in category_counts]

# 绘制饼图
wedges, texts, autotexts = ax.pie(
    category_counts,
    labels=None,  # 不直接显示标签，改用图例
    autopct=make_autopct(category_counts),  # 使用自定义的格式化函数
    pctdistance=0.75,  # 百分比标签离中心的距离
    textprops=dict(color="black", fontsize=10, weight="bold"),  # 标签样式
    startangle=90,
    explode=explode,
    colors=plt.cm.Paired.colors,
    wedgeprops=dict(width=0.4)  # 控制饼图厚度（甜甜圈样式）
)

# 添加图例并美化
legend = ax.legend(wedges, category_counts.index,
                   title=f"{statistic}分布",
                   loc="center left",
                   bbox_to_anchor=(1, 0, 0.5, 1),
                   prop={'size': 12})

# 设置标题
ax.set_title(f"{statistic}分布", fontsize=16)

# 保证饼图为圆形
ax.axis('equal')

# 自动调整布局，防止被截断
plt.tight_layout()

# 保存图片到本地，使用 bbox_inches='tight' 避免白边
output_folder = f"output/{statistic}"
os.makedirs(output_folder, exist_ok=True)
save_path = f"{output_folder}/{statistic}分布.png"

plt.savefig(save_path, dpi=300, bbox_inches='tight')

# 添加水印
try:
    add_watermark(
        save_path,
        save_path,
        watermark_text="葵妈考研",
        opacity=30,
        scale=0.8,
        angle=30,
        color="gray",
        compress=True,
        quality=10,
    )
    print(f"饼图已生成并添加水印: {save_path}")
except Exception as e:
    print(f"添加水印失败: {e}")

# 显示图表
plt.show()

# 保存统计数据到Excel
excel_path = f"{output_folder}/{statistic}分布统计.xlsx"
result_df.to_excel(excel_path, index=False)
print(f"统计数据已保存到Excel: {excel_path}")
