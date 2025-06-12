import pandas as pd
import matplotlib.pyplot as plt
import os
import sys
from pathlib import Path

# 把项目根目录添加到系统路径
project_dir = str(Path(__file__).resolve().parents[2])
sys.path.append(project_dir)
from src.utils.watermark_generator import add_watermark

# 读取Excel文件
file_path = 'data/哈工计算机25考研复试信息表_纠正后_合并.xlsx'
df = pd.read_excel(file_path)

# 统计本科学校并按人数降序排序
school_counts = df['本科学校（必填）'].value_counts().sort_values(ascending=False)

# 反转排序结果（使人数最多的学校在顶部）
school_counts = school_counts.iloc[::-1]


# 设置中文字体
plt.rcParams["font.sans-serif"] = ["SimHei"]
plt.rcParams["axes.unicode_minus"] = False

# 创建水平条形图，增加图形尺寸以容纳更多内容
plt.figure(figsize=(12, 10), dpi=100)
bars = school_counts.plot(kind='barh', color='orange')

# 设置图表标题和坐标轴标签
plt.title('生源院校分布', fontsize=18, pad=20)  # 增加标题与图表的间距
plt.xlabel('人数', fontsize=16, labelpad=15)
plt.ylabel('本科院校', fontsize=16, labelpad=15)

# 调整y轴标签字体大小
plt.yticks(fontsize=12)

# 使用tick_params调整刻度标签与坐标轴的间距
plt.tick_params(axis='y', pad=10)  # 增加y轴刻度标签的间距

# 添加数值标签
for i, v in enumerate(school_counts):
    bars.text(v + 0.2, i, str(v), color='black', va='center', fontweight='bold')

# 手动调整边距，增加左侧空间以容纳较长的学校名称
plt.subplots_adjust(left=0.3, right=0.95, top=0.9, bottom=0.05)



# 保存图片到本地，你可以修改保存路径和文件名
output_folder = "output/本科学校"
os.makedirs(output_folder, exist_ok=True)
save_path = f"{output_folder}/本科学校分布.png"
plt.savefig(save_path, bbox_inches='tight') 

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
    print(f"图表已生成: {save_path}")
except Exception as e:
    print(f"添加水印失败: {e}")

# 显示图表
plt.show()
