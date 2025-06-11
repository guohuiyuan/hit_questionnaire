import pandas as pd
import matplotlib.pyplot as plt
import sys
from pathlib import Path
import os

# 把项目根目录添加到系统路径
project_dir = str(Path(__file__).resolve().parents[2])
sys.path.append(project_dir)
from src.utils.watermark_generator import add_watermark

# 专业映射
major_mapping = {
    "所有校区": "所有校区",
    "013-计算学部": "校本部",
    "903-哈尔滨工业大学（深圳）": "深圳校区",
    "902-哈尔滨工业大学（威海）": "威海校区",
    "013-081200-00": "本部计学",
    "013-083500-00": "本部软学",
    "013-083900-00": "本部网学",
    "013-085400-11": "本部计专",
    "013-085400-12": "本部软专",
    "013-085400-13": "本部网专",
    "013-085400-40": "苏州计专",
    "013-085400-41": "郑州计专",
    "013-085400-42": "郑州软专",
    "013-085400-43": "郑州网专",
    "013-085400-44": "重庆计专",
    "013-085400-45": "重庆软专",
    "013-085400-46": "重庆网专",
    "013-085400-60": "工程联培",
    "013-087600-00": "本部智科",
    "903-081200-00": "深圳计学",
    "903-085400-31": "深圳计专",
    "902-081200-00": "威海计学",
    "902-083900-00": "威海网学",
    "902-085400-24": "威海计专",
}

statistic = "问卷66人机试"
output_folder = f"output/{statistic}分布"
os.makedirs(output_folder, exist_ok=True)
# 读取文件
input_file = "哈工计算机25考研复试信息表_纠正后_合并.xlsx"
input_folder = "data"

# 创建ExcelWriter对象，用于保存所有校区的统计数据
excel_writer = pd.ExcelWriter(
    os.path.join(output_folder, f"各校区{statistic}分布统计.xlsx"), engine="openpyxl"
)

excel_file = pd.ExcelFile(os.path.join(input_folder, input_file))
# 获取指定工作表中的数据
df = excel_file.parse("Sheet1")
data = df
data["完整专业代码"] = (
    data["院系所码与名称"].str.split("-").str[0]
    + "-"
    + data["专业代码与名称"].str.split("-").str[0]
    + "-"
    + data["研究方向"].str.split("-").str[0]
)
data[statistic] = data["复试机试成绩（总分160）（必填）"]
for major_code, major_name in major_mapping.items():
    print(f"正在处理: {major_name}")

    # 只生成所有校区的统计数据
    if major_code != "所有校区":
        continue

    # 筛选专业数据
    if major_code == "所有校区":
        major_df = data.copy()
    elif (
        major_code == "013-计算学部"
        or major_code == "903-哈尔滨工业大学（深圳）"
        or major_code == "902-哈尔滨工业大学（威海）"
    ):
        major_df = data[data["院系所码与名称"] == major_code].copy()
    else:
        major_df = data[data["完整专业代码"] == major_code].copy()
    if len(major_df) == 0:
        print(f"警告: 专业 {major_name} 没有数据")
        continue

    print(f"处理 {major_name}，数据量: {len(major_df)}")

    # 定义分数段
    max_score = major_df[statistic].max() // 10 * 10 + 10
    min_score = major_df[statistic].min() // 10 * 10
    print(f"分数段范围: {min_score}~{max_score}")
    bins = list(range(int(min_score), int(max_score) + 10, 10))
    labels = [f"[{bins[i]},{bins[i + 1]})" for i in range(len(bins) - 1)]

    # 对初试总分进行分组统计
    major_df["分数段"] = pd.cut(
        major_df[statistic], bins=bins, labels=labels, right=False
    )
    score_distribution = (
        major_df["分数段"].value_counts().reindex(labels).reset_index(name="人数")
    )

    # 将统计结果保存到Excel的不同sheet中
    sheet_name = (
        major_name if len(major_name) <= 31 else major_name[:31]
    )  # Excel sheet名不能超过31个字符
    score_distribution.to_excel(excel_writer, sheet_name=sheet_name, index=False)

    # 设置中文字体
    plt.rcParams["font.sans-serif"] = ["SimHei"]

    # 绘制矩形图
    plt.figure(figsize=(10, 6))  # 设置图表大小
    width = 0.8  # 矩形宽度
    plt.xticks(rotation=30)
    bars = plt.bar(score_distribution["分数段"], score_distribution["人数"], width)
    plt.xlabel(f"{statistic}分数段")
    plt.ylabel("人数")
    plt.title(f"{major_name}{statistic}分布矩形图")

    # 在柱子上方添加数值
    for bar in bars:
        height = bar.get_height()
        plt.text(
            bar.get_x() + bar.get_width() / 2.0,
            height,
            f"{height}",
            ha="center",
            va="bottom",
            fontsize=10,
        )

    # 保存图片
    file_path = f"{major_name}{statistic}分布矩形图.png"
    file_path = os.path.join(output_folder, file_path)
    plt.savefig(file_path, dpi=300, bbox_inches="tight")  # 提高图片清晰度
    plt.close()

    # 添加水印
    try:
        add_watermark(
            file_path,
            file_path,
            watermark_text="葵妈考研",
            compress=True,
            quality=10,
        )
    except Exception as e:
        print(f"添加水印失败: {e}")

    print(f"已生成: {file_path}")
    print(f"{statistic}分布表：")
    print(score_distribution)
    print("-" * 50)  # 分隔线

# 保存Excel文件
excel_writer.close()
print(
    f"\n所有校区统计数据已保存至: {os.path.join(output_folder, f'各校区{statistic}分布统计.xlsx')}"
)
