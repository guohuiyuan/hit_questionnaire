"""
一志愿去向统计
"""

import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os
from pathlib import Path
import sys

# 把项目根目录添加到系统路径
project_dir = str(Path(__file__).resolve().parents[2])
sys.path.append(project_dir)
from src.utils.watermark_generator import add_watermark

# 设置中文字体
plt.rcParams["font.sans-serif"] = ["SimHei"]
plt.rcParams["axes.unicode_minus"] = False


def load_data(excel_path):
    """加载Excel数据"""
    data = pd.read_excel(excel_path)
    data["完整专业代码"] = (
        data["院系所码与名称"].str.split("-").str[0]
        + "-"
        + data["专业代码与名称"].str.split("-").str[0]
        + "-"
        + data["研究方向"].str.split("-").str[0]
    )
    return data


def analyze_major_data(major_df, bin_size=10):
    """分析单个专业的数据"""
    # 生成分数段
    max_score = (major_df["初试总分"].max() // bin_size + 1) * bin_size
    min_score = (major_df["初试总分"].min() // bin_size) * bin_size
    bins = list(range(min_score, max_score + bin_size, bin_size))
    labels = [f"[{bins[i]},{bins[i+1]})" for i in range(len(bins) - 1)]

    # 分析数据
    major_df["分数段"] = pd.cut(
        major_df["初试总分"], bins=bins, labels=labels, right=False
    )
    grouped = major_df.groupby(
        ["分数段", "复试及格", "录取状态", "一志愿录取"], observed=False
    ).size()

    # 填充缺失的组合
    grouped = grouped.reindex(
        pd.MultiIndex.from_product(
            [labels, ["是", "否"], ["已录取", "未录取"], ["是", "否"]],
            names=["分数段", "复试及格", "录取状态", "一志愿录取"],
        ),
        fill_value=0,
    )

    # 提取各部分数据
    failed_interview = grouped.loc[(slice(None), "否", "未录取", "否")]
    passed_not_admitted = grouped.loc[(slice(None), "是", "未录取", "否")]
    transferred = grouped.loc[(slice(None), "是", "已录取", "否")]
    first_choice = (
        grouped.loc[(slice(None), "是", "已录取", "是")].groupby("分数段").sum()
    )

    # 返回分析结果
    return pd.DataFrame(
        {
            "复试不及格": failed_interview,
            "复试及格未录取": passed_not_admitted,
            "调剂录取": transferred,
            "一志愿录取": first_choice,
        },
        index=labels,
    ).fillna(0)


def generate_chart(
    data, major_name, output_dir, watermark_text="葵妈考研", figsize=(12, 6)
):
    """生成并保存图表"""
    fig, ax = plt.subplots(figsize=figsize)
    colors = ["#1F77B4", "#FF7F0E", "#2CA02C", "#D62728"]
    width = 0.8 / len(data.columns)
    x = np.arange(len(data.index))

    for i, col in enumerate(data.columns):
        values = data[col]
        ax.bar(x + (i - 1.5) * width, values, width, label=col, color=colors[i])
        for j, val in enumerate(values):
            if val > 0:
                ax.text(
                    x[j] + (i - 1.5) * width,
                    val,
                    str(int(val)),
                    ha="center",
                    va="bottom",
                    fontsize=9,
                )

    ax.set_xticks(x)
    ax.set_xticklabels(data.index, rotation=30, ha="center")
    ax.set_xlabel("初试总分分数段")
    ax.set_ylabel("人数")
    ax.set_title(f"{major_name}一志愿去向")
    ax.legend()
    plt.tight_layout()

    output_path = os.path.join(output_dir, f"{major_name}一志愿去向.png")
    plt.savefig(output_path, bbox_inches="tight", dpi=300)
    plt.close()

    # 添加水印
    try:
        add_watermark(
            output_path,
            output_path,
            watermark_text=watermark_text,
            opacity=30,
            scale=0.8,
            angle=30,
            color="gray",
            compress=True,
            quality=10,
        )
        print(f"图表已生成: {output_path}")
    except Exception as e:
        print(f"添加水印失败: {e}")


def process_all_majors(excel_path, major_mapping, output_dir=".", bin_size=10):
    """处理所有专业数据"""
    os.makedirs(output_dir, exist_ok=True)

    try:
        # 加载数据
        data = load_data(excel_path)

        # 准备Excel写入器
        excel_writer = pd.ExcelWriter(os.path.join(output_dir, "analysis_results.xlsx"))

        for major_code, major_name in major_mapping.items():
            print(f"正在处理: {major_name}")

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

            # 分析数据
            result = analyze_major_data(major_df, bin_size)

            # 保存分析结果
            summary = pd.DataFrame({"总计": result.sum()}).T
            pd.concat([result, summary]).to_excel(
                excel_writer, sheet_name=major_name[:31]
            )

            # 生成图表（如果有数据）
            if result.sum().sum() > 0:
                generate_chart(result, major_name, output_dir)

        # 保存Excel文件
        excel_writer.close()
        print(f"分析结果已保存到: {os.path.join(output_dir, 'analysis_results.xlsx')}")

    except Exception as e:
        print(f"处理过程中出错: {e}")
        if "excel_writer" in locals():
            excel_writer.close()


def main():
    """主函数"""
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

    # 处理所有专业数据
    process_all_majors(
        excel_path="data/总复试成绩单.xlsx",
        major_mapping=major_mapping,
        output_dir="./output/一志愿去向",
    )


if __name__ == "__main__":
    main()
