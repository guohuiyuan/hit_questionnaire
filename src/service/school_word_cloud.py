import pandas as pd
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import numpy as np
import os
import sys
from pathlib import Path

# 把项目根目录添加到系统路径
project_dir = str(Path(__file__).resolve().parents[2])
sys.path.append(project_dir)
from src.utils.watermark_generator import add_watermark

# 读取 Excel 文件
excel_file = pd.ExcelFile("data/哈工计算机25考研复试信息表_纠正后_合并.xlsx")
df = excel_file.parse("Sheet1")

# 设置中文字体
plt.rcParams["font.sans-serif"] = ["SimHei"]
plt.rcParams["axes.unicode_minus"] = False

# 提取“对26考生的备考建议”列的数据
# statistic = "对26考生的备考建议"
statistic = "本科学校（必填）"
text = " ".join([str(x) for x in df[statistic] if pd.notna(x)])

# 设置字体路径
font_path = "C:/Windows/Fonts/simhei.ttf"


# ----------------------------
# ✅ 使用数学公式生成心形 mask
# ----------------------------
def generate_heart_mask(size=500):
    # 创建坐标网格
    x = np.linspace(-1.5, 1.5, size)
    y = np.linspace(-1.5, 1.5, size)
    X, Y = np.meshgrid(x, y)

    # 心形公式
    heart = (X**2 + Y**2 - 1) ** 3 - X**2 * Y**3 <= 0

    # 将 True/False 转换为 uint8（0 和 255）
    mask = 255 * np.invert(heart).astype(np.uint8)

    # ✅ 翻转 mask，让心“正过来”
    mask = np.flipud(mask)
    return mask


# 生成大小为 1000x1000 的心形 mask
heart_mask = generate_heart_mask(size=1000)

# ----------------------------
# ✅ 创建词云
# ----------------------------
wordcloud = WordCloud(
    font_path=font_path,
    colormap="magma",
    background_color="white",
    mask=heart_mask,
    width=1600,
    height=1200,
    contour_width=1,
    contour_color="firebrick",
    max_words=500,
    min_font_size=8,
    max_font_size=100,
    random_state=42,
).generate(text)

# 显示词云
plt.figure(figsize=(10, 10))
plt.imshow(wordcloud, interpolation="bilinear")
plt.axis("off")

# 保存图片
output_folder = f"output/{statistic}心形词云"
os.makedirs(output_folder, exist_ok=True)
save_path = f"{output_folder}/{statistic}心形词云.png"
plt.savefig(save_path, bbox_inches="tight", dpi=300)

# 添加水印
try:
    add_watermark(
        save_path,
        save_path,
        watermark_text="",
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

# 展示图片
plt.show()
