import pandas as pd
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import os
import sys
from pathlib import Path

# 把项目根目录添加到系统路径
project_dir = str(Path(__file__).resolve().parents[2])
sys.path.append(project_dir)
from src.utils.watermark_generator import add_watermark

# 读取文件
excel_file = pd.ExcelFile("data/哈工计算机25考研复试信息表_纠正后_合并.xlsx")

# 获取指定工作表中的数据
df = excel_file.parse("Sheet1")

# 设置中文字体
plt.rcParams["font.sans-serif"] = ["SimHei"]
plt.rcParams["axes.unicode_minus"] = False

# 提取本科学校（必填）列的数据
# statistic = "本科学校（必填）"
statistic = "对26考生的备考建议"
text = " ".join([str(x) for x in df[statistic] if pd.notna(x)])

# 设置词云字体，Windows 下 WenQuanYi Zen Hei 字体路径示例
font_path = "C:/Windows/Fonts/simhei.ttf"
# 生成词云，设置字体
wordcloud = WordCloud(
    font_path=font_path,
    colormap="magma",
    background_color="white",
    width=1600,
    height=1200,
).generate(text)

# 显示词云
plt.imshow(wordcloud, interpolation="bilinear")
plt.axis("off")

# 保存图片到本地，你可以修改保存路径和文件名
output_folder = f"output/{statistic}词云"
os.makedirs(output_folder, exist_ok=True)
save_path = f"{output_folder}/{statistic}词云.png"
plt.savefig(save_path)

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
