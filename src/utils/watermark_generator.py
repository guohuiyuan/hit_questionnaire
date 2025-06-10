"""
图片水印工具（带压缩功能）
功能：添加水印 + 图片压缩
"""

import os
import sys
from PIL import Image, ImageDraw, ImageFont
import argparse
import random


def compress_image(image, output_path, quality=85, optimize=True):
    """
    压缩图片并保存
    :param image: PIL Image对象
    :param output_path: 输出路径
    :param quality: 压缩质量 (1-100)
    :param optimize: 是否优化
    :return: None
    """
    try:
        if output_path.lower().endswith((".jpg", ".jpeg")):
            image.save(output_path, quality=quality, optimize=optimize)
        elif output_path.lower().endswith(".png"):
            # 使用量化和调色板优化实现PNG有损压缩
            if image.mode != "RGB":
                image = image.convert("RGB")

            # 根据quality调整量化参数
            # quality越高，保留的颜色越多，压缩率越低
            colors = max(2, min(256, int(256 * (quality / 100))))

            # 使用自适应调色板进行量化
            image = image.quantize(colors=colors, method=2, kmeans=0, palette=None)

            # 保存PNG，使用zlib压缩级别6（平衡速度和大小）
            image.save(output_path, compress_level=6)
        else:
            image.save(output_path)
    except Exception as e:
        print(f"压缩图片时出错: {output_path} - {e}")


def add_watermark(
    image_path,
    output_path,
    watermark_text,
    opacity=30,
    scale=0.8,
    angle=30,
    position="tiled",
    color="gray",
    compress=False,
    quality=85,
):
    """
    添加水印并可选压缩图片
    """
    try:
        with Image.open(image_path).convert("RGBA") as img:
            width, height = img.size

            # 创建水印层
            watermark = Image.new("RGBA", img.size, (0, 0, 0, 0))
            draw = ImageDraw.Draw(watermark)

            # 字体处理
            font = None
            for font_path in [
                "C:/Windows/Fonts/simhei.ttf",
                "C:/Windows/Fonts/simsun.ttc",
                "/System/Library/Fonts/PingFang.ttc",
                "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
            ]:
                try:
                    min_font_size = 10
                    max_font_size = 50
                    font_size = int(width * scale / len(watermark_text))
                    font_size = max(min(font_size, max_font_size), min_font_size)
                    font = ImageFont.truetype(font_path, font_size)
                    break
                except (IOError, OSError):
                    continue

            if font is None:
                try:
                    font = ImageFont.load_default()
                    print("警告: 未找到中文字体，使用默认字体。可能无法正确显示中文。")
                except Exception as e:
                    print(f"错误: 无法加载字体 - {e}")
                    return

            # 颜色处理
            if isinstance(color, str):
                color_map = {
                    "white": (255, 255, 255),
                    "black": (0, 0, 0),
                    "gray": (128, 128, 128),
                    "lightgray": (211, 211, 211),
                    "darkgray": (169, 169, 169),
                    "silver": (192, 192, 192),
                }
                rgb_color = color_map.get(color.lower(), (128, 128, 128))
            else:
                rgb_color = color

            fill_color = (*rgb_color, int(opacity * 2.55))  # 转换为0-255范围

            # 水印位置处理
            if position == "tiled":
                bbox = draw.textbbox((0, 0), watermark_text, font=font)
                text_width = bbox[2] - bbox[0]
                text_height = bbox[3] - bbox[1]

                step_x = text_width + int(text_width * 0.2)
                step_y = text_height + int(text_height * 0.2)

                for x in range(0, width, step_x):
                    for y in range(0, height, step_y):
                        draw.text((x, y), watermark_text, font=font, fill=fill_color)

            elif position == "center":
                bbox = draw.textbbox((0, 0), watermark_text, font=font)
                text_width = bbox[2] - bbox[0]
                text_height = bbox[3] - bbox[1]
                x = (width - text_width) / 2
                y = (height - text_height) / 2
                draw.text((x, y), watermark_text, font=font, fill=fill_color)

            elif position == "random":
                for _ in range(int(width * height / 10000)):
                    x = int(width * random.random())
                    y = int(height * random.random())
                    draw.text((x, y), watermark_text, font=font, fill=fill_color)

            # 旋转水印
            if angle != 0:
                watermark = watermark.rotate(angle, expand=True)
                watermark = watermark.resize(img.size, Image.BICUBIC)

            # 合并水印
            merged = Image.alpha_composite(img.convert("RGBA"), watermark)

            # 处理输出格式
            if output_path.lower().endswith((".jpg", ".jpeg")):
                merged = merged.convert("RGB")

            # 压缩处理
            if compress:
                compress_image(merged, output_path, quality=quality)
            else:
                merged.save(output_path)

            print(f"已处理: {image_path} -> {output_path}")

    except Exception as e:
        print(f"处理图片时出错: {image_path} - {e}")


def process_directory(
    input_dir, output_dir, watermark_text, compress=False, quality=85, **kwargs
):
    """
    处理目录中的所有图片
    """
    os.makedirs(output_dir, exist_ok=True)
    image_extensions = (".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".gif")

    for filename in os.listdir(input_dir):
        if filename.lower().endswith(image_extensions):
            input_path = os.path.join(input_dir, filename)
            output_path = os.path.join(output_dir, filename)
            add_watermark(
                input_path,
                output_path,
                watermark_text,
                compress=compress,
                quality=quality,
                **kwargs,
            )


def main():
    parser = argparse.ArgumentParser(description="图片水印工具（带压缩功能）")
    parser.add_argument("-i", "--input", required=True, help="输入图片路径或目录")
    parser.add_argument("-o", "--output", required=True, help="输出图片路径或目录")
    parser.add_argument("-t", "--text", required=True, help="水印文字")
    parser.add_argument(
        "-a", "--alpha", type=int, default=30, help="水印透明度 (0-100), 默认: 30"
    )
    parser.add_argument(
        "-s", "--scale", type=float, default=0.8, help="水印缩放比例, 默认: 0.8"
    )
    parser.add_argument(
        "-r", "--rotate", type=int, default=30, help="水印旋转角度, 默认: 30度"
    )
    parser.add_argument(
        "-p",
        "--position",
        choices=["tiled", "center", "random"],
        default="tiled",
        help="水印位置模式: tiled(平铺), center(居中), random(随机), 默认: tiled",
    )
    parser.add_argument(
        "-c",
        "--color",
        default="gray",
        help='水印颜色，支持: white, black, gray, lightgray, darkgray, silver 或 RGB元组(如"128,128,128")',
    )
    parser.add_argument("--compress", action="store_true", help="启用图片压缩")
    parser.add_argument(
        "--quality",
        type=int,
        default=85,
        help="压缩质量 (1-100), 默认: 85 (仅当启用压缩时有效)",
    )

    args = parser.parse_args()

    # 处理颜色参数
    if "," in args.color:
        try:
            rgb = tuple(int(x.strip()) for x in args.color.split(","))
            if len(rgb) == 3 and all(0 <= x <= 255 for x in rgb):
                args.color = rgb
            else:
                print("警告: 无效的RGB格式，使用默认灰色。")
                args.color = "gray"
        except:
            print("警告: 无法解析颜色参数，使用默认灰色。")
            args.color = "gray"

    # 处理输入输出
    if os.path.isfile(args.input):
        add_watermark(
            args.input,
            args.output,
            args.text,
            opacity=args.alpha,
            scale=args.scale,
            angle=args.rotate,
            position=args.position,
            color=args.color,
            compress=args.compress,
            quality=args.quality,
        )
    elif os.path.isdir(args.input):
        process_directory(
            args.input,
            args.output,
            args.text,
            compress=args.compress,
            quality=args.quality,
            opacity=args.alpha,
            scale=args.scale,
            angle=args.rotate,
            position=args.position,
            color=args.color,
        )
    else:
        print(f"错误: 输入路径不存在 - {args.input}")
        sys.exit(1)


if __name__ == "__main__":
    main()
