import pandas as pd
import os

output_folder = "output/调剂统计"
os.makedirs(output_folder, exist_ok=True)

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
    "902-083900-00": "威海网安",
    "902-085400-24": "威海计专",
}

# 创建ExcelWriter对象
excel_path = f"{output_folder}/各专业调剂录取统计.xlsx"
with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:

    # 读取文件
    excel_file = pd.ExcelFile("data/总复试成绩单.xlsx")
    # 获取指定工作表中的数据
    df = excel_file.parse("Sheet1")

    # 构建完整的专业映射键（针对每一行数据）
    df["完整专业代码"] = (
        df["院系所码与名称"].str.split("-").str[0]
        + "-"
        + df["专业代码与名称"].str.split("-").str[0]
        + "-"
        + df["研究方向"].str.split("-").str[0]
    )

    # 筛选录取状态为已录取且一志愿未录取的数据
    filtered_df = df[(df["录取状态"] == "已录取") & (df["一志愿录取"] == "否")]

    # 记录有调剂数据的专业数量
    valid_major_count = 0

    # 循环遍历 major_mapping 中的每一个专业
    for code, major in major_mapping.items():
        # 筛选专业数据
        if code == "所有校区":
            major_df = df.copy()
        elif (
            code == "013-计算学部"
            or code == "903-哈尔滨工业大学（深圳）"
            or code == "902-哈尔滨工业大学（威海）"
        ):
            major_df = df[df["院系所码与名称"] == code].copy()
        else:
            major_df = df[df["完整专业代码"] == code].copy()
        if len(major_df) == 0:
            print(f"警告: 专业 {major} 没有数据")
            continue
        # 只处理有调剂数据的专业
        if not major_df.empty:
            # 统计录取专业的个数
            stats_df = major_df["录取专业"].value_counts().reset_index()
            stats_df.columns = ["录取专业", "个数"]
            # 按照个数列进行升序排序
            stats_df = stats_df.sort_values(by="个数")
            transposed_major_count = stats_df.set_index("录取专业").T.reset_index(
                drop=True
            )

            # 将统计结果写入Excel的不同sheet
            sheet_name = major[:31]  # 限制sheet名称长度不超过31个字符
            transposed_major_count.to_excel(writer, sheet_name=sheet_name, index=False)

            print(f"已统计专业: {major}，调剂人数: {stats_df['个数'].sum()}")
            valid_major_count += 1
        else:
            print(f"专业 {major} 没有调剂数据，跳过生成sheet")

    # 如果没有任何专业有调剂数据，删除空文件
    if valid_major_count == 0:
        os.remove(excel_path)
        print("所有专业均无调剂数据，未生成Excel文件")
    else:
        print(
            f"\n已将 {valid_major_count} 个有调剂数据的专业统计结果保存至: {excel_path}"
        )
