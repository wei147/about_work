import os
import pandas as pd

# 必需的字段，包括但不限于 "自定义sku" 和 "渠道sku"
EXPECTED_COLUMNS = ["平台", "渠道", "渠道sku", "平台ID", "自定义sku"]  # 不包括"仓区/市场"和"来源"，因为这是要动态添加的

# 渠道映射关系，定义你的规则
channel_mapping = {
    "us": "美国仓",
    "pl": "欧洲仓",
    "de": "欧洲仓",
    "uk": "英国仓",
    "gb": "英国仓",
    "ca": "加拿大仓",
    "fba": "FBA仓",
    "jp": "日本仓",
}

# 定义需要保留的列
REQUIRED_COLUMNS = ["平台", "渠道", "渠道sku", "平台ID", "自定义sku", "来源", "仓区/市场"]

# 定义特殊来源的特殊映射规则
special_source_mapping = {
    "HZ-sunflow": {  # 如果来源是 HZ-sunflow，定义特殊的渠道映射规则
    },
}

# 定义需要过滤的平台
# PLATFORM_FILTER = ["HomeDepot", "Wayfair", "Amazon", "Lowes", "OverStock"]  # 这些平台的数据将被过滤掉
# PLATFORM_FILTER = ["Wayfair", "Amazon", "OverStock"]  # 这些平台的数据将被过滤掉
PLATFORM_FILTER = []


def find_mapping_files(directory):
    """遍历子目录，查找所有SKU映射.csv和sku映射.xls文件"""
    mapping_files = []

    for sub_folder_name in os.listdir(directory):

        sub_folder_path = os.path.join(directory, sub_folder_name)
        if os.path.isdir(sub_folder_path) and any(keyword in sub_folder_name for keyword in ["组", "站", "院"]):
            if os.path.isdir(sub_folder_path):
                for file in os.listdir(sub_folder_path):
                    if "汇总表" in file:
                        continue
                    if file.lower().startswith('sku映射') and (
                            file.lower().endswith('.csv') or file.lower().endswith('.xls') or file.lower().endswith(
                        '.xlsx')):
                        print(f"正在遍历文件夹: {sub_folder_path}")
                        mapping_files.append(os.path.join(sub_folder_path, file))

    return mapping_files


def check_and_concatenate(files, log_file):
    """检查文件的表头并合并数据"""
    data_frames = []  # 存储所有数据帧

    # 必需的字段集合（不包括来源和仓区/市场）
    required_columns = set(EXPECTED_COLUMNS)

    for file in files:
        try:
            # 根据文件格式读取文件
            if file.endswith('.csv'):
                df = pd.read_csv(file, encoding='gbk')
            else:
                df = pd.read_excel(file)

            # 获取表头并转换为集合进行比较
            file_columns = set(df.columns)

            # 检查是否包含所有必需字段
            missing_columns = required_columns - file_columns
            if missing_columns:
                log_file.write(f"文件 {file} 缺少以下必要字段: {', '.join(missing_columns)}，跳过文件。\n")
                continue

            # 平台筛选：过滤掉在 PLATFORM_FILTER 列表中的平台
            df = df[~df["平台"].isin(PLATFORM_FILTER)]

            # 获取文件夹名称作为"来源"
            source = os.path.basename(os.path.dirname(file))
            df["来源"] = source  # 动态添加"来源"列

            # 检查是否存在"仓区/市场"列，若存在则仅对为空地行进行处理
            if "仓区/市场" in df.columns:
                df["仓区/市场"] = df.apply(
                    lambda row: row["仓区/市场"] if pd.notnull(row["仓区/市场"]) else map_channel_to_market(row["渠道"],
                                                                                                            source),
                    axis=1
                )
            else:
                df["仓区/市场"] = df["渠道"].apply(lambda channel: map_channel_to_market(channel, source))

            # 保留指定的列
            df = df[REQUIRED_COLUMNS]

            # 过滤掉自定义sku为空的行
            df = df.dropna(subset=["自定义sku"])

            # 将处理后的数据添加到列表中
            data_frames.append(df)

        except Exception as e:
            log_file.write(f"处理文件时出错: {file}, 错误信息: {str(e)}\n")

    # 合并数据帧
    if data_frames:
        combined_data = pd.concat(data_frames, ignore_index=True)
    else:
        combined_data = pd.DataFrame()

    return combined_data


def map_channel_to_market(channel, source):
    """
    根据渠道列内容和来源判断仓区
    source: 文件夹名称，即来源
    """
    if not isinstance(channel, str):
        return "未匹配到对应市场"

    # 检查是否有针对特定来源的特殊规则
    if source in special_source_mapping:
        special_market = special_source_mapping[source].get(channel)
        if special_market:
            return special_market

    channel = channel.lower()
    for keyword, market in channel_mapping.items():
        if channel.startswith(keyword) or channel.endswith(keyword) or keyword in channel:
            return market

    return "未匹配到对应市场"


def process_all_subfolders(source_folder, output_file, log_file_path):
    """遍历所有子文件夹并处理每个文件夹中的映射文件"""
    all_mapping_files = []

    with open(log_file_path, 'w', encoding='utf-8') as log_file:
        mapping_files = find_mapping_files(source_folder)
        log_file.write(f"在文件夹 {source_folder} 中找到 {len(mapping_files)} 个映射文件\n")
        all_mapping_files.extend(mapping_files)

        # 合并所有子文件夹中的映射数据
        combined_data = check_and_concatenate(all_mapping_files, log_file)

        if not combined_data.empty:
            # 在这里添加替换代码
            # 将"未匹配到对应市场"替换为"美国仓"
            combined_data["仓区/市场"] = combined_data["仓区/市场"].replace("未匹配到对应市场", "美国仓")

            # 保存最终结果
            combined_data.to_csv(output_file, index=False, encoding='utf-8-sig')
            log_file.write(f"SKU映射汇总表已保存至: {output_file}\n")
            print(f"SKU映射汇总表已保存至: {output_file}")
        else:
            log_file.write("没有成功合并任何数据。\n")
            print("没有成功合并任何数据。")


def main():
    source_folder = r"Y:\1-公共资料查询\9-库存查询\库存查询原始表"
    output_file = r"Y:\1-公共资料查询\9-库存查询\库存查询原始表\配置文件\SKU映射汇总表_加来源test.csv"
    log_file_path = os.path.join(r"Y:\1-公共资料查询\9-库存查询\库存查询原始表\配置文件", "运行日志.txt")

    process_all_subfolders(source_folder, output_file, log_file_path)


if __name__ == "__main__":
    main()
