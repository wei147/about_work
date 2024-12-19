
import os
import pandas as pd
from datetime import datetime


def allegro_update_inventory_status(input_file_path):
    """
    更新库存文件的函数，保留原文件格式并将结果保存为同样的文件名。

    参数:
        input_file_path (str): 需要更新的库存文件路径。

    返回:
        str: 更新后的文件路径（即原文件路径）。
    """
    # 读取整个文件，保留表头前的所有数据
    with open(input_file_path, 'r', encoding='utf-8') as f:
        # 读取表头前的所有内容（假设表头在第4行，从0计数就是第3行）
        all_lines = f.readlines()

    # 将表头的行号设定为第四行，从0计数就是第3行
    header_line = 3

    # 从第四行（即数据行）开始读取CSV内容
    df = pd.read_csv(input_file_path, header=header_line)

    # 根据 'Quantity' 列修改 'Offer Status' 列
    df['Offer Status'] = df['Quantity'].apply(lambda x: 'Ended' if x < 10 else 'Active')

    # 在所有处理的行中，'Action' 列设置为 'Edit all'
    df['Action'] = 'Edit all'

    # 将修改后的数据保存到临时CSV文件
    result_file = 'temp_result.csv'
    df.to_csv(result_file, index=False)

    # 读取处理后的内容
    with open(result_file, 'r', encoding='utf-8') as f:
        result_lines = f.readlines()

    # 将表头前的数据和处理后的数据重新组合
    with open(input_file_path, 'w', encoding='utf-8') as f:
        f.writelines(all_lines[:header_line + 1])  # 保留表头前的内容和表头行
        f.writelines(result_lines[1:])  # 写入处理后的数据，跳过重复的表头行

    # 删除临时文件
    os.remove(result_file)

    print(f"Updated file saved as: {input_file_path}")
    return input_file_path


def process_allegro_files_in_folder():
    current_date = datetime.now().strftime("%Y-%m-%d")
    folder_path = rf"Y:\1-公共资料查询\9-库存查询\库存查询结果表\{current_date}\欧洲组-梁月华"  # 替换为实际文件夹路径
    # process_allegro_files_in_folder(folder_path)
    """
    在指定文件夹中查找包含 'Allegro' 且后缀为 .csv 的文件，并进行库存更新处理。

    参数:
        folder_path (str): 文件夹路径。
    """
    # 遍历文件夹中的所有文件
    for file_name in os.listdir(folder_path):
        # 检查文件名是否包含 "Allegro" 且后缀为 .csv
        if "Allegro" in file_name and file_name.endswith(".csv"):
            input_file_path = os.path.join(folder_path, file_name)
            print(f"Processing file: {input_file_path}")

            # 调用 allegro_update_inventory_status 函数进行处理

            allegro_update_inventory_status(input_file_path)



if __name__ == '__main__':
    process_allegro_files_in_folder()



