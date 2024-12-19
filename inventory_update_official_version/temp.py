
import pandas as pd

input_file = r"Y:\1-公共资料查询\10-RPA结果查询\运营中心\抓取tribesigns独立站评论数据到速卖通\速卖通导入模版 - 副本.xlsx"  # 替换为你的输入Excel文件路径
output_file = r"Y:\1-公共资料查询\10-RPA结果查询\运营中心\抓取tribesigns独立站评论数据到速卖通\速卖通导入模版 - 副本result.xlsx"  # 替换为你希望输出的Excel文件路径

# def check_duplicate_comments(input_excel, output_excel):
#     # 读取Excel文件到一个新的DataFrame中
#     df = pd.read_excel(input_excel)
#
#     # 确保DataFrame有'商品ID'和'评论'这两列
#     required_columns = ['AE商品ID', '评价内容']
#     if not all(column in df.columns for column in required_columns):
#         raise ValueError(f"Excel文件中必须包含以下列：{required_columns}")
#
#     # 创建一个新列来存储标识信息，初始值为空字符串
#     df['标识'] = ''
#
#     # 使用groupby和apply来检查重复评论，并设置标识
#     def mark_duplicates(group):
#         # 对评论进行去重检查，找出重复项的索引（这里用布尔值表示）
#         duplicates = group['评价内容'].duplicated(keep=False)
#         # 对重复项设置标识信息
#         group.loc[duplicates, '标识'] = '已存在相同的评论'
#         return group
#
#     # 按'商品ID'分组，并应用上述函数
#     df = df.groupby('AE商品ID', group_keys=False).apply(mark_duplicates)
#
#     # 将结果写入新的Excel文件，不改变原始文件
#     df.to_excel(output_excel, index=False)
#
#
# check_duplicate_comments(input_file, output_file)


# 读取Excel文件时，将'商品id'列作为字符串读取，以避免精度丢失
# 同时读取所有列，不改变其他列的数据类型
df = pd.read_excel(input_file, dtype={'AE商品ID': str})

# 检查是否包含必要的列

# 新增一个标识列，默认为空。确保此列在最后
df['标识'] = ''

# 使用groupby和duplicated检查重复评论
# `keep='first'` 表示保留第一次出现，不标识
duplicates = df.duplicated(subset=['AE商品ID', '评价内容'], keep='first')

# 将标识填充为“已存在相同的评论”对于重复的行
df.loc[duplicates, '标识'] = '已存在相同的评论'

# 保存到新的Excel文件，指定 `index=False` 以避免保存索引
df.to_excel(output_file, index=False)

print(f"处理完成，结果已保存到 {output_file}")
