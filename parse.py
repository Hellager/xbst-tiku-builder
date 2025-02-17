import os
import pandas as pd
import xlrd
import openpyxl
import xlsxwriter
from collections import Counter
from os import PathLike
from typing import Union, Optional
import string


def convert_xls_to_xlsx(file_path: Union[str, bytes, PathLike[str], PathLike[bytes]],
                        output_path: Union[str, bytes, PathLike[str], PathLike[bytes]],
                        title_row: list):
    try:
        # 使用 xlrd 读取 xls 文件
        df = pd.read_excel(file_path, engine='xlrd')

        # 将标题行数据插入到 DataFrame 的第一行
        if len(df) > 0 and all(df.iloc[0] == title_row):
            # 如果第一行数据与标题行数据相同,则不需要插入
            pass
        else:
            df.loc[-1] = title_row
            df.index = df.index + 1
            df = df.sort_index()

        # 使用 xlsxwriter 将 DataFrame 保存为 xlsx 文件
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, header=False)

        print(f"已将 {file_path} 转换为 {output_path}")
        return output_path
    except Exception as e:
        print(f"【错误】处理文件 {file_path} 时发生异常：{str(e)}")
        return None

def find_title_rows(folder_path):
    title_info = {}
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            # 筛选Excel文件
            if file.lower().endswith(('.xls', '.xlsx')):
                file_path = os.path.join(root, file)
                try:
                    # 检查是否存在同名 xlsx 文件
                    xlsx_file_path = os.path.splitext(file_path)[0] + '.xlsx'
                    if os.path.exists(xlsx_file_path):
                        # 直接读取 xlsx 文件
                        df = pd.read_excel(xlsx_file_path, header=None, engine='openpyxl')
                        file_path = xlsx_file_path
                    else:
                        # 将 xls 文件转换为 xlsx 文件
                        output_path = xlsx_file_path

                        # 读取 xls 文件并获取标题行数据
                        df = pd.read_excel(file_path, header=None, engine='xlrd')
                        title_row = []
                        for _, row in df.iterrows():
                            # 检查第一列并去除前后空格
                            first_col = str(row[0]).strip() if pd.notna(row[0]) else ""
                            if first_col == "序号":
                                # 处理整行数据并格式化为字符串
                                title_row = row.fillna("").astype(str).tolist()
                                break
                        new_file_path = convert_xls_to_xlsx(file_path, output_path, title_row)
                        file_path = new_file_path
                        df = pd.read_excel(output_path, header=None, engine='openpyxl')

                    # 遍历所有行
                    for _, row in df.iterrows():
                        # 检查第一列并去除前后空格
                        first_col = str(row[0]).strip() if pd.notna(row[0]) else ""
                        if first_col == "序号":
                            # 处理整行数据并格式化为字符串
                            title_row = row.fillna("").astype(str).tolist()
                            title_str = ", ".join(title_row)
                            title_info[file_path] = title_str
                            # print(f"{rel_path} - {title_str}")
                except Exception as e:
                    print(f"【错误】处理文件 {file_path} 时发生异常：{str(e)}")
    return title_info

def count_title_occurrences(title_info):
    """
    统计所有文件对应的列名
    """
    title_counts = Counter(title_info.values())
    for title, count in title_counts.items():
        print(f"列名: {title}, 出现次数: {count}")


def count_options_characters(title_info):
    """
    统计所有文件对应的 A后字符并打印出字符对应统计数量
    """
    char_counts = {}
    for file_path in title_info.keys():
        # 使用 pandas 解析 xlsx 文件,找到行首为 "选项" 列,列中数据应为 Axxx 格式
        df = pd.read_excel(file_path, header=None, engine='openpyxl')
        option_col = None
        for col in df.columns:
            if df[col].str.contains("选项").any():
                option_col = col
                break
        if option_col is not None:
            option_rows = df[df[option_col] == "选项"].index
            for row in option_rows:
                option = df.iloc[row + 1, df.columns.get_loc(option_col)]
                if str(option).startswith("A"):
                    char = str(option)[1]
                    if char in char_counts:
                        char_counts[char] += 1
                    else:
                        char_counts[char] = 1

    # 打印出字符对应统计数量
    for char, count in char_counts.items():
        print(f"选项后字符 {char} 出现了 {count} 次")

def count_answers_characters(title_info):
    """
    统计所有文件对应的答案后字符并打印出字符对应统计数量
    """
    char_counts = {}
    for file_path in title_info.keys():
        # 使用 pandas 解析 xlsx 文件,找到行首为 "选项" 列,列中数据应为 Axxx 格式
        df = pd.read_excel(file_path, header=None, engine='openpyxl')
        option_col = None
        for col in df.columns:
            if df[col].str.contains("答案").any():
                option_col = col
                break

        if option_col is not None:
            for value in df.iloc[:, option_col]:
                if isinstance(value, str) and value.strip() and value[0] in string.ascii_letters:
                    if len(str(value)) > 1:
                        char = str(value)[1]
                        if char in char_counts:
                            char_counts[char] += 1
                        else:
                            char_counts[char] = 1
                        break
                    else:
                        continue  # 跳过当前文件,开始遍历下一个文件

    # 打印出字符对应统计数量
    for char, count in char_counts.items():
        print(f"答案第二个字符 {char} 出现了 {count} 次")
