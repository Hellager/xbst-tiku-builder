'''
格式应为 考题 答案 选项A 选项B ....
'''

import os
import shutil
import pandas as pd
import numpy as np
import re


def copy_and_convert_files(input_folder, output_folder):
    """
    遍历输入文件夹及其子文件夹中的文件,若文件后缀为 xlsx,则直接复制到输出文件夹,
    若文件后缀为 xls,则使用 pandas 读取文件内容后,将内容保存至输出文件夹的同名文件.
    """
    total_files = 0
    copied_files = 0
    converted_files = 0

    for root, dirs, files in os.walk(input_folder):
        for filename in files:
            input_file = os.path.join(root, filename)
            output_file = os.path.join(output_folder, os.path.relpath(input_file, input_folder))
            total_files += 1

            # 检查文件后缀
            if filename.endswith('.xlsx'):
                # 直接复制 xlsx 文件
                os.makedirs(os.path.dirname(output_file), exist_ok=True)
                shutil.copy(input_file, output_file)
                copied_files += 1
            elif filename.endswith('.xls'):
                # 读取 xls 文件内容,并保存为 xlsx 文件
                os.makedirs(os.path.dirname(output_file), exist_ok=True)
                df = pd.read_excel(input_file, engine='xlrd')
                df.to_excel(output_file, index=False, engine='openpyxl')
                converted_files += 1

    print(f"总共处理了 {total_files} 个文件")
    print(f"成功复制了 {copied_files} 个 .xlsx 文件")
    print(f"成功转换了 {converted_files} 个 .xls 文件")

def clear_xls_files(folder_path):
    """
    遍历给定文件夹及其子目录中的所有 .xls 文件,并删除这些文件。
    """
    for root, dirs, files in os.walk(folder_path):
        for filename in files:
            if filename.endswith('.xls'):
                file_path = os.path.join(root, filename)
                os.remove(file_path)
                print(f"已删除文件: {os.path.relpath(file_path, folder_path)}")

# 定义清洗函数
def clean_option(opt):
    """ 优化后的选项清洗逻辑 """
    s = str(opt)  # 确保转换为字符串
    # 步骤1：去除首字符
    if len(s) >= 1:
        s = s[1:]
    # 步骤2：检查新首字符
    if len(s) >= 1 and s[0] in {'.', '-', '、'}:
        s = s[1:]
    # 步骤3：去除两端空格
    return s.strip()

def build_formatted_files(output_folder):
    """
    遍历 output_folder 中的 xlsx 文件，处理并保存。处理步骤：
    1. 将 '题目名称' 列重命名为 '考题' 并置为第一列
    2. 将 '答案' 列移至第二列
    3. 根据 '选项' 列中的 '|' 数量动态创建选项列
    4. 处理选项内容，去除前导符号
    """
    total_files = 0
    formatted_files = 0
    for root, dirs, files in os.walk(output_folder):
        for filename in files:
            if filename.endswith('.xlsx'):
                total_files += 1
                file_path = os.path.abspath(os.path.join(root, filename))
                if not os.path.exists(file_path):
                    print(f"文件 {file_path} 不存在，跳过")
                    continue
                try:
                    # 读取Excel，不自动识别表头
                    df = pd.read_excel(file_path, header=None, engine="openpyxl", dtype=str)

                    # 寻找列名所在行（含'序号'的行）
                    header_row = None
                    for i, row in df.iterrows():
                        if '序号' in row.values:
                            header_row = i
                            break
                    if header_row is None:
                        print(f"无法识别列名所在行，跳过文件: {filename}")
                        continue

                    # 设置列名并截取数据
                    df.columns = df.iloc[header_row]
                    df = df.drop(header_row).reset_index(drop=True)

                    # 步骤1：重命名列
                    df.rename(columns={'题目名称': '考题'}, inplace=True)

                    # 步骤3-4：处理选项列
                    if '选项' in df.columns:
                        # 确保转换为有效列表
                        df['选项'] = df['选项'].fillna('').apply(
                            lambda x: [
                                part.strip()  # 去除每个选项首尾空格
                                for part in re.sub(r'[\|\n]+', '|', str(x).strip())  # 合并所有分隔符
                                .strip('|')  # 去除首尾分隔符
                                .split('|')
                                if part.strip()  # 过滤空内容
                            ] if pd.notna(x) else []
                        )

                        # 计算最大选项数（保留非空判断）
                        non_empty = df['选项'][df['选项'].apply(len) > 0]
                        max_options = non_empty.apply(len).max() if not non_empty.empty else 0

                        option_cols = [f'选项{chr(65 + i)}' for i in range(max_options)]

                        # 应用新的清洗函数
                        df[option_cols] = df['选项'].apply(
                            lambda x: [clean_option(y) for y in x[:max_options]]  # 新清洗逻辑
                                      + [''] * (max_options - len(x))  # 填充空值
                        ).tolist()

                        df = df.drop('选项', axis=1)
                    else:
                        option_cols = []

                    # 调整列顺序：考题、答案 + 动态选项列
                    required_cols = ['考题', '答案'] + option_cols
                    # 确保其他列不被丢弃（可根据需求调整）
                    # other_cols = [col for col in df.columns if col not in required_cols]
                    df = df[required_cols]

                    # 在最终写入Excel文件前添加以下代码
                    if not df.empty:
                        # 创建处理副本
                        temp_df = df.copy()

                        # 获取前两列（自动处理列数不足的情况）
                        cols_to_check = temp_df.columns[:2]

                        # 初始化空值条件
                        empty_condition = pd.Series(False, index=temp_df.index)

                        for col in cols_to_check:
                            # 处理空值并去除空格（兼容NaN和空白字符串）
                            col_series = temp_df[col].fillna('').astype(str).str.strip()
                            # 累积空值条件（逻辑或）
                            empty_condition |= (col_series == '')

                        # 反向筛选并重置索引
                        df = df[~empty_condition].reset_index(drop=True)

                    # 新增答案列处理（添加在选项处理之后）
                    if '答案' in df.columns:
                        df['答案'] = df['答案'].fillna('').apply(
                            lambda x: ''.join(re.findall(r'^[A-Za-z]+', str(x))).upper()
                        )

                        # 将空字符串转为NaN并删除对应行
                        df['答案'] = df['答案'].replace('', np.nan)
                        df.dropna(subset=['答案'], inplace=True)

                    # 保存文件
                    df.to_excel(file_path, index=False)
                    formatted_files += 1
                    print(f"已处理并保存文件: {filename}")
                except Exception as e:
                    print(f"处理文件 {filename} 时出错: {e}")

    print(f"总共遍历了 {total_files} 个文件，处理并保存了 {formatted_files} 个文件")


def clear_folder(folder_path):
    """
    清空指定文件夹中的所有内容和子目录，保留空文件夹本身

    参数：
    folder_path (str): 要清空的文件夹路径

    异常：
    ValueError: 当路径不存在或不是目录时抛出
    """
    # 验证路径有效性
    if not os.path.exists(folder_path):
        raise ValueError(f"路径不存在: {folder_path}")
    if not os.path.isdir(folder_path):
        raise ValueError(f"路径不是目录: {folder_path}")

    # 遍历并删除内容
    for entry in os.listdir(folder_path):
        full_path = os.path.join(folder_path, entry)
        try:
            if os.path.isfile(full_path) or os.path.islink(full_path):
                os.unlink(full_path)  # 删除文件或符号链接
            else:
                shutil.rmtree(full_path)  # 递归删除子目录
        except Exception as e:
            print(f"删除失败 {full_path}: {str(e)}")

def main():
    # input_folder = input("请输入要扫描的文件夹路径：").strip()
    # output_folder = input("请输入要保存的文件夹路径：").strip()

    input_folder = "D:\\Project\\Gitee\\build_tiku_for_souti\\raw\\2025版题库"
    output_folder = "D:\\Project\\Gitee\\build_tiku_for_souti\\output\\2025版题库"

    if os.path.isdir(input_folder):
        # 检查输出文件夹是否存在,如果不存在则创建
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
            print(f"创建输出文件夹 {output_folder}")
        else:
            # 清空输出文件夹
            clear_folder(output_folder)

        copy_and_convert_files(input_folder, output_folder)
        clear_xls_files(output_folder)
        build_formatted_files(output_folder)
    else:
        print("错误：输入的路径不存在或不是文件夹")
