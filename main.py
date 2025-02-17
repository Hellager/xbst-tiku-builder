from parse import *
from format import *


def main():
    input_folder = input("请输入要扫描的文件夹路径：").strip()
    if os.path.isdir(input_folder):
        title_info = find_title_rows(input_folder)
        count_title_occurrences(title_info)
        count_options_characters(title_info)
        count_answers_characters(title_info)

        continue_build = input("是否确认生成题库(y/n)：").strip()
        if continue_build == "y":
            output_folder = input("请输入要保存的文件夹路径：").strip()
            if os.path.isdir(input_folder):
                if not os.path.exists(output_folder):
                    os.makedirs(output_folder)
                    print(f"创建输出文件夹 {output_folder}")
                else:
                    clear_folder(output_folder)

                copy_and_convert_files(input_folder, output_folder)
                clear_xls_files(output_folder)
                build_formatted_files(output_folder)
            else:
                print("错误：输入的路径不存在或不是文件夹")
        else:
            exit(-1)
    else:
        print("错误：输入的路径不存在或不是文件夹")
        exit(-1)

if __name__ == "__main__":
    main()
