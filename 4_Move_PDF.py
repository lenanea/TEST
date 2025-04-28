import os
import shutil
import re
import pandas as pd

# 配置路径
# 配置
A_dir = r"C:\RPA\Finance_001\Download\Renamed"  # 替换为A目录的路径
B_dir = r"C:\RPA\Finance_001\Output"  # 替换为B目录的路径
excel_file = r"C:\RPA\Finance_001\Download\export.XLSX"  # 替换为Excel文件路径
document_number_column = "Document Number"  # Excel中的列名，Document Number
customer_no_column = "Customer No."  # Excel中的列名，Customer No.

# 读取 Excel 文件
print("正在读取 Excel 文件...")
try:
    df = pd.read_excel(excel_file)

    # 确保列名正确，去除空格，并将数据转换为字符串类型
    df[document_number_column] = df[document_number_column].astype(str).str.strip()
    df[customer_no_column] = df[customer_no_column].astype(str).str.strip()
    # print("Excel 文件读取成功，预览数据：")
    # print(df.head())  # 打印前几行数据以便检查
except Exception as e:
    print(f"无法读取 Excel 文件。错误：{e}")
    exit()

# 获取 B 目录的第一层子文件夹名称
# print("正在获取 B 目录的第一层子文件夹...")
try:
    b_folders = [folder for folder in os.listdir(B_dir) if os.path.isdir(os.path.join(B_dir, folder))]
    # print(f"B 目录的子文件夹：{b_folders}")
except Exception as e:
    print(f"无法读取 B 目录。错误：{e}")
    exit()

# 遍历 A 目录下的 PDF 文件
# print("开始遍历 A 目录中的文件...")
for file_name in os.listdir(A_dir):
    if file_name.endswith(".pdf"):  # 判断是否是 PDF 文件
        print(f"正在处理文件：{file_name}")

        # 第一步：提取 PDF 文件名中的 x（第一个加号之前的部分）
        match = re.match(r"([^+]+)\+.*", file_name)  # 匹配文件名格式
        if match:
            x = match.group(1).strip()  # 获取第一个加号之前的部分，并去除可能的空格
            # print(f"提取到的 x 值为：{x}")
        else:
            print(f"文件名 {file_name} 不符合格式，跳过处理。")
            continue

        # 第二步：直接在 Excel 的 Document Number 列中查找 x 对应的 Customer No.
        # print(f"正在查找 {x} 对应的 Customer No...")
        try:
            # 使用布尔索引匹配 Document Number
            y = df.loc[df[document_number_column] == x, customer_no_column]
            if y.empty:
                # print(f"未找到 Document Number 为 {x} 的对应 Customer No.，跳过文件 {file_name}。")
                continue
            else:
                y = y.values[0]  # 获取第一个匹配的值
                print(f"找到的 Customer No. 值为：{y}")
        except Exception as e:
            print(f"查找 Customer No. 时出错：{e}")
            continue

        # 第三步：检查 B 目录中是否存在名称为 y 的子文件夹
        # print(f"检查 B 目录中是否存在名为 {y} 的子文件夹...")
        if y in b_folders:
            print(f"找到匹配的子文件夹：{y}")
        else:
            print(f"未找到匹配的子文件夹 {y}，跳过文件 {file_name}。")
            continue

        # 第四步：移动文件到对应的子文件夹
        source_path = os.path.join(A_dir, file_name)  # A 目录中的文件路径
        target_folder = os.path.join(B_dir, y)  # B 目录中的目标文件夹路径
        target_path = os.path.join(target_folder, file_name)  # 目标文件路径

        print(f"正在将文件 {file_name} 移动到 {target_folder}...")
        try:
            shutil.move(source_path, target_path)
            print(f"文件 {file_name} 已成功移动到 {target_folder}。")
        except Exception as e:
            print(f"移动文件 {file_name} 时出错：{e}")

print("所有文件处理完成！")