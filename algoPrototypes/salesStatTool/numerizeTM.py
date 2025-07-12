import pandas as pd
import os
import numpy as np # 虽然可能不直接用，但pd.to_numeric可能产生NaN，numpy与NaN相关

def numerize_excel_columns(input_file_path, columns_to_numerize, output_suffix="_numerized"):
    """
    读取指定的Excel文件，将其中特定列的内容转换为纯数字格式，并输出新的Excel文件。

    参数:
    input_file_path (str): 输入的原始Excel文件的完整路径。
    columns_to_numerize (list): 一个包含列名的列表，这些列将被尝试转换为数字。
                                例如: ['购买数量', '买家实付金额', '退款金额']
    output_suffix (str): 添加到原始文件名前缀以构成输出文件名的新后缀。

    返回:
    str: 生成的Excel文件的完整路径，如果成功。
    None: 如果发生错误。
    """
    # --- 1. 检查输入文件是否存在 ---
    if not os.path.exists(input_file_path):
        print(f"错误: 输入文件未找到于 '{os.path.abspath(input_file_path)}'")
        return None
    if not (input_file_path.lower().endswith(".xlsx") or input_file_path.lower().endswith(".xls")):
        print(f"错误: 输入文件 '{os.path.basename(input_file_path)}' 不是一个有效的Excel文件。")
        return None

    # --- 2. 构建输出文件路径 ---
    input_dir = os.path.dirname(input_file_path)
    base_name = os.path.basename(input_file_path)
    file_name_without_ext = os.path.splitext(base_name)[0]
    output_file_path = os.path.join(input_dir, f"{file_name_without_ext}{output_suffix}.xlsx")

    try:
        # --- 3. 读取Excel文件 ---
        # 先将所有内容作为字符串读取，以避免Excel自动转换可能导致的问题，
        # 并且可以处理列名中的空格等。
        try:
            df = pd.read_excel(input_file_path, dtype=str)
        except UnicodeDecodeError: # 备用引擎
            df = pd.read_excel(input_file_path, dtype=str, engine='openpyxl')
        
        if df.empty:
            print(f"警告: 输入文件 '{base_name}' 为空或未能读取到数据。")
            # 仍然尝试保存一个空的（或只有表头的）输出文件，或者根据需要返回None
            # df.to_excel(output_file_path, index=False)
            # return output_file_path
            return None


        # 清理列名中的潜在空格
        df.columns = df.columns.str.strip()

        # --- 4. 转换指定列为数字 ---
        converted_cols_count = 0
        for col_name in columns_to_numerize:
            stripped_col_name = col_name.strip() # 确保列表中的列名也去除空格
            if stripped_col_name in df.columns:
                # 清理该列单元格数据：去除两端空格，并将多种常见空值表示统一替换为Pandas的NaN
                # 这一步对于to_numeric很重要，因为带空格的" 123 "可能转换失败或得到错误结果
                df[stripped_col_name] = df[stripped_col_name].str.strip()
                df[stripped_col_name] = df[stripped_col_name].replace(
                    ['--', '', 'None', 'nan', '#NULL!', None], np.nan, regex=False
                )
                
                # 尝试将列转换为数字，无法转换的将变为NaN
                df[stripped_col_name] = pd.to_numeric(df[stripped_col_name], errors='coerce')
                
                # 可选：将转换后产生的NaN值替换为0，这样手动求和时不会因NaN出错
                # 如果希望保留NaN以标识原始数据问题，可以注释掉下面这行
                df[stripped_col_name] = df[stripped_col_name].fillna(0)
                
                print(f"列 '{stripped_col_name}' 已尝试转换为数字格式 (非数字转为0)。")
                converted_cols_count += 1
            else:
                print(f"警告: 指定的列 '{stripped_col_name}' 在Excel文件中未找到，将跳过转换。")
        
        if converted_cols_count == 0 and columns_to_numerize:
            print("警告: 没有一列被成功找到并转换。请检查 `columns_to_numerize` 列表中的列名是否与Excel表头完全匹配。")

        # --- 5. 保存处理后的DataFrame到新的Excel文件 ---
        # index=False 表示不将DataFrame的索引写入到Excel文件中
        df.to_excel(output_file_path, index=False, engine='openpyxl') # 使用 openpyxl 引擎通常能更好地保留格式
        
        print(f"\n处理完成。转换后的文件已保存到: {output_file_path}")
        return output_file_path

    except Exception as e:
        print(f"处理过程中发生错误: {e}")
        import traceback
        traceback.print_exc()
        return None

# ---- 主程序入口 ----
if __name__ == "__main__":
    # --- 用户配置区域 ---
    # 1. 指定原始Excel文件所在的目录和文件名
    directory_path = r"F:\étude\Ecole\E4\E4stage\E4stageProjet\samples\TM"  # 替换为您的目录路径
    original_xlsx_filename = "ExportOrderList24220674019.xlsx"  # 替换为您的原始文件名


    # 2. 列出需要转换为纯数字的列的准确名称 (必须与Excel表头完全一致)
    cols_to_make_numeric = [
        "购买数量",         # 对应原始E列
        "商品价格",         # 对应原始D列
        "买家实付金额",     # 对应原始M列
        "退款金额",         # 对应原始O列
        "买家应付货款",     # 对应原始L列 (如果也需要转换)
        # "其他可能需要转换的金额或数量列..."
    ]
    # --- 用户配置结束 ---

    input_full_path = os.path.join(directory_path, original_xlsx_filename)
    
    print(f"开始处理文件: {input_full_path}")
    print(f"将尝试转换以下列为数字: {', '.join(cols_to_make_numeric)}")

    output_file = numerize_excel_columns(input_full_path, cols_to_make_numeric)

    if output_file:
        print("脚本执行完毕。")
    else:
        print("脚本执行过程中发生错误，未能生成输出文件。")