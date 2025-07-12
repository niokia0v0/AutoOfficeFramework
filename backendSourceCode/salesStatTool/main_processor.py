import os
import sys
import argparse
import pandas as pd
import numpy as np
from datetime import datetime

#0.文件结构：确保 main_processor.py, identifier.py, TMProcess.py, JDProcess.py, PDDProcess.py, DYProcess.py 这六个文件都在同一个目录下。
#准备数据：
#创建一个输入目录，例如 C:\my_input_data。
#将所有平台的原始 .csv 和 .xlsx 文件放入这个目录。
#创建一个空的输出目录，例如 D:\my_output_data。

#1.执行命令(在当前路径下)：
#基本用法（默认重命名）：
#python main_processor.py C:\my_input_data D:\my_output_data
#指定为覆盖模式：
#python main_processor.py C:\my_input_data D:\my_output_data --on-conflict overwrite
#指定为跳过模式：
#python main_processor.py C:\my_input_data D:\my_output_data --on-conflict skip
#查看帮助：
#python main_processor.py -h

# 动态添加脚本所在目录到Python路径，以便能找到其他模块
# 这使得脚本可以从任何位置被调用
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# 导入自定义模块
try:
    import identifier
    import TMProcess
    import JDProcess
    import PDDProcess
    import DYProcess
except ImportError as e:
    print(f"错误：无法导入必要的处理模块。请确保所有处理脚本 "
          f"(identifier.py, TMProcess.py, etc.) 与 main_processor.py 在同一目录下。")
    print(f"具体错误: {e}")
    sys.exit(1)

# --- 平台与处理函数的映射 ---
PROCESSOR_MAP = {
    "TM": TMProcess.process_tmall_data,
    "JD": JDProcess.process_jingdong_data,
    "PDD": PDDProcess.process_pdd_data,
    "DY": DYProcess.process_douyin_data,
}

def get_safe_output_path(output_dir, base_filename, on_conflict_policy):
    """
    根据文件冲突策略，计算一个安全的输出文件路径。
    
    Args:
        output_dir (str): 输出目录。
        base_filename (str): 基础输出文件名。
        on_conflict_policy (str): 'skip', 'overwrite', or 'rename'.

    Returns:
        str or None: 如果可以写入，则返回完整的输出路径；如果策略为'skip'且文件存在，则返回None。
    """
    output_path = os.path.join(output_dir, base_filename)
    
    if not os.path.exists(output_path):
        return output_path

    if on_conflict_policy == 'skip':
        print(f"  -> 文件 '{base_filename}' 已存在，策略为【跳过】。")
        return None
        
    if on_conflict_policy == 'overwrite':
        print(f"  -> 文件 '{base_filename}' 已存在，策略为【覆盖】。")
        return output_path

    if on_conflict_policy == 'rename':
        name, ext = os.path.splitext(base_filename)
        counter = 1
        while True:
            new_filename = f"{name} ({counter}){ext}"
            new_path = os.path.join(output_dir, new_filename)
            if not os.path.exists(new_path):
                print(f"  -> 文件 '{base_filename}' 已存在，策略为【重命名】为 '{new_filename}'。")
                return new_path
            counter += 1
    
    return None

def read_dataframe_from_file(file_path):
    """
    根据文件扩展名，从文件读取数据到Pandas DataFrame，并进行基础清洗。
    
    Args:
        file_path (str): 输入文件的完整路径。

    Returns:
        pd.DataFrame or None: 成功则返回DataFrame，失败返回None。
    """
    file_ext = os.path.splitext(file_path)[1].lower()
    df = None
    try:
        if file_ext == '.csv':
            df = pd.read_csv(file_path, dtype=str, keep_default_na=False, encoding='utf-8-sig')
        elif file_ext in ['.xlsx', '.xls']:
            df = pd.read_excel(file_path, dtype=str, engine='openpyxl', keep_default_na=False)
        
        if df is not None:
            # 统一进行基础清洗
            df.columns = [col.strip().replace('"', '') for col in df.columns]
            for col in df.columns:
                if df[col].dtype == 'object':
                    # 使用replace将常见空值字符串替换为np.nan，以便后续处理
                    df[col] = df[col].astype(str).str.strip().replace(
                        ['-', '--', '', 'None', 'nan', '#NULL!', 'null', '\t'], np.nan, regex=False
                    )
            return df

    except Exception as e:
        print(f"  -> 错误：读取文件时发生错误: {e}")
        return None

    return df

def main():
    """
    主执行函数，负责整个处理流程。
    """
    parser = argparse.ArgumentParser(
        description="电商平台订单数据自动化处理工具。",
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument("input_dir", help="包含原始数据文件 (.csv, .xlsx) 的输入目录。")
    parser.add_argument("output_dir", help="用于存放处理后Excel文件的输出目录。")
    parser.add_argument(
        "--on-conflict",
        choices=['skip', 'overwrite', 'rename'],
        default='rename',
        help="当输出文件已存在时的处理策略:\n"
             "  skip:      跳过已存在的文件，不进行处理。\n"
             "  overwrite: 覆盖已存在的文件。\n"
             "  rename:    在文件名后添加序号 (file (1).xlsx) (默认)。"
    )
    args = parser.parse_args()

    # --- 准备工作 ---
    start_time = datetime.now()
    print("-" * 60)
    print(f"处理开始时间: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"输入目录: {os.path.abspath(args.input_dir)}")
    print(f"输出目录: {os.path.abspath(args.output_dir)}")
    print(f"文件冲突策略: {args.on_conflict.upper()}")
    print("-" * 60)

    if not os.path.isdir(args.input_dir):
        print(f"错误：输入目录不存在 -> '{args.input_dir}'")
        sys.exit(1)
        
    os.makedirs(args.output_dir, exist_ok=True)
    
    processed_files = 0
    skipped_files = 0
    failed_files = 0
    
    # --- 核心处理循环 ---
    for filename in sorted(os.listdir(args.input_dir)):
        if filename.lower().endswith(('.csv', '.xlsx', '.xls')):
            input_path = os.path.join(args.input_dir, filename)
            
            print(f"发现文件: '{filename}'")
            
            # 1. 识别平台
            platform = identifier.identify_platform(input_path)
            if not platform:
                print("  -> 平台识别失败，跳过此文件。\n")
                failed_files += 1
                continue
            print(f"  -> 识别为【{platform}】平台。")

            # 2. 计算安全输出路径
            base_name_no_ext = os.path.splitext(filename)[0]
            output_filename = f"{platform}_output_{base_name_no_ext}.xlsx"
            output_path = get_safe_output_path(args.output_dir, output_filename, args.on_conflict)
            
            if output_path is None:
                skipped_files += 1
                print("") 
                continue

            # 3. 读取数据为DataFrame
            print("  -> 正在读取数据...")
            df_raw = read_dataframe_from_file(input_path)
            if df_raw is None:
                print("  -> 读取失败，跳过此文件。\n")
                failed_files += 1
                continue

            # 4. 调用对应的平台处理函数
            print("  -> 正在处理数据...")
            processor_func = PROCESSOR_MAP.get(platform)
            try:
                result_workbook = processor_func(df_raw)
            except Exception as e:
                print(f"  -> 错误: 在处理【{platform}】数据时发生异常: {e}")
                import traceback
                traceback.print_exc() 
                result_workbook = None

            # 5. 保存结果
            if result_workbook:
                print(f"  -> 正在保存到: '{os.path.basename(output_path)}'")
                try:
                    result_workbook.save(output_path)
                    print("  -> 保存成功！\n")
                    processed_files += 1
                except Exception as e:
                    print(f"  -> 错误：保存文件失败: {e}\n")
                    failed_files += 1
            else:
                print("  -> 数据处理失败，未生成结果文件。\n")
                failed_files += 1
    
    # --- 结束总结 ---
    end_time = datetime.now()
    duration = end_time - start_time
    print("-" * 60)
    print("所有文件处理完毕！")
    print(f"总耗时: {duration}")
    print(f"成功处理: {processed_files}个文件")
    print(f"跳过处理: {skipped_files}个文件")
    print(f"处理失败: {failed_files}个文件")
    print("-" * 60)

if __name__ == "__main__":
    main()