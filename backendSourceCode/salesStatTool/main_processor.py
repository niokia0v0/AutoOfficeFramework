import os
import sys
import argparse
import pandas as pd
import numpy as np
from datetime import datetime

# 动态添加脚本所在目录到Python路径
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
          f"(identifier.py, TMProcess.py, etc.) 与 main_processor.py 在同一目录下。", flush=True)
    print(f"具体错误: {e}", flush=True)
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
        print(f"  -> 文件 '{base_filename}' 已存在，策略【跳过】。", flush=True)
        return None
        
    if on_conflict_policy == 'overwrite':
        print(f"  -> 文件 '{base_filename}' 已存在，策略【覆盖】。", flush=True)
        return output_path

    if on_conflict_policy == 'rename':
        name, ext = os.path.splitext(base_filename)
        counter = 1
        while True:
            new_filename = f"{name} ({counter}){ext}"
            new_path = os.path.join(output_dir, new_filename)
            if not os.path.exists(new_path):
                print(f"  -> 文件 '{base_filename}' 已存在，策略【重命名】，另存为 '{new_filename}'。", flush=True)
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
            try:
                # 优先尝试用 utf-8-sig 读取
                df = pd.read_csv(file_path, dtype=str, keep_default_na=False, encoding='utf-8-sig')
            except UnicodeDecodeError:
                # 如果UTF-8解码失败，则回退到GBK编码再次尝试
                print(f"  -> UTF-8解码失败，尝试使用GBK编码读取完整文件...")
                df = pd.read_csv(file_path, dtype=str, keep_default_na=False, encoding='gbk')
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
        print(f"  -> 错误：读取文件时发生错误: {e}", flush=True)
        return None

    return df

def main():
    """
    主执行函数，负责整个处理流程。
    """
    parser = argparse.ArgumentParser(
        description="电商平台销售数据处理工具。",
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument("input_dir", help="包含原始数据文件 (.csv, .xlsx) 的输入目录。")
    parser.add_argument("output_dir", help="用于存放处理后Excel文件的输出目录。")
    parser.add_argument(
        "--on-conflict",
        choices=['skip', 'overwrite', 'rename'],
        default='skip',
        help="当输出文件已存在时的处理策略:\n"
             "  skip:      跳过已存在的文件，不进行处理。\n"
             "  overwrite: 覆盖已存在的文件。\n"
             "  rename:    在文件名后添加序号 (file (1).xlsx) (默认)。"
    )
    args = parser.parse_args()

    # --- 准备工作 ---
    start_time = datetime.now()
    print("-" * 60, flush=True)
    print(f"处理开始时间: {start_time.strftime('%Y-%m-%d %H:%M:%S')}", flush=True)
    print(f"输入目录: {os.path.abspath(args.input_dir)}", flush=True)
    print(f"输出目录: {os.path.abspath(args.output_dir)}", flush=True)
    print(f"文件冲突策略: {args.on_conflict.upper()}", flush=True)
    print("-" * 60, flush=True)

    if not os.path.isdir(args.input_dir):
        print(f"错误：输入目录不存在 -> '{args.input_dir}'", flush=True)
        sys.exit(1)
        
    os.makedirs(args.output_dir, exist_ok=True)
    
    processed_files = 0
    skipped_files = 0
    failed_files = 0
    unidentified_files = 0
    
    # --- 核心处理循环 ---
    for filename in sorted(os.listdir(args.input_dir)):
        if filename.lower().endswith(('.csv', '.xlsx', '.xls')):
            input_path = os.path.join(args.input_dir, filename)
            
            print(f"发现文件: '{filename}'", flush=True)
            
            # 1. 识别平台
            platform = identifier.identify_platform(input_path)
            if not platform:
                print("  -> 平台识别失败，跳过此文件。\n", flush=True)
                unidentified_files += 1
                continue
            print(f"  -> 识别为【{platform}】平台。", flush=True)

            # 2. 计算安全输出路径
            base_name_no_ext = os.path.splitext(filename)[0]
            
            # 根据平台标识符构建不同的输出文件名
            if platform == "TM_RECENT":
                output_filename = f"TM_output_recent_{base_name_no_ext}.xlsx"
            elif platform == "TM_HISTORY":
                output_filename = f"TM_output_history_{base_name_no_ext}.xlsx"
            else:
                output_filename = f"{platform}_output_{base_name_no_ext}.xlsx"

            output_path = get_safe_output_path(args.output_dir, output_filename, args.on_conflict)
            
            if output_path is None:
                skipped_files += 1
                print("", flush=True) 
                continue

            # 3. 读取数据为DataFrame
            print("  -> 正在读取数据...", flush=True)
            df_raw = read_dataframe_from_file(input_path)
            if df_raw is None:
                print("  -> 读取失败，跳过此文件。\n", flush=True)
                failed_files += 1
                continue

            # 4. 调用对应的平台处理函数
            print("  -> 正在处理数据...", flush=True)
            # 从具体标识符（如 'TM_RECENT'）中提取基础平台名（'TM'）用于查找处理器
            base_platform = platform.split('_')[0]
            processor_func = PROCESSOR_MAP.get(base_platform)
            
            if not processor_func:
                print(f"  -> 错误：未找到平台 '{base_platform}' 对应的处理器，跳过此文件。\n", flush=True)
                unidentified_files += 1
                continue
                
            try:
                result_workbook = processor_func(df_raw)
            except Exception as e:
                print(f"  -> 错误: 在处理【{platform}】数据时发生异常: {e}", flush=True)
                import traceback
                # 异常堆栈信息也需要flush
                traceback.print_exc(file=sys.stdout)
                sys.stdout.flush()
                result_workbook = None

            # 5. 保存结果
            if result_workbook:
                print(f"  -> 正在保存到: '{os.path.basename(output_path)}'", flush=True)
                try:
                    result_workbook.save(output_path)
                    print("  -> 保存成功！\n", flush=True)
                    processed_files += 1
                except Exception as e:
                    print(f"  -> 错误：保存文件失败: {e}\n", flush=True)
                    failed_files += 1
            else:
                print("  -> 数据处理失败，未生成结果文件。\n", flush=True)
                failed_files += 1
    
    # --- 结束总结 ---
    end_time = datetime.now()
    duration = end_time - start_time
    print("-" * 60, flush=True)
    print("所有文件处理完毕！", flush=True)
    print(f"处理总耗时: {duration}", flush=True)
    print(f"成功处理: {processed_files}个文件", flush=True)
    print(f"跳过处理 (同名文件): {skipped_files}个文件", flush=True)
    print(f"平台未识别: {unidentified_files}个文件", flush=True)
    print(f"处理失败 (错误): {failed_files}个文件", flush=True)
    print("-" * 60, flush=True)

if __name__ == "__main__":
    main()