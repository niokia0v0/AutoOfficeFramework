import os
import sys
import argparse
import pandas as pd
import numpy as np
from datetime import datetime
import io
import traceback

#解决前后端通信时的编码问题。
sys.stdin = io.TextIOWrapper(sys.stdin.buffer, encoding='utf-8')
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='surrogateescape')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='surrogateescape')

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
    print(f"错误：无法导入必要的处理模块。", file=sys.stderr, flush=True)
    print(f"具体错误: {e}", file=sys.stderr, flush=True)
    sys.exit(1)

# 平台处理工具映射表。
# 这里的键必须与 identifier.py 中 PLATFORM_FINGERPRINTS 的键完全对应。
# 由于 TMProcess.py 内部已能够区分"TM_RECENT"、"TM_HISTORY"这两种格式，所以它们可以指向同一个处理函数。
PROCESSOR_MAP = {
    "TM_RECENT": TMProcess.process_tmall_data,  # 天猫近期订单 (.xlsx)
    "TM_HISTORY": TMProcess.process_tmall_data, # 天猫历史订单 (.csv)
    "JD": JDProcess.process_jingdong_data,      # 京东
    "PDD": PDDProcess.process_pdd_data,         # 拼多多
    "DY": DYProcess.process_douyin_data,        # 抖店
}

# --- 协议与通信函数 ---
def send_status_update(file_path, status, message=""):
    """
    向标准输出发送格式化的状态更新信息，供前端GUI解析。
    """
    # 打印格式化的字符串，并立即刷新缓冲区，确保前端能实时收到。
    print(f"##STATUS##|{file_path}|{status}|{message}", flush=True)

# --- 文件处理辅助函数 ---
def get_safe_output_path(output_dir, input_filename, platform, on_conflict_policy):
    """
    根据文件冲突策略，计算一个安全的输出文件路径。
    """
    # 从输入文件名中分离出基础名和扩展名
    base_name_no_ext = os.path.splitext(input_filename)[0]
    
    # 根据平台标识符构建不同的输出文件名，以区分结果
    if platform == "TM_RECENT":
        output_filename = f"TM_recent_output_{base_name_no_ext}.xlsx"
    elif platform == "TM_HISTORY":
        output_filename = f"TM_history_output_{base_name_no_ext}.xlsx"
    else:
        output_filename = f"{platform}_output_{base_name_no_ext}.xlsx"
        
    # 组合成完整路径
    output_path = os.path.join(output_dir, output_filename)
    
    if not os.path.exists(output_path):
        return output_path

    # 根据冲突策略进行处理
    if on_conflict_policy == 'skip':
        return None # 返回None表示跳过
        
    if on_conflict_policy == 'overwrite':
        return output_path # 直接返回原路径，后续操作会覆盖

    if on_conflict_policy == 'rename':
        # 循环尝试在文件名后添加序号 (1), (2), ... 直到找到一个不冲突的名称
        name, ext = os.path.splitext(output_filename)
        counter = 1
        while True:
            new_filename = f"{name} ({counter}){ext}"
            new_path = os.path.join(output_dir, new_filename)
            if not os.path.exists(new_path):
                return new_path
            counter += 1
    
    return None

def read_dataframe_from_file(file_path):
    """
    根据文件扩展名，从文件读取数据到Pandas DataFrame，并进行基础清洗。
    """
    file_ext = os.path.splitext(file_path)[1].lower()
    df = None
    try:
        # 根据扩展名选择不同的读取方式
        if file_ext == '.csv':
            try:
                # 优先尝试utf-8-sig
                df = pd.read_csv(file_path, dtype=str, keep_default_na=False, encoding='utf-8-sig')
            except UnicodeDecodeError:
                # 如果失败，回退到GBK编码
                print(f"  -> {os.path.basename(file_path)}: UTF-8解码失败，尝试GBK编码...", flush=True)
                df = pd.read_csv(file_path, dtype=str, keep_default_na=False, encoding='gbk')
        elif file_ext in ['.xlsx', '.xls']:
            df = pd.read_excel(file_path, dtype=str, engine='openpyxl', keep_default_na=False)
        
        if df is not None:
            # 对读取到的数据进行统一的基础清洗
            df.columns = [col.strip().replace('"', '') for col in df.columns]
            for col in df.columns:
                if df[col].dtype == 'object':
                    df[col] = df[col].astype(str).str.strip().replace(
                        ['-', '--', '', 'None', 'nan', '#NULL!', 'null', '\t'], np.nan, regex=False
                    )
            return df

    except Exception as e:
        print(f"  -> 错误：读取文件 '{os.path.basename(file_path)}' 时发生错误: {e}", flush=True)
        return None

    return df

# --- 主逻辑函数 ---
def main():
    """
    主执行函数，负责整个处理流程。
    """
    # 1. 解析命令行参数
    parser = argparse.ArgumentParser(description="电商平台销售数据处理后端引擎。")
    parser.add_argument("--output-dir", help="输出目录。如果未提供，则输出到源文件所在目录。")
    parser.add_argument("--on-conflict", choices=['skip', 'overwrite', 'rename'], default='rename', help="文件冲突处理策略。")
    args = parser.parse_args()

    # 2. 打印启动信息
    start_time = datetime.now()
    print("-" * 60, flush=True)
    print(f"后端引擎启动: {start_time.strftime('%Y-%m-%d %H:%M:%S')}", flush=True)
    output_mode = f"指定目录: {os.path.abspath(args.output_dir)}" if args.output_dir else "源文件目录模式"
    print(f"输出模式: {output_mode}", flush=True)
    print(f"文件冲突策略: {args.on_conflict.upper()}", flush=True)
    print("等待从前端接收任务列表...", flush=True)
    print("-" * 60, flush=True)

    # 初始化一个字典，用于统计各种处理结果的数量。
    # 键是状态码（与send_status_update中使用的保持一致），值是计数。
    status_counts = {
        "SUCCESS": 0,
        "SKIPPED": 0,
        "UNIDENTIFIED": 0,
        "FAILURE": 0,
    }

    # 3. 核心处理循环：从标准输入逐行读取文件路径
    for file_path in map(str.strip, sys.stdin):
        if not file_path:
            continue # 跳过空行

        print(f"\n开始处理任务: '{os.path.basename(file_path)}'", flush=True)
        send_status_update(file_path, "PROCESSING", "开始处理...")

        # 3.1 识别平台
        platform = identifier.identify_platform(file_path)
        if not platform:
            print("  -> 平台识别失败，跳过此文件。", flush=True)
            send_status_update(file_path, "UNIDENTIFIED", "未能识别平台类型")
            status_counts["UNIDENTIFIED"] += 1
            continue
        print(f"  -> 识别为【{platform}】平台。", flush=True)

        # 3.2 计算输出路径
        output_dir = args.output_dir if args.output_dir else os.path.dirname(file_path)
        os.makedirs(output_dir, exist_ok=True)
        output_path = get_safe_output_path(output_dir, os.path.basename(file_path), platform, args.on_conflict)
        
        if output_path is None:
            skipped_name = os.path.basename(get_safe_output_path(output_dir, os.path.basename(file_path), platform, 'rename')).replace(' (1)', '')
            print("  -> 输出文件已存在，根据策略跳过。", flush=True)
            send_status_update(file_path, "SKIPPED", f"文件 '{skipped_name}' 已存在")
            status_counts["SKIPPED"] += 1
            continue

        # 3.3 读取数据
        df_raw = read_dataframe_from_file(file_path)
        if df_raw is None:
            send_status_update(file_path, "FAILURE", "读取文件时发生错误")
            status_counts["FAILURE"] += 1
            continue

        # 3.4 调用处理工具
        print("  -> 正在调用处理工具...", flush=True)
        
        # 使用从 identifier 返回的平台标识符 (如 "TM_RECENT") 作为键来查找处理工具。
        processor_func = PROCESSOR_MAP.get(platform)
        
        if not processor_func:
            print(f"  -> 错误：未找到平台 '{platform}' 对应的处理工具，跳过。", flush=True)
            send_status_update(file_path, "UNIDENTIFIED", f"未找到平台'{platform}'的处理工具")
            status_counts["UNIDENTIFIED"] += 1
            continue
            
        result_workbook = None
        try:
            result_workbook = processor_func(df_raw)
        except Exception as e:
            print(f"  -> 错误: 在处理【{platform}】数据时发生异常: {e}", flush=True)
            exc_str = traceback.format_exc()
            print(exc_str, flush=True)
            send_status_update(file_path, "FAILURE", f"处理时发生异常: {e}")
            status_counts["FAILURE"] += 1
            continue

        # 3.5 保存结果
        if result_workbook:
            print(f"  -> 正在保存到: '{os.path.basename(output_path)}'", flush=True)
            try:
                result_workbook.save(output_path)
                print("  -> 保存成功！", flush=True)
                send_status_update(file_path, "SUCCESS", f"已保存到: {output_path}")
                status_counts["SUCCESS"] += 1
            except Exception as e:
                print(f"  -> 错误：保存文件失败: {e}", flush=True)
                send_status_update(file_path, "FAILURE", f"保存文件时发生错误: {e}")
                status_counts["FAILURE"] += 1
        else:
            print("  -> 数据处理失败，未生成结果文件。", flush=True)
            send_status_update(file_path, "FAILURE", "处理工具未返回有效结果")
            status_counts["FAILURE"] += 1

    # 4. 结束总结
    end_time = datetime.now()
    duration = end_time - start_time
    print("\n" + "-" * 60, flush=True)
    print("所有任务处理完毕！", flush=True)
    print(f"处理总耗时: {duration}", flush=True)
    print("处理结果统计:", flush=True)
    print(f"  - 成功: {status_counts['SUCCESS']} 个文件", flush=True)
    print(f"  - 跳过 (文件已存在): {status_counts['SKIPPED']} 个文件", flush=True)
    print(f"  - 平台未识别: {status_counts['UNIDENTIFIED']} 个文件", flush=True)
    print(f"  - 失败 (发生错误): {status_counts['FAILURE']} 个文件", flush=True)
    print("-" * 60, flush=True)

if __name__ == "__main__":
    main()