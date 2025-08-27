# -*- coding: utf-8 -*-

"""
一个通用的、安全的CSV到XLSX转换工具原型。独立打包为exe后使用。

功能:
1. 通过拖拽方式接收一个或多个CSV文件进行批量转换。
2. 使用 chardet 库自动检测文件编码，极大提高对不同来源文件的兼容性。
3. 自动嗅探CSV文件的分隔符 (通常是逗号或分号)。
4. 智能分析每一列的数据内容，以防止在Excel中打开时发生常见的数据损坏问题：
    - 长数字（如订单号、身份证号）因超出15位精度而被截断或末尾变为0。
    - 纯数字字符串因过长而被Excel默认以不直观的科学计数法显示。
    - 文本型数字（如邮政编码、物料编码）的前导零被自动去除。
    - 看起来像日期的字符串（如'10-12'可能代表尺寸或编号）被错误地格式化为日期。
5. 检测并净化潜在的公式注入攻击。对于包含高风险函数（如HYPERLINK）或可疑命令
   （如DDE攻击）的单元格，会在其内容前添加一个单引号，使其在Excel中被强制
   解析为纯文本，从而阻止恶意代码执行，同时保留原始内容供用户审查。
6. 对于确实混合了数字和文本的列（例如包含数值和'无'、'N/A'等说明的列），为了保证
   100%不丢失任何原始信息，该列会整体作为文本处理。
7. 将处理后的数据保存为与原CSV文件同名（带前缀）、扩展名为 .xlsx 的新文件，并存放在
   与原文件相同的目录下。
"""

import pandas as pd
import sys
import os
import re
import csv
import numpy as np
import chardet

# 主动采纳Pandas未来的行为，以消除FutureWarning。
pd.set_option('future.no_silent_downcasting', True)

# --- 全局配置区 ---

# 预扫描的行数。程序会读取文件的前这么多行来分析数据结构，
SAMPLE_ROWS = 1000

# 数字长度阈值。当一列中检测到任何长度大于或等于此值的纯数字字符串时，
# 该列将被强制转换为文本格式，以同时解决Excel的15位精度丢失和12位科学计数法显示问题。
PRECISION_THRESHOLD = 12

# 危险公式关键字列表（检测时不区分大小写）。
# 这些关键字常用于调用外部程序、链接或服务，是公式注入攻击的常见特征。
RISKY_KEYWORDS = [
    'cmd', 'powershell', 'exec', '.exe', 'call', 'register.id',
    'urlmon', 'webservice', 'filterxml', 'hyperlink'
]

# 编译后的正则表达式，用于快速查找危险公式。
# 它的逻辑是：匹配一个以公式符号（=, +, -, @）开头，
# 并且后面包含了DDE攻击特征（|...!)或任何一个RISKY_KEYWORDS的字符串。
# 使用非捕获组 (?:...) 是为了优化性能并消除Pandas的UserWarning。
FORMULA_INJECTION_PATTERN = re.compile(
    r'^\s*[\=\+\-\@].*(?:\|.*!|' + '|'.join(RISKY_KEYWORDS) + ')',
    re.IGNORECASE
)

def detect_encoding_and_delimiter(file_path):
    """
    使用 chardet 和 csv.Sniffer 自动检测文件的编码和分隔符。
    
    Args:
        file_path (str): CSV文件的完整路径。
        
    Returns:
        tuple: (encoding, delimiter)，如果成功则返回检测结果，否则返回默认值。
    """
    encoding, delimiter = 'gbk', ',' # 默认回退值
    
    # 1. 使用 chardet 检测编码
    try:
        with open(file_path, 'rb') as f:
            raw_data = f.read(10000) # 读取文件头部一小部分字节用于检测
            if not raw_data: # 文件为空
                return None, None
            detection = chardet.detect(raw_data)
            encoding = detection['encoding']
            confidence = detection['confidence']
            print(f"  -> chardet检测到编码: {encoding} (置信度: {confidence:.0%})")
            # 对于常见的中文编码，如gb2312是gbk的子集，直接使用gbk更稳妥
            if encoding and 'gb' in encoding.lower():
                encoding = 'gbk'
    except Exception as e:
        print(f"  -> 警告: 编码检测失败: {e}。将使用默认编码。")

    # 2. 使用检测到的编码来嗅探分隔符
    try:
        with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
            sniffer = csv.Sniffer()
            delimiter = sniffer.sniff(f.read(2048)).delimiter
            print(f"  -> Sniffer检测到分隔符: '{delimiter}'")
    except Exception:
        print(f"  -> 警告: 无法自动检测分隔符，将默认使用逗号 ','。")

    return encoding, delimiter

def analyze_columns(file_path, encoding, delimiter):
    """
    通过预扫描文件来分析哪些列需要被强制指定为文本类型。
    这是实现“智能转换”的核心步骤，它为后续的完整读取制定规则。
    
    Args:
        file_path (str): CSV文件的完整路径。
        encoding (str): 文件编码。
        delimiter (str): CSV分隔符。
        
    Returns:
        dict: 一个适用于Pandas read_csv的dtype字典，例如 {'订单号': str}。
    """
    print("  -> (1/4) 正在预扫描文件以分析数据结构...")
    forced_text_cols = set()
    try:
        # 第一遍读取：只读样本行，且将所有列都当作字符串，以100%保留原始格式
        df_sample = pd.read_csv(
            file_path,
            encoding=encoding,
            sep=delimiter,
            dtype=str,
            nrows=SAMPLE_ROWS,
            keep_default_na=False,
            engine='python' # 使用更健壮的Python引擎以更好地处理复杂CSV
        )

        for col in df_sample.columns:
            # 创建一个没有空值和空白的临时Series用于分析
            series_non_null = df_sample[col].str.strip().replace('', np.nan).dropna()
            if series_non_null.empty:
                continue

            # 规则1: 前导零保护 (例如 '007', '012345')
            if series_non_null.str.match(r'^0[0-9]+$').any():
                forced_text_cols.add(col)
                continue

            # 规则2: 可疑日期格式保护 (例如 '10-12', '5-1')
            if series_non_null.str.match(r'^\d{1,2}-\d{1,2}$').any():
                forced_text_cols.add(col)
                continue
                
            # 规则3: 长数字保护 (解决精度丢失和科学计数法显示问题)
            numeric_strings = series_non_null[series_non_null.str.isdigit()]
            if not numeric_strings.empty and numeric_strings.str.len().max() >= PRECISION_THRESHOLD:
                forced_text_cols.add(col)
                continue

        if forced_text_cols:
            print("  -> 分析完成。以下列将被强制转换为文本以保留原始格式:")
            for col in sorted(list(forced_text_cols)):
                print(f"     - {col}")
        else:
            print("  -> 分析完成。未发现需要强制转换格式的列。")

        return {col: str for col in forced_text_cols}

    except Exception as e:
        print(f"  -> 预扫描失败: {e}。将尝试按默认方式读取。")
        return {}


def sanitize_dataframe(df):
    """
    遍历DataFrame，查找并净化所有潜在的公式注入单元格。
    
    Args:
        df (pd.DataFrame): 待处理的DataFrame。
        
    Returns:
        pd.DataFrame: 经过净化处理的DataFrame。
    """
    print("  -> (3/4) 正在扫描并净化潜在的恶意公式...")
    sanitized_count = 0
    
    # 只选择数据类型为'object'（通常是字符串）的列进行检查，以提高效率
    for col in df.select_dtypes(include=['object']).columns:
        # 使用.astype(str)确保所有内容都为字符串，然后用正则表达式进行匹配
        mask = df[col].astype(str).str.contains(FORMULA_INJECTION_PATTERN, na=False, regex=True)
        
        if mask.any():
            # 对所有匹配到的危险单元格，在其内容前添加一个单引号
            df.loc[mask, col] = "'" + df.loc[mask, col].astype(str)
            sanitized_count += mask.sum()
                
    if sanitized_count > 0:
        print(f"  -> 净化完成。共处理了 {sanitized_count} 个有风险的单元格。")
    else:
        print("  -> 扫描完成。未发现需要净化的风险单元格。")
        
    return df

def main():
    """主执行函数，处理命令行参数和文件转换流程。"""
    files_to_process = sys.argv[1:]

    if not files_to_process:
        print("csv到xlsx转换工具")
        print("用法: 请将一个或多个CSV文件拖拽到本程序的图标上。")
        input("\n按回车键退出...")
        return

    for file_path in files_to_process:
        print("-" * 60)
        print(f"开始处理文件: {os.path.basename(file_path)}")

        try:
            if not os.path.exists(file_path) or not file_path.lower().endswith('.csv'):
                print("  -> 错误: 文件不存在或不是一个CSV文件。")
                continue

            encoding, delimiter = detect_encoding_and_delimiter(file_path)
            if not encoding:
                print("  -> 错误: 文件为空或无法确定文件编码。")
                continue

            dtype_map = analyze_columns(file_path, encoding, delimiter)

            print("  -> (2/4) 正在读取完整文件...")
            # 第二遍读取：使用分析得出的规则，精确地读取完整文件
            df = pd.read_csv(
                file_path, encoding=encoding, sep=delimiter, dtype=dtype_map,
                keep_default_na=False, engine='python'
            )
            
            # 基础清洗：去除列名和所有字符串单元格的前后空格
            df.columns = [str(col).strip() for col in df.columns]
            for col in df.select_dtypes(include=['object']).columns:
                df[col] = df[col].str.strip()

            df = sanitize_dataframe(df)

            base_name = os.path.splitext(os.path.basename(file_path))[0]
            output_dir = os.path.dirname(file_path)
            output_path = os.path.join(output_dir, f"xlsx_{base_name}.xlsx")

            print(f"  -> (4/4) 正在写入Excel文件: {os.path.basename(output_path)}")
            df.to_excel(output_path, index=False)

            print(f"\n成功! 文件已保存至:\n{output_path}\n")

        except PermissionError:
            print(f"\n错误: 权限不足。文件 '{os.path.basename(file_path)}' 可能正被Excel或其他程序占用。")
            print("请关闭相关程序后重试。\n")
        except pd.errors.EmptyDataError:
            print(f"\n错误: 文件 '{os.path.basename(file_path)}' 为空或格式不正确，无法处理。\n")
        except Exception as e:
            print(f"\n发生未知错误: {e}\n")

    input("所有文件处理完毕。按回车键退出...")

if __name__ == '__main__':
    main()