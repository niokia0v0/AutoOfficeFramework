import pandas as pd
import openpyxl
import os
from typing import Dict, Any

#用于将源数据的敏感/隐私信息脱敏，以供公开演示
#翻译修正：desens->anonym

# ==============================================================================
#                                  配置区
# ==============================================================================

# 定义输入和输出文件夹路径
INPUT_FOLDER = r'C:\Users\LENOVO\Desktop'
OUTPUT_FOLDER = r'C:\Users\LENOVO\Desktop'

# 定义源文件名
INPUT_FILENAME = '发票记录格式.xlsx'

# 定义需要脱敏的、有明确列名的列及其脱敏后的前缀
# 格式: '原始列名': '脱敏前缀'
DESENSITIZATION_MAP = {
    '发票号': '发票号',
    '客户名称': '客户名称',
    '单号': '单号'
}

# ==============================================================================
#                                核心脱敏函数
# ==============================================================================

def desensitize_column(series: pd.Series, prefix: str, mapping_dict: Dict[str, str]) -> pd.Series:
    """
    对一个Pandas Series进行脱敏处理。
    
    - 相同的原始值会被映射到相同的脱敏结果。
    - 空白或占位符值保持不变。
    
    Args:
        series (pd.Series): 需要脱敏的列数据。
        prefix (str): 脱敏后生成字符串的前缀 (如 "客户")。
        mapping_dict (Dict[str, str]): 用于存储映射关系的字典，实现跨行唯一性。
        
    Returns:
        pd.Series: 脱敏处理后的新Series。
    """
    
    # 从当前映射字典的大小开始计数，确保ID是连续的
    next_id = len(mapping_dict) + 1
    
    def apply_desensitization(value: Any) -> Any:
        nonlocal next_id
        # 检查是否为空值、空字符串或常见占位符，如果是则保持原样
        if pd.isna(value) or str(value).strip() in ['', '-', '--', 'nan', 'None']:
            return value
        
        value_str = str(value)
        # 如果当前值尚未被映射
        if value_str not in mapping_dict:
            # 创建新的脱敏ID并存入映射
            mapping_dict[value_str] = f"{prefix}{next_id}"
            next_id += 1
            
        return mapping_dict[value_str]

    return series.apply(apply_desensitization)

# ==============================================================================
#                                  主执行逻辑
# ==============================================================================

def main():
    """
    主函数，执行发票文件的读取、脱敏和保存操作，并完整保留原始格式。
    """
    # 构造输入文件的完整路径
    input_path = os.path.join(INPUT_FOLDER, INPUT_FILENAME)

    # 生成输出文件名，在原始文件名前添加 "anonym_" 前缀
    output_filename = f"anonym_{INPUT_FILENAME}"
    # 构造输出文件的完整路径
    output_path = os.path.join(OUTPUT_FOLDER, output_filename)

    print(f"开始处理文件: {input_path}")

    # --- 步骤 1: 使用 Pandas 读取数据内容 ---
    # 这一步仅用于方便地对数据进行批量脱敏处理，会忽略所有格式信息。
    # 使用 dtype=str 确保发票号、单号等不会被错误地识别为数字。
    try:
        df = pd.read_excel(input_path, header=0, dtype=str)
    except FileNotFoundError:
        print(f"  -> 错误: 输入文件未找到 -> {input_path}")
        return
    except Exception as e:
        print(f"  -> 错误: 使用Pandas读取文件失败: {e}")
        return
        
    print("  -> 数据内容读取成功，开始进行脱敏处理...")

    # --- 步骤 2: 在内存中对 DataFrame 进行脱敏 ---
    
    # 缓存，用于存储每个脱敏列的映射关系
    # 格式: {'原始列名': {'原始值1': '脱敏值1', ...}, ...}
    mapping_cache: Dict[str, Dict[str, str]] = {}

    # a. 处理在 DESENSITIZATION_MAP 中定义的带名称的列
    for col_name, prefix in DESENSITIZATION_MAP.items():
        if col_name in df.columns:
            print(f"    -> 正在脱敏列: '{col_name}'")
            # 为该列获取或创建映射字典
            col_mapping = mapping_cache.setdefault(col_name, {})
            df[col_name] = desensitize_column(df[col_name], prefix, col_mapping)
        else:
            print(f"    -> 警告: 未在文件中找到列 '{col_name}'，跳过脱敏。")

    # b. 特殊处理第17和18列 (地址和联系方式)，它们没有固定列名
    # 通过列的整数位置来定位，这样更健壮
    if len(df.columns) >= 18:
        # 第17列 (Pandas索引为16)
        col_17_name = df.columns[16]
        print(f"    -> 正在脱敏第17列 (地址): '{col_17_name}'")
        col_17_mapping = mapping_cache.setdefault('address_col', {})
        df[col_17_name] = desensitize_column(df[col_17_name], '地址', col_17_mapping)

        # 第18列 (Pandas索引为17)
        col_18_name = df.columns[17]
        print(f"    -> 正在脱敏第18列 (联系方式): '{col_18_name}'")
        col_18_mapping = mapping_cache.setdefault('contact_col', {})
        df[col_18_name] = desensitize_column(df[col_18_name], '联系方式', col_18_mapping)
    else:
        print("    -> 警告: 文件列数不足18列，无法脱敏地址和联系方式列。")

    print("  -> 数据脱敏处理完成。")
    print("  -> 开始将脱敏数据写回并保留原始格式...")
    
    # --- 步骤 3: 使用 Openpyxl 写入数据以保留格式 ---
    # 这是一个关键步骤。我们加载原始工作簿（包含所有格式），
    # 然后仅用脱敏后的数据更新每个单元格的 .value 属性。
    try:
        # 加载原始工作簿，这将保留所有单元格的颜色、字体、边框等格式
        workbook = openpyxl.load_workbook(input_path)
        sheet = workbook.active
        
        # 遍历脱敏后的 DataFrame，并将值逐一写回 openpyxl 工作表对象
        # header=True，所以数据从第2行开始
        # iterrows() 的索引是0-based，Excel行号是1-based，所以行号是 `idx + 2`
        for idx, row in df.iterrows():
            # 列号是1-based，所以是 `col_idx + 1`
            for col_idx, col_name in enumerate(df.columns):
                cell_to_update = sheet.cell(row=idx + 2, column=col_idx + 1)
                # 只更新值，不改变任何其他属性
                cell_to_update.value = row[col_name]

        # 保存被修改过的工作簿到新文件
        workbook.save(output_path)
        print(f"  -> 保存成功！脱敏后的文件已保存至: {output_path}")

    except PermissionError:
        # 更新错误提示，使用动态生成的 `output_filename` 变量
        print(f"  -> 错误: 保存文件失败。请确保文件 '{output_filename}' 未被其他程序打开。")
    except Exception as e:
        print(f"  -> 错误: 使用Openpyxl写入文件时发生未知错误: {e}")

if __name__ == "__main__":
    main()
    print("-" * 50)
    print("所有操作已完成。")