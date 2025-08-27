import os
import pandas as pd
from typing import Dict, Any

#用于将源数据的敏感/隐私信息脱敏，以供公开演示

# ==============================================================================
#                                  配置区
# ==============================================================================

# 存放原始数据文件的文件夹
INPUT_DIR = r"C:\Users\LENOVO\Desktop\my_input_data"

# 保存脱敏后文件的文件夹
OUTPUT_DIR = r"C:\Users\LENOVO\Desktop\my_output_data"

# 定义需要脱敏的列名及其脱敏后的前缀
# 格式: '原始列名': '脱敏前缀'
DESENSITIZATION_MAP = {
    # --- 天猫历史 & 天猫近期 ---
    '子订单编号': '子订单编号',
    '主订单编号': '主订单编号',
    '标题': '商品标题',          # 兼容天猫历史
    '商品标题': '商品标题',      # 兼容天猫近期
    '支付单号': '支付单号',
    '商品ID': '商品ID',
    '物流单号': '物流单号',

    # --- 京东 ---
    '订单编号': '订单编号',
    '父单号': '父单号',
    '售后服务单号': '售后服务单号',
    '商品编号': '商品编号',
    '商品名称': '商品名称',
    '商户订单号': '商户订单号',

    # --- 拼多多 ---
    '商品': '商品',
    '订单号': '订单号',
    '商品id': '商品ID',
    '样式ID': '样式ID',
    '快递单号': '快递单号',

    # --- 抖音 ---
    # '主订单编号' 和 '商品ID' 已在上面定义
    '选购商品': '选购商品'
    # '售后状态' 有特殊处理逻辑，不在此处定义
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
        prefix (str): 脱敏后生成字符串的前缀 (如 "订单编号")。
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
    主函数，遍历输入目录，执行脱敏并保存。
    """
    if not os.path.isdir(INPUT_DIR):
        print(f"错误: 输入文件夹不存在 -> {INPUT_DIR}")
        return

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    print(f"输入文件夹: {os.path.abspath(INPUT_DIR)}")
    print(f"输出文件夹: {os.path.abspath(OUTPUT_DIR)}")
    print("-" * 50)

    # 缓存，用于存储每个脱敏列的映射关系
    # 格式: {'原始列名': {'原始值1': '脱敏值1', ...}, ...}
    mapping_cache: Dict[str, Dict[str, str]] = {}

    for filename in sorted(os.listdir(INPUT_DIR)):
        input_path = os.path.join(INPUT_DIR, filename)
        file_ext = os.path.splitext(filename)[1].lower()

        if file_ext not in ['.csv', '.xlsx']:
            continue
            
        print(f"正在处理: {filename}...")

        # --- 1. 读取文件 ---
        df = None
        try:
            if file_ext == '.csv':
                try:
                    df = pd.read_csv(input_path, dtype=str, keep_default_na=False, encoding='utf-8-sig')
                except UnicodeDecodeError:
                    print("  -> UTF-8 解码失败，尝试使用 GBK 编码...")
                    df = pd.read_csv(input_path, dtype=str, keep_default_na=False, encoding='gbk')
            elif file_ext == '.xlsx':
                df = pd.read_excel(input_path, dtype=str, engine='openpyxl', keep_default_na=False)
        except Exception as e:
            print(f"  -> 错误: 读取文件失败: {e}")
            continue

        if df is None:
            continue
            
        # --- 2. 循环处理需要脱敏的列 ---
        for col_name, prefix in DESENSITIZATION_MAP.items():
            if col_name in df.columns:
                print(f"  -> 正在脱敏列: '{col_name}'")
                # 为该列获取或创建映射字典
                col_mapping = mapping_cache.setdefault(col_name, {})
                df[col_name] = desensitize_column(df[col_name], prefix, col_mapping)
        
        # --- 3. 处理特殊规则 (抖音售后状态) ---
        dy_special_col = '售后状态'
        if dy_special_col in df.columns:
            print(f"  -> 正在处理特殊规则列: '{dy_special_col}'")
            def process_dy_status(status):
                if isinstance(status, str) and '-' in status:
                    # 分割字符串并保留分隔符及之后的部分
                    return '-' + status.split('-', 1)[1]
                return status
            df[dy_special_col] = df[dy_special_col].apply(process_dy_status)

        # --- 4. 保存文件 ---
        output_filename = f"desens_{filename}"
        output_path = os.path.join(OUTPUT_DIR, output_filename)
        try:
            if file_ext == '.csv':
                # 保存为 utf-8-sig 编码，确保 Excel 能正确打开
                df.to_csv(output_path, index=False, encoding='utf-8-sig')
            elif file_ext == '.xlsx':
                df.to_excel(output_path, index=False, engine='openpyxl')
            print(f"  -> 保存成功: {output_filename}\n")
        except Exception as e:
            print(f"  -> 错误: 保存文件失败: {e}\n")

    print("-" * 50)
    print("所有文件处理完毕！")

if __name__ == "__main__":
    main()