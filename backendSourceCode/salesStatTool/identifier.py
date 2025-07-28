import pandas as pd
import os

# --- 平台指纹定义 ---
# 定义每个平台独有的、稳定存在的列名集合作为“指纹”

PLATFORM_FINGERPRINTS = {
    "TM_RECENT": {  # 天猫/淘宝 (三个月内 .xlsx)
        '子订单编号',
        '主订单编号',
        '商品标题',
        '买家实付金额',
        '退款金额',
        '商品ID'
    },
    "TM_HISTORY": { # 天猫/淘宝 (历史数据 .csv)
        '主订单编号',
        '子订单编号',
        '标题',
        '商家编码',
        '退款金额',
        '订单状态'
    },
    "JD": {  # 京东
        '订单编号',
        '订单下单时间',
        '费用名称',
        '应结金额',
        '收支方向',
        '结算状态'
    },
    "PDD": { # 拼多多
        '商品',
        '订单号',
        '商品总价(元)',
        '商家实收金额(元)',
        '商品id',
        '售后状态'
    },
    "DY": {  # 抖店
        '主订单编号',
        '选购商品',
        '商品金额',
        '订单提交时间',
        '订单完成时间',
        '售后状态'
    }
}

# --- 核心识别函数 ---

def identify_platform(file_path):
    """
    通过读取文件表头并与预定义的指纹比对，来识别文件所属的电商平台。

    Args:
        file_path (str): 需要识别的文件的完整路径 (.csv, .xls, .xlsx)。

    Returns:
        str: 代表平台的字符串 (e.g., "TM_RECENT", "JD", "TM_HISTORY")。
             如果无法识别或文件有问题，则返回 None。
    """
    if not os.path.exists(file_path):
        print(f"识别错误: 文件不存在 -> {file_path}")
        return None

    try:
        # 根据文件扩展名选择合适的读取方式，只读取表头 (nrows=0)
        file_ext = os.path.splitext(file_path)[1].lower()
        
        df_header = None
        if file_ext == '.csv':
            try:
                # 优先尝试用 utf-8-sig 读取，它能处理带BOM的CSV
                df_header = pd.read_csv(file_path, nrows=0, encoding='utf-8-sig', keep_default_na=False)
            except UnicodeDecodeError:
                # 如果UTF-8解码失败，则回退到GBK编码再次尝试
                print(f"  -> UTF-8解码失败，尝试使用GBK编码读取表头...")
                df_header = pd.read_csv(file_path, nrows=0, encoding='gbk', keep_default_na=False)

        elif file_ext in ['.xlsx', '.xls']:
            df_header = pd.read_excel(file_path, nrows=0, engine='openpyxl')
        else:
            print(f"识别警告: 不支持的文件类型 -> {os.path.basename(file_path)}")
            return None
            
        # 清理列名中的空格和潜在的引号
        header_columns = {col.strip().replace('"', '') for col in df_header.columns}

        # 逐一比对指纹
        for platform, fingerprint in PLATFORM_FINGERPRINTS.items():
            # issubset() 检查指纹中的所有列名是否都存在于文件的表头中
            if fingerprint.issubset(header_columns):
                return platform
        
        # 如果所有指纹都未匹配
        return None

    except pd.errors.EmptyDataError:
        print(f"识别警告: 文件为空 -> {os.path.basename(file_path)}")
        return None
    except Exception as e:
        print(f"识别错误: 读取文件 '{os.path.basename(file_path)}' 表头时发生错误: {e}")
        return None

# ---- 主程序入口 (用于独立测试) ----
if __name__ == "__main__":
    # --- 测试配置 ---
    # 将这里的路径修改为存放所有待测文件的目录
    TEST_DIRECTORY = r"C:\Users\LENOVO\Desktop"
    # --- 测试配置结束 ---
    
    print(f"--- 开始测试平台识别模块 ---")
    print(f"测试目录: {os.path.abspath(TEST_DIRECTORY)}\n")

    if not os.path.isdir(TEST_DIRECTORY):
        print(f"错误: 指定的测试目录不存在 -> {TEST_DIRECTORY}")
    else:
        # 使用 os.walk 遍历目录及其所有子目录
        for root, _, files in os.walk(TEST_DIRECTORY):
            for filename in files:
                # 只处理指定类型的文件
                if filename.lower().endswith(('.csv', '.xlsx', '.xls')):
                    file_full_path = os.path.join(root, filename)
                    
                    # 获取相对于测试根目录的相对路径，使输出更简洁
                    relative_path = os.path.relpath(file_full_path, TEST_DIRECTORY)
                    
                    print(f"正在分析文件: {relative_path}")
                    platform_result = identify_platform(file_full_path)
                    
                    if platform_result:
                        print(f"  -> 识别结果: 【{platform_result}】\n")
                    else:
                        print(f"  -> 识别结果: 【未能识别】\n")
                        
    print("--- 测试结束 ---")