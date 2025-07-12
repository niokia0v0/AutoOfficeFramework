import pandas as pd
import os
import re
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
import numpy as np

# --- 配置区 ---

# Pandas 显示选项
pd.set_option('future.no_silent_downcasting', True)

# 京东原始列名映射
JD_COL_ORDER_ID = '订单编号'
JD_COL_ORDER_STATUS = '订单状态'
JD_COL_PRODUCT_ID = '商品编号'
JD_COL_PRODUCT_NAME = '商品名称'
JD_COL_QUANTITY = '商品数量'
JD_COL_AMOUNT_DUE = '应结金额'
JD_COL_FEE_NAME = '费用名称'
JD_COL_DIRECTION = '收支方向'
JD_COL_AFTER_SALES_ID = '售后服务单号'
JD_COL_COMMISSION_RATIO = '佣金比例'

# 费用名称常量
FEE_NAME_GOODS = '货款'
FEE_NAME_COMMISSION = '佣金'
FEE_NAME_TRANSACTION = '交易服务费'
FEE_NAME_AD_COMMISSION = '广告联合活动降扣佣金'
FEE_NAME_JINGDOU = '京豆'
FEE_NAME_PRODUCT_INSURANCE = '商品保险服务费'
FEE_NAME_FREIGHT_INSURANCE = '运费保险服务费'

# 订单状态常量
STATUS_COMPLETED = '已完成'

# 收支方向常量
DIRECTION_INCOME = '收入'
DIRECTION_EXPENSE = '支出'

# 输出到详情页的列定义
DETAIL_SHEET_COLUMNS_JD = [
    '订单编号', '父单号', '订单状态', '订单下单时间', '订单完成时间', '售后服务单号', '售后退款时间',
    '商品编号', '商品名称', '商品数量', '扣点类型', '佣金比例', '费用名称', '应结金额',
    '收支方向', '结算状态', '预计结算时间', '账单生成时间', '到账时间', '商户订单号',
    '资金动账备注', '费用项含义', '备注', '留用时间', '费用说明'
]

# --- 内部功能函数 ---

def _convert_numeric_columns(df):
    """将DataFrame中的指定列转换为数值类型。"""
    df[JD_COL_AMOUNT_DUE] = pd.to_numeric(df[JD_COL_AMOUNT_DUE], errors='coerce').fillna(0)
    df[JD_COL_QUANTITY] = pd.to_numeric(df[JD_COL_QUANTITY], errors='coerce').fillna(0)
    if JD_COL_COMMISSION_RATIO in df.columns:
        df[JD_COL_COMMISSION_RATIO] = pd.to_numeric(df[JD_COL_COMMISSION_RATIO], errors='coerce')
    return df

def _filter_and_prepare_data(df_numeric):
    """筛选有效数据（已完成状态，有商品ID）。"""
    df_completed = df_numeric[df_numeric[JD_COL_ORDER_STATUS] == STATUS_COMPLETED].copy()
    
    # 确保售后相关列存在，以便后续筛选
    for col in [JD_COL_AFTER_SALES_ID, '售后退款时间']:
        if col not in df_completed.columns:
            df_completed[col] = np.nan
            
    df_processed = df_completed[df_completed[JD_COL_PRODUCT_ID].notna()].copy()
    
    if df_processed.empty:
        print(f"数据中没有找到“已完成”且包含有效商品ID('{JD_COL_PRODUCT_ID}')的行。无法生成报告。")
        return None
        
    return df_processed

def _aggregate_product_data(df_processed, original_df_numeric):
    """按商品ID聚合数据，计算各项收入与支出。"""
    product_summary = {}
    for product_id, group in df_processed.groupby(JD_COL_PRODUCT_ID):
        product_name = group[JD_COL_PRODUCT_NAME].dropna().iloc[0] if not group[JD_COL_PRODUCT_NAME].dropna().empty else "未知商品"

        sales_group = group[(group[JD_COL_FEE_NAME] == FEE_NAME_GOODS) & (group[JD_COL_DIRECTION] == DIRECTION_INCOME)]
        returns_group = group[
            (group[JD_COL_FEE_NAME] == FEE_NAME_GOODS) & 
            (group[JD_COL_DIRECTION] == DIRECTION_EXPENSE) & 
            (group[JD_COL_AFTER_SALES_ID].notna()) & 
            (group[JD_COL_AFTER_SALES_ID] != '')
        ]
        
        expenses_group = group[group[JD_COL_DIRECTION] == DIRECTION_EXPENSE]
        commission = expenses_group[expenses_group[JD_COL_FEE_NAME] == FEE_NAME_COMMISSION][JD_COL_AMOUNT_DUE].sum()
        transaction_fee = expenses_group[expenses_group[JD_COL_FEE_NAME] == FEE_NAME_TRANSACTION][JD_COL_AMOUNT_DUE].sum()
        ad_commission = expenses_group[expenses_group[JD_COL_FEE_NAME] == FEE_NAME_AD_COMMISSION][JD_COL_AMOUNT_DUE].sum()
        jingdou = expenses_group[expenses_group[JD_COL_FEE_NAME] == FEE_NAME_JINGDOU][JD_COL_AMOUNT_DUE].sum()

        related_order_ids = group[JD_COL_ORDER_ID].unique()
        orders_df = original_df_numeric[original_df_numeric[JD_COL_ORDER_ID].isin(related_order_ids)]
        product_insurance = orders_df[(orders_df[JD_COL_FEE_NAME] == FEE_NAME_PRODUCT_INSURANCE) & (orders_df[JD_COL_DIRECTION] == DIRECTION_EXPENSE)][JD_COL_AMOUNT_DUE].sum()
        freight_insurance = orders_df[(orders_df[JD_COL_FEE_NAME] == FEE_NAME_FREIGHT_INSURANCE) & (orders_df[JD_COL_DIRECTION] == DIRECTION_EXPENSE)][JD_COL_AMOUNT_DUE].sum()
        
        product_summary[str(product_id)] = {
            'name': product_name,
            'sales_quantity': sales_group[JD_COL_QUANTITY].sum(),
            'sales_amount': sales_group[JD_COL_AMOUNT_DUE].sum(),
            'return_quantity': returns_group[JD_COL_QUANTITY].sum(),
            'return_amount': returns_group[JD_COL_AMOUNT_DUE].sum(),
            'commission': commission,
            'transaction_fee': transaction_fee,
            'ad_commission': ad_commission,
            'jingdou': jingdou,
            'product_insurance': product_insurance,
            'freight_insurance': freight_insurance,
            'total_product_expenses': commission + transaction_fee + ad_commission + jingdou + product_insurance + freight_insurance,
            'sales_detail_df': sales_group,
            'returns_detail_df': returns_group,
        }
    return product_summary

def _create_summary_sheet(wb, product_summary):
    """在工作簿中创建并填充销售总结页，精确匹配原型逻辑。"""
    ws = wb.active
    ws.title = "销售总结"
    bold_font = Font(bold=True)
    
    headers = [
        "商品编号", "商品名称", "销售数量", "销售额", "佣金支出", "交易服务费支出",
        "广告降扣支出", "京豆支出", "商品保险费支出", "运费保险费支出", "产品总支出"
    ]
    ws.append(headers)
    for cell in ws[1]: cell.font = bold_font
    
    sorted_ids = sorted(product_summary.keys())
    grand_total_return_qty, grand_total_return_amt = 0, 0.0
    for prod_id in sorted_ids:
        item = product_summary[prod_id]
        ws.append([
            prod_id, item['name'], item['sales_quantity'], item['sales_amount'],
            item['commission'], item['transaction_fee'], item['ad_commission'], item['jingdou'],
            item['product_insurance'], item['freight_insurance'], item['total_product_expenses']
        ])
        grand_total_return_qty += item['return_quantity']
        grand_total_return_amt += item['return_amount']

    max_data_row = ws.max_row
    current_row = max_data_row + 2
    ws.cell(row=current_row, column=1, value="总计 (不含退款)").font = bold_font
    for i, _ in enumerate(headers[2:], 3):
        col_letter = get_column_letter(i)
        ws.cell(row=current_row, column=i, value=f"=SUM({col_letter}2:{col_letter}{max_data_row})").font = bold_font

    current_row += 2
    ws.cell(row=current_row, column=1, value="退款商品明细").font = bold_font
    current_row += 1
    
    refund_start_row = current_row
    has_refunds = False
    for prod_id in sorted_ids:
        item = product_summary[prod_id]
        if item['return_quantity'] > 0:
            has_refunds = True
            ws.append([prod_id, item['name'], item['return_quantity'], item['return_amount']])
            current_row += 1
    refund_end_row = current_row - 1
    
    ws.cell(row=current_row, column=1, value="总计退款").font = bold_font
    if has_refunds:
        ws.cell(row=current_row, column=3, value=f"=SUM(C{refund_start_row}:C{refund_end_row})").font = bold_font
        ws.cell(row=current_row, column=4, value=f"=SUM(D{refund_start_row}:D{refund_end_row})").font = bold_font
    else:
        ws.cell(row=current_row, column=3, value=0).font = bold_font
        ws.cell(row=current_row, column=4, value=0).font = bold_font
        
    current_row += 2
    ws.cell(row=current_row, column=1, value="总计 (计算退款)").font = bold_font
    ws.cell(row=current_row, column=3, value=f"={get_column_letter(3)}{max_data_row+2} - {abs(grand_total_return_qty)}").font = bold_font
    ws.cell(row=current_row, column=4, value=f"={get_column_letter(4)}{max_data_row+2} + {grand_total_return_amt}").font = bold_font
    
    ws.column_dimensions['A'].width = 20; ws.column_dimensions['B'].width = 70
    for col in "CDEFGHIJK": ws.column_dimensions[col].width = 18
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, (int, float)) or (isinstance(cell.value, str) and cell.value.startswith("=")):
                if cell.column == 3: cell.number_format = '#,##0'
                else: cell.number_format = '#,##0.00'

def _create_detail_sheets(wb, product_summary):
    """为每个商品创建并填充详情页。"""
    for prod_id, item in product_summary.items():
        if item['sales_detail_df'].empty and item['returns_detail_df'].empty: continue

        sheet_name_raw = f"{prod_id}_{item['name']}"
        sheet_name = re.sub(r'[\\/\*\[\]\:?]', '_', sheet_name_raw)[:31]
        try: ws = wb.create_sheet(sheet_name)
        except: ws = wb.create_sheet(f"{prod_id}_detail")
            
        header_written = False
        def write_df_section(df, title, qty, amt):
            nonlocal header_written
            if df.empty: return
            ws.append([title]); ws[ws.max_row][0].font = Font(bold=True)
            if not header_written:
                ws.append(DETAIL_SHEET_COLUMNS_JD); header_written = True
                for cell in ws[ws.max_row]: cell.font = Font(bold=True)
            
            for r in df.reindex(columns=DETAIL_SHEET_COLUMNS_JD).fillna('').itertuples(index=False):
                ws.append(list(r))
            
            total_row = ["总计", *[""]*8, qty, *[""]*3, amt, *[""]*11]
            ws.append(total_row)
            for cell in ws[ws.max_row]: cell.font = Font(bold=True)
            ws.append([])
        
        write_df_section(item['sales_detail_df'], "销售明细", item['sales_quantity'], item['sales_amount'])
        write_df_section(item['returns_detail_df'], "退款明细", item['return_quantity'], item['return_amount'])

# --- 主处理函数 (新架构) ---

def process_jingdong_data(df_raw):
    """
    处理京东结算DataFrame，生成包含总结和明细的Excel Workbook对象。
    """
    if df_raw is None or df_raw.empty:
        print("京东处理错误：输入的DataFrame为空。")
        return None
        
    df_numeric = _convert_numeric_columns(df_raw.copy())
    df_processed = _filter_and_prepare_data(df_numeric.copy())
    if df_processed is None: return None
        
    product_summary = _aggregate_product_data(df_processed, df_numeric)
    
    wb = Workbook()
    _create_summary_sheet(wb, product_summary)
    _create_detail_sheets(wb, product_summary)
    
    if "销售总结" in wb.sheetnames and wb.sheetnames[0] != "销售总结":
        wb.move_sheet("销售总结", -len(wb.sheetnames)+1)
        
    return wb

# ---- 主程序入口 (用于独立测试) ----
if __name__ == "__main__":
    TEST_INPUT_DIR = r"C:\Users\LENOVO\Desktop"
    TEST_OUTPUT_DIR = r"C:\Users\LENOVO\Desktop"
    TEST_FILENAME = "订单结算明细对账_2025-05-01_2025-05-31 (1).csv"

    input_file = os.path.join(TEST_INPUT_DIR, TEST_FILENAME)

    if not os.path.exists(input_file):
        print(f"错误: 测试输入文件未找到于 '{os.path.abspath(input_file)}'")
    else:
        print(f"--- 独立测试: 读取文件 {TEST_FILENAME} ---")
        try:
            df_test_raw = pd.read_csv(input_file, dtype=str, na_values=['--'], keep_default_na=True, encoding='utf-8-sig')
            df_test_raw.columns = [col.strip() for col in df_test_raw.columns]
            for col in df_test_raw.columns:
                if df_test_raw[col].dtype == 'object':
                    df_test_raw[col] = df_test_raw[col].str.strip().replace(['--', '', 'None', 'nan', np.nan], np.nan, regex=False)
        except Exception as e:
            print(f"测试中读取文件失败: {e}")
            df_test_raw = None

        if df_test_raw is not None:
            print("--- 独立测试: 调用 process_jingdong_data ---")
            workbook_result = process_jingdong_data(df_test_raw)
            
            if workbook_result:
                os.makedirs(TEST_OUTPUT_DIR, exist_ok=True)
                file_name_without_ext = os.path.splitext(TEST_FILENAME)[0]
                output_filename = f"JD_output_{file_name_without_ext}.xlsx"
                output_file_path = os.path.join(TEST_OUTPUT_DIR, output_filename)
                
                try:
                    workbook_result.save(output_file_path)
                    print(f"\n脚本独立测试成功。输出文件位于: {output_file_path}")
                except Exception as e:
                    print(f"\n测试中保存文件失败: {e}")
            else:
                print("\n脚本独立测试失败，未能生成Workbook对象。")