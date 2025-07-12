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

# 拼多多原始列名映射
PDD_COL_PRODUCT_NAME = '商品'
PDD_COL_ORDER_ID = '订单号'
PDD_COL_ORDER_STATUS = '订单状态'
PDD_COL_PRODUCT_TOTAL_PRICE = '商品总价(元)'
PDD_COL_STORE_DISCOUNT = '店铺优惠折扣(元)'
PDD_COL_PLATFORM_DISCOUNT = '平台优惠折扣(元)'
PDD_COL_USER_ACTUAL_PAYMENT = '用户实付金额(元)'
PDD_COL_MERCHANT_ACTUAL_RECEIPT = '商家实收金额(元)'
PDD_COL_QUANTITY = '商品数量(件)'
PDD_COL_SHIPPING_TIME = '发货时间'
PDD_COL_CONFIRM_RECEIPT_TIME = '确认收货时间'
PDD_COL_PRODUCT_ID = '商品id'
PDD_COL_PRODUCT_SPEC = '商品规格'
PDD_COL_AFTER_SALES_STATUS = '售后状态'
PDD_COL_LOGISTICS_NO = '快递单号'
PDD_COL_LOGISTICS_COMPANY = '快递公司'
PDD_COL_ORDER_TRANSACTION_TIME = '订单成交时间'

# 订单状态常量
STATUS_REFUND_SUCCESS = '退款成功'

# 输出到详情页的列定义
DETAIL_SHEET_COLUMNS_PDD = [
    '订单号', '订单状态', '售后状态', '商品ID', '商品名称', '商品规格', '商品数量(件)',
    '用户实付金额(元)', '商家实收金额(元)', '订单成交时间', '发货时间', '确认收货时间',
    '快递单号', '快递公司'
]

# --- 内部功能函数 ---

def _prepare_and_validate_data(df):
    """验证数据，转换数值列，并进行数据筛选。"""
    numeric_cols = [
        PDD_COL_QUANTITY, PDD_COL_USER_ACTUAL_PAYMENT, PDD_COL_MERCHANT_ACTUAL_RECEIPT,
        PDD_COL_PRODUCT_TOTAL_PRICE, PDD_COL_STORE_DISCOUNT, PDD_COL_PLATFORM_DISCOUNT,
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    df_no_cancel = df[~df[PDD_COL_ORDER_STATUS].str.contains('取消', na=False)].copy()
    df_processed = df_no_cancel[df_no_cancel[PDD_COL_PRODUCT_ID].notna()].copy()
    
    if df_processed.empty:
        print("数据中没有找到未取消且包含有效商品ID的行。无法生成报告。")
        return None
        
    df_processed[PDD_COL_PRODUCT_ID] = df_processed[PDD_COL_PRODUCT_ID].astype(str)
    return df_processed

def _format_df_for_detail(df_source, p_id, p_name, is_refund=False):
    """根据详情页列定义，格式化DataFrame并计算金额。"""
    if df_source.empty: return pd.DataFrame(columns=DETAIL_SHEET_COLUMNS_PDD)
    
    df_target = pd.DataFrame()
    df_target['订单号'] = df_source.get(PDD_COL_ORDER_ID)
    df_target['订单状态'] = df_source.get(PDD_COL_ORDER_STATUS)
    df_target['售后状态'] = df_source.get(PDD_COL_AFTER_SALES_STATUS)
    df_target['商品ID'] = p_id
    df_target['商品名称'] = p_name
    df_target['商品规格'] = df_source.get(PDD_COL_PRODUCT_SPEC)
    df_target['商品数量(件)'] = df_source.get(PDD_COL_QUANTITY)
    df_target['订单成交时间'] = df_source.get(PDD_COL_ORDER_TRANSACTION_TIME)
    df_target['发货时间'] = df_source.get(PDD_COL_SHIPPING_TIME)
    df_target['确认收货时间'] = df_source.get(PDD_COL_CONFIRM_RECEIPT_TIME)
    df_target['快递单号'] = df_source.get(PDD_COL_LOGISTICS_NO)
    df_target['快递公司'] = df_source.get(PDD_COL_LOGISTICS_COMPANY)

    if not is_refund:
        df_target['用户实付金额(元)'] = df_source.get(PDD_COL_USER_ACTUAL_PAYMENT, 0)
        df_target['商家实收金额(元)'] = df_source.get(PDD_COL_PRODUCT_TOTAL_PRICE, 0) - df_source.get(PDD_COL_STORE_DISCOUNT, 0)
    else:
        df_target['用户实付金额(元)'] = -df_source.get(PDD_COL_USER_ACTUAL_PAYMENT, 0)
        df_target['商家实收金额(元)'] = -(df_source.get(PDD_COL_USER_ACTUAL_PAYMENT, 0) + df_source.get(PDD_COL_PLATFORM_DISCOUNT, 0))

    return df_target.reindex(columns=DETAIL_SHEET_COLUMNS_PDD).fillna('')

def _aggregate_product_data(df_processed):
    """按商品ID聚合数据，划分为收入和两类退款。"""
    product_data_map = {}
    for product_id, group_df in df_processed.groupby(PDD_COL_PRODUCT_ID):
        product_name = group_df[PDD_COL_PRODUCT_NAME].iloc[0] if not group_df[PDD_COL_PRODUCT_NAME].empty else "未知商品"

        all_refunds_df = group_df[group_df[PDD_COL_AFTER_SALES_STATUS].str.contains(STATUS_REFUND_SUCCESS, na=False)]
        unshipped_refund_df = all_refunds_df[all_refunds_df[PDD_COL_ORDER_STATUS].str.contains('未发货', na=False)]
        shipped_refund_df = all_refunds_df[all_refunds_df[PDD_COL_ORDER_STATUS].str.contains('已发货', na=False)]

        product_data_map[product_id] = {
            'name': product_name,
            'income_df': _format_df_for_detail(group_df, product_id, product_name),
            'unshipped_refund_df': _format_df_for_detail(unshipped_refund_df, product_id, product_name, is_refund=True),
            'shipped_refund_df': _format_df_for_detail(shipped_refund_df, product_id, product_name, is_refund=True)
        }
    return product_data_map

def _create_summary_sheet(wb, product_data_map):
    """在工作簿中创建并填充销售总结页。"""
    ws = wb.active
    ws.title = "销售总结"
    bold_font = Font(bold=True)
    center_align = Alignment(horizontal='center', vertical='center')
    current_row = 1

    def write_summary_section(title, df_key, headers):
        nonlocal current_row
        ws.cell(row=current_row, column=1, value=title).font = bold_font
        current_row += 1
        ws.append(headers)
        for cell in ws[current_row]: cell.font = bold_font; cell.alignment = center_align
        current_row += 1
        
        total_qty, total_user_pay, total_receipt = 0, 0.0, 0.0
        for prod_id in sorted(product_data_map.keys()):
            item = product_data_map[prod_id]
            df = item[df_key]
            if not df.empty:
                qty, user_pay, receipt = df['商品数量(件)'].sum(), df['用户实付金额(元)'].sum(), df['商家实收金额(元)'].sum()
                ws.append([prod_id, item['name'], qty, user_pay, receipt])
                total_qty += qty; total_user_pay += user_pay; total_receipt += receipt
                current_row += 1
        
        total_row_title = title.replace("各商品", "").replace("汇总", "总计").strip()
        ws.append([total_row_title, "", total_qty, total_user_pay, total_receipt])
        for cell in ws[current_row]: cell.font = bold_font
        current_row += 2
        return total_qty, total_user_pay, total_receipt

    income_qty, income_user, income_receipt = write_summary_section(
        "各商品收入汇总 (所有未取消订单)", 'income_df',
        ["商品ID", "商品名称", "总销售数量", "用户实付总额(参考)", "总销售额"]
    )
    unshipped_qty, unshipped_user, unshipped_receipt = write_summary_section(
        "各商品支出汇总 (未发货退款)", 'unshipped_refund_df',
        ["商品ID", "商品名称", "退款数量", "用户实付总额(退款)", "总退款额"]
    )
    shipped_qty, shipped_user, shipped_receipt = write_summary_section(
        "各商品支出汇总 (已发货退款)", 'shipped_refund_df',
        ["商品ID", "商品名称", "退款数量", "用户实付(退款)", "退款额"]
    )

    net_qty1 = income_qty - unshipped_qty
    net_user1 = income_user + unshipped_user
    net_receipt1 = income_receipt + unshipped_receipt
    ws.cell(row=current_row, column=1, value="净总计(已发货退款订单按售出计算)").font = bold_font
    ws.cell(row=current_row, column=3, value=net_qty1).font = bold_font
    ws.cell(row=current_row, column=4, value=net_user1).font = bold_font
    ws.cell(row=current_row, column=5, value=net_receipt1).font = bold_font
    current_row += 1

    net_qty2 = income_qty - unshipped_qty - shipped_qty
    net_user2 = income_user + unshipped_user + shipped_user
    net_receipt2 = income_receipt + unshipped_receipt + shipped_receipt
    ws.cell(row=current_row, column=1, value="净总计(已发货退款订单按退款计算)").font = bold_font
    ws.cell(row=current_row, column=3, value=net_qty2).font = bold_font
    ws.cell(row=current_row, column=4, value=net_user2).font = bold_font
    ws.cell(row=current_row, column=5, value=net_receipt2).font = bold_font
    
    for col_letter, width in [('A', 35), ('B', 60), ('C', 18), ('D', 22), ('E', 22)]:
        ws.column_dimensions[col_letter].width = width
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, (int, float)):
                if cell.column == 3: cell.number_format = '#,##0'
                if cell.column in [4, 5]: cell.number_format = '#,##0.00'

def _create_detail_sheets(wb, product_data_map):
    """为每个商品创建并填充详情页。"""
    bold_font = Font(bold=True)
    center_align = Alignment(horizontal='center', vertical='center')

    def write_section(ws, df, title, total_title):
        if df.empty: return
        ws.append([title]); ws[ws.max_row][0].font = bold_font
        ws.append(DETAIL_SHEET_COLUMNS_PDD)
        for cell in ws[ws.max_row]: cell.font = bold_font; cell.alignment = center_align
        for _, row_data in df.iterrows(): ws.append(row_data.tolist())
        
        qty_sum = df['商品数量(件)'].sum()
        user_pay_sum = df['用户实付金额(元)'].sum()
        receipt_sum = df['商家实收金额(元)'].sum()
        total_row_data = [total_title, "", "", "", "", "", qty_sum, user_pay_sum, receipt_sum]
        ws.append(total_row_data)
        for cell in ws[ws.max_row]: cell.font = bold_font
        ws.append([])

    for prod_id in sorted(product_data_map.keys()):
        item = product_data_map[prod_id]
        sheet_name_raw = f"{prod_id}_{item['name']}"
        sheet_name = re.sub(r'[\\/\*\[\]\:?]', '_', sheet_name_raw)[:31]
        try: ws = wb.create_sheet(sheet_name)
        except: ws = wb.create_sheet(f"{prod_id}_detail")
            
        write_section(ws, item['income_df'], "收入明细 (所有未取消订单)", "收入总计")
        write_section(ws, item['unshipped_refund_df'], "支出明细 (未发货退款)", "未发货退款总计")
        write_section(ws, item['shipped_refund_df'], "支出明细 (已发货退款)", "已发货退款总计")

        for col_idx, width in enumerate([25, 20, 15, 25, 60, 22, 15, 18, 18, 20, 20, 20, 25, 15], 1):
             ws.column_dimensions[get_column_letter(col_idx)].width = width

# --- 主处理函数 (新架构) ---

def process_pdd_data(df_raw):
    """
    处理拼多多订单DataFrame，生成包含总结和明细的Excel Workbook对象。
    """
    if df_raw is None or df_raw.empty:
        print("拼多多处理错误：输入的DataFrame为空。")
        return None

    df_processed = _prepare_and_validate_data(df_raw)
    if df_processed is None: return None
        
    product_data_map = _aggregate_product_data(df_processed)
    
    wb = Workbook()
    _create_summary_sheet(wb, product_data_map)
    _create_detail_sheets(wb, product_data_map)
    
    if "销售总结" in wb.sheetnames and wb.sheetnames[0] != "销售总结":
         wb.move_sheet("销售总结", -len(wb.sheetnames)+1)
        
    return wb

# ---- 主程序入口 (用于独立测试) ----
if __name__ == "__main__":
    TEST_INPUT_DIR = r"C:\Users\LENOVO\Desktop"
    TEST_OUTPUT_DIR = r"C:\Users\LENOVO\Desktop"
    TEST_FILENAME = "orders_export2025.csv"

    input_file = os.path.join(TEST_INPUT_DIR, TEST_FILENAME)

    if not os.path.exists(input_file):
        print(f"错误: 测试输入文件未找到于 '{os.path.abspath(input_file)}'")
    else:
        print(f"--- 独立测试: 读取文件 {TEST_FILENAME} ---")
        try:
            df_test_raw = pd.read_csv(input_file, dtype=str, encoding='utf-8-sig')
            df_test_raw.columns = df_test_raw.columns.str.strip()
            for col in df_test_raw.columns:
                if df_test_raw[col].dtype == 'object':
                    df_test_raw[col] = df_test_raw[col].astype(str).str.strip().replace(
                        ['-', '--', '', 'None', 'nan', '#NULL!', None, 'null', '\t'], np.nan, regex=False)
        except Exception as e:
            print(f"测试中读取文件失败: {e}")
            df_test_raw = None

        if df_test_raw is not None:
            print("--- 独立测试: 调用 process_pdd_data ---")
            workbook_result = process_pdd_data(df_test_raw)
            
            if workbook_result:
                os.makedirs(TEST_OUTPUT_DIR, exist_ok=True)
                file_name_without_ext = os.path.splitext(TEST_FILENAME)[0]
                output_filename = f"PDD_output_{file_name_without_ext}.xlsx"
                output_file_path = os.path.join(TEST_OUTPUT_DIR, output_filename)
                
                try:
                    workbook_result.save(output_file_path)
                    print(f"\n脚本独立测试成功。输出文件位于: {output_file_path}")
                except Exception as e:
                    print(f"\n测试中保存文件失败: {e}")
            else:
                print("\n脚本独立测试失败，未能生成Workbook对象。")