import pandas as pd
import json
import numpy as np
from io import BytesIO
import xlsxwriter
from datetime import timedelta

def extract_single_sku(combined_sku):
    if pd.isna(combined_sku) or combined_sku == "": return ""
    sku_str = str(combined_sku).strip()
    for char in ['*', '-', '+']:
        if char in sku_str: sku_str = sku_str.split(char)[0].strip()
    return sku_str

def process_sales_and_products(df_sales):
    if 'skus_json' not in df_sales.columns: return df_sales
    df_sales['skus_list'] = df_sales['skus_json'].apply(lambda x: json.loads(x) if isinstance(x, str) else [])
    df_exploded = df_sales.explode('skus_list').reset_index(drop=True)
    
    def parse_sku_array(item):
        if isinstance(item, list) and len(item) >= 4:
            return pd.Series({'product_name': str(item[0]), 'qty': float(item[1]), 'product_sku': extract_single_sku(item[2]), 'price': float(item[3])})
        elif isinstance(item, dict):
            return pd.Series({'product_name': item.get('name', ''), 'product_sku': extract_single_sku(item.get('sku', '')), 'qty': float(item.get('quantity', 1)), 'price': float(item.get('price', 0.0))})
        return pd.Series({'product_name': '', 'qty': 0, 'price': 0.0})

    parsed_skus = df_exploded['skus_list'].apply(parse_sku_array)
    df_final = pd.concat([df_exploded.drop(columns=['skus_json', 'skus_list']), parsed_skus], axis=1)
    df_final['product_total'] = df_final['price'] * df_final['qty']
    return df_final

def calculate_financials(df_sales, df_profiles, df_payments, df_tabby, df_tamara, df_emkan, df_jnt, df_aramex, df_beez, manual_expenses):
    df = process_sales_and_products(df_sales)
    order_col = [c for c in df.columns if 'رقم الطلب' in c][0] if any('رقم الطلب' in c for c in df.columns) else 'Order Number'
    
    df['gateway_fee'] = 0.0
    df['shipping_cost'] = 0.0
    df['total_cost'] = 0.0
    order_totals = df.groupby(order_col)['product_total'].sum()

    # 1. التكلفة من ملف البروفايلات (Prescription No.)
    if df_profiles is not None and not df_profiles.empty:
        prof_order_col = [c for c in df_profiles.columns if 'Prescription No' in c or 'رقم الوصفة' in c or 'رقم الطلب' in c][0]
        # افتراض أن عمود التكلفة اسمه 'Net Amount' أو 'Cost'
        cost_col = [c for c in df_profiles.columns if 'Amount' in c or 'Cost' in c or 'التكلفة' in c][-1]
        cost_map = df_profiles.groupby(prof_order_col)[cost_col].sum().to_dict()
        df['total_cost'] = df.apply(lambda row: cost_map.get(row[order_col], 0.0) * (row['product_total'] / max(order_totals.get(row[order_col], 1), 1)), axis=1)

    # 2. بوابات الدفع (مدى، فيزا، إمكان، تابي، تمارا)
    def map_fees(df_fee, order_c, fee_c):
        if df_fee is not None and not df_fee.empty:
            f_map = df_fee.groupby(order_c)[fee_c].sum().to_dict()
            df['gateway_fee'] += df.apply(lambda row: f_map.get(str(row[order_col]), 0.0) * (row['product_total'] / max(order_totals.get(row[order_col], 1), 1)), axis=1)

    if df_payments is not None: map_fees(df_payments, df_payments.columns[0], [c for c in df_payments.columns if 'رسوم' in c][0])
    if df_tabby is not None: map_fees(df_tabby, [c for c in df_tabby.columns if 'Order' in c][0], [c for c in df_tabby.columns if 'Fee' in c or 'Deduction' in c][-1])
    if df_tamara is not None: map_fees(df_tamara, [c for c in df_tamara.columns if 'Order ID' in c][0], [c for c in df_tamara.columns if 'Fees' in c][0])
    if df_emkan is not None: map_fees(df_emkan, df_emkan.columns[0], df_emkan.columns[-1]) # يفترض أن الإدخال اليدوي لإمكان تم ترتيبه

    # 3. تكاليف الشحن
    if df_jnt is not None: map_fees(df_jnt, [c for c in df_jnt.columns if 'Client order' in c][0], [c for c in df_jnt.columns if 'Charge' in c][0])
    if df_aramex is not None: map_fees(df_aramex, df_aramex.columns[0], df_aramex.columns[-1])
    
    # بيز (18.5 ثابت أو من ملف)
    if df_beez is not None and not df_beez.empty:
        map_fees(df_beez, df_beez.columns[0], df_beez.columns[-1])
    elif 'شركة الشحن' in df.columns:
        bosta_orders = df[df['شركة الشحن'].astype(str).str.contains('بيز|beez', case=False, na=False)][order_col].unique()
        for o_id in bosta_orders:
            items = df[df[order_col] == o_id]['qty'].sum()
            df.loc[df[order_col] == o_id, 'shipping_cost'] += (18.5 / max(items, 1)) * df.loc[df[order_col] == o_id, 'qty']

    # 4. الحسابات النهائية
    df['marketing_commission'] = df['product_total'] * 0.08
    df['net_profit'] = df['product_total'] - df['total_cost'] - df['gateway_fee'] - df['shipping_cost'] - df['marketing_commission']
    df['profit_margin'] = np.where(df['product_total'] > 0, (df['net_profit'] / df['product_total']), 0)
    
    # 5. حساب الزخم (Momentum) للمنتجات
    if 'تاريخ الطلب' in df.columns:
        df['تاريخ الطلب'] = pd.to_datetime(df['تاريخ الطلب'], errors='coerce')
        max_date = df['تاريخ الطلب'].max()
        last_7_days = df[df['تاريخ الطلب'] >= (max_date - timedelta(days=7))]
        prev_21_days = df[(df['تاريخ الطلب'] >= (max_date - timedelta(days=28))) & (df['تاريخ الطلب'] < (max_date - timedelta(days=7)))]
        
        l7_sales = last_7_days.groupby('product_name')['qty'].sum()
        p21_sales = prev_21_days.groupby('product_name')['qty'].sum() / 3 # متوسط أسبوعي
        
        df['momentum'] = df['product_name'].map(l7_sales / p21_sales.replace(0, 1)).fillna(0)

    total_manual_expenses = sum([float(exp['amount']) for exp in manual_expenses]) if manual_expenses else 0.0
    return df, total_manual_expenses

def export_advanced_excel(df, rfm_df, shipping_stats, payment_stats):
    """توليد ملف Excel احترافي مع فلاتر، ألوان، ورسوم بيانية لـ Power BI"""
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    
    # 1. شيت تفاصيل الطلبات (للـ Power BI)
    df.to_excel(writer, sheet_name='قاعدة البيانات الشاملة', index=False)
    workbook = writer.book
    worksheet = writer.sheets['قاعدة البيانات الشاملة']
    
    # التنسيق الاحترافي
    header_format = workbook.add_format({'bold': True, 'bg_color': '#16425b', 'font_color': 'white', 'border': 1, 'align': 'center'})
    money_format = workbook.add_format({'num_format': '#,##0.00'})
    
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)
        worksheet.set_column(col_num, col_num, 15)
        if 'profit' in str(value).lower() or 'cost' in str(value).lower() or 'fee' in str(value).lower() or 'total' in str(value).lower():
            worksheet.set_column(col_num, col_num, 15, money_format)
            
    worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
    
    # 2. شيت المنتجات النازفة
    bleeders = df[df['net_profit'] < 0].groupby('product_name').agg({'qty':'sum', 'net_profit':'sum'}).sort_values('net_profit')
    bleeders.to_excel(writer, sheet_name='المنتجات النازفة (تحذير)')
    
    # 3. شيت العملاء RFM
    if rfm_df is not None: rfm_df.to_excel(writer, sheet_name='تحليل العملاء RFM', index=False)
    
    # 4. إضافة رسم بياني (Chart) داخل الإكسيل
    chart_sheet = workbook.add_worksheet('ملخص الرسوم البيانية')
    chart = workbook.add_chart({'type': 'column'})
    
    # بيانات تلخيصية للرسم
    summary = df.groupby('شركة الشحن')['net_profit'].sum().reset_index()
    summary.to_excel(writer, sheet_name='ملخص الرسوم البيانية', startrow=1, index=False)
    
    chart.add_series({
        'categories': ['ملخص الرسوم البيانية', 2, 0, len(summary)+1, 0],
        'values':     ['ملخص الرسوم البيانية', 2, 1, len(summary)+1, 1],
        'name':       'صافي الربح حسب شركة الشحن',
        'fill':       {'color': '#1f7a8c'}
    })
    chart_sheet.insert_chart('E2', chart)
    
    writer.close()
    return output.getvalue()
