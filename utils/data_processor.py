import pandas as pd
import json
import numpy as np

def explode_skus(df):
    """تفكيك الـ skus_json وتحويل الطلبات إلى صفوف منتجات مفردة"""
    df['skus_json'] = df['skus_json'].apply(lambda x: json.loads(x) if isinstance(x, str) else [])
    df = df.explode('skus_json')
    # استخراج البيانات من الـ JSON المفكك (الاسم، الكمية، السعر)
    df_skus = pd.json_normalize(df['skus_json'])
    df = df.reset_index(drop=True).join(df_skus.reset_index(drop=True))
    return df

def calculate_net_profit(df_sales, df_shipping, df_payments, other_expenses, marketing_costs):
    """دمج المبيعات بالمصاريف وحساب الربح الصافي"""
    # 1. دمج المصاريف (الشحن، الدفع، التسويق) بناءً على رقم الطلب
    # 2. حساب عمولة شركة 'هاي تك' (8% من إجمالي المبيعات)
    df_sales['marketing_comm_high_tech'] = df_sales['order_amount'] * 0.08
    
    # 3. حساب صافي الربح
    # Net Profit = (Sales - Discount) - (COGS) - (Shipping) - (Gateway Fees) - (Marketing)
    df_sales['net_profit'] = df_sales['order_amount'] - df_sales['cogs'] - df_sales['shipping_fee'] - df_sales['gateway_fees'] - df_sales['marketing_comm_high_tech']
    
    return df_sales
