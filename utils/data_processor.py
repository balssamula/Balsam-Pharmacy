# utils/data_processor.py
import pandas as pd
import json

def process_sales_file(file_path):
    """تفكيك مبيعات سلة و JSON الخاص بالمنتجات"""
    df = pd.read_excel(file_path)
    # تفكيك عمود الـ JSON إلى منتجات مفردة
    df['skus_list'] = df['skus_json'].apply(json.loads)
    df = df.explode('skus_list')
    # استخراج تفاصيل المنتج من JSON
    skus_df = pd.json_normalize(df['skus_list'])
    df = df.drop(columns=['skus_json', 'skus_list']).join(skus_df)
    return df

def apply_financial_logic(df, shipping_fees, gateway_fees):
    """دمج ملفات المصاريف (الشحن + الدفع) وحساب الربح الصافي"""
    # ربط طلبات الشحن
    df = df.merge(shipping_fees, on='رقم الطلب', how='left')
    # ربط بوابات الدفع
    df = df.merge(gateway_fees, on='رقم الطلب', how='left')
    
    # حساب المصاريف التراكمية
    df['total_shipping'] = df['shipping_amount'] # رسوم الشحن الفعلية
    df['gateway_fees_calc'] = df.apply(lambda x: calculate_gateway_fee(x['amount'], x['payment_method']), axis=1)
    df['marketing_comm'] = df['total_price'] * 0.08 # عمولة هاي تك 8%
    
    # صافي الربح
    df['net_profit'] = df['total_price'] - df['cost'] - df['total_shipping'] - df['gateway_fees_calc'] - df['marketing_comm']
    return df

def calculate_gateway_fee(amount, method):
    """معادلة رسوم بوابات الدفع (مثال: مدى +0.8 ريال، تابي نسبة+ثابت)"""
    if method == 'mada': return (amount * 0.005) + 0.8
    if method == 'tabby': return (amount * 0.02) + 1.5
    # إلخ لكل وسيلة دفع...
    return 0
