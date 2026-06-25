# utils/financial_engine.py
import pandas as pd
import json
import numpy as np

def extract_single_sku(combined_sku):
    """تنظيف واستخراج الـ SKU النقي لتسهيل المطابقة"""
    if pd.isna(combined_sku) or combined_sku == "": return ""
    sku_str = str(combined_sku).strip()
    for char in ['*', '-', '+']:
        if char in sku_str: sku_str = sku_str.split(char)[0].strip()
    return sku_str

def process_sales_and_products(df_sales):
    """تفكيك مبيعات سلة (skus_json) وتحويلها لصفوف منتجات مفردة"""
    if 'skus_json' not in df_sales.columns:
        return df_sales
        
    df_sales['skus_list'] = df_sales['skus_json'].apply(
        lambda x: json.loads(x) if isinstance(x, str) else []
    )
    df_exploded = df_sales.explode('skus_list').reset_index(drop=True)
    
    # 🧠 [خوارزمية قراءة هيكل سلة]: 
    # مثال: ["Humana Baby Milk", 2, "17257*3", 445.44, 721.62]
    def parse_sku_array(item):
        if isinstance(item, list) and len(item) >= 4:
            return pd.Series({
                'product_name': str(item[0]),
                'qty': float(item[1]),
                'product_sku': extract_single_sku(item[2]),
                'price': float(item[3]),
                'cost': float(item[4]) if len(item) > 4 else 0.0
            })
        elif isinstance(item, dict):
            return pd.Series({
                'product_name': item.get('name', ''),
                'product_sku': extract_single_sku(item.get('sku', '')),
                'qty': float(item.get('quantity', 1)),
                'price': float(item.get('price', 0.0)),
                'cost': float(item.get('cost_price', 0.0))
            })
        return pd.Series({'product_name': '', 'qty': 0, 'price': 0.0, 'cost': 0.0})

    parsed_skus = df_exploded['skus_list'].apply(parse_sku_array)
    df_final = pd.concat([df_exploded.drop(columns=['skus_json', 'skus_list']), parsed_skus], axis=1)
    
    df_final['product_total'] = df_final['price'] * df_final['qty']
    df_final['total_cost'] = df_final['cost'] * df_final['qty']
    
    return df_final

def calculate_financials(df_sales, df_payments=None, df_tabby=None, df_tamara=None, df_jnt=None, manual_expenses=None):
    """
    محرك الحسابات المالية المركزي:
    دمج المبيعات + رسوم البوابات + تابي + تمارا + الشحن (J&T و بوسطة)
    """
    df = process_sales_and_products(df_sales)
    
    # توحيد اسم عمود رقم الطلب
    order_col = [c for c in df.columns if 'رقم الطلب' in c][0] if any('رقم الطلب' in c for c in df.columns) else 'Order Number'
    df['gateway_fee'] = 0.0
    df['shipping_cost'] = 0.0
    
    order_totals = df.groupby(order_col)['product_total'].sum()

    # 1️⃣ دمج رسوم الدفع الإلكتروني (مدى، فيزا)
    if df_payments is not None and not df_payments.empty:
        pay_col = [c for c in df_payments.columns if 'رقم الطلب' in c][0]
        fee_col = [c for c in df_payments.columns if 'الرسوم' in c][0]
        fees_map = df_payments.groupby(pay_col)[fee_col].sum().to_dict()
        df['gateway_fee'] += df.apply(lambda row: fees_map.get(row[order_col], 0.0) * (row['product_total'] / max(order_totals.get(row[order_col], 1), 1)), axis=1)

    # 2️⃣ دمج رسوم تابي (Tabby Settlement)
    if df_tabby is not None and not df_tabby.empty:
        tabby_order_col = [c for c in df_tabby.columns if 'Order Number' in c][0]
        tabby_fee_col = [c for c in df_tabby.columns if 'Total Deduction' in c or 'Total Fee' in c][-1]
        tabby_map = df_tabby.groupby(tabby_order_col)[tabby_fee_col].sum().to_dict()
        df['gateway_fee'] += df.apply(lambda row: tabby_map.get(str(row[order_col]), 0.0) * (row['product_total'] / max(order_totals.get(row[order_col], 1), 1)), axis=1)

    # 3️⃣ دمج رسوم تمارا (Tamara Invoice)
    if df_tamara is not None and not df_tamara.empty:
        tamara_order_col = [c for c in df_tamara.columns if 'Merchant Order ID' in c][0]
        tamara_fee_col = [c for c in df_tamara.columns if 'Total Fees' in c][0]
        tamara_map = df_tamara.groupby(tamara_order_col)[tamara_fee_col].sum().to_dict()
        df['gateway_fee'] += df.apply(lambda row: tamara_map.get(row[order_col], 0.0) * (row['product_total'] / max(order_totals.get(row[order_col], 1), 1)), axis=1)

    # 4️⃣ دمج رسوم الشحن J&T
    if df_jnt is not None and not df_jnt.empty:
        jnt_order_col = [c for c in df_jnt.columns if 'Client order No' in c][0]
        jnt_fee_col = [c for c in df_jnt.columns if 'Total Charge' in c][0]
        jnt_map = df_jnt.groupby(jnt_order_col)[jnt_fee_col].sum().to_dict()
        
        def allocate_shipping(row):
            o_id = str(row[order_col]).strip()
            items_in_order = df[df[order_col].astype(str).str.strip() == o_id]['qty'].sum()
            order_shipping = jnt_map.get(o_id, 0.0)
            return (order_shipping / max(items_in_order, 1)) * row['qty']
            
        df['shipping_cost'] += df.apply(allocate_shipping, axis=1)
        
    # 5️⃣ تسعير الشحن الافتراضي الثابت لشركة بيز (18.5 ريال للطلب) وتوزيعه
    if 'شركة الشحن' in df.columns:
        bosta_orders = df[df['شركة الشحن'].astype(str).str.contains('بيز', case=False, na=False)][order_col].unique()
        for o_id in bosta_orders:
            items_in_order = df[df[order_col] == o_id]['qty'].sum()
            df.loc[df[order_col] == o_id, 'shipping_cost'] = (18.5 / max(items_in_order, 1)) * df.loc[df[order_col] == o_id, 'qty']

    # 6️⃣ حساب عمولة شركة هاي تك للتسويق (8% من المبيعات)
    df['marketing_commission'] = df['product_total'] * 0.08
    
    # 7️⃣ حساب الربح الصافي
    # صافي الربح = إجمالي بيع المنتج - تكلفته - رسوم الدفع - رسوم الشحن - عمولة التسويق
    df['net_profit'] = df['product_total'] - df['total_cost'] - df['gateway_fee'] - df['shipping_cost'] - df['marketing_commission']
    
    # 8️⃣ تجميع المصروفات اليدوية
    total_manual_expenses = sum([float(exp['amount']) for exp in manual_expenses]) if manual_expenses else 0.0
    
    return df, total_manual_expenses
