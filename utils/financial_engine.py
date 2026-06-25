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
        
    # تحويل النص إلى كائنات JSON
    df_sales['skus_list'] = df_sales['skus_json'].apply(
        lambda x: json.loads(x) if isinstance(x, str) else []
    )
    
    # تفجير القائمة لتصبح كل قطعة في صف مستقل (Explode)
    df_exploded = df_sales.explode('skus_list').reset_index(drop=True)
    
    # استخراج بيانات المنتج (الاسم، السعر، الكمية، التكلفة إن وجدت)
    def parse_sku_dict(sku_dict):
        if not isinstance(sku_dict, dict): return pd.Series({'product_name': '', 'qty': 0, 'price': 0.0})
        # التوافق مع هيكلية سلة المعتادة
        return pd.Series({
            'product_name': sku_dict.get('name', ''),
            'product_sku': extract_single_sku(sku_dict.get('sku', '')),
            'qty': float(sku_dict.get('quantity', 1)),
            'price': float(sku_dict.get('price', 0.0)),
            'cost': float(sku_dict.get('cost_price', 0.0)) # يسحب التكلفة إذا كانت مسجلة بالمنصة
        })

    parsed_skus = df_exploded['skus_list'].apply(parse_sku_dict)
    df_final = pd.concat([df_exploded.drop(columns=['skus_json', 'skus_list']), parsed_skus], axis=1)
    
    # حساب إجمالي قيمة المنتج في الطلب
    df_final['product_total'] = df_final['price'] * df_final['qty']
    return df_final

def calculate_financials(df_sales, df_payments=None, df_shipping=None, manual_expenses=None):
    """
    المحرك المالي: دمج المصروفات، الرسوم، وحساب الربحية
    manual_expenses: list of dicts [{'desc': 'اعلانات تيك توك', 'amount': 1500}, ...]
    """
    df = process_sales_and_products(df_sales)
    
    # التأكد من وجود عمود رقم الطلب القياسي
    order_col = [c for c in df.columns if 'رقم الطلب' in c or 'Order Number' in c]
    order_col = order_col[0] if order_col else 'رقم الطلب'
    
    # 1. معالجة رسوم بوابات الدفع (مدى، فيزا، تمارا، تابي)
    df['gateway_fee'] = 0.0
    if df_payments is not None and not df_payments.empty:
        pay_order_col = [c for c in df_payments.columns if 'رقم الطلب' in c or 'Order Number' in c][0]
        fee_col = [c for c in df_payments.columns if 'رسوم' in c or 'Fee' in c or 'Deduction' in c][0]
        
        # تجميع الرسوم برقم الطلب (لضمان عدم التكرار)
        fees_map = df_payments.groupby(pay_order_col)[fee_col].sum().to_dict()
        
        # توزيع رسوم الدفع على المنتجات بنسبة وتناسب (حسب قيمة المنتج من إجمالي الطلب)
        order_totals = df.groupby(order_col)['product_total'].sum()
        
        def allocate_fee(row):
            o_id = row[order_col]
            total_order_val = order_totals.get(o_id, 1)
            if total_order_val == 0: total_order_val = 1
            order_fee = fees_map.get(o_id, 0.0)
            return order_fee * (row['product_total'] / total_order_val)
            
        df['gateway_fee'] = df.apply(allocate_fee, axis=1)
    
    # 2. معالجة تكاليف الشحن (بيز، أرامكس، سمسا، جي آند تي)
    df['shipping_cost'] = 0.0
    if df_shipping is not None and not df_shipping.empty:
        ship_order_col = [c for c in df_shipping.columns if 'رقم الطلب' in c or 'Reference' in c or 'Client order' in c][0]
        ship_fee_col = [c for c in df_shipping.columns if 'Total' in c or 'Charge' in c or 'تكلفة' in c][0]
        shipping_map = df_shipping.groupby(ship_order_col)[ship_fee_col].sum().to_dict()
        
        def allocate_shipping(row):
            o_id = row[order_col]
            # توزيع تكلفة الشحن بالتساوي على عدد القطع في الطلب
            items_in_order = df[df[order_col] == o_id]['qty'].sum()
            if items_in_order == 0: items_in_order = 1
            order_shipping = shipping_map.get(o_id, 0.0)
            return (order_shipping / items_in_order) * row['qty']
            
        df['shipping_cost'] = df.apply(allocate_shipping, axis=1)
    else:
        # افتراضي: إذا لم يرفع ملف شحن، وكان الطلب عبر "بيز"، نحتسب 18.5 ريال
        df.loc[df['شركة الشحن'].astype(str).str.contains('بيز|bosta', case=False, na=False), 'shipping_cost'] = 18.5
    
    # 3. حساب عمولة شركة هاي تك للتسويق (8% من المبيعات)
    df['marketing_commission'] = df['product_total'] * 0.08
    
    # 4. حساب صافي الربح الفردي (Net Profit)
    # ملاحظة: إذا كانت التكلفة (cost) تساوي 0، سيحسب النظام الهامش الإجمالي قبل تكلفة البضاعة.
    df['net_profit'] = df['product_total'] - df['cost'] - df['gateway_fee'] - df['shipping_cost'] - df['marketing_commission']
    
    # 5. تجميع المصروفات اليدوية العامة (OPEX)
    total_manual_expenses = sum([float(exp['amount']) for exp in manual_expenses]) if manual_expenses else 0.0
    
    return df, total_manual_expenses
