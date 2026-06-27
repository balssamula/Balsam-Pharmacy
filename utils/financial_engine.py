import pandas as pd
import json
import numpy as np
from io import BytesIO
import gc

def extract_single_sku(combined_sku):
    if pd.isna(combined_sku) or str(combined_sku).strip() == "": return ""
    sku_str = str(combined_sku).strip()
    for char in ['*', '-', '+']:
        if char in sku_str: sku_str = sku_str.split(char)[0].strip()
    return sku_str

def fast_json_loads(val):
    if not isinstance(val, str) or not val.strip(): return []
    try: return json.loads(val, strict=False)
    except:
        try: return json.loads(val.replace('\\', '\\\\'), strict=False)
        except: return []

def process_sales_and_products(df_sales):
    col_map = {
        'الكمية': 'qty', 'اسم المنتج': 'product_name', 'أسم المنتج': 'product_name',
        'سعر المنتج': 'price', 'السعر': 'price', 'الإجمالي': 'product_total', 'المجموع': 'product_total'
    }
    df_sales = df_sales.rename(columns=col_map)
    
    if 'skus_json' in df_sales.columns:
        df_sales['skus_list'] = [fast_json_loads(x) for x in df_sales['skus_json']]
        df_exploded = df_sales.explode('skus_list').reset_index(drop=True)
        
        extracted = []
        for item in df_exploded['skus_list']:
            if isinstance(item, list) and len(item) >= 4:
                extracted.append({
                    'product_name': str(item[0]),
                    'qty': float(item[1]) if str(item[1]).replace('.','',1).isdigit() else 1.0,
                    'product_sku': extract_single_sku(item[2]),
                    'price': float(item[3]) if str(item[3]).replace('.','',1).isdigit() else 0.0,
                    'cost': float(item[4]) if len(item) > 4 and str(item[4]).replace('.','',1).isdigit() else 0.0
                })
            elif isinstance(item, dict):
                extracted.append({
                    'product_name': str(item.get('name', 'منتج غير محدد')),
                    'product_sku': extract_single_sku(item.get('sku', '')),
                    'qty': float(item.get('quantity', 1)),
                    'price': float(item.get('price', 0.0)),
                    'cost': float(item.get('cost_price', 0.0))
                })
            else:
                extracted.append({'product_name': 'منتج غير محدد', 'qty': 1.0, 'product_sku': '', 'price': 0.0, 'cost': 0.0})
        
        df_extracted = pd.DataFrame(extracted, index=df_exploded.index)
        df_final = pd.concat([df_exploded.drop(columns=['skus_json', 'skus_list']), df_extracted], axis=1)
        
        df_final['qty'] = pd.to_numeric(df_final['qty'], errors='coerce').fillna(1).astype('float32')
        df_final['price'] = pd.to_numeric(df_final['price'], errors='coerce').fillna(0).astype('float32')
        df_final['cost'] = pd.to_numeric(df_final['cost'], errors='coerce').fillna(0).astype('float32')
        
        has_sku = df_final['product_sku'].astype(str).str.strip() != ""
        df_final['product_display'] = df_final['product_name']
        df_final.loc[has_sku, 'product_display'] = df_final['product_name'] + " (SKU: " + df_final['product_sku'].astype(str) + ")"
        
        df_final['product_total'] = (df_final['price'] * df_final['qty']).astype('float32')
        df_final['total_cost'] = (df_final['cost'] * df_final['qty']).astype('float32')
        
        del df_sales, df_exploded, df_extracted
        gc.collect()
        return df_final
    else:
        for col in ['qty', 'product_total', 'total_cost', 'price', 'product_name', 'product_display']:
            if col not in df_sales.columns:
                if col == 'qty': df_sales[col] = 1.0
                elif col in ['product_name', 'product_display']: df_sales[col] = 'منتج غير محدد'
                else: df_sales[col] = 0.0
        df_sales['qty'] = pd.to_numeric(df_sales['qty'], errors='coerce').fillna(1).astype('float32')
        return df_sales

def calculate_financials(df_sales, df_profiles, df_payments, df_tabby, df_tamara, df_emkan, df_jnt, df_aramex, df_beez, manual_expenses):
    df = process_sales_and_products(df_sales)
    
    order_col = [c for c in df.columns if 'رقم الطلب' in c]
    order_col = order_col[0] if order_col else ('Order Number' if 'Order Number' in df.columns else None)
    
    df['gateway_fee'] = np.zeros(len(df), dtype='float32')
    df['shipping_cost'] = np.zeros(len(df), dtype='float32')
    if 'total_cost' not in df.columns: df['total_cost'] = np.zeros(len(df), dtype='float32')
    
    if order_col and order_col in df.columns:
        order_totals = df.groupby(order_col)['product_total'].sum()

        if df_profiles is not None and not df_profiles.empty:
            prof_order_col = [c for c in df_profiles.columns if 'Prescription No' in c or 'رقم الوصفة' in c or 'رقم الطلب' in c]
            if prof_order_col:
                cost_col = [c for c in df_profiles.columns if 'Amount' in c or 'Cost' in c or 'التكلفة' in c][-1]
                cost_map = df_profiles.groupby(prof_order_col[0])[cost_col].sum().to_dict()
                df['total_cost'] = df.apply(lambda row: cost_map.get(row[order_col], 0.0) * (row['product_total'] / max(order_totals.get(row[order_col], 1), 1)), axis=1).astype('float32')

        # 💡 تم تحديث دالة map_fees لتستقبل عمود الوجهة (target_col) لمنع اختلاط الشحن بالدفع
        def map_fees(df_fee, order_c, fee_c, target_col='gateway_fee'):
            if df_fee is not None and not df_fee.empty:
                f_map = df_fee.groupby(order_c)[fee_c].sum().to_dict()
                allocated_fee = df.apply(lambda row: f_map.get(str(row[order_col]), f_map.get(row[order_col], 0.0)) * (row['product_total'] / max(order_totals.get(row[order_col], 1), 1)), axis=1).astype('float32')
                df[target_col] += allocated_fee

        # رسوم البوابات تذهب لعمود gateway_fee
        if df_payments is not None: map_fees(df_payments, df_payments.columns[0], [c for c in df_payments.columns if 'رسوم' in c][0], 'gateway_fee')
        if df_tabby is not None: map_fees(df_tabby, [c for c in df_tabby.columns if 'Order' in c][0], [c for c in df_tabby.columns if 'Fee' in c or 'Deduction' in c][-1], 'gateway_fee')
        if df_tamara is not None: map_fees(df_tamara, [c for c in df_tamara.columns if 'Order ID' in c][0], [c for c in df_tamara.columns if 'Fees' in c][0], 'gateway_fee')
        if df_emkan is not None: map_fees(df_emkan, df_emkan.columns[0], df_emkan.columns[-1], 'gateway_fee')
        
        # 💡 رسوم الشحن تذهب لعمود shipping_cost
        if df_jnt is not None: map_fees(df_jnt, [c for c in df_jnt.columns if 'Client order' in c][0], [c for c in df_jnt.columns if 'Charge' in c][0], 'shipping_cost')
        if df_aramex is not None: map_fees(df_aramex, df_aramex.columns[0], df_aramex.columns[-1], 'shipping_cost')
        if df_beez is not None and not df_beez.empty: map_fees(df_beez, df_beez.columns[0], df_beez.columns[-1], 'shipping_cost')
        
        elif 'شركة الشحن' in df.columns:
            bosta_orders = df[df['شركة الشحن'].astype(str).str.contains('بيز|beez', case=False, na=False)][order_col].unique()
            for o_id in bosta_orders:
                items = df[df[order_col] == o_id]['qty'].sum()
                df.loc[df[order_col] == o_id, 'shipping_cost'] += (18.5 / max(items, 1)) * df.loc[df[order_col] == o_id, 'qty']

    df['marketing_commission'] = (df['product_total'] * 0.08).astype('float32')
    df['net_profit'] = (df['product_total'] - df['total_cost'] - df['gateway_fee'] - df['shipping_cost'] - df['marketing_commission']).astype('float32')
    
    if 'تاريخ الطلب' in df.columns and 'product_display' in df.columns and 'qty' in df.columns:
        df['تاريخ الطلب'] = pd.to_datetime(df['تاريخ الطلب'], errors='coerce')
        max_date = df['تاريخ الطلب'].max()
        if pd.notnull(max_date):
            last_7_days = df[df['تاريخ الطلب'] >= (max_date - pd.Timedelta(days=7))]
            prev_21_days = df[(df['تاريخ الطلب'] >= (max_date - pd.Timedelta(days=28))) & (df['تاريخ الطلب'] < (max_date - pd.Timedelta(days=7)))]
            
            l7 = last_7_days.groupby('product_display')['qty'].sum()
            p21 = prev_21_days.groupby('product_display')['qty'].sum() / 3
            df['momentum'] = df['product_display'].map(l7 / p21.replace(0, 1)).fillna(0).astype('float32')
        else:
            df['momentum'] = 0.0
    else:
        df['momentum'] = 0.0
        
    df['brand'] = df['product_name'].apply(lambda x: str(x).split(' ')[0] if pd.notna(x) else 'أخرى')
    total_manual_expenses = sum([float(exp['amount']) for exp in manual_expenses]) if manual_expenses else 0.0
    
    gc.collect() 
    return df, total_manual_expenses

def export_advanced_excel(df, rfm_df):
    output = BytesIO()
    export_df = df.copy()
    
    for col in export_df.columns:
        if pd.api.types.is_datetime64_any_dtype(export_df[col]):
            export_df[col] = export_df[col].dt.tz_localize(None)
            
    for col in export_df.select_dtypes(include=['object']).columns:
        export_df[col] = export_df[col].fillna("")

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        export_df.to_excel(writer, sheet_name='البيانات الشاملة', index=False)
        if rfm_df is not None: 
            rfm_df.to_excel(writer, sheet_name='RFM العملاء', index=False)
        
        bleeders = export_df[export_df['net_profit'] < 0].groupby('product_display').agg({'qty':'sum', 'net_profit':'sum'}).sort_values('net_profit')
        bleeders.to_excel(writer, sheet_name='المنتجات النازفة')
        
        if 'brand' in export_df.columns:
            brand_prof = export_df.groupby('brand').agg({'product_total':'sum', 'net_profit':'sum'}).sort_values('net_profit', ascending=False)
            brand_prof.to_excel(writer, sheet_name='ربحية العلامات التجارية')
    
    gc.collect()
    return output.getvalue()
