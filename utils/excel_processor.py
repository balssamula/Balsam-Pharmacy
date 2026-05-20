import pandas as pd
import numpy as np
from utils.helpers import (
    normalize_order_number, normalize_sku, normalize_text, normalize_city,
    determine_branch, get_branch_number, is_gift_or_promotion
)
from utils.database import DB_PATH, now_str

def prepare_salla_frame(df_salla: pd.DataFrame) -> pd.DataFrame:
    df = df_salla.copy()
    df["order_number"] = df["رقم الطلب"].apply(normalize_order_number)
    df["sku"] = df["SKU"].apply(normalize_sku)
    df["product_name"] = df["اسم المنتج"].apply(normalize_text)
    df["quantity"] = pd.to_numeric(df["الكمية"], errors="coerce").fillna(0)
    df["customer_name"] = df["اسم العميل"].apply(normalize_text)
    df["customer_phone"] = df["رقم الجوال"].apply(normalize_text)
    df["city"] = df["المدينة"].apply(normalize_text)
    df["order_status"] = df["حالة الطلب"].apply(normalize_text)
    df["order_date"] = df["تاريخ الطلب"].apply(normalize_text)
    df["total_amount"] = pd.to_numeric(df["إجمالي الطلب"], errors="coerce").fillna(0)

    # استبعاد عملاء الهدية والدعاية
    df = df[~df["customer_name"].apply(is_gift_or_promotion)]

    df = df[
        (df["order_number"] != "")
        & (df["sku"] != "")
        & (df["quantity"] != 0)
        & (df["order_status"] != "محذوف")
    ].copy()

    branch_info = df.apply(lambda row: determine_branch(row["order_status"], row["city"]), axis=1)
    df["pharmacy_name"] = branch_info.apply(lambda value: value[0])
    df["branch_number"] = branch_info.apply(lambda value: value[1])

    grouped = (
        df.groupby(["order_number", "sku"], as_index=False)
        .agg(
            {
                "product_name": "first",
                "quantity": "sum",
                "customer_name": "first",
                "customer_phone": "first",
                "city": "first",
                "order_status": "first",
                "order_date": "first",
                "total_amount": "first",
                "pharmacy_name": "first",
                "branch_number": "first",
            }
        )
        .rename(
            columns={
                "product_name": "salla_product_name",
                "quantity": "salla_qty",
                "pharmacy_name": "salla_pharmacy_name",
                "branch_number": "salla_branch_number",
            }
        )
    )
    return grouped

def prepare_abc_frame(df_abc: pd.DataFrame) -> pd.DataFrame:
    df = df_abc.copy()
    
    EXCLUDED_PROFILE = "FREE GIFTS FOR CUSTOMERS"
    if "نوع البروفايل" in df.columns:
        df = df[df["نوع البروفايل"].astype(str).str.strip() != EXCLUDED_PROFILE].copy()
    
    if "اسم الصنف" in df.columns:
        df = df[~df["اسم الصنف"].astype(str).str.upper().str.contains("DELIVERY FEE", na=False)].copy()
    
    if "رقم الصنف" in df.columns:
        df = df[df["رقم الصنف"].astype(str).str.strip() != "16133"].copy()

    df["order_number"] = df["رقم الطلب"].apply(normalize_order_number)
    df["sku"] = df["رقم الصنف"].apply(normalize_sku)
    df["abc_product_name"] = df["اسم الصنف"].apply(normalize_text)
    df["abc_qty"] = pd.to_numeric(df["Net Sold Qty"], errors="coerce").fillna(0)
    df["invoice_number"] = df["رقم الفاتورة"].apply(normalize_text)
    df["invoice_date"] = df["التاريخ"].apply(normalize_text)
    df["abc_pharmacy_name"] = df["رقم الصيدلية"].apply(normalize_text)
    df["abc_pharmacist_name"] = df["الصيدلي"].apply(normalize_text) if "الصيدلي" in df.columns else ""
    df["all_abc_pharmacies"] = df["abc_pharmacy_name"]
    df["abc_branch_number"] = df["abc_pharmacy_name"].apply(get_branch_number)
    df["profile_type"] = df["نوع البروفايل"].apply(normalize_text) if "نوع البروفايل" in df.columns else ""
    
    if "Receipt Classification" in df.columns:
        df["receipt_classification"] = df["Receipt Classification"].apply(normalize_text)
    else:
        df["receipt_classification"] = ""

    df = df[
        (df["sku"] != "")
        & (df["order_number"] != "")
    ].copy()

    grouped = (
        df.groupby(["order_number", "sku"], as_index=False)
        .agg(
            {
                "abc_qty": "sum",
                "invoice_number": "first",
                "invoice_date": "first",
                "abc_product_name": "first",
                "abc_pharmacy_name": "first",
                "abc_pharmacist_name": "first",
                "abc_branch_number": "first",
                "profile_type": lambda values: " | ".join(
                    sorted({normalize_text(value) for value in values if normalize_text(value)})
                ),
                "receipt_classification": lambda values: " | ".join(
                    sorted({normalize_text(value) for value in values if normalize_text(value)})
                ),
                "all_abc_pharmacies": lambda values: " | ".join(
                    sorted({normalize_text(value) for value in values if normalize_text(value)})
                ),
            }
        )
    )
    grouped["other_branch_details"] = grouped.apply(
        lambda row: (
            f"تم بيع نفس الطلب/الصنف في فروع أخرى: {row['all_abc_pharmacies']}"
            if " | " in row["all_abc_pharmacies"]
            else ""
        ),
        axis=1,
    )
    return grouped

def update_balances(abc_file, salla_file):
    """تحديث أرصدة الفروع بناءً على ملف ABC"""
    try:
        df_abc = pd.read_excel(abc_file, skiprows=4)
        df_salla = pd.read_excel(salla_file)
        
        def get_abc_col(branch_num):
            return pd.to_numeric(df_abc.iloc[:, branch_num + 1], errors='coerce').fillna(0)
        
        item_key = df_abc.iloc[:, 0]
        
        tabuk_calc = np.floor(((get_abc_col(8) + get_abc_col(10) + get_abc_col(11) + get_abc_col(12) +
                               get_abc_col(14) + get_abc_col(15) + get_abc_col(16) + get_abc_col(17)) / 2) + get_abc_col(13))
        
        f9_calc = np.floor(((get_abc_col(1) + get_abc_col(3)) / 2) + get_abc_col(9))
        
        def create_map(values):
            return dict(zip(item_key, values.astype(int)))
        
        maps = {
            'tabuk': create_map(tabuk_calc),
            'f9': create_map(f9_calc),
            'f1': create_map(get_abc_col(1)), 'f2': create_map(get_abc_col(2)),
            'f3': create_map(get_abc_col(3)), 'f4': create_map(get_abc_col(4)),
            'f5': create_map(get_abc_col(5)), 'f6': create_map(get_abc_col(6)),
            'f7': create_map(get_abc_col(7)), 'f8': create_map(get_abc_col(8)),
            'f10': create_map(get_abc_col(10)), 'f11': create_map(get_abc_col(11)),
            'f12': create_map(get_abc_col(12)), 'f14': create_map(get_abc_col(14)),
            'f15': create_map(get_abc_col(15)), 'f16': create_map(get_abc_col(16)),
            'f17': create_map(get_abc_col(17))
        }
        
        df_updated = df_salla.copy()
        salla_id_col = 3
        
        col_mapping = {
            5: 'tabuk', 7: 'f8', 9: 'f9', 11: 'f11', 13: 'f15', 15: 'f16',
            17: 'f10', 21: 'f12', 23: 'f14', 25: 'f1', 27: 'f2', 29: 'f3',
            31: 'f4', 33: 'f5', 35: 'f6', 37: 'f7', 39: 'f17'
        }
        
        for col_idx, map_name in col_mapping.items():
            df_updated.iloc[:, col_idx] = df_updated.iloc[:, salla_id_col].map(maps[map_name]).fillna(0).astype(int)
        
        cols_to_check = list(col_mapping.keys())
        old_data = df_salla.iloc[:, cols_to_check].apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)
        new_data = df_updated.iloc[:, cols_to_check]
        
        is_different = (new_data.values != old_data.values).any(axis=1)
        has_balance = new_data.sum(axis=1) > 0
        
        df_final = df_updated[is_different & has_balance]
        
        return df_final, len(df_final)
    except Exception as e:
        return None, str(e)

def process_excel(uploaded_file, uploaded_by: str):
    from utils.database import persist_reconciliation_results
    
    df_salla = pd.read_excel(uploaded_file, sheet_name="سلة")
    df_abc = pd.read_excel(uploaded_file, sheet_name="abc")
    results = classify_cases(df_salla, df_abc)
    upload_batch_id = persist_reconciliation_results(results, uploaded_file.name, uploaded_by)
    return results, upload_batch_id
