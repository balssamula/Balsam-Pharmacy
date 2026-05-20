import pandas as pd
import numpy as np
from utils.helpers import (
    normalize_order_number, normalize_sku, normalize_text, 
    determine_branch, get_branch_number, is_gift_or_promotion
)

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
    df["abc_pharmacist_name"] = df["الصيدلي"].apply(normalize_text) if "
