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

    df = df[~df["customer_name"].apply(is_gift_or_promotion)]
    df = df[(df["order_number"] != "") & (df["sku"] != "") & (df["quantity"] != 0) & (df["order_status"] != "محذوف")].copy()

    branch_info = df.apply(lambda row: determine_branch(row["order_status"], row["city"]), axis=1)
    df["pharmacy_name"] = branch_info.apply(lambda x: x[0])
    df["branch_number"] = branch_info.apply(lambda x: x[1])

    grouped = df.groupby(["order_number", "sku"], as_index=False).agg({
        "product_name": "first", "quantity": "sum", "customer_name": "first",
        "customer_phone": "first", "city": "first", "order_status": "first",
        "order_date": "first", "total_amount": "first", "pharmacy_name": "first", "branch_number": "first"
    }).rename(columns={
        "product_name": "salla_product_name", "quantity": "salla_qty",
        "pharmacy_name": "salla_pharmacy_name", "branch_number": "salla_branch_number"
    })
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
    df["profile_type"] = df["نوع البروفايل"].apply(normalize_text) if "نوع البروفايل" in df.columns else ""
    if "Receipt Classification" in df.columns:
        df["receipt_classification"] = df["Receipt Classification"].apply(normalize_text)
    else:
        df["receipt_classification"] = ""

    df = df[(df["sku"] != "") & (df["order_number"] != "")].copy()

    grouped = df.groupby(["order_number", "sku"], as_index=False).agg({
        "abc_qty": "sum", "invoice_number": "first", "invoice_date": "first",
        "abc_product_name": "first", "abc_pharmacy_name": "first", "abc_pharmacist_name": "first",
        "profile_type": lambda x: " | ".join(sorted({normalize_text(v) for v in x if normalize_text(v)})),
        "receipt_classification": lambda x: " | ".join(sorted({normalize_text(v) for v in x if normalize_text(v)})),
        "all_abc_pharmacies": lambda x: " | ".join(sorted({normalize_text(v) for v in x if normalize_text(v)}))
    })
    grouped["other_branch_details"] = grouped.apply(
        lambda row: f"تم بيع نفس الطلب/الصنف في فروع أخرى: {row['all_abc_pharmacies']}" if " | " in row["all_abc_pharmacies"] else "",
        axis=1
    )
    return grouped

def classify_cases(df_salla: pd.DataFrame, df_abc: pd.DataFrame) -> pd.DataFrame:
    salla_grouped = prepare_salla_frame(df_salla)
    abc_grouped = prepare_abc_frame(df_abc)
    merged = pd.merge(salla_grouped, abc_grouped, on=["order_number", "sku"], how="outer", indicator=True)

    for col in ["salla_qty", "abc_qty", "total_amount"]:
        if col in merged.columns:
            merged[col] = pd.to_numeric(merged[col], errors="coerce").fillna(0)
        else:
            merged[col] = 0

    text_cols = ["salla_product_name", "abc_product_name", "customer_name", "customer_phone", "city",
                 "order_status", "order_date", "invoice_number", "invoice_date", "salla_pharmacy_name",
                 "abc_pharmacy_name", "abc_pharmacist_name", "profile_type", "receipt_classification",
                 "all_abc_pharmacies", "other_branch_details"]
    for col in text_cols:
        if col not in merged.columns:
            merged[col] = ""
        merged[col] = merged[col].fillna("").astype(str)

    merged["product_name"] = merged["salla_product_name"]
    merged.loc[merged["product_name"].eq(""), "product_name"] = merged.loc[merged["product_name"].eq(""), "abc_product_name"]
    merged["pharmacy_name"] = merged["salla_pharmacy_name"]
    merged.loc[merged["pharmacy_name"].eq(""), "pharmacy_name"] = merged.loc[merged["pharmacy_name"].eq(""), "abc_pharmacy_name"]
    merged["branch_number"] = merged["salla_branch_number"]
    merged.loc[merged["branch_number"].eq(""), "branch_number"] = merged.loc[merged["branch_number"].eq(""), "abc_branch_number"] if "abc_branch_number" in merged.columns else ""

    merged["difference"] = merged["salla_qty"] - merged["abc_qty"]
    merged["case_type"] = ""
    merged["case_reason"] = ""

    addition_mask = (merged["_merge"] == "both") & (merged["salla_qty"] > merged["abc_qty"]) & (merged["salla_qty"] > 0)
    merged.loc[addition_mask, "case_type"] = "addition"
    merged.loc[addition_mask, "case_reason"] = "كمية الطلب أعلى من كمية الفاتورة."

    return_mask = (merged["_merge"] == "both") & (merged["abc_qty"] > merged["salla_qty"])
    merged.loc[return_mask, "case_type"] = "return"
    merged.loc[return_mask, "case_reason"] = "كمية الفاتورة أعلى من كمية الطلب."

    orphan_salla_mask = (merged["_merge"] == "left_only") & (merged["salla_qty"] > 0)
    merged.loc[orphan_salla_mask, "case_type"] = "orphan_salla"
    merged.loc[orphan_salla_mask, "case_reason"] = "سطر طلب موجود في سلة ولم يُعثر على سطر مطابق له في ABC."

    orphan_abc_mask = (merged["_merge"] == "right_only") & (merged["abc_qty"] != 0)
    merged.loc[orphan_abc_mask, "case_type"] = "orphan_abc"
    merged.loc[orphan_abc_mask, "case_reason"] = "سطر فاتورة موجود في ABC ولم يُعثر على سطر مطابق له في سلة."

    result = merged[merged["case_type"] != ""].copy()
    result["case_label"] = result["case_type"]
    result["item_key"] = result.apply(lambda r: f"{r['pharmacy_name']}||{r['order_number']}||{r['sku']}||{r['case_type']}", axis=1)

    return result[["item_key", "order_number", "invoice_number", "sku", "product_name",
                   "salla_product_name", "abc_product_name", "pharmacy_name", "salla_pharmacy_name",
                   "abc_pharmacy_name", "abc_pharmacist_name", "branch_number", "salla_qty",
                   "abc_qty", "difference", "case_type", "case_label", "case_reason",
                   "customer_name", "customer_phone", "city", "order_status", "order_date",
                   "invoice_date", "profile_type", "receipt_classification", "all_abc_pharmacies",
                   "other_branch_details", "total_amount"]]

def process_excel(uploaded_file, uploaded_by: str):
    from utils.database import persist_reconciliation_results
    df_salla = pd.read_excel(uploaded_file, sheet_name="سلة")
    df_abc = pd.read_excel(uploaded_file, sheet_name="abc")
    results = classify_cases(df_salla, df_abc)
    upload_batch_id = persist_reconciliation_results(results, uploaded_file.name, uploaded_by)
    return results, upload_batch_id

def persist_reconciliation_results(results: pd.DataFrame, uploaded_file_name: str, uploaded_by: str):
    import uuid
    upload_batch_id = uuid.uuid4().hex
    timestamp = now_str()
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    
    cur.execute("""
        INSERT INTO uploads (upload_batch_id, file_name, uploaded_by, uploaded_at, total_cases,
            total_additions, total_returns, total_orphan_salla, total_orphan_abc, is_active)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
    """, (upload_batch_id, uploaded_file_name, uploaded_by, timestamp, len(results),
          int((results["case_type"] == "addition").sum()), int((results["case_type"] == "return").sum()),
          int((results["case_type"] == "orphan_salla").sum()), int((results["case_type"] == "orphan_abc").sum())))
    
    for _, row in results.iterrows():
        cur.execute("""
            INSERT OR REPLACE INTO reconciliation_items (
                item_key, upload_batch_id, order_number, invoice_number, sku, product_name,
                salla_product_name, abc_product_name, pharmacy_name, salla_pharmacy_name,
                abc_pharmacy_name, abc_pharmacist_name, branch_number, salla_qty, abc_qty, difference,
                case_type, case_label, case_reason, status, customer_name, customer_phone,
                city, order_status, order_date, invoice_date, profile_type, receipt_classification,
                all_abc_pharmacies, other_branch_details, total_amount, first_seen_at, last_seen_at, active)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
        """, (row["item_key"], upload_batch_id, str(row["order_number"]), str(row.get("invoice_number", "")),
              str(row["sku"]), str(row["product_name"])[:200], str(row.get("salla_product_name", ""))[:200],
              str(row.get("abc_product_name", ""))[:200], str(row["pharmacy_name"]),
              str(row.get("salla_pharmacy_name", "")), str(row.get("abc_pharmacy_name", "")),
              str(row.get("abc_pharmacist_name", "")), str(row.get("branch_number", "")),
              float(row["salla_qty"]), float(row["abc_qty"]), float(row["difference"]),
              str(row["case_type"]), str(row["case_label"]), str(row.get("case_reason", ""))[:500],
              "قيد المتابعة", str(row.get("customer_name", ""))[:100], str(row.get("customer_phone", "")),
              str(row.get("city", "")), str(row.get("order_status", "")), str(row.get("order_date", "")),
              str(row.get("invoice_date", "")), str(row.get("profile_type", "")), str(row.get("receipt_classification", "")),
              str(row.get("all_abc_pharmacies", "")), str(row.get("other_branch_details", "")),
              float(row.get("total_amount", 0)), timestamp, timestamp))
    
    cur.execute("UPDATE reconciliation_items SET active = CASE WHEN upload_batch_id = ? THEN 1 ELSE 0 END", (upload_batch_id,))
    cur.execute("UPDATE uploads SET is_active = 0")
    cur.execute("UPDATE uploads SET is_active = 1 WHERE upload_batch_id = ?", (upload_batch_id,))
    session_name = datetime.now().strftime("%Y-%m-%d %H:%M")
    cur.execute("UPDATE uploads SET session_name = ? WHERE upload_batch_id = ?", (session_name, upload_batch_id))
    
    conn.commit()
    conn.close()
    return upload_batch_id

from utils.database import DB_PATH, now_str
import sqlite3
from datetime import datetime
