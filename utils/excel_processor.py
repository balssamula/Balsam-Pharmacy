import pandas as pd
import numpy as np
import sqlite3
import uuid
from datetime import datetime
from utils.helpers import (
    normalize_order_number, normalize_sku, normalize_text,
    determine_branch, get_branch_number, is_gift_or_promotion, now_str
)
from utils.database import DB_PATH

def prepare_salla_frame(df_salla: pd.DataFrame) -> pd.DataFrame:
    """معالجة شيت سلة"""
    df = df_salla.copy()
    
    # تحديد الأعمدة المطلوبة
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
    
    # الأعمدة الجديدة من سلة
    df["discount"] = pd.to_numeric(df["الخصم"], errors="coerce").fillna(0) if "الخصم" in df.columns else 0
    df["shipping_cost"] = pd.to_numeric(df["تكلفة الشحن"], errors="coerce").fillna(0) if "تكلفة الشحن" in df.columns else 0
    df["payment_method"] = df["طريقة الدفع"].apply(normalize_text) if "طريقة الدفع" in df.columns else ""
    df["tax"] = pd.to_numeric(df["الضريبة"], errors="coerce").fillna(0) if "الضريبة" in df.columns else 0
    df["coupon_discount"] = pd.to_numeric(df["قيمة خصم الكوبون"], errors="coerce").fillna(0) if "قيمة خصم الكوبون" in df.columns else 0
    df["offer_discount"] = pd.to_numeric(df["قيمة خصم العروض الخاصة"], errors="coerce").fillna(0) if "قيمة خصم العروض الخاصة" in df.columns else 0
    df["total_discount"] = pd.to_numeric(df["إجمالي الخصم"], errors="coerce").fillna(0) if "إجمالي الخصم" in df.columns else 0
    
    # استبعاد عملاء الهدية والدعاية
    df = df[~df["customer_name"].apply(is_gift_or_promotion)]
    
    # استبعاد الطلبات المحذوفة أو الفارغة
    df = df[
        (df["order_number"] != "")
        & (df["sku"] != "")
        & (df["quantity"] != 0)
        & (df["order_status"] != "محذوف")
    ].copy()
    
    # تحديد الفرع
    branch_info = df.apply(lambda row: determine_branch(row["order_status"], row["city"]), axis=1)
    df["pharmacy_name"] = branch_info.apply(lambda x: x[0])
    df["branch_number"] = branch_info.apply(lambda x: x[1])
    
    # تجميع البيانات
    grouped = df.groupby(["order_number", "sku"], as_index=False).agg({
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
        "discount": "first",
        "shipping_cost": "first",
        "payment_method": "first",
        "tax": "first",
        "coupon_discount": "first",
        "offer_discount": "first",
        "total_discount": "first"
    }).rename(columns={
        "product_name": "salla_product_name",
        "quantity": "salla_qty",
        "pharmacy_name": "salla_pharmacy_name",
        "branch_number": "salla_branch_number",
    })
    
    return grouped

def prepare_abc_frame(df_abc: pd.DataFrame) -> pd.DataFrame:
    """معالجة شيت ABC مع الأعمدة الجديدة"""
    df = df_abc.copy()
    
    # استبعاد FREE GIFTS
    EXCLUDED_PROFILE = "FREE GIFTS FOR CUSTOMERS"
    if "نوع البروفايل" in df.columns:
        df = df[df["نوع البروفايل"].astype(str).str.strip() != EXCLUDED_PROFILE].copy()
    
    # استبعاد DELIVERY FEE
    if "اسم الصنف" in df.columns:
        df = df[~df["اسم الصنف"].astype(str).str.upper().str.contains("DELIVERY FEE", na=False)].copy()
    
    # استبعاد SKU 16133
    if "رقم الصنف" in df.columns:
        df = df[df["رقم الصنف"].astype(str).str.strip() != "16133"].copy()
    
    # تحديد الأعمدة المطلوبة - ديناميكياً حسب المتاح
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
    
    # الأعمدة الجديدة من ABC
    df["batch_no"] = df["Batch No."].apply(normalize_text) if "Batch No." in df.columns else ""
    df["expiry_date"] = df["Expiry"].apply(normalize_text) if "Expiry" in df.columns else ""
    df["sale_price"] = pd.to_numeric(df["Sale Price"], errors="coerce").fillna(0) if "Sale Price" in df.columns else 0
    df["total_sale"] = pd.to_numeric(df["Total Sale"], errors="coerce").fillna(0) if "Total Sale" in df.columns else 0
    df["cost_price"] = pd.to_numeric(df["Cost Price"], errors="coerce").fillna(0) if "Cost Price" in df.columns else 0
    df["total_cost"] = pd.to_numeric(df["Total Cost"], errors="coerce").fillna(0) if "Total Cost" in df.columns else 0
    df["vat"] = pd.to_numeric(df["VAT %"], errors="coerce").fillna(0) if "VAT %" in df.columns else 0
    df["total_vat"] = pd.to_numeric(df["Total VAT."], errors="coerce").fillna(0) if "Total VAT." in df.columns else 0
    df["total_after_vat"] = pd.to_numeric(df["Total After VAT"], errors="coerce").fillna(0) if "Total After VAT" in df.columns else 0
    df["receipt_no"] = df["Receipt No."].apply(normalize_text) if "Receipt No." in df.columns else ""
    df["branch_city"] = df["Branch City"].apply(normalize_text) if "Branch City" in df.columns else ""
    
    if "Receipt Classification" in df.columns:
        df["receipt_classification"] = df["Receipt Classification"].apply(normalize_text)
    else:
        df["receipt_classification"] = ""
    
    # تصفية البيانات
    df = df[
        (df["sku"] != "")
        & (df["order_number"] != "")
    ].copy()
    
    # تجميع البيانات
    grouped = df.groupby(["order_number", "sku"], as_index=False).agg({
        "abc_qty": "sum",
        "invoice_number": "first",
        "invoice_date": "first",
        "abc_product_name": "first",
        "abc_pharmacy_name": "first",
        "abc_pharmacist_name": "first",
        "profile_type": lambda x: " | ".join(sorted({normalize_text(v) for v in x if normalize_text(v)})),
        "receipt_classification": lambda x: " | ".join(sorted({normalize_text(v) for v in x if normalize_text(v)})),
        "all_abc_pharmacies": lambda x: " | ".join(sorted({normalize_text(v) for v in x if normalize_text(v)})),
        "batch_no": "first",
        "expiry_date": "first",
        "sale_price": "mean",
        "total_sale": "sum",
        "cost_price": "mean",
        "total_cost": "sum",
        "vat": "mean",
        "total_vat": "sum",
        "total_after_vat": "sum",
        "receipt_no": "first",
        "branch_city": "first"
    })
    
    grouped["other_branch_details"] = grouped.apply(
        lambda row: f"تم بيع نفس الطلب/الصنف في فروع أخرى: {row['all_abc_pharmacies']}" if " | " in row["all_abc_pharmacies"] else "",
        axis=1
    )
    
    return grouped

def classify_cases(df_salla: pd.DataFrame, df_abc: pd.DataFrame) -> pd.DataFrame:
    """تصنيف الحالات (إضافة/إرجاع/طلب بدون فاتورة/فاتورة بدون طلب)"""
    salla_grouped = prepare_salla_frame(df_salla)
    abc_grouped = prepare_abc_frame(df_abc)
    merged = pd.merge(salla_grouped, abc_grouped, on=["order_number", "sku"], how="outer", indicator=True)
    
    # تعبئة القيم المفقودة
    for col in ["salla_qty", "abc_qty", "total_amount"]:
        if col in merged.columns:
            merged[col] = pd.to_numeric(merged[col], errors="coerce").fillna(0)
        else:
            merged[col] = 0
    
    # الأعمدة النصية
    text_cols = [
        "salla_product_name", "abc_product_name", "customer_name", "customer_phone", "city",
        "order_status", "order_date", "invoice_number", "invoice_date", "salla_pharmacy_name",
        "abc_pharmacy_name", "abc_pharmacist_name", "profile_type", "receipt_classification",
        "all_abc_pharmacies", "other_branch_details", "payment_method", "batch_no", "expiry_date",
        "receipt_no", "branch_city"
    ]
    for col in text_cols:
        if col not in merged.columns:
            merged[col] = ""
        merged[col] = merged[col].fillna("").astype(str)
    
    # تعبئة اسم المنتج
    merged["product_name"] = merged["salla_product_name"]
    merged.loc[merged["product_name"].eq(""), "product_name"] = merged.loc[merged["product_name"].eq(""), "abc_product_name"]
    
    # تعبئة اسم الصيدلية
    merged["pharmacy_name"] = merged["salla_pharmacy_name"]
    merged.loc[merged["pharmacy_name"].eq(""), "pharmacy_name"] = merged.loc[merged["pharmacy_name"].eq(""), "abc_pharmacy_name"]
    
    # تعبئة رقم الفرع
    merged["branch_number"] = merged["salla_branch_number"]
    merged.loc[merged["branch_number"].eq(""), "branch_number"] = merged.loc[merged["branch_number"].eq(""), "abc_branch_number"] if "abc_branch_number" in merged.columns else ""
    
    # حساب الفرق
    merged["difference"] = merged["salla_qty"] - merged["abc_qty"]
    merged["case_type"] = ""
    merged["case_reason"] = ""
    
    # تصنيف الحالات
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
    
    # الأعمدة المطلوبة للنتيجة
    ordered_columns = [
        "item_key", "order_number", "invoice_number", "sku", "product_name",
        "salla_product_name", "abc_product_name", "pharmacy_name",
        "salla_pharmacy_name", "abc_pharmacy_name", "abc_pharmacist_name",
        "branch_number", "salla_qty", "abc_qty", "difference",
        "case_type", "case_label", "case_reason", "customer_name",
        "customer_phone", "city", "order_status", "order_date",
        "invoice_date", "profile_type", "receipt_classification",
        "all_abc_pharmacies", "other_branch_details", "total_amount",
        # الأعمدة الجديدة
        "payment_method", "discount", "shipping_cost", "tax",
        "coupon_discount", "offer_discount", "total_discount",
        "batch_no", "expiry_date", "sale_price", "total_sale",
        "cost_price", "total_cost", "vat", "total_vat",
        "total_after_vat", "receipt_no", "branch_city"
    ]
    
    for col in ordered_columns:
        if col not in result.columns:
            result[col] = ""
    
    return result[ordered_columns]

def process_excel(uploaded_file, uploaded_by: str):
    """معالجة ملف Excel وإدراج النتائج في قاعدة البيانات"""
    df_salla = pd.read_excel(uploaded_file, sheet_name="سلة")
    df_abc = pd.read_excel(uploaded_file, sheet_name="abc")
    results = classify_cases(df_salla, df_abc)
    
    upload_batch_id = uuid.uuid4().hex
    timestamp = now_str()
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    
    # إدراج بيانات الرفع
    cur.execute("""
        INSERT INTO uploads (upload_batch_id, file_name, uploaded_by, uploaded_at, total_cases,
            total_additions, total_returns, total_orphan_salla, total_orphan_abc, is_active)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
    """, (upload_batch_id, uploaded_file.name, uploaded_by, timestamp, len(results),
          int((results["case_type"] == "addition").sum()), int((results["case_type"] == "return").sum()),
          int((results["case_type"] == "orphan_salla").sum()), int((results["case_type"] == "orphan_abc").sum())))
    
    # إدراج العناصر
    for _, row in results.iterrows():
        cur.execute("""
            INSERT OR REPLACE INTO reconciliation_items (
                item_key, upload_batch_id, order_number, invoice_number, sku, product_name,
                salla_product_name, abc_product_name, pharmacy_name, salla_pharmacy_name,
                abc_pharmacy_name, abc_pharmacist_name, branch_number, salla_qty, abc_qty, difference,
                case_type, case_label, case_reason, status, customer_name, customer_phone,
                city, order_status, order_date, invoice_date, profile_type, receipt_classification,
                all_abc_pharmacies, other_branch_details, pharmacist_note, total_amount, 
                payment_method, discount, shipping_cost, tax, coupon_discount, offer_discount,
                batch_no, expiry_date, sale_price, total_sale, cost_price, total_cost, vat, total_vat,
                total_after_vat, receipt_no, branch_city, first_seen_at, last_seen_at, active
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
        """, (
            row["item_key"], upload_batch_id, str(row["order_number"]), str(row.get("invoice_number", "")),
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
            "", float(row.get("total_amount", 0)),
            str(row.get("payment_method", "")), float(row.get("discount", 0)), float(row.get("shipping_cost", 0)),
            float(row.get("tax", 0)), float(row.get("coupon_discount", 0)), float(row.get("offer_discount", 0)),
            str(row.get("batch_no", "")), str(row.get("expiry_date", "")), float(row.get("sale_price", 0)),
            float(row.get("total_sale", 0)), float(row.get("cost_price", 0)), float(row.get("total_cost", 0)),
            float(row.get("vat", 0)), float(row.get("total_vat", 0)), float(row.get("total_after_vat", 0)),
            str(row.get("receipt_no", "")), str(row.get("branch_city", "")), timestamp, timestamp
        ))
    
    # تعطيل العناصر القديمة وتفعيل الجلسة الحالية
    cur.execute("UPDATE reconciliation_items SET active = CASE WHEN upload_batch_id = ? THEN 1 ELSE 0 END", (upload_batch_id,))
    cur.execute("UPDATE uploads SET is_active = 0")
    cur.execute("UPDATE uploads SET is_active = 1 WHERE upload_batch_id = ?", (upload_batch_id,))
    session_name = datetime.now().strftime("%Y-%m-%d %H:%M")
    cur.execute("UPDATE uploads SET session_name = ? WHERE upload_batch_id = ?", (session_name, upload_batch_id))
    
    conn.commit()
    conn.close()
    return results, upload_batch_id

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
