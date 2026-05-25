import pandas as pd
import numpy as np
import sqlite3
import uuid
import re
from datetime import datetime
from utils.helpers import (
    normalize_order_number, normalize_sku, normalize_text,
    determine_branch, get_branch_number, is_gift_or_promotion, now_str
)
from utils.database import DB_PATH

def find_column(df, possible_names):
    """البحث عن عمود في DataFrame بأسماء محتملة"""
    for name in possible_names:
        if name in df.columns:
            return name
        clean_name = str(name).strip()
        for col in df.columns:
            if str(col).strip() == clean_name:
                return col
    return None

def prepare_salla_frame(df_salla: pd.DataFrame) -> pd.DataFrame:
    """معالجة شيت سلة - تجميع كميات نفس SKU لنفس رقم الطلب"""
    df = df_salla.copy()
    
    # تحديد الأعمدة المطلوبة في شيت سلة
    order_col = find_column(df, ['رقم الطلب', 'Order Number', 'order_number'])
    sku_col = find_column(df, ['SKU', 'Sku', 'sku'])
    product_col = find_column(df, ['اسم المنتج', 'Product Name', 'product_name'])
    qty_col = find_column(df, ['الكمية', 'Quantity', 'qty'])
    customer_col = find_column(df, ['اسم العميل', 'Customer Name', 'customer_name'])
    phone_col = find_column(df, ['رقم الجوال', 'Phone', 'phone'])
    city_col = find_column(df, ['المدينة', 'City', 'city'])
    status_col = find_column(df, ['حالة الطلب', 'Order Status', 'order_status'])
    date_col = find_column(df, ['تاريخ الطلب', 'Order Date', 'order_date'])
    total_col = find_column(df, ['إجمالي الطلب', 'Total', 'total'])
    discount_col = find_column(df, ['الخصم', 'Discount', 'discount'])
    shipping_col = find_column(df, ['تكلفة الشحن', 'Shipping Cost', 'shipping_cost'])
    payment_col = find_column(df, ['طريقة الدفع', 'Payment Method', 'payment_method'])
    tax_col = find_column(df, ['الضريبة', 'Tax', 'tax'])
    coupon_col = find_column(df, ['قيمة خصم الكوبون', 'Coupon Discount', 'coupon_discount'])
    offer_col = find_column(df, ['قيمة خصم العروض الخاصة', 'Offer Discount', 'offer_discount'])
    
    # تطبيق الدوال
    df["order_number"] = df[order_col].apply(normalize_order_number) if order_col else ""
    df["sku"] = df[sku_col].apply(normalize_sku) if sku_col else ""
    df["product_name"] = df[product_col].apply(normalize_text) if product_col else ""
    df["quantity"] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0) if qty_col else 0
    df["customer_name"] = df[customer_col].apply(normalize_text) if customer_col else ""
    df["customer_phone"] = df[phone_col].apply(normalize_text) if phone_col else ""
    df["city"] = df[city_col].apply(normalize_text) if city_col else ""
    df["order_status"] = df[status_col].apply(normalize_text) if status_col else ""
    df["order_date"] = df[date_col].apply(normalize_text) if date_col else ""
    df["total_amount"] = pd.to_numeric(df[total_col], errors="coerce").fillna(0) if total_col else 0
    df["discount"] = pd.to_numeric(df[discount_col], errors="coerce").fillna(0) if discount_col else 0
    df["shipping_cost"] = pd.to_numeric(df[shipping_col], errors="coerce").fillna(0) if shipping_col else 0
    df["payment_method"] = df[payment_col].apply(normalize_text) if payment_col else ""
    df["tax"] = pd.to_numeric(df[tax_col], errors="coerce").fillna(0) if tax_col else 0
    df["coupon_discount"] = pd.to_numeric(df[coupon_col], errors="coerce").fillna(0) if coupon_col else 0
    df["offer_discount"] = pd.to_numeric(df[offer_col], errors="coerce").fillna(0) if offer_col else 0
    
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
    
    # تجميع البيانات حسب رقم الطلب و SKU
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
        "offer_discount": "first"
    }).rename(columns={
        "product_name": "salla_product_name",
        "quantity": "salla_qty",
        "pharmacy_name": "salla_pharmacy_name",
        "branch_number": "salla_branch_number",
    })
    
    return grouped


def prepare_abc_frame(df_abc: pd.DataFrame) -> pd.DataFrame:
    """معالجة شيت ABC - الحفاظ على كل فرع كسطر منفصل"""
    df = df_abc.copy()
    
    # إزالة صفوف الإجمالي (subtotal)
    df = df[~df.iloc[:, 0].astype(str).str.contains('SUBTOTAL', na=False, case=False)]
    
    # البحث عن الأعمدة بأسمائها أولاً
    order_col = find_column(df, ['رقم الطلب', 'Order Number', 'order_number'])
    sku_col = find_column(df, ['رقم الصنف', 'Item No.', 'Item Number', 'item_number'])
    product_col = find_column(df, ['اسم الصنف', 'Product', 'product'])
    qty_col = find_column(df, ['Net Sold Qty', 'Net Qty.', 'Net Sold Quantity'])
    invoice_col = find_column(df, ['رقم الفاتورة', 'Receipt No.', 'receipt_no', 'invoice_number'])
    date_col = find_column(df, ['التاريخ', 'Date', 'date', 'Sales Date'])
    pharmacy_col = find_column(df, ['رقم الصيدلية', 'Branch', 'branch'])
    pharmacist_col = find_column(df, ['الصيدلي', 'Username', 'username'])
    profile_col = find_column(df, ['نوع البروفايل', 'Profile', 'profile'])
    
    # إذا لم يتم العثور على الأعمدة بأسمائها، نستخدم المواقع
    if order_col is None and len(df.columns) > 30:
        order_col = df.columns[30]
        invoice_col = df.columns[28]
        date_col = df.columns[29]
        profile_col = df.columns[0]
        sku_col = df.columns[1]
        product_col = df.columns[2]
        qty_col = df.columns[9] if len(df.columns) > 9 else None
        pharmacy_col = df.columns[37] if len(df.columns) > 37 else None
        pharmacist_col = df.columns[44] if len(df.columns) > 44 else None
    
    if order_col is None:
        raise ValueError("لم يتم العثور على عمود رقم الطلب في شيت ABC")
    
    # تطبيق الدوال
    df["order_number"] = df[order_col].apply(normalize_order_number)
    df["sku"] = df[sku_col].apply(normalize_sku) if sku_col else ""
    df["abc_product_name"] = df[product_col].apply(normalize_text) if product_col else ""
    df["abc_qty"] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0) if qty_col else 0
    df["invoice_number"] = df[invoice_col].apply(normalize_text) if invoice_col else ""
    df["invoice_date"] = df[date_col].apply(normalize_text) if date_col else ""
    df["abc_pharmacy_name"] = df[pharmacy_col].apply(normalize_text) if pharmacy_col else ""
    df["abc_pharmacist_name"] = df[pharmacist_col].apply(normalize_text) if pharmacist_col else ""
    df["profile_type"] = df[profile_col].apply(normalize_text) if profile_col else ""
    
    df["all_abc_pharmacies"] = df["abc_pharmacy_name"]
    df["receipt_classification"] = ""
    
    # استبعاد FREE GIFTS
    EXCLUDED_PROFILE = "FREE GIFTS FOR CUSTOMERS"
    df = df[df["profile_type"] != EXCLUDED_PROFILE].copy() if "profile_type" in df.columns else df
    
    # استبعاد DELIVERY FEE
    df = df[~df["abc_product_name"].str.upper().str.contains("DELIVERY FEE", na=False)] if "abc_product_name" in df.columns else df
    
    # استبعاد SKU غير الصالحة
    df = df[~df["sku"].isin(["", "0", "1", "200", "16133"])].copy()
    
    # استبعاد البيانات غير الصالحة
    df = df[
        (df["sku"] != "")
        & (df["order_number"] != "")
    ].copy()
    
    if df.empty:
        return pd.DataFrame()
    
    # تجميع البيانات حسب رقم الطلب، SKU، والفرع (كل فرع على حدة)
    grouped = df.groupby(["order_number", "sku", "abc_pharmacy_name"], as_index=False).agg({
        "abc_qty": "sum",
        "invoice_number": lambda x: " | ".join(sorted(set(str(v) for v in x if v))),
        "invoice_date": "first",
        "abc_product_name": "first",
        "abc_pharmacist_name": "first",
        "profile_type": lambda x: " | ".join(sorted({normalize_text(v) for v in x if normalize_text(v)})),
        "receipt_classification": lambda x: " | ".join(sorted({normalize_text(v) for v in x if normalize_text(v)})),
        "all_abc_pharmacies": lambda x: " | ".join(sorted({normalize_text(v) for v in x if normalize_text(v)}))
    })
    
    return grouped


def classify_cases(df_salla: pd.DataFrame, df_abc: pd.DataFrame) -> pd.DataFrame:
    """
    تصنيف الحالات حسب المنطق المطلوب
    """
    salla_grouped = prepare_salla_frame(df_salla)
    abc_grouped = prepare_abc_frame(df_abc)
    
    if salla_grouped.empty and abc_grouped.empty:
        return pd.DataFrame()
    
    # حساب المجموع الكلي لكميات ABC لكل order_number و sku
    abc_total_qty = abc_grouped.groupby(["order_number", "sku"])["abc_qty"].sum().reset_index()
    abc_total_qty.rename(columns={"abc_qty": "abc_total_qty"}, inplace=True)
    
    # دمج سلة مع المجموع الكلي
    salla_with_total = pd.merge(salla_grouped, abc_total_qty, on=["order_number", "sku"], how="left")
    salla_with_total["abc_total_qty"] = salla_with_total["abc_total_qty"].fillna(0)
    
    # تحديد ما إذا كان المجموع الكلي مطابقاً لكمية سلة
    salla_with_total["total_matched"] = salla_with_total["salla_qty"] == salla_with_total["abc_total_qty"]
    
    # دمج مع بيانات ABC لكل فرع على حدة
    merged = pd.merge(
        salla_with_total, 
        abc_grouped, 
        on=["order_number", "sku"], 
        how="outer", 
        indicator=True
    )
    
    # تعبئة القيم المفقودة
    for col in ["salla_qty", "abc_qty", "total_amount", "abc_total_qty"]:
        if col in merged.columns:
            merged[col] = pd.to_numeric(merged[col], errors="coerce").fillna(0)
        else:
            merged[col] = 0
    
    # تعبئة عمود total_matched
    if "total_matched" not in merged.columns:
        merged["total_matched"] = False
    merged["total_matched"] = merged["total_matched"].fillna(False)
    
    # الأعمدة النصية
    text_cols = [
        "salla_product_name", "abc_product_name", "customer_name", "customer_phone", "city",
        "order_status", "order_date", "invoice_number", "invoice_date", "salla_pharmacy_name",
        "abc_pharmacy_name", "abc_pharmacist_name", "profile_type", "receipt_classification",
        "all_abc_pharmacies", "other_branch_details", "payment_method"
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
    abc_mask = (merged["abc_pharmacy_name"] != "") & (merged["abc_pharmacy_name"] != "nan")
    merged.loc[abc_mask, "pharmacy_name"] = merged.loc[abc_mask, "abc_pharmacy_name"]
    empty_mask = (merged["pharmacy_name"] == "") | (merged["pharmacy_name"] == "nan")
    merged.loc[empty_mask, "pharmacy_name"] = merged.loc[empty_mask, "salla_pharmacy_name"]
    
    # حساب الفرق
    merged["difference"] = merged["salla_qty"] - merged["abc_qty"]
    
    # ========== التصنيف ==========
    merged["case_type"] = ""
    merged["case_reason"] = ""
    
    # 1. فاتورة بدون طلب
    orphan_abc_mask = (merged["_merge"] == "right_only") & (merged["abc_qty"] != 0)
    merged.loc[orphan_abc_mask, "case_type"] = "orphan_abc"
    merged.loc[orphan_abc_mask, "case_reason"] = f"📄 فاتورة موجودة في ABC بكمية {merged.loc[orphan_abc_mask, 'abc_qty']} ولم يُعثر عليها في سلة."
    
    # 2. طلب بدون فاتورة
    orphan_salla_mask = (merged["_merge"] == "left_only") & (merged["salla_qty"] > 0)
    merged.loc[orphan_salla_mask, "case_type"] = "orphan_salla"
    merged.loc[orphan_salla_mask, "case_reason"] = f"🛒 طلب موجود في سلة بكمية {merged.loc[orphan_salla_mask, 'salla_qty']} ولم يُعثر عليه في ABC."
    
    # 3. العنصر موجود في كليهما
    both_mask = (merged["_merge"] == "both")
    
    # 3a. توزيع على عدة فروع (المجموع مطابق)
    multiple_branches_mask = both_mask & (merged["total_matched"] == True) & (merged["abc_qty"] > 0)
    merged.loc[multiple_branches_mask, "case_type"] = "addition"
    merged.loc[multiple_branches_mask, "case_reason"] = (
        f"⚠️ تنبيه: الكمية الإجمالية في ABC ({merged.loc[multiple_branches_mask, 'abc_total_qty']}) "
        f"مطابقة لكمية سلة ({merged.loc[multiple_branches_mask, 'salla_qty']})، ولكنها موزعة على عدة فروع. "
        f"هذا الفرع تظهر فيه كمية {merged.loc[multiple_branches_mask, 'abc_qty']}. "
        f"يرجى المراجعة والتأكد من أي فرع صحيح تم خروج الصنف منه حقيقة."
    )
    
    # 3b. إضافة عادية (كمية سلة أكبر)
    addition_mask = (
        both_mask & 
        (merged["total_matched"] == False) & 
        (merged["salla_qty"] > merged["abc_qty"]) &
        (merged["abc_qty"] > 0)
    )
    merged.loc[addition_mask, "case_type"] = "addition"
    merged.loc[addition_mask, "case_reason"] = (
        f"➕ كمية الطلب في سلة ({merged.loc[addition_mask, 'salla_qty']}) "
        f"أكبر من كمية الفاتورة في هذا الفرع ({merged.loc[addition_mask, 'abc_qty']}). "
        f"(إجمالي ABC: {merged.loc[addition_mask, 'abc_total_qty']})"
    )
    
    # 3c. فرع بدون كمية (abc_qty = 0)
    zero_in_branch_mask = (
        both_mask & 
        (merged["total_matched"] == False) & 
        (merged["salla_qty"] > 0) &
        (merged["abc_qty"] == 0)
    )
    merged.loc[zero_in_branch_mask, "case_type"] = "addition"
    merged.loc[zero_in_branch_mask, "case_reason"] = (
        f"⚠️ تنبيه: الطلب موجود في سلة بكمية {merged.loc[zero_in_branch_mask, 'salla_qty']} "
        f"ولكن هذا الفرع لا تظهر فيه أي كمية في ABC. "
        f"(إجمالي ABC في جميع الفروع: {merged.loc[zero_in_branch_mask, 'abc_total_qty']})"
    )
    
    # 3d. إرجاع (كمية ABC أكبر)
    return_mask = (
        both_mask & 
        (merged["abc_qty"] > merged["salla_qty"])
    )
    merged.loc[return_mask, "case_type"] = "return"
    merged.loc[return_mask, "case_reason"] = (
        f"🔄 كمية الفاتورة في هذا الفرع ({merged.loc[return_mask, 'abc_qty']}) "
        f"أكبر من كمية الطلب في سلة ({merged.loc[return_mask, 'salla_qty']})."
    )
    
    # تصفية النتيجة
    result = merged[merged["case_type"] != ""].copy()
    
    if result.empty:
        return pd.DataFrame()
    
    result["case_label"] = result["case_type"]
    
    # إنشاء item_key فريد
    result["item_key"] = result.apply(
        lambda r: f"{r['pharmacy_name']}||{r['order_number']}||{r['sku']}||{r['case_type']}", 
        axis=1
    )
    
    # الأعمدة المطلوبة (بدون duplicate_warning)
    ordered_columns = [
        "item_key", "order_number", "invoice_number", "sku", "product_name",
        "salla_product_name", "abc_product_name", "pharmacy_name",
        "salla_pharmacy_name", "abc_pharmacy_name", "abc_pharmacist_name",
        "salla_qty", "abc_qty", "difference",
        "case_type", "case_label", "case_reason", "customer_name",
        "customer_phone", "city", "order_status", "order_date",
        "invoice_date", "profile_type", "receipt_classification",
        "all_abc_pharmacies", "other_branch_details", "total_amount",
        "payment_method", "discount", "shipping_cost", "tax",
        "coupon_discount", "offer_discount"
    ]
    
    for col in ordered_columns:
        if col not in result.columns:
            result[col] = ""
    
    result = result[ordered_columns]
    
    return result


def process_excel(uploaded_file, uploaded_by: str):
    """معالجة ملف Excel وإدراج النتائج في قاعدة البيانات"""
    df_salla = pd.read_excel(uploaded_file, sheet_name="سلة")
    df_abc = pd.read_excel(uploaded_file, sheet_name="abc")
    results = classify_cases(df_salla, df_abc)
    
    upload_batch_id = uuid.uuid4().hex
    timestamp = now_str()
    
    conn = sqlite3.connect(DB_PATH, timeout=60.0)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    cur = conn.cursor()
    
    try:
        # إدراج بيانات الرفع
        cur.execute("""
            INSERT INTO uploads (upload_batch_id, file_name, uploaded_by, uploaded_at, total_cases,
                total_additions, total_returns, total_orphan_salla, total_orphan_abc, is_active)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
        """, (upload_batch_id, uploaded_file.name, uploaded_by, timestamp, len(results),
              int((results["case_type"] == "addition").sum()) if not results.empty else 0,
              int((results["case_type"] == "return").sum()) if not results.empty else 0,
              int((results["case_type"] == "orphan_salla").sum()) if not results.empty else 0,
              int((results["case_type"] == "orphan_abc").sum()) if not results.empty else 0))
        
        if not results.empty:
            # إعداد بيانات الإدراج
            insert_df = results.copy()
            insert_df['upload_batch_id'] = upload_batch_id
            insert_df['status'] = 'قيد المتابعة'
            insert_df['pharmacist_note'] = ''
            insert_df['first_seen_at'] = timestamp
            insert_df['last_seen_at'] = timestamp
            insert_df['active'] = 1
            insert_df['hidden_from_pharmacy'] = 0
            insert_df['is_item_locked'] = 0
            insert_df['item_locked_by'] = ''
            insert_df['item_locked_at'] = ''
            insert_df['performed_by'] = ''
            insert_df['performed_at'] = ''
            
            # قائمة الأعمدة الموجودة في قاعدة البيانات
            valid_columns = [
                'item_key', 'upload_batch_id', 'order_number', 'invoice_number', 'sku',
                'product_name', 'salla_product_name', 'abc_product_name', 'pharmacy_name',
                'salla_pharmacy_name', 'abc_pharmacy_name', 'abc_pharmacist_name',
                'branch_number', 'salla_branch_number', 'salla_qty', 'abc_qty', 'difference',
                'case_type', 'case_label', 'case_reason', 'status', 'performed_by', 'performed_at',
                'customer_name', 'customer_phone', 'city', 'order_status', 'order_date',
                'invoice_date', 'profile_type', 'receipt_classification', 'all_abc_pharmacies',
                'other_branch_details', 'pharmacist_note', 'total_amount', 'first_seen_at',
                'last_seen_at', 'active', 'hidden_from_pharmacy', 'payment_method',
                'discount', 'shipping_cost', 'tax', 'coupon_discount', 'offer_discount',
                'is_item_locked', 'item_locked_by', 'item_locked_at'
            ]
            
            # إزالة الأعمدة غير الموجودة
            cols_to_drop = [col for col in insert_df.columns if col not in valid_columns]
            if cols_to_drop:
                insert_df = insert_df.drop(columns=cols_to_drop)
            
            # إضافة الأعمدة المفقودة
            for col in valid_columns:
                if col not in insert_df.columns:
                    if col in ['performed_by', 'performed_at', 'item_locked_by', 'item_locked_at', 'branch_number', 'salla_branch_number']:
                        insert_df[col] = ''
                    elif col in ['is_item_locked']:
                        insert_df[col] = 0
                    else:
                        insert_df[col] = ''
            
            # ترتيب الأعمدة
            insert_df = insert_df[valid_columns]
            
            # ========== الإدراج على دفعات (batches) لتجنب "too many SQL variables" ==========
            batch_size = 500  # عدد الصفوف في كل دفعة
            total_rows = len(insert_df)
            
            for start in range(0, total_rows, batch_size):
                end = min(start + batch_size, total_rows)
                batch_df = insert_df.iloc[start:end]
                batch_df.to_sql('reconciliation_items', conn, if_exists='append', index=False, method='multi')
                print(f"Inserted rows {start} to {end} of {total_rows}")
        
        # تعطيل العناصر القديمة وتفعيل الجلسة الحالية
        cur.execute("UPDATE reconciliation_items SET active = CASE WHEN upload_batch_id = ? THEN 1 ELSE 0 END", (upload_batch_id,))
        cur.execute("UPDATE uploads SET is_active = 0")
        cur.execute("UPDATE uploads SET is_active = 1 WHERE upload_batch_id = ?", (upload_batch_id,))
        session_name = datetime.now().strftime("%Y-%m-%d %H:%M")
        cur.execute("UPDATE uploads SET session_name = ? WHERE upload_batch_id = ?", (session_name, upload_batch_id))
        
        conn.commit()
        
    except Exception as e:
        print(f"Error in process_excel: {e}")
        conn.rollback()
        raise
    finally:
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
