import os
import re
import sqlite3
import uuid
from datetime import datetime

import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="نظام بلسم - مطابقة الطلبات والفواتير",
    layout="wide",
    initial_sidebar_state="expanded",
)


DB_DIR = "data"
DB_PATH = os.path.join(DB_DIR, "pharmacy_reconciliation.db")
PHARMACY_COUNT = 17
SPECIAL_ORDER_NUMBERS = {"0", "123456"}
EXCLUDED_PROFILE = "FREE GIFTS FOR CUSTOMERS"
CASE_LABELS = {
    "addition": "إضافة",
    "return": "إرجاع",
    "orphan_salla": "طلب بدون فاتورة",
    "orphan_abc": "فاتورة بدون طلب",
    "branch_mismatch": "اختلاف فرع",
    "special_review": "مراجعة رقم طلب خاص",
}
STATUS_DONE = "تم"
STATUS_PENDING = "قيد المتابعة"


st.markdown(
    """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@300;400;500;700;800&display=swap');

    * { font-family: 'Tajawal', sans-serif; }

    .hero {
        background:
            radial-gradient(circle at top right, rgba(255,255,255,0.18), transparent 28%),
            linear-gradient(135deg, #0f4c5c 0%, #1f7a8c 50%, #16425b 100%);
        border-radius: 24px;
        padding: 2.2rem;
        color: white;
        margin-bottom: 1.6rem;
        box-shadow: 0 18px 40px rgba(22, 66, 91, 0.20);
    }

    .hero h1 {
        margin: 0;
        font-size: 2.2rem;
        font-weight: 800;
    }

    .hero p {
        margin-top: 0.6rem;
        font-size: 1rem;
        opacity: 0.95;
    }

    .section-title {
        font-size: 1.15rem;
        font-weight: 800;
        color: #16425b;
        border-right: 5px solid #1f7a8c;
        padding-right: 0.65rem;
        margin: 1rem 0 0.8rem;
    }

    .note-card {
        background: linear-gradient(135deg, #f4fbfc 0%, #ffffff 100%);
        border: 1px solid #d7ebef;
        border-radius: 16px;
        padding: 1rem 1.1rem;
        margin-bottom: 1rem;
    }

    .action-card {
        background: white;
        border: 1px solid #e4eef1;
        border-right: 6px solid #1f7a8c;
        border-radius: 18px;
        padding: 1rem;
        margin-bottom: 0.85rem;
        box-shadow: 0 8px 22px rgba(15, 76, 92, 0.07);
    }

    .pill {
        display: inline-block;
        padding: 0.28rem 0.75rem;
        border-radius: 999px;
        font-size: 0.78rem;
        font-weight: 700;
    }

    .pill-green { background: #dff7e8; color: #0f7a3a; }
    .pill-amber { background: #fff0c2; color: #8a5b00; }
    .pill-red { background: #ffe0df; color: #a32929; }
    .pill-blue { background: #dff1ff; color: #0f5488; }
    .pill-slate { background: #eef3f5; color: #445b66; }

    .metric-box {
        background: white;
        border-radius: 18px;
        padding: 1rem;
        border: 1px solid #e6eef0;
        box-shadow: 0 8px 20px rgba(15, 76, 92, 0.06);
    }

    .stButton button {
        width: 100%;
        border-radius: 10px;
        font-weight: 800;
    }
</style>
""",
    unsafe_allow_html=True,
)


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def pharmacy_names():
    return [f"Balsam Alula Pharmacy {i:02d}" for i in range(1, PHARMACY_COUNT + 1)]


def ensure_database():
    os.makedirs(DB_DIR, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY,
            password TEXT NOT NULL,
            role TEXT NOT NULL,
            pharmacist_name TEXT DEFAULT '',
            last_login TEXT DEFAULT ''
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS last_access (
            pharmacy_name TEXT PRIMARY KEY,
            last_login TEXT DEFAULT '',
            pharmacist_name TEXT DEFAULT ''
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS uploads (
            upload_batch_id TEXT PRIMARY KEY,
            file_name TEXT,
            uploaded_by TEXT,
            uploaded_at TEXT,
            total_cases INTEGER DEFAULT 0,
            total_additions INTEGER DEFAULT 0,
            total_returns INTEGER DEFAULT 0,
            total_orphan_salla INTEGER DEFAULT 0,
            total_orphan_abc INTEGER DEFAULT 0,
            total_branch_mismatch INTEGER DEFAULT 0,
            total_special_review INTEGER DEFAULT 0
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS reconciliation_items (
            item_key TEXT PRIMARY KEY,
            upload_batch_id TEXT,
            order_number TEXT,
            invoice_number TEXT,
            sku TEXT,
            product_name TEXT,
            salla_product_name TEXT,
            abc_product_name TEXT,
            pharmacy_name TEXT,
            salla_pharmacy_name TEXT,
            abc_pharmacy_name TEXT,
            branch_number TEXT,
            salla_qty REAL DEFAULT 0,
            abc_qty REAL DEFAULT 0,
            difference REAL DEFAULT 0,
            case_type TEXT,
            case_label TEXT,
            case_reason TEXT,
            status TEXT DEFAULT 'قيد المتابعة',
            performed_by TEXT DEFAULT '',
            performed_at TEXT DEFAULT '',
            customer_name TEXT DEFAULT '',
            customer_phone TEXT DEFAULT '',
            city TEXT DEFAULT '',
            order_status TEXT DEFAULT '',
            order_date TEXT DEFAULT '',
            invoice_date TEXT DEFAULT '',
            total_amount REAL DEFAULT 0,
            first_seen_at TEXT,
            last_seen_at TEXT,
            active INTEGER DEFAULT 1
        )
        """
    )

    cur.execute(
        """
        CREATE INDEX IF NOT EXISTS idx_reconciliation_active_pharmacy
        ON reconciliation_items (active, pharmacy_name, case_type)
        """
    )

    for index, name in enumerate(pharmacy_names(), start=1):
        cur.execute(
            """
            INSERT OR IGNORE INTO users (username, password, role, pharmacist_name, last_login)
            VALUES (?, ?, 'pharmacy', '', '')
            """,
            (name, f"balsam{index}"),
        )

    cur.execute(
        """
        INSERT OR IGNORE INTO users (username, password, role, pharmacist_name, last_login)
        VALUES ('admin', 'admin123', 'admin', 'Manager', '')
        """
    )

    conn.commit()
    conn.close()


def normalize_text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_city(value) -> str:
    city = normalize_text(value).upper()
    city = city.replace("-", " ").replace("_", " ")
    city = re.sub(r"\s+", " ", city)
    return city


def normalize_order_number(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    text = text.replace(".0", "")
    if text.lower() in {"nan", "none", "null"}:
        return ""
    match = re.search(r"\d+", text)
    return match.group(0) if match else text


def normalize_sku(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if not text:
        return ""
    text = text.replace(".0", "")
    if re.fullmatch(r"\d+", text):
        return text
    return ""


def numeric_value(value) -> float:
    return float(pd.to_numeric(pd.Series([value]), errors="coerce").fillna(0).iloc[0])


def extract_branch_from_status(status_text):
    if not status_text or pd.isna(status_text):
        return ""
    match = re.search(r"فرع\s*(\d+)", str(status_text))
    if match:
        return f"{int(match.group(1)):02d}"
    return ""


def determine_branch(order_status: str, city: str) -> tuple[str, str]:
    branch_num = extract_branch_from_status(order_status)
    if branch_num:
        return f"Balsam Alula Pharmacy {branch_num}", branch_num

    normalized_city = normalize_city(city)
    delivered_statuses = ["تم التوصيل", "ملغي", "مسترجع", "محذوف"]
    if any(status in normalize_text(order_status) for status in delivered_statuses):
        if normalized_city in {"AL ULA", "ALULA", "AL-ULA"}:
            return "Balsam Alula Pharmacy 09", "09"
        return "Balsam Alula Pharmacy 13", "13"

    return "Balsam Alula Pharmacy 13", "13"


def get_branch_number(pharmacy_name: str) -> str:
    match = re.search(r"(\d{2})$", normalize_text(pharmacy_name))
    return match.group(1) if match else ""


def update_last_access(pharmacy_name: str, pharmacist_name: str):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    current_time = now_str()
    cur.execute(
        "UPDATE users SET pharmacist_name = ?, last_login = ? WHERE username = ?",
        (pharmacist_name, current_time, pharmacy_name),
    )
    cur.execute(
        """
        INSERT INTO last_access (pharmacy_name, last_login, pharmacist_name)
        VALUES (?, ?, ?)
        ON CONFLICT(pharmacy_name) DO UPDATE SET
            last_login = excluded.last_login,
            pharmacist_name = excluded.pharmacist_name
        """,
        (pharmacy_name, current_time, pharmacist_name),
    )
    conn.commit()
    conn.close()


def fetch_user(username: str, password: str):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute(
        "SELECT username, role, pharmacist_name FROM users WHERE username = ? AND password = ?",
        (username, password),
    )
    user = cur.fetchone()
    conn.close()
    return user


def get_all_last_logins() -> pd.DataFrame:
    conn = sqlite3.connect(DB_PATH)
    try:
        return pd.read_sql_query(
            """
            SELECT pharmacy_name, last_login, pharmacist_name
            FROM last_access
            ORDER BY last_login DESC
            """,
            conn,
        )
    finally:
        conn.close()


def get_latest_upload_summary():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute(
        """
        SELECT upload_batch_id, file_name, uploaded_by, uploaded_at, total_cases,
               total_additions, total_returns, total_orphan_salla, total_orphan_abc,
               total_branch_mismatch, total_special_review
        FROM uploads
        ORDER BY uploaded_at DESC
        LIMIT 1
        """
    )
    row = cur.fetchone()
    conn.close()
    return row


def build_item_key(row: pd.Series) -> str:
    parts = [
        normalize_text(row.get("pharmacy_name")),
        normalize_text(row.get("order_number")),
        normalize_text(row.get("sku")),
        normalize_text(row.get("case_type")),
    ]
    return "||".join(parts)


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

    df = df[(df["order_number"] != "") & (df["sku"] != "") & (df["quantity"] != 0)].copy()

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
    if "نوع البروفايل" in df.columns:
        df = df[df["نوع البروفايل"].astype(str).str.strip() != EXCLUDED_PROFILE].copy()

    df["order_number"] = df["رقم الطلب"].apply(normalize_order_number)
    df["sku"] = df["رقم الصنف"].apply(normalize_sku)
    df["abc_product_name"] = df["اسم الصنف"].apply(normalize_text)
    df["abc_qty"] = pd.to_numeric(df["Net Sold Qty"], errors="coerce").fillna(0)
    df["invoice_number"] = df["رقم الفاتورة"].apply(normalize_text)
    df["invoice_date"] = df["التاريخ"].apply(normalize_text)
    df["abc_pharmacy_name"] = df["رقم الصيدلية"].apply(normalize_text)
    df["abc_branch_number"] = df["abc_pharmacy_name"].apply(get_branch_number)

    df = df[(df["sku"] != "") & (df["order_number"] != "")].copy()

    grouped = (
        df.groupby(["order_number", "sku"], as_index=False)
        .agg(
            {
                "abc_qty": "sum",
                "invoice_number": "first",
                "invoice_date": "first",
                "abc_product_name": "first",
                "abc_pharmacy_name": "first",
                "abc_branch_number": "first",
            }
        )
    )
    return grouped


def classify_cases(df_salla: pd.DataFrame, df_abc: pd.DataFrame) -> pd.DataFrame:
    salla_grouped = prepare_salla_frame(df_salla)
    abc_grouped = prepare_abc_frame(df_abc)

    merged = pd.merge(
        salla_grouped,
        abc_grouped,
        on=["order_number", "sku"],
        how="outer",
        indicator=True,
    )

    for column in [
        "salla_qty",
        "abc_qty",
        "total_amount",
    ]:
        if column in merged.columns:
            merged[column] = pd.to_numeric(merged[column], errors="coerce").fillna(0)

    text_columns = [
        "salla_product_name",
        "abc_product_name",
        "customer_name",
        "customer_phone",
        "city",
        "order_status",
        "order_date",
        "invoice_number",
        "invoice_date",
        "salla_pharmacy_name",
        "abc_pharmacy_name",
        "salla_branch_number",
        "abc_branch_number",
    ]
    for column in text_columns:
        if column not in merged.columns:
            merged[column] = ""
        merged[column] = merged[column].fillna("").astype(str)

    merged["product_name"] = merged["salla_product_name"]
    missing_product = merged["product_name"].eq("")
    merged.loc[missing_product, "product_name"] = merged.loc[missing_product, "abc_product_name"]

    merged["pharmacy_name"] = merged["salla_pharmacy_name"]
    missing_pharmacy = merged["pharmacy_name"].fillna("").eq("")
    merged.loc[missing_pharmacy, "pharmacy_name"] = merged.loc[missing_pharmacy, "abc_pharmacy_name"]

    merged["branch_number"] = merged["salla_branch_number"]
    missing_branch = merged["branch_number"].fillna("").eq("")
    merged.loc[missing_branch, "branch_number"] = merged.loc[missing_branch, "abc_branch_number"]

    merged["difference"] = merged["salla_qty"] - merged["abc_qty"]
    merged["case_type"] = ""
    merged["case_reason"] = ""

    special_mask = merged["order_number"].isin(SPECIAL_ORDER_NUMBERS)
    merged.loc[special_mask, "case_type"] = "special_review"
    merged.loc[special_mask, "case_reason"] = "رقم طلب خاص لا يمكن الاعتماد عليه في الترحيل التلقائي."

    addition_mask = (
        (merged["case_type"] == "")
        & (merged["_merge"] == "both")
        & (merged["salla_qty"] > merged["abc_qty"])
        & (merged["salla_qty"] > 0)
    )
    merged.loc[addition_mask, "case_type"] = "addition"
    merged.loc[addition_mask, "case_reason"] = "كمية الطلب أعلى من كمية الفاتورة."

    return_mask = (
        (merged["case_type"] == "")
        & (merged["_merge"] == "both")
        & (merged["abc_qty"] > merged["salla_qty"])
    )
    merged.loc[return_mask, "case_type"] = "return"
    merged.loc[return_mask, "case_reason"] = "كمية الفاتورة أعلى من كمية الطلب."

    orphan_salla_mask = (
        (merged["case_type"] == "")
        & (merged["_merge"] == "left_only")
        & (merged["salla_qty"] > 0)
    )
    merged.loc[orphan_salla_mask, "case_type"] = "orphan_salla"
    merged.loc[orphan_salla_mask, "case_reason"] = "سطر طلب موجود في سلة ولم يُعثر على سطر مطابق له في ABC."

    orphan_abc_mask = (
        (merged["case_type"] == "")
        & (merged["_merge"] == "right_only")
        & (merged["abc_qty"] != 0)
    )
    merged.loc[orphan_abc_mask, "case_type"] = "orphan_abc"
    merged.loc[orphan_abc_mask, "case_reason"] = "سطر فاتورة موجود في ABC ولم يُعثر على سطر مطابق له في سلة."

    result = merged[merged["case_type"] != ""].copy()
    result["case_label"] = result["case_type"].map(CASE_LABELS)
    result["item_key"] = result.apply(build_item_key, axis=1)
    result["status"] = STATUS_PENDING
    result["performed_by"] = ""
    result["performed_at"] = ""

    ordered_columns = [
        "item_key",
        "order_number",
        "invoice_number",
        "sku",
        "product_name",
        "salla_product_name",
        "abc_product_name",
        "pharmacy_name",
        "salla_pharmacy_name",
        "abc_pharmacy_name",
        "branch_number",
        "salla_qty",
        "abc_qty",
        "difference",
        "case_type",
        "case_label",
        "case_reason",
        "customer_name",
        "customer_phone",
        "city",
        "order_status",
        "order_date",
        "invoice_date",
        "total_amount",
    ]
    return result[ordered_columns]


def persist_reconciliation_results(results: pd.DataFrame, uploaded_file_name: str, uploaded_by: str):
    upload_batch_id = uuid.uuid4().hex
    timestamp = now_str()
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    cur.execute("BEGIN")

    cur.execute(
        """
        INSERT INTO uploads (
            upload_batch_id, file_name, uploaded_by, uploaded_at, total_cases,
            total_additions, total_returns, total_orphan_salla, total_orphan_abc,
            total_branch_mismatch, total_special_review
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            upload_batch_id,
            uploaded_file_name,
            uploaded_by,
            timestamp,
            len(results),
            int((results["case_type"] == "addition").sum()),
            int((results["case_type"] == "return").sum()),
            int((results["case_type"] == "orphan_salla").sum()),
            int((results["case_type"] == "orphan_abc").sum()),
            int((results["case_type"] == "branch_mismatch").sum()),
            int((results["case_type"] == "special_review").sum()),
        ),
    )

    existing_map = {}
    existing_rows = cur.execute(
        "SELECT item_key, status, performed_by, performed_at, first_seen_at FROM reconciliation_items"
    ).fetchall()
    for item_key, status, performed_by, performed_at, first_seen_at in existing_rows:
        existing_map[item_key] = {
            "status": status,
            "performed_by": performed_by,
            "performed_at": performed_at,
            "first_seen_at": first_seen_at or timestamp,
        }

    for _, row in results.iterrows():
        previous = existing_map.get(row["item_key"], {})
        cur.execute(
            """
            INSERT INTO reconciliation_items (
                item_key, upload_batch_id, order_number, invoice_number, sku, product_name,
                salla_product_name, abc_product_name, pharmacy_name, salla_pharmacy_name,
                abc_pharmacy_name, branch_number, salla_qty, abc_qty, difference,
                case_type, case_label, case_reason, status, performed_by, performed_at,
                customer_name, customer_phone, city, order_status, order_date,
                invoice_date, total_amount, first_seen_at, last_seen_at, active
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
            ON CONFLICT(item_key) DO UPDATE SET
                upload_batch_id = excluded.upload_batch_id,
                order_number = excluded.order_number,
                invoice_number = excluded.invoice_number,
                sku = excluded.sku,
                product_name = excluded.product_name,
                salla_product_name = excluded.salla_product_name,
                abc_product_name = excluded.abc_product_name,
                pharmacy_name = excluded.pharmacy_name,
                salla_pharmacy_name = excluded.salla_pharmacy_name,
                abc_pharmacy_name = excluded.abc_pharmacy_name,
                branch_number = excluded.branch_number,
                salla_qty = excluded.salla_qty,
                abc_qty = excluded.abc_qty,
                difference = excluded.difference,
                case_type = excluded.case_type,
                case_label = excluded.case_label,
                case_reason = excluded.case_reason,
                customer_name = excluded.customer_name,
                customer_phone = excluded.customer_phone,
                city = excluded.city,
                order_status = excluded.order_status,
                order_date = excluded.order_date,
                invoice_date = excluded.invoice_date,
                total_amount = excluded.total_amount,
                first_seen_at = reconciliation_items.first_seen_at,
                last_seen_at = excluded.last_seen_at,
                active = 1
            """,
            (
                row["item_key"],
                upload_batch_id,
                row["order_number"],
                row["invoice_number"],
                row["sku"],
                row["product_name"],
                row["salla_product_name"],
                row["abc_product_name"],
                row["pharmacy_name"],
                row["salla_pharmacy_name"],
                row["abc_pharmacy_name"],
                row["branch_number"],
                numeric_value(row["salla_qty"]),
                numeric_value(row["abc_qty"]),
                numeric_value(row["difference"]),
                row["case_type"],
                row["case_label"],
                row["case_reason"],
                previous.get("status", STATUS_PENDING),
                previous.get("performed_by", ""),
                previous.get("performed_at", ""),
                row["customer_name"],
                row["customer_phone"],
                row["city"],
                row["order_status"],
                row["order_date"],
                row["invoice_date"],
                numeric_value(row["total_amount"]),
                previous.get("first_seen_at", timestamp),
                timestamp,
            ),
        )

    cur.execute(
        """
        UPDATE reconciliation_items
        SET active = CASE WHEN upload_batch_id = ? THEN 1 ELSE 0 END
        """,
        (upload_batch_id,),
    )

    conn.commit()
    conn.close()
    return upload_batch_id


def process_excel(uploaded_file, uploaded_by: str):
    df_salla = pd.read_excel(uploaded_file, sheet_name="سلة")
    df_abc = pd.read_excel(uploaded_file, sheet_name="abc")
    results = classify_cases(df_salla, df_abc)
    upload_batch_id = persist_reconciliation_results(results, uploaded_file.name, uploaded_by)
    return results, upload_batch_id


def fetch_active_items(pharmacy_name: str | None = None) -> pd.DataFrame:
    conn = sqlite3.connect(DB_PATH)
    query = """
        SELECT order_number, invoice_number, sku, product_name, pharmacy_name, branch_number,
               salla_qty, abc_qty, difference, case_type, case_label, case_reason, status,
               performed_by, performed_at, customer_name, customer_phone, city, order_status,
               order_date, invoice_date, total_amount, salla_pharmacy_name, abc_pharmacy_name
        FROM reconciliation_items
        WHERE active = 1
    """
    params = []
    if pharmacy_name:
        query += " AND pharmacy_name = ?"
        params.append(pharmacy_name)
    query += " ORDER BY case_type, order_number DESC, sku"
    try:
        return pd.read_sql_query(query, conn, params=params)
    finally:
        conn.close()


def mark_case_done(order_number: str, sku: str, pharmacy_name: str, case_type: str, performed_by: str):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute(
        """
        UPDATE reconciliation_items
        SET status = ?, performed_by = ?, performed_at = ?
        WHERE active = 1 AND order_number = ? AND sku = ? AND pharmacy_name = ? AND case_type = ?
        """,
        (STATUS_DONE, performed_by, now_str(), order_number, sku, pharmacy_name, case_type),
    )
    conn.commit()
    conn.close()


def status_pill(status: str) -> str:
    if status == STATUS_DONE:
        return '<span class="pill pill-green">مغلق</span>'
    return '<span class="pill pill-amber">قيد المتابعة</span>'


def case_pill(case_type: str) -> str:
    mapping = {
        "addition": "pill-blue",
        "return": "pill-red",
        "orphan_salla": "pill-amber",
        "orphan_abc": "pill-slate",
        "branch_mismatch": "pill-red",
        "special_review": "pill-slate",
    }
    css_class = mapping.get(case_type, "pill-slate")
    return f'<span class="pill {css_class}">{CASE_LABELS.get(case_type, case_type)}</span>'


def render_metrics(df: pd.DataFrame):
    cols = st.columns(6)
    metrics = [
        ("إجمالي الحالات", len(df)),
        ("إضافات", int((df["case_type"] == "addition").sum())),
        ("إرجاعات", int((df["case_type"] == "return").sum())),
        ("طلبات بدون فاتورة", int((df["case_type"] == "orphan_salla").sum())),
        ("فواتير بدون طلب", int((df["case_type"] == "orphan_abc").sum())),
        ("تم إنجازها", int((df["status"] == STATUS_DONE).sum())),
    ]
    for col, (label, value) in zip(cols, metrics):
        with col:
            st.markdown(
                f"""
                <div class="metric-box">
                    <div style="font-size:0.88rem;color:#5a7380;">{label}</div>
                    <div style="font-size:2rem;font-weight:800;color:#16425b;">{value}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )


def render_case_cards(df: pd.DataFrame, allow_actions: bool, pharmacist_name: str, pharmacy_name: str):
    if df.empty:
        st.success("لا توجد حالات في هذا القسم.")
        return

    for idx, row in df.iterrows():
        st.markdown(
            f"""
            <div class="action-card">
                <div style="display:flex;justify-content:space-between;gap:1rem;align-items:center;flex-wrap:wrap;">
                    <div>
                        {case_pill(row['case_type'])}
                        &nbsp; {status_pill(row['status'])}
                    </div>
                    <div style="font-weight:700;color:#48606a;">الفرع: {row['pharmacy_name'] or 'غير محدد'}</div>
                </div>
                <div style="margin-top:0.8rem;display:grid;grid-template-columns:repeat(4, minmax(120px, 1fr));gap:0.9rem;">
                    <div><strong>رقم الطلب</strong><br>{row['order_number']}</div>
                    <div><strong>رقم الفاتورة</strong><br>{row['invoice_number'] or 'غير متوفر'}</div>
                    <div><strong>SKU</strong><br>{row['sku']}</div>
                    <div><strong>المنتج</strong><br>{row['product_name'][:70]}</div>
                    <div><strong>كمية سلة</strong><br>{int(row['salla_qty']) if pd.notna(row['salla_qty']) else 0}</div>
                    <div><strong>كمية ABC</strong><br>{int(row['abc_qty']) if pd.notna(row['abc_qty']) else 0}</div>
                    <div><strong>الفرق</strong><br>{row['difference']}</div>
                    <div><strong>الحالة</strong><br>{row['case_reason']}</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if allow_actions and row["status"] != STATUS_DONE and row["case_type"] in {"addition", "return"}:
            button_label = "تأكيد الإضافة" if row["case_type"] == "addition" else "تأكيد الإرجاع"
            if st.button(button_label, key=f"{row['case_type']}_{row['order_number']}_{row['sku']}_{idx}"):
                mark_case_done(
                    order_number=row["order_number"],
                    sku=row["sku"],
                    pharmacy_name=pharmacy_name,
                    case_type=row["case_type"],
                    performed_by=pharmacist_name,
                )
                st.rerun()


def render_admin_dashboard():
    st.markdown(
        """
        <div class="hero">
            <h1>لوحة المدير العام</h1>
            <p>مطابقة أكثر دقة بين سلة و ABC مع فصل الحالات الفعلية عن السطور غير المربوطة وحفظ الإنجاز بين كل رفعة وأخرى.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    latest = get_latest_upload_summary()
    if latest:
        _, file_name, uploaded_by, uploaded_at, *_ = latest
        st.markdown(
            f"""
            <div class="note-card">
                <strong>آخر رفع:</strong> {file_name} &nbsp; | &nbsp;
                <strong>بواسطة:</strong> {uploaded_by} &nbsp; | &nbsp;
                <strong>التاريخ:</strong> {uploaded_at}
            </div>
            """,
            unsafe_allow_html=True,
        )

    with st.expander("رفع ملف الطلبات والفواتير", expanded=True):
        uploaded_file = st.file_uploader("اختر ملف Excel", type=["xlsx"])
        if uploaded_file and st.button("معالجة الملف و ترحيل الحالات", use_container_width=True):
            with st.spinner("جاري قراءة الملف وتصنيف الحالات بدقة..."):
                results, upload_batch_id = process_excel(uploaded_file, st.session_state.username)
            st.success(f"تمت المعالجة بنجاح. رقم دفعة الرفع: {upload_batch_id}")
            st.session_state.last_processed_preview = results.head(20)

    df = fetch_active_items()
    if df.empty:
        st.info("لا توجد بيانات فعالة بعد. ارفع الملف من الأعلى لبدء التحليل.")
        return

    render_metrics(df)

    st.markdown('<div class="section-title">آخر دخول للصيدليات</div>', unsafe_allow_html=True)
    last_logins = get_all_last_logins()
    if not last_logins.empty:
        cols = st.columns(4)
        for idx, (_, row) in enumerate(last_logins.head(8).iterrows()):
            with cols[idx % 4]:
                st.markdown(
                    f"""
                    <div class="note-card">
                        <strong>{row['pharmacy_name']}</strong><br>
                        <span style="color:#58707a;">{row['pharmacist_name'] or 'غير مسجل'}</span><br>
                        <span style="color:#58707a;">{row['last_login'][:16] if row['last_login'] else 'لم يدخل بعد'}</span>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

    tab1, tab2, tab3, tab4, tab5 = st.tabs(
        ["الإضافات", "الإرجاعات", "طلبات بدون فاتورة", "فواتير بدون طلب", "حالات المراجعة"]
    )

    with tab1:
        st.dataframe(
            df[df["case_type"] == "addition"][
                ["order_number", "sku", "product_name", "pharmacy_name", "salla_qty", "abc_qty", "difference", "status", "performed_by"]
            ],
            use_container_width=True,
        )

    with tab2:
        st.dataframe(
            df[df["case_type"] == "return"][
                ["order_number", "invoice_number", "sku", "product_name", "pharmacy_name", "salla_qty", "abc_qty", "difference", "status", "performed_by"]
            ],
            use_container_width=True,
        )

    with tab3:
        st.dataframe(
            df[df["case_type"] == "orphan_salla"][
                ["order_number", "sku", "product_name", "pharmacy_name", "salla_qty", "order_status", "customer_name", "city"]
            ],
            use_container_width=True,
        )

    with tab4:
        st.dataframe(
            df[df["case_type"] == "orphan_abc"][
                ["order_number", "invoice_number", "sku", "product_name", "pharmacy_name", "abc_qty", "invoice_date"]
            ],
            use_container_width=True,
        )

    with tab5:
        review_df = df[df["case_type"].isin(["branch_mismatch", "special_review"])]
        st.dataframe(
            review_df[
                [
                    "order_number",
                    "invoice_number",
                    "sku",
                    "product_name",
                    "pharmacy_name",
                    "salla_pharmacy_name",
                    "abc_pharmacy_name",
                    "case_label",
                    "case_reason",
                ]
            ],
            use_container_width=True,
        )


def render_pharmacy_dashboard():
    pharmacy_name = st.session_state.username
    pharmacist_name = st.session_state.pharmacist_name or ""
    branch_number = get_branch_number(pharmacy_name)

    st.markdown(
        f"""
        <div class="hero">
            <h1>{pharmacy_name}</h1>
            <p>فرع رقم {branch_number} | الصيدلي: {pharmacist_name or 'غير مسجل بعد'}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    df = fetch_active_items(pharmacy_name)
    if df.empty:
        st.info("لا توجد حالات نشطة لهذا الفرع حاليًا.")
        return

    render_metrics(df)

    additions_df = df[df["case_type"] == "addition"].copy()
    returns_df = df[df["case_type"] == "return"].copy()
    review_df = df[df["case_type"].isin(["orphan_salla", "orphan_abc", "branch_mismatch", "special_review"])].copy()

    st.markdown('<div class="section-title">الإضافات المطلوبة</div>', unsafe_allow_html=True)
    render_case_cards(additions_df, True, pharmacist_name, pharmacy_name)

    st.markdown('<div class="section-title">الإرجاعات المطلوبة</div>', unsafe_allow_html=True)
    render_case_cards(returns_df, True, pharmacist_name, pharmacy_name)

    st.markdown('<div class="section-title">حالات تحتاج مراجعة</div>', unsafe_allow_html=True)
    render_case_cards(review_df, False, pharmacist_name, pharmacy_name)


ensure_database()


for key, default_value in {
    "logged_in": False,
    "username": "",
    "user_role": "",
    "pharmacist_name": "",
    "last_processed_preview": None,
}.items():
    if key not in st.session_state:
        st.session_state[key] = default_value


with st.sidebar:
    st.title("نظام بلسم")
    st.caption("مطابقة الطلبات والفواتير")
    st.markdown("---")

    if not st.session_state.logged_in:
        username = st.text_input("اسم المستخدم")
        password = st.text_input("كلمة المرور", type="password")
        if st.button("دخول", use_container_width=True):
            user = fetch_user(username, password)
            if user:
                st.session_state.logged_in = True
                st.session_state.username = user[0]
                st.session_state.user_role = user[1]
                st.session_state.pharmacist_name = user[2] or ""
                st.rerun()
            else:
                st.error("بيانات الدخول غير صحيحة.")
    else:
        st.success(st.session_state.username)
        if st.button("تسجيل خروج", use_container_width=True):
            for key in ["logged_in", "username", "user_role", "pharmacist_name", "last_processed_preview"]:
                st.session_state[key] = False if key == "logged_in" else ""
            st.session_state.last_processed_preview = None
            st.rerun()


if not st.session_state.logged_in:
    st.markdown(
        """
        <div class="hero">
            <h1>نظام بلسم لمراقبة إدخالات الفواتير</h1>
            <p>يعرض الإضافات والإرجاعات الفعلية، ويفصل السطور غير المربوطة وحالات اختلاف الفرع، ويحافظ على حالة كل فرع بين عمليات الرفع.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <div class="note-card">
            <strong>ما الذي تغيّر في هذه النسخة؟</strong><br>
            1. لم تعد الحالات تضيع عند كل تشغيل.<br>
            2. تم فصل الإرجاع الفعلي عن "فاتورة بدون طلب" و"طلب بدون فاتورة".<br>
            3. أصبح ترحيل حالة كل صيدلية تلقائيًا عند رفع ملف جديد إذا بقي نفس البند مفتوحًا.
        </div>
        """,
        unsafe_allow_html=True,
    )
elif st.session_state.user_role == "pharmacy" and not st.session_state.pharmacist_name:
    st.markdown("### الرجاء إدخال اسم الصيدلي")
    pharmacist_name_input = st.text_input("اسم الصيدلي")
    if st.button("حفظ الاسم", use_container_width=True):
        if pharmacist_name_input.strip():
            st.session_state.pharmacist_name = pharmacist_name_input.strip()
            update_last_access(st.session_state.username, st.session_state.pharmacist_name)
            st.success("تم حفظ الاسم بنجاح.")
            st.rerun()
else:
    if st.session_state.user_role == "pharmacy":
        update_last_access(st.session_state.username, st.session_state.pharmacist_name)
        render_pharmacy_dashboard()
    else:
        render_admin_dashboard()


st.markdown("---")
st.markdown(
    """
    <div style="text-align:center;color:#607783;padding:0.6rem 0 0.8rem;">
        نظام بلسم لمطابقة الطلبات والفواتير © 2026
    </div>
    """,
    unsafe_allow_html=True,
)
