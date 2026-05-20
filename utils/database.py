import os
import sqlite3
import json
from datetime import datetime
import pandas as pd
import streamlit as st

DB_DIR = "data"
DB_PATH = os.path.join(DB_DIR, "pharmacy_reconciliation.db")
PHARMACY_COUNT = 17

def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def pharmacy_names():
    return [f"Balsam Alula Pharmacy {i:02d}" for i in range(1, PHARMACY_COUNT + 1)]

def init_database():
    """تهيئة قاعدة البيانات مع دعم الترقية"""
    os.makedirs(DB_DIR, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    # التحقق من وجود جدول users وإضافة الأعمدة المفقودة
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='users'")
    table_exists = cur.fetchone()
    
    if not table_exists:
        # إنشاء الجدول من الصفر
        cur.execute("""
            CREATE TABLE users (
                username TEXT PRIMARY KEY,
                password TEXT NOT NULL,
                role TEXT NOT NULL,
                pharmacist_name TEXT DEFAULT '',
                last_login TEXT DEFAULT '',
                can_view_dashboard INTEGER DEFAULT 1,
                can_view_balances INTEGER DEFAULT 0,
                can_view_monitoring INTEGER DEFAULT 0,
                can_manage_users INTEGER DEFAULT 0
            )
        """)
    else:
        # إضافة الأعمدة المفقودة
        cur.execute("PRAGMA table_info(users)")
        existing_columns = [row[1] for row in cur.fetchall()]
        
        new_columns = {
            "can_view_dashboard": "INTEGER DEFAULT 1",
            "can_view_balances": "INTEGER DEFAULT 0",
            "can_view_monitoring": "INTEGER DEFAULT 0",
            "can_manage_users": "INTEGER DEFAULT 0"
        }
        
        for col_name, col_type in new_columns.items():
            if col_name not in existing_columns:
                try:
                    cur.execute(f"ALTER TABLE users ADD COLUMN {col_name} {col_type}")
                except:
                    pass

    # Last access table
    cur.execute("""
        CREATE TABLE IF NOT EXISTS last_access (
            pharmacy_name TEXT PRIMARY KEY,
            last_login TEXT DEFAULT '',
            pharmacist_name TEXT DEFAULT ''
        )
    """)

    # Uploads table
    cur.execute("""
        CREATE TABLE IF NOT EXISTS uploads (
            upload_batch_id TEXT PRIMARY KEY,
            session_name TEXT DEFAULT '',
            file_name TEXT,
            uploaded_by TEXT,
            uploaded_at TEXT,
            total_cases INTEGER DEFAULT 0,
            total_additions INTEGER DEFAULT 0,
            total_returns INTEGER DEFAULT 0,
            total_orphan_salla INTEGER DEFAULT 0,
            total_orphan_abc INTEGER DEFAULT 0,
            total_branch_mismatch INTEGER DEFAULT 0,
            total_special_review INTEGER DEFAULT 0,
            is_locked INTEGER DEFAULT 0,
            locked_by TEXT DEFAULT '',
            locked_at TEXT DEFAULT '',
            is_active INTEGER DEFAULT 0
        )
    """)

    # Reconciliation items table
    cur.execute("""
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
            abc_pharmacist_name TEXT,
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
            profile_type TEXT DEFAULT '',
            profile_type_from_abc TEXT DEFAULT '',
            receipt_classification TEXT DEFAULT '',
            all_abc_pharmacies TEXT DEFAULT '',
            other_branch_details TEXT DEFAULT '',
            pharmacist_note TEXT DEFAULT '',
            total_amount REAL DEFAULT 0,
            first_seen_at TEXT,
            last_seen_at TEXT,
            active INTEGER DEFAULT 1,
            hidden_from_pharmacy INTEGER DEFAULT 0,
            hidden_by TEXT DEFAULT '',
            hidden_at TEXT DEFAULT ''
        )
    """)

    # Insert default admin if not exists
    cur.execute("SELECT * FROM users WHERE username = 'admin'")
    if not cur.fetchone():
        cur.execute("""
            INSERT INTO users 
            (username, password, role, pharmacist_name, can_view_dashboard, can_view_balances, can_view_monitoring, can_manage_users)
            VALUES ('admin', 'admin123', 'admin', 'مدير النظام', 1, 1, 1, 1)
        """)

    # Insert default pharmacies if not exists
    for index, name in enumerate(pharmacy_names(), start=1):
        cur.execute("SELECT * FROM users WHERE username = ?", (name,))
        if not cur.fetchone():
            cur.execute("""
                INSERT INTO users 
                (username, password, role, pharmacist_name, can_view_dashboard, can_view_balances, can_view_monitoring, can_manage_users)
                VALUES (?, ?, 'pharmacy', '', 1, 0, 0, 0)
            """, (name, f"balsam{index}"))

    conn.commit()
    conn.close()

def get_user_permissions(username: str):
    """الحصول على صلاحيات المستخدم"""
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        SELECT role, can_view_dashboard, can_view_balances, can_view_monitoring, can_manage_users, pharmacist_name
        FROM users WHERE username = ?
    """, (username,))
    result = cur.fetchone()
    conn.close()
    if result:
        return {
            "role": result[0],
            "can_view_dashboard": bool(result[1]),
            "can_view_balances": bool(result[2]),
            "can_view_monitoring": bool(result[3]),
            "can_manage_users": bool(result[4]),
            "pharmacist_name": result[5] or ""
        }
    return None

def update_user_permissions(username: str, permissions: dict):
    """تحديث صلاحيات المستخدم"""
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        UPDATE users 
        SET can_view_dashboard = ?, can_view_balances = ?, can_view_monitoring = ?, can_manage_users = ?, pharmacist_name = ?
        WHERE username = ?
    """, (
        permissions.get("can_view_dashboard", 0),
        permissions.get("can_view_balances", 0),
        permissions.get("can_view_monitoring", 0),
        permissions.get("can_manage_users", 0),
        permissions.get("pharmacist_name", ""),
        username
    ))
    conn.commit()
    conn.close()

def add_user(username: str, password: str, role: str, pharmacist_name: str = ""):
    """إضافة مستخدم جديد"""
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    try:
        cur.execute("""
            INSERT INTO users 
            (username, password, role, pharmacist_name, can_view_dashboard, can_view_balances, can_view_monitoring, can_manage_users)
            VALUES (?, ?, ?, ?, 1, 0, 0, 0)
        """, (username, password, role, pharmacist_name))
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        return False
    finally:
        conn.close()

def delete_user(username: str):
    """حذف مستخدم"""
    if username == "admin":
        return False
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("DELETE FROM users WHERE username = ?", (username,))
    conn.commit()
    conn.close()
    return True

def get_all_users():
    """الحصول على جميع المستخدمين"""
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("""
        SELECT username, role, pharmacist_name, last_login, 
               can_view_dashboard, can_view_balances, can_view_monitoring, can_manage_users
        FROM users ORDER BY role, username
    """, conn)
    conn.close()
    return df

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
    try:
        cur.execute("""
            SELECT upload_batch_id, file_name, uploaded_by, uploaded_at, total_cases,
                   total_additions, total_returns, total_orphan_salla, total_orphan_abc,
                   total_branch_mismatch, total_special_review, is_locked, session_name
            FROM uploads
            WHERE is_active = 1
            ORDER BY uploaded_at DESC
            LIMIT 1
        """)
        return cur.fetchone()
    except:
        return None
    finally:
        conn.close()

def get_all_sessions() -> pd.DataFrame:
    conn = sqlite3.connect(DB_PATH)
    try:
        return pd.read_sql_query("""
            SELECT upload_batch_id, session_name, file_name, uploaded_by, uploaded_at, 
                   total_cases, total_additions, total_returns, is_locked, is_active
            FROM uploads
            ORDER BY uploaded_at DESC
        """, conn)
    finally:
        conn.close()

def get_completed_items(pharmacy_name: str = None) -> pd.DataFrame:
    conn = sqlite3.connect(DB_PATH)
    query = """
        SELECT order_number, invoice_number, sku, product_name, case_type, case_label,
               performed_by, performed_at, status, item_key, pharmacy_name
        FROM reconciliation_items
        WHERE active = 1 AND status = 'تم'
    """
    params = []
    if pharmacy_name:
        query += " AND pharmacy_name = ?"
        params.append(pharmacy_name)
    query += " ORDER BY performed_at DESC"
    
    try:
        return pd.read_sql_query(query, conn, params=params if params else None)
    finally:
        conn.close()

def fetch_active_items(pharmacy_name: str = None, include_hidden: bool = False) -> pd.DataFrame:
    conn = sqlite3.connect(DB_PATH)
    
    cur = conn.cursor()
    cur.execute("""
        SELECT upload_batch_id FROM uploads 
        WHERE is_active = 1
        ORDER BY uploaded_at DESC LIMIT 1
    """)
    active_session = cur.fetchone()
    
    if not active_session:
        conn.close()
        return pd.DataFrame()
    
    active_batch_id = active_session[0]
    
    # التحقق من وجود عمود is_locked
    cur.execute("PRAGMA table_info(uploads)")
    existing_columns = [row[1] for row in cur.fetchall()]
    has_lock_column = "is_locked" in existing_columns
    
    if has_lock_column:
        cur.execute("SELECT is_locked FROM uploads WHERE upload_batch_id = ?", (active_batch_id,))
        lock_result = cur.fetchone()
        is_locked = lock_result[0] if lock_result else 0
    else:
        is_locked = 0
    
    query = """
        SELECT order_number, invoice_number, sku, product_name, pharmacy_name, branch_number,
               salla_qty, abc_qty, difference, case_type, case_label, case_reason, status,
               performed_by, performed_at, customer_name, customer_phone, city, order_status,
               order_date, invoice_date, total_amount, first_seen_at, last_seen_at,
               profile_type, profile_type_from_abc, receipt_classification, 
               all_abc_pharmacies, other_branch_details, pharmacist_note, item_key,
               abc_pharmacy_name, abc_pharmacist_name,
               ? as is_locked
        FROM reconciliation_items
        WHERE active = 1 AND upload_batch_id = ?
    """
    params = [1 if is_locked else 0, active_batch_id]
    
    if pharmacy_name:
        query += " AND pharmacy_name = ?"
        params.append(pharmacy_name)
    
    query += " ORDER BY case_type, order_number DESC, sku"
    
    try:
        df = pd.read_sql_query(query, conn, params=params)
        return df
    except Exception as e:
        return pd.DataFrame()
    finally:
        conn.close()

def hide_item_from_pharmacy(item_key: str, hidden_by: str):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        UPDATE reconciliation_items 
        SET hidden_from_pharmacy = 1, hidden_by = ?, hidden_at = ?
        WHERE item_key = ?
    """, (hidden_by, now_str(), item_key))
    conn.commit()
    conn.close()

def unhide_item_from_pharmacy(item_key: str):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        UPDATE reconciliation_items 
        SET hidden_from_pharmacy = 0, hidden_by = '', hidden_at = ''
        WHERE item_key = ?
    """, (item_key,))
    conn.commit()
    conn.close()

def save_case_note(order_number: str, sku: str, pharmacy_name: str, case_type: str, note: str):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        UPDATE reconciliation_items
        SET pharmacist_note = ?
        WHERE active = 1 AND order_number = ? AND sku = ? AND pharmacy_name = ? AND case_type = ?
    """, (note, order_number, sku, pharmacy_name, case_type))
    conn.commit()
    conn.close()

def mark_case_done(order_number: str, sku: str, pharmacy_name: str, case_type: str, performed_by: str):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        UPDATE reconciliation_items
        SET status = 'تم', performed_by = ?, performed_at = ?
        WHERE active = 1 AND order_number = ? AND sku = ? AND pharmacy_name = ? AND case_type = ?
    """, (performed_by, now_str(), order_number, sku, pharmacy_name, case_type))
    conn.commit()
    conn.close()

def reopen_case(order_number: str, sku: str, pharmacy_name: str, case_type: str):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        UPDATE reconciliation_items
        SET status = 'قيد المتابعة', performed_by = '', performed_at = ''
        WHERE active = 1 AND order_number = ? AND sku = ? AND pharmacy_name = ? AND case_type = ?
    """, (order_number, sku, pharmacy_name, case_type))
    conn.commit()
    conn.close()

def persist_reconciliation_results(results: pd.DataFrame, uploaded_file_name: str, uploaded_by: str):
    import uuid
    upload_batch_id = uuid.uuid4().hex
    timestamp = now_str()
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    cur.execute("BEGIN")

    # إدراج بيانات الرفع
    cur.execute("""
        INSERT INTO uploads (
            upload_batch_id, file_name, uploaded_by, uploaded_at, total_cases,
            total_additions, total_returns, total_orphan_salla, total_orphan_abc,
            total_branch_mismatch, total_special_review, is_locked, is_active
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, 1)
    """, (
        upload_batch_id, uploaded_file_name, uploaded_by, timestamp,
        len(results),
        int((results["case_type"] == "addition").sum()),
        int((results["case_type"] == "return").sum()),
        int((results["case_type"] == "orphan_salla").sum()),
        int((results["case_type"].isin(["orphan_abc", "post_cutoff_abc"])).sum()),
        int((results["case_type"] == "branch_mismatch").sum()),
        int((results["case_type"] == "special_review").sum()),
    ))

    # إدراج العناصر
    for _, row in results.iterrows():
        cur.execute("""
            INSERT OR REPLACE INTO reconciliation_items (
                item_key, upload_batch_id, order_number, invoice_number, sku, product_name,
                salla_product_name, abc_product_name, pharmacy_name, salla_pharmacy_name,
                abc_pharmacy_name, abc_pharmacist_name, branch_number, salla_qty, abc_qty, difference,
                case_type, case_label, case_reason, status, performed_by, performed_at,
                customer_name, customer_phone, city, order_status, order_date,
                invoice_date, profile_type, profile_type_from_abc, receipt_classification, 
                all_abc_pharmacies, other_branch_details, pharmacist_note, total_amount, 
                first_seen_at, last_seen_at, active, hidden_from_pharmacy
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, 0)
        """, (
            row["item_key"], upload_batch_id, str(row["order_number"]), str(row.get("invoice_number", "")),
            str(row["sku"]), str(row["product_name"])[:200], str(row.get("salla_product_name", ""))[:200],
            str(row.get("abc_product_name", ""))[:200], str(row["pharmacy_name"]),
            str(row.get("salla_pharmacy_name", "")), str(row.get("abc_pharmacy_name", "")),
            str(row.get("abc_pharmacist_name", "")), str(row.get("branch_number", "")),
            float(row["salla_qty"]), float(row["abc_qty"]), float(row["difference"]),
            str(row["case_type"]), str(row["case_label"]), str(row.get("case_reason", ""))[:500],
            "قيد المتابعة", "", "", str(row.get("customer_name", ""))[:100], str(row.get("customer_phone", "")),
            str(row.get("city", "")), str(row.get("order_status", "")), str(row.get("order_date", "")),
            str(row.get("invoice_date", "")), str(row.get("profile_type", "")), str(row.get("profile_type_from_abc", "")),
            str(row.get("receipt_classification", "")), str(row.get("all_abc_pharmacies", "")),
            str(row.get("other_branch_details", "")), "", float(row.get("total_amount", 0)), timestamp, timestamp
        ))

    # تعطيل العناصر القديمة
    cur.execute("""
        UPDATE reconciliation_items
        SET active = CASE WHEN upload_batch_id = ? THEN 1 ELSE 0 END
    """, (upload_batch_id,))

    # تعطيل الجلسات الأخرى وتفعيل الجلسة الحالية
    cur.execute("UPDATE uploads SET is_active = 0")
    cur.execute("UPDATE uploads SET is_active = 1 WHERE upload_batch_id = ?", (upload_batch_id,))
    
    # إنشاء اسم الجلسة
    session_name = datetime.now().strftime("%Y-%m-%d %H:%M")
    cur.execute("UPDATE uploads SET session_name = ? WHERE upload_batch_id = ?", (session_name, upload_batch_id))

    conn.commit()
    conn.close()
    
    return upload_batch_id
