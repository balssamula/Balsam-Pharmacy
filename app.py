import streamlit as st
import pandas as pd
import sqlite3
import os
import re
from datetime import datetime

st.set_page_config(page_title="نظام بلسم - إدارة الصيدليات", layout="wide", initial_sidebar_state="expanded")

# ========== CSS Professional ==========
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@300;400;500;700&display=swap');
    * { font-family: 'Tajawal', sans-serif; }
    
    .main-header {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        padding: 2rem;
        border-radius: 20px;
        color: white;
        text-align: center;
        margin-bottom: 2rem;
        box-shadow: 0 4px 15px rgba(0,0,0,0.2);
    }
    
    .stat-card {
        background: white;
        padding: 1.2rem;
        border-radius: 15px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        text-align: center;
        margin: 10px;
        transition: transform 0.3s;
    }
    
    .stat-card:hover {
        transform: translateY(-5px);
    }
    
    .stat-number {
        font-size: 2rem;
        font-weight: bold;
        color: #2a5298;
    }
    
    .success-badge {
        background: linear-gradient(135deg, #28a745, #20c997);
        color: white;
        padding: 8px 15px;
        border-radius: 25px;
        font-size: 0.8rem;
        text-align: center;
        display: inline-block;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
    }
    
    .pending-badge {
        background: linear-gradient(135deg, #ffc107, #fd7e14);
        color: #333;
        padding: 8px 15px;
        border-radius: 25px;
        font-size: 0.8rem;
        text-align: center;
        display: inline-block;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
    }
    
    .pharmacy-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1.5rem;
        border-radius: 15px;
        color: white;
        text-align: center;
        margin: 1rem 0;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }
    
    .info-card {
        background: linear-gradient(135deg, #f8f9fa, #e9ecef);
        padding: 1rem;
        border-radius: 10px;
        margin: 0.5rem 0;
        border-right: 4px solid #2a5298;
    }
    
    .stButton button {
        width: 100%;
        border-radius: 8px;
        transition: all 0.3s;
        font-weight: bold;
    }
    
    .stButton button:hover {
        transform: scale(1.02);
        box-shadow: 0 4px 10px rgba(0,0,0,0.2);
    }
    
    .action-card {
        background: white;
        border-radius: 15px;
        padding: 1rem;
        margin-bottom: 1rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        border-right: 4px solid #2a5298;
        transition: all 0.3s;
    }
    
    .action-card:hover {
        box-shadow: 0 4px 15px rgba(0,0,0,0.15);
        transform: translateX(-3px);
    }
    
    .invoice-badge {
        background: #e9ecef;
        padding: 4px 10px;
        border-radius: 20px;
        font-family: monospace;
        font-size: 0.8rem;
        display: inline-block;
    }
    
    .last-login-card {
        background: linear-gradient(135deg, #f8f9fa, #ffffff);
        padding: 0.8rem;
        border-radius: 10px;
        margin: 0.5rem;
        border: 1px solid #dee2e6;
    }
    
    .section-title {
        font-size: 1.3rem;
        font-weight: bold;
        color: #2a5298;
        margin: 1rem 0;
        padding-right: 0.5rem;
        border-right: 4px solid #2a5298;
    }
    
    hr {
        margin: 1.5rem 0;
        background: linear-gradient(90deg, #2a5298, transparent);
        height: 2px;
        border: none;
    }
</style>
""", unsafe_allow_html=True)

# ========== DATABASE SETUP ==========
os.makedirs("data", exist_ok=True)

def init_database():
    """تهيئة قاعدة البيانات من الصفر"""
    db_path = 'data/pharmacy.db'
    
    if os.path.exists(db_path):
        try:
            os.remove(db_path)
        except:
            pass
    
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    
    # Users table
    c.execute('''CREATE TABLE users (
        username TEXT PRIMARY KEY,
        password TEXT,
        role TEXT,
        pharmacist_name TEXT,
        last_login TEXT
    )''')
    
    # Adjustments table with invoice number
    c.execute('''CREATE TABLE adjustments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        order_number TEXT,
        invoice_number TEXT,
        sku TEXT,
        product_name TEXT,
        branch_number TEXT,
        pharmacist TEXT,
        pharmacy_name TEXT,
        salla_qty REAL,
        abc_qty REAL,
        difference REAL,
        action TEXT,
        status TEXT,
        performed_by TEXT,
        performed_at TEXT,
        timestamp TEXT,
        order_status TEXT,
        city TEXT,
        customer_name TEXT,
        customer_phone TEXT,
        order_date TEXT,
        total_amount REAL,
        invoice_date TEXT
    )''')
    
    # Last access table
    c.execute('''CREATE TABLE last_access (
        pharmacy_name TEXT PRIMARY KEY,
        last_login TEXT,
        pharmacist_name TEXT
    )''')
    
    # Insert pharmacies
    for i in range(1, 18):
        pharmacy_name = f"Balsam Alula Pharmacy {i:02d}"
        c.execute("INSERT INTO users (username, password, role, pharmacist_name) VALUES (?, ?, 'pharmacy', ?)",
                 (pharmacy_name, f"balsam{i}", ""))
    
    # Insert admin
    c.execute("INSERT INTO users (username, password, role, pharmacist_name) VALUES ('admin', 'admin123', 'admin', 'Manager')")
    
    conn.commit()
    conn.close()
    
    return [f"Balsam Alula Pharmacy {i:02d}" for i in range(1, 18)]

# Initialize database
try:
    pharmacies_list = init_database()
except Exception as e:
    pharmacies_list = [f"Balsam Alula Pharmacy {i:02d}" for i in range(1, 18)]

# ========== HELPER FUNCTIONS ==========
def extract_branch_from_status(status_text):
    if not status_text or pd.isna(status_text):
        return None
    match = re.search(r'فرع\s*(\d+)', str(status_text))
    if match:
        return f"{int(match.group(1)):02d}"
    return None

def determine_branch(order_status, city):
    branch_num = extract_branch_from_status(order_status)
    if branch_num:
        return f"Balsam Alula Pharmacy {branch_num}", branch_num
    
    excluded_statuses = ['تم التوصيل', 'ملغي', 'مسترجع', 'محذوف']
    if any(s in str(order_status) for s in excluded_statuses):
        if city == 'AL ULA':
            return "Balsam Alula Pharmacy 09", "09"
        else:
            return "Balsam Alula Pharmacy 13", "13"
    
    return "Balsam Alula Pharmacy 13", "13"

def is_valid_sku(sku):
    if pd.isna(sku):
        return False
    sku_str = str(sku).strip()
    invalid_values = ['', '0', '1', '200', 'nan', 'NaN', 'None', 'null']
    if sku_str in invalid_values:
        return False
    return sku_str.replace('.', '').isdigit()

def update_pharmacist_name(pharmacy_name, pharmacist_name):
    try:
        conn = sqlite3.connect('data/pharmacy.db')
        c = conn.cursor()
        c.execute("UPDATE users SET pharmacist_name = ?, last_login = ? WHERE username = ?", 
                 (pharmacist_name, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), pharmacy_name))
        c.execute("INSERT OR REPLACE INTO last_access (pharmacy_name, last_login, pharmacist_name) VALUES (?, ?, ?)",
                 (pharmacy_name, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), pharmacist_name))
        conn.commit()
        conn.close()
        return True
    except:
        return False

def get_pharmacist_name(pharmacy_name):
    try:
        conn = sqlite3.connect('data/pharmacy.db')
        c = conn.cursor()
        c.execute("SELECT pharmacist_name FROM users WHERE username = ?", (pharmacy_name,))
        result = c.fetchone()
        conn.close()
        return result[0] if result and result[0] else ""
    except:
        return ""

def get_last_login(pharmacy_name):
    try:
        conn = sqlite3.connect('data/pharmacy.db')
        c = conn.cursor()
        c.execute("SELECT last_login FROM last_access WHERE pharmacy_name = ?", (pharmacy_name,))
        result = c.fetchone()
        conn.close()
        return result[0] if result else "لم يدخل بعد"
    except:
        return "لم يدخل بعد"

def get_all_last_logins():
    """الحصول على آخر دخول لجميع الصيدليات"""
    try:
        conn = sqlite3.connect('data/pharmacy.db')
        df = pd.read_sql_query("SELECT pharmacy_name, last_login, pharmacist_name FROM last_access ORDER BY last_login DESC", conn)
        conn.close()
        return df
    except:
        return pd.DataFrame()

def process_excel(uploaded_file):
    """معالجة ملف Excel مع استبعاد FREE GIFTS"""
    try:
        df_salla = pd.read_excel(uploaded_file, sheet_name="سلة")
        df_abc = pd.read_excel(uploaded_file, sheet_name="abc")
        
        # استبعاد FREE GIFTS FROM ABC
        if 'نوع البروفايل' in df_abc.columns:
            df_abc = df_abc[df_abc['نوع البروفايل'] != "FREE GIFTS FOR CUSTOMERS"]
            st.info(f"✅ تم استبعاد الصفوف التي تحتوي على FREE GIFTS FOR CUSTOMERS")
        
        # Clean salla
        df_salla = df_salla[df_salla['رقم الطلب'].notna()]
        valid_mask = df_salla['SKU'].apply(is_valid_sku)
        df_salla = df_salla[valid_mask]
        
        # Determine branch
        branch_info = df_salla.apply(
            lambda row: determine_branch(row['حالة الطلب'], row['المدينة']), 
            axis=1
        )
        df_salla['الفرع'] = branch_info.apply(lambda x: x[0])
        df_salla['رقم_الفرع'] = branch_info.apply(lambda x: x[1])
        
        # Group salla
        salla_grouped = df_salla.groupby(['رقم الطلب', 'SKU', 'اسم المنتج', 'الفرع', 'رقم_الفرع']).agg({
            'الكمية': 'sum',
            'حالة الطلب': 'first',
            'المدينة': 'first',
            'اسم العميل': 'first',
            'رقم الجوال': 'first',
            'تاريخ الطلب': 'first',
            'إجمالي الطلب': 'first'
        }).reset_index()
        
        # Group ABC with invoice info
        if 'رقم الصنف' in df_abc.columns and 'Net Sold Qty' in df_abc.columns:
            abc_grouped = df_abc.groupby(['رقم الطلب', 'رقم الصنف']).agg({
                'Net Sold Qty': 'sum',
                'رقم الفاتورة': 'first',
                'التاريخ': 'first'
            }).reset_index()
            abc_grouped.rename(columns={
                'رقم الصنف': 'SKU', 
                'Net Sold Qty': 'كمية_ABC',
                'رقم الفاتورة': 'رقم_الفاتورة',
                'التاريخ': 'تاريخ_الفاتورة'
            }, inplace=True)
        else:
            abc_grouped = pd.DataFrame(columns=['رقم الطلب', 'SKU', 'كمية_ABC', 'رقم_الفاتورة', 'تاريخ_الفاتورة'])
        
        # Merge
        merged = pd.merge(salla_grouped, abc_grouped, on=['رقم الطلب', 'SKU'], how='outer')
        merged = merged.fillna(0)
        
        # Calculate difference
        merged['الفرق'] = merged['الكمية'] - merged['كمية_ABC']
        merged['نوع_الاجراء'] = merged['الفرق'].apply(
            lambda x: 'إضافة' if x > 0 else ('إرجاع' if x < 0 else 'مطابق')
        )
        
        result = merged[merged['نوع_الاجراء'].isin(['إضافة', 'إرجاع'])].copy()
        
        # Save to database
        conn = sqlite3.connect('data/pharmacy.db')
        cursor = conn.cursor()
        
        for _, row in result.iterrows():
            cursor.execute("""
                SELECT status, performed_by FROM adjustments 
                WHERE order_number=? AND sku=? AND pharmacy_name=?
                ORDER BY timestamp DESC LIMIT 1
            """, (str(row['رقم الطلب']), str(row['SKU']), row['الفرع']))
            
            existing = cursor.fetchone()
            status = existing[0] if existing else "لم يبدأ"
            performed_by = existing[1] if existing else ""
            
            cursor.execute("""
                INSERT OR REPLACE INTO adjustments 
                (order_number, invoice_number, sku, product_name, branch_number, pharmacist, pharmacy_name,
                 salla_qty, abc_qty, difference, action, status, performed_by, performed_at,
                 order_status, city, customer_name, customer_phone, order_date, total_amount, invoice_date, timestamp)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                str(row['رقم الطلب']), str(row.get('رقم_الفاتورة', '')), str(row['SKU']), str(row['اسم المنتج'])[:100],
                str(row['رقم_الفرع']), '', str(row['الفرع']),
                float(row['الكمية']), float(row['كمية_ABC']), float(row['الفرق']),
                str(row['نوع_الاجراء']), status, performed_by, performed_by,
                str(row['حالة الطلب'])[:50], str(row['المدينة']), str(row.get('اسم العميل', ''))[:50],
                str(row.get('رقم الجوال', '')), str(row.get('تاريخ الطلب', '')), 
                float(row.get('إجمالي الطلب', 0)), str(row.get('تاريخ_الفاتورة', '')),
                datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ))
        
        conn.commit()
        conn.close()
        
        return result
        
    except Exception as e:
        st.error(f"❌ خطأ في معالجة الملف: {str(e)}")
        return None

def record_adjustment(order_number, sku, pharmacy_name, action):
    try:
        conn = sqlite3.connect('data/pharmacy.db')
        c = conn.cursor()
        
        c.execute("""
            UPDATE adjustments 
            SET status = 'تم', performed_by = ?, performed_at = ?, timestamp = ?
            WHERE order_number = ? AND sku = ? AND pharmacy_name = ? AND action = ?
        """, (st.session_state.pharmacist_name, datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
              datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
              str(order_number), str(sku), pharmacy_name, action))
        
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        st.error(f"خطأ: {str(e)}")
        return False

def get_pharmacy_data(pharmacy_name):
    try:
        conn = sqlite3.connect('data/pharmacy.db')
        df = pd.read_sql_query("""
            SELECT order_number, invoice_number, sku, product_name, salla_qty, abc_qty, difference, 
                   action, status, order_status, city, customer_name, customer_phone, order_date,
                   performed_by, performed_at, invoice_date
            FROM adjustments 
            WHERE pharmacy_name = ?
            ORDER BY order_number DESC
        """, conn, params=(pharmacy_name,))
        conn.close()
        return df
    except:
        return pd.DataFrame()

def get_all_adjustments():
    try:
        conn = sqlite3.connect('data/pharmacy.db')
        df = pd.read_sql_query("""
            SELECT order_number, invoice_number, sku, product_name, branch_number, pharmacy_name, 
                   salla_qty, abc_qty, difference, action, status, performed_by, performed_at,
                   order_status, city, customer_name, customer_phone, order_date, invoice_date
            FROM adjustments 
            ORDER BY order_number DESC
        """, conn)
        conn.close()
        return df
    except:
        return pd.DataFrame()

# ========== SESSION STATE ==========
for key in ['logged_in', 'username', 'user_role', 'pharmacist_name', 'processed_data']:
    if key not in st.session_state:
        st.session_state[key] = None if key != 'logged_in' else False

# ========== LOGIN ==========
with st.sidebar:
    st.image("https://img.icons8.com/color/96/000000/pharmacy.png", width=80)
    st.title("🌟 نظام بلسم")
    st.markdown("---")
    
    if not st.session_state.logged_in:
        st.subheader("🔐 تسجيل الدخول")
        username = st.text_input("👤 اسم المستخدم")
        password = st.text_input("🔒 كلمة المرور", type="password")
        
        if st.button("🚪 دخول", use_container_width=True):
            try:
                conn = sqlite3.connect('data/pharmacy.db')
                c = conn.cursor()
                c.execute("SELECT role, pharmacist_name FROM users WHERE username=? AND password=?", (username, password))
                user = c.fetchone()
                conn.close()
                
                if user:
                    st.session_state.logged_in = True
                    st.session_state.username = username
                    st.session_state.user_role = user[0]
                    if user[0] == "pharmacy":
                        st.session_state.pharmacist_name = user[1] if user[1] else ""
                    st.rerun()
                else:
                    st.error("❌ خطأ في البيانات")
            except Exception as e:
                st.error(f"خطأ: {str(e)}")
    else:
        st.success(f"مرحباً {st.session_state.username}")
        if st.session_state.user_role == "admin":
            st.info("👑 مدير عام")
        else:
            st.info(f"🏥 {st.session_state.username}")
        st.markdown("---")
        if st.button("🚪 تسجيل خروج", use_container_width=True):
            for key in ['logged_in', 'username', 'user_role', 'pharmacist_name', 'processed_data']:
                st.session_state[key] = None
            st.rerun()

# ========== PHARMACIST NAME INPUT ==========
if st.session_state.user_role == "pharmacy" and st.session_state.logged_in:
    if not st.session_state.pharmacist_name:
        st.markdown("### 👤 الرجاء إدخال اسمك")
        col1, col2 = st.columns([2, 1])
        with col1:
            pharmacist_name_input = st.text_input("اسم الصيدلي")
        with col2:
            if st.button("تأكيد الاسم"):
                if pharmacist_name_input.strip():
                    st.session_state.pharmacist_name = pharmacist_name_input
                    if update_pharmacist_name(st.session_state.username, pharmacist_name_input):
                        st.success(f"✅ تم تسجيل اسمك: {pharmacist_name_input}")
                        st.rerun()
        st.stop()

# ========== MAIN CONTENT ==========
if not st.session_state.logged_in:
    st.markdown("""
    <div class="main-header">
        <h1 style="font-size: 2.5rem;">📊 نظام بلسم لإدارة الصيدليات</h1>
        <p style="font-size: 1.2rem;">نظام متكامل لمراقبة طلبات سلة ومطابقتها مع فواتير ABC</p>
    </div>
    """, unsafe_allow_html=True)

elif st.session_state.user_role == "admin":
    # ========== ADMIN DASHBOARD ==========
    st.markdown('<div class="main-header"><h1>👑 لوحة تحكم المدير العام</h1></div>', unsafe_allow_html=True)
    
    # Last logins section - Top of page
    st.markdown('<div class="section-title">📅 آخر دخول للصيدليات</div>', unsafe_allow_html=True)
    last_logins_df = get_all_last_logins()
    if len(last_logins_df) > 0:
        cols = st.columns(4)
        for i, (_, row) in enumerate(last_logins_df.head(8).iterrows()):
            with cols[i % 4]:
                st.markdown(f"""
                <div class="last-login-card">
                    <strong>🏥 {row['pharmacy_name'][-5:]}</strong><br>
                    <small>👤 {row['pharmacist_name'] if row['pharmacist_name'] else 'غير مسجل'}</small><br>
                    <small>📅 {row['last_login'][:16] if row['last_login'] != 'لم يدخل بعد' else 'لم يدخل'}</small>
                </div>
                """, unsafe_allow_html=True)
    st.markdown("<hr>", unsafe_allow_html=True)
    
    # File upload
    with st.expander("📂 رفع ملف الطلبات والفواتير", expanded=True):
        uploaded_file = st.file_uploader("اختر ملف Excel", type=['xlsx'])
        if uploaded_file:
            if st.button("🔄 معالجة الملف", use_container_width=True):
                with st.spinner("جاري معالجة الملف..."):
                    result = process_excel(uploaded_file)
                    if result is not None:
                        st.session_state.processed_data = result
                        st.success("✅ تمت المعالجة بنجاح!")
                        st.balloons()
    
    df = get_all_adjustments()
    
    if len(df) > 0:
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            st.metric("📋 إجمالي الطلبات", len(df['order_number'].unique()))
        with col2:
            st.metric("➕ إضافات", len(df[df['action'] == 'إضافة']))
        with col3:
            st.metric("➖ إرجاعات", len(df[df['action'] == 'إرجاع']))
        with col4:
            st.metric("✅ تم إنجازها", len(df[df['status'] == 'تم']))
        with col5:
            st.metric("⏳ قيد الانتظار", len(df[df['status'] != 'تم']))
        
        st.markdown("<hr>", unsafe_allow_html=True)
        
        tab1, tab2, tab3 = st.tabs(["📈 الإضافات", "📉 الإرجاعات", "🏥 أداء الفروع"])
        
        with tab1:
            additions_df = df[df['action'] == 'إضافة']
            if len(additions_df) > 0:
                additions_df['الحالة'] = additions_df['status'].apply(
                    lambda x: "✅ تمت" if x == "تم" else "⏳ قيد الانتظار"
                )
                st.dataframe(additions_df[['order_number', 'sku', 'product_name', 'salla_qty', 'abc_qty', 
                               'difference', 'pharmacy_name', 'الحالة', 'performed_by']], 
                               use_container_width=True)
            else:
                st.success("🎉 لا توجد إضافات")
        
        with tab2:
            returns_df = df[df['action'] == 'إرجاع']
            if len(returns_df) > 0:
                returns_df['الحالة'] = returns_df['status'].apply(
                    lambda x: "✅ تم" if x == "تم" else "⏳ قيد الانتظار"
                )
                st.dataframe(returns_df[['order_number', 'invoice_number', 'sku', 'product_name', 
                               'salla_qty', 'abc_qty', 'difference', 'pharmacy_name', 
                               'الحالة', 'performed_by', 'invoice_date']], 
                               use_container_width=True)
            else:
                st.success("🎉 لا توجد إرجاعات")
        
        with tab3:
            branch_stats = df.groupby('pharmacy_name').agg({
                'order_number': 'nunique',
                'action': 'count',
                'status': lambda x: (x == 'تم').sum()
            }).reset_index()
            branch_stats.columns = ['الفرع', 'الطلبات', 'الإجراءات', 'تم']
            branch_stats['متبقي'] = branch_stats['الإجراءات'] - branch_stats['تم']
            branch_stats['آخر دخول'] = branch_stats['الفرع'].apply(get_last_login)
            st.dataframe(branch_stats, use_container_width=True)
    else:
        st.info("📂 لا توجد بيانات")

else:
    # ========== PHARMACY DASHBOARD ==========
    pharmacy_name = st.session_state.username
    branch_num = pharmacy_name.split()[-1]
    
    st.markdown(f"""
    <div class="pharmacy-header">
        <h1>🏥 {pharmacy_name}</h1>
        <p>فرع رقم {branch_num} | الصيدلي: {st.session_state.pharmacist_name}</p>
    </div>
    """, unsafe_allow_html=True)
    
    # File upload
    with st.expander("📂 رفع ملف الطلبات والفواتير", expanded=False):
        uploaded_file = st.file_uploader("اختر ملف Excel", type=['xlsx'])
        if uploaded_file:
            if st.button("🔄 معالجة الملف"):
                with st.spinner("جاري المعالجة..."):
                    result = process_excel(uploaded_file)
                    if result is not None:
                        st.session_state.processed_data = result
                        st.success("✅ تمت المعالجة!")
                        st.rerun()
    
    df = get_pharmacy_data(pharmacy_name)
    
    if len(df) > 0:
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("➕ إضافات", len(df[df['action'] == 'إضافة']))
        with col2:
            st.metric("➖ إرجاعات", len(df[df['action'] == 'إرجاع']))
        with col3:
            st.metric("✅ تم", len(df[df['status'] == 'تم']))
        
        st.markdown("<hr>", unsafe_allow_html=True)
        
        # Additions
        additions_df = df[df['action'] == 'إضافة']
        if len(additions_df) > 0:
            st.markdown('<div class="section-title">✅ الأصناف التي تحتاج إلى إضافة</div>', unsafe_allow_html=True)
            for idx, row in additions_df.iterrows():
                with st.container():
                    st.markdown(f"""
                    <div class="action-card">
                        <table style="width: 100%;">
                            <tr>
                                <td style="width: 15%"><strong>📋 رقم الطلب</strong><br>{row['order_number']}</td>
                                <td style="width: 12%"><strong>🏷️ SKU</strong><br>{row['sku']}</td>
                                <td style="width: 30%"><strong>📦 المنتج</strong><br>{row['product_name'][:40]}</td>
                                <td style="width: 10%"><strong>📊 الكمية</strong><br>{int(row['salla_qty'])}</td>
                                <td style="width: 10%"><strong>➕ الفرق</strong><br>+{int(row['difference'])}</td>
                                <td style="width: 23%">
                    """, unsafe_allow_html=True)
                    
                    if row['status'] == 'تم':
                        st.markdown(f'<span class="success-badge">✅ تمت الإضافة بواسطة {row["performed_by"][:15] if row["performed_by"] else ""}</span>', unsafe_allow_html=True)
                    else:
                        if st.button(f"➕ تمت الإضافة", key=f"add_{idx}"):
                            if record_adjustment(row['order_number'], row['sku'], pharmacy_name, 'إضافة'):
                                st.success("✅ تم!")
                                st.rerun()
                    
                    st.markdown(f"""
                                </td>
                            </tr>
                        </table>
                    </div>
                    """, unsafe_allow_html=True)
        
        # Returns
        returns_df = df[df['action'] == 'إرجاع']
        if len(returns_df) > 0:
            st.markdown('<div class="section-title">🔄 الأصناف التي تحتاج إلى إرجاع</div>', unsafe_allow_html=True)
            for idx, row in returns_df.iterrows():
                with st.container():
                    st.markdown(f"""
                    <div class="action-card">
                        <table style="width: 100%;">
                            <tr>
                                <td style="width: 12%"><strong>📋 رقم الطلب</strong><br>{row['order_number']}</td>
                                <td style="width: 15%"><strong>🧾 رقم الفاتورة</strong><br><span class="invoice-badge">{row['invoice_number'] if row['invoice_number'] else 'غير موجود'}</span></td>
                                <td style="width: 10%"><strong>🏷️ SKU</strong><br>{row['sku']}</td>
                                <td style="width: 25%"><strong>📦 المنتج</strong><br>{row['product_name'][:35]}</td>
                                <td style="width: 10%"><strong>📊 كمية ABC</strong><br>{int(row['abc_qty'])}</td>
                                <td style="width: 10%"><strong>📅 تاريخ الفاتورة</strong><br>{row['invoice_date'][:10] if row['invoice_date'] else 'غير محدد'}</td>
                                <td style="width: 18%">
                    """, unsafe_allow_html=True)
                    
                    if row['status'] == 'تم':
                        st.markdown(f'<span class="success-badge">✅ تم الإرجاع بواسطة {row["performed_by"][:15] if row["performed_by"] else ""}</span>', unsafe_allow_html=True)
                    else:
                        if st.button(f"🔄 تم الإرجاع", key=f"return_{idx}"):
                            if record_adjustment(row['order_number'], row['sku'], pharmacy_name, 'إرجاع'):
                                st.success("✅ تم!")
                                st.rerun()
                    
                    st.markdown(f"""
                                </td>
                            </tr>
                        </table>
                    </div>
                    """, unsafe_allow_html=True)
        
        if len(additions_df) == 0 and len(returns_df) == 0:
            st.success("🎉 لا توجد إضافات أو إرجاعات مطلوبة")
    else:
        st.info("📭 لا توجد طلبات مخصصة لهذا الفرع")

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666; padding: 1rem;">
    <p>نظام بلسم لإدارة الصيدليات | جميع الحقوق محفوظة © 2026</p>
</div>
""", unsafe_allow_html=True)
