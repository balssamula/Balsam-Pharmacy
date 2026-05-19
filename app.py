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
    }
    
    .stat-card {
        background: white;
        padding: 1.2rem;
        border-radius: 15px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        text-align: center;
        margin: 10px;
    }
    
    .stat-number {
        font-size: 2rem;
        font-weight: bold;
        color: #2a5298;
    }
    
    .success-badge {
        background-color: #28a745;
        color: white;
        padding: 5px 10px;
        border-radius: 20px;
        font-size: 0.8rem;
        text-align: center;
        display: inline-block;
    }
    
    .pending-badge {
        background-color: #ffc107;
        color: #333;
        padding: 5px 10px;
        border-radius: 20px;
        font-size: 0.8rem;
        text-align: center;
        display: inline-block;
    }
    
    .pharmacy-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1rem;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin: 1rem 0;
    }
    
    .stButton button {
        width: 100%;
        border-radius: 8px;
        transition: all 0.3s;
    }
    
    .stButton button:hover {
        transform: scale(1.02);
    }
    
    div[data-testid="stDataFrame"] {
        direction: rtl;
    }
    
    .dataframe th {
        background-color: #2a5298;
        color: white;
        padding: 10px;
        text-align: center;
    }
    
    /* تنسيق الجدول */
    .styled-table {
        border-collapse: collapse;
        width: 100%;
        direction: rtl;
    }
    
    .styled-table th {
        background-color: #2a5298;
        color: white;
        padding: 12px;
        text-align: center;
        border: 1px solid #ddd;
    }
    
    .styled-table td {
        padding: 10px;
        text-align: center;
        border: 1px solid #ddd;
    }
    
    .styled-table tr:nth-child(even) {
        background-color: #f2f2f2;
    }
</style>
""", unsafe_allow_html=True)

# ========== DATABASE SETUP ==========
os.makedirs("data", exist_ok=True)

def init_database():
    """تهيئة قاعدة البيانات"""
    conn = sqlite3.connect('data/pharmacy.db')
    c = conn.cursor()
    
    # Users table with pharmacist name
    c.execute('''CREATE TABLE IF NOT EXISTS users (
                 username TEXT PRIMARY KEY, 
                 password TEXT, 
                 role TEXT,
                 pharmacist_name TEXT,
                 last_login TEXT
             )''')
    
    # Adjustments table (orders needing action)
    c.execute('''CREATE TABLE IF NOT EXISTS adjustments (
                 id INTEGER PRIMARY KEY AUTOINCREMENT,
                 order_number TEXT, 
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
                 total_amount REAL
             )''')
    
    # Last login tracking
    c.execute('''CREATE TABLE IF NOT EXISTS last_access (
                 pharmacy_name TEXT PRIMARY KEY,
                 last_login TEXT,
                 pharmacist_name TEXT
             )''')
    
    # Insert default users (pharmacies 01-17)
    for i in range(1, 18):
        pharmacy_name = f"Balsam Alula Pharmacy {i:02d}"
        c.execute("INSERT OR IGNORE INTO users (username, password, role, pharmacist_name) VALUES (?, ?, 'pharmacy', ?)", 
                 (pharmacy_name, f"balsam{i}", ""))
    
    # Admin user
    c.execute("INSERT OR IGNORE INTO users (username, password, role, pharmacist_name) VALUES ('admin', 'admin123', 'admin', 'Manager')")
    
    conn.commit()
    conn.close()
    return [f"Balsam Alula Pharmacy {i:02d}" for i in range(1, 18)]

pharmacies_list = init_database()

# ========== HELPER FUNCTIONS ==========
def extract_branch_from_status(status_text):
    """استخراج رقم الفرع من حالة الطلب"""
    if not status_text or pd.isna(status_text):
        return None
    match = re.search(r'فرع\s*(\d+)', str(status_text))
    if match:
        return f"{int(match.group(1)):02d}"
    return None

def determine_branch(order_status, city):
    """تحديد الفرع بناءً على حالة الطلب والمدينة"""
    branch_num = extract_branch_from_status(order_status)
    if branch_num:
        return f"Balsam Alula Pharmacy {branch_num}", branch_num
    
    excluded_statuses = ['تم التوصيل', 'ملغي', 'مسترجع', 'محذوف', 'Delivered', 'Cancelled', 'Returned', 'Deleted']
    if any(s in str(order_status) for s in excluded_statuses):
        if city == 'AL ULA':
            return "Balsam Alula Pharmacy 09", "09"
        else:
            return "Balsam Alula Pharmacy 13", "13"
    
    return "Balsam Alula Pharmacy 13", "13"

def is_valid_sku(sku):
    """التحقق من صحة SKU"""
    if pd.isna(sku):
        return False
    sku_str = str(sku).strip()
    invalid_values = ['', '0', '1', '200', 'nan', 'NaN', 'None', 'null']
    if sku_str in invalid_values:
        return False
    return sku_str.isdigit()

def update_pharmacist_name(pharmacy_name, pharmacist_name):
    """تحديث اسم الصيدلي"""
    conn = sqlite3.connect('data/pharmacy.db')
    c = conn.cursor()
    c.execute("UPDATE users SET pharmacist_name = ? WHERE username = ?", (pharmacist_name, pharmacy_name))
    c.execute("INSERT OR REPLACE INTO last_access (pharmacy_name, last_login, pharmacist_name) VALUES (?, ?, ?)",
             (pharmacy_name, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), pharmacist_name))
    conn.commit()
    conn.close()

def get_pharmacist_name(pharmacy_name):
    """الحصول على اسم الصيدلي"""
    conn = sqlite3.connect('data/pharmacy.db')
    c = conn.cursor()
    c.execute("SELECT pharmacist_name FROM users WHERE username = ?", (pharmacy_name,))
    result = c.fetchone()
    conn.close()
    return result[0] if result and result[0] else ""

def get_last_login(pharmacy_name):
    """الحصول على آخر دخول للفرع"""
    conn = sqlite3.connect('data/pharmacy.db')
    c = conn.cursor()
    c.execute("SELECT last_login FROM last_access WHERE pharmacy_name = ?", (pharmacy_name,))
    result = c.fetchone()
    conn.close()
    return result[0] if result else "لم يدخل بعد"

def process_excel(uploaded_file):
    """معالجة ملف Excel وحفظ البيانات في قاعدة البيانات"""
    try:
        df_salla = pd.read_excel(uploaded_file, sheet_name="سلة")
        df_abc = pd.read_excel(uploaded_file, sheet_name="abc")
        
        # تنظيف بيانات سلة - إزالة الصفوف الفارغة
        df_salla = df_salla[df_salla['رقم الطلب'].notna()]
        
        # تصفية SKU الصالحة
        valid_skus = df_salla['SKU'].apply(is_valid_skus)
        df_salla = df_salla[valid_skus]
        
        # تحديد الفرع لكل طلب
        branch_info = df_salla.apply(
            lambda row: determine_branch(row['حالة الطلب'], row['المدينة']), 
            axis=1
        )
        df_salla['الفرع'] = branch_info.apply(lambda x: x[0])
        df_salla['رقم_الفرع'] = branch_info.apply(lambda x: x[1])
        
        # تجميع كميات سلة
        salla_grouped = df_salla.groupby(['رقم الطلب', 'SKU', 'اسم المنتج', 'الفرع', 'رقم_الفرع']).agg({
            'الكمية': 'sum',
            'حالة الطلب': 'first',
            'المدينة': 'first',
            'اسم العميل': 'first',
            'رقم الجوال': 'first',
            'تاريخ الطلب': 'first',
            'إجمالي الطلب': 'first'
        }).reset_index()
        
        # تجميع كميات ABC
        abc_grouped = df_abc.groupby(['رقم الطلب', 'رقم الصنف']).agg({
            'Net Sold Qty': 'sum',
            'الصيدلي': 'first'
        }).reset_index()
        abc_grouped.rename(columns={'رقم الصنف': 'SKU', 'Net Sold Qty': 'كمية_ABC'}, inplace=True)
        
        # دمج البيانات
        merged = pd.merge(salla_grouped, abc_grouped, on=['رقم الطلب', 'SKU'], how='outer')
        merged = merged.fillna(0)
        
        # حساب الفرق
        merged['الفرق'] = merged['الكمية'] - merged['كمية_ABC']
        merged['نوع_الاجراء'] = merged['الفرق'].apply(
            lambda x: 'إضافة' if x > 0 else ('إرجاع' if x < 0 else 'مطابق')
        )
        
        # تصفية الإضافات والإرجاعات فقط
        result = merged[merged['نوع_الاجراء'].isin(['إضافة', 'إرجاع'])].copy()
        
        # حفظ في قاعدة البيانات
        conn = sqlite3.connect('data/pharmacy.db')
        cursor = conn.cursor()
        
        # مسح البيانات القديمة (اختياري - يمكن الاحتفاظ بالتاريخ)
        # cursor.execute("DELETE FROM adjustments WHERE 1=1")
        
        for _, row in result.iterrows():
            # التحقق من وجود التعديل مسبقاً
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
                (order_number, sku, product_name, branch_number, pharmacist, pharmacy_name,
                 salla_qty, abc_qty, difference, action, status, performed_by, performed_at,
                 order_status, city, customer_name, customer_phone, order_date, total_amount, timestamp)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                str(row['رقم الطلب']), str(row['SKU']), row['اسم المنتج'],
                row['رقم_الفرع'], row.get('الصيدلي', ''), row['الفرع'],
                float(row['الكمية']), float(row['كمية_ABC']), float(row['الفرق']),
                row['نوع_الاجراء'], status, performed_by, performed_by,
                str(row['حالة الطلب']), str(row['المدينة']), str(row.get('اسم العميل', '')),
                str(row.get('رقم الجوال', '')), str(row.get('تاريخ الطلب', '')), 
                float(row.get('إجمالي الطلب', 0)), datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ))
        
        conn.commit()
        conn.close()
        
        return result
        
    except Exception as e:
        st.error(f"❌ خطأ في معالجة الملف: {str(e)}")
        return None

def record_adjustment(order_number, sku, pharmacy_name, action):
    """تسجيل التعديل (إضافة أو إرجاع)"""
    conn = sqlite3.connect('data/pharmacy.db')
    c = conn.cursor()
    
    # تحديث حالة التعديل
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

def get_pharmacy_data(pharmacy_name):
    """الحصول على بيانات الصيدلية"""
    conn = sqlite3.connect('data/pharmacy.db')
    df = pd.read_sql_query("""
        SELECT order_number, sku, product_name, salla_qty, abc_qty, difference, 
               action, status, order_status, city, customer_name, customer_phone, order_date
        FROM adjustments 
        WHERE pharmacy_name = ?
        ORDER BY order_number DESC
    """, conn, params=(pharmacy_name,))
    conn.close()
    return df

def get_all_adjustments():
    """الحصول على جميع التعديلات للمدير"""
    conn = sqlite3.connect('data/pharmacy.db')
    df = pd.read_sql_query("""
        SELECT order_number, sku, product_name, branch_number, pharmacy_name, 
               salla_qty, abc_qty, difference, action, status, performed_by, performed_at,
               order_status, city, customer_name, customer_phone, order_date
        FROM adjustments 
        ORDER BY order_number DESC
    """, conn)
    conn.close()
    return df

# ========== SESSION STATE ==========
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
if 'username' not in st.session_state:
    st.session_state.username = None
if 'user_role' not in st.session_state:
    st.session_state.user_role = None
if 'pharmacist_name' not in st.session_state:
    st.session_state.pharmacist_name = None
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None
if 'last_file' not in st.session_state:
    st.session_state.last_file = None

# ========== LOGIN ==========
with st.sidebar:
    st.image("https://img.icons8.com/color/96/000000/pharmacy.png", width=80)
    st.title("🌟 نظام بلسم")
    st.markdown("---")
    
    if not st.session_state.logged_in:
        st.subheader("🔐 تسجيل الدخول")
        username = st.text_input("👤 اسم المستخدم", key="login_user")
        password = st.text_input("🔒 كلمة المرور", type="password", key="login_pass")
        
        if st.button("🚪 دخول", use_container_width=True):
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
                    # طلب اسم الصيدلي إذا لم يكن مسجلاً
                    pharmacist_name = user[1] if user[1] else ""
                    st.session_state.pharmacist_name = pharmacist_name
                
                st.rerun()
            else:
                st.error("❌ اسم المستخدم أو كلمة المرور غير صحيحة")
    else:
        st.success(f"مرحباً {st.session_state.username}")
        
        if st.session_state.user_role == "admin":
            st.info("👑 مدير عام")
        else:
            st.info(f"🏥 {st.session_state.username}")
            # إظهار آخر دخول
            last_login = get_last_login(st.session_state.username)
            if last_login != "لم يدخل بعد":
                st.caption(f"📅 آخر دخول: {last_login}")
        
        st.markdown("---")
        
        if st.button("🚪 تسجيل خروج", use_container_width=True):
            for key in ['logged_in', 'username', 'user_role', 'pharmacist_name', 'processed_data']:
                st.session_state[key] = None
            st.rerun()

# ========== PHARMACIST NAME INPUT (FOR PHARMACY) ==========
if st.session_state.user_role == "pharmacy" and st.session_state.logged_in:
    if not st.session_state.pharmacist_name:
        st.markdown("### 👤 الرجاء إدخال اسمك")
        pharmacist_name_input = st.text_input("اسم الصيدلي")
        if st.button("تأكيد الاسم"):
            if pharmacist_name_input.strip():
                st.session_state.pharmacist_name = pharmacist_name_input
                update_pharmacist_name(st.session_state.username, pharmacist_name_input)
                st.success(f"✅ تم تسجيل اسمك: {pharmacist_name_input}")
                st.rerun()
            else:
                st.error("❌ الرجاء إدخال اسم صحيح")
        st.stop()

# ========== MAIN CONTENT ==========
if not st.session_state.logged_in:
    st.markdown("""
    <div class="main-header">
        <h1 style="font-size: 2.5rem;">📊 نظام بلسم لإدارة الصيدليات</h1>
        <p style="font-size: 1.2rem;">نظام متكامل لمراقبة طلبات سلة ومطابقتها مع فواتير ABC</p>
        <p style="font-size: 1rem; margin-top: 1rem;">✅ إدارة الفروع | ✅ متابعة الإضافات والإرجاعات | ✅ تقارير فورية</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-number">17</div>
            <div class="stat-label">🏥 فرع</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div class="stat-card">
            <div class="stat-number">1000+</div>
            <div class="stat-label">📦 طلب شهرياً</div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        all_adj = get_all_adjustments()
        pending = len(all_adj[all_adj['status'] != 'تم']) if len(all_adj) > 0 else 0
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-number">{pending}</div>
            <div class="stat-label">⏳ إجراءات معلقة</div>
        </div>
        """, unsafe_allow_html=True)
    with col4:
        completed = len(all_adj[all_adj['status'] == 'تم']) if len(all_adj) > 0 else 0
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-number">{completed}</div>
            <div class="stat-label">✅ تم إنجازها</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.info("👈 الرجاء تسجيل الدخول من القائمة الجانبية للاستمرار")

elif st.session_state.user_role == "admin":
    # ========== ADMIN DASHBOARD ==========
    st.markdown('<div class="main-header"><h1>👑 لوحة تحكم المدير العام</h1></div>', unsafe_allow_html=True)
    
    # File upload section
    with st.expander("📂 رفع ملف الطلبات والفواتير", expanded=True):
        uploaded_file = st.file_uploader("اختر ملف Excel", type=['xlsx'], key="admin_upload")
        
        if uploaded_file and st.session_state.last_file != uploaded_file.name:
            if st.button("🔄 معالجة الملف", use_container_width=True):
                with st.spinner("جاري معالجة الملف..."):
                    result = process_excel(uploaded_file)
                    if result is not None:
                        st.session_state.processed_data = result
                        st.session_state.last_file = uploaded_file.name
                        st.success("✅ تمت المعالجة بنجاح!")
                        st.balloons()
    
    # Get all adjustments
    df = get_all_adjustments()
    
    if len(df) > 0:
        # Statistics
        st.markdown("### 📊 إحصائيات سريعة")
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            st.metric("📋 إجمالي الطلبات", len(df['order_number'].unique()))
        with col2:
            additions = len(df[df['action'] == 'إضافة'])
            st.metric("➕ إضافات مطلوبة", additions)
        with col3:
            returns = len(df[df['action'] == 'إرجاع'])
            st.metric("➖ إرجاعات مطلوبة", returns)
        with col4:
            completed = len(df[df['status'] == 'تم'])
            st.metric("✅ تم إنجازها", completed)
        with col5:
            pending = len(df[df['status'] != 'تم'])
            st.metric("⏳ قيد الانتظار", pending)
        
        st.markdown("---")
        
        # Tabs
        tab1, tab2, tab3 = st.tabs(["📈 الإضافات", "📉 الإرجاعات", "🏥 أداء الفروع"])
        
        with tab1:
            additions_df = df[df['action'] == 'إضافة'].copy()
            if len(additions_df) > 0:
                # إضافة عمود الحالة
                additions_df['حالة الإجراء'] = additions_df['status'].apply(
                    lambda x: "✅ تمت الإضافة" if x == "تم" else "⏳ لم تبدأ"
                )
                display_cols = ['order_number', 'sku', 'product_name', 'salla_qty', 'abc_qty', 
                               'difference', 'pharmacy_name', 'حالة الإجراء', 'performed_by', 'performed_at',
                               'order_status', 'city', 'customer_name']
                st.dataframe(additions_df[display_cols], use_container_width=True)
            else:
                st.success("🎉 لا توجد إضافات مطلوبة!")
        
        with tab2:
            returns_df = df[df['action'] == 'إرجاع'].copy()
            if len(returns_df) > 0:
                returns_df['حالة الإجراء'] = returns_df['status'].apply(
                    lambda x: "✅ تم الإرجاع" if x == "تم" else "⏳ لم يبدأ"
                )
                display_cols = ['order_number', 'sku', 'product_name', 'salla_qty', 'abc_qty', 
                               'difference', 'pharmacy_name', 'حالة الإجراء', 'performed_by', 'performed_at',
                               'order_status', 'city', 'customer_name']
                st.dataframe(returns_df[display_cols], use_container_width=True)
            else:
                st.success("🎉 لا توجد إرجاعات مطلوبة!")
        
        with tab3:
            st.subheader("🏥 أداء الفروع")
            branch_stats = df.groupby(['pharmacy_name', 'branch_number']).agg({
                'order_number': 'nunique',
                'action': 'count',
                'status': lambda x: (x == 'تم').sum()
            }).reset_index()
            branch_stats.columns = ['اسم الفرع', 'رقم الفرع', 'عدد الطلبات', 'عدد الإجراءات', 'تم إنجازها']
            branch_stats['المتبقي'] = branch_stats['عدد الإجراءات'] - branch_stats['تم إنجازها']
            
            # إضافة آخر دخول لكل فرع
            branch_stats['آخر دخول'] = branch_stats['اسم الفرع'].apply(get_last_login)
            
            st.dataframe(branch_stats, use_container_width=True)
    else:
        st.info("📂 لا توجد بيانات. الرجاء رفع ملف Excel أولاً")

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
    
    # Upload file for pharmacy (optional - data already loaded by admin)
    with st.expander("📂 رفع ملف الطلبات والفواتير (إذا لم يرفعه المدير)", expanded=False):
        uploaded_file = st.file_uploader("اختر ملف Excel", type=['xlsx'], key="pharmacy_upload")
        
        if uploaded_file and st.session_state.last_file != uploaded_file.name:
            if st.button("🔄 معالجة الملف", use_container_width=True):
                with st.spinner("جاري معالجة الملف..."):
                    result = process_excel(uploaded_file)
                    if result is not None:
                        st.session_state.processed_data = result
                        st.session_state.last_file = uploaded_file.name
                        st.success("✅ تمت المعالجة بنجاح!")
                        st.rerun()
    
    # Get pharmacy data
    df = get_pharmacy_data(pharmacy_name)
    
    if len(df) > 0:
        # Statistics
        col1, col2, col3 = st.columns(3)
        with col1:
            additions = len(df[df['action'] == 'إضافة'])
            st.metric("➕ إضافات مطلوبة", additions)
        with col2:
            returns = len(df[df['action'] == 'إرجاع'])
            st.metric("➖ إرجاعات مطلوبة", returns)
        with col3:
            completed = len(df[df['status'] == 'تم'])
            st.metric("✅ تم إنجازها", completed)
        
        st.markdown("---")
        
        # Additions section
        additions_df = df[df['action'] == 'إضافة'].copy()
        if len(additions_df) > 0:
            st.subheader("✅ الأصناف التي تحتاج إلى إضافة")
            for idx, row in additions_df.iterrows():
                with st.container():
                    st.markdown(f"""
                    <table class="styled-table" style="width: 100%; margin-bottom: 10px;">
                        <tr>
                            <th style="width: 12%">رقم الطلب</th>
                            <th style="width: 10%">SKU</th>
                            <th style="width: 25%">اسم المنتج</th>
                            <th style="width: 8%">كمية سلة</th>
                            <th style="width: 8%">كمية ABC</th>
                            <th style="width: 8%">الفرق</th>
                            <th style="width: 10%">حالة الطلب</th>
                            <th style="width: 10%">المدينة</th>
                            <th style="width: 9%">الإجراء</th>
                        </tr>
                        <tr>
                            <td style="text-align: center">{row['order_number']}</td>
                            <td style="text-align: center">{row['sku']}</td>
                            <td style="text-align: center">{row['product_name'][:40]}</td>
                            <td style="text-align: center">{int(row['salla_qty'])}</td>
                            <td style="text-align: center">{int(row['abc_qty'])}</td>
                            <td style="text-align: center; color: {'#28a745' if row['difference'] > 0 else '#dc3545'}">+{int(row['difference'])}</td>
                            <td style="text-align: center">{row['order_status'][:20]}</td>
                            <td style="text-align: center">{row['city']}</td>
                            <td style="text-align: center">
                    """, unsafe_allow_html=True)
                    
                    if row['status'] == 'تم':
                        st.markdown(f'<span class="success-badge">✅ تمت الإضافة بواسطة {row["performed_by"] if row["performed_by"] else "غير معروف"}</span>', unsafe_allow_html=True)
                    else:
                        if st.button(f"➕ تمت الإضافة", key=f"add_{idx}_{row['sku']}"):
                            if record_adjustment(row['order_number'], row['sku'], pharmacy_name, 'إضافة'):
                                st.success("✅ تم تسجيل الإضافة بنجاح!")
                                st.rerun()
                    
                    st.markdown(f"""
                            </td>
                        </tr>
                    </table>
                    """, unsafe_allow_html=True)
        
        # Returns section
        returns_df = df[df['action'] == 'إرجاع'].copy()
        if len(returns_df) > 0:
            st.subheader("🔄 الأصناف التي تحتاج إلى إرجاع")
            for idx, row in returns_df.iterrows():
                with st.container():
                    st.markdown(f"""
                    <table class="styled-table" style="width: 100%; margin-bottom: 10px;">
                        <tr>
                            <th style="width: 12%">رقم الطلب</th>
                            <th style="width: 10%">SKU</th>
                            <th style="width: 25%">اسم المنتج</th>
                            <th style="width: 8%">كمية سلة</th>
                            <th style="width: 8%">كمية ABC</th>
                            <th style="width: 8%">الفرق</th>
                            <th style="width: 10%">حالة الطلب</th>
                            <th style="width: 10%">المدينة</th>
                            <th style="width: 9%">الإجراء</th>
                        </tr>
                        <tr>
                            <td style="text-align: center">{row['order_number']}</td>
                            <td style="text-align: center">{row['sku']}</td>
                            <td style="text-align: center">{row['product_name'][:40]}</td>
                            <td style="text-align: center">{int(row['salla_qty'])}</td>
                            <td style="text-align: center">{int(row['abc_qty'])}</td>
                            <td style="text-align: center; color: {'#28a745' if row['difference'] > 0 else '#dc3545'}">{int(row['difference'])}</td>
                            <td style="text-align: center">{row['order_status'][:20]}</td>
                            <td style="text-align: center">{row['city']}</td>
                            <td style="text-align: center">
                    """, unsafe_allow_html=True)
                    
                    if row['status'] == 'تم':
                        st.markdown(f'<span class="success-badge">✅ تم الإرجاع بواسطة {row["performed_by"] if row["performed_by"] else "غير معروف"}</span>', unsafe_allow_html=True)
                    else:
                        if st.button(f"🔄 تم الإرجاع", key=f"return_{idx}_{row['sku']}"):
                            if record_adjustment(row['order_number'], row['sku'], pharmacy_name, 'إرجاع'):
                                st.success("✅ تم تسجيل الإرجاع بنجاح!")
                                st.rerun()
                    
                    st.markdown(f"""
                            </td>
                        </tr>
                    </table>
                    """, unsafe_allow_html=True)
        
        if len(additions_df) == 0 and len(returns_df) == 0:
            st.info("🎉 لا توجد إضافات أو إرجاعات مطلوبة لهذا الفرع")
    else:
        st.info("📭 لا توجد طلبات مخصصة لهذا الفرع. الرجاء التواصل مع المدير لرفع ملف Excel.")

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666; padding: 1rem;">
    <p>نظام بلسم لإدارة الصيدليات | جميع الحقوق محفوظة © 2026</p>
</div>
""", unsafe_allow_html=True)
