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
    
    .pharmacy-selector {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        padding: 1.5rem;
        border-radius: 15px;
        text-align: center;
        cursor: pointer;
        transition: all 0.3s ease;
        margin: 10px;
        border: 2px solid transparent;
    }
    
    .pharmacy-selector:hover {
        transform: translateY(-5px);
        border-color: #2a5298;
        box-shadow: 0 10px 20px rgba(0,0,0,0.1);
    }
    
    .pharmacy-selector.selected {
        border-color: #28a745;
        background: #e8f5e9;
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
    
    .stat-label {
        font-size: 0.9rem;
        color: #666;
        margin-top: 5px;
    }
    
    .success-badge {
        background-color: #28a745;
        color: white;
        padding: 5px 10px;
        border-radius: 20px;
        font-size: 0.8rem;
    }
    
    .warning-badge {
        background-color: #ffc107;
        color: #333;
        padding: 5px 10px;
        border-radius: 20px;
        font-size: 0.8rem;
    }
    
    .info-badge {
        background-color: #17a2b8;
        color: white;
        padding: 5px 10px;
        border-radius: 20px;
        font-size: 0.8rem;
    }
    
    .pharmacy-title {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1rem;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin: 1rem 0;
    }
    
    .action-button {
        width: 100%;
        padding: 0.5rem;
        border-radius: 8px;
        font-weight: bold;
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
    
    /* تنسيق الأعمدة */
    .dataframe th {
        background-color: #2a5298;
        color: white;
        padding: 10px;
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)

# ========== DATABASE SETUP ==========
os.makedirs("data", exist_ok=True)

def init_database():
    conn = sqlite3.connect('data/pharmacy.db')
    c = conn.cursor()
    
    # Users table
    c.execute('''CREATE TABLE IF NOT EXISTS users
                 (username TEXT PRIMARY KEY, password TEXT, role TEXT, pharmacy_name TEXT)''')
    
    # Adjustments table
    c.execute('''CREATE TABLE IF NOT EXISTS adjustments
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  order_number TEXT, sku TEXT, product_name TEXT,
                  branch_number TEXT, pharmacist TEXT, pharmacy_name TEXT,
                  salla_qty REAL, abc_qty REAL, difference REAL,
                  action TEXT, status TEXT, performed_by TEXT,
                  timestamp TEXT, order_status TEXT, city TEXT)''')
    
    # Orders summary table
    c.execute('''CREATE TABLE IF NOT EXISTS orders_summary
                 (order_number TEXT PRIMARY KEY,
                  customer_name TEXT, phone TEXT, city TEXT,
                  order_status TEXT, branch_number TEXT, pharmacist TEXT,
                  order_date TEXT, total_amount REAL, processed BOOLEAN DEFAULT 0)''')
    
    # Insert default users (17 pharmacies + admin)
    pharmacies = []
    for i in range(1, 18):
        pharmacy_name = f"Balsam Alula Pharmacy {i:02d}"
        pharmacies.append(pharmacy_name)
        c.execute("INSERT OR IGNORE INTO users VALUES (?, ?, 'pharmacy', ?)", 
                 (pharmacy_name, f"balsam{i}", pharmacy_name))
    
    # Admin user
    c.execute("INSERT OR IGNORE INTO users VALUES ('admin', 'admin123', 'admin', NULL)")
    
    conn.commit()
    conn.close()
    return pharmacies

pharmacies_list = init_database()

# ========== HELPER FUNCTIONS ==========
def extract_branch_from_status(status_text):
    """استخراج رقم الفرع من حالة الطلب مثل 'تم الاستلام من فرع 14'"""
    if not status_text:
        return None
    match = re.search(r'فرع\s*(\d+)', str(status_text))
    if match:
        return f"{int(match.group(1)):02d}"
    return None

def determine_branch(order_status, city):
    """تحديد الفرع بناءً على حالة الطلب والمدينة"""
    # استخراج رقم الفرع من الحالة إذا كان موجوداً
    branch_num = extract_branch_from_status(order_status)
    if branch_num:
        return f"Balsam Alula Pharmacy {branch_num}", branch_num
    
    # إذا كانت الحالة تم التوصيل أو ملغي أو مسترجع أو محذوف
    excluded_statuses = ['تم التوصيل', 'ملغي', 'مسترجع', 'محذوف', 'Delivered', 'Cancelled', 'Returned', 'Deleted']
    if any(s in str(order_status) for s in excluded_statuses):
        if city == 'AL ULA':
            return "Balsam Alula Pharmacy 09", "09"
        else:
            return "Balsam Alula Pharmacy 13", "13"
    
    # افتراضي
    return "Balsam Alula Pharmacy 13", "13"

def is_valid_sku(sku):
    """التحقق من صحة SKU"""
    if pd.isna(sku):
        return False
    sku_str = str(sku).strip()
    # استبعاد القيم غير الصالحة
    invalid_values = ['', '0', '1', '200', 'nan', 'NaN', 'None']
    if sku_str in invalid_values:
        return False
    # يجب أن يكون رقماً
    return sku_str.isdigit()

def process_excel(uploaded_file):
    """معالجة ملف Excel مع كل الشروط"""
    try:
        # قراءة الشيتات
        df_salla = pd.read_excel(uploaded_file, sheet_name="سلة")
        df_abc = pd.read_excel(uploaded_file, sheet_name="abc")
        
        # ===== معالجة شيت سلة =====
        # حذف الصفوف المكررة (حيث رقم الطلب فارغ)
        df_salla = df_salla[df_salla['رقم الطلب'].notna()]
        
        # تصفية حسب SKU الصالحة
        df_salla = df_salla[df_salla['SKU'].apply(is_valid_sku)]
        
        # تصفية الحالات المستثناة (ملغي، مسترجع، محذوف)
        # لكن ملاحظة: هذه الحالات يتم توجيهها إلى فرع 9 أو 13 وليس حذفها
        # حسب فهمي، لا يتم حذفها بل يتم توجيهها
        
        # إنشاء عمود الفرع ورقم الفرع
        branch_info = df_salla.apply(
            lambda row: determine_branch(row['حالة الطلب'], row['المدينة']), 
            axis=1
        )
        df_salla['الفرع'] = branch_info.apply(lambda x: x[0])
        df_salla['رقم_الفرع'] = branch_info.apply(lambda x: x[1])
        
        # إضافة عمود الصيدلي (من شيت abc سيتم ربطه لاحقاً)
        df_salla['الصيدلي'] = ''
        
        # تجميع كميات سلة حسب رقم الطلب و SKU والفرع
        salla_grouped = df_salla.groupby(['رقم الطلب', 'SKU', 'اسم المنتج', 'الفرع', 'رقم_الفرع']).agg({
            'الكمية': 'sum',
            'حالة الطلب': 'first',
            'المدينة': 'first',
            'اسم العميل': 'first',
            'رقم الجوال': 'first',
            'تاريخ الطلب': 'first',
            'إجمالي الطلب': 'first'
        }).reset_index()
        
        # ===== معالجة شيت ABC =====
        # تجميع كميات ABC
        abc_grouped = df_abc.groupby(['رقم الطلب', 'رقم الصنف']).agg({
            'Net Sold Qty': 'sum',
            'الصيدلي': 'first',
            'رقم الصيدلية': 'first'
        }).reset_index()
        abc_grouped.rename(columns={
            'رقم الصنف': 'SKU', 
            'Net Sold Qty': 'كمية_ABC',
            'رقم الصيدلية': 'رقم_الصيدلية'
        }, inplace=True)
        
        # ربط بيانات الصيدلي من abc إلى سلة
        pharmacist_map = abc_grouped.groupby('رقم الطلب').agg({
            'الصيدلي': lambda x: ' / '.join(x.unique())
        }).reset_index()
        
        # دمج البيانات
        merged = pd.merge(salla_grouped, abc_grouped, on=['رقم الطلب', 'SKU'], how='outer')
        merged = merged.fillna(0)
        
        # إضافة أسماء الصيادلة
        merged = pd.merge(merged, pharmacist_map, on='رقم الطلب', how='left')
        if 'الصيدلي_y' in merged.columns:
            merged['الصيدلي'] = merged['الصيدلي_y'].fillna('')
        elif 'الصيدلي' in merged.columns:
            pass
        else:
            merged['الصيدلي'] = ''
        
        # حساب الفرق
        merged['الفرق'] = merged['الكمية'] - merged['كمية_ABC']
        merged['نوع_الاجراء'] = merged['الفرق'].apply(
            lambda x: 'إضافة' if x > 0 else ('إرجاع' if x < 0 else 'مطابق')
        )
        
        # تصفية فقط الإضافات والإرجاعات
        result = merged[merged['نوع_الاجراء'].isin(['إضافة', 'إرجاع'])].copy()
        
        # إعادة تسمية الأعمدة للعرض
        result = result.rename(columns={
            'رقم الطلب': 'رقم الطلب',
            'SKU': 'SKU',
            'اسم المنتج': 'اسم المنتج',
            'الكمية': 'كمية سلة',
            'كمية_ABC': 'كمية ABC',
            'الفرق': 'الفرق',
            'نوع_الاجراء': 'نوع الإجراء',
            'الفرع': 'الفرع',
            'رقم_الفرع': 'رقم الفرع',
            'الصيدلي': 'الصيدلي',
            'حالة الطلب': 'حالة الطلب',
            'المدينة': 'المدينة',
            'اسم العميل': 'اسم العميل',
            'رقم الجوال': 'رقم الجوال',
            'تاريخ الطلب': 'تاريخ الطلب',
            'إجمالي الطلب': 'إجمالي الطلب'
        })
        
        # حفظ البيانات في قاعدة البيانات
        conn = sqlite3.connect('data/pharmacy.db')
        cursor = conn.cursor()
        
        for _, row in result.iterrows():
            # التحقق من وجود التعديل مسبقاً
            cursor.execute("""
                SELECT status FROM adjustments 
                WHERE order_number=? AND sku=? AND pharmacy_name=?
            """, (str(row['رقم الطلب']), str(row['SKU']), row['الفرع']))
            
            existing = cursor.fetchone()
            status = existing[0] if existing else "لم يبدأ"
            
            cursor.execute("""
                INSERT OR REPLACE INTO adjustments 
                (order_number, sku, product_name, branch_number, pharmacist, pharmacy_name,
                 salla_qty, abc_qty, difference, action, status, order_status, city, timestamp)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                str(row['رقم الطلب']), str(row['SKU']), row['اسم المنتج'],
                row['رقم الفرع'], row.get('الصيدلي', ''), row['الفرع'],
                float(row['كمية سلة']), float(row['كمية ABC']), float(row['الفرق']),
                row['نوع الإجراء'], status, row['حالة الطلب'], row['المدينة'],
                datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            ))
            
            # حفظ ملخص الطلب
            cursor.execute("""
                INSERT OR REPLACE INTO orders_summary 
                (order_number, customer_name, phone, city, order_status, branch_number, pharmacist, order_date, total_amount)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                str(row['رقم الطلب']), row.get('اسم العميل', ''), str(row.get('رقم الجوال', '')),
                row.get('المدينة', ''), row.get('حالة الطلب', ''), row['رقم الفرع'],
                row.get('الصيدلي', ''), str(row.get('تاريخ الطلب', '')), float(row.get('إجمالي الطلب', 0))
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
        SET status = 'تم', performed_by = ?, timestamp = ?
        WHERE order_number = ? AND sku = ? AND pharmacy_name = ? AND action = ?
    """, (st.session_state.username, datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
          str(order_number), str(sku), pharmacy_name, action))
    
    # إذا لم يتم التحديث، أدخل سجلاً جديداً
    if c.rowcount == 0:
        c.execute("""
            INSERT INTO adjustments 
            (order_number, sku, pharmacy_name, action, status, performed_by, timestamp)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (str(order_number), str(sku), pharmacy_name, action, 'تم', 
              st.session_state.username, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    
    conn.commit()
    conn.close()

def get_adjustment_status(order_number, sku, pharmacy_name, action):
    """الحصول على حالة التعديل"""
    conn = sqlite3.connect('data/pharmacy.db')
    c = conn.cursor()
    c.execute("""
        SELECT status FROM adjustments 
        WHERE order_number = ? AND sku = ? AND pharmacy_name = ? AND action = ?
        ORDER BY timestamp DESC LIMIT 1
    """, (str(order_number), str(sku), pharmacy_name, action))
    result = c.fetchone()
    conn.close()
    return result[0] if result else "لم يبدأ"

# ========== SESSION STATE ==========
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
if 'username' not in st.session_state:
    st.session_state.username = None
if 'user_role' not in st.session_state:
    st.session_state.user_role = None
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None
if 'selected_pharmacy' not in st.session_state:
    st.session_state.selected_pharmacy = None

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
            c.execute("SELECT role FROM users WHERE username=? AND password=?", (username, password))
            user = c.fetchone()
            conn.close()
            
            if user:
                st.session_state.logged_in = True
                st.session_state.username = username
                st.session_state.user_role = user[0]
                st.rerun()
            else:
                st.error("❌ اسم المستخدم أو كلمة المرور غير صحيحة")
    else:
        st.success(f"مرحباً {st.session_state.username}")
        
        # عرض معلومات المستخدم
        if st.session_state.user_role == "admin":
            st.info("👑 أنت مسجل كمدير عام")
        else:
            st.info(f"🏥 الصيدلية: {st.session_state.username}")
        
        st.markdown("---")
        
        if st.button("🚪 تسجيل خروج", use_container_width=True):
            for key in ['logged_in', 'username', 'user_role', 'processed_data', 'selected_pharmacy']:
                st.session_state[key] = None
            st.rerun()

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
        st.markdown("""
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
        st.markdown("""
        <div class="stat-card">
            <div class="stat-number">99%</div>
            <div class="stat-label">⚡ دقة المطابقة</div>
        </div>
        """, unsafe_allow_html=True)
    with col4:
        st.markdown("""
        <div class="stat-card">
            <div class="stat-number">24/7</div>
            <div class="stat-label">🕐 دعم مستمر</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.info("👈 الرجاء تسجيل الدخول من القائمة الجانبية للاستمرار")

elif st.session_state.user_role == "admin":
    # ========== ADMIN DASHBOARD ==========
    st.markdown('<div class="main-header"><h1>👑 لوحة تحكم المدير العام</h1></div>', unsafe_allow_html=True)
    
    # Upload section
    with st.expander("📂 رفع ملف الطلبات والفواتير", expanded=True):
        uploaded_file = st.file_uploader("اختر ملف Excel", type=['xlsx'], key="admin_upload")
        
        if uploaded_file:
            if st.button("🔄 معالجة الملف", use_container_width=True):
                with st.spinner("جاري معالجة الملف..."):
                    result = process_excel(uploaded_file)
                    if result is not None:
                        st.session_state.processed_data = result
                        st.success("✅ تمت المعالجة بنجاح!")
                        st.balloons()
    
    if st.session_state.processed_data is not None:
        df = st.session_state.processed_data
        
        # Statistics
        st.markdown("### 📊 إحصائيات سريعة")
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            st.metric("📋 إجمالي الطلبات", len(df['رقم الطلب'].unique()))
        with col2:
            additions = len(df[df['نوع الإجراء'] == 'إضافة'])
            st.metric("➕ إضافات مطلوبة", additions)
        with col3:
            returns = len(df[df['نوع الإجراء'] == 'إرجاع'])
            st.metric("➖ إرجاعات مطلوبة", returns)
        with col4:
            completed = len(df[df['status'] == 'تم']) if 'status' in df.columns else 0
            st.metric("✅ تم إنجازها", completed)
        with col5:
            pending = additions + returns - completed
            st.metric("⏳ قيد الانتظار", pending)
        
        st.markdown("---")
        
        # Tabs for different views
        tab1, tab2, tab3, tab4 = st.tabs(["📈 الإضافات", "📉 الإرجاعات", "🏥 حسب الفرع", "📋 سجل التعديلات"])
        
        with tab1:
            additions_df = df[df['نوع الإجراء'] == 'إضافة']
            if len(additions_df) > 0:
                display_cols = ['رقم الطلب', 'SKU', 'اسم المنتج', 'كمية سلة', 'كمية ABC', 'الفرق', 'رقم الفرع', 'الصيدلي', 'حالة الطلب', 'المدينة']
                st.dataframe(additions_df[display_cols], use_container_width=True)
            else:
                st.success("🎉 لا توجد إضافات مطلوبة!")
        
        with tab2:
            returns_df = df[df['نوع الإجراء'] == 'إرجاع']
            if len(returns_df) > 0:
                display_cols = ['رقم الطلب', 'SKU', 'اسم المنتج', 'كمية سلة', 'كمية ABC', 'الفرق', 'رقم الفرع', 'الصيدلي', 'حالة الطلب', 'المدينة']
                st.dataframe(returns_df[display_cols], use_container_width=True)
            else:
                st.success("🎉 لا توجد إرجاعات مطلوبة!")
        
        with tab3:
            st.subheader("🏥 توزيع الطلبات حسب الفرع")
            branch_stats = df.groupby(['رقم الفرع', 'الفرع']).agg({
                'رقم الطلب': 'nunique',
                'نوع الإجراء': 'count'
            }).reset_index()
            branch_stats.columns = ['رقم الفرع', 'اسم الفرع', 'عدد الطلبات', 'عدد الإجراءات']
            st.dataframe(branch_stats, use_container_width=True)
        
        with tab4:
            st.subheader("📋 سجل التعديلات")
            conn = sqlite3.connect('data/pharmacy.db')
            history = pd.read_sql_query("""
                SELECT order_number, sku, product_name, pharmacy_name, action, status, performed_by, timestamp 
                FROM adjustments ORDER BY timestamp DESC LIMIT 100
            """, conn)
            conn.close()
            if len(history) > 0:
                st.dataframe(history, use_container_width=True)
            else:
                st.info("لا توجد تعديلات مسجلة بعد")
    
    # Select pharmacy to view details
    st.markdown("---")
    st.subheader("🏥 متابعة أداء الصيدليات")
    
    col1, col2 = st.columns([1, 2])
    with col1:
        selected_pharmacy = st.selectbox("اختر الفرع", pharmacies_list)
    
    if selected_pharmacy:
        with col2:
            conn = sqlite3.connect('data/pharmacy.db')
            branch_orders = pd.read_sql_query("""
                SELECT COUNT(DISTINCT order_number) as orders_count,
                       SUM(CASE WHEN action='إضافة' THEN 1 ELSE 0 END) as additions_count,
                       SUM(CASE WHEN action='إرجاع' THEN 1 ELSE 0 END) as returns_count,
                       SUM(CASE WHEN status='تم' THEN 1 ELSE 0 END) as completed_count
                FROM adjustments WHERE pharmacy_name=?
            """, conn, params=(selected_pharmacy,))
            conn.close()
            
            if len(branch_orders) > 0:
                c1, c2, c3, c4 = st.columns(4)
                with c1:
                    st.metric("📦 طلبات", int(branch_orders['orders_count'].iloc[0]))
                with c2:
                    st.metric("➕ إضافات", int(branch_orders['additions_count'].iloc[0]))
                with c3:
                    st.metric("➖ إرجاعات", int(branch_orders['returns_count'].iloc[0]))
                with c4:
                    st.metric("✅ منجز", int(branch_orders['completed_count'].iloc[0]))

else:
    # ========== PHARMACY DASHBOARD ==========
    pharmacy_name = st.session_state.username
    
    st.markdown(f"""
    <div class="pharmacy-title">
        <h1>🏥 {pharmacy_name}</h1>
        <p>مراجعة وإدارة الإضافات والإرجاعات</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Upload file for pharmacy
    with st.expander("📂 رفع ملف الطلبات والفواتير", expanded=True):
        uploaded_file = st.file_uploader("اختر ملف Excel", type=['xlsx'], key="pharmacy_upload")
        
        if uploaded_file:
            if st.button("🔄 معالجة الملف", use_container_width=True):
                with st.spinner("جاري معالجة الملف..."):
                    result = process_excel(uploaded_file)
                    if result is not None:
                        st.session_state.processed_data = result
                        st.success("✅ تمت المعالجة بنجاح!")
                        st.balloons()
    
    if st.session_state.processed_data is not None:
        df = st.session_state.processed_data
        
        # Filter data for this pharmacy
        pharmacy_df = df[df['الفرع'] == pharmacy_name]
        
        if len(pharmacy_df) > 0:
            # Statistics for this pharmacy
            col1, col2, col3 = st.columns(3)
            with col1:
                additions = len(pharmacy_df[pharmacy_df['نوع الإجراء'] == 'إضافة'])
                st.metric("➕ إضافات مطلوبة", additions)
            with col2:
                returns = len(pharmacy_df[pharmacy_df['نوع الإجراء'] == 'إرجاع'])
                st.metric("➖ إرجاعات مطلوبة", returns)
            with col3:
                # حساب المنجز
                conn = sqlite3.connect('data/pharmacy.db')
                completed = pd.read_sql_query("""
                    SELECT COUNT(*) as completed FROM adjustments 
                    WHERE pharmacy_name=? AND status='تم'
                """, conn, params=(pharmacy_name,))
                conn.close()
                st.metric("✅ تم إنجازها", completed['completed'].iloc[0] if len(completed) > 0 else 0)
            
            st.markdown("---")
            
            # Additions section
            additions_df = pharmacy_df[pharmacy_df['نوع الإجراء'] == 'إضافة']
            if len(additions_df) > 0:
                st.subheader("✅ الأصناف التي تحتاج إلى إضافة")
                for idx, row in additions_df.iterrows():
                    with st.container():
                        cols = st.columns([1.5, 1.5, 3, 1, 1, 1.5])
                        
                        # Order number
                        cols[0].markdown(f"**📋 رقم الطلب**<br>{row['رقم الطلب']}", unsafe_allow_html=True)
                        
                        # SKU
                        cols[1].markdown(f"**🏷️ SKU**<br>{row['SKU']}", unsafe_allow_html=True)
                        
                        # Product name
                        product_name = row['اسم المنتج'][:40] + "..." if len(str(row['اسم المنتج'])) > 40 else row['اسم المنتج']
                        cols[2].markdown(f"**📦 المنتج**<br>{product_name}", unsafe_allow_html=True)
                        
                        # Quantities
                        cols[3].markdown(f"**📊 الكمية**<br>{int(row['كمية سلة'])}", unsafe_allow_html=True)
                        
                        # Difference
                        cols[4].markdown(f"**➕ الفرق**<br>+{int(row['الفرق'])}", unsafe_allow_html=True)
                        
                        # Action button
                        status = get_adjustment_status(row['رقم الطلب'], row['SKU'], pharmacy_name, 'إضافة')
                        if status == 'تم':
                            cols[5].markdown('<div class="success-badge" style="text-align: center;">✅ تمت الإضافة</div>', unsafe_allow_html=True)
                        else:
                            if cols[5].button("➕ تمت الإضافة", key=f"add_{row['رقم الطلب']}_{row['SKU']}"):
                                record_adjustment(row['رقم الطلب'], row['SKU'], pharmacy_name, 'إضافة')
                                st.success("✅ تم تسجيل الإضافة بنجاح!")
                                st.rerun()
                        st.divider()
            else:
                st.info("🎉 لا توجد إضافات مطلوبة لهذا الفرع")
            
            # Returns section
            returns_df = pharmacy_df[pharmacy_df['نوع الإجراء'] == 'إرجاع']
            if len(returns_df) > 0:
                st.subheader("🔄 الأصناف التي تحتاج إلى إرجاع")
                for idx, row in returns_df.iterrows():
                    with st.container():
                        cols = st.columns([1.5, 1.5, 3, 1, 1, 1.5])
                        
                        # Order number
                        cols[0].markdown(f"**📋 رقم الطلب**<br>{row['رقم الطلب']}", unsafe_allow_html=True)
                        
                        # SKU
                        cols[1].markdown(f"**🏷️ SKU**<br>{row['SKU']}", unsafe_allow_html=True)
                        
                        # Product name
                        product_name = row['اسم المنتج'][:40] + "..." if len(str(row['اسم المنتج'])) > 40 else row['اسم المنتج']
                        cols[2].markdown(f"**📦 المنتج**<br>{product_name}", unsafe_allow_html=True)
                        
                        # Quantities
                        cols[3].markdown(f"**📊 الكمية في ABC**<br>{int(row['كمية ABC'])}", unsafe_allow_html=True)
                        
                        # Difference
                        cols[4].markdown(f"**➖ الفرق**<br>{int(row['الفرق'])}", unsafe_allow_html=True)
                        
                        # Action button
                        status = get_adjustment_status(row['رقم الطلب'], row['SKU'], pharmacy_name, 'إرجاع')
                        if status == 'تم':
                            cols[5].markdown('<div class="success-badge" style="text-align: center;">✅ تم الإرجاع</div>', unsafe_allow_html=True)
                        else:
                            if cols[5].button("🔄 تم الإرجاع", key=f"return_{row['رقم الطلب']}_{row['SKU']}"):
                                record_adjustment(row['رقم الطلب'], row['SKU'], pharmacy_name, 'إرجاع')
                                st.success("✅ تم تسجيل الإرجاع بنجاح!")
                                st.rerun()
                        st.divider()
            else:
                st.info("🎉 لا توجد إرجاعات مطلوبة لهذا الفرع")
            
            # Show order summary
            st.markdown("---")
            with st.expander("📋 عرض ملخص الطلبات"):
                summary_cols = ['رقم الطلب', 'اسم العميل', 'المدينة', 'حالة الطلب', 'رقم الفرع', 'إجمالي الطلب']
                available_cols = [c for c in summary_cols if c in pharmacy_df.columns]
                st.dataframe(pharmacy_df[available_cols].drop_duplicates('رقم الطلب'), use_container_width=True)
        
        else:
            st.info(f"📭 لا توجد طلبات مخصصة لـ {pharmacy_name} حالياً")
    else:
        st.info("📂 الرجاء رفع ملف Excel لعرض الطلبات")

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666; padding: 1rem;">
    <p>نظام بلسم لإدارة الصيدليات | جميع الحقوق محفوظة © 2026</p>
</div>
""", unsafe_allow_html=True)
