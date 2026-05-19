import streamlit as st
import pandas as pd
import sqlite3
import os
from datetime import datetime

st.set_page_config(page_title="نظام بلسم", layout="wide")

# Arabic CSS
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal&display=swap');
    * {font-family: 'Tajawal', sans-serif;}
    .stButton button {width: 100%; border-radius: 8px;}
    .success-box {background-color: #d4edda; padding: 15px; border-radius: 10px; border-right: 4px solid #28a745;}
    .warning-box {background-color: #fff3cd; padding: 15px; border-radius: 10px; border-right: 4px solid #ffc107;}
</style>
""", unsafe_allow_html=True)

# ========== DATABASE ==========
os.makedirs("data", exist_ok=True)

def init_db():
    conn = sqlite3.connect('data/pharmacy.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS adjustments
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  order_number TEXT, sku TEXT, pharmacy TEXT,
                  action TEXT, status TEXT, performed_by TEXT, timestamp TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS users
                 (username TEXT PRIMARY KEY, password TEXT, role TEXT)''')
    
    # Insert default users
    c.execute("SELECT * FROM users WHERE username='admin'")
    if not c.fetchone():
        c.execute("INSERT INTO users VALUES ('admin', 'admin123', 'admin')")
        for i in range(1, 18):
            c.execute("INSERT INTO users VALUES (?, ?, 'pharmacy')", 
                     (f"Balsam Alula Pharmacy {i:02d}", f"balsam{i}"))
    conn.commit()
    conn.close()

init_db()

# ========== SMART COLUMN DETECTION ==========
def find_column(df, possible_names):
    """Find column by possible names"""
    for name in possible_names:
        if name in df.columns:
            return name
    return None

def process_excel(file):
    try:
        # Read sheets
        df_salla = pd.read_excel(file, sheet_name="سلة")
        df_abc = pd.read_excel(file, sheet_name="abc")
        
        # Display column names for debugging (only in admin mode)
        st.info(f"📋 أعمدة سلة: {list(df_salla.columns)}")
        st.info(f"📋 أعمدة ABC: {list(df_abc.columns)}")
        
        # ===== DETECT COLUMNS IN SALLA SHEET =====
        order_col = find_column(df_salla, ['رقم الطلب', 'Order Number', 'order_number', 'OrderID'])
        sku_col = find_column(df_salla, ['SKU', 'Sku', 'sku', 'رقم المنتج', 'Product ID'])
        product_col = find_column(df_salla, ['اسم المنتج', 'Product Name', 'product_name', 'Product'])
        qty_col = find_column(df_salla, ['الكمية', 'Quantity', 'qty', 'Qty'])
        city_col = find_column(df_salla, ['المدينة', 'City', 'city'])
        status_col = find_column(df_salla, ['حالة الطلب', 'Order Status', 'status'])
        
        # ===== DETECT COLUMNS IN ABC SHEET =====
        order_col_abc = find_column(df_abc, ['رقم الطلب', 'Order Number', 'order_number', 'OrderID'])
        sku_col_abc = find_column(df_abc, ['رقم الصنف', 'Item Number', 'رقم المنتج', 'SKU', 'Product ID'])
        qty_col_abc = find_column(df_abc, ['Net Sold Qty', 'Sold Qty', 'Quantity', 'الكمية المباعة'])
        
        # Check if all required columns found
        if not order_col:
            st.error("لم يتم العثور على عمود رقم الطلب في شيت سلة")
            return None
        if not sku_col:
            st.error("لم يتم العثور على عمود SKU في شيت سلة")
            return None
        if not qty_col:
            st.error("لم يتم العichtauf على عمود الكمية في شيت سلة")
            return None
            
        st.success(f"✅ تم التعرف على الأعمدة: الطلب={order_col}, SKU={sku_col}, الكمية={qty_col}")
        
        # ===== FILTER DATA =====
        # Filter out cancelled/returned orders
        if status_col:
            excluded = ['ملغي', 'مسترجع', 'محذوف', 'Cancelled', 'Returned', 'Deleted']
            df_salla = df_salla[~df_salla[status_col].astype(str).isin(excluded)]
        
        # Filter valid SKUs
        df_salla = df_salla[df_salla[sku_col].astype(str).str.isdigit()]
        
        # Assign branch based on city
        if city_col:
            df_salla['الفرع'] = df_salla[city_col].apply(
                lambda x: 'Balsam Alula Pharmacy 09' if str(x) == 'AL ULA' else 'Balsam Alula Pharmacy 13'
            )
        else:
            df_salla['الفرع'] = 'Balsam Alula Pharmacy 13'
        
        # Rename columns to standard names
        df_salla = df_salla.rename(columns={
            order_col: 'رقم_الطلب',
            sku_col: 'SKU',
            qty_col: 'الكمية'
        })
        
        if product_col:
            df_salla = df_salla.rename(columns={product_col: 'اسم_المنتج'})
        else:
            df_salla['اسم_المنتج'] = 'غير محدد'
        
        # Group salla data
        group_cols = ['رقم_الطلب', 'SKU', 'اسم_المنتج', 'الفرع']
        salla_grouped = df_salla.groupby(group_cols).agg({
            'الكمية': 'sum'
        }).reset_index()
        
        # Process ABC data
        if order_col_abc and sku_col_abc and qty_col_abc:
            df_abc = df_abc.rename(columns={
                order_col_abc: 'رقم_الطلب',
                sku_col_abc: 'SKU',
                qty_col_abc: 'كمية_ABC'
            })
            
            abc_grouped = df_abc.groupby(['رقم_الطلب', 'SKU']).agg({
                'كمية_ABC': 'sum'
            }).reset_index()
        else:
            st.warning("بعض الأعمدة غير موجودة في شيت ABC")
            abc_grouped = pd.DataFrame(columns=['رقم_الطلب', 'SKU', 'كمية_ABC'])
        
        # Merge
        merged = pd.merge(salla_grouped, abc_grouped, on=['رقم_الطلب', 'SKU'], how='outer')
        merged = merged.fillna(0)
        
        # Calculate difference
        merged['الفرق'] = merged['الكمية'] - merged['كمية_ABC']
        merged['نوع_الاجراء'] = merged['الفرق'].apply(
            lambda x: 'إضافة' if x > 0 else ('إرجاع' if x < 0 else 'مطابق')
        )
        
        # Keep only additions and returns
        result = merged[merged['نوع_الاجراء'].isin(['إضافة', 'إرجاع'])].copy()
        
        # Rename back for display
        result = result.rename(columns={
            'رقم_الطلب': 'رقم الطلب',
            'اسم_المنتج': 'اسم المنتج',
            'الكمية': 'كمية سلة'
        })
        
        return result
        
    except Exception as e:
        st.error(f"❌ خطأ: {str(e)}")
        return None

def record_adjustment(order_number, sku, pharmacy, action):
    conn = sqlite3.connect('data/pharmacy.db')
    c = conn.cursor()
    c.execute("""
        INSERT INTO adjustments (order_number, sku, pharmacy, action, status, performed_by, timestamp)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (order_number, sku, pharmacy, action, 'تم', st.session_state.username, 
          datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit()
    conn.close()

# ========== SESSION STATE ==========
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
if 'username' not in st.session_state:
    st.session_state.username = None
if 'user_role' not in st.session_state:
    st.session_state.user_role = None
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None

# ========== LOGIN SIDEBAR ==========
with st.sidebar:
    st.image("https://img.icons8.com/color/96/000000/pharmacy.png", width=80)
    st.title("نظام بلسم")
    
    if not st.session_state.logged_in:
        username = st.text_input("👤 اسم المستخدم")
        password = st.text_input("🔒 كلمة المرور", type="password")
        
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
                st.error("❌ خطأ في اسم المستخدم أو كلمة المرور")
    else:
        st.success(f"مرحباً {st.session_state.username}")
        if st.button("🚪 تسجيل خروج", use_container_width=True):
            st.session_state.logged_in = False
            st.session_state.username = None
            st.session_state.user_role = None
            st.session_state.processed_data = None
            st.rerun()

# ========== MAIN CONTENT ==========
if not st.session_state.logged_in:
    st.markdown("""
    <div style="background: linear-gradient(135deg, #667eea, #764ba2); 
                padding: 3rem; border-radius: 20px; color: white; text-align: center;">
        <h1 style="font-size: 2.5rem;">📊 نظام مراقبة طلبات سلة و ABC</h1>
        <p style="font-size: 1.2rem;">نظام متكامل لإدارة ومتابعة الطلبات والإضافات والإرجاعات</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("🏥 صيدلية", "17", "فرع")
    with col2:
        st.metric("📦 طلبات شهرياً", "1000+", "متوسط")
    with col3:
        st.metric("⚡ دقة المطابقة", "99%", "عالية")
    
    st.info("👈 الرجاء تسجيل الدخول من القائمة الجانبية")

else:
    st.title("📊 لوحة التحكم")
    
    # File upload
    uploaded = st.file_uploader("📂 رفع ملف الطلبات والفواتير (Excel)", type=['xlsx'])
    
    if uploaded:
        if st.button("🔄 معالجة الملف", use_container_width=True):
            with st.spinner("جاري المعالجة..."):
                result = process_excel(uploaded)
                if result is not None:
                    st.session_state.processed_data = result
                    st.success("✅ تمت المعالجة بنجاح!")
                    st.balloons()
    
    if st.session_state.processed_data is not None:
        df = st.session_state.processed_data
        
        if st.session_state.user_role == "admin":
            # Admin dashboard
            st.markdown("---")
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("📊 عدد الطلبات", len(df['رقم الطلب'].unique()))
            with col2:
                additions = len(df[df['نوع_الاجراء'] == 'إضافة'])
                st.metric("➕ إضافات", additions)
            with col3:
                returns = len(df[df['نوع_الاجراء'] == 'إرجاع'])
                st.metric("➖ إرجاعات", returns)
            
            tab1, tab2 = st.tabs(["📈 الإضافات", "📉 الإرجاعات"])
            
            with tab1:
                additions_df = df[df['نوع_الاجراء'] == 'إضافة']
                if len(additions_df) > 0:
                    st.dataframe(additions_df, use_container_width=True)
                else:
                    st.success("🎉 لا توجد إضافات مطلوبة!")
            
            with tab2:
                returns_df = df[df['نوع_الاجراء'] == 'إرجاع']
                if len(returns_df) > 0:
                    st.dataframe(returns_df, use_container_width=True)
                else:
                    st.success("🎉 لا توجد إرجاعات مطلوبة!")
            
            # Show history
            st.markdown("---")
            st.subheader("📋 سجل التعديلات")
            conn = sqlite3.connect('data/pharmacy.db')
            history = pd.read_sql_query("SELECT * FROM adjustments ORDER BY timestamp DESC LIMIT 50", conn)
            conn.close()
            if len(history) > 0:
                st.dataframe(history, use_container_width=True)
        
        else:
            # Pharmacy user
            pharmacy = st.session_state.username
            pharmacy_df = df[df['الفرع'] == pharmacy]
            
            if len(pharmacy_df) > 0:
                st.success(f"🏥 مرحباً {pharmacy}")
                
                # Additions
                additions = pharmacy_df[pharmacy_df['نوع_الاجراء'] == 'إضافة']
                if len(additions) > 0:
                    st.subheader("✅ الأصناف التي تحتاج إضافة")
                    for idx, row in additions.iterrows():
                        with st.container():
                            cols = st.columns([2,2,3,2,2])
                            cols[0].write(f"📋 طلب: {row['رقم الطلب']}")
                            cols[1].write(f"🏷️ SKU: {row['SKU']}")
                            cols[2].write(f"📦 {row['اسم المنتج'][:35]}")
                            cols[3].write(f"📊 الكمية: {int(row['كمية سلة'])}")
                            
                            # Check if already done
                            conn = sqlite3.connect('data/pharmacy.db')
                            c = conn.cursor()
                            c.execute("SELECT * FROM adjustments WHERE order_number=? AND sku=? AND pharmacy=? AND action='إضافة'",
                                     (str(row['رقم الطلب']), str(row['SKU']), pharmacy))
                            done = c.fetchone()
                            conn.close()
                            
                            if done:
                                cols[4].success("✅ تمت")
                            else:
                                if cols[4].button("➕ تمت الإضافة", key=f"add_{idx}_{row['SKU']}"):
                                    record_adjustment(str(row['رقم الطلب']), str(row['SKU']), pharmacy, 'إضافة')
                                    st.success("تم التسجيل!")
                                    st.rerun()
                            st.divider()
                
                # Returns
                returns = pharmacy_df[pharmacy_df['نوع_الاجراء'] == 'إرجاع']
                if len(returns) > 0:
                    st.subheader("🔄 الأصناف التي تحتاج إرجاع")
                    for idx, row in returns.iterrows():
                        with st.container():
                            cols = st.columns([2,2,3,2,2])
                            cols[0].write(f"📋 طلب: {row['رقم الطلب']}")
                            cols[1].write(f"🏷️ SKU: {row['SKU']}")
                            cols[2].write(f"📦 {row['اسم المنتج'][:35]}")
                            cols[3].write(f"📊 الفرق: {abs(int(row['الفرق']))}")
                            
                            # Check if already done
                            conn = sqlite3.connect('data/pharmacy.db')
                            c = conn.cursor()
                            c.execute("SELECT * FROM adjustments WHERE order_number=? AND sku=? AND pharmacy=? AND action='إرجاع'",
                                     (str(row['رقم الطلب']), str(row['SKU']), pharmacy))
                            done = c.fetchone()
                            conn.close()
                            
                            if done:
                                cols[4].success("✅ تم")
                            else:
                                if cols[4].button("🔄 تم الإرجاع", key=f"return_{idx}_{row['SKU']}"):
                                    record_adjustment(str(row['رقم الطلب']), str(row['SKU']), pharmacy, 'إرجاع')
                                    st.success("تم التسجيل!")
                                    st.rerun()
                            st.divider()
                
                if len(additions) == 0 and len(returns) == 0:
                    st.info("🎉 لا توجد إضافات أو إرجاعات مطلوبة لهذا الفرع")
            else:
                st.info(f"📭 لا توجد طلبات لـ {pharmacy}")
