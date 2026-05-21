import streamlit as st
from utils.database import init_database, fetch_user, update_last_access, get_user_permissions
from utils.helpers import get_branch_number

# تهيئة قاعدة البيانات
init_database()

# تحديث صلاحيات المدير
def fix_admin_permissions():
    import sqlite3
    import os
    db_path = 'data/pharmacy_reconciliation.db'
    if os.path.exists(db_path):
        conn = sqlite3.connect(db_path)
        cur = conn.cursor()
        try:
            # تحديث صلاحيات admin
            cur.execute("""
                UPDATE users 
                SET can_view_dashboard = 1, 
                    can_view_balances = 1, 
                    can_view_monitoring = 1, 
                    can_manage_users = 1 
                WHERE username = 'admin'
            """)
            conn.commit()
            
            # إضافة الأعمدة إذا لم تكن موجودة
            cur.execute("PRAGMA table_info(users)")
            existing_columns = [row[1] for row in cur.fetchall()]
            
            if "can_view_balances" not in existing_columns:
                cur.execute("ALTER TABLE users ADD COLUMN can_view_balances INTEGER DEFAULT 0")
            if "can_view_monitoring" not in existing_columns:
                cur.execute("ALTER TABLE users ADD COLUMN can_view_monitoring INTEGER DEFAULT 0")
            if "can_manage_users" not in existing_columns:
                cur.execute("ALTER TABLE users ADD COLUMN can_manage_users INTEGER DEFAULT 0")
            
            conn.commit()
        except Exception as e:
            print(f"خطأ: {e}")
        finally:
            conn.close()

fix_admin_permissions()

st.set_page_config(
    page_title="نظام بلسم العلا - مطابقة الطلبات والفواتير",
    layout="wide",
)

# CSS المشترك
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@300;400;500;700;800&display=swap');
    * { font-family: 'Tajawal', sans-serif; }
    .hero {
        background: linear-gradient(135deg, #0f4c5c 0%, #1f7a8c 50%, #16425b 100%);
        border-radius: 24px;
        padding: 2rem;
        color: white;
        margin-bottom: 1rem;
        text-align: center;
    }
    .metric-box {
        background: white;
        border-radius: 18px;
        padding: 1rem;
        border: 1px solid #e6eef0;
        text-align: center;
        margin: 0.5rem;
    }
    .stButton button { width: 100%; border-radius: 10px; }
    .note-card {
        background: linear-gradient(135deg, #f4fbfc 0%, #ffffff 100%);
        border: 1px solid #d7ebef;
        border-radius: 16px;
        padding: 1rem;
        margin-bottom: 1rem;
    }
    .section-title {
        font-size: 1.15rem;
        font-weight: 800;
        color: #16425b;
        border-right: 5px solid #1f7a8c;
        padding-right: 0.65rem;
        margin: 1rem 0 0.8rem;
    }
</style>
""", unsafe_allow_html=True)

# Session State
for key, default_value in {
    "logged_in": False,
    "username": "",
    "user_role": "",
    "pharmacist_name": "",
    "page": "dashboard",
}.items():
    if key not in st.session_state:
        st.session_state[key] = default_value

# Sidebar Login
with st.sidebar:
    st.title("نظام بلسم العلا")
    st.caption("مطابقة طلبات سلة والفواتير")
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
        st.success(f"مرحباً {st.session_state.username}")
        
        # طلب اسم الصيدلي للصيادلة
        if st.session_state.user_role == "pharmacy" and not st.session_state.pharmacist_name:
            pharmacist_input = st.text_input("👤 اسم الصيدلي", key="pharmacist_name_input")
            if st.button("💾 حفظ الاسم", use_container_width=True):
                if pharmacist_input.strip():
                    st.session_state.pharmacist_name = pharmacist_input.strip()
                    update_last_access(st.session_state.username, st.session_state.pharmacist_name)
                    st.success("✅ تم حفظ الاسم")
                    st.rerun()
        
        # قائمة الأدوات للمدير
        if st.session_state.user_role == "admin":
            st.markdown("---")
            st.markdown("### 📂 أدوات المدير")
            
            # زر لوحة التحكم الرئيسية
            if st.button("📊 لوحة التحكم الرئيسية", use_container_width=True):
                st.session_state.page = "dashboard"
                st.rerun()
            
            # زر تحديث الأرصدة
            if st.button("🔄 تحديث الأرصدة", use_container_width=True):
                st.session_state.page = "balances"
                st.rerun()
            
            # زر مراقبة تعديلات الصيدليات
            if st.button("👥 مراقبة تعديلات الصيدليات", use_container_width=True):
                st.session_state.page = "monitoring"
                st.rerun()
            
            # زر إدارة المستخدمين
            if st.button("👥 إدارة المستخدمين والصلاحيات", use_container_width=True):
                st.session_state.page = "users"
                st.rerun()
        
        st.markdown("---")
        if st.button("🚪 تسجيل خروج", use_container_width=True):
            for key in ["logged_in", "username", "user_role", "pharmacist_name", "page"]:
                st.session_state[key] = False if key == "logged_in" else "dashboard"
            st.rerun()

# Main Content
if not st.session_state.logged_in:
    st.markdown("""
    <div class="hero">
        <h1>نظام بلسم العلا لمراقبة إدخالات الفواتير</h1>
        <p>نظام متكامل لمطابقة طلبات سلة والفواتير</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
        <div class="metric-box">
            <div style="font-size:1.5rem;font-weight:800;">17</div>
            <div>فرع</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div class="metric-box">
            <div style="font-size:1.5rem;font-weight:800;">1000+</div>
            <div>طلب شهرياً</div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown("""
        <div class="metric-box">
            <div style="font-size:1.5rem;font-weight:800;">99%</div>
            <div>دقة المطابقة</div>
        </div>
        """, unsafe_allow_html=True)
        
elif st.session_state.user_role == "pharmacy":
    if not st.session_state.pharmacist_name:
        st.info("👈 الرجاء إدخال اسم الصيدلي من القائمة الجانبية")
    else:
        try:
            from pages import pharmacy_dashboard
            pharmacy_dashboard.show()
        except ImportError:
            st.error("❌ صفحة الصيدلية غير متوفرة")
else:  # admin
    page = st.session_state.get("page", "dashboard")
    
    if page == "users":
        try:
            from pages import users_management
            users_management.show()
        except ImportError:
            st.error("❌ صفحة إدارة المستخدمين غير متوفرة")
    elif page == "balances":
        try:
            from pages import balances_updater
            balances_updater.show()
        except ImportError:
            st.error("❌ صفحة تحديث الأرصدة غير متوفرة")
    elif page == "monitoring":
        try:
            from pages import monitoring
            monitoring.show()
        except ImportError:
            st.error("❌ صفحة مراقبة التعديلات غير متوفرة")
    else:
        try:
            from pages import admin_dashboard
            admin_dashboard.show()
        except ImportError:
            st.error("❌ صفحة لوحة التحكم غير متوفرة")

st.markdown("---")
st.markdown(
    """
    <div style="text-align:center;color:#607783;padding:0.6rem 0 0.8rem;">
        نظام بلسم العلا لمطابقة الطلبات والفواتير © 2026
    </div>
    """,
    unsafe_allow_html=True,
)
