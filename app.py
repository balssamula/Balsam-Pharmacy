import streamlit as st
from utils.database import init_database, fetch_user, update_last_access, get_user_permissions
from utils.helpers import get_branch_number

# تهيئة قاعدة البيانات
init_database()

def fix_admin_permissions():
    import sqlite3
    conn = sqlite3.connect('data/pharmacy_reconciliation.db')
    cur = conn.cursor()
    cur.execute("""
        UPDATE users 
        SET can_view_dashboard = 1, 
            can_view_balances = 1, 
            can_view_monitoring = 1, 
            can_manage_users = 1 
        WHERE username = 'admin'
    """)
    conn.commit()
    conn.close()
    print("✅ تم تحديث صلاحيات المدير")

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
    .pill-completed { background: #28a745; color: white; }
    .stButton button { width: 100%; border-radius: 10px; }
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
            
            if st.button("📊 لوحة التحكم", use_container_width=True):
                st.session_state.page = "dashboard"
                st.rerun()
            
            if st.button("👥 المستخدمين والصلاحيات", use_container_width=True):
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
elif st.session_state.user_role == "pharmacy":
    if not st.session_state.pharmacist_name:
        st.info("👈 الرجاء إدخال اسم الصيدلي من القائمة الجانبية")
    else:
        from pages import pharmacy_dashboard
        pharmacy_dashboard.show()
else:  # admin
    page = st.session_state.get("page", "dashboard")
    
    if page == "users":
        from pages import users_management
        users_management.show()
    else:
        from pages import admin_dashboard
        admin_dashboard.show()

st.markdown("---")
st.markdown(
    """
    <div style="text-align:center;color:#607783;padding:0.6rem 0 0.8rem;">
        نظام بلسم العلا لمطابقة الطلبات والفواتير © 2026
    </div>
    """,
    unsafe_allow_html=True,
)
