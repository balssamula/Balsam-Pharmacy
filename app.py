import streamlit as st
from utils.database import init_database, fetch_user, update_last_access, get_user_permissions
from utils.helpers import get_branch_number

# تهيئة قاعدة البيانات
init_database()

st.set_page_config(
    page_title="نظام بلسم العلا - مطابقة الطلبات والفواتير",
    layout="wide",
    initial_sidebar_state="expanded",
)

# CSS المشترك
st.markdown(
    """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@300;400;500;700;800&display=swap');
    * { font-family: 'Tajawal', sans-serif; }
    .hero {
        background: linear-gradient(135deg, #0f4c5c 0%, #1f7a8c 50%, #16425b 100%);
        border-radius: 24px;
        padding: 2.2rem;
        color: white;
        margin-bottom: 1.6rem;
        box-shadow: 0 18px 40px rgba(22, 66, 91, 0.20);
    }
    .hero h1 { margin: 0; font-size: 2.2rem; font-weight: 800; }
    .hero p { margin-top: 0.6rem; font-size: 1rem; opacity: 0.95; }
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
    .metric-box {
        background: white;
        border-radius: 18px;
        padding: 1rem;
        border: 1px solid #e6eef0;
        box-shadow: 0 8px 20px rgba(15, 76, 92, 0.06);
        text-align: center;
    }
    .stButton button { width: 100%; border-radius: 10px; font-weight: 800; }
    .session-card {
        background: #f8f9fa;
        border-radius: 12px;
        padding: 0.8rem;
        margin: 0.3rem 0;
        border-right: 3px solid #1f7a8c;
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
    .pill-cancel { background: #ffd8d8; color: #8f1f1f; }
    .pill-payment { background: #fff0c2; color: #8a5b00; }
    .pill-completed { background: #28a745; color: white; }
    .lock-badge { display: inline-block; padding: 0.2rem 0.6rem; border-radius: 20px; font-size: 0.7rem; font-weight: bold; }
    .lock-closed { background: #d9534f; color: white; }
    .lock-open { background: #5cb85c; color: white; }
    .diff-positive { color: #0f7a3a; font-weight: 800; }
    .diff-negative { color: #a32929; font-weight: 800; }
</style>
""",
    unsafe_allow_html=True,
)

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
        
        # قائمة الأدوات للمدير حسب الصلاحيات
        if st.session_state.user_role == "admin":
            st.markdown("---")
            st.markdown("### 📂 أدوات المدير")
            
            permissions = get_user_permissions(st.session_state.username)
            
            if permissions and permissions.get("can_view_dashboard"):
                if st.button("📊 لوحة التحكم", use_container_width=True):
                    st.session_state.page = "dashboard"
                    st.rerun()
            
            if permissions and permissions.get("can_view_balances"):
                if st.button("🔄 تحديث الأرصدة", use_container_width=True):
                    st.session_state.page = "balances"
                    st.rerun()
            
            if permissions and permissions.get("can_view_monitoring"):
                if st.button("👥 مراقبة التعديلات", use_container_width=True):
                    st.session_state.page = "monitoring"
                    st.rerun()
            
            if permissions and permissions.get("can_manage_users"):
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
    st.markdown(
        """
        <div class="hero">
            <h1>نظام بلسم العلا لمراقبة إدخالات الفواتير</h1>
            <p>يعرض الإضافات والإرجاعات الفعلية، ويفصل السطور غير المربوطة وحالات اختلاف الفرع، ويحافظ على حالة كل فرع بين عمليات الرفع.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
elif st.session_state.user_role == "pharmacy":
    if not st.session_state.pharmacist_name:
        st.info("👈 الرجاء إدخال اسم الصيدلي من القائمة الجانبية")
    else:
        from pages import pharmacy_dashboard
        pharmacy_dashboard.show()
else:  # admin
    permissions = get_user_permissions(st.session_state.username)
    page = st.session_state.get("page", "dashboard")
    
    if page == "balances" and permissions and permissions.get("can_view_balances"):
        from pages import balances_updater
        balances_updater.show()
    elif page == "monitoring" and permissions and permissions.get("can_view_monitoring"):
        from pages import monitoring
        monitoring.show()
    elif page == "users" and permissions and permissions.get("can_manage_users"):
        from pages import users_management
        users_management.show()
    else:
        if permissions and permissions.get("can_view_dashboard"):
            from pages import admin_dashboard
            admin_dashboard.show()
        else:
            st.error("⚠️ ليس لديك صلاحية الوصول إلى لوحة التحكم")

st.markdown("---")
st.markdown(
    """
    <div style="text-align:center;color:#607783;padding:0.6rem 0 0.8rem;">
        نظام بلسم العلا لمطابقة الطلبات والفواتير © 2026
    </div>
    """,
    unsafe_allow_html=True,
)