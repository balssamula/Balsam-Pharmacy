import streamlit as st
from utils.database import init_database, fetch_user, update_last_access
from utils.helpers import get_branch_number

# تهيئة قاعدة البيانات
init_database()

st.set_page_config(
    page_title="نظام بلسم العلا",
    layout="wide",
)

# CSS بسيط
st.markdown("""
<style>
    * { font-family: 'Tajawal', sans-serif; }
    .hero {
        background: linear-gradient(135deg, #0f4c5c 0%, #1f7a8c 100%);
        border-radius: 24px;
        padding: 2rem;
        color: white;
        text-align: center;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# Session State
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
if 'username' not in st.session_state:
    st.session_state.username = ""
if 'user_role' not in st.session_state:
    st.session_state.user_role = ""
if 'pharmacist_name' not in st.session_state:
    st.session_state.pharmacist_name = ""

# Sidebar Login
with st.sidebar:
    st.title("نظام بلسم العلا")
    
    if not st.session_state.logged_in:
        username = st.text_input("اسم المستخدم")
        password = st.text_input("كلمة المرور", type="password")
        if st.button("دخول"):
            user = fetch_user(username, password)
            if user:
                st.session_state.logged_in = True
                st.session_state.username = user[0]
                st.session_state.user_role = user[1]
                st.session_state.pharmacist_name = user[2] or ""
                st.rerun()
            else:
                st.error("بيانات الدخول غير صحيحة")
    else:
        st.success(f"مرحباً {st.session_state.username}")
        if st.button("تسجيل خروج"):
            st.session_state.logged_in = False
            st.session_state.username = ""
            st.session_state.user_role = ""
            st.session_state.pharmacist_name = ""
            st.rerun()

# Main Content
if not st.session_state.logged_in:
    st.markdown("""
    <div class="hero">
        <h1>نظام بلسم العلا</h1>
        <p>نظام متكامل لمطابقة طلبات سلة والفواتير</p>
    </div>
    """, unsafe_allow_html=True)
else:
    st.markdown("""
    <div class="hero">
        <h1>مرحباً بك في النظام</h1>
        <p>تم تسجيل الدخول بنجاح</p>
    </div>
    """, unsafe_allow_html=True)
    
    if st.session_state.user_role == "admin":
        st.info("👑 أنت مسجل كمدير عام")
        from pages.admin_dashboard import show
        show()
    else:
        st.info(f"🏥 مرحباً في {st.session_state.username}")
        st.info("سيتم إضافة صفحات الصيدليات قريباً")
