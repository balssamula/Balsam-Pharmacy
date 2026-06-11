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

st.markdown(
    """
    <style>
    /* إخفاء قسم روابط التنقل التلقائي في السايدبار */
    [data-testid="stSidebarNav"] {
        display: none !important;
    }
    </style>
    """,
    unsafe_allow_html=True
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
        text-align: right;
    }
</style>
""", unsafe_allow_html=True)

# تتبع حالة تسجيل الدخول وحفظها في الجلسة (Session State)
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    # ------------------ نموذج تسجيل الدخول ------------------
    st.title("🔐 تسجيل الدخول إلى النظام")
    username = st.text_input("اسم المستخدم")
    password = st.text_input("كلمة المرور", type="password")
    
    if st.button("تسجيل الدخول", use_container_width=True):
        user = fetch_user(username, password)
        if user:
            st.session_state.logged_in = True
            st.session_state.username = username
            st.session_state.user_role = user.get("role", "pharmacy")  # حفظ الدور: admin, manager, pharmacy
            st.session_state.pharmacist_name = user.get("pharmacist_name", "")
            update_last_access(username)
            st.rerun()
        else:
            st.error("❌ اسم المستخدم أو كلمة المرور غير صحيحة")
else:
    # ------------------ بعد تسجيل الدخول بنجاح ------------------
    permissions = get_user_permissions(st.session_state.username)
    
    # بناء القائمة الجانبية المخصصة بناءً على الصلاحيات والأدوار
    st.sidebar.title(f"👋 أهلاً، {st.session_state.username}")
    
    # إعداد الخيارات المتاحة في القائمة الجانبية
    nav_options = {}
    
    if permissions and permissions.get("can_view_dashboard"):
        if st.session_state.user_role in ["admin", "manager"]:
            nav_options["لوحة التحكم الإدارية"] = "dashboard"
        else:
            nav_options["لوحة تحكم الصيدلية"] = "dashboard"
            
    if permissions and permissions.get("can_view_balances"):
        nav_options["💰 تحديث الأرصدة والمالية"] = "balances"
        
    if permissions and permissions.get("can_view_monitoring"):
        nav_options["📈 مراقبة العمليات والتحليلات"] = "monitoring"
        
    if st.session_state.user_role in ["admin", "manager"]:
        nav_options["📑 تفصيلي المنتجات"] = "product_details"
        nav_options["📊 تحليل مبيعات الشهور"] = "sales_analysis"
        # ✨ تم الإضافة هنا: خيار العروض الفعالة الجديد بالاسم المطلوب لـ admin و manager
        nav_options["🛍️ العروض الحالية الفعالة بالمتجر"] = "promotions"
        
    if permissions and permissions.get("can_manage_users"):
        nav_options["👥 إدارة المستخدمين والصلاحيات"] = "users"
        
    # عرض القائمة الجانبية للمسؤولين والصيادلة لاختيار الصفحة
    selected_page_label = st.sidebar.radio("🗂️ الانتقال السريع:", list(nav_options.keys()))
    page = nav_options[selected_page_label]
    
    # تسجيل الخروج
    if st.sidebar.button("🚪 تسجيل الخروج", use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.username = None
        st.session_state.user_role = None
        st.rerun()
