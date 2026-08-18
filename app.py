import streamlit as st
from PIL import Image
from utils.database import init_database, fetch_user, update_last_access, get_user_permissions
from utils.helpers import get_branch_number
from streamlit_javascript import st_javascript

# تهيئة قاعدة البيانات
init_database()

app_icon = Image.open("لوجو--جديد.png")

# إعدادات الصفحة
st.set_page_config(
    page_title="نظام بلسم العلا - مطابقة الطلبات والفواتير",
    page_icon=app_icon,  # 💡 تمرير اللوجو كأيقونة للتطبيق هنا
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
    .session-card {
        background: #f8f9fa;
        border-radius: 12px;
        padding: 0.8rem;
        margin: 0.3rem 0;
        border-right: 3px solid #1f7a8c;
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
if 'page' not in st.session_state:
    st.session_state.page = "dashboard"

# Sidebar Login
with st.sidebar:
    st.title("🌟 نظام بلسم العلا")
    st.caption("مطابقة طلبات سلة والفواتير")
    st.caption("Balsam Alula Pharmacy")
    st.markdown("---")

    # 1. إذا لم يكن مسجلاً للدخول، نعرض حقول الإدخال وزر الدخول فقط
    if not st.session_state.logged_in:
        username = st.text_input("👤 اسم المستخدم")
        password = st.text_input("🔒 كلمة المرور", type="password")
        
        # 💡 تم إصلاح المسافة البادئة ليكون الزر داخل شرط عدم تسجيل الدخول
        if st.button("🚪 دخول", use_container_width=True):
            user = fetch_user(username, password)
            if user:
                st.session_state.logged_in = True
                st.session_state.username = user[0]
                st.session_state.user_role = user[1]
                
                # تفريغ اسم الصيدلي إجبارياً للصيدليات ليطلب إدخال الشيفت من جديد
                if user[1] == "pharmacy":
                    st.session_state.pharmacist_name = ""
                    st.session_state.login_recorded = False
                else:
                    st.session_state.pharmacist_name = user[2] or ""
                    # تسجيل دخول الإدارة في سجل العمليات
                    from utils.database import log_action
                    log_action(user[0], user[1], "النظام", "عام", "عام", "تسجيل دخول", f"تم تسجيل دخول {user[1]} للنظام")
                    
                st.rerun()
            else:
                st.error("❌ بيانات الدخول غير صحيحة.")
                
    # 2. في حالة كان المستخدم مسجلاً للدخول بالفعل (هذا الـ else يتبع الـ if الأولى)
    else:
        st.success(f"مرحباً {st.session_state.username}")

        # التقاط "آخر نشاط" بصمت مع كل ضغطة زر أو تفاعل من المستخدم
        from utils.database import update_last_seen
        update_last_seen(st.session_state.username)
        
        # طلب اسم الصيدلي والفترة للصيادلة
        if st.session_state.user_role == "pharmacy":
            
            # جلب الـ IP الفعلي من جهاز العميل بصمت في الخلفية عبر ipify
            real_ip = st_javascript("await fetch('https://api.ipify.org?format=json').then(r => r.json()).then(d => d.ip).catch(() => 'غير معروف')")
            
            # عرض اسم الصيدلي الحالي وزر لتغييره
            if st.session_state.get('pharmacist_name'):
                st.info(f"الصيدلي الحالي: {st.session_state.pharmacist_name}")
                if st.button("🔄 تسجيل صيدلي مختلف"):
                    st.session_state.pharmacist_name = ""
                    st.rerun()
            
            # إذا لم يكن هناك اسم مسجل، نظهر نموذج الإدخال
            if not st.session_state.get('pharmacist_name'):
                with st.form("pharmacist_login_form"):
                    st.markdown("### 👤 تسجيل بيانات الشيفت")
                    name = st.text_input("اسم الصيدلي (ثلاثي)")
                    shift = st.radio("الفترة (الشيفت)", ["صباحاً ☀️", "مساءً 🌙"])
                    
                    # ننتظر حتى يتم جلب الـ IP 
                    final_ip = real_ip if real_ip and real_ip != 0 else "جاري الجلب..."
                    st.caption(f"🌐 IP الجهاز الحالي: {final_ip}")
                    
                    submitted = st.form_submit_button("💾 بدء العمل")
                    
                    if submitted:
                        if name.strip():
                            full_name = f"{name.strip()} - {shift}"
                            st.session_state.pharmacist_name = full_name
                            
                            # التأكد من التقاط الـ IP الصحيح قبل الحفظ
                            ip_to_save = real_ip if real_ip and real_ip != 0 else "غير معروف"
                            
                            from utils.database import update_last_access, log_action
                            update_last_access(st.session_state.username, full_name, ip_to_save)
                            
                            # تسجيل دخول الصيدلية وبدء الشيفت في سجل العمليات
                            log_action(full_name, "pharmacy", st.session_state.username, "عام", "عام", "بدء شيفت", f"تم تسجيل الدخول وبدء الشيفت من IP: {ip_to_save}")
                            
                            st.session_state.login_recorded = True
                            st.session_state.current_ip = ip_to_save
                            st.rerun()
                        else:
                            st.error("❌ يرجى إدخال اسم الصيدلي")
            
            # عرض الـ IP فقط إذا تم التسجيل
            if st.session_state.get('pharmacist_name') and st.session_state.get('current_ip'):
                st.success(f"✅ تم تسجيل الدخول من IP: {st.session_state.current_ip}")
        
        # قائمة الأدوات حسب الصلاحيات
        if st.session_state.user_role in ["admin", "manager"]:
            st.markdown("---")
            st.markdown("### 📂 أدوات المطابقة والتعديلات")
            
            permissions = get_user_permissions(st.session_state.username)
            
            if permissions and permissions.get("can_view_dashboard"):
                if st.button("📊 لوحة تحكم المطابقات ورفع الملفات", use_container_width=True):
                    st.session_state.page = "dashboard"
                    st.rerun()
                      
            if permissions and permissions.get("can_view_monitoring"):
                if st.button("👥 مراقبة التعديلات", use_container_width=True):
                    st.session_state.page = "monitoring"
                    st.rerun()
            
            if permissions and permissions.get("can_manage_users"):
                if st.button("👥 إدارة المستخدمين", use_container_width=True):
                    st.session_state.page = "users"
                    st.rerun()

            st.markdown("---")
            st.markdown("### 📦 تقارير إضافية")
            if st.button("📦 تفصيلي المنتجات من سلة", use_container_width=True):
                st.session_state.page = "product_details"
                st.rerun()

            if st.button("🔄 تحديث الأرصدة", use_container_width=True):
                st.session_state.page = "balances"
                st.rerun()

            if st.button("📈 تحليل الصيدليات الشامل", use_container_width=True):
                st.session_state.page = "comprehensive_analysis"
                st.rerun()
                
            if st.button("📊 تحليل مبيعات الشهور", use_container_width=True):
                st.session_state.page = "sales_analysis"
                st.rerun()

        st.markdown("---")
        st.markdown("### 🛍️ العروض")

        if st.button("🛍️ العروض الحالية الفعالة بالمتجر", use_container_width=True):
            st.session_state.page = "promotion_viewer"
            st.rerun()

        if st.button("🏷️ حاسبة العروض الترويجية", use_container_width=True):
            st.session_state.page = "promotions"
            st.rerun()
                
        if st.button("🚪 تسجيل خروج", use_container_width=True):
            # تسجيل الخروج في السجل الشامل قبل مسح الجلسة
            if st.session_state.get('logged_in'):
                from utils.database import log_action
                p_name = st.session_state.get('pharmacist_name') or st.session_state.get('username')
                role = st.session_state.get('user_role')
                ph_name = st.session_state.get('username') if role == 'pharmacy' else "النظام"
                
                log_action(p_name, role, ph_name, "عام", "عام", "تسجيل خروج", "تم تسجيل الخروج من النظام")

            # تفريغ المتغيرات
            st.session_state.logged_in = False
            st.session_state.username = ""
            st.session_state.user_role = ""
            st.session_state.pharmacist_name = ""
            st.session_state.page = "dashboard"
            st.session_state.login_recorded = False
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
            <div>🏥 فرع</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div class="metric-box">
            <div style="font-size:1.5rem;font-weight:800;">1000+</div>
            <div>📦 طلب شهرياً</div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown("""
        <div class="metric-box">
            <div style="font-size:1.5rem;font-weight:800;">99%</div>
            <div>⚡ دقة المطابقة</div>
        </div>
        """, unsafe_allow_html=True)
        
elif st.session_state.user_role == "pharmacy":
    if not st.session_state.pharmacist_name:
        st.info("👈 الرجاء إدخال اسم الصيدلي من القائمة الجانبية")
    else:
        from pages import pharmacy_dashboard
        pharmacy_dashboard.show()
else:  # admin or manager
    page = st.session_state.get("page", "dashboard")
    permissions = get_user_permissions(st.session_state.username)
    
    if page == "users" and permissions and permissions.get("can_manage_users"):
        from pages import users_management
        users_management.show()
    elif page == "monitoring" and permissions and permissions.get("can_view_monitoring"):
        from pages import monitoring
        monitoring.show()
    # ========== صفحة تفصيلي المنتجات - تظهر لـ admin و manager فقط ==========
    elif page == "product_details" and st.session_state.user_role in ["admin", "manager"]:
        from pages import product_details
        product_details.show()

    elif page == "balances" and permissions and st.session_state.user_role in ["admin", "manager"]:
        from pages import balances_updater
        balances_updater.show()

    elif page == "comprehensive_analysis" and st.session_state.user_role in ["admin", "manager"]:
        from pages import comprehensive_analysis
        comprehensive_analysis.show()

    # 💡 توجيه صفحة تحليل مبيعات الشهور
    elif page == "sales_analysis" and st.session_state.user_role in ["admin", "manager"]:
        from pages import sales_analysis
        sales_analysis.show()

    # ✨ [تعديل 2]: توجيه واستدعاء صفحة حاسبة العروض الجديدة عند اختيارها
    elif page == "promotions" and st.session_state.user_role in ["admin", "manager"]:
        from pages import promotion_calculator
        promotion_calculator.show()
        
    # 🆕 توجيه صفحة العروض الحالية
    elif page == "promotion_viewer" and st.session_state.user_role in ["admin", "manager"]:
        from pages import promotion_viewer
        promotion_viewer.show()
        
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
        نظام بلسم العلا لمطابقة الطلبات والفواتير (V1) © 2026
    </div>
    """,
    unsafe_allow_html=True,
)
