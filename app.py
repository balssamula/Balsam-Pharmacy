import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime
import hashlib
import base64
from pathlib import Path

# إعداد الصفحة
st.set_page_config(
    page_title="نظام مراقبة طلبات سلة و ABC",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS مخصص للغة العربية والتصميم الاحترافي
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@300;400;500;700&display=swap');
    
    * {
        font-family: 'Tajawal', sans-serif;
    }
    
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 15px;
        color: white;
        text-align: center;
        margin-bottom: 2rem;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }
    
    .pharmacy-card {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        padding: 1.5rem;
        border-radius: 10px;
        text-align: center;
        cursor: pointer;
        transition: transform 0.3s;
        margin: 10px;
        border: 2px solid transparent;
    }
    
    .pharmacy-card:hover {
        transform: translateY(-5px);
        border-color: #667eea;
    }
    
    .success-alert {
        background-color: #d4edda;
        border-right: 4px solid #28a745;
        padding: 1rem;
        border-radius: 8px;
        margin: 10px 0;
    }
    
    .warning-alert {
        background-color: #fff3cd;
        border-right: 4px solid #ffc107;
        padding: 1rem;
        border-radius: 8px;
        margin: 10px 0;
    }
    
    .danger-alert {
        background-color: #f8d7da;
        border-right: 4px solid #dc3545;
        padding: 1rem;
        border-radius: 8px;
        margin: 10px 0;
    }
    
    .info-alert {
        background-color: #d1ecf1;
        border-right: 4px solid #17a2b8;
        padding: 1rem;
        border-radius: 8px;
        margin: 10px 0;
    }
    
    .button-completed {
        background-color: #28a745;
        color: white;
        border: none;
        padding: 10px 20px;
        border-radius: 8px;
        cursor: pointer;
        font-weight: bold;
    }
    
    .button-pending {
        background-color: #ffc107;
        color: #333;
        border: none;
        padding: 10px 20px;
        border-radius: 8px;
        cursor: pointer;
        font-weight: bold;
    }
    
    .stats-card {
        background: white;
        padding: 1rem;
        border-radius: 10px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        text-align: center;
        margin: 10px;
    }
    
    .stats-number {
        font-size: 2rem;
        font-weight: bold;
        color: #667eea;
    }
    
    /* تنسيق الجداول */
    .dataframe {
        font-size: 14px;
        text-align: right;
    }
    
    .dataframe th {
        background-color: #667eea;
        color: white;
        padding: 10px;
        text-align: center;
    }
    
    .dataframe td {
        padding: 8px;
        text-align: center;
    }
    
    /* زر الصوت */
    .audio-btn {
        position: fixed;
        bottom: 20px;
        right: 20px;
        background-color: #667eea;
        color: white;
        padding: 10px;
        border-radius: 50%;
        cursor: pointer;
        z-index: 1000;
    }
    </style>
""", unsafe_allow_html=True)

# دالة لتشغيل الصوت
def play_sound(sound_type):
    sounds = {
        "success": "🔔",
        "warning": "⚠️",
        "info": "ℹ️",
        "error": "❌"
    }
    sound_emoji = sounds.get(sound_type, "🔔")
    st.markdown(f"""
        <script>
            // تشغيل صوت باستخدام Web Speech API
            var msg = new SpeechSynthesisUtterance();
            msg.text = '{sound_emoji}';
            msg.lang = 'ar-SA';
            window.speechSynthesis.speak(msg);
        </script>
    """, unsafe_allow_html=True)

# تهيئة المجلدات
os.makedirs("data", exist_ok=True)
os.makedirs("uploads", exist_ok=True)

# ملف المستخدمين وكلمات المرور
def init_users():
    users_file = "data/users.json"
    if not os.path.exists(users_file):
        users = {}
        for i in range(1, 18):
            pharmacy_num = f"{i:02d}"
            users[f"Balsam Alula Pharmacy {pharmacy_num}"] = {
                "password": f"balsam{i}",
                "name": f"صيدلية بلسم العلا {pharmacy_num}"
            }
        # إضافة مستخدم مدير
        users["admin"] = {
            "password": "admin123",
            "name": "المدير العام"
        }
        with open(users_file, "w", encoding="utf-8") as f:
            json.dump(users, f, ensure_ascii=False, indent=2)
    return users_file

# ملف التعديلات
def init_adjustments():
    adj_file = "data/adjustments.json"
    if not os.path.exists(adj_file):
        with open(adj_file, "w", encoding="utf-8") as f:
            json.dump([], f, ensure_ascii=False, indent=2)
    return adj_file

# معالجة ملف Excel
def process_excel(uploaded_file):
    if uploaded_file is not None:
        # قراءة الشيتات
        df_salla = pd.read_excel(uploaded_file, sheet_name="سلة")
        df_abc = pd.read_excel(uploaded_file, sheet_name="abc")
        
        # تنظيف بيانات سلة
        df_salla = df_salla[df_salla['حالة الطلب'].notna()]
        excluded_statuses = ['ملغي', 'مسترجع', 'محذوف']
        df_salla = df_salla[~df_salla['حالة الطلب'].isin(excluded_statuses)]
        
        # التحقق من SKU صحيح
        df_salla = df_salla[df_salla['SKU'].astype(str).str.isdigit()]
        
        # توجيه الفروع
        df_salla['الفرع'] = df_salla['المدينة'].apply(
            lambda x: 'Balsam Alula Pharmacy 09' if x == 'AL ULA' else 'Balsam Alula Pharmacy 13'
        )
        
        # تجميع كميات سلة حسب رقم الطلب و SKU
        salla_grouped = df_salla.groupby(['رقم الطلب', 'SKU', 'اسم المنتج', 'الفرع']).agg({
            'الكمية': 'sum',
            'اجمالي بعد الخصم': 'first'
        }).reset_index()
        
        # تجميع كميات ABC
        abc_grouped = df_abc.groupby(['رقم الطلب', 'رقم الصنف']).agg({
            'Net Sold Qty': 'sum',
            'Total Sales After VAT.': 'first'
        }).reset_index()
        abc_grouped.rename(columns={'رقم الصنف': 'SKU', 'Net Sold Qty': 'كمية ABC'}, inplace=True)
        
        # دمج البيانات
        merged = pd.merge(salla_grouped, abc_grouped, on=['رقم الطلب', 'SKU'], how='outer', suffixes=('_سلة', '_abc'))
        merged = merged.fillna(0)
        
        # حساب الفرق
        merged['الفرق'] = merged['الكمية_سلة'] - merged['كمية ABC']
        merged['حالة'] = merged['الفرق'].apply(
            lambda x: 'إضافة' if x > 0 else ('إرجاع' if x < 0 else 'مطابق')
        )
        
        # تصفية فقط الإضافات والإرجاعات
        result = merged[merged['حالة'].isin(['إضافة', 'إرجاع'])].copy()
        
        # تحميل التعديلات السابقة
        with open(init_adjustments(), "r", encoding="utf-8") as f:
            adjustments = json.load(f)
        
        # إضافة حالة التعديل لكل صف
        def get_adjustment_status(row):
            for adj in adjustments:
                if (adj['order_number'] == int(row['رقم الطلب']) and 
                    adj['sku'] == int(row['SKU']) and
                    adj['pharmacy'] == row['الفرع']):
                    return adj['status']
            return "لم يبدأ"
        
        result['حالة التعديل'] = result.apply(get_adjustment_status, axis=1)
        
        # إضافة عمود الصيدلية
        result['الصيدلية'] = result['الفرع']
        
        return result, df_salla, df_abc
    
    return None, None, None

# دالة لتسجيل التعديل
def record_adjustment(order_number, sku, pharmacy, action, status, performed_by):
    adj_file = init_adjustments()
    with open(adj_file, "r", encoding="utf-8") as f:
        adjustments = json.load(f)
    
    # البحث عن التعديل الموجود
    existing = None
    for adj in adjustments:
        if adj['order_number'] == order_number and adj['sku'] == sku and adj['pharmacy'] == pharmacy:
            existing = adj
            break
    
    if existing:
        existing['status'] = status
        existing['action'] = action
        existing['performed_by'] = performed_by
        existing['timestamp'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    else:
        adjustments.append({
            "id": len(adjustments) + 1,
            "order_number": order_number,
            "sku": sku,
            "pharmacy": pharmacy,
            "action": action,
            "status": status,
            "performed_by": performed_by,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        })
    
    with open(adj_file, "w", encoding="utf-8") as f:
        json.dump(adjustments, f, ensure_ascii=False, indent=2)

# دالة التحقق من تسجيل الدخول
def check_login(username, password, users):
    if username in users:
        return users[username]['password'] == password
    return False

# تهيئة Session State
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
if 'current_user' not in st.session_state:
    st.session_state.current_user = None
if 'user_type' not in st.session_state:
    st.session_state.user_type = None
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None
if 'df_salla' not in st.session_state:
    st.session_state.df_salla = None
if 'df_abc' not in st.session_state:
    st.session_state.df_abc = None
if 'uploaded_file_name' not in st.session_state:
    st.session_state.uploaded_file_name = None

# الشريط الجانبي
with st.sidebar:
    st.image("https://img.icons8.com/color/96/000000/pharmacy.png", width=80)
    st.title("🌟 نظام بلسم")
    st.markdown("---")
    
    if not st.session_state.logged_in:
        st.subheader("🔐 تسجيل الدخول")
        username = st.text_input("اسم المستخدم")
        password = st.text_input("كلمة المرور", type="password")
        
        users = json.load(open(init_users(), "r", encoding="utf-8"))
        
        if st.button("دخول 🚪", use_container_width=True):
            if check_login(username, password, users):
                st.session_state.logged_in = True
                st.session_state.current_user = username
                if username == "admin":
                    st.session_state.user_type = "admin"
                else:
                    st.session_state.user_type = "pharmacy"
                play_sound("success")
                st.rerun()
            else:
                st.error("❌ اسم المستخدم أو كلمة المرور غير صحيحة")
                play_sound("error")
    else:
        st.success(f"مرحباً {st.session_state.current_user} 👋")
        if st.button("تسجيل خروج 🚪", use_container_width=True):
            st.session_state.logged_in = False
            st.session_state.current_user = None
            st.session_state.user_type = None
            play_sound("info")
            st.rerun()

# المحتوى الرئيسي
if not st.session_state.logged_in:
    st.markdown("""
    <div class="main-header">
        <h1>📊 نظام مراقبة طلبات سلة و ABC</h1>
        <p>نظام متكامل لإدارة ومتابعة الطلبات والإضافات والإرجاعات</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
        <div class="stats-card">
            <div class="stats-number">🏥 17</div>
            <p>صيدلية</p>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div class="stats-card">
            <div class="stats-number">📦 1000+</div>
            <p>طلب شهرياً</p>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown("""
        <div class="stats-card">
            <div class="stats-number">⚡ 99%</div>
            <p>دقة المطابقة</p>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    st.info("👈 الرجاء تسجيل الدخول من القائمة الجانبية للاستمرار")

else:
    # تحميل البيانات إذا كان هناك ملف مرفوع
    uploaded_file = st.file_uploader("📂 رفع ملف الطلبات والفواتير", type=['xlsx'], key="file_uploader")
    
    if uploaded_file is not None:
        if st.session_state.uploaded_file_name != uploaded_file.name:
            st.session_state.processed_data, st.session_state.df_salla, st.session_state.df_abc = process_excel(uploaded_file)
            st.session_state.uploaded_file_name = uploaded_file.name
            play_sound("success")
            st.success("✅ تم معالجة الملف بنجاح!")
    
    # عرض المحتوى حسب نوع المستخدم
    if st.session_state.user_type == "admin":
        st.markdown("""
        <div class="main-header">
            <h1>👑 لوحة تحكم المدير</h1>
            <p>مراقبة جميع الإضافات والإرجاعات عبر الصيدليات</p>
        </div>
        """, unsafe_allow_html=True)
        
        if st.session_state.processed_data is not None:
            df = st.session_state.processed_data
            
            # إحصائيات سريعة
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("📊 إجمالي الطلبات", len(df['رقم الطلب'].unique()))
            with col2:
                additions = len(df[df['حالة'] == 'إضافة'])
                st.metric("➕ إضافات مطلوبة", additions)
            with col3:
                returns = len(df[df['حالة'] == 'إرجاع'])
                st.metric("➖ إرجاعات مطلوبة", returns)
            with col4:
                completed = len(df[df['حالة التعديل'] != 'لم يبدأ'])
                st.metric("✅ تم إنجازها", completed)
            
            st.markdown("---")
            
            # تبويبات للإضافات والإرجاعات
            tab1, tab2 = st.tabs(["📈 الإضافات المطلوبة", "📉 الإرجاعات المطلوبة"])
            
            with tab1:
                additions_df = df[df['حالة'] == 'إضافة'].copy()
                if len(additions_df) > 0:
                    display_cols = ['رقم الطلب', 'SKU', 'اسم المنتج', 'الكمية_سلة', 'كمية ABC', 'الفرق', 'الصيدلية', 'حالة التعديل']
                    st.dataframe(additions_df[display_cols], use_container_width=True, height=400)
                else:
                    st.info("🎉 لا توجد إضافات مطلوبة حالياً")
            
            with tab2:
                returns_df = df[df['حالة'] == 'إرجاع'].copy()
                if len(returns_df) > 0:
                    display_cols = ['رقم الطلب', 'SKU', 'اسم المنتج', 'الكمية_سلة', 'كمية ABC', 'الفرق', 'الصيدلية', 'حالة التعديل']
                    st.dataframe(returns_df[display_cols], use_container_width=True, height=400)
                else:
                    st.info("🎉 لا توجد إرجاعات مطلوبة حالياً")
            
            # سجل التعديلات
            st.markdown("---")
            st.subheader("📋 سجل التعديلات")
            with open(init_adjustments(), "r", encoding="utf-8") as f:
                adjustments = json.load(f)
            if adjustments:
                adjustments_df = pd.DataFrame(adjustments)
                st.dataframe(adjustments_df, use_container_width=True)
            else:
                st.info("لا توجد تعديلات مسجلة بعد")
        
        else:
            st.warning("⚠️ الرجاء رفع ملف الطلبات والفواتير أولاً")
    
    else:  # مستخدم صيدلية
        pharmacy_name = st.session_state.current_user
        
        st.markdown(f"""
        <div class="main-header">
            <h1>🏥 {pharmacy_name}</h1>
            <p>مراجعة وإدارة الإضافات والإرجاعات</p>
        </div>
        """, unsafe_allow_html=True)
        
        if st.session_state.processed_data is not None:
            df = st.session_state.processed_data
            pharmacy_df = df[df['الصيدلية'] == pharmacy_name].copy()
            
            if len(pharmacy_df) > 0:
                # إحصائيات الصيدلية
                col1, col2, col3 = st.columns(3)
                with col1:
                    additions = len(pharmacy_df[pharmacy_df['حالة'] == 'إضافة'])
                    st.metric("➕ إضافات مطلوبة", additions)
                with col2:
                    returns = len(pharmacy_df[pharmacy_df['حالة'] == 'إرجاع'])
                    st.metric("➖ إرجاعات مطلوبة", returns)
                with col3:
                    completed = len(pharmacy_df[pharmacy_df['حالة التعديل'] != 'لم يبدأ'])
                    st.metric("✅ تم إنجازها", completed)
                
                st.markdown("---")
                
                # تبويب الإضافات
                st.subheader("📈 الأصناف التي تحتاج إلى إضافة")
                additions_df = pharmacy_df[pharmacy_df['حالة'] == 'إضافة'].copy()
                if len(additions_df) > 0:
                    for idx, row in additions_df.iterrows():
                        col1, col2, col3, col4, col5, col6 = st.columns([2, 2, 3, 1, 1, 2])
                        with col1:
                            st.write(f"**رقم الطلب:** {int(row['رقم الطلب'])}")
                        with col2:
                            st.write(f"**SKU:** {int(row['SKU'])}")
                        with col3:
                            st.write(f"**المنتج:** {row['اسم المنتج']}")
                        with col4:
                            st.write(f"**المطلوب:** {int(row['الكمية_سلة'])}")
                        with col5:
                            st.write(f"**الموجود:** {int(row['كمية ABC'])}")
                        with col6:
                            if row['حالة التعديل'] == 'تمت الإضافة':
                                st.success("✅ تمت الإضافة")
                            else:
                                if st.button(f"➕ تمت الإضافة", key=f"add_{idx}"):
                                    record_adjustment(
                                        int(row['رقم الطلب']),
                                        int(row['SKU']),
                                        pharmacy_name,
                                        "إضافة",
                                        "تمت الإضافة",
                                        pharmacy_name
                                    )
                                    play_sound("success")
                                    st.rerun()
                        st.markdown("---")
                else:
                    st.info("🎉 لا توجد إضافات مطلوبة حالياً")
                
                # تبويب الإرجاعات
                st.subheader("📉 الأصناف التي تحتاج إلى إرجاع")
                returns_df = pharmacy_df[pharmacy_df['حالة'] == 'إرجاع'].copy()
                if len(returns_df) > 0:
                    for idx, row in returns_df.iterrows():
                        col1, col2, col3, col4, col5, col6 = st.columns([2, 2, 3, 1, 1, 2])
                        with col1:
                            st.write(f"**رقم الطلب:** {int(row['رقم الطلب'])}")
                        with col2:
                            st.write(f"**SKU:** {int(row['SKU'])}")
                        with col3:
                            st.write(f"**المنتج:** {row['اسم المنتج']}")
                        with col4:
                            st.write(f"**المطلوب:** {int(row['الكمية_سلة'])}")
                        with col5:
                            st.write(f"**الموجود:** {int(row['كمية ABC'])}")
                        with col6:
                            if row['حالة التعديل'] == 'تم الإرجاع':
                                st.success("✅ تم الإرجاع")
                            else:
                                if st.button(f"🔄 تم الإرجاع", key=f"return_{idx}"):
                                    record_adjustment(
                                        int(row['رقم الطلب']),
                                        int(row['SKU']),
                                        pharmacy_name,
                                        "إرجاع",
                                        "تم الإرجاع",
                                        pharmacy_name
                                    )
                                    play_sound("success")
                                    st.rerun()
                        st.markdown("---")
                else:
                    st.info("🎉 لا توجد إرجاعات مطلوبة حالياً")
            else:
                st.info(f"📭 لا توجد طلبات مخصصة لـ {pharmacy_name} حالياً")
        else:
            st.warning("⚠️ الرجاء رفع ملف الطلبات والفواتير أولاً")
            st.info("📌 سيتم توجيه الطلبات حسب المدينة: AL ULA → فرع 9، باقي المدن → فرع 13")

# تشغيل الصوت عند التحميل
if st.session_state.get('play_startup_sound', True):
    play_sound("info")
    st.session_state.play_startup_sound = False