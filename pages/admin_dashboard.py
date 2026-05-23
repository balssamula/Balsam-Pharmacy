import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from utils.database import (
    get_latest_upload_summary, get_all_sessions, get_session_items, 
    lock_session, unlock_session, activate_session, delete_session,
    fetch_active_items, get_all_last_logins, get_completed_items,
    reopen_case_by_item_key, hide_item_from_pharmacy, unhide_item_from_pharmacy,
    lock_item, unlock_item, save_case_note,
    get_manager_last_login, get_login_history, get_client_ip
)
from utils.helpers import (
    is_cancelled_or_returned_status, is_pending_payment_status,
    get_tab_label, numeric_value
)
from utils.excel_processor import process_excel

# تهيئة حالة العرض
if 'show_sessions' not in st.session_state:
    st.session_state.show_sessions = True
if 'show_comparison' not in st.session_state:
    st.session_state.show_comparison = True

def export_to_excel(dataframes_dict: dict) -> bytes:
    # ... (نفس الكود السابق)
    pass

def compare_sessions(session1_id: str, session2_id: str) -> pd.DataFrame:
    # ... (نفس الكود السابق)
    pass

def show():
    st.markdown("""
    <div class="hero">
        <h1>👑 لوحة التحكم الإدارية</h1>
        <p>إدارة الطلبات والفواتير - متابعة الإضافات والإرجاعات - إدارة الجلسات</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1, 5])
    with col1:
        if st.button("🔄 تحديث الصفحة", use_container_width=True):
            st.rerun()
    with col2:
        if st.button("📥 تصدير إلى Excel", use_container_width=True):
            st.session_state.show_export = True
    
    # ========== معلومات المدير العام ==========
    st.markdown('<div class="section-title">👑 معلومات المدير العام</div>', unsafe_allow_html=True)
    
    manager_info = get_manager_last_login()
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"""
        <div class="note-card">
            <strong>👤 آخر دخول للمدير العام:</strong><br>
            📅 {manager_info['last_login']}<br>
            🌐 IP: {manager_info['last_ip']}<br>
            👤 الاسم: {manager_info['pharmacist_name']}
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        login_history = get_login_history(10)
        if not login_history.empty:
            st.markdown("### 📋 آخر محاولات الدخول")
            st.dataframe(login_history, use_container_width=True)
        else:
            st.info("لا توجد سجلات دخول بعد")
    
    st.markdown("---")
    
    with st.expander("📂 رفع ملف الطلبات والفواتير", expanded=True):
        uploaded_file = st.file_uploader("اختر ملف Excel", type=["xlsx"])
        if uploaded_file:
            if st.button("🔄 معالجة الملف", use_container_width=True, type="primary"):
                with st.spinner("جاري معالجة الملف..."):
                    results, upload_batch_id = process_excel(uploaded_file, st.session_state.username)
                if results is not None:
                    st.success(f"✅ تمت المعالجة بنجاح!")
                    st.balloons()
                    st.rerun()
    
    latest = get_latest_upload_summary()
    if latest:
        batch_id, file_name, uploaded_by, uploaded_at, total_cases, additions, returns, orphan_salla, orphan_abc, post_cutoff, is_locked, session_name = latest
        lock_status = "🔒 مقفلة" if is_locked else "🔓 مفتوحة"
        st.markdown(f"""
        <div class="note-card">
            <strong>📋 الجلسة النشطة:</strong> {session_name or 'غير مسماة'} &nbsp; | &nbsp;
            <strong>الملف:</strong> {file_name} &nbsp; | &nbsp;
            <strong>بواسطة:</strong> {uploaded_by} &nbsp; | &nbsp;
            <strong>التاريخ:</strong> {uploaded_at[:16] if uploaded_at else ''} &nbsp; | &nbsp;
            <strong>الحالة:</strong> {lock_status}
        </div>
        """, unsafe_allow_html=True)
    
    # ========== إدارة الجلسات السابقة مع إمكانية الإخفاء ==========
    col1, col2 = st.columns([1, 5])
    with col1:
        btn_label = "🙈 إخفاء الجلسات" if st.session_state.show_sessions else "👁️ إظهار الجلسات"
        if st.button(btn_label, use_container_width=True):
            st.session_state.show_sessions = not st.session_state.show_sessions
            st.rerun()

    if st.session_state.show_sessions:
        st.markdown('<div class="section-title">📋 إدارة الجلسات السابقة</div>', unsafe_allow_html=True)
        sessions_df = get_all_sessions()
        if not sessions_df.empty:
            for _, session in sessions_df.iterrows():
                col1, col2, col3, col4, col5 = st.columns([2, 2, 1.5, 1.5, 2])
                session_name_val = session.get('session_name', session['upload_batch_id'][:8])
                with col1:
                    st.markdown(f'<div class="session-card"><strong>📅 {session_name_val}</strong><br><small>{session["file_name"][:35]}</small></div>', unsafe_allow_html=True)
                with col2:
                    is_active = session.get('is_active', 0)
                    active_badge = "✅ نشطة" if is_active else "⏸ غير نشطة"
                    st.markdown(f'<div class="session-card"><small>👤 {session["uploaded_by"]}<br>📅 {session["uploaded_at"][:16]}<br>{active_badge}</small></div>', unsafe_allow_html=True)
                with col3:
                    st.markdown(f'<div class="session-card"><small>📊 {int(session.get("total_cases", 0))} حالة<br>➕ {int(session.get("total_additions", 0))}<br>➖ {int(session.get("total_returns", 0))}</small></div>', unsafe_allow_html=True)
                with col4:
                    is_locked = session.get('is_locked', 0)
                    lock_text = "🔒 مقفلة" if is_locked else "🔓 مفتوحة"
                    st.markdown(f'<div class="session-card" style="text-align:center;"><small>{lock_text}</small></div>', unsafe_allow_html=True)
                with col5:
                    btn1, btn2, btn3, btn4 = st.columns(4)
                    with btn1:
                        if not is_locked:
                            if st.button(f"🔒", key=f"lock_{session['upload_batch_id']}"):
                                lock_session(session['upload_batch_id'], st.session_state.username)
                                st.rerun()
                        else:
                            if st.button(f"🔓", key=f"unlock_{session['upload_batch_id']}"):
                                unlock_session(session['upload_batch_id'])
                                st.rerun()
                    with btn2:
                        if not is_active:
                            if st.button(f"⭐", key=f"activate_{session['upload_batch_id']}"):
                                activate_session(session['upload_batch_id'])
                                st.rerun()
                    with btn3:
                        if st.button(f"👁️", key=f"view_{session['upload_batch_id']}"):
                            st.session_state.view_session_id = session['upload_batch_id']
                            st.rerun()
                    with btn4:
                        if st.button(f"🗑️", key=f"delete_{session['upload_batch_id']}"):
                            delete_session(session['upload_batch_id'])
                            st.rerun()
            st.markdown("---")
    else:
        st.info("📂 إدارة الجلسات مخفية. اضغط على زر 'إظهار الجلسات' لعرضها.")
    
    # ========== مقارنة الجلسات مع إمكانية الإخفاء ==========
    col1, col2 = st.columns([1, 5])
    with col1:
        btn_label2 = "🙈 إخفاء المقارنة" if st.session_state.show_comparison else "👁️ إظهار المقارنة"
        if st.button(btn_label2, use_container_width=True):
            st.session_state.show_comparison = not st.session_state.show_comparison
            st.rerun()

    if st.session_state.show_comparison:
        st.markdown('<div class="section-title">🔄 مقارنة الجلسات</div>', unsafe_allow_html=True)
        sessions_list = get_all_sessions()
        if not sessions_list.empty:
            session_options = {f"{row['session_name']} ({row['uploaded_at'][:16]})": row['upload_batch_id'] 
                              for _, row in sessions_list.iterrows()}
            col1, col2, col3 = st.columns([2, 2, 1])
            with col1:
                session1 = st.selectbox("اختر الجلسة الأولى", list(session_options.keys()), key="session1")
            with col2:
                session2 = st.selectbox("اختر الجلسة الثانية", list(session_options.keys()), key="session2")
            with col3:
                if st.button("📊 مقارنة", use_container_width=True):
                    with st.spinner("جاري المقارنة..."):
                        comparison_df = compare_sessions(session_options[session1], session_options[session2])
                        st.session_state.comparison_result = comparison_df
                        st.success(f"✅ تمت المقارنة!")
            if st.session_state.get('comparison_result') is not None:
                st.dataframe(st.session_state.comparison_result, use_container_width=True)
                excel_data = export_to_excel({"مقارنة_الجلسات": st.session_state.comparison_result})
                st.download_button("📥 تحميل تقرير المقارنة", data=excel_data, 
                    file_name=f"session_comparison_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    else:
        st.info("📊 قسم مقارنة الجلسات مخفي. اضغط على زر 'إظهار المقارنة' لعرضه.")
    
    if st.session_state.get('view_session_id'):
        st.markdown(f'<div class="section-title">📄 عرض الجلسة المحددة</div>', unsafe_allow_html=True)
        session_items = get_session_items(st.session_state.view_session_id)
        if not session_items.empty:
            st.dataframe(session_items, use_container_width=True)
        if st.button("إغلاق العرض", use_container_width=True):
            del st.session_state.view_session_id
            st.rerun()
    
    # باقي الكود (جلب البيانات، الفلاتر، التبويبات) كما هو...
    df = fetch_active_items(include_hidden=True)
    if df.empty:
        st.info("📂 لا توجد بيانات فعالة بعد. ارفع ملف Excel من الأعلى لبدء التحليل.")
        return
    
    # ... باقي الكود كما هو ...
