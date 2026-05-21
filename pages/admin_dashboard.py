import streamlit as st
import pandas as pd
from utils.database import (
    get_latest_upload_summary, get_all_sessions, get_session_items, 
    lock_session, unlock_session, activate_session, delete_session,
    fetch_active_items, get_all_last_logins, get_completed_items,
    reopen_case_by_item_key
)
from utils.helpers import (
    is_cancelled_or_returned_status, is_pending_payment_status,
    get_tab_label
)
from utils.ui_components import render_metrics
from utils.excel_processor import process_excel

def show():
    st.markdown("""
    <div class="hero">
        <h1>👑 لوحة التحكم الإدارية</h1>
        <p>إدارة الطلبات والفواتير - متابعة الإضافات والإرجاعات - إدارة الجلسات</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns([1, 6])
    with col1:
        if st.button("🔄 تحديث الصفحة", use_container_width=True):
            st.rerun()
    
    with st.expander("📂 رفع ملف الطلبات والفواتير", expanded=True):
        uploaded_file = st.file_uploader("اختر ملف Excel (يحتوي على شيتين: 'سلة' و 'abc')", type=["xlsx"])
        if uploaded_file:
            if st.button("🔄 معالجة الملف", use_container_width=True, type="primary"):
                with st.spinner("جاري معالجة الملف..."):
                    results, upload_batch_id = process_excel(uploaded_file, st.session_state.username)
                if results is not None:
                    st.success(f"✅ تمت المعالجة بنجاح! عدد الحالات: {len(results)}")
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
    
    # إدارة الجلسات السابقة
    st.markdown('<div class="section-title">📋 إدارة الجلسات السابقة</div>', unsafe_allow_html=True)
    sessions_df = get_all_sessions()
    if not sessions_df.empty:
        for _, session in
