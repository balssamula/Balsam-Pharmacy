import streamlit as st
import pandas as pd
from utils.database import (
    get_latest_upload_summary, get_all_sessions, get_session_items, 
    lock_session, unlock_session, activate_session, delete_session,
    fetch_active_items, get_all_last_logins, get_completed_items,
    reopen_case_by_item_key, get_tab_completed_counts
)
from utils.helpers import (
    is_cancelled_or_returned_status, is_pending_payment_status,
    get_tab_label, get_saudi_time, numeric_value
)
from utils.ui_components import render_metrics
from utils.excel_processor import process_excel

def styled_frame(input_df):
    if input_df.empty:
        return input_df
    def row_style(row):
        case_type = row.get("نوع الحالة", "")
        order_status = row.get("حالة الطلب", "")
        status = row.get("الحالة", "")
        
        if status == "تم":
            color = "background-color: #d4edda"
        elif is_cancelled_or_returned_status(order_status):
            color = "background-color: #ffe5e5"
        elif is_pending_payment_status(order_status):
            color = "background-color: #fff4d6"
        elif case_type == "إرجاع":
            color = "background-color: #ffe0df"
        elif case_type == "إضافة":
            color = "background-color: #dff1ff"
        else:
            color = "background-color: #ffe9cc"
        return [color] * len(row)
    
    display_df = input_df.copy()
    display_df = display_df.rename(columns={
        "order_number": "رقم الطلب",
        "invoice_number": "رقم الفاتورة",
        "sku": "SKU",
        "product_name": "المنتج",
        "pharmacy_name": "الفرع",
        "salla_qty": "كمية سلة",
        "abc_qty": "كمية ABC",
        "difference": "الفرق",
        "case_label": "نوع الحالة",
        "status": "الحالة",
        "performed_by": "تم بواسطة",
        "performed_at": "تاريخ التنفيذ",
        "order_status": "حالة الطلب",
        "city": "المدينة",
        "profile_type": "نوع البروفايل"
    })
    return display_df.style.apply(row_style, axis=1)

def show():
    st.markdown(f"""
    <div class="hero">
        <h1>👑 لوحة التحكم الإدارية</h1>
        <p>إدارة الطلبات والفواتير - متابعة الإضافات والإرجاعات - إدارة الجلسات</p>
        <p>🕐 آخر تحديث: {get_saudi_time()}</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns([1, 6])
    with col1:
        if st.button("🔄 تحديث", use_container_width=True):
            st.rerun()
    
    # رفع ملف الطلبات والفواتير
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
    
    # عرض آخر جلسة نشطة
    latest = get_latest_upload_summary()
    if latest:
        batch_id, file_name, uploaded_by, uploaded_at, total_cases, additions, returns, orphan_salla, orphan_abc, is_locked, session_name = latest
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
        for _, session in sessions_df.iterrows():
            col1, col2, col3, col4, col5 = st.columns([2, 2, 1.5, 1.5, 2])
            
            session_name_val = session.get('session_name', '')
            if not session_name_val or pd.isna(session_name_val):
                session_name_val = session['upload_batch_id'][:8]
            
            with col1:
                st.markdown(f"""
                <div class="session-card">
                    <strong>📅 {session_name_val}</strong><br>
                    <small>{session['file_name'][:35] if session['file_name'] else ''}</small>
                </div>
                """, unsafe_allow_html=True)
            
            with col2:
                is_active = session.get('is_active', 0)
                active_badge = "✅ نشطة" if is_active else "⏸ غير نشطة"
                st.markdown(f"""
                <div class="session-card">
                    <small>👤 {session['uploaded_by']}<br>
                    📅 {session['uploaded_at'][:16] if session['uploaded_at'] else ''}<br>
                    {active_badge}</small>
                </div>
                """, unsafe_allow_html=True)
            
            with col3:
                st.markdown(f"""
                <div class="session-card">
                    <small>📊 {int(session.get('total_cases', 0))} حالة<br>
                    ➕ {int(session.get('total_additions', 0))}<br>
                    ➖ {int(session.get('total_returns', 0))}</small>
                </div>
                """, unsafe_allow_html=True)
            
            with col4:
                is_locked = session.get('is_locked', 0)
                lock_text = "🔒 مقفلة" if is_locked else "🔓 مفتوحة"
                st.markdown(f"""
                <div class="session-card" style="text-align: center;">
                    <small>{lock_text}</small>
                </div>
                """, unsafe_allow_html=True)
            
            with col5:
                btn1, btn2, btn3, btn4 = st.columns(4)
                with btn1:
                    if not is_locked:
                        if st.button(f"🔒", key=f"lock_{session['upload_batch_id']}", help="قفل الجلسة"):
                            lock_session(session['upload_batch_id'], st.session_state.username)
                            st.rerun()
                    else:
                        if st.button(f"🔓", key=f"unlock_{session['upload_batch_id']}", help="فتح الجلسة"):
                            unlock_session(session['upload_batch_id'])
                            st.rerun()
                with btn2:
                    if not is_active:
                        if st.button(f"⭐", key=f"activate_{session['upload_batch_id']}", help="تفعيل الجلسة"):
                            activate_session(session['upload_batch_id'])
                            st.rerun()
                with btn3:
                    if st.button(f"👁️", key=f"view_{session['upload_batch_id']}", help="عرض الجلسة"):
                        st.session_state.view_session_id = session['upload_batch_id']
                        st.rerun()
                with btn4:
                    if st.button(f"🗑️", key=f"delete_{session['upload_batch_id']}", help="حذف الجلسة"):
                        delete_session(session['upload_batch_id'])
                        st.rerun()
        
        st.markdown("---")
    
    if st.session_state.get('view_session_id'):
        st.markdown(f'<div class="section-title">📄 عرض الجلسة المحددة</div>', unsafe_allow_html=True)
        session_items = get_session_items(st.session_state.view_session_id)
        if not session_items.empty:
            st.dataframe(session_items, use_container_width=True)
        if st.button("إغلاق العرض", use_container_width=True):
            del st.session_state.view_session_id
            st.rerun()
    
    df = fetch_active_items(include_hidden=True)
    if df.empty:
        st.info("📂 لا توجد بيانات فعالة بعد. ارفع ملف Excel من الأعلى لبدء التحليل.")
        return
    
    df['difference'] = df.apply(lambda row: numeric_value(row['salla_qty']) - numeric_value(row['abc_qty']), axis=1)
    render_metrics(df)
    
    # فلاتر
    col1, col2 = st.columns(2)
    with col1:
        branch_options = ["الكل"] + sorted(df["pharmacy_name"].dropna().astype(str).unique().tolist())
        selected_branch = st.selectbox("🏥 فلتر الفرع", branch_options)
    with col2:
        status_filter = st.selectbox("📌 فلتر الحالة", ["الكل", "قيد المتابعة", "تم"])
    
    filtered_df = df.copy()
    if selected_branch != "الكل":
        filtered_df = filtered_df[filtered_df["pharmacy_name"] == selected_branch]
    if status_filter != "الكل":
        filtered_df = filtered_df[filtered_df["status"] == ("تم" if status_filter == "تم" else "قيد المتابعة")]
    
    tab_completed = get_tab_completed_counts()
    
    additions_count = len(filtered_df[filtered_df["case_type"] == "addition"])
    returns_count = len(filtered_df[filtered_df["case_type"] == "return"])
    orphan_salla_count = len(filtered_df[filtered_df["case_type"] == "orphan_salla"])
    orphan_abc_count = len(filtered_df[filtered_df["case_type"] == "orphan_abc"])
    
    completed_df = get_completed_items()
    if selected_branch != "الكل":
        completed_df = completed_df[completed_df["pharmacy_name"] == selected_branch]
    
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        get_tab_label("📈 الإضافات", tab_completed.get("addition", 0), additions_count + tab_completed.get("addition", 0)),
        get_tab_label("📉 الإرجاعات", tab_completed.get("return", 0), returns_count + tab_completed.get("return", 0)),
        get_tab_label("📦 طلبات بدون فاتورة", tab_completed.get("orphan_salla", 0), orphan_salla_count + tab_completed.get("orphan_salla", 0)),
        get_tab_label("🧾 فواتير بدون طلب", tab_completed.get("orphan_abc", 0), orphan_abc_count + tab_completed.get("orphan_abc", 0)),
        get_tab_label("✅ تم الانتهاء", len(completed_df), len(completed_df))
    ])
    
    with tab1:
        additions = filtered_df[filtered_df["case_type"] == "addition"]
        st.dataframe(styled_frame(additions), use_container_width=True)
    with tab2:
        returns = filtered_df[filtered_df["case_type"] == "return"]
        st.dataframe(styled_frame(returns), use_container_width=True)
    with tab3:
        orphan_salla = filtered_df[filtered_df["case_type"] == "orphan_salla"]
        st.dataframe(styled_frame(orphan_salla), use_container_width=True)
    with tab4:
        orphan_abc = filtered_df[filtered_df["case_type"] == "orphan_abc"]
        st.dataframe(styled_frame(orphan_abc), use_container_width=True)
    with tab5:
        if not completed_df.empty:
            st.dataframe(styled_frame(completed_df), use_container_width=True)
            for idx, row in completed_df.iterrows():
                col1, col2, col3, col4, col5 = st.columns([2, 2, 2, 2, 1])
                with col1:
                    st.write(f"طلب: {row['order_number']}")
                with col2:
                    st.write(f"SKU: {row['sku']}")
                with col3:
                    st.write(f"الفرع: {row['pharmacy_name']}")
                with col4:
                    st.write(f"تم بواسطة: {row['performed_by']}")
                with col5:
                    if st.button(f"🔓 إعادة فتح", key=f"reopen_{idx}"):
                        if 'item_key' in row:
                            reopen_case_by_item_key(row['item_key'])
                            st.rerun()
                st.divider()
        else:
            st.info("لا توجد طلبات مكتملة")
    
    st.markdown('<div class="section-title">👥 آخر دخول للصيدليات</div>', unsafe_allow_html=True)
    last_logins = get_all_last_logins()
    if not last_logins.empty:
        cols = st.columns(4)
        for idx, (_, row) in enumerate(last_logins.head(8).iterrows()):
            with cols[idx % 4]:
                st.markdown(f"""
                <div class="note-card">
                    <strong>🏥 {row['pharmacy_name'][-10:]}</strong><br>
                    <span>👤 {row['pharmacist_name'] or 'غير مسجل'}</span><br>
                    <span>📅 {row['last_login'][:16] if row['last_login'] else 'لم يدخل'}</span>
                </div>
                """, unsafe_allow_html=True)
