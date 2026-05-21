import streamlit as st
import pandas as pd
from datetime import datetime
from utils.database import (
    get_latest_upload_summary, get_all_sessions, get_session_items, 
    lock_session, unlock_session, activate_session, delete_session,
    fetch_active_items, get_all_last_logins, get_completed_items,
    reopen_case_by_item_key
)
from utils.helpers import (
    is_cancelled_or_returned_status, is_pending_payment_status,
    get_tab_label, status_pill, case_pill
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
    
    # إدارة الجلسات السابقة (نفس الكود السابق)
    st.markdown('<div class="section-title">📋 إدارة الجلسات السابقة</div>', unsafe_allow_html=True)
    sessions_df = get_all_sessions()
    # ... (نفس الكود السابق)
    
    df = fetch_active_items(include_hidden=True)
    if df.empty:
        st.info("📂 لا توجد بيانات فعالة بعد. ارفع ملف Excel من الأعلى لبدء التحليل.")
        return
    
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
    
    # فصل البيانات
    active_mask = ~filtered_df["order_status"].apply(is_cancelled_or_returned_status)
    payment_mask = filtered_df["order_status"].apply(is_pending_payment_status)
    cancelled_mask = filtered_df["order_status"].apply(is_cancelled_or_returned_status)
    
    additions_df = filtered_df[(filtered_df["case_type"] == "addition") & active_mask]
    returns_df = filtered_df[(filtered_df["case_type"] == "return") & active_mask]
    orphan_salla_df = filtered_df[(filtered_df["case_type"] == "orphan_salla") & active_mask]
    orphan_abc_df = filtered_df[(filtered_df["case_type"] == "orphan_abc") & active_mask]
    post_cutoff_df = filtered_df[filtered_df["case_type"] == "post_cutoff_abc"]
    payment_df = filtered_df[payment_mask]
    cancelled_df = filtered_df[cancelled_mask]
    
    completed_df = get_completed_items()
    if selected_branch != "الكل":
        completed_df = completed_df[completed_df["pharmacy_name"] == selected_branch]
    
    # التبويبات
    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
        get_tab_label("📈 الإضافات", len(additions_df[additions_df["status"] == "تم"]), len(additions_df)),
        get_tab_label("📉 الإرجاعات", len(returns_df[returns_df["status"] == "تم"]), len(returns_df)),
        get_tab_label("📦 طلبات بدون فاتورة", len(orphan_salla_df[orphan_salla_df["status"] == "تم"]), len(orphan_salla_df)),
        get_tab_label("🧾 فواتير بدون طلب", len(orphan_abc_df[orphan_abc_df["status"] == "تم"]), len(orphan_abc_df)),
        get_tab_label("⏰ فواتير بعد آخر طلب", len(post_cutoff_df[post_cutoff_df["status"] == "تم"]), len(post_cutoff_df)),
        get_tab_label("💰 بانتظار الدفع", 0, len(payment_df)),
        get_tab_label("⚠️ ملغي/مسترجع", 0, len(cancelled_df)),
        get_tab_label("✅ تم الانتهاء", len(completed_df), len(completed_df))
    ])
    
    def styled_frame(input_df, title=""):
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
            "order_number": "رقم الطلب", "invoice_number": "رقم الفاتورة", "sku": "SKU",
            "product_name": "المنتج", "pharmacy_name": "الفرع", "salla_qty": "كمية سلة",
            "abc_qty": "كمية ABC", "difference": "الفرق", "case_label": "نوع الحالة",
            "status": "الحالة", "performed_by": "تم بواسطة", "performed_at": "تاريخ التنفيذ",
            "order_status": "حالة الطلب", "city": "المدينة", "profile_type": "نوع البروفايل"
        })
        return display_df.style.apply(row_style, axis=1)
    
    # تظليل التبويبات الرئيسية
    with tab1:
        st.markdown('<div style="background-color:#e8f4f8; padding:10px; border-radius:10px;">', unsafe_allow_html=True)
        st.dataframe(styled_frame(additions_df), use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    with tab2:
        st.markdown('<div style="background-color:#e8f4f8; padding:10px; border-radius:10px;">', unsafe_allow_html=True)
        st.dataframe(styled_frame(returns_df), use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    with tab3:
        st.markdown('<div style="background-color:#e8f4f8; padding:10px; border-radius:10px;">', unsafe_allow_html=True)
        st.dataframe(styled_frame(orphan_salla_df), use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    with tab4:
        st.markdown('<div style="background-color:#e8f4f8; padding:10px; border-radius:10px;">', unsafe_allow_html=True)
        st.dataframe(styled_frame(orphan_abc_df), use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    with tab5:
        st.dataframe(styled_frame(post_cutoff_df), use_container_width=True)
    with tab6:
        st.dataframe(styled_frame(payment_df), use_container_width=True)
    with tab7:
        st.dataframe(styled_frame(cancelled_df), use_container_width=True)
    with tab8:
        if not completed_df.empty:
            st.dataframe(styled_frame(completed_df), use_container_width=True)
            st.markdown("#### 🔓 إعادة فتح الطلبات المكتملة")
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
    
    # آخر دخول للصيدليات
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
