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
    get_manager_last_login, get_login_history
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
    output = BytesIO()
    tab_colors = {
        "الإضافات": "4472C4", "الإرجاعات": "ED7D31",
        "طلبات_بدون_فاتورة": "70AD47", "فواتير_بدون_طلب": "FFC000",
        "فواتير_بعد_آخر_طلب": "9B59B6", "بانتظار_الدفع": "3498DB",
        "ملغي_ومسترجع": "E74C3C", "تم_الانتهاء": "27AE60",
        "مقارنة_الجلسات": "2A5298"
    }
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in dataframes_dict.items():
            if df is not None and not df.empty:
                df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
                worksheet = writer.sheets[sheet_name[:31]]
                header_fill = PatternFill(start_color=tab_colors.get(sheet_name, "2A5298"), 
                                         end_color=tab_colors.get(sheet_name, "2A5298"), fill_type="solid")
                header_font = Font(color="FFFFFF", bold=True, size=12)
                for col in range(1, len(df.columns) + 1):
                    cell = worksheet.cell(row=1, column=col)
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                for col in range(1, len(df.columns) + 1):
                    column_letter = get_column_letter(col)
                    max_length = 0
                    for row in range(1, len(df) + 2):
                        cell_value = worksheet.cell(row=row, column=col).value
                        if cell_value:
                            max_length = max(max_length, len(str(cell_value)))
                    worksheet.column_dimensions[column_letter].width = min(max_length + 2, 40)
                for row in range(2, len(df) + 2):
                    for col in range(1, len(df.columns) + 1):
                        cell = worksheet.cell(row=row, column=col)
                        if row % 2 == 0:
                            cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        if worksheet.cell(row=1, column=col).value == "الفرق":
                            diff_value = cell.value
                            if diff_value:
                                if diff_value > 0:
                                    cell.font = Font(color="008000", bold=True)
                                elif diff_value < 0:
                                    cell.font = Font(color="FF0000", bold=True)
            else:
                empty_df = pd.DataFrame({"ملاحظة": ["لا توجد بيانات"]})
                empty_df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    output.seek(0)
    return output.getvalue()

def compare_sessions(session1_id: str, session2_id: str) -> pd.DataFrame:
    import sqlite3
    from utils.database import DB_PATH
    conn = sqlite3.connect(DB_PATH)
    query = """
        SELECT order_number, invoice_number, sku, product_name, pharmacy_name,
               salla_qty, abc_qty, (salla_qty - abc_qty) as difference,
               case_type, order_status, abc_pharmacist_name
        FROM reconciliation_items 
        WHERE upload_batch_id IN (?, ?) AND active = 1
    """
    df = pd.read_sql_query(query, conn, params=(session1_id, session2_id))
    conn.close()
    df['required_action'] = df['difference'].apply(
        lambda x: 'إضافة' if x > 0 else ('إرجاع' if x < 0 else 'مطابق')
    )
    return df.rename(columns={
        "order_number": "رقم الطلب", "invoice_number": "رقم الفاتورة",
        "sku": "SKU", "product_name": "المنتج", "pharmacy_name": "الفرع",
        "salla_qty": "كمية سلة", "abc_qty": "كمية ABC",
        "difference": "الفرق", "order_status": "حالة الطلب",
        "abc_pharmacist_name": "الصيدلي"
    })

def styled_dataframe(input_df):
    if input_df.empty:
        return None
    def highlight_rows(row):
        if row.get('status') == "تم":
            if row.get('case_type') == "addition":
                return ['background-color: #d4edda'] * len(row)
            elif row.get('case_type') == "return":
                return ['background-color: #f8d7da'] * len(row)
            else:
                return ['background-color: #d1ecf1'] * len(row)
        return [''] * len(row)
    
    display_df = input_df.copy()
    display_df = display_df.rename(columns={
        "order_number": "رقم الطلب", "invoice_number": "رقم الفاتورة",
        "sku": "SKU", "product_name": "المنتج", "pharmacy_name": "الفرع",
        "salla_qty": "كمية سلة", "abc_qty": "كمية ABC",
        "difference": "الفرق", "order_status": "حالة الطلب",
        "case_label": "نوع الحالة", "status": "الحالة"
    })
    return display_df.style.apply(highlight_rows, axis=1)

def render_table_with_click(df, tab_name):
    if df.empty:
        st.success("لا توجد بيانات في هذا القسم.")
        return
    styled_df = styled_dataframe(df)
    if styled_df is not None:
        event = st.dataframe(
            styled_df,
            use_container_width=True,
            height=400,
            selection_mode="single-row",
            on_select="rerun"
        )
        if event.selection.rows:
            selected_idx = event.selection.rows[0]
            if 0 <= selected_idx < len(df):
                row = df.iloc[selected_idx]
                item_key = row.get('item_key', '')
                st.markdown(f"""
                <div style="background:#f0f2f6;border-radius:10px;padding:1rem;margin-top:1rem;border-right:4px solid #1f7a8c;">
                    <h4 style="margin:0 0 0.5rem 0;">🛠️ إجراءات الصف المحدد</h4>
                    <p><strong>📋 رقم الطلب:</strong> {row['order_number']} | 
                    <strong>🏷️ SKU:</strong> {row['sku']} | 
                    <strong>📦 المنتج:</strong> {row['product_name'][:50]}</p>
                </div>
                """, unsafe_allow_html=True)
                cols = st.columns(5)
                is_hidden = row.get('hidden_from_pharmacy', 0) == 1
                if cols[0].button("🙈 إخفاء من الصيدلية" if not is_hidden else "👁️ إظهار للصيدلية", key=f"hide_{tab_name}_{selected_idx}", use_container_width=True):
                    if is_hidden:
                        unhide_item_from_pharmacy(item_key)
                    else:
                        hide_item_from_pharmacy(item_key, st.session_state.username)
                    st.rerun()
                is_locked = row.get('is_item_locked', 0) == 1
                if cols[1].button("🔒 قفل التعديل" if not is_locked else "🔓 فتح التعديل", key=f"lock_{tab_name}_{selected_idx}", use_container_width=True):
                    if is_locked:
                        unlock_item(item_key)
                    else:
                        lock_item(item_key, st.session_state.username)
                    st.rerun()
                if row['status'] == "تم":
                    if cols[2].button("🔄 إعادة فتح", key=f"reopen_{tab_name}_{selected_idx}", use_container_width=True):
                        reopen_case_by_item_key(item_key)
                        st.rerun()
                note = cols[3].text_input("📝 ملحوظة", value=row.get('pharmacist_note', ''), key=f"note_{tab_name}_{selected_idx}")
                if cols[4].button("💾 حفظ", key=f"save_note_{tab_name}_{selected_idx}", use_container_width=True):
                    save_case_note(row['order_number'], row['sku'], row['pharmacy_name'], row['case_type'], note)
                    st.rerun()

def show():
    st.markdown("""
    <div class="hero">
        <h1>👑 لوحة التحكم الإدارية</h1>
        <p>إدارة الطلبات والفواتير - متابعة الإضافات والإرجاعات - إدارة الجلسات</p>
    </div>
    """, unsafe_allow_html=True)
    
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
    
    # ========== إدارة الجلسات السابقة ==========
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
    
    # ========== مقارنة الجلسات ==========
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
    
    # ========== جلب البيانات النشطة ==========
    df = fetch_active_items(include_hidden=True)
    if df.empty:
        st.info("📂 لا توجد بيانات فعالة بعد. ارفع ملف Excel من الأعلى لبدء التحليل.")
        return
    
    # إحصائيات سريعة - استبعاد الملغي والمسترجع
    active_mask = ~df["order_status"].apply(is_cancelled_or_returned_status)
    active_df = df[active_mask]
    
    col1, col2, col3, col4, col5, col6, col7 = st.columns(7)
    with col1:
        st.metric("📊 إجمالي الحالات", len(active_df))
    with col2:
        st.metric("➕ إضافات", len(active_df[active_df["case_type"] == "addition"]))
    with col3:
        st.metric("➖ إرجاعات", len(active_df[active_df["case_type"] == "return"]))
    with col4:
        st.metric("📦 طلبات بدون فاتورة", len(active_df[active_df["case_type"] == "orphan_salla"]))
    with col5:
        st.metric("🧾 فواتير بدون طلب", len(active_df[active_df["case_type"] == "orphan_abc"]))
    with col6:
        st.metric("⏰ فواتير بعد آخر طلب", len(active_df[active_df["case_type"] == "post_cutoff_abc"]))
    with col7:
        st.metric("✅ تم إنجازها", len(df[df["status"] == "تم"]))
    
    # فلاتر
    col1, col2, col3 = st.columns(3)
    with col1:
        branch_options = ["الكل"] + sorted(df["pharmacy_name"].dropna().astype(str).unique().tolist())
        selected_branch = st.selectbox("🏥 فلتر الفرع", branch_options)
    with col2:
        status_filter = st.selectbox("📌 فلتر حالة الإجراء", ["الكل", "قيد المتابعة", "تم"])
    with col3:
        order_status_options = ["الكل", "تم التوصيل", "ملغي", "مسترجع", "بانتظار الدفع", "تم الاستلام من فرع"]
        selected_order_status = st.selectbox("📋 فلتر حالة الطلب", order_status_options)
    
    # تطبيق الفلاتر
    filtered_df = df.copy()
    if selected_branch != "الكل":
        filtered_df = filtered_df[filtered_df["pharmacy_name"] == selected_branch]
    if status_filter != "الكل":
        filtered_df = filtered_df[filtered_df["status"] == ("تم" if status_filter == "تم" else "قيد المتابعة")]
    if selected_order_status != "الكل":
        if selected_order_status == "تم الاستلام من فرع":
            filtered_df = filtered_df[filtered_df["order_status"].str.contains("تم الاستلام من فرع", na=False)]
        else:
            filtered_df = filtered_df[filtered_df["order_status"] == selected_order_status]
    
    # فصل البيانات - استبعاد الملغي والمسترجع من التبويبات الرئيسية
    active_mask_filtered = ~filtered_df["order_status"].apply(is_cancelled_or_returned_status)
    payment_mask = filtered_df["order_status"].apply(is_pending_payment_status)
    cancelled_mask = filtered_df["order_status"].apply(is_cancelled_or_returned_status)
    
    additions_df = filtered_df[(filtered_df["case_type"] == "addition") & active_mask_filtered]
    returns_df = filtered_df[(filtered_df["case_type"] == "return") & active_mask_filtered]
    orphan_salla_df = filtered_df[(filtered_df["case_type"] == "orphan_salla") & active_mask_filtered]
    orphan_abc_df = filtered_df[(filtered_df["case_type"] == "orphan_abc") & active_mask_filtered]
    post_cutoff_df = filtered_df[(filtered_df["case_type"] == "post_cutoff_abc") & active_mask_filtered]
    payment_df = filtered_df[payment_mask]
    cancelled_df = filtered_df[cancelled_mask]
    
    completed_df = get_completed_items()
    if selected_branch != "الكل":
        completed_df = completed_df[completed_df["pharmacy_name"] == selected_branch]

    col1, col2 = st.columns([1, 6])
    with col1:
        if st.button("🔄 تحديث الصفحة", use_container_width=True):
            st.rerun()
    with col2:
        if st.button("📥 تصدير إلى Excel", use_container_width=True):
            st.session_state.show_export = True
    
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
                    
    # تصدير Excel
    if st.session_state.get('show_export', False):
        export_data = {
            "الإضافات": additions_df, "الإرجاعات": returns_df,
            "طلبات_بدون_فاتورة": orphan_salla_df, "فواتير_بدون_طلب": orphan_abc_df,
            "فواتير_بعد_آخر_طلب": post_cutoff_df, "بانتظار_الدفع": payment_df,
            "ملغي_ومسترجع": cancelled_df, "تم_الانتهاء": completed_df
        }
        excel_data = export_to_excel(export_data)
        st.download_button(
            "📥 تحميل التقرير",
            data=excel_data,
            file_name=f"balsam_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            use_container_width=True,
        )
        st.session_state.show_export = False
    
    # ========== التبويبات الملونة ==========
    st.markdown("""
    <style>
    .stTabs [data-baseweb="tab-list"] button:nth-child(1) { background-color: #4472C4; color: white; border-radius: 10px 10px 0 0; }
    .stTabs [data-baseweb="tab-list"] button:nth-child(2) { background-color: #ED7D31; color: white; border-radius: 10px 10px 0 0; }
    .stTabs [data-baseweb="tab-list"] button:nth-child(3) { background-color: #70AD47; color: white; border-radius: 10px 10px 0 0; }
    .stTabs [data-baseweb="tab-list"] button:nth-child(4) { background-color: #FFC000; color: white; border-radius: 10px 10px 0 0; }
    .stTabs [data-baseweb="tab-list"] button:nth-child(5) { background-color: #6c757d; color: white; border-radius: 10px 10px 0 0; }
    .stTabs [data-baseweb="tab-list"] button:nth-child(6) { background-color: #6c757d; color: white; border-radius: 10px 10px 0 0; }
    .stTabs [data-baseweb="tab-list"] button:nth-child(7) { background-color: #6c757d; color: white; border-radius: 10px 10px 0 0; }
    .stTabs [data-baseweb="tab-list"] button:nth-child(8) { background-color: #6c757d; color: white; border-radius: 10px 10px 0 0; }
    /* التبويب النشط */
    .stTabs [data-baseweb="tab-list"] button[aria-selected="true"] {
        transform: translateY(-2px) !important;
        box-shadow: 0 4px 8px rgba(0,0,0,0.2) !important;
    }
    /* التبويب غير النشط */
    .stTabs [data-baseweb="tab-list"] button[aria-selected="false"] {
        opacity: 0.85 !important;
    }
    /* تأثير hover */
    .stTabs [data-baseweb="tab-list"] button:hover {
        transform: translateY(-2px) !important;
        opacity: 1 !important;
    }
    </style>
    """, unsafe_allow_html=True)
    
    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
        f"📈 الإضافات ({len(additions_df)})",
        f"📉 الإرجاعات ({len(returns_df)})",
        f"📦 طلبات بدون فاتورة ({len(orphan_salla_df)})",
        f"🧾 فواتير بدون طلب ({len(orphan_abc_df)})",
        f"⏰ فواتير بعد آخر طلب ({len(post_cutoff_df)})",
        f"💰 بانتظار الدفع ({len(payment_df)})",
        f"⚠️ ملغي/مسترجع ({len(cancelled_df)})",
        f"✅ تم الانتهاء ({len(completed_df)})"
    ])
    
    with tab1:
        render_table_with_click(additions_df, "additions")
    with tab2:
        render_table_with_click(returns_df, "returns")
    with tab3:
        render_table_with_click(orphan_salla_df, "orphan_salla")
    with tab4:
        render_table_with_click(orphan_abc_df, "orphan_abc")
    with tab5:
        render_table_with_click(post_cutoff_df, "post_cutoff")
    with tab6:
        render_table_with_click(payment_df, "payment")
    with tab7:
        render_table_with_click(cancelled_df, "cancelled")
    with tab8:
        if not completed_df.empty:
            render_table_with_click(completed_df, "completed")
        else:
            st.info("لا توجد طلبات مكتملة")
    
    # ========== آخر دخول للصيدليات ==========
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

    # ========== معلومات المدير العام ==========
    st.markdown('<div class="section-title">👑 معلومات المدير العام</div>', unsafe_allow_html=True)
    manager_info = get_manager_last_login()
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"""
        <div class="note-card">
            <strong>👤 آخر دخول للمدير العام:</strong><br>
            📅 {manager_info['last_login']}<br>
            🌐 IP: {manager_info['last_ip']}
        </div>
        """, unsafe_allow_html=True)
    with col2:
        login_history = get_login_history(10)
        if not login_history.empty:
            st.markdown("### 📋 آخر محاولات الدخول")
            st.dataframe(login_history, use_container_width=True)
    
    st.markdown("---")
