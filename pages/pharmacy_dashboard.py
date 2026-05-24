import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from utils.database import (
    fetch_active_items, get_completed_items, get_tab_completed_counts, 
    get_old_orders, get_old_invoices, get_old_invoices_stats,
    check_duplicate_across_branches, get_all_duplicate_items
)
from utils.helpers import (
    is_cancelled_or_returned_status, is_pending_payment_status, 
    get_branch_number, get_branch_location, get_tab_label, numeric_value,
    get_saudi_time
)
from utils.ui_components import render_metrics, render_completed_table

def export_to_excel(dataframes_dict: dict, pharmacy_name: str) -> bytes:
    """تصدير البيانات إلى ملف Excel مع تنسيق احترافي"""
    output = BytesIO()
    
    tab_colors = {
        "الإضافات": "4472C4",
        "الإرجاعات": "ED7D31",
        "طلبات_بدون_فاتورة": "70AD47",
        "فواتير_بدون_طلب": "FFC000",
        "فواتير_بعد_آخر_طلب": "9B59B6",
        "بانتظار_الدفع": "3498DB",
        "ملغي_ومسترجع": "E74C3C",
        "تم_الانتهاء": "27AE60",
        "فواتير_قديمة": "6c757d"
    }
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in dataframes_dict.items():
            if df is not None and not df.empty:
                df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
                
                worksheet = writer.sheets[sheet_name[:31]]
                
                header_fill = PatternFill(start_color=tab_colors.get(sheet_name, "2A5298"), 
                                         end_color=tab_colors.get(sheet_name, "2A5298"), 
                                         fill_type="solid")
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
                empty_df = pd.DataFrame({"ملاحظة": ["لا توجد بيانات في هذا التبويب"]})
                empty_df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    
    output.seek(0)
    return output.getvalue()

@st.fragment
def render_single_case_card(row, idx, allow_actions, pharmacist_name, pharmacy_name):
    """عرض بطاقة حالة واحدة مع تفاصيلها وأزرار الإجراءات"""
    
    # تعيين لون ومسمى شارة التمييز (Badge) بناءً على نوع الحالة الفعلي
    case_type = row.get('case_type', '')
    
    if case_type == 'addition':
        badge_text = "➕ إضافة مخزنية عادية"
        badge_style = "background:#dff1ff; color:#084298;"
    elif case_type == 'orphan_salla':
        badge_text = "🛒 طلب مبيعات مفقود الفاتورة"
        badge_style = "background:#fff3cd; color:#856404; font-weight:bold;"
    elif case_type == 'return':
        badge_text = "🔄 إرجاع مخزني عادي"
        badge_style = "background:#ffe0df; color:#491217;"
    elif case_type == 'orphan_abc':
        badge_text = "📄 فاتورة توريد مفقودة الطلب"
        badge_style = "background:#f8d7da; color:#721c24; font-weight:bold;"
    elif case_type == 'post_cutoff_abc':
        badge_text = "⏰ فاتورة بعد آخر طلب"
        badge_style = "background:#e2e8f0; color:#475569;"
    else:
        badge_text = row.get('case_label', 'حالة عامة')
        badge_style = "background:#e2e8f0; color:#475569;"

    # حساب الفروقات ديناميكياً لضمان دقة الأرقام
    salla_numeric = int(row.get('salla_qty', 0)) if pd.notna(row.get('salla_qty', 0)) else 0
    abc_numeric = int(row.get('abc_qty', 0)) if pd.notna(row.get('abc_qty', 0)) else 0
    diff_value = salla_numeric - abc_numeric

    order_status = row.get('order_status', 'غير متوفرة')
    invoice_date = row.get('invoice_date', '')
    order_date = row.get('order_date', '')
    unique_key = f"{case_type}_{row.get('order_number', '')}_{row.get('sku', '')}_{idx}"
    
    # التحقق من وجود مكررات في فروع أخرى
    order_number = str(row.get('order_number', ''))
    sku = str(row.get('sku', ''))
    duplicates = check_duplicate_across_branches(order_number, sku, pharmacy_name)
    
    duplicate_warning = ""
    if duplicates:
        duplicate_warning = f"""
        <div style="background:#fff3cd; border-right:4px solid #ffc107; padding:0.5rem; margin-top:0.5rem; border-radius:8px;">
            <span style="color:#856404;">⚠️ <strong>تنبيه:</strong> هذا المنتج (SKU: {sku}) موجود أيضاً في الفروع التالية:</span><br>
        """
        for dup in duplicates:
            duplicate_warning += f"""
            <span style="font-size:0.85rem;">• {dup['pharmacy']} (الحالة: {dup['status']}, نوع: {dup['case_type']})</span><br>
            """
        duplicate_warning += "</div>"

    # عرض تصميم البطاقة مع تاريخ الفاتورة والتنبيهات
    st.markdown(f"""
    <div style="background:#f8f9fa; border-radius:16px; padding:1.2rem; margin-bottom:0.8rem; border-right:5px solid #1f7a8c; box-shadow:0 4px 12px rgba(0,0,0,0.03);">
        <div style="display:flex; justify-content:space-between; margin-bottom:0.8rem; align-items:center;">
            <span style="{badge_style} padding:0.3rem 1rem; border-radius:20px; font-size:0.85rem;">{badge_text}</span>
            <div style="text-align:left;">
                <span style="color:#6c757d; font-size:0.8rem; margin-left:1rem;">📅 تاريخ الطلب: {order_date[:16] if order_date else 'غير محدد'}</span>
                <span style="color:#6c757d; font-size:0.8rem;">📅 تاريخ الفاتورة: {invoice_date[:16] if invoice_date else 'غير محدد'}</span>
            </div>
        </div>
        <div style="display:flex; flex-wrap:wrap; gap:1rem;">
            <div style="flex:2; min-width:200px;">
                <strong>📋 رقم الطلب / الفاتورة:</strong> {row.get('order_number', '') if pd.notna(row.get('order_number')) else row.get('invoice_number', '')}<br>
                <strong>🏷️ SKU:</strong> {row.get('sku', '')}<br>
                <strong>📦 المنتج:</strong> {row.get('product_name', '')[:60]}
            </div>
            <div style="flex:1; min-width:120px;">
                <strong>📊 الكميات المكتشفة:</strong><br>
                🛒 سلة: {salla_numeric}<br>
                📄 ABC: {abc_numeric}<br>
                <strong>📊 الفرق الفعلي:</strong> <span style="color:{'#28a745' if diff_value > 0 else '#dc3545' if diff_value < 0 else '#6c757d'}; font-weight:bold;">{'+' if diff_value > 0 else ''}{diff_value}</span>
            </div>
            <div style="flex:1.5; min-width:160px;">
                <strong>🧾 الفاتورة/الصيدلي:</strong><br>
                {row.get('invoice_number', '')}/{row.get('abc_pharmacist_name', 'غير معروف')}<br>
                <strong>📌 حالة الطلب:</strong> <span style="color:#d9534f;">{order_status}</span><br>
                <strong>🎯 الإجراء المطلوب:</strong> <span style="color:{'#28a745' if diff_value > 0 else '#dc3545' if diff_value < 0 else '#6c757d'}; font-weight:bold;">{'إضافة' if diff_value > 0 else 'إرجاع' if diff_value < 0 else 'مطابق'}</span>
            </div>
        </div>
        {duplicate_warning}
    </div>
    """, unsafe_allow_html=True)
    
    # حقل الملاحظات والأزرار التفاعلية
    note_key = f"note_{unique_key}"
    note_value = st.text_area("📝 ملحوظة الصيدلي", value=row.get("pharmacist_note", "") or "", key=note_key, height=60)
    
    btn_col1, btn_col2 = st.columns([1, 4])
    with btn_col1:
        if st.button("💾 حفظ", key=f"save_{unique_key}", use_container_width=True):
            from utils.database import save_case_note
            save_case_note(row['order_number'], row['sku'], pharmacy_name, case_type, note_value)
            st.toast("📋 تم حفظ الملاحظة بنجاح!", icon="💾")

    if allow_actions and row.get("status") != "تم" and diff_value != 0 and case_type in {"addition", "return", "orphan_salla", "orphan_abc"}:
        button_label = "✅ تأكيد الإضافة" if diff_value > 0 else "🔄 تأكيد الإرجاع"
        with btn_col2:
            if st.button(button_label, key=f"done_{unique_key}", use_container_width=True):
                from utils.database import save_case_note, mark_case_done
                save_case_note(row['order_number'], row['sku'], pharmacy_name, case_type, note_value)
                mark_case_done(row['order_number'], row['sku'], pharmacy_name, case_type, pharmacist_name)
                st.toast("✅ تم تأكيد الحالة وإغلاقها!", icon="🚀")
                st.rerun()

def render_case_cards_pharmacy(df: pd.DataFrame, allow_actions: bool, pharmacist_name: str, pharmacy_name: str):
    """عرض بطاقات الحالات للصيدلي"""
    if df.empty:
        st.success("🎉 لا توجد حالات في هذا القسم.")
        return

    for idx, row in df.iterrows():
        render_single_case_card(row, idx, allow_actions, pharmacist_name, pharmacy_name)
        st.markdown("---")

def render_old_invoices_pharmacy(old_invoices_df, pharmacy_name, pharmacist_name):
    """عرض الفواتير القديمة للصيدلي"""
    if old_invoices_df.empty:
        return
    
    st.warning(f"⚠️ توجد {len(old_invoices_df)} فاتورة قديمة (أكثر من 6 أشهر) لم يتم إكمالها")
    
    for idx, row in old_invoices_df.iterrows():
        days_old = int(row['days_old'])
        if days_old > 365:
            card_color = "#ffcccc"
            badge = "🔴 قديم جداً"
        elif days_old > 180:
            card_color = "#ffe0cc"
            badge = "🟠 قديم"
        else:
            card_color = "#fff3cd"
            badge = "🟡 يحتاج مراجعة"
        
        diff_value = numeric_value(row['difference'])
        required_action = "إضافة" if diff_value > 0 else "إرجاع" if diff_value < 0 else "مطابق"
        invoice_date = row.get('invoice_date', '')
        
        with st.container():
            st.markdown(f"""
            <div style="background:{card_color};border-radius:16px;padding:1rem;margin-bottom:1rem;border-right:4px solid #dc3545;">
                <div style="display:flex;justify-content:space-between;margin-bottom:0.5rem;">
                    <span style="background:#dc3545;color:white;padding:0.2rem 0.8rem;border-radius:20px;font-size:0.8rem;">{badge}</span>
                    <span style="color:#6c757d;">📅 تاريخ الفاتورة: {invoice_date[:16] if invoice_date else ''} | ⏰ {days_old} يوم</span>
                </div>
                <div style="display:flex;flex-wrap:wrap;gap:1rem;">
                    <div style="flex:2;">
                        <strong>📋 رقم الفاتورة:</strong> {row['invoice_number']}<br>
                        <strong>🏷️ SKU:</strong> {row['sku']}<br>
                        <strong>📦 المنتج:</strong> {row['product_name'][:60]}
                    </div>
                    <div style="flex:1;">
                        <strong>📊 الكميات:</strong><br>
                        🛒 سلة: {int(row['salla_qty'])}<br>
                        📄 ABC: {int(row['abc_qty'])}<br>
                        <strong>📊 الفرق:</strong> {diff_value}
                    </div>
                    <div style="flex:1.5;">
                        <strong>🧾 رقم الطلب/الصيدلي:</strong><br>
                        {row['order_number']}/{row['abc_pharmacist_name'] or 'غير معروف'}<br>
                        <strong>📌 حالة الطلب:</strong> {row['order_status']}<br>
                        <strong>🎯 المطلوب:</strong> {required_action}
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            note_key = f"old_invoice_note_{idx}"
            note_value = st.text_area("📝 ملحوظة", value=row.get('pharmacist_note', ''), key=note_key, height=60)
            
            col1, col2 = st.columns([1, 4])
            with col1:
                if st.button("💾 حفظ", key=f"save_old_invoice_{idx}"):
                    from utils.database import save_case_note
                    save_case_note(row['order_number'], row['sku'], pharmacy_name, row['case_type'], note_value)
                    st.success("✅ تم حفظ الملحوظة")
                    st.rerun()
            
            if row['status'] != "تم":
                with col2:
                    if st.button("✅ تأكيد الإكمال", key=f"complete_old_invoice_{idx}"):
                        from utils.database import mark_case_done
                        mark_case_done(row['order_number'], row['sku'], pharmacy_name, row['case_type'], pharmacist_name)
                        st.success("✅ تم تأكيد إكمال الفاتورة")
                        st.rerun()
            st.markdown("---")

def render_old_orders_pharmacy(old_orders_df, pharmacy_name, pharmacist_name):
    """عرض الطلبات القديمة للصيدلي"""
    if old_orders_df.empty:
        return
    
    for idx, row in old_orders_df.iterrows():
        days_old = int(row['days_old'])
        if days_old > 365:
            card_color = "#ffcccc"
            badge = "🔴 قديم جداً"
        elif days_old > 180:
            card_color = "#ffe0cc"
            badge = "🟠 قديم"
        else:
            card_color = "#fff3cd"
            badge = "🟡 يحتاج مراجعة"
        
        diff_value = numeric_value(row['difference'])
        required_action = "إضافة" if diff_value > 0 else "إرجاع" if diff_value < 0 else "مطابق"
        
        with st.container():
            st.markdown(f"""
            <div style="background:{card_color};border-radius:16px;padding:1rem;margin-bottom:1rem;border-right:4px solid #dc3545;">
                <div style="display:flex;justify-content:space-between;margin-bottom:0.5rem;">
                    <span style="background:#dc3545;color:white;padding:0.2rem 0.8rem;border-radius:20px;font-size:0.8rem;">{badge}</span>
                    <span style="color:#6c757d;">📅 {row['order_date'][:16] if row['order_date'] else ''} | ⏰ {days_old} يوم</span>
                </div>
                <div style="display:flex;flex-wrap:wrap;gap:1rem;">
                    <div style="flex:2;">
                        <strong>📋 رقم الطلب:</strong> {row['order_number']}<br>
                        <strong>🏷️ SKU:</strong> {row['sku']}<br>
                        <strong>📦 المنتج:</strong> {row['product_name'][:60]}
                    </div>
                    <div style="flex:1;">
                        <strong>📊 الكميات:</strong><br>
                        🛒 سلة: {int(row['salla_qty'])}<br>
                        📄 ABC: {int(row['abc_qty'])}<br>
                        <strong>📊 الفرق:</strong> {diff_value}
                    </div>
                    <div style="flex:1.5;">
                        <strong>🧾 الفاتورة/الصيدلي:</strong><br>
                        {row['invoice_number']}/{row['abc_pharmacist_name'] or 'غير معروف'}<br>
                        <strong>📌 حالة الطلب:</strong> {row['order_status']}<br>
                        <strong>🎯 المطلوب:</strong> {required_action}
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            note_key = f"old_note_{idx}"
            note_value = st.text_area("📝 ملحوظة", value=row.get('pharmacist_note', ''), key=note_key, height=60)
            
            col1, col2 = st.columns([1, 4])
            with col1:
                if st.button("💾 حفظ", key=f"save_old_{idx}"):
                    from utils.database import save_case_note
                    save_case_note(row['order_number'], row['sku'], pharmacy_name, row['case_type'], note_value)
                    st.success("✅ تم حفظ الملحوظة")
                    st.rerun()
            
            if row['status'] != "تم":
                with col2:
                    if st.button("✅ تأكيد الإكمال", key=f"complete_old_{idx}"):
                        from utils.database import mark_case_done
                        mark_case_done(row['order_number'], row['sku'], pharmacy_name, row['case_type'], pharmacist_name)
                        st.success("✅ تم تأكيد إكمال الطلب")
                        st.rerun()
            st.markdown("---")

def show():
    pharmacy_name = st.session_state.username
    pharmacist_name = st.session_state.pharmacist_name or ""
    branch_number = get_branch_number(pharmacy_name)
    branch_location = get_branch_location(branch_number)

    st.markdown(f"""
    <div class="hero">
        <h1>🏥 {pharmacy_name}</h1>
        <p>فرع رقم {branch_number} | الموقع: {branch_location} | الصيدلي: {pharmacist_name}</p>
        <p>🕐 آخر تحديث: {get_saudi_time()}</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("🔄 تحديث الصفحة", use_container_width=True):
            st.rerun()
    with col2:
        if st.button("📥 تصدير إلى Excel", use_container_width=True):
            st.session_state.show_export_pharmacy = True

    # التحقق من وجود مكررات
    duplicate_items = get_all_duplicate_items(pharmacy_name)
    if not duplicate_items.empty:
        with st.expander("⚠️ تنبيه: منتجات مكررة عبر الفروع", expanded=True):
            st.warning(f"تم العثور على {len(duplicate_items)} منتج مكرر (نفس SKU ورقم الطلب) موجود في أكثر من فرع")
            st.dataframe(
                duplicate_items[['order_number', 'sku', 'product_name', 'pharmacy_name', 'case_type', 'status']],
                use_container_width=True
            )

    df = fetch_active_items(pharmacy_name, include_hidden=False)
    
    if df.empty:
        st.info("📭 لا توجد حالات نشطة لهذا الفرع حاليًا.")
        completed_df = get_completed_items(pharmacy_name)
        if not completed_df.empty:
            st.markdown("---")
            st.markdown('<div class="section-title">✅ الطلبات المكتملة</div>', unsafe_allow_html=True)
            render_completed_table(completed_df, is_admin=False)
        
        # عرض الطلبات القديمة والفواتير القديمة
        st.markdown("---")
        st.markdown('<div class="section-title">📅 العناصر القديمة (أكثر من 6 أشهر)</div>', unsafe_allow_html=True)
        
        tab_old_orders, tab_old_invoices = st.tabs(["📦 الطلبات القديمة", "🧾 الفواتير القديمة"])
        
        with tab_old_orders:
            months_orders = st.selectbox("عدد الأشهر للبحث (طلبات)", [3, 6, 9, 12, 18, 24], index=1, key="old_orders_months_empty")
            old_orders_df = get_old_orders(pharmacy_name=pharmacy_name, months=months_orders)
            if not old_orders_df.empty:
                render_old_orders_pharmacy(old_orders_df, pharmacy_name, pharmacist_name)
            else:
                st.success(f"🎉 لا توجد طلبات قديمة (أكثر من {months_orders} أشهر)")
        
        with tab_old_invoices:
            months_invoices = st.selectbox("عدد الأشهر للبحث (فواتير)", [3, 6, 9, 12, 18, 24], index=1, key="old_invoices_months_empty")
            old_invoices_df = get_old_invoices(pharmacy_name=pharmacy_name, months=months_invoices)
            if not old_invoices_df.empty:
                render_old_invoices_pharmacy(old_invoices_df, pharmacy_name, pharmacist_name)
            else:
                st.success(f"🎉 لا توجد فواتير قديمة (أكثر من {months_invoices} أشهر)")
        return

    is_locked = False
    if 'is_locked' in df.columns and not df.empty:
        is_locked = df['is_locked'].iloc[0] == 1
    allow_actions = not is_locked

    active_mask = ~df["order_status"].apply(is_cancelled_or_returned_status)
    active_df = df[active_mask]
    
    total = len(active_df)
    additions = len(active_df[active_df["case_type"] == "addition"])
    returns = len(active_df[active_df["case_type"] == "return"])
    orphan_salla = len(active_df[active_df["case_type"] == "orphan_salla"])
    orphan_abc = len(active_df[active_df["case_type"] == "orphan_abc"])
    post_cutoff = len(active_df[active_df["case_type"] == "post_cutoff_abc"])
    completed = len(df[df["status"] == "تم"])
    
    col1, col2, col3, col4, col5, col6, col7 = st.columns(7)
    with col1:
        st.metric("📊 إجمالي الحالات", total)
    with col2:
        st.metric("➕ إضافات", additions)
    with col3:
        st.metric("➖ إرجاعات", returns)
    with col4:
        st.metric("📦 طلبات بدون فاتورة", orphan_salla)
    with col5:
        st.metric("🧾 فواتير بدون طلب", orphan_abc)
    with col6:
        st.metric("⏰ فواتير بعد آخر طلب", post_cutoff)
    with col7:
        st.metric("✅ تم إنجازها", completed)

    # تجهيز البيانات للتبويبات
    branch_add_df = df[df['case_type'].isin(['addition', 'orphan_salla']) & active_mask].copy()
    total_additions_merged = len(branch_add_df)
    completed_additions_merged = len(branch_add_df[branch_add_df["status"] == "تم"])
    pending_additions_merged = total_additions_merged - completed_additions_merged
    
    branch_ret_df = df[df['case_type'].isin(['return', 'orphan_abc']) & active_mask].copy()
    total_returns_merged = len(branch_ret_df)
    completed_returns_merged = len(branch_ret_df[branch_ret_df["status"] == "تم"])
    pending_returns_merged = total_returns_merged - completed_returns_merged
    
    post_cutoff_df = df[(df["case_type"] == "post_cutoff_abc") & active_mask].copy()
    total_post_cutoff = len(post_cutoff_df)
    completed_post_cutoff = len(post_cutoff_df[post_cutoff_df["status"] == "تم"])
    
    payment_df = df[df["order_status"].apply(is_pending_payment_status) & active_mask].copy()
    total_payment = len(payment_df)
    
    cancelled_df = df[df["order_status"].apply(is_cancelled_or_returned_status)].copy()
    total_cancelled = len(cancelled_df)
    
    completed_df = get_completed_items(pharmacy_name)
    total_completed = len(completed_df)
    
    # جلب الفواتير القديمة
    old_invoices_df = get_old_invoices(pharmacy_name=pharmacy_name, months=6)
    old_orders_df = get_old_orders(pharmacy_name=pharmacy_name, months=6)
    
    # إعداد أعداد التبويبات
    tab_additions_count = f"{completed_additions_merged}/{total_additions_merged}" if total_additions_merged > 0 else "0"
    tab_returns_count = f"{completed_returns_merged}/{total_returns_merged}" if total_returns_merged > 0 else "0"
    
    # تصدير Excel
    if st.session_state.get('show_export_pharmacy', False):
        additions_merged = df[df['case_type'].isin(['addition', 'orphan_salla'])].copy()
        returns_merged = df[df['case_type'].isin(['return', 'orphan_abc'])].copy()
        
        additions_merged['نوع التفصيلي'] = additions_merged['case_type'].map({
            'addition': 'إضافة عادية', 'orphan_salla': 'طلب بدون فاتورة'
        })
        returns_merged['نوع التفصيلي'] = returns_merged['case_type'].map({
            'return': 'إرجاع عادي', 'orphan_abc': 'فاتورة بدون طلب'
        })
        
        export_data = {
            "الإضافات_والطلبات_بدون_فاتورة": additions_merged,
            "الإرجاعات_والفواتير_بدون_طلب": returns_merged,
            "فواتير_بعد_آخر_طلب": post_cutoff_df,
            "بانتظار_الدفع": payment_df,
            "ملغي_ومسترجع": cancelled_df,
            "تم_الانتهاء": completed_df,
            "الطلبات_القديمة": old_orders_df,
            "الفواتير_القديمة": old_invoices_df
        }
        excel_data = export_to_excel(export_data, pharmacy_name)
        st.download_button(
            "📥 تحميل التقرير",
            data=excel_data,
            file_name=f"balsam_pharmacy_{pharmacy_name.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.session_state.show_export_pharmacy = False
    
    # تلوين التبويبات
    st.markdown("""
    <style>
    button[data-baseweb="tab"]:nth-child(1) { background-color: #4472C4; color: white; border-radius: 10px 10px 0 0; }
    button[data-baseweb="tab"]:nth-child(2) { background-color: #ED7D31; color: white; border-radius: 10px 10px 0 0; }
    button[data-baseweb="tab"]:nth-child(3) { background-color: #9B59B6; color: white; border-radius: 10px 10px 0 0; }
    button[data-baseweb="tab"]:nth-child(4) { background-color: #3498DB; color: white; border-radius: 10px 10px 0 0; }
    button[data-baseweb="tab"]:nth-child(5) { background-color: #E74C3C; color: white; border-radius: 10px 10px 0 0; }
    button[data-baseweb="tab"]:nth-child(6) { background-color: #27AE60; color: white; border-radius: 10px 10px 0 0; }
    button[data-baseweb="tab"]:nth-child(7) { background-color: #6c757d; color: white; border-radius: 10px 10px 0 0; }
    button[data-baseweb="tab"]:nth-child(8) { background-color: #6c757d; color: white; border-radius: 10px 10px 0 0; }
    button[data-baseweb="tab"][aria-selected="true"] { transform: translateY(-2px) !important; box-shadow: 0 4px 8px rgba(0,0,0,0.2) !important; }
    button[data-baseweb="tab"][aria-selected="false"] { opacity: 0.85 !important; }
    button[data-baseweb="tab"]:hover { transform: translateY(-2px) !important; opacity: 1 !important; }
    </style>
    """, unsafe_allow_html=True)
    
    # ========== التبويبات مع إضافة تبويب الفواتير القديمة ==========
    tab_additions, tab_returns, tab_post_cutoff, tab_payment, tab_cancelled, tab_completed, tab_old_orders, tab_old_invoices = st.tabs([
        f"📥 الإضافات والطلبات المفقودة ({tab_additions_count})",
        f"📤 الإرجاعات والفواتير المعلقة ({tab_returns_count})",
        f"⏰ فواتير بعد آخر طلب ({completed_post_cutoff}/{total_post_cutoff})" if total_post_cutoff > 0 else f"⏰ فواتير بعد آخر طلب (0)",
        f"💰 بانتظار الدفع ({total_payment})",
        f"⚠️ ملغي/مسترجع ({total_cancelled})",
        f"✅ تم الانتهاء ({total_completed})",
        f"📦 طلبات قديمة ({len(old_orders_df)})",
        f"🧾 فواتير قديمة ({len(old_invoices_df)})"
    ])
    
    with tab_additions:
        st.markdown("""
        <div style="background: #dff1ff20; padding: 0.5rem 1rem; border-radius: 12px; margin-bottom: 0.75rem;">
            <span style="font-size: 0.9rem;">🔵 <strong>الإضافات العادية</strong>: كمية الطلب أعلى من الفاتورة | 🟡 <strong>طلبات بدون فاتورة</strong>: طلب موجود في سلة وغير موجود في ABC</span>
        </div>
        """, unsafe_allow_html=True)
        
        if not branch_add_df.empty:
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("📊 إجمالي الحالات", total_additions_merged)
            with col2:
                st.metric("✅ تم الإنجاز", completed_additions_merged)
            with col3:
                st.metric("⏳ قيد المتابعة", pending_additions_merged)
            st.markdown("---")
        
        render_case_cards_pharmacy(branch_add_df, allow_actions, pharmacist_name, pharmacy_name)
    
    with tab_returns:
        st.markdown("""
        <div style="background: #ffe0df20; padding: 0.5rem 1rem; border-radius: 12px; margin-bottom: 0.75rem;">
            <span style="font-size: 0.9rem;">🔴 <strong>الإرجاعات العادية</strong>: كمية الفاتورة أعلى من الطلب | 🟡 <strong>فواتير بدون طلب</strong>: فاتورة موجودة في ABC وغير موجودة في سلة</span>
        </div>
        """, unsafe_allow_html=True)
        
        if not branch_ret_df.empty:
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("📊 إجمالي الحالات", total_returns_merged)
            with col2:
                st.metric("✅ تم الإنجاز", completed_returns_merged)
            with col3:
                st.metric("⏳ قيد المتابعة", pending_returns_merged)
            st.markdown("---")
        
        render_case_cards_pharmacy(branch_ret_df, allow_actions, pharmacist_name, pharmacy_name)
    
    with tab_post_cutoff:
        if not post_cutoff_df.empty:
            col1, col2 = st.columns(2)
            with col1:
                st.metric("📊 إجمالي الحالات", total_post_cutoff)
            with col2:
                st.metric("✅ تم الإنجاز", completed_post_cutoff)
            st.markdown("---")
        render_case_cards_pharmacy(post_cutoff_df, False, pharmacist_name, pharmacy_name)
    
    with tab_payment:
        if not payment_df.empty:
            st.info(f"💰 عدد الطلبات التي تنتظر الدفع: {total_payment}")
            st.markdown("---")
        render_case_cards_pharmacy(payment_df, False, pharmacist_name, pharmacy_name)
    
    with tab_cancelled:
        if not cancelled_df.empty:
            st.warning(f"⚠️ عدد الطلبات الملغية أو المسترجعة: {total_cancelled}")
            st.markdown("---")
        render_case_cards_pharmacy(cancelled_df, False, pharmacist_name, pharmacy_name)
    
    with tab_completed:
        render_completed_table(completed_df, is_admin=False)
    
    with tab_old_orders:
        st.markdown("### 📅 الطلبات القديمة (أكثر من 6 أشهر)")
        
        months_orders = st.selectbox("عدد الأشهر للبحث (طلبات)", [3, 6, 9, 12, 18, 24], index=1, key="old_orders_months")
        old_orders_dynamic = get_old_orders(pharmacy_name=pharmacy_name, months=months_orders)
        
        if not old_orders_dynamic.empty:
            render_old_orders_pharmacy(old_orders_dynamic, pharmacy_name, pharmacist_name)
        else:
            st.success(f"🎉 لا توجد طلبات قديمة (أكثر من {months_orders} أشهر)")
    
    with tab_old_invoices:
        st.markdown("""
        <div style="background: linear-gradient(135deg, #1a1a1a, #333); padding: 1rem; border-radius: 10px; margin-bottom: 1rem;">
            <h3 style="color: white; margin: 0;">🧾 الفواتير القديمة (أكثر من 6 أشهر)</h3>
            <p style="color: #ccc; margin: 0.5rem 0 0 0;">⚠️ الفواتير التي مر عليها أكثر من 6 أشهر والموجودة في ABC ولكن لم تكتمل</p>
        </div>
        """, unsafe_allow_html=True)
        
        months_invoices = st.selectbox("عدد الأشهر للبحث (فواتير)", [3, 6, 9, 12, 18, 24], index=1, key="old_invoices_months")
        old_invoices_dynamic = get_old_invoices(pharmacy_name=pharmacy_name, months=months_invoices)
        
        if not old_invoices_dynamic.empty:
            render_old_invoices_pharmacy(old_invoices_dynamic, pharmacy_name, pharmacist_name)
            
            # زر تصدير الفواتير القديمة
            if st.button("📥 تصدير الفواتير القديمة إلى Excel", use_container_width=True):
                excel_data = export_to_excel({"الفواتير_القديمة": old_invoices_dynamic}, pharmacy_name)
                st.download_button(
                    "📥 تحميل التقرير",
                    data=excel_data,
                    file_name=f"old_invoices_{pharmacy_name.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    use_container_width=True
                )
        else:
            st.success(f"🎉 لا توجد فواتير قديمة (أكثر من {months_invoices} أشهر)")
