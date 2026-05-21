import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from utils.database import fetch_active_items, get_completed_items, get_tab_completed_counts
from utils.helpers import (
    is_cancelled_or_returned_status, is_pending_payment_status, 
    get_branch_number, get_branch_location, get_tab_label, numeric_value,
    get_saudi_time
)
from utils.ui_components import render_metrics, render_completed_table

def export_to_excel(dataframes_dict: dict, pharmacy_name: str) -> bytes:
    """تصدير البيانات إلى ملف Excel مع تنسيق احترافي"""
    output = BytesIO()
    
    # الألوان لكل تبويب
    tab_colors = {
        "الإضافات": "4472C4",      # أزرق
        "الإرجاعات": "ED7D31",      # برتقالي
        "طلبات_بدون_فاتورة": "70AD47",  # أخضر
        "فواتير_بدون_طلب": "FFC000",    # ذهبي
        "فواتير_بعد_آخر_طلب": "9B59B6",  # بنفسجي
        "بانتظار_الدفع": "3498DB",   # أزرق فاتح
        "ملغي_ومسترجع": "E74C3C",    # أحمر
        "تم_الانتهاء": "27AE60"      # أخضر داكن
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

def render_case_cards_pharmacy(df: pd.DataFrame, allow_actions: bool, pharmacist_name: str, pharmacy_name: str):
    if df.empty:
        st.success("لا توجد حالات في هذا القسم.")
        return

    for idx, row in df.iterrows():
        diff_value = numeric_value(row['difference'])
        
        if diff_value > 0:
            required_action = "إضافة"
            action_color = "#28a745"
            button_label = "✅ تأكيد الإضافة"
        elif diff_value < 0:
            required_action = "إرجاع"
            action_color = "#dc3545"
            button_label = "🔄 تأكيد الإرجاع"
        else:
            required_action = "مطابق"
            action_color = "#6c757d"
            button_label = None
        
        order_status = row.get('order_status', 'غير متوفرة')
        unique_key = f"{row['case_type']}_{row['order_number']}_{row['sku']}_{idx}"
        
        with st.container():
            st.markdown(f"""
            <div style="background:#f8f9fa;border-radius:16px;padding:1rem;margin-bottom:1rem;border-right:4px solid #1f7a8c;box-shadow:0 2px 8px rgba(0,0,0,0.05);">
                <div style="display:flex;justify-content:space-between;margin-bottom:0.5rem;">
                    <span style="background:#dff1ff;color:#0f5488;padding:0.2rem 0.8rem;border-radius:20px;font-size:0.8rem;">{row['case_label']}</span>
                    <span style="color:#6c757d;font-size:0.8rem;">📅 {row['order_date'][:16] if row['order_date'] else ''}</span>
                </div>
                <div style="display:flex;flex-wrap:wrap;gap:1rem;">
                    <div style="flex:2;">
                        <strong>📋 رقم الطلب:</strong> {row['order_number']}<br>
                        <strong>🏷️ SKU:</strong> {row['sku']}<br>
                        <strong>📦 المنتج:</strong> {row['product_name'][:60]}
                    </div>
                    <div style="flex:1;">
                        <strong>📊 الكميات:</strong><br>
                        🛒 سلة: {int(row['salla_qty']) if pd.notna(row['salla_qty']) else 0}<br>
                        📄 ABC: {int(row['abc_qty']) if pd.notna(row['abc_qty']) else 0}<br>
                        <strong>📊 الفرق:</strong> <span style="color:{'#28a745' if diff_value > 0 else '#dc3545' if diff_value < 0 else '#6c757d'};font-weight:bold;">{'+' if diff_value > 0 else ''}{diff_value}</span>
                    </div>
                    <div style="flex:1.5;">
                        <strong>🧾 الفاتورة/الصيدلي:</strong><br>
                        {row['invoice_number']}/{row.get('abc_pharmacist_name', 'غير معروف')}<br>
                        <strong>📌 حالة الطلب:</strong> <span style="color:#d9534f;">{order_status}</span><br>
                        <strong>🎯 المطلوب:</strong> <span style="color:{action_color};">{required_action}</span>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            note_key = f"note_{unique_key}"
            note_value = st.text_area("📝 ملحوظة الصيدلي", value=row.get("pharmacist_note", "") or "", key=note_key, height=60)
            
            btn_col1, btn_col2 = st.columns([1, 4])
            with btn_col1:
                if st.button("💾 حفظ", key=f"save_{unique_key}", use_container_width=True):
                    from utils.database import save_case_note
                    save_case_note(row['order_number'], row['sku'], pharmacy_name, row['case_type'], note_value)
                    st.rerun()
            
            if allow_actions and row["status"] != "تم" and button_label and row["case_type"] in {"addition", "return", "orphan_salla", "orphan_abc"}:
                with btn_col2:
                    if st.button(button_label, key=f"done_{unique_key}", use_container_width=True):
                        from utils.database import save_case_note, mark_case_done
                        save_case_note(row['order_number'], row['sku'], pharmacy_name, row['case_type'], note_value)
                        mark_case_done(row['order_number'], row['sku'], pharmacy_name, row['case_type'], pharmacist_name)
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

    df = fetch_active_items(pharmacy_name, include_hidden=False)
    
    if df.empty:
        st.info("📭 لا توجد حالات نشطة لهذا الفرع حاليًا.")
        completed_df = get_completed_items(pharmacy_name)
        if not completed_df.empty:
            st.markdown("---")
            st.markdown('<div class="section-title">✅ الطلبات المكتملة</div>', unsafe_allow_html=True)
            render_completed_table(completed_df, is_admin=False)
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

    additions_df = df[(df["case_type"] == "addition") & active_mask].copy()
    returns_df = df[(df["case_type"] == "return") & active_mask].copy()
    orphan_salla_df = df[(df["case_type"] == "orphan_salla") & active_mask].copy()
    orphan_abc_df = df[(df["case_type"] == "orphan_abc") & active_mask].copy()
    post_cutoff_df = df[(df["case_type"] == "post_cutoff_abc") & active_mask].copy()
    cancelled_df = df[df["order_status"].apply(is_cancelled_or_returned_status)].copy()
    payment_df = df[df["order_status"].apply(is
