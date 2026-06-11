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
    check_duplicate_across_branches, get_all_duplicate_items,
    save_case_note, mark_case_done
)
from utils.helpers import (
    is_cancelled_or_returned_status, is_pending_payment_status, 
    get_branch_number, get_branch_location, get_tab_label, numeric_value,
    get_saudi_time
)
from utils.ui_components import render_metrics, render_completed_table

def to_safe_int(val):
    """تحويل آمن ومطلق لأي قيمة نصية أو فارغة إلى عدد صحيح لمنع الانهيار الحسابي"""
    if pd.isna(val) or str(val).strip() in ["", "nan", "None"]:
        return 0
    try:
        return int(float(str(val).strip()))
    except:
        return 0
        
def export_to_excel(dataframes_dict: dict, pharmacy_name: str) -> bytes:
    """تصدير البيانات إلى ملف Excel مع تنسيق واختيار الأعمدة بشكل احترافي"""
    output = BytesIO()
    
    tab_colors = {
        "الاضافات والطلبات المفقودة": "4472C4",
        "الارجاعات والزيادات": "ED7D31",
        "فواتير معلقة بين الفروع": "9B59B6",
        "فواتير بعد اخر طلب": "9B59B6",
        "بانتظار الدفع": "3498DB",
        "الملغيات والمسترجعات": "E74C3C",
        "الفواتير القديمة (أرشيف)": "6c757d",
        "الطلبات القديمة": "6c757d"
    }
    
    columns_mapping = {
        'order_date': 'تاريخ الطلب',
        'order_number': 'رقم الطلب',
        'invoice_date': 'تاريخ الفاتورة',
        'invoice_number': 'رقم الفاتورة',
        'customer_phone': 'رقم جوال العميل',
        'sku': 'رقم المنتج SKU',
        'product_name': 'اسم المنتج',
        'salla_qty': 'كمية سلة',
        'abc_qty': 'كمية ABC',
        'difference': 'الفرق',
        'order_status': 'حالة الطلب',
        'profile_type': 'نوع البروفايل',
        'abc_pharmacist_name': 'الصيدلي المسؤول',
        'pharmacist_note': 'ملاحظات الصيدلية',
        'status': 'حالة التسوية',
        'case_type': 'نوع الحالة',
        'case_reason': 'سبب الحالة',
        'city': 'المدينة',
        'item_key': 'المفتاح الشامل للصنف'
    }

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in dataframes_dict.items():
            if df is not None and not df.empty:
                available_cols = [col for col in columns_mapping.keys() if col in df.columns]
                df_filtered = df[available_cols].copy()
                df_filtered = df_filtered.rename(columns=columns_mapping)
                df_filtered.to_excel(writer, sheet_name=sheet_name[:31], index=False)
                worksheet = writer.sheets[sheet_name[:31]]
                header_fill = PatternFill(start_color=tab_colors.get(sheet_name, "2A5298"), fill_type="solid")
                header_font = Font(color="FFFFFF", bold=True)
                for col in range(1, len(df_filtered.columns) + 1):
                    cell = worksheet.cell(row=1, column=col)
                    cell.fill = header_fill
                    cell.font = header_font
                    worksheet.column_dimensions[get_column_letter(col)].width = 22
            else:
                pd.DataFrame({"ملاحظة": ["لا توجد بيانات لهذا التبويب"]}).to_excel(writer, sheet_name=sheet_name[:31], index=False)
    output.seek(0)
    return output.getvalue()

def export_to_excel_brief(dataframes_dict: dict) -> bytes:
    """تصدير ملف إكسيل مختصر مقتصراً فقط وحصرياً على التبويبات الثلاثة الأساسية بناءً على طلبك"""
    output = BytesIO()
    allowed_tabs = ["الاضافات والطلبات المفقودة", "الارجاعات والزيادات", "فواتير معلقة بين الفروع"]
    
    tab_colors = {
        "الاضافات والطلبات المفقودة": "4472C4",
        "الارجاعات والزيادات": "ED7D31",
        "فواتير معلقة بين الفروع": "9B59B6"
    }
    
    target_columns = {
        'order_date': 'تاريخ الطلب',
        'invoice_date': 'تاريخ الفاتورة',
        'order_number': 'رقم الطلب',
        'invoice_number': 'رقم الفاتورة',
        'customer_phone': 'رقم جوال العميل',
        'sku': 'رقم المنتج',
        'product_name': 'اسم المنتج',
        'salla_qty': 'كمية سلة',
        'abc_qty': 'كمية abc',
        'diff_qty': 'الفرق',
        'order_status': 'حالة الطلب',
        'abc_pharmacist_name': 'اسم الصيدلي',
        'profile_type': 'نوع البروفايل'
    }
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name in allowed_tabs:
            df = dataframes_dict.get(sheet_name)
            if df is not None and not df.empty:
                df_brief = df.copy()
                salla_q = pd.to_numeric(df_brief.get('salla_qty', 0), errors='coerce').fillna(0).astype(int)
                abc_q = pd.to_numeric(df_brief.get('abc_qty', 0), errors='coerce').fillna(0).astype(int)
                df_brief['diff_qty'] = salla_q - abc_q
                
                for col in target_columns.keys():
                    if col not in df_brief.columns: df_brief[col] = "N/A"
                        
                df_brief = df_brief[list(target_columns.keys())].rename(columns=target_columns)
                df_brief.to_excel(writer, sheet_name=sheet_name[:31], index=False)
                worksheet = writer.sheets[sheet_name[:31]]
                header_fill = PatternFill(start_color=tab_colors.get(sheet_name, "2A5298"), fill_type="solid")
                header_font = Font(color="FFFFFF", bold=True)
                for col in range(1, len(df_brief.columns) + 1):
                    worksheet.cell(row=1, column=col).fill = header_fill
                    worksheet.cell(row=1, column=col).font = header_font
                    worksheet.column_dimensions[get_column_letter(col)].width = 20
            else:
                pd.DataFrame({"ملاحظة": ["لا توجد بيانات معلقة حالياً"]}).to_excel(writer, sheet_name=sheet_name[:31], index=False)
                
    output.seek(0)
    return output.getvalue()

def render_single_case_card(row, idx, allow_actions, pharmacist_name, pharmacy_name, tab_id=""):
    salla_numeric = int(row.get('salla_qty', 0)) if pd.notna(row.get('salla_qty', 0)) else 0
    abc_numeric = int(row.get('abc_qty', 0)) if pd.notna(row.get('abc_qty', 0)) else 0
    diff_value = salla_numeric - abc_numeric
    case_type = row.get('case_type', '')

    if diff_value > 0:
        badge_text, badge_color, badge_text_color, diff_style, required_action = "إضافة مخزنية عادية ➕", "#dff1ff", "#084298", "color: #28a745; font-weight: bold;", "<span style='color: #28a745; font-weight: bold;'>إضافة</span>"
    elif diff_value < 0:
        badge_text, badge_color, badge_text_color, diff_style, required_action = "إرجاع مخزني عادي 🔄", "#ffe0df", "#491217", "color: #dc3545; font-weight: bold;", "<span style='color: #dc3545; font-weight: bold;'>إرجاع</span>"
    else:
        diff_style, required_action = "color: #6c757d; font-weight: bold;", "<span style='color: #6c757d; font-weight: bold;'>مطابق</span>"
        if case_type == 'orphan_salla': badge_text, badge_color, badge_text_color = "طلب مبيعات مفقود الفاتورة 🛒", "#fff3cd", "#856404"
        elif case_type == 'orphan_abc': badge_text, badge_color, badge_text_color = "فاتورة توريد مفقودة الطلب 📄", "#f8d7da", "#721c24"
        else: badge_text, badge_color, badge_text_color = "حالة تسوية عامة", "#e2e8f0", "#475569"
   
    order_status = row.get('order_status', 'غير متوفرة')
    invoice_date = row.get('invoice_date', '')
    order_date = row.get('order_date', '')
    order_number = str(row.get('order_number', ''))
    sku = str(row.get('sku', ''))
    
    duplicates = []
    try: duplicates = check_duplicate_across_branches(order_number, sku, pharmacy_name)
    except: pass
    
    with st.container():
        st.markdown(f"""
        <div style="background:#f8f9fa; border-radius:16px; padding:1rem; margin-bottom:1rem; border-right:5px solid #1f7a8c; box-shadow:0 2px 8px rgba(0,0,0,0.05);">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:0.75rem; flex-wrap:wrap;">
                <span style="background:{badge_color}; color:{badge_text_color}; padding:0.25rem 0.75rem; border-radius:20px; font-size:0.8rem; font-weight:bold;">{badge_text}</span>
                <div>
                    <span style="color:#6c757d; font-size:0.75rem;">📅 الطلب: {order_date[:16] if order_date else 'غير محدد'}</span>
                    <span style="color:#6c757d; font-size:0.75rem; margin-right:0.5rem;">📅 الفاتورة: {invoice_date[:16] if invoice_date else 'غير محدد'}</span>
                </div>
            </div>
        """, unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns([2, 1, 1.5])
        with col1:
            st.markdown(f"- **📋 رقم الطلب:** {row.get('order_number', 'N/A')}\n- **🏷️ SKU / رقم المنتج:** {row.get('sku', 'N/A')}\n- **📦 المنتج:** {str(row.get('product_name', 'N/A'))[:60]}")
            if case_type in ['addition', 'orphan_salla']:
                phone = row.get('customer_phone', 'N/A')
                if pd.notna(phone) and str(phone).strip() not in ["", "nan", "None"]: st.markdown(f"- **📱 جوال العميل:** `{phone}`")
        with col2:
            st.markdown(f"- **🛒 كمية سلة:** {salla_numeric}\n- **📄 كمية ABC:** {abc_numeric}\n- **📊 الفرق:** <span style='{diff_style}'>{'+' if diff_value > 0 else ''}{diff_value}</span>\n- **🎯 المطلوب:** {required_action}", unsafe_allow_html=True)
        with col3:
            st.markdown(f"- **🧾 رقم الفاتورة:** {row.get('invoice_number', 'N/A')}\n- **👤 الصيدلي:** {row.get('abc_pharmacist_name', 'غير معروف')}\n- **⚙️ حالة الطلب:** {order_status}")
            if case_type in ['return', 'orphan_abc']:
                profile = row.get('profile_type', 'N/A')
                if pd.notna(profile) and str(profile).strip() not in ["", "nan", "None"]: st.markdown(f"- **📄 نوع البروفايل:** `{profile}`")
            
        if duplicates:
            dup_warning_html = '<div style="background:#fff3cd; border-right:4px solid #ff9800; padding:0.75rem; margin-top:0.75rem; border-radius:10px; margin-bottom:0.75rem;"><span style="color:#856404; font-weight:bold;">⚠️ تنبيه هام: هذا الصنف مكرر في فروع أخرى!</span>'
            for dup in duplicates: dup_warning_html += f'<div style="font-size:0.85rem; color:#66521a;">🏥 <strong>{dup.get("pharmacy", "غير معروف")}</strong> | الإجراء: {dup.get("status", "غير معروف")}</div>'
            st.markdown(dup_warning_html + '</div>', unsafe_allow_html=True)
            
        st.markdown("</div>", unsafe_allow_html=True)
        
        note_key = f"note_{tab_id}_{case_type}_{idx}_{row.get('order_number', '')}_{row.get('sku', '')}"
        note_value = st.text_area("📝 ملحوظة الصيدلي", value=row.get("pharmacist_note", "") or "", key=note_key, height=60)
        
        btn_col1, btn_col2, btn_col3 = st.columns([1, 1.5, 1.5])
        with btn_col1:
            if st.button("💾 حفظ الملحوظة", key=f"save_{note_key}", use_container_width=True):
                save_case_note(row['order_number'], row['sku'], pharmacy_name, case_type, note_value)
                st.toast("📋 تم حفظ الملاحظة بنجاح!", icon="💾")
                
        if allow_actions and row.get("status") != "تم":
            if case_type == "branch_conflict":
                with btn_col2:
                    if st.button("📥 تأكيد الإضافة (تم البيع من فرعي)", key=f"conf_add_{note_key}", use_container_width=True):
                        save_case_note(row['order_number'], row['sku'], pharmacy_name, case_type, f"[فرع صحيح] | {note_value}")
                        mark_case_done(row['order_number'], row['sku'], pharmacy_name, case_type, pharmacist_name)
                        st.toast("✅ تم الاعتماد!"); st.rerun()
                with btn_col3:
                    if st.button("🔄 تأكيد الإرجاع (ليس من فرعي)", key=f"conf_ret_{note_key}", use_container_width=True):
                        save_case_note(row['order_number'], row['sku'], pharmacy_name, case_type, f"[فرع مخطئ] | {note_value}")
                        mark_case_done(row['order_number'], row['sku'], pharmacy_name, case_type, pharmacist_name)
                        st.toast("🔄 تم العكس!"); st.rerun()
            elif case_type in {"addition", "orphan_salla", "return", "orphan_abc"}:
                button_label = "✅ تأكيد الإضافة" if case_type in {"addition", "orphan_salla"} else "🔄 تأكيد الإرجاع"
                with btn_col2:
                    if st.button(button_label, key=f"done_{note_key}", use_container_width=True):
                        save_case_note(row['order_number'], row['sku'], pharmacy_name, case_type, note_value)
                        mark_case_done(row['order_number'], row['sku'], pharmacy_name, case_type, pharmacist_name)
                        st.toast("🚀 تم التأكيد!"); st.rerun()
        st.markdown("---")

def render_case_cards_pharmacy(df, allow_actions, pharmacist_name, pharmacy_name, tab_id=""):
    """رسم وإدارة بطاقات العرض للتبويبات العادية مع معالجة تسلسل الـ index الفريد لكل لسان"""
    if df is not None and not df.empty:
        for i, (orig_idx, row) in enumerate(df.iterrows()):
            render_single_case_card(row, i, allow_actions, pharmacist_name, pharmacy_name, tab_id=tab_id)
    else:
        st.success("🎉 ممتاز! التبويب الحالي متطابق بالكامل ولا توجد أي فروقات مخزنية معلقة هنا.")

def show():
    pharmacy_name = st.session_state.username
    pharmacist_name = st.session_state.get("pharmacist_name", "") or ""
    branch_number = get_branch_number(pharmacy_name)
    branch_location = get_branch_location(branch_number)

    st.markdown(f"""
    <div class="hero">
        <h1>🏥 {pharmacy_name}</h1>
        <p>فرع رقم {branch_number} | الموقع: {branch_location} | الصيدلي: {pharmacist_name}</p>
        <p>🕐 آخر تحديث: {get_saudi_time()}</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1.2, 1.5])
    with col1:
        if st.button("🔄 تحديث الصفحة", use_container_width=True): st.rerun()
    with col2:
        if st.button("📥 تصدير الملف الكامل Excel", use_container_width=True): st.session_state.show_export_pharmacy = True
    with col3:
        if st.button("📋 تصدير ملف مختصر Excel", use_container_width=True): st.session_state.show_export_brief_pharmacy = True

    df = fetch_active_items(pharmacy_name, include_hidden=False)
    old_invoices_df = get_old_invoices(pharmacy_name=pharmacy_name, months=6)
    old_orders_df = get_old_orders(pharmacy_name=pharmacy_name, months=6)

    if df.empty:
        st.info("📭 لا توجد حالات نشطة لهذا الفرع حاليًا.")
        return

    is_locked = df['is_item_locked'].iloc[0] == 1 if 'is_item_locked' in df.columns else False
    allow_actions = not is_locked

    active_mask = ~df["order_status"].apply(is_cancelled_or_returned_status)
    active_df = df[active_mask].copy()

    branch_add_df = df[df['case_type'].isin(['addition', 'orphan_salla']) & active_mask].copy()
    total_additions_merged = len(branch_add_df)
    completed_additions_merged = len(branch_add_df[branch_add_df["status"] == "تم"])
    
    branch_ret_df = df[df['case_type'].isin(['return', 'orphan_abc']) & active_mask].copy()
    total_returns_merged = len(branch_ret_df)
    completed_returns_merged = len(branch_ret_df[branch_ret_df["status"] == "تم"])
    
    conflicts_df = df[df["case_type"] == "branch_conflict"].copy()
    total_conflicts = len(conflicts_df)
    completed_conflicts = len(conflicts_df[conflicts_df["status"] == "تم"])

    post_cutoff_df = df[(df["case_type"] == "post_cutoff_abc") & active_mask].copy()
    total_post_cutoff = len(post_cutoff_df)
    completed_post_cutoff = len(post_cutoff_df[post_cutoff_df["status"] == "تم"])
    
    payment_df = df[df["order_status"].apply(is_pending_payment_status) & (df["case_type"] != "branch_conflict") & active_mask].copy()
    total_payment = len(payment_df)
    
    cancelled_df = df[df["order_status"].apply(is_cancelled_or_returned_status) & (df["case_type"] != "branch_conflict")].copy()
    total_cancelled = len(cancelled_df)
    
    completed_df = get_completed_items(pharmacy_name)
    total_completed = len(completed_df)

    if st.session_state.get('show_export_pharmacy', False):
        excel_sheets = {
            "الاضافات والطلبات المفقودة": branch_add_df,
            "الارجاعات والزيادات": branch_ret_df,
            "فواتير معلقة بين الفروع": conflicts_df,
            "فواتير بعد اخر طلب": post_cutoff_df,
            "بانتظار الدفع": payment_df,
            "الملغيات والمسترجعات": cancelled_df,
            "تم الانتهاء": completed_df
        }
        excel_data = export_to_excel(excel_sheets, pharmacy_name)
        st.download_button(label="💾 تحميل ملف Excel الموحد", data=excel_data, file_name=f"Full_Report_{pharmacy_name}.xlsx", use_container_width=True)
        st.session_state.show_export_pharmacy = False

    if st.session_state.get('show_export_brief_pharmacy', False):
        excel_sheets_brief = {
            "الاضافات والطلبات المفقودة": branch_add_df,
            "الارجاعات والزيادات": branch_ret_df,
            "فواتير معلقة بين الفروع": conflicts_df
        }
        excel_data_brief = export_to_excel_brief(excel_sheets_brief)
        st.download_button(label="📊 تحميل ملف Excel المختصر", data=excel_data_brief, file_name=f"Brief_Report_{pharmacy_name}.xlsx", use_container_width=True)
        st.session_state.show_export_brief_pharmacy = False
        
    tab_additions, tab_returns, tab_conflicts, tab_post_cutoff, tab_payment, tab_cancelled = st.tabs([
        f"📥 الإضافات والطلبات ({completed_additions_merged}/{total_additions_merged})",
        f"📤 الإرجاعات والزيادات ({completed_returns_merged}/{total_returns_merged})",
        f"📊 فواتير معلقة بين الفروع ({completed_conflicts}/{total_conflicts})", 
        f"⏰ فواتير بعد آخر طلب ({completed_post_cutoff}/{total_post_cutoff})",
        f"💰 بانتظار الدفع ({total_payment})",
        f"⚠️ ملغي/مسترجع ({total_cancelled})",
        f"✅ تم الانتهاء ({total_completed})"
    ])
    
    with tab_additions: render_case_cards_pharmacy(branch_add_df, allow_actions, pharmacist_name, pharmacy_name, tab_id="add")
    with tab_returns: render_case_cards_pharmacy(branch_ret_df, allow_actions, pharmacist_name, pharmacy_name, tab_id="ret")
    with tab_conflicts: render_case_cards_pharmacy(conflicts_df, allow_actions, pharmacist_name, pharmacy_name, tab_id="conf")
    with tab_post_cutoff: render_case_cards_pharmacy(post_cutoff_df, False, pharmacist_name, pharmacy_name, tab_id="cutoff")
    with tab_payment: render_case_cards_pharmacy(payment_df, False, pharmacist_name, pharmacy_name, tab_id="pay")
    with tab_cancelled: render_case_cards_pharmacy(cancelled_df, False, pharmacist_name, pharmacy_name, tab_id="cancel")
    with tab_completed: render_completed_table(completed_df, False, pharmacist_name, pharmacy_name, tab_id="complete")
