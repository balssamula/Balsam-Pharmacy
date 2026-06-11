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
        else: badge_text,
