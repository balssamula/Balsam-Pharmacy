import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_template_excel():
    """إنشاء ملف نموذج فارغ مفرغ يحتوي على الأعمدة المطلوبة بالتنسيق الجديد المعزز"""
    template_data = {
        "رقم المنتج (SKU)": ["1001", "1002", "1003", "1004"],
        "اسم المنتج": ["بندول نايت 20 قرص", "فيتامين سي 1000 ملجم", "معجون أسنان كولجيت", "حليب أطفال نيدو 400 جرام"],
        "عنوان العرض": [
            "عرض 1+1 مجانا", 
            "خصم 50% على الحبة الثانية", 
            "خصم 20%", 
            "2 حبة بسعر 99.95 ريال"
        ],
        "السعر غير شامل الضريبة": [100.00, 50.00, 150.00, 60.00],
        "الضريبة": [0.15, 0.15, 0.15, 0.15]
    }
    df_template = pd.DataFrame(template_data)
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_template.to_excel(writer, index=False, sheet_name="نموذج العروض")
        
        # تنسيق السيرفر والنموذج بشكل سريع
        worksheet = writer.sheets["نموذج العروض"]
        header_fill = PatternFill(start_color="1F7A8C", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, name="Segoe UI")
        for col in range(1, len(df_template.columns) + 1):
            cell = worksheet.cell(row=1, column=col)
            cell.fill = header_fill
            cell.font = header_font
            worksheet.column_dimensions[get_column_letter(col)].width = 25
            
    output.seek(0)
    return output.getvalue()

def style_excel_professionally(df, output_bytesio):
    """إعادة فتح ملف الإكسيل وتطبيق تنسيق تجميلي واحترافي مكثف بألوان مخصصة وعمود سعر عريض"""
    output_styled = BytesIO()
    
    # تحميل الملف المكتوب لتعديل ستايل الخلايا
    with pd.ExcelWriter(output_styled, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="تقرير العروض الذكي")
        worksheet = writer.sheets["تقرير العروض الذكي"]
        
        # إعدادات الألوان والخطوط والحدود
        header_fill = PatternFill(start_color="1F7A8C", end_color="1F7A8C", fill_type="solid") # لون كحلي مميز للعنوان
        header_font = Font(color="FFFFFF", bold=True, size=11, name="Segoe UI")
        
        bold_price_font = Font(bold=True, size=11, name="Segoe UI", color="0F4C5C") # خط عريض مخصص لأعمدة الأسعار
        normal_font = Font(size=11, name="Segoe UI")
        
        center_align = Alignment(horizontal="center", vertical="center")
        right_align = Alignment(horizontal="right", vertical="center")
        
        thin_border = Border(
            left=Side(style='thin', color='D3D3D3'),
            right=Side(style='thin', color='D3D3D3'),
            top=Side(
