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
            top=Side(style='thin', color='D3D3D3'),
            bottom=Side(style='thin', color='D3D3D3')
        )
        
        # الأسماء المعبرة عن أعمدة الأسعار لتطبيق الخط العريض (Bold) عليها
        price_columns_keywords = ["سعر", "السعر", "الخصم", "قيمة"]
        
        # 1. تنسيق صف العناوين الرئيسي
        for col_idx in range(1, len(df.columns) + 1):
            cell = worksheet.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
            cell.border = thin_border
            
        # 2. تنسيق صفوف البيانات والأسعار والحدود
        for row_idx in range(2, len(df) + 2):
            for col_idx in range(1, len(df.columns) + 1):
                cell = worksheet.cell(row=row_idx, column=col_idx)
                col_name = df.columns[col_idx - 1]
                
                # تطبيق الخط العريض إذا كان العمود يخص الأسعار أو القيم المالية
                if any(keyword in col_name for keyword in price_columns_keywords):
                    cell.font = bold_price_font
                else:
                    cell.font = normal_font
                    
                cell.border = thin_border
                cell.alignment = center_align
                
        # 3. احتساب وضبط عرض الأعمدة تلقائياً بناء على المحتوى لمنع ظهور رموز الاختصار ###
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            worksheet.column_dimensions[col_letter].width = max(max_len + 4, 15)
            
    output_styled.seek(0)
    return output_styled.getvalue()

def calculate_promotions_stream(uploaded_file):
    """حساب وتفصيل العروض الترويجية بشكل آمن مع دمج أعمدة المنتج والـ SKU"""
    df = pd.read_excel(uploaded_file)
    
    # التحقق من وجود الأعمدة المحدثة في الملف المرفوع
    required_cols = ["رقم المنتج (SKU)", "اسم المنتج", "عنوان العرض", "السعر غير شامل الضريبة", "الضريبة"]
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"العمود الأساسي المطلوب '{col}' مفقود! يرجى تحميل النموذج الجديد وتعبئته.")

    df["عنوان العرض"] = df["عنوان العرض"].astype(str).str.strip()

    total_qty_list, discount_pct_list, price_before_list, discount_val_list, price_after_list = [], [], [], [], []

    for idx, row in df.iterrows():
        p_base = float(row.get("السعر غير شامل الضريبة", 0))
        tax_pct = float(row.get("الضريبة", 0))
        promo_text = row["عنوان العرض"]
        
        p_tax = p_base * (1 + tax_pct)
        qty, disc_pct, p_before, p_after = 1, 0.0, p_tax, p_tax

        match_free = re.search(r"عرض\s*(\d+)\s*\+\s*(\d+)\s*مجانا", promo_text)
        match_free_alt = re.search(r"عرض\s*(\d+)\s*حبة\s*\+\s*(\d+)\s*مجانا", promo_text)

        if match_free or match_free_alt:
            m = match_free if match_free else match_free_alt
            buy_qty, free_qty = int(m.group(1)), int(m.group(2))
            qty = buy_qty + free_qty
            p_before = p_tax * qty
            p_after = p_tax * buy_qty
            disc_pct = (free_qty / qty) if qty > 0 else 0.0
            
        elif "على الحبة الثانية" in promo_text and "%" in promo_text:
            try:
                pct_val = float(re.search(r"خصم\s*(\d+)%", promo_text).group(1)) / 100
                qty = 2
                p_before = p_tax * 2
                p_after = p_tax + (p_tax * (1 - pct_val))
                disc_pct = ((p_before - p_after) / p_before) if p_before > 0 else 0.0
            except: pass
            
        elif "خصم" in promo_text and "%" in promo_text:
            try:
                pct_val = float(re.search(r"خصم\s*(\d+)%", promo_text).group(1)) / 100
                qty = 1
                p_before = p_tax
                p_after = p_tax * (1 - pct_val)
                disc_pct = pct_val
            except: pass
            
        elif "بسعر" in promo_text:
            qty_match = re.search(r"(\d+)\s*حبة", promo_text)
            price_match = re.search(r"بسعر\s*([\d\.]+)", promo_text)
            if qty_match and price_match:
                qty = int(qty_match.group(1))
                p_after = float(price_match.group(1))
                p_before = p_tax * qty
                disc_pct = ((p_before - p_after) / p_before) if p_before > 0 else 0.0

        disc_val = p_before - p_after
        
        total_qty_list.append(qty)
        discount_pct_list.append(round(disc_pct * 100, 2))
        price_before_list.append(round(p_before, 2))
        discount_val_list.append(round(disc_val, 2))
        price_after_list.append(round(p_after, 2))

    # دمج وترتيب الأعمدة المحسوبة داخل جدول البيانات الرئيسي
    df["عدد حبات العرض"] = total_qty_list
    df["نسبة الخصم الإجمالية %"] = discount_pct_list
    df["السعر قبل العرض شامل الضريبة"] = price_before_list
    df["قيمة الخصم شامل الضريبة"] = discount_val_list
    df["السعر بعد العرض شامل الضريبة"] = price_after_list

    # إرسال الملف المكتمل لتجميله وتنسيقه باحترافية كاملة
    styled_excel_bytes = style_excel_professionally(df, None)
    return styled_excel_bytes, df

def show():
    st.markdown("<div class='hero'><h1>🏷️ حاسبة ومُنسّق العروض الترويجية المطور</h1><p>تصدير تقارير محاسبية منسقة بالكامل للأعضاء والإدارة مع تمييز دقيق لبيانات أسماء المنتجات والأسعار العريضة</p></div>", unsafe_allow_html=True)
    
    st.markdown("### 📋 الخطوة 1: تنزيل نموذج الهيكل الجديد")
    st.info("قم بتحميل الملف المفرغ أدناه والمطور، والذي يضم الآن أعمدة (رقم المنتج SKU واسم المنتج):")
    
    template_bytes = create_template_excel()
    st.download_button(
        label="📥 تحميل نموذج العروض المطور والمعدّل (Excel Template)",
        data=template_bytes,
        file_name="Promotions_Enhanced_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    
    st.markdown("---")
    
    st.markdown("### 📂 الخطوة 2: رفع وتحليل ملف العروض")
    uploaded_file = st.file_uploader("ارفع ملف العروض المكتمل بعد تعبئته بناءً على هيكل النموذج الجديد", type=["xlsx"])
    
    if uploaded_file is not None:
        with st.spinner("⏳ جاري تفكيك العروض وتطبيق التنسيق الجمالي على جداول الأسعار العريضة..."):
            try:
                excel_bytes, result_df = calculate_promotions_stream(uploaded_file)
                st.success("✨ تم الحساب وإعادة تنسيق الملف بجدول احترافي بالكامل وجعل عمود السعر عريضاً (Bold)!")
                
                st.markdown("### 📊 معاينة تفاعلية سريعة للملف الجاهز للتنزيل:")
                st.dataframe(result_df.head(10), use_container_width=True)
                
                st.download_button(
                    label="💾 تحميل الملف الاحترافي المنسق والجاهز للإدارة (Excel)",
                    data=excel_bytes,
                    file_name="Styled_Promotions_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
            except ValueError as ve:
                st.error(f"⚠️ {str(ve)}")
            except Exception as e:
                st.error(f"❌ حدث خطأ غير متوقع أثناء التصميم والحساب: {str(e)}")
