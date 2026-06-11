# pages/promotion_viewer.py
import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def apply_excel_style(writer, sheet_name, df):
    """تطبيق التنسيقات على ملف Excel"""
    workbook = writer.book
    worksheet = workbook[sheet_name]
    
    # ألوان التنسيق
    header_fill = PatternFill(start_color="1F7A8C", end_color="1F7A8C", fill_type="solid")
    header_font = Font(name="Tajawal", size=12, bold=True, color="FFFFFF")
    alt_row_fill = PatternFill(start_color="E6F3F5", end_color="E6F3F5", fill_type="solid")
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # تنسيق رأس الجدول
    for col_idx, col_name in enumerate(df.columns, 1):
        cell = worksheet.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
        # ضبط عرض العمود
        worksheet.column_dimensions[get_column_letter(col_idx)].width = max(20, len(str(col_name)) + 5)
    
    # تنسيق الصفوف
    for row_idx in range(2, len(df) + 2):
        for col_idx in range(1, len(df.columns) + 1):
            cell = worksheet.cell(row=row_idx, column=col_idx)
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
            # تلوين الصفوف الزوجية
            if (row_idx - 2) % 2 == 1:
                cell.fill = alt_row_fill
    
    # تثبيت الصف الأول
    worksheet.freeze_panes = 'A2'

def show():
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@300;400;500;700;800&display=swap');
        * { font-family: 'Tajawal', sans-serif; }
        .offers-header {
            background: linear-gradient(135deg, #0f4c5c 0%, #1f7a8c 50%, #16425b 100%);
            border-radius: 24px;
            padding: 1.5rem;
            color: white;
            margin-bottom: 1.5rem;
            text-align: center;
        }
        .offer-card {
            background: white;
            border-radius: 16px;
            padding: 1rem;
            margin-bottom: 1rem;
            border-right: 4px solid #1f7a8c;
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        }
        .price-cell {
            background: #dff7e8;
            color: #0f7a3a;
            font-weight: bold;
            padding: 0.2rem 0.5rem;
            border-radius: 8px;
            display: inline-block;
        }
        .download-section {
            background: #f0f7f9;
            border-radius: 16px;
            padding: 1rem;
            margin-top: 1rem;
        }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="offers-header">
        <h1>🛍️ العروض الحالية الفعالة بالمتجر</h1>
        <p>جميع العروض الترويجية والخصومات المطبقة على المنتجات</p>
    </div>
    """, unsafe_allow_html=True)

    # رفع ملف العروض
    st.subheader("📂 رفع ملف Excel")
    
    uploaded_file = st.file_uploader(
        "قم برفع ملف Excel الذي يحتوي على العروض والأسعار المخفضة",
        type=["xlsx", "xls"],
        help="الملف يجب أن يحتوي على ورقة 'عرض خاص' وورقة 'سعر مخفض'"
    )
    
    if uploaded_file is not None:
        with st.spinner("جاري معالجة العروض..."):
            # قراءة الملف
            excel_data = pd.ExcelFile(uploaded_file)
            sheet_names = excel_data.sheet_names
            
            # البحث عن ورقة العروض
            offers_sheet = None
            discounted_sheet = None
            
            for sheet in sheet_names:
                if "عرض خاص" in sheet or "عروض" in sheet:
                    offers_sheet = sheet
                if "سعر مخفض" in sheet or "مخفض" in sheet:
                    discounted_sheet = sheet
            
            if not offers_sheet:
                offers_sheet = sheet_names[0]
            if not discounted_sheet and len(sheet_names) > 1:
                discounted_sheet = sheet_names[1]
            
            # قراءة البيانات
            df_raw = pd.read_excel(uploaded_file, sheet_name=offers_sheet)
            df_discounted = pd.read_excel(uploaded_file, sheet_name=discounted_sheet) if discounted_sheet else pd.DataFrame()
            
            # دالة لاستخراج أرقام المنتجات فقط (بدون سنوات التواريخ)
            def extract_product_numbers(text):
                if pd.isna(text):
                    return []
                text = str(text)
                excluded_years = {'2024', '2025', '2026', '2027', '2028', '2029', '2030'}
                pattern = r'(?:^|[-/\s]+)(\d{4,6})(?:[-/\s]|$)'
                matches = re.findall(pattern, text)
                numbers = [m for m in matches if m not in excluded_years and not m.startswith('20')]
                return numbers
            
            # دالة لاستخراج اسم العرض
            def extract_offer_name(text):
                text = str(text) if not pd.isna(text) else ""
                
                if "خصم" in text:
                    if "القطعة الثانية" in text:
                        match = re.search(r'(خصم \d+% على القطعة الثانية)', text)
                        return match.group(1) if match else "خصم على القطعة الثانية"
                    elif "الحبة الثانية" in text:
                        match = re.search(r'(خصم \d+% على الحبة الثانية)', text)
                        return match.group(1) if match else "خصم على الحبة الثانية"
                    elif "ريال" in text:
                        match = re.search(r'(خصم \d+ ريال على الحبة الثانية)', text)
                        return match.group(1) if match else "خصم بقيمة محددة"
                elif "عرض" in text and "مجاناً" in text:
                    match = re.search(r'(عرض \d+\+\d+ مجاناً?)', text)
                    return match.group(1) if match else "عرض مجاني"
                elif "صفقة اليوم" in text:
                    return "صفقة اليوم"
                elif "حبة بسعر" in text or "حبات بسعر" in text:
                    match = re.search(r'(\d+حبات? بسعر [\d.]+ ريال)', text)
                    return match.group(1) if match else "عرض كميات"
                return "عرض خاص"
            
            # دالة لاستخراج التواريخ
            def extract_dates(text):
                text = str(text) if not pd.isna(text) else ""
                date_match = re.findall(r'(\d{4}-\d{2}-\d{2} \d{2}:\d{2})', text)
                start = date_match[0] if len(date_match) > 0 else None
                end = date_match[1] if len(date_match) > 1 else None
                return start, end
            
            # ========== 1. معالجة العروض ==========
            all_offers = []
            
            for idx, row in df_raw.iterrows():
                text = row.iloc[0] if len(row) > 0 else ""
                if pd.isna(text) or text == "nan":
                    continue
                
                numbers = extract_product_numbers(text)
                offer_name = extract_offer_name(text)
                start_date, end_date = extract_dates(text)
                
                for num in numbers:
                    all_offers.append({
                        "رقم الصنف": num,
                        "اسم العرض": offer_name,
                        "بداية العرض": start_date,
                        "نهاية العرض": end_date
                    })
            
            df_offers = pd.DataFrame(all_offers)
            
            # ========== 2. معالجة الأسعار المخفضة ==========
            df_prices = pd.DataFrame()
            
            if not df_discounted.empty:
                sku_col = None
                price_col = None
                end_date_col = None
                promo_col = None
                
                for col in df_discounted.columns:
                    col_str = str(col).lower()
                    if "رمز" in col_str or "sku" in col_str or "كود" in col_str or "رقم" in col_str or "رمز المنتج" in col_str:
                        sku_col = col
                    if "سعر" in col_str or "price" in col_str or "مخفض" in col_str:
                        price_col = col
                    if "نهاية" in col_str or "تاريخ نهاية" in col_str or "expiry" in col_str:
                        end_date_col = col
                    if "عنوان" in col_str or "ترويجي" in col_str or "promo" in col_str:
                        promo_col = col
                
                if sku_col is None and len(df_discounted.columns) > 0:
                    sku_col = df_discounted.columns[0]
                if price_col is None and len(df_discounted.columns) > 1:
                    price_col = df_discounted.columns[1]
                
                df_prices = pd.DataFrame()
                df_prices["رقم الصنف"] = df_discounted[sku_col].astype(str).str.strip()
                df_prices["السعر المخفض"] = df_discounted[price_col] if price_col else ""
                df_prices["تاريخ نهاية التخفيض"] = df_discounted[end_date_col] if end_date_col else ""
                df_prices["العنوان الترويجي"] = df_discounted[promo_col] if promo_col else ""
                
                df_prices = df_prices[df_prices["رقم الصنف"].notna()]
                df_prices = df_prices[df_prices["رقم الصنف"] != "nan"]
                df_prices = df_prices[df_prices["رقم الصنف"] != ""]
            
            # ========== 3. دمج البيانات ==========
            if not df_prices.empty:
                df_merged = df_offers.merge(df_prices, on="رقم الصنف", how="left")
            else:
                df_merged = df_offers.copy()
                df_merged["السعر المخفض"] = ""
                df_merged["تاريخ نهاية التخفيض"] = ""
                df_merged["العنوان الترويجي"] = ""
            
            # إضافة الأصناف التي لها سعر مخفض فقط
            if not df_prices.empty:
                discounted_only = df_prices[~df_prices["رقم الصنف"].isin(df_offers["رقم الصنف"])].copy()
                discounted_only["اسم العرض"] = ""
                discounted_only["بداية العرض"] = ""
                discounted_only["نهاية العرض"] = ""
                df_final = pd.concat([df_merged, discounted_only], ignore_index=True)
            else:
                df_final = df_merged
            
            # إعادة ترتيب الأعمدة
            final_columns = ["رقم الصنف", "اسم العرض", "بداية العرض", "نهاية العرض", 
                            "السعر المخفض", "تاريخ نهاية التخفيض", "العنوان الترويجي"]
            df_final = df_final[[col for col in final_columns if col in df_final.columns]]
            
            # تنظيف البيانات
            df_final = df_final[df_final["رقم الصنف"].notna()]
            df_final = df_final[df_final["رقم الصنف"] != "nan"]
            df_final = df_final[df_final["رقم الصنف"] != ""]
            df_final = df_final[~df_final["رقم الصنف"].isin(["2024", "2025", "2026", "2027", "2028", "2029", "2030"])]
            
            # عرض الإحصائيات
            st.success(f"✅ تم معالجة {len(df_final)} صنف بنجاح")
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("📦 إجمالي الأصناف", len(df_final))
            with col2:
                st.metric("🏷️ أنواع العروض", df_final["اسم العرض"].nunique() if "اسم العرض" in df_final else 0)
            with col3:
                st.metric("🏷️ منتجات مخفضة", df_final["السعر المخفض"].notna().sum() if "السعر المخفض" in df_final else 0)
            with col4:
                st.metric("🎯 عروض نشطة", len(df_final[df_final["اسم العرض"] != ""]) if "اسم العرض" in df_final else 0)
            
            # عرض جدول البيانات
            st.subheader("📋 قائمة العروض والأسعار المخفضة")
            st.dataframe(df_final, use_container_width=True, height=400)
            
            # ========== خيارات التحميل ==========
            st.subheader("💾 تحميل النتائج")
            st.markdown('<div class="download-section">', unsafe_allow_html=True)
            
            col1, col2 = st.columns(2)
            
            # ===== 1. ملف مفصل (مع شيتات منفصلة لكل عرض) =====
            with col1:
                st.markdown("### 📑 ملف مفصل")
                st.caption("يحتوي على شيت منفصل لكل نوع عرض + شيت رئيسي")
                
                def create_detailed_excel(df):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        # الشيت الرئيسي
                        df.to_excel(writer, sheet_name="النتيجة النهائية", index=False)
                        apply_excel_style(writer, "النتيجة النهائية", df)
                        
                        # شيت منفصل لكل نوع عرض
                        if "اسم العرض" in df:
                            for offer_type in df["اسم العرض"].dropna().unique():
                                if offer_type and offer_type != "":
                                    type_df = df[df["اسم العرض"] == offer_type]
                                    if len(type_df) > 0:
                                        sheet_name = offer_type[:31]
                                        type_df.to_excel(writer, sheet_name=sheet_name, index=False)
                                        apply_excel_style(writer, sheet_name, type_df)
                    return output.getvalue()
                
                detailed_excel = create_detailed_excel(df_final)
                st.download_button(
                    label="📥 تحميل ملف مفصل (مع شيتات)",
                    data=detailed_excel,
                    file_name="العروض_المفصلة_شيتات.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            # ===== 2. ملف مبسط (جدول واحد فقط مع تنسيق) =====
            with col2:
                st.markdown("### 📄 ملف مبسط")
                st.caption("جدول واحد فقط مع تنسيق احترافي للألوان")
                
                def create_simple_excel(df):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, sheet_name="العروض والأسعار المخفضة", index=False)
                        apply_excel_style(writer, "العروض والأسعار المخفضة", df)
                    return output.getvalue()
                
                simple_excel = create_simple_excel(df_final)
                st.download_button(
                    label="📥 تحميل ملف مبسط (جدول واحد)",
                    data=simple_excel,
                    file_name="العروض_والأسعار_المخفضة.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            st.markdown('</div>', unsafe_allow_html=True)
            
            # فلترة حسب نوع العرض
            if "اسم العرض" in df_final:
                st.subheader("🔍 تصفية حسب نوع العرض")
                offer_types = ["الكل"] + sorted(df_final["اسم العرض"].dropna().unique().tolist())
                selected_type = st.selectbox("اختر نوع العرض", offer_types)
                
                if selected_type != "الكل":
                    filtered_df = df_final[df_final["اسم العرض"] == selected_type]
                else:
                    filtered_df = df_final
                
                st.dataframe(filtered_df, use_container_width=True)
            
            # عرض الأصناف المكررة
            duplicates = df_final[df_final["رقم الصنف"].duplicated(keep=False)]
            if len(duplicates) > 0:
                with st.expander(f"⚠️ الأصناف المكررة ({len(duplicates)} صنف له أكثر من عرض)"):
                    st.dataframe(duplicates.sort_values("رقم الصنف"), use_container_width=True)
            
    else:
        st.info("📂 يرجى رفع ملف Excel لعرض العروض")
        
        with st.expander("ℹ️ تعليمات"):
            st.markdown("""
            **كيفية استخدام هذه الصفحة:**
            1. قم برفع ملف Excel الذي يحتوي على العروض والأسعار المخفضة
            2. يجب أن يحتوي الملف على:
               - ورقة باسم "عرض خاص" أو "العروض"
               - ورقة باسم "سعر مخفض" (تحتوي على: رمز المنتج، السعر المخفض، تاريخ نهاية التخفيض، العنوان الترويجي)
            3. سيتم فك العروض تلقائيًا (تقسيم الأرقام المتعددة)
            4. سيتم دمج العروض مع الأسعار المخفضة حسب رقم المنتج
            
            **خيارات التحميل:**
            - **ملف مفصل:** يحتوي على شيت منفصل لكل نوع عرض + شيت رئيسي
            - **ملف مبسط:** جدول واحد فقط مع تنسيق احترافي للألوان
            
            **ملاحظة:** تم تحسين الكود لاستبعاد السنوات (2024-2030) من عمود رقم الصنف.
            """)
