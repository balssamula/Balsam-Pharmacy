# pages/promotion_viewer.py
import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="منظومة العروض", layout="wide")

def apply_excel_style(writer, sheet_name, df):
    """تطبيق التنسيقات على ملف Excel"""
    workbook = writer.book
    worksheet = workbook[sheet_name]
    
    header_fill = PatternFill(start_color="1F7A8C", end_color="1F7A8C", fill_type="solid")
    header_font = Font(name="Tajawal", size=12, bold=True, color="FFFFFF")
    alt_row_fill = PatternFill(start_color="E6F3F5", end_color="E6F3F5", fill_type="solid")
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    for col_idx, col_name in enumerate(df.columns, 1):
        cell = worksheet.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
        worksheet.column_dimensions[get_column_letter(col_idx)].width = max(20, len(str(col_name)) + 5)
    
    for row_idx in range(2, len(df) + 2):
        for col_idx in range(1, len(df.columns) + 1):
            cell = worksheet.cell(row=row_idx, column=col_idx)
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if (row_idx - 2) % 2 == 1:
                cell.fill = alt_row_fill
    
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
        <p>جميع العروض الترويجية والخصومات المطبقة على المنتجات مع دعم المجموعات</p>
    </div>
    """, unsafe_allow_html=True)

    st.subheader("📂 رفع ملف Excel")
    
    uploaded_file = st.file_uploader(
        "قم برفع ملف Excel",
        type=["xlsx", "xls"],
        help="يجب أن يحتوي الملف على أوراق: 'عرض خاص'، 'سعر مخفض'، 'اسعار المنتجات'"
    )
    
    if uploaded_file is not None:
        with st.spinner("جاري معالجة العروض..."):
            excel_data = pd.ExcelFile(uploaded_file)
            sheet_names = excel_data.sheet_names
            
            # تحديد الأوراق المطلوبة
            offers_sheet = None
            discounted_sheet = None
            prices_sheet = None
            
            for sheet in sheet_names:
                if "عرض خاص" in sheet or "عروض" in sheet:
                    offers_sheet = sheet
                if "سعر مخفض" in sheet or "مخفض" in sheet:
                    discounted_sheet = sheet
                if "اسعار المنتجات" in sheet or "سعر المنتج" in sheet:
                    prices_sheet = sheet
            
            if not offers_sheet:
                offers_sheet = sheet_names[0]
            
            # قراءة البيانات
            df_raw = pd.read_excel(uploaded_file, sheet_name=offers_sheet, header=None)
            df_discounted = pd.read_excel(uploaded_file, sheet_name=discounted_sheet) if discounted_sheet else pd.DataFrame()
            df_regular_prices = pd.read_excel(uploaded_file, sheet_name=prices_sheet) if prices_sheet else pd.DataFrame()
            
            # ========== 1. معالجة الأسعار العادية ==========
            price_map = {}  # {sku: {name, price}}
            if not df_regular_prices.empty:
                sku_col = None
                price_col = None
                name_col = None
                for col in df_regular_prices.columns:
                    col_str = str(col).lower()
                    if "رمز" in col_str or "sku" in col_str or "كود" in col_str or "رقم" in col_str:
                        sku_col = col
                    if "سعر" in col_str or "price" in col_str:
                        price_col = col
                    if "اسم" in col_str or "name" in col_str:
                        name_col = col
                
                if sku_col:
                    for _, row in df_regular_prices.iterrows():
                        sku = str(row[sku_col]).strip()
                        if sku and sku != "nan":
                            price_map[sku] = {
                                "name": row[name_col] if name_col else "",
                                "price": row[price_col] if price_col else ""
                            }
            
            # ========== 2. معالجة الأسعار المخفضة ==========
            discounted_map = {}  # {sku: {discounted_price, end_date, promo_title}}
            if not df_discounted.empty:
                sku_col = None
                price_col = None
                end_date_col = None
                promo_col = None
                for col in df_discounted.columns:
                    col_str = str(col).lower()
                    if "رمز" in col_str or "sku" in col_str or "كود" in col_str or "رقم" in col_str:
                        sku_col = col
                    if "سعر" in col_str or "price" in col_str or "مخفض" in col_str:
                        price_col = col
                    if "نهاية" in col_str or "تاريخ نهاية" in col_str:
                        end_date_col = col
                    if "عنوان" in col_str or "ترويجي" in col_str:
                        promo_col = col
                
                if sku_col:
                    for _, row in df_discounted.iterrows():
                        sku = str(row[sku_col]).strip()
                        if sku and sku != "nan":
                            discounted_map[sku] = {
                                "discounted_price": row[price_col] if price_col else "",
                                "end_date": row[end_date_col] if end_date_col else "",
                                "promo_title": row[promo_col] if promo_col else ""
                            }
            
            # ========== 3. معالجة العروض ==========
            # هيكل التخزين: {sku: {offers: [], group_sku: None, group_qty: None, products_in_group: []}}
            products_data = {}
            
            for idx, row in df_raw.iterrows():
                text = row.iloc[0] if len(row) > 0 else ""
                if pd.isna(text) or text == "nan":
                    continue
                
                text = str(text)
                
                # استخراج المجموعات (مثل 1500*6)
                group_pattern = r'(\d{4,6})\*(\d+)'
                groups = re.findall(group_pattern, text)
                
                # استخراج الأرقام المتعددة المفصولة بـ - (مثل 15000-16000-12000)
                dash_pattern = r'(\d{4,6})-(\d{4,6})-(\d{4,6})'
                dash_matches = re.findall(dash_pattern, text)
                
                # استخراج الأرقام المفردة (المنتجات العادية)
                excluded_years = {'2024', '2025', '2026', '2027', '2028', '2029', '2030'}
                single_pattern = r'(?:^|[-/\s]+)(\d{4,6})(?:[-/\s]|$)'
                singles = re.findall(single_pattern, text)
                singles = [s for s in singles if s not in excluded_years and not s.startswith('20')]
                
                # استخراج اسم العرض والتواريخ
                offer_name = extract_offer_name(text)
                start_date, end_date = extract_dates(text)
                is_daily_deal = "صفقة اليوم" in text
                
                # ===== حالة 1: منتج مركب به * (مثل 1500*6) =====
                for base_sku, qty in groups:
                    group_sku = f"{base_sku}*{qty}"
                    if base_sku not in products_data:
                        products_data[base_sku] = {
                            "offers": [],
                            "group_sku": group_sku,
                            "group_qty": int(qty),
                            "products_in_group": [base_sku]
                        }
                    # إضافة العرض الحالي
                    products_data[base_sku]["offers"].append({
                        "name": offer_name,
                        "start": start_date,
                        "end": end_date,
                        "is_daily_deal": is_daily_deal,
                        "group_sku": group_sku,
                        "group_qty": int(qty)
                    })
                
                # ===== حالة 2: منتجات متعددة في عرض واحد (مثل 15000-16000-12000) =====
                for multi in dash_matches:
                    group_key = f"{multi[0]}-{multi[1]}-{multi[2]}"
                    for sku in multi:
                        if sku not in products_data:
                            products_data[sku] = {
                                "offers": [],
                                "group_sku": group_key,
                                "group_qty": None,
                                "products_in_group": list(multi)
                            }
                        else:
                            # تحديث معلومات المجموعة إذا لم تكن موجودة
                            if not products_data[sku].get("group_sku"):
                                products_data[sku]["group_sku"] = group_key
                                products_data[sku]["products_in_group"] = list(multi)
                        
                        products_data[sku]["offers"].append({
                            "name": offer_name,
                            "start": start_date,
                            "end": end_date,
                            "is_daily_deal": is_daily_deal,
                            "group_sku": group_key,
                            "group_qty": None
                        })
                
                # ===== حالة 3: منتج عادي (رقم فردي) =====
                for sku in singles:
                    if sku not in products_data:
                        products_data[sku] = {
                            "offers": [],
                            "group_sku": None,
                            "group_qty": None,
                            "products_in_group": []
                        }
                    
                    products_data[sku]["offers"].append({
                        "name": offer_name,
                        "start": start_date,
                        "end": end_date,
                        "is_daily_deal": is_daily_deal,
                        "group_sku": None,
                        "group_qty": None
                    })
            
            # ========== 4. بناء النتيجة النهائية (صف واحد لكل منتج) ==========
            final_results = []
            
            for sku, data in products_data.items():
                # جلب بيانات السعر العادي
                product_info = price_map.get(sku, {"name": "", "price": ""})
                product_name = product_info["name"]
                product_price = product_info["price"]
                
                # جلب بيانات السعر المخفض للمنتج العادي
                disc_info = discounted_map.get(sku, {})
                disc_price = disc_info.get("discounted_price", "")
                disc_promo = disc_info.get("promo_title", "")
                
                # جلب بيانات السعر المخفض للمجموعة (إذا وجدت)
                group_sku = data.get("group_sku")
                group_qty = data.get("group_qty")
                group_disc_info = discounted_map.get(group_sku, {}) if group_sku else {}
                group_disc_price = group_disc_info.get("discounted_price", "")
                group_disc_promo = group_disc_info.get("promo_title", "")
                
                # دمج العروض المتعددة لنفس المنتج في سطر واحد
                offers_list = data["offers"]
                unique_offers = {}
                for offer in offers_list:
                    key = offer["name"]
                    if key not in unique_offers:
                        unique_offers[key] = offer
                
                offer_names = " | ".join([o["name"] for o in unique_offers.values()])
                offer_starts = " | ".join([o["start"] for o in unique_offers.values() if o["start"]])
                offer_ends = " | ".join([o["end"] for o in unique_offers.values() if o["end"]])
                is_daily = "نعم" if any(o["is_daily_deal"] for o in unique_offers.values()) else ""
                
                # ملاحظة: إذا كان المنتج جزء من مجموعة متعددة
                note = ""
                if data.get("products_in_group") and len(data["products_in_group"]) > 1:
                    note = f"عرض يشمل المنتجات: {', '.join(data['products_in_group'])}"
                
                final_results.append({
                    "رقم المنتج": sku,
                    "رقم المنتج للمجموعة": group_sku if group_sku else "",
                    "اسم المنتج": product_name,
                    "سعر المنتج": product_price,
                    "اسم العرض الخاص": offer_names,
                    "بداية العرض": offer_starts,
                    "نهاية العرض": offer_ends,
                    "سعر مخفض للمنتج": disc_price,
                    "سعر مخفض للمجموعة": group_disc_price,
                    "عدد حبات المجموعة": group_qty if group_qty else "",
                    "العنوان الترويجي للمنتج": disc_promo,
                    "العنوان الترويجي للمجموعة": group_disc_promo,
                    "صفقة اليوم": is_daily,
                    "ملاحظة": note
                })
            
            # إضافة المنتجات التي لها سعر مخفض فقط (بدون عرض)
            for sku, disc_info in discounted_map.items():
                if sku not in products_data:
                    product_info = price_map.get(sku, {"name": "", "price": ""})
                    final_results.append({
                        "رقم المنتج": sku,
                        "رقم المنتج للمجموعة": "",
                        "اسم المنتج": product_info["name"],
                        "سعر المنتج": product_info["price"],
                        "اسم العرض الخاص": "",
                        "بداية العرض": "",
                        "نهاية العرض": "",
                        "سعر مخفض للمنتج": disc_info.get("discounted_price", ""),
                        "سعر مخفض للمجموعة": "",
                        "عدد حبات المجموعة": "",
                        "العنوان الترويجي للمنتج": disc_info.get("promo_title", ""),
                        "العنوان الترويجي للمجموعة": "",
                        "صفقة اليوم": "",
                        "ملاحظة": "منتج مخفض بدون عرض"
                    })
            
            df_final = pd.DataFrame(final_results)
            
            # تنظيف البيانات
            df_final = df_final[df_final["رقم المنتج"].notna()]
            df_final = df_final[df_final["رقم المنتج"] != "nan"]
            df_final = df_final[df_final["رقم المنتج"] != ""]
            df_final = df_final[~df_final["رقم المنتج"].isin(["2024", "2025", "2026", "2027", "2028", "2029", "2030"])]
            
            # ترتيب الأعمدة
            column_order = [
                "رقم المنتج", "رقم المنتج للمجموعة", "اسم المنتج", "سعر المنتج",
                "اسم العرض الخاص", "بداية العرض", "نهاية العرض",
                "سعر مخفض للمنتج", "سعر مخفض للمجموعة", "عدد حبات المجموعة",
                "العنوان الترويجي للمنتج", "العنوان الترويجي للمجموعة",
                "صفقة اليوم", "ملاحظة"
            ]
            df_final = df_final[[col for col in column_order if col in df_final.columns]]
            
            # عرض الإحصائيات
            st.success(f"✅ تم معالجة {len(df_final)} منتج بنجاح")
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("📦 إجمالي المنتجات", len(df_final))
            with col2:
                st.metric("🏷️ مع عروض", df_final["اسم العرض الخاص"].str.len().gt(0).sum())
            with col3:
                st.metric("🏷️ منتجات مخفضة", df_final["سعر مخفض للمنتج"].notna().sum())
            with col4:
                st.metric("🎯 مجموعات", df_final["رقم المنتج للمجموعة"].str.len().gt(0).sum())
            
            # عرض الجدول
            st.subheader("📋 قائمة المنتجات والعروض")
            st.dataframe(df_final, use_container_width=True, height=400)
            
            # ========== خيارات التحميل ==========
            st.subheader("💾 تحميل النتائج")
            
            col1, col2 = st.columns(2)
            
            with col1:
                def create_detailed_excel(df):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, sheet_name="النتيجة النهائية", index=False)
                        apply_excel_style(writer, "النتيجة النهائية", df)
                        
                        if "اسم العرض الخاص" in df:
                            for offer_type in df["اسم العرض الخاص"].dropna().unique():
                                if offer_type and offer_type != "":
                                    type_df = df[df["اسم العرض الخاص"] == offer_type]
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
            
            with col2:
                def create_simple_excel(df):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, sheet_name="العروض والمنتجات", index=False)
                        apply_excel_style(writer, "العروض والمنتجات", df)
                    return output.getvalue()
                
                simple_excel = create_simple_excel(df_final)
                st.download_button(
                    label="📥 تحميل ملف مبسط (جدول واحد)",
                    data=simple_excel,
                    file_name="العروض_والمنتجات.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            # فلترة حسب نوع العرض
            if "اسم العرض الخاص" in df_final:
                st.subheader("🔍 تصفية حسب نوع العرض")
                offer_types = ["الكل"] + sorted(df_final["اسم العرض الخاص"].dropna().unique().tolist())
                selected_type = st.selectbox("اختر نوع العرض", offer_types)
                
                if selected_type != "الكل":
                    filtered_df = df_final[df_final["اسم العرض الخاص"] == selected_type]
                else:
                    filtered_df = df_final
                
                st.dataframe(filtered_df, use_container_width=True)
            
            # عرض المنتجات ذات العروض المتعددة
            duplicates = df_final[df_final["رقم المنتج"].duplicated(keep=False)]
            if len(duplicates) > 0:
                with st.expander(f"⚠️ منتجات لها أكثر من عرض ({len(duplicates)} منتج)"):
                    st.dataframe(duplicates.sort_values("رقم المنتج"), use_container_width=True)
            
    else:
        st.info("📂 يرجى رفع ملف Excel لعرض العروض")
        
        with st.expander("ℹ️ تعليمات"):
            st.markdown("""
            **كيفية استخدام هذه الصفحة:**
            
            1. قم برفع ملف Excel الذي يحتوي على:
               - ورقة "عرض خاص" (العروض الترويجية)
               - ورقة "سعر مخفض" (المنتجات المخفضة)
               - ورقة "اسعار المنتجات" (الأسعار العادية للمنتجات)
            
            2. سيتم معالجة:
               - **المنتجات العادية** (مثل: 9974)
               - **المجموعات** (مثل: 1500*6)
               - **المنتجات المتعددة في عرض واحد** (مثل: 15000-16000-12000)
            
            3. **دمج العروض:** إذا كان نفس المنتج لديه أكثر من عرض، سيتم دمجها في صف واحد
            
            **الأعمدة الناتجة:**
            - رقم المنتج
            - رقم المنتج للمجموعة
            - اسم المنتج (من شيت اسعار المنتجات)
            - سعر المنتج
            - اسم العرض الخاص
            - بداية العرض
            - نهاية العرض
            - سعر مخفض للمنتج
            - سعر مخفض للمجموعة
            - عدد حبات المجموعة
            - العنوان الترويجي للمنتج
            - العنوان الترويجي للمجموعة
            - صفقة اليوم
            - ملاحظة
            """)

# ========== دوال مساعدة ==========
def extract_offer_name(text):
    text = str(text) if not pd.isna(text) else ""
    if "خصم" in text and "القطعة الثانية" in text:
        match = re.search(r'(خصم \d+% على القطعة الثانية)', text)
        return match.group(1) if match else "خصم على القطعة الثانية"
    if "خصم" in text and "الحبة الثانية" in text:
        match = re.search(r'(خصم \d+% على الحبة الثانية)', text)
        if match: return match.group(1)
        match = re.search(r'(خصم \d+ ريال على الحبة الثانية)', text)
        return match.group(1) if match else "خصم على الحبة الثانية"
    if "عرض" in text and "مجاناً" in text:
        match = re.search(r'(عرض \d+\+\d+ مجاناً?)', text)
        return match.group(1) if match else "عرض مجاني"
    if "صفقة اليوم" in text:
        return "صفقة اليوم"
    if "حبة بسعر" in text or "حبات بسعر" in text:
        match = re.search(r'(\d+حبات? بسعر [\d.]+ ريال)', text)
        return match.group(1) if match else "عرض كميات"
    return "عرض خاص"

def extract_dates(text):
    text = str(text) if not pd.isna(text) else ""
    date_match = re.findall(r'(\d{4}-\d{2}-\d{2} \d{2}:\d{2})', text)
    start = date_match[0] if len(date_match) > 0 else None
    end = date_match[1] if len(date_match) > 1 else None
    return start, end
