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
            
            # ========== عرض اختيار الأعمدة ==========
            st.subheader("🔧 تحديد أعمدة شيت 'اسعار المنتجات'")
            col1, col2, col3 = st.columns(3)
            with col1:
                sku_col = st.selectbox("عمود رقم المنتج (SKU)", options=list(df_regular_prices.columns) if not df_regular_prices.empty else ["-"])
            with col2:
                name_col = st.selectbox("عمود اسم المنتج", options=list(df_regular_prices.columns) if not df_regular_prices.empty else ["-"])
            with col3:
                price_col = st.selectbox("عمود السعر", options=list(df_regular_prices.columns) if not df_regular_prices.empty else ["-"])
            
            st.subheader("🔧 تحديد أعمدة شيت 'سعر مخفض'")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                disc_sku_col = st.selectbox("عمود رقم المنتج", options=list(df_discounted.columns) if not df_discounted.empty else ["-"])
            with col2:
                disc_price_col = st.selectbox("عمود السعر المخفض", options=list(df_discounted.columns) if not df_discounted.empty else ["-"])
            with col3:
                disc_end_col = st.selectbox("عمود تاريخ النهاية (اختياري)", options=["لا يوجد"] + list(df_discounted.columns) if not df_discounted.empty else ["-"])
            with col4:
                disc_promo_col = st.selectbox("عمود العنوان الترويجي (اختياري)", options=["لا يوجد"] + list(df_discounted.columns) if not df_discounted.empty else ["-"])
            
            # ========== 1. بناء قاموس الأسعار العادية ==========
            price_map = {}
            if not df_regular_prices.empty and sku_col != "-":
                for _, row in df_regular_prices.iterrows():
                    sku_val = str(row[sku_col]).strip()
                    if sku_val and sku_val != "nan":
                        price_map[sku_val] = {
                            "name": row[name_col] if name_col != "-" else "",
                            "price": row[price_col] if price_col != "-" else ""
                        }
            
            # ========== 2. بناء قاموس الأسعار المخفضة ==========
            discounted_map = {}
            if not df_discounted.empty and disc_sku_col != "-":
                for _, row in df_discounted.iterrows():
                    sku_raw = str(row[disc_sku_col]).strip()
                    if not sku_raw or sku_raw == "nan":
                        continue
                    
                    price_val = row[disc_price_col] if disc_price_col != "-" else ""
                    end_val = row[disc_end_col] if disc_end_col != "لا يوجد" else ""
                    promo_val = row[disc_promo_col] if disc_promo_col != "لا يوجد" else ""
                    
                    discounted_map[sku_raw] = {
                        "price": price_val,
                        "end_date": end_val,
                        "promo": promo_val
                    }
            
            # ========== 3. معالجة العروض من شيت "عرض خاص" ==========
            products = {}  # {base_sku: {offers: [], group_sku: None, group_qty: None}}
            
            for _, row in df_raw.iterrows():
                text = str(row.iloc[0]) if len(row) > 0 and pd.notna(row.iloc[0]) else ""
                if not text or text == "nan":
                    continue
                
                # استخراج أرقام المنتجات من النص
                numbers = re.findall(r'(?<!\d)(\d{4,6})(?!\d)', text)
                # استخراج المجموعات (مثل 16265*6)
                groups = re.findall(r'(\d{4,6})\*(\d+)', text)
                # استخراج المنتجات المتعددة (مثل 15000-16000-12000)
                dash_groups = re.findall(r'(\d{4,6})-(\d{4,6})-(\d{4,6})', text)
                
                # اسم العرض
                offer_name = "عرض خاص"
                if "خصم" in text:
                    if "القطعة الثانية" in text:
                        match = re.search(r'(خصم \d+% على القطعة الثانية)', text)
                        offer_name = match.group(1) if match else "خصم على القطعة الثانية"
                    elif "الحبة الثانية" in text:
                        match = re.search(r'(خصم \d+% على الحبة الثانية)', text)
                        offer_name = match.group(1) if match else "خصم على الحبة الثانية"
                elif "عرض" in text and "مجاناً" in text:
                    match = re.search(r'(عرض \d+\+\d+ مجاناً?)', text)
                    offer_name = match.group(1) if match else "عرض مجاني"
                elif "صفقة اليوم" in text:
                    offer_name = "صفقة اليوم"
                
                # التواريخ
                dates = re.findall(r'(\d{4}-\d{2}-\d{2} \d{2}:\d{2})', text)
                start_date = dates[0] if len(dates) > 0 else ""
                end_date = dates[1] if len(dates) > 1 else ""
                is_daily = "نعم" if "صفقة اليوم" in text else ""
                
                # معالجة الأرقام العادية
                for num in numbers:
                    if num not in products:
                        products[num] = {"offers": [], "group_sku": None, "group_qty": None}
                    products[num]["offers"].append({
                        "name": offer_name, "start": start_date, "end": end_date, "daily": is_daily
                    })
                
                # معالجة المجموعات (مثل 16265*6)
                for base, qty in groups:
                    group_sku = f"{base}*{qty}"
                    if base not in products:
                        products[base] = {"offers": [], "group_sku": group_sku, "group_qty": int(qty)}
                    else:
                        products[base]["group_sku"] = group_sku
                        products[base]["group_qty"] = int(qty)
                    # نضيف العرض أيضًا للمنتج الأساسي
                    products[base]["offers"].append({
                        "name": offer_name, "start": start_date, "end": end_date, "daily": is_daily
                    })
                
                # معالجة المنتجات المتعددة (مثل 15000-16000-12000)
                for group in dash_groups:
                    group_sku = "-".join(group)
                    for sku in group:
                        if sku not in products:
                            products[sku] = {"offers": [], "group_sku": group_sku, "group_qty": None}
                        else:
                            products[sku]["group_sku"] = group_sku
                        products[sku]["offers"].append({
                            "name": offer_name, "start": start_date, "end": end_date, "daily": is_daily
                        })
            
            # ========== 4. بناء النتيجة النهائية ==========
            results = []
            for base_sku, data in products.items():
                # معلومات السعر العادي
                prod_info = price_map.get(base_sku, {"name": "", "price": ""})
                product_name = prod_info["name"]
                product_price = prod_info["price"]
                
                # معلومات المجموعة
                group_sku = data.get("group_sku")
                group_qty = data.get("group_qty")
                group_info = price_map.get(group_sku, {"name": "", "price": ""}) if group_sku else {"name": "", "price": ""}
                group_name = group_info["name"]
                group_price = group_info["price"]
                
                # الأسعار المخفضة
                disc_prod = discounted_map.get(base_sku, {"price": "", "promo": ""})
                disc_group = discounted_map.get(group_sku, {"price": "", "promo": ""}) if group_sku else {"price": "", "promo": ""}
                
                # دمج العروض
                offers = data["offers"]
                unique_offers = {}
                for off in offers:
                    if off["name"] not in unique_offers:
                        unique_offers[off["name"]] = off
                
                offer_names = " | ".join([o["name"] for o in unique_offers.values()])
                offer_starts = " | ".join([o["start"] for o in unique_offers.values() if o["start"]])
                offer_ends = " | ".join([o["end"] for o in unique_offers.values() if o["end"]])
                is_daily = "نعم" if any(o["daily"] == "نعم" for o in unique_offers.values()) else ""
                
                results.append({
                    "رقم المنتج": base_sku,
                    "رقم المنتج للمجموعة": group_sku if group_sku else "",
                    "اسم المنتج": product_name,
                    "سعر المنتج": product_price,
                    "اسم المنتج للمجموعة": group_name,
                    "سعر المنتج للمجموعة": group_price,
                    "اسم العرض الخاص": offer_names,
                    "بداية العرض": offer_starts,
                    "نهاية العرض": offer_ends,
                    "سعر مخفض للمنتج": disc_prod["price"],
                    "سعر مخفض للمجموعة": disc_group["price"],
                    "عدد حبات المجموعة": group_qty if group_qty else "",
                    "العنوان الترويجي للمنتج": disc_prod["promo"],
                    "العنوان الترويجي للمجموعة": disc_group["promo"],
                    "صفقة اليوم": is_daily
                })
            
            # إضافة المنتجات التي لها سعر مخفض فقط (بدون عرض)
            for sku, disc in discounted_map.items():
                if sku not in products and not re.search(r'[*\-]', sku):
                    prod_info = price_map.get(sku, {"name": "", "price": ""})
                    results.append({
                        "رقم المنتج": sku,
                        "رقم المنتج للمجموعة": "",
                        "اسم المنتج": prod_info["name"],
                        "سعر المنتج": prod_info["price"],
                        "اسم المنتج للمجموعة": "",
                        "سعر المنتج للمجموعة": "",
                        "اسم العرض الخاص": "",
                        "بداية العرض": "",
                        "نهاية العرض": "",
                        "سعر مخفض للمنتج": disc["price"],
                        "سعر مخفض للمجموعة": "",
                        "عدد حبات المجموعة": "",
                        "العنوان الترويجي للمنتج": disc["promo"],
                        "العنوان الترويجي للمجموعة": "",
                        "صفقة اليوم": ""
                    })
            
            df_final = pd.DataFrame(results)
            df_final = df_final.drop_duplicates(subset=["رقم المنتج"], keep="first")
            
            # عرض النتيجة
            st.success(f"✅ تم معالجة {len(df_final)} منتج")
            st.dataframe(df_final, use_container_width=True, height=400)
            
            # خيارات التحميل
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
                                if offer_type:
                                    type_df = df[df["اسم العرض الخاص"] == offer_type]
                                    sheet_name = offer_type[:31]
                                    type_df.to_excel(writer, sheet_name=sheet_name, index=False)
                                    apply_excel_style(writer, sheet_name, type_df)
                    return output.getvalue()
                st.download_button("📥 تحميل ملف مفصل", data=create_detailed_excel(df_final), file_name="العروض_المفصلة.xlsx", use_container_width=True)
            
            with col2:
                def create_simple_excel(df):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, sheet_name="العروض والمنتجات", index=False)
                        apply_excel_style(writer, "العروض والمنتجات", df)
                    return output.getvalue()
                st.download_button("📥 تحميل ملف مبسط", data=create_simple_excel(df_final), file_name="العروض_المنتجات.xlsx", use_container_width=True)
            
            # فلترة
            st.subheader("🔍 تصفية حسب رقم المنتج")
            search = st.text_input("أدخل رقم المنتج")
            if search:
                filtered = df_final[(df_final["رقم المنتج"] == search) | (df_final["رقم المنتج للمجموعة"] == search)]
                st.dataframe(filtered, use_container_width=True)
    
    else:
        st.info("📂 يرجى رفع ملف Excel")
