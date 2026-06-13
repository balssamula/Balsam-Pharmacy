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

def parse_composite_sku(sku):
    """
    تحليل الرقم المركب (مثل 1500*6 أو 15000-16000-12000)
    إرجاع: (base_sku, group_sku, group_qty, individual_skus, is_composite)
    """
    sku = str(sku).strip()
    
    # حالة: 1500*6
    star_match = re.match(r'^(\d{4,6})\*(\d+)$', sku)
    if star_match:
        base_sku = star_match.group(1)
        qty = int(star_match.group(2))
        group_sku = f"{base_sku}*{qty}"
        return base_sku, group_sku, qty, [base_sku], True
    
    # حالة: 15000-16000-12000
    dash_match = re.match(r'^(\d{4,6})-(\d{4,6})-(\d{4,6})$', sku)
    if dash_match:
        individual_skus = list(dash_match.groups())
        group_sku = "-".join(individual_skus)
        return individual_skus[0], group_sku, None, individual_skus, True
    
    # حالة: رقم عادي
    return sku, None, None, [sku], False

def extract_offer_name(text):
    """استخراج اسم العرض من النص"""
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
    """استخراج التواريخ من النص"""
    text = str(text) if not pd.isna(text) else ""
    date_match = re.findall(r'(\d{4}-\d{2}-\d{2} \d{2}:\d{2})', text)
    start = date_match[0] if len(date_match) > 0 else None
    end = date_match[1] if len(date_match) > 1 else None
    return start, end

def extract_numbers_from_text(text):
    """استخراج الأرقام الفردية من النص (للعروض)"""
    if pd.isna(text):
        return []
    text = str(text)
    excluded_years = {'2024', '2025', '2026', '2027', '2028', '2029', '2030'}
    pattern = r'(?:^|[-/\s]+)(\d{4,6})(?:[-/\s]|$)'
    matches = re.findall(pattern, text)
    return [m for m in matches if m not in excluded_years and not m.startswith('20')]

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
                if "اسعار المنتجات" in sheet or "سعر المنتج" in sheet or "أسعار المنتجات" in sheet:
                    prices_sheet = sheet
            
            if not offers_sheet:
                offers_sheet = sheet_names[0]
            
            # قراءة البيانات
            df_raw = pd.read_excel(uploaded_file, sheet_name=offers_sheet, header=None)
            df_discounted = pd.read_excel(uploaded_file, sheet_name=discounted_sheet) if discounted_sheet else pd.DataFrame()
            df_regular_prices = pd.read_excel(uploaded_file, sheet_name=prices_sheet) if prices_sheet else pd.DataFrame()
            
            # ========== عرض أسماء الأعمدة للمستخدم لاختيارها ==========
            st.subheader("🔧 تحديد أعمدة شيت 'اسعار المنتجات'")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                sku_col_choice = st.selectbox(
                    "اختر عمود رقم المنتج (SKU)",
                    options=list(df_regular_prices.columns) if not df_regular_prices.empty else ["لا يوجد"],
                    index=0 if not df_regular_prices.empty else None
                )
            with col2:
                name_col_choice = st.selectbox(
                    "اختر عمود اسم المنتج",
                    options=list(df_regular_prices.columns) if not df_regular_prices.empty else ["لا يوجد"],
                    index=min(1, len(df_regular_prices.columns)-1) if not df_regular_prices.empty else None
                )
            with col3:
                price_col_choice = st.selectbox(
                    "اختر عمود السعر",
                    options=list(df_regular_prices.columns) if not df_regular_prices.empty else ["لا يوجد"],
                    index=min(2, len(df_regular_prices.columns)-1) if not df_regular_prices.empty else None
                )
            
            # ========== 1. معالجة الأسعار العادية ==========
            price_map = {}  # {sku: {name, price}}
            if not df_regular_prices.empty and sku_col_choice != "لا يوجد":
                for _, row in df_regular_prices.iterrows():
                    sku = str(row[sku_col_choice]).strip()
                    if sku and sku != "nan":
                        price_map[sku] = {
                            "name": row[name_col_choice] if name_col_choice != "لا يوجد" and pd.notna(row[name_col_choice]) else "",
                            "price": row[price_col_choice] if price_col_choice != "لا يوجد" and pd.notna(row[price_col_choice]) else ""
                        }
            
            # ========== 2. معالجة الأسعار المخفضة (مع فك الأرقام المركبة) ==========
            discounted_map = {}  # {sku: {discounted_price, end_date, promo_title, group_sku, group_qty}}
            
            if not df_discounted.empty:
                # تحديد أعمدة شيت السعر المخفض
                st.subheader("🔧 تحديد أعمدة شيت 'سعر مخفض'")
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    disc_sku_col = st.selectbox(
                        "عمود رقم المنتج",
                        options=list(df_discounted.columns),
                        index=0
                    )
                with col2:
                    disc_price_col = st.selectbox(
                        "عمود السعر المخفض",
                        options=list(df_discounted.columns),
                        index=min(1, len(df_discounted.columns)-1)
                    )
                with col3:
                    disc_end_col = st.selectbox(
                        "عمود تاريخ النهاية (اختياري)",
                        options=["لا يوجد"] + list(df_discounted.columns),
                        index=0
                    )
                with col4:
                    disc_promo_col = st.selectbox(
                        "عمود العنوان الترويجي (اختياري)",
                        options=["لا يوجد"] + list(df_discounted.columns),
                        index=0
                    )
                
                for _, row in df_discounted.iterrows():
                    sku_raw = str(row[disc_sku_col]).strip()
                    if not sku_raw or sku_raw == "nan":
                        continue
                    
                    # فك الرقم المركب
                    base_sku, group_sku, group_qty, individual_skus, is_composite = parse_composite_sku(sku_raw)
                    
                    price_val = row[disc_price_col] if disc_price_col and pd.notna(row[disc_price_col]) else ""
                    end_val = row[disc_end_col] if disc_end_col != "لا يوجد" and pd.notna(row[disc_end_col]) else ""
                    promo_val = row[disc_promo_col] if disc_promo_col != "لا يوجد" and pd.notna(row[disc_promo_col]) else ""
                    
                    if is_composite:
                        # تخزين للمجموعة
                        discounted_map[group_sku] = {
                            "discounted_price": price_val,
                            "end_date": end_val,
                            "promo_title": promo_val,
                            "group_sku": group_sku,
                            "group_qty": group_qty,
                            "is_composite": True
                        }
                        # تخزين لكل منتج فردي في المجموعة (لربطه لاحقًا)
                        for ind_sku in individual_skus:
                            if ind_sku not in discounted_map:
                                discounted_map[ind_sku] = {
                                    "discounted_price": price_val,
                                    "end_date": end_val,
                                    "promo_title": promo_val,
                                    "group_sku": group_sku,
                                    "group_qty": group_qty,
                                    "is_composite": False
                                }
                    else:
                        # منتج عادي
                        discounted_map[sku_raw] = {
                            "discounted_price": price_val,
                            "end_date": end_val,
                            "promo_title": promo_val,
                            "group_sku": None,
                            "group_qty": None,
                            "is_composite": False
                        }
            
            # ========== 3. معالجة العروض (شيت عرض خاص) ==========
            products_data = {}  # {base_sku: {offers: [], group_sku: None, group_qty: None, is_composite: False}}
            
            for idx, row in df_raw.iterrows():
                text = row.iloc[0] if len(row) > 0 else ""
                if pd.isna(text) or text == "nan":
                    continue
                
                text = str(text)
                
                # استخراج الأرقام الفردية
                numbers = extract_numbers_from_text(text)
                offer_name = extract_offer_name(text)
                start_date, end_date = extract_dates(text)
                is_daily_deal = "صفقة اليوم" in text
                
                # استخراج المجموعات (مثل 1500*6)
                group_pattern = r'(\d{4,6})\*(\d+)'
                groups = re.findall(group_pattern, text)
                
                # استخراج الأرقام المتعددة المفصولة بـ -
                dash_pattern = r'(\d{4,6})-(\d{4,6})-(\d{4,6})'
                dash_matches = re.findall(dash_pattern, text)
                
                # معالجة المجموعات (مثل 1500*6)
                for base_sku, qty in groups:
                    group_sku = f"{base_sku}*{qty}"
                    if base_sku not in products_data:
                        products_data[base_sku] = {
                            "offers": [],
                            "group_sku": group_sku,
                            "group_qty": int(qty),
                            "is_composite": True
                        }
                    products_data[base_sku]["offers"].append({
                        "name": offer_name,
                        "start": start_date,
                        "end": end_date,
                        "is_daily_deal": is_daily_deal
                    })
                
                # معالجة المنتجات المتعددة في عرض واحد (مثل 15000-16000-12000)
                for multi in dash_matches:
                    group_sku = "-".join(multi)
                    for sku in multi:
                        if sku not in products_data:
                            products_data[sku] = {
                                "offers": [],
                                "group_sku": group_sku,
                                "group_qty": None,
                                "is_composite": True
                            }
                        products_data[sku]["offers"].append({
                            "name": offer_name,
                            "start": start_date,
                            "end": end_date,
                            "is_daily_deal": is_daily_deal
                        })
                
                # معالجة الأرقام الفردية (منتجات عادية)
                for sku in numbers:
                    if sku not in products_data:
                        products_data[sku] = {
                            "offers": [],
                            "group_sku": None,
                            "group_qty": None,
                            "is_composite": False
                        }
                    products_data[sku]["offers"].append({
                        "name": offer_name,
                        "start": start_date,
                        "end": end_date,
                        "is_daily_deal": is_daily_deal
                    })
            
            # ========== 4. دمج بيانات المجموعة مع المنتج العادي في صف واحد ==========
            # سنقوم بإنشاء قاموس جديد حيث المفتاح هو المنتج العادي (base_sku)
            merged_data = {}
            
            for sku, data in products_data.items():
                # إذا كان sku يحتوي على * أو -، فهو منتج مركب (مجموعة)، نتعامل معه بشكل منفصل
                if '*' in sku or '-' in sku:
                    # نبحث عن المنتج العادي المقابل (قاعدة)
                    base_match = re.match(r'^(\d+)', sku)
                    if base_match:
                        base_sku = base_match.group(1)
                        if base_sku not in merged_data:
                            # ننشئ سجل جديد للمنتج العادي ونضيف معلومات المجموعة
                            merged_data[base_sku] = {
                                "offers": [],
                                "group_sku": sku,
                                "group_qty": data["group_qty"],
                                "is_composite": True
                            }
                        # نضيف عروض المجموعة إلى المنتج العادي
                        merged_data[base_sku]["offers"].extend(data["offers"])
                else:
                    # منتج عادي
                    if sku not in merged_data:
                        merged_data[sku] = {
                            "offers": [],
                            "group_sku": None,
                            "group_qty": None,
                            "is_composite": False
                        }
                    merged_data[sku]["offers"].extend(data["offers"])
            
            # ========== 5. بناء النتيجة النهائية ==========
            final_results = []
            
            for base_sku, data in merged_data.items():
                # جلب معلومات السعر العادي
                product_info = price_map.get(base_sku, {"name": "", "price": ""})
                product_name = product_info["name"]
                product_price = product_info["price"]
                
                # جلب معلومات المجموعة (إن وجدت)
                group_sku = data.get("group_sku")
                group_qty = data.get("group_qty")
                
                # اسم وسعر المنتج للمجموعة
                group_product_name = ""
                group_product_price = ""
                if group_sku:
                    group_info = price_map.get(group_sku, {})
                    group_product_name = group_info.get("name", "")
                    group_product_price = group_info.get("price", "")
                    if not group_product_name and group_sku:
                        base_match = re.match(r'^(\d+)', group_sku)
                        if base_match:
                            base_info = price_map.get(base_match.group(1), {})
                            group_product_name = base_info.get("name", "")
                            group_product_price = base_info.get("price", "")
                
                # جلب بيانات السعر المخفض للمنتج الفردي
                disc_info = discounted_map.get(base_sku, {})
                disc_price = disc_info.get("discounted_price", "")
                disc_promo = disc_info.get("promo_title", "")
                
                # جلب بيانات السعر المخفض للمجموعة
                group_disc_info = discounted_map.get(group_sku, {}) if group_sku else {}
                group_disc_price = group_disc_info.get("discounted_price", "")
                group_disc_promo = group_disc_info.get("promo_title", "")
                
                # دمج العروض المتعددة لنفس المنتج
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
                
                final_results.append({
                    "رقم المنتج": base_sku,
                    "رقم المنتج للمجموعة": group_sku if group_sku else "",
                    "اسم المنتج": product_name,
                    "سعر المنتج": product_price,
                    "اسم المنتج للمجموعة": group_product_name,
                    "سعر المنتج للمجموعة": group_product_price,
                    "اسم العرض الخاص": offer_names,
                    "بداية العرض": offer_starts,
                    "نهاية العرض": offer_ends,
                    "سعر مخفض للمنتج": disc_price,
                    "سعر مخفض للمجموعة": group_disc_price,
                    "عدد حبات المجموعة": group_qty if group_qty else "",
                    "العنوان الترويجي للمنتج": disc_promo,
                    "العنوان الترويجي للمجموعة": group_disc_promo,
                    "صفقة اليوم": is_daily
                })
            
            # إضافة المنتجات التي لها سعر مخفض فقط (بدون عرض في شيت العروض)
            for sku, disc_info in discounted_map.items():
                # تجاهل المنتجات المركبة (سنضيفها مع المنتج العادي)
                if '*' in sku or '-' in sku:
                    continue
                if sku not in merged_data:
                    product_info = price_map.get(sku, {"name": "", "price": ""})
                    final_results.append({
                        "رقم المنتج": sku,
                        "رقم المنتج للمجموعة": "",
                        "اسم المنتج": product_info["name"],
                        "سعر المنتج": product_info["price"],
                        "اسم المنتج للمجموعة": "",
                        "سعر المنتج للمجموعة": "",
                        "اسم العرض الخاص": "",
                        "بداية العرض": "",
                        "نهاية العرض": "",
                        "سعر مخفض للمنتج": disc_info.get("discounted_price", ""),
                        "سعر مخفض للمجموعة": "",
                        "عدد حبات المجموعة": "",
                        "العنوان الترويجي للمنتج": disc_info.get("promo_title", ""),
                        "العنوان الترويجي للمجموعة": "",
                        "صفقة اليوم": ""
                    })
            
            df_final = pd.DataFrame(final_results)
            
            # تنظيف البيانات
            df_final = df_final[df_final["رقم المنتج"].notna()]
            df_final = df_final[df_final["رقم المنتج"] != "nan"]
            df_final = df_final[df_final["رقم المنتج"] != ""]
            df_final = df_final[~df_final["رقم المنتج"].isin(["2024", "2025", "2026", "2027", "2028", "2029", "2030"])]
            
            # إزالة أي تكرارات متبقية (نفس رقم المنتج ونفس رقم المجموعة)
            df_final = df_final.drop_duplicates(subset=["رقم المنتج", "رقم المنتج للمجموعة"], keep="first")
            
            # ترتيب الأعمدة
            column_order = [
                "رقم المنتج", "رقم المنتج للمجموعة", "اسم المنتج", "سعر المنتج",
                "اسم المنتج للمجموعة", "سعر المنتج للمجموعة",
                "اسم العرض الخاص", "بداية العرض", "نهاية العرض",
                "سعر مخفض للمنتج", "سعر مخفض للمجموعة", "عدد حبات المجموعة",
                "العنوان الترويجي للمنتج", "العنوان الترويجي للمجموعة",
                "صفقة اليوم"
            ]
            df_final = df_final[[col for col in column_order if col in df_final.columns]]
            
            # عرض الإحصائيات
            st.success(f"✅ تم معالجة {len(df_final)} منتج بنجاح (بدون تكرار)")
            
            col1, col2, col3, col4, col5 = st.columns(5)
            with col1:
                st.metric("📦 إجمالي المنتجات", len(df_final))
            with col2:
                st.metric("🏷️ مع عروض", df_final["اسم العرض الخاص"].str.len().gt(0).sum())
            with col3:
                st.metric("🏷️ منتجات مخفضة", df_final["سعر مخفض للمنتج"].notna().sum())
            with col4:
                st.metric("🎯 مجموعات", df_final["رقم المنتج للمجموعة"].str.len().gt(0).sum())
            with col5:
                st.metric("📊 عروض فريدة", df_final["اسم العرض الخاص"].nunique())
            
            # ========== تصفية برقم المنتج ==========
            st.subheader("🔍 تصفية برقم المنتج")
            search_term = st.text_input("أدخل رقم المنتج (فردي أو مجموعة):", placeholder="مثال: 9974 أو 16265*6")
            
            if search_term:
                search_term = search_term.strip()
                filtered_df = df_final[
                    (df_final["رقم المنتج"] == search_term) | 
                    (df_final["رقم المنتج للمجموعة"] == search_term) |
                    (df_final["اسم المنتج"].str.contains(search_term, na=False))
                ]
                if len(filtered_df) > 0:
                    st.dataframe(filtered_df, use_container_width=True)
                else:
                    st.warning(f"⚠️ لم يتم العثور على منتج بالرقم '{search_term}'")
            
            # عرض الجدول الكامل
            st.subheader("📋 قائمة جميع المنتجات والعروض")
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
            if "اسم العرض الخاص" in df_final and len(df_final["اسم العرض الخاص"].dropna().unique()) > 0:
                st.subheader("🔍 تصفية حسب نوع العرض")
                offer_types = ["الكل"] + sorted(df_final["اسم العرض الخاص"].dropna().unique().tolist())
                selected_type = st.selectbox("اختر نوع العرض", offer_types)
                
                if selected_type != "الكل":
                    filtered_df = df_final[df_final["اسم العرض الخاص"] == selected_type]
                else:
                    filtered_df = df_final
                
                st.dataframe(filtered_df, use_container_width=True)
            
    else:
        st.info("📂 يرجى رفع ملف Excel لعرض العروض")
        
        with st.expander("ℹ️ تعليمات"):
            st.markdown("""
            **كيفية استخدام هذه الصفحة:**
            
            1. قم برفع ملف Excel الذي يحتوي على:
               - ورقة "عرض خاص" (العروض الترويجية)
               - ورقة "سعر مخفض" (المنتجات المخفضة)
               - ورقة "اسعار المنتجات" (الأسعار العادية للمنتجات)
            
            2. سيتم عرض أسماء الأعمدة في كل شيت، يمكنك اختيار الأعمدة الصحيحة من القوائم المنسدلة
            
            3. سيتم معالجة:
               - **المنتجات العادية** (مثل: 9974)
               - **المجموعات** (مثل: 1500*6)
               - **المنتجات المتعددة في عرض واحد** (مثل: 15000-16000-12000)
            
            4. **دمج العروض:** إذا كان نفس المنتج لديه أكثر من عرض، سيتم دمجها في صف واحد
            
            5. يمكنك **التصفية برقم المنتج** (فردي أو مجموعة) للبحث عن منتج معين
            
            **الأعمدة الناتجة:**
            - رقم المنتج
            - رقم المنتج للمجموعة
            - اسم المنتج
            - سعر المنتج
            - اسم المنتج للمجموعة (جديد)
            - سعر المنتج للمجموعة (جديد)
            - اسم العرض الخاص
            - بداية العرض
            - نهاية العرض
            - سعر مخفض للمنتج
            - سعر مخفض للمجموعة
            - عدد حبات المجموعة
            - العنوان الترويجي للمنتج
            - العنوان الترويجي للمجموعة
            - صفقة اليوم
            """)
