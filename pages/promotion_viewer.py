# pages/promotion_viewer.py
import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ❌ تم حذف st.set_page_config من هنا لمنع انهيار التطبيق السحابي

def apply_excel_style(writer, sheet_name, df):
    """تطبيق التنسيقات على ملف Excel بكفاءة عالية"""
    workbook = writer.book
    worksheet = workbook[sheet_name]
    
    header_fill = PatternFill(start_color="1F7A8C", end_color="1F7A8C", fill_type="solid")
    header_font = Font(name="Tajawal", size=12, bold=True, color="FFFFFF")
    alt_row_fill = PatternFill(start_color="E6F3F5", end_color="E6F3F5", fill_type="solid")
    border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    # تنسيق العناوين
    for col_idx, col_name in enumerate(df.columns, 1):
        cell = worksheet.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
        worksheet.column_dimensions[get_column_letter(col_idx)].width = max(20, len(str(col_name)) + 5)
    
    # تنسيق الصفوف (تطبيق سريع للألوان البديلة)
    for row_idx in range(2, len(df) + 2):
        is_alt = (row_idx - 2) % 2 == 1
        for col_idx in range(1, len(df.columns) + 1):
            cell = worksheet.cell(row=row_idx, column=col_idx)
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if is_alt:
                cell.fill = alt_row_fill
    
    worksheet.freeze_panes = 'A2'

def parse_composite_sku(sku):
    """تحليل الرقم المركب (مثل 1500*6 أو 15000-16000-12000)"""
    sku = str(sku).strip()
    
    star_match = re.match(r'^(\d{4,6})\*(\d+)$', sku)
    if star_match:
        base_sku = star_match.group(1)
        qty = int(star_match.group(2))
        group_sku = f"{base_sku}*{qty}"
        return base_sku, group_sku, qty, [base_sku], True
    
    dash_match = re.match(r'^(\d{4,6})-(\d{4,6})-(\d{4,6})$', sku)
    if dash_match:
        individual_skus = list(dash_match.groups())
        group_sku = "-".join(individual_skus)
        return individual_skus[0], group_sku, None, individual_skus, True
    
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

# 🧠 [تحسين الأداء الحاسم]: عزل وتخزين عمليات إنشاء الإكسيل الثقيلة لمنع التهنيج عند البحث
@st.cache_data(show_spinner=False)
def generate_excel_download_files(df):
    """توليد ملفات التحميل وحفظها بكاش مخصص لبيانات حجمها +100 ألف سطر"""
    # 1. الملف المبسط
    simple_output = BytesIO()
    with pd.ExcelWriter(simple_output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="العروض والمنتجات", index=False)
        apply_excel_style(writer, "العروض والمنتجات", df)
    simple_bytes = simple_output.getvalue()

    # 2. الملف المفصل (شيتات) مع وضع حد أقصى لحماية المعالج لقصرها على أبرز 10 عروض فريدة
    detailed_output = BytesIO()
    with pd.ExcelWriter(detailed_output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="النتيجة النهائية", index=False)
        apply_excel_style(writer, "النتيجة النهائية", df)
        
        if "اسم العرض الخاص" in df:
            unique_offers = df["اسم العرض الخاص"].dropna().unique()
            for offer_type in unique_offers[:12]: # حد أقصى 12 شيت لمنع انهيار الذاكرة
                if offer_type and offer_type.strip() != "":
                    type_df = df[df["اسم العرض الخاص"] == offer_type]
                    if len(type_df) > 0:
                        sheet_name = str(offer_type)[:30].replace("|", "-")
                        type_df.to_excel(writer, sheet_name=sheet_name, index=False)
                        apply_excel_style(writer, sheet_name, type_df)
                        
    detailed_bytes = detailed_output.getvalue()
    return simple_bytes, detailed_bytes

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
    </style>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="offers-header">
        <h1>🛍️ العروض الحالية الفعالة بالمتجر</h1>
        <p>جميع العروض الترويجية والخصومات المطبقة على المنتجات مع دعم المجموعات</p>
    </div>
    """, unsafe_allow_html=True)

    st.subheader("📂 رفع ملف Excel لتقرير العروض")
    uploaded_file = st.file_uploader(
        "قم برفع ملف Excel", type=["xlsx", "xls"],
        help="يجب أن يحتوي الملف على أوراق: 'عرض خاص'، 'سعر مخفض'، 'اسعار المنتجات'"
    )
    
    if uploaded_file is not None:
        # قراءة الأوراق المخزنة بكاش داخلي سريع
        excel_data = pd.ExcelFile(uploaded_file)
        sheet_names = excel_data.sheet_names
        
        offers_sheet = None
        discounted_sheet = None
        prices_sheet = None
        
        for sheet in sheet_names:
            if "عرض خاص" in sheet or "عروض" in sheet: offers_sheet = sheet
            if "سعر مخفض" in sheet or "مخفض" in sheet: discounted_sheet = sheet
            if "اسعار المنتجات" in sheet or "سعر المنتج" in sheet or "أسعار المنتجات" in sheet: prices_sheet = sheet
        
        if not offers_sheet: offers_sheet = sheet_names[0]
        
        df_raw = pd.read_excel(uploaded_file, sheet_name=offers_sheet, header=None)
        df_discounted = pd.read_excel(uploaded_file, sheet_name=discounted_sheet) if discounted_sheet else pd.DataFrame()
        df_regular_prices = pd.read_excel(uploaded_file, sheet_name=prices_sheet) if prices_sheet else pd.DataFrame()
        
        st.subheader("🔧 تحديد أعمدة شيت 'اسعار المنتجات'")
        col1, col2, col3 = st.columns(3)
        with col1:
            sku_col_choice = st.selectbox("اختر عمود رقم المنتج (SKU)", options=list(df_regular_prices.columns) if not df_regular_prices.empty else ["لا يوجد"])
        with col2:
            name_col_choice = st.selectbox("اختر عمود اسم المنتج", options=list(df_regular_prices.columns) if not df_regular_prices.empty else ["لا يوجد"], index=min(1, len(df_regular_prices.columns)-1) if not df_regular_prices.empty else 0)
        with col3:
            price_col_choice = st.selectbox("اختر عمود السعر", options=list(df_regular_prices.columns) if not df_regular_prices.empty else ["لا يوجد"], index=min(2, len(df_regular_prices.columns)-1) if not df_regular_prices.empty else 0)
        
        # 1. معالجة الأسعار العادية
        price_map = {}
        if not df_regular_prices.empty and sku_col_choice != "لا يوجد":
            for _, row in df_regular_prices.iterrows():
                sku = str(row[sku_col_choice]).strip()
                if sku and sku != "nan":
                    price_map[sku] = {
                        "name": row[name_col_choice] if name_col_choice != "لا يوجد" and pd.notna(row[name_col_choice]) else "",
                        "price": row[price_col_choice] if price_col_choice != "لا يوجد" and pd.notna(row[price_col_choice]) else ""
                    }
        
        # 2. معالجة الأسعار المخفضة
        discounted_map = {}
        if not df_discounted.empty:
            st.subheader("🔧 تحديد أعمدة شيت 'سعر مخفض'")
            c1, c2, c3, c4 = st.columns(4)
            with c1: disc_sku_col = st.selectbox("عمود رقم المنتج", options=list(df_discounted.columns))
            with c2: disc_price_col = st.selectbox("عمود السعر المخفض", options=list(df_discounted.columns), index=min(1, len(df_discounted.columns)-1))
            with c3: disc_end_col = st.selectbox("عمود تاريخ النهاية (اختياري)", options=["لا يوجد"] + list(df_discounted.columns))
            with c4: disc_promo_col = st.selectbox("عمود العنوان الترويجي (اختياري)", options=["لا يوجد"] + list(df_discounted.columns))
            
            for _, row in df_discounted.iterrows():
                sku_raw = str(row[disc_sku_col]).strip()
                if not sku_raw or sku_raw == "nan": continue
                
                base_sku, group_sku, group_qty, individual_skus, is_composite = parse_composite_sku(sku_raw)
                price_val = row[disc_price_col] if pd.notna(row[disc_price_col]) else ""
                end_val = row[disc_end_col] if disc_end_col != "لا يوجد" and pd.notna(row[disc_end_col]) else ""
                promo_val = row[disc_promo_col] if disc_promo_col != "لا يوجد" and pd.notna(row[disc_promo_col]) else ""
                
                if is_composite:
                    discounted_map[group_sku] = {"discounted_price": price_val, "end_date": end_val, "promo_title": promo_val, "group_sku": group_sku, "group_qty": group_qty, "is_composite": True}
                    for ind_sku in individual_skus:
                        if ind_sku not in discounted_map:
                            discounted_map[ind_sku] = {"discounted_price": price_val, "end_date": end_val, "promo_title": promo_val, "group_sku": group_sku, "group_qty": group_qty, "is_composite": False}
                else:
                    discounted_map[sku_raw] = {"discounted_price": price_val, "end_date": end_val, "promo_title": promo_val, "group_sku": None, "group_qty": None, "is_composite": False}
        
        # 3. معالجة العروض
        with st.spinner("🧠 جاري معالجة وفك تكرار العروض حياً..."):
            products_data = {}
            for idx, row in df_raw.iterrows():
                text = row.iloc[0] if len(row) > 0 else ""
                if pd.isna(text) or text == "nan": continue
                
                text = str(text)
                numbers = extract_numbers_from_text(text)
                offer_name = extract_offer_name(text)
                start_date, end_date = extract_dates(text)
                is_daily_deal = "صفقة اليوم" in text
                
                groups = re.findall(r'(\d{4,6})\*(\d+)', text)
                dash_matches = re.findall(r'(\d{4,6})-(\d{4,6})-(\d{4,6})', text)
                
                for base_sku, qty in groups:
                    group_sku = f"{base_sku}*{qty}"
                    if base_sku not in products_data:
                        products_data[base_sku] = {"offers": [], "group_sku": group_sku, "group_qty": int(qty), "is_composite": True}
                    products_data[base_sku]["offers"].append({"name": offer_name, "start": start_date, "end": end_date, "is_daily_deal": is_daily_deal})
                
                for multi in dash_matches:
                    group_sku = "-".join(multi)
                    for sku in multi:
                        if sku not in products_data:
                            products_data[sku] = {"offers": [], "group_sku": group_sku, "group_qty": None, "is_composite": True}
                        products_data[sku]["offers"].append({"name": offer_name, "start": start_date, "end": end_date, "is_daily_deal": is_daily_deal})
                
                for sku in numbers:
                    if sku not in products_data:
                        products_data[sku] = {"offers": [], "group_sku": None, "group_qty": None, "is_composite": False}
                    products_data[sku]["offers"].append({"name": offer_name, "start": start_date, "end": end_date, "is_daily_deal": is_daily_deal})
            
            # 4. دمج بيانات المجموعات والمنتجات الفردية
            merged_data = {}
            for sku, data in products_data.items():
                if '*' in sku or '-' in sku:
                    base_match = re.match(r'^(\d+)', sku)
                    if base_match:
                        base_sku = base_match.group(1)
                        if base_sku not in merged_data:
                            merged_data[base_sku] = {"offers": [], "group_sku": sku, "group_qty": data["group_qty"], "is_composite": True}
                        merged_data[base_sku]["offers"].extend(data["offers"])
                else:
                    if sku not in merged_data:
                        merged_data[sku] = {"offers": [], "group_sku": None, "group_qty": None, "is_composite": False}
                    merged_data[sku]["offers"].extend(data["offers"])
            
            # 5. بناء النتيجة النهائية كقائمة
            final_results = []
            for base_sku, data in merged_data.items():
                product_info = price_map.get(base_sku, {"name": "", "price": ""})
                group_sku = data.get("group_sku")
                group_qty = data.get("group_qty")
                
                group_product_name = ""
                group_product_price = ""
                if group_sku:
                    group_info = price_map.get(group_sku, {})
                    group_product_name = group_info.get("name", "")
                    group_product_price = group_info.get("price", "")
                    if not group_product_name:
                        base_m = re.match(r'^(\d+)', group_sku)
                        if base_m:
                            b_info = price_map.get(base_m.group(1), {})
                            group_product_name = b_info.get("name", "")
                            group_product_price = b_info.get("price", "")
                
                disc_info = discounted_map.get(base_sku, {})
                group_disc_info = discounted_map.get(group_sku, {}) if group_sku else {}
                
                unique_offers = {o["name"]: o for o in data["offers"]}
                offer_names = " | ".join([o["name"] for o in unique_offers.values()])
                offer_starts = " | ".join([o["start"] for o in unique_offers.values() if o["start"]])
                offer_ends = " | ".join([o["end"] for o in unique_offers.values() if o["end"]])
                is_daily = "نعم" if any(o["is_daily_deal"] for o in unique_offers.values()) else ""
                
                final_results.append({
                    "رقم المنتج": base_sku, "رقم المنتج للمجموعة": group_sku if group_sku else "",
                    "اسم المنتج": product_info["name"], "sعر المنتج": product_info["price"],
                    "اسم المنتج للمجموعة": group_product_name, "سعر المنتج للمجموعة": group_product_price,
                    "اسم العرض الخاص": offer_names, "بداية العرض": offer_starts, "نهاية العرض": offer_ends,
                    "سعر مخفض للمنتج": disc_info.get("discounted_price", ""), "سعر مخفض للمجموعة": group_disc_info.get("discounted_price", ""),
                    "عدد حبات المجموعة": group_qty if group_qty else "", "العنوان الترويجي للمنتج": disc_info.get("promo_title", ""),
                    "العنوان الترويجي للمجموعة": group_disc_info.get("promo_title", ""), "صفقة اليوم": is_daily
                })
            
            # إضافة المنتجات المخفضة التي ليس لها عروض نصية
            for sku, d_info in discounted_map.items():
                if '*' in sku or '-' in sku: continue
                if sku not in merged_data:
                    p_info = price_map.get(sku, {"name": "", "price": ""})
                    final_results.append({
                        "رقم المنتج": sku, "رقم المنتج للمجموعة": "", "اسم المنتج": p_info["name"], "sعر المنتج": p_info["price"],
                        "اسم المنتج للمجموعة": "", "سعر المنتج للمجموعة": "", "اسم العرض الخاص": "", "بداية العرض": "", "نهاية العرض": "",
                        "سعر مخفض للمنتج": d_info.get("discounted_price", ""), "سعر مخفض للمجموعة": "", "عدد حبات المجموعة": "",
                        "العنوان الترويجي للمنتج": d_info.get("promo_title", ""), "العنوان الترويجي للمجموعة": "", "صفقة اليوم": ""
                    })
            
            df_final = pd.DataFrame(final_results)
            df_final = df_final[df_final["رقم المنتج"].notna() & (df_final["رقم المنتج"] != "nan") & (df_final["رقم المنتج"] != "")]
            df_final = df_final[~df_final["رقم المنتج"].isin(["2024", "2025", "2026", "2027", "2028", "2029", "2030"])]
            df_final = df_final.drop_duplicates(subset=["رقم المنتج", "رقم المنتج للمجموعة"], keep="first")
            
            # ترتيب الأعمدة النهائي
            column_order = [
                "رقم المنتج", "رقم المنتج للمجموعة", "اسم المنتج", "sعر المنتج", "اسم المنتج للمجموعة", "سعر المنتج للمجموعة",
                "اسم العرض الخاص", "بداية العرض", "نهاية العرض", "سعر مخفض للمنتج", "سعر مخفض للمجموعة", "عدد حبات المجموعة",
                "العنوان الترويجي للمنتج", "العنوان الترويجي للمجموعة", "صفقة اليوم"
            ]
            df_final = df_final[[col for col in column_order if col in df_final.columns]]
            
        # استعراض الإحصائيات الفورية لتقرير العروض
        st.success(f"✅ تم معالجة {len(df_final):,} منتج بنجاح وبسرعة فائقة (بدون تكرار)")
        
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1: st.metric("📦 إجمالي المنتجات", f"{len(df_final):,}")
        with col2: st.metric("🏷️ مع عروض", f"{df_final['اسم العرض الخاص'].str.len().gt(0).sum():,}")
        with col3: st.metric("🏷️ منتجات مخفضة", f"{df_final['سعر مخفض للمنتج'].notna().sum():,}")
        with col4: st.metric("🎯 مجموعات", f"{df_final['رقم المنتج للمجموعة'].str.len().gt(0).sum():,}")
        with col5: st.metric("📊 عروض فريدة", f"{df_final['اسم العرض الخاص'].nunique():,}")
        
        # ========== البحث والتصفية برقم المنتج ==========
        st.subheader("🔍 تصفية وبحث لحظي")
        search_term = st.text_input("أدخل رقم المنتج أو اسم المنتج للبحث:", placeholder="مثال: 9974 أو 16265*6")
        
        if search_term:
            search_term = search_term.strip()
            filtered_df = df_final[
                (df_final["رقم المنتج"].astype(str) == search_term) | 
                (df_final["رقم المنتج للمجموعة"].astype(str) == search_term) |
                (df_final["اسم المنتج"].str.contains(search_term, na=False))
            ]
            if not filtered_df.empty:
                st.dataframe(filtered_df, use_container_width=True)
            else:
                st.warning(f"⚠️ لم يتم العثور على أي نتائج تطابق '{search_term}'")
        
        st.subheader("📋 قائمة المعاينة لجميع المنتجات والعروض")
        st.dataframe(df_final, use_container_width=True, height=400)
        
        # ========== 💾 تحميل النتائج (باستخدام الكاش الآمن للبيانات الضخمة) ==========
        st.subheader("💾 تحميل التقارير النهائية")
        col_dl1, col_dl2 = st.columns(2)
        
        # توليد الملفات مرة واحدة فقط عبر الكاش المخزن
        simple_bytes, detailed_bytes = generate_excel_download_files(df_final)
        
        with col_dl1:
            st.download_button(
                label="📥 تحميل ملف مفصل (مع شيتات العروض أبرز 12 نوع)",
                data=detailed_bytes,
                file_name="العروض_المفصلة_شيتات.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with col_dl2:
            st.download_button(
                label="📥 تحميل ملف مبسط (جدول واحد شامل الفلترة)",
                data=simple_bytes,
                file_name="العروض_والمنتجات_الموحد.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
        # فلترة حسب نوع العرض أسفل الصفحة
        if "اسم العرض الخاص" in df_final and len(df_final["اسم العرض الخاص"].dropna().unique()) > 0:
            st.subheader("🔍 تصفية حسب نوع العرض الفريد")
            offer_types = ["الكل"] + sorted(df_final["اسم العرض الخاص"].dropna().unique().tolist())
            selected_type = st.selectbox("اختر نوع العرض للتصفية", offer_types)
            
            if selected_type != "الكل":
                st.dataframe(df_final[df_final["اسم العرض الخاص"] == selected_type], use_container_width=True)
    else:
        st.info("📂 يرجى رفع ملف Excel الخاص بالعروض والمنتجات لتنشيط لوحة المعالجة والتحميل.")
