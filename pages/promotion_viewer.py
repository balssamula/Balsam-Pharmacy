# pages/promotion_viewer.py
import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def apply_excel_style(writer, sheet_name, df):
    """تطبيق التنسيقات الاحترافية على ملف Excel بكفاءة عالية"""
    workbook = writer.book
    worksheet = workbook[sheet_name]
    
    header_fill = PatternFill(start_color="1F7A8C", end_color="1F7A8C", fill_type="solid")
    header_font = Font(name="Tajawal", size=12, bold=True, color="FFFFFF")
    alt_row_fill = PatternFill(start_color="E6F3F5", end_color="E6F3F5", fill_type="solid")
    border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    for col_idx, col_name in enumerate(df.columns, 1):
        cell = worksheet.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
        worksheet.column_dimensions[get_column_letter(col_idx)].width = max(20, len(str(col_name)) + 5)
    
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
    """
    [المحرك المطور]: تحليل الأرقام المركبة والسلاسل الطويلة والمجموعات 
    مع تصفية المسافات لضمان عدم سقوط أي جروب (مثل 14028 * 2)
    """
    sku = str(sku).strip().replace(" ", "")
    if not sku or sku == "nan" or sku == "":
        return None, None, None, [], False
        
    if '*' in sku or '-' in sku:
        if '*' in sku and '-' not in sku:
            match = re.match(r'^(\d+)\*(\d+)$', sku)
            if match:
                base_sku = match.group(1)
                qty = int(match.group(2))
                return base_sku, sku, qty, [base_sku], True
        
        parts = sku.split('-')
        individual_skus = []
        for p in parts:
            if '*' in p:
                m = re.match(r'^(\d+)\*(\d+)$', p)
                if m: individual_skus.append(m.group(1))
            else:
                m = re.match(r'^(\d+)$', p)
                if m: individual_skus.append(p)
        if individual_skus:
            return individual_skus[0], sku, None, individual_skus, True

    return sku, None, None, [sku], False

def extract_offer_name(text):
    """استخراج اسم العرض التفصيلي من النصوص"""
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
        match = re.search(r'صفقة اليوم\s*:\s*([^:]+?)(?=\s*(?:اذا اشترى|نسبة من|يبدأ بتاريخ|$))', text)
        if match: return f"صفقة اليوم : {match.group(1).strip()}"
        return "صفقة اليوم"
    if "حبة بسعر" in text or "حبات بسعر" in text:
        match = re.search(r'(\d+حبات? بسعر [\d.]+ ريال)', text)
        return match.group(1) if match else "عرض كميات"
    return "عرض خاص"

def extract_dates(text):
    """استخراج النطاق الزمني للعروض"""
    text = str(text) if not pd.isna(text) else ""
    date_match = re.findall(r'(\d{4}-\d{2}-\d{2} \d{2}:\d{2})', text)
    start = date_match[0] if len(date_match) > 0 else None
    end = date_match[1] if len(date_match) > 1 else None
    return start, end

def extract_numbers_from_text(text):
    """عزل الأرقام الفردية للمنتجات بعد استبعاد السلاسل المركبة والمجموعات"""
    if pd.isna(text):
        return []
    text = str(text)
    text_clean = re.sub(r'\d{3,6}(?:-\d{3,6})+', '', text)
    text_clean = re.sub(r'\d{3,6}\s*\*\s*\d+', '', text_clean)
    
    excluded_years = {'2024', '2025', '2026', '2027', '2028', '2029', '2030'}
    pattern = r'(?:^|[^0-9])(\d{3,6})(?:[^0-9]|$)'
    matches = re.findall(pattern, text_clean)
    return [m for m in matches if m not in excluded_years and not m.startswith('20')]

@st.cache_data(show_spinner=False)
def generate_excel_download_files(df):
    """توليد ملفات إكسيل التحميل بكفاءة عالية مستندة للذاكرة المؤقتة للبيانات الضخمة"""
    simple_output = BytesIO()
    with pd.ExcelWriter(simple_output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="العروض والمنتجات", index=False)
        apply_excel_style(writer, "العروض والمنتجات", df)
    simple_bytes = simple_output.getvalue()

    detailed_output = BytesIO()
    with pd.ExcelWriter(detailed_output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="النتيجة النهائية", index=False)
        apply_excel_style(writer, "النتيجة النهائية", df)
        
        if "اسم العرض الخاص" in df:
            unique_offers = df["اسم العرض الخاص"].dropna().unique()
            for offer_type in unique_offers[:12]:
                if offer_type and offer_type.strip() != "":
                    type_df = df[df["اسم العرض الخاص"] == offer_type]
                    if len(type_df) > 0:
                        sheet_name = str(offer_type)[:30].replace("|", "-").replace(":", "-")
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
            border-radius: 24px; padding: 1.5rem; color: white;
            margin-bottom: 1.5rem; text-align: center;
        }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="offers-header">
        <h1>🛍/ مركز معالجة وإدارة عروض المتجر</h1>
        <p>تفكيك السلاسل المركبة الشاملة، ربط الأسعار والقنوات الترويجية الموازية، وتأمين جداول التسويات</p>
    </div>
    """, unsafe_allow_html=True)

    st.subheader("📂 رفع ملف التقرير الشامل")
    uploaded_file = st.file_uploader("قم برفع ملف المبيعات والعروض المشترك لبلسم العلا", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
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
        
        # تنظيف شيت الأسعار القياسية من الترويسات العشوائية
        if not df_regular_prices.empty:
            if not ('رمز المنتج' in str(df_regular_prices.columns) or 'sku' in str(df_regular_prices.columns).lower()):
                if df_regular_prices.iloc[0].astype(str).str.contains('رمز المنتج|sku', case=False, na=False).any():
                    df_regular_prices.columns = df_regular_prices.iloc[0]
                    df_regular_prices = df_regular_prices[1:].reset_index(drop=True)

        st.subheader("🔧 مطابقة وتأكيد أعمدة النظام")
        col1, col2, col3 = st.columns(3)
        with col1: sku_col_choice = st.selectbox("عمود معرف المنتج (SKU) - شيت الأسعار", options=list(df_regular_prices.columns) if not df_regular_prices.empty else ["لا يوجد"], index=0)
        with col2: name_col_choice = st.selectbox("عمود اسم المنتج - شيت الأسعار", options=list(df_regular_prices.columns) if not df_regular_prices.empty else ["لا يوجد"], index=1 if len(df_regular_prices.columns) > 1 else 0)
        with col3: price_col_choice = st.selectbox("عمود السعر القياسي - شيت الأسعار", options=list(df_regular_prices.columns) if not df_regular_prices.empty else ["لا يوجد"], index=2 if len(df_regular_prices.columns) > 2 else 0)
        
        price_map = {}
        if not df_regular_prices.empty and sku_col_choice != "لا يوجد":
            for _, row in df_regular_prices.iterrows():
                sku = str(row[sku_col_choice]).strip()
                if sku and sku != "nan":
                    price_map[sku] = {
                        "name": row[name_col_choice] if name_col_choice != "لا يوجد" and pd.notna(row[name_col_choice]) else "",
                        "price": row[price_col_choice] if price_col_choice != "لا يوجد" and pd.notna(row[price_col_choice]) else ""
                    }
        
        # 🧠 [الفصل التام لقنوات التخفيضات]: عزل الفردي عن المجموعات لمنع الاختلاط
        individual_discount_map = {}
        group_discount_map = {}
        sku_master = {} # ديل ريادي مركزي يجمع كل الأصناف الفريدة المكتشفة بالملف
        
        if not df_discounted.empty:
            st.subheader("🔧 إعدادات شيت الأسعار المخفضة")
            c1, c2, c3, c4 = st.columns(4)
            with c1: disc_sku_col = st.selectbox("عمود معرف المنتج (SKU)", options=list(df_discounted.columns), index=0)
            with c2: disc_price_col = st.selectbox("عمود السعر المخفض الحقيقي", options=list(df_discounted.columns), index=2 if len(df_discounted.columns) > 2 else 0)
            with c3: disc_end_col = st.selectbox("عمود نهاية التخفيض", options=["لا يوجد"] + list(df_discounted.columns), index=4 if len(df_discounted.columns) > 4 else 0)
            with c4: disc_promo_col = st.selectbox("عمود الترويجات للخصم", options=["لا يوجد"] + list(df_discounted.columns), index=5 if len(df_discounted.columns) > 5 else 0)
            
            for _, row in df_discounted.iterrows():
                sku_raw = str(row[disc_sku_col]).strip()
                if not sku_raw or sku_raw == "nan": continue
                
                base_sku, group_sku, group_qty, individual_skus, is_composite = parse_composite_sku(sku_raw)
                price_val = row[disc_price_col] if pd.notna(row[disc_price_col]) else ""
                end_val = row[disc_end_col] if disc_end_col != "لا يوجد" and pd.notna(row[disc_end_col]) else ""
                promo_val = row[disc_promo_col] if disc_promo_col != "لا يوجد" and pd.notna(row[disc_promo_col]) else ""
                
                disc_payload = {"discounted_price": price_val, "end_date": end_val, "promo_title": promo_val}
                
                if is_composite:
                    group_discount_map[group_sku] = disc_payload
                    for ind_sku in individual_skus:
                        if ind_sku not in sku_master: sku_master[ind_sku] = {"groups": set(), "offers": []}
                        sku_master[ind_sku]["groups"].add(group_sku)
                else:
                    individual_discount_map[base_sku] = disc_payload
                    if base_sku not in sku_master: sku_master[base_sku] = {"groups": set(), "offers": []}
        
        # 3. معالجة وتفجير تقارير شيت العروض النصية
        with st.spinner("🧠 جاري تشغيل المحرك المطور وفك تداخل عروض الأصناف المشتركة..."):
            offers_col = df_raw.columns[0]
            
            for idx, row in df_raw.iterrows():
                text = str(row[offers_col]).strip()
                if not text or text == "nan": continue
                
                offer_name = extract_offer_name(text)
                start_date, end_date = extract_dates(text)
                is_daily_deal = "صفقة اليوم" in text
                offer_payload = {"name": offer_name, "start": start_date, "end": end_date, "is_daily_deal": is_daily_deal}
                
                # التقاط عروض الـ Star Groups
                groups_star = re.findall(r'(\d{3,6})\s*\*\s*(\d+)', text)
                for base_sku, qty in groups_star:
                    group_sku = f"{base_sku}*{qty.strip()}"
                    if base_sku not in sku_master: sku_master[base_sku] = {"groups": set(), "offers": []}
                    sku_master[base_sku]["groups"].add(group_sku)
                    sku_master[base_sku]["offers"].append(offer_payload)
                
                # التقاط عروض السلاسل الكبرى الممتدة بالشرطة لأي طول (مثل صبغات كوليستون)
                groups_dash = re.findall(r'(\d{3,6}(?:-\d{3,6})+)', text)
                for d_seq in groups_dash:
                    individual_skus = d_seq.split('-')
                    for sku in individual_skus:
                        if sku not in sku_master: sku_master[sku] = {"groups": set(), "offers": []}
                        sku_master[sku]["groups"].add(d_seq)
                        sku_master[sku]["offers"].append(offer_payload)
                
                # التقاط العروض الموجهة للأصناف الفردية
                numbers = extract_numbers_from_text(text)
                for sku in numbers:
                    if sku not in sku_master: sku_master[sku] = {"groups": set(), "offers": []}
                    sku_master[sku]["offers"].append(offer_payload)
            
            # 4. بناء الهيكل السحابي الموحد وتعبئة كافة المصفوفات بدقة متناهية
            final_results = []
            
            for base_sku in sorted(sku_master.keys()):
                data = sku_master[base_sku]
                product_info = price_map.get(base_sku, {"name": "", "price": ""})
                
                # سحب بيانات الخصم الفردي النقي
                ind_disc = individual_discount_map.get(base_sku, {})
                disc_price_item = ind_disc.get("discounted_price", "")
                disc_promo_item = ind_disc.get("promo_title", "")
                disc_end_item = ind_disc.get("end_date", "")
                
                # تجميع العروض المترابطة
                unique_offers = {o["name"]: o for o in data["offers"]}
                offer_names = " | ".join([o["name"] for o in unique_offers.values()])
                offer_starts = " | ".join([o["start"] for o in unique_offers.values() if o["start"]])
                offer_ends = " | ".join([o["end"] for o in unique_offers.values() if o["end"]])
                is_daily = "نعم" if any(o["is_daily_deal"] for o in unique_offers.values()) else ""
                
                # 🧠 [محرك الدمج المتوازي العبقري]: صب المجموعات المتعددة للصنف في قنوات متوازية
                g_skus_list, g_names_list, g_prices_list, g_disc_prices_list, g_qtys_list, g_promos_list, g_ends_list = [], [], [], [], [], [], []
                
                for g_sku in sorted(list(data["groups"])):
                    g_skus_list.append(g_sku)
                    
                    # محاولة جلب السعر والاسم للمجموعة مباشرة
                    g_info = price_map.get(g_sku, {})
                    g_name = g_info.get("name", "")
                    g_price = g_info.get("price", "")
                    
                    if not g_name:
                        base_m = re.match(r'^(\d+)', g_sku)
                        if base_m:
                            b_info = price_map.get(base_m.group(1), {})
                            g_name = b_info.get("name", "")
                            g_price = b_info.get("price", "")
                    
                    g_names_list.append(str(g_name))
                    g_prices_list.append(str(g_price))
                    
                    # استخلاص خصومات المجموعات من قنواتها الصحيحة
                    g_disc = group_discount_map.get(g_sku, {})
                    g_disc_prices_list.append(str(g_disc.get("discounted_price", "")))
                    g_promos_list.append(str(g_disc.get("promo_title", "")))
                    g_ends_list.append(str(g_disc.get("end_date", "")))
                    
                    # جلب عدد حبات المجموعة للـ Star Group
                    qty_val = ""
                    if '*' in g_sku and '-' not in g_sku:
                        q_match = re.search(r'\*(\d+)', g_sku)
                        if q_match: qty_val = q_match.group(1)
                    g_qtys_list.append(qty_val)
                
                # الدمج النهائي للجروبات باستخدام الأنبوب الفاصل (Pipe) في حال تعددها للمنتج الواحد
                group_sku_str = " | ".join(g_skus_list) if g_skus_list else ""
                group_name_str = " | ".join(g_names_list) if any(x != "" for x in g_names_list) else ""
                group_price_str = " | ".join(g_prices_list) if any(x != "" for x in g_prices_list) else ""
                group_disc_price_str = " | ".join(g_disc_prices_list) if any(x != "" for x in g_disc_prices_list) else ""
                group_qty_str = " | ".join(g_qtys_list) if any(x != "" for x in g_qtys_list) else ""
                group_promo_str = " | ".join(g_promos_list) if any(x != "" for x in g_promos_list) else ""
                group_end_str = " | ".join(g_ends_list) if any(x != "" for x in g_ends_list) else ""
                
                final_results.append({
                    "رقم المنتج": base_sku, "رقم المنتج للمجموعة": group_sku_str,
                    "اسم المنتج": product_info["name"], "سعر المنتج": product_info["price"],
                    "اسم المنتج للمجموعة": group_name_str, "سعر المنتج للمجموعة": group_price_str,
                    "اسم العرض الخاص": offer_names, "بداية العرض": offer_starts, "نهاية العرض": offer_ends,
                    "سعر مخفض للمنتج": disc_price_item, "سعر مخفض للمجموعة": group_disc_price_str, "عدد حبات المجموعة": group_qty_str,
                    "العنوان الترويجي للمنتج": disc_promo_item, "العنوان الترويجي للمجموعة": group_promo_str,
                    "تاريخ نهاية التخفيض للمنتج": disc_end_item, "تاريخ نهاية التخفيض للمجموعة": group_end_str,
                    "صفقة اليوم": is_daily
                })
            
            df_final = pd.DataFrame(final_results)
            df_final = df_final[df_final["رقم المنتج"].notna() & (df_final["رقم المنتج"] != "nan") & (df_final["رقم المنتج"] != "")]
            df_final = df_final[~df_final["رقم المنتج"].isin(["2024", "2025", "2026", "2027", "2028", "2029", "2030"])]
            df_final = df_final.drop_duplicates(subset=["رقم المنتج", "رقم المنتج للمجموعة"], keep="first")
            
            # ترتيب الأعمدة النهائي والمثالي متضمناً قنوات التواريخ المستردة كاملة
            column_order = [
                "رقم المنتج", "رقم المنتج للمجموعة", "اسم المنتج", "سعر المنتج", "اسم المنتج للمجموعة", "سعر المنتج للمجموعة",
                "اسم العرض الخاص", "بداية العرض", "نهاية العرض", "سعر مخفض للمنتج", "سعر مخفض للمجموعة", "عدد حبات المجموعة",
                "العنوان الترويجي للمنتج", "العنوان الترويجي للمجموعة", "تاريخ نهاية التخفيض للمنتج", "تاريخ نهاية التخفيض للمجموعة", "صفقة اليوم"
            ]
            df_final = df_final[[col for col in column_order if col in df_final.columns]]
            
        st.success(f"✅ [نجاح مطلق]: تم إنجاز تسويات العروض حياً، وحصر كافة القنوات الموازية للأصناف فريداً!")
        
        # استعراض بطاقات الأداء اللحظية المحصنة
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1: st.metric("📦 إجمالي السلع المدرجة", f"{len(df_final):,}")
        with col2: st.metric("🏷/ سلع مربوطة بعروض", f"{(df_final['اسم العرض الخاص'].astype(str).str.strip().str.len().gt(0)).sum():,}")
        with col3: st.metric("💰 تخفيضات الأسعار المفردة", f"{(df_final['سعر مخفض للمنتج'].astype(str).str.strip().str.len().gt(0)).sum():,}")
        with col4: st.metric("🎯 مجموعات تم حصرها وشحنها", f"{(df_final['رقم المنتج للمجموعة'].astype(str).str.strip().str.len().gt(0)).sum():,}")
        with col5: st.metric("📊 أنواع العروض", f"{df_final['اسم العرض الخاص'].nunique():,}")
        
        # محرك الاستعلام البحثي المباشر
        st.subheader("🔍 استعلام وبحث سريع في قاعدة التسويات")
        search_term = st.text_input("أدخل رقم صنف أو اسم المنتج للمعاين اللحظية للمجموعات المتكاملة:", placeholder="مثال: 9969")
        if search_term:
            search_term = search_term.strip()
            filtered_df = df_final[(df_final["رقم المنتج"].astype(str) == search_term) | (df_final["رقم المنتج للمجموعة"].astype(str).str.contains(search_term)) | (df_final["اسم المنتج"].str.contains(search_term, na=False))]
            if not filtered_df.empty: st.dataframe(filtered_df, use_container_width=True)
            else: st.warning(f"⚠️ لم يتم العثور على أي عروض مطابقة للمعيار: '{search_term}'")
            
        st.subheader("📋 شاشة العرض والمراقبة الشاملة للجدول المركزي المطور")
        st.dataframe(df_final, use_container_width=True, height=400)
        
        # أزرار تنزيل التقارير السحابية الفورية
        st.subheader("💾 استخراج وحفظ التقارير الحصينة")
        col_dl1, col_dl2 = st.columns(2)
        simple_bytes, detailed_bytes = generate_excel_download_files(df_final)
        
        with col_dl1:
            st.download_button(label="📥 استخراج الملف الشامل (توزيع شيتات العروض تلقائياً)", data=detailed_bytes, file_name="تقرير_العروض_المفصل_المطور.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with col_dl2:
            st.download_button(label="📥 استخراج الملف الموحد (جدول المطابقة المركزي المحمي)", data=simple_bytes, file_name="العروض_والمنتجات_الموحد_المطور.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            
        if "اسم العرض الخاص" in df_final and len(df_final["اسم العرض الخاص"].dropna().unique()) > 0:
            st.subheader("📊 تصفية فرز مخصصة بنوع الخصم")
            offer_types = ["الكل"] + sorted(df_final["اسم العرض الخاص"].dropna().unique().tolist())
            selected_type = st.selectbox("اختر نوع المعاملة المالية لفرزها:", offer_types)
            if selected_type != "الكل": st.dataframe(df_final[df_final["اسم العرض الخاص"] == selected_type], use_container_width=True)
    else:
        st.info("📂 بانتظار رفع ملف العروض المحدث لتنشيط المعالجة الفورية واستعراض القنوات المتوازية.")
