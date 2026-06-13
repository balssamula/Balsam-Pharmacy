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

def extract_sku_from_text(text):
    """استخراج أرقام المنتجات الفردية من النص (تجاهل السنوات)"""
    if pd.isna(text):
        return []
    text = str(text)
    excluded = {'2024','2025','2026','2027','2028','2029','2030'}
    # البحث عن أرقام مكونة من 4-6 أرقام تسبقها أو تليها مسافة أو شرطة
    pattern = r'(?:^|[-\s/]+)(\d{4,6})(?:[-\s/]|$)'
    matches = re.findall(pattern, text)
    return [m for m in matches if m not in excluded]

def extract_composite_skus(text):
    """استخراج الصيغ المركبة مثل 1500*6 أو 15000-16000-12000"""
    if pd.isna(text):
        return []
    text = str(text)
    composites = []
    # بحث عن مجموعات الضرب
    star_matches = re.findall(r'(\d{4,6})\*(\d+)', text)
    for base, qty in star_matches:
        composites.append({
            'type': 'star',
            'base': base,
            'full': f"{base}*{qty}",
            'qty': int(qty)
        })
    # بحث عن المجموعات المتعددة بالشرطة (3 أرقام)
    dash_matches = re.findall(r'(\d{4,6})-(\d{4,6})-(\d{4,6})', text)
    for a,b,c in dash_matches:
        composites.append({
            'type': 'dash',
            'base': a,  # نأخذ أول رقم كقاعدة (يمكن أن يكون أي)
            'full': f"{a}-{b}-{c}",
            'members': [a,b,c],
            'qty': None
        })
    return composites

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
            df_offers_raw = pd.read_excel(uploaded_file, sheet_name=offers_sheet, header=None)
            df_discounted = pd.read_excel(uploaded_file, sheet_name=discounted_sheet) if discounted_sheet else pd.DataFrame()
            df_prices = pd.read_excel(uploaded_file, sheet_name=prices_sheet) if prices_sheet else pd.DataFrame()
            
            # ========== تحديد الأعمدة في شيت اسعار المنتجات ==========
            st.subheader("🔧 تحديد أعمدة شيت 'اسعار المنتجات'")
            col1, col2, col3 = st.columns(3)
            with col1:
                sku_col = st.selectbox("عمود رقم المنتج", options=df_prices.columns, index=0 if len(df_prices.columns)>0 else None)
            with col2:
                name_col = st.selectbox("عمود اسم المنتج", options=df_prices.columns, index=min(1, len(df_prices.columns)-1) if len(df_prices.columns)>1 else 0)
            with col3:
                price_col = st.selectbox("عمود السعر العادي", options=df_prices.columns, index=min(2, len(df_prices.columns)-1) if len(df_prices.columns)>2 else 0)
            
            # بناء قاموس الأسعار العادية
            price_map = {}
            if not df_prices.empty:
                for _, row in df_prices.iterrows():
                    sku = str(row[sku_col]).strip()
                    if sku and sku != "nan":
                        price_map[sku] = {
                            "name": row[name_col] if name_col else "",
                            "price": row[price_col] if price_col else ""
                        }
            
            # ========== معالجة الأسعار المخفضة ==========
            # أولاً: تحديد أعمدة شيت سعر مخفض
            st.subheader("🔧 تحديد أعمدة شيت 'سعر مخفض'")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                disc_sku_col = st.selectbox("عمود رقم المنتج", options=df_discounted.columns, index=0 if len(df_discounted.columns)>0 else None)
            with col2:
                disc_price_col = st.selectbox("عمود السعر المخفض", options=df_discounted.columns, index=min(1, len(df_discounted.columns)-1) if len(df_discounted.columns)>1 else 0)
            with col3:
                disc_end_col = st.selectbox("عمود تاريخ النهاية (اختياري)", options=["لا يوجد"] + list(df_discounted.columns), index=0)
            with col4:
                disc_promo_col = st.selectbox("عمود العنوان الترويجي (اختياري)", options=["لا يوجد"] + list(df_discounted.columns), index=0)
            
            discounted_map = {}  # {sku: {discounted_price, promo_title, end_date}}
            if not df_discounted.empty:
                for _, row in df_discounted.iterrows():
                    sku = str(row[disc_sku_col]).strip()
                    if not sku or sku == "nan":
                        continue
                    price_val = row[disc_price_col] if disc_price_col and pd.notna(row[disc_price_col]) else ""
                    end_val = row[disc_end_col] if disc_end_col != "لا يوجد" and pd.notna(row[disc_end_col]) else ""
                    promo_val = row[disc_promo_col] if disc_promo_col != "لا يوجد" and pd.notna(row[disc_promo_col]) else ""
                    discounted_map[sku] = {
                        "discounted_price": price_val,
                        "end_date": end_val,
                        "promo_title": promo_val
                    }
            
            # ========== معالجة العروض ==========
            # هيكل لتخزين البيانات: المفتاح هو المنتج الأساسي (base_sku)
            # وكل منتج يحمل قائمة عروض ومعلومات المجموعة
            products = {}  # base_sku -> {offers: [], group_sku: None, group_qty: None, group_members: []}
            
            for idx, row in df_offers_raw.iterrows():
                text = row.iloc[0] if len(row) > 0 else ""
                if pd.isna(text) or text == "nan":
                    continue
                text = str(text)
                
                # استخراج الأرقام الفردية
                single_skus = extract_sku_from_text(text)
                # استخراج الصيغ المركبة
                composites = extract_composite_skus(text)
                
                offer_name = extract_offer_name(text)
                start_date, end_date = extract_dates(text)
                is_daily = "صفقة اليوم" in text
                
                # معالجة الأرقام الفردية
                for sku in single_skus:
                    if sku not in products:
                        products[sku] = {
                            "offers": [],
                            "group_sku": None,
                            "group_qty": None,
                            "group_members": []
                        }
                    products[sku]["offers"].append({
                        "name": offer_name,
                        "start": start_date,
                        "end": end_date,
                        "is_daily": is_daily
                    })
                
                # معالجة الصيغ المركبة
                for comp in composites:
                    if comp['type'] == 'star':
                        base = comp['base']
                        full = comp['full']
                        qty = comp['qty']
                        if base not in products:
                            products[base] = {
                                "offers": [],
                                "group_sku": None,
                                "group_qty": None,
                                "group_members": []
                            }
                        # تسجيل معلومات المجموعة للمنتج الأساسي
                        products[base]["group_sku"] = full
                        products[base]["group_qty"] = qty
                        products[base]["offers"].append({
                            "name": offer_name,
                            "start": start_date,
                            "end": end_date,
                            "is_daily": is_daily
                        })
                    elif comp['type'] == 'dash':
                        members = comp['members']
                        full = comp['full']
                        # لكل عضو في المجموعة، نضيف نفس معلومات المجموعة
                        for member in members:
                            if member not in products:
                                products[member] = {
                                    "offers": [],
                                    "group_sku": None,
                                    "group_qty": None,
                                    "group_members": []
                                }
                            products[member]["group_sku"] = full
                            products[member]["group_members"] = members
                            products[member]["offers"].append({
                                "name": offer_name,
                                "start": start_date,
                                "end": end_date,
                                "is_daily": is_daily
                            })
            
            # ========== بناء النتيجة النهائية ==========
            result_rows = []
            for base_sku, data in products.items():
                # تجميع العروض (قد تتكرر لوجود أكثر من عرض لنفس المنتج)
                offers_list = data["offers"]
                # دمج العروض المتطابقة (نفس الاسم)
                unique_offers = {}
                for offer in offers_list:
                    key = offer["name"]
                    if key not in unique_offers:
                        unique_offers[key] = offer
                
                offer_names = " | ".join([o["name"] for o in unique_offers.values()])
                offer_starts = " | ".join([o["start"] for o in unique_offers.values() if o["start"]])
                offer_ends = " | ".join([o["end"] for o in unique_offers.values() if o["end"]])
                is_daily_val = "نعم" if any(o["is_daily"] for o in unique_offers.values()) else ""
                
                # معلومات المنتج العادي
                prod_info = price_map.get(base_sku, {"name": "", "price": ""})
                product_name = prod_info["name"]
                product_price = prod_info["price"]
                
                # معلومات المجموعة
                group_sku = data.get("group_sku")
                group_qty = data.get("group_qty")
                group_members = data.get("group_members", [])
                
                # اسم المجموعة وسعرها
                group_name = ""
                group_price = ""
                if group_sku:
                    group_info = price_map.get(group_sku, {})
                    group_name = group_info.get("name", "")
                    group_price = group_info.get("price", "")
                    if not group_name and group_members:
                        # إذا كان لدينا أعضاء المجموعة، قد نأخذ اسم أول عضو
                        first_member = group_members[0] if group_members else None
                        if first_member:
                            group_name = price_map.get(first_member, {}).get("name", "")
                
                # السعر المخفض والعنوان الترويجي للمنتج العادي
                disc_info = discounted_map.get(base_sku, {})
                disc_price = disc_info.get("discounted_price", "")
                disc_promo = disc_info.get("promo_title", "")
                disc_end = disc_info.get("end_date", "")
                
                # السعر المخفض والعنوان الترويجي للمجموعة
                group_disc_info = discounted_map.get(group_sku, {}) if group_sku else {}
                group_disc_price = group_disc_info.get("discounted_price", "")
                group_disc_promo = group_disc_info.get("promo_title", "")
                
                result_rows.append({
                    "رقم المنتج": base_sku,
                    "رقم المنتج للمجموعة": group_sku if group_sku else "",
                    "اسم المنتج": product_name,
                    "سعر المنتج": product_price,
                    "اسم المنتج للمجموعة": group_name,
                    "سعر المنتج للمجموعة": group_price,
                    "اسم العرض الخاص": offer_names,
                    "بداية العرض": offer_starts,
                    "نهاية العرض": offer_ends,
                    "سعر مخفض للمنتج": disc_price,
                    "سعر مخفض للمجموعة": group_disc_price,
                    "عدد حبات المجموعة": group_qty if group_qty else "",
                    "العنوان الترويجي للمنتج": disc_promo,
                    "العنوان الترويجي للمجموعة": group_disc_promo,
                    "صفقة اليوم": is_daily_val
                })
            
            # إضافة المنتجات التي لها سعر مخفض فقط (لم تظهر في العروض)
            for sku, disc_info in discounted_map.items():
                if sku not in products:
                    # تجنب إضافة السنوات
                    if sku in ['2024','2025','2026','2027','2028','2029','2030']:
                        continue
                    prod_info = price_map.get(sku, {"name": "", "price": ""})
                    result_rows.append({
                        "رقم المنتج": sku,
                        "رقم المنتج للمجموعة": "",
                        "اسم المنتج": prod_info["name"],
                        "سعر المنتج": prod_info["price"],
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
            
            df_final = pd.DataFrame(result_rows)
            
            # إزالة التكرارات (نفس رقم المنتج قد يظهر أكثر من مرة بسبب عروض متعددة، لكننا دمجناها)
            # التأكد من عدم وجود صفوف مكررة بالكامل
            df_final = df_final.drop_duplicates(subset=["رقم المنتج", "رقم المنتج للمجموعة"], keep="first")
            
            # تنظيف إضافي: إزالة صفوف السنوات التي قد تبقى
            df_final = df_final[~df_final["رقم المنتج"].isin(["2024","2025","2026","2027","2028","2029","2030"])]
            
            # عرض الإحصائيات
            st.success(f"✅ تم معالجة {len(df_final)} منتج بنجاح (بدون تكرار)")
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("📦 إجمالي المنتجات", len(df_final))
            with col2:
                st.metric("🏷️ مع عروض", df_final["اسم العرض الخاص"].str.len().gt(0).sum())
            with col3:
                st.metric("🏷️ منتجات مخفضة", df_final["سعر مخفض للمنتج"].notna().sum())
            with col4:
                st.metric("🎯 مجموعات", df_final["رقم المنتج للمجموعة"].str.len().gt(0).sum())
            
            st.subheader("📋 قائمة المنتجات والعروض")
            st.dataframe(df_final, use_container_width=True, height=400)
            
            # تصفية برقم المنتج
            st.subheader("🔍 تصفية برقم المنتج")
            search_term = st.text_input("أدخل رقم المنتج (فردي أو مجموعة):")
            if search_term:
                filtered = df_final[(df_final["رقم المنتج"] == search_term) | (df_final["رقم المنتج للمجموعة"] == search_term)]
                st.dataframe(filtered, use_container_width=True)
            
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
                                if offer_type and offer_type != "":
                                    type_df = df[df["اسم العرض الخاص"] == offer_type]
                                    if len(type_df) > 0:
                                        sheet_name = offer_type[:31]
                                        type_df.to_excel(writer, sheet_name=sheet_name, index=False)
                                        apply_excel_style(writer, sheet_name, type_df)
                    return output.getvalue()
                
                st.download_button(
                    label="📥 تحميل ملف مفصل (مع شيتات)",
                    data=create_detailed_excel(df_final),
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
                
                st.download_button(
                    label="📥 تحميل ملف مبسط (جدول واحد)",
                    data=create_simple_excel(df_final),
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
                    st.dataframe(df_final[df_final["اسم العرض الخاص"] == selected_type], use_container_width=True)
    else:
        st.info("📂 يرجى رفع ملف Excel لعرض العروض")
        with st.expander("ℹ️ تعليمات"):
            st.markdown("""
            **كيفية استخدام هذه الصفحة:**
            1. قم برفع ملف Excel الذي يحتوي على أوراق: 'عرض خاص'، 'سعر مخفض'، 'اسعار المنتجات'
            2. حدد الأعمدة الصحيحة لكل شيت من القوائم المنسدلة
            3. سيتم استخراج العروض والأسعار تلقائياً
            4. يمكنك التصفية برقم المنتج أو نوع العرض
            5. تحميل النتائج كملف Excel مفصل أو مبسط
            """)
