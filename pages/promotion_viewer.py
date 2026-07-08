# pages/promotion_viewer.py
import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
import re

# ========== دوال التنسيق ==========
def apply_excel_style(writer, sheet_name, df):
    """تطبيق التنسيقات الاحترافية على ملف Excel"""
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

@st.cache_data
def generate_empty_template():
    """توليد النموذج الفارغ"""
    output = BytesIO()
    wb = openpyxl.Workbook()
    
    header_fill = PatternFill(start_color="1F7A8C", end_color="1F7A8C", fill_type="solid")
    font_headers = Font(name="Tajawal", size=11, bold=True, color="FFFFFF")
    align_center = Alignment(horizontal="center", vertical="center")
    border_thin = Border(
        left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1')
    )
    
    ws1 = wb.active
    ws1.title = "عرض خاص"
    ws1.views.sheetView[0].rightToLeft = True
    cell_1 = ws1.cell(row=1, column=1, value="  العروض الخاصة  ")
    cell_1.fill = header_fill
    cell_1.font = font_headers
    cell_1.alignment = align_center
    cell_1.border = border_thin
    ws1.column_dimensions['A'].width = 60
    ws1.row_dimensions[1].height = 24
    
    ws2 = wb.create_sheet(title="سعر مخفض")
    ws2.views.sheetView[0].rightToLeft = True
    headers_discounted = ["رمز المنتج sku", "أسم المنتج", "السعر المخفض", "تاريخ نهاية التخفيض", "العنوان الترويجي"]
    for col_idx, text in enumerate(headers_discounted, 1):
        cell = ws2.cell(row=1, column=col_idx, value=text)
        cell.fill = header_fill
        cell.font = font_headers
        cell.alignment = align_center
        cell.border = border_thin
        ws2.column_dimensions[get_column_letter(col_idx)].width = 24
    ws2.row_dimensions[1].height = 24
        
    ws3 = wb.create_sheet(title="اسعار المنتجات")
    ws3.views.sheetView[0].rightToLeft = True
    headers_prices = ["رمز المنتج sku", "أسم المنتج", "سعر المنتج"]
    for col_idx, text in enumerate(headers_prices, 1):
        cell = ws3.cell(row=2, column=col_idx, value=text)
        cell.fill = header_fill
        cell.font = font_headers
        cell.alignment = align_center
        cell.border = border_thin
        ws3.column_dimensions[get_column_letter(col_idx)].width = 24
    ws3.row_dimensions[1].height = 18
    ws3.row_dimensions[2].height = 24
        
    wb.save(output)
    return output.getvalue()

# ========== دوال تحليل العروض ==========
def extract_offer_name(text):
    """استخراج اسم العرض التفصيلي مع دعم جميع الصيغ"""
    text = str(text) if not pd.isna(text) else ""
    
    # 1️⃣ خصم على القطعة الثانية
    if "خصم" in text and "القطعة الثانية" in text:
        match = re.search(r'خصم\s*(\d+)\s*%\s*على\s*القطعة\s*الثانية', text)
        if match: return f"خصم {match.group(1)}% على القطعة الثانية"
        match = re.search(r'خصم\s*(\d+)\s*%\s*على\s*القطعة', text)
        if match: return f"خصم {match.group(1)}% على القطعة الثانية"
        return "خصم على القطعة الثانية"
    
    # 2️⃣ خصم على الحبة الثانية
    if "خصم" in text and "الحبة الثانية" in text:
        match = re.search(r'خصم\s*(\d+)\s*%\s*على\s*الحبة\s*الثانية', text)
        if match: return f"خصم {match.group(1)}% على الحبة الثانية"
        match = re.search(r'خصم\s*(\d+)\s*%\s*على\s*الحبة', text)
        if match: return f"خصم {match.group(1)}% على الحبة الثانية"
        match = re.search(r'خصم\s*(\d+)\s*ريال\s*على\s*الحبة\s*الثانية', text)
        if match: return f"خصم {match.group(1)} ريال على الحبة الثانية"
        return "خصم على الحبة الثانية"
    
    # 3️⃣ عروض مجانية (2+1 مجاناً، 2 +1 مجانا، إلخ)
    if "مجاناً" in text or "مجانا" in text:
        match = re.search(r'(\d+)\s*\+\s*(\d+)\s*مجاناً?', text)
        if match: return f"عرض {match.group(1)}+{match.group(2)} مجاناً"
        match = re.search(r'(\d+)\s*\+\s*(\d+)\s*مجانا', text)
        if match: return f"عرض {match.group(1)}+{match.group(2)} مجاناً"
        # حالة خاصة: "1+1 مجانا" أو "2+1 مجانا" بدون كلمة عرض
        match = re.search(r'(\d+)\s*\+\s*(\d+)\s*مجاناً?', text)
        if match: return f"عرض {match.group(1)}+{match.group(2)} مجاناً"
        return "عرض مجاني"
    
    # 4️⃣ صفقة اليوم
    if "صفقة اليوم" in text:
        match = re.search(r'صفقة اليوم\s*:\s*([^:]+?)(?=\s*(?:اذا اشترى|نسبة من|يبدأ بتاريخ|$))', text)
        if match:
            full_text = match.group(1).strip()
            if len(full_text) > 50:
                full_text = full_text[:50] + "..."
            return f"صفقة اليوم : {full_text}"
        return "صفقة اليوم"
    
    # 5️⃣ عروض الكميات (6حبات بسعر 77 ريال)
    if "حبة بسعر" in text or "حبات بسعر" in text:
        match = re.search(r'(\d+)\s*حبات?\s*بسعر\s*([\d.]+)\s*ريال', text)
        if match: return f"{match.group(1)}حبات بسعر {match.group(2)} ريال"
        return "عرض كميات"
    
    # 6️⃣ عرض خاص - خصم
    if "عرض خاص" in text and "خصم" in text:
        match = re.search(r'عرض خاص\s*-\s*خصم\s*(\d+)\s*%\s*على\s*الحبة\s*الثانية', text)
        if match: return f"عرض خاص - خصم {match.group(1)}% على الحبة الثانية"
        return "عرض خاص"
    
    # 7️⃣ صيغة "2حبة بسعر 65.50 ريال" (بدون مسافة)
    match = re.search(r'(\d+)\s*حبة\s*بسعر\s*([\d.]+)\s*ريال', text)
    if match: return f"{match.group(1)}حبة بسعر {match.group(2)} ريال"
    
    # 8️⃣ حالة خاصة: أرقام فقط بدون عرض واضح
    if "/" in text:
        parts = text.split("/")
        if len(parts) >= 2:
            potential_name = parts[0].strip()
            if re.search(r'[^\d\s\-]', potential_name):
                return potential_name
    
    return "عرض خاص"

def extract_discount_percentage(text, original_price=None, discounted_price=None, promo_text=None):
    """
    استخراج نسبة الخصم من مصادر متعددة:
    1. من النص المباشر (خصم 50%)
    2. من النص الترويجي (12حبة ب29.95 ريال)
    3. من مقارنة السعر الأصلي مع السعر المخفض
    4. من قيمة الخصم بالريال (خصم 17 ريال)
    """
    text = str(text) if not pd.isna(text) else ""
    
    # 1️⃣ خصم بنسبة مئوية مباشرة
    match = re.search(r'خصم\s*(\d+)\s*%', text)
    if match:
        return f"{match.group(1)}%"
    
    # 2️⃣ إذا كان هناك نص ترويجي منفصل
    if promo_text and isinstance(promo_text, str):
        # 2أ: البحث عن نسبة مئوية في النص الترويجي
        match = re.search(r'خصم\s*(\d+)\s*%', promo_text)
        if match:
            return f"{match.group(1)}%"
        
        # 2ب: البحث عن سعر محدد في النص الترويجي (مثل 12حبة ب29.95 ريال)
        promo_price = extract_price_from_promo(promo_text)
        if promo_price and original_price:
            try:
                original = float(str(original_price).replace(',', '').strip())
                if original > 0 and promo_price < original:
                    discount_amount = original - promo_price
                    percentage = (discount_amount / original) * 100
                    return f"{round(percentage, 0)}%"
            except (ValueError, TypeError):
                pass
        
        # 2ج: البحث عن قيمة خصم بالريال (مثل خصم 17 ريال)
        discount_value = extract_discount_value_from_promo(promo_text)
        if discount_value and original_price:
            try:
                original = float(str(original_price).replace(',', '').strip())
                if original > 0:
                    percentage = (discount_value / original) * 100
                    return f"{round(percentage, 0)}%"
            except (ValueError, TypeError):
                pass
    
    # 3️⃣ حساب النسبة من السعر الأصلي والمخفض
    if original_price and discounted_price:
        try:
            original = float(str(original_price).replace(',', '').strip())
            discounted = float(str(discounted_price).replace(',', '').strip())
            if original > 0 and discounted > 0 and original > discounted:
                discount_amount = original - discounted
                percentage = (discount_amount / original) * 100
                return f"{round(percentage, 0)}%"
        except (ValueError, TypeError):
            pass
    
    # 4️⃣ عروض مجانية (2+1 مجاناً)
    match = re.search(r'(\d+)\s*\+\s*(\d+)\s*مجاناً?', text)
    if match:
        free_qty = int(match.group(2))
        total_qty = int(match.group(1)) + free_qty
        percentage = (free_qty / total_qty) * 100
        return f"{round(percentage, 0)}%"
    
    # 5️⃣ صيغة "2 +1 مجانا"
    match = re.search(r'(\d+)\s*\+\s*(\d+)\s*مجانا', text)
    if match:
        free_qty = int(match.group(2))
        total_qty = int(match.group(1)) + free_qty
        percentage = (free_qty / total_qty) * 100
        return f"{round(percentage, 0)}%"
    
    return ""

def extract_offer_quantity(text):
    """استخراج عدد حبات العرض من نص العرض"""
    text = str(text) if not pd.isna(text) else ""
    
    # 1️⃣ عروض مجانية (2+1 مجاناً)
    match = re.search(r'(\d+)\s*\+\s*(\d+)\s*مجاناً?', text)
    if match:
        paid = int(match.group(1))
        free = int(match.group(2))
        return f"{paid + free}"
    
    # 2️⃣ صيغة "2 +1 مجانا"
    match = re.search(r'(\d+)\s*\+\s*(\d+)\s*مجانا', text)
    if match:
        paid = int(match.group(1))
        free = int(match.group(2))
        return f"{paid + free}"
    
    # 3️⃣ عروض الكميات (6حبات بسعر 77 ريال)
    match = re.search(r'(\d+)\s*حبات?\s*بسعر', text)
    if match:
        return match.group(1)
    
    # 4️⃣ خصم على الحبة الثانية
    if "الحبة الثانية" in text:
        return "2"
    
    # 5️⃣ خصم على القطعة الثانية
    if "القطعة الثانية" in text:
        return "2"
    
    return ""

def extract_special_offer_sku(text):
    """استخراج رقم المنتج الخاص بالعرض (المجموعة) من النص"""
    text = str(text) if not pd.isna(text) else ""
    
    # البحث عن صيغة "رقم-رقم-رقم" (مثل 5209-17164-4846)
    match = re.search(r'(\d{3,6}-\d{3,6}-\d{3,6})', text)
    if match:
        return match.group(1)
    
    # البحث عن صيغة "رقم*رقم" (مثل 16265*6)
    match = re.search(r'(\d{3,6}\*\d+)', text)
    if match:
        return match.group(1)
    
    return ""

def extract_dates(text):
    """استخراج النطاق الزمني للعروض"""
    text = str(text) if not pd.isna(text) else ""
    date_match = re.findall(r'(\d{4}-\d{2}-\d{2} \d{2}:\d{2})', text)
    start = date_match[0] if len(date_match) > 0 else None
    end = date_match[1] if len(date_match) > 1 else None
    return start, end

def extract_numbers_from_text(text):
    """عزل الأرقام الفردية للمنتجات بعد استبعاد السلاسل المركبة"""
    if pd.isna(text):
        return []
    text = str(text)
    text_clean = re.sub(r'\d{3,6}(?:-\d{3,6})+', '', text)
    text_clean = re.sub(r'\d{3,6}\s*\*\s*\d+', '', text_clean)
    
    excluded_years = {'2024', '2025', '2026', '2027', '2028', '2029', '2030'}
    pattern = r'(?:^|[^0-9])(\d{3,6})(?:[^0-9]|$)'
    matches = re.findall(pattern, text_clean)
    return [m for m in matches if m not in excluded_years and not m.startswith('20')]

def parse_composite_sku(sku):
    """تحليل الأرقام المركبة والسلاسل الطويلة"""
    sku = str(sku).strip().replace(" ", "")
    if not sku or sku == "nan" or sku == "":
        return None, None, None, [], False, False
        
    if '-' in sku:
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
            return individual_skus[0], sku, None, individual_skus, False, True
            
    if '*' in sku:
        match = re.match(r'^(\d+)\*(\d+)$', sku)
        if match:
            base_sku = match.group(1)
            qty = int(match.group(2))
            return base_sku, sku, qty, [base_sku], True, False

    return sku, None, None, [sku], False, False

# ========== دوال التصدير ==========
@st.cache_data(show_spinner=False)
def generate_excel_download_files(df):
    """توليد ملفات إكسيل التحميل"""
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

# ========== دالة العرض الرئيسية ==========
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
        <h1>🛍️ مركز معالجة وإدارة عروض المتجر</h1>
        <p>فصل أوتوماتيكي لسلاسل العروض الخاصة، تفعيل كشافات البحث الفوري، وتأمين الترويجات والتواريخ</p>
    </div>
    """, unsafe_allow_html=True)

    st.subheader("📂 رفع ملف التقرير الشامل")
    
    col_upload, col_template = st.columns([3, 1])
    with col_upload:
        uploaded_file = st.file_uploader("قم برفع ملف المبيعات والعروض المشترك", type=["xlsx", "xls"])
    with col_template:
        st.markdown("<div style='text-align: center; padding-top: 1.6rem;'>", unsafe_allow_html=True)
        template_bytes = generate_empty_template()
        st.download_button(
            label="📥 نموذج ملف العروض",
            data=template_bytes,
            file_name="نموذج_جميع_عروض_المتجر_الفعالة.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.markdown("</div>", unsafe_allow_html=True)
    
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
        
        # ========== معالجة الأسعار ==========
        if not df_regular_prices.empty:
            sku_col_idx, name_col_idx, price_col_idx = 0, 1, 2
            for i, c in enumerate(df_regular_prices.columns):
                c_str = str(c).lower()
                if 'sku' in c_str or 'رمز' in c_str: sku_col_idx = i
                elif 'اسم' in c_str or 'أسم' in c_str: name_col_idx = i
                elif 'سعر' in c_str: price_col_idx = i
            
            sku_col_choice = df_regular_prices.columns[sku_col_idx]
            name_col_choice = df_regular_prices.columns[name_col_idx]
            price_col_choice = df_regular_prices.columns[price_col_idx]
            
            price_map = {}
            for _, row in df_regular_prices.iterrows():
                sku = str(row[sku_col_choice]).strip()
                if sku and sku != "nan":
                    price_map[sku] = {
                        "name": row[name_col_choice] if pd.notna(row[name_col_choice]) else "",
                        "price": row[price_col_choice] if pd.notna(row[price_col_choice]) else ""
                    }
        else:
            price_map = {}
        
        # ========== معالجة الأسعار المخفضة ==========
        individual_discount_map = {}
        group_discount_map = {}
        sku_master = {}

        if not df_discounted.empty:
            # ... (تحديد الأعمدة كما هو)
    
            for _, row in df_discounted.iterrows():
                sku_raw = str(row[disc_sku_col]).strip()
                if not sku_raw or sku_raw == "nan": continue
        
                base_sku, group_sku, group_qty, individual_skus, is_star, is_dash = parse_composite_sku(sku_raw)
                price_val = row[disc_price_col] if pd.notna(row[disc_price_col]) else ""
                end_val = row[disc_end_col] if disc_end_col and pd.notna(row[disc_end_col]) else ""
                promo_val = row[disc_promo_col] if disc_promo_col and pd.notna(row[disc_promo_col]) else ""
        
                # 🆕 حساب نسبة الخصم من السعر الأصلي والمخفض
                original_price = price_map.get(base_sku, {}).get("price", "")
                discount_percentage = extract_discount_percentage(
                    text=promo_val,
                    original_price=original_price,
                    discounted_price=price_val,
                    promo_text=promo_val
                )
        
                disc_payload = {
                    "discounted_price": price_val,
                    "end_date": end_val,
                    "promo_title": promo_val,
                    "discount_percentage": discount_percentage  # 🆕 إضافة النسبة المحسوبة
                }
        
                if is_star or is_dash:
                    group_discount_map[group_sku] = disc_payload
                    for ind_sku in individual_skus:
                        if ind_sku not in sku_master:
                            sku_master[ind_sku] = {"groups": set(), "special_offer_skus": set(), "offers": []}
                        if is_star: sku_master[ind_sku]["groups"].add(group_sku)
                        if is_dash: sku_master[ind_sku]["special_offer_skus"].add(group_sku)
                else:
                    individual_discount_map[base_sku] = disc_payload
                    if base_sku not in sku_master:
                        sku_master[base_sku] = {"groups": set(), "special_offer_skus": set(), "offers": []}
        
        # ========== معالجة العروض ==========
        with st.spinner("🧠 جاري معالجة العروض..."):
            offers_col = df_raw.columns[0]
            
            for idx, row in df_raw.iterrows():
                text = str(row[offers_col]).strip()
                if not text or text == "nan": continue
                
                # استخراج البيانات الأساسية
                offer_name = extract_offer_name(text)
                start_date, end_date = extract_dates(text)
                is_daily_deal = "صفقة اليوم" in text
                discount_percentage = extract_discount_percentage(text)
                offer_quantity = extract_offer_quantity(text)
                special_offer_sku = extract_special_offer_sku(text)
                
                offer_payload = {
                    "name": offer_name,
                    "start": start_date,
                    "end": end_date,
                    "is_daily_deal": is_daily_deal,
                    "discount_percentage": discount_percentage,
                    "offer_quantity": offer_quantity,
                    "special_offer_sku": special_offer_sku
                }
                
                # استخراج المجموعات (مثل 16265*6)
                groups_star = re.findall(r'(\d{3,6})\s*\*\s*(\d+)', text)
                for base_sku, qty in groups_star:
                    group_sku = f"{base_sku}*{qty.strip()}"
                    if base_sku not in sku_master:
                        sku_master[base_sku] = {"groups": set(), "special_offer_skus": set(), "offers": []}
                    sku_master[base_sku]["groups"].add(group_sku)
                    sku_master[base_sku]["offers"].append(offer_payload)
                
                # استخراج السلاسل المفصولة بـ -
                groups_dash = re.findall(r'(\d{3,6}(?:-\d{3,6})+)', text)
                for d_seq in groups_dash:
                    individual_skus = d_seq.split('-')
                    for sku in individual_skus:
                        if sku not in sku_master:
                            sku_master[sku] = {"groups": set(), "special_offer_skus": set(), "offers": []}
                        sku_master[sku]["special_offer_skus"].add(d_seq)
                        sku_master[sku]["offers"].append(offer_payload)
                
                # استخراج الأرقام الفردية
                numbers = extract_numbers_from_text(text)
                for sku in numbers:
                    if sku not in sku_master:
                        sku_master[sku] = {"groups": set(), "special_offer_skus": set(), "offers": []}
                    sku_master[sku]["offers"].append(offer_payload)
            
            # ========== بناء النتيجة النهائية ==========
            final_results = []
            
            for base_sku in sorted(sku_master.keys()):
                data = sku_master[base_sku]
                product_info = price_map.get(base_sku, {"name": "", "price": ""})
                
                ind_disc = individual_discount_map.get(base_sku, {})
                disc_price_item = ind_disc.get("discounted_price", "")
                disc_promo_item = ind_disc.get("promo_title", "")
                disc_end_item = ind_disc.get("end_date", "")
                
                # دمج العروض المتعددة
                unique_offers = {}
                for o in data["offers"]:
                    key = o["name"]
                    if key not in unique_offers:
                        unique_offers[key] = o
                
                offer_names = " | ".join([o["name"] for o in unique_offers.values()])
                offer_starts = " | ".join([o["start"] for o in unique_offers.values() if o["start"]])
                offer_ends = " | ".join([o["end"] for o in unique_offers.values() if o["end"]])
                is_daily = "نعم" if any(o["is_daily_deal"] for o in unique_offers.values()) else ""
                
                # دمج نسب الخصم وعدد الحبات
                discount_percentages = [o.get("discount_percentage", "") for o in unique_offers.values() if o.get("discount_percentage")]
                offer_quantities = [o.get("offer_quantity", "") for o in unique_offers.values() if o.get("offer_quantity")]
                special_offer_skus = [o.get("special_offer_sku", "") for o in unique_offers.values() if o.get("special_offer_sku")]
                
                discount_percentage_str = " | ".join(set([x for x in discount_percentages if x]))
                offer_quantity_str = " | ".join(set([x for x in offer_quantities if x]))
                special_offer_sku_str = " | ".join(set([x for x in special_offer_skus if x]))
                
                # معالجة المجموعات
                g_skus_list = []
                g_names_list = []
                g_prices_list = []
                g_qtys_list = []
                g_disc_prices_list = []
                g_promos_list = []
                g_ends_list = []
                
                for g_sku in sorted(list(data["groups"])):
                    g_skus_list.append(g_sku)
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
                    
                    qty_val = ""
                    q_match = re.search(r'\*(\d+)', g_sku)
                    if q_match: qty_val = q_match.group(1)
                    g_qtys_list.append(qty_val)
                    
                    g_disc = group_discount_map.get(g_sku, {})
                    g_disc_prices_list.append(str(g_disc.get("discounted_price", "")))
                    g_promos_list.append(str(g_disc.get("promo_title", "")))
                    g_ends_list.append(str(g_disc.get("end_date", "")))
                
                # إضافة عروض خاصة
                for sp_sku in sorted(list(data["special_offer_skus"])):
                    g_disc = group_discount_map.get(sp_sku, {})
                    if g_disc:
                        g_disc_prices_list.append(str(g_disc.get("discounted_price", "")))
                        g_promos_list.append(str(g_disc.get("promo_title", "")))
                        g_ends_list.append(str(g_disc.get("end_date", "")))
                
                group_sku_str = " | ".join(g_skus_list) if g_skus_list else ""
                group_name_str = " | ".join(g_names_list) if any(x != "" for x in g_names_list) else ""
                group_price_str = " | ".join(g_prices_list) if any(x != "" for x in g_prices_list) else ""
                group_disc_price_str = " | ".join([x for x in g_disc_prices_list if x != ""]) if g_disc_prices_list else ""
                group_qty_str = " | ".join(g_qtys_list) if any(x != "" for x in g_qtys_list) else ""
                group_promo_str = " | ".join([x for x in g_promos_list if x != ""]) if g_promos_list else ""
                group_end_str = " | ".join([x for x in g_ends_list if x != ""]) if g_ends_list else ""
                
                final_results.append({
                    "رقم المنتج": base_sku,
                    "رقم المنتج للمجموعة": group_sku_str,
                    "رقم منتج العرض الخاص": special_offer_sku_str,
                    "اسم المنتج": product_info["name"],
                    "سعر المنتج": product_info["price"],
                    "اسم المنتج للمجموعة": group_name_str,
                    "سعر المنتج للمجموعة": group_price_str,
                    "اسم العرض الخاص": offer_names,
                    "نسبة الخصم": discount_percentage_str,
                    "عدد حبات العرض": offer_quantity_str,
                    "بداية العرض": offer_starts,
                    "نهاية العرض": offer_ends,
                    "سعر مخفض للمنتج": disc_price_item,
                    "سعر مخفض للمجموعة": group_disc_price_str,
                    "عدد حبات المجموعة": group_qty_str,
                    "العنوان الترويجي للمنتج": disc_promo_item,
                    "العنوان الترويجي للمجموعة": group_promo_str,
                    "تاريخ نهاية التخفيض للمنتج": disc_end_item,
                    "تاريخ نهاية التخفيض للمجموعة": group_end_str,
                    "صفقة اليوم": is_daily
                })
            
            df_final = pd.DataFrame(final_results)
            df_final = df_final[df_final["رقم المنتج"].notna() & (df_final["رقم المنتج"] != "nan") & (df_final["رقم المنتج"] != "")]
            df_final = df_final[~df_final["رقم المنتج"].isin(["2024", "2025", "2026", "2027", "2028", "2029", "2030"])]
            df_final = df_final.drop_duplicates(subset=["رقم المنتج", "رقم المنتج للمجموعة", "رقم منتج العرض الخاص"], keep="first")
            
            column_order = [
                "رقم المنتج", "رقم المنتج للمجموعة", "رقم منتج العرض الخاص",
                "اسم المنتج", "سعر المنتج", "اسم المنتج للمجموعة", "سعر المنتج للمجموعة",
                "اسم العرض الخاص", "نسبة الخصم", "عدد حبات العرض",
                "بداية العرض", "نهاية العرض",
                "سعر مخفض للمنتج", "سعر مخفض للمجموعة", "عدد حبات المجموعة",
                "العنوان الترويجي للمنتج", "العنوان الترويجي للمجموعة",
                "تاريخ نهاية التخفيض للمنتج", "تاريخ نهاية التخفيض للمجموعة",
                "صفقة اليوم"
            ]
            df_final = df_final[[col for col in column_order if col in df_final.columns]]
            
        st.success(f"✅ تم معالجة {len(df_final)} منتج بنجاح")
        
        # عرض الإحصائيات
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1: st.metric("📦 إجمالي المنتجات", f"{len(df_final):,}")
        with col2: st.metric("🏷️ مع عروض", f"{(df_final['اسم العرض الخاص'].astype(str).str.strip().str.len().gt(0)).sum():,}")
        with col3: st.metric("💰 نسبة الخصم", f"{df_final['نسبة الخصم'].astype(str).str.strip().str.len().gt(0).sum():,}")
        with col4: st.metric("🎯 عدد حبات العرض", f"{df_final['عدد حبات العرض'].astype(str).str.strip().str.len().gt(0).sum():,}")
        with col5: st.metric("📊 أنواع العروض", f"{df_final['اسم العرض الخاص'].nunique():,}")
        
        # ========== تصفية ==========
        st.subheader("🔍 استعلام وبحث سريع")
        search_term = st.text_input("أدخل رقم صنف، رقم مجموعة، أو سلسلة العرض الخاص للبحث:", placeholder="مثال: 5929")
        if search_term:
            search_term = search_term.strip()
            filtered_df = df_final[
                (df_final["رقم المنتج"].astype(str) == search_term) |
                (df_final["رقم المنتج للمجموعة"].astype(str).str.contains(search_term, na=False, regex=False)) |
                (df_final["رقم منتج العرض الخاص"].astype(str).str.contains(search_term, na=False, regex=False)) |
                (df_final["اسم المنتج"].str.contains(search_term, na=False, regex=False))
            ]
            if not filtered_df.empty:
                st.dataframe(filtered_df, use_container_width=True)
            else:
                st.warning(f"⚠️ لم يتم العثور على نتائج مطابقة لـ: '{search_term}'")
        
        # ========== عرض الجدول ==========
        st.subheader("📋 شاشة العرض والمراقبة الشاملة")
        st.dataframe(df_final, use_container_width=True, height=400)
        
        # ========== خيارات التحميل ==========
        st.subheader("💾 استخراج وحفظ التقارير")
        col_dl1, col_dl2 = st.columns(2)
        simple_bytes, detailed_bytes = generate_excel_download_files(df_final)
        
        with col_dl1:
            st.download_button(
                label="📥 تحميل الملف المفصل (مع شيتات منفصلة)",
                data=detailed_bytes,
                file_name="تقرير_العروض_المفصل.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with col_dl2:
            st.download_button(
                label="📥 تحميل الملف المبسط (جدول واحد)",
                data=simple_bytes,
                file_name="العروض_والمنتجات_الموحد.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        # ========== تصفية حسب نوع العرض ==========
        if "اسم العرض الخاص" in df_final and len(df_final["اسم العرض الخاص"].dropna().unique()) > 0:
            st.subheader("📊 تصفية حسب نوع العرض")
            offer_types = ["الكل"] + sorted(df_final["اسم العرض الخاص"].dropna().unique().tolist())
            selected_type = st.selectbox("اختر نوع العرض:", offer_types)
            if selected_type != "الكل":
                st.dataframe(df_final[df_final["اسم العرض الخاص"] == selected_type], use_container_width=True)
    else:
        st.info("📂 يرجى رفع ملف Excel لبدء المعالجة")

def calculate_discount_percentage(original_price, discounted_price, promo_text=""):
    """
    حساب نسبة الخصم بناءً على السعر الأصلي والسعر المخفض والنص الترويجي
    """
    try:
        original = float(str(original_price).replace(',', '').strip()) if original_price else 0
        discounted = float(str(discounted_price).replace(',', '').strip()) if discounted_price else 0
    except (ValueError, TypeError):
        return ""
    
    if original <= 0 or discounted <= 0:
        return ""
    
    # 1️⃣ إذا كان النص الترويجي يحتوي على نسبة مئوية مباشرة
    if promo_text and isinstance(promo_text, str):
        match = re.search(r'خصم\s*(\d+)\s*%', promo_text)
        if match:
            return f"{match.group(1)}%"
    
    # 2️⃣ حساب النسبة من السعرين
    if original > discounted:
        discount_amount = original - discounted
        percentage = (discount_amount / original) * 100
        return f"{round(percentage, 0)}%"
    
    return ""

def extract_price_from_promo(promo_text):
    """
    استخراج السعر من النص الترويجي (مثل 12حبة ب29.95 ريال)
    """
    if not promo_text or not isinstance(promo_text, str):
        return None
    
    # البحث عن صيغ مختلفة
    patterns = [
        r'(\d+)\s*حبات?\s*ب\s*([\d.]+)\s*ريال',
        r'(\d+)\s*حبة\s*بسعر\s*([\d.]+)\s*ريال',
        r'(\d+)\s*حبات?\s*بسعر\s*([\d.]+)\s*ريال',
        r'بسعر\s*([\d.]+)\s*ريال',
        r'ب\s*([\d.]+)\s*ريال'
    ]
    
    for pattern in patterns:
        match = re.search(pattern, promo_text)
        if match:
            # إذا كان هناك مجموعتان (عدد + سعر) أو مجموعة واحدة (سعر فقط)
            if len(match.groups()) >= 2:
                return float(match.group(2))
            else:
                return float(match.group(1))
    
    return None

def extract_discount_value_from_promo(promo_text):
    """
    استخراج قيمة الخصم بالريال من النص الترويجي (مثل خصم 17 ريال)
    """
    if not promo_text or not isinstance(promo_text, str):
        return None
    
    match = re.search(r'خصم\s*([\d.]+)\s*ريال', promo_text)
    if match:
        return float(match.group(1))
    
    return None
