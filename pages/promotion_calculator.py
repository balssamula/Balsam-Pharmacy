import streamlit as st
import pandas as pd
import re
from io import BytesIO

def calculate_promotions_stream(uploaded_file):
    """حساب العروض الترويجية من الملف المرفوع مباشرة وإرجاع حزمة الـ Bytes للتنزيل"""
    df = pd.read_excel(uploaded_file)
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
            disc_pct = free_qty / qty
        elif "على الحبة الثانية" in promo_text and "%" in promo_text:
            try:
                pct_val = float(re.search(r"خصم\s*(\d+)%", promo_text).group(1)) / 100
                qty = 2
                p_before = p_tax * 2
                p_after = p_tax + (p_tax * (1 - pct_val))
                disc_pct = (p_before - p_after) / p_before
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
                disc_pct = (p_before - p_after) / p_before

        disc_val = p_before - p_after
        total_qty_list.append(qty)
        discount_pct_list.append(round(disc_pct * 100, 2))
        price_before_list.append(round(p_before, 2))
        discount_val_list.append(round(disc_val, 2))
        price_after_list.append(round(p_after, 2))

    df["عدد حبات العرض"] = total_qty_list
    df["نسبة الخصم الإجمالية %"] = discount_pct_list
    df["السعر قبل العرض شامل الضريبة"] = price_before_list
    df["قيمة الخصم شامل الضريبة"] = discount_val_list
    df["السعر بعد العرض شامل الضريبة"] = price_after_list

    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)
    return output.getvalue(), df

def show():
    st.markdown("<div class='hero'><h1>🏷️ حاسبة العروض والخصومات الترويجية</h1><p>ارفع ملف إكسيل العروض لحساب التفاصيل الإجمالية ونسب الخصم والأسعار قبل وبعد العرض تلقائياً</p></div>", unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader("📂 اختر ملف إكسيل العروض المدخل الأساسي", type=["xlsx"])
    if uploaded_file is not None:
        with st.spinner("⏳ جاري تحليل وهندسة العروض الرياضية الذكية..."):
            try:
                excel_bytes, result_df = calculate_promotions_stream(uploaded_file)
                st.success("✅ تمت معالجة وتفكيك معادلات العروض بنجاح!")
                
                st.markdown("### 📊 معاينة سريعة للنتائج المحسوبة:")
                st.dataframe(result_df.head(10), use_container_width=True)
                
                st.download_button(
                    label="💾 تحميل ملف العروض المحسوب بالكامل Excel",
                    data=excel_bytes,
                    file_name="Calculated_Promotions_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"❌ حدث خطأ أثناء المعالجة، تأكد من مطابقة أسماء الأعمدة: {str(e)}")
