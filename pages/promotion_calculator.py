import streamlit as st
import pandas as pd
import re
from io import BytesIO

def create_template_excel():
    """إنشاء ملف نموذج فارغ يحتوي على الأعمدة المطلوبة فقط لمساعد المدير على تعبئته"""
    template_data = {
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
    output.seek(0)
    return output.getvalue()

def calculate_promotions_stream(uploaded_file):
    """حساب العروض الترويجية مع حماية كاملة ضد أخطاء القسمة على صفر"""
    df = pd.read_excel(uploaded_file)
    
    # التأكد من توفر الأعمدة الأساسية وتنبيه المستخدم بشكل ودي
    required_cols = ["عنوان العرض", "السعر غير شامل الضريبة", "الضريبة"]
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"العمود المطلوبة '{col}' غير موجود في الملف المرفوع!")

    df["عنوان العرض"] = df["عنوان العرض"].astype(str).str.strip()

    total_qty_list, discount_pct_list, price_before_list, discount_val_list, price_after_list = [], [], [], [], []

    for idx, row in df.iterrows():
        p_base = float(row.get("السعر غير شامل الضريبة", 0))
        tax_pct = float(row.get("الضريبة", 0))
        promo_text = row["عنوان العرض"]
        
        # السعر الفردي شامل الضريبة للحبة الواحدة
        p_tax = p_base * (1 + tax_pct)

        qty, disc_pct, p_before, p_after = 1, 0.0, p_tax, p_tax

        # النمط 1: عروض المجاني (مثال: عرض 1+1 مجاناً)
        match_free = re.search(r"عرض\s*(\d+)\s*\+\s*(\d+)\s*مجانا", promo_text)
        match_free_alt = re.search(r"عرض\s*(\d+)\s*حبة\s*\+\s*(\d+)\s*مجانا", promo_text)

        if match_free or match_free_alt:
            m = match_free if match_free else match_free_alt
            buy_qty, free_qty = int(m.group(1)), int(m.group(2))
            qty = buy_qty + free_qty
            p_before = p_tax * qty
            p_after = p_tax * buy_qty
            # حماية ضد القسمة على صفر في إجمالي الكمية
            disc_pct = (free_qty / qty) if qty > 0 else 0.0
            
        # النمط 2: خصم نسبة على الحبة الثانية
        elif "على الحبة الثانية" in promo_text and "%" in promo_text:
            try:
                pct_val = float(re.search(r"خصم\s*(\d+)%", promo_text).group(1)) / 100
                qty = 2
                p_before = p_tax * 2
                p_after = p_tax + (p_tax * (1 - pct_val))
                # حماية ضد القسمة على صفر في السعر قبل العرض
                disc_pct = ((p_before - p_after) / p_before) if p_before > 0 else 0.0
            except: pass
            
        # النمط 3: خصم نسبة مئوية مباشرة على الإجمالي
        elif "خصم" in promo_text and "%" in promo_text:
            try:
                pct_val = float(re.search(r"خصم\s*(\d+)%", promo_text).group(1)) / 100
                qty = 1
                p_before = p_tax
                p_after = p_tax * (1 - pct_val)
                disc_pct = pct_val
            except: pass
            
        # النمط 4: عدد حبات بسعر ثابت
        elif "بسعر" in promo_text:
            qty_match = re.search(r"(\d+)\s*حبة", promo_text)
            price_match = re.search(r"بسعر\s*([\d\.]+)", promo_text)
            if qty_match and price_match:
                qty = int(qty_match.group(1))
                p_after = float(price_match.group(1))
                p_before = p_tax * qty
                # حماية ضد القسمة على صفر في السعر قبل العرض
                disc_pct = ((p_before - p_after) / p_before) if p_before > 0 else 0.0

        disc_val = p_before - p_after
        
        # إضافة النتائج للقوائم مع تقريبها لخانة عشرية ثنائية
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
    
    # قسم تحميل النموذج الجاهز
    st.markdown("### 📋 الخطوة 1: تحميل نموذج البيانات وتعبئته")
    st.info("لتفادي الأخطاء، قم بتحميل هذا النموذج المفرغ وتعبئة البيانات داخله بالصيغة الصحيحة:")
    
    template_bytes = create_template_excel()
    st.download_button(
        label="📥 تحميل نموذج ملف العروض الفارغ (Excel Template)",
        data=template_bytes,
        file_name="Promotions_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    
    st.markdown("---")
    
    # قسم رفع الملف بعد التعبئة
    st.markdown("### 📂 الخطوة 2: رفع ملف العروض المكتمل")
    uploaded_file = st.file_uploader("اختر ملف إكسيل العروض المدخل الأساسي بعد ملئه", type=["xlsx"])
    
    if uploaded_file is not None:
        with st.spinner("⏳ جاري تحليل وهندسة العروض الرياضية الذكية وتأمين الحسابات..."):
            try:
                excel_bytes, result_df = calculate_promotions_stream(uploaded_file)
                st.success("✅ تمت معالجة وتفكيك معادلات العروض بأمان ودون أي تقسيم صِفري!")
                
                st.markdown("### 📊 معاينة سريعة للنتائج المحسوبة (أول 10 صفوف):")
                st.dataframe(result_df.head(10), use_container_width=True)
                
                st.download_button(
                    label="💾 تحميل ملف العروض المحسوب بالكامل Excel",
                    data=excel_bytes,
                    file_name="Calculated_Promotions_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
            except ValueError as ve:
                st.error(f"⚠️ {str(ve)}")
            except Exception as e:
                st.error(f"❌ حدث خطأ أثناء المعالجة: {str(e)}")
