# pages/promotion_viewer.py
import streamlit as st
import pandas as pd
import re
from io import BytesIO

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
        .offer-title {
            font-size: 1.1rem;
            font-weight: 800;
            color: #16425b;
        }
        .offer-dates {
            font-size: 0.8rem;
            color: #607783;
            margin-top: 0.3rem;
        }
        .offer-badge {
            display: inline-block;
            background: #dff1ff;
            color: #0f5488;
            border-radius: 999px;
            padding: 0.2rem 0.7rem;
            font-size: 0.7rem;
            font-weight: 700;
            margin: 0.2rem;
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
    st.subheader("📂 رفع ملف العروض")
    
    uploaded_file = st.file_uploader(
        "قم برفع ملف Excel الذي يحتوي على العروض",
        type=["xlsx", "xls"],
        help="الملف يجب أن يحتوي على ورقة باسم 'عرض خاص' أو 'العروض'"
    )
    
    if uploaded_file is not None:
        with st.spinner("جاري معالجة العروض..."):
            # قراءة الملف
            excel_data = pd.ExcelFile(uploaded_file)
            sheet_names = excel_data.sheet_names
            
            # البحث عن ورقة العروض
            offers_sheet = None
            for sheet in sheet_names:
                if "عرض خاص" in sheet or "عروض" in sheet:
                    offers_sheet = sheet
                    break
            
            if not offers_sheet:
                offers_sheet = sheet_names[0]
            
            df_raw = pd.read_excel(uploaded_file, sheet_name=offers_sheet)
            
            # معالجة العروض
            all_offers = []
            
            for idx, row in df_raw.iterrows():
                text = str(row.iloc[0]) if len(row) > 0 else ""
                if pd.isna(text) or text == "nan":
                    continue
                
                # استخراج الأرقام
                numbers_match = re.findall(r'(\d{4,6})(?:[-/\s]|$)', text)
                numbers = numbers_match if numbers_match else []
                
                # استخراج اسم العرض
                offer_name = ""
                if "خصم" in text:
                    if "القطعة الثانية" in text:
                        match = re.search(r'(خصم \d+% على القطعة الثانية)', text)
                        offer_name = match.group(1) if match else "خصم على القطعة الثانية"
                    elif "الحبة الثانية" in text:
                        match = re.search(r'(خصم \d+% على الحبة الثانية)', text)
                        offer_name = match.group(1) if match else "خصم على الحبة الثانية"
                    elif "ريال" in text:
                        match = re.search(r'(خصم \d+ ريال على الحبة الثانية)', text)
                        offer_name = match.group(1) if match else "خصم بقيمة محددة"
                elif "عرض" in text and "مجاناً" in text:
                    match = re.search(r'(عرض \d+\+\d+ مجاناً?)', text)
                    offer_name = match.group(1) if match else "عرض مجاني"
                elif "صفقة اليوم" in text:
                    offer_name = "صفقة اليوم"
                elif "حبة بسعر" in text or "حبات بسعر" in text:
                    match = re.search(r'(\d+حبات? بسعر [\d.]+ ريال)', text)
                    offer_name = match.group(1) if match else "عرض كميات"
                else:
                    offer_name = "عرض خاص"
                
                # استخراج التواريخ
                start_date = None
                end_date = None
                date_match = re.findall(r'(\d{4}-\d{2}-\d{2} \d{2}:\d{2})', text)
                if len(date_match) >= 1:
                    start_date = date_match[0]
                if len(date_match) >= 2:
                    end_date = date_match[1]
                
                # إضافة كل رقم على حدة
                for num in numbers:
                    all_offers.append({
                        "رقم الصنف": num,
                        "اسم العرض": offer_name,
                        "بداية العرض": start_date,
                        "نهاية العرض": end_date
                    })
            
            df_offers = pd.DataFrame(all_offers)
            
            st.success(f"✅ تم معالجة {len(df_offers)} عرض بنجاح")
            
            # عرض الإحصائيات
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("📦 إجمالي العروض", len(df_offers))
            with col2:
                st.metric("🏷️ أنواع العروض", df_offers["اسم العرض"].nunique())
            with col3:
                st.metric("🔢 أصناف مشمولة", df_offers["رقم الصنف"].nunique())
            
            # فلترة حسب نوع العرض
            st.subheader("🔍 تصفية حسب نوع العرض")
            offer_types = ["الكل"] + sorted(df_offers["اسم العرض"].dropna().unique().tolist())
            selected_type = st.selectbox("اختر نوع العرض", offer_types)
            
            if selected_type != "الكل":
                filtered_df = df_offers[df_offers["اسم العرض"] == selected_type]
            else:
                filtered_df = df_offers
            
            # عرض الجدول
            st.subheader("📋 قائمة العروض المفصلة")
            st.dataframe(filtered_df, use_container_width=True, height=400)
            
            # عرض العروض على شكل بطاقات
            st.subheader("🎴 عرض البطاقات")
            show_cards = st.checkbox("عرض على شكل بطاقات", value=False)
            
            if show_cards:
                for _, offer in filtered_df.iterrows():
                    with st.container():
                        st.markdown(f"""
                        <div class="offer-card">
                            <div class="offer-title">🆔 {offer['رقم الصنف']}</div>
                            <div><strong>🏷️ {offer['اسم العرض']}</strong></div>
                            <div class="offer-dates">
                                📅 بداية: {offer['بداية العرض'] if offer['بداية العرض'] else 'غير محدد'} 
                                | 🔚 نهاية: {offer['نهاية العرض'] if offer['نهاية العرض'] else 'غير محدد'}
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
            
            # تحميل الملف
            st.subheader("💾 تحميل النتيجة")
            
            def to_excel(df):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name="العروض المفصلة", index=False)
                    
                    # ورقة منفصلة لكل نوع عرض
                    for offer_type in df["اسم العرض"].dropna().unique():
                        if offer_type:
                            type_df = df[df["اسم العرض"] == offer_type]
                            sheet_name = offer_type[:31]
                            type_df.to_excel(writer, sheet_name=sheet_name, index=False)
                return output.getvalue()
            
            excel_data = to_excel(filtered_df)
            st.download_button(
                label="📥 تحميل ملف Excel",
                data=excel_data,
                file_name="العروض_المفصلة.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # عرض الأصناف المكررة
            duplicates = filtered_df[filtered_df["رقم الصنف"].duplicated(keep=False)]
            if len(duplicates) > 0:
                with st.expander(f"⚠️ الأصناف المكررة ({len(duplicates)} عرض)"):
                    st.dataframe(duplicates.sort_values("رقم الصنف"), use_container_width=True)
                    
    else:
        st.info("📂 يرجى رفع ملف Excel لعرض العروض")
        
        with st.expander("ℹ️ تعليمات"):
            st.markdown("""
            **كيفية استخدام هذه الصفحة:**
            1. قم برفع ملف Excel الذي يحتوي على العروض
            2. يجب أن يحتوي الملف على ورقة باسم "عرض خاص" أو "العروض"
            3. سيتم فك العروض تلقائيًا (تقسيم الأرقام المتعددة)
            4. يمكنك تصفية العروض حسب النوع
            5. يمكنك تحميل النتيجة كملف Excel
            
            **صيغة الملف المدعومة:**
            - الأرقام المفصولة بـ `-` أو `/` أو مسافة
            - مثال: `9974-9952-9958 / خصم 50% على القطعة الثانية`
            """)
