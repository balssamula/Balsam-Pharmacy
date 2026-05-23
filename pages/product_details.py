import streamlit as st
import pandas as pd
import json
from io import BytesIO
from datetime import datetime

def show():
    st.markdown("""
    <div class="hero">
        <h1>📦 تفصيلي المنتجات</h1>
        <p>تحليل تفاصيل المنتجات من ملف الطلبات وربطها بملف المنتجات</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.info("""
    **📌 تعليمات استخدام هذه الصفحة:**
    1. قم برفع ملف `orders.xlsx` (يحتوي على عمودي 'رقم الطلب' و 'skus_json')
    2. قم برفع ملف `products.xlsx` (يحتوي على عمودي 'SKU' و 'ProductName')
    3. سيتم معالجة البيانات واستخراج تفاصيل المنتجات الرئيسية والفرعية
    4. يمكنك تحميل النتيجة كملف Excel
    """)
    
    col1, col2 = st.columns(2)
    
    with col1:
        orders_file = st.file_uploader("📊 رفع ملف الطلبات (orders.xlsx)", type=["xlsx"], key="orders_file")
        if orders_file:
            st.success("✅ تم رفع ملف الطلبات")
    
    with col2:
        products_file = st.file_uploader("📦 رفع ملف المنتجات (products.xlsx)", type=["xlsx"], key="products_file")
        if products_file:
            st.success("✅ تم رفع ملف المنتجات")
    
    if orders_file and products_file:
        if st.button("🔄 معالجة البيانات", use_container_width=True, type="primary"):
            with st.spinner("جاري معالجة البيانات..."):
                try:
                    # قراءة الملفات
                    df_orders = pd.read_excel(orders_file)
                    df_products = pd.read_excel(products_file)
                    
                    # تنظيف أسماء الأعمدة
                    df_orders.columns = df_orders.columns.str.strip()
                    df_products.columns = df_products.columns.str.strip()
                    
                    st.info(f"📊 تم قراءة {len(df_orders)} طلب و {len(df_products)} منتج")
                    
                    # التحقق من وجود الأعمدة المطلوبة
                    required_order_cols = ['رقم الطلب', 'skus_json']
                    required_product_cols = ['SKU', 'ProductName']
                    
                    missing_order = [col for col in required_order_cols if col not in df_orders.columns]
                    missing_product = [col for col in required_product_cols if col not in df_products.columns]
                    
                    if missing_order:
                        st.error(f"❌ الأعمدة المفقودة في ملف الطلبات: {missing_order}")
                        st.stop()
                    if missing_product:
                        st.error(f"❌ الأعمدة المفقودة في ملف المنتجات: {missing_product}")
                        st.stop()
                    
                    # إنشاء قاموس المنتجات للبحث السريع
                    product_map = dict(zip(df_products['SKU'].astype(str), df_products['ProductName']))
                    
                    final_rows = []
                    processed_orders = 0
                    failed_orders = 0
                    
                    for _, row in df_orders.iterrows():
                        order_id = row['رقم الطلب']
                        try:
                            # تحويل JSON إلى قائمة
                            json_data = json.loads(row['skus_json'])
                            
                            for item in json_data:
                                # --- سطر المنتج الأساسي ---
                                main_sku_combined = str(item[2]) if len(item) > 2 else ""
                                final_rows.append([
                                    order_id, 
                                    item[1] if len(item) > 1 else "",  # المنتج
                                    item[3] if len(item) > 3 else 0,   # الكمية
                                    main_sku_combined,                  # SKU فردي
                                    main_sku_combined,                  # SKU مجمع
                                    item[4] if len(item) > 4 else 0,   # سعر الوحدة
                                    item[4] * item[3] if len(item) > 3 and len(item) > 4 else 0,  # الإجمالي
                                    "أساسي"                              # النوع
                                ])
                                
                                # --- معالجة المنتجات الفرعية ---
                                if len(item) > 5 and isinstance(item[5], list):
                                    sub_items_list = item[5]
                                    
                                    for idx, sub in enumerate(sub_items_list):
                                        raw_sku_combined = str(sub[2]) if len(sub) > 2 else ""
                                        
                                        # تحديد الفاصل وتقسيم الـ SKU
                                        if '+' in raw_sku_combined:
                                            delim = '+'
                                        elif '-' in raw_sku_combined:
                                            delim = '-'
                                        else:
                                            delim = None
                                        
                                        if delim:
                                            sku_parts = [s.split('*')[0].strip() for s in raw_sku_combined.split(delim)]
                                        else:
                                            sku_parts = [raw_sku_combined]
                                        
                                        # ربط المنتج الفرعي بالـ SKU المنفرد حسب الترتيب
                                        try:
                                            current_sku_single = sku_parts[idx]
                                        except IndexError:
                                            current_sku_single = sku_parts[-1] if sku_parts else raw_sku_combined
                                        
                                        # جلب الاسم من المرجع أو استخدام الاسم الموجود
                                        found_name = product_map.get(current_sku_single, sub[0] if len(sub) > 0 else "")
                                        
                                        final_rows.append([
                                            order_id,
                                            found_name,
                                            sub[3] if len(sub) > 3 else 0,   # الكمية
                                            current_sku_single,               # SKU فردي
                                            raw_sku_combined,                 # SKU مجمع
                                            sub[4] if len(sub) > 4 else 0,    # سعر الوحدة
                                            (sub[4] * sub[3]) if len(sub) > 3 and len(sub) > 4 else 0,  # الإجمالي
                                            "فرعي"                             # النوع
                                        ])
                            
                            processed_orders += 1
                            
                        except Exception as e:
                            failed_orders += 1
                            continue
                    
                    # إنشاء الجدول النهائي
                    columns = ['رقم الطلب', 'المنتج', 'الكمية', 'SKU فردي', 'SKU مجمع (للمراجعة)', 'سعر الوحدة', 'الإجمالي', 'النوع']
                    result_df = pd.DataFrame(final_rows, columns=columns)
                    
                    st.success(f"✅ تمت المعالجة بنجاح!")
                    st.info(f"📊 الإحصائيات: {processed_orders} طلب تمت معالجتها، {failed_orders} طلب فشل")
                    st.info(f"📦 عدد الصفوف المنتجة: {len(result_df)} (أساسي: {len(result_df[result_df['النوع'] == 'أساسي'])}, فرعي: {len(result_df[result_df['النوع'] == 'فرعي'])})")
                    
                    # عرض عينة من النتائج
                    st.subheader("📋 عينة من النتائج")
                    st.dataframe(result_df.head(20), use_container_width=True)
                    
                    # زر تحميل النتيجة
                    output = BytesIO()
                    result_df.to_excel(output, index=False)
                    output.seek(0)
                    
                    st.download_button(
                        "📥 تحميل ملف النتائج (Excel)",
                        data=output,
                        file_name=f"product_details_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                    
                    # إحصائيات إضافية
                    st.subheader("📊 إحصائيات إضافية")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("عدد الطلبات الفريدة", result_df['رقم الطلب'].nunique())
                    with col2:
                        st.metric("عدد المنتجات الفريدة", result_df['SKU فردي'].nunique())
                    with col3:
                        st.metric("إجمالي الكمية", int(result_df['الكمية'].sum()))
                    
                except Exception as e:
                    st.error(f"❌ حدث خطأ أثناء المعالجة: {str(e)}")
    
    else:
        st.info("📂 الرجاء رفع ملفي الطلبات والمنتجات لبدء المعالجة")