import streamlit as st
import pandas as pd
import json
import re
from io import BytesIO
from datetime import datetime

def extract_single_sku(combined_sku):
    """استخراج SKU الفردي من SKU المجمع"""
    if pd.isna(combined_sku) or combined_sku == "":
        return ""
    
    sku_str = str(combined_sku).strip()
    
    # إزالة أي أرقام بعد * (مثل 14373*6 -> 14373)
    if '*' in sku_str:
        sku_str = sku_str.split('*')[0].strip()
    
    # إذا كان هناك عدة SKU مفصولة بـ - أو +، نأخذ أول واحد
    if '-' in sku_str:
        sku_str = sku_str.split('-')[0].strip()
    if '+' in sku_str:
        sku_str = sku_str.split('+')[0].strip()
    
    # إزالة أي أحرف غير رقمية
    sku_str = re.sub(r'[^0-9]', '', sku_str)
    
    return sku_str

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
    4. يتم البحث عن اسم المنتج باستخدام `SKU مجمع (للمراجعة)` بعد استخراج SKU الفردي
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
                    
                    # إنشاء قاموس المنتجات للبحث السريع (باستخدام SKU كـ string)
                    product_map = {}
                    for _, row in df_products.iterrows():
                        sku = str(row['SKU']).strip()
                        # إزالة .0 من نهاية الرقم إذا وجدت
                        if sku.endswith('.0'):
                            sku = sku[:-2]
                        name = str(row['ProductName']).strip()
                        product_map[sku] = name
                    
                    final_rows = []
                    processed_orders = 0
                    failed_orders = 0
                    
                    for idx, row in df_orders.iterrows():
                        order_id = row['رقم الطلب']
                        try:
                            # تحويل JSON إلى قائمة
                            skus_json = row['skus_json']
                            if isinstance(skus_json, str):
                                json_data = json.loads(skus_json)
                            else:
                                json_data = skus_json
                            
                            for item in json_data:
                                # هيكل البيانات المتوقع في item:
                                # item[0] = اسم المنتج الأساسي (أو معرف)
                                # item[1] = الكمية (أحياناً)
                                # item[2] = SKU المجمع
                                # item[3] = الكمية (في بعض الحالات)
                                # item[4] = سعر الوحدة
                                # item[5] = قائمة المنتجات الفرعية
                                
                                # تحديد SKU المجمع
                                combined_sku = ""
                                quantity = 0
                                unit_price = 0
                                product_name = ""
                                
                                # البحث عن SKU في positions المختلفة
                                for pos in range(len(item)):
                                    val = str(item[pos]) if item[pos] is not None else ""
                                    # إذا كان القيم عبارة عن SKU (أرقام فقط أو تحتوي على * أو -)
                                    if re.match(r'^[\d\*\-]+$', val) and len(val) > 2:
                                        combined_sku = val
                                        break
                                
                                # إذا لم نجد SKU، نستخدم item[2]
                                if not combined_sku and len(item) > 2:
                                    combined_sku = str(item[2]) if item[2] else ""
                                
                                # البحث عن الكمية
                                for pos in range(len(item)):
                                    val = item[pos]
                                    if isinstance(val, (int, float)) and val > 0 and val < 10000:
                                        # قد تكون الكمية أو السعر
                                        if quantity == 0:
                                            quantity = float(val)
                                        elif unit_price == 0 and val != quantity:
                                            unit_price = float(val)
                                
                                # إذا لم نجد الكمية، نستخدم item[3]
                                if quantity == 0 and len(item) > 3 and isinstance(item[3], (int, float)):
                                    quantity = float(item[3])
                                
                                # البحث عن اسم المنتج
                                if len(item) > 0 and isinstance(item[0], str) and not re.match(r'^[\d\*\-]+$', item[0]):
                                    product_name = item[0]
                                
                                # استخراج SKU الفردي من SKU المجمع
                                single_sku = extract_single_sku(combined_sku)
                                
                                # جلب اسم المنتج من قاموس المنتجات
                                if single_sku and single_sku in product_map:
                                    product_name = product_map[single_sku]
                                
                                total = quantity * unit_price if quantity and unit_price else 0
                                
                                # إضافة المنتج الأساسي
                                if combined_sku:
                                    final_rows.append([
                                        order_id,
                                        product_name if product_name else combined_sku,
                                        quantity,
                                        single_sku,
                                        combined_sku,
                                        unit_price,
                                        total,
                                        "أساسي"
                                    ])
                                
                                # --- معالجة المنتجات الفرعية ---
                                if len(item) > 5 and isinstance(item[5], list):
                                    sub_items_list = item[5]
                                    
                                    for sub_idx, sub in enumerate(sub_items_list):
                                        if not isinstance(sub, list) or len(sub) < 3:
                                            continue
                                        
                                        # استخراج بيانات المنتج الفرعي
                                        sub_combined_sku = str(sub[2]) if len(sub) > 2 else ""
                                        sub_quantity = sub[1] if len(sub) > 1 and isinstance(sub[1], (int, float)) else 0
                                        sub_unit_price = sub[3] if len(sub) > 3 and isinstance(sub[3], (int, float)) else 0
                                        sub_total = sub[4] if len(sub) > 4 and isinstance(sub[4], (int, float)) else (sub_quantity * sub_unit_price)
                                        
                                        # اسم المنتج الفرعي
                                        sub_name = sub[0] if len(sub) > 0 and isinstance(sub[0], str) else ""
                                        
                                        # استخراج SKU الفردي للمنتج الفرعي
                                        sub_single_sku = extract_single_sku(sub_combined_sku)
                                        
                                        # جلب اسم المنتج من قاموس المنتجات
                                        if sub_single_sku and sub_single_sku in product_map:
                                            sub_name = product_map[sub_single_sku]
                                        
                                        # تجنب إضافة منتجات فرعية مكررة أو بدون اسم
                                        if sub_combined_sku and sub_quantity > 0:
                                            final_rows.append([
                                                order_id,
                                                sub_name if sub_name else sub_combined_sku,
                                                sub_quantity,
                                                sub_single_sku,
                                                sub_combined_sku,
                                                sub_unit_price,
                                                sub_total,
                                                "فرعي"
                                            ])
                            
                            processed_orders += 1
                            
                        except Exception as e:
                            failed_orders += 1
                            st.warning(f"⚠️ خطأ في معالجة الطلب {order_id}: {str(e)[:100]}")
                            continue
                    
                    # إنشاء الجدول النهائي
                    columns = ['رقم الطلب', 'المنتج', 'الكمية', 'SKU فردي', 'SKU مجمع (للمراجعة)', 'سعر الوحدة', 'الإجمالي', 'النوع']
                    result_df = pd.DataFrame(final_rows, columns=columns)
                    
                    # تنظيف البيانات
                    result_df = result_df[result_df['المنتج'].notna()]
                    result_df = result_df[result_df['المنتج'] != ""]
                    result_df = result_df[result_df['الكمية'] > 0]
                    
                    # إزالة الصفوف المكررة تقريباً
                    result_df = result_df.drop_duplicates(subset=['رقم الطلب', 'SKU مجمع (للمراجعة)'])
                    
                    st.success(f"✅ تمت المعالجة بنجاح!")
                    st.info(f"📊 الإحصائيات: {processed_orders} طلب تمت معالجتها، {failed_orders} طلب فشل")
                    st.info(f"📦 عدد الصفوف المنتجة: {len(result_df)} (أساسي: {len(result_df[result_df['النوع'] == 'أساسي'])}, فرعي: {len(result_df[result_df['النوع'] == 'فرعي'])})")
                    
                    # عرض عينة من النتائج
                    st.subheader("📋 عينة من النتائج")
                    display_df = result_df.head(20).copy()
                    display_df['الكمية'] = display_df['الكمية'].round(2)
                    display_df['سعر الوحدة'] = display_df['سعر الوحدة'].round(2)
                    display_df['الإجمالي'] = display_df['الإجمالي'].round(2)
                    st.dataframe(display_df, use_container_width=True)
                    
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
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("عدد الطلبات الفريدة", result_df['رقم الطلب'].nunique())
                    with col2:
                        st.metric("عدد المنتجات الفريدة", result_df['SKU فردي'].nunique())
                    with col3:
                        total_qty = result_df['الكمية'].sum()
                        st.metric("إجمالي الكمية", f"{total_qty:,.2f}")
                    with col4:
                        total_amount = result_df['الإجمالي'].sum()
                        st.metric("إجمالي القيمة", f"{total_amount:,.2f}")
                    
                    # عرض المنتجات الأكثر مبيعاً
                    st.subheader("🏆 أكثر 10 منتجات مبيعاً")
                    top_products = result_df.groupby(['SKU فردي', 'المنتج']).agg({
                        'الكمية': 'sum',
                        'الإجمالي': 'sum'
                    }).reset_index().sort_values('الكمية', ascending=False).head(10)
                    top_products['الكمية'] = top_products['الكمية'].round(2)
                    top_products['الإجمالي'] = top_products['الإجمالي'].round(2)
                    st.dataframe(top_products, use_container_width=True)
                    
                except Exception as e:
                    st.error(f"❌ حدث خطأ أثناء المعالجة: {str(e)}")
                    st.exception(e)
    
    else:
        st.info("📂 الرجاء رفع ملفي الطلبات والمنتجات لبدء المعالجة")
