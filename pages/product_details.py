import streamlit as st
import pandas as pd
import json
import re
from io import BytesIO
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go

def extract_single_sku(combined_sku):
    """استخراج SKU الفردي من SKU المجمع"""
    if pd.isna(combined_sku) or combined_sku == "":
        return ""
    
    sku_str = str(combined_sku).strip()
    if '*' in sku_str:
        sku_str = sku_str.split('*')[0].strip()
    if '-' in sku_str:
        sku_str = sku_str.split('-')[0].strip()
    if '+' in sku_str:
        sku_str = sku_str.split('+')[0].strip()
    
    # إزالة أي أحرف غير رقمية
    if sku_str.replace('.', '').isdigit():
        sku_str = re.sub(r'[^0-9]', '', sku_str)
    return sku_str

def safe_float_convert(value):
    """تحويل آمن إلى float"""
    if pd.isna(value):
        return 0.0
    try:
        return float(value)
    except (ValueError, TypeError):
        return 0.0

def show():
    st.markdown("""
    <div class="hero">
        <h1>📦 تفصيلي المنتجات وتحليلات متقدمة</h1>
        <p>تحليل تفاصيل المنتجات من ملف الطلبات مع إحصائيات متقدمة ورسوم بيانية تفاعلية</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.info("""
    **📌 تعليمات استخدام هذه الصفحة:**
    1. قم برفع ملف `orders.xlsx` (يحتوي على أعمدة: رقم الطلب، skus_json، الخصم، تكلفة الشحن، طريقة الدفع، الضريبة، تاريخ الطلب)
    2. قم برفع ملف `products.xlsx` (يحتوي على عمودي 'SKU' و 'ProductName')
    3. سيتم معالجة البيانات واستخراج تفاصيل المنتجات
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
                    df_orders = pd.read_excel(orders_file)
                    df_products = pd.read_excel(products_file)
                    
                    df_orders.columns = df_orders.columns.str.strip()
                    df_products.columns = df_products.columns.str.strip()
                    
                    st.info(f"📊 تم قراءة {len(df_orders)} طلب و {len(df_products)} منتج")
                    
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
                    
                    product_map = {}
                    for _, row in df_products.iterrows():
                        sku = str(row['SKU']).strip()
                        if sku.endswith('.0'): sku = sku[:-2]
                        name = str(row['ProductName']).strip()
                        product_map[sku] = name
                    
                    order_info_map = {}
                    for _, row in df_orders.iterrows():
                        order_id = row['رقم الطلب']
                        order_info_map[order_id] = {
                            'الخصم': safe_float_convert(row.get('الخصم', 0)),
                            'تكلفة الشحن': safe_float_convert(row.get('تكلفة الشحن', 0)),
                            'طريقة الدفع': str(row.get('طريقة الدفع', 'غير محدد')),
                            'الضريبة': safe_float_convert(row.get('الضريبة', 0)),
                            'تاريخ الطلب': row.get('تاريخ الطلب', ''),
                            'قيمة خصم الكوبون': safe_float_convert(row.get('قيمة خصم الكوبون', 0)),
                            'قيمة خصم العروض الخاصة': safe_float_convert(row.get('قيمة خصم العروض الخاصة', 0))
                        }
                    
                    raw_rows = []
                    processed_orders = 0
                    failed_orders = 0
                    
                    for _, row in df_orders.iterrows():
                        order_id = row['رقم الطلب']
                        order_info = order_info_map.get(order_id, {})
                        
                        try:
                            skus_json = row['skus_json']
                            json_data = json.loads(skus_json) if isinstance(skus_json, str) else skus_json
                            
                            for item in json_data:
                                combined_sku = ""
                                quantity = 0
                                unit_price = 0
                                product_name = ""
            
                                # 1. استخراج الـ SKU بأمان (الاعتماد على الموقع الافتراضي أولاً وهو الدليل رقم 2)
                                if len(item) > 2 and item[2] and str(item[2]).replace('*', '').replace('-', '').isdigit():
                                    combined_sku = str(item[2]).strip()
                                else:
                                    # إذا لم يكن في الموقع 2، نبحث في باقي المواقع مع استثناء الموقع 1 (لأنه دائما للكمية)
                                    for pos in range(len(item)):
                                        if pos == 1: continue # تخطي موقع الكمية
                                        val = str(item[pos]) if item[pos] is not None else ""
                                        if re.match(r'^[\d\*\-]+$', val) and len(val) > 2:
                                            combined_sku = val
                                            break
            
                                # 2. استخراج الكمية والسعر بشكل صريح وصحيح
                                if len(item) > 1 and isinstance(item[1], (int, float)):
                                    quantity = float(item[1])
                
                                if len(item) > 3 and isinstance(item[3], (int, float)):
                                    unit_price = float(item[3])
            
                                # كخطة بديلة إذا كانت المصفوفة غير مرتبة
                                if quantity == 0 or unit_price == 0:
                                    for pos in range(len(item)):
                                        val = item[pos]
                                        if isinstance(val, (int, float)) and val > 0:
                                            if quantity == 0:
                                                quantity = float(val)
                                            elif unit_price == 0 and val != quantity:
                                                unit_price = float(val)
                                
                                if quantity == 0 and len(item) > 3 and isinstance(item[3], (int, float)):
                                    quantity = float(item[3])
                                
                                if len(item) > 0 and isinstance(item[0], str) and not re.match(r'^[\d\*\-]+$', item[0]):
                                    product_name = item[0]
                                
                                single_sku = extract_single_sku(combined_sku)
                                if single_sku and single_sku in product_map:
                                    product_name = product_map[single_sku]
                                
                                total = quantity * unit_price if quantity and unit_price else 0
                                
                                if combined_sku:
                                    raw_rows.append({
                                        'رقم الطلب': order_id,
                                        'المنتج': product_name if product_name else combined_sku,
                                        'الكمية': quantity,
                                        'SKU فردي': combined_sku, # نحتفظ بالاس كيو يو المجمع كفردي هنا لكي لا يُدمج
                                        'SKU مجمع (للمراجعة)': combined_sku,
                                        'سعر الوحدة': unit_price,
                                        'الإجمالي': total,
                                        'النوع': 'أساسي (مجموعة)' if '*' in combined_sku or '-' in combined_sku else 'أساسي',
                                        'الخصم': order_info.get('الخصم', 0),
                                        'تكلفة الشحن': order_info.get('تكلفة الشحن', 0),
                                        'طريقة الدفع': order_info.get('طريقة الدفع', 'غير محدد'),
                                        'الضريبة': order_info.get('الضريبة', 0),
                                        'تاريخ الطلب': order_info.get('تاريخ الطلب', ''),
                                        'قيمة خصم الكوبون': order_info.get('قيمة خصم الكوبون', 0),
                                        'قيمة خصم العروض الخاصة': order_info.get('قيمة خصم العروض الخاصة', 0)
                                    })
                                
                                # معالجة المنتجات الفرعية (المفككة)
                                if len(item) > 5 and isinstance(item[5], list):
                                    for sub in item[5]:
                                        if not isinstance(sub, list) or len(sub) < 3: continue
                                        
                                        sub_combined_sku = str(sub[2]) if len(sub) > 2 else ""
                                        sub_quantity = sub[1] if len(sub) > 1 and isinstance(sub[1], (int, float)) else 0
                                        sub_unit_price = sub[3] if len(sub) > 3 and isinstance(sub[3], (int, float)) else 0
                                        sub_total = sub[4] if len(sub) > 4 and isinstance(sub[4], (int, float)) else (sub_quantity * sub_unit_price)
                                        
                                        sub_name = sub[0] if len(sub) > 0 and isinstance(sub[0], str) else ""
                                        sub_single_sku = extract_single_sku(sub_combined_sku)
                                        
                                        if sub_single_sku and sub_single_sku in product_map:
                                            sub_name = product_map[sub_single_sku]
                                        
                                        # حساب الكمية الحقيقية = كمية المنتج الفرعي × كمية المجموعة الأساسية
                                        actual_sub_qty = sub_quantity
                                        
                                        if sub_combined_sku and sub_quantity > 0:
                                            raw_rows.append({
                                                'رقم الطلب': order_id,
                                                'المنتج': sub_name if sub_name else sub_combined_sku,
                                                'الكمية': actual_sub_qty,
                                                'SKU فردي': sub_single_sku, # هذا سيُدمج مع الأساسي
                                                'SKU مجمع (للمراجعة)': sub_combined_sku,
                                                'سعر الوحدة': sub_unit_price,
                                                'الإجمالي': sub_total * quantity, # ضرب الاجمالي في كمية المجموعة
                                                'النوع': 'فرعي (مفكك)',
                                                'الخصم': order_info.get('الخصم', 0),
                                                'تكلفة الشحن': order_info.get('تكلفة الشحن', 0),
                                                'طريقة الدفع': order_info.get('طريقة الدفع', 'غير محدد'),
                                                'الضريبة': order_info.get('الضريبة', 0),
                                                'تاريخ الطلب': order_info.get('تاريخ الطلب', ''),
                                                'قيمة خصم الكوبون': order_info.get('قيمة خصم الكوبون', 0),
                                                'قيمة خصم العروض الخاصة': order_info.get('قيمة خصم العروض الخاصة', 0)
                                            })
                            processed_orders += 1
                        except Exception as e:
                            failed_orders += 1
                            continue
                    
                    # إنشاء DataFrame النهائي
                    result_df = pd.DataFrame(raw_rows)
                    
                    numeric_cols = ['الكمية', 'سعر الوحدة', 'الإجمالي', 'الخصم', 'تكلفة الشحن', 'الضريبة', 'قيمة خصم الكوبون', 'قيمة خصم العروض الخاصة']
                    for col in numeric_cols:
                        if col in result_df.columns: result_df[col] = result_df[col].apply(safe_float_convert)
                    
                    result_df = result_df[result_df['المنتج'].notna() & (result_df['المنتج'] != "") & (result_df['الكمية'] > 0)]
                    
                    # 💡 المنطق الذكي للتجميع: يتم الدمج بناءً على الـ SKU الفردي ورقم الطلب، 
                    # لكن المنتجات التي نوعها "أساسي (مجموعة)" ستبقى مفصولة لأن الـ SKU الفردي لها يحتوي على (*) 
                    group_cols = [
                        'رقم الطلب', 'SKU فردي', 'الخصم', 'تكلفة الشحن', 'طريقة الدفع', 
                        'الضريبة', 'تاريخ الطلب', 'قيمة خصم الكوبون', 'قيمة خصم العروض الخاصة'
                    ]
                    
                    # تجميع وحفظ أول قيمة للمنتج والسعر والمجمع
                    result_df = result_df.groupby(group_cols, as_index=False).agg({
                        'الكمية': 'sum',
                        'الإجمالي': 'sum',
                        'المنتج': 'first',
                        'سعر الوحدة': 'first',
                        'SKU مجمع (للمراجعة)': 'first',
                        'النوع': 'first' # للحفاظ على التصنيف
                    })
                    
                    if 'تاريخ الطلب' in result_df.columns:
                        result_df['تاريخ الطلب'] = pd.to_datetime(result_df['تاريخ الطلب'], errors='coerce')
                        
                    # ترتيب الأعمدة ليكون شكلها منطقي
                    columns_order = [
                        'رقم الطلب', 'المنتج', 'الكمية', 'SKU فردي', 'SKU مجمع (للمراجعة)', 
                        'سعر الوحدة', 'الإجمالي', 'النوع', 'الخصم', 'تكلفة الشحن', 'طريقة الدفع', 
                        'الضريبة', 'تاريخ الطلب', 'قيمة خصم الكوبون', 'قيمة خصم العروض الخاصة'
                    ]
                    result_df = result_df[[col for col in columns_order if col in result_df.columns]]
                    
                    st.success(f"✅ تمت المعالجة بنجاح وتم دمج المنتجات الفرعية مع الأساسية!")
                    st.info(f"📊 الإحصائيات: {processed_orders} طلب تمت معالجتها، {failed_orders} طلب فشل")
                    st.info(f"📦 عدد الصفوف المنتجة: {len(result_df)}")
                    
                    # ========== التبويبات المتقدمة ==========
                    tab1, tab2, tab3, tab4, tab5 = st.tabs([
                        "📋 جدول البيانات التفصيلي",
                        "📊 إحصائيات وتحليلات",
                        "📈 رسوم بيانية تفاعلية",
                        "💰 تحليل المبيعات",
                        "🏷️ تحليل طرق الدفع"
                    ])
                    
                    with tab1:
                        st.subheader("📋 جدول البيانات التفصيلي")
                        st.dataframe(result_df, use_container_width=True)
                        output = BytesIO()
                        result_df.to_excel(output, index=False)
                        output.seek(0)
                        st.download_button(
                            "📥 تحميل ملف النتائج (Excel)",
                            data=output,
                            file_name=f"product_details_merged_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                        )
                    
                    with tab2:
                        st.subheader("📊 إحصائيات وتحليلات")
                        col1, col2, col3, col4 = st.columns(4)
                        with col1: st.metric("عدد الطلبات الفريدة", result_df['رقم الطلب'].nunique())
                        with col2: st.metric("عدد المنتجات الفريدة", result_df['SKU فردي'].nunique())
                        with col3: st.metric("إجمالي الكمية", f"{result_df['الكمية'].sum():,.2f}")
                        with col4: st.metric("إجمالي المبيعات", f"{result_df['الإجمالي'].sum():,.2f} ₴")
                        
                        st.markdown("---")
                        st.subheader("🏆 أكثر 10 منتجات مبيعاً")
                        top_products = result_df.groupby(['SKU فردي', 'المنتج']).agg({
                            'الكمية': 'sum', 'الإجمالي': 'sum'
                        }).reset_index().sort_values('الكمية', ascending=False).head(10)
                        st.dataframe(top_products, use_container_width=True)
                        
                        st.subheader("📦 توزيع المنتجات حسب النوع")
                        type_stats = result_df.groupby('النوع').agg({'الكمية': 'sum', 'الإجمالي': 'sum'}).reset_index()
                        st.dataframe(type_stats, use_container_width=True)
                    
                    with tab3:
                        st.subheader("📈 الرسوم البيانية التفاعلية")
                        if 'تاريخ الطلب' in result_df.columns and not result_df['تاريخ الطلب'].isna().all():
                            st.markdown("### 📅 المبيعات اليومية")
                            daily_sales = result_df.groupby(result_df['تاريخ الطلب'].dt.date).agg({'الإجمالي': 'sum'}).reset_index().sort_values('تاريخ الطلب')
                            fig1 = px.line(daily_sales, x='تاريخ الطلب', y='الإجمالي', title='المبيعات اليومية', labels={'الإجمالي': 'المبيعات (₴)', 'تاريخ الطلب': 'التاريخ'})
                            st.plotly_chart(fig1, use_container_width=True)
                        
                        st.markdown("### 🏷️ أفضل 10 منتجات من حيث المبيعات")
                        top_10 = result_df.groupby('المنتج')['الإجمالي'].sum().sort_values(ascending=False).head(10).reset_index()
                        if not top_10.empty:
                            fig2 = px.bar(top_10, x='الإجمالي', y='المنتج', orientation='h', title='أفضل 10 منتجات من حيث المبيعات')
                            st.plotly_chart(fig2, use_container_width=True)
                    
                    with tab4:
                        st.subheader("💰 تحليل المبيعات")
                        total_revenue = result_df['الإجمالي'].sum()
                        total_quantity = result_df['الكمية'].sum()
                        avg_unit_price = total_revenue / total_quantity if total_quantity > 0 else 0
                        col1, col2, col3 = st.columns(3)
                        with col1: st.metric("إجمالي الإيرادات", f"{total_revenue:,.2f} ₴")
                        with col2: st.metric("إجمالي الكمية", f"{total_quantity:,.0f}")
                        with col3: st.metric("متوسط سعر الوحدة", f"{avg_unit_price:,.2f} ₴")
                    
                    with tab5:
                        st.subheader("🏷️ تحليل طرق الدفع")
                        payment_stats = result_df.groupby('طريقة الدفع').agg({'الإجمالي': 'sum', 'رقم الطلب': 'nunique'}).reset_index()
                        payment_stats.columns = ['طريقة الدفع', 'إجمالي المبيعات', 'عدد الطلبات']
                        payment_stats = payment_stats.sort_values('إجمالي المبيعات', ascending=False)
                        st.dataframe(payment_stats, use_container_width=True)
                        fig5 = px.pie(payment_stats, values='إجمالي المبيعات', names='طريقة الدفع', title='نسبة المبيعات حسب طريقة الدفع', hole=0.3)
                        st.plotly_chart(fig5, use_container_width=True)
                        
                except Exception as e:
                    st.error(f"❌ حدث خطأ أثناء المعالجة: {str(e)}")
                    st.exception(e)
    else:
        st.info("📂 الرجاء رفع ملفي الطلبات والمنتجات لبدء المعالجة")
