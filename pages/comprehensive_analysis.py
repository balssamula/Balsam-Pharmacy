import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import timedelta
from sklearn.linear_model import LinearRegression
from sklearn.ensemble import RandomForestRegressor

def show():
    st.markdown("""
    <div class="hero">
        <h1>📈 التحليل الشامل والتنبؤات للصيدليات</h1>
        <p>تحليل المبيعات، المشتريات، والمصروفات مع نماذج الذكاء الاصطناعي للتنبؤ بالمستقبل والمخزون</p>
    </div>
    """, unsafe_allow_html=True)

    uploaded_file = st.file_uploader("📂 ارفع ملف Excel (يحتوي على شيتات: المبيعات، المشتريات، المصروفات)", type=["xlsx", "xls"])

    if uploaded_file:
        try:
            with st.spinner("جاري قراءة البيانات وبناء النماذج التنبؤية..."):
                # قراءة الشيتات الثلاثة
                xls = pd.ExcelFile(uploaded_file)
                sheet_names = xls.sheet_names
                
                # افتراض ترتيب الشيتات كما ذكرت
                df_sales = pd.read_excel(xls, sheet_name=sheet_names[0])
                df_purchases = pd.read_excel(xls, sheet_name=sheet_names[1])
                df_expenses = pd.read_excel(xls, sheet_name=sheet_names[2])

                # تسمية أعمدة المصروفات حسب طلبك
                df_expenses.columns = ["المبلغ", "بيان المصروف", "مركز التكلفة"]
                
                # دوال مساعدة لتوحيد أسماء الأعمدة المتوقعة (بناءً على أمثلة سلة وABC المعتادة)
                def get_col(df, possible_names):
                    for name in possible_names:
                        if name in df.columns: return name
                    return df.columns[0] # احتياطي
                
                # تحديد الأعمدة الأساسية للمبيعات
                date_col = get_col(df_sales, ['التاريخ', 'تاريخ الطلب', 'Date', 'تاريخ'])
                time_col = get_col(df_sales, ['الوقت', 'وقت الطلب', 'Time'])
                product_col = get_col(df_sales, ['المنتج', 'اسم المنتج', 'الصنف', 'Product'])
                branch_col = get_col(df_sales, ['الفرع', 'الصيدلية', 'Branch'])
                user_col = get_col(df_sales, ['المستخدم', 'الصيدلي', 'الكاشير', 'البائع', 'User'])
                qty_col = get_col(df_sales, ['الكمية', 'العدد', 'Qty'])
                total_col = get_col(df_sales, ['الإجمالي', 'المبلغ', 'Total'])
                category_col = get_col(df_sales, ['التصنيف', 'القسم', 'الكاتيجوري', 'Category'])

                # تنظيف وتحضير التواريخ والأوقات
                df_sales[date_col] = pd.to_datetime(df_sales[date_col], errors='coerce')
                
                # استخراج الساعة وتحديد الليل والنهار
                if time_col in df_sales.columns:
                    # إذا كان الوقت مفصولاً
                    df_sales['Hour'] = pd.to_datetime(df_sales[time_col], format='%H:%M:%S', errors='coerce').dt.hour
                else:
                    # إذا كان الوقت مدمجاً مع التاريخ
                    df_sales['Hour'] = df_sales[date_col].dt.hour

                df_sales['Hour'] = df_sales['Hour'].fillna(12).astype(int)
                df_sales['فترة البيع'] = df_sales['Hour'].apply(lambda x: 'نهاراً ☀️' if 6 <= x < 18 else 'ليلاً 🌙')

                # تحويل القيم الرقمية
                df_sales[total_col] = pd.to_numeric(df_sales[total_col], errors='coerce').fillna(0)
                df_sales[qty_col] = pd.to_numeric(df_sales[qty_col], errors='coerce').fillna(1)
                
                pur_qty_col = get_col(df_purchases, ['الكمية', 'العدد', 'Qty'])
                pur_product_col = get_col(df_purchases, ['المنتج', 'اسم المنتج', 'الصنف', 'Product'])
                df_purchases[pur_qty_col] = pd.to_numeric(df_purchases[pur_qty_col], errors='coerce').fillna(0)
                df_expenses["المبلغ"] = pd.to_numeric(df_expenses["المبلغ"], errors='coerce').fillna(0)

                # ================= واجهة التبويبات =================
                tab1, tab2, tab3, tab4, tab5 = st.tabs([
                    "📊 مؤشرات الأداء", 
                    "⏰ تحليل الأوقات (ساعات الذروة)", 
                    "👥 الفروع، المنتجات والمستخدمين", 
                    "🤖 تنبؤات المبيعات والمستقبل (ML)", 
                    "📦 تقديرات المخزون والطلبيات"
                ])

                # ---------------- التبويب الأول: مؤشرات الأداء ----------------
                with tab1:
                    st.markdown("### 📈 ملخص الأداء المالي")
                    total_sales = df_sales[total_col].sum()
                    total_expenses = df_expenses["المبلغ"].sum()
                    pur_total_col = get_col(df_purchases, ['الإجمالي', 'المبلغ', 'Total'])
                    total_purchases = pd.to_numeric(df_purchases[pur_total_col], errors='coerce').fillna(0).sum()
                    net_profit = total_sales - total_purchases - total_expenses

                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("إجمالي المبيعات", f"{total_sales:,.2f} ريال")
                    c2.metric("إجمالي المشتريات (تكلفة)", f"{total_purchases:,.2f} ريال")
                    c3.metric("إجمالي المصروفات", f"{total_expenses:,.2f} ريال")
                    c4.metric("صافي الدخل المؤقت", f"{net_profit:,.2f} ريال", delta=float(net_profit))

                    st.markdown("---")
                    st.markdown("### 💸 تحليل المصروفات حسب مركز التكلفة")
                    exp_summary = df_expenses.groupby("مركز التكلفة")["المبلغ"].sum().reset_index()
                    fig_exp = px.pie(exp_summary, values='المبلغ', names='مركز التكلفة', hole=0.4, title="توزيع المصروفات")
                    st.plotly_chart(fig_exp, use_container_width=True)

                # ---------------- التبويب الثاني: تحليل الأوقات ----------------
                with tab2:
                    st.markdown("### ⏰ تحليل أوقات البيع (الليل vs النهار)")
                    c1, c2 = st.columns(2)
                    
                    with c1:
                        time_sales = df_sales.groupby('فترة البيع')[total_col].sum().reset_index()
                        fig_time = px.pie(time_sales, values=total_col, names='فترة البيع', color='فترة البيع',
                                          color_discrete_map={'نهاراً ☀️':'#f1c40f', 'ليلاً 🌙':'#2c3e50'})
                        st.plotly_chart(fig_time, use_container_width=True)

                    with c2:
                        st.markdown("### 📈 أكثر الساعات ذروة في المبيعات")
                        hourly_sales = df_sales.groupby('Hour')[total_col].sum().reset_index()
                        fig_hour = px.bar(hourly_sales, x='Hour', y=total_col, 
                                          labels={'Hour': 'ساعة اليوم (0-23)', total_col: 'حجم المبيعات'},
                                          text_auto='.2s')
                        st.plotly_chart(fig_hour, use_container_width=True)
                        
                    st.markdown("### 🕒 تحليل المنتج مع توقيت بيعه")
                    top_products = df_sales.groupby(product_col)[total_col].sum().nlargest(10).index
                    df_top_prod = df_sales[df_sales[product_col].isin(top_products)]
                    prod_time_matrix = df_top_prod.pivot_table(index=product_col, columns='فترة البيع', values=total_col, aggfunc='sum').fillna(0)
                    st.dataframe(prod_time_matrix.style.background_gradient(cmap='Blues'), use_container_width=True)

                # ---------------- التبويب الثالث: الفروع والمستخدمين ----------------
                with tab3:
                    st.markdown("### 🏢 تحليل المنتج مع الفرع")
                    branch_prod = df_sales.pivot_table(index=product_col, columns=branch_col, values=total_col, aggfunc='sum').fillna(0)
                    # عرض أعلى 15 منتج مبيعاً لتسهيل القراءة
                    branch_prod['الإجمالي'] = branch_prod.sum(axis=1)
                    branch_prod = branch_prod.sort_values(by='الإجمالي', ascending=False).head(15).drop(columns=['الإجمالي'])
                    fig_bp = px.imshow(branch_prod, aspect="auto", color_continuous_scale='Viridis', title="خريطة حرارية: أفضل المنتجات لكل فرع")
                    st.plotly_chart(fig_bp, use_container_width=True)

                    st.markdown("---")
                    c1, c2 = st.columns(2)
                    with c1:
                        st.markdown("### 👤 تحليل المنتج مع المستخدم (الصيدلي)")
                        user_prod = df_sales.groupby([user_col, product_col])[total_col].sum().reset_index()
                        top_user_prod = user_prod.sort_values(total_col, ascending=False).groupby(user_col).head(3)
                        st.dataframe(top_user_prod, use_container_width=True)
                    
                    with c2:
                        st.markdown("### 🏷️ تحليل مبيعات المستخدم مع الكاتيجوري")
                        if category_col in df_sales.columns:
                            user_cat = df_sales.pivot_table(index=user_col, columns=category_col, values=total_col, aggfunc='sum').fillna(0)
                            st.dataframe(user_cat.style.background_gradient(cmap='Greens'), use_container_width=True)
                        else:
                            st.info("عمود التصنيف/الكاتيجوري غير موجود في بيانات المبيعات المرفوعة.")

                # ---------------- التبويب الرابع: التنبؤات (Machine Learning) ----------------
                with tab4:
                    st.markdown("### 🤖 نماذج التنبؤ بالمبيعات والعوائد (Forecasting Models)")
                    st.info("تستخدم هذه النماذج خوارزميات Scikit-Learn (Linear Regression & Random Forest) لتحليل الاتجاهات التاريخية والتنبؤ بالـ 30 يوماً القادمة.")
                    
                    # تجهيز السلسلة الزمنية
                    daily_sales = df_sales.groupby(date_col)[total_col].sum().reset_index()
                    daily_sales = daily_sales.sort_values(date_col)
                    
                    if len(daily_sales) > 5: # نحتاج بيانات كافية للتدريب
                        daily_sales['Days_Since_Start'] = (daily_sales[date_col] - daily_sales[date_col].min()).dt.days
                        daily_sales['Day_of_Week'] = daily_sales[date_col].dt.dayofweek
                        
                        X = daily_sales[['Days_Since_Start', 'Day_of_Week']]
                        y = daily_sales[total_col]
                        
                        # تدريب النموذج
                        model = RandomForestRegressor(n_estimators=100, random_state=42)
                        model.fit(X, y)
                        
                        # التنبؤ بـ 30 يوم قادمة
                        last_date = daily_sales[date_col].max()
                        future_dates = [last_date + timedelta(days=x) for x in range(1, 31)]
                        future_days_since = [(d - daily_sales[date_col].min()).days for d in future_dates]
                        future_dow = [d.weekday() for d in future_dates]
                        
                        X_future = pd.DataFrame({'Days_Since_Start': future_days_since, 'Day_of_Week': future_dow})
                        predictions = model.predict(X_future)
                        
                        future_df = pd.DataFrame({
                            'التاريخ': future_dates,
                            'المبيعات المتوقعة': predictions,
                            'النوع': 'تنبؤ (مستقبل)'
                        })
                        
                        historical_df = pd.DataFrame({
                            'التاريخ': daily_sales[date_col],
                            'المبيعات المتوقعة': daily_sales[total_col],
                            'النوع': 'بيانات فعلية (تاريخي)'
                        })
                        
                        forecast_plot_df = pd.concat([historical_df, future_df])
                        
                        fig_forecast = px.line(forecast_plot_df, x='التاريخ', y='المبيعات المتوقعة', color='النوع',
                                               title="📈 التنبؤ بإجمالي العوائد للـ 30 يوماً القادمة",
                                               color_discrete_map={'بيانات فعلية (تاريخي)':'#2980b9', 'تنبؤ (مستقبل)':'#e74c3c'})
                        st.plotly_chart(fig_forecast, use_container_width=True)
                        
                        st.success(f"💰 إجمالي العائد المتوقع للـ 30 يوماً القادمة: **{predictions.sum():,.2f} ريال**")
                        
                        st.markdown("### 🔮 التنبؤ بأداء الفروع القادم")
                        # نموذج مبسط انحدار خطي لكل فرع
                        branch_forecast = []
                        for branch in df_sales[branch_col].unique():
                            b_sales = df_sales[df_sales[branch_col] == branch].groupby(date_col)[total_col].sum().reset_index()
                            if len(b_sales) > 3:
                                b_sales['d'] = (b_sales[date_col] - b_sales[date_col].min()).dt.days
                                lr = LinearRegression()
                                lr.fit(b_sales[['d']], b_sales[total_col])
                                future_d = np.array([[b_sales['d'].max() + 15]]) # التنبؤ لمنتصف الشهر القادم
                                expected_daily = lr.predict(future_d)[0]
                                branch_forecast.append({"الفرع": branch, "المتوسط اليومي المتوقع": max(0, expected_daily)})
                        
                        if branch_forecast:
                            st.dataframe(pd.DataFrame(branch_forecast).style.format({"المتوسط اليومي المتوقع": "{:.2f}"}), use_container_width=True)
                    else:
                        st.warning("البيانات الزمنية المتاحة غير كافية لبناء نموذج تنبؤ موثوق (نحتاج أكثر من 5 أيام).")

                # ---------------- التبويب الخامس: المخزون والطلبيات ----------------
                with tab5:
                    st.markdown("### 📦 وضع تقديرات للمخزون وطلبيات النواقص")
                    st.info("يقوم النظام بدمج المشتريات السابقة مطروحاً منها المبيعات لمعرفة (المخزون الدفتري)، ثم حساب معدل الحرق اليومي (Run Rate) للتنبؤ بالكميات المطلوبة لـ 30 يوماً.")
                    
                    # تجميع المبيعات الكلية لكل منتج
                    sales_qty = df_sales.groupby(product_col)[qty_col].sum().reset_index().rename(columns={qty_col: 'إجمالي المباع'})
                    
                    # تجميع المشتريات الكلية لكل منتج
                    purchases_qty = df_purchases.groupby(pur_product_col)[pur_qty_col].sum().reset_index().rename(columns={pur_product_col: product_col, pur_qty_col: 'إجمالي المشتريات'})
                    
                    # دمج المخزون
                    inventory_df = pd.merge(purchases_qty, sales_qty, on=product_col, how='outer').fillna(0)
                    inventory_df['المخزون الحالي (التقديري)'] = inventory_df['إجمالي المشتريات'] - inventory_df['إجمالي المباع']
                    
                    # حساب معدل البيع اليومي (افتراض فترة البيانات 90 يوم من المثال)
                    days_in_data = max((df_sales[date_col].max() - df_sales[date_col].min()).days, 1)
                    inventory_df['معدل البيع اليومي'] = inventory_df['إجمالي المباع'] / days_in_data
                    
                    # الكمية المطلوبة لتغطية 30 يوم
                    inventory_df['الاحتياج لـ 30 يوم'] = np.ceil(inventory_df['معدل البيع اليومي'] * 30)
                    
                    # تحديد الطلبية: الاحتياج - المخزون الحالي (إذا كان الناتج موجباً)
                    inventory_df['الكمية المقترح طلبها'] = np.where(
                        inventory_df['الاحتياج لـ 30 يوم'] > inventory_df['المخزون الحالي (التقديري)'],
                        inventory_df['الاحتياج لـ 30 يوم'] - inventory_df['المخزون الحالي (التقديري)'],
                        0
                    )
                    
                    # تصفية المنتجات التي تحتاج لطلب
                    reorder_df = inventory_df[inventory_df['الكمية المقترح طلبها'] > 0].sort_values('الكمية المقترح طلبها', ascending=False)
                    
                    c1, c2 = st.columns(2)
                    c1.metric("عدد الأصناف التي تحتاج إعادة طلب", f"{len(reorder_df)} صنف")
                    if len(reorder_df) > 0:
                        st.dataframe(reorder_df[[product_col, 'المخزون الحالي (التقديري)', 'الاحتياج لـ 30 يوم', 'الكمية المقترح طلبها']].style.background_gradient(subset=['الكمية المقترح طلبها'], cmap='Reds'), use_container_width=True)
                    else:
                        st.success("المخزون الحالي كافٍ لجميع المنتجات للفترة القادمة.")

        except Exception as e:
            st.error(f"حدث خطأ أثناء معالجة الملف. يرجى التأكد من أن الأعمدة متوافقة: {e}")
