import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from datetime import timedelta
from sklearn.linear_model import LinearRegression
from sklearn.ensemble import RandomForestRegressor

def show():
    st.markdown("""
    <div class="hero">
        <h1>📈 التحليل الشامل والتنبؤات للصيدليات</h1>
        <p>تحليل المبيعات، المشتريات، والمصروفات مع نماذج الذكاء الاصطناعي</p>
    </div>
    """, unsafe_allow_html=True)

    uploaded_file = st.file_uploader("📂 ارفع ملف Excel (المبيعات، المشتريات، المصروفات)", type=["xlsx", "xls"])

    if uploaded_file:
        try:
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = xls.sheet_names
            
            # 💡 [قراءة ذكية]: التحقق مما إذا كانت العناوين في الصف الثاني (مثل ملفات سلة)
            df_sales = pd.read_excel(xls, sheet_name=sheet_names[0])
            if any(str(c).startswith("Unnamed") for c in df_sales.columns[:3]):
                df_sales = pd.read_excel(xls, sheet_name=sheet_names[0], header=1)
                
            df_purchases = pd.read_excel(xls, sheet_name=sheet_names[1])
            if any(str(c).startswith("Unnamed") for c in df_purchases.columns[:3]):
                df_purchases = pd.read_excel(xls, sheet_name=sheet_names[1], header=1)

            df_expenses = pd.read_excel(xls, sheet_name=sheet_names[2])
            if any(str(c).startswith("Unnamed") for c in df_expenses.columns[:3]):
                df_expenses = pd.read_excel(xls, sheet_name=sheet_names[2], header=1)

            # تنظيف أسماء الأعمدة
            df_sales.columns = df_sales.columns.astype(str).str.strip()
            df_purchases.columns = df_purchases.columns.astype(str).str.strip()
            
            # 💡 تأمين شيت المصروفات لأول 3 أعمدة فقط لمنع الانهيار إذا كان الملف يحتوي أعمدة فارغة
            df_expenses = df_expenses.iloc[:, :3]
            df_expenses.columns = ["المبلغ", "بيان المصروف", "مركز التكلفة"]

            # دالة مساعدة للبحث عن الأعمدة
            def guess_col(cols_list, possible_names):
                for name in possible_names:
                    for c in cols_list:
                        if name in str(c): return cols_list.index(c)
                return 0 # الافتراضي

            cols_sales = list(df_sales.columns)

            # 💡 [الواجهة الذكية]: السماح للمستخدم بتصحيح الأعمدة إذا كان الرسم البياني غير منطقي
            with st.expander("⚙️ إعدادات مطابقة الأعمدة (افتح هنا للتأكد من دقة البيانات)", expanded=True):
                st.info("حاول النظام التعرف على الأعمدة تلقائياً. إذا وجدت الرسوم البيانية غير دقيقة، يرجى اختيار الأعمدة الصحيحة من هنا:")
                
                c1, c2, c3 = st.columns(3)
                date_col = c1.selectbox("عمود التاريخ:", cols_sales, index=guess_col(cols_sales, ['التاريخ', 'تاريخ', 'Date']))
                total_col = c2.selectbox("عمود إجمالي المبيعات:", cols_sales, index=guess_col(cols_sales, ['الإجمالي', 'المبلغ', 'Total', 'الاجمالي']))
                qty_col = c3.selectbox("عمود الكمية:", cols_sales, index=guess_col(cols_sales, ['الكمية', 'العدد', 'Qty']))
                
                c4, c5, c6 = st.columns(3)
                product_col = c4.selectbox("عمود اسم المنتج:", cols_sales, index=guess_col(cols_sales, ['المنتج', 'الصنف', 'Product']))
                branch_col = c5.selectbox("عمود الفرع/الصيدلية:", cols_sales, index=guess_col(cols_sales, ['الفرع', 'الصيدلية', 'Branch']))
                user_col = c6.selectbox("عمود الصيدلي/المستخدم:", cols_sales, index=guess_col(cols_sales, ['المستخدم', 'الصيدلي', 'الكاشير', 'User']))
                
                time_col = st.selectbox("عمود الوقت (اختياري، اتركه إذا كان مدمجاً مع التاريخ):", ["غير موجود"] + cols_sales, index=0)

            with st.spinner("جاري معالجة البيانات وبناء النماذج التنبؤية..."):
                # --- معالجة التواريخ والأوقات ---
                df_sales[date_col] = pd.to_datetime(df_sales[date_col], errors='coerce')
                
                if time_col != "غير موجود":
                    df_sales['Hour'] = pd.to_datetime(df_sales[time_col], errors='coerce').dt.hour
                else:
                    df_sales['Hour'] = df_sales[date_col].dt.hour

                df_sales['Hour'] = df_sales['Hour'].fillna(12).astype(int)
                df_sales['فترة البيع'] = df_sales['Hour'].apply(lambda x: 'نهاراً ☀️' if 6 <= x < 18 else 'ليلاً 🌙')

                # --- معالجة الأرقام وتحويلها بشكل آمن (لمنع ظهور أرقام مشوهة) ---
                df_sales[total_col] = pd.to_numeric(df_sales[total_col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                df_sales[qty_col] = pd.to_numeric(df_sales[qty_col].astype(str).str.replace(',', ''), errors='coerce').fillna(1)
                
                pur_cols = list(df_purchases.columns)
                pur_qty_col = pur_cols[guess_col(pur_cols, ['الكمية', 'العدد', 'Qty'])]
                pur_product_col = pur_cols[guess_col(pur_cols, ['المنتج', 'الصنف', 'Product'])]
                pur_total_col = pur_cols[guess_col(pur_cols, ['الإجمالي', 'المبلغ', 'Total'])]

                df_purchases[pur_qty_col] = pd.to_numeric(df_purchases[pur_qty_col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                df_purchases[pur_total_col] = pd.to_numeric(df_purchases[pur_total_col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                df_expenses["المبلغ"] = pd.to_numeric(df_expenses["المبلغ"].astype(str).str.replace(',', ''), errors='coerce').fillna(0)

                # إزالة الصفوف الفارغة للحفاظ على جودة التحليل
                df_sales = df_sales.dropna(subset=[date_col, product_col])

                # ================= واجهة التبويبات =================
                tab1, tab2, tab3, tab4, tab5 = st.tabs([
                    "📊 مؤشرات الأداء", 
                    "⏰ تحليل الأوقات (ساعات الذروة)", 
                    "👥 الفروع، المنتجات والمستخدمين", 
                    "🤖 تنبؤات المبيعات والمستقبل (ML)", 
                    "📦 تقديرات المخزون والطلبيات"
                ])

                with tab1:
                    st.markdown("### 📈 ملخص الأداء المالي")
                    total_sales = df_sales[total_col].sum()
                    total_expenses = df_expenses["المبلغ"].sum()
                    total_purchases = df_purchases[pur_total_col].sum()
                    net_profit = total_sales - total_purchases - total_expenses

                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("إجمالي المبيعات", f"{total_sales:,.2f} ريال")
                    c2.metric("إجمالي المشتريات (تكلفة)", f"{total_purchases:,.2f} ريال")
                    c3.metric("إجمالي المصروفات", f"{total_expenses:,.2f} ريال")
                    c4.metric("صافي الدخل المؤقت", f"{net_profit:,.2f} ريال", delta=float(net_profit))

                    st.markdown("---")
                    st.markdown("### 💸 تحليل المصروفات حسب مركز التكلفة")
                    if not df_expenses.empty and df_expenses["المبلغ"].sum() > 0:
                        exp_summary = df_expenses.groupby("مركز التكلفة")["المبلغ"].sum().reset_index()
                        fig_exp = px.pie(exp_summary, values='المبلغ', names='مركز التكلفة', hole=0.4, title="توزيع المصروفات")
                        st.plotly_chart(fig_exp, use_container_width=True)
                    else:
                        st.info("لا توجد مصروفات مسجلة في الشيت المرفق.")

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

                with tab3:
                    st.markdown("### 🏢 تحليل المنتج مع الفرع (أفضل 15 منتج)")
                    branch_prod = df_sales.pivot_table(index=product_col, columns=branch_col, values=total_col, aggfunc='sum').fillna(0)
                    branch_prod['الإجمالي'] = branch_prod.sum(axis=1)
                    branch_prod = branch_prod.sort_values(by='الإجمالي', ascending=False).head(15).drop(columns=['الإجمالي'])
                    fig_bp = px.imshow(branch_prod, aspect="auto", color_continuous_scale='Viridis')
                    st.plotly_chart(fig_bp, use_container_width=True)

                    st.markdown("---")
                    st.markdown("### 👤 أفضل مبيعات المنتجات لكل صيدلي")
                    user_prod = df_sales.groupby([user_col, product_col])[total_col].sum().reset_index()
                    top_user_prod = user_prod.sort_values(total_col, ascending=False).groupby(user_col).head(3)
                    st.dataframe(top_user_prod, use_container_width=True)

                with tab4:
                    st.markdown("### 🤖 نماذج التنبؤ بالمبيعات والعوائد (Machine Learning)")
                    
                    # 💡 الحل الجذري لمنع خطأ Unnamed: 0 عند عمل التجميع
                    daily_sales = df_sales.groupby(date_col)[total_col].sum().to_frame(name='المبيعات_اليومية').reset_index()
                    daily_sales = daily_sales.sort_values(date_col)
                    
                    if len(daily_sales) > 5:
                        daily_sales['Days_Since_Start'] = (daily_sales[date_col] - daily_sales[date_col].min()).dt.days
                        daily_sales['Day_of_Week'] = daily_sales[date_col].dt.dayofweek
                        
                        X = daily_sales[['Days_Since_Start', 'Day_of_Week']]
                        y = daily_sales['المبيعات_اليومية']
                        
                        model = RandomForestRegressor(n_estimators=100, random_state=42)
                        model.fit(X, y)
                        
                        last_date = daily_sales[date_col].max()
                        future_dates = [last_date + timedelta(days=x) for x in range(1, 31)]
                        future_days_since = [(d - daily_sales[date_col].min()).days for d in future_dates]
                        future_dow = [d.weekday() for d in future_dates]
                        
                        X_future = pd.DataFrame({'Days_Since_Start': future_days_since, 'Day_of_Week': future_dow})
                        predictions = model.predict(X_future)
                        
                        future_df = pd.DataFrame({'التاريخ': future_dates, 'المبيعات': predictions, 'النوع': 'تنبؤ (مستقبل)'})
                        historical_df = pd.DataFrame({'التاريخ': daily_sales[date_col], 'المبيعات': daily_sales['المبيعات_اليومية'], 'النوع': 'بيانات فعلية'})
                        
                        forecast_plot_df = pd.concat([historical_df, future_df])
                        fig_forecast = px.line(forecast_plot_df, x='التاريخ', y='المبيعات', color='النوع',
                                               title="📈 التنبؤ بإجمالي العوائد للـ 30 يوماً القادمة",
                                               color_discrete_map={'بيانات فعلية':'#2980b9', 'تنبؤ (مستقبل)':'#e74c3c'})
                        st.plotly_chart(fig_forecast, use_container_width=True)
                        st.success(f"💰 إجمالي العائد المتوقع للـ 30 يوماً القادمة: **{predictions.sum():,.2f} ريال**")
                        
                        st.markdown("### 🔮 التنبؤ بأداء الفروع القادم (المتوسط اليومي المتوقع)")
                        branch_forecast = []
                        for branch in df_sales[branch_col].unique():
                            b_sales = df_sales[df_sales[branch_col] == branch].groupby(date_col)[total_col].sum().reset_index()
                            if len(b_sales) > 3:
                                b_sales['d'] = (b_sales[date_col] - b_sales[date_col].min()).dt.days
                                lr = LinearRegression()
                                lr.fit(b_sales[['d']], b_sales[total_col])
                                expected_daily = lr.predict(np.array([[b_sales['d'].max() + 15]]))[0]
                                branch_forecast.append({"الفرع": branch, "المتوسط اليومي": max(0, expected_daily)})
                        if branch_forecast:
                            st.dataframe(pd.DataFrame(branch_forecast).style.format({"المتوسط اليومي": "{:.2f}"}), use_container_width=True)
                    else:
                        st.warning("البيانات الزمنية المتاحة غير كافية لبناء نموذج تنبؤ موثوق (نحتاج أكثر من 5 أيام).")

                with tab5:
                    st.markdown("### 📦 تقديرات المخزون والطلبيات")
                    st.info("يقوم النظام بدمج المشتريات وطرح المبيعات لمعرفة (المخزون الدفتري)، ثم يتنبأ بالكمية المطلوبة لـ 30 يوماً بناءً على معدل البيع.")
                    
                    sales_qty_df = df_sales.groupby(product_col)[qty_col].sum().reset_index().rename(columns={qty_col: 'إجمالي المباع'})
                    purchases_qty_df = df_purchases.groupby(pur_product_col)[pur_qty_col].sum().reset_index().rename(columns={pur_product_col: product_col, pur_qty_col: 'إجمالي المشتريات'})
                    
                    inventory_df = pd.merge(purchases_qty_df, sales_qty_df, on=product_col, how='outer').fillna(0)
                    inventory_df['المخزون الحالي'] = inventory_df['إجمالي المشتريات'] - inventory_df['إجمالي المباع']
                    
                    days_in_data = max((df_sales[date_col].max() - df_sales[date_col].min()).days, 1)
                    inventory_df['الاحتياج لـ 30 يوم'] = np.ceil((inventory_df['إجمالي المباع'] / days_in_data) * 30)
                    
                    inventory_df['الكمية المقترح طلبها'] = np.where(
                        inventory_df['الاحتياج لـ 30 يوم'] > inventory_df['المخزون الحالي'],
                        inventory_df['الاحتياج لـ 30 يوم'] - inventory_df['المخزون الحالي'], 0)
                    
                    reorder_df = inventory_df[inventory_df['الكمية المقترح طلبها'] > 0].sort_values('الكمية المقترح طلبها', ascending=False)
                    
                    c1, c2 = st.columns(2)
                    c1.metric("أصناف تحتاج إعادة طلب", f"{len(reorder_df)} صنف")
                    if len(reorder_df) > 0:
                        st.dataframe(reorder_df[[product_col, 'المخزون الحالي', 'الاحتياج لـ 30 يوم', 'الكمية المقترح طلبها']].style.background_gradient(subset=['الكمية المقترح طلبها'], cmap='Reds'), use_container_width=True)
                    else:
                        st.success("المخزون الحالي كافٍ لجميع المنتجات للفترة القادمة.")

        except Exception as e:
            st.error(f"حدث خطأ أثناء معالجة الملف: {e}")
