import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import io
from datetime import timedelta
from sklearn.linear_model import LinearRegression
from sklearn.ensemble import RandomForestRegressor

def export_to_excel(dfs_dict):
    """دالة لتحويل مجموعة من DataFrames إلى ملف Excel متعدد الشيتات"""
    output = io.BytesIO()
    # نستخدم 엔진 xlsxwriter المرفق في ملف المتطلبات الخاص بك
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df in dfs_dict.items():
            if df is not None and not df.empty:
                # التأكد من أن اسم الشيت لا يتجاوز 31 حرف (قيد في Excel)
                safe_sheet_name = str(sheet_name)[:31]
                df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                
                # تنسيق العرض (توسيع الأعمدة قليلاً لتكون مقروءة)
                worksheet = writer.sheets[safe_sheet_name]
                for i, col in enumerate(df.columns):
                    column_len = max(df[col].astype(str).map(len).max(), len(str(col))) + 2
                    worksheet.set_column(i, i, min(column_len, 40))
    return output.getvalue()

def show():
    st.markdown("""
    <div class="hero">
        <h1>📈 التحليل الشامل والتنبؤات للصيدليات</h1>
        <p>تحليل المبيعات، المشتريات، المصروفات، وتصدير التقارير الكاملة</p>
    </div>
    """, unsafe_allow_html=True)

    uploaded_file = st.file_uploader("📂 ارفع ملف Excel (المبيعات، المشتريات، المصروفات)", type=["xlsx", "xls"])

    if uploaded_file:
        try:
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = xls.sheet_names
            
            # قراءة ذكية للشيتات (تخطي الصف الأول لو كان فارغاً)
            df_sales = pd.read_excel(xls, sheet_name=sheet_names[0])
            if any(str(c).startswith("Unnamed") for c in df_sales.columns[:3]):
                df_sales = pd.read_excel(xls, sheet_name=sheet_names[0], header=1)
                
            df_purchases = pd.read_excel(xls, sheet_name=sheet_names[1])
            if any(str(c).startswith("Unnamed") for c in df_purchases.columns[:3]):
                df_purchases = pd.read_excel(xls, sheet_name=sheet_names[1], header=1)

            df_expenses = pd.read_excel(xls, sheet_name=sheet_names[2])
            if any(str(c).startswith("Unnamed") for c in df_expenses.columns[:3]):
                df_expenses = pd.read_excel(xls, sheet_name=sheet_names[2], header=1)

            # تنظيف الأعمدة
            df_sales.columns = df_sales.columns.astype(str).str.strip()
            df_purchases.columns = df_purchases.columns.astype(str).str.strip()
            df_expenses = df_expenses.iloc[:, :3]
            df_expenses.columns = ["المبلغ", "بيان المصروف", "مركز التكلفة"]

            def guess_col(cols_list, possible_names):
                for name in possible_names:
                    for c in cols_list:
                        if name in str(c): return cols_list.index(c)
                return 0 

            cols_sales = list(df_sales.columns)

            with st.expander("⚙️ إعدادات مطابقة الأعمدة", expanded=False):
                c1, c2, c3 = st.columns(3)
                date_col = c1.selectbox("عمود التاريخ:", cols_sales, index=guess_col(cols_sales, ['التاريخ', 'تاريخ', 'Date']))
                total_col = c2.selectbox("عمود الإجمالي:", cols_sales, index=guess_col(cols_sales, ['الإجمالي', 'المبلغ', 'Total', 'الاجمالي']))
                qty_col = c3.selectbox("عمود الكمية:", cols_sales, index=guess_col(cols_sales, ['الكمية', 'العدد', 'Qty']))
                
                c4, c5, c6 = st.columns(3)
                product_col = c4.selectbox("عمود المنتج:", cols_sales, index=guess_col(cols_sales, ['المنتج', 'الصنف', 'Product']))
                branch_col = c5.selectbox("عمود الفرع:", cols_sales, index=guess_col(cols_sales, ['الفرع', 'الصيدلية', 'Branch']))
                user_col = c6.selectbox("عمود الصيدلي:", cols_sales, index=guess_col(cols_sales, ['المستخدم', 'الصيدلي', 'الكاشير', 'User']))
                
                time_col = st.selectbox("عمود الوقت:", ["غير موجود"] + cols_sales, index=0)

            with st.spinner("جاري معالجة البيانات وبناء النماذج التنبؤية..."):
                # ----------------- معالجة البيانات -----------------
                df_sales[date_col] = pd.to_datetime(df_sales[date_col], errors='coerce')
                
                if time_col != "غير موجود":
                    df_sales['Hour'] = pd.to_datetime(df_sales[time_col], errors='coerce').dt.hour
                else:
                    df_sales['Hour'] = df_sales[date_col].dt.hour

                df_sales['Hour'] = df_sales['Hour'].fillna(12).astype(int)
                df_sales['فترة البيع'] = df_sales['Hour'].apply(lambda x: 'نهاراً ☀️' if 6 <= x < 18 else 'ليلاً 🌙')

                df_sales[total_col] = pd.to_numeric(df_sales[total_col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                df_sales[qty_col] = pd.to_numeric(df_sales[qty_col].astype(str).str.replace(',', ''), errors='coerce').fillna(1)
                
                pur_cols = list(df_purchases.columns)
                pur_qty_col = pur_cols[guess_col(pur_cols, ['الكمية', 'العدد', 'Qty'])]
                pur_product_col = pur_cols[guess_col(pur_cols, ['المنتج', 'الصنف', 'Product'])]
                pur_total_col = pur_cols[guess_col(pur_cols, ['الإجمالي', 'المبلغ', 'Total'])]

                df_purchases[pur_qty_col] = pd.to_numeric(df_purchases[pur_qty_col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                df_purchases[pur_total_col] = pd.to_numeric(df_purchases[pur_total_col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                df_expenses["المبلغ"] = pd.to_numeric(df_expenses["المبلغ"].astype(str).str.replace(',', ''), errors='coerce').fillna(0)

                df_sales = df_sales.dropna(subset=[date_col, product_col])

                # ================= 1. ملخص الأداء =================
                total_sales = df_sales[total_col].sum()
                total_expenses = df_expenses["المبلغ"].sum()
                total_purchases = df_purchases[pur_total_col].sum()
                net_profit = total_sales - total_purchases - total_expenses
                
                df_financial = pd.DataFrame([{
                    "إجمالي المبيعات": total_sales,
                    "إجمالي المشتريات": total_purchases,
                    "إجمالي المصروفات": total_expenses,
                    "صافي الدخل المؤقت": net_profit
                }])
                
                exp_summary = df_expenses.groupby("مركز التكلفة")["المبلغ"].sum().reset_index()

                # ================= 2. تحليل الأوقات =================
                time_sales = df_sales.groupby('فترة البيع')[total_col].sum().reset_index()
                hourly_sales = df_sales.groupby('Hour')[total_col].sum().reset_index()
                
                top_products = df_sales.groupby(product_col)[total_col].sum().nlargest(10).index
                prod_time_matrix = df_sales[df_sales[product_col].isin(top_products)].pivot_table(
                    index=product_col, columns='فترة البيع', values=total_col, aggfunc='sum'
                ).fillna(0).reset_index()

                # ================= 3. تفصيلي الفروع والمستخدمين =================
                # بناء التقرير الشامل الذي طلبه المستخدم (الفرع + الصيدلي + المنتج)
                df_comprehensive = df_sales.groupby([branch_col, user_col, product_col]).agg(
                    إجمالي_المبيعات=(total_col, 'sum'),
                    إجمالي_الكمية=(qty_col, 'sum')
                ).reset_index().sort_values(by=[branch_col, 'إجمالي_المبيعات'], ascending=[True, False])

                branch_prod = df_sales.pivot_table(index=product_col, columns=branch_col, values=total_col, aggfunc='sum').fillna(0)
                branch_prod['الإجمالي'] = branch_prod.sum(axis=1)
                branch_prod = branch_prod.sort_values(by='الإجمالي', ascending=False)
                df_branch_prod_export = branch_prod.reset_index()

                user_prod = df_sales.groupby([user_col, product_col])[total_col].sum().reset_index()
                top_user_prod = user_prod.sort_values(total_col, ascending=False).groupby(user_col).head(3)

                # ================= 4. التنبؤات (ML) =================
                daily_sales = df_sales.groupby(date_col)[total_col].sum().to_frame(name='المبيعات').reset_index().sort_values(date_col)
                df_forecast = pd.DataFrame()
                df_branch_forecast = pd.DataFrame()
                forecast_plot_df = pd.DataFrame()

                if len(daily_sales) > 5:
                    daily_sales['Days_Since_Start'] = (daily_sales[date_col] - daily_sales[date_col].min()).dt.days
                    daily_sales['Day_of_Week'] = daily_sales[date_col].dt.dayofweek
                    
                    X = daily_sales[['Days_Since_Start', 'Day_of_Week']]
                    y = daily_sales['المبيعات']
                    
                    model = RandomForestRegressor(n_estimators=100, random_state=42)
                    model.fit(X, y)
                    
                    last_date = daily_sales[date_col].max()
                    future_dates = [last_date + timedelta(days=x) for x in range(1, 31)]
                    future_days_since = [(d - daily_sales[date_col].min()).days for d in future_dates]
                    future_dow = [d.weekday() for d in future_dates]
                    
                    predictions = model.predict(pd.DataFrame({'Days_Since_Start': future_days_since, 'Day_of_Week': future_dow}))
                    df_forecast = pd.DataFrame({'التاريخ': future_dates, 'المبيعات_المتوقعة': predictions, 'النوع': 'تنبؤ (مستقبل)'})
                    
                    historical_df = pd.DataFrame({'التاريخ': daily_sales[date_col], 'المبيعات_المتوقعة': daily_sales['المبيعات'], 'النوع': 'بيانات فعلية'})
                    forecast_plot_df = pd.concat([historical_df, df_forecast])
                    
                    b_forecast_list = []
                    for branch in df_sales[branch_col].unique():
                        b_sales = df_sales[df_sales[branch_col] == branch].groupby(date_col)[total_col].sum().reset_index()
                        if len(b_sales) > 3:
                            b_sales['d'] = (b_sales[date_col] - b_sales[date_col].min()).dt.days
                            lr = LinearRegression()
                            lr.fit(b_sales[['d']], b_sales[total_col])
                            expected_daily = lr.predict(np.array([[b_sales['d'].max() + 15]]))[0]
                            b_forecast_list.append({"الفرع": branch, "المتوسط اليومي المتوقع": max(0, expected_daily)})
                    df_branch_forecast = pd.DataFrame(b_forecast_list)

                # ================= 5. المخزون والطلبيات =================
                sales_qty_df = df_sales.groupby(product_col)[qty_col].sum().reset_index().rename(columns={qty_col: 'إجمالي المباع'})
                purchases_qty_df = df_purchases.groupby(pur_product_col)[pur_qty_col].sum().reset_index().rename(columns={pur_product_col: product_col, pur_qty_col: 'إجمالي المشتريات'})
                
                inventory_df = pd.merge(purchases_qty_df, sales_qty_df, on=product_col, how='outer').fillna(0)
                inventory_df['المخزون الحالي'] = inventory_df['إجمالي المشتريات'] - inventory_df['إجمالي المباع']
                days_in_data = max((df_sales[date_col].max() - df_sales[date_col].min()).days, 1)
                inventory_df['الاحتياج لـ 30 يوم'] = np.ceil((inventory_df['إجمالي المباع'] / days_in_data) * 30)
                inventory_df['الكمية المقترح طلبها'] = np.where(inventory_df['الاحتياج لـ 30 يوم'] > inventory_df['المخزون الحالي'], inventory_df['الاحتياج لـ 30 يوم'] - inventory_df['المخزون الحالي'], 0)
                reorder_df = inventory_df[inventory_df['الكمية المقترح طلبها'] > 0].sort_values('الكمية المقترح طلبها', ascending=False)

                # ================= زر التحميل (Excel Export) =================
                st.markdown("### 📥 تصدير جميع التقارير والتنبؤات")
                excel_sheets = {
                    "الملخص المالي": df_financial,
                    "مبيعات الصيدلي والفرع (شامل)": df_comprehensive,
                    "نواقص المخزون والطلبيات": reorder_df,
                    "تنبؤات المستقبل (30 يوم)": df_forecast,
                    "تنبؤات الفروع": df_branch_forecast,
                    "المصروفات": exp_summary,
                    "المنتجات بالفرع": df_branch_prod_export,
                    "ساعات الذروة": hourly_sales
                }
                excel_data = export_to_excel(excel_sheets)
                
                st.download_button(
                    label="💾 تحميل التقرير الشامل (Excel)",
                    data=excel_data,
                    file_name="Comprehensive_Pharmacy_Analysis.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
                st.markdown("---")

                # ================= عرض الواجهة (Tabs) =================
                tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 مؤشرات الأداء", "⏰ أوقات الذروة", "👥 فروع ومستخدمين", "🤖 تنبؤات المبيعات", "📦 المخزون"])

                with tab1:
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("إجمالي المبيعات", f"{total_sales:,.2f} ريال")
                    c2.metric("إجمالي المشتريات", f"{total_purchases:,.2f} ريال")
                    c3.metric("المصروفات", f"{total_expenses:,.2f} ريال")
                    c4.metric("صافي الدخل المؤقت", f"{net_profit:,.2f} ريال", delta=float(net_profit))
                    if not exp_summary.empty:
                        st.plotly_chart(px.pie(exp_summary, values='المبلغ', names='مركز التكلفة', hole=0.4), use_container_width=True)

                with tab2:
                    c1, c2 = st.columns(2)
                    c1.plotly_chart(px.pie(time_sales, values=total_col, names='فترة البيع', color='فترة البيع', color_discrete_map={'نهاراً ☀️':'#f1c40f', 'ليلاً 🌙':'#2c3e50'}), use_container_width=True)
                    c2.plotly_chart(px.bar(hourly_sales, x='Hour', y=total_col, text_auto='.2s'), use_container_width=True)
                    st.dataframe(prod_time_matrix, use_container_width=True)

                with tab3:
                    st.markdown("### 📋 تفصيلي مبيعات الفروع والصيادلة لكل منتج")
                    st.dataframe(df_comprehensive, use_container_width=True)
                    
                    st.markdown("### 🔥 خريطة حرارية: المنتجات بالفرع")
                    st.plotly_chart(px.imshow(branch_prod.head(15).drop(columns=['الإجمالي']), aspect="auto"), use_container_width=True)

                with tab4:
                    if not forecast_plot_df.empty:
                        st.plotly_chart(px.line(forecast_plot_df, x='التاريخ', y='المبيعات_المتوقعة', color='النوع', title="التنبؤ لـ 30 يوم", color_discrete_map={'بيانات فعلية':'#2980b9', 'تنبؤ (مستقبل)':'#e74c3c'}), use_container_width=True)
                        if not df_branch_forecast.empty:
                            st.dataframe(df_branch_forecast.style.format({"المتوسط اليومي المتوقع": "{:.2f}"}), use_container_width=True)
                    else:
                        st.warning("البيانات الزمنية غير كافية لعمل تنبؤات موثوقة.")

                with tab5:
                    c1, c2 = st.columns(2)
                    c1.metric("أصناف تحتاج إعادة طلب", f"{len(reorder_df)} صنف")
                    if not reorder_df.empty:
                        st.dataframe(reorder_df.style.background_gradient(subset=['الكمية المقترح طلبها'], cmap='Reds'), use_container_width=True)
                    else:
                        st.success("المخزون الحالي كافٍ لجميع المنتجات.")

        except Exception as e:
            st.error(f"حدث خطأ أثناء المعالجة: {e}")
