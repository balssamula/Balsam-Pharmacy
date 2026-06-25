# pages/sales_analysis.py
import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from utils.financial_engine import calculate_financials

@st.cache_data(show_spinner=False)
def generate_bi_excel_report(df_main, prod_profit, rfm_df, manual_expenses):
    """توليد ملف إكسيل BI متكامل ومقسم لشيتات للإدارة"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 1. شيت البيانات المالية الشاملة
        df_main.to_excel(writer, sheet_name="البيانات المالية التفصيلية", index=False)
        
        # 2. شيت ربحية المنتجات
        prod_profit.sort_values('net_profit', ascending=False).to_excel(writer, sheet_name="ربحية المنتجات", index=False)
        
        # 3. شيت المنتجات النازفة (الخاسرة)
        loss_prods = prod_profit[prod_profit['net_profit'] < 0].sort_values('net_profit')
        loss_prods.to_excel(writer, sheet_name="المنتجات الخاسرة", index=False)
        
        # 4. شيت تحليل العملاء (RFM)
        if rfm_df is not None:
            rfm_df.sort_values('إجمالي الإنفاق', ascending=False).to_excel(writer, sheet_name="تحليل العملاء", index=False)
            
        # 5. شيت المصروفات اليدوية
        if manual_expenses:
            pd.DataFrame(manual_expenses).to_excel(writer, sheet_name="المصروفات الإضافية", index=False)
            
    return output.getvalue()

def show():
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@300;400;500;700;800&display=swap');
        * { font-family: 'Tajawal', sans-serif; }
        .kpi-card {background: #ffffff; border-radius: 12px; padding: 24px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); border-right: 5px solid #1f7a8c; margin-bottom: 15px;}
        .kpi-title {color: #6c757d; font-size: 15px; font-weight: 600; margin-bottom: 8px;}
        .kpi-value {color: #16425b; font-size: 30px; font-weight: 800;}
        </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1>📊 لوحة القيادة الذكية - بلسم العلا (BI Dashboard)</h1>", unsafe_allow_html=True)
    
    # ==========================================
    # 1. محطة الرفع والمصروفات
    # ==========================================
    with st.expander("📥 محطة رفع الملفات المالية وإدارة المصروفات", expanded=True):
        st.markdown("### 📂 رفع التقارير الشهرية")
        col1, col2, col3 = st.columns(3)
        col4, col5 = st.columns(2)
        
        with col1: sales_file = st.file_uploader("1️⃣ ملف المبيعات (سلة)", type=["csv", "xlsx"])
        with col2: payment_file = st.file_uploader("2️⃣ تقرير بوابات الدفع (مدى/فيزا)", type=["csv", "xlsx"])
        with col3: tabby_file = st.file_uploader("3️⃣ كشف حساب Tabby", type=["csv", "xlsx"])
        with col4: tamara_file = st.file_uploader("4️⃣ فاتورة Tamara", type=["csv", "xlsx"])
        with col5: jnt_file = st.file_uploader("5️⃣ فاتورة شحن J&T", type=["csv", "xlsx"])
        
        st.markdown("---")
        st.markdown("### 💸 سجل المصروفات التشغيلية والتسويقية الإضافية (OPEX)")
        
        if 'manual_expenses' not in st.session_state:
            st.session_state.manual_expenses = []
            
        c_desc, c_amt, c_btn = st.columns([3, 2, 1])
        with c_desc: exp_desc = st.text_input("بيان المصروف (مثال: رسوم سناب شات، إيجار، رواتب)")
        with c_amt: exp_amt = st.number_input("المبلغ (SAR)", min_value=0.0, step=100.0)
        with c_btn: 
            st.write("") 
            if st.button("➕ إدراج المصروف", use_container_width=True):
                if exp_desc and exp_amt > 0:
                    st.session_state.manual_expenses.append({'desc': exp_desc, 'amount': exp_amt})
                    st.rerun()
                    
        if st.session_state.manual_expenses:
            for i, exp in enumerate(st.session_state.manual_expenses):
                ct, cd = st.columns([5, 1])
                with ct: st.info(f"🏷️ {exp['desc']} | 💰 {exp['amount']:,.2f} ر.س")
                with cd: 
                    if st.button("❌ حذف", key=f"del_{i}"):
                        st.session_state.manual_expenses.pop(i)
                        st.rerun()

    # ==========================================
    # 2. معالجة البيانات وبناء الـ Dashboard
    # ==========================================
    if sales_file:
        with st.spinner("🧠 جاري معالجة ملايين البيانات، وتوزيع التكاليف على المنتجات..."):
            try:
                # قراءة الملفات
                df_sales = pd.read_excel(sales_file) if sales_file.name.endswith('xlsx') else pd.read_csv(sales_file)
                df_pay = pd.read_csv(payment_file) if payment_file and payment_file.name.endswith('csv') else (pd.read_excel(payment_file) if payment_file else pd.DataFrame())
                df_tabby = pd.read_csv(tabby_file) if tabby_file and tabby_file.name.endswith('csv') else (pd.read_excel(tabby_file, skiprows=10) if tabby_file else pd.DataFrame())
                df_tamara = pd.read_csv(tamara_file) if tamara_file and tamara_file.name.endswith('csv') else (pd.read_excel(tamara_file, skiprows=26) if tamara_file else pd.DataFrame())
                df_jnt = pd.read_csv(jnt_file) if jnt_file and jnt_file.name.endswith('csv') else (pd.read_excel(jnt_file, sheet_name="DETAILS") if jnt_file else pd.DataFrame())

                # تشغيل المحرك المالي
                df, total_opex = calculate_financials(df_sales, df_pay, df_tabby, df_tamara, df_jnt, st.session_state.manual_expenses)
                
                # حساب الـ KPIs
                total_revenue = df['product_total'].sum()
                total_cogs = df['total_cost'].sum()
                gross_profit = total_revenue - total_cogs
                operating_profit = df['net_profit'].sum()
                net_profit_final = operating_profit - total_opex
                profit_margin = (net_profit_final / total_revenue * 100) if total_revenue > 0 else 0
                
                st.markdown("---")
                
                # عرض بطاقات الأداء
                mc1, mc2, mc3, mc4 = st.columns(4)
                mc1.markdown(f"<div class='kpi-card'><div class='kpi-title'>إجمالي المبيعات (Revenue)</div><div class='kpi-value'>SAR {total_revenue:,.0f}</div></div>", unsafe_allow_html=True)
                mc2.markdown(f"<div class='kpi-card'><div class='kpi-title'>إجمالي التكلفة (COGS)</div><div class='kpi-value'>SAR {total_cogs:,.0f}</div></div>", unsafe_allow_html=True)
                mc3.markdown(f"<div class='kpi-card'><div class='kpi-title'>صافي الربح الفعلي</div><div class='kpi-value' style='color:{'#2a9d8f' if net_profit_final > 0 else '#e63946'}'>SAR {net_profit_final:,.0f}</div></div>", unsafe_allow_html=True)
                mc4.markdown(f"<div class='kpi-card'><div class='kpi-title'>هامش الربح (Margin)</div><div class='kpi-value'>{profit_margin:,.1f}%</div></div>", unsafe_allow_html=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                # تجهيز جداول التحليل للإكسيل
                prod_profit = df.groupby('product_name').agg({
                    'qty': 'sum',
                    'product_total': 'sum',
                    'net_profit': 'sum'
                }).reset_index()
                
                rfm = None # تهيئة متغير العملاء

                # التبويبات الاحترافية
                tab1, tab2, tab3 = st.tabs(["🛍️ تحليل المنتجات والربحية", "🚚 أداء الشحن والتمركز", "👥 تحليل RFM وسلوك العملاء"])
                
                with tab1:
                    st.subheader("تحليل الربحية الدقيق لكل منتج (بعد خصم كافة التكاليف)")
                    
                    c_win, c_loss = st.columns(2)
                    with c_win:
                        st.success("⭐ أعلى 10 منتجات دراً للربح")
                        top_prods = prod_profit.sort_values('net_profit', ascending=False).head(10)
                        fig_top = px.bar(top_prods, x='net_profit', y='product_name', orientation='h', color_discrete_sequence=['#2a9d8f'])
                        st.plotly_chart(fig_top, use_container_width=True)
                        
                    with c_loss:
                        st.error("⚠️ منتجات تنزف الأرباح (تباع بخسارة بعد خصم العمولات والشحن)")
                        loss_prods = prod_profit[prod_profit['net_profit'] < 0].sort_values('net_profit').head(10)
                        if not loss_prods.empty:
                            fig_loss = px.bar(loss_prods, x='net_profit', y='product_name', orientation='h', color_discrete_sequence=['#e63946'])
                            st.plotly_chart(fig_loss, use_container_width=True)
                        else:
                            st.info("لا يوجد منتجات مباعة بخسارة.")
                            
                with tab2:
                    st.subheader("أين تذهب المبيعات؟ (تحليل مناطق التوصيل وشركات الشحن)")
                    col_city, col_ship = st.columns(2)
                    
                    with col_city:
                        if 'المدينة' in df.columns:
                            city_perf = df.groupby('المدينة').agg({'product_total': 'sum'}).reset_index().sort_values('product_total', ascending=False).head(10)
                            fig_city = px.pie(city_perf, names='المدينة', values='product_total', hole=0.4, title="أعلى 10 مدن في المبيعات")
                            st.plotly_chart(fig_city, use_container_width=True)
                            
                    with col_ship:
                        if 'شركة الشحن' in df.columns:
                            ship_perf = df.groupby('شركة الشحن').agg({'product_total': 'sum'}).reset_index()
                            fig_ship = px.bar(ship_perf, x='شركة الشحن', y='product_total', title="حجم المبيعات حسب شركة الشحن/الفرع", color_discrete_sequence=['#1f7a8c'])
                            st.plotly_chart(fig_ship, use_container_width=True)
                
                with tab3:
                    st.subheader("تحليل العملاء (RFM) - الولاء مقابل خطر الفقدان")
                    date_col = [c for c in df.columns if 'تاريخ' in c or 'Date' in c]
                    name_col = [c for c in df.columns if 'اسم العميل' in c or 'Customer Name' in c]
                    
                    if date_col and name_col:
                        date_c = date_col[0]
                        name_c = name_col[0]
                        
                        df[date_c] = pd.to_datetime(df[date_c], errors='coerce')
                        recent_date = df[date_c].max()
                        
                        rfm = df.groupby(name_c).agg({
                            date_c: lambda x: (recent_date - x.max()).days, # Recency
                            'رقم الطلب': 'nunique', # Frequency
                            'product_total': 'sum' # Monetary
                        }).reset_index()
                        rfm.columns = ['اسم العميل', 'أيام منذ آخر طلب', 'عدد الطلبات', 'إجمالي الإنفاق']
                        
                        # تصنيف العملاء
                        def classify_customer(row):
                            if row['إجمالي الإنفاق'] >= 1000 and row['أيام منذ آخر طلب'] <= 30: return "🌟 عميل VIP نشط"
                            elif row['إجمالي الإنفاق'] >= 500 and row['أيام منذ آخر طلب'] > 30: return "🚨 VIP في خطر (Churn Risk)"
                            elif row['عدد الطلبات'] > 1: return "🔄 عميل متكرر"
                            else: return "👤 عميل جديد/عادي"
                            
                        rfm['تصنيف العميل'] = rfm.apply(classify_customer, axis=1)
                        
                        r1, r2 = st.columns(2)
                        with r1:
                            st.success("🌟 قائمة عملاء الـ VIP (أنفقوا > 1000 ر.س ومستمرون بالطلب)")
                            vip_active = rfm[rfm['تصنيف العميل'] == '🌟 عميل VIP نشط'].sort_values('إجمالي الإنفاق', ascending=False)
                            st.dataframe(vip_active[['اسم العميل', 'عدد الطلبات', 'إجمالي الإنفاق']].head(10), use_container_width=True)
                        
                        with r2:
                            st.error("🚨 عملاء ذو قيمة عالية لم يشتروا منذ فترة (يجب التواصل معهم فوراً)")
                            risk_customers = rfm[rfm['تصنيف العميل'] == '🚨 VIP في خطر (Churn Risk)'].sort_values('إجمالي الإنفاق', ascending=False)
                            st.dataframe(risk_customers[['اسم العميل', 'أيام منذ آخر طلب', 'إجمالي الإنفاق']].head(10), use_container_width=True)
                            
                    else:
                        st.warning("لم يتم العثور على أعمدة 'تاريخ الطلب' أو 'اسم العميل' في الملف المرفوع.")
                
                # ==========================================
                # 3. زر تحميل التقرير النهائي (BI Excel Report)
                # ==========================================
                st.markdown("---")
                st.subheader("💾 استخراج التقارير الإدارية (BI Report)")
                st.info("هذا الزر يقوم بتوليد ملف Excel احترافي يحتوي على كافة التحليلات، الأرباح، المنتجات الخاسرة، وقاعدة بيانات العملاء مقسمة في شيتات جاهزة للطباعة والعرض على الإدارة.")
                
                excel_bytes = generate_bi_excel_report(df, prod_profit, rfm, st.session_state.manual_expenses)
                st.download_button(
                    label="📥 تحميل تقرير ذكاء الأعمال المالي (Excel)",
                    data=excel_bytes,
                    file_name="Balsam_BI_Financial_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"حدث خطأ فني أثناء مطابقة الملفات: {str(e)}")
    else:
        st.info("📂 يرجى رفع ملف المبيعات (سلة) كحد أدنى لبدء لوحة ذكاء الأعمال.")
