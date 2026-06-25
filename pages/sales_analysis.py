import streamlit as st
import pandas as pd
import plotly.express as px
from utils.financial_engine import calculate_financials

def show():
    st.markdown("""
        <style>
        .metric-card {background-color: #f8f9fa; border-radius: 10px; padding: 20px; border-right: 5px solid #1f7a8c; box-shadow: 0 4px 6px rgba(0,0,0,0.05);}
        .metric-title {color: #6c757d; font-size: 14px; font-weight: 600; margin-bottom: 8px;}
        .metric-value {color: #2b2d42; font-size: 28px; font-weight: 800;}
        </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1>📊 نظام ذكاء الأعمال (BI) - لوحة الأداء الشامل</h1>", unsafe_allow_html=True)
    
    # ==========================================
    # 1. نظام رفع الملفات وإدخال المصروفات
    # ==========================================
    with st.expander("📥 محطة رفع البيانات المالية وإدارة المصروفات", expanded=True):
        st.markdown("### 📄 رفع التقارير الشهرية")
        c1, c2, c3 = st.columns(3)
        with c1: sales_file = st.file_uploader("ملف المبيعات (سلة)", type=["csv", "xlsx"])
        with c2: payment_file = st.file_uploader("ملف بوابات الدفع (مدى/تمارا/تابي)", type=["csv", "xlsx"])
        with c3: shipping_file = st.file_uploader("ملف فواتير الشحن (أرامكس/بيز...)", type=["csv", "xlsx", "pdf"])
        
        st.markdown("---")
        st.markdown("### 💸 سجل المصروفات التشغيلية والتسويقية الإضافية")
        
        # إدارة المصروفات الديناميكية باستخدام session_state
        if 'manual_expenses' not in st.session_state:
            st.session_state.manual_expenses = []
            
        col_desc, col_amt, col_btn = st.columns([3, 2, 1])
        with col_desc: exp_desc = st.text_input("بيان المصروف (مثال: إعلانات سناب شات، إيجار مستودع)")
        with col_amt: exp_amt = st.number_input("المبلغ (ريال)", min_value=0.0, step=100.0)
        with col_btn: 
            st.write("") # للمحاذاة
            if st.button("➕ إضافة المصروف", use_container_width=True):
                if exp_desc and exp_amt > 0:
                    st.session_state.manual_expenses.append({'desc': exp_desc, 'amount': exp_amt})
                    st.rerun()
                    
        # عرض المصروفات المضافة
        if st.session_state.manual_expenses:
            st.markdown("**قائمة المصروفات المضافة للشهر الحالي:**")
            for i, exp in enumerate(st.session_state.manual_expenses):
                col_text, col_del = st.columns([5, 1])
                with col_text: st.info(f"🏷️ {exp['desc']} | 💰 {exp['amount']:,.2f} ر.س")
                with col_del: 
                    if st.button("❌ حذف", key=f"del_exp_{i}"):
                        st.session_state.manual_expenses.pop(i)
                        st.rerun()

    # ==========================================
    # 2. معالجة البيانات وبناء اللوحة
    # ==========================================
    if sales_file:
        with st.spinner("🧠 جاري تشغيل محرك الحسابات المالية ودمج المصروفات..."):
            try:
                # قراءة الملفات
                df_sales = pd.read_csv(sales_file) if sales_file.name.endswith('csv') else pd.read_excel(sales_file)
                df_pay = pd.read_csv(payment_file) if payment_file and payment_file.name.endswith('csv') else (pd.read_excel(payment_file) if payment_file else pd.DataFrame())
                df_ship = pd.read_csv(shipping_file) if shipping_file and shipping_file.name.endswith('csv') else (pd.read_excel(shipping_file) if shipping_file else pd.DataFrame())
                
                # تنفيذ المحرك المالي
                df, total_opex = calculate_financials(df_sales, df_pay, df_ship, st.session_state.manual_expenses)
                
                # حساب الـ KPIs
                total_revenue = df['product_total'].sum()
                total_cogs = df['cost'].sum()
                gross_profit = total_revenue - total_cogs
                operating_profit = df['net_profit'].sum()
                net_profit_final = operating_profit - total_opex
                profit_margin = (net_profit_final / total_revenue * 100) if total_revenue > 0 else 0
                
                st.markdown("---")
                
                # عرض المؤشرات العليا
                mc1, mc2, mc3, mc4 = st.columns(4)
                mc1.markdown(f"<div class='metric-card'><div class='metric-title'>إجمالي الإيرادات</div><div class='metric-value'>SAR {total_revenue:,.0f}</div></div>", unsafe_allow_html=True)
                mc2.markdown(f"<div class='metric-card'><div class='metric-title'>صافي الربح النهائي</div><div class='metric-value' style='color:{'#2a9d8f' if net_profit_final > 0 else '#e63946'}'>SAR {net_profit_final:,.0f}</div></div>", unsafe_allow_html=True)
                mc3.markdown(f"<div class='metric-card'><div class='metric-title'>هامش الربح التشغيلي</div><div class='metric-value'>{profit_margin:,.1f}%</div></div>", unsafe_allow_html=True)
                mc4.markdown(f"<div class='metric-card'><div class='metric-title'>إجمالي المصروفات التشغيلية</div><div class='metric-value'>SAR {total_opex + df['marketing_commission'].sum():,.0f}</div></div>", unsafe_allow_html=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                # التبويبات الاحترافية
                tab1, tab2, tab3 = st.tabs(["🛍️ أداء المنتجات والربحية", "🗺️ تحليل الفروع والنمو", "👥 سلوك العملاء والولاء"])
                
                with tab1:
                    st.subheader("المنتجات الأكثر دراً للربح مقابل المنتجات النازفة")
                    prod_profit = df.groupby('product_name').agg({
                        'qty': 'sum',
                        'product_total': 'sum',
                        'net_profit': 'sum'
                    }).reset_index()
                    
                    c_win, c_loss = st.columns(2)
                    with c_win:
                        st.success("⭐ أعلى 5 منتجات ربحية (نجوم المبيعات)")
                        top_prods = prod_profit.sort_values('net_profit', ascending=False).head(5)
                        fig_top = px.bar(top_prods, x='net_profit', y='product_name', orientation='h', color_discrete_sequence=['#2a9d8f'])
                        st.plotly_chart(fig_top, use_container_width=True)
                        
                    with c_loss:
                        st.error("⚠️ منتجات تنزف الربح (تباع بخسارة وتحتاج إيقاف عروضها)")
                        loss_prods = prod_profit[prod_profit['net_profit'] < 0].sort_values('net_profit').head(5)
                        if not loss_prods.empty:
                            fig_loss = px.bar(loss_prods, x='net_profit', y='product_name', orientation='h', color_discrete_sequence=['#e63946'])
                            st.plotly_chart(fig_loss, use_container_width=True)
                        else:
                            st.info("ممتاز! لا يوجد منتجات تباع بخسارة.")
                            
                with tab2:
                    st.subheader("قياس أداء التوصيل والنمو حسب الفروع")
                    if 'شركة الشحن' in df.columns:
                        branch_perf = df.groupby('شركة الشحن').agg({'product_total': 'sum', 'net_profit':'sum'}).reset_index()
                        fig_branch = px.pie(branch_perf, names='شركة الشحن', values='product_total', hole=0.4)
                        st.plotly_chart(fig_branch, use_container_width=True)
                        st.caption("💡 تلميح: إذا تراجعت نسبة (استلام من الفرع) مقابل الشحن الخارجي، راجع أداء موظفي المبيعات داخل فروع العلا.")
                
                with tab3:
                    st.subheader("تحليل العملاء وخطر الفقدان (Churn Risk)")
                    if 'اسم العميل' in df.columns and 'تاريخ الطلب' in df.columns:
                        df['تاريخ الطلب'] = pd.to_datetime(df['تاريخ الطلب'], errors='coerce')
                        recent_date = df['تاريخ الطلب'].max()
                        
                        rfm = df.groupby('اسم العميل').agg({
                            'تاريخ الطلب': lambda x: (recent_date - x.max()).days, # Recency
                            'رقم الطلب': 'nunique', # Frequency
                            'product_total': 'sum' # Monetary
                        }).reset_index()
                        rfm.columns = ['اسم العميل', 'أيام منذ آخر طلب', 'عدد الطلبات', 'إجمالي الإنفاق']
                        
                        risk_customers = rfm[(rfm['أيام منذ آخر طلب'] > 30) & (rfm['إجمالي الإنفاق'] > 500)]
                        st.warning(f"🚨 يوجد {len(risk_customers)} عميل ذو قيمة عالية (أنفقوا > 500 ريال) لم يشتروا منذ أكثر من 30 يوماً. يجب إرسال رسائل استرجاع لهم.")
                        st.dataframe(risk_customers.sort_values('إجمالي الإنفاق', ascending=False).head(10))

            except Exception as e:
                st.error(f"حدث خطأ أثناء معالجة البيانات: {str(e)}")
    else:
        st.info("يرجى رفع ملف المبيعات للبدء في تحليل ذكاء الأعمال.")
