import streamlit as st
import pandas as pd
import plotly.express as px
from utils.financial_engine import calculate_financials, export_advanced_excel

def show():
    st.set_page_config(layout="wide", page_title="نظام ذكاء الأعمال ERP - بلسم العلا")
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700&display=swap');
        * { font-family: 'Tajawal', sans-serif; }
        .kpi {background: #fff; padding: 20px; border-radius: 10px; border-top: 5px solid #1f7a8c; box-shadow: 0 4px 6px rgba(0,0,0,0.1); text-align:center;}
        .kpi-title {color: #888; font-size: 14px;}
        .kpi-val {color: #111; font-size: 28px; font-weight: bold;}
        .alert-box {padding: 15px; border-radius: 8px; margin-bottom: 10px; font-weight: bold;}
        .alert-danger {background-color: #ffe5e5; color: #d63031; border-left: 5px solid #d63031;}
        .alert-success {background-color: #e5f9e5; color: #27ae60; border-left: 5px solid #27ae60;}
        </style>
    """, unsafe_allow_html=True)

    st.title("🚀 لوحة القيادة الاستراتيجية الشاملة (Strategic KPI Dashboard)")

    with st.expander("📥 1. مركز معالجة البيانات (Data Ingestion)", expanded=True):
        c1, c2, c3 = st.columns(3)
        sales_file = c1.file_uploader("📦 مبيعات سلة", type=["xlsx", "csv"])
        profiles_file = c2.file_uploader("💊 فواتير البروفايلات (لحساب التكلفة)", type=["xlsx", "csv"])
        payment_file = c3.file_uploader("💳 بوابات الدفع (مدى/فيزا)", type=["xlsx", "csv"])
        
        c4, c5, c6 = st.columns(3)
        tabby_file = c4.file_uploader("🟢 تابي", type=["xlsx", "csv"])
        tamara_file = c5.file_uploader("🟣 تمارا", type=["xlsx", "csv"])
        emkan_file = c6.file_uploader("🔵 إمكان (يدوي برقم الطلب)", type=["xlsx", "csv"])
        
        c7, c8, c9 = st.columns(3)
        jnt_file = c7.file_uploader("🚚 J&T", type=["xlsx", "csv"])
        aramex_file = c8.file_uploader("🚚 Aramex", type=["xlsx", "csv", "pdf"])
        beez_file = c9.file_uploader("🚚 Beez", type=["xlsx", "csv"])

        st.markdown("---")
        if 'manual_exp' not in st.session_state: st.session_state.manual_exp = []
        c_n, c_a, c_b = st.columns([3,2,1])
        with c_n: exp_name = st.text_input("بيان المصروف الثابت (إيجار/تسويق)")
        with c_a: exp_amt = st.number_input("المبلغ", min_value=0.0)
        with c_b: 
            st.write(""); 
            if st.button("إضافة"): st.session_state.manual_exp.append({'desc': exp_name, 'amount': exp_amt}); st.rerun()

        if st.session_state.manual_exp:
            for i, exp in enumerate(st.session_state.manual_exp):
                col_txt, col_del = st.columns([5,1])
                col_txt.info(f"🏷️ {exp['desc']} | 💰 {exp['amount']:,.2f} ر.س")
                if col_del.button("❌", key=f"del_{i}"): st.session_state.manual_exp.pop(i); st.rerun()

    if sales_file:
        with st.spinner("جاري دمج وتحليل ملايين السجلات..."):
            
            # 🧠 [التعديل الجذري]: دالة قراءة ذكية محصنة من أخطاء الـ TypeError وملفات الـ PDF
            def load_df(f, **kwargs):
                if not f: return None
                fname = f.name.lower()
                try:
                    if fname.endswith('.xlsx') or fname.endswith('.xls'):
                        return pd.read_excel(f, **kwargs)
                    elif fname.endswith('.csv'):
                        kwargs.pop('sheet_name', None) # 👈 مسح خاصية الشيت إذا كان الملف CSV لمنع الانهيار
                        return pd.read_csv(f, **kwargs)
                    else:
                        return None # تجاهل ملفات PDF أو الصور بأمان تام
                except Exception as e:
                    st.warning(f"لم نتمكن من قراءة الملف {f.name} آلياً، يرجى التأكد من صيغته.")
                    return None
            
            try:
                df_sales = load_df(sales_file)
                df_prof = load_df(profiles_file)
                df_pay = load_df(payment_file)
                df_tabby = load_df(tabby_file, skiprows=10)
                df_tamara = load_df(tamara_file, skiprows=26)
                df_emkan = load_df(emkan_file)
                df_jnt = load_df(jnt_file, sheet_name="DETAILS")
                df_aramex = load_df(aramex_file)
                df_beez = load_df(beez_file)

                # تشغيل المحرك المالي
                df, total_opex = calculate_financials(df_sales, df_prof, df_pay, df_tabby, df_tamara, df_emkan, df_jnt, df_aramex, df_beez, st.session_state.manual_exp)
                
                # إجمالي العمليات
                total_rev = df['product_total'].sum()
                total_gateway = df['gateway_fee'].sum()
                total_ship = df['shipping_cost'].sum()
                total_expenses = total_opex + total_gateway + total_ship + df['marketing_commission'].sum()
                net_profit = df['net_profit'].sum() - total_opex
                margin = (net_profit / total_rev * 100) if total_rev else 0

                # --- 1. KPI Dashboard ---
                st.markdown("### 📊 المؤشرات العليا (KPIs)")
                k1, k2, k3, k4, k5 = st.columns(5)
                k1.markdown(f"<div class='kpi'><div class='kpi-title'>إجمالي المبيعات</div><div class='kpi-val'>SAR {total_rev:,.0f}</div></div>", unsafe_allow_html=True)
                k2.markdown(f"<div class='kpi'><div class='kpi-title'>إجمالي المصروفات</div><div class='kpi-val'>SAR {total_expenses:,.0f}</div></div>", unsafe_allow_html=True)
                k3.markdown(f"<div class='kpi'><div class='kpi-title'>صافي الربح</div><div class='kpi-val' style='color:{'#27ae60' if net_profit>0 else '#d63031'}'>SAR {net_profit:,.0f}</div></div>", unsafe_allow_html=True)
                k4.markdown(f"<div class='kpi'><div class='kpi-title'>هامش الربح</div><div class='kpi-val'>{margin:.1f}%</div></div>", unsafe_allow_html=True)
                
                delivery_growth = len(df[df['المدينة'] != 'العلا']) / len(df) * 100 if 'المدينة' in df.columns else 0
                k5.markdown(f"<div class='kpi'><div class='kpi-title'>مبيعات خارج العلا</div><div class='kpi-val'>{delivery_growth:.1f}%</div></div>", unsafe_allow_html=True)

                # --- 7. Proactive Alerts ---
                st.markdown("### 🚨 الإنذارات الاستراتيجية الآلية")
                if delivery_growth > 15:
                    st.markdown("<div class='alert-box alert-success'>🎯 فرصة تسويقية: طلبات التوصيل الخارجي ممتازة (>15%). نوصي بزيادة ميزانية الإعلانات الرقمية في هذه المدن!</div>", unsafe_allow_html=True)
                if (df['net_profit'] < 0).any():
                    st.markdown(f"<div class='alert-box alert-danger'>⚠️ نزيف أموال: يوجد {len(df[df['net_profit'] < 0]['product_name'].unique())} منتجات تباع بخسارة! يجب إزالة العروض عنها فوراً.</div>", unsafe_allow_html=True)

                tab1, tab2, tab3, tab4 = st.tabs(["👥 العملاء (RFM)", "📦 أداء المنتجات", "🚚 المصروفات (شحن/دفع)", "📤 تصدير BI"])

                with tab1:
                    st.subheader("تحليل العملاء (RFM) ومخاطر الفقدان")
                    if 'تاريخ الطلب' in df.columns and 'اسم العميل' in df.columns:
                        rfm = df.groupby('اسم العميل').agg({
                            'تاريخ الطلب': lambda x: (df['تاريخ الطلب'].max() - x.max()).days,
                            'رقم الطلب': 'nunique',
                            'product_total': 'sum'
                        }).reset_index()
                        rfm.columns = ['العميل', 'أيام الانقطاع', 'الطلبات', 'الإنفاق']
                        
                        c1, c2 = st.columns(2)
                        c1.success("🌟 عملاء VIP (للرعاية الخاصة)")
                        c1.dataframe(rfm[(rfm['الإنفاق'] > 1000) & (rfm['أيام الانقطاع'] <= 30)].sort_values('الإنفاق', ascending=False).head(10))
                        
                        c2.error("🚨 عملاء في خطر الفقدان (يجب الاتصال بهم!)")
                        c2.dataframe(rfm[(rfm['الإنفاق'] > 500) & (rfm['أيام الانقطاع'] > 45)].sort_values('الإنفاق', ascending=False).head(10))

                with tab2:
                    st.subheader("تحليل ربحية المنتجات (Profitability)")
                    c1, c2 = st.columns(2)
                    prod_stats = df.groupby('product_name').agg({'product_total':'sum', 'net_profit':'sum', 'momentum':'mean'}).reset_index()
                    
                    c1.markdown("#### ✅ المنتجات التي تجلب الربح الحقيقي")
                    c1.dataframe(prod_stats.sort_values('net_profit', ascending=False).head(10))
                    
                    c2.markdown("#### 🩸 المنتجات النازفة (تباع بخسارة)")
                    c2.dataframe(prod_stats[prod_stats['net_profit'] < 0].sort_values('net_profit').head(10))

                with tab3:
                    st.subheader("تحليل المصروفات (مُطابق لتقارير الإدارة)")
                    c1, c2 = st.columns(2)
                    
                    with c1:
                        st.markdown("#### 🚚 تحليل شركات الشحن")
                        if 'شركة الشحن' in df.columns:
                            ship_stats = df.groupby('شركة الشحن').agg(
                                الطلبات=('رقم الطلب', 'nunique'),
                                قيمة_الطلبات=('product_total', 'sum'),
                                الرسوم_المدفوعة=('shipping_cost', 'sum')
                            ).reset_index()
                            ship_stats['الفرق'] = ship_stats['قيمة_الطلبات'] - ship_stats['الرسوم_المدفوعة']
                            st.dataframe(ship_stats, use_container_width=True)
                    
                    with c2:
                        st.markdown("#### 💳 تحليل بوابات الدفع")
                        if 'طريقة الدفع' in df.columns:
                            pay_stats = df.groupby('طريقة الدفع').agg(
                                الطلبات=('رقم الطلب', 'nunique'),
                                قيمة_الطلبات=('product_total', 'sum'),
                                الرسوم_المستقطعة=('gateway_fee', 'sum')
                            ).reset_index()
                            pay_stats['نسبة الرسوم'] = (pay_stats['الرسوم_المستقطعة'] / pay_stats['قيمة_الطلبات'] * 100).fillna(0).map("{:.2f}%".format)
                            st.dataframe(pay_stats, use_container_width=True)

                with tab4:
                    st.subheader("📥 استخراج تقارير Power BI و PDF")
                    st.info("تم تجهيز ملف Excel احترافي يحتوي على كافة الطلبات، الحسابات الدقيقة، الفلاتر، والرسوم البيانية المدمجة الجاهزة للربط المباشر مع Power BI.")
                    
                    excel_data = export_advanced_excel(df, rfm if 'تاريخ الطلب' in df.columns else None, None, None)
                    
                    c1, c2 = st.columns(2)
                    c1.download_button("📊 تحميل ملف قاعدة البيانات (Power BI Excel)", data=excel_data, file_name="Balsam_BI_Database.xlsx", use_container_width=True)
                    
                    st.markdown("""
                        <script>
                        function printReport() { window.print(); }
                        </script>
                        <button onclick="printReport()" style="width:100%; padding:10px; background:#e74c3c; color:white; border:none; border-radius:5px; cursor:pointer;">📄 طباعة التقرير كـ PDF</button>
                    """, unsafe_allow_html=True)
            
            except Exception as e:
                st.error(f"حدث خطأ أثناء معالجة البيانات: {str(e)}")
