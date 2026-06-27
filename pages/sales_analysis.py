import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from utils.financial_engine import calculate_financials, export_advanced_excel
import gc

@st.cache_data(max_entries=1, show_spinner=False)
def load_and_process_data(sales_file, profiles_file, payment_file, tabby_file, tamara_file, emkan_file, jnt_file, aramex_file, beez_file, manual_exp):
    def load_df(f, **kwargs):
        if not f: return None
        fname = f.name.lower()
        try:
            if fname.endswith('.xlsx') or fname.endswith('.xls'): return pd.read_excel(f, **kwargs)
            elif fname.endswith('.csv'): kwargs.pop('sheet_name', None); return pd.read_csv(f, **kwargs)
            return None
        except Exception: return None

    df_sales = load_df(sales_file)
    df, opex = calculate_financials(
        df_sales, load_df(profiles_file), load_df(payment_file), 
        load_df(tabby_file, skiprows=10), load_df(tamara_file, skiprows=26), load_df(emkan_file), 
        load_df(jnt_file, sheet_name="DETAILS"), load_df(aramex_file), load_df(beez_file), 
        manual_exp
    )
    gc.collect() 
    return df, opex

def show():
    st.set_page_config(layout="wide", page_title="نظام ذكاء الأعمال ERP - بلسم العلا")
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700;800&display=swap');
        * { font-family: 'Tajawal', sans-serif; }
        .kpi {background: #fff; padding: 15px; border-radius: 8px; border-top: 4px solid #1f7a8c; box-shadow: 0 2px 5px rgba(0,0,0,0.05); text-align:center; margin-bottom:10px;}
        .kpi-title {color: #555; font-size: 13px; font-weight: bold;}
        .kpi-val {color: #111; font-size: 24px; font-weight: 800;}
        .alert-box {padding: 15px; border-radius: 8px; margin-bottom: 10px; font-weight: bold;}
        .alert-danger {background-color: #ffe5e5; color: #d63031; border-right: 5px solid #d63031;}
        .alert-success {background-color: #e5f9e5; color: #27ae60; border-right: 5px solid #27ae60;}
        .alert-warning {background-color: #fff3cd; color: #f39c12; border-right: 5px solid #f39c12;}
        </style>
    """, unsafe_allow_html=True)

    st.title("🎯 لوحة القيادة الاستراتيجية الشاملة (Strategic Dashboard)")

    with st.expander("📥 1. مركز معالجة البيانات وتوزيع المصروفات", expanded=False):
        c1, c2, c3 = st.columns(3)
        sales_file = c1.file_uploader("📦 مبيعات سلة", type=["xlsx", "csv"])
        profiles_file = c2.file_uploader("💊 البروفايلات (للتكلفة)", type=["xlsx", "csv"])
        payment_file = c3.file_uploader("💳 بوابات الدفع (مدى/فيزا)", type=["xlsx", "csv"])
        
        c4, c5, c6 = st.columns(3)
        tabby_file = c4.file_uploader("🟢 تابي", type=["xlsx", "csv"])
        tamara_file = c5.file_uploader("🟣 تمارا", type=["xlsx", "csv"])
        emkan_file = c6.file_uploader("🔵 إمكان", type=["xlsx", "csv"])
        
        c7, c8, c9 = st.columns(3)
        jnt_file = c7.file_uploader("🚚 J&T (اكسيل)", type=["xlsx", "csv"])
        aramex_file = c8.file_uploader("🚚 Aramex (اكسيل فقط)", type=["xlsx", "csv"])
        beez_file = c9.file_uploader("🚚 Beez (اكسيل)", type=["xlsx", "csv"])

        st.markdown("---")
        if 'manual_exp' not in st.session_state: st.session_state.manual_exp = []
        c_n, c_a, c_b = st.columns([3,2,1])
        with c_n: exp_name = st.text_input("بيان المصروف الثابت (إيجار/تسويق/رواتب)")
        with c_a: exp_amt = st.number_input("المبلغ", min_value=0.0)
        with c_b: 
            st.write(""); 
            if st.button("إضافة مصروف"): st.session_state.manual_exp.append({'desc': exp_name, 'amount': exp_amt}); st.rerun()

        if st.session_state.manual_exp:
            for i, exp in enumerate(st.session_state.manual_exp):
                col_txt, col_del = st.columns([5,1])
                col_txt.info(f"🏷️ {exp['desc']} | 💰 {exp['amount']:,.2f} ر.س")
                if col_del.button("❌", key=f"del_{i}"): st.session_state.manual_exp.pop(i); st.rerun()

    if sales_file:
        with st.spinner("جاري دمج وتحليل البيانات الضخمة بأمان..."):
            
            df, total_opex = load_and_process_data(
                sales_file, profiles_file, payment_file, tabby_file, tamara_file, 
                emkan_file, jnt_file, aramex_file, beez_file, st.session_state.manual_exp
            )

            st.markdown("### 1️⃣ المقاييس الاستراتيجية (Executive KPIs)")
            total_rev = df['product_total'].sum()
            total_cogs = df['total_cost'].sum()
            total_gateway = df['gateway_fee'].sum()
            total_ship = df['shipping_cost'].sum()
            total_expenses = total_opex + total_gateway + total_ship + df['marketing_commission'].sum()
            net_profit = df['net_profit'].sum() - total_opex
            margin = (net_profit / total_rev * 100) if total_rev else 0

            k1, k2, k3, k4, k5 = st.columns(5)
            k1.markdown(f"<div class='kpi'><div class='kpi-title'>إجمالي المبيعات</div><div class='kpi-val'>SAR {total_rev:,.0f}</div></div>", unsafe_allow_html=True)
            k2.markdown(f"<div class='kpi'><div class='kpi-title'>إجمالي المصروفات</div><div class='kpi-val'>SAR {total_expenses:,.0f}</div></div>", unsafe_allow_html=True)
            k3.markdown(f"<div class='kpi'><div class='kpi-title'>صافي الربح</div><div class='kpi-val' style='color:{'#27ae60' if net_profit>0 else '#d63031'}'>SAR {net_profit:,.0f}</div></div>", unsafe_allow_html=True)
            k4.markdown(f"<div class='kpi'><div class='kpi-title'>هامش الربح (Margin)</div><div class='kpi-val'>{margin:.1f}%</div></div>", unsafe_allow_html=True)
            
            outside_sales_ratio = len(df[~df['المدينة'].isin(['العلا', 'تبوك'])]) / len(df) * 100 if 'المدينة' in df.columns and len(df) > 0 else 0
            k5.markdown(f"<div class='kpi'><div class='kpi-title'>طلبات خارج العلا وتبوك</div><div class='kpi-val'>{outside_sales_ratio:.1f}%</div></div>", unsafe_allow_html=True)

            st.markdown("### 🚨 نظام الإنذارات الاستباقية")
            if outside_sales_ratio > 15:
                st.markdown("<div class='alert-box alert-success'>✅ نجاح تسويقي: طلبات التوصيل الخارجي ممتازة (>15%). ينصح النظام بزيادة ميزانية الإعلانات!</div>", unsafe_allow_html=True)
            if 'طريقة الدفع' in df.columns:
                pickup_orders = len(df[df['شركة الشحن'].astype(str).str.contains('استلام', na=False)]) if 'شركة الشحن' in df.columns else 0
                if len(df) > 0 and pickup_orders < (len(df) * 0.10):
                    st.markdown("<div class='alert-box alert-warning'>⚠️ تراجع الفروع: نسبة 'الاستلام من الفرع' منخفضة جداً. هناك مشكلة محتملة في أداء الفروع.</div>", unsafe_allow_html=True)
            
            loss_makers = df.groupby('product_display')['net_profit'].sum()
            loss_count = len(loss_makers[loss_makers < -10])
            if loss_count > 0:
                st.markdown(f"<div class='alert-box alert-danger'>🛑 نزيف أموال: تم اكتشاف {loss_count} منتجاً تباع بخسارة صريحة. راجع التقرير المستخرج لإزالتها!</div>", unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            tab_rfm, tab_prod, tab_opex, tab_branch, tab_export = st.tabs([
                "👥 العملاء (RFM & Geo)", "📦 المنتجات (Profit & Momentum)", 
                "💳 المصروفات (OPEX)", "🏢 أداء الفروع", "📤 تصدير التقارير (BI/PDF)"
            ])

            with tab_rfm:
                st.subheader("تحليل العملاء وخطر الفقدان (مُدمج ببيانات الاتصال)")
                if 'تاريخ الطلب' in df.columns and 'اسم العميل' in df.columns:
                    recent_date = df['تاريخ الطلب'].max()
                    mob_col = 'رقم الجوال' if 'رقم الجوال' in df.columns else None
                    group_cols = ['اسم العميل', mob_col] if mob_col else ['اسم العميل']
                    
                    rfm = df.groupby(group_cols).agg({'تاريخ الطلب': lambda x: (recent_date - x.max()).days, 'رقم الطلب': 'nunique', 'product_total': 'sum'}).reset_index()
                    rfm.columns = ['العميل', 'رقم الجوال', 'أيام الانقطاع', 'الطلبات', 'الإنفاق'] if mob_col else ['العميل', 'أيام الانقطاع', 'الطلبات', 'الإنفاق']
                    
                    c1, c2 = st.columns(2)
                    c1.success("👑 عملاء VIP الوفيون (أعلى 10 نشطون)")
                    c1.dataframe(rfm[(rfm['الإنفاق'] > 1000) & (rfm['أيام الانقطاع'] <= 30)].sort_values('الإنفاق', ascending=False).head(10), use_container_width=True)
                    
                    c2.error("🚨 عملاء VIP في خطر الفقدان (أعلى 10 عرضة للفقدان)")
                    c2.dataframe(rfm[(rfm['الإنفاق'] > 500) & (rfm['أيام الانقطاع'] > 45)].sort_values('الإنفاق', ascending=False).head(10), use_container_width=True)

            with tab_prod:
                st.subheader("تحليل الأداء والربحية للمنتجات والعلامات التجارية")
                prod_stats = df.groupby('product_display').agg({'qty':'sum', 'product_total':'sum', 'net_profit':'sum', 'momentum':'mean'}).reset_index()
                
                c1, c2 = st.columns(2)
                c1.markdown("#### ✅ أعلى 20 منتجاً دراً للربح")
                c1.dataframe(prod_stats.sort_values('net_profit', ascending=False).head(20)[['product_display', 'qty', 'net_profit']], use_container_width=True)
                c2.markdown("#### 🩸 أعلى 20 منتجاً نازفاً للربح")
                c2.dataframe(prod_stats[prod_stats['net_profit'] < 0].sort_values('net_profit').head(20)[['product_display', 'net_profit', 'momentum']], use_container_width=True)

            with tab_opex:
                st.subheader("تحليل دقيق لشركات الشحن وبوابات الدفع")
                c1, c2 = st.columns(2)
                with c1:
                    if 'شركة الشحن' in df.columns:
                        ship_stats = df.groupby('شركة الشحن').agg(الطلبات=('رقم الطلب', 'nunique'), الإيرادات=('product_total', 'sum'), الرسوم_المدفوعة=('shipping_cost', 'sum')).reset_index()
                        ship_stats['نسبة الرسوم'] = (ship_stats['الرسوم_المدفوعة'] / ship_stats['الإيرادات'] * 100).fillna(0).map("{:.1f}%".format)
                        st.dataframe(ship_stats.sort_values('الرسوم_المدفوعة', ascending=False), use_container_width=True)
                with c2:
                    if 'طريقة الدفع' in df.columns:
                        pay_stats = df.groupby('طريقة الدفع').agg(الطلبات=('رقم الطلب', 'nunique'), الإيرادات=('product_total', 'sum'), الرسوم=('gateway_fee', 'sum')).reset_index()
                        pay_stats['تكلفة البوابة %'] = (pay_stats['الرسوم'] / pay_stats['الإيرادات'] * 100).fillna(0).map("{:.2f}%".format)
                        st.dataframe(pay_stats.sort_values('الرسوم', ascending=False), use_container_width=True)

            # --- 5. أداء الفروع (مع المستهدفات القابلة للتعديل) ---
            with tab_branch:
                st.subheader("🏢 أداء الفروع (الاستلام من الفرع)")
                st.info("💡 أدخل المستهدف الشهري لكل فرع في الجدول أدناه، وسيقوم النظام بحساب نسبة الإنجاز فوراً!")
                
                if 'شركة الشحن' in df.columns:
                    branch_orders = df[df['شركة الشحن'].astype(str).str.contains('فرع|استلام|branch', case=False, na=False)]
                    
                    if not branch_orders.empty:
                        branch_df = branch_orders.groupby('شركة الشحن').agg(إجمالي_المبيعات=('product_total', 'sum'), الطلبات=('رقم الطلب', 'nunique')).reset_index()
                        branch_df.rename(columns={'شركة الشحن': 'اسم الفرع'}, inplace=True)
                        
                        # 💡 نظام المستهدفات الديناميكي
                        if 'branch_targets' not in st.session_state:
                            default_targets = pd.DataFrame({
                                'اسم الفرع': branch_df['اسم الفرع'].tolist(),
                                'المستهدف_الشهري': [100000] * len(branch_df)
                            })
                            st.session_state.branch_targets = default_targets
                        
                        # مزامنة الفروع الجديدة إن وجدت
                        current_branches = branch_df['اسم الفرع'].tolist()
                        saved_branches = st.session_state.branch_targets['اسم الفرع'].tolist()
                        missing = [b for b in current_branches if b not in saved_branches]
                        if missing:
                            new_df = pd.DataFrame({'اسم الفرع': missing, 'المستهدف_الشهري': [100000]*len(missing)})
                            st.session_state.branch_targets = pd.concat([st.session_state.branch_targets, new_df], ignore_index=True)

                        # 📝 عرض محرر البيانات للمستخدم لتعديل المستهدفات
                        st.markdown("#### 🎯 قم بتعديل المستهدف هنا:")
                        edited_targets = st.data_editor(
                            st.session_state.branch_targets, 
                            hide_index=True, 
                            use_container_width=True,
                            disabled=["اسم الفرع"] # قفل اسم الفرع للسماح بتعديل الرقم فقط
                        )
                        st.session_state.branch_targets = edited_targets
                        
                        # دمج المستهدفات المُعدلة مع المبيعات الفعلية
                        final_branch_df = branch_df.merge(edited_targets, on='اسم الفرع', how='left')
                        final_branch_df['المستهدف_الشهري'] = pd.to_numeric(final_branch_df['المستهدف_الشهري'], errors='coerce').fillna(1)
                        final_branch_df['نسبة الإنجاز'] = (final_branch_df['إجمالي_المبيعات'] / final_branch_df['المستهدف_الشهري'] * 100).map("{:.1f}%".format)
                        final_branch_df['الحالة'] = np.where(final_branch_df['إجمالي_المبيعات'] >= final_branch_df['المستهدف_الشهري'], "✅ محقق", "🔴 متراجع")
                        
                        st.markdown("#### 📊 النتيجة الفورية:")
                        st.dataframe(final_branch_df.sort_values('إجمالي_المبيعات', ascending=False), use_container_width=True)
                    else:
                        st.warning("لم يتم العثور على طلبات مسجلة كـ 'استلام من الفرع'.")

            with tab_export:
                st.subheader("📥 استخراج تقارير الإدارة العليا (بكامل البيانات)")
                excel_data = export_advanced_excel(df, rfm if 'تاريخ الطلب' in df.columns and 'اسم العميل' in df.columns else None)
                st.download_button("📊 تحميل التقرير الكامل (Excel) بجميع الأعمدة", data=excel_data, file_name="Balsam_BI_Full_Report.xlsx", use_container_width=True)
