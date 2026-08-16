import streamlit as st
import pandas as pd
import sqlite3
from io import BytesIO
from datetime import datetime
from utils.database import DB_PATH, get_action_logs

def show():
    st.markdown(
        """
        <div class="hero">
            <h1>👥 مراقبة الأداء وسجل العمليات</h1>
            <p>متابعة إنجازات الصيادلة والسجل الشامل لكافة التحركات والتعديلات بالنظام</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    
    # تبويبات لتقسيم واجهة المراقبة
    tab1, tab2, tab3 = st.tabs(["📋 سجل العمليات الشامل (Logs)", "👤 الصيادلة المسجلون", "✅ سجل التسويات المكتملة"])
    
    with tab1:
        st.markdown('### 🕒 السجل الزمني لكافة العمليات بالنظام')
        st.info("يعرض هذا الجدول كل خطوة تمت في النظام بدءاً من رفع الملف، مروراً بالملاحظات والنقل والاعتمادات بالوقت والتاريخ ومن قام بها.")
        
        logs_df = get_action_logs(limit=1000)
        
        if not logs_df.empty:
            # إضافة فلاتر بحث سريعة
            col1, col2 = st.columns(2)
            search_order = col1.text_input("🔍 بحث برقم الطلب (في السجل)")
            search_sku = col2.text_input("🏷️ بحث برقم المنتج (SKU)")
            
            filtered_logs = logs_df.copy()
            if search_order:
                filtered_logs = filtered_logs[filtered_logs["order_number"].astype(str).str.contains(search_order, na=False)]
            if search_sku:
                filtered_logs = filtered_logs[filtered_logs["sku"].astype(str).str.contains(search_sku, na=False)]
                
            filtered_logs = filtered_logs.rename(columns={
                "action_date": "التاريخ والوقت",
                "performed_by": "بواسطة",
                "role": "الصلاحية",
                "pharmacy_name": "الفرع / الجهة",
                "order_number": "رقم الطلب",
                "sku": "SKU",
                "action_type": "نوع الإجراء",
                "action_details": "التفاصيل"
            })
            
            st.dataframe(filtered_logs, use_container_width=True, height=500)
            
            output = BytesIO()
            filtered_logs.to_excel(output, index=False)
            output.seek(0)
            st.download_button(
                "📥 تصدير السجل الشامل (Excel)",
                data=output,
                file_name=f"Full_Action_Logs_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        else:
            st.warning("لا توجد عمليات مسجلة في النظام حتى الآن.")

    with tab2:
        st.markdown('### 👤 آخر دخول للصيادلة')
        conn = sqlite3.connect(DB_PATH)
        try:
            pharmacists_df = pd.read_sql_query("""
                SELECT username, pharmacist_name, last_login, last_ip
                FROM users
                WHERE role = 'pharmacy' AND pharmacist_name != ''
                ORDER BY last_login DESC
            """, conn)
            
            if not pharmacists_df.empty:
                st.dataframe(pharmacists_df.rename(columns={
                    "username": "اسم المستخدم",
                    "pharmacist_name": "اسم الصيدلي",
                    "last_login": "آخر دخول",
                    "last_ip": "IP الجهاز"
                }), use_container_width=True)
            else:
                st.info("لا يوجد صيادلة مسجلون بعد")
        finally:
            conn.close()

    with tab3:
        st.markdown('### ✅ الحالات التي تم إنجازها (التسويات)')
        conn = sqlite3.connect(DB_PATH)
        try:
            adjustments_df = pd.read_sql_query("""
                SELECT order_number, sku, product_name, pharmacy_name, case_type, 
                       status, performed_by, performed_at, pharmacist_note
                FROM reconciliation_items
                WHERE performed_by != '' AND status = 'تم'
                ORDER BY performed_at DESC
                LIMIT 500
            """, conn)
            
            if not adjustments_df.empty:
                adjustments_df = adjustments_df.rename(columns={
                    "order_number": "رقم الطلب",
                    "sku": "SKU",
                    "product_name": "المنتج",
                    "pharmacy_name": "الصيدلية",
                    "case_type": "نوع الإجراء",
                    "status": "الحالة",
                    "performed_by": "تم بواسطة",
                    "performed_at": "تاريخ التنفيذ",
                    "pharmacist_note": "ملحوظة"
                })
                st.dataframe(adjustments_df, use_container_width=True)
                
                output = BytesIO()
                adjustments_df.to_excel(output, index=False)
                output.seek(0)
                st.download_button(
                    "📥 تصدير سجل الإنجازات والتسويات إلى Excel",
                    data=output,
                    file_name=f"pharmacy_completed_adjustments_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            else:
                st.info("لا توجد تسويات مكتملة حتى الآن.")
        finally:
            conn.close()
