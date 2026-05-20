import streamlit as st
import pandas as pd
from utils.database import get_all_users, get_latest_upload_summary, fetch_active_items
from utils.ui_components import render_metrics
from utils.excel_processor import process_excel

def show():
    st.markdown(
        """
        <div class="hero">
            <h1>👑 لوحة التحكم الإدارية</h1>
            <p>مرحباً بك في لوحة التحكم الرئيسية - إدارة الطلبات والفواتير</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    
    # رفع ملف الطلبات والفواتير
    with st.expander("📂 رفع ملف الطلبات والفواتير", expanded=True):
        st.markdown("#### قم برفع ملف Excel يحتوي على شيتين: 'سلة' و 'abc'")
        uploaded_file = st.file_uploader("اختر ملف Excel", type=["xlsx"], key="reconciliation_upload")
        
        if uploaded_file:
            if st.button("🔄 معالجة الملف وترحيل الحالات", use_container_width=True):
                with st.spinner("جاري قراءة الملف وتصنيف الحالات..."):
                    results, upload_batch_id = process_excel(uploaded_file, st.session_state.username)
                if results is not None:
                    st.success(f"✅ تمت المعالجة بنجاح! عدد الحالات: {len(results)}")
                    st.balloons()
                    st.session_state.processed_data = results
                    st.rerun()
    
    # عرض آخر ملف تم رفعه
    latest = get_latest_upload_summary()
    if latest:
        st.markdown(f"""
        <div class="note-card">
            <strong>📋 آخر جلسة نشطة:</strong><br>
            الملف: {latest[1]}<br>
            بواسطة: {latest[2]}<br>
            التاريخ: {latest[3]}<br>
            الحالات: {latest[4]} (إضافات: {latest[5]}, إرجاعات: {latest[6]})
        </div>
        """, unsafe_allow_html=True)
    
    # عرض البيانات إذا كانت موجودة
    if st.session_state.get('processed_data') is not None and len(st.session_state.processed_data) > 0:
        df = st.session_state.processed_data
        
        st.markdown("### 📊 البيانات المعالجة")
        
        # إحصائيات سريعة
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("إجمالي الحالات", len(df))
        with col2:
            st.metric("إضافات", len(df[df['case_type'] == 'addition']))
        with col3:
            st.metric("إرجاعات", len(df[df['case_type'] == 'return']))
        
        # عرض الجدول
        st.dataframe(df[['order_number', 'sku', 'product_name', 'pharmacy_name', 
                         'case_type', 'salla_qty', 'abc_qty', 'difference']], 
                     use_container_width=True)
        
        # زر تنزيل النتائج
        output = pd.ExcelWriter('temp.xlsx', engine='openpyxl')
        df.to_excel(output, index=False)
        output.close()
        with open('temp.xlsx', 'rb') as f:
            st.download_button(
                "📥 تحميل النتائج كـ Excel",
                data=f,
                file_name="reconciliation_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    
    # عرض المستخدمين
    st.markdown("### 📋 المستخدمين المسجلين")
    users_df = get_all_users()
    if not users_df.empty:
        st.dataframe(users_df[['username', 'role', 'pharmacist_name', 'last_login']], use_container_width=True)
    else:
        st.info("لا توجد مستخدمين")
