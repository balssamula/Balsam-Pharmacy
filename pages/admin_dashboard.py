import streamlit as st
import pandas as pd
from utils.database import get_all_users, get_latest_upload_summary
from utils.ui_components import render_metrics

def show():
    st.markdown(
        """
        <div class="hero">
            <h1>👑 لوحة التحكم الإدارية</h1>
            <p>مرحباً بك في لوحة التحكم الرئيسية</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    
    latest = get_latest_upload_summary()
    if latest:
        st.info(f"📂 آخر ملف مرفوع: {latest[1]} - بواسطة: {latest[2]} - التاريخ: {latest[3]}")
    
    # عرض المستخدمين
    st.subheader("📋 المستخدمين المسجلين")
    users_df = get_all_users()
    if not users_df.empty:
        st.dataframe(users_df[['username', 'role', 'pharmacist_name', 'last_login']], use_container_width=True)
    else:
        st.info("لا توجد مستخدمين")
    
    st.success("✅ النظام يعمل بشكل صحيح!")
