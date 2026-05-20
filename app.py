import streamlit as st
import pandas as pd
from utils.database import get_all_users

def show():
    st.markdown(
        """
        <div class="hero">
            <h1>👑 لوحة التحكم الإدارية</h1>
            <p>مرحباً بك في لوحة التحكم</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    
    st.success("✅ النظام يعمل بشكل صحيح!")
    
    # عرض المستخدمين
    st.subheader("📋 المستخدمين المسجلين")
    users_df = get_all_users()
    if not users_df.empty:
        st.dataframe(users_df, use_container_width=True)
    else:
        st.info("لا توجد مستخدمين")
