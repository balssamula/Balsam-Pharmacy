import streamlit as st
import pandas as pd
from utils.database import get_all_users, add_user, delete_user, update_user_permissions

def show():
    st.markdown(
        """
        <div class="hero">
            <h1>👥 إدارة المستخدمين والصلاحيات</h1>
            <p>إضافة وتعديل وحذف المستخدمين وتحديد صلاحياتهم</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    
    # إضافة مستخدم جديد
    with st.expander("➕ إضافة مستخدم جديد", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            new_username = st.text_input("اسم المستخدم")
            new_password = st.text_input("كلمة المرور", type="password")
        with col2:
            new_role = st.selectbox("نوع المستخدم", ["pharmacy", "admin"])
            new_pharmacist_name = st.text_input("اسم الصيدلي (للفروع فقط)")
        
        if st.button("➕ إضافة مستخدم", use_container_width=True):
            if new_username and new_password:
                if add_user(new_username, new_password, new_role, new_pharmacist_name):
                    st.success(f"✅ تم إضافة المستخدم {new_username}")
                    st.rerun()
                else:
                    st.error("❌ اسم المستخدم موجود مسبقاً")
            else:
                st.error("❌ الرجاء إدخال اسم المستخدم وكلمة المرور")
    
    # عرض المستخدمين
    st.markdown("### 📋 قائمة المستخدمين")
    users_df = get_all_users()
    
    if not users_df.empty:
        for idx, row in users_df.iterrows():
            with st.container():
                col1, col2, col3, col4 = st.columns([1.5, 2, 2, 1])
                
                with col1:
                    st.markdown(f"**👤 {row['username']}**")
                    st.caption(f"نوع: {row['role']}")
                
                with col2:
                    st.markdown(f"**اسم الصيدلي:** {row['pharmacist_name'] or '-'}")
                    st.caption(f"آخر دخول: {row['last_login'][:16] if row['last_login'] else 'لم يدخل'}")
                
                with col3:
                    if row['username'] != "admin":
                        new_pharm_name = st.text_input("اسم الصيدلي", value=row['pharmacist_name'] or "", key=f"name_{idx}")
                        perms = {
                            "can_view_dashboard": st.checkbox("لوحة التحكم الرئيسية", value=bool(row['can_view_dashboard']), key=f"dash_{idx}"),
                            "can_view_balances": st.checkbox("تحديث الأرصدة", value=bool(row['can_view_balances']), key=f"bal_{idx}"),
                            "can_view_monitoring": st.checkbox("مراقبة التعديلات", value=bool(row['can_view_monitoring']), key=f"mon_{idx}"),
                            "can_manage_users": st.checkbox("إدارة المستخدمين", value=bool(row['can_manage_users']), key=f"users_{idx}"),
                            "pharmacist_name": new_pharm_name
                        }
                        if st.button("💾 حفظ التعديلات", key=f"save_{idx}"):
                            update_user_permissions(row['username'], perms)
                            st.success(f"✅ تم تحديث صلاحيات {row['username']}")
                            st.rerun()
                    else:
                        st.info("المستخدم admin لديه جميع الصلاحيات")
                
                with col4:
                    if row['username'] != "admin":
                        if st.button("🗑️ حذف", key=f"delete_{idx}"):
                            if delete_user(row['username']):
                                st.success(f"✅ تم حذف المستخدم {row['username']}")
                                st.rerun()
                
                st.divider()
    else:
        st.info("لا توجد مستخدمين")
