import streamlit as st
from utils.database import fetch_active_items, get_completed_items
from utils.helpers import get_branch_number, is_cancelled_or_returned_status, is_pending_payment_status
from utils.ui_components import render_metrics, render_case_cards, render_completed_table, get_tab_label

def show():
    pharmacy_name = st.session_state.username
    pharmacist_name = st.session_state.pharmacist_name or ""
    branch_number = get_branch_number(pharmacy_name)

    st.markdown(
        f"""
        <div class="hero">
            <h1>{pharmacy_name}</h1>
            <p>فرع رقم {branch_number} | الصيدلي: {pharmacist_name}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if st.button("🔄 تحديث الصفحة", use_container_width=True):
        st.rerun()

    df = fetch_active_items(pharmacy_name, include_hidden=False)
    
    if df.empty:
        st.info("لا توجد حالات نشطة لهذا الفرع حاليًا.")
        completed_df = get_completed_items(pharmacy_name)
        if not completed_df.empty:
            st.markdown("---")
            st.markdown('<div class="section-title">✅ الطلبات المكتملة</div>', unsafe_allow_html=True)
            render_completed_table(completed_df, is_admin=False)
        return

    render_metrics(df)

    active_non_cancelled = ~df["order_status"].apply(is_cancelled_or_returned_status)
    active_non_payment = ~df["order_status"].apply(is_pending_payment_status)
    active_operational = active_non_cancelled & active_non_payment
    
    additions_df = df[(df["case_type"] == "addition") & active_operational].copy()
    returns_df = df[(df["case_type"] == "return") & active_operational].copy()
    
    total_items = len(additions_df) + len(returns_df)
    
    tab1, tab2 = st.tabs([
        get_tab_label("الإضافات", len(additions_df), total_items),
        get_tab_label("الإرجاعات", len(returns_df), total_items),
    ])

    with tab1:
        render_case_cards(additions_df, True, pharmacist_name, pharmacy_name, is_admin=False)
    with tab2:
        render_case_cards(returns_df, True, pharmacist_name, pharmacy_name, is_admin=False)
