import streamlit as st
import pandas as pd
from utils.database import fetch_active_items, get_completed_items, get_tab_completed_counts
from utils.helpers import (
    is_cancelled_or_returned_status, is_pending_payment_status,
    get_branch_number, get_branch_location, get_tab_label, numeric_value,
    get_saudi_time, status_pill, case_pill
)
from utils.ui_components import render_metrics

def show():
    pharmacy_name = st.session_state.username
    pharmacist_name = st.session_state.pharmacist_name or ""
    branch_number = get_branch_number(pharmacy_name)
    branch_location = get_branch_location(branch_number)

    st.markdown(f"""
    <div class="hero">
        <h1>🏥 {pharmacy_name}</h1>
        <p>فرع رقم {branch_number} | الموقع: {branch_location} | الصيدلي: {pharmacist_name}</p>
        <p>🕐 آخر تحديث: {get_saudi_time()}</p>
    </div>
    """, unsafe_allow_html=True)

    if st.button("🔄 تحديث الصفحة", use_container_width=True):
        st.rerun()

    df = fetch_active_items(pharmacy_name, include_hidden=False)
    
    if df.empty:
        st.info("📭 لا توجد حالات نشطة لهذا الفرع حاليًا.")
        completed_df = get_completed_items(pharmacy_name)
        if not completed_df.empty:
            st.markdown("---")
            st.markdown('<div class="section-title">✅ الطلبات المكتملة</div>', unsafe_allow_html=True)
            from utils.ui_components import render_completed_table
            render_completed_table(completed_df, is_admin=False)
        return

    allow_actions = True

    col1, col2, col3, col4, col5, col6 = st.columns(6)
    with col1:
        st.metric("📊 إجمالي الحالات", len(df))
    with col2:
        st.metric("➕ إضافات", len(df[df["case_type"] == "addition"]))
    with col3:
        st.metric("➖ إرجاعات", len(df[df["case_type"] == "return"]))
    with col4:
        st.metric("📦 طلبات بدون فاتورة", len(df[df["case_type"] == "orphan_salla"]))
    with col5:
        st.metric("🧾 فواتير بدون طلب", len(df[df["case_type"] == "orphan_abc"]))
    with col6:
        st.metric("✅ تم إنجازها", len(df[df["status"] == "تم"]))

    additions_df = df[df["case_type"] == "addition"].copy()
    returns_df = df[df["case_type"] == "return"].copy()
    orphan_salla_df = df[df["case_type"] == "orphan_salla"].copy()
    orphan_abc_df = df[df["case_type"] == "orphan_abc"].copy()
    
    completed_df = get_completed_items(pharmacy_name)
    tab_completed = get_tab_completed_counts(pharmacy_name)

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        get_tab_label("📈 الإضافات", tab_completed.get("addition", 0), len(additions_df) + tab_completed.get("addition", 0)),
        get_tab_label("📉 الإرجاعات", tab_completed.get("return", 0), len(returns_df) + tab_completed.get("return", 0)),
        get_tab_label("📦 طلبات بدون فاتورة", tab_completed.get("orphan_salla", 0), len(orphan_salla_df) + tab_completed.get("orphan_salla", 0)),
        get_tab_label("🧾 فواتير بدون طلب", tab_completed.get("orphan_abc", 0), len(orphan_abc_df) + tab_completed.get("orphan_abc", 0)),
        get_tab_label("✅ تم الانتهاء", len(completed_df), len(completed_df))
    ])

    with tab1:
        for idx, row in additions_df.iterrows():
            with st.container():
                diff = numeric_value(row['difference'])
                required = "إضافة" if diff > 0 else "إرجاع"
                st.markdown(f"""
                <div style="background:white;border-radius:15px;padding:1rem;margin-bottom:1rem;border-right:4px solid #1f7a8c;">
                    <div><strong>📋 رقم الطلب:</strong> {row['order_number']} | <strong>🏷️ SKU:</strong> {row['sku']}</div>
                    <div><strong>📦 المنتج:</strong> {row['product_name'][:60]}</div>
                    <div><strong>📊 كمية سلة:</strong> {int(row['salla_qty'])} | <strong>📊 كمية ABC:</strong> {int(row['abc_qty'])} | <strong>📊 الفرق:</strong> {'+' if diff > 0 else ''}{diff}</div>
                    <div><strong>🎯 المطلوب:</strong> {required}</div>
                    <div><strong>🧾 رقم الفاتورة/الصيدلي:</strong> {row['invoice_number']}/{row.get('abc_pharmacist_name', 'غير معروف')}</div>
                </div>
                """, unsafe_allow_html=True)
                note = st.text_area("ملحوظة", key=f"note_{idx}", height=60)
                if st.button("✅ تأكيد الإضافة", key=f"add_{idx}"):
                    from utils.database import mark_case_done, save_case_note
                    if note:
                        save_case_note(row['order_number'], row['sku'], pharmacy_name, row['case_type'], note)
                    mark_case_done(row['order_number'], row['sku'], pharmacy_name, row['case_type'], pharmacist_name)
                    st.rerun()
                st.divider()
    
    with tab2:
        for idx, row in returns_df.iterrows():
            with st.container():
                diff = numeric_value(row['difference'])
                required = "إضافة" if diff > 0 else "إرجاع"
                st.markdown(f"""
                <div style="background:white;border-radius:15px;padding:1rem;margin-bottom:1rem;border-right:4px solid #dc3545;">
                    <div><strong>📋 رقم الطلب:</strong> {row['order_number']} | <strong>🏷️ SKU:</strong> {row['sku']}</div>
                    <div><strong>📦 المنتج:</strong> {row['product_name'][:60]}</div>
                    <div><strong>📊 كمية سلة:</strong> {int(row['salla_qty'])} | <strong>📊 كمية ABC:</strong> {int(row['abc_qty'])} | <strong>📊 الفرق:</strong> {'+' if diff > 0 else ''}{diff}</div>
                    <div><strong>🎯 المطلوب:</strong> {required}</div>
                    <div><strong>🧾 رقم الفاتورة/الصيدلي:</strong> {row['invoice_number']}/{row.get('abc_pharmacist_name', 'غير معروف')}</div>
                </div>
                """, unsafe_allow_html=True)
                note = st.text_area("ملحوظة", key=f"note_{idx}", height=60)
                if st.button("🔄 تأكيد الإرجاع", key=f"return_{idx}"):
                    from utils.database import mark_case_done, save_case_note
                    if note:
                        save_case_note(row['order_number'], row['sku'], pharmacy_name, row['case_type'], note)
                    mark_case_done(row['order_number'], row['sku'], pharmacy_name, row['case_type'], pharmacist_name)
                    st.rerun()
                st.divider()
    
    with tab3:
        for idx, row in orphan_salla_df.iterrows():
            with st.container():
                diff = numeric_value(row['difference'])
                st.markdown(f"""
                <div style="background:white;border-radius:15px;padding:1rem;margin-bottom:1rem;border-right:4px solid #ffc107;">
                    <div><strong>📋 رقم الطلب:</strong> {row['order_number']} | <strong>🏷️ SKU:</strong> {row['sku']}</div>
                    <div><strong>📦 المنتج:</strong> {row['product_name'][:60]}</div>
                    <div><strong>📊 كمية سلة:</strong> {int(row['salla_qty'])} | <strong>📊 كمية ABC:</strong> {int(row['abc_qty'])}</div>
                    <div><strong>🧾 رقم الفاتورة/الصيدلي:</strong> {row['invoice_number']}/{row.get('abc_pharmacist_name', 'غير معروف')}</div>
                </div>
                """, unsafe_allow_html=True)
                note = st.text_area("ملحوظة", key=f"note_{idx}", height=60)
                if st.button("✅ تأكيد الإضافة", key=f"orphan_add_{idx}"):
                    from utils.database import mark_case_done, save_case_note
                    if note:
                        save_case_note(row['order_number'], row['sku'], pharmacy_name, row['case_type'], note)
                    mark_case_done(row['order_number'], row['sku'], pharmacy_name, row['case_type'], pharmacist_name)
                    st.rerun()
                st.divider()
    
    with tab4:
        for idx, row in orphan_abc_df.iterrows():
            with st.container():
                diff = numeric_value(row['difference'])
                st.markdown(f"""
                <div style="background:white;border-radius:15px;padding:1rem;margin-bottom:1rem;border-right:4px solid #17a2b8;">
                    <div><strong>📋 رقم الطلب:</strong> {row['order_number']} | <strong>🏷️ SKU:</strong> {row['sku']}</div>
                    <div><strong>📦 المنتج:</strong> {row['product_name'][:60]}</div>
                    <div><strong>📊 كمية ABC:</strong> {int(row['abc_qty'])}</div>
                    <div><strong>🧾 رقم الفاتورة/الصيدلي:</strong> {row['invoice_number']}/{row.get('abc_pharmacist_name', 'غير معروف')}</div>
                </div>
                """, unsafe_allow_html=True)
                note = st.text_area("ملحوظة", key=f"note_{idx}", height=60)
                if st.button("🔄 تأكيد الإرجاع", key=f"orphan_return_{idx}"):
                    from utils.database import mark_case_done, save_case_note
                    if note:
                        save_case_note(row['order_number'], row['sku'], pharmacy_name, row['case_type'], note)
                    mark_case_done(row['order_number'], row['sku'], pharmacy_name, row['case_type'], pharmacist_name)
                    st.rerun()
                st.divider()
    
    with tab5:
        from utils.ui_components import render_completed_table
        render_completed_table(completed_df, is_admin=False)
