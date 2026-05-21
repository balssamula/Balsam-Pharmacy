import streamlit as st
import pandas as pd
from utils.database import fetch_active_items, get_completed_items, get_tab_completed_counts
from utils.helpers import (
    is_cancelled_or_returned_status, is_pending_payment_status, 
    get_branch_number, get_branch_location, get_tab_label, numeric_value,
    get_saudi_time
)
from utils.ui_components import render_metrics, render_completed_table

def render_case_cards_pharmacy(df: pd.DataFrame, allow_actions: bool, pharmacist_name: str, pharmacy_name: str):
    if df.empty:
        st.success("لا توجد حالات في هذا القسم.")
        return

    for idx, row in df.iterrows():
        diff_value = numeric_value(row['difference'])
        required_action = "إضافة" if diff_value > 0 else "إرجاع" if diff_value < 0 else "مطابق"
        order_status = row.get('order_status', 'غير متوفرة')
        
        # تحديد لون الخلفية حسب حالة الطلب
        status_color = "#fff3cd" if "بانتظار الدفع" in order_status else "#f8f9fa"
        
        with st.container():
            st.markdown(f"""
            <div style="background:{status_color};border-radius:16px;padding:1rem;margin-bottom:1rem;border-right:4px solid #1f7a8c;box-shadow:0 2px 8px rgba(0,0,0,0.05);">
                <div style="display:flex;justify-content:space-between;margin-bottom:0.5rem;">
                    <span style="background:#dff1ff;color:#0f5488;padding:0.2rem 0.8rem;border-radius:20px;font-size:0.8rem;">{row['case_label']}</span>
                    <span style="color:#6c757d;font-size:0.8rem;">📅 {row['order_date'][:16] if row['order_date'] else ''}</span>
                </div>
                <div style="display:flex;flex-wrap:wrap;gap:1rem;">
                    <div style="flex:2;">
                        <strong>📋 رقم الطلب:</strong> {row['order_number']}<br>
                        <strong>🏷️ SKU:</strong> {row['sku']}<br>
                        <strong>📦 المنتج:</strong> {row['product_name'][:60]}
                    </div>
                    <div style="flex:1;">
                        <strong>📊 الكميات:</strong><br>
                        🛒 سلة: {int(row['salla_qty']) if pd.notna(row['salla_qty']) else 0}<br>
                        📄 ABC: {int(row['abc_qty']) if pd.notna(row['abc_qty']) else 0}<br>
                        <strong>📊 الفرق:</strong> <span style="color:{'#28a745' if diff_value > 0 else '#dc3545' if diff_value < 0 else '#6c757d'};font-weight:bold;">{'+' if diff_value > 0 else ''}{diff_value}</span>
                    </div>
                    <div style="flex:1.5;">
                        <strong>🧾 الفاتورة/الصيدلي:</strong><br>
                        {row['invoice_number']}/{row.get('abc_pharmacist_name', 'غير معروف')}<br>
                        <strong>📌 حالة الطلب:</strong> <span style="color:#d9534f;">{order_status}</span><br>
                        <strong>🎯 المطلوب:</strong> <span style="color:{'#28a745' if diff_value > 0 else '#dc3545' if diff_value < 0 else '#6c757d'};">{required_action}</span>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            note_key = f"note_{idx}"
            note_value = st.text_area("📝 ملحوظة الصيدلي", value=row.get("pharmacist_note", "") or "", key=note_key, height=60)
            
            btn_col1, btn_col2 = st.columns([1, 4])
            with btn_col1:
                if st.button("💾 حفظ", key=f"save_{note_key}", use_container_width=True):
                    from utils.database import save_case_note
                    save_case_note(row['order_number'], row['sku'], pharmacy_name, row['case_type'], note_value)
                    st.rerun()
            
            if allow_actions and row["status"] != "تم" and row["case_type"] in {"addition", "return", "orphan_salla", "orphan_abc"}:
                button_label = "✅ تأكيد الإضافة" if diff_value > 0 else "🔄 تأكيد الإرجاع"
                with btn_col2:
                    if st.button(button_label, key=f"done_{note_key}", use_container_width=True):
                        from utils.database import save_case_note, mark_case_done
                        save_case_note(row['order_number'], row['sku'], pharmacy_name, row['case_type'], note_value)
                        mark_case_done(row['order_number'], row['sku'], pharmacy_name, row['case_type'], pharmacist_name)
                        st.rerun()
            st.markdown("---")

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
            render_completed_table(completed_df, is_admin=False)
        return

    # التحقق من القفل
    is_locked = False
    if 'is_locked' in df.columns and not df.empty:
        is_locked = df['is_locked'].iloc[0] == 1
    allow_actions = not is_locked

    # إحصائيات سريعة
    total = len(df)
    additions = len(df[df["case_type"] == "addition"])
    returns = len(df[df["case_type"] == "return"])
    orphan_salla = len(df[df["case_type"] == "orphan_salla"])
    orphan_abc = len(df[df["case_type"] == "orphan_abc"])
    post_cutoff = len(df[df["case_type"] == "post_cutoff_abc"])
    completed = len(df[df["status"] == "تم"])
    
    col1, col2, col3, col4, col5, col6, col7 = st.columns(7)
    with col1:
        st.metric("📊 إجمالي الحالات", total)
    with col2:
        st.metric("➕ إضافات", additions)
    with col3:
        st.metric("➖ إرجاعات", returns)
    with col4:
        st.metric("📦 طلبات بدون فاتورة", orphan_salla)
    with col5:
        st.metric("🧾 فواتير بدون طلب", orphan_abc)
    with col6:
        st.metric("⏰ فواتير بعد آخر طلب", post_cutoff)
    with col7:
        st.metric("✅ تم إنجازها", completed)

    # فصل البيانات حسب النوع
    additions_df = df[df["case_type"] == "addition"].copy()
    returns_df = df[df["case_type"] == "return"].copy()
    orphan_salla_df = df[df["case_type"] == "orphan_salla"].copy()
    orphan_abc_df = df[df["case_type"] == "orphan_abc"].copy()
    post_cutoff_df = df[df["case_type"] == "post_cutoff_abc"].copy()
    payment_df = df[df["order_status"].apply(is_pending_payment_status)].copy()
    cancelled_df = df[df["order_status"].apply(is_cancelled_or_returned_status)].copy()
    
    completed_df = get_completed_items(pharmacy_name)
    tab_completed = get_tab_completed_counts(pharmacy_name)
    
    # أعداد التبويبات
    tab1_completed = tab_completed.get("addition", 0)
    tab2_completed = tab_completed.get("return", 0)
    tab3_completed = tab_completed.get("orphan_salla", 0)
    tab4_completed = tab_completed.get("orphan_abc", 0)
    tab5_completed = tab_completed.get("post_cutoff_abc", 0)
    
    # عرض التبويبات
    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
        get_tab_label("📈 الإضافات", tab1_completed, len(additions_df) + tab1_completed),
        get_tab_label("📉 الإرجاعات", tab2_completed, len(returns_df) + tab2_completed),
        get_tab_label("📦 طلبات بدون فاتورة", tab3_completed, len(orphan_salla_df) + tab3_completed),
        get_tab_label("🧾 فواتير بدون طلب", tab4_completed, len(orphan_abc_df) + tab4_completed),
        get_tab_label("⏰ فواتير بعد آخر طلب", tab5_completed, len(post_cutoff_df) + tab5_completed),
        get_tab_label("💰 بانتظار الدفع", 0, len(payment_df)),
        get_tab_label("⚠️ ملغي/مسترجع", 0, len(cancelled_df)),
        get_tab_label("✅ تم الانتهاء", len(completed_df), len(completed_df))
    ])

    with tab1:
        render_case_cards_pharmacy(additions_df, allow_actions, pharmacist_name, pharmacy_name)
    with tab2:
        render_case_cards_pharmacy(returns_df, allow_actions, pharmacist_name, pharmacy_name)
    with tab3:
        render_case_cards_pharmacy(orphan_salla_df, allow_actions, pharmacist_name, pharmacy_name)
    with tab4:
        render_case_cards_pharmacy(orphan_abc_df, allow_actions, pharmacist_name, pharmacy_name)
    with tab5:
        render_case_cards_pharmacy(post_cutoff_df, False, pharmacist_name, pharmacy_name)
    with tab6:
        render_case_cards_pharmacy(payment_df, False, pharmacist_name, pharmacy_name)
    with tab7:
        render_case_cards_pharmacy(cancelled_df, False, pharmacist_name, pharmacy_name)
    with tab8:
        render_completed_table(completed_df, is_admin=False)
