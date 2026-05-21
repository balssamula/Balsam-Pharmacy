import streamlit as st
import pandas as pd
from utils.database import fetch_active_items, get_completed_items, get_tab_completed_counts
from utils.helpers import (
    is_cancelled_or_returned_status, is_pending_payment_status, 
    get_branch_number, get_branch_location, get_tab_label, numeric_value,
    get_saudi_time, status_pill, case_pill
)
from utils.ui_components import render_metrics

def render_case_cards_pharmacy(df: pd.DataFrame, allow_actions: bool, pharmacist_name: str, pharmacy_name: str):
    """عرض البطاقات في صفحة الصيدليات مع حساب الفرق بشكل صحيح"""
    if df.empty:
        st.success("لا توجد حالات في هذا القسم.")
        return

    for idx, row in df.iterrows():
        diff_value = numeric_value(row['difference'])
        # تحديد المطلوب (إضافة أو إرجاع)
        required_action = "إضافة" if diff_value > 0 else "إرجاع"
        
        with st.container():
            st.markdown(f"""
            <div class="action-card">
                <div style="display:flex;justify-content:space-between;gap:1rem;align-items:center;flex-wrap:wrap;">
                    <div>
                        {case_pill(row['case_type'])}&nbsp; {status_pill(row['status'])}
                    </div>
                    <div style="font-weight:700;color:#48606a;">🏥 {row['pharmacy_name'] or 'غير محدد'}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown(f"**📋 رقم الطلب:** {row['order_number']}")
                st.markdown(f"**🏷️ SKU:** {row['sku']}")
                st.markdown(f"**📦 المنتج:** {row['product_name'][:50]}...")
            
            with col2:
                st.markdown(f"**📊 كمية سلة:** {int(row['salla_qty']) if pd.notna(row['salla_qty']) else 0}")
                st.markdown(f"**📊 كمية ABC:** {int(row['abc_qty']) if pd.notna(row['abc_qty']) else 0}")
                st.markdown(f"**📊 الفرق:** {'+' if diff_value > 0 else ''}{diff_value}")
                st.markdown(f"**🎯 المطلوب:** {required_action}")
            
            with st.expander("📋 تفاصيل إضافية"):
                detail_cols = st.columns(3)
                with detail_cols[0]:
                    st.markdown(f"**حالة الطلب:** {row['order_status'] or 'غير متوفرة'}")
                    st.markdown(f"**المدينة:** {row['city'] or 'غير متوفرة'}")
                with detail_cols[1]:
                    st.markdown(f"**تاريخ الطلب:** {row['order_date'][:16] if row['order_date'] else 'غير متوفر'}")
                    st.markdown(f"**تاريخ الفاتورة:** {row['invoice_date'][:16] if row['invoice_date'] else 'غير متوفر'}")
                with detail_cols[2]:
                    pharmacist_display = row.get('abc_pharmacist_name', '') or ''
                    if not pharmacist_display:
                        pharmacist_display = 'غير معروف'
                    st.markdown(f"**الصيدلي:** {pharmacist_display}")
                    st.markdown(f"**نوع البروفايل:** {row.get('profile_type', '') or 'غير متوفر'}")
                st.markdown(f"**التفصيل:** {row.get('case_reason', '')}")
            
            note_key = f"note_{idx}"
            note_value = st.text_area(
                "📝 ملحوظة الصيدلي",
                value=row.get("pharmacist_note", "") or "",
                key=note_key,
                height=60,
            )
            
            btn_col1, btn_col2 = st.columns([1, 4])
            with btn_col1:
                if st.button("💾 حفظ", key=f"save_{note_key}", use_container_width=True):
                    from utils.database import save_case_note
                    save_case_note(
                        order_number=row["order_number"],
                        sku=row["sku"],
                        pharmacy_name=pharmacy_name,
                        case_type=row["case_type"],
                        note=note_value,
                    )
                    st.rerun()
            
            if allow_actions and row["status"] != "تم" and row["case_type"] in {"addition", "return", "orphan_salla", "orphan_abc"}:
                button_label = "✅ تأكيد الإضافة" if diff_value > 0 else "🔄 تأكيد الإرجاع"
                with btn_col2:
                    if st.button(button_label, key=f"done_{note_key}", use_container_width=True):
                        from utils.database import save_case_note, mark_case_done
                        save_case_note(
                            order_number=row["order_number"],
                            sku=row["sku"],
                            pharmacy_name=pharmacy_name,
                            case_type=row["case_type"],
                            note=note_value,
                        )
                        mark_case_done(
                            order_number=row["order_number"],
                            sku=row["sku"],
                            pharmacy_name=pharmacy_name,
                            case_type=row["case_type"],
                            performed_by=pharmacist_name,
                        )
                        st.rerun()
            
            st.markdown("---")

def show():
    pharmacy_name = st.session_state.username
    pharmacist_name = st.session_state.pharmacist_name or ""
    branch_number = get_branch_number(pharmacy_name)
    branch_location = get_branch_location(branch_number)

    st.markdown(
        f"""
        <div class="hero">
            <h1>🏥 {pharmacy_name}</h1>
            <p>فرع رقم {branch_number} | الموقع: {branch_location} | الصيدلي: {pharmacist_name}</p>
            <p>🕐 آخر تحديث: {get_saudi_time()}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

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

    is_locked = df['is_locked'].iloc[0] == 1 if not df.empty else False
    
    if is_locked:
        st.warning("🔒 هذه الجلسة مقفلة ولا يمكن إجراء تعديلات عليها.")
        allow_actions = False
    else:
        allow_actions = True

    # إحصائيات سريعة
    total = len(df)
    additions = len(df[df["case_type"] == "addition"])
    returns = len(df[df["case_type"] == "return"])
    orphan_salla = len(df[df["case_type"] == "orphan_salla"])
    orphan_abc = len(df[df["case_type"] == "orphan_abc"])
    completed = len(df[df["status"] == "تم"])
    
    col1, col2, col3, col4, col5, col6 = st.columns(6)
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
        st.metric("✅ تم إنجازها", completed)

    active_non_cancelled = ~df["order_status"].apply(is_cancelled_or_returned_status)
    active_non_payment = ~df["order_status"].apply(is_pending_payment_status)
    active_operational = active_non_cancelled & active_non_payment
    
    additions_df = df[(df["case_type"] == "addition") & active_operational].copy()
    returns_df = df[(df["case_type"] == "return") & active_operational].copy()
    orphan_salla_df = df[(df["case_type"] == "orphan_salla") & active_operational].copy()
    orphan_abc_df = df[(df["case_type"] == "orphan_abc") & active_operational].copy()
    post_cutoff_df = df[df["case_type"] == "post_cutoff_abc"].copy()
    cancelled_df = df[df["order_status"].apply(is_cancelled_or_returned_status)].copy()
    payment_pending_df = df[df["order_status"].apply(is_pending_payment_status)].copy()
    
    completed_df = get_completed_items(pharmacy_name)
    
    # الحصول على أعداد المنجزات داخل كل تبويب
    tab_completed = get_tab_completed_counts(pharmacy_name)
    
    # إعداد التبويبات مع الأعداد الصحيحة (المنجز/الإجمالي داخل نفس التبويب)
    tab1_total = len(additions_df) + len(df[df["case_type"] == "addition"])
    tab1_completed = tab_completed.get("addition", 0)
    
    tab2_total = len(returns_df) + len(df[df["case_type"] == "return"])
    tab2_completed = tab_completed.get("return", 0)
    
    tab3_total = len(orphan_salla_df) + len(df[df["case_type"] == "orphan_salla"])
    tab3_completed = tab_completed.get("orphan_salla", 0)
    
    tab4_total = len(orphan_abc_df) + len(df[df["case_type"] == "orphan_abc"])
    tab4_completed = tab_completed.get("orphan_abc", 0)
    
    tab8_total = len(completed_df)
    tab8_completed = tab8_total
    
    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
        get_tab_label("📈 الإضافات", tab1_completed, tab1_total),
        get_tab_label("📉 الإرجاعات", tab2_completed, tab2_total),
        get_tab_label("📦 طلبات بدون فاتورة", tab3_completed, tab3_total),
        get_tab_label("🧾 فواتير بدون طلب", tab4_completed, tab4_total),
        get_tab_label("⏰ فواتير بعد آخر طلب", 0, len(post_cutoff_df)),
        get_tab_label("💰 بانتظار الدفع", 0, len(payment_pending_df)),
        get_tab_label("⚠️ ملغي/مسترجع", 0, len(cancelled_df)),
        get_tab_label("✅ تم الانتهاء", tab8_completed, tab8_total)
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
        render_case_cards_pharmacy(payment_pending_df, False, pharmacist_name, pharmacy_name)
    with tab7:
        render_case_cards_pharmacy(cancelled_df, False, pharmacist_name, pharmacy_name)
    with tab8:
        from utils.ui_components import render_completed_table
        render_completed_table(completed_df, is_admin=False)
