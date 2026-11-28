import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from io import BytesIO
from datetime import datetime
from utils.excel_processor import update_balances

def show():
    st.markdown(
        """
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@300;400;500;700;800&display=swap');
            * { font-family: 'Tajawal', sans-serif; }
        </style>
        <div class="hero">
            <h1>🔄 تحديث أرصدة الفروع</h1>
            <p>رفع ملفات ABC و Salla لتحديث الأرصدة وبناء مستند التحديث الثنائي المتوافق مع المنصة</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    
    balance_tabs = st.tabs(["📦 كميات الفروع", "🏬 كميات الفروع مع المستودع"])

    with balance_tabs[0]:
        st.info("تُحسب الكميات بالمعادلات الأساسية الحالية دون إضافة رصيد المستودع.")
        col1, col2, col3 = st.columns(3)
        with col1:
            abc_file = st.file_uploader("📊 رفع ملف ABC (يبدأ من الصف 5)", type=["xlsx"], key="abc_balances_base")
        with col2:
            salla_file = st.file_uploader("📋 رفع ملف Salla", type=["xlsx"], key="salla_balances_base")
        with col3:
            exclude_file = st.file_uploader("🚫 رفع ملف الاستبعادات (اختياري)", type=["xlsx"], key="exclude_balances_base", help="ملف يحتوي على أرقام SKU المطلوب استبعادها من التحديث")
        _render_balance_update(abc_file, salla_file, exclude_file, mode_key="base")

    with balance_tabs[1]:
        st.info(
            "ارفع ملف ABC الذي يحتوي على عمود المستودع في T، وارفع ملف SKU للمنتجات المستهدفة. "
            "تُضاف 25% من رصيد T إلى المعادلة الأساسية في العمود F (تبوك القادسية وباقي المدن) "
            "لهذه المنتجات فقط، ثم يُقرّب الناتج للأعلى. جميع المنتجات الأخرى وبقية الفروع تبقى على المعادلات الأساسية."
        )
        st.download_button(
            "📥 تنزيل نموذج قائمة المنتجات المستهدفة",
            data=_warehouse_sku_template(),
            file_name="Warehouse_Bonus_Products.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="warehouse_sku_template",
        )
        col1, col2, col3, col4 = st.columns([1.1, 1.1, 1, 1])
        with col1:
            abc_file_wh = st.file_uploader(
                "🏬 رفع ملف ABC شامل رصيد المستودع (العمود T)",
                type=["xlsx"], key="abc_balances_warehouse",
                help="يتحقق النظام من أن عنوان العمود T هو Balsam Alula Warehouse."
            )
        with col2:
            warehouse_products_file = st.file_uploader(
                "📦 رفع منتجات تشمل كمية المستودع",
                type=["xlsx", "csv"], key="warehouse_bonus_products",
                help="ملف Excel أو CSV يحتوي على SKU المنتجات المستهدفة في العمود الأول. ستُطبق زيادة المستودع على هذه المنتجات فقط."
            )
        with col3:
            salla_file_wh = st.file_uploader("📋 رفع ملف Salla", type=["xlsx"], key="salla_balances_warehouse")
        with col4:
            exclude_file_wh = st.file_uploader(
                "🚫 رفع ملف الاستبعادات (اختياري)", type=["xlsx"], key="exclude_balances_warehouse",
                help="ملف يحتوي على أرقام SKU المطلوب استبعادها من التحديث."
            )
        _render_balance_update(
            abc_file_wh, salla_file_wh, exclude_file_wh, mode_key="warehouse",
            include_warehouse=True, warehouse_products_file=warehouse_products_file
        )


def _warehouse_sku_template():
    frame = pd.DataFrame(columns=["SKU"])
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        frame.to_excel(writer, index=False, sheet_name="Warehouse Products")
    return output.getvalue()


def _normalize_balance_sku(value):
    if value is None or pd.isna(value):
        return ""
    text = str(value).strip()
    return text[:-2] if text.endswith(".0") else text


def _render_balance_update(
    abc_file, salla_file, exclude_file, mode_key,
    include_warehouse=False, warehouse_products_file=None,
):
    if include_warehouse and not warehouse_products_file:
        st.info("ارفع ملف «منتجات تشمل كمية المستودع» لتحديد المنتجات التي ستضاف لها 25% في العمود F.")
        return
    if not abc_file or not salla_file:
        return

    warehouse_skus = None
    if include_warehouse:
        try:
            product_frame = (
                pd.read_csv(BytesIO(warehouse_products_file.getvalue()))
                if warehouse_products_file.name.lower().endswith(".csv")
                else pd.read_excel(BytesIO(warehouse_products_file.getvalue()))
            )
            sku_column = product_frame.columns[0]
            for column in product_frame.columns:
                label = str(column).strip().lower()
                if "sku" in label or "رمز" in label or "رقم" in label:
                    sku_column = column
                    break
            warehouse_skus = {
                _normalize_balance_sku(value)
                for value in product_frame[sku_column].dropna().tolist()
                if _normalize_balance_sku(value)
            }
            if not warehouse_skus:
                st.error("لم أجد أرقام SKU في ملف المنتجات. ضع SKU في العمود الأول ثم أعد رفع الملف.")
                return

            abc_skus_frame = pd.read_excel(
                BytesIO(abc_file.getvalue()), skiprows=4, usecols=[0]
            )
            abc_skus = {
                _normalize_balance_sku(value)
                for value in abc_skus_frame.iloc[:, 0].dropna().tolist()
                if _normalize_balance_sku(value)
            }
            matched_skus = warehouse_skus & abc_skus
            unmatched_skus = warehouse_skus - abc_skus
            if not matched_skus:
                st.error("لا يوجد SKU من ملف المنتجات مطابق لملف ABC.")
                return
            st.caption(
                f"منتجات مستهدفة: {len(warehouse_skus):,} · "
                f"مطابقة في ABC: {len(matched_skus):,} · غير موجودة في ABC: {len(unmatched_skus):,}"
            )
            if unmatched_skus:
                st.warning("لن تُضاف زيادة المستودع للـSKU غير الموجودة في ABC: " + ", ".join(sorted(unmatched_skus)[:8]))
            warehouse_skus = matched_skus
        except (ValueError, TypeError, KeyError, OSError, pd.errors.ParserError) as exc:
            st.error(f"تعذر قراءة ملف المنتجات المستهدفة: {type(exc).__name__}")
            return

    if st.button("🔄 تنفيذ تحديث الأرصدة", use_container_width=True, key=f"update_balances_{mode_key}"):
        with st.spinner("جاري تحديث الأرصدة وبناء الترويسات الثنائية..."):
            result_df, result = update_balances(
                abc_file, salla_file,
                warehouse_bonus_rate=0.25 if include_warehouse else 0.0,
                require_warehouse=include_warehouse,
                warehouse_skus=warehouse_skus,
            )
            if result_df is not None:
                    
                excluded_count = 0
                excluded_df = pd.DataFrame()
                    
                if exclude_file is not None:
                    try:
                        exclude_df_upload = pd.read_excel(exclude_file)
                        sku_col = exclude_df_upload.columns[0]
                        for col in exclude_df_upload.columns:
                            if 'sku' in str(col).lower() or 'رمز' in str(col) or 'رقم' in str(col):
                                sku_col = col
                                break
                            
                        excluded_skus = exclude_df_upload[sku_col].astype(str).str.replace('.0', '', regex=False).str.strip().tolist()
                            
                        temp_skus = result_df.iloc[:, 3].astype(str).str.replace('.0', '', regex=False).str.strip()
                            
                        # 💡 فصل الأصناف المستبعدة في جدول خاص لعرضها
                        mask = temp_skus.isin(excluded_skus)
                        excluded_df = result_df[mask].copy()
                        result_df = result_df[~mask]
                            
                        excluded_count = len(excluded_df)
                            
                    except Exception as e:
                        st.warning(f"⚠️ لم يتم استبعاد الأصناف. تأكد من صحة ملف الاستبعادات: {e}")

                # 💡 عرض رسالة النجاح وجدول المستبعدات إذا وجدت
                if excluded_count > 0:
                    st.success(f"✅ تم التحديث بنجاح! تم تجهيز {len(result_df):,} صنف وتم استبعاد {excluded_count:,} صنف.")
                    with st.expander(f"🚫 عرض قائمة الأصناف المستبعدة ({excluded_count} صنف)"):
                        st.dataframe(excluded_df, use_container_width=True)
                else:
                    st.success(f"✅ تم التحديث بنجاح! عدد الأصناف المحدثة والمعدلة: {len(result_df):,}")
                        
                st.dataframe(result_df.head(20), use_container_width=True)
                    
                # 🧠 بناء المصنف سحابياً ببنية الصفين المتطابقة مع منصة سلة
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Salla Product Quantities Sheet"
                    
                font_title = Font(name="Tajawal", size=11, bold=True, color="FFFFFF")
                font_headers = Font(name="Tajawal", size=10, bold=True, color="FFFFFF")
                font_data = Font(name="Tajawal", size=10)
                    
                fill_title_prod = PatternFill(start_color="16425B", end_color="16425B", fill_type="solid")
                fill_title_qty = PatternFill(start_color="1F7A8C", end_color="1F7A8C", fill_type="solid")
                fill_headers = PatternFill(start_color="4A5568", end_color="4A5568", fill_type="solid")
                    
                align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
                border_thin = Border(
                    left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'),
                    top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1')
                )
                    
                ws.merge_cells("A1:E1")
                ws["A1"] = "بيانات المنتج"
                ws.merge_cells("F1:AO1")
                ws["F1"] = "الكميات"
                    
                for col in range(1, 6):
                    cell = ws.cell(row=1, column=col)
                    cell.font = font_title
                    cell.fill = fill_title_prod
                    cell.alignment = align_center
                    cell.border = border_thin
                        
                for col in range(6, 42):
                    cell = ws.cell(row=1, column=col)
                    cell.font = font_title
                    cell.fill = fill_title_qty
                    cell.alignment = align_center
                    cell.border = border_thin
                        
                headers_row2 = [
                    "No.", "النوع", "أسم المنتج", "رمز المنتج sku", "إعداد الكمية", 
                    "الكمية في فرع تبوك القادسية وباقي المدن", "العرض في فرع تبوك القادسية وباقي المدن", 
                    "الكمية في فرع تبوك - صيدلية بلسم العلا 8 النظيم", "العرض في فرع تبوك - صيدلية بلسم العلا 8 النظيم", 
                    "الكمية في فرع العلا - صيدلية بلسم العلا 9", "العرض في فرع العلا - صيدلية بلسم العلا 9", 
                    "الكمية في فرع تبوك - صيدلية بلسم العلا 11 البوادي", "العرض في فرع تبوك - صيدلية بلسم العلا 11 البوادي", 
                    "الكمية في فرع تبوك - صيدلية بلسم العلا 15 الصفا", "العرض في فرع تبوك - صيدلية بلسم العلا 15 الصفا", 
                    "الكمية في فرع تبوك - صيدلية بلسم العلا 16 النخيل", "العرض في فرع تبوك - صيدلية بلسم العلا 16 النخيل", 
                    "الكمية في فرع تبوك - صيدلية بلسم العلا 10 الريان", "العرض في فرع تبوك - صيدلية بلسم العلا 10 الريان", 
                    "الكمية في فرع تبوك صيدلية بلسم العلا 13 القادسية", "العرض في فرع تبوك صيدلية بلسم العلا 13 القادسية", 
                    "الكمية في فرع تبوك - صيدلية بلسم العلا 12 العليا", "العرض في فرع تبوك - صيدلية بلسم العلا 12 العليا", 
                    "الكمية في فرع تبوك - صيدلية بلسم العلا 14 المصيف", "العرض في فرع تبوك - صيدلية بلسم العلا 14 المصيف", 
                    "الكمية في فرع العلا - صيدلية بلسم العلا 1", "العرض في فرع العلا - صيدلية بلسم العلا 1", 
                    "الكمية في فرع العلا -- صيدلية بلسم العلا 2", "العرض في فرع العلا -- صيدلية بلسم العلا 2", 
                    "الكمية في فرع العلا - صيدلية بلسم العلا 3", "العرض في فرع العلا - صيدلية بلسم العلا 3", 
                    "الكمية في فرع العلا - صيدلية بلسم العلا 4 العذيب", "العرض في فرع العلا - صيدلية بلسم العلا 4 العذيب", 
                    "الكمية في فرع العلا - صيدلية بلسم العلا 5", "العرض في فرع العلا - صيدلية بلسم العلا 5", 
                    "الكمية في فرع العلا - صيدلية بلسم العلا 6 ابو راكة", "العرض في فرع العلا - صيدلية بلسم العلا 6 ابو راكة", 
                    "الكمية في فرع العلا - صيدلية بلسم العلا 7", "العرض في فرع العلا - صيدلية بلسم العلا 7", 
                    "الكمية في فرع تبوك - صيدلية بلسم العلا 17 الروضة", "العرض في فرع تبوك - صيدلية بلسم العلا 17 الروضة"
                ]
                    
                for col_idx, h_text in enumerate(headers_row2, 1):
                    cell = ws.cell(row=2, column=col_idx, value=h_text)
                    cell.font = font_headers
                    cell.fill = fill_headers
                    cell.alignment = align_center
                    cell.border = border_thin
                        
                offer_columns = [i + 1 for i, val in enumerate(headers_row2) if str(val).startswith("العرض في")]
                col_13_offer_idx = headers_row2.index("العرض في فرع تبوك صيدلية بلسم العلا 13 القادسية") + 1
                col_7_offer_idx = headers_row2.index("العرض في فرع العلا - صيدلية بلسم العلا 7") + 1

                for r_idx, row_values in enumerate(result_df.values, start=3):
                    for c_idx, val in enumerate(row_values, start=1):
                        if c_idx in offer_columns:
                            if c_idx == col_13_offer_idx or c_idx == col_7_offer_idx:
                                val = "لا"
                            else:
                                val = "نعم"
                        else:
                            if isinstance(val, (np.integer, np.floating)):
                                val = val.item()
                            elif pd.isna(val):
                                val = ""
                                    
                        cell = ws.cell(row=r_idx, column=c_idx, value=val)
                        cell.font = font_data
                        cell.alignment = align_center
                        cell.border = border_thin
                            
                ws.row_dimensions[1].height = 28
                ws.row_dimensions[2].height = 24
                    
                for col_cols in ws.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col_cols)
                    col_letter = openpyxl.utils.get_column_letter(col_cols[0].column)
                    ws.column_dimensions[col_letter].width = min(max(max_len + 3, 11), 32)
                    
                output = BytesIO()
                wb.save(output)
                output.seek(0)
                    
                st.download_button(
                    "📥 تحميل ملف Salla المحدث",
                    data=output,
                    file_name=f"salla_updated_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            else:
                st.error(f"❌ خطأ في التحديث: {result}")
