import os
import re
import sqlite3
import uuid
from datetime import datetime

import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="نظام بلسم - مطابقة الطلبات والفواتير",
    layout="wide",
    initial_sidebar_state="expanded",
)


DB_DIR = "data"
DB_PATH = os.path.join(DB_DIR, "pharmacy_reconciliation.db")
PHARMACY_COUNT = 17
SPECIAL_ORDER_NUMBERS = {"0", "123456"}
EXCLUDED_PROFILE = "FREE GIFTS FOR CUSTOMERS"
CASE_LABELS = {
    "addition": "إضافة",
    "return": "إرجاع",
    "orphan_salla": "طلب بدون فاتورة",
    "orphan_abc": "فاتورة بدون طلب",
    "branch_mismatch": "اختلاف فرع",
    "special_review": "مراجعة رقم طلب خاص",
}
STATUS_DONE = "تم"
STATUS_PENDING = "قيد المتابعة"


st.markdown(
    """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@300;400;500;700;800&display=swap');

    * { font-family: 'Tajawal', sans-serif; }

    .hero {
        background:
            radial-gradient(circle at top right, rgba(255,255,255,0.18), transparent 28%),
            linear-gradient(135deg, #0f4c5c 0%, #1f7a8c 50%, #16425b 100%);
        border-radius: 24px;
        padding: 2.2rem;
        color: white;
        margin-bottom: 1.6rem;
        box-shadow: 0 18px 40px rgba(22, 66, 91, 0.20);
    }

    .hero h1 {
        margin: 0;
        font-size: 2.2rem;
        font-weight: 800;
    }

    .hero p {
        margin-top: 0.6rem;
        font-size: 1rem;
        opacity: 0.95;
    }

    .section-title {
        font-size: 1.15rem;
        font-weight: 800;
        color: #16425b;
        border-right: 5px solid #1f7a8c;
        padding-right: 0.65rem;
        margin: 1rem 0 0.8rem;
    }

    .note-card {
        background: linear-gradient(135deg, #f4fbfc 0%, #ffffff 100%);
        border: 1px solid #d7ebef;
        border-radius: 16px;
        padding: 1rem 1.1rem;
        margin-bottom: 1rem;
    }

    .action-card {
        background: white;
        border: 1px solid #e4eef1;
        border-right: 6px solid #1f7a8c;
        border-radius: 18px;
        padding: 1rem;
        margin-bottom: 0.85rem;
        box-shadow: 0 8px 22px rgba(15, 76, 92, 0.07);
    }

    .pill {
        display: inline-block;
        padding: 0.28rem 0.75rem;
        border-radius: 999px;
        font-size: 0.78rem;
        font-weight: 700;
    }

    .pill-green { background: #dff7e8; color: #0f7a3a; }
