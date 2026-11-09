"""
app.py

Streamlit web interface for the WhatsApp attendance agent.
Run with: streamlit run app.py
"""

import os
from datetime import date

import pandas as pd
import streamlit as st

from agent_logic import (
    extract_attendance_records,
    fill_attendance_excel,
    parse_alias_text,
)

st.set_page_config(page_title="WhatsApp Attendance Agent", page_icon="📋", layout="wide")


def get_default_api_key():
    """Read the Groq key from Streamlit secrets (cloud) or from the environment (local)."""
    try:
        return st.secrets["GROQ_API_KEY"]
    except Exception:
        return os.getenv("GROQ_API_KEY", "")

st.title("📋 WhatsApp Attendance Agent")
st.caption(
    "WhatsApp hazri ka text paste karein, company template upload karein, "
    "aur final Excel download karein."
)

# ---------------- Sidebar settings ----------------
with st.sidebar:
    st.header("⚙️ Settings")

    api_key = st.text_input(
        "Groq API Key",
        type="password",
        value=get_default_api_key(),
        help="Aap GROQ_API_KEY environment variable bhi set kar sakte hain.",
    )

    name_col = st.number_input(
        "Naam wala column number",
        min_value=1,
        value=2,
        step=1,
        help="Template mein jis column mein employees ke naam hain. NAME column B hai, is liye 2.",
    )

    header_row = st.number_input(
        "Din/tareekh wali header row number",
        min_value=1,
        value=2,
        step=1,
        help="Template ki wo row jis mein din (1, 2, 3...) ya tareekhen likhi hain. Aapke template mein ye row 2 hai.",
    )

    sheet_name = st.text_input(
        "Sheet ka naam (khali chhor dein to pehli sheet)",
        value="",
    )

    sheet_month = st.date_input(
        "Sheet ka mahina (us mahine ka koi bhi din chunein)",
        value=date.today(),
        help="Template ka mahina. Is mahine se alag tareekh wali entries skip hongi.",
    )

    default_date = st.date_input(
        "Default tareekh (agar chat mein tareekh na ho)",
        value=date.today(),
    )

# ---------------- Main inputs ----------------
col_left, col_right = st.columns(2)

with col_left:
    chat_text = st.text_area(
        "WhatsApp chat paste karein",
        height=320,
        placeholder="01/11/2026, 08:52 - +92 321 5856154: Present at Askari grid site\n...",
    )

    alias_text = st.text_area(
        "Naam ya number ka mapping (optional)",
        height=120,
        placeholder="+92 321 5856154 = Sarfraz Khan\nsarfraz = Sarfraz Khan",
        help="Har line mein: chat wala naam ya number = Excel template wala naam. Jo naam template mein hai wahi likhein.",
    )

with col_right:
    uploaded_file = st.file_uploader("Company Excel template (.xlsx)", type=["xlsx"])

# ---------------- Processing ----------------
if st.button("🚀 Attendance process karein", type="primary", use_container_width=True):
    if not api_key.strip():
        st.error("Pehle sidebar mein Groq API key daalein.")
    elif not chat_text.strip():
        st.error("WhatsApp chat ka text paste karein.")
    elif uploaded_file is None:
        st.error("Excel template upload karein.")
    else:
        try:
            with st.spinner("Groq se attendance nikal rahe hain..."):
                records = extract_attendance_records(
                    chat_text,
                    api_key.strip(),
                    default_date=default_date,
                )

            if not records:
                st.warning("Chat mein koi attendance entry nahi mili.")
            else:
                st.subheader("1. Extract hui attendance")
                preview_df = pd.DataFrame(
                    [
                        {
                            "Naam": r["name"],
                            "Tareekh": r["date"].strftime("%d/%m/%Y"),
                            "Status": r["status"],
                        }
                        for r in records
                    ]
                )
                st.dataframe(preview_df, use_container_width=True, hide_index=True)

                with st.spinner("Excel template update ho raha hai..."):
                    excel_bytes, summary = fill_attendance_excel(
                        uploaded_file.getvalue(),
                        records,
                        name_col=int(name_col),
                        header_row=int(header_row),
                        sheet_month=sheet_month,
                        sheet_name=sheet_name.strip() or None,
                        aliases=parse_alias_text(alias_text),
                    )

                st.subheader("2. Result")
                st.success(
                    f"{summary['written']} out of {summary['total_records']} entries "
                    "Excel mein darj ho gayi."
                )

                if summary["unmatched"]:
                    st.warning(
                        "Yeh naam template mein nahi mile. Inhe mapping box mein "
                        "'chat wala naam = template wala naam' ki shakl mein likhein:"
                    )
                    for name, suggestion in summary["unmatched"].items():
                        if suggestion:
                            st.write(f"- **{name}** (kya yeh **{suggestion}** hai?)")
                        else:
                            st.write(f"- **{name}**")

                if summary["skipped_dates"]:
                    st.info(
                        "Yeh entries skip hui (mahina ya din template se match nahi karta): "
                        + ", ".join(summary["skipped_dates"])
                    )

                st.download_button(
                    label="📥 Final Excel download karein",
                    data=excel_bytes,
                    file_name="attendance_final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )

        except ValueError as exc:
            st.error(str(exc))
        except Exception as exc:  # Show unexpected errors without crashing the app
            st.error(f"Kuch ghalat ho gaya: {exc}")
