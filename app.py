import re
import io
from datetime import datetime, date

import pdfplumber
import pandas as pd
import streamlit as st


# --- Regex για γραμμή report ---
# Πιάνει και περιπτώσεις που ο αριθμός "κολλάει" με την ημερομηνία, π.χ. 4341906/7/25
LINE_RE = re.compile(r"^\s*(\d{1,2}/\d{1,2}/\d{2})\s+(\d{5,8})\s+(\d+)\s+([123S][A-Z0-9]{2,6})\s+(.*)$")

# Πιάνει κωδικό τμήματος που ξεκινά με 1/2/3/S, π.χ. 3DW1, 2DA1, 3T08, 2TS1
DEPTCODE_RE = re.compile(r"\b([123S][A-Z0-9]{2,6})\b")


def dept_from_access_deptcode(code: str) -> str:
    c = (code or "").strip().upper()
    if c.startswith("1"):
        return "Γραμμή 1"
    if c.startswith("2"):
        return "Γραμμή 2"
    if c.startswith("3"):
        return "Γραμμή 3"
    if c.startswith("S"):
        return "Τραμ"
    return "Άγνωστο"


def extract_open_from_access_pdf(file_bytes: bytes) -> pd.DataFrame:
    rows = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.splitlines():
                line = line.strip()
                if not line:
                    continue

                m = LINE_RE.match(line)
                if not m:
                    continue

                entoli = m.group(1)
                hmer = m.group(2)
                rest = m.group(3)

                dept_candidates = DEPTCODE_RE.findall(rest)
                dept_code = dept_candidates[-1] if dept_candidates else ""

                rows.append({
                    "Τμήμα": dept_from_access_deptcode(dept_code),
                    "Εντολή": entoli,
                    "Ημ/νία": hmer,
                    "Τμήμα_κωδ": dept_code,
                    "Raw": line,
                })

    df = pd.DataFrame(rows)
    if not df.empty:
        df["Ημ/νία_dt"] = pd.to_datetime(df["Ημ/νία"], format="%d/%m/%y", errors="coerce")
        today = pd.Timestamp(date.today())
        df["Ημέρες_ανοικτή"] = (today - df["Ημ/νία_dt"]).dt.days
    return df


def make_excel_bytes(summary_df: pd.DataFrame, details_df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Σύνοψη", index=False)
        details_df.to_excel(writer, sheet_name="Ανοιχτές_Λίστα", index=False)
    return output.getvalue()


# ---------------- Streamlit UI ----------------
st.set_page_config(page_title="Access Open Orders", layout="wide")
st.title("Ανοιχτές (Κενές) Εντολές Εργασίας από Access PDF")

uploaded = st.file_uploader("Ανέβασε το PDF (Access) με τις Κενές/Ανοιχτές εντολές", type=["pdf"])

if uploaded is None:
    st.info("Ανέβασε το PDF αναφοράς από Access για να δεις σύνοψη ανά τμήμα και λίστα ανοιχτών.")
    st.stop()

st.success(f"Ανέβηκε: {uploaded.name} ({uploaded.size} bytes)")

df_open = extract_open_from_access_pdf(uploaded.read())

if df_open.empty:
    st.error("Δεν βρέθηκαν γραμμές εντολών. Αν θες, κάνε copy-paste 2-3 γραμμές όπως φαίνονται στο PDF.")
    st.stop()

# Quick filters
st.subheader("Γρήγορα φίλτρα")
c1, c2, c3, c4, c5 = st.columns(5)
if "quick_dept" not in st.session_state:
    st.session_state.quick_dept = "Όλα"

with c1:
    if st.button("Όλα", use_container_width=True):
        st.session_state.quick_dept = "Όλα"
with c2:
    if st.button("Μόνο Γραμμή 1", use_container_width=True):
        st.session_state.quick_dept = "Γραμμή 1"
with c3:
    if st.button("Μόνο Γραμμή 2", use_container_width=True):
        st.session_state.quick_dept = "Γραμμή 2"
with c4:
    if st.button("Μόνο Γραμμή 3", use_container_width=True):
        st.session_state.quick_dept = "Γραμμή 3"
with c5:
    if st.button("Μόνο Τραμ", use_container_width=True):
        st.session_state.quick_dept = "Τραμ"

# Filters
col1, col2 = st.columns(2)
with col1:
    dept_options = sorted(df_open["Τμήμα"].unique().tolist())
    default_dept = dept_options
    if st.session_state.quick_dept != "Όλα" and st.session_state.quick_dept in dept_options:
        default_dept = [st.session_state.quick_dept]
    dept = st.multiselect("Τμήμα", dept_options, default=default_dept)

with col2:
    age_bucket = st.selectbox("Παλαιότητα", ["Όλες", "> 7 μέρες", "> 30 μέρες"], index=0)

filtered = df_open[df_open["Τμήμα"].isin(dept)].copy()
if age_bucket == "> 7 μέρες":
    filtered = filtered[filtered["Ημέρες_ανοικτή"] > 7]
elif age_bucket == "> 30 μέρες":
    filtered = filtered[filtered["Ημέρες_ανοικτή"] > 30]

# Summary
summary = (
    filtered.groupby("Τμήμα")["Εντολή"]
    .count()
    .rename("Ανοιχτές (Access)")
    .reset_index()
    .sort_values("Τμήμα")
)

aging = (
    filtered.groupby("Τμήμα")["Ημέρες_ανοικτή"]
    .agg(
        Σύνολο="count",
        Πάνω_από_7=lambda s: int((s > 7).sum()),
        Πάνω_από_30=lambda s: int((s > 30).sum()),
        Max_ημέρες=lambda s: int(pd.to_numeric(s, errors="coerce").max()) if len(s) else 0,
    )
    .reset_index()
    .sort_values("Τμήμα")
)

st.subheader("Σύνοψη ανά τμήμα")
st.dataframe(summary, use_container_width=True)

st.subheader("Παλαιότητα ανά τμήμα")
st.dataframe(aging, use_container_width=True)

st.subheader("Λίστα ανοιχτών")
show_cols = ["Τμήμα", "Εντολή", "Ημ/νία", "Ημέρες_ανοικτή", "Τμήμα_κωδ", "Raw"]
filtered_view = filtered[show_cols].sort_values(["Τμήμα", "Ημέρες_ανοικτή"], ascending=[True, False])
st.dataframe(filtered_view, use_container_width=True, height=520)

# Excel export
excel_bytes = make_excel_bytes(aging, filtered_view)
stamp = datetime.now().strftime("%Y%m%d_%H%M")
st.download_button(
    "⬇️ Κατέβασε Excel",
    data=excel_bytes,
    file_name=f"access_open_orders_{stamp}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

