
### `app.py` (βασικό skeleton: upload + στήλες Τμήμα + backlog)
```python
import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="Αστοχίες Γραμμών", layout="wide")

def split_section(inst):
    if pd.isna(inst):
        return (None, None)
    s = str(inst).strip()
    m = re.match(r'^([^\s(]+)\s*(?:\((.*?)\))?$', s)
    if not m:
        return (s, None)
    code = m.group(1).strip() if m.group(1) else None
    desc = m.group(2).strip() if m.group(2) else None
    return (code, desc)

def infer_line(inst):
    if pd.isna(inst):
        return None
    s = str(inst).strip()
    code = re.split(r'\s|\(', s)[0].strip().upper()
    if code.startswith(("L1", "1")): return "Γραμμή 1"
    if code.startswith(("L2", "2")): return "Γραμμή 2"
    if code.startswith(("L3", "3")): return "Γραμμή 3"
    if code.startswith(("LS", "TWS")) or "TRAM" in code: return "Τραμ"
    if "L1" in code: return "Γραμμή 1"
    if "L2" in code: return "Γραμμή 2"
    if "L3" in code: return "Γραμμή 3"
    if "LS" in code or "TWS" in code: return "Τραμ"
    return "Άγνωστο"

st.title("Αστοχίες Γραμμών")

uploaded = st.file_uploader("Ανέβασε Excel (Astoxies export)", type=["xlsx"])

if not uploaded:
    st.info("Ανέβασε ένα .xlsx για να ξεκινήσουμε.")
    st.stop()

df = pd.read_excel(uploaded, sheet_name=0)

# Ημερομηνία
if "Ημερομηνία" in df.columns:
    df["Ημερομηνία"] = pd.to_datetime(df["Ημερομηνία"], errors="coerce")
elif "Ημ/νία" in df.columns:
    df["Ημερομηνία"] = pd.to_datetime(df["Ημ/νία"], errors="coerce")
else:
    st.error("Δεν βρέθηκε στήλη 'Ημ/νία' ή 'Ημερομηνία'.")
    st.stop()

# Γραμμή (αν δεν υπάρχει ήδη)
if "Γραμμή" not in df.columns and "Εγκατάσταση" in df.columns:
    df["Γραμμή"] = df["Εγκατάσταση"].apply(infer_line)

# Τμήμα (από Εγκατάσταση)
if "Εγκατάσταση" in df.columns:
    df[["Τμήμα_Κωδικός","Τμήμα_Περιγραφή"]] = df["Εγκατάσταση"].apply(lambda x: pd.Series(split_section(x)))
else:
    st.error("Δεν βρέθηκε στήλη 'Εγκατάσταση'.")
    st.stop()

# Reference date = max ημερομηνία αρχείου
ref_date = df["Ημερομηνία"].max()

# Backlog days (ανοιχτά)
fixed_col = "Επισκευάστηκε" if "Επισκευάστηκε" in df.columns else None
if fixed_col:
    fixed = df[fixed_col].astype(str).str.upper().str.strip()
    is_open = ~fixed.isin(["ΝΑΙ", "YES", "Y", "1", "TRUE"])
else:
    is_open = pd.Series(True, index=df.index)

df["Ημέρες_Ανοιχτό"] = None
df.loc[is_open, "Ημέρες_Ανοιχτό"] = (ref_date - df.loc[is_open, "Ημερομηνία"]).dt.days

# Sidebar filters
st.sidebar.header("Φίλτρα")
line = st.sidebar.multiselect("Γραμμή", sorted(df["Γραμμή"].dropna().unique().tolist()))
section = st.sidebar.multiselect("Τμήμα_Κωδικός", sorted(df["Τμήμα_Κωδικός"].dropna().unique().tolist()))
sos = st.sidebar.multiselect("SOS", sorted(df["SOS"].dropna().unique().tolist())) if "SOS" in df.columns else []
fault = st.sidebar.multiselect("Αστοχία", sorted(df["Αστοχία"].dropna().unique().tolist())) if "Αστοχία" in df.columns else []

f = df.copy()
if line: f = f[f["Γραμμή"].isin(line)]
if section: f = f[f["Τμήμα_Κωδικός"].isin(section)]
if "SOS" in f.columns and sos: f = f[f["SOS"].isin(sos)]
if "Αστοχία" in f.columns and fault: f = f[f["Αστοχία"].isin(fault)]

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("Σύνολο ανά Γραμμή")
    st.dataframe(
        f.groupby("Γραμμή").size().reset_index(name="Πλήθος").sort_values("Πλήθος", ascending=False),
        use_container_width=True
    )

with col2:
    st.subheader("Top 10 Τμήματα")
    st.dataframe(
        f.groupby(["Γραμμή","Τμήμα_Κωδικός"]).size().reset_index(name="Πλήθος")
         .sort_values("Πλήθος", ascending=False).head(10),
        use_container_width=True
    )

st.subheader("Backlog (όχι επισκευασμένα) — ταξινόμηση κατά Ημέρες_Ανοιχτό")
open_mask = is_open.reindex(f.index).fillna(True)
open_df = f[open_mask].copy()
open_df = open_df.sort_values("Ημέρες_Ανοιχτό", ascending=False)
st.dataframe(open_df, use_container_width=True)
