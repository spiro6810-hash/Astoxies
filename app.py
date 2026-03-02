
import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="Αστοχίες Γραμμών", layout="wide")

# -------------------------
# Helpers
# -------------------------
def split_section(inst):
    """
    Από '3TN3 (ΝΟΜ-ΑΓ.ΠΑΡ)' βγάζει:
    - Τμήμα_Κωδικός = 3TN3
    - Τμήμα_Περιγραφή = ΝΟΜ-ΑΓ.ΠΑΡ
    """
    if pd.isna(inst):
        return (None, None)
    s = str(inst).strip()
    m = re.match(r"^([^\s(]+)\s*(?:\((.*?)\))?$", s)
    if not m:
        return (s, None)
    code = m.group(1).strip() if m.group(1) else None
    desc = m.group(2).strip() if m.group(2) else None
    return (code, desc)

def infer_line(inst):
    """
    Κατάταξη σε Γραμμή από prefix εγκατάστασης (όπως δουλεύεις ήδη).
    """
    if pd.isna(inst):
        return None
    s = str(inst).strip()
    code = re.split(r"\s|\(", s)[0].strip().upper()

    if code.startswith(("L1", "1")):
        return "Γραμμή 1"
    if code.startswith(("L2", "2")):
        return "Γραμμή 2"
    if code.startswith(("L3", "3")):
        return "Γραμμή 3"
    if code.startswith(("LS", "TWS")) or "TRAM" in code:
        return "Τραμ"

    # fallback αν το L1/L2/L3 δεν είναι μπροστά
    if "L1" in code:
        return "Γραμμή 1"
    if "L2" in code:
        return "Γραμμή 2"
    if "L3" in code:
        return "Γραμμή 3"
    if "LS" in code or "TWS" in code:
        return "Τραμ"

    return "Άγνωστο"

def normalize_yes(series):
    """
    Μετατρέπει διάφορες τιμές ΝΑΙ/YES/1/TRUE σε True.
    """
    s = series.astype(str).str.upper().str.strip()
    return s.isin(["ΝΑΙ", "YES", "Y", "1", "TRUE", "T"])

# -------------------------
# UI
# -------------------------
st.title("Αστοχίες Γραμμών (Excel από Access)")

uploaded = st.file_uploader("Ανέβασε Excel (.xlsx) με αστοχίες", type=["xlsx"])

if not uploaded:
    st.info("Ανέβασε ένα .xlsx για να ξεκινήσουμε.")
    st.stop()

# Read excel
df = pd.read_excel(uploaded, sheet_name=0)

# -------------------------
# Validate columns & build core fields
# -------------------------
# Date column
if "Ημερομηνία" in df.columns:
    df["Ημερομηνία"] = pd.to_datetime(df["Ημερομηνία"], errors="coerce")
elif "Ημ/νία" in df.columns:
    df["Ημερομηνία"] = pd.to_datetime(df["Ημ/νία"], errors="coerce")
else:
    st.error("Δεν βρέθηκε στήλη 'Ημ/νία' ή 'Ημερομηνία'.")
    st.stop()

# Installation column
if "Εγκατάσταση" not in df.columns:
    st.error("Δεν βρέθηκε στήλη 'Εγκατάσταση'.")
    st.stop()

# Add Line if missing
if "Γραμμή" not in df.columns:
    df["Γραμμή"] = df["Εγκατάσταση"].apply(infer_line)

# Add Section fields
df[["Τμήμα_Κωδικός", "Τμήμα_Περιγραφή"]] = df["Εγκατάσταση"].apply(lambda x: pd.Series(split_section(x)))

# Reference date = max date in file (για backlog)
ref_date = df["Ημερομηνία"].max()

# Fixed / Open
if "Επισκευάστηκε" in df.columns:
    is_fixed = normalize_yes(df["Επισκευάστηκε"])
    is_open = ~is_fixed
else:
    # Αν δεν έχεις στήλη, θεωρούμε όλα ανοιχτά
    is_open = pd.Series(True, index=df.index)

# Days open (μόνο για ανοιχτά)
df["Ημέρες_Ανοιχτό"] = pd.NA
df.loc[is_open, "Ημέρες_Ανοιχτό"] = (ref_date - df.loc[is_open, "Ημερομηνία"]).dt.days

# Month for grouping
df["Μήνας"] = df["Ημερομηνία"].dt.to_period("M").astype(str)

# -------------------------
# Sidebar filters
# -------------------------
st.sidebar.header("Φίλτρα")

# Date range filter (safe defaults)
min_date = df["Ημερομηνία"].min()
max_date = df["Ημερομηνία"].max()

date_from = st.sidebar.date_input("Από", value=min_date.date() if pd.notna(min_date) else None)
date_to = st.sidebar.date_input("Έως", value=max_date.date() if pd.notna(max_date) else None)

line_vals = sorted([x for x in df["Γραμμή"].dropna().unique().tolist()])
lines = st.sidebar.multiselect("Γραμμή", line_vals)

section_vals = sorted([x for x in df["Τμήμα_Κωδικός"].dropna().unique().tolist()])
sections = st.sidebar.multiselect("Τμήμα_Κωδικός", section_vals)

fault_vals = sorted([x for x in df["Αστοχία"].dropna().unique().tolist()]) if "Αστοχία" in df.columns else []
faults = st.sidebar.multiselect("Αστοχία", fault_vals) if fault_vals else []

sos_vals = sorted([x for x in df["SOS"].dropna().unique().tolist()]) if "SOS" in df.columns else []
sos = st.sidebar.multiselect("SOS", sos_vals) if sos_vals else []

only_open = st.sidebar.checkbox("Μόνο ανοιχτές (όχι επισκευασμένες)", value=False)

# Apply filters
f = df.copy()

if date_from and date_to:
    d1 = pd.to_datetime(date_from)
    d2 = pd.to_datetime(date_to) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    f = f[(f["Ημερομηνία"] >= d1) & (f["Ημερομηνία"] <= d2)]

if lines:
    f = f[f["Γραμμή"].isin(lines)]

if sections:
    f = f[f["Τμήμα_Κωδικός"].isin(sections)]

if "Αστοχία" in f.columns and faults:
    f = f[f["Αστοχία"].isin(faults)]

if "SOS" in f.columns and sos:
    f = f[f["SOS"].isin(sos)]

if only_open:
    open_mask = is_open.reindex(f.index).fillna(True)
    f = f[open_mask]

# -------------------------
# KPIs
# -------------------------
k1, k2, k3 = st.columns(3)
k1.metric("Σύνολο εγγραφών (με φίλτρα)", f"{len(f):,}".replace(",", "."))
k2.metric("Reference date", ref_date.strftime("%d/%m/%Y") if pd.notna(ref_date) else "-")
k3.metric("Ανοιχτές (με φίλτρα)", f"{int(is_open.reindex(f.index).fillna(True).sum()):,}".replace(",", "."))

st.divider()

# -------------------------
# Summary tables
# -------------------------
c1, c2 = st.columns([1, 1])

with c1:
    st.subheader("Σύνολο ανά Γραμμή")
    t_line = (f.groupby("Γραμμή")
              .size()
              .reset_index(name="Πλήθος")
              .sort_values("Πλήθος", ascending=False))
    st.dataframe(t_line, use_container_width=True)

with c2:
    st.subheader("Top 10 Τμήματα (ανά Γραμμή)")
    t_sections = (f.groupby(["Γραμμή", "Τμήμα_Κωδικός"])
                  .size()
                  .reset_index(name="Πλήθος")
                  .sort_values("Πλήθος", ascending=False)
                  .head(10))
    st.dataframe(t_sections, use_container_width=True)

st.subheader("Ανά μήνα (pivot)")
pivot = (f.pivot_table(index="Μήνας", columns="Γραμμή", values="Τμήμα_Κωδικός",
                       aggfunc="size", fill_value=0)
         .sort_index()
         .reset_index())
st.dataframe(pivot, use_container_width=True)

st.divider()

# -------------------------
# Backlog table
# -------------------------
st.subheader("Backlog (όχι επισκευασμένα) — ταξινόμηση κατά Ημέρες_Ανοιχτό")

open_mask = is_open.reindex(f.index).fillna(True)
open_df = f[open_mask].copy()

if len(open_df) == 0:
    st.info("Με τα φίλτρα που έχεις βάλει, δεν υπάρχουν ανοιχτές εγγραφές.")
else:
    # ωραίο sort: πρώτα τα μεγαλύτερα days open
    open_df = open_df.sort_values("Ημέρες_Ανοιχτό", ascending=False)

    # Διάλεξε “χρήσιμες” στήλες αν υπάρχουν
    cols_preferred = [
        "Ημερομηνία", "Γραμμή", "Τμήμα_Κωδικός", "Τμήμα_Περιγραφή",
        "Αστοχία", "SOS", "Εργασία", "Προτεινόμενη Ενέργεια",
        "Ενέργεια Επισκευής", "Ημ/νία Επισκευής", "Επισκευάστηκε",
        "Τροχιά", "ΧΘ", "TrackID", "Ημέρες_Ανοιχτό"
    ]
    cols = [c for c in cols_preferred if c in open_df.columns]

    st.dataframe(open_df[cols], use_container_width=True)

# -------------------------
# Download filtered data
# -------------------------
st.divider()
st.subheader("Export")

csv = f.to_csv(index=False).encode("utf-8-sig")
st.download_button("Κατέβασε τα δεδομένα (CSV)", data=csv, file_name="astoxies_filtered.csv", mime="text/csv")
