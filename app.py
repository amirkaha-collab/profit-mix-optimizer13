import itertools
from io import BytesIO

import pandas as pd
import streamlit as st

# =============================
# UI
# =============================
st.set_page_config(page_title="מנוע הקצאה – שילובי מסלולים (1–3)", layout="wide")

DARK_RTL_CSS = """
<style>
html, body, [class*="css"]  { direction: rtl; text-align: right; }
div[data-testid="stAppViewContainer"]{ background: #0b0f17; }
div[data-testid="stSidebar"]{ background: #0a0d14; border-left: 1px solid rgba(255,255,255,0.08); }
h1, h2, h3, h4, h5, h6, p, label, span, div { color: rgba(255,255,255,0.92) !important; }
.stButton button, .stDownloadButton button { border-radius: 14px; }
[data-testid="stDataFrame"] { background: rgba(255,255,255,0.02); border-radius: 14px; border: 1px solid rgba(255,255,255,0.08); padding: 6px; }
div[data-baseweb="select"] > div { background: rgba(255,255,255,0.04) !important; }
</style>
"""
st.markdown(DARK_RTL_CSS, unsafe_allow_html=True)

st.title("מנוע הקצאה – שילובי מסלולים (1–3)")
st.caption("העלה Excel/CSV. אם זה Excel בפורמט רחב (כמו הקובץ שלך), המערכת תחלץ ציון מתוך השורה 'מדד שארפ' בכל לשונית.")

# =============================
# Parsers
# =============================
REQUIRED_COLS = {"provider", "track", "score"}


def _as_str(x) -> str:
    return "" if pd.isna(x) else str(x).strip()


def parse_normalized_table(df: pd.DataFrame) -> pd.DataFrame:
    """Expect columns: provider, track, score (case-insensitive)."""
    if df is None or df.empty:
        raise ValueError("Empty table")

    cols = {str(c).strip().lower(): c for c in df.columns}
    missing = [c for c in REQUIRED_COLS if c not in cols]
    if missing:
        raise KeyError(f"Missing columns: {missing}")

    out = df[[cols["provider"], cols["track"], cols["score"]]].copy()
    out.columns = ["provider", "track", "score"]
    out["provider"] = out["provider"].astype(str).str.strip()
    out["track"] = out["track"].astype(str).str.strip()
    out["score"] = pd.to_numeric(out["score"], errors="coerce")
    out = out.dropna(subset=["provider", "track", "score"])
    return out


def parse_hebrew_matrix_sheet(sheet_df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    """Parse wide Hebrew matrix sheets.

    Expected layout:
      row 0: ["פרמטר", <provider1>, <provider2>, ...]
      col 0: metric names; we pick the first row containing "שארפ" as the score row.
    """
    if sheet_df is None or sheet_df.empty:
        raise ValueError("Empty sheet")

    df = sheet_df.copy()
    df.columns = list(range(df.shape[1]))
    df[0] = df[0].apply(_as_str)

    providers = [_as_str(x) for x in df.iloc[0, 1:].tolist() if _as_str(x)]
    if not providers:
        raise ValueError("Could not detect provider header row")

    sharpe_rows = df[0].str.contains("שארפ", na=False)
    if not sharpe_rows.any():
        raise ValueError("Could not find Sharpe row (שארפ)")

    sharpe_idx = sharpe_rows[sharpe_rows].index[0]
    scores = df.loc[sharpe_idx, 1 : 1 + len(providers) - 1]

    out = pd.DataFrame(
        {
            "provider": providers,
            "track": [sheet_name] * len(providers),
            "score": pd.to_numeric(list(scores), errors="coerce"),
        }
    )
    out = out.dropna(subset=["provider", "track", "score"])
    return out


def parse_excel_workbook(file_bytes: bytes) -> pd.DataFrame:
    xls = pd.ExcelFile(BytesIO(file_bytes))

    # Try normalized table first
    for s in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=s)
            return parse_normalized_table(df)
        except Exception:
            pass

    # Fallback: wide Hebrew matrix per-sheet
    frames = []
    for s in xls.sheet_names:
        try:
            sheet_df = pd.read_excel(xls, sheet_name=s, header=None)
            frames.append(parse_hebrew_matrix_sheet(sheet_df, s))
        except Exception:
            continue

    if not frames:
        raise ValueError("לא הצלחתי לפרש את קובץ האקסל. אם תרצה, העלה CSV עם עמודות: provider,track,score")
    return pd.concat(frames, ignore_index=True)


def parse_upload(uploaded_file) -> pd.DataFrame:
    name = (uploaded_file.name or "").lower()
    data = uploaded_file.getvalue()

    if name.endswith(".csv"):
        df = pd.read_csv(BytesIO(data))
        return parse_normalized_table(df)

    if name.endswith(".xlsx") or name.endswith(".xls"):
        return parse_excel_workbook(data)

    raise ValueError("סוג קובץ לא נתמך. העלה Excel או CSV")


# =============================
# Sidebar inputs
# =============================
with st.sidebar:
    st.header("קלט נתונים")
    uploaded = st.file_uploader("קובץ מסלולים (Excel/CSV)", type=["xlsx", "xls", "csv"])

    st.divider()
    st.header("הגדרות שילוב")
    k = st.radio("כמה מסלולים לשלב?", options=[1, 2, 3], horizontal=True)


# =============================
# Main
# =============================
if not uploaded:
    st.info("העלה קובץ כדי להתחיל.")
    st.stop()

try:
    data = parse_upload(uploaded)
except Exception as e:
    st.error(f"שגיאה בקריאת הקובץ: {e}")
    st.stop()

# Basic cleanup
for col in ["provider", "track"]:
    data[col] = data[col].astype(str).str.strip()

data = data.dropna(subset=["provider", "track", "score"]).copy()

tracks = sorted(data["track"].unique().tolist())
providers = sorted(data["provider"].unique().tolist())

c1, c2, c3 = st.columns(3)
with c1:
    st.metric("מספר מסלולים בקובץ", len(tracks))
with c2:
    st.metric("מספר גופים", len(providers))
with c3:
    st.metric("מספר שורות", len(data))

st.subheader("בחירת מסלולים ומשקלים")

# Select tracks
selected_tracks = st.multiselect(
    "בחר מסלולים",
    options=tracks,
    default=tracks[:k],
    max_selections=k,
)

if len(selected_tracks) != k:
    st.warning("בחר בדיוק את מספר המסלולים שהוגדר.")
    st.stop()

# Weights
st.caption("משקל לכל מסלול (אחוזים). אם לא מסתכם ל-100, נבצע נרמול אוטומטי.")
weights = {}
cols = st.columns(k)
for i, tr in enumerate(selected_tracks):
    with cols[i]:
        weights[tr] = st.number_input(f"משקל: {tr}", min_value=0.0, max_value=100.0, value=round(100.0 / k, 2), step=1.0)

w_sum = sum(weights.values())
if w_sum <= 0:
    st.error("סכום משקלים חייב להיות גדול מ-0")
    st.stop()

norm_weights = {tr: (w / w_sum) for tr, w in weights.items()}

# Build per-track tables
per_track = {}
for tr in selected_tracks:
    sub = data[data["track"] == tr][["provider", "score"]].copy()
    sub = sub.dropna().sort_values("score", ascending=False)
    per_track[tr] = sub

# Cartesian combinations: choose 1 provider per selected track
combos = list(itertools.product(*[per_track[tr]["provider"].tolist() for tr in selected_tracks]))

rows = []
for combo in combos:
    total = 0.0
    row = {}
    for tr, prov in zip(selected_tracks, combo):
        score = float(per_track[tr].loc[per_track[tr]["provider"] == prov, "score"].iloc[0])
        total += norm_weights[tr] * score
        row[tr] = prov
        row[f"{tr} – ציון"] = score
    row["ציון משוקלל"] = total
    rows.append(row)

result = pd.DataFrame(rows)
result = result.sort_values("ציון משוקלל", ascending=False).reset_index(drop=True)

st.subheader("התוצאות")

top_n = st.slider("כמה תוצאות להציג?", min_value=5, max_value=min(200, len(result)), value=min(30, len(result)))

st.dataframe(result.head(top_n), use_container_width=True)

# Download
csv_bytes = result.to_csv(index=False).encode("utf-8-sig")
st.download_button(
    "הורד CSV של כל השילובים",
    data=csv_bytes,
    file_name="mix_combinations.csv",
    mime="text/csv",
)

st.subheader("Top בכל מסלול")
for tr in selected_tracks:
    st.markdown(f"#### {tr}")
    st.dataframe(per_track[tr].head(15), use_container_width=True)
