import streamlit as st
import pandas as pd
import joblib
import re
import unicodedata
from datetime import datetime
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
from openpyxl.styles import PatternFill

# ===================== ألوان لوجو إجادة =====================
PRIMARY   = "#0B3D91"   # كحلي اللوجو
SECONDARY = "#1565C0"   # أزرق متوسط
ACCENT    = "#42A5F5"   # سماوي
SUCCESS   = "#1E88E5"   # أزرق نجاح
DANGER    = "#90A4AE"   # رمادي مزرق
WARNING   = "#FB8C00"   # برتقالي
BG        = "#F5F9FC"   # خلفية فاتحة
CARD      = "#FFFFFF"
TEXT_DARK = "#0D2137"
TEXT_MUTED = "#5A6A7A"

st.set_page_config(
    page_title="إجادة | لوحة متابعة التحصيل",
    page_icon="logo.png",
    layout="wide",
    initial_sidebar_state="collapsed"
)

_THEME_VARS = (
    "@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Arabic:wght@400;500;600;700&display=swap');"
    ":root{--primary:%s;--secondary:%s;--accent:%s;--bg:%s;--card:%s;--ink:%s;--muted:%s;"
    "--line:#DDE3EC;--tint:#E8EEF8;}"
) % (PRIMARY, SECONDARY, ACCENT, BG, CARD, TEXT_DARK, TEXT_MUTED)

st.markdown("<style>" + _THEME_VARS + """
    html, body, [class*="css"], .stMarkdown, p, label, button, input, textarea, th, td {
        font-family: 'IBM Plex Sans Arabic', 'Cairo', sans-serif;
    }
    .stApp { background: var(--bg); color: var(--ink); }
    .block-container { max-width: 1320px; padding-top: 1.4rem; padding-bottom: 3rem; }
    #MainMenu, footer, header { visibility: hidden; }

    /* ---------- الهيدر ---------- */
    .ejada-header {
        background: var(--primary);
        border-radius: 10px;
        padding: 14px 20px;
        display: flex; align-items: center; gap: 16px;
        margin-bottom: 14px;
    }
    .ejada-header .logo-wrap {
        background: #fff; border-radius: 8px; padding: 6px;
        display: flex; align-items: center; justify-content: center;
        width: 64px; height: 64px; flex-shrink: 0;
    }
    .ejada-header h1 {
        margin: 0; font-size: 1.35rem; font-weight: 700; line-height: 1.4;
        color: #fff !important;
    }
    .ejada-header p { margin: 2px 0 0 0; font-size: .9rem; color: rgba(255,255,255,.78); }

    /* ---------- العناوين ---------- */
    h1, h2, h3, h4, h5 { color: var(--ink) !important; font-weight: 600 !important; letter-spacing: 0; }
    h2, [data-testid="stHeadingWithActionElements"] h2 { font-size: 1.3rem !important; }
    h3 { font-size: 1.1rem !important; }
    h5 { font-size: 1rem !important; }
    [data-testid="stCaptionContainer"] { color: var(--muted); }

    /* ---------- شريط التنقل ---------- */
    div[role="radiogroup"] {
        background: var(--card);
        border: 1px solid var(--line);
        border-radius: 8px;
        padding: 6px; gap: 4px;
    }
    div[role="radiogroup"] label {
        font-weight: 500 !important; color: var(--ink) !important;
        padding: 6px 12px; border-radius: 6px; margin: 0;
        transition: background .15s;
    }
    div[role="radiogroup"] label:hover { background: var(--bg); }
    div[role="radiogroup"] label:has(input:checked) {
        background: var(--tint); color: var(--primary) !important;
    }
    div[role="radiogroup"] label:has(input:checked) p { font-weight: 700; color: var(--primary); }

    /* ---------- المؤشرات ---------- */
    div[data-testid="stMetric"] {
        background: var(--card);
        border: 1px solid var(--line);
        border-radius: 8px;
        padding: 14px 16px;
    }
    div[data-testid="stMetricLabel"] { color: var(--muted) !important; font-weight: 500 !important; font-size: .85rem; }
    div[data-testid="stMetricValue"] {
        color: var(--ink); font-weight: 700; font-variant-numeric: tabular-nums;
    }

    /* ---------- الأزرار ---------- */
    .stButton > button, .stDownloadButton > button {
        border-radius: 8px; font-weight: 600; padding: .5rem 1.1rem;
        border: 1px solid var(--line); background: var(--card); color: var(--ink);
        box-shadow: none; transition: background .15s, border-color .15s;
    }
    .stButton > button:hover { border-color: var(--primary); color: var(--primary); }
    .stButton > button[kind="primary"], button[data-testid="stBaseButton-primary"],
    .stDownloadButton > button, button[data-testid="stBaseButton-downloadButton"] {
        background: var(--primary) !important; color: #fff !important;
        border: 1px solid var(--primary) !important;
    }
    .stButton > button[kind="primary"]:hover, button[data-testid="stBaseButton-primary"]:hover,
    .stDownloadButton > button:hover, button[data-testid="stBaseButton-downloadButton"]:hover {
        background: var(--secondary) !important; border-color: var(--secondary) !important; color: #fff !important;
    }
    button:focus-visible, input:focus-visible, textarea:focus-visible {
        outline: 2px solid var(--accent) !important; outline-offset: 2px;
    }

    /* ---------- رفع الملفات ---------- */
    section[data-testid="stFileUploaderDropzone"], div[data-testid="stFileUploaderDropzone"] {
        background: var(--card); border: 1px dashed #B8C4D6; border-radius: 8px;
    }

    /* ---------- الجداول والتنبيهات والـ expander ---------- */
    div[data-testid="stDataFrame"] {
        border: 1px solid var(--line); border-radius: 8px; overflow: hidden; background: var(--card);
    }
    div[data-testid="stAlert"] { border-radius: 8px; }
    div[data-testid="stExpander"] details {
        background: var(--card); border: 1px solid var(--line) !important; border-radius: 8px;
    }
    button[role="tab"] { font-weight: 600; }
    button[role="tab"][aria-selected="true"] { color: var(--primary); }

    .block-card {
        background: var(--card); border: 1px solid var(--line);
        border-radius: 8px; padding: 16px 18px; margin-bottom: 10px;
    }
    hr { border: none; border-top: 1px solid var(--line); margin: 1.1rem 0; }

    @media (max-width: 640px) {
        .ejada-header { padding: 12px 14px; }
        .ejada-header h1 { font-size: 1.1rem; }
        .ejada-header p { display: none; }
    }
</style>""", unsafe_allow_html=True)

# ---------- ثيم الرسومات (Plotly) ----------
import plotly.io as pio
pio.templates["ejada"] = go.layout.Template(layout=dict(
    font=dict(family="IBM Plex Sans Arabic, Cairo, sans-serif", color=TEXT_DARK, size=13),
    colorway=[PRIMARY, ACCENT, "#7FA6D9", WARNING, DANGER],
    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
    xaxis=dict(gridcolor="#E6EBF2", linecolor="#C9D2DE", zeroline=False, ticks="outside", tickcolor="#C9D2DE"),
    yaxis=dict(gridcolor="#E6EBF2", linecolor="rgba(0,0,0,0)", zeroline=False),
    title=dict(font=dict(size=15, color=TEXT_DARK), x=0.5),
    legend=dict(bgcolor="rgba(0,0,0,0)", borderwidth=0),
    hoverlabel=dict(font_family="IBM Plex Sans Arabic, Cairo, sans-serif"),
    margin=dict(l=40, r=20, t=50, b=40),
))
pio.templates.default = "plotly_white+ejada"


def clean_text(text):
    text = str(text)
    text = re.sub(r'\+?\d+', ' ', text)
    text = re.sub(r'[^\u0600-\u06FF\s]', ' ', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def find_column(df, possible_names, keywords=None):
    for name in possible_names:
        if name in df.columns:
            return name
    if keywords:
        for c in df.columns:
            c_str = str(c).upper()
            if any(kw.upper() in c_str for kw in keywords):
                return c
    return None


def unify_result_columns(frame):
    """توحيد أعمدة النتيجة المترادفة (ناجحه / ناجحة / …) في عمود واحد"""
    frame = frame.copy()
    frame.columns = frame.columns.astype(str).str.strip().str.replace(r"\s+", " ", regex=True)

    def _key(name):
        s = str(name).strip().replace("ة", "ه").replace("ى", "ي")
        s = s.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
        return " ".join(s.split()).lower()

    matched = []
    for c in frame.columns:
        if c == "__مصدر_الملف__":
            continue
        ck = _key(c)
        if ck in ("ناجحه", "ناجح", "نجاح", "success", "successful", "classification", "تصنيف", "نتيجه", "نتيجة") or "ناجح" in ck:
            matched.append(c)
    if len(matched) >= 1:
        series = frame[matched[0]]
        for c in matched[1:]:
            series = series.combine_first(frame[c])
        frame = frame.drop(columns=matched)
        frame["ناجحه"] = series
    return frame


@st.cache_resource
def load_models():
    model = joblib.load("model.pkl")
    vectorizer = joblib.load("vectorizer.pkl")
    return model, vectorizer


# ===================== الهيدر مع اللوجو =====================
# اللوجو مضمّن داخل الكود عشان يظهر دائمًا بدون الاعتماد على ملف خارجي
EJADA_LOGO_B64 = """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"""

st.markdown(f"""
<div class="ejada-header">
    <div class="logo-wrap">
        <img src="data:image/png;base64,{EJADA_LOGO_B64}" width="52" height="52"
             style="object-fit: contain; display: block;" alt="إجادة" />
    </div>
    <div>
        <h1>لوحة متابعة التحصيل</h1>
        <p>شركة إجادة لتحصيل الديون — تقارير أداء المكالمات والمحصلين</p>
    </div>
</div>
""", unsafe_allow_html=True)


def _ops_find_col(df, candidates):
    cols = list(df.columns)
    lower_map = {str(c).strip().lower(): c for c in cols}
    for cand in candidates:
        if cand in cols:
            return cand
        key = str(cand).strip().lower()
        if key in lower_map:
            return lower_map[key]
    for c in cols:
        cl = str(c).strip().lower()
        for cand in candidates:
            if str(cand).strip().lower() in cl or cl in str(cand).strip().lower():
                return c
    return None


def _ops_read_excel(uploaded):
    if str(getattr(uploaded, "name", "")).lower().endswith(".csv"):
        try:
            df = pd.read_csv(uploaded, encoding="utf-8-sig")
        except UnicodeDecodeError:
            uploaded.seek(0)
            df = pd.read_csv(uploaded, encoding="cp1256")
        df = df.dropna(axis=1, how="all")
        df.columns = df.columns.astype(str).str.strip()
        return df, "csv"
    xl = pd.ExcelFile(uploaded)
    best, best_score = xl.sheet_names[0], -1
    for name in xl.sheet_names:
        try:
            tmp = pd.read_excel(xl, sheet_name=name, nrows=5)
            score = len([c for c in tmp.columns if not str(c).startswith("Unnamed")])
            if score > best_score:
                best, best_score = name, score
        except Exception:
            continue
    df = pd.read_excel(xl, sheet_name=best)
    df = df.dropna(axis=1, how="all")
    df.columns = df.columns.astype(str).str.strip()
    if len(df) > 0:
        first = df.iloc[0]
        emptyish = first.isna().all() or all(str(v).strip() in {"", "nan", "None"} for v in first.values)
        if emptyish:
            df = df.iloc[1:].reset_index(drop=True)
    return df, best


def _ops_parse_date(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    ts = pd.to_datetime(v, errors="coerce", dayfirst=True)
    if pd.isna(ts):
        return None
    return ts.date()


def _ops_to_excel_download(df_dict, filename):
    import io as _io
    if not any(d is not None and len(d) for d in df_dict.values()):
        st.info("مفيش بيانات للتحميل")
        return
    buf = _io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sheet, d in df_dict.items():
            if d is not None and len(d):
                d.to_excel(writer, index=False, sheet_name=str(sheet)[:31])
    buf.seek(0)
    st.download_button(
        "تحميل النتيجة Excel",
        data=buf.getvalue(),
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


# ===================== اختيار المرحلة =====================
mode = st.radio(
    "اختر الصفحة",
    [
        "① معالجة الشيت",
        "② لوحة التحليل",
        "③ دمج الشيتات + Pivot",
        "④ الوعود",
        "⑤ الإهمال",
        "⑥ الجدولة المتعثرة",
        "⑦ أخطاء الحالات",
        "⑧ التوزيع",
    ],
    horizontal=True,
    label_visibility="collapsed",
    key="main_nav",
)

st.markdown("---")

# ============================================================
# المرحلة 1: التصنيف
# ============================================================
if mode.startswith("①"):
    st.subheader("المرحلة 1 — معالجة الشيت")
    st.caption("ارفع الشيت الخام → التصنيف + الوقت + مدة المكالمة + تمييز الدبليكات → حمّل النتيجة وراجعها")

    try:
        model, vectorizer = load_models()
    except Exception as e:
        st.error(f"تعذر تحميل الموديل: {e}")
        st.info("تأكد من وجود: model.pkl و vectorizer.pkl في نفس المجلد")
        st.stop()

    uploaded_file = st.file_uploader(
        "ارفع ملف Excel الخام",
        type=["xlsx", "xls"],
        key="classify_upload"
    )

    if uploaded_file is not None:
        df_raw = pd.read_excel(uploaded_file)
        df_raw.columns = df_raw.columns.str.strip()
        original_count = len(df_raw)
        st.success(f"تم استلام الملف — {original_count:,} صف")

        st.markdown("##### تأكيد أسماء الأعمدة")
        all_cols = ["— لا يوجد —"] + list(df_raw.columns)

        def get_default_index(candidates, keywords=None):
            found = find_column(df_raw, candidates, keywords)
            if found and found in all_cols:
                return all_cols.index(found)
            return 0

        c1, c2 = st.columns(2)
        with c1:
            notes_col = st.selectbox("عمود الإفادة / Notes *مطلوب*", all_cols,
                index=get_default_index(["Notes", "NOTES", "ملاحظات", "الملاحظات", "الإفادة"]))
            created_col = st.selectbox("عمود التاريخ (CREATED ON)", all_cols,
                index=get_default_index(["CREATED ON", "Created On", "Created"], ["CREATED", "تاريخ"]))
            collector_col = st.selectbox("عمود اسم المحصل", all_cols,
                index=get_default_index(["اسم المحصل", "المحصل", "Collector", "Agent"], ["محصل", "collector", "agent"]))
        with c2:
            account_col = st.selectbox("عمود رقم الأكونت / Account", all_cols,
                index=get_default_index(["رقم الاكونت", "Account", "ACCOUNT"], ["ACCOUNT", "اكونت", "حساب"]))
            claims_col = st.selectbox("عمود CLAIMS", all_cols,
                index=get_default_index(["CLAIMS", "Claims", "Claim"], ["CLAIM"]))
            debitor_col = st.selectbox("عمود DEBITOR / المدين", all_cols,
                index=get_default_index(["DEBITOR", "Debitor", "Debtor", "المدين"], ["DEBITOR", "DEBTOR", "مدين"]))

        def clean_sel(v):
            return None if v == "— لا يوجد —" else v

        notes_col = clean_sel(notes_col)
        created_col = clean_sel(created_col)
        collector_col = clean_sel(collector_col)
        account_col = clean_sel(account_col)
        claims_col = clean_sel(claims_col)
        debitor_col = clean_sel(debitor_col)

        if not notes_col:
            st.error("يجب اختيار عمود الإفادة (Notes)")
            st.stop()

        if st.button("تشغيل التصنيف", type="primary", use_container_width=True):
            df = df_raw.copy()

            with st.spinner("جاري التصنيف..."):
                df[notes_col] = df[notes_col].apply(clean_text)
                X = vectorizer.transform(df[notes_col])
                df["ناجحه"] = model.predict(X)
                empty_mask = df[notes_col].isna() | (df[notes_col].astype(str).str.strip() == "")
                df.loc[empty_mask, "ناجحه"] = 0

            # تمييز الدبليكات بدون حذف
            dup_cols = [c for c in [collector_col, account_col, claims_col, notes_col, debitor_col] if c]
            dup_count = 0
            df["دبليكات"] = ""
            if len(dup_cols) >= 2:
                is_dup = df.duplicated(subset=dup_cols, keep=False)
                df.loc[is_dup, "دبليكات"] = "نعم"
                dup_count = int(is_dup.sum())

            # الوقت + مدة المكالمة
            if created_col:
                df[created_col] = pd.to_datetime(df[created_col], errors="coerce")
                # ناتج Excel بالظبط: MOD(Created_On, 1) = كسر من اليوم (0→1)
                def _excel_mod_time(ts):
                    if pd.isna(ts):
                        return None
                    # نفس حساب MOD في Excel
                    return (
                        ts.hour * 3600
                        + ts.minute * 60
                        + ts.second
                        + ts.microsecond / 1_000_000
                    ) / 86400.0

                df["الوقت"] = df[created_col].apply(_excel_mod_time)
                if collector_col:
                    df = df.sort_values(by=[collector_col, created_col], ascending=[True, True]).reset_index(drop=True)
                    df["مدة المكالمة"] = np.nan
                    for i in range(1, len(df)):
                        same = (
                            pd.notna(df.loc[i, collector_col]) and pd.notna(df.loc[i - 1, collector_col])
                            and str(df.loc[i, collector_col]).strip() == str(df.loc[i - 1, collector_col]).strip()
                        )
                        has_t = pd.notna(df.loc[i, created_col]) and pd.notna(df.loc[i - 1, created_col])
                        if same and has_t:
                            delta = (df.loc[i, created_col] - df.loc[i - 1, created_col]).total_seconds() / 60
                            if 0 < delta <= 180:
                                df.loc[i, "مدة المكالمة"] = round(delta, 2)

            today = datetime.today().strftime("%Y-%m-%d")
            output_name = f"EJADA_classified_{today}.xlsx"
            yellow_fill = PatternFill(start_color="FFF59D", end_color="FFF59D", fill_type="solid")

            with pd.ExcelWriter(output_name, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="النتائج")
                ws = writer.sheets["النتائج"]
                # فورمات العرض زي Excel: h:mm:ss AM/PM على قيمة MOD الرقمية
                time_col_idx = None
                dup_col_idx = None
                for idx, cell in enumerate(ws[1], start=1):
                    if cell.value == "الوقت":
                        time_col_idx = idx
                    if cell.value == "دبليكات":
                        dup_col_idx = idx
                if time_col_idx is not None:
                    for row in range(2, ws.max_row + 1):
                        cell = ws.cell(row=row, column=time_col_idx)
                        if cell.value is not None and cell.value != "":
                            try:
                                cell.value = float(cell.value)
                            except (TypeError, ValueError):
                                pass
                            cell.number_format = "h:mm:ss AM/PM"
                if dup_col_idx is not None:
                    for row in range(2, ws.max_row + 1):
                        if ws.cell(row=row, column=dup_col_idx).value == "نعم":
                            for col in range(1, ws.max_column + 1):
                                ws.cell(row=row, column=col).fill = yellow_fill

            st.success(
                f"اكتمل التصنيف — **{len(df):,}** إفادة "
                f"| دبليكات مميزة بالأصفر: **{dup_count}** (لم يُحذف شيء)"
            )
            with open(output_name, "rb") as f:
                st.download_button(
                    "تحميل الشيت المصنف",
                    data=f,
                    file_name=output_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            st.info("بعد ما تراجع الشيت وتمسح الدبليكات اللي عايزها، ارجع للمرحلة ② وارفع الشيت النهائي للتحليل.")

# ============================================================
# المرحلة 2: التحليل فقط
# ============================================================
elif mode.startswith("②"):
    st.subheader("المرحلة 2 — لوحة التحليل")
    st.caption("ارفع الشيت بعد ما تراجعُه وتظبطه — هتطلع لك لوحة تحليل جاهزة للسكرين شوت")

    uploaded_files = st.file_uploader(
        "ارفع شيت واحد أو أكثر (بعد المراجعة) — هيتدمجو مع بعض للتحليل",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="analysis_upload"
    )

    if not uploaded_files:
        st.info("ارفع شيت واحد أو أكتر عشان نبني لوحة التحليل على الإجمالي")
        st.stop()

    def _score_sheet(xl, name):
        try:
            tmp = pd.read_excel(xl, sheet_name=name, nrows=30)
            tmp = tmp.dropna(axis=1, how="all")
            cols = [str(c).strip() for c in tmp.columns]
            score = len(cols)
            keywords = ["ناجحه", "ناجح", "notes", "ملاحظات", "محصل", "created", "مدة",
                        "account", "اكونت", "claims", "debitor", "مدين", "دبليك", "الوقت"]
            joined = " ".join(cols).lower()
            for kw in keywords:
                if kw in joined:
                    score += 15
            if len(cols) <= 2:
                score -= 20
            if "count of" in joined or "sum of" in joined or "تقرير" in joined:
                score -= 25
            return score
        except Exception:
            return -999

    def _read_best_table(file_obj, filename):
        """يقرأ أفضل شيت داتا من الملف مع اكتشاف الهيدر"""
        xl = pd.ExcelFile(file_obj)
        scored = sorted(((_score_sheet(xl, n), n) for n in xl.sheet_names), reverse=True)
        chosen_sheet = scored[0][1]
        df_local = None
        header_used = 0
        for hdr in [0, 1, 2, 3, 4, 5]:
            trial = pd.read_excel(xl, sheet_name=chosen_sheet, header=hdr)
            trial = trial.dropna(axis=1, how="all")
            trial.columns = trial.columns.astype(str).str.strip().str.replace(r"\s+", " ", regex=True)
            unnamed = [c for c in trial.columns if str(c).startswith("Unnamed")]
            trial2 = trial.drop(columns=unnamed, errors="ignore") if unnamed else trial
            cols = list(trial2.columns)
            if len(cols) < 2:
                continue
            joined = " ".join(cols).lower()
            if len(cols) == 1 and ("تقرير" in cols[0] or "count of" in joined):
                continue
            hits = sum(
                1 for c in cols
                for kw in ["ناجحه", "ناجح", "notes", "محصل", "created", "مدة", "account", "اكونت", "claims", "مدين", "دبليك"]
                if kw in str(c).lower()
            )
            if hits >= 1 or (len(cols) >= 4 and hdr == 0):
                df_local = trial2
                header_used = hdr
                break
            if df_local is None or len(cols) > len(df_local.columns):
                df_local = trial2
                header_used = hdr
        if df_local is None:
            df_local = pd.read_excel(xl, sheet_name=chosen_sheet)
            df_local = df_local.dropna(axis=1, how="all")
            df_local.columns = df_local.columns.astype(str).str.strip()
        df_local["__مصدر_الملف__"] = filename
        return df_local, chosen_sheet, header_used

    frames = []
    read_log = []
    for f in uploaded_files:
        try:
            part, sheet_name, hdr = _read_best_table(f, f.name)
            frames.append(part)
            read_log.append({
                "الملف": f.name,
                "الشيت": sheet_name,
                "صف الهيدر": hdr + 1,
                "الصفوف": len(part),
                "الأعمدة": len([c for c in part.columns if c != "__مصدر_الملف__"])
            })
        except Exception as e:
            st.error(f"تعذر قراءة الملف {f.name}: {e}")

    if not frames:
        st.error("لم يتم قراءة أي ملف بنجاح")
        st.stop()

    df = pd.concat(frames, ignore_index=True, sort=False)
    # توحيد أسماء الأعمدة + دمج أعمدة ناجحه/ناجحة في عمود واحد
    df.columns = df.columns.astype(str).str.strip().str.replace(r"\s+", " ", regex=True)
    df = unify_result_columns(df)

    st.success(
        f"تم دمج **{len(uploaded_files)}** ملف — الإجمالي **{len(df):,}** صف | **{len(df.columns)}** عمود"
    )
    st.dataframe(pd.DataFrame(read_log), use_container_width=True)

    with st.expander("عرض الأعمدة المكتشفة + عينة من البيانات المدموجة", expanded=True):
        st.write([c for c in df.columns if c != "__مصدر_الملف__"])
        st.dataframe(df.head(8), use_container_width=True)

    cols_join = " ".join(map(str, df.columns)).lower()
    if len([c for c in df.columns if c != "__مصدر_الملف__"]) <= 2 or "count of" in cols_join:
        st.warning(
            "يبدو إن بعض الملفات Pivot / تقرير ملخص مش جدول البيانات. "
            "تأكد إن الملفات فيها أعمدة زي: Notes / ناجحه / اسم المحصل / CREATED ON."
        )

    if len(df.columns) == 0:
        st.error("الملف مفيهوش أعمدة مقروءة. تأكد إن الصف الأول فيه أسماء الأعمدة.")
        st.stop()

    st.markdown("##### تأكيد الأعمدة المستخدمة في التحليل")
    st.caption("اختَر العمود الصحيح من القوائم تحت")

    all_cols = ["— لا يوجد —"] + [c for c in df.columns if c != "__مصدر_الملف__"]

    def _default(cands, kws=None):
        found = find_column(df, cands, kws)
        return all_cols.index(found) if found and found in all_cols else 0

    a1, a2 = st.columns(2)
    with a1:
        success_col = st.selectbox(
            "عمود نتيجة التصنيف (ناجحه) *مطلوب*",
            all_cols,
            index=_default(["ناجحه", "ناجحة", "ناجح", "success", "Success", "result"], ["ناجح", "success"])
        )
        collector_col = st.selectbox(
            "عمود اسم المحصل",
            all_cols,
            index=_default(["اسم المحصل", "المحصل", "Collector", "Agent"], ["محصل", "collector", "agent"])
        )
        duration_col = st.selectbox(
            "عمود مدة المكالمة",
            all_cols,
            index=_default(["مدة المكالمة", "المدة", "Duration"], ["مدة", "duration"])
        )
        account_col = st.selectbox(
            "عمود رقم الحساب / Account",
            all_cols,
            index=_default(
                ["Customer Account Number", "رقم الاكونت", "Account", "ACCOUNT", "رقم الحساب"],
                ["ACCOUNT", "اكونت", "حساب", "Customer"]
            )
        )
    with a2:
        created_col = st.selectbox(
            "عمود التاريخ (CREATED ON)",
            all_cols,
            index=_default(["CREATED ON", "Created On", "Created"], ["CREATED", "تاريخ"])
        )
        dup_col = st.selectbox(
            "عمود الدبليكات",
            all_cols,
            index=_default(["دبليكات", "دبليك"], ["دبليك"])
        )
        time_col = st.selectbox(
            "عمود الوقت",
            all_cols,
            index=_default(["الوقت", "وقت"], ["وقت"])
        )
        assign_col = st.selectbox(
            "عمود تاريخ الإسناد (Claim/Assign Date)",
            all_cols,
            index=_default(
                ["Claim/Assign Date", "Claim Assign Date", "تاريخ الإسناد", "Assign Date"],
                ["ASSIGN", "إسناد", "اسناد"]
            )
        )
        incident_col = st.selectbox(
            "عمود تاريخ الحادث (Claim/Debit Date)",
            all_cols,
            index=_default(
                ["Claim/Debit Date", "Claim Debit Date", "تاريخ الحادث", "Debit Date", "Due Date"],
                ["DEBIT", "حادث", "DEBT"]
            )
        )

    def _clean(v):
        return None if v == "— لا يوجد —" else v

    success_col = _clean(success_col)
    collector_col = _clean(collector_col)
    duration_col = _clean(duration_col)
    account_col = _clean(account_col)
    created_col = _clean(created_col)
    dup_col = _clean(dup_col)
    time_col = _clean(time_col)
    assign_col = _clean(assign_col)
    incident_col = _clean(incident_col)

    if success_col is None:
        st.error("لازم تختار عمود نتيجة التصنيف (ناجحه أو ما يعادله)")
        st.stop()

    st.markdown("##### بيانات ظهور التقرير")
    rmeta1, rmeta2 = st.columns(2)
    with rmeta1:
        report_company = st.text_input(
            "اسم الشركة / المحفظة",
            value="سايكو",
            placeholder="مثال: سايكو"
        ).strip()
    with rmeta2:
        report_period = st.text_input(
            "اسم الفترة",
            value="الفترة الأولى",
            placeholder="مثال: الفترة الأولى"
        ).strip()
    rmeta3, rmeta4 = st.columns(2)
    with rmeta3:
        period_from = st.text_input(
            "من الساعة",
            value="1",
            placeholder="مثال: 1"
        ).strip()
    with rmeta4:
        period_to = st.text_input(
            "إلى الساعة",
            value="4",
            placeholder="مثال: 4"
        ).strip()

    if not st.button("عرض لوحة التحليل", type="primary", use_container_width=True, key="run_analysis"):
        st.stop()

    # تطبيع عمود النجاح (يقبل 0/1 أو نعم/لا أو true/false)
    raw_s = df[success_col]
    if raw_s.dtype == object:
        mapped = raw_s.astype(str).str.strip().str.lower().map({
            "1": 1, "0": 0, "نعم": 1, "لا": 0, "true": 1, "false": 0,
            "ناجحة": 1, "غير ناجحة": 0, "yes": 1, "no": 0
        })
        df[success_col] = pd.to_numeric(mapped, errors="coerce").fillna(
            pd.to_numeric(raw_s, errors="coerce")
        ).fillna(0).astype(int)
    else:
        df[success_col] = pd.to_numeric(raw_s, errors="coerce").fillna(0).astype(int)

    total = len(df)
    success_count = int((df[success_col] == 1).sum())
    fail_count = int((df[success_col] == 0).sum())
    success_rate = (success_count / total * 100) if total else 0
    has_duration = duration_col is not None and df[duration_col].notna().any()
    duration = pd.to_numeric(df[duration_col], errors="coerce").dropna() if has_duration else pd.Series(dtype=float)
    # إجمالي الحسابات = عدد الحسابات الفريدة
    if account_col and account_col in df.columns:
        total_accounts = int(df[account_col].nunique())
    else:
        total_accounts = None
    remaining_dups = 0
    if dup_col and dup_col in df.columns:
        remaining_dups = int((df[dup_col].astype(str).str.strip() == "نعم").sum())

    # ========== عنوان التقرير ==========
    report_date = datetime.today().strftime("%Y-%m-%d")
    # عنوان ديناميكي: نشاط {الفترة} لمحصلين {الشركة}
    company_part = report_company if report_company else "الشركة"
    period_part = report_period if report_period else "الفترة"
    if period_from and period_to:
        period_range = f"من الساعة {period_from} إلى الساعة {period_to}"
        report_title = f"نشاط {period_part} ({period_range}) لمحصلين {company_part}"
        period_display = f"{period_part} — {period_range}"
    else:
        report_title = f"نشاط {period_part} لمحصلين {company_part}"
        period_display = period_part

    st.markdown(f"""
    <div style="background:linear-gradient(135deg,{PRIMARY},{SECONDARY}); color:white;
         padding:14px 20px; border-radius:14px; margin-bottom:16px;">
        <div style="font-size:1.25rem; font-weight:700;">{report_title}</div>
        <div style="opacity:0.9; font-size:0.95rem;">شركة إجادة لتحصيل الديون | {report_date} | إجمالي السجلات: {total:,}</div>
    </div>
    """, unsafe_allow_html=True)

    # ========== KPIs بنفس أسماء اللوحة ==========
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("اجمالي المكالمات", f"{total:,}")
    k2.metric("اجمالي الحسابات", f"{total_accounts:,}" if total_accounts is not None else "—")
    k3.metric("نسبة نجاح المكالمات", f"{success_rate:.2f}%")
    k4.metric(
        "متوسط مدة المكالمة",
        f"{duration.mean():.2f}" if has_duration and len(duration) else "—"
    )

    # ========== تقرير HTML للموبايل / الكمبيوتر (بدون اعتماد على Plotly CDN) ==========
    COLOR_OK = "#1E88E5"
    COLOR_BAD = "#90A4AE"
    COLOR_MID = "#42A5F5"

    def _css_bars(items, value_key, label_key, color=PRIMARY, max_v=None):
        """رسم أعمدة CSS بسيطة تشتغل على الآيفون بدون JavaScript"""
        if not items:
            return "<p style='color:#5A6A7A'>لا توجد بيانات</p>"
        vals = [float(x[value_key]) for x in items]
        mx = max(vals) if max_v is None else max_v
        mx = mx if mx > 0 else 1
        rows = []
        for it in items:
            v = float(it[value_key])
            pct = max(2, round(v / mx * 100, 1))
            rows.append(
                f"<div class='bar-row'>"
                f"<div class='bar-label'>{it[label_key]}</div>"
                f"<div class='bar-track'><div class='bar-fill' style='width:{pct}%;background:{color}'></div></div>"
                f"<div class='bar-val'>{int(v) if float(v).is_integer() else round(v,1)}</div>"
                f"</div>"
            )
        return "".join(rows)

    def _donut_legend(parts):
        # parts: list of (label, value, color)
        total_v = sum(p[1] for p in parts) or 1
        lis = []
        for label, val, color in parts:
            pct = round(val / total_v * 100, 1)
            lis.append(
                f"<div class='legend-item'>"
                f"<span class='dot' style='background:{color}'></span>"
                f"<span>{label}</span>"
                f"<strong>{val:,}</strong>"
                f"<span class='pct'>({pct}%)</span>"
                f"</div>"
            )
        return "".join(lis)

    # توزيع النجاح
    success_legend = _donut_legend([
        ("ناجحة", success_count, COLOR_OK),
        ("غير ناجحة", fail_count, COLOR_BAD),
    ])

    # متوسط مدة المكالمة حسب النتيجة
    duration_bars = ""
    duration_bins_legend = ""
    if has_duration and len(duration):
        avg_map = df.groupby(success_col)[duration_col].mean().to_dict()
        dur_items = []
        if 0 in avg_map:
            dur_items.append({"label": "غير ناجحة", "value": round(float(avg_map[0]), 1)})
        if 1 in avg_map:
            dur_items.append({"label": "ناجحة", "value": round(float(avg_map[1]), 1)})
        duration_bars = _css_bars(dur_items, "value", "label", COLOR_OK)

        short = int((duration <= 5).sum())
        mid = int(((duration > 5) & (duration <= 15)).sum())
        long_ = int((duration > 15).sum())
        duration_bins_legend = _donut_legend([
            ("≤ 5 د", short, COLOR_OK),
            ("5–15 د", mid, COLOR_MID),
            ("> 15 د", long_, COLOR_BAD),
        ])

    # المحصلين
    collector_bars = ""
    collector_success_bars = ""
    collector_table_html = ""
    if collector_col and collector_col in df.columns:
        stats_r = (
            df.groupby(collector_col)
            .agg(اجمالي=(success_col, "count"), ناجحة=(success_col, "sum"))
            .reset_index()
        )
        stats_r["نسبة%"] = (stats_r["ناجحة"] / stats_r["اجمالي"] * 100).round(1)
        stats_r = stats_r.sort_values("اجمالي", ascending=False)
        items_all = [{"label": str(r[collector_col]), "value": int(r["اجمالي"])} for _, r in stats_r.iterrows()]
        items_ok = [{"label": str(r[collector_col]), "value": int(r["ناجحة"])} for _, r in stats_r.iterrows()]
        collector_bars = _css_bars(items_all, "value", "label", PRIMARY)
        collector_success_bars = _css_bars(items_ok, "value", "label", COLOR_OK)
        rows = ""
        for _, r in stats_r.iterrows():
            rows += (
                f"<tr><td>{r[collector_col]}</td><td>{int(r['اجمالي'])}</td>"
                f"<td>{int(r['ناجحة'])}</td><td>{r['نسبة%']}%</td></tr>"
            )
        collector_table_html = f"""
        <h3 class="sec-title">أداء المحصلين</h3>
        <div class="table-wrap">
          <table>
            <thead><tr><th>المحصل</th><th>اجمالي المكالمات</th><th>الناجحة</th><th>نسبة النجاح</th></tr></thead>
            <tbody>{rows}</tbody>
          </table>
        </div>
        """

    # سنة الإسناد
    year_bars = ""
    if assign_col and assign_col in df.columns:
        tmp_a = df.copy()
        tmp_a[assign_col] = pd.to_datetime(tmp_a[assign_col], errors="coerce")
        tmp_a = tmp_a.dropna(subset=[assign_col])
        tmp_a = tmp_a[tmp_a[success_col] == 1]
        if len(tmp_a):
            tmp_a["سنة"] = tmp_a[assign_col].dt.year.astype(int)
            by_year = tmp_a.groupby("سنة").size().reset_index(name="عدد").sort_values("سنة")
            year_items = [{"label": str(int(r["سنة"])), "value": int(r["عدد"])} for _, r in by_year.iterrows()]
            year_bars = _css_bars(year_items, "value", "label", PRIMARY)

    # سنة الحادث
    incident_bars = ""
    if incident_col and incident_col in df.columns:
        tmp_i = df.copy()
        tmp_i[incident_col] = pd.to_datetime(tmp_i[incident_col], errors="coerce")
        tmp_i = tmp_i.dropna(subset=[incident_col])
        tmp_i = tmp_i[tmp_i[success_col] == 1]
        if len(tmp_i):
            tmp_i["سنة_الحادث"] = tmp_i[incident_col].dt.year.astype(int)
            by_inc = tmp_i.groupby("سنة_الحادث").size().reset_index(name="عدد").sort_values("سنة_الحادث")
            inc_items = [{"label": str(int(r["سنة_الحادث"])), "value": int(r["عدد"])} for _, r in by_inc.iterrows()]
            incident_bars = _css_bars(inc_items, "value", "label", COLOR_OK)

    avg_dur_txt = f"{duration.mean():.2f}" if has_duration and len(duration) else "—"
    accounts_txt = f"{total_accounts:,}" if total_accounts is not None else "—"
    logo_img = (
        f'<img class="logo" src="data:image/png;base64,{EJADA_LOGO_B64}" '
        f'alt="إجادة" />'
    )

    year_section = ""
    if year_bars:
        year_section = f"""
        <div class="card-block">
          <h3 class="sec-title">عدد المكالمات الناجحة من سنة الإسناد</h3>
          {year_bars}
        </div>
        """
    incident_section = ""
    if incident_bars:
        incident_section = f"""
        <div class="card-block">
          <h3 class="sec-title">عدد المكالمات الناجحة من سنين الحادث</h3>
          {incident_bars}
        </div>
        """

    duration_section = ""
    if duration_bars or duration_bins_legend:
        duration_section = f"""
        <div class="card-block">
          <h3 class="sec-title">متوسط مدة المكالمة حسب النتيجة</h3>
          {duration_bars}
        </div>
        <div class="card-block">
          <h3 class="sec-title">توزيع مدد المكالمات</h3>
          <div class="legend">{duration_bins_legend}</div>
        </div>
        """

    collector_section = ""
    if collector_bars:
        collector_section = f"""
        <div class="card-block">
          <h3 class="sec-title">اجمالي المكالمات لكل محصل</h3>
          {collector_bars}
        </div>
        <div class="card-block">
          <h3 class="sec-title">عدد المكالمات الناجحة لكل محصل</h3>
          {collector_success_bars}
        </div>
        """

    html_report = f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover" />
<meta name="apple-mobile-web-app-capable" content="yes" />
<title>{report_title} — إجادة</title>
<style>
  * {{ box-sizing: border-box; -webkit-tap-highlight-color: transparent; }}
  html, body {{
    margin: 0; padding: 0; background: #F5F9FC; color: #0D2137;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Tahoma, Arial, sans-serif;
    -webkit-text-size-adjust: 100%;
  }}
  .page {{ width: 100%; max-width: 980px; margin: 0 auto; padding: 12px; }}
  .header {{
    background: linear-gradient(135deg, #0B3D91, #1565C0);
    color: #fff; border-radius: 14px; padding: 14px;
    display: flex; align-items: center; gap: 12px; margin-bottom: 12px;
  }}
  .logo {{
    width: 56px; height: 56px; object-fit: contain;
    background: #fff; border-radius: 10px; padding: 6px; flex-shrink: 0;
  }}
  .header h1 {{ margin: 0; font-size: 1.1rem; }}
  .header p {{ margin: 4px 0 0; opacity: .92; font-size: .82rem; }}
  .cards {{
    display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin-bottom: 12px;
  }}
  .card {{
    background: #fff; border-radius: 12px; padding: 12px;
    box-shadow: 0 1px 4px rgba(11,61,145,.08); border-top: 3px solid #1565C0;
  }}
  .card .label {{ color: #5A6A7A; font-size: .75rem; margin-bottom: 4px; }}
  .card .value {{ font-size: 1.15rem; font-weight: 700; color: #0B3D91; }}
  .card-block {{
    background: #fff; border-radius: 12px; padding: 12px;
    box-shadow: 0 1px 4px rgba(11,61,145,.08); margin-bottom: 10px;
  }}
  .sec-title {{ margin: 0 0 10px; color: #0B3D91; font-size: 1rem; }}
  .bar-row {{
    display: grid; grid-template-columns: 34% 1fr 48px;
    gap: 6px; align-items: center; margin-bottom: 8px;
  }}
  .bar-label {{ font-size: .78rem; color: #0D2137; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
  .bar-track {{ background: #E8EEF5; border-radius: 999px; height: 12px; overflow: hidden; }}
  .bar-fill {{ height: 100%; border-radius: 999px; }}
  .bar-val {{ font-size: .78rem; font-weight: 700; color: #0B3D91; text-align: left; }}
  .legend {{ display: flex; flex-direction: column; gap: 8px; }}
  .legend-item {{
    display: grid; grid-template-columns: 12px 1fr auto auto;
    gap: 8px; align-items: center; font-size: .88rem;
  }}
  .dot {{ width: 12px; height: 12px; border-radius: 50%; display: inline-block; }}
  .pct {{ color: #5A6A7A; }}
  .table-wrap {{ width: 100%; overflow-x: auto; -webkit-overflow-scrolling: touch; }}
  table {{ width: 100%; min-width: 420px; border-collapse: collapse; font-size: .86rem; background: #fff; }}
  th, td {{ padding: 10px 8px; border-bottom: 1px solid #E8EEF5; text-align: right; }}
  th {{ background: #0B3D91; color: #fff; }}
  .footer {{ text-align: center; color: #5A6A7A; font-size: .75rem; margin: 16px 0; }}
  .hint {{
    background: #E3F2FD; color: #0D2137; border-radius: 10px;
    padding: 10px 12px; font-size: .8rem; margin-bottom: 12px; line-height: 1.5;
  }}
  @media (min-width: 700px) {{
    .page {{ padding: 20px; }}
    .cards {{ grid-template-columns: repeat(4, 1fr); gap: 12px; }}
    .card .value {{ font-size: 1.35rem; }}
    .header h1 {{ font-size: 1.35rem; }}
    .grid-2 {{ display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }}
  }}
  @media (min-width: 1024px) {{
    .page {{ padding: 24px; }}
    .card .value {{ font-size: 1.45rem; }}
  }}
</style>
</head>
<body>
  <div class="page">
    <div class="header">
      {logo_img}
      <div>
        <h1>{report_title}</h1>
        <p>شركة إجادة لتحصيل الديون | {report_date}</p>
        <p style="margin:2px 0 0;opacity:.9;font-size:.8rem;">{company_part} — {period_display}</p>
      </div>
    </div>

    <div class="hint">
      على الآيفون: إذا وصل الملف على واتساب اضغط عليه → شارك → حفظ في Files ثم افتحه من تطبيق Files بـ Safari.
    </div>

    <div class="cards">
      <div class="card"><div class="label">اجمالي المكالمات</div><div class="value">{total:,}</div></div>
      <div class="card"><div class="label">اجمالي الحسابات</div><div class="value">{accounts_txt}</div></div>
      <div class="card"><div class="label">نسبة نجاح المكالمات</div><div class="value">{success_rate:.2f}%</div></div>
      <div class="card"><div class="label">متوسط مدة المكالمة</div><div class="value">{avg_dur_txt}</div></div>
    </div>

    <div class="grid-2">
      <div class="card-block">
        <h3 class="sec-title">تصنيف المكالمات</h3>
        <div class="legend">{success_legend}</div>
      </div>
      {duration_section}
    </div>

    {collector_section}
    {collector_table_html}
    {year_section}
    {incident_section}

    <div class="footer">شركة إجادة لتحصيل الديون — يعمل بدون إنترنت على الموبايل والكمبيوتر</div>
  </div>
</body>
</html>
"""

    safe_co = "".join(ch for ch in company_part if ch.isalnum() or ch in (" ", "-", "_")).strip().replace(" ", "_") or "report"
    safe_pe = "".join(ch for ch in period_part if ch.isalnum() or ch in (" ", "-", "_")).strip().replace(" ", "_") or "period"
    html_name = f"EJADA_{safe_co}_{safe_pe}_{report_date}.html"
    st.download_button(
        label="تحميل تقرير HTML (موبايل + كمبيوتر)",
        data=html_report.encode("utf-8"),
        file_name=html_name,
        mime="text/html",
        use_container_width=True,
    )
    st.info(
        "على الآيفون: احفظ الملف في Files ثم افتحه من Safari. "
        "واتساب أحيانًا لا يعرض HTML مباشرة."
    )

    # ========== صف 1: توزيع + ملخص + مدة ==========
    # ألوان متناسقة مع الأزرق (هوية إجادة)
    COLOR_OK = "#1E88E5"    # أزرق نجاح واضح
    COLOR_BAD = "#90A4AE"   # رمادي مزرق للفشل (مشقوش مع الأزرق)
    COLOR_MID = "#42A5F5"   # سماوي للفئات المتوسطة

    r1c1, r1c2, r1c3 = st.columns([1.15, 1, 1.15])

    with r1c1:
        st.markdown("##### تصنيف المكالمات")
        fig_pie = go.Figure(data=[go.Pie(
            labels=["ناجحة", "غير ناجحة"],
            values=[success_count, fail_count],
            hole=0.55,
            marker=dict(colors=[COLOR_OK, COLOR_BAD], line=dict(color="white", width=3)),
            textinfo="label+percent+value",
            texttemplate="%{label}<br>%{value:,}<br>(%{percent})",
            textfont=dict(size=13, color="#0D2137"),
            hovertemplate="%{label}: %{value:,} (%{percent})<extra></extra>"
        )])
        fig_pie.update_layout(
            height=310, margin=dict(t=10, b=10, l=10, r=10), showlegend=False,
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            annotations=[dict(
                text=f"<b>{success_rate:.1f}%</b><br>نجاح",
                x=0.5, y=0.5, font_size=16, showarrow=False, font_color=COLOR_OK
            )]
        )
        st.plotly_chart(fig_pie, use_container_width=True)

    with r1c2:
        st.markdown("##### ملخص تنفيذي")
        ratio_txt = "مرتفع" if success_rate >= 70 else ("متوسط" if success_rate >= 50 else "منخفض")
        ratio_color = COLOR_OK if success_rate >= 70 else (WARNING if success_rate >= 50 else COLOR_BAD)
        st.markdown(f"""
        <div class="block-card" style="line-height:1.95; font-size:0.95rem;">
            • مستوى النجاح: <b style="color:{ratio_color}">{ratio_txt}</b> ({success_rate:.1f}%)<br>
            • مكالمات ناجحة: <b style="color:{COLOR_OK}">{success_count:,}</b><br>
            • مكالمات غير ناجحة: <b style="color:{COLOR_BAD}">{fail_count:,}</b><br>
            • اجمالي المكالمات: <b>{total:,}</b><br>
            {"• متوسط مدة المكالمة: <b>" + f"{duration.mean():.1f}" + " د</b><br>" if has_duration and len(duration) else ""}
            {"• دبليكات متبقية: <b style='color:" + COLOR_BAD + "'>" + str(remaining_dups) + "</b>" if remaining_dups else "• لا توجد دبليكات معلّمة"}
        </div>
        """, unsafe_allow_html=True)

    with r1c3:
        st.markdown("##### متوسط مدة المكالمة حسب النتيجة")
        if has_duration and len(duration):
            avg_dur = df.groupby(success_col)[duration_col].mean().reset_index()
            avg_dur["النتيجة"] = avg_dur[success_col].map({1: "ناجحة", 0: "غير ناجحة"})
            avg_dur["label"] = avg_dur[duration_col].map(lambda x: f"{x:.1f} د")
            fig_avg = go.Figure()
            for _, row in avg_dur.iterrows():
                color = COLOR_OK if row["النتيجة"] == "ناجحة" else COLOR_BAD
                fig_avg.add_trace(go.Bar(
                    x=[row["النتيجة"]],
                    y=[row[duration_col]],
                    text=[row["label"]],
                    textposition="outside",
                    textfont=dict(size=14, color="#0D2137"),
                    marker_color=color,
                    name=row["النتيجة"],
                    hovertemplate="%{x}: %{y:.1f} دقيقة<extra></extra>"
                ))
            fig_avg.update_layout(
                height=310, showlegend=False, margin=dict(t=20, b=30, l=10, r=10),
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                yaxis_title="دقيقة", xaxis_title="",
                uniformtext_minsize=12, uniformtext_mode="show"
            )
            fig_avg.update_yaxes(showgrid=True, gridcolor="#E8EEF5", range=[0, max(avg_dur[duration_col].max() * 1.25, 1)])
            st.plotly_chart(fig_avg, use_container_width=True)
        else:
            st.info("لا توجد بيانات مدة مكالمة")

    # ========== صف 2: تحليل المدة ==========
    if has_duration and len(duration):
        st.markdown("##### تحليل مدة المكالمة")
        d1, d2 = st.columns([1.6, 1])

        with d1:
            # histogram with count labels on bars
            counts, bins = np.histogram(duration.dropna(), bins=20)
            centers = 0.5 * (bins[:-1] + bins[1:])
            fig_hist = go.Figure(go.Bar(
                x=centers,
                y=counts,
                text=[str(int(c)) if c > 0 else "" for c in counts],
                textposition="outside",
                textfont=dict(size=11, color="#0D2137"),
                marker_color=PRIMARY,
                hovertemplate="المدة ≈ %{x:.1f} د<br>العدد: %{y}<extra></extra>"
            ))
            fig_hist.update_layout(
                height=300, margin=dict(t=25, b=40, l=10, r=10), bargap=0.08,
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                xaxis_title="المدة (دقيقة)", yaxis_title="عدد المكالمات"
            )
            fig_hist.update_xaxes(showgrid=False)
            fig_hist.update_yaxes(showgrid=True, gridcolor="#E2E8F0", range=[0, max(counts.max() * 1.2, 1)])
            st.plotly_chart(fig_hist, use_container_width=True)

        with d2:
            short = int((duration <= 5).sum())
            mid = int(((duration > 5) & (duration <= 15)).sum())
            long_ = int((duration > 15).sum())
            fig_bins = go.Figure(data=[go.Pie(
                labels=["≤ 5 د", "5–15 د", "> 15 د"],
                values=[short, mid, long_],
                hole=0.48,
                marker=dict(colors=[COLOR_OK, COLOR_MID, COLOR_BAD], line=dict(color="white", width=2)),
                textinfo="label+percent+value",
                texttemplate="%{label}<br>%{value}<br>(%{percent})",
                textfont=dict(size=12, color="#0D2137"),
                hovertemplate="%{label}: %{value} (%{percent})<extra></extra>"
            )])
            fig_bins.update_layout(
                height=300, margin=dict(t=10, b=10, l=10, r=10), showlegend=False,
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)"
            )
            st.plotly_chart(fig_bins, use_container_width=True)

        dm1, dm2, dm3, dm4 = st.columns(4)
        dm1.metric("متوسط مدة المكالمات", f"{duration.mean():.1f} د")
        dm2.metric("أقل مدة مكالمة", f"{duration.min():.1f} د")
        dm3.metric("أعلى مدة مكالمة", f"{duration.max():.1f} د")
        dm4.metric("الوسيط", f"{duration.median():.1f} د")


    # ========== عدد المكالمات الناجحة حسب سنة الإسناد ==========
    if assign_col and assign_col in df.columns:
        st.markdown("##### عدد المكالمات الناجحة من سنة الإسناد")
        tmp_a = df.copy()
        tmp_a[assign_col] = pd.to_datetime(tmp_a[assign_col], errors="coerce")
        tmp_a = tmp_a.dropna(subset=[assign_col])
        tmp_a = tmp_a[tmp_a[success_col] == 1]
        if len(tmp_a):
            tmp_a["سنة_الإسناد"] = tmp_a[assign_col].dt.year.astype(int)
            by_year = (
                tmp_a.groupby("سنة_الإسناد")
                .size()
                .reset_index(name="عدد_المكالمات_الناجحة")
                .sort_values("سنة_الإسناد")
            )
            fig_year = px.bar(
                by_year,
                x="سنة_الإسناد",
                y="عدد_المكالمات_الناجحة",
                text="عدد_المكالمات_الناجحة",
                color_discrete_sequence=[PRIMARY]
            )
            fig_year.update_traces(textposition="outside", marker_line_width=0)
            fig_year.update_layout(
                title="عدد المكالمات الناجحة من سنة الإسناد",
                height=340,
                xaxis_title="",
                yaxis_title="",
                margin=dict(t=50, b=30, l=10, r=10),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                showlegend=False
            )
            fig_year.update_xaxes(type="category", showgrid=False)
            fig_year.update_yaxes(showgrid=True, gridcolor="#E8EEF5")
            st.plotly_chart(fig_year, use_container_width=True)
        else:
            st.info("لا توجد مكالمات ناجحة بتاريخ إسناد صالح")

    # ========== عدد المكالمات الناجحة من سنين الحادث ==========
    if incident_col and incident_col in df.columns:
        st.markdown("##### عدد المكالمات الناجحة من سنين الحادث")
        tmp_i = df.copy()
        tmp_i[incident_col] = pd.to_datetime(tmp_i[incident_col], errors="coerce")
        tmp_i = tmp_i.dropna(subset=[incident_col])
        tmp_i = tmp_i[tmp_i[success_col] == 1]
        if len(tmp_i):
            tmp_i["سنة_الحادث"] = tmp_i[incident_col].dt.year.astype(int)
            by_inc = (
                tmp_i.groupby("سنة_الحادث")
                .size()
                .reset_index(name="عدد_المكالمات_الناجحة")
                .sort_values("سنة_الحادث")
            )
            fig_inc = px.bar(
                by_inc,
                x="سنة_الحادث",
                y="عدد_المكالمات_الناجحة",
                text="عدد_المكالمات_الناجحة",
                color_discrete_sequence=["#1E88E5"]
            )
            fig_inc.update_traces(textposition="outside", marker_line_width=0)
            fig_inc.update_layout(
                title="عدد المكالمات الناجحة من سنين الحادث",
                height=340,
                xaxis_title="",
                yaxis_title="",
                margin=dict(t=50, b=30, l=10, r=10),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                showlegend=False
            )
            fig_inc.update_xaxes(type="category", showgrid=False)
            fig_inc.update_yaxes(showgrid=True, gridcolor="#E8EEF5")
            st.plotly_chart(fig_inc, use_container_width=True)
        else:
            st.info("لا توجد مكالمات ناجحة بتاريخ حادث صالح")

    # ========== صف 3: أداء المحصلين (بنفس شكل اللوحة) ==========
    if collector_col and collector_col in df.columns:
        st.markdown("##### أداء المحصلين")

        stats = (
            df.groupby(collector_col)
            .agg(اجمالي_المكالمات=(success_col, "count"), عدد_المكالمات_الناجحة=(success_col, "sum"))
            .reset_index()
        )
        stats["نسبة_النجاح"] = (
            stats["عدد_المكالمات_الناجحة"] / stats["اجمالي_المكالمات"] * 100
        ).round(2)
        if has_duration:
            avg_d = df.groupby(collector_col)[duration_col].mean().round(2)
            stats = stats.merge(
                avg_d.rename("متوسط_المدة"),
                left_on=collector_col,
                right_index=True,
                how="left"
            )
        stats = stats.sort_values("اجمالي_المكالمات", ascending=True)  # للبار الأفقي

        left, right = st.columns(2)

        with left:
            fig_vol = px.bar(
                stats,
                y=collector_col,
                x="اجمالي_المكالمات",
                orientation="h",
                text="اجمالي_المكالمات",
                color_discrete_sequence=[PRIMARY]
            )
            fig_vol.update_traces(textposition="outside", marker_line_width=0)
            fig_vol.update_layout(
                title="اجمالي المكالمات لكل محصل",
                height=340,
                xaxis_title="",
                yaxis_title="",
                margin=dict(t=45, b=20, l=10, r=30),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                showlegend=False
            )
            fig_vol.update_xaxes(showgrid=True, gridcolor="#E8EEF5")
            fig_vol.update_yaxes(showgrid=False)
            st.plotly_chart(fig_vol, use_container_width=True)

        with right:
            # ترتيب تنازلي للشكل المساحي
            stats_area = stats.sort_values("عدد_المكالمات_الناجحة", ascending=False)
            fig_area = go.Figure()
            fig_area.add_trace(go.Scatter(
                x=stats_area[collector_col],
                y=stats_area["عدد_المكالمات_الناجحة"],
                mode="lines+markers",
                fill="tozeroy",
                line=dict(color=SECONDARY, width=3),
                fillcolor="rgba(21, 101, 192, 0.25)",
                marker=dict(size=8, color=PRIMARY),
                text=stats_area["عدد_المكالمات_الناجحة"],
                textposition="top center",
                hovertemplate="%{x}<br>الناجحة: %{y}<extra></extra>"
            ))
            fig_area.update_layout(
                title="عدد المكالمات الناجحة لكل محصل",
                height=340,
                xaxis_title="",
                yaxis_title="",
                margin=dict(t=45, b=60, l=10, r=10),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)"
            )
            fig_area.update_xaxes(tickangle=25, showgrid=False)
            fig_area.update_yaxes(showgrid=True, gridcolor="#E8EEF5")
            st.plotly_chart(fig_area, use_container_width=True)

        # بطاقات سريعة
        stats_desc = stats.sort_values("اجمالي_المكالمات", ascending=False)
        if len(stats_desc):
            most = stats_desc.iloc[0]
            best = stats_desc.sort_values("نسبة_النجاح", ascending=False).iloc[0]
            b1, b2, b3 = st.columns(3)
            b1.metric(
                "الأكثر مكالمات",
                str(most[collector_col]),
                delta=f"{int(most['اجمالي_المكالمات'])} مكالمة"
            )
            b2.metric(
                "أعلى نسبة نجاح",
                str(best[collector_col]),
                delta=f"{best['نسبة_النجاح']}%"
            )
            b3.metric("إجمالي المكالمات الناجحة", f"{success_count:,}")

        display_df = stats_desc.rename(columns={
            collector_col: "المحصل",
            "اجمالي_المكالمات": "اجمالي المكالمات",
            "عدد_المكالمات_الناجحة": "المكالمات الناجحة",
            "نسبة_النجاح": "نسبة النجاح %",
            "متوسط_المدة": "متوسط المدة"
        })
        st.dataframe(display_df, use_container_width=True, height=260)

    # ========== توزيع زمني لو فيه تاريخ ==========
    if created_col and created_col in df.columns:
        tmp = df.copy()
        tmp[created_col] = pd.to_datetime(tmp[created_col], errors="coerce")
        tmp = tmp.dropna(subset=[created_col])
        if len(tmp):
            st.markdown("##### التوزيع الزمني للإفادات")
            t1, t2 = st.columns(2)
            with t1:
                by_day = tmp.groupby(tmp[created_col].dt.date).size().reset_index(name="العدد")
                by_day.columns = ["اليوم", "العدد"]
                fig_day = px.line(by_day, x="اليوم", y="العدد", markers=True, color_discrete_sequence=[PRIMARY])
                fig_day.update_layout(
                    title="عدد الإفادات حسب اليوم", height=280,
                    margin=dict(t=40, b=30), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)"
                )
                st.plotly_chart(fig_day, use_container_width=True)
            with t2:
                tmp["ساعة"] = tmp[created_col].dt.hour
                by_hour = tmp.groupby("ساعة").size().reset_index(name="العدد")
                fig_hour = px.bar(by_hour, x="ساعة", y="العدد", color_discrete_sequence=[SECONDARY])
                fig_hour.update_layout(
                    title="عدد الإفادات حسب الساعة", height=280,
                    margin=dict(t=40, b=30), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)"
                )
                st.plotly_chart(fig_hour, use_container_width=True)

            # نجاح حسب اليوم
            by_day_s = tmp.groupby(tmp[created_col].dt.date).agg(
                إجمالي=(success_col, "count"),
                ناجحة=(success_col, "sum")
            ).reset_index()
            by_day_s.columns = ["اليوم", "إجمالي", "ناجحة"]
            by_day_s["نسبة النجاح %"] = (by_day_s["ناجحة"] / by_day_s["إجمالي"] * 100).round(1)
            fig_sr = px.bar(by_day_s, x="اليوم", y="نسبة النجاح %", color="نسبة النجاح %",
                            color_continuous_scale=["#FECACA", SUCCESS])
            fig_sr.update_layout(
                title="نسبة النجاح حسب اليوم", height=280, coloraxis_showscale=False,
                margin=dict(t=40, b=30), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)"
            )
            st.plotly_chart(fig_sr, use_container_width=True)

    st.caption("شركة إجادة لتحصيل الديون — تقرير متابعة نشاط المحصلين")


# ============================================================
# المرحلة 3: دمج شيتات مصنّفة + Pivot + تحميل
# ============================================================
elif mode.startswith("③"):
    st.subheader("المرحلة 3 — دمج الشيتات وعمل Pivot")
    st.caption("ارفع أي عدد من شيتات النشاط بعد التصنيف → دمجها في شيت واحد + Pivot Table → تحميل النتيجة")

    uploaded_files = st.file_uploader(
        "ارفع الشيتات المصنّفة (واحد أو أكثر)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="merge_upload"
    )

    if not uploaded_files:
        st.info("ارفع شيت أو أكثر بعد التصنيف عشان ندمجهم ونطلع Pivot")
        st.stop()

    def _score_sheet_merge(xl, name):
        try:
            tmp = pd.read_excel(xl, sheet_name=name, nrows=30)
            tmp = tmp.dropna(axis=1, how="all")
            cols = [str(c).strip().lower() for c in tmp.columns]
            score = len(cols)
            for kw in ["ناجحه", "ناجح", "notes", "محصل", "collector", "created", "مدة", "account", "اكونت"]:
                if any(kw in c for c in cols):
                    score += 15
            if len(cols) <= 2:
                score -= 20
            joined = " ".join(cols)
            if "count of" in joined or "تقرير" in joined:
                score -= 25
            return score
        except Exception:
            return -999

    def _norm_key(name):
        """توحيد شكل اسم العمود للمقارنة (ة→ه، تشكيل مبسّط)"""
        s = str(name).strip()
        s = s.replace("ة", "ه").replace("ى", "ي")
        s = s.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
        s = " ".join(s.split())
        return s.lower()

    def _unify_columns(frame):
        """توحيد الأعمدة المترادفة داخل كل شيت (ناجحه/ناجحة/…) قبل الدمج"""
        frame = frame.copy()
        frame.columns = frame.columns.astype(str).str.strip().str.replace(r"\s+", " ", regex=True)

        # مجموعات مترادفات → الاسم الموحّد
        groups = {
            "ناجحه": [
                "ناجحه", "ناجحة", "ناجح", "نجاح", "success", "successful",
                "classification", "تصنيف", "نتيجه", "نتيجة"
            ],
            "مدة المكالمة": ["مدة المكالمة", "مده المكالمه", "المدة", "المده", "duration", "call duration"],
            "الوقت": ["الوقت", "وقت", "time"],
            "دبليكات": ["دبليكات", "دوبليكيت", "duplicate", "duplicates"],
        }

        # ابحث عن أعمدة تطابق كل مجموعة واجمعها في عمود واحد
        used = set()
        for target, aliases in groups.items():
            alias_keys = {_norm_key(a) for a in aliases}
            matched = []
            for c in frame.columns:
                if c in used or c == "__مصدر_الملف__":
                    continue
                ck = _norm_key(c)
                # تطابق تام أو يحتوي كلمة ناجح/success للنجاح فقط
                if ck in alias_keys:
                    matched.append(c)
                elif target == "ناجحه" and ("ناجح" in ck or ck == "success"):
                    matched.append(c)
            if not matched:
                continue
            # ادمج القيم: أول قيمة غير فارغة من الأعمدة المتطابقة
            series = frame[matched[0]]
            for c in matched[1:]:
                series = series.combine_first(frame[c])
            frame = frame.drop(columns=matched)
            frame[target] = series
            used.add(target)

        return frame

    def _read_file_merge(file_obj, filename):
        xl = pd.ExcelFile(file_obj)
        scored = sorted(((_score_sheet_merge(xl, n), n) for n in xl.sheet_names), reverse=True)
        sheet = scored[0][1]
        df_local = None
        for hdr in [0, 1, 2, 3, 4, 5]:
            trial = pd.read_excel(xl, sheet_name=sheet, header=hdr)
            trial = trial.dropna(axis=1, how="all")
            trial.columns = trial.columns.astype(str).str.strip().str.replace(r"\s+", " ", regex=True)
            unnamed = [c for c in trial.columns if str(c).startswith("Unnamed")]
            trial = trial.drop(columns=unnamed, errors="ignore")
            if len(trial.columns) < 2:
                continue
            hits = sum(
                1 for c in trial.columns
                for kw in ["ناجحه", "ناجح", "notes", "محصل", "collector", "created", "مدة", "account"]
                if kw in str(c).lower()
            )
            if hits >= 1 or (len(trial.columns) >= 4 and hdr == 0):
                df_local = trial
                break
            if df_local is None or len(trial.columns) > len(df_local.columns):
                df_local = trial
        if df_local is None:
            df_local = pd.read_excel(xl, sheet_name=sheet)
            df_local.columns = df_local.columns.astype(str).str.strip()
        df_local = _unify_columns(df_local)
        df_local["__مصدر_الملف__"] = filename
        return df_local, sheet

    frames, logs = [], []
    for f in uploaded_files:
        try:
            part, sheet = _read_file_merge(f, f.name)
            frames.append(part)
            logs.append({"الملف": f.name, "الشيت": sheet, "الصفوف": len(part), "الأعمدة": len(part.columns) - 1})
        except Exception as e:
            st.error(f"تعذر قراءة {f.name}: {e}")

    if not frames:
        st.error("لم يتم قراءة أي ملف")
        st.stop()

    df = pd.concat(frames, ignore_index=True, sort=False)
    df.columns = df.columns.astype(str).str.strip().str.replace(r"\s+", " ", regex=True)
    # توحيد نهائي بعد الدمج لو فضلت أعمدة متشابهة
    df = _unify_columns(df)

    st.success(f"تم دمج **{len(uploaded_files)}** ملف — الإجمالي **{len(df):,}** صف")
    st.dataframe(pd.DataFrame(logs), use_container_width=True)

    st.markdown("##### تأكيد أعمدة الـ Pivot")
    all_cols = ["— لا يوجد —"] + [c for c in df.columns if c != "__مصدر_الملف__"]

    def _def(cands, kws=None):
        found = find_column(df, cands, kws)
        return all_cols.index(found) if found and found in all_cols else 0

    c1, c2, c3 = st.columns(3)
    with c1:
        collector_col = st.selectbox(
            "عمود المحصل", all_cols,
            index=_def(["Collector", "اسم المحصل", "المحصل"], ["محصل", "collector", "agent"])
        )
        success_col = st.selectbox(
            "عمود النتيجة (ناجحه)", all_cols,
            index=_def(["ناجحه", "ناجحة", "success"], ["ناجح", "success"])
        )
    with c2:
        account_col = st.selectbox(
            "عمود رقم الحساب", all_cols,
            index=_def(["Customer Account Number", "Account", "رقم الاكونت", "رقم الحساب"], ["ACCOUNT", "اكونت", "حساب"])
        )
        duration_col = st.selectbox(
            "عمود مدة المكالمة", all_cols,
            index=_def(["مدة المكالمة", "المدة"], ["مدة", "duration"])
        )
    with c3:
        created_col = st.selectbox(
            "عمود التاريخ", all_cols,
            index=_def(["Created on", "CREATED ON", "Created On"], ["CREATED", "تاريخ"])
        )
        notes_col = st.selectbox(
            "عمود الإفادة (اختياري)", all_cols,
            index=_def(["Notes", "ملاحظات", "الإفادة"], ["notes", "ملاحظة"])
        )

    def _cl(v):
        return None if v == "— لا يوجد —" else v

    collector_col = _cl(collector_col)
    success_col = _cl(success_col)
    account_col = _cl(account_col)
    duration_col = _cl(duration_col)
    created_col = _cl(created_col)
    notes_col = _cl(notes_col)

    if not st.button("دمج + إنشاء Pivot + تحميل", type="primary", use_container_width=True):
        st.stop()

    work = df.copy()

    # تطبيع عمود النجاح إن وجد
    if success_col:
        raw_s = work[success_col]
        if raw_s.dtype == object:
            mapped = raw_s.astype(str).str.strip().str.lower().map({
                "1": 1, "0": 0, "نعم": 1, "لا": 0, "true": 1, "false": 0,
                "ناجحة": 1, "غير ناجحة": 0, "yes": 1, "no": 0
            })
            work[success_col] = pd.to_numeric(mapped, errors="coerce").fillna(
                pd.to_numeric(raw_s, errors="coerce")
            ).fillna(0).astype(int)
        else:
            work[success_col] = pd.to_numeric(raw_s, errors="coerce").fillna(0).astype(int)

    if duration_col:
        work[duration_col] = pd.to_numeric(work[duration_col], errors="coerce")

    # ---- Pivot المحصلين ----
    pivot_collector = None
    if collector_col:
        agg = {collector_col: "count"}
        rename_map = {collector_col: "اجمالي المكالمات"}
        if success_col:
            work["_ناجحة_"] = (work[success_col] == 1).astype(int)
            agg["_ناجحة_"] = "sum"
        if account_col:
            # نستخدم nunique عبر groupby منفصل
            pass
        if duration_col:
            agg[duration_col] = "mean"

        g = work.groupby(collector_col, dropna=False)
        pivot_collector = g.size().reset_index(name="اجمالي المكالمات")

        if success_col:
            ok = g["_ناجحة_"].sum().reset_index(name="المكالمات الناجحة")
            pivot_collector = pivot_collector.merge(ok, on=collector_col, how="left")
            pivot_collector["المكالمات غير الناجحة"] = (
                pivot_collector["اجمالي المكالمات"] - pivot_collector["المكالمات الناجحة"]
            )
            pivot_collector["نسبة النجاح %"] = (
                pivot_collector["المكالمات الناجحة"] / pivot_collector["اجمالي المكالمات"] * 100
            ).round(2)

        if account_col:
            acc = g[account_col].nunique().reset_index(name="اجمالي الحسابات")
            pivot_collector = pivot_collector.merge(acc, on=collector_col, how="left")

        if duration_col:
            dur = g[duration_col].mean().reset_index(name="متوسط مدة المكالمة")
            pivot_collector = pivot_collector.merge(dur, on=collector_col, how="left")
            pivot_collector["متوسط مدة المكالمة"] = pivot_collector["متوسط مدة المكالمة"].round(2)

        pivot_collector = pivot_collector.rename(columns={collector_col: "المحصل"})
        pivot_collector = pivot_collector.sort_values("اجمالي المكالمات", ascending=False)

        # صف إجمالي
        total_row = {"المحصل": "الإجمالي"}
        total_row["اجمالي المكالمات"] = int(pivot_collector["اجمالي المكالمات"].sum())
        if "المكالمات الناجحة" in pivot_collector.columns:
            total_row["المكالمات الناجحة"] = int(pivot_collector["المكالمات الناجحة"].sum())
            total_row["المكالمات غير الناجحة"] = int(pivot_collector["المكالمات غير الناجحة"].sum())
            total_row["نسبة النجاح %"] = round(
                total_row["المكالمات الناجحة"] / total_row["اجمالي المكالمات"] * 100, 2
            ) if total_row["اجمالي المكالمات"] else 0
        if "اجمالي الحسابات" in pivot_collector.columns and account_col:
            total_row["اجمالي الحسابات"] = int(work[account_col].nunique())
        if "متوسط مدة المكالمة" in pivot_collector.columns and duration_col:
            total_row["متوسط مدة المكالمة"] = round(float(work[duration_col].mean()), 2)
        pivot_collector = pd.concat([pivot_collector, pd.DataFrame([total_row])], ignore_index=True)

    # ---- Pivot يومي إن وجد تاريخ ----
    pivot_daily = None
    if created_col and created_col in work.columns:
        work["_تاريخ_"] = pd.to_datetime(work[created_col], errors="coerce")
        tmp = work.dropna(subset=["_تاريخ_"]).copy()
        if len(tmp):
            tmp["_يوم_"] = tmp["_تاريخ_"].dt.date
            daily = tmp.groupby("_يوم_").size().reset_index(name="اجمالي المكالمات")
            if success_col:
                daily_ok = tmp.groupby("_يوم_")["_ناجحة_"].sum().reset_index(name="المكالمات الناجحة")
                daily = daily.merge(daily_ok, on="_يوم_", how="left")
                daily["نسبة النجاح %"] = (
                    daily["المكالمات الناجحة"] / daily["اجمالي المكالمات"] * 100
                ).round(2)
            daily = daily.rename(columns={"_يوم_": "اليوم"}).sort_values("اليوم")
            pivot_daily = daily

    # عرض
    st.markdown("##### معاينة البيانات المدموجة")
    st.dataframe(work.drop(columns=["__مصدر_الملف__"], errors="ignore").head(20), use_container_width=True)

    if pivot_collector is not None:
        st.markdown("##### Pivot — حسب المحصل")
        st.dataframe(pivot_collector, use_container_width=True)

    if pivot_daily is not None:
        st.markdown("##### Pivot — حسب اليوم")
        st.dataframe(pivot_daily, use_container_width=True)

    # تحميل
    today = datetime.today().strftime("%Y-%m-%d")
    out_name = f"EJADA_merged_pivot_{today}.xlsx"
    export_df = work.drop(columns=["_ناجحة_", "_تاريخ_", "_يوم_"], errors="ignore")

    with pd.ExcelWriter(out_name, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="البيانات المدموجة")
        if pivot_collector is not None:
            pivot_collector.to_excel(writer, index=False, sheet_name="Pivot المحصلين")
        if pivot_daily is not None:
            pivot_daily.to_excel(writer, index=False, sheet_name="Pivot يومي")
        # ملخص سريع
        summary = pd.DataFrame([
            {"البند": "عدد الملفات", "القيمة": len(uploaded_files)},
            {"البند": "اجمالي الصفوف", "القيمة": len(export_df)},
            {"البند": "اجمالي المكالمات", "القيمة": len(export_df)},
            {"البند": "اجمالي الحسابات", "القيمة": int(export_df[account_col].nunique()) if account_col else "—"},
            {"البند": "المكالمات الناجحة", "القيمة": int((work[success_col] == 1).sum()) if success_col else "—"},
            {"البند": "نسبة النجاح %", "القيمة": round(float((work[success_col] == 1).mean() * 100), 2) if success_col else "—"},
        ])
        summary.to_excel(writer, index=False, sheet_name="ملخص")

    with open(out_name, "rb") as f:
        st.download_button(
            "تحميل الشيت المدموج + الـ Pivot",
            data=f,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    st.success("جاهز — الملف فيه: البيانات المدموجة + Pivot المحصلين + Pivot يومي + ملخص")


# ============================================================
# ④ الوعود
# ============================================================
elif mode.startswith("④"):
    st.subheader("④ الوعود — قائمة ومكسورة")
    st.caption("ارفع ملف المحفظة → Sub State = واعد بالسداد + تاريخ المتابعة")
    company = st.radio("الشركة", ["سايكو", "ولاء"], horizontal=True, key="promises_company")
    tab_stand, tab_broken = st.tabs(["الوعود القائمة", "الوعود المكسورة"])
    SALES_CANDS = ["Sales Person", "SalesPerson", "Create By", "Created By", "Collector", "المحصل", "المحصّل"]
    SUB_CANDS = ["Sub State", "SubState", "الحالة الفرعية"]
    DUE_CANDS = ["Follow up Due Date", "Follow Up Due Date", "Due Date", "تاريخ المتابعة"]
    NET_CANDS = ["Net Amount", "NetAmount", "صافي المبلغ", "مبلغ المديونية"]
    TARGET_SUB = "واعد بالسداد"

    def _run_promises(uploaded, due_mode, label):
        df, sheet = _ops_read_excel(uploaded)
        sales_col = _ops_find_col(df, SALES_CANDS)
        sub_col = _ops_find_col(df, SUB_CANDS)
        due_col = _ops_find_col(df, DUE_CANDS)
        net_col = _ops_find_col(df, NET_CANDS)
        missing = [n for n, c in [("المحصل", sales_col), ("Sub State", sub_col), ("Due Date", due_col)] if not c]
        if missing:
            st.error("أعمدة ناقصة: " + ", ".join(missing))
            st.write("الأعمدة الموجودة:", list(df.columns))
            return
        work = df[df[sub_col].astype(str).str.strip() == TARGET_SUB].copy()
        due_dates = work[due_col].map(_ops_parse_date)
        today = datetime.today().date()
        if due_mode == "today":
            work = work[due_dates == today].copy()
            mode_txt = f"تاريخ المتابعة = اليوم ({today})"
        else:
            work = work[due_dates.apply(lambda d: d is not None and d < today)].copy()
            mode_txt = f"تاريخ المتابعة قبل اليوم (< {today})"
        st.info(f"{company} | {mode_txt} | الشيت: {sheet}")
        k1, k2, k3 = st.columns(3)
        k1.metric(label, f"{len(work):,}")
        if net_col:
            net_sum = pd.to_numeric(work[net_col], errors="coerce").fillna(0).sum()
            k2.metric("إجمالي المديونية", f"{net_sum:,.0f}")
        k3.metric("عدد المحصلين", f"{work[sales_col].nunique()}")
        if len(work) == 0:
            st.warning("لا توجد نتائج بهذه الفلاتر")
            return
        summary = work.groupby(sales_col).size().reset_index(name=label)
        if net_col:
            amounts = pd.to_numeric(work[net_col], errors="coerce").fillna(0)
            amt = work.assign(_amt=amounts).groupby(sales_col)["_amt"].sum().reset_index(name="المديونية")
            summary = summary.merge(amt, on=sales_col, how="left")
        summary = summary.sort_values(label, ascending=False)
        st.dataframe(summary, use_container_width=True)
        st.dataframe(work.head(50), use_container_width=True)
        _ops_to_excel_download({"التفاصيل": work, "تجميع المحصلين": summary}, f"EJADA_{company}_{label}_{datetime.today().strftime('%Y-%m-%d')}.xlsx")

    with tab_stand:
        up = st.file_uploader("ملف المحفظة — وعود قائمة", type=["xlsx", "xls"], key="prom_stand")
        if up and st.button("تحليل الوعود القائمة", type="primary", key="btn_stand"):
            _run_promises(up, "today", "الوعود القائمة")
    with tab_broken:
        up2 = st.file_uploader("ملف المحفظة — وعود مكسورة", type=["xlsx", "xls"], key="prom_broken")
        if up2 and st.button("تحليل الوعود المكسورة", type="primary", key="btn_broken"):
            _run_promises(up2, "before", "الوعود المكسورة")

# ============================================================
# ⑤ الإهمال
# ============================================================
elif mode.startswith("⑤"):
    st.subheader("⑤ الإهمال")
    company = st.radio("الشركة", ["سايكو", "ولاء"], horizontal=True, key="neglect_company")

    # زي الأصلي: زرّين — إهمال / متابعة
    m1, m2 = st.columns(2)
    if "neglect_mode" not in st.session_state:
        st.session_state["neglect_mode"] = "neglect"
    with m1:
        if st.button(
            "الإهمال (أول مرة)",
            use_container_width=True,
            type="primary" if st.session_state["neglect_mode"] == "neglect" else "secondary",
            key="neg_mode_first",
        ):
            st.session_state["neglect_mode"] = "neglect"
            st.rerun()
    with m2:
        if st.button(
            "متابعة الإهمال",
            use_container_width=True,
            type="primary" if st.session_state["neglect_mode"] == "followup" else "secondary",
            key="neg_mode_follow",
        ):
            st.session_state["neglect_mode"] = "followup"
            st.rerun()

    st.divider()

    SALES_CANDS = ["Sales Person", "SalesPerson", "Create By", "Created By", "Collector", "المحصل", "المحصّل"]
    SUB_CANDS = ["Sub State", "SubState", "الحالة الفرعية"]
    DUE_CANDS = ["Follow up Due Date", "Follow Up Due Date", "Due Date", "تاريخ المتابعة"]
    LAST_CANDS = [
        "Follow up Last Date", "Follow Up Last Date", "Followup Last Date",
        "Last Follow Up", "Last Follow-up", "تاريخ آخر متابعة", "تاريخ اخر متابعة", "آخر متابعة",
    ]
    NET_CANDS = ["Net Amount", "NetAmount", "صافي المبلغ", "مبلغ المديونية"]
    ACC_CANDS = ["Customer Account Number", "Account Number", "Account ID", "Claim", "رقم الحساب", "رقم المطالبة"]
    DEFAULT_STATES = [
        "واعد بالسداد",
        "تم الاتصال",
        "تم الاتصال واتساب",
        "جدولة",
        "اعفاء بانتظار المستند",
        "إعفاء بانتظار المستند",
        "إعفاء || بإنتظار المستند",
        "رافض السداد",
        "رافض السداد نهائيا",
        "مسجون",
        "مسافر",
        "لا يرد",
        "لا يرد مع التكرار",
        "لايرد مع التكرار",
        "مغلق مع التكرار",
        "تم ابلاغ العميل اتصال",
        "تم ابلاغ العميل - اتصال",
        "تم ابلاغ العميل واتساب",
        "تم ابلاغ العميل - واتساب",
        "تم ابلاغ العميل - واتسب",
    ]

    # -------------------- أول مرة --------------------
    if st.session_state["neglect_mode"] == "neglect":
        st.caption(
            "من المحفظة: حالات Sub State المحددة + تاريخ المتابعة ≠ اليوم + "
            "مرّ على آخر متابعة أكثر من 7 أيام"
        )
        up = st.file_uploader("📂 ملف المحفظة", type=["xlsx", "xls", "csv"], key="neg_first_file")

        if up is not None:
            df, sheet = _ops_read_excel(up)
            sales_col = _ops_find_col(df, SALES_CANDS)
            sub_col = _ops_find_col(df, SUB_CANDS)
            due_col = _ops_find_col(df, DUE_CANDS)
            last_col = _ops_find_col(df, LAST_CANDS)
            net_col = _ops_find_col(df, NET_CANDS)

            missing = [n for n, c in [("المحصل", sales_col), ("Sub State", sub_col), ("Follow up Due Date", due_col)] if not c]
            if missing:
                st.error("أعمدة ناقصة: " + ", ".join(missing))
                st.write(list(df.columns))
                st.stop()

            # كل الحالات الموجودة في الملف عشان تختار منها
            file_states = sorted({str(x).strip() for x in df[sub_col].dropna().tolist() if str(x).strip() and str(x).strip().lower() != "nan"})
            default_sel = [s for s in DEFAULT_STATES if s in file_states] or file_states[: min(9, len(file_states))]
            selected_states = st.multiselect(
                "حالات Sub State المستهدفة للإهمال",
                options=file_states,
                default=default_sel,
                key="neg_states_dyn",
            )

            if st.button("استخراج تقرير الإهمال", type="primary", key="btn_neg_first"):
                if not selected_states:
                    st.error("اختَر حالة واحدة على الأقل")
                    st.stop()

                today = datetime.today().date()
                total0 = len(df)
                work = df[df[sub_col].astype(str).str.strip().isin(selected_states)].copy()
                after_sub = len(work)

                # 1) استبعاد Follow up Due Date == اليوم
                due_dates = [_ops_parse_date(v) for v in work[due_col]]
                work = work[[d != today for d in due_dates]].copy()
                after_due = len(work)

                # 2) فرق الأيام من Follow up Last Date — أكتر من 7 أيام
                if last_col:
                    last_dates = [_ops_parse_date(v) for v in work[last_col]]
                    # زي الأصلي: لو مفيش تاريخ → 0 (وبالتالي يتشال لأن مش > 7)
                    days = [(today - d).days if d else 0 for d in last_dates]
                    work = work.copy()
                    work["فرق_الأيام"] = days
                    work = work[work["فرق_الأيام"] > 7].copy()
                else:
                    work = work.copy()
                    work["فرق_الأيام"] = 0
                    st.warning("عمود Follow up Last Date غير موجود — لازم يكون موجود لحساب فرق الأيام")
                    work = work.iloc[0:0].copy()
                after_days = len(work)

                st.info(
                    f"{company} | الشيت: {sheet} | اليوم: {today}\n\n"
                    f"من الملف: **{total0:,}** → بعد Sub State: **{after_sub:,}** → "
                    f"بعد استبعاد متابعته اليوم: **{after_due:,}** → بعد >7 أيام: **{after_days:,}**"
                )

                c1, c2, c3, c4 = st.columns(4)
                c1.metric("حالات الإهمال", f"{len(work):,}")
                c2.metric("عدد المحصلين", f"{work[sales_col].nunique() if len(work) else 0}")
                if net_col and len(work):
                    c3.metric(
                        "إجمالي المديونية",
                        f"{pd.to_numeric(work[net_col], errors='coerce').fillna(0).sum():,.0f}",
                    )
                if len(work) and "فرق_الأيام" in work.columns:
                    c4.metric("متوسط فرق الأيام", f"{pd.to_numeric(work['فرق_الأيام'], errors='coerce').mean():.1f}")

                if len(work) == 0:
                    st.warning("لا توجد حالات إهمال بهذه الشروط")
                else:
                    summary = (
                        work.groupby(sales_col).size()
                        .reset_index(name="عدد حالات الإهمال")
                        .sort_values("عدد حالات الإهمال", ascending=False)
                    )
                    st.markdown("##### تجميع حسب المحصل")
                    st.dataframe(summary, use_container_width=True)
                    st.markdown("##### تفاصيل الحالات")
                    st.dataframe(work, use_container_width=True)
                    _ops_to_excel_download(
                        {"تقرير الإهمال": work, "تجميع المحصلين": summary},
                        f"EJADA_{company}_neglect_{today}.xlsx",
                    )

    # -------------------- متابعة --------------------
    else:
        st.caption(
            "تقرير إهمال قديم + محفظة حديثة → تحديث تاريخ آخر متابعة + "
            "تم التغطية إذا مرّ أقل من 8 أيام"
        )
        c1, c2 = st.columns(2)
        with c1:
            new_f = st.file_uploader("📂 محفظة اليوم الحديثة", type=["xlsx", "xls", "csv"], key="neg_new")
        with c2:
            old_f = st.file_uploader("📂 تقرير الإهمال القديم", type=["xlsx", "xls", "csv"], key="neg_old")

        if new_f and old_f and st.button("بدء المطابقة ومتابعة الإهمال", type="primary", key="btn_neg_follow"):
            df_new, _ = _ops_read_excel(new_f)
            df_old, _ = _ops_read_excel(old_f)
            acc_new = _ops_find_col(df_new, ACC_CANDS)
            acc_old = _ops_find_col(df_old, ACC_CANDS)
            last_new = _ops_find_col(df_new, LAST_CANDS)

            if not acc_new or not acc_old:
                st.error("تعذر إيجاد عمود رقم الحساب/المطالبة للمطابقة")
                st.write("المحفظة الحديثة:", list(df_new.columns))
                st.write("تقرير الإهمال:", list(df_old.columns))
                st.stop()
            if not last_new:
                st.error("تعذر إيجاد Follow up Last Date في المحفظة الحديثة")
                st.write(list(df_new.columns))
                st.stop()

            def _norm_id(v):
                s = str(v).strip()
                if s.lower() in {"nan", "none", ""}:
                    return ""
                if s.endswith(".0"):
                    s = s[:-2]
                return s

            # XLOOKUP: أكبر تاريخ آخر متابعة لكل حساب من المحفظة الحديثة
            lookup = {}
            for _, row in df_new.iterrows():
                k = _norm_id(row[acc_new])
                d = _ops_parse_date(row[last_new])
                if k and d and (k not in lookup or d > lookup[k]):
                    lookup[k] = d

            result = df_old.copy()
            target_col = "تاريخ اخر متابعة"
            result[target_col] = result[acc_old].map(lambda x: lookup.get(_norm_id(x)))
            today = datetime.today().date()
            result["فرق_الأيام"] = result[target_col].map(lambda v: (today - v).days if v else None)
            # زي الأصلي: أقل من 8 أيام = تم التغطية
            result["حالة_التغطية"] = result["فرق_الأيام"].map(
                lambda d: "تم التغطية" if d is not None and d < 8 else "لم يتم التغطية"
            )

            covered = int((result["حالة_التغطية"] == "تم التغطية").sum())
            not_cov = int((result["حالة_التغطية"] == "لم يتم التغطية").sum())
            a, b, c = st.columns(3)
            a.metric("إجمالي الحالات", f"{len(result):,}")
            b.metric("تم التغطية", f"{covered:,}")
            c.metric("لم يتم التغطية", f"{not_cov:,}")
            st.dataframe(result, use_container_width=True)
            _ops_to_excel_download(
                {"متابعة الإهمال": result},
                f"EJADA_{company}_neglect_followup_{today}.xlsx",
            )

# ============================================================
# ⑥ الجدولة المتعثرة
# ============================================================
elif mode.startswith("⑥"):
    st.subheader("⑥ الجدولة المتعثرة")
    st.caption("Sub State = جدولة + آخر سداد → منتظمة ≤60 يوم / متعثرة >60")
    company = st.radio("الشركة", ["سايكو", "ولاء"], horizontal=True, key="sched_company")
    wallet_f = st.file_uploader("ملف المحفظة", type=["xlsx", "xls"], key="sched_wallet")
    pay_f = st.file_uploader("ملف السدادات (اختياري)", type=["xlsx", "xls"], key="sched_pay")
    SUB_CANDS = ["Sub State", "SubState", "الحالة الفرعية"]
    DEB_CANDS = ["Debitor", "Debtor", "رقم المدين", "المدين"]
    ACC_CANDS = ["Customer Account Number", "Account Number", "رقم الحساب"]
    PAY_DATE_CANDS = ["Date of Creation", "Payment Date", "Creation Date", "تاريخ السداد", "تاريخ الإنشاء"]
    if wallet_f and st.button("تحليل الجدولة", type="primary"):
        df_w, _ = _ops_read_excel(wallet_f)
        sub_col = _ops_find_col(df_w, SUB_CANDS)
        deb_col = _ops_find_col(df_w, DEB_CANDS) or _ops_find_col(df_w, ACC_CANDS)
        if not sub_col or not deb_col:
            st.error("محتاج Sub State و Debitor/Account")
            st.write(list(df_w.columns))
            st.stop()
        sched = df_w[df_w[sub_col].astype(str).str.strip() == "جدولة"].copy()
        last_pay = {}
        if pay_f:
            df_p, _ = _ops_read_excel(pay_f)
            p_deb = _ops_find_col(df_p, DEB_CANDS) or _ops_find_col(df_p, ACC_CANDS)
            p_date = _ops_find_col(df_p, PAY_DATE_CANDS)
            if p_deb and p_date:
                for _, row in df_p.iterrows():
                    k = str(row[p_deb]).strip()
                    d = _ops_parse_date(row[p_date])
                    if k and d and (k not in last_pay or d > last_pay[k]):
                        last_pay[k] = d
        today = datetime.today().date()
        sched["تاريخ_آخر_سداد"] = sched[deb_col].map(lambda x: last_pay.get(str(x).strip()))
        sched["أيام_منذ_آخر_سداد"] = sched["تاريخ_آخر_سداد"].map(lambda d: (today - d).days if d else None)
        sched["حالة_الجدولة"] = sched["أيام_منذ_آخر_سداد"].map(
            lambda d: "منتظمة" if d is not None and d <= 60 else ("متعثرة" if d is not None else "بدون سداد")
        )
        a, b, c, d = st.columns(4)
        a.metric("إجمالي الجدولة", f"{len(sched):,}")
        b.metric("منتظمة", f"{int((sched['حالة_الجدولة']=='منتظمة').sum()):,}")
        c.metric("متعثرة", f"{int((sched['حالة_الجدولة']=='متعثرة').sum()):,}")
        d.metric("بدون سداد", f"{int((sched['حالة_الجدولة']=='بدون سداد').sum()):,}")
        st.dataframe(sched.head(100), use_container_width=True)
        _ops_to_excel_download({"الجدولة": sched}, f"EJADA_{company}_schedule_{today}.xlsx")

# ============================================================
# ⑦ أخطاء الحالات
# ============================================================
elif mode.startswith("⑦"):
    st.subheader("⑦ أخطاء الحالات")
    company = st.radio("الشركة", ["سايكو", "ولاء"], horizontal=True, key="case_company")

    if "case_mode" not in st.session_state:
        st.session_state["case_mode"] = "scan"
    m1, m2 = st.columns(2)
    with m1:
        if st.button(
            "فحص أخطاء الحالات",
            use_container_width=True,
            type="primary" if st.session_state["case_mode"] == "scan" else "secondary",
            key="case_mode_scan",
        ):
            st.session_state["case_mode"] = "scan"
            st.rerun()
    with m2:
        if st.button(
            "متابعة أخطاء الحالات",
            use_container_width=True,
            type="primary" if st.session_state["case_mode"] == "followup" else "secondary",
            key="case_mode_follow",
        ):
            st.session_state["case_mode"] = "followup"
            st.rerun()

    st.divider()

    SUB_CANDS = ["Sub State", "SubState", "الحالة الفرعية"]
    PAY_CANDS = ["Payment", "Paid Amount", "المبلغ المدفوع"]
    NET_CANDS = ["Net Amount", "NetAmount", "صافي المبلغ"]
    ACC_CANDS = [
        "Customer Account Number", "Account Number", "Account ID", "Claim",
        "رقم الحساب", "رقم المطالبة", "Debitor", "Debtor", "رقم المدين",
    ]
    MUST_POSITIVE = {"جدولة", "سدد كامل المديونية"}

    def _norm_id(v):
        s = str(v).strip()
        if s.lower() in {"nan", "none", ""}:
            return ""
        if s.endswith(".0"):
            s = s[:-2]
        return s

    _AR_JUNK = re.compile(r"[\u0640\u064B-\u065F\u0670\u200b-\u200f\u202a-\u202e\u2066-\u2069\ufeff]")
    DISCOUNT_WORDS = ("خصم", "خصوم", "تخفيض")  # ضيف أي كلمة تانية بتستخدموها للخصم

    def _state_key(value):
        """توحيد النص: شيل التطويل والتشكيل والحروف المخفية + توحيد الهمزات والتاء المربوطة"""
        s = unicodedata.normalize("NFKC", str(value or ""))
        s = _AR_JUNK.sub("", s)
        for a, b in (("ة", "ه"), ("أ", "ا"), ("إ", "ا"), ("آ", "ا"), ("ى", "ي")):
            s = s.replace(a, b)
        return " ".join(s.split()).lower()

    def _num(v):
        """تحويل آمن لرقم: الفاضي/النص غير الرقمي = 0 (بيدعم الفاصلة والأرقام العربي)"""
        if v is None or (not isinstance(v, str) and pd.isna(v)):
            return 0.0
        if isinstance(v, str):
            v = v.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩٫٬", "0123456789.,")).replace(",", "").strip()
        n = pd.to_numeric(v, errors="coerce")
        return 0.0 if pd.isna(n) else float(n)

    def _match_payment_state(value):
        """تصنيف Sub State لحالة سداد/جدولة معروفة (مع أو بدون خصم)"""
        key = _state_key(value)
        if not key:
            return None
        disc = any(w in key for w in DISCOUNT_WORDS)
        closed = "مقفل" in key or "مغلق" in key
        paid = ("سدد" in key or "سداد" in key) and ("مديون" in key or "كامل" in key)
        partial = "جزء" in key or "جزئ" in key

        if "جدوله" in key and closed:
            return "جدولة مقفلة بخصم" if disc else "جدولة مقفلة"
        if paid and partial:  # سداد جزئي: حالة سداد عادية، من غير فحص على المتبقي
            return "سداد جزئي"
        if paid:
            return "سدد كامل المديونية بخصم" if disc else "سدد كامل المديونية"
        if key.startswith("جدوله"):
            return "جدولة"
        return None

    def _errors_for_row(sub, pay, net):
        errs = []
        matched = _match_payment_state(sub)
        is_pay_state = matched is not None
        if matched in MUST_POSITIVE and pay == 0:
            errs.append("Payment صفر مع حالة سداد")
        if not is_pay_state and pay > 0:
            errs.append("Payment > 0 بدون حالة سداد")
        if matched == "سدد كامل المديونية" and net >= 50:
            errs.append("متبقي بعد سداد كامل")
        elif matched == "جدولة مقفلة" and net >= 50:
            errs.append("متبقي مع جدولة مقفلة")
        elif matched == "جدولة مقفلة بخصم" and net == 0:
            errs.append("Net Amount صفر مع جدولة مقفلة بخصم")
        elif matched == "سدد كامل المديونية بخصم" and net == 0:
            errs.append("Net Amount صفر مع سداد بخصم")
        # جدولة عادية + متبقي / سداد بخصم + Payment و Net فيهم رقم → طبيعي مش خطأ
        return errs

    def _detect_errors_df(df, sub_col, pay_col, net_col):
        rows = []
        for _, row in df.iterrows():
            sub = str(row[sub_col]).strip() if pd.notna(row[sub_col]) else ""
            pay = _num(row[pay_col]) if pay_col else 0.0
            net = _num(row[net_col]) if net_col else 0.0
            errs = _errors_for_row(sub, pay, net)
            if errs:
                rec = row.to_dict()
                rec["سبب_الخطأ"] = " | ".join(errs)
                rows.append(rec)
        return pd.DataFrame(rows)

    # ---------- فحص أول مرة ----------
    if st.session_state["case_mode"] == "scan":
        st.caption("ارفع المحفظة → كشف تعارض Sub State مع Payment / Net Amount")
        up = st.file_uploader("ملف المحفظة", type=["xlsx", "xls", "csv"], key="case_file")
        if up and st.button("فحص أخطاء الحالات", type="primary", key="btn_case_scan"):
            df, _ = _ops_read_excel(up)
            sub_col = _ops_find_col(df, SUB_CANDS)
            pay_col = _ops_find_col(df, PAY_CANDS)
            net_col = _ops_find_col(df, NET_CANDS)
            if not sub_col:
                st.error("محتاج عمود Sub State")
                st.stop()
            if not pay_col or not net_col:
                st.error("محتاج عمود Payment وعمود Net Amount عشان الفحص يكون دقيق")
                st.write("الأعمدة الموجودة:", list(df.columns))
                st.stop()
            err_df = _detect_errors_df(df, sub_col, pay_col, net_col)
            st.metric("عدد الأخطاء", f"{len(err_df):,}")
            with st.expander("🔍 تصنيف الحالات الموجودة في الملف"):
                d = df[sub_col].astype(str).str.strip().value_counts().reset_index()
                d.columns = ["Sub State", "العدد"]
                d["التصنيف"] = d["Sub State"].map(lambda v: _match_payment_state(v) or "غير معروف (عادي)")
                d["النص الخام"] = d["Sub State"].map(repr)
                st.dataframe(d, use_container_width=True)
                if len(err_df):
                    st.write("أسباب الأخطاء:")
                    st.dataframe(err_df["سبب_الخطأ"].value_counts().reset_index(), use_container_width=True)
            if len(err_df):
                st.dataframe(err_df, use_container_width=True)
                _ops_to_excel_download(
                    {"أخطاء الحالات": err_df},
                    f"EJADA_{company}_case_errors_{datetime.today().strftime('%Y-%m-%d')}.xlsx",
                )
            else:
                st.success("لا توجد أخطاء ظاهرة بالقواعد الحالية")

    # ---------- متابعة ----------
    else:
        st.caption(
            "تقرير أخطاء قديم + محفظة جديدة → الحالة الفرعية الجديدة جنب القديمة | "
            "تم تغير الحالة / لم يتم تغيير الحالة | + أخطاء جديدة"
        )
        c1, c2 = st.columns(2)
        with c1:
            old_f = st.file_uploader("📂 شيت أخطاء الحالات السابق", type=["xlsx", "xls", "csv"], key="case_old")
        with c2:
            new_f = st.file_uploader("📂 المحفظة الجديدة", type=["xlsx", "xls", "csv"], key="case_new")

        if old_f and new_f and st.button("بدء متابعة أخطاء الحالات", type="primary", key="btn_case_follow"):
            df_old, _ = _ops_read_excel(old_f)
            df_new, _ = _ops_read_excel(new_f)

            acc_old = _ops_find_col(df_old, ACC_CANDS)
            acc_new = _ops_find_col(df_new, ACC_CANDS)
            sub_old = _ops_find_col(df_old, SUB_CANDS)
            sub_new = _ops_find_col(df_new, SUB_CANDS)
            pay_new = _ops_find_col(df_new, PAY_CANDS)
            net_new = _ops_find_col(df_new, NET_CANDS)

            if not acc_old or not acc_new:
                st.error("محتاج عمود رقم الحساب/المطالبة في التقرير القديم والمحفظة الجديدة")
                st.write("التقرير القديم:", list(df_old.columns))
                st.write("المحفظة:", list(df_new.columns))
                st.stop()
            if not sub_new:
                st.error("محتاج عمود Sub State في المحفظة الجديدة")
                st.stop()
            if not pay_new or not net_new:
                st.error("محتاج عمود Payment وعمود Net Amount في المحفظة الجديدة")
                st.write("المحفظة:", list(df_new.columns))
                st.stop()

            def _row_still_error(row):
                """هل الصف لسه فيه خطأ حسب قواعد الفحص؟"""
                sub = str(row[sub_new]).strip() if pd.notna(row[sub_new]) else ""
                pay = _num(row[pay_new]) if pay_new else 0.0
                net = _num(row[net_new]) if net_new else 0.0
                return bool(_errors_for_row(sub, pay, net))

            # خريطة المحفظة الجديدة: حساب → الصف
            new_by_acc = {}
            for _, row in df_new.iterrows():
                k = _norm_id(row[acc_new])
                if k:
                    new_by_acc[k] = row

            # كل أخطاء المحفظة الجديدة (للأخطاء الجديدة)
            current_err = _detect_errors_df(df_new, sub_new, pay_new, net_new)

            follow_rows = []
            old_keys = set()

            for _, row in df_old.iterrows():
                k = _norm_id(row[acc_old])
                if not k:
                    continue
                old_keys.add(k)
                rec = row.to_dict()
                old_sub = ""
                if sub_old and sub_old in row.index and pd.notna(row[sub_old]):
                    old_sub = str(row[sub_old]).strip()
                # لو مفيش عمود sub في القديم، جرّب اسم شائع
                if not old_sub:
                    for alt in ("Sub State", "SubState", "الحالة الفرعية"):
                        if alt in row.index and pd.notna(row[alt]):
                            old_sub = str(row[alt]).strip()
                            break

                if k not in new_by_acc:
                    rec["الحالة الفرعية الجديدة"] = ""
                    rec["حالة_المتابعة"] = "غير موجود في المحفظة الجديدة"
                    follow_rows.append(rec)
                    continue

                new_row = new_by_acc[k]
                new_sub = str(new_row[sub_new]).strip() if pd.notna(new_row[sub_new]) else ""
                rec["الحالة الفرعية الجديدة"] = new_sub

                still_err = _row_still_error(new_row)
                # اتغيرت وبقت صح = الحالة الجديدة مش خطأ
                if not still_err:
                    rec["حالة_المتابعة"] = "تم تغير الحالة"
                    rec["سبب_الخطأ"] = ""  # امسح سبب الخطأ القديم
                else:
                    rec["حالة_المتابعة"] = "لم يتم تغيير الحالة"
                    pay = _num(new_row[pay_new]) if pay_new else 0.0
                    net = _num(new_row[net_new]) if net_new else 0.0
                    errs_now = _errors_for_row(new_sub, pay, net)
                    if errs_now:
                        rec["سبب_الخطأ"] = " | ".join(errs_now)
                follow_rows.append(rec)

            follow_df = pd.DataFrame(follow_rows)

            # أخطاء جديدة: في المحفظة ومش في التقرير القديم
            new_only_rows = []
            if len(current_err):
                for _, row in current_err.iterrows():
                    k = _norm_id(row[acc_new])
                    if k and k not in old_keys:
                        rec = row.to_dict()
                        rec["الحالة الفرعية الجديدة"] = str(row[sub_new]).strip() if sub_new in row.index and pd.notna(row[sub_new]) else ""
                        rec["حالة_المتابعة"] = "خطأ جديد"
                        new_only_rows.append(rec)
            new_only_df = pd.DataFrame(new_only_rows)

            combined = pd.concat([follow_df, new_only_df], ignore_index=True, sort=False)

            # رتّب الأعمدة: Sub State القديمة ثم الحالة الفرعية الجديدة ثم حالة المتابعة
            preferred = []
            if sub_old and sub_old in combined.columns:
                preferred.append(sub_old)
            elif "Sub State" in combined.columns:
                preferred.append("Sub State")
            for col in ("الحالة الفرعية الجديدة", "حالة_المتابعة", "سبب_الخطأ"):
                if col in combined.columns:
                    preferred.append(col)
            rest = [c for c in combined.columns if c not in preferred]
            combined = combined[preferred + rest]

            n_fixed = int((follow_df["حالة_المتابعة"] == "تم تغير الحالة").sum()) if len(follow_df) else 0
            n_same = int((follow_df["حالة_المتابعة"] == "لم يتم تغيير الحالة").sum()) if len(follow_df) else 0
            n_missing = int((follow_df["حالة_المتابعة"] == "غير موجود في المحفظة الجديدة").sum()) if len(follow_df) else 0
            n_new = len(new_only_df)

            a, b, c, d = st.columns(4)
            a.metric("تم تغير الحالة", f"{n_fixed:,}")
            b.metric("لم يتم تغيير الحالة", f"{n_same:,}")
            c.metric("أخطاء جديدة", f"{n_new:,}")
            d.metric("غير موجود بالمحفظة", f"{n_missing:,}")

            st.markdown("##### نتيجة المتابعة")
            st.dataframe(combined, use_container_width=True)

            sheets = {"متابعة أخطاء الحالات": combined}
            if len(follow_df):
                sheets["الحالات القديمة"] = follow_df
            if len(new_only_df):
                sheets["أخطاء جديدة"] = new_only_df
            _ops_to_excel_download(
                sheets,
                f"EJADA_{company}_case_errors_followup_{datetime.today().strftime('%Y-%m-%d')}.xlsx",
            )

# ============================================================
# ⑧ التوزيع
# ============================================================
elif mode.startswith("⑧"):
    st.subheader("⑧ التوزيع")
    st.caption("توزيع الحالات على المحصلين")
    company = st.radio("الشركة", ["سايكو", "ولاء"], horizontal=True, key="dist_company")
    up = st.file_uploader("ملف المحفظة", type=["xlsx", "xls"], key="dist_file")
    SALES_CANDS = ["Sales Person", "SalesPerson", "Collector", "المحصل", "المحصّل", "Create By"]
    SUB_CANDS = ["Sub State", "SubState", "الحالة الفرعية"]
    NET_CANDS = ["Net Amount", "NetAmount", "صافي المبلغ"]
    if up and st.button("تحليل التوزيع", type="primary"):
        df, _ = _ops_read_excel(up)
        sales_col = _ops_find_col(df, SALES_CANDS)
        sub_col = _ops_find_col(df, SUB_CANDS)
        net_col = _ops_find_col(df, NET_CANDS)
        if not sales_col:
            st.error("محتاج عمود المحصل")
            st.write(list(df.columns))
            st.stop()
        dist = df.groupby(sales_col).size().reset_index(name="عدد الحالات")
        if net_col:
            amounts = pd.to_numeric(df[net_col], errors="coerce").fillna(0)
            amt = df.assign(_a=amounts).groupby(sales_col)["_a"].sum().reset_index(name="إجمالي المديونية")
            dist = dist.merge(amt, on=sales_col, how="left")
        if sub_col:
            pivot = pd.crosstab(df[sales_col], df[sub_col].astype(str).str.strip())
            dist = dist.merge(pivot.reset_index(), on=sales_col, how="left")
        dist = dist.sort_values("عدد الحالات", ascending=False)
        st.metric("إجمالي الحالات", f"{int(dist['عدد الحالات'].sum()):,}")
        st.metric("عدد المحصلين", f"{len(dist):,}")
        st.dataframe(dist, use_container_width=True)
        fig = px.bar(dist.head(20), x=sales_col, y="عدد الحالات", color_discrete_sequence=[PRIMARY])
        fig.update_layout(title="أعلى 20 محصل", height=400)
        st.plotly_chart(fig, use_container_width=True)
        _ops_to_excel_download({"التوزيع": dist}, f"EJADA_{company}_distribution_{datetime.today().strftime('%Y-%m-%d')}.xlsx")
