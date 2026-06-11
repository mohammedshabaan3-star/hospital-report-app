"""وحدة تقارير Snapshot المتقدمة (تحميل كسول + عمودا النسب + كاش الملفات).

المبادئ:
- لا تُقرأ أي ملفات Excel ولا تُنفّذ أي عمليات تجميع/دمج قبل الضغط على زر "🔍 بدء التحليل".
- أدوات التصفية (التاريخ، المحافظة، المستشفى، التخصص) تُعرض فوراً عند الفتح
  وتُجلب خياراتها من قاعدة البيانات (خفيف) دون قراءة Excel.
- الملفات المقروءة تُخزَّن في st.session_state لتفادي إعادة القراءة في كل rerun.
- نتائج التحليل تُخزَّن في st.session_state["snapshot_analysis_result"].
- يُضاف عمودا النسب بنفس منطق add_report_ratio_columns في app.py (منسوخ محلياً
  لتجنّب استيراد app.py الذي ينفّذ كود Streamlit عند الاستيراد).
"""

import os
import sqlite3
import calendar

import pandas as pd
import streamlit as st

try:
    import plotly.express as px
    PLOTLY_AVAILABLE = True
except Exception:  # pragma: no cover - plotly متاح في بيئة التشغيل
    PLOTLY_AVAILABLE = False


# نفس مسار قاعدة البيانات المستخدم في app.py
DB_PATH = os.path.join("data", "reports.db")

# الأسماء المحتملة لعمود (إجمالي الحالات) داخل ملف الإكسيل المرفوع
TOTAL_CASES_COLUMN_CANDIDATES = [
    "إجمالي الحالات", "اجمالى الحالات", "إجمالى الحالات", "اجمالي الحالات",
    "اجمالى عدد الحالات", "إجمالي عدد الحالات", "اجمالي عدد الحالات",
]
# أعمدة فئات الحالات التي يُشتق منها الإجمالي عند غياب عمود صريح للإجمالي
SNAPSHOT_CASE_COMPONENT_COLUMNS = [
    "عدد الحالات التي تمت", "عدد الحالات الجديدة", "عدد الحالات الجارية",
    "عدد حالات تم التأجيل بناء على الحالة الصحية للمريض",
]
# عمود الحالات المنفّذة المستخدم للحساب
CASES_DONE_COLUMN = "عدد الحالات التي تمت"


def _db_conn():
    """اتصال قراءة فقط بقاعدة البيانات بنفس إعدادات app.py."""
    conn = sqlite3.connect(DB_PATH, timeout=30.0, check_same_thread=False)
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def _adjusted_capacity(row):
    """الطاقة المعدّلة للفترة بنفس منطق app.add_report_ratio_columns._adjusted_capacity."""
    cap = pd.to_numeric(row.get("capacity", 0), errors="coerce")
    if not cap or cap == 0 or pd.isna(cap):
        return 0.0
    period_type = str(row.get("period_type", "شهري")).strip()
    if period_type == "شهري":
        return float(cap)
    try:
        d_from = pd.to_datetime(row["date_from"])
        d_to = pd.to_datetime(row["date_to"])
        days_in_month = calendar.monthrange(d_from.year, d_from.month)[1]
        days_in_period = max((d_to - d_from).days + 1, 1)
        return float(cap) / days_in_month * days_in_period
    except Exception:
        return float(cap)


def add_report_ratio_columns(df, totals):
    """إضافة عمودي النسب المئوية للعرض فقط (نسخة مطابقة لمنطق app.py).

    - (نسبة التنفيذ من الطاقة الاستيعابية) = cases / الطاقة المعدّلة للفترة * 100
    - (نسبة التنفيذ من إجمالي الحالات) = cases / إجمالي_الحالات * 100
    لا تنتج NaN ولا Infinity ولا قسمة على صفر.
    """
    if df.empty:
        df["نسبة التنفيذ من الطاقة الاستيعابية"] = pd.Series(dtype=float)
        df["نسبة التنفيذ من إجمالي الحالات"] = pd.Series(dtype=float)
        return df

    adj_cap = df.apply(_adjusted_capacity, axis=1)

    cap_ratio = df["cases"].astype(float) / adj_cap.replace(0, float("nan"))
    cap_ratio = cap_ratio.replace([float("inf"), -float("inf")], float("nan")).fillna(0) * 100
    df["نسبة التنفيذ من الطاقة الاستيعابية"] = cap_ratio.round(2)

    total_series = df.apply(
        lambda row: totals.get(
            (str(row.get("hospital", "")).strip(), str(row.get("procedure", "")).strip()), 0
        ) or 0,
        axis=1
    ).astype(float)
    total_ratio = df["cases"].astype(float) / total_series.replace(0, float("nan"))
    total_ratio = total_ratio.replace([float("inf"), -float("inf")], float("nan")).fillna(0) * 100
    df["نسبة التنفيذ من إجمالي الحالات"] = total_ratio.round(2)

    return df


@st.cache_data(ttl=300)
def _load_snapshots_meta():
    """جلب بيانات الملفات الوصفية من DB فقط (خفيف — لا يقرأ Excel)."""
    try:
        with _db_conn() as c:
            return pd.read_sql(
                "SELECT id, filename, report_date, upload_date, file_path "
                "FROM snapshots ORDER BY report_date ASC",
                c,
            )
    except Exception:
        return pd.DataFrame(columns=["id", "filename", "report_date", "upload_date", "file_path"])


@st.cache_data(ttl=300)
def _load_capacity_lookup():
    """خريطة (المستشفى، التخصص) -> الطاقة الاستيعابية من جدول التخصصات (DB فقط)."""
    lookup = {}
    try:
        with _db_conn() as c:
            df = pd.read_sql(
                "SELECT u.hospital, s.procedure, s.capacity "
                "FROM specialties s JOIN users u ON s.username = u.username",
                c,
            )
        for _, r in df.iterrows():
            key = (str(r["hospital"]).strip(), str(r["procedure"]).strip())
            lookup[key] = pd.to_numeric(r["capacity"], errors="coerce")
    except Exception:
        return {}
    return lookup


@st.cache_data(ttl=300)
def _load_filter_options():
    """خيارات الفلاتر (المحافظات، المستشفيات، التخصصات) من DB فقط."""
    gov_opts, hosp_opts, proc_opts = [], [], []
    try:
        with _db_conn() as c:
            gov_opts = pd.read_sql(
                "SELECT DISTINCT governorate FROM users "
                "WHERE governorate IS NOT NULL AND governorate <> '' ORDER BY governorate", c
            )["governorate"].tolist()
            hosp_opts = pd.read_sql(
                "SELECT DISTINCT hospital FROM users "
                "WHERE hospital IS NOT NULL AND hospital <> '' ORDER BY hospital", c
            )["hospital"].tolist()
            proc_opts = pd.read_sql(
                "SELECT DISTINCT procedure FROM specialties "
                "WHERE procedure IS NOT NULL AND procedure <> '' ORDER BY procedure", c
            )["procedure"].tolist()
    except Exception:
        pass
    return gov_opts, hosp_opts, proc_opts


def _compute_totals_from_latest(files_cache, ordered_dates):
    """إجمالي الحالات (totals) من أحدث ملف ضمن الفترة، مفهرسة حسب (المستشفى، التخصص)."""
    totals = {}
    if not ordered_dates:
        return totals
    latest_df = files_cache.get(ordered_dates[-1])
    if latest_df is None or latest_df.empty:
        return totals
    dfx = latest_df.copy()
    if "المستشفى" not in dfx.columns or "تصنيف الاجراء" not in dfx.columns:
        return totals
    total_col = next((col for col in TOTAL_CASES_COLUMN_CANDIDATES if col in dfx.columns), None)
    if total_col:
        dfx["__total__"] = pd.to_numeric(dfx[total_col], errors="coerce").fillna(0)
    else:
        comp = [col for col in SNAPSHOT_CASE_COMPONENT_COLUMNS if col in dfx.columns]
        if not comp:
            return totals
        dfx["__total__"] = sum(pd.to_numeric(dfx[col], errors="coerce").fillna(0) for col in comp)
    dfx["المستشفى"] = dfx["المستشفى"].astype(str).str.strip()
    dfx["تصنيف الاجراء"] = dfx["تصنيف الاجراء"].astype(str).str.strip()
    grouped = dfx.groupby(["المستشفى", "تصنيف الاجراء"])["__total__"].sum()
    return {key: float(val) for key, val in grouped.items()}


def _build_analysis(files_cache, filtered_files, start_date, end_date,
                    gov_filter, hosp_filter, proc_filter, capacity_lookup):
    """دمج الملفات وحساب جدول النتائج مع عمودي النسب (يعمل على الكاش بلا قراءة جديدة)."""
    frames = []
    for file in filtered_files:
        raw = files_cache.get(file["report_date"])
        if raw is None or raw.empty:
            continue
        if "المستشفى" not in raw.columns or "تصنيف الاجراء" not in raw.columns:
            continue
        if CASES_DONE_COLUMN not in raw.columns:
            continue
        part = pd.DataFrame({
            "المحافظه": raw["المحافظه"].astype(str).str.strip() if "المحافظه" in raw.columns else "",
            "المستشفى": raw["المستشفى"].astype(str).str.strip(),
            "تصنيف الاجراء": raw["تصنيف الاجراء"].astype(str).str.strip(),
            "cases": pd.to_numeric(raw[CASES_DONE_COLUMN], errors="coerce").fillna(0),
        })
        frames.append(part)

    if not frames:
        return pd.DataFrame()

    df_all = pd.concat(frames, ignore_index=True)

    if gov_filter:
        df_all = df_all[df_all["المحافظه"].isin(gov_filter)]
    if hosp_filter:
        df_all = df_all[df_all["المستشفى"].isin(hosp_filter)]
    if proc_filter:
        df_all = df_all[df_all["تصنيف الاجراء"].isin(proc_filter)]

    if df_all.empty:
        return pd.DataFrame()

    grouped = (
        df_all.groupby(["المحافظه", "المستشفى", "تصنيف الاجراء"], as_index=False)["cases"].sum()
    )

    grouped["hospital"] = grouped["المستشفى"]
    grouped["procedure"] = grouped["تصنيف الاجراء"]
    grouped["capacity"] = grouped.apply(
        lambda r: capacity_lookup.get((str(r["hospital"]).strip(), str(r["procedure"]).strip()), 0),
        axis=1,
    )
    # فترة مخصصة: تُستخدم لحساب الطاقة المعدّلة بنفس منطق app.py
    grouped["period_type"] = "مخصص"
    grouped["date_from"] = str(start_date)
    grouped["date_to"] = str(end_date)

    ordered_dates = [f["report_date"] for f in filtered_files]
    totals = _compute_totals_from_latest(files_cache, ordered_dates)

    grouped = add_report_ratio_columns(grouped, totals)
    return grouped


def display_snapshot_analysis():
    """الشاشة الرئيسية: تحميل كسول مع عمودي النسب وكاش الملفات."""
    st.title("📊 تقارير Snapshot المتقدمة")

    snapshots = _load_snapshots_meta()
    if snapshots.empty:
        st.info("❌ لا توجد ملفات snapshot مرفوعة بعد.")
        return

    snapshots = snapshots.copy()
    snapshots["report_date"] = pd.to_datetime(snapshots["report_date"]).dt.date
    min_d, max_d = min(snapshots["report_date"]), max(snapshots["report_date"])

    gov_opts, hosp_opts, proc_opts = _load_filter_options()

    # --- أدوات التصفية تُعرض فوراً (دون قراءة أي ملف Excel) ---
    st.markdown("### 🔎 أدوات التصفية")
    c1, c2 = st.columns(2)
    with c1:
        start_date = st.date_input(
            "📅 من تاريخ", value=min_d, min_value=min_d, max_value=max_d, key="snap_start"
        )
        gov_filter = st.multiselect("🏙️ المحافظه", gov_opts, key="snap_gov")
        hosp_filter = st.multiselect("🏥 المستشفى", hosp_opts, key="snap_hosp")
    with c2:
        end_date = st.date_input(
            "📅 إلى تاريخ", value=max_d, min_value=min_d, max_value=max_d, key="snap_end"
        )
        proc_filter = st.multiselect("🛠️ التخصص", proc_opts, key="snap_proc")

    if start_date > end_date:
        st.error("❌ تاريخ البداية لا يمكن أن يكون بعد تاريخ النهاية.")
        return

    filtered_files = [
        {"filename": r["filename"], "report_date": r["report_date"], "file_path": r["file_path"]}
        for _, r in snapshots.iterrows()
        if start_date <= r["report_date"] <= end_date
    ]
    st.info(f"📁 عدد الملفات في الفترة: {len(filtered_files)}")
    if not filtered_files:
        st.warning("⚠️ لا توجد ملفات في الفترة المحددة.")
        return

    # --- زر بدء التحليل: لا قراءة ملفات ولا حسابات قبل الضغط ---
    if st.button("🔍 بدء التحليل", key="btn_snapshot_analysis"):
        cache_key = f"snap_files_{start_date}_{end_date}_{len(filtered_files)}"
        if cache_key not in st.session_state:
            with st.spinner("⏳ جاري قراءة الملفات..."):
                files_cache = {}
                for file in filtered_files:
                    try:
                        files_cache[file["report_date"]] = pd.read_excel(
                            file["file_path"], engine="openpyxl", dtype=str, na_filter=False
                        )
                    except Exception:
                        files_cache[file["report_date"]] = None
                st.session_state[cache_key] = files_cache
        files_cache = st.session_state[cache_key]

        with st.spinner("⏳ جاري حساب المؤشرات..."):
            capacity_lookup = _load_capacity_lookup()
            result_df = _build_analysis(
                files_cache, filtered_files, start_date, end_date,
                gov_filter, hosp_filter, proc_filter, capacity_lookup,
            )
        st.session_state["snapshot_analysis_result"] = {
            "df": result_df,
            "start_date": start_date,
            "end_date": end_date,
        }

    if "snapshot_analysis_result" not in st.session_state:
        st.info("ℹ️ اختر الفلاتر المطلوبة ثم اضغط 'بدء التحليل'.")
        return

    result = st.session_state["snapshot_analysis_result"]
    df = result["df"]
    if df is None or df.empty:
        st.warning("⚠️ لا توجد بيانات مطابقة للفلاتر المحددة.")
        return

    display_cols = [
        "المحافظه", "المستشفى", "تصنيف الاجراء", "cases", "capacity",
        "نسبة التنفيذ من الطاقة الاستيعابية", "نسبة التنفيذ من إجمالي الحالات",
    ]
    display_cols = [c for c in display_cols if c in df.columns]
    df_show = df[display_cols].rename(columns={"cases": "عدد الحالات", "capacity": "الطاقة الاستيعابية"})

    st.markdown("### 📋 جدول النتائج")
    st.info(f"📊 عدد الصفوف: {len(df_show):,}")
    st.dataframe(df_show, use_container_width=True)

    # --- مؤشرات سريعة ---
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"### ✅ **إجمالي الحالات المنفّذة: {int(df['cases'].sum()):,}**")
    with col2:
        avg_cap = df["نسبة التنفيذ من الطاقة الاستيعابية"].replace(0, pd.NA).dropna()
        avg_val = float(avg_cap.mean()) if not avg_cap.empty else 0.0
        st.markdown(f"### 📈 **متوسط نسبة الطاقة: {avg_val:.2f}%**")

    # --- رسوم بيانية ---
    if PLOTLY_AVAILABLE:
        st.markdown("### 📊 الرسوم البيانية")
        top_hosp = (
            df.groupby("المستشفى", as_index=False)["cases"].sum()
            .sort_values("cases", ascending=False).head(15)
        )
        fig_hosp = px.bar(
            top_hosp, x="المستشفى", y="cases",
            title="🏥 أعلى المستشفيات تنفيذاً للحالات",
            labels={"المستشفى": "المستشفى", "cases": "عدد الحالات"},
        )
        st.plotly_chart(fig_hosp, use_container_width=True)

        by_proc = df.groupby("تصنيف الاجراء", as_index=False)["cases"].sum()
        fig_proc = px.bar(
            by_proc, x="تصنيف الاجراء", y="cases",
            title="🛠️ توزيع الحالات حسب التخصص",
            labels={"تصنيف الاجراء": "التخصص", "cases": "عدد الحالات"},
        )
        st.plotly_chart(fig_proc, use_container_width=True)

    # --- تنزيل النتيجة ---
    from io import BytesIO
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_show.to_excel(writer, index=False, sheet_name="تقارير Snapshot")
    output.seek(0)
    st.download_button(
        "⬇️ تنزيل التقرير",
        output,
        f"تقارير_snapshot_{result['start_date']}_الى_{result['end_date']}.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_snapshot_report",
    )
