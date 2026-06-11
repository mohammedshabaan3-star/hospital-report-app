"""نظام تقارير المستشفيات - التطبيق الرئيسي

هذا التطبيق يوفر نظام شامل لإدارة تقارير المستشفيات مع إمكانيات:
- إدارة المستخدمين والمستشفيات
- إنشاء التقارير اليومية والدورية
- تحليل البيانات ومقارنة الملفات
- تصدير التقارير بصيغ مختلفة
"""

# استيراد المكتبات الأساسية
import streamlit as st
import sqlite3
import pandas as pd
import os
from datetime import date
from pathlib import Path
from io import BytesIO
import shutil
import warnings
import bcrypt
import calendar

# تعطيل التحذيرات غير الضرورية
warnings.filterwarnings('ignore', category=UserWarning)
warnings.filterwarnings('ignore', category=FutureWarning)

# مكتبات الرسوم البيانية
import plotly.express as px
import plotly.graph_objects as go

# مكتبات إضافية
import pdfkit
from sklearn.ensemble import IsolationForest

# استيراد دوال التحسين
try:
    from fix_advanced_analysis import (
        load_excel_files_optimized,
        calculate_daily_increase_optimized,
        calculate_waiting_optimized
    )
    OPTIMIZED_ANALYSIS = True
except ImportError:
    OPTIMIZED_ANALYSIS = False

# استيراد دوال إصلاح التكرارات
try:
    from fix_duplicate_labels import (
        fix_duplicate_index, safe_merge, safe_concat,
        clean_dataframe_for_merge, fix_excel_merge_duplicates
    )
    DUPLICATE_FIX_AVAILABLE = True
except ImportError:
    DUPLICATE_FIX_AVAILABLE = False
    def safe_merge(left, right, **kwargs):
        return pd.merge(left, right, **kwargs)
    def safe_concat(dfs, **kwargs):
        return pd.concat(dfs, **kwargs)
    def fix_duplicate_index(df):
        return df.reset_index(drop=True)
    def clean_dataframe_for_merge(df, key_columns=None):
        return df.drop_duplicates().reset_index(drop=True)

# استيراد الوحدات الجديدة
try:
    from modules.database import init_database, get_db_connection, get_user_data, get_hospitals_stats
    from modules.ui_components import (
        show_header, show_stats_cards, show_progress_bar, 
        create_enhanced_chart, show_data_table, create_download_section,
        show_success_message, show_error_message, show_warning_message
    )
    from modules.analytics import (
        get_hospital_performance_metrics, get_trend_analysis,
        create_performance_dashboard, create_trend_analysis_dashboard,
        create_comparative_report, generate_insights
    )
    from modules.reports import (
        display_hospital_comparison_report, display_needs_and_issues_report,
        display_capacity_analysis_report
    )
    MODULES_AVAILABLE = True
except ImportError:
    MODULES_AVAILABLE = False
    # دوال بديلة للتوافق مع الإصدار القديم
    def show_header():
        st.markdown("<h2 style='color:#1a237e;'>🏥 نظام تقارير المستشفيات</h2>", unsafe_allow_html=True)
    def show_success_message(msg):
        st.success(msg)
    def show_error_message(msg):
        st.error(msg)
    def show_warning_message(msg):
        st.warning(msg)



# استيراد وحدة تقارير Snapshot
try:
    from modules.snapshot_reports import display_snapshot_analysis
    SNAPSHOT_REPORTS_AVAILABLE = True
except ImportError:
    SNAPSHOT_REPORTS_AVAILABLE = False

# تحسينات الأداء (للتوافق مع الإصدار القديم)
try:
    from performance_improvements import (
        optimize_dataframe_display, paginate_dataframe,
        initialize_performance_improvements
    )
except ImportError:
    def optimize_dataframe_display(df, max_rows=1000):
        return df
    def paginate_dataframe(df, page_size=50):
        return df
    def initialize_performance_improvements():
        pass

# ==================== إعدادات التطبيق ====================

# مسارات الملفات وقواعد البيانات
DB_PATH = "data/reports.db"
EXCEL_PATH = "data.xlsx"
UPLOAD_DIR = "uploads/snapshots"

# إعدادات PDF
WKHTMLTOPDF_PATH = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
try:
    config = pdfkit.configuration(wkhtmltopdf=WKHTMLTOPDF_PATH)
except Exception:
    config = None

# إعدادات صفحة Streamlit
st.set_page_config(
    layout="wide",
    page_title="نظام تقارير المستشفيات",
    page_icon="🏥",
    initial_sidebar_state="collapsed"
)

# تحسين أداء pandas
pd.options.mode.chained_assignment = None  # تعطيل تحذيرات SettingWithCopyWarning

# إنشاء المجلدات اللازمة
os.makedirs("data", exist_ok=True)
os.makedirs(UPLOAD_DIR, exist_ok=True)

# --- تخصيص CSS للواجهة ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap');
    html, body, [class*="css"] {
        font-family: 'Cairo', sans-serif;
    }
    .stApp {
        background-color: #f8f9fa;
    }
    .css-1d391kg {
        background-color: #1a237e !important;
        color: white;
    }
    h1, h2, h3, h4 {
        color: #1a237e;
    }
    .stButton>button {
        background-color: #1a237e;
        color: white;
        border: none;
        border-radius: 8px;
        padding: 10px 20px;
        font-weight: 600;
    }
    .stButton>button:hover {
        background-color: #3949ab;
    }
    .stSelectbox, .stTextInput, .stDateInput, .stMultiselect {
        margin-bottom: 10px;
    }
    .stDataFrame {
        border: 1px solid #ddd;
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
    }
</style>
""", unsafe_allow_html=True)

# عرض رأس التطبيق المحسن
show_header()

# ==================== وظائف قاعدة البيانات ====================

def backup_database():
    """إنشاء نسخة احتياطية من قاعدة البيانات
    
    Returns:
        bool: True إذا تم إنشاء النسخة بنجاح
    """
    backup_path = "data/reports_backup.db"
    try:
        if os.path.exists(DB_PATH):
            shutil.copy(DB_PATH, backup_path)
            return True
    except Exception as e:
        st.error(f"خطأ في إنشاء النسخة الاحتياطية: {e}")
    return False

@st.cache_resource
def ensure_db():
    """إنشاء وتهيئة قاعدة البيانات مع جميع الجداول اللازمة"""
    try:
        with sqlite3.connect(DB_PATH) as conn:
            cur = conn.cursor()
            cur.execute("PRAGMA foreign_keys = ON")
            
            # جدول المستخدمين
            cur.execute("""CREATE TABLE IF NOT EXISTS users (
                username TEXT PRIMARY KEY,
                password TEXT NOT NULL,
                hospital TEXT,
                governorate TEXT,
                hospital_type TEXT,
                role TEXT CHECK(role IN ('admin','user')) NOT NULL,
                permissions TEXT DEFAULT ''
            )""")
            
            # جدول التخصصات
            cur.execute("""CREATE TABLE IF NOT EXISTS specialties (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT,
                procedure TEXT,
                capacity REAL,
                FOREIGN KEY(username) REFERENCES users(username)
            )""")
            
            # جدول التقارير
            cur.execute("""CREATE TABLE IF NOT EXISTS reports (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT,
                period_type TEXT,
                date_from TEXT,
                date_to TEXT,
                procedure TEXT,
                capacity REAL,
                cases INTEGER,
                notes TEXT,
                pdf TEXT,
                FOREIGN KEY(username) REFERENCES users(username)
            )""")
            
            # جدول لقطات الشاشة
            cur.execute("""CREATE TABLE IF NOT EXISTS snapshots (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                filename TEXT NOT NULL,
                report_date TEXT NOT NULL,
                upload_date TEXT NOT NULL,
                file_path TEXT NOT NULL
            )""")
            
            # إنشاء الفهارس لتحسين الأداء
            cur.execute("CREATE INDEX IF NOT EXISTS idx_reports_username ON reports(username)")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_reports_dates ON reports(date_from, date_to)")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_reports_procedure ON reports(procedure)")
            cur.execute("CREATE INDEX IF NOT EXISTS idx_specialties_username ON specialties(username)")
            # قيود التفرد (محمية ضد البيانات المكررة الموجودة مسبقاً)
            for unique_idx in (
                "CREATE UNIQUE INDEX IF NOT EXISTS idx_reports_unique ON reports(username, period_type, date_from, date_to, procedure)",
                "CREATE UNIQUE INDEX IF NOT EXISTS idx_specialties_unique ON specialties(username, procedure)",
            ):
                try:
                    cur.execute(unique_idx)
                except sqlite3.IntegrityError:
                    pass
            
            # التحقق من وجود عمود hospital_type
            cur.execute("PRAGMA table_info(users)")
            columns = [col[1] for col in cur.fetchall()]
            if "hospital_type" not in columns:
                cur.execute("ALTER TABLE users ADD COLUMN hospital_type TEXT DEFAULT 'عام'")
            
            # إنشاء مستخدم إداري افتراضي
            cur.execute("SELECT COUNT(*) FROM users WHERE username='admin'")
            if cur.fetchone()[0] == 0:
                cur.execute("""
                    INSERT INTO users 
                    (username, password, hospital, governorate, hospital_type, role, permissions) 
                    VALUES ('admin', ?, 'وزارة الصحة', 'القاهرة', 'عام', 'admin', 'all')
                """, (hash_password("admin123"),))
            else:
                cur.execute("""
                    UPDATE users SET hospital_type = 'عام' 
                    WHERE username = 'admin' AND hospital_type IS NULL
                """)
            
            conn.commit()
            
    except Exception as e:
        st.error(f"خطأ في إعداد قاعدة البيانات: {e}")
        raise

# ---------- استيراد البيانات من Excel ----------
@st.cache_data(ttl=300)
def import_excel_data():
    if not os.path.exists(EXCEL_PATH):
        st.warning("⚠️ لم يتم العثور على ملف data.xlsx")
        return
    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name=0, skiprows=2, usecols=[1, 2, 3, 4, 5])
        df.columns = ["governorate", "hospital_type", "hospital", "procedure", "capacity"]
        df.dropna(subset=["hospital", "procedure", "capacity"], inplace=True)
        df["capacity"] = pd.to_numeric(df["capacity"], errors="coerce").fillna(0).astype(float)
        df["hospital_type"] = df["hospital_type"].fillna("عام").astype(str).str.strip()
        df["hospital"] = df["hospital"].astype(str).str.strip()
        df["governorate"] = df["governorate"].astype(str).str.strip()
        df["procedure"] = df["procedure"].astype(str).str.strip()
    except Exception as e:
        st.error(f"❌ خطأ في قراءة ملف Excel: {e}")
        return

    with sqlite3.connect(DB_PATH) as conn:
        c = conn.cursor()
        hospitals = df[["hospital", "governorate", "hospital_type"]].drop_duplicates()
        for _, row in hospitals.iterrows():
            username = row["hospital"].strip().replace(" ", "_").lower()
            c.execute("""
                INSERT OR REPLACE INTO users (username, password, hospital, governorate, hospital_type, role, permissions)
                VALUES (?, ?, ?, ?, ?, 'user', '')
            """, (username, hash_password("1234"), row["hospital"].strip(), row["governorate"].strip(), row["hospital_type"].strip()))
        for _, row in df.iterrows():
            username = row["hospital"].strip().replace(" ", "_").lower()
            procedure = str(row["procedure"]).strip()
            capacity = float(row["capacity"])
            c.execute("INSERT OR IGNORE INTO specialties (username, procedure, capacity) VALUES (?, ?, ?)",
                      (username, procedure, capacity))
        conn.commit()
    st.success("✅ تم استيراد البيانات من data.xlsx بنجاح، مع الحفاظ على تصنيفات المستشفيات.")
    st.cache_data.clear()

# ---------- تصدير البيانات إلى Excel ----------
@st.cache_data(ttl=300)
def export_to_excel():
    with db_conn() as c:
        df = pd.read_sql("""
            SELECT u.hospital, u.governorate, u.hospital_type, s.procedure, s.capacity
            FROM specialties s
            JOIN users u ON s.username = u.username
        """, c)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="البيانات")
    output.seek(0)
    return output

# ---------- الاتصال بقاعدة البيانات ----------
def db_conn():
    if MODULES_AVAILABLE:
        return get_db_connection()
    else:
        conn = sqlite3.connect(DB_PATH, timeout=30.0, check_same_thread=False)
        conn.execute("PRAGMA foreign_keys = ON")
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA synchronous=NORMAL")
        conn.execute("PRAGMA cache_size=10000")
        return conn


# ---------- دوال كلمات المرور ----------
def hash_password(plain: str) -> str:
    """تشفير كلمة المرور باستخدام bcrypt."""
    return bcrypt.hashpw(plain.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")


def check_password(plain: str, hashed: str, username: str = None) -> bool:
    """التحقق من كلمة المرور مع دعم كلمات المرور القديمة (نص صريح).

    إذا كان الهاش المخزّن نصاً صريحاً قديماً وتطابق، تتم ترقيته تلقائياً إلى bcrypt.
    """
    if not hashed:
        return False
    if hashed.startswith(("$2a$", "$2b$", "$2y$")):
        try:
            return bcrypt.checkpw(plain.encode("utf-8"), hashed.encode("utf-8"))
        except ValueError:
            return False
    # توافق مع كلمات المرور القديمة المخزّنة كنص صريح
    if plain == hashed:
        if username:
            try:
                with db_conn() as c:
                    c.execute("UPDATE users SET password=? WHERE username=?",
                              (hash_password(plain), username))
            except Exception:
                pass
        return True
    return False


def update_password(username: str, new_pass: str) -> bool:
    """تحديث كلمة مرور المستخدم بعد تشفيرها. ترجع False إذا كانت فارغة."""
    if not new_pass or not new_pass.strip():
        return False
    with db_conn() as c:
        c.execute("UPDATE users SET password=? WHERE username=?",
                  (hash_password(new_pass.strip()), username))
    return True


# تهيئة قاعدة البيانات مرة واحدة عند بدء التطبيق
ensure_db()

# ---------- صلاحيات الأدمن ----------
@st.cache_data(ttl=600)
def get_user_info(username):
    with db_conn() as c:
        cur = c.cursor()
        cur.execute("SELECT hospital, governorate, role FROM users WHERE username = ?", (username,))
        return cur.fetchone()

@st.cache_data(ttl=600)
def get_permissions(username):
    with db_conn() as c:
        cur = c.cursor()
        cur.execute("SELECT permissions FROM users WHERE username=?", (username,))
        row = cur.fetchone()
        return row[0].split(',') if row and row[0] else []

def has_permission(perm):
    return st.session_state.user == "admin" or perm in get_permissions(st.session_state.user)

# ---------- حالة الجلسة ----------
if "logged" not in st.session_state:
    st.session_state.logged = False
    st.session_state.user = ""
    st.session_state.pending = []

# ---------- تسجيل الدخول ----------
def login():
    st.subheader("🔐 تسجيل الدخول")
    u = st.text_input("اسم المستخدم")
    p = st.text_input("كلمة المرور", type="password")
    if st.button("دخول"):
        with db_conn() as c:
            cur = c.cursor()
            cur.execute("SELECT password FROM users WHERE username = ?", (u.strip(),))
            row = cur.fetchone()
            if row and check_password(p.strip(), row[0], u.strip()):
                st.session_state.logged = True
                st.session_state.user = u.strip()
                st.rerun()
            else:
                st.error("❌ اسم المستخدم أو كلمة المرور غير صحيحة.")

# ---------- تسجيل الخروج ----------
def logout():
    """تسجيل خروج آمن مع تنظيف كامل للجلسة"""
    # حذف جميع مفاتيح الجلسة المخصصة
    keys_to_delete = [key for key in st.session_state.keys() if key not in ['logged', 'user', 'pending']]
    for key in keys_to_delete:
        del st.session_state[key]
    
    # إعادة تعيين القيم الأساسية
    st.session_state.logged = False
    st.session_state.user = ""
    st.session_state.pending = []
    
    # مسح الكاش
    st.cache_data.clear()
    
    # إعادة التحميل
    st.rerun()

# ---------- واجهة المستخدم ----------
def user_view():
    uname = st.session_state.user
    hosp, gov, _ = get_user_info(uname)
    st.sidebar.title(f"🏥 {hosp}")
    menu = st.sidebar.radio("القائمة", ["➕ تقرير", "📋 تقاريري", "🔐 الحساب", "🚪 خروج"])

    if menu == "➕ تقرير":
        st.title("📝 تقرير جديد")
        period = st.radio("نوع التقرير", ["يومي", "أسبوعي", "شهري"])
        dfrom = dto = date.today()
        if period != "يومي":
            dfrom = st.date_input("من تاريخ", value=date.today())
            dto = st.date_input("إلى تاريخ", value=date.today())
        else:
            dfrom = dto = st.date_input("📆 التاريخ", value=date.today())

        # تخزين التخصصات في حالة الجلسة
        if f"specialties_{uname}" not in st.session_state:
            with db_conn() as c:
                df = pd.read_sql("SELECT procedure, capacity FROM specialties WHERE username = ?", c, params=[uname])
            st.session_state[f"specialties_{uname}"] = df
        else:
            df = st.session_state[f"specialties_{uname}"]
            
        if df.empty:
            st.warning("⚠️ لا توجد تخصصات مسجلة.")
            return

        # اختيار طريقة الإدخال: دفعة واحدة لكل التخصصات (جديد) أو تخصص واحد (الطريقة القديمة كاحتياطي)
        entry_mode = st.radio(
            "🧾 طريقة الإدخال",
            ["📋 إدخال جميع التخصصات دفعة واحدة", "➕ إدخال تخصص واحد (الطريقة القديمة)"]
        )

        if entry_mode == "📋 إدخال جميع التخصصات دفعة واحدة":
            st.subheader("📋 أدخل عدد الحالات لكل تخصص ثم احفظ دفعة واحدة")
            with st.form(f"batch_report_form_{uname}"):
                batch_inputs = {}
                for _, row in df.iterrows():
                    proc_name = row["procedure"]
                    cap_val = row["capacity"]
                    batch_inputs[proc_name] = st.number_input(
                        f"🔧 {proc_name} — ⚡ السعة الاستيعابية: {cap_val}",
                        min_value=0, step=1, key=f"batch_{uname}_{proc_name}"
                    )
                submitted = st.form_submit_button("📥 حفظ التقرير لجميع التخصصات")
            if submitted:
                saved, skipped = 0, 0
                with db_conn() as c:
                    for _, row in df.iterrows():
                        proc_name = row["procedure"]
                        cap_val = row["capacity"]
                        cases_val = batch_inputs.get(proc_name, 0)
                        try:
                            c.execute("""INSERT INTO reports (
                                username, period_type, date_from, date_to, procedure, capacity, cases, notes, pdf
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                            (uname, period, str(dfrom), str(dto), proc_name, cap_val, cases_val, "", ""))
                            saved += 1
                        except sqlite3.IntegrityError:
                            skipped += 1
                if saved:
                    st.success(f"✅ تم حفظ {saved} تخصص دفعة واحدة.")
                if skipped:
                    st.warning(f"⚠️ تم تخطي {skipped} تخصص (يوجد تقرير مكرر لنفس الفترة مسبقًا).")
                if not saved and not skipped:
                    st.info("ℹ️ لا توجد تخصصات للحفظ.")
            return

        # ----- الطريقة القديمة (محفوظة كما هي كخيار احتياطي داخلي) -----
        proc = st.selectbox("🔧 اختر التخصص", df["procedure"].unique())
        cap = df[df["procedure"] == proc]["capacity"].values[0]
        st.markdown(f"⚡ السعة الاستيعابية: **{cap}**")
        cases = st.number_input("📌 عدد الحالات", min_value=0)
        notes = st.text_area("💬 ملاحظات / أعطال / احتياجات")
        pdf = st.file_uploader("📎 ملف PDF", type="pdf")

        if st.button("➕ إضافة مؤقتًا"):
            pdf_path = ""
            if pdf:
                file_bytes = pdf.getbuffer()
                if bytes(file_bytes[:4]) != b'%PDF':
                    st.error("❌ الملف المرفوع ليس ملف PDF صحيح.")
                    return
                path = Path(UPLOAD_DIR) / f"{uname}_{proc}_{dfrom}.pdf"
                with open(path, "wb") as f:
                    f.write(file_bytes)
                pdf_path = str(path)
            st.session_state.pending.append({
                "period": period,
                "from": str(dfrom),
                "to": str(dto),
                "procedure": proc,
                "capacity": cap,
                "cases": cases,
                "notes": notes,
                "pdf": pdf_path
            })
            st.success("✅ تم الإضافة مؤقتًا.")

        if st.session_state.pending:
            st.subheader("📋 الإدخالات المؤقتة")
            st.dataframe(pd.DataFrame(st.session_state.pending)[["procedure", "cases", "capacity"]])
            if st.button("📥 حفظ التقرير النهائي"):
                with db_conn() as c:
                    for e in st.session_state.pending:
                        c.execute("""INSERT INTO reports (
                            username, period_type, date_from, date_to, procedure, capacity, cases, notes, pdf
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                        (uname, e["period"], e["from"], e["to"], e["procedure"],
                         e["capacity"], e["cases"], e["notes"], e["pdf"]))
                st.session_state.pending = []
                st.success("✅ تم حفظ التقرير.")

    elif menu == "📋 تقاريري":
        st.title("📊 تقاريري السابقة")
        with db_conn() as c:
            df = pd.read_sql("SELECT * FROM reports WHERE username = ?", c, params=[uname])
        st.dataframe(df)

    elif menu == "🔐 الحساب":
        st.title("🔐 تعديل بيانات الحساب")
        st.text_input("اسم المستخدم", value=uname, disabled=True)
        new_pass = st.text_input("🔑 كلمة مرور جديدة", type="password")
        if st.button("🔄 تحديث كلمة المرور"):
            if update_password(uname, new_pass):
                st.success("✅ تم تحديث كلمة المرور. الرجاء تسجيل الدخول مجددًا.")
                logout()
            else:
                st.error("⚠️ كلمة المرور لا يمكن أن تكون فارغة.")

    elif menu == "🚪 خروج":
        if st.button("✅ تأكيد تسجيل الخروج", type="primary", use_container_width=True):
            logout()

# ---------- مساعدات شاشة التقارير ----------
# استعلام الأساس لشاشة التقارير (نفس الجداول ونفس الربط الحالي)
REPORTS_BASE_QUERY = "FROM reports r JOIN users u ON r.username = u.username"

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


def get_report_filter_options():
    """جلب قوائم الفلاتر المتاحة (المستخدمون، التخصصات، المحافظات، التصنيفات) من قاعدة البيانات."""
    with db_conn() as c:
        user_opts = pd.read_sql(f"SELECT DISTINCT r.username {REPORTS_BASE_QUERY} ORDER BY r.username", c)["username"].tolist()
        proc_opts = pd.read_sql(f"SELECT DISTINCT r.procedure {REPORTS_BASE_QUERY} ORDER BY r.procedure", c)["procedure"].tolist()
        gov_opts = pd.read_sql(f"SELECT DISTINCT u.governorate {REPORTS_BASE_QUERY} WHERE u.governorate IS NOT NULL ORDER BY u.governorate", c)["governorate"].tolist()
        type_opts = pd.read_sql(f"SELECT DISTINCT u.hospital_type {REPORTS_BASE_QUERY} WHERE u.hospital_type IS NOT NULL ORDER BY u.hospital_type", c)["hospital_type"].tolist()
    return user_opts, proc_opts, gov_opts, type_opts


def build_reports_where(ufilter, pfilter, gfilter, tfilter, d1, d2):
    """بناء جملة WHERE المُعلَّمة (parameterized) لتصفية التقارير داخل SQL (نفس المنطق الحالي)."""
    clauses = []
    params = []
    if ufilter:
        clauses.append(f"r.username IN ({','.join(['?'] * len(ufilter))})")
        params.extend(ufilter)
    if pfilter:
        clauses.append(f"r.procedure IN ({','.join(['?'] * len(pfilter))})")
        params.extend(pfilter)
    if gfilter:
        clauses.append(f"u.governorate IN ({','.join(['?'] * len(gfilter))})")
        params.extend(gfilter)
    if tfilter:
        clauses.append(f"u.hospital_type IN ({','.join(['?'] * len(tfilter))})")
        params.extend(tfilter)
    if d1 and d2:
        clauses.append("r.date_from >= ? AND r.date_to <= ?")
        params.extend([str(d1), str(d2)])
    where = (" WHERE " + " AND ".join(clauses)) if clauses else ""
    return where, params


def load_reports_df(where, params):
    """تحميل التقارير المُفلترة بنفس الاستعلام والأعمدة الحالية."""
    with db_conn() as c:
        return pd.read_sql(
            f"SELECT r.*, u.hospital, u.governorate, u.hospital_type {REPORTS_BASE_QUERY}{where}",
            c, params=params
        )


@st.cache_data(ttl=300)
def get_latest_snapshot_total_cases():
    """جلب (إجمالي الحالات) من أحدث ملف إكسيل مرفوع حسب التاريخ، مفهرسة حسب (المستشفى، التخصص).

    يعتمد على عمود (إجمالي الحالات) إن وُجد صراحةً في الملف، وإلا يُشتق الإجمالي من
    مجموع أعمدة فئات الحالات الموجودة. قراءة فقط — لا يغيّر أي جدول أو معادلة قائمة.
    يُرجع dict بالشكل: {(hospital, procedure): total_cases}.
    """
    totals = {}
    try:
        with db_conn() as c:
            snap = pd.read_sql(
                "SELECT file_path FROM snapshots ORDER BY report_date DESC, id DESC LIMIT 1", c
            )
        if snap.empty:
            return totals
        file_path = snap.iloc[0]["file_path"]
        if not file_path or not os.path.exists(file_path):
            return totals
        dfx = pd.read_excel(file_path)
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
        totals = {key: float(val) for key, val in grouped.items()}
    except Exception:
        return {}
    return totals


def add_report_ratio_columns(df, totals):
    """إضافة عمودي النسب المئوية للعرض فقط دون تعديل أي معادلة قائمة.

    نسبة التنفيذ من الطاقة الاستيعابية = cases / الطاقة المعدّلة للفترة * 100
      الطاقة المعدّلة:
        شهري  → capacity (كما هي)
        يومي  → capacity / أيام_الشهر
        أسبوعي أو مخصص → capacity / أيام_الشهر * أيام_الفترة
    نسبة التنفيذ من إجمالي الحالات = cases / إجمالي_الحالات * 100
    لا تنتج NaN ولا Infinity ولا قسمة على صفر.
    """
    if df.empty:
        df["نسبة التنفيذ من الطاقة الاستيعابية"] = pd.Series(dtype=float)
        df["نسبة التنفيذ من إجمالي الحالات"] = pd.Series(dtype=float)
        return df

    def _adjusted_capacity(row):
        cap = pd.to_numeric(row.get("capacity", 0), errors="coerce")
        if not cap or cap == 0 or pd.isna(cap):
            return 0.0
        period_type = str(row.get("period_type", "شهري")).strip()
        if period_type == "شهري":
            return float(cap)
        try:
            d_from = pd.to_datetime(row["date_from"])
            d_to   = pd.to_datetime(row["date_to"])
            days_in_month  = calendar.monthrange(d_from.year, d_from.month)[1]
            days_in_period = max((d_to - d_from).days + 1, 1)
            return float(cap) / days_in_month * days_in_period
        except Exception:
            return float(cap)

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


# ---------- واجهة الأدمن ----------
def admin_view():
    st.sidebar.title("👤 Admin")
    tabs = ["🔐 الحساب", "🚪 خروج"]
    if has_permission("reports"):
        tabs.insert(0, "📊 التقارير")
    if has_permission("summary"):
        tabs.insert(5, "📈 تقرير إحصائي مجمع")
    if has_permission("hospitals"):
        tabs.insert(1, "🏥 المستشفيات")
    if has_permission("specialties"):
        tabs.insert(2, "⚙️ التخصصات")
    if st.session_state.user == "admin":
        tabs.insert(3, "🧑‍💼 صلاحيات الأدمن")
    if has_permission("reset_password"):
        tabs.insert(4, "🔑 إعادة تعيين كلمة مرور")
    if has_permission("admin"):
        tabs.insert(6, "📂 مقارنة ملفات إكسيل (Snapshot Analysis)")
    if has_permission("admin") and SNAPSHOT_REPORTS_AVAILABLE:
        tabs.insert(7, "📊 تقارير Snapshot المتقدمة")

    menu = st.sidebar.radio("القائمة", tabs)

    if menu == "📊 التقارير" and has_permission("reports"):
        st.title("📊 تقارير المستشفيات")
        # جلب قوائم الفلاتر المتاحة من قاعدة البيانات
        user_opts, proc_opts, gov_opts, type_opts = get_report_filter_options()

        ufilter = st.multiselect("📌 المستخدم", user_opts)
        pfilter = st.multiselect("🔧 التخصص", proc_opts)
        gfilter = st.multiselect("🏙️ المحافظه", gov_opts)
        tfilter = st.multiselect("🏷️ تصنيف المستشفى", type_opts)
        d1 = st.date_input("📅 من تاريخ")
        d2 = st.date_input("📅 إلى تاريخ")

        # --- زر بدء التحليل ---
        if st.button("🔍 بدء التحليل", key="btn_start_reports_analysis"):
            where, params = build_reports_where(ufilter, pfilter, gfilter, tfilter, d1, d2)
            df = load_reports_df(where, params)
            df = add_report_ratio_columns(df, get_latest_snapshot_total_cases())
            st.session_state["reports_analysis_df"] = df
            st.session_state["reports_analysis_ready"] = True

        if not st.session_state.get("reports_analysis_ready", False):
            st.info("ℹ️ اختر الفلاتر المطلوبة ثم اضغط 'بدء التحليل'.")
        else:
            df = st.session_state["reports_analysis_df"]

            # تحسين عرض البيانات
            df_display = optimize_dataframe_display(df)
            if len(df) > 100:
                df_display = paginate_dataframe(df_display)
            st.dataframe(df_display)

            # 🔍 كشف التقارير المكررة
            with st.expander("🔍 كشف التقارير المكررة"):
                with db_conn() as c:
                    df_dups = pd.read_sql("""
                        SELECT username, period_type, date_from, date_to, procedure,
                               COUNT(*) AS total_count,
                               MIN(id) AS original_id,
                               GROUP_CONCAT(id) AS all_ids
                        FROM reports
                        GROUP BY username, period_type, date_from, date_to, procedure
                        HAVING COUNT(*) > 1
                        ORDER BY total_count DESC
                    """, c)
                if df_dups.empty:
                    st.success("✅ لا توجد تقارير مكررة.")
                else:
                    st.warning(f"⚠️ تم اكتشاف {len(df_dups)} مجموعة تكرار.")
                    st.dataframe(df_dups)
                    if has_permission("delete_reports"):
                        if st.button("🗑️ حذف التكرارات (الاحتفاظ بالأقدم)"):
                            with db_conn() as c:
                                c.execute("""
                                    DELETE FROM reports
                                    WHERE id NOT IN (
                                        SELECT MIN(id)
                                        FROM reports
                                        GROUP BY username, period_type, date_from, date_to, procedure
                                    )
                                """)
                            st.success("✅ تم حذف التقارير المكررة.")
                            st.rerun()
                    else:
                        st.info("🔒 لا تملك صلاحية حذف التقارير.")

            output = BytesIO()
            df.to_excel(output, index=False, engine="openpyxl")
            output.seek(0)
            st.download_button("⬇️ تحميل Excel", output, "filtered_reports.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            st.subheader("📌 حالة إرسال التقارير")
            with db_conn() as c:
                df_users = pd.read_sql("SELECT username, hospital FROM users WHERE role='user'", c)
            sent_users = df["username"].unique()
            sent_hospitals = df_users[df_users["username"].isin(sent_users)]["hospital"].unique()
            unsent_hospitals = df_users[~df_users["username"].isin(sent_users)]["hospital"].unique()

            col1, col2 = st.columns(2)
            with col1:
                st.markdown("#### ✅ أرسلت")
                st.success(f"{len(sent_hospitals)} مستشفى")
                st.write(sent_hospitals)
            with col2:
                st.markdown("#### 🚫 لم ترسل")
                st.error(f"{len(unsent_hospitals)} مستشفى")
                st.write(unsent_hospitals)

            if has_permission("delete_reports"):
                st.subheader("🗑️ حذف تقرير محدد")
                if not df.empty:
                    selected_id = st.selectbox("اختر رقم التقرير", df["id"].sort_values(ascending=False))
                    if st.button("🗑️ حذف التقرير"):
                        with db_conn() as c:
                            c.execute("DELETE FROM reports WHERE id=?", (selected_id,))
                        st.success(f"✅ تم حذف التقرير رقم {selected_id}")
                        st.rerun()
            else:
                st.info("🔒 لا تملك صلاحية حذف التقارير.")

    # 📈 تقرير إحصائي مجمع
    elif menu == "📈 تقرير إحصائي مجمع" and has_permission("summary"):
        st.title("📈 تقرير إحصائي متقدم")

        # --- التقرير الأصلي من قاعدة البيانات ---
        if "summary_data" not in st.session_state or st.button("🔄 تحديث البيانات"):
            with db_conn() as c:
                df = pd.read_sql("""
                    SELECT r.procedure, r.cases, r.date_from, r.date_to, u.hospital, u.governorate, u.hospital_type,
                           s.capacity
                    FROM reports r 
                    JOIN users u ON r.username = u.username
                    LEFT JOIN specialties s ON r.username = s.username AND r.procedure = s.procedure
                """, c)
            st.session_state["summary_data"] = df
        else:
            df = st.session_state["summary_data"]

        selected_hosp = st.sidebar.multiselect("🏥 المستشفى", df["hospital"].dropna().unique())
        selected_gov = st.sidebar.multiselect("📍 المحافظه", df["governorate"].dropna().unique())
        selected_type = st.sidebar.multiselect("🏷️ تصنيف المستشفى", df["hospital_type"].dropna().unique())
        selected_proc = st.sidebar.multiselect("🛠️ التخصص", df["procedure"].dropna().unique())
        d1 = st.sidebar.date_input("📅 من تاريخ")
        d2 = st.sidebar.date_input("📅 إلى تاريخ")

        if selected_hosp:
            df = df[df["hospital"].isin(selected_hosp)]
        if selected_gov:
            df = df[df["governorate"].isin(selected_gov)]
        if selected_type:
            df = df[df["hospital_type"].isin(selected_type)]
        if selected_proc:
            df = df[df["procedure"].isin(selected_proc)]
        if d1 and d2:
            df = df[(df["date_from"] >= str(d1)) & (df["date_to"] <= str(d2))]

        df["نسبة_الإشغال"] = (df["cases"] / df["capacity"]).replace([float('inf'), -float('inf')], None) * 100

        def get_status(p):
            try:
                if p >= 100:
                    return "🔴 فوق الطاقة"
                elif p >= 80:
                    return "🟠 قريب من الحد"
                else:
                    return "🟢 ضمن النطاق"
            except:
                return "⚪ غير معروف"

        df["الحالة"] = df["نسبة_الإشغال"].apply(get_status)

        col1, col2, col3 = st.columns(3)
        col1.metric("📊 عدد التخصصات", df["procedure"].nunique())
        col2.metric("🏥 عدد المستشفيات", df["hospital"].nunique())
        col3.metric("📦 إجمالي الحالات", int(df["cases"].sum()))

        st.markdown("### 🏥 نسب الإشغال حسب المستشفى")
        hospital_load = df.groupby("hospital").agg({"cases": "sum", "capacity": "sum"}).reset_index()
        hospital_load["نسبة_الإشغال"] = (hospital_load["cases"] / hospital_load["capacity"]) * 100
        hospital_load["الحالة"] = hospital_load["نسبة_الإشغال"].apply(get_status)
        st.dataframe(hospital_load.sort_values(by="نسبة_الإشغال", ascending=False))

        df["period"] = pd.to_datetime(df["date_from"])
        time_trend = df.groupby(["period", "procedure"])["cases"].sum().reset_index()
        fig_time = px.line(time_trend, x="period", y="cases", color="procedure", markers=True,
                           title="📈 الاتجاه الزمني للحالات حسب التخصص")
        st.plotly_chart(fig_time, use_container_width=True)

        agg_by = st.selectbox("📊 التجميع حسب", ["procedure", "hospital", "governorate"])
        grouped = df.groupby(agg_by).agg({"cases": "sum", "capacity": "sum"}).reset_index()
        grouped["نسبة_الإشغال"] = (grouped["cases"] / grouped["capacity"]).replace([float('inf'), -float('inf')], None) * 100
        st.dataframe(grouped)
        st.markdown(f"### ✅ الإجمالي: {grouped['cases'].sum():,} حالة")

        fig = px.bar(grouped, x=agg_by, y="cases", color="نسبة_الإشغال", title="📊 عدد الحالات حسب " + agg_by)
        st.plotly_chart(fig, use_container_width=True)

        heatmap_data = df.pivot_table(values="cases", index="governorate", columns="procedure", aggfunc="sum").fillna(0)
        st.markdown("### 🗺️ مصفوفة التوزيع حسب المحافظه والتخصص")
        st.dataframe(heatmap_data.style.background_gradient(cmap='Oranges'))

        top_pressure = grouped.sort_values(by="نسبة_الإشغال", ascending=False).head(5)
        st.markdown("### 🔴 أعلى 5 تخصصات ضغطًا على السعة")
        st.dataframe(top_pressure)

        pie_data = df.groupby("procedure")["cases"].sum().reset_index()
        fig_pie = px.pie(pie_data, names="procedure", values="cases", title="📌 توزيع الحالات حسب التخصص")
        st.plotly_chart(fig_pie, use_container_width=True)

        st.markdown("### 📆 التغير الشهري")
        df["شهر"] = pd.to_datetime(df["date_from"]).dt.to_period("M")
        trend = df.groupby(["شهر", "procedure"])["cases"].sum().reset_index()
        trend["شهر"] = trend["شهر"].astype(str)
        fig_month = px.line(trend, x="شهر", y="cases", color="procedure", title="📈 تطور الحالات شهريًا")
        st.plotly_chart(fig_month, use_container_width=True)

        st.markdown("### 🔝 أعلى التخصصات")
        top_n = st.slider("🔢 عدد الأعلى", 1, 2, 15)
        top = df.groupby("procedure")["cases"].sum().reset_index().sort_values("cases", ascending=False).head(top_n)
        st.dataframe(top)

        st.markdown("### 🌦️ التحليل الموسمي")
        df["شهر_عددي"] = pd.to_datetime(df["date_from"]).dt.month
        seasonal = df.groupby(["شهر_عددي", "procedure"])["cases"].sum().reset_index()
        fig_season = px.line(seasonal, x="شهر_عددي", y="cases", color="procedure", title="🌦️ الاتجاه الموسمي للحالات")
        st.plotly_chart(fig_season, use_container_width=True)

        df["duration_days"] = (pd.to_datetime(df["date_to"]) - pd.to_datetime(df["date_from"])).dt.days + 1
        df["risk_index"] = df["نسبة_الإشغال"] * df["duration_days"]
        high_risk = df.sort_values("risk_index", ascending=False).head(10)
        st.markdown("### ⚠️ التخصصات الأعلى خطورة")
        st.dataframe(high_risk[["procedure", "hospital", "risk_index"]])

        all_specialties = df["procedure"].unique()
        active_specialties = df[df["cases"] > 0]["procedure"].unique()
        inactive = set(all_specialties) - set(active_specialties)
        if inactive:
            st.warning("🔍 التخصصات التي لم تسجل حالات:")
            st.write(list(inactive))

        radar_data = df.groupby("hospital").agg({"cases":"sum", "capacity":"sum"}).reset_index()
        radar_data["نسبة_الإشغال"] = radar_data["cases"] / radar_data["capacity"] * 100
        fig_radar = go.Figure()
        fig_radar.add_trace(go.Scatterpolar(r=radar_data["نسبة_الإشغال"], theta=radar_data["hospital"], fill='toself', name='نسبة الإشغال'))
        fig_radar.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0,150])), showlegend=True, title="📡 رادار الأداء حسب المستشفى")
        st.plotly_chart(fig_radar, use_container_width=True)

        gov_perf = df.groupby("governorate").agg({"cases":"sum", "capacity":"sum"}).reset_index()
        gov_perf["نسبة_الإشغال"] = gov_perf["cases"] / gov_perf["capacity"] * 100
        best = gov_perf.sort_values("نسبة_الإشغال", ascending=False).head(1)
        worst = gov_perf.sort_values("نسبة_الإشغال", ascending=True).head(1)
        st.success(f"🏆 أفضل محافظة: {best['governorate'].values[0]} بنسبة إشغال {best['نسبة_الإشغال'].values[0]:.2f}%")
        st.error(f"⚠️ أضعف محافظة: {worst['governorate'].values[0]} بنسبة إشغال {worst['نسبة_الإشغال'].values[0]:.2f}%")

        st.markdown("### 🔍 مؤشر التحميل الزمني لكل تخصص")
        df["شهر"] = pd.to_datetime(df["date_from"]).dt.to_period("M")
        load_index = df.groupby(["procedure", "شهر"])["cases"].sum().unstack().fillna(0)
        st.dataframe(load_index.style.background_gradient(cmap='Blues'))

        st.markdown("### 🚨 التخصصات الخارجة عن النطاق (Anomaly Detection)")
        anomaly_data = df.groupby("procedure")["cases"].sum().reset_index()
        model = IsolationForest(contamination=0.15, random_state=42)
        anomaly_data["is_outlier"] = model.fit_predict(anomaly_data[["cases"]])
        outliers = anomaly_data[anomaly_data["is_outlier"] == -1]
        if not outliers.empty:
            st.error("❗ التخصصات التي بها حالات غير معتادة:")
            st.dataframe(outliers[["procedure", "cases"]])
        else:
            st.success("✅ لا توجد تخصصات شاذة حاليًا.")

        st.markdown("### ⚡ مؤشر الاستجابة")
        df["days"] = (pd.to_datetime(df["date_to"]) - pd.to_datetime(df["date_from"])).dt.days + 1
        df["مؤشر_الاستجابة"] = (df["cases"] / df["capacity"]) / df["days"]
        response_index = df.groupby("procedure")["مؤشر_الاستجابة"].mean().reset_index()
        st.dataframe(response_index.sort_values(by="مؤشر_الاستجابة", ascending=False))

        st.markdown("### 🧭 التخصصات التي لم تُبلغ عن حالات")
        all_specialties = df["procedure"].unique()
        missing = df.groupby("hospital")["procedure"].unique().apply(lambda reported: list(set(all_specialties) - set(reported)))
        missing_df = missing.reset_index().rename(columns={0: "تخصصات مفقودة"})
        st.dataframe(missing_df)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="البيانات", index=False)
            grouped.to_excel(writer, sheet_name="ملخص حسب " + agg_by, index=False)
            hospital_load.to_excel(writer, sheet_name="الأداء حسب المستشفى", index=False)
            high_risk.to_excel(writer, sheet_name="الأعلى خطورة", index=False)
        output.seek(0)
        filename = f"summary_{str(d1)}_{str(d2)}.xlsx"
        st.download_button("⬇️ تحميل التقرير المفصل", output, filename)

        st.markdown("## مقارنة بين فترتين زمنيتين")
        col1, col2 = st.columns(2)
        with col1:
            from1 = st.date_input("📅 فترة أولى - من", key="f1_from")
            to1 = st.date_input("📅 فترة أولى - إلى", key="f1_to")
        with col2:
            from2 = st.date_input("📅 فترة ثانية - من", key="f2_from")
            to2 = st.date_input("📅 فترة ثانية - إلى", key="f2_to")
        df1 = df[(df["date_from"] >= str(from1)) & (df["date_to"] <= str(to1))]
        df2 = df[(df["date_from"] >= str(from2)) & (df["date_to"] <= str(to2))]
        sum1 = df1["cases"].sum()
        sum2 = df2["cases"].sum()
        diff = sum2 - sum1
        st.metric("📊 الفرق بين الفترتين", f"{diff:+,}", delta_color="inverse")

    # 🏥 إدارة المستشفيات
    elif menu == "🏥 المستشفيات" and has_permission("hospitals"):
        st.title("🏥 إدارة المستشفيات")

        # --- زر الاستيراد ---
        st.subheader("📥 استيراد من ملف Excel")
        if st.button("🔄 استيراد من data.xlsx"):
            import_excel_data()
            st.rerun()

        # --- زر التصدير ---
        st.subheader("📤 تصدير إلى Excel")
        excel_data = export_to_excel()
        st.download_button(
            "⬇️ تنزيل ملف البيانات",
            excel_data,
            "hospitals_export.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # --- إضافة مستشفى جديد ---
        st.subheader("➕ إضافة مستشفى جديد")
        col1, col2 = st.columns(2)
        with col1:
            name = st.text_input("اسم المستشفى", key="add_name")
            gov = st.text_input("المحافظه", key="add_gov")
            hosp_type = st.text_input("تصنيف المستشفى", placeholder="مثال: عام، تخصصي، جامعي ...", value="عام", key="add_type")
        with col2:
            uname = st.text_input("اسم المستخدم", value=name.strip().replace(" ", "_").lower() if name else "", key="add_username")
            pw = st.text_input("كلمة المرور", value="1234", type="password", key="add_password")
            confirm_pw = st.text_input("تأكيد كلمة المرور", type="password", key="add_confirm_pw")

        if st.button("➕ إضافة"):
            if not all([name, gov, hosp_type, uname, pw]):
                st.error("❌ يرجى تعبئة جميع الحقول.")
            elif confirm_pw != pw:
                st.error("❌ كلمة المرور وتأكيدها غير متطابقتين.")
            else:
                with db_conn() as c:
                    cur = c.cursor()
                    cur.execute("SELECT COUNT(*) FROM users WHERE username=?", (uname,))
                    if cur.fetchone()[0] > 0:
                        st.error("❌ اسم المستخدم موجود بالفعل.")
                    else:
                        cur.execute("""
                            INSERT INTO users (username, password, hospital, governorate, hospital_type, role, permissions)
                            VALUES (?, ?, ?, ?, ?, 'user', '')
                        """, (uname, hash_password(pw), name, gov, hosp_type.strip()))
                        st.success(f"✅ تم إضافة المستشفى: {name} (اسم المستخدم: {uname})")
                        st.rerun()

        # --- عرض المستشفيات ---
        with db_conn() as c:
            df_hosp = pd.read_sql("SELECT username, hospital, governorate, hospital_type FROM users WHERE role='user'", c)
            df_count = pd.read_sql("SELECT username, COUNT(*) AS specialties FROM specialties GROUP BY username", c)
            df_all = pd.merge(df_hosp, df_count, on="username", how="left").fillna(0)

        df_all = df_all.rename(columns={
            "hospital": "اسم المستشفى",
            "governorate": "المحافظة",
            "hospital_type": "تصنيف المستشفى",
            "specialties": "عدد التخصصات"
        })
        st.dataframe(df_all)

        # --- تعديل بيانات المستشفى ---
        st.subheader("✏️ تعديل بيانات مستشفى")
        if not df_hosp.empty:
            selected_hospital = st.selectbox("اختر مستشفى للتعديل", df_hosp["hospital"].unique())
            current_data = df_hosp[df_hosp["hospital"] == selected_hospital].iloc[0]
            
            with db_conn() as c:
                cur = c.cursor()
                cur.execute("SELECT username, password FROM users WHERE hospital=?", (selected_hospital,))
                user_data = cur.fetchone()
            
            col1, col2 = st.columns(2)
            with col1:
                new_name = st.text_input("اسم المستشفى", value=current_data["hospital"], key="edit_name")
                new_gov = st.text_input("المحافظة", value=current_data["governorate"], key="edit_gov")
                new_type = st.text_input("تصنيف المستشفى", value=current_data["hospital_type"], key="edit_type")
            with col2:
                new_username = st.text_input("اسم المستخدم", value=user_data[0] if user_data else "", key="edit_username")
                new_password = st.text_input("كلمة المرور الجديدة", type="password", placeholder="اترك فارغاً لعدم التغيير", key="edit_password")
                confirm_new_pw = st.text_input("تأكيد كلمة المرور", type="password", key="edit_confirm_pw")
            
            if st.button("💾 حفظ التعديلات"):
                if not all([new_name, new_gov, new_type, new_username]):
                    st.error("❌ يرجى تعبئة جميع الحقول المطلوبة.")
                elif new_password and new_password != confirm_new_pw:
                    st.error("❌ كلمة المرور وتأكيدها غير متطابقتين.")
                else:
                    with db_conn() as c:
                        cur = c.cursor()
                        # التحقق من عدم تكرار اسم المستخدم
                        if new_username != user_data[0]:
                            cur.execute("SELECT COUNT(*) FROM users WHERE username=?", (new_username,))
                            if cur.fetchone()[0] > 0:
                                st.error("❌ اسم المستخدم الجديد موجود بالفعل.")
                                return
                        
                        # تحديث بيانات المستشفى
                        if new_password:
                            cur.execute("""
                                UPDATE users SET username=?, password=?, hospital=?, governorate=?, hospital_type=?
                                WHERE hospital=?
                            """, (new_username, hash_password(new_password), new_name, new_gov, new_type, selected_hospital))
                        else:
                            cur.execute("""
                                UPDATE users SET username=?, hospital=?, governorate=?, hospital_type=?
                                WHERE hospital=?
                            """, (new_username, new_name, new_gov, new_type, selected_hospital))
                        
                        # تحديث اسم المستخدم في جدول التخصصات
                        if new_username != user_data[0]:
                            cur.execute("UPDATE specialties SET username=? WHERE username=?", (new_username, user_data[0]))
                            cur.execute("UPDATE reports SET username=? WHERE username=?", (new_username, user_data[0]))
                        
                        st.success(f"✅ تم تحديث بيانات المستشفى بنجاح")
                        st.rerun()

        # --- حذف مستشفى ---
        st.subheader("🗑️ حذف مستشفى")
        del_hosp = st.selectbox("اختر مستشفى للحذف", ["اختر مستشفى"] + df_hosp["hospital"].tolist())
        if del_hosp != "اختر مستشفى":
            code = df_hosp[df_hosp["hospital"] == del_hosp]["username"].values[0]
            st.warning(f"⚠️ سيتم حذف جميع بيانات المستشفى '{del_hosp}' نهائياً")
            if st.button("🗑️ تأكيد الحذف", type="primary"):
                with db_conn() as c:
                    c.execute("DELETE FROM reports WHERE username=?", (code,))
                    c.execute("DELETE FROM specialties WHERE username=?", (code,))
                    c.execute("DELETE FROM users WHERE username=?", (code,))
                st.success(f"✅ تم حذف المستشفى: {del_hosp}")
                st.rerun()

    # ⚙️ إدارة التخصصات
    elif menu == "⚙️ التخصصات" and has_permission("specialties"):
        st.title("⚙️ إدارة التخصصات")
        with db_conn() as c:
            df_hosp = pd.read_sql("SELECT username, hospital FROM users WHERE role='user'", c)
            all_specs = pd.read_sql("SELECT DISTINCT procedure FROM specialties", c)
        if df_hosp.empty:
            st.warning("🚫 لا توجد مستشفيات.")
            return
        hosp = st.selectbox("اختر مستشفى", df_hosp["hospital"])
        uname = df_hosp[df_hosp["hospital"] == hosp]["username"].values[0]
        with db_conn() as c:
            df_spec = pd.read_sql("SELECT id, procedure, capacity FROM specialties WHERE username = ?", c, params=[uname])
        if not df_spec.empty:
            row_map = {f"{r['procedure']} (سعة: {r['capacity']})": r for _, r in df_spec.iterrows()}
            selected = st.selectbox("🔧 تخصص للتعديل أو الحذف", list(row_map.keys()))
            r = row_map[selected]
            new_proc = st.text_input("✏️ الاسم", value=r["procedure"])
            new_cap = st.number_input("🔄 السعة", min_value=0.0, value=float(r["capacity"]))
            col1, col2 = st.columns(2)
            with col1:
                if st.button("💾 تعديل"):
                    with db_conn() as c:
                        c.execute("UPDATE specialties SET procedure=?, capacity=? WHERE id=?", (new_proc, new_cap, r["id"]))
                    st.success("✅ تم التعديل.")
                    st.rerun()
            with col2:
                if st.button("🗑️ حذف"):
                    with db_conn() as c:
                        c.execute("DELETE FROM specialties WHERE id=?", (r["id"],))
                    st.success("🗑️ تم الحذف.")
                    st.rerun()
        st.markdown("### ➕ إضافة تخصص جديد")
        new_proc_add = st.selectbox("🆕 تخصص عام", all_specs["procedure"].unique())
        new_cap_add = st.number_input("📌 السعة", min_value=0.0)
        if st.button("➕ أضف التخصص"):
            with db_conn() as c:
                c.execute("INSERT INTO specialties (username, procedure, capacity) VALUES (?, ?, ?)",
                          (uname, new_proc_add.strip(), new_cap_add))
            st.success("✅ تم الإضافة.")
            st.rerun()

    # 🧑‍💼 إدارة صلاحيات الأدمن
    elif menu == "🧑‍💼 صلاحيات الأدمن" and st.session_state.user == "admin":
        st.title("🧑‍💼 إدارة صلاحيات الأدمن")
        st.subheader("➕ إضافة أدمن جديد")
        new_admin_user = st.text_input("👤 اسم المستخدم الجديد").strip().lower().replace(" ", "_")
        new_admin_pass = st.text_input("🔑 كلمة المرور", type="password")
        new_admin_hospital = st.text_input("🏥 جهة العمل", value="وزارة الصحة")
        new_admin_gov = st.text_input("📍 المحافظه", value="القاهرة")
        perms = ["reports", "hospitals", "specialties", "reset_password", "delete_reports", "summary", "admin"]
        selected_perms = st.multiselect("✅ الصلاحيات", perms)
        if st.button("➕ إضافة الأدمن"):
            if new_admin_user and new_admin_pass:
                with db_conn() as c:
                    c.execute("INSERT OR IGNORE INTO users VALUES (?, ?, ?, ?, ?, ?, ?)",
                              (new_admin_user, hash_password(new_admin_pass), new_admin_hospital, new_admin_gov, "عام", "admin", ",".join(selected_perms)))
                st.success(f"✅ تم إضافة الأدمن '{new_admin_user}'")
                st.rerun()
            else:
                st.warning("⚠️ الرجاء إدخال اسم مستخدم وكلمة مرور.")

        st.subheader("🛠️ تعديل صلاحيات الأدمن الحاليين")
        with db_conn() as c:
            admins = pd.read_sql("SELECT username FROM users WHERE role='admin' AND username != 'admin'", c)
        selected = st.selectbox("👥 اختر الأدمن", admins["username"]) if not admins.empty else None
        all_perms = ["reports", "hospitals", "specialties", "reset_password", "delete_reports", "summary", "admin"]
        current = get_permissions(selected) if selected else []
        if selected:
            for perm in all_perms:
                state = perm in current
                new_state = st.checkbox(f"{perm} - {perm.upper()}", value=state, key=perm)
                if new_state and not state:
                    current.append(perm)
                elif not new_state and state:
                    current.remove(perm)
            if st.button("💾 تحديث الصلاحيات"):
                with db_conn() as c:
                    c.execute("UPDATE users SET permissions=? WHERE username=?", (",".join(current), selected))
                st.success("✅ تم تحديث الصلاحيات.")
                st.rerun()
            if st.button("🗑️ حذف الأدمن المحدد"):
                with db_conn() as c:
                    c.execute("DELETE FROM users WHERE username=?", (selected,))
                st.success(f"🗑️ تم حذف الأدمن '{selected}'")
                st.rerun()

    # 🔑 إعادة تعيين كلمة مرور
    elif menu == "🔑 إعادة تعيين كلمة مرور" and has_permission("reset_password"):
        st.title("🔒 إعادة تعيين كلمة مرور")
        with db_conn() as c:
            df_users = pd.read_sql("SELECT username, hospital FROM users WHERE role='user'", c)
        if not df_users.empty:
            hosp = st.selectbox("اختر مستشفى", df_users["hospital"])
            uname = df_users[df_users["hospital"] == hosp]["username"].values[0]
            new_pass = st.text_input("🔑 كلمة مرور جديدة")
            if st.button("🔄 تحديث"):
                if update_password(uname, new_pass):
                    st.success("✅ تم تحديث كلمة المرور.")
                else:
                    st.error("⚠️ كلمة المرور لا يمكن أن تكون فارغة.")

    # 🔐 الحساب
    elif menu == "🔐 الحساب":
        uname = st.session_state.user
        st.title("🔐 تعديل الحساب")
        st.text_input("اسم المستخدم", value=uname, disabled=True)
        new_pass = st.text_input("كلمة مرور جديدة", type="password")
        if st.button("تحديث"):
            if update_password(uname, new_pass):
                st.success("✅ تم التحديث.")
                logout()
            else:
                st.error("⚠️ كلمة المرور لا يمكن أن تكون فارغة.")

    # 📂 مقارنة ملفات إكسيل (Snapshot Analysis)
    elif menu == "📂 مقارنة ملفات إكسيل (Snapshot Analysis)" and has_permission("admin"):
        st.header("📂 مقارنة ملفات إكسيل (مقارنة حسب التاريخ)")

        # --- تبويبات فرعية ---
        sub_menu = st.radio("التحليل", [
            "📊 رفع ومقارنة ملفين",
            "📈 تحليل متقدم للملفات المحفوظة"
        ])

        if sub_menu == "📊 رفع ومقارنة ملفين":
            # --- رفع ملف جديد ---
            st.subheader("📤 رفع ملف تقرير")
            uploaded_file = st.file_uploader("اختر ملف إكسل", type=["xlsx"], key="upload_snapshot")
            report_date = st.date_input("📅 تاريخ التقرير", value=date.today())

            if st.button("💾 حفظ الملف") and uploaded_file is not None:
                try:
                    df_new = pd.read_excel(uploaded_file)
                    required_cols = [
                        "المحافظه", "المستشفى", "تصنيف المستشفى", "تصنيف الاجراء",
                        "عدد الحالات التي تمت", "عدد الحالات الجديدة", "عدد الحالات الجارية",
                        "عدد حالات تم التأجيل بناء على الحالة الصحية للمريض"
                    ]
                    if not all(col in df_new.columns for col in required_cols):
                        st.error("❌ الملف لا يحتوي على الأعمدة المطلوبة.")
                    else:
                        filename = f"{report_date}_{uploaded_file.name}"
                        file_path = os.path.join(UPLOAD_DIR, filename)
                        with open(file_path, "wb") as f:
                            f.write(uploaded_file.getbuffer())
                        with db_conn() as c:
                            c.execute("INSERT INTO snapshots (filename, report_date, upload_date, file_path) VALUES (?, ?, ?, ?)",
                                      (filename, str(report_date), str(date.today()), file_path))
                        st.success(f"✅ تم حفظ الملف: {filename}")
                        st.rerun()
                except Exception as e:
                    st.error(f"❌ خطأ في حفظ الملف: {e}")

            # --- جلب الملفات المحفوظة ---
            with db_conn() as c:
                snapshots = pd.read_sql("""
                    SELECT id, filename, report_date, upload_date, file_path 
                    FROM snapshots 
                    ORDER BY report_date DESC
                """, c)

            if snapshots.empty:
                st.info("❌ لا توجد ملفات محفوظة بعد.")
            else:
                snapshots["report_date"] = pd.to_datetime(snapshots["report_date"]).dt.date
                all_dates = sorted(snapshots["report_date"].unique(), reverse=True)

                st.subheader("🔍 اختر التواريخ للمقارنة")
                col1, col2 = st.columns(2)
                with col1:
                    selected_later = st.selectbox(
                        "📌 التقرير الحديث (لاحق)",
                        options=all_dates,
                        format_func=lambda d: f"{d.year}-{d.month:02d}-{d.day:02d}",
                        key="select_later"
                    )
                with col2:
                    earlier_options = [d for d in all_dates if d < selected_later]
                    if not earlier_options:
                        st.selectbox("📌 التقرير القديم (سابق)", options=[], disabled=True)
                        st.warning("⚠️ لا توجد تواريخ أقدم.")
                        selected_earlier = None
                    else:
                        selected_earlier = st.selectbox(
                            "📌 التقرير القديم (سابق)",
                            options=earlier_options,
                            format_func=lambda d: f"{d.year}-{d.month:02d}-{d.day:02d}",
                            key="select_earlier"
                        )

                # --- تنفيذ المقارنة ---
                if st.button("📊 تنفيذ المقارنة", key="run_comparison") and selected_earlier:
                    try:
                        # --- جلب الملفات ---
                        later_file = snapshots[snapshots["report_date"] == selected_later].iloc[0]
                        earlier_file = snapshots[snapshots["report_date"] == selected_earlier].iloc[0]

                        file_path_later = later_file["file_path"]
                        file_path_earlier = earlier_file["file_path"]

                        df_later = pd.read_excel(file_path_later)
                        df_earlier = pd.read_excel(file_path_earlier)

                        required_cols = [
                            "المحافظه", "المستشفى", "تصنيف المستشفى", "تصنيف الاجراء",
                            "عدد الحالات التي تمت", "عدد الحالات الجديدة", "عدد الحالات الجارية",
                            "عدد حالات تم التأجيل بناء على الحالة الصحية للمريض"
                        ]

                        if not all(col in df_later.columns for col in required_cols):
                            st.error("❌ الملف الحديث يفتقد أعمدة.")
                        elif not all(col in df_earlier.columns for col in required_cols):
                            st.error("❌ الملف القديم يفتقد أعمدة.")
                        else:
                            # --- حساب "المنتظر" للملف الحديث ---
                            df_later["المنتظر"] = df_later["عدد الحالات الجديدة"] + df_later["عدد الحالات الجارية"]
                            
                            # --- حساب "المنتظر" للملف القديم ---
                            df_earlier["المنتظر_القديم"] = df_earlier["عدد الحالات الجديدة"] + df_earlier["عدد الحالات الجارية"]

                            # --- إزالة التكرارات قبل الدمج ---
                            merge_columns = ["المحافظه", "المستشفى", "تصنيف المستشفى", "تصنيف الاجراء"]
                            
                            if DUPLICATE_FIX_AVAILABLE:
                                df_later_clean, df_earlier_clean = fix_excel_merge_duplicates(
                                    df_later[required_cols + ["المنتظر"]],
                                    df_earlier[["المحافظه", "المستشفى", "تصنيف المستشفى", "تصنيف الاجراء", "عدد الحالات التي تمت", "المنتظر_القديم"]],
                                    merge_columns
                                )
                            else:
                                df_later_clean = df_later[required_cols + ["المنتظر"]].drop_duplicates(
                                    subset=merge_columns
                                ).reset_index(drop=True)
                                
                                df_earlier_clean = df_earlier[["المحافظه", "المستشفى", "تصنيف المستشفى", "تصنيف الاجراء", "عدد الحالات التي تمت", "المنتظر_القديم"]].drop_duplicates(
                                    subset=merge_columns
                                ).reset_index(drop=True)
                            
                            # --- دمج الملفات ---
                            merged = safe_merge(
                                df_later_clean,
                                df_earlier_clean,
                                on=merge_columns,
                                how="left",
                                suffixes=("_الحديث", "_القديم")
                            )
                            merged["عدد الحالات التي تمت_القديم"] = merged["عدد الحالات التي تمت_القديم"].fillna(0)
                            merged["المنتظر_القديم"] = merged["المنتظر_القديم"].fillna(0)
                            merged["الزيادة"] = merged["عدد الحالات التي تمت_الحديث"] - merged["عدد الحالات التي تمت_القديم"]

                            # --- جلب البيانات من النظام خلال الفترة ---
                            try:
                                start_date = selected_earlier.isoformat()
                                end_date = selected_later.isoformat()

                                with db_conn() as c:
                                    df_db = pd.read_sql("""
                                        SELECT 
                                            u.governorate AS المحافظه,
                                            u.hospital AS المستشفى,
                                            r.procedure AS تصنيف_الاجراء,
                                            SUM(r.cases) AS الحالات_المسجلة
                                        FROM reports r
                                        JOIN users u ON r.username = u.username
                                        WHERE r.date_from >= ?
                                          AND r.date_to <= ?
                                        GROUP BY u.governorate, u.hospital, r.procedure
                                    """, c, params=[start_date, end_date])

                                # --- تنظيف ودمج ---
                                for col in ["المحافظه", "المستشفى", "تصنيف_الاجراء"]:
                                    df_db[col] = df_db[col].astype(str).str.strip()
                                df_db.rename(columns={"تصنيف_الاجراء": "تصنيف الاجراء"}, inplace=True)
                                
                                # إزالة التكرارات من بيانات قاعدة البيانات
                                df_db = df_db.drop_duplicates(
                                    subset=["المحافظه", "المستشفى", "تصنيف الاجراء"]
                                ).reset_index(drop=True)

                                for col in ["المحافظه", "المستشفى", "تصنيف الاجراء"]:
                                    merged[col] = merged[col].astype(str).str.strip()

                                merged = safe_merge(
                                    merged,
                                    df_db,
                                    on=["المحافظه", "المستشفى", "تصنيف الاجراء"],
                                    how="left"
                                )
                                merged["الحالات_المسجلة"] = merged["الحالات_المسجلة"].fillna(0)

                            except Exception as e:
                                st.warning(f"⚠️ تعذر جلب البيانات من النظام: {e}")
                                merged["الحالات_المسجلة"] = 0

                            # --- حفظ النتيجة ---
                            st.session_state.comparison_result = merged
                            st.success("✅ تم تنفيذ المقارنة. يمكنك الآن استخدام الفلاتر.")

                    except Exception as e:
                        st.error(f"❌ خطأ في معالجة الملفات: {e}")

                # --- عرض النتائج والفلاتر ---
                if "comparison_result" in st.session_state:
                    merged = st.session_state.comparison_result

                    st.subheader("⚙️ تصفية النتائج")
                    col1, col2 = st.columns(2)
                    with col1:
                        gov_filter = st.multiselect("📍 المحافظه", merged["المحافظه"].unique(), key="gov_filter")
                        hosp_filter = st.multiselect("🏥 المستشفى", merged["المستشفى"].unique(), key="hosp_filter")
                    with col2:
                        type_filter = st.multiselect("🏷️ تصنيف المستشفى", merged["تصنيف المستشفى"].unique(), key="type_filter")
                        proc_filter = st.multiselect("🛠️ تصنيف الاجراء", merged["تصنيف الاجراء"].unique(), key="proc_filter")

                    df_filtered = merged.copy()
                    if gov_filter:
                        df_filtered = df_filtered[df_filtered["المحافظه"].isin(gov_filter)]
                    if hosp_filter:
                        df_filtered = df_filtered[df_filtered["المستشفى"].isin(hosp_filter)]
                    if type_filter:
                        df_filtered = df_filtered[df_filtered["تصنيف المستشفى"].isin(type_filter)]
                    if proc_filter:
                        df_filtered = df_filtered[df_filtered["تصنيف الاجراء"].isin(proc_filter)]

                    st.subheader("📊 النتائج بعد المقارنة")
                    st.dataframe(df_filtered)

                    # --- تنزيل Excel ---
                    output = BytesIO()
                    df_filtered.to_excel(output, index=False, engine="openpyxl")
                    output.seek(0)
                    st.download_button(
                        "⬇️ تنزيل Excel",
                        output,
                        file_name=f"comparison_{selected_later}_vs_{selected_earlier}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_excel"
                    )

                    # --- تجميع البيانات ---
                    st.subheader("📊 تجميع البيانات")
                    group_by_options = {
                        "حسب المحافظة": "المحافظه",
                        "حسب المستشفى": "المستشفى",
                        "حسب تصنيف المستشفى": "تصنيف المستشفى",
                        "حسب التخصص": "تصنيف الاجراء",
                        "حسب المستشفى والتخصص معًا": ["المستشفى", "تصنيف الاجراء"],
                        "حسب التخصص والمستشفى معًا": ["تصنيف الاجراء", "المستشفى"]
                    }
                    selected_group = st.selectbox(
                        "اختر طريقة التجميع",
                        options=list(group_by_options.keys()),
                        key="group_select_pdf"
                    )
                    group_col = group_by_options[selected_group]

                    # فقط الأعمدة الموجودة في df_filtered
                    available_cols = df_filtered.columns.tolist()
                    value_cols = [col for col in [
                        "عدد الحالات التي تمت_الحديث",
                        "عدد الحالات التي تمت_القديم",
                        "الزيادة",
                        "الحالات_المسجلة",
                        "المنتظر",
                        "المنتظر_القديم"
                    ] if col in available_cols]

                    # --- دالة لإضافة إجماليات فرعية ---
                    def add_subtotals(df, group_cols, value_cols):
                        if not isinstance(group_cols, list):
                            group_cols = [group_cols]
                        df_grouped = df.groupby(group_cols)[value_cols].sum().round(0).astype(int)
                        if len(group_cols) != 2:
                            return df_grouped
                        outer_col, inner_col = group_cols[0], group_cols[1]
                        result_list = []
                        for outer_val in df_grouped.index.get_level_values(outer_col).unique():
                            sub_df = df_grouped.xs(outer_val)
                            for idx in sub_df.index:
                                result_list.append(sub_df.loc[idx].to_dict())
                            total_dict = sub_df.sum(axis=0).to_dict()
                            result_list.append(total_dict)
                        return pd.DataFrame(result_list)

                    if isinstance(group_col, list) and len(group_col) == 2:
                        grouped = add_subtotals(df_filtered, group_col, value_cols)
                    else:
                        grouped = df_filtered.groupby(group_col)[value_cols].sum().round(0).astype(int)
                    
                    # إضافة عمود الزيادة إذا لم يكن موجوداً
                    if "عدد الحالات التي تمت_الحديث" in grouped.columns and "عدد الحالات التي تمت_القديم" in grouped.columns:
                        if "الزيادة" not in grouped.columns:
                            grouped["الزيادة"] = grouped["عدد الحالات التي تمت_الحديث"] - grouped["عدد الحالات التي تمت_القديم"]

                    # --- إعادة التسمية ---
                    grouped = grouped.rename(columns={
                        "عدد الحالات التي تمت_الحديث": "الحالات (الحديث)",
                        "عدد الحالات التي تمت_القديم": "الحالات (القديم)", 
                        "الزيادة": "الزيادة",
                        "الحالات_المسجلة": "المسجلة في النظام",
                        "المنتظر": "المنتظر",
                        "المنتظر_القديم": "المنتظر (القديم)"
                    })

                    st.dataframe(grouped)

                    # --- حساب الإجماليات ---
                    total_excel = int(df_filtered["عدد الحالات التي تمت_الحديث"].sum())
                    total_previous = int(df_filtered["عدد الحالات التي تمت_القديم"].sum())
                    total_diff = int(df_filtered["الزيادة"].sum()) if "الزيادة" in df_filtered.columns else total_excel - total_previous
                    total_reported = int(df_filtered["الحالات_المسجلة"].sum())
                    total_waiting = int(df_filtered["المنتظر"].sum())
                    total_waiting_old = int(df_filtered["المنتظر_القديم"].sum()) if "المنتظر_القديم" in df_filtered.columns else 0

                    # --- زر تصدير PDF ---
                    if st.button("🖨️ تصدير PDF", key="generate_pdf"):
                        table_html = grouped.to_html(table_id="grouped-table", border=0, classes="table", escape=False)

                        html = f"""
                        <html>
                        <head>
                            <meta charset="UTF-8">
                            <style>
                                @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap');
                                body {{
                                    font-family: 'Cairo', 'Arial', sans-serif;
                                    direction: rtl;
                                    text-align: right;
                                    background: #f8f9fa;
                                    padding: 20px;
                                    color: #333;
                                }}
                                .header {{
                                    text-align: center;
                                    border-bottom: 2px solid #1a237e;
                                    padding-bottom: 10px;
                                    margin-bottom: 20px;
                                    color: #1a237e;
                                }}
                                .header h1 {{
                                    margin: 0;
                                    font-size: 24px;
                                }}
                                .summary {{
                                    background-color: #e3f2fd;
                                    padding: 15px;
                                    border-radius: 8px;
                                    margin: 20px 0;
                                    font-size: 16px;
                                }}
                                .summary strong {{
                                    color: #1a237e;
                                }}
                                table {{
                                    width: 100%;
                                    border-collapse: collapse;
                                    margin: 20px 0;
                                    font-size: 14px;
                                }}
                                th, td {{
                                    border: 1px solid #ddd;
                                    padding: 10px;
                                    text-align: center;
                                }}
                                th {{
                                    background-color: #1a237e;
                                    color: white;
                                }}
                                tr:nth-child(even) {{
                                    background-color: #f2f2f2;
                                }}
                                .footer {{
                                    text-align: center;
                                    margin-top: 30px;
                                    color: #777;
                                    font-size: 12px;
                                }}
                            </style>
                        </head>
                        <body>
                            <div class="header">
                                <h1>📄 تقرير المقارنة</h1>
                                <p><strong>الفترة:</strong> {selected_earlier} إلى {selected_later}</p>
                                <p><strong>التجميع:</strong> {selected_group}</p>
                            </div>
                            <div class="summary">
                                <p>📌 <strong>إجمالي الحالات (الحديث):</strong> {total_excel:,}</p>
                                <p>📌 <strong>إجمالي الحالات (القديم):</strong> {total_previous:,}</p>
                                <p>📈 <strong>الزيادة الكلية:</strong> {total_diff:+,}</p>
                                <p>✅ <strong>المسجلة في النظام:</strong> {total_reported:,}</p>
                                <p>⏳ <strong>الحالات المنتظرة (الحديث):</strong> {total_waiting:,}</p>
                                <p>⏳ <strong>الحالات المنتظرة (القديم):</strong> {total_waiting_old:,}</p>
                            </div>
                            <h3>📊 تقرير التجميع ({selected_group})</h3>
                            {table_html}
                            <div class="footer">
                                <p>تم إنشاء هذا التقرير تلقائيًا بواسطة نظام تقارير المستشفيات</p>
                                <p>تاريخ الإنشاء: {date.today()}</p>
                            </div>
                        </body>
                        </html>
                        """

                        pdf_path = f"تقرير_مقارنة_{selected_earlier}_إلى_{selected_later}.pdf"
                        try:
                            if config:
                                pdfkit.from_string(html, pdf_path, configuration=config)
                                with open(pdf_path, "rb") as f:
                                    st.download_button(
                                        "⬇️ تنزيل PDF",
                                        f,
                                        pdf_path,
                                        "application/pdf",
                                        key="download_pdf_final"
                                    )
                                os.remove(pdf_path)
                            else:
                                st.error("❌ تم تعطيل PDF: تحقق من تثبيت wkhtmltopdf.")
                        except Exception as e:
                            st.error(f"❌ خطأ في إنشاء PDF: {e}")

            # --- عرض الملفات المحفوظة ---
            st.subheader("📁 الملفات المحفوظة")
            if not snapshots.empty:
                for _, row in snapshots.iterrows():
                    cols = st.columns([3, 2, 1, 1])
                    cols[0].write(row["filename"])
                    cols[1].write(row["report_date"])
                    try:
                        if os.path.exists(row["file_path"]):
                            if cols[3].button("🗑️", key=f"del_{row['id']}"):
                                os.remove(row["file_path"])
                                with db_conn() as c:
                                    c.execute("DELETE FROM snapshots WHERE id = ?", (row["id"],))
                                st.success(f"✅ تم حذف الملف: {row['filename']}")
                                st.rerun()
                        else:
                            st.warning(f"⚠️ الملف غير موجود: {row['file_path']}")
                            if st.button("❌ إزالة من القاعدة", key=f"fix_{row['id']}"):
                                with db_conn() as c:
                                    c.execute("DELETE FROM snapshots WHERE id = ?", (row["id"],))
                                st.rerun()
                    except Exception as e:
                        st.error(f"❌ خطأ في الحذف: {e}")
            else:
                st.info("لا توجد ملفات محفوظة بعد.")

        elif sub_menu == "📈 تحليل متقدم للملفات المحفوظة":
            st.subheader("📈 تحليل متقدم للملفات المحفوظة")

            with db_conn() as c:
                snapshots = pd.read_sql("""
                    SELECT id, filename, report_date, upload_date, file_path 
                    FROM snapshots 
                    ORDER BY report_date ASC
                """, c)

            if snapshots.empty or len(snapshots) < 2:
                st.info("❌ يُرجى رفع ملفين على الأقل لإجراء التحليل.")
            else:
                snapshots["report_date"] = pd.to_datetime(snapshots["report_date"]).dt.date
                # --- اختيار الفترة الزمنية ---
                st.markdown("### 📅 اختيار الفترة للتحليل")
                col1, col2 = st.columns(2)
                with col1:
                    start_date = st.date_input(
                        "من تاريخ",
                        value=min(snapshots["report_date"]),
                        min_value=min(snapshots["report_date"]),
                        max_value=max(snapshots["report_date"])
                    )
                with col2:
                    end_date = st.date_input(
                        "إلى تاريخ",
                        value=max(snapshots["report_date"]),
                        min_value=min(snapshots["report_date"]),
                        max_value=max(snapshots["report_date"])
                    )

                if start_date > end_date:
                    st.error("❌ تاريخ البداية لا يمكن أن يكون بعد تاريخ النهاية.")
                else:
                    # --- تصفية الملفات حسب الفترة ---
                    filtered_files = []
                    for _, row in snapshots.iterrows():
                        if start_date <= row['report_date'] <= end_date:
                            filtered_files.append({
                                'filename': row['filename'],
                                'report_date': row['report_date'],
                                'file_path': row['file_path']
                            })
                    if len(filtered_files) < 1:
                        st.warning("⚠️ لا توجد ملفات كافية في الفترة المحددة.")
                    else:
                        st.info(f"📁 عدد الملفات في الفترة: {len(filtered_files)}")

                        # --- زر بدء التحليل ---
                        if st.button("🔍 بدء التحليل", key="btn_advanced_analysis"):
                            st.session_state["advanced_analysis_triggered"] = True
                            st.session_state["advanced_analysis_params"] = {
                                "start_date": start_date,
                                "end_date": end_date,
                                "filtered_files": filtered_files
                            }

                        if not st.session_state.get("advanced_analysis_triggered", False):
                            st.info("ℹ️ اضغط 'بدء التحليل' لبدء قراءة الملفات وتنفيذ التحليل.")
                        else:
                            # --- قراءة جميع الملفات في الفترة ---
                            with st.spinner("⏳ جاري قراءة الملفات..."):
                                all_data = []
                                progress_bar = st.progress(0)
                                status_text = st.empty()
                                
                                for idx, file in enumerate(filtered_files):
                                    try:
                                        status_text.text(f"📂 قراءة الملف {idx + 1}/{len(filtered_files)}: {file['filename']}")
                                        
                                        # قراءة الملف بكفاءة عالية
                                        df_file = pd.read_excel(
                                            file["file_path"], 
                                            engine='openpyxl',
                                            dtype=str,  # قراءة كل شيء كنص أولاً
                                            na_filter=False  # تعطيل الكشف التلقائي عن NA لتحسين الأداء
                                        )
                                        
                                        df_file["report_date"] = file["report_date"]
                                        required_cols = [
                                            "المحافظه", "المستشفى", "تصنيف المستشفى", "تصنيف الاجراء",
                                            "عدد الحالات التي تمت"
                                        ]
                                        
                                        # التحقق من الأعمدة
                                        missing_cols = [col for col in required_cols if col not in df_file.columns]
                                        if missing_cols:
                                            st.error(f"❌ الملف {file['filename']} يفتقد الأعمدة: {', '.join(missing_cols)}")
                                            continue
                                        
                                        # اختيار الأعمدة المطلوبة فقط
                                        df_file = df_file[required_cols + ["report_date"]].copy()
                                        
                                        # تحويل عدد الحالات إلى رقم
                                        df_file["عدد الحالات التي تمت"] = pd.to_numeric(
                                            df_file["عدد الحالات التي تمت"], 
                                            errors='coerce'
                                        ).fillna(0)
                                        
                                        df_file.rename(columns={"عدد الحالات التي تمت": "cases"}, inplace=True)
                                        
                                        # إزالة الصفوف الفارغة
                                        df_file = df_file.dropna(subset=["المستشفى", "تصنيف الاجراء"])
                                        
                                        all_data.append(df_file)
                                        progress_bar.progress((idx + 1) / len(filtered_files))
                                        
                                    except Exception as e:
                                        st.error(f"❌ خطأ في قراءة الملف {file['filename']}: {str(e)}")
                                        continue
                                
                                progress_bar.empty()
                                status_text.empty()
    
                            if not all_data:
                                st.error("❌ لم يتم قراءة أي ملف بنجاح. يرجى التحقق من:")
                                st.markdown("""
                                - صحة تنسيق ملفات Excel
                                - وجود الأعمدة المطلوبة: المحافظه، المستشفى، تصنيف المستشفى، تصنيف الاجراء، عدد الحالات التي تمت
                                - عدم وجود ملفات تالفة
                                """)
                            else:
                                with st.spinner("⏳ جاري دمج ومعالجة البيانات..."):
                                    try:
                                        # دمج البيانات بكفاءة
                                        df_all = pd.concat(all_data, ignore_index=True, copy=False)
                                        
                                        # تحويل التاريخ
                                        df_all["report_date"] = pd.to_datetime(df_all["report_date"])
                                        
                                        # تنظيف البيانات
                                        df_all["cases"] = pd.to_numeric(df_all["cases"], errors='coerce').fillna(0)
                                        
                                        # إزالة المسافات الزائدة
                                        for col in ["المحافظه", "المستشفى", "تصنيف المستشفى", "تصنيف الاجراء"]:
                                            df_all[col] = df_all[col].astype(str).str.strip()
                                        
                                        st.success(f"✅ تم تحميل {len(all_data)} ملف بنجاح ({len(df_all):,} سجل)")
                                    except Exception as e:
                                        st.error(f"❌ خطأ في معالجة البيانات: {str(e)}")
                                        st.stop()
    
                                # --- التصفية المتقدمة ---
                                st.markdown("### ⚙️ التصفية المتقدمة")
                                col1, col2 = st.columns(2)
                                with col1:
                                    gov_filter = st.multiselect("📍 المحافظه", df_all["المحافظه"].unique(), key="adv_gov")
                                    hosp_filter = st.multiselect("🏥 المستشفى", df_all["المستشفى"].unique(), key="adv_hosp")
                                with col2:
                                    type_filter = st.multiselect("🏷️ تصنيف المستشفى", df_all["تصنيف المستشفى"].unique(), key="adv_type")
                                    proc_filter = st.multiselect("🛠️ التخصص", df_all["تصنيف الاجراء"].unique(), key="adv_proc")
    
                                df_filtered = df_all.copy()
                                if gov_filter:
                                    df_filtered = df_filtered[df_filtered["المحافظه"].isin(gov_filter)]
                                if hosp_filter:
                                    df_filtered = df_filtered[df_filtered["المستشفى"].isin(hosp_filter)]
                                if type_filter:
                                    df_filtered = df_filtered[df_filtered["تصنيف المستشفى"].isin(type_filter)]
                                if proc_filter:
                                    df_filtered = df_filtered[df_filtered["تصنيف الاجراء"].isin(proc_filter)]
    
                                st.markdown("### 📊 البيانات حسب الزمن")
                                st.info(f"📊 إجمالي السجلات: {len(df_filtered):,}")
                                if len(df_filtered) > 10000:
                                    st.warning("⚠️ عدد السجلات كبير جداً. سيتم عرض أول 10,000 سجل فقط.")
                                    st.dataframe(df_filtered.head(10000))
                                else:
                                    st.dataframe(df_filtered)
    
                                # --- تحليل الزيادة اليومية مع المنتظر ---
                                st.markdown("### 📈 تقرير الزيادة اليومية")
    
                                with st.spinner("⏳ جاري حساب الزيادة اليومية..."):
                                    # تجميع الحالات لكل يوم مع تفاصيل المستشفى والتخصص
                                    daily_details = df_filtered.groupby(["report_date", "المستشفى", "تصنيف الاجراء"]).agg({
                                        "cases": "sum"
                                    }).reset_index()
                                    daily_details = daily_details.sort_values(["report_date", "المستشفى", "تصنيف الاجراء"])
                                    
                                    # حساب الزيادة لكل مستشفى وتخصص
                                    daily_details['الزيادة'] = daily_details.groupby(["المستشفى", "تصنيف الاجراء"])['cases'].diff().fillna(0).apply(lambda x: max(x, 0)).astype(int)
                                
                                with st.spinner("⏳ جاري حساب الحالات المنتظرة..."):
                                    # تحميل جميع الملفات مرة واحدة لتحسين الأداء مع كاش في session_state
                                    cache_key = f"files_cache_{start_date}_{end_date}_{len(filtered_files)}"
                                    if cache_key not in st.session_state:
                                        files_cache = {}
                                        status_text = st.empty()
                                        for idx, file in enumerate(filtered_files):
                                            try:
                                                status_text.text(f"📂 تحميل ملف {idx + 1}/{len(filtered_files)} للمنتظر...")
                                                temp_df = pd.read_excel(
                                                    file["file_path"],
                                                    engine='openpyxl',
                                                    usecols=["المستشفى", "تصنيف الاجراء", "عدد الحالات الجديدة", "عدد الحالات الجارية"],
                                                    dtype=str,
                                                    na_filter=False
                                                )
                                                files_cache[file["report_date"]] = temp_df
                                            except Exception:
                                                files_cache[file["report_date"]] = None
                                        st.session_state[cache_key] = files_cache
                                        status_text.empty()
                                    else:
                                        files_cache = st.session_state[cache_key]
                                    
                                    # حساب المنتظر باستخدام عمليات متجهة بدلاً من الحلقات
                                    daily_waiting = []
                                    progress_bar = st.progress(0)
                                    status_text = st.empty()
                                    total_rows = len(daily_details)
                                    
                                    for idx, (_, row) in enumerate(daily_details.iterrows()):
                                        current_date = row["report_date"].date()
                                        hospital = row["المستشفى"]
                                        procedure = row["تصنيف الاجراء"]
                                        
                                        temp_df = files_cache.get(current_date)
                                        waiting = 0
                                        
                                        if temp_df is not None:
                                            try:
                                                if "عدد الحالات الجديدة" in temp_df.columns and "عدد الحالات الجارية" in temp_df.columns:
                                                    # استخدام عمليات متجهة
                                                    mask = (temp_df["المستشفى"] == hospital) & (temp_df["تصنيف الاجراء"] == procedure)
                                                    filtered_temp = temp_df[mask]
                                                    
                                                    if not filtered_temp.empty:
                                                        new_cases = pd.to_numeric(filtered_temp["عدد الحالات الجديدة"], errors='coerce').fillna(0).sum()
                                                        ongoing_cases = pd.to_numeric(filtered_temp["عدد الحالات الجارية"], errors='coerce').fillna(0).sum()
                                                        waiting = int(new_cases + ongoing_cases)
                                            except Exception as e:
                                                waiting = 0
                                        
                                        daily_waiting.append(waiting)
                                        
                                        # تحديث شريط التقدم كل 50 سجل
                                        if idx % 50 == 0:
                                            progress = min((idx + 1) / total_rows, 1.0)
                                            progress_bar.progress(progress)
                                            status_text.text(f"⏳ معالجة السجل {idx + 1}/{total_rows} ({progress*100:.1f}%)")
                                    
                                    progress_bar.empty()
                                    status_text.empty()
                                    daily_details['المنتظر'] = daily_waiting
                                    st.success(f"✅ تم حساب المنتظر لـ {len(daily_waiting):,} سجل")
                                
                                # حساب الإجماليات للعرض
                                daily_totals = daily_details.groupby("report_date").agg({
                                    "cases": "sum",
                                    "الزيادة": "sum",
                                    "المنتظر": "sum"
                                }).reset_index()
                                daily_totals = daily_totals.sort_values("report_date")
    
    
                                
    
    
                                # حساب الإجمالي الكلي للزيادة خلال الفترة
                                total_daily_increase = daily_totals["الزيادة"].sum()
                                # حساب إجمالي المنتظر من أحدث ملف إكسيل
                                if len(filtered_files) > 0:
                                    latest_file = filtered_files[-1]
                                    try:
                                        latest_df = pd.read_excel(latest_file["file_path"])
                                        if "عدد الحالات الجديدة" in latest_df.columns and "عدد الحالات الجارية" in latest_df.columns:
                                            if gov_filter:
                                                latest_df = latest_df[latest_df["المحافظه"].isin(gov_filter)]
                                            if hosp_filter:
                                                latest_df = latest_df[latest_df["المستشفى"].isin(hosp_filter)]
                                            if type_filter:
                                                latest_df = latest_df[latest_df["تصنيف المستشفى"].isin(type_filter)]
                                            if proc_filter:
                                                latest_df = latest_df[latest_df["تصنيف الاجراء"].isin(proc_filter)]
                                            total_expected = int((latest_df["عدد الحالات الجديدة"] + latest_df["عدد الحالات الجارية"]).sum())
                                        else:
                                            total_expected = 0
                                    except:
                                        total_expected = 0
                                else:
                                    total_expected = 0
    
                                # عرض الجدول التفصيلي
                                daily_display = daily_details.copy()
                                daily_display["report_date"] = daily_display["report_date"].dt.strftime("%Y-%m-%d")
                                daily_display.rename(columns={
                                    "report_date": "التاريخ",
                                    "المستشفى": "المستشفى",
                                    "تصنيف الاجراء": "تصنيف الاجراء",
                                    "cases": "عدد الحالات",
                                    "الزيادة": "الزيادة",
                                    "المنتظر": "المنتظر"
                                }, inplace=True)
    
                                st.info(f"📊 إجمالي السجلات: {len(daily_display):,}")
                                if len(daily_display) > 5000:
                                    st.warning("⚠️ عدد السجلات كبير. سيتم عرض أول 5,000 سجل فقط في الجدول.")
                                    st.dataframe(daily_display.head(5000), use_container_width=True)
                                else:
                                    st.dataframe(daily_display, use_container_width=True)
    
                                # عرض الإجماليات
                                col1, col2 = st.columns(2)
                                with col1:
                                    st.markdown(f"### ✅ **إجمالي الزيادة: {int(total_daily_increase):,} حالة**")
                                with col2:
                                    st.markdown(f"### ⏳ **إجمالي المنتظر: {int(total_expected):,} حالة**")
    
                                # --- رسوم بيانية ---
                                st.markdown("### 📉 الرسوم البيانية")
    
                                fig_line = px.line(
                                    daily_totals,
                                    x="report_date",
                                    y="cases",
                                    markers=True,
                                    title="📈 إجمالي الحالات اليومية",
                                    labels={"report_date": "التاريخ", "cases": "عدد الحالات"}
                                )
                                st.plotly_chart(fig_line, use_container_width=True)
    
                                fig_bar = px.bar(
                                    daily_totals,
                                    x="report_date",
                                    y="الزيادة",
                                    title="📊 الزيادة اليومية",
                                    labels={"report_date": "التاريخ", "الزيادة": "الزيادة"}
                                )
                                st.plotly_chart(fig_bar, use_container_width=True)
                                
                                # رسم بياني للمنتظر
                                fig_expected = px.line(
                                    daily_totals,
                                    x="report_date",
                                    y="المنتظر",
                                    markers=True,
                                    title="⏳ الحالات المنتظرة (متوسط متحرك)",
                                    labels={"report_date": "التاريخ", "المنتظر": "المنتظر"}
                                )
                                st.plotly_chart(fig_expected, use_container_width=True)
    
                                # --- تصدير التقرير النهائي مع الأعمدة الجديدة ---
                                st.markdown("### 📥 تنزيل التقرير النهائي")
                                output = BytesIO()
                                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                                    # ورقة التقرير اليومي مع الأعمدة الجديدة
                                    daily_display.to_excel(writer, index=False, sheet_name="التقرير اليومي")
                                    # ورقة البيانات الخام
                                    df_filtered.to_excel(writer, index=False, sheet_name="البيانات الخام")
                                    # ورقة الملخص
                                    summary_data = pd.DataFrame({
                                        "المؤشر": ["إجمالي الزيادة", "إجمالي المنتظر", "عدد الأيام", "متوسط الزيادة اليومية"],
                                        "القيمة": [int(total_daily_increase), int(total_expected), len(daily_display), int(total_daily_increase/len(daily_display)) if len(daily_display) > 0 else 0]
                                    })
                                    summary_data.to_excel(writer, index=False, sheet_name="الملخص")
                                output.seek(0)
                                st.download_button(
                                    "⬇️ تنزيل التقرير النهائي مع الأعمدة الجديدة",
                                    output,
                                    f"تقرير_يومي_متقدم_{start_date}_الى_{end_date}.xlsx",
                                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    key="download_advanced_report"
                                )

    # 📊 تقارير Snapshot المتقدمة
    elif menu == "📊 تقارير Snapshot المتقدمة" and has_permission("admin") and SNAPSHOT_REPORTS_AVAILABLE:
        st.title("📊 تقارير Snapshot المتقدمة")
        if not st.session_state.get("snapshot_analysis_triggered", False):
            st.info("ℹ️ اضغط الزر أدناه لبدء تحميل وتحليل بيانات Snapshot.")
            if st.button("🔍 بدء تحليل Snapshot", key="btn_snapshot_analysis"):
                st.session_state["snapshot_analysis_triggered"] = True
                st.rerun()
        else:
            display_snapshot_analysis()

    elif menu == "🚪 خروج":
        st.title("🚪 تسجيل الخروج")
        st.warning("⚠️ هل أنت متأكد من تسجيل الخروج؟")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("✅ نعم، تسجيل الخروج", type="primary", use_container_width=True):
                logout()
        with col2:
            if st.button("❌ إلغاء", use_container_width=True):
                st.rerun()

# ========== تشغيل التطبيق ==========

# تهيئة تحسينات الأداء
initialize_performance_improvements()

if not st.session_state.logged:
    login()
else:
    if st.session_state.user == "admin" or get_user_info(st.session_state.user)[2] == "admin":
        admin_view()
    else:
        user_view()