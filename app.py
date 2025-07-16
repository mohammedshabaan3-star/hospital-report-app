# app.py

import streamlit as st
import sqlite3
import pandas as pd
import os
from datetime import date
from pathlib import Path
from io import BytesIO
import shutil

# الشعار
st.set_page_config(layout="wide", page_title="نظام تقارير المستشفيات")
st.image("assets/logo.png", width=150)

# المسارات
DB_PATH = "data/reports.db"
EXCEL_PATH = "data.xlsx"
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

# إنشاء مجلد البيانات إن لم يكن موجودًا
Path("data").mkdir(exist_ok=True)

# 🔄 نسخ احتياطي لقاعدة البيانات
def backup_database():
    backup_path = "data/reports_backup.db"
    if os.path.exists(DB_PATH):
        shutil.copy(DB_PATH, backup_path)
        return True
    return False

# 📥 استيراد بيانات من ملف Excel
def import_excel_data():
    df = pd.read_excel(EXCEL_PATH, sheet_name=0, skiprows=2, usecols=[1, 2, 3, 4])
    df.columns = ["governorate", "hospital", "procedure", "capacity"]
    df.dropna(subset=["hospital", "procedure", "capacity"], inplace=True)
    df["capacity"] = pd.to_numeric(df["capacity"], errors="coerce").fillna(0).astype(float)

    with sqlite3.connect(DB_PATH) as conn:
        c = conn.cursor()

        # إضافة المستشفيات كمستخدمين
        hospitals = df[["hospital", "governorate"]].drop_duplicates()
        for _, row in hospitals.iterrows():
            username = row["hospital"].strip().replace(" ", "_").lower()
            c.execute("""INSERT OR IGNORE INTO users 
                         VALUES (?, ?, ?, ?, ?)""",
                      (username, "1234", row["hospital"].strip(), row["governorate"].strip(), "user"))

        # إضافة التخصصات لكل مستشفى
        for _, row in df.iterrows():
            username = row["hospital"].strip().replace(" ", "_").lower()
            procedure = str(row["procedure"]).strip()
            capacity = float(row["capacity"])
            c.execute("SELECT COUNT(*) FROM specialties WHERE username=? AND procedure=?", (username, procedure))
            if c.fetchone()[0] == 0:
                c.execute("""INSERT INTO specialties (username, procedure, capacity)
                             VALUES (?, ?, ?)""", (username, procedure, capacity))
        conn.commit()

# ⚙️ إنشاء قاعدة البيانات إن لم تكن موجودة
def ensure_db():
    if not os.path.exists(DB_PATH):
        with sqlite3.connect(DB_PATH) as conn:
            cur = conn.cursor()
            cur.execute("""CREATE TABLE IF NOT EXISTS users (
                username TEXT PRIMARY KEY,
                password TEXT NOT NULL,
                hospital TEXT,
                governorate TEXT,
                role TEXT CHECK(role IN ('admin','user')) NOT NULL
            )""")
            cur.execute("""CREATE TABLE IF NOT EXISTS specialties (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT,
                procedure TEXT,
                capacity REAL,
                FOREIGN KEY(username) REFERENCES users(username)
            )""")
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
            cur.execute("""INSERT OR IGNORE INTO users 
                (username, password, hospital, governorate, role)
                VALUES ('admin', 'admin123', 'وزارة الصحة', 'القاهرة', 'admin')""")
        import_excel_data()

# 🔌 الاتصال بقاعدة البيانات
def db_conn():
    ensure_db()
    return sqlite3.connect(DB_PATH)

# 👤 جلب بيانات مستخدم
def get_user_info(username):
    with db_conn() as c:
        cur = c.cursor()
        cur.execute("SELECT hospital, governorate, role FROM users WHERE username = ?", (username,))
        return cur.fetchone()

# 🔑 تسجيل الدخول
def login():
    st.subheader("🔐 تسجيل الدخول")
    u = st.text_input("اسم المستخدم")
    p = st.text_input("كلمة المرور", type="password")
    if st.button("دخول"):
        with db_conn() as c:
            cur = c.cursor()
            cur.execute("SELECT password FROM users WHERE username = ?", (u.strip(),))
            row = cur.fetchone()
            if row and row[0] == p.strip():
                st.session_state.logged = True
                st.session_state.user = u.strip()
                st.rerun()
            else:
                st.error("❌ اسم المستخدم أو كلمة المرور غير صحيحة.")

# 🚪 تسجيل الخروج
def logout():
    st.session_state.logged = False
    st.session_state.user = ""
    st.session_state.pending = []
    st.rerun()

# 🧠 الحالة الافتراضية عند بدء البرنامج
if "logged" not in st.session_state:
    st.session_state.logged = False
    st.session_state.user = ""
    st.session_state.pending = []
# 👤 واجهة المستخدم العادي
def user_view():
    uname = st.session_state.user
    hosp, gov, _ = get_user_info(uname)
    st.sidebar.title(f"🏥 {hosp}")
    menu = st.sidebar.radio("القائمة", ["➕ تقرير", "📋 تقاريري", "🔐 الحساب", "🚪 خروج"])

    if menu == "➕ تقرير":
        st.title("📝 تقرير جديد")
        
        period = st.radio("نوع التقرير", ["يومي", "أسبوعي", "شهري"])
        dfrom = st.date_input("📅 من تاريخ", value=date.today())
        if period != "يومي":
            dto = st.date_input("📅 إلى تاريخ", value=date.today())
        else:
            dto = dfrom  # نفس اليوم في التقرير اليومي

        with db_conn() as c:
            df = pd.read_sql("SELECT procedure, capacity FROM specialties WHERE username = ?", c, params=[uname])
        
        if df.empty:
            st.warning("⚠️ لا توجد تخصصات مسجلة.")
            return

        proc = st.selectbox("🔧 اختر التخصص", df["procedure"].unique())
        cap = df[df["procedure"] == proc]["capacity"].values[0]
        st.markdown(f"⚡ السعة: **{cap}**")
        cases = st.number_input("📌 عدد الحالات", min_value=0)
        notes = st.text_area("💬 ملاحظات / أعطال / احتياجات")
        pdf = st.file_uploader("📎 ملف PDF", type="pdf")

        if st.button("➕ إضافة مؤقتًا"):
            pdf_path = ""
            if pdf:
                path = Path(UPLOAD_DIR) / f"{uname}_{proc}_{dfrom}.pdf"
                with open(path, "wb") as f:
                    f.write(pdf.getbuffer())
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
                            username, period_type, date_from, date_to,
                            procedure, capacity, cases, notes, pdf
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                        (uname, e["period"], e["from"], e["to"], e["procedure"],
                         e["capacity"], e["cases"], e["notes"], e["pdf"]))
                st.session_state.pending = []
                st.success("✅ تم حفظ التقرير.")

    elif menu == "📋 تقاريري":
        with db_conn() as c:
            df = pd.read_sql("SELECT * FROM reports WHERE username = ?", c, params=[uname])
        st.title("📊 تقاريري السابقة")
        st.dataframe(df)

    elif menu == "🔐 الحساب":
        st.title("🔐 تعديل بيانات الحساب")
        st.text_input("اسم المستخدم", value=uname, disabled=True)
        new_pass = st.text_input("كلمة مرور جديدة", type="password")
        if st.button("🔄 تحديث كلمة المرور"):
            with db_conn() as c:
                c.execute("UPDATE users SET password=? WHERE username=?", (new_pass, uname))
            st.success("✅ تم تحديث كلمة المرور. سجّل دخولك مجددًا.")
            logout()

    elif menu == "🚪 خروج":
        logout()
def reports_view():
    st.title("📊 جميع التقارير")
    with db_conn() as c:
        df = pd.read_sql("""
            SELECT r.*, u.hospital, u.governorate 
            FROM reports r 
            JOIN users u ON r.username = u.username
        """, c)

    ufilter = st.multiselect("📌 تصفية حسب المستخدم", df["username"].unique())
    pfilter = st.multiselect("📌 حسب التخصص", df["procedure"].unique())
    gfilter = st.multiselect("📌 حسب المحافظة", df["governorate"].unique())
    d1 = st.date_input("📅 من تاريخ")
    d2 = st.date_input("📅 إلى تاريخ")

    if ufilter:
        df = df[df["username"].isin(ufilter)]
    if pfilter:
        df = df[df["procedure"].isin(pfilter)]
    if gfilter:
        df = df[df["governorate"].isin(gfilter)]
    if d1 and d2:
        df = df[(df["date_from"] >= str(d1)) & (df["date_to"] <= str(d2))]

    st.dataframe(df)

    # حالة إرسال التقارير
    st.subheader("📌 حالة إرسال التقارير")
    with db_conn() as c:
        df_users = pd.read_sql("SELECT username, hospital FROM users WHERE role='user'", c)
    sent_users = df["username"].unique()
    sent_hospitals = df_users[df_users["username"].isin(sent_users)]["hospital"].unique()
    unsent_hospitals = df_users[~df_users["username"].isin(sent_users)]["hospital"].unique()
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("#### ✅ مستشفيات أرسلت التقارير")
        if len(sent_hospitals) > 0:
            st.success(f"تم الإرسال من: {len(sent_hospitals)} مستشفى")
            st.write(sent_hospitals)
        else:
            st.warning("⚠️ لا توجد مستشفيات أرسلت تقارير ضمن التصفية الحالية.")

    with col2:
        st.markdown("#### 🚫 مستشفيات لم ترسل التقارير")
        if len(unsent_hospitals) > 0:
            st.error(f"عددها: {len(unsent_hospitals)}")
            st.write(unsent_hospitals)
        else:
            st.info("✅ كل المستشفيات أرسلت تقارير بناءً على التصفية.")

    # حذف تقرير
    st.subheader("🗑️ حذف تقرير معين")
    if not df.empty:
        selected_id = st.selectbox("📄 اختر رقم التقرير للحذف", df["id"].sort_values(ascending=False))
        r = df[df["id"] == selected_id].iloc[0]
        if st.button("🗑️ حذف هذا التقرير"):
            with db_conn() as c:
                c.execute("DELETE FROM reports WHERE id=?", (selected_id,))
            st.success(f"✅ تم حذف التقرير رقم {selected_id}.")
            st.rerun()

    # تصدير
    output = BytesIO()
    df.to_excel(output, index=False, engine="openpyxl")
    output.seek(0)
    st.download_button("⬇️ تحميل Excel", output, "hospital_reports.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with st.expander("🛠️ أدوات إضافية"):
        if st.button("📥 تحميل بيانات المستشفيات والتخصصات من Excel"):
            import_excel_data()
            st.success("✅ تم التحديث من Excel.")

        if st.button("💾 حفظ نسخة احتياطية"):
            if backup_database():
                st.success("✅ تم حفظ النسخة الاحتياطية.")
            else:
                st.error("❌ تعذر العثور على قاعدة البيانات.")
def hospitals_view():
    st.title("🏥 إدارة المستشفيات")

    name = st.text_input("اسم المستشفى")
    gov = st.text_input("المحافظة")
    uname = name.strip().replace(" ", "_").lower()
    pw = st.text_input("كلمة المرور", value="1234")

    if st.button("➕ إضافة مستشفى"):
        with db_conn() as c:
            c.execute("INSERT OR IGNORE INTO users VALUES (?, ?, ?, ?, ?)",
                      (uname, pw, name, gov, "user"))
        st.success("✅ تمت إضافة المستشفى.")

    # عرض المستشفيات الحالية
    with db_conn() as c:
        df_hosp = pd.read_sql("SELECT username, hospital FROM users WHERE role='user'", c)
        df_count = pd.read_sql("SELECT username, COUNT(*) AS specialties FROM specialties GROUP BY username", c)
        df_all = pd.merge(df_hosp, df_count, on="username", how="left").fillna(0)

    st.dataframe(df_all.rename(columns={"hospital": "اسم المستشفى", "specialties": "عدد التخصصات"}))

    # حذف مستشفى
    del_hosp = st.selectbox("🗑️ اختر مستشفى لحذفه", df_all["hospital"].unique())
    code = df_all[df_all["hospital"] == del_hosp]["username"].values[0]

    if st.button("🗑️ حذف المستشفى بالكامل"):
        with db_conn() as c:
            c.execute("DELETE FROM reports WHERE username=?", (code,))
            c.execute("DELETE FROM specialties WHERE username=?", (code,))
            c.execute("DELETE FROM users WHERE username=?", (code,))
        st.success(f"✅ تم حذف '{del_hosp}' بالكامل.")
        st.rerun()
def specialties_view():
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

    if df_spec.empty:
        st.info("📭 لا توجد تخصصات لهذا المستشفى.")
    else:
        row_map = {f"{r['procedure']} (سعة: {r['capacity']})": r for _, r in df_spec.iterrows()}
        selected = st.selectbox("🔧 اختر تخصص للتعديل أو الحذف", list(row_map.keys()))
        r = row_map[selected]

        new_proc = st.text_input("✏️ تعديل اسم التخصص", value=r["procedure"])
        new_cap = st.number_input("🔄 تعديل السعة", min_value=0.0, value=float(r["capacity"]))

        col1, col2 = st.columns(2)
        with col1:
            if st.button("💾 تعديل"):
                with db_conn() as c:
                    c.execute("UPDATE specialties SET procedure=?, capacity=? WHERE id=?",
                              (new_proc.strip(), new_cap, r["id"]))
                st.success("✅ تم التعديل.")
                st.rerun()
        with col2:
            if st.button("🗑️ حذف"):
                with db_conn() as c:
                    c.execute("DELETE FROM specialties WHERE id=?", (r["id"],))
                st.success("🗑️ تم الحذف.")
                st.rerun()

    st.markdown("### ➕ إضافة تخصص جديد")
    new_proc_add = st.selectbox("🆕 اختر تخصص عام", all_specs["procedure"].unique())
    new_cap_add = st.number_input("📌 السعة الاستيعابية", min_value=0.0)

    if st.button("➕ أضف التخصص"):
        with db_conn() as c:
            c.execute("INSERT INTO specialties (username, procedure, capacity) VALUES (?, ?, ?)",
                      (uname, new_proc_add.strip(), new_cap_add))
        st.success("✅ تم إضافة التخصص.")
        st.rerun()
def reset_password_view():
    st.title("🔒 إعادة تعيين كلمة مرور مستخدم")

    with db_conn() as c:
        df_users = pd.read_sql("SELECT username, hospital FROM users WHERE role='user'", c)

    if df_users.empty:
        st.warning("🚫 لا توجد مستشفيات مسجلة.")
        return

    selected_hosp = st.selectbox("🏥 اختر مستشفى", df_users["hospital"])
    uname = df_users[df_users["hospital"] == selected_hosp]["username"].values[0]
    new_pass = st.text_input("🔑 كلمة المرور الجديدة")

    if st.button("🔄 تعيين كلمة المرور"):
        with db_conn() as c:
            c.execute("UPDATE users SET password=? WHERE username=?", (new_pass, uname))
        st.success(f"✅ تم تعيين كلمة مرور جديدة لـ '{selected_hosp}'")
def admin_account_view():
    st.title("🔐 تعديل بيانات الأدمن")

    uname = st.session_state.user
    st.text_input("اسم المستخدم", value=uname, disabled=True)
    new_pass = st.text_input("كلمة مرور جديدة", type="password")

    if st.button("🔄 تحديث كلمة المرور"):
        with db_conn() as c:
            c.execute("UPDATE users SET password=? WHERE username=?", (new_pass, uname))
        st.success("✅ تم تحديث كلمة المرور.")
        logout()
def admin_view():
    st.sidebar.title("👤 الأدمن")
    menu = st.sidebar.radio("القائمة", [
        "📊 التقارير",
        "🏥 مستشفيات",
        "⚙️ تخصصات",
        "🔐 الحساب",
        "🔒 استرجاع كلمة المرور",
        "🚪 خروج"
    ])

    if menu == "📊 التقارير":
        reports_view()
    elif menu == "🏥 مستشفيات":
        hospitals_view()
    elif menu == "⚙️ تخصصات":
        specialties_view()
    elif menu == "🔐 الحساب":
        admin_account_view()
    elif menu == "🔒 استرجاع كلمة المرور":
        reset_password_view()
    elif menu == "🚪 خروج":
        logout()
# تأكد من استيراد البيانات من Excel إن وُجد
if os.path.exists(EXCEL_PATH):
    import_excel_data()

# التحكم في حالة المستخدم
if not st.session_state.logged:
    login()
else:
    _, _, role = get_user_info(st.session_state.user)
    if role == "admin":
        admin_view()
    else:
        user_view()
# تذييل وحقوق ملكية البرنامج
st.markdown("""
<hr>
<p style='text-align: center; color: gray; font-size: small;'>
جميع الحقوق محفوظة © 2025 — تصميم وتطوير: <strong>Mohamed Shaban</strong> | Hospital Reports System
</p>
""", unsafe_allow_html=True)
