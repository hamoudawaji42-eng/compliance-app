import os
import re
import json
import psycopg2
import pandas as pd
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, abort

app = Flask(__name__)

# تحديد المسار المطلق للملفات لضمان عملها بشكل صحيح على خوادم النشر (Render)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, "data.xlsx")
# رابط قاعدة بيانات Supabase (يتخزن في Render كمتغير باسم DATABASE_URL)
DATABASE_URL = os.environ.get("DATABASE_URL")


def get_conn():
    """فتح اتصال بقاعدة بيانات Supabase"""
    return psycopg2.connect(DATABASE_URL, sslmode="require")

def init_db():
    """إنشاء قاعدة البيانات وجدول العملاء والسجلات"""
    conn = get_conn()
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS audit_history (
            id SERIAL PRIMARY KEY,
            client_name TEXT NOT NULL,
            company_name TEXT NOT NULL,
            phone TEXT NOT NULL,
            company_tier TEXT,
            violations_count INTEGER,
            total_fine DOUBLE PRECISION,
            compliance_rate TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            violations_detail TEXT
        )
    ''')
    conn.commit()
    cursor.close()
    conn.close()

init_db()

def save_to_db(client_name, company_name, phone, company_tier, total_fine, selected_count, compliance_rate, violations_detail):
    """حفظ سجل العميل في قاعدة البيانات مع تفاصيل المخالفات المختارة"""
    try:
        tier_names = {'a': 'فئة (أ): 50 عامل فأعلى', 'b': 'فئة (ب): 21 إلى 49 عامل', 'c': 'فئة (ج): 20 عامل فأقل'}
        conn = get_conn()
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO audit_history 
            (client_name, company_name, phone, company_tier, violations_count, total_fine, compliance_rate, violations_detail)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
        ''', (client_name, company_name, phone, tier_names.get(company_tier, company_tier), 
              selected_count, total_fine, f"{compliance_rate}%",
              json.dumps(violations_detail, ensure_ascii=False)))
        conn.commit()
        cursor.close()
        conn.close()
    except Exception as e:
        print(f"Error saving to DB: {e}")

def get_clean_violations(company_tier="a"):
    if not os.path.exists(EXCEL_FILE):
        return []

    try:
        xls = pd.ExcelFile(EXCEL_FILE)
        violations = []
        v_id = 1

        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)

            for idx, row in df.iterrows():
                row_str = " ".join([str(val).strip() for val in row.values if pd.notna(val)])

                if any(phrase in row_str for phrase in ["الموقع", "إن وزير", "الجهة المُصدرة", "الغرامة المالية بالريال", "وصف المخالفة تصنيف", "نوع الوثيقة", "رقم الصادر", "تاريخ الصادر"]):
                    continue

                match_start_num = re.match(r"^\s*(\d+)\s+", row_str)

                if match_start_num:
                    v_num = int(match_start_num.group(1))
                    if v_num > 500:
                        continue

                    raw_numbers = re.findall(r"\b\d{3,6}\b", row_str)
                    valid_fines = []
                    for num in raw_numbers:
                        val = int(num)
                        if val != v_num and val not in valid_fines and val in [300, 500, 1000, 1500, 2000, 2500, 3000, 5000, 10000, 15000, 20000, 25000, 50000]:
                            valid_fines.append(val)

                    valid_fines.sort()

                    fine_amount = 500
                    if valid_fines:
                        if company_tier == "a":
                            fine_amount = valid_fines[-1]
                        elif company_tier == "b":
                            fine_amount = valid_fines[len(valid_fines) // 2]
                        else:
                            fine_amount = valid_fines[0]

                    has_warning = "إنذار" in row_str

                    violations.append({
                        "id": v_id,
                        "category": sheet,
                        "title": row_str,
                        "fine": fine_amount,
                        "warning": has_warning,
                    })
                    v_id += 1

        return violations
    except Exception as e:
        print(f"Error: {e}")
        return []

@app.route("/", methods=["GET", "POST"])
def index():
    company_tier = request.form.get("company_tier", "a")
    violations = get_clean_violations(company_tier)
    total_available = len(violations)

    selected_ids = []
    total_fine = 0
    selected_count = 0
    compliance_rate = 100.0

    client_name = request.form.get("client_name", "")
    company_name = request.form.get("company_name", "")
    phone = request.form.get("phone", "")

    if request.method == "POST" and "calculate" in request.form:
        selected_ids = [int(x) for x in request.form.getlist("violations")]
        selected_count = len(selected_ids)

        selected_violations = []
        for v in violations:
            if v["id"] in selected_ids:
                total_fine += v["fine"]
                selected_violations.append(v)

        if total_available > 0:
            compliance_rate = round(((total_available - selected_count) / total_available) * 100, 1)

        if client_name and phone:
            save_to_db(client_name, company_name, phone, company_tier, total_fine, selected_count, compliance_rate, selected_violations)

    return render_template(
        "index.html",
        violations=violations,
        selected_ids=selected_ids,
        total_fine=total_fine,
        selected_count=selected_count,
        compliance_rate=compliance_rate,
        total_available=total_available,
        company_tier=company_tier,
        client_name=client_name,
        company_name=company_name,
        phone=phone
    )

@app.route("/clients")
def clients_list():
    """صفحة لاستعراض والبحث في تقارير العملاء المخزنة"""
    search_query = request.args.get("search", "")
    conn = get_conn()
    cursor = conn.cursor()

    if search_query:
        cursor.execute('''
            SELECT id, client_name, company_name, phone, company_tier, violations_count, total_fine, compliance_rate,
                   to_char(created_at, 'YYYY-MM-DD HH24:MI:SS')
            FROM audit_history 
            WHERE client_name ILIKE %s OR company_name ILIKE %s OR phone ILIKE %s
            ORDER BY created_at DESC
        ''', (f"%{search_query}%", f"%{search_query}%", f"%{search_query}%"))
    else:
        cursor.execute('''
            SELECT id, client_name, company_name, phone, company_tier, violations_count, total_fine, compliance_rate,
                   to_char(created_at, 'YYYY-MM-DD HH24:MI:SS')
            FROM audit_history 
            ORDER BY created_at DESC
        ''')

    records = cursor.fetchall()
    cursor.close()
    conn.close()
    return render_template("clients.html", records=records, search_query=search_query)

@app.route("/report/<int:record_id>")
def view_report(record_id):
    """عرض تقرير محفوظ سابقاً بنفس المخالفات وقت الفحص، جاهز للطباعة/التصدير"""
    conn = get_conn()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT id, client_name, company_name, phone, company_tier, violations_count, total_fine, compliance_rate,
               to_char(created_at, 'YYYY-MM-DD HH24:MI:SS'), violations_detail
        FROM audit_history WHERE id = %s
    ''', (record_id,))
    row = cursor.fetchone()
    cursor.close()
    conn.close()

    if not row:
        abort(404)

    violations_detail = []
    if row[9]:
        try:
            violations_detail = json.loads(row[9])
        except Exception:
            violations_detail = []

    record = {
        "id": row[0],
        "client_name": row[1],
        "company_name": row[2],
        "phone": row[3],
        "company_tier": row[4],
        "violations_count": row[5],
        "total_fine": row[6],
        "compliance_rate": row[7],
        "created_at": row[8],
    }

    return render_template("report.html", record=record, violations=violations_detail)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
