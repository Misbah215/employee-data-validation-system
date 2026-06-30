import pandas as pd
import os
import win32com.client as win32
import sqlite3
from datetime import datetime
from email_validator import validate_email, EmailNotValidError

# ---------------- DB SETUP ----------------
conn = sqlite3.connect("system.db")
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS employees (
    EmpID TEXT,
    EmpName TEXT,
    ManagerName TEXT,
    ManagerEmail TEXT
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS error_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    timestamp TEXT,
    error_type TEXT,
    message TEXT
)
""")

conn.commit()

cursor.execute("DELETE FROM error_logs")
conn.commit()

# ---------------- LOAD EXCEL (RUN ONLY ONCE) ----------------
def load_excel_to_db():
    df = pd.read_excel("input.xlsx", sheet_name="EmployeeData")
    df.to_sql("employees", conn, if_exists="replace", index=False)
    print("✅ Excel loaded into DB")

# ---------------- LOG ----------------
def log_error(error_type, message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute(
        "INSERT INTO error_logs (timestamp, error_type, message) VALUES (?, ?, ?)",
        (timestamp, error_type, message)
    )
    conn.commit()

# ---------------- EMAIL VALIDATION ----------------
def is_valid_email(email):
    try:
        validate_email(email)
        return True
    except EmailNotValidError:
        return False

# ---------------- SEND EMAIL ----------------
def send_email(to_email, manager_name, file_path):
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        mail.To = to_email
        mail.Subject = f"Employee Report - {manager_name}"

        mail.Body = f"""Hello {manager_name},

Please find attached your team’s employee report.
Only validated records were processed. Invalid entries were logged.

Regards,
HR Team
"""

        mail.Attachments.Add(os.path.abspath(file_path))
        mail.Send()

        print(f"✅ Email sent to {to_email}")

    except Exception as e:
        log_error("EMAIL_ERROR", f"{to_email} -> {e}")

# ---------------- MAIN ----------------

# ✅ RUN THIS ONLY FIRST TIME
#load_excel_to_db()

df = pd.read_sql("SELECT * FROM employees", conn)

df.columns = df.columns.str.strip()

print("✅ Data fetched from database")
print("Columns:", df.columns)

required_cols = ["EmpID", "EmpName", "ManagerName", "ManagerEmail"]

for col in required_cols:
    if col not in df.columns:
        print(f"❌ Missing column: {col}")
        exit()

valid_rows = []

# ✅ MISSING DATA CHECK
for idx, row in df.iterrows():
    is_valid_row = True
    for col in required_cols:
        if pd.isna(row[col]):
            log_error("MISSING_DATA", f"Row {idx} missing {col}")
            is_valid_row = False
    if is_valid_row:
        valid_rows.append(row)

df_clean = pd.DataFrame(valid_rows)

final_rows = []

# ✅ EMAIL + DOMAIN VALIDATION
for _, row in df_clean.iterrows():
    email_raw = row["ManagerEmail"]

    if pd.isna(email_raw):
        log_error("MISSING_DATA", "Email missing")
        continue

    email = str(email_raw).strip().lower()

    print("Checking email:", email)

    if not is_valid_email(email):
        log_error("VALIDATION_ERROR", email)
        continue

    if not email.endswith("@shell.com"):
        log_error("INVALID_DOMAIN", email)
        continue

    row["ManagerEmail"] = email
    final_rows.append(row)

final_df = pd.DataFrame(final_rows)

if final_df.empty:
    print("❌ No valid data to process")
    conn.close()
    exit()

grouped = final_df.groupby("ManagerEmail")

print(f"✅ Managers found: {len(grouped)}")

folder = "reports_" + datetime.now().strftime("%Y%m%d")
os.makedirs(folder, exist_ok=True)

for email, group in grouped:
    manager_name = group["ManagerName"].iloc[0]

    file_path = os.path.join(folder, f"{manager_name}.xlsx")
    group.to_excel(file_path, index=False)

    print(f"✅ File created for {manager_name}")

    send_email(email, manager_name, file_path)

error_count = pd.read_sql("SELECT COUNT(*) as c FROM error_logs", conn)["c"][0]

summary = pd.DataFrame([{
    "Total Employees Processed": len(final_df),
    "Total Managers": len(grouped),
    "Total Errors": error_count
}])

summary.to_excel("summary.xlsx", index=False)

print("✅ Summary report generated")

conn.close()
