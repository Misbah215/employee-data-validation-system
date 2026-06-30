import pandas as pd
import os
import win32com.client as win32
import sqlite3
from datetime import datetime
from email_validator import validate_email, EmailNotValidError

# ---------------- DATABASE SETUP ----------------
conn = sqlite3.connect("system.db")
cursor = conn.cursor()

# ✅ Create tables
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

# ✅ Clear logs (clean run)
cursor.execute("DELETE FROM error_logs")
conn.commit()


# ---------------- LOAD EXCEL (ONE-TIME) ----------------
def load_excel_to_db():
    df = pd.read_excel("input.xlsx", sheet_name="EmployeeData")
    df.to_sql("employees", conn, if_exists="replace", index=False)
    print("✅ Excel data loaded into DB")


# ---------------- LOG FUNCTION ----------------
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

I hope this message finds you well.

Please find attached the employee report for your team. The data has been validated and processed, and any invalid records were logged in the system for review.

Kindly review the report and let us know if any updates are required.

Best regards,  
HR Team
"""

        mail.Attachments.Add(os.path.abspath(file_path))
        mail.Send()

        print(f"✅ Email sent to {to_email}")

    except Exception as e:
        log_error("EMAIL_ERROR", f"{to_email} -> {e}")
        print(f"❌ Email failed for {to_email}")


# ---------------- MAIN ----------------

# ✅ STEP 1: OPTIONAL (run only first time)
#load_excel_to_db()

# ✅ STEP 2: READ FROM DATABASE (MAIN CHANGE 🔥)
df = pd.read_sql("SELECT * FROM employees", conn)

print("✅ Data fetched from database")

# ✅ Required columns
required_cols = ["EmpID", "EmpName", "ManagerName", "ManagerEmail"]

for col in required_cols:
    if col not in df.columns:
        log_error("DATA_ERROR", f"Missing column: {col}")
        exit()

valid_rows = []

# ✅ Check missing data
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

# ✅ Validate email + domain restriction
for idx, row in df_clean.iterrows():
    email = row["ManagerEmail"]

    # Format check
    if not is_valid_email(email):
        log_error("VALIDATION_ERROR", email)
        continue

    # Domain restriction
    if not email.endswith("@shell.com"):
        log_error("INVALID_DOMAIN", email)
        continue

    final_rows.append(row)

final_df = pd.DataFrame(final_rows)

# ✅ Group data
grouped = final_df.groupby("ManagerEmail")

print(f"✅ Managers found: {len(grouped)}")

# ✅ Folder
folder = "reports_" + datetime.now().strftime("%Y%m%d")
os.makedirs(folder, exist_ok=True)

# ✅ Process
for email, group in grouped:
    manager_name = group["ManagerName"].iloc[0]

    file_path = os.path.join(folder, f"{manager_name}.xlsx")
    group.to_excel(file_path, index=False)

    print(f"✅ File created for {manager_name}")

    send_email(email, manager_name, file_path)

# ✅ Summary
error_count = pd.read_sql(
    "SELECT COUNT(*) as count FROM error_logs", conn)["count"][0]

summary = pd.DataFrame([{
    "Total Employees Processed": len(final_df),
    "Total Managers": len(grouped),
    "Total Errors": error_count
}])

summary.to_excel("summary.xlsx", index=False)

print("✅ Summary report generated")

# ✅ Close DB
conn.close()