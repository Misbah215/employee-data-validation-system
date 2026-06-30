import sqlite3
import pandas as pd

conn = sqlite3.connect("system.db")

print("\n🔍 ERROR LOGS:\n")
print(pd.read_sql("SELECT * FROM error_logs", conn))

conn.close()