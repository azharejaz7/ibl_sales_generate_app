import pyodbc

server = "192.168.10.2,1433"
database = "PS_Trade"
username = "sa"
password = "29031982"
driver = "ODBC Driver 17 for SQL Server"

try:
    conn = pyodbc.connect(f"DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}")
    print("✅ Connection successful!")
except Exception as e:
    print(f"❌ Connection failed: {e}")
