# testing_db_quick.py 20250917

"""
Test connecting to and querying a MS SQL database.

Uses the pacakes:
pyodbc

# For Windows Authentication (if you're on a trusted network):
# CONNECTION_STRING = (
#     "DRIVER={ODBC Driver 17 for SQL Server};"
#     "SERVER=YourServerName;"
#     "DATABASE=YourDatabaseName;"
#     "Trusted_Connection=yes;"
# )

"""

import pyodbc

# Replace with your actual connection details
DRIVER_NAME = 'SQL Server' # Verify your driver name
SERVER_NAME = 'fcvodsysqlprod\\GAFORSYTHPROD'
DATABASE_NAME = 'Justice'
USERNAME = 'YourUsername'
PASSWORD = 'YourPassword'

# Connection string using f-string for clarity
# For Windows Authentication (if you're on a trusted network):
CONNECTION_STRING = (
    f"DRIVER={DRIVER_NAME};"
    f"SERVER={SERVER_NAME};"
    f"DATABASE={DATABASE_NAME};"
    f"Trusted_Connection=yes;"
)

conn = None # Initialize connection

try:
    # 1. Connect to the database
    conn = pyodbc.connect(CONNECTION_STRING)
    cursor = conn.cursor()
    print("Successfully connected to the database.")
    
    # Test
    cursor.execute("select * from Justice.dbo.sCacheTable;")
    for row in cursor.fetchall():
        print(row)

except pyodbc.Error as ex:
    sqlstate = ex.args[0]
    print(f"Database Error: {sqlstate}")
    if conn:
        conn.rollback() # Rollback in case of an error

finally:
    # 5. Close the connection
    if conn:
        conn.close()
        print("Connection closed.")
