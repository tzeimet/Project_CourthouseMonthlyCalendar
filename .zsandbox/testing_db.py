# testing_db.py 20250917

"""
Test connecting to and querying a MS SQL database.

Uses the pacakes:
pyodbc

CONNECTION_STRING = (
    "DRIVER={ODBC Driver 17 for SQL Server};"  # Use the exact driver name you have
    "SERVER=YourServerName\SQLEXPRESS;"        # e.g., 'localhost' or 'SERVERNAME\INSTANCE'
    "DATABASE=YourDatabaseName;"
    "UID=YourUsername;"                        # SQL Server login
    "PWD=YourPassword;"                        # SQL Server password
)
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

# The name of the stored procedure and its parameters
STORED_PROC_NAME = "YourStoredProcedureName"
PARAM1_VALUE = 101
PARAM2_VALUE = "New Data"

# ODBC call syntax for a stored procedure: {CALL procedure_name(?, ?)}
# Use a '?' for each parameter the stored procedure accepts.
SQL_CALL = f"{{CALL {STORED_PROC_NAME}(?, ?)}}"


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

    # 2. Execute the stored procedure
    # Pass the SQL_CALL string and a tuple of the parameter values
    cursor.execute(SQL_CALL, (PARAM1_VALUE, PARAM2_VALUE))
    print(f"Executed stored procedure: {STORED_PROC_NAME}")

    # 3. Handle results (if the stored procedure returns data)
    # Use fetchall(), fetchone(), or iterate over the cursor
    if cursor.description:
        rows = cursor.fetchall()
        print("\nStored Procedure Results:")
        for row in rows:
            print(row)
    
    # 4. Commit changes (if the stored procedure modifies data)
    conn.commit()
    print("\nTransaction committed.")

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
