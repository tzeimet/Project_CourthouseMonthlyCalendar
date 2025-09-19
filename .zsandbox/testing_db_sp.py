# testing_db_sp.py 20250917

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
STORED_PROC_NAME = "Justice.fc.sp_getCourtSessionsByYear"
PMONTH = 9

SQL_CALL = f"{{CALL {STORED_PROC_NAME} (?)}}"

# Use a '?' for each parameter the stored procedure accepts.
sql = """\
SET NOCOUNT ON;
EXEC Justice.fc.sp_getCourtSessionsByYear @pMonth=?;
"""

conn = None # Initialize connection

try:
    # 1. Connect to the database
    conn = pyodbc.connect(CONNECTION_STRING)
    cursor = conn.cursor()
    print("Successfully connected to the database.")
    
     # 2. Execute the stored procedure
    #cursor.execute(sql, (9))
    cursor.execute(SQL_CALL, (9))
    print(f"Executed stored procedure: {STORED_PROC_NAME}")

    # 3. Handle results (if the stored procedure returns data)
    # Use fetchall(), fetchone(), or iterate over the cursor
    while True:
        if cursor.description:
            #print(cursor.description)
            print("\nStored Procedure Results:")
            for row in cursor.fetchall():
                print(row)
        if not cursor.nextset():
            break

#    # 4. Commit changes (if the stored procedure modifies data)
#    conn.commit()
#    print("\nTransaction committed.")

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
