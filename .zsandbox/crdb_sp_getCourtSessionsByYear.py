# crdb_sp_getCourtSessionsByYear.py 20250922

import duckdb
from icecream import ic
import io
import pandas as pd
import pyodbc
from sqlalchemy import create_engine
import sys
import urllib

#==================================================================================================
# Connection details
DRIVER_NAME = 'SQL Server' # Verify your driver name
SERVER_NAME = 'fcvodsysqlprod\\GAFORSYTHPROD'
DATABASE_NAME = 'Justice'
USERNAME = 'YourUsername'
PASSWORD = 'YourPassword'

# Construct the ODBC connection string with Trusted_Connection=yes
odbc_conn_str = (
    f"DRIVER={{{DRIVER_NAME}}};"  # Note the double curly braces for the driver name
    f"SERVER={SERVER_NAME};"
    f"DATABASE={DATABASE_NAME};"
    f"Trusted_Connection=yes;"
)

# URL-encode the entire ODBC connection string
params = urllib.parse.quote_plus(odbc_conn_str)

# 3. Create the Database URI using the mssql+pyodbc dialect
# The format is: dialect+driver:///?odbc_connect=params
DB_URI = f"mssql+pyodbc:///?odbc_connect={params}"

# 4. Create the SQLAlchemy Engine
engine = create_engine(DB_URI)
print("SQLAlchemy Engine created successfully using Windows Authentication.")

STORED_PROC_NAME = "Justice.fc.sp_getCourtSessionsByYear"
PARAM_VALUE = 0

sql_query = f"""
SET NOCOUNT ON;
EXEC {STORED_PROC_NAME} @pMonth=?;
"""
#==================================================================================================
# Fetch Data from MS SQL into Pandas ---

try:
    print("Executing SP and fetching data into Pandas...")
    df = pd.read_sql(
        sql=sql_query,
        con=engine,      # Pass the SQLAlchemy Engine
        params=(PARAM_VALUE,) 
    )
    print(f"\nSuccessfully retrieved {len(df)} rows.")
    # Convert SessionDate column from varchar to datetime (TIMESTAMP_NS)
    df['SessionDate'] = pd.to_datetime(df['SessionDate'])
    print(df.head())

except Exception as e:
    print(f"Unkown exception: {e}")
    exit()

#==================================================================================================
# DuckDB DB name.
DUCKDB_FILENAME = "sp_getCourtSessionsByYear.db"
DUCKDB_TABLE_NAME = "courtsession"

# Load Data into In-Memory DuckDB ---
# Connect to an in-memory DuckDB database (:memory: is the key)
#con_duckdb = duckdb.connect(database=':memory:', read_only=False)
con_duckdb = duckdb.connect(database=DUCKDB_FILENAME, read_only=False)

# Load the Pandas DataFrame directly into DuckDB
# DuckDB creates a new virtual, temporary table named 'ms_sql_data' based on the DataFrame structure
con_duckdb.register("temp_df", df)
print(f"Data loaded into in-memory DuckDB table '{DUCKDB_TABLE_NAME}'.")

# Create permanent table using the structure of the registered DataFrame.
con_duckdb.sql(f"CREATE OR REPLACE TABLE {DUCKDB_TABLE_NAME} AS SELECT * FROM temp_df;")

# Verify and Query DuckDB Data ---
# You can now run high-performance OLAP queries on the in-memory data
result = con_duckdb.execute(f"""
    SELECT 
        COUNT(*) AS TotalRecords
    FROM {DUCKDB_TABLE_NAME};
""").fetchdf()
print("\nDuckDB Query Result:")
print(result)

con_duckdb.sql(f"describe {DUCKDB_TABLE_NAME};")

con_duckdb.sql(f"select count(*) from {DUCKDB_TABLE_NAME};")

# Find Sessions that have same SessionDescription.
con_duckdb.sql(f"""
select
    SessionDescription
    ,count(SessionDescription)
from
    {DUCKDB_TABLE_NAME}
group by
    SessionDescription
having
    count(SessionDescription) > 1
order by
    SessionDescription
;
"""
)
con_duckdb.sql(f"""
select
    StartTime
    ,SessionDescription
    ,count(SessionDescription)
from
    {DUCKDB_TABLE_NAME}
group by
    StartTime
    ,SessionDescription
having
    count(SessionDescription) > 1
order by
    StartTime
    ,SessionDescription
;
"""
)

# Close the DuckDB connection when done
con_duckdb.close()

#==================================================================================================
#==================================================================================================
import duckdb

# DuckDB DB name.
DUCKDB_FILENAME = "sp_getCourtSessionsByYear.db"

# Connect to DuckDB
con_duckdb = duckdb.connect(database=DUCKDB_FILENAME, read_only=False)

# Python List Data
# Each inner list/tuple is a row. The order determines the column index.
courtsession_mapping_list = [
    ['xxxxxxxxxxx','xxxxxxxxxxxxxxx',0]
    ,['SU CR SP SET JURY TRIAL','SUPERIOR COURT SPECIAL SET CRIMINAL JURY TRIAL (${JudicialOfficer}$-$CourtRoom}$)',0]
    ,['SU CR JURY TRIALS','SUPERIOR COURT CRIMINAL JURY TRIAL (${JudicialOfficer}$-$CourtRoom}$)',0]
    ,['SU CV SP SET BENCH TRIAL','SUPERIOR COURT SPECIAL SET CIVIL BENCH TRIAL (${JudicialOfficer}$-$CourtRoom}',0]
    ,['SU CV BENCH TRIALS','SUPERIOR COURT CIVIL BENCH TRIALS (${JudicialOfficer}$-$CourtRoom}$)',0]
    ,['SU CV JURY TRIALS','SUPERIOR COURT CIVIL JURY TRIALS (${JudicialOfficer}$-$CourtRoom}$)',0]
    ,['ST CR JURY TRIALS','STATE COURT CRIMINAL JURY TRIALS (${JudicialOfficer}$-$CourtRoom}$)',1]
    ,['ST CV JURY TRIALS','STATE COURT CIVIL JURY TRIALS (${JudicialOfficer}$-$CourtRoom}$)',1]
    ,['ST CV SP SET JURY TRIAL','STATE COURT SPECIAL SET CIVIL JURY TRIAL (${JudicialOfficer}$-$CourtRoom}$)',1]
    ,['ST CV BENCH TRIALS','STATE COURT CIVIL BENCH TRIALS (${JudicialOfficer}$-$CourtRoom}$)',1]
]

# Create the table schema first
DUCKDB_TABLE_NAME = "courtsession_mapping"
con_duckdb.sql(f"""
    drop table {DUCKDB_TABLE_NAME};
    CREATE TABLE {DUCKDB_TABLE_NAME} (
        OdysseyCourtSession VARCHAR
        ,CalendarFormat VARCHAR
        ,DisplayOrder int
    );
""")
con_duckdb.sql(f"describe {DUCKDB_TABLE_NAME};")

# Insert the entire list of records using the built-in VALUES clause
# DuckDB handles the list insertion automatically when passed as a parameter.
for record in courtsession_mapping_list:
    con_duckdb.execute(f"INSERT INTO {DUCKDB_TABLE_NAME} VALUES (?, ?, ?)", record)

# Verify
result = con_duckdb.sql(f"SELECT * FROM {DUCKDB_TABLE_NAME};").fetchdf()
print("Table created successfully:")
print(result)

# Close the database connection.
con_duckdb.close()

