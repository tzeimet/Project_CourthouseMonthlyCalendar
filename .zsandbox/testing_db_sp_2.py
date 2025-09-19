
import duckdb
from icecream import ic
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import pandas as pd
import pandas as pd
from pandas.tseries.offsets import CustomBusinessDay # Keep this for the calendar definition
from sqlalchemy import create_engine
import sys
import urllib

# Connection details
DRIVER_NAME = 'SQL Server' # Verify your driver name
SERVER_NAME = 'fcvodsysqlprod\\GAFORSYTHPROD'
DATABASE_NAME = 'Justice'

# Connection string using f-string for clarity
# For Windows Authentication (if you're on a trusted network):
CONNECTION_STRING = (
    f"DRIVER={DRIVER_NAME};"
    f"SERVER={SERVER_NAME};"
    f"DATABASE={DATABASE_NAME};"
    f"Trusted_Connection=yes;"
)

# 1. Construct the ODBC connection string with Trusted_Connection=yes
odbc_conn_str = (
    f"DRIVER={{{DRIVER_NAME}}};"  # Note the double curly braces for the driver name
    f"SERVER={SERVER_NAME};"
    f"DATABASE={DATABASE_NAME};"
    f"Trusted_Connection=yes;"
)

# 2. URL-encode the entire ODBC connection string
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

#try:
df = pd.read_sql(
    sql=sql_query,
    con=engine,      # Pass the SQLAlchemy Engine
    params=(PARAM_VALUE,) 
)
df['SessionDate'] = pd.to_datetime(df['SessionDate'])
#df['StartTime'] = pd.to_datetime(df['StartTime']).dt.time

print(f"\nSuccessfully retrieved {len(df)} rows.")
print(df.head())

# 1. Define your custom holiday list (e.g., Good Friday, Memorial Day)
# You should replace this with your actual list of observed holidays
custom_holidays = [
    pd.Timestamp('2025-03-29') # Example: Good Friday (a date in your sample data)
    ,pd.Timestamp('2025-05-27') # Example: Memorial Day (a Monday)
    ,pd.Timestamp('2025-07-04')  # Example: Independence Day
    ,pd.Timestamp('2025-12-25')  # Example: Christmas Day
    ,pd.Timestamp('2025-12-31')  # Example: New Year's Eve Day
]    

# 2. Define the Custom Business Day offset
workday_calendar = CustomBusinessDay(
    weekmask='Mon Tue Wed Thu Fri',
    holidays=custom_holidays
)

# 3. Calculate the date that is one custom workday after the previous date
# .shift(1) moves the SessionDate column down by one row (i.e., the previous row's date)
expected_next_date = df['SessionDate'].shift(1) + workday_calendar

# 4. Compare the current SessionDate with the expected next date
# This creates a boolean Series where True means the dates are consecutive workdays.
is_consecutive = (df['SessionDate'] == expected_next_date)

# Since the very first row has no predecessor, its value will be False (or NaT/NaN comparison)
# You only need to identify the *second* row in the consecutive pair, so the calculation is correct.

# 5. Filter the DataFrame to show only the rows that are the *second* day of a consecutive pair
# Note: You can also use .shift(-1) to identify the *first* day of the pair.
consecutive_rows = df[is_consecutive]

print("Rows that are consecutive (The second date in the pair):")
print(consecutive_rows)

except Exception as e:
    print(f"\nAn error occurred: {e}")
finally:
    engine.dispose()

