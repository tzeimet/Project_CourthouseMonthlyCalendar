# testing_db_sp_3.py 20250919

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

print(f"\nSuccessfully retrieved {len(df)} rows.")
print(df.head())
#===================================================================
# Crucial step: Ensure SessionDate is datetime and SORTED by both date AND description
df['SessionDate'] = pd.to_datetime(df['SessionDate'])
df = df.sort_values(by=['SessionDescription', 'SessionDate']).reset_index(drop=True)

custom_holidays = [
    pd.Timestamp('2024-07-04')
]

workday_calendar = CustomBusinessDay(
    weekmask='Mon Tue Wed Thu Fri',
    holidays=custom_holidays
)

# --- 1. IDENTIFY CONSECUTIVE DATES (Within the sorted DataFrame) ---

date_series_shifted = df['SessionDate'].shift(1)

expected_next_date = date_series_shifted.apply(
    lambda x: x + workday_calendar if pd.notna(x) else pd.NaT
)

# Boolean Series: True if the current date is consecutive to the previous one
is_date_consecutive = (df['SessionDate'] == expected_next_date)

# --- 2. IDENTIFY DESCRIPTION BREAKS ---

# True if the current description is DIFFERENT from the previous description
is_desc_change = (df['SessionDescription'] != df['SessionDescription'].shift(1))

# --- 3. CREATE A GROUP ID FOR CONSECUTIVE BLOCKS ---

# A break in the group sequence occurs if:
# 1. The session date is NOT consecutive to the previous one OR
# 2. The SessionDescription has changed from the previous row.
# We include the first row (index 0) as the start of the first group (handled by NaT).
group_break = (~is_date_consecutive) | is_desc_change

# Use .cumsum() to assign a unique ID to each group.
df['Group_ID'] = group_break.cumsum()


# --- 4. FILTER AND AGGREGATE TO GET DATE RANGES ---

# Count the size of each group ID
group_sizes = df.groupby('Group_ID').size()

# Get the IDs of groups that have 2 or more dates (consecutive)
long_group_ids = group_sizes[group_sizes >= 2].index

# Filter the main DataFrame to keep only the long consecutive groups
consecutive_df = df[df['Group_ID'].isin(long_group_ids)].copy()

# Calculate the Start Date, End Date, and keep the Session Description (using first/min/max is fine here)
date_ranges_with_desc = consecutive_df.groupby('Group_ID').agg(
    SessionDescription=('SessionDescription', 'first'), # Keep the common description
    Start_Date=('SessionDate', 'min'),
    End_Date=('SessionDate', 'max'),
    Count=('SessionDate', 'size')
).reset_index(drop=True)

print("\n--- Consecutive Session Date Ranges (Grouped by Description) ---")
print(date_ranges_with_desc)

date_ranges_with_desc['SessionDescription'].unique()