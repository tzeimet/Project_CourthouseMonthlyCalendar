# prep_court_session.py 20250930

from collections import defaultdict
from datetime import date, timedelta
import duckdb
from itertools import chain
import pandas as pd

#==================================================================================================
#==================================================================================================

ddb_conn = duckdb.connect(database=r'court_calendar_test.db')

sql_qry = """
-- ================================================================================================
drop table if exists tmp_courtsession;
CREATE TEMPORARY TABLE tmp_courtsession
(
  SessionDate date
  ,StartTime varchar
  ,SessionDescription varchar
  ,Color varchar
  ,JudicialOfficerCode varchar
  ,DisplayOrder int
);
insert into
  tmp_courtsession
select
  strftime(SessionDate,'%Y-%m-%d') as SessionDate
  ,case 
    when cs_m.CalendarFormat is not null
    then ''
    else cs.StartTime
  end as StartTime
  ,case 
    when cs_m.CalendarFormat is not null
    then
      replace(
        replace(
          cs_m.CalendarFormat
          ,'$CourtRoom}$'
          ,cs.CourtRoomCode
        )
        ,'${JudicialOfficer}$'
        ,cs.JudicialOfficerCode
      )
    else 
      concat(
        strftime(
          concat(
            '2025-01-01 '
            ,cs.StartTime
          )::datetime
          ,'%-I:%M '
        )
        ,regexp_replace(
          cs.SessionDescription
          ,'\\([A-Z]{3}\\) '
          ,''
        )
        ,' ('
        ,left(reverse(cs.JudicialOfficerCode),1)
        ,')'
      )
  end as SessionDescription
  ,j.Color as Color
  ,cs.JudicialOfficerCode as JudicialOfficerCode
  ,if(cs_m.DisplayOrder is null,3,cs_m.DisplayOrder) as DisplayOrder
from
  courtsession cs 
  left outer join judge j
  on
   j.OysseyCode = cs.JudicialOfficerCode
  left outer join courtsession_mapping cs_m 
  on
    cs.SessionDescription ilike concat(
      '%'
      ,replace(
        replace(
          cs_m.OdysseyCourtSession
          ,'-'
          ,''
        )
        ,' '
        ,'%'
      )
      ,'%'
    )
union
select
  strftime(Date,'%Y-%m-%d') as SessionDate
  ,'' as StartTime
  ,Name as SessionDescription
  ,Color as Color
  ,'' as JudicialOfficerCode
  ,DisplayOrder as DisplayOrder
from
  special_date
;
-- ================================================================================================
select
  SessionDate
  ,StartTime
  ,SessionDescription
  ,Color
  ,JudicialOfficerCode
  ,DisplayOrder
  ,strftime(Date,'%U')::int as week
from
  tmp_courtsession
order by
  SessionDate
  ,DisplayOrder
  ,StartTime
  ,SessionDescription
;
-- ================================================================================================
"""
court_session_list = ddb_conn.execute(sql_qry).fetchall()

ddb_conn.close()

week_sessions = []
for year_week in range(0,int(date(2025, 12, 31).strftime("%W"))+1):
    print(year_week)
    week_sessions = [session for session in court_session_list if int(session[0].strftime("%W")) == year_week]
    for s in [session for session in week_sessions if session[1] == '']:
        print(s)
    break

session_list = court_session_list
df = pd.DataFrame(
    session_list
    ,columns=[
        'SessionDate'
        ,'StartTime'
        ,'SessionDescription'
        ,'Color'
        ,'JudicialOfficerCode'
        ,'DisplayOrder'
    ]
)
# Ensure SessionDate is datetime type and extract WeekOfYear
df['SessionDate'] = pd.to_datetime(df['SessionDate'])
df['WeekOfYear'] = df['SessionDate'].dt.isocalendar().week

# Create a rank based on SessionDescription for each SessionDate group
df['session_rank'] = df.groupby('SessionDate')['SessionDescription'].rank(method='first').astype(int)

# Identify all unique descriptions across the ENTIRE dataset for final alignment
unique_descriptions = sorted(df['SessionDescription'].unique())

# Map the unique descriptions to a consistent rank for alignment across dates
description_rank_map = {desc: rank + 1 for rank, desc in enumerate(unique_descriptions)}

# Pivot to align descriptions: 
# Index: Week, Rank | Columns: Date | Values: Description
pivot_df = df.pivot_table(
    index=['WeekOfYear', 'session_rank'],
    columns='SessionDate',
    values='SessionDescription',
    aggfunc='first'
)

# Fill NaN values with the empty string '' (Requirement 1 & 2)
aligned_df = pivot_df.fillna('')

# Reset the index to get a final, clean table
final_df = aligned_df.reset_index(drop=True)

final_output = defaultdict(list)
dates = aligned_df.columns.tolist()

for week_year, group in aligned_df.groupby(level=0):
    weekly_output = []
    
    # Iterate through each aligned row (which represents the session index)
    for index, row in group.iterrows():
        # Build the list of tuples for this session index across all dates in the week
        session_index_row = []
        for date_col in dates:
            session_desc = row[date_col]
            
            # The row index in the 'aligned_df' IS the session_rank - 1
            if session_desc != '':
                session_index_row.append((date_col.date(), session_desc))
            else:
                # Add the tuple with the date and empty string '' (Requirement 2)
                session_index_row.append((date_col.date(), ''))
                
        # Only add to the final output if the date is one of the target dates
        weekly_output.append(session_index_row)

    final_output[week_year] = weekly_output


print("\n--- Example Output for Week 1 ---")
# Week 1: 2025-01-01, 2025-01-02, 2025-01-03
# The columns in the list represent the days with data.
print(final_output[1])


session_list_2 = [(session[0],session[1]) for session in session_list if session[1] == '']



sql = """
drop table if exists tmp_courtsession_2;
create table tmp_courtsession_2
as
select
  SessionDate
  ,StartTime
  ,SessionDescription
  ,Color
  ,JudicialOfficerCode
  ,DisplayOrder
  ,dayofweek(SessionDate)::int as dow
  ,strftime(SessionDate,'%U')::int as week
from
  tmp_courtsession
where
  StartTime = ''
order by
  SessionDate
  ,DisplayOrder
  ,StartTime
  ,SessionDescription
;
"""
# ------------------------------------------------------------------------------------------------+
def date_range_generator(start_date: date, end_date: date):
    """Generates a sequence of dates from start_date up to and including end_date."""
    # 1. Calculate the duration (a timedelta object)
    delta = end_date - start_date
    # 2. Iterate using range() for the number of days in the duration
    # delta.days gives the number of days, +1 makes the range inclusive of the end_date.
    for i in range(delta.days + 1):
        # 3. Yield the start_date plus the timedelta for the current iteration (i days)
        yield start_date + timedelta(days=i)
# ------------------------------------------------------------------------------------------------+

ddb_conn.sql(sql)

ddb_conn = duckdb.connect(database=r"court_calendar_test.db")

sql_qry="""
-- Get list of unique sessions in each week, sorting
-- and adding a row_num so to have ability to maintain 
-- same sort order.
select
  SessionDescription
  ,week
  ,row_number() over (order by week,DisplayOrder,SessionDescription) row_num
from
  (
    select distinct
      SessionDescription
      ,DisplayOrder
      ,week
    from
      courtsession_tmp
    where
      StartTime = ''
    order by
      week
      ,DisplayOrder
      ,SessionDescription
  ) z
  order by
    week
    ,DisplayOrder
    ,SessionDescription
;
"""
week_sessions = ddb_conn.execute(sql_qry).fetchall()

sql_qry = """
-- ================================================================================================
select
  SessionDate
  ,StartTime
  ,SessionDescription
  ,Color
  ,JudicialOfficerCode
  ,DisplayOrder
  ,strftime(SessionDate,'%U')::int as week
from
  courtsession_tmp
order by
  SessionDate
  ,DisplayOrder
  ,StartTime
  ,SessionDescription
;
-- ================================================================================================
"""
court_session_list = ddb_conn.execute(sql_qry).fetchall()

date_list = [s[0] for s in court_session_list]
first_week = int(min(date_list).strftime("%U"))
last_week = int(max(date_list).strftime("%U"))
sessions = []
for dt in date_range_generator(min(date_list),max(date_list)): # dt = date(2025,1,1)
    # Skip weekend dates.
    if dt.isoweekday() <= 5:
        date_sessions = [s for s in court_session_list if s[0] == dt and s[1] == '']
        # From sorted list of all the sessions of the given week,
        # clear those that are not in the current date's list. 
        # For those that are, add any needed additional data, such as color.
        for week_session in [ws for ws in week_sessions if ws[1] == int(dt.strftime("%U"))]:
            # Get matching date session to current week session, if exists.
            matching_session = None
            try:
                matching_session = next(
                    date_session for date_session in date_sessions if week_session[0] == date_session[2]
                )
            except StopIteration:
                pass
            # Using SessionDescription
            if matching_session:
                # Add matching session with week session row_num to maintain order.
                sessions.append(list(matching_session)+[week_session[2]])
            else:
                # Add blank session with week session row_num to maintain order.
                # [SessionDate,StatrDate='',SessionDescription='',Color='',DisplayOrder='',row_num]
                sessions.append([dt,'','','','',week_session[2]])

for s in sessions:
    print(s)
