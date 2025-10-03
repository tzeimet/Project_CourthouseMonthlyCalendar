# get_sessions_testing.py 20251002

from collections import defaultdict
from datetime import date, timedelta
import duckdb
from itertools import chain
import pandas as pd

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
#ddb_conn = duckdb.connect(database=':memory:')
ddb_conn = duckdb.connect(database=r'.zsandbox\court_calendar.db')
# ------------------------------------------------------------------------------------------------+

sql_qry="""
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
  ,Week int
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
  ,strftime(SessionDate,'%U')::int as Week
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
  ,strftime(Date,'%U')::int as Week
from
  special_date
;
-- ================================================================================================
-- Get list of unique sessions in each week, sorting
-- and adding a row_num so to have ability to maintain 
-- same sort order.
select
  SessionDescription
  ,Week
  ,row_number() over (order by Week,DisplayOrder,SessionDescription) row_num
from
  (
    select distinct
      SessionDescription
      ,DisplayOrder
      ,Week
    from
      tmp_courtsession
    where
      StartTime = ''
    order by
      Week
      ,DisplayOrder
      ,SessionDescription
  ) z
  order by
    Week
    ,DisplayOrder
    ,SessionDescription
;
-- ================================================================================================
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
  tmp_courtsession
order by
  SessionDate
  ,DisplayOrder
  ,StartTime
  ,SessionDescription
;
-- ================================================================================================
"""
court_session_list = []
sessions = []

court_session_list = ddb_conn.execute(sql_qry).fetchall()

date_list = [s[0] for s in court_session_list]
breakpoint()
for dt in date_range_generator(min(date_list),max(date_list)):
    # Skip weekend dates.
    if dt.isoweekday() <= 5:
        date_sessions = [s for s in court_session_list if s[0] == dt and s[1] == '']
        # From sorted list of all the sessions of the given week,
        # clear those that are not in the current date's list. 
        # For those that are, add any needed additional data, such as color.
        for week_session in [ws for ws in week_sessions if ws[1] == int(dt.strftime("%U"))]:
            # Get matching date session to current week session, if exists.
            matching_date_session = None
            try:
                matching_date_session = next(
                    date_session for date_session in date_sessions if week_session[0] == date_session[2]
                )
            except StopIteration:
                pass
            # Using SessionDescription
            if matching_date_session:
                # Add matching session with week session row_num to maintain order.
                # [SessionDate,StatrDate,SessionDescription,Color,JudicialOfficerCode,DisplayOrder,week,row_num]
                sessions.append(list(matching_date_session)+[week_session[2]])
            else:
                # Add blank session with week session row_num to maintain order.
                # [SessionDate,StatrDate='',SessionDescription='',Color='',JudicialOfficerCode='',DisplayOrder=,week,row_num]
                sessions.append([dt,'','','','',999,week_session[1],week_session[2]])

for s in sessions:
    print(s)
