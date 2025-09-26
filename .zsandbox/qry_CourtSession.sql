-- qry_CourtSession.sql 20250922

describe courtsession;

select
  *
from
  courtsession_mapping cs_m 
;
select
  *
from
  courtsession cs 
where
  year(cs.SessionDate) = 2025 and month(SessionDate) = 1
;

select * from CourtSession cs
where cs.SessionDescription ilike '%su%cv%bench%' and year(cs.SessionDate) = 2025 and month(SessionDate) = 1 order by cs.SessionDate;

select
  cs_m.OdysseyCourtSession
  ,replace(
    replace(
      cs_m.OdysseyCourtSession
      ,'-'
      ,''
    )
    ,' '
    ,'%'
  ) as test_str
from
  courtsession_mapping cs_m
;
-- ================================================================================================
drop table if exists tmp_courtsessions;
CREATE TEMPORARY TABLE tmp_courtsessions AS
select
  strftime(SessionDate,'%Y-%m-%d') as SessionDate
  ,StartTime as StartTime
  ,SessionDescription as SessionDescription_orig
  ,cs.JudicialOfficerCode
  ,replace(
    replace(
      CalendarFormat
      ,'$CourtRoom}$'
      ,CourtRoomCode
    )
    ,'${JudicialOfficer}$'
    ,cs.JudicialOfficerCode
  ) as SessionDescription
  ,DisplayOrder as DisplayOrder
from
  courtsession cs 
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
where
  year(cs.SessionDate) = 2025 and month(SessionDate) = 1
;
-- select * from courtsession;
-- select * from tmp_courtsessions;
select
    cs.SessionDate
    ,cs.StartTime
    ,cs.SessionDescription_orig
    ,cs.SessionDescription
    ,cs.JudicialOfficerCode
from
  (
    select
      SessionDate
      ,StartTime
      ,SessionDescription_orig
      ,SessionDescription
      ,JudicialOfficerCode
      ,0 as DisplayOrder0
      ,DisplayOrder
    from
      tmp_courtsessions
    where
      SessionDescription is not null
    union
    select
      SessionDate
      ,StartTime
      ,SessionDescription_orig
      ,concat(
        strftime(
          concat(
            '2025-01-01 '
            ,StartTime
          )::datetime
          ,'%-I:%M '
        )
        ,regexp_replace(
          SessionDescription_orig
          ,'\([A-Z]{3}\) '
          ,''
        )
        ,' ('
        ,left(reverse(JudicialOfficerCode),1)
        ,')'
      ) as SessionDescription
      ,JudicialOfficerCode
      ,1 as DisplayOrder0
      ,DisplayOrder
    from
      tmp_courtsessions
    where
      SessionDescription is null
  ) cs
order by
  cs.SessionDate
  ,cs.StartTime
  ,cs.DisplayOrder0
  ,cs.DisplayOrder
  ,cs.SessionDescription
;
