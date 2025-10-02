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
  *
from
  tmp_courtsession
order by
  SessionDate
  ,DisplayOrder
  ,StartTime
  ,SessionDescription
;
-- ------------------------------------------------------------------------------------------------
-- Tests
-- ------------------------------------------------------------------------------------------------
select
  *
from
  tmp_courtsession
where
  StartTime = ''
  and SessionDate = '2025-01-06'
order by
  SessionDate
  ,DisplayOrder
  ,StartTime
  ,SessionDescription
;
select
  cs1.SessionDescription
  ,cs2.SessionDescription
from
  tmp_courtsession cs1
  left outer join tmp_courtsession cs2
  on
    cs2.SessionDate = cs1.SessionDate
    and cs2.StartTime = cs1.StartTime
    and cs2.SessionDescription = cs1.SessionDescription
where
  cs1.StartTime between '2025-01-06' and '2025-01-10'
order by
  cs1.SessionDate
  ,cs1.DisplayOrder
  ,cs1.StartTime
  ,cs1.SessionDescription
;
-- ------------------------------------------------------------------------------------------------
select
  cs1.SessionDate
  ,cs1.SessionDescription
  ,cs1.SessionDate
  ,cs2.SessionDescription
from
  tmp_courtsession cs1
  left outer join tmp_courtsession cs2
  on
    cs2.SessionDescription = cs1.SessionDescription
where
    1=1
    and cs1.SessionDate between '2025-01-06' and '2025-01-10'
    and cs1.StartTime = ''
    and cs2.SessionDate between '2025-01-06' and '2025-01-10'
    and cs2.StartTime = ''
    and cs1.SessionDate <> cs2.SessionDate
;
-- ------------------------------------------------------------------------------------------------
select
  dayofweek(SessionDate) as weekday
  ,weekofyear(SessionDate) as yearweek
  ,SessionDate as SessionDate
  ,SessionDescription as SessionDescription
from
  tmp_courtsession
where
  StartTime - ''
;
select
  SessionDate
  ,StartTime
  ,SessionDescription
  ,Color
  ,JudicialOfficerCode
  ,DisplayOrder
  ,month(SessionDate) as month
  ,weekofyear(SessionDate) as yearweek
  ,dayofweek(SessionDate) as weekday
from
  tmp_courtsession
where
  1-1
  and StartTime - ''
  and month(SessionDate) - 1
  and weekofyear(SessionDate) - 5
--
order by
  month
  ,weekday
  ,DisplayOrder
;

