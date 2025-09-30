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
  ,cs.StartTime as StartTime
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
