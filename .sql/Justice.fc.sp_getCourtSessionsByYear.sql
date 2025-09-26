-- Justice.fc.sp_getCourtSessionsByYear.sql 20250917

USE [Justice]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET NOCOUNT ON
go

alter procedure fc.sp_getCourtSessionsByYear
-- declare
  @pMonthOrYear int = 9  -- @pMonthOrYear = month for testing.
as
begin
  SET NOCOUNT ON;
  declare 
    @vBeginDate date
    ,@vEndDate date

  -- Check if @pMonthOrYear if current year s/b used.r
  if (
    @pMonthOrYear is null 
    or @pMonthOrYear < 1 
    or @pMonthOrYear > 12
  )
  and @pMonthOrYear < 2024
  begin
    select
      @vBeginDate = datefromparts(year(getdate()),12,1)
      ,@vEndDate = datefromparts(year(getdate()),12,31)
  end -- if a 4-digit year
  else if @pMonthOrYear >= 2024
  begin
    select
      @vBeginDate = datefromparts(@pMonthOrYear,1,1)
      ,@vEndDate = datefromparts(@pMonthOrYear,12,31)
  end
  else
  begin
    select
      @vBeginDate = datefromparts(year(getdate()),@pMonthOrYear,1)
      ,@vEndDate = eomonth(@vBeginDate)
  end -- else

  drop table if exists #courtroom;
  create table #courtroom
  (
    CourtRoomCode varchar(10)
  );
  insert into #courtroom
  (
    CourtRoomCode
  )
   values
   ('401')
  ,('402')
  ,('501')
  ,('502')
  ,('503')
  ,('504')
  ,('JAR')
  ;
  -- select CourtRoomCode from #courtroom
  drop table if exists #CtSessionCourtHouse
  select
    --CtS.CourtSessionID
    cast(CtS.SessionDate as Date) as SessionDate
    ,left(cast(CtS.StartTime as time),5) as StartTime
    --,CtS.EndTime
    ,CtS.Description as SessionDescription
    ,CtS.Deleted as Deleted
    ,Cts.CalendarID as CalendarID
    ,uC_Clndr.Description CalendarDescription
    --
    ,CtS.TimestampCreate as CtS_TimestampCreate
    ,CtS.UserIDCreate as CtS_CreateUserID
    ,concat(AU_CtS_Create.NameFirst,' ',AU_CtS_Create.NameLast) as CtS_CreateUserName
    ,CtS.TimestampChange as CtS_TimestampChange
    ,CtS.UserIDChange as CtS_ChangeUserID
    ,concat(AU_CtS_Change.NameFirst,' ',AU_CtS_Change.NameLast) as CtS_ChangeUserName
    ,CtS.TimestampMajorChange as CtS_TimestampMajorChange
    ,S.TimestampCreate as S_TimestampCreate
    ,S.TimestampChange as S_TimestampChange
    ,'|||' as x
    --
    ,L.CalendarResourceTypeID as LocationTypeID
    ,L.CalendarResourceTypeCode as LocationTypeCode
    ,L.CalendarResourceTypeDescription as LocationTypeDescription
    ,L.CalendarResourceID as LocationID
    ,L.CalendarResourceCode as LocationCode
    ,L.CalendarResourceDescription as LocationDescription
    --
    ,JO.CalendarResourceTypeID as JudicialOfficerTypeID
    ,JO.CalendarResourceTypeCode as JudicialOfficerTypeCode
    ,JO.CalendarResourceTypeDescription as JudicialOfficerTypeDescription
    ,JO.CalendarResourceID as JudicialOfficerID
    ,JO.CalendarResourceCode as JudicialOfficerCode
    ,JO.CalendarResourceDescription as JudicialOfficerDescription
    --
    ,CTRM.CalendarResourceTypeID as CourtRoomTypeID
    ,CTRM.CalendarResourceTypeCode as CourtRoomTypeCode
    ,CTRM.CalendarResourceTypeDescription as CourtRoomTypeDescription
    ,CTRM.CalendarResourceID as CourtRoomID
    ,CTRM.CalendarResourceCode as CourtRoomCode
    ,CTRM.CalendarResourceDescription as CourtRoomDescription
  into #CtSessionCourtHouse
  from
    Justice.dbo.CtSession CtS
    inner join Justice.dbo.CtSessionBlock CtSB 
    on
      CtSB.CourtSessionID = CtS.CourtSessionID
    left outer join Justice.dbo.Setting S 
    on
      S.CourtSessionBlockID = CtSB.CourtSessionBlockID
    --
    left outer join Operations.dbo.AppUser AU_CtS_Create
    on
      AU_CtS_Create.UserID = CtS.UserIDCreate
    left outer join Operations.dbo.AppUser AU_CtS_Change
    on
      AU_CtS_Change.UserID = CtS.UserIDChange
    --
    inner join Justice.dbo.uCode uC_Clndr
    on
      uC_Clndr.CodeID = Cts.CalendarID
    -- select top 1 * from Justice.dbo.CtSession
    left outer join 
    (
      select
        CtSCRB.CourtSessionID
        ,CtSCR.CalendarResourceBucketID
        ,CtSCRB.CalendarResourceTypeID as CalendarResourceTypeID
        ,uC_CRT.Code as CalendarResourceTypeCode
        ,uC_CRT.Description as CalendarResourceTypeDescription
        ,CtSCR.CalendarResourceID as CalendarResourceID
        ,uC_CR.Code as CalendarResourceCode
        ,uC_CR.Description as CalendarResourceDescription
      from
        Justice.dbo.CtSessionClndrResBucket CtSCRB
        inner join Justice.dbo.uCode uC_CRT
        on
          uC_CRT.CodeID = CtSCRB.CalendarResourceTypeID
          and uC_CRT.Code = 'L'
        inner join Justice.dbo.xCtSessionClndrRes CtSCR
        on
          CtSCR.CalendarResourceBucketID = CtSCRB.CalendarResourceBucketID
        inner join Justice.dbo.uCode uC_CR
        on
          uC_CR.CodeID = CtSCR.CalendarResourceID
    ) L
    on
      L.CourtSessionID = CtS.CourtSessionID
    --
    left outer join 
    (
      select
        CtSCRB.CourtSessionID
        ,CtSCR.CalendarResourceBucketID
        ,CtSCRB.CalendarResourceTypeID as CalendarResourceTypeID
        ,uC_CRT.Code as CalendarResourceTypeCode
        ,uC_CRT.Description as CalendarResourceTypeDescription
        ,CtSCR.CalendarResourceID as CalendarResourceID
        ,uC_CR.Code as CalendarResourceCode
        ,uC_CR.Description as CalendarResourceDescription
      from
        Justice.dbo.CtSessionClndrResBucket CtSCRB
        inner join Justice.dbo.uCode uC_CRT
        on
          uC_CRT.CodeID = CtSCRB.CalendarResourceTypeID
          and uC_CRT.Code = 'JO'
        inner join Justice.dbo.xCtSessionClndrRes CtSCR
        on
          CtSCR.CalendarResourceBucketID = CtSCRB.CalendarResourceBucketID
        inner join Justice.dbo.uCode uC_CR
        on
          uC_CR.CodeID = CtSCR.CalendarResourceID
    ) JO
    on
      JO.CourtSessionID = CtS.CourtSessionID
    --
    left outer join 
    (
      select
        CtSCRB.CourtSessionID
        ,CtSCR.CalendarResourceBucketID
        ,CtSCRB.CalendarResourceTypeID as CalendarResourceTypeID
        ,uC_CRT.Code as CalendarResourceTypeCode
        ,uC_CRT.Description as CalendarResourceTypeDescription
        ,CtSCR.CalendarResourceID as CalendarResourceID
        ,uC_CR.Code as CalendarResourceCode
        ,uC_CR.Description as CalendarResourceDescription
      from
        Justice.dbo.CtSessionClndrResBucket CtSCRB
        inner join Justice.dbo.uCode uC_CRT
        on
          uC_CRT.CodeID = CtSCRB.CalendarResourceTypeID
          and uC_CRT.Code = 'CTRM'
        inner join Justice.dbo.xCtSessionClndrRes CtSCR
        on
          CtSCR.CalendarResourceBucketID = CtSCRB.CalendarResourceBucketID
        inner join Justice.dbo.uCode uC_CR
        on
          uC_CR.CodeID = CtSCR.CalendarResourceID
    ) CTRM
    on
      CTRM.CourtSessionID = CtS.CourtSessionID
  where
    1=1
    and cast(CtS.SessionDate as date) between @vBeginDate and @vEndDate
    --and cast(CtS.SessionDate as date) = '2025-08-28'
  ;
  -- select * from #CtSessionCourtHouse

/*
  select
    *
  from
    #CtSessionCourtHouse
  where
      1=1
      and Deleted = 0
      and 
      (
        LocationCode = 'FCCH'
        or
        LocationCode is null and CourtRoomCode in (select CourtRoomCode from #courtroom)
      )
  order by
    SessionDate
    ,StartTime
    ,SessionDescription
    ,CalendarDescription
  ;
*/

  select distinct
    SessionDate
    ,StartTime
    ,SessionDescription
    ,CalendarDescription
    ,JudicialOfficerCode
    ,JudicialOfficerDescription
    ,CourtRoomCode
    ,CourtRoomDescription
  from
    #CtSessionCourtHouse
  where
      1=1
      and Deleted = 0
      and 
      (
        LocationCode = 'FCCH'
        or
        LocationCode is null and CourtRoomCode in (select CourtRoomCode from #courtroom)
      )
  order by
    SessionDate
    ,StartTime
    ,SessionDescription
    ,CalendarDescription
  ;
end -- proc
;

/*
select distinct
  JudicialOfficerCode
  ,JudicialOfficerDescription
from
  #CtSessionCourtHouse
where
    1=1
    and Deleted = 0
    and 
    (
      LocationCode = 'FCCH'
      or
      LocationCode is null and CourtRoomCode in (select CourtRoomCode from #courtroom)
    )
order by
  JudicialOfficerDescription
;
*/

/*
exec Justice.fc.sp_getCourtSessionsByYear
  @pMonthOrYear=9
;
exec Justice.fc.sp_getCourtSessionsByYear
  @pMonthOrYear=2025
;
exec Justice.fc.sp_getCourtSessionsByYear
  @pMonthOrYear=0
;
*/
