if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ideal_1DF]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[ideal_1DF]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE PROCEDURE ideal_1DF

@nYear int, @mStart int, @mEnd int, @MTD_IN_ID int, @MTD_OUT_ID int, @MTD_TAX_ID int, @lgtCode int

AS

set nocount on

Create table #AgData (	AgID int, 
			AgName  nvarchar(255), 
			lPreiod int, 
			Sum3a Money, 
			Sum3 Money, 
			Sum4a money, 
			Sum4 Money, 
			InType money, 
			lgType money)

insert into #AgData(AgID, AgName, lPreiod, Sum3a, Sum3, Sum4a, Sum4, InType, lgType)

SELECT 	P_JOURNAL.AG_ID, 
	AGENTS.AG_NAME, 
	pp_year*100+pp_month, 
	Sum(P_JOURNAL.JP_SUM), 
	0, 0, 0, 1, 0
FROM P_PERIOD 	RIGHT JOIN (AGENTS RIGHT JOIN (P_JOURNAL RIGHT JOIN MTD_DEPENDS ON P_JOURNAL.MTC_ID = MTD_DEPENDS.MTD_ID_LEFT) ON AGENTS.AG_ID = P_JOURNAL.AG_ID) ON P_PERIOD.PP_ID = P_JOURNAL.W_PERIOD
WHERE (((MTD_DEPENDS.MTD_ID_RIGHT)=@MTD_IN_ID) AND ((P_JOURNAL.JP_DONE)=1) AND ((P_PERIOD.PP_YEAR)=@nYear) AND ((P_PERIOD.PP_MONTH)>=@mStart And (P_PERIOD.PP_MONTH)<=@mEnd))
GROUP BY P_JOURNAL.AG_ID, AGENTS.AG_NAME, pp_year*100+pp_month
HAVING (((Sum(P_JOURNAL.JP_SUM))<>0))
ORDER BY AGENTS.AG_NAME


insert into #AgData(AgID, AgName, lPreiod, Sum3a, Sum3, Sum4a, Sum4, InType, lgType)
SELECT 	P_JOURNAL.AG_ID, 
	AGENTS.AG_NAME, 
	pp_year*100+pp_month,
	0, 
	Sum(P_JOURNAL.JP_SUM),
	0, 0, 1, 0
FROM P_PERIOD RIGHT JOIN (AGENTS RIGHT JOIN (P_JOURNAL RIGHT JOIN MTD_DEPENDS ON P_JOURNAL.MTC_ID = MTD_DEPENDS.MTD_ID_LEFT) ON AGENTS.AG_ID = P_JOURNAL.AG_ID) ON P_PERIOD.PP_ID = P_JOURNAL.W_PERIOD
WHERE (((MTD_DEPENDS.MTD_ID_RIGHT)=@MTD_OUT_ID) AND ((P_JOURNAL.JP_DONE)=1) AND ((P_PERIOD.PP_YEAR)=@nYear) AND ((P_PERIOD.PP_MONTH)>=@mStart And (P_PERIOD.PP_MONTH)<=@mEnd))
GROUP BY P_JOURNAL.AG_ID, AGENTS.AG_NAME, pp_year*100+pp_month
HAVING (((Sum(P_JOURNAL.JP_SUM))<>0))
ORDER BY AGENTS.AG_NAME


insert into #AgData(AgID, AgName, lPreiod, Sum3a, Sum3, Sum4a, Sum4, InType, lgType)

SELECT 	P_JOURNAL.AG_ID, 
	AGENTS.AG_NAME, 
	pp_year*100+pp_month, 
	0, 0, 
	Sum(P_JOURNAL.JP_SUM), 
	Sum(P_JOURNAL.JP_SUM), 
	Case 
		When isnull(JP_T1, 0) = 0 then 1
		else JP_T1
	end, 0
FROM P_PERIOD RIGHT JOIN (AGENTS RIGHT JOIN (P_JOURNAL RIGHT JOIN MTD_DEPENDS ON P_JOURNAL.MTC_ID = MTD_DEPENDS.MTD_ID_LEFT) ON AGENTS.AG_ID = P_JOURNAL.AG_ID) ON P_PERIOD.PP_ID = P_JOURNAL.W_PERIOD
WHERE (((MTD_DEPENDS.MTD_ID_RIGHT)=@MTD_TAX_ID) AND ((P_JOURNAL.JP_DONE)=1) AND ((P_PERIOD.PP_YEAR)=@nYear) AND ((P_PERIOD.PP_MONTH)>=@mStart And (P_PERIOD.PP_MONTH)<=@mEnd))
GROUP BY P_JOURNAL.AG_ID, AGENTS.AG_NAME, pp_year*100+pp_month, JP_T1
HAVING (((Sum(P_JOURNAL.JP_SUM))<>0))
ORDER BY AGENTS.AG_NAME

insert into #AgData(AgID, AgName, lPreiod, Sum3a, Sum3, Sum4a, Sum4, InType, lgType)

SELECT 	P_JOURNAL.AG_ID,
	AGENTS.AG_NAME,
	PP_YEAR*100+PP_MONTH,
	0, 0, 0, 0, 1,
	Case
		When isnull(P_JOURNAL.JP_T3,0) = 0 then 0
		else P_JOURNAL.JP_T3
	end

FROM P_PERIOD LEFT JOIN (AGENTS RIGHT JOIN P_JOURNAL ON AGENTS.AG_ID = P_JOURNAL.AG_ID) ON P_PERIOD.PP_ID = P_JOURNAL.W_PERIOD
WHERE (((P_JOURNAL.JP_DONE)=1) AND ((P_JOURNAL.MTC_ID)=@lgtCode) AND ((P_PERIOD.PP_MONTH)>=@mStart And (P_PERIOD.PP_MONTH)<=@mEnd) AND ((P_PERIOD.PP_YEAR)=@nYear) AND ((P_JOURNAL.JP_SUM)<>0))
GROUP BY P_JOURNAL.AG_ID, AGENTS.AG_NAME, PP_YEAR*100+PP_MONTH, P_JOURNAL.JP_T3
ORDER BY AGENTS.AG_NAME

SELECT 	AGID as AgID, 
	AGNAME as AgName, 
	Sum(Sum3a) as Sim3a, 
	Sum(Sum3) as Sum3, 
	Sum(Sum4a) as Sum4a, 
	Sum(Sum4) as Sum4, 
	InType, Max(dbo.Num2Pos(lgtype, 1)) AS lgt_1, Max(dbo.Num2Pos(lgtype, 2)) AS lgt_2, Max(dbo.Num2Pos(lgtype, 3)) AS lgt_3, Max(dbo.Num2Pos(lgtype, 4)) AS lgt_4
FROM #AgData
GROUP BY AGID, AGNAME, InType
ORDER BY AGNAME, AGID
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO