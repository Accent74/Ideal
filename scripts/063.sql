USE [payroll]
GO

/****** Object:  StoredProcedure [dbo].[ideal_1DF]    Script Date: 01/16/2012 10:59:09 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ideal_1DF]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ideal_1DF]
GO

/****** Object:  StoredProcedure [dbo].[ideal_get_ag_list]    Script Date: 01/16/2012 10:59:09 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ideal_get_ag_list]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ideal_get_ag_list]
GO

/****** Object:  StoredProcedure [dbo].[ideal_loadpay1]    Script Date: 01/16/2012 10:59:09 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ideal_loadpay1]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ideal_loadpay1]
GO

/****** Object:  StoredProcedure [dbo].[ideal_loadpay2]    Script Date: 01/16/2012 10:59:09 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ideal_loadpay2]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[ideal_loadpay2]
GO

USE [payroll]
GO

/****** Object:  StoredProcedure [dbo].[ideal_1DF]    Script Date: 01/16/2012 10:59:09 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO



CREATE PROCEDURE [dbo].[ideal_1DF]

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

/****** Object:  StoredProcedure [dbo].[ideal_get_ag_list]    Script Date: 01/16/2012 10:59:09 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [dbo].[ideal_get_ag_list] 
	-- Add the parameters for the stored procedure here
		@pid int,
		@dno nvarchar(255),
		@advanceID int,
		@payID int ,
		@debtID int
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;
declare @Acc661 nvarchar(3)
set @Acc661 = '661'

	create table #Rep_tabel (ag_id int, ag_name nvarchar(255))

	insert into #Rep_tabel  (ag_id, ag_name)
	
	SELECT DISTINCT P_JOURNAL.AG_ID, AGENTS.AG_NAME
	FROM P_DOCUMENTS 
		INNER JOIN P_JOURNAL ON P_DOCUMENTS.PD_ID = P_JOURNAL.PD_ID
		INNER JOIN AGENTS 	ON AGENTS.AG_ID = P_JOURNAL.AG_ID
	
	WHERE (((P_JOURNAL.C_PERIOD)=@pid) AND ((P_JOURNAL.JP_SUM)<>0) AND ((P_DOCUMENTS.PD_NO)=@dno))
	ORDER BY AGENTS.AG_NAME

	SELECT P_JOURNAL.AG_ID, 
			#Rep_tabel.AG_NAME, 
			sum(CASE 
				WHEN (P_JOURNAL.JP_CR=@Acc661 And P_JOURNAL.MTC_ID <> @advanceID) THEN P_JOURNAL.JP_SUM
				ELSE 0
			END),

			sum(CASE 
				WHEN (P_JOURNAL.JP_DB=@Acc661 And P_JOURNAL.MTC_ID<>@payID) THEN P_JOURNAL.JP_SUM
				ELSE 0
			END),
					
			sum(CASE 
				WHEN (P_JOURNAL.JP_CR=@Acc661 And P_JOURNAL.MTC_ID<>@advanceID) THEN P_JOURNAL.JP_SUM
				ELSE 0
			END)
			-- -			
			--sum(CASE 
			--	WHEN (P_JOURNAL.JP_DB=@Acc661) THEN P_JOURNAL.JP_SUM
			--	ELSE 0
			--END)
					
	FROM P_JOURNAL 
		INNER JOIN #Rep_tabel ON P_JOURNAL.AG_ID = #Rep_tabel.ag_id
	WHERE (((P_JOURNAL.JP_SUM)<>0) AND ((P_JOURNAL.C_PERIOD)=@pid) AND ((P_JOURNAL.[MTC_ID])<>@debtID))
	GROUP BY P_JOURNAL.AG_ID, #Rep_tabel.AG_NAME
	ORDER BY #Rep_tabel.AG_NAME

	drop table #Rep_tabel
END

GO

/****** Object:  StoredProcedure [dbo].[ideal_loadpay1]    Script Date: 01/16/2012 10:59:09 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE procedure [dbo].[ideal_loadpay1]
	@PeriodID int, 
	@PayTypePropID int, 
	@MtdID int, 
	@mtdTaxID1 int, 
	@mtdTaxID2 int, 
	@AgPropValue int
as
begin
	SET NOCOUNT ON;

	create table #ideal_sum_by_ag_mtc_period (AG_VATNO nvarchar(50),
							MTC_ID int, SumJP_SUM money)
	insert into #ideal_sum_by_ag_mtc_period (AG_VATNO,MTC_ID, SumJP_SUM)
	SELECT AGENTS.AG_VATNO, P_JOURNAL.MTC_ID, Sum(P_JOURNAL.JP_SUM)
	FROM AGENTS INNER JOIN P_JOURNAL ON AGENTS.AG_ID = P_JOURNAL.AG_ID
	WHERE (((P_JOURNAL.JP_DONE)=1) AND ((P_JOURNAL.C_PERIOD)=@PeriodID))
	GROUP BY AGENTS.AG_VATNO, P_JOURNAL.MTC_ID

	create table #ideal_pay_by_deps(AG_VATNO nvarchar(50), 
							PROP_LONG int, 
							W_PERIOD int, 
							PP_MONTH int,
							PP_YEAR int,
							MtdSum money)
	insert into #ideal_pay_by_deps (AG_VATNO, PROP_LONG, W_PERIOD, PP_MONTH, PP_YEAR, MtdSum) 							

	SELECT	AGENTS.AG_VATNO, 
			AG_PROPS.PROP_LONG, 
			P_JOURNAL.W_PERIOD, 
			P_PERIOD.PP_MONTH, 
			P_PERIOD.PP_YEAR, 
			Sum(P_JOURNAL.JP_SUM) AS MtdSum
	FROM P_PERIOD RIGHT JOIN ((AGENTS LEFT JOIN AG_PROPS ON AGENTS.AG_ID = AG_PROPS.EL_ID) RIGHT JOIN (MTD_DEPENDS LEFT JOIN P_JOURNAL ON MTD_DEPENDS.MTD_ID_LEFT = P_JOURNAL.MTC_ID) ON AGENTS.AG_ID = P_JOURNAL.AG_ID) ON P_PERIOD.PP_ID = P_JOURNAL.W_PERIOD
	GROUP BY AGENTS.AG_VATNO, AG_PROPS.PROP_LONG, P_JOURNAL.W_PERIOD, P_PERIOD.PP_MONTH, P_PERIOD.PP_YEAR, P_JOURNAL.C_PERIOD, P_JOURNAL.JP_DONE, MTD_DEPENDS.MTD_ID_RIGHT, AG_PROPS.PROP_ID
	HAVING (((Sum(P_JOURNAL.JP_SUM))<>0) AND ((P_JOURNAL.C_PERIOD)=@PeriodID) AND ((P_JOURNAL.JP_DONE)=1) AND ((MTD_DEPENDS.MTD_ID_RIGHT)=@MtdID) AND ((AG_PROPS.PROP_ID)=@PayTypePropID));

	Select	#ideal_pay_by_deps.AG_VATNO, 
			#ideal_pay_by_deps.W_Period, 
			#ideal_pay_by_deps.PP_MONTH, 
			#ideal_pay_by_deps.PP_YEAR, 
			#ideal_pay_by_deps.MtdSum, 
			Sum(#ideal_sum_by_ag_mtc_period.SumJP_SUM) 
	FROM #ideal_pay_by_deps Left Join #ideal_sum_by_ag_mtc_period On #ideal_pay_by_deps.AG_VATNO = #ideal_sum_by_ag_mtc_period.AG_VATNO
	WHERE ((((#ideal_sum_by_ag_mtc_period.MTC_ID)=@mtdTaxID1)) Or (((#ideal_sum_by_ag_mtc_period.MTC_ID)=@mtdTaxID2))) And #ideal_pay_by_deps.PROP_LONG = @AgPropValue
	GROUP BY #ideal_pay_by_deps.AG_VATNO, #ideal_pay_by_deps.PROP_LONG, #ideal_pay_by_deps.W_PERIOD, #ideal_pay_by_deps.PP_MONTH, #ideal_pay_by_deps.PP_YEAR, #ideal_pay_by_deps.MtdSum
		
end
GO

/****** Object:  StoredProcedure [dbo].[ideal_loadpay2]    Script Date: 01/16/2012 10:59:09 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER OFF
GO

CREATE procedure [dbo].[ideal_loadpay2]
	@PeriodID int, 
	@PayTypePropID int, 
	@MtdID int, 
	@mtdTaxID1 int, 
	@mtdTaxID2 int, 
	@AgPropValue int
as
begin

	SET NOCOUNT ON;
	create table #ideal_sum_by_ag_mtc_period (AG_VATNO nvarchar(50),
							MTC_ID int, SumJP_SUM money)
	insert into #ideal_sum_by_ag_mtc_period (AG_VATNO,MTC_ID, SumJP_SUM)
	SELECT AGENTS.AG_VATNO, P_JOURNAL.MTC_ID, Sum(P_JOURNAL.JP_SUM)
	FROM AGENTS INNER JOIN P_JOURNAL ON AGENTS.AG_ID = P_JOURNAL.AG_ID
	WHERE (((P_JOURNAL.JP_DONE)=1) AND ((P_JOURNAL.C_PERIOD)=@PeriodID))
	GROUP BY AGENTS.AG_VATNO, P_JOURNAL.MTC_ID

	create table #ideal_pay_by_deps(AG_VATNO nvarchar(50), 
							PROP_LONG int, 
							W_PERIOD int, 
							PP_MONTH int,
							PP_YEAR int,
							MtdSum money)
	insert into #ideal_pay_by_deps (AG_VATNO, PROP_LONG, W_PERIOD, PP_MONTH, PP_YEAR, MtdSum) 							

	SELECT	AGENTS.AG_VATNO, 
			AG_PROPS.PROP_LONG, 
			P_JOURNAL.W_PERIOD, 
			P_PERIOD.PP_MONTH, 
			P_PERIOD.PP_YEAR, 
			Sum(P_JOURNAL.JP_SUM) AS MtdSum
	FROM P_PERIOD RIGHT JOIN ((AGENTS LEFT JOIN AG_PROPS ON AGENTS.AG_ID = AG_PROPS.EL_ID) RIGHT JOIN (MTD_DEPENDS LEFT JOIN P_JOURNAL ON MTD_DEPENDS.MTD_ID_LEFT = P_JOURNAL.MTC_ID) ON AGENTS.AG_ID = P_JOURNAL.AG_ID) ON P_PERIOD.PP_ID = P_JOURNAL.W_PERIOD
	GROUP BY AGENTS.AG_VATNO, AG_PROPS.PROP_LONG, P_JOURNAL.W_PERIOD, P_PERIOD.PP_MONTH, P_PERIOD.PP_YEAR, P_JOURNAL.C_PERIOD, P_JOURNAL.JP_DONE, MTD_DEPENDS.MTD_ID_RIGHT, AG_PROPS.PROP_ID
	HAVING (((Sum(P_JOURNAL.JP_SUM))<>0) AND ((P_JOURNAL.C_PERIOD)=@PeriodID) AND ((P_JOURNAL.JP_DONE)=1) AND ((MTD_DEPENDS.MTD_ID_RIGHT)=@MtdID) AND ((AG_PROPS.PROP_ID)=@PayTypePropID));

	Select	#ideal_pay_by_deps.AG_VATNO, 
			#ideal_pay_by_deps.W_Period, 
			#ideal_pay_by_deps.PP_MONTH, 
			#ideal_pay_by_deps.PP_YEAR, 
			#ideal_pay_by_deps.MtdSum, 
			Sum(#ideal_sum_by_ag_mtc_period.SumJP_SUM) 
	FROM #ideal_pay_by_deps Left Join #ideal_sum_by_ag_mtc_period On #ideal_pay_by_deps.AG_VATNO = #ideal_sum_by_ag_mtc_period.AG_VATNO
	WHERE ((((#ideal_sum_by_ag_mtc_period.MTC_ID)=@mtdTaxID1)) Or (((#ideal_sum_by_ag_mtc_period.MTC_ID)=@mtdTaxID2))) and #ideal_pay_by_deps.PROP_LONG <> @AgPropValue
	GROUP BY #ideal_pay_by_deps.AG_VATNO, #ideal_pay_by_deps.PROP_LONG, #ideal_pay_by_deps.W_PERIOD, #ideal_pay_by_deps.PP_MONTH, #ideal_pay_by_deps.PP_YEAR, #ideal_pay_by_deps.MtdSum
		
end
GO


