ALTER procedure [dbo].[ideal_loadpay1]
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
							MTC_ID int, SumJP_SUM money, long3 int)
	insert into #ideal_sum_by_ag_mtc_period (AG_VATNO,MTC_ID, SumJP_SUM, long3)
	SELECT AGENTS.AG_VATNO, P_JOURNAL.MTC_ID, Sum(P_JOURNAL.JP_SUM), p_journal.jp_pl3
	FROM AGENTS INNER JOIN P_JOURNAL ON AGENTS.AG_ID = P_JOURNAL.AG_ID
	WHERE (((P_JOURNAL.JP_DONE)=1) AND ((P_JOURNAL.C_PERIOD)=@PeriodID))
	GROUP BY AGENTS.AG_VATNO, P_JOURNAL.MTC_ID, p_journal.jp_pl3

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
	HAVING (((Sum(P_JOURNAL.JP_SUM))<>0) AND ((P_JOURNAL.C_PERIOD)=@PeriodID) AND ((P_JOURNAL.JP_DONE)=1) AND ((MTD_DEPENDS.MTD_ID_RIGHT)=@MtdID))

	Select	#ideal_pay_by_deps.AG_VATNO, 
			#ideal_pay_by_deps.W_Period, 
			#ideal_pay_by_deps.PP_MONTH, 
			#ideal_pay_by_deps.PP_YEAR, 
			#ideal_pay_by_deps.MtdSum, 
			Sum(#ideal_sum_by_ag_mtc_period.SumJP_SUM),
			long3  
	FROM #ideal_pay_by_deps Left Join #ideal_sum_by_ag_mtc_period On #ideal_pay_by_deps.AG_VATNO = #ideal_sum_by_ag_mtc_period.AG_VATNO
	WHERE ((((#ideal_sum_by_ag_mtc_period.MTC_ID)=@mtdTaxID1)) Or (((#ideal_sum_by_ag_mtc_period.MTC_ID)=@mtdTaxID2))) 
	GROUP BY #ideal_pay_by_deps.AG_VATNO, #ideal_pay_by_deps.PROP_LONG, #ideal_pay_by_deps.W_PERIOD, #ideal_pay_by_deps.PP_MONTH, #ideal_pay_by_deps.PP_YEAR, #ideal_pay_by_deps.MtdSum, long3
		
end


GO


