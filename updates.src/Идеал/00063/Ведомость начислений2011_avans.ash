аЯрЁБс                >  ўџ	                               ўџџџ       џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ0
	sum2 = 0
	arFonds = Array( "2000р", "2001р" )
	Set Repcalc = WorkArea.CreateIntObject("CommonCalc")
	Repcalc.setPeriod "month", Workarea.Period.Month, Workarea.Period.Year
	d = Workarea.Period.FirstDay

	max = Workarea.UTable( USRTBL_LGT ).GetValue( 1, USRTBL_MAXCY, 2 )
	percent = Workarea.UTable( USRTBL_LGT ).GetValue( 1, USRTBL_ECB, 2 )
	With Sheet( 1 )
		CodeID = WorkArea.GetCodeID( "2000р" )
		With WorkArea.Code( CodeID )
			Sheet( 1 ).Cell( 5, 7 ).Value = .DbCode
			Sheet( 1 ).Cell( 5, 8 ).Value = .CrCode
			percent = .Percent
			Sheet( 1 ).Cell( 5, 2 ).Value = percent
		End With
		FOT = CalcDepsSumLimit_Pens( CodeID, Repcalc, max, False )
		FOT1 = FOT
		sum1 = sum1 + Round2( FOT * percent / 100 , 2 )

		Repcalc.Clear
		CodeID = WorkArea.GetCodeID( "2001р" )
		With WorkArea.Code( CodeID )
			percent = .Percent
			Sheet( 1 ).Cell( 6, 7 ).Value = .DbCode
			Sheet( 1 ).Cell( 6, 8 ).Value = .CrCode
			Sheet( 1 ).Cell( 6, 2 ).Value = percent
		End With
		FOT = CalcDepsSumLimit_Pens( CodeID, Repcalc, max, True )
		FOT2 = FOT
		sum2 = sum2 + Round2( FOT * percent / 100 , 2 )

		.Cell( 5, 3 ).value = FOT1
		.Cell( 5, 5 ).value = sum1
		.Cell( 6, 3 ).value = FOT2
		.Cell( 6, 5 ).value = sum2

	End With
End Sub
'________________________________________________
Sub BuildLastReport
	Dim percent, i, sum1, sumT, CodeID
	Dim arFonds

	sum1 = 0

	arFonds = Array( MTD_NDFL_ID, MTD_ESV_ID, MTD_ESVB_ID, MTD_ESVT_ID)
	percent  = Array( 15, 3.6, 2, 2.6)

	With Sheet( 1 )
		For i = 0 To UBound(arFonds)
			GetSumByCode arFonds(i), sum1, sumT
			cRow = LastRow + i
			.Cell( cRow, 2 ).Value = percent(i)
			.Cell( cRow, 3 ).value = sum1
			With WorkArea.Code( arFonds(i) )
				Sheet( 1 ).Cell( cRow, 4 ).Value = .DbCode
				Sheet( 1 ).Cell( cRow, 5 ).Value = .CrCode
			End With
		Next
	End With
End Sub
'=================
Sub GetSumByCode(mtdID, sum1, sumT)
	Dim rs
	InitRs rs, _
		"SELECT SUM(JP_SUM) AS PSUM, sum(jp_tariff) as pt FROM P_JOURNAL WHERE JP_DONE = 1 " & _
			" AND C_PERIOD = " & WorkArea.Period.ID & _
			" AND MTC_ID = " & mtdID
	If Not rs.eof	Then
		sum1 = rs.fields("PSUM")
		sumT = rs.fields("pt")
	End If
End Sub
'________________________________________________
Function CalcDepsSumLimit( CodeID, Repcalc, d, max )
	Dim i
	Dim TotalSum
	Dim AgSum
	Dim flAllowDogovor

	With Workarea.Code(CodeID)
		flAllowDogovor = Iif( ((.Mode And 8) = 0) Or (.Type <> 103), False, True )
	End With

	TotalSum = 0
	For i = 1 To Repcalc.Count
		If flAllowDogovor Or (Not Repcalc.Agent(i).Prop("в№ѓфютющ фюуютю№", d)) Then
			AgSum = Repcalc.SumBy("CodeSumDependsW", i, CodeID)
			If AgSum > max Then AgSum = max
			TotalSum = TotalSum + AgSum
		End If
	Next
	CalcDepsSumLimit = TotalSum
End Function

'________________________________________________
Function CalcDepsSumLimit_Pens( CodeID, Repcalc, max, IsAgInvalid )
	Dim i
	Dim TotalSum
	Dim AgSum

	TotalSum = 0

	For i = 1 To Repcalc.Count
		If Repcalc.Agent(i).Prop("Шэтрышф") = IsAgInvalid Then
			AgSum = Repcalc.SumBy("CodeSumDependsW", i, CodeID)
			If AgSum > max Then AgSum = max
			TotalSum = TotalSum + AgSum
		End If
	Next

	CalcDepsSumLimit_Pens = TotalSum
End Function
<     ш                Ь Arial D:\STDs.!!!\Zap\Шфхры\Шфхры       џџџџ                 џџ 
 CSheetPageSheet1Тхфюьюёђќ эрїшёыхэшщ         61    7                                     X   3   O   B   F   *   0   2   1   *   1   1   *   &   &   1   F                                                     џџ  CRow џџ  CCell6     џџ   (   54><>ABL  =0G8A;5=89                      " џџџџ                        " џџџџ                        " џџџџ                         " џџџџ                          " џџџџ                          " џџџџ                              џџ      /=20@L  2 0 1 2   3.                       " џџџџ                        " џџџџ                         ! џџ+       0G8A;5=85                    ! џџ+       %                     ! џџ+     $"                    " џџ+                        ! џџ+  
   !C<<0                    ! џџ+                        ! џџ+    
   515B                    ! џџ+       @548B                               B>3>:                                                 "                         "                          "                         "                          "                           "                              џџ    "   !  >1I89  ( 020=A)                       џџ)    xЏ                         " џџ                        " џџ                        " џџ                              " џџ                        " џџ                            " џџ       6 5 7                        џџ    (   !  8=20;84K  ( 020=A)                       џџ    H                         " џџ                          џџ                        " џџ                                џџ                        " џџ                            " џџ       6 5 7                          " џџџџ                        " џџџџ                        " џџџџ                        " џџџџ                         ! џџ+       #45@60=85                    ! џџ+       %                     ! џџ+    
   !C<<0                    ! џџ+    
   515B                    ! џџ+       @548B                       џџ    
   $                        џџ                         " џџ                         " џџ                         " џџ                            џџ       !  ( 107>2K9)                       џџ                         " џџ                         " џџ                         " џџ                            џџ        !  ( 1>;L=8G=K5)                       џџ                         " џџ                         " џџ                         " џџ                            џџ       !  ( )                       џџ                         " џџ                         " џџ                         " џџ                              " џџџџ                        " џџџџ                         % џџ,       >4@0745;5=85                     џџ,  
   :;04                    % џџ+  
   20=A                    % џџ+     $                    % џџ+       2K40G5                     % џџ-       !>B@C4=8:                     џџ-                       " џџ	                       % џџ+                       % џџ+                        "                           "                                  " џџ                                " џџ                                " џџ                                     џџџџ :   CE30;B5@: _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _                       џџџџ                         џџџџ    /   1 1 1   /                       џџџџ                         џџџџ                         џџџџ                                      <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                                2   2      	 <     ш    Д           Ь Arial ющёђт ћ (F7) ёюёђюџэшџ фыџ ю <     ш   Ш           Ь Arial                      шЭe <     ш            ((( Ь Arial џџ                   а#g <     ш           џџџ Ь Arial џџ                   ing  <     ш           ((( Ь Arial џџ               htRange) <     ш           џџр Ь Arial џџ               №ѓђъш    <     ш               Ь Arial шђх F1  (F7)    1      шЭe  <     ш    №           Ь Arial шђх F1 ыхээюую §ыхьхэђр 
.av <     ш   №           Ь Arial Sub ShtBook_OnLoadEnd Su <     ш    Ш           Ь Arial аDb                   
   
 <     ш            РРР Ь Arial ђюф F1 ьхэђ ёяшёър 2 _ 43.ad/    рџџ                              рџџ                               рџџ                                џџџ                               
 рџџ                              	 џџџџ                              џџџџ                               џџџџ                              џџр                              џџџџ                             џџр                                  ммм                              ммм                                   ШШШ                                  ШШШ                              џџџ                               РРР                               џџџ                              џџџ                                  РРР                              ШШШ                               РРР                               ШШШ                               РРР                                  џџџџ                               џџџџ                                џџџџ                               џџџ                               џџџ                                рџџ                                 џџџ                                 џџџ                               * џџџџ                               ) џџџџ                             ( рџџ                                 % ррр                              $ ррр                                џџр                              ! џџр                              " ррр                                 # рррR o o t   E n t r y                                               џџџџџџџџ                               p}яљCаЬ"         C o n t e n t s                                                  џџџџ   џџџџ                                    #   9       S u m m a r y I n f o r m a t i o n                           (  џџџџџџџџџџџџ                                        q                                                                          џџџџџџџџџџџџ                                                                        	   
                                 ўџџџ§џџџўџџџўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџўџџџ$   %   &   '   (   )   *       џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ               ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ                             & ррр                               ' џџр                              + ццњ                             , ццњ                              - ццњ                              . џџџџ                                 m a r y I n f o r m a t i o n                           (  џџџџџџџџџџџџ                                        q                                                                          џџџџџџџџџџџџ                                                                        	   
                              ўџџџџџџџџџџџџџџџўџџџ§џџџўџџџўџџџ                !   +   џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ    џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ               ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџўџ   рђљOhЋ +'Гй   рђљOhЋ +'Гй0   A        x            Ќ      И      Ф      а      м   	   ш   
   є                     ,      4     у     #   Тхфюьюёђќ эрїшёыхэшщ2011_avans.ash                                                          499 @   pGz  @   @ѕCаЬ@   @kчєЖ§Ц      Ръіхэђ 7.0                                                                                                                                                                   џОOption Explicit
'#include "definition.avb"

Const StartNumberOfRows = 16
Const LastRow = 9
Dim cRow

Const MTD_SALARY_NAME 	= "0117"
Const MTD_NDFL_NAME 		= "2400р"
Const MTD_ESV_NAME 			= "2410р"
Const MTD_ESVB_NAME			= "2420р"
Const MTD_ESVT_NAME 		= "2430р"
Const MTD_AVANS_NAME 		= "3020"

Dim MTD_SALARY_ID  
Dim MTD_NDFL_ID 	
Dim MTD_ESV_ID 	
Dim MTD_ESVB_ID	
Dim MTD_ESVT_ID 	
Dim MTD_AVANS_ID		
Dim BuhName
'-----------------------------------------------------------------------------------------------
Sub ShtBook_OnLoad
	Dim ut
	Set ut = Workarea.UTable(USRTBL_ONE)
	BuhName =ut.GetValue(1, "дШЮ сѓѕурыђх№р", 2)

	With Sheet( 1 )
		.DisplayGrid = False
		.DisplayHeadings = True
	End With
	
	EnableButtons 1
	PeriodText = WorkArea.Period.RepTitle
	Sheet( 1 ).Cell( 2, 1 ).Value = WorkArea.Period.RepTitle
	Sheet( 1 ).Range( 5, 1, 8, 1 ).Mergecells = True
	Sheet( 1 ).Range( 5, 1, 8, 1 ).Value = WorkArea.Agent(1237).Name
	Sheet( 1 ).Range( 5, 1, 8, 1 ).Alignment = acRight

	MTD_SALARY_ID 		= Workarea.GetCodeID(MTD_SALARY_NAME)
	MTD_NDFL_ID 		= Workarea.GetCodeID(MTD_NDFL_NAME)
	MTD_ESV_ID 			= Workarea.GetCodeID(MTD_ESV_NAME)
	MTD_ESVB_ID 		= Workarea.GetCodeID(MTD_ESVB_NAME)
	MTD_ESVT_ID 		= Workarea.GetCodeID(MTD_ESVT_NAME)
	MTD_AVANS_ID 		= Workarea.GetCodeID(MTD_AVANS_NAME)

	BuildReport
End Sub

'________________________________________________
Sub ShtBook_OnRequery
	BuildReport
End Sub

'-----------------------------------------------------------------------------------------------
Sub ShtBook_OnPeriodClick(Buttons)
	If WorkArea.Period.Browse Then
		PeriodText = WorkArea.Period.RepTitle
		Sheet( 1 ).Cell( 2, 1 ).Value = WorkArea.Period.RepTitle
	End If
End Sub

'-----------------------------------------------------------------------------------------------
Sub BuildReport
	BuildMainReport
	BuildSecondaryReport
	BuildLastReport
End Sub
'-----------------------------------------------------------------------------------------------
Sub BuildMainReport
	Dim i, cols
	Dim repCC
	Set repCC = WorkArea.CreateIntObject("CommonCalc")
	repcc.setPeriod "month", Workarea.Period.Month, Workarea.Period.Year

	cols = 5
	cRow = StartNumberOfRows
	With Sheet( 1 )
		.Rows = StartNumberOfRows
		' юїшёђшыш ёђ№юъѓ шђюуют
		For i = 2 To Cols
			.Cell( StartNumberOfRows, i ).Value = 0.0
		Next
	End With

	With Sheet( 1 )
		.Cell( StartNumberOfRows, 3).Value = repCC.TotalSumBy ("CodeSumW", MTD_SALARY_ID)
		.Cell( StartNumberOfRows, 4).Value = repCC.TotalSumBy ("CodeSumW", MTD_NDFL_ID)
		.Cell( StartNumberOfRows, 5).Value = repCC.TotalSumBy ("CodeSumW", MTD_AVANS_ID)
		.Range( 1, StartNumberOfRows, .Columns, StartNumberOfRows ).Font.Bold = True

		For i = 1 To repcc.CountAgent
			cRow = .InsertRow2
			.Cell(cRow, 1 ).CellParam = 1
			.Cell(cRow, 1).Value = repCC.Agent(i).Name
			.Cell(cRow, 2).Value = repCC.Agent(i).Prop("Юъырф")
			.Cell(cRow, 3).Value = repCC.SumBy ("CodeSumW", i, MTD_SALARY_ID)
			.Cell(cRow, 4).Value = repCC.SumBy ("CodeSumW", i, MTD_NDFL_ID)
			.Cell(cRow, 5).Value = repCC.SumBy ("CodeSumW", i, MTD_AVANS_ID)
		Next

		.Range( 1, StartNumberOfRows - 2, Cols, .Rows).Autofit
		.Range( 1, StartNumberOfRows + 1, Cols, .Rows).Stripe
		.Range( 1, StartNumberOfRows + 1, Cols, .Rows).Alignment = acRight
		.Range( 1, StartNumberOfRows + 1, Cols, .Rows).SetBorder acBrdOutline + acBrdInsideVert, acLineNormal
		With .Range( 2, StartNumberOfRows, Cols, StartNumberOfRows )
			.Font.Bold = True
			.Alignment = acRight
			.Backўџ   рђљOhЋ +'Гй   рђљOhЋ +'Гй0   A        x            Ќ      И      Ф      а      м   	   ш   
   є                     ,      4     у     #   Тхфюьюёђќ эрїшёыхэшщ2011_avans.ash                                                          500 @   PёeKz  @    іфљCаЬ@   @kчєЖ§Ц      Ръіхэђ 7.0                                                                                                                                                                   џРOption Explicit
'#include "definition.avb"

Const StartNumberOfRows = 16
Const LastRow = 9
Dim cRow

Const MTD_SALARY_NAME 	= "0117"
Const MTD_NDFL_NAME 		= "2400р"
Const MTD_ESV_NAME 			= "2410р"
Const MTD_ESVB_NAME			= "2420р"
Const MTD_ESVT_NAME 		= "2430р"
Const MTD_AVANS_NAME 		= "3020"

Dim MTD_SALARY_ID  
Dim MTD_NDFL_ID 	
Dim MTD_ESV_ID 	
Dim MTD_ESVB_ID	
Dim MTD_ESVT_ID 	
Dim MTD_AVANS_ID		
Dim BuhName
'-----------------------------------------------------------------------------------------------
Sub ShtBook_OnLoad
	Dim ut
	Set ut = Workarea.UTable(USRTBL_ONE)
	BuhName =ut.GetValue(1, "дШЮ сѓѕурыђх№р", 2)

	With Sheet( 1 )
		.DisplayGrid = False
		.DisplayHeadings = True
	End With
	
	EnableButtons 1
	PeriodText = WorkArea.Period.RepTitle
	Sheet( 1 ).Cell( 2, 1 ).Value = WorkArea.Period.RepTitle
	Sheet( 1 ).Range( 5, 1, 8, 1 ).Mergecells = True
	Sheet( 1 ).Range( 5, 1, 8, 1 ).Value = WorkArea.Agent(1237).Name
	Sheet( 1 ).Range( 5, 1, 8, 1 ).Alignment = acRight

	MTD_SALARY_ID 		= Workarea.GetCodeID(MTD_SALARY_NAME)
	MTD_NDFL_ID 		= Workarea.GetCodeID(MTD_NDFL_NAME)
	MTD_ESV_ID 			= Workarea.GetCodeID(MTD_ESV_NAME)
	MTD_ESVB_ID 		= Workarea.GetCodeID(MTD_ESVB_NAME)
	MTD_ESVT_ID 		= Workarea.GetCodeID(MTD_ESVT_NAME)
	MTD_AVANS_ID 		= Workarea.GetCodeID(MTD_AVANS_NAME)

	BuildReport
End Sub

'________________________________________________
Sub ShtBook_OnRequery
	BuildReport
End Sub

'-----------------------------------------------------------------------------------------------
Sub ShtBook_OnPeriodClick(Buttons)
	If WorkArea.Period.Browse Then
		PeriodText = WorkArea.Period.RepTitle
		Sheet( 1 ).Cell( 2, 1 ).Value = WorkArea.Period.RepTitle
	End If
End Sub

'-----------------------------------------------------------------------------------------------
Sub BuildReport
	BuildMainReport
	BuildSecondaryReport
	BuildLastReport
End Sub
'-----------------------------------------------------------------------------------------------
Sub BuildMainReport
	Dim i, cols
	Dim repCC
	Set repCC = WorkArea.CreateIntObject("CommonCalc")
	repcc.setPeriod "month", Workarea.Period.Month, Workarea.Period.Year

	cols = 5
	cRow = StartNumberOfRows
	With Sheet( 1 )
		.Rows = StartNumberOfRows
		' юїшёђшыш ёђ№юъѓ шђюуют
		For i = 2 To Cols
			.Cell( StartNumberOfRows, i ).Value = 0.0
		Next
	End With

	With Sheet( 1 )
		.Cell( StartNumberOfRows, 3).Value = repCC.TotalSumBy ("CodeSumW", MTD_SALARY_ID)
		.Cell( StartNumberOfRows, 4).Value = repCC.TotalSumBy ("CodeSumW", MTD_NDFL_ID)
		.Cell( StartNumberOfRows, 5).Value = repCC.TotalSumBy ("CodeSumW", MTD_AVANS_ID)
		.Range( 1, StartNumberOfRows, .Columns, StartNumberOfRows ).Font.Bold = True

		For i = 1 To repcc.CountAgent
			cRow = .InsertRow2
			.Cell(cRow, 1 ).CellParam = 1
			.Cell(cRow, 1).Value = repCC.Agent(i).Name
			.Cell(cRow, 2).Value = repCC.Agent(i).Prop("Юъырф")
			.Cell(cRow, 3).Value = repCC.SumBy ("CodeSumW", i, MTD_SALARY_ID)
			.Cell(cRow, 4).Value = repCC.SumBy ("CodeSumW", i, MTD_NDFL_ID)
			.Cell(cRow, 5).Value = repCC.SumBy ("CodeSumW", i, MTD_AVANS_ID)
		Next

		.Range( 1, StartNumberOfRows - 2, Cols, .Rows).Autofit
		.Range( 1, StartNumberOfRows + 1, Cols, .Rows).Stripe
		.Range( 1, StartNumberOfRows + 1, Cols, .Rows).Alignment = acRight
		.Range( 1, StartNumberOfRows + 1, Cols, .Rows).SetBorder acBrdOutline + acBrdInsideVert, acLineNormal
		With .Range( 2, StartNumberOfRows, Cols, StartNumberOfRows )
			.Font.Bold = True
			.Alignment = acRight
			.BackColor = 14745599
			.SetBorder acBrdGrid
		End With
		cRow = .InsertRow2
		cRow = .InsertRow2
		.Range (1,cRow,2,cRow).MergeCells = True
		.Cell(cRow,1).Value = "Сѓѕурыђх№:___________________"
		.Range (3,cRow,6,cRow).MergeCells = True
		If BuhName <>"" Then
			.Cell(cRow,3).Value ="/ "&BuhName&" /"
		End If 
	End With
End Sub
'________________________________________________
Sub BuildSecondaryReport
	Dim max, percent, i, CodeID, FOT, FOT1, FOT2, d, sum1, sum2
	Dim arFonds, Repcalc
	sum1 = 