аЯрЁБс                >  ўџ	               $          &      ўџџџ    %   џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ
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
 CSheetPageSheet1Тхфюьюёђќ эрїшёыхэшщ	   '      61    7                                     Щ   4   4   4   /   S   \   R   >   *   1   1   *   &   &   1   F   '                                                                                                                      ' џџ  CRow џџ  CCell6     џџ (   54><>ABL  =0G8A;5=89                      џџџџ                       " џџџџ                       " џџџџ                       " џџџџ                         " џџџџ                          " џџџџ                          " џџџџ                              џџ    /=20@L  2 0 1 5   3.                       џџџџ                       " џџџџ                       " џџџџ                         џџџџ                         џџџџ                         џџџџ                         џџџџ                        ! џџ+       0G8A;5=85                    ! џџ+       %                     ! џџ+     $"                    " џџ+                        ! џџ+  
   !C<<0                    ! џџ+                        ! џџ+    
   515B                    ! џџ+       @548B                               B>3>:                     "                           "                         "                          "                         "                          "                           "                              џџ    "   !  >1I89  ( 020=A)                     " џџ)                                " џџ   Ц                        " џџ                        " џџ                              " џџ                        " џџ                            " џџ                               џџ    (   !  8=20;84K  ( 020=A)                     " џџ    H                         " џџ                              " џџ                        " џџ                              " џџ                        " џџ                            " џџ                               џџџџ                        џџџџ                       " џџџџ                       " џџџџ                       " џџџџ                       " џџџџ                         џџџџ                         џџџџ                        ! џџ+       #45@60=85                    ! џџ+       %                     ! џџ+    
   !C<<0                    ! џџ+    
   515B                    ! џџ+       @548B                       џџ    
   $                      " џџ          .@                    " џџ    ИЫ                         ! џџ                            ! џџ                               џџ       !  ( 107>2K9)                     " џџ    ЭЬЬЬЬЬ@                    " џџ    8Ў2                         ! џџ                            ! џџ       6 6 1                        џџ        !  ( 1>;L=8G=K5)                     " џџ           @                    " џџ                                ! џџ                            ! џџ                               џџ       !  ( )                     " џџ    ЭЬЬЬЬЬ@                    " џџ                                ! џџ                            ! џџ                               џџ1       !  ( 45:@5B)                     " џџ1           @                    " џџ1                          џџ1                              џџ1                               џџ1       !  ( 8=20;84K)                     " џџ1    ЭЬЬЬЬЬ@                    " џџ1                          џџ1                              џџ1                               џџ    "   >5==K9  A1>@  1 , 5 %                     " џџ          ј?                    " џџ    @                         ! џџ       6 4 2                     ! џџ                                 џџџџ                        " џџџџ                       " џџџџ                             џџџџ                                 џџџџ                          	   % џџ,       >4@0745;5=85                     џџ,  
   :;04                    % џџ+  
   20=A                    % џџ+     $                    % џџ+     !                    5 џџ+     !  ( 45:@5B)                      џџ+     !  ( 8=20;84K)                     5 џџ+     >5==K9  A1>@                     џџ+       2K40G5                  	   % џџ-       !>B@C4=8:                     џџ-                       " џџ	                       % џџ+                       % џџ+                        џџ	                        џџџџ                      5 џџ	                           џџџџ                    	      3                         "                                  "       Ц                        "      ИЫ                         "      8Ў2                         "                                  "                                  "      @                         "      pjl                      	   " џџ
    0   0?CABO=  N1>2  <8B@V2=0                    " џџ
    @ї                        " џџ
    pђ                        " џџ
    є(                         " џџ
    Xј	                         " џџ
                                " џџ
                                " џџ
    Ќ'                         " џџ
    дЩџџџџџ                  	   " џџ    8   0=8;>2  !5@3V9  >;>48<8@>28G                    " џџ    №њ                        " џџ    Рсф                         " џџ    Ј!                         " џџ    `=                         " џџ                                " џџ                                " џџ    шn                         " џџ    ;гџџџџџ                  	   " џџ    <   ;5:A552  0A8;89  ;048<8@>28G                      " џџ    еC                        " џџ    @ѓ­                         " џџ    '                         " џџ     C                         " џџ                                " џџ                                " џџ    ј                         " џџ    иљнџџџџџ                  	   " џџ    0   C47L  N4<8;0  5>=V4V2=0                    " џџ     -1                        " џџ     z                         " џџ    РІ                         " џџ     e                         " џџ                                " џџ                                " џџ    Рд                         " џџ    шџџџџџ                  	   " џџ    >   >A:0;LGC:  !2VB;0=0  5==04V52=0                    " џџ    К                        " џџ    XO                         " џџ    y                         " џџ    @л                         " џџ                                " џџ                                " џџ    А0                         " џџ    єz№џџџџџ                  	   " џџ    ,   #;VB5=:>  20=  20=>28G                    " џџ     ?Ћ                        " џџ    @@                         " џџ    DD	                         " џџ     N                         " џџ                                " џџ                                " џџ    і                          " џџ    wѓџџџџџ                  	   " џџ    .   C;0G>:  20=  02@8;>28G                    " џџ    @]Ц                         " џџ    P1                         " џџ    №+                         " џџ    Щ                         " џџ                                " џџ                                " џџ     О                          " џџ    hLіџџџџџ                  	   " џџ    :   !<>;0  ;5:A0=4@  ;5:A0=4@>28G                    " џџ                                " џџ     №                         " џџ    (Р                         " џџ    я                          " џџ                                " џџ                                " џџ    c                          " џџ    ,эњџџџџџ                  	   " џџ    8   0@G5=:>  !5@3V9  =0B>;V9>28G                    " џџ                                " џџ     ў6                         " џџ    ѓ                         " џџ    Єњ                         " џџ                                " џџ                                " џџ    Tг                          " џџ    |>ѕџџџџџ                  	   " џџ    4   20=>20  !2VB;0=0  !5@3VW2=0                    " џџ                                " џџ    XХ7                         " џџ    t                         " џџ                             " џџ                                " џџ                                " џџ    ж                          " џџ    lѕџџџџџ                  	   " џџ    .   5=>2  048<  8:>;09>28G                    " џџ                                " џџ    @p1                         " џџ    &                         " џџ    xЧ                         " џџ                                " џџ                                " џџ    иН                          " џџ    Tіџџџџџ                  	   " џџ    6   C@02;52  ;53  >;>48<8@>28G                    " џџ                                " џџ    O                         " џџ    МЅ                         " џџ    РЈ                          " џџ                                " џџ                                " џџ    PF                          " џџ    4kќџџџџџ                  	   " џџ    2   "0@0=>2  3>@  =0B>;V9>28G                    " џџ                                " џџ    O                         " џџ    МЅ                         " џџ    РЈ                          " џџ                                " џџ                                " џџ    PF                          " џџ    4kќџџџџџ                  	   " џџ    @   C@;0:0  >;>48<8@  >=ABO=B8=>28G                    " џџ                                " џџ    H-5                         " џџ    \А                         " џџ    <ъ                         " џџ                                " џџ                                " џџ    LЬ                          " џџ    ѕџџџџџ                  	   " џџ    <   0I5=:>  ;5:A0=4@  =0B>;V9>28G                    " џџ                                " џџ     $                         " џџ    xK                         " џџ    Q                         " џџ                                " џџ                                " џџ                               " џџ    hжјџџџџџ                  	   " џџ    ,   0;NB0  =4@V9  20=>28G                    " џџ                                " џџ    s                         " џџ    №Ь                         " џџ    иr                          " џџ                                " џџ                                " џџ    Ј/                          " џџ    §џџџџџ                  	   " џџ    4   (C;L30  V:B>@  <8:>;09>28G                      " џџ                                " џџ    E                         " џџ    Ш
                         " џџ    h                         " џџ                                " џџ                                " џџ    0                         " џџ     cђџџџџџ                  	   " џџ    2   $04TT2  ;;O  ;5:A0=4@>28G                    " џџ                                " џџ    O                         " џџ    МЅ                         " џџ    РЈ                          " џџ                                " џџ                                " џџ    PF                          " џџ    4kќџџџџџ                         џџџџ :   CE30;B5@: _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _                       џџџџ                         џџџџ <   /   0?CABO=  N1>2L  <8B@852=0  /                       џџџџ                         џџџџ                         џџџџ                                      <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                                2   2      
 <     ш            РРР Ь Arial ђюф F1 ьхэђ ёяшёър 2 _ 43.ad <     ш    Ш           Ь Arial аDb                   
    <     ш   №           Ь Arial Sub ShtBook_OnLoadEnd Su  <     ш    №           Ь Arial шђх F1 ыхээюую §ыхьхэђр 
.av <     ш               Ь Arial шђх F1  (F7)    1      шЭe <     ш           џџр Ь Arial џџ               №ѓђъш    <     ш           ((( Ь Arial џџ               htRange) <     ш           џџџ Ь Arial џџ                   ing  <     ш            ((( Ь Arial џџ                   а#g <     ш   Ш           Ь Arial                      шЭe	 <     ш    Д           Ь Arial ющёђт ћ (F7) ёюёђюџэшџ фыџ ю4    џџџ                               ШШШ                               ШШШ                                  ммм                                  ммм                              џџр                                  џџџџ                             џџр                              џџџџ                              џџџџ                              	 џџџџ                             
 рџџ                               џџџ                                рџџ                                рџџ                               рџџ                              џџџ                                џџџ                                 рџџ                                 џџџ                                џџџ                               џџџџ                               џџџџ                                џџџџ                               РРР                                  ШШШ                               РРР                               ШШШ                               РРР                              џџџ                                  џџџ                              РРР                              . џџџџ                              - ццњ                              , ццњ                              + ццњ                             ' џџр                              & ррр                               # ррр                             " ррр                                 ! џџр                                џџр                              $ ррр                              % ррр                              ( рџџ                                 ) џџџџ                             * џџџџ                               / џџџџ   РРР    РРР    РРР    РРР 0 ццњ                                 1 џџџџ   ddd    ddd    ddd    ddd 2 џџр                              3 џџр    ddd    ddd    ddd    ddd     џџџџ                             џџр                              џџџџ                              џџR o o t   E n t r y                                               џџџџџџџџ                               ASб'         C o n t e n t s                                                  џџџџ   џџџџ                                    (   g       S u m m a r y I n f o r m a t i o n                           (  џџџџџџџџџџџџ                                        q                                                                          џџџџџџџџџџџџ                                                                        	   
                                                                      !   "   #   ўџџџўџџџ§џџџўџџџўџџџ)   *   +   ,   -   .   /   0   1   2   3   4   5   6   7       џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ               ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџўџ   рђљOhЋ +'Гй   рђљOhЋ +'Гй0   A        x            Ќ      И      Ф      а      м   	   ш   
   є                     ,      4     у     #   Тхфюьюёђќ эрїшёыхэшщ2011_avans.ash                                                          546 @   №ЎC§  @   @ћqASб@   @kчєЖ§Ц      Ръіхэђ 7.0                                                                                                                                                                   џР$Option Explicit
'#include "definition.avb"

Const StartNumberOfRows = 19
Const LastRow = 9
Const MAX_COLUMNS = 9

Dim cRow

Const MTD_AVANS_NAME 	= "3020"

Const MTD_SALARY_NAME = "0117"
Const MTD_NDFL_NAME 	= "2400р"
Const MTD_ESV_NAME 		= "2410р"
Const MTD_ESVB_NAME		= "2420р"
Const MTD_ESVT_NAME 	= "2430р"

Const MTD_BABY_NAME 	= "2450р"
Const MTD_INV_NAME 		= "2460р"

Const MTD_ARMY_NAME 	= "2500р"		'тюхээћщ ёсю№ 1,5%

Dim MTD_SALARY_ID  
Dim MTD_NDFL_ID 	
Dim MTD_ESV_ID 	
Dim MTD_ESVB_ID	
Dim MTD_ESVT_ID 	
Dim MTD_AVANS_ID		
Dim MTD_ARMY_ID 	
Dim MTD_BABY_ID
Dim MTD_INV_ID

Dim BuhName
'-----------------------------------------------------------------------------------------------
Sub ShtBook_OnLoad
	Dim ut
	Set ut = Workarea.UTable(USRTBL_ONE)
	BuhName =ut.GetValue(1, "дШЮ сѓѕурыђх№р", 2)

	With Sheet( 1 )
		.DisplayGrid = False
		.DisplayHeadings = True
		.Columns = MAX_COLUMNS
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
	MTD_ARMY_ID 		= Workarea.GetCodeID(MTD_ARMY_NAME)

	MTD_BABY_ID 		= Workarea.GetCodeID(MTD_BABY_NAME)
	MTD_INV_ID 		= Workarea.GetCodeID(MTD_INV_NAME)

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
	Sheet1.AutoFit 1 + 2
	Sheet1.DisplayHeadings = False
End Sub
'-----------------------------------------------------------------------------------------------
Sub BuildMainReport
	Dim i, cols
	Dim repCC, j
	Set repCC = WorkArea.CreateIntObject("CommonCalc")
	repcc.setPeriod "month", Workarea.Period.Month, Workarea.Period.Year

	cols = 7
	cRow = StartNumberOfRows
	With Sheet( 1 )
		.Rows = StartNumberOfRows
		' юїшёђшыш ёђ№юъѓ шђюуют
		For i = 2 To Cols
			.Cell( StartNumberOfRows, i ).Value = 0.0
		Next
	End With

	With Sheet( 1 )
		.Cell( StartNumberOfRows, 3).Value = repCC.TotalSumBy ("CodeSumW", MTD_SALARY_ID) 'MTD_AVANS_ID
		.Cell( StartNumberOfRows, 4).Value = repCC.TotalSumBy ("CodeSumW", MTD_NDFL_ID)
		.Cell( StartNumberOfRows, 5).Value = repCC.TotalSumBy ("CodeSumW", MTD_ESV_ID) + _
																		repCC.TotalSumBy ("CodeSumW", MTD_ESVB_ID) + _
																		repCC.TotalSumBy ("CodeSumW", MTD_ESVT_ID)
		.Cell( StartNumberOfRows, 6).Value = 	repCC.TotalSumBy ("CodeSumW", MTD_BABY_ID)
		.Cell( StartNumberOfRows, 7).Value = 	repCC.TotalSumBy ("CodeSumW", MTD_INV_ID)
		.Cell( StartNumberOfRows, 8).Value = repCC.TotalSumBy ("CodeSumW", MTD_ARMY_ID)
		.Cell(StartNumberOfRows, 9).Value = .Cell(StartNumberOfRows, 3).Value - _
																.Cell(StartNumberOfRows, 4).Value - _
																.Cell(StartNumberOfRows, 5).Value - _
																.Cell(StartNumberOfRows, 6).Value - _
																.Cell(StartNumberOfRows, 7).Value - _
																.Cell(StartNumberOfRows, 8).Value

		.Range( 1, StartNumberOfRows, .Columns, StartNumberOfRows ).Font.Bold = True

		For i = 1 To repcc.CountAgent
			cRow = .InsertRow2
			.Cell(cRow, 1 ).CellParam = 1
			.Cell(cRow, 1).Value = repCC.Agent(i).Name
			.Cell(cRow, 2).Value = repCC.Agent(i).Prop("Юъырф")
			.Cell(cRow, 3).Value = repCC.SumBy("CodeSumW", i, MTD_SALARY_ID) 'MTD_AVANS_ID
			.Cell(cRow, 4).Value = repCC.SumBy("CodeSumW", i, MTD_NDFL_ID)

			.Cell(cRow, 5).Value = repCC.SumBy("CodeSumW", i, MTD_ESV_ID) + _
																		repCC.SumBy("CodeSumW", i, MTD_ESVB_ID) + _
																		repCC.SumBy("CodeSumW", i, MTD_ESVT_ID)
			.Cell(cRow, 6).Value = repCC.SumBy("CodeSumW", i, MTD_BABY_ID)
			.Cell(cRow, 7).Value = repCC.SumBy("CodeSumW", i, MTD_INV_ID)
			.Cell(cRow, 8).Value = repCC.SumBy("CodeSumW", i, MTD_ARMY_ID)
			' totals
			.Cell(cRow, 9).Value = .Cell(cRow, 3).Value

			For j = 4 To 8
				.Cell(cRow, 9).Value = .Cell(cRow, 9).Value - .Cell(cRow, j).Value
			Next

			If .Cell(cRow, 3).Value=0 And .Cell(cRow, 4).Value=0 And .Cell(cRow, 5).Value=0 Then
				 '.Range(1,cRow,cols,cRow).RowHeight = 0
				.DeleteRow(cRow)
				cRow = cRow - 1
			End If
		Next
		.Range( 1, StartNumberOfRows - 2, .Columns, .Rows).Autofit
		.Range( 1, StartNumberOfRows + 1, .Columns, .Rows).Stripe
		.Range( 1, StartNumberOfRows + 1, .Columns, .Rows).Alignment = acRight
		.Range( 1, StartNumberOfRows + 1, .Columns, .Rows).SetBorder acBrdOutline + acBrdInsideVert, acLineNormal

		With .Range( 2, StartNumberOfRows, .Columns, StartNumberOfRows )
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
	sum1 = 0
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

	arFonds = Array( MTD_NDFL_ID, MTD_ESV_ID, MTD_ESVB_ID, MTD_ESVT_ID, MTD_BABY_ID, MTD_INV_ID, MTD_ARMY_ID)

	If Workarea.Period.Year < 2016 Then
		percent  = Array( 15, 3.6, 3.6, 2, 2.6, 2, 1.5)
	Else
		percent  = Array( 18, 0, 0, 0, 0, 0, 1.5)
	End If

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