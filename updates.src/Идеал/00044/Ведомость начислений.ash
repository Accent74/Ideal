аЯрЁБс                >  ўџ	               %          '      ўџџџ    &   џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџCodeID, Repcalc, d, max )
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
'---
'
'---
Sub Calc2010(CodeID, Repcalc, d, max)
	Dim i
	Dim AgSum1, AgSum2
	Dim flAllowDogovor

'	With Workarea.Code(CodeID)
'		flAllowDogovor = Iif( ((.Mode And 8) = 0) Or (.Type <> 103), False, True )
'	End With

	For i = 1 To Repcalc.Count
'		If flAllowDogovor Then
			With Repcalc.Agent(i)

				AgSum1 = CalcSumByPeriod(.ID, CodeID, CDate("01/01/2009"), CDate("12/01/2009"))
				AgSum2 = CalcSumByPeriod(.ID, CodeID, CDate("13/01/2009"), CDate("31/01/2009"))

				If (AgSum1 + AgSum2) > max Then AgSum2 = max - AgSum1

				If .Prop(PROP_TRUD) Then
					If .Prop("бюёђюшђ т ђ№ѓфютћѕ юђэюјхэшџѕ") Then
						Sheet1.Cell(8, 3).Value = Sheet1.Cell(8, 3).Value + AgSum2
					Else
						Sheet1.Cell(9, 3).Value = Sheet1.Cell(9, 3).Value + AgSum2
					End If
				Else
					Sheet1.Cell(7, 3).Value = Sheet1.Cell(7, 3).Value + AgSum1 + AgSum2
					Sheet1.Cell(7, 5).Value = Sheet1.Cell(7, 5).Value + round2(AgSum1 * 0.013, 2) + round2(AgSum2 * 0.016, 2)
				End If
			End With
'		End If
	Next

End Sub

'---
'
'---
Sub Calc2030(CodeID, Repcalc, d, max)
	Dim i
	Dim AgSum1, AgSum2

	Sheet1.Cell(11, 3).Value = 0
	Sheet1.Cell(11, 5).Value = 0

	For i = 1 To Repcalc.Count
		With Repcalc.Agent(i)
			If Not .Prop(PROP_TRUD, d) Then
				AgSum1 = CalcSumByPeriod(.ID, CodeID, CDate("01/01/2009"), CDate("12/01/2009"))
				AgSum2 = CalcSumByPeriod(.ID, CodeID, CDate("13/01/2009"), CDate("31/01/2009"))

				If (AgSum1 + AgSum2) > max Then AgSum2 = max - AgSum1

				Sheet1.Cell(11, 3).Value = Sheet1.Cell(11, 3).Value + AgSum1 + AgSum2
				Sheet1.Cell(11, 5).Value = Sheet1.Cell(11, 5).Value + round2(AgSum1 * 0.015, 2) + round2(AgSum2 * 0.014, 2)
			End If
		End With
	Next

End Sub<     ш                Ь Arial D:\STDs.!!!\Zap\Шфхры\Шфхры       џџџџ                 џџ 
 CSheetPageSheet1Тхфюьюёђќ эрїшёыхэшщ         61    7                                     Ю   .   4   +   4   +   %   %   %   +   >   4   *   &   &   1   F                                                                                      џџ  CRow џџ  CCell6     џџ   (   54><>ABL  =0G8A;5=89                        џџ      /=20@L  2 0 0 9   3.                      ! џџ#       0G8A;5=85                    ! џџ#       %                     ! џџ#     $"                    ! џџ#                        ! џџ#  
   !C<<0                    ! џџ#                        ! џџ#    
   515B                    ! џџ#       @548B                               B>3>:                                                 "                         "                         =Sum(E5:E11)"    @E                        "                          "                           "                              џџ    .   5=A8>==K9  D>=4  ( 3 3 . 2 % )                       џџ    р                         " џџ   lм                        " џџ   ї(\ТЫ@                    " џџ  hffffц@                    " џџ                        " џџ       9 2                     " џџ       6 5 1                        џџ    (   5=A8>==K9  D>=4  ( 4 % )                       џџ    @                          " џџ   6n                        " џџ   уzЎGс3@                    " џџ       X@                    " џџ                        " џџ       9 2                     " џџ       6 5 1                        џџ    .   $>=4  157@01>B8FK  ( 1 . 6 % )                       џџ    Ш2                          " џџ   lм                        " џџ   иЃp=
з4@                    " џџџџ !ыQИr@                    " џџ                        " џџ       9 2                     " џџ       6 5 3                      " џџ         ( 1 , 6   % )                       џџ    >                          " џџ   sU                         " џџ   Ѕp=
з#(@                   =round2(B8*C8/100, 2)" џџ  эQИы!@                    " џџ                              =G7" џџ       9 2                    =H7" џџ       6 5 3                      " џџ         ( 2 , 2 % )                       џџ    №U                          " џџ   цЊ                         " џџ                       =round2(B9*C9/100, 2)" џџ  Ѕp=
зЃ8@                    " џџ                       =G7" џџ       9 2                    =H7" џџ       6 5 3                        џџ    2   5AG0AB=K9  A;CG09  ( 1 . 0 2 % )                       џџ    и'                          " џџ   lм                        " џџ   Ѕp=
з#(@                    " џџ  |ЎGсzH@                    " џџ                               " џџ       9 2                     " џџ       6 5 6                        џџ    2   5B@C4>A?>A>1=>ABL  ( 1 . 4 % )                       џџ    А6                          " џџ   lм                        " џџ   53333s0@                    " џџ  *\Тѕ(Q@                    " џџ                               " џџ       9 2                     " џџ       6 5 2                        % џџ$       >4@0745;5=85                    % џџ$    
   !C<<0                      џџ#     0G8A;5=>                      џџ"                          џџ#                          џџ#     #45@60=>                      џџ#                          џџ#                          џџ"                          џџ"                          џџ#       2K40G5                     % џџ%       !>B@C4=8:                    ! џџ%       ;L3>BK                      џџ#       0 1 1 0                       џџ#       0 5 5 0                       џџ#    
   A53>                      џџ#       2 1 0 0                       џџ#       2 1 1 0                       џџ#       2 1 2 0                       џџ#       2 1 3 0                       џџ#    
   A53>                      џџ"                               
   B>3>                    "                                  " џџ     ЂJ                        " џџ    Xы9                         " џџ    X                        " џџ    XV                         " џџ     І                         " џџ     т                         " џџ    Рj                         " џџ    {l                         " џџ    @                         "  
       "                     "  
                         "  
     6n                        "  
    Xы9                         "  
    X!Ј                        "  
                          "  
                          "  
                          "  
                          "  
                          "  
    X!Ј                         " џџ    2   >;4V=  5==04V9   ><0=>28G                    " џџ                          " џџ                         " џџ    Xы9                         " џџ    Xы9                         " џџ                         " џџ                         " џџ                         " џџ                         " џџ                          " џџ    Xы9                          " џџ       20=>2                    " џџ                          " џџ     z                         " џџ                         " џџ     z                         " џџ                         " џџ                         " џџ                         " џџ                         " џџ                          " џџ     z                          " џџ    ,   5B@>2  <8B@>  20=>28G                    " џџ                          " џџ     z                         " џџ                         " џџ     z                         " џџ                         " џџ                         " џџ                         " џџ                         " џџ                          " џџ     z                          " џџ    4    09G5=:>  ;53  8:>;09>28G                      " џџ                          " џџ     z                         " џџ                         " џџ     z                         " џџ                         " џџ                         " џџ                         " џџ                         " џџ                          " џџ     z                          "         @01>G85                    "                           "       lм                        "                            "       lм                        "      XV                         "       І                         "       т                         "      Рj                         "      {l                         "      ш№o                         " џџ    0   =35;LT2  =4@V9  02;>28G                    " џџ    h(.                         " џџ     z                         " џџ                         " џџ     z                         " џџ    АТ
                         " џџ     q                         " џџ    8                         " џџ    @                          " џџ    pя                         " џџ    "k                          " џџ    2   V=>3@04>2  235=  0@?>28G                    " џџ                          " џџ     z                         " џџ                         " џџ     z                         " џџ    XЏ                         " џџ     q                         " џџ    8                         " џџ    @                          " џџ    м                         " џџ    ш5d                          " џџ    B   >;>1>@>4L:>  >;>48<8@  0A8;L>28G                    " џџ    h(.                         " џџ     z                         " џџ                         " џџ     z                         " џџ    ЈЛ
                         " џџ     q                         " џџ    8                         " џџ     В                          " џџ    H                         " џџ    Ињj                          " џџ    0   >=G0@C:  !5@3V9  5B@>28G                    " џџ                          " џџ     z                         " џџ                         " џџ     z                         " џџ    O                         " џџ     q                         " џџ                         " џџ                         " џџ    Р                         " џџ    Qe                          " џџ    0   AV?>2  ;5:A0=4@  L2>28G                    " џџ    h(.                         " џџ     z                         " џџ                         " џџ     z                         " џџ    ЈЛ
                         " џџ     q                         " џџ    8                         " џџ     В                          " џџ    H                         " џџ    Ињj                          " џџ    4   @0?:>  <8B@>  =0B>;V9>28G                    " џџ                          " џџ     z                         " џџ                         " џџ     z                         " џџ    O                         " џџ     q                         " џџ                         " џџ                         " џџ    Р                         " џџ    Qe                       d   d   d   d    <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                                2   2   
   	 <     ш    Д           Ь Arial ющёђт ћ (F7) ёюёђюџэшџ фыџ ю <     ш   Ш           Ь Arial                      шЭe <     ш            ((( Ь Arial џџ                   а#g <     ш           џџџ Ь Arial џџ                   ing  <     ш           ((( Ь Arial џџ               htRange) <     ш           џџр Ь Arial џџ               №ѓђъш    <     ш               Ь Arial шђх F1  (F7)    1      шЭe  <     ш    №           Ь Arial шђх F1 ыхээюую §ыхьхэђр 
.av <     ш   №           Ь Arial Sub ShtBook_OnLoadEnd Su <     ш    Ш           Ь Arial аDb                   
   -    рџџ                              рџџ                               рџџ                                џџџ                               
 рџџ                              	 џџџџ                              џџџџ                               џџџџ                              џџр                              џџџџ                             џџр                                  ммм                              ммм                                   ШШШ                                  ШШШ                              џџџ                               РРР                               џџџ                              џџџ                                  РРР                              ШШШ                               РРР                               ШШШ                               РРР                                  џџџџ                               џџџџ                                џџџџ                               џџџ                               џџџ                                рџџ                                 џџџ                                 џџџ                               ( рџџ                                 % ррр                              $ ррр                                џџр                              ! џџр                              " ррр                                 # ррр                             & ррр                               ' џџр                              ) Рџџ                                 * Џюю                                 + џџџџ                               , рџџ                                    ш           џџр Ь Arial џџ               №ѓђъш    <     ш               Ь Arial шђх F1  (F7)    1      шЭe  <     ш    №           Ь Arial шђх F1 ыхээюую §ыхьхэђр 
.av <     ш   №           Ь Arial Sub ShtBook_OnLoadEnd Su <     ш    Ш           Ь Arial аDb                   
   -    рџџ                              рџџ                               рџџ                                џџџ        R o o t   E n t r y                                               џџџџџџџџ                               T61Щ$         C o n t e n t s                                                  џџџџ   џџџџ                                    <   7       S u m m a r y I n f o r m a t i o n                           (  џџџџџџџџџџџџ                                        i                                                                          џџџџџџџџџџџџ                                                                        	   
                                                                      ўџџџўџџџ§џџџўџџџўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ=   >   ?   @   A   B   C   D   E   F   G   H   I   J   K   L   M   N   O   P   Q   R   S   T   U   V   W   X   Y   Z   [       џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ               ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџўџ   рђљOhЋ +'Гй   рђљOhЋ +'Гй0   9        x            Є      А      М      Ш      д   	   р   
   ь      ј               $      ,     у        Тхфюьюёђќ эрїшёыхэшщ.ash                                                            446 @   №фѕ  @   PЖ<61Щ@   @kчєЖ§Ц      Ръіхэђ 7.0                                                                                                                                                                           R o o t   E n t r y                                               џџџџџџџџ                                bу1Щ(         C o n t e n t s                                                  џџџџ   џџџџ                                    )   M       S u m m a r y I n f o r m a t i o n                           (  џџџџџџџџџџџџ                                        i                                                                          џџџџџџџџџџџџ                                                                        	   
                                                                      ўџџџџџџџџџџџџџџџџџџџўџџџ§џџџўџџџўџџџ*   +   ,   -   .   /   0   1   2   3   4   5   6   7   8   9   :   ;   \   џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ]   ^   _   `   a   b   c   d   e   f   g   h       џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ               ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџўџ   рђљOhЋ +'Гй   рђљOhЋ +'Гй0   9        x            Є      А      М      Ш      д   	   р   
   ь      ј               $      ,     у        Тхфюьюёђќ эрїшёыхэшщ.ash                                                            447 @   рLі  @   @рGу1Щ@   @kчєЖ§Ц      Ръіхэђ 7.0                                                                                                                                                                           џJOption Explicit
'#include "definition.avb"

Const adCmdText = 1
Const adParamInput = 1

Const StartNumberOfRows = 15
Const StartNumberOfColumns = 2

Const MTC_PRIVILEGE_NAME = "0801"
Const MTC_WPF_NAME = "дЮв"

Dim cRow

Dim DicCodes
Dim TotalsCols( 2 )
Dim arCodes
Dim PrivilegeID, WPF_ID

'-----------------------------------------------------------------------------------------------
Sub ShtBook_OnLoad
	PrivilegeID = WorkArea.GetCodeID(MTC_PRIVILEGE_NAME)
	WPF_ID = GetCodeID( MTC_WPF_NAME )
	Set DicCodes = CreateObject("Scripting.Dictionary")
	arCodes = GetCodes
	With Sheet( 1 )
		.DisplayGrid = False
		.DisplayHeadings = True
	End With
	EnableButtons 1
	PeriodText = WorkArea.Period.RepTitle
	Sheet( 1 ).Cell( 2, 1 ).Value = WorkArea.Period.RepTitle
	BuildReport
End Sub

'________________________________________________
Sub ShtBook_OnRequery
	BuildReport
	recalc
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
	Sheet1.DisplayHeadings = False
End Sub

'________________________________________________
' Function GetCodes
' тћсш№рхь ъюфћ, чр шёъыўїхэшхь ёяхі-ъюфют "тћяырђр", "фюыу ёюђ№ѓфэшър" ш "фюыу я№хфя№шџђшџ"
'________________________________________________
Function GetCodes
	Dim SqlQr, ar
	Select Case App.AppType
	Case "DAO"
		SqlQr = "SELECT MTD_ID, MTD_CODE, 0 FROM MTD_CODES WHERE ( MTD_TYPE = 101 ) AND ( IIF( MTD_MODE IS NULL, 0, MTD_MODE ) <> 2 ) ORDER BY 3, 2" & _
				vbNewLine & "UNION" & vbNewLine & _
				"SELECT MTD_ID, MTD_CODE, 1 FROM MTD_CODES WHERE ( MTD_TYPE = 102 ) AND ( IIF( MTD_MODE IS NULL, 0, MTD_MODE ) <> 3 ) AND ( IIF( MTD_MODE IS NULL, 0, MTD_MODE ) <> 1 ) ORDER BY 3, 2"
		GetRows_DAO SqlQr, ar
		GetCodes = ar
	Case "OLEDB"
		SqlQr = "SELECT MTD_ID, MTD_CODE, 0 FROM MTD_CODES WHERE ( MTD_TYPE = 101 ) AND ( ISNULL( MTD_MODE, 0 ) <> 2 ) ORDER BY 3, 2" & _
				vbNewLine & "UNION" & vbNewLine & _
				"SELECT MTD_ID, MTD_CODE, 1 FROM MTD_CODES WHERE ( MTD_TYPE = 102 ) AND ( ISNULL( MTD_MODE, 0 ) <> 3 ) AND ( ISNULL( MTD_MODE, 0 ) <> 1 ) ORDER BY 3, 2"
		GetRows_ADO SQLText, adCmdText, ar
		GetCodes = ar
	End Select
End Function

'-----------------------------------------------------------------------------------------------
Sub BuildMainReport
	Dim GroupID, AgID, PeriodID, k, i, n
	Dim hdr, RecSet, arData

	DicCodes.RemoveAll

	cRow = StartNumberOfRows
	With Sheet( 1 )
		.Rows = StartNumberOfRows
		.Columns = 8
		' юїшёђшыш ёђ№юъѓ шђюуют
		For i = 2 To .Columns
			.Cell( StartNumberOfRows, i ).Value = 0.0
		Next
	End With

	PeriodID = WorkArea.Period.ID
	Set RecSet = GetData( PeriodID )
	If RecSet.EOF Then Exit Sub

	Select Case App.AppType
	Case "DAO"
		RecSet.MoveLast
		n = RecSet.RecordCount
		RecSet.MoveFirst
		arData = RecSet.GetRows( n )
	Case Else
		arData = RecSet.GetRows
	End Select

	' тћс№рыш тёх ъюфћ, яю ъюђю№ћь хёђќ юсю№юђћ
	For i = 0 To UBound( arData, 2 )
		DicCodes( arData( 4, i )) = 0
	Next

	With Sheet( 1 )
		If ( DicCodes.Count + 2 ) > 8 Then
			.Columns = 2 + DicCodes.Count
		End If

		TotalsCols( 0 ) = 2
		' ашёѓхь јряъѓ ё ях№хїэхь ъюфют эрїшёыхэшщ ш ѓфх№црэшщ
		k = 3
		' arCodes( 2, i ) = 0 шыш 1. Эѓцэю фыџ ђюую, їђюс т ѓфх№црэшџѕ ёфхырђќ ёфтшу эр юфшэ ёђюысхі,
		' їђюсћ юёђртшђќ ьхёђю фыџ ёюысір шђюуют яю ѓфх№црэшџь
		hdr = "1:Эрїшёыхэю"
		For i = 0 To UBound( arCodes, 2 )
			If i = 0 Then
			ElseIf ( arCodes( 2, i - 1 ) = 0 ) And ( arCodes( 2, i ) = 1 ) Then
				hdr = hdr & vbLf & "2:Тёхую" & vbLf & "1:гфх№црэю"
				TotalsCols( 1 ) = k
			End If
			If DicCodes.Exists( arCodes( 0, i )) Then
				DicCodes( arCodes( 0, i )) = k + arCodes( 2, i )
				hdr = hdr & vbLf & "2:" & arCodes( 1, i )
				k = k + 1
			End If
		Next
		hdr = hdr & vbLf & "2:Тёхую" & vbLf & "1:Ъ тћфрїх"

		.MakeHeader StartNumberOfRows - 2, 3, hdr
		TotalsCols( 2 ) = .Columns - 1

		For i = 0 To UBound( arData, 2 )
			If GroupID <> arData( 0, i ) Then
				cRow = .InsertRow2

				.Cell( cRow, 1 ).CellParam = 1
				.Range( 1, cRow, .Columns, cRow ).Font.Bold = True

				.Cell(cRow, 1).Value = arData( 2, i )
				GroupID = arData( 0, i )
			End If

			If AgID <> arData( 1, i ) Then
				cRow = .InsertRow2

				.Cell( cRow, 1 ).Value = arData( 3, i )
				.Cell( cRow, 2 ).Value = GetPrivilege( PrivilegeID, arData( 1, i ), PeriodID )
				AgID = arData( 1, i )
			End If
			.Cell( cRow, DicCodes( arData( 4, i ))).Value = arData( 5, i )
		Next
		CalcTotals
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
	End With
End Sub

'-----------------------------------------------------------------------------------------------
Function GetData(PeriodID)
	Dim SqlQr

	Select Case App.AppType
	Case "DAO"
		With WorkArea.DAODataBase.CreateQueryDef("")
			.Sql = "PARAMETERS PeriodID Long; " & _
			       "Select AG_TREE.P0, P_JOURNAL.AG_ID, Groups.AG_NAME, AGENTS.AG_NAME, MTC_ID, Sum(JP_SUM) " & _
			       "From MTD_CODES Inner Join " & _
			       "    (P_JOURNAL Inner Join " & _
			       "       (AGENTS Inner Join " & _
			       "      (AG_TREE Inner Join AGENTS Groups On AG_TREE.P0 = Groups.AG_ID) " & _
			       "                                        On AGENTS.AG_ID = AG_TREE.ID) " & _
			       "                                        On P_JOURNAL.AG_ID = AGENTS.AG_ID) " & _
			       "                                        On MTD_CODES.MTD_ID = P_JOURNAL.MTC_ID " & _
			       "Where JP_DONE = 1 And SHORTCUT = 0 And ( IIF( MTD_MODE IS NULL, 0, MTD_MODE ) <> 1 ) And ( IIF( MTD_MODE IS NULL, 0, MTD_MODE ) <> 2 ) And ( IIF( MTD_MODE IS NULL, 0, MTD_MODE ) <> 3 ) " & _
			       "                  And ( IIF( MTD_MASK IS NULL, 0, MTD_MASK ) < 8 )" & _
			       "                  And C_PERIOD = [PeriodID] " & _
			       "Group By AG_TREE.P0, P_JOURNAL.AG_ID, Groups.AG_NAME, AGENTS.AG_NAME, MTC_ID " & _
			       "Having Sum(JP_SUM) <> 0 " & _
			       "Order By Groups.AG_NAME, AGENTS.AG_NAME "

			.Parameters(0).Value = PeriodID

			Set GetData = .OpenRecordSet
		End With
	Case "OLEDB"
		SqlQr = "Select AG_TREE.P0, P_JOURNAL.AG_ID, Groups.AG_NAME, AGENTS.AG_NAME, MTC_ID, Sum(JP_SUM) " & _
		        "From P_JOURNAL Inner Join AGENTS On P_JOURNAL.AG_ID = AGENTS.AG_ID " & _
		        "               Inner Join AG_TREE On P_JOURNAL.AG_ID = ID " & _
		        "               Inner Join MTD_CODES On MTC_ID = MTD_ID " & _
		        "               Left Join AGENTS Groups On P0 = Groups.AG_ID " & _
		        "Where JP_DONE = 1 And SHORTCUT = 0 AND ( ISNULL( MTD_MODE, 0 ) <> 1 ) AND ( ISNULL( MTD_MODE, 0 ) <> 2 ) AND ( ISNULL( MTD_MODE, 0 ) <> 3 ) " & _
		        "                  And ( IsNull( MTD_MASK, 0 ) < 8 ) " & _
		        "                  And C_PERIOD = ? " & _
		        "Group By AG_TREE.P0, P_JOURNAL.AG_ID, Groups.AG_NAME, AGENTS.AG_NAME, MTC_ID " & _
		        "Having Sum(JP_SUM) <> 0 " & _
		        "Order By Groups.AG_NAME, AGENTS.AG_NAME "

		Set GetData = QueryWithParams( SqlQr, adCmdText, Array( adInteger, PeriodID ))
	End Select
End Function

'-----------------------------------------------------------------------------------------------
Function GetPrivilege(MtcID, AgID, PeriodID)
	Dim RecSet
	Dim SqlQr

	Select Case App.AppType
		Case "DAO"
			With WorkArea.DAODataBase.CreateQueryDef("")
				.Sql = "PARAMETERS MtcID Long, AgID Long, PeriodID Long; " & _
				       "Select JP_S1 From P_JOURNAL " & _
				       "Where JP_DONE = 1 And (Not IsNull(JP_S1)) And " & _
				       "MTC_ID = [MtcID] And AG_ID = [AgID] And C_PERIOD = [PeriodID]"

				.Parameters(0).Value = MtcID
				.Parameters(1).Value = AgID
				.Parameters(2).Value = PeriodID

				Set RecSet = .OpenRecordSet
			End With
		Case "OLEDB"
			SqlQr = "Select JP_S1 From P_JOURNAL " & _
			        "Where JP_DONE = 1 And IsNull(JP_S1, 0) <> 0 And " & _
			        "MTC_ID = ? And AG_ID = ? And C_PERIOD = ? "
			Set RecSet = QueryWithParams( SqlQr, adCmdText, Array( adInteger, MtcID, adInteger, AgID, adInteger, PeriodID ))
	End Select

	If RecSet.EOF Then
		GetPrivilege = 0
	Else
		GetPrivilege = RecSet.Fields(0).Value
	End If
End Function

'-----------------------------------------------------------------------------------------------
Sub CalcTotals
	Dim i, j
	ReDim GroupTotals(Sheet( 1 ).Columns)
	ReDim Totals(Sheet( 1 ).Columns)

	CalcTotalsInRows

	SetZero Totals
	SetZero GroupTotals

	For i = Sheet( 1 ).Rows To StartNumberOfRows + 1 Step -1
		If Sheet( 1 ).Cell(i, 1).CellParam = 1 Then
			For j = StartNumberOfColumns + 1 To Sheet( 1 ).Columns
				Sheet( 1 ).Cell(i, j).Value = GroupTotals(j)
				Totals(j) = Totals(j) + GroupTotals(j)
			Next

			SetZero GroupTotals
		Else
			For j = StartNumberOfColumns + 1 To Sheet( 1 ).Columns
				GroupTotals(j) = GroupTotals(j) + Sheet( 1 ).Cell(i, j).Value
			Next
		End IfџxJOption Explicit
'#include "definition.avb"

Const adCmdText = 1
Const adParamInput = 1

Const StartNumberOfRows = 15
Const StartNumberOfColumns = 2

Const MTC_PRIVILEGE_NAME = "0801"
Const MTC_WPF_NAME = "дЮв"

Dim cRow

Dim DicCodes
Dim TotalsCols( 2 )
Dim arCodes
Dim PrivilegeID, WPF_ID

'-----------------------------------------------------------------------------------------------
Sub ShtBook_OnLoad
	PrivilegeID = WorkArea.GetCodeID(MTC_PRIVILEGE_NAME)
	WPF_ID = GetCodeID( MTC_WPF_NAME )
	Set DicCodes = CreateObject("Scripting.Dictionary")
	arCodes = GetCodes
	With Sheet( 1 )
		.DisplayGrid = False
		.DisplayHeadings = True
	End With
	EnableButtons 1
	PeriodText = WorkArea.Period.RepTitle
	Sheet( 1 ).Cell( 2, 1 ).Value = WorkArea.Period.RepTitle
	BuildReport
End Sub

'________________________________________________
Sub ShtBook_OnRequery
	BuildReport
	recalc
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
	Sheet1.DisplayHeadings = False
End Sub

'________________________________________________
' Function GetCodes
' тћсш№рхь ъюфћ, чр шёъыўїхэшхь ёяхі-ъюфют "тћяырђр", "фюыу ёюђ№ѓфэшър" ш "фюыу я№хфя№шџђшџ"
'________________________________________________
Function GetCodes
	Dim SqlQr, ar
	Select Case App.AppType
	Case "DAO"
		SqlQr = "SELECT MTD_ID, MTD_CODE, 0 FROM MTD_CODES WHERE ( MTD_TYPE = 101 ) AND ( IIF( MTD_MODE IS NULL, 0, MTD_MODE ) <> 2 ) ORDER BY 3, 2" & _
				vbNewLine & "UNION" & vbNewLine & _
				"SELECT MTD_ID, MTD_CODE, 1 FROM MTD_CODES WHERE ( MTD_TYPE = 102 ) AND ( IIF( MTD_MODE IS NULL, 0, MTD_MODE ) <> 3 ) AND ( IIF( MTD_MODE IS NULL, 0, MTD_MODE ) <> 1 ) ORDER BY 3, 2"
		GetRows_DAO SqlQr, ar
		GetCodes = ar
	Case "OLEDB"
		SqlQr = "SELECT MTD_ID, MTD_CODE, 0 FROM MTD_CODES WHERE ( MTD_TYPE = 101 ) AND ( ISNULL( MTD_MODE, 0 ) <> 2 ) ORDER BY 3, 2" & _
				vbNewLine & "UNION" & vbNewLine & _
				"SELECT MTD_ID, MTD_CODE, 1 FROM MTD_CODES WHERE ( MTD_TYPE = 102 ) AND ( ISNULL( MTD_MODE, 0 ) <> 3 ) AND ( ISNULL( MTD_MODE, 0 ) <> 1 ) ORDER BY 3, 2"
		GetRows_ADO SQLText, adCmdText, ar
		GetCodes = ar
	End Select
End Function

'-----------------------------------------------------------------------------------------------
Sub BuildMainReport
	Dim GroupID, AgID, PeriodID, k, i, n
	Dim hdr, RecSet, arData

	DicCodes.RemoveAll

	cRow = StartNumberOfRows
	With Sheet( 1 )
		.Rows = StartNumberOfRows
		.Columns = 8
		' юїшёђшыш ёђ№юъѓ шђюуют
		For i = 2 To .Columns
			.Cell( StartNumberOfRows, i ).Value = 0.0
		Next
	End With

	PeriodID = WorkArea.Period.ID
	Set RecSet = GetData( PeriodID )
	If RecSet.EOF Then Exit Sub

	Select Case App.AppType
	Case "DAO"
		RecSet.MoveLast
		n = RecSet.RecordCount
		RecSet.MoveFirst
		arData = RecSet.GetRows( n )
	Case Else
		arData = RecSet.GetRows
	End Select

	' тћс№рыш тёх ъюфћ, яю ъюђю№ћь хёђќ юсю№юђћ
	For i = 0 To UBound( arData, 2 )
		DicCodes( arData( 4, i )) = 0
	Next

	With Sheet( 1 )
		If ( DicCodes.Count + 2 ) > 8 Then
			.Columns = 2 + DicCodes.Count
		End If

		TotalsCols( 0 ) = 2
		' ашёѓхь јряъѓ ё ях№хїэхь ъюфют эрїшёыхэшщ ш ѓфх№црэшщ
		k = 3
		' arCodes( 2, i ) = 0 шыш 1. Эѓцэю фыџ ђюую, їђюс т ѓфх№црэшџѕ ёфхырђќ ёфтшу эр юфшэ ёђюысхі,
		' їђюсћ юёђртшђќ ьхёђю фыџ ёюысір шђюуют яю ѓфх№црэшџь
		hdr = "1:Эрїшёыхэю"
		For i = 0 To UBound( arCodes, 2 )
			If i = 0 Then
			ElseIf ( arCodes( 2, i - 1 ) = 0 ) And ( arCodes( 2, i ) = 1 ) Then
				hdr = hdr & vbLf & "2:Тёхую" & vbLf & "1:гфх№црэю"
				TotalsCols( 1 ) = k
			End If
			If DicCodes.Exists( arCodes( 0, i )) Then
				DicCodes( arCodes( 0, i )) = k + arCodes( 2, i )
				hdr = hdr & vbLf & "2:" & arCodes( 1, i )
				k = k + 1
			End If
		Next
		hdr = hdr & vbLf & "2:Тёхую" & vbLf & "1:Ъ тћфрїх"

		.MakeHeader StartNumberOfRows - 2, 3, hdr
		TotalsCols( 2 ) = .Columns - 1

		For i = 0 To UBound( arData, 2 )
			If GroupID <> arData( 0, i ) Then
				cRow = .InsertRow2

				.Cell( cRow, 1 ).CellParam = 1
				.Range( 1, cRow, .Columns, cRow ).Font.Bold = True

				.Cell(cRow, 1).Value = arData( 2, i )
				GroupID = arData( 0, i )
			End If

			If AgID <> arData( 1, i ) Then
				cRow = .InsertRow2

				.Cell( cRow, 1 ).Value = arData( 3, i )
				.Cell( cRow, 2 ).Value = GetPrivilege( PrivilegeID, arData( 1, i ), PeriodID )
				AgID = arData( 1, i )
			End If
			.Cell( cRow, DicCodes( arData( 4, i ))).Value = arData( 5, i )
		Next
		CalcTotals
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
	End With
End Sub

'-----------------------------------------------------------------------------------------------
Function GetData(PeriodID)
	Dim SqlQr

	Select Case App.AppType
	Case "DAO"
		With WorkArea.DAODataBase.CreateQueryDef("")
			.Sql = "PARAMETERS PeriodID Long; " & _
			       "Select AG_TREE.P0, P_JOURNAL.AG_ID, Groups.AG_NAME, AGENTS.AG_NAME, MTC_ID, Sum(JP_SUM) " & _
			       "From MTD_CODES Inner Join " & _
			       "    (P_JOURNAL Inner Join " & _
			       "       (AGENTS Inner Join " & _
			       "      (AG_TREE Inner Join AGENTS Groups On AG_TREE.P0 = Groups.AG_ID) " & _
			       "                                        On AGENTS.AG_ID = AG_TREE.ID) " & _
			       "                                        On P_JOURNAL.AG_ID = AGENTS.AG_ID) " & _
			       "                                        On MTD_CODES.MTD_ID = P_JOURNAL.MTC_ID " & _
			       "Where JP_DONE = 1 And SHORTCUT = 0 And ( IIF( MTD_MODE IS NULL, 0, MTD_MODE ) <> 1 ) And ( IIF( MTD_MODE IS NULL, 0, MTD_MODE ) <> 2 ) And ( IIF( MTD_MODE IS NULL, 0, MTD_MODE ) <> 3 ) " & _
			       "                  And ( IIF( MTD_MASK IS NULL, 0, MTD_MASK ) < 8 )" & _
			       "                  And C_PERIOD = [PeriodID] " & _
			       "Group By AG_TREE.P0, P_JOURNAL.AG_ID, Groups.AG_NAME, AGENTS.AG_NAME, MTC_ID " & _
			       "Having Sum(JP_SUM) <> 0 " & _
			       "Order By Groups.AG_NAME, AGENTS.AG_NAME "

			.Parameters(0).Value = PeriodID

			Set GetData = .OpenRecordSet
		End With
	Case "OLEDB"
		SqlQr = "Select AG_TREE.P0, P_JOURNAL.AG_ID, Groups.AG_NAME, AGENTS.AG_NAME, MTC_ID, Sum(JP_SUM) " & _
		        "From P_JOURNAL Inner Join AGENTS On P_JOURNAL.AG_ID = AGENTS.AG_ID " & _
		        "               Inner Join AG_TREE On P_JOURNAL.AG_ID = ID " & _
		        "               Inner Join MTD_CODES On MTC_ID = MTD_ID " & _
		        "               Left Join AGENTS Groups On P0 = Groups.AG_ID " & _
		        "Where JP_DONE = 1 And SHORTCUT = 0 AND ( ISNULL( MTD_MODE, 0 ) <> 1 ) AND ( ISNULL( MTD_MODE, 0 ) <> 2 ) AND ( ISNULL( MTD_MODE, 0 ) <> 3 ) " & _
		        "                  And ( IsNull( MTD_MASK, 0 ) < 8 ) " & _
		        "                  And C_PERIOD = ? " & _
		        "Group By AG_TREE.P0, P_JOURNAL.AG_ID, Groups.AG_NAME, AGENTS.AG_NAME, MTC_ID " & _
		        "Having Sum(JP_SUM) <> 0 " & _
		        "Order By Groups.AG_NAME, AGENTS.AG_NAME "

		Set GetData = QueryWithParams( SqlQr, adCmdText, Array( adInteger, PeriodID ))
	End Select
End Function

'-----------------------------------------------------------------------------------------------
Function GetPrivilege(MtcID, AgID, PeriodID)
	Dim RecSet
	Dim SqlQr

	Select Case App.AppType
		Case "DAO"
			With WorkArea.DAODataBase.CreateQueryDef("")
				.Sql = "PARAMETERS MtcID Long, AgID Long, PeriodID Long; " & _
				       "Select JP_S1 From P_JOURNAL " & _
				       "Where JP_DONE = 1 And (Not IsNull(JP_S1)) And " & _
				       "MTC_ID = [MtcID] And AG_ID = [AgID] And C_PERIOD = [PeriodID]"

				.Parameters(0).Value = MtcID
				.Parameters(1).Value = AgID
				.Parameters(2).Value = PeriodID

				Set RecSet = .OpenRecordSet
			End With
		Case "OLEDB"
			SqlQr = "Select JP_S1 From P_JOURNAL " & _
			        "Where JP_DONE = 1 And IsNull(JP_S1, 0) <> 0 And " & _
			        "MTC_ID = ? And AG_ID = ? And C_PERIOD = ? "
			Set RecSet = QueryWithParams( SqlQr, adCmdText, Array( adInteger, MtcID, adInteger, AgID, adInteger, PeriodID ))
	End Select

	If RecSet.EOF Then
		GetPrivilege = 0
	Else
		GetPrivilege = RecSet.Fields(0).Value
	End If
End Function

'-----------------------------------------------------------------------------------------------
Sub CalcTotals
	Dim i, j
	ReDim GroupTotals(Sheet( 1 ).Columns)
	ReDim Totals(Sheet( 1 ).Columns)

	CalcTotalsInRows

	SetZero Totals
	SetZero GroupTotals

	For i = Sheet( 1 ).Rows To StartNumberOfRows + 1 Step -1
		If Sheet( 1 ).Cell(i, 1).CellParam = 1 Then
			For j = StartNumberOfColumns + 1 To Sheet( 1 ).Columns
				Sheet( 1 ).Cell(i, j).Value = GroupTotals(j)
				Totals(j) = Totals(j) + GroupTotals(j)
			Next

			SetZero GroupTotals
		Else
			For j = StartNumberOfColumns + 1 To Sheet( 1 ).Columns
				GroupTotals(j) = GroupTotals(j) + Sheet( 1 ).Cell(i, j).Value
			Next
		End If
	Next

	For j = StartNumberOfColumns + 1 To Sheet( 1 ).Columns
		Sheet( 1 ).Cell(StartNumberOfRows, j).Value = Totals(j)
	Next
End Sub

'-----------------------------------------------------------------------------------------------
Sub CalcTotalsInRows
	Dim i, j, k
	Dim Sum
	With Sheet( 1 )
		For i = StartNumberOfRows + 1 To .Rows
			For k = 1 To UBound( TotalsCols )
				Sum = 0
				For j = TotalsCols( k - 1 ) + 1 To TotalsCols( k ) - 1
					Sum = Sum + .Cell( i, j ).Value
				Next
				.Cell( i, TotalsCols( k )).Value = Sum
			Next
			.Cell( i, .Columns ).Value = .Cell( i, TotalsCols( 1 ) ).Value - .Cell( i, TotalsCols( 2 )).Value
		Next
	End With
End Sub

'-----------------------------------------------------------------------------------------------
Sub SetZero(Arr)
	Dim i
	For i = 1 To UBound(Arr)
		Arr(i) = 0
	Next
End Sub

'-----------------------------------------------------------------------------------------------
Sub MergeCells(OnRow, CellsText)
	Dim j
	Dim FirstPos, LastPos

	FirstPos = 1
	LastPos = 0

	For j = 1 To Sheet( 1 ).Columns
		If Sheet( 1 ).Cell(OnRow, j).Value = CellsText Then
			FirstPos = j
			Exit For
		End If
	Next

	For j = Sheet( 1 ).Columns To 1 Step -1
		If Sheet( 1 ).Cell(OnRow, j).Value = CellsText Then
			LastPos = j 
			Exit For
		End If
	Next

	If LastPos >= FirstPos Then Sheet( 1 ).Range(FirstPos, OnRow, LastPos, OnRow).MergeCells = True
End Sub

'________________________________________________
Sub BuildSecondaryReport
	Dim max, percent, i, CodeID, FOT, d, sum1
	Dim arFonds, Repcalc
	arFonds = Array( "2000", "2001", "2010", "2020", "2030" )
	Set Repcalc = WorkArea.CreateIntObject("CommonCalc")
	Repcalc.setPeriod "month", Workarea.Period.Month, Workarea.Period.Year
	d = Workarea.Period.FirstDay

	max = Workarea.UTable( USRTBL_LGT ).GetValue( 1, USRTBL_MAXCY, 2 )
	With Sheet( 1 )
		CodeID = WorkArea.GetCodeID( "2000" )
		With WorkArea.Code( CodeID )
			percent = .Percent
			Sheet( 1 ).Cell( 5, 7 ).Value = .DbCode
			Sheet( 1 ).Cell( 5, 8 ).Value = .CrCode
		End With
		.Cell( 5, 2 ).Value = percent
		FOT = CalcDepsSumLimit_Pens( CodeID, Repcalc, max, False )
		sum1 = Round2( FOT * percent / 100 , 2 )
		.Cell( 5, 3 ).value = FOT
		.Cell( 5, 5 ).value = Round2( FOT * percent / 100 , 2 )

		Repcalc.Clear
		CodeID = WorkArea.GetCodeID( "2001" )
		With WorkArea.Code( CodeID )
			percent = .Percent
			Sheet( 1 ).Cell( 6, 7 ).Value = .DbCode
			Sheet( 1 ).Cell( 6, 8 ).Value = .CrCode
		End With
		.Cell( 6, 2 ).Value = percent
		FOT = CalcDepsSumLimit_Pens( CodeID, Repcalc, max, True )
		sum1 = sum1 + Round2( FOT * percent / 100 , 2 )
		.Cell( 6, 3 ).value = FOT
		.Cell( 6, 5 ).value = Round2( FOT * percent / 100 , 2 )

		Repcalc.Clear

		CodeID = WorkArea.GetCodeID( "2010" )

		With WorkArea.Code( CodeID )
			percent = .Percent
			Sheet( 1 ).Cell( 7, 7 ).Value = .DbCode
			Sheet( 1 ).Cell( 7, 8 ).Value = .CrCode
		End With

		Sheet1.Cell(7, 3).Value = 0
		Sheet1.Cell(8, 3).Value = 0
		Sheet1.Cell(9, 3).Value = 0

		If Workarea.Period.Year < 2009 Then
			FOT = CalcDepsSumLimit( CodeID, Repcalc, d, max )
			sum1 = sum1 + Round2( FOT * percent / 100 , 2 )
			.Cell( 7, 3 ).value = FOT
			.Cell( 7, 5 ).value = Round2( FOT * percent / 100 , 2 )
		Else
			Calc2010 CodeID, Repcalc, d, max 
		End If

		Repcalc.Clear
		CodeID = WorkArea.GetCodeID( "2020" )
		With WorkArea.Code( CodeID )
			percent = .Percent
			Sheet( 1 ).Cell( 10, 7 ).Value = .DbCode
			Sheet( 1 ).Cell( 10, 8 ).Value = .CrCode
		End With
		.Cell( 10, 2 ).Value = percent
		FOT = CalcDepsSumLimit( CodeID, Repcalc, d, max )
		sum1 = sum1 + Round2( FOT * percent / 100 , 2 )
		.Cell( 10, 3 ).value = FOT
		.Cell( 10, 5 ).value = Round2( FOT * percent / 100 , 2 )

		Repcalc.Clear
		CodeID = WorkArea.GetCodeID( "2030" )
		With WorkArea.Code( CodeID )
			percent = .Percent
			Sheet( 1 ).Cell( 11, 7 ).Value = .DbCode
			Sheet( 1 ).Cell( 11, 8 ).Value = .CrCode
		End With
		.Cell( 11, 2 ).Value = percent
		Calc2030 CodeID, Repcalc, d, max

'		.Cell( 4, 5 ).value = sum1
	End With
End Sub

'-----------------------------------------------------------------------------------------------
' Function GetCodeID( CodeName ) - тх№эѓђќ шфхэђшєшърђю№ яю шьхэш ъюфр/ьхђюфр 
'-----------------------------------------------------------------------------------------------
Function GetCodeID( CodeName )
	Dim SqlQr, ar
	GetCodeID = 0
	SqlQr = "Select MTD_ID From MTD_CODES Where MTD_NAME = '" & CodeName & "'"
	Select Case App.AppType
	Case "DAO"
		If GetRows_DAO( SqlQr, ar ) Then GetCodeID = ar( 0, 0 )
	Case "OLEDB"
		If GetRows_ADO( SqlQr, ar ) Then GetCodeID = ar( 0, 0 )
	End Select
End Function

'________________________________________________
' Function QueryWithParams( SqlText, Options, arParams )
' Тћяюыэхэшх чря№юёр ё ях№хфрїхщ т эхую яр№рьхђ№ют
' Тючт№рљрхђёџ юсњхъђ ђшяр RecordSet
' ADO
'________________________________________________
Function QueryWithParams( SqlText, Options, arParams )
	Dim Cn, Cmd, Prm, i, val1

	Set Cn = WorkArea.AdoConnection
	Cn.CursorLocation = adUseServer
	Set Cmd = CreateObject( "ADODB.Command" )
	Set Cmd.ActiveConnection = Cn
	Cmd.CommandText = SqlText
'	Cmd.CommandTimeout = 600

	' ѓёђрэртыштрхь яр№рьхђ№ћ чря№юёр
	For i = 0 To UBound( arParams ) Step 2
		val1 = arParams( i + 1 )
		If IsNull( val1 ) Then
			Set Prm = Cmd.CreateParameter( , arParams( i ), adParamInput, 10 )
		ElseIf arParams( i ) = adVarChar Then
			Set Prm = Cmd.CreateParameter( , arParams( i ), adParamInput, Iif( Len( val1 ) = 0, 1, Len( val1 )), val1 )
		Else
			Set Prm = Cmd.CreateParameter( , arParams( i ), adParamInput )
			Prm.Value = val1
		End If
		Cmd.Parameters.Append Prm
	Next
	Set QueryWithParams = Cmd.Execute( ,, Options )
End Function

'________________________________________________
Function GetRows_DAO( SQLText, ByRef ar )
	Dim rs, n
	ar = Array()
	GetRows_DAO = False
	Set rs = WorkArea.DAODataBase.OpenRecordset( SqlText )
	If rs.eof Then Exit Function
	rs.MoveLast
	n = rs.RecordCount
	rs.MoveFirst
	ar = rs.GetRows( n )
	GetRows_DAO = True
End Function

'________________________________________________
Function GetRows_ADO( SQLText, Options, ByRef ar )
	ar = Array()
	GetRows_ADO = False
	With WorkArea.AdoConnection.Execute( SqlText, , Options )
		If .eof Then Exit Function
		ar = .GetRows
	End With
	GetRows_ADO = True
End Function

'________________________________________________
Function CalcDepsSumLimit( CodeID, Repcalc, d, max )
	Dim i
	Next

	For j = StartNumberOfColumns + 1 To Sheet( 1 ).Columns
		Sheet( 1 ).Cell(StartNumberOfRows, j).Value = Totals(j)
	Next
End Sub

'-----------------------------------------------------------------------------------------------
Sub CalcTotalsInRows
	Dim i, j, k
	Dim Sum
	With Sheet( 1 )
		For i = StartNumberOfRows + 1 To .Rows
			For k = 1 To UBound( TotalsCols )
				Sum = 0
				For j = TotalsCols( k - 1 ) + 1 To TotalsCols( k ) - 1
					Sum = Sum + .Cell( i, j ).Value
				Next
				.Cell( i, TotalsCols( k )).Value = Sum
			Next
			.Cell( i, .Columns ).Value = .Cell( i, TotalsCols( 1 ) ).Value - .Cell( i, TotalsCols( 2 )).Value
		Next
	End With
End Sub

'-----------------------------------------------------------------------------------------------
Sub SetZero(Arr)
	Dim i
	For i = 1 To UBound(Arr)
		Arr(i) = 0
	Next
End Sub

'-----------------------------------------------------------------------------------------------
Sub MergeCells(OnRow, CellsText)
	Dim j
	Dim FirstPos, LastPos

	FirstPos = 1
	LastPos = 0

	For j = 1 To Sheet( 1 ).Columns
		If Sheet( 1 ).Cell(OnRow, j).Value = CellsText Then
			FirstPos = j
			Exit For
		End If
	Next

	For j = Sheet( 1 ).Columns To 1 Step -1
		If Sheet( 1 ).Cell(OnRow, j).Value = CellsText Then
			LastPos = j 
			Exit For
		End If
	Next

	If LastPos >= FirstPos Then Sheet( 1 ).Range(FirstPos, OnRow, LastPos, OnRow).MergeCells = True
End Sub

'________________________________________________
Sub BuildSecondaryReport
	Dim max, percent, i, CodeID, FOT, d, sum1
	Dim arFonds, Repcalc
	arFonds = Array( "2000", "2001", "2010", "2020", "2030" )
	Set Repcalc = WorkArea.CreateIntObject("CommonCalc")
	Repcalc.setPeriod "month", Workarea.Period.Month, Workarea.Period.Year
	d = Workarea.Period.FirstDay

	max = Workarea.UTable( USRTBL_LGT ).GetValue( 1, USRTBL_MAXCY, 2 )
	With Sheet( 1 )
		CodeID = WorkArea.GetCodeID( "2000" )
		With WorkArea.Code( CodeID )
			percent = .Percent
			Sheet( 1 ).Cell( 5, 7 ).Value = .DbCode
			Sheet( 1 ).Cell( 5, 8 ).Value = .CrCode
		End With
		.Cell( 5, 2 ).Value = percent
		FOT = CalcDepsSumLimit_Pens( CodeID, Repcalc, max, False )
		sum1 = Round2( FOT * percent / 100 , 2 )
		.Cell( 5, 3 ).value = FOT
		.Cell( 5, 5 ).value = Round2( FOT * percent / 100 , 2 )

		Repcalc.Clear
		CodeID = WorkArea.GetCodeID( "2001" )
		With WorkArea.Code( CodeID )
			percent = .Percent
			Sheet( 1 ).Cell( 6, 7 ).Value = .DbCode
			Sheet( 1 ).Cell( 6, 8 ).Value = .CrCode
		End With
		.Cell( 6, 2 ).Value = percent
		FOT = CalcDepsSumLimit_Pens( CodeID, Repcalc, max, True )
		sum1 = sum1 + Round2( FOT * percent / 100 , 2 )
		.Cell( 6, 3 ).value = FOT
		.Cell( 6, 5 ).value = Round2( FOT * percent / 100 , 2 )

		Repcalc.Clear

		CodeID = WorkArea.GetCodeID( "2010" )

		With WorkArea.Code( CodeID )
			percent = .Percent
			Sheet( 1 ).Cell( 7, 7 ).Value = .DbCode
			Sheet( 1 ).Cell( 7, 8 ).Value = .CrCode
		End With

		Sheet1.Cell(7, 3).Value = 0
		Sheet1.Cell(8, 3).Value = 0
		Sheet1.Cell(9, 3).Value = 0
		Sheet1.Cell(7, 5).Value = 0


		If Workarea.Period.Year < 2009 Then
			FOT = CalcDepsSumLimit( CodeID, Repcalc, d, max )
			sum1 = sum1 + Round2( FOT * percent / 100 , 2 )
			.Cell( 7, 3 ).value = FOT
			.Cell( 7, 5 ).value = Round2( FOT * percent / 100 , 2 )
		Else
			Calc2010 CodeID, Repcalc, d, max 
		End If

		Repcalc.Clear
		CodeID = WorkArea.GetCodeID( "2020" )
		With WorkArea.Code( CodeID )
			percent = .Percent
			Sheet( 1 ).Cell( 10, 7 ).Value = .DbCode
			Sheet( 1 ).Cell( 10, 8 ).Value = .CrCode
		End With
		.Cell( 10, 2 ).Value = percent
		FOT = CalcDepsSumLimit( CodeID, Repcalc, d, max )
		sum1 = sum1 + Round2( FOT * percent / 100 , 2 )
		.Cell( 10, 3 ).value = FOT
		.Cell( 10, 5 ).value = Round2( FOT * percent / 100 , 2 )

		Repcalc.Clear
		CodeID = WorkArea.GetCodeID( "2030" )
		With WorkArea.Code( CodeID )
			percent = .Percent
			Sheet( 1 ).Cell( 11, 7 ).Value = .DbCode
			Sheet( 1 ).Cell( 11, 8 ).Value = .CrCode
		End With
		.Cell( 11, 2 ).Value = percent
		Calc2030 CodeID, Repcalc, d, max

'		.Cell( 4, 5 ).value = sum1
	End With
End Sub

'-----------------------------------------------------------------------------------------------
' Function GetCodeID( CodeName ) - тх№эѓђќ шфхэђшєшърђю№ яю шьхэш ъюфр/ьхђюфр 
'-----------------------------------------------------------------------------------------------
Function GetCodeID( CodeName )
	Dim SqlQr, ar
	GetCodeID = 0
	SqlQr = "Select MTD_ID From MTD_CODES Where MTD_NAME = '" & CodeName & "'"
	Select Case App.AppType
	Case "DAO"
		If GetRows_DAO( SqlQr, ar ) Then GetCodeID = ar( 0, 0 )
	Case "OLEDB"
		If GetRows_ADO( SqlQr, ar ) Then GetCodeID = ar( 0, 0 )
	End Select
End Function

'________________________________________________
' Function QueryWithParams( SqlText, Options, arParams )
' Тћяюыэхэшх чря№юёр ё ях№хфрїхщ т эхую яр№рьхђ№ют
' Тючт№рљрхђёџ юсњхъђ ђшяр RecordSet
' ADO
'________________________________________________
Function QueryWithParams( SqlText, Options, arParams )
	Dim Cn, Cmd, Prm, i, val1

	Set Cn = WorkArea.AdoConnection
	Cn.CursorLocation = adUseServer
	Set Cmd = CreateObject( "ADODB.Command" )
	Set Cmd.ActiveConnection = Cn
	Cmd.CommandText = SqlText
'	Cmd.CommandTimeout = 600

	' ѓёђрэртыштрхь яр№рьхђ№ћ чря№юёр
	For i = 0 To UBound( arParams ) Step 2
		val1 = arParams( i + 1 )
		If IsNull( val1 ) Then
			Set Prm = Cmd.CreateParameter( , arParams( i ), adParamInput, 10 )
		ElseIf arParams( i ) = adVarChar Then
			Set Prm = Cmd.CreateParameter( , arParams( i ), adParamInput, Iif( Len( val1 ) = 0, 1, Len( val1 )), val1 )
		Else
			Set Prm = Cmd.CreateParameter( , arParams( i ), adParamInput )
			Prm.Value = val1
		End If
		Cmd.Parameters.Append Prm
	Next
	Set QueryWithParams = Cmd.Execute( ,, Options )
End Function

'________________________________________________
Function GetRows_DAO( SQLText, ByRef ar )
	Dim rs, n
	ar = Array()
	GetRows_DAO = False
	Set rs = WorkArea.DAODataBase.OpenRecordset( SqlText )
	If rs.eof Then Exit Function
	rs.MoveLast
	n = rs.RecordCount
	rs.MoveFirst
	ar = rs.GetRows( n )
	GetRows_DAO = True
End Function

'________________________________________________
Function GetRows_ADO( SQLText, Options, ByRef ar )
	ar = Array()
	GetRows_ADO = False
	With WorkArea.AdoConnection.Execute( SqlText, , Options )
		If .eof Then Exit Function
		ar = .GetRows
	End With
	GetRows_ADO = True
End Function

'________________________________________________
Function CalcDepsSumLimit( 