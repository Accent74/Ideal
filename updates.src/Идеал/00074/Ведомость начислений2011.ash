аЯрЁБс                >  ўџ	                               ўџџџ       A   џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџt rs = WorkArea.DAODataBase.OpenRecordset( SqlText )
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
	Dim TotalSum
	Dim AgSum
	Dim flAllowDogovor

	With Workarea.Code(CodeID)
		flAllowDogovor = Iif( ((.Mode And 8) = 0) Or (.Type <> 103), False, True )
	End With

	TotalSum = 0
	For i = 1 To Repcalc.Count
		If flAllowDogovor Or (Not Repcalc.Agent(i).Prop("в№ѓфютющ фюуютю№", d)) Then
			AgSum = Repcalc.SumBy("CodeSumDependsC", i, CodeID)
			If AgSum > max Then AgSum = max
			TotalSum = TotalSum + AgSum
		End If
	Next
	CalcDepsSumLimit = TotalSum
End Function
'________________________________________________
Function CalcDepsSumLimitAll( CodeID, Repcalc, d, max )
	Dim i
	Dim TotalSum
	Dim AgSum

	TotalSum = 0
	For i = 1 To Repcalc.Count
		AgSum = Repcalc.SumBy("CodeSumDependsC", i, CodeID)
		If AgSum > max Then AgSum = max
		TotalSum = TotalSum + AgSum
	Next
	CalcDepsSumLimitAll = TotalSum
End Function
'________________________________________________
Function CalcDepsSumLimit_Pens( Code, Repcalc, max, FirstDay, ByRef FOT_GPD)
	Dim i, CodeMode, IsAgInvalid 
	Dim TotalSum
	Dim AgSum, Ag
	Dim flAllowDogovor

	Fot_GPD = 0

	flAllowDogovor = ((Code.mode And 8) = 8) 'тъыўїрђќ УЯФ
	CodeMode = (Code.mode And 7)
	IsAgInvalid = (CodeMode = 2)

	TotalSum = 0

	For i = 1 To Repcalc.Count
		Set Ag = Repcalc.Agent(i)

		If (Ag.Prop("Шэтрышф") = IsAgInvalid Or CodeMode = 0) And _
			Ag.Prop("в№ѓфютющ фюуютю№", FirstDay) = flAllowDogovor  Then

			AgSum = Repcalc.SumBy("CodeSumDependsC", i, Code.ID)
			If AgSum > max Then AgSum = max
			TotalSum = TotalSum + AgSum

			If Ag.Prop("в№ѓфютющ фюуютю№", FirstDay) Then 
				Fot_GPD = Fot_GPD + AgSum
			End If
		End If
	Next

	CalcDepsSumLimit_Pens = TotalSum
End Function
<     ш                Ь Arial D:\STDs.!!!\Zap\Шфхры\Шфхры       џџџџ                 џџ 
 CSheetPageSheet1Тхфюьюёђќ эрїшёыхэшщ         61    7                                     
  .      |   Ё   [   %   :   J   1   b   :   @   U   g      4   4   ?                                                           џџ  CRow џџ  CCell6     џџ   (   54><>ABL  =0G8A;5=89                      " џџџџ                        " џџџџ                        " џџџџ    >O  D8@<0                    " џџџџ     >O  D8@<0                    " џџџџ     >O  D8@<0                    " џџџџ     >O  D8@<0                        џџ      /=20@L  2 0 1 2   3.                       " џџџџ                        " џџџџ                         ! џџ       0G8A;5=85                    ! џџ       %                     ! џџ;     $"                    ! џџ<                        ! џџ;  
   !C<<0                    ! џџ<                        ! џџ    
   515B                    ! џџ       @548B                               B>3>:                                                =C5+C6+C7+C8"   
                         "                         =E5+E6+E7+E8"   
                         "                          "                           "                              џџ       !  >1I89                      џџ*    xЏ                         " џџ                       " џџ/                        " џџ                       " џџ/                        " џџ                            " џџ       6 5 7                        џџ)       !  8=20;84K                        џџ    H                         " џџ                         џџ   р                         " џџ                         џџ                        " џџ                            " џџ       6 5 7                        џџ       !                        џџ    xK                         " џџ                         џџ                        " џџ                       " џџ                        " џџ                            " џџ       6 5 7                        џџ       !  1>;L=8G=K5                      џџ    р                         " џџ                       " џџ                        " џџ                       " џџ                        " џџ                            " џџ       6 5 7                          " џџџџ                        " џџџџ                        " џџџџ                        " џџџџ                         ! џџ       #45@60=85                    ! џџ       %                     ! џџ;  
   !C<<0                    " џџ<                        ! џџ    
   515B                    ! џџ       @548B                       џџ    
   $                        џџ                         " џџ                       " џџ                        " џџ       6 6 1                     " џџ       6 4 1 4                        џџ       !  ( 107>2K9)                       џџ    ЭЬЬЬЬЬ@                    " џџ*                       " џџ*                        " џџ       6 6 1                     " џџ    
   6 5 7 / 1                        џџ        !  ( 1>;L=8G=K5)                       џџ                         " џџџџ                      " џџџџ                       " џџ       6 6 1                     " џџ    
   6 5 7 / 3                        џџ       !  ( )                       џџ    ЭЬЬЬЬЬ@                    " џџ                       " џџ                        " џџ       6 6 1                     " џџ    
   6 5 7 / 2                        % џџ:       >4@0745;5=85                    % џџ:    
   !C<<0                     % џџ9       !>B@C4=8:                    ! џџ9       ;L3>BK                     "                           "                                       џџџџ   :   CE30;B5@: _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _                       џџџџ                        џџџџ                         џџџџ                         џџџџ                         џџџџ                                      <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                              <     ш                Ь Arial                                2   2      
 <     ш            РРР Ь Arial ђюф F1 ьхэђ ёяшёър 2 _ 43.ad <     ш    Ш           Ь Arial аDb                   
    <     ш   №           Ь Arial Sub ShtBook_OnLoadEnd Su  <     ш    №           Ь Arial шђх F1 ыхээюую §ыхьхэђр 
.av <     ш               Ь Arial шђх F1  (F7)    1      шЭe <     ш           џџр Ь Arial џџ               №ѓђъш    <     ш           ((( Ь Arial џџ               htRange) <     ш           џџџ Ь Arial џџ                   ing  <     ш            ((( Ь Arial џџ                   а#g <     ш   Ш           Ь Arial                      шЭe	 <     ш    Д           Ь Arial ющёђт ћ (F7) ёюёђюџэшџ фыџ ю=    џџџ                               ШШШ                               ШШШ                                  ммм                                  ммм                              џџр                                  џџџџ                             џџр                              џџџџ                              џџџџ                              	 џџџџ                             
 рџџ                               џџџ                                рџџ                                рџџ                               рџџ                              џџџ                                џџџ                                 рџџ                                 џџџ                                џџџ                               џџџџ                               џџџџ                                џџџџ                               РРР                                  ШШШ                               РРР                               ШШШ                               РРР                              џџџ                                  џџџ                              РРР                              . џџџџ                              - ццњ                              , ццњ                              + ццњ                             ' џџр                              & ррр                               # ррр                             " ррр                                 ! џџр                                џџр                              $ ррр                              % ррр                              ( рџџ                                 ) џџџџ                             * џџџџ                               / џџџџ                           РРР 0 ццњ                                 1 жжж                              2 жжж                              3 жжж                             4 еее                              5 еее                              6 ддд                              7 ддд                              8 ввв                              9 ммм                              : ммм                              ; ммм    ddd    ddd            ddd < ммм            ddd    ddd    ddd    ; ммм    ddd    ddd            ddd < ммм            ddd    ddd    ddd                  ; ммм    ddd    ddd            ddd < мR o o t   E n t r y                                               џџџџџџџџ                               `Ю         C o n t e n t s                                                  џџџџ   џџџџ                                       yi       S u m m a r y I n f o r m a t i o n                           (  џџџџџџџџџџџџ                                        i                                                                          џџџџџџџџџџџџ                                                R o o t   E n t r y                                               џџџџџџџџ                               аXdЮ         C o n t e n t s                                                  џџџџ   џџџџ                                       Шi       S u m m a r y I n f o r m a t i o n                           (  џџџџџџџџџџџџ                                        i                                                                          џџџџџџџџџџџџ                                                                        	   
                                 ўџџџџџџџўџџџ§џџџўџџџўџџџ   ?   џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ@   B   §џџџC   D   E   F   G   H   I   J   K   L   M   N   O   P   Q   R   S   T   U   V   W   X   Y   Z   [   \   ]       џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ               ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџўџ   рђљOhЋ +'Гй   рђљOhЋ +'Гй0   9        x            Є      А      М      Ш      д   	   р   
   ь      ј               $      ,     у        Тхфюьюёђќ эрїшёыхэшщ2011                                                            614 @    ж  @    
[Ю@   @kчєЖ§Ц      Ръіхэђ 7.0                                                                                                                                                                           џAIOption Explicit
'#include "definition.avb"
'#include "ado.inc"
Const adInteger          =   3 ' vbLong
Const adVarChar          = 200
Const adParamInput = 1
Const adUseServer        = 2

Const StartNumberOfRows = 18
Const LastRow = 11
Const StartNumberOfColumns = 2

Const MTC_PRIVILEGE_NAME = "0801"
Const MTC_WPF_NAME = "дЮв"

Dim cRow

Dim DicCodes
Dim TotalsCols( 2 )
Dim arCodes
Dim PrivilegeID
Dim BuhName

Const MTD_NDFL_NAME 		= "2400"
Const MTD_ESV_NAME 			= "2410"
Const MTD_ESVB_NAME			= "2420"
Const MTD_ESVT_NAME 		= "2430"

Dim MTD_NDFL_ID 	
Dim MTD_ESV_ID 	
Dim MTD_ESVB_ID	
Dim MTD_ESVT_ID 	
'-----------------------------------------------------------------------------------------------
Sub ShtBook_OnLoad
	Dim ut

	Set ut = Workarea.UTable(USRTBL_ONE)

	BuhName =ut.GetValue(1, "дШЮ сѓѕурыђх№р", 2)
	PrivilegeID = WorkArea.GetCodeID(MTC_PRIVILEGE_NAME)

	Set DicCodes = CreateObject("Scripting.Dictionary")
	arCodes = GetCodes

	With Sheet( 1 )
		.DisplayGr                        	   
                                 ўџџџўџџџџџџџџџџџџџџџџџџџџџџџџџџџ§џџџўџџџўџџџ    !   "   #   $   %   &   '   (   )   *   +   ,   -   .   /   0   1   2   3   4   5   6   7   8   9   :   ;   <   =   >       џџџџџџџџ§џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ               ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџўџ   рђљOhЋ +'Гй   рђљOhЋ +'Гй0   9        x            Є      А      М      Ш      д   	   р   
   ь      ј               $      ,     у        Тхфюьюёђќ эрїшёыхэшщ2011                                                            615 @   ДЗ  @   pчuЮ@   @kчєЖ§Ц      Ръіхэђ 7.0                                                                                                                                                                           џђHOption Explicit
'#include "definition.avb"
'#include "ado.inc"
Const adInteger          =   3 ' vbLong
Const adVarChar          = 200
Const adParamInput = 1
Const adUseServer        = 2

Const StartNumberOfRows = 18
Const LastRow = 11
Const StartNumberOfColumns = 2

Const MTC_PRIVILEGE_NAME = "0801"
Const MTC_WPF_NAME = "дЮв"

Dim cRow

Dim DicCodes
Dim TotalsCols( 2 )
Dim arCodes
Dim PrivilegeID
Dim BuhName

Const MTD_NDFL_NAME 		= "2400"
Const MTD_ESV_NAME 			= "2410"
Const MTD_ESVB_NAME			= "2420"
Const MTD_ESVT_NAME 		= "2430"

Dim MTD_NDFL_ID 	
Dim MTD_ESV_ID 	
Dim MTD_ESVB_ID	
Dim MTD_ESVT_ID 	
'-----------------------------------------------------------------------------------------------
Sub ShtBook_OnLoad
	Dim ut

	Set ut = Workarea.UTable(USRTBL_ONE)

	BuhName =ut.GetValue(1, "дШЮ сѓѕурыђх№р", 2)
	PrivilegeID = WorkArea.GetCodeID(MTC_PRIVILEGE_NAME)

	Set DicCodes = CreateObject("Scripting.Dictionary")
	arCodes = GetCodes

	With Sheet( 1 )
		.DisplayGrid = False
		.DisplayHeadings = True
	End With

	EnableButtons 1
	PeriodText = WorkArea.Period.RepTitle

	MTD_NDFL_ID 		= Workarea.GetCodeID(MTD_NDFL_NAME)
	MTD_ESV_ID 			= Workarea.GetCodeID(MTD_ESV_NAME)
	MTD_ESVB_ID 		= Workarea.GetCodeID(MTD_ESVB_NAME)
	MTD_ESVT_ID 		= Workarea.GetCodeID(MTD_ESVT_NAME)

	Sheet( 1 ).Cell( 2, 1 ).Value = WorkArea.Period.RepTitle
	Sheet( 1 ).Range( 5, 1, 8, 1 ).Mergecells = True
	Sheet( 1 ).Range( 5, 1, 8, 1 ).Value = WorkArea.Agent(1237).Name
	Sheet( 1 ).Range( 5, 1, 8, 1 ).Alignment = acRight
	BuildReport

	Sheet1.DisplayHeadings False
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

'________________________________________________
' Function GetCodes
' тћсш№рхь ъюфћ, чр шёъыўїхэшхь ёяхі-ъюфют "тћяырђр", "фюыу ёюђ№ѓфэшър" ш "фюыу я№хфя№шџђшџ"
'________________________________________________
Function GetCodes
	Dim SqlQr, ar
	Select Case App.AppType
	Case "DAO"
		SqlQr = "SELECT MTD_ID, MTD_name, 0 FROM MTD_CODES WHERE ( MTD_TYPE = 101 ) AND ( IIF( MTD_MODE IS NULL, 0, MTD_MODE ) <> 2 ) ORDER BY 3, 2" & _
				vbNewLine & "UNION" & vbNewLine & _
				"SELECT MTD_ID, MTD_name, 1 FROM MTD_CODES WHERE ( MTD_TYPE = 102 ) AND ( IIF( MTD_MODE IS NULL, 0, MTD_MODE ) <> 3 ) AND ( IIF( MTD_MODE IS NULL, 0, MTD_MODE ) <> 1 ) ORDER BY 3, 2"
		GetRows_DAO SqlQr, ar
		GetCodes = ar
	Case "OLEDB"
		SqlQr = "SELECT MTD_ID, MTD_name, 0 FROM MTD_CODES WHERE ( MTD_TYPE = 101 ) AND ( ISNULL( MTD_MODE, 0 ) <> 2 ) " & _
				" UNION all " & _
				"SELECT MTD_ID, MTD_name, 1 FROM MTD_CODES WHERE ( MTD_TYPE = 102 ) AND ( ISNULL( MTD_MODE, 0 ) <> 3 ) AND ( ISNULL( MTD_MODE, 0 ) <> 1 ) ORDER BY 3, 2"
		GetRows_ADO SQLQr, adCmdText, ar
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
		.Columns = 8
		.Rows = StartNumberOfRows
		' юїшёђшыш ёђ№юъѓ шђюуют
		For i = 2 To .Columns
			.Cell( StartNumberOfRows, i ).Value = 0.0
		Next
	End With

	PeriodID = WorkArea.Period.ID
	arData = GetData( PeriodID )

	If UBound(arData) = -1 Then Exit Sub

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
				.Cell( StartNumberOfRows, DicCodes( arData( 4, i ))).Value = 0. 'arData( 6, i )
				GroupID = arData( 0, i )
			End If

			If AgID <> arData( 1, i ) Then
				cRow = .InsertRow2

				.Cell( cRow, 1 ).Value = arData( 3, i )
				.Cell( cRow, 2 ).Value = GetPrivilege( PrivilegeID, arData( 1, i ), PeriodID )
				AgID = arData( 1, i )
			End If
			.Cell( cRow, DicCodes( arData( 4, i ))).Value = arData( 5, i )

'			.Cell( StartNumberOfRows, DicCodes( arData( 4, i )) ).Value = .Cell( StartNumberOfRows, DicCodes( arData( 4, i ))).Value + arData( 6, i )

		Next
		CalcTotals
		.Range( 1, StartNumberOfRows - 2, .Columns, .Rows).Autofit
		.Range( 1, StartNumberOfRows + 1, .Columns, .Rows).Stripe
		.Range( 1, StartNumberOfRows + 1, .Columns, .Rows).Alignment = acRight
		.Range( 1, StartNumberOfRows + 1, .Columns, .Rows).SetBorder acBrdOutline + acBrdInsideVert, acLineNormal

		With .Range( 3, StartNumberOfRows, .Columns, StartNumberOfRows )
			.Font.Bold = True
			.Alignment = acRight
			.BackColor = 14745599
			.SetBorder acBrdGrid
		End With

		cRow = .InsertRow2
		cRow = .InsertRow2
		.Cell(cRow,1).Value = "Сѓѕурыђх№:___________________"
  		.Range (2,cRow,6,cRow).MergeCells = True
		If BuhName <>"" Then
			.Cell(cRow,2).Value ="/ "&BuhName&" /"
		End If 
	End With
End Sub

'-----------------------------------------------------------------------------------------------
Function GetData(PeriodID)
	Dim SqlQr, Rs, n

	GetData = Array()

	If App.AppType = "DAO" Then
		With WorkArea.DAODataBase.CreateQueryDef("")
			.Sql = "PARAMETERS PeriodID Long; " & _
			       "Select AG_TREE.P0, P_JOURNAL.AG_ID, Groups.AG_NAME, AGENTS.AG_NAME, MTC_ID, Sum(JP_SUM),  " & _
						"sum(iif(JP_SUM <>0, JP_Tariff,0)) " & _
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

			Set Rs = .OpenRecordSet
			If Not rs.eof Then
				Rs.MoveLast
				n = Rs.RecordCount
				Rs.MoveFirst
				GetData = Rs.GetRows( n )
			End If
		End With
	Else
		SqlQr = "Select AG_TREE.P0, P_JOURNAL.AG_ID, Groups.AG_NAME, AGENTS.AG_NAME, MTC_ID, Sum(JP_SUM), sum(Case WHEN JP_SUM <> 0 Then JP_Tariff Else 0 End)  " & _
		        "From P_JOURNAL Inner Join AGENTS On P_JOURNAL.AG_ID = AGENTS.AG_ID " & _
		        "               Inner Join AG_TREE On P_JOURNAL.AG_ID = ID " & _
		        "               Inner Join MTD_CODES On MTD_CODES.MTD_ID = P_JOURNAL.MTC_ID  " & _
		        "               Left Join AGENTS Groups On P0 = Groups.AG_ID " & _
		        "Where JP_DONE = 1 And SHORTCUT = 0 AND ( ISNULL( MTD_MODE, 0 ) <> 1 ) AND ( ISNULL( MTD_MODE, 0 ) <> 2 ) AND ( ISNULL( MTD_MODE, 0 ) <> 3 ) " & _
		        "                  And ( IsNull( MTD_MASK, 0 ) < 8 ) " & _
		        "                  And W_PERIOD = " & CStr(PeriodID)  & _
		        "Group By AG_TREE.P0, P_JOURNAL.AG_ID, Groups.AG_NAME, AGENTS.AG_NAME, MTC_ID " & _
		        "Having Sum(JP_SUM) <> 0 " & _
		        "Order By Groups.AG_NAME, AGENTS.AG_NAME "

		Set Rs = QueryWithParams( SqlQr, adCmdText, Array())

		If Not Rs.EOF Then
			GetData = Rs.GetRows
		End If
	End If

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

	SetZero Totals		' шђюуш 3
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

	For j = StartNumberOfColumns + 1 To Sheet(1).Columns	' шђюуш 3
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

Sub display_code_calc(Code, Row, FOT_Sum, FOT_Percent)
	Sheet( 1 ).Cell( Row, 7 ).Value = Code.DbCode
	Sheet( 1 ).Cell( Row, 8 ).Value = Code.CrCode
	Sheet( 1 ).Cell( Row, 2 ).Value = Code.percent
	Sheet( 1 ).Cell( Row, 3 ).value = FOT_Sum
	Sheet( 1 ).Cell( Row, 5 ).value = FOT_Percent
End Sub
'________________________________________________
Sub BuildSecondaryReport
	Dim max, percent, i, CodeID, FOT, FOT_GPD
	Dim arFonds, Repcalc
	Dim Code, FirstDay

	arFonds = Array( "2000", "2001", "2003", "2002")

	Set Repcalc = WorkArea.CreateIntObject("CommonCalc")
	Repcalc.setPeriod "month", Workarea.Period.Month, Workarea.Period.Year
	FirstDay = Workarea.Period.FirstDay

	max = Workarea.UTable( USRTBL_LGT ).GetValue( 1, USRTBL_MAXCY, 2 )
	percent = Workarea.UTable( USRTBL_LGT ).GetValue( 1, USRTBL_ECB, 2 )

	For i = 0 To UBound(arFonds)
		CodeID = WorkArea.GetCodeID( arFonds(i))

		If CodeID <> 0 Then
			Set Code = WorkArea.Code(CodeID)
			percent = Code.Percent

			FOT = CalcDepsSumLimit_Pens( Code, Repcalc, max, FirstDay, FOT_GPD)
			display_code_calc Code, 5 + i, FOT, Round2( FOT * percent / 100 , 2 )
			Repcalc.Clear
		End If

	Next

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
				Sheet( 1 ).Cell( cRow, 5 ).Value = .DbCode
				Sheet( 1 ).Cell( cRow, 6 ).Value = .CrCode
			End With
		Next
	End With
End Sub
'=================
Sub GetSumByMtdT(mtdID, sum1, sumT)
	Dim rs
	Set rs = App.WorkArea.DAODataBase.OpenRecordset( _
		"SELECT SUM(JP_SUM) AS PSUM, sum(jp_tariff) as pt FROM P_JOURNAL WHERE JP_DONE = 1 " & _
			" AND C_PERIOD = " & WorkArea.Period.ID & _
			" AND MTD_ID = " & mtdID )
	If Not rs.eof	Then
		sum1 = rs.fields("PSUM")
		sumT = rs.fields("pt")
	End If
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
	Seid = False
		.DisplayHeadings = True
	End With

	EnableButtons 1
	PeriodText = WorkArea.Period.RepTitle

	MTD_NDFL_ID 		= Workarea.GetCodeID(MTD_NDFL_NAME)
	MTD_ESV_ID 			= Workarea.GetCodeID(MTD_ESV_NAME)
	MTD_ESVB_ID 		= Workarea.GetCodeID(MTD_ESVB_NAME)
	MTD_ESVT_ID 		= Workarea.GetCodeID(MTD_ESVT_NAME)

	Sheet( 1 ).Cell( 2, 1 ).Value = WorkArea.Period.RepTitle
	Sheet( 1 ).Range( 5, 1, 8, 1 ).Mergecells = True
	Sheet( 1 ).Range( 5, 1, 8, 1 ).Value = WorkArea.Agent(1237).Name
	Sheet( 1 ).Range( 5, 1, 8, 1 ).Alignment = acRight
	BuildReport

	Sheet1.DisplayHeadings False
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

'---------------------------џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ