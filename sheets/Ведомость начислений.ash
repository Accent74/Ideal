ĐĎŕĄąá                >  ţ˙	               ;          =      ţ˙˙˙    <   [     ˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙
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
		With Repcalc.Agent(i)
			If flAllowDogovor Or (Not .Prop("Ňđóäîâîé äîăîâîđ", d)) Then
				AgSum = CalcSumByPeriod(.ID, CodeID, d, workarea.period.lastday)'Repcalc.SumBy("CodeSumDependsW", i, CodeID)
				If AgSum > max Then AgSum = max
				TotalSum = TotalSum + AgSum
			End If
		End With
	Next
	CalcDepsSumLimit = TotalSum
End Function

'________________________________________________
Function CalcDepsSumLimit_Pens( CodeID, Repcalc, max, IsAgInvalid, d )
	Dim i
	Dim TotalSum
	Dim AgSum

	TotalSum = 0

	For i = 1 To Repcalc.Count
		With Repcalc.Agent(i)
			If .Prop("Číâŕëčä") = IsAgInvalid Then
				AgSum = CalcSumByPeriod(.ID, CodeID, d, workarea.period.lastday) 'Repcalc.SumBy("CodeSumDependsW", i, CodeID)
				If AgSum > max Then AgSum = max
				TotalSum = TotalSum + AgSum
			End If
		End With
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

				If Month(d) = 1 And Year(d) = 2009 Then
					AgSum1 = CalcSumByPeriod(.ID, CodeID, CDate("01/01/2009"), CDate("12/01/2009"))
					AgSum2 = CalcSumByPeriod(.ID, CodeID, CDate("13/01/2009"), CDate("31/01/2009"))
				Else
					AgSum1 = 0
					AgSum2 = CalcSumByPeriod(.ID, CodeID, d, workarea.period.lastday)
				End If

				If (AgSum1 + AgSum2) > max Then AgSum2 = max - AgSum1

				If .Prop(PROP_TRUD) Then
					If .Prop("Ńîńňîčň â ňđóäîâűő îňíîřĺíč˙ő") Then
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

				If Month(d) = 1 And Year(d) = 2009 Then
					AgSum1 = CalcSumByPeriod(.ID, CodeID, CDate("01/01/2009"), CDate("12/01/2009"))
					AgSum2 = CalcSumByPeriod(.ID, CodeID, CDate("13/01/2009"), CDate("31/01/2009"))
				Else
					AgSum1 = 0
					AgSum2 = CalcSumByPeriod(.ID, CodeID, CDate("13/01/2009"), CDate("31/01/2009"))
				End If

				If (AgSum1 + AgSum2) > max Then AgSum2 = max - AgSum1

				Sheet1.Cell(11, 3).Value = Sheet1.Cell(11, 3).Value + AgSum1 + AgSum2
				Sheet1.Cell(11, 5).Value = Sheet1.Cell(11, 5).Value + round2(AgSum1 * 0.015, 2) + round2(AgSum2 * 0.014, 2)
			End If
		End With
	Next

End Sub<     č                Ě Arial D:\STDs.!!!\Zap\Čäĺŕë\Čäĺŕë       ˙˙˙˙                 ˙˙ 
 CSheetPageSheet1Âĺäîěîńňü íŕ÷čńëĺíčé
         61    7                                     X   .   :   :   +   %   "   "   +   ?   4   +   +   +   4   :   @                                                                    ˙˙  CRow ˙˙  CCell6     ˙˙   (   54><>ABL  =0G8A;5=89                        ˙˙      0@B  2 0 0 9   3.                      ! ˙˙#       0G8A;5=85                    ! ˙˙#       %                     ! ˙˙#     $"                    ! ˙˙#                        ! ˙˙#  
   !C<<0                    ! ˙˙#                        ! ˙˙#    
   515B                    ! ˙˙#       @548B                               B>3>:                                                 "                         "                         =Sum(E5:E11)"    0f                        "                          "                           "                              ˙˙    .   5=A8>==K9  D>=4  ( 3 3 . 2 % )                       ˙˙    ŕ                         " ˙˙   ę                        " ˙˙   ÷(\ÂË@                    " ˙˙  hffffŽ@                    " ˙˙                        " ˙˙       9 2                     " ˙˙       6 5 1                        ˙˙    (   5=A8>==K9  D>=4  ( 4 % )                       ˙˙    @                          " ˙˙                        " ˙˙   ăzŽGá3@                    " ˙˙                              " ˙˙                        " ˙˙       9 2                     " ˙˙       6 5 1                        ˙˙    .   $>=4  157@01>B8FK  ( 1 . 6 % )                       ˙˙    >                          " ˙˙  A                         " ˙˙   ŘŁp=
×4@                    " ˙˙˙˙ 433333a@                    " ˙˙                        " ˙˙       9 2                     " ˙˙       6 5 3                      " ˙˙         ( 1 , 6   % )                       ˙˙    >                          " ˙˙                           " ˙˙   Ľp=
×#(@                   =round2(B8*C8/100, 2)" ˙˙       0@                    " ˙˙                              =G7" ˙˙       9 2                    =H7" ˙˙       6 5 3                      " ˙˙         ( 2 , 2 % )                       ˙˙    đU                          " ˙˙   -1                        " ˙˙                       =round2(B9*C9/100, 2)" ˙˙       F@                    " ˙˙                       =G7" ˙˙       9 2                    =H7" ˙˙       6 5 3                        ˙˙    2   5AG0AB=K9  A;CG09  ( 1 . 0 2 % )                       ˙˙    Ř'                          " ˙˙  A                         " ˙˙   Ľp=
×#(@                    " ˙˙  ŻGázîU@                    " ˙˙                               " ˙˙       9 2                     " ˙˙       6 5 6                        ˙˙    2   5B@C4>A?>A>1=>ABL  ( 1 . 4 % )                       ˙˙    °6                          " ˙˙  A                         " ˙˙   53333s0@                    " ˙˙  ^@                    " ˙˙                               " ˙˙       9 2                     " ˙˙       6 5 2                     
   % ˙˙$       >4@0745;5=85                    % ˙˙$    
   !C<<0                      ˙˙#     0G8A;5=>                      ˙˙"                          ˙˙#     #45@60=>                      ˙˙#                          ˙˙#                          ˙˙#                          ˙˙"                          ˙˙#       2K40G5                  
   % ˙˙%       !>B@C4=8:                    ! ˙˙%       ;L3>BK                      ˙˙#       0 1 1 0                       ˙˙#    
   A53>                      ˙˙#       2 1 0 0                       ˙˙#       2 1 1 0                       ˙˙#       2 1 2 0                       ˙˙#       2 1 3 0                       ˙˙#    
   A53>                      ˙˙"                     
          
   B>3>                    "                                  " ˙˙     ę                        " ˙˙     ę                        " ˙˙    K                         " ˙˙    ü
                         " ˙˙    0u                          " ˙˙                               " ˙˙    Đ
W                         " ˙˙    0ú                      
   "  
       ? ? ? ? ? ? ? ? ? ?                     "  
                         "  
     Q%                        "  
     Q%                        "  
    K                         "  
    ü
                         "  
    0u                          "  
                               "  
    Đ
W                         "  
    0FÎ                      
   " ˙˙       1                     " ˙˙    h(.                         " ˙˙    [                         " ˙˙    [                         " ˙˙    `b                         " ˙˙    ŔÔ                         " ˙˙    0u                          " ˙˙                               " ˙˙    đ8	                         " ˙˙    TR                       
   " ˙˙       2                     " ˙˙                          " ˙˙                             " ˙˙                             " ˙˙    `ă                         " ˙˙    @                         " ˙˙                         " ˙˙                         " ˙˙     đ                         " ˙˙    ŕĽ~                       
   " ˙˙       3                     " ˙˙                          " ˙˙     -1                        " ˙˙     -1                        " ˙˙    ŔĆ-                         " ˙˙                             " ˙˙                         " ˙˙                         " ˙˙    @á3                         " ˙˙    ŔKý                       
   "         p p p p p p                     "                           "       ´Ä                        "       ´Ä                        "                            "                            "                            "                            "                            "       ´Ä                      
   " ˙˙       f f g                     " ˙˙                          " ˙˙     ´Ä                        " ˙˙     ´Ä                        " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                          " ˙˙     ´Ä                      d   d   d   d    <     č                Ě Arial                              <     č                Ě Arial                              <     č                Ě Arial                              <     č                Ě Arial                              <     č                Ě Arial                              <     č                Ě Arial                                2   2   
    <     č    Č           Ě Arial ĐDb                   
    <     č   đ           Ě Arial Sub ShtBook_OnLoadEnd Su  <     č    đ           Ě Arial čňĺ F1 ëĺííîăî ýëĺěĺíňŕ 
.av <     č               Ě Arial čňĺ F1  (F7)    1      čÍe <     č           ˙˙ŕ Ě Arial ˙˙               đóňęč    <     č           ((( Ě Arial ˙˙               htRange) <     č           ˙˙˙ Ě Arial ˙˙                   ing  <     č            ((( Ě Arial ˙˙                   Đ#g <     č   Č           Ě Arial                      čÍe	 <     č    ´           Ě Arial îéńňâ ű (F7) ńîńňî˙íč˙ äë˙ î-    ˙˙˙                               ČČČ                               ČČČ                                  ÜÜÜ                                  ÜÜÜ        ˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙ý˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙                      ˙˙ŕ                                  ˙˙˙˙                             ˙˙ŕ                              ˙˙˙˙                              ˙˙˙˙                              	 ˙˙˙˙                             
 ŕ˙˙                               ˙˙˙                                ŕ˙˙                                ŕ˙˙                               ŕ˙˙                              ˙˙˙                                ˙˙˙                                 ŕ˙˙                                 ˙˙˙                                ˙˙˙                               ˙˙˙˙                               ˙˙˙˙                                ˙˙˙˙                               ŔŔŔ                                  ČČČ                               ŔŔŔ                               ČČČ                               ŔŔŔ                              ˙˙˙                                  ˙˙˙                              ŔŔŔ                              , ŕ˙˙                                + ˙˙˙˙                               * Żîî                                 ) Ŕ˙˙                                 ' ˙˙ŕ                              & ŕŕŕ                               # ŕŕŕ                             " ŕŕŕ                                 ! ˙˙ŕ                                ˙˙ŕ                              $ ŕŕŕ                              % ŕŕŕ                              ( ŕ˙˙                                   bject("Scripting.Dictionary")
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
' âűáčđŕĺě ęîäű, çŕ čńęëţ÷ĺíčĺě ńďĺö-ęîäîâ "âűďëŕňŕ", "äîëă ńîňđóäíčęŕ" č "äîëă ďđĺäďđč˙ňč˙"
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
		SqlQr = "SELECT MTD_ID, MTD_CODE, 0 FROM MTD_CODES WHERE ( MTD_TYPE = 101 ) AND ( ISNULL( MTD_MODE, 0 ) <> 2 ) " & _
				vbNewLine & "UNION all " & vbNewLine & _
				"SELECT MTD_ID, MTD_CODE, 1 FROM MTD_CODES WHERE ( MTD_TYPE = 102 ) AND ( ISNULL( MTD_MODE, 0 ) <> 3 ) AND ( ISNULL( MTD_MODE, 0 ) <> 1 ) ORDER BY 3, 2"
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
		.Rows = StartNumberOfRows
		.Columns = 8
		' î÷čńňčëč ńňđîęó čňîăîâ
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

	' âűáđŕëč âńĺ ęîäű, ďî ęîňîđűě ĺńňü îáîđîňű
	For i = 0 To UBound( arData, 2 )
		DicCodes( arData( 4, i )) = 0
	Next

	With Sheet( 1 )
		If ( DicCodes.Count + 2 ) > 8 Then
			.Columns = 2 + DicCodes.Count
		End If

		TotalsCols( 0 ) = 2
		' Đčńóĺě řŕďęó ń ďĺđĺ÷íĺě ęîäîâ íŕ÷čńëĺíčé č óäĺđćŕíčé
		k = 3
		' arCodes( 2, i ) = 0 čëč 1. Íóćíî äë˙ ňîăî, ÷ňîá â óäĺđćŕíč˙ő ńäĺëŕňü ńäâčă íŕ îäčí ńňîëáĺö,
		' ÷ňîáű îńňŕâčňü ěĺńňî äë˙ ńîëáöŕ čňîăîâ ďî óäĺđćŕíč˙ě
		hdr = "1:Íŕ÷čńëĺíî"
		For i = 0 To UBound( arCodes, 2 )
			If i = 0 Then
			ElseIf ( arCodes( 2, i - 1 ) = 0 ) And ( arCodes( 2, i ) = 1 ) Then
				hdr = hdr & vbLf & "2:Âńĺăî" & vbLf & "1:Óäĺđćŕíî"
				TotalsCols( 1 ) = k
			End If
			If DicCodes.Exists( arCodes( 0, i )) Then
				DicCodes( arCodes( 0, i )) = k + arCodes( 2, i )
				hdr = hdr & vbLf & "2:" & arCodes( 1, i )
				k = k + 1
			End If
		Next
		hdr = hdr & vbLf & "2:Âńĺăî" & vbLf & "1:Ę âűäŕ÷ĺ"

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
		SqlQr = "Select AG_TREE.P0, P_JOURNAL.AG_ID, Groups.AG_NAME, AGENTS.AG_NAME, P_JOURNAL.MTC_ID, Sum(JP_SUM) " & _
		        "From P_JOURNAL Inner Join AGENTS On P_JOURNAL.AG_ID = AGENTS.AG_ID " & _
		        "               Inner Join AG_TREE On P_JOURNAL.AG_ID = ID " & _
		        "               Inner Join MTD_CODES  m On P_JOURNAL.MTC_ID = m.MTD_ID " & _
		        "               Left Join AGENTS Groups On P0 = Groups.AG_ID " & _
		        "Where JP_DONE = 1 And SHORTCUT = 0 AND ( ISNULL( MTD_MODE, 0 ) <> 1 ) AND ( ISNULL( MTD_MODE, 0 ) <> 2 ) AND ( ISNULL( MTD_MODE, 0 ) <> 3 ) " & _
		        "                  And ( IsNull( MTD_MASK, 0 ) < 8 ) " & _
		        "                  And C_PERIOD = ? " & _
		        "Group By AG_TREE.P0, P_JOURNAL.AG_ID, Groups.AG_NAME, AGENTS.AG_NAME, P_JOURNAL.MTC_ID " & _
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
		FOT = CalcDepsSumLimit_Pens( CodeID, Repcalc, max, False, d )
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
		FOT = CalcDepsSumLimit_Pens( CodeID, Repcalc, max, True,d )
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


		If Year(d) < 2009 Then
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

		If Month(d) = 1 And Year(d) = 2009 Then
			Calc2030 CodeID, Repcalc, d, max
		Else
			FOT = CalcDepsSumLimit( CodeID, Repcalc, d, max )
			sum1 = sum1 + Round2( FOT * percent / 100 , 2 )
			.Cell( 11, 3 ).value = FOT
			.Cell( 11, 5 ).value = Round2( FOT * percent / 100 , 2 )
		End If

'		.Cell( 4, 5 ).value = sum1
	End With
End Sub

'-----------------------------------------------------------------------------------------------
' Function GetCodeID( CodeName ) - âĺđíóňü čäĺíňčôčęŕňîđ ďî čěĺíč ęîäŕ/ěĺňîäŕ 
'-----------------------------------------------------------------------------------------------
Function GetCodeID( CodeName )
	Dim SqlQr, ar
	GetCodeID = 0
	SqlQr = "Select MTD_ID From MTD_CODES Where MTD_NAME = '" & CodeName & "'"
	Select Case App.AppType
	Case "DAO"
		If GetRows_DAO( SqlQr, ar ) Then GetCodeID = ar( 0, 0 )
	Case "OLEDB"
		If GetRows_ADO( SqlQr, adCmdText, ar ) Then GetCodeID = ar( 0, 0 )
	End Select
End Function

'________________________________________________
' Function QueryWithParams( SqlText, Options, arParams )
' Âűďîëíĺíčĺ çŕďđîńŕ ń ďĺđĺäŕ÷ĺé â íĺăî ďŕđŕěĺňđîâ
' Âîçâđŕůŕĺňń˙ îáúĺęň ňčďŕ RecordSet
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

	' óńňŕíŕâëčâŕĺě ďŕđŕěĺňđű çŕďđîńŕ
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
Function GetRows_ADO( SQLText, Options, ByRef ar )R o o t   E n t r y                                               ˙˙˙˙˙˙˙˙                               
ýEhĘ>         C o n t e n t s                                                  ˙˙˙˙   ˙˙˙˙                                    ?   	v       S u m m a r y I n f o r m a t i o n                           (  ˙˙˙˙˙˙˙˙˙˙˙˙                                        i                                                                          ˙˙˙˙˙˙˙˙˙˙˙˙                                                                        	   
                                             ý˙˙˙         ţ˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙ţ˙˙˙ý˙˙˙ţ˙˙˙ţ˙˙˙@   A   B   C   D   E   F   G   H   I   J   K   L   M   N   O   P   Q   R   S   T   U   V   W   X   Y   Z   [   \   ]   ^       ˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙               ţ˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙ţ˙   ŕňůOhŤ +'łŮ   ŕňůOhŤ +'łŮ0   9        x            ¤      °      ź      Č      Ô   	   ŕ   
   ě      ř               $      ,     ă        Âĺäîěîńňü íŕ÷čńëĺíčé.ash                                                            463 @   ŔŁ  @   @÷üEhĘ@   @kçôśýĆ      Ŕęöĺíň 7.0                                                                                                                                                                           ˙M'Option Explicit
'#include "definition.avb"
'#include "ado.inc"

Const StartNumberOfRows = 15
Const StartNumberOfColumns = 2

Const MTC_PRIVILEGE_NAME = "0801"
Const MTC_WPF_NAME = "ÔÎŇ"

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
' âűáčđŕĺě ęîäű, çŕ čńęëţ÷ĺíčĺě ńďĺö-ęîäîâ "âűďëŕňŕ", "äîëă ńîňđóäíčęŕ" č "äîëă ďđĺäďđč˙ňč˙"
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
		SqlQr = "SELECT MTD_ID, MTD_CODE, 0 FROM MTD_CODES WHERE ( MTD_TYPE = 101 ) AND ( ISNULL( MTD_MODE, 0 ) <> 2 ) " & _
				vbNewLine & "UNION all " & vbNewLine & _
				"SELECT MTD_ID, MTD_CODE, 1 FROM MTD_CODES WHERE ( MTD_TYPE = 102 ) AND ( ISNULL( MTD_MODE, 0 ) <> 3 ) AND ( ISNULL( MTD_MODE, 0 ) <> 1 ) ORDER BY 3, 2"
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
		.Rows = StartNumberOfRows
		.Columns = 8
		' î÷čńňčëč ńňđîęó čňîăîâ
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

	' âűáđŕëč âńĺ ęîäű, ďî ęîňîđűě ĺńňü îáîđîňű
	For i = 0 To UBound( arData, 2 )
		DicCodes( arData( 4, i )) = 0
	Next

	With Sheet( 1 )
		If ( DicCodes.Count + 2 ) > 8 Then
			.Columns = 2 + DicCodes.Count
		End If

		TotalsCols( 0 ) = 2
		' Đčńóĺě řŕďęó ń ďĺđĺ÷íĺě ęîäîâ íŕ÷čńëĺíčé č óäĺđćŕíčé
		k = 3
		' arCodes( 2, i ) = 0 čëč 1. Íóćíî äë˙ ňîăî, ÷ňîá â óäĺđćŕíč˙ő ńäĺëŕňü ńäâčă íŕ îäčí ńňîëáĺö,
		' ÷ňîáű îńňŕâčňü ěĺńňî äë˙ ńîëáöŕ čňîăîâ ďî óäĺđćŕíč˙ě
		hdr = "1:Íŕ÷čńëĺíî"
		For i = 0 To UBound( arCodes, 2 )
			If i = 0 Then
			ElseIf ( arCodes( 2, i - 1 ) = 0 ) And ( arCodes( 2, i ) = 1 ) Then
				hdr = hdr & vbLf & "2:Âńĺăî" & vbLf & "1:Óäĺđćŕíî"
				TotalsCols( 1 ) = k
			End If
			If DicCodes.Exists( arCodes( 0, i )) Then
				DicCodes( arCodes( 0, i )) = k + arCodes( 2, i )
				hdr = hdr & vbLf & "2:" & arCodes( 1, i )
				k = k + 1
			End If
		Next
		hdr = hdr & vbLf & "2:Âńĺăî" & vbLf & "1:Ę âűäŕ÷ĺ"

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
		SqlQr = "Select AG_TREE.P0, P_JOURNAL.AG_ID, Groups.AG_NAME, AGENTS.AG_NAME, P_JOURNAL.MTC_ID, Sum(JP_SUM) " & _
		        "From P_JOURNAL Inner Join AGENTS On P_JOURNAL.AG_ID = AGENTS.AG_ID " & _
		        "               Inner Join AG_TREE On P_JOURNAL.AG_ID = ID " & _
		        "               Inner Join MTD_CODES  m On P_JOURNAL.MTC_ID = m.MTD_ID " & _
		        "               Left Join AGENTS Groups On P0 = Groups.AG_ID " & _
		        "Where JP_DONE = 1 And SHORTCUT = 0 AND ( ISNULL( MTD_MODE, 0 ) <> 1 ) AND ( ISNULL( MTD_MODE, 0 ) <> 2 ) AND ( ISNULL( MTD_MODE, 0 ) <> 3 ) " & _
		        "                  And ( IsNull( MTD_MASK, 0 ) < 8 ) " & _
		        "                  And C_PERIOD = ? " & _
		        "Group By AG_TREE.P0, P_JOURNAL.AG_ID, Groups.AG_NAME, AGENTS.AG_NAME, P_JOURNAL.MTC_ID " & _
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
		FOT = CalcDepsSumLimit_Pens( CodeID, Repcalc, max, False, d )
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
		FOT = CalcDepsSumLimit_Pens( CodeID, Repcalc, max, True,d )
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


		If Year(d) < 2009 Then
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

		If Month(d) = 1 And Year(d) = 2009 Then
			Calc2030 CodeID, Repcalc, d, max
		Else
			FOT = CalcDepsSumLimit( CodeID, Repcalc, d, max )
			sum1 = sum1 + Round2( FOT * percent / 100 , 2 )
			.Cell( 11, 3 ).value = FOT
			.Cell( 11, 5 ).value = Round2( FOT * percent / 100 , 2 )
		End If

'		.Cell( 4, 5 ).value = sum1
	End With
End Sub

'-----------------------------------------------------------------------------------------------
' Function GetCodeID( CodeName ) - âĺđíóňü čäĺíňčôčęŕňîđ ďî čěĺíč ęîäŕ/ěĺňîäŕ 
'-----------------------------------------------------------------------------------------------
Function GetCodeID( CodeName )
	Dim SqlQr, ar
	GetCodeID = 0
	SqlQr = "Select MTD_ID From MTD_CODES Where MTD_NAME = '" & CodeName & "'"
	Select Case App.AppType
	Case "DAO"
		If GetRows_DAO( SqlQr, ar ) Then GetCodeID = ar( 0, 0 )
	Case "OLEDB"
		If GetRows_ADO( SqlQr, adCmdText, ar ) Then GetCodeID = ar( 0, 0 )
	End Select
End Function

'________________________________________________
' Function QueryWithParams( SqlText, Options, arParams )
' Âűďîëíĺíčĺ çŕďđîńŕ ń ďĺđĺäŕ÷ĺé â íĺăî ďŕđŕěĺňđîâ
' Âîçâđŕůŕĺňń˙ îáúĺęň ňčďŕ RecordSet
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

	' óńňŕíŕâëčâŕĺě ďŕđŕěĺňđű çŕďđîńŕ
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
Function GetRows_ADO( SQLText, Options, ByRef ar )" ˙˙                         " ˙˙                         " ˙˙    `ěS                         " ˙˙    ¨2                         " ˙˙    °­                         " ˙˙    lk                          " ˙˙    č                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŹĚ                         " ˙˙    ´E                          " ˙˙    4   C4V=>2  ;5:AV9  V:B>@>28G                    " ˙˙                          " ˙˙    ŔŘ§                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŔŘ§                         " ˙˙    HE                         " ˙˙    `[                         " ˙˙    °­                         " ˙˙    Đ                         " ˙˙                         " ˙˙                         " ˙˙    O                         " ˙˙    ¸Ďm                         " ˙˙    	:                          " ˙˙    @   54254T2  ;5:A0=4@  >;>48<8@>28G                    " ˙˙                          " ˙˙    Ë¤                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ë¤                         " ˙˙    dÔ                         " ˙˙    ŔK                         " ˙˙    ŕĽ                         " ˙˙     ý                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    $Ă                         " ˙˙    \                          " ˙˙    4   "0@A:>2  5B@>  0;5=B8=>28G                    " ˙˙                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    pýŞ                         " ˙˙    pýŞ                         " ˙˙    °š                         " ˙˙    dk                         " ˙˙    äľ                         " ˙˙                             " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    xá                         " ˙˙    ř                          " ˙˙    :   "5@OT2  ;5:A0=4@  VB@>D0=>28G                    " ˙˙                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    PĹ                         " ˙˙    PĹ                         " ˙˙                             " ˙˙    Tó                         " ˙˙    Üů                         " ˙˙    /                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    4­#                         " ˙˙    ÝĄ                          " ˙˙    2   ">;?0  =0B>;V9  =4@V9>28G                    " ˙˙    h(.                         " ˙˙    ë	                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    PŐz                         " ˙˙    `Ŕ                         " ˙˙    ĚE                         " ˙˙    °§                         " ˙˙    ŘS                         " ˙˙    čË                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    <                         " ˙˙    $łs                          " ˙˙    0   %><O:  ;5:AV9  0A8;L>28G                    " ˙˙                          " ˙˙                               " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ô                         " ˙˙    "ż                         " ˙˙    dŁ                         " ˙˙    Ň                         " ˙˙    té                         " ˙˙    Ŕ%                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    "                         " ˙˙    ř                          " ˙˙    8   (0;V<>2  =4@V9  >;>48<8@>28G                    " ˙˙                          " ˙˙    ŕyŻ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŕyŻ                         " ˙˙    ´_                         " ˙˙    p                         " ˙˙    8Á                         " ˙˙                             " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ä°                         " ˙˙    üČ                          "         0:552A:89  D- ;                    "                           "      Ŕáä                         "                            "                            "                            "                            "      Ŕáä                         "      ¨!                         "      ŕ                         "      đI                         "      _                         "                            "                            "                            "      V)                         "      ¸ť                          " ˙˙    :   =8I5=:>  !5@3V9  ;5:A0=4@>28G                    " ˙˙                          " ˙˙    Ŕáä                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ŕáä                         " ˙˙    ¨!                         " ˙˙    ŕ                         " ˙˙    đI                         " ˙˙    _                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    V)                         " ˙˙    ¸ť                          "         0@:5B8=3                    "                           "       Y                         "                            "                            "                            "                            "       Y                         "      Ä6                         "                                "                               "      Ŕ                         "                            "                            "                            "      ü&                         "      ě\Ů                          " ˙˙    2   0:0@5=:>   8<<0  0A8;V2=0                    " ˙˙    h(.                         " ˙˙    [                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    [                         " ˙˙    `b                         " ˙˙    ŔÔ                         " ˙˙    0u                          " ˙˙                               " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    đ8	                         " ˙˙    TR                          " ˙˙    ,   (T=DT;L4  ;V1  ;V1>28G                    " ˙˙                          " ˙˙    Ë¤                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ë¤                         " ˙˙    dÔ                         " ˙˙    ŔK                         " ˙˙    ŕĽ                         " ˙˙     ý                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    $Ă                         " ˙˙    \                          "      @   >4@. 1 3 . 0 5 . 0 3 . 1 1 6 M W E ( !25B;8G=K9)                     "                           "      Xâ                        "      0U                         "      âQ                         "                            "      XÚ                         "      `ô	                        "      ÄK                        "      lQ.                         "      ů                         "      I                         "                            "                            "                            "      HŹ                        "      Hp                         " ˙˙    8   0H52  >;>48<8@  5==04V9>28G                    " ˙˙                          " ˙˙    ŔŘ§                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŔŘ§                         " ˙˙    HE                         " ˙˙    `[                         " ˙˙    °­                         " ˙˙    Đ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    (P                         " ˙˙                              " ˙˙    >   >;>4V=  >;>48<8@  >;>48<8@>28G                    " ˙˙                          " ˙˙    đĹ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ô,	                         " ˙˙    Äň                         " ˙˙    ŕÖ                         " ˙˙    (                         " ˙˙    L                         " ˙˙    ň                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    X                         " ˙˙    ll                          " ˙˙    ,   0B5;;>  0E0@  3>@>28G                    " ˙˙                          " ˙˙    ŕČ                         " ˙˙    0U                         " ˙˙    âQ                         " ˙˙                         " ˙˙                         " ˙˙     y                         " ˙˙    pÝ                         " ˙˙    k                         " ˙˙    _˙˙˙˙˙˙                    " ˙˙                              " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ć                         " ˙˙    t:e                          " ˙˙    2   C:C9  >ABO=B8=  1@0<>28G                    " ˙˙                          " ˙˙    Ŕáä                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ŕáä                         " ˙˙    ¨!                         " ˙˙    ŕ                         " ˙˙    đI                         " ˙˙    _                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    V)                         " ˙˙    ¸ť                          " ˙˙    0   OE  !5@3V9  >;>48<8@>28G                    " ˙˙                          " ˙˙     Ii                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     Ii                         " ˙˙    D9                         " ˙˙                             " ˙˙                             " ˙˙    ¸Ą                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                             " ˙˙    EV                          " ˙˙    4   0?:>  "5BO=0  ;5:A0=4@V2=0                    " ˙˙                          " ˙˙    ŔŘ§                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŔŘ§                         " ˙˙    HE                         " ˙˙    `[                         " ˙˙    °­                         " ˙˙    Đ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    (P                         " ˙˙                              " ˙˙    2   @8E>4L:>  !5<5=  @BC@>28G                    " ˙˙                          " ˙˙    č(                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    č(                         " ˙˙    <Ň                         " ˙˙    Í                          " ˙˙    ,3                          " ˙˙    ¸=                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    4                         " ˙˙    ´ý                           " ˙˙    H    ><0=>2AL:89  >;>48<8@  @>=VA;02>28G                    " ˙˙                          " ˙˙    Ë¤                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ë¤                         " ˙˙    dÔ                         " ˙˙    ŔK                         " ˙˙    ŕĽ                         " ˙˙     ý                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    $Ă                         " ˙˙    \                          " ˙˙    <   !2VB;8G=0  0@30@8B0  8:>;0W2=0                    " ˙˙    h(.                         " ˙˙    ŕý                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŕý                         " ˙˙                         " ˙˙    p                          " ˙˙    %                          " ˙˙    ,                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ć                          " ˙˙    Ě                          " ˙˙    :   !2VB;8G=89  ;5:AV9  0A8;L>28G                    " ˙˙                          " ˙˙    ÷s                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ­                         " ˙˙    Ľy                         " ˙˙                              " ˙˙    ¨n                         " ˙˙    T7                         " ˙˙    ¸ş                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ô÷                         " ˙˙    0­c                          " ˙˙    8   !C;L=V:>2  !5@3V9  0;5@V9>28G                    " ˙˙                          " ˙˙    Ő                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ő                         " ˙˙    Đă                         " ˙˙    ŔE                         " ˙˙    ŕ"                         " ˙˙     H                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    &                         " ˙˙    đ
Ż                          " ˙˙    6   !CE>@C:>2  <8B@>  V:B>@>28G                    " ˙˙                          " ˙˙    ŔŘ§                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŔŘ§                         " ˙˙    HE                         " ˙˙    `[                         " ˙˙    °­                         " ˙˙    Đ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    (P                         " ˙˙                              " ˙˙    0   "8BN:  0:A8<  8:>;09>28G                    " ˙˙                          " ˙˙     Ým                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     Ým                         " ˙˙    Ěâ                         " ˙˙    2                         " ˙˙    @                         " ˙˙    Ŕ¨                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    L×                         " ˙˙    ´Z                          " ˙˙    8   "@5BLO:>2  0=8W;  ;5:AV9>28G                    " ˙˙                          " ˙˙     7                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     7                          " ˙˙    Ü*                         " ˙˙    P4                         " ˙˙    (                         " ˙˙    ö                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    lď                         " ˙˙    4H                          " ˙˙    .   $>=>B>2  20=  8E09;>28G                    " ˙˙                          " ˙˙    Ő                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ő                         " ˙˙    Đă                         " ˙˙    ŔE                         " ˙˙    ŕ"                         " ˙˙     H                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    &                         " ˙˙    đ
Ż                          " ˙˙    :   (8<0=AL:89  5==04V9  235=>28G                    " ˙˙                          " ˙˙     |                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     |                         " ˙˙    ¨.                         " ˙˙     î                         " ˙˙     w                         " ˙˙     á                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ¨t                         " ˙˙    Xx                          "      @   >4@. 1 3 . 0 7 . 0 3 . 1 1 5 M W E ( !>;>4  . !. )                     "                           "      ŕE                        "                            "                            "                            "                            "      ŕE                        "      pĚJ                         "      pX
                         "      8,                         "                               "                            "                            "                            "       k]                         "      @Ú§                         " ˙˙    <   5=5FL:89  ;5:Ai 9  5==04i 9>28G                    " ˙˙                          " ˙˙    Ë¤                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ë¤                         " ˙˙    dÔ                         " ˙˙    ŔK                         " ˙˙    ŕĽ                         " ˙˙     ý                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    $Ă                         " ˙˙    \                          " ˙˙    6   0<0T2  =0B>;i 9  8:>;09>28G                    " ˙˙                          " ˙˙     {                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     {                         " ˙˙    dß                         " ˙˙    Đx                         " ˙˙    h<                         " ˙˙    Ř˝                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    tR                         " ˙˙    ,Fe                          " ˙˙    4   !>;>4  >;>48<8@  !5@3V9>28G                    " ˙˙                          " ˙˙    Ŕáä                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ŕáä                         " ˙˙    ¨!                         " ˙˙    ŕ                         " ˙˙    đI                         " ˙˙    _                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    V)                         " ˙˙    ¸ť                          "      :   >4@074. " >4@O4=K5  @- BK  #!0"                     "                           "      ŕÚ                         "                            "                            "                            "                            "      ŕÚ                         "      <                         "      đĆ                         "      xc                         "      HŐ                          "                            "                            "                            "      ě                         "      ôĆq                          " ˙˙    8   $54>@>2  8:>;0  >;>48<8@>28G                    " ˙˙                          " ˙˙    ŕÚ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŕÚ                         " ˙˙    <                         " ˙˙    đĆ                         " ˙˙    xc                         " ˙˙    HŐ                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ě                         " ˙˙    ôĆq                          "      0   >4@074. - >1I5E>7. @0AE>4K                    "                           "      (F                        "                            "                            "                            "      Dó                        "      l9                        "      Ô´U                        "      {                         "      Ôô;                         "      Dç$                         "                            "                            "      źŐ                        "      ¸­                        "      ´˙                         " ˙˙    6   ;5:A0=4@>20  "5BO=0  ;;V2=0                    " ˙˙                          " ˙˙    xđĽ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    xđĽ                         " ˙˙    ţ                         " ˙˙    Q                         " ˙˙     Š                         " ˙˙    ˙                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Dř                         " ˙˙    4ř                          " ˙˙    .   E:>7>2  >;>48<8@  ;;VG                    " ˙˙                          " ˙˙    ŔŘ§                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŔŘ§                         " ˙˙    HE                         " ˙˙    `[                         " ˙˙    °­                         " ˙˙    Đ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    (P                         " ˙˙                              " ˙˙    0   5=5FL:0  ;5=0  235=V2=0                    " ˙˙                          " ˙˙     Ii                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     Ii                         " ˙˙    D9                         " ˙˙                             " ˙˙                             " ˙˙    ¸Ą                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                             " ˙˙    EV                          " ˙˙    8   C30T2AL:0  0@8=0  V:B>@V2=0                    " ˙˙                          " ˙˙    Ő                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ő                         " ˙˙    Đă                         " ˙˙    ŔE                         " ˙˙    ŕ"                         " ˙˙     H                         " ˙˙                         " ˙˙                         " ˙˙    `M/                         " ˙˙    đáU                         " ˙˙    ˝                          " ˙˙    .   5@=V:  @B5<  !5@3V9>28G                    " ˙˙                          " ˙˙     ˇ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     ˇ                         " ˙˙     z                         " ˙˙    Š                         " ˙˙    ŔÔ                         " ˙˙    @                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     !                         " ˙˙    `	                          " ˙˙    0   V=>:C@>20  ;L30  .@VW2=0                    " ˙˙                          " ˙˙    ŔŘ§                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŔŘ§                         " ˙˙    HE                         " ˙˙    `[                         " ˙˙    °­                         " ˙˙    Đ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    (P                         " ˙˙                              " ˙˙    2   >;>E>20  ;5=0  V:B>@V2=0                    " ˙˙                          " ˙˙     źž                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     źž                         " ˙˙                             " ˙˙    Đ                         " ˙˙    Hč                         " ˙˙    ř$                         " ˙˙                         " ˙˙                         " ˙˙                             " ˙˙    Üť                         " ˙˙    Dł                          " ˙˙    0   >@1C=>20  ;5=0  5B@V2=0                    " ˙˙                          " ˙˙     |                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     |                         " ˙˙    ¨.                         " ˙˙     î                         " ˙˙     w                         " ˙˙     á                          " ˙˙                         " ˙˙                         " ˙˙    @KL                         " ˙˙    čżf                         " ˙˙    ź+                          " ˙˙    2   @5AA   0WA0  ;5:A0=4@V2=0                    " ˙˙                          " ˙˙    ŔŘ§                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŔŘ§                         " ˙˙    HE                         " ˙˙    `[                         " ˙˙    °­                         " ˙˙    Đ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    (P                         " ˙˙                              " ˙˙    0   >1@N:  !5@3V9  0A8;L>28G                    " ˙˙                          " ˙˙     |                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     |                         " ˙˙    ¨.                         " ˙˙     î                         " ˙˙     w                         " ˙˙     á                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ¨t                         " ˙˙    Xx                          " ˙˙    0   20=>2  =0B>;V9  20=>28G                    " ˙˙                          " ˙˙    ŕÚ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŕÚ                         " ˙˙    <                         " ˙˙    đĆ                         " ˙˙    xc                         " ˙˙    HŐ                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ě                         " ˙˙    ôĆq                          " ˙˙    B   V@VG5=:>  ;5:A0=4@  ;5:A0=4@>28G                    " ˙˙                          " ˙˙    @]Ć                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    @]Ć                         " ˙˙    řŽ                         " ˙˙     ÷                         " ˙˙    Đű                         " ˙˙    °0                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ó#                         " ˙˙    (˘                          " ˙˙    2   >;>4=89  0;5@V9  02;>28G                    " ˙˙    h(.                         " ˙˙    ŔĎj                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŔĎj                         " ˙˙    ¤                         " ˙˙    ŕ"                         " ˙˙    p                         " ˙˙    ¤                          " ˙˙                         " ˙˙                         " ˙˙    @w                         " ˙˙    DŐ'                         " ˙˙    |úB                          " ˙˙    4   @02G5=:>  0@8=0  !5@3VW2=0                    " ˙˙    h(.                         " ˙˙    Ŕ\                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ŕ\                         " ˙˙                         " ˙˙    `m                          " ˙˙    X                          " ˙˙    Đ                           " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Š                          " ˙˙    8ł                          " ˙˙    4   C4@O2F520  04VO  0A8;V2=0                    " ˙˙                          " ˙˙    ŔŘ§                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŔŘ§                         " ˙˙    HE                         " ˙˙    `[                         " ˙˙    °­                         " ˙˙    Đ                         " ˙˙                         " ˙˙                         " ˙˙    ŔĆ-                         " ˙˙    čL                         " ˙˙    ŘÁ[                          " ˙˙    @   C7=5F>20  /@>A;020  ;5:A0=4@V2=0                    " ˙˙                          " ˙˙     z                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     z                         " ˙˙    ŔŚ                         " ˙˙     q                         " ˙˙    8                         " ˙˙    ť                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ŕ                         " ˙˙    @d                          " ˙˙    *   C?G8:  0@VO  20=V2=0                    " ˙˙                          " ˙˙    ŕÚ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŕÚ                         " ˙˙    <                         " ˙˙    đĆ                         " ˙˙    xc                         " ˙˙    HŐ                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ě                         " ˙˙    ôĆq                          " ˙˙    2   8A5=:>  5@<0=  8E09;>28G                    " ˙˙                          " ˙˙    @@                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    @@                         " ˙˙    ÄP	                         " ˙˙     H                         " ˙˙    R                          " ˙˙    pb                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    \M                         " ˙˙    äČ4                          " ˙˙    8   ><0:>2AL:89  5>=V4  5B@>28G                    " ˙˙                          " ˙˙    xçh                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ä˝                         " ˙˙    \Ľq                         " ˙˙    o                         " ˙˙    F                         " ˙˙    #                         " ˙˙    Ž                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    °                         " ˙˙    Ź]                          " ˙˙    6   060@>2  >;>48<8@  <8B@>28G                    " ˙˙                          " ˙˙    ŕÚ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŕÚ                         " ˙˙    <                         " ˙˙    đĆ                         " ˙˙    xc                         " ˙˙    HŐ                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ě                         " ˙˙    ôĆq                          " ˙˙    :   0=H8;V=  ;5:A0=4@  59=V9>28G                    " ˙˙                          " ˙˙     şŰ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    @3                         " ˙˙    @Ó                        " ˙˙    d)'                         " ˙˙    ¸j                         " ˙˙    \ľ                         " ˙˙                              " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    |é0                         " ˙˙    ÄéÝ                          " ˙˙    4   0=H8;V=0  ==0  5==04VW2=0                    " ˙˙                          " ˙˙     $                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     $                         " ˙˙    R                         " ˙˙    ť                          " ˙˙    ŕ.                          " ˙˙    @8                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     u                         " ˙˙    ŕ)                          " ˙˙    4   >@>7>20  V=0V40  !5<5=V2=0                    " ˙˙                          " ˙˙    ŕÚ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŕÚ                         " ˙˙    <                         " ˙˙    đĆ                         " ˙˙    xc                         " ˙˙    HŐ                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ě                         " ˙˙    ôĆq                          " ˙˙    8   070@>2  ;5:A0=4@  !5@3V9>28G                    " ˙˙    h(.                         " ˙˙    ż:                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ż:                         " ˙˙    H                         " ˙˙    Č,                         " ˙˙    dK                          " ˙˙    <Z                          " ˙˙                         " ˙˙                         " ˙˙    `ă                         " ˙˙    S                         " ˙˙     l                           " ˙˙    8   52V4><0  @8=0  >;>48<8@V2=0                    " ˙˙                          " ˙˙    ŔŘ§                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŔŘ§                         " ˙˙    HE                         " ˙˙    `[                         " ˙˙    °­                         " ˙˙    Đ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    (P                         " ˙˙                              " ˙˙    :   545;V=  !5@3V9  >ABO=B8=>28G                      " ˙˙    h(.                         " ˙˙    @T                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    @T                         " ˙˙    Tď                         " ˙˙     ż                         " ˙˙    _                         " ˙˙    đŇ                          " ˙˙                         " ˙˙                         " ˙˙     Ii                         " ˙˙    *{                         " ˙˙    ,*                          " ˙˙    8   02;5=:>  VB0;V9  VB0;V9>28G                    " ˙˙                          " ˙˙     z                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     z                         " ˙˙    ŔŚ                         " ˙˙     q                         " ˙˙    8                         " ˙˙    ť                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ŕ                         " ˙˙    @d                          " ˙˙    2   ;V3V=  <8B@>  8:>;09>28G                    " ˙˙                          " ˙˙     sU                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ß                         " ˙˙    Ri                         " ˙˙    Ô:                         " ˙˙                             " ˙˙                             " ˙˙    ¸Ą                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    $                         " ˙˙    \MV                          " ˙˙    :   @>E>@>2  !5@3V9  >;>48<8@>28G                    " ˙˙                          " ˙˙    ŕÚ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŕÚ                         " ˙˙    <                         " ˙˙    đĆ                         " ˙˙    xc                         " ˙˙    HŐ                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ě                         " ˙˙    ôĆq                          " ˙˙    .    >B0@C  0@V0==0  >=V2=0                    " ˙˙                          " ˙˙    Ő                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ő                         " ˙˙    Đă                         " ˙˙    ŔE                         " ˙˙    ŕ"                         " ˙˙     H                         " ˙˙                         " ˙˙                         " ˙˙    <F                         " ˙˙    ĚÚˇ                         " ˙˙    ´Ä                          " ˙˙    (    O1:V=  .@V9  .@V9>28G                    " ˙˙                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     <´                         " ˙˙     <´                         " ˙˙    ŕ                         " ˙˙    ¨                         " ˙˙    TÍ                         " ˙˙    ô                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Đ                          " ˙˙    ĐŻ                          " ˙˙    4    O?>;>2  VB0;V9  V:B>@>28G                    " ˙˙                          " ˙˙    ŕ;f                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŕ;f                         " ˙˙    `Č                         " ˙˙    p                         " ˙˙    ¸                         " ˙˙                              " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    v                         " ˙˙    PĹS                          " ˙˙    <   !28@84>20  :A0=0  >;>48<8@V2=0                    " ˙˙                          " ˙˙    ŔŘ§                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŔŘ§                         " ˙˙    HE                         " ˙˙    `[                         " ˙˙    °­                         " ˙˙    Đ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    (P                         " ˙˙                              " ˙˙    >   !525@5=GC:  =4@V9  !B0=VA;02>28G                    " ˙˙                          " ˙˙     ˇ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     ˇ                         " ˙˙     z                         " ˙˙    Š                         " ˙˙    ŔÔ                         " ˙˙    @                         " ˙˙                         " ˙˙                         " ˙˙    @@                         " ˙˙    ŕ'a                         " ˙˙     óU                          " ˙˙    @   !525@5=GC:  0B0;VO  ;5:A0=4@V2=0                    " ˙˙                          " ˙˙    h(.                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    h(.                         " ˙˙    ľ                         " ˙˙    Tě                          " ˙˙    ü:                          " ˙˙    G                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ü#                         " ˙˙    l&                          " ˙˙    <   !84>@5=:>  5==04V9  8:>;09>28G                    " ˙˙                          " ˙˙    5˝                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    5˝                         " ˙˙    č[                         " ˙˙    ŔČ                         " ˙˙    `ä                         " ˙˙     "                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ¨+"                         " ˙˙    Ř	                          " ˙˙    4   !8<>=>28G  @8=0  3=0BL52=0                    " ˙˙                          " ˙˙    ŕ;f                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŕ;f                         " ˙˙    `Č                         " ˙˙    p                         " ˙˙    ¸                         " ˙˙                              " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    v                         " ˙˙    PĹS                          " ˙˙    :   !8@>B:V=  !5@3V9  !B0=VA;02>28G                    " ˙˙                          " ˙˙    ŔĎj                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŔĎj                         " ˙˙    čq                         " ˙˙    ŕ"                         " ˙˙    p                         " ˙˙    ¤                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    HJ                         " ˙˙    xW                          " ˙˙    >   !V@VG5=:>  0;5=B8=  0;5=B8=>28G                    " ˙˙                          " ˙˙    ŔĎj                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŔĎj                         " ˙˙    čq                         " ˙˙    ŕ"                         " ˙˙    p                         " ˙˙    ¤                          " ˙˙                         " ˙˙                         " ˙˙     %&                         " ˙˙    čo9                         " ˙˙    Ř_1                          " ˙˙    2   "@CEV=  8:>;0  ;5:AV9>28G                    " ˙˙                          " ˙˙     z                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     z                         " ˙˙    ŔŚ                         " ˙˙     q                         " ˙˙    8                         " ˙˙    ť                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ŕ                         " ˙˙    @d                          " ˙˙    :   $5>DV;0:B>2  =4@V9  V:B>@>28G                    " ˙˙                          " ˙˙     $ô                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     $ô                         " ˙˙    M#                         " ˙˙     â                         " ˙˙     q                         " ˙˙     w                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ,                         " ˙˙    Č                          " ˙˙    F   $5>DV;0:B>20  0B5@8=0  >;>48<8@V2=0                    " ˙˙    h(.                         " ˙˙    š*                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    š*                         " ˙˙                         " ˙˙    ŔÚ                          " ˙˙    °6                          " ˙˙     A                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    S                         " ˙˙    pf)                          " ˙˙    0   %V=528G  ;;O  8:>;09>28G                    " ˙˙                          " ˙˙     z                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     z                         " ˙˙    ŔŚ                         " ˙˙     q                         " ˙˙    8                         " ˙˙    ť                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ŕ                         " ˙˙    @d                          " ˙˙    *   &C?@C=  04VO  5B@V2=0                    " ˙˙    h(.                         " ˙˙    @Ź'                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    @Ź'                         " ˙˙                         " ˙˙     Ë                          " ˙˙    Č2                          " ˙˙    đ<                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ř:                         " ˙˙    hq&                          " ˙˙    ,   (0;L=>2  048<  20=>28G                    " ˙˙                          " ˙˙    ŕÚ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŕÚ                         " ˙˙    <                         " ˙˙    đĆ                         " ˙˙    xc                         " ˙˙    HŐ                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ě                         " ˙˙    ôĆq                          "      >   >4@074. - M:A?;C0B. @0AE>4K  >D8A0                    "                           "       Đü                        "                            "                            "                            "                            "       Đü                        "      <ę~                         "      j                         "      `´	                         "      ř                         "                            "                            "                            "      $)Ł                         "      üŚY                         " ˙˙    .   C@O:  OG5A;02  5B@>28G                    " ˙˙    h(.                         " ˙˙     z                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     z                         " ˙˙    |ş
                         " ˙˙     q                         " ˙˙    8                         " ˙˙    ť                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    |                         " ˙˙    ňj                          " ˙˙    ,   >@BO:   0VA0  $54>@V2=0                    " ˙˙                          " ˙˙    @T                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    @T                         " ˙˙    Ű                         " ˙˙     ż                         " ˙˙    _                         " ˙˙    đŇ                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    8Í                         " ˙˙    p                          " ˙˙    8   >HC;L:>  ;5:A0=4@  235=>28G                    " ˙˙                          " ˙˙    @ľd                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    @ľd                         " ˙˙    DŁ                         " ˙˙                              " ˙˙    č                          " ˙˙    °                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    |Â                         " ˙˙    ÄňR                          " ˙˙    2   5B@V2=89  5B@>  0E0@>28G                    " ˙˙    h(.                         " ˙˙    Âg                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Âg                         " ˙˙    \                         " ˙˙    @                         " ˙˙     	                         " ˙˙    `                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Đ                         " ˙˙    äń[                          " ˙˙    &    59:>  20=  20=>28G                    " ˙˙    h(.                         " ˙˙    ŕpr                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŕpr                         " ˙˙     	                         " ˙˙    đI                         " ˙˙    ř$                         " ˙˙    ČŻ                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ŕž                         " ˙˙     ˛d                          " ˙˙    2   &810  8E09;>  =0B>;V9>28G                    " ˙˙                          " ˙˙     z                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     z                         " ˙˙    ŔŚ                         " ˙˙     q                         " ˙˙    8                         " ˙˙    ť                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ŕ                         " ˙˙    @d                          " ˙˙    >   &V@C;L=V:>2  >;>48<8@  >@8A>28G                    " ˙˙                          " ˙˙     ˇ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     ˇ                         " ˙˙     z                         " ˙˙    Š                         " ˙˙    ŔÔ                         " ˙˙    @                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     !                         " ˙˙    `	                          " ˙˙    .   /:8<5=:>  V=0  "8E>=V2=0                    " ˙˙                          " ˙˙    @T                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    @T                         " ˙˙    Ű                         " ˙˙     ż                         " ˙˙    _                         " ˙˙    đŇ                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    8Í                         " ˙˙    p                          "         !B@>9:0  ?3B  '.                     "                           "      ÷s                         "                            "                            "                            "                            "      ÷s                         "      řÄ                         "      ŔQ                         "      ŕ(                         "       ˛                          "                            "                            "                            "      ¸ń                         "      Č_                          " ˙˙    8   830=>2  5==04V9  8:>;09>28G                    " ˙˙                          " ˙˙    ÷s                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ÷s                         " ˙˙    řÄ                         " ˙˙    ŔQ                         " ˙˙    ŕ(                         " ˙˙     ˛                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ¸ń                         " ˙˙    Č_                          "      &   "25@4>A?. 48A:. 20;:8                    "                           "      Č
                        "                            "                            "                            "      2Ž                         "      PÁ3                        "      ČY                        "      4[9                         "      (C                         "      5                         "                            "                            "       z                         "      8h:                        "      Yů                         " ˙˙    .   C@>2  20=  0;5=B8=>28G                    " ˙˙                          " ˙˙    `ţÍ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    `ţÍ                         " ˙˙    dÉ                         " ˙˙    °                         " ˙˙    X                         " ˙˙    h<                         " ˙˙                         " ˙˙                         " ˙˙    @KL                         " ˙˙    q                         " ˙˙    L\                          " ˙˙    D   >@10B5=:>  ;048A;02  >;>48<8@>28G                    " ˙˙                          " ˙˙                             " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                             " ˙˙    p                         " ˙˙    @                         " ˙˙                              " ˙˙    `ę                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    °                         " ˙˙    Đ}                          " ˙˙    2   0284>2  <8B@>  V:B>@>28G                    " ˙˙    h(.                         " ˙˙    ŕpr                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŕpr                         " ˙˙     	                         " ˙˙    đI                         " ˙˙    ř$                         " ˙˙    ČŻ                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ŕž                         " ˙˙     ˛d                          " ˙˙    <   C@02;L>2  ;5:A0=4@  @83>@>28G                    " ˙˙    h(.                         " ˙˙    ŕpr                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŕpr                         " ˙˙     	                         " ˙˙    đI                         " ˙˙    ř$                         " ˙˙    ČŻ                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ŕž                         " ˙˙     ˛d                          " ˙˙    8   V=G5=:>  @BC@  ;5:A0=4@>28G                    " ˙˙    h(.                         " ˙˙    ŕpr                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŕpr                         " ˙˙     	                         " ˙˙    đI                         " ˙˙    ř$                         " ˙˙    ČŻ                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ŕž                         " ˙˙     ˛d                          " ˙˙    B   025@8=AL:89  >;>48<8@  =4@V9>28G                    " ˙˙                          " ˙˙    a                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    a                         " ˙˙    |L                         " ˙˙    ŔÎ                         " ˙˙    `g                         " ˙˙     ×                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    <Z                         " ˙˙    Ds                          " ˙˙    4   0H8@V=  0;5@V9  V:B>@>28G                    " ˙˙                          " ˙˙     źž                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     źž                         " ˙˙                             " ˙˙    Đ                         " ˙˙    Hč                         " ˙˙    ř$                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    \r"                         " ˙˙    ÄI                          " ˙˙    :   0H8@V=  ;5:A0=4@  0;5@V9>28G                    " ˙˙    h(.                         " ˙˙     z                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     z                         " ˙˙    |ş
                         " ˙˙     q                         " ˙˙    8                         " ˙˙    ť                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    |                         " ˙˙    ňj                          " ˙˙    0   @C?:>  !5@3V9  8E09;>28G                    " ˙˙                          " ˙˙    @]Ć                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    @]Ć                         " ˙˙    řŽ                         " ˙˙     ÷                         " ˙˙    Đű                         " ˙˙    °0                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ó#                         " ˙˙    (˘                          " ˙˙    6   02;N4>2  5@<0=   02V;V9>28G                    " ˙˙    h(.                         " ˙˙    ŕpr                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŕpr                         " ˙˙     	                         " ˙˙    đI                         " ˙˙    ř$                         " ˙˙    ČŻ                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ŕž                         " ˙˙     ˛d                          " ˙˙    :   >@3C=>2  8E09;>  0;5=B8=>28G                    " ˙˙    h(.                         " ˙˙    ŕpr                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŕpr                         " ˙˙     	                         " ˙˙    đI                         " ˙˙    ř$                         " ˙˙    ČŻ                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ŕž                         " ˙˙     ˛d                          " ˙˙    >   0H8=AL:89  >;>48<8@  V:B>@>28G                    " ˙˙                          " ˙˙    (+                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     (                         " ˙˙    Č4S                         " ˙˙                             " ˙˙    ,Ş                         " ˙˙    ¤j                          " ˙˙    ź                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ¤Ź                         " ˙˙    $D                          " ˙˙    2   >?>2  0A8;L  OG5A;02>28G                    " ˙˙    h(.                         " ˙˙     z                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     z                         " ˙˙    |ş
                         " ˙˙     q                         " ˙˙    8                         " ˙˙    ť                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    |                         " ˙˙    ňj                          " ˙˙    .   >?>2  ;51  =0B>;V9>28G                    " ˙˙    h(.                         " ˙˙    ŕpr                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŕpr                         " ˙˙     	                         " ˙˙    đI                         " ˙˙    ř$                         " ˙˙    ČŻ                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ŕž                         " ˙˙     ˛d                          " ˙˙    0    0B8G  0A8;L  8:>;09>28G                    " ˙˙    h(.                         " ˙˙     z                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     z                         " ˙˙    |ş
                         " ˙˙     q                         " ˙˙    8                         " ˙˙    ť                          " ˙˙                         " ˙˙                         " ˙˙    ŔĆ-                         " ˙˙    <ć<                         " ˙˙    Ä+=                          " ˙˙    6   !84>@5=:>  5=8A  5==048528G                    " ˙˙                          " ˙˙     ˇ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     X*                         " ˙˙     sá                         " ˙˙    ´                          " ˙˙    L                         " ˙˙    XA                         " ˙˙    |Z                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ôˇ(                         " ˙˙    Ěť¸                          " ˙˙    0   !:0G:>  =4@V9    235=>28G                    " ˙˙                          " ˙˙     7                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     7                          " ˙˙    Ü*                         " ˙˙    P4                         " ˙˙    (                         " ˙˙    ö                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    lď                         " ˙˙    4H                          " ˙˙    2   "@CE0=>2  !5@3V9  048<>28G                    " ˙˙                          " ˙˙     .c                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    HŔ[                         " ˙˙    čîž                         " ˙˙    ř                         " ˙˙    źŃ                         " ˙˙    Źč                         " ˙˙    \%                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ź{"                         " ˙˙    ,s                          " ˙˙    *   (C<0=  ;;O  @>:>?>28G                    " ˙˙                          " ˙˙    `ő                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    `ő                         " ˙˙    ö                         " ˙˙    0ć                         " ˙˙    s                         " ˙˙    ¨Ţ                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ô-                         " ˙˙    lÇv                          " ˙˙    2   /:C=V=  0;5@V9  V:B>@>28G                    " ˙˙    h(.                         " ˙˙     z                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     z                         " ˙˙    |ş
                         " ˙˙     q                         " ˙˙    8                         " ˙˙    ť                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    |                         " ˙˙    ňj                          "         "5E>B45;                    "                           "      °Ę@                        "                            "                            "      Ŕjg                        "       ,                        "      a°                        "      DĐî                        "      Ô#I                         "      Z                          "      Ěđ                         "      ŹĽ:                         "                            "      ŔĆ-                         "      XŤÖ                        "      8śŮ                         " ˙˙    *   7V=  235=  5>=V4>28G                    " ˙˙                          " ˙˙     ćŞ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     ćŞ                         " ˙˙    ,ś                         " ˙˙     k                         " ˙˙    ľ                         " ˙˙                             " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ,Ý                         " ˙˙    Ô                          " ˙˙    <   5935;L7V<5@  <<0=CW;  .E8<>28G                    " ˙˙                          " ˙˙    `ă                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    `ă                         " ˙˙    S                         " ˙˙    0u                          " ˙˙    L                          " ˙˙    (#                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    4	                         " ˙˙    ,Ú                          " ˙˙    <   >3>D0;>20  ;5=0  >ABO=B8=V2=0                    " ˙˙                          " ˙˙    @ľd                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    @ľd                         " ˙˙    DŁ                         " ˙˙                              " ˙˙    č                          " ˙˙    °                          " ˙˙                         " ˙˙                         " ˙˙    ŔĆ-                         " ˙˙    <?                         " ˙˙    ,%                          " ˙˙    6   0E@0=T2  !5@3V9  8:>;09>28G                    " ˙˙    h(.                         " ˙˙     {                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     {                         " ˙˙    źň
                         " ˙˙    Đx                         " ˙˙    h<                         " ˙˙    Ř˝                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ěe                         " ˙˙    Ô2l                          " ˙˙    <   545@=8:>2  235=V9  8:>;09>28G                    " ˙˙    h(.                         " ˙˙    Ŕ0F                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ŕ0F                         " ˙˙    tG                         " ˙˙    `g                         " ˙˙    ŘY                          " ˙˙    Đk                          " ˙˙    ř.                         " ˙˙                         " ˙˙                         " ˙˙    tŁ                         " ˙˙    L0                          " ˙˙    <   @0G>20  !2VB;0=0  >ABO=B8=V2=0                    " ˙˙                          " ˙˙    ŔĆ-                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŔĆ-                         " ˙˙    §                         " ˙˙    `ę                          " ˙˙    :                          " ˙˙    PF                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ě                         " ˙˙    ôł%                          " ˙˙    2   <8B@VT2  235=  !5@3V9>28G                    " ˙˙                          " ˙˙    @X                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    @X                         " ˙˙    řÜ                         " ˙˙     Ĺ                         " ˙˙    Hq                          " ˙˙    đ                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    P                         " ˙˙    đäH                          " ˙˙    2   01V@0  V=0  ;5:A0=4@V2=0                    " ˙˙                          " ˙˙    Ŕw                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ŕw                         " ˙˙    Ü5                         " ˙˙    `a                         " ˙˙    °0                         " ˙˙    Đś                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ź~                         " ˙˙    a                          " ˙˙    6   5<;O=AL:89  =4@V9  3>@528G                    " ˙˙    h(.                         " ˙˙     {                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     {                         " ˙˙    źň
                         " ˙˙    Đx                         " ˙˙    h<                         " ˙˙    Ř˝                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ěe                         " ˙˙    Ô2l                          " ˙˙    4   >7;5=:>  5=8A  ;5:AV9>28G                    " ˙˙                          " ˙˙    đĹ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ô,	                         " ˙˙    Äň                         " ˙˙    ŕÖ                         " ˙˙    (                         " ˙˙    L                         " ˙˙    ň                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    X                         " ˙˙    ll                          " ˙˙    8   >;>:>;>2  ;5:AV9  V:B>@>28G                    " ˙˙                          " ˙˙    Ŕw                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Pł$                         " ˙˙    ¸                         " ˙˙    t                         " ˙˙    D                         " ˙˙    Ô                         " ˙˙    ď                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                             " ˙˙    t                          " ˙˙    6   >=>20;>2  .@V9  OG5A;02>28G                    " ˙˙                          " ˙˙     $                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     $                         " ˙˙    R                         " ˙˙    ť                          " ˙˙    ŕ.                          " ˙˙    @8                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     u                         " ˙˙    ŕ)                          " ˙˙    4   C1:V=  <8B@>  =0B>;V9>28G                    " ˙˙                          " ˙˙    Ŕw                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Pł$                         " ˙˙    ¸                         " ˙˙    t                         " ˙˙    D                         " ˙˙    Ô                         " ˙˙    ď                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                             " ˙˙    t                          " ˙˙    8   C7L<V=  ;5:A0=4@  V:B>@>28G                    " ˙˙                          " ˙˙    pŮ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    dĽ
                         " ˙˙    tä                         " ˙˙    řú                          " ˙˙    ř                         " ˙˙    üG                         " ˙˙    d^                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    P1)                         " ˙˙    $äş                          " ˙˙    ,   C:C9  0284  5=ECA>28G                    " ˙˙                          " ˙˙     PĂ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     l-                         " ˙˙     źđ                         " ˙˙    TĎ"                         " ˙˙    lĐ                         " ˙˙    hh                         " ˙˙    ěq                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    z+                         " ˙˙    BĹ                          " ˙˙    B   59:>2AL:89  >ABO=B8=  5=04V9>28G                    " ˙˙                          " ˙˙     ćŞ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     ćŞ                         " ˙˙    ,ś                         " ˙˙     k                         " ˙˙    ľ                         " ˙˙                             " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ,Ý                         " ˙˙    Ô                          " ˙˙    2   i =BCA  =4@i 9  8:>;09>28G                    " ˙˙                          " ˙˙    š*                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    š*                         " ˙˙    Ř5                         " ˙˙    ŔÚ                          " ˙˙    °6                          " ˙˙     A                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    č                         " ˙˙    0#                          " ˙˙    4   0:A0T2  02;>  =0B>;V9>28G                    " ˙˙    h(.                         " ˙˙    [                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    [                         " ˙˙    `b                         " ˙˙    ŔÔ                         " ˙˙    0u                          " ˙˙                               " ˙˙    $                         " ˙˙                         " ˙˙                         " ˙˙    Î                         " ˙˙    lż=                          " ˙˙    0   0<:V=0  N4<8;0  20=V2=0                    " ˙˙    h(.                         " ˙˙    `!`                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    `!`                         " ˙˙    °                         " ˙˙    0ě                         " ˙˙    {                          " ˙˙    ¨                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    
                         " ˙˙    ĚV                          " ˙˙    8   070@>2  <8B@>  ;5:A0=4@>28G                    " ˙˙                          " ˙˙    0W                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    K                         " ˙˙    @˘                         " ˙˙    h                         " ˙˙    ,~                          " ˙˙    ¤                          " ˙˙    ä%                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    X                         " ˙˙    $J                          " ˙˙    *   V:C;V=  3>@  .@V9>28G                    " ˙˙                          " ˙˙    ŔĎj                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŔĎj                         " ˙˙    čq                         " ˙˙    ŕ"                         " ˙˙    p                         " ˙˙    ¤                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    HJ                         " ˙˙    xW                          " ˙˙    (   =CE  <8B@>  ;53>28G                    " ˙˙                          " ˙˙    ŔŘ§                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŔŘ§                         " ˙˙    HE                         " ˙˙    `[                         " ˙˙    °­                         " ˙˙    Đ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    (P                         " ˙˙                              " ˙˙    4   @5E>2  VB0;V9  8:>;09>28G                    " ˙˙                          " ˙˙    @@                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    @@                         " ˙˙    ÄP	                         " ˙˙     H                         " ˙˙    R                          " ˙˙    pb                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    \M                         " ˙˙    äČ4                          " ˙˙    2   >?>2  04V<  >;>48<8@>28G                    " ˙˙                          " ˙˙    @ľd                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    @ľd                         " ˙˙    DŁ                         " ˙˙                              " ˙˙    č                          " ˙˙    °                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    |Â                         " ˙˙    ÄňR                          " ˙˙    ,   @O4:>  N1>2  0284V2=0                    " ˙˙                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ŕjg                        " ˙˙                         " ˙˙    Ŕjg                        " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                          " ˙˙    Ŕjg                         " ˙˙    6    54L:>  0;8=0  ;5:A0=4@V2=0                    " ˙˙                          " ˙˙    `;                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    `;                         " ˙˙    tŚ                         " ˙˙    °0                         " ˙˙    ,L                          " ˙˙    h[                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ¸~
                         " ˙˙    ¨1                          " ˙˙    <    C4AL:>9  !B0=i A;02  ;5:Ai 9>28G                    " ˙˙                          " ˙˙     |                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙     |                         " ˙˙    ¨.                         " ˙˙     î                         " ˙˙     w                         " ˙˙     á                          " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ¨t                         " ˙˙    Xx                          " ˙˙    .   "C=8:  ;53  =0B>;V9>28G                    " ˙˙                          " ˙˙    ŕÚ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    E                         " ˙˙    hŕĎ                         " ˙˙    P                         " ˙˙    t(                         " ˙˙                             " ˙˙    $?                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    đ%                         " ˙˙    xUŞ                          " ˙˙    .   &C?@C=  ;5:Ai 9  .@i 9>28G                    " ˙˙                          " ˙˙    Ŕáä                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ŕáä                         " ˙˙    ¨!                         " ˙˙    ŕ                         " ˙˙    đI                         " ˙˙    _                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    V)                         " ˙˙    ¸ť                          " ˙˙    0   (0@5FL  <8B@>  V:B>@>28G                    " ˙˙                          " ˙˙    ŔĎj                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŔĎj                         " ˙˙    čq                         " ˙˙    ŕ"                         " ˙˙    p                         " ˙˙    ¤                          " ˙˙    á                         " ˙˙                         " ˙˙                         " ˙˙    Ř+)                         " ˙˙    čŁA                          " ˙˙    ,   (0EBL>@  3>@  235=>28G                    " ˙˙                          " ˙˙    Ë¤                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    6%                         " ˙˙    Ę                         " ˙˙    ź5                         " ˙˙    `
                         " ˙˙    0                         " ˙˙    (6                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    t{$                         " ˙˙    Ľ                          "      :   -;5:B@>:0@48>3@0DK  -   ?>4@074.                     "                           "      `äx                        "                            "                            "                            "      ŕp'                         "      @U                         "      ¸3<                         "      ¤S                         "       )                         "      t                         "                            "                            "                            "      p0K                         "      Đ$U                         " ˙˙    8   01C@:V=  235=  ;5:A0=4@>28G                    " ˙˙                          " ˙˙    `ţÍ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    `ţÍ                         " ˙˙    dÉ                         " ˙˙    °                         " ˙˙    X                         " ˙˙    h<                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    Ô3%                         " ˙˙    Ę¨                          " ˙˙    ,   030=>2  .@V9  $54>@>28G                    " ˙˙                          " ˙˙     ćŞ                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ŕp'                         " ˙˙    ŕVŇ                         " ˙˙    Tj                         " ˙˙    ô4                         " ˙˙    H                         " ˙˙    C                         " ˙˙                         " ˙˙                         " ˙˙                         " ˙˙    ü%                         " ˙˙    DZŹ                       d   d   d   d    <     č                Ě Arial                              <     č                Ě Arial                              <     č                Ě Arial                              <     č                Ě Arial                              <     č                Ě Arial                              <     č                Ě Arial                                2   2   
   	 <     č    ´           Ě Arial îéńňâ ű (F7) ńîńňî˙íč˙ äë˙ î <     č   Č           Ě Arial                      čÍe <     č            ((( Ě Arial ˙˙                   Đ#g <     č           ˙˙˙ Ě Arial ˙˙                   ing  <     č           ((( Ě Arial ˙˙               htRange) <     č           ˙˙ŕ Ě Arial ˙˙               đóňęč    <     č               Ě Arial čňĺ F1  (F7)    1      čÍe  <     č    đ           Ě Arial čňĺ F1 ëĺííîăî ýëĺěĺíňŕ 
.av <     č   đ           Ě Arial Sub ShtBook_OnLoadEnd Su <     č    Č           Ě Arial ĐDb                   
   -    ŕ˙˙                              ŕ˙˙                               ŕ˙˙                                ˙˙˙                               
 ŕ˙˙                              	 ˙˙˙˙                              ˙˙˙˙                               ˙˙˙˙                              ˙˙ŕ                              ˙˙˙˙                             ˙˙ŕ                                  ÜÜÜ                              ÜÜÜ                                   ČČČ                                  ČČČ                              ˙˙˙                               ŔŔŔ                               ˙˙˙                              ˙˙˙                                  ŔŔŔ                              ČČČ                               ŔŔŔ                               ČČČ                               ŔŔŔ                                  ˙˙˙˙                               ˙˙˙˙                                ˙˙˙˙                               ˙˙˙                               ˙˙˙                                ŕ˙˙                                 ˙˙˙                                 ˙˙˙                               ( ŕ˙˙                                 % ŕŕŕ                              $ ŕŕŕ                                ˙˙ŕ                              ! ˙˙ŕ                              " ŕŕŕ                                 # ŕŕŕ                             & ŕŕŕ                               ' ˙˙ŕ                              ) Ŕ˙˙                                 * Żîî                                 + ˙˙˙˙                               , ŕ˙˙                                                       " ˙˙R o o t   E n t r y                                               ˙˙˙˙˙˙˙˙                               `ü?ôAźÉ7        C o n t e n t s                                                  ˙˙˙˙   ˙˙˙˙                                    8  äŁ      S u m m a r y I n f o r m a t i o n                           (  ˙˙˙˙˙˙˙˙˙˙˙˙                                        i                                                                          ˙˙˙˙˙˙˙˙˙˙˙˙                                                                	  
                                               !  "  #  $  %  &  '  (  )  *  +  ,  -  .  /  0  1  ţ˙˙˙ţ˙˙˙ý˙˙˙ý˙˙˙ý˙˙˙ţ˙˙˙ţ˙˙˙9  :  ;  <  =  >  ?  @  A  B  C  D  E  F  G  H  I  J  K  L  M  N  O  P  Q  R  S  T  U  V  W      ˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙                        	   
                                                                      !   "   #   $   %   &   '   (   )   *   +   ,   -   .   /   0   1   2   3   4   5   6   7   8   9   :   ;   <   =   >   ?   @   A   B   C   D   E   F   G   H   I   J   K   L   M   N   O   P   Q   R   S   T   U   V   W   X   Y   Z   [   \   ]   ^   _   `   a   b   c   d   e   f   g   h   i   j   k   l   m   n   o   p   q   r   s   t   u   v   w   x   y   z   {   |   }   ~                                                                                                          Ą   ˘   Ł   ¤   Ľ   Ś   §   ¨   Š   Ş   Ť   Ź   ­   Ž   Ż   °   ą   ˛   ł   ´   ľ   ś   ˇ   ¸   š   ş   ť   ź   ˝   ž   ż   Ŕ   Á   Â   Ă   Ä   Ĺ   Ć   Ç   Č   É   Ę   Ë   Ě   Í   Î   Ď   Đ   Ń   Ň   Ó   Ô   Ő   Ö   ×   Ř   Ů   Ú   Ű   Ü   Ý   Ţ   ß   ŕ   á   â   ă   ä   ĺ   ć   ç   č   é   ę   ë   ě   í   î   ď   đ   ń   ň   ó   ô   ő   ö   ÷   ř   ů   ú   ű   ü   ý   ţ   ˙                     ţ˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙ţ˙   ŕňůOhŤ +'łŮ   ŕňůOhŤ +'łŮ0   9        x            ¤      °      ź      Č      Ô   	   ŕ   
   ě      ř               $      ,     ă        Âĺäîěîńňü íŕ÷čńëĺíčé.ash                                                            458 @   PŻ]ţ  @   pA;ôAźÉ@   @kçôśýĆ      Ŕęöĺíň 7.0                                                                                                                                                                           ˙M'Option Explicit
'#include "definition.avb"

Const adCmdText = 1
Const adParamInput = 1

Const StartNumberOfRows = 15
Const StartNumberOfColumns = 2

Const MTC_PRIVILEGE_NAME = "0801"
Const MTC_WPF_NAME = "ÔÎŇ"

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
' âűáčđŕĺě ęîäű, çŕ čńęëţ÷ĺíčĺě ńďĺö-ęîäîâ "âűďëŕňŕ", "äîëă ńîňđóäíčęŕ" č "äîëă ďđĺäďđč˙ňč˙"
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
		SqlQr = "SELECT MTD_ID, MTD_CODE, 0 FROM MTD_CODES WHERE ( MTD_TYPE = 101 ) AND ( ISNULL( MTD_MODE, 0 ) <> 2 ) " & _
				vbNewLine & "UNION all " & vbNewLine & _
				"SELECT MTD_ID, MTD_CODE, 1 FROM MTD_CODES WHERE ( MTD_TYPE = 102 ) AND ( ISNULL( MTD_MODE, 0 ) <> 3 ) AND ( ISNULL( MTD_MODE, 0 ) <> 1 ) ORDER BY 3, 2"
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
		.Rows = StartNumberOfRows
		.Columns = 8
		' î÷čńňčëč ńňđîęó čňîăîâ
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

	' âűáđŕëč âńĺ ęîäű, ďî ęîňîđűě ĺńňü îáîđîňű
	For i = 0 To UBound( arData, 2 )
		DicCodes( arData( 4, i )) = 0
	Next

	With Sheet( 1 )
		If ( DicCodes.Count + 2 ) > 8 Then
			.Columns = 2 + DicCodes.Count
		End If

		TotalsCols( 0 ) = 2
		' Đčńóĺě řŕďęó ń ďĺđĺ÷íĺě ęîäîâ íŕ÷čńëĺíčé č óäĺđćŕíčé
		k = 3
		' arCodes( 2, i ) = 0 čëč 1. Íóćíî äë˙ ňîăî, ÷ňîá â óäĺđćŕíč˙ő ńäĺëŕňü ńäâčă íŕ îäčí ńňîëáĺö,
		' ÷ňîáű îńňŕâčňü ěĺńňî äë˙ ńîëáöŕ čňîăîâ ďî óäĺđćŕíč˙ě
		hdr = "1:Íŕ÷čńëĺíî"
		For i = 0 To UBound( arCodes, 2 )
			If i = 0 Then
			ElseIf ( arCodes( 2, i - 1 ) = 0 ) And ( arCodes( 2, i ) = 1 ) Then
				hdr = hdr & vbLf & "2:Âńĺăî" & vbLf & "1:Óäĺđćŕíî"
				TotalsCols( 1 ) = k
			End If
			If DicCodes.Exists( arCodes( 0, i )) Then
				DicCodes( arCodes( 0, i )) = k + arCodes( 2, i )
				hdr = hdr & vbLf & "2:" & arCodes( 1, i )
				k = k + 1
			End If
		Next
		hdr = hdr & vbLf & "2:Âńĺăî" & vbLf & "1:Ę âűäŕ÷ĺ"

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
		FOT = CalcDepsSumLimit_Pens( CodeID, Repcalc, max, False, d )
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
		FOT = CalcDepsSumLimit_Pens( CodeID, Repcalc, max, True,d )
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


		If Year(d) < 2009 Then
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

		If Month(d) = 1 And Year(d) = 2009 Then
			Calc2030 CodeID, Repcalc, d, max
		Else
			FOT = CalcDepsSumLimit( CodeID, Repcalc, d, max )
			sum1 = sum1 + Round2( FOT * percent / 100 , 2 )
			.Cell( 11, 3 ).value = FOT
			.Cell( 11, 5 ).value = Round2( FOT * percent / 100 , 2 )
		End If

'		.Cell( 4, 5 ).value = sum1
	End With
End Sub

'-----------------------------------------------------------------------------------------------
' Function GetCodeID( CodeName ) - âĺđíóňü čäĺíňčôčęŕňîđ ďî čěĺíč ęîäŕ/ěĺňîäŕ 
'-----------------------------------------------------------------------------------------------
Function GetCodeID( CodeName )
	Dim SqlQr, ar
	GetCodeID = 0
	SqlQr = "Select MTD_ID From MTD_CODES Where MTD_NAME = '" & CodeName & "'"
	Select Case App.AppType
	Case "DAO"
		If GetRows_DAO( SqlQr, ar ) Then GetCodeID = ar( 0, 0 )
	Case "OLEDB"
		If GetRows_ADO( SqlQr, adCmdText, ar ) Then GetCodeID = ar( 0, 0 )
	End Select
End Function

'________________________________________________
' Function QueryWithParams( SqlText, Options, arParams )
' Âűďîëíĺíčĺ çŕďđîńŕ ń ďĺđĺäŕ÷ĺé â íĺăî ďŕđŕěĺňđîâ
' Âîçâđŕůŕĺňń˙ îáúĺęň ňčďŕ RecordSet
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

	' óńňŕíŕâëčâŕĺě ďŕđŕěĺňđű çŕďđîńŕ
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
	ar = AR o o t   E n t r y                                               ˙˙˙˙˙˙˙˙                                @BźÉ]        C o n t e n t s                                                  ˙˙˙˙   ˙˙˙˙                                    ^  0k       S u m m a r y I n f o r m a t i o n                           (  ˙˙˙˙˙˙˙˙˙˙˙˙                                        i                                                                          ˙˙˙˙˙˙˙˙˙˙˙˙                                                ˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙ţ˙˙˙ý˙˙˙ý˙˙˙ý˙˙˙ţ˙˙˙ţ˙˙˙_  `  a  b  c  d  e  f  g  h  i  j  k  l  m  n  o  p  q  r  s  t  u  v  w  x  y  z  {  |  }      ˙˙˙˙˙˙˙˙                        	   
                                    ţ˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙˙