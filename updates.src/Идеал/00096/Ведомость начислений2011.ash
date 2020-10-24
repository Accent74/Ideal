–ѕа°±б                >  ю€	               0          >      ю€€€    1       €€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€Prm.Value = val1
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
	Dim TotalSum
	Dim AgSum
	Dim flAllowDogovor

	With Workarea.Code(CodeID)
		flAllowDogovor = Iif( ((.Mode And 8) = 0) Or (.Type <> 103), False, True )
	End With

	TotalSum = 0
	For i = 1 To Repcalc.Count
		If flAllowDogovor Or (Not Repcalc.Agent(i).Prop("“рудовой договор", d)) Then
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
Function CalcDepsSumLimit_Pens( Code, Repcalc, max, min, FirstDay, ByRef FOT_GPD)
	Dim i, CodeMode, IsAgInvalid 
	Dim TotalSum
	Dim AgSum, Ag
	Dim flAllowDogovor

	Fot_GPD = 0

	flAllowDogovor = ((Code.mode And 8) = 8) 'включать √ѕƒ
	CodeMode = (Code.mode And 7)
	IsAgInvalid = (CodeMode = 2)

	TotalSum = 0

	For i = 1 To Repcalc.Count
		Set Ag = Repcalc.Agent(i)

		If (Ag.Prop("»нвалид") = IsAgInvalid Or CodeMode = 0) And _
			Ag.Prop("“рудовой договор", FirstDay) = flAllowDogovor  Then

			AgSum = Repcalc.SumBy("CodeSumDependsC", i, Code.ID)

			If AgSum <> 0 Then
				If AgSum > max Then AgSum = max
				If AgSum < min Then AgSum = min
				TotalSum = TotalSum + AgSum
			End If

			If Ag.Prop("“рудовой договор", FirstDay) Then 
				Fot_GPD = Fot_GPD + AgSum
			End If
		End If
	Next

	CalcDepsSumLimit_Pens = TotalSum
End Function

<     и    †           ћ Arial D:\STDs.!!!\Zap\»деал\»деал       €€€€                 €€ 
 CSheetPageSheet1¬едомость начислений         61    7Н                                     ®   F   o   4   :   R   U   \   4   4   >   R   4   ?   ?   g   ]   :   :   @                                                                                   €€  CRow €€  CCell6     €€   (   54><>ABL  =0G8A;5=89 В                   А  " €€€€     В                 А  " €€€€     В                 А  " €€€€Б    "  " "  В                 А  " €€€€Б     "  " "  В                 А  " €€€€В     "  " "  В                 А  " €€€€Г     "  " "  В                 А А     €€      0@B  2 0 1 5   3.  В                   А  " €€€€     В                 А  " €€€€     В                 А А  ! €€       0G8A;5=8O В                 А  ! €€       %  В                 А  ! €€; Б    $" В                 А  ! €€< Б    В                 А  ! €€; Б 
   !C<<0 В                 А  ! €€< Б    В                 А  ! €€    
   515B В                 А  ! €€       @548B В                 А А            B>3>:  В                 А           В                 А =C5+C6+C7+C8"   Б †6С     В                 А  "   Б    В                 А =E5+E6+E7+E8"   Б THџ      В                 А  "   Б    В                 А  "        В                 А  "        В                 А А    €€       !  ( 107>2K9)  В                 А    €€*    а      В                 А  " €€ Б А®     В                 А  " €€/ Б    В                 А  " €€ Б иэd      В                 А  " €€/ Б    В                 А  " €€       9 2  В                 А  " €€       6 5 1  В                 А А    €€)       !  ( 8=20;84K)  В                 А    €€    ДH      В                 А  " €€ Б  О~     В                 А    €€ Б    В                 А  " €€ Б lJv      В                 А    €€ Б    В                 А  " €€       9 2  В                 А  " €€       6 5 1  В                 А А    €€       !  ( )  В                 А    €€    xK      Б                 А  " €€ Б    В                 А    €€ Б    В                 А  " €€ Б          В                 А  " €€ Б    В                 А  " €€         В                 А  " €€         В                 А А    €€        !  ( 1>;L=8G=K5)  В                 А    €€    а      Б                 А  " €€ Б    В                 А  " €€ Б    В                 А  " €€ Б          В                 А  " €€ Б    В                 А  " €€         В                 А  " €€         В                 А А    €€?       !  ( 45:@5B=K5)  В                 А    €€?    а      Б                 А  " €€? Б    В                 А  " €€? Б    В                 А  " €€? Б          В                 А  " €€? Б    В                 А  " €€?       6 5 3  В                 А  " €€?         В                   А А  ! €€       #45@60=8O В                 А  ! €€       %  В                 А  ! €€; Б 
   !C<<0 В                 А  " €€< Б    В                 А  ! €€    
   515B В                 А  ! €€       @548B В                 А А    €€@    
   $   В                 А    €€@      В                 А  " €€@ Б `у      В                 А    €€@ Б    В                 А  " €€@         В                 А  " €€@         В                 А А    €€       !  ( 107>2K9)  В                 А    €€    Ќћћћћћ@ В                 А  " €€ Б @г	      В                 А  " €€ Б    В                 А  " €€         В                 А  " €€         В                 А А    €€       !  ( 8=20;84K)  В                 А    €€    Ќћћћћћ@ В                 А  " €€* Б дҐ2      В                 А  " €€* Б    В                 А  " €€         В                 А  " €€         В                 А А    €€       !  ( )  В                 А    €€      В                 А  " €€€€Б          В                 А  " €€€€Б    В                 А  " €€         В                 А  " €€         В                 А А    €€>        !  ( 1>;L=8G=K5)  В                 А    €€?    Ќћћћћћ@ В                 А  " €€? Б          В                 А  " €€? Б    В                 А  " €€?         В                 А  " €€?         В                 А А    €€?       !  ( 45:@5B=K5)  В                 А    €€?      В                 А  " €€? Б          В                 А  " €€? Б    В                 А  " €€?       6 5 7  В                 А  " €€         В                 А А    €€    "   >5==K9  A1>@  1 , 5 %  В                 А    €€          ш? В                 А  " €€ Б №7      В                 А    €€ Б    В                 А  " €€       6 4 2  В                 А  " €€         В                   А А   €€$       >4@0745;5=85 В                 А   €€$        В                 А   €€# Б    0G8A;5=> В                 А   €€" Б    В                 А   €€" В    В                 А   €€# Б    #45@60=> В                 А   €€" Б    В                 А   €€" В    В                 А   €€" Г    В                 А   €€" Д    В                 А   €€# В      2K40G5 В                 А А   €€%       !>B@C4=8: В                 А   €€%      В                 А   €€#    "   A=>2=0O  70@?;0B0 В                 А   €€#       @5<8O В                 А   €€#    
   A53> В                 А   €€#       >5==K9  A1>@ В                 А   €€#       !  ( 107>2K9)  В                 А   €€#       !  ( 8=20;84K)  В                 А   €€#       $ В                 А   €€#    
   A53> В                 А   €€" А   В                 А А    €€      В                 А    €€             В                 А  "      †№.     В                 А  "       Zb     В                 А  "      †6С     В                 А  "      №7      В                 А  "      @г	      В                 А  "      дҐ2      В                 А  "      `у      В                 А  "      `ЎH     В                 А  "      @^H     В                 А А  "      $   ??0@0B  C?@02;5=8O В                 А  "        В                 А  "      †№.     В                 А  "       Zb     В                 А  "      †6С     В                 А  "      №7      В                 А  "      @г	      В                 А  "      дҐ2      В                 А  "      `у      В                 А  "      `ЎH     В                 А  "      @^H     В                 А А  " €€    .   OB;>2  02;>  0A8;L>28G В                 А  " €€       2 8 5 5 9 0 0 0 3 0  В                 А  " €€     Zb     В                 А  " €€     Zb     В                 А  " €€     іƒ     В                 А  " €€    АO      В                 А  " €€      В                 А  " €€     т+      В                 А  " €€    АГ∞      В                 А  " €€     ≈о      В                 А  " €€     п’     В                 А А  " €€    2   @0A>;  V:B>@  8:>;09>28G В                 А  " €€       1 9 4 8 1 1 2 9 1 6  В                 А  " €€    А®     В                 А  " €€      В                 А  " €€    А®     В                 А  " €€    ∞      В                 А  " €€    @г	      В                 А  " €€      В                 А  " €€    0Ј'      В                 А  " €€     є5      В                 А  " €€    `п№      В                 А А  " €€    8   ":0;5=:>  ;L30  ;5:A0=4@V2=0 В                 А  " €€       2 7 4 4 9 1 5 5 6 3  В                 А  " €€     Џє      В                 А  " €€      В                 А  " €€     Џє      В                 А  " €€    ђ…      В                 А  " €€      В                 А  " €€    д∞      В                 А  " €€    ∞я      В                 А  " €€    @Z$      В                 А  " €€    аХ      В                   А А    €€€€   :   CE30;B5@: _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _  В                                  <     и    †           ћ Arial                              <     и    †           ћ Arial                              <     и    †           ћ Arial                              <     и    †           ћ Arial                              <     и    †           ћ Arial                              <     и    †           ћ Arial                                2   2      	 <     и    і           ћ Arial ойств ы (F7) состо€ни€ дл€ о <     и   »           ћ Arial                      иЌe <     и    †       ((( ћ Arial €€              В     –#g <     и   †       €€€ ћ Arial €€              В     ing  <     и   †       ((( ћ Arial €€              В htRange) <     и   †       €€а ћ Arial €€              В рутки    <     и   †           ћ Arial ите F1  (F7)    1      иЌe  <     и    р           ћ Arial ите F1 ленного элемента 
.av <     и   р           ћ Arial Sub ShtBook_OnLoadEnd Su <     и    »           ћ Arial –Db                   
   
 <     и    †       јјј ћ Arial тод F1 мент списка 2 _ 43.adG    а€€                              а€€                               а€€                                €€€                               
 а€€                              	 €€€€                              €€€€                               €€€€                              €€а                              €€€€                             €€а                                  №№№                              №№№                                   »»»                                  »»»                              €€€                               јјј                               €€€                              €€€                                  јјј                              »»»                               јјј                               »»»                               јјј                                  €€€€                               €€€€                                €€€€                               €€€                               €€€                                а€€                                 €€€                                 €€€                               / €€€€                           јјј * €€€€                               ) €€€€                             ( а€€                                 % ааа                              $ ааа                                €€а                              ! €€а                              " ааа                                 # ааа                             & ааа                               ' €€а                              + жжъ                             , жжъ                              - жжъ                              . €€€€                              < №№№            ddd    ddd    ddd ; №№№    ddd    ddd            ddd : №№№                              9 №№№                              8 “““                              7 ‘‘‘                              6 ‘‘‘                              5 ’’’                              4 ’’’                              3 ÷÷÷                             2 ÷÷÷                              1 ÷÷÷                              0 жжъ                                 = €€€€   јјј    јјј    јјј    јјј > €€€    (((    (((    (((    ((( ? €€€€   (((    (((    (((    ((( F €€€                              E €€€                               @ €€€€   ddd    ddd    ddd    ddd A €€€€                              B €€€€                             C а€€                               D а€€                                                      и   †       ((( ћ Arial €€              В htRange) <     и   †       €€а ћ Arial €€              В рутки    <     и   †           ћ Arial ите F1  (F7)    1      иЌe  <     и    р           ћ Arial ите F1 ленного элемента 
.av <     и   р           ћ Arial Sub ShtBook_OnLoadEnd Su <     и    »           ћ Arial –DR o o t   E n t r y                                               €€€€€€€€                                iс•–#   А      C o n t e n t s                                                  €€€€   €€€€                                    $   ЙМ       S u m m a r y I n f o r m a t i o n                           (  €€€€€€€€€€€€                                        m                                                                          €€€€€€€€€€€€                                                €€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€                        	   
                                                               ю€€€ю€€€э€€€э€€€ю€€€ю€€€%   &   '   (   )   *   +   ,   -   .   /   2   €€€€€€€€3   4   5   6   7   8   9   :   ;   <   =   U   €€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€V   W   X   Y   Z   [   \   ]   ^   _   `   a   b   c   d       €€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€               ю€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€ю€   аЕЯтщOhЂС +'≥ў   аЕЯтщOhЂС +'≥ў0   =        x      А      ®      і      ј      ћ      Ў   	   д   
   р      ь               А(      0     г        ¬едомость начислений2011.ash                                                            817 @    ТЇЛ  @   †∆шhс•–@   @kзфґэ∆      јкцент 7.0                                                                                                                                                                       €\ZOption Explicit
'#include "definition.avb"
'#include "ado.inc"

Const adInteger          =   3 ' vbLong
Const adVarChar          = 200
Const adParamInput = 1
Const adUseServer        = 2

Const StartNumberOfRows = 22
Const LastRow = 12
Const StartNumberOfColumns = 2

Const MTC_PRIVILEGE_NAME = "0801"
Const MTC_WPF_NAME = "‘ќ“"

Dim MTD_LGT_ID
Dim cRow

Dim DicCodes
Dim TotalsCols( 2 )
Dim arCodes
Dim PrivilegeID
Dim BuhName

Const MTD_NDFL_NAME 		= "2400"
Const MTD_ESV_NAME 			= "2410"
Const MTD_ESV_INV_NAME 	= "2460"
Const MTD_ESVB_NAME			= "2420"
Const MTD_ESVT_NAME 		= "2430"
Const MTD_ESVD_NAME 		= "2450"		'декретные
Const MTD_ARMY_NAME 		= "2500"		'военный сбор 1,5%

Dim MTD_NDFL_ID 	
Dim MTD_ESV_ID 	
Dim MTD_ESV_INV_ID
Dim MTD_ESVB_ID	
Dim MTD_ESVT_ID 	
Dim MTD_ESVD_ID 	
Dim MTD_ARMY_ID 	

Dim MTD_HOLYDAY_NO_PAY

'-----------------------------------------------------------------------------------------------
Sub ShtBook_OnLoad
	Dim ut

	Set ut = Workarea.UTable(USRTBL_ONE)

	BuhName =ut.GetValue(1, "‘»ќ бухгалтера", 2)
	PrivilegeID = WorkArea.GetCodeID(MTC_PRIVILEGE_NAME)

	MTD_HOLYDAY_NO_PAY = workarea.getcodeid("8001")
	MTD_LGT_ID = Workarea.GetCodeID("2440")

	Set DicCodes = CreateObject("Scripting.Dictionary")
	arCodes = GetCodes

	With Sheet( 1 )
		.DisplayGrid = False
		.DisplayHeadings = True
	End With

	EnableButtons 1
	PeriodText = WorkArea.Period.RepTitle

	MTD_NDFL_ID 		= Array(Workarea.GetCodeID(MTD_NDFL_NAME), Workarea.GetCodeID(MTD_NDFL_NAME & "а"))
	MTD_ESV_ID 			= Array(Workarea.GetCodeID(MTD_ESV_NAME), Workarea.GetCodeID(MTD_ESV_NAME & "а"))
	MTD_ESV_INV_ID		= Array(Workarea.GetCodeID(MTD_ESV_INV_NAME), Workarea.GetCodeID(MTD_ESV_INV_NAME & "а"))
	MTD_ESVB_ID 		= Array(Workarea.GetCodeID(MTD_ESVB_NAME), Workarea.GetCodeID(MTD_ESVB_NAME & "а"))
	MTD_ESVT_ID 		= Array(Workarea.GetCodeID(MTD_ESVT_NAME), Workarea.GetCodeID(MTD_ESVT_NAME & "а"))
	MTD_ESVD_ID 		= Array(Workarea.GetCodeID(MTD_ESVD_NAME), Workarea.GetCodeID(MTD_ESVD_NAME & "а"))
	MTD_ARMY_ID 		= Array(Workarea.GetCodeID(MTD_ARMY_NAME), Workarea.GetCodeID(MTD_ARMY_NAME & "а"))

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
' выбираем коды, за исключением спец-кодов "выплата", "долг сотрудника" и "долг предпри€ти€"
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

'---
'
'---
Sub InitGrid
	Dim i

	With Sheet( 1 )
		.Rows = StartNumberOfRows - 3
		.Rows = .Rows + 3

		.Cell(StartNumberOfRows - 2, 1).Value = "ѕодразделение"
		.Cell(StartNumberOfRows - 1, 1).Value = "—отрудник"
		.Cell(StartNumberOfRows - 2, 2).Value = "»ЌЌ"
'		.Cell(StartNumberOfRows - 1, 2).Value = "льготы"

		.Range(1, StartNumberOfRows - 2, 1, StartNumberOfRows - 1).SetBorder acBrdOutline	
		.Range(2, StartNumberOfRows - 2, 2, StartNumberOfRows - 1).SetBorder acBrdOutline	
		.Range(1, StartNumberOfRows - 2, 2, StartNumberOfRows - 1).Backcolor StdBackColor(1)
		
		.Range(1, StartNumberOfRows, 2, StartNumberOfRows).Backcolor StdBackColor(2)
		.Range(1, StartNumberOfRows, 2, StartNumberOfRows).SetBorder acBrdGrid

		.Columns = 8

		' очистили строку итогов
		For i = 2 To .Columns
			.Cell( StartNumberOfRows, i ).Value = 0.0
		Next
	End With

End Sub
'---
'
'---
Sub BuildMainReport
	Dim GroupID, AgID, PeriodID, k, i, n
	Dim hdr, RecSet, arData, Flag

	DicCodes.RemoveAll

	cRow = StartNumberOfRows
	InitGrid

	PeriodID = WorkArea.Period.ID
	arData = GetData( PeriodID )

	If UBound(arData) = -1 Then Exit Sub

	' выбрали все коды, по которым есть обороты
	For i = 0 To UBound( arData, 2 )
		If arData(4, i) <> MTD_HOLYDAY_NO_PAY Then
			DicCodes( arData( 4, i )) = 0
		End If
	Next

	With Sheet( 1 )
		If ( DicCodes.Count + 2 ) > 8 Then
			.Columns = 2 + DicCodes.Count
		End If

		TotalsCols( 0 ) = 2
		' –исуем шапку с перечнем кодов начислений и удержаний
		k = 3
		' arCodes( 2, i ) = 0 или 1. Ќужно дл€ того, чтоб в удержани€х сделать сдвиг на один столбец,
		' чтобы оставить место дл€ солбца итогов по удержани€м
		hdr = "1:Ќачислено"

		For i = 0 To UBound( arCodes, 2 )
'			If arData(4, i) <> MTD_HOLYDAY_NO_PAY Then
				If i = 0 Then
				ElseIf ( arCodes( 2, i - 1 ) = 0 ) And ( arCodes( 2, i ) = 1 ) Then
					hdr = hdr & vbLf & "2:¬сего" & vbLf & "1:”держано"
					TotalsCols( 1 ) = k
				End If
		R o o t   E n t r y                                               €€€€€€€€                               ІDт•–?   А      C o n t e n t s                                                  €€€€   €€€€                                    @   ЧМ       S u m m a r y I n f o r m a t i o n                           (  €€€€€€€€€€€€                                        m                                                                          €€€€€€€€€€€€                                                                        	   
                                                               ю€€€€€€€э€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€ю€€€э€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€ю€€€ю€€€A   B   C   D   E   F   G   H   I   J   K   L   M   N   O   P   Q   R   S   T   e   €€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€f   g   h   i   j   k   l   m   n   o   p   q   r   s   t   u   v   w       €€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€		If DicCodes.Exists( arCodes( 0, i )) Then				
					DicCodes( arCodes( 0, i )) = k + arCodes( 2, i )
					hdr = hdr & vbLf & "2:" & arCodes( 1, i )
					k = k + 1
				End If
'			End If
		Next

		hdr = hdr & vbLf & "2:¬сего" & vbLf & "1:  выдаче"
		.MakeHeader StartNumberOfRows - 2, 3, hdr
		.columns = k + 2

		If k = 5 Then 
			TotalsCols( 2 ) = .Columns - 2 
		Else 
			TotalsCols( 2 ) = .Columns - 1
		End If

		Flag = False

		For i = 0 To UBound( arData, 2 )
			If GroupID <> arData( 0, i ) Then
				cRow = .InsertRow2

				.Cell( cRow, 1 ).CellParam = 1
				.Range( 1, cRow, .Columns, cRow ).Font.Bold = True

				.Cell(cRow, 1).Value = arData( 2, i )

				If DicCodes.Exists(arData( 4, i )) Then
					.Cell( StartNumberOfRows, DicCodes( arData( 4, i ))).Value = 0. 'arData( 6, i )
				End If

				GroupID = arData( 0, i )
			End If

			If AgID <> arData( 1, i ) Then
				cRow = .InsertRow2

				.Cell( cRow, 1 ).Value = arData( 3, i )
				.Cell( cRow, 2 ).Value = arData( 7, i) 'GetPrivilege( PrivilegeID, arData( 1, i ), PeriodID )
				AgID = arData( 1, i )

			 	Flag = (arData(4, i) = MTD_HOLYDAY_NO_PAY)

				If Flag Then
					With Sheet1.Range(3, cRow, Sheet1.Columns, cRow)
						.MergeCells = True
						.Value = "ѕриказ на неоплачиваемый отпуск"
						.Alignment acLeft
						.setborder acBrdOutline	
					End With
				End If
			
			End If

			If Not Flag Then
				With .Cell( cRow, DicCodes( arData( 4, i )))
					.Value = arData( 5, i )
					.CellParam = arData( 4, i )

					If MTD_LGT_ID = arData( 4, i ) Then 
						.ToolTip = "Ћьгота по Ќƒ‘Ћ не учитываетс€ в итогах"
					End If
				End With
			End If

'			.Cell( StartNumberOfRows, DicCodes( arData( 4, i )) ).Value = .Cell( StartNumberOfRows, DicCodes( arData( 4, i ))).Value + arData( 6, i )

		Next

		CalcTotals
		.Range( 1, StartNumberOfRows - 2, TotalsCols(2)+1, .Rows).Autofit
		.Range( 1, StartNumberOfRows + 1, TotalsCols(2)+1, .Rows).Stripe
		.Range( 1, StartNumberOfRows + 1, TotalsCols(2)+1, .Rows).Alignment = acRight
		.Range( 1, StartNumberOfRows + 1, TotalsCols(2)+1, .Rows).SetBorder acBrdOutline + acBrdInsideVert, acLineNormal
	
		With .Range(1, StartNumberOfRows - 2 , Sheet(1).Columns, StartNumberOfRows - 1)
			.Multiline = True
			.Alignment = acCenter
		End With

'		With .Range( 3, StartNumberOfRows, .Columns, StartNumberOfRows )
		With .Range( 3, StartNumberOfRows, TotalsCols(2)+1, StartNumberOfRows )
			.Font.Bold = True
			.Alignment = acRight
			.BackColor = 14745599
			.SetBorder acBrdGrid
		End With

		cRow = .InsertRow2
		cRow = .InsertRow2
		.Cell(cRow,1).Value = "Ѕухгалтер:___________________"
  		'.Range (2,cRow,6,cRow).MergeCells = True
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
						"sum(iif(JP_SUM <>0, JP_Tariff,0)), AGENTS.AG_VATNO " & _
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
			       "                  And MTD_TYPE <> 103 " & _
			       "Group By AG_TREE.P0, P_JOURNAL.AG_ID, Groups.AG_NAME, AGENTS.AG_NAME, AGENTS.AG_VATNO, MTC_ID " & _
			       "Having Sum(JP_SUM) <> 0 or MTC_id = " & MTD_HOLYDAY_NO_PAY & " " & _
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
 		       	 "                  And MTD_TYPE <> 103 " & _
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
				.Sql = "PARAMETERS MtcI               ю€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€ю€   аЕЯтщOhЂС +'≥ў   аЕЯтщOhЂС +'≥ў0   =        x      А      ®      і      ј      ћ      Ў   	   д   
   р      ь               А(      0     г        ¬едомость начислений2011.ash                                                            818 @   ∞JэѓЛ  @   P;т•–@   @kзфґэ∆      јкцент 7.0                                                                                                                                                                       €XZOption Explicit
'#include "definition.avb"
'#include "ado.inc"

Const adInteger          =   3 ' vbLong
Const adVarChar          = 200
Const adParamInput = 1
Const adUseServer        = 2

Const StartNumberOfRows = 22
Const LastRow = 12
Const StartNumberOfColumns = 2

Const MTC_PRIVILEGE_NAME = "0801"
Const MTC_WPF_NAME = "‘ќ“"

Dim MTD_LGT_ID
Dim cRow

Dim DicCodes
Dim TotalsCols( 2 )
Dim arCodes
Dim PrivilegeID
Dim BuhName

Const MTD_NDFL_NAME 		= "2400"
Const MTD_ESV_NAME 			= "2410"
Const MTD_ESV_INV_NAME 	= "2460"
Const MTD_ESVB_NAME			= "2420"
Const MTD_ESVT_NAME 		= "2430"
Const MTD_ESVD_NAME 		= "2450"		'декретные
Const MTD_ARMY_NAME 		= "2500"		'военный сбор 1,5%

Dim MTD_NDFL_ID 	
Dim MTD_ESV_ID 	
Dim MTD_ESV_INV_ID
Dim MTD_ESVB_ID	
Dim MTD_ESVT_ID 	
Dim MTD_ESVD_ID 	
Dim MTD_ARMY_ID 	

Dim MTD_HOLYDAY_NO_PAY

'-----------------------------------------------------------------------------------------------
Sub ShtBook_OnLoad
	Dim ut

	Set ut = Workarea.UTable(USRTBL_ONE)

	BuhName =ut.GetValue(1, "‘»ќ бухгалтера", 2)
	PrivilegeID = WorkArea.GetCodeID(MTC_PRIVILEGE_NAME)

	MTD_HOLYDAY_NO_PAY = workarea.getcodeid("8001")
	MTD_LGT_ID = Workarea.GetCodeID("2440")

	Set DicCodes = CreateObject("Scripting.Dictionary")
	arCodes = GetCodes

	With Sheet( 1 )
		.DisplayGrid = False
		.DisplayHeadings = True
	End With

	EnableButtons 1
	PeriodText = WorkArea.Period.RepTitle

	MTD_NDFL_ID 		= Array(Workarea.GetCodeID(MTD_NDFL_NAME), Workarea.GetCodeID(MTD_NDFL_NAME & "а"))
	MTD_ESV_ID 			= Array(Workarea.GetCodeID(MTD_ESV_NAME), Workarea.GetCodeID(MTD_ESV_NAME & "а"))
	MTD_ESV_INV_ID		= Array(Workarea.GetCodeID(MTD_ESV_INV_NAME), Workarea.GetCodeID(MTD_ESV_INV_NAME & "а"))
	MTD_ESVB_ID 		= Array(Workarea.GetCodeID(MTD_ESVB_NAME), Workarea.GetCodeID(MTD_ESVB_NAME & "а"))
	MTD_ESVT_ID 		= Array(Workarea.GetCodeID(MTD_ESVT_NAME), Workarea.GetCodeID(MTD_ESVT_NAME & "а"))
	MTD_ESVD_ID 		= Array(Workarea.GetCodeID(MTD_ESVD_NAME), Workarea.GetCodeID(MTD_ESVD_NAME & "а"))
	MTD_ARMY_ID 		= Array(Workarea.GetCodeID(MTD_ARMY_NAME), Workarea.GetCodeID(MTD_ARMY_NAME & "а"))

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
' выбираем коды, за исключением спец-кодов "выплата", "долг сотрудника" и "долг предпри€ти€"
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

'---
'
'---
Sub InitGrid
	Dim i

	With Sheet( 1 )
		.Rows = StartNumberOfRows - 3
		.Rows = .Rows + 3

		.Cell(StartNumberOfRows - 2, 1).Value = "ѕодразделение"
		.Cell(StartNumberOfRows - 1, 1).Value = "—отрудник"
		.Cell(StartNumberOfRows - 2, 2).Value = "»ЌЌ"
'		.Cell(StartNumberOfRows - 1, 2).Value = "льготы"

		.Range(1, StartNumberOfRows - 2, 1, StartNumberOfRows - 1).SetBorder acBrdOutline	
		.Range(2, StartNumberOfRows - 2, 2, StartNumberOfRows - 1).SetBorder acBrdOutline	
		.Range(1, StartNumberOfRows - 2, 2, StartNumberOfRows - 1).Backcolor StdBackColor(1)
		
		.Range(1, StartNumberOfRows, 2, StartNumberOfRows).Backcolor StdBackColor(2)
		.Range(1, StartNumberOfRows, 2, StartNumberOfRows).SetBorder acBrdGrid

		.Columns = 8

		' очистили строку итогов
		For i = 2 To .Columns
			.Cell( StartNumberOfRows, i ).Value = 0.0
		Next
	End With

End Sub
'---
'
'---
Sub BuildMainReport
	Dim GroupID, AgID, PeriodID, k, i, n
	Dim hdr, RecSet, arData, Flag

	DicCodes.RemoveAll

	cRow = StartNumberOfRows
	InitGrid

	PeriodID = WorkArea.Period.ID
	arData = GetData( PeriodID )

	If UBound(arData) = -1 Then Exit Sub

	' выбрали все коды, по которым есть обороты
	For i = 0 To UBound( arData, 2 )
		If arData(4, i) <> MTD_HOLYDAY_NO_PAY Then
			DicCodes( arData( 4, i )) = 0
		End If
	Next

	With Sheet( 1 )
		If ( DicCodes.Count + 2 ) > 8 Then
			.Columns = 2 + DicCodes.Count
		End If

		TotalsCols( 0 ) = 2
		' –исуем шапку с перечнем кодов начислений и удержаний
		k = 3
		' arCodes( 2, i ) = 0 или 1. Ќужно дл€ того, чтоб в удержани€х сделать сдвиг на один столбец,
		' чтобы оставить место дл€ солбца итогов по удержани€м
		hdr = "1:Ќачислено"

		For i = 0 To UBound( arCodes, 2 )
'			If arData(4, i) <> MTD_HOLYDAY_NO_PAY Then
				If i = 0 Then
				ElseIf ( arCodes( 2, i - 1 ) = 0 ) And ( arCodes( 2, i ) = 1 ) Then
					hdr = hdr & vbLf & "2:¬сего" & vbLf & "1:”держано"
					TotalsCols( 1 ) = k
				End If
				If DicCodes.Exists( arCodes( 0, i )) Then				
					DicCodes( arCodes( 0, i )) = k + arCodes( 2, i )
					hdr = hdr & vbLf & "2:" & arCodes( 1, i )
					k = k + 1
				End If
'			End If
		Next

		hdr = hdr & vbLf & "2:¬сего" & vbLf & "1:  выдаче"
		.MakeHeader StartNumberOfRows - 2, 3, hdr
		.columns = k + 2

		If k = 5 Then 
			TotalsCols( 2 ) = .Columns - 2 
		Else 
			TotalsCols( 2 ) = .Columns - 1
		End If

		Flag = False

		For i = 0 To UBound( arData, 2 )
			If GroupID <> arData( 0, i ) Then
				cRow = .InsertRow2

				.Cell( cRow, 1 ).CellParam = 1
				.Range( 1, cRow, .Columns, cRow ).Font.Bold = True

				.Cell(cRow, 1).Value = arData( 2, i )

				If DicCodes.Exists(arData( 4, i )) Then
					.Cell( StartNumberOfRows, DicCodes( arData( 4, i ))).Value = 0. 'arData( 6, i )
				End If

				GroupID = arData( 0, i )
			End If

			If AgID <> arData( 1, i ) Then
				cRow = .InsertRow2

				.Cell( cRow, 1 ).Value = arData( 3, i )
				.Cell( cRow, 2 ).Value = arData( 7, i) 'GetPrivilege( PrivilegeID, arData( 1, i ), PeriodID )
				AgID = arData( 1, i )

			 	Flag = (arData(4, i) = MTD_HOLYDAY_NO_PAY)

				If Flag Then
					With Sheet1.Range(3, cRow, Sheet1.Columns, cRow)
						.MergeCells = True
						.Value = "ѕриказ на неоплачиваемый отпуск"
						.Alignment acLeft
						.setborder acBrdOutline	
					End With
				End If
			
			End If

			If Not Flag Then
				With .Cell( cRow, DicCodes( arData( 4, i )))
					.Value = arData( 5, i )
					.CellParam = arData( 4, i )

					If MTD_LGT_ID = arData( 4, i ) Then 
						.ToolTip = "Ћьгота по Ќƒ‘Ћ не учитываетс€ в итогах"
					End If
				End With
			End If

'			.Cell( StartNumberOfRows, DicCodes( arData( 4, i )) ).Value = .Cell( StartNumberOfRows, DicCodes( arData( 4, i ))).Value + arData( 6, i )

		Next

		CalcTotals
		.Range( 1, StartNumberOfRows - 2, TotalsCols(2)+1, .Rows).Autofit
		.Range( 1, StartNumberOfRows + 1, TotalsCols(2)+1, .Rows).Stripe
		.Range( 1, StartNumberOfRows + 1, TotalsCols(2)+1, .Rows).Alignment = acRight
		.Range( 1, StartNumberOfRows + 1, TotalsCols(2)+1, .Rows).SetBorder acBrdOutline + acBrdInsideVert, acLineNormal
	
		With .Range(1, StartNumberOfRows - 2 , Sheet(1).Columns, StartNumberOfRows - 1)
			.Multiline = True
			.Alignment = acCenter
		End With

'		With .Range( 3, StartNumberOfRows, .Columns, StartNumberOfRows )
		With .Range( 3, StartNumberOfRows, TotalsCols(2)+1, StartNumberOfRows )
			.Font.Bold = True
			.Alignment = acRight
			.BackColor = 14745599
			.SetBorder acBrdGrid
		End With

		cRow = .InsertRow2
		cRow = .InsertRow2
		.Cell(cRow,1).Value = "Ѕухгалтер:___________________"
  		'.Range (2,cRow,6,cRow).MergeCells = True
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
						"sum(iif(JP_SUM <>0, JP_Tariff,0)), AGENTS.AG_VATNO " & _
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
			       "                  And MTD_TYPE <> 103 " & _
			       "Group By AG_TREE.P0, P_JOURNAL.AG_ID, Groups.AG_NAME, AGENTS.AG_NAME, AGENTS.AG_VATNO, MTC_ID " & _
			       "Having Sum(JP_SUM) <> 0 or MTC_id = " & MTD_HOLYDAY_NO_PAY & " " & _
			       "Order By Groups.AG_NAME, AGENTS.AG_NAME "

			.Parameters(0).Value = PeriodID

			Set Rs = .OpenRecordSet
			If Not rs.eof Then
				Rs.MoveLast
				n = Rs.RecordCount
				Rs.MoveFirD Long, AgID Long, PeriodID Long; " & _
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

	SetZero Totals		' итоги 3
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
				If Sheet( 1 ).Cell(i, j).cellParam <> MTD_LGT_ID Then
					If  IsNumeric(Sheet( 1 ).Cell(i, j).Value) Then
						GroupTotals(j) = GroupTotals(j) + Sheet( 1 ).Cell(i, j).Value
					End If
				End If
			Next
		End If
	Next

	For j = StartNumberOfColumns + 1 To Sheet(1).Columns	' итоги 3
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
					If .Cell(i, j).cellParam <> MTD_LGT_ID And IsNumeric(.Cell( i, j ).Value) Then 
						Sum = Sum + .Cell( i, j ).Value
					End If
				Next

				.Cell( i, TotalsCols( k )).Value = Sum
			Next
			.Cell(i, TotalsCols(2)+1).Value = .Cell(i, TotalsCols(1)).Value - .Cell( i, TotalsCols(2)).Value

'			.Cell( i, .Columns ).Value = .Cell( i, TotalsCols( 1 ) ).Value - .Cell( i, TotalsCols( 2 )).Value
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

	'If LastPos >= FirstPos Then Sheet( 1 ).Range(FirstPos, OnRow, LastPos, OnRow).MergeCells = True
End Sub
'________________________________________________

Sub display_code_calc(Code, Row, FOT_Sum, FOT_Percent)
	Sheet(1).Cell(Row, 7).Value = Code.DbCode
	Sheet(1).Cell(Row, 8).Value = Code.CrCode
	Sheet(1).Cell(Row, 2).Value = Code.percent
	Sheet(1).Cell(Row, 3).value = FOT_Sum
	Sheet(1).Cell(Row, 5).value = FOT_Percent
End Sub
'________________________________________________
Sub BuildSecondaryReport
	Dim max, percent, i, CodeID, FOT, FOT_GPD, sum1, sumT
	Dim arFonds, Repcalc, BaseSum
	Dim Code, FirstDay, min

	arFonds = Array( "2000", "2001", "2003", "2002", "2004")

	Set Repcalc = WorkArea.CreateIntObject("CommonCalc")
	Repcalc.setPeriod "month", Workarea.Period.Month, Workarea.Period.Year
	FirstDay = Workarea.Period.FirstDay

	max = Workarea.UTable( USRTBL_LGT ).GetValue( 1, USRTBL_MAXCY, 2 )
	min = max / 17

	percent = Workarea.UTable( USRTBL_LGT ).GetValue( 1, USRTBL_ECB, 2 )
stop
	For i = 0 To UBound(arFonds)
		CodeID = WorkArea.GetCodeID( arFonds(i))

		If CodeID <> 0 Then
			Set Code = WorkArea.Code(CodeID)
			percent = Code.Percent

			FOT = repcalc.totalsumby("CodeSumW", CodeID)
			BaseSum = GetTariffByMtcPP(CodeID, Workarea.Period.ID)

'			display_code_calc Code, 5 + i, Round2( FOT * 100 / percent , 2), FOT
			display_code_calc Code, 5 + i, BaseSum, FOT
			Repcalc.Clear

		End If
	Next
	
End Sub
'________________________________________________
Sub BuildLastReport
	Dim percent, i, sum1, sumT, CodeID
	Dim arFonds

	sum1 = 0

	arFonds = Array( MTD_NDFL_ID, MTD_ESV_ID, MTD_ESV_INV_ID, MTD_ESVB_ID, MTD_ESVT_ID, MTD_ESVD_ID, MTD_ARMY_ID)
	percent  = Array( 15, 3.6, 3.6, 2, 2.6, 2, 1.5)

	With Sheet( 1 )
		For i = 0 To UBound(arFonds)
			GetSumByCode arFonds(i), sum1, sumT
			cRow = LastRow + i
			.Cell( cRow, 2 ).Value = percent(i)
			.Cell( cRow, 3 ).value = round2(sum1, 2)
			With WorkArea.Code( arFonds(i)(0) )
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
		sum1 = rs.fields("PSUM").Value
		sumT = rs.fields("pt").Value
	End If
End Sub
'=================
Sub GetSumByCode(mtdID, sum1, sumT)
	Dim rs1, rs2

	InitRs rs1, _
		"SELECT SUM(JP_SUM) AS PSUM, sum(jp_tariff) as pt FROM P_JOURNAL WHERE JP_DONE = 1 " & _
			" AND C_PERIOD = " & WorkArea.Period.ID & _
			" AND MTC_ID = " & mtdID(0)

	InitRs rs2, _
		"SELECT SUM(JP_SUM) AS PSUM, sum(jp_tariff) as pt FROM P_JOURNAL WHERE JP_DONE = 1 " & _
			" AND C_PERIOD = " & WorkArea.Period.ID & _
			" AND MTC_ID = " & mtdID(1)

	If Not rs1.eof	Then
		sum1 = checknull(rs1.fields("PSUM").Value, 0)
		sumT = checknull(rs1.fields("pt").Value, 0)
	Else
		Sum1 = 0
	End If

	If Not rs2.eof	Then
		sum1 = sum1 + checknull(rs2.fields("PSUM").Value, 0)
		sumT = sumT + checknull(rs2.fields("pt").Value, 0)
	End If

End Sub
'=================
Sub GetSumByCodeDEK(mtdID, sum1, sumT, Inv)
	Dim rs
	If Inv = "инвалид" Then
		InitRs rs, _
			"SELECT SUM(JP_SUM) AS PSUM, sum(jp_tariff) as pt FROM P_JOURNAL WHERE JP_DONE = 1 " & _
				" AND C_PERIOD = " & WorkArea.Period.ID & _
				" AND MTC_ID = " & mtdID & " And JP_PL3 = 1"
	Else
		InitRs rs, _
			"SELECT SUM(JP_SUM) AS PSUM, sum(jp_tariff) as pt FROM P_JOURNAL WHERE JP_DONE = 1 " & _
				" AND C_PERIOD = " & WorkArea.Period.ID & _
				" AND MTC_ID = " & mtdID & " And (JP_PL3 = 0 or (JP_PL3 is Null))"
	End If
	
	If Not rs.eof	Then
		sum1 = rs.fields("PSUM").Value
		sumT = rs.fields("pt").Value
	End If
End Sub

'________________________________________________
' Function QueryWithParams( SqlText, Options, arParams )
' ¬ыполнение запроса с передачей в него параметров
' ¬озвращаетс€ объект типа RecordSet
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

	' устанавливаем параметры запроса
	For i = 0 To UBound( arParams ) Step 2
		val1 = arParams( i + 1 )
		If IsNull( val1 ) Then
			Set Prm = Cmd.CreateParameter( , arParams( i ), adParamInput, 10 )
		ElseIf arParams( i ) = adVarChar Then
			Set Prm = Cmd.CreateParameter( , arParams( i ), adParamInput, Iif( Len( val1 ) = 0, 1, Len( val1 )), val1 )
		Else
			Set Prm = Cmd.CreateParameter( , arParams( i ), adParamInput )st
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
 		       	 "                  And MTD_TYPE <> 103 " & _
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

	SetZero Totals		' итоги 3
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
				If Sheet( 1 ).Cell(i, j).cellParam <> MTD_LGT_ID Then
					If  IsNumeric(Sheet( 1 ).Cell(i, j).Value) Then
						GroupTotals(j) = GroupTotals(j) + Sheet( 1 ).Cell(i, j).Value
					End If
				End If
			Next
		End If
	Next

	For j = StartNumberOfColumns + 1 To Sheet(1).Columns	' итоги 3
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
					If .Cell(i, j).cellParam <> MTD_LGT_ID And IsNumeric(.Cell( i, j ).Value) Then 
						Sum = Sum + .Cell( i, j ).Value
					End If
				Next

				.Cell( i, TotalsCols( k )).Value = Sum
			Next
			.Cell(i, TotalsCols(2)+1).Value = .Cell(i, TotalsCols(1)).Value - .Cell( i, TotalsCols(2)).Value

'			.Cell( i, .Columns ).Value = .Cell( i, TotalsCols( 1 ) ).Value - .Cell( i, TotalsCols( 2 )).Value
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

	'If LastPos >= FirstPos Then Sheet( 1 ).Range(FirstPos, OnRow, LastPos, OnRow).MergeCells = True
End Sub
'________________________________________________

Sub display_code_calc(Code, Row, FOT_Sum, FOT_Percent)
	Sheet(1).Cell(Row, 7).Value = Code.DbCode
	Sheet(1).Cell(Row, 8).Value = Code.CrCode
	Sheet(1).Cell(Row, 2).Value = Code.percent
	Sheet(1).Cell(Row, 3).value = FOT_Sum
	Sheet(1).Cell(Row, 5).value = FOT_Percent
End Sub
'________________________________________________
Sub BuildSecondaryReport
	Dim max, percent, i, CodeID, FOT, FOT_GPD, sum1, sumT
	Dim arFonds, Repcalc, BaseSum
	Dim Code, FirstDay, min

	arFonds = Array( "2000", "2001", "2003", "2002", "2004")

	Set Repcalc = WorkArea.CreateIntObject("CommonCalc")
	Repcalc.setPeriod "month", Workarea.Period.Month, Workarea.Period.Year
	FirstDay = Workarea.Period.FirstDay

	max = Workarea.UTable( USRTBL_LGT ).GetValue( 1, USRTBL_MAXCY, 2 )
	min = max / 17

	percent = Workarea.UTable( USRTBL_LGT ).GetValue( 1, USRTBL_ECB, 2 )

	For i = 0 To UBound(arFonds)
		CodeID = WorkArea.GetCodeID( arFonds(i))

		If CodeID <> 0 Then
			Set Code = WorkArea.Code(CodeID)
			percent = Code.Percent

			FOT = repcalc.totalsumby("CodeSumW", CodeID)
			BaseSum = GetTariffByMtcPP(CodeID, Workarea.Period.ID)

'			display_code_calc Code, 5 + i, Round2( FOT * 100 / percent , 2), FOT
			display_code_calc Code, 5 + i, BaseSum, FOT
			Repcalc.Clear

		End If
	Next
	
End Sub
'________________________________________________
Sub BuildLastReport
	Dim percent, i, sum1, sumT, CodeID
	Dim arFonds

	sum1 = 0

	arFonds = Array( MTD_NDFL_ID, MTD_ESV_ID, MTD_ESV_INV_ID, MTD_ESVB_ID, MTD_ESVT_ID, MTD_ESVD_ID, MTD_ARMY_ID)
	percent  = Array( 15, 3.6, 3.6, 2, 2.6, 2, 1.5)

	With Sheet( 1 )
		For i = 0 To UBound(arFonds)
			GetSumByCode arFonds(i), sum1, sumT
			cRow = LastRow + i
			.Cell( cRow, 2 ).Value = percent(i)
			.Cell( cRow, 3 ).value = round2(sum1, 2)
			With WorkArea.Code( arFonds(i)(0) )
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
		sum1 = rs.fields("PSUM").Value
		sumT = rs.fields("pt").Value
	End If
End Sub
'=================
Sub GetSumByCode(mtdID, sum1, sumT)
	Dim rs1, rs2

	InitRs rs1, _
		"SELECT SUM(JP_SUM) AS PSUM, sum(jp_tariff) as pt FROM P_JOURNAL WHERE JP_DONE = 1 " & _
			" AND C_PERIOD = " & WorkArea.Period.ID & _
			" AND MTC_ID = " & mtdID(0)

	InitRs rs2, _
		"SELECT SUM(JP_SUM) AS PSUM, sum(jp_tariff) as pt FROM P_JOURNAL WHERE JP_DONE = 1 " & _
			" AND C_PERIOD = " & WorkArea.Period.ID & _
			" AND MTC_ID = " & mtdID(1)

	If Not rs1.eof	Then
		sum1 = checknull(rs1.fields("PSUM").Value, 0)
		sumT = checknull(rs1.fields("pt").Value, 0)
	Else
		Sum1 = 0
	End If

	If Not rs2.eof	Then
		sum1 = sum1 + checknull(rs2.fields("PSUM").Value, 0)
		sumT = sumT + checknull(rs2.fields("pt").Value, 0)
	End If

End Sub
'=================
Sub GetSumByCodeDEK(mtdID, sum1, sumT, Inv)
	Dim rs
	If Inv = "инвалид" Then
		InitRs rs, _
			"SELECT SUM(JP_SUM) AS PSUM, sum(jp_tariff) as pt FROM P_JOURNAL WHERE JP_DONE = 1 " & _
				" AND C_PERIOD = " & WorkArea.Period.ID & _
				" AND MTC_ID = " & mtdID & " And JP_PL3 = 1"
	Else
		InitRs rs, _
			"SELECT SUM(JP_SUM) AS PSUM, sum(jp_tariff) as pt FROM P_JOURNAL WHERE JP_DONE = 1 " & _
				" AND C_PERIOD = " & WorkArea.Period.ID & _
				" AND MTC_ID = " & mtdID & " And (JP_PL3 = 0 or (JP_PL3 is Null))"
	End If
	
	If Not rs.eof	Then
		sum1 = rs.fields("PSUM").Value
		sumT = rs.fields("pt").Value
	End If
End Sub

'________________________________________________
' Function QueryWithParams( SqlText, Options, arParams )
' ¬ыполнение запроса с передачей в него параметров
' ¬озвращаетс€ объект типа RecordSet
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

	' устанавливаем параметры запроса
	For i = 0 To UBound( arParams ) Step 2
		val1 = arParams( i + 1 )
		If IsNull( val1 ) Then
			Set Prm = Cmd.CreateParameter( , arParams( i ), adParamInput, 10 )
		ElseIf arParams( i ) = adVarChar Then
			Set Prm = Cmd.CreateParameter( , arParams( i ), adParamInput, Iif( Len( val1 ) = 0, 1, Len( val1 )), val1 )
		Else
			Set Prm = Cmd.CreateParameter( , arParams( i ), adParamInput )
			