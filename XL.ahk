#Persistent
#SingleInstance, Force

; GUI Layout and stuff
;---------------------------------------------
;---------------------------------------------

Gui, Default:New, +AlwaysOnTop, Informal Hearing Tool
Gui, Default:Font, cFFFFFF
Gui, Default:Color, d01919
Gui, Default:Show, w400 h520, 

Gui, Default:Add, Text, x10 y10, Adjust/Deny
Gui, Default:Add, Radio, vAdjust, Adjust
Gui, Default:Add, Radio, vDeny, Deny
Gui, Default:Add, Radio, vWithdraw, Withdraw

Gui, Default:Add, Text, x130 y10, Show/No Show
Gui, Default:Add, Radio, vShow gSNP, Show
Gui, Default:Add, Radio, vNoShow gSNP, No Show
Gui, Default:Add, Radio, vPhone gSNP, Phone

Gui, Default:Add, Text, x250 y10, Owner/Agent
Gui, Default:Add, Radio, vOwner gOAR, Owner
Gui, Default:Add, Radio, vAgent gOAR, Agent
Gui, Default:Add, Radio, vRental gOAR, Rental

Gui, Default:Add, Text, x10 y100, Date
Gui, Default:Add, DateTime, cBlack x10 y120 w90 vDate, M/d/yyyy

Gui, Default:Add, Text, x110 y100, Time
Gui, Default:Add, ComboBox, cBlack x110 y120 w70 vTime, 8:00 AM|8:30 AM|9:00 AM|9:30 AM|10:00 AM|10:30 AM|11:00 AM|11:30 AM|12:00 PM|12:30 PM|1:00 PM|1:30 PM|2:00 PM|2:30 PM|3:00 PM|3:30 PM|4:00 PM|4:30 PM

Gui, Default:Add, Text, x190 y100, Protest #
Gui, Default:Add, Edit, cBlack x190 y120 w70 number limit4 vProtestNumber,

Gui, Default:Add, Text, x270 y100, Name
Gui, Default:Add, Edit, cBlack x270 y120 w120 vName,

Gui, Default:Add, Text, x10 y150, Account Number
Gui, Default:Add, Edit, cBlack x10 y170 w60 number limit5 vAccount1,
Gui, Default:Add, Edit, cBlack x80 y170 w30 number limit2 vAccount2,
Gui, Default:Add, Edit, cBlack x120 y170 w30 number limit2 vAccount3,
Gui, Default:Add, Edit, cBlack x160 y170 w60 number limit5 vAccount4,

Gui, Default:Add, Text, x10 y200, Notes
Gui, Default:Add, Edit, cBLack x10 y220 w380 h50 vNotes,

Gui, Default:Add, Text, x10 y280, Current Value
Gui, Default:Add, Edit, cBlack x10 y300 w70 number vCurrent,

Gui, Default:Add, Text, x90 y280, Sq Ft
Gui, Default:Add, Edit, cBlack x90 y300 w40 number limit5 vSQFT,

Gui, Default:Add, Text, x140 y280, Year Built
Gui, Default:Add, Edit, cBlack x140 y300 w50 number limit4 vYearBuilt,

Gui, Default:Add, Text, x200 y280, Quality
Gui, Default:Add, DropDownList, vQuality, Low|Low +|Fair|Fair +|Avg|Avg +|Good|Good +|V. Good|V. Good +|Excellent

Gui, Default:Add, Text, x10 y330, Proposed
Gui, Default:Add, Edit, cBlack x10 y350 w70 number vProposed,

Gui, Default:Add, Text, x90 y330, Sq Ft
Gui, Default:Add, Edit, cBlack x90 y350 w40 number limit4 vSQFT2,

Gui, Default:Add, Text, x140 y330, Year Built
Gui, Default:Add, Edit, cBlack x140 y350 w50 number limit4 vYearBuilt2,

Gui, Default:Add, Text, x200 y330, Quality
Gui, Default:Add, DropDownList, vQuality2, Low|Low +|Fair|Fair +|Avg|Avg +|Good|Good +|V. Good|V. Good +|Excellent

Gui, Default:Add, Text, x10 y380, Land
Gui, Default:Add, Edit, CBlack x10 y400 w70 number limit9 vLand,

Gui, Default:Add, Radio, x10 y430 vVacant gImprovement, Vacant 
Gui, Default:Add, Radio, x10 y450 vSingle gImprovement, One Imp
Gui, Default:Add, Radio, x10 y470 vMulti gImprovement, Multiple Imps

Gui, Default:Add, Text, x250 y380, Reason
Gui, Default:Add, Radio, vComp gReason, Comp. Sales
Gui, Default:Add, Radio, vCost gReason, Cost
Gui, Default:Add, Radio, vIncome gReason, Income
Gui, Default:Add, Radio, vAppraisal gReason, Appraisal 
Gui, Default:Add, Radio, vData gReason, Data Correction 
Gui, Default:Add, Radio, vOther gReason, Other/See Notes

Gui, Default:Add, Text, x120 y380, Decision
Gui, Default:Add, Radio, vAdjMkt gDecision, Adj. Mkt
Gui, Default:Add, Radio, vDeny2 gDecision, Deny 
Gui, Default:Add, Radio, vWithdraw2 gDecision, Withdraw 
Gui, Default:Add, Radio, vLate gDecision, Late File 
Gui, Default:Add, Radio, vAMaT gDecision, Adj. Mkt and Taxable
Gui, Default:Add, Radio, vAdjTax gDecision, Adj. Taxable

Gui, Default:Add, Button, x10 y490 gDoEverything, Submit

Gui, Default:Add, Button, x70 y490 gReset, Reset

return


; Labels
;-------------------------------------------------
;-------------------------------------------------

Reset:  ; Reloads GUI when pressed
	Reload
return

Clear:  ; Clears all variables on GUI

Comp := 
Cost := 
Income := 
Appraisal := 
Data := 
Other := 
AdjMkt := 
Deny := 
Deny2 := 
Withdraw := 
Withdraw2 := 
Late := 
AMaT := 
AdjTax := 
Adjust := 
Show := 
NoShow := 
Phone := 
Owner := 
Agent := 
Rental := 
Date := 
Time := 
ProtestNumber := 
Name := 
Account1 := 
Account2 := 
Account3 := 
Account4 := 
Notes := 
Current := 
SQFT := 
SQFT2 := 
YearBuilt := 
YearBuilt2 := 
Quality := 
Quality2 := 
Proposed := 
Land := 
Vacant := 
Single := 
Multi := 

return


; Exit and pause stuff
;-----------------------------------------------

DefaultGuiClose:

ExitApp

return

Esc::Pause


; Does all the stuff
;-----------------------------------------------
;-----------------------------------------------

DoEverything:
	Deploy()
return

Deploy() {

	Gui, Default:Submit, Nohide
	Gui, Default:Minimize
	MsgBox, Make sure there are no tabs open in RealWare, and the lookup field is empty, then press OK.
	

;Declare global variables from GUI and File
;-----------------------------------------------
#Include C:\Users\%A_UserName%\Documents\ahk\coordinates.ahk

global Comp 
global Cost 
global Income 
global Appraisal 
global Data 
global Other 
global AdjMkt 
global Deny 
global Deny2 
global Withdraw 
global Withdraw2 
global Late
global AMaT 
global AdjTax 
global Adjust 
global Show 
global NoShow 
global Phone 
global Owner 
global Agent 
global Rental 
global Date
global Time 
global ProtestNumber 
global Name 
global Account1 
global Account2 
global Account3 
global Account4 
global Notes 
global Current 
global SQFT  
global SQFT2  
global YearBuilt 
global YearBuilt2  
global Quality 
global Quality2 
global Proposed 
global Land 
global Vacant  
global Single  
global Multi 

;Change Date to acceptable format for Excel
;-----------------------------------------------
global NewDate
global NewYear
global NewMonth
global NewDay
global FinalDate

StringTrimRight, NewDate, Date, 6

NewYear := SubStr(NewDate, 1, 4)
NewDay := SubStr(NewDate, 7, 2)
NewMonth := SubStr(NewDate, 5, 2)

FinalDate := NewMonth . "/" . NewDay . "/" . NewYear


; Progress bar stuff
;-----------------------------------------------
width := A_ScreenWidth - 220
height := A_ScreenHeight
static TheProgress

Gui, ProgressBar:New, +AlwaysOnTop, Informal Hearing Tool
Gui, ProgressBar:Font, c39ff14
Gui, ProgressBar:Color, c000000
Gui, ProgressBar:Add, Text, x10 y5, Working...
Gui, ProgressBar:Show, w220 h60 x%width% y2
Gui, ProgressBar:Add, Progress, w200 h30 c39ff14 Background4d4e4f Range0-100 vTheProgress

;-Excel Stuff
;-----------------------------------

xl := ComObjCreate("Excel.Application")
xl.Visible := true
path = C:\Users\%A_UserName%\Documents\ahk\Protest.xlsx
xl.Workbooks.Open(Path)

xl.range("G3").value := FinalDate
xl.range("AB3").value := Time
xl.range("AS3").value := ProtestNumber
xl.range("L5").value := Account1
xl.range("V5").value := Account2
xl.range("Z5").value := Account3 
xl.range("AD5").value := Account4 
xl.range("L8").value := Name
xl.range("A21").value := Notes
xl.range("N36").value := Current
xl.range("AC36").value := SQFT 
xl.range("AQ36").value := YearBuilt 
xl.range("AV36").value := Quality
xl.range("N38").value := Proposed
xl.range("AC38").value := SQFT2 
xl.range("AQ38").value := YearBuilt2
xl.range("AV38").value := Quality2
xl.range("K55").value := Land
xl.range("K57").value := Proposed
xl.Range("N59").value := FieldNumber

GoSub, SNP
Gosub, OAR
GoSub, Decision
GoSub, Reason

xl.ActiveSheet.Printout
xl.ActiveWorkbook.saved := True
xl.Quit

GoSub, Improvement
Gui, ProgressBar:Destroy
GoSub, Clear

return

; SubRoutines
;---------------------------------------------------
;---------------------------------------------------

SNP:
; Marks respective box on excel form
;---------------------------------------------------
If (Show = 1)
	{
	xl.range("AR7").value := "x"
	}
else If (NoShow = 1)
	{
	xl.range("AR11").value := "x"
	}
else If (Phone = 1)
	{
	xl.range("AR9").value := "x"
	}
return

OAR:
; Marks respective box on excel form
;---------------------------------------------------
If (Owner = 1)
	{
	xl.range("AX7").value := "x"
	}
else If (Agent = 1)
	{
	xl.range("AX9").value := "x"
	}
else If (Rental = 1)
	{
	xl.range("AX11").value := "x"
	}	
return

Decision:
; Marks respective box on excel form
;---------------------------------------------------
If (AdjMkt = 1)
	{
	xl.range("G43").value := "x"
	}
else If (Deny = 1)
	{
	xl.range("G45").value := "x"
	}
else If (Withdraw2 = 1)
	{
	xl.range("G47").value := "x"
	}
else If (Late = 1)
	{
	xl.range("G49").value := "x"
	}
else If (AMaT = 1)
	{
	xl.range("G51").value := "x"
	}
else If (AdjTax = 1)
	{
	xl.range("G53").value := "x"
	}	
return

Reason:
; Marks respective box on excel form
;----------------------------------------------------
If (Comp = 1)
	{
	xl.range("AJ43").value := "x"
	}
else If (Cost = 1)
	{
	xl.range("AJ45").value := "x"
	}
else If (Income = 1)
	{
	xl.range("AJ47").value := "x"
	}
else If (Appraisal = 1)
	{
	xl.range("AJ49").value := "x"
	}
else If (Data = 1)
	{
	xl.range("AJ51").value := "x"
	}
else If (Other = 1)
	{
	xl.range("AJ53").value := "x"
	}
return

Improvement:
; Inputs data into RealWare - User Input for Multiple Improvements if Imps > 1
;-------------------------------------------------------

SetKeyDelay, 50
SetControlDelay, 100
Num := StrLen(ProtestNumber)

WinActivate RealWare
WinMaximize RealWare

; Account changes and other stuff
;-------------------------------------------------------

{
If (Single = 1)
{
	IF (Adjust = 1)
	{
		MouseClick left, %Realx%, %Realy% ;  Selects Real Major
		GuiControl,, TheProgress, +1
		Sleep 100
		GuiControl,, TheProgress, +1
		MouseClick left, %Accountx%, %Accounty%  ;  Selects Account Minor
		GuiControl,, TheProgress, +1
		Sleep 100
		GuiControl,, TheProgress, +1
		MouseClick left, %Lookupx%, %Lookupy%  ;  Selects Account Lookup
		GuiControl,, TheProgress, +1
		Sleep 100
		Send,  %Account1%%Account2%%Account3%%Account4%  ; Enters account number 
		GuiControl,, TheProgress, +1
		SendInput {Enter}
		GuiControl,, TheProgress, +1
		Loop, 5
			{
				Sleep 500
				GuiControl,, TheProgress, +1
			}
		MouseClick left, %LandValuex%, %LandValuey%  ; Selects Land Value
		GuiControl,, TheProgress, +1
		Sleep 50
		Send, %Land%
		GuiControl,, TheProgress, +1
		Sleep 500
		GuiControl,, TheProgress, +1
		MouseClick left, %DefaultApproachx%, %DefaultApproachy%  ; Selects Default Approach
		GuiControl,, TheProgress, +1
		Sleep 500
		GuiControl,, TheProgress, +1
		Send, Reconciled
		GuiControl,, TheProgress, +1
		Sleep 500
		GuiControl,, TheProgress, +1
		SendInput {ctrl down}s{ctrl up}  ;  Save
		Sleep 500
		GuiControl,, TheProgress, +1
		SendInput {Enter}
		GuiControl,, TheProgress, +1
		Loop, 5
			{
				Sleep 500
				GuiControl,, TheProgress, +2
			}
		MouseClick left, %Improvementsx%, %Improvementsy%  ; Selects Improvements Minor
		Loop, 13
			{
				Sleep 500
				GuiControl,, TheProgress, +1
			}
		MouseClick left, %Reconciledx%, %Reconciledy%  ; Selects Reconciled Tab
		GuiControl,, TheProgress, +1
		Loop, 5
			{
				Sleep 500
				GuiControl,, TheProgress, +2
			}
		MouseClick left, %TotalValuex%, %TotalValuey%  ; Selects Total Reconciled Value
		GuiControl,, TheProgress, +1
		Send, %Proposed%
		GuiControl,, TheProgress, +1
		Sleep 500
		GuiControl,, TheProgress, +1
		SendInput {Tab}
		GuiControl,, TheProgress, +1
		Sleep 500
		SendInput {ctrl down}s{ctrl up}
		Loop, 5
			{
				Sleep 500
				GuiControl,, TheProgress, +2
			}
		MouseClick left, %Appealsx%, %Appealsy%  ; Selects Appeals Major
		GuiControl,, TheProgress, +1
		Sleep 50
		MouseClick left, %Detailx%, %Detaily%  ; Selects Details Minor
		GuiControl,, TheProgress, +1
		Sleep 50
		MouseClick left, %Lookupx%, %Lookupy%  ;  Selects Account Lookup and enters Protest Number
		GuiControl,, TheProgress, +1
		Sleep 100
		If (Num = 1)
			{
				Send, %TaxYear%00000%ProtestNumber%
				SendInput {Enter}
				Sleep 2500
				GuiControl,, TheProgress, +1
			}
		else If (Num = 2)
			{
				Send, %TaxYear%0000%ProtestNumber%
				SendInput {Enter}
				Sleep 2500
				GuiControl,, TheProgress, +1
			}
		else If (Num = 3)
			{
				Send, %TaxYear%000%ProtestNumber%
				SendInput {Enter}
				Sleep 2500
				GuiControl,, TheProgress, +1
			}
		else If (Num = 4)
			{
				Send, %TaxYear%00%ProtestNumber%
				SendInput {Enter}
				Sleep 2500
				GuiControl,, TheProgress, +1
			}
		MouseClick left, %Decisionx%, %Decisiony%  ; Selects Appeal Decision
		GuiControl,, TheProgress, +1
		Sleep 100
		Send, Adjust
		GuiControl,, TheProgress, +1
		Sleep 500
		MouseClick left, %Reasonx%, %Reasony%  ; Selects Adjust/Deny Reason
		If (Comp = 1)
			{
				SendRaw, Comp. Sales
				SendInput {Tab}
				GuiControl,, TheProgress, +1
			}
		else If (Cost = 1)
			{
				SendRaw, Cost
				SendInput {Tab}
				GuiControl,, TheProgress, +1
			}
		else If (Income = 1)
			{
				SendRaw, Income
				SendInput {Tab}
				GuiControl,, TheProgress, +1
			}
		else If (Appraisal = 1)
			{
				SendRaw, Appraisal
				SendInput {Tab}
				GuiControl,, TheProgress, +1
			}
		else If (Data = 1)
			{
				SendRaw, Data Correction
				SendInput {Tab}
				GuiControl,, TheProgress, +1
			}
		else If (Other = 1)
			{
				SendRaw, Other/See Note
				SendInput {Tab}
				GuiControl,, TheProgress, +1
			}
		Sleep 100
		GuiControl,, TheProgress, +1
		SendInput {ctrl down}s{ctrl up}
		Loop, 5
			{
				Sleep 500
				GuiControl,, TheProgress, +1
			}
		MouseClick left, %Notesx%, %Notesy%  ; Selects Notes Minor
		Loop, 3
			{
				Sleep 500
				GuiControl,, TheProgress, +2
			}
		MouseClick left, %NoteFieldx%, %NoteFieldy%  ; Selects Note Field
		GuiControl,, TheProgress, +1
		Sleep 500
		GuiControl,, TheProgress, +1
		SendInput %Notes%
		GuiControl,, TheProgress, +1
		Sleep 100
		GuiControl,, TheProgress, +1
		SendInput {Tab}
		GuiControl,, TheProgress, +1
		Sleep 100
		GuiControl,, TheProgress, +1
		SendInput {Space}
		GuiControl,, TheProgress, +1
		Sleep 100
		GuiControl,, TheProgress, +1
		SendInput {ctrl down}s{ctrl up}
		GuiControl,, TheProgress, +1
		Sleep 500
		MouseClick left, %Summaryx%, %Summaryy%  ; Selects Summary Minor
		Loop, 7
			{
				Sleep 500
				GuiControl,, TheProgress, +1
			}
		MouseClick left, %Showx%, %Showy%  ; Selects Show/No Show/Phone
		If (Show = 1)
			{
				Send, Show
				GuiControl,, TheProgress, +1
			}
		else If (NoShow = 1)
			{
				Send, No Show
				GuiControl,, TheProgress, +1
			}
		else If (Phone = 1)
			{
				Send, Phone
				GuiControl,, TheProgress, +1
			}
		Sleep 500
		MouseClick left, %DecisionByx%, %DecisionByy%
		Sleep 500
		Send, %FieldNumber%
		GuiControl,, TheProgress, +1
		Sleep 500
		SendInput {ctrl down}s{ctrl up}
	}
	else If (Deny = 1)
	{
		MouseClick left, %Appealsx%, %Appealsy%  ; Selects Appeals Major
		GuiControl,, TheProgress, +2
		Sleep 50
		GuiControl,, TheProgress, +2
		MouseClick left, %Detailx%, %Detaily%  ; Selects Details Minor
		GuiControl,, TheProgress, +2
		Sleep 50
		GuiControl,, TheProgress, +2
		MouseClick left, %Lookupx%, %Lookupy%  ;  Selects Account Lookup
		GuiControl,, TheProgress, +2
		Sleep 100
		If (Num = 1)
			{
				Send, %TaxYear%00000%ProtestNumber%
				SendInput {Enter}
				Loop, 5
				{
					Sleep 500
					GuiControl,, TheProgress, +3
				}
			}
		else If (Num = 2)
			{
				Send, %TaxYear%0000%ProtestNumber%
				SendInput {Enter}
				Loop, 5
				{
					Sleep 500
					GuiControl,, TheProgress, +3
				}
			}
		else If (Num = 3)
			{
				Send, %TaxYear%000%ProtestNumber%
				SendInput {Enter}
				Loop, 5
				{
					Sleep 500
					GuiControl,, TheProgress, +3
				}
			}
		else If (Num = 4)
			{
				Send, %TaxYear%00%ProtestNumber%
				SendInput {Enter}
				Loop, 5
				{
					Sleep 500
					GuiControl,, TheProgress, +3
				}
			}
		MouseClick left, %Decisionx%, %Decisiony%  ; Selects Appeal Decision
		GuiControl,, TheProgress, +2
		Send, Deny
		GuiControl,, TheProgress, +2
		Sleep 500
		GuiControl,, TheProgress, +2
		MouseClick left, %Reasonx%, %Reasony%  ; Selects Adjust/Deny Reason
		GuiControl,, TheProgress, +2
		If (Comp = 1)
			{
				SendRaw, Comp. Sales
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Cost = 1)
			{
				SendRaw, Cost
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Income = 1)
			{
				SendRaw, Income
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Appraisal = 1)
			{
				SendRaw, Appraisal
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Data = 1)
			{
				SendRaw, Data Correction
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Other = 1)
			{
				SendRaw, Other/See Note
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		Sleep 100
		GuiControl,, TheProgress, +2
		SendInput {ctrl down}s{ctrl up}
		Loop, 5
			{
				Sleep 500
				GuiControl,, TheProgress, +2
			}
		MouseClick left, %Notesx%, %Notesy%  ; Selects Notes Minor
		Loop, 3
			{
				Sleep 500
				GuiControl,, TheProgress, +2
			}
		MouseClick left, %NoteFieldx%, %NoteFieldy%  ; Selects Note Field
		GuiControl,, TheProgress, +2
		Sleep 500
		GuiControl,, TheProgress, +2
		SendInput %Notes%
		GuiControl,, TheProgress, +2
		Sleep 100
		GuiControl,, TheProgress, +2
		SendInput {Tab}
		GuiControl,, TheProgress, +2
		Sleep 100
		GuiControl,, TheProgress, +2
		SendInput {Space}
		GuiControl,, TheProgress, +2
		Sleep 100	
		GuiControl,, TheProgress, +2
		SendInput {ctrl down}s{ctrl up}
		GuiControl,, TheProgress, +2
		Sleep 500
		MouseClick left, %Summaryx%, %Summaryy%  ; Selects Summary Minor
		Loop, 7
			{
				Sleep 500
				GuiControl,, TheProgress, +2
			}
		MouseClick left, %Showx%, %Showy%  ; Selects Show/No Show/Phone
		If (Show = 1)
			{
				Send, Show
				GuiControl,, TheProgress, +3
			}
		else If (NoShow = 1)
			{
				Send, No Show
				GuiControl,, TheProgress, +3
			}
		else If (Phone = 1)
			{
				Send, Phone
				GuiControl,, TheProgress, +3
			}
		Sleep 500
		GuiControl,, TheProgress, +3
		MouseClick left, %DecisionByx%, %DecisionByy%
		Sleep 500
		GuiControl,, TheProgress, +3
		Send, %FieldNumber%
		GuiControl,, TheProgress, +3
		Sleep 500
		GuiControl,, TheProgress, +3
		SendInput {ctrl down}s{ctrl up}
	}
	else If (Withdraw = 1)
	{
		MouseClick left, %Appealsx%, %Appealsy%  ; Selects Appeals Major
		GuiControl,, TheProgress, +2
		Sleep 50
		GuiControl,, TheProgress, +3
		MouseClick left, %Detailx%, %Detaily%  ; Selects Details Minor
		GuiControl,, TheProgress, +2
		Sleep 50
		GuiControl,, TheProgress, +3
		MouseClick left, %Lookupx%, %Lookupy%  ;  Selects Account Lookup
		GuiControl,, TheProgress, +2
		Sleep 100	
		If (Num = 1)
			{
				Send, %TaxYear%00000%ProtestNumber%
				SendInput {Enter}
				Loop, 5
					{
						GuiControl,, TheProgress, +3
						Sleep 500
					}
			}
		else If (Num = 2)
			{
				Send, %TaxYear%0000%ProtestNumber%
				SendInput {Enter}
				Loop, 5
					{
						GuiControl,, TheProgress, +3
						Sleep 500
					}
			}
		else If (Num = 3)
			{
				Send, %TaxYear%000%ProtestNumber%
				SendInput {Enter}
				Loop, 5
					{
						GuiControl,, TheProgress, +3
						Sleep 500
					}
			}
		else If (Num = 4)
			{
				Send, %TaxYear%00%ProtestNumber%
				SendInput {Enter}
				Loop, 5
					{
						GuiControl,, TheProgress, +3
						Sleep 500
					}
			}
		MouseClick left, %Decisionx%, %Decisiony%  ; Selects Appeal Decision
		GuiControl,, TheProgress, +2
		Send, Withdrawn
		GuiControl,, TheProgress, +2
		Sleep 500
		GuiControl,, TheProgress, +2
		MouseClick left, %Reasonx%, %Reasony%  ; Selects Adjust/Deny Reason
		GuiControl,, TheProgress, +2
		If (Comp = 1)
			{
				SendRaw, Comp. Sales
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Cost = 1)
			{
				Send, Cost
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Income = 1)
			{
				Send, Income
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Appraisal = 1)
			{
				Send, Appraisal
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Data = 1)
			{
				SendRaw, Data Correction
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Other = 1)
			{
				SendRaw, Other/See Note
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		Sleep 100
		GuiControl,, TheProgress, +2
		SendInput {ctrl down}s{ctrl up}
		Loop, 5
			{
				GuiControl,, TheProgress, +3
				Sleep 500
			}
		MouseClick left, %Notesx%, %Notesy%  ; Selects Notes Minor
		Loop, 3
			{
				GuiControl,, TheProgress, +2
				Sleep 500
			}
		MouseClick left, %NoteFieldx%, %NoteFieldy%  ; Selects Note Field
		GuiControl,, TheProgress, +2
		Sleep 500
		GuiControl,, TheProgress, +2
		SendInput %Notes%
		GuiControl,, TheProgress, +2
		Sleep 100
		GuiControl,, TheProgress, +2
		SendInput {Tab}
		GuiControl,, TheProgress, +2
		Sleep 100
		GuiControl,, TheProgress, +2
		SendInput {Space}
		GuiControl,, TheProgress, +2
		Sleep 100
		GuiControl,, TheProgress, +2
		SendInput {ctrl down}s{ctrl up}
		GuiControl,, TheProgress, +2
		Sleep 500
		GuiControl,, TheProgress, +2
		MouseClick left, %Summaryx%, %Summaryy%  ; Selects Summary Minor
		Loop, 7
			{
				GuiControl,, TheProgress, +2
				Sleep 500
			}
		MouseClick left, %Showx%, %Showy%  ; Selects Show/No Show/Phone
		If (Show = 1)
			{
				Send, Show
				GuiControl,, TheProgress, +2
			}
		else If (NoShow = 1)
			{
				Send, No Show
				GuiControl,, TheProgress, +2
			}
		else If (Phone = 1)
			{
				Send, Phone
				GuiControl,, TheProgress, +2
			}
		Sleep 500
		Mouseclick left, %DecisionByx%, %DecisionByy%
		Sleep 200
		Send, %FieldNumber%
		GuiControl,, TheProgress, +2
		Sleep 500
		GuiControl,, TheProgress, +2
		SendInput {ctrl down}s{ctrl up}
	}
}
else If (Multi = 1)
{
	IF (Adjust = 1)
	{
		MouseClick left, %Realx%, %Realy% ;  Selects Real Major
		GuiControl,, TheProgress, +1
		Sleep 100
		GuiControl,, TheProgress, +1
		MouseClick left, %Accountx%, %Accounty%  ;  Selects Account Minor
		GuiControl,, TheProgress, +1
		Sleep 100
		GuiControl,, TheProgress, +1
		MouseClick left, %Lookupx%, %Lookupy%  ;  Selects Account Lookup
		GuiControl,, TheProgress, +1
		Sleep 100
		Send,  %Account1%%Account2%%Account3%%Account4%
		GuiControl,, TheProgress, +1
		SendInput {Enter}
		GuiControl,, TheProgress, +1
		Loop, 5
			{
				Sleep 500
				GuiControl,, TheProgress, +1
			}
		MouseClick left, %LandValuex%, %LandValuey%  ; Selects Land Value
		GuiControl,, TheProgress, +1
		Sleep 50
		Send, %Land%
		GuiControl,, TheProgress, +1
		Sleep 500
		GuiControl,, TheProgress, +1
		MouseClick left, %DefaultApproachx%, %DefaultApproachy%  ; Selects Default Approach
		GuiControl,, TheProgress, +1
		Sleep 500
		GuiControl,, TheProgress, +1
		Send, Reconciled
		GuiControl,, TheProgress, +1
		Sleep 500
		GuiControl,, TheProgress, +1
		SendInput {ctrl down}s{ctrl up}
		Sleep 500
		GuiControl,, TheProgress, +1
		SendInput {Enter}
		GuiControl,, TheProgress, +1
		Loop, 5
			{
				Sleep 500
				GuiControl,, TheProgress, +1
			}
		MouseClick left, %Improvementsx%, %Improvementsy%  ; Selects Improvements Minor
		Loop, 13
			{
				Sleep 500
				GuiControl,, TheProgress, +1
			}
		MouseClick left, %Reconciledx%, %Reconciledy%  ; Selects Reconciled Tab
		GuiControl,, TheProgress, +1
		Loop, 5
			{
				Sleep 500
				GuiControl,, TheProgress, +1
			}
		InputBox, Imps, Improvements, How many Improvements?
		Sleep 100
		MouseClick left, %ImpValuex%, %ImpValuey%  ; Selects Imps Reconciled Value
		Loop, %Imps%
			{
				InputBox, Imp, Improvements, Improvement Value?,
				Sleep 100
				MouseClick left, %Impvaluex%, %ImpValuey%
				Sleep 100
				SendInput %Imp%
				Sleep 100
			}
		Sleep 500
		SendInput {ctrl down}s{ctrl up}
		Loop, 5
			{
				Sleep 500
				GuiControl,, TheProgress, +1
			}
		MouseClick left, %Appealsx%, %Appealsy%  ; Selects Appeals Major
		GuiControl,, TheProgress, +1
		Sleep 50
		MouseClick left, %Detailx%, %Detaily%  ; Selects Details Minor
		GuiControl,, TheProgress, +2
		Sleep 50
		MouseClick left, %Lookupx%, %Lookupy%  ;  Selects Account Lookup
		GuiControl,, TheProgress, +1
		Sleep 100
		If (Num = 1)
			{
				Send, %TaxYear%00000%ProtestNumber%
				SendInput {Enter}
				Sleep 2500
				GuiControl,, TheProgress, +1
	
			}
		else If (Num = 2)
			{
				Send, %TaxYear%0000%ProtestNumber%
				SendInput {Enter}
				Sleep 2500
				GuiControl,, TheProgress, +1
	
			}
		else If (Num = 3)
			{
				Send, %TaxYear%000%ProtestNumber%
				SendInput {Enter}
				Sleep 2500
				GuiControl,, TheProgress, +1
	
			}
		else If (Num = 4)
			{
				Send, %TaxYear%00%ProtestNumber%
				SendInput {Enter}
				Sleep 2500
				GuiControl,, TheProgress, +1
	
			}
		MouseClick left, %Decisionx%, %Decisiony%  ; Selects Appeal Decision
		GuiControl,, TheProgress, +2
		Sleep 100
		Send, Adjust
		GuiControl,, TheProgress, +2
		Sleep 500
		MouseClick left, %Reasonx%, %Reasony%  ; Selects Adjust/Deny Reason
		If (Comp = 1)
			{
				SendRaw, Comp. Sales
				SendInput {Tab}
				GuiControl,, TheProgress, +1
	
			}
		else If (Cost = 1)
			{
				SendRaw, Cost
				SendInput {Tab}
				GuiControl,, TheProgress, +1
	
			}
		else If (Income = 1)
			{
				SendRaw, Income
				SendInput {Tab}
				GuiControl,, TheProgress, +1
	
			}
		else If (Appraisal = 1)
			{
				SendRaw, Appraisal
				SendInput {Tab}
				GuiControl,, TheProgress, +1
	
			}
		else If (Data = 1)
			{
				SendRaw, Data Correction
				SendInput {Tab}
				GuiControl,, TheProgress, +1
	
			}
		else If (Other = 1)
			{
				SendRaw, Other/See Note
				SendInput {Tab}
				GuiControl,, TheProgress, +1
	
			}
		Sleep 100
		GuiControl,, TheProgress, +1
		SendInput {ctrl down}s{ctrl up}
		Loop, 5
			{
				Sleep 500
				GuiControl,, TheProgress, +1
			}
		MouseClick left, %Notesx%, %Notesy%  ; Selects Notes Minor
		Loop, 3
			{
				Sleep 500
				GuiControl,, TheProgress, +1
			}
		MouseClick left, %NoteFieldx%, %NoteFieldy%  ; Selects Note Field
		GuiControl,, TheProgress, +2
		Sleep 500
		GuiControl,, TheProgress, +2
		SendInput %Notes%
		GuiControl,, TheProgress, +2
		Sleep 100
		GuiControl,, TheProgress, +2
		SendInput {Tab}
		GuiControl,, TheProgress, +2
		Sleep 100
		GuiControl,, TheProgress, +2
		SendInput {Space}
		GuiControl,, TheProgress, +2
		Sleep 100
		GuiControl,, TheProgress, +2
		SendInput {ctrl down}s{ctrl up}
		GuiControl,, TheProgress, +2
		Sleep 500
		MouseClick left, %Summaryx%, %Summaryy%  ; Selects Summary Minor
		Loop, 7
			{
				Sleep 500
				GuiControl,, TheProgress, +1
			}
		MouseClick left, %Showx%, %Showy%  ; Selects Show/No Show/Phone
		If (Show = 1)
			{
				Send, Show
				GuiControl,, TheProgress, +1
	
			}
		else If (NoShow = 1)
			{
				Send, No Show
				GuiControl,, TheProgress, +1
	
			}
		else If (Phone = 1)
			{
				Send, Phone
				GuiControl,, TheProgress, +1
	
			}
		Sleep 500
		MouseClick left, %DecisionByx%, %DecisionByy%
		Sleep 500
		Send, %FieldNumber%
		GuiControl,, TheProgress, +1
		Sleep 500
		SendInput {ctrl down}s{ctrl up}
	}
	else If (Deny = 1)
	{
		MouseClick left, %Appealsx%, %Appealsy%  ; Selects Appeals Major
		GuiControl,, TheProgress, +2
		Sleep 50
		GuiControl,, TheProgress, +2
		MouseClick left, %Detailx%, %Detaily%  ; Selects Details Minor
		GuiControl,, TheProgress, +2
		Sleep 50
		GuiControl,, TheProgress, +2
		MouseClick left, %Lookupx%, %Lookupy%  ; Selects Account Lookup
		GuiControl,, TheProgress, +2
		Sleep 100
		If (Num = 1)
			{
				Send, %TaxYear%00000%ProtestNumber%
				SendInput {Enter}
				Loop, 5
				{
					Sleep 500
					GuiControl,, TheProgress, +2
				}
			}
		else If (Num = 2)
			{
				Send, %TaxYear%0000%ProtestNumber%
				SendInput {Enter}
				Loop, 5
				{
					Sleep 500
					GuiControl,, TheProgress, +2
				}
			}
		else If (Num = 3)
			{
				Send, %TaxYear%000%ProtestNumber%
				SendInput {Enter}
				Loop, 5
				{
					Sleep 500
					GuiControl,, TheProgress, +2
				}
			}
		else If (Num = 4)
			{
				Send, %TaxYear%00%ProtestNumber%
				SendInput {Enter}
				Loop, 5
				{
					Sleep 500
					GuiControl,, TheProgress, +2
				}
			}
		MouseClick left, %Decisionx%, %Decisiony%  ; Selects Appeal Decision
		GuiControl,, TheProgress, +2
		Send, Deny
		GuiControl,, TheProgress, +2
		Sleep 500
		GuiControl,, TheProgress, +2
		MouseClick left, %Reasonx%, %Reasony%  ; Selects Adjust/Deny Reason
		GuiControl,, TheProgress, +2
		If (Comp = 1)
			{
				SendRaw, Comp. Sales
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Cost = 1)
			{
				SendRaw, Cost
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Income = 1)
			{
				SendRaw, Income
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Appraisal = 1)
			{
				SendRaw, Appraisal
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Data = 1)
			{
				SendRaw, Data Correction
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Other = 1)
			{
				SendRaw, Other/See Note
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		Sleep 100
		GuiControl,, TheProgress, +2
		SendInput {ctrl down}s{ctrl up}
		Loop, 5
			{
				Sleep 500
				GuiControl,, TheProgress, +2
			}
		MouseClick left, %Notesx%, %Notesy%  ; Selects Notes Minor
		Loop, 3
			{
				Sleep 500
				GuiControl,, TheProgress, +2
			}
		MouseClick left, %NoteFieldx%, %NoteFieldy%  ; Selects Note Field
		GuiControl,, TheProgress, +2
		Sleep 500
		GuiControl,, TheProgress, +2
		SendInput %Notes%
		GuiControl,, TheProgress, +2
		Sleep 100
		GuiControl,, TheProgress, +2
		SendInput {Tab}
		GuiControl,, TheProgress, +2
		Sleep 100
		GuiControl,, TheProgress, +2
		SendInput {Space}
		GuiControl,, TheProgress, +2
		Sleep 100	
		GuiControl,, TheProgress, +2
		SendInput {ctrl down}s{ctrl up}
		GuiControl,, TheProgress, +2
		Sleep 500
		MouseClick left, %Summaryx%, %Summaryy%  ; Selects Summary Minor
		Loop, 7
			{
				Sleep 500
				GuiControl,, TheProgress, +2
			}
		MouseClick left, %Showx%, %Showy%  ; Selects Show/No Show/Phone
		If (Show = 1)
			{
				Send, Show
				GuiControl,, TheProgress, +2
			}
		else If (NoShow = 1)
			{
				Send, No Show
				GuiControl,, TheProgress, +2
			}
		else If (Phone = 1)
			{
				Send, Phone
				GuiControl,, TheProgress, +2
			}
		Sleep 500
		GuiControl,, TheProgress, +2
		MouseClick left, %DecisionByx%, %DecisionByy%
		Sleep 500
		GuiControl,, TheProgress, +1
		Send, %FieldNumber%
		GuiControl,, TheProgress, +1
		Sleep 500
		GuiControl,, TheProgress, +1
		SendInput {ctrl down}s{ctrl up}
	}
	else If (Withdraw = 1)
	{
		MouseClick left, %Appealsx%, %Appealsy%  ; Selects Appeals Major
		GuiControl,, TheProgress, +2
		Sleep 50
		GuiControl,, TheProgress, +2
		MouseClick left, %Detailx%, %Detaily%  ; Selects Details Minor
		GuiControl,, TheProgress, +2
		Sleep 50
		GuiControl,, TheProgress, +2
		MouseClick left, %Lookupx%, %Lookupy%  ; Selects Account Lookup
		GuiControl,, TheProgress, +2
		Sleep 100	
		If (Num = 1)
			{
				Send, %TaxYear%00000%ProtestNumber%
				SendInput {Enter}
				Loop, 5
					{
						GuiControl,, TheProgress, +2
						Sleep 500
					}
			}
		else If (Num = 2)
			{
				Send, %TaxYear%0000%ProtestNumber%
				SendInput {Enter}
				Loop, 5
					{
						GuiControl,, TheProgress, +2
						Sleep 500
					}
			}
		else If (Num = 3)
			{
				Send, %TaxYear%000%ProtestNumber%
				SendInput {Enter}
				Loop, 5
					{
						GuiControl,, TheProgress, +2
						Sleep 500
					}
			}
		else If (Num = 4)
			{
				Send, %TaxYear%00%ProtestNumber%
				SendInput {Enter}
				Loop, 5
					{
						GuiControl,, TheProgress, +2
						Sleep 500
					}
			}
		MouseClick left, %Decisionx%, %Decisiony%  ; Selects Appeal Decision
		GuiControl,, TheProgress, +2
		Send, Withdrawn
		GuiControl,, TheProgress, +2
		Sleep 500
		GuiControl,, TheProgress, +2
		MouseClick left, %Reasonx%, %Reasony%  ; Selects Adjust/Deny Reason
		GuiControl,, TheProgress, +2
		If (Comp = 1)
			{
				SendRaw, Comp. Sales
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Cost = 1)
			{
				Send, Cost
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Income = 1)
			{
				Send, Income
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Appraisal = 1)
			{
				Send, Appraisal
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Data = 1)
			{
				SendRaw, Data Correction
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Other = 1)
			{
				SendRaw, Other/See Note
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		Sleep 100
		GuiControl,, TheProgress, +2
		SendInput {ctrl down}s{ctrl up}
		Loop, 5
			{
				GuiControl,, TheProgress, +2
				Sleep 500
			}
		MouseClick left, %Notesx%, %Notesy%  ; Selects Notes Minor
		Loop, 3
			{
				GuiControl,, TheProgress, +2
				Sleep 500
			}
		MouseClick left, %NoteFieldx%, %NoteFieldy%  ; Selects Note Field
		GuiControl,, TheProgress, +2
		Sleep 500
		GuiControl,, TheProgress, +2
		SendInput %Notes%
		GuiControl,, TheProgress, +2
		Sleep 100
		GuiControl,, TheProgress, +2
		SendInput {Tab}
		GuiControl,, TheProgress, +2
		Sleep 100
		GuiControl,, TheProgress, +2
		SendInput {Space}
		GuiControl,, TheProgress, +2
		Sleep 100
		GuiControl,, TheProgress, +2
		SendInput {ctrl down}s{ctrl up}
		GuiControl,, TheProgress, +2
		Sleep 500
		GuiControl,, TheProgress, +2
		MouseClick left, %Summaryx%, %Summaryy%  ; Selects Summary Minor
		Loop, 7
			{
				GuiControl,, TheProgress, +2
				Sleep 500
			}
		MouseClick left, %Showx%, %Showy%  ; Selects Show/No Show/Phone
		If (Show = 1)
			{
				Send, Show
				GuiControl,, TheProgress, +2
			}
		else If (NoShow = 1)
			{
				Send, No Show
				GuiControl,, TheProgress, +2
			}
		else If (Phone = 1)
			{
				Send, Phone
				GuiControl,, TheProgress, +2
			}
		Sleep 500
		Mouseclick left, %DecisionByx%, %DecisionByy%
		Sleep 200
		Send, %FieldNumber%
		GuiControl,, TheProgress, +1
		Sleep 500
		GuiControl,, TheProgress, +2
		SendInput {ctrl down}s{ctrl up}
	}
}
else If (Vacant = 1)
{
	IF (Adjust = 1)
	{
		MouseClick left, %Realx%, %Realy% ;  Selects Real Major
		GuiControl,, TheProgress, +2
		Sleep 100
		MouseClick left, %Accountx%, %Accounty%  ;  Selects Account Minor
		GuiControl,, TheProgress, +2
		Sleep 100
		MouseClick left, %Lookupx%, %Lookupy%  ; Selects Account Lookup
		GuiControl,, TheProgress, +2
		Sleep 100
		GuiControl,, TheProgress, +2
		Send, %Account1%%Account2%%Account3%%Account4%
		GuiControl,, TheProgress, +2
		SendInput {Enter}
		GuiControl,, TheProgress, +2
		Loop, 7
			{
				GuiControl,, TheProgress, +2
				Sleep 500
			}
		MouseClick left, %LandValuex%, %LandValuey%  ; Selects Land Value
		GuiControl,, TheProgress, +2
		Sleep 50
		Send, %Proposed%
		GuiControl,, TheProgress, +2
		Sleep 500
		GuiControl,, TheProgress, +2
		SendInput {ctrl down}s{ctrl up}
		Loop, 5
			{
				GuiControl,, TheProgress, +2
				Sleep 500
			}
		MouseClick left, %Appealsx%, %Appealsy%  ; Selects Appeals Major
		GuiControl,, TheProgress, +2
		Sleep 50
		MouseClick left, %Detailx%, %Detaily%  ; Selects Details Minor
		GuiControl,, TheProgress, +2
		Sleep 50
		MouseClick left, %Lookupx%, %Lookupy%  ; Selects Account Lookup
		GuiControl,, TheProgress, +2
		Sleep 100	
		If (Num = 1)
			{
				Send, %TaxYear%00000%ProtestNumber%
				SendInput {Enter}
				Loop, 5
					{
						GuiControl,, TheProgress, +2
						Sleep 500
					}
			}
		else If (Num = 2)
			{
				Send, %TaxYear%0000%ProtestNumber%
				SendInput {Enter}
				Loop, 5
					{
						GuiControl,, TheProgress, +2
						Sleep 500
					}
			}
		else If (Num = 3)
			{
				Send, %TaxYear%000%ProtestNumber%
				SendInput {Enter}
				Loop, 5
					{
						GuiControl,, TheProgress, +2
						Sleep 500
					}
			}
		else If (Num = 4)
			{
				Send, %TaxYear%00%ProtestNumber%
				SendInput {Enter}
				Loop, 5
					{
						GuiControl,, TheProgress, +2  
						Sleep 500
					}
			}
		MouseClick left, %Decisionx%, %Decisiony%  ; Selects Appeal Decision
		GuiControl,, TheProgress, +1
		Send, Adjust
		GuiControl,, TheProgress, +1
		Sleep 500
		GuiControl,, TheProgress, +1
		MouseClick left, %Reasonx%, %Reasony%  ; Selects Adjust/Deny Reason
		GuiControl,, TheProgress, +1
		If (Comp = 1)
			{
				SendRaw, Comp. Sales
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Cost = 1)
			{
				SendRaw, Cost
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Income = 1)
			{
				SendRaw, Income
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Appraisal = 1)
			{
				SendRaw, Appraisal
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Data = 1)
			{
				SendRaw, Data Correction
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Other = 1)
			{
				SendRaw, Other/See Note
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		Sleep 100
		SendInput {ctrl down}s{ctrl up}
		Loop, 5
			{
				GuiControl,, TheProgress, +1
				Sleep 500
			}
		MouseClick left, %Notesx%, %Notesy%  ; Selects Notes Minor
		Loop, 3
			{
				GuiControl,, TheProgress, +1  
				Sleep 500
			}
		MouseClick left, %NoteFieldx%, %NoteFieldy%  ; Selects Note Field
		Sleep 500
		GuiControl,, TheProgress, +1
		SendInput %Notes%
		GuiControl,, TheProgress, +1
		Sleep 100
		SendInput {Tab}
		GuiControl,, TheProgress, +1
		SendInput {Space}
		Sleep 100
		GuiControl,, TheProgress, +1
		SendInput {ctrl down}s{ctrl up}
		Sleep 500
		MouseClick left, %Summaryx%, %Summaryy%  ; Selects Summary Minor
		Loop, 7
			{
				GuiControl,, TheProgress, +1
				Sleep 500
			}
		MouseClick left, %Showx%, %Showy%  ; Selects Show/No Show/Phone
		If (Show = 1)
			{
				Send, Show
				GuiControl,, TheProgress, +1
			}
		else If (NoShow = 1)
			{
				Send, No Show
				GuiControl,, TheProgress, +1
			}
		else If (Phone = 1)
			{
				Send, Phone
				GuiControl,, TheProgress, +1
			}
		Sleep 500
		MouseClick left, %DecisionByx%, %DecisionByy%
		Sleep 500
		Send, %FieldNumber%
		GuiControl,, TheProgress, +1
		Sleep 500
		GuiControl,, TheProgress, +1
		SendInput {ctrl down}s{ctrl up}
		GuiControl,, TheProgress, +1
	}
	else If (Deny = 1)
	{
		MouseClick left, %Appealsx%, %Appealsy%  ; Selects Appeals Major
		GuiControl,, TheProgress, +3
		Sleep 50
		GuiControl,, TheProgress, +3
		MouseClick left, %Detailx%, %Detaily%  ; Selects Details Minor
		GuiControl,, TheProgress, +3
		Sleep 50
		GuiControl,, TheProgress, +3
		MouseClick left, %Lookupx%, %Lookupy%  ; Selects Account Lookup
		GuiControl,, TheProgress, +3
		Sleep 100
		If (Num = 1)
			{
				Send, %TaxYear%00000%ProtestNumber%
				SendInput {Enter}
				Loop, 5
				{
					Sleep 500
					GuiControl,, TheProgress, +2
				}
			}
		else If (Num = 2)
			{
				Send, %TaxYear%0000%ProtestNumber%
				SendInput {Enter}
				Loop, 5
				{
					Sleep 500
					GuiControl,, TheProgress, +2
				}
			}
		else If (Num = 3)
			{
				Send, %TaxYear%000%ProtestNumber%
				SendInput {Enter}
				Loop, 5
				{
					Sleep 500
					GuiControl,, TheProgress, +2
				}
			}
		else If (Num = 4)
			{
				Send, %TaxYear%00%ProtestNumber%
				SendInput {Enter}
				Loop, 5
				{
					Sleep 500
					GuiControl,, TheProgress, +2
				}
			}
		MouseClick left, %Decisionx%, %Decisiony%  ; Selects Appeal Decision
		GuiControl,, TheProgress, +3
		Send, Deny
		GuiControl,, TheProgress, +3
		Sleep 500
		GuiControl,, TheProgress, +3
		MouseClick left, %Reasonx%, %Reasony%  ; Selects Adjust/Deny Reason
		GuiControl,, TheProgress, +3
		If (Comp = 1)
			{
				SendRaw, Comp. Sales
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Cost = 1)
			{
				SendRaw, Cost
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Income = 1)
			{
				SendRaw, Income
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Appraisal = 1)
			{
				SendRaw, Appraisal
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Data = 1)
			{
				SendRaw, Data Correction
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Other = 1)
			{
				SendRaw, Other/See Note
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		Sleep 100
		GuiControl,, TheProgress, +2
		SendInput {ctrl down}s{ctrl up}
		Loop, 5
			{
				Sleep 500
				GuiControl,, TheProgress, +2
			}
		MouseClick left, %Notesx%, %Notesy%  ; Selects Notes Minor
		Loop, 3
			{
				Sleep 500
				GuiControl,, TheProgress, +3
			}
		MouseClick left, %NoteFieldx%, %NoteFieldy%  ; Selects Note Field
		GuiControl,, TheProgress, +2
		Sleep 500
		GuiControl,, TheProgress, +2
		SendInput %Notes%
		GuiControl,, TheProgress, +2
		Sleep 100
		GuiControl,, TheProgress, +2
		SendInput {Tab}
		GuiControl,, TheProgress, +2
		Sleep 100
		GuiControl,, TheProgress, +2
		SendInput {Space}
		GuiControl,, TheProgress, +2
		Sleep 100	
		GuiControl,, TheProgress, +2
		SendInput {ctrl down}s{ctrl up}
		GuiControl,, TheProgress, +2
		Sleep 500
		MouseClick left, %Summaryx%, %Summaryy%  ; Selects Summary Minor
		Loop, 7
			{
				Sleep 500
				GuiControl,, TheProgress, +2
			}
		MouseClick left, %Showx%, %Showy%  ; Selects Show/No Show/Phone
		If (Show = 1)
			{
				Send, Show
				GuiControl,, TheProgress, +2
			}
		else If (NoShow = 1)
			{
				Send, No Show
				GuiControl,, TheProgress, +2
			}
		else If (Phone = 1)
			{
				Send, Phone
				GuiControl,, TheProgress, +2
			}
		Sleep 500
		GuiControl,, TheProgress, +2
		MouseClick left, %DecisionByx%, %DecisionByy%
		Sleep 500
		GuiControl,, TheProgress, +1
		Send, %FieldNumber%
		GuiControl,, TheProgress, +2
		Sleep 500
		GuiControl,, TheProgress, +1
		SendInput {ctrl down}s{ctrl up}
	}
	else If (Withdraw = 1)
	{
		MouseClick left, %Appealsx%, %Appealsy%  ; Selects Appeals Major
		GuiControl,, TheProgress, +3
		Sleep 50
		GuiControl,, TheProgress, +3
		MouseClick left, %Detailx%, %Detaily%  ; Selects Details Minor
		GuiControl,, TheProgress, +3
		Sleep 50
		GuiControl,, TheProgress, +3
		MouseClick left, %Lookupx%, %Lookupy%  ; Selects Account Lookup
		GuiControl,, TheProgress, +3
		Sleep 100	
		If (Num = 1)
			{
				Send, %TaxYear%00000%ProtestNumber%
				SendInput {Enter}
				Loop, 5
					{
						GuiControl,, TheProgress, +2
						Sleep 500
					}
			}
		else If (Num = 2)
			{
				Send, %TaxYear%0000%ProtestNumber%
				SendInput {Enter}
				Loop, 5
					{
						GuiControl,, TheProgress, +2
						Sleep 500
					}
			}
		else If (Num = 3)
			{
				Send, %TaxYear%000%ProtestNumber%
				SendInput {Enter}
				Loop, 5
					{
						GuiControl,, TheProgress, +2
						Sleep 500
					}
			}
		else If (Num = 4)
			{
				Send, %TaxYear%00%ProtestNumber%
				SendInput {Enter}
				Loop, 5
					{
						GuiControl,, TheProgress, +2
						Sleep 500
					}
			}
		MouseClick left, %Decisionx%, %Decisiony%  ; Selects Appeal Decision
		GuiControl,, TheProgress, +2
		Send, Withdrawn
		GuiControl,, TheProgress, +2
		Sleep 500
		GuiControl,, TheProgress, +2
		MouseClick left, %Reasonx%, %Reasony%  ; Selects Adjust/Deny Reason
		GuiControl,, TheProgress, +2
		If (Comp = 1)
			{
				SendRaw, Comp. Sales
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Cost = 1)
			{
				Send, Cost
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Income = 1)
			{
				Send, Income
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Appraisal = 1)
			{
				Send, Appraisal
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Data = 1)
			{
				SendRaw, Data Correction
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		else If (Other = 1)
			{
				SendRaw, Other/See Note
				SendInput {Tab}
				GuiControl,, TheProgress, +2
			}
		Sleep 100
		GuiControl,, TheProgress, +2
		SendInput {ctrl down}s{ctrl up}
		Loop, 5
			{
				GuiControl,, TheProgress, +3
				Sleep 500
			}
		MouseClick left, %Notesx%, %Notesy%  ; Selects Notes Minor
		Loop, 3
			{
				GuiControl,, TheProgress, +3
				Sleep 500
			}
		MouseClick left, %NoteFieldx%, %NoteFieldy%  ; Selects Note Field
		GuiControl,, TheProgress, +2
		Sleep 500
		GuiControl,, TheProgress, +2
		SendInput %Notes%
		GuiControl,, TheProgress, +2
		Sleep 100
		GuiControl,, TheProgress, +2
		SendInput {Tab}
		GuiControl,, TheProgress, +2
		Sleep 100
		GuiControl,, TheProgress, +2
		SendInput {Space}
		GuiControl,, TheProgress, +2
		Sleep 100
		GuiControl,, TheProgress, +2
		SendInput {ctrl down}s{ctrl up}
		GuiControl,, TheProgress, +2
		Sleep 500
		GuiControl,, TheProgress, +2
		MouseClick left, %Summaryx%, %Summaryy%  ; Selects Summary Minor
		Loop, 7
			{
				GuiControl,, TheProgress, +3
				Sleep 500
			}
		MouseClick left, %Showx%, %Showy%  ; Selects Show/No Show/Phone
		If (Show = 1)
			{
				Send, Show
				GuiControl,, TheProgress, +2
			}
		else If (NoShow = 1)
			{
				Send, No Show
				GuiControl,, TheProgress, +2
			}
		else If (Phone = 1)
			{
				Send, Phone
				GuiControl,, TheProgress, +2
			}
		Sleep 500
		Mouseclick left, %DecisionByx%, %DecisionByy%
		Sleep 200
		Send, %FieldNumber%
		GuiControl,, TheProgress, +2
		Sleep 500
		GuiControl,, TheProgress, +2
		SendInput {ctrl down}s{ctrl up}
	}	
}
return
}

}