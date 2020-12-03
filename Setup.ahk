#SingleInstance, Force

;-----GUI Stuff----------
Gui, New, +AlwaysOnTop, Setup Tool
Gui, Font, c39ff14
Gui, Color, c000000
Gui, Show, w150 h170    
Gui, Add, Text, x10 y10, Tax Year
Gui, Add, Edit, cBlack x10 y30 w50 number limit4 vTaxYear,
Gui, Add, Text, x10 y60, Field Number
Gui, Add, Edit, cBlack x10 y80 w50 number limit3 vFieldNumber, 
Gui, Add, Text, x10 y110, Press to Start Setup
Gui, Add, Button, x10 y130 gStart, Start Setup

return

GuiClose:
    ExitApp

return

Start:

Gui, Submit, NoHide
Gui, Minimize

FileDelete, C:\Users\%A_UserName%\Documents\ahk\coordinates.ahk

;---Write Tax Year and Field Number to Coordinates file--------
FileAppend, TaxYear := %TaxYear%`nFieldNumber := %FieldNumber%`n, C:\Users\%A_UserName%\Documents\ahk\coordinates.ahk

;-----Get Monitor Info---------
SysGet, MonPrim, MonitorPrimary
If (MonPrim = 1)
    Mo = 2
else 
    Mo = 1

SysGet, Mon, Monitor, %Mo%

IF (Mo = 2) 
    {
        MonS := (MonLeft + 50)
        MonT := (MonTop + 50)
    }
else if (Mo = 1) 
    {
        MonS := (MonRight - 350)
        MonT := (MonTop + 50)
    }

;----Populate and Loop Through Array, Writing Coordinates to File-----------------
Array := ["Lookup", "Real", "Account", "LandValue", "DefaultApproach", "Improvements", "Reconciled", "ImpValue", "TotalValue", "Appeals", "Detail", "Decision", "Reason", "Notes", "NoteField", "Summary", "Show", "DecisionBy"]
MsgBox % Array.Length()
Loop 
    {
    if (A_Index <= Array.Length()) 
        {
            thing := Array[A_Index]
            WinActivate, RealWare
            WinMaximize, RealWare
            SplashImage, C:\Users\%A_UserName%\Documents\ahk\images\%thing%.jpg, M x%MonS% y%MonT%, Move Mouse to %thing% and Right Click, 
            Loop 
                {
                    key := GetKeyState("RButton")
                    if (key = 1) 
                        {    
                            MouseGetPos, x, y
                            FileAppend, %thing%x := %x%`n%thing%y := %y%`n, C:\Users\%A_UserName%\Documents\ahk\coordinates.ahk
                            sleep 500
                            break
                        } 
                    else if (key = 0)
                        continue
                }  
        } 
    else if (A_Index > Array.Length()) 
        {
            MsgBox, Setup Complete
            SplashImage, Off
            ExitApp
            break
        }
    }
return