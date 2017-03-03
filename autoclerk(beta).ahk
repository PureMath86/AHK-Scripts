;
;   	Auto-Clerk Script (BETA)
;
;
; 	Developer: Bryan Arnold 
;
;	Created: 10/1/2016
;
; 	Last Updated: 10/20/2016
;
;

;------------------------------------------------------------;
; Reminders:
;
; ^ = Ctrl
; + = Shift
; ! = Alt
; # = Win
;------------------------------------------------------------;

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
 #Warn  ; Enable warnings to assist with detecting common errors.
 SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
 SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance,Force

; #NoTrayIcon


;
; Get File-Path Function
;

FileSelectFile, SelectedFile, 3, , Open a file, Excel Documents (*.csv; *.xlsm; *xlsx)
if SelectedFile =
    MsgBox, The user didn't select anything.
else
    MsgBox, The user selected the following:`n%SelectedFile%

;
; File-Path
;
 
; SelectedFile := "`n%SelectedFile%"

oWorkbook := ComObjGet(SelectedFile) ; 

/*
;
; File-Path
;

FilePath := "C:\Users\Frey\Desktop\Bryan\Book1.xlsx" ; 

oWorkbook := ComObjGet(FilePath) ; 
*/

;------------------------------------------------------------;
;
; Hard Quit / Forced Quit
;
;------------------------------------------------------------;

^+0::exitapp

;------------------------------------------------------------;
;
; Launch
;
;------------------------------------------------------------;

^+1::
{
SendLevel 0 ; Disables Hot-Key/String Calls						<----------------------------- Level Default

;------------------------------------------------------------;
;
; Automate 1 Item PPO
;
;------------------------------------------------------------;

WinActivate, http://www.personable.com/?sp=100&w=1337&h=986&m=&key=55BE0968-7E19-455A-A547-1B14CC77DD08&cb=& - Internet Explorer ahk_class IEFrame

Sleep, 500

;
; Clear out first info line Delivery Date, FOB, etc.
;

Send, {Backspace}{Tab}{Backspace}{Tab}{Backspace}{Tab}{Backspace}{Tab}{Backspace}{Tab}{Backspace}{Tab}{Backspace}{Tab}{Backspace}

Sleep, 100

;
; Go back to beginning of first info line
;

Send, {LShift Down}{Tab}{Tab}{Tab}{Tab}{Tab}{Tab}{Tab}{LShift Up}

;
; Get Delivery Date
;

Value := oWorkbook.Sheets(7).Range("E2").Value ; 

Sleep, 100

;
; Input Delivery Date
;

Send % Value

Send, {Tab}


;
; Get Supplier Terms
;

Value := oWorkbook.Sheets(7).Range("F2").Value ; 

Sleep, 100

;
; Input Supplier Terms
;

Send % Value

Send, {Tab}

;
; Input Created By
;

Send % "BMA"

Send, {Tab}

;
; Get Supplier Num
;

Value := oWorkbook.Sheets(7).Range("G2").Value ; 

Sleep, 100

;
; Input Supplier Num
;

if Value is number
{
Value := Floor(Value)

  if (Value = 0)
  {
  Send % "NEED"
  }
    else
    {
    Send % Value
    }
}
  else
  {
  Send % Value
  }

Send, {Tab}

;
; Get Ship Via
;

Value := oWorkbook.Sheets(7).Range("H2").Value ;

Sleep, 100

if ( Value = "NEED")
{
 ;Do Nothing
}
  else if ( Value = "Need")
  {
  Value := "NEED"
  }
    else if ( Value = "need")
    {
    Value := "NEED"
    }
      else
      {
        ; Do Nothing 
      }

;
; Input Ship Via
;

Send % Value

Send, {Tab}

;
; Get Trans PO
;

if ( Value = "NEED")
{
 ; Do Nothing
}
  else if (Value = "FOB Pickup")
  {
  Value := "FOB"
  }
    else if (Value = "DLVD")
    {
    Value := "NA"
    }
      else
      {
      Value := oWorkbook.Sheets(7).Range("I2").Value ; 
      }

;
; Input Trans PO
;

if Value is number
{
Value := Floor(Value)
  if (Value = 0)
  {
   Send % "NEED"
  }
    else
    {
    Send % Value
    }
}
  else
  {
  Send % Value
  }

Send, {Tab}

;
; Get SO Num ID
;

Value := oWorkbook.Sheets(7).Range("J2").Value ;

Sleep, 100							;
								; Note:
Value := Floor(Value) ;  <------------------------------------- ;   If String (or empty String), 
								;     Value := 0 
;								;   via Floor function.
; Input SO Num ID						;
;

if (Value = 0)
{
Send % "NEED"									
}
  else
  {
  Send % Value
  }

Send, {Tab}

;
; Get Ship Date
;

Value := oWorkbook.Sheets(7).Range("K2").Value ; 

Sleep, 100

;
; Input Ship Date
;

Send % Value

Send, {Tab}

;------------------------------------------------------------;
;
; Get and Input - Item/Count/Qty/UnitCost/Origin
;
;------------------------------------------------------------;

SendLevel 1 ; Enables 1st Level Hot-Key/String Calls					<----------------------------- Level Change

Item := oWorkbook.Sheets(7).Range("L2").Value ;    Item
Count := oWorkbook.Sheets(7).Range("M2").Value ;   Count

Sleep, 100

Count := Floor(Count) ; Make Integer

Send, %Item%

Sleep, 500

Send, %Count%

Sleep, 500

Send, {Space}

Sleep, 100

SendLevel 0 ; Disables Hot-Key/String Calls						<----------------------------- Level Default

Send, {Tab}{Tab}

Value := oWorkbook.Sheets(7).Range("N2").Value ;   Qty

Sleep, 100

Send % Floor(Value)

Send, {Tab}

Value := oWorkbook.Sheets(7).Range("O2").Value ;   UnitCost

Sleep, 100

Value := Floor(Value) ; <--------------------------------------------------------------------- String Killer

if (Value = 0)
{
Send, 1
}
  else
  {
   Send % FloorDecimal(oWorkbook.Sheets(7).Range("O2").Value) 
  }

Send, {Tab}{Tab}

Value := oWorkbook.Sheets(7).Range("P2").Value ;   Origin

Sleep, 100

Send % Value

Send, {Tab}{Tab}{Tab}{Tab}{Down}

;------------------------------------------------------------;
;
; Pallet - Handler
;
;------------------------------------------------------------;

;
; Get Blonde Pallets
;

Value := oWorkbook.Sheets(7).Range("Q2").Value ; 

Sleep, 100

;
; Input Blonde Pallets
;

Value := Floor(Value) ; <--------------------------------------------------------------------- String Killer

If (Value = 0)
{
 ; Do Nothing
}
    else
    {
     Send % Value

     Send % " Blonde Pallets"

     Send, {Enter}
     }

;
; Get Chep Pallets
;

Value := oWorkbook.Sheets(7).Range("R2").Value ; 

Sleep, 100

;
; Input Chep Pallets
;

Value := Floor(Value) ; <--------------------------------------------------------------------- String Killer

If (Value = 0)
{
 ; Do Nothing
}
    else
    {
    Send % Value

    Send % " Chep Pallets"

    Send, {Enter}
    }

;
; Get Peco Pallets
;

Value := oWorkbook.Sheets(7).Range("S2").Value ; 

Sleep, 100

;
; Input Peco Pallets
;


Value := Floor(Value) ; <--------------------------------------------------------------------- String Killer

If (Value = 0)
{
 ; Do Nothing
}
    else
    {
    Send % Value

    Send % " Peco Pallets"

    Send, {Enter}
    }

Sleep, 100

Send, {Down}

;------------------------------------------------------------;
;
; CPO - Handler
;
;------------------------------------------------------------;

;
; Get CPO # 
;

Value := oWorkbook.Sheets(7).Range("T2").Value ; 

Sleep, 100

;
; Input CPO #
;

Send % "CPO "

Send, {#}{Space}

if Value is number
{
Value := Floor(Value)
  if (Value = 0)
  {
  Send % "NA"
  }
    else
    {
    Send % Value
    }
}
  else
  {
  Send % Value
  }

Sleep, 100

} ; <--------------------------------------------------------- End of Auto 1 Item PPO (Body)
Return























^+2::
{
SendLevel 0 ; Disables Hot-Key/String Calls						<----------------------------- Level Default

;------------------------------------------------------------;
;
; Automate 1 Item PPO
;
;------------------------------------------------------------;

WinActivate, http://www.personable.com/?sp=100&w=1337&h=986&m=&key=55BE0968-7E19-455A-A547-1B14CC77DD08&cb=& - Internet Explorer ahk_class IEFrame

Sleep, 500

;
; Clear out first info line Delivery Date, FOB, etc.
;

Send, {Backspace}{Tab}{Backspace}{Tab}{Backspace}{Tab}{Backspace}{Tab}{Backspace}{Tab}{Backspace}{Tab}{Backspace}{Tab}{Backspace}

Sleep, 100

;
; Go back to beginning of first info line
;

Send, {LShift Down}{Tab}{Tab}{Tab}{Tab}{Tab}{Tab}{Tab}{LShift Up}

;
; Get Delivery Date
;

Value := oWorkbook.Sheets(7).Range("E2").Value ; 

Sleep, 100

;
; Input Delivery Date
;

Send % Value

Send, {Tab}


;
; Get Supplier Terms
;

Value := oWorkbook.Sheets(7).Range("F2").Value ; 

Sleep, 100

;
; Input Supplier Terms
;

Send % Value

Send, {Tab}

;
; Input Created By
;

Send % "BMA"

Send, {Tab}

;
; Get Supplier Num
;

Value := oWorkbook.Sheets(7).Range("G2").Value ; 

Sleep, 100

;
; Input Supplier Num
;

if Value is number
{
Value := Floor(Value)

  if (Value = 0)
  {
  Send % "NEED"
  }
    else
    {
    Send % Value
    }
}
  else
  {
  Send % Value
  }

Send, {Tab}

;
; Get Ship Via
;

Value := oWorkbook.Sheets(7).Range("H2").Value ;

Sleep, 100

if ( Value = "NEED")
{
 ;Do Nothing
}
  else if ( Value = "Need")
  {
  Value := "NEED"
  }
    else if ( Value = "need")
    {
    Value := "NEED"
    }
      else
      {
        ; Do Nothing 
      }

;
; Input Ship Via
;

Send % Value

Send, {Tab}

;
; Get Trans PO
;

if ( Value = "NEED")
{
 ; Do Nothing
}
  else if (Value = "FOB Pickup")
  {
  Value := "FOB"
  }
    else if (Value = "DLVD")
    {
    Value := "NA"
    }
      else
      {
      Value := oWorkbook.Sheets(7).Range("I2").Value ; 
      }

;
; Input Trans PO
;

if Value is number
{
Value := Floor(Value)
  if (Value = 0)
  {
   Send % "NEED"
  }
    else
    {
    Send % Value
    }
}
  else
  {
  Send % Value
  }

Send, {Tab}

;
; Get SO Num ID
;

Value := oWorkbook.Sheets(7).Range("J2").Value ;

Sleep, 100							;
								; Note:
Value := Floor(Value) ;  <------------------------------------- ;   If String (or empty String), 
								;     Value := 0 
;								;   via Floor function.
; Input SO Num ID						;
;

if (Value = 0)
{
Send % "NEED"									
}
  else
  {
  Send % Value
  }

Send, {Tab}

;
; Get Ship Date
;

Value := oWorkbook.Sheets(7).Range("K2").Value ; 

Sleep, 100

;
; Input Ship Date
;

Send % Value

Send, {Tab}

;------------------------------------------------------------;
;
; Get and Input - Item/Count/Qty/UnitCost/Origin
;
;------------------------------------------------------------;

SendLevel 1 ; Enables 1st Level Hot-Key/String Calls					<----------------------------- Level Change

Item := oWorkbook.Sheets(7).Range("L2").Value ;    Item
Count := oWorkbook.Sheets(7).Range("M2").Value ;   Count

Sleep, 100

Count := Floor(Count) ; Make Integer

Send, %Item%

Sleep, 500

Send, %Count%

Sleep, 500

Send, {Space}

Sleep, 100

SendLevel 0 ; Disables Hot-Key/String Calls						<----------------------------- Level Default

Send, {Tab}{Tab}

Value := oWorkbook.Sheets(7).Range("N2").Value ;   Qty

Sleep, 100

Send % Floor(Value)

Send, {Tab}

Value := oWorkbook.Sheets(7).Range("O2").Value ;   UnitCost

Sleep, 100

Value := Floor(Value) ; <--------------------------------------------------------------------- String Killer

if (Value = 0)
{
Send, 1
}
  else
  {
   Send % FloorDecimal(oWorkbook.Sheets(7).Range("O2").Value) 
  }

Send, {Tab}{Tab}

Value := oWorkbook.Sheets(7).Range("P2").Value ;   Origin

Sleep, 100

Send % Value

Send, {Tab}{Tab}{Tab}{Tab}{Tab}{Down}

;------------------------------------------------------------;
;
; Pallet - Handler
;
;------------------------------------------------------------;

;
; Get Blonde Pallets
;

Value := oWorkbook.Sheets(7).Range("Q2").Value ; 

Sleep, 100

;
; Input Blonde Pallets
;

Value := Floor(Value) ; <--------------------------------------------------------------------- String Killer

If (Value = 0)
{
 ; Do Nothing
}
    else
    {
     Send % Value

     Send % " Blonde Pallets"

     Send, {Enter}
     }

;
; Get Chep Pallets
;

Value := oWorkbook.Sheets(7).Range("R2").Value ; 

Sleep, 100

;
; Input Chep Pallets
;

Value := Floor(Value) ; <--------------------------------------------------------------------- String Killer

If (Value = 0)
{
 ; Do Nothing
}
    else
    {
    Send % Value

    Send % " Chep Pallets"

    Send, {Enter}
    }

;
; Get Peco Pallets
;

Value := oWorkbook.Sheets(7).Range("S2").Value ; 

Sleep, 100

;
; Input Peco Pallets
;


Value := Floor(Value) ; <--------------------------------------------------------------------- String Killer

If (Value = 0)
{
 ; Do Nothing
}
    else
    {
    Send % Value

    Send % " Peco Pallets"

    Send, {Enter}
    }

Sleep, 100

Send, {Down}

;------------------------------------------------------------;
;
; CPO - Handler
;
;------------------------------------------------------------;

;
; Get CPO # 
;

Value := oWorkbook.Sheets(7).Range("T2").Value ; 

Sleep, 100

;
; Input CPO #
;

Send % "CPO "

Send, {#}{Space}

if Value is number
{
Value := Floor(Value)
  if (Value = 0)
  {
  Send % "NA"
  }
    else
    {
    Send % Value
    }
}
  else
  {
  Send % Value
  }

Sleep, 100

} ; <--------------------------------------------------------- End of Auto 1 Item PPO (Body)
Return











































^+9::
{
SendLevel 0 ; Disables Hot-Key/String Calls						<----------------------------- Level Default

;------------------------------------------------------------;
;
; Automate 1 Drop-Off TPO
;
;------------------------------------------------------------;

WinActivate, http://www.personable.com/?sp=100&w=1337&h=986&m=&key=55BE0968-7E19-455A-A547-1B14CC77DD08&cb=& - Internet Explorer ahk_class IEFrame

Sleep, 500

;
; Clear out first info line Delivery Date, Created By, etc.
;

Send, {Backspace}{Tab}{Backspace}{Tab}{Backspace}{Tab}{Backspace}{Tab}{Backspace}{Tab}{Backspace}

Sleep, 100

;
; Go back to beginning of first info line
;

Send, {LShift Down}{Tab}{Tab}{Tab}{Tab}{Tab}{LShift Up}

;
; Get Delivery Date
;

Value := oWorkbook.Sheets(8).Range("E2").Value ; 

Sleep, 100

;
; Input Delivery Date
;

Send % Value

Send, {Tab}

Sleep, 100

;
; Input Created By
;

Send % "BMA"

Send, {Tab}

;
; Get Supplier Num
;

Value := oWorkbook.Sheets(8).Range("F2").Value ; 

Sleep, 100

;
; Input Supplier Num
;

if Value is number
{
Value := Floor(Value)

  if (Value = 0)
  {
  Send % "NEED"
  }
    else
    {
     Send % Value
    }
}
  else
  {
   Send % Value
  }                                                  

Send, {Tab}

;
; Get Prod PO
;

Value := oWorkbook.Sheets(8).Range("G2").Value ; 

Sleep, 100

;
; Input Prod PO
;

; This should always be a number or say "NEED"...

Value := Floor(Value) ; <---------------------------------------------------------------- String Killer

if (Value = 0)
{
Send % "NEED"
}
  else
  {
  Send % Value
  }

Send, {Tab}

;
; Get SO Num ID
;

Value := oWorkbook.Sheets(8).Range("H2").Value ;

Sleep, 100							;
								; Note:
Value := Floor(Value) ;  <------------------------------------- ;   If String (or empty String), 
								;     Value := 0 
;								;   via Floor function.
; Input SO Num ID						;
;

if (Value = 0)
{
Send % "NEED"									
}
  else
  {
  Send % Value
  }

Send, {Tab}

;
; Get Ship Date
;

Value := oWorkbook.Sheets(8).Range("I2").Value ; 

Sleep, 100

;
; Input Ship Date
;

Send % Value

Send, {Tab}

;------------------------------------------------------------;
;
; Get and Input - FrtRate/EstUnld/Reefer/Info
;
;------------------------------------------------------------;

Send % "FREIGHT*"

Value :=FloorDecimal(oWorkbook.Sheets(8).Range("J2").Value) ;   Freight Rate

Sleep, 100

Send, {Space}

Sleep, 500

Send, {Tab}{Tab}

Sleep, 100

Send, 1

Sleep, 100

Send, {Tab}

Send % Value

Sleep, 100

Value := oWorkbook.Sheets(8).Range("K2").Value ;   Estimated Unloading       ;
									     ; Note:
Value := Floor(Value) ; <--------------------------------------------------- ;   If String (or empty String), 
									     ;     Value := 0 
if (Value != 0)								     ;   via Floor function.
{
Send, {Tab}{Tab}{Tab}{Tab}
	
Sleep, 100

Send % "FREIGHT*:UNLOADING FEE"

Sleep, 500

Send, {Tab}{Tab}

Sleep, 100

Send, 1

Sleep, 100

Send, {Tab}

Send % FloorDecimal(oWorkbook.Sheets(8).Range("K2").Value)

Sleep, 100

Send, {Tab}{Tab}{Tab}{Tab}{Tab}{Down}
}
  else
  {
  Send, {LShift Down}{Tab}{Tab}{LShift Up}

  Sleep, 100

  Send, {Right 20}

  Sleep, 100

  Send, {Space}{+}

  Send % " Unloading"

  Sleep, 100

  Send, {Down}{Down}
  }

SendLevel 1 ; Enables 1st Level Hot-Key/String Calls					<----------------------------- Level Change

Value := oWorkbook.Sheets(8).Range("P2").Value ;   Loading Location

Send, %Value%

Sleep, 100

Send % "LL" ; This is the Syntax for Loading Locations. See: Hot-Strings in Script1. 	<----------------------------- NOTE

Sleep, 1500 ; It takes it a while to spit out the keys...                               <----------------------------- .str

SendLevel 0 ; Disables Hot-Key/String Calls						<----------------------------- Level Default

Send, {Enter}

Sleep, 100

Value := oWorkbook.Sheets(8).Range("N2").Value ;   Appt Time

;
; Get the Correct Time
;

if Value is number
{
  if (Value < 1)
  {
  Value := 24*60*60*Value ; Convert to Seconds
  }
    else
    {
     ; Do Nothing
    }

Value := Floor(Value)

  if (Value = 0)
  {
  Value := "NEED"
  }
    else
    {
    Value := FormatSeconds(Value)
    }
}
  else
  {
   ; Do Nothing
  }

;
; Input the Correct Time 
;

Send % "Appt "

Sleep, 100

Send % Value

Value := oWorkbook.Sheets(8).Range("O2").Value ;   Confirmation #

Send, {Enter}

Sleep, 100

Send % "Confirmation "

Send, {#}{Space}

if Value is number
{
Value := Floor(Value)
  if (Value = 0)
  {
  Send % "NA"
  }
    else
    {
    Send % Value
    }
}
  else
  {
  Send % Value
  }

/*

Value := oWorkbook.Sheets(8).Range("L2").Value ;   REEFER (+loading) Hot-String This... Put in Script1

*/
} <--------------------------------------------------------- End of Auto 1 Drop TPO (Body)
Return











;
; Time Function - Seconds |---> H:MM (w/ AM or PM marker)
;

FormatSeconds(NumberOfSeconds)  ; Convert the specified number of seconds to h:mm tt format.
{
    time = 19990101  ; *Midnight* of an arbitrary date.
    time += %NumberOfSeconds%, seconds

    FormatTime, hmm, %time%, h:mm tt
    return hmm
}

;
; Function for Dollar Amounts $ ... d . c c
;

FloorDecimal(num) {

  num:=Floor(num*100)
  SetFormat Float, 0.2
  return num/100

}
