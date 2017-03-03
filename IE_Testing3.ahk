
/*

Auto-Log-In for QB (Personable)

*/

navOpenNewForegroundTab = 65536
navOpenInBackgroundTab = 4096
navOpenInNewWindow = 1
navOpenInNewTab = 2048

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
 ;#Warn  ; Enable warnings to assist with detecting common errors.
 SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
 SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance,Force

;------------------------------------------------------------;
; Reminders:
;
; ^ = Ctrl
; + = Shift
; ! = Alt
; # = Win
;------------------------------------------------------------;


InputBox, Input1q, Auto-Log-In, Please enter the Personable e-mail address:, , 640, 480
if ErrorLevel
    MsgBox, CANCEL was pressed.
else
{

InputBox, Input2q, Auto-Log-In, Please enter the Personable password:, HIDE , 640, 480
if ErrorLevel
    MsgBox, CANCEL was pressed.
else
{

wb := ComObjCreate("InternetExplorer.Application") ; create a IE instance
IeHwnd := wb.Hwnd
wb.Visible := True

WinMaximize, A ; Maximize the Active Window
;WinMaximize, ahk_class IEFrame ; Maximizes the Web Browser

;
; tab 1
;
wb.Navigate("http://www.personable.com/login.asp")

While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
   Sleep, 10
   
;----------------------------------------------------------QB-Log-------------------------------------------------------------;

;Extra Wait
Sleep, 500

QBLogin(Input1q,Input2q)

;Extra Wait
Sleep, 500

;----------------------------------------------------------ASIDE-------------------------------------------------------------;
;
;  Fill the E-mail / P-word with Gibberish
;

Loop, 12 ;  Loop to be thorough --OVERKILL!!!
{
goop1 := Rand(9999,99999999)
goop2 := Rand(9999,99999999)
goop3 := Rand(9999,99999999)
goop := Rand(15,150)
; Replaced
Input1q := "%goop2%%goop%%goop1%%goop3%"
}

Loop, 13   ;  Loop to be thorough --OVERKILL!!!
{
goop1 := Rand(9999,99999999)
goop2 := Rand(9999,99999999)
goop3 := Rand(9999,99999999)
goop := Rand(15,150)
; Replaced
Input2q := "%goop3%%goop%%goop1%%goop2%"
}

;
;  Kill the E-mail / P-word and gibberish
;

goop1 :=    ; hard reset
goop2 :=    ; hard reset
goop3 :=    ; hard reset
goop :=    ; hard reset
Input1q :=    ; hard reset 
Input2q :=    ; hard reset
;-----------------------------------------------------------------------------------------------------------------------------;

} }  ;  <-------------------------------------------------------------------------------------------- End of Line ***dAft pUnk***



;
;  User Defined Functions (UDF)
;

New_IE(Url) {
	For wb in ComObjCreate("Shell.Application").Windows
	{
		If instr( wb.LocationUrl, URL) && InStr( wb.FullName, "iexplore.exe" )
			return wb
		else
			continue
	}
}

ClickLink(PXL,Text=""){
ComObjError(false)
Links := PXL.Document.Links
Loop % Links.Length
   If (Links[A_Index-1].InnerText = Text ) { ; search for Text
      Links[A_Index-1].Click() ;click it when you find it
      break
   }
ComObjError(True)
}

TabR(){
Sleep, 100
Send, ^{tab}
Sleep, 100
}

TabL(){
Sleep, 100
Send, ^+{tab}
Sleep, 100
}

SSLogin(Input1,Input2){
Send %Input1%
Sleep, 500
Send, {tab}
Sleep, 100
Send %Input2%
Sleep, 500
Send, {enter}	
}

QBLogin(Input1,Input2){
; Move Cursor to Page
Loop, 3
{
	Send, {F6}
	Sleep, 100
}
Send, {tab}
Sleep, 100
; E-mail
Send %Input1%
Sleep, 500
Send, {tab}
Sleep, 100
; P-word
Send %Input2%
Sleep, 500
; Go to Login Button
Send, {tab}
Sleep, 100
Send, {enter}	
}

; Close the Current Tab
KillTab(){
	Sleep, 100
	Send, ^{f4} 
	Sleep, 100
}

;
; RANDOM FUNCTIONS
;

; Random Numbers (Integers)
Rand( a=0.0, b=1 ) {
   IfEqual,a,,Random,,% r := b = 1 ? Rand(0,0xFFFFFFFF) : b
   Else Random,r,a,b
   Return r
}





