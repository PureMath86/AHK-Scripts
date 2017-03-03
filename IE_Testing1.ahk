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

InputBox, Input2q, Auto-Log-In, Please enter the  Personable password:, HIDE , 640, 480
if ErrorLevel
    MsgBox, CANCEL was pressed.
else
{

InputBox, Input1s, Auto-Log-In, Please enter the SS e-mail address:, , 640, 480
if ErrorLevel
    MsgBox, CANCEL was pressed.
else
{
	
InputBox, Input2s, Auto-Log-In, Please enter the SS password:, HIDE , 640, 480
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
wb.Navigate("https://www.securesheet-cloud.com/register/login.aspx")

While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
   Sleep, 10

;
; tab 2
;
wb.Navigate("about:blank", navOpenInNewTab)

while !wb2
	wb2 := New_IE("about:blank")

wb2.Navigate("https://www.securesheet-cloud.com/register/login.aspx")

While wb2.readyState != 4 || wb2.document.readyState != "complete" || wb2.busy ; wait for the page to load
   Sleep, 10

;
; tab 3
;
wb.Navigate("about:blank", navOpenInNewTab)

while !wb3
	wb3 := New_IE("about:blank")

wb3.Navigate("https://www.securesheet-cloud.com/register/login.aspx")

While wb3.readyState != 4 || wb3.document.readyState != "complete" || wb3.busy ; wait for the page to load
   Sleep, 10

;
; tab 4
;
wb.Navigate("about:blank", navOpenInNewTab)

while !wb4
	wb4 := New_IE("about:blank")

wb4.Navigate("https://www.securesheet-cloud.com/register/login.aspx")

While wb4.readyState != 4 || wb4.document.readyState != "complete" || wb4.busy ; wait for the page to load
   Sleep, 10

;
; Main Body
;

; Go to the First Tab
TabR()

Loop, 4
{
if (a_index =1)
	{
	While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
	Sleep, 10
	
	SSLogin(Input1s,Input2s)
	
	TabR()
	}
else
	{
	While wb%a_index%.readyState != 4 || wb%a_index%.document.readyState != "complete" || wb%a_index%.busy ; wait for the page to load
	Sleep, 10
	
	SSLogin(Input1s,Input2s)
	
	TabR()
	}
}

; Back to the First Tab

; Scroll Down
Loop, 4
{
if (a_index =1)
	{
	While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
	Sleep, 10
	
	Send,  {PGDN}   ;  Scroll Down (in case the link is not visible)

	TabR()
	}
else
	{
	While wb%a_index%.readyState != 4 || wb%a_index%.document.readyState != "complete" || wb%a_index%.busy ; wait for the page to load
	Sleep, 10
	
	Send,  {PGDN}   ;  Scroll Down (in case the link is not visible)

	TabR()
	}
}

; Back to the First Tab

; Go To Master Current SS
Loop, 4
{
if (a_index =1)
	{
	While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
	Sleep, 10
	
	ClickLink(wb,Text:="Master Current SS")  ;  
	
	TabR()
	}
else
	{
	While wb%a_index%.readyState != 4 || wb%a_index%.document.readyState != "complete" || wb%a_index%.busy ; wait for the page to load
	Sleep, 10
	
	ClickLink(wb%a_index%,Text:="Master Current SS")  ;  
	
	TabR()
	}
}

; Back to the First Tab

; Sort Ship Dates
Loop, 4
{
if (a_index =1)
	{
	While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
	Sleep, 10
	
	ClickLink(wb,Text:="D")  ;  Sorts the Ship Dates
	
	TabR()
	}
else
	{
	While wb%a_index%.readyState != 4 || wb%a_index%.document.readyState != "complete" || wb%a_index%.busy ; wait for the page to load
	Sleep, 10

	ClickLink(wb%a_index%,Text:="D")  ;  Sorts the Ship Dates
	
	TabR()
	}
}

; Back to the First Tab


;----------------------------------------------------------QB-Log-------------------------------------------------------------;
;
; tab (temporary)
;
wb.Navigate("about:blank", navOpenInNewTab)

while !wb1
	wb1 := New_IE("about:blank")

wb1.Navigate("http://www.personable.com/login.asp")

While wb1.readyState != 4 || wb1.document.readyState != "complete" || wb1.busy ; wait for the page to load
   Sleep, 10

;Extra Wait
Sleep, 500

QBLogin(Input1q,Input2q)

;----------------------------------------------------------ASIDE-------------------------------------------------------------;
;
;  Fill the E-mail / P-word with Gibberish
;

Loop, 12  ;  Loop to be thorough --OVERKILL!!!
{
goop1 := Rand(9999,99999999)
goop2 := Rand(9999,99999999)
goop3 := Rand(9999,99999999)
goop := Rand(15,150)
; Replaced
Input1s := "%goop1%%goop2%%goop%%goop3%"
}

Loop, 13   ;  Loop to be thorough --OVERKILL!!!
{
goop1 := Rand(9999,99999999)
goop2 := Rand(9999,99999999)
goop3 := Rand(9999,99999999)
goop := Rand(15,150)
; Replaced
Input2s := "%goop1%%goop%%goop3%%goop2%"
}

Loop, 14  ;  Loop to be thorough --OVERKILL!!!
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

Input1s :=    ; hard reset 
Input2s :=    ; hard reset
goop1 :=    ; hard reset
goop2 :=    ; hard reset
goop3 :=    ; hard reset
goop :=    ; hard reset
Input1q :=    ; hard reset 
Input2q :=    ; hard reset
;-----------------------------------------------------------------------------------------------------------------------------;




/*
; Note: May be subject to change upon updating.
AM := "dgFH__ctl2_D39" ; AM Column ID 
Search := "txtSearch" ; Search Entry Field ID

wb.document.all.%AM%.value := "NA" ; Set DropDown box to NA
wb.document.all.%Search%.value := "Cancel" ; Set the Search Box to Cancel

; Function for Later ---^
SetTo(Input1,Input2){
wb.document.all.%Input1%.value := "%Input2%"
}

Send, ^{f4} ; close the current tab
Send, {f6} ; select the URL
ClickLink(wb,Text:="Search")  ;  Search in Tab 1
ClickLink(wb,Text:="Print Page")  ;  Get the Copy page for Tab 1
*/


;
;  ...to be continued...
;

} } } } ;  <-------------------------------------------------------------------------------------------- End of Line ***dAft pUnk***



;
; Loop (switching between tabs) Template
;

/*

Loop, 4
{
if (a_index =1)
	{
	While wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ; wait for the page to load
	Sleep, 10
	
	}
else
	{
	While wb%a_index%.readyState != 4 || wb%a_index%.document.readyState != "complete" || wb%a_index%.busy ; wait for the page to load
	Sleep, 10
	
	}
}

*/

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
