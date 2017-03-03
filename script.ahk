;
;
; 	Developer: Bryan Arnold 
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
SetTitleMatchMode RegEx ; 




^#!space::
{
 Pwb :=  ComObjCreate("InternetExplorer.Application")
 Pwb.Visible:=True
 Pwb.Navigate("http://www.personable.com/login.asp")
 Pwb.Navigate("https://www.securesheet-cloud.com/register/login.aspx", 2048)
 Pwb.Navigate("https://www.securesheet-cloud.com/register/login.aspx", 2048)
 Pwb.Navigate("https://www.securesheet-cloud.com/register/login.aspx", 2048)
 Pwb.Navigate("https://www.securesheet-cloud.com/register/login.aspx", 2048)
 Pwb.Navigate("https://docs.google.com/document/u/0/", 2048)
} 
Return




CapsLock::Run, calc

#CapsLock::Run, cmd

+CapsLock::Capslock






; Window Activation Hot-Keys

#c::WinActivate Pulover's Macro Creator v5.0.5
#h::WinActivate AutoHotkey Help
#i::WinActivate Internet Explorer
#n::WinActivate Notepad
#o::WinActivate Outlook
#s::WinActivate Skype
#x::WinActivate Excel

#b::WinActivate Microsoft Excel - Book1.xlsm

#w::WinActivate Winspector

; Launch Hot-Keys

^+a::Run C:\Program Files\AutoHotkey\AU3_Spy.exe
^+c::RUN C:\Users\Frey\Desktop\macros\PuloversMacroCreator-Portable\MacroCreatorPortable\x64\MacroCreator\MacroCreator.exe
^+h::Run C:\Program Files\AutoHotkey\AutoHotkey.chm
^+i::Run C:\Program Files\Internet Explorer\iexplore.exe
^+n::Run C:\Windows\system32\notepad.exe
^+o::Run C:\Program Files (x86)\Microsoft Office\Office14\OUTLOOK.EXE
^+s::Run C:\Program Files (x86)\Skype\Phone\Skype.exe
^+x::Run C:\Program Files (x86)\Microsoft Office\Office14\excel.exe

^+b::Run C:\Users\Frey\Desktop\Bryan\Book1.xlsx

; Hot-Strings
;
;   :*:something here::
;
;   ::something here::   <-- soft version (only auto-fills after an 'end' key, i.e. comma, space, etc.)
;

:*:@bbbb::bryanarnold@freyproduce.com

:*:@ssss::sborton@freyproduce.com

;------------------------------------------------------------;
;
;
;	Product Hot-Strings
;
;
;------------------------------------------------------------;

;
; SEEDLESS
;

:*:Sdls36::SEEDLESS:SDLS BINS 36
:*:Sdls60::SEEDLESS:SDLS BINS 60

;
; PUMPKIN MINI*
;

/*

PUMPKIN MINI*:MINI PUMPKIN-ACE
PUMPKIN MINI*:STRIPED MINI PUMPKIN-ACE
PUMPKIN MINI*:STRIPED MINI PUMPKIN 16/3ct
PUMPKIN MINI*:STRIPED MINI BULK
PUMPKIN MINI*:MINI PUMPKIN-LOWES
PUMPKIN MINI*:MINI PUMPKIN-Menards
PUMPKIN MINI*:STRIPED MINI PUMPKIN 14/3ct
PUMPKIN MINI*:MINI PUMPKIN 18#
PUMPKIN MINI*:MINI PUMPKIN BULK
PUMPKIN MINI*:MINI PUMPKINS 14/5ct
PUMPKIN MINI*:STRIPED MINI BULK 42ct
PUMPKIN MINI*:WHITE MINI  PUMPKIN BULK
PUMPKIN MINI*:WHITE MINI  PUMPKIN BULK:WHITE MINI PUMPKINS 5CT

*/

;
; PUMPKIN PIE*
;

:*:Pies12::PUMPKIN PIE*:PIE PUMPKIN 12ct

:*:Pies170::PUMPKIN PIE*:PIE PUMPKIN BIN 170ct

:*:Pies160::PUMPKIN PIE*:PIE PUMPKIN-Menards

:*:Pies300::PUMPKIN PIE*:PIE PUMPKIN 12ct

/*

PUMPKIN PIE*:PIE PUMPKIN 15ct
PUMPKIN PIE*:PIE PUMPKIN-ACE

PUMPKIN PIE*:PIE PUMPKIN 32#
PUMPKIN PIE*:PIE PUMPKIN BIN 110ct
PUMPKIN PIE*:PIE PUMPKIN BIN 140ct
PUMPKIN PIE*:PIE PUMPKIN BIN 200ct (60/40)
PUMPKIN PIE*:PIE PUMPKIN BIN 225ct
PUMPKIN PIE*:PIE PUMPKIN BIN 170ct
PUMPKIN PIE*:PIE PUMPKIN-LOWES
PUMPKIN PIE*:PIE PUMPKIN-Menards
PUMPKIN PIE*:PIE PUMPKIN BIN 200ct
PUMPKIN PIE*:PIE PUMPKIN 35#
PUMPKIN PIE*:PIE PUMPKIN 9ct
PUMPKIN PIE*:PIE PUMPKIN BIN 525#
PUMPKIN PIE*:PIE PUMPKIN BIN 600#

/*

;
; PUMPKINS*
;

/*

PUMPKINS*:WHT PUMPKIN-ACE
PUMPKINS*:PUMPKIN-ACE
PUMPKINS*:PUMPKIN BIN 21
PUMPKINS*:PUMPKIN BIN 18
PUMPKINS*:PUMPKIN-LOWES-Jumbo
PUMPKINS*:WHT PUMPKIN 6
PUMPKINS*:PUMPKIN-LOWES-Large
PUMPKINS*:PUMPKIN-LOWES-Med
PUMPKINS*:PUMPKIN-Menards 20
PUMPKINS*:PUMPKIN-Menards 50
PUMPKINS*:WHT PUMPKIN-Menards
PUMPKINS*:PUMPKIN - CAN

*/

:*:Pump25::PUMPKINS*:PUMPKIN BIN 25

:*:Pump20::PUMPKINS*:PUMPKIN BIN 20 ; Menard's

:*:Pump55::PUMPKINS*:PUMPKIN BIN 55

:*:Pump28::PUMPKINS*:PUMPKIN BIN 28 ; Lowe's

:*:Pump30::PUMPKINS*:PUMPKIN BIN 30
:*:Pump32::PUMPKINS*:PUMPKIN BIN 32
:*:Pump35::PUMPKINS*:PUMPKIN BIN 35
:*:Pump40::PUMPKINS*:PUMPKIN BIN 40
:*:Pump45::PUMPKINS*:PUMPKIN BIN 45

:*:Pump50::PUMPKINS*:PUMPKIN BIN 50 ; Lowe's + Menard's

:*:Pump60::PUMPKINS*:PUMPKIN BIN 60
:*:Pump70::PUMPKINS*:PUMPKIN BIN 70
:*:Pump16::PUMPKINS*:PUMPKIN SAM'S JUMBO

;
; AUTUMN COULEUR
;

/*

AUTUMN COULEUR:AUTUMN COULEUR-ACE
AUTUMN COULEUR:AUTUMN COULEUR-MENARDS:AUTUMN COULEUR DRC 12
AUTUMN COULEUR:AUTUMN COULEUR-MENARDS:AUTUMN COULEUR DRC 10
AUTUMN COULEUR:AUTUMN COULEUR-MENARDS:AUTUMN COULEUR DRC 8
AUTUMN COULEUR:AUTUMN COULEUR-MENARDS:AUTUMN COULEUR 40-MENARDS
AUTUMN COULEUR:AUTUMN COULEUR-LOWES:AUTUMN COULEUR 50-LOWES
AUTUMN COULEUR:AUTUMN COULEUR-LOWES:AUTUMN COULEUR 40-LOWES

*/

:*:AC100::AUTUMN COULEUR:AUTUMN COULEUR 100
:*:POTW100::AUTUMN COULEUR:AUTUMN COULEUR 100

:*:AC120::AUTUMN COULEUR:AUTUMN COULEUR 120
:*:POTW120::AUTUMN COULEUR:AUTUMN COULEUR 120

:*:AC140::AUTUMN COULEUR:AUTUMN COULEUR 140
:*:POTW140::AUTUMN COULEUR:AUTUMN COULEUR 140

:*:AC50::AUTUMN COULEUR:AUTUMN COULEUR 50
:*:AC40::AUTUMN COULEUR:AUTUMN COULEUR 40
:*:AC35::AUTUMN COULEUR:AUTUMN COULEUR 35
:*:AC30::AUTUMN COULEUR:AUTUMN COULEUR 30
:*:AC20::AUTUMN COULEUR:AUTUMN COULEUR 20
:*:AC45::AUTUMN COULEUR:AUTUMN COULEUR 45
:*:AC60::AUTUMN COULEUR:AUTUMN COULEUR 60

;------------------------------------------------------------;
;
;
;	Transportation Hot-Strings
;
;
;------------------------------------------------------------;


:*:All City Logistics::All City Logistics, LLP

:*:Armstrong Transport::Armstrong Transport Group, Inc.

:*:ATNT::ATNT Transport Inc.

:*:Carlyn Transport::Carlyn Transport, LLC

:*:Carroll Fulmer Logistics::Carroll Fulmer Logistics Corp.

:*:CH Robinson::C.H. Robinson Worldwide, Inc.

:*:D&L Transport LLC::D&L Transport, LLC

:*:Direct Freight::Direct Freight Solutions, LLC

:*:GS Enterprises::GS Enterprises, Inc.

:*:H&N Logistics::H&N Logistics,LLC

:*:J B Hunt::JB Hunt Transport, Inc.

:*:Kenda 4 Kenda::Kenda 4  Kenda

:*:McCord Trucking::McCord Trucking LLC

:*:PFL Logistics::PFL Logistics LLC

:*:QTI::QTI Service Corp.

:*:RJM Trucking::RJM Trucking

:*:Roark Trucking::Roark Trucking

:*:Ronnie L Hull Trucking LLC::Ronnie L. Hull Trucking LLC

:*:Turner Freight Transporters LLC::Turner Freight Transporters, LLC

;------------------------------------------------------------;
;
;
;	"Bill To" Hot-Strings
;
;
;------------------------------------------------------------;

/*



*/

;------------------------------------------------------------;
;
;
;	Supplier Hot-Strings
;
;
;------------------------------------------------------------;

/*



*/

;------------------------------------------------------------;
;
;
;	Loading Location Hot-Strings
;
;
;------------------------------------------------------------;

;
; Frey 
;

:*:PoseyvilleLL::
{
Send % "Pickup @ Poseyville, IN"
Sleep, 100
Send, {Enter}
Sleep, 100
Send % "812-874-3373"
}
Return

:*:Kennett, MO-Frey BrothersLL::
{
Send % "Pickup @ Kennett, MO"
Sleep, 100
Send, {Enter}
Sleep, 100
Send % "573-717-1252"
}
Return

:*:Wayne CityLL::
{
Send % "Pickup @ Wayne City, IL"
Sleep, 100
Send, {Enter}
Sleep, 100
Send % "618-835-2536"
}
Return

:*:KeenesLL::
{
Send % "Pickup @ Keenes, IL"
Sleep, 100
Send, {Enter}
Sleep, 100
Send % "618-835-2536"
}
Return

:*:NashvilleLL::
{
Send % "Pickup @ Nashville, IL"
Sleep, 100
Send, {Enter}
Sleep, 100
Send % "618-824-6515"
}
Return

;
; Other
;

:*:Gary SidesLL::
{
Send % "Pickup @ Sparta, NC"
Sleep, 100
Send, {Enter}
Sleep, 100
Send % "618-835-2536"
}
Return

:*:Wallendal Supply, Inc.LL::
{
Send % "Pickup @ Grand Marsh, WI"
Sleep, 100
Send, {Enter}
Sleep, 100
Send % "608-931-6779"
Sleep, 100
Send, {Enter}
Sleep, 100
Send % "608-215-6696"
}
Return

:*:Howell FarmsLL::
{
Send % "Pickup @ Middletown, IN"
Sleep, 100
Send, {Enter}
Sleep, 100
Send % "765-759-7432"
}
Return

:*:JB LaddLL::
{
Send % "Pickup @ Peru, IN"
Sleep, 100
Send, {Enter}
Sleep, 100
Send % "812-455-1565"
}
Return

:*:Henry DelagrangeLL::
{
Send % "Pickup @ Camden, MI"
Sleep, 100
Send, {Enter}
Sleep, 100
Send % "517-368-5892"
}
Return

:*:Sam SchwartzLL::
{
Send % "Pickup @ Reading, MI"
Sleep, 100
Send, {Enter}
Sleep, 100
Send % "517-296-4472"
}
Return


:*:Pine Hill FarmLL::
{
Send % "Pickup @ Hagerstown, IN"
Sleep, 100
Send, {Enter}
Sleep, 100
Send % "FAX: 765-334-4180"
}
Return

:*:Manley BrothersLL::
{
Send % "Pickup @ Bainbridge, GA"
Sleep, 100
Send, {Enter}
Sleep, 100
Send % "229-246-1226"
}
Return

/*

:*:LL::
{
Send % "Pickup @ Bainbridge, GA"
Sleep, 100
Send, {Enter}
Sleep, 100
Send % "229-246-1226"
}
Return

;
;                       ^
;                       |
;                       |
; Template ---/
;


*/

