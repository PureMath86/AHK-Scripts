
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
 #Warn  ; Enable warnings to assist with detecting common errors.
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

;
;  Get the URL of current tab in IE
;

#w::
{
ClipSaved := ClipboardAll ; Save the Clipboard  

clipboard =  ; Start off empty to allow ClipWait to detect when the text has arrived.

Send, {f6} ; select the URL

Sleep, 100

Send, ^c
ClipWait ; Wait for the clipboard to contain text

MsgBox , , , %Clipboard%,

Clipboard := ClipSaved ; Restore the original clipboard. Note: NOT ClipboardAll.

ClipSaved :=   ; Free the memory 
}
