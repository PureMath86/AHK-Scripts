: 1
; Repurpose Those Function Keys

;Launch Notepad
F7::Run C:\Windows\Notepad.exe 
return

; 2
; Open Webpages Quickly

; Launch Google
^+g::Run "www.google.com" ; use ctrl+Shift+g
return

; 3
; Open Downloads Folder Quickly

; Open Downloads folder
^+d::Run "C:\Users\User1\Downloads" ; ctrl+shift+d
return

; 4
; Make Tilde Work for You (Explorer Example)

; Press ~ to move up a folder in Explorer
#IfWinActive, ahk_class CabinetWClass
`::Send !{Up} 
#IfWinActive
return

; 5
; Make the Numpad Work for You (Volume Control)

; Custom volume buttons
+NumpadAdd:: Send {Volume_Up} ;shift + numpad plus
+NumpadSub:: Send {Volume_Down} ;shift + numpad minus
break::Send {Volume_Mute} ; Break key mutes
return

; 6
; Lock Keys

; Default state of lock keys
SetNumlockState, AlwaysOn
SetCapsLockState, AlwaysOff
SetScrollLockState, AlwaysOff
return




