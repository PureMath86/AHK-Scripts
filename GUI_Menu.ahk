
; Beginning of script's auto-execute section.
; -~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~- ;
; Create the popup menu by adding some items to it.
Menu, MyMenu, Add, Item&1, MenuHandler
Menu, MyMenu, Add, Item&2, MenuHandler

; Add a separator line.
Menu, MyMenu, Add 

; Create another menu destined to become a submenu of the above menu.
Menu, Submenu1, Add, Item&1, MenuHandler
Menu, Submenu1, Add, Item&2, MenuHandler

; Create a submenu in the first menu (a right-arrow indicator). 
; When the user selects it, the second menu is displayed.
Menu, MyMenu, Add, My &Submenu, :Submenu1

; Add a separator line.
Menu, MyMenu, Add  

; Add another menu item beneath the submenu.
Menu, MyMenu, Add, Item&3, MenuHandler  
return  
; -~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~-~- ;
; End of script's auto-execute section.

; MenuHandler Example
MenuHandler:
MsgBox You selected %A_ThisMenuItem% from the menu %A_ThisMenu%.
return

; Hot-Key to show the Menu
#z::Menu, MyMenu, Show  ; i.e. press the Win-Z hotkey to show the menu.
