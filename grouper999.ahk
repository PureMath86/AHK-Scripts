;***********save clipboard to restore later******************* 
Store:=ClipboardAll  ; Store full version of Clipboard
  clipboard = ; Empty the clipboard
  SendInput, ^c ; changed from Send 
  ClipWait, 1
    If ErrorLevel ; Added errorLevel checking
      {
        MsgBox, No text was sent to clipboard
        Return
      }
;***********remove blank lines*******************       
Loop,parse,Clipboard,`n`r ; loops through lines
{
  New:=RegExReplace(A_LoopField,"^\s?$","") ; if line only has white space
    If (new!="") 
      NewText.=new "`r`n"  
 }
 Clipboard:=NewText
 NewText:=""
 
;***********clean up the text******************* 
Trimmed:=RegExReplace(clipboard, "m)^\s?(.*)\s?$","$1") ; trim spaces from both sides
;***********************wrap special characthers********************************.
StringReplace, Trimmed, Trimmed, &, '||'&'||',All ; wrap commas
StringReplace, Trimmed, Trimmed, ', '',All ; double up ' so it escapes it
 
;***********************Create groups of 999********************************.
Loop,
{
  NewStr := RegExReplace(Trimmed, "`r`n", "','",1,998,1)  ;998 is max
  AddBreak := RegExReplace(NewStr, "`r`n", "`n`r*****************" . (A_Index*1000)+1 . "*****************`r",1,1,1)  
   If (AddBreak=NewStr) ;check to see if done looping through all items
    break
   Trimmed:=AddBreak ; 
  }
Clipboard := RegExReplace(NewStr, "','$", "") ; Delete ending ','
 
;***********restore clipboard******************* 
SendEvent , ^v
Sleep, 50
Clipboard:=Store
 
;***********wipe out vars of <strong> SQL in list</strong>******************* 
New:="" , newstr:="" , newText:="" , Trimmed:="" , Addbreak:=""  ; Empty vars
return
