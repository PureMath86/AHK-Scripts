
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

oWorkbook := ComObjGet(SelectedFile) ; 

;
;
;


Value := oWorkbook.Sheets(12).Range("C1").Value ; Get the value to look-up

Sleep, 100 

If (Value = "")
{
MsgBox The cell is Blank...
}
else
{
MsgBox, Shit...
}
