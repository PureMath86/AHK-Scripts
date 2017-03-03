#SingleInstance,, Force

w:=1150,h:=800,m:=50
Gui,+hWndn
Gui,Show,w%w% h%h%
Loop,% (w,x:=-1,d:=DllCall("GetDC",ptr,n))
{
Loop,% (h,y:=-1,j:=++x*(3/w)-2.1)
{
k:=++y*(2.5/h)-1.25,a:=b:=i:=0
while,(a*a+b*b<4&&i<m)
t:=a*a-b*b+j,b:=2*a*b+k,a:=t,i++
DllCall("SetPixel",ptr,d,int,x,int,y,int,i=m?0:i/m*255)
}
}
MsgBox


 ; Exit if the GUI is closed 
 GuiClose: 
 GuiEscape: 
   ExitApp
