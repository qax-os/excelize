Attribute VB_Name = "MyRibbon"

'namespace=vba-files/ribbons

'/*
'[Ribbon Menu Action]
'Ribbon Buttom 1 example call code
'
'*/
Public Sub btn1(ByRef control As Office.IRibbonControl)

 MsgBox "Click Ribbon Button 1"

End Sub

'/*
'[Ribbon Menu Action]
'Ribbon Buttom 2 example call code
'
'*/
Public Sub btn2(ByRef control As Office.IRibbonControl)

 MsgBox "Click Ribbon Button 2"
 
End Sub


'/*
'[Ribbon Menu Action]
'Ribbon Buttom 2 example call code
'
'*/
Public Sub btn3(ByRef control As Office.IRibbonControl)

 MsgBox "Click Ribbon Button 3 on Second Group"
 
End Sub