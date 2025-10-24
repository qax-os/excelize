Attribute VB_Name = "Alert"


'namespace=vba-files\Helpers


'/*
'
'This comment block is used by XVBA to
' show the sub info
'
'@return void
'*/
Public  Sub Show()
  Call Xlog.message(0,"Run Alert Show Sub")
  MsgBox "Alert",,"title" 
End Sub
