Attribute VB_Name = "pageController"


'namespace=vba-files\Controllers


'/*
'
'This comment block is used by XVBA to
' show the sub info
'
'@return void
'*/
Public  Sub index()

  'XVBA auto-complete just work with namespace like "pageView"  
  Call pageView.publish
End Sub

'/*
'
'Test VBA Immediate Window Simulate 
'
'*/
Public Sub testXdebugPrint()

   Dim test(1)
  'Add an Object
  Set test(0) = Sheets(1)
  'Add a String
  test(1) = "Test Xdebug Output"

  Xdebug.printx test
  

End Sub

'/*
'
'Test VBA Immediate Window Simulate for Error
'
'*/
Public Sub testXdebugPrintError()

  
  On Error GoTo ErrorHandle:
    'throw an error
    d = 1/0
   'Your code here
  
  ErrorHandle:
    Xdebug.errorSource = "pageConsoller.index"
    Xdebug.printError
   

End Sub