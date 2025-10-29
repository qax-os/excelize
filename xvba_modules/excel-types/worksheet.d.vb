'/*
'Represents a worksheet.
'
'Example
'
'Worksheets(1).Visible = False
'
'*/
Public Class Worksheet()


'/*
'Makes the current sheet the active sheet.
'
'Example
'ActiveWorkbook.SendMail recipients:="Jean Selva"
'*/
Public Sub Activate() 

End Sub


Public Sub Calculate() 

End Sub

Public Sub Copy() 

End Sub

Public Sub Delete() 
End Sub

Public Sub Move()

End Sub

Public Sub Past() 

End Sub

Public Sub PastSpecial() 

End Sub

Public Sub Select() 
End Sub

Public Sub SaveAs()

End Sub  

'/*
'
'Returns a Range object that represents a cell or a range of cells.
'
'Example: 
'
'Worksheets("Sheet1").Range("A1").Value = 3.14159
'
'*/
Public Property Range As Range


'/*
'
'Returns a Range object that represents all the rows on the specified worksheet.
'
'Example: 
'
'Worksheets("Sheet1").Rows(3).Delete
'
'*/
Public Property Row As Range


'/*
'
'Returns a Long value that represents the index number of the 
'object within the collection of similar objects.orksheets("Sheet1").Rows(3).Delete
'
'*/
Public Property Index As Long


End Class