'/*
'A collection of all the Workbook objects that are currently open in the Microsoft Excel application.
'
'*/
Public Class Workbooks() 

'/*
'Creates a new workbook. The new workbook becomes the active workbook.
'
'@param {Variant}  Template:[Optional] Determines how the new workbook is created. 
'If this argument is a string specifying the name of an existing Microsoft Excel file, 
'the new workbook is created with the specified file as a template.
'*/
Public Function Add(Template) 

End Function


'/*
'True if Microsoft Excel can check out a specified workbook from a server. 
'Read/write Boolean.
'
'@param {String}  FileName:[Required] 	The name of the file to check out.
'
'*/
Public Function CanCheckOut (FileName) 

End Function


Public Function CheckOut (FileName) 

End Function


Public Function Close () 

End Function

'/*
'
'Opens a workbook.
'
'Example
'Workbooks.Open "ANALYSIS.XLS" 
'ActiveWorkbook.RunAutoMacros xlAutoOpen
'*/
Public Function Open  (FileName, UpdateLinks, ReadOnly, Format, Password, WriteResPassword, IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, Notify, Converter, AddToMru, Local, CorruptLoad) As Workbook

End Function

'/*
'Returns a Workbook object representing a database.
'*/
Public Function OpenDatabase (FileName, CommandText, CommandType, BackgroundQuery, ImportDataAs)

End Function


'/*
'Loads and parses a text file as a new workbook with a single sheet that contains the parsed text-file data.
'*/
Public Function  (FileName, Origin, StartRow, DataType, TextQualifier, ConsecutiveDelimiter, Tab, Semicolon, Comma, Space, Other, OtherChar, FieldInfo, TextVisualLayout, DecimalSeparator, ThousandsSeparator, TrailingMinusNumbers, Local)

End Function

'/*
'Opens an XML data file. Returns a Workbook object.
'*/
Public Function OpenXML (FileName, Stylesheets, LoadOption)

End Function


'/*
'When used without an object qualifier, this property returns an 
'Application object that represents the Microsoft Excel application.
'
'*/
Public Property Application As Application


'/*
'Returns a Long value that represents the number of objects in the collection.
'
'*/
Public Property Count As Long

'/*
'Returns a 32-bit integer that indicates the application in which this object was created. Read-only Long.
'
'*/
Public Property Creator As Integer



'/*
'Returns a single object from a collection.
'
'*/
Public Property Item(Index) As Object

'/*
'Returns the parent object for the specified object. Read-only.
'
'*/
Public Property Parent As Object

End Class