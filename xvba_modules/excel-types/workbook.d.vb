'/*
'Represents a Microsoft Excel workbook.
'
'
'*/
Public Class Workbook()

'/*
'Returns a String that represents the complete path to the 
'workbook/file that this workbook object represents.
'
'*/
Public Property Path As String


Public Property Worksheets As Worksheet

'/*
'True if a backup file is created when this file is saved. Read-only Boolean.
'
'@type {Boolean}
'*/
Public Property CreateBackup As Boolean

'/*
'Returns a String value that represents the name of the object.
'
'@type {Boolean}
'*/
Public Property Name As String

'/*
'True if no changes have been made to the specified workbook since it was last saved. Read/write Boolean.
'
'Example:
'If Not ActiveWorkbook.Saved Then 
' MsgBox "This workbook contains unsaved changes." 
'End If
'
'@type {Boolean}
'*/
Public Property Saved As String

'/*
'Returns a Sheets collection that represents all the sheets in the specified workbook. Read-only Sheets object.
'
'Example:
'Set newSheet = Sheets.Add(Type:=xlWorksheet) 
'For i = 1 To Sheets.Count 
'    newSheet.Cells(i, 1).Value = Sheets(i).Name 
'   Next i
'
'@type {Sheets}
'*/
Public Property Sheets As Worksheets


'/*
'
'Activates the first window associated with the workbook.
'
'*/
Public Function Activate() 

End Function

'/*
'
'Closes the object.
'
'Example
'
'Workbooks("BOOK1.XLS").Close SaveChanges:=False
'
'@param {Variant}  SaveChanges:[Optional] True or false
'@param {Variant}  FileName:[Optional] Saves changes under this file name.
'@param {Variant}  RouteWorkbook:[Optional] True or False
'*/
Public Function Close(SaveChanges,FileName,RouteWorkbook) 

End Function

'/*
'Saves changes to the specified workbook.
'*/
Public Function Save() 

End Function

'/*
'Saves changes to the workbook in a different file.
'
'@param {Variant}  FileName:[Optional]
'@param {Variant}  FileFormat:[Optional]
'@param {Variant}  Password:[Optional]
'@param {Variant}  WriteResPassword:[Optional]
'@param {Variant}  ReadOnlyRecommended:[Optional]
'@param {Variant}  CreateBackup:[Optional]
'@param {Variant}  AccessMode:[Optional]
'@param {Variant}  ConflictResolution:[Optional]
'@param {Variant}  AddToMru:[Optional]
'@param {Variant}  TextCodepage:[Optional]
'@param {Variant}  TextVisualLayout:[Optional]
'@param {Variant}  Local:[Optional]
'*/
Public Function SaveAs(FileName, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AccessMode, ConflictResolution, AddToMru, TextCodepage, TextVisualLayout, Local)

End Function


'/*
'Exports the data that has been mapped to the specified XML schema map to an XML data file.
'
'@param {String}  FileName:[Required]
'@param {XmlMap}  Map:[Required]
'*/
Public Function SaveAsXMLData(FileName, Map) 

End Function


'/*
'Saves a copy of the workbook to a file but doesn't modify the open workbook in memory.
'
'Example
'ActiveWorkbook.SaveCopyAs "C:\TEMP\XXXX.XLS"
'
'@param {Variant}  FileName:[Required]
'*/
Public Function SaveCopyAs(FileName, Map) 

End Function

'/*
'Sends the workbook by using the installed mail system.
'
'Example
'ActiveWorkbook.SendMail recipients:="Jean Selva"
'
'@param {Variant}  Recipients:[Required]
'@param {Variant}  Subject:[Optional]
'@param {Variant}  ReturnReceipt:[Optional]
'*/
Public Function SendMail(Recipients, Subject, ReturnReceipt) 

End Function

'/*
'
'The ExportAsFixedFormat method is used to publish 
'a workbook to either the PDF or XPS format.
'
'*/
Public Function ExportAsFixedFormat(Type, FileName, Quality, IncludeDocProperties, IgnorePrintAreas, From, To, OpenAfterPublish, FixedFormatExtClassPtr)
End Function

End Class