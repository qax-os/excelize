'/*
'Represents the entire Microsoft Excel application.
'
'
'*/
Public Class Application()

'/*
'Returns an object that represents the active sheet (the sheet on top) 
'in the active workbook or in the specified window or workbook. 
'Returns Nothing if no sheet is active.
'
'*/
Public Property ActiveSheet As Worksheet

'/*
'Returns a Workbook object that represents the workbook in the 
'active window (the window on top). 
'Returns Nothing if there are no windows open or if either 
'the Info window or the Clipboard window is the active window. Read-only.
'
'
'*/
Public Property ActiveWorkbook As Workbook

'/*
'True if Microsoft Excel displays certain alerts and messages while 
'a macro is running. Read/write Boolean.
'
'Example
'Application.DisplayAlerts = False 
'Workbooks("BOOK1.XLS").Close 
'Application.DisplayAlerts = True
'
'*/
Public Property DisplayAlerts As Boolean


'/*
'Returns or sets an XlCalculation value that represents the calculation mode.
'
'Example
'Application.Calculation = xlCalculationManual 
'Application.Calculation = xlAutomatic
'
'@type {XlCalculation}
'*/
Public Property Calculation As XlCalculation

'/*
'True if events are enabled for the specified object. Read/write Boolean.
'
'@type {Boolean}
'*/
Public Property EnableEvents As Boolean

'/*
'True if screen updating is turned on. Read/write Boolean.
'@type {Boolean}
'*/
Public Property ScreenUpdating As Boolean


'/*
'Returns a Range object that represents the active cell in the active window 
'(the window on top) or in the specified window. If the window isn't displaying 
'a worksheet, this property fails. Read-only.
'
'@type {Object.<Range>}
'
'*/
Public Property ActiveCell As Range

'/*
'Returns a Chart object that represents the active chart (either an embedded chart or a chart sheet). 
'An embedded chart is considered active when it's either selected or activated. When no chart is active, 
'this property returns Nothing.
'
'Example: 
'ActiveChart.HasLegend = True
'
'@type {Object.<Chart>}
'*/
Public Property ActiveChart As Chart

'/*
'Returns a Workbooks collection that represents all the open workbooks. Read-only.
'
'@type {Object.<Collection>} Workbooks Collection
'*/
Public Property ThisWorkbook As Workbook

'/*
'Returns a Workbooks collection that represents all the open workbooks. Read-only.
'
'@type {Object.<Collection>} Workbooks Collection
'*/
Public Property Workbooks As Workbooks

'/*
'Activates a Microsoft application. If the application is already running, 
'this method activates the running application. 
'If the application isn't running, this method starts a new instance of the application.
'
'Example: (This example starts and activates Word.)
'
'Application.ActivateMicrosoftApp xlMicrosoftWord
'
'@param {XlMSApplication}  index 
'*/    
Public Sub ActivateMicrosoftApp( index As XlMSApplication) 

End Sub

'/*
'An event occurs when all pending refresh activity (both synchronous and asynchronous) 
'and all of the resultant calculation activities have been completed.
'
'*/
Public Event AfterCalculate()

'/*
'Occurs when a new workbook is created.
'
'Example:
'
'Private Sub App_NewWorkbook(ByVal Wb As Workbook) 
'Application.Windows.Arrange xlArrangeStyleTiled End Sub
'   
'@param {Workbook} Wb
'*/
Public Event NewWorkbook(ByVal Wb As Workbook) 

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

End Class