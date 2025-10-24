'/*
'Represents a cell, a row, a column, a selection of cells containing one or more contiguous blocks of cells, or a 3D range.
'
'
'*/
Public Class Range()

'/*
'(Range)
'
'Activates a single cell, which must be inside the current selection. 
'To select a range of cells, use the Select method., 
'
'Example
'
'This example selects cells A1:C3 on Sheet1 and then makes cell B2 the active cell.
'
' Worksheets("Sheet1").Activate 
' Range("A1:C3").Select 
' Range("B2").Activate
'
'*/    
Public Sub Activate() 

End Sub


'/*
'Adds a comment to the range.
'
'Example:
'
'Worksheets(1).Range("E5").AddComment "Current Sales"
'
'@param {String} text
'*/
Public  Sub AddComment(text As String)

End Sub

Public  Sub AddCommentThreaded()

End Sub

Public  Sub AdvancedFilter()

End Sub

Public  Sub AllocateChanges()

End Sub

Public  Sub ApplyName()

End Sub

Public  Sub ApplyOutLineStyles()

End Sub

Public  Sub AutoComplete()

End Sub

Public  Sub AutoFill()

End Sub

Public  Sub AutoFilter()

End Sub

Public  Sub AutioFit()

End Sub

Public  Sub AutoOutline()

End Sub

Public  Sub BorderAround()

End Sub

Public  Sub Calculate()

End Sub

Public  Sub CalculateRowMajorOrder()

End Sub

Public  Sub CheckSpelling()

End Sub

Public  Sub Clear()

End Sub

Public  Sub ClearComments()

End Sub

Public  Sub ClearContents()

End Sub

Public  Sub ClearFormats()

End Sub

Public  Sub ClearHyperlinks()

End Sub

Public  Sub ClearNotes()

End Sub

Public  Sub ClearOutline()

End Sub

Public  Sub ColumnDifferences()

End Sub

Public  Sub Consolidate()

End Sub

Public  Sub ConvertToLinkedDataType()

End Sub

Public  Sub Copy()

End Sub

Public  Sub CopyFromRecordset()

End Sub

Public  Sub CopyPicture()

End Sub

Public  Sub CreateNames()

End Sub

Public  Sub Cut()

End Sub

Public  Function DataSeries (Rowcol, Type, Date, Step, Stop, Trend)

End Function

Public  Function DataTypeToText()

End Function

Public  Function Delete (Shift)

End Function

Public  Function DialogBox()

End Function

Public  Function Dirty()

End Function
Public  Function DiscardChanges()

End Function
Public  Function EditionOptions (Type, Option, Name, Reference, Appearance, ChartSize, Format)

End Function
Public  Function ExportAsFixedFormat (Type, FileName, Quality, IncludeDocProperties, IgnorePrintAreas, From, To, OpenAfterPublish, FixedFormatExtClassPtr)

End Function
Public  Function FillDown()

End Function
Public  Function FillLeft()

End Function

Public  Function FillRight()

End Function

Public  Function FillUp()

End Function

Public  Function Find (What, After, LookIn, LookAt, SearchOrder, SearchDirection, MatchCase, MatchByte, SearchFormat)

End Function

Public  Function FindNext (After)

End Function

Public  Function FindPrevious (Before)

End Function

Public  Function FlashFill ()

End Function


Public  Function FunctionWizard ()

End Function


Public  Function Group (Start, End, By, Periods)

End Function

Public  Function Insert (Shift, CopyOrigin)

End Function


Public  Function InsertIndent (InsertAmount)

End Function


Public  Function Justify()

End Function


Public  Function ListNames()

End Function

Public  Function Merge(Across)

End Function

Public  Function NavigateArrow(TowardPrecedent, ArrowNumber, LinkNumber)

End Function

Public  Function NoteText(Text, Start, Length)

End Function

Public  Function Parse(ParseLine, Destination)

End Function

Public  Function PasteSpecial(Paste, Operation, SkipBlanks, Transpose)

End Function


Public  Function PrintOut(From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate, PrToFileName)

End Function

Public  Function PrintPreview(EnableChanges)

End Function

Public  Function RemoveDuplicates(Columns , Header)

End Function

Public  Function RemoveSubtotal()

End Function


Public  Function Replace(What, Replacement, LookAt, SearchOrder, MatchCase, MatchByte, SearchFormat, ReplaceFormat)

End Function


Public  Function RowDifferences(Comparison)

End Function

Public  Function Run(Arg1, Arg2, Arg3, Arg4, Arg5...)

End Function

Public  Function Select()

End Function

Public  Function SetCellDataTypeFromCell(Range, LanguageCulture)

End Function


Public  Function SetPhonetic()

End Function


Public  Function Show()

End Function


Public  Function ShowCard()

End Function


Public  Function ShowDependents(Remove)

End Function


Public  Function ShowErrors()

End Function


Public  Function ShowPrecedents(Remove)

End Function


Public  Function Sort(Key1, Order1, Key2, Type, Order2, Key3, Order3, Header, OrderCustom, MatchCase, Orientation, SortMethod, DataOption1, DataOption2, DataOption3)

End Function


Public  Function SortSpecial(SortMethod, Key1, Order1, Type, Key2, Order2, Key3, Order3, Header, OrderCustom, MatchCase, Orientation, DataOption1, DataOption2, DataOption3)

End Function

Public  Function Speak(SpeakDirection, SpeakFormulas)

End Function


Public  Function SpecialCells(Type, Value)

End Function

Public  Function SubscribeTo(Edition, Format)

End Function

Public  Function Subtotal(GroupBy, Function, TotalList, Replace, PageBreaks, SummaryBelowData)

End Function


Public  Function Table(RowInput, ColumnInput)

End Function


Public  Function TextToColumns(Destination, DataType, TextQualifier, ConsecutiveDelimiter, Tab, Semicolon, Comma, Space, Other, OtherChar, FieldInfo, DecimalSeparator, ThousandsSeparator, TrailingMinusNumbers

End Function


Public  Function Ungroup()

End Function


Public  Function UnMerge()

End Function

Public Property AddIndent As Variant

Public Function Address(RowAbsolute, ColumnAbsolute, ReferenceStyle, External, RelativeTo) As Variant
End Function

Public Function AddressLocal(RowAbsolute, ColumnAbsolute, ReferenceStyle, External, RelativeTo) As Range
End Function

Public Property Application As Application 

Public Property Areas As Areas

Public Property Borders As Variant

Public Property Cells As Range 

Public Function Characters(Start, Length) As Variant
End Function

Public Property Column As Long

Public Property Columns As Long
        
Public Property ColumnWidth As Variant

Public Property Comment As Variant

Public Property CommentThreaded As Variant


Public Property Count As Long


Public Property CountLarge As Variant


Public Property Creator As Integer


Public Property CurrentArray As Range

'/*
'Returns a Range object that represents the current region. 
'The current region is a range bounded by any combination of blank rows and blank columns. Read-only.
'
'Example
'Worksheets("Sheet1").Activate 
'ActiveCell.CurrentRegion.Select
'*/
Public Property CurrentRegion As Range

Public Property Dependents As Range

Public Property DirectDependents As Range

Public Property DirectPrecedents As Range

Public Property DisplayFormat As Object


Public Function End(Direction) As Object
End Function

Public Property EntireColumn As Range


Public Property EntireRow As Range


Public Property Errors As Errors


Public Property Font As Font


Public Property FormatConditions As Range

Public Property Formula As Variant

Public Property FormulaArray As Variant

Public Property FormulaHidden As Variant


Public Property FormulaLocal As Object

Public Property FormulaR1C1 As Varaint

Public Property FormulaR1C1Local As Varaint

Public Property HasArray As Varaint

Public Property HasFormula As Varaint

Public Property HasRichDataType As Varaint

Public Property Height As Double

Public Property Hidden As Variant 

Public Property HorizontalAlignment As Variant

Public Property Hyperlinks As Object

Public Property ID As String

Public Property IndentLevel As Variant

Public Property Interior As Double

Public Function Item (RowIndex, ColumnIndex) As Range
End Function


Public Property Left As Variant


Public Property LinkedDataTypeState As Variant

Public Property ListHeaderRows As Variant


Public Property ListObject As Object

Public Property LocationInTable As Variant

Public Property Locked As Variant

Public Property MDX As String

Public Property MergeArea As Range

Public Property MergeCells As Boolean

Public Property Name As Variant

Public Property Next As Range

Public Property NumberFormat As Variant

Public Property NumberFormatLocal As Variant

Public Function Offset (RowOffset, ColumnOffset) As Range

End Function


Public Property Orientation As Variant

Public Property OutlineLevel As Variant

Public Property PageBreak As Variant

Public Property Parent As Variant

Public Property Phonetic As Variant
                    
Public Property Phonetics As Variant

Public Property PivotCell As Variant


Public Property PivotField As Variant


Public Property PivotItem As Variant

Public Property PivotTable As Variant

Public Property Precedents As Range

Public Property PrefixCharacter As Variant


Public Property Previous As Range

Public Property QueryTable As Variant

Public Function Range (Cell1, Cell2) As Range
End Function

Public Property ReadingOrder As Variant

Public Function Resize (RowSize, ColumnSize) As Range
End Function
        

Public Property Row As Long

Public Property RowHeight As Double

Public Property Rows As Range

Public Property ServerActions As Variant


Public Property ShowDetail As Variant


Public Property ShrinkToFit As Variant


Public Property SoundNote As Variant


Public Property SparklineGroups As Variant


Public Property Style As Variant

Public Property Summary As Variant


Public Property Text As String

Public Property Top As Variant

Public Property UseStandardHeight As Variant

Public Property UseStandardWidth As Variant

Public Property Validation As Variant

Public Function Value (RangeValueDataType) As Variant
End Function

Public Property Value2 As Variant


Public Property VerticalAlignment As Variant

Public Property Width As Double

                   
                    
Public Property Worksheet As Worksheet


Public Property WrapText As Variant

Public Property XPath As XPath
                        


End Class
