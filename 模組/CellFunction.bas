Attribute VB_Name = "CellFunction"
Option Explicit
'匡
Sub SelectAllTable()
    ActiveCell.CurrentRegion.Select
End Sub
'眔程(row length)
Function NotSpaceRow(ByVal columnName As String)
    Dim columnindex As String
    '侣excel程65535
    columnindex = columnName & "65536"
    Dim myRange As Range
    Set myRange = ThisWorkbook.Sheets(1).Range(columnindex).End(xlUp)
    myRange.Select
    NotSpaceRow = myRange.Row
    Set myRange = Nothing
End Function
'眔程(row length)
Function NotSpaceRowBySheetName(ByVal columnName As String, ByVal sheetName As String)
    Dim columnindex As String
    '侣excel程65535
    columnindex = columnName & "65536"
    Dim myRange As Range
    Set myRange = ThisWorkbook.Sheets(sheetName).Range(columnindex).End(xlUp)
    myRange.Select
    NotSpaceRowBySheetName = myRange.Row
    Set myRange = Nothing
End Function
'眔程逆(column length)
Function NotSpaceColumns(ByVal rowIndex As Integer)
    Dim selectedRow As String
    selectedRow = "IV" & rowIndex
    Dim myRange As Range
    Set myRange = ThisWorkbook.Sheets(1).Range(selectedRow).End(xlToLeft)
    myRange.Select
    NotSpaceColumns = myRange.Column
    Set myRange = Nothing
End Function
Function NotSpaceColumnsBySheetNam(ByVal rowIndex As Integer, ByVal sheetName As String)
    Dim selectedRow As String
    selectedRow = "IV" & rowIndex
    Dim myRange As Range
    Set myRange = ThisWorkbook.Sheets(sheetName).Range(selectedRow).End(xlToLeft)
    myRange.Select
    NotSpaceColumnsBySheetNam = myRange.Column
    Set myRange = Nothing
End Function
