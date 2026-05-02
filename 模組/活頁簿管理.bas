Attribute VB_Name = "活頁簿管理"
Option Explicit

Sub 建立大量工作表並建立超連結()
  Application.ScreenUpdating = False
  Call 工作表大量命名
  Application.DisplayAlerts = False
  Sheets("總表").Delete
  Application.DisplayAlerts = True
  Call 建立所有工作表的索引
  Application.ScreenUpdating = True
End Sub

Sub 工作表大量命名()

'刪除現有工作表
Dim q  As Integer
Application.DisplayAlerts = False
    Sheets("工作表2").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("工作表3").Select
    ActiveWindow.SelectedSheets.Delete
Application.DisplayAlerts = True

'增加工作表
Dim my_name, pre_name As String
Dim i, j As Byte
Dim my_range As Range
Sheets(1).Name = "總表"
Sheets("總表").Activate
Range("A1").Activate
my_name = ""
pre_name = ""
my_name = ActiveCell.Text
Set my_range = ActiveCell

Do While my_name <> pre_name
    my_name = my_name
    Sheets("總表").Activate
    my_range.Offset(1, 0).Activate
    Set my_range = ActiveCell
    my_name = ActiveCell.Text
On Error Resume Next
j = i + 2
    If my_name <> "" Then
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(j).Name = my_name
    End If
i = i + 1
'MsgBox my_name & my_name
Loop
Sheets(1).Name = "總表"
Sheets("總表").Activate
Range("A1").Activate

End Sub

Sub 建立所有工作表的索引()
    Dim mySht As Worksheet
    Dim ShtName As String
    Dim myRng As Range
    Dim i, j As Integer
    Dim myWb As Workbook
    Set myWb = ActiveWorkbook
    Set mySht = myWb.Worksheets.Add(Count:=1, Before:=Sheets(1))
    
    On Error Resume Next
   
    mySht.Name = "工作表索引"
    Columns(1).Clear                            '儲存格的清除
    Set myRng = Range("A1")
    Range("A1").Value = "工作表索引"            '標題的設定
    
    For Each mySht In Worksheets
        ShtName = mySht.Name
        If mySht.Name = "工作表索引" Then
            Sheets("工作表索引").Activate
            Columns(1).Clear                            '儲存格的清除
            Set myRng = Range("A1")
            Range("A1").Value = "工作表索引"            '標題的設定

        Else
        
     '   Cells(Rows.Count, 2).End(xlUp).Offset(1).Value = mySht.Name     '參考 說明()
    '    MsgBox mySht.Name
        Set myRng = myRng.Offset(1)
        myRng.Activate
        ActiveCell.Hyperlinks.Add Anchor:=Selection, SubAddress:="'" & ShtName & "'" & "!A1", Address:="", TextToDisplay:=ShtName
    
        End If
    Next
    'MsgBox Rows.Count
    Set mySht = Nothing


End Sub

Sub 保護所有工作表()

     ' Loop through all sheets in the workbook.
     Dim i As Integer
     For i = 1 To Sheets.Count
        Sheets(i).Protect
     Next i
End Sub

Sub 取消保護所有工作表()
     ' Loop through all sheets in the workbook.
     Dim i As Integer
     For i = 1 To Sheets.Count
        Sheets(i).Unprotect
     Next i
End Sub
Sub 清除內容()
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.ClearContents
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.Select
End Sub

Sub 調整欄寬符合文字內容()
    ActiveSheet.Columns.AutoFit
End Sub
