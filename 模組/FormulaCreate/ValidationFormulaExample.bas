Attribute VB_Name = "ValidationFormulaExample"
Option Explicit
'*************************************************************************************
'模組名稱: ValidationFormulaExample
'功能說明: 以 VBA 建立資料驗證規則，示範清單驗證、整數範圍驗證
'          與自訂公式驗證三種常見場景
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/16
'
'*************************************************************************************

Sub CreateValidationExamples()
    Dim ws As Worksheet
    Set ws = GetOrCreateValidSheet("資料驗證範例")
    ws.Cells.Clear

    ' 標題
    ws.Range("A1").Value = "驗證類型"
    ws.Range("B1").Value = "輸入儲存格"
    ws.Range("C1").Value = "說明"

    ws.Range("A2").Value = "清單驗證"
    ws.Range("A3").Value = "整數範圍驗證"
    ws.Range("A4").Value = "自訂公式驗證（非空白）"

    ws.Range("C2").Value = "只能選擇：北區 / 中區 / 南區"
    ws.Range("C3").Value = "只能輸入 1 ~ 100 的整數"
    ws.Range("C4").Value = "儲存格不得為空白"

    ' ── 清單驗證（B2）
    Call AddListValidation(ws.Range("B2"), "北區,中區,南區")

    ' ── 整數範圍驗證（B3）
    Call AddIntegerValidation(ws.Range("B3"), 1, 100)

    ' ── 自訂公式驗證（B4）
    Call AddFormulaValidation(ws.Range("B4"), "LEN(B4)>0", "輸入不得為空白！")

    ws.Columns("A:C").AutoFit
    MsgBox "三種資料驗證規則已建立完成！請點選 B2、B3、B4 測試。", vbInformation, "完成"
End Sub

' 建立下拉清單驗證
Private Sub AddListValidation(ByVal rng As Range, ByVal listItems As String)
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:=listItems
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .InputTitle = "請選擇區域"
        .InputMessage = "從清單選取一個區域"
        .ShowError = True
        .ErrorTitle = "輸入錯誤"
        .ErrorMessage = "請從清單中選擇有效選項！"
    End With
End Sub

' 建立整數範圍驗證
Private Sub AddIntegerValidation(ByVal rng As Range, ByVal minVal As Long, ByVal maxVal As Long)
    With rng.Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:=CStr(minVal), _
             Formula2:=CStr(maxVal)
        .ShowInput = True
        .InputTitle = "請輸入整數"
        .InputMessage = "允許範圍：" & minVal & " 到 " & maxVal
        .ShowError = True
        .ErrorTitle = "超出範圍"
        .ErrorMessage = "請輸入 " & minVal & " 到 " & maxVal & " 的整數！"
    End With
End Sub

' 建立自訂公式驗證
Private Sub AddFormulaValidation(ByVal rng As Range, ByVal formula As String, ByVal errMsg As String)
    With rng.Validation
        .Delete
        .Add Type:=xlValidateCustom, _
             AlertStyle:=xlValidAlertWarning, _
             Formula1:="=" & formula
        .ShowError = True
        .ErrorTitle = "驗證警告"
        .ErrorMessage = errMsg
    End With
End Sub

Private Function GetOrCreateValidSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateValidSheet = ws
End Function
