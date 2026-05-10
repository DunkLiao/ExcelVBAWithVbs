Attribute VB_Name = "ConditionalFormulaExample"
Option Explicit

' ============================================================
' 模組名稱：ConditionalFormulaExample
' 功能說明：示範複合條件公式應用
'           包含 IF/AND/OR/IFERROR/IFS/SWITCH 組合範例
' ============================================================

Sub CreateConditionalFormulaExample()
    Dim ws      As Worksheet
    Dim wsName  As String
    Dim lastRow As Long
    Dim i       As Long
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    wsName = "複合條件公式範例"
    
    ' 若工作表已存在則刪除
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(wsName).Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrHandler
    
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = wsName
    
    ' --- 原始資料區 (A:C) ---
    ws.Range("A1").Value = "姓名"
    ws.Range("B1").Value = "分數"
    ws.Range("C1").Value = "出勤天數"
    
    ' 填入範例資料
    Dim arrNames(9)   As String
    Dim arrScores(9)  As Integer
    Dim arrDays(9)    As Integer
    arrNames(0) = "王小明"  : arrScores(0) = 92  : arrDays(0) = 22
    arrNames(1) = "李小華"  : arrScores(1) = 45  : arrDays(1) = 18
    arrNames(2) = "張大同"  : arrScores(2) = 78  : arrDays(2) = 25
    arrNames(3) = "陳美玲"  : arrScores(3) = 60  : arrDays(3) = 15
    arrNames(4) = "林志遠"  : arrScores(4) = 85  : arrDays(4) = 24
    arrNames(5) = "黃建國"  : arrScores(5) = 33  : arrDays(5) = 10
    arrNames(6) = "吳淑芬"  : arrScores(6) = 70  : arrDays(6) = 20
    arrNames(7) = "蔡明哲"  : arrScores(7) = 55  : arrDays(7) = 17
    arrNames(8) = "許文雄"  : arrScores(8) = 88  : arrDays(8) = 23
    arrNames(9) = "葉麗珍"  : arrScores(9) = 40  : arrDays(9) = 12
    
    For i = 0 To 9
        ws.Cells(i + 2, 1).Value = arrNames(i)
        ws.Cells(i + 2, 2).Value = arrScores(i)
        ws.Cells(i + 2, 3).Value = arrDays(i)
    Next i
    
    lastRow = 11
    
    ' --- 公式欄位標題 (D:H) ---
    ws.Range("D1").Value = "IF成績等第"
    ws.Range("E1").Value = "AND雙條件合格"
    ws.Range("F1").Value = "OR任一合格"
    ws.Range("G1").Value = "IFERROR保護"
    ws.Range("H1").Value = "IFS多層判斷"
    
    ' D欄：巢狀 IF 成績等第
    ws.Range("D2:D" & lastRow).Formula = _
        "=IF(B2>=90,""優"",IF(B2>=80,""良"",IF(B2>=70,""中"",IF(B2>=60,""可"",""差""))))"
    
    ' E欄：AND - 分數>=60 且 出勤>=20 才合格
    ws.Range("E2:E" & lastRow).Formula = _
        "=IF(AND(B2>=60,C2>=20),""合格"",""不合格"")"
    
    ' F欄：OR - 分數>=80 或 出勤>=24 任一達標
    ws.Range("F2:F" & lastRow).Formula = _
        "=IF(OR(B2>=80,C2>=24),""達標"",""未達標"")"
    
    ' G欄：IFERROR - 防止除零錯誤，計算每出勤日得分
    ws.Range("G2:G" & lastRow).Formula = _
        "=IFERROR(ROUND(B2/C2,2),""N/A"")"
    
    ' H欄：IFS 多層判斷（Excel 2019/365）
    ws.Range("H2:H" & lastRow).Formula = _
        "=IFERROR(IFS(B2>=90,""A"",B2>=80,""B"",B2>=70,""C"",B2>=60,""D"",B2<60,""F""),""N/A"")"
    
    ' --- 樣式設定 ---
    With ws.Range("A1:H1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.Columns("A:H").AutoFit
    
    Application.ScreenUpdating = True
    MsgBox "複合條件公式範例建立完成！" & vbCrLf & _
           "請查看「" & wsName & "」工作表。", vbInformation, "完成"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub