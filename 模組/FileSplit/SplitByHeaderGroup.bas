Attribute VB_Name = "SplitByHeaderGroup"
Option Explicit

' ============================================================
' 模組名稱：SplitByHeaderGroup
' 功能說明：依指定分組欄位將工作表資料分割至各子工作表
'           分組欄位由使用者輸入欄號（預設為第1欄）
' 使用方式：確認資料有標題列，執行後輸入分組欄號
' ============================================================

Sub SplitByHeaderGroup()
    Dim wsSrc       As Worksheet
    Dim wsNew       As Worksheet
    Dim lastRow     As Long
    Dim lastCol     As Long
    Dim groupCol    As Long
    Dim i           As Long
    Dim groupKey    As String
    Dim wsName      As String
    Dim headerRange As Range
    Dim colInput    As String
    
    ' 使用目前作用中工作表
    Set wsSrc = ActiveSheet
    
    On Error GoTo ErrHandler
    
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
    
    If lastRow < 2 Then
        MsgBox "工作表中沒有足夠的資料（至少需要標題列與一列資料）。", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 詢問分組欄號
    colInput = InputBox("請輸入分組欄號（數字，例如輸入 1 代表第 A 欄）：", _
                        "分組欄號", "1")
    If colInput = "" Then
        MsgBox "已取消操作。", vbInformation, "取消"
        Exit Sub
    End If
    
    If Not IsNumeric(colInput) Then
        MsgBox "請輸入有效的欄號數字。", vbExclamation, "輸入錯誤"
        Exit Sub
    End If
    
    groupCol = CLng(colInput)
    If groupCol < 1 Or groupCol > lastCol Then
        MsgBox "欄號超出資料範圍（1 到 " & lastCol & "）。", vbExclamation, "輸入錯誤"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' 使用字典儲存各分組工作表
    Dim dictSheets As Object
    Set dictSheets = CreateObject("Scripting.Dictionary")
    
    ' 儲存標題列
    Set headerRange = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(1, lastCol))
    
    ' 逐列讀取，依分組鍵建立工作表並寫入資料
    For i = 2 To lastRow
        groupKey = CStr(wsSrc.Cells(i, groupCol).Value)
        If groupKey = "" Then groupKey = "(空白)"
        
        ' 清除工作表名稱中的非法字元
        wsName = CleanSheetName(groupKey)
        
        ' 若工作表尚未建立則新增
        If Not dictSheets.Exists(wsName) Then
            ' 刪除同名舊工作表
            On Error Resume Next
            Application.DisplayAlerts = False
            ThisWorkbook.Sheets(wsName).Delete
            Application.DisplayAlerts = True
            On Error GoTo ErrHandler
            
            Set wsNew = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            wsNew.Name = wsName
            
            ' 複製標題列
            headerRange.Copy Destination:=wsNew.Range("A1")
            wsNew.Rows(1).Font.Bold = True
            wsNew.Rows(1).Interior.Color = RGB(68, 114, 196)
            wsNew.Rows(1).Font.Color = RGB(255, 255, 255)
            
            dictSheets.Add wsName, 2  ' 下一個可寫入列從第2列開始
        End If
        
        ' 寫入資料列
        Set wsNew = ThisWorkbook.Sheets(wsName)
        Dim writeRow As Long
        writeRow = dictSheets(wsName)
        
        wsSrc.Range(wsSrc.Cells(i, 1), wsSrc.Cells(i, lastCol)).Copy _
            Destination:=wsNew.Cells(writeRow, 1)
        
        dictSheets(wsName) = writeRow + 1
    Next i
    
    ' 自動調整各工作表欄寬
    Dim shKey As Variant
    For Each shKey In dictSheets.Keys
        ThisWorkbook.Sheets(CStr(shKey)).Columns.AutoFit
    Next shKey
    
    Application.ScreenUpdating = True
    MsgBox "分割完成！共建立 " & dictSheets.Count & " 個群組工作表。", _
           vbInformation, "完成"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub

' 清除工作表名稱中的非法字元
Private Function CleanSheetName(ByVal sName As String) As String
    Dim illegalChars As String
    Dim c As String
    Dim result As String
    Dim j As Integer
    
    illegalChars = "\/?*[]:"
    result = sName
    
    For j = 1 To Len(illegalChars)
        c = Mid(illegalChars, j, 1)
        result = Replace(result, c, "_")
    Next j
    
    ' 最大31字元
    If Len(result) > 31 Then result = Left(result, 31)
    
    CleanSheetName = result
End Function