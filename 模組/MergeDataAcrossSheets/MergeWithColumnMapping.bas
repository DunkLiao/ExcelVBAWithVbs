Attribute VB_Name = "MergeWithColumnMapping"
Option Explicit

' ============================================================
' 模組名稱：MergeWithColumnMapping
' 功能說明：依欄位標題名稱對應，將多個工作表合併至主表
'           即使各工作表欄位順序不同，也能正確對應合併
' 使用方式：確認每個來源工作表第一列為標題列，執行此巨集
' ============================================================

Sub MergeWithColumnMapping()
    Dim wsMaster    As Worksheet
    Dim wsSrc       As Worksheet
    Dim masterName  As String
    Dim nextRow     As Long
    Dim i           As Long
    Dim j           As Long
    Dim srcLastRow  As Long
    Dim srcLastCol  As Long
    Dim masterCols  As Long
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    masterName = "欄位對應合併"
    
    ' 刪除舊的主表
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(masterName).Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True
    
    ' 新增主表並置於最後
    Set wsMaster = ThisWorkbook.Sheets.Add( _
        After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsMaster.Name = masterName
    
    masterCols = 0
    nextRow = 1
    
    Dim dictMasterHeader As Object
    Set dictMasterHeader = CreateObject("Scripting.Dictionary")
    
    ' 第一輪：收集所有工作表的欄位標題，建立主表標題列
    For Each wsSrc In ThisWorkbook.Sheets
        If wsSrc.Name <> masterName Then
            srcLastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
            For j = 1 To srcLastCol
                Dim hdr As String
                hdr = Trim(CStr(wsSrc.Cells(1, j).Value))
                If hdr <> "" And Not dictMasterHeader.Exists(hdr) Then
                    masterCols = masterCols + 1
                    dictMasterHeader.Add hdr, masterCols
                    wsMaster.Cells(1, masterCols).Value = hdr
                End If
            Next j
        End If
    Next wsSrc
    
    ' 加入「來源工作表」欄
    masterCols = masterCols + 1
    wsMaster.Cells(1, masterCols).Value = "來源工作表"
    
    ' 設定標題列樣式
    With wsMaster.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With
    nextRow = 2
    
    ' 第二輪：逐工作表讀取資料並依欄位名稱對應寫入
    For Each wsSrc In ThisWorkbook.Sheets
        If wsSrc.Name <> masterName Then
            srcLastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
            srcLastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
            
            If srcLastRow >= 2 Then
                ' 建立此來源工作表的欄位對應
                Dim dictSrcCol As Object
                Set dictSrcCol = CreateObject("Scripting.Dictionary")
                
                For j = 1 To srcLastCol
                    Dim srcHdr As String
                    srcHdr = Trim(CStr(wsSrc.Cells(1, j).Value))
                    If srcHdr <> "" Then
                        dictSrcCol(srcHdr) = j
                    End If
                Next j
                
                ' 逐列寫入主表
                For i = 2 To srcLastRow
                    Dim masterKey As Variant
                    For Each masterKey In dictMasterHeader.Keys
                        If dictSrcCol.Exists(masterKey) Then
                            wsMaster.Cells(nextRow, dictMasterHeader(masterKey)).Value = _
                                wsSrc.Cells(i, dictSrcCol(masterKey)).Value
                        End If
                    Next masterKey
                    ' 記錄來源工作表名稱
                    wsMaster.Cells(nextRow, masterCols).Value = wsSrc.Name
                    nextRow = nextRow + 1
                Next i
            End If
        End If
    Next wsSrc
    
    wsMaster.Columns.AutoFit
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "欄位對應合併完成！" & vbCrLf & _
           "共合併 " & (nextRow - 2) & " 筆資料。" & vbCrLf & _
           "結果已存至「" & masterName & "」工作表。", vbInformation, "完成"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub