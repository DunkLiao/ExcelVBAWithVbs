Option Explicit
Attribute VB_Name = "MergeExcelByFilePattern"
'*************************************************************************************
'模組名稱: 依檔名模式合併 Excel
'功能說明: 掃描指定資料夾，將符合檔名模式（萬用字元）的 Excel 檔案
'          第一張工作表的資料合併至新活頁簿
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/27
'
'*************************************************************************************
Sub MergeExcelByFilePattern()
    On Error GoTo ErrorHandler

    Dim folderPath As String
    Dim filePattern As String
    Dim fileName As String
    Dim wbSource As Workbook
    Dim wbDest As Workbook
    Dim wsDest As Worksheet
    Dim wsSource As Worksheet
    Dim lastRow As Long
    Dim destRow As Long
    Dim isFirstFile As Boolean

    folderPath = InputBox("請輸入要掃描的資料夾路徑（例如：C:\Reports\）", "資料夾路徑")
    If folderPath = "" Then Exit Sub

    If Right(folderPath, 1) <> "" Then folderPath = folderPath & ""

    filePattern = InputBox("請輸入檔名模式（例如：銷售_*.xlsx）", "檔名模式", "*.xlsx")
    If filePattern = "" Then Exit Sub

    fileName = Dir(folderPath & filePattern)
    If fileName = "" Then
        MsgBox "在指定資料夾中找不到符合模式的檔案：" & filePattern, vbExclamation, "找不到檔案"
        Exit Sub
    End If

    Set wbDest = Workbooks.Add
    Set wsDest = wbDest.Worksheets(1)
    wsDest.Name = "合併結果"
    destRow = 1
    isFirstFile = True

    Do While fileName <> ""
        Application.StatusBar = "正在處理：" & fileName
        Set wbSource = Workbooks.Open(folderPath & fileName, ReadOnly:=True)
        Set wsSource = wbSource.Worksheets(1)
        lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row

        If isFirstFile Then
            ' 第一個檔案含標題列
            wsSource.Range("A1:A" & lastRow).EntireRow.Copy _
                Destination:=wsDest.Rows(destRow)
            destRow = destRow + lastRow
            isFirstFile = False
        Else
            ' 後續檔案跳過標題列（第 2 列起）
            If lastRow >= 2 Then
                wsSource.Range("A2:A" & lastRow).EntireRow.Copy _
                    Destination:=wsDest.Rows(destRow)
                destRow = destRow + lastRow - 1
            End If
        End If

        wbSource.Close SaveChanges:=False
        fileName = Dir
    Loop

    wsDest.Columns.AutoFit
    Application.StatusBar = False

    MsgBox "合併完成！共 " & (destRow - 1) & " 列資料（含標題）。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    If Not wbSource Is Nothing Then wbSource.Close SaveChanges:=False
    MsgBox "合併時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub
