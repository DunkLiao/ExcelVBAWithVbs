Attribute VB_Name = "MergeExcelWithDataBar"
Option Explicit
'*************************************************************************************
'模組名稱: MergeExcelWithDataBar
'功能說明: 合併同一資料夾下的多個 Excel 檔案，並在合併結果的數值欄套用資料條格式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/28
'
'*************************************************************************************

Sub TestMergeWithDataBar()
    Dim folderPath As String
    folderPath = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xlsx;*.xls),*.xlsx;*.xls", _
        Title:="請選擇任一要合併的 Excel 檔案")
    If folderPath = "False" Then
        MsgBox "已取消操作。", vbExclamation, "取消"
        Exit Sub
    End If
    folderPath = Left(folderPath, InStrRev(folderPath, "\"))
    Call MergeFilesAndApplyDataBar(folderPath)
End Sub

Sub MergeFilesAndApplyDataBar(ByVal folderPath As String)
    Dim masterWb      As Workbook
    Dim masterWs      As Worksheet
    Dim srcWb         As Workbook
    Dim srcWs         As Worksheet
    Dim fileName      As String
    Dim lastRow       As Long
    Dim masterLastRow As Long
    Dim dataColIndex  As Integer
    Dim dataColLetter As String

    Set masterWb = Workbooks.Add
    Set masterWs = masterWb.Worksheets(1)
    masterWs.Name = "合併結果"
    masterLastRow = 1

    Application.ScreenUpdating = False

    fileName = Dir(folderPath & "*.xlsx")
    If fileName = "" Then fileName = Dir(folderPath & "*.xls")

    If fileName = "" Then
        MsgBox "資料夾內找不到 Excel 檔案，請確認路徑。", vbExclamation, "錯誤"
        masterWb.Close SaveChanges:=False
        Application.ScreenUpdating = True
        Exit Sub
    End If

    Do While fileName <> ""
        Set srcWb = Workbooks.Open(folderPath & fileName, ReadOnly:=True)
        Set srcWs = srcWb.Worksheets(1)
        lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row

        If masterLastRow = 1 Then
            srcWs.Range("A1").CurrentRegion.Copy
            masterWs.Range("A1").PasteSpecial Paste:=xlPasteValues
            masterLastRow = lastRow
        Else
            If lastRow > 1 Then
                srcWs.Range("A2:" & srcWs.Cells(lastRow, srcWs.UsedRange.Columns.Count).Address).Copy
                masterWs.Cells(masterLastRow + 1, 1).PasteSpecial Paste:=xlPasteValues
                masterLastRow = masterLastRow + lastRow - 1
            End If
        End If

        Application.CutCopyMode = False
        srcWb.Close SaveChanges:=False
        fileName = Dir()
    Loop

    If masterLastRow > 1 Then
        dataColIndex = 2
        dataColLetter = Chr(64 + dataColIndex)
        With masterWs.Range(dataColLetter & "2:" & dataColLetter & masterLastRow) _
                .FormatConditions.AddDatabar
            .MinPoint.Modify newtype:=xlConditionValueAutomaticMin
            .MaxPoint.Modify newtype:=xlConditionValueAutomaticMax
            .BarColor.Color = RGB(99, 180, 99)
            .BarFillType = xlDataBarFillGradient
            .ShowValue = True
        End With
    End If

    masterWs.Columns.AutoFit
    Application.ScreenUpdating = True
    MsgBox "合併完成！共 " & (masterLastRow - 1) & " 筆資料，已套用資料條格式。", _
           vbInformation, "完成"
End Sub
