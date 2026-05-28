Attribute VB_Name = "SplitSheetByCategory"
Option Explicit
'*************************************************************************************
'模組名稱: SplitSheetByCategory
'功能說明: 依據指定的分類欄位將工作表資料切割，每個分類獨立儲存為一個新的 Excel 檔案
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/28
'
'*************************************************************************************

Sub TestSplitByCategory()
    Dim ws          As Worksheet
    Set ws = ActiveSheet
    If ws.UsedRange.Rows.Count < 2 Then
        MsgBox "工作表內沒有足夠的資料，請先填入資料後再執行。", vbExclamation, "錯誤"
        Exit Sub
    End If
    Call SplitSheetByCategory(ws, 1)
End Sub

Sub SplitSheetByCategory(ByVal srcWs As Worksheet, ByVal catColIndex As Integer)
    Dim wb          As Workbook
    Dim newWb       As Workbook
    Dim newWs       As Worksheet
    Dim lastRow     As Long
    Dim i           As Long
    Dim catValue    As String
    Dim categories  As Object
    Dim savePath    As String
    Dim safeFile    As String

    Set wb = srcWs.Parent
    lastRow = srcWs.Cells(srcWs.Rows.Count, catColIndex).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "沒有資料可供切割。", vbExclamation, "錯誤"
        Exit Sub
    End If

    Set categories = CreateObject("Scripting.Dictionary")
    For i = 2 To lastRow
        catValue = CStr(srcWs.Cells(i, catColIndex).Value)
        If Len(catValue) > 0 And Not categories.Exists(catValue) Then
            categories.Add catValue, catValue
        End If
    Next i

    If categories.Count = 0 Then
        MsgBox "分類欄位為空，無法切割。", vbExclamation, "錯誤"
        Exit Sub
    End If

    savePath = Left(wb.FullName, InStrRev(wb.FullName, "\"))
    If savePath = "" Then savePath = Environ("TEMP") & "\"

    Application.ScreenUpdating = False

    Dim key As Variant
    For Each key In categories.Keys
        catValue = CStr(key)
        Set newWb = Workbooks.Add
        Set newWs = newWb.Worksheets(1)
        newWs.Name = Left(catValue, 31)

        srcWs.Rows(1).Copy
        newWs.Rows(1).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False

        Dim newRow As Long
        newRow = 2
        For i = 2 To lastRow
            If CStr(srcWs.Cells(i, catColIndex).Value) = catValue Then
                srcWs.Rows(i).Copy
                newWs.Rows(newRow).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                newRow = newRow + 1
            End If
        Next i

        newWs.Columns.AutoFit

        safeFile = catValue
        Dim invalidChars As Variant
        invalidChars = Array("\", "/", ":", "*", "?", Chr(34), "<", ">", "|")
        Dim c As Variant
        For Each c In invalidChars
            safeFile = Replace(safeFile, c, "_")
        Next c

        On Error Resume Next
        newWb.SaveAs Filename:=savePath & safeFile & ".xlsx", _
                     FileFormat:=xlOpenXMLWorkbook
        On Error GoTo 0
        newWb.Close SaveChanges:=False
    Next key

    Application.ScreenUpdating = True
    MsgBox "切割完成！共產生 " & categories.Count & " 個檔案。" & vbCrLf & _
           "儲存位置：" & savePath, vbInformation, "完成"
End Sub
