Attribute VB_Name = "MergeExcelByColumnName"
Option Explicit
'*************************************************************************************
'模組名稱: MergeExcelByColumnName
'功能說明: 依指定欄位名稱，合併同一資料夾下所有 Excel 檔案中符合的欄位資料
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************
' 範例使用入口
Sub TestMergeByColumnName()
    Call MergeExcelByColumnName
End Sub

' 依欄位名稱合併多個 Excel 檔案的資料
Sub MergeExcelByColumnName()
    On Error GoTo ErrorHandler

    Dim folderPath As String
    Dim columnNames As String
    Dim colArr() As String
    Dim fileName As String
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim targetRow As Long
    Dim lastRow As Long
    Dim i As Integer
    Dim j As Integer
    Dim isFirstFile As Boolean

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇包含 Excel 檔案的資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作", vbInformation, "取消"
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
    End With

    columnNames = InputBox("請輸入要合併的欄位名稱（以逗號分隔）：" & vbCrLf & _
                           "例如：姓名,部門,金額", "設定合併欄位", "")
    If columnNames = "" Then
        MsgBox "未輸入欄位名稱，已取消", vbInformation, "取消"
        Exit Sub
    End If

    colArr = Split(columnNames, ",")
    For i = 0 To UBound(colArr)
        colArr(i) = Trim(colArr(i))
    Next i

    On Error Resume Next
    Set wsTarget = ThisWorkbook.Worksheets("依欄位合併")
    On Error GoTo 0
    If wsTarget Is Nothing Then
        Set wsTarget = ThisWorkbook.Worksheets.Add
        wsTarget.Name = "依欄位合併"
    Else
        wsTarget.Cells.Clear
    End If

    For i = 0 To UBound(colArr)
        wsTarget.Cells(1, i + 1).Value = colArr(i)
    Next i
    wsTarget.Rows(1).Font.Bold = True

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    targetRow = 2
    isFirstFile = True

    fileName = Dir(folderPath & "\*.xlsx")
    Do While fileName <> ""
        If fileName <> ThisWorkbook.Name Then
            Set wbSource = Workbooks.Open(folderPath & "\" & fileName)
            Set wsSource = wbSource.Worksheets(1)
            lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row

            If lastRow >= 2 Then
                Dim srcColIndexes() As Integer
                ReDim srcColIndexes(UBound(colArr))

                For i = 0 To UBound(colArr)
                    srcColIndexes(i) = 0
                    For j = 1 To wsSource.UsedRange.Columns.Count
                        If Trim(CStr(wsSource.Cells(1, j).Value)) = colArr(i) Then
                            srcColIndexes(i) = j
                            Exit For
                        End If
                    Next j
                Next i

                Dim r As Long
                For r = 2 To lastRow
                    For i = 0 To UBound(colArr)
                        If srcColIndexes(i) > 0 Then
                            wsTarget.Cells(targetRow, i + 1).Value = _
                                wsSource.Cells(r, srcColIndexes(i)).Value
                        End If
                    Next i
                    targetRow = targetRow + 1
                Next r
            End If

            wbSource.Close SaveChanges:=False
        End If
        fileName = Dir()
    Loop

    wsTarget.UsedRange.Columns.AutoFit
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    wsTarget.Activate

    MsgBox "依欄位名稱合併完成！共 " & targetRow - 2 & " 筆資料。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "合併時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub
