Attribute VB_Name = "MergeFixedWidthTextFiles"
Option Explicit

' ============================================================
' 範例：合併同一資料夾下所有固定寬度格式 (.txt) 文字檔
' 功能：使用者定義各欄位的起始位置與寬度，程式讀取每一列，
'       依定義拆解欄位後，寫入目標工作表的對應儲存格
' 說明：固定寬度格式例如：
'         姓名(1-10)  部門(11-20)  金額(21-28)
' ============================================================

' 固定寬度欄位定義型別
Private Type FieldDef
    FieldName   As String   ' 欄位名稱
    StartPos    As Long     ' 起始字元位置（1 起算）
    FieldWidth  As Long     ' 欄位寬度（字元數）
End Type

Sub MergeFixedWidthTextFiles()
    On Error GoTo ErrHandler

    Dim strFolder   As String
    Dim strFile     As String
    Dim intIn       As Integer
    Dim strLine     As String
    Dim wsDest      As Worksheet
    Dim lngDestRow  As Long
    Dim blnFirst    As Boolean
    Dim lngFilesCount As Long

    ' --- 定義固定寬度欄位（依實際格式修改）---
    Dim nFields     As Long
    nFields = 3
    Dim fields(1 To 3) As FieldDef
    fields(1).FieldName = "姓名"
    fields(1).StartPos = 1
    fields(1).FieldWidth = 10

    fields(2).FieldName = "部門"
    fields(2).StartPos = 11
    fields(2).FieldWidth = 10

    fields(3).FieldName = "金額"
    fields(3).StartPos = 21
    fields(3).FieldWidth = 8

    ' --- 選擇資料夾 ---
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇包含固定寬度 TXT 檔案的資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation, "提示"
            Exit Sub
        End If
        strFolder = .SelectedItems(1)
    End With

    If Right(strFolder, 1) <> "\" Then strFolder = strFolder & "\"

    ' --- 建立目標工作表 ---
    Set wsDest = ThisWorkbook.Worksheets.Add
    wsDest.Name = "FixedWidth_" & Format(Now(), "mmddHHmm")
    lngDestRow = 1
    blnFirst = True
    lngFilesCount = 0

    Application.ScreenUpdating = False

    strFile = Dir(strFolder & "*.txt")
    If strFile = "" Then
        MsgBox "找不到任何 TXT 檔案。", vbExclamation, "警告"
        GoTo CleanUp
    End If

    Do While strFile <> ""
        intIn = FreeFile
        Open strFolder & strFile For Input As #intIn

        ' 第一個檔案寫入標題列
        If blnFirst Then
            Dim lngCol As Long
            For lngCol = 1 To nFields
                wsDest.Cells(lngDestRow, lngCol).Value = fields(lngCol).FieldName
            Next lngCol
            lngDestRow = lngDestRow + 1
            blnFirst = False
        End If

        ' 逐列讀取並依欄位定義拆解
        Dim blnIsHeader As Boolean
        blnIsHeader = True
        Do While Not EOF(intIn)
            Line Input #intIn, strLine
            ' 跳過每個檔案的第一列（假設為標題列）
            If blnIsHeader Then
                blnIsHeader = False
            Else
                If Len(strLine) > 0 Then
                    For lngCol = 1 To nFields
                        Dim lngStart As Long
                        Dim lngWidth As Long
                        lngStart = fields(lngCol).StartPos
                        lngWidth = fields(lngCol).FieldWidth
                        If lngStart <= Len(strLine) Then
                            Dim strVal As String
                            strVal = Mid(strLine, lngStart, lngWidth)
                            wsDest.Cells(lngDestRow, lngCol).Value = Trim(strVal)
                        End If
                    Next lngCol
                    lngDestRow = lngDestRow + 1
                End If
            End If
        Loop

        Close #intIn
        lngFilesCount = lngFilesCount + 1
        strFile = Dir()
    Loop

    wsDest.Columns.AutoFit

CleanUp:
    Application.ScreenUpdating = True

    If lngFilesCount > 0 Then
        MsgBox "固定寬度文字檔合併完成！" & vbCrLf & _
               "共合併 " & lngFilesCount & " 個檔案，" & _
               "合計 " & lngDestRow - 2 & " 列資料。", vbInformation, "完成"
    End If
    Exit Sub

ErrHandler:
    On Error Resume Next
    Close #intIn
    On Error GoTo 0
    Application.ScreenUpdating = True
    MsgBox "合併過程發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub