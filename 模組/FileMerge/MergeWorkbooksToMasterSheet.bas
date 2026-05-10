Attribute VB_Name = "MergeWorkbooksToMasterSheet"
Option Explicit

' ============================================================
' 模組名稱：MergeWorkbooksToMasterSheet
' 功能說明：將指定資料夾中所有 Excel 活頁簿的第一個工作表
'           合併到目前活頁簿的「合併主表」工作表
' 使用方式：執行 MergeWorkbooksToMasterSheet 並選擇來源資料夾
' ============================================================

Sub MergeWorkbooksToMasterSheet()
    Dim sFolder     As String
    Dim sFile       As String
    Dim wbSource    As Workbook
    Dim wsSource    As Worksheet
    Dim wsMaster    As Worksheet
    Dim masterName  As String
    Dim nextRow     As Long
    Dim srcLastRow  As Long
    Dim srcLastCol  As Long
    Dim headerCopied As Boolean
    Dim fileCount   As Long
    
    On Error GoTo ErrHandler
    
    ' 選擇來源資料夾
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇包含 Excel 活頁簿的資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation, "取消"
            Exit Sub
        End If
        sFolder = .SelectedItems(1)
    End With
    
    If Right(sFolder, 1) <> "\" Then sFolder = sFolder & "\"
    
    ' 取得第一個 xlsx/xls 檔案
    sFile = Dir(sFolder & "*.xlsx")
    If sFile = "" Then
        sFile = Dir(sFolder & "*.xls")
        If sFile = "" Then
            MsgBox "指定資料夾中找不到 Excel 檔案。", vbExclamation, "提示"
            Exit Sub
        End If
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 建立或清空主表工作表
    masterName = "合併主表"
    On Error Resume Next
    ThisWorkbook.Sheets(masterName).Delete
    On Error GoTo ErrHandler
    
    Set wsMaster = ThisWorkbook.Sheets.Add
    wsMaster.Name = masterName
    nextRow = 1
    headerCopied = False
    fileCount = 0
    
    ' 逐一處理每個 Excel 檔案
    Do While sFile <> ""
        ' 跳過自身檔案
        If sFolder & sFile <> ThisWorkbook.FullName Then
            Set wbSource = Workbooks.Open(Filename:=sFolder & sFile, _
                                          ReadOnly:=True, UpdateLinks:=0)
            Set wsSource = wbSource.Sheets(1)
            
            srcLastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
            srcLastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
            
            If srcLastRow > 0 And srcLastCol > 0 Then
                If Not headerCopied Then
                    ' 複製標題列（含來源檔案欄）
                    wsSource.Range(wsSource.Cells(1, 1), _
                                   wsSource.Cells(1, srcLastCol)).Copy _
                        Destination:=wsMaster.Cells(1, 1)
                    wsMaster.Cells(1, srcLastCol + 1).Value = "來源檔案"
                    wsMaster.Rows(1).Font.Bold = True
                    wsMaster.Rows(1).Interior.Color = RGB(68, 114, 196)
                    wsMaster.Rows(1).Font.Color = RGB(255, 255, 255)
                    nextRow = 2
                    headerCopied = True
                End If
                
                ' 複製資料列（略過標題）
                If srcLastRow >= 2 Then
                    wsSource.Range(wsSource.Cells(2, 1), _
                                   wsSource.Cells(srcLastRow, srcLastCol)).Copy _
                        Destination:=wsMaster.Cells(nextRow, 1)
                    
                    ' 填入來源檔案名稱
                    Dim r As Long
                    For r = nextRow To nextRow + srcLastRow - 2
                        wsMaster.Cells(r, srcLastCol + 1).Value = sFile
                    Next r
                    
                    nextRow = nextRow + srcLastRow - 1
                End If
            End If
            
            wbSource.Close SaveChanges:=False
            fileCount = fileCount + 1
        End If
        
        sFile = Dir()
    Loop
    
    ' 自動調整欄寬
    If fileCount > 0 Then
        wsMaster.Columns.AutoFit
    End If
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "合併完成！" & vbCrLf & _
           "共處理 " & fileCount & " 個活頁簿。" & vbCrLf & _
           "結果已存至「" & masterName & "」工作表。", vbInformation, "完成"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    On Error Resume Next
    If Not wbSource Is Nothing Then wbSource.Close SaveChanges:=False
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub