Attribute VB_Name = "ExportPDFByGroup"
Option Explicit

' ============================================================
' 模組名稱：ExportPDFByGroup
' 功能說明：依工作表名稱前綴分組，將同群組的工作表
'           合併匯出為同一個 PDF 檔案
' 使用方式：設定前綴分隔符號（預設 "_"），執行後選擇輸出資料夾
' 範例：工作表 "Q1_北區" 與 "Q1_南區" 會合併輸出為 "Q1.pdf"
' ============================================================

Sub ExportPDFByGroup()
    Dim sFolder     As String
    Dim ws          As Worksheet
    Dim dictGroups  As Object
    Dim separator   As String
    Dim groupKey    As String
    Dim prefixEnd   As Integer
    Dim groupVar    As Variant
    Dim arrSheets() As String
    Dim shList      As String
    Dim pdfPath     As String
    Dim exportCount As Long
    
    On Error GoTo ErrHandler
    
    ' 選擇輸出資料夾
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇 PDF 輸出資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation, "取消"
            Exit Sub
        End If
        sFolder = .SelectedItems(1)
    End With
    
    If Right(sFolder, 1) <> "\" Then sFolder = sFolder & "\"
    
    ' 前綴分隔符號
    separator = InputBox("請輸入前綴分隔符號（預設為底線 _ ）：", "分隔符號", "_")
    If separator = "" Then separator = "_"
    
    Set dictGroups = CreateObject("Scripting.Dictionary")
    
    ' 掃描所有工作表，依前綴分組
    For Each ws In ThisWorkbook.Sheets
        prefixEnd = InStr(ws.Name, separator)
        If prefixEnd > 1 Then
            groupKey = Left(ws.Name, prefixEnd - 1)
        Else
            groupKey = ws.Name
        End If
        
        If dictGroups.Exists(groupKey) Then
            dictGroups(groupKey) = dictGroups(groupKey) & "|" & ws.Name
        Else
            dictGroups.Add groupKey, ws.Name
        End If
    Next ws
    
    Application.ScreenUpdating = False
    exportCount = 0
    
    ' 逐群組匯出 PDF
    For Each groupVar In dictGroups.Keys
        shList = CStr(dictGroups(groupVar))
        arrSheets = Split(shList, "|")
        
        ' 選取群組工作表
        ThisWorkbook.Sheets(arrSheets(0)).Select
        
        Dim k As Integer
        For k = 1 To UBound(arrSheets)
            ThisWorkbook.Sheets(arrSheets(k)).Select Replace:=False
        Next k
        
        ' 設定 PDF 輸出路徑
        pdfPath = sFolder & CStr(groupVar) & ".pdf"
        
        ' 匯出 PDF
        ActiveWindow.SelectedSheets.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=pdfPath, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
        
        exportCount = exportCount + 1
    Next groupVar
    
    ' 回到第一張工作表
    ThisWorkbook.Sheets(1).Select
    
    Application.ScreenUpdating = True
    MsgBox "PDF 匯出完成！" & vbCrLf & _
           "共匯出 " & exportCount & " 個群組 PDF。" & vbCrLf & _
           "輸出位置：" & sFolder, vbInformation, "完成"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub