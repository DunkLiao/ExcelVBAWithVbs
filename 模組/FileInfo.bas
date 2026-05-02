Attribute VB_Name = "FileInfo"
Option Explicit

Sub 列出檔案相關資訊()
    Dim selectFolder As String
    
    selectFolder = Application.InputBox("請輸入資料夾路徑:", _
                "將會列出檔案路徑、大小及時間")
    
    If selectFolder = "" Then
        MsgBox "請輸入路徑"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Call 列出檔案清單(selectFolder)
    Call 列出子目錄清單(selectFolder)
    MsgBox "列出檔案資訊完畢!"
    Application.ScreenUpdating = True
End Sub


Sub 設定標題(ByVal sheetIndex As Integer)
    '清除內容並將將儲存格格式設為文字格式
    Dim pt As Range
    Dim myRange As Range
    Dim i As Integer
    Set pt = ThisWorkbook.Sheets(sheetIndex).Range("a2")
    For i = 1 To 3
        pt.Worksheet.Columns(i).ClearContents
    Next
    Set myRange = ThisWorkbook.Sheets(sheetIndex).Range("A1:C65536")
    myRange.NumberFormatLocal = "@"
            
    '設定標題
    ThisWorkbook.Sheets(sheetIndex).Range("A1").Value = "路徑"
    ThisWorkbook.Sheets(sheetIndex).Range("B1").Value = "大小"
    ThisWorkbook.Sheets(sheetIndex).Range("C1").Value = "修改時間"
End Sub

Sub 列出檔案清單(ByVal theDir As String)
    Dim pt As Range
                
    Set pt = Sheet1.Range("a2")
    Call 設定標題(1)
                
    If Len(Dir(theDir, vbDirectory)) > 0 Then
        If (GetAttr(theDir) And vbDirectory) = vbDirectory Then
            Call FileIOUtility.RetrivalFileList(theDir, pt, 0)
        End If
    End If
    
    pt.Worksheet.Columns("A:B").AutoFit
End Sub

Sub 列出子目錄清單(ByVal theDir As String)
    Dim pt As Range
                    
    Set pt = Sheet2.Range("a2")
    Call 設定標題(2)
    
    If Len(Dir(theDir, vbDirectory)) > 0 Then
        If (GetAttr(theDir) And vbDirectory) = vbDirectory Then
        Call FileIOUtility.RetrivalAllSubFolderList(theDir, pt)
        End If
    End If
    
    pt.Worksheet.Columns("A:B").AutoFit
End Sub


