Attribute VB_Name = "MergeExcelByKeyColumn"
Option Explicit

' ============================================================
' 範例：以關鍵欄位（Key Column）水平合併兩個 Excel 檔案
' 功能：選擇主檔與副檔，指定兩者的關鍵欄位欄號，
'       依關鍵值比對後，將副檔的資料欄位附加至主檔右側
' ============================================================

Sub MergeExcelByKeyColumn()
    On Error GoTo ErrHandler

    Dim strMainFile As String
    Dim strSubFile  As String
    Dim wbMain      As Workbook
    Dim wbSub       As Workbook
    Dim wsMain      As Worksheet
    Dim wsSub       As Worksheet
    Dim wsDest      As Worksheet
    Dim lngKeyColMain   As Long
    Dim lngKeyColSub    As Long
    Dim lngMainLastRow  As Long
    Dim lngSubLastRow   As Long
    Dim lngMainLastCol  As Long
    Dim lngSubLastCol   As Long
    Dim lngRow      As Long
    Dim lngSubRow   As Long
    Dim strKeyVal   As String
    Dim strInput    As String

    ' --- 選擇主檔 ---
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "請選擇主檔（Main File）"
        .Filters.Clear
        .Filters.Add "Excel 檔案", "*.xlsx;*.xls;*.xlsm"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation, "提示"
            Exit Sub
        End If
        strMainFile = .SelectedItems(1)
    End With

    ' --- 選擇副檔 ---
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "請選擇副檔（Sub File）"
        .Filters.Clear
        .Filters.Add "Excel 檔案", "*.xlsx;*.xls;*.xlsm"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation, "提示"
            Exit Sub
        End If
        strSubFile = .SelectedItems(1)
    End With

    ' --- 取得關鍵欄位欄號 ---
    strInput = InputBox("請輸入主檔的關鍵欄位欄號（如：1 代表 A 欄）：", "主檔關鍵欄", "1")
    If Not IsNumeric(strInput) Or CLng(strInput) < 1 Then
        MsgBox "欄號輸入無效，已取消。", vbExclamation, "警告"
        Exit Sub
    End If
    lngKeyColMain = CLng(strInput)

    strInput = InputBox("請輸入副檔的關鍵欄位欄號（如：1 代表 A 欄）：", "副檔關鍵欄", "1")
    If Not IsNumeric(strInput) Or CLng(strInput) < 1 Then
        MsgBox "欄號輸入無效，已取消。", vbExclamation, "警告"
        Exit Sub
    End If
    lngKeyColSub = CLng(strInput)

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wbMain = Workbooks.Open(Filename:=strMainFile, ReadOnly:=True)
    Set wsMain = wbMain.Sheets(1)
    Set wbSub = Workbooks.Open(Filename:=strSubFile, ReadOnly:=True)
    Set wsSub = wbSub.Sheets(1)

    lngMainLastRow = wsMain.Cells(wsMain.Rows.Count, lngKeyColMain).End(xlUp).Row
    lngSubLastRow = wsSub.Cells(wsSub.Rows.Count, lngKeyColSub).End(xlUp).Row
    lngMainLastCol = wsMain.Cells(1, wsMain.Columns.Count).End(xlToLeft).Column
    lngSubLastCol = wsSub.Cells(1, wsSub.Columns.Count).End(xlToLeft).Column

    ' --- 建立目標工作表 ---
    Set wsDest = ThisWorkbook.Worksheets.Add
    wsDest.Name = "KeyMerged_" & Format(Now(), "mmddHHmm")

    ' 複製主檔全部資料
    wsMain.Cells(1, 1).Resize(lngMainLastRow, lngMainLastCol).Copy _
        Destination:=wsDest.Cells(1, 1)

    ' 寫入副檔標題（排除 Key 欄）至目標表右側
    Dim lngDestStartCol As Long
    lngDestStartCol = lngMainLastCol + 1
    Dim lngSubCol As Long
    Dim lngWriteCol As Long
    lngWriteCol = lngDestStartCol
    For lngSubCol = 1 To lngSubLastCol
        If lngSubCol <> lngKeyColSub Then
            wsDest.Cells(1, lngWriteCol).Value = wsSub.Cells(1, lngSubCol).Value
            lngWriteCol = lngWriteCol + 1
        End If
    Next lngSubCol

    ' 逐列比對主檔 Key，查找副檔對應資料
    For lngRow = 2 To lngMainLastRow
        strKeyVal = CStr(wsDest.Cells(lngRow, lngKeyColMain).Value)
        For lngSubRow = 2 To lngSubLastRow
            If CStr(wsSub.Cells(lngSubRow, lngKeyColSub).Value) = strKeyVal Then
                lngWriteCol = lngDestStartCol
                For lngSubCol = 1 To lngSubLastCol
                    If lngSubCol <> lngKeyColSub Then
                        wsDest.Cells(lngRow, lngWriteCol).Value = wsSub.Cells(lngSubRow, lngSubCol).Value
                        lngWriteCol = lngWriteCol + 1
                    End If
                Next lngSubCol
                Exit For
            End If
        Next lngSubRow
    Next lngRow

    wbMain.Close SaveChanges:=False
    wbSub.Close SaveChanges:=False

    wsDest.Columns.AutoFit
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "關鍵欄位合併完成！結果在工作表：" & wsDest.Name, vbInformation, "完成"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "合併過程發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub