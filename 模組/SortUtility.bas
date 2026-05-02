Attribute VB_Name = "SortUtility"
Sub 依照修改日期遞減排序()
Attribute 依照修改日期遞減排序.VB_ProcData.VB_Invoke_Func = " \n14"
'檔案清單
    ActiveWorkbook.Worksheets("檔案清單").ListObjects("檔案清單_表格").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("檔案清單").ListObjects("檔案清單_表格").Sort.SortFields.Add _
        Key:=Range("檔案清單_表格[[#All],[修改時間]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("檔案清單").ListObjects("檔案清單_表格").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
'子目錄清單
    ActiveWorkbook.Worksheets("子目錄清單").ListObjects("檔案清單_目錄").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("子目錄清單").ListObjects("檔案清單_目錄").Sort.SortFields.Add _
        Key:=Range("檔案清單_目錄[[#All],[修改時間]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("子目錄清單").ListObjects("檔案清單_目錄").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
