Attribute VB_Name = "SavingPictureTool"
Option Explicit
'*************************************************************************************
'專案名稱: 底層元件
'功能描述: 將worksheet另存成圖檔
'
'版權所有: Dunk
'程式撰寫: Dunk
'撰寫日期：2018/3/23
'
'改版日期:
'改版備註:
'
'*************************************************************************************

'Sub test()
'    PrintScreenToPngFile "123", "d:\output\aaa1.png"
'    PrintScreenToPngFile "345", "d:\output\aaa2.png"
'    PrintScreenToPngFileByRange "123", "A1", "d:\output\aaa3.png"
'    SaveChartToJpg "789", "圖表 1"
'     SaveChartToJpgWithDestFile "789", "圖表 1", "d:\output\12355.jpg"
'End Sub

'另存圖表成PNG(指定range)
Function PrintScreenToPngFileByRange(ByVal sheetName As String, ByVal rangeDisplay As String _
                                                                , ByVal destFile As String)
    Dim Rng As range
    Dim wsSource As Worksheet

    Set wsSource = Sheets(sheetName)
    Set Rng = wsSource.range(rangeDisplay)
    Rng.CopyPicture

    ' Excel 存圖檔的精簡程式碼。
    With wsSource.ChartObjects.Add(1, 1, Rng.Width, Rng.Height)  '新增 圖表
        .Chart.Paste                                                '貼上 圖片
        .Chart.Export Filename:=destFile, Filtername:="png"
        .Delete                                                      '刪除 圖表
    End With

    Set Rng = Nothing
    Set wsSource = Nothing
End Function


Function PrintScreenToPngFile(ByVal sheetName As String, ByVal destFile As String)
    Dim Rng As range
    Dim wsSource As Worksheet

    Set wsSource = Sheets(sheetName)
    Set Rng = wsSource.UsedRange
    Rng.CopyPicture

    ' Excel 存圖檔的精簡程式碼。
    With wsSource.ChartObjects.Add(1, 1, Rng.Width, Rng.Height)  '新增 圖表
        .Chart.Paste                                                '貼上 圖片
        .Chart.Export Filename:=destFile, Filtername:="png"
        .Delete                                                      '刪除 圖表
    End With

    Set Rng = Nothing
    Set wsSource = Nothing
End Function

'存一般圖的程式 存在當下的資料夾
Function SaveChartToJpg(ByVal sheetName As String, ByVal chartName As String)
    Dim MyFullName As String
    MyFullName = Replace(Replace(ThisWorkbook.FullName, ".xlsm", ""), ".xls", "") & _
                 "_" & chartName

    '建立存圖路徑的程式
    With Sheets(sheetName).ChartObjects(chartName).Chart
        .Export Filename:=MyFullName & ".jpg", Filtername:="jpg"
    End With
End Function

'存一般圖的程式 存在指定的檔案
Function SaveChartToJpgWithDestFile(ByVal sheetName As String, ByVal chartName As String, _
                                    ByVal destFileName As String)

'建立存圖路徑的程式
    With Sheets(sheetName).ChartObjects(chartName).Chart
        .Export Filename:=destFileName, Filtername:="jpg"
    End With
End Function

