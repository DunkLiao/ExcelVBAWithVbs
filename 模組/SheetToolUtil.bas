Attribute VB_Name = "SheetToolUtil"
'*************************************************************************************
' 模組名稱：SheetToolUtil（工作表工具模組）
' 功能說明：提供工作表常用操作的公用函式，包含：
'           - 取得最後一列 / 最後一欄的位置
'           - 跨工作表 / 跨活頁簿的範圍複製
'           - 範圍內的搜尋、清除與取代
'           - 外部連結管理
'
' 版權所有：台灣版權
' 程式作者：Dunk
' 建立日期：2017/8/14
'
' 修改紀錄：
'   2017/8/16  新增清除指定範圍內容（僅清文字數值）
'   2018/10/31 新增取代指定字元
'   2023/6/15  新增複製 UsedRange
'   2025/05/03 全面重構：移除 Select/Selection 反模式、修正型別宣告、補上中文說明
'*************************************************************************************
Option Explicit

' ══════════════════════════════════════════════════════════════════════
' 【區段一】取得最後列 / 欄位置
' ══════════════════════════════════════════════════════════════════════

'--------------------------------------------------
' 函式：GetLastRowNum
' 用途：依工作表名稱與欄位英文代號，取得該欄最後一列的列號
' 參數：sheetName   - 工作表名稱
'       columnAlpha - 欄位英文代號（例如 "A"、"B"）
' 回傳：最後一列列號（Long）；若工作表不存在則觸發錯誤
' 備註：自動判斷 xlsx（1048576 列）或舊版 xls（65536 列）格式
'--------------------------------------------------
Function GetLastRowNum(ByVal sheetName As String, ByVal columnAlpha As String) As Long
    Dim lastRowNum As Long
    On Error GoTo xlsFormat
    ' xlsx 格式：最大列數 1,048,576
    lastRowNum = ThisWorkbook.Sheets(sheetName).Range(columnAlpha & "1048576").End(xlUp).Row
    GetLastRowNum = lastRowNum
    Exit Function
xlsFormat:
    ' 舊版 xls 格式：最大列數 65,536
    lastRowNum = ThisWorkbook.Sheets(sheetName).Range(columnAlpha & "65536").End(xlUp).Row
    GetLastRowNum = lastRowNum
End Function

'--------------------------------------------------
' 函式：GetLastRowNumBySheet
' 用途：依工作表物件與欄位英文代號，取得該欄最後一列的列號
' 參數：mySheet     - 工作表物件（Worksheet）
'       columnAlpha - 欄位英文代號（例如 "A"、"B"）
' 回傳：最後一列列號（Long）
' 備註：自動判斷 xlsx / xls 格式
'--------------------------------------------------
Function GetLastRowNumBySheet(ByVal mySheet As Worksheet, ByVal columnAlpha As String) As Long
    Dim lastRowNum As Long
    On Error GoTo xlsFormat
    ' xlsx 格式：最大列數 1,048,576
    lastRowNum = mySheet.Range(columnAlpha & "1048576").End(xlUp).Row
    GetLastRowNumBySheet = lastRowNum
    Exit Function
xlsFormat:
    ' 舊版 xls 格式：最大列數 65,536
    lastRowNum = mySheet.Range(columnAlpha & "65536").End(xlUp).Row
    GetLastRowNumBySheet = lastRowNum
End Function

'--------------------------------------------------
' 函式：GetLastColumnBySheet
' 用途：依工作表物件與指定列號，取得該列最後一欄的欄號
' 參數：ws     - 工作表物件（Worksheet）
'       rowNum - 要查詢的列號
' 回傳：最後一欄欄號（Long）
'--------------------------------------------------
Function GetLastColumnBySheet(ByVal ws As Worksheet, ByVal rowNum As Variant) As Long
    Dim lastColumnNum As Long
    With ws
        lastColumnNum = .Cells(rowNum, .Columns.Count).End(xlToLeft).Column
    End With
    GetLastColumnBySheet = lastColumnNum
End Function

'--------------------------------------------------
' 函式：GetLastColumn
' 用途：依工作表名稱與指定列號，取得該列最後一欄的欄號
' 參數：sheetName - 工作表名稱
'       rowNum    - 要查詢的列號
' 回傳：最後一欄欄號（Long）
'--------------------------------------------------
Function GetLastColumn(ByVal sheetName As String, ByVal rowNum As Variant) As Long
    Dim ws As Worksheet
    Dim lastColumnNum As Long
    Set ws = ThisWorkbook.Sheets(sheetName)
    With ws
        lastColumnNum = .Cells(rowNum, .Columns.Count).End(xlToLeft).Column
    End With
    Set ws = Nothing
    GetLastColumn = lastColumnNum
End Function

'--------------------------------------------------
' 函式：GetLastColumnAddressBySheet
' 用途：依工作表物件與指定列號，取得該列最後一欄的儲存格位址（無 $ 號）
' 參數：ws     - 工作表物件（Worksheet）
'       rowNum - 要查詢的列號
' 回傳：儲存格位址字串，例如 "D5"（String）
'--------------------------------------------------
Function GetLastColumnAddressBySheet(ByVal ws As Worksheet, ByVal rowNum As Variant) As String
    Dim lastColumnNum As Long
    Dim addr As String
    With ws
        lastColumnNum = .Cells(rowNum, .Columns.Count).End(xlToLeft).Column
        addr = .Cells(rowNum, lastColumnNum).Address(False, False)
    End With
    GetLastColumnAddressBySheet = addr
End Function

'--------------------------------------------------
' 函式：GetColumnNameBySheet
' 用途：依工作表物件與欄號，取得對應的欄位英文代號
' 參數：ws        - 工作表物件（Worksheet）
'       columnNum - 欄號（數字）
' 回傳：欄位英文代號，例如 "A"、"AB"（String）
'--------------------------------------------------
Function GetColumnNameBySheet(ByVal ws As Worksheet, ByVal columnNum As Variant) As String
    Dim addr As String
    With ws
        ' 取得第 1 列指定欄的位址，再移除列號 "1" 得到欄名
        addr = Replace(.Cells(1, columnNum).Address(False, False), "1", "")
    End With
    GetColumnNameBySheet = addr
End Function


' ══════════════════════════════════════════════════════════════════════
' 【區段二】範圍複製
' ══════════════════════════════════════════════════════════════════════

'--------------------------------------------------
' 函式：CopyRange
' 用途：將來源工作表（依名稱）指定欄的某範圍列，以「僅貼值」方式
'       複製到目的工作表（依名稱）的指定起始位置
' 參數：sourceSheetName - 來源工作表名稱
'       sourceColumnName - 來源欄位英文代號（例如 "A"）
'       sourceRowStart  - 來源起始列號
'       sourceRowEnd    - 來源結束列號
'       destSheetName   - 目的工作表名稱
'       destColumnName  - 目的欄位英文代號
'       destRowStart    - 目的起始列號
'--------------------------------------------------
Function CopyRange(ByVal sourceSheetName As String, ByVal sourceColumnName As String _
                 , ByVal sourceRowStart As Long, ByVal sourceRowEnd As Long _
                 , ByVal destSheetName As String, ByVal destColumnName As String _
                 , ByVal destRowStart As Long)

    Dim srcRange As Range
    Dim dstRange As Range
    Set srcRange = ThisWorkbook.Sheets(sourceSheetName).Range( _
        sourceColumnName & sourceRowStart & ":" & sourceColumnName & sourceRowEnd)
    Set dstRange = ThisWorkbook.Sheets(destSheetName).Range(destColumnName & destRowStart)

    srcRange.Copy
    dstRange.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
                          SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

    Set srcRange = Nothing
    Set dstRange = Nothing
End Function

'--------------------------------------------------
' 函式：CopyRangeWithPlace
' 用途：將來源工作表（依名稱）指定矩形範圍，以「僅貼值」方式
'       複製到目的工作表（依名稱）的指定起始位置
'       （與 CopyRange 的差異在於：本函式可指定起始欄與結束欄兩個欄位）
' 參數：sourceSheetName      - 來源工作表名稱
'       sourceColumnNameStart - 來源起始欄位英文代號（例如 "A"）
'       sourceRowStart        - 來源起始列號
'       sourceColumnNameEnd   - 來源結束欄位英文代號（例如 "C"）
'       sourceRowEnd          - 來源結束列號
'       destSheetName         - 目的工作表名稱
'       destColumnName        - 目的起始欄位英文代號
'       destRowStart          - 目的起始列號
'--------------------------------------------------
Function CopyRangeWithPlace(ByVal sourceSheetName As String _
                           , ByVal sourceColumnNameStart As String, ByVal sourceRowStart As Long _
                           , ByVal sourceColumnNameEnd As String, ByVal sourceRowEnd As Long _
                           , ByVal destSheetName As String _
                           , ByVal destColumnName As String, ByVal destRowStart As Long)

    Dim srcRange As Range
    Dim dstRange As Range
    Set srcRange = ThisWorkbook.Sheets(sourceSheetName).Range( _
        sourceColumnNameStart & sourceRowStart & ":" & sourceColumnNameEnd & sourceRowEnd)
    Set dstRange = ThisWorkbook.Sheets(destSheetName).Range(destColumnName & destRowStart)

    srcRange.Copy
    dstRange.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
                          SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

    Set srcRange = Nothing
    Set dstRange = Nothing
End Function

'--------------------------------------------------
' 函式：CopyRangeBySheet
' 用途：將來源工作表物件指定欄的某範圍列，以「僅貼值」方式
'       複製到目的工作表物件的指定起始位置
' 參數：sourceSheet      - 來源工作表物件（Worksheet）
'       sourceColumnName - 來源欄位英文代號
'       sourceRowStart   - 來源起始列號
'       sourceRowEnd     - 來源結束列號
'       destSheet        - 目的工作表物件（Worksheet）
'       destColumnName   - 目的欄位英文代號
'       destRowStart     - 目的起始列號
'--------------------------------------------------
Function CopyRangeBySheet(ByVal sourceSheet As Worksheet, ByVal sourceColumnName As String _
                         , ByVal sourceRowStart As Long, ByVal sourceRowEnd As Long _
                         , ByVal destSheet As Worksheet, ByVal destColumnName As String _
                         , ByVal destRowStart As Long)

    Dim srcRange As Range
    Dim dstRange As Range
    Set srcRange = sourceSheet.Range( _
        sourceColumnName & sourceRowStart & ":" & sourceColumnName & sourceRowEnd)
    Set dstRange = destSheet.Range(destColumnName & destRowStart)

    srcRange.Copy
    dstRange.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
                          SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

    Set srcRange = Nothing
    Set dstRange = Nothing
End Function

'--------------------------------------------------
' 函式：CopyRangeByFile
' 用途：將來源活頁簿（依檔名）的指定範圍，以「僅貼值」方式
'       複製到目的活頁簿（依檔名）的指定範圍
'       ＊兩個活頁簿必須都已開啟
' 參數：sourceFileName    - 來源活頁簿檔名（含副檔名）
'       sourceSheetName   - 來源工作表名稱
'       sourceRangeDisplay - 來源範圍位址字串（例如 "A1:C10"）
'       destFileName      - 目的活頁簿檔名（含副檔名）
'       destSheetName     - 目的工作表名稱
'       destRangeDisplay  - 目的起始範圍位址字串
'--------------------------------------------------
Function CopyRangeByFile(ByVal sourceFileName As String, ByVal sourceSheetName As String _
                        , ByVal sourceRangeDisplay As String _
                        , ByVal destFileName As String, ByVal destSheetName As String _
                        , ByVal destRangeDisplay As String)

    ' 複製來源範圍到剪貼簿
    Workbooks(sourceFileName).Worksheets(sourceSheetName).Range(sourceRangeDisplay).Copy

    ' 以僅貼值方式貼到目的範圍
    Workbooks(destFileName).Worksheets(destSheetName).Range(destRangeDisplay).PasteSpecial _
        Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    Application.CutCopyMode = False
End Function

'--------------------------------------------------
' 函式：CopyRangeFomulaBySheet
' 用途：將來源工作表物件指定欄的某範圍列，以「僅貼公式」方式
'       複製到目的工作表物件的指定起始位置
' 參數：sourceSheet      - 來源工作表物件（Worksheet）
'       sourceColumnName - 來源欄位英文代號
'       sourceRowStart   - 來源起始列號
'       sourceRowEnd     - 來源結束列號
'       destSheet        - 目的工作表物件（Worksheet）
'       destColumnName   - 目的欄位英文代號
'       destRowStart     - 目的起始列號
'--------------------------------------------------
Function CopyRangeFomulaBySheet(ByVal sourceSheet As Worksheet, ByVal sourceColumnName As String _
                               , ByVal sourceRowStart As Long, ByVal sourceRowEnd As Long _
                               , ByVal destSheet As Worksheet, ByVal destColumnName As String _
                               , ByVal destRowStart As Long)

    Dim srcRange As Range
    Dim dstRange As Range
    Set srcRange = sourceSheet.Range( _
        sourceColumnName & sourceRowStart & ":" & sourceColumnName & sourceRowEnd)
    Set dstRange = destSheet.Range(destColumnName & destRowStart)

    srcRange.Copy
    dstRange.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
                          SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

    Set srcRange = Nothing
    Set dstRange = Nothing
End Function

'--------------------------------------------------
' 函式：CopyRangeFormatBySheet
' 用途：將來源工作表物件指定欄的某範圍列，以「僅貼格式」方式
'       複製到目的工作表物件的指定起始位置
' 參數：sourceSheet      - 來源工作表物件（Worksheet）
'       sourceColumnName - 來源欄位英文代號
'       sourceRowStart   - 來源起始列號
'       sourceRowEnd     - 來源結束列號
'       destSheet        - 目的工作表物件（Worksheet）
'       destColumnName   - 目的欄位英文代號
'       destRowStart     - 目的起始列號
'--------------------------------------------------
Function CopyRangeFormatBySheet(ByVal sourceSheet As Worksheet, ByVal sourceColumnName As String _
                               , ByVal sourceRowStart As Long, ByVal sourceRowEnd As Long _
                               , ByVal destSheet As Worksheet, ByVal destColumnName As String _
                               , ByVal destRowStart As Long)

    Dim srcRange As Range
    Dim dstRange As Range
    Set srcRange = sourceSheet.Range( _
        sourceColumnName & sourceRowStart & ":" & sourceColumnName & sourceRowEnd)
    Set dstRange = destSheet.Range(destColumnName & destRowStart)

    srcRange.Copy
    dstRange.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                          SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False

    Set srcRange = Nothing
    Set dstRange = Nothing
End Function

'--------------------------------------------------
' 函式：CopyUsedRangeToAnotherSheetWithoutFormatting
' 用途：將來源工作表的 UsedRange（已使用範圍）以「僅貼值」方式
'       複製到目的工作表的指定位置（不複製格式）
' 參數：wsSource         - 來源工作表物件（Worksheet）
'       wsDest           - 目的工作表物件（Worksheet）
'       destRangeDisplay - 目的起始範圍位址字串（例如 "A1"）
' 備註：原始版本有 Bug（使用未宣告的 destRange 變數），已修正為 destRangeDisplay
'--------------------------------------------------
Private Function CopyUsedRangeToAnotherSheetWithoutFormatting( _
    ByVal wsSource As Worksheet, ByVal wsDest As Worksheet, ByVal destRangeDisplay As String)

    wsSource.UsedRange.Copy
    wsDest.Range(destRangeDisplay).PasteSpecial _
        Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
End Function

'--------------------------------------------------
' 函式：CopySheetFromFile
' 用途：開啟外部 Excel 檔案，將其第一張工作表複製到目前活頁簿的最後，
'       並重新命名後關閉來源檔案
' 參數：fileName        - 來源檔案完整路徑
'       resultSheetName - 複製後的工作表新名稱
' 備註：需搭配 FileIOUtility 模組的 GetFileNameWithoutFolder 函式
'--------------------------------------------------
Function CopySheetFromFile(ByVal fileName As String, ByVal resultSheetName As String)
    Dim windowName As String
    windowName = FileIOUtility.GetFileNameWithoutFolder(fileName)

    Workbooks.Open fileName:=fileName
    Windows(windowName).Activate

    ' 將來源第一張工作表複製到目前活頁簿最後
    Sheets(1).Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = resultSheetName

    Windows(windowName).Close SaveChanges:=False
End Function


' ══════════════════════════════════════════════════════════════════════
' 【區段三】搜尋
' ══════════════════════════════════════════════════════════════════════

'--------------------------------------------------
' 函式：FindFirstValueRowInRange
' 用途：在指定範圍內搜尋（完全符合），回傳第一個符合值所在的列號
' 參數：myRange - 搜尋範圍（Range）
'       findStr - 要搜尋的字串
' 回傳：找到時回傳列號（Long）；找不到時回傳 -1
'--------------------------------------------------
Function FindFirstValueRowInRange(ByVal myRange As Range, ByVal findStr As String) As Long
    Dim findRange As Range
    With myRange
        Set findRange = .Find(findStr, After:=.Cells(.Cells.Count), _
                              LookIn:=xlValues, LookAt:=xlWhole)
    End With

    If Not findRange Is Nothing Then
        FindFirstValueRowInRange = findRange.Row
        Set findRange = Nothing
    Else
        FindFirstValueRowInRange = -1
    End If
End Function

'--------------------------------------------------
' 函式：FindValueRowsInRange
' 用途：在指定範圍內搜尋（完全符合），回傳所有符合值所在的列號陣列
' 參數：myRange - 搜尋範圍（Range）
'       findStr - 要搜尋的字串
' 回傳：包含所有符合列號的 Variant 陣列；若無符合則回傳空陣列
' 備註：原始版本的 firstAddress 未宣告，已修正
'--------------------------------------------------
Function FindValueRowsInRange(ByVal myRange As Range, ByVal findStr As String) As Variant
    Dim findRange As Range
    Dim resultArray() As Variant
    Dim firstAddress As String
    Dim counter As Long

    ReDim resultArray(65536)
    counter = 0

    With myRange
        Set findRange = .Find(findStr, After:=.Cells(.Cells.Count), _
                              LookIn:=xlValues, LookAt:=xlWhole)
        If Not findRange Is Nothing Then
            firstAddress = findRange.Address
            Do
                Set findRange = .FindNext(findRange)
                resultArray(counter) = findRange.Row
                counter = counter + 1
            Loop While Not findRange Is Nothing And findRange.Address <> firstAddress
        End If
    End With

    If counter > 0 Then
        counter = counter - 1
    End If

    ReDim Preserve resultArray(counter)
    FindValueRowsInRange = resultArray
End Function

'--------------------------------------------------
' 函式：FindFirstValueColumnInRange
' 用途：在指定範圍內搜尋（完全符合），回傳第一個符合值所在的欄號
' 參數：myRange - 搜尋範圍（Range）
'       findStr - 要搜尋的字串
' 回傳：找到時回傳欄號（Long）；找不到時回傳 -1
'--------------------------------------------------
Function FindFirstValueColumnInRange(ByVal myRange As Range, ByVal findStr As String) As Long
    Dim findRange As Range
    With myRange
        Set findRange = .Find(findStr, After:=.Cells(.Cells.Count), _
                              LookIn:=xlValues, LookAt:=xlWhole)
    End With

    If Not findRange Is Nothing Then
        FindFirstValueColumnInRange = findRange.Column
        Set findRange = Nothing
    Else
        FindFirstValueColumnInRange = -1
    End If
End Function

'--------------------------------------------------
' 函式：FindFirstValueAbsAddressInRange
' 用途：在指定範圍內搜尋（完全符合），回傳第一個符合值的絕對位址（含 $）
' 參數：myRange - 搜尋範圍（Range）
'       findStr - 要搜尋的字串
' 回傳：找到時回傳絕對位址字串，例如 "$A$5"（String）；找不到時回傳 ""
'--------------------------------------------------
Function FindFirstValueAbsAddressInRange(ByVal myRange As Range, ByVal findStr As String) As String
    Dim findRange As Range
    With myRange
        Set findRange = .Find(findStr, After:=.Cells(.Cells.Count), _
                              LookIn:=xlValues, LookAt:=xlWhole)
    End With

    If Not findRange Is Nothing Then
        FindFirstValueAbsAddressInRange = findRange.Address
        Set findRange = Nothing
    Else
        FindFirstValueAbsAddressInRange = ""
    End If
End Function

'--------------------------------------------------
' 函式：FindFirstValueAddressInRangeFullMatch
' 用途：在指定範圍內搜尋（完全符合），回傳第一個符合值的相對位址（不含 $）
' 參數：myRange - 搜尋範圍（Range）
'       findStr - 要搜尋的字串
' 回傳：找到時回傳相對位址字串，例如 "A5"（String）；找不到時回傳 ""
'--------------------------------------------------
Function FindFirstValueAddressInRangeFullMatch(ByVal myRange As Range, ByVal findStr As String) As String
    Dim findRange As Range
    With myRange
        Set findRange = .Find(findStr, After:=.Cells(.Cells.Count), _
                              LookIn:=xlValues, LookAt:=xlWhole)
    End With

    If Not findRange Is Nothing Then
        FindFirstValueAddressInRangeFullMatch = Replace(findRange.Address, "$", "")
        Set findRange = Nothing
    Else
        FindFirstValueAddressInRangeFullMatch = ""
    End If
End Function

'--------------------------------------------------
' 函式：FindFirstValueAddressInRangePartialMatch
' 用途：在指定範圍內搜尋（部分符合），回傳第一個符合值的相對位址（不含 $）
' 參數：myRange - 搜尋範圍（Range）
'       findStr - 要搜尋的字串（支援部分比對）
' 回傳：找到時回傳相對位址字串，例如 "A5"（String）；找不到時回傳 ""
'--------------------------------------------------
Function FindFirstValueAddressInRangePartialMatch(ByVal myRange As Range, ByVal findStr As String) As String
    Dim findRange As Range
    With myRange
        Set findRange = .Find(findStr, After:=.Cells(.Cells.Count), _
                              LookIn:=xlValues, LookAt:=xlPart)
    End With

    If Not findRange Is Nothing Then
        FindFirstValueAddressInRangePartialMatch = Replace(findRange.Address, "$", "")
        Set findRange = Nothing
    Else
        FindFirstValueAddressInRangePartialMatch = ""
    End If
End Function


' ══════════════════════════════════════════════════════════════════════
' 【區段四】資料清除與取代
' ══════════════════════════════════════════════════════════════════════

'--------------------------------------------------
' 函式：ClearContent
' 用途：清除指定範圍內所有內容（值、公式），保留格式
' 參數：myRange - 要清除的範圍（Range）
'--------------------------------------------------
Function ClearContent(ByVal myRange As Range)
    myRange.ClearContents
End Function

'--------------------------------------------------
' 函式：ClearContentOnlyText
' 用途：清除指定範圍內的常數內容（文字、數值），保留公式與格式
' 參數：myRange - 要清除的範圍（Range）
' 備註：若範圍內無常數儲存格，忽略錯誤並正常結束
'--------------------------------------------------
Function ClearContentOnlyText(ByVal myRange As Range)
    On Error GoTo NoData
    myRange.SpecialCells(xlCellTypeConstants, 23).ClearContents
NoData:
End Function

'--------------------------------------------------
' 函式：ReplaceValueInRange
' 用途：在指定範圍內，將所有符合（完全符合）的值以 replaceString 取代
' 參數：myRange       - 搜尋範圍（Range）
'       findStr       - 要搜尋的字串
'       replaceString - 取代後的新值
' 備註：原始版本的 firstAddress 未宣告，已修正
'--------------------------------------------------
Function ReplaceValueInRange(ByVal myRange As Range, ByVal findStr As String, ByVal replaceString As Variant)
    Dim findRange As Range
    Dim firstAddress As String

    With myRange
        Set findRange = .Find(findStr, After:=.Cells(.Cells.Count), _
                              LookIn:=xlValues, LookAt:=xlWhole)
        If Not findRange Is Nothing Then
            firstAddress = findRange.Address
            Do
                findRange.Value = replaceString
                Set findRange = .FindNext(findRange)
            Loop While Not findRange Is Nothing And findRange.Address <> firstAddress
        End If
    End With

    Set findRange = Nothing
End Function


' ══════════════════════════════════════════════════════════════════════
' 【區段五】活頁簿連結與其他工具
' ══════════════════════════════════════════════════════════════════════

'--------------------------------------------------
' 函式：BreakAllLinks
' 用途：中斷目前活頁簿內所有外部 Excel 連結（xlLinkTypeExcelLinks）
' 備註：若活頁簿中沒有任何外部連結，函式安全地結束
'--------------------------------------------------
Function BreakAllLinks()
    Dim links As Variant
    Dim i As Integer

    links = ThisWorkbook.LinkSources(Type:=xlLinkTypeExcelLinks)
    On Error GoTo leave

    For i = 1 To UBound(links)
        ThisWorkbook.BreakLink Name:=links(i), Type:=xlLinkTypeExcelLinks
    Next i
leave:
    Exit Function
End Function

'--------------------------------------------------
' 函式：CheckValueFromFile
' 用途：開啟外部 Excel 檔案，檢查其第一張工作表指定儲存格的值
'       是否包含指定字串，然後關閉檔案
' 參數：fileName   - 外部檔案完整路徑
'       rangeValue - 要檢查的儲存格位址（例如 "A1"）
'       findText   - 要搜尋的字串
' 回傳：True 表示找到，False 表示未找到（Boolean）
' 備註：需搭配 FileIOUtility 模組的 GetFileNameWithoutFolder 函式
'--------------------------------------------------
Function CheckValueFromFile(ByVal fileName As String, ByVal rangeValue As String, ByVal findText As String) As Boolean
    Dim windowName As String
    Dim isFind As Boolean
    windowName = FileIOUtility.GetFileNameWithoutFolder(fileName)

    Workbooks.Open fileName:=fileName
    Windows(windowName).Activate

    isFind = (InStr(1, Sheets(1).Range(rangeValue).Value, findText) > 0)

    Windows(windowName).Close SaveChanges:=False
    CheckValueFromFile = isFind
End Function