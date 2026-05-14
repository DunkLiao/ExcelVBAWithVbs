Attribute VB_Name = "EngineeringFormulaExample"
Option Explicit
'*************************************************************************************
'模組名稱: EngineeringFormulaExample
'功能說明: 在 Excel 中建立工程函數公式範例（CONVERT、DELTA、GESTEP、BIN2DEC 等）
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/14
'
'*************************************************************************************

' 測試用入口
Sub TestEngineeringFormulas()
    Call CreateEngineeringFormulas("工程函數範例")
End Sub

' 建立工程函數公式範例
Sub CreateEngineeringFormulas(ByVal sheetName As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = GetOrCreateEngWs(sheetName)
    ws.Cells.Clear

    ' 設定標題
    ws.Range("A1").Value = "函數名稱"
    ws.Range("B1").Value = "公式結果"
    ws.Range("C1").Value = "說明"
    ws.Range("A1:C1").Font.Bold = True

    ' CONVERT - 公里轉英里
    ws.Range("A2").Value = "CONVERT（公里→英里）"
    ws.Range("B2").Formula = "=CONVERT(1,""km"",""mi"")"
    ws.Range("C2").Value = "將 1 公里換算為英里"

    ' CONVERT - 攝氏轉華氏
    ws.Range("A3").Value = "CONVERT（攝氏→華氏）"
    ws.Range("B3").Formula = "=CONVERT(100,""C"",""F"")"
    ws.Range("C3").Value = "將 100°C 換算為°F"

    ' CONVERT - 公斤轉磅
    ws.Range("A4").Value = "CONVERT（公斤→磅）"
    ws.Range("B4").Formula = "=CONVERT(1,""kg"",""lbm"")"
    ws.Range("C4").Value = "將 1 公斤換算為磅"

    ' DELTA - 相等判斷（相等回傳 1）
    ws.Range("A5").Value = "DELTA（相等）"
    ws.Range("B5").Formula = "=DELTA(5,5)"
    ws.Range("C5").Value = "5 = 5，回傳 1"

    ' DELTA - 不相等（不相等回傳 0）
    ws.Range("A6").Value = "DELTA（不相等）"
    ws.Range("B6").Formula = "=DELTA(3,5)"
    ws.Range("C6").Value = "3 <> 5，回傳 0"

    ' GESTEP - 大於等於步階
    ws.Range("A7").Value = "GESTEP（大於步階）"
    ws.Range("B7").Formula = "=GESTEP(10,5)"
    ws.Range("C7").Value = "10 >= 5，回傳 1"

    ' GESTEP - 小於步階
    ws.Range("A8").Value = "GESTEP（小於步階）"
    ws.Range("B8").Formula = "=GESTEP(3,5)"
    ws.Range("C8").Value = "3 < 5，回傳 0"

    ' BIN2DEC - 二進位轉十進位
    ws.Range("A9").Value = "BIN2DEC（2 進位→10 進位）"
    ws.Range("B9").Formula = "=BIN2DEC(""1010"")"
    ws.Range("C9").Value = "二進位 1010 = 十進位 10"

    ' DEC2BIN - 十進位轉二進位
    ws.Range("A10").Value = "DEC2BIN（10 進位→2 進位）"
    ws.Range("B10").Formula = "=DEC2BIN(10)"
    ws.Range("C10").Value = "十進位 10 = 二進位 1010"

    ' DEC2HEX - 十進位轉十六進位
    ws.Range("A11").Value = "DEC2HEX（10 進位→16 進位）"
    ws.Range("B11").Formula = "=DEC2HEX(255)"
    ws.Range("C11").Value = "十進位 255 = 十六進位 FF"

    ' HEX2DEC - 十六進位轉十進位
    ws.Range("A12").Value = "HEX2DEC（16 進位→10 進位）"
    ws.Range("B12").Formula = "=HEX2DEC(""FF"")"
    ws.Range("C12").Value = "十六進位 FF = 十進位 255"

    ' OCT2DEC - 八進位轉十進位
    ws.Range("A13").Value = "OCT2DEC（8 進位→10 進位）"
    ws.Range("B13").Formula = "=OCT2DEC(17)"
    ws.Range("C13").Value = "八進位 17 = 十進位 15"

    ' COMPLEX - 建立複數
    ws.Range("A14").Value = "COMPLEX（建立複數）"
    ws.Range("B14").Formula = "=COMPLEX(3,4)"
    ws.Range("C14").Value = "建立複數 3+4i"

    ' IMABS - 複數絕對值
    ws.Range("A15").Value = "IMABS（複數模）"
    ws.Range("B15").Formula = "=IMABS(""3+4i"")"
    ws.Range("C15").Value = "複數 3+4i 的模 = 5"

    ws.Columns("A:C").AutoFit

    MsgBox "工程函數公式已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立工程函數公式時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 取得或建立工作表
Private Function GetOrCreateEngWs(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateEngWs = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateEngWs Is Nothing Then
        Set GetOrCreateEngWs = ThisWorkbook.Worksheets.Add
        GetOrCreateEngWs.Name = sheetName
    End If
End Function
