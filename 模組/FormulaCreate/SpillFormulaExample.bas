Option Explicit
Attribute VB_Name = "SpillFormulaExample"
'*************************************************************************************

'模組名稱: SpillFormulaExample

'功能說明: 示範 Excel 動態陣列溢位公式（UNIQUE、SORT、FILTER、SEQUENCE）的建立方式

'

'版權所有: Dunk

'程式設計: Dunk

'撰寫日期: 2026/5/27

'

'*************************************************************************************




Sub CreateSpillFormulaExample()

    Dim ws As Worksheet

    Dim i As Integer



    ' 建立新工作表

    Set ws = ThisWorkbook.Worksheets.Add

    ws.Name = "SpillFormulas"



    ' 填入標題

    ws.Range("A1").Value = "溢位公式範例"

    ws.Range("A1").Font.Bold = True

    ws.Range("A1").Font.Size = 13



    ' 來源資料

    ws.Range("A3").Value = "來源資料"

    ws.Range("A3").Font.Bold = True



    Dim srcData As Variant

    srcData = Array("蘋果", "香蕉", "蘋果", "芒果", "香蕉", "草莓", "芒果", "蘋果")

    For i = 0 To UBound(srcData)

        ws.Cells(4 + i, 1).Value = srcData(i)

    Next i



    Dim numData As Variant

    numData = Array(5, 3, 8, 2, 7, 4, 6, 1)

    For i = 0 To UBound(numData)

        ws.Cells(4 + i, 2).Value = numData(i)

    Next i



    ws.Range("A3:B3").EntireColumn.AutoFit



    ' UNIQUE 公式

    ws.Range("D3").Value = "UNIQUE 去重"

    ws.Range("D3").Font.Bold = True

    ws.Range("D4").Formula2 = "=UNIQUE(A4:A11)"



    ' SORT 公式

    ws.Range("F3").Value = "SORT 排序"

    ws.Range("F3").Font.Bold = True

    ws.Range("F4").Formula2 = "=SORT(A4:A11)"



    ' SEQUENCE 公式

    ws.Range("H3").Value = "SEQUENCE 序列"

    ws.Range("H3").Font.Bold = True

    ws.Range("H4").Formula2 = "=SEQUENCE(8,1,1,1)"



    ' FILTER 公式（篩選數值 >= 5）

    ws.Range("J3").Value = "FILTER (數值>=5)"

    ws.Range("J3").Font.Bold = True

    ws.Range("J4").Formula2 = "=FILTER(A4:B11,B4:B11>=5,""無資料"")"



    ' 自動調整欄寬

    ws.Columns("A:K").AutoFit



    MsgBox "溢位公式範例已建立完成，工作表：" & ws.Name, vbInformation, "完成"

End Sub

