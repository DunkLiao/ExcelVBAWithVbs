Option Explicit
'*************************************************************************************
'模組名稱: WebFormulaExample
'功能說明: 示範 Excel 網路函數的用法，包含 ENCODEURL、FILTERXML、HYPERLINK 等
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/27
'
'*************************************************************************************
Sub WebFormulaExample()
    Dim ws As Worksheet
    Dim row As Long

    On Error GoTo ErrHandler

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("WebFormula").Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True

    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "WebFormula"
    row = 1

    ' 標題列
    ws.Cells(row, 1).Value = "函數名稱"
    ws.Cells(row, 2).Value = "公式範例"
    ws.Cells(row, 3).Value = "說明"
    ws.Rows(row).Font.Bold = True
    ws.Rows(row).Interior.Color = RGB(68, 114, 196)
    ws.Rows(row).Font.Color = RGB(255, 255, 255)
    row = row + 1

    ' ENCODEURL：將文字轉換為 URL 安全格式
    ws.Cells(row, 1).Value = "ENCODEURL"
    ws.Cells(row, 2).Formula = "=ENCODEURL(""Excel VBA 教學"")"
    ws.Cells(row, 3).Value = "將含空格或中文的文字編碼為 URL 格式"
    row = row + 1

    ' ENCODEURL 組合 Google 搜尋 URL
    ws.Cells(row, 1).Value = "ENCODEURL（搜尋連結）"
    ws.Cells(row, 2).Formula = "=""https://www.google.com/search?q=""&ENCODEURL(""Excel VBA 自動化"")"
    ws.Cells(row, 3).Value = "組合 Google 搜尋 URL 字串"
    row = row + 1

    ' HYPERLINK 結合 ENCODEURL
    ws.Cells(row, 1).Value = "HYPERLINK+ENCODEURL"
    ws.Cells(row, 2).Formula = "=HYPERLINK(""https://www.google.com/search?q=""&ENCODEURL(""VBA 教學""),""點此搜尋"")"
    ws.Cells(row, 3).Value = "建立帶 URL 編碼的可點擊超連結"
    row = row + 1

    ' FILTERXML 擷取 XML 第一個節點
    ws.Cells(row, 1).Value = "FILTERXML（節點1）"
    ws.Cells(row, 2).Formula = "=FILTERXML(""<r><v>Apple</v><v>Banana</v><v>Cherry</v></r>"",""/r/v[1]"")"
    ws.Cells(row, 3).Value = "從 XML 字串擷取第一個節點值"
    row = row + 1

    ' FILTERXML 擷取 XML 第二個節點
    ws.Cells(row, 1).Value = "FILTERXML（節點2）"
    ws.Cells(row, 2).Formula = "=FILTERXML(""<r><v>Apple</v><v>Banana</v><v>Cherry</v></r>"",""/r/v[2]"")"
    ws.Cells(row, 3).Value = "從 XML 字串擷取第二個節點值"
    row = row + 1

    ' FILTERXML 計算節點數
    ws.Cells(row, 1).Value = "FILTERXML（計數）"
    ws.Cells(row, 2).Formula = "=COUNTA(FILTERXML(""<r><v>A</v><v>B</v><v>C</v></r>"",""/r/v""))"
    ws.Cells(row, 3).Value = "計算 XML 中節點的總數量"
    row = row + 1

    ' WEBSERVICE 說明（僅示意，需網路連線）
    ws.Cells(row, 1).Value = "WEBSERVICE（說明）"
    ws.Cells(row, 2).Value = "=WEBSERVICE(""https://api.example.com/data"")"
    ws.Cells(row, 3).Value = "呼叫 Web API 取得回傳資料（需網路，部分版本支援）"
    row = row + 1

    ws.Columns("A:C").AutoFit

    MsgBox "網路函數範例建立完成！共 " & (row - 2) & " 個範例。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub
