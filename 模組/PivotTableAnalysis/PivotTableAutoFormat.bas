Option Explicit
Attribute VB_Name = "PivotTableAutoFormat"
'*************************************************************************************
'模組名稱: PivotTableAutoFormat
'功能說明: 自動套用格式設定至活頁簿中所有樞紐分析表（樣式、數字格式、字型等）
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************

Sub ApplyPivotTableAutoFormat()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim df As PivotField
    Dim foundCount As Integer

    On Error GoTo ErrHandler

    foundCount = 0

    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            foundCount = foundCount + 1

            pt.TableStyle2 = "PivotStyleMedium9"
            pt.ShowTableStyleRowStripes = True
            pt.ShowTableStyleColumnStripes = False

            For Each df In pt.DataFields
                df.NumberFormat = "#,##0"
            Next df

            With pt.TableRange1.Font
                .Name = "微軟正黑體"
                .Size = 10
            End With

            pt.RowAxisLayout xlCompactRow
            pt.HasAutoFormat = True

        Next pt
    Next ws

    If foundCount = 0 Then
        MsgBox "目前活頁簿中沒有找到任何樞紐分析表。", vbExclamation, "提示"
    Else
        MsgBox "已自動套用格式至 " & foundCount & " 個樞紐分析表。", vbInformation, "完成"
    End If

    Exit Sub

ErrHandler:
    MsgBox "錯誤：" & Err.Description, vbCritical, "自動格式化樞紐分析表失敗"
End Sub