Option Explicit
Attribute VB_Name = "SplitSheetByUserInput"
'*************************************************************************************

'模組名稱: SplitSheetByUserInput

'功能說明: 依使用者輸入的欄位值，將工作表資料切割為多個獨立工作表

'

'版權所有: Dunk

'程式設計: Dunk

'撰寫日期: 2026/5/27

'

'*************************************************************************************




Sub SplitSheetByUserInput()

    Dim ws As Worksheet

    Dim wsDst As Worksheet

    Dim lastRow As Long

    Dim lastCol As Long

    Dim i As Long

    Dim colIndex As Integer

    Dim keyVal As String

    Dim colInput As String

    Dim sheetName As String

    Dim dictSheets As Object



    Set ws = ActiveSheet

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column



    If lastRow < 2 Then

        MsgBox "工作表資料不足，至少需要標題列與一筆資料。", vbExclamation, "提示"

        Exit Sub

    End If



    colInput = InputBox("請輸入要依據切割的欄號（例如：2 代表 B 欄）：", "切割工作表", "2")

    If colInput = "" Then Exit Sub



    On Error Resume Next

    colIndex = CInt(colInput)

    On Error GoTo 0



    If colIndex < 1 Or colIndex > lastCol Then

        MsgBox "欄號無效，請輸入 1 到 " & lastCol & " 之間的數字。", vbExclamation, "錯誤"

        Exit Sub

    End If



    Set dictSheets = CreateObject("Scripting.Dictionary")



    Application.ScreenUpdating = False



    ' 掃描所有資料列，依指定欄位值分組

    For i = 2 To lastRow

        keyVal = Trim(CStr(ws.Cells(i, colIndex).Value))

        If keyVal = "" Then keyVal = "(空白)"



        ' 建立新工作表（若尚未建立）

        If Not dictSheets.Exists(keyVal) Then

            sheetName = Left(keyVal, 31)

            sheetName = Replace(sheetName, "/", "-")

            sheetName = Replace(sheetName, "\\", "-")

            sheetName = Replace(sheetName, "*", "-")

            sheetName = Replace(sheetName, "?", "-")

            sheetName = Replace(sheetName, "[", "(")

            sheetName = Replace(sheetName, "]", ")")



            Dim nameIdx As Integer

            Dim finalName As String

            finalName = sheetName

            nameIdx = 1

            Dim existWs As Worksheet

            Dim nameExists As Boolean

            Do

                nameExists = False

                For Each existWs In ThisWorkbook.Worksheets

                    If existWs.Name = finalName Then

                        nameExists = True

                        Exit For

                    End If

                Next existWs

                If nameExists Then

                    finalName = Left(sheetName, 28) & "_" & nameIdx

                    nameIdx = nameIdx + 1

                End If

            Loop While nameExists



            Set wsDst = ThisWorkbook.Worksheets.Add( _

                After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))

            wsDst.Name = finalName

            ws.Rows(1).Copy wsDst.Rows(1)

            dictSheets.Add keyVal, wsDst

        End If



        Dim targetWs As Worksheet

        Set targetWs = dictSheets(keyVal)

        Dim targetRow As Long

        targetRow = targetWs.Cells(targetWs.Rows.Count, 1).End(xlUp).Row + 1

        ws.Rows(i).Copy targetWs.Rows(targetRow)

    Next i



    Dim wsItem As Variant

    For Each wsItem In dictSheets.Keys

        dictSheets(wsItem).Columns.AutoFit

    Next wsItem



    Application.ScreenUpdating = True

    MsgBox "切割完成！共建立 " & dictSheets.Count & " 個工作表。", vbInformation, "完成"

End Sub

