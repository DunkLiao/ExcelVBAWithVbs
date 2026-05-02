Attribute VB_Name = "ShapUtil"
Option Explicit
'*************************************************************************************
'專案名稱: 風管系統
'功能描述: 報表標示紅框工具
'
'版權所有: Dunk
'程式撰寫: Dunk
'撰寫日期：2024/7/11
'
'改版日期:
'改版備註:
'
'*************************************************************************************
'清除所有紅框
Function ClearAllOval()
    Dim Object As Object
    
    For Each Object In ActiveSheet.Shapes
        If InStrRev(Object.Name, "Oval") > 0 Then
            Object.Delete
        End If
    Next
    
    Set Object = Nothing
End Function

'設定紅框(實線)
Function MarkRedMarkWithSelection()
        ActiveSheet.Shapes.AddShape(msoShapeOval, Selection.Left, Selection.Top, Selection.Width, Selection.Height).Select
        Selection.ShapeRange.Fill.Visible = msoFalse
        
        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 0, 0)
            .Transparency = 0
        End With
End Function

'設定紅框(虛線)
Function MarkRedMarkWithSelectionLineDash()
        ActiveSheet.Shapes.AddShape(msoShapeOval, Selection.Left, Selection.Top, Selection.Width, Selection.Height).Select
        Selection.ShapeRange.Fill.Visible = msoFalse
        
        With Selection.ShapeRange.Line
            .Visible = msoTrue
            .DashStyle = msoLineDash
            .ForeColor.RGB = RGB(255, 0, 0)
            .Transparency = 0
        End With
End Function
