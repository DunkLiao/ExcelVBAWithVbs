Attribute VB_Name = "LookMutiInCell"
Option Explicit
'*************************************************************************************
'盡嘿: ┏糷じン
'磞瓃: 琩高纗
'https://www.extendoffice.com/documents/excel/2706-excel-vlookup-return-multiple-values-in-one-cell.html
'舦┮Τ:
'祘Α级糶: Dunk
'级糶ら戳2023/2/17
'
'эら戳:
'э称爹:
'
'*************************************************************************************
'琩高逆(场才)
Function ConcatenateIf(CriteriaRange As Range, Condition As Variant, ConcatenateRange As Range, Optional Separator As String = ",") As Variant
    'Updateby Extendoffice
    Dim xResult As String
    Dim i As Long
    On Error Resume Next
    If CriteriaRange.Count <> ConcatenateRange.Count Then
        ConcatenateIf = CVErr(xlErrRef)
        Exit Function
    End If
    For i = 1 To CriteriaRange.Count
        If CriteriaRange.Cells(i).Value = Condition Then
            xResult = xResult & Separator & ConcatenateRange.Cells(i).Value
        End If
    Next i
    If xResult <> "" Then
        xResult = VBA.Mid(xResult, VBA.Len(Separator) + 1)
    End If
    ConcatenateIf = xResult
Exit Function
End Function

'琩高逆(场だ才)
Function ConcatenateIfPartial(CriteriaRange As Range, Condition As Variant, ConcatenateRange As Range, Optional Separator As String = ",") As Variant
    'Updateby Extendoffice
    Dim xResult As String
    Dim i As Long
    On Error Resume Next
    If CriteriaRange.Count <> ConcatenateRange.Count Then
        ConcatenateIfPartial = CVErr(xlErrRef)
        Exit Function
    End If
    For i = 1 To CriteriaRange.Count
        If InStrRev(CriteriaRange.Cells(i).Value, Condition) > 0 Then
            xResult = xResult & Separator & ConcatenateRange.Cells(i).Value
        End If
    Next i
    If xResult <> "" Then
        xResult = VBA.Mid(xResult, VBA.Len(Separator) + 1)
    End If
    ConcatenateIfPartial = xResult
Exit Function
End Function
