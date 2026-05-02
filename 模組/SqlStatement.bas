Attribute VB_Name = "SqlStatement"
Option Explicit
'*************************************************************************************
'專案名稱: VBA專案
'功能描述: SQL指令彙整
'
'版權所有: 合作金庫商業銀行
'程式撰寫: Dunk
'撰寫日期：2015/7/30
'
'改版日期:
'改版備註:
'*************************************************************************************
Function 取得資料庫表格比對格式(ByVal tableName As String) As Variant
    取得資料庫表格比對格式 = "SELECT   a.TABLE_NAME as [Table ID],a.COLUMN_NAME as [Column name]," & Chr(13) & "" & Chr(10) & "case a.DATA_TYPE " & Chr(13) & "" & Chr(10) & "when 'char' then a.DATA_TYP" _
        & "E +'(' +cast(a.CHARACTER_MAXIMUM_LENGTH as varchar)+ ')'" & Chr(13) & "" & Chr(10) & "when 'varchar' then a.DATA_TYPE +'(' +cast(a.CHARACTER_MA" _
        & "XIMUM_LENGTH as varchar)+ ')'" & Chr(13) & "" & Chr(10) & "when 'decimal' then a.DATA_TYPE +'(' +cast(a.NUMERIC_PRECISION as varchar)+','+cast(" _
        & "a.NUMERIC_SCALE AS varchar)+ ')'" & Chr(13) & "" & Chr(10) & "ELSE a.DATA_TYPE" & Chr(13) & "" & Chr(10) & "end as [Data type]," & Chr(13) & "" & Chr(10) & "N'N' as [New column(Y/N)]," & Chr(13) & "" & Chr(10) & "N'N' as [Data " _
        & "type changed" & Chr(13) & "" & Chr(10) & "N: no change" & Chr(13) & "" & Chr(10) & "Data type]," & Chr(13) & "" & Chr(10) & "N'' as [Default value]," & Chr(13) & "" & Chr(10) & "N'unchanged' as [transporting type]," & Chr(13) & "" & Chr(10) & "a.ORDINAL_P" _
        & "OSITION as [New column id]," & Chr(13) & "" & Chr(10) & "a.ORDINAL_POSITION as [Original column id]" & Chr(13) & "" & Chr(10) & "FROM INFORMATION_SCHEMA.COLUMNS a " & Chr(13) & "" & Chr(10) & "inner " _
        & "join INFORMATION_SCHEMA.TABLES b on" & Chr(13) & "" & Chr(10) & "a.TABLE_NAME = b.TABLE_NAME" & Chr(13) & "" & Chr(10) & "where b.TABLE_TYPE='BASE TABLE' and a.TABLE_NAME " _
        & "='" & tableName & "'" & Chr(13) & "" & Chr(10) & "order by b.TABLE_NAME,a.ORDINAL_POSITION"
End Function

Function 取得資料表資料(ByVal tableName As String) As Variant
    取得資料表資料 = "select * from " & tableName
End Function
