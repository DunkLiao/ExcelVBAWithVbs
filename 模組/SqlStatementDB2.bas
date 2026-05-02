Attribute VB_Name = "SqlStatementDB2"
Option Explicit
'*************************************************************************************
'專案名稱: 雪梨分行集中
'功能描述: SQL指令彙整
'
'版權所有: 合作金庫商業銀行
'程式撰寫: Dunk
'撰寫日期：2015/10/26
'
'改版日期:
'改版備註:
'*************************************************************************************

Function 取得資料表資料(ByVal tableName As String, ByVal schemaName As String) As Variant
    取得資料表資料 = "select * from " & schemaName & "." & tableName
End Function

Function 取得資料表結構(ByVal tableName As String, ByVal schemaName As String) As Variant
    取得資料表結構 = "   SELECT TABNAME AS TABLEID, COLNAME AS COLUMNNAME,     CASE TYPENAME " & _
      "WHEN 'CHARACTER' THEN LOWER('CHAR') || '(' || LENGTH || ')' " & _
      "WHEN 'VARCHAR' THEN LOWER(TYPENAME) || '(' || LENGTH || ')' " & _
      "WHEN 'DECIMAL' THEN LOWER(TYPENAME) || '(' || LENGTH || ',' || SCALE || ')'" & _
      "Else LOWER (TypeName) " & _
    "END AS DATATYPEDESC, COLNO + 1 AS ORIGINALCOLUMNID " & _
  "FROM SYSCAT.Columns " & _
  "WHERE TABSCHEMA = '" & schemaName & "' AND TABNAME = '" & tableName & "' " & _
  "ORDER BY TABLEID, COLNO;"
End Function
