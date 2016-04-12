VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatabaseCoreOracle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' OracleのCore処理
'
'====================================================================================================
Implements IDatabaseCore


'====================================================================================================
' 定数
'====================================================================================================
' カラム定義取得クエリ
Private Const cstColumnDefinitionQuery = "SELECT utc.table_name,utc.column_id,utc.column_name,ucc.comments,utc.data_type,CASE utc.data_type WHEN 'CHAR' THEN TO_CHAR(utc.data_length) WHEN 'VARCHAR2' THEN TO_CHAR(utc.data_length) WHEN 'NUMBER' THEN utc.data_precision || ',' || utc.data_scale ELSE ' ' END AS data_length,DECODE(utc.nullable,'N','1','0') AS is_required,DECODE(pkc.table_name,NULL,'0','1') AS is_primary_key FROM user_tab_columns utc LEFT JOIN user_col_comments ucc ON(ucc.table_name = utc.table_name AND ucc.column_name = utc.column_name) LEFT JOIN (SELECT cc.table_name,cc.column_name FROM user_cons_columns cc INNER JOIN user_constraints cons ON(cons.constraint_name = cc.constraint_name AND cons.constraint_type = 'P')) pkc ON (pkc.table_name = utc.table_name AND pkc.column_name = utc.column_name) WHERE utc.table_name IN(${tableNameIn}) ORDER BY table_name, column_id"
' テーブル名取得クエリ
Private Const cstTableNameQuery = "SELECT table_name from user_tables"
Private Const cstDataEntryQueryPrefix = "INSERT ALL "
Private Const cstDataEntryQuerySuffix = " SELECT * FROM dual"
Private Const cstDataEntryQuery = "INTO ${tableName} (${columns}) VALUES (${values})"
' 区切り文字
Private Const cstDelimiter = ","


'====================================================================================================
' データベースの接続文字列を取得します
'----------------------------------------------------------------------------------------------------
' OUT: 接続文字列
'====================================================================================================
Public Function IDatabaseCore_GetConnectStr() As String
    IDatabaseCore_GetConnectStr = "Driver={Microsoft ODBC for Oracle};" _
        & "Server=" & Setting.ServerName & ";" _
        & "UID=" & Setting.UserId & ";" _
        & "PWD=" & Setting.Password & ";"
End Function


'====================================================================================================
' テーブル名取得クエリを取得します
'----------------------------------------------------------------------------------------------------
' OUT: テーブル名取得クエリ
'====================================================================================================
Public Function IDatabaseCore_GetTableNameQuery() As String
    IDatabaseCore_GetTableNameQuery = cstTableNameQuery
End Function


'====================================================================================================
' カラム定義取得クエリを取得します
'----------------------------------------------------------------------------------------------------
' IN : tableSettings テーブル設定の連装配列
' OUT: カラム定義取得クエリ
'====================================================================================================
Public Function IDatabaseCore_GetColumnDefinitionQuery(tableSettings As Object) As String
    Dim query As String
    Dim xTableName As Variant
    Dim tableNameInStr As String

    For Each xTableName In tableSettings.Keys
        tableNameInStr = tableNameInStr & "'" & xTableName & "'" & cstDelimiter
    Next
    tableNameInStr = Left(tableNameInStr, Len(tableNameInStr) - Len(cstDelimiter))
    query = Replace(cstColumnDefinitionQuery, "${tableNameIn}", tableNameInStr)

    IDatabaseCore_GetColumnDefinitionQuery = query
End Function


'====================================================================================================
' データ投入クエリの接頭辞を取得します
'----------------------------------------------------------------------------------------------------
' OUT: データ投入クエリの接頭辞
'====================================================================================================
Public Function IDatabaseCore_GetDataEntryQueryPrefix() As String
    IDatabaseCore_GetDataEntryQueryPrefix = cstDataEntryQueryPrefix
End Function


'====================================================================================================
' データ投入クエリの接尾辞を取得します
'----------------------------------------------------------------------------------------------------
' OUT: データ投入クエリの接尾辞
'====================================================================================================
Public Function IDatabaseCore_GetDataEntryQuerySuffix() As String
    IDatabaseCore_GetDataEntryQuerySuffix = cstDataEntryQuerySuffix
End Function


'====================================================================================================
' データ投入クエリを取得します
'----------------------------------------------------------------------------------------------------
' OUT: データ投入クエリ
'====================================================================================================
Public Function IDatabaseCore_GetDataEntryQuery() As String
    IDatabaseCore_GetDataEntryQuery = cstDataEntryQuery
End Function


'====================================================================================================
' データ種別が文字列かどうか判定する
'----------------------------------------------------------------------------------------------------
' IN : xDataType データ型
' OUT: True:文字列、False:文字列以外
'====================================================================================================
Public Function IDatabaseCore_IsDataTypeString(ByVal xDataType As String) As Boolean
    If InStr(UCase(xDataType), "CHAR") Then
        IDatabaseCore_IsDataTypeString = True
        Exit Function
    End If
    IDatabaseCore_IsDataTypeString = False
End Function


'====================================================================================================
' データ種別が日付かどうか判定する
'----------------------------------------------------------------------------------------------------
' IN : xDataType データ型
' OUT: True:日付、False:日付以外
'====================================================================================================
Public Function IDatabaseCore_IsDataTypeDate(ByVal xDataType As String) As Boolean
    xDataType = UCase(xDataType)
    If InStr(xDataType, "TIMESTAMP") Or xDataType = "DATE" Then
        IDatabaseCore_IsDataTypeDate = True
        Exit Function
    End If
    IDatabaseCore_IsDataTypeDate = False
End Function


'====================================================================================================
' データ型がタイムスタンプかどうか判定する
'----------------------------------------------------------------------------------------------------
' IN : xDataType データ型
' OUT: True:タイムスタンプ、False:タイムスタンプ以外
'====================================================================================================
Public Function IDatabaseCore_IsDataTypeTimestamp(ByVal xDataType As String) As Boolean
    xDataType = UCase(xDataType)
    If xDataType = "TIMESTAMP" Then
        IDatabaseCore_IsDataTypeTimestamp = True
        Exit Function
    End If
    IDatabaseCore_IsDataTypeTimestamp = False
End Function
