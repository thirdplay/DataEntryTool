VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatabaseCorePostgreSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' PostgreSqlのCore処理
'
'====================================================================================================
Implements IDatabaseCore


'====================================================================================================
' 定数
'====================================================================================================
' カラム定義取得クエリ
Private Const cstColumnDefinitionQuery As String = _
    "SELECT isc.table_name, isc.ordinal_position AS column_id, isc.column_name, COALESCE(pd.description, '') AS comments, isc.data_type AS data_type, CASE isc.data_type WHEN 'character' THEN TO_CHAR(isc.character_maximum_length, 'FM99999999') WHEN 'character varying' THEN TO_CHAR(isc.character_maximum_length, 'FM99999999') WHEN 'numeric' THEN TO_CHAR(isc.numeric_precision, 'FM99999999') || ',' || TO_CHAR(isc.numeric_scale, 'FM99999999') ELSE '' END AS data_length, CASE isc.is_nullable WHEN 'NO' THEN '1' ELSE '0' END AS is_required, CASE WHEN pkc.column_name IS NOT NULL THEN '1' ELSE '0' END AS is_primary_key FROM information_schema.columns isc INNER JOIN pg_stat_all_tables psat ON( psat.schemaname = isc.table_schema AND psat.relname = isc.table_name ) INNER JOIN pg_attribute pa ON( pa.attrelid = psat.relid AND pa.attname = isc.column_name ) LEFT JOIN pg_description pd ON( pd.objoid = psat.relid AND pd.objsubid != 0 and pd.objsubid = pa.attnum ) LEFT JOIN ( SELECT tc.table_catalog, " & _
    "tc.table_schema, tc.table_name, ccu.column_name FROM information_schema.table_constraints tc INNER JOIN information_schema.constraint_column_usage ccu ON( ccu.table_catalog = tc.table_catalog AND ccu.table_schema = tc.table_schema AND ccu.table_name = tc.table_name AND ccu.constraint_name = tc.constraint_name  ) WHERE tc.constraint_type = 'PRIMARY KEY' ) pkc ON ( pkc.table_catalog = isc.table_catalog AND pkc.table_schema = isc.table_schema AND pkc.table_name = isc.table_name AND pkc.column_name = isc.column_name ) WHERE isc.table_name='${tableName}'"
' テーブル名取得クエリ
Private Const cstTableNameQuery = "SELECT relname AS table_name FROM pg_stat_user_tables"
' データ投入クエリ
Private Const cstDataEntryQuery = "INSERT INTO ${tableName} (${columns}) VALUES (${values}) RETURNING *;"


'====================================================================================================
' データベースの接続文字列を取得します
'----------------------------------------------------------------------------------------------------
' OUT: 接続文字列
'====================================================================================================
Public Function IDatabaseCore_GetConnectStr() As String
    IDatabaseCore_GetConnectStr = "Driver={PostgreSQL Unicode};" _
        & "SERVER=" & Setting.ServerName & ";" _
        & "DATABASE=" & Setting.DatabaseName & ";" _
        & "UID=" & Setting.UserId & ";" _
        & "PWD=" & Setting.Password & ";" _
        & "PORT=" & Setting.Port & ";"
End Function


'====================================================================================================
' カラム定義取得クエリを取得します
'----------------------------------------------------------------------------------------------------
' OUT: カラム定義取得クエリ
'====================================================================================================
Public Function IDatabaseCore_GetColumnDefinitionQuery() As String
    IDatabaseCore_GetColumnDefinitionQuery = cstColumnDefinitionQuery
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
' データ投入クエリの接頭辞を取得します
'----------------------------------------------------------------------------------------------------
' OUT: データ投入クエリの接頭辞
'====================================================================================================
Public Function IDatabaseCore_GetDataEntryQueryPrefix() As String
    IDatabaseCore_GetDataEntryQueryPrefix = ""
End Function


'====================================================================================================
' データ投入クエリの接尾辞を取得します
'----------------------------------------------------------------------------------------------------
' OUT: データ投入クエリの接尾辞
'====================================================================================================
Public Function IDatabaseCore_GetDataEntryQuerySuffix() As String
    IDatabaseCore_GetDataEntryQuerySuffix = ""
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
    If InStr(UCase(xDataType), "CHAR") Or UCase(xDataType) = "TEXT" Then
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
    If xDataType = "DATE" Then
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
    If InStr(xDataType, "TIMESTAMP") Then
        IDatabaseCore_IsDataTypeTimestamp = True
        Exit Function
    End If
    IDatabaseCore_IsDataTypeTimestamp = False
End Function
