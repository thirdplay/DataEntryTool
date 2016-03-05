VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PostgreSqlDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' PostgreSqlDao
'
'====================================================================================================
Implements IDataEntryDao


'====================================================================================================
' 定数
'====================================================================================================
' カラム情報取得クエリ
Private Const cstColumnDefinitionGetQuery As String = _
    "SELECT isc.ordinal_position - 1 AS column_id, isc.column_name, COALESCE(pd.description, '') AS comments, isc.data_type AS data_type, CASE isc.data_type WHEN 'character' THEN TO_CHAR(isc.character_maximum_length, 'FM99999999') WHEN 'character varying' THEN TO_CHAR(isc.character_maximum_length, 'FM99999999') WHEN 'numeric' THEN TO_CHAR(isc.numeric_precision, 'FM99999999') || ',' || TO_CHAR(isc.numeric_scale, 'FM99999999') ELSE '' END AS data_length, CASE isc.is_nullable WHEN 'NO' THEN '1' ELSE '0' END AS is_nullable, CASE WHEN pkc.column_name IS NOT NULL THEN '1' ELSE '0' END AS is_primary_key FROM information_schema.columns isc INNER JOIN pg_stat_all_tables psat ON( psat.schemaname = isc.table_schema AND psat.relname = isc.table_name ) INNER JOIN pg_attribute pa ON( pa.attrelid = psat.relid AND pa.attname = isc.column_name ) LEFT JOIN pg_description pd ON( pd.objoid = psat.relid AND pd.objsubid != 0 and pd.objsubid = pa.attnum ) LEFT JOIN ( SELECT tc.table_catalog, tc.table_schema, tc.table_name, " & _
    "ccu.column_name FROM information_schema.table_constraints tc INNER JOIN information_schema.constraint_column_usage ccu ON( ccu.table_catalog = tc.table_catalog AND ccu.table_schema = tc.table_schema AND ccu.table_name = tc.table_name AND ccu.constraint_name = tc.constraint_name  ) WHERE tc.constraint_type = 'PRIMARY KEY' ) pkc ON ( pkc.table_catalog = isc.table_catalog AND pkc.table_schema = isc.table_schema AND pkc.table_name = isc.table_name AND pkc.column_name = isc.column_name ) WHERE isc.table_name='${tableName}' ORDER BY isc.ordinal_position"


'====================================================================================================
' メンバ変数
'====================================================================================================
Dim mConnect As Object      ' 接続オブジェクト


'====================================================================================================
' コンストラクタ
'====================================================================================================
Private Sub Class_Initialize()
End Sub


'====================================================================================================
' デストラクタ
'====================================================================================================
Private Sub Class_Terminate()
    Call IDataEntryDao_Disconnect
End Sub


'====================================================================================================
' DB接続
'----------------------------------------------------------------------------------------------------
' IN : setting データベース情報
'====================================================================================================
Public Sub IDataEntryDao_Connect(setting As Setting)
On Error GoTo ErrHandler
    Set mConnect = CreateObject("ADODB.Connection")

    mConnect.Open "Driver={PostgreSQL Unicode};" _
        & "SERVER=" & setting.ServerName & ";" _
        & "DATABASE=" & setting.DatabaseName & ";" _
        & "UID=" & setting.UserId & ";" _
        & "PWD=" & setting.Password & ";" _
        & "PORT=" & setting.Port & ";"
    Exit Sub
ErrHandler:
    Set mConnect = Nothing
    Err.Raise Err.Number, Err.Source, "DB接続失敗" & vbNewLine & Err.Description, Err.HelpFile, Err.HelpContext
    Exit Sub
End Sub


'====================================================================================================
' DB切断
'====================================================================================================
Public Sub IDataEntryDao_Disconnect()
    If Not mConnect Is Nothing Then
        mConnect.Close
        Set mConnect = Nothing
    End If
End Sub


'====================================================================================================
' トランザクション開始
'====================================================================================================
Public Sub IDataEntryDao_BeginTrans()
    Debug.Print "PostgreSql:BeginTrans"
End Sub


'====================================================================================================
' コミット
'====================================================================================================
Public Sub IDataEntryDao_CommitTrans()
    Debug.Print "PostgreSql:CommitTrans"
End Sub


'====================================================================================================
' ロールバック
'====================================================================================================
Public Sub IDataEntryDao_RollbackTrans()
    Debug.Print "PostgreSql:RollbackTrans"
End Sub


'====================================================================================================
' クエリ実行
'----------------------------------------------------------------------------------------------------
' IN : query クエリ文字列
'====================================================================================================
Public Sub IDataEntryDao_Execute(query)
    Debug.Print "PostgreSqlDao.Execute"
End Sub


'====================================================================================================
' 指定されたテーブルのカラム情報リストを取得します
'----------------------------------------------------------------------------------------------------
' IN : tableName テーブル名
' OUT: カラム情報リスト
'====================================================================================================
Public Function IDataEntryDao_GetColumnDefinitions(tableName As String) As Collection
    Dim recordSet As Object
    Dim query As String
    Dim cd As ColumnDefinition
    Dim list As Collection
    Set list = New Collection

    query = Replace(cstColumnDefinitionGetQuery, "${tableName}", tableName)
    Set recordSet = mConnect.Execute(query)
    Do Until recordSet.EOF
        Set cd = New ColumnDefinition
        With ci
            .ColumnId = recordSet("column_id")
            .ColumnName = recordSet("column_name")
            .Comments = recordSet("comments")
            .DataType = recordSet("data_type")
            .DataLength = recordSet("data_length")
            .IsNullable = recordSet("is_nullable")
            .IsPrimaryKey = recordSet("is_primary_key")
        End With
        Call list.Add(cd)
        recordSet.MoveNext
    Loop
    Set recordSet = Nothing

    Set IDataEntryDao_GetColumnDefinitions = list
End Function


