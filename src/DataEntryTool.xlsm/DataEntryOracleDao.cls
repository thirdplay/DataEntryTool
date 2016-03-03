VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataEntryOracleDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' OracleDao
'
'====================================================================================================
Implements IDataEntryDao


'====================================================================================================
' 定数
'====================================================================================================
' カラム情報取得クエリ
Private Const cstColumnInfoGetQuery As String = "SELECT utc.column_id, utc.column_name, ucc.comments, utc.data_type, CASE utc.data_type WHEN 'CHAR' THEN TO_CHAR(utc.data_length) WHEN 'VARCHAR2' THEN TO_CHAR(utc.data_length) WHEN 'NUMBER' THEN utc.data_precision || ',' || utc.data_scale ELSE ' ' END AS data_length, DECODE(utc.nullable,'N','0','1') AS is_nullable, DECODE(pkc.table_name, NULL, '0', '1') AS is_primary_key FROM user_tab_columns utc LEFT JOIN user_col_comments ucc ON(ucc.table_name = utc.table_name AND ucc.column_name = utc.column_name) LEFT JOIN (SELECT cc.table_name, cc.column_name FROM user_cons_columns cc INNER JOIN user_constraints cons ON(cons.constraint_name = cc.constraint_name AND cons.constraint_type = 'P')) pkc ON (pkc.table_name = utc.table_name AND pkc.column_name = utc.column_name) WHERE utc.table_name='${tableName}' ORDER BY utc.column_id"


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
' IN : setting 設定
'====================================================================================================
Public Sub IDataEntryDao_Connect(setting As Setting)
On Error GoTo ErrHandler
    If mConnect Is Nothing Then
        Set mConnect = CreateObject("ADODB.Connection")

        mConnect.Open "Driver={Microsoft ODBC for Oracle};" _
            & "Server=" & setting.ServerName & ";" _
            & "UID=" & setting.UserId & ";" _
            & "PWD=" & setting.Password & ";"
        Exit Sub
    End If
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
    MsgBox "Oracle:BeginTrans"
End Sub


'====================================================================================================
' コミット
'====================================================================================================
Public Sub IDataEntryDao_CommitTrans()
    MsgBox "Oracle:CommitTrans"
End Sub


'====================================================================================================
' ロールバック
'====================================================================================================
Public Sub IDataEntryDao_RollbackTrans()
    MsgBox "Oracle:RollbackTrans"
End Sub


'====================================================================================================
' 指定されたテーブルのカラム情報リストを取得します。
'----------------------------------------------------------------------------------------------------
' IN : tableName テーブル名
' OUT: カラム情報リスト
'====================================================================================================
Public Function IDataEntryDao_GetColumnInfoList(tableName As String) As Collection
    Dim recordSet As Object
    Dim query As String
    Dim ci As ColumnInfo
    Dim list As Collection
    Set list = New Collection

    query = Replace(cstColumnInfoGetQuery, "${tableName}", tableName)
    Set recordSet = mConnect.Execute(query)
    Do Until recordSet.EOF
        Set ci = New ColumnInfo
        With ci
            .ColumnId = recordSet("column_id")
            .ColumnName = recordSet("column_name")
            .Comments = recordSet("comments")
            .DataType = recordSet("data_type")
            .DataLength = recordSet("data_length")
            .IsNullable = recordSet("is_nullable")
            .IsPrimaryKey = recordSet("is_primary_key")
        End With
        Call list.Add(ci)
        recordSet.MoveNext
    Loop
    Set IDataEntryDao_GetColumnInfoList = list
End Function


'====================================================================================================
' クエリ実行
'----------------------------------------------------------------------------------------------------
' IN : query クエリ文字列
'====================================================================================================
'Public Sub Execute(query)
'End Sub
