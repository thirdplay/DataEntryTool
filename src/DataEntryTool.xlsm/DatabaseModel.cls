VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatabaseModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' データベースモデル
'
'====================================================================================================

'====================================================================================================
' 定数
'====================================================================================================
' Insertクエリ
Private Const cstInsertQuery = "INSERT INTO ${tableName} (${columns}) VALUES (${values})"
' Updateクエリ
Private Const cstInsertQuery = "UPDATE ${tableName} SET (${set}) WHERE (${where})"
' Deleteクエリ
Private Const cstInsertQuery = "DELETE FROM ${tableName} WHERE (${where})"


'====================================================================================================
' メンバ変数
'====================================================================================================
Dim mDatabaseCore As IDatabaseCore      ' データベースコアインターフェース
Dim mConnect As Object                  ' 接続オブジェクト


'====================================================================================================
' コンストラクタ
'====================================================================================================
Private Sub Class_Initialize()
End Sub


'====================================================================================================
' デストラクタ
'====================================================================================================
Private Sub Class_Terminate()
    Call Disconnect
End Sub


'====================================================================================================
' モデルを構成する
'====================================================================================================
Public Sub Setup
    Call Connect
End Sub


'====================================================================================================
' DB接続
'====================================================================================================
Public Sub Connect()
On Error GoTo ErrHandler
    If mConnect Is Nothing Then
        ' データベースコアの生成
        If Setting.DatabaseType = cstDatabaseTypeOracle Then
            Set mDatabaseCore = New DatabaseCoreOracle
        ElseIf Setting.DatabaseType = cstDatabaseTypePostgreSQL Then
            Set mDatabaseCore = New DatabaseCorePostgreSql
        End If

        ' データベース接続
        Set mConnect = CreateObject("ADODB.Connection")
        mConnect.Open mDatabaseCore.GetConnectStr
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
Public Sub Disconnect()
    If Not mConnect Is Nothing Then
        mConnect.Close
        Set mConnect = Nothing
        Set mDatabaseCore = Nothing
    End If
End Sub


'====================================================================================================
' トランザクション開始
'====================================================================================================
Public Sub BeginTrans()
    mConnect.BeginTrans
End Sub


'====================================================================================================
' コミット
'====================================================================================================
Public Sub CommitTrans()
    mConnect.CommitTrans
End Sub


'====================================================================================================
' ロールバック
'====================================================================================================
Public Sub RollbackTrans()
    mConnect.RollbackTrans
End Sub


'====================================================================================================
' クエリを実行します
'----------------------------------------------------------------------------------------------------
' IN : query クエリ文字列
' OUT: 処理件数
'====================================================================================================
Public Function ExecuteQuery(query As String) As Long
    ExecuteQuery = 0
End Function


'====================================================================================================
' 指定されたテーブルのカラム定義リストを取得します
'----------------------------------------------------------------------------------------------------
' IN : tableName テーブル名
' OUT: カラム定義リスト
'====================================================================================================
Public Function GetColumnDefinitions(tableName As String) As Collection
    Dim recordSet As Object
    Dim query As String
    Dim cd As ColumnDefinition
    Dim list As Collection
    Set list = New Collection

    query = Replace(mDatabaseCore.GetColumnDefinitionQuery, "${tableName}", tableName)
    Set recordSet = mConnect.Execute(query)
    Do Until recordSet.EOF
        Set cd = New ColumnDefinition
        With cd
            .ColumnId = recordSet("column_id")
            .ColumnName = recordSet("column_name")
            .Comments = recordSet("comments")
            .DataType = recordSet("data_type")
            .DataLength = recordSet("data_length")
            .IsRequired = recordSet("is_required")
            .IsPrimaryKey = recordSet("is_primary_key")
        End With
        Call list.Add(cd)
        recordSet.MoveNext
    Loop
    Set GetColumnDefinitions = list
End Function


'====================================================================================================
' クエリを生成します
'----------------------------------------------------------------------------------------------------
' IN : xEntryData 投入データ
' OUT: クエリリスト
'====================================================================================================
Public Function GenerateQueries(xEntryData As EntryData) As Collection
    Select Case xEntryData.EntryType
    Case EntryType.Register
        Set GenerateQueries = GenerateInsertQueries(xEntryData)
    Case EntryType.Update
        Set GenerateQueries = GenerateUpdateQueries(xEntryData)
    Case EntryType.Remove
        Set GenerateQueries = GenerateDeleteQueries(xEntryData)
    End Select
End Function


'====================================================================================================
' Insertクエリリストを生成します
'----------------------------------------------------------------------------------------------------
' IN : xEntryData 投入データ
' OUT: クエリリスト
'====================================================================================================
Public Function GenerateInsertQueries(xEntryData As EntryData) As Collection
    Dim i As Long
    Dim query As String
    Dim queries As Collection

    Set queries = New Collection
    For i = 1 To xEntryData.RecordRange.Rows.Count
        query = cstInsertQuery
        query = Replace(query, "${tableName}", xEntryData.TableName)
        query = Replace(query, "${columns}", "") 'GetColumnsPhrase(xEntryData.ColumnsDefinitions)
        query = Replace(query, "${values}", "") 'GetValuesPhrase(xEntryData.ColumnsDefinitions, xEntryData.RecordRange.Rows(i))
        Call queries.Add(query)
    Next
    Set GenerateInsertQueries = queries
End Function


'====================================================================================================
' Updateクエリリストを生成します
'----------------------------------------------------------------------------------------------------
' IN : xEntryData 投入データ
' OUT: クエリリスト
'====================================================================================================
Public Function GenerateUpdateQueries(xEntryData As EntryData) As Collection
    Dim i As Long
    Dim query As String
    Dim queries As Collection

    Set queries = New Collection
    For i = 1 To xEntryData.RecordRange.Rows.Count
        query = cstUpdateQuery
        query = Replace(query, "${tableName}", xEntryData.TableName)
        query = Replace(query, "${set}", "") 'GetSetPhrase(xEntryData.ColumnsDefinitions, xEntryData.RecordRange.Rows(i))
        query = Replace(query, "${where}", "") 'GetWherePhrase(xEntryData.ColumnsDefinitions, xEntryData.RecordRange.Rows(i))
        Call queries.Add(query)
    Next
    Set GenerateUpdateQueries = queries
End Function


'====================================================================================================
' Deleteクエリリストを生成します
'----------------------------------------------------------------------------------------------------
' IN : xEntryData 投入データ
' OUT: クエリリスト
'====================================================================================================
Public Function GenerateDeleteQueries(xEntryData As EntryData) As Collection
    Dim i As Long
    Dim query As String
    Dim queries As Collection

    Set queries = New Collection
    For i = 1 To xEntryData.RecordRange.Rows.Count
        query = cstDeleteQuery
        query = Replace(query, "${tableName}", xEntryData.TableName)
        query = Replace(query, "${where}", "") 'GetWherePhrase(xEntryData.ColumnsDefinitions, xEntryData.RecordRange.Rows(i))
        Call queries.Add(query)
    Next
    Set GenerateDeleteQueries = Nothing
End Function

