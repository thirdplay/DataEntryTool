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
Private Const cstUpdateQuery = "UPDATE ${tableName} SET (${set}) WHERE (${where})"
' Deleteクエリ
Private Const cstDeleteQuery = "DELETE FROM ${tableName} WHERE (${where})"


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
        query = Replace(query, "${columns}", GetColumnPhrase(xEntryData.ColumnDefinitions))
        query = Replace(query, "${values}", GetValuePhrase(xEntryData.ColumnDefinitions, xEntryData.RecordRange.Rows(i)))
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


'====================================================================================================
' Insert文のColumn句を取得します
'----------------------------------------------------------------------------------------------------
' IN : xColumnDefinitions カラム定義リスト
' OUT: Column句
'====================================================================================================
Private Function GetColumnPhrase(xColumnDefinitions As Collection) As String
    Dim result As String
    Dim cd As ColumnDefinition

    For Each cd In xColumnDefinitions
        result = result + cd.ColumnName + ","
    Next
    GetColumnPhrase = Left(result, Len(result) - 1)
End Function


'====================================================================================================
' Insert文のValue句を取得します
'----------------------------------------------------------------------------------------------------
' IN : record レコード
' OUT: Value句
'====================================================================================================
Private Function GetValuePhrase(xColumnDefinitions As Collection, record As Range) As String
    Dim result As String
    Dim i As Long
    Dim cd As ColumnDefinition

    For i = 1 To record.Columns.Count
        result = result + GetItemValue(record.Cells(1, i), xColumnDefinitions(i).DataType) + ","
    Next
    GetValuePhrase = Left(result, Len(result) - 1)
End Function


'====================================================================================================
' データ型に対応したデータ値の項目値を取得します
'----------------------------------------------------------------------------------------------------
' IN : dataValue データ値
'    : dataType データ型
' OUT: Value句
'====================================================================================================
Private Function GetItemValue(ByVal dataValue As String, ByVal dataType As String) As String
    Dim itemValue As String

    If dataValue = "" Then
        itemValue = "NULL"
    ' 文字列
    ElseIf mDatabaseCore.IsDataTypeString(dataType) Then
        itemValue = Replace(dataValue, "'", "''")
        itemValue = "'" & itemValue & "'"
    ' 日付
    ElseIf mDatabaseCore.IsDataTypeDate(dataType) Then
        itemValue = "TO_DATE(" & dataValue & "," & Setting.DateFormat & ")"
    ' その他
    Else
        itemValue = dataValue
    End If
    GetItemValue = itemValue
End Function
