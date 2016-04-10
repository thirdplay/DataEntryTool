Attribute VB_Name = "Database"
Option Explicit
Option Private Module

'====================================================================================================
'
' データベースモジュール
'
'====================================================================================================

'====================================================================================================
' メンバ変数
'====================================================================================================
Private mDatabaseCore As IDatabaseCore      ' データベースコアインターフェース
Private mConnect As Object                  ' 接続オブジェクト


'====================================================================================================
' データベース接続
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
        mConnect.CursorLocation = adUseClient
    End If
    Exit Sub
ErrHandler:
    Set mConnect = Nothing
    Err.Raise Err.Number, Err.Source, "データベース接続失敗" & vbNewLine & Err.Description, Err.HelpFile, Err.HelpContext
    Exit Sub
End Sub


'====================================================================================================
' データベース切断
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
    Call mConnect.BeginTrans
End Sub


'====================================================================================================
' コミット
'====================================================================================================
Public Sub CommitTrans()
    Call mConnect.CommitTrans
End Sub


'====================================================================================================
' ロールバック
'====================================================================================================
Public Sub RollbackTrans()
    Call mConnect.RollbackTrans
End Sub


'====================================================================================================
' クエリを実行します
'----------------------------------------------------------------------------------------------------
' IN : query クエリ文字列
' OUT: 処理件数
'====================================================================================================
Public Function ExecuteQuery(query As String) As Long
    Dim procCnt As Long
    Dim rs As Object

    Set rs = mConnect.Execute(CommandText:=query)
    Do Until rs Is Nothing
        procCnt = procCnt + rs.RecordCount
        Set rs = rs.NextRecordset
    Loop
    ExecuteQuery = procCnt
End Function


'====================================================================================================
' 指定されたテーブルのカラム定義リストを取得します
'----------------------------------------------------------------------------------------------------
' IN : xTableName テーブル名
' OUT: カラム定義リスト
'====================================================================================================
Public Function GetColumnDefinitions(xTableName As String) As Collection
    Dim recordSet As Object
    Dim query As String
    Dim cd As ColumnDefinition
    Dim list As Collection
    Set list = New Collection

    query = Replace(mDatabaseCore.GetColumnDefinitionQuery, "${tableName}", xTableName)
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
' データ型が文字列かどうか判定する
'----------------------------------------------------------------------------------------------------
' IN : xDataType データ型
' OUT: True:文字列、False:文字列以外
'====================================================================================================
Public Function IsDataTypeString(ByVal xDataType As String) As Boolean
    IsDataTypeString = mDatabaseCore.IsDataTypeString(xDataType)
End Function


'====================================================================================================
' データ型が日付かどうか判定する
'----------------------------------------------------------------------------------------------------
' IN : xDataType データ型
' OUT: True:日付、False:日付以外
'====================================================================================================
Public Function IsDataTypeDate(ByVal xDataType As String) As Boolean
    IsDataTypeDate = mDatabaseCore.IsDataTypeDate(xDataType)
End Function


'====================================================================================================
' データ型がタイムスタンプかどうか判定する
'----------------------------------------------------------------------------------------------------
' IN : xDataType データ型
' OUT: True:タイムスタンプ、False:タイムスタンプ以外
'====================================================================================================
Public Function IsDataTypeTimestamp(ByVal xDataType As String) As Boolean
    IsDataTypeTimestamp = mDatabaseCore.IsDataTypeTimestamp(xDataType)
End Function


'====================================================================================================
' クエリの接尾辞を取得します
'----------------------------------------------------------------------------------------------------
' OUT: 接尾辞
'====================================================================================================
Public Function GetQuerySuffix() As String
    GetQuerySuffix = mDatabaseCore.GetQuerySuffix()
End Function


