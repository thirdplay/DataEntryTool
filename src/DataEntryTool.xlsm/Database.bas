Attribute VB_Name = "Database"
Option Explicit
Option Private Module

'====================================================================================================
'
' データベースモジュール
'
'====================================================================================================

'====================================================================================================
' 定数
'====================================================================================================
' カラム定義取得クエリの区切り文字
Private Const cstColumnDefinitionDelimiter = " UNION ALL "
' カラム定義取得クエリの接頭辞
Private Const cstColumnDefinitionPrefix = "SELECT * FROM ("
' カラム定義取得クエリの接尾辞
Private Const cstColumnDefinitionSuffix = ") tbl ORDER BY table_name, column_id"
' 区切り文字
Private Const cstDelimiter = ","


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
'====================================================================================================
Public Sub ExecuteQuery(query As String)
    Call mConnect.Execute(CommandText:=query)
End Sub


'====================================================================================================
' 指定されたテーブルのカラム定義リストを取得します
'----------------------------------------------------------------------------------------------------
' IN : tableSettings テーブル設定連想配列
' OUT: カラム定義リスト
'====================================================================================================
Public Function GetColumnDefinitions(tableSettings As Object) As Object
    Dim rs As Object
    Dim tableNames As Object
    Dim query As String
    Dim xTableName As Variant
    Dim td As TableDefinition
    Dim cd As ColumnDefinition
    Dim dic As Object

    ' テーブル名連想配列の取得
    Set rs = mConnect.Execute(mDatabaseCore.GetTableNameQuery)
    Set tableNames = CreateObject("Scripting.Dictionary")
    Do Until rs.EOF
        Call tableNames.Add(rs("table_name").Value, rs("table_name").Value)
        rs.MoveNext
    Loop

    ' クエリを作りながらテーブル名の存在チェックをする
    query = cstColumnDefinitionPrefix
    For Each xTableName In tableSettings.Keys
        query = query & Replace(mDatabaseCore.GetColumnDefinitionQuery, "${tableName}", xTableName) & cstColumnDefinitionDelimiter
        If Not tableNames.Exists(xTableName) Then
            Err.Raise 1000, , "テーブル[" & xTableName & "]のカラム定義が取得できません。"
        End If
    Next
    query = Left(query, Len(query) - Len(cstColumnDefinitionDelimiter)) & cstColumnDefinitionSuffix
    Debug.Print query

    ' カラム定義取得クエリの実行
    Set rs = mConnect.Execute(query)

    ' テーブル定義の連装配列を作成する
    Set dic = CreateObject("Scripting.Dictionary")
    Do Until rs.EOF
        ' テーブル定義生成
        If Not dic.Exists(rs("table_name").Value) Then
            Set td = New TableDefinition
            td.TableName = rs("table_name").Value
            td.ColumnDefinitions = New Collection
            Call dic.Add(td.TableName, td)
        End If

        ' カラム定義生成
        Set cd = New ColumnDefinition
        With cd
            .ColumnId = rs("column_id").Value
            .ColumnName = rs("column_name").Value
            .Comments = rs("comments").Value
            .DataType = rs("data_type").Value
            .DataLength = rs("data_length").Value
            .IsRequired = rs("is_required").Value
            .IsPrimaryKey = rs("is_primary_key").Value
        End With
        Call dic(rs("table_name").Value).ColumnDefinitions.Add(cd)
        rs.MoveNext
    Loop
    Set GetColumnDefinitions = dic
End Function


'====================================================================================================
' データ投入クエリを生成します
'----------------------------------------------------------------------------------------------------
' IN : xEntryData 投入データ
' OUT: クエリ文字列
'====================================================================================================
Public Function MakeDataEntryQuery(xEntryData As EntryData) As String
    Dim rr As Range
    Dim query As String

    query = mDatabaseCore.GetDataEntryQueryPrefix
    For Each rr In xEntryData.RecordRange.Rows
        query = query & mDatabaseCore.GetDataEntryQuery
        query = Replace(query, "${tableName}", xEntryData.TableName)
        query = Replace(query, "${columns}", GetColumnPhrase(xEntryData.ColumnDefinitions))
        query = Replace(query, "${values}", GetValuePhrase(xEntryData.ColumnDefinitions, rr))
    Next
    query = query & mDatabaseCore.GetDataEntryQuerySuffix
    MakeDataEntryQuery = query
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
        result = result & cd.ColumnName & cstDelimiter
    Next
    GetColumnPhrase = Left(result, Len(result) - Len(cstDelimiter))
End Function


'====================================================================================================
' Insert文のValue句を取得します
'----------------------------------------------------------------------------------------------------
' IN : cd カラム定義リスト
'    : record レコード
' OUT: Value句
'====================================================================================================
Private Function GetValuePhrase(cd As Collection, record As Range) As String
    Dim result As String
    Dim i As Long

    For i = 1 To record.Columns.Count
        result = result & GetItemValue(record.Cells(1, i), cd(i).DataType) & cstDelimiter
    Next
    GetValuePhrase = Left(result, Len(result) - Len(cstDelimiter))
End Function


'====================================================================================================
' データ型に対応したデータ値の項目値を取得します
'----------------------------------------------------------------------------------------------------
' IN : dataValue データ値
'    : xDataType データ型
' OUT: Value句
'====================================================================================================
Private Function GetItemValue(ByVal dataValue As String, ByVal xDataType As String) As String
    Dim itemValue As String

    itemValue = dataValue
    If itemValue = "" Then
        itemValue = "NULL"
    ' 文字列
    ElseIf mDatabaseCore.IsDataTypeString(xDataType) Then
        itemValue = Replace(itemValue, "'", "''")                                   ' 単一引用符エスケープ
        itemValue = Replace(itemValue, vbLf, "'" & Setting.LinefeedCode & "'")      ' 改行コード変換
        itemValue = "'" & itemValue & "'"                                           ' 単一引用符付与
    ' 日付
    ElseIf mDatabaseCore.IsDataTypeDate(xDataType) Then
        itemValue = "TO_DATE('" & itemValue & "','" & Setting.DateFormat & "')"
    ' タイムスタンプ
    ElseIf mDatabaseCore.IsDataTypeTimestamp(xDataType) Then
        itemValue = "TO_TIMESTAMP('" & itemValue & "','" & Setting.TimestampFormat & "')"
    End If
    GetItemValue = itemValue
End Function
