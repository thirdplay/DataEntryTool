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
' OUT: 処理件数
'====================================================================================================
Public Function ExecuteQuery(query As String)
    Dim procCnt As Long
    Call mConnect.Execute(CommandText:=query, RecordsAffected:=procCnt)
    ExecuteQuery = procCnt
End Function


'====================================================================================================
' テーブル名を取得します
'----------------------------------------------------------------------------------------------------
' OUT: テーブル名
'====================================================================================================
Public Function GetTableName()
    Set GetTableName = mConnect.Execute(mDatabaseCore.GetTableNameQuery)
End Function


'====================================================================================================
' カラム定義を取得します
'----------------------------------------------------------------------------------------------------
' IN : tableSettings テーブル設定の連装配列
' OUT: カラム定義
'====================================================================================================
Public Function GetColumnDefinition(tableSettings As Object)
    Set GetColumnDefinition = mConnect.Execute(mDatabaseCore.GetColumnDefinitionQuery(tableSettings))
End Function


'====================================================================================================
' データ型に対応したデータ値の項目値を取得します
'----------------------------------------------------------------------------------------------------
' IN : dataValue データ値
'    : xDataType データ型
' OUT: Value句
'====================================================================================================
Public Function GetItemValue(ByVal dataValue As String, ByVal xDataType As String) As String
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


'Private Sub GetRecordsetData()
'   iExecutionCount = iExecutionCount + 1
'   If rst.State <> adStateClosed Then
'      rst.Close
'   End If
'   rst.Open _
'      "Select * From Pubs..Publishers, Pubs..Titles, Pubs..Authors", _
'      con, adOpenKeyset, adLockOptimistic, adAsyncExecute
'End Sub
'
'Private Sub con_ExecuteComplete(ByVal RecordsAffected As Long, _
'    ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, _
'    ByVal pCommand As ADODB.Command, _
'    ByVal pRecordset As ADODB.Recordset, _
'    ByVal pConnection As ADODB.Connection)
'
'   ' When the ADO recordset has been populated with data, begin opening
'   ' the next ADO recordset.
'   GetRecordsetData
'End Sub
