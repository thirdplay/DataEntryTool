VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataEntryDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' データ投入Dao
'
'====================================================================================================

'====================================================================================================
' メンバ変数
'====================================================================================================
Private mDatabaseCore As IDatabaseCore                  ' データベースコアインターフェース
Private WithEvents mConnection As ADODB.Connection      ' 接続オブジェクト
Attribute mConnection.VB_VarHelpID = -1
Private mAsyncRecordset As ADODB.Recordset              ' 非同期処理のレコードセット
Private mAsyncRecordsAffected As Long                   ' 非同期処理の更新件数
Private mIsAsyncProc As Boolean                         ' 非同期処理フラグ


'====================================================================================================
' コンストラクタ
'====================================================================================================
Public Sub Class_Initialize()
End Sub


'====================================================================================================
' デストラクタ
'====================================================================================================
Public Sub Class_Terminate()
    Call Me.Disconnect
End Sub


'====================================================================================================
' データベース接続
'====================================================================================================
Public Sub Connect()
On Error GoTo ErrHandler
    If mConnection Is Nothing Then
        ' データベースコアの生成
        If Setting.DatabaseType = cstDatabaseTypeOracle Then
            Set mDatabaseCore = New DatabaseCoreOracle
        ElseIf Setting.DatabaseType = cstDatabaseTypePostgreSQL Then
            Set mDatabaseCore = New DatabaseCorePostgreSql
        End If

        ' データベース接続
        Set mConnection = New ADODB.Connection
        Call mConnection.Open(mDatabaseCore.GetConnectStr(Setting.ServerName, Setting.UserId, Setting.Password, Setting.Port, Setting.DatabaseName))
        mConnection.CursorLocation = adUseClient
    End If
    Exit Sub
ErrHandler:
    Set mConnection = Nothing
    Err.Raise Err.Number, Err.Source, "データベース接続失敗" & vbNewLine & Err.Description, Err.HelpFile, Err.HelpContext
    Exit Sub
End Sub


'====================================================================================================
' データベース切断
'====================================================================================================
Public Sub Disconnect()
    If Not mConnection Is Nothing Then
        mConnection.Close
        Set mConnection = Nothing
        Set mDatabaseCore = Nothing
    End If
End Sub


'====================================================================================================
' トランザクション開始
'====================================================================================================
Public Sub BeginTrans()
    Call mConnection.BeginTrans
End Sub


'====================================================================================================
' コミット
'====================================================================================================
Public Sub CommitTrans()
    Call mConnection.CommitTrans
End Sub


'====================================================================================================
' ロールバック
'====================================================================================================
Public Sub RollbackTrans()
    Call mConnection.RollbackTrans
End Sub


'====================================================================================================
' 更新クエリを実行します
'----------------------------------------------------------------------------------------------------
' IN : query クエリ文字列
' OUT: 処理件数
'====================================================================================================
Public Function ExecuteUpdate(query As String) As Long
    Dim procCnt As Long
    Call mConnection.Execute(CommandText:=query, RecordsAffected:=procCnt)
    ExecuteUpdate = procCnt
End Function


'====================================================================================================
' テーブル名を取得します
'----------------------------------------------------------------------------------------------------
' OUT: テーブル名
'====================================================================================================
Public Function GetTableName() As ADODB.Recordset
    Set GetTableName = mConnection.Execute(mDatabaseCore.GetTableNameQuery)
End Function


'====================================================================================================
' カラム定義を取得します
'----------------------------------------------------------------------------------------------------
' IN : tableSettings テーブル設定の連装配列
' OUT: カラム定義
'====================================================================================================
Public Function GetColumnDefinition(tableSettings As Object) As ADODB.Recordset
    Call ExecuteAsync(mDatabaseCore.GetColumnDefinitionQuery(tableSettings))
    Call WaitAsync
    Set GetColumnDefinition = mAsyncRecordset
End Function


'====================================================================================================
' データ型に対応したデータ値の項目値を取得します
'----------------------------------------------------------------------------------------------------
' IN : dataValue データ値
'    : xDataType データ型
' OUT: 項目値
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


'====================================================================================================
' クエリ実行(非同期)
'----------------------------------------------------------------------------------------------------
' IN : query クエリ文字列
'====================================================================================================
Private Sub ExecuteAsync(query As String)
    mIsAsyncProc = True
    Call mConnection.Execute(CommandText:=query, Options:=adAsyncExecute)
End Sub


'====================================================================================================
' 非同期処理が終了するまで待つ
'====================================================================================================
Private Sub WaitAsync()
    Do While mIsAsyncProc
        Call ApplicationEx.CheckEvents
    Loop
End Sub


'====================================================================================================
' コネクションの実行完了イベント
'----------------------------------------------------------------------------------------------------
' IN : RecordsAffected 更新件数
'    : pError ADODBエラー
'    : adStatus イベントステータス
'    : pCommand コマンド
'    : pRecordset レコードセット
'    : pConnection コネクション
'====================================================================================================
Private Sub mConnection_ExecuteComplete( _
    ByVal RecordsAffected As Long, _
    ByVal pError As ADODB.Error, _
    ByRef adStatus As ADODB.EventStatusEnum, _
    ByVal pCommand As ADODB.Command, _
    ByVal pRecordset As ADODB.Recordset, _
    ByVal pConnection As ADODB.Connection _
)
    Set mAsyncRecordset = pRecordset
    mAsyncRecordsAffected = RecordsAffected
    mIsAsyncProc = False
End Sub
