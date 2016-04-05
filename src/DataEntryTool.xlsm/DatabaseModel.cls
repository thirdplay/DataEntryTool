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
' DB接続
'----------------------------------------------------------------------------------------------------
' IN : xSetting 設定
'====================================================================================================
Public Sub Connect(xSetting As Setting)
On Error GoTo ErrHandler
    If mConnect Is Nothing Then
        ' データベースコアの生成
        If xSetting.DatabaseType = cstDatabaseTypeOracle Then
            Set mDatabaseCore = New DatabaseCoreOracle
        ElseIf xSetting.DatabaseType = cstDatabaseTypePostgreSQL Then
            Set mDatabaseCore = New DatabaseCorePostgreSql
        End If

        ' データベース接続
        Set mConnect = CreateObject("ADODB.Connection")
        mConnect.Open mDatabaseCore.GetConnectStr(xSetting)
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
' クエリ実行
'----------------------------------------------------------------------------------------------------
' IN : query クエリ文字列
'====================================================================================================
Public Sub Execute(query)
    Debug.Print "Execute"
End Sub


'====================================================================================================
' 指定されたテーブルのカラム定義リストを取得します
'----------------------------------------------------------------------------------------------------
' IN : tableName テーブル名
' OUT: カラム定義リスト
'====================================================================================================
Public Function GetColumnDefinitions(TableName As String) As Collection
    Dim recordSet As Object
    Dim query As String
    Dim cd As ColumnDefinition
    Dim list As Collection
    Set list = New Collection

    query = Replace(mDatabaseCore.GetColumnDefinitionQuery, "${tableName}", TableName)
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


