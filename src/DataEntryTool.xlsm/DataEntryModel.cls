VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataEntryModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' データ投入モデル
'
'====================================================================================================

'====================================================================================================
' メンバ変数
'====================================================================================================
Private mDatabaseModel As DatabaseModel   ' データベースモデル


'====================================================================================================
' コンストラクタ
'====================================================================================================
Private Sub Class_Initialize()
End Sub


'====================================================================================================
' デストラクタ
'====================================================================================================
Private Sub Class_Terminate()
    Set mDatabaseModel = Nothing
End Sub


'====================================================================================================
' モデルを構成します
'====================================================================================================
Public Sub Setup()
    Set mDatabaseModel = New DatabaseModel
    Call mDatabaseModel.Connect
End Sub


'====================================================================================================
' テーブル設定リストのテーブル情報をDBから取得し、返却します
'----------------------------------------------------------------------------------------------------
' IN : tableSettings テーブル設定リスト
' OUT: テーブル定義リスト
'====================================================================================================
Public Function GetTableDefinitions(tableSettings As Collection) As Collection
    Dim ts As TableSetting
    Dim td As TableDefinition
    Dim list As Collection

    Set list = New Collection
    For Each ts In tableSettings
        Set td = New TableDefinition
        td.ColumnDefinitions = mDatabaseModel.GetColumnDefinitions(ts.PhysicsName)
        If td.ColumnDefinitions.Count = 0 Then
            Err.Raise 100, , "テーブル[" & ts.PhysicsName & "]のカラム定義が取得できません。"
        End If
        td.TableName = ts.PhysicsName
        Call list.Add(td)
    Next

    Set GetTableDefinitions = list
End Function


'====================================================================================================
' データ投入を実行します
'----------------------------------------------------------------------------------------------------
' IN : xEntryData 投入データ
' OUT: 処理件数
'====================================================================================================
Public Function ExecuteDataEntry(xEntryData As EntryData) As Long
On Error GoTo ErrHandler
    Dim i As Long
    Dim queries As Collection
    Dim execCnt As Long
    execCnt = 0

    ' トランザクション開始
    Call mDatabaseModel.BeginTrans

    ' クエリ生成
    Set queries = mDatabaseModel.GenerateQueries(xEntryData)

    ' クエリ実行
    For i = 1 To queries.Count
        Debug.Print queries(i)
        'execCnt = execCnt + mDatabaseModel.ExecuteQuery(query)
    Next

    ' コミット
    Call mDatabaseModel.CommitTrans

    ExecuteDataEntry = execCnt
    Exit Function
ErrHandler:
    ' ロールバック
    Call mDatabaseModel.RollbackTrans
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    Exit Function
End Function

