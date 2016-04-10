Attribute VB_Name = "TableSheetCreation"
Option Explicit
Option Private Module

'====================================================================================================
'
' テーブルシート作成モジュール
'
'====================================================================================================

'====================================================================================================
' テーブルシート作成の実行
'====================================================================================================
Public Sub Execute()
On Error GoTo Finally
    Dim tableSettings As Object
    Dim tableDefinitions As Collection

    ' 初期化
    If Not Initialize() Then
        Exit Sub
    End If

    ' テーブル設定リストの取得
    Set tableSettings = DataEntrySheet.GetTableSettings(False)
    If tableSettings.Count = 0 Then
        MsgBoxEx.Warning "作成対象のテーブルがありません。", "テーブル一覧にテーブル物理名を入力してください。"
        Exit Sub
    End If

    ' テーブル設定リストを元に、テーブル定義リストを取得
    Set tableDefinitions = GetTableDefinitions(tableSettings)
    ' テーブル定義リストを元に、テーブルシートを作成する
    Call TableSheet.CreateTableSheet(tableDefinitions)
    ' テーブル設定にハイパーリンクを設定する
    Call DataEntrySheet.SetHyperlink(tableSettings)

Finally:
    ' 終了化
    Call Finalize

    ' 実行結果の表示
    Call ApplicationEx.ShowExecutionResult(Err.Number = 0, "テーブルシートの作成")
End Sub


'====================================================================================================
' 初期化
'----------------------------------------------------------------------------------------------------
' OUT: True:成功、False:失敗
'====================================================================================================
Private Function Initialize() As Boolean
    ' 画面描画の抑制
    Call ApplicationEx.SuppressScreenDrawing(True)

    ' 設定モジュールの構成
    Call Setting.Setup
    If Not Setting.CheckDbSetting() Then
        Initialize = False
        Exit Function
    End If

    ' データベース接続
    Call Database.Connect
    Initialize = True
End Function


'====================================================================================================
' 終了化
'====================================================================================================
Private Sub Finalize()
    ' データベース切断
    Call Database.Disconnect

    ' 画面描画の抑制解除
    Call ApplicationEx.SuppressScreenDrawing(False)
End Sub


'====================================================================================================
' テーブル設定リストのテーブル情報をDBから取得し、返却します
'----------------------------------------------------------------------------------------------------
' IN : tableSettings テーブル設定リスト
' OUT: テーブル定義リスト
'====================================================================================================
Private Function GetTableDefinitions(tableSettings As Object) As Collection
    Dim ts As TableSetting
    Dim td As TableDefinition
    Dim list As Collection
    Dim xKey As Variant

    Set list = New Collection
    For Each xKey In tableSettings
        Set ts = tableSettings(xKey)
        Set td = New TableDefinition
        td.ColumnDefinitions = Database.GetColumnDefinitions(ts.PhysicsName)
        If td.ColumnDefinitions.Count = 0 Then
            Err.Raise 1000, , "テーブル[" & ts.PhysicsName & "]のカラム定義が取得できません。"
        End If
        td.TableName = ts.PhysicsName
        Call list.Add(td)
    Next

    Set GetTableDefinitions = list
End Function


