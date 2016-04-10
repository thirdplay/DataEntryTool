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

    ' 画面描画の抑制
    Call ApplicationEx.SuppressScreenDrawing(True)
    ' 設定モジュールの構成
    Call Setting.Setup
    If Not Setting.CheckDbSetting() Then
        Exit Sub
    End If
    ' データベース接続
    Call Database.Connect

    ' テーブル設定リストの取得
    Set tableSettings = DataEntrySheet.GetTableSettings(False)
    If tableSettings.Count = 0 Then
        MsgBoxEx.Warning "作成対象のテーブルがありません。", "テーブル一覧にテーブル物理名を入力してください。"
        Exit Sub
    End If

    ' テーブル設定リストを元に、テーブル定義リストを取得
    Set tableDefinitions = TableSheetCreationModel.GetTableDefinitions(tableSettings)
    ' テーブル定義リストを元に、テーブルシートを作成する
    Call TableSheet.CreateTableSheet(tableDefinitions)
    ' テーブル設定にハイパーリンクを設定する
    Call DataEntrySheet.SetHyperlink(tableSettings)

Finally:
    ' データベース切断
    Call Database.Disconnect
    ' 画面描画の抑制解除
    Call ApplicationEx.SuppressScreenDrawing(False)

    ' 実行結果の表示
    If Err.Number <> 0 Then
        MsgBoxEx.Error "テーブルシートの作成に失敗しました", Err.Description
    Else
        MsgBox "テーブルシートの作成が完了しました"
    End If
End Sub


