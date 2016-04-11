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
    Dim tableDefinitions As Object

    ' マクロ起動
    If Not ApplicationEx.StartupMacro(MacroType.Database) Then
        Exit Sub
    End If

    ' テーブル設定リストの取得
    Set tableSettings = DataEntrySheet.GetTableSettings(False)
    If tableSettings.Count = 0 Then
        MsgBoxEx.Warning "作成対象のテーブルがありません。", "テーブル一覧にテーブル物理名を入力してください。"
        Exit Sub
    End If

    ' テーブル設定リストを元に、テーブル定義リストを取得
    Set tableDefinitions = Database.GetColumnDefinitions(tableSettings)
    ' テーブル定義リストを元に、テーブルシートを作成する
    Call TableSheet.CreateTableSheet(tableDefinitions)
    ' テーブル設定にハイパーリンクを設定する
    Call DataEntrySheet.SetHyperlink(tableSettings)

Finally:
    ' マクロ停止
    Call ApplicationEx.ShutdownMacro

    ' 実行結果の表示
    Call ApplicationEx.ShowExecutionResult("テーブルシートの作成")
End Sub
