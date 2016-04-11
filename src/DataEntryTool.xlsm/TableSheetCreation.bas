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
    Set tableDefinitions = Database.GetColumnDefinitions(tableSettings)
    ' テーブル定義リストを元に、テーブルシートを作成する
    Call TableSheet.CreateTableSheet(tableDefinitions)
    ' テーブル設定にハイパーリンクを設定する
    Call DataEntrySheet.SetHyperlink(tableSettings)

Finally:
    ' 終了化
    Call Finalize

    ' 実行結果の表示
    Call ApplicationEx.ShowExecutionResult("テーブルシートの作成")
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
