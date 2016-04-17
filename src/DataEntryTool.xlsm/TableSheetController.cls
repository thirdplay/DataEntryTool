VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableSheetController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' テーブルシートのコントローラ
'
'====================================================================================================

'====================================================================================================
' 定数
'====================================================================================================
Private Const cstProcessName = "テーブルシート作成"
Private Const cstOperationContens = "テーブルシートの作成"


'====================================================================================================
' メンバ変数
'====================================================================================================
Private mDataEntrySheet As DataEntrySheet       ' データ投入シート
Private mTableSheet As TableSheet               ' テーブルシート
Private mTableSheetLogic As TableSheetLogic     ' テーブルシートロジック


'====================================================================================================
' 実行
'====================================================================================================
Public Sub Execute()
On Error GoTo Finally
    Dim tableSettings As Dictionary
    Dim tableDefinitions As Dictionary

    ' 初期化
    Call Initialize

    ' テーブル設定の取得
    Set tableSettings = mDataEntrySheet.GetTableSettings(False)
    If tableSettings.Count = 0 Then
        Err.Raise ErrNumber.Warning, , "作成対象のテーブルがありません。" & vbNewLine & "テーブル一覧にテーブル物理名を入力してください。"
    End If

    ' プログレスバー表示
    Call StatusBar.ShowProgress(cstProcessName, 100, 10)

    ' テーブル定義の取得
    Set tableDefinitions = mTableSheetLogic.GetColumnDefinitions(tableSettings)
    Call StatusBar.IncreaseProgress(50)

    ' テーブルシートの作成
    Call mTableSheet.CreateTableSheet(tableSettings, tableDefinitions)
    Call StatusBar.IncreaseProgress(40)

    ' ハイパーリンクの設定
    Call mDataEntrySheet.SetHyperlink(tableSettings)
    Call StatusBar.IncreaseProgress(10)

Finally:
    ' 終了化
    Call Finalize
    ' 実行結果の表示
    Call ApplicationEx.ShowExecutionResult(cstOperationContens)
    ' ステータスバー非表示
    Call StatusBar.Hide
End Sub


'====================================================================================================
' 初期化
'====================================================================================================
Private Sub Initialize()
    ' マクロ起動
    Call ApplicationEx.StartupMacro(SettingType.Database)

    ' インスタンス生成
    Set mDataEntrySheet = New DataEntrySheet
    Set mTableSheet = New TableSheet
    Set mTableSheetLogic = New TableSheetLogic
End Sub


'====================================================================================================
' 終了化
'====================================================================================================
Private Sub Finalize()
    ' インスタンス解放
    Set mDataEntrySheet = Nothing
    Set mTableSheet = Nothing
    Set mTableSheetLogic = Nothing

    ' マクロ停止
    Call ApplicationEx.ShutdownMacro
End Sub
