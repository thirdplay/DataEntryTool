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
Private Const cstProcessName = "テーブルシートの作成"


'====================================================================================================
' メンバ変数
'====================================================================================================
Private mTableSheetLogic As TableSheetLogic     ' テーブルシートロジック
Private mDataEntrySheet As DataEntrySheet       ' データ投入シート
Private mTableSheet As TableSheet               ' テーブルシート


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
    Call ProgressBarForm.ResetProgress
    Call ProgressBarForm.SetMainLabelText(cstProcessName)
    Call ProgressBarForm.AddMessageToDetailsBox(cstProcessName & "開始")
    Call ProgressBarForm.Show

    ' テーブル定義の取得
    Call ProgressBarForm.AddMessageToDetailsBox("テーブル定義の取得")
    Set tableDefinitions = mTableSheetLogic.GetColumnDefinitions(tableSettings)
    Call ProgressBarForm.IncreaseProgressByPercent(50)

    ' テーブルシートの作成
    Call ProgressBarForm.AddMessageToDetailsBox("テーブルシートの作成")
    Call mTableSheet.CreateTableSheet(tableSettings, tableDefinitions)
    Call ProgressBarForm.IncreaseProgressByPercent(40)

    ' ハイパーリンクの設定
    Call ProgressBarForm.AddMessageToDetailsBox("ハイパーリンクの設定")
    Call mDataEntrySheet.SetHyperlink(tableSettings)
    Call ProgressBarForm.IncreaseProgressByPercent(10)
    Call ProgressBarForm.AddMessageToDetailsBox("正常終了")

Finally:
    ' 終了化
    Call Finalize
    ' 実行結果の表示
    Call ApplicationEx.ShowExecutionResult(cstProcessName)

    ' プログレスバー非表示
    If Err.Number = 0 Then
        Unload ProgressBarForm
    End If
End Sub


'====================================================================================================
' 初期化
'====================================================================================================
Private Sub Initialize()
    ' マクロ起動
    Call ApplicationEx.StartupMacro(SettingType.Database)

    ' インスタンス生成
    Set mTableSheetLogic = New TableSheetLogic
    Set mDataEntrySheet = New DataEntrySheet
    Set mTableSheet = New TableSheet
End Sub


'====================================================================================================
' 終了化
'====================================================================================================
Private Sub Finalize()
    ' インスタンス解放
    Set mTableSheetLogic = Nothing
    Set mDataEntrySheet = Nothing
    Set mTableSheet = Nothing

    ' マクロ停止
    Call ApplicationEx.ShutdownMacro
End Sub
