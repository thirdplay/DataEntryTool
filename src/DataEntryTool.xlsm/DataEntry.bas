Attribute VB_Name = "DataEntry"
Option Explicit
Option Private Module

'====================================================================================================
'
' データ投入モジュール
'
'====================================================================================================

'====================================================================================================
' データ投入の実行
'----------------------------------------------------------------------------------------------------
' IN : xEntryType 投入種類
'====================================================================================================
Private Sub Execute()
On Error GoTo Finally
    Dim tableSettings As Object
    Dim ts As TableSetting
    Dim ed As EntryData
    Dim xKey As Variant
    Dim procCnt As Long

    ' マクロ起動
    If Not ApplicationEx.StartupMacro(MacroType.DataEntry) Then
        Exit Sub
    End If

    ' 処理件数のクリア
    Call TableSheet.ClearProcessingCount

    ' 対象テーブル設定の取得
    Set tableSettings = DataEntrySheet.GetTableSettings(True)
    If tableSettings.Count = 0 Then
        MsgBoxEx.Warning "データ投入対象のデータがありません。", _
            "下記手順を実施してデータ投入対象のデータを設定してください。" & vbNewLine & _
            "  ・テーブル一覧のデータ投入対象列に空文字以外の値を設定する。" & vbNewLine & _
            "  ・データ投入対象のテーブルシートにデータを入力する。"
        Exit Sub
    End If

    ' 対象テーブル設定を全て処理
    For Each xKey In tableSettings
        Set ts = tableSettings(xKey)

        ' 対象テーブルのテーブルデータの取得
        Set ed = DataEntryModel.GetEntryData(ts.PhysicsName)

        ' データ投入実行
        procCnt = DataEntryModel.ExecuteDataEntry(ed)

        ' 処理件数の書き込み
        Call TableSheet.WriteProcessingCount(ts, procCnt)
    Next
Finally:
    ' マクロ停止
    Call ApplicationEx.ShutdownMacro

    ' 実行結果の表示
    Call ApplicationEx.ShowExecutionResult("データ投入")
End Sub
