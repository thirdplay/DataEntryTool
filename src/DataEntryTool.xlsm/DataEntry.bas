Attribute VB_Name = "DataEntry"
Option Explicit
Option Private Module

'====================================================================================================
'
' データ投入モジュール
'
'====================================================================================================

'====================================================================================================
' データ登録
'====================================================================================================
Public Sub RegisterData()
    Call Execute(EntryType.Register)
End Sub


'====================================================================================================
' データ更新
'====================================================================================================
Public Sub UpdateData()
    Call Execute(EntryType.Update)
End Sub


'====================================================================================================
' データ削除
'====================================================================================================
Public Sub DeleteData()
    Call Execute(EntryType.Delete)
End Sub


'====================================================================================================
' データ投入の実行
'----------------------------------------------------------------------------------------------------
' IN : xEntryType 投入種類
'====================================================================================================
Private Sub Execute(xEntryType As EntryType)
On Error GoTo Finally
    Dim tableSettings As Object
    Dim ts As TableSetting
    Dim ed As EntryData
    Dim xKey As Variant
    Dim procCnt As Long
    Dim operationDic As Object
    Set operationDic = GetOperationDic()

    ' マクロ起動
    Call ApplicationEx.StartupMacro(MacroType.DataEntry)

    ' 処理件数のクリア
    Call DataEntrySheet.ClearProcessingCount

    ' 対象テーブル設定の取得
    Set tableSettings = DataEntrySheet.GetTableSettings(True)
    If tableSettings.Count = 0 Then
        Err.Raise ErrNumber.Warning, , "データ投入対象のデータがありません。" & vbNewLine & vbNewLine & _
            "下記手順を実施してデータ投入対象のデータを設定してください。" & vbNewLine & _
            "  ・テーブル一覧のデータ投入対象列に空文字以外の値を設定する。" & vbNewLine & _
            "  ・データ投入対象のテーブルシートにデータを入力する。"
    End If

    ' 対象テーブル設定を全て処理
    For Each xKey In tableSettings
        Set ts = tableSettings(xKey)

        ' 対象テーブルのテーブルデータの取得
        Set ed = TableSheet.GetEntryData(ts.PhysicsName)

        ' データ投入実行
        procCnt = DataEntryModel.ExecuteDataEntry(xEntryType, ed)

        ' 処理件数の書き込み
        Call DataEntrySheet.WriteProcessingCount(ts, procCnt)
    Next
Finally:
    ' マクロ停止
    Call ApplicationEx.ShutdownMacro

    ' 実行結果の表示
    Call ApplicationEx.ShowExecutionResult("データ" & operationDic(xEntryType))
End Sub


'====================================================================================================
' 投入種類に対応した投入文字列を格納した連装配列を取得します
'----------------------------------------------------------------------------------------------------
' OUT: 連装配列
'====================================================================================================
Private Function GetOperationDic()
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    Call dic.Add(EntryType.Register, "登録")
    Call dic.Add(EntryType.Update, "更新")
    Call dic.Add(EntryType.Delete, "削除")
    Set GetOperationDic = dic
End Function
