Attribute VB_Name = "DataEntry"
Option Explicit
Option Private Module

'====================================================================================================
'
' データ投入モジュール
'
'====================================================================================================

'====================================================================================================
' データを登録します
'====================================================================================================
Public Sub RegisterData()
    Call Execute(EntryType.Register)
End Sub


'====================================================================================================
' データを更新します
'====================================================================================================
Public Sub UpdateData()
    Call Execute(EntryType.Update)
End Sub


'====================================================================================================
' データを削除します
'====================================================================================================
Public Sub RemoveData()
    Call Execute(EntryType.Remove)
End Sub


'====================================================================================================
' データ投入の実行
'----------------------------------------------------------------------------------------------------
' IN : xEntryType 投入種類
'====================================================================================================
Private Sub Execute(xEntryType As EntryType)
On Error GoTo Finally
    Dim operationDic As Object
    Dim tableSettings As Collection
    Dim ts As TableSetting
    Dim ed As EntryData
    Dim procCount As Long

    ' 画面描画の抑制
    Call ApplicationEx.SuppressScreenDrawing(True)
    ' 設定モジュールの構成
    Call Setting.Setup
    If Not Setting.CheckDataEntrySetting() Then
        Exit Sub
    End If

    ' 処理件数のクリア
    Call ProcessCountModel.ClearProcessingCount

    ' 対象テーブル設定の取得
    Set tableSettings = TableSettingModel.GetTableSettings(True)
    If tableSettings.Count = 0 Then
        MsgBoxEx.Warning "データ投入対象のデータがありません。" & vbNewLine & vbNewLine & _
            "下記手順を実施してデータ投入対象のデータを設定してください。" & vbNewLine & _
            "  ・テーブル一覧のデータ投入対象列に空文字以外の値を設定する。" & vbNewLine & _
            "  ・データ投入対象のテーブルシートにデータを入力する。"
            
        Exit Sub
    End If

    ' 対象テーブル設定を全て処理
    For Each ts In tableSettings
        ' 対象テーブルのテーブルデータの取得
        Set ed = EntryDataModel.GetEntryData(xEntryType, ts.PhysicsName)

        ' データ投入実行
        procCount = DataEntryModel.ExecuteDataEntry(ed)

        ' 処理件数の書き込み
        Call ProcessCountModel.WriteProcessingCount(ts, procCount)
    Next
Finally:
    ' 画面描画の抑制解除
    Call ApplicationEx.SuppressScreenDrawing(False)

    ' 実行結果の表示
    Set operationDic = GetOperationDic
    If Err.Number <> 0 Then
        MsgBoxEx.Error "データ" & operationDic(xEntryType) & "に失敗しました" & vbNewLine & Err.Description
    Else
        MsgBox "データ" & operationDic(xEntryType) & "が完了しました"
    End If
End Sub


'====================================================================================================
' 投入種類に対応した投入文字列を格納する辞書を取得します
'----------------------------------------------------------------------------------------------------
' OUT: 投入辞書
'====================================================================================================
Private Function GetOperationDic()
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    Call dic.Add(EntryType.Register, "登録")
    Call dic.Add(EntryType.Update, "更新")
    Call dic.Add(EntryType.Remove, "削除")
    Set GetOperationDic = dic
End Function

