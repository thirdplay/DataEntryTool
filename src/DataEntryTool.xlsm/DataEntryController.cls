VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataEntryController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' データ投入のコントローラ
'
'====================================================================================================

'====================================================================================================
' メンバ変数
'====================================================================================================
Private mBaseController As BaseController       ' 基本コントローラー
Private mDataEntrySheet As DataEntrySheet       ' データ投入シート
Private mTableSheet As TableSheet               ' テーブルシート
Private mDataEntryLogic As DataEntryLogic       ' データ投入ロジック
Private mProcessName As String                  ' 処理名


'====================================================================================================
' 実行
'----------------------------------------------------------------------------------------------------
' IN : xEntryType 投入種別
'====================================================================================================
Public Sub Execute(xEntryType As EntryType)
On Error GoTo Finally
    Dim tableSettings As Dictionary
    Dim xKey As Variant
    Dim ts As TableSetting
    Dim ed As EntryData
    Dim procCnt As Long

    ' 初期化
    Call Initialize(xEntryType)

    ' 処理件数のクリア
    Call mDataEntrySheet.ClearProcessingCount

    ' 対象テーブル設定の取得
    Set tableSettings = mDataEntrySheet.GetTableSettings(True)
    If tableSettings.Count = 0 Then
        Err.Raise ErrNumber.Warning, , "データ投入対象のデータがありません。" & vbNewLine & vbNewLine & _
            "下記手順を実施してデータ投入対象のデータを設定してください。" & vbNewLine & _
            "  ・テーブル一覧のデータ投入対象列に空文字以外の値を設定する。" & vbNewLine & _
            "  ・データ投入対象のテーブルシートにデータを入力する。"
    End If

    ' 対象テーブル設定を全て処理
    For Each xKey In tableSettings
        ' 対象テーブルのテーブルデータの取得
        Set ts = tableSettings(xKey)
        Set ed = mTableSheet.GetEntryData(ts.PhysicsName)

        ' データ投入実行
        procCnt = mDataEntryLogic.ExecuteDataEntry(xEntryType, ed)

        ' 処理件数の書き込み
        Call mDataEntrySheet.WriteProcessingCount(ts, procCnt)
    Next

Finally:
    ' 終了化
    Call Finalize
End Sub


'====================================================================================================
' 初期化
'----------------------------------------------------------------------------------------------------
' IN : xEntryType 投入種別
'====================================================================================================
Private Sub Initialize(xEntryType As EntryType)
    ' 起動
    Set mBaseController = New BaseController
    Call mBaseController.Startup(SettingType.DataEntry)

    ' インスタンス生成
    Set mDataEntrySheet = New DataEntrySheet
    Set mTableSheet = New TableSheet
    Set mDataEntryLogic = New DataEntryLogic
    mProcessName = GetProcessName(xEntryType)
End Sub


'====================================================================================================
' 終了化
'====================================================================================================
Private Sub Finalize()
    ' インスタンス解放
    Set mDataEntrySheet = Nothing
    Set mTableSheet = Nothing
    Set mDataEntryLogic = Nothing

    ' 停止
    Call mBaseController.Shutdown

    ' 実行結果の表示
    Call mBaseController.ShowExecutionResult(mProcessName)
    Set mBaseController = Nothing
End Sub


'====================================================================================================
' 処理名を取得します
'----------------------------------------------------------------------------------------------------
' IN : xEntryType 投入種別
' OUT: 処理名
'====================================================================================================
Private Function GetProcessName(xEntryType As EntryType)
    Select Case xEntryType
        Case EntryType.Register
            GetProcessName = "データ登録"
        Case EntryType.Update
            GetProcessName = "データ更新"
        Case EntryType.Delete
            GetProcessName = "データ削除"
    End Select
End Function
