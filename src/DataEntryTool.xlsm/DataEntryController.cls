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
' テーブルシートのコントローラ
'
'====================================================================================================

'====================================================================================================
' 定数
'====================================================================================================
Private Const cstProcessName = "データ"
Private Const cstOperationContens = "データの"


'====================================================================================================
' メンバ変数
'====================================================================================================
Private mDataEntrySheet As DataEntrySheet       ' データ投入シート
Private mTableSheet As TableSheet               ' テーブルシート
Private mDataEntryLogic As DataEntryLogic       ' データ投入ロジック
Private mEntryTypeName As String                ' 投入種別名


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

    ' プログレスバー表示
    Call StatusBar.ShowProgress(cstProcessName & mEntryTypeName, GetRecordTotal(tableSettings), 0)

    ' 対象テーブル設定を全て処理
    For Each xKey In tableSettings
        Call ApplicationEx.CheckEvents

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
    ' 実行結果の表示
    Call ApplicationEx.ShowExecutionResult(cstOperationContens & mEntryTypeName)
    ' ステータスバー非表示
    Call StatusBar.Hide
End Sub


'====================================================================================================
' 初期化
'----------------------------------------------------------------------------------------------------
' IN : xEntryType 投入種別
'====================================================================================================
Private Sub Initialize(xEntryType As EntryType)
    ' マクロ起動
    Call ApplicationEx.StartupMacro(SettingType.DataEntry)

    ' インスタンス生成
    Set mDataEntrySheet = New DataEntrySheet
    Set mTableSheet = New TableSheet
    Set mDataEntryLogic = New DataEntryLogic
    mEntryTypeName = GetEntryTypeName(xEntryType)
End Sub


'====================================================================================================
' 終了化
'====================================================================================================
Private Sub Finalize()
    ' インスタンス解放
    Set mDataEntrySheet = Nothing
    Set mTableSheet = Nothing
    Set mDataEntryLogic = Nothing

    ' マクロ停止
    Call ApplicationEx.ShutdownMacro
End Sub


'====================================================================================================
' 投入種別名を取得します
'----------------------------------------------------------------------------------------------------
' IN : xEntryType 投入種別
' OUT: 投入種別名
'====================================================================================================
Private Function GetEntryTypeName(xEntryType As EntryType)
    Select Case xEntryType
        Case EntryType.Register
            GetEntryTypeName = "登録"
        Case EntryType.Update
            GetEntryTypeName = "更新"
        Case EntryType.Delete
            GetEntryTypeName = "削除"
    End Select
End Function


'====================================================================================================
' レコード総数を取得します
'----------------------------------------------------------------------------------------------------
' IN : tableSettings テーブル設定の連装配列
' OUT: レコード総数
'====================================================================================================
Private Function GetRecordTotal(tableSettings As Dictionary)
    Dim xKey As Variant
    Dim result As Long

    For Each xKey In tableSettings
        result = result + tableSettings(xKey).DataCount
    Next
    GetRecordTotal = result
End Function
