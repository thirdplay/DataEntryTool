VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DmlOutputController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' DML出力のコントローラー
'
'====================================================================================================

'====================================================================================================
' メンバ変数
'====================================================================================================
Private mBaseController As BaseController       ' 基本コントローラー
Private mDataEntrySheet As DataEntrySheet       ' データ投入シート
Private mTableSheet As TableSheet               ' テーブルシート
Private mDmlOutputLogic As DmlOutputLogic       ' DML出力ロジック
Private mProcessName As String                  ' 処理名


'====================================================================================================
' 実行
'----------------------------------------------------------------------------------------------------
' IN : xEntryType 投入種別
'====================================================================================================
Public Sub Execute(xEntryType As EntryType)
On Error GoTo Finally
    Dim tableSettings As Dictionary
    Dim xOutputDir As String
    Dim xKey As Variant
    Dim ts As TableSetting
    Dim ed As EntryData

    ' 初期化
    Call Initialize(xEntryType)

    ' 対象テーブル設定の取得
    Set tableSettings = mDataEntrySheet.GetTableSettings(True)
    If tableSettings.Count = 0 Then
        Err.Raise ErrNumber.Warning, , "データ投入対象のデータがありません。" & vbNewLine & vbNewLine & _
            "下記手順を実施してデータ投入対象のデータを設定してください。" & vbNewLine & _
            "  ・テーブル一覧のデータ投入対象列に空文字以外の値を設定する。" & vbNewLine & _
            "  ・データ投入対象のテーブルシートにデータを入力する。"
    End If

    ' 出力先の取得
    xOutputDir = mDmlOutputLogic.GetOutputDir
    If xOutputDir = "" Then
        Err.Raise ErrNumber.Abort
    End If

    ' 対象テーブル設定を全て処理
    For Each xKey In tableSettings
        ' 対象テーブルのテーブルデータの取得
        Set ts = tableSettings(xKey)
        Set ed = mTableSheet.GetEntryData(ts.PhysicsName)

        ' DML出力
        Call mDmlOutputLogic.OutputDml(xEntryType, ed)
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
    Set mDmlOutputLogic = New DmlOutputLogic
    mProcessName = GetProcessName(xEntryType)
End Sub


'====================================================================================================
' 終了化
'====================================================================================================
Private Sub Finalize()
    ' インスタンス解放
    Set mDataEntrySheet = Nothing
    Set mTableSheet = Nothing
    Set mDmlOutputLogic = Nothing

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
            GetProcessName = "登録DML出力"
        Case EntryType.Update
            GetProcessName = "更新DML出力"
        Case EntryType.Delete
            GetProcessName = "削除DML出力"
    End Select
End Function
