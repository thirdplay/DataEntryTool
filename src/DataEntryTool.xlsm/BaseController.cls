VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BaseController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' 基本コントローラー
'
'====================================================================================================

'====================================================================================================
' メンバ変数
'====================================================================================================
Private mCalculation As Long            ' 自動計算方法(退避用)


'====================================================================================================
' 起動
' ---------------------------------------------------------------------------------------------------
' IN : xSettingType 設定種別
'====================================================================================================
Public Sub Startup(xSettingType As SettingType)
    ' 描画/計算抑制
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    mCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.Cursor = xlWait

    ' 設定モジュールの構成
    Call Setting.Setup(xSettingType)

    ' データ投入Daoファクトリの初期化
    If (xSettingType And SettingType.Database) = SettingType.Database Then
        Call DataEntryDaoFactory.Initialize
    End If
End Sub


'====================================================================================================
' 停止
'====================================================================================================
Public Sub Shutdown()
    ' データ投入Daoファクトリの終了化
    Call DataEntryDaoFactory.Finalize

    ' 描画/計算抑制解除
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = mCalculation
    Application.Cursor = xlDefault
End Sub


'====================================================================================================
' 実行結果を表示します
'----------------------------------------------------------------------------------------------------
' IN : operationContens 操作内容
'====================================================================================================
Public Sub ShowExecutionResult(operationContens As String)
    Select Case Err.Number
        Case 0
            MsgBox operationContens & "が完了しました。", vbOKOnly + vbInformation
        Case ErrNumber.Warning
            MsgBox Err.Description, vbOKOnly + vbExclamation
        Case ErrNumber.Abort
        Case Else
            MsgBox operationContens & "に失敗しました。" & vbNewLine & vbNewLine & Err.Description, vbOKOnly + vbCritical
    End Select
End Sub
