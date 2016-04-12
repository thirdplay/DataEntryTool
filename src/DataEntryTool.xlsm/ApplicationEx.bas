Attribute VB_Name = "ApplicationEx"
Option Explicit
Option Private Module

'====================================================================================================
'
' アプリケーションの拡張モジュール
'
'====================================================================================================

'====================================================================================================
' メンバ変数
'====================================================================================================
Private mMacroType As MacroType     ' マクロ種別
Private mCalculation As Long        ' 自動計算方法(退避用)
Private mTime As Variant

'====================================================================================================
' マクロ起動
'----------------------------------------------------------------------------------------------------
' IN : xMacroType マクロ種別
'====================================================================================================
Public Sub StartupMacro(xMacroType As MacroType)
    mMacroType = xMacroType
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    'Application.EnableEvents = False
    mCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.Cursor = xlWait

    ' 設定モジュールの構成
    Call Setting.Setup(mMacroType)

    ' データベース接続
    If (mMacroType And MacroType.Database) = MacroType.Database Then
        Call Database.Connect
    End If
End Sub


'====================================================================================================
' マクロ停止
'====================================================================================================
Public Sub ShutdownMacro()
    ' データベース切断
    If (mMacroType And MacroType.Database) = MacroType.Database Then
        Call Database.Disconnect
    End If

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    'Application.EnableEvents = True
    Application.Calculation = mCalculation
    Application.Cursor = xlDefault
End Sub


'====================================================================================================
' 実行結果を表示します
'----------------------------------------------------------------------------------------------------
' IN : result          実行結果
'    : operationDetail 操作内容
'====================================================================================================
Public Sub ShowExecutionResult(operationDetail As String)
    If Err.Number = 0 Then
        MsgBox operationDetail & "が完了しました。", vbOKOnly + vbInformation
    ElseIf Err.Number = ErrNumber.Warning Then
        MsgBox Err.Description, vbOKOnly + vbExclamation
    Else
        MsgBox operationDetail & "に失敗しました。" & vbNewLine & vbNewLine & Err.Description, vbOKOnly + vbCritical
    End If
End Sub
