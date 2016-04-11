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

'====================================================================================================
' マクロ起動
'----------------------------------------------------------------------------------------------------
' IN : xMacroType マクロ種別
' OUT: True:成功、False:失敗
'====================================================================================================
Public Function StartupMacro(xMacroType As MacroType) As Boolean
    mMacroType = xMacroType
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    mCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.Cursor = xlWait

    ' 設定モジュールの構成
    If Not Setting.Setup(mMacroType) Then
        StartupMacro = False
    End If

    ' データベース接続
    If (mMacroType And MacroType.Database) = MacroType.Database Then
        Call Database.Connect
    End If
    StartupMacro = True
End Function


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
    Application.EnableEvents = True
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
        MsgBoxEx.Information operationDetail & "が完了しました。"
    Else
        MsgBoxEx.Error operationDetail & "に失敗しました。", Err.Description
    End If
End Sub
