Attribute VB_Name = "ApplicationEx"
Option Explicit
Option Private Module

'====================================================================================================
'
' アプリケーションの拡張モジュール
'
'====================================================================================================

'====================================================================================================
' 画面描画を抑制します
'----------------------------------------------------------------------------------------------------
' IN : isSuppress 抑制フラグ(True:抑制する、False:抑制しない)
'====================================================================================================
Public Sub SuppressScreenDrawing(isSuppress As Boolean)
    Application.ScreenUpdating = Not isSuppress
    Application.DisplayAlerts = Not isSuppress
    'Application.EnableEvents = False
    'mCalculation = Application.Calculation
    'Application.Calculation = xlCalculationManual
    'Application.Cirspr = xlWait 'xlDefault
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
