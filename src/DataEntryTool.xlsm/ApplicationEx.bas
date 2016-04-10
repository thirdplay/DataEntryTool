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
End Sub


'====================================================================================================
' 実行結果を表示します
'----------------------------------------------------------------------------------------------------
' IN : result          実行結果
'    : operationDetail 操作内容
'====================================================================================================
Public Sub ShowExecutionResult(result As Boolean, operationDetail As String)
    If result Then
        MsgBoxEx.Information operationDetail & "が完了しました。"
    Else
        MsgBoxEx.Error operationDetail & "に失敗しました。", Err.Description
    End If
End Sub

