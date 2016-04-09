Attribute VB_Name = "MsgBoxEx"
Option Explicit
Option Private Module

'====================================================================================================
'
' メッセージボックスの拡張モジュール
'
'====================================================================================================

'====================================================================================================
' 警告メッセージを表示します
'----------------------------------------------------------------------------------------------------
' IN : message メッセージ
'    : detailMessage 詳細メッセージ
'====================================================================================================
Public Sub Warning(ByVal message As String, ByVal Optional detailMessage As String = "")
    If detailMessage <> "" Then
        message = message & vbNewLine & vbNewLine & detailMessage
    End If
    MsgBox message, vbOKOnly + vbExclamation
End Sub


'====================================================================================================
' エラーメッセージを表示します
'----------------------------------------------------------------------------------------------------
' IN : message メッセージ
'    : detailMessage 詳細メッセージ
'====================================================================================================
Public Sub Error(message As String, detailMessage As String)
    MsgBox message & vbNewLine & vbNewLine & detailMessage, vbOKOnly + vbCritical
End Sub


