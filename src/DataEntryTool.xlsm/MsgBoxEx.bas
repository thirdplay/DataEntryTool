Attribute VB_Name = "MsgBoxEx"
Option Explicit
Option Private Module

'====================================================================================================
'
' ���b�Z�[�W�{�b�N�X�̊g�����W���[��
'
'====================================================================================================

'====================================================================================================
' �x�����b�Z�[�W��\�����܂�
'----------------------------------------------------------------------------------------------------
' IN : message ���b�Z�[�W
'    : detailMessage �ڍ׃��b�Z�[�W
'====================================================================================================
Public Sub Warning(ByVal message As String, ByVal Optional detailMessage As String = "")
    If detailMessage <> "" Then
        message = message & vbNewLine & vbNewLine & detailMessage
    End If
    MsgBox message, vbOKOnly + vbExclamation
End Sub


'====================================================================================================
' �G���[���b�Z�[�W��\�����܂�
'----------------------------------------------------------------------------------------------------
' IN : message ���b�Z�[�W
'    : detailMessage �ڍ׃��b�Z�[�W
'====================================================================================================
Public Sub Error(message As String, detailMessage As String)
    MsgBox message & vbNewLine & vbNewLine & detailMessage, vbOKOnly + vbCritical
End Sub


