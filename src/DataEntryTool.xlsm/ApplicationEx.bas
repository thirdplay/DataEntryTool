Attribute VB_Name = "ApplicationEx"
Option Explicit
Option Private Module

'====================================================================================================
'
' �A�v���P�[�V�����̊g�����W���[��
'
'====================================================================================================

'====================================================================================================
' ��ʕ`���}�����܂�
'----------------------------------------------------------------------------------------------------
' IN : isSuppress �}���t���O(True:�}������AFalse:�}�����Ȃ�)
'====================================================================================================
Public Sub SuppressScreenDrawing(isSuppress As Boolean)
    Application.ScreenUpdating = Not isSuppress
    Application.DisplayAlerts = Not isSuppress
End Sub


'====================================================================================================
' ���s���ʂ�\�����܂�
'----------------------------------------------------------------------------------------------------
' IN : result          ���s����
'    : operationDetail ������e
'====================================================================================================
Public Sub ShowExecutionResult(result As Boolean, operationDetail As String)
    If result Then
        MsgBoxEx.Information operationDetail & "���������܂����B"
    Else
        MsgBoxEx.Error operationDetail & "�Ɏ��s���܂����B", Err.Description
    End If
End Sub

