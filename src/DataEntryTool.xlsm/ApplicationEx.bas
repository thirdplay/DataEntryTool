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
    'Application.EnableEvents = False
    'mCalculation = Application.Calculation
    'Application.Calculation = xlCalculationManual
    'Application.Cirspr = xlWait 'xlDefault
End Sub


'====================================================================================================
' ���s���ʂ�\�����܂�
'----------------------------------------------------------------------------------------------------
' IN : result          ���s����
'    : operationDetail ������e
'====================================================================================================
Public Sub ShowExecutionResult(operationDetail As String)
    If Err.Number = 0 Then
        MsgBoxEx.Information operationDetail & "���������܂����B"
    Else
        MsgBoxEx.Error operationDetail & "�Ɏ��s���܂����B", Err.Description
    End If
End Sub
