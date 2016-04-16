Attribute VB_Name = "ApplicationEx"
Option Explicit
Option Private Module

'====================================================================================================
'
' �A�v���P�[�V�����̊g�����W���[��
'
'====================================================================================================

'====================================================================================================
' �����o�ϐ�
'====================================================================================================
Private mCalculation As Long            ' �����v�Z���@(�ޔ�p)


'====================================================================================================
' �}�N���N��
' ---------------------------------------------------------------------------------------------------
' IN : xSettingType �ݒ���
'====================================================================================================
Public Sub StartupMacro(xSettingType As SettingType)
    ' �`��/�v�Z�}��
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    'Application.EnableEvents = False
    mCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.Cursor = xlWait

    ' �ݒ胂�W���[���̍\��
    Call Setting.Setup(xSettingType)

    ' �f�[�^����Dao�t�@�N�g���̏�����
    If (xSettingType And SettingType.Database) = SettingType.Database Then
        Call DataEntryDaoFactory.Initialize
    End If
End Sub


'====================================================================================================
' �}�N����~
'====================================================================================================
Public Sub ShutdownMacro()
    ' �f�[�^����Dao�t�@�N�g���̏I����
    Call DataEntryDaoFactory.Finalize

    ' �`��/�v�Z�}������
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    'Application.EnableEvents = True
    Application.Calculation = mCalculation
    Application.Cursor = xlDefault
End Sub


'====================================================================================================
' ���s���ʂ�\�����܂�
'----------------------------------------------------------------------------------------------------
' IN : result          ���s����
'    : operationDetail ������e
'====================================================================================================
Public Sub ShowExecutionResult(operationDetail As String)
    If Err.Number = 0 Then
        MsgBox operationDetail & "���������܂����B", vbOKOnly + vbInformation
    ElseIf Err.Number = ErrNumber.Warning Then
        MsgBox Err.Description, vbOKOnly + vbExclamation
    Else
        MsgBox operationDetail & "�Ɏ��s���܂����B" & vbNewLine & vbNewLine & Err.Description, vbOKOnly + vbCritical
    End If
End Sub
