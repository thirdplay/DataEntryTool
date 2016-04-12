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
Private mMacroType As MacroType     ' �}�N�����
Private mCalculation As Long        ' �����v�Z���@(�ޔ�p)
Private mTime As Variant

'====================================================================================================
' �}�N���N��
'----------------------------------------------------------------------------------------------------
' IN : xMacroType �}�N�����
'====================================================================================================
Public Sub StartupMacro(xMacroType As MacroType)
    mMacroType = xMacroType
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    'Application.EnableEvents = False
    mCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.Cursor = xlWait

    ' �ݒ胂�W���[���̍\��
    Call Setting.Setup(mMacroType)

    ' �f�[�^�x�[�X�ڑ�
    If (mMacroType And MacroType.Database) = MacroType.Database Then
        Call Database.Connect
    End If
End Sub


'====================================================================================================
' �}�N����~
'====================================================================================================
Public Sub ShutdownMacro()
    ' �f�[�^�x�[�X�ؒf
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
