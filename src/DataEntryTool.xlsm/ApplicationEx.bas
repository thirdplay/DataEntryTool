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

'====================================================================================================
' �}�N���N��
'----------------------------------------------------------------------------------------------------
' IN : xMacroType �}�N�����
' OUT: True:�����AFalse:���s
'====================================================================================================
Public Function StartupMacro(xMacroType As MacroType) As Boolean
    mMacroType = xMacroType
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    mCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.Cursor = xlWait

    ' �ݒ胂�W���[���̍\��
    If Not Setting.Setup(mMacroType) Then
        StartupMacro = False
    End If

    ' �f�[�^�x�[�X�ڑ�
    If (mMacroType And MacroType.Database) = MacroType.Database Then
        Call Database.Connect
    End If
    StartupMacro = True
End Function


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
    Application.EnableEvents = True
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
        MsgBoxEx.Information operationDetail & "���������܂����B"
    Else
        MsgBoxEx.Error operationDetail & "�Ɏ��s���܂����B", Err.Description
    End If
End Sub
