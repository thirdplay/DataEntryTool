VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BaseController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' ��{�R���g���[���[
'
'====================================================================================================

'====================================================================================================
' �����o�ϐ�
'====================================================================================================
Private mCalculation As Long            ' �����v�Z���@(�ޔ�p)


'====================================================================================================
' �N��
' ---------------------------------------------------------------------------------------------------
' IN : xSettingType �ݒ���
'====================================================================================================
Public Sub Startup(xSettingType As SettingType)
    ' �`��/�v�Z�}��
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
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
' ��~
'====================================================================================================
Public Sub Shutdown()
    ' �f�[�^����Dao�t�@�N�g���̏I����
    Call DataEntryDaoFactory.Finalize

    ' �`��/�v�Z�}������
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = mCalculation
    Application.Cursor = xlDefault
End Sub


'====================================================================================================
' ���s���ʂ�\�����܂�
'----------------------------------------------------------------------------------------------------
' IN : operationContens ������e
'====================================================================================================
Public Sub ShowExecutionResult(operationContens As String)
    If Err.Number = 0 Then
        MsgBox operationContens & "���������܂����B", vbOKOnly + vbInformation
    ElseIf Err.Number = ErrNumber.Warning Then
        MsgBox Err.Description, vbOKOnly + vbExclamation
    Else
        MsgBox operationContens & "�Ɏ��s���܂����B" & vbNewLine & vbNewLine & Err.Description, vbOKOnly + vbCritical
    End If
End Sub