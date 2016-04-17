Attribute VB_Name = "StatusBar"
Option Explicit
Option Private Module

'====================================================================================================
'
' �X�e�[�^�X�o�[�̃��W���[��
'
'====================================================================================================

'====================================================================================================
' �萔
'====================================================================================================
' �X�e�[�^�X�o�[�̏���
Private Const cstStatusBarFormat = "${processName} ${progressRate}% ${progressBar}"


'====================================================================================================
' �����o�ϐ�
'====================================================================================================
Private mDisplayStatusBar As Boolean    ' �X�e�[�^�X�o�[�\���t���O(�ޔ�p)
Private mIsDisplay As Boolean           ' �\���t���O
Private mProcessName As String          ' �v���Z�X��
Private mProgressCnt As Long            ' �i���J�E���g
Private mProgressMax As Long            ' �i���ő�J�E���g
Private mStatusBarContens As String     ' �X�e�[�^�X�o�[�̓��e


'====================================================================================================
' �X�e�[�^�X�o�[�ɐi���󋵂�\�����܂�
' ---------------------------------------------------------------------------------------------------
' IN : processName �v���Z�X��
'    : progressMax �i���J�E���g�̍ő�l
'    : progressCnt �����i���J�E���g
'====================================================================================================
Public Sub ShowProgress(ByVal processName As String, Optional ByVal progressMax As Long = 100, Optional progressCnt As Long = 0)
    If Not mIsDisplay Then
        ' �X�e�[�^�X�o�[�̕\��
        mDisplayStatusBar = Application.DisplayStatusBar
        Application.DisplayStatusBar = True

        ' �i���󋵂̏�����
        mIsDisplay = True
        mProcessName = processName
        mProgressMax = progressMax
        mProgressCnt = 0

        ' �i���󋵂̑���
        Call IncreaseProgress(progressCnt)
    End If
End Sub


'====================================================================================================
' �X�e�[�^�X�o�[���\���ɂ��܂�
'====================================================================================================
Public Sub Hide()
    If mIsDisplay Then
        mIsDisplay = False
        mProcessName = ""
        mProgressCnt = 0
        mProgressMax = 0
        mStatusBarContens = ""

        ' �X�e�[�^�X�o�[�̕\���ݒ�𕜌�����
        Application.StatusBar = False
        Application.DisplayStatusBar = mDisplayStatusBar
    End If
End Sub


'====================================================================================================
' �i���󋵂𑝂₵�܂�
' ---------------------------------------------------------------------------------------------------
' IN : progressCnt �i���J�E���g
'====================================================================================================
Public Sub IncreaseProgress(Optional ByVal progressCnt As Long = 1)
    Dim statusBarContens As String
    Dim progressRate As Byte
    Dim progressStatus As Byte

    ' �\�����ȊO�͏������Ȃ�'
    If Not mIsDisplay Then
        Exit Sub
    End If

    ' �i�����̍X�V
    mProgressCnt = mProgressCnt + progressCnt
    If mProgressCnt > mProgressMax Then
        mProgressCnt = mProgressMax
    End If

    ' �i���󋵂̍쐬
    progressRate = CByte(mProgressCnt / mProgressMax * 100)
    progressStatus = progressRate / 10
    statusBarContens = cstStatusBarFormat
    statusBarContens = Replace(statusBarContens, "${processName}", mProcessName)
    statusBarContens = Replace(statusBarContens, "${progressRate}", progressRate)
    statusBarContens = Replace(statusBarContens, "${progressBar}", String(progressStatus, "��") & String(10 - progressStatus, "��"))

    ' �X�e�[�^�X�o�[�̓��e�ɕω�������ꍇ�A�X�V
    If statusBarContens <> mStatusBarContens Then
        Application.StatusBar = statusBarContens
        mStatusBarContens = statusBarContens
    End If
End Sub
