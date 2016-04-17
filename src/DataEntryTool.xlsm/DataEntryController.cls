VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataEntryController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' �e�[�u���V�[�g�̃R���g���[��
'
'====================================================================================================

'====================================================================================================
' �萔
'====================================================================================================
Private Const cstProcessName = "�f�[�^"


'====================================================================================================
' �����o�ϐ�
'====================================================================================================
Private mDataEntrySheet As DataEntrySheet       ' �f�[�^�����V�[�g
Private mTableSheet As TableSheet               ' �e�[�u���V�[�g
Private mDataEntryLogic As DataEntryLogic       ' �f�[�^�������W�b�N
Private mEntryTypeName As String                ' ������ʖ�


'====================================================================================================
' ���s
'----------------------------------------------------------------------------------------------------
' IN : xEntryType �������
'====================================================================================================
Public Sub Execute(xEntryType As EntryType)
On Error GoTo Finally
    Dim tableSettings As Dictionary
    Dim xKey As Variant
    Dim ts As TableSetting
    Dim ed As EntryData
    Dim procCnt As Long

    ' ������
    Call Initialize(xEntryType)

    ' ���������̃N���A
    Call mDataEntrySheet.ClearProcessingCount

    ' �Ώۃe�[�u���ݒ�̎擾
    Set tableSettings = mDataEntrySheet.GetTableSettings(True)
    If tableSettings.Count = 0 Then
        Err.Raise ErrNumber.Warning, , "�f�[�^�����Ώۂ̃f�[�^������܂���B" & vbNewLine & vbNewLine & _
            "���L�菇�����{���ăf�[�^�����Ώۂ̃f�[�^��ݒ肵�Ă��������B" & vbNewLine & _
            "  �E�e�[�u���ꗗ�̃f�[�^�����Ώۗ�ɋ󕶎��ȊO�̒l��ݒ肷��B" & vbNewLine & _
            "  �E�f�[�^�����Ώۂ̃e�[�u���V�[�g�Ƀf�[�^����͂���B"
    End If

    ' �v���O���X�o�[�\��
    Call StatusBar.ShowProgress(cstProcessName & mEntryTypeName, GetRecordTotal(tableSettings), 0)

    ' �Ώۃe�[�u���ݒ��S�ď���
    For Each xKey In tableSettings
        Call ApplicationEx.CheckEvents

        ' �Ώۃe�[�u���̃e�[�u���f�[�^�̎擾
        Set ts = tableSettings(xKey)
        Set ed = mTableSheet.GetEntryData(ts.PhysicsName)

        ' �f�[�^�������s
        procCnt = mDataEntryLogic.ExecuteDataEntry(xEntryType, ed)

        ' ���������̏�������
        Call mDataEntrySheet.WriteProcessingCount(ts, procCnt)
    Next

Finally:
    ' �I����
    Call Finalize
    ' ���s���ʂ̕\��
    Call ApplicationEx.ShowExecutionResult(cstProcessName & mEntryTypeName)
    ' �X�e�[�^�X�o�[��\��
    Call StatusBar.Hide
End Sub


'====================================================================================================
' ������
'----------------------------------------------------------------------------------------------------
' IN : xEntryType �������
'====================================================================================================
Private Sub Initialize(xEntryType As EntryType)
    ' �}�N���N��
    Call ApplicationEx.StartupMacro(SettingType.DataEntry)

    ' �C���X�^���X����
    Set mDataEntrySheet = New DataEntrySheet
    Set mTableSheet = New TableSheet
    Set mDataEntryLogic = New DataEntryLogic
    mEntryTypeName = GetEntryTypeName(xEntryType)
End Sub


'====================================================================================================
' �I����
'====================================================================================================
Private Sub Finalize()
    ' �C���X�^���X���
    Set mDataEntrySheet = Nothing
    Set mTableSheet = Nothing
    Set mDataEntryLogic = Nothing

    ' �}�N����~
    Call ApplicationEx.ShutdownMacro
End Sub


'====================================================================================================
' ������ʖ����擾���܂�
'----------------------------------------------------------------------------------------------------
' IN : xEntryType �������
' OUT: ������ʖ�
'====================================================================================================
Private Function GetEntryTypeName(xEntryType As EntryType)
    Select Case xEntryType
        Case EntryType.Register
            GetEntryTypeName = "�o�^"
        Case EntryType.Update
            GetEntryTypeName = "�X�V"
        Case EntryType.Delete
            GetEntryTypeName = "�폜"
    End Select
End Function


'====================================================================================================
' ���R�[�h�������擾���܂�
'----------------------------------------------------------------------------------------------------
' IN : tableSettings �e�[�u���ݒ�̘A���z��
' OUT: ���R�[�h����
'====================================================================================================
Private Function GetRecordTotal(tableSettings As Dictionary)
    Dim xKey As Variant
    Dim result As Long

    For Each xKey In tableSettings
        result = result + tableSettings(xKey).DataCount
    Next
    GetRecordTotal = result
End Function