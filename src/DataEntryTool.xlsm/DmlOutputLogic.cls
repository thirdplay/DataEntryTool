VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DmlOutputLogic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' �f�[�^�������W�b�N
'
'====================================================================================================

'====================================================================================================
' �萔
'====================================================================================================
' �N�G���̐ړ���
Private Const cstQuerySuffix = ";" & vbNewLine
' DML�t�@�C���g���q
Private Const cstFileExt = ".sql"


'====================================================================================================
' �����o�ϐ�
'====================================================================================================
Private mDataEntryLogic As DataEntryLogic


'====================================================================================================
' �R���X�g���N�^
'====================================================================================================
Public Sub Class_Initialize()
    Set mDataEntryLogic = New DataEntryLogic
End Sub


'====================================================================================================
' �f�X�g���N�^
'====================================================================================================
Public Sub Class_Terminate()
    Set mDataEntryLogic = Nothing
End Sub


'====================================================================================================
' DML���o�͂��܂�
'----------------------------------------------------------------------------------------------------
' IN : xOutputDir �o�̓f�B���N�g��
'    : xEntryType �������
'    : xEntryData �����f�[�^
'====================================================================================================
Public Sub OutputDml(xOutputDir As String, xEntryType As EntryType, xEntryData As EntryData)
On Error GoTo Finally
    Dim stream As ADODB.Stream
    Dim queries As Collection
    Dim query

    ' �f�[�^�����N�G���̐���
    Set queries = mDataEntryLogic.MakeQueries(xEntryType, xEntryData)

    ' �X�g���[���I�[�v��
    Set stream = New ADODB.Stream
    stream.Mode = adModeReadWrite   ' �ǂݎ��/�������݌���
    stream.Type = adTypeText        '�e�L�X�g�f�[�^
    stream.Charset = "Shift_JIS"    '�����R�[�h
    stream.Open

    For Each query In queries
        Call stream.WriteText(query & cstQuerySuffix, adWriteChar)
    Next
    stream.SaveToFile xOutputDir & xEntryData.TableName & cstFileExt, adSaveCreateOverWrite
Finally:
    ' �X�g���[���N���[�Y
    stream.Close
    Set stream = Nothing

    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Sub