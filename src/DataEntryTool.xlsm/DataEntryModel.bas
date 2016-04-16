Attribute VB_Name = "DataEntryModel"
Option Explicit
Option Private Module

'====================================================================================================
'
' �f�[�^�������f���̃��W���[��
'
'====================================================================================================

'====================================================================================================
' �f�[�^���������s���܂�
'----------------------------------------------------------------------------------------------------
' IN : xEntryType �������
'    : xEntryData �����f�[�^
' OUT: ��������
'====================================================================================================
Public Function ExecuteDataEntry(xEntryType As EntryType, xEntryData As EntryData) As Long
On Error GoTo ErrHandler
    Dim i As Long
    Dim queries As Collection
    Dim query As String
    Dim procCnt As Long

    ' �g�����U�N�V�����J�n
    Call Database.BeginTrans

    ' �N�G������
    Select Case xEntryType
    Case EntryType.Register
        Set queries= xEntryData.MakeInsertQueries()
    Case EntryType.Update
        Set queries = xEntryData.MakeUpdateQueries()
    Case EntryType.Delete
        Set queries = xEntryData.MakeDeleteQueries()
    End Select

    ' �f�[�^����
    For i = 1 To queries.Count
        procCnt = procCnt + Database.ExecuteQuery (queries(i))
    Next

    ' �R�~�b�g
    Call Database.CommitTrans

    ExecuteDataEntry = procCnt
    Exit Function
ErrHandler:
    ' ���[���o�b�N
    Call Database.RollbackTrans
    Err.Raise Err.Number, Err.Source, _
        "[�������]" & vbNewLine & _
        "�e�[�u����:" & xEntryData.TableName & vbNewLine & _
        "�s��:" & (cstTableRecordBase + i - 1) & vbNewLine & _
        "[�G���[���e]" & vbNewLine & _
        Err.Description, Err.HelpFile, Err.HelpContext
    Exit Function
End Function
