VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataEntryLogic"
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
' Insert�N�G��
Private Const cstInsertQuery = "INSERT INTO ${tableName} (${columns}) VALUES (${values})"
' Update�N�G��
Private Const cstUpdateQuery = "UPDATE ${tableName} SET ${set} WHERE ${where}"
' Delete�N�G��
Private Const cstDeleteQuery = "DELETE FROM ${tableName} WHERE ${where}"
' ������r����
Private Const cstEqualComparisonCriteria = "${column} = ${value}"
' ��؂蕶��
Private Const cstDelimiter = ","
' Where��̋�؂蕶��
Private Const cstWherePhraseDelimiter = " AND "


'====================================================================================================
' �����o�ϐ�
'====================================================================================================
Private mDataEntryDao As DataEntryDao


'====================================================================================================
' �R���X�g���N�^
'====================================================================================================
Public Sub Class_Initialize()
    Set mDataEntryDao = DataEntryDaoFactory.GetDataEntryDao
End Sub


'====================================================================================================
' �f�X�g���N�^
'====================================================================================================
Public Sub Class_Terminate()
    Set mDataEntryDao = Nothing
End Sub


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
    Call mDataEntryDao.BeginTrans

    ' �N�G������
    Set queries = MakeQueries(xEntryType, xEntryData)
    ' �f�[�^����
    For i = 1 To queries.Count
        Call ApplicationEx.CheckEvents
        procCnt = procCnt + mDataEntryDao.ExecuteUpdate(queries(i))
        Call StatusBar.IncreaseProgress
    Next

    ' �R�~�b�g
    Call mDataEntryDao.CommitTrans

    ExecuteDataEntry = procCnt
    Exit Function
ErrHandler:
    ' ���[���o�b�N
    Call mDataEntryDao.RollbackTrans
    Err.Raise Err.Number, Err.Source, _
        "[�������]" & vbNewLine & _
        "�e�[�u����:" & xEntryData.TableName & vbNewLine & _
        "�s��:" & (cstTableRecordBase + i - 1) & vbNewLine & _
        "[�G���[���e]" & vbNewLine & _
        Err.Description, Err.HelpFile, Err.HelpContext
    Exit Function
End Function


'====================================================================================================
' �N�G���𐶐����܂�
'----------------------------------------------------------------------------------------------------
' IN : xEntryType �������
'    : xEntryData �����f�[�^
' OUT: �N�G�����X�g
'====================================================================================================
Private Function MakeQueries(xEntryType As EntryType, xEntryData As EntryData)
    Select Case xEntryType
    Case EntryType.Register
        Set MakeQueries = MakeInsertQueries(xEntryData)
    Case EntryType.Update
        Set MakeQueries = MakeUpdateQueries(xEntryData)
    Case EntryType.Delete
        Set MakeQueries = MakeDeleteQueries(xEntryData)
    End Select
End Function


'====================================================================================================
' Insert�N�G���𐶐����܂�
'----------------------------------------------------------------------------------------------------
' IN : xEntryData �����f�[�^
' OUT: �N�G�����X�g
'====================================================================================================
Private Function MakeInsertQueries(xEntryData As EntryData) As Collection
    Dim i As Long
    Dim query As String
    Dim queries As Collection

    Set queries = New Collection
    For i = 1 To xEntryData.RecordRange.Rows.Count
        query = cstInsertQuery
        query = Replace(query, "${tableName}", xEntryData.TableName)
        query = Replace(query, "${columns}", GetColumnPhrase(xEntryData.ColumnDefinitions))
        query = Replace(query, "${values}", GetValuePhrase(xEntryData.ColumnDefinitions, xEntryData.RecordRange.Rows(i)))
        Call queries.Add(query)
    Next
    Set MakeInsertQueries = queries
End Function


'====================================================================================================
' Update�N�G���𐶐����܂�
'----------------------------------------------------------------------------------------------------
' IN : xEntryData �����f�[�^
' OUT: �N�G�����X�g
'====================================================================================================
Private Function MakeUpdateQueries(xEntryData As EntryData) As Collection
    Dim i As Long
    Dim query As String
    Dim queries As Collection

    Set queries = New Collection
    For i = 1 To xEntryData.RecordRange.Rows.Count
        query = cstUpdateQuery
        query = Replace(query, "${tableName}", xEntryData.TableName)
        query = Replace(query, "${set}", GetSetPhrase(xEntryData.ColumnDefinitions, xEntryData.RecordRange.Rows(i)))
        query = Replace(query, "${where}", GetWherePhrase(xEntryData.ColumnDefinitions, xEntryData.RecordRange.Rows(i)))
        Call queries.Add(query)
    Next
    Set MakeUpdateQueries = queries
End Function


'====================================================================================================
' Delete�N�G���𐶐����܂�
'----------------------------------------------------------------------------------------------------
' IN : xEntryData �����f�[�^
' OUT: �N�G�����X�g
'====================================================================================================
Private Function MakeDeleteQueries(xEntryData As EntryData) As Collection
    Dim i As Long
    Dim query As String
    Dim queries As Collection

    Set queries = New Collection
    For i = 1 To xEntryData.RecordRange.Rows.Count
        query = cstDeleteQuery
        query = Replace(query, "${tableName}", xEntryData.TableName)
        query = Replace(query, "${where}", GetWherePhrase(xEntryData.ColumnDefinitions, xEntryData.RecordRange.Rows(i)))
        Call queries.Add(query)
    Next
    Set MakeDeleteQueries = queries
End Function


'====================================================================================================
' Insert����Column����擾���܂�
'----------------------------------------------------------------------------------------------------
' IN : xColumnDefinitions �J������`���X�g
' OUT: Column��
'====================================================================================================
Private Function GetColumnPhrase(xColumnDefinitions As Collection) As String
    Dim result As String
    Dim cd As ColumnDefinition

    For Each cd In xColumnDefinitions
        result = result & cd.ColumnName & cstDelimiter
    Next
    GetColumnPhrase = Left(result, Len(result) - Len(cstDelimiter))
End Function


'====================================================================================================
' Insert����Value����擾���܂�
'----------------------------------------------------------------------------------------------------
' IN : cd �J������`���X�g
'    : record ���R�[�h
' OUT: Value��
'====================================================================================================
Private Function GetValuePhrase(cd As Collection, record As Range) As String
    Dim result As String
    Dim i As Long

    For i = 1 To record.Columns.Count
        result = result & mDataEntryDao.GetItemValue(record.Cells(1, i), cd(i).DataType) & cstDelimiter
    Next
    GetValuePhrase = Left(result, Len(result) - Len(cstDelimiter))
End Function


'====================================================================================================
' Update����Set����擾���܂�
'----------------------------------------------------------------------------------------------------
' IN : cd �J������`���X�g
'    : record ���R�[�h
' OUT: Set��
'====================================================================================================
Private Function GetSetPhrase(cd As Collection, record As Range) As String
    Dim result As String
    Dim i As Long

    For i = 1 To record.Columns.Count
        If cd(i).IsPrimaryKey = "" Then
            result = result & cstEqualComparisonCriteria
            result = Replace(result, "${column}", cd(i).ColumnName)
            result = Replace(result, "${value}", mDataEntryDao.GetItemValue(record.Cells(1, i), cd(i).DataType))
            result = result & cstDelimiter
        End If
    Next
    GetSetPhrase = Left(result, Len(result) - Len(cstDelimiter))
End Function


'====================================================================================================
' Where����擾���܂�
'----------------------------------------------------------------------------------------------------
' IN : cd �J������`���X�g
'    : record ���R�[�h
' OUT: Value��
'====================================================================================================
Private Function GetWherePhrase(cd As Collection, record As Range) As String
    Dim result As String
    Dim i As Long

    For i = 1 To record.Columns.Count
        If cd(i).IsPrimaryKey <> "" Then
            result = result & cstEqualComparisonCriteria
            result = Replace(result, "${column}", cd(i).ColumnName)
            result = Replace(result, "${value}", mDataEntryDao.GetItemValue(record.Cells(1, i), cd(i).DataType))
            result = result & cstWherePhraseDelimiter
        End If
    Next
    GetWherePhrase = Left(result, Len(result) - Len(cstWherePhraseDelimiter))
End Function