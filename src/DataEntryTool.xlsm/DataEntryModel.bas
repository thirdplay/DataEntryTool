Attribute VB_Name = "DataEntryModel"
Option Explicit
Option Private Module

'====================================================================================================
'
' �f�[�^�������f���̃��W���[��
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
Dim mDatabaseModel As DatabaseModel      ' �f�[�^�x�[�X���f��


'====================================================================================================
' �����f�[�^���擾���܂�
'----------------------------------------------------------------------------------------------------
' IN : xEntryType �������
'    : xTableName �e�[�u����
' OUT: �����f�[�^
'====================================================================================================
Public Function GetEntryData(xEntryType As EntryType, xTableName As String) As EntryData
    Dim ed As EntryData
    Dim cd As ColumnDefinition
    Dim rowIndex As Long
    Dim dataCount As Long
    Dim columnRange As Range
    Dim cr As Range

    With ThisWorkbook.Worksheets(xTableName)
        Set ed = New EntryData
        ed.EntryType = xEntryType
        ed.TableName = xTableName

        ' �J������`���X�g�̍쐬
        ed.ColumnDefinitions = New Collection
        dataCount = WorkSheetEx.GetColDataCount(xTableName, ColumnDefinitionRow.ColumnName, 1)
        Set columnRange = .Range(.Cells(1, 1), .Cells(ColumnDefinitionRow.Max, dataCount))
        For Each cr In columnRange.Columns
            Set cd = New ColumnDefinition
            cd.ColumnName = cr.Cells(ColumnDefinitionRow.ColumnName, 1)
            cd.DataType = cr.Cells(ColumnDefinitionRow.DataType, 1)
            cd.IsPrimaryKey = cr.Cells(ColumnDefinitionRow.IsPrimaryKey, 1)
            Call ed.ColumnDefinitions.Add(cd)
        Next

        ' ���R�[�h�͈͂̐ݒ�
        dataCount = WorkSheetEx.GetRowDataCount(xTableName, cstTableRecordBase, 1)
        ed.RecordRange = .Range(.Cells(cstTableRecordBase, 1), .Cells(cstTableRecordBase + dataCount - 1, ed.ColumnDefinitions.Count))

        Set GetEntryData = ed
    End With
End Function


'====================================================================================================
' �f�[�^���������s���܂�
'----------------------------------------------------------------------------------------------------
' IN : xEntryData �����f�[�^
' OUT: ��������
'====================================================================================================
Public Function ExecuteDataEntry(xEntryData As EntryData) As Long
On Error GoTo ErrHandler
    Dim i As Long
    Dim queries As Collection
    Dim procCnt As Long
    procCnt = 0

    ' �g�����U�N�V�����J�n
    Set mDatabaseModel = DatabaseModelFactory.Create()
    Call mDatabaseModel.BeginTrans

    ' �N�G������
    Set queries = MakeQueries(xEntryData)

    ' �N�G�����s
    For i = 1 To queries.Count
        procCnt = procCnt + mDatabaseModel.ExecuteQuery(queries(i))
    Next

    ' �R�~�b�g
    Call mDatabaseModel.CommitTrans

    ExecuteDataEntry = procCnt
    Exit Function
ErrHandler:
    ' ���[���o�b�N
    Call mDatabaseModel.RollbackTrans
    Err.Raise Err.Number, Err.Source, _
        "[�������]" & vbNewLine & _
        "�e�[�u����:" & xEntryData.TableName & vbNewLine & _
        "�f�[�^�s:" & (cstTableRecordBase + i - 1) & vbNewLine & _
        "[�G���[�ڍ�]" & vbNewLine & _
        Err.Description, Err.HelpFile, Err.HelpContext
    Exit Function
End Function


'====================================================================================================
' �N�G���𐶐����܂�
'----------------------------------------------------------------------------------------------------
' IN : xEntryData �����f�[�^
' OUT: �N�G�����X�g
'====================================================================================================
Private Function MakeQueries(xEntryData As EntryData) As Collection
    Select Case xEntryData.EntryType
    Case EntryType.Register
        Set MakeQueries = MakeInsertQueries(xEntryData)
    Case EntryType.Update
        Set MakeQueries = MakeUpdateQueries(xEntryData)
    Case EntryType.Remove
        Set MakeQueries = MakeDeleteQueries(xEntryData)
    End Select
End Function


'====================================================================================================
' Insert�N�G�����X�g�𐶐����܂�
'----------------------------------------------------------------------------------------------------
' IN : xEntryData �����f�[�^
' OUT: �N�G�����X�g
'====================================================================================================
Private Function MakeInsertQueries(xEntryData As EntryData) As Collection
    Dim rr As Range
    Dim query As String
    Dim queries As Collection

    Set queries = New Collection
    For Each rr In xEntryData.RecordRange.Rows
        query = cstInsertQuery
        query = Replace(query, "${tableName}", xEntryData.TableName)
        query = Replace(query, "${columns}", GetColumnPhrase(xEntryData.ColumnDefinitions))
        query = Replace(query, "${values}", GetValuePhrase(xEntryData.ColumnDefinitions, rr))
        Call queries.Add(query)
    Next
    Set MakeInsertQueries = queries
End Function


'====================================================================================================
' Update�N�G�����X�g�𐶐����܂�
'----------------------------------------------------------------------------------------------------
' IN : xEntryData �����f�[�^
' OUT: �N�G�����X�g
'====================================================================================================
Private Function MakeUpdateQueries(xEntryData As EntryData) As Collection
    Dim rr As Range
    Dim query As String
    Dim queries As Collection

    Set queries = New Collection
    For Each rr In xEntryData.RecordRange.Rows
        query = cstUpdateQuery
        query = Replace(query, "${tableName}", xEntryData.TableName)
        query = Replace(query, "${set}", GetSetPhrase(xEntryData.ColumnDefinitions, rr))
        query = Replace(query, "${where}", GetWherePhrase(xEntryData.ColumnDefinitions, rr))
        Call queries.Add(query)
    Next
    Set MakeUpdateQueries = queries
End Function


'====================================================================================================
' Delete�N�G�����X�g�𐶐����܂�
'----------------------------------------------------------------------------------------------------
' IN : xEntryData �����f�[�^
' OUT: �N�G�����X�g
'====================================================================================================
Private Function MakeDeleteQueries(xEntryData As EntryData) As Collection
    Dim rr As Range
    Dim query As String
    Dim queries As Collection

    Set queries = New Collection
    For Each rr In xEntryData.RecordRange.Rows
        query = cstDeleteQuery
        query = Replace(query, "${tableName}", xEntryData.TableName)
        query = Replace(query, "${where}", GetWherePhrase(xEntryData.ColumnDefinitions, rr))
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
        result = result & GetItemValue(record.Cells(1, i), cd(i).DataType) & cstDelimiter
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
            result = Replace(result, "${value}", GetItemValue(record.Cells(1, i), cd(i).DataType))
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
            result = Replace(result, "${value}", GetItemValue(record.Cells(1, i), cd(i).DataType))
            result = result & cstWherePhraseDelimiter
        End If
    Next
    GetWherePhrase = Left(result, Len(result) - Len(cstWherePhraseDelimiter))
End Function


'====================================================================================================
' �f�[�^�^�ɑΉ������f�[�^�l�̍��ڒl���擾���܂�
'----------------------------------------------------------------------------------------------------
' IN : dataValue �f�[�^�l
'    : xDataType �f�[�^�^
' OUT: Value��
'====================================================================================================
Private Function GetItemValue(ByVal dataValue As String, ByVal xDataType As String) As String
    Dim itemValue As String

    itemValue = dataValue
    If itemValue = "" Then
        itemValue = "NULL"
    ' ������
    ElseIf mDatabaseModel.IsDataTypeString(xDataType) Then
        itemValue = Replace(itemValue, "'", "''")                                   ' �P����p���G�X�P�[�v
        itemValue = Replace(itemValue, vbLf, "'" & Setting.LinefeedCode & "'")      ' ���s�R�[�h�ϊ�
        itemValue = "'" & itemValue & "'"                                           ' �P����p���t�^
    ' ���t
    ElseIf mDatabaseModel.IsDataTypeDate(xDataType) Then
        itemValue = "TO_DATE('" & itemValue & "','" & Setting.DateFormat & "')"
    ' �^�C���X�^���v
    ElseIf mDatabaseModel.IsDataTypeTimestamp(xDataType) Then
        itemValue = "TO_TIMESTAMP('" & itemValue & "','" & Setting.TimestampFormat & "')"
    End If
    GetItemValue = itemValue
End Function
