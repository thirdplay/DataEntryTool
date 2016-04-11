Attribute VB_Name = "Database"
Option Explicit
Option Private Module

'====================================================================================================
'
' �f�[�^�x�[�X���W���[��
'
'====================================================================================================

'====================================================================================================
' �萔
'====================================================================================================
' �J������`�擾�N�G���̋�؂蕶��
Private Const cstColumnDefinitionDelimiter = " UNION ALL "
' �J������`�擾�N�G���̐ړ���
Private Const cstColumnDefinitionPrefix = "SELECT * FROM ("
' �J������`�擾�N�G���̐ڔ���
Private Const cstColumnDefinitionSuffix = ") tbl ORDER BY table_name, column_id"
' ��؂蕶��
Private Const cstDelimiter = ","


'====================================================================================================
' �����o�ϐ�
'====================================================================================================
Private mDatabaseCore As IDatabaseCore      ' �f�[�^�x�[�X�R�A�C���^�[�t�F�[�X
Private mConnect As Object                  ' �ڑ��I�u�W�F�N�g


'====================================================================================================
' �f�[�^�x�[�X�ڑ�
'====================================================================================================
Public Sub Connect()
On Error GoTo ErrHandler
    If mConnect Is Nothing Then
        ' �f�[�^�x�[�X�R�A�̐���
        If Setting.DatabaseType = cstDatabaseTypeOracle Then
            Set mDatabaseCore = New DatabaseCoreOracle
        ElseIf Setting.DatabaseType = cstDatabaseTypePostgreSQL Then
            Set mDatabaseCore = New DatabaseCorePostgreSql
        End If

        ' �f�[�^�x�[�X�ڑ�
        Set mConnect = CreateObject("ADODB.Connection")
        mConnect.Open mDatabaseCore.GetConnectStr
        mConnect.CursorLocation = adUseClient
    End If
    Exit Sub
ErrHandler:
    Set mConnect = Nothing
    Err.Raise Err.Number, Err.Source, "�f�[�^�x�[�X�ڑ����s" & vbNewLine & Err.Description, Err.HelpFile, Err.HelpContext
    Exit Sub
End Sub


'====================================================================================================
' �f�[�^�x�[�X�ؒf
'====================================================================================================
Public Sub Disconnect()
    If Not mConnect Is Nothing Then
        mConnect.Close
        Set mConnect = Nothing
        Set mDatabaseCore = Nothing
    End If
End Sub


'====================================================================================================
' �g�����U�N�V�����J�n
'====================================================================================================
Public Sub BeginTrans()
    Call mConnect.BeginTrans
End Sub


'====================================================================================================
' �R�~�b�g
'====================================================================================================
Public Sub CommitTrans()
    Call mConnect.CommitTrans
End Sub


'====================================================================================================
' ���[���o�b�N
'====================================================================================================
Public Sub RollbackTrans()
    Call mConnect.RollbackTrans
End Sub


'====================================================================================================
' �N�G�������s���܂�
'----------------------------------------------------------------------------------------------------
' IN : query �N�G��������
'====================================================================================================
Public Sub ExecuteQuery(query As String)
    Call mConnect.Execute(CommandText:=query)
End Sub


'====================================================================================================
' �w�肳�ꂽ�e�[�u���̃J������`���X�g���擾���܂�
'----------------------------------------------------------------------------------------------------
' IN : tableSettings �e�[�u���ݒ�A�z�z��
' OUT: �J������`���X�g
'====================================================================================================
Public Function GetColumnDefinitions(tableSettings As Object) As Object
    Dim rs As Object
    Dim tableNames As Object
    Dim query As String
    Dim xTableName As Variant
    Dim td As TableDefinition
    Dim cd As ColumnDefinition
    Dim dic As Object

    ' �e�[�u�����A�z�z��̎擾
    Set rs = mConnect.Execute(mDatabaseCore.GetTableNameQuery)
    Set tableNames = CreateObject("Scripting.Dictionary")
    Do Until rs.EOF
        Call tableNames.Add(rs("table_name").Value, rs("table_name").Value)
        rs.MoveNext
    Loop

    ' �N�G�������Ȃ���e�[�u�����̑��݃`�F�b�N������
    query = cstColumnDefinitionPrefix
    For Each xTableName In tableSettings.Keys
        query = query & Replace(mDatabaseCore.GetColumnDefinitionQuery, "${tableName}", xTableName) & cstColumnDefinitionDelimiter
        If Not tableNames.Exists(xTableName) Then
            Err.Raise 1000, , "�e�[�u��[" & xTableName & "]�̃J������`���擾�ł��܂���B"
        End If
    Next
    query = Left(query, Len(query) - Len(cstColumnDefinitionDelimiter)) & cstColumnDefinitionSuffix
    Debug.Print query

    ' �J������`�擾�N�G���̎��s
    Set rs = mConnect.Execute(query)

    ' �e�[�u����`�̘A���z����쐬����
    Set dic = CreateObject("Scripting.Dictionary")
    Do Until rs.EOF
        ' �e�[�u����`����
        If Not dic.Exists(rs("table_name").Value) Then
            Set td = New TableDefinition
            td.TableName = rs("table_name").Value
            td.ColumnDefinitions = New Collection
            Call dic.Add(td.TableName, td)
        End If

        ' �J������`����
        Set cd = New ColumnDefinition
        With cd
            .ColumnId = rs("column_id").Value
            .ColumnName = rs("column_name").Value
            .Comments = rs("comments").Value
            .DataType = rs("data_type").Value
            .DataLength = rs("data_length").Value
            .IsRequired = rs("is_required").Value
            .IsPrimaryKey = rs("is_primary_key").Value
        End With
        Call dic(rs("table_name").Value).ColumnDefinitions.Add(cd)
        rs.MoveNext
    Loop
    Set GetColumnDefinitions = dic
End Function


'====================================================================================================
' �f�[�^�����N�G���𐶐����܂�
'----------------------------------------------------------------------------------------------------
' IN : xEntryData �����f�[�^
' OUT: �N�G��������
'====================================================================================================
Public Function MakeDataEntryQuery(xEntryData As EntryData) As String
    Dim rr As Range
    Dim query As String

    query = mDatabaseCore.GetDataEntryQueryPrefix
    For Each rr In xEntryData.RecordRange.Rows
        query = query & mDatabaseCore.GetDataEntryQuery
        query = Replace(query, "${tableName}", xEntryData.TableName)
        query = Replace(query, "${columns}", GetColumnPhrase(xEntryData.ColumnDefinitions))
        query = Replace(query, "${values}", GetValuePhrase(xEntryData.ColumnDefinitions, rr))
    Next
    query = query & mDatabaseCore.GetDataEntryQuerySuffix
    MakeDataEntryQuery = query
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
    ElseIf mDatabaseCore.IsDataTypeString(xDataType) Then
        itemValue = Replace(itemValue, "'", "''")                                   ' �P����p���G�X�P�[�v
        itemValue = Replace(itemValue, vbLf, "'" & Setting.LinefeedCode & "'")      ' ���s�R�[�h�ϊ�
        itemValue = "'" & itemValue & "'"                                           ' �P����p���t�^
    ' ���t
    ElseIf mDatabaseCore.IsDataTypeDate(xDataType) Then
        itemValue = "TO_DATE('" & itemValue & "','" & Setting.DateFormat & "')"
    ' �^�C���X�^���v
    ElseIf mDatabaseCore.IsDataTypeTimestamp(xDataType) Then
        itemValue = "TO_TIMESTAMP('" & itemValue & "','" & Setting.TimestampFormat & "')"
    End If
    GetItemValue = itemValue
End Function
