Attribute VB_Name = "Database"
Option Explicit
Option Private Module

'====================================================================================================
'
' �f�[�^�x�[�X���W���[��
'
'====================================================================================================

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
' OUT: ��������
'====================================================================================================
Public Function ExecuteQuery(query As String) As Long
    Dim procCnt As Long
    Dim rs As Object

    Set rs = mConnect.Execute(CommandText:=query)
    Do Until rs Is Nothing
        procCnt = procCnt + rs.RecordCount
        Set rs = rs.NextRecordset
    Loop
    ExecuteQuery = procCnt
End Function


'====================================================================================================
' �w�肳�ꂽ�e�[�u���̃J������`���X�g���擾���܂�
'----------------------------------------------------------------------------------------------------
' IN : xTableName �e�[�u����
' OUT: �J������`���X�g
'====================================================================================================
Public Function GetColumnDefinitions(xTableName As String) As Collection
    Dim recordSet As Object
    Dim query As String
    Dim cd As ColumnDefinition
    Dim list As Collection
    Set list = New Collection

    query = Replace(mDatabaseCore.GetColumnDefinitionQuery, "${tableName}", xTableName)
    Set recordSet = mConnect.Execute(query)
    Do Until recordSet.EOF
        Set cd = New ColumnDefinition
        With cd
            .ColumnId = recordSet("column_id")
            .ColumnName = recordSet("column_name")
            .Comments = recordSet("comments")
            .DataType = recordSet("data_type")
            .DataLength = recordSet("data_length")
            .IsRequired = recordSet("is_required")
            .IsPrimaryKey = recordSet("is_primary_key")
        End With
        Call list.Add(cd)
        recordSet.MoveNext
    Loop
    Set GetColumnDefinitions = list
End Function


'====================================================================================================
' �f�[�^�^�������񂩂ǂ������肷��
'----------------------------------------------------------------------------------------------------
' IN : xDataType �f�[�^�^
' OUT: True:������AFalse:������ȊO
'====================================================================================================
Public Function IsDataTypeString(ByVal xDataType As String) As Boolean
    IsDataTypeString = mDatabaseCore.IsDataTypeString(xDataType)
End Function


'====================================================================================================
' �f�[�^�^�����t���ǂ������肷��
'----------------------------------------------------------------------------------------------------
' IN : xDataType �f�[�^�^
' OUT: True:���t�AFalse:���t�ȊO
'====================================================================================================
Public Function IsDataTypeDate(ByVal xDataType As String) As Boolean
    IsDataTypeDate = mDatabaseCore.IsDataTypeDate(xDataType)
End Function


'====================================================================================================
' �f�[�^�^���^�C���X�^���v���ǂ������肷��
'----------------------------------------------------------------------------------------------------
' IN : xDataType �f�[�^�^
' OUT: True:�^�C���X�^���v�AFalse:�^�C���X�^���v�ȊO
'====================================================================================================
Public Function IsDataTypeTimestamp(ByVal xDataType As String) As Boolean
    IsDataTypeTimestamp = mDatabaseCore.IsDataTypeTimestamp(xDataType)
End Function


'====================================================================================================
' �N�G���̐ڔ������擾���܂�
'----------------------------------------------------------------------------------------------------
' OUT: �ڔ���
'====================================================================================================
Public Function GetQuerySuffix() As String
    GetQuerySuffix = mDatabaseCore.GetQuerySuffix()
End Function


