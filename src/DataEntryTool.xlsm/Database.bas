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
' OUT: ��������
'====================================================================================================
Public Function ExecuteQuery(query As String)
    Dim procCnt As Long
    Call mConnect.Execute(CommandText:=query, RecordsAffected:=procCnt)
    ExecuteQuery = procCnt
End Function


'====================================================================================================
' �e�[�u�������擾���܂�
'----------------------------------------------------------------------------------------------------
' OUT: �e�[�u����
'====================================================================================================
Public Function GetTableName()
    Set GetTableName = mConnect.Execute(mDatabaseCore.GetTableNameQuery)
End Function


'====================================================================================================
' �J������`���擾���܂�
'----------------------------------------------------------------------------------------------------
' IN : tableSettings �e�[�u���ݒ�̘A���z��
' OUT: �J������`
'====================================================================================================
Public Function GetColumnDefinition(tableSettings As Object)
    Set GetColumnDefinition = mConnect.Execute(mDatabaseCore.GetColumnDefinitionQuery(tableSettings))
End Function


'====================================================================================================
' �f�[�^�^�ɑΉ������f�[�^�l�̍��ڒl���擾���܂�
'----------------------------------------------------------------------------------------------------
' IN : dataValue �f�[�^�l
'    : xDataType �f�[�^�^
' OUT: Value��
'====================================================================================================
Public Function GetItemValue(ByVal dataValue As String, ByVal xDataType As String) As String
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


'Private Sub GetRecordsetData()
'   iExecutionCount = iExecutionCount + 1
'   If rst.State <> adStateClosed Then
'      rst.Close
'   End If
'   rst.Open _
'      "Select * From Pubs..Publishers, Pubs..Titles, Pubs..Authors", _
'      con, adOpenKeyset, adLockOptimistic, adAsyncExecute
'End Sub
'
'Private Sub con_ExecuteComplete(ByVal RecordsAffected As Long, _
'    ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, _
'    ByVal pCommand As ADODB.Command, _
'    ByVal pRecordset As ADODB.Recordset, _
'    ByVal pConnection As ADODB.Connection)
'
'   ' When the ADO recordset has been populated with data, begin opening
'   ' the next ADO recordset.
'   GetRecordsetData
'End Sub
