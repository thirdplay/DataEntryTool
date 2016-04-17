VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataEntryDao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' �f�[�^����Dao
'
'====================================================================================================

'====================================================================================================
' �����o�ϐ�
'====================================================================================================
Private mDatabaseCore As IDatabaseCore                  ' �f�[�^�x�[�X�R�A�C���^�[�t�F�[�X
Private WithEvents mConnection As ADODB.Connection      ' �ڑ��I�u�W�F�N�g
Attribute mConnection.VB_VarHelpID = -1
Private mAsyncRecordset As ADODB.Recordset              ' �񓯊������̃��R�[�h�Z�b�g
Private mAsyncRecordsAffected As Long                   ' �񓯊������̍X�V����
Private mIsAsyncProc As Boolean                         ' �񓯊������t���O


'====================================================================================================
' �R���X�g���N�^
'====================================================================================================
Public Sub Class_Initialize()
End Sub


'====================================================================================================
' �f�X�g���N�^
'====================================================================================================
Public Sub Class_Terminate()
    Call Me.Disconnect
End Sub


'====================================================================================================
' �f�[�^�x�[�X�ڑ�
'====================================================================================================
Public Sub Connect()
On Error GoTo ErrHandler
    If mConnection Is Nothing Then
        ' �f�[�^�x�[�X�R�A�̐���
        If Setting.DatabaseType = cstDatabaseTypeOracle Then
            Set mDatabaseCore = New DatabaseCoreOracle
        ElseIf Setting.DatabaseType = cstDatabaseTypePostgreSQL Then
            Set mDatabaseCore = New DatabaseCorePostgreSql
        End If

        ' �f�[�^�x�[�X�ڑ�
        Set mConnection = New ADODB.Connection
        Call mConnection.Open(mDatabaseCore.GetConnectStr(Setting.ServerName, Setting.UserId, Setting.Password, Setting.Port, Setting.DatabaseName))
        mConnection.CursorLocation = adUseClient
    End If
    Exit Sub
ErrHandler:
    Set mConnection = Nothing
    Err.Raise Err.Number, Err.Source, "�f�[�^�x�[�X�ڑ����s" & vbNewLine & Err.Description, Err.HelpFile, Err.HelpContext
    Exit Sub
End Sub


'====================================================================================================
' �f�[�^�x�[�X�ؒf
'====================================================================================================
Public Sub Disconnect()
    If Not mConnection Is Nothing Then
        mConnection.Close
        Set mConnection = Nothing
        Set mDatabaseCore = Nothing
    End If
End Sub


'====================================================================================================
' �g�����U�N�V�����J�n
'====================================================================================================
Public Sub BeginTrans()
    Call mConnection.BeginTrans
End Sub


'====================================================================================================
' �R�~�b�g
'====================================================================================================
Public Sub CommitTrans()
    Call mConnection.CommitTrans
End Sub


'====================================================================================================
' ���[���o�b�N
'====================================================================================================
Public Sub RollbackTrans()
    Call mConnection.RollbackTrans
End Sub


'====================================================================================================
' �X�V�N�G�������s���܂�
'----------------------------------------------------------------------------------------------------
' IN : query �N�G��������
' OUT: ��������
'====================================================================================================
Public Function ExecuteUpdate(query As String) As Long
    Dim procCnt As Long
    Call mConnection.Execute(CommandText:=query, RecordsAffected:=procCnt)
    ExecuteUpdate = procCnt
End Function


'====================================================================================================
' �e�[�u�������擾���܂�
'----------------------------------------------------------------------------------------------------
' OUT: �e�[�u����
'====================================================================================================
Public Function GetTableName() As ADODB.Recordset
    Set GetTableName = mConnection.Execute(mDatabaseCore.GetTableNameQuery)
End Function


'====================================================================================================
' �J������`���擾���܂�
'----------------------------------------------------------------------------------------------------
' IN : tableSettings �e�[�u���ݒ�̘A���z��
' OUT: �J������`
'====================================================================================================
Public Function GetColumnDefinition(tableSettings As Object) As ADODB.Recordset
    Call ExecuteAsync(mDatabaseCore.GetColumnDefinitionQuery(tableSettings))
    Call WaitAsync
    Set GetColumnDefinition = mAsyncRecordset
End Function


'====================================================================================================
' �f�[�^�^�ɑΉ������f�[�^�l�̍��ڒl���擾���܂�
'----------------------------------------------------------------------------------------------------
' IN : dataValue �f�[�^�l
'    : xDataType �f�[�^�^
' OUT: ���ڒl
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


'====================================================================================================
' �N�G�����s(�񓯊�)
'----------------------------------------------------------------------------------------------------
' IN : query �N�G��������
'====================================================================================================
Private Sub ExecuteAsync(query As String)
    mIsAsyncProc = True
    Call mConnection.Execute(CommandText:=query, Options:=adAsyncExecute)
End Sub


'====================================================================================================
' �񓯊��������I������܂ő҂�
'====================================================================================================
Private Sub WaitAsync()
    Do While mIsAsyncProc
        Call ApplicationEx.CheckEvents
    Loop
End Sub


'====================================================================================================
' �R�l�N�V�����̎��s�����C�x���g
'----------------------------------------------------------------------------------------------------
' IN : RecordsAffected �X�V����
'    : pError ADODB�G���[
'    : adStatus �C�x���g�X�e�[�^�X
'    : pCommand �R�}���h
'    : pRecordset ���R�[�h�Z�b�g
'    : pConnection �R�l�N�V����
'====================================================================================================
Private Sub mConnection_ExecuteComplete( _
    ByVal RecordsAffected As Long, _
    ByVal pError As ADODB.Error, _
    ByRef adStatus As ADODB.EventStatusEnum, _
    ByVal pCommand As ADODB.Command, _
    ByVal pRecordset As ADODB.Recordset, _
    ByVal pConnection As ADODB.Connection _
)
    Set mAsyncRecordset = pRecordset
    mAsyncRecordsAffected = RecordsAffected
    mIsAsyncProc = False
End Sub