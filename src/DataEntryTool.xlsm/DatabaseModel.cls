VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatabaseModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' �f�[�^�x�[�X���f��
'
'====================================================================================================

'====================================================================================================
' �����o�ϐ�
'====================================================================================================
Dim mDatabaseCore As IDatabaseCore      ' �f�[�^�x�[�X�R�A�C���^�[�t�F�[�X
Dim mConnect As Object                  ' �ڑ��I�u�W�F�N�g


'====================================================================================================
' �R���X�g���N�^
'====================================================================================================
Private Sub Class_Initialize()
End Sub


'====================================================================================================
' �f�X�g���N�^
'====================================================================================================
Private Sub Class_Terminate()
    Call Disconnect
End Sub


'====================================================================================================
' DB�ڑ�
'----------------------------------------------------------------------------------------------------
' IN : xSetting �ݒ�
'====================================================================================================
Public Sub Connect(xSetting As Setting)
On Error GoTo ErrHandler
    If mConnect Is Nothing Then
        ' �f�[�^�x�[�X�R�A�̐���
        If xSetting.DatabaseType = cstDatabaseTypeOracle Then
            Set mDatabaseCore = New DatabaseCoreOracle
        ElseIf xSetting.DatabaseType = cstDatabaseTypePostgreSQL Then
            Set mDatabaseCore = New DatabaseCorePostgreSql
        End If

        ' �f�[�^�x�[�X�ڑ�
        Set mConnect = CreateObject("ADODB.Connection")
        mConnect.Open mDatabaseCore.GetConnectStr(xSetting)
        Exit Sub
    End If
ErrHandler:
    Set mConnect = Nothing
    Err.Raise Err.Number, Err.Source, "DB�ڑ����s" & vbNewLine & Err.Description, Err.HelpFile, Err.HelpContext
    Exit Sub
End Sub


'====================================================================================================
' DB�ؒf
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
    mConnect.BeginTrans
End Sub


'====================================================================================================
' �R�~�b�g
'====================================================================================================
Public Sub CommitTrans()
    mConnect.CommitTrans
End Sub


'====================================================================================================
' ���[���o�b�N
'====================================================================================================
Public Sub RollbackTrans()
    mConnect.RollbackTrans
End Sub


'====================================================================================================
' �N�G�����s
'----------------------------------------------------------------------------------------------------
' IN : query �N�G��������
'====================================================================================================
Public Sub Execute(query)
    Debug.Print "Execute"
End Sub


'====================================================================================================
' �w�肳�ꂽ�e�[�u���̃J������`���X�g���擾���܂�
'----------------------------------------------------------------------------------------------------
' IN : tableName �e�[�u����
' OUT: �J������`���X�g
'====================================================================================================
Public Function GetColumnDefinitions(TableName As String) As Collection
    Dim recordSet As Object
    Dim query As String
    Dim cd As ColumnDefinition
    Dim list As Collection
    Set list = New Collection

    query = Replace(mDatabaseCore.GetColumnDefinitionQuery, "${tableName}", TableName)
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

