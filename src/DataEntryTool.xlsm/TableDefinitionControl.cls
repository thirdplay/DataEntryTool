VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableDefinitionControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' �e�[�u����`�R���g���[��
'
'====================================================================================================

'====================================================================================================
' �����o�ϐ�
'====================================================================================================
Private mView As DataEntryView              ' �f�[�^�����r���[
Private mModel As TableDefinitionModel      ' �e�[�u����`���f��


'====================================================================================================
' �R���X�g���N�^
'====================================================================================================
Private Sub Class_Initialize()
End Sub


'====================================================================================================
' �f�X�g���N�^
'====================================================================================================
Private Sub Class_Terminate()
    Set mView = Nothing
    Set mModel = Nothing
End Sub


'====================================================================================================
' �r���[�̐ݒ�
'====================================================================================================
Public Property Let View(View As DataEntryView)
    Set mView = View
End Property


'====================================================================================================
' ���f���̐ݒ�
'====================================================================================================
Public Property Let Model(Model As TableDefinitionModel)
    Set mModel = Model
End Property


'====================================================================================================
' ���s
'====================================================================================================
Public Sub Execute()
    Dim dbInfo As DatabaseInfo
    Dim tableNames As Collection
    
    ' DB���̎擾
    Set dbInfo = mView.GetDatabaseInfo
    ' �e�[�u�����ꗗ�̎擾
    Set tableNames = mView.GetTableNameList

    ' �e�[�u�����ꗗ�����ɁAModel����e�[�u�����ꗗ���擾
    'Set tableInfoList = mModel.GetTableInfoList(dbInfo)

    ' 3. �e�[�u�����ꗗ�����ɁA�e�[�u���V�[�g���č쐬
    'Call mView.RecreateTableDefinition tableInfoList
    
    Set dbInfo = Nothing
End Sub
