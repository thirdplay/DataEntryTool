VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataEntryControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' �f�[�^�����R���g���[��
'
'====================================================================================================

'====================================================================================================
' �����o�ϐ�
'====================================================================================================
Private mView As DataEntryView          ' �f�[�^�����r���[
Private mModel As DataEntryModel        ' �f�[�^�������f��


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
Public Property Let Model(Model As DataEntryModel)
    Set mModel = Model
End Property


'====================================================================================================
' �e�[�u����`�č쐬
'====================================================================================================
Public Sub RecreateTableDefinition()
On Error GoTo Finally
    Dim setting As Setting
    Dim tableNames As Collection
    Dim tableInfoList As Collection
    
    ' �ݒ�ƃe�[�u�����ꗗ�̎擾
    Set setting = mView.GetSetting
    Set tableNames = mView.GetTableNameList

    ' �e�[�u�����ꗗ�����ɁAModel����e�[�u�����ꗗ���擾
    mModel.Setting = setting
    Set tableInfoList = mModel.GetTableInfoList(tableNames)

    ' �e�[�u�����ꗗ�����ɁA�e�[�u����`�V�[�g���č쐬����
    Call mView.RecreateTableDefinition(tableInfoList)
Finally:
    Set setting = Nothing
    Set tableNames = Nothing
    Set tableInfoList = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Sub
