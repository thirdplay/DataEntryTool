VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataEntryModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' �f�[�^�������f��
'
'====================================================================================================

'====================================================================================================
' �����o�ϐ�
'====================================================================================================
Private mSetting As Setting     ' �f�[�^�x�[�X���


'====================================================================================================
' �R���X�g���N�^
'====================================================================================================
Private Sub Class_Initialize()
End Sub


'====================================================================================================
' �f�X�g���N�^
'====================================================================================================
Private Sub Class_Terminate()
    Set mSetting = Nothing
End Sub


'====================================================================================================
' �ݒ�̐ݒ�
'====================================================================================================
Public Property Let Setting(setting As Setting)
    Set mSetting = setting
End Property


'====================================================================================================
' �e�[�u�����ꗗ�̃e�[�u������DB����擾���A�ԋp���܂�
'----------------------------------------------------------------------------------------------------
' IN : tableNames �e�[�u�������X�g
' OUT: �e�[�u����񃊃X�g
'====================================================================================================
Public Function GetTableInfoList(tableNames As Collection) As Collection
On Error GoTo Finally
    Dim i As Long
    Dim dao As IDataEntryDao
    Dim ti AS TableInfo
    Dim tiList As Collection

    Set tiList = New Collection
    Set dao = DaoFactory.Create(mSetting.DatabaseType)

    Call dao.Connect(mSetting)
    For i = 1 To tableNames.Count
        Set ti = New TableInfo
        ti.ColumnInfoList = dao.GetColumnInfoList(tableNames(i))
        ti.TableName = tableNames(i)
        Call tiList.Add(ti)
    Next
    Call dao.Disconnect

    Set GetTableInfoList = tiList
Finally:
    Set dao = Nothing
    Set ti = Nothing
    Set tiList = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Function
