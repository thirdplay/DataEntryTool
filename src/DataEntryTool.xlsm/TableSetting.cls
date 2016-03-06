VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' �e�[�u���ݒ�
'
'====================================================================================================

'====================================================================================================
' �����o�ϐ�
'====================================================================================================
Private mRow As String                      ' �s�ԍ�
Private mPhysicsName As String              ' ������
Private mLogicalName As String              ' �_����
Private mDataEntryTarget As String          ' �f�[�^�����Ώ�
Private mProcessCount As String             ' ��������


'====================================================================================================
' �s�ԍ��̎擾/�ݒ�
'====================================================================================================
Public Property Get Row() As String
    Row = mRow
End Property
Public Property Let Row(Row As String)
    mRow = Row
End Property


'====================================================================================================
' �������̎擾/�ݒ�
'====================================================================================================
Public Property Get PhysicsName() As String
    PhysicsName = mPhysicsName
End Property
Public Property Let PhysicsName(PhysicsName As String)
    mPhysicsName = PhysicsName
End Property


'====================================================================================================
' �_�����̎擾/�ݒ�
'====================================================================================================
Public Property Get LogicalName() As String
    LogicalName = mLogicalName
End Property
Public Property Let LogicalName(LogicalName As String)
    mLogicalName = LogicalName
End Property


'====================================================================================================
' �f�[�^�����Ώۂ̎擾/�ݒ�
'====================================================================================================
Public Property Get DataEntryTarget() As String
    DataEntryTarget = mDataEntryTarget
End Property
Public Property Let DataEntryTarget(DataEntryTarget As String)
    mDataEntryTarget = DataEntryTarget
End Property


'====================================================================================================
' ���������̎擾/�ݒ�
'====================================================================================================
Public Property Get ProcessCount() As String
    ProcessCount = mProcessCount
End Property
Public Property Let ProcessCount(ProcessCount As String)
    mProcessCount = ProcessCount
End Property