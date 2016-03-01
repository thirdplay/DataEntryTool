VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataEntryView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' �f�[�^�����r���[
'
'====================================================================================================

'====================================================================================================
' �萔
'====================================================================================================
' ���O��`
Private Const cstServerName = "ServerName"              ' �T�[�o��
Private Const cstPort = "Port"                          ' �|�[�g
Private Const cstDatabaseName = "DatabaseName"          ' �f�[�^�x�[�X��
Private Const cstUserId = "UserId"                      ' ���[�UID
Private Const cstPassword = "Password"                  ' �p�X���[�h
Private Const cstTableBase = "TableBase"                ' �e�[�u���ꗗ�̊�Z��
' �V�[�g��
Private Const cstSheetMain = "�f�[�^�����c�[��"         ' ���C���V�[�g
Private Const cstSheetTemplate = "�e���v���[�g"         ' �e���v���[�g�V�[�g

' �e�[�u���ꗗ�̎Q�ƒl
Private Enum TableListIndex
    PhysicsName = 1
    LogicName
    Target
    Result
End Enum


'====================================================================================================
' �����o�ϐ�
'====================================================================================================
Private mDatabaseInfo As DatabaseInfo           ' �f�[�^�x�[�X���


'====================================================================================================
' �R���X�g���N�^
'====================================================================================================
Private Sub Class_Initialize()
    Set mDatabaseInfo = New DatabaseInfo
    With mDatabaseInfo
        .ServerName = ThisWorkbook.Worksheets(cstSheetMain).Range(cstServerName).Value
        .Port = ThisWorkbook.Worksheets(cstSheetMain).Range(cstPort).Value
        .DatabaseName = ThisWorkbook.Worksheets(cstSheetMain).Range(cstDatabaseName).Value
        .UserId = ThisWorkbook.Worksheets(cstSheetMain).Range(cstUserId).Value
        .Password = ThisWorkbook.Worksheets(cstSheetMain).Range(cstPassword).Value
    End With
End Sub


'====================================================================================================
' �f�X�g���N�^
'====================================================================================================
Private Sub Class_Terminate()
    Set mDatabaseInfo = Nothing
End Sub


'====================================================================================================
' �f�[�^�x�[�X����ԋp���܂�
'----------------------------------------------------------------------------------------------------
' OUT:�f�[�^�x�[�X���
'====================================================================================================
Public Function GetDatabaseInfo() As DatabaseInfo
    Set GetDatabaseInfo = mDatabaseInfo
End Function


'====================================================================================================
' �e�[�u�����ꗗ��ԋp���܂�
'----------------------------------------------------------------------------------------------------
' OUT:�e�[�u�����ꗗ
'====================================================================================================
Public Function GetTableNameList() As Collection
    Dim list As Collection
    Dim index As Long
    
    With ThisWorkbook.Worksheets(cstSheetMain)
        Set list = New Collection
        index = .Range(cstTableBase).Row + 1
        
        ' �Ώۗ�ɒl���ݒ肳��Ă���e�[�u������ǉ�����
        Do While .Cells(index, TableListIndex.PhysicsName).Value <> ""
            If .Cells(index, TableListIndex.Target).Value <> "" Then
                list.Add (.Cells(index, TableListIndex.PhysicsName))
            End If
            index = index + 1
        Loop
    End With
    Set GetTableNameList = list
    
End Function

'�e�[�u���f�[�^�ꗗ�擾

'�e�[�u���V�[�g�̍č쐬


