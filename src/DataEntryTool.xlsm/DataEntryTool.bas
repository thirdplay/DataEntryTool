Attribute VB_Name = "DataEntryTool"
Option Explicit
'Option Private Module

'====================================================================================================
'
' �f�[�^�����c�[���̃��W���[��
'
'====================================================================================================

'====================================================================================================
' Excel�t�@�C���I�[�v���C�x���g
'====================================================================================================
Private Sub Auto_Open()
On Error GoTo ErrHandler
    ' �Q�Ɛݒ�
    Call ThisWorkbook.VBProject.References.AddFromGuid(cstGuidAdodb, 6, 1)
    Call ThisWorkbook.VBProject.References.AddFromGuid(cstGuidScripting, 1, 0)
    Exit Sub
ErrHandler:
    If Err.Number = ErrNumber.AlreadyReferenceConfigured Then
        Resume Next
    Else
        MsgBox "Error Number : " & Err.Number & vbNewLine & Err.Description
    End If
End Sub


'====================================================================================================
' �e�[�u���V�[�g�쐬
'====================================================================================================
Public Sub CreateTableSheet()
    Dim controller As TableSheetController

    Set controller = New TableSheetController
    Call controller.Execute
End Sub

'====================================================================================================
' �f�[�^�o�^
'====================================================================================================
Public Sub RegisterData()
    Call ExecuteDataEntry(EntryType.Register)
End Sub


'====================================================================================================
' �f�[�^�X�V
'====================================================================================================
Public Sub UpdateData()
    Call ExecuteDataEntry(EntryType.Update)
End Sub


'====================================================================================================
' �f�[�^�폜
'====================================================================================================
Public Sub DeleteData()
    Call ExecuteDataEntry(EntryType.Delete)
End Sub


'====================================================================================================
' �f�[�^�����̎��s
'----------------------------------------------------------------------------------------------------
' IN : xEntryType �������
'====================================================================================================
Private Sub ExecuteDataEntry(xEntryType As EntryType)
    Dim controller As DataEntryController

    Set controller = New DataEntryController
    Call controller.Execute(xEntryType)
End Sub
