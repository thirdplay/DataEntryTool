Attribute VB_Name = "WorkBookEx"
Option Explicit
Option Private Module

'====================================================================================================
'
' ���[�N�u�b�N�̊g�����W���[��
'
'====================================================================================================

'====================================================================================================
' �V�[�g�����݂��邩���肵�܂�
'----------------------------------------------------------------------------------------------------
' IN : sheetName �V�[�g��
' OUT: ���݂���ꍇ��true�A����ȊO��false
'====================================================================================================
Public Function ExistsSheet(sheetName As String) As Boolean
    Dim sheet As Worksheet
    For Each sheet In ThisWorkbook.Worksheets
        If sheet.Name = sheetName Then
            ExistsSheet = True
            Exit Function
        End If
    Next
    ExistsSheet = False
End Function


