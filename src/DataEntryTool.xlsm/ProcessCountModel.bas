Attribute VB_Name = "ProcessCountModel"
Option Explicit
Option Private Module

'====================================================================================================
'
' �����������f���̃��W���[��
'
'====================================================================================================

'====================================================================================================
' �����������N���A���܂�
'====================================================================================================
Public Sub ClearProcessingCount()
    Dim baseRowIndex As Long
    Dim dataCount As Long

    With ThisWorkbook.Worksheets(cstSheetMain)
        ' �e�[�u���ݒ萔�̎擾
        baseRowIndex = .Range(cstTableBase).Row + 1
        dataCount = WorkSheetEx.GetRowDataCount(cstSheetMain, baseRowIndex, TableSettingCol.PhysicsName)

        ' ���������̃N���A
        Call .Range(.Cells(baseRowIndex, TableSettingCol.ProcessCount), .Cells(baseRowIndex + dataCount, TableSettingCol.ProcessCount)).ClearContents
    End With
End Sub


'====================================================================================================
' ������������������
'----------------------------------------------------------------------------------------------------
' IN : xTableSetting �e�[�u���ݒ�
'    : procCount ��������
'====================================================================================================
Public Sub WriteProcessingCount(xTableSetting As TableSetting, procCount As Long)
    With ThisWorkbook.Worksheets(cstSheetMain)
        .Cells(xTableSetting.Row, TableSettingCol.ProcessCount).Value = procCount
    End With
End Sub

