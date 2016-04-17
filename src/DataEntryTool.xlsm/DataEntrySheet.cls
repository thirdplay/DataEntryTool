VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataEntrySheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' �e�[�u�������V�[�g
'
'====================================================================================================

'====================================================================================================
' �e�[�u���ݒ胊�X�g��ԋp���܂�
'----------------------------------------------------------------------------------------------------
' IN : isEntryTarget �����Ώۃt���O(True:�����Ώۂ̂�,False:�S��)
' OUT: �e�[�u���ݒ�̘A���z��
'====================================================================================================
Public Function GetTableSettings(isEntryTarget As Boolean) As Dictionary
    Dim rowIndex As Long
    Dim ts As TableSetting
    Dim isTarget As Boolean
    Dim xTableName As String
    Dim dataCnt As Long
    Dim tableRange As Range
    Dim tr As Range
    Dim dic As Dictionary
    Set dic = New Dictionary

    With ThisWorkbook.Worksheets(cstSheetMain)
        rowIndex = .Range(cstTableBase).Row + 1
        dataCnt = WorkSheetEx.GetRowDataCount(cstSheetMain, rowIndex, TableSettingCol.PhysicsName)
        If dataCnt > 0 Then
            Set tableRange = .Range(.Cells(rowIndex, TableSettingCol.PhysicsName), .Cells(rowIndex + dataCnt - 1, TableSettingCol.Max))

            dataCnt = 0
            For Each tr In tableRange.Rows
                ' �����Ώۃt���O��true�̏ꍇ�A�����ΏۊO�̃e�[�u���͏������Ȃ�
                xTableName = tr.Cells(1, TableSettingCol.PhysicsName).Value
                isTarget = True
                If isEntryTarget Then
                    If tr.Cells(1, TableSettingCol.DataEntryTarget).Value = "" Then
                        isTarget = False
                    ElseIf Not WorkBookEx.ExistsSheet(xTableName) Then
                        Err.Raise ErrNumber.Error, , "�e�[�u��[" & xTableName & "]�̃V�[�g�����݂��܂���B" & vbNewLine & "�e�[�u���V�[�g�쐬���s���A�e�[�u���V�[�g���쐬���Ă��������B"
                    Else
                        dataCnt = WorkSheetEx.GetRowDataCount(xTableName, cstTableRecordBase, 1)
                        If dataCnt = 0 Then
                            isTarget = False
                        End If
                    End If
                End If

                If isTarget Then
                    If dic.Exists(tr.Cells(1, TableSettingCol.PhysicsName).Value) Then
                        Err.Raise ErrNumber.Error, , "�e�[�u��[" & xTableName & "]���d�����Ă��܂��B"
                    End If
                    Set ts = New TableSetting
                    ts.Row = tr.Row
                    ts.PhysicsName = tr.Cells(1, TableSettingCol.PhysicsName).Value
                    ts.LogicalName = tr.Cells(1, TableSettingCol.LogicalName).Value
                    ts.DataEntryTarget = tr.Cells(1, TableSettingCol.DataEntryTarget).Value
                    ts.DataCount = dataCnt
                    Call dic.Add(ts.PhysicsName, ts)
                End If
            Next
        End If
    End With
    Set GetTableSettings = dic
End Function


'====================================================================================================
' �e�[�u���ݒ�Ƀn�C�p�[�����N��ݒ肵�܂�
'----------------------------------------------------------------------------------------------------
' IN : tablSeetings �e�[�u���ݒ�̘A���z��
'====================================================================================================
Public Sub SetHyperlink(tableSettings As Dictionary)
    Dim xKey As Variant
    Dim ts As TableSetting

    With ThisWorkbook.Worksheets(cstSheetMain)
        For Each xKey In tableSettings
            Set ts = tableSettings(xKey)
            If .Cells(ts.Row, TableSettingCol.LogicalName) <> "" Then
                Call .Hyperlinks.Add(Anchor:=.Cells(ts.Row, TableSettingCol.LogicalName), Address:="#" & ts.PhysicsName & "!A1")
            End If
        Next
    End With
End Sub


'====================================================================================================
' �����������N���A���܂�
'====================================================================================================
Public Sub ClearProcessingCount()
    Dim baseRowIndex As Long
    Dim dataCnt As Long

    With ThisWorkbook.Worksheets(cstSheetMain)
        ' �e�[�u���ݒ萔�̎擾
        baseRowIndex = .Range(cstTableBase).Row + 1
        dataCnt = WorkSheetEx.GetRowDataCount(cstSheetMain, baseRowIndex, TableSettingCol.PhysicsName)

        ' ���������̃N���A
        Call .Range(.Cells(baseRowIndex, TableSettingCol.ProcessCount), .Cells(baseRowIndex + dataCnt, TableSettingCol.ProcessCount)).ClearContents
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