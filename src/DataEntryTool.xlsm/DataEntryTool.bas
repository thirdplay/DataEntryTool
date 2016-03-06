Attribute VB_Name = "DataEntryTool"
Option Explicit


'====================================================================================================
' テーブルシートを作成します
'====================================================================================================
Public Sub CreateTableSheet()
On Error GoTo Finally
    Dim control As DataEntryControl
    Dim Start As Single
    Dim Finish As Single
    Start = Timer

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set control = New DataEntryControl

    With control
        .View = New DataEntryView
        .Model = New DataEntryModel
        Call .CreateTableSheet
    End With

Finally:
    Finish = Timer
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Set control = Nothing
    If Err.Number <> 0 Then
        MsgBox "テーブルシートの作成に失敗しました" & vbNewLine & Err.Description
    Else
        Debug.Print "処理時間:" & (Finish - Start)
        MsgBox "テーブルシートの作成が完了しました" & ":" & (Finish - Start)
    End If
End Sub


'====================================================================================================
' データを登録します
'====================================================================================================
Public Sub RegisterData()
On Error GoTo Finally
    Dim control As DataEntryControl

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set control = New DataEntryControl

    With control
        .View = New DataEntryView
        .Model = New DataEntryModel
        Call .EntryData(EntryType.Register)
    End With

Finally:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Set control = Nothing
    If Err.Number <> 0 Then
        MsgBox "データの登録に失敗しました" & vbNewLine & Err.Description
    Else
        MsgBox "データの登録が完了しました"
    End If
End Sub


'====================================================================================================
' データを更新します
'====================================================================================================
Public Sub UpdateData()
    MsgBox "UpdateData"
End Sub


'====================================================================================================
' データを削除します
'====================================================================================================
Public Sub RemoveData()
    MsgBox "RemoveData"
End Sub


