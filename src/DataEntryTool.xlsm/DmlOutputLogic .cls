VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DmlOutputLogic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' データ投入ロジック
'
'====================================================================================================

'====================================================================================================
' メンバ変数
'====================================================================================================
Private mDataEntryLogic As DataEntryLogic


'====================================================================================================
' コンストラクタ
'====================================================================================================
Public Sub Class_Initialize()
    Set mDataEntryLogic = New DataEntryLogic
End Sub


'====================================================================================================
' デストラクタ
'====================================================================================================
Public Sub Class_Terminate()
    Set mDataEntryLogic = Nothing
End Sub


'====================================================================================================
' 出力先を取得します
'----------------------------------------------------------------------------------------------------
' OUT: 出力先
'====================================================================================================
Public Function GetOutputDir()
    Dim xOutputDir As String

    xOutputDir = Setting.OutputDir
    If xOutputDir = "" Then
        With Application.FileDialog(msoFileDialogFolderPicker)
            If .Show = True Then
                xOutputDir = .SelectedItems(1)
            Else
                GetOutputDir = ""
                Exit Function
            End If
        End With
    End If

    If Right(xOutputDir, 1) <> "\" Then
        xOutputDir = xOutputDir & "\"
    End If
    GetOutputDir = xOutputDir
End Function


'====================================================================================================
' DMLを出力します
'----------------------------------------------------------------------------------------------------
' IN : xEntryType 投入種別
'    : xEntryData 投入データ
' OUT: 処理件数
'====================================================================================================
Public Function OutputDml(xEntryType As EntryType, xEntryData As EntryData) As Long
    ' データ投入クエリの生成
    'queries = mDataEntryLogic.MakeQueries(xEntryType, ed)

End Function
