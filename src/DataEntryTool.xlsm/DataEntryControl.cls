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
' データ投入コントロール
'
'====================================================================================================

'====================================================================================================
' メンバ変数
'====================================================================================================
Private mView As DataEntryView          ' データ投入ビュー
Private mModel As DataEntryModel        ' データ投入モデル


'====================================================================================================
' コンストラクタ
'====================================================================================================
Private Sub Class_Initialize()
End Sub


'====================================================================================================
' デストラクタ
'====================================================================================================
Private Sub Class_Terminate()
    Set mView = Nothing
    Set mModel = Nothing
End Sub


'====================================================================================================
' ビューの設定
'====================================================================================================
Public Property Let View(View As DataEntryView)
    Set mView = View
End Property


'====================================================================================================
' モデルの設定
'====================================================================================================
Public Property Let Model(Model As DataEntryModel)
    Set mModel = Model
End Property


'====================================================================================================
' テーブル定義再作成
'====================================================================================================
Public Sub RecreateTableDefinition()
On Error GoTo Finally
    Dim setting As Setting
    Dim tableNames As Collection
    Dim tableInfoList As Collection

    ' 設定とテーブル名一覧の取得
    Set setting = mView.GetSetting
    Set tableNames = mView.GetTableNameList

    ' テーブル名一覧を元に、Modelからテーブル情報一覧を取得
    Call mModel.Setup(setting)
    Set tableInfoList = mModel.GetTableInfoList(tableNames)

    ' テーブル情報一覧を元に、テーブル定義シートを再作成する
    Call mView.RecreateTableDefinition(tableInfoList)
Finally:
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Sub

