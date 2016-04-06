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
Public Property Let View(xView As DataEntryView)
    Set mView = xView
    Call mView.Setup
End Property


'====================================================================================================
' モデルの設定
'====================================================================================================
Public Property Let Model(xModel As DataEntryModel)
    Set mModel = xModel
    Call mModel.Setup
End Property


'====================================================================================================
' テーブルシートを作成します
'====================================================================================================
Public Sub CreateTableSheet()
    Dim tableSettings As Collection
    Dim tableDefinitions As Collection

    ' テーブル設定リストの取得
    Debug.Print "DataEntryControl.CreateTableSheet:GetTableSettings"
    Set tableSettings = mView.GetTableSettings

    ' テーブル設定リストを元に、Modelからテーブル定義を取得
    Debug.Print "DataEntryControl.CreateTableSheet:GetTableDefinitions"
    Set tableDefinitions = mModel.GetTableDefinitions(tableSettings)

    ' テーブル定義を元に、テーブルシートを作成する
    Debug.Print "DataEntryControl.CreateTableSheet:CreateTableSheet"
    Call mView.CreateTableSheet(tableDefinitions)
End Sub


'====================================================================================================
' データを投入します
' ---------------------------------------------------------------------------------------------------
' IN : xEntryType 投入種類
'====================================================================================================
Public Sub EntryData(xEntryType As EntryType)
    Dim tableSettings As Collection
    Dim ts As TableSetting
    Dim ed As EntryData
    Dim count As Long

    ' 処理件数のクリア
    Call mView.ClearProcessingCount

    ' 対象テーブル設定の取得
    Debug.Print "DataEntryControl.CreateTableSheet:GetTableSettingsByCriteria"
    Set tableSettings = mView.GetTableSettingsByCriteria(True)

    ' 対象テーブル設定を全て処理
    For Each ts In tableSettings
        ' 対象テーブルのテーブルデータの取得
        Set ed = mView.GetEntryData(ts.PhysicsName)

        ' データ投入実行
        count = mModel.ExecuteDataEntry(xEntryType, ed)

        ' 処理件数の書き込み
        Call mView.WriteProcessingCount(ts, count)
    Next
End Sub

