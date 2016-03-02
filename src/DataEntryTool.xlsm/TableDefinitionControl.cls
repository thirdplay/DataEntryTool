VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableDefinitionControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' テーブル定義コントロール
'
'====================================================================================================

'====================================================================================================
' メンバ変数
'====================================================================================================
Private mView As DataEntryView              ' データ投入ビュー
Private mModel As TableDefinitionModel      ' テーブル定義モデル


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
Public Property Let Model(Model As TableDefinitionModel)
    Set mModel = Model
End Property


'====================================================================================================
' 実行
'====================================================================================================
Public Sub Execute()
    Dim dbInfo As DatabaseInfo
    Dim tableNames As Collection
    Dim tableInfoList As Collection
    
    ' DB情報の取得
    Set dbInfo = mView.GetDatabaseInfo
    ' テーブル名一覧の取得
    Set tableNames = mView.GetTableNameList

    ' テーブル名一覧を元に、Modelからテーブル情報一覧を取得
    mModel.DatabaseInfo = dbInfo
    Set tableInfoList = mModel.GetTableInfoList(tableNames)

    ' 3. テーブル情報一覧を元に、テーブルシートを再作成
    'Call mView.RecreateTableDefinition tableInfoList
    
    Set dbInfo = Nothing
End Sub

