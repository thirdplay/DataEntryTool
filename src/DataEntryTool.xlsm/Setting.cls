VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' 設定
'
'====================================================================================================

'====================================================================================================
' メンバ変数
'====================================================================================================
Private mDatabaseType As String     ' データベース種類
Private mServerName As String       ' サーバ名
Private mPort As String             ' ポート
Private mDatabaseName As String     ' データベース名
Private mUserId As String           ' ユーザID
Private mPassword As String         ' パスワード


'====================================================================================================
' コンストラクタ
'====================================================================================================
Private Sub Class_Initialize()
End Sub


'====================================================================================================
' デストラクタ
'====================================================================================================
Private Sub Class_Terminate()
End Sub


'====================================================================================================
' データベース種類の取得/設定
'====================================================================================================
Public Property Get DatabaseType() As String
    DatabaseType = mDatabaseType
End Property
Public Property Let DatabaseType(DatabaseType As String)
    mDatabaseType = DatabaseType
End Property


'====================================================================================================
' サーバ名の取得/設定
'====================================================================================================
Public Property Get ServerName() As String
    ServerName = mServerName
End Property
Public Property Let ServerName(ServerName As String)
    mServerName = ServerName
End Property


'====================================================================================================
' ポートの取得/設定
'====================================================================================================
Public Property Get Port() As String
    Port = mPort
End Property
Public Property Let Port(Port As String)
    mPort = Port
End Property


'====================================================================================================
' データベース名の取得/設定
'====================================================================================================
Public Property Get DatabaseName() As String
    DatabaseName = mDatabaseName
End Property
Public Property Let DatabaseName(DatabaseName As String)
    mDatabaseName = DatabaseName
End Property


'====================================================================================================
' ユーザIDの取得/設定
'====================================================================================================
Public Property Get UserId() As String
    UserId = mUserId
End Property
Public Property Let UserId(UserId As String)
    mUserId = UserId
End Property


'====================================================================================================
' パスワードの取得/設定
'====================================================================================================
Public Property Get Password() As String
    Password = mPassword
End Property
Public Property Let Password(Password As String)
    mPassword = Password
End Property

