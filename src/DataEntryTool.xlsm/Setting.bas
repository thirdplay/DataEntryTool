Attribute VB_Name = "Setting"
Option Explicit
Option Private Module

'====================================================================================================
'
' 設定モジュール
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
Private mLinefeedCode As String     ' 改行コード
Private mDateFormat As String       ' 日付書式


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


'====================================================================================================
' 改行コードの取得/設定
'====================================================================================================
Public Property Get LinefeedCode() As String
    Dim result As String
    result = "|| CHR(10) ||"
    If mLinefeedCode = cstLinefeedCodeCRLF Then
        result = "|| CHR(13) " & result
    End If
    LinefeedCode = result
End Property
Public Property Let LinefeedCode(LinefeedCode As String)
    mLinefeedCode = LinefeedCode
End Property


'====================================================================================================
' 日付書式の取得/設定
'====================================================================================================
Public Property Get DateFormat() As String
    DateFormat = mDateFormat
End Property
Public Property Let DateFormat(DateFormat As String)
    mDateFormat = DateFormat
End Property
