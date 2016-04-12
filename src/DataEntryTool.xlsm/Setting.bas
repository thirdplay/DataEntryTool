Attribute VB_Name = "Setting"
Option Explicit
Option Private Module

'====================================================================================================
'
' 設定モジュール
'
'====================================================================================================

'====================================================================================================
' 定数
'====================================================================================================
' 名前定義
Private Const cstDatabaseType = "DatabaseType"              ' データベース種類
Private Const cstServerName = "ServerName"                  ' サーバ名
Private Const cstPort = "Port"                              ' ポート
Private Const cstDatabaseName = "DatabaseName"              ' データベース名
Private Const cstUserId = "UserId"                          ' ユーザID
Private Const cstPassword = "Password"                      ' パスワード
Private Const cstLinefeedCode = "LinefeedCode"              ' 改行コード
Private Const cstDateFormat = "DateFormat"                  ' 日付書式
Private Const cstTimestampFormat = "TimestampFormat"        ' タイムスタンプ書式


'====================================================================================================
' メンバ変数
'====================================================================================================
Private mDatabaseType As String         ' データベース種類
Private mServerName As String           ' サーバ名
Private mPort As String                 ' ポート
Private mDatabaseName As String         ' データベース名
Private mUserId As String               ' ユーザID
Private mPassword As String             ' パスワード
Private mLinefeedCode As String         ' 改行コード
Private mDateFormat As String           ' 日付書式
Private mTimestampFormat As String      ' タイムスタンプ書式


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


'====================================================================================================
' タイムスタンプ書式の取得/設定
'====================================================================================================
Public Property Get TimestampFormat() As String
    TimestampFormat = mTimestampFormat
End Property
Public Property Let TimestampFormat(TimestampFormat As String)
    mTimestampFormat = TimestampFormat
End Property


'====================================================================================================
' 設定モジュールを構成します
'----------------------------------------------------------------------------------------------------
' IN : xMacroType マクロ種別
'====================================================================================================
Public Sub Setup(xMacroType As MacroType)
    With ThisWorkbook.Worksheets(cstSheetMain)
        Setting.DatabaseType = .Range(cstDatabaseType).Value
        Setting.ServerName = .Range(cstServerName).Value
        Setting.Port = .Range(cstPort).Value
        Setting.DatabaseName = .Range(cstDatabaseName).Value
        Setting.UserId = .Range(cstUserId).Value
        Setting.Password = .Range(cstPassword).Value
        Setting.LinefeedCode = .Range(cstLinefeedCode).Value
        Setting.DateFormat = .Range(cstDateFormat).Value
        Setting.TimestampFormat = .Range(cstTimestampFormat).Value
    End With

    ' 設定モジュールのチェック
    If xMacroType = MacroType.Database Then
        Call CheckDbSetting
    ElseIf xMacroType = MacroType.DataEntry Then
        Call Setting.CheckDataEntrySetting
    End If
End Sub


'====================================================================================================
' データベース設定をチェックします
'====================================================================================================
Private Sub CheckDbSetting()
On Error GoTo ErrHandler
    Call CheckInputValue(Setting.DatabaseType, "データベース種類")
    Call CheckInputValue(Setting.ServerName, "サーバ名")
    Call CheckInputValue(Setting.UserId, "ユーザID")
    Call CheckInputValue(Setting.Password, "パスワード")
    If Setting.DatabaseType = cstDatabaseTypePostgreSQL Then
        Call CheckInputValue(Setting.Port, "ポート")
        Call CheckInputValue(Setting.DatabaseName, "データベース名")
    End If
    Exit Sub
ErrHandler:
    Err.Raise ErrNumber.Warning, , "データベース設定の" & Err.Description
End Sub


'====================================================================================================
' データ投入設定をチェックします
'====================================================================================================
Private Sub CheckDataEntrySetting()
On Error GoTo ErrHandler
    Call CheckDbSetting
    Call CheckInputValue(Setting.LinefeedCode, "改行コード")
    Call CheckInputValue(Setting.DateFormat, "日付書式")
    Call CheckInputValue(Setting.TimestampFormat, "タイムスタンプ書式")
    Exit Sub
ErrHandler:
    Err.Raise ErrNumber.Warning, , "データ投入設定の" & Err.Description
End Sub


'====================================================================================================
' 入力値をチェックします
'----------------------------------------------------------------------------------------------------
' IN : inputValue 入力値
'    : itemName 項目名
'====================================================================================================
Private Sub CheckInputValue(inputValue As String, itemName As String)
    If inputValue = "" Then
        Err.Raise 1000, , "[" & itemName & "]を入力してください。"
    End If
End Sub
