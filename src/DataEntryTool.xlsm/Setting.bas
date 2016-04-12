Attribute VB_Name = "Setting"
Option Explicit
Option Private Module

'====================================================================================================
'
' �ݒ胂�W���[��
'
'====================================================================================================

'====================================================================================================
' �萔
'====================================================================================================
' ���O��`
Private Const cstDatabaseType = "DatabaseType"              ' �f�[�^�x�[�X���
Private Const cstServerName = "ServerName"                  ' �T�[�o��
Private Const cstPort = "Port"                              ' �|�[�g
Private Const cstDatabaseName = "DatabaseName"              ' �f�[�^�x�[�X��
Private Const cstUserId = "UserId"                          ' ���[�UID
Private Const cstPassword = "Password"                      ' �p�X���[�h
Private Const cstLinefeedCode = "LinefeedCode"              ' ���s�R�[�h
Private Const cstDateFormat = "DateFormat"                  ' ���t����
Private Const cstTimestampFormat = "TimestampFormat"        ' �^�C���X�^���v����


'====================================================================================================
' �����o�ϐ�
'====================================================================================================
Private mDatabaseType As String         ' �f�[�^�x�[�X���
Private mServerName As String           ' �T�[�o��
Private mPort As String                 ' �|�[�g
Private mDatabaseName As String         ' �f�[�^�x�[�X��
Private mUserId As String               ' ���[�UID
Private mPassword As String             ' �p�X���[�h
Private mLinefeedCode As String         ' ���s�R�[�h
Private mDateFormat As String           ' ���t����
Private mTimestampFormat As String      ' �^�C���X�^���v����


'====================================================================================================
' �f�[�^�x�[�X��ނ̎擾/�ݒ�
'====================================================================================================
Public Property Get DatabaseType() As String
    DatabaseType = mDatabaseType
End Property
Public Property Let DatabaseType(DatabaseType As String)
    mDatabaseType = DatabaseType
End Property


'====================================================================================================
' �T�[�o���̎擾/�ݒ�
'====================================================================================================
Public Property Get ServerName() As String
    ServerName = mServerName
End Property
Public Property Let ServerName(ServerName As String)
    mServerName = ServerName
End Property


'====================================================================================================
' �|�[�g�̎擾/�ݒ�
'====================================================================================================
Public Property Get Port() As String
    Port = mPort
End Property
Public Property Let Port(Port As String)
    mPort = Port
End Property


'====================================================================================================
' �f�[�^�x�[�X���̎擾/�ݒ�
'====================================================================================================
Public Property Get DatabaseName() As String
    DatabaseName = mDatabaseName
End Property
Public Property Let DatabaseName(DatabaseName As String)
    mDatabaseName = DatabaseName
End Property


'====================================================================================================
' ���[�UID�̎擾/�ݒ�
'====================================================================================================
Public Property Get UserId() As String
    UserId = mUserId
End Property
Public Property Let UserId(UserId As String)
    mUserId = UserId
End Property


'====================================================================================================
' �p�X���[�h�̎擾/�ݒ�
'====================================================================================================
Public Property Get Password() As String
    Password = mPassword
End Property
Public Property Let Password(Password As String)
    mPassword = Password
End Property


'====================================================================================================
' ���s�R�[�h�̎擾/�ݒ�
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
' ���t�����̎擾/�ݒ�
'====================================================================================================
Public Property Get DateFormat() As String
    DateFormat = mDateFormat
End Property
Public Property Let DateFormat(DateFormat As String)
    mDateFormat = DateFormat
End Property


'====================================================================================================
' �^�C���X�^���v�����̎擾/�ݒ�
'====================================================================================================
Public Property Get TimestampFormat() As String
    TimestampFormat = mTimestampFormat
End Property
Public Property Let TimestampFormat(TimestampFormat As String)
    mTimestampFormat = TimestampFormat
End Property


'====================================================================================================
' �ݒ胂�W���[�����\�����܂�
'----------------------------------------------------------------------------------------------------
' IN : xMacroType �}�N�����
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

    ' �ݒ胂�W���[���̃`�F�b�N
    If xMacroType = MacroType.Database Then
        Call CheckDbSetting
    ElseIf xMacroType = MacroType.DataEntry Then
        Call Setting.CheckDataEntrySetting
    End If
End Sub


'====================================================================================================
' �f�[�^�x�[�X�ݒ���`�F�b�N���܂�
'====================================================================================================
Private Sub CheckDbSetting()
On Error GoTo ErrHandler
    Call CheckInputValue(Setting.DatabaseType, "�f�[�^�x�[�X���")
    Call CheckInputValue(Setting.ServerName, "�T�[�o��")
    Call CheckInputValue(Setting.UserId, "���[�UID")
    Call CheckInputValue(Setting.Password, "�p�X���[�h")
    If Setting.DatabaseType = cstDatabaseTypePostgreSQL Then
        Call CheckInputValue(Setting.Port, "�|�[�g")
        Call CheckInputValue(Setting.DatabaseName, "�f�[�^�x�[�X��")
    End If
    Exit Sub
ErrHandler:
    Err.Raise ErrNumber.Warning, , "�f�[�^�x�[�X�ݒ��" & Err.Description
End Sub


'====================================================================================================
' �f�[�^�����ݒ���`�F�b�N���܂�
'====================================================================================================
Private Sub CheckDataEntrySetting()
On Error GoTo ErrHandler
    Call CheckDbSetting
    Call CheckInputValue(Setting.LinefeedCode, "���s�R�[�h")
    Call CheckInputValue(Setting.DateFormat, "���t����")
    Call CheckInputValue(Setting.TimestampFormat, "�^�C���X�^���v����")
    Exit Sub
ErrHandler:
    Err.Raise ErrNumber.Warning, , "�f�[�^�����ݒ��" & Err.Description
End Sub


'====================================================================================================
' ���͒l���`�F�b�N���܂�
'----------------------------------------------------------------------------------------------------
' IN : inputValue ���͒l
'    : itemName ���ږ�
'====================================================================================================
Private Sub CheckInputValue(inputValue As String, itemName As String)
    If inputValue = "" Then
        Err.Raise 1000, , "[" & itemName & "]����͂��Ă��������B"
    End If
End Sub
