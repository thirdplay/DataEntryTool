Attribute VB_Name = "Setting"
Option Explicit
Option Private Module

'====================================================================================================
'
' �ݒ胂�W���[��
'
'====================================================================================================

'====================================================================================================
' �����o�ϐ�
'====================================================================================================
Private mDatabaseType As String     ' �f�[�^�x�[�X���
Private mServerName As String       ' �T�[�o��
Private mPort As String             ' �|�[�g
Private mDatabaseName As String     ' �f�[�^�x�[�X��
Private mUserId As String           ' ���[�UID
Private mPassword As String         ' �p�X���[�h
Private mLinefeedCode As String     ' ���s�R�[�h
Private mDateFormat As String       ' ���t����


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
