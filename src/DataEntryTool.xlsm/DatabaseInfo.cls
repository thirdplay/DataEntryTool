VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatabaseInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================================================
'
' �f�[�^�x�[�X���
'
'====================================================================================================

'====================================================================================================
' �����o�ϐ�
'====================================================================================================
Private mServerName As String       ' �T�[�o��
Private mPort As String             ' �|�[�g
Private mDatabaseName As String     ' �f�[�^�x�[�X��
Private mUserId As String           ' ���[�UID
Private mPassword As String         ' �p�X���[�h


'====================================================================================================
' �R���X�g���N�^
'====================================================================================================
Private Sub Class_Initialize()
End Sub


'====================================================================================================
' �f�X�g���N�^
'====================================================================================================
Private Sub Class_Terminate()
End Sub


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
