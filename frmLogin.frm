VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function LogonUser Lib "Advapi32" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As Any, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function NetUserChangePassword Lib "netapi32.dll" (ByVal sDomain As String, ByVal sUserName As String, ByVal sOldPassword As String, ByVal sNewPassword As String) As Long

'Purpose   :    Checks if a the NT password for a user is correct.
'Inputs    :    UserName                The username
'               Password                The password
'               [Domain]                If DOMAIN is omitted uses the local account database.
'Outputs   :    Returns True if the password and user name are valid.
'Author    :    Andrew Baker
'Date      :    25/03/2000
'Notes     :    Windows NT and 2000 ONLY. Will work on any machine.
'               Slower than the UserCheckPassword function, but more reliable.

Function UserValidate(sUserName As String, sPassword As String, Optional sDomain As String) As Boolean
    Dim lReturn As Long
    Const NERR_BASE = 2100
    Const NERR_PasswordCantChange = NERR_BASE + 143
    Const NERR_PasswordHistConflict = NERR_BASE + 144
    Const NERR_PasswordTooShort = NERR_BASE + 145
    Const NERR_PasswordTooRecent = NERR_BASE + 146
    
    If Len(sDomain) = 0 Then
        sDomain = Environ$("USERDOMAIN")
    End If
    
    'Call API to check password.
    lReturn = NetUserChangePassword(StrConv(sDomain, vbUnicode), StrConv(sUserName, vbUnicode), StrConv(sPassword, vbUnicode), StrConv(sPassword, vbUnicode))
    
    'Test return value.
    Select Case lReturn
    Case 0, NERR_PasswordCantChange, NERR_PasswordHistConflict, NERR_PasswordTooShort, NERR_PasswordTooRecent
        UserValidate = True
    Case Else
        UserValidate = False
    End Select
End Function



'Purpose     :  Return the error message associated with LastDLLError
'Inputs      :  lLastDLLError               The error number of the last DLL error (from Err.LastDllError)
'Outputs     :  Returns the error message associated with the DLL error number
'Author      :  Andrew Baker
'Date        :  13/11/2000 10:14
'Notes       :
'Revisions   :

Public Function DLLErrorText(ByVal lLastDLLError As Long) As String
    Dim sBuff As String * 256
    Dim lCount As Long
    Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100, FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
    Const FORMAT_MESSAGE_FROM_HMODULE = &H800, FORMAT_MESSAGE_FROM_STRING = &H400
    Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000, FORMAT_MESSAGE_IGNORE_INSERTS = &H200
    Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
    
    lCount = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, lLastDLLError, 0&, sBuff, Len(sBuff), ByVal 0)
    If lCount Then
        DLLErrorText = Left$(sBuff, lCount - 2)    'Remove line feeds
    End If
    
End Function

Private Sub Form_Load()
    txtUserName.Text = Environ$("USERNAME")
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    'check for correct password
    Screen.MousePointer = 11
    If UserValidate(txtUserName.Text, txtPassword.Text) Then
        Screen.MousePointer = 0
        Unload Me
    Else
        Screen.MousePointer = 0
        MsgBox "Invalid Username/Password, try again!", , "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If

End Sub
