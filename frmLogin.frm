VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00E6F1EF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " User Login ..."
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5640
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3840
      TabIndex        =   3
      Top             =   2400
      Width           =   1485
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2280
      TabIndex        =   2
      Top             =   2400
      Width           =   1485
   End
   Begin VB.TextBox txtPwd 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   1965
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1680
      Width           =   3360
   End
   Begin VB.TextBox txtUser 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1965
      MaxLength       =   30
      TabIndex        =   0
      Top             =   1215
      Width           =   3360
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1710
      TabIndex        =   8
      Top             =   1740
      Width           =   75
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1710
      TabIndex        =   7
      Top             =   1245
      Width           =   75
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   1740
      Width           =   945
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   240
      Picture         =   "frmLogin.frx":2CFA
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Login ..."
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   870
      TabIndex        =   5
      Top             =   285
      Width           =   1965
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   1245
      Width           =   1080
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   3525
      Picture         =   "frmLogin.frx":393E
      Top             =   113
      Width           =   1800
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdOK_Click()
    Dim RS As New ADODB.Recordset
    Dim strPwd As String
    
    S = "Select UserName, UserPwd, UserType From UserList " & _
        "Where UserName = '" & txtUser.Text & "' " & _
        "And UserPwd = '" & txtPwd.Text & "' "
    RS.Open S, oConn, adOpenStatic, adLockReadOnly, adCmdText
    If RS.RecordCount > 0 Then
        
        'User Name
        strUser = RS!UserName
        
        'User password
        strPwd = RS!UserPwd
        
        'User Type
        strUserType = RS!UserType
        
        Call CloseRS(RS)
        
        'Case Sensitive
        If strUser <> txtUser.Text Or strPwd <> txtPwd.Text Then
            
            MsgBox "User name atau password salah !", vbCritical
        
        Else
        
            Unload Me
            frmSplash.Show
                        
        End If
        
    Else
        
        Call CloseRS(RS)
        MsgBox "User name atau password salah !", vbCritical
        
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub

Private Sub txtPwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdOK.Value = True
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
