VERSION 5.00
Begin VB.Form frmSetJlhTmptParkir 
   BackColor       =   &H00E6F1EF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setting Jumlah Tempat Parkir"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7410
   Icon            =   "frmSetJlhTmptParkir.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   7410
   Begin VB.Frame Frame1 
      BackColor       =   &H00E6F1EF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   315
      TabIndex        =   6
      Top             =   1260
      Width           =   6720
      Begin VB.TextBox txtJlhTmptParkir 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3945
         MaxLength       =   15
         TabIndex        =   0
         Text            =   "0"
         Top             =   345
         Width           =   1845
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Tempat Parkir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   600
         TabIndex        =   8
         Top             =   405
         Width           =   2700
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "= "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3540
         TabIndex        =   7
         Top             =   405
         Width           =   270
      End
   End
   Begin VB.CommandButton cmdSimpan 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&SIMPAN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2730
      TabIndex        =   1
      Top             =   3165
      Width           =   2055
   End
   Begin VB.CommandButton cmdKeluar 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&KELUAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4830
      TabIndex        =   2
      Top             =   3165
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H002A5CE4&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   -60
      ScaleHeight     =   255
      ScaleWidth      =   10020
      TabIndex        =   5
      Top             =   3840
      Width           =   10050
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H002A5CE4&
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   -15
      ScaleHeight     =   600
      ScaleWidth      =   10020
      TabIndex        =   3
      Top             =   0
      Width           =   10050
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Setting Jumlah Tempat Parkir"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   30
         Width           =   4755
      End
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   150
      Picture         =   "frmSetJlhTmptParkir.frx":2CFA
      Top             =   3045
      Width           =   1800
   End
End
Attribute VB_Name = "frmSetJlhTmptParkir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdKeluar_Click()
    
    Unload Me
    
End Sub

Private Sub cmdSimpan_Click()
    
    If Val(txtJlhTmptParkir.Text) = 0 Then
        MsgBox "Jumlah tempat parkir belum diisi !", vbCritical
        Exit Sub
    End If
    
    'Jumlah tempat parkir
    S = "Update SettingID Set SetValue = '" & _
        Format(txtJlhTmptParkir, "0") & "' " & _
        "Where SetID = 'Jumlah Tempat Parkir'"
    oConn.Execute S
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    'Load Data
    Dim rs As New ADODB.Recordset
    
    S = "Select SetID, SetValue From SettingID " & _
        "Where SetID = 'Jumlah Tempat Parkir'"
    rs.Open S, oConn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then
        
        txtJlhTmptParkir = Format(rs!SetValue, "#,##0")
        
    End If
    
    Call CloseRS(rs)
    
End Sub

Private Sub txtJlhTmptParkir_GotFocus()
    
    Call GotFocus(txtJlhTmptParkir)
    
End Sub

Private Sub txtJlhTmptParkir_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> vbKeyBack And IsNumeric(Chr(KeyAscii)) = False Then
    
       KeyAscii = 0
       
    End If
    
End Sub

Private Sub txtJlhTmptParkir_LostFocus()
    
    Call LostFocus(txtJlhTmptParkir)
    
End Sub

