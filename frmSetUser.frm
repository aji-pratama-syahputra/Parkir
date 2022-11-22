VERSION 5.00
Begin VB.Form frmSetUser 
   BackColor       =   &H00E6F1EF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setting Operator"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7410
   Icon            =   "frmSetUser.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   7410
   Begin VB.CommandButton cmdHapus 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&HAPUS"
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
      Left            =   4710
      TabIndex        =   5
      Top             =   5482
      Width           =   1290
   End
   Begin VB.CommandButton cmdBaru 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&BARU"
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
      Left            =   2070
      TabIndex        =   7
      Top             =   5482
      Width           =   1290
   End
   Begin VB.TextBox txtKet 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   2730
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3120
      Width           =   4335
   End
   Begin VB.ComboBox cboTipe 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmSetUser.frx":2CFA
      Left            =   2730
      List            =   "frmSetUser.frx":2D07
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2520
      Width           =   2205
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2730
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1920
      Width           =   4335
   End
   Begin VB.ComboBox cboOp 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2730
      TabIndex        =   0
      Top             =   1320
      Width           =   2205
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
      Left            =   3390
      TabIndex        =   4
      Top             =   5482
      Width           =   1290
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
      Left            =   6030
      TabIndex        =   6
      Top             =   5482
      Width           =   1290
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H002A5CE4&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   -60
      ScaleHeight     =   255
      ScaleWidth      =   10020
      TabIndex        =   10
      Top             =   6165
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
      TabIndex        =   8
      Top             =   0
      Width           =   10050
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Setting Operator"
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
         TabIndex        =   9
         Top             =   30
         Width           =   2625
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   2460
      TabIndex        =   18
      Top             =   3165
      Width           =   90
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan"
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
      Left            =   315
      TabIndex        =   17
      Top             =   3165
      Width           =   1440
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   2460
      TabIndex        =   16
      Top             =   2565
      Width           =   90
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipe operator"
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
      Left            =   315
      TabIndex        =   15
      Top             =   2565
      Width           =   1650
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   2460
      TabIndex        =   14
      Top             =   1965
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   315
      TabIndex        =   13
      Top             =   1965
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Left            =   2460
      TabIndex        =   12
      Top             =   1365
      Width           =   90
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama operator"
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
      Left            =   315
      TabIndex        =   11
      Top             =   1365
      Width           =   1860
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   150
      Picture         =   "frmSetUser.frx":2D35
      Top             =   5355
      Width           =   1800
   End
End
Attribute VB_Name = "frmSetUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cboOp_Change()
    Dim rsT As New ADODB.Recordset
    Dim strP As String
    
    strP = "Select * From UserList Where " & _
           "UserName = '" & Trim(cboOp.Text) & "'"
    rsT.Open strP, oConn, adOpenStatic, adLockReadOnly, adCmdText
    If rsT.RecordCount > 0 Then
        
        txtPassword.Text = rsT!UserPwd
        txtKet.Text = rsT!Keterangan
        
        If rsT!UserType = "Administrator" Then
            cboTipe.ListIndex = 0
        ElseIf rsT!UserType = "Operator IN" Then
            cboTipe.ListIndex = 1
        ElseIf rsT!UserType = "Operator OUT" Then
            cboTipe.ListIndex = 2
        End If
        
    Else
    
        txtPassword.Text = ""
        cboTipe.ListIndex = -1
        txtKet = ""
        
    End If
    
    Call CloseRS(rsT)

End Sub

Private Sub cboOp_Click()
    Call cboOp_Change
End Sub

Private Sub cboOp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cboTipe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmdBaru_Click()
    
    cboOp.Text = ""
    cboOp.SetFocus
    txtPassword.Text = ""
    txtKet = ""
    cboTipe.ListIndex = -1
    
End Sub

Private Sub cmdHapus_Click()
        
    'Periksa sebelumnya ada atau tidak
    Dim rsD As New ADODB.Recordset
    
    S = "Select UserName From UserList " & _
        "Where UserName = '" & Trim(cboOp.Text) & "'"
    rsD.Open S, oConn, adOpenStatic, adLockReadOnly, adCmdText
    
    If rsD.RecordCount <= 0 Then
        
        Call CloseRS(rsD)
        MsgBox "Tidak ada data untuk dihapus !", vbCritical, "Delete"
        cboOp.SetFocus
        Exit Sub
        
    End If
    
    Call CloseRS(rsD)
        
    'Konfirmasi Hapus
    If MsgBox("Apakah anda yakin untuk menghapus data ini ?", _
              vbQuestion + vbYesNo, "Hapus") = vbNo Then
        Exit Sub
    End If

    On Error GoTo errDelete
    
    'Mulai hapus
    oConn.BeginTrans
    
    'Sebelumnya sudah ada, hapus
    S = "Delete From UserList " & _
        "Where UserName = '" & Trim(cboOp.Text) & "'"
    oConn.Execute S
    
    'Hapus Selesai
    oConn.CommitTrans
    
    'Kosongkan semua
    cmdBaru.Value = True
    Call RefreshData
    
    Exit Sub

errDelete:
    oConn.RollbackTrans
    MsgBox "Proses Hapus Gagal !" & vbCrLf & Err.Description, vbInformation, "Error"
    
End Sub

Private Sub cmdKeluar_Click()
    
    Unload Me
    
End Sub

Private Sub cmdSimpan_Click()
        
    'Periksa input
    If Trim(cboOp.Text) = "" Then
        
        MsgBox "Nama operator belum diisi !", vbCritical
        cboOp.SetFocus
        Exit Sub
    
    ElseIf Trim(txtPassword.Text) = "" Then
        
        MsgBox "Password belum diisi !", vbCritical
        txtPassword.SetFocus
        Exit Sub
    
    ElseIf Trim(cboTipe.Text) = "" Then
        
        MsgBox "Tipe Operator belum diisi !", vbCritical
        cboTipe.SetFocus
        Exit Sub
        
    End If
    
    On Error GoTo errSave
    
    'Mulai Simpan
    oConn.BeginTrans
        
    'Hapus
    S = "Delete From UserList " & _
        "Where UserName = '" & Trim(cboOp) & "'"
    oConn.Execute S
    
    'Simpan
    S = "Insert Into UserList(UserName, UserPwd, UserType, Keterangan) Values ('" & _
        Trim(cboOp.Text) & "','" & Trim(txtPassword.Text) & "','" & _
        cboTipe.Text & "','" & txtKet.Text & "')"
    oConn.Execute S
    
    'Simpan Selesai
    oConn.CommitTrans
    
    'Kosongkan semua
    cmdBaru.Value = True
    Call RefreshData
    
    Exit Sub

errSave:
    oConn.RollbackTrans
    MsgBox "Proses Simpan Gagal !" & vbCrLf & Err.Description, vbInformation, "Error"
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub

Private Sub Form_Load()

    Call RefreshData
    
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub RefreshData()

    Dim rs As New ADODB.Recordset
    
    S = "Select UserName From UserList Order by UserName"
    rs.Open S, oConn, adOpenStatic, adLockReadOnly, adCmdText
    
    cboOp.Clear
    For I = 1 To rs.RecordCount
        
        cboOp.AddItem rs!UserName
        rs.MoveNext
        
    Next I
    
    Call CloseRS(rs)
    
End Sub
