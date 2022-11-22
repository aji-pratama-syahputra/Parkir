VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSetBiaya 
   BackColor       =   &H00E6F1EF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setting Biaya Parkir"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7410
   Icon            =   "frmSetBiaya.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   7410
   Begin VB.Frame Frame1 
      BackColor       =   &H00E6F1EF&
      Height          =   1440
      Left            =   540
      TabIndex        =   18
      Top             =   2640
      Width           =   6345
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Hapus dari tabel"
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
         Left            =   3180
         TabIndex        =   9
         Top             =   780
         Width           =   2925
      End
      Begin VB.CommandButton cmdInsert 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Masukkan ke tabel"
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
         Left            =   210
         TabIndex        =   8
         Top             =   780
         Width           =   2925
      End
      Begin VB.TextBox txtBiaya 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4845
         MaxLength       =   10
         TabIndex        =   7
         Text            =   "0"
         Top             =   270
         Width           =   1260
      End
      Begin VB.TextBox txtSampai 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3165
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "0"
         Top             =   270
         Width           =   495
      End
      Begin VB.TextBox txtDari 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1290
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "0"
         Top             =   270
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "jam = Rp."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3795
         TabIndex        =   21
         Top             =   330
         Width           =   945
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "jam  sampai"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1890
         TabIndex        =   20
         Top             =   330
         Width           =   1170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Biaya dari "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   210
         TabIndex        =   19
         Top             =   330
         Width           =   1035
      End
   End
   Begin VB.TextBox txtBiayaJam 
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
      Left            =   4245
      MaxLength       =   10
      TabIndex        =   4
      Text            =   "0"
      Top             =   1590
      Width           =   1770
   End
   Begin VB.TextBox txtBiayaStatis 
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
      Left            =   4245
      MaxLength       =   10
      TabIndex        =   3
      Text            =   "0"
      Top             =   960
      Width           =   1770
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00E6F1EF&
      Caption         =   "Biaya parkir berdasarkan range waktu"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   5430
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00E6F1EF&
      Caption         =   "Biaya parkir per jam"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   1
      Top             =   1620
      Width           =   3090
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E6F1EF&
      Caption         =   "Biaya parkir statis"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   3090
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
      TabIndex        =   11
      Top             =   6735
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
      TabIndex        =   12
      Top             =   6735
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
      TabIndex        =   15
      Top             =   7425
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
      TabIndex        =   13
      Top             =   0
      Width           =   10050
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Setting Biaya Parkir"
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
         TabIndex        =   14
         Top             =   30
         Width           =   3270
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Tbl 
      Height          =   2085
      Left            =   540
      TabIndex        =   10
      Top             =   4155
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   3678
      _Version        =   393216
      BackColorBkg    =   15135215
      HighLight       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "=   Rp."
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
      Left            =   3330
      TabIndex        =   17
      Top             =   1650
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "=   Rp."
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
      Left            =   3330
      TabIndex        =   16
      Top             =   1020
      Width           =   825
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   150
      Picture         =   "frmSetBiaya.frx":2CFA
      Top             =   6615
      Width           =   1800
   End
End
Attribute VB_Name = "frmSetBiaya"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdDelete_Click()
    
    If Tbl.Row > 0 Then
        If Tbl.Rows = 2 Then
            Tbl.Rows = 1
        Else
            Tbl.RemoveItem Tbl.Row
        End If
    End If
    
End Sub

Private Sub cmdInsert_Click()
    
    If Val(txtDari.Text) >= Val(txtSampai.Text) Then
        MsgBox "Range jam salah !", vbCritical
        Exit Sub
    ElseIf Val(txtBiaya.Text) = 0 Then
        MsgBox "Biaya parkir belum diisi !", vbCritical
        Exit Sub
    Else
        
        'Cek apakah ada yang sama
        With Tbl
            For I = 1 To .Rows - 1
                If Val(txtDari.Text) >= Val(.TextMatrix(I, 0)) And _
                   Val(txtSampai.Text) <= Val(.TextMatrix(I, 1)) Then
                   
                   MsgBox "Range jam sudah termasuk dalam range [" & _
                          .TextMatrix(I, 0) & "jam - " & _
                          .TextMatrix(I, 1) & "jam].", vbCritical
                   Exit Sub
                   
                End If
            Next I
        End With
        
    End If
    
    'Tambah
    Tbl.AddItem txtDari.Text & vbTab & txtSampai.Text & vbTab & Format(txtBiaya.Text, "#,##0")
    
End Sub

Private Sub cmdKeluar_Click()
    
    Unload Me
    
End Sub

Private Sub cmdSimpan_Click()
    
    If Option1.Value Then
        
        'BIAYA STATIS
        
        If Val(txtBiayaStatis) = 0 Then
            MsgBox "Biaya parkir belum diisi !", vbCritical
            Exit Sub
        End If
        
        'Tipe biaya
        S = "Update SettingID Set SetValue = '1' " & _
            "Where SetID = 'Tipe Biaya'"
        oConn.Execute S
        
        'Besar biaya
        S = "Update SettingID Set SetValue = '" & _
            Format(txtBiayaStatis, "0") & "' " & _
            "Where SetID = 'Besar Biaya'"
        oConn.Execute S
        
    ElseIf Option2.Value Then
        
        'BIAYA PER JAM
        
        If Val(txtBiayaJam) = 0 Then
            MsgBox "Biaya parkir belum diisi !", vbCritical
            Exit Sub
        End If
        
        'Tipe biaya
        S = "Update SettingID Set SetValue = '2' " & _
            "Where SetID = 'Tipe Biaya'"
        oConn.Execute S
        
        'Besar biaya
        S = "Update SettingID Set SetValue = '" & _
            Format(txtBiayaJam, "0") & "' " & _
            "Where SetID = 'Besar Biaya'"
        oConn.Execute S
        
    Else
        
        'BIAYA RANGE
        
        If Tbl.Rows = 1 Then
            MsgBox "Biaya parkir belum diisi !", vbCritical
            Exit Sub
        End If
        
        'Tipe biaya
        S = "Update SettingID Set SetValue = '3' " & _
            "Where SetID = 'Tipe Biaya'"
        oConn.Execute S
        
        'Besar biaya
        oConn.Execute "Delete From BiayaParkir"
        For I = 1 To Tbl.Rows - 1
            S = "Insert into BiayaParkir(Jam1, Jam2, Biaya) Values (" & _
                Tbl.TextMatrix(I, 0) & "," & Tbl.TextMatrix(I, 1) & "," & _
                Format(Tbl.TextMatrix(I, 2), "0") & ")"
            oConn.Execute S
        Next I
        
    End If
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    With Tbl
        .FixedCols = 0
        .Cols = 3
        .Rows = 1
        
        .ColAlignment(0) = 4
        .FixedAlignment(0) = 4
        .ColWidth(0) = 2000
        .TextMatrix(0, 0) = "Dari Jam"
        
        .ColAlignment(1) = 4
        .FixedAlignment(1) = 4
        .ColWidth(1) = 2000
        .TextMatrix(0, 1) = "Sampai Jam"
        
        .ColAlignment(2) = 4
        .FixedAlignment(2) = 4
        .ColWidth(2) = 2250
        .TextMatrix(0, 2) = "Biaya (Rp.)"
        
    End With
    
    'Load Data
    Dim rs As New ADODB.Recordset
    
    S = "Select SetID, SetValue From SettingID " & _
        "Where SetID = 'Tipe Biaya'"
    rs.Open S, oConn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount > 0 Then
        
        If rs!SetValue = "1" Then
            
            'Opsi-1
            Option1.Value = True
            
            Call CloseRS(rs)
            S = "Select SetID, SetValue From SettingID " & _
                "Where SetID = 'Besar Biaya'"
            rs.Open S, oConn, adOpenStatic, adLockReadOnly, adCmdText
            
            If rs.RecordCount > 0 Then
                txtBiayaStatis = Format(rs!SetValue, "#,##0")
            Else
                txtBiayaStatis = "0"
            End If
            
        ElseIf rs!SetValue = "2" Then
            
            'Opsi-2
            Option2.Value = True
            
            Call CloseRS(rs)
            S = "Select SetID, SetValue From SettingID " & _
                "Where SetID = 'Besar Biaya'"
            rs.Open S, oConn, adOpenStatic, adLockReadOnly, adCmdText
            
            If rs.RecordCount > 0 Then
                txtBiayaJam = Format(rs!SetValue, "#,##0")
            Else
                txtBiayaJam = "0"
            End If
            
        ElseIf rs!SetValue = "3" Then
            
            'Opsi-3
            Option3.Value = True
            
            Call CloseRS(rs)
            S = "Select Jam1, Jam2, Biaya From BiayaParkir " & _
                "Order by Jam1"
            rs.Open S, oConn, adOpenStatic, adLockReadOnly, adCmdText
            
            Tbl.Rows = 1
            
            For I = 1 To rs.RecordCount
                Tbl.AddItem rs!Jam1 & vbTab & _
                            rs!Jam2 & vbTab & _
                            Format(rs!Biaya, "#,##0")
                rs.MoveNext
            Next I
            
        End If
        
    End If
    
    Call CloseRS(rs)
    
End Sub

Private Sub Option1_Click()
    
    txtBiayaStatis.Enabled = True
    txtBiayaJam.Enabled = False
    txtDari.Enabled = False
    txtSampai.Enabled = False
    txtBiaya.Enabled = False
    cmdInsert.Enabled = False
    cmdDelete.Enabled = False
    
End Sub

Private Sub Option2_Click()
    
    txtBiayaStatis.Enabled = False
    txtBiayaJam.Enabled = True
    txtDari.Enabled = False
    txtSampai.Enabled = False
    txtBiaya.Enabled = False
    cmdInsert.Enabled = False
    cmdDelete.Enabled = False

End Sub

Private Sub Option3_Click()
    
    txtBiayaStatis.Enabled = False
    txtBiayaJam.Enabled = False
    txtDari.Enabled = True
    txtSampai.Enabled = True
    txtBiaya.Enabled = True
    cmdInsert.Enabled = True
    cmdDelete.Enabled = True

End Sub

Private Sub txtBiayaStatis_GotFocus()
    
    Call GotFocus(txtBiayaStatis)
    
End Sub

Private Sub txtBiayaStatis_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> vbKeyBack And IsNumeric(Chr(KeyAscii)) = False Then
    
       KeyAscii = 0
       
    End If
    
End Sub

Private Sub txtBiayaStatis_LostFocus()
    
    Call LostFocus(txtBiayaStatis)
    
End Sub

Private Sub txtBiaya_GotFocus()
    
    Call GotFocus(txtBiaya)
    
End Sub

Private Sub txtBiaya_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> vbKeyBack And IsNumeric(Chr(KeyAscii)) = False Then
    
       KeyAscii = 0
       
    End If
    
End Sub

Private Sub txtBiaya_LostFocus()
    
    Call LostFocus(txtBiaya)
    
End Sub

Private Sub txtBiayaJam_GotFocus()
    
    Call GotFocus(txtBiayaJam)
    
End Sub

Private Sub txtBiayaJam_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> vbKeyBack And IsNumeric(Chr(KeyAscii)) = False Then
    
       KeyAscii = 0
       
    End If
    
End Sub

Private Sub txtBiayaJam_LostFocus()
    
    Call LostFocus(txtBiayaJam)
    
End Sub

Private Sub txtDari_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> vbKeyBack And IsNumeric(Chr(KeyAscii)) = False Then
    
       KeyAscii = 0
       
    End If
    
End Sub

Private Sub txtSampai_KeyPress(KeyAscii As Integer)
    
    If KeyAscii <> vbKeyBack And IsNumeric(Chr(KeyAscii)) = False Then
    
       KeyAscii = 0
       
    End If
    
End Sub

