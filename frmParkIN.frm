VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form frmParkIN 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " GRAND ANGKASA - Secure Parking  [Parkir Masuk]"
   ClientHeight    =   9960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10020
   Icon            =   "frmParkIN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9960
   ScaleWidth      =   10020
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H002A5CE4&
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   -15
      ScaleHeight     =   300
      ScaleWidth      =   10020
      TabIndex        =   15
      Top             =   9645
      Width           =   10050
      Begin VB.PictureBox picScroll 
         Appearance      =   0  'Flat
         BackColor       =   &H002A5CE4&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   240
         ScaleHeight     =   285
         ScaleWidth      =   9990
         TabIndex        =   31
         Top             =   30
         Width           =   9990
         Begin VB.Label lblScroll 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GRAND ANGKASA - Secure Parking [Parkir Masuk], Operator: YULIANA, 23 September 2006"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Left            =   90
            TabIndex        =   32
            Top             =   15
            Width           =   8835
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H002A5CE4&
      ForeColor       =   &H80000008&
      Height          =   1830
      Left            =   -120
      Picture         =   "frmParkIN.frx":2CFA
      ScaleHeight     =   1800
      ScaleWidth      =   10185
      TabIndex        =   9
      Top             =   0
      Width           =   10215
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00000000&
         Height          =   1230
         Left            =   6210
         ScaleHeight     =   1170
         ScaleWidth      =   3780
         TabIndex        =   12
         Top             =   495
         Width           =   3840
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFFFF&
            X1              =   -75
            X2              =   3800
            Y1              =   285
            Y2              =   285
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "W A K T U   S E K A R A N G"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   15
            TabIndex        =   14
            Top             =   30
            Width           =   3750
         End
         Begin VB.Label lblTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "10:55:33"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   45
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   1020
            Left            =   90
            TabIndex        =   13
            Top             =   225
            Width           =   3615
         End
      End
      Begin VB.Label lblTgl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sabtu, 23/09/2006"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6210
         TabIndex        =   10
         Top             =   60
         Width           =   3825
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   9495
      Top             =   1905
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E6F1EF&
      Height          =   7830
      Left            =   -15
      ScaleHeight     =   7770
      ScaleWidth      =   10020
      TabIndex        =   11
      Top             =   1830
      Width           =   10080
      Begin VB.Timer tmrScroll 
         Interval        =   20
         Left            =   195
         Top             =   4890
      End
      Begin VB.CommandButton cmdBaru 
         Caption         =   "&Baru"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2872
         TabIndex        =   8
         Top             =   6330
         Width           =   2175
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00E0E0E0&
         Height          =   1170
         Left            =   4440
         ScaleHeight     =   1110
         ScaleWidth      =   5055
         TabIndex        =   24
         Top             =   4740
         Width           =   5115
         Begin VB.Label lblUser 
            BackStyle       =   0  'Transparent
            Caption         =   "Yuliana"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   3150
            TabIndex        =   30
            Top             =   615
            Width           =   1740
         End
         Begin VB.Label lblSisa 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "100"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   3180
            TabIndex        =   29
            Top             =   128
            Width           =   630
         End
         Begin VB.Shape Shape4 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   390
            Left            =   3090
            Top             =   120
            Width           =   1860
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2910
            TabIndex        =   28
            Top             =   150
            Width           =   105
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama operator"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   120
            TabIndex        =   27
            Top             =   630
            Width           =   2130
         End
         Begin VB.Shape Shape3 
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   390
            Left            =   3090
            Top             =   600
            Width           =   1860
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sisa tempat parkir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   120
            TabIndex        =   26
            Top             =   150
            Width           =   2625
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2910
            TabIndex        =   25
            Top             =   630
            Width           =   105
         End
      End
      Begin VB.CommandButton cmdList 
         Caption         =   "&List Kendaraan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5122
         TabIndex        =   6
         Top             =   7020
         Width           =   2175
      End
      Begin VB.TextBox txtKet 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1665
         Left            =   4440
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   2130
         Width           =   5115
      End
      Begin VB.TextBox txtNoPlat 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   4440
         MaxLength       =   30
         TabIndex        =   0
         Top             =   1305
         Width           =   5115
      End
      Begin VB.TextBox txtJamMasuk 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   3900
         Width           =   5115
      End
      Begin VB.CommandButton cmdKeluar 
         Caption         =   "&Keluar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7365
         TabIndex        =   7
         Top             =   7020
         Width           =   2190
      End
      Begin VB.CommandButton cmdCetak 
         Caption         =   "&Cetak"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2865
         TabIndex        =   5
         Top             =   7020
         Width           =   2175
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "&Hapus"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7380
         TabIndex        =   4
         Top             =   6330
         Width           =   2175
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5122
         TabIndex        =   3
         Top             =   6330
         Width           =   2175
      End
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
         Height          =   2010
         Left            =   15
         TabIndex        =   16
         Top             =   5850
         Width           =   2550
         _cx             =   4498
         _cy             =   3545
         FlashVars       =   ""
         Movie           =   "0"
         Src             =   "0"
         WMode           =   "Window"
         Play            =   -1  'True
         Loop            =   -1  'True
         Quality         =   "High"
         SAlign          =   ""
         Menu            =   0   'False
         Base            =   ""
         AllowScriptAccess=   "always"
         Scale           =   "ShowAll"
         DeviceFont      =   0   'False
         EmbedMovie      =   0   'False
         BGColor         =   ""
         SWRemote        =   ""
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   540
         Left            =   420
         TabIndex        =   23
         Top             =   2220
         Width           =   3420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   540
         Left            =   3885
         TabIndex        =   22
         Top             =   2220
         Width           =   285
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Plat Polisi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   540
         Left            =   420
         TabIndex        =   21
         Top             =   1395
         Width           =   3420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   540
         Left            =   3885
         TabIndex        =   20
         Top             =   1395
         Width           =   285
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Jam Masuk"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   540
         Left            =   420
         TabIndex        =   19
         Top             =   3990
         Width           =   3420
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   540
         Left            =   3885
         TabIndex        =   18
         Top             =   3990
         Width           =   285
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PARKIR MASUK"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00143DA7&
         Height          =   630
         Left            =   0
         TabIndex        =   17
         Top             =   195
         Width           =   9930
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00143DA7&
         BorderWidth     =   3
         X1              =   2580
         X2              =   7395
         Y1              =   855
         Y2              =   855
      End
   End
End
Attribute VB_Name = "frmParkIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdBaru_Click()

    txtNoPlat.Text = ""
    txtKet.Text = ""
    txtJamMasuk.Text = ""
    Call RefreshJlhTempat
    Call RefreshTgl
    txtNoPlat.SetFocus
    
End Sub

Private Sub cmdCetak_Click()
    
    'Periksa sebelumnya ada atau tidak
    Dim rsD As New ADODB.Recordset
    
    S = "Select NoPlat From Parking " & _
        "Where NoPlat = '" & Trim(txtNoPlat) & "' And " & _
        "Tanggal = " & FormatTgl(Date)
    rsD.Open S, oConn, adOpenStatic, adLockReadOnly, adCmdText
    
    If rsD.RecordCount <= 0 Then
        
        Call CloseRS(rsD)
        MsgBox "Tidak ada data untuk dicetak !", vbCritical, "Delete"
        txtNoPlat.SetFocus
        Exit Sub
        
    End If
    
    'Laporan
    cNoPlat = txtNoPlat.Text
    Unload frmCetakBon
    frmCetakBon.Show 1
    
End Sub

Private Sub cmdHapus_Click()
    
    'Periksa sebelumnya ada atau tidak
    Dim rsD As New ADODB.Recordset
    
    S = "Select NoPlat From Parking " & _
        "Where NoPlat = '" & Trim(txtNoPlat) & "' And " & _
        "Tanggal = " & FormatTgl(Date)
    rsD.Open S, oConn, adOpenStatic, adLockReadOnly, adCmdText
    
    If rsD.RecordCount <= 0 Then
        
        Call CloseRS(rsD)
        MsgBox "Tidak ada data untuk dihapus !", vbCritical, "Delete"
        txtNoPlat.SetFocus
        Exit Sub
        
    End If
    
    Call CloseRS(rsD)
        
    'Periksa apakah no. plat sudah keluar atau belum
    If IsNoPlatInUse(Trim(txtNoPlat.Text)) Then
        
        MsgBox "No. Plat Polisi '" & txtNoPlat.Text & "' sudah keluar !", vbCritical
        txtNoPlat.SetFocus
        Exit Sub
        
    End If
        
    'Konfirmasi Hapus
    If MsgBox("Apakah anda yakin untuk menghapus data ini ?", _
              vbQuestion + vbYesNo, "Hapus") = vbNo Then
        Exit Sub
    End If

    On Error GoTo errDelete
    
    'Mulai hapus
    oConn.BeginTrans
    
    'Update
    S = "Update SisaTempat Set Sisa = Sisa + 1 " & _
        "Where Tanggal = " & FormatTgl(Date)
    oConn.Execute S
    
    'Sebelumnya sudah ada, hapus
    S = "Delete From Parking " & _
        "Where NoPlat = '" & Trim(txtNoPlat) & "' And " & _
        "Tanggal = " & FormatTgl(Date)
    oConn.Execute S
    
    'Hapus Selesai
    oConn.CommitTrans
    
    'Kosongkan semua
    cmdBaru.Value = True
    
    Exit Sub

errDelete:
    oConn.RollbackTrans
    MsgBox "Proses Hapus Gagal !" & vbCrLf & Err.Description, vbInformation, "Error"
    
End Sub

Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdList_Click()
    frmList.Show vbModal
End Sub

Private Sub cmdSimpan_Click()
    
    'Periksa input
    If Trim(txtNoPlat.Text) = "" Then
        
        MsgBox "No. Plat Polisi belum diisi !", vbCritical
        txtNoPlat.SetFocus
        Exit Sub
    
    ElseIf IsNoPlatInUse(Trim(txtNoPlat.Text)) Then
        
        MsgBox "No. Plat Polisi '" & txtNoPlat.Text & "' sudah keluar !", vbCritical
        txtNoPlat.SetFocus
        Exit Sub
    
    Else
        
        'Tempat parkir
        Call RefreshJlhTempat
        
        'Tempat parkir sudah habis
        If Val(lblSisa) = 0 Then
            MsgBox "Sisa tempat parkir sudah habis !", vbCritical
            Exit Sub
        End If
        
    End If
    
    'Kalau kosong, ambil waktu sekarang
    If Trim(txtJamMasuk.Text) = "" Then
        txtJamMasuk.Text = lblTime
    End If
    
    On Error GoTo errSave
    
    'Mulai Simpan
    oConn.BeginTrans
    
    'Periksa sebelumnya ada atau tidak
    Dim RS As New ADODB.Recordset
    
    S = "Select NoPlat From Parking " & _
        "Where NoPlat = '" & Trim(txtNoPlat) & "' And " & _
        "Tanggal = " & FormatTgl(Date)
    RS.Open S, oConn, adOpenStatic, adLockReadOnly, adCmdText
    If RS.RecordCount > 0 Then
        
        Call CloseRS(RS)
        
        'Sebelumnya sudah ada, hapus
        S = "Delete From Parking " & _
            "Where NoPlat = '" & Trim(txtNoPlat) & "' And " & _
            "Tanggal = " & FormatTgl(Date)
        oConn.Execute S
        
    Else
        
        Call CloseRS(RS)
        
        'Belum pernah ada, update tempat parkir
        S = "Select Sisa From SisaTempat " & _
        "Where Tanggal = " & FormatTgl(Date)
        RS.Open S, oConn, adOpenStatic, adLockReadOnly, adCmdText
        If RS.RecordCount <= 0 Then
            S = "Insert Into SisaTempat(Tanggal, Sisa) Values " & _
                "(" & FormatTgl(Date) & "," & lblSisa & ")"
            oConn.Execute S
        End If
        
        Call CloseRS(RS)
        
        'Update
        S = "Update SisaTempat Set Sisa = Sisa - 1 " & _
            "Where Tanggal = " & FormatTgl(Date)
        oConn.Execute S
        
    End If
    
    'Simpan ke tabel Parking
    S = "Insert Into Parking(Tanggal, NoPlat, JamMasuk, OpMasuk, " & _
        "JamKeluar, OpKeluar, Biaya, Keterangan) Values (" & _
        FormatTgl(Date) & ",'" & Trim(txtNoPlat.Text) & "',#" & _
        txtJamMasuk.Text & "#,'" & lblUser & "',NULL,NULL,0,'" & _
        Trim(txtKet.Text) & "')"
    oConn.Execute S
    
    'Simpan Selesai
    oConn.CommitTrans
    
    If MsgBox("Cetak bon parkir ?", vbQuestion + vbYesNo, "Cetak") = vbYes Then
        
        'Laporan
        cNoPlat = txtNoPlat.Text
        Unload frmCetakBon
        frmCetakBon.Show 1
        
    End If
    
    'Kosongkan semua
    cmdBaru.Value = True
    
    Exit Sub

errSave:
    oConn.RollbackTrans
    MsgBox "Proses Simpan Gagal !" & vbCrLf & Err.Description, vbInformation, "Error"
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    
    'Flash
    ShockwaveFlash1.Movie = App.Path & "\Inside.swf"
    
    'User
    lblUser = strUser
    
    'Waktu
    lblTime.Caption = Format(Time, "hh:nn:ss")
    txtJamMasuk.Text = ""
    
    'Refresh Tanggal
    Call RefreshTgl
    
    'Jumlah Tempat Parkir
    Call RefreshJlhTempat
    
    'Scroll
    lblScroll = "GRAND ANGKASA - Secure Parking [Operator: " & _
                strUser & "] " & Format(Date, "dd MMMM yyyy")
    
    picScroll.Left = Picture3.Width
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("Apakah anda yakin untuk keluar ?", vbQuestion + vbYesNo, "Keluar") = vbNo Then
        Cancel = True
    Else
        End
    End If
End Sub

Private Sub Timer1_Timer()

    lblTime.Caption = Format(Time, "hh:nn:ss")
    
End Sub

Private Sub RefreshJlhTempat()
    
    Dim rsP As New ADODB.Recordset
    
    'Sisa Tempat Parkir
    S = "Select Sisa From SisaTempat " & _
        "Where Tanggal = " & FormatTgl(Date)
    rsP.Open S, oConn, adOpenStatic, adLockReadOnly, adCmdText
    If rsP.RecordCount > 0 Then
        
        'Sebelumnya sudah ada
        lblSisa = rsP!Sisa
        
    Else
        
        'Sebelumnya tidak ada
        
        Call CloseRS(rsP)
        
        'Ambil dari SettingID
        S = "Select SetValue From SettingID " & _
            "Where SetID = 'Jumlah Tempat Parkir'"
        rsP.Open S, oConn, adOpenStatic, adLockReadOnly, adCmdText
        If rsP.RecordCount > 0 Then
            lblSisa = rsP!SetValue
        Else
            lblSisa = "0"
        End If
        
    End If
    
    Call CloseRS(rsP)
    
End Sub

Private Sub RefreshTgl()

    i = Weekday(Date, vbSunday)
    Select Case i
        Case 1: lblTgl = "Minggu"
        Case 2: lblTgl = "Senin"
        Case 3: lblTgl = "Selasa"
        Case 4: lblTgl = "Rabu"
        Case 5: lblTgl = "Kamis"
        Case 6: lblTgl = "Jum'at"
        Case 7: lblTgl = "Sabtu"
    End Select
    lblTgl = lblTgl & ", " & Format(Date, "dd/mm/yyyy")
    
End Sub

Private Sub tmrScroll_Timer()
    
    picScroll.Left = picScroll.Left - 15
    If picScroll.Left + picScroll.Width + 1000 < Picture3.Left Then
        picScroll.Left = Picture3.Width
    End If
    
End Sub

Private Sub txtNoPlat_Change()
    Dim rsT As New ADODB.Recordset
    Dim strP As String
    
    strP = "Select * From Parking Where " & _
             "NoPlat = '" & Trim(txtNoPlat.Text) & "' " & _
             "And tanggal = " & FormatTgl(Date)
    rsT.Open strP, oConn, adOpenStatic, adLockReadOnly, adCmdText
    If rsT.RecordCount > 0 Then
        txtKet.Text = rsT!Keterangan
        txtJamMasuk.Text = Format(rsT!JamMasuk, "hh:nn:ss")
    Else
        txtKet.Text = ""
        txtJamMasuk.Text = ""
    End If
    
    Call CloseRS(rsT)
End Sub

Private Function IsNoPlatInUse(pcNoPlat) As Boolean
    Dim rsP As New ADODB.Recordset
    Dim strP As String
    
    strP = "Select JamKeluar From Parking " & _
           "Where NoPlat = '" & pcNoPlat & "' AND " & _
           "Tanggal = " & FormatTgl(Date)
    rsP.Open strP, oConn, adOpenStatic, adLockReadOnly, adCmdText
    If rsP.RecordCount > 0 Then
        IsNoPlatInUse = Not IsNull(rsP!JamKeluar)
    Else
        IsNoPlatInUse = False
    End If
    
    Call CloseRS(rsP)
End Function

Private Sub txtNoPlat_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then SendKeys "{TAB}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub


