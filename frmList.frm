VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmList 
   BackColor       =   &H00E6F1EF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " List Kendaraan"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12270
   Icon            =   "frmList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   12270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNoPlat 
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
      Left            =   2895
      MaxLength       =   30
      TabIndex        =   1
      Top             =   8355
      Width           =   4905
   End
   Begin VB.CommandButton cmdKeluar 
      BackColor       =   &H00C0E0FF&
      Caption         =   "KELUAR"
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
      Left            =   10065
      TabIndex        =   3
      Top             =   8295
      Width           =   2055
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00C0E0FF&
      Caption         =   "REFRESH"
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
      Left            =   7965
      TabIndex        =   2
      Top             =   8295
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid Tbl 
      Height          =   7305
      Left            =   150
      TabIndex        =   0
      Top             =   900
      Width           =   11970
      _ExtentX        =   21114
      _ExtentY        =   12885
      _Version        =   393216
      BackColorBkg    =   14737632
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
      Caption         =   "CARI NO. PLAT POLISI ="
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   150
      TabIndex        =   6
      Top             =   8415
      Width           =   2580
   End
   Begin VB.Label lblTgl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal: 23 Oktober 2006"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   150
      TabIndex        =   5
      Top             =   615
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   10320
      Picture         =   "frmList.frx":2CFA
      Top             =   105
      Width           =   1800
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List Kendaraan ..."
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   150
      TabIndex        =   4
      Top             =   150
      Width           =   2595
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    Dim rsT As New ADODB.Recordset
    
    If txtNoPlat.Text = "" Then
        
        'Tampilkan semua
        Tbl.Rows = 1
        
        S = " Select * From Parking " & _
            " Where Tanggal = " & FormatTgl(Date) & _
            " Order by JamMasuk"
        rsT.Open S, oConn, adOpenStatic, adLockReadOnly, adCmdText
        For I = 1 To rsT.RecordCount
            Tbl.AddItem I & "." & vbTab & _
                        rsT!NoPlat & vbTab & _
                        Format(rsT!JamMasuk, "hh:nn:ss") & vbTab & _
                        rsT!OpMasuk & vbTab & _
                        Format(rsT!JamKeluar, "hh:nn:ss") & vbTab & _
                        rsT!OpKeluar & vbTab & _
                        Format(rsT!Biaya, "#,###") & vbTab & _
                        rsT!Keterangan
            rsT.MoveNext
        Next I
        Call CloseRS(rsT)
        
    Else
        
        'Pilih
        Tbl.Rows = 1
        'Dim rsT As New ADODB.Recordset
        
        S = " Select * From Parking " & _
            " Where Tanggal = " & FormatTgl(Date) & _
            " And NoPlat like '%" & txtNoPlat.Text & "%' Order by JamMasuk"
        rsT.Open S, oConn, adOpenStatic, adLockReadOnly, adCmdText
        For I = 1 To rsT.RecordCount
            Tbl.AddItem I & "." & vbTab & _
                        rsT!NoPlat & vbTab & _
                        Format(rsT!JamMasuk, "hh:nn:ss") & vbTab & _
                        rsT!OpMasuk & vbTab & _
                        Format(rsT!JamKeluar, "hh:nn:ss") & vbTab & _
                        rsT!OpKeluar & vbTab & _
                        Format(rsT!Biaya, "#,###") & vbTab & _
                        rsT!Keterangan
            rsT.MoveNext
        Next I
        Call CloseRS(rsT)
        
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    
    'Tanggal
    lblTgl = "Tanggal: " & Format(Date, "dd MMMM yyyy")
    
    With Tbl
        .Cols = 8
        .Rows = 1
        
        'Kolom 0 - Nomor
        .TextMatrix(0, 0) = "No."
        .FixedAlignment(0) = 4
        .ColAlignment(0) = 6
        .ColWidth(0) = 700
        
        'Kolom 1 - Nomor plat
        .TextMatrix(0, 1) = "Plat Polisi"
        .FixedAlignment(1) = 4
        .ColAlignment(1) = 1
        .ColWidth(1) = 1500
        
        'Kolom 2 - Jam Masuk
        .TextMatrix(0, 2) = "Jam Masuk"
        .ColAlignment(2) = 4
        .ColWidth(2) = 1200
        
        'Kolom 3 - Op Masuk
        .TextMatrix(0, 3) = "Operator Masuk"
        .ColAlignment(3) = 4
        .ColWidth(3) = 1800
        
        'Kolom 4 - Jam Keluar
        .TextMatrix(0, 4) = "Jam Keluar"
        .ColAlignment(4) = 4
        .ColWidth(4) = 1200
        
        'Kolom 5 - Op Keluar
        .TextMatrix(0, 5) = "Operator Keluar"
        .ColAlignment(5) = 4
        .ColWidth(5) = 1800
        
        'Kolom 6 - Biaya
        .TextMatrix(0, 6) = "Biaya (Rp.)"
        .ColAlignment(6) = 6
        .ColWidth(6) = 1500
        
        'Kolom 7 - Keterangan
        .TextMatrix(0, 7) = "Keterangan"
        .FixedAlignment(7) = 4
        .ColAlignment(7) = 1
        .ColWidth(7) = 2165
        
    End With
    
    Dim rsT As New ADODB.Recordset
    
    S = " Select * From Parking " & _
        " Where Tanggal = " & FormatTgl(Date) & _
        " Order by JamMasuk"
    rsT.Open S, oConn, adOpenStatic, adLockReadOnly, adCmdText
    For I = 1 To rsT.RecordCount
        Tbl.AddItem I & "." & vbTab & _
                    rsT!NoPlat & vbTab & _
                    Format(rsT!JamMasuk, "hh:nn:ss") & vbTab & _
                    rsT!OpMasuk & vbTab & _
                    Format(rsT!JamKeluar, "hh:nn:ss") & vbTab & _
                    rsT!OpKeluar & vbTab & _
                    Format(rsT!Biaya, "#,###") & vbTab & _
                    rsT!Keterangan
        rsT.MoveNext
    Next I
    
    Call CloseRS(rsT)
    
End Sub

Private Sub txtNoPlat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdRefresh.Value = True
End Sub
