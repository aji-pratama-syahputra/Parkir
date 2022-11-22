VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H00B3C4C7&
   Caption         =   " GRAND ANGKASA - Secure Parking"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10215
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnPengaturan 
      Caption         =   "Pengaturan"
      Begin VB.Menu mnSetBiaya 
         Caption         =   "Setting Biaya Parkir"
      End
      Begin VB.Menu mnSetJlhTmptParkir 
         Caption         =   "Setting Jumlah Tempat Parkir"
      End
      Begin VB.Menu mnP1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSetOperator 
         Caption         =   "Setting Operator"
      End
   End
   Begin VB.Menu mnLaporan 
      Caption         =   "Laporan"
      Begin VB.Menu mnLaporanParkirTgl 
         Caption         =   "Laporan Parkir Per Tanggal"
      End
      Begin VB.Menu mnLaporanParkirBln 
         Caption         =   "Laporan Parkir Per Bulan"
      End
      Begin VB.Menu mnL1 
         Caption         =   "-"
      End
      Begin VB.Menu mnLaporanTerimaTgl 
         Caption         =   "Laporan Penerimaan Per Tanggal"
      End
      Begin VB.Menu mnLaporanTerimaBln 
         Caption         =   "Laporan Penerimaan Per Bulan"
      End
      Begin VB.Menu mnL2 
         Caption         =   "-"
      End
      Begin VB.Menu mnLaporanSisaParkir 
         Caption         =   "Daftar Sisa Kendaraan"
      End
   End
   Begin VB.Menu mnAbout 
      Caption         =   "About"
   End
   Begin VB.Menu mnWindow 
      Caption         =   "Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnTileH 
         Caption         =   "Tile Horizontally"
      End
      Begin VB.Menu mnTileV 
         Caption         =   "Tile Vertically"
      End
      Begin VB.Menu mnCascade 
         Caption         =   "Cascade"
      End
   End
   Begin VB.Menu mnKeluar 
      Caption         =   "Keluar"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub MDIForm_Unload(Cancel As Integer)
    
    If MsgBox("Apakah anda yakin untuk keluar ?", vbQuestion + vbYesNo, "Keluar") = vbNo Then
        Cancel = True
    Else
        End
    End If
    
End Sub

Private Sub mnAbout_Click()
    
    frmAbout.Show vbModal
    
End Sub

Private Sub mnCascade_Click()
    
    Me.Arrange vbCascade
    
End Sub

Private Sub mnKeluar_Click()

    Unload Me
    
End Sub

Private Sub mnLaporanParkirBln_Click()
    
    frmLapParkirBln1.Show
    frmLapParkirBln1.ZOrder 0
    
End Sub

Private Sub mnLaporanParkirTgl_Click()
    
    frmLapParkirTgl1.Show
    frmLapParkirTgl1.ZOrder 0
    
End Sub

Private Sub mnLaporanSisaParkir_Click()
    
    frmLapSisaParkir1.Show
    frmLapSisaParkir1.ZOrder 0
    
End Sub

Private Sub mnLaporanTerimaBln_Click()
    
    frmLapTerimaBln1.Show
    frmLapTerimaBln1.ZOrder 0
    
End Sub

Private Sub mnLaporanTerimaTgl_Click()
    
    frmLapTerimaTgl1.Show
    frmLapTerimaTgl1.ZOrder 0
    
End Sub

Private Sub mnSetBiaya_Click()

    frmSetBiaya.Show
    frmSetBiaya.ZOrder 0
    
End Sub

Private Sub mnSetJlhTmptParkir_Click()
    
    frmSetJlhTmptParkir.Show
    frmSetJlhTmptParkir.ZOrder 0
    
End Sub

Private Sub mnSetOperator_Click()
    
    frmSetUser.Show
    frmSetUser.ZOrder 0
    
End Sub

Private Sub mnTileH_Click()

    Me.Arrange vbTileHorizontal
    
End Sub

Private Sub mnTileV_Click()

    Me.Arrange vbTileVertical
    
End Sub
