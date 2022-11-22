VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmCetakBon 
   BackColor       =   &H00E6F1EF&
   Caption         =   "Cetak Bon - Parkir"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10320
   Icon            =   "frmCetakBon.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9315
   ScaleWidth      =   10320
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1410
      Top             =   1995
   End
   Begin CRVIEWERLibCtl.CRViewer CR1 
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5370
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmCetakBon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'REPORT - OBJECT
Private rsCetak As New ADODB.Recordset      'Report Recordset
Private cRep As New CRAXDDRT.Report         'Report Object
Private cApp As New CRAXDDRT.Application    'Report Application

Private Sub Form_Load()
    Dim cF As String
    
    'Sumber rpt
    Set cRep = cApp.OpenReport(App.Path & "\Report\CetakBon.rpt")
            
    'SQL Laporan
    cF = "Select Tanggal, NoPlat, JamMasuk, OpMasuk, " & _
         "JamKeluar, OpKeluar, Biaya, Keterangan From Parking " & _
         "Where Tanggal = " & FormatTgl(Date) & " and NoPlat = '" & cNoPlat & "'"
    rsCetak.Open cF, oConn, adOpenStatic, adLockReadOnly, adCmdText
    If rsCetak.RecordCount > 0 Then
        cRep.Database.SetDataSource rsCetak
        CR1.ReportSource = cRep
        With CR1
            .Left = 0: .Top = 0
            .Width = Me.ScaleWidth
            .Height = Me.ScaleHeight
            .ViewReport
            .Zoom 100
            .Visible = True
        End With
                
    Else
        MsgBox "Tidak ada data untuk dicetak !", vbCritical, "No Data"
        Timer1.Enabled = True
    End If

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState <> vbMinimized Then
        CR1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cApp = Nothing
    Set cRep = Nothing
    Call CloseRS(rsCetak)
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Unload Me
End Sub



