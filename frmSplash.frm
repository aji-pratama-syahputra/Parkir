VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   " GRAND ANGKASA - Secure Parking"
   ClientHeight    =   4995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7020
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   4995
      Left            =   0
      ScaleHeight     =   4965
      ScaleWidth      =   6990
      TabIndex        =   0
      Top             =   0
      Width           =   7020
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   690
         Top             =   1245
      End
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
         Height          =   4965
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   6990
         _cx             =   12330
         _cy             =   8758
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
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private nWkt As Integer

Private Sub Form_Load()
    
    nWkt = 0
    ShockwaveFlash1.Movie = App.Path & "\Splash.swf"
    
End Sub

Private Sub ShockwaveFlash1_FSCommand(ByVal command As String, ByVal args As String)
    If command = "END" And args = "END" Then
        Timer1.Enabled = True
    End If
End Sub

Private Sub Timer1_Timer()
    
    nWkt = nWkt + 1
    If nWkt = 3 Then
        
        Timer1.Enabled = False
        Unload Me
        
        'Periksa operator
        If strUserType = "Operator IN" Then
            
            'Operator Masuk
            frmParkIN.Show
            
        ElseIf strUserType = "Operator OUT" Then
            
            'Operator Keluar
            frmParkOUT.Show
            
        Else
            
            'Administrator
            frmMain.Show
            
        End If
        
    End If
    
End Sub
