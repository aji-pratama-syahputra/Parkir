VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLapSisaParkir1 
   BackColor       =   &H00E6F1EF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Daftar Sisa Kendaraan"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7125
   Icon            =   "frmLapSisaParkir1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   7125
   Begin VB.CommandButton cmdKeluar 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&KELUAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4335
      TabIndex        =   2
      Top             =   3525
      Width           =   1650
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2640
      TabIndex        =   1
      Top             =   3525
      Width           =   1650
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E6F1EF&
      Caption         =   " TANGGAL "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   1080
      TabIndex        =   6
      Top             =   1635
      Width           =   4905
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   1560
         TabIndex        =   0
         Top             =   450
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   19726339
         CurrentDate     =   38985
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H002A5CE4&
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   -15
      ScaleHeight     =   600
      ScaleWidth      =   10020
      TabIndex        =   4
      Top             =   0
      Width           =   10050
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Laporan daftar sisa kendaraan"
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
         TabIndex        =   5
         Top             =   30
         Width           =   4830
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H002A5CE4&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   -15
      ScaleHeight     =   255
      ScaleWidth      =   10020
      TabIndex        =   3
      Top             =   4650
      Width           =   10050
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   210
      Picture         =   "frmLapSisaParkir1.frx":2CFA
      Top             =   3840
      Width           =   1800
   End
End
Attribute VB_Name = "frmLapSisaParkir1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdKeluar_Click()

    Unload Me
    
End Sub

Private Sub cmdOK_Click()
    
    Tgl1 = DTPicker1.Value
    
    Unload Me
    Unload frmLapSisaParkir2
    frmLapSisaParkir2.Show
    
End Sub

Private Sub Form_Load()
    
    DTPicker1.Value = Date
    
End Sub
