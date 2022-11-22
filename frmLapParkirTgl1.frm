VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLapParkirTgl1 
   BackColor       =   &H00E6F1EF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Parkir Per Tanggal"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7125
   Icon            =   "frmLapParkirTgl1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
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
      TabIndex        =   10
      Top             =   5280
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
      TabIndex        =   9
      Top             =   5280
      Width           =   1650
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E6F1EF&
      Caption         =   " "
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
      TabIndex        =   18
      Top             =   3810
      Width           =   4905
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2850
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   450
         Width           =   1680
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   465
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   450
         Width           =   1680
      End
      Begin VB.CheckBox chkOperator 
         BackColor       =   &H00E6F1EF&
         Caption         =   "FILTER OPERATOR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   165
         TabIndex        =   6
         Top             =   -75
         Width           =   2235
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s/d"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2340
         TabIndex        =   19
         Top             =   525
         Width           =   300
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E6F1EF&
      Caption         =   " "
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
      TabIndex        =   16
      Top             =   2490
      Width           =   4905
      Begin VB.CheckBox chkJam 
         BackColor       =   &H00E6F1EF&
         Caption         =   "FILTER JAM MASUK"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   165
         TabIndex        =   3
         Top             =   -75
         Width           =   2280
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   360
         Left            =   465
         TabIndex        =   4
         Top             =   450
         Width           =   1680
         _ExtentX        =   2963
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
         CustomFormat    =   "hh:mm:ss"
         Format          =   19726338
         UpDown          =   -1  'True
         CurrentDate     =   38985
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   360
         Left            =   2850
         TabIndex        =   5
         Top             =   450
         Width           =   1680
         _ExtentX        =   2963
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
         CustomFormat    =   "hh:mm:ss"
         Format          =   19726338
         UpDown          =   -1  'True
         CurrentDate     =   38985
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s/d"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2340
         TabIndex        =   17
         Top             =   525
         Width           =   300
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E6F1EF&
      Caption         =   " "
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
      TabIndex        =   14
      Top             =   1200
      Width           =   4905
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   465
         TabIndex        =   1
         Top             =   450
         Width           =   1680
         _ExtentX        =   2963
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
      Begin VB.CheckBox chkTgl 
         BackColor       =   &H00E6F1EF&
         Caption         =   "FILTER TANGGAL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   165
         TabIndex        =   0
         Top             =   -75
         Value           =   1  'Checked
         Width           =   2115
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   360
         Left            =   2850
         TabIndex        =   2
         Top             =   450
         Width           =   1680
         _ExtentX        =   2963
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s/d"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2340
         TabIndex        =   15
         Top             =   525
         Width           =   300
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
      TabIndex        =   12
      Top             =   0
      Width           =   10050
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Laporan parkir per tanggal"
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
         TabIndex        =   13
         Top             =   30
         Width           =   4455
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
      TabIndex        =   11
      Top             =   6405
      Width           =   10050
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   210
      Picture         =   "frmLapParkirTgl1.frx":2CFA
      Top             =   5595
      Width           =   1800
   End
End
Attribute VB_Name = "frmLapParkirTgl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkJam_Click()

    DTPicker3.Enabled = chkJam.Value
    DTPicker4.Enabled = chkJam.Value
    
End Sub

Private Sub chkOperator_Click()

    Combo1.Enabled = chkOperator.Value
    Combo2.Enabled = chkOperator.Value
    
End Sub

Private Sub chkTgl_Click()

    DTPicker1.Enabled = chkTgl.Value
    DTPicker2.Enabled = chkTgl.Value
    
End Sub

Private Sub RefreshOp()

    Dim RS As New ADODB.Recordset
    
    S = "Select UserName From UserList Order by UserName"
    RS.Open S, oConn, adOpenStatic, adLockReadOnly, adCmdText
    
    Combo1.Clear
    Combo2.Clear
    For i = 1 To RS.RecordCount
        
        Combo1.AddItem RS!UserName
        Combo2.AddItem RS!UserName
        RS.MoveNext
        
    Next i
    
    Call CloseRS(RS)
    
End Sub

Private Sub cmdKeluar_Click()

    Unload Me
    
End Sub

Private Sub cmdOK_Click()
    
    If chkTgl.Value And (DTPicker1.Value > DTPicker2.Value) Then
        MsgBox "Filter tanggal salah !", vbCritical
        Exit Sub
    ElseIf chkJam.Value And (DTPicker3.Value > DTPicker4.Value) Then
        MsgBox "Filter jam salah !", vbCritical
        Exit Sub
    End If
    
    BTgl = chkTgl.Value
    Tgl1 = DTPicker1.Value
    Tgl2 = DTPicker2.Value
    
    BJam = chkJam.Value
    Jam1 = Format(DTPicker3.Value, "hh:nn:ss")
    Jam2 = Format(DTPicker4.Value, "hh:nn:ss")
    
    BOp = chkOperator.Value
    Op1 = Combo1.Text
    Op2 = Combo2.Text
    
    Unload Me
    Unload frmLapParkirTgl2
    frmLapParkirTgl2.Show
    
End Sub

Private Sub Form_Load()
    
    chkTgl.Value = 1
    Call chkTgl_Click
    chkJam.Value = 0
    Call chkJam_Click
    chkOperator.Value = 0
    Call chkOperator_Click
    
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    
    Call RefreshOp
    
End Sub
