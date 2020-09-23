VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FMenu 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   7620
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   14610
   Icon            =   "FMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   14610
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   8040
      Top             =   360
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7245
      Width           =   14610
      _ExtentX        =   25770
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STMIK Ganesha "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   2490
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jl. Kiaracondong 416, Bandung"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   4755
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Left            =   12600
      TabIndex        =   2
      Top             =   360
      Width           =   870
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Left            =   13560
      TabIndex        =   1
      Top             =   360
      Width           =   690
   End
   Begin VB.Menu mnm 
      Caption         =   "&File"
      Begin VB.Menu mna 
         Caption         =   "Absensi Mahasiswa"
      End
      Begin VB.Menu mnline1 
         Caption         =   "-"
      End
      Begin VB.Menu mne 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mno 
      Caption         =   "&Data"
      Begin VB.Menu mnma 
         Caption         =   "Mahasisiwa"
      End
      Begin VB.Menu mnline2 
         Caption         =   "-"
      End
      Begin VB.Menu mnd 
         Caption         =   "Dosen"
      End
      Begin VB.Menu mnline3 
         Caption         =   "-"
      End
      Begin VB.Menu mnmata 
         Caption         =   "Matakuliah"
      End
   End
   Begin VB.Menu mnl 
      Caption         =   "&Laporan"
      Begin VB.Menu mnlapo 
         Caption         =   "Laporan Absensi"
      End
   End
End
Attribute VB_Name = "FMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    FMenu.Caption = "Aplikasi Absensi STMIK GANESHA"
End Sub

Private Sub mna_Click()
    FAbsensi.Show vbModal
End Sub

Private Sub mnd_Click()
    FDosen.Show vbModal
End Sub

Private Sub mnk_Click()
k = MsgBox("Yakin Akan Keluar Aplikasi....?", vbQuestion + vbYesNoCancel, "Instruksi Aplikasi")
If k = vbYes Then
    End
End If
End Sub

Private Sub mne_Click()
p = MsgBox("Keluar Aplikasi ?", vbQuestion + vbYesNoCancel, "Instruction Aplication")
If p = vbYes Then
    End
End If
End Sub

Private Sub mnlapo_Click()
FLaporanAbsensi.Show vbModal
End Sub

Private Sub mnma_Click()
    FMahasiswa.Show vbModal
End Sub

Private Sub mnmata_Click()
    FMatakuliah.Show vbModal
End Sub

Private Sub Timer1_Timer()
If Label13.Caption <> Str(Time) Then Label14.Caption = Str(Time)
End Sub
