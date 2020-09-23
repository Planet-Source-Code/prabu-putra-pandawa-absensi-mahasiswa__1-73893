VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FLaporanAbsensi 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7200
   ControlBox      =   0   'False
   Icon            =   "FLaporanAbsensi.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6285
   ScaleWidth      =   7200
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2880
      Top             =   5760
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   3840
      Top             =   5520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton CKELUAR 
      Appearance      =   0  'Flat
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      Picture         =   "FLaporanAbsensi.frx":E2E1
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton Cprint 
      Appearance      =   0  'Flat
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      Picture         =   "FLaporanAbsensi.frx":E46B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   9128
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Laporan Absensi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   120
      TabIndex        =   3
      Top             =   5880
      Width           =   2370
   End
   Begin VB.Shape Shape4 
      Height          =   5415
      Left            =   0
      Top             =   0
      Width           =   7095
   End
End
Attribute VB_Name = "FLaporanAbsensi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CHAPUS_Click()

End Sub

Private Sub CKELUAR_Click()
Unload Me
keluar
End Sub

Private Sub Cprint_Click()
With CrystalReport1
        .ReportFileName = App.Path & "\ReportAbsensi.rpt"
        .WindowState = crptMaximized
        .RetrieveDataFiles
        .Action = 1
End With
End Sub

Private Sub Form_Activate()
FLaporanAbsensi.Caption = "Data Absensi Mahasiswa"
End Sub

Private Sub Form_Load()
koneksi

RsAbsensi.Open "SELECT * FROM TABSENSI ", XKoneksi, adOpenDynamic, adLockPessimistic
Set DataGrid1.DataSource = RsAbsensi

DataGrid1.Columns(0).Width = 1500
DataGrid1.Columns(1).Width = 2500
DataGrid1.Columns(2).Width = 1500
DataGrid1.Columns(3).Width = 2700
DataGrid1.Columns(4).Width = 1700
DataGrid1.Columns(5).Width = 1500
DataGrid1.Columns(6).Width = 1500
DataGrid1.Columns(7).Width = 1500
End Sub

Private Sub Timer1_Timer()
Label8.ForeColor = QBColor(Rnd * 15)
End Sub
