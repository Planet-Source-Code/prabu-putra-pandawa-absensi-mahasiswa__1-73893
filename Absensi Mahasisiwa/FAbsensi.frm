VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FAbsensi 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form5"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   ControlBox      =   0   'False
   Icon            =   "FAbsensi.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   8640
      Top             =   120
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   1095
      Left            =   120
      TabIndex        =   38
      Top             =   6840
      Width           =   4935
      Begin VB.CommandButton CKELUAR 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
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
         Left            =   3480
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FAbsensi.frx":E2E1
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CHAPUS 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "&Delete"
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
         Left            =   2640
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FAbsensi.frx":E46B
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CEDIT 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "&Edit"
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
         Left            =   1800
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FAbsensi.frx":E5F5
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CSIMPAN 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "&Save"
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
         Left            =   960
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FAbsensi.frx":EC5F
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CBARU 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         Caption         =   "&New"
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
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FAbsensi.frx":F2C9
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cari Data Absensi Mahasiswa"
      Height          =   3135
      Left            =   5280
      TabIndex        =   31
      Top             =   4800
      Width           =   6495
      Begin VB.CommandButton CAHIR 
         Appearance      =   0  'Flat
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5880
         TabIndex        =   36
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton CMAJU 
         Appearance      =   0  'Flat
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5400
         TabIndex        =   35
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton CMUNDUR 
         Appearance      =   0  'Flat
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   34
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton CAWAL 
         Appearance      =   0  'Flat
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4440
         TabIndex        =   33
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox TxtCari 
         Height          =   525
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   3495
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2055
         Left            =   120
         TabIndex        =   37
         Top             =   960
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   3625
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
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cari Data Mahasiswa"
      Height          =   3135
      Left            =   5280
      TabIndex        =   24
      Top             =   1560
      Width           =   6495
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5880
         TabIndex        =   29
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5400
         TabIndex        =   28
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   27
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4440
         TabIndex        =   26
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox TxtCari2 
         Height          =   525
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   3375
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   2055
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   3625
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
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Input Data Absensi Mahasiswa"
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   4935
      Begin VB.TextBox TxtNpm 
         Height          =   285
         Left            =   1560
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TxtNama 
         Height          =   285
         Left            =   1560
         TabIndex        =   12
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox TxtJurusan 
         Height          =   285
         Left            =   1560
         TabIndex        =   11
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox TxtClass 
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox ComboKoMtk 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   2295
      End
      Begin VB.ComboBox ComboNip 
         Height          =   315
         Left            =   2520
         TabIndex        =   8
         Top             =   1920
         Width           =   2295
      End
      Begin VB.ComboBox ComboKeterangan 
         Height          =   315
         ItemData        =   "FAbsensi.frx":F933
         Left            =   120
         List            =   "FAbsensi.frx":F935
         TabIndex        =   7
         Top             =   3120
         Width           =   2295
      End
      Begin VB.ComboBox Combohari 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   2520
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2520
         TabIndex        =   14
         Top             =   2520
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16187393
         CurrentDate     =   39089
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jurusan         :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   1350
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class             :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Npm              :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama            :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nip  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2520
         TabIndex        =   19
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ko Matakuliah :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   1485
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2520
         TabIndex        =   17
         Top             =   2280
         Width           =   900
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hari :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   510
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   2880
         Width           =   1230
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   705
      ScaleWidth      =   4905
      TabIndex        =   2
      Top             =   5280
      Width           =   4935
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Simpan Data"
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
         Left            =   1560
         TabIndex        =   4
         Top             =   0
         Width           =   1755
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dengan Benar Dan Teliti"
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
         Left            =   1560
         TabIndex        =   3
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   4680
      Top             =   6480
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
      Left            =   120
      TabIndex        =   45
      Top             =   600
      Width           =   4755
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
      Left            =   120
      TabIndex        =   44
      Top             =   240
      Width           =   2490
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
      Height          =   330
      Left            =   9120
      TabIndex        =   1
      Top             =   120
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
      Height          =   330
      Left            =   10080
      TabIndex        =   0
      Top             =   120
      Width           =   690
   End
End
Attribute VB_Name = "FAbsensi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CAHIR_Click()
If RsAbsensi.RecordCount = 0 Then
    MsgBox "Empty Date In Table.....!!!"
    Exit Sub
End If

RsAbsensi.MoveLast
tampil

CSIMPAN.Enabled = False
CEDIT.Enabled = True
CHAPUS.Enabled = True
Frame1.Enabled = True
TxtNpm.SetFocus
CBARU.Caption = "&New"

End Sub

Private Sub CAWAL_Click()
If RsAbsensi.RecordCount = 0 Then
    MsgBox "Empty Date In Table.....!!!"
    Exit Sub
End If
RsAbsensi.MoveFirst
tampil

CSIMPAN.Enabled = False
CEDIT.Enabled = True
CHAPUS.Enabled = True
Frame1.Enabled = True
TxtNpm.SetFocus
CBARU.Caption = "&New"

End Sub

Private Sub CBARU_Click()
If CBARU.Caption = "&New" Then
    CBARU.Caption = "&Cancel"
    CSIMPAN.Enabled = True
    CEDIT.Enabled = False
    CHAPUS.Enabled = False
    Frame1.Enabled = True
    TxtNpm.SetFocus
    bersih
    Else
        CBARU.Caption = "&New"
        CSIMPAN.Caption = "&Save"
        Frame1.Enabled = False
        CSIMPAN.Enabled = False
        CHAPUS.Enabled = True
        CEDIT.Enabled = True
End If
End Sub

Private Sub CCari2_Click()
If TxtCari2 = "" Then
MsgBox "Pleace Input Your npm"
TxtCari2.SetFocus
Else
On Error Resume Next
RsMahasiswa.Close
RsMahasiswa.Open "select *from Tmahasiswa where npm='" & TxtCari2.Text
         
Set DataGrid1.DataSource = RsMahasiswa
    'With RsMahasiswa
        '.MoveFirst
        '.Find "npm='" & TxtCari2.Text + "'"
        
        DataGrid1.Columns(0).Width = 1000
        DataGrid1.Columns(1).Width = 2000
        DataGrid1.Columns(2).Width = 2000
        DataGrid1.Columns(3).Width = 1000
           ' If .EOF = True Then
          
            '    Exit Sub
           ' End If
   ' End With
End If
End Sub

Private Sub CEDIT_Click()
If RsAbsensi.RecordCount = 0 Then
MsgBox "Empty Date In Table.....!!!"
Else
If TxtNpm = "" Or TxtNama = "" Or TxtJurusan = "" Or TxtClass = "" Or ComboKoMtk = "" Or ComboNip = "" Or Combohari = "" Or DTPicker1 = "" Or ComboKeterangan = "" Then
    MsgBox "Pleace Input Complete Date"
    Frame1.Enabled = True
    TxtNpm.SetFocus
Else
    h = MsgBox("Are You Sure To Edit Date....!!!", vbQuestion + vbYesNo, "Information Aplication")
    If h = vbYes Then
    With RsAbsensi
        !npm = TxtNpm
        !nama = TxtNama
        !Class = TxtClass
        !jurusan = TxtJurusan
        !ko_matakuliah = Left(ComboKoMtk.Text, 6)
        !nip = Left(ComboNip.Text, 9)
        !hari = Combohari
        !tanggal = DTPicker1
        !keterangan = ComboKeterangan
        .Update
    End With
MsgBox "Editing Succesfull"
    bersih
    TxtNpm.SetFocus
End If
End If
End If
End Sub

Private Sub CHAPUS_Click()
If RsAbsensi.RecordCount = 0 Then
    MsgBox "Empty Date.....!!!"
    Exit Sub
Else
    Frame1.Enabled = False
    tampil
    h = MsgBox("Are You Sure To Delete Date....!!!", vbQuestion + vbYesNo, "Warning Aplication")
    If h = vbYes Then
    On Error Resume Next
    With RsAbsensi
        .Delete
        .MoveFirst
    End With
bersih
Frame1.Enabled = False
CBARU.Enabled = True
End If
End If

End Sub

Private Sub CKELUAR_Click()
Unload Me
keluar
End Sub

Private Sub CMAJU_Click()
If RsAbsensi.RecordCount = 0 Then
MsgBox "Empty Date In Table.....!!!"
Exit Sub
End If
With RsAbsensi
    RsAbsensi.MoveNext
End With
If RsAbsensi.EOF Then
    RsAbsensi.MoveFirst
End If
tampil
CEDIT.Enabled = True
CHAPUS.Enabled = True
CSIMPAN.Enabled = False
Frame1.Enabled = True
TxtNpm.SetFocus
CBARU.Caption = "&New"

End Sub

Private Sub CMUNDUR_Click()
If RsAbsensi.RecordCount = 0 Then
MsgBox "Empty Date In Table.....!!!"
Exit Sub
End If
With RsAbsensi
    RsAbsensi.MovePrevious
 End With
If RsAbsensi.BOF Then
    RsAbsensi.MoveLast
End If
tampil
CSIMPAN.Enabled = False
CEDIT.Enabled = True
CHAPUS.Enabled = True
Frame1.Enabled = True
TxtNpm.SetFocus
CBARU.Caption = "&New"

End Sub

Private Sub Command1_Click()
If TxtCari = "" Then
MsgBox "Pleace Input Your npm"
TxtCari.SetFocus
Else
On Error Resume Next
RsAbsensi.Close
RsAbsensi.Open ("select * from Tabsensi where npm ='") & TxtCari.Text

Set DataGrid1.DataSource = RsAbsensi
End If
End Sub

Private Sub Command3_Click()
If RsMahasiswa.RecordCount = 0 Then
    MsgBox "Empty Date In Table.....!!!"
    Exit Sub
End If

RsMahasiswa.MoveLast
tampil_mahasiswa
bersih2
CSIMPAN.Enabled = False
CEDIT.Enabled = True
CHAPUS.Enabled = True
Frame1.Enabled = True
TxtNpm.SetFocus
CBARU.Caption = "&New"

End Sub

Private Sub Command4_Click()
If RsMahasiswa.RecordCount = 0 Then
MsgBox "Empty Date In Table.....!!!"
Exit Sub
End If
With RsMahasiswa
    RsMahasiswa.MoveNext
End With
If RsMahasiswa.EOF Then
    RsMahasiswa.MoveFirst
End If
bersih2
tampil_mahasiswa
CEDIT.Enabled = True
CHAPUS.Enabled = True
CSIMPAN.Enabled = False
Frame1.Enabled = True
TxtNpm.SetFocus
CBARU.Caption = "&New"

End Sub

Private Sub Command5_Click()
If RsMahasiswa.RecordCount = 0 Then
MsgBox "Empty Date In Table.....!!!"
Exit Sub
End If
With RsMahasiswa
    RsMahasiswa.MovePrevious
 End With
If RsMahasiswa.BOF Then
    RsMahasiswa.MoveLast
End If
bersih2
tampil_mahasiswa
CSIMPAN.Enabled = False
CEDIT.Enabled = True
CHAPUS.Enabled = True
Frame1.Enabled = True
TxtNpm.SetFocus
CBARU.Caption = "&New"

End Sub

Private Sub Command6_Click()
If RsMahasiswa.RecordCount = 0 Then
    MsgBox "Empty Date In Table.....!!!"
    Exit Sub
End If
RsMahasiswa.MoveFirst
tampil_mahasiswa
bersih2
CSIMPAN.Enabled = False
CEDIT.Enabled = True
CHAPUS.Enabled = True
Frame1.Enabled = True

CBARU.Caption = "&New"

End Sub

Private Sub Command7_Click()
tampil_mahasiswa
bersih2
Frame1.Enabled = True
CSIMPAN.Enabled = True
CEDIT.Enabled = False
CHAPUS.Enabled = False
CBARU.Caption = "&New"
End Sub

Private Sub CSIMPAN_Click()
On Error GoTo pesan
If TxtNpm = "" Or TxtNama = "" Or TxtJurusan = "" Or TxtClass = "" Or ComboKoMtk = "" Or ComboNip = "" Or Combohari = "" Or DTPicker1 = "" Or ComboKeterangan = "" Then
    MsgBox "Pleace Input Complete Date"
Else
a = MsgBox("Are You Sure To Save Date", vbQuestion + vbYesNo, "Information")
If a = vbYes Then
    With RsAbsensi
        .AddNew
        !npm = TxtNpm
        !nama = TxtNama
        !Class = TxtClass
        !jurusan = TxtJurusan
        !ko_matakuliah = Left(ComboKoMtk.Text, 6)
        !nip = Left(ComboNip.Text, 9)
        !hari = Combohari
        !tanggal = DTPicker1
        !keterangan = ComboKeterangan
        .Update
    End With
MsgBox "Save Date Is Succesfull"
    bersih
    TxtNpm.SetFocus
End If
End If
Exit Sub
pesan:
MsgBox "Dont Save Bacause Error Date....'" & TxtNpm & "'....Duplicate"
RsAbsensi.CancelUpdate
tampil
End Sub

Private Sub DataGrid1_Click()
tampil
CSIMPAN.Enabled = False
CEDIT.Enabled = True
CHAPUS.Enabled = True
Frame1.Enabled = True
TxtNpm.SetFocus
CBARU.Caption = "&New"
End Sub

Private Sub DataGrid2_Click()
tampil_mahasiswa
bersih2
Frame1.Enabled = True
CSIMPAN.Enabled = True
CEDIT.Enabled = False
CHAPUS.Enabled = False
CBARU.Caption = "&New"
End Sub

Private Sub Form_Activate()
FAbsensi.Caption = "Form Absensi Mahasiswa "
CSIMPAN.Enabled = False
Frame1.Enabled = False
bersih

ComboKeterangan.AddItem "Hadir"
ComboKeterangan.AddItem "Tidak Hadir"
Combohari.AddItem "Senin"
Combohari.AddItem "Selasa"
Combohari.AddItem "Rabu"
Combohari.AddItem "Kamis"
Combohari.AddItem "Jum'at"
Combohari.AddItem "Sabtu"
Combohari.AddItem "Minggu"
End Sub

Private Sub Form_Load()
koneksi

RsAbsensi.Open "SELECT * FROM TABSENSI ", XKoneksi, adOpenDynamic, adLockPessimistic
Set DataGrid1.DataSource = RsAbsensi

RsMahasiswa.Open "select * from tmahasiswa order by npm", XKoneksi, adOpenDynamic, adLockPessimistic
Set DataGrid2.DataSource = RsMahasiswa

DataGrid2.Columns(0).Width = 1500
DataGrid2.Columns(1).Width = 2500
DataGrid2.Columns(2).Width = 2500
DataGrid2.Columns(3).Width = 1500

DataGrid1.Columns(0).Width = 1500
DataGrid1.Columns(1).Width = 2500
DataGrid1.Columns(3).Width = 2700
DataGrid1.Columns(4).Width = 1700
DataGrid1.Columns(5).Width = 1500
DataGrid1.Columns(6).Width = 1500
DataGrid1.Columns(7).Width = 1500
DataGrid1.Columns(8).Width = 1700
End Sub

Sub bersih()
On Error Resume Next
TxtNpm = ""
TxtNama = ""
TxtJurusan = ""
TxtClass = ""
ComboKoMtk.Clear
ComboNip.Clear
Combohari = ""
ComboKeterangan = ""
        With RsMatakuliah
            .Close
            .Open "Select * From Tmatakuliah Order  By ko_matakuliah", XKoneksi, adOpenDynamic, adLockPessimistic
            .MoveFirst
            Do
                 ComboKoMtk.AddItem (!ko_matakuliah + " - " + !nama_matakuliah)
                .MoveNext
            Loop Until .EOF
        End With
        
         With RsDosen
            .Close
            .Open "Select * From Tdosen Order  By nip", XKoneksi, adOpenDynamic, adLockPessimistic
            .MoveFirst
            Do
                 ComboNip.AddItem (!nip + " - " + !nama)
                .MoveNext
            Loop Until .EOF
        End With
End Sub

Sub bersih2()
ComboKoMtk = ""
ComboNip = ""
Combohari = ""
ComboKeterangan = ""
End Sub

Sub tampil()
On Error Resume Next
With RsAbsensi
    TxtNpm.Text = !npm
    TxtNama.Text = !nama
    TxtJurusan.Text = !jurusan
    TxtClass.Text = !Class
    ComboKoMtk.Text = !ko_matakuliah
    ComboNip.Text = !nip
    Combohari.Text = !hari
    DTPicker1 = !tanggal
    ComboKeterangan = !keterangan
End With
End Sub

Sub tampil_mahasiswa()
On Error Resume Next
With RsMahasiswa
    TxtNpm.Text = !npm
    TxtNama.Text = !nama
    TxtJurusan.Text = !jurusan
    TxtClass.Text = !Class
End With
End Sub

Private Sub Timer1_Timer()
If Label13.Caption <> Str(Time) Then Label14.Caption = Str(Time)
End Sub

Private Sub Timer2_Timer()
If Label15.Left = Label15.Left - 90 Or Label16.Left = Label16.Left - 90 Then
    Timer1.Enabled = True
    Else
    If Label15.Left > 0 Then
        Label15.Left = Label15.Left - 90
        Label16.Left = Label16.Left - 90
    Else
    Label15.Left = FMenu.Width
    Label16.Left = FMenu.Width
End If
End If
Label15.ForeColor = QBColor(Rnd * 15)
Label16.ForeColor = QBColor(Rnd * 15)
End Sub


Private Sub TxtCari_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
Dim a As String
a = TxtCari.Text
RsAbsensi.Open "select *from tabsensi where npm='" & a & "'  ", Dataabsensi, adOpenStatic, adLockOptimistic
    With RsAbsensi
        Set DataGrid1.DataSource = RsAbsensi
        .MoveFirst
        .Find "npm='" & a & "'"
            TxtNpm = !npm
            TxtNama = !nama
            TxtClass = !Class
            TxtJurusan = !jurusan
            ComboKoMtk = !ko_matakuliah
            ComboNip = !nip
            Combohari = !hari
            DTPicker1 = !tanggal
            ComboKeterangan = !keterangan
        Set DataGrid1.DataSource = RsAbsensi
        DataGrid1.Columns(0).Width = 1000
        DataGrid1.Columns(1).Width = 2000
        DataGrid1.Columns(2).Width = 1000
        DataGrid1.Columns(3).Width = 2000
        DataGrid1.Columns(4).Width = 2000
        DataGrid1.Columns(5).Width = 2000
        DataGrid1.Columns(6).Width = 2000
        DataGrid1.Columns(7).Width = 2000
            If .EOF = True Then
          
                Exit Sub
            End If
        End With
End If
End Sub

Private Sub TxtCari2_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
Dim a As String
a = TxtCari2
With RsMahasiswa
    .MoveFirst
    .Find "npm='" & a & "'"
    TxtNpm = !npm
    TxtNama = !nama
    TxtClass = !Class
    TxtJurusan = !jurusan
    If .EOF = True Then
    MsgBox "Data '" & a & "' Not Review"
    TxtCari2 = ""
    Exit Sub
    End If
End With
Frame1.Enabled = True
TxtNpm.SetFocus
bersih2
Else
    Frame1.Enabled = False
End If
End Sub

Private Sub TxtClass_KeyPress(KeyAscii As Integer)
X = Len(TxtClass)
If KeyAscii = 13 Then
If X > 50 Then
MsgBox "Pleace Input Kode of 50 Digit Only"
Else
TxtJurusan.SetFocus
End If
End If

End Sub


Private Sub TxtJurusan_KeyPress(KeyAscii As Integer)
X = Len(TxtJurusan)
If KeyAscii = 13 Then
If X > 50 Then
MsgBox "Pleace Input Kode of 50 Digit Only"
Else
ComboKoMtk.SetFocus
End If
End If
End Sub

Private Sub TXTNAMA_KeyPress(KeyAscii As Integer)
X = Len(TxtNama)
If KeyAscii = 13 Then
If X > 50 Then
MsgBox "Pleace Input Kode of 50 Digit Only"
Else
TxtClass.SetFocus
End If
End If
End Sub

Private Sub TxtNpm_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
Dim a As String
a = TxtNpm
With RsMahasiswa
    .MoveFirst
    .Find "npm='" & a & "'"
    TxtNpm = !npm
    TxtNama = !nama
    TxtClass = !Class
    TxtJurusan = !jurusan
    If .EOF = True Then
    MsgBox "Data '" & a & "' Not Review"
    TxtNpm = ""
    Exit Sub
    End If
End With
End If

End Sub



