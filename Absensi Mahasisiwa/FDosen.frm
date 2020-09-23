VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FDosen 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form2"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7545
   ControlBox      =   0   'False
   Icon            =   "FDosen.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   7545
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   2880
      Top             =   2040
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
      Left            =   4920
      Picture         =   "FDosen.frx":E2E1
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton CHAPUS 
      Appearance      =   0  'Flat
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
      Left            =   3480
      Picture         =   "FDosen.frx":E46B
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton CEDIT 
      Appearance      =   0  'Flat
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
      Left            =   2400
      Picture         =   "FDosen.frx":E5F5
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton CSIMPAN 
      Appearance      =   0  'Flat
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
      Left            =   1320
      Picture         =   "FDosen.frx":EC5F
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton CBARU 
      Appearance      =   0  'Flat
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
      Left            =   240
      Picture         =   "FDosen.frx":F2C9
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6600
      Width           =   1095
   End
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
      Left            =   5760
      TabIndex        =   12
      Top             =   1440
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
      Left            =   5280
      TabIndex        =   11
      Top             =   1440
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
      Left            =   4800
      TabIndex        =   10
      Top             =   1440
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
      Left            =   4320
      TabIndex        =   9
      Top             =   1440
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5880
      Top             =   840
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3975
      Begin VB.TextBox TXTKODE 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TXTNAMA 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nip"
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
         TabIndex        =   4
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
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
         TabIndex        =   3
         Top             =   600
         Width           =   555
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3255
      Left            =   240
      TabIndex        =   13
      Top             =   3000
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5741
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
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
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jl. Kiaracondong 416, Bandung"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   0
      TabIndex        =   20
      Top             =   360
      Width           =   2475
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STMIK Ganesha "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   1890
   End
   Begin VB.Shape Shape4 
      Height          =   975
      Left            =   4800
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Shape Shape3 
      Height          =   975
      Left            =   120
      Top             =   6480
      Width           =   4575
   End
   Begin VB.Shape Shape2 
      Height          =   3495
      Left            =   120
      Top             =   2880
      Width           =   6615
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   4200
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   8
      Top             =   960
      Width           =   705
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5160
      TabIndex        =   7
      Top             =   960
      Width           =   555
   End
   Begin VB.Label Label8 
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
      Left            =   3480
      TabIndex        =   6
      Top             =   2040
      Width           =   1755
   End
   Begin VB.Label Label9 
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
      Left            =   3480
      TabIndex        =   5
      Top             =   2400
      Width           =   3375
   End
End
Attribute VB_Name = "FDosen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CAHIR_Click()
If RsDosen.RecordCount = 0 Then
    MsgBox "Empty Date In Table.....!!!"
    Exit Sub
End If

RsDosen.MoveLast
tampil

CSIMPAN.Enabled = False
CEDIT.Enabled = True
CHAPUS.Enabled = True
Frame1.Enabled = True
TXTKODE.SetFocus
CBARU.Caption = "&New"
End Sub

Private Sub CAWAL_Click()
If RsDosen.RecordCount = 0 Then
    MsgBox "Empty Date In Table.....!!!"
    Exit Sub
End If
RsDosen.MoveFirst
tampil

CSIMPAN.Enabled = False
CEDIT.Enabled = True
CHAPUS.Enabled = True
Frame1.Enabled = True
TXTKODE.SetFocus
CBARU.Caption = "&New"
End Sub

Private Sub CBARU_Click()
If CBARU.Caption = "&New" Then
    CBARU.Caption = "&Cancel"
    CSIMPAN.Enabled = True
    CEDIT.Enabled = False
    CHAPUS.Enabled = False
    Frame1.Enabled = True
    TXTKODE.SetFocus
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

Private Sub CEDIT_Click()
If RsDosen.RecordCount = 0 Then
    MsgBox "Empty Date In Table.....!!!"
Else
On Error GoTo pesan
If TXTKODE = "" Or TxtNama = "" Then
    MsgBox "Pleace Edit Your Date"
    Frame1.Enabled = True
    TXTKODE.SetFocus
Else
    h = MsgBox("Are You Sure To Edit Date....!!!", vbQuestion + vbYesNo, "Information Aplication")
    If h = vbYes Then
    With RsDosen
        !nip = TXTKODE
        !nama = TxtNama
        .Update
    End With
MsgBox "Editing Succesfull"
    bersih
    Frame1.Enabled = True
    TXTKODE.SetFocus
End If
End If
End If
Exit Sub
pesan:
MsgBox "Dont Edit Save Bacause Error Date....'" & TXTKODE & "'....Duplicate"
RsDosen.CancelUpdate
tampil
End Sub

Private Sub CHAPUS_Click()
If RsDosen.RecordCount = 0 Then
    MsgBox "Empty Date In Table.....!!!"
    Exit Sub
Else
    Frame1.Enabled = False
    tampil
    h = MsgBox("Are You Sure To Delet Date....!!!", vbQuestion + vbYesNo, "Warning Aplication")
    If h = vbYes Then
    On Error Resume Next
    With RsDosen
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
If RsDosen.RecordCount = 0 Then
MsgBox "Empty Date In Table.....!!!"
Exit Sub
End If
With RsDosen
    RsDosen.MoveNext
End With
If RsDosen.EOF Then
    RsDosen.MoveFirst
End If
tampil
CEDIT.Enabled = True
CHAPUS.Enabled = True
CSIMPAN.Enabled = False
Frame1.Enabled = True
TXTKODE.SetFocus
CBARU.Caption = "&New"
End Sub

Private Sub CMUNDUR_Click()
If RsDosen.RecordCount = 0 Then
MsgBox "Empty Date In Table.....!!!"
Exit Sub
End If
With RsDosen
    RsDosen.MovePrevious
 End With
If RsDosen.BOF Then
    RsDosen.MoveLast
End If
tampil
CSIMPAN.Enabled = False
CEDIT.Enabled = True
CHAPUS.Enabled = True
Frame1.Enabled = True
TXTKODE.SetFocus
CBARU.Caption = "&New"
End Sub

Private Sub CSIMPAN_Click()
On Error GoTo pesan
If CSIMPAN.Caption = "&Save" Then
If TXTKODE = "" Or TxtNama = "" Then
    MsgBox "Pleace Input Complete Date"
    TXTKODE.SetFocus
Else
a = MsgBox("Are You Sure To Save Date ", vbQuestion + vbYesNo, "Information")
If a = vbYes Then
    With RsDosen
        .AddNew
        !nip = TXTKODE
        !nama = TxtNama
        .Update
    End With
MsgBox "Save Date is Succesfull"
    bersih
    TXTKODE.SetFocus
End If
End If
End If
Exit Sub
pesan:
MsgBox "Dont Save Bacause Error Date....'" & TXTKODE & "'....Duplicate"
RsDosen.CancelUpdate
tampil

End Sub

Private Sub DataGrid1_Click()
tampil
CSIMPAN.Enabled = False
CEDIT.Enabled = True
CHAPUS.Enabled = True
Frame1.Enabled = True
TXTKODE.SetFocus
CBARU.Caption = "&New"
End Sub

Private Sub Form_Activate()
Frame1.Enabled = False
CSIMPAN.Enabled = False
FDosen.Caption = "Form Input Data Dosen"
bersih
End Sub

Private Sub Form_Load()
koneksi

RsDosen.Open "select * from tdosen order by nip", XKoneksi, adOpenDynamic, adLockPessimistic
Set DataGrid1.DataSource = RsDosen

DataGrid1.Columns(0).Width = 2000
DataGrid1.Columns(1).Width = 2500
End Sub

Sub bersih()
TXTKODE = ""
TxtNama = ""
End Sub

Sub tampil()
On Error Resume Next
With RsDosen
    TXTKODE.Text = !nip
    TxtNama.Text = !nama
End With
End Sub

Private Sub Timer1_Timer()
If Label3.Caption <> Str(Time) Then Label7.Caption = Str(Time)
End Sub

Private Sub Timer2_Timer()
If Label8.Left = Label8.Left - 90 Or Label9.Left = Label9.Left - 90 Then
    Timer1.Enabled = True
    Else
    If Label8.Left > 0 Then
        Label8.Left = Label8.Left - 90
        Label9.Left = Label9.Left - 90
    Else
    Label8.Left = FMenu.Width
    Label9.Left = FMenu.Width
End If
End If
Label8.ForeColor = QBColor(Rnd * 15)
Label9.ForeColor = QBColor(Rnd * 15)
End Sub

Private Sub TXTKODE_KeyPress(KeyAscii As Integer)
X = Len(TXTKODE)
If KeyAscii = 13 Then
If X > 10 Then
MsgBox "Pleace Input Kode of 10 Digit Only"
Else
TxtNama.SetFocus
End If
End If
End Sub

Private Sub TXTNAMA_KeyPress(KeyAscii As Integer)
X = Len(TxtNama)
If KeyAscii = 13 Then
If X > 50 Then
MsgBox "Pleace Input Kode of 50 Digit Only"
End If
End If
End Sub
