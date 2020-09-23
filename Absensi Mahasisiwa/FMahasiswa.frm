VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FMahasiswa 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form4"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7875
   ControlBox      =   0   'False
   Icon            =   "FMahasiswa.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   ScaleHeight     =   7635
   ScaleWidth      =   7875
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   3720
      Top             =   2640
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
      Left            =   3720
      Picture         =   "FMahasiswa.frx":E2E1
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6720
      Width           =   1215
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
      Left            =   2520
      Picture         =   "FMahasiswa.frx":E46B
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6720
      Width           =   1215
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
      Picture         =   "FMahasiswa.frx":EAD5
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6720
      Width           =   1215
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
      Picture         =   "FMahasiswa.frx":F13F
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6720
      Width           =   1095
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
      Left            =   5280
      Picture         =   "FMahasiswa.frx":F7A9
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6720
      Width           =   2415
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
      Height          =   735
      Left            =   7080
      TabIndex        =   17
      Top             =   1680
      Width           =   615
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
      Height          =   735
      Left            =   6480
      TabIndex        =   16
      Top             =   1680
      Width           =   615
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
      Height          =   735
      Left            =   5880
      TabIndex        =   15
      Top             =   1680
      Width           =   615
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
      Height          =   735
      Left            =   5280
      TabIndex        =   14
      Top             =   1680
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6840
      Top             =   840
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4935
      Begin VB.TextBox TXTCLASS 
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox TXTJURUSAN 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox TXTNAMA 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox TXTNPM 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   1575
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
         TabIndex        =   6
         Top             =   600
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Npm"
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
         TabIndex        =   5
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class"
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
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jurusan"
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
         Top             =   960
         Width           =   750
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3015
      Left            =   240
      TabIndex        =   13
      Top             =   3360
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      BackColor       =   -2147483644
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
      TabIndex        =   24
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
      TabIndex        =   23
      Top             =   0
      Width           =   1890
   End
   Begin VB.Shape Shape4 
      Height          =   975
      Left            =   120
      Top             =   6600
      Width           =   4935
   End
   Begin VB.Shape Shape3 
      Height          =   975
      Left            =   5160
      Top             =   6600
      Width           =   2655
   End
   Begin VB.Shape Shape2 
      Height          =   975
      Left            =   5160
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      Height          =   3255
      Left            =   120
      Top             =   3240
      Width           =   7695
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
      Left            =   4320
      TabIndex        =   10
      Top             =   2880
      Width           =   3375
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
      Left            =   4320
      TabIndex        =   9
      Top             =   2520
      Width           =   1755
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
      Left            =   6240
      TabIndex        =   8
      Top             =   960
      Width           =   555
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
      Left            =   5280
      TabIndex        =   7
      Top             =   960
      Width           =   705
   End
End
Attribute VB_Name = "FMahasiswa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CAHIR_Click()
If RsMahasiswa.RecordCount = 0 Then
    MsgBox "Empty Date In Tabel.....!!!"
    Exit Sub
End If

RsMahasiswa.MoveLast
tampil

CSIMPAN.Enabled = False
CEDIT.Enabled = True
CHAPUS.Enabled = True
Frame1.Enabled = True
TxtNpm.SetFocus
CBARU.Caption = "&New"
End Sub

Private Sub CAWAL_Click()
If RsMahasiswa.RecordCount = 0 Then
    MsgBox "Empty Date In Tabel.....!!!"
    Exit Sub
End If
RsMahasiswa.MoveFirst
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

Private Sub CEDIT_Click()
If RsMahasiswa.RecordCount = 0 Then
    MsgBox "Empty Date In Tabel.....!!!"
Else
On Error GoTo pesan
If TxtNpm = "" Or TxtNama = "" Or TxtJurusan = "" Or TxtClass = "" Then
    MsgBox "Pleace Edit Your Date"
    Frame1.Enabled = True
    TxtNpm.SetFocus
Else
    h = MsgBox("Are You Sure To Edit Date....!!!", vbQuestion + vbYesNo, "Information Aplication")
    If h = vbYes Then
    With RsMahasiswa
        !npm = TxtNpm
        !nama = TxtNama
        !jurusan = TxtJurusan
        !Class = TxtClass
        .Update
    End With
MsgBox "Editing Succesfull"
    bersih
    Frame1.Enabled = True
    TxtNpm.SetFocus
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
If RsMahasiswa.RecordCount = 0 Then
    MsgBox "Empty Date Date.....!!!"
    Exit Sub
Else
    tampil
    Frame1.Enabled = False
    h = MsgBox("Are You Sure To Delet Date....!!!", vbQuestion + vbYesNo, "Warning Aplication")
    If h = vbYes Then
    On Error Resume Next
    With RsMahasiswa
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
If RsMahasiswa.RecordCount = 0 Then
MsgBox "Empty Date In Tabel.....!!!"
Exit Sub
End If
With RsMahasiswa
    RsMahasiswa.MoveNext
End With
If RsMahasiswa.EOF Then
    RsMahasiswa.MoveFirst
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
If RsMahasiswa.RecordCount = 0 Then
MsgBox "Empty Date In Tabel.....!!!"
Exit Sub
End If
With RsMahasiswa
    RsMahasiswa.MovePrevious
 End With
If RsMahasiswa.BOF Then
    RsMahasiswa.MoveLast
End If
tampil
CSIMPAN.Enabled = False
CEDIT.Enabled = True
CHAPUS.Enabled = True
Frame1.Enabled = True
TxtNpm.SetFocus
CBARU.Caption = "&New"

End Sub

Private Sub CSIMPAN_Click()
On Error GoTo pesan
If CSIMPAN.Caption = "&Save" Then
If TxtNpm = "" Or TxtNama = "" Or TxtJurusan = "" Or TxtClass = "" Then
    MsgBox "Pleace Input Complete Date"
    TxtNpm.SetFocus
Else
a = MsgBox("Are You Sure To Save Date.....!!!", vbQuestion + vbYesNo, "Information")
If a = vbYes Then
    With RsMahasiswa
        .AddNew
        !npm = TxtNpm
        !nama = TxtNama
        !jurusan = TxtJurusan
        !Class = TxtClass
        .Update
    End With
MsgBox "Save Date is Succesfull"
    bersih
    TxtNpm.SetFocus
End If
End If
End If
Exit Sub
pesan:
MsgBox "Dont Save Bacause Error Date....'" & TxtNpm & "'....Duplicate"
RsMatakuliah.CancelUpdate
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

Private Sub Form_Activate()
CSIMPAN.Enabled = False
Frame1.Enabled = False
FMahasiswa.Caption = "Form Input Data Mahasiswa"
End Sub

Private Sub Form_Load()
koneksi

RsMahasiswa.Open "select * from tmahasiswa order by npm", XKoneksi, adOpenDynamic, adLockPessimistic
Set DataGrid1.DataSource = RsMahasiswa

DataGrid1.Columns(0).Width = 1500
DataGrid1.Columns(1).Width = 2500
DataGrid1.Columns(2).Width = 2500
DataGrid1.Columns(3).Width = 1500
End Sub

Sub bersih()
TxtNpm = ""
TxtNama = ""
TxtJurusan = ""
TxtClass = ""
End Sub

Sub tampil()
On Error Resume Next
With RsMahasiswa
    TxtNpm.Text = !npm
    TxtNama.Text = !nama
    TxtJurusan.Text = !jurusan
    TxtClass.Text = !Class
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

Private Sub TxtClass_Change()
X = Len(TxtClass)
If KeyAscii = 13 Then
If X > 20 Then
MsgBox "Pleace Input Kode of 20 Digit Only"
End If
End If
End Sub

Private Sub TxtJurusan_KeyPress(KeyAscii As Integer)
X = Len(TxtJurusan)
If KeyAscii = 13 Then
If X > 50 Then
MsgBox "Pleace Input Kode of 50 Digit Only"
Else
TxtClass.SetFocus
End If
End If
End Sub

Private Sub TXTNAMA_KeyPress(KeyAscii As Integer)
X = Len(TxtNama)
If KeyAscii = 13 Then
If X > 50 Then
MsgBox "Pleace Input Kode of 50 Digit Only"
Else
TxtJurusan.SetFocus
End If
End If
End Sub

Private Sub TxtNpm_KeyPress(KeyAscii As Integer)
X = Len(TxtNpm)
If KeyAscii = 13 Then
If X > 10 Then
MsgBox "Pleace Input Kode of 50 Digit Only"
Else
TxtNama.SetFocus
End If
End If
End Sub
