VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FMatakuliah 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form3"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10425
   ControlBox      =   0   'False
   Icon            =   "FMatakuliah.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   7095
   ScaleWidth      =   10425
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   6360
      ScaleHeight     =   825
      ScaleWidth      =   3705
      TabIndex        =   21
      Top             =   1800
      Width           =   3735
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
         Left            =   240
         TabIndex        =   23
         Top             =   360
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
         Left            =   240
         TabIndex        =   22
         Top             =   0
         Width           =   1755
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   7560
      Top             =   1200
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
      Height          =   615
      Left            =   7320
      TabIndex        =   20
      Top             =   4920
      Width           =   855
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
      Height          =   615
      Left            =   7320
      TabIndex        =   19
      Top             =   4320
      Width           =   855
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
      Height          =   615
      Left            =   7320
      TabIndex        =   18
      Top             =   3720
      Width           =   855
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
      Height          =   615
      Left            =   7320
      TabIndex        =   17
      Top             =   3120
      Width           =   855
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
      Height          =   615
      Left            =   8520
      Picture         =   "FMatakuliah.frx":E2E1
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4920
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
      Height          =   615
      Left            =   8520
      Picture         =   "FMatakuliah.frx":E46B
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4320
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
      Height          =   615
      Left            =   8520
      Picture         =   "FMatakuliah.frx":EAD5
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3720
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
      Height          =   615
      Left            =   8520
      Picture         =   "FMatakuliah.frx":F13F
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3120
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
      Left            =   7320
      Picture         =   "FMatakuliah.frx":F7A9
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   8160
      Top             =   1200
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   5895
      Begin VB.ComboBox combonip 
         Height          =   315
         Left            =   1920
         TabIndex        =   8
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox txtsks 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtkodematakuliah 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtnamamatakuliah 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label Label5 
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
         TabIndex        =   7
         Top             =   1320
         Width           =   315
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumalah Sks"
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
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Matakuliah"
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
         Width           =   1605
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Matakuliah"
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
         Top             =   600
         Width           =   1665
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3735
      Left            =   240
      TabIndex        =   11
      Top             =   3120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   6588
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      BackColor       =   16777215
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
      TabIndex        =   25
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
      Left            =   120
      TabIndex        =   24
      Top             =   480
      Width           =   4755
   End
   Begin VB.Shape Shape4 
      Height          =   975
      Left            =   7200
      Top             =   6000
      Width           =   2535
   End
   Begin VB.Shape Shape3 
      Height          =   2655
      Left            =   8400
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      Height          =   2655
      Left            =   7200
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      Height          =   3975
      Left            =   120
      Top             =   3000
      Width           =   6975
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
      Left            =   6960
      TabIndex        =   10
      Top             =   1200
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
      Left            =   6120
      TabIndex        =   9
      Top             =   1200
      Width           =   705
   End
End
Attribute VB_Name = "FMatakuliah"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CAHIR_Click()
If RsMatakuliah.RecordCount = 0 Then
    MsgBox "Empty Date In Tabel.....!!!"
    Exit Sub
End If

RsMatakuliah.MoveLast
tampil

CSIMPAN.Enabled = False
CEDIT.Enabled = True
CHAPUS.Enabled = True
Frame4.Enabled = True
txtkodematakuliah.SetFocus
CBARU.Caption = "&New"
End Sub

Private Sub CAWAL_Click()
If RsMatakuliah.RecordCount = 0 Then
    MsgBox "Empty Date In Tabel.....!!!"
    Exit Sub
End If
RsMatakuliah.MoveFirst
tampil

CSIMPAN.Enabled = False
CEDIT.Enabled = True
CHAPUS.Enabled = True
Frame4.Enabled = True
txtkodematakuliah.SetFocus
CBARU.Caption = "&New"
End Sub

Private Sub CBARU_Click()
If CBARU.Caption = "&New" Then
    CBARU.Caption = "&Cancel"
    CSIMPAN.Enabled = True
    CEDIT.Enabled = False
    CHAPUS.Enabled = False
    Frame4.Enabled = True
    txtkodematakuliah.SetFocus
    bersih
    Else
        CBARU.Caption = "&New"
        CSIMPAN.Caption = "&Save"
        Frame4.Enabled = False
        CSIMPAN.Enabled = False
        CHAPUS.Enabled = True
        CEDIT.Enabled = True
End If

End Sub

Private Sub CEDIT_Click()
If RsMatakuliah.RecordCount = 0 Then
MsgBox "Empty Date In Tabel.....!!!"
Else
On Error GoTo pesan
If txtkodematakuliah = "" Or txtnamamatakuliah = "" Or txtsks = "" Or ComboNip = "" Then
    MsgBox "Pleace Edit Your Date"
     Frame4.Enabled = True
     txtkodematakuliah.SetFocus
Else
    h = MsgBox("Are You Sure To Edit Date....!!!", vbQuestion + vbYesNo, "Warning Aplication")
    If h = vbYes Then
        With RsMatakuliah
            !ko_matakuliah = txtkodematakuliah
            !nama_matakuliah = txtnamamatakuliah
            !sks = txtsks
            !nip = Left(ComboNip.Text, 9)
            .Update
        End With
    MsgBox "Editing Succesfull"
        bersih
        Frame4.Enabled = True
        txtkodematakuliah.SetFocus
    End If
End If
End If
Exit Sub
pesan:
MsgBox "Dont Edit Save Bacause Error Date....'" & txtkodematakuliah & "'....Duplicate"
RsMatakuliah.CancelUpdate
tampil
End Sub

Private Sub CHAPUS_Click()
If RsMatakuliah.RecordCount = 0 Then
    MsgBox "Empty Date In Tabel.....!!!"
    Exit Sub
Else
    tampil
    Frame4.Enabled = False
    h = MsgBox("Are You Sure To Delet Date....!!!", vbQuestion + vbYesNo, "Warning Aplication")
    If h = vbYes Then
    On Error Resume Next
    With RsMatakuliah
        .Delete
        .MoveFirst
    End With
bersih
Frame4.Enabled = False
CBARU.Enabled = True
End If
End If
End Sub

Private Sub CKELUAR_Click()
Unload Me
keluar
End Sub

Private Sub CMAJU_Click()
If RsMatakuliah.RecordCount = 0 Then
MsgBox "Empty Date In Tabel.....!!!"
Exit Sub
End If
With RsMatakuliah
    RsMatakuliah.MoveNext
End With
If RsMatakuliah.EOF Then
    RsMatakuliah.MoveFirst
End If
tampil
CEDIT.Enabled = True
CHAPUS.Enabled = True
CSIMPAN.Enabled = False
Frame4.Enabled = True
txtkodematakuliah.SetFocus
CBARU.Caption = "&New"
End Sub

Private Sub CMUNDUR_Click()
If RsMatakuliah.RecordCount = 0 Then
MsgBox "Empty Date In Tabel.....!!!"
Exit Sub
End If
With RsMatakuliah
    RsMatakuliah.MovePrevious
 End With
If RsMatakuliah.BOF Then
    RsMatakuliah.MoveLast
End If
tampil
CSIMPAN.Enabled = False
CEDIT.Enabled = True
CHAPUS.Enabled = True
Frame4.Enabled = True
txtkodematakuliah.SetFocus
CBARU.Caption = "&New"
End Sub

Private Sub CSIMPAN_Click()
On Error GoTo pesan
If CSIMPAN.Caption = "&Save" Then
If txtkodematakuliah = "" Or txtnamamatakuliah = "" Or txtsks = "" Or ComboNip = "" Then
    MsgBox "Pleace Input Complete Date"
    txtkodematakuliah.SetFocus
Else
a = MsgBox("Are You Sure To Save Date.....!!!", vbQuestion + vbYesNo, "Information")
If a = vbYes Then
    With RsMatakuliah
        .AddNew
        !ko_matakuliah = txtkodematakuliah
        !nama_matakuliah = txtnamamatakuliah
        !sks = txtsks
        !nip = Left(ComboNip.Text, 9)
        .Update
    End With
MsgBox "Save Date is Succesfull"
    bersih
    txtkodematakuliah.SetFocus
MsgBox "Save Date Sucses"
End If
End If
End If
Exit Sub
pesan:
MsgBox "Dont Save Bacause Error Date....'" & txtkodematakuliah & "'....Duplicate"
RsMatakuliah.CancelUpdate
tampil
End Sub

Private Sub DataGrid1_Click()
tampil
CSIMPAN.Enabled = False
CEDIT.Enabled = True
CHAPUS.Enabled = True
Frame4.Enabled = True
txtkodematakuliah.SetFocus
CBARU.Caption = "&New"
End Sub


Private Sub Form_Activate()
FMatakuliah.Caption = "Form Input Data Matakuliah"
Frame4.Enabled = False
CSIMPAN.Enabled = False
bersih
End Sub

Private Sub Form_Load()
    koneksi
    
    RsMatakuliah.Open "select * from tmatakuliah ", XKoneksi, adOpenDynamic, adLockPessimistic
    Set DataGrid1.DataSource = RsMatakuliah
    
    DataGrid1.Columns(0).Width = 1300
    DataGrid1.Columns(1).Width = 3000
    DataGrid1.Columns(2).Width = 1200
    DataGrid1.Columns(3).Width = 1500
End Sub


Sub bersih()
On Error Resume Next
txtkodematakuliah = ""
txtnamamatakuliah = ""
txtsks = ""
        ComboNip.Clear
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


Sub tampil()
On Error Resume Next
With RsMatakuliah
    txtkodematakuliah.Text = !ko_matakuliah
    txtnamamatakuliah.Text = !nama_matakuliah
    txtsks.Text = !sks
    ComboNip = !nip
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

Private Sub txtkodematakuliah_KeyPress(KeyAscii As Integer)
X = Len(txtkodematakuliah)
If KeyAscii = 13 Then
If X > 10 Then
MsgBox "Pleace Input Kode of 10 Digit Only"
Else
txtnamamatakuliah.SetFocus
End If
End If
End Sub

Private Sub txtnamamatakuliah_KeyPress(KeyAscii As Integer)
X = Len(txtnamamatakuliah)
If KeyAscii = 13 Then
If X > 50 Then
MsgBox "Pleace Input Kode of 50 Digit Only"
Else
txtsks.SetFocus
End If
End If
End Sub

Private Sub txtsks_KeyPress(KeyAscii As Integer)
X = Len(txtsks)
If KeyAscii = 13 Then
If X > 4 Then
MsgBox "Pleace Input Kode of 4 Digit Only"
Else
ComboNip.SetFocus
End If
End If
End Sub
