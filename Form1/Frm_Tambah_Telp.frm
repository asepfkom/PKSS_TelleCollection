VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form Frm_Tambah_Telp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tambah No.Telepon"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4530
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdTambah 
      Caption         =   "&Tambah"
      Height          =   405
      Left            =   2190
      TabIndex        =   5
      Top             =   1650
      Width           =   1005
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   405
      Left            =   3240
      TabIndex        =   4
      Top             =   1650
      Width           =   1095
   End
   Begin VB.TextBox TxtDeskripsi 
      Height          =   735
      Left            =   1740
      MaxLength       =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   630
      Width           =   2655
   End
   Begin TDBNumber6Ctl.TDBNumber TxtNo 
      Height          =   255
      Left            =   2010
      TabIndex        =   1
      Top             =   300
      Width           =   525
      _Version        =   65536
      _ExtentX        =   926
      _ExtentY        =   450
      Calculator      =   "Frm_Tambah_Telp.frx":0000
      Caption         =   "Frm_Tambah_Telp.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "Frm_Tambah_Telp.frx":008C
      Keys            =   "Frm_Tambah_Telp.frx":00AA
      Spin            =   "Frm_Tambah_Telp.frx":00F4
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "####0;;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "####0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999
      MinValue        =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   -1
      ValueVT         =   2088828933
      Value           =   1
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Label Label4 
      Caption         =   "108"
      Height          =   195
      Left            =   2610
      TabIndex        =   7
      Top             =   330
      Width           =   465
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   195
      Left            =   1740
      TabIndex        =   6
      Top             =   330
      Width           =   225
   End
   Begin VB.Label Label2 
      Caption         =   "Deskripsi layanan:"
      Height          =   255
      Left            =   150
      TabIndex        =   2
      Top             =   630
      Width           =   1425
   End
   Begin VB.Label Label1 
      Caption         =   "No.Layanan"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   270
      Width           =   1245
   End
End
Attribute VB_Name = "Frm_Tambah_Telp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBatal_Click()
    TxtNo.Value = 1
    TxtDeskripsi.Text = ""
    Me.Hide
End Sub

Private Sub CmdTambah_Click()
    Dim cmdsql As String
    Dim M_OBJRS As ADODB.Recordset
    
    On Error GoTo salah
    If TxtNo.Value = 1 Or TxtNo.Value = 0 Then
        MsgBox "Inputkan no.telepon!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    cmdsql = "insert into tbllayanantelkom (nolayanan,desclayanan) values ('"
    cmdsql = cmdsql + "0" + Trim(CStr(TxtNo.Value)) + "108','"
    cmdsql = cmdsql + Trim(IIf(IsNull(TxtDeskripsi.Text), "", TxtDeskripsi.Text)) + "')"
    
    Set M_OBJRS = New ADODB.Recordset
    M_OBJRS.CursorLocation = adUseClient
    M_OBJRS.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    Set M_OBJRS = Nothing
    
    MsgBox "Data berhasil disimpan!", vbOKOnly + vbInformation, "Informasi"
    FrmCC_Colection.CmbPhone.AddItem "0" + Trim(TxtNo.Value) + "108"
    
    TxtNo.Value = 1
    TxtDeskripsi.Text = ""
    
    Me.Hide
    Exit Sub
salah:
    MsgBox "No layanan sudah ada!", vbOKOnly + vbCritical, "Error"
    Exit Sub
End Sub
