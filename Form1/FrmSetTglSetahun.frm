VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmSetTglSetahun 
   Caption         =   "Set Tanggal Setahun ..."
   ClientHeight    =   1545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8640
   Icon            =   "FrmSetTglSetahun.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   1545
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmbmingguKe 
      Height          =   315
      Left            =   2415
      TabIndex        =   8
      Top             =   150
      Width           =   1005
   End
   Begin VB.ComboBox CmbTahun 
      Height          =   315
      Left            =   5610
      TabIndex        =   6
      Top             =   180
      Width           =   945
   End
   Begin VB.ComboBox CmbBulan 
      Height          =   315
      Left            =   3990
      TabIndex        =   4
      Top             =   150
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Insert"
      Height          =   315
      Index           =   2
      Left            =   5610
      TabIndex        =   3
      Top             =   1170
      Width           =   900
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   315
      Index           =   1
      Left            =   7515
      TabIndex        =   2
      Top             =   1170
      Width           =   900
   End
   Begin TDBDate6Ctl.TDBDate TglPertama 
      Height          =   315
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   795
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Calendar        =   "FrmSetTglSetahun.frx":000C
      Caption         =   "FrmSetTglSetahun.frx":0124
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmSetTglSetahun.frx":0190
      Keys            =   "FrmSetTglSetahun.frx":01AE
      Spin            =   "FrmSetTglSetahun.frx":020C
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd-mmm-yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   0
      Format          =   "dd-mmm-yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__-___-____"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   37468
      CenturyMode     =   0
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   15
      TabIndex        =   10
      Top             =   1215
      Visible         =   0   'False
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
   End
   Begin TDBDate6Ctl.TDBDate TglPertama 
      Height          =   315
      Index           =   1
      Left            =   1530
      TabIndex        =   11
      Top             =   795
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Calendar        =   "FrmSetTglSetahun.frx":0234
      Caption         =   "FrmSetTglSetahun.frx":034C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmSetTglSetahun.frx":03B8
      Keys            =   "FrmSetTglSetahun.frx":03D6
      Spin            =   "FrmSetTglSetahun.frx":0434
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd-mmm-yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   0
      Format          =   "dd-mmm-yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__-___-____"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   37468
      CenturyMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate TglPertama 
      Height          =   315
      Index           =   2
      Left            =   2940
      TabIndex        =   13
      Top             =   810
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Calendar        =   "FrmSetTglSetahun.frx":045C
      Caption         =   "FrmSetTglSetahun.frx":0574
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmSetTglSetahun.frx":05E0
      Keys            =   "FrmSetTglSetahun.frx":05FE
      Spin            =   "FrmSetTglSetahun.frx":065C
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd-mmm-yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   0
      Format          =   "dd-mmm-yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__-___-____"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   37468
      CenturyMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate TglPertama 
      Height          =   315
      Index           =   3
      Left            =   4320
      TabIndex        =   15
      Top             =   810
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Calendar        =   "FrmSetTglSetahun.frx":0684
      Caption         =   "FrmSetTglSetahun.frx":079C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmSetTglSetahun.frx":0808
      Keys            =   "FrmSetTglSetahun.frx":0826
      Spin            =   "FrmSetTglSetahun.frx":0884
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd-mmm-yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   0
      Format          =   "dd-mmm-yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__-___-____"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   37468
      CenturyMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate TglPertama 
      Height          =   315
      Index           =   4
      Left            =   5745
      TabIndex        =   17
      Top             =   795
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Calendar        =   "FrmSetTglSetahun.frx":08AC
      Caption         =   "FrmSetTglSetahun.frx":09C4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmSetTglSetahun.frx":0A30
      Keys            =   "FrmSetTglSetahun.frx":0A4E
      Spin            =   "FrmSetTglSetahun.frx":0AAC
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd-mmm-yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   0
      Format          =   "dd-mmm-yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__-___-____"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   37468
      CenturyMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate TglPertama 
      Height          =   315
      Index           =   5
      Left            =   7170
      TabIndex        =   19
      Top             =   795
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Calendar        =   "FrmSetTglSetahun.frx":0AD4
      Caption         =   "FrmSetTglSetahun.frx":0BEC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmSetTglSetahun.frx":0C58
      Keys            =   "FrmSetTglSetahun.frx":0C76
      Spin            =   "FrmSetTglSetahun.frx":0CD4
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd-mmm-yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   0
      Format          =   "dd-mmm-yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__-___-____"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   37468
      CenturyMode     =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Hari Keenam:"
      Height          =   300
      Index           =   4
      Left            =   7140
      TabIndex        =   20
      Top             =   555
      Width           =   1440
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Hari Kelima:"
      Height          =   300
      Index           =   3
      Left            =   5715
      TabIndex        =   18
      Top             =   555
      Width           =   1380
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Hari Keempat:"
      Height          =   300
      Index           =   2
      Left            =   4305
      TabIndex        =   16
      Top             =   555
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Hari Ketiga:"
      Height          =   300
      Index           =   1
      Left            =   2970
      TabIndex        =   14
      Top             =   555
      Width           =   1260
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Hari Kedua:"
      Height          =   300
      Index           =   0
      Left            =   1530
      TabIndex        =   12
      Top             =   540
      Width           =   1365
   End
   Begin VB.Label Label4 
      Caption         =   "Minggu Ke :"
      Height          =   270
      Left            =   1530
      TabIndex        =   9
      Top             =   180
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Tahun :"
      Height          =   270
      Left            =   5010
      TabIndex        =   7
      Top             =   210
      Width           =   600
   End
   Begin VB.Label Label2 
      Caption         =   "Bulan :"
      Height          =   270
      Left            =   3465
      TabIndex        =   5
      Top             =   195
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Hari Pertama :"
      Height          =   300
      Index           =   5
      Left            =   120
      TabIndex        =   1
      Top             =   525
      Width           =   1365
   End
End
Attribute VB_Name = "FrmSetTglSetahun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
Select Case Index
    Case 1
        Unload Me
    Case 2
        If Len(CmbBulan.Text) = 0 Or Len(CmbTahun.Text) = 0 Or Len(CmbmingguKe.Text) = 0 Then
            MsgBox "Data Tidak Lengkap", vbInformation + vbOKOnly, "Informasi"
            Exit Sub
        Else
            ProgressBar1.Visible = True
            If CmbmingguKe.Text = "" Or CmbBulan.Text = "" Or CmbTahun.Text = "" Then
                MsgBox "Minimal Minggu, Bulan dan Tahun harus di isi", vbInformation + vbOKOnly, "Telegrandi"
                Exit Sub
            End If
            Call IsiTanggal
            
            MsgBox "done"
        End If
    Case 0
End Select
End Sub

Private Sub IsiTanggal()
Dim cmdsql As String
Dim m_msgbox As Variant
On Error GoTo addErr
Dim m_objtanggal As ADODB.Recordset
Set m_objtanggal = New ADODB.Recordset
m_objtanggal.CursorLocation = adUseClient
m_objtanggal.Open "Select * from TblTanggal where Minggu =" + CmbmingguKe.Text + " and Bulan = " + CmbBulan.Text + "  and Tahun =" + CmbTahun.Text + " ", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If m_objtanggal.RecordCount <> 0 Then
    m_msgbox = MsgBox("Data sudah pernah ada Update dengan yang baru", vbYesNo + vbQuestion, "Telegrandi")
    If m_msgbox = vbNo Then
        Exit Sub
    Else
        M_OBJCONN.Execute "Delete from tbltanggal where Minggu =" + CmbmingguKe.Text + " and Bulan = " + CmbBulan.Text + "  and Tahun =" + CmbTahun.Text + " "
    End If
End If
m_objtanggal.Requery
    If TglPertama(0).ValueIsNull = False Then
        m_objtanggal.AddNew
        m_objtanggal!TGL = Format(TglPertama(0).Value, "yyyy/mm/dd")
        m_objtanggal!Minggu = CmbmingguKe.Text
        m_objtanggal!Bulan = CmbBulan.Text
        m_objtanggal!tahun = CmbTahun.Text
        m_objtanggal.UPDATE
    End If
    If TglPertama(1).ValueIsNull = False Then
        m_objtanggal.AddNew
        m_objtanggal!TGL = Format(TglPertama(1).Value, "yyyy/mm/dd")
        m_objtanggal!Minggu = CmbmingguKe.Text
        m_objtanggal!Bulan = CmbBulan.Text
        m_objtanggal!tahun = CmbTahun.Text
        m_objtanggal.UPDATE
    End If
    If TglPertama(2).ValueIsNull = False Then
        m_objtanggal.AddNew
        m_objtanggal!TGL = Format(TglPertama(2).Value, "yyyy/mm/dd")
        m_objtanggal!Minggu = CmbmingguKe.Text
        m_objtanggal!Bulan = CmbBulan.Text
        m_objtanggal!tahun = CmbTahun.Text
        m_objtanggal.UPDATE
    End If
    If TglPertama(3).ValueIsNull = False Then
        m_objtanggal.AddNew
        m_objtanggal!TGL = Format(TglPertama(3).Value, "yyyy/mm/dd")
        m_objtanggal!Minggu = CmbmingguKe.Text
        m_objtanggal!Bulan = CmbBulan.Text
        m_objtanggal!tahun = CmbTahun.Text
        m_objtanggal.UPDATE
    End If
    If TglPertama(4).ValueIsNull = False Then
        m_objtanggal.AddNew
        m_objtanggal!TGL = Format(TglPertama(4).Value, "yyyy/mm/dd")
        m_objtanggal!Minggu = CmbmingguKe.Text
        m_objtanggal!Bulan = CmbBulan.Text
        m_objtanggal!tahun = CmbTahun.Text
        m_objtanggal.UPDATE
    End If
    If TglPertama(5).ValueIsNull = False Then
        m_objtanggal.AddNew
        m_objtanggal!TGL = Format(TglPertama(5).Value, "yyyy/mm/dd")
        m_objtanggal!Minggu = CmbmingguKe.Text
        m_objtanggal!Bulan = CmbBulan.Text
        m_objtanggal!tahun = CmbTahun.Text
        m_objtanggal.UPDATE
    End If
Set m_objtanggal = Nothing
Exit Sub
addErr:
    MsgBox Err.Description
    Set m_objtanggal = Nothing
    Exit Sub
End Sub

Private Sub Form_Load()
Dim m_spv As ADODB.Recordset
Dim i As Integer
For i = 1 To 5
    CmbmingguKe.AddItem i
Next i
CmbTahun.AddItem 2005
CmbTahun.AddItem 2006
CmbTahun.AddItem 2007
CmbTahun.AddItem 2008
CmbTahun.AddItem 2009
CmbTahun.AddItem 2010
For i = 1 To 12
    CmbBulan.AddItem i
Next i
End Sub

