VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FrmDailyReportSubmission 
   Caption         =   "Daily Report Submission"
   ClientHeight    =   1830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8640
   Icon            =   "DailyReportSubmission.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   1830
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Set Tanggal"
      Height          =   480
      Index           =   3
      Left            =   75
      TabIndex        =   24
      Top             =   75
      Width           =   900
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Index           =   0
      Left            =   3510
      TabIndex        =   12
      Top             =   15
      Width           =   2085
   End
   Begin VB.ComboBox CmbmingguKe 
      Height          =   315
      Left            =   2415
      TabIndex        =   9
      Top             =   420
      Width           =   1005
   End
   Begin VB.ComboBox CmbTahun 
      Height          =   315
      Left            =   5610
      TabIndex        =   7
      Top             =   450
      Width           =   945
   End
   Begin VB.ComboBox CmbBulan 
      Height          =   315
      Left            =   3990
      TabIndex        =   5
      Top             =   420
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Proses"
      Height          =   315
      Index           =   2
      Left            =   5610
      TabIndex        =   4
      Top             =   1440
      Width           =   900
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Index           =   1
      Left            =   7515
      TabIndex        =   3
      Top             =   1440
      Width           =   900
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Preview"
      Height          =   315
      Index           =   0
      Left            =   6555
      TabIndex        =   2
      Top             =   1440
      Width           =   900
   End
   Begin TDBDate6Ctl.TDBDate TglPertama 
      Height          =   315
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   1065
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Calendar        =   "DailyReportSubmission.frx":000C
      Caption         =   "DailyReportSubmission.frx":0124
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "DailyReportSubmission.frx":0190
      Keys            =   "DailyReportSubmission.frx":01AE
      Spin            =   "DailyReportSubmission.frx":020C
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
   Begin Crystal.CrystalReport RPT 
      Left            =   255
      Top             =   2265
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   15
      TabIndex        =   13
      Top             =   1485
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
      TabIndex        =   14
      Top             =   1065
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Calendar        =   "DailyReportSubmission.frx":0234
      Caption         =   "DailyReportSubmission.frx":034C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "DailyReportSubmission.frx":03B8
      Keys            =   "DailyReportSubmission.frx":03D6
      Spin            =   "DailyReportSubmission.frx":0434
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
      TabIndex        =   16
      Top             =   1080
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Calendar        =   "DailyReportSubmission.frx":045C
      Caption         =   "DailyReportSubmission.frx":0574
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "DailyReportSubmission.frx":05E0
      Keys            =   "DailyReportSubmission.frx":05FE
      Spin            =   "DailyReportSubmission.frx":065C
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
      TabIndex        =   18
      Top             =   1080
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Calendar        =   "DailyReportSubmission.frx":0684
      Caption         =   "DailyReportSubmission.frx":079C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "DailyReportSubmission.frx":0808
      Keys            =   "DailyReportSubmission.frx":0826
      Spin            =   "DailyReportSubmission.frx":0884
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
      TabIndex        =   20
      Top             =   1065
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Calendar        =   "DailyReportSubmission.frx":08AC
      Caption         =   "DailyReportSubmission.frx":09C4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "DailyReportSubmission.frx":0A30
      Keys            =   "DailyReportSubmission.frx":0A4E
      Spin            =   "DailyReportSubmission.frx":0AAC
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
      TabIndex        =   22
      Top             =   1065
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Calendar        =   "DailyReportSubmission.frx":0AD4
      Caption         =   "DailyReportSubmission.frx":0BEC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "DailyReportSubmission.frx":0C58
      Keys            =   "DailyReportSubmission.frx":0C76
      Spin            =   "DailyReportSubmission.frx":0CD4
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
      TabIndex        =   23
      Top             =   825
      Width           =   1440
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Hari Kelima:"
      Height          =   300
      Index           =   3
      Left            =   5715
      TabIndex        =   21
      Top             =   825
      Width           =   1380
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Hari Keempat:"
      Height          =   300
      Index           =   2
      Left            =   4305
      TabIndex        =   19
      Top             =   825
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Hari Ketiga:"
      Height          =   300
      Index           =   1
      Left            =   2970
      TabIndex        =   17
      Top             =   825
      Width           =   1260
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Hari Kedua:"
      Height          =   300
      Index           =   0
      Left            =   1530
      TabIndex        =   15
      Top             =   810
      Width           =   1365
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Supervisor :"
      Height          =   270
      Left            =   2580
      TabIndex        =   11
      Top             =   45
      Width           =   825
   End
   Begin VB.Label Label4 
      Caption         =   "Minggu Ke :"
      Height          =   270
      Left            =   1530
      TabIndex        =   10
      Top             =   450
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Tahun :"
      Height          =   270
      Left            =   5010
      TabIndex        =   8
      Top             =   480
      Width           =   600
   End
   Begin VB.Label Label2 
      Caption         =   "Bulan :"
      Height          =   270
      Left            =   3465
      TabIndex        =   6
      Top             =   465
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Hari Pertama :"
      Height          =   300
      Index           =   5
      Left            =   120
      TabIndex        =   1
      Top             =   795
      Width           =   1365
   End
End
Attribute VB_Name = "FrmDailyReportSubmission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SHOW_PRN()
    RPT.RetrieveDataFiles
    RPT.WindowLeft = 0
    RPT.WindowTop = 0
    RPT.WindowState = crptMaximized
    RPT.WindowShowPrintBtn = True
    RPT.WindowShowRefreshBtn = True
    RPT.WindowShowSearchBtn = True
    RPT.WindowShowPrintSetupBtn = True
    RPT.WindowControls = True
    RPT.PrintReport
    RPT.Reset
End Sub

Private Sub Isi_AgentDailyReport()
Dim m_objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
Dim cmdsql As String
On Error GoTo Isi_AgentErr
    m_objrs.CursorLocation = adUseClient
    m_objrs1.CursorLocation = adUseClient
        cmdsql = "Select USERTBL.*, UserTblTarget.Absent1,UserTblTarget.Absent2,UserTblTarget.Absent3,UserTblTarget.Absent4, UserTblTarget.target1, UserTblTarget.target2, UserTblTarget.target3, UserTblTarget.target4 from "
        cmdsql = cmdsql + "  USERTBL INNER JOIN"
        cmdsql = cmdsql + "  UserTblTarget ON USERTBL.USERID = UserTblTarget.UserId"
        cmdsql = cmdsql + "  where USERTBL.SPVCODE ='" + Combo3(0).Text + "' AND USERTBL.USERTYPE =1 and UserTblTarget.bulan =" + CmbBulan.Text + " and UserTblTarget.tahun =" + CmbTahun.Text + " "
        m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    m_objrs1.Open "Select * from DailySubmissionRpt where bulan =" + CmbBulan.Text + " and tahun =" + CmbTahun.Text + " and SPV ='" + Combo3(0).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If m_objrs1.RecordCount <> 0 Then
        M_OBJCONN.Execute "Update DailySubmissionRpt set [" + CmbmingguKe.Text + "]=0 where bulan =" + CmbBulan.Text + " and tahun =" + CmbTahun.Text + " and SPV ='" + Combo3(0).Text + "'"
    Else
        ProgressBar1.Max = m_objrs.RecordCount + 2
        While Not m_objrs.EOF
            m_objrs1.AddNew
            m_objrs1!TSRNAME = CStr(IIf(IsNull(m_objrs!agent), "", m_objrs!agent))
            m_objrs1!SPV = CStr(IIf(IsNull(m_objrs!SPVCODE), "", m_objrs!SPVCODE))
            m_objrs1!AOC = CStr(IIf(IsNull(m_objrs!USERID), "", m_objrs!USERID))
            m_objrs1!Absent1 = IIf(IsNull(m_objrs!Absent1), 0, m_objrs!Absent1)
            m_objrs1!Absent2 = IIf(IsNull(m_objrs!Absent2), 0, m_objrs!Absent2)
            m_objrs1!Absent3 = IIf(IsNull(m_objrs!Absent3), 0, m_objrs!Absent3)
            m_objrs1!Absent4 = IIf(IsNull(m_objrs!Absent4), 0, m_objrs!Absent4)
            m_objrs1!target1 = IIf(IsNull(m_objrs!target1), 0, m_objrs!target1)
            m_objrs1!target2 = IIf(IsNull(m_objrs!target2), 0, m_objrs!target2)
            m_objrs1!target3 = IIf(IsNull(m_objrs!target3), 0, m_objrs!target3)
            m_objrs1!target4 = IIf(IsNull(m_objrs!target4), 0, m_objrs!target4)
            m_objrs1.UPDATE
            m_objrs.MoveNext
        Wend
    End If
    Set m_objrs = Nothing
    Set m_objrs1 = Nothing
Exit Sub
Isi_AgentErr:
    MsgBox Err.Description
End Sub

Private Sub Command1_Click(Index As Integer)
Dim m_objtgl As ADODB.Recordset
Select Case Index
    Case 1
        Unload Me
    Case 2
        If Len(CmbBulan.Text) = 0 Or Len(CmbTahun.Text) = 0 Or Len(CmbmingguKe.Text) = 0 Or Len(Combo3(0).Text) = 0 Then
            MsgBox "Data Tidak Lengkap", vbInformation + vbOKOnly, "Informasi"
            Exit Sub
        Else
            ProgressBar1.Visible = True
            Call isi_Agent
            'Call Isi_AgentDailyReport
            'Call hitungSubmission
            Call hitungSubmission
            MsgBox "done"
        End If
    Case 0
        If CmbmingguKe.Text = "" Or CmbBulan.Text = "" Or CmbTahun.Text = "" Then
            MsgBox "Minggu, Bulan dan tahun harus di isi", vbInformation + vbOKOnly, "Informasi"
            Exit Sub
        End If
        Call Isi_TempData
        Call isi_HitunganSubmission
        RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
        RPT.Formulas(2) = "@Bulan = totext('" + CmbBulan.Text + "')"
        RPT.Formulas(3) = "@Tahun = totext('" + CmbTahun.Text + "')"
        RPT.Formulas(4) = "@Spv = totext('" + CStr(Combo3(0).Text) + "')"
        RPT.ReportFileName = App.Path + "\Report\RptDailyTrackingRpt.rpt"
        Call SHOW_PRN
    Case 3
        If CmbmingguKe.Text = "" Or CmbBulan.Text = "" Or CmbTahun.Text = "" Then
            MsgBox "Minggu, Bulan dan tahun harus di isi", vbInformation + vbOKOnly, "Informasi"
            Exit Sub
        End If
        Set m_objtgl = New ADODB.Recordset
        m_objtgl.CursorLocation = adUseClient
        m_objtgl.Open "Select * from TblTanggal where minggu = " + CmbmingguKe + " and bulan = " + CmbBulan + " and tahun =" + CmbTahun + " order by tgl", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        While Not m_objtgl.EOF
            Select Case m_objtgl.Bookmark
                Case 1
                    TglPertama(0).Value = IIf(IsNull(m_objtgl!TGL), "", m_objtgl!TGL)
                Case 2
                    TglPertama(1).Value = IIf(IsNull(m_objtgl!TGL), "", m_objtgl!TGL)
                Case 3
                    TglPertama(2).Value = IIf(IsNull(m_objtgl!TGL), "", m_objtgl!TGL)
                Case 4
                    TglPertama(3).Value = IIf(IsNull(m_objtgl!TGL), "", m_objtgl!TGL)
                Case 5
                    TglPertama(4).Value = IIf(IsNull(m_objtgl!TGL), "", m_objtgl!TGL)
                Case 6
                    TglPertama(5).Value = IIf(IsNull(m_objtgl!TGL), "", m_objtgl!TGL)
            End Select
            m_objtgl.MoveNext
        Wend
        If m_objtgl.RecordCount = 0 Then
            MsgBox "Master Tanggal Belum di set....", vbInformation + vbOKOnly, "Telegrandi"
        End If
        Set m_objtgl = Nothing
End Select
End Sub

Private Sub Isi_TempData()
Dim m_objrs As ADODB.Recordset
Dim m_objacc As ADODB.Recordset

M_RPTCONN.Execute "Delete * from RptDailyReport"

Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient

Set m_objacc = New ADODB.Recordset
m_objacc.CursorLocation = adUseClient
m_objacc.Open "Select * from RptDailyReport", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
m_objrs.Open "Select * from UserTblTarget where SpvCode ='" + Combo3(0).Text + "' and Bulan = " + CmbBulan + " and tahun =" + CmbTahun + "", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not m_objrs.EOF
m_objacc.AddNew
    m_objacc!AOC = m_objrs!USERID
    m_objacc!TSRNAME = m_objrs!NAMAAGENT
    m_objacc!SPV = m_objrs!SPVCODE
    m_objacc!ABSENT = m_objrs!Absent1
    m_objacc!TARGET = m_objrs!target1
    m_objacc!Absent2 = m_objrs!Absent2
    m_objacc!target2 = m_objrs!target2
    m_objacc!Absent3 = m_objrs!Absent3
    m_objacc!target3 = m_objrs!target3
    m_objacc!Absent4 = m_objrs!Absent4
    m_objacc!target4 = m_objrs!target4
    m_objacc!Absent5 = m_objrs!Absent5
    m_objacc!target5 = m_objrs!target5
    m_objacc!Bulan = m_objrs!Bulan
    m_objacc!tahun = m_objrs!tahun
m_objacc.UPDATE
    m_objrs.MoveNext
Wend
Set m_objacc = Nothing
Set m_objrs = Nothing
End Sub

Private Sub isi_HitunganSubmission()
Dim m_objrs As ADODB.Recordset
Dim cmdsql As String
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open "Select * from RekapSubmission where SpvCode ='" + Combo3(0).Text + "' and Bulan =" + CmbBulan.Text + " and tahun =" + CmbTahun.Text + "", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not m_objrs.EOF
    cmdsql = "Update RptDailyReport set "
    Select Case m_objrs!Minggu
    Case 1
        Select Case m_objrs!hari
        Case 1
            cmdsql = cmdsql + " [1] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        Case 2
            cmdsql = cmdsql + " [2] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        Case 3
            cmdsql = cmdsql + " [3] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        Case 4
            cmdsql = cmdsql + " [4] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        Case 5
            cmdsql = cmdsql + " [5] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        Case 6
            cmdsql = cmdsql + " [6] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        End Select
    Case 2
        Select Case m_objrs!hari
        Case 1
            cmdsql = cmdsql + " [7] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        Case 2
            cmdsql = cmdsql + " [8] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        Case 3
            cmdsql = cmdsql + " [9] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        Case 4
            cmdsql = cmdsql + " [10] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        Case 5
            cmdsql = cmdsql + " [11] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        Case 6
            cmdsql = cmdsql + " [12] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        End Select
    Case 3
        Select Case m_objrs!hari
        Case 1
            cmdsql = cmdsql + " [13] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        Case 2
            cmdsql = cmdsql + " [14] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        Case 3
            cmdsql = cmdsql + " [15] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        Case 4
            cmdsql = cmdsql + " [16] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        Case 5
            cmdsql = cmdsql + " [17] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        Case 6
            cmdsql = cmdsql + " [18] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        End Select
    Case 4
        Select Case m_objrs!hari
        Case 1
            cmdsql = cmdsql + " [19] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        Case 2
            cmdsql = cmdsql + " [20] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        Case 3
            cmdsql = cmdsql + " [21] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        Case 4
            cmdsql = cmdsql + " [22] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        Case 5
            cmdsql = cmdsql + " [23] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        Case 6
            cmdsql = cmdsql + " [24] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        End Select
    Case 5
        Select Case m_objrs!hari
        Case 1
            cmdsql = cmdsql + " [25] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        Case 2
            cmdsql = cmdsql + " [26] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        Case 3
            cmdsql = cmdsql + " [27] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        Case 4
            cmdsql = cmdsql + " [28] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        Case 5
            cmdsql = cmdsql + " [29] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        Case 6
            cmdsql = cmdsql + " [30] = " + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
        End Select
    
    End Select
        cmdsql = cmdsql + " where AOC = '" + m_objrs!USERID + "' "
        M_RPTCONN.Execute cmdsql
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
End Sub

Private Sub isi_Agent()
Dim m_objrs  As New ADODB.Recordset
Dim m_cek As New ADODB.Recordset
Dim cmdsql As String
Dim m_msgbox As Variant
Dim i As Integer
m_objrs.CursorLocation = adUseClient
m_objrs.Open "Select * from UserTblTarget where Bulan =" + CmbBulan.Text + " and tahun = " + CmbTahun.Text + " and spvcode ='" + Combo3(0).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

If m_objrs.RecordCount <> 0 Then
    
'    Exit Sub
End If
ProgressBar1.Max = m_objrs.RecordCount + 1
m_cek.CursorLocation = adUseClient
m_cek.Open "Select * from RekapSubmission  where minggu =" + CmbmingguKe + " and Bulan =" + CmbBulan.Text + " and tahun = " + CmbTahun.Text + " and spvcode ='" + Combo3(0).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If m_cek.RecordCount <> 0 Then
    m_msgbox = MsgBox("Data sudah ada.. Proses ulang??..", vbYesNo + vbCritical, "Informas")
    If m_msgbox = vbNo Then
        Exit Sub
    End If
End If
Set m_cek = Nothing
M_OBJCONN.Execute "delete from RekapSubmission  where minggu =" + CmbmingguKe + " and Bulan =" + CmbBulan.Text + " and tahun = " + CmbTahun.Text + " and spvcode ='" + Combo3(0).Text + "'"
While Not m_objrs.EOF
ProgressBar1.Value = m_objrs.Bookmark
    For i = 1 To 6
        cmdsql = " Insert into RekapSubmission (UserId, hari, Minggu, Bulan, tahun, Jumlah, Tgl, SpvCode)"
        cmdsql = cmdsql + " Values "
        cmdsql = cmdsql + " ('" + m_objrs!USERID + "',"
        cmdsql = cmdsql + " " + CStr(i) + ","
        cmdsql = cmdsql + " " + CmbmingguKe.Text + ","
        cmdsql = cmdsql + " " + CmbBulan.Text + ","
        cmdsql = cmdsql + " " + CmbTahun.Text + ","
        cmdsql = cmdsql + " 0,"
        Select Case i
            Case 1
                If TglPertama(0).ValueIsNull Then
                    cmdsql = cmdsql + " null ,"
                Else
                    cmdsql = cmdsql + " '" + Format(TglPertama(0).Value, "yyyy/mm/dd") + "',"
                End If
'                    cmdsql = cmdsql + " " + CStr(m_objrs!absent1) + ","
'                    cmdsql = cmdsql + " " + CStr(m_objrs!target1) + ","
            Case 2
                If TglPertama(1).ValueIsNull Then
                    cmdsql = cmdsql + " null ,"
                Else
                    cmdsql = cmdsql + " '" + Format(TglPertama(1).Value, "yyyy/mm/dd") + "',"
                End If
 '                   cmdsql = cmdsql + " " + CStr(m_objrs!absent2) + ","
 '                   cmdsql = cmdsql + " " + CStr(m_objrs!target2) + ","
            Case 3
                If TglPertama(2).ValueIsNull Then
                    cmdsql = cmdsql + " null ,"
                Else
                    cmdsql = cmdsql + " '" + Format(TglPertama(2).Value, "yyyy/mm/dd") + "',"
                End If
 '                   cmdsql = cmdsql + " " + CStr(m_objrs!absent3) + ","
 '                   cmdsql = cmdsql + " " + CStr(m_objrs!target3) + ","
            Case 4
                If TglPertama(3).ValueIsNull Then
                    cmdsql = cmdsql + " null ,"
                Else
                    cmdsql = cmdsql + " '" + Format(TglPertama(3).Value, "yyyy/mm/dd") + "',"
                End If
 '                   cmdsql = cmdsql + " " + CStr(m_objrs!absent4) + ","
 '                   cmdsql = cmdsql + " " + CStr(m_objrs!target4) + ","
            Case 5
                If TglPertama(4).ValueIsNull Then
                    cmdsql = cmdsql + " null ,"
                Else
                    cmdsql = cmdsql + " '" + Format(TglPertama(4).Value, "yyyy/mm/dd") + "',"
                End If
            Case 6
                If TglPertama(5).ValueIsNull Then
                    cmdsql = cmdsql + " null ,"
                Else
                    cmdsql = cmdsql + " '" + Format(TglPertama(5).Value, "yyyy/mm/dd") + "',"
                End If
 '                   cmdsql = cmdsql + " " + CStr(m_objrs!absent5) + ","
 '                   cmdsql = cmdsql + " " + CStr(m_objrs!target5) + ","
        End Select
        cmdsql = cmdsql + "'" + m_objrs!SPVCODE + "')"
        M_OBJCONN.Execute cmdsql
    Next i
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
End Sub

Private Sub hitungSubmission()
Dim m_objrs As New ADODB.Recordset
Dim cmdsql As String
Dim i As Integer
m_objrs.CursorLocation = adUseClient
For i = 1 To 6
    Select Case i
        Case 1
            If TglPertama(0).ValueIsNull Then
            Else
                cmdsql = "Select agent,count(agent) as jumlah from cc_custtbl where kethslkerja ='I'"
                cmdsql = cmdsql + " and tglstatus between '" + Format(TglPertama(0).Value, "yyyy/mm/dd") & " 00:00" + "' and '" + Format(TglPertama(0).Value, "yyyy/mm/dd") & " 23:59" + "' group by agent"
            End If
        Case 2
            If TglPertama(1).ValueIsNull Then
            Else
                cmdsql = "Select agent,count(agent) as jumlah from cc_custtbl where kethslkerja ='I'"
                cmdsql = cmdsql + " and tglstatus between '" + Format(TglPertama(1).Value, "yyyy/mm/dd") & " 00:00" + "' and '" + Format(TglPertama(1).Value, "yyyy/mm/dd") & " 23:59" + "' group by agent"
            End If
        Case 3
            If TglPertama(2).ValueIsNull Then
            Else
                cmdsql = "Select agent,count(agent) as jumlah from cc_custtbl where kethslkerja ='I'"
                cmdsql = cmdsql + " and tglstatus between '" + Format(TglPertama(2).Value, "yyyy/mm/dd") & " 00:00" + "' and '" + Format(TglPertama(2).Value, "yyyy/mm/dd") & " 23:59" + "' group by agent"
            End If
        Case 4
            If TglPertama(3).ValueIsNull Then
            Else
                cmdsql = "Select agent,count(agent) as jumlah from cc_custtbl where kethslkerja ='I'"
                cmdsql = cmdsql + " and tglstatus between '" + Format(TglPertama(3).Value, "yyyy/mm/dd") & " 00:00" + "' and '" + Format(TglPertama(3).Value, "yyyy/mm/dd") & " 23:59" + "' group by agent"
            End If
        Case 5
            If TglPertama(4).ValueIsNull Then
            Else
                cmdsql = "Select agent,count(agent) as jumlah from cc_custtbl where kethslkerja ='I'"
                cmdsql = cmdsql + " and tglstatus between '" + Format(TglPertama(4).Value, "yyyy/mm/dd") & " 00:00" + "' and '" + Format(TglPertama(4).Value, "yyyy/mm/dd") & " 23:59" + "' group by agent"
            End If
        Case 6
            If TglPertama(5).ValueIsNull Then
            Else
                cmdsql = "Select agent,count(agent) as jumlah from cc_custtbl where kethslkerja ='I'"
                cmdsql = cmdsql + " and tglstatus between '" + Format(TglPertama(5).Value, "yyyy/mm/dd") & " 00:00" + "' and '" + Format(TglPertama(5).Value, "yyyy/mm/dd") & " 23:59" + "' group by agent"
            End If
    End Select
    If cmdsql = "" Then
        GoTo Nexti:
    Else
        m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    End If
While Not m_objrs.EOF
    cmdsql = "Update RekapSubmission "
    cmdsql = cmdsql + " set Jumlah = "
    cmdsql = cmdsql + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH))
    cmdsql = cmdsql + " where minggu = " + CmbmingguKe.Text + " and hari =" + CStr(i) + " and bulan =" + CmbBulan.Text + " and tahun =" + CmbTahun.Text + " and SPVcode ='" + Combo3(0).Text + "' and UserId ='" + CStr(IIf(IsNull(m_objrs!agent), "", m_objrs!agent)) + "'"
    M_OBJCONN.Execute cmdsql
    m_objrs.MoveNext
Wend
cmdsql = ""
Set m_objrs = Nothing
Nexti:
Next i
End Sub


Private Sub Form_Load()
Dim m_spv As ADODB.Recordset
Dim i As Integer
Set m_spv = New ADODB.Recordset
m_spv.CursorLocation = adUseClient
m_spv.CursorLocation = adUseClient
'm_spv.Open "SELECT * FROM spvtbl ORDER BY SPVCODE", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
m_spv.Open "select distinct SPVTBL.SPVCODE from SPVTBL, USERTBL where SPVTBL.SPVCODE = USERTBL.SPVCODE AND USERTYPE = '6'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

While Not m_spv.EOF
    Combo3(0).AddItem m_spv!SPVCODE
    m_spv.MoveNext
Wend
Set m_spv = Nothing
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

