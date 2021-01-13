VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmDailyRpt 
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4095
   LinkTopic       =   "Form2"
   ScaleHeight     =   1575
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   1245
      Width           =   4080
      _ExtentX        =   7197
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Proses Daily Report"
      Height          =   420
      Left            =   1185
      TabIndex        =   0
      Top             =   600
      Width           =   1650
   End
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   315
      Left            =   1305
      TabIndex        =   2
      Top             =   120
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Calendar        =   "FrmDailyRpt.frx":0000
      Caption         =   "FrmDailyRpt.frx":0118
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmDailyRpt.frx":0184
      Keys            =   "FrmDailyRpt.frx":01A2
      Spin            =   "FrmDailyRpt.frx":0200
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
End
Attribute VB_Name = "FrmDailyRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call HBatchMgmRptTracking
    Call HBatchLeadsRptTracking
    MsgBox "Proses Selesai", vbInformation + vbOKOnly, "Telegrandi"
End Sub

Private Sub HBatchMgmRptTracking()
Dim m_objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
Dim cmdsql As String

On Error GoTo HBatchMgmRptTrackingErr
ProgressBar1.Visible = True
m_objrs.CursorLocation = adUseClient
cmdsql = "select agent,recsource,kethslkerja,count(custid) AS JUMLAH from mgm where tglstatus between '" + Format(TDBDate1.Value, "mm/dd/yyyy") & " 00:00'" + " And '" + Format(TDBDate1.Value, "mm/dd/yyyy") & " 00:00" + "'"
cmdsql = cmdsql + " Group By agent , RECSOURCE, KETHSLKERJA"
m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

M_OBJCONN.Execute "Delete From BatchMgmRptTracking where tglstatus = '" + Format(TDBDate1.Value, "mm/dd/yyyy") + "'"

m_objrs1.CursorLocation = adUseClient
m_objrs1.Open "select * from BatchMgmRptTracking where tglstatus between '" + Format(TDBDate1.Value, "mm/dd/yyyy") & " 00:00'" + " And '" + Format(TDBDate1.Value, "mm/dd/yyyy") & " 00:00" + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

ProgressBar1.Max = m_objrs.RecordCount + 2
While Not m_objrs.EOF
    m_objrs1.AddNew
    ProgressBar1.Value = m_objrs.Bookmark
    m_objrs1!TGLSTATUS = Format(TDBDate1.Value, "mm/dd/yyyy")
    m_objrs1!Aoc = m_objrs!agent
    m_objrs1!batch = m_objrs!RECSOURCE
    m_objrs1!KETHSLKERJA = m_objrs!KETHSLKERJA
    m_objrs1!JUMLAH = m_objrs!JUMLAH
    m_objrs1.UPDATE
    m_objrs.MoveNext
Wend
m_objrs1.Close
m_objrs.Close
ProgressBar1.Value = ProgressBar1.Max
ProgressBar1.Visible = False
Set m_objrs1 = Nothing
Set m_objrs = Nothing
Exit Sub
HBatchMgmRptTrackingErr:
    MsgBox Err.Description
    Set m_objrs1 = Nothing
    Set m_objrs = Nothing
Exit Sub
End Sub

Private Sub HBatchLeadsRptTracking()
Dim m_objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
Dim cmdsql As String

On Error GoTo HBatchLeadsRptTrackingErr
ProgressBar1.Visible = True
m_objrs.CursorLocation = adUseClient
cmdsql = "select agent,recsource,kethslkerja,count(custid) AS JUMLAH from cc_Custtbl where tglstatus between '" + Format(TDBDate1.Value, "mm/dd/yyyy") & " 00:00'" + " And '" + Format(TDBDate1.Value, "mm/dd/yyyy") & " 00:00" + "'"
cmdsql = cmdsql + " Group By agent , RECSOURCE, KETHSLKERJA"
m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

M_OBJCONN.Execute "Delete From BatchLeadsRptTracking where tglstatus = '" + Format(TDBDate1.Value, "mm/dd/yyyy") + "'"

m_objrs1.CursorLocation = adUseClient
m_objrs1.Open "select * from BatchLeadsRptTracking where tglstatus between '" + Format(TDBDate1.Value, "mm/dd/yyyy") & " 00:00'" + " And '" + Format(TDBDate1.Value, "mm/dd/yyyy") & " 00:00" + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

ProgressBar1.Max = m_objrs.RecordCount + 2
While Not m_objrs.EOF
    m_objrs1.AddNew
    ProgressBar1.Value = m_objrs.Bookmark
    m_objrs1!TGLSTATUS = Format(Date, "mm/dd/yyyy")
    m_objrs1!Aoc = m_objrs!agent
    m_objrs1!batch = m_objrs!RECSOURCE
    m_objrs1!KETHSLKERJA = m_objrs!KETHSLKERJA
    m_objrs1!JUMLAH = m_objrs!JUMLAH
    m_objrs1.UPDATE
    m_objrs.MoveNext
Wend
m_objrs1.Close
m_objrs.Close
ProgressBar1.Value = ProgressBar1.Max
ProgressBar1.Visible = False
Set m_objrs1 = Nothing
Set m_objrs = Nothing
Exit Sub
HBatchLeadsRptTrackingErr:
    MsgBox Err.Description
    Set m_objrs1 = Nothing
    Set m_objrs = Nothing
Exit Sub
End Sub

Private Sub Form_Load()
    TDBDate1.Value = Date
End Sub
