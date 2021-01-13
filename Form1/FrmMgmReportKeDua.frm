VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FrmMgmReportKeDua 
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11400
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   3405
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport RPT 
      Left            =   6420
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox CmbCek 
      Height          =   315
      Left            =   8925
      TabIndex        =   24
      Top             =   2370
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Preview"
      Height          =   360
      Index           =   0
      Left            =   8880
      TabIndex        =   12
      Top             =   2730
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   360
      Index           =   1
      Left            =   10230
      TabIndex        =   13
      Top             =   2730
      Width           =   1125
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   9060
      TabIndex        =   7
      Top             =   1695
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   6570
      TabIndex        =   6
      Top             =   1695
      Width           =   2265
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose One..."
      Height          =   1035
      Left            =   5550
      TabIndex        =   14
      Top             =   570
      Width           =   5805
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   1
         Left            =   3630
         TabIndex        =   2
         Top             =   195
         Width           =   2130
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   0
         Left            =   1215
         TabIndex        =   1
         Top             =   195
         Width           =   2085
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Index           =   1
         Left            =   3630
         TabIndex        =   5
         Top             =   540
         Width           =   2130
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Index           =   0
         Left            =   1230
         TabIndex        =   4
         Top             =   540
         Width           =   2085
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Agent        :"
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   210
         Width           =   1125
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Supervisor :"
         Height          =   270
         Index           =   1
         Left            =   90
         TabIndex        =   3
         Top             =   555
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "to"
         Height          =   300
         Index           =   2
         Left            =   3375
         TabIndex        =   18
         Top             =   225
         Width           =   270
      End
      Begin VB.Label Label1 
         Caption         =   "to"
         Height          =   300
         Index           =   6
         Left            =   3405
         TabIndex        =   17
         Top             =   570
         Width           =   270
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3405
      Left            =   -30
      TabIndex        =   16
      Top             =   -15
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   6006
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   5535
      TabIndex        =   15
      Top             =   3165
      Visible         =   0   'False
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
   End
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   315
      Index           =   1
      Left            =   9060
      TabIndex        =   10
      Top             =   2010
      Width           =   1440
      _Version        =   65536
      _ExtentX        =   2540
      _ExtentY        =   556
      Calendar        =   "FrmMgmReportKeDua.frx":0000
      Caption         =   "FrmMgmReportKeDua.frx":0118
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmMgmReportKeDua.frx":0184
      Keys            =   "FrmMgmReportKeDua.frx":01A2
      Spin            =   "FrmMgmReportKeDua.frx":0200
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
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   315
      Index           =   0
      Left            =   6570
      TabIndex        =   8
      Top             =   2040
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Calendar        =   "FrmMgmReportKeDua.frx":0228
      Caption         =   "FrmMgmReportKeDua.frx":0340
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmMgmReportKeDua.frx":03AC
      Keys            =   "FrmMgmReportKeDua.frx":03CA
      Spin            =   "FrmMgmReportKeDua.frx":0428
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
   Begin TDBTime6Ctl.TDBTime DTimeLastCall 
      Height          =   300
      Index           =   0
      Left            =   7950
      TabIndex        =   9
      Top             =   2025
      Width           =   885
      _Version        =   65536
      _ExtentX        =   1561
      _ExtentY        =   529
      Caption         =   "FrmMgmReportKeDua.frx":0450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmMgmReportKeDua.frx":04BC
      Spin            =   "FrmMgmReportKeDua.frx":050C
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "hh:nn"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "hh:nn"
      HighlightText   =   0
      Hour12Mode      =   1
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxTime         =   0.99999
      MidnightMode    =   0
      MinTime         =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__:__"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   0.870289351851852
   End
   Begin TDBTime6Ctl.TDBTime DTimeLastCall 
      Height          =   300
      Index           =   1
      Left            =   10485
      TabIndex        =   11
      Top             =   2010
      Width           =   885
      _Version        =   65536
      _ExtentX        =   1561
      _ExtentY        =   529
      Caption         =   "FrmMgmReportKeDua.frx":0534
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmMgmReportKeDua.frx":05A0
      Spin            =   "FrmMgmReportKeDua.frx":05F0
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "hh:nn"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "hh:nn"
      HighlightText   =   0
      Hour12Mode      =   1
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxTime         =   0.99999
      MidnightMode    =   0
      MinTime         =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__:__"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   0.870289351851852
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Status Cek :"
      Height          =   315
      Left            =   7635
      TabIndex        =   25
      Top             =   2385
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Label Label1 
      Caption         =   "to"
      Height          =   255
      Index           =   1
      Left            =   8850
      TabIndex        =   23
      Top             =   1710
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "From Batch :"
      Height          =   300
      Index           =   0
      Left            =   5595
      TabIndex        =   22
      Top             =   1725
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Date :"
      Height          =   300
      Index           =   5
      Left            =   5595
      TabIndex        =   21
      Top             =   2055
      Width           =   930
   End
   Begin VB.Label Label1 
      Caption         =   "to"
      Height          =   300
      Index           =   4
      Left            =   8850
      TabIndex        =   20
      Top             =   2040
      Width           =   270
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   5565
      TabIndex        =   19
      Top             =   75
      Width           =   5745
   End
End
Attribute VB_Name = "FrmMgmReportKeDua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
Private Sub ReportPTPNego()
Dim Rsptp As ADODB.Recordset
Dim m_msgbox As Variant
Dim CMDSQL As String
Dim LAgent As String
Dim jml As String
Dim Lf_cek As String
Dim Lvol As String


On Error GoTo eddder:
Set Rsptp = New ADODB.Recordset
Rsptp.CursorLocation = adUseClient
If Option1(0).Value Then
CMDSQL = "select agent,f_cek,count(agent) as JML,sum(promisepay) as VOL from reportPTP  where "
CMDSQL = CMDSQL + " agent in (select userid from usertbl where userid >='" + Combo2(0).Text + "' and userid<='" + Combo2(1).Text + "') and "
CMDSQL = CMDSQL + " RECSOURCE Between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and promisedate between  "
CMDSQL = CMDSQL + " '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "' "
CMDSQL = CMDSQL + "  group by agent, f_cek "

Else
If Option1(1).Value Then
CMDSQL = "select agent,f_cek, count(agent) as JML,sum(promisepay) as VOL from reportPTP  where "
CMDSQL = CMDSQL + " agent in (select userid from usertbl where spvcode >='" + Combo3(0).Text + "' and SPVCODE<='" + Combo3(1).Text + "') and "
CMDSQL = CMDSQL + " RECSOURCE Between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and promisedate between  "
CMDSQL = CMDSQL + " '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "' "
CMDSQL = CMDSQL + "  group by agent,f_cek "
End If
End If

Rsptp.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
While Not Rsptp.EOF
LAgent = Trim(IIf(IsNull(Rsptp!agent), "", Rsptp!agent))
jml = Trim(IIf(IsNull(Rsptp!jml), 0, Rsptp!jml))
Lf_cek = Trim(IIf(IsNull(Rsptp!F_CEK), "", Rsptp!F_CEK))
Lvol = Trim(IIf(IsNull(Rsptp!vol), 0, Rsptp!vol))
If Lf_cek = "PTP" Then
jml = jml
Else
jml = 0
End If
M_RPTCONN.Execute "UPDATE TrackingRptPerPrgBatch set PTP_BARU =" + jml + ",VolPTP_Baru=" + Lvol + "  where AOC = '" + LAgent + "'"
Rsptp.MoveNext
Wend
Set Rsptp = Nothing
CMDSQL = Empty
LAgent = Empty
jml = Empty
Lf_cek = Empty

Exit Sub
eddder:
    If err.number = -2147217871 Then
        m_msgbox = MsgBox(err.Description, vbRetryCancel, "Aplikasi")
        If m_msgbox = vbRetry Then
            WaitSecs (5)
            Resume
        End If
    Else
        MsgBox err.Description
    End If
    Resume Next

End Sub


Private Sub Isi_Report_PTP_Jatuh_Tempo()
Dim M_objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
Dim CMDSQL As String
On Error GoTo Isi_AgentErr

M_RPTCONN.Execute "Delete * from TrackingPTP"
Set M_objrs = New ADODB.Recordset
Set m_objrs1 = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
m_objrs1.CursorLocation = adUseClient
If Option1(1).Value = True Then
CMDSQL = "Select * from mgm where f_cek='PTP' and tdbdatePTP between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and recsource between  '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and agent in (select userid from usertbl where SPVCODE >='" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "') ORDER BY AGENT"
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    If Option1(0).Value = True Then
       CMDSQL = "Select * from mgm where f_cek='PTP' and tdbdatePTP between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and  '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' and agent between '" + Combo2(0).Text + "' and '" + Combo2(1).Text + "' ORDER BY AGENT"
        M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    End If
End If
m_objrs1.Open "Select * from TrackingPTP", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText

ProgressBar1.Max = M_objrs.RecordCount + 1
While Not M_objrs.EOF
    ProgressBar1.Value = M_objrs.Bookmark
    m_objrs1.AddNew
    m_objrs1!agent = CStr(IIf(IsNull(M_objrs!agent), "", M_objrs!agent))
    m_objrs1!CustId = CStr(IIf(IsNull(M_objrs!CustId), "", M_objrs!CustId))
    m_objrs1!Name = CStr(IIf(IsNull(M_objrs!Name), "", M_objrs!Name))
    m_objrs1!TglPtp = CStr(IIf(IsNull(M_objrs!TdbDatePTP), "2020-12-30", M_objrs!TdbDatePTP))
    m_objrs1!ttlptp = CStr(IIf(IsNull(M_objrs!ttlptp), "0", M_objrs!ttlptp))
    m_objrs1!BaseOn = CStr(IIf(IsNull(M_objrs!CmbBaseOn), "", M_objrs!CmbBaseOn))
    m_objrs1!Principle = CStr(IIf(IsNull(M_objrs!Principal), "0", M_objrs!Principal))
    m_objrs1!amountwo = CStr(IIf(IsNull(M_objrs!amountwo), "0", M_objrs!amountwo))
    m_objrs1.update
    M_objrs.MoveNext
Wend
Set M_objrs = Nothing
Set m_objrs1 = Nothing
Exit Sub

Isi_AgentErr:
    MsgBox err.Description
    'Resume

End Sub

Private Sub AmbilDtYgDiFU_PerAgent()
Dim m_hst As New ADODB.Recordset
Dim tglawal As String
Dim tglakhir As String
Dim LAgent As String
Dim CMDSQL As String
Dim m_msgbox As Variant
Dim STATUS As String
On Error GoTo eddder
tglawal = Format(TDBDate1(0).Value, "yyyymmdd") & Format(DTimeLastCall(0).Value, "hhnn")
tglakhir = Format(TDBDate1(1).Value, "yyyymmdd") & Format(DTimeLastCall(1).Value, "hhnn")
m_hst.CursorLocation = adUseClient
If Option1(0).Value Then
CMDSQL = "SELECT mgm_hst.agent, count(mgm_hst.f_cek)as jml, mgm_hst.f_cek,  sum(mgm.ttlptp)as jmlPTP from mgm_hst INNER JOIN "
CMDSQL = CMDSQL + "(SELECT custid, agent,max(tgl) as tglmax FROM mgm_hst where tgl BETWEEN "
CMDSQL = CMDSQL + " '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  and "
CMDSQL = CMDSQL + " custid in(select custid from mgm where agent between  '" + Combo2(0).Text + "' and '" + Combo2(1).Text + "' ) "
CMDSQL = CMDSQL + " group by custid,agent) as a  on mgm_hst.custid = a.custid and mgm_hst.tgl = a.tglmax  "
CMDSQL = CMDSQL + "inner join mgm on mgm.custid = a.custid and a.agent=mgm.agent where "
CMDSQL = CMDSQL + " recsource BETWEEN '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'"
CMDSQL = CMDSQL + " group by mgm_hst.agent, mgm_hst.f_cek order by mgm_hst.agent"
Else
If Option1(1).Value Then
CMDSQL = "SELECT mgm_hst.agent, count(mgm_hst.f_cek)as jml, mgm_hst.f_cek,  sum(mgm.ttlptp)as jmlPTP from mgm_hst INNER JOIN "
CMDSQL = CMDSQL + "(SELECT custid, agent,max(tgl) as tglmax FROM mgm_hst where tgl BETWEEN "
CMDSQL = CMDSQL + " '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  and "
CMDSQL = CMDSQL + " custid in(select custid from mgm where agent in (select userid from usertbl where "
CMDSQL = CMDSQL + " SPVCODE >= '" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "'))"
CMDSQL = CMDSQL + " group by custid,agent) as a  on mgm_hst.custid = a.custid and mgm_hst.tgl = a.tglmax  "
CMDSQL = CMDSQL + "inner join mgm on mgm.custid = a.custid and a.agent=mgm.agent where "
CMDSQL = CMDSQL + " recsource BETWEEN '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'"
CMDSQL = CMDSQL + " group by mgm_hst.agent, mgm_hst.f_cek order by mgm_hst.agent"
 End If
End If
m_hst.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_hst.RecordCount + 1
While Not m_hst.EOF
    ProgressBar1.Value = m_hst.Bookmark
        LAgent = CStr(IIf(IsNull(m_hst!agent), "", m_hst!agent))
        CMDSQL = "Update TrackingRptPerPrgBatch Set "
        Select Case Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 2)
            Case "NK"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 5)
            Case "MV"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 5)
            Case "WN"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 5)
            Case "BP"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
            Case "PT"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
            Case "RP"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 4)
            Case "NA"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 4)
            Case "SP"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
              Case "PO"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
             Case "OP"
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 3)
            Case Else
                STATUS = Left(IIf(IsNull(m_hst!F_CEK), "", m_hst!F_CEK), 4)
            
        End Select
        CMDSQL = CMDSQL + "[" + STATUS + "]"
        CMDSQL = CMDSQL + "=" + CStr(IIf(IsNull(m_hst!jml), 0, m_hst!jml)) + " "
        If STATUS = "PTP" Then
        CMDSQL = CMDSQL + ", VOLPTP =" + CStr(IIf(IsNull(m_hst!jmlPTP), 0, m_hst!jmlPTP)) + " "
        End If
        CMDSQL = CMDSQL + " Where AOC = '" + LAgent + "'"
        If IsNull(m_hst!F_CEK) Then
        Else
            If m_hst!F_CEK = "" Then
            Else
                If m_hst!jml = 0 Then
                Else
                   
                    M_RPTCONN.Execute CMDSQL
                End If
            End If
        End If
    m_hst.MoveNext
Wend
Set m_hst = Nothing
Exit Sub
eddder:
    If err.number = -2147217871 Then
        m_msgbox = MsgBox(err.Description, vbRetryCancel, "Aplikasi")
        If m_msgbox = vbRetry Then
            WaitSecs (5)
            Resume
        End If
    Else
  '      MsgBox Err.Description
    End If
    Resume Next
End Sub

Private Sub Combo1_Click(Index As Integer)
    Call Combo1_LostFocus(Index)
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim M_objrs As New ADODB.Recordset
On Error GoTo Combo1_LostFocusErr
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open "Select * from datasourcetbl where kodeds ='" + Combo1(Index).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If Not M_objrs.EOF Then
        Combo1(Index).Text = M_objrs!KODEDS
    Else
        Combo1(Index).Text = Empty
    End If
Exit Sub
Combo1_LostFocusErr:
    MsgBox err.Description
End Sub

Private Sub Combo2_Click(Index As Integer)
    Call Combo2_LostFocus(Index)
End Sub

Private Sub Combo2_LostFocus(Index As Integer)
Dim M_objrs As New ADODB.Recordset
On Error GoTo Combo2_LostFocusErr
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open "Select * from usertbl where AKTIF = 0 AND USERID ='" + Combo2(Index).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If Not M_objrs.EOF Then
        Combo2(Index).Text = M_objrs!USERID
    Else
        Combo2(Index).Text = Empty
    End If
Exit Sub
Combo2_LostFocusErr:
    MsgBox err.Description
End Sub

Private Sub Command1_Click(Index As Integer)
On Error GoTo Command1_ClickeR
If TDBDate1(0).ValueIsNull And TDBDate1(1).ValueIsNull Then
    TDBDate1(0).Value = "01/01/1990"
    TDBDate1(1).Value = "31/12/2020"
End If

If Combo1(0).Text = Empty And Combo1(1).Text = Empty Then
    Combo1(0).Text = "-----"
    Combo1(1).Text = "ZZZZZ"
End If
If Option1(0).Value = False And Option1(1).Value = False Then
If Combo2(0).Text = Empty And Combo2(1).Text = Empty Then
    Combo2(0).Text = "-----"
    Combo2(1).Text = "ZZZZZ"
End If
End If
ProgressBar1.Visible = True
Select Case Index
    Case 0
    Select Case ListView1.SelectedItem.Text
        Case 1 'mgm
           Call Isi_Agent_mgm
            Call hitung_JmlData_PerAgent_mgm
            Call AmbilDtYgDiFU_PerAgent
            Call ReportPTPNego
'           Call Isi_Settled_Payment
'            Call Isi_Progess_OF_PAyment
'            Call hitung_JmlData_PerAgent_PTP
          '' Call Hitung_JmlLeadsPerAgent
'          Call Hitung_Vol_PTP
           Call hitung_BatchCallInitilized_PerAgent_mgm
           Call Hitung_Number_of_Payment
            Call Hitung_Volume_of_Payment
        
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.TxtUsername.Text) + "')"
            RPT.Formulas(2) = "@TglShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
            RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
            
            'RPT.ReportFileName = App.Path + "\Report\Tracking ReportAgent.rpt"
            Call SHOW_PRN
        
        
        Case 2 'PTP DUE DATE
            Call Isi_Report_PTP_Jatuh_Tempo
            WaitSecs (2)
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.TxtUsername.Text) + "')"
            RPT.Formulas(2) = "@TglAwalShow = totext('" + CStr(TDBDate1(0).Text & " " & DTimeLastCall(0).Text) + "')"
            RPT.Formulas(3) = "@TglAkhirShow = totext('" + CStr(TDBDate1(1).Text & " " & DTimeLastCall(1).Text) + "')"
            RPT.Formulas(4) = "@TglAwal = totext('" + CStr(Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(0).Value, "hh:nn"))) + "')"
            RPT.Formulas(5) = "@TglAkhir = totext('" + CStr(Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & CStr(Format(DTimeLastCall(1).Value, "hh:nn"))) + "')"
            
            RPT.ReportFileName = "D:\COLLECTION\Report\RptPTPJatuhTempo.rpt"
            Call SHOW_PRN
            
'        Case 3
'            Call Isi_Agent_mgm
'            Call isi_PTP
            
           End Select
    ProgressBar1.Visible = False
    Case 1
        Unload Me
End Select
ProgressBar1.Visible = False
Exit Sub
Command1_ClickeR:
    MsgBox err.Description
End Sub

Private Sub isi_PTP()
Dim M_objrs As New ADODB.Recordset
Dim JUMLAH As String
Dim agent As String
Dim TGL As Integer
Dim CMDSQL As String

On Error GoTo hitung_JmlDataer
M_objrs.CursorLocation = adUseClient
CMDSQL = "SELECT mgm.AGENT,TBLNEGOPTP.PROMISEDATE,TBLNEGOPTP.Promisepay from TBLNEGOPTP,mgm "
CMDSQL = CMDSQL + "Where mgm.CustId = TBLNEGOPTP.CustId AND tblnegoptp.promisedate between between '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "'"

'M_OBJRS.Open "Select recsource, count(custid) as jml from mgm  where recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "' group by recsource", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_objrs.RecordCount + 2
While Not M_objrs.EOF
    ProgressBar1.Value = M_objrs.Bookmark
    JUMLAH = IIf(IsNull(M_objrs!PromisePay), "", DatePart("D", M_objrs!PromisePay))
    TGL = IIf(IsNull(M_objrs!PromiseDate), 0, DatePart("D", M_objrs!PromiseDate))
    agent = CStr(IIf(IsNull(M_objrs!agent), "", M_objrs!agent))
    CMDSQL = "Update TrackingRptPerPrgBatch set TGL '" + TGL + "' = " + JUMLAH + " where USERID ='" + agent + "'"
    M_RPTCONN.Execute CMDSQL
    M_objrs.MoveNext
Wend
Set M_objrs = Nothing
Exit Sub
hitung_JmlDataer:
MsgBox err.Description
End Sub

Private Sub Hitung_JmlLeadsPerAgent()
Dim M_objrs As ADODB.Recordset
Dim m_objrs1 As ADODB.Recordset
Dim CMDSQL As String
Dim LAgent As String

On Error GoTo Hitung_JmlLeadsPerAgent
    Me.MousePointer = vbHourglass
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    CMDSQL = "Select Agent, count(custid) as jumlah from CC_Custtbl where tglsource between '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "' and RecSourceRef >='" + Combo1(0).Text + "' and RecSourceRef <='" + Combo1(1).Text + "' group by Agent"
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = M_objrs.RecordCount + 2
    While Not M_objrs.EOF
        ProgressBar1.Value = M_objrs.Bookmark
        LAgent = CStr(IIf(IsNull(M_objrs!agent), "", M_objrs!agent))
        CMDSQL = "Update TrackingRptPerPrgBatch Set "
        CMDSQL = CMDSQL + " LEADS "
        CMDSQL = CMDSQL + "=" + CStr(IIf(IsNull(M_objrs!JUMLAH), 0, M_objrs!JUMLAH)) + ""
        CMDSQL = CMDSQL + " Where AOC = '" + LAgent + "'"
        M_RPTCONN.Execute CMDSQL
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
    Me.MousePointer = vbNormal
Exit Sub
Hitung_JmlLeadsPerAgent:
Me.MousePointer = vbNormal
MsgBox err.Description
End Sub

Private Sub Hitung_TrackingReport_mgm()
Dim M_objrs As ADODB.Recordset
Dim m_objrs1 As ADODB.Recordset
Dim CMDSQL As String
On Error GoTo Hitung_TrackingReportErr
    Me.MousePointer = vbHourglass
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open "Select recsource, kethslkerja, count(custid) as jumlah from mgm  where tglcall between '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "' and recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "'  group by recsource, kethslkerja", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = M_objrs.RecordCount + 2
    While Not M_objrs.EOF
        ProgressBar1.Value = M_objrs.Bookmark
        CMDSQL = "Update TrackingRptPerPrgBatch Set "
        CMDSQL = CMDSQL + Trim(CStr(M_objrs!KETHSLKERJA))
        CMDSQL = CMDSQL + "=" + CStr(IIf(IsNull(M_objrs!JUMLAH), 0, M_objrs!JUMLAH)) + ""
        CMDSQL = CMDSQL + " where BATCH = '" + IIf(IsNull(M_objrs!RECSOURCE), "", M_objrs!RECSOURCE) + "'"
        If IsNull(M_objrs!KETHSLKERJA) Then
        Else
            If M_objrs!KETHSLKERJA = Empty Then
            Else
            M_RPTCONN.Execute CMDSQL
            End If
        End If
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
    Me.MousePointer = vbNormal
Exit Sub
Hitung_TrackingReportErr:
Me.MousePointer = vbNormal
MsgBox err.Description
End Sub

Private Sub hitung_JmlData_mgm()
Dim M_objrs As New ADODB.Recordset
Dim JUMLAH As String
Dim batch As String
Dim CMDSQL As String

On Error GoTo hitung_JmlDataer
M_objrs.CursorLocation = adUseClient
M_objrs.Open "Select recsource, count(custid) as jml from mgm  where recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "' group by recsource", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_objrs.RecordCount + 2
While Not M_objrs.EOF
    ProgressBar1.Value = M_objrs.Bookmark
    JUMLAH = CStr(IIf(IsNull(M_objrs!jml), "", M_objrs!jml))
    batch = CStr(IIf(IsNull(M_objrs!RECSOURCE), "", M_objrs!RECSOURCE))
    CMDSQL = "Update TrackingRptPerPrgBatch set DATASIZE =" + JUMLAH + " where BATCH ='" + batch + "'"
    M_RPTCONN.Execute CMDSQL
    M_objrs.MoveNext
Wend
Set M_objrs = Nothing
Exit Sub
hitung_JmlDataer:
MsgBox err.Description
End Sub

Private Sub hitung_BatchCallInitilized_mgm()
Dim M_objrs As New ADODB.Recordset
Dim JUMLAH As String
Dim batch As String
Dim CMDSQL As String
Dim LAgent As String
On Error GoTo hitung_BatchCallInitilizeder
M_objrs.CursorLocation = adUseClient
CMDSQL = "Select recsource ,count(userid) as jml from TblPhoneMonitorHst where tgl between '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "' group by recsource "
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_objrs.RecordCount + 2
While Not M_objrs.EOF
    ProgressBar1.Value = M_objrs.Bookmark
    JUMLAH = CStr(IIf(IsNull(M_objrs!jml), "", M_objrs!jml))
    LAgent = CStr(IIf(IsNull(M_objrs!RECSOURCE), "", M_objrs!RECSOURCE))
    CMDSQL = "Update TrackingRptPerPrgBatch set CALLSINITIATED =" + JUMLAH + " where BATCH ='" + LAgent + "'"
    If JUMLAH < "1" Then
    Else
        M_RPTCONN.Execute CMDSQL
    End If
    M_objrs.MoveNext
Wend
Set M_objrs = Nothing
Exit Sub
hitung_BatchCallInitilizeder:
    MsgBox err.Description
End Sub

' batas akhir
Private Sub Isi_Agent_mgm()
Dim M_objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
On Error GoTo Isi_AgentErr

M_RPTCONN.Execute "Delete * from TrackingRptPerPrgBatch"

M_objrs.CursorLocation = adUseClient
m_objrs1.CursorLocation = adUseClient
If Option1(1).Value = True Then
    M_objrs.Open "Select * from usertbl where AKTIF = 0 AND SPVCODE >='" + Combo3(0).Text + "' and SPVCODE <= '" + Combo3(1).Text + "' AND USERTYPE =1", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    M_objrs.Open "Select * from usertbl where AKTIF = 0 AND userid >='" + Combo2(0).Text + "' and userid <= '" + Combo2(1).Text + "' AND USERTYPE =1", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
End If
m_objrs1.Open "Select * from TrackingRptPerPrgBatch", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_objrs.RecordCount + 2
While Not M_objrs.EOF
    m_objrs1.AddNew
    m_objrs1!TEAM = CStr(IIf(IsNull(M_objrs!TEAM), "", M_objrs!TEAM))
    m_objrs1!TSRNAME = CStr(IIf(IsNull(M_objrs!agent), "", M_objrs!agent))
    m_objrs1!TEAM = CStr(IIf(IsNull(M_objrs!SPVCODE), "", M_objrs!SPVCODE))
    m_objrs1!aoc = CStr(IIf(IsNull(M_objrs!USERID), "", M_objrs!USERID))
    m_objrs1.update
    M_objrs.MoveNext
Wend
Set M_objrs = Nothing
Set m_objrs1 = Nothing
Exit Sub

Isi_AgentErr:
    MsgBox err.Description
End Sub

Private Sub hitung_JmlData_PerAgent_mgm()

Dim M_objrs As New ADODB.Recordset
Dim JUMLAH As String
Dim JUMLAHVOL As String
Dim batch As String
Dim CMDSQL As String
Dim LAgent As String
Dim m_msgbox As Variant
On Error GoTo hitung_JmlDataer
M_objrs.CursorLocation = adUseClient
'cmdsql = "Select Agent, count(custid) as jml from mgm  where left(recsource,3) <>'PRE' and recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "' AND tglsource between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
CMDSQL = "Select Agent, count(custid) as jml,sum(AmountWo) as JMLVOL from mgm  where recsource between'" + Combo1(0).Text + "' and '" + Combo1(1).Text + "' AND tglsource <= '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "' group by Agent"
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_objrs.RecordCount + 2
While Not M_objrs.EOF
    ProgressBar1.Value = M_objrs.Bookmark
    JUMLAH = CStr(IIf(IsNull(M_objrs!jml), "0", M_objrs!jml))
    JUMLAHVOL = CStr(IIf(IsNull(M_objrs!JMLVOL), "0", M_objrs!JMLVOL))
    LAgent = CStr(IIf(IsNull(M_objrs!agent), "", M_objrs!agent))
    CMDSQL = "Update TrackingRptPerPrgBatch set DATASIZE =" + JUMLAH + ", JMLVOL= " + JUMLAHVOL + "  where AOC ='" + LAgent + "'"
    M_RPTCONN.Execute CMDSQL
    M_objrs.MoveNext
Wend
Set M_objrs = Nothing
Exit Sub
hitung_JmlDataer:
    If err.number = -2147217871 Then
        m_msgbox = MsgBox(err.Description, vbRetryCancel, "Aplikasi")
        If m_msgbox = vbRetry Then
            WaitSecs (3)
            Resume
        End If
    Else
        MsgBox err.Description
    End If
    Resume Next
End Sub
Private Sub Hitung_TrackingReportPerAgent_mgm()
Dim M_objrs As ADODB.Recordset
Dim m_objrs1 As ADODB.Recordset
Dim CMDSQL As String
Dim LAgent As String

On Error GoTo Hitung_TrackingReportErr
    Me.MousePointer = vbHourglass
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    CMDSQL = "Select AGENT, kethslkerja, count(AGENT) as jumlah from mgm where tglcall between '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "' and recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "'group by AGENT, kethslkerja"
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = M_objrs.RecordCount + 1
    While Not M_objrs.EOF
        ProgressBar1.Value = M_objrs.Bookmark
'         WaitSecs (0.5)
        LAgent = CStr(IIf(IsNull(M_objrs!agent), "", M_objrs!agent))
        CMDSQL = "Update TrackingRptPerPrgBatch Set "
        CMDSQL = CMDSQL + "[" + Trim(CStr(M_objrs!KETHSLKERJA)) + "]"
        CMDSQL = CMDSQL + "=" + CStr(IIf(IsNull(M_objrs!JUMLAH), 0, M_objrs!JUMLAH)) + ""
        CMDSQL = CMDSQL + " Where AOC = '" + LAgent + "'"
        If IsNull(M_objrs!KETHSLKERJA) Then
        Else
            If M_objrs!KETHSLKERJA = "[]" Then
            Else
                If M_objrs!JUMLAH = 0 Then
                Else
                   
                    M_RPTCONN.Execute CMDSQL
                End If
            End If
        End If
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
    Me.MousePointer = vbNormal
Exit Sub
Hitung_TrackingReportErr:
Me.MousePointer = vbNormal
MsgBox err.Description
'Resume
End Sub

Private Sub hitung_BatchCallInitilized_PerAgent_mgm()
Dim M_objrs As New ADODB.Recordset
Dim JUMLAH As String
Dim batch As String
Dim CMDSQL As String
Dim LAgent As String
Dim m_msgbox As Variant

On Error GoTo hitung_BatchCallInitilizeder
M_objrs.CursorLocation = adUseClient
CMDSQL = "Select agent,count(agent) as jml from mgm_hst where tgl between '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  and custid in(select custid from mgm where recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "')group by agent order by  agent"
'CMDSQL = "Select userid,count(userid) as jml from mgm_hst where tgl between '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  and left(RecSource,3) <> 'PRE' and custid in(select custid from mgm where recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "')group by userid"
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_objrs.RecordCount + 2
While Not M_objrs.EOF
    ProgressBar1.Value = M_objrs.Bookmark
    JUMLAH = CStr(IIf(IsNull(M_objrs!jml), "", M_objrs!jml))
    LAgent = CStr(IIf(IsNull(M_objrs!agent), "", M_objrs!agent))
    CMDSQL = "Update TrackingRptPerPrgBatch set CALLSINITIATED =" + JUMLAH + " where AOC ='" + LAgent + "'"
    If JUMLAH < "1" Then
    Else
        M_RPTCONN.Execute CMDSQL
    End If
    M_objrs.MoveNext
Wend
Set M_objrs = Nothing
Exit Sub
hitung_BatchCallInitilizeder:
m_msgbox = MsgBox(err.Description, vbRetryCancel, "Aplikasi")
If m_msgbox = vbRetry Then
    WaitSecs (3)
    Resume
End If
End Sub
Private Sub Hitung_Number_of_Payment()
Dim M_objrs As New ADODB.Recordset
Dim JUMLAH As String
Dim batch As String
Dim CMDSQL As String
Dim LAgent As String
Dim m_msgbox As Variant
On Error GoTo hitung_JmlDataer

    M_objrs.CursorLocation = adUseClient

CMDSQL = "select agent, count(custid) as jml from (select distinct custid,agent from HtgNumberOfPayment where paydate BETWEEN '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' AND '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "'  and recsource between'" + Combo1(0).Text + "' and '" + Combo1(1).Text + "') a group by agent"
'CMDSQL = CMDSQL + " and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'"

   ' CMDSQL = "Select Agent, count(custid) as jml,sum(AmountWo) as JMLVOL from mgm  where F_CEK ='PTP' AND recsource >='" + Combo1(0).Text + "' and recsource <='" + Combo1(1).Text + "' AND TGLINCOMING BETWEEN '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " " & DTimeLastCall(0).Value & "' AND '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " " & DTimeLastCall(1).Value & "'  group by Agent"
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = M_objrs.RecordCount + 2
    While Not M_objrs.EOF
        ProgressBar1.Value = M_objrs.Bookmark
        JUMLAH = CStr(IIf(IsNull(M_objrs!jml), "0", M_objrs!jml))
        LAgent = CStr(IIf(IsNull(M_objrs!agent), "", M_objrs!agent))
        CMDSQL = "Update TrackingRptPerPrgBatch set NPayment =" + JUMLAH + " where AOC ='" + LAgent + "'"
        M_RPTCONN.Execute CMDSQL
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
    Exit Sub
hitung_JmlDataer:
        If err.number = -2147217871 Then
            m_msgbox = MsgBox(err.Description, vbRetryCancel, "Aplikasi")
            If m_msgbox = vbRetry Then
                WaitSecs (3)
                Resume
            End If
        Else
            MsgBox err.Description
        End If

End Sub

Private Sub Hitung_Volume_of_Payment()
Dim M_objrs As New ADODB.Recordset
Dim JUMLAH As String
Dim batch As String
Dim CMDSQL As String
Dim LAgent As String
Dim LRECSOURCE As String
Dim m_msgbox As Variant

On Error GoTo hitung_JmlDataer

    M_objrs.CursorLocation = adUseClient

CMDSQL = "select agent, sum(Payment) as jml from (select * from HtgVolumeOfPayment where paydate BETWEEN '" + Format(TDBDate1(0).Value, "yyyy-mm-dd") & " " & DTimeLastCall(0).Value & "' AND '" + Format(TDBDate1(1).Value, "yyyy-mm-dd") & " " & DTimeLastCall(1).Value & "' and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "') a group by agent"
'CMDSQL = CMDSQL + " and recsource between '" + Combo1(0).Text + "' and '" + Combo1(1).Text + "'"
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = M_objrs.RecordCount + 2
    While Not M_objrs.EOF
        ProgressBar1.Value = M_objrs.Bookmark
        JUMLAH = CStr(IIf(IsNull(M_objrs!jml), "0", M_objrs!jml))
        LAgent = CStr(IIf(IsNull(M_objrs!agent), "", M_objrs!agent))
        CMDSQL = "Update TrackingRptPerPrgBatch set VolPayment =" + JUMLAH + " where AOC ='" + LAgent + "'"
        M_RPTCONN.Execute CMDSQL
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
    Exit Sub
hitung_JmlDataer:
        If err.number = -2147217871 Then
            m_msgbox = MsgBox(err.Description, vbRetryCancel, "Aplikasi")
            If m_msgbox = vbRetry Then
                WaitSecs (3)
                Resume
            End If
        Else
            MsgBox err.Description
        End If
End Sub

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




Private Sub hitung_BatchCallInitilized_PerAgent_Leads()
Dim M_objrs As New ADODB.Recordset
Dim JUMLAH As String
Dim batch As String
Dim CMDSQL As String
Dim LAgent As String
On Error GoTo hitung_BatchCallInitilizeder

M_objrs.CursorLocation = adUseClient
CMDSQL = "Select userid,count(userid) as jml from TblPhoneMonitorHst where tgl between '" + Format(TDBDate1(0).Value, "mm/dd/yyyy") & " " & DTimeLastCall(0).Value & "' and '" + Format(TDBDate1(1).Value, "mm/dd/yyyy") & " " & DTimeLastCall(1).Value & "' and recsource between 'mgm-REF' and 'mgm-REF'  group by userid"
M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = M_objrs.RecordCount + 2
While Not M_objrs.EOF
    ProgressBar1.Value = M_objrs.Bookmark
    JUMLAH = CStr(IIf(IsNull(M_objrs!jml), "", M_objrs!jml))
    LAgent = CStr(IIf(IsNull(M_objrs!USERID), "", M_objrs!USERID))
    CMDSQL = "Update TrackingRptPerPrgBatch set CALLSINITIATED =" + JUMLAH + " where AOC ='" + LAgent + "'"
    If JUMLAH < "1" Then
    Else
        M_RPTCONN.Execute CMDSQL
    End If
    M_objrs.MoveNext
Wend
Set M_objrs = Nothing
Exit Sub
hitung_BatchCallInitilizeder:
    MsgBox err.Description
End Sub

Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "No", 4 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Report", 50 * TXT
End Sub

Private Sub Form_Load()
Dim listitem As listitem
Dim M_objrs As ADODB.Recordset
Set M_objrs = New ADODB.Recordset
DTimeLastCall(0).Text = "00:00"
DTimeLastCall(1).Text = "23:59"
M_objrs.CursorLocation = adUseClient
CmbCek.AddItem "Not Check"
CmbCek.AddItem "Accept"
CmbCek.AddItem "RETURN"

 Option1(0).Value = True
Option1(1).Visible = False
Combo2(0).Text = MDIForm1.TxtUsername.Text
Combo2(1).Text = MDIForm1.TxtUsername.Text
Combo2(0).Enabled = False
Combo2(1).Enabled = False


M_objrs.Open "SELECT * FROM usertbl WHERE AKTIF = 0 ORDER BY USERID", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not M_objrs.EOF
    Combo2(0).AddItem M_objrs!USERID
    Combo2(1).AddItem M_objrs!USERID
    M_objrs.MoveNext
Wend
Set M_objrs = Nothing

Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open "SELECT * FROM DATASOURCETBL ORDER BY KODEDS", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not M_objrs.EOF
    Combo1(0).AddItem M_objrs!KODEDS
    Combo1(1).AddItem M_objrs!KODEDS
    M_objrs.MoveNext
Wend
Set M_objrs = Nothing

Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
'm_objrs.Open "SELECT * FROM spvtbl ORDER BY SPVCODE", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
M_objrs.Open "select distinct SPVTBL.SPVCODE from SPVTBL, usertbl where SPVTBL.SPVCODE = usertbl.SPVCODE AND USERTYPE = '6'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

While Not M_objrs.EOF
    Combo3(0).AddItem M_objrs!SPVCODE
    Combo3(1).AddItem M_objrs!SPVCODE
    M_objrs.MoveNext
Wend
Set M_objrs = Nothing

Call header
'Set LISTITEM = ListView1.ListItems.ADD(, , "1")
'    LISTITEM.SubItems(1) = "Individual Call List Data Tracking Summary"
'Set LISTITEM = ListView1.ListItems.ADD(, , "2")
'    LISTITEM.SubItems(1) = "Individual Leads Data Tracking Summary"
'Set LISTITEM = ListView1.ListItems.ADD(, , "3")
'    LISTITEM.SubItems(1) = "Team Leads Data Tracking Summary"
'Set LISTITEM = ListView1.ListItems.ADD(, , "4")
'    LISTITEM.SubItems(1) = "Team Productivity Based On Data Source"
'Set LISTITEM = ListView1.ListItems.ADD(, , "5")
'    LISTITEM.SubItems(1) = "Team Leads Productivity Based On Data Source"
'Set LISTITEM = ListView1.ListItems.ADD(, , "6")
'    LISTITEM.SubItems(1) = "Team Call List Data Tracking Summary"
'Set listitem = listview1.ListItems.ADD(, , "1")
'    listitem.SubItems(1) = "Tracking Report Daily PerAgent"
Set listitem = ListView1.ListItems.ADD(, , "2")
    listitem.SubItems(1) = "Report PTP Due Date PerAgent"
'Set listitem = ListView1.ListItems.ADD(, , "3")
'    listitem.SubItems(1) = "Report Detail PTP"
'Set listitem = ListView1.ListItems.ADD(, , "8")
'    listitem.SubItems(1) = "Application Submitted"
''Set LISTITEM = ListView1.ListItems.ADD(, , "9")
''    LISTITEM.SubItems(1) = "Daily Submission Report"
End Sub


Private Sub ListView1_Click()
    Label2.Caption = ListView1.SelectedItem.SubItems(1)
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
    Case 0
        If Option1(Index).Value = False Then
            Option1(1).Value = False
        Else
            Combo2(0).Enabled = True
            Combo2(1).Enabled = True
            Combo3(0).Enabled = False
            Combo3(1).Enabled = False
        End If
    Case 1
        If Option1(Index).Value = False Then
            Option1(0).Value = False
        Else
            Combo2(0).Enabled = False
            Combo2(1).Enabled = False
            Combo3(0).Enabled = True
            Combo3(1).Enabled = True
        End If
End Select
End Sub
