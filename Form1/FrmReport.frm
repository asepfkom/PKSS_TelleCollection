VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "TIDATE6.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmReport 
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5535
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   4605
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport RPT 
      Left            =   1905
      Top             =   3900
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Data Quality Report  (Per Batch)"
      Height          =   300
      Index           =   2
      Left            =   300
      TabIndex        =   0
      Top             =   150
      Width           =   2655
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   300
      Left            =   135
      TabIndex        =   21
      Top             =   4080
      Visible         =   0   'False
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Performance Report (Per TSR)"
      Height          =   300
      Index           =   1
      Left            =   285
      TabIndex        =   6
      Top             =   2430
      Width           =   2580
   End
   Begin VB.Frame Frame1 
      Height          =   1380
      Index           =   1
      Left            =   135
      TabIndex        =   16
      Top             =   2460
      Width           =   5370
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   0
         Left            =   1635
         TabIndex        =   7
         Top             =   390
         Width           =   1305
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   1
         Left            =   3330
         TabIndex        =   8
         Top             =   390
         Width           =   1305
      End
      Begin TDBDate6Ctl.TDBDate TDBDate1 
         Height          =   315
         Index           =   1
         Left            =   3330
         TabIndex        =   10
         Top             =   780
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   556
         Calendar        =   "FrmReport.frx":0000
         Caption         =   "FrmReport.frx":0118
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmReport.frx":0184
         Keys            =   "FrmReport.frx":01A2
         Spin            =   "FrmReport.frx":0200
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
         Left            =   1620
         TabIndex        =   9
         Top             =   780
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   556
         Calendar        =   "FrmReport.frx":0228
         Caption         =   "FrmReport.frx":0340
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmReport.frx":03AC
         Keys            =   "FrmReport.frx":03CA
         Spin            =   "FrmReport.frx":0428
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
         Alignment       =   1  'Right Justify
         Caption         =   "Date :"
         Height          =   300
         Index           =   5
         Left            =   645
         TabIndex        =   20
         Top             =   840
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "to"
         Height          =   300
         Index           =   4
         Left            =   3075
         TabIndex        =   19
         Top             =   825
         Width           =   270
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "From TSR :"
         Height          =   300
         Index           =   3
         Left            =   660
         TabIndex        =   18
         Top             =   420
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "to"
         Height          =   300
         Index           =   2
         Left            =   3045
         TabIndex        =   17
         Top             =   420
         Width           =   270
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Tracking Report  (Per Batch)"
      Height          =   300
      Index           =   0
      Left            =   285
      TabIndex        =   3
      Top             =   1290
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Height          =   1110
      Index           =   0
      Left            =   135
      TabIndex        =   13
      Top             =   1320
      Width           =   5370
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   3360
         TabIndex        =   5
         Top             =   390
         Width           =   1305
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1650
         TabIndex        =   4
         Top             =   390
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "to"
         Height          =   300
         Index           =   1
         Left            =   3060
         TabIndex        =   15
         Top             =   420
         Width           =   285
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "From Batch :"
         Height          =   300
         Index           =   0
         Left            =   675
         TabIndex        =   14
         Top             =   420
         Width           =   930
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   315
      Index           =   1
      Left            =   4440
      TabIndex        =   12
      Top             =   4065
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Preview"
      Height          =   315
      Index           =   0
      Left            =   3315
      TabIndex        =   11
      Top             =   4065
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Height          =   1110
      Index           =   2
      Left            =   135
      TabIndex        =   22
      Top             =   180
      Width           =   5370
      Begin VB.ComboBox Combo3 
         Height          =   315
         Index           =   0
         Left            =   1650
         TabIndex        =   1
         Top             =   390
         Width           =   1305
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Index           =   1
         Left            =   3360
         TabIndex        =   2
         Top             =   390
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "From Batch :"
         Height          =   300
         Index           =   7
         Left            =   675
         TabIndex        =   24
         Top             =   420
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "to"
         Height          =   300
         Index           =   6
         Left            =   3060
         TabIndex        =   23
         Top             =   420
         Width           =   285
      End
   End
End
Attribute VB_Name = "FrmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo3_Click(Index As Integer)
    Call Combo3_LostFocus(Index)
End Sub

Private Sub Combo3_LostFocus(Index As Integer)
Dim m_objrs As New ADODB.Recordset
On Error GoTo Combo3_LostFocusErr
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open "Select * from datasourcetbl where kodeds ='" + Combo3(Index).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If Not m_objrs.EOF Then
        Combo3(Index).Text = m_objrs!KODEDS
    Else
        Combo3(Index).Text = Empty
    End If
Exit Sub
Combo3_LostFocusErr:
    MsgBox Err.Description
End Sub

Private Sub Combo1_Click(Index As Integer)
    Call Combo1_LostFocus(Index)
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim m_objrs As New ADODB.Recordset
On Error GoTo Combo1_LostFocusErr
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open "Select * from datasourcetbl where kodeds ='" + Combo1(Index).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If Not m_objrs.EOF Then
        Combo1(Index).Text = m_objrs!KODEDS
    Else
        Combo1(Index).Text = Empty
    End If
Exit Sub
Combo1_LostFocusErr:
    MsgBox Err.Description
End Sub

Private Sub Combo2_Click(Index As Integer)
    Call Combo2_LostFocus(Index)
End Sub

Private Sub Combo2_LostFocus(Index As Integer)
Dim m_objrs As New ADODB.Recordset
On Error GoTo Combo2_LostFocusErr
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open "Select * from USERTBL where USERID ='" + Combo2(Index).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If Not m_objrs.EOF Then
        Combo2(Index).Text = m_objrs!USERID
    Else
        Combo2(Index).Text = Empty
    End If
Exit Sub
Combo2_LostFocusErr:
    MsgBox Err.Description
End Sub

Private Sub Command1_Click(Index As Integer)
On Error GoTo Command1_ClickeR
Select Case Index
    Case 0
        If Option1(0).Value Then
        If Combo1(0).Text = Empty Or Combo1(1).Text = Empty Then
            MsgBox "Batch Database Code Harus DiIsi", vbInformation + vbOKOnly, "Telegrandi"
            Combo1(0).SetFocus
            Exit Sub
        End If
            ProgressBar1.Visible = True
            Call ISI_DATASOURCE
            Call Hitung_TrackingReport
            Call hitung_JmlData
            Call hitung_BatchCallInitilized
            RPT.Reset
            RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
            RPT.ReportFileName = App.Path + "\Report\TrackingRptPerPrgBatch.rpt"
            Call SHOW_PRN
        Else
            If Option1(1).Value Then
    '            Call HItung_PerformanceReport
                If Combo2(0).Text = Empty Or Combo2(1).Text = Empty Or TDBDate1(0).ValueIsNull Or TDBDate1(1).ValueIsNull Then
                    MsgBox "Tanggal Dan Batch Database Code Harus DiIsi", vbInformation + vbOKOnly, "Telegrandi"
                    Combo2(0).SetFocus
                    Exit Sub
                End If
                ProgressBar1.Visible = True
                Call Isi_Agent
                Call hitung_JmlData_PerAgent
                Call Hitung_TrackingReportPerAgent
                Call hitung_BatchCallInitilized_PerAgent
                RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
                RPT.ReportFileName = App.Path + "\Report\PerformanceRpt.rpt"
                Call SHOW_PRN
                
            Else
                If Combo3(0).Text = Empty Or Combo3(1).Text = Empty Then
                    MsgBox "Batch Database Code Harus DiIsi", vbInformation + vbOKOnly, "Telegrandi"
                    Combo1(0).SetFocus
                    Exit Sub
                End If
                ProgressBar1.Visible = True
                Call ISI_DATASOURCE
                Call Hitung_TrackingReport
                Call Hitung_DataQuality
                Call hitung_JmlData
                RPT.Formulas(1) = "@User = totext('" + CStr(MDIForm1.Text1.Text) + "')"
                RPT.ReportFileName = App.Path + "\Report\DataQuality.rpt"
                Call SHOW_PRN
            End If
        End If
    Case 1
        Unload Me
End Select
    ProgressBar1.Visible = False
Exit Sub
Command1_ClickeR:
    MsgBox Err.Description
End Sub

Private Sub Isi_Agent()
Dim m_objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
Dim M_DATASOURCE As ADODB.Recordset
On Error GoTo Isi_AgentErr

M_RPTCONN.Execute "Delete * from TrackingRptPerPrgBatch"

m_objrs.CursorLocation = adUseClient
m_objrs1.CursorLocation = adUseClient

m_objrs.Open "Select * from UserTbl where userid >='" + Combo2(0).Text + "' and userid <= '" + Combo2(1).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
m_objrs1.Open "Select * from TrackingRptPerPrgBatch", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_objrs.RecordCount + 2
While Not m_objrs.EOF
    ProgressBar1.Value = m_objrs.Bookmark
    Set M_DATASOURCE = New ADODB.Recordset
    M_DATASOURCE.CursorLocation = adUseClient
    M_DATASOURCE.Open "Select * from DataSourceTbl", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    While Not M_DATASOURCE.EOF
        m_objrs1.AddNew
        m_objrs1!TEAM = CStr(IIf(IsNull(m_objrs!TEAM), "", m_objrs!TEAM))
        m_objrs1!TSRNAME = CStr(m_objrs!USERID)
        m_objrs1!batch = CStr(M_DATASOURCE!KODEDS)
        m_objrs1.UPDATE
        M_DATASOURCE.MoveNext
    Wend
    Set M_DATASOURCE = Nothing
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
Set m_objrs1 = Nothing
Exit Sub

Isi_AgentErr:
    MsgBox Err.Description
End Sub

Private Sub hitung_JmlData_PerAgent()
Dim m_objrs As New ADODB.Recordset
Dim JUMLAH As String
Dim batch As String
Dim CMDSQL As String
Dim LAgent As String

On Error GoTo hitung_JmlDataer
m_objrs.CursorLocation = adUseClient
m_objrs.Open "Select Agent, recsource, count(custid) as jml from cc_custtbl  group by Agent, recsource", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_objrs.RecordCount + 2
While Not m_objrs.EOF
    ProgressBar1.Value = m_objrs.Bookmark
    JUMLAH = CStr(IIf(IsNull(m_objrs!jml), "", m_objrs!jml))
    batch = CStr(IIf(IsNull(m_objrs!RECSOURCE), "", m_objrs!RECSOURCE))
    LAgent = CStr(IIf(IsNull(m_objrs!agent), "", m_objrs!agent))
    CMDSQL = "Update TrackingRptPerPrgBatch set DATASIZE =" + JUMLAH + " where BATCH ='" + batch + "' and TSRNAME ='" + LAgent + "'"
    M_RPTCONN.Execute CMDSQL
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
Exit Sub
hitung_JmlDataer:
MsgBox Err.Description
End Sub

Private Sub Hitung_TrackingReportPerAgent()
Dim m_objrs As ADODB.Recordset
Dim m_objrs1 As ADODB.Recordset
Dim CMDSQL As String
Dim LAgent As String

On Error GoTo Hitung_TrackingReportErr
    Me.MousePointer = vbHourglass
    Set m_objrs = New ADODB.Recordset
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open "Select Agent, recsource, kethslkerja, count(custid) as jumlah from cc_custtbl where tglstatus >= '" + Format(TDBDate1(0).Text, "mm/dd/yy") + "' and tglstatus <= '" + Format(TDBDate1(1).Text, "mm/dd/yy") + "' group by Agent, recsource, kethslkerja", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = m_objrs.RecordCount + 2
    While Not m_objrs.EOF
        ProgressBar1.Value = m_objrs.Bookmark
        LAgent = CStr(IIf(IsNull(m_objrs!agent), "", m_objrs!agent))
        CMDSQL = "Update TrackingRptPerPrgBatch Set "
        CMDSQL = CMDSQL + m_objrs!KETHSLKERJA
        CMDSQL = CMDSQL + "=" + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH)) + ""
        CMDSQL = CMDSQL + " where BATCH = '" + IIf(IsNull(m_objrs!RECSOURCE), "", m_objrs!RECSOURCE) + "'"
        CMDSQL = CMDSQL + " and TSRNAME = '" + LAgent + "'"
        If IsNull(m_objrs!KETHSLKERJA) Then
        Else
            If m_objrs!KETHSLKERJA = Empty Then
            Else
                If m_objrs!JUMLAH = 0 Then
                Else
                    M_RPTCONN.Execute CMDSQL
                End If
            End If
        End If
        m_objrs.MoveNext
    Wend
    Set m_objrs = Nothing
    Me.MousePointer = vbNormal
Exit Sub
Hitung_TrackingReportErr:
Me.MousePointer = vbNormal
MsgBox Err.Description
End Sub

Private Sub hitung_BatchCallInitilized_PerAgent()
Dim m_objrs As New ADODB.Recordset
Dim JUMLAH As String
Dim batch As String
Dim CMDSQL As String
Dim LAgent As String
On Error GoTo hitung_BatchCallInitilizeder
m_objrs.CursorLocation = adUseClient
m_objrs.Open "Select Agent,recsource, count(recsource) as jml from PHONENO_CALL where TGL >= '" + Format(TDBDate1(0).Text, "mm/dd/yy") + "' and TGL <= '" + Format(TDBDate1(1).Text, "mm/dd/yy") + "' group by Agent, recsource", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_objrs.RecordCount + 2
While Not m_objrs.EOF
    ProgressBar1.Value = m_objrs.Bookmark
    JUMLAH = CStr(IIf(IsNull(m_objrs!jml), "", m_objrs!jml))
    batch = CStr(IIf(IsNull(m_objrs!RECSOURCE), "", m_objrs!RECSOURCE))
    LAgent = CStr(IIf(IsNull(m_objrs!agent), "", m_objrs!agent))
    CMDSQL = "Update TrackingRptPerPrgBatch set CALLSINITIATED =" + JUMLAH + " where BATCH ='" + batch + "' and TSRNAME ='" + LAgent + "'"
    If JUMLAH < "1" Then
    Else
        M_RPTCONN.Execute CMDSQL
    End If
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
Exit Sub
hitung_BatchCallInitilizeder:
    MsgBox Err.Description
End Sub

Private Sub Hitung_DataQuality()
Dim m_objrs As ADODB.Recordset
Dim m_objrs1 As ADODB.Recordset
Dim CMDSQL As String
On Error GoTo Hitung_TrackingReportErr
    Me.MousePointer = vbHourglass
    Set m_objrs = New ADODB.Recordset
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open "Select recsource, kd_cls, count(custid) as jumlah from cc_custtbl group by recsource, kd_cls", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = m_objrs.RecordCount + 2
    While Not m_objrs.EOF
        ProgressBar1.Value = m_objrs.Bookmark
        CMDSQL = "Update TrackingRptPerPrgBatch Set "
        CMDSQL = CMDSQL + CStr(m_objrs!KD_CLS)
        CMDSQL = CMDSQL + "=" + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH)) + ""
        CMDSQL = CMDSQL + " where BATCH = '" + IIf(IsNull(m_objrs!RECSOURCE), "", m_objrs!RECSOURCE) + "'"
        If IsNull(m_objrs!KD_CLS) Then
        Else
            If m_objrs!KD_CLS = Empty Then
            Else
            M_RPTCONN.Execute CMDSQL
            End If
        End If
        m_objrs.MoveNext
    Wend
    Set m_objrs = Nothing
    Call hitung_JmlData
    Call hitung_BatchCallInitilized
    Me.MousePointer = vbNormal
Exit Sub
Hitung_TrackingReportErr:
Me.MousePointer = vbNormal
MsgBox Err.Description
End Sub

Private Sub ISI_DATASOURCE()
Dim m_objrs As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset

m_objrs.CursorLocation = adUseClient
m_objrs1.CursorLocation = adUseClient
M_RPTCONN.Execute "Delete * from TrackingRptPerPrgBatch"
If Option1(0).Value Then
    m_objrs.Open "Select * from datasourcetbl where kodeds >= '" + Combo1(0).Text + "' and kodeds <= '" + Combo1(1).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
Else
    m_objrs.Open "Select * from datasourcetbl where kodeds >= '" + Combo3(0).Text + "' and kodeds <= '" + Combo3(1).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
End If

m_objrs1.Open "Select * from TrackingRptPerPrgBatch", M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_objrs.RecordCount + 2
    
While Not m_objrs.EOF
    ProgressBar1.Value = m_objrs.Bookmark
    m_objrs1.AddNew
    m_objrs1![batch] = CStr(m_objrs!KODEDS)
    m_objrs1.UPDATE
    m_objrs.MoveNext
Wend

Set m_objrs = Nothing
Set m_objrs1 = Nothing

End Sub

Private Sub Hitung_TrackingReport()
Dim m_objrs As ADODB.Recordset
Dim m_objrs1 As ADODB.Recordset
Dim CMDSQL As String
On Error GoTo Hitung_TrackingReportErr
    Me.MousePointer = vbHourglass
    Set m_objrs = New ADODB.Recordset
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open "Select recsource, kethslkerja, count(custid) as jumlah from cc_custtbl group by recsource, kethslkerja", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ProgressBar1.Max = m_objrs.RecordCount + 2
    While Not m_objrs.EOF
        ProgressBar1.Value = m_objrs.Bookmark
        CMDSQL = "Update TrackingRptPerPrgBatch Set "
        CMDSQL = CMDSQL + m_objrs!KETHSLKERJA
        CMDSQL = CMDSQL + "=" + CStr(IIf(IsNull(m_objrs!JUMLAH), 0, m_objrs!JUMLAH)) + ""
        CMDSQL = CMDSQL + " where BATCH = '" + IIf(IsNull(m_objrs!RECSOURCE), "", m_objrs!RECSOURCE) + "'"
        If IsNull(m_objrs!KETHSLKERJA) Then
        Else
            If m_objrs!KETHSLKERJA = Empty Then
            Else
            M_RPTCONN.Execute CMDSQL
            End If
        End If
        m_objrs.MoveNext
    Wend
    Set m_objrs = Nothing
    Me.MousePointer = vbNormal
Exit Sub
Hitung_TrackingReportErr:
Me.MousePointer = vbNormal
MsgBox Err.Description
End Sub

Private Sub hitung_JmlData()
Dim m_objrs As New ADODB.Recordset
Dim JUMLAH As String
Dim batch As String
Dim CMDSQL As String

On Error GoTo hitung_JmlDataer
m_objrs.CursorLocation = adUseClient
m_objrs.Open "Select recsource, count(custid) as jml from cc_custtbl group by recsource", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_objrs.RecordCount + 2
While Not m_objrs.EOF
    ProgressBar1.Value = m_objrs.Bookmark
    JUMLAH = CStr(IIf(IsNull(m_objrs!jml), "", m_objrs!jml))
    batch = CStr(IIf(IsNull(m_objrs!RECSOURCE), "", m_objrs!RECSOURCE))
    CMDSQL = "Update TrackingRptPerPrgBatch set DATASIZE =" + JUMLAH + " where BATCH ='" + batch + "'"
    M_RPTCONN.Execute CMDSQL
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
Exit Sub
hitung_JmlDataer:
MsgBox Err.Description
End Sub

Private Sub hitung_BatchCallInitilized()
Dim m_objrs As New ADODB.Recordset
Dim JUMLAH As String
Dim batch As String
Dim CMDSQL As String

On Error GoTo hitung_BatchCallInitilizeder
m_objrs.CursorLocation = adUseClient
m_objrs.Open "Select recsource, count(recsource) as jml from PHONENO_CALL group by recsource", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
ProgressBar1.Max = m_objrs.RecordCount + 2
While Not m_objrs.EOF
    ProgressBar1.Value = m_objrs.Bookmark
    JUMLAH = CStr(IIf(IsNull(m_objrs!jml), "", m_objrs!jml))
    batch = CStr(IIf(IsNull(m_objrs!RECSOURCE), "", m_objrs!RECSOURCE))
    CMDSQL = "Update TrackingRptPerPrgBatch set CALLSINITIATED =" + JUMLAH + " where BATCH ='" + batch + "'"
    M_RPTCONN.Execute CMDSQL
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
Exit Sub
hitung_BatchCallInitilizeder:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
Dim m_objrs As ADODB.Recordset
Me.MousePointer = vbHourglass

On Error GoTo Show_Form
Option1(2).Value = True
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open "Select * from DATASOURCETBL ORDER BY KODEDS", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not m_objrs.EOF
    Combo1(0).AddItem IIf(IsNull(m_objrs!KODEDS), "", m_objrs!KODEDS)
    Combo1(1).AddItem IIf(IsNull(m_objrs!KODEDS), "", m_objrs!KODEDS)
    Combo3(0).AddItem IIf(IsNull(m_objrs!KODEDS), "", m_objrs!KODEDS)
    Combo3(1).AddItem IIf(IsNull(m_objrs!KODEDS), "", m_objrs!KODEDS)
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing

Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open "Select * from usertbl ORDER BY USERID", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not m_objrs.EOF
    Combo2(0).AddItem IIf(IsNull(m_objrs!USERID), "", m_objrs!USERID)
    Combo2(1).AddItem IIf(IsNull(m_objrs!USERID), "", m_objrs!USERID)
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
    Me.MousePointer = vbNormal
Exit Sub
Show_Form:
    Me.MousePointer = vbNormal
    MsgBox Err.Description
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
    Case 0
        Combo2(0).Text = Empty
        Combo2(1).Text = Empty
        Combo3(0).Text = Empty
        Combo3(1).Text = Empty
        TDBDate1(0).Value = Empty
        TDBDate1(1).Value = Empty
            
        Frame1(0).Enabled = True
        Frame1(1).Enabled = False
        Frame1(2).Enabled = False
    Case 1
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
        Combo3(0).Text = Empty
        Combo3(1).Text = Empty
        
        Frame1(1).Enabled = True
        Frame1(0).Enabled = False
        Frame1(2).Enabled = False
    Case 2
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
        Combo2(0).Text = Empty
        Combo2(1).Text = Empty
        TDBDate1(0).Value = Empty
        TDBDate1(1).Value = Empty
        
        Frame1(2).Enabled = True
        Frame1(0).Enabled = False
        Frame1(1).Enabled = False
End Select
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
