VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rptCallTrackingServer4 
   Caption         =   "Report Time on System [Card Server 4]"
   ClientHeight    =   4635
   ClientLeft      =   4335
   ClientTop       =   2655
   ClientWidth     =   5955
   Icon            =   "rptCallTracking.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   5955
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Height          =   4125
      Left            =   0
      TabIndex        =   0
      Top             =   540
      Width           =   5985
      Begin VB.Frame Frame2 
         BackColor       =   &H000080FF&
         Caption         =   "Pilih User"
         Height          =   1875
         Left            =   1560
         TabIndex        =   12
         Top             =   1080
         Width           =   3675
         Begin VB.ComboBox cmbtlfullname 
            Height          =   315
            Left            =   420
            TabIndex        =   16
            Top             =   1320
            Width           =   2985
         End
         Begin VB.OptionButton Opt_Team 
            BackColor       =   &H000080FF&
            Caption         =   "Per Team"
            Height          =   315
            Left            =   120
            TabIndex        =   15
            Top             =   1020
            Width           =   1875
         End
         Begin VB.ComboBox cmbagentfullname 
            Height          =   315
            ItemData        =   "rptCallTracking.frx":058A
            Left            =   420
            List            =   "rptCallTracking.frx":058C
            TabIndex        =   14
            Top             =   600
            Width           =   2985
         End
         Begin VB.OptionButton Opt_Agent 
            BackColor       =   &H000080FF&
            Caption         =   "Per Agent"
            Height          =   315
            Left            =   180
            TabIndex        =   13
            Top             =   300
            Value           =   -1  'True
            Width           =   1875
         End
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "&Exit"
         Default         =   -1  'True
         Height          =   375
         Index           =   1
         Left            =   3900
         TabIndex        =   4
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "&Show Report"
         Height          =   375
         Index           =   0
         Left            =   2520
         TabIndex        =   2
         Top             =   3480
         Width           =   1335
      End
      Begin VB.ComboBox cmbReportName 
         Height          =   315
         ItemData        =   "rptCallTracking.frx":058E
         Left            =   1560
         List            =   "rptCallTracking.frx":059E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   3645
      End
      Begin MSComctlLib.ProgressBar PB 
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   3060
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   661
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin TDBDate6Ctl.TDBDate TdTglCall1 
         Height          =   315
         Left            =   1590
         TabIndex        =   7
         Top             =   330
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   556
         Calendar        =   "rptCallTracking.frx":0616
         Caption         =   "rptCallTracking.frx":072E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "rptCallTracking.frx":079A
         Keys            =   "rptCallTracking.frx":07B8
         Spin            =   "rptCallTracking.frx":0816
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
         Format          =   "dd-mm-yyyy"
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
         Text            =   "__-__-____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   37468
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate TdTglCall2 
         Height          =   315
         Left            =   3900
         TabIndex        =   8
         Top             =   330
         Visible         =   0   'False
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   556
         Calendar        =   "rptCallTracking.frx":083E
         Caption         =   "rptCallTracking.frx":0956
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "rptCallTracking.frx":09C2
         Keys            =   "rptCallTracking.frx":09E0
         Spin            =   "rptCallTracking.frx":0A3E
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
         Format          =   "dd-mm-yyyy"
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
         Text            =   "__-__-____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   37468
         CenturyMode     =   0
      End
      Begin Crystal.CrystalReport RPT 
         Left            =   6990
         Top             =   2520
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "rptCallTracking.frx":0A66
         Left            =   240
         List            =   "rptCallTracking.frx":0A68
         TabIndex        =   11
         Top             =   1620
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Between"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3000
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. call :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   810
         TabIndex        =   9
         Top             =   390
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Report Name"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1395
      End
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   4
      Left            =   60
      Picture         =   "rptCallTracking.frx":0A6A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   420
   End
   Begin VB.Label Label4 
      BackColor       =   &H00F1E5DB&
      BackStyle       =   0  'Transparent
      Caption         =   "Report Call Tracking"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   510
      TabIndex        =   5
      Top             =   60
      Width           =   4335
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "rptCallTracking.frx":1574
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6000
   End
End
Attribute VB_Name = "rptCallTrackingServer4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim connIcentra As New ADODB.Connection
Dim M_RPTCONN As New ADODB.Connection
Public Sub bukakoneksiicentra()
Set connIcentra = New ADODB.Connection
 'Server Icentra 4
 'STRKONEKSI = "Driver={PostgreSQL ANSI}; Server=192.168.10.4; PORT=5432; Database=icentra; UID=icentra; PWD=jengkolman"
 STRKONEKSI = "Driver={PostgreSQL ANSI}; Server=localhost; PORT=5432; Database=icentra_4; UID=admin; PWD=admin321"
 connIcentra.Open STRKONEKSI
End Sub

Public Sub tutupkoneksiicentra()
Set connIcentra = Nothing
End Sub








Private Sub cmbagentfullname_Click()
Dim M_OBJRSnew As New ADODB.Recordset
bukakoneksiicentra
strsql = "select * from ("
strsql = strsql + " select tblbaru.acd_agent_id ,tblbaru.userid,tblbaru.nama,tblbaru.team,acd_agent.fullname from ("
strsql = strsql + " select  acd_agent.name as userid,acd_agent.fullname as nama, acd_queue.name as team,acd_agent.acd_agent_id  from acd_agent_group,acd_agent,acd_queue where acd_agent_group.acd_agent_id=acd_agent.acd_agent_id"
strsql = strsql + " and  acd_agent_group.acd_queue_id=acd_queue.acd_queue_id )  as tblbaru,acd_agent where tblbaru.team=acd_agent.name ) tbluser "
strsql = strsql + "  where  tbluser.userid='" + cmbagentfullname.Text + "' "


    Set M_OBJRSnew = New ADODB.Recordset
    M_OBJRSnew.CursorLocation = adUseClient
    M_OBJRSnew.Open strsql, connIcentra, adOpenDynamic, adLockOptimistic
   ' cmbAgent.CLEAR
    While Not M_OBJRSnew.EOF
        Combo1.Text = IIf(IsNull(M_OBJRSnew!acd_agent_id), "", M_OBJRSnew!acd_agent_id)
        M_OBJRSnew.MoveNext
    Wend
    Set M_OBJRSnew = Nothing
  
    
tutupkoneksiicentra
End Sub









Private Sub cmdexit_Click(Index As Integer)
Unload Me
End Sub

Private Sub cmdShow_Click(Index As Integer)
Select Case Index
Case 0
    If cmbReportName.Text = "" Then MsgBox "Pilih jenis reportnya terlebih dahulu": Exit Sub
    'If cmbagentfullname.Text = "" Then MsgBox "Pilih Agent terlebih dahulu": Exit Sub
    '@@ 31-05-2011 Jika Memilih berdasarkan agent
    If Opt_Agent.Value Then
        If cmbagentfullname.Text = "" Then
            MsgBox "Pilih Agent terlebih dahulu"
        End If
    End If
    '@@ 31-05-2011 Jika Memilih berdasarkan team
    If Opt_Team.Value Then
        If cmbtlfullname.Text = "" Then
            MsgBox "Pilih Team Leader terlebih dahulu"
        End If
    End If
    
    If TdTglCall1.ValueIsNull And TdTglCall2.ValueIsNull Then
        TdTglCall1.Value = "01/01/1990"
        TdTglCall2.Value = "31/12/2020"
    End If

    Select Case Left(cmbReportName.Text, 1)
    Case 1
          sub101
          WaitSecs (2)
          RPT.Reset
         ' RPT.Formulas(1) = "@User = totext('" + CStr(Form1.txtUserName.Text) + "')"
          RPT.Formulas(1) = "@User = totext('Admin')"
          RPT.Formulas(2) = "@TglShow = totext('" + CStr(TdTglCall1.Text) + "')"
          RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TdTglCall2.Text) + "')"
          RPT.ReportFileName = "D:\tnis_report\RptTaltimenew.rpt"
          Call SHOW_PRN
    Case 2
          sub102
          WaitSecs (2)
          RPT.Reset
          ' RPT.Formulas(1) = "@User = totext('" + CStr(Form1.txtUserName.Text) + "')"
          RPT.Formulas(1) = "@User = totext('Admin')"
          RPT.Formulas(2) = "@TglShow = totext('" + CStr(TdTglCall1.Text) + "')"
          RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TdTglCall2.Text) + "')"
          RPT.ReportFileName = "D:\tnis_report\Rptfavorite.rpt"
          Call SHOW_PRN
    Case 3
        sub103
         WaitSecs (2)
          RPT.Reset
         ' RPT.Formulas(1) = "@User = totext('" + CStr(Form1.Text1.Text) + "')"
            RPT.Formulas(1) = "@User = totext('Admin')"
          RPT.Formulas(2) = "@TglShow = totext('" + CStr(TdTglCall1.Text) + "')"
          RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TdTglCall2.Text) + "')"
          RPT.ReportFileName = "D:\tnis_report\RptTaltimenewsumary.rpt"
          Call SHOW_PRN
    Case 4
          sub104
          WaitSecs (2)
          RPT.Reset
         ' RPT.Formulas(1) = "@User = totext('" + CStr(Form1.txtUserName.Text) + "')"
           RPT.Formulas(1) = "@User = totext('Admin')"
          RPT.Formulas(2) = "@TglShow = totext('" + CStr(TdTglCall1.Text) + "')"
          RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TdTglCall2.Text) + "')"
          RPT.ReportFileName = "D:\tnis_report\RptTaltimenewsumarytl.rpt"
          Call SHOW_PRN
    Case 5
          sub105
          WaitSecs (2)
          RPT.Reset
         ' RPT.Formulas(1) = "@User = totext('" + CStr(Form1.txtUserName.Text) + "')"
           RPT.Formulas(1) = "@User = totext('Admin')"
          RPT.Formulas(2) = "@TglShow = totext('" + CStr(TdTglCall1.Text) + "')"
          RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(TdTglCall2.Text) + "')"
          RPT.ReportFileName = "D:\tnis_report\RptIncomingCall.rpt"
          Call SHOW_PRN
    End Select
Case 1
Unload Me

End Select

End Sub
Public Sub sub101()
Dim strsql As String
Dim tglawal As String
Dim tglakhir As String
Dim M_OBJRSnew As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
bukakoneksiicentra
If Not (TdTglCall1.ValueIsNull) And Not (TdTglCall2.ValueIsNull) Then
   mwhere = Empty
   If Len(mwhere) = 0 Then
        mwhere = mwhere + " and date(calldate) between '"
        mwhere = mwhere + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' and '"
        mwhere = mwhere + Format(TdTglCall2.Value, "yyyy-mm-dd") + "'"
    Else
        mwhere = mwhere + " and date(calldate) between '"
        mwhere = mwhere + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' and '"
        mwhere = mwhere + Format(TdTglCall2.Value, "yyyy-mm-dd") + "'"
    End If
End If


'If cmbtlfullname.Text <> Empty Then
'    If Len(mwhere) = 0 Then
'        mwhere = mwhere + " where fullname ='" + cmbtlfullname.Text + "'"
'    Else
'        mwhere = mwhere + " and  fullname ='" + cmbtlfullname.Text + "'"
'    End If
'End If


'@@ 31-05-2011 Jika yang dipilih per agent
If Opt_Agent.Value Then
    If cmbagentfullname.Text <> Empty Then
        If Len(mwhere) = 0 Then
            mwhere = mwhere + " bill_call.acd_agent_id ='" + Combo1.Text + "'"
        Else
            mwhere = mwhere + " and  bill_call.acd_agent_id ='" + Combo1.Text + "'"
        End If
    End If
 End If

'@@ 31-05-2011 Jika yang dipilih per team
If Opt_Team.Value Then
    If cmbtlfullname.Text <> Empty Then
        If Len(mwhere) = 0 Then
            mwhere = mwhere + " bill_call.acd_agent_id in ("
            mwhere = mwhere + "select acd_agent_id from acd_agent where parent_id in ("
            mwhere = mwhere + "select acd_agent_id from acd_agent where fullname='"
            mwhere = mwhere + Trim(cmbtlfullname.Text) + "')) "
        Else
            mwhere = mwhere + " and bill_call.acd_agent_id in ("
            mwhere = mwhere + "select acd_agent_id from acd_agent where parent_id in ("
            mwhere = mwhere + "select acd_agent_id from acd_agent where fullname='"
            mwhere = mwhere + Trim(cmbtlfullname.Text) + "')) "
        End If
    End If
 End If



strsql = " select * from( "
strsql = strsql + " select * from ("
strsql = strsql + " select tblbaru.acd_agent_id ,tblbaru.userid,tblbaru.nama,tblbaru.team,acd_agent.fullname from ("
strsql = strsql + "select  acd_agent.name as userid,acd_agent.fullname as nama, acd_queue.name as team,acd_agent.acd_agent_id  from acd_agent_group,acd_agent,acd_queue where acd_agent_group.acd_agent_id=acd_agent.acd_agent_id"
strsql = strsql + " and  acd_agent_group.acd_queue_id=acd_queue.acd_queue_id )  as tblbaru,acd_agent where tblbaru.team=acd_agent.name ) tbluser,bill_call"
strsql = strsql + " where tbluser.acd_agent_id= bill_call.acd_agent_id "
strsql = strsql + mwhere + " ) usertbl"



Set M_OBJRSnew = New ADODB.Recordset
M_OBJRSnew.CursorLocation = adUseClient
M_OBJRSnew.Open strsql, connIcentra, adOpenDynamic, adLockOptimistic
 tutupkoneksiicentra
 M_RPTCONN.Execute ("delete from tblcall")
Set m_objrs1 = New ADODB.Recordset
m_objrs1.CursorLocation = adUseClient
strsql = "select * from tblcall "
m_objrs1.Open strsql, M_RPTCONN, adOpenDynamic, adLockOptimistic

PB.Max = M_OBJRSnew.RecordCount + 1
While Not M_OBJRSnew.EOF
   PB.Value = M_OBJRSnew.Bookmark
DoEvents
   m_objrs1.AddNew
   m_objrs1("agent") = IIf(IsNull(M_OBJRSnew("USERID")), "", M_OBJRSnew("USERID"))
   m_objrs1("NAMAAGENT") = IIf(IsNull(M_OBJRSnew("NAMA")), "", M_OBJRSnew("NAMA"))
   m_objrs1("calldate") = IIf(IsNull(M_OBJRSnew("calldate")), Null, M_OBJRSnew("calldate"))
   m_objrs1("stoptime") = IIf(IsNull(M_OBJRSnew("stoptime")), Null, M_OBJRSnew("stoptime"))
   m_objrs1("dst") = IIf(IsNull(M_OBJRSnew("dst")), "", M_OBJRSnew("dst"))
   m_objrs1("dur") = IIf(IsNull(M_OBJRSnew("duration")), "", M_OBJRSnew("duration"))
  m_objrs1("TEAM") = IIf(IsNull(M_OBJRSnew("FULLNAME")), "", M_OBJRSnew("FULLNAME"))
  ' m_objrs1("caledcity") = IIf(IsNull(m_objrsnew("calledcity")), "", m_objrsnew("calledcity"))
   m_objrs1.Update
   M_OBJRSnew.MoveNext
Wend
Set M_OBJRSnew = Nothing
End Sub



Private Sub Form_Load()
Set M_RPTCONN = New ADODB.Connection
M_RPTCONN.Open "Provider=MSDASQL.1;Persist Security Info=False;User ID=Admin;Data Source=TINS_REPORT"
    
    isicombo
End Sub
Public Sub isicombo()
Dim M_OBJRSnew As New ADODB.Recordset
bukakoneksiicentra
strsql = "select * from ("
strsql = strsql + " select tblbaru.acd_agent_id ,tblbaru.userid,tblbaru.nama,tblbaru.team,acd_agent.fullname from ("
strsql = strsql + " select  acd_agent.name as userid,acd_agent.fullname as nama, acd_queue.name as team,acd_agent.acd_agent_id  from acd_agent_group,acd_agent,acd_queue where acd_agent_group.acd_agent_id=acd_agent.acd_agent_id"
strsql = strsql + " and  acd_agent_group.acd_queue_id=acd_queue.acd_queue_id )  as tblbaru,acd_agent where tblbaru.team=acd_agent.name ) tbluser ORDER BY userid asc"

    Set M_OBJRSnew = New ADODB.Recordset
    M_OBJRSnew.CursorLocation = adUseClient
    M_OBJRSnew.Open strsql, connIcentra, adOpenDynamic, adLockOptimistic
   ' cmbAgent.CLEAR
    While Not M_OBJRSnew.EOF
       ' cmbagentfullname.AddItem CStr(M_OBJRSnew("nama"))
       cmbagentfullname.AddItem CStr(M_OBJRSnew("userid"))
        M_OBJRSnew.MoveNext
    Wend
    Set M_OBJRSnew = Nothing
    
    strsql = " select distinct(fullname)  from ("
    strsql = strsql + " select tblbaru.acd_agent_id ,tblbaru.userid,tblbaru.nama,tblbaru.team,acd_agent.fullname from ("
    strsql = strsql + " select  acd_agent.name as userid,acd_agent.fullname as nama, acd_queue.name as team,acd_agent.acd_agent_id  from acd_agent_group,acd_agent,acd_queue where acd_agent_group.acd_agent_id=acd_agent.acd_agent_id"
    strsql = strsql + " and  acd_agent_group.acd_queue_id=acd_queue.acd_queue_id )  as tblbaru,acd_agent where tblbaru.team=acd_agent.name ) tbluser where fullname <>'Team Leader 1' ORDER BY FULLNAME"
    
    Set M_OBJRSnew = New ADODB.Recordset
    M_OBJRSnew.CursorLocation = adUseClient
    M_OBJRSnew.Open strsql, connIcentra, adOpenDynamic, adLockOptimistic
'    cmbAgent.CLEAR
    While Not M_OBJRSnew.EOF
        cmbtlfullname.AddItem CStr(M_OBJRSnew("fullname"))
        M_OBJRSnew.MoveNext
    Wend
    Set M_OBJRSnew = Nothing
    
    
tutupkoneksiicentra
End Sub




Private Sub SHOW_PRN()
'    RPT.Action = 1
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
Public Sub sub102()
Dim strsql As String
Dim tglawal As String
Dim tglakhir As String
Dim M_OBJRSnew As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
bukakoneksiicentra


If Not (TdTglCall1.ValueIsNull) And Not (TdTglCall2.ValueIsNull) Then
mwhere = Empty
   If Len(mwhere) = 0 Then
        mwhere = mwhere + " and date(calldate) between '"
        mwhere = mwhere + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' and '"
        mwhere = mwhere + Format(TdTglCall2.Value, "yyyy-mm-dd") + "'"
    Else
        mwhere = mwhere + " and date(calldate) between '"
        mwhere = mwhere + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' and '"
        mwhere = mwhere + Format(TdTglCall2.Value, "yyyy-mm-dd") + "'"
    End If
End If


'If cmbtlfullname.Text <> Empty Then
'    If Len(mwhere) = 0 Then
'        mwhere = mwhere + " where fullname ='" + cmbtlfullname.Text + "'"
'    Else
'        mwhere = mwhere + " and  fullname ='" + cmbtlfullname.Text + "'"
'    End If
'End If

'
'If cmbagentfullname.Text <> Empty Then
'    If Len(mwhere) = 0 Then
'        mwhere = mwhere + " where bill_call.acd_agent_id ='" + Combo1.Text + "'"
'    Else
'        mwhere = mwhere + " and  bill_call.acd_agent_id ='" + Combo1.Text + "'"
'    End If
'End If


'@@ 31-05-2011 Jika yang dipilih per agent
If Opt_Agent.Value Then
    If cmbagentfullname.Text <> Empty Then
        If Len(mwhere) = 0 Then
            mwhere = mwhere + " bill_call.acd_agent_id ='" + Combo1.Text + "'"
        Else
            mwhere = mwhere + " and  bill_call.acd_agent_id ='" + Combo1.Text + "'"
        End If
    End If
 End If

'@@ 31-05-2011 Jika yang dipilih per team
If Opt_Team.Value Then
    If cmbtlfullname.Text <> Empty Then
        If Len(mwhere) = 0 Then
            mwhere = mwhere + " bill_call.acd_agent_id in ("
            mwhere = mwhere + "select acd_agent_id from acd_agent where parent_id in ("
            mwhere = mwhere + "select acd_agent_id from acd_agent where fullname='"
            mwhere = mwhere + Trim(cmbtlfullname.Text) + "')) "
        Else
            mwhere = mwhere + " and bill_call.acd_agent_id in ("
            mwhere = mwhere + "select acd_agent_id from acd_agent where parent_id in ("
            mwhere = mwhere + "select acd_agent_id from acd_agent where fullname='"
            mwhere = mwhere + Trim(cmbtlfullname.Text) + "')) "
        End If
    End If
 End If


strsql = " select USERID,NAMA ,COUNT(DST) AS attemp,DST ,SUM(DURation) AS duration from("
strsql = strsql + " select * from ("
strsql = strsql + " select tblbaru.acd_agent_id ,tblbaru.userid,tblbaru.nama,tblbaru.team,acd_agent.fullname from ("
strsql = strsql + " select  acd_agent.name as userid,acd_agent.fullname as nama, acd_queue.name as team,acd_agent.acd_agent_id  from acd_agent_group,acd_agent,acd_queue where acd_agent_group.acd_agent_id=acd_agent.acd_agent_id"
strsql = strsql + " and  acd_agent_group.acd_queue_id=acd_queue.acd_queue_id )  as tblbaru,acd_agent where tblbaru.team=acd_agent.name ) tbluser,bill_call"
strsql = strsql + " where tbluser.acd_agent_id= bill_call.acd_agent_id"
strsql = strsql + mwhere + ") usertbl  GROUP BY  USERID,NAMA,dst "

Set M_OBJRSnew = New ADODB.Recordset
M_OBJRSnew.CursorLocation = adUseClient
M_OBJRSnew.Open strsql, connIcentra, adOpenDynamic, adLockOptimistic
 tutupkoneksiicentra
 M_RPTCONN.Execute ("delete from tblcall")
Set m_objrs1 = New ADODB.Recordset
m_objrs1.CursorLocation = adUseClient
strsql = "select * from tblcall"
m_objrs1.Open strsql, M_RPTCONN, adOpenDynamic, adLockOptimistic

PB.Max = M_OBJRSnew.RecordCount + 1
While Not M_OBJRSnew.EOF
   PB.Value = M_OBJRSnew.Bookmark
DoEvents
   m_objrs1.AddNew
   m_objrs1("agent") = IIf(IsNull(M_OBJRSnew("USERID")), "", M_OBJRSnew("USERID"))
   m_objrs1("NAMAAGENT") = IIf(IsNull(M_OBJRSnew("NAMA")), "", M_OBJRSnew("NAMA"))
   m_objrs1("dst") = IIf(IsNull(M_OBJRSnew("dst")), "", M_OBJRSnew("dst"))
   m_objrs1("dur") = IIf(IsNull(M_OBJRSnew("duration")), "", M_OBJRSnew("duration"))
   m_objrs1("attemp") = IIf(IsNull(M_OBJRSnew("attemp")), Null, M_OBJRSnew("attemp"))
 
   m_objrs1.Update
   M_OBJRSnew.MoveNext
Wend
Set M_OBJRSnew = Nothing
End Sub
Public Sub sub103()
Dim strsql As String
Dim tglawal As String
Dim tglakhir As String
Dim M_OBJRSnew As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
bukakoneksiicentra


If Not (TdTglCall1.ValueIsNull) And Not (TdTglCall2.ValueIsNull) Then
   mwhere = Empty
   If Len(mwhere) = 0 Then
        mwhere = mwhere + " and date(calldate) between '"
        mwhere = mwhere + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' and '"
        mwhere = mwhere + Format(TdTglCall2.Value, "yyyy-mm-dd") + "'"
    Else
        mwhere = mwhere + " and date(calldate) between '"
        mwhere = mwhere + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' and '"
        mwhere = mwhere + Format(TdTglCall2.Value, "yyyy-mm-dd") + "'"
    End If
End If

'
'If cmbtlfullname.Text <> Empty Then
'    If Len(mwhere) = 0 Then
'        mwhere = mwhere + " where fullname ='" + cmbtlfullname.Text + "'"
'    Else
'        mwhere = mwhere + " and  fullname ='" + cmbtlfullname.Text + "'"
'    End If
'End If
'
'
'If cmbagentfullname.Text <> Empty Then
'    If Len(mwhere) = 0 Then
'        mwhere = mwhere + " where bill_call.acd_agent_id ='" + Combo1.Text + "'"
'    Else
'        mwhere = mwhere + " and  bill_call.acd_agent_id ='" + Combo1.Text + "'"
'    End If
'End If


'@@ 31-05-2011 Jika yang dipilih per agent
If Opt_Agent.Value Then
    If cmbagentfullname.Text <> Empty Then
        If Len(mwhere) = 0 Then
            mwhere = mwhere + " bill_call.acd_agent_id ='" + Combo1.Text + "'"
        Else
            mwhere = mwhere + " and  bill_call.acd_agent_id ='" + Combo1.Text + "'"
        End If
    End If
 End If

'@@ 31-05-2011 Jika yang dipilih per team
If Opt_Team.Value Then
    If cmbtlfullname.Text <> Empty Then
        If Len(mwhere) = 0 Then
            mwhere = mwhere + " bill_call.acd_agent_id in ("
            mwhere = mwhere + "select acd_agent_id from acd_agent where parent_id in ("
            mwhere = mwhere + "select acd_agent_id from acd_agent where fullname='"
            mwhere = mwhere + Trim(cmbtlfullname.Text) + "')) "
        Else
            mwhere = mwhere + " and bill_call.acd_agent_id in ("
            mwhere = mwhere + "select acd_agent_id from acd_agent where parent_id in ("
            mwhere = mwhere + "select acd_agent_id from acd_agent where fullname='"
            mwhere = mwhere + Trim(cmbtlfullname.Text) + "')) "
        End If
    End If
 End If


strsql = " select userid,NAMA ,fullname,SUM(DURation) AS duration,count(dst) as jmllead from("
strsql = strsql + " select * from ("
strsql = strsql + " select tblbaru.acd_agent_id ,tblbaru.userid,tblbaru.nama,tblbaru.team,acd_agent.fullname from ("
strsql = strsql + " select  acd_agent.name as userid,acd_agent.fullname as nama, acd_queue.name as team,acd_agent.acd_agent_id  from acd_agent_group,acd_agent,acd_queue where acd_agent_group.acd_agent_id=acd_agent.acd_agent_id"
strsql = strsql + " and  acd_agent_group.acd_queue_id=acd_queue.acd_queue_id )  as tblbaru,acd_agent where tblbaru.team=acd_agent.name ) tbluser,bill_call"
strsql = strsql + " where tbluser.acd_agent_id= bill_call.acd_agent_id"
strsql = strsql + mwhere + " ) usertbl group by  userid,nama,fullname order by nama,fullname "
 
Set M_OBJRSnew = New ADODB.Recordset
M_OBJRSnew.CursorLocation = adUseClient
M_OBJRSnew.Open strsql, connIcentra, adOpenDynamic, adLockOptimistic
 tutupkoneksiicentra
 M_RPTCONN.Execute ("delete from tblcall")
Set m_objrs1 = New ADODB.Recordset
m_objrs1.CursorLocation = adUseClient
strsql = "select * from tblcall"
m_objrs1.Open strsql, M_RPTCONN, adOpenDynamic, adLockOptimistic

PB.Max = M_OBJRSnew.RecordCount + 1
While Not M_OBJRSnew.EOF
   PB.Value = M_OBJRSnew.Bookmark
DoEvents
   m_objrs1.AddNew
   m_objrs1("agent") = IIf(IsNull(M_OBJRSnew("USERID")), "", M_OBJRSnew("USERID"))
   m_objrs1("NAMAAGENT") = IIf(IsNull(M_OBJRSnew("NAMA")), "", M_OBJRSnew("NAMA"))
   m_objrs1("team") = IIf(IsNull(M_OBJRSnew("fullname")), "", M_OBJRSnew("fullname"))
   m_objrs1("dur") = IIf(IsNull(M_OBJRSnew("duration")), "", M_OBJRSnew("duration"))
    m_objrs1("attemp") = IIf(IsNull(M_OBJRSnew("jmllead")), "", M_OBJRSnew("jmllead"))
'  m_objrs1("jml") = 1
   m_objrs1.Update
   M_OBJRSnew.MoveNext
Wend
Set M_OBJRSnew = Nothing

End Sub

Sub WaitSecs(Seconds As Single)
 Dim a As Long
 Seconds = Seconds + Timer
 While Seconds > Timer
  a = DoEvents
 Wend
End Sub
Public Sub sub104()
Dim strsql As String
Dim tglawal As String
Dim tglakhir As String
Dim M_OBJRSnew As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
bukakoneksiicentra
If Not (TdTglCall1.ValueIsNull) And Not (TdTglCall2.ValueIsNull) Then
    mwhere = Empty
   If Len(mwhere) = 0 Then
        mwhere = mwhere + " and date(calldate) between '"
        mwhere = mwhere + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' and '"
        mwhere = mwhere + Format(TdTglCall2.Value, "yyyy-mm-dd") + "'"
    Else
        mwhere = mwhere + " and date(calldate) between '"
        mwhere = mwhere + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' and '"
        mwhere = mwhere + Format(TdTglCall2.Value, "yyyy-mm-dd") + "'"
    End If
End If


'If cmbtlfullname.Text <> Empty Then
'    If Len(mwhere) = 0 Then
'        mwhere = mwhere + " where fullname ='" + cmbtlfullname.Text + "'"
'    Else
'        mwhere = mwhere + " and  fullname ='" + cmbtlfullname.Text + "'"
'    End If
'End If
'
'
'If cmbagentfullname.Text <> Empty Then
'    If Len(mwhere) = 0 Then
'        mwhere = mwhere + " where bill_call.acd_agent_id ='" + Combo1.Text + "'"
'    Else
'        mwhere = mwhere + " and  bill_call.acd_agent_id ='" + Combo1.Text + "'"
'    End If
'End If



'@@ 31-05-2011 Jika yang dipilih per agent
If Opt_Agent.Value Then
    If cmbagentfullname.Text <> Empty Then
        If Len(mwhere) = 0 Then
            mwhere = mwhere + " bill_call.acd_agent_id ='" + Combo1.Text + "'"
        Else
            mwhere = mwhere + " and  bill_call.acd_agent_id ='" + Combo1.Text + "'"
        End If
    End If
 End If

'@@ 31-05-2011 Jika yang dipilih per team
If Opt_Team.Value Then
    If cmbtlfullname.Text <> Empty Then
        If Len(mwhere) = 0 Then
            mwhere = mwhere + " bill_call.acd_agent_id in ("
            mwhere = mwhere + "select acd_agent_id from acd_agent where parent_id in ("
            mwhere = mwhere + "select acd_agent_id from acd_agent where fullname='"
            mwhere = mwhere + Trim(cmbtlfullname.Text) + "')) "
        Else
            mwhere = mwhere + " and bill_call.acd_agent_id in ("
            mwhere = mwhere + "select acd_agent_id from acd_agent where parent_id in ("
            mwhere = mwhere + "select acd_agent_id from acd_agent where fullname='"
            mwhere = mwhere + Trim(cmbtlfullname.Text) + "')) "
        End If
    End If
 End If



strsql = " select * from( "
strsql = strsql + " select * from ("
strsql = strsql + " select tblbaru.acd_agent_id ,tblbaru.userid,tblbaru.nama,tblbaru.team,acd_agent.fullname from ("
strsql = strsql + "select  acd_agent.name as userid,acd_agent.fullname as nama, acd_queue.name as team,acd_agent.acd_agent_id  from acd_agent_group,acd_agent,acd_queue where acd_agent_group.acd_agent_id=acd_agent.acd_agent_id"
strsql = strsql + " and  acd_agent_group.acd_queue_id=acd_queue.acd_queue_id )  as tblbaru,acd_agent where tblbaru.team=acd_agent.name ) tbluser,bill_call"
strsql = strsql + " where tbluser.acd_agent_id= bill_call.acd_agent_id"
strsql = strsql + mwhere + " ) usertbl "



Set M_OBJRSnew = New ADODB.Recordset
M_OBJRSnew.CursorLocation = adUseClient
M_OBJRSnew.Open strsql, connIcentra, adOpenDynamic, adLockOptimistic
 tutupkoneksiicentra
 M_RPTCONN.Execute ("delete from tblcall")
Set m_objrs1 = New ADODB.Recordset
m_objrs1.CursorLocation = adUseClient
strsql = "select * from tblcall "
m_objrs1.Open strsql, M_RPTCONN, adOpenDynamic, adLockOptimistic

PB.Max = M_OBJRSnew.RecordCount + 1
While Not M_OBJRSnew.EOF
   PB.Value = M_OBJRSnew.Bookmark
DoEvents
   m_objrs1.AddNew
   m_objrs1("agent") = IIf(IsNull(M_OBJRSnew("USERID")), "", M_OBJRSnew("USERID"))
   m_objrs1("NAMAAGENT") = IIf(IsNull(M_OBJRSnew("NAMA")), "", M_OBJRSnew("NAMA"))
   m_objrs1("calldate") = IIf(IsNull(M_OBJRSnew("calldate")), Null, M_OBJRSnew("calldate"))
   m_objrs1("stoptime") = IIf(IsNull(M_OBJRSnew("stoptime")), Null, M_OBJRSnew("stoptime"))
   m_objrs1("dst") = IIf(IsNull(M_OBJRSnew("dst")), "", M_OBJRSnew("dst"))
   m_objrs1("dur") = IIf(IsNull(M_OBJRSnew("duration")), "", M_OBJRSnew("duration"))
  m_objrs1("TEAM") = IIf(IsNull(M_OBJRSnew("FULLNAME")), "", M_OBJRSnew("FULLNAME"))
  ' m_objrs1("caledcity") = IIf(IsNull(m_objrsnew("calledcity")), "", m_objrsnew("calledcity"))
   m_objrs1.Update
   M_OBJRSnew.MoveNext
Wend
Set M_OBJRSnew = Nothing
End Sub



Public Sub sub105()
Dim strsql As String
Dim tglawal As String
Dim tglakhir As String
Dim M_OBJRSnew As New ADODB.Recordset
Dim m_objrs1 As New ADODB.Recordset
bukakoneksiicentra
If Not (TdTglCall1.ValueIsNull) And Not (TdTglCall2.ValueIsNull) Then
   If Len(mwhere) = 0 Then
        mwhere = mwhere + " where date(calldate) between '"
        mwhere = mwhere + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' and '"
        mwhere = mwhere + Format(TdTglCall2.Value, "yyyy-mm-dd") + "'"
    Else
        mwhere = mwhere + " and date(calldate) between '"
        mwhere = mwhere + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' and '"
        mwhere = mwhere + Format(TdTglCall2.Value, "yyyy-mm-dd") + "'"
    End If
End If


If cmbtlfullname.Text <> Empty Then
    If Len(mwhere) = 0 Then
        mwhere = mwhere + " where fullname ='" + cmbtlfullname.Text + "'"
    Else
        mwhere = mwhere + " and  fullname ='" + cmbtlfullname.Text + "'"
    End If
End If


If cmbagentfullname.Text <> Empty Then
    If Len(mwhere) = 0 Then
        mwhere = mwhere + " where nama ='" + cmbagentfullname.Text + "'"
    Else
        mwhere = mwhere + " and  nama ='" + cmbagentfullname.Text + "'"
    End If
End If





strsql = " select * from( "
strsql = strsql + " select * from ("
strsql = strsql + " select tblbaru.acd_agent_id ,tblbaru.userid,tblbaru.nama,tblbaru.team,acd_agent.fullname from ("
strsql = strsql + "select  acd_agent.name as userid,acd_agent.fullname as nama, acd_queue.name as team,acd_agent.acd_agent_id  from acd_agent_group,acd_agent,acd_queue where acd_agent_group.acd_agent_id=acd_agent.acd_agent_id"
strsql = strsql + " and  acd_agent_group.acd_queue_id=acd_queue.acd_queue_id )  as tblbaru,acd_agent where tblbaru.team=acd_agent.name ) tbluser,bill_call"
strsql = strsql + " where tbluser.acd_agent_id= bill_call.acd_agent_id) usertbl "



Set M_OBJRSnew = New ADODB.Recordset
M_OBJRSnew.CursorLocation = adUseClient
M_OBJRSnew.Open strsql + mwhere, connIcentra, adOpenDynamic, adLockOptimistic
 tutupkoneksiicentra
 M_RPTCONN.Execute ("delete from tblcall")
Set m_objrs1 = New ADODB.Recordset
m_objrs1.CursorLocation = adUseClient
strsql = "select * from tblcall "
m_objrs1.Open strsql, M_RPTCONN, adOpenDynamic, adLockOptimistic

PB.Max = M_OBJRSnew.RecordCount + 1
While Not M_OBJRSnew.EOF
   PB.Value = M_OBJRSnew.Bookmark
DoEvents
   m_objrs1.AddNew
   m_objrs1("agent") = IIf(IsNull(M_OBJRSnew("USERID")), "", M_OBJRSnew("USERID"))
   m_objrs1("NAMAAGENT") = IIf(IsNull(M_OBJRSnew("NAMA")), "", M_OBJRSnew("NAMA"))
   m_objrs1("calldate") = IIf(IsNull(M_OBJRSnew("calldate")), Null, M_OBJRSnew("calldate"))
   m_objrs1("stoptime") = IIf(IsNull(M_OBJRSnew("stoptime")), Null, M_OBJRSnew("stoptime"))
   m_objrs1("dst") = IIf(IsNull(M_OBJRSnew("dst")), "", M_OBJRSnew("dst"))
   m_objrs1("dur") = IIf(IsNull(M_OBJRSnew("duration")), "", M_OBJRSnew("duration"))
  m_objrs1("TEAM") = IIf(IsNull(M_OBJRSnew("FULLNAME")), "", M_OBJRSnew("FULLNAME"))
  ' m_objrs1("caledcity") = IIf(IsNull(m_objrsnew("calledcity")), "", m_objrsnew("calledcity"))
   m_objrs1.Update
   M_OBJRSnew.MoveNext
Wend
Set M_OBJRSnew = Nothing
End Sub


'Private Sub Opt_Agent_Click()
'    If Opt_Agent.Value Then
'        cmbagentfullname.Enabled = True
'        cmbtlfullname.Enabled = False
'    Else
'        cmbagentfullname.Enabled = False
'    End If
'End Sub
'
'
'
'Private Sub Opt_Team_Click()
'    If Opt_Team.Value Then
'        cmbagentfullname.Enabled = True
'        cmbtlfullname.Enabled = False
'    Else
'        cmbagentfullname.Enabled = False
'    End If
'End Sub

Private Sub TdTglCall1_Change()
    TdTglCall2.Value = TdTglCall1.Value
End Sub
