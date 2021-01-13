VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form_rpt_last_status 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report Last Statuscall"
   ClientHeight    =   11070
   ClientLeft      =   300
   ClientTop       =   570
   ClientWidth     =   15735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11070
   ScaleWidth      =   15735
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Criteria Report"
      Height          =   2595
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   15705
      Begin VB.TextBox txtcustid 
         Height          =   315
         Left            =   1500
         TabIndex        =   26
         Top             =   1815
         Width           =   2115
      End
      Begin VB.CheckBox Check_all1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Check All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5640
         TabIndex        =   23
         Top             =   2220
         Width           =   1455
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   3165
         TabIndex        =   20
         Top             =   1425
         Width           =   2385
      End
      Begin VB.ComboBox cboagentname 
         Height          =   315
         Left            =   1500
         TabIndex        =   19
         Top             =   1425
         Width           =   1635
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E87211&
         Caption         =   "Show Phone Number"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   13995
         TabIndex        =   15
         Top             =   -1290
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.CommandButton SSCommand2 
         BackColor       =   &H00F1E5DB&
         Cancel          =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9480
         Picture         =   "Form_rpt_last_status.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   915
         Width           =   1605
      End
      Begin VB.CommandButton cmdCari 
         Height          =   360
         Left            =   9480
         Picture         =   "Form_rpt_last_status.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   480
         Width           =   1605
      End
      Begin VB.CommandButton SSCommand1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Export to Excel"
         Height          =   810
         Left            =   9480
         Picture         =   "Form_rpt_last_status.frx":0C34
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1350
         Width           =   1590
      End
      Begin VB.ComboBox cbocampaign 
         Height          =   315
         Left            =   1515
         TabIndex        =   4
         Top             =   645
         Width           =   4035
      End
      Begin VB.ComboBox cbostatuscall 
         Height          =   315
         Left            =   1515
         TabIndex        =   3
         Top             =   1035
         Width           =   4035
      End
      Begin VB.TextBox TxtPath 
         Enabled         =   0   'False
         Height          =   285
         Left            =   13275
         TabIndex        =   2
         Top             =   -930
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSComDlg.CommonDialog Cd_save 
         Left            =   13755
         Top             =   -570
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "*.xls"
      End
      Begin TDBDate6Ctl.TDBDate tgl_call 
         Height          =   315
         Index           =   0
         Left            =   1515
         TabIndex        =   16
         Top             =   240
         Width           =   1365
         _Version        =   65536
         _ExtentX        =   2408
         _ExtentY        =   556
         Calendar        =   "Form_rpt_last_status.frx":139A
         Caption         =   "Form_rpt_last_status.frx":14B2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Form_rpt_last_status.frx":151E
         Keys            =   "Form_rpt_last_status.frx":153C
         Spin            =   "Form_rpt_last_status.frx":159A
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
      Begin TDBDate6Ctl.TDBDate tgl_call 
         Height          =   315
         Index           =   1
         Left            =   3480
         TabIndex        =   17
         Top             =   240
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   556
         Calendar        =   "Form_rpt_last_status.frx":15C2
         Caption         =   "Form_rpt_last_status.frx":16DA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Form_rpt_last_status.frx":1746
         Keys            =   "Form_rpt_last_status.frx":1764
         Spin            =   "Form_rpt_last_status.frx":17C2
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
         Left            =   13275
         Top             =   -570
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSComctlLib.ListView List_Supervisor 
         Height          =   1695
         Left            =   5640
         TabIndex        =   24
         Top             =   480
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Cust Number"
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
         Height          =   315
         Index           =   0
         Left            =   105
         TabIndex        =   25
         Top             =   1830
         Width           =   1425
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Supervisor"
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
         Height          =   315
         Index           =   5
         Left            =   5640
         TabIndex        =   22
         Top             =   120
         Width           =   1425
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Telesales "
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
         Height          =   315
         Index           =   6
         Left            =   105
         TabIndex        =   21
         Top             =   1425
         Width           =   1425
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2970
         TabIndex        =   18
         Top             =   270
         Width           =   825
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Tanggal Call"
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
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Campaign"
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
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   6
         Top             =   645
         Width           =   1425
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Status Call "
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
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1040
         Width           =   1425
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   6150
      Left            =   0
      TabIndex        =   11
      Top             =   3615
      Width           =   17730
      Begin VB.TextBox txtlead 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   14760
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   225
         Width           =   915
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5280
         Left            =   45
         TabIndex        =   12
         Top             =   720
         Width           =   15660
         _ExtentX        =   27623
         _ExtentY        =   9313
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
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
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         X1              =   45
         X2              =   15690
         Y1              =   630
         Y2              =   630
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Jml Lead"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   13830
         TabIndex        =   13
         Top             =   270
         Width           =   915
      End
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   0
      Left            =   30
      Picture         =   "Form_rpt_last_status.frx":17EA
      Stretch         =   -1  'True
      Top             =   240
      Width           =   420
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Report Last Statuscall"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   300
      Width           =   3585
   End
   Begin VB.Image Image2 
      Height          =   825
      Left            =   -1980
      Picture         =   "Form_rpt_last_status.frx":22F4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   17700
   End
End
Attribute VB_Name = "Form_rpt_last_status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sGetSPV As String
Private Sub cboagentname_Click()
    cboagentname_LostFocus
End Sub

Private Sub cboagentname_DropDown()
'LOAD_AGENT
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient

strsql = "select * from usertbl where aktif='1'  and   kdlevel='1'"

M_objrs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
cboagentname.CLEAR
    While Not M_objrs.EOF
      cboagentname.AddItem IIf(IsNull(M_objrs!USERID), "", M_objrs!USERID)
        M_objrs.MoveNext
    Wend
 Set M_objrs = Nothing


End Sub

Private Sub cboagentname_LostFocus()
Dim mobjr As New ADODB.Recordset
Set mobjr = New ADODB.Recordset
   mobjr.CursorLocation = adUseClient
   
strsql = "select * from usertbl where userid='" + cboagentname.Text + "'"
mobjr.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
If Not mobjr.EOF Then
Combo3.Text = IIf(IsNull(mobjr!AGENT), "", mobjr!AGENT)
End If
Set mobjr = Nothing

End Sub

Private Sub cbocampaign_DropDown()
sStrsql = "select * from datasourcetbl where   status ='1' order by  tglentry,  kodeds "
Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    cbocampaign.CLEAR
    While Not M_objrs.EOF
        cbocampaign.AddItem IIf(IsNull(M_objrs!KODEDS), "", M_objrs!KODEDS)
        
        M_objrs.MoveNext
    Wend
Set M_objrs = Nothing
End Sub
Public Sub load_spv()
    If MDIForm1.txtlevel.Text = "Agent" Then
        sStrsql = " select userid , agent  from usertbl where  userid in  (select distinct spvcode  from usertbl where  spvcode= '" + MDIForm1.TxtUsername.Text + "') and aktif='1'"
    ElseIf MDIForm1.txtlevel.Text = "Supervisor" Then
        sStrsql = " select userid , agent  from usertbl where  userid = '" + MDIForm1.TxtUsername.Text + "' and  aktif ='1'"
    Else
        sStrsql = "select userid , agent  from  usertbl  where  aktif ='1' and  kdlevel ='2'"
    End If
    
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        CBOTEAMNAME.CLEAR
        While Not M_objrs.EOF
                CBOTEAMNAME.AddItem IIf(IsNull(M_objrs!USERID), "", M_objrs!USERID) & "!" & IIf(IsNull(M_objrs!AGENT), "", M_objrs!AGENT)
                M_objrs.MoveNext
        Wend
        
    Set M_objrs = Nothing
End Sub

Private Sub cbostatuscall_DropDown()
    load_statuscall
End Sub

Private Sub CBOTEAMNAME_Click()
CBOTEAMNAME_LostFocus
cboagentname.CLEAR
Combo3.CLEAR
End Sub
Private Sub CBOTEAMNAME_DropDown()
' Dim clsspv As New clsTbluser
'    Set clsspv = New clsTbluser
'    Set M_objrs = clsspv.FindRecordUser("", "", "2", "1", "", "")
'    CBOTEAMNAME.CLEAR
'    While Not M_objrs.EOF
'        CBOTEAMNAME.AddItem IIf(IsNull(M_objrs!tbluser_userid), "", M_objrs!tbluser_userid)
'        M_objrs.MoveNext
'    Wend
' Set clsspv = Nothing
' Set M_objrs = Nothing


End Sub

Public Sub load_statuscall()
    sStrsql = " select tblstatuscall_kdstscall,  tblstatuscall_keterangan  from tblstatuscall  "
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        cbostatuscall.CLEAR
        While Not M_objrs.EOF
                cbostatuscall.AddItem IIf(IsNull(M_objrs!tblstatuscall_keterangan), "", M_objrs!tblstatuscall_keterangan)
                M_objrs.MoveNext
        Wend
    Set M_objrs = Nothing
End Sub

Private Sub CBOTEAMNAME_LostFocus()
Dim mobjr As New ADODB.Recordset
Set mobjr = New ADODB.Recordset
   mobjr.CursorLocation = adUseClient
   
strsql = "select * from tbluser where tbluser_userid='" + CBOTEAMNAME.Text + "'"
mobjr.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
If Not mobjr.EOF Then
Combo2.Text = IIf(IsNull(mobjr!tbluser_name), "", mobjr!tbluser_name)
End If
Set mobjr = Nothing

End Sub

Private Sub CmdCari_Click()
    Dim MOBJ As ADODB.Recordset
    Dim jml As Double
    Dim getSpvcode As String
    Dim getSpv_name As String
    Dim getUserid As String
    Dim getCampaign_code As String
    Dim getCampaign_name As String
    Dim strsql1 As String
    If cboagentname.Text <> "" Then
        intvrl = InStr(1, cboagentname.Text, "!", vbTextCompare)
        If intvrl <> 0 Then
            ArrayString = Split(cboagentname.Text, "!", 2, vbTextCompare)
            getUserid = ArrayString(0)
            getUser_name = ArrayString(1)
        End If
    End If


'     strsql = " SELECT date(a.tgl) as ""Call Date"",to_char(a.tgl,'HH24:MI:SS') AS ""Call Time"",a.custid as ""Customer Number"",b.name as ""Customer Name"",b.curbal as ""Outstanding"",c.promisepay as ""PTP Volume"",b.region as ""Kota"",a.lastcall as ""Last Status"",a.statuscall as ""Group Call"",a.hst as ""Status Account"",recsource as ""Campaign Name"",a.agent as ""Agent Code"",b.nama_agent as ""Agent Name"""
'     strsql = strsql + " FROM (select y.* from (SELECT max(id) as last_id,date(tgl) as tgl_akhir,custid FROM mgm_hst where coalesce(lastcall,'')<>'' group by 2,3) x,mgm_hst y where x.last_id=y.id) a "
'     strsql = strsql + " LEFT JOIN mgm b ON (a.custid=b.custid) "
'     strsql = strsql + " LEFT JOIN tblnegoptp c ON (a.custid=c.custid) "
     
     strsql = " SELECT date(a.tglcall) as ""Call Date"",to_char(a.tglcall,'HH24:MI:SS') AS ""Call Time"",a.custid as ""Customer Number"",a.name as ""Customer Name"",a.curbal as ""Outstanding"",a.amountnew as ""PTP Volume"",a.region as ""Kota"",a.statuscall as ""Last Status"",a.stscallcust as ""Group Call"",a.remarks as ""Status Account"",recsource as ""Campaign Name"",a.agent as ""Agent Code"",a.nama_agent as ""Agent Name"""
     strsql = strsql + " FROM mgm a"
    
    mwhere = " WHERE 1=1 "
    
    If Not (tgl_call(0).ValueIsNull) And Not (tgl_call(1).ValueIsNull) Then
        If Len(mwhere) = 0 Then
            mwhere = " where  date(tglcall) between '" + Format(tgl_call(0).Value, "yyyy-mm-dd") + "' "
            mwhere = mwhere + " and '" + Format(tgl_call(1).Value, "yyyy-mm-dd") + "'"
        Else
            mwhere = mwhere + "  and date(tglcall) between '" + Format(tgl_call(0).Value, "yyyy-mm-dd") + "' "
            mwhere = mwhere + " and '" + Format(tgl_call(1).Value, "yyyy-mm-dd") + "'"
        End If
    Else
        'MsgBox "Tanggal Call Harus Diisi", vbInformation, "Informasi"
    End If
    
    If cbocampaign.Text <> Empty Then
        If Len(mwhere) = 0 Then
            mwhere = mwhere + " where recsource like '%" + cbocampaign.Text + "%'"
        Else
            mwhere = mwhere + " and recsource like '%" + cbocampaign.Text + "%'"
        End If
    End If
    
    If txtcustid.Text <> Empty Then
        If Len(mwhere) = 0 Then
            mwhere = mwhere + " where custid = '" + txtcustid.Text + "'"
        Else
            mwhere = mwhere + " and custid = '" + txtcustid.Text + "'"
        End If
    End If
    
    If sGetSPV <> Empty Then
        If Len(mwhere) = 0 Then
            mwhere = mwhere + " where a.agent in (select userid from  usertbl where  spvcode in (" + sGetSPV + "))"
        Else
            mwhere = mwhere + " and  a.agent in (select userid from usertbl  where  spvcode in (" + sGetSPV + "))"
        End If
    End If
    
    If cboagentname.Text <> Empty Then
        If Len(mwhere) = 0 Then
            mwhere = mwhere + " where a.agent ='" + cboagentname.Text + "'"
        Else
            mwhere = mwhere + " and  a.agent ='" + cboagentname.Text + "'"
        End If
    End If

    If cbostatuscall.Text <> Empty Then
        If Len(mwhere) = 0 Then
            mwhere = "where     statuscall ='" + cbostatuscall.Text + "'"
        Else
            mwhere = mwhere + " and      statuscall ='" + cbostatuscall.Text + "'"
        End If
    End If
    
        strsqlJML = " SELECT SUM(AMOUNT) AS ttl FROM (" + strsql + mwhere + " ) AS MGM"

    strsqlJML = " SELECT SUM(AMOUNT) AS ttl FROM (" + strsql + mwhere + " ) AS MGM"

    Set MOBJ = New ADODB.Recordset
    MOBJ.CursorLocation = adUseClient
    MOBJ.Open strsql + mwhere + " ORDER BY a.tglcall", M_OBJCONN, adOpenKeyset, adLockOptimistic
    
    txtlead.Text = MOBJ.RecordCount
    Set DataGrid1.DATASOURCE = MOBJ
    cmdCari.Enabled = True
End Sub
Private Sub Combo3_Click()
Combo3_LostFocus
End Sub

Private Sub Combo3_DropDown()
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient

strsql = "select * from usertbl where aktif='1'  and kdlevel='1'"

M_objrs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
Combo3.CLEAR
    While Not M_objrs.EOF
      Combo3.AddItem IIf(IsNull(M_objrs!AGENT), "", M_objrs!AGENT)
        M_objrs.MoveNext
    Wend
 Set M_objrs = Nothing

End Sub

Private Sub Combo3_LostFocus()
Dim mobjr As New ADODB.Recordset
Set mobjr = New ADODB.Recordset
   mobjr.CursorLocation = adUseClient
   
strsql = "select * from usertbl where agent='" + Combo3.Text + "'"
mobjr.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
If Not mobjr.EOF Then
    cboagentname.Text = IIf(IsNull(mobjr!USERID), "", mobjr!USERID)
End If
Set mobjr = Nothing

End Sub
Private Sub Form_Load()
    List_Supervisor.ColumnHeaders.ADD 1, , "Kode Supervisor", 1000
    List_Supervisor.ColumnHeaders.ADD 2, , "Nama Supervisor", 5000

    Call load_spv1
    If MDIForm1.txtlevel.Text = "Supervisor" Then
       Check_all1.Value = 1
       Check_all1_Click
       Check_all1.Enabled = False
       List_Supervisor.Enabled = False
    Else
       List_Supervisor.Enabled = True
       Check_all1.Enabled = True
    End If
End Sub
Public Sub load_spv1()
    Dim listv As ListItem
    If MDIForm1.txtlevel.Text = "Supervisor" Then
        sStrsql = " select userid , agent  from usertbl where  userid = '" + MDIForm1.TxtUsername.Text + "' and  aktif ='1'"
    Else
        sStrsql = "select userid , agent  from usertbl  where  aktif ='1' and  level_name ='Supervisor'"
    End If
    
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        List_Supervisor.ListItems.CLEAR
        While Not M_objrs.EOF
                Set listv = List_Supervisor.ListItems.ADD(, , IIf(IsNull(M_objrs!USERID), "", M_objrs!USERID))
                listv.SubItems(1) = IIf(IsNull(M_objrs!AGENT), "", M_objrs!AGENT)
                M_objrs.MoveNext
        Wend
    Set M_objrs = Nothing
End Sub

Private Sub Check_all1_Click()
    Dim i As Integer
    i = 0
    If Check_all1.Value = 1 Then
        For i = 1 To List_Supervisor.ListItems.Count
            List_Supervisor.ListItems(i).Checked = True
        Next i
    ElseIf Check_all1.Value = 0 Then
        For i = 1 To List_Supervisor.ListItems.Count
            List_Supervisor.ListItems(i).Checked = False
        Next i
    End If
    Call GetSPVs
End Sub
Private Sub GetSPVs()
    Dim sWhere As String
    sWhere = ""
    sWhere = GETSPV
    If sWhere <> "" Then
        sGetSPV = ""
        sGetSPV = sWhere
        Exit Sub
    End If
End Sub
Public Function GETSPV() As Variant
    Dim row As Double
    row = 1
    strsql = ""
    For i = 1 To List_Supervisor.ListItems.Count
       If List_Supervisor.ListItems(i).Checked = True Then
            If row = 1 Then
                strsql = "'" + List_Supervisor.ListItems(i).Text + "'"
            Else
                strsql = strsql + ",'" + List_Supervisor.ListItems(i).Text + "'"
            End If
            row = row + 1
      End If
    Next i
    GETSPV = strsql
End Function

Private Sub List_Supervisor_Click()
    Call GetSPVs
End Sub

Private Sub SSCommand1_Click()
    export_data
End Sub

Private Sub SSCommand2_Click()
Unload Me
End Sub
Public Sub export_data()
    Dim MOBJ As ADODB.Recordset
    Dim jml As Double
    Dim getSpvcode As String
    Dim getSpv_name As String
    Dim getUserid As String
    Dim getCampaign_code As String
    Dim getCampaign_name As String
    Dim strsql1 As String
    Dim strsql As String
        If cboagentname.Text <> "" Then
        intvrl = InStr(1, cboagentname.Text, "!", vbTextCompare)
        If intvrl <> 0 Then
            ArrayString = Split(cboagentname.Text, "!", 2, vbTextCompare)
            getUserid = ArrayString(0)
            getUser_name = ArrayString(1)
        End If
    End If
    
'     strsql = " SELECT date(a.tgl) as ""Call Date"",to_char(a.tgl,'HH24:MI:SS') AS ""Call Time"",a.custid as ""Customer Number"",b.name as ""Customer Name"",b.curbal as ""Outstanding"",c.promisepay as ""PTP Volume"",b.region as ""Kota"",a.lastcall as ""Last Status"",a.statuscall as ""Group Call"",a.hst as ""Status Account"",recsource as ""Campaign Name"",a.agent as ""Agent Code"",b.nama_agent as ""Agent Name"""
'     strsql = strsql + " FROM (select y.* from (SELECT max(id) as last_id,date(tgl) as tgl_akhir,custid FROM mgm_hst where coalesce(lastcall,'')<>'' group by 2,3) x,mgm_hst y where x.last_id=y.id) a "
'     strsql = strsql + " LEFT JOIN mgm b ON (a.custid=b.custid) "
'     strsql = strsql + " LEFT JOIN tblnegoptp c ON (a.custid=c.custid) "
    strsql = " SELECT date(a.tglcall) as ""Call Date"",to_char(a.tglcall,'HH24:MI:SS') AS ""Call Time"",a.custid as ""Customer Number"",a.name as ""Customer Name"",a.curbal as ""Outstanding"",a.amountnew as ""PTP Volume"",a.region as ""Kota"",a.statuscall as ""Last Status"",a.stscallcust as ""Group Call"",a.remarks as ""Status Account"",recsource as ""Campaign Name"",a.agent as ""Agent Code"",a.nama_agent as ""Agent Name"""
     strsql = strsql + " FROM mgm a"
    
    mwhere = " WHERE 1=1 "
    
    If Not (tgl_call(0).ValueIsNull) And Not (tgl_call(1).ValueIsNull) Then
        If Len(mwhere) = 0 Then
            mwhere = " where  date(tglcall) between '" + Format(tgl_call(0).Value, "yyyy-mm-dd") + "' "
            mwhere = mwhere + " and '" + Format(tgl_call(1).Value, "yyyy-mm-dd") + "'"
        Else
            mwhere = mwhere + "  and date(tglcall) between '" + Format(tgl_call(0).Value, "yyyy-mm-dd") + "' "
            mwhere = mwhere + " and '" + Format(tgl_call(1).Value, "yyyy-mm-dd") + "'"
        End If
    Else
        'MsgBox "Tanggal Call Harus Diisi", vbInformation, "Informasi"
    End If
    
    If cbocampaign.Text <> Empty Then
        If Len(mwhere) = 0 Then
            mwhere = mwhere + " where recsource like '%" + cbocampaign.Text + "%'"
        Else
            mwhere = mwhere + " and recsource like '%" + cbocampaign.Text + "%'"
        End If
    End If
    
    If txtcustid.Text <> Empty Then
        If Len(mwhere) = 0 Then
            mwhere = mwhere + " where custid = '" + txtcustid.Text + "'"
        Else
            mwhere = mwhere + " and custid = '" + txtcustid.Text + "'"
        End If
    End If
    
    If sGetSPV <> Empty Then
        If Len(mwhere) = 0 Then
            mwhere = mwhere + " where a.agent in (select userid from  usertbl where  spvcode in (" + sGetSPV + "))"
        Else
            mwhere = mwhere + " and  a.agent in (select userid from usertbl  where  spvcode in (" + sGetSPV + "))"
        End If
    End If
    
    If cboagentname.Text <> Empty Then
        If Len(mwhere) = 0 Then
            mwhere = mwhere + " where a.agent ='" + cboagentname.Text + "'"
        Else
            mwhere = mwhere + " and  a.agent ='" + cboagentname.Text + "'"
        End If
    End If

    If cbostatuscall.Text <> Empty Then
        If Len(mwhere) = 0 Then
            mwhere = "where     statuscall ='" + cbostatuscall.Text + "'"
        Else
            mwhere = mwhere + " and      statuscall ='" + cbostatuscall.Text + "'"
        End If
    End If
    
    
        strsqlJML = " SELECT SUM(AMOUNT) AS ttl FROM (" + strsql + mwhere + " ) AS MGM"
        isi_data (strsql + mwhere + " ORDER BY a.tglcall")
    
End Sub
Private Sub isi_data(strsql As String)
On Error GoTo Salah
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim ListItem As ListItem
    Dim cmdsql_update As String
    Dim objExcel        As Excel.Application
    Dim objBook         As Excel.Workbook
    Dim objSheet        As Excel.Worksheet
    Dim i As Integer
    Dim m_msgbox As String
    
    i = 1
    

    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
  
   

    'Jika data tidak ada, maka keluar dari fungsi ini!
'    If M_OBJRS.RecordCount = 0 Then
'        MsgBox "Data  tidak ada!", vbOKOnly + vbInformation, "Informasi"
'        Exit Sub
'    End If

   
Form_Save:
    Cd_save.ShowSave
    TxtPath.Text = Cd_save.FileName
    
    'Cek apakah user menekan tombol cancel pada dialog save
    If TxtPath.Text = Empty Then
        'Tanyakan ke user.. apakah benar2 akan membatalkan proses download???
        m_msgbox = MsgBox("Anda ingin Download dibatalkan?", vbYesNo + vbQuestion, "Konfirmasi")
        'Jika user benar-benar akan membatalkan proses download, keluar dari fungsi ini!
        If m_msgbox = vbYes Then
              MsgBox "Download dibatalkan!", vbOKOnly + vbInformation, "Informasi"
            Exit Sub
        End If
        If m_msgbox = vbNo Then '-> jika user tidak membatalkan proses download
          GoTo Form_Save        '-> maka goto form_save
        End If
    End If
    
    
    
        
    
'    lblstatus.Caption = "Status download: Mengisi field... silahkan tunggu!"
    
'    If WriteRecordsetToCSv(M_OBJRS, TXTPATH.Text + ".CSV", ",") Then
'        MsgBox " export berhasil"
'    End If
    

 Set objExcel = CreateObject("Excel.Application")
    Set objBook = objExcel.Workbooks.ADD
    Set objSheet = objBook.ActiveSheet
        
    
'    lblstatus.Caption = "Status download: Mengisi field... silahkan tunggu!"
    
    
    On Error GoTo Salah
    'Proses pengsisian nama field ke excel
    Dim X, Y    As Integer
        If M_objrs.State = 1 Then
            X = 0
            Y = M_objrs.fields().Count - 1
            Do Until X > Y
                DoEvents
                objSheet.Cells(1, i).Value = CStr(M_objrs.fields(X).Name)
                i = i + 1
                X = X + 1
            Loop
        End If
    
   ' lblstatus.Caption = "Status download: Membuat file excel... silahkan tunggu!"
    objSheet.Range("A2").CopyFromRecordset M_objrs '-> Proses pengisian data dimulai dari Cell A2
    objBook.SaveAs TxtPath.Text, xlWorkbookNormal
    objExcel.Quit
    Set objExcel = Nothing: Set objBook = Nothing: Set objSheet = Nothing
    
    

'        CSVData = RecordsetToCSV(M_objrs, True)
'
'        Open "" & TxtPath + ".csv" For Binary Access Write As #1
'            Put #1, , CSVData
'        Close #1
        
    
        
    
'    lblstatus.Caption = "Status download: Mengisi field... silahkan tunggu!"
    
    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    Set M_objrs = Nothing
 
Salah:
    Exit Sub
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
    'RPT.Action = 1
    'RPT.Reset
End Sub
