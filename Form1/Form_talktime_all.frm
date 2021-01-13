VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form_talktime_all 
   BackColor       =   &H00C0FFC0&
   Caption         =   "TalkTime"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13140
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   13140
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7560
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "E&xport"
      Height          =   495
      Left            =   10200
      TabIndex        =   11
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   11640
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   8
      Text            =   "0"
      Top             =   7920
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Proce&ss"
      Height          =   495
      Left            =   8760
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7095
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   12515
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin TDBTime6Ctl.TDBTime txtWaktu1 
      Height          =   315
      Left            =   2820
      TabIndex        =   0
      Top             =   240
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   556
      Caption         =   "Form_talktime_all.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "Form_talktime_all.frx":006C
      Spin            =   "Form_talktime_all.frx":00BC
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   16777215
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
      Value           =   1.02960316199441E-317
   End
   Begin TDBDate6Ctl.TDBDate TdTglCall1 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Calendar        =   "Form_talktime_all.frx":00E4
      Caption         =   "Form_talktime_all.frx":01FC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "Form_talktime_all.frx":0268
      Keys            =   "Form_talktime_all.frx":0286
      Spin            =   "Form_talktime_all.frx":02E4
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
      Left            =   4740
      TabIndex        =   2
      Top             =   240
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Calendar        =   "Form_talktime_all.frx":030C
      Caption         =   "Form_talktime_all.frx":0424
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "Form_talktime_all.frx":0490
      Keys            =   "Form_talktime_all.frx":04AE
      Spin            =   "Form_talktime_all.frx":050C
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
   Begin TDBTime6Ctl.TDBTime TxtWaktu2 
      Height          =   315
      Left            =   6180
      TabIndex        =   3
      Top             =   240
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   556
      Caption         =   "Form_talktime_all.frx":0534
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "Form_talktime_all.frx":05A0
      Spin            =   "Form_talktime_all.frx":05F0
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   16777215
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
      Value           =   1.02960316199441E-317
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Data:"
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
      Index           =   0
      Left            =   11040
      TabIndex        =   9
      Top             =   8040
      Width           =   1455
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   3720
      TabIndex        =   5
      Top             =   300
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
      Left            =   240
      TabIndex        =   4
      Top             =   300
      Width           =   1455
   End
End
Attribute VB_Name = "Form_talktime_all"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSTEMP As ADODB.Recordset

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    On Error GoTo err
    CommonDialog1.Filter = "Excel Files (*.xls)|*.xls"
    CommonDialog1.ShowSave
    ConvertToExcel RSTEMP, CommonDialog1.FileName
err:
End Sub

Private Sub Form_Load()
    Call init_header
    Call koneksi
End Sub

Private Sub koneksi()
    Set RSTEMP = New ADODB.Recordset
    RSTEMP.ActiveConnection = M_OBJCONN
    RSTEMP.CursorLocation = adUseClient
    RSTEMP.CursorType = adOpenDynamic
    RSTEMP.LockType = adLockOptimistic
End Sub

Private Sub Command1_Click()
    M_OBJCONN.Execute "DELETE FROM mandiri.temp_tbltalktime; "
    Call func_talktime("dbname=icentra host=192.168.10.4 port=5432 password=jengkolman user=icentra")
    Call func_talktime("dbname=icentra host=192.168.10.5 port=5432 password=jengkolman user=icentra")
    Call show_data
End Sub

Private Sub func_talktime(strconfig As String)
    Dim STRSQL      As String
    Dim StrKoneksi  As String
    
    STRSQL = "SELECT ota_log_record.agent_id as agent,"
    STRSQL = STRSQL + "tbl_agent.nama as namaagent,"
    STRSQL = STRSQL + "ota_bill_record.calldate,"
    STRSQL = STRSQL + "ota_bill_record.stop_time as stoptime,"
    STRSQL = STRSQL + "ota_bill_record.duration as dur,"
    STRSQL = STRSQL + "tbl_agent.team,"
    STRSQL = STRSQL + "ota_bill_record.dst,"
    STRSQL = STRSQL + "ota_log_record.nocase as custid,"
    STRSQL = STRSQL + "ota_bill_record.billsec "
    STRSQL = STRSQL + "FROM "
    STRSQL = STRSQL + "(select  acd_agent.name as userid,"
    STRSQL = STRSQL + "acd_agent.fullname as nama,"
    STRSQL = STRSQL + "acd_queue.name as team,"
    STRSQL = STRSQL + "acd_agent.acd_agent_id "
    STRSQL = STRSQL + "FROM acd_agent_group, acd_agent, acd_queue "
    STRSQL = STRSQL + "WHERE acd_agent_group.acd_agent_id = acd_agent.acd_agent_id "
    STRSQL = STRSQL + "and  acd_agent_group.acd_queue_id=acd_queue.acd_queue_id ) as tbl_agent, "
    STRSQL = STRSQL + "ota_bill_record , ota_log_record "
    STRSQL = STRSQL + "WHERE "
    STRSQL = STRSQL + "tbl_agent.userid=ota_log_record.agent_id and "
    STRSQL = STRSQL + "ota_bill_record.uniqueid=ota_log_record.unique_id and "
    'StrSql = StrSql + "date(ota_bill_record.calldate) between '"
    STRSQL = STRSQL + " ota_bill_record.calldate between ''"
    STRSQL = STRSQL + Format(TdTglCall1.Value, "yyyy-mm-dd") & " " & txtWaktu1.Value & "'' and ''"
    STRSQL = STRSQL + Format(TdTglCall2.Value, "yyyy-mm-dd") & " " & TxtWaktu2.Value & "'' "
    'Strsql = Strsql + " tbl_agent.userid Between '" + CmbAgent1.Text + "' and '"
    'Strsql = Strsql + CmbAgent2.Text + "'"
    
    M_OBJCONN.Execute "INSERT INTO mandiri.temp_tbltalktime(agent,namaagent,calldate,stoptime,dur,team,dst,custid,billsec) " & _
                        "SELECT * FROM dblink( '" & strconfig & "',E'" & STRSQL & "')  as t " & _
                        "(agent varchar,namaagent varchar,calldate timestamp without time zone,stoptime timestamp without time zone,dur integer,team varchar,dst varchar,custid varchar,billsec integer);"
End Sub

Private Sub init_header()
    ListView1.ColumnHeaders.ADD , , "No"
    ListView1.ColumnHeaders.ADD , , "DeskColl"
    ListView1.ColumnHeaders.ADD , , "DeskColl Name"
    ListView1.ColumnHeaders.ADD , , "Team"
    ListView1.ColumnHeaders.ADD , , "Talk Time"
    ListView1.ColumnHeaders.ADD , , "Connect Time"
End Sub

Private Sub show_data()
    Dim lst     As ListItem
    Dim i       As Integer
    Dim total_connect As Long
    
    If RSTEMP.state = 1 Then RSTEMP.Close
    RSTEMP.Open "SELECT agent,namaagent,team,sum(dur) as tot_dur,sum(billsec) as tot_bill FROM mandiri.temp_tbltalktime GROUP BY agent,team,namaagent ORDER BY agent; "
    If RSTEMP.RecordCount > 0 Then
        While Not RSTEMP.EOF
            Set lst = ListView1.ListItems.ADD(, , i)
            lst.SubItems(1) = cnull(RSTEMP!AGENT)
            lst.SubItems(2) = cnull(RSTEMP!NamaAgent)
            lst.SubItems(3) = cnull(RSTEMP!TEAM)
            lst.SubItems(4) = create_duration(cnull(RSTEMP!tot_dur)) 'TalkTime
            'total_connect = cnull(rsTemp!tot_dur) - cnull(rsTemp!tot_bill)
            total_connect = cnull(RSTEMP!tot_bill)
            lst.SubItems(5) = create_duration(total_connect) 'Total Connect
            i = i + 1
            RSTEMP.MoveNext
        Wend
        Text1.Text = RSTEMP.RecordCount
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RSTEMP = Nothing
End Sub
