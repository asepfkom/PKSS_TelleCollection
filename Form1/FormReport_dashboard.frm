VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FormReport_dashboard 
   Caption         =   "Report"
   ClientHeight    =   2820
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   6900
   StartUpPosition =   3  'Windows Default
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
      Height          =   360
      Left            =   4905
      Picture         =   "FormReport_dashboard.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2145
      Width           =   1620
   End
   Begin VB.CommandButton SSCommand1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Export to Excel"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   4905
      Picture         =   "FormReport_dashboard.frx":0646
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   690
      Width           =   1620
   End
   Begin VB.TextBox TxtPath 
      Enabled         =   0   'False
      Height          =   480
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   160
   End
   Begin VB.ComboBox cmbagent 
      Height          =   315
      Left            =   1425
      TabIndex        =   1
      Top             =   1185
      Width           =   2220
   End
   Begin MSComDlg.CommonDialog CD_Save 
      Left            =   3300
      Top             =   15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TDBDate6Ctl.TDBDate TdTglCall1 
      Height          =   315
      Left            =   1425
      TabIndex        =   6
      Top             =   735
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   556
      Calendar        =   "FormReport_dashboard.frx":0DAC
      Caption         =   "FormReport_dashboard.frx":0EC4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FormReport_dashboard.frx":0F30
      Keys            =   "FormReport_dashboard.frx":0F4E
      Spin            =   "FormReport_dashboard.frx":0FAC
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
      Left            =   3210
      TabIndex        =   7
      Top             =   735
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   556
      Calendar        =   "FormReport_dashboard.frx":0FD4
      Caption         =   "FormReport_dashboard.frx":10EC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FormReport_dashboard.frx":1158
      Keys            =   "FormReport_dashboard.frx":1176
      Spin            =   "FormReport_dashboard.frx":11D4
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "to"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2880
      TabIndex        =   9
      Top             =   720
      Width           =   195
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   285
      TabIndex        =   8
      Top             =   705
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Agent"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   285
      TabIndex        =   2
      Top             =   1170
      Width           =   570
   End
   Begin VB.Line Line1 
      X1              =   225
      X2              =   11580
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dashboard Agent"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   225
      TabIndex        =   0
      Top             =   75
      Width           =   2205
   End
End
Attribute VB_Name = "FormReport_dashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim sQuerySelectTemp As String
Dim sQuerySelectTempID As String
Dim sGetagent As String

Private Sub cmbagent_Change()
    load_agent
End Sub
Public Sub buatheader()
Dim list As ListItem
Set list = List_ShowReport.ListItems.ADD(, , "Report Type : Agent Dashboard Report")
Set list = List_ShowReport.ListItems.ADD(, , "interval : " + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd"))
Set list = List_ShowReport.ListItems.ADD(, , "Agent   : " + cmbagent.Text + "")
Set list = List_ShowReport.ListItems.ADD(, , "")
End Sub

Private Sub CmdProses_Click()
Dim sId As String
Dim ExlObj As Excel.Application
    List_ShowReport.ListItems.CLEAR
    Call buatheader
    sId = "report" + Format(FungsiWaktuServer, "ddhhmmss")
    
    Call createreportPTP(ExlObj)
    'Call itungtotalamount
    'Call itung_outbound(sid)
End Sub
Public Sub createreportPTP(ELIN As Excel.Application)
    
    Dim sQuerySelect As String
    Dim nilai, TotalPtp, nilai_standing, Totalstanding As Double

    arrayAlphabet = Array("", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", _
    "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", _
    "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ", _
    "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV", "CW", "CX", "CY", "CZ", _
    "DA", "DB", "DC", "DD", "DE", "DF", "DG", "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP", "DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", "DY", "DZ", _
    "EA", "EB", "EC", "ED", "EE", "EF", "EG", "EH", "EI", "EJ", "EK", "EL", "EM", "EN", "EO", "EP", "EQ", "ER", "ES", "ET", "EU", "EV", "EW", "EX", "EY", "EZ", _
    "FA", "FB", "FC", "FD", "FE", "FF", "FG", "FH", "FI", "FJ", "FK", "FL", "FM", "FN", "FO", "FP", "FQ", "FR", "FS", "FT", "FU", "FV", "FW", "FX", "FY", "FZ")
    
'------------------------------ptp-----------------------------------------------------------------------------------------------------------------

    If Not (TdTglCall1.ValueIsNull) And Not (TdTglCall2.ValueIsNull) Then
        sWhere = sWhere + "  and date(tglcall) between '" + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' "
        sWhere = sWhere + " and '" + Format(TdTglCall2.Value, "yyyy-mm-dd") + "'"
        sWhere1 = sWhere1 + "  and date(tgl) between '" + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' "
        sWhere1 = sWhere1 + " and '" + Format(TdTglCall2.Value, "yyyy-mm-dd") + "'"
    Else
        MsgBox "Tanggal Call Harus Diisi", vbInformation, "Informasi"
        Exit Sub
    End If
    
    If cmbagent.Text <> "" Then
         sWhere = sWhere + " and agent = '" + cmbagent.Text + "'"
         sWhere1 = sWhere1 + " and a.agent = '" + cmbagent.Text + "'"
    End If

    
    sQuerySelect = " select tglcall,to_char(tglcall,'HH24:MI:SS') as jam_call,name,"
    sQuerySelect = sQuerySelect + vbCrLf + " custid ,nama_agent,dateptp,amountnew from mgm where statuscall='PTP'"
    sQuerySelect = sQuerySelect + vbCrLf + sWhere
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sQuerySelect, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    Nomor = 0
    If rs.RecordCount = 0 Then
        MsgBox "Data Tidak Tersedia", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    Set ExlObj = CreateObject("excel.application")
    ExlObj.Workbooks.ADD
    ExlObj.Visible = True
    
    ExlObj.Range("A1:J1").MergeCells = True
    
    With ExlObj.ActiveSheet
        .Cells(1, 1).Value = "Report Type : Agent Dashboard Report"
        .Cells(1, 1).Font.Name = "Arial"
        .Cells(1, 1).Font.Size = "14"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Color = &HFF0000
        .Cells(2, 1).Value = "Interval : " & Format(TdTglCall1.Value, "DD-MM-YYYY") & "  to  " & Format(TdTglCall2.Value, "DD-MM-YYYY") & ""
        .Cells(2, 1).Font.Name = "Arial"
        .Cells(2, 1).Font.Size = "10"
        .Cells(2, 1).Font.Bold = True
        .Cells(3, 1).Value = "Agent  : " + cmbagent.Text + ""
        .Cells(3, 1).Font.Name = "Arial"
        .Cells(3, 1).Font.Size = "10"
        .Cells(3, 1).Font.Bold = True
        
        .Cells(5, 1).Value = "Promise to Pay Report"
        .Cells(5, 1).Font.Name = "Arial"
        .Cells(5, 1).Font.Size = "10"
        .Cells(5, 1).Font.Bold = True
        ExlObj.Range("A5:C5").MergeCells = True
        
        kolom = 1
        .Cells(6, kolom).Value = "No."
        .Cells(6, kolom).Font.Bold = True
        .Cells(6, kolom).Font.Size = 10
        .Cells(6, kolom).Borders.LineStyle = xlContinuous
        .Cells(6, kolom).Interior.Color = &HC0C000
        .Cells(6, kolom).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(6, kolom + 1).Value = "Tgl Call"
        .Cells(6, kolom + 1).Font.Bold = True
        .Cells(6, kolom + 1).Font.Size = 10
        .Cells(6, kolom + 1).Borders.LineStyle = xlContinuous
        .Cells(6, kolom + 1).Interior.Color = &HC0C000
        .Cells(6, kolom + 1).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(6, kolom + 2).Value = "Jam Call"
        .Cells(6, kolom + 2).Font.Bold = True
        .Cells(6, kolom + 2).Font.Size = 10
        .Cells(6, kolom + 2).Borders.LineStyle = xlContinuous
        .Cells(6, kolom + 2).Interior.Color = &HC0C000
        .Cells(6, kolom + 2).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(6, kolom + 3).Value = "Nama Customer"
        .Cells(6, kolom + 3).Font.Bold = True
        .Cells(6, kolom + 3).Font.Size = 10
        .Cells(6, kolom + 3).Borders.LineStyle = xlContinuous
        .Cells(6, kolom + 3).Interior.Color = &HC0C000
        .Cells(6, kolom + 3).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(6, kolom + 4).Value = "No Customer"
        .Cells(6, kolom + 4).Font.Bold = True
        .Cells(6, kolom + 4).Font.Size = 10
        .Cells(6, kolom + 4).Borders.LineStyle = xlContinuous
        .Cells(6, kolom + 4).Interior.Color = &HC0C000
        .Cells(6, kolom + 4).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(6, kolom + 5).Value = "DCR Name"
        .Cells(6, kolom + 5).Font.Bold = True
        .Cells(6, kolom + 5).Font.Size = 10
        .Cells(6, kolom + 5).Borders.LineStyle = xlContinuous
        .Cells(6, kolom + 5).Interior.Color = &HC0C000
        .Cells(6, kolom + 5).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(6, kolom + 6).Value = "Tgl PTP"
        .Cells(6, kolom + 6).Font.Bold = True
        .Cells(6, kolom + 6).Font.Size = 10
        .Cells(6, kolom + 6).Borders.LineStyle = xlContinuous
        .Cells(6, kolom + 6).Interior.Color = &HC0C000
        .Cells(6, kolom + 6).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(6, kolom + 7).Value = "PTP Amount"
        .Cells(6, kolom + 7).Font.Bold = True
        .Cells(6, kolom + 7).Font.Size = 10
        .Cells(6, kolom + 7).Borders.LineStyle = xlContinuous
        .Cells(6, kolom + 7).Interior.Color = &HC0C000
        .Cells(6, kolom + 7).HorizontalAlignment = xlCenter
    j = 6
    nilai = 0
    TotalPtp = 0
 While Not rs.EOF
        Nomor = Nomor + 1
        j = j + 1
        .Cells(j, 1) = Nomor
        .Cells(j, 1).Borders.LineStyle = xlContinuous
        .Cells(j, 1).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j, 2) = Format(cnull(rs!tglcall), "yyyy-mm-dd")
        .Cells(j, 2).Borders.LineStyle = xlContinuous
        .Cells(j, 2).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j, 3) = cnull(rs!jam_call)
        .Cells(j, 3).Borders.LineStyle = xlContinuous
        .Cells(j, 3).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j, 4) = cnull(rs!Name)
        .Cells(j, 4).Borders.LineStyle = xlContinuous
        .Cells(j, 4).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j, 5) = "'" + cnull(rs!CustId)
        .Cells(j, 5).Borders.LineStyle = xlContinuous
        .Cells(j, 5).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j, 6) = cnull(rs!nama_agent)
        .Cells(j, 6).Borders.LineStyle = xlContinuous
        .Cells(j, 6).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j, 7) = Format(cnull(rs!dateptp), "yyyy-mm-dd")
        .Cells(j, 7).Borders.LineStyle = xlContinuous
        .Cells(j, 7).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j, 8) = cnull(rs!AmountNew)
        .Cells(j, 8).Borders.LineStyle = xlContinuous
        .Cells(j, 8).HorizontalAlignment = xlCenter
        nilai = IIf(IsNull(rs!AmountNew), 0, rs!AmountNew)
        TotalPtp = TotalPtp + nilai
        rs.MoveNext
        
    Wend
        .Cells(j + 1, 6) = "Grand Total"
        .Cells(j + 1, 6).Font.Bold = True
        .Cells(j + 1, 6).Font.Color = vbRed
        ExlObj.Range(arrayAlphabet(6) & j + 1 & ":" & arrayAlphabet(7) & j + 1).Merge
        ExlObj.Range(arrayAlphabet(6) & j + 1 & ":" & arrayAlphabet(7) & j + 1).HorizontalAlignment = xlCenter
        ExlObj.Range(arrayAlphabet(6) & j + 1 & ":" & arrayAlphabet(7) & j + 1).Interior.Color = &HC0C000
        ExlObj.Range(arrayAlphabet(6) & j + 1 & ":" & arrayAlphabet(7) & j + 1).Borders.LineStyle = xlContinuous
        .Cells(j + 1, 7).Borders.LineStyle = xlContinuous
        .Cells(j + 1, 6).HorizontalAlignment = xlCenter
        .Cells(j + 1, 8) = TotalPtp
        .Cells(j + 1, 8).Font.Bold = True
        .Cells(j + 1, 8).Font.Color = vbRed
        .Cells(j + 1, 8).Interior.Color = &HC0C000
        .Cells(j + 1, 8).Borders.LineStyle = xlContinuous
        .Cells(j + 1, 8).HorizontalAlignment = xlCenter
        
 '------------------------------outbound call-----------------------------------------------------------------------------------------------------------------
        j = j + 4
        .Cells(j, 1).Value = "Outbound Call Report"
        .Cells(j, 1).Font.Name = "Arial"
        .Cells(j, 1).Font.Size = "10"
        .Cells(j, 1).Font.Bold = True
        
        sQuerySelect = "select tgl,a.custid as custid1,name,b.agent as agent1,phoneno,curbal,lastcall,hst   from mgm_hst a "
        sQuerySelect = sQuerySelect + vbCrLf + " left join mgm c on (a.custid=c.custid)"
        sQuerySelect = sQuerySelect + vbCrLf + " left join usertbl b on (a.agent=b.userid) where lastcall<>'' "
        sQuerySelect = sQuerySelect + vbCrLf + sWhere1 + "ORDER BY tgl,custid1"
        Set rs2 = New ADODB.Recordset
        rs2.CursorLocation = adUseClient
        rs2.Open sQuerySelect, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        j = j + 1
        kolom = 1
        .Cells(j, kolom).Value = "No."
        .Cells(j, kolom).Font.Bold = True
        .Cells(j, kolom).Font.Size = 10
        .Cells(j, kolom).Borders.LineStyle = xlContinuous
        .Cells(j, kolom).Interior.Color = &HC0C000
        .Cells(j, kolom).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j, kolom + 1).Value = "Call Time"
        .Cells(j, kolom + 1).Font.Bold = True
        .Cells(j, kolom + 1).ColumnWidth = 20
        .Cells(j, kolom + 1).Font.Size = 10
        .Cells(j, kolom + 1).Borders.LineStyle = xlContinuous
        .Cells(j, kolom + 1).Interior.Color = &HC0C000
        .Cells(j, kolom + 1).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j, kolom + 2).Value = "No Customer"
        .Cells(j, kolom + 2).Font.Bold = True
        .Cells(j, kolom + 2).ColumnWidth = 40
        .Cells(j, kolom + 2).Font.Size = 10
        .Cells(j, kolom + 2).Borders.LineStyle = xlContinuous
        .Cells(j, kolom + 2).Interior.Color = &HC0C000
        .Cells(j, kolom + 2).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j, kolom + 3).Value = "Nama Customer"
        .Cells(j, kolom + 3).Font.Bold = True
        .Cells(j, kolom + 3).ColumnWidth = 30
        .Cells(j, kolom + 3).Font.Size = 10
        .Cells(j, kolom + 3).Borders.LineStyle = xlContinuous
        .Cells(j, kolom + 3).Interior.Color = &HC0C000
        .Cells(j, kolom + 3).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j, kolom + 4).Value = "DCR Name"
        .Cells(j, kolom + 4).Font.Bold = True
        .Cells(j, kolom + 4).ColumnWidth = 25
        .Cells(j, kolom + 4).Font.Size = 10
        .Cells(j, kolom + 4).Borders.LineStyle = xlContinuous
        .Cells(j, kolom + 4).Interior.Color = &HC0C000
        .Cells(j, kolom + 4).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j, kolom + 5).Value = "No Telp."
        .Cells(j, kolom + 5).Font.Bold = True
        .Cells(j, kolom + 5).ColumnWidth = 15
        .Cells(j, kolom + 5).Font.Size = 10
        .Cells(j, kolom + 5).Borders.LineStyle = xlContinuous
        .Cells(j, kolom + 5).Interior.Color = &HC0C000
        .Cells(j, kolom + 5).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j, kolom + 6).Value = "Outstanding"
        .Cells(j, kolom + 6).Font.Bold = True
        .Cells(j, kolom + 6).ColumnWidth = 15
        .Cells(j, kolom + 6).Font.Size = 10
        .Cells(j, kolom + 6).Borders.LineStyle = xlContinuous
        .Cells(j, kolom + 6).Interior.Color = &HC0C000
        .Cells(j, kolom + 6).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j, kolom + 7).Value = "Status Account"
        .Cells(j, kolom + 7).Font.Bold = True
        .Cells(j, kolom + 7).ColumnWidth = 20
        .Cells(j, kolom + 7).Font.Size = 10
        .Cells(j, kolom + 7).HorizontalAlignment = xlCenter
        .Cells(j, kolom + 7).Interior.Color = &HC0C000
        .Cells(j, kolom + 7).Borders.LineStyle = xlContinuous
        '-----------------------------------------------------
        .Cells(j, kolom + 8).Value = "Remarks"
        .Cells(j, kolom + 8).Font.Bold = True
        .Cells(j, kolom + 8).ColumnWidth = 80
        .Cells(j, kolom + 8).Font.Size = 10
        .Cells(j, kolom + 8).Borders.LineStyle = xlContinuous
        .Cells(j, kolom + 8).Interior.Color = &HC0C000
        .Cells(j, kolom + 8).HorizontalAlignment = xlCenter
        Nomor = 0
        
    nilai_standing = 0
    Totalstanding = 0
 While Not rs2.EOF
        Nomor = Nomor + 1
        j = j + 1
        .Cells(j, 1) = Nomor
        .Cells(j, 1).Borders.LineStyle = xlContinuous
        .Cells(j, 1).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j, 2) = cnull(rs2!TGL)
        .Cells(j, 2).Borders.LineStyle = xlContinuous
        '-----------------------------------------------------
        .Cells(j, 3) = "'" + cnull(rs2!custid1)
        .Cells(j, 3).Borders.LineStyle = xlContinuous
        '-----------------------------------------------------
        .Cells(j, 4) = cnull(rs2!Name)
        .Cells(j, 4).Borders.LineStyle = xlContinuous
        '-----------------------------------------------------
        .Cells(j, 5) = cnull(rs2!agent1)
        .Cells(j, 5).Borders.LineStyle = xlContinuous
        '-----------------------------------------------------
        .Cells(j, 6) = "'" + cnull(rs2!phoneno)
        .Cells(j, 6).Borders.LineStyle = xlContinuous
        '-----------------------------------------------------
        .Cells(j, 7) = cnull(rs2!curbal)
        .Cells(j, 7).Borders.LineStyle = xlContinuous
        .Cells(j, 7).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j, 8) = cnull(rs2!lastcall)
        .Cells(j, 8).Borders.LineStyle = xlContinuous
        '-----------------------------------------------------
        .Cells(j, 9) = cnull(rs2!hst)
        .Cells(j, 9).Borders.LineStyle = xlContinuous
        nilai_standing = cnull(rs2!curbal)
        Totalstanding = Totalstanding + nilai_standing
        rs2.MoveNext
        
    Wend
        .Cells(j + 1, 1) = "Grand Total"
        .Cells(j + 1, 1).Font.Bold = True
        .Cells(j + 1, 1).Font.Color = vbRed
        ExlObj.Range(arrayAlphabet(1) & j + 1 & ":" & arrayAlphabet(6) & j + 1).Merge
        ExlObj.Range(arrayAlphabet(1) & j + 1 & ":" & arrayAlphabet(6) & j + 1).Interior.Color = &HC0C000
        ExlObj.Range(arrayAlphabet(1) & j + 1 & ":" & arrayAlphabet(6) & j + 1).HorizontalAlignment = xlCenter
        ExlObj.Range(arrayAlphabet(1) & j + 1 & ":" & arrayAlphabet(6) & j + 1).Borders.LineStyle = xlContinuous
        .Cells(j + 1, 6).Borders.LineStyle = xlContinuous
        .Cells(j + 1, 1).HorizontalAlignment = xlCenter
        .Cells(j + 1, 7) = Totalstanding
        .Cells(j + 1, 7).Interior.Color = &HC0C000
        .Cells(j + 1, 7).Font.Color = vbRed
        .Cells(j + 1, 7).Font.Bold = True
        .Cells(j + 1, 7).Borders.LineStyle = xlContinuous
        .Cells(j + 1, 7).HorizontalAlignment = xlCenter
        
        '------------------------------Account Summary Report-----------------------------------------------------------------------------------------------------------------
        j = j + 4
        .Cells(j, 1).Value = "Account Summary Report"
        .Cells(j, 1).Font.Name = "Arial"
        .Cells(j, 1).Font.Size = "10"
        .Cells(j, 1).Font.Bold = True
        
            strsql = " select date(tglcall)::varchar as tglcall1,total_Call1 as totalcall1,jumlah_polis as Jumlah_Polis "
        strsql = strsql + vbCrLf + " ,jml_paid as Already_Paid,jml_bp as BP,jml_ptp as PTP,jml_schedule as Schedule_Call,jml_left_msg as Left_Message,jml_nego as Negosiasi,jml_busy as Busy,jml_dead as Dead,jml_invalid as Invalid,jml_mailbox as Mailbox,jml_pndah_alamat as Pindah_Alamat,jml_salbung as Salah_Sambung,jml_tdk_ditempat as Tidak_Ada_di_Tempat,jml_tdk_diangkat as Tidak_Diangkat,jml_unknow as Unknow,jml_data_retur as Data_Retur,nominal_ptp ,tagihan  from("
        strsql = strsql + vbCrLf + "  SELECT  date(tglcall) as tglcall, "
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'Already Paid' then 1 else 0 end) as jml_paid,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'BP' then 1 else 0 end) as jml_bp,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'PTP' then 1 else 0 end) as jml_ptp,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'Schedule Call' then 1 else 0 end) as jml_schedule,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'Left Message' then 1 else 0 end) as jml_left_msg,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'Negosiasi' then 1 else 0 end) as jml_nego,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'Busy' then 1 else 0 end) as jml_busy,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'Dead' then 1 else 0 end) as jml_dead,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'Invalid' then 1 else 0 end) as jml_invalid,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'Mailbox' then 1 else 0 end) as jml_mailbox,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'Pindah Alamat' then 1 else 0 end) as jml_pndah_alamat,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'Salah Sambung' then 1 else 0 end) as jml_salbung,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'Tidak Ada di Tempat' then 1 else 0 end) as jml_tdk_ditempat,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'Tidak Diangkat' then 1 else 0 end) as jml_tdk_diangkat,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'Unknow' then 1 else 0 end) as jml_unknow,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'Data Retur' then 1 else 0 end) as jml_data_retur,"
        strsql = strsql + vbCrLf + "  sum(curbal) as tagihan,"
        strsql = strsql + vbCrLf + "  sum(amountnew) as nominal_ptp"
        strsql = strsql + vbCrLf + "  FROM mgm "
        strsql = strsql + vbCrLf + "  WHERE agent in ('" + cmbagent.Text + "')  "
        strsql = strsql + vbCrLf + "  AND date(tglcall) between '" + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' and '" + Format(TdTglCall2.Value, "yyyy-mm-dd") + "' "
        strsql = strsql + vbCrLf + "  group by date(tglcall)"
        strsql = strsql + vbCrLf + "  order by date(tglcall)"
        strsql = strsql + vbCrLf + " )x"
        strsql = strsql + vbCrLf + " left join"
        strsql = strsql + vbCrLf + " (select date(tgl) as tgl, count(date(tgl)) as total_Call1 from mgm_hst where custid in(select custid from mgm where agent in ('" + cmbagent.Text + "') AND date(tglcall) between '" + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' and '" + Format(TdTglCall2.Value, "yyyy-mm-dd") + "'  ) "
        strsql = strsql + vbCrLf + " and agent in ('" + cmbagent.Text + "') AND date(tgl) between '" + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' and '" + Format(TdTglCall2.Value, "yyyy-mm-dd") + "' group by date(tgl) order by date(tgl) "
        strsql = strsql + vbCrLf + " )y on x.tglcall=y.tgl"
        strsql = strsql + vbCrLf + " left join"
        strsql = strsql + vbCrLf + " (select date(tglcall) as tglcall1, count(date(tglcall)) as jumlah_polis from mgm where statuscall <> '' and agent in ('" + cmbagent.Text + "') AND date(tglcall) between '" + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' and '" + Format(TdTglCall2.Value, "yyyy-mm-dd") + "'    group by date(tglcall) order by date(tglcall)"
        strsql = strsql + vbCrLf + " )z on x.tglcall=z.tglcall1"
        strsql = strsql + vbCrLf + " UNION ALL"
        strsql = strsql + vbCrLf + " select 'TOTAL',sum(totalcall1),sum(Jumlah_Polis) "
        strsql = strsql + vbCrLf + " ,sum(Already_Paid),sum(BP),sum(PTP),sum(Schedule_Call),sum(Left_Message),sum(Negosiasi),sum(Busy),sum(Dead),sum(Invalid),sum(Mailbox),sum(Pindah_Alamat),sum(Salah_Sambung),sum(Tidak_Ada_di_Tempat),sum(Tidak_Diangkat),sum(Unknow),sum(Data_Retur),sum(Nominal_PTP),sum(Tagihan) from ("
        strsql = strsql + vbCrLf + " select date(tglcall) as tglcall1,total_Call1 as totalcall1,jumlah_polis as Jumlah_Polis "
        strsql = strsql + vbCrLf + " ,jml_paid as Already_Paid,jml_bp as BP,jml_ptp as PTP,jml_schedule as Schedule_Call,jml_left_msg as Left_Message,jml_nego as Negosiasi,jml_busy as Busy,jml_dead as Dead,jml_invalid as Invalid,jml_mailbox as Mailbox,jml_pndah_alamat as Pindah_Alamat,jml_salbung as Salah_Sambung,jml_tdk_ditempat as Tidak_Ada_di_Tempat,jml_tdk_diangkat as Tidak_Diangkat,jml_unknow as Unknow,jml_data_retur as Data_Retur,nominal_ptp ,tagihan  from("
        strsql = strsql + vbCrLf + "  SELECT  date(tglcall) as tglcall, "
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'Already Paid' then 1 else 0 end) as jml_paid,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'BP' then 1 else 0 end) as jml_bp,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'PTP' then 1 else 0 end) as jml_ptp,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'Schedule Call' then 1 else 0 end) as jml_schedule,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'Left Message' then 1 else 0 end) as jml_left_msg,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'Negosiasi' then 1 else 0 end) as jml_nego,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'Busy' then 1 else 0 end) as jml_busy,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'Dead' then 1 else 0 end) as jml_dead,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'Invalid' then 1 else 0 end) as jml_invalid,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'Mailbox' then 1 else 0 end) as jml_mailbox,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'Pindah Alamat' then 1 else 0 end) as jml_pndah_alamat,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'Salah Sambung' then 1 else 0 end) as jml_salbung,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'Tidak Ada di Tempat' then 1 else 0 end) as jml_tdk_ditempat,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'Tidak Diangkat' then 1 else 0 end) as jml_tdk_diangkat,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'Unknow' then 1 else 0 end) as jml_unknow,"
        strsql = strsql + vbCrLf + "  sum(CASE WHEN statuscall = 'Data Retur' then 1 else 0 end) as jml_data_retur,"
        strsql = strsql + vbCrLf + "  sum(curbal) as tagihan,"
        strsql = strsql + vbCrLf + "  sum(amountnew) as nominal_ptp"
        strsql = strsql + vbCrLf + "  FROM mgm "
        strsql = strsql + vbCrLf + "  WHERE agent in ('" + cmbagent.Text + "')  "
        strsql = strsql + vbCrLf + "  AND date(tglcall) between '" + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' and '" + Format(TdTglCall2.Value, "yyyy-mm-dd") + "' "
        strsql = strsql + vbCrLf + "  group by date(tglcall)"
        strsql = strsql + vbCrLf + "  order by date(tglcall)"
        strsql = strsql + vbCrLf + " )x"
        strsql = strsql + vbCrLf + " left join"
        strsql = strsql + vbCrLf + " (select date(tgl) as tgl, count(date(tgl)) as total_Call1 from mgm_hst where custid in(select custid from mgm where agent in ('" + cmbagent.Text + "') AND date(tglcall) between '" + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' and '" + Format(TdTglCall2.Value, "yyyy-mm-dd") + "'  ) "
        strsql = strsql + vbCrLf + " and agent in ('" + cmbagent.Text + "') AND date(tgl) between '" + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' and '" + Format(TdTglCall2.Value, "yyyy-mm-dd") + "' group by date(tgl) order by date(tgl) "
        strsql = strsql + vbCrLf + " )y on x.tglcall=y.tgl"
        strsql = strsql + vbCrLf + " left join"
        strsql = strsql + vbCrLf + " (select date(tglcall) as tglcall1, count(date(tglcall)) as jumlah_polis from mgm where statuscall <> '' and agent in ('" + cmbagent.Text + "') AND date(tglcall) between '" + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' and '" + Format(TdTglCall2.Value, "yyyy-mm-dd") + "'    group by date(tglcall) order by date(tglcall)"
        strsql = strsql + vbCrLf + " )z on x.tglcall=z.tglcall1"
        strsql = strsql + vbCrLf + " ) abc"
        Set rs3 = New ADODB.Recordset
        rs3.CursorLocation = adUseClient
        rs3.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        j = j + 1
        kolom = 1
        .Cells(j, kolom).Value = "Tanggal"
        .Cells(j, kolom).Font.Bold = True
        .Cells(j, kolom).ColumnWidth = 20
        .Cells(j, kolom).Font.Size = 10
        ExlObj.Range(arrayAlphabet(kolom) & j & ":" & arrayAlphabet(kolom) & j + 1).Merge
        ExlObj.Range(arrayAlphabet(kolom) & j & ":" & arrayAlphabet(kolom) & j + 1).Interior.Color = &HC0C000
        ExlObj.Range(arrayAlphabet(kolom) & j & ":" & arrayAlphabet(kolom) & j + 1).HorizontalAlignment = xlCenter
        ExlObj.Range(arrayAlphabet(kolom) & j & ":" & arrayAlphabet(kolom) & j + 1).VerticalAlignment = xlCenter
        ExlObj.Range(arrayAlphabet(kolom) & j & ":" & arrayAlphabet(kolom) & j + 1).Borders.LineStyle = xlContinuous
        '-----------------------------------------------------
        .Cells(j, kolom + 1).Value = "Total Call"
        .Cells(j, kolom + 1).Font.Bold = True
        .Cells(j, kolom + 1).ColumnWidth = 20
        .Cells(j, kolom + 1).Font.Size = 10
        ExlObj.Range(arrayAlphabet(kolom + 1) & j & ":" & arrayAlphabet(kolom + 1) & j + 1).Merge
        ExlObj.Range(arrayAlphabet(kolom + 1) & j & ":" & arrayAlphabet(kolom + 1) & j + 1).HorizontalAlignment = xlCenter
        ExlObj.Range(arrayAlphabet(kolom + 1) & j & ":" & arrayAlphabet(kolom + 1) & j + 1).Interior.Color = &HC0C000
        ExlObj.Range(arrayAlphabet(kolom + 1) & j & ":" & arrayAlphabet(kolom + 1) & j + 1).VerticalAlignment = xlCenter
        ExlObj.Range(arrayAlphabet(kolom + 1) & j & ":" & arrayAlphabet(kolom + 1) & j + 1).Borders.LineStyle = xlContinuous
        '-----------------------------------------------------
        .Cells(j, kolom + 2).Value = "Jumlah Polis"
        .Cells(j, kolom + 2).Font.Bold = True
        .Cells(j, kolom + 2).ColumnWidth = 20
        .Cells(j, kolom + 2).Font.Size = 10
        ExlObj.Range(arrayAlphabet(kolom + 2) & j & ":" & arrayAlphabet(kolom + 2) & j + 1).Merge
        ExlObj.Range(arrayAlphabet(kolom + 2) & j & ":" & arrayAlphabet(kolom + 2) & j + 1).HorizontalAlignment = xlCenter
        ExlObj.Range(arrayAlphabet(kolom + 2) & j & ":" & arrayAlphabet(kolom + 2) & j + 1).Interior.Color = &HC0C000
        ExlObj.Range(arrayAlphabet(kolom + 2) & j & ":" & arrayAlphabet(kolom + 2) & j + 1).VerticalAlignment = xlCenter
        ExlObj.Range(arrayAlphabet(kolom + 2) & j & ":" & arrayAlphabet(kolom + 2) & j + 1).Borders.LineStyle = xlContinuous
        '-----------------------------------------------------
        .Cells(j, kolom + 3).Value = "Status"
        .Cells(j, kolom + 3).Font.Bold = True
        .Cells(j, kolom + 3).Font.Size = 10
        ExlObj.Range(arrayAlphabet(kolom + 3) & j & ":" & arrayAlphabet(kolom + 18) & j).Merge
        ExlObj.Range(arrayAlphabet(kolom + 3) & j & ":" & arrayAlphabet(kolom + 18) & j).HorizontalAlignment = xlCenter
        ExlObj.Range(arrayAlphabet(kolom + 3) & j & ":" & arrayAlphabet(kolom + 18) & j).Interior.Color = &HC0C000
        ExlObj.Range(arrayAlphabet(kolom + 3) & j & ":" & arrayAlphabet(kolom + 18) & j).Borders.LineStyle = xlContinuous
        '-----------------------------------------------------
        .Cells(j + 1, kolom + 3).Value = "Already Paid"
        .Cells(j + 1, kolom + 3).Font.Bold = True
        .Cells(j + 1, kolom + 3).Font.Size = 10
        .Cells(j + 1, kolom + 3).Interior.Color = &HC0E0FF
        .Cells(j + 1, kolom + 3).Borders.LineStyle = xlContinuous
        .Cells(j + 1, kolom + 3).HorizontalAlignment = xlCenter
        .Cells(j + 1, kolom + 3).VerticalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j + 1, kolom + 4).Value = "BP"
        .Cells(j + 1, kolom + 4).Font.Bold = True
        .Cells(j + 1, kolom + 4).ColumnWidth = 20
        .Cells(j + 1, kolom + 4).Font.Size = 10
        .Cells(j + 1, kolom + 4).Interior.Color = &H80FFFF
        .Cells(j + 1, kolom + 4).Borders.LineStyle = xlContinuous
        .Cells(j + 1, kolom + 4).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j + 1, kolom + 5).Value = "PTP"
        .Cells(j + 1, kolom + 5).Font.Bold = True
        .Cells(j + 1, kolom + 5).ColumnWidth = 20
        .Cells(j + 1, kolom + 5).Font.Size = 10
        .Cells(j + 1, kolom + 5).Interior.Color = &HC0E0FF
        .Cells(j + 1, kolom + 5).Borders.LineStyle = xlContinuous
        .Cells(j + 1, kolom + 5).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j + 1, kolom + 6).Value = "Schedule Call"
        .Cells(j + 1, kolom + 6).Font.Bold = True
        .Cells(j + 1, kolom + 6).ColumnWidth = 20
        .Cells(j + 1, kolom + 6).Font.Size = 10
        .Cells(j + 1, kolom + 6).Interior.Color = &H80FFFF
        .Cells(j + 1, kolom + 6).Borders.LineStyle = xlContinuous
        .Cells(j + 1, kolom + 6).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j + 1, kolom + 7).Value = "Left Message"
        .Cells(j + 1, kolom + 7).Font.Bold = True
        .Cells(j + 1, kolom + 7).ColumnWidth = 20
        .Cells(j + 1, kolom + 7).Font.Size = 10
        .Cells(j + 1, kolom + 7).Interior.Color = &HC0E0FF
        .Cells(j + 1, kolom + 7).HorizontalAlignment = xlCenter
        .Cells(j + 1, kolom + 7).Borders.LineStyle = xlContinuous
        '-----------------------------------------------------
        .Cells(j + 1, kolom + 8).Value = "Negosiasi"
        .Cells(j + 1, kolom + 8).Font.Bold = True
        .Cells(j + 1, kolom + 8).Font.Size = 10
        .Cells(j + 1, kolom + 8).Interior.Color = &H80FFFF
        .Cells(j + 1, kolom + 8).Borders.LineStyle = xlContinuous
        .Cells(j + 1, kolom + 8).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j + 1, kolom + 9).Value = "Busy"
        .Cells(j + 1, kolom + 9).Font.Bold = True
        .Cells(j + 1, kolom + 9).ColumnWidth = 20
        .Cells(j + 1, kolom + 9).Font.Size = 10
        .Cells(j + 1, kolom + 9).Interior.Color = &HC0E0FF
        .Cells(j + 1, kolom + 9).Borders.LineStyle = xlContinuous
        .Cells(j + 1, kolom + 9).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j + 1, kolom + 10).Value = "Dead"
        .Cells(j + 1, kolom + 10).Font.Bold = True
        .Cells(j + 1, kolom + 10).ColumnWidth = 20
        .Cells(j + 1, kolom + 10).Font.Size = 10
        .Cells(j + 1, kolom + 10).Interior.Color = &H80FFFF
        .Cells(j + 1, kolom + 10).Borders.LineStyle = xlContinuous
        .Cells(j + 1, kolom + 10).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j + 1, kolom + 11).Value = "Invalid"
        .Cells(j + 1, kolom + 11).Font.Bold = True
        .Cells(j + 1, kolom + 11).ColumnWidth = 20
        .Cells(j + 1, kolom + 11).Font.Size = 10
        .Cells(j + 1, kolom + 11).Interior.Color = &HC0E0FF
        .Cells(j + 1, kolom + 11).Borders.LineStyle = xlContinuous
        .Cells(j + 1, kolom + 11).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j + 1, kolom + 12).Value = "Mailbox"
        .Cells(j + 1, kolom + 12).Font.Bold = True
        .Cells(j + 1, kolom + 12).ColumnWidth = 20
        .Cells(j + 1, kolom + 12).Font.Size = 10
        .Cells(j + 1, kolom + 12).Interior.Color = &H80FFFF
        .Cells(j + 1, kolom + 12).Borders.LineStyle = xlContinuous
        .Cells(j + 1, kolom + 12).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j + 1, kolom + 13).Value = "Pindah Alamat"
        .Cells(j + 1, kolom + 13).Font.Bold = True
        .Cells(j + 1, kolom + 13).ColumnWidth = 20
        .Cells(j + 1, kolom + 13).Font.Size = 10
        .Cells(j + 1, kolom + 13).Interior.Color = &HC0E0FF
        .Cells(j + 1, kolom + 13).Borders.LineStyle = xlContinuous
        .Cells(j + 1, kolom + 13).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j + 1, kolom + 14).Value = "Salah Sambung"
        .Cells(j + 1, kolom + 14).Font.Bold = True
        .Cells(j + 1, kolom + 14).ColumnWidth = 20
        .Cells(j + 1, kolom + 14).Font.Size = 10
        .Cells(j + 1, kolom + 14).Interior.Color = &H80FFFF
        .Cells(j + 1, kolom + 14).Borders.LineStyle = xlContinuous
        .Cells(j + 1, kolom + 14).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j + 1, kolom + 15).Value = "Tidak Ada di Tempat"
        .Cells(j + 1, kolom + 15).Font.Bold = True
        .Cells(j + 1, kolom + 15).ColumnWidth = 20
        .Cells(j + 1, kolom + 15).Font.Size = 10
        .Cells(j + 1, kolom + 15).Interior.Color = &HC0E0FF
        .Cells(j + 1, kolom + 15).Borders.LineStyle = xlContinuous
        .Cells(j + 1, kolom + 15).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j + 1, kolom + 16).Value = "Tidak Diangkat"
        .Cells(j + 1, kolom + 16).Font.Bold = True
        .Cells(j + 1, kolom + 16).ColumnWidth = 20
        .Cells(j + 1, kolom + 16).Font.Size = 10
        .Cells(j + 1, kolom + 16).Interior.Color = &H80FFFF
        .Cells(j + 1, kolom + 16).Borders.LineStyle = xlContinuous
        .Cells(j + 1, kolom + 16).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j + 1, kolom + 17).Value = "Unknow"
        .Cells(j + 1, kolom + 17).Font.Bold = True
        .Cells(j + 1, kolom + 17).ColumnWidth = 20
        .Cells(j + 1, kolom + 17).Font.Size = 10
        .Cells(j + 1, kolom + 17).Interior.Color = &HC0E0FF
        .Cells(j + 1, kolom + 17).Borders.LineStyle = xlContinuous
        .Cells(j + 1, kolom + 17).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j + 1, kolom + 18).Value = "Data Retur"
        .Cells(j + 1, kolom + 18).Font.Bold = True
        .Cells(j + 1, kolom + 18).ColumnWidth = 20
        .Cells(j + 1, kolom + 18).Font.Size = 10
        .Cells(j + 1, kolom + 18).Interior.Color = &HC0E0FF
        .Cells(j + 1, kolom + 18).Borders.LineStyle = xlContinuous
        .Cells(j + 1, kolom + 18).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j, kolom + 19).Value = "Nominal PTP"
        .Cells(j, kolom + 19).Font.Bold = True
        .Cells(j, kolom + 19).ColumnWidth = 20
        .Cells(j, kolom + 19).Font.Size = 10
        ExlObj.Range(arrayAlphabet(kolom + 19) & j & ":" & arrayAlphabet(kolom + 19) & j + 1).Merge
        ExlObj.Range(arrayAlphabet(kolom + 19) & j & ":" & arrayAlphabet(kolom + 19) & j + 1).HorizontalAlignment = xlCenter
        ExlObj.Range(arrayAlphabet(kolom + 19) & j & ":" & arrayAlphabet(kolom + 19) & j + 1).Interior.Color = &HC0C000
        ExlObj.Range(arrayAlphabet(kolom + 19) & j & ":" & arrayAlphabet(kolom + 19) & j + 1).VerticalAlignment = xlCenter
        ExlObj.Range(arrayAlphabet(kolom + 19) & j & ":" & arrayAlphabet(kolom + 19) & j + 1).Borders.LineStyle = xlContinuous
        '-----------------------------------------------------
        .Cells(j, kolom + 20).Value = "Tagihan"
        .Cells(j, kolom + 20).Font.Bold = True
        .Cells(j, kolom + 20).ColumnWidth = 20
        .Cells(j, kolom + 20).Font.Size = 10
        ExlObj.Range(arrayAlphabet(kolom + 20) & j & ":" & arrayAlphabet(kolom + 20) & j + 1).Merge
        ExlObj.Range(arrayAlphabet(kolom + 20) & j & ":" & arrayAlphabet(kolom + 20) & j + 1).HorizontalAlignment = xlCenter
        ExlObj.Range(arrayAlphabet(kolom + 20) & j & ":" & arrayAlphabet(kolom + 20) & j + 1).Interior.Color = &HC0C000
        ExlObj.Range(arrayAlphabet(kolom + 20) & j & ":" & arrayAlphabet(kolom + 20) & j + 1).VerticalAlignment = xlCenter
        ExlObj.Range(arrayAlphabet(kolom + 20) & j & ":" & arrayAlphabet(kolom + 20) & j + 1).Borders.LineStyle = xlContinuous
        Nomor = 0
       j = j + 1
    nilai_standing = 0
    Totalstanding = 0
 While Not rs3.EOF
        Nomor = Nomor + 1
        j = j + 1
        .Cells(j, 1) = cnull(rs3!tglcall1)
        .Cells(j, 1).Borders.LineStyle = xlContinuous
        .Cells(j, 1).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j, 2) = cnull(rs3!totalcall1)
        .Cells(j, 2).Borders.LineStyle = xlContinuous
        .Cells(j, 2).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j, 3) = cnull(rs3!Jumlah_Polis)
        .Cells(j, 3).Borders.LineStyle = xlContinuous
        .Cells(j, 3).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j, 4) = cnull(rs3!Already_Paid)
        .Cells(j, 4).Borders.LineStyle = xlContinuous
        .Cells(j, 4).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j, 5) = cnull(rs3!BP)
        .Cells(j, 5).Borders.LineStyle = xlContinuous
        .Cells(j, 5).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j, 6) = cnull(rs3!ptp)
        .Cells(j, 6).Borders.LineStyle = xlContinuous
        .Cells(j, 6).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j, 7) = cnull(rs3!Schedule_Call)
        .Cells(j, 7).Borders.LineStyle = xlContinuous
        .Cells(j, 7).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j, 8) = cnull(rs3!Left_Message)
        .Cells(j, 8).Borders.LineStyle = xlContinuous
        '-----------------------------------------------------
        .Cells(j, 9) = cnull(rs3!Negosiasi)
        .Cells(j, 9).Borders.LineStyle = xlContinuous
        .Cells(j, 9).HorizontalAlignment = xlCenter
         '-----------------------------------------------------
        .Cells(j, 10) = cnull(rs3!Busy)
        .Cells(j, 10).Borders.LineStyle = xlContinuous
        .Cells(j, 10).HorizontalAlignment = xlCenter
         '-----------------------------------------------------
        .Cells(j, 11) = cnull(rs3!Dead)
        .Cells(j, 11).Borders.LineStyle = xlContinuous
        .Cells(j, 11).HorizontalAlignment = xlCenter
         '-----------------------------------------------------
        .Cells(j, 12) = cnull(rs3!Invalid)
        .Cells(j, 12).Borders.LineStyle = xlContinuous
        .Cells(j, 12).HorizontalAlignment = xlCenter
         '-----------------------------------------------------
        .Cells(j, 13) = cnull(rs3!Mailbox)
        .Cells(j, 13).Borders.LineStyle = xlContinuous
        .Cells(j, 13).HorizontalAlignment = xlCenter
         '-----------------------------------------------------
        .Cells(j, 14) = cnull(rs3!Pindah_Alamat)
        .Cells(j, 14).Borders.LineStyle = xlContinuous
        .Cells(j, 14).HorizontalAlignment = xlCenter
         '-----------------------------------------------------
        .Cells(j, 15) = cnull(rs3!Salah_Sambung)
        .Cells(j, 15).Borders.LineStyle = xlContinuous
        .Cells(j, 15).HorizontalAlignment = xlCenter
         '-----------------------------------------------------
        .Cells(j, 16) = cnull(rs3!Tidak_Ada_di_Tempat)
        .Cells(j, 16).Borders.LineStyle = xlContinuous
        .Cells(j, 16).HorizontalAlignment = xlCenter
         '-----------------------------------------------------
        .Cells(j, 17) = cnull(rs3!Tidak_Diangkat)
        .Cells(j, 17).Borders.LineStyle = xlContinuous
        .Cells(j, 17).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j, 18) = cnull(rs3!Unknow)
        .Cells(j, 18).Borders.LineStyle = xlContinuous
        .Cells(j, 18).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j, 19) = cnull(rs3!Data_Retur)
        .Cells(j, 19).Borders.LineStyle = xlContinuous
        .Cells(j, 19).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j, 20) = cnull(rs3!nominal_ptp)
        .Cells(j, 20).Borders.LineStyle = xlContinuous
        .Cells(j, 20).HorizontalAlignment = xlCenter
        '-----------------------------------------------------
        .Cells(j, 21) = cnull(rs3!tagihan)
        .Cells(j, 21).Borders.LineStyle = xlContinuous
        .Cells(j, 21).HorizontalAlignment = xlCenter
        rs3.MoveNext
        
    Wend
        .Cells(j, 1).Font.Bold = True
        .Cells(j, 1).Interior.Color = &HC0C000
        .Cells(j, 1).Font.Color = vbRed
        '-----------------------------------------------
        .Cells(j, 2).Font.Bold = True
        .Cells(j, 2).Interior.Color = &HC0C000
        .Cells(j, 2).Font.Color = vbRed
        '-----------------------------------------------
        .Cells(j, 3).Font.Bold = True
        .Cells(j, 3).Interior.Color = &HC0C000
        .Cells(j, 3).Font.Color = vbRed
        '-----------------------------------------------
        .Cells(j, 4).Font.Bold = True
        .Cells(j, 4).Interior.Color = &HC0C000
        .Cells(j, 4).Font.Color = vbRed
        '-----------------------------------------------
        .Cells(j, 5).Font.Bold = True
        .Cells(j, 5).Interior.Color = &HC0C000
        .Cells(j, 5).Font.Color = vbRed
        '-----------------------------------------------
        .Cells(j, 6).Font.Bold = True
        .Cells(j, 6).Interior.Color = &HC0C000
        .Cells(j, 6).Font.Color = vbRed
        '-----------------------------------------------
        .Cells(j, 7).Font.Bold = True
        .Cells(j, 7).Interior.Color = &HC0C000
        .Cells(j, 7).Font.Color = vbRed
        '-----------------------------------------------
        .Cells(j, 8).Font.Bold = True
        .Cells(j, 8).Interior.Color = &HC0C000
        .Cells(j, 8).Font.Color = vbRed
        '-----------------------------------------------
        .Cells(j, 9).Font.Bold = True
        .Cells(j, 9).Interior.Color = &HC0C000
        .Cells(j, 9).Font.Color = vbRed
        '-----------------------------------------------
        .Cells(j, 10).Font.Bold = True
        .Cells(j, 10).Interior.Color = &HC0C000
        .Cells(j, 10).Font.Color = vbRed
        '-----------------------------------------------
        .Cells(j, 11).Font.Bold = True
        .Cells(j, 11).Interior.Color = &HC0C000
        .Cells(j, 11).Font.Color = vbRed
        '-----------------------------------------------
        .Cells(j, 12).Font.Bold = True
        .Cells(j, 12).Interior.Color = &HC0C000
        .Cells(j, 12).Font.Color = vbRed
        '-----------------------------------------------
        .Cells(j, 13).Font.Bold = True
        .Cells(j, 13).Interior.Color = &HC0C000
        .Cells(j, 13).Font.Color = vbRed
        '-----------------------------------------------
        .Cells(j, 14).Font.Bold = True
        .Cells(j, 14).Interior.Color = &HC0C000
        .Cells(j, 14).Font.Color = vbRed
        '-----------------------------------------------
        .Cells(j, 15).Font.Bold = True
        .Cells(j, 15).Interior.Color = &HC0C000
        .Cells(j, 15).Font.Color = vbRed
        '-----------------------------------------------
        .Cells(j, 16).Font.Bold = True
        .Cells(j, 16).Interior.Color = &HC0C000
        .Cells(j, 16).Font.Color = vbRed
        '-----------------------------------------------
        .Cells(j, 17).Font.Bold = True
        .Cells(j, 17).Interior.Color = &HC0C000
        .Cells(j, 17).Font.Color = vbRed
        '-----------------------------------------------
        .Cells(j, 18).Font.Bold = True
        .Cells(j, 18).Interior.Color = &HC0C000
        .Cells(j, 18).Font.Color = vbRed
        '-----------------------------------------------
        .Cells(j, 19).Font.Bold = True
        .Cells(j, 19).Interior.Color = &HC0C000
        .Cells(j, 19).Font.Color = vbRed
        '-----------------------------------------------
        .Cells(j, 20).Font.Bold = True
        .Cells(j, 20).Interior.Color = &HC0C000
        .Cells(j, 20).Font.Color = vbRed
        '-----------------------------------------------
        .Cells(j, 21).Font.Bold = True
        .Cells(j, 21).Interior.Color = &HC0C000
        .Cells(j, 21).Font.Color = vbRed
    End With
End Sub
Public Sub load_agent()
    sStrsql = " select userid from usertbl where aktif='1' AND kdlevel='1' "
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        cmbagent.CLEAR
        While Not M_objrs.EOF
                cmbagent.AddItem IIf(IsNull(M_objrs!USERID), "", M_objrs!USERID)
                M_objrs.MoveNext
        Wend
    Set M_objrs = Nothing
End Sub
Private Sub cmbagent_DropDown()
    load_agent
End Sub
Private Sub SSCommand1_Click()
Dim ExlObj As Excel.Application
    Call createreportPTP(ExlObj)
End Sub
Private Sub SSCommand2_Click()
    Unload Me
End Sub
