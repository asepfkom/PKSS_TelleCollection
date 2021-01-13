VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmConfidenceList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Confidence report with previous month"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10410
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   10410
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1200
      TabIndex        =   14
      Top             =   600
      Width           =   1515
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      Top             =   600
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5160
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtptp2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   8400
      TabIndex        =   8
      Text            =   "0"
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox txtpayment2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   8400
      TabIndex        =   7
      Text            =   "0"
      Top             =   5400
      Width           =   1695
   End
   Begin VB.TextBox txtptp1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   2040
      TabIndex        =   6
      Text            =   "0"
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox txtpayment1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   2040
      TabIndex        =   5
      Text            =   "0"
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Export Excel"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Process"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin MSComctlLib.ListView LvPTPPayment 
      Height          =   4140
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   7303
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin TDBDate6Ctl.TDBDate txt_tgl 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   1515
      _Version        =   65536
      _ExtentX        =   2672
      _ExtentY        =   503
      Calendar        =   "FrmConfidenceList.frx":0000
      Caption         =   "FrmConfidenceList.frx":0118
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmConfidenceList.frx":0184
      Keys            =   "FrmConfidenceList.frx":01A2
      Spin            =   "FrmConfidenceList.frx":0200
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   16777215
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd/mm/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd/mm/yyyy"
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
      ShowContextMenu =   1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__/__/____"
      ValidateMode    =   0
      ValueVT         =   6815745
      Value           =   39876
      CenturyMode     =   0
   End
   Begin VB.Label Label2 
      Caption         =   "Acc Type"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label10 
      BackColor       =   &H000080FF&
      Caption         =   "Prev PTP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6480
      TabIndex        =   12
      Top             =   5760
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackColor       =   &H000080FF&
      Caption         =   "Prev Performance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6480
      TabIndex        =   11
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H000080FF&
      Caption         =   "PTP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   10
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      Caption         =   "Performance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   9
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Tanggal"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "FrmConfidenceList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private M_objrs As ADODB.Recordset
Private m_objrs2 As ADODB.Recordset
Private dTanggal_awal As Date
Private dTanggal_akhir As Date
Private dTanggal_awal_old As Date
Private dTanggal_akhir_old As Date
Private CMDSQL As String

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Command1_Click()
    Call load_ptp
End Sub

Private Sub Command2_Click()
    Call My_Export_Excel
End Sub

Private Sub Form_Load()
    Call koneksi
    Call create_header
    
    If M_objrs.State = 1 Then M_objrs.Close
    M_objrs.Open "SELECT distinct acc_type FROM mgm;"
    
    Combo1.AddItem "All"
    While Not M_objrs.EOF
        If cnull(M_objrs!acc_type) <> "" Then
            Combo1.AddItem cnull(M_objrs!acc_type)
        End If
        M_objrs.MoveNext
    Wend
    
    Combo1.Text = "All"
    
    txt_tgl.Value = Date
End Sub

Private Sub koneksi()
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.ActiveConnection = M_OBJCONN
    M_objrs.CursorType = adOpenDynamic
    M_objrs.LockType = adLockOptimistic
    
    Set m_objrs2 = New ADODB.Recordset
    m_objrs2.CursorLocation = adUseClient
    m_objrs2.ActiveConnection = M_OBJCONN
    m_objrs2.CursorType = adOpenDynamic
    m_objrs2.LockType = adLockOptimistic
End Sub

Private Sub create_header()
    LvPTPPayment.ColumnHeaders.ADD , , "TL"
    LvPTPPayment.ColumnHeaders.ADD , , "Name"
    LvPTPPayment.ColumnHeaders.ADD , , "Performance"
    LvPTPPayment.ColumnHeaders.ADD , , "PTP"
    LvPTPPayment.ColumnHeaders.ADD , , "Prev Performance"
    LvPTPPayment.ColumnHeaders.ADD , , "Prev PTP"
End Sub

Private Sub load_ptp()
    Dim ListItem As ListItem
    Dim dTotalPtp As Double
    Dim dTotalPayment As Double
    Dim dprev_TotalPtp As Double
    Dim dprev_TotalPayment As Double
    
    If M_objrs.State = 1 Then M_objrs.Close
    M_objrs.Open this_query(Combo1.Text)
    
    dTotalPtp = 0
    dTotalPayment = 0
    dprev_TotalPtp = 0
    dprev_TotalPayment = 0
    
    LvPTPPayment.ListItems.CLEAR
    
    If M_objrs.RecordCount > 0 Then
        'no = 0s
        'Dim TotalPtpValid As Long
        PB1.Max = M_objrs.RecordCount
        While Not M_objrs.EOF
            PB1.Value = M_objrs.Bookmark
            'no = no + 1
            Set ListItem = LvPTPPayment.ListItems.ADD(, , IIf(IsNull(M_objrs("team")), "", M_objrs("team")))
            ListItem.SubItems(1) = IIf(IsNull(M_objrs("name_tl")), "", M_objrs("name_tl"))
            ListItem.SubItems(2) = IIf(IsNull(M_objrs("t_payment")), "0", Format(M_objrs("t_payment"), "#,###,###"))
            ListItem.SubItems(3) = IIf(IsNull(M_objrs("t_ptp")), "0", Format(M_objrs("t_ptp"), "#,###,##"))
            ListItem.SubItems(4) = IIf(IsNull(M_objrs("old_payment")), "0", Format(M_objrs("old_payment"), "#,###,###"))
            ListItem.SubItems(5) = IIf(IsNull(M_objrs("old_ptp")), "0", Format(M_objrs("old_ptp"), "#,###,###"))
            
            dTotalPtp = dTotalPtp + M_objrs("t_ptp")
            dTotalPayment = dTotalPayment + M_objrs("t_payment")
            dprev_TotalPtp = dprev_TotalPtp + M_objrs("old_ptp")
            dprev_TotalPayment = dprev_TotalPayment + M_objrs("old_payment")
            
            M_objrs.MoveNext
        Wend
    End If
    
    txtpayment1.Text = Format(dTotalPayment, "#,###,###")
    txtpayment2.Text = Format(dprev_TotalPayment, "#,###,###")
    txtptp1.Text = Format(dTotalPtp, "#,###,###")
    txtptp2.Text = Format(dprev_TotalPtp, "#,###,###")
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set M_objrs = Nothing
    Set m_objrs2 = Nothing
End Sub

Private Sub My_Export_Excel()
    Dim a           As Long
    Dim B           As Long
    Dim ExlObj      As Excel.Application
    Dim Exlsheet    As Excel.Worksheet
    Dim ListCustId  As String
    Dim iRow        As Integer
    Dim i           As Integer
    Dim xl_abjad    As Variant
    Dim xx          As Integer
    
    xl_abjad = Array("", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    
    Set ExlObj = CreateObject("excel.application")
    ExlObj.Workbooks.ADD
    ExlObj.Visible = True
    'ExlObj.Sheets.DELETE
    
    For xx = 0 To Combo1.ListCount - 1
        If M_objrs.State = 1 Then M_objrs.Close
        M_objrs.Open this_query(Combo1.list(xx))
        
        If ExlObj.Worksheets.Count <> Combo1.ListCount Then
            ExlObj.Worksheets.ADD
        End If
        
        With ExlObj.Worksheets(xx + 1)
            ExlObj.Worksheets(xx + 1).Visible = True
            .Name = Combo1.list(xx)
            
            .Range("A1:I1").MergeCells = True
            .Range("A2:I2").MergeCells = True
            .Range("A4:I4").Font.Bold = True
            .Range("A5:I5").Font.Bold = True
            .Range("A6:I6").Font.Bold = True
            
            .Range("A5:I6").Borders.LineStyle = xlContinuous
            
            .Cells(1, 1).Value = "Confidence report with previous month"
            .Cells(1, 1).Font.Name = "Verdana"
            .Cells(1, 1).Font.Bold = True
            .Cells(2, 1).Value = "Day : " & DateDiff("d", dTanggal_awal, dTanggal_akhir)
            .Cells(2, 1).Font.Name = "Verdana"
            .Cells(2, 1).Font.Bold = True
            
            .Cells(4, 1).Value = "Month"
            .Cells(4, 2).Value = MonthName(Month(dTanggal_awal))
            .Cells(4, 7).Value = "Month"
            .Cells(4, 8).Value = MonthName(Month(dTanggal_awal_old))
            
            .Cells(5, 1).Value = "TL"
            .Cells(5, 2).Value = "Target"
            .Cells(5, 3).Value = "Performance"
            .Cells(5, 4).Value = "PTP sd Akhir Bulan"
            .Cells(5, 5).Value = "Confidence"
            .Cells(5, 6).Value = "Short to Target"
            .Cells(5, 7).Value = "Prev Performance"
            .Cells(5, 8).Value = "Prev PTP"
            .Cells(5, 9).Value = "Prev Confidence"
            
            .Cells(6, 1).Value = "Date"
            .Cells(6, 3).Value = dTanggal_awal
            .Cells(6, 7).Value = dTanggal_awal_old
            
            iRow = 6
            If M_objrs.RecordCount > 0 Then
                PB1.Max = M_objrs.RecordCount
                i = 0
                M_objrs.MoveFirst
                Do Until M_objrs.EOF
                    i = i + 1
                    iRow = iRow + 1
                    PB1.Value = M_objrs.Bookmark
                    .Cells(iRow, 1).Value = IIf(IsNull(M_objrs("name_tl")), "", M_objrs("name_tl"))
                    'Target
                    .Cells(iRow, 2).Value = ""
                    ' Performance
                    .Cells(iRow, 3).Value = IIf(IsNull(M_objrs("t_payment")), "0", Format(M_objrs("t_payment"), "#,###,###"))
                    ' PTP
                    .Cells(iRow, 4).Value = IIf(IsNull(M_objrs("t_ptp")), "0", Format(M_objrs("t_ptp"), "#,###,##"))
                    ' Confidence
                    .Cells(iRow, 5).Value = Format("=" & xl_abjad(3) & iRow & "+" & xl_abjad(4) & iRow, "#,###,###")
                    ' Short to Target
                    .Cells(iRow, 6).Value = Format("=" & xl_abjad(5) & iRow & "-" & xl_abjad(2) & iRow, "#,###,###")
                    ' OLD Performance
                    .Cells(iRow, 7).Value = IIf(IsNull(M_objrs("old_payment")), "0", Format(M_objrs("old_payment"), "#,###,###"))
                    ' OLD PTP
                    .Cells(iRow, 8).Value = IIf(IsNull(M_objrs("old_ptp")), "0", Format(M_objrs("old_ptp"), "#,###,###"))
                    ' OLD Confidence
                    .Cells(iRow, 9).Value = Format("=" & xl_abjad(7) & iRow & "+" & xl_abjad(8) & iRow, "#,###,###")
                    
                    .Range("A" & iRow & ":" & xl_abjad(9) & iRow).Borders.LineStyle = xlContinuous
                    
                    M_objrs.MoveNext
                Loop
                
                .Cells(iRow + 1, 1) = "Total"
                
                'OTOMATISASI CELL
                For iColom = 1 To 9
                    .Cells(7, iColom).EntireColumn.AutoFit
                    If iColom + 2 <= 9 Then
                        .Cells(iRow + 1, iColom + 2).Value = "=sum(" & xl_abjad(iColom + 2) & 7 & ":" & xl_abjad(iColom + 2) & iRow & ")"
                    End If
                Next
                
                .Range("A" & iRow + 1 & ":" & xl_abjad(9) & iRow + 1).Borders.LineStyle = xlContinuous
            
            End If
            
        End With
            
    Next xx
    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    PB1.Value = 0
    Command1.Enabled = True
    
    Set ExlObj = Nothing
End Sub

Private Function this_query(sAcctype As String) As String
    Dim sqlfilter As String
    
    sqlfilter = ""
    
    If sAcctype <> "All" Then
        sqlfilter = " AND acc_type='" & sAcctype & "'"
    End If
    
    dTanggal_awal = Format(txt_tgl.Value, "yyyy-mm-dd")
    dTanggal_akhir = DateAdd("d", -1, DateAdd("m", 1, Format(dTanggal_awal, "yyyy-mm-01")))
    
    dTanggal_awal_old = DateAdd("m", -1, Format(dTanggal_awal, "yyyy-mm-dd"))
    dTanggal_akhir_old = DateAdd("d", -1, Format(dTanggal_awal, "yyyy-mm-dd"))
    
'    cmdsql = "SELECT xx.*,yy.t_ptp as old_ptp,yy.t_payment as old_payment FROM ("
'    ' NEW PTP AND PAYMENT
'    cmdsql = cmdsql & "SELECT team,name_tl,sum(total_ptp) as t_ptp,sum(total_payment) as t_payment FROM (SELECT a.agent,total_ptp,total_payment FROM (SELECT m.agent, sum(ptp.promisepay) as total_ptp FROM tblnegoptp as ptp, (SELECT custid,agent,acc_type FROM mgm) as m WHERE ptp.custid = m.custid AND (date(ptp.promisedate) BETWEEN '" & Format(dTanggal_awal, "yyyy-mm-dd") & "' AND '" & Format(dTanggal_akhir, "yyyy-mm-dd") & "') AND m.custid IS NOT NULL " & sqlfilter & " GROUP BY agent) a, " & _
'            "(SELECT agent,sum(payment) as total_payment FROM tbllunas WHERE (date_part('month',paydate)=" & Month(dTanggal_awal) & " AND date_part('year',paydate)=" & Year(dTanggal_awal) & ") GROUP BY agent) b WHERE a.agent=b.agent ORDER BY agent) x, " & _
'            "(SELECT v.userid,team,w.agent as name_tl FROM usertbl v,(SELECT userid,agent FROM usertbl) w WHERE v.team=w.userid ) y WHERE x.agent=y.userid GROUP BY team,name_tl ORDER BY team) xx, "
'    ' OLD PTP AND PAYMENT
'    cmdsql = cmdsql & "(SELECT team,name_tl,sum(total_ptp) as t_ptp,sum(total_payment) as t_payment FROM (SELECT a.agent,total_ptp,total_payment FROM (SELECT m.agent, sum(ptp.promisepay) as total_ptp FROM tblnegoptp as ptp, (SELECT custid,agent,acc_type FROM mgm) as m WHERE ptp.custid = m.custid AND (date(ptp.promisedate) BETWEEN '" & Format(dTanggal_awal_old, "yyyy-mm-dd") & "' AND '" & Format(dTanggal_akhir_old, "yyyy-mm-dd") & "') AND m.custid IS NOT NULL " & sqlfilter & " GROUP BY agent) a, " & _
'            "(SELECT agent,sum(payment) as total_payment FROM tbllunas WHERE (date_part('month',paydate)=" & Month(dTanggal_awal_old) & " AND date_part('year',paydate)=" & Year(dTanggal_awal_old) & ") GROUP BY agent) b WHERE a.agent=b.agent ORDER BY agent) x, " & _
'            "(SELECT v.userid,team,w.agent as name_tl FROM usertbl v,(SELECT userid,agent FROM usertbl) w WHERE v.team=w.userid ) y WHERE x.agent=y.userid GROUP BY team,name_tl ORDER BY team) yy WHERE xx.team=yy.team "
'----- query yang lama----
'    cmdsql = "SELECT xx.*,yy.t_ptp as old_ptp,yy.t_payment as old_payment FROM "
'    cmdsql = cmdsql & "(SELECT x.team,y.agent as name_tl,sum(promisepay) as t_ptp,sum(payment) as t_payment FROM (SELECT b.agent,d.team,a.custid,a.promisedate,a.promisepay,c.paydate,c.payment FROM (SELECT * FROM tblnegoptp WHERE date(promisedate) BETWEEN '" & Format(dTanggal_awal, "yyyy-mm-dd") & "' AND '" & Format(dTanggal_akhir, "yyyy-mm-dd") & "') a,mgm b,(SELECT * FROM tbllunas WHERE date_part('month',paydate)=" & Month(dTanggal_awal) & " AND date_part('year',paydate)=" & Year(dTanggal_awal) & ") c,usertbl d WHERE a.custid=b.custid AND a.custid=c.custid AND a.promisedate<=c.paydate AND b.agent=d.userid) x,usertbl y WHERE x.team=y.userid GROUP BY x.team,y.agent) xx,"
'    cmdsql = cmdsql & "(SELECT x.team,y.agent as name_tl,sum(promisepay) as t_ptp,sum(payment) as t_payment FROM (SELECT b.agent,d.team,a.custid,a.promisedate,a.promisepay,c.paydate,c.payment FROM (SELECT * FROM tblnegoptp WHERE date(promisedate) BETWEEN '" & Format(dTanggal_awal_old, "yyyy-mm-dd") & "' AND '" & Format(dTanggal_akhir_old, "yyyy-mm-dd") & "') a,mgm b,(SELECT * FROM tbllunas WHERE date_part('month',paydate)=" & Month(dTanggal_awal_old) & " AND date_part('year',paydate)=" & Year(dTanggal_akhir_old) & ") c,usertbl d WHERE a.custid=b.custid AND a.custid=c.custid AND a.promisedate<=c.paydate AND b.agent=d.userid) x,usertbl y WHERE x.team=y.userid GROUP BY x.team,y.agent) yy WHERE xx.team=yy.team ORDER BY xx.team "
'
    
' Query yang baru creator : Budi date : 2014-07-17
CMDSQL = " SELECT xx.*,yy.t_ptp as old_ptp,yy.t_payment as old_payment FROM"
CMDSQL = CMDSQL + " (SELECT x.team,y.agent as name_tl,sum(promisepay) as t_ptp,sum(payment) as t_payment FROM"
CMDSQL = CMDSQL + "     (   SELECT b.agent,d.team,a.custid,a.promisedate,a.promisepay,c.paydate,c.payment FROM"
CMDSQL = CMDSQL + " ("
CMDSQL = CMDSQL + "     SELECT * FROM tblnegoptp aa,(select custid as custid2, paydate from tbllunas where"
CMDSQL = CMDSQL + "             date_part('month',paydate)=" & Month(dTanggal_awal) & " AND date_part('year',paydate)=" & Year(dTanggal_awal) & ") bb WHERE date(promisedate)"
CMDSQL = CMDSQL + "             Between '" & Format(dTanggal_awal, "yyyy-mm-dd") & "' AND '" & Format(dTanggal_akhir, "yyyy-mm-dd") & "' "
CMDSQL = CMDSQL + " and aa.custid=bb.custid2 and promisedate>=bb.paydate"
CMDSQL = CMDSQL + "         ) a,mgm b,"
CMDSQL = CMDSQL + "         (SELECT * FROM tbllunas WHERE date_part('month',paydate)=" & Month(dTanggal_awal) & " AND date_part('year',paydate)=" & Year(dTanggal_awal) & ") c,"
CMDSQL = CMDSQL + "         usertbl d WHERE a.custid=b.custid AND a.custid=c.custid AND b.agent=d.userid"
CMDSQL = CMDSQL + "     ) x,usertbl y WHERE x.team=y.userid GROUP BY x.team,y.agent) xx,"
' Query  data bulan sebelumnya
CMDSQL = CMDSQL + "  (SELECT x.team,y.agent as name_tl,sum(promisepay) as t_ptp,sum(payment) as t_payment FROM"
CMDSQL = CMDSQL + "     (SELECT b.agent,d.team,a.custid,a.promisedate,a.promisepay,c.paydate,c.payment FROM"
CMDSQL = CMDSQL + "         ("
CMDSQL = CMDSQL + "         SELECT * FROM tblnegoptp aa,(select custid as custid2, paydate from tbllunas where"
CMDSQL = CMDSQL + "             date_part('month',paydate)=" & Month(dTanggal_awal_old) & " AND date_part('year',paydate)=" & Year(dTanggal_awal_old) & ") bb WHERE date(promisedate)"
CMDSQL = CMDSQL + "             Between '" & Format(dTanggal_awal_old, "yyyy-mm-dd") & "' AND '" & Format(dTanggal_akhir_old, "yyyy-mm-dd") & "' "
CMDSQL = CMDSQL + " and aa.custid=bb.custid2 and promisedate>=bb.paydate"
CMDSQL = CMDSQL + "         ) a,mgm b,"
CMDSQL = CMDSQL + "         (SELECT * FROM tbllunas WHERE date_part('month',paydate)=" & Month(dTanggal_awal_old) & " AND date_part('year',paydate)=" & Year(dTanggal_awal_old) & ") c,"
CMDSQL = CMDSQL + "         usertbl d WHERE a.custid=b.custid AND a.custid=c.custid AND  b.agent=d.userid) x,usertbl y WHERE x.team=y.userid"
CMDSQL = CMDSQL + "          GROUP BY x.team,y.agent) yy"
CMDSQL = CMDSQL + "  WHERE xx.team=yy.team ORDER BY xx.team"

    
    
    
    this_query = CMDSQL
End Function
