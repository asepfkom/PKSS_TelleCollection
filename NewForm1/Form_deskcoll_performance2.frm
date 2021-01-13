VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form_deskcoll_performance2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Deskcoll Performance"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13275
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   13275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD 
      Left            =   8760
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   7320
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   7320
      Width           =   495
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6615
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   11668
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11400
      TabIndex        =   6
      Top             =   7320
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6960
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Export To Excel"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11400
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Process"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker tgl_laporan 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MMMM-yyyy"
      Format          =   96141315
      CurrentDate     =   41610
   End
   Begin MSComCtl2.DTPicker tgl_laporan2 
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "MMMM-yyyy"
      Format          =   96141315
      CurrentDate     =   41610
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Data (s)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   7440
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bulan dan Tahun"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Group"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form_deskcoll_performance2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rs_calc As ADODB.Recordset
Private rs_temp As ADODB.Recordset

Private tgl_lap As Date
Private tgl_lap2 As Date
Private sql_str As String

Private sqlfilter As String
Private m_SortColumn As Integer
Private m_SortOrder As Integer

Private curr_jmldata As Integer

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Combo2_Click()
    Combo3.ListIndex = Combo2.ListIndex
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Combo3_Click()
    Combo2.ListIndex = Combo3.ListIndex
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Command1_Click()
    Dim lst As listItem
    Dim sqlstr As String
    Dim avg_performance As Double
    Dim avg_rank As Double
    Dim iRank As Double
    
    Dim Old_Tanggal As Date
    
    ListView1.ListItems.CLEAR
    
    sqlfilter = ""
    sqlstr = ""
    
    If Combo1.Text <> "" Then
        sqlfilter = " And a.team='" & Combo1.Text & "'"
    End If
    
    Command1.Enabled = False
    Command2.Enabled = False
    tgl_lap = Format(tgl_laporan.Value, "yyyy-mm-01")
    tgl_lap2 = DateAdd("d", -1, DateAdd("m", 1, Format(tgl_laporan2.Value, "yyyy-mm-01")))
    
    If tgl_lap > tgl_lap2 Then
        MsgBox "Tanggal awal harus lebih besar dari tanggal akhir!!", vbOKOnly + vbInformation, "INFO"
        Exit Sub
    End If
    
'    sqlstr = "SELECT a.userid,a.nama,b.paid_hours,c.total_mct,c.paydate_x FROM (SELECT userid,agent as nama,team FROM usertbl WHERE usertype=1) a,"
'    sqlstr = sqlstr & "(SELECT y.userid,sum(x.hours) as paid_hours FROM tblabsen x, usertbl y WHERE x.nopeg=y.nik_absensi AND date_part('month',tanggal)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',tanggal)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') GROUP BY userid,nopeg ) b,"
'    sqlstr = sqlstr & "(SELECT agent,sum(payment) as total_mct,to_char(paydate,'yyyy-mm')||'-01' as paydate_x FROM tbllunas WHERE date_part('month',paydate)=date_part('month',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') AND date_part('year',paydate)=date_part('year',timestamp '" & Format(tgl_lap, "yyyy-mm-dd") & "') Group By agent,paydate_x) c WHERE a.userid=b.userid AND a.userid=c.agent " & sqlFilter
    sqlstr = "SELECT a.userid,a.nama,b.paid_hours,c.total_mct,(CASE WHEN b.paid_hours > 0 THEN (c.total_mct/b.paid_hours) ELSE 0 END ) as mct_ph,c.paydate_x FROM (SELECT userid,agent as nama,team FROM usertbl WHERE usertype=1) a,"
    sqlstr = sqlstr & "(SELECT y.userid,sum(x.hours) as paid_hours,to_char(tanggal,'yyyy-mm')||'-01' as tglabsen FROM tblabsen x, usertbl y WHERE x.nopeg=y.nik_absensi AND date(tanggal) between '" & Format(tgl_lap, "yyyy-mm-dd") & "' AND '" & Format(tgl_lap2, "yyyy-mm-dd") & "' GROUP BY tglabsen,userid,nopeg ) b,"
    sqlstr = sqlstr & "(SELECT agent,sum(payment) as total_mct,to_char(paydate,'yyyy-mm')||'-01' as paydate_x FROM tbllunas WHERE date(paydate) between '" & Format(tgl_lap, "yyyy-mm-dd") & "' AND '" & Format(tgl_lap2, "yyyy-mm-dd") & "' Group By agent,paydate_x) c WHERE a.userid=b.userid AND a.userid=c.agent AND b.tglabsen=c.paydate_x " & sqlfilter
    
    If rs_calc.state = 1 Then rs_calc.Close
    rs_calc.Open sqlstr & " ORDER BY c.paydate_x,mct_ph desc "
    
    If rs_calc.RecordCount > 0 Then
    
        M_OBJCONN.Execute "DELETE FROM temp_performance;"
        
        While Not rs_calc.EOF
        
            If Old_Tanggal <> cnull(rs_calc!paydate_x) Then
                iRank = 1
                If rs_temp.state = 1 Then rs_temp.Close
                rs_temp.Open "SELECT count(a.userid) as jmluser FROM (" & rs_calc.Source & ") a WHERE a.paydate_x='" & Format(rs_calc!paydate_x, "yyyy-mm-dd") & "'"
                curr_jmldata = cnull(rs_temp!jmluser)
            End If
        
            avg_rank = Format(iRank / curr_jmldata, "0.00")
        
            M_OBJCONN.Execute "INSERT INTO temp_performance(userid,username,paid_hours,mct,mct_ph,bulan,avg_rank,rank) " & _
                                "VALUES ('" & cnull(rs_calc!Userid) & "','" & cnull(rs_calc!Nama) & "','" & cnull(rs_calc!paid_hours) & "','" & cnull(rs_calc!total_mct) & "','" & cnull(rs_calc!mct_ph) & "','" & cnull(rs_calc!paydate_x) & "'," & iRank & "," & avg_rank & ")"
            
            iRank = iRank + 1
            Old_Tanggal = cnull(rs_calc!paydate_x)
            
            rs_calc.MoveNext
        Wend
        
        If rs_temp.state = 1 Then rs_temp.Close
        rs_temp.Open "SELECT distinct bulan FROM temp_performance GROUP BY bulan"
        curr_jmldata = rs_temp.RecordCount
        
        If rs_calc.state = 1 Then rs_calc.Close
        rs_calc.Open "SELECT userid,username,sum(paid_hours) as paid_hours,sum(mct) as mct,sum(mct_ph) as mct_ph,sum(avg_rank)/ " & curr_jmldata & " as avg_rank_, sum(rank)/ " & curr_jmldata & " as Rank_ FROM temp_performance GROUP BY userid,username ORDER BY avg_rank_ "
        If rs_calc.RecordCount > 0 Then
            DoEvents
            ProgressBar1.Max = rs_calc.RecordCount
            While Not rs_calc.EOF
                DoEvents
                ProgressBar1.Value = rs_calc.Bookmark
                Set lst = ListView1.ListItems.ADD(, , cnull(rs_calc!Userid))
                lst.SubItems(1) = cnull(rs_calc!UserName)
                lst.SubItems(2) = cnull(rs_calc!paid_hours)
                lst.SubItems(3) = Format(cnull(rs_calc!mct), "#,###,###")
                lst.SubItems(4) = Format(cnull(rs_calc!mct_ph), "#,###,###")
                lst.SubItems(5) = rs_calc.Bookmark 'Format(cnull(rs_calc!avg_rank_), "#")
                lst.SubItems(6) = Format(cnull(rs_calc!Rank_), "0.00")

                rs_calc.MoveNext
            Wend

        End If
        Text1.Text = rs_calc.RecordCount
        Command2.Enabled = True
    
    End If
    
    Command1.Enabled = True
    
End Sub

Private Sub Command2_Click()
    CD.Filter = "Excel Files (*.xls)|*.xls"
    CD.ShowSave
    
    If rs_calc.state = 1 Then rs_calc.Close
    rs_calc.Open "SELECT * FROM temp_performance ORDER BY bulan,avg_rank"
    
    If CD.FileName <> "" Then
        If rs_calc.RecordCount > 0 Then
            convert_this rs_calc, CD.FileName
        Else
            MsgBox "Tidak ada data yang didownload!!", vbOKOnly + vbInformation, "INFO"
        End If
    End If
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Call koneksi
    
    ' ---- OPSI TL ----
    If rs_temp.state = 1 Then rs_temp.Close
    rs_temp.Open "SELECT distinct team as team_TL FROM usertbl WHERE team is not null AND lower(team) not in ('reserved','septian','wulan','admin')"
    Combo1.CLEAR
    Combo1.AddItem "ALL"
    Do Until rs_temp.EOF
        Combo1.AddItem IIf(IsNull(rs_temp!team_TL), "", rs_temp!team_TL)
        rs_temp.MoveNext
    Loop
    ' ------------------
    
    ListView1.ColumnHeaders.ADD , , "User ID"
    ListView1.ColumnHeaders.ADD , , "Name", 3500
    ListView1.ColumnHeaders.ADD , , "Paid Hours"
    ListView1.ColumnHeaders.ADD , , "MCT"
    ListView1.ColumnHeaders.ADD , , "MCT/PH"
    ListView1.ColumnHeaders.ADD , , "Avg Rank %"
    ListView1.ColumnHeaders.ADD , , "Rank"
    
    Command2.Enabled = False
End Sub

Private Sub koneksi()
    Set rs_calc = New ADODB.Recordset
    rs_calc.CursorLocation = adUseClient
    rs_calc.CursorType = adOpenDynamic
    rs_calc.LockType = adLockOptimistic
    rs_calc.ActiveConnection = M_OBJCONN
    
    Set rs_temp = New ADODB.Recordset
    rs_temp.CursorLocation = adUseClient
    rs_temp.CursorType = adOpenDynamic
    rs_temp.LockType = adLockOptimistic
    rs_temp.ActiveConnection = M_OBJCONN
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs_temp = Nothing
    Set rs_calc = Nothing
End Sub

Private Sub convert_this(M_Objrs As ADODB.Recordset, Txtpath As String)
    Dim listItem        As listItem
    Dim cmdsql_update   As String
    Dim objExcel        As Excel.Application
    Dim objBook         As Excel.Workbook
    Dim objSheet        As Excel.Worksheet
    Dim i As Double
    Dim m_msgbox As String
    Dim iCell           As Integer
    Dim iLastColumn     As Integer
    Dim arrAlpha
    Dim ilastrow        As Integer

    i = 1

    'Cek apakah user menekan tombol cancel pada dialog save
    If Txtpath = Empty Then
        MsgBox "Nama file tidak boleh kosong, download dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If

    'Set excel
    Set objExcel = CreateObject("Excel.Application")
    Set objBook = objExcel.Workbooks.ADD
    Set objSheet = objBook.ActiveSheet

'    lblstatus.Caption = "Status download: Mengisi field... silahkan tunggu!"

    arrAlpha = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD")

    On Error GoTo SALAH
    'Proses pengsisian nama field ke excel
    Dim x, Y    As Double
    If M_Objrs.state = 1 Then
        x = 0
        Y = M_Objrs.fields().Count - 1
        Do Until x > Y
            DoEvents
            objSheet.Cells(1, i).Value = CStr(M_Objrs.fields(x).Name)
            i = i + 1
            x = x + 1
        Loop
    End If

   ' lblstatus.Caption = "Status download: Membuat file excel... silahkan tunggu!"
    objSheet.Range("A2").CopyFromRecordset M_Objrs '-> Proses pengisian data dimulai dari Cell A2

'    For x = 3 To iLastColumn + 2
'        objSheet.Cells(iCell, x).Value = "=sum(" & arrAlpha(x - 1) & "2:" & arrAlpha(x - 1) & iCell - 1 & ")"
'    Next x

    ilastrow = M_Objrs.RecordCount + 3

    If rs_temp.state = 1 Then rs_temp.Close
    rs_temp.Open "SELECT userid,username,sum(paid_hours) as paid_hours,sum(mct) as mct,sum(mct_ph) as mct_ph,round(sum(avg_rank)/ " & curr_jmldata & ",2) as avg_rank_, round(sum(rank)/ " & curr_jmldata & ",2) as Rank_ FROM temp_performance GROUP BY userid,username ORDER BY userid "
    If rs_temp.state = 1 Then
        x = 0
        Y = rs_temp.fields().Count - 1
        i = 1
        Do Until x > Y
            DoEvents
            objSheet.Cells(ilastrow, i).Value = CStr(rs_temp.fields(x).Name)
            i = i + 1
            x = x + 1
        Loop
    End If
    objSheet.Range("A" & ilastrow + 1).CopyFromRecordset rs_temp

    objBook.SaveAs Txtpath, xlWorkbookNormal
    objExcel.Quit

    Set objExcel = Nothing: Set objBook = Nothing: Set objSheet = Nothing
    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    'Set M_Objrs = Nothing

    Exit Sub

SALAH:
    MsgBox err.Description
    Exit Sub
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   ListView1.SortKey = ColumnHeader.Index - 1
   ListView1.Sorted = True
End Sub
