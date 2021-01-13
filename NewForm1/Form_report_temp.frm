VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form_report_temp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form Report Temp"
   ClientHeight    =   9900
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12225
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   12225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Download SUM To Excel"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   9480
      Width           =   2535
   End
   Begin VB.TextBox txtsum 
      Enabled         =   0   'False
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
      Left            =   2640
      TabIndex        =   15
      Top             =   9360
      Width           =   1695
   End
   Begin VB.CommandButton btnreporttemp 
      Caption         =   "Report Temp"
      Height          =   375
      Left            =   9480
      TabIndex        =   14
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton btnsumtemp 
      Caption         =   "Rep.Sum Temp"
      Height          =   375
      Left            =   10680
      TabIndex        =   13
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmd_exit 
      BackColor       =   &H008080FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmd_download 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Download To Excel"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9000
      Width           =   2535
   End
   Begin VB.TextBox Txtpath 
      Height          =   285
      Left            =   3000
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txt_jml 
      Enabled         =   0   'False
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
      Left            =   2640
      TabIndex        =   7
      Top             =   8880
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmd_proses 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Process"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin TDBDate6Ctl.TDBDate tgl_mulai1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1845
      _Version        =   65536
      _ExtentX        =   3254
      _ExtentY        =   661
      Calendar        =   "Form_report_temp.frx":0000
      Caption         =   "Form_report_temp.frx":0118
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "Form_report_temp.frx":0184
      Keys            =   "Form_report_temp.frx":01A2
      Spin            =   "Form_report_temp.frx":0200
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
      Format          =   "dd, mmm yyyy"
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
      Text            =   "__, ___ ____"
      ValidateMode    =   0
      ValueVT         =   6815745
      Value           =   39876
      CenturyMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate tgl_akhir1 
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   720
      Width           =   1845
      _Version        =   65536
      _ExtentX        =   3254
      _ExtentY        =   661
      Calendar        =   "Form_report_temp.frx":0228
      Caption         =   "Form_report_temp.frx":0340
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "Form_report_temp.frx":03AC
      Keys            =   "Form_report_temp.frx":03CA
      Spin            =   "Form_report_temp.frx":0428
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
      Format          =   "dd, mmm yyyy"
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
      Text            =   "__, ___ ____"
      ValidateMode    =   0
      ValueVT         =   6815745
      Value           =   39876
      CenturyMode     =   0
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   345
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.ListView lv_report_temp 
      Height          =   6945
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   12250
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
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lv_report_sum 
      Height          =   6945
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Visible         =   0   'False
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   12250
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
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComDlg.CommonDialog CD_save 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      BackColor       =   &H00F1E5DB&
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Sum (Account)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   9480
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00F1E5DB&
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Account"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   9000
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   4
      Left            =   0
      Picture         =   "Form_report_temp.frx":0450
      Stretch         =   -1  'True
      Top             =   60
      Width           =   420
   End
   Begin VB.Label Label4 
      BackColor       =   &H00F1E5DB&
      BackStyle       =   0  'Transparent
      Caption         =   "Report Temp"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   510
      TabIndex        =   4
      Top             =   0
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "Form_report_temp.frx":0F5A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12600
   End
   Begin VB.Label Label2 
      BackColor       =   &H00F1E5DB&
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   720
      Width           =   375
   End
End
Attribute VB_Name = "Form_report_temp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnreporttemp_Click()
    lv_report_temp.Visible = True
    lv_report_sum.Visible = False
End Sub

Private Sub btnsumtemp_Click()
    lv_report_temp.Visible = False
    lv_report_sum.Visible = True
End Sub

Private Sub cmd_download_Click()
'    Me.MousePointer = vbArrowHourglass
'    cmd_download.Enabled = False
'    Call My_Export_Excel
'    Me.MousePointer = vbArrow
'19102016TIAN
    Dim objExcel As New Excel.Application
        Dim objExcelSheet As Excel.Worksheet
        Dim col, row As Integer
        Dim a As String
        If lv_report_temp.ListItems.Count > 0 Then
            objExcel.Workbooks.ADD
            Set objExcelSheet = objExcel.Worksheets.ADD
         
        
            For col = 1 To lv_report_temp.ColumnHeaders.Count
                objExcelSheet.Cells(1, col).Value = lv_report_temp.ColumnHeaders(col)
            Next
         
            For row = 2 To lv_report_temp.ListItems.Count + 1
                For col = 1 To lv_report_temp.ColumnHeaders.Count
                If col = 1 Then
                        objExcelSheet.Cells(row, col).Value = lv_report_temp.ListItems(row - 1).Text
                Else
                    '" 'cararandy 29032016 "
                    Dim hasil1 As String
                        hasil1 = "'" + lv_report_temp.ListItems(row - 1).SubItems(col - 1)
                        objExcelSheet.Cells(row, col).Value = hasil1
                    End If
                Next
            Next
         
            objExcelSheet.Columns.AutoFit
            Cd_save.ShowOpen
            a = Cd_save.FileName
         
            objExcelSheet.SaveAs a & ".xls"
            MsgBox "Export Completed", vbInformation, Me.Caption
         
            objExcel.Workbooks.Open a & ".xls"
            objExcel.Visible = True
        Else
            MsgBox "No data to export", vbInformation, Me.Caption
        End If
End Sub

Private Sub My_Export_Excel()
    Dim a           As Long
    Dim B           As Long
    Dim ExlObj      As Excel.Application
    Dim listcustid  As String
    Dim rs          As ADODB.Recordset
    Dim RS2         As ADODB.Recordset
    Dim iRow        As Integer
    Dim i           As Integer
    Dim sQuery      As String
    Dim totalcall   As Double
    Dim totaldata   As Double
    Dim ratarata   As Double
    
'    Strsql = "SELECT * FROM ("
'    Strsql = Strsql + " SELECT '" & CmbApprove.Text & "' as Approved,* "
'    Strsql = Strsql + " FROM tblsendptp WHERE custid in (" & listcustid & ")) As a"
'    Strsql = Strsql + " LEFT JOIN (SELECT custid, OpenDate, b_d FROM mgm WHERE custid in (" & listcustid & ")) As b"
'    Strsql = Strsql + " on a.custid = b.custid"
'
'    Set RS = New ADODB.Recordset
'    RS.CursorLocation = adUseClient
'    RS.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        
    jam_mulai = "00:00:00"
    jam_selesai = "23:59:59"
    
    tgl_mulai = Format(tgl_mulai1.Value, "DD-MM-YYYY")
    tgl_akhir = Format(tgl_akhir1.Value, "DD-MM-YYYY")
    
    M_OBJCONN.Execute "DROP VIEW IF EXISTS mandiri.view_temp_jumlah_call "
    
    cQuery = "CREATE mandiri.VIEW view_temp_jumlah_call AS ("
    cQuery = cQuery + "SELECT a.custid, tgl_call, jumlah, f_cek_new, descol  FROM ("
    cQuery = cQuery + " SELECT * FROM mandiri.tbl_temp_jumlah_call WHERE tgl_call"
    cQuery = cQuery + " BETWEEN '" & tgl_mulai & " " & jam_mulai & "' AND '" & tgl_akhir & " " & jam_selesai & "') As a LEFT JOIN "
    cQuery = cQuery + " (SELECT custid, f_cek_new, b.agent as descol FROM mandiri.mgm a, mandiri.usertbl b "
    cQuery = cQuery + " WHERE a.agent = b. userid AND tglcall BETWEEN '" & tgl_mulai & " " & jam_mulai & "' AND '" & tgl_akhir & " " & jam_selesai & "') As b "
    cQuery = cQuery + " on a.custid = b.custid )"
    
    M_OBJCONN.Execute cQuery
    'M_OBJCONN.Execute "ALTER VIEW view_temp_jumlah_call OWNER TO new_owner"
    
    'QUERY SELECT DETAIL DATA
    cQuery = "SELECT * FROM mandiri.view_temp_jumlah_call order by custid"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    'QUERY ITUNG JUMLAH DATA
    sQuery = "SELECT tgl_call, sum(jumlah) as total_call, count(jumlah) as total_data "
    sQuery = sQuery + " FROM mandiri.view_temp_jumlah_call group by tgl_call order by tgl_call"
    Set RS2 = New ADODB.Recordset
    RS2.CursorLocation = adUseClient
    RS2.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic

    Set ExlObj = CreateObject("excel.application")
    ExlObj.Workbooks.ADD
    ExlObj.Visible = True
    
    ExlObj.Range("A1:N1").MergeCells = True
    'ExlObj.Range("A2:N2").MergeCells = True
    ExlObj.Range("A4:N4").Font.Bold = True
    
    
    With ExlObj.ActiveSheet
        .Cells(1, 1).Value = "Report Temp - Tanggal " & Format(tgl_mulai1.Value, "DD-MM-YYYY") & " Sampai " & Format(tgl_akhir1.Value, "DD-MM-YYYY")
        .Cells(1, 1).Font.Name = "Verdana"
        .Cells(1, 1).Font.Bold = True
        .Cells(4, 1).Value = "NO"
        .Cells(4, 2).Value = "CARD NUMBER"
        .Cells(4, 3).Value = "TANGGAL CALL"
        .Cells(4, 4).Value = "JUMLAH CALL"
        .Cells(4, 5).Value = "STATUS CALL TERAKHIR"
        .Cells(4, 6).Value = "AGENT"
        .Cells(4, 9).Value = "TANGGAL"
        .Cells(4, 10).Value = "TOTAL CALL"
        .Cells(4, 11).Value = "TOTAL DATA"
        .Cells(4, 12).Value = "RATA - RATA"
        
        iRow = 4
        If RS2.RecordCount > 0 Then
            ProgressBar1.Max = RS2.RecordCount
            i = 0
            Do Until RS2.EOF
                i = i + 1
                iRow = iRow + 1
                ProgressBar1.Value = RS2.Bookmark
                .Cells(iRow, 9).Value = IIf(IsNull(RS2!tgl_call), "", Format(RS2("tgl_call"), "yyyy-mm-dd"))
                .Cells(iRow, 10).Value = IIf(IsNull(RS2!total_call), "", RS2!total_call)
                .Cells(iRow, 11).Value = IIf(IsNull(RS2!total_data), "", RS2!total_data)
                    totalcall = IIf(IsNull(RS2!total_call), "", RS2!total_call)
                    totaldata = IIf(IsNull(RS2!total_data), "", RS2!total_data)
                    ratarata = totalcall / totaldata
                .Cells(iRow, 12).Value = ratarata
                RS2.MoveNext
            Loop
        End If
        
        iRow = 4
        If rs.RecordCount > 0 Then
            ProgressBar1.Max = rs.RecordCount
            i = 0
            Do Until rs.EOF
                i = i + 1
                iRow = iRow + 1
                ProgressBar1.Value = rs.Bookmark
                .Cells(iRow, 1).Value = i
                .Cells(iRow, 2).Value = IIf(IsNull(rs!CustId), "", rs!CustId)
                .Cells(iRow, 3).Value = IIf(IsNull(rs("tgl_call")), "", Format(rs("tgl_call"), "yyyy-mm-dd"))
                .Cells(iRow, 4).Value = IIf(IsNull(rs!JUMLAH), "", rs!JUMLAH)
                .Cells(iRow, 5).Value = IIf(IsNull(rs!f_cek_new), "", rs!f_cek_new)
                .Cells(iRow, 6).Value = IIf(IsNull(rs!Descol), "", rs!Descol)
                
                rs.MoveNext
            Loop
        End If
    
        'OTOMATISASI CELL
        For iColom = 1 To 14
            ExlObj.Cells(4, iColom).EntireColumn.AutoFit
        Next
        
        MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
        ProgressBar1.Value = 0
        cmd_download.Enabled = True
    
        Set ExlObj = Nothing
        Set rs = Nothing

        'StartMeUp (Txtlocation.Text)
        'FILL COLOR CELL
        'ExlObj.Range(.Cells(NoUrut, 1), .Cells(NoUrut, 7)).Interior.Color = RGB(6, 207, 250)
    End With
End Sub

Private Sub cmd_exit_Click()
    Unload Me
End Sub

Private Sub cmd_proses_Click()
    Me.MousePointer = vbArrowHourglass
    Call Create_Table_Temp
    Call Create_Table_Temp_sum
    'Call export_data
    Me.MousePointer = vbArrow
End Sub

Private Sub Create_Table_Temp_sum()
    Dim cQuery As String
    Dim jam_mulai As String
    Dim jam_selesai As String
    Dim tgl_mulai As String
    Dim tgl_akhir As String
    Dim rs As ADODB.Recordset
    
    jam_mulai = "00:00:00"
    jam_selesai = "23:59:59"
    
    tgl_mulai = Format(tgl_mulai1.Value, "YYYY-MM-DD")
    tgl_akhir = Format(tgl_akhir1.Value, "YYYY-MM-DD")

    cQuery = " SELECT a.custid, sum(jumlah), f_cek_new, descol  FROM ( SELECT * FROM mandiri.tbl_temp_jumlah_call WHERE tgl_call BETWEEN '" & tgl_mulai & " " & jam_mulai & "' AND '" & tgl_akhir & " " & jam_selesai & "') As a LEFT JOIN  (SELECT custid, f_cek_new, b.agent as descol FROM mandiri.mgm a, mandiri.usertbl b  WHERE a.agent = b. userid) As b  on a.custid = b.custid group by 1,3,4"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    lv_report_sum.ListItems.CLEAR
    If rs.RecordCount > 0 Then
        ProgressBar1.Max = rs.RecordCount
        While Not rs.EOF
            ProgressBar1.Value = rs.Bookmark
            Set LS = lv_report_sum.ListItems.ADD(, , cnull(Trim(rs("descol"))))
                LS.SubItems(1) = Trim(rs("custid"))
                LS.SubItems(2) = Trim(rs("sum"))
                LS.SubItems(3) = cnull(Trim(rs("f_cek_new")))
        rs.MoveNext

        Wend
        Warna_Row_Listview Form_report_temp, lv_report_sum, &HFFFF80, vbWhite
    Else
        MsgBox "Data tidak ditemukan!!", vbOKOnly + vbInformation, "INFO"
    End If

    txtsum.Text = Format(rs.RecordCount, "##,###")
    
End Sub

Private Sub Create_Table_Temp()
    Dim cQuery As String
    Dim jam_mulai As String
    Dim jam_selesai As String
    Dim tgl_mulai As String
    Dim tgl_akhir As String
    Dim rs As ADODB.Recordset
    
    jam_mulai = "00:00:00"
    jam_selesai = "23:59:59"
    
    tgl_mulai = Format(tgl_mulai1.Value, "YYYY-MM-DD")
    tgl_akhir = Format(tgl_akhir1.Value, "YYYY-MM-DD")
    

        
    'cQuery = " CREATE TABLE tbl_temp_jumlah_call AS"
'    cQuery = "SELECT a.custid, tgl, temp, f_cek_new FROM ( "
'    cQuery = cQuery + " SELECT custid, tgl, count(stop_time) as temp"
'    cQuery = cQuery + " FROM mgm_hst"
'    cQuery = cQuery + " WHERE tgl BETWEEN '" & tgl_mulai & " " & jam_mulai & "' AND '" & tgl_akhir & " " & jam_selesai & "' "
'    cQuery = cQuery + " AND custid IN (SELECT custid FROM mgm) GROUP BY custid, tgl ) AS a  "
'    cQuery = cQuery + " INNER JOIN (SELECT max(tglcall), custid, f_cek_new FROM mgm"
'    cQuery = cQuery + " GROUP BY custid, f_cek_new) As b"
'    cQuery = cQuery + " ON a.custid = b.custid"
'    cQuery = cQuery + " WHERE temp > 0 order by tgl"
    cQuery = "SELECT a.custid, tgl_call, jumlah, f_cek_new, descol  FROM ("
    cQuery = cQuery + " SELECT * FROM mandiri.tbl_temp_jumlah_call WHERE tgl_call"
    cQuery = cQuery + " BETWEEN '" & tgl_mulai & " " & jam_mulai & "' AND '" & tgl_akhir & " " & jam_selesai & "') As a INNER JOIN "
    cQuery = cQuery + " (SELECT custid, f_cek_new, b.agent as descol FROM mandiri.mgm a, mandiri.usertbl b "
    cQuery = cQuery + " WHERE a.agent = b. userid ) As b "
    cQuery = cQuery + " on a.custid = b.custid"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    lv_report_temp.ListItems.CLEAR
    If rs.RecordCount > 0 Then
        ProgressBar1.Max = rs.RecordCount
        While Not rs.EOF
            ProgressBar1.Value = rs.Bookmark
            Set LS = lv_report_temp.ListItems.ADD(, , cnull(Trim(rs("descol"))))
                LS.SubItems(1) = Trim(rs("custid"))
                LS.SubItems(2) = Trim(Format(rs("tgl_call"), "DD-MM-YYYY"))
                LS.SubItems(3) = Trim(rs("jumlah"))
                LS.SubItems(4) = cnull(Trim(rs("f_cek_new")))
        rs.MoveNext

        Wend
        Warna_Row_Listview Form_report_temp, lv_report_temp, &HFFFF80, vbWhite
    Else
        MsgBox "Data tidak ditemukan!!", vbOKOnly + vbInformation, "INFO"
    End If

    txt_jml.Text = Format(rs.RecordCount, "##,###")
    
End Sub

Private Sub export_data()
    exportdata " SELECT * FROM mandiri.tbl_temp_jumlah_call ", ""
End Sub

Public Sub exportdata(mwheere As String, XUPDATE As String)
    Dim CMDSQL As String
    Dim M_objrs As New ADODB.Recordset
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
    M_objrs.Open mwheere, obj_koneksi, adOpenDynamic, adLockOptimistic, adCmdText
    
    'Cek jumlah data
    
    'Jika data tidak ada, maka keluar dari fungsi ini!
    
    If MsgBox("Apakah Anda Yakin Ingin Men-Download Report", vbQuestion + vbYesNo, "Question") = vbYes Then
Form_Save:

    CommonDialog1.ShowSave
    TxtPath.Text = CommonDialog1.FileName
    
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
    
    'Mengambil waktu dari server
   
    
    'Set excel
    Set objExcel = CreateObject("Excel.Application")
    Set objBook = objExcel.Workbooks.ADD
    Set objSheet = objBook.ActiveSheet
        
    
    
    
    
    On Error GoTo Salah
    'Proses pengsisian nama field ke excel
    Dim X, Y    As Integer
        If M_objrs.state = 1 Then
            X = 0
            Y = M_objrs.fields().Count - 1
            Do Until X > Y
                DoEvents
                objSheet.Cells(1, i).Value = CStr(M_objrs.fields(X).Name)
                i = i + 1
                X = X + 1
            Loop
        End If
    
    
    objSheet.Range("A2").CopyFromRecordset M_objrs '-> Proses pengisian data dimulai dari Cell A2
    objBook.SaveAs TxtPath.Text, xlWorkbookNormal
    objExcel.Quit
    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    objExcel.Workbooks.Open TxtPath.Text & ".xls"
    objExcel.Visible = True
    Set objExcel = Nothing: Set objBook = Nothing: Set objSheet = Nothing
    TxtPath.Text = ""
    Set M_objrs = Nothing
    Exit Sub
End If
Salah:
'MsgBox Err.Description
    Exit Sub
    
End Sub



Private Sub Command1_Click()
        Dim objExcel As New Excel.Application
        Dim objExcelSheet As Excel.Worksheet
        Dim col, row As Integer
        Dim a As String
        If lv_report_sum.ListItems.Count > 0 Then
            objExcel.Workbooks.ADD
            Set objExcelSheet = objExcel.Worksheets.ADD
         
        
            For col = 1 To lv_report_sum.ColumnHeaders.Count
                objExcelSheet.Cells(1, col).Value = lv_report_sum.ColumnHeaders(col)
            Next
         
            For row = 2 To lv_report_sum.ListItems.Count + 1
                For col = 1 To lv_report_sum.ColumnHeaders.Count
                If col = 1 Then
                        objExcelSheet.Cells(row, col).Value = lv_report_sum.ListItems(row - 1).Text
                Else
                    '" 'cararandy 29032016 "
                    Dim hasil1 As String
                        hasil1 = "'" + lv_report_sum.ListItems(row - 1).SubItems(col - 1)
                        objExcelSheet.Cells(row, col).Value = hasil1
                    End If
                Next
            Next
         
            objExcelSheet.Columns.AutoFit
            Cd_save.ShowOpen
            a = Cd_save.FileName
         
            objExcelSheet.SaveAs a & ".xls"
            MsgBox "Export Completed", vbInformation, Me.Caption
         
            objExcel.Workbooks.Open a & ".xls"
            objExcel.Visible = True
        Else
            MsgBox "No data to export", vbInformation, Me.Caption
        End If
End Sub

Private Sub Form_Load()
    tgl_mulai1.Value = Now
    tgl_akhir1.Value = Now
    
    Call Header_Report_Temp
    Call Header_Report_Temp_sum
End Sub

Private Sub Header_Report_Temp()
    lv_report_temp.ColumnHeaders.ADD 1, , "Agent", 2500
    lv_report_temp.ColumnHeaders.ADD 2, , "CustID", 2500
    lv_report_temp.ColumnHeaders.ADD 3, , "Tanggal", 1700
    lv_report_temp.ColumnHeaders.ADD 4, , "Jumlah Temp", 3500
    lv_report_temp.ColumnHeaders.ADD 5, , "Status Call", 1700
End Sub

Private Sub Header_Report_Temp_sum()
    lv_report_sum.ColumnHeaders.ADD 1, , "Agent", 2500
    lv_report_sum.ColumnHeaders.ADD 2, , "CustID", 2500
    lv_report_sum.ColumnHeaders.ADD 3, , "Jumlah Temp", 3500
    lv_report_sum.ColumnHeaders.ADD 4, , "Status Call", 1700
End Sub

