VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form_Report_Submit 
   Caption         =   "Result Report"
   ClientHeight    =   9390
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   17610
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   17610
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   8010
      Left            =   30
      TabIndex        =   2
      Top             =   1320
      Width           =   17520
      _ExtentX        =   30903
      _ExtentY        =   14129
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
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1185
      Left            =   15
      TabIndex        =   0
      Top             =   90
      Width           =   18885
      Begin Threed.SSCommand cmd_search_visit 
         Height          =   735
         Left            =   7800
         TabIndex        =   8
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1296
         _Version        =   196610
         Caption         =   "Search"
         ButtonStyle     =   2
      End
      Begin VB.TextBox TxtPath 
         Height          =   285
         Left            =   5640
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   360
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form_Report_Submit.frx":0000
         Left            =   1440
         List            =   "Form_Report_Submit.frx":0010
         TabIndex        =   4
         Top             =   175
         Width           =   3375
      End
      Begin VB.CommandButton cmdSettingColumn 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   17595
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Form_Report_Submit.frx":0075
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3060
         Width           =   405
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Preview"
         Height          =   375
         Left            =   9255
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1200
         Width           =   3705
      End
      Begin MSComDlg.CommonDialog Cd_save 
         Left            =   13800
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "Ms. Excel 97/2000/XP|*.xls"
      End
      Begin Crystal.CrystalReport RPT 
         Left            =   13800
         Top             =   600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin Threed.SSCommand sscommand1 
         Height          =   735
         Left            =   8880
         TabIndex        =   9
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1296
         _Version        =   196610
         Caption         =   "Export"
         ButtonStyle     =   2
      End
      Begin Threed.SSCommand cmdexit_visit 
         Height          =   735
         Left            =   9960
         TabIndex        =   7
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1296
         _Version        =   196610
         ForeColor       =   -2147483637
         BackColor       =   192
         Caption         =   "Exit"
         ButtonStyle     =   2
      End
      Begin VB.Label Label1 
         Caption         =   "Choose Report"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   4
      Left            =   60
      Picture         =   "Form_Report_Submit.frx":068F
      Stretch         =   -1  'True
      Top             =   60
      Width           =   420
   End
End
Attribute VB_Name = "Form_Report_Submit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function createQuery()
    Dim mwhere As String
    Dim basicField As String
    Dim LoginManager As Boolean
    Dim formulirField As String
    Dim sFieldManager, sFieldAgree As String
    
    If Combo1.Text = "TRACKING SUMMARY REPORT AGENT" Then
        STRSQL = "select row_number() over(),tsra.*, hst.initiated from ("
        STRSQL = STRSQL + " select agent as ""AGENT"","
        STRSQL = STRSQL + " count(custid) as ""DATASIZE"","
        STRSQL = STRSQL + " sum(amountwo) as ""JML VOL"","
        STRSQL = STRSQL + " sum(case when tglcall is not null then 1 end) as ""Data Utilized"","
        STRSQL = STRSQL + " sum(case when tglcall is not null then amountwo end) as ""Volume Utilized"","
        STRSQL = STRSQL + " (sum(case when tglcall is not null then 1 end)::numeric(3,0)) / (count(custid)::numeric(3,0)) * 100 as ""% Utilized"","
        STRSQL = STRSQL + " sum(case when f_cek_new = 'POP' then 1 else 0 end) as ""POP"","
        STRSQL = STRSQL + " sum(case when f_cek_new = 'SP-' then 1 else 0 end) as ""SP"","
        STRSQL = STRSQL + " sum(case when f_cek_new = 'BP-' then 1 else 0 end) as ""BP"","
        STRSQL = STRSQL + " sum(case when f_cek_new = 'PO-' then 1 else 0 end) as ""PTP PAIDOFF"","
        STRSQL = STRSQL + " sum(case when f_cek_new = 'PTP-NE' then 1 else 0 end) as ""PTP NEW"","
        STRSQL = STRSQL + " sum(case when f_cek_new = 'PTP-PO' then 1 else 0 end) as ""PTP POP"","
        STRSQL = STRSQL + " sum(case when f_cek_new = 'POP' or f_cek_new = 'SP-' or f_cek_new = 'BP-' or f_cek_new = 'PO-' or f_cek_new = 'PTP-NE'"
        STRSQL = STRSQL + " or f_cek_new = 'PTP-PO' then 1 else 0 end) as ""Total PTP"","
        STRSQL = STRSQL + " sum(case when f_cek_new = 'POP' or f_cek_new = 'SP-' or f_cek_new = 'BP-' or f_cek_new = 'PO-' or f_cek_new = 'PTP-NE'"
        STRSQL = STRSQL + " or f_cek_new = 'PTP-PO' then 1 else 0 end)::numeric(3,0)/sum(case when tglcall is not null then 1 end)::numeric(3,0)*100 as ""% PTP"","
        STRSQL = STRSQL + " sum(case when statuscall = 'VALID' then 1 else 0 end) as ""VALID"","
        STRSQL = STRSQL + " sum(case when statuscall = 'SKIP' then 1 else 0 end) as ""SKIP"","
        STRSQL = STRSQL + " sum(case when statuscall = 'Prospect' then 1 else 0 end) as ""PROSPECT"","
        STRSQL = STRSQL + " sum(case when statuscall = 'On Nego' then 1 else 0 end) as ""ON NEGO"","
        STRSQL = STRSQL + " sum(case when statuscall = 'On Process' then 1 else 0 end) as ""ON PROCESS"","
        STRSQL = STRSQL + " '' as ""RESULT POP"", '' as ""RESULT SP"", '' as ""RESULT VL"", '' as ""RESULT SK"",'' as ""RESULT PR"", '' as ""RESULT ON"", '' as ""RESULT PO"", '' as ""RESULT PTP PO"""
        STRSQL = STRSQL + " from mgm group by 1 ) tsra inner join (select agent, count(agent) as initiated from mgm_hst group by 1) as hst on tsra.""AGENT"" = hst.agent "
    ElseIf Combo1.Text = "REPORT PAYMENT NEW" Then
        STRSQL = "select row_number() over(),mgm.agent,mgm.tglptpnew, dateptp.prd, mgm.name, mgm.custid, mgm.region, '' as dbbulan, mgm.principal, mgm.amountwo, dateptp.promisepay from mgm inner join (select custid, max(promisedate) prd, promisepay from tblnegoptp group by 1,3) dateptp on mgm.custid = dateptp.custid where mgm.tglptpnew is not null"
    ElseIf Combo1.Text = "REPORT PTP JATUH TEMPO" Then
        STRSQL = "select row_number() over(),dateptp.prd, mgm.name, mgm.custid, mgm.region, mgm.agent, mgm.amountwo, dateptp.promisepay, mgm.ptpvia, mgm.tglcall, mgm.result_ptp from mgm inner join (select custid, max(promisedate) prd, promisepay from tblnegoptp group by 1,3) dateptp on mgm.custid = dateptp.custid where dateptp.prd is not null"
    ElseIf Combo1.Text = "Outbound Call Report" Then
        STRSQL = "select to_char(tglcall, 'YYYY-MM-DD') as calldate, to_char(tglcall, 'HH24:mm:ss') as calltime, custid, name, curbal, amountwo, region, statuscall,remarks,to_char(tglsource, 'Month'), agent from mgm"
    End If
    createQuery = STRSQL
    
End Function

Public Sub cariData()
    Dim STRSQL  As String
    Dim objVISIT As New ADODB.Recordset
    On Error GoTo ER
        
    Set objVISIT = New ADODB.Recordset
    objVISIT.CursorLocation = adUseClient
        
    STRSQL = createQuery '<<<----------- CREATE QUERY
    objVISIT.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
    'Text4.Text = objVISIT.RecordCount
    If objVISIT.RecordCount = 0 Then
        MsgBox "Data not found", vbInformation + vbOKOnly, "TINS"
        Exit Sub
    End If
       
    Set DataGrid1.DATASOURCE = objVISIT
    Set objVISIT = Nothing
    mwhere = ""
    
    Exit Sub
    
ER:
    MsgBox "Sorry, TINS Error: " + err.Description, vbCritical + vbOKOnly, "TINS"
    'cmbSaveColumn.Text = Empty
End Sub



Private Sub cmd_search_visit_Click()
    If Combo1.Text = "" Then
        MsgBox "Choose The Report"
        Exit Sub
    End If
    TxtPath.Text = Combo1.Text
    DataGrid1.Refresh
    cariData
End Sub

Private Sub cmdexit_visit_1_Click()

End Sub

Private Sub cmdexit_visit_Click()
    Unload Me
End Sub

Public Sub PRIVIEWDATA()
Dim STRSQL  As String
Dim m_objrs2 As New ADODB.Recordset
   
STRSQL = createQuery
    
m_objrs2.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
jml = m_objrs2.RecordCount

M_RPTCONN.Execute "delete from tblvoluntry_detail "
While Not m_objrs2.EOF
    If Combo2.Text = "Formulir Pembukaan Rekening (anak ke orang tua)" Then
        Call SAVE_FGM(m_objrs2)
    ElseIf Combo2.Text = "Formulir PermataProteksi MasaDepan" Then
        Call SAVE_PPMD(m_objrs2)
    ElseIf Combo2.Text = "Formulir PermataTabungan Berhadiah" Then
        Call SAVE_PermataBebas(m_objrs2)
    ElseIf Combo2.Text = "Konversi Tabungan Permata" Then
        Call SAVE_CONVERT(m_objrs2)
    End If
    m_objrs2.MoveNext
Wend

If Combo2.Text = "Formulir Pembukaan Rekening (anak ke orang tua)" Then
   RPT.ReportFileName = "D:\REPORT_COLLECTION_PERMATA\Rptpembukaan_tambahan.rpt"
ElseIf Combo2.Text = "Formulir PermataProteksi MasaDepan" Then
   RPT.ReportFileName = "D:\REPORT_COLLECTION_PERMATA\RptPPMD.rpt"
ElseIf Combo2.Text = "Formulir PermataTabungan Berhadiah" Then
   RPT.ReportFileName = "D:\REPORT_COLLECTION_PERMATA\RptPermataBebas.rpt"
ElseIf Combo2.Text = "Konversi Tabungan Permata" Then
   RPT.ReportFileName = "D:\REPORT_COLLECTION_PERMATA\RptCONVERT.rpt"

ElseIf Combo2.Text = Empty Then
    MsgBox "Pilihan Formulir harus diisi", vbCritical + vbOKOnly, "TINS"
    Exit Sub
End If
        
WaitSecs (2)
Call SHOW_PRN
Set m_objrs2 = Nothing
        
Set objVISIT = Nothing
End Sub



Private Sub isi_dataSTATUS(STRSQL As String)
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
    M_objrs.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic

    'Jika data tidak ada, maka keluar dari fungsi ini!
    If M_objrs.RecordCount = 0 Then
        MsgBox "Data Blank!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
   
form_save:
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
          GoTo form_save        '-> maka goto form_save
        End If
    End If
    
    'Set excel
    Set objExcel = CreateObject("Excel.Application")
    Set objBook = objExcel.Workbooks.ADD
    Set objSheet = objBook.ActiveSheet
            
    On Error GoTo SALAH
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
    
   ' lblstatus.Caption = "Status download: Membuat file excel... silahkan tunggu!"
    objSheet.Range("A2").CopyFromRecordset M_objrs '-> Proses pengisian data dimulai dari Cell A2
    objBook.SaveAs TxtPath.Text, xlWorkbookNormal
    objExcel.Quit
    Set objExcel = Nothing: Set objBook = Nothing: Set objSheet = Nothing
    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    Set M_objrs = Nothing
 
SALAH:
    Exit Sub

End Sub






Private Sub sscommand1_Click()
    Dim strQuery As String
    'SSCommand1.Enabled = False
    strQuery = createQuery
    isi_dataSTATUS strQuery
    sscommand1.Enabled = True
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


Private Function getFieldFromColumn() As String
    Dim rs As ADODB.Recordset
    Dim CMDSQL As String
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
        
    CMDSQL = "SELECT ARRAY_TO_STRING(ARRAY(SELECT column_name from tblcolumn_report where save_name = '" + cmbSaveColumn.Text + "'), ',') AS SQLSTMT"
    rs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

    getFieldFromColumn = Empty
    If Not rs.EOF Then
        getFieldFromColumn = cnull(rs(0))
    End If
    
    Set rs = Nothing
'getFieldFromColumn
End Function

Private Sub cmdSaveColumn_Click()
    If cmbSaveColumn.Text = Empty Or lvw2.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    saveColumnSetting
    
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub SSCommand2_Click()

End Sub

Private Sub SSCommand_1_Click()

End Sub

Private Sub sscommand1_1_Click()

End Sub
