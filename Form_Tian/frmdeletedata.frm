VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmdeletedata 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PullOut Data"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11730
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   11730
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Choose File"
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000E&
         Caption         =   "Cara Pakai"
         Height          =   2175
         Left            =   7920
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   3735
         Begin VB.Label Label7 
            BackColor       =   &H8000000E&
            Caption         =   "_"
            BeginProperty Font 
               Name            =   "Showcard Gothic"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3120
            TabIndex        =   12
            Top             =   120
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   $"frmdeletedata.frx":0000
            Height          =   1815
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Delete History"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Export"
         Height          =   255
         Left            =   1200
         TabIndex        =   16
         Top             =   1920
         Width           =   735
      End
      Begin VB.Timer Timer1 
         Interval        =   600
         Left            =   11160
         Top             =   360
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Check"
         Height          =   495
         Left            =   8760
         TabIndex        =   8
         Top             =   1680
         Width           =   1215
      End
      Begin VB.ComboBox cbosheet 
         Height          =   315
         Left            =   1350
         TabIndex        =   4
         Top             =   990
         Width           =   2565
      End
      Begin VB.CommandButton cmdbrowse 
         BackColor       =   &H00C0FFC0&
         Caption         =   "...."
         Height          =   315
         Left            =   9840
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   585
         Width           =   555
      End
      Begin VB.TextBox txtlocation 
         Height          =   315
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   8445
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   495
         Left            =   10200
         TabIndex        =   1
         Top             =   1680
         Width           =   1335
      End
      Begin MSComDlg.CommonDialog CD_save 
         Left            =   0
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label3 
         Caption         =   "Show History"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         Caption         =   "Cara Pakai"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10320
         TabIndex        =   9
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Sheet"
         Height          =   255
         Left            =   150
         TabIndex        =   7
         Top             =   1020
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "Location"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Choose File"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Height          =   5055
      Left            =   0
      TabIndex        =   14
      Top             =   2280
      Width           =   11775
      Begin VB.CommandButton Command4 
         Caption         =   "Search"
         Height          =   255
         Left            =   2040
         TabIndex        =   22
         Top             =   960
         Width           =   735
      End
      Begin MSComctlLib.ListView LvPTP 
         Height          =   4620
         Left            =   2880
         TabIndex        =   15
         Top             =   240
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   8149
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
         Appearance      =   1
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
      Begin TDBDate6Ctl.TDBDate dtpropsal 
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   450
         Calendar        =   "frmdeletedata.frx":0156
         Caption         =   "frmdeletedata.frx":026E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmdeletedata.frx":02DA
         Keys            =   "frmdeletedata.frx":02F8
         Spin            =   "frmdeletedata.frx":0356
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   0
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
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   3.54028054673894E-316
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate TDBDate1 
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Top             =   600
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   450
         Calendar        =   "frmdeletedata.frx":037E
         Caption         =   "frmdeletedata.frx":0496
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmdeletedata.frx":0502
         Keys            =   "frmdeletedata.frx":0520
         Spin            =   "frmdeletedata.frx":057E
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   0
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
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   3.54028054673894E-316
         CenturyMode     =   0
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2940
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   5186
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
         Appearance      =   1
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
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Found : "
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         Height          =   255
         Left            =   1440
         TabIndex        =   20
         Top             =   720
         Width           =   135
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Perkiraan Tanggal"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmdeletedata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, B, c As Integer
Public M_XLSCONN As New ADODB.Connection

Private Sub cbosheet_Click()
    If txtlocation.text <> "" Then
        If M_XLSCONN.State = adStateOpen Then M_XLSCONN.Close
        M_XLSCONN.Open "Provider = Microsoft.Jet.OleDb.4.0;data source = " & CommonDialog1.FileName & ";Extended Properties=Excel 8.0;"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        ssql = "SELECT * FROM [" & cbosheet.text & "] "
        M_objrs.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
        Set M_objrs = Nothing
    End If
End Sub

Private Sub cmdbrowse_Click()
    With CommonDialog1
        .DialogTitle = "Import From File"
        '.Filter = "Excel Files|*.xls;*.xlsx"
        .Filter = "Excel Files|*.xls"
        .ShowOpen
    End With
    txtlocation.text = ""
    If CommonDialog1.FileName = "" Then Exit Sub
    txtlocation.text = CommonDialog1.FileName
    If M_XLSCONN.State = adStateOpen Then M_XLSCONN.Close
    M_XLSCONN.Open "Provider = Microsoft.Jet.OleDb.4.0;data source = " & CommonDialog1.FileName & ";Extended Properties=Excel 8.0;"
    Set M_objrs = M_XLSCONN.OpenSchema(adSchemaTables)
    cbosheet.CLEAR
    If M_objrs.EOF And M_objrs.BOF Then Exit Sub
    While Not M_objrs.EOF
        cbosheet.AddItem IIf(IsNull(M_objrs!table_name), "", M_objrs!table_name)
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
    Set M_XLSCONN = Nothing

End Sub

Private Sub Command1_Click()

    qs = "select * from information_schema.columns where table_name = 'tbllogpullout'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If M_objrs.RecordCount = 0 Then
        qc = "Create table tbllogpullout (tanggal timestamp without time zone, jml int, tabel varchar, eksekusiby varchar);"
        M_OBJCONN.Execute qc
    End If
    
    qs = "select to_char(now(),'yyyymmddhhmiss') as tanggal"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic
    zzz = M_objrs!Tanggal
    
    qi = "Insert into tbllogpullout values (now()," & c & "," & "'backuppullout_" & zzz & "', '" & MDIForm1.TxtUsername.text & "');"
    M_OBJCONN.Execute qi
    
    qc = "Create table backuppullout_" & zzz & " as select * from mgm where custid in (select uniqnya from tbltemp_uniqdelete)"
    M_OBJCONN.Execute qc
    
    qd = "delete from mgm where custid in (select uniqnya from tbltemp_uniqdelete)"
    M_OBJCONN.Execute qd
    
    If Check1.Value = 1 Then
        qd = "delete from mgm_hst where custid in (select uniqnya from tbltemp_uniqdelete)"
        M_OBJCONN.Execute qd
    End If
    
    MsgBox "Data Berhasil di Delete sebanyak : " & c & " Data"
    
End Sub

Private Sub Command2_Click()
    Dim str_sql As String
    
    qs = "select * from information_schema.columns where table_name = 'tbltemp_uniqdelete'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic
        
    If M_objrs.RecordCount = 0 Then
        qc = "create table tbltemp_uniqdelete ( id serial, uniqnya varchar );"
        M_OBJCONN.Execute qc
    End If
        qd = "delete from tbltemp_uniqdelete;"
        M_OBJCONN.Execute qd
        
        ssql = "SELECT * FROM [" & cbosheet.text & "]   "
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
            
        Set rsTemporary = New ADODB.Recordset
        rsTemporary.CursorLocation = adUseClient
        rsTemporary.CursorType = adOpenDynamic
        rsTemporary.ActiveConnection = M_OBJCONN
        rsTemporary.LockType = adLockOptimistic
            
        rs.Open ssql, M_XLSCONN, adOpenKeyset, adLockOptimistic
    
        While Not rs.EOF
        
            str_sql = "INSERT INTO tbltemp_uniqdelete (uniqnya) Values ( '" + rs(0) + "' " + ");"
            M_OBJCONN.Execute str_sql
            
            rs.MoveNext
        Wend
            
        qs = "select * from tbltemp_uniqdelete"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic
        
        a = rs.RecordCount
                
        qs = " select distinct * from tbltemp_uniqdelete"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic
        B = rs.RecordCount
        
        If a > B Then
            MsgBox "Pastikan Uniq tidak ada yang double"
            Exit Sub
        End If
        
        qs = "select custid from mgm where custid in (select uniqnya from tbltemp_uniqdelete)"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic
        c = rs.RecordCount
        
        Dim teks As String
        teks = "Jumlah Data Excel : " & a & vbCrLf & "Jumlah Data Didatabase setelah dicheck : " & c & vbCrLf
        
        If a <> c Then
            teks = teks & "Status : Tidak Sesuai"
        Else
            teks = teks & "Status : Sesuai"
        End If
        MsgBox teks
        Command1.Enabled = True
End Sub

Private Sub Command3_Click()
    Dim objExcel As New Excel.Application
    Dim objExcelSheet As Excel.Worksheet
    Dim col, row As Integer
    Dim a As String
    If LvPTP.ListItems.Count > 0 Then
        objExcel.Workbooks.ADD
        Set objExcelSheet = objExcel.Worksheets.ADD
     
    
        For col = 1 To LvPTP.ColumnHeaders.Count
            objExcelSheet.Cells(1, col).Value = LvPTP.ColumnHeaders(col)
        Next
     
        For row = 2 To LvPTP.ListItems.Count + 1
            For col = 1 To LvPTP.ColumnHeaders.Count
            If col = 1 Then
                    objExcelSheet.Cells(row, col).Value = LvPTP.ListItems(row - 1).text
            Else
                '" 'cararandy 29032016 "
                Dim hasil1 As String
                    hasil1 = "'" + LvPTP.ListItems(row - 1).SubItems(col - 1)
                    objExcelSheet.Cells(row, col).Value = hasil1
                End If
            Next
        Next
     
        objExcelSheet.Columns.AutoFit
        CD_save.ShowOpen
        a = CD_save.FileName
     
        'objExcelSheet.SaveAs a & ".xlsx"
        objExcelSheet.SaveAs a & ".xls"
        MsgBox "Export Completed", vbInformation, Me.Caption
     
        'objExcel.Workbooks.Open a & ".xlsx"
        objExcel.Workbooks.Open a & ".xls"
        objExcel.Visible = True
    Else
        MsgBox "No data to export", vbInformation, Me.Caption
    End If

End Sub

Private Sub Command4_Click()
    Dim abcde1, abcde2 As String
    abcde1 = Format(cnull(dtpropsal.Value), "yyyy-mm-dd")
    abcde2 = Format(cnull(TDBDate1.Value), "yyyy-mm-dd")
    
    If abcde1 = "" Or abcde2 = "" Then
        MsgBox "Harap pilih tanggal"
        Exit Sub
    End If
    
    qs = "select * from ("
    qs = qs & " select table_name, tanggal::timestamp without time zone from (" & vbCrLf
    qs = qs & " select table_name, thn||'-'||bln||'-'||tgl||' '||jm||':'||min as tanggal from (" & vbCrLf
    qs = qs & " select table_name, left(tgl,4) thn, substring(tgl from 5 for 2) as bln, substring(tgl from 7 for 2) as tgl, substring(tgl from 9 for 2) as jm, substring(tgl from 11 for 2) as min from (" & vbCrLf
    qs = qs & " select table_name, replace(table_name,'backuppullout_','') tgl from (" & vbCrLf
    qs = qs & " select distinct table_name from information_schema.columns  where table_name ilike '%backuppullout_%'" & vbCrLf
    qs = qs & " ) a" & vbCrLf
    qs = qs & " ) b" & vbCrLf
    qs = qs & " ) c" & vbCrLf
    qs = qs & " ) d" & vbCrLf
    qs = qs & " ) e where tanggal between '" & abcde1 & " 00:00:00' and '" & abcde2 & " 23:59:59' order by tanggal"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If M_objrs.RecordCount > 0 Then
        MsgBox "Data ditemukan"
    Else
        MsgBox "Data tidak ditemukan"
    End If
    
    ListView1.ListItems.CLEAR
    
    While Not M_objrs.EOF
        Set ListItem = ListView1.ListItems.ADD(, , cnull(M_objrs("table_name")))
            ListItem.SubItems(1) = cnull(M_objrs("tanggal"))
        M_objrs.MoveNext
    Wend
    
    Set M_objrs = Nothing
End Sub

Private Sub Form_Load()
    headerhst
End Sub

Private Sub search()
    tabel = ListView1.SelectedItem.text
    
    Field = "name,addrnow,homeno,mobileno,addrpt,officeno,nocard,region,dob,recsource,custid,curbal,pay_dt,lastpay,product_desc,batchdiskon,remarks_old,afaxno,delq_history"
        
    sStrsql = "select " & Field & " from " & tabel
            
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    LvPTP.ListItems.CLEAR
    While Not M_objrs.EOF
        Set ListItem = LvPTP.ListItems.ADD(, , cnull(M_objrs("name")))
            ListItem.SubItems(1) = cnull(M_objrs("AddrNow"))
            ListItem.SubItems(2) = cnull(M_objrs("homeno"))
            ListItem.SubItems(3) = cnull(M_objrs("mobileno"))
            ListItem.SubItems(4) = cnull(M_objrs("addrpt"))
            ListItem.SubItems(5) = cnull(M_objrs("officeno"))
            ListItem.SubItems(6) = cnull(M_objrs("nocard"))
            ListItem.SubItems(7) = cnull(M_objrs("region"))
            ListItem.SubItems(8) = cnull(M_objrs("DOB"))
            ListItem.SubItems(9) = cnull(M_objrs("recsource"))
            ListItem.SubItems(10) = cnull(M_objrs("CustId"))
            ListItem.SubItems(11) = cnull(M_objrs("curbal"))
            ListItem.SubItems(12) = cnull(M_objrs("pay_dt"))
            ListItem.SubItems(13) = cnull(M_objrs("lastpay"))
            ListItem.SubItems(14) = cnull(M_objrs("product_desc"))
            ListItem.SubItems(15) = cnull(M_objrs("batchdiskon"))
            ListItem.SubItems(16) = cnull(M_objrs("remarks_old"))
            ListItem.SubItems(17) = cnull(M_objrs("afaxno"))
            ListItem.SubItems(18) = cnull(M_objrs("delq_history"))
        M_objrs.MoveNext
    Wend
    
    Label10.Caption = "Found : " & M_objrs.RecordCount
    
    Set M_objrs = Nothing

End Sub

Private Sub headerhst()
    LvPTP.ColumnHeaders.CLEAR
    With LvPTP.ColumnHeaders
        .ADD 1, , "NAME"
        .ADD 2, , "ADDRESSNOW"
        .ADD 3, , "HOMEPHONE"
        .ADD 4, , "MOBILEPHONE"
        .ADD 5, , "ADDRESSOFFICE"
        .ADD 6, , "OFFICEPHONE"
        .ADD 7, , "CARDNO"
        .ADD 8, , "REGION"
        .ADD 9, , "DOB"
        .ADD 10, , "RECSOURCE"
        .ADD 11, , "CUSTID"
        .ADD 12, , "CURBAL"
        .ADD 13, , "PAYDATE"
        .ADD 14, , "LASTPAY"
        .ADD 15, , "ECDESC"
        .ADD 16, , "MIN DISKON"
        .ADD 17, , "REMARKSOLD"
        .ADD 18, , "ECPHONE"
        .ADD 19, , "APPLID"
    End With
    
    ListView1.ColumnHeaders.CLEAR
    With ListView1.ColumnHeaders
        .ADD 1, , "File", 0
        .ADD 2, , "TANGGAL", TXT * 100
    End With

End Sub

Private Sub Label3_Click()
    If frmdeletedata.Height <> 7785 Then
        frmdeletedata.Height = 7785
        Label3.Caption = "Hide History"
    Else
        frmdeletedata.Height = 2730
        Label3.Caption = "Show History"
    End If
End Sub

Private Sub Label5_Click()
    Frame2.Visible = True
End Sub

Private Sub Label7_Click()
    Frame2.Visible = False
End Sub

Private Sub ListView1_DblClick()
    If ListView1.ListItems.Count = 0 Then
        Exit Sub
    End If
    Call search
End Sub

Private Sub Timer1_Timer()
    If Label5.BackColor = &H8000000D Then
        Label5.BackColor = &H8000000F
        Label3.BackColor = &H8000000F
    Else
        Label5.BackColor = &H8000000D
        Label3.BackColor = &H8000000D
    End If
End Sub
