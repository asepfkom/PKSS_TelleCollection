VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form form_hardcopy 
   Caption         =   "Form HardCopy"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14025
   LinkTopic       =   "Form2"
   ScaleHeight     =   8700
   ScaleWidth      =   14025
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8940
      Left            =   -15
      TabIndex        =   0
      Top             =   -15
      Width           =   14085
      Begin VB.TextBox txt_jmldata 
         Enabled         =   0   'False
         Height          =   345
         Left            =   7170
         TabIndex        =   16
         Top             =   8145
         Width           =   1305
      End
      Begin VB.CommandButton cmd_check_data 
         BackColor       =   &H0080FFFF&
         Caption         =   "Print && Check Data"
         Height          =   705
         Left            =   10620
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   7875
         Width           =   990
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000005&
         Caption         =   "Frame2"
         Height          =   255
         Left            =   1320
         TabIndex        =   10
         Top             =   7785
         Visible         =   0   'False
         Width           =   825
         Begin VB.TextBox TxtPath 
            Height          =   285
            Left            =   2985
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   240
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.TextBox txt_pass 
            Height          =   315
            Left            =   1620
            TabIndex        =   11
            Top             =   465
            Visible         =   0   'False
            Width           =   1560
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   255
            Left            =   2985
            TabIndex        =   13
            Top             =   525
            Visible         =   0   'False
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   450
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
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Set Password"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00333333&
            Height          =   345
            Index           =   1
            Left            =   390
            TabIndex        =   14
            Top             =   510
            Visible         =   0   'False
            Width           =   1335
         End
      End
      Begin TDBDate6Ctl.TDBDate tgl_upload 
         Height          =   345
         Left            =   4890
         TabIndex        =   8
         Top             =   8160
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   609
         Calendar        =   "form_hardcopy.frx":0000
         Caption         =   "form_hardcopy.frx":0118
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "form_hardcopy.frx":0184
         Keys            =   "form_hardcopy.frx":01A2
         Spin            =   "form_hardcopy.frx":0200
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "mm/dd/yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "mm/dd/yyyy"
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
         Value           =   43731
         CenturyMode     =   0
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000005&
         Caption         =   "Check All"
         Height          =   330
         Left            =   90
         TabIndex        =   6
         Top             =   7755
         Width           =   1350
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FF80&
         Caption         =   "Print Hard Copy && Excel"
         Height          =   345
         Left            =   11700
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   8235
         Width           =   1950
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF80&
         Caption         =   "Search"
         Height          =   345
         Left            =   11700
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   7860
         UseMaskColor    =   -1  'True
         Width           =   1950
      End
      Begin VB.ComboBox Combo7 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "form_hardcopy.frx":0228
         Left            =   1305
         List            =   "form_hardcopy.frx":0235
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   8160
         Width           =   2475
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   6825
         Left            =   90
         TabIndex        =   1
         Top             =   900
         Width           =   13875
         _ExtentX        =   24474
         _ExtentY        =   12039
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
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
      Begin Crystal.CrystalReport RPT 
         Left            =   10125
         Top             =   7815
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSComDlg.CommonDialog Cd_save 
         Left            =   10095
         Top             =   8205
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "Ms. Excel 97/2000/XP|*.xls"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Data"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00333333&
         Height          =   345
         Index           =   2
         Left            =   6060
         TabIndex        =   17
         Top             =   8190
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Report Verifikasi && Hardcopy"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Index           =   0
         Left            =   2010
         TabIndex        =   9
         Top             =   270
         Width           =   3585
      End
      Begin VB.Image Image2 
         Height          =   825
         Left            =   0
         Picture         =   "form_hardcopy.frx":0249
         Stretch         =   -1  'True
         Top             =   0
         Width           =   19245
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Upload"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00333333&
         Height          =   345
         Index           =   0
         Left            =   3855
         TabIndex        =   7
         Top             =   8205
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status Data"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00333333&
         Height          =   345
         Index           =   19
         Left            =   -165
         TabIndex        =   3
         Top             =   8190
         Width           =   1335
      End
   End
End
Attribute VB_Name = "form_hardcopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim get_noapp As String
Dim where_get_noapp As String
Dim where_get_noapp_vw As String

Private Sub Check1_Click()
    If ListView2.ListItems.Count > 0 Then
        i = ListView2.ListItems.Count
        If Check1.Value = vbChecked Then
            For a = 1 To i
                ListView2.ListItems(a).Checked = True
            Next a
        Else
            For a = 1 To i
                ListView2.ListItems(a).Checked = False
            Next a
        End If
    Else
        If Check1.Value = 1 Then
            MsgBox "tidak ada data"
            Check1.Value = vbUnchecked
            Exit Sub
        End If
    End If
End Sub

Private Sub cmd_check_data_Click()
    Call getapp
    
    query = "delete from tbl_report;" & vbCrLf
    query = query & "insert into tbl_report select * from vw_report_last6 where 1=1 " & where_get_noapp & ";"
    M_OBJCONN.Execute query
    
    'Excel
    Dim strQuery As String
    strQuery = createQuery
    isi_dataSTATUS strQuery
    '----------
End Sub

Private Sub Command1_Click()
    Call Last6TT
    Call Last6Kan
    Call Last6Econ
    Call Last6App
    Dim M_M_objrs As New ADODB.Recordset
    
    Dim c1 As Integer
    Dim buka As Integer

    sStrsql = "select column_name from information_schema.columns  where table_name = 'tbl_header_mgm_show' order by ordinal_position"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic

    ListView2.ListItems.CLEAR
    ListView2.ColumnHeaders.CLEAR
    If M_objrs.RecordCount > 0 Then
        c1 = M_objrs.RecordCount
        With ListView2.ColumnHeaders
            For i = 1 To c1
                .ADD i, , cnull(M_objrs!column_name)
                M_objrs.MoveNext
            Next i
        End With
    End If
    Set M_objrs = Nothing

    wherebuka = " and f_open = 0 "
    
    If tgl_upload.Value <> "" Then
        datewhere = " and tanggal_upload = '" & Format(tgl_upload.Value, "yyyy-mm-dd") & "'"
    End If
    
    mwhere = ""
    If Combo7.text = "OK" Then
        mwhere = mwhere & " and coalesce(status_data,'') <> '' "
    ElseIf Combo7.text = "TIDAK OK" Then
        mwhere = mwhere & " and coalesce(status_data,'') = '' "
    End If


    sStrsql = "select * from mgm_show_list_report where 1=1 " & mwhere & wherebuka & datewhere
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic

    While Not M_objrs.EOF

        Set ListItem = ListView2.ListItems.ADD(, , cnull(M_objrs("id_mgm")))
            For i = 1 To c1 - 1
                ListItem.SubItems(i) = cnull(M_objrs(i))
            Next i
        M_objrs.MoveNext
    Wend
    
    Set M_objrs = Nothing
    
    txt_jmldata.text = ListView2.ListItems.Count
    'Excel
    TxtPath.text = "Report Detail PKSS"
    DataGrid1.Refresh
    cariData
    '----------------
    
End Sub

Private Sub Command2_Click()
    Call getapp

    query = "delete from tbl_report;" & vbCrLf
    query = query & "insert into tbl_report select * from vw_report_last6 where 1=1 " & where_get_noapp & ";"
    M_OBJCONN.Execute query
    
    
    'WaitSecs (2)
    RPT.ReportFileName = "D:\PhoneVerif\rpt_verification.rpt"
    'RPT.ReportFileName = "D:\KANTOR\desktop development\PKSS\Additional\rpt_verification - Copy.rpt"
    'RPT.ReportFileName = "E:\rpt_verification.rpt"
    Call SHOW_PRN
    
    'Excel
    Dim strQuery As String
    strQuery = createQuery
    isi_dataSTATUS strQuery
    '----------
    
    M_OBJCONN.Execute "update mgm_verification_submit set f_open=1 where 1=1 " & where_get_noapp
    M_OBJCONN.Execute "update mgm_verification set f_open=1 where 1=1 " & where_get_noapp
    
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

Private Sub getapp()
    get_noapp = ""
    where_get_noapp = ""
    where_get_noapp_vw = ""
    For i = 1 To ListView2.ListItems.Count
        If ListView2.ListItems(i).Checked Then
            get_noapp = get_noapp & "'" & ListView2.ListItems(i).SubItems(4) & "',"
        End If
    Next i
    
    If get_noapp <> "" Then
        get_noapp = Left(get_noapp, Len(get_noapp) - 1)
        where_get_noapp = " and noapp in (" & get_noapp & ")"
        where_get_noapp_vw = " and ""No.Aplikasi"" in (" & get_noapp & ")"
    End If
End Sub

Public Sub cariData()
    Dim strsql  As String
    Dim objVISIT As New ADODB.Recordset
    On Error GoTo ER
        
    Set objVISIT = New ADODB.Recordset
    objVISIT.CursorLocation = adUseClient
        
    strsql = createQuery '<<<----------- CREATE QUERY
    objVISIT.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    If objVISIT.RecordCount = 0 Then
        'MsgBox "Data not found", vbInformation + vbOKOnly, "TINS"
        Exit Sub
    End If
       
    Set DataGrid1.DATASOURCE = objVISIT
    Set objVISIT = Nothing
    mwhere = ""
    
    Exit Sub
    
ER:
    MsgBox "Sorry, TINS Error: " + err.Description, vbCritical + vbOKOnly, "TINS"
End Sub

Public Function createQuery()
    Call getapp
    
    strsql = "select * from vw_tbl_report where 1=1 " & where_get_noapp_vw
    createQuery = strsql
End Function

Private Sub isi_dataSTATUS(strsql As String)
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim ListItem As ListItem
    Dim cmdsql_update As String
    Dim objExcel        As Excel.Application
    Dim objBook         As Excel.Workbook
    Dim objSheet        As Excel.Worksheet
    Dim i As Integer
    Dim m_msgbox As String
    Dim passwords As String
    Dim paths() As String
    Dim locates As String
    
    i = 1
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic

    'Jika data tidak ada, maka keluar dari fungsi ini!
    If M_objrs.RecordCount = 0 Then
        MsgBox "Data Blank!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
   
Form_Save:
    Cd_save.ShowSave
    TxtPath.text = Cd_save.FileName
    
    On Error GoTo errors
    paths = Split(TxtPath.text, "\")
    For q = 0 To 10
        zz = paths(q)
        ff = q
    Next q
errors:
    locates = ""
    For W = 0 To ff - 1
        If W = 0 Then
            locates = locates & paths(W)
        Else
            locates = locates & "\" & paths(W)
        End If
    Next W
    
    locates = locates
    
    'Cek apakah user menekan tombol cancel pada dialog save
    If TxtPath.text = Empty Then
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
    
    'Set excel
    Set objExcel = CreateObject("Excel.Application")
    Set objBook = objExcel.Workbooks.ADD
    Set objSheet = objBook.ActiveSheet
            
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
    
    'tandain
'    Set objVISIT = New ADODB.Recordset
'    objVISIT.CursorLocation = adUseClient
'    strsql = "SELECT array_to_string(array(select substr('abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789', trunc(random() * 62)::integer + 1, 1) FROM generate_series(1, 10)), '');"
'    objVISIT.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'
'    passwords = objVISIT(0)
'
'    objBook.SaveAs TxtPath.text, xlWorkbookNormal
'    ', passwords
'    Call logwktcti(locates, passwords)
    objExcel.Quit
    Set objExcel = Nothing: Set objBook = Nothing: Set objSheet = Nothing
    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    Set M_objrs = Nothing
 
Salah:
    Exit Sub

End Sub

Public Function Last6TT()
Dim M_objrs As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim CMDSQL As String
On Error Resume Next

    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    CMDSQL = "select noapp from vw_report_last3"
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

    M_OBJCONN.Execute "delete from last6tt"

    If M_objrs.RecordCount > 0 Then
        For z = 1 To M_objrs.RecordCount
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            CMDSQL = "select id_mgm,noapp,tt_statuscall,to_char (tt_tglcall,'yyyy-mm-dd hh24:mi:ss')::timestamp as tt_tglcall,tt_remarks from mgm_verification_hst_tt where noapp = '" & M_objrs!noapp & "' order by id desc"
            rs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic

            If rs.RecordCount > 0 Then
                For i = 1 To 6
                    If i = 1 Then
                        M_OBJCONN.Execute "insert into last6tt values ('" & rs!id_mgm & "','" & rs!noapp & "','" & rs!tt_tglcall & "','" & rs!tt_statuscall & "','" & rs!tt_remarks & "');"
                    ElseIf i = 2 Then
                        M_OBJCONN.Execute "update last6tt set status2 = '" & rs!tt_statuscall & "', tanggal2 = '" & rs!tt_tglcall & "', remarks2 = '" & rs!tt_remarks & "' where noapp = '" & rs!noapp & "';"
                    ElseIf i = 3 Then
                        M_OBJCONN.Execute "update last6tt set status3 = '" & rs!tt_statuscall & "', tanggal3 = '" & rs!tt_tglcall & "', remarks3 = '" & rs!tt_remarks & "' where noapp = '" & rs!noapp & "';"
                    ElseIf i = 4 Then
                        M_OBJCONN.Execute "update last6tt set status4 = '" & rs!tt_statuscall & "', tanggal4 = '" & rs!tt_tglcall & "', remarks4 = '" & rs!tt_remarks & "' where noapp = '" & rs!noapp & "';"
                    ElseIf i = 5 Then
                        M_OBJCONN.Execute "update last6tt set status5 = '" & rs!tt_statuscall & "', tanggal5 = '" & rs!tt_tglcall & "', remarks5 = '" & rs!tt_remarks & "' where noapp = '" & rs!noapp & "';"
                    ElseIf i = 6 Then
                        M_OBJCONN.Execute "update last6tt set status6 = '" & rs!tt_statuscall & "', tanggal6 = '" & rs!tt_tglcall & "', remarks6 = '" & rs!tt_remarks & "' where noapp = '" & rs!noapp & "';"
                    End If
                    rs.MoveNext
                Next i
            End If
            M_objrs.MoveNext
        Next z
    End If
    'MsgBox "Done"
End Function
Public Function Last6Kan()
Dim M_objrs As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim CMDSQL As String
On Error Resume Next

    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    CMDSQL = "select noapp from vw_report_last3"
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

    M_OBJCONN.Execute "delete from last6kan"

    If M_objrs.RecordCount > 0 Then
        For z = 1 To M_objrs.RecordCount
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            CMDSQL = "select id_mgm,noapp,kan_statuscall,to_char (kan_tglcall,'yyyy-mm-dd hh24:mi:ss')::timestamp as kan_tglcall,kan_remarks from mgm_verification_hst_kantor where noapp = '" & M_objrs!noapp & "' order by id desc"
            rs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic

            If rs.RecordCount > 0 Then
                For i = 1 To 6
                    If i = 1 Then
                        M_OBJCONN.Execute "insert into last6kan values ('" & rs!id_mgm & "','" & rs!noapp & "','" & rs!kan_tglcall & "','" & rs!kan_statuscall & "','" & rs!kan_remarks & "');"
                    ElseIf i = 2 Then
                        M_OBJCONN.Execute "update last6kan set kan_status2 = '" & rs!kan_statuscall & "', kan_tanggal2 = '" & rs!kan_tglcall & "', kan_remarks2 = '" & rs!kan_remarks & "' where noapp = '" & rs!noapp & "';"
                    ElseIf i = 3 Then
                        M_OBJCONN.Execute "update last6kan set kan_status3 = '" & rs!kan_statuscall & "', kan_tanggal3 = '" & rs!kan_tglcall & "', kan_remarks3 = '" & rs!kan_remarks & "' where noapp = '" & rs!noapp & "';"
                    ElseIf i = 4 Then
                        M_OBJCONN.Execute "update last6kan set kan_status4 = '" & rs!kan_statuscall & "', kan_tanggal4 = '" & rs!kan_tglcall & "', kan_remarks4 = '" & rs!kan_remarks & "' where noapp = '" & rs!noapp & "';"
                    ElseIf i = 5 Then
                        M_OBJCONN.Execute "update last6kan set kan_status5 = '" & rs!kan_statuscall & "', kan_tanggal5 = '" & rs!kan_tglcall & "', kan_remarks5 = '" & rs!kan_remarks & "' where noapp = '" & rs!noapp & "';"
                    ElseIf i = 6 Then
                        M_OBJCONN.Execute "update last6kan set kan_status6 = '" & rs!kan_statuscall & "', kan_tanggal6 = '" & rs!kan_tglcall & "', kan_remarks6 = '" & rs!kan_remarks & "' where noapp = '" & rs!noapp & "';"
                    End If
                    rs.MoveNext
                Next i
            End If
            M_objrs.MoveNext
        Next z
    End If
    'MsgBox "Done"
End Function
Public Function Last6Econ()
Dim M_objrs As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim CMDSQL As String
On Error Resume Next

    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    CMDSQL = "select noapp from vw_report_last3"
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

    M_OBJCONN.Execute "delete from last6econ"

    If M_objrs.RecordCount > 0 Then
        For z = 1 To M_objrs.RecordCount
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            CMDSQL = "select id_mgm,noapp,econ_statuscall,to_char (econ_tglcall,'yyyy-mm-dd hh24:mi:ss')::timestamp as econ_tglcall,econ_remarks from mgm_verification_hst_econ where noapp = '" & M_objrs!noapp & "' order by id desc"
            rs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic

            If rs.RecordCount > 0 Then
                For i = 1 To 6
                    If i = 1 Then
                        M_OBJCONN.Execute "insert into last6econ values ('" & rs!id_mgm & "','" & rs!noapp & "','" & rs!econ_tglcall & "','" & rs!econ_statuscall & "','" & rs!econ_remarks & "');"
                    ElseIf i = 2 Then
                        M_OBJCONN.Execute "update last6econ set econ_status2 = '" & rs!econ_statuscall & "', econ_tanggal2 = '" & rs!econ_tglcall & "', econ_remarks2 = '" & rs!econ_remarks & "' where noapp = '" & rs!noapp & "';"
                    ElseIf i = 3 Then
                        M_OBJCONN.Execute "update last6econ set econ_status3 = '" & rs!econ_statuscall & "', econ_tanggal3 = '" & rs!econ_tglcall & "', econ_remarks3 = '" & rs!econ_remarks & "' where noapp = '" & rs!noapp & "';"
                    ElseIf i = 4 Then
                        M_OBJCONN.Execute "update last6econ set econ_status4 = '" & rs!econ_statuscall & "', econ_tanggal4 = '" & rs!econ_tglcall & "', econ_remarks4 = '" & rs!econ_remarks & "' where noapp = '" & rs!noapp & "';"
                    ElseIf i = 5 Then
                        M_OBJCONN.Execute "update last6econ set econ_status5 = '" & rs!econ_statuscall & "', econ_tanggal5 = '" & rs!econ_tglcall & "', econ_remarks5 = '" & rs!econ_remarks & "' where noapp = '" & rs!noapp & "';"
                    ElseIf i = 6 Then
                        M_OBJCONN.Execute "update last6econ set econ_status6 = '" & rs!econ_statuscall & "', econ_tanggal6 = '" & rs!econ_tglcall & "', econ_remarks6 = '" & rs!econ_remarks & "' where noapp = '" & rs!noapp & "';"
                    End If
                    rs.MoveNext
                Next i
            End If
            M_objrs.MoveNext
        Next z
    End If
    'MsgBox "Done"
End Function
Public Function Last6App()
Dim M_objrs As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim CMDSQL As String
On Error Resume Next

    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    CMDSQL = "select noapp from vw_report_last3"
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

    M_OBJCONN.Execute "delete from last6app"

    If M_objrs.RecordCount > 0 Then
        For z = 1 To M_objrs.RecordCount
            Set rs = New ADODB.Recordset
            rs.CursorLocation = adUseClient
            CMDSQL = "select id_mgm,noapp,app_statuscall,to_char (app_tglcall,'yyyy-mm-dd hh24:mi:ss')::timestamp as app_tglcall,app_remarks from mgm_verification_hst_app where noapp = '" & M_objrs!noapp & "' order by id desc"
            rs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic

            If rs.RecordCount > 0 Then
                For i = 1 To 6
                    If i = 1 Then
                        M_OBJCONN.Execute "insert into last6app values ('" & rs!id_mgm & "','" & rs!noapp & "','" & rs!app_tglcall & "','" & rs!app_statuscall & "','" & rs!app_remarks & "');"
                    ElseIf i = 2 Then
                        M_OBJCONN.Execute "update last6app set app_status2 = '" & rs!app_statuscall & "', app_tanggal2 = '" & rs!app_tglcall & "', app_remarks2 = '" & rs!app_remarks & "' where noapp = '" & rs!noapp & "';"
                    ElseIf i = 3 Then
                        M_OBJCONN.Execute "update last6app set app_status3 = '" & rs!app_statuscall & "', app_tanggal3 = '" & rs!app_tglcall & "', app_remarks3 = '" & rs!app_remarks & "' where noapp = '" & rs!noapp & "';"
                    ElseIf i = 4 Then
                        M_OBJCONN.Execute "update last6app set app_status4 = '" & rs!app_statuscall & "', app_tanggal4 = '" & rs!app_tglcall & "', app_remarks4 = '" & rs!app_remarks & "' where noapp = '" & rs!noapp & "';"
                    ElseIf i = 5 Then
                        M_OBJCONN.Execute "update last6app set app_status5 = '" & rs!app_statuscall & "', app_tanggal5 = '" & rs!app_tglcall & "', app_remarks5 = '" & rs!app_remarks & "' where noapp = '" & rs!noapp & "';"
                    ElseIf i = 6 Then
                        M_OBJCONN.Execute "update last6app set app_status6 = '" & rs!app_statuscall & "', app_tanggal6 = '" & rs!app_tglcall & "', app_remarks6 = '" & rs!app_remarks & "' where noapp = '" & rs!noapp & "';"
                    End If
                    rs.MoveNext
                Next i
            End If
            M_objrs.MoveNext
        Next z
    End If
    'MsgBox "Done"
End Function
