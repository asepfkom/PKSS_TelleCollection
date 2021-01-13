VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form_Report_dika 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Report MIS"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12225
   FillColor       =   &H0080FF80&
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   12225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Criteria Report"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4740
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   12225
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Agent"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3870
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   6375
         Begin VB.TextBox txt_jumlah_acc 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5280
            TabIndex        =   14
            Top             =   3480
            Width           =   975
         End
         Begin VB.CheckBox CheckAll_MGR 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check All"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   3480
            Width           =   1455
         End
         Begin VB.CommandButton Cmd_Refersh3 
            BackColor       =   &H00F1E5DB&
            Caption         =   "REFRESH"
            Height          =   255
            Left            =   3960
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   4440
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtJmlAgent 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7680
            Locked          =   -1  'True
            TabIndex        =   8
            Text            =   "0"
            Top             =   4440
            Width           =   975
         End
         Begin VB.CheckBox CheckAll_Agent 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check All"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   4440
            Width           =   1455
         End
         Begin MSComctlLib.ListView LVAgent 
            Height          =   3120
            Left            =   90
            TabIndex        =   10
            Top             =   300
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   5503
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   12582912
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   0
            Enabled         =   0   'False
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
         Begin MSComctlLib.ProgressBar ProgressBar3 
            Height          =   255
            Left            =   2040
            TabIndex        =   11
            Top             =   4440
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
         End
         Begin VB.Label Label30 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Jumlah User :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   5640
            TabIndex        =   12
            Top             =   4440
            Width           =   2055
         End
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
         Height          =   1380
         Left            =   10500
         Picture         =   "Form_Report_Berto.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   195
         Width           =   1620
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
         Height          =   360
         Left            =   10500
         Picture         =   "Form_Report_Berto.frx":0766
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1665
         Width           =   1620
      End
      Begin MSComDlg.CommonDialog Cd_save 
         Left            =   6840
         Top             =   1200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "*.xlsx"
      End
      Begin MSComCtl2.DTPicker tgl_tracking 
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   260
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM-yyyy"
         Format          =   112525315
         CurrentDate     =   42706
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Bulan Dan Tahun"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Report MIS"
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
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   3585
   End
   Begin VB.Image Image2 
      Height          =   825
      Left            =   0
      Picture         =   "Form_Report_Berto.frx":0DAC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12240
   End
End
Attribute VB_Name = "Form_Report_dika"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheckAll_MGR_Click()
    If CheckAll_MGR.Value = 1 Then
        If LVAgent.ListItems.Count <> 0 Then
            For i = 1 To LVAgent.ListItems.Count
                LVAgent.ListItems(i).Checked = True
            Next i
        End If
    ElseIf CheckAll_MGR.Value = 0 Then
        If LVAgent.ListItems.Count <> 0 Then
            For i = 1 To LVAgent.ListItems.Count
                LVAgent.ListItems(i).Checked = False
            Next i
        End If
    End If
End Sub

Private Sub Form_Load()
    tgl_tracking.Value = Now
    
    CheckAll_MGR.Value = 1
        
    Call HeaderLvAgent
    Call ISIAGENT
    Call CheckAll_MGR_Click

End Sub

Private Sub ISIAGENT()
    Dim sQuery As String
    Dim Rs_Agent As ADODB.Recordset
    Dim Nomor As Double
    Dim list As ListItem
    
    sQuery = "SELECT * FROM usertbl WHERE aktif = '1' AND kdlevel='1' order by agent "
    Set Rs_Agent = New ADODB.Recordset
    Rs_Agent.CursorLocation = adUseClient
    Rs_Agent.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    LVAgent.ListItems.CLEAR
    
    If Rs_Agent.RecordCount > 0 Then
        While Not Rs_Agent.EOF
            Nomor = Nomor + 1
            Set list = LVAgent.ListItems.ADD(, , Nomor)
                list.SubItems(1) = Trim(Rs_Agent("userid"))
            Rs_Agent.MoveNext
        Wend
    End If
    
    txt_jumlah_acc = Rs_Agent.RecordCount
End Sub

Private Sub HeaderLvAgent()
    LVAgent.ColumnHeaders.ADD 1, , "No", 600
    LVAgent.ColumnHeaders.ADD 2, , "AGENT", 5000
End Sub


Private Sub SSCommand1_Click()
    Dim cQuery As String
    Dim RS_Report As ADODB.Recordset
    Dim ListCustId As String
    Dim K As Integer
    Dim cek As Integer
    Dim sql_tahun As String
    Dim sql_bulan As String
    Dim tgl_track As Date
    
    tgl_track = Format(tgl_tracking.Value, "yyyy-mm-dd")
    
    sql_tahun = Format(tgl_track, "yyyy")
    sql_bulan = Format(tgl_track, "mm")
    
    If LVAgent.ListItems.Count = 0 Then
        MsgBox "Agent Tidak Tersedia", vbOKOnly + vbInformation, "Informasi"
    Exit Sub
    End If
    
    For K = 1 To LVAgent.ListItems.Count
        If LVAgent.ListItems(K).Checked = True Then
            cek = cek + 1
        End If
    Next K
    
    If cek = 0 Then
        MsgBox "Pilih Agent Terlebih Dahulu!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    For i = 1 To txt_jumlah_acc.Text
        If LVAgent.ListItems(i).Checked = True Then
            ListCustId = ListCustId & "'" & LVAgent.ListItems(i).SubItems(1) & "',"
        End If
    Next i
    
    ListCustId = Mid(ListCustId, 1, Len(ListCustId) - 1)
    
    cQuery = " DROP TABLE IF EXISTS tbl_report_dika "
    M_OBJCONN.Execute cQuery
    
    cQuery = "CREATE TABLE tbl_report_dika AS ( "
    cQuery = cQuery + vbCrLf + " SELECT  date(tglcall) as tgl_call,recsource, b.tblstatuscall_kdstscall,statuscall, agent, "
    cQuery = cQuery + vbCrLf + " CASE WHEN statuscall = 'Already Paid' then count(a.id) end as jml_apaid,"
    cQuery = cQuery + vbCrLf + " CASE WHEN statuscall = 'Schedule Call' then count(a.id) end as jml_schedule,"
    cQuery = cQuery + vbCrLf + " CASE WHEN statuscall = 'Left Message' then count(a.id) end as jml_left_message,"
    cQuery = cQuery + vbCrLf + " CASE WHEN statuscall = 'PTP' then count(a.id) end as jml_ptp,"
    cQuery = cQuery + vbCrLf + " CASE WHEN statuscall = 'Negosiasi' then count(a.id) end as jml_negosiasi,"
    cQuery = cQuery + vbCrLf + " CASE WHEN statuscall = 'Invalid' then count(a.id) end as jml_invalid,"
    cQuery = cQuery + vbCrLf + " CASE WHEN statuscall = 'Mailbox' then count(a.id) end as jml_mailbox,"
    cQuery = cQuery + vbCrLf + " CASE WHEN statuscall = 'Unknow' then count(a.id) end as jml_unknow,"
    cQuery = cQuery + vbCrLf + " CASE WHEN statuscall = 'Dead' then count(a.id) end as jml_dead,"
    cQuery = cQuery + vbCrLf + " CASE WHEN statuscall = 'Busy' then count(a.id) end as jml_busy,"
    cQuery = cQuery + vbCrLf + " CASE WHEN statuscall = 'Pindah Alamat' then count(a.id) end as jml_pindah_alamat,"
    cQuery = cQuery + vbCrLf + " CASE WHEN statuscall = 'Salah Sambung' then count(a.id) end as jml_salbung,"
    cQuery = cQuery + vbCrLf + " CASE WHEN statuscall = 'Tidak Ada di Tempat' then count(a.id) end as jml_tdk_ditempat,"
    cQuery = cQuery + vbCrLf + " CASE WHEN statuscall = 'Tidak Diangkat' then count(a.id) end as jml_tdk_diangkat,"
    cQuery = cQuery + vbCrLf + " CASE WHEN statuscall = 'BP' then count(a.id) end as jml_bp,"
    cQuery = cQuery + vbCrLf + " CASE WHEN statuscall = 'Data Retur' then count(a.id) end as jml_data_retur,"
    cQuery = cQuery + vbCrLf + " CASE WHEN coalesce(statuscall,'') = '' then count(a.id) end as jml_new_data"
    '--------------------------------------------------------------------------
    cQuery = cQuery + vbCrLf + " FROM mgm a, tblstatuscall b"
    cQuery = cQuery + vbCrLf + " WHERE a.statuscall=b.tblstatuscall_keterangan "
    cQuery = cQuery + vbCrLf + " AND agent in (" + ListCustId + ")  "
    cQuery = cQuery + vbCrLf + " AND date_part('year', tglcall) = '" & sql_tahun & "' "
    cQuery = cQuery + vbCrLf + " AND date_part('month', tglcall) = '" & sql_bulan & "' "
    cQuery = cQuery + vbCrLf + " group by agent,date(tglcall),recsource,statuscall,tblstatuscall_kdstscall "
    cQuery = cQuery + vbCrLf + " order by agent, tgl_call ) "
    
    M_OBJCONN.Execute cQuery
    
    Call report_by_name
End Sub

Private Sub report_by_name()
'    Dim ExlObj As Excel.Application
    Dim TGL, kolom, Baris As Integer
    Dim sQuery As String
    Dim totalcall As Integer
    Dim RS_Report As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim tgl_excel, tgl_cek_excel, tgl_rs, agent_excel, agent_rs, speed_regular, cc_pl As String
    Dim baris_pas_ngisi, kolom_pas_ngisi, baris_isi, baris_sekarang As Integer
    Dim bulan_tahun_mulai_series, bulan_tahun_akhir_series, mulai_series, akhir_series As String
    Dim nilai, totalan_speed_cc, totalan_speed_pl, totalan_reg_cc, totalan_reg_pl, totalan_sebulan, totalan_sebulan1, totalan_sebulan2 As Double
    Dim totalan_na_speed_cc, totalan_na_speed_pl, totalan_na_reg_cc, totalan_na_reg_pl As Double
    
    
    arrayAlphabet = Array("", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", _
    "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", _
    "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ", _
    "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV", "CW", "CX", "CY", "CZ", _
    "DA", "DB", "DC", "DD", "DE", "DF", "DG", "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP", "DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", "DY", "DZ", _
    "EA", "EB", "EC", "ED", "EE", "EF", "EG", "EH", "EI", "EJ", "EK", "EL", "EM", "EN", "EO", "EP", "EQ", "ER", "ES", "ET", "EU", "EV", "EW", "EX", "EY", "EZ", _
    "FA", "FB", "FC", "FD", "FE", "FF", "FG", "FH", "FI", "FJ", "FK", "FL", "FM", "FN", "FO", "FP", "FQ", "FR", "FS", "FT", "FU", "FV", "FW", "FX", "FY", "FZ")
    
    sQuery = "SELECT * FROM tbl_report_dika limit 1"
    Set RS_Report = New ADODB.Recordset
    RS_Report.CursorLocation = adUseClient
    RS_Report.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If RS_Report.RecordCount = 0 Then
        MsgBox "Data Tidak Tersedia", vbOKOnly + vbInformation, "Informasi"
    Exit Sub
    End If
    
    Set RS_Report = Nothing
    Set ELIN2 = CreateObject("excel.application")
    ELIN2.Workbooks.ADD
    ELIN2.Visible = True
    
    ELIN2.Range("A1:DT1").MergeCells = True
'--REPORT BY NAME
    '------ELIN CODING------
    With ELIN2.ActiveSheet
    .Cells(1, 1).Value = "Report Type   : Call Activity Report"
        .Cells(1, 1).Font.Name = "Arial"
        .Cells(1, 1).Font.Size = "12"
        '.Cells(1, 1).Font.Bold = True
        .Cells(2, 1).Value = "Interval      : " & Format(tgl_tracking.Value, "DD-MM-YYYY") & ""
        .Cells(2, 1).Font.Name = "Arial"
        .Cells(2, 1).Font.Size = "12"
        '.Cells(2, 1).Font.Bold = True
        
    kolom = 1
    
        .Cells(4, kolom).Value = "Agent Name"
        .Cells(4, kolom).Font.Bold = True
        .Cells(4, kolom).Font.Size = 12
        '.Cells(4, kolom).Interior.Color = &H808080
        .Cells(4, kolom).Borders.LineStyle = xlContinuous
        .Cells(4, kolom).Borders.Weight = 3
        .Cells(4, kolom).ColumnWidth = 15
        ELIN2.Range(arrayAlphabet(kolom) & "4:" & arrayAlphabet(kolom) & "6").Merge
        ELIN2.Range(arrayAlphabet(kolom) & "4:" & arrayAlphabet(kolom) & "6").HorizontalAlignment = xlCenter
        ELIN2.Range(arrayAlphabet(kolom) & "4:" & arrayAlphabet(kolom) & "6").VerticalAlignment = xlCenter
        'ELIN2.Range(arrayAlphabet(kolom) & "4:" & arrayAlphabet(kolom) & "6").Interior.Color = &H808080
        ELIN2.Range(arrayAlphabet(kolom) & "4:" & arrayAlphabet(kolom) & "6").Borders.LineStyle = xlContinuous
        ELIN2.Range(arrayAlphabet(kolom) & "4:" & arrayAlphabet(kolom) & "6").Borders.Weight = 3
        
        .Cells(4, kolom + 1).Value = "Total Call"
        .Cells(4, kolom + 1).Font.Bold = True
        .Cells(4, kolom + 1).Font.Size = 12
        '.Cells(4, kolom + 1).Interior.Color = &H808080
        .Cells(4, kolom + 1).Borders.LineStyle = xlContinuous
        .Cells(4, kolom + 1).Borders.Weight = 3
        .Cells(4, kolom + 1).ColumnWidth = 15
        ELIN2.Range(arrayAlphabet(kolom + 1) & "4:" & arrayAlphabet(kolom + 1) & "6").Merge
        ELIN2.Range(arrayAlphabet(kolom + 1) & "4:" & arrayAlphabet(kolom + 1) & "6").HorizontalAlignment = xlCenter
        ELIN2.Range(arrayAlphabet(kolom + 1) & "4:" & arrayAlphabet(kolom + 1) & "6").VerticalAlignment = xlCenter
        'ELIN2.Range(arrayAlphabet(kolom + 1) & "4:" & arrayAlphabet(kolom + 1) & "6").Interior.Color = &H808080
        ELIN2.Range(arrayAlphabet(kolom + 1) & "4:" & arrayAlphabet(kolom + 1) & "6").Borders.LineStyle = xlContinuous
        ELIN2.Range(arrayAlphabet(kolom + 1) & "4:" & arrayAlphabet(kolom + 1) & "6").Borders.Weight = 3
        
        .Cells(4, kolom + 2).Value = "Jumlah Polis"
        .Cells(4, kolom + 2).Font.Bold = True
        .Cells(4, kolom + 2).Font.Size = 12
        '.Cells(4, kolom + 2).Interior.Color = &H808080
        .Cells(4, kolom + 2).Borders.LineStyle = xlContinuous
        .Cells(4, kolom + 2).Borders.Weight = 3
        .Cells(4, kolom + 2).ColumnWidth = 15
        ELIN2.Range(arrayAlphabet(kolom + 2) & "4:" & arrayAlphabet(kolom + 2) & "6").Merge
        ELIN2.Range(arrayAlphabet(kolom + 2) & "4:" & arrayAlphabet(kolom + 2) & "6").HorizontalAlignment = xlCenter
        ELIN2.Range(arrayAlphabet(kolom + 2) & "4:" & arrayAlphabet(kolom + 2) & "6").VerticalAlignment = xlCenter
        'ELIN2.Range(arrayAlphabet(kolom + 2) & "4:" & arrayAlphabet(kolom + 2) & "6").Interior.Color = &H808080
        ELIN2.Range(arrayAlphabet(kolom + 2) & "4:" & arrayAlphabet(kolom + 2) & "6").Borders.LineStyle = xlContinuous
        ELIN2.Range(arrayAlphabet(kolom + 2) & "4:" & arrayAlphabet(kolom + 2) & "6").Borders.Weight = 3
        
        
            kolom = kolom + 3
                .Cells(4, kolom).Value = "Status"
                .Cells(4, kolom).Font.Bold = True
                
                '----------------------------------------------------------------------------------------
                .Cells(5, kolom).Value = "Already Paid"
                ELIN2.Range(arrayAlphabet(kolom) & "5:" & arrayAlphabet(kolom) & "6").Merge
                ELIN2.Range(arrayAlphabet(kolom) & "5:" & arrayAlphabet(kolom) & "6").HorizontalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom) & "5:" & arrayAlphabet(kolom) & "6").VerticalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom) & "5:" & arrayAlphabet(kolom) & "6").Borders.LineStyle = xlContinuous
                '----------------------------------------------------------------------------------------
                
                .Cells(5, kolom + 1).Value = "BP/Broken Promise"
                ELIN2.Range(arrayAlphabet(kolom + 1) & "5:" & arrayAlphabet(kolom + 1) & "6").Merge
                ELIN2.Range(arrayAlphabet(kolom + 1) & "5:" & arrayAlphabet(kolom + 1) & "6").HorizontalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 1) & "5:" & arrayAlphabet(kolom + 1) & "6").VerticalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 1) & "5:" & arrayAlphabet(kolom + 1) & "6").Borders.LineStyle = xlContinuous
                '----------------------------------------------------------------------------------------
            ELIN2.Range(arrayAlphabet(kolom) & "4:" & arrayAlphabet(kolom + 15) & "4").Merge
            ELIN2.Range(arrayAlphabet(kolom) & "4:" & arrayAlphabet(kolom + 15) & "4").HorizontalAlignment = xlCenter
            ELIN2.Range(arrayAlphabet(kolom) & "4:" & arrayAlphabet(kolom + 15) & "4").Interior.Color = &HFF00FF
            ELIN2.Range(arrayAlphabet(kolom) & "4:" & arrayAlphabet(kolom + 15) & "4").Borders.LineStyle = xlContinuous
            ELIN2.Range(arrayAlphabet(kolom) & "4:" & arrayAlphabet(kolom + 15) & "4").Borders.Weight = 3
            
                '----------------------------------------------------------------------------------------
                .Cells(5, kolom + 2).Value = "busy"
                ELIN2.Range(arrayAlphabet(kolom + 2) & "5:" & arrayAlphabet(kolom + 2) & "6").Merge
                ELIN2.Range(arrayAlphabet(kolom + 2) & "5:" & arrayAlphabet(kolom + 2) & "6").HorizontalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 2) & "5:" & arrayAlphabet(kolom + 2) & "6").VerticalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 2) & "5:" & arrayAlphabet(kolom + 2) & "6").Borders.LineStyle = xlContinuous
                '----------------------------------------------------------------------------------------
                .Cells(5, kolom + 3).Value = "Data Retur"
                ELIN2.Range(arrayAlphabet(kolom + 3) & "5:" & arrayAlphabet(kolom + 3) & "6").Merge
                ELIN2.Range(arrayAlphabet(kolom + 3) & "5:" & arrayAlphabet(kolom + 3) & "6").HorizontalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 3) & "5:" & arrayAlphabet(kolom + 3) & "6").VerticalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 3) & "5:" & arrayAlphabet(kolom + 3) & "6").Borders.LineStyle = xlContinuous
                '----------------------------------------------------------------------------------------
                .Cells(5, kolom + 4).Value = "Invalid"
                ELIN2.Range(arrayAlphabet(kolom + 4) & "5:" & arrayAlphabet(kolom + 4) & "6").Merge
                ELIN2.Range(arrayAlphabet(kolom + 4) & "5:" & arrayAlphabet(kolom + 4) & "6").HorizontalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 4) & "5:" & arrayAlphabet(kolom + 4) & "6").VerticalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 4) & "5:" & arrayAlphabet(kolom + 4) & "6").Borders.LineStyle = xlContinuous
                '----------------------------------------------------------------------------------------
                .Cells(5, kolom + 5).Value = "KP/Keep Promise"
                ELIN2.Range(arrayAlphabet(kolom + 5) & "5:" & arrayAlphabet(kolom + 5) & "6").Merge
                ELIN2.Range(arrayAlphabet(kolom + 5) & "5:" & arrayAlphabet(kolom + 5) & "6").HorizontalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 5) & "5:" & arrayAlphabet(kolom + 5) & "6").VerticalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 5) & "5:" & arrayAlphabet(kolom + 5) & "6").Borders.LineStyle = xlContinuous
                '----------------------------------------------------------------------------------------
                .Cells(5, kolom + 6).Value = "Left Message"
                ELIN2.Range(arrayAlphabet(kolom + 6) & "5:" & arrayAlphabet(kolom + 6) & "6").Merge
                ELIN2.Range(arrayAlphabet(kolom + 6) & "5:" & arrayAlphabet(kolom + 6) & "6").HorizontalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 6) & "5:" & arrayAlphabet(kolom + 6) & "6").VerticalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 6) & "5:" & arrayAlphabet(kolom + 6) & "6").Borders.LineStyle = xlContinuous
                '----------------------------------------------------------------------------------------
                .Cells(5, kolom + 7).Value = "Mailbox"
                ELIN2.Range(arrayAlphabet(kolom + 7) & "5:" & arrayAlphabet(kolom + 7) & "6").Merge
                ELIN2.Range(arrayAlphabet(kolom + 7) & "5:" & arrayAlphabet(kolom + 7) & "6").HorizontalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 7) & "5:" & arrayAlphabet(kolom + 7) & "6").VerticalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 7) & "5:" & arrayAlphabet(kolom + 7) & "6").Borders.LineStyle = xlContinuous
                '----------------------------------------------------------------------------------------
                .Cells(5, kolom + 8).Value = "Negosiasi"
                ELIN2.Range(arrayAlphabet(kolom + 8) & "5:" & arrayAlphabet(kolom + 8) & "6").Merge
                ELIN2.Range(arrayAlphabet(kolom + 8) & "5:" & arrayAlphabet(kolom + 8) & "6").HorizontalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 8) & "5:" & arrayAlphabet(kolom + 8) & "6").VerticalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 8) & "5:" & arrayAlphabet(kolom + 8) & "6").Borders.LineStyle = xlContinuous
                '----------------------------------------------------------------------------------------
                .Cells(5, kolom + 9).Value = "Pindah Alamat"
                ELIN2.Range(arrayAlphabet(kolom + 9) & "5:" & arrayAlphabet(kolom + 9) & "6").Merge
                ELIN2.Range(arrayAlphabet(kolom + 9) & "5:" & arrayAlphabet(kolom + 9) & "6").HorizontalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 9) & "5:" & arrayAlphabet(kolom + 9) & "6").VerticalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 9) & "5:" & arrayAlphabet(kolom + 9) & "6").Borders.LineStyle = xlContinuous
                '----------------------------------------------------------------------------------------
                .Cells(5, kolom + 10).Value = "PTP"
                ELIN2.Range(arrayAlphabet(kolom + 10) & "5:" & arrayAlphabet(kolom + 10) & "6").Merge
                ELIN2.Range(arrayAlphabet(kolom + 10) & "5:" & arrayAlphabet(kolom + 10) & "6").HorizontalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 10) & "5:" & arrayAlphabet(kolom + 10) & "6").VerticalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 10) & "5:" & arrayAlphabet(kolom + 10) & "6").Borders.LineStyle = xlContinuous
                '----------------------------------------------------------------------------------------
                .Cells(5, kolom + 11).Value = "Salah Sambung"
                ELIN2.Range(arrayAlphabet(kolom + 11) & "5:" & arrayAlphabet(kolom + 11) & "6").Merge
                ELIN2.Range(arrayAlphabet(kolom + 11) & "5:" & arrayAlphabet(kolom + 11) & "6").HorizontalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 11) & "5:" & arrayAlphabet(kolom + 11) & "6").VerticalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 11) & "5:" & arrayAlphabet(kolom + 11) & "6").Borders.LineStyle = xlContinuous
                '----------------------------------------------------------------------------------------
                .Cells(5, kolom + 12).Value = "Schedule Call"
                ELIN2.Range(arrayAlphabet(kolom + 12) & "5:" & arrayAlphabet(kolom + 12) & "6").Merge
                ELIN2.Range(arrayAlphabet(kolom + 12) & "5:" & arrayAlphabet(kolom + 12) & "6").HorizontalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 12) & "5:" & arrayAlphabet(kolom + 12) & "6").VerticalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 12) & "5:" & arrayAlphabet(kolom + 12) & "6").Borders.LineStyle = xlContinuous
                '----------------------------------------------------------------------------------------
                .Cells(5, kolom + 13).Value = "Tidak diangkat"
                ELIN2.Range(arrayAlphabet(kolom + 13) & "5:" & arrayAlphabet(kolom + 13) & "6").Merge
                ELIN2.Range(arrayAlphabet(kolom + 13) & "5:" & arrayAlphabet(kolom + 13) & "6").HorizontalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 13) & "5:" & arrayAlphabet(kolom + 13) & "6").VerticalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 13) & "5:" & arrayAlphabet(kolom + 13) & "6").Borders.LineStyle = xlContinuous
                '----------------------------------------------------------------------------------------
                .Cells(5, kolom + 14).Value = "Tidak Ada di Tempat"
                ELIN2.Range(arrayAlphabet(kolom + 14) & "5:" & arrayAlphabet(kolom + 14) & "6").Merge
                ELIN2.Range(arrayAlphabet(kolom + 14) & "5:" & arrayAlphabet(kolom + 14) & "6").HorizontalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 14) & "5:" & arrayAlphabet(kolom + 14) & "6").VerticalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 14) & "5:" & arrayAlphabet(kolom + 14) & "6").Borders.LineStyle = xlContinuous
                '----------------------------------------------------------------------------------------
                .Cells(5, kolom + 15).Value = "Unknown"
                ELIN2.Range(arrayAlphabet(kolom + 15) & "5:" & arrayAlphabet(kolom + 15) & "6").Merge
                ELIN2.Range(arrayAlphabet(kolom + 15) & "5:" & arrayAlphabet(kolom + 15) & "6").HorizontalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 15) & "5:" & arrayAlphabet(kolom + 15) & "6").VerticalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom + 15) & "5:" & arrayAlphabet(kolom + 15) & "6").Borders.LineStyle = xlContinuous
                '----------------------------------------------------------------------------------------
          Baris = 7
        kolom = 2
        For i = 1 To LVAgent.ListItems.Count
            If LVAgent.ListItems(i).Checked = True Then
                
                agent_excel = LVAgent.ListItems(i).ListSubItems(1)
                
                .Cells(Baris, kolom - 1).Value = agent_excel
                .Cells(Baris, kolom - 1).Font.Bold = True
                .Cells(Baris, kolom - 1).Interior.Color = &HC0C0C0
                .Cells(Baris, kolom - 1).HorizontalAlignment = xlCenter
                .Cells(Baris, kolom - 1).Borders.LineStyle = xlContinuous
                .Cells(Baris, kolom - 1).Borders.Weight = 3
                Baris = Baris + 1
           
                
        'ngisi total call
                
                strsql = " SELECT agent,count(ID) as jml FROM MGM_HST where id >0 and agent = '" & agent_excel & "'"
                strsql = strsql + "  Group by agent "

                Set rs = New ADODB.Recordset
                rs.CursorLocation = adUseClient
                rs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If rs.RecordCount > 0 Then
                    baris_isi = 6
                    kolom = 2
                    nilai = 0
                    totalan_call = 0
                    For ELIN = 1 To rs.RecordCount
                        baris_isi = baris_isi + 1
                        .Cells(baris_isi, kolom).Value = IIf(IsNull(RS_Report!jml), "0", RS_Report!jml)
                        .Cells(baris_isi, kolom).HorizontalAlignment = xlCenter
                        '.cells(baris_pas_ngisi, kolom).Interior.Color = &H808080
                        .Cells(baris_isi, kolom).Borders.LineStyle = xlContinuous
                        nilai = IIf(IsNull(RS_Report!jml), "0", RS_Report!jml)
                        
                        totalan_call = totalan_call + nilai
                        
                        rs.MoveNext
                    Next ELIN
                    '.Cells(baris_pas_ngisi + 1, kolom).Font.Color = &HFFFFFF
                    '.Cells(baris_pas_ngisi + 1, kolom).Value = totalan_call
                    
                End If
                
                'ngisi total already paid
                
                sQuery = " SELECT userid ,COALESCE(sum_apaid,0) AS ELIN FROM ("
                sQuery = sQuery + "(SELECT userid FROM usertbl WHERE spvcode = 'SPV1' and USERID = '" & agent_excel & "' AND aktif = '1' ORDER BY ID) as A LEFT JOIN"
                sQuery = sQuery + "(select agent,sum(jml_Apaid) AS sum_apaid from tbl_report_dika  group  by agent order by agent ) as b ON a.userid = b.agent)"
                Set RS_Report = New ADODB.Recordset
                RS_Report.CursorLocation = adUseClient
                RS_Report.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If RS_Report.RecordCount > 0 Then
                    baris_pas_ngisi = 6
                    kolom = 4
                    nilai = 0
                    totalan_apaid = 0
                    For ELIN = 1 To RS_Report.RecordCount
                        baris_pas_ngisi = baris_pas_ngisi + 1
                        .Cells(baris_pas_ngisi, kolom).Value = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        .Cells(baris_pas_ngisi, kolom).HorizontalAlignment = xlCenter
                        '.cells(baris_pas_ngisi, kolom).Interior.Color = &H808080
                        .Cells(baris_pas_ngisi, kolom).Borders.LineStyle = xlContinuous
                        nilai = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        
                        totalan_apaid = totalan_apaid + nilai
                        
                        RS_Report.MoveNext
                    Next ELIN
                    .Cells(baris_pas_ngisi + 1, kolom).Font.Color = &HFFFFFF
                    .Cells(baris_pas_ngisi + 1, kolom).Value = totalan_apaid
                    
                End If
                   
                
                'NGISI total BP
                
                sQuery = " SELECT userid ,COALESCE(sum_bp,0) AS ELIN FROM ("
                sQuery = sQuery + "(SELECT userid FROM usertbl WHERE spvcode = 'SPV1' and USERID = '" & agent_excel & "' AND aktif = '1' ORDER BY ID) as A LEFT JOIN"
                sQuery = sQuery + "(select agent,sum(jml_bp) AS sum_bp from tbl_report_dika  group  by agent order by agent ) as b ON a.userid = b.agent)"
                Set RS_Report = New ADODB.Recordset
                RS_Report.CursorLocation = adUseClient
                RS_Report.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If RS_Report.RecordCount > 0 Then
                    baris_pas_ngisi = 6
                    nilai = 0
                    totalan_bp = 0
                    For dika = 1 To RS_Report.RecordCount
                        baris_pas_ngisi = baris_pas_ngisi + 1
                        .Cells(baris_pas_ngisi, kolom + 1).Value = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        .Cells(baris_pas_ngisi, kolom + 1).HorizontalAlignment = xlCenter
                        .Cells(baris_pas_ngisi, kolom + 1).Borders.LineStyle = xlContinuous
                        .Cells(baris_pas_ngisi, kolom + 1).Borders(xlEdgeRight).Weight = 3
                        nilai = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        
                        totalan_bp = totalan_bp + nilai
                        
                        RS_Report.MoveNext
                    Next dika
                    
                    .Cells(baris_pas_ngisi + 1, kolom).Font.Color = &HFFFFFF
                    .Cells(baris_pas_ngisi + 1, kolom).Value = totalan_bp
                    '------------------
                End If
               
                'NGISI TOTAL busy
                
                sQuery = " SELECT userid ,COALESCE(sum_busy,0) AS ELIN FROM ("
                sQuery = sQuery + "(SELECT userid FROM usertbl WHERE spvcode = 'SPV1' and USERID = '" & agent_excel & "' AND aktif = '1' ORDER BY ID) as A LEFT JOIN"
                sQuery = sQuery + "(select agent,sum(jml_busy) AS sum_busy from tbl_report_dika  group  by agent order by agent ) as b ON a.userid = b.agent)"
                Set RS_Report = New ADODB.Recordset
                RS_Report.CursorLocation = adUseClient
                RS_Report.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If RS_Report.RecordCount > 0 Then
                    baris_pas_ngisi = 6
                    nilai = 0
                    totalan_busy = 0
                    
                    For dika = 1 To RS_Report.RecordCount
                        baris_pas_ngisi = baris_pas_ngisi + 1
                        .Cells(baris_pas_ngisi, kolom + 2).Value = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        .Cells(baris_pas_ngisi, kolom + 2).HorizontalAlignment = xlCenter
                        .Cells(baris_pas_ngisi, kolom + 2).Borders.LineStyle = xlContinuous
                        
                        nilai = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        
                        totalan_busy = totalan_busy + nilai
                        
                        RS_Report.MoveNext
                    Next dika
                    .Cells(baris_pas_ngisi + 1, kolom + 2).Font.Color = &HFFFFFF
                    .Cells(baris_pas_ngisi + 1, kolom + 2).Value = totalan_busy
                End If
                
                'NGISI TOTAL DATA RETUR
                
                sQuery = " SELECT userid ,COALESCE(sum_data_retur,0) AS ELIN FROM ("
                sQuery = sQuery + "(SELECT userid FROM usertbl WHERE spvcode = 'SPV1'  and USERID = '" & agent_excel & "' AND aktif = '1' ORDER BY ID) as A LEFT JOIN"
                sQuery = sQuery + "(select agent,sum(jml_data_retur) AS sum_data_retur from tbl_report_dika  group  by agent order by agent ) as b ON a.userid = b.agent)"
                Set RS_Report = New ADODB.Recordset
                RS_Report.CursorLocation = adUseClient
                RS_Report.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If RS_Report.RecordCount > 0 Then
                    baris_pas_ngisi = 47
                    nilai = 0
                    totalan_retur = 0
                    For dika = 1 To RS_Report.RecordCount
                        baris_pas_ngisi = baris_pas_ngisi + 1
                        .Cells(baris_pas_ngisi, kolom + 3).Value = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        .Cells(baris_pas_ngisi, kolom + 3).HorizontalAlignment = xlCenter
                        .Cells(baris_pas_ngisi, kolom + 3).Borders.LineStyle = xlContinuous
                        .Cells(baris_pas_ngisi, kolom + 3).Borders(xlEdgeRight).Weight = 3
                        nilai = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        
                        totalan_retur = totalan_retur + nilai
                        
                        RS_Report.MoveNext
                    Next dika
                    .Cells(baris_pas_ngisi + 1, kolom + 2).Font.Color = &HFFFFFF
                    .Cells(baris_pas_ngisi + 1, kolom + 2).Value = totalan_retur
                End If
                
                'NGISI TOTAL invalid
                
                sQuery = " SELECT userid ,COALESCE(sum_invalid,0) AS ELIN FROM ("
                sQuery = sQuery + "(SELECT userid FROM usertbl WHERE spvcode = 'SPV1' and USERID = '" & agent_excel & "' AND aktif = '1' ORDER BY ID) as A LEFT JOIN"
                sQuery = sQuery + "(select agent,sum(jml_invalid) AS sum_invalid from tbl_report_dika  group  by agent order by agent ) as b ON a.userid = b.agent)"
                Set RS_Report = New ADODB.Recordset
                RS_Report.CursorLocation = adUseClient
                RS_Report.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If RS_Report.RecordCount > 0 Then
                    baris_pas_ngisi = 6
                    nilai = 0
                    totalan_invalid = 0
                    For dika = 1 To RS_Report.RecordCount
                        baris_pas_ngisi = baris_pas_ngisi + 1
                        .Cells(baris_pas_ngisi, kolom + 4).Value = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        .Cells(baris_pas_ngisi, kolom + 4).HorizontalAlignment = xlCenter
                        .Cells(baris_pas_ngisi, kolom + 4).Borders.LineStyle = xlContinuous
                        
                        nilai = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        
                        totalan_invalid = totalan_invalid + nilai
                        
                        RS_Report.MoveNext
                    Next dika
                     .Cells(baris_pas_ngisi + 1, kolom + 4).Font.Color = &HFFFFFF
                     .Cells(baris_pas_ngisi + 1, kolom + 4).Value = totalan_invalid
                End If
                
                'NGISI TOTAL KP / KEEP PROMISE
                
                sQuery = " SELECT userid ,COALESCE(sum_keep_promise,0) AS ELIN FROM ("
                sQuery = sQuery + "(SELECT userid FROM usertbl WHERE spvcode = 'SPV1' and USERID = '" & agent_excel & "' AND aktif = '1' ORDER BY ID) as A LEFT JOIN"
                sQuery = sQuery + "(select agent,sum(jml_dead) AS sum_keep_promise from tbl_report_dika  group  by agent order by agent ) as b ON a.userid = b.agent)"
                Set RS_Report = New ADODB.Recordset
                RS_Report.CursorLocation = adUseClient
                RS_Report.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If RS_Report.RecordCount > 0 Then
                    baris_pas_ngisi = 6
                    nilai = 0
                    totalan_kp = 0
                    For dika = 1 To RS_Report.RecordCount
                        baris_pas_ngisi = baris_pas_ngisi + 1
                        .Cells(baris_pas_ngisi, kolom + 5).Value = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        .Cells(baris_pas_ngisi, kolom + 5).HorizontalAlignment = xlCenter
                        .Cells(baris_pas_ngisi, kolom + 5).Borders.LineStyle = xlContinuous
                        .Cells(baris_pas_ngisi, kolom + 5).Borders(xlEdgeRight).Weight = 3
                        nilai = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        
                        totalan_kp = totalan_kp + nilai
                        
                        RS_Report.MoveNext
                    Next dika
                    .Cells(baris_pas_ngisi + 1, kolom + 4).Font.Color = &HFFFFFF
                     .Cells(baris_pas_ngisi + 1, kolom + 4).Value = totalan_kp
                End If
                
                
                'NGISI TOTAL left message
                
                sQuery = " SELECT userid ,COALESCE(sum_left_message,0) AS ELIN FROM ("
                sQuery = sQuery + "(SELECT userid FROM usertbl WHERE spvcode = 'SPV1' and USERID = '" & agent_excel & "' AND aktif = '1' ORDER BY ID) as A LEFT JOIN"
                sQuery = sQuery + "(select agent,sum(jml_left_message) AS sum_left_message from tbl_report_dika  group  by agent order by agent ) as b ON a.userid = b.agent)"
                Set RS_Report = New ADODB.Recordset
                RS_Report.CursorLocation = adUseClient
                RS_Report.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If RS_Report.RecordCount > 0 Then
                    baris_pas_ngisi = 6
                    nilai = 0
                    totalan_lm = 0
                    For dika = 1 To RS_Report.RecordCount
                        baris_pas_ngisi = baris_pas_ngisi + 1
                        .Cells(baris_pas_ngisi, kolom + 6).Value = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        .Cells(baris_pas_ngisi, kolom + 6).HorizontalAlignment = xlCenter
                        .Cells(baris_pas_ngisi, kolom + 6).Borders.LineStyle = xlContinuous
                        
                        nilai = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        
                        totalan_lm = totalan_lm + nilai
                        
                        RS_Report.MoveNext
                    Next dika
                    .Cells(baris_pas_ngisi + 1, kolom + 6).Font.Color = &HFFFFFF
                    .Cells(baris_pas_ngisi + 1, kolom + 6).Value = totalan_lm
                End If
                
                'NGISI TOTAL mailbox
                
                sQuery = " SELECT userid ,COALESCE(sum_mailbox,0) AS ELIN FROM ("
                sQuery = sQuery + "(SELECT userid FROM usertbl WHERE spvcode = 'SPV1' and USERID = '" & agent_excel & "' AND aktif = '1' ORDER BY ID) as A LEFT JOIN"
                sQuery = sQuery + "(select agent,sum(jml_mailbox) AS sum_mailbox from tbl_report_dika  group  by agent order by agent ) as b ON a.userid = b.agent)"
                Set RS_Report = New ADODB.Recordset
                RS_Report.CursorLocation = adUseClient
                RS_Report.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If RS_Report.RecordCount > 0 Then
                    baris_pas_ngisi = 6
                    nilai = 0
                    totalan_mailbox = 0
                    For dika = 1 To RS_Report.RecordCount
                        baris_pas_ngisi = baris_pas_ngisi + 1
                        .Cells(baris_pas_ngisi, kolom + 7).Value = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        .Cells(baris_pas_ngisi, kolom + 7).HorizontalAlignment = xlCenter
                        .Cells(baris_pas_ngisi, kolom + 7).Borders.LineStyle = xlContinuous
                        .Cells(baris_pas_ngisi, kolom + 7).Borders(xlEdgeRight).Weight = 3
                        nilai = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        
                        totalan_mailbox = totalan_mailbox + nilai
                        
                        RS_Report.MoveNext
                    Next dika
                    .Cells(baris_pas_ngisi + 1, kolom + 6).Font.Color = &HFFFFFF
                    .Cells(baris_pas_ngisi + 1, kolom + 6).Value = totalan_mailbox
                End If
                
                
                'NGISI TOTAL negosiasi
                
                sQuery = " SELECT userid ,COALESCE(sum_negosiasi,0) AS ELIN FROM ("
                sQuery = sQuery + "(SELECT userid FROM usertbl WHERE spvcode = 'SPV1' and USERID = '" & agent_excel & "' AND aktif = '1' ORDER BY ID) as A LEFT JOIN"
                sQuery = sQuery + "(select agent,sum(jml_negosiasi) AS sum_negosiasi from tbl_report_dika  group  by agent order by agent ) as b ON a.userid = b.agent)"
                Set RS_Report = New ADODB.Recordset
                RS_Report.CursorLocation = adUseClient
                RS_Report.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If RS_Report.RecordCount > 0 Then
                    baris_pas_ngisi = 6
                    nilai = 0
                    totalan_nego = 0
                    For dika = 1 To RS_Report.RecordCount
                        baris_pas_ngisi = baris_pas_ngisi + 1
                        .Cells(baris_pas_ngisi, kolom + 8).Value = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        .Cells(baris_pas_ngisi, kolom + 8).HorizontalAlignment = xlCenter
                        .Cells(baris_pas_ngisi, kolom + 8).Borders.LineStyle = xlContinuous
                        
                        nilai = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        
                        totalan_nego = totalan_nego + nilai
                        
                        RS_Report.MoveNext
                    Next dika
                     .Cells(baris_pas_ngisi + 1, kolom + 8).Font.Color = &HFFFFFF
                     .Cells(baris_pas_ngisi + 1, kolom + 8).Value = totalan_nego
                End If
                
                'NGISI TOTAL PINDAH ALAMAT
                
                sQuery = " SELECT userid ,COALESCE(sum_pindah_alamat,0) AS ELIN FROM ("
                sQuery = sQuery + "(SELECT userid FROM usertbl WHERE spvcode = 'SPV1' and USERID = '" & agent_excel & "' AND aktif = '1' ORDER BY ID) as A LEFT JOIN"
                sQuery = sQuery + "(select agent,sum(jml_pindah_alamat) AS sum_pindah_alamat from tbl_report_dika  group  by agent order by agent ) as b ON a.userid = b.agent)"
                Set RS_Report = New ADODB.Recordset
                RS_Report.CursorLocation = adUseClient
                RS_Report.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If RS_Report.RecordCount > 0 Then
                    baris_pas_ngisi = 6
                    nilai = 0
                    totalan_pindah = 0
                    For dika = 1 To RS_Report.RecordCount
                        baris_pas_ngisi = baris_pas_ngisi + 1
                        .Cells(baris_pas_ngisi, kolom + 9).Value = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        .Cells(baris_pas_ngisi, kolom + 9).HorizontalAlignment = xlCenter
                        .Cells(baris_pas_ngisi, kolom + 9).Borders.LineStyle = xlContinuous
                        .Cells(baris_pas_ngisi, kolom + 9).Borders(xlEdgeRight).Weight = 3
                        nilai = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        
                        totalan_pindah = totalan_pindah + nilai
                        
                        RS_Report.MoveNext
                    Next dika
                    .Cells(baris_pas_ngisi + 1, kolom + 8).Font.Color = &HFFFFFF
                     .Cells(baris_pas_ngisi + 1, kolom + 8).Value = totalan_pindah
                End If
                
                'NGISI TOTAL PTP
                
                sQuery = " SELECT userid ,COALESCE(sum_ptp,0) AS ELIN FROM ("
                sQuery = sQuery + "(SELECT userid FROM usertbl WHERE spvcode = 'SPV1' and USERID = '" & agent_excel & "' AND aktif = '1' ORDER BY ID) as A LEFT JOIN"
                sQuery = sQuery + "(select agent,sum(jml_ptp) AS sum_ptp from tbl_report_dika  group  by agent order by agent ) as b ON a.userid = b.agent)"
                Set RS_Report = New ADODB.Recordset
                RS_Report.CursorLocation = adUseClient
                RS_Report.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If RS_Report.RecordCount > 0 Then
                    baris_pas_ngisi = 6
                    nilai = 0
                    totalan_ptp = 0
                    For dika = 1 To RS_Report.RecordCount
                        baris_pas_ngisi = baris_pas_ngisi + 1
                        .Cells(baris_pas_ngisi, kolom + 10).Value = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        .Cells(baris_pas_ngisi, kolom + 10).HorizontalAlignment = xlCenter
                        .Cells(baris_pas_ngisi, kolom + 10).Borders.LineStyle = xlContinuous
                        .Cells(baris_pas_ngisi, kolom + 10).Borders(xlEdgeRight).Weight = 3
                        nilai = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        
                        totalan_ptp = totalan_ptp + nilai
                        
                        RS_Report.MoveNext
                    Next dika
                    .Cells(baris_pas_ngisi + 1, kolom + 9).Font.Color = &HFFFFFF
                     .Cells(baris_pas_ngisi + 1, kolom + 9).Value = totalan_ptp
                End If
                
                'NGISI TOTAL salah sambung
                
                sQuery = " SELECT userid ,COALESCE(sum_salbung,0) AS ELIN FROM ("
                sQuery = sQuery + "(SELECT userid FROM usertbl WHERE spvcode = 'SPV1' and USERID = '" & agent_excel & "' AND aktif = '1' ORDER BY ID) as A LEFT JOIN"
                sQuery = sQuery + "(select agent,sum(jml_salbung) AS sum_salbung from tbl_report_dika  group  by agent order by agent ) as b ON a.userid = b.agent)"
                Set RS_Report = New ADODB.Recordset
                RS_Report.CursorLocation = adUseClient
                RS_Report.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If RS_Report.RecordCount > 0 Then
                    baris_pas_ngisi = 6
                    nilai = 0
                    totalan_salbung = 0
                    For dika = 1 To RS_Report.RecordCount
                        baris_pas_ngisi = baris_pas_ngisi + 1
                        .Cells(baris_pas_ngisi, kolom + 11).Value = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        .Cells(baris_pas_ngisi, kolom + 11).HorizontalAlignment = xlCenter
                        .Cells(baris_pas_ngisi, kolom + 11).Borders.LineStyle = xlContinuous
                        .Cells(baris_pas_ngisi, kolom + 11).Borders(xlEdgeRight).Weight = 3
                        nilai = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        
                        totalan_salbung = totalan_salbung + nilai
                        
                        RS_Report.MoveNext
                    Next dika
                    .Cells(baris_pas_ngisi + 1, kolom + 10).Font.Color = &HFFFFFF
                     .Cells(baris_pas_ngisi + 1, kolom + 10).Value = totalan_salbung
                End If
                
                'NGISI TOTAL schedule call
                
                sQuery = " SELECT userid ,COALESCE(sum_schedule,0) AS ELIN FROM ("
                sQuery = sQuery + "(SELECT userid FROM usertbl WHERE spvcode = 'SPV1' and USERID = '" & agent_excel & "' AND aktif = '1' ORDER BY ID) as A LEFT JOIN"
                sQuery = sQuery + "(select agent,sum(jml_schedule) AS sum_schedule from tbl_report_dika  group  by agent order by agent ) as b ON a.userid = b.agent)"
                Set RS_Report = New ADODB.Recordset
                RS_Report.CursorLocation = adUseClient
                RS_Report.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If RS_Report.RecordCount > 0 Then
                    baris_pas_ngisi = 6
                    nilai = 0
                    totalan_schedule = 0
                    For dika = 1 To RS_Report.RecordCount
                        baris_pas_ngisi = baris_pas_ngisi + 1
                        .Cells(baris_pas_ngisi, kolom + 12).Value = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        .Cells(baris_pas_ngisi, kolom + 12).HorizontalAlignment = xlCenter
                        .Cells(baris_pas_ngisi, kolom + 12).Borders.LineStyle = xlContinuous
                        .Cells(baris_pas_ngisi, kolom + 12).Borders(xlEdgeRight).Weight = 3
                        nilai = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        
                        totalan_schedule = totalan_schedule + nilai
                        
                        RS_Report.MoveNext
                    Next dika
                    .Cells(baris_pas_ngisi + 1, kolom + 11).Font.Color = &HFFFFFF
                     .Cells(baris_pas_ngisi + 1, kolom + 11).Value = totalan_schedule
                End If
                
                'NGISI TOTAL tidak diangkat
                
                sQuery = " SELECT userid ,COALESCE(sum_tdk_diangkat,0) AS ELIN FROM ("
                sQuery = sQuery + "(SELECT userid FROM usertbl WHERE spvcode = 'SPV1' and USERID = '" & agent_excel & "' AND aktif = '1' ORDER BY ID) as A LEFT JOIN"
                sQuery = sQuery + "(select agent,sum(jml_tdk_diangkat) AS sum_tdk_diangkat from tbl_report_dika  group  by agent order by agent ) as b ON a.userid = b.agent)"
                Set RS_Report = New ADODB.Recordset
                RS_Report.CursorLocation = adUseClient
                RS_Report.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If RS_Report.RecordCount > 0 Then
                    baris_pas_ngisi = 6
                    nilai = 0
                    totalan_tdk_diangkat = 0
                    For dika = 1 To RS_Report.RecordCount
                        baris_pas_ngisi = baris_pas_ngisi + 1
                        .Cells(baris_pas_ngisi, kolom + 13).Value = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        .Cells(baris_pas_ngisi, kolom + 13).HorizontalAlignment = xlCenter
                        .Cells(baris_pas_ngisi, kolom + 13).Borders.LineStyle = xlContinuous
                        .Cells(baris_pas_ngisi, kolom + 13).Borders(xlEdgeRight).Weight = 3
                        nilai = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        
                        totalan_tdk_diangkat = totalan_tdk_diangkat + nilai
                        
                        RS_Report.MoveNext
                    Next dika
                    .Cells(baris_pas_ngisi + 1, kolom + 12).Font.Color = &HFFFFFF
                     .Cells(baris_pas_ngisi + 1, kolom + 12).Value = totalan_tdk_diangkat
                End If
                
                'NGISI TOTAL tidak ada ditempat
                
                sQuery = " SELECT userid ,COALESCE(sum_tdk_ditempat,0) AS ELIN FROM ("
                sQuery = sQuery + "(SELECT userid FROM usertbl WHERE spvcode = 'SPV1' and USERID = '" & agent_excel & "' AND aktif = '1' ORDER BY ID) as A LEFT JOIN"
                sQuery = sQuery + "(select agent,sum(jml_tdk_ditempat) AS sum_tdk_ditempat from tbl_report_dika  group  by agent order by agent ) as b ON a.userid = b.agent)"
                Set RS_Report = New ADODB.Recordset
                RS_Report.CursorLocation = adUseClient
                RS_Report.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If RS_Report.RecordCount > 0 Then
                    baris_pas_ngisi = 6
                    nilai = 0
                    totalan_tdk_ditempat = 0
                    For dika = 1 To RS_Report.RecordCount
                        baris_pas_ngisi = baris_pas_ngisi + 1
                        .Cells(baris_pas_ngisi, kolom + 14).Value = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        .Cells(baris_pas_ngisi, kolom + 14).HorizontalAlignment = xlCenter
                        .Cells(baris_pas_ngisi, kolom + 14).Borders.LineStyle = xlContinuous
                        .Cells(baris_pas_ngisi, kolom + 14).Borders(xlEdgeRight).Weight = 3
                        nilai = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        
                        totalan_tdk_ditempat = totalan_tdk_ditempat + nilai
                        
                        RS_Report.MoveNext
                    Next dika
                    .Cells(baris_pas_ngisi + 1, kolom + 13).Font.Color = &HFFFFFF
                     .Cells(baris_pas_ngisi + 1, kolom + 13).Value = totalan_tdk_ditempat
                End If
                
                'NGISI TOTAL unknown
                
                sQuery = " SELECT userid ,COALESCE(sum_unknow,0) AS ELIN FROM ("
                sQuery = sQuery + "(SELECT userid FROM usertbl WHERE spvcode = 'SPV1' and USERID = '" & agent_excel & "' AND aktif = '1' ORDER BY ID) as A LEFT JOIN"
                sQuery = sQuery + "(select agent,sum(jml_unknow) AS sum_unknow from tbl_report_dika  group  by agent order by agent ) as b ON a.userid = b.agent)"
                Set RS_Report = New ADODB.Recordset
                RS_Report.CursorLocation = adUseClient
                RS_Report.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If RS_Report.RecordCount > 0 Then
                    baris_pas_ngisi = 6
                    nilai = 0
                    totalan_unknow = 0
                    For dika = 1 To RS_Report.RecordCount
                        baris_pas_ngisi = baris_pas_ngisi + 1
                        .Cells(baris_pas_ngisi, kolom + 15).Value = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        .Cells(baris_pas_ngisi, kolom + 15).HorizontalAlignment = xlCenter
                        .Cells(baris_pas_ngisi, kolom + 15).Borders.LineStyle = xlContinuous
                        .Cells(baris_pas_ngisi, kolom + 15).Borders(xlEdgeRight).Weight = 3
                        nilai = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
                        
                        totalan_unknow = totalan_unknow + nilai
                        
                        RS_Report.MoveNext
                    Next dika
                    .Cells(baris_pas_ngisi + 1, kolom + 14).Font.Color = &HFFFFFF
                     .Cells(baris_pas_ngisi + 1, kolom + 14).Value = totalan_unknow
                End If
                
                
                 End If
        Next i
                kolom = kolom + 10
                '----TOTAL----
                .Cells(46, kolom).Value = "TOTAL"
                .Cells(46, kolom).Font.Bold = True
                .Cells(46, kolom).Font.Size = 12
                .Cells(46, kolom).Interior.Color = &H808080
                .Cells(46, kolom).Borders.LineStyle = xlContinuous
                .Cells(46, kolom).Borders.Weight = 3
                .Cells(46, kolom).ColumnWidth = 15
                ELIN2.Range(arrayAlphabet(kolom) & "46:" & arrayAlphabet(kolom) & "47").Merge
                ELIN2.Range(arrayAlphabet(kolom) & "46:" & arrayAlphabet(kolom) & "47").HorizontalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom) & "46:" & arrayAlphabet(kolom) & "47").VerticalAlignment = xlCenter
                ELIN2.Range(arrayAlphabet(kolom) & "46:" & arrayAlphabet(kolom) & "47").Interior.Color = &H808080
                ELIN2.Range(arrayAlphabet(kolom) & "46:" & arrayAlphabet(kolom) & "47").Borders.LineStyle = xlContinuous
                ELIN2.Range(arrayAlphabet(kolom) & "46:" & arrayAlphabet(kolom) & "47").Borders.Weight = 3
                
                
'                sQuery = " SELECT c.userid ,coalesce(sum_conr_pl, 0) + coalesce(sum_nar_pl, 0) + coalesce(sum_msr_pl, 0) + "
'                sQuery = sQuery + " coalesce(sum_canr_pl, 0)+coalesce(sum_conr_cc, 0) + coalesce(sum_nar_cc, 0) + coalesce(sum_msr_cc, 0) + "
'                sQuery = sQuery + " coalesce(sum_canr_cc, 0)+ coalesce(sum_cons_pl, 0) + coalesce(sum_nas_pl, 0) + coalesce(sum_mss_pl, 0) + "
'                sQuery = sQuery + " coalesce(sum_cans_pl, 0)+coalesce(sum_cons_cc, 0) + coalesce(sum_nas_cc, 0) + coalesce(sum_mss_cc, 0) + "
'                sQuery = sQuery + " coalesce(sum_cans_cc, 0)+ coalesce(sum_nons_pl, 0) + coalesce(sum_nons_cc, 0) + coalesce(sum_nonr_pl, 0) + coalesce(sum_nonr_cc, 0) + "
'
'                sQuery = sQuery + " coalesce(sum_conp_pl, 0) + coalesce(sum_nap_pl, 0) + coalesce(sum_msp_pl, 0) + "
'                sQuery = sQuery + " coalesce(sum_canp_pl, 0) + coalesce(sum_nonp_pl,0) + "
'                sQuery = sQuery + " coalesce(sum_conp_cc, 0) + coalesce(sum_nap_cc, 0) + coalesce(sum_msp_cc, 0) + "
'                sQuery = sQuery + " coalesce(sum_canp_cc, 0) + coalesce(sum_nonp_cc, 0) "
'                sQuery = sQuery + "as elin FROM ("
'                sQuery = sQuery + " (SELECT userid FROM usertbl WHERE spvcode = 'TL07' AND aktif = '1'  AND f_login_support = '1' order by agentcode) as A LEFT JOIN"
'                sQuery = sQuery + " (select agent,sum(jumlah_confirmed_regular_pl) as sum_conr_pl,sum(jumlah_na_regular_pl)as sum_nar_pl,"
'                sQuery = sQuery + " sum(jumlah_ms_regular_pl) as sum_msr_pl,sum(jumlah_can_regular_pl) as sum_canr_pl,sum(jumlah_confirmed_regular_cc) as sum_conr_cc,"
'                sQuery = sQuery + " sum(jumlah_na_regular_cc)as sum_nar_cc,sum(jumlah_ms_regular_cc) as sum_msr_cc,sum(jumlah_can_regular_cc) as sum_canr_cc,"
'                sQuery = sQuery + " sum(jumlah_confirmed_speed_pl) as sum_cons_pl,sum(jumlah_na_speed_pl)as sum_nas_pl,sum(jumlah_ms_speed_pl) as sum_mss_pl,"
'                sQuery = sQuery + " sum(jumlah_can_speed_pl) as sum_cans_pl,sum(jumlah_confirmed_speed_cc) as sum_cons_cc,sum(jumlah_na_speed_cc)as sum_nas_cc,"
'                sQuery = sQuery + " sum(jumlah_ms_speed_cc) as sum_mss_cc,sum(jumlah_can_speed_cc) as sum_cans_cc, sum(jumlah_non_speed_pl) as sum_nons_pl,"
'                sQuery = sQuery + " sum(jumlah_non_speed_cc)as sum_nons_cc,sum(jumlah_non_regular_pl) as sum_nonr_pl,sum(jumlah_non_regular_cc) as sum_nonr_cc,"
'                sQuery = sQuery + " sum(jumlah_confirmed_pa_pl) as sum_conp_pl,sum(jumlah_na_pa_pl) as sum_nap_pl,sum(jumlah_ms_pa_pl) as sum_msp_pl,"
'                sQuery = sQuery + " sum(jumlah_can_pa_pl) as sum_canp_pl,sum(jumlah_non_pa_pl) as sum_nonp_pl,"
'                sQuery = sQuery + " sum(jumlah_confirmed_pa_cc) as sum_conp_cc,sum(jumlah_na_pa_cc) as sum_nap_cc,sum(jumlah_ms_pa_cc) as sum_msp_cc,"
'                sQuery = sQuery + " sum(jumlah_can_pa_cc) as sum_canp_cc,sum(jumlah_non_pa_cc) as sum_nonp_cc from tbl_report_berto"
'                sQuery = sQuery + " group by agent order by agent ) as b ON a.userid = b.agent)as c LEFT JOIN usertbl d on c.agent=d.userid order by agentcode"
'
'                Set RS_Report = New ADODB.Recordset
'                RS_Report.CursorLocation = adUseClient
'                RS_Report.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'                If RS_Report.RecordCount > 0 Then
'                    baris_pas_ngisi = 47
'                    Nilai = 0
'                    totalan_all = 0
'                    For dika = 1 To RS_Report.RecordCount
'                        baris_pas_ngisi = baris_pas_ngisi + 1
'                        .Cells(baris_pas_ngisi, kolom).Value = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
'                        .Cells(baris_pas_ngisi, kolom).HorizontalAlignment = xlCenter
'                        .Cells(baris_pas_ngisi, kolom).Borders.LineStyle = xlContinuous
'                        .Cells(baris_pas_ngisi, kolom).Borders(xlEdgeRight).Weight = 3
'                        .Cells(baris_pas_ngisi, kolom).Font.Bold = True
'                        Nilai = IIf(IsNull(RS_Report!ELIN), "0", RS_Report!ELIN)
'
'                        totalan_all = totalan_all + Nilai
'
'                        RS_Report.MoveNext
'                    Next dika
'                    .Cells(baris_pas_ngisi + 1, kolom).Value = totalan_all
'                    .Cells(baris_pas_ngisi + 1, kolom).Font.Bold = True
'                    .Cells(baris_pas_ngisi + 1, kolom).Font.Size = 15
'                    .Cells(baris_pas_ngisi + 1, kolom).Font.Color = &HFFFFFF
'                    ELIN2.Range(arrayAlphabet(kolom) & baris_pas_ngisi + 1 & ":" & arrayAlphabet(kolom) & baris_pas_ngisi + 1).HorizontalAlignment = xlCenter
'                    ELIN2.Range(arrayAlphabet(kolom) & baris_pas_ngisi + 1 & ":" & arrayAlphabet(kolom) & baris_pas_ngisi + 1).Interior.Color = &H808080
'                    ELIN2.Range(arrayAlphabet(kolom) & baris_pas_ngisi + 1 & ":" & arrayAlphabet(kolom) & baris_pas_ngisi + 1).Font.Bold = True
'                    ELIN2.Range(arrayAlphabet(kolom) & baris_pas_ngisi + 1 & ":" & arrayAlphabet(kolom) & baris_pas_ngisi + 1).Borders.LineStyle = xlContinuous
'                    ELIN2.Range(arrayAlphabet(kolom) & baris_pas_ngisi + 1 & ":" & arrayAlphabet(kolom) & baris_pas_ngisi + 1).Borders.Weight = 3
'                End If
     End With
           
End Sub

Private Sub SSCommand2_Click()
    Unload Me
End Sub

