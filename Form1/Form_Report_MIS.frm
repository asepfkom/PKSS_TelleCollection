VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form_Report_MIS 
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
         TabIndex        =   4
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
            TabIndex        =   12
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
            TabIndex        =   11
            Top             =   3480
            Width           =   1455
         End
         Begin VB.CommandButton Cmd_Refersh3 
            BackColor       =   &H00F1E5DB&
            Caption         =   "REFRESH"
            Height          =   255
            Left            =   3960
            Style           =   1  'Graphical
            TabIndex        =   7
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
            TabIndex        =   6
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
            TabIndex        =   5
            Top             =   4440
            Width           =   1455
         End
         Begin MSComctlLib.ListView LVAgent 
            Height          =   3120
            Left            =   90
            TabIndex        =   8
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
            TabIndex        =   9
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
            TabIndex        =   10
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
         Left            =   10515
         Picture         =   "Form_Report_MIS.frx":0000
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
         Picture         =   "Form_Report_MIS.frx":0766
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
      Begin TDBDate6Ctl.TDBDate TdTglCall1 
         Height          =   315
         Left            =   1320
         TabIndex        =   13
         Top             =   315
         Width           =   1365
         _Version        =   65536
         _ExtentX        =   2408
         _ExtentY        =   556
         Calendar        =   "Form_Report_MIS.frx":0DAC
         Caption         =   "Form_Report_MIS.frx":0EC4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Form_Report_MIS.frx":0F30
         Keys            =   "Form_Report_MIS.frx":0F4E
         Spin            =   "Form_Report_MIS.frx":0FAC
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
         Left            =   3105
         TabIndex        =   14
         Top             =   315
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   556
         Calendar        =   "Form_Report_MIS.frx":0FD4
         Caption         =   "Form_Report_MIS.frx":10EC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "Form_Report_MIS.frx":1158
         Keys            =   "Form_Report_MIS.frx":1176
         Spin            =   "Form_Report_MIS.frx":11D4
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
         Left            =   2775
         TabIndex        =   16
         Top             =   300
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
         Left            =   180
         TabIndex        =   15
         Top             =   285
         Width           =   780
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
      Picture         =   "Form_Report_MIS.frx":11FC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12240
   End
End
Attribute VB_Name = "Form_Report_MIS"
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
    'tgl_tracking.Value = Now
    
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
    cQuery = cQuery + vbCrLf + " SELECT  date(tgl) as tgl_call, "
    cQuery = cQuery + vbCrLf + "  sum(CASE WHEN lastcall = 'Already Paid' then 1 else 0 end) as jml_paid,"
    cQuery = cQuery + vbCrLf + "  sum(CASE WHEN lastcall = 'BP' then 1 else 0 end) as jml_bp,"
    cQuery = cQuery + vbCrLf + "  sum(CASE WHEN lastcall = 'PTP' then 1 else 0 end) as jml_ptp,"
    cQuery = cQuery + vbCrLf + "  sum(CASE WHEN lastcall = 'Schedule Call' then 1 else 0 end) as jml_schedule,"
    cQuery = cQuery + vbCrLf + "  sum(CASE WHEN lastcall = 'Left Message' then 1 else 0 end) as jml_left_msg,"
    cQuery = cQuery + vbCrLf + "  sum(CASE WHEN lastcall = 'Negosiasi' then 1 else 0 end) as jml_nego,"
    cQuery = cQuery + vbCrLf + "  sum(CASE WHEN lastcall = 'Busy' then 1 else 0 end) as jml_busy,"
    cQuery = cQuery + vbCrLf + "  sum(CASE WHEN lastcall = 'Dead' then 1 else 0 end) as jml_dead,"
    cQuery = cQuery + vbCrLf + "  sum(CASE WHEN lastcall = 'Invalid' then 1 else 0 end) as jml_invalid,"
    cQuery = cQuery + vbCrLf + "  sum(CASE WHEN lastcall = 'Mailbox' then 1 else 0 end) as jml_mailbox,"
    cQuery = cQuery + vbCrLf + "  sum(CASE WHEN lastcall = 'Pindah Alamat' then 1 else 0 end) as jml_pndah_alamat,"
    cQuery = cQuery + vbCrLf + "  sum(CASE WHEN lastcall = 'Salah Sambung' then 1 else 0 end) as jml_salbung,"
    cQuery = cQuery + vbCrLf + "  sum(CASE WHEN lastcall = 'Tidak Ada di Tempat' then 1 else 0 end) as jml_tdk_ditempat,"
    cQuery = cQuery + vbCrLf + "  sum(CASE WHEN lastcall = 'Tidak Diangkat' then 1 else 0 end) as jml_tdk_diangkat,"
    cQuery = cQuery + vbCrLf + "  sum(CASE WHEN lastcall = 'Unknow' then 1 else 0 end) as jml_unknow,"
    cQuery = cQuery + vbCrLf + "  sum(CASE WHEN lastcall = 'Data Retur' then 1 else 0 end) as jml_data_retur"
    cQuery = cQuery + vbCrLf + " FROM ("
    cQuery = cQuery + vbCrLf + " select distinct date(tgl) as tgl,agent,lastcall,custid from mgm_hst where coalesce(lastcall,'')<>''"
    cQuery = cQuery + vbCrLf + " ) a"
    cQuery = cQuery + vbCrLf + " WHERe agent in (" + ListCustId + ")   "
    cQuery = cQuery + vbCrLf + " AND date(tgl) between '" + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' and '" + Format(TdTglCall2.Value, "yyyy-mm-dd") + "' "
    cQuery = cQuery + vbCrLf + " group by date(tgl),AGENT "
    cQuery = cQuery + vbCrLf + " order by tgl_call )"
    
    M_OBJCONN.Execute cQuery
    
    Call Isi_Excel
End Sub

Private Sub Isi_Excel()
    Dim ExlObj As Excel.Application
    Dim TGL, kolom, Baris, kolom_tgl As Integer
    Dim sQuery As String
    Dim rs_new As ADODB.Recordset
    Dim RS_Report As ADODB.Recordset
    Dim RS_total As ADODB.Recordset
    Dim tgl_excel, tgl_cek_excel, tgl_rs, agent_excel, agent_rs, speed_regular, cc_pl As String
    Dim baris_pas_ngisi, kolom_pas_ngisi, baris_sekarang As Integer
    Dim bulan_tahun_mulai_series, bulan_tahun_akhir_series, mulai_series, akhir_series As String
    Dim nilai, totalan_speed_cc, totalan_speed_pl, totalan_reg_cc, totalan_reg_pl, totalan_pa_cc, totalan_pa_pl, totalan_sebulan, totalan_sebulan1, totalan_sebulan2 As Double
    Dim totalan_confirmed_speed_cc, totalan_confirmed_speed_pl, totalan_confirmed_reg_cc, totalan_confirmed_reg_pl, totalan_confirmed_pa_cc, totalan_confirmed_pa_pl As Double
    Dim total_lm, total_sc, total_ap, total_connect, total_bp, total_ptp, total_nego, total_contact, total_b, total_d, total_in, total_mb, total_under, total_ssl As Double
    Dim total_unconnect, total_tadt, total_td, total_un, total_dr, total_grand_total As Double
    
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
    Set ExlObj = CreateObject("excel.application")
    ExlObj.Workbooks.ADD
    ExlObj.Visible = True
    
    ExlObj.Range("A1:AL1").MergeCells = True
    ExlObj.Range("A1:AL500").Interior.Color = &HFFFFFF
    
    
    With ExlObj.ActiveSheet
        .Cells(1, 1).Value = "REPORT MIS " & Format(TdTglCall1.Value, "DD-MM-YYYY") & "  -  " & Format(TdTglCall2.Value, "DD-MM-YYYY") & ""
        .Cells(1, 1).Font.Name = "Arial"
        .Cells(1, 1).Font.Size = "14"
        .Cells(1, 1).Font.Bold = True
        
        .Cells(3, 1).Value = "PT.DIKA - BCA COLLECTION"
        .Cells(3, 1).Font.Color = &HFF00FF
        .Cells(3, 1).Font.Bold = True
        ExlObj.Range("A3:C4").MergeCells = True
        ExlObj.Range("A3:C4").HorizontalAlignment = xlCenter
        ExlObj.Range("A3:C4").VerticalAlignment = xlCenter
        '------nasional----------
        .Cells(5, 1).Value = "Nasional"
        .Cells(5, 1).Font.Color = &HFFFFFF
        .Cells(5, 1).Font.Size = "12"
        .Cells(5, 1).Font.Bold = True
         ExlObj.Range("A5:C5").Interior.Color = &HC0C000
         ExlObj.Range("A5:C5").MergeCells = True
        '----------------------------------------------------------
        .Cells(7, 2).Value = "Connect"
        .Cells(7, 2).Font.Bold = True
        .Cells(7, 2).Font.Color = &H800000
         ExlObj.Range("B7:C7").MergeCells = True
         ExlObj.Range("B7:C7").Interior.Color = &HC0C0C0
        '----------------------------------------------------------
        .Cells(8, 3).Value = "Left Message"
        .Cells(8, 3).ColumnWidth = 18
        '----------------------------------------------------------
        .Cells(9, 3).Value = "Schedule Call"
        .Cells(9, 3).ColumnWidth = 18
        '----------------------------------------------------------
        '----------------------------------------------------------
        .Cells(11, 2).Value = "Contact"
        .Cells(11, 2).Font.Bold = True
        .Cells(11, 2).Font.Color = &HC0C000
         ExlObj.Range("B11:C11").MergeCells = True
         ExlObj.Range("B11:C11").Interior.Color = &HC0C0C0
        '----------------------------------------------------------
        .Cells(12, 3).Value = "Already Paid"
        '.Cells(12, 3).Borders.LineStyle = xlContinuous
        '----------------------------------------------------------
        .Cells(13, 3).Value = "BP"
        '.Cells(13, 3).Borders.LineStyle = xlContinuous
        '----------------------------------------------------------
        .Cells(14, 3).Value = "Negosiasi"
        '.Cells(14, 3).Borders.LineStyle = xlContinuous
        '----------------------------------------------------------
        .Cells(15, 3).Value = "PTP"
        '.Cells(15, 3).Borders.LineStyle = xlContinuous
        '----------------------------------------------------------
        '----------------------------------------------------------
        .Cells(17, 2).Value = "Unconnect"
        .Cells(17, 2).Font.Bold = True
        .Cells(17, 2).Font.Color = vbRed
         ExlObj.Range("B17:C17").MergeCells = True
         ExlObj.Range("B17:C17").Interior.Color = &HC0C0C0
        '----------------------------------------------------------
        .Cells(18, 3).Value = "Busy"
        '.Cells(18, 3).Borders.LineStyle = xlContinuous
        '----------------------------------------------------------
        .Cells(19, 3).Value = "Dead"
        '.Cells(15, 2).Borders.LineStyle = xlContinuous
        '----------------------------------------------------------
        .Cells(20, 3).Value = "Invalid"
        '.Cells(16, 2).Borders.LineStyle = xlContinuous
        '----------------------------------------------------------
        .Cells(21, 3).Value = "Mailbox"
        '.Cells(21, 3).Borders.LineStyle = xlContinuous
        '----------------------------------------------------------
        .Cells(22, 3).Value = "Pindah Alamat"
        '.Cells(18, 2).Borders.LineStyle = xlContinuous
        '----------------------------------------------------------
        .Cells(23, 3).Value = "Salah Sambung"
        '.Cells(19, 2).Borders.LineStyle = xlContinuous
        '----------------------------------------------------------
        .Cells(24, 3).Value = "Tidak Ada di Tempat"
        '----------------------------------------------------------
        .Cells(25, 3).Value = "Tidak Diangkat"
        '----------------------------------------------------------
        .Cells(26, 3).Value = "Unknow"
        '----------------------------------------------------------
        .Cells(27, 3).Value = "Data Retur"
        '----------------------------------------------------------
        .Cells(3, 4).Value = "MTD"
        .Cells(3, 4).Font.Color = &HFFFFFF
        .Cells(3, 4).Font.Size = "12"
        .Cells(3, 4).Font.Bold = True
         ExlObj.Range("D3:E3").Interior.Color = &HFF0000
         ExlObj.Range("D3:E3").MergeCells = True
         ExlObj.Range("D3:E3").HorizontalAlignment = xlCenter
        '----------------------------------------------------------
        .Cells(4, 4).Value = "#"
        .Cells(4, 4).Font.Color = &HFFFFFF
        .Cells(4, 4).Font.Size = "12"
        .Cells(4, 4).Font.Bold = True
        .Cells(4, 4).Interior.Color = &HFF0000
        .Cells(4, 4).HorizontalAlignment = xlCenter
        '----------------------------------------------------------
        .Cells(4, 5).Value = "%"
        .Cells(4, 5).Font.Color = &HFFFFFF
        .Cells(4, 5).Font.Size = "12"
        .Cells(4, 5).Font.Bold = True
        .Cells(4, 5).Interior.Color = &HFF0000
        .Cells(4, 5).HorizontalAlignment = xlCenter
        '----------------------------------------------------------
        Dim RSdate As ADODB.Recordset
        Dim RSdate2 As ADODB.Recordset
        Dim RSdate3 As ADODB.Recordset
        Dim maxHari As Integer
        Dim jumlah As Integer
        Dim hari As String
        Dim tgl_awal As String
        Set RSdate = New ADODB.Recordset
        RSdate.CursorLocation = adUseClient
        Set RSdate2 = New ADODB.Recordset
        RSdate2.CursorLocation = adUseClient
        Set RSdate3 = New ADODB.Recordset
        RSdate3.CursorLocation = adUseClient
                  
        Baris = 7
        kolom_region = 4
        kolom_total = 4
        kolom_hari = 4
        kolom = 6
        kolom_tgl = 6
        baris_tgl = 2
        baris_region = 0
        baris_pas_ngisi = 8
        baris_status = 8
        kolom_status = 4
        strsql = "select distinct date(tglcall) as tgl1 from mgm where date(tglcall) between '" + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' and '" + Format(TdTglCall2.Value, "yyyy-mm-dd") + "' order by date(tglcall)"
        RSdate2.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        
        total_lm = 0
        total_sc = 0
        total_connect = 0
        total_ap = 0
        total_bp = 0
        total_nego = 0
        total_ptp = 0
        total_contact = 0
        total_b = 0
        total_d = 0
        total_in = 0
        total_mb = 0
        total_pa = 0
        total_ssl = 0
        total_tadt = 0
        total_td = 0
        total_un = 0
        total_dr = 0
        total_grand_total = 0
        
        While Not RSdate2.EOF
            .Cells(4, kolom_tgl) = cnull(RSdate2!Tgl1)
            hari = .Cells(4, kolom_tgl)
            .Cells(4, kolom_tgl).ColumnWidth = 15
            ExlObj.Range(arrayAlphabet(kolom_tgl) & 3 & ":" & arrayAlphabet(kolom_tgl) & 4).Font.Bold = True
            ExlObj.Range(arrayAlphabet(kolom_tgl) & 3 & ":" & arrayAlphabet(kolom_tgl) & 4).MergeCells = True
            ExlObj.Range(arrayAlphabet(kolom_tgl) & 3 & ":" & arrayAlphabet(kolom_tgl) & 4).HorizontalAlignment = xlCenter
            ExlObj.Range(arrayAlphabet(kolom_tgl) & 3 & ":" & arrayAlphabet(kolom_tgl) & 4).VerticalAlignment = xlCenter
            ExlObj.Range(arrayAlphabet(kolom_tgl) & 3 & ":" & arrayAlphabet(kolom_tgl) & 4).Interior.Color = &HFF0000
            ExlObj.Range(arrayAlphabet(kolom_tgl) & 3 & ":" & arrayAlphabet(kolom_tgl) & 4).Font.Color = &HFFFFFF
        
        
                sQuery = "SELECT * FROM tbl_report_dika WHERE tgl_call='" + hari + "'"
                Set RS_Report = New ADODB.Recordset
                RS_Report.CursorLocation = adUseClient
                RS_Report.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                    If RS_Report.RecordCount <> 0 Then
                    
                        .Cells(baris_pas_ngisi, kolom).Value = IIf(IsNull(RS_Report!jml_left_msg), "0", RS_Report!jml_left_msg)
                        .Cells(baris_pas_ngisi, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 1, kolom).Value = IIf(IsNull(RS_Report!jml_schedule), "0", RS_Report!jml_schedule)
                        .Cells(baris_pas_ngisi + 1, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        
                        .Cells(Baris, kolom_hari + 2).Value = (.Cells(baris_pas_ngisi, kolom).Value + .Cells(baris_pas_ngisi + 1, kolom).Value)
                        .Cells(Baris, kolom_hari + 2).HorizontalAlignment = xlCenter
                        .Cells(Baris, kolom_hari + 2).Interior.Color = &HC0C0C0
                        .Cells(Baris, kolom_hari + 2).Font.Color = &H800000
                        
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 4, kolom).Value = IIf(IsNull(RS_Report!jml_paid), "0", RS_Report!jml_paid)
                        .Cells(baris_pas_ngisi + 4, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 5, kolom).Value = IIf(IsNull(RS_Report!jml_bp), "0", RS_Report!jml_bp)
                        .Cells(baris_pas_ngisi + 5, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 6, kolom).Value = IIf(IsNull(RS_Report!jml_nego), "0", RS_Report!jml_nego)
                        .Cells(baris_pas_ngisi + 6, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 7, kolom).Value = IIf(IsNull(RS_Report!jml_ptp), "0", RS_Report!jml_ptp)
                        .Cells(baris_pas_ngisi + 7, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        
                        .Cells(Baris + 4, kolom_hari + 2).Value = (.Cells(baris_pas_ngisi + 4, kolom).Value + .Cells(baris_pas_ngisi + 5, kolom).Value + .Cells(baris_pas_ngisi + 6, kolom).Value + .Cells(baris_pas_ngisi + 7, kolom).Value)
                        .Cells(Baris + 4, kolom_hari + 2).HorizontalAlignment = xlCenter
                        .Cells(Baris + 4, kolom_hari + 2).Interior.Color = &HC0C0C0
                        .Cells(Baris + 4, kolom_hari + 2).Font.Color = &HC0C000
                        .Cells(Baris + 4, kolom_hari + 2).Font.Bold = True
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 10, kolom).Value = IIf(IsNull(RS_Report!jml_busy), "0", RS_Report!jml_busy)
                        .Cells(baris_pas_ngisi + 10, kolom).HorizontalAlignment = xlCenter
                        
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 11, kolom).Value = IIf(IsNull(RS_Report!jml_dead), "0", RS_Report!jml_dead)
                        .Cells(baris_pas_ngisi + 11, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 12, kolom).Value = IIf(IsNull(RS_Report!jml_invalid), "0", RS_Report!jml_invalid)
                        .Cells(baris_pas_ngisi + 12, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 13, kolom).Value = IIf(IsNull(RS_Report!jml_mailbox), "0", RS_Report!jml_mailbox)
                        .Cells(baris_pas_ngisi + 13, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 14, kolom).Value = IIf(IsNull(RS_Report!jml_pndah_alamat), "0", RS_Report!jml_pndah_alamat)
                        .Cells(baris_pas_ngisi + 14, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 15, kolom).Value = IIf(IsNull(RS_Report!jml_salbung), "0", RS_Report!jml_salbung)
                        .Cells(baris_pas_ngisi + 15, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 16, kolom).Value = IIf(IsNull(RS_Report!jml_tdk_ditempat), "0", RS_Report!jml_tdk_ditempat)
                        .Cells(baris_pas_ngisi + 16, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 17, kolom).Value = IIf(IsNull(RS_Report!jml_tdk_diangkat), "0", RS_Report!jml_tdk_diangkat)
                        .Cells(baris_pas_ngisi + 17, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 18, kolom).Value = IIf(IsNull(RS_Report!jml_unknow), "0", RS_Report!jml_unknow)
                        .Cells(baris_pas_ngisi + 18, kolom).HorizontalAlignment = xlCenter
                         '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 19, kolom).Value = IIf(IsNull(RS_Report!jml_data_retur), "0", RS_Report!jml_data_retur)
                        .Cells(baris_pas_ngisi + 19, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        
                        .Cells(Baris + 10, kolom_hari + 2).Value = (.Cells(baris_pas_ngisi + 10, kolom).Value + .Cells(baris_pas_ngisi + 11, kolom).Value + .Cells(baris_pas_ngisi + 12, kolom).Value + .Cells(baris_pas_ngisi + 13, kolom).Value + .Cells(baris_pas_ngisi + 14, kolom).Value + .Cells(baris_pas_ngisi + 15, kolom).Value + .Cells(baris_pas_ngisi + 16, kolom).Value + .Cells(baris_pas_ngisi + 17, kolom).Value + .Cells(baris_pas_ngisi + 18, kolom).Value + .Cells(baris_pas_ngisi + 19, kolom).Value)
                        .Cells(Baris + 10, kolom_hari + 2).HorizontalAlignment = xlCenter
                        .Cells(Baris + 10, kolom_hari + 2).Interior.Color = &HC0C0C0
                        .Cells(Baris + 10, kolom_hari + 2).Font.Color = vbRed
                        '---------------------------------------------------------------------
                        
                        .Cells(Baris - 2, kolom_region + 2).Value = (.Cells(Baris, kolom_hari + 2).Value + .Cells(Baris + 4, kolom_hari + 2).Value + .Cells(Baris + 10, kolom_hari + 2).Value)
                        .Cells(Baris - 2, kolom_region + 2).HorizontalAlignment = xlCenter
                        .Cells(Baris - 2, kolom_region + 2).Interior.Color = &HC0C000
                        .Cells(Baris - 2, kolom_region + 2).Font.Color = &HFFFFFF
                        .Cells(Baris - 2, kolom_region + 2).Font.Size = "12"
                        .Cells(Baris - 2, kolom_region + 2).Font.Bold = True
                    
                    End If
                        total_lm = total_lm + (.Cells(baris_pas_ngisi, kolom).Value)
                        .Cells(baris_status, kolom_status).Value = total_lm
                        .Cells(baris_status, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status, kolom_status).Font.Color = &H800000
                        
                        total_sc = total_sc + (.Cells(baris_pas_ngisi + 1, kolom).Value)
                        .Cells(baris_status + 1, kolom_status).Value = total_sc
                        .Cells(baris_status + 1, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 1, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 1, kolom_status).Font.Color = &H800000
                    '---------------------------total connect------------------------------------------
                        If .Cells(Baris, kolom_total).Value = "" Then
                            .Cells(Baris, kolom_total).Value = (.Cells(baris_pas_ngisi, kolom).Value + .Cells(baris_pas_ngisi + 1, kolom).Value)
                        Else
                            .Cells(Baris, kolom_total).Value = (.Cells(baris_pas_ngisi, kolom).Value + .Cells(baris_pas_ngisi + 1, kolom).Value)
                        End If
                        .Cells(Baris, kolom_total).HorizontalAlignment = xlCenter
                        .Cells(Baris, kolom_total).Interior.Color = &HC0C0C0
                        .Cells(Baris, kolom_total).Font.Color = &H800000
                        
                        total_connect = total_connect + (.Cells(Baris, kolom_total).Value)
                                                
                        .Cells(Baris, kolom_total).Value = total_connect
                        
                        '---------------------------------------------------------------------
                        total_ap = total_ap + (.Cells(baris_pas_ngisi + 4, kolom).Value)
                        .Cells(baris_status + 4, kolom_status).Value = total_ap
                        .Cells(baris_status + 4, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 4, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 4, kolom_status).Font.Color = &HC0C000
                        
                        total_bp = total_bp + (.Cells(baris_pas_ngisi + 5, kolom).Value)
                        .Cells(baris_status + 5, kolom_status).Value = total_bp
                        .Cells(baris_status + 5, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 5, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 5, kolom_status).Font.Color = &HC0C000
                        
                        total_nego = total_nego + (.Cells(baris_pas_ngisi + 6, kolom).Value)
                        .Cells(baris_status + 6, kolom_status).Value = total_nego
                        .Cells(baris_status + 6, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 6, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 6, kolom_status).Font.Color = &HC0C000
                        
                        total_ptp = total_ptp + (.Cells(baris_pas_ngisi + 7, kolom).Value)
                        .Cells(baris_status + 7, kolom_status).Value = total_ptp
                        .Cells(baris_status + 7, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 7, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 7, kolom_status).Font.Color = &HC0C000
                        '-----------------------total contact-----------------------------------------------
                        .Cells(Baris + 4, kolom_total).Value = (.Cells(baris_pas_ngisi + 4, kolom).Value + .Cells(baris_pas_ngisi + 5, kolom).Value + .Cells(baris_pas_ngisi + 6, kolom).Value + .Cells(baris_pas_ngisi + 7, kolom).Value)
                        .Cells(Baris + 4, kolom_total).HorizontalAlignment = xlCenter
                        .Cells(Baris + 4, kolom_total).Interior.Color = &HC0C0C0
                        .Cells(Baris + 4, kolom_total).Font.Color = &HC0C000
                        .Cells(Baris + 4, kolom_total).Font.Bold = True
                        
                        total_contact = total_contact + (.Cells(Baris + 4, kolom_total).Value)
                        
                        .Cells(Baris + 4, kolom_total).Value = total_contact
                        '---------------------------------------------------------------------
                        total_b = total_b + (.Cells(baris_pas_ngisi + 10, kolom).Value)
                        .Cells(baris_status + 10, kolom_status).Value = total_b
                        .Cells(baris_status + 10, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 10, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 10, kolom_status).Font.Color = vbRed
                        
                        total_d = total_d + (.Cells(baris_pas_ngisi + 11, kolom).Value)
                        .Cells(baris_status + 11, kolom_status).Value = total_d
                        .Cells(baris_status + 11, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 11, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 11, kolom_status).Font.Color = vbRed
                        
                        total_in = total_in + (.Cells(baris_pas_ngisi + 12, kolom).Value)
                        .Cells(baris_status + 12, kolom_status).Value = total_in
                        .Cells(baris_status + 12, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 12, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 12, kolom_status).Font.Color = vbRed
                        
                        total_mb = total_mb + (.Cells(baris_pas_ngisi + 13, kolom).Value)
                        .Cells(baris_status + 13, kolom_status).Value = total_mb
                        .Cells(baris_status + 13, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 13, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 13, kolom_status).Font.Color = vbRed
                        
                        total_pa = total_pa + (.Cells(baris_pas_ngisi + 14, kolom).Value)
                        .Cells(baris_status + 14, kolom_status).Value = total_pa
                        .Cells(baris_status + 14, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 14, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 14, kolom_status).Font.Color = vbRed
                        
                        total_ssl = total_ssl + (.Cells(baris_pas_ngisi + 15, kolom).Value)
                        .Cells(baris_status + 15, kolom_status).Value = total_ssl
                        .Cells(baris_status + 15, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 15, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 15, kolom_status).Font.Color = vbRed
                        
                        total_tadt = total_tadt + (.Cells(baris_pas_ngisi + 16, kolom).Value)
                        .Cells(baris_status + 16, kolom_status).Value = total_tadt
                        .Cells(baris_status + 16, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 16, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 16, kolom_status).Font.Color = vbRed
                        
                        total_td = total_td + (.Cells(baris_pas_ngisi + 17, kolom).Value)
                        .Cells(baris_status + 17, kolom_status).Value = total_td
                        .Cells(baris_status + 17, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 17, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 17, kolom_status).Font.Color = vbRed
                        
                        total_un = total_un + (.Cells(baris_pas_ngisi + 18, kolom).Value)
                        .Cells(baris_status + 18, kolom_status).Value = total_un
                        .Cells(baris_status + 18, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 18, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 18, kolom_status).Font.Color = vbRed
                        
                        total_dr = total_dr + (.Cells(baris_pas_ngisi + 19, kolom).Value)
                        .Cells(baris_status + 19, kolom_status).Value = total_dr
                        .Cells(baris_status + 19, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 19, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 19, kolom_status).Font.Color = vbRed
                        '------------------------total unconnect---------------------------------------------
                        .Cells(Baris + 10, kolom_total).Value = (.Cells(baris_pas_ngisi + 10, kolom).Value + .Cells(baris_pas_ngisi + 11, kolom).Value + .Cells(baris_pas_ngisi + 12, kolom).Value + .Cells(baris_pas_ngisi + 13, kolom).Value + .Cells(baris_pas_ngisi + 14, kolom).Value + .Cells(baris_pas_ngisi + 15, kolom).Value + .Cells(baris_pas_ngisi + 16, kolom).Value + .Cells(baris_pas_ngisi + 17, kolom).Value + .Cells(baris_pas_ngisi + 18, kolom).Value + .Cells(baris_pas_ngisi + 19, kolom).Value)
                        .Cells(Baris + 10, kolom_total).HorizontalAlignment = xlCenter
                        .Cells(Baris + 10, kolom_total).Interior.Color = &HC0C0C0
                        .Cells(Baris + 10, kolom_total).Font.Color = vbRed
                        
                        total_unconnect = total_unconnect + (.Cells(Baris + 10, kolom_total).Value)
                        
                        .Cells(Baris + 10, kolom_total).Value = total_unconnect
                        
            baris_tgl = kolom
            kolom_tgl = kolom_tgl + 1
            kolom_region = kolom_region + 1
            kolom_hari = kolom_hari + 1
            kolom = kolom + 1
            
            RSdate2.MoveNext
                
                
        Wend
        Dim persen_connect, persen_contact, persen_unconnect, total_persen_all As Double
        Dim baris_persen, kolom_persen, nasional As Integer
        
        kolom_region = 4
        baris_region = 5
        kolom_persen = 5
        baris_persen = 5
        .Cells(baris_region, kolom_region).Value = (.Cells(Baris, kolom_total).Value + .Cells(Baris + 4, kolom_total).Value + .Cells(Baris + 10, kolom_total).Value)
        nasional = (.Cells(baris_region, kolom_region).Value)
        .Cells(baris_region, kolom_region).HorizontalAlignment = xlCenter
        .Cells(baris_region, kolom_region).Interior.Color = &HC0C000
        .Cells(baris_region, kolom_region).Font.Color = &HFFFFFF
        .Cells(baris_region, kolom_region).Font.Size = "12"
        .Cells(baris_region, kolom_region).Font.Bold = True
        '----------------------------------------------------------------
        persen_connect = ((.Cells(Baris, kolom_total).Value / .Cells(baris_region, kolom_region).Value) * 100)
        If persen_connect = 0 Then
        .Cells(baris_persen + 2, kolom_persen).Value = "0%"
        Else
        .Cells(baris_persen + 2, kolom_persen).Value = Format(persen_connect, "#.##") + "%"
        End If
        .Cells(baris_persen + 2, kolom_persen).HorizontalAlignment = xlCenter
        .Cells(baris_persen + 2, kolom_persen).Interior.Color = &HC0C0C0
        .Cells(baris_persen + 2, kolom_persen).Font.Color = &H800000
        '----------------------------------------------------------------
        persen_contact = ((.Cells(Baris + 4, kolom_total).Value / .Cells(baris_region, kolom_region).Value) * 100)
        If persen_contact = 0 Then
            .Cells(baris_persen + 6, kolom_persen).Value = "0%"
        Else
            .Cells(baris_persen + 6, kolom_persen).Value = Format(persen_contact, "#.##") + "%"
        End If
        .Cells(baris_persen + 6, kolom_persen).HorizontalAlignment = xlCenter
        .Cells(baris_persen + 6, kolom_persen).Interior.Color = &HC0C0C0
        .Cells(baris_persen + 6, kolom_persen).Font.Color = &HC0C000
        .Cells(baris_persen + 6, kolom_persen).Font.Bold = True
        '----------------------------------------------------------------
        persen_unconnect = ((.Cells(Baris + 10, kolom_total).Value / .Cells(baris_region, kolom_region).Value) * 100)
        If persen_unconnect = 0 Then
            .Cells(baris_persen + 12, kolom_persen).Value = "0%"
        Else
            .Cells(baris_persen + 12, kolom_persen).Value = Format(persen_unconnect, "#.##") + "%"
        End If
        .Cells(baris_persen + 12, kolom_persen).HorizontalAlignment = xlCenter
        .Cells(baris_persen + 12, kolom_persen).Interior.Color = &HC0C0C0
        .Cells(baris_persen + 12, kolom_persen).Font.Color = vbRed
        '----------------------------------------------------------------
        total_persen_all = (persen_connect + persen_contact + persen_unconnect)
        .Cells(baris_region, kolom_region + 1).Value = CStr(total_persen_all) + "%"
        .Cells(baris_region, kolom_region + 1).HorizontalAlignment = xlCenter
        .Cells(baris_region, kolom_region + 1).Interior.Color = &HC0C000
        .Cells(baris_region, kolom_region + 1).Font.Color = &HFFFFFF
        .Cells(baris_region, kolom_region + 1).Font.Size = "12"
        .Cells(baris_region, kolom_region + 1).Font.Bold = True
        
        ExlObj.Range(arrayAlphabet(2) & "6:" & arrayAlphabet(kolom - 1) & baris_pas_ngisi + 20).Borders(xlInsideHorizontal).LineStyle = xlDash
        
        Call query_region(ExlObj)
        
        Set ExlObj = Nothing
    End With

End Sub
Private Sub query_region(ELIN2 As Excel.Application)
Dim strsql As String
Dim listagent, region As String
Dim baris_baru As Integer
Dim rs_region As ADODB.Recordset
Dim rs_region2 As ADODB.Recordset
Dim rs_region3 As ADODB.Recordset
Dim total_lm, total_sc, total_ap, total_connect, total_bp, total_ptp, total_nego, total_contact, total_b, total_d, total_in, total_mb, total_under, total_ssl As Double
Dim total_unconnect, total_tadt, total_td, total_un, total_dr, total_grand_total As Double
    
arrayAlphabet = Array("", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", _
    "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", _
    "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ", _
    "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV", "CW", "CX", "CY", "CZ", _
    "DA", "DB", "DC", "DD", "DE", "DF", "DG", "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP", "DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", "DY", "DZ", _
    "EA", "EB", "EC", "ED", "EE", "EF", "EG", "EH", "EI", "EJ", "EK", "EL", "EM", "EN", "EO", "EP", "EQ", "ER", "ES", "ET", "EU", "EV", "EW", "EX", "EY", "EZ", _
    "FA", "FB", "FC", "FD", "FE", "FF", "FG", "FH", "FI", "FJ", "FK", "FL", "FM", "FN", "FO", "FP", "FQ", "FR", "FS", "FT", "FU", "FV", "FW", "FX", "FY", "FZ")

    For i = 1 To txt_jumlah_acc.Text
        If LVAgent.ListItems(i).Checked = True Then
            ListCustId = ListCustId & "'" & LVAgent.ListItems(i).SubItems(1) & "',"
        End If
    Next i
    
    ListCustId = Mid(ListCustId, 1, Len(ListCustId) - 1)
    listagent = ListCustId
    
    strsql = " DROP TABLE IF EXISTS tbl_report_dika_by_region "
    M_OBJCONN.Execute strsql
    
    strsql = " CREATE TABLE tbl_report_dika_by_region AS ( "
    strsql = strsql + vbCrLf + " SELECT  REGION,date(tgl) as tgl_call,"
    strsql = strsql + vbCrLf + " sum(CASE WHEN lastcall = 'Already Paid' then 1 else 0 end) as jml_paid,"
    strsql = strsql + vbCrLf + " sum(CASE WHEN lastcall = 'BP' then 1 else 0 end) as jml_bp,"
    strsql = strsql + vbCrLf + " sum(CASE WHEN lastcall = 'PTP' then 1 else 0 end) as jml_ptp,"
    strsql = strsql + vbCrLf + " sum(CASE WHEN lastcall = 'Schedule Call' then 1 else 0 end) as jml_schedule,"
    strsql = strsql + vbCrLf + " sum(CASE WHEN lastcall = 'Left Message' then 1 else 0 end) as jml_left_msg,"
    strsql = strsql + vbCrLf + " sum(CASE WHEN lastcall = 'Negosiasi' then 1 else 0 end) as jml_nego,"
    strsql = strsql + vbCrLf + " sum(CASE WHEN lastcall = 'Busy' then 1 else 0 end) as jml_busy,"
    strsql = strsql + vbCrLf + " sum(CASE WHEN lastcall = 'Dead' then 1 else 0 end) as jml_dead,"
    strsql = strsql + vbCrLf + " sum(CASE WHEN lastcall = 'Invalid' then 1 else 0 end) as jml_invalid,"
    strsql = strsql + vbCrLf + " sum(CASE WHEN lastcall = 'Mailbox' then 1 else 0 end) as jml_mailbox,"
    strsql = strsql + vbCrLf + " sum(CASE WHEN lastcall = 'Pindah Alamat' then 1 else 0 end) as jml_pndah_alamat,"
    strsql = strsql + vbCrLf + " sum(CASE WHEN lastcall = 'Salah Sambung' then 1 else 0 end) as jml_salbung,"
    strsql = strsql + vbCrLf + " sum(CASE WHEN lastcall = 'Tidak Ada di Tempat' then 1 else 0 end) as jml_tdk_ditempat,"
    strsql = strsql + vbCrLf + " sum(CASE WHEN lastcall = 'Tidak Diangkat' then 1 else 0 end) as jml_tdk_diangkat,"
    strsql = strsql + vbCrLf + " sum(CASE WHEN lastcall = 'Unknow' then 1 else 0 end) as jml_unknow,"
    strsql = strsql + vbCrLf + " sum(CASE WHEN lastcall = 'Data Retur' then 1 else 0 end) as jml_data_retur"
    strsql = strsql + vbCrLf + " FROM ("
    strsql = strsql + vbCrLf + " select distinct date(tgl) as tgl,a.agent,lastcall,a.custid,b.REGION from mgm_hst a left join mgm b on (a.custid=b.custid) where coalesce(lastcall,'')<>'' order by date(tgl)"
    strsql = strsql + vbCrLf + " ) a"
    strsql = strsql + vbCrLf + " WHERe agent in (" + listagent + ")"
    strsql = strsql + vbCrLf + " AND date(tgl) between '" + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' and '" + Format(TdTglCall2.Value, "yyyy-mm-dd") + "' "
    strsql = strsql + vbCrLf + " group by region,date(tgl)"
    strsql = strsql + vbCrLf + " order by region,tgl_call )"
    
    M_OBJCONN.Execute strsql
    
        Baris = 31
        kolom_region = 4
        kolom_total = 4
        kolom_hari = 4
        kolom = 6
        kolom_tgl = 6
        baris_tgl = 2
        baris_region = 0
        baris_pas_ngisi = 32
        baris_status = 32
        kolom_status = 4
        baris_region_r = 29
        kolom_region_r = 4
        baris_persen_r = 29
        kolom_persen_r = 5
    strsql = "SELECT distinct region FROM tbl_report_dika_by_region order by region "
    Set rs_region = New ADODB.Recordset
    rs_region.CursorLocation = adUseClient
    rs_region.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    With ELIN2.ActiveSheet
        baris_baru = 29
        
        While Not rs_region.EOF
                    
            '------by region----------
            .Cells(baris_baru, 1).Value = cnull(rs_region!region)
            region = .Cells(baris_baru, 1).Value
            .Cells(baris_baru, 1).Font.Color = &HFFFFFF
            .Cells(baris_baru, 1).Font.Size = "12"
            .Cells(baris_baru, 1).Font.Bold = True
             ELIN2.Range(arrayAlphabet(1) & baris_baru & ":" & arrayAlphabet(3) & baris_baru).Interior.Color = &HC0C000
             ELIN2.Range(arrayAlphabet(1) & baris_baru & ":" & arrayAlphabet(3) & baris_baru).MergeCells = True
            '----------------------------------------------------------
            .Cells(baris_baru + 2, 2).Value = "Connect"
            .Cells(baris_baru + 2, 2).Font.Bold = True
            .Cells(baris_baru + 2, 2).Font.Color = &H800000
             ELIN2.Range(arrayAlphabet(2) & baris_baru + 2 & ":" & arrayAlphabet(3) & baris_baru + 2).MergeCells = True
             ELIN2.Range(arrayAlphabet(2) & baris_baru + 2 & ":" & arrayAlphabet(3) & baris_baru + 2).Interior.Color = &HC0C0C0
            '----------------------------------------------------------
            .Cells(baris_baru + 3, 3).Value = "Left Message"
            .Cells(baris_baru + 3, 3).ColumnWidth = 18
            '----------------------------------------------------------
            .Cells(baris_baru + 4, 3).Value = "Schedule Call"
            .Cells(baris_baru + 4, 3).ColumnWidth = 18
            '----------------------------------------------------------
            '----------------------------------------------------------
            .Cells(baris_baru + 6, 2).Value = "Contact"
            .Cells(baris_baru + 6, 2).Font.Bold = True
            .Cells(baris_baru + 6, 2).Font.Color = &HC0C000
            .Cells(baris_baru + 6, 2).Font.Bold = True
             ELIN2.Range(arrayAlphabet(2) & baris_baru + 6 & ":" & arrayAlphabet(3) & baris_baru + 6).MergeCells = True
             ELIN2.Range(arrayAlphabet(2) & baris_baru + 6 & ":" & arrayAlphabet(3) & baris_baru + 6).Interior.Color = &HC0C0C0
            '----------------------------------------------------------
            .Cells(baris_baru + 7, 3).Value = "Already Paid"
            '.Cells(12, 3).Borders.LineStyle = xlContinuous
            '----------------------------------------------------------
            .Cells(baris_baru + 8, 3).Value = "BP"
            '.Cells(13, 3).Borders.LineStyle = xlContinuous
            '----------------------------------------------------------
            .Cells(baris_baru + 9, 3).Value = "Negosiasi"
            '.Cells(14, 3).Borders.LineStyle = xlContinuous
            '----------------------------------------------------------
            .Cells(baris_baru + 10, 3).Value = "PTP"
            '.Cells(15, 3).Borders.LineStyle = xlContinuous
            '----------------------------------------------------------
            '----------------------------------------------------------
            .Cells(baris_baru + 12, 2).Value = "Unconnect"
            .Cells(baris_baru + 12, 2).Font.Bold = True
            .Cells(baris_baru + 12, 2).Font.Color = vbRed
             ELIN2.Range(arrayAlphabet(2) & baris_baru + 12 & ":" & arrayAlphabet(3) & baris_baru + 12).MergeCells = True
             ELIN2.Range(arrayAlphabet(2) & baris_baru + 12 & ":" & arrayAlphabet(3) & baris_baru + 12).Interior.Color = &HC0C0C0
            '----------------------------------------------------------
            .Cells(baris_baru + 13, 3).Value = "Busy"
            '.Cells(18, 3).Borders.LineStyle = xlContinuous
            '----------------------------------------------------------
            .Cells(baris_baru + 14, 3).Value = "Dead"
            '.Cells(15, 2).Borders.LineStyle = xlContinuous
            '----------------------------------------------------------
            .Cells(baris_baru + 15, 3).Value = "Invalid"
            '.Cells(16, 2).Borders.LineStyle = xlContinuous
            '----------------------------------------------------------
            .Cells(baris_baru + 16, 3).Value = "Mailbox"
            '.Cells(21, 3).Borders.LineStyle = xlContinuous
            '----------------------------------------------------------
            .Cells(baris_baru + 17, 3).Value = "Pindah Alamat"
            '.Cells(18, 2).Borders.LineStyle = xlContinuous
            '----------------------------------------------------------
            .Cells(baris_baru + 18, 3).Value = "Salah Sambung"
            '.Cells(19, 2).Borders.LineStyle = xlContinuous
            '----------------------------------------------------------
            .Cells(baris_baru + 19, 3).Value = "Tidak Ada di Tempat"
            '----------------------------------------------------------
            .Cells(baris_baru + 20, 3).Value = "Tidak Diangkat"
            '----------------------------------------------------------
            .Cells(baris_baru + 21, 3).Value = "Unknow"
            '----------------------------------------------------------
            .Cells(baris_baru + 22, 3).Value = "Data Retur"
            '----------------------------------------------------------
            
        strsql = "select distinct date(tglcall) as tgl1 from mgm where date(tglcall) between '" + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' and '" + Format(TdTglCall2.Value, "yyyy-mm-dd") + "' order by date(tglcall)"
        Set rs_region2 = New ADODB.Recordset
        rs_region2.CursorLocation = adUseClient
        rs_region2.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        
        total_lm = 0
        total_sc = 0
        total_connect = 0
        total_ap = 0
        total_bp = 0
        total_nego = 0
        total_ptp = 0
        total_contact = 0
        total_b = 0
        total_d = 0
        total_in = 0
        total_mb = 0
        total_pa = 0
        total_ssl = 0
        total_tadt = 0
        total_td = 0
        total_un = 0
        total_dr = 0
        total_unconnect = 0
        total_grand_total = 0
        
        While Not rs_region2.EOF
            hari = cnull(rs_region2!Tgl1)
        
                sQuery = "SELECT * FROM tbl_report_dika_by_region WHERE tgl_call='" + hari + "' and region='" + region + "'"
                Set rs_region3 = New ADODB.Recordset
                rs_region3.CursorLocation = adUseClient
                rs_region3.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If rs_region3.RecordCount <> 0 Then
                    
                        .Cells(baris_pas_ngisi, kolom).Value = IIf(IsNull(rs_region3!jml_left_msg), "0", rs_region3!jml_left_msg)
                        .Cells(baris_pas_ngisi, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 1, kolom).Value = IIf(IsNull(rs_region3!jml_schedule), "0", rs_region3!jml_schedule)
                        .Cells(baris_pas_ngisi + 1, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        
                        .Cells(Baris, kolom_hari + 2).Value = (.Cells(baris_pas_ngisi, kolom).Value + .Cells(baris_pas_ngisi + 1, kolom).Value)
                        .Cells(Baris, kolom_hari + 2).HorizontalAlignment = xlCenter
                        .Cells(Baris, kolom_hari + 2).Interior.Color = &HC0C0C0
                        .Cells(Baris, kolom_hari + 2).Font.Color = &H800000
                        
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 4, kolom).Value = IIf(IsNull(rs_region3!jml_paid), "0", rs_region3!jml_paid)
                        .Cells(baris_pas_ngisi + 4, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 5, kolom).Value = IIf(IsNull(rs_region3!jml_bp), "0", rs_region3!jml_bp)
                        .Cells(baris_pas_ngisi + 5, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 6, kolom).Value = IIf(IsNull(rs_region3!jml_nego), "0", rs_region3!jml_nego)
                        .Cells(baris_pas_ngisi + 6, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 7, kolom).Value = IIf(IsNull(rs_region3!jml_ptp), "0", rs_region3!jml_ptp)
                        .Cells(baris_pas_ngisi + 7, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        
                        .Cells(Baris + 4, kolom_hari + 2).Value = (.Cells(baris_pas_ngisi + 4, kolom).Value + .Cells(baris_pas_ngisi + 5, kolom).Value + .Cells(baris_pas_ngisi + 6, kolom).Value + .Cells(baris_pas_ngisi + 7, kolom).Value)
                        .Cells(Baris + 4, kolom_hari + 2).HorizontalAlignment = xlCenter
                        .Cells(Baris + 4, kolom_hari + 2).Interior.Color = &HC0C0C0
                        .Cells(Baris + 4, kolom_hari + 2).Font.Color = &HC0C000
                        .Cells(Baris + 4, kolom_hari + 2).Font.Bold = True
                        
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 10, kolom).Value = IIf(IsNull(rs_region3!jml_busy), "0", rs_region3!jml_busy)
                        .Cells(baris_pas_ngisi + 10, kolom).HorizontalAlignment = xlCenter
                        
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 11, kolom).Value = IIf(IsNull(rs_region3!jml_dead), "0", rs_region3!jml_dead)
                        .Cells(baris_pas_ngisi + 11, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 12, kolom).Value = IIf(IsNull(rs_region3!jml_invalid), "0", rs_region3!jml_invalid)
                        .Cells(baris_pas_ngisi + 12, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 13, kolom).Value = IIf(IsNull(rs_region3!jml_mailbox), "0", rs_region3!jml_mailbox)
                        .Cells(baris_pas_ngisi + 13, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 14, kolom).Value = IIf(IsNull(rs_region3!jml_pndah_alamat), "0", rs_region3!jml_pndah_alamat)
                        .Cells(baris_pas_ngisi + 14, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 15, kolom).Value = IIf(IsNull(rs_region3!jml_salbung), "0", rs_region3!jml_salbung)
                        .Cells(baris_pas_ngisi + 15, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 16, kolom).Value = IIf(IsNull(rs_region3!jml_tdk_ditempat), "0", rs_region3!jml_tdk_ditempat)
                        .Cells(baris_pas_ngisi + 16, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 17, kolom).Value = IIf(IsNull(rs_region3!jml_tdk_diangkat), "0", rs_region3!jml_tdk_diangkat)
                        .Cells(baris_pas_ngisi + 17, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 18, kolom).Value = IIf(IsNull(rs_region3!jml_unknow), "0", rs_region3!jml_unknow)
                        .Cells(baris_pas_ngisi + 18, kolom).HorizontalAlignment = xlCenter
                         '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 19, kolom).Value = IIf(IsNull(rs_region3!jml_data_retur), "0", rs_region3!jml_data_retur)
                        .Cells(baris_pas_ngisi + 19, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        
                        .Cells(Baris + 10, kolom_hari + 2).Value = (.Cells(baris_pas_ngisi + 10, kolom).Value + .Cells(baris_pas_ngisi + 11, kolom).Value + .Cells(baris_pas_ngisi + 12, kolom).Value + .Cells(baris_pas_ngisi + 13, kolom).Value + .Cells(baris_pas_ngisi + 14, kolom).Value + .Cells(baris_pas_ngisi + 15, kolom).Value + .Cells(baris_pas_ngisi + 16, kolom).Value + .Cells(baris_pas_ngisi + 17, kolom).Value + .Cells(baris_pas_ngisi + 18, kolom).Value + .Cells(baris_pas_ngisi + 19, kolom).Value)
                        .Cells(Baris + 10, kolom_hari + 2).HorizontalAlignment = xlCenter
                        .Cells(Baris + 10, kolom_hari + 2).Interior.Color = &HC0C0C0
                        .Cells(Baris + 10, kolom_hari + 2).Font.Color = vbRed
                        '---------------------------------------------------------------------
                        
                        .Cells(Baris - 2, kolom_region + 2).Value = (.Cells(Baris, kolom_hari + 2).Value + .Cells(Baris + 4, kolom_hari + 2).Value + .Cells(Baris + 10, kolom_hari + 2).Value)
                        .Cells(Baris - 2, kolom_region + 2).HorizontalAlignment = xlCenter
                        .Cells(Baris - 2, kolom_region + 2).Interior.Color = &HC0C000
                        .Cells(Baris - 2, kolom_region + 2).Font.Color = &HFFFFFF
                        .Cells(Baris - 2, kolom_region + 2).Font.Size = "12"
                        .Cells(Baris - 2, kolom_region + 2).Font.Bold = True
                        
                ElseIf rs_region3.RecordCount = 0 Then
                    
                        .Cells(baris_pas_ngisi, kolom).Value = "0"
                        .Cells(baris_pas_ngisi, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 1, kolom).Value = "0"
                        .Cells(baris_pas_ngisi + 1, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        
                        .Cells(Baris, kolom_hari + 2).Value = "0"
                        .Cells(Baris, kolom_hari + 2).HorizontalAlignment = xlCenter
                        .Cells(Baris, kolom_hari + 2).Interior.Color = &HC0C0C0
                        .Cells(Baris, kolom_hari + 2).Font.Color = &H800000
                        
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 4, kolom).Value = "0"
                        .Cells(baris_pas_ngisi + 4, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 5, kolom).Value = "0"
                        .Cells(baris_pas_ngisi + 5, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 6, kolom).Value = "0"
                        .Cells(baris_pas_ngisi + 6, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 7, kolom).Value = "0"
                        .Cells(baris_pas_ngisi + 7, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        
                        .Cells(Baris + 4, kolom_hari + 2).Value = "0"
                        .Cells(Baris + 4, kolom_hari + 2).HorizontalAlignment = xlCenter
                        .Cells(Baris + 4, kolom_hari + 2).Interior.Color = &HC0C0C0
                        .Cells(Baris + 4, kolom_hari + 2).Font.Color = &HC0C000
                        .Cells(Baris + 4, kolom_hari + 2).Font.Bold = True
                        
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 10, kolom).Value = "0"
                        .Cells(baris_pas_ngisi + 10, kolom).HorizontalAlignment = xlCenter
                        
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 11, kolom).Value = "0"
                        .Cells(baris_pas_ngisi + 11, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 12, kolom).Value = "0"
                        .Cells(baris_pas_ngisi + 12, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 13, kolom).Value = "0"
                        .Cells(baris_pas_ngisi + 13, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 14, kolom).Value = "0"
                        .Cells(baris_pas_ngisi + 14, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 15, kolom).Value = "0"
                        .Cells(baris_pas_ngisi + 15, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 16, kolom).Value = "0"
                        .Cells(baris_pas_ngisi + 16, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 17, kolom).Value = "0"
                        .Cells(baris_pas_ngisi + 17, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 18, kolom).Value = "0"
                        .Cells(baris_pas_ngisi + 18, kolom).HorizontalAlignment = xlCenter
                         '---------------------------------------------------------------------
                        .Cells(baris_pas_ngisi + 19, kolom).Value = "0"
                        .Cells(baris_pas_ngisi + 19, kolom).HorizontalAlignment = xlCenter
                        '---------------------------------------------------------------------
                        
                        .Cells(Baris + 10, kolom_hari + 2).Value = "0"
                        .Cells(Baris + 10, kolom_hari + 2).HorizontalAlignment = xlCenter
                        .Cells(Baris + 10, kolom_hari + 2).Interior.Color = &HC0C0C0
                        .Cells(Baris + 10, kolom_hari + 2).Font.Color = vbRed
                        '---------------------------------------------------------------------
                        
                        .Cells(Baris - 2, kolom_region + 2).Value = "0"
                        .Cells(Baris - 2, kolom_region + 2).HorizontalAlignment = xlCenter
                        .Cells(Baris - 2, kolom_region + 2).Interior.Color = &HC0C000
                        .Cells(Baris - 2, kolom_region + 2).Font.Color = &HFFFFFF
                        .Cells(Baris - 2, kolom_region + 2).Font.Size = "12"
                        .Cells(Baris - 2, kolom_region + 2).Font.Bold = True
                        
                End If
                        total_lm = total_lm + (.Cells(baris_pas_ngisi, kolom).Value)
                        .Cells(baris_status, kolom_status).Value = total_lm
                        .Cells(baris_status, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status, kolom_status).Font.Color = &H800000
                        
                        total_sc = total_sc + (.Cells(baris_pas_ngisi + 1, kolom).Value)
                        .Cells(baris_status + 1, kolom_status).Value = total_sc
                        .Cells(baris_status + 1, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 1, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 1, kolom_status).Font.Color = &H800000
                    '---------------------------total connect------------------------------------------
                        If .Cells(Baris, kolom_total).Value = "" Then
                            .Cells(Baris, kolom_total).Value = (.Cells(baris_pas_ngisi, kolom).Value + .Cells(baris_pas_ngisi + 1, kolom).Value)
                        Else
                            .Cells(Baris, kolom_total).Value = (.Cells(baris_pas_ngisi, kolom).Value + .Cells(baris_pas_ngisi + 1, kolom).Value)
                        End If
                        .Cells(Baris, kolom_total).HorizontalAlignment = xlCenter
                        .Cells(Baris, kolom_total).Interior.Color = &HC0C0C0
                        .Cells(Baris, kolom_total).Font.Color = &H800000
                        
                        total_connect = total_connect + (.Cells(Baris, kolom_total).Value)
                                                
                        .Cells(Baris, kolom_total).Value = total_connect
                        
                        '---------------------------------------------------------------------
                        total_ap = total_ap + (.Cells(baris_pas_ngisi + 4, kolom).Value)
                        .Cells(baris_status + 4, kolom_status).Value = total_sc
                        .Cells(baris_status + 4, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 4, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 4, kolom_status).Font.Color = &HC0C000
                        
                        total_bp = total_bp + (.Cells(baris_pas_ngisi + 5, kolom).Value)
                        .Cells(baris_status + 5, kolom_status).Value = total_bp
                        .Cells(baris_status + 5, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 5, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 5, kolom_status).Font.Color = &HC0C000
                        
                        total_nego = total_nego + (.Cells(baris_pas_ngisi + 6, kolom).Value)
                        .Cells(baris_status + 6, kolom_status).Value = total_nego
                        .Cells(baris_status + 6, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 6, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 6, kolom_status).Font.Color = &HC0C000
                        
                        total_ptp = total_ptp + (.Cells(baris_pas_ngisi + 7, kolom).Value)
                        .Cells(baris_status + 7, kolom_status).Value = total_ptp
                        .Cells(baris_status + 7, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 7, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 7, kolom_status).Font.Color = &HC0C000
                        '-----------------------total contact-----------------------------------------------
                        .Cells(Baris + 4, kolom_total).Value = (.Cells(baris_pas_ngisi + 4, kolom).Value + .Cells(baris_pas_ngisi + 5, kolom).Value + .Cells(baris_pas_ngisi + 6, kolom).Value + .Cells(baris_pas_ngisi + 7, kolom).Value)
                        .Cells(Baris + 4, kolom_total).HorizontalAlignment = xlCenter
                        .Cells(Baris + 4, kolom_total).Interior.Color = &HC0C0C0
                        .Cells(Baris + 4, kolom_total).Font.Color = &HC0C000
                        .Cells(Baris + 4, kolom_total).Font.Bold = True
                        
                        total_contact = total_contact + (.Cells(Baris + 4, kolom_total).Value)
                        
                        .Cells(Baris + 4, kolom_total).Value = total_contact
                        '---------------------------------------------------------------------
                        total_b = total_b + (.Cells(baris_pas_ngisi + 10, kolom).Value)
                        .Cells(baris_status + 10, kolom_status).Value = total_b
                        .Cells(baris_status + 10, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 10, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 10, kolom_status).Font.Color = vbRed
                        
                        total_d = total_d + (.Cells(baris_pas_ngisi + 11, kolom).Value)
                        .Cells(baris_status + 11, kolom_status).Value = total_d
                        .Cells(baris_status + 11, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 11, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 11, kolom_status).Font.Color = vbRed
                        
                        total_in = total_in + (.Cells(baris_pas_ngisi + 12, kolom).Value)
                        .Cells(baris_status + 12, kolom_status).Value = total_in
                        .Cells(baris_status + 12, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 12, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 12, kolom_status).Font.Color = vbRed
                        
                        total_mb = total_mb + (.Cells(baris_pas_ngisi + 13, kolom).Value)
                        .Cells(baris_status + 13, kolom_status).Value = total_mb
                        .Cells(baris_status + 13, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 13, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 13, kolom_status).Font.Color = vbRed
                        
                        total_pa = total_pa + (.Cells(baris_pas_ngisi + 14, kolom).Value)
                        .Cells(baris_status + 14, kolom_status).Value = total_pa
                        .Cells(baris_status + 14, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 14, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 14, kolom_status).Font.Color = vbRed
                        
                        total_ssl = total_ssl + (.Cells(baris_pas_ngisi + 15, kolom).Value)
                        .Cells(baris_status + 15, kolom_status).Value = total_ssl
                        .Cells(baris_status + 15, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 15, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 15, kolom_status).Font.Color = vbRed
                        
                        total_tadt = total_tadt + (.Cells(baris_pas_ngisi + 16, kolom).Value)
                        .Cells(baris_status + 16, kolom_status).Value = total_tadt
                        .Cells(baris_status + 16, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 16, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 16, kolom_status).Font.Color = vbRed
                        
                        total_td = total_td + (.Cells(baris_pas_ngisi + 17, kolom).Value)
                        .Cells(baris_status + 17, kolom_status).Value = total_td
                        .Cells(baris_status + 17, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 17, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 17, kolom_status).Font.Color = vbRed
                        
                        total_un = total_un + (.Cells(baris_pas_ngisi + 18, kolom).Value)
                        .Cells(baris_status + 18, kolom_status).Value = total_un
                        .Cells(baris_status + 18, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 18, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 18, kolom_status).Font.Color = vbRed
                        
                        total_dr = total_dr + (.Cells(baris_pas_ngisi + 19, kolom).Value)
                        .Cells(baris_status + 19, kolom_status).Value = total_dr
                        .Cells(baris_status + 19, kolom_status).HorizontalAlignment = xlCenter
                        .Cells(baris_status + 19, kolom_status).Interior.Color = &HC0C0C0
                        .Cells(baris_status + 19, kolom_status).Font.Color = vbRed
                        '------------------------total unconnect---------------------------------------------
                        .Cells(Baris + 10, kolom_total).Value = (.Cells(baris_pas_ngisi + 10, kolom).Value + .Cells(baris_pas_ngisi + 11, kolom).Value + .Cells(baris_pas_ngisi + 12, kolom).Value + .Cells(baris_pas_ngisi + 13, kolom).Value + .Cells(baris_pas_ngisi + 14, kolom).Value + .Cells(baris_pas_ngisi + 15, kolom).Value + .Cells(baris_pas_ngisi + 16, kolom).Value + .Cells(baris_pas_ngisi + 17, kolom).Value + .Cells(baris_pas_ngisi + 18, kolom).Value + .Cells(baris_pas_ngisi + 19, kolom).Value)
                        .Cells(Baris + 10, kolom_total).HorizontalAlignment = xlCenter
                        .Cells(Baris + 10, kolom_total).Interior.Color = &HC0C0C0
                        .Cells(Baris + 10, kolom_total).Font.Color = vbRed
                        
                        total_unconnect = total_unconnect + (.Cells(Baris + 10, kolom_total).Value)
                        
                        .Cells(Baris + 10, kolom_total).Value = total_unconnect
                        
            baris_tgl = kolom
            kolom_tgl = kolom_tgl + 1
            kolom_region = kolom_region + 1
            kolom_hari = kolom_hari + 1
            kolom = kolom + 1
                
                
                
                
                
            rs_region2.MoveNext
        Wend
                Dim persen_connect_r, persen_contact_r, persen_unconnect_r, total_persen_all_r As Double
        
                
                .Cells(baris_region_r, kolom_region_r).Value = (.Cells(Baris, kolom_total).Value + .Cells(Baris + 4, kolom_total).Value + .Cells(Baris + 10, kolom_total).Value)
                total_by_region = (.Cells(baris_region_r, kolom_region_r).Value)
                .Cells(baris_region_r, kolom_region_r).HorizontalAlignment = xlCenter
                .Cells(baris_region_r, kolom_region_r).Interior.Color = &HC0C000
                .Cells(baris_region_r, kolom_region_r).Font.Color = &HFFFFFF
                .Cells(baris_region_r, kolom_region_r).Font.Size = "12"
                .Cells(baris_region_r, kolom_region_r).Font.Bold = True
                '----------------------------------------------------------------
                persen_connect_r = ((.Cells(Baris, kolom_total).Value / .Cells(baris_region_r, kolom_region_r).Value) * 100)
                
                If persen_connect_r = 0 Then
                    .Cells(baris_persen_r + 2, kolom_persen_r).Value = "0%"
                    .Cells(baris_persen_r + 2, kolom_persen_r).HorizontalAlignment = xlCenter
                    .Cells(baris_persen_r + 2, kolom_persen_r).Interior.Color = &HC0C0C0
                    .Cells(baris_persen_r + 2, kolom_persen_r).Font.Color = &H800000
                Else
                    .Cells(baris_persen_r + 2, kolom_persen_r).Value = Format(persen_connect_r, "#.##") + "%"
                    .Cells(baris_persen_r + 2, kolom_persen_r).HorizontalAlignment = xlCenter
                    .Cells(baris_persen_r + 2, kolom_persen_r).Interior.Color = &HC0C0C0
                    .Cells(baris_persen_r + 2, kolom_persen_r).Font.Color = &H800000
                End If
                
                '----------------------------------------------------------------
                persen_contact_r = ((.Cells(Baris + 4, kolom_total).Value / .Cells(baris_region_r, kolom_region_r).Value) * 100)
                
                If persen_contact_r = 0 Then
                    .Cells(baris_persen_r + 6, kolom_persen_r).Value = "0%"
                    .Cells(baris_persen_r + 6, kolom_persen_r).HorizontalAlignment = xlCenter
                    .Cells(baris_persen_r + 6, kolom_persen_r).Interior.Color = &HC0C0C0
                    .Cells(baris_persen_r + 6, kolom_persen_r).Font.Color = &HC0C000
                    .Cells(baris_persen_r + 6, kolom_persen_r).Font.Bold = True
                Else
                    .Cells(baris_persen_r + 6, kolom_persen_r).Value = Format(persen_contact_r, "#.##") + "%"
                    .Cells(baris_persen_r + 6, kolom_persen_r).HorizontalAlignment = xlCenter
                    .Cells(baris_persen_r + 6, kolom_persen_r).Interior.Color = &HC0C0C0
                    .Cells(baris_persen_r + 6, kolom_persen_r).Font.Color = &HC0C000
                    .Cells(baris_persen_r + 6, kolom_persen_r).Font.Bold = True
                End If
                
                '----------------------------------------------------------------
                persen_unconnect_r = ((.Cells(Baris + 10, kolom_total).Value / .Cells(baris_region_r, kolom_region_r).Value) * 100)
                
                If persen_unconnect_r = 0 Then
                    .Cells(baris_persen_r + 12, kolom_persen_r).Value = "0%"
                    .Cells(baris_persen_r + 12, kolom_persen_r).HorizontalAlignment = xlCenter
                    .Cells(baris_persen_r + 12, kolom_persen_r).Interior.Color = &HC0C0C0
                    .Cells(baris_persen_r + 12, kolom_persen_r).Font.Color = vbRed
                Else
                    .Cells(baris_persen_r + 12, kolom_persen_r).Value = Format(persen_unconnect_r, "#.##") + "%"
                    .Cells(baris_persen_r + 12, kolom_persen_r).HorizontalAlignment = xlCenter
                    .Cells(baris_persen_r + 12, kolom_persen_r).Interior.Color = &HC0C0C0
                    .Cells(baris_persen_r + 12, kolom_persen_r).Font.Color = vbRed
                End If
                '----------------------------------------------------------------
                nasional = (.Cells(5, 4).Value)
                total_persen_all_r = ((total_by_region / nasional) * 100)
                
                If total_persen_all_r = 0 Then
                    .Cells(baris_region_r, kolom_region_r + 1).Value = "0%"
                    .Cells(baris_region_r, kolom_region_r + 1).HorizontalAlignment = xlCenter
                    .Cells(baris_region_r, kolom_region_r + 1).Interior.Color = &HC0C000
                    .Cells(baris_region_r, kolom_region_r + 1).Font.Color = &HFFFFFF
                    .Cells(baris_region_r, kolom_region_r + 1).Font.Size = "12"
                    .Cells(baris_region_r, kolom_region_r + 1).Font.Bold = True
                Else
                    .Cells(baris_region_r, kolom_region_r + 1).Value = CStr(total_persen_all_r) + "%"
                    .Cells(baris_region_r, kolom_region_r + 1).HorizontalAlignment = xlCenter
                    .Cells(baris_region_r, kolom_region_r + 1).Interior.Color = &HC0C000
                    .Cells(baris_region_r, kolom_region_r + 1).Font.Color = &HFFFFFF
                    .Cells(baris_region_r, kolom_region_r + 1).Font.Size = "12"
                    .Cells(baris_region_r, kolom_region_r + 1).Font.Bold = True
                End If
            rs_region.MoveNext
            
            ELIN2.Range(arrayAlphabet(2) & baris_baru + 2 & " :" & arrayAlphabet(kolom - 1) & baris_pas_ngisi + 20).Borders(xlInsideHorizontal).LineStyle = xlDash
            
            baris_baru = baris_baru + 24
            Baris = Baris + 24
            baris_status = baris_status + 24
            kolom_region = 4
            kolom_total = 4
            kolom_hari = 4
            kolom = 6
            kolom_tgl = 6
            baris_tgl = 2
            baris_region = 0
            kolom_region_r = 4
            kolom_persen_r = 5
            baris_region_r = baris_region_r + 24
            baris_pas_ngisi = baris_pas_ngisi + 24
            baris_persen_r = baris_persen_r + 24
            
        
        Wend
    End With
End Sub
Private Sub SSCommand2_Click()
    Unload Me
End Sub

