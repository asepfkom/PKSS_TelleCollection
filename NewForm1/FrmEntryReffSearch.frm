VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form FrmEntryReffSearch 
   Caption         =   "Entry Refferall Data"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13170
   Icon            =   "FrmEntryReffSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   13170
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      ForeColor       =   &H00000000&
      Height          =   7605
      Left            =   45
      TabIndex        =   22
      Top             =   1515
      Width           =   13080
      Begin VB.TextBox TxtJmlDtMgm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   11925
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   7980
         Width           =   3045
      End
      Begin MSComctlLib.ListView LstVwSearchMgm 
         Height          =   7440
         Left            =   45
         TabIndex        =   8
         Top             =   120
         Width           =   12990
         _ExtentX        =   22913
         _ExtentY        =   13123
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   1425
      Left            =   60
      TabIndex        =   15
      Top             =   60
      Width           =   13065
      Begin VB.TextBox TxtNama 
         Height          =   375
         Left            =   1800
         TabIndex        =   0
         Top             =   315
         Width           =   2895
      End
      Begin VB.TextBox TxtTelpRumah 
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   690
         Width           =   2505
      End
      Begin VB.TextBox TxtTelpKantor 
         Height          =   375
         Left            =   6360
         TabIndex        =   2
         Top             =   285
         Width           =   2505
      End
      Begin VB.TextBox TxtHandPhone 
         Height          =   375
         Left            =   6360
         TabIndex        =   3
         Top             =   660
         Width           =   2505
      End
      Begin VB.ComboBox cmbRecsource 
         Height          =   315
         Left            =   9585
         TabIndex        =   5
         Top             =   1425
         Visible         =   0   'False
         Width           =   1695
      End
      Begin Threed.SSCommand CmdSave 
         Height          =   390
         Left            =   9600
         TabIndex        =   6
         Top             =   705
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   688
         _Version        =   196610
         MousePointer    =   16
         Caption         =   "&Search"
      End
      Begin TDBDate6Ctl.TDBDate TdbDOB 
         Height          =   360
         Left            =   9585
         TabIndex        =   4
         Top             =   240
         Width           =   1470
         _Version        =   65536
         _ExtentX        =   2593
         _ExtentY        =   635
         Calendar        =   "FrmEntryReffSearch.frx":000C
         Caption         =   "FrmEntryReffSearch.frx":0124
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmEntryReffSearch.frx":0190
         Keys            =   "FrmEntryReffSearch.frx":01AE
         Spin            =   "FrmEntryReffSearch.frx":020C
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
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
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   1.15962147735399E-317
         CenturyMode     =   0
      End
      Begin Threed.SSCommand CmdCancel 
         Height          =   390
         Left            =   10350
         TabIndex        =   7
         Top             =   705
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   688
         _Version        =   196610
         MousePointer    =   16
         Caption         =   "&Cancel"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama :"
         Height          =   345
         Left            =   675
         TabIndex        =   21
         Top             =   345
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Telp Rumah : "
         Height          =   345
         Left            =   720
         TabIndex        =   20
         Top             =   705
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Telp Kantor : "
         Height          =   345
         Left            =   5280
         TabIndex        =   19
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "HandPhone : "
         Height          =   345
         Left            =   5280
         TabIndex        =   18
         Top             =   675
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DOB : "
         Height          =   345
         Left            =   8490
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Batch : "
         Height          =   345
         Left            =   8505
         TabIndex        =   16
         Top             =   1425
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.TextBox TxtNamaReff 
      Enabled         =   0   'False
      Height          =   375
      Left            =   8595
      TabIndex        =   13
      Top             =   3540
      Width           =   2895
   End
   Begin VB.TextBox TxtRecsourceRef 
      Enabled         =   0   'False
      Height          =   375
      Left            =   8460
      TabIndex        =   11
      Top             =   3150
      Width           =   2895
   End
   Begin VB.TextBox TxtIdReff 
      Height          =   375
      Left            =   8535
      TabIndex        =   9
      Top             =   2700
      Width           =   2295
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "NamaRef :"
      Enabled         =   0   'False
      Height          =   345
      Left            =   7470
      TabIndex        =   14
      Top             =   3570
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Recsource Ref :"
      Enabled         =   0   'False
      Height          =   345
      Left            =   7335
      TabIndex        =   12
      Top             =   3180
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Id Ref :"
      Height          =   345
      Left            =   7410
      TabIndex        =   10
      Top             =   2730
      Width           =   1095
   End
End
Attribute VB_Name = "FrmEntryReffSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public okReff As Boolean
Dim NoTelpSama As String

Private Sub cmbRecsource_LostFocus()
Dim m_obj As New ADODB.Recordset
m_obj.CursorLocation = adUseClient
m_obj.Open "Select * from DATASOURCETBL WHERE KODEDS = '" + cmbRecsource.Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If m_obj.RecordCount <> 0 Then
    cmbRecsource.Text = m_obj!KODEDS
Else
    cmbRecsource.Text = ""
End If
Set m_obj = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    If Len(TxtNama.Text) < 2 Then
        MsgBox "Nama harus diisi", vbInformation + vbOKOnly, "Telegrandi"
        Exit Sub
    End If
    If Len(TxtTelpRumah.Text) < 2 And Len(TxtTelpKantor.Text) < 2 And Len(TxtHandPhone.Text) < 2 Then
        MsgBox "Minimal salah satu dari telp harus diisi", vbInformation + vbOKOnly, "Telegrandi"
        Exit Sub
    End If
'    If Len(cmbRecsource.Text) < 2 Then
'        MsgBox "Batch harus diisi", vbInformation + vbOKOnly, "Telegrandi"
'        Exit Sub
'    End If
    'CmdSave.Enabled = False
    Call cari_duplicate
End Sub

Private Sub cari_duplicate()
    Dim CMDSQL As String
    Dim mrs_cek As ADODB.Recordset
    Dim kriteria1 As String
    Dim kriteria2 As String
    Dim CUSTID1 As String
Dim Bookmark As String, CUSTID As String, RECSTATUS As String, PRIOR As String, CUSTIDREF As String, NAMAREF As String, NAME As String, NEXTACTDATE As String, NEXTACT As String, REMARKS As String, agent As String, _
                            NAMAAGENT As String, RECSOURCEREF As String, TGLSTATUS As String, KETHSLKERJA As String, KdComplaint As String, RemarkComplaint As String
    ' kriteria pertama
    'nama ama notelp
    If Len(TxtNama.Text) > 2 And Len(TxtTelpRumah.Text) > 2 Then
        kriteria2 = Left(TxtTelpRumah.Text, 5)
        CMDSQL = "Select * from cc_custtbl where name like '%" + TxtNama.Text + "%' "
        CMDSQL = CMDSQL + " and (HOMENO like '%" + kriteria2 + "%' or HOMENO2 like '%" + kriteria2 + "%') "
        'CMDSQL = CMDSQL + " and (HOMENO like '%" + kriteria2 + "%' or HOMENO2 like '%" + kriteria2 + "%' or mobileno like '%" + kriteria2 + "%' or mobileno2 like '%" + kriteria2 + "%' or officeno like '%" + kriteria2 + "%' or officeno2 like '%" + kriteria2 + "%') "
    
    Set mrs_cek = New ADODB.Recordset
    mrs_cek.CursorLocation = adUseClient
        
        mrs_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If mrs_cek.RecordCount <> 0 Then
            ' paksain keluar deh...
            
            MsgBox " Nama dan Telp Rumah Ada yg sama", vbInformation + vbOKOnly, "Telegrandi"
            
            CUSTID1 = Empty
            While Not mrs_cek.EOF
                'Show Ke ListView
                CUSTID = IIf(IsNull(mrs_cek("custid")), "", mrs_cek("custid"))
                NoTelpSama = IIf(IsNull(mrs_cek("HOMENO")), "", mrs_cek("HOMENO")) + " - " + IIf(IsNull(mrs_cek("HOMENO2")), "", mrs_cek("HOMENO2"))
                PRIOR = IIf(IsNull(mrs_cek("PRIOR")), "", mrs_cek("PRIOR"))
                CUSTIDREF = IIf(IsNull(mrs_cek("CUSTIDREF")), "", mrs_cek("CUSTIDREF"))
                NAMAREF = IIf(IsNull(mrs_cek("NAMAREF")), "", mrs_cek("NAMAREF"))
                NAME = IIf(IsNull(mrs_cek("NAME")), "", mrs_cek("NAME"))
                NEXTACTDATE = IIf(IsNull(mrs_cek("NEXTACTDATE")), "", Format(mrs_cek("NEXTACTDATE"), "yyyy/mm/dd hh:mm"))
                NEXTACT = IIf(IsNull(mrs_cek("NEXTACT")), "", mrs_cek("NEXTACT"))
                REMARKS = IIf(IsNull(mrs_cek("REMARKS")), "", mrs_cek("REMARKS"))
                agent = IIf(IsNull(mrs_cek("AGENT")), "", mrs_cek("AGENT"))
                NAMAAGENT = IIf(IsNull(mrs_cek("NamaAGENT")), "", mrs_cek("NamaAGENT"))
                RECSOURCEREF = IIf(IsNull(mrs_cek("RECSOURCEREF")), "", mrs_cek("RECSOURCEREF"))
                TGLSTATUS = Format(IIf(IsNull(mrs_cek("TGLSTATUS")), "", mrs_cek("TGLSTATUS")), "yyyy/mm/dd")
                KETHSLKERJA = IIf(IsNull(mrs_cek("Kethslkerja")), "", mrs_cek("Kethslkerja"))
                KdComplaint = IIf(IsNull(mrs_cek("KdComplaint")), "", mrs_cek("KdComplaint"))
                RemarkComplaint = IIf(IsNull(mrs_cek("RemarkComplaint")), "", mrs_cek("RemarkComplaint"))
                Bookmark = mrs_cek.Bookmark
                Call show_refferall(Bookmark, CUSTID, RECSTATUS, PRIOR, CUSTIDREF, NAMAREF, NAME, NEXTACTDATE, NEXTACT, REMARKS, agent, _
                                        NAMAAGENT, RECSOURCEREF, TGLSTATUS, KETHSLKERJA, KdComplaint, RemarkComplaint, kriteria2)
            
                CUSTID1 = "REFI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC"))
                
                CMDSQL = "Insert into TBL_DUPLIKASI (CUSTID, CUSTIDBARU, NAMABARU, HOMENOBARU, OFFICENOBARU, MOBILENOBARU, AGENTBARU, CHBARUNAME, CHBARUID, "
                If TdbDOB.ValueIsNull = False Then
                    CMDSQL = CMDSQL + "DOB,"
                End If
                CMDSQL = CMDSQL + " RECSOURCEBARU) values "
                CMDSQL = CMDSQL + "('" + mrs_cek!CUSTID + "',"
                CMDSQL = CMDSQL + "'" + CUSTID1 + "',"
                CMDSQL = CMDSQL + "'" + TxtNama.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtTelpRumah.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtTelpKantor.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtHandPhone.Text + "',"
                CMDSQL = CMDSQL + "'" + MDIForm1.Text1.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtNamaReff.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtIdReff.Text + "',"
                If TdbDOB.ValueIsNull = False Then
                    CMDSQL = CMDSQL + "'" + Format(TdbDOB.Value, "yyyy/mm/dd") + "',"
                End If
                CMDSQL = CMDSQL + "'" + cmbRecsource.Text + "')"
                M_OBJCONN.Execute CMDSQL
                mrs_cek.MoveNext
            Wend
            Set mrs_cek = Nothing
            
          '  Unload Me
            Exit Sub
            
        End If
        Set mrs_cek = Nothing
    End If
    If Len(TxtNama.Text) > 2 And Len(TxtTelpKantor.Text) > 2 Then
        kriteria2 = Left(TxtTelpKantor.Text, 5)
        CMDSQL = "Select * from cc_custtbl where name like '%" + TxtNama.Text + "%' "
        CMDSQL = CMDSQL + " and (officeno like '%" + kriteria2 + "%' or officeno2 like '%" + kriteria2 + "%') "
        'CMDSQL = CMDSQL + " and (HOMENO like '%" + kriteria2 + "%' or HOMENO2 like '%" + kriteria2 + "%' or mobileno like '%" + kriteria2 + "%' or mobileno2 like '%" + kriteria2 + "%' or officeno like '%" + kriteria2 + "%' or officeno2 like '%" + kriteria2 + "%') "
    Set mrs_cek = New ADODB.Recordset
    mrs_cek.CursorLocation = adUseClient
        
        mrs_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If mrs_cek.RecordCount <> 0 Then
            MsgBox " Nama dan Telp Kantor Ada yg sama", vbInformation + vbOKOnly, "Telegrandi"
            
            CUSTID1 = Empty
            While Not mrs_cek.EOF
                'Show Ke ListView
                CUSTID = IIf(IsNull(mrs_cek("custid")), "", mrs_cek("custid"))
                NoTelpSama = IIf(IsNull(mrs_cek("OFFICENO")), "", mrs_cek("OFFICENO")) + " - " + IIf(IsNull(mrs_cek("OFFICENO2")), "", mrs_cek("OFFICENO2"))
                PRIOR = IIf(IsNull(mrs_cek("PRIOR")), "", mrs_cek("PRIOR"))
                CUSTIDREF = IIf(IsNull(mrs_cek("CUSTIDREF")), "", mrs_cek("CUSTIDREF"))
                NAMAREF = IIf(IsNull(mrs_cek("NAMAREF")), "", mrs_cek("NAMAREF"))
                NAME = IIf(IsNull(mrs_cek("NAME")), "", mrs_cek("NAME"))
                NEXTACTDATE = IIf(IsNull(mrs_cek("NEXTACTDATE")), "", Format(mrs_cek("NEXTACTDATE"), "yyyy/mm/dd hh:mm"))
                NEXTACT = IIf(IsNull(mrs_cek("NEXTACT")), "", mrs_cek("NEXTACT"))
                REMARKS = IIf(IsNull(mrs_cek("REMARKS")), "", mrs_cek("REMARKS"))
                agent = IIf(IsNull(mrs_cek("AGENT")), "", mrs_cek("AGENT"))
                NAMAAGENT = IIf(IsNull(mrs_cek("NamaAGENT")), "", mrs_cek("NamaAGENT"))
                RECSOURCEREF = IIf(IsNull(mrs_cek("RECSOURCEREF")), "", mrs_cek("RECSOURCEREF"))
                TGLSTATUS = Format(IIf(IsNull(mrs_cek("TGLSTATUS")), "", mrs_cek("TGLSTATUS")), "yyyy/mm/dd")
                KETHSLKERJA = IIf(IsNull(mrs_cek("Kethslkerja")), "", mrs_cek("Kethslkerja"))
                KdComplaint = IIf(IsNull(mrs_cek("KdComplaint")), "", mrs_cek("KdComplaint"))
                RemarkComplaint = IIf(IsNull(mrs_cek("RemarkComplaint")), "", mrs_cek("RemarkComplaint"))
                Bookmark = mrs_cek.Bookmark
                Call show_refferall(Bookmark, CUSTID, RECSTATUS, PRIOR, CUSTIDREF, NAMAREF, NAME, NEXTACTDATE, NEXTACT, REMARKS, agent, _
                                        NAMAAGENT, RECSOURCEREF, TGLSTATUS, KETHSLKERJA, KdComplaint, RemarkComplaint, kriteria2)
            
                CUSTID1 = "REFI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC"))
                
                CMDSQL = "Insert into TBL_DUPLIKASI (CUSTID, CUSTIDBARU, NAMABARU, HOMENOBARU, OFFICENOBARU, MOBILENOBARU, AGENTBARU, CHBARUNAME, CHBARUID, "
                If TdbDOB.ValueIsNull = False Then
                    CMDSQL = CMDSQL + "DOB,"
                End If
                CMDSQL = CMDSQL + " RECSOURCEBARU) values "
                CMDSQL = CMDSQL + "('" + mrs_cek!CUSTID + "',"
                CMDSQL = CMDSQL + "'" + CUSTID1 + "',"
                CMDSQL = CMDSQL + "'" + TxtNama.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtTelpRumah.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtTelpKantor.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtHandPhone.Text + "',"
                CMDSQL = CMDSQL + "'" + MDIForm1.Text1.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtNamaReff.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtIdReff.Text + "',"
                If TdbDOB.ValueIsNull = False Then
                    CMDSQL = CMDSQL + "'" + Format(TdbDOB.Value, "yyyy/mm/dd") + "',"
                End If
                CMDSQL = CMDSQL + "'" + cmbRecsource.Text + "')"
                M_OBJCONN.Execute CMDSQL
                mrs_cek.MoveNext
            Wend
            Set mrs_cek = Nothing
            
 '           Unload Me
            Exit Sub
        End If
        Set mrs_cek = Nothing
        
    End If
    If Len(TxtNama.Text) > 2 And Len(TxtHandPhone.Text) > 2 Then
        kriteria2 = Left(TxtHandPhone.Text, 8)
        CMDSQL = "Select * from cc_custtbl where name like '%" + TxtNama.Text + "%' "
        CMDSQL = CMDSQL + " and (mobileno like '%" + kriteria2 + "%' or mobileno2 like '%" + kriteria2 + "%') "
    Set mrs_cek = New ADODB.Recordset
    mrs_cek.CursorLocation = adUseClient
        
        mrs_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If mrs_cek.RecordCount <> 0 Then
            MsgBox "Nama dan Handphone Ada yg sama", vbInformation + vbOKOnly, "Telegrandi"
            
            CUSTID1 = Empty
            While Not mrs_cek.EOF
                'Show Ke ListView
                CUSTID = IIf(IsNull(mrs_cek("custid")), "", mrs_cek("custid"))
                NoTelpSama = IIf(IsNull(mrs_cek("MOBILENO")), "", mrs_cek("MOBILENO")) + " - " + IIf(IsNull(mrs_cek("MOBILENO2")), "", mrs_cek("MOBILENO2"))
                PRIOR = IIf(IsNull(mrs_cek("PRIOR")), "", mrs_cek("PRIOR"))
                CUSTIDREF = IIf(IsNull(mrs_cek("CUSTIDREF")), "", mrs_cek("CUSTIDREF"))
                NAMAREF = IIf(IsNull(mrs_cek("NAMAREF")), "", mrs_cek("NAMAREF"))
                NAME = IIf(IsNull(mrs_cek("NAME")), "", mrs_cek("NAME"))
                NEXTACTDATE = IIf(IsNull(mrs_cek("NEXTACTDATE")), "", Format(mrs_cek("NEXTACTDATE"), "yyyy/mm/dd hh:mm"))
                NEXTACT = IIf(IsNull(mrs_cek("NEXTACT")), "", mrs_cek("NEXTACT"))
                REMARKS = IIf(IsNull(mrs_cek("REMARKS")), "", mrs_cek("REMARKS"))
                agent = IIf(IsNull(mrs_cek("AGENT")), "", mrs_cek("AGENT"))
                NAMAAGENT = IIf(IsNull(mrs_cek("NamaAGENT")), "", mrs_cek("NamaAGENT"))
                RECSOURCEREF = IIf(IsNull(mrs_cek("RECSOURCEREF")), "", mrs_cek("RECSOURCEREF"))
                TGLSTATUS = Format(IIf(IsNull(mrs_cek("TGLSTATUS")), "", mrs_cek("TGLSTATUS")), "yyyy/mm/dd")
                KETHSLKERJA = IIf(IsNull(mrs_cek("Kethslkerja")), "", mrs_cek("Kethslkerja"))
                KdComplaint = IIf(IsNull(mrs_cek("KdComplaint")), "", mrs_cek("KdComplaint"))
                RemarkComplaint = IIf(IsNull(mrs_cek("RemarkComplaint")), "", mrs_cek("RemarkComplaint"))
                Bookmark = mrs_cek.Bookmark
                Call show_refferall(Bookmark, CUSTID, RECSTATUS, PRIOR, CUSTIDREF, NAMAREF, NAME, NEXTACTDATE, NEXTACT, REMARKS, agent, _
                                        NAMAAGENT, RECSOURCEREF, TGLSTATUS, KETHSLKERJA, KdComplaint, RemarkComplaint, kriteria2)
            
                CUSTID1 = "REFI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC"))
                
                CMDSQL = "Insert into TBL_DUPLIKASI (CUSTID, CUSTIDBARU, NAMABARU, HOMENOBARU, OFFICENOBARU, MOBILENOBARU, AGENTBARU, CHBARUNAME, CHBARUID, "
                If TdbDOB.ValueIsNull = False Then
                    CMDSQL = CMDSQL + "DOB,"
                End If
                CMDSQL = CMDSQL + " RECSOURCEBARU) values "
                CMDSQL = CMDSQL + "('" + mrs_cek!CUSTID + "',"
                CMDSQL = CMDSQL + "'" + CUSTID1 + "',"
                CMDSQL = CMDSQL + "'" + TxtNama.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtTelpRumah.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtTelpKantor.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtHandPhone.Text + "',"
                CMDSQL = CMDSQL + "'" + MDIForm1.Text1.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtNamaReff.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtIdReff.Text + "',"
                If TdbDOB.ValueIsNull = False Then
                    CMDSQL = CMDSQL + "'" + Format(TdbDOB.Value, "yyyy/mm/dd") + "',"
                End If
                CMDSQL = CMDSQL + "'" + cmbRecsource.Text + "')"
                
                M_OBJCONN.Execute CMDSQL
                mrs_cek.MoveNext
            Wend
            Set mrs_cek = Nothing
            
 '           Unload Me
            Exit Sub
        End If
        Set mrs_cek = Nothing
    
    End If
    If Len(TxtNama.Text) > 2 And TdbDOB.ValueIsNull = False Then
        kriteria2 = Format(TdbDOB.Value, "yyyy/mm/dd")
        CMDSQL = "Select * from cc_custtbl where name like '%" + TxtNama.Text + "%' "
        CMDSQL = CMDSQL + " and birthd = '" + Format(TdbDOB.Value, "yyyy/mm/dd") + "'"
        Set mrs_cek = New ADODB.Recordset
            mrs_cek.CursorLocation = adUseClient

        mrs_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If mrs_cek.RecordCount <> 0 Then
            MsgBox "Nama dan DOB Ada yg sama", vbInformation + vbOKOnly, "Telegrandi"
            
            CUSTID1 = Empty
            While Not mrs_cek.EOF
                'Show Ke ListView
                CUSTID = IIf(IsNull(mrs_cek("custid")), "", mrs_cek("custid"))
                PRIOR = IIf(IsNull(mrs_cek("PRIOR")), "", mrs_cek("PRIOR"))
                CUSTIDREF = IIf(IsNull(mrs_cek("CUSTIDREF")), "", mrs_cek("CUSTIDREF"))
                NAMAREF = IIf(IsNull(mrs_cek("NAMAREF")), "", mrs_cek("NAMAREF"))
                NAME = IIf(IsNull(mrs_cek("NAME")), "", mrs_cek("NAME"))
                NEXTACTDATE = IIf(IsNull(mrs_cek("NEXTACTDATE")), "", Format(mrs_cek("NEXTACTDATE"), "yyyy/mm/dd hh:mm"))
                NEXTACT = IIf(IsNull(mrs_cek("NEXTACT")), "", mrs_cek("NEXTACT"))
                REMARKS = IIf(IsNull(mrs_cek("REMARKS")), "", mrs_cek("REMARKS"))
                agent = IIf(IsNull(mrs_cek("AGENT")), "", mrs_cek("AGENT"))
                NAMAAGENT = IIf(IsNull(mrs_cek("NamaAGENT")), "", mrs_cek("NamaAGENT"))
                RECSOURCEREF = IIf(IsNull(mrs_cek("RECSOURCEREF")), "", mrs_cek("RECSOURCEREF"))
                TGLSTATUS = Format(IIf(IsNull(mrs_cek("TGLSTATUS")), "", mrs_cek("TGLSTATUS")), "yyyy/mm/dd")
                KETHSLKERJA = IIf(IsNull(mrs_cek("Kethslkerja")), "", mrs_cek("Kethslkerja"))
                KdComplaint = IIf(IsNull(mrs_cek("KdComplaint")), "", mrs_cek("KdComplaint"))
                RemarkComplaint = IIf(IsNull(mrs_cek("RemarkComplaint")), "", mrs_cek("RemarkComplaint"))
                Bookmark = mrs_cek.Bookmark
                Call show_refferall(Bookmark, CUSTID, RECSTATUS, PRIOR, CUSTIDREF, NAMAREF, NAME, NEXTACTDATE, NEXTACT, REMARKS, agent, _
                                        NAMAAGENT, RECSOURCEREF, TGLSTATUS, KETHSLKERJA, KdComplaint, RemarkComplaint, kriteria2)
            
                CUSTID1 = "REFI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC"))
                
                CMDSQL = "Insert into TBL_DUPLIKASI (CUSTID, CUSTIDBARU, NAMABARU, HOMENOBARU, OFFICENOBARU, MOBILENOBARU, AGENTBARU, CHBARUNAME, CHBARUID, "
                If TdbDOB.ValueIsNull = False Then
                    CMDSQL = CMDSQL + "DOB,"
                End If
                CMDSQL = CMDSQL + " RECSOURCEBARU) values "
                CMDSQL = CMDSQL + "('" + mrs_cek!CUSTID + "',"
                CMDSQL = CMDSQL + "'" + CUSTID1 + "',"
                CMDSQL = CMDSQL + "'" + TxtNama.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtTelpRumah.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtTelpKantor.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtHandPhone.Text + "',"
                CMDSQL = CMDSQL + "'" + MDIForm1.Text1.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtNamaReff.Text + "',"
                CMDSQL = CMDSQL + "'" + TxtIdReff.Text + "',"
                If TdbDOB.ValueIsNull = False Then
                    CMDSQL = CMDSQL + "'" + Format(TdbDOB.Value, "yyyy/mm/dd") + "',"
                End If
                CMDSQL = CMDSQL + "'" + cmbRecsource.Text + "')"
                
                M_OBJCONN.Execute CMDSQL
                mrs_cek.MoveNext
            Wend
            Set mrs_cek = Nothing
            
'            Unload Me
            Exit Sub
        End If
        Set mrs_cek = Nothing
    End If
        With FrmEntryReff
            .TxtRecsourceRef.Text = TxtRecsourceRef.Text
            .TxtIdReff.Text = TxtIdReff.Text
            .TxtNamaReff.Text = TxtNamaReff.Text
            .TxtIdReff.Enabled = False
             Me.Hide
             .Show vbModal
             If .okReff Then
             Else
             End If
        End With
        
End Sub

Private Sub Form_Load()
Dim m_objrs As New ADODB.Recordset
    TdbDOB.Value = Empty
    TxtNama.Text = Empty
    TxtTelpRumah.Text = Empty
    TxtTelpKantor.Text = Empty
    TxtHandPhone.Text = Empty
    Call HEADER_VIEW_Refferall
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open "Select * from DATASOURCETBL WHERE STATUS = 'I' ", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    While Not m_objrs.EOF
        cmbRecsource.AddItem m_objrs!KODEDS
        m_objrs.MoveNext
    Wend
    Set m_objrs = Nothing
End Sub

Private Sub show_refferall(Bookmark As String, CUSTID As String, RECSTATUS As String, PRIOR As String, CUSTIDREF As String, NAMAREF As String, NAME As String, NEXTACTDATE As String, NEXTACT As String, REMARKS As String, agent As String, _
                            NAMAAGENT As String, RECSOURCEREF As String, TGLSTATUS As String, KETHSLKERJA As String, KdComplaint As String, RemarkComplaint As String, NoTelpSama As String)
Dim listitem As listitem

Set listitem = LstVwSearchMgm.ListItems.ADD(, , Bookmark)
    listitem.SubItems(1) = CUSTID
    listitem.SubItems(2) = PRIOR
    listitem.SubItems(3) = CUSTIDREF
    listitem.SubItems(4) = NAMAREF
    listitem.SubItems(5) = NAME
    listitem.SubItems(6) = NoTelpSama
    listitem.SubItems(7) = NEXTACTDATE
    listitem.SubItems(8) = NEXTACT
    listitem.SubItems(9) = REMARKS
    listitem.SubItems(10) = agent
    listitem.SubItems(11) = NAMAAGENT
    listitem.SubItems(12) = RECSOURCEREF
    listitem.SubItems(13) = TGLSTATUS
    listitem.SubItems(14) = KETHSLKERJA
    listitem.SubItems(15) = KdComplaint
    listitem.SubItems(16) = RemarkComplaint
    
End Sub

Private Sub HEADER_VIEW_Refferall()
    LstVwSearchMgm.ColumnHeaders.ADD 1, , "No", 3 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 2, , "Cust Id", 5 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 3, , "Priority", 5 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 4, , "Ref Id", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 5, , "Ref Name", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 6, , "Nama Customer", 25 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 7, , "Nomor Telp", 25 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 8, , "Tgl Schedule", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 9, , "Next Action", 12 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 10, , "Remarks", 17 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 11, , "SalesCode", 8 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 12, , "Agent", 8 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 13, , "DataBase", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 14, , "LastCall Date", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 15, , "Sts LastCall", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 16, , "Code", 5 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 17, , "Complaint Note", 15 * TXT
End Sub
