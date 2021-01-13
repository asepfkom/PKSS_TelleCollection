VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTodayList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Today List"
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15300
   ControlBox      =   0   'False
   Icon            =   "FrmTodayList.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   15300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   330
      Left            =   14295
      TabIndex        =   8
      Top             =   15
      Width           =   720
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8775
      Left            =   0
      TabIndex        =   0
      Top             =   135
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   15478
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Card Holder Data"
      TabPicture(0)   =   "FrmTodayList.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LblTarget(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Referall Data"
      TabPicture(1)   =   "FrmTodayList.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         ForeColor       =   &H00000000&
         Height          =   8370
         Left            =   -74955
         TabIndex        =   4
         Top             =   345
         Width           =   15090
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
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   7980
            Width           =   3045
         End
         Begin MSComctlLib.ListView LstVwSearchMgm 
            Height          =   7845
            Left            =   0
            TabIndex        =   6
            Top             =   120
            Width           =   15000
            _ExtentX        =   26458
            _ExtentY        =   13838
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
      Begin VB.Frame Frame4 
         BackColor       =   &H80000004&
         Height          =   8385
         Left            =   60
         TabIndex        =   1
         Top             =   345
         Width           =   15075
         Begin VB.TextBox Text2 
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
            Left            =   11910
            Locked          =   -1  'True
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   7980
            Width           =   3045
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   7845
            Left            =   -30
            TabIndex        =   3
            Top             =   120
            Width           =   15015
            _ExtentX        =   26485
            _ExtentY        =   13838
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   12582912
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
      Begin VB.Label LblTarget 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Index           =   0
         Left            =   -71805
         TabIndex        =   7
         Top             =   -15
         Width           =   9480
      End
   End
End
Attribute VB_Name = "FrmTodayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FOLLOWUPABIS As Boolean

Private Sub HEADER_VIEW_Refferall()
    ListView1.ColumnHeaders.ADD 1, , "No", 3 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Cust Id", 5 * TXT
    ListView1.ColumnHeaders.ADD 3, , "Priority", 5 * TXT
    ListView1.ColumnHeaders.ADD 4, , "Ref Id", 10 * TXT
    ListView1.ColumnHeaders.ADD 5, , "Ref Name", 10 * TXT
    ListView1.ColumnHeaders.ADD 6, , "Nama Customer", 25 * TXT
    ListView1.ColumnHeaders.ADD 7, , "Tgl Schedule", 10 * TXT
    ListView1.ColumnHeaders.ADD 8, , "Next Action", 12 * TXT
    ListView1.ColumnHeaders.ADD 9, , "Remarks", 17 * TXT
    ListView1.ColumnHeaders.ADD 10, , "SalesCode", 8 * TXT
    ListView1.ColumnHeaders.ADD 11, , "Agent", 8 * TXT
    ListView1.ColumnHeaders.ADD 12, , "DataBase", 10 * TXT
    ListView1.ColumnHeaders.ADD 13, , "LastCall Date", 10 * TXT
    ListView1.ColumnHeaders.ADD 14, , "Sts LastCall", 10 * TXT
    ListView1.ColumnHeaders.ADD 15, , "Code", 5 * TXT
    ListView1.ColumnHeaders.ADD 16, , "Complaint Note", 15 * TXT
    ListView1.ColumnHeaders.ADD 17, , "Check", 10 * TXT
    ListView1.ColumnHeaders.ADD 18, , "ID", 10 * TXT
End Sub

Private Sub HEADER_VIEW_MGM()
    LstVwSearchMgm.ColumnHeaders.ADD 1, , "No", 3 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 2, , "Cust Id", 5 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 3, , "Priority", 5 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 4, , "Nama Customer", 25 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 5, , "Tgl Schedule", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 6, , "Next Action", 12 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 7, , "Remarks", 17 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 8, , "SalesCode", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 9, , "Agent", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 10, , "DataBase", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 11, , "LastCall Date", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 12, , "Sts LastCall", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 13, , "Code", 5 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 14, , "Complaint Note", 15 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 15, , "Check", 10 * TXT
    'LstVwSearchMgm.ColumnHeaders.ADD 16, , "ID", 10 * TXT
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call HEADER_VIEW_Refferall
    Call HEADER_VIEW_MGM
    Call show_Search_Refferal
    Call show_Search_mgmData
    
'    If FOLLOWUPABIS = False Then
'        CmdClose.Enabled = True
'    Else
'        CmdClose.Enabled = False
'    End If
End Sub

Private Sub show_Search_Refferal()
Dim listitem As listitem
Dim m_cari As New ADODB.Recordset
Dim time1 As Date
Dim time2 As Date
time1 = Format(Now, "hh:nn")
If Right(Time, 2) < 46 Then
    time2 = Time
End If
time2 = Format(Now, "hh:nn")

m_cari.CursorLocation = adUseClient
m_cari.Open "Select * from cc_custtbl where agent = '" + MDIForm1.Text1.Text + "' and (NEXTACTDATE BETWEEN '" + Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") & " 00:00" + "' AND '" + Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") & " 23:59" + "') ", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

'ProgressBar1.Max = m_cari.RecordCount + 1
Text2.Text = m_cari.RecordCount & " Data"
ListView1.ListItems.Clear

While Not m_cari.EOF
'ProgressBar1.Value = m_cari.Bookmark
Set listitem = ListView1.ListItems.ADD(, , m_cari.Bookmark)
    listitem.SubItems(1) = IIf(IsNull(m_cari("custid")), "", m_cari("custid"))
    Select Case m_cari("RECSTATUS")
    Case "1A"
        listitem.SubItems(2) = "Available"
    Case ""
        listitem.SubItems(2) = "Available"
    Case Else
        listitem.SubItems(2) = IIf(IsNull(m_cari("PRIOR")), "", m_cari("PRIOR"))
    End Select
    listitem.SubItems(3) = IIf(IsNull(m_cari("CUSTIDREF")), "", m_cari("CUSTIDREF"))
    listitem.SubItems(4) = IIf(IsNull(m_cari("NAMAREF")), "", m_cari("NAMAREF"))
    listitem.SubItems(5) = IIf(IsNull(m_cari("NAME")), "", m_cari("NAME"))
    listitem.SubItems(6) = IIf(IsNull(m_cari("NEXTACTDATE")), "", Format(m_cari("NEXTACTDATE"), "yyyy/mm/dd hh:mm"))
    listitem.SubItems(7) = IIf(IsNull(m_cari("NEXTACT")), "", m_cari("NEXTACT"))
    listitem.SubItems(8) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS"))
    listitem.SubItems(9) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
    listitem.SubItems(10) = IIf(IsNull(m_cari("NamaAGENT")), "", m_cari("NamaAGENT"))
    listitem.SubItems(11) = IIf(IsNull(m_cari("RECSOURCEREF")), "", m_cari("RECSOURCEREF"))
    listitem.SubItems(12) = Format(IIf(IsNull(m_cari("TGLSTATUS")), "", m_cari("TGLSTATUS")), "yyyy/mm/dd")
    listitem.SubItems(13) = IIf(IsNull(m_cari("Kethslkerja")), "", m_cari("Kethslkerja"))
    listitem.SubItems(14) = IIf(IsNull(m_cari("KdComplaint")), "", m_cari("KdComplaint"))
    listitem.SubItems(15) = IIf(IsNull(m_cari("RemarkComplaint")), "", m_cari("RemarkComplaint"))
    listitem.SubItems(16) = IIf(IsNull(m_cari("F_CEK")), "", m_cari("F_CEK"))
    listitem.SubItems(17) = IIf(IsNull(m_cari("Nomor")), "", m_cari("Nomor"))
    m_cari.MoveNext
    FOLLOWUPABIS = True
Wend
Set m_cari = Nothing
End Sub

Private Sub show_Search_mgmData()
Dim listitem As listitem
Dim m_cari As New ADODB.Recordset
Dim i As Integer
i = 1
On Error GoTo HELL
m_cari.CursorLocation = adUseClient
m_cari.Open "Select * from cc_custtbl where agent = '" + MDIForm1.Text1.Text + "' and (NEXTACTDATE BETWEEN '" + Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") & " 00:00" + "' AND '" + Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") & " 23:59" + "') ", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LstVwSearchMgm.ListItems.Clear
    Me.MousePointer = vbHourglass
'    ProgressBar1.Max = m_cari.RecordCount + 1
    While Not m_cari.EOF
 '   ProgressBar1.Value = m_cari.Bookmark
        Set listitem = LstVwSearchMgm.ListItems.ADD(, , m_cari.Bookmark)
        listitem.SubItems(1) = IIf(IsNull(m_cari("CUSTID")), "", m_cari("CUSTID"))
        listitem.SubItems(2) = IIf(IsNull(m_cari("PRIOR")), "", m_cari("PRIOR"))
        listitem.SubItems(3) = IIf(IsNull(m_cari("NAME")), "", m_cari("NAME"))
        listitem.SubItems(4) = IIf(IsNull(m_cari("NEXTACTDATE")), "", Format(m_cari("NEXTACTDATE"), "yyyy/mm/dd hh:nn"))
        listitem.SubItems(5) = IIf(IsNull(m_cari("NEXTACT")), "", m_cari("NEXTACT"))
        listitem.SubItems(6) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS"))
        listitem.SubItems(7) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
        listitem.SubItems(8) = IIf(IsNull(m_cari("NamaAGENT")), "", m_cari("NamaAGENT"))
        listitem.SubItems(9) = IIf(IsNull(m_cari("RECSOURCE")), "", m_cari("RECSOURCE"))
        listitem.SubItems(10) = IIf(IsNull(m_cari("TGLSTATUS")), "", Format(m_cari("TGLSTATUS"), "DD/MM/YYYY"))
        listitem.SubItems(11) = IIf(IsNull(m_cari("Kethslkerja")), "", m_cari("Kethslkerja"))
        listitem.SubItems(12) = IIf(IsNull(m_cari("KdComplaint")), "", m_cari("KdComplaint"))
        listitem.SubItems(13) = IIf(IsNull(m_cari("RemarkComplaint")), "", m_cari("RemarkComplaint"))
        listitem.SubItems(14) = IIf(IsNull(m_cari("F_CEK")), "", m_cari("F_CEK"))
        'LISTITEM.SubItems(15) = IIf(IsNull(m_cari("[NO]")), "", m_cari("[NO]"))
        m_cari.MoveNext
    Wend
        If LstVwSearchMgm.ListItems.count = 0 Then
            TxtJmlDtMgm.Text = "Tidak Ada Data"
        Else
            FOLLOWUPABIS = True
            TxtJmlDtMgm.Text = "Total " + CStr(m_cari.RecordCount) + " Records"
        End If
LstVwSearchMgm.SortKey = 2
LstVwSearchMgm.Sorted = True
'ProgressBar1.Value = 0
'ProgressBar1.Visible = False
MousePointer = vbNormal
Set m_cari = Nothing
Exit Sub
HELL:
    Me.MousePointer = vbNormal
    MsgBox Err.Description
  ''  Resume
End Sub


Private Sub ListView1_DblClick()
    If ListView1.ListItems.count = 0 Then
        Exit Sub
    End If
    Status_Form = 2
    TodayList = True
    FRMCUST_CC.Show vbModal
    Me.MousePointer = vbNormal
End Sub

Private Sub LstVwSearchMgm_DblClick()
    If LstVwSearchMgm.ListItems.count = 0 Then
        Exit Sub
    End If
    Status_Form = 2
    TodayList = True
    FRMCUST_CC_MGM.Show vbModal
    Me.MousePointer = vbNormal
End Sub
