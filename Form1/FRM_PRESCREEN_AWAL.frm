VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FRM_PRESCREEN_AWAL 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Schedule Reminder"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11535
   ForeColor       =   &H00000000&
   Icon            =   "FRM_PRESCREEN_AWAL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   8595
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11490
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Cancel          =   -1  'True
         Caption         =   "&Tutup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   10530
         TabIndex        =   2
         Top             =   150
         Width           =   870
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   8415
         Left            =   30
         TabIndex        =   1
         Top             =   120
         Width           =   10470
         _ExtentX        =   18468
         _ExtentY        =   14843
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "0 data"
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   10710
         TabIndex        =   3
         Top             =   8250
         Width           =   705
      End
   End
End
Attribute VB_Name = "FRM_PRESCREEN_AWAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub HEADER_PRESCREEN()
    ListView1.ColumnHeaders.ADD 1, , "Customers Id", 15 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Level", 10 * TXT
    ListView1.ColumnHeaders.ADD 3, , "Customers Name", 40 * TXT
    ListView1.ColumnHeaders.ADD 4, , "Alamat", 15 * TXT
    ListView1.ColumnHeaders.ADD 5, , "Tanggal Lahir", 15 * TXT
    ListView1.ColumnHeaders.ADD 6, , "Next Action Date", 15 * TXT
    ListView1.ColumnHeaders.ADD 7, , "Next Action", 50 * TXT
    ListView1.ColumnHeaders.ADD 8, , "Data Source", 15 * TXT
    ListView1.ColumnHeaders.ADD 9, , "Agent", 15 * TXT
    ListView1.ColumnHeaders.ADD 10, , "LastCall Date", 15 * TXT
    ListView1.ColumnHeaders.ADD 11, , "Sts LastCall", 15 * TXT
    ListView1.ColumnHeaders.ADD 12, , "Complaint Code", 15 * TXT
    ListView1.ColumnHeaders.ADD 13, , "Complaint Note", 15 * TXT
End Sub

Private Sub Command1_Click(Index As Integer)
Dim M_DATA As New CLS_FRMSEARCH
Dim m_objrs As ADODB.Recordset
Dim PANJANG As Integer
Select Case Index
    Case 0
        Unload Me
End Select
Set M_DATA = Nothing
Set m_objrs = Nothing
End Sub

Private Sub Form_Load()
Dim LISTITEM As LISTITEM
Dim m_objrs As ADODB.Recordset
Dim M_DATA As New CLS_FRMSEARCH
MDIForm1.Text5.Visible = True
MDIForm1.ProgressBar1.Visible = True
Call HEADER_PRESCREEN
Set m_objrs = M_DATA.QUERY_SEARCH(M_OBJCONN, "AGENT = '" + MDIForm1.Text1.Text + "' AND (NEXTACTDATE < '" + Format(MDIForm1.TDBDate1.Value & " 23:59", "mm/dd/yyyy hh:mm") + "' AND NEXTACTDATE > '" + Format(MDIForm1.TDBDate1.Value & " 00:00", "mm/dd/yyyy hh:mm") + "') AND LEFT(RECSTATUS,1) <> '0' AND LEFT(RECSTATUS,2) <> '3A' AND LEFT(RECSTATUS,2) <> 'XX'", MDIForm1.Text3.Text)
MDIForm1.ProgressBar1.Max = m_objrs.RecordCount + 1
Label1.Caption = m_objrs.RecordCount & " Data"
    While Not m_objrs.EOF
    MDIForm1.ProgressBar1.Value = m_objrs.Bookmark
    Set LISTITEM = ListView1.ListItems.ADD(, , IIf(IsNull(m_objrs("CUSTID")), "", JADI_QUOTE(m_objrs("CUSTID"))))
        Select Case m_objrs("RECSTATUS")
        Case "1A"
            LISTITEM.SubItems(1) = "Available"
        Case Else
            LISTITEM.SubItems(1) = IIf(IsNull(m_objrs("PRIOR")), "", JADI_QUOTE(m_objrs("PRIOR")))
        End Select
        LISTITEM.SubItems(2) = IIf(IsNull(m_objrs("NAME")), "", JADI_QUOTE(m_objrs("NAME")))
        LISTITEM.SubItems(3) = IIf(IsNull(m_objrs("ADDRNOW")), "", JADI_QUOTE(m_objrs("ADDRNOW")))
        LISTITEM.SubItems(4) = IIf(IsNull(m_objrs("BIRTHD")), "", Format(m_objrs("BIRTHD"), "dd-mmm-yyyy"))
        LISTITEM.SubItems(5) = IIf(IsNull(m_objrs("NEXTACTDATE")), "", Format(m_objrs("NEXTACTDATE"), "yyyy/mm/dd hh:mm"))
        LISTITEM.SubItems(6) = IIf(IsNull(m_objrs("NEXTACT")), "", m_objrs("NEXTACT"))
        LISTITEM.SubItems(7) = IIf(IsNull(m_objrs("RECSOURCE")), "", m_objrs("RECSOURCE"))
        LISTITEM.SubItems(8) = IIf(IsNull(m_objrs("AGENT")), "", m_objrs("AGENT"))
        LISTITEM.SubItems(9) = IIf(IsNull(m_objrs("TglStatus")), "", m_objrs("TglStatus"))
        LISTITEM.SubItems(10) = IIf(IsNull(m_objrs("StsLastCall")), "", m_objrs("StsLastCall"))
        LISTITEM.SubItems(11) = IIf(IsNull(m_objrs("KdComplaint")), "", m_objrs("KdComplaint"))
        LISTITEM.SubItems(12) = IIf(IsNull(m_objrs("RemarkComplaint")), "", m_objrs("RemarkComplaint"))
        
    m_objrs.MoveNext
    Wend
SCHEDULE_VIEW = False
Set m_objrs = Nothing
Set M_DATA = Nothing
MDIForm1.Text5.Text = Empty
MDIForm1.Text5.Visible = False
MDIForm1.ProgressBar1.Value = 0
MDIForm1.ProgressBar1.Visible = False
End Sub

Private Sub ListView1_DblClick()
If ListView1.ListItems.Count = 0 Then
    Exit Sub
End If
    ADD_CUST = False
    VIEW_OK = False
    SCREENER_AWAL = True
    SCREENER = False
    REBUT_DATA = False
    VIEW_AVAIL_AWAL = False
    SCREENER_APPROV = False
    SCR_SPV_CARI = False
    Select Case UCase(MDIForm1.Text3.Text)
    Case "CREDIT CARD"
        FRMCUST_CC.Show vbModal
    Case Else
    End Select
End Sub
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
   ListView1.SortKey = ColumnHeader.Index - 1
   ListView1.Sorted = True
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If ListView1.ListItems.Count = 0 Then
    Exit Sub
End If
If KeyAscii = 13 Then
    ADD_CUST = False
    VIEW_OK = False
    SCREENER_AWAL = True
    SCREENER = False
    VIEW_AVAIL_AWAL = False
    SCREENER_APPROV = False
    REBUT_DATA = False
    SCR_SPV_CARI = False
    Select Case UCase(MDIForm1.Text3.Text)
    Case "CREDIT CARD"
        FRMCUST_CC.Show vbModal
    Case Else
    End Select
End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Command1_Click(0)
End If
End Sub
