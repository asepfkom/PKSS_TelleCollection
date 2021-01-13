VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form VIEWCUSTAVAIL_AGENT 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   9780
   ClientLeft      =   -3345
   ClientTop       =   450
   ClientWidth     =   11805
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   Icon            =   "VIEWCUSTAVAIL_AGENT.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9780
   ScaleWidth      =   11805
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
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
      Left            =   10935
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   10350
      Width           =   3045
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14025
      TabIndex        =   2
      Top             =   10305
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   10245
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Width           =   15195
      Begin MSComctlLib.ListView ListView1 
         Height          =   10080
         Left            =   15
         TabIndex        =   1
         Top             =   135
         Width           =   15135
         _ExtentX        =   26696
         _ExtentY        =   17780
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
      Height          =   375
      Left            =   300
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   765
      Visible         =   0   'False
      Width           =   9060
   End
End
Attribute VB_Name = "VIEWCUSTAVAIL_AGENT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub HEADER_VIEW_ALL()
    ListView1.ColumnHeaders.ADD 1, , "No", 3 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Cust Id", 5 * TXT
    ListView1.ColumnHeaders.ADD 3, , "Priority", 5 * TXT
    ListView1.ColumnHeaders.ADD 4, , "Tgl Schedule", 10 * TXT
    ListView1.ColumnHeaders.ADD 5, , "Nama Customer", 20 * TXT
    ListView1.ColumnHeaders.ADD 6, , "Next Action", 17 * TXT
    ListView1.ColumnHeaders.ADD 7, , "Team Leader", 15 * TXT
    ListView1.ColumnHeaders.ADD 8, , "Agent", 15 * TXT
    ListView1.ColumnHeaders.ADD 9, , "Data Source", 15 * TXT
    ListView1.ColumnHeaders.ADD 10, , "Sts LastCall", 15 * TXT
    ListView1.ColumnHeaders.ADD 11, , "Complaint Code", 15 * TXT
    ListView1.ColumnHeaders.ADD 12, , "Complaint Note", 15 * TXT
End Sub
  
Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim m_objrs As ADODB.Recordset
Dim M_DATA As New VIEW
Dim LISTITEM As LISTITEM
Dim M_AGENT As String
Dim M_DATAS As String
Dim M_SPV As String
Dim i As Integer
i = 1
On Error GoTo HELL
Call HEADER_VIEW_ALL
    Text2.Text = "View All"
    MDIForm1.ProgressBar1.Visible = True
    If B_AVAIL = True Then
        Me.Caption = "Tampilkan Data Available"
        Set m_objrs = M_DATA.QUERY_VIEW_ALL_AVAILABLE(M_OBJCONN, "USERTBL.USERID = '" + MDIForm1.Text1.Text + "'", " NAME", MDIForm1.Text3.Text)
    End If
    If B_INCOMING = True Then
        Me.Caption = "Tampilkan Data Incoming"
        Set m_objrs = M_DATA.QUERY_INCOMING(M_OBJCONN, "USERTBL.USERID = '" + MDIForm1.Text1.Text + "'", " NAME", MDIForm1.Text3.Text)
    End If
    MDIForm1.ProgressBar1.Max = m_objrs.RecordCount + 1
    While Not m_objrs.EOF
    MDIForm1.ProgressBar1.Value = m_objrs.Bookmark
        Set LISTITEM = ListView1.ListItems.ADD(, , m_objrs.Bookmark)
        LISTITEM.SubItems(1) = IIf(IsNull(m_objrs("CUSTID")), "", m_objrs("CUSTID"))
        Select Case IIf(IsNull(m_objrs("RECSTATUS")), "", m_objrs("RECSTATUS"))
            Case "XX"
                LISTITEM.SubItems(2) = "Reject"
            Case "3A"
                LISTITEM.SubItems(2) = "Approval"
            Case "1A"
                LISTITEM.SubItems(2) = "Available"
            Case "2C"
                LISTITEM.SubItems(2) = "Incoming"
        End Select
        LISTITEM.SubItems(3) = IIf(IsNull(m_objrs("NEXTACTDATE")), "", Format(m_objrs("NEXTACTDATE"), "yyyy/mm/dd hh:mm"))
        LISTITEM.SubItems(4) = IIf(IsNull(m_objrs("NAME")), "", m_objrs("NAME"))
        LISTITEM.SubItems(5) = IIf(IsNull(m_objrs("NEXTACT")), "", m_objrs("NEXTACT"))
        LISTITEM.SubItems(6) = IIf(IsNull(m_objrs("SPVNAME")), "", m_objrs("SPVNAME"))
        LISTITEM.SubItems(7) = IIf(IsNull(m_objrs("AGENT")), "", m_objrs("AGENT"))
        LISTITEM.SubItems(8) = IIf(IsNull(m_objrs("RECSOURCE")), "", m_objrs("RECSOURCE"))
        LISTITEM.SubItems(9) = IIf(IsNull(m_objrs("StsLastCall")), "", m_objrs("StsLastCall"))
        LISTITEM.SubItems(10) = IIf(IsNull(m_objrs("KdComplaint")), "", m_objrs("KdComplaint"))
        LISTITEM.SubItems(11) = IIf(IsNull(m_objrs("RemarkComplaint")), "", m_objrs("RemarkComplaint"))
        m_objrs.MoveNext
        Wend
    If ListView1.ListItems.Count = 0 Then
        If B_AVAIL = True Then
            Text1.Text = "Tidak Ada Data Available"
        Else
            Text1.Text = "Tidak Ada Data Incoming"
        End If
    Else
        Text1.Text = "Total " + CStr(m_objrs.RecordCount) + " Records"
    End If
ListView1.SortKey = 2
ListView1.Sorted = True
MDIForm1.ProgressBar1.Value = 0
MDIForm1.ProgressBar1.Visible = False
Set M_DATA = Nothing
Set m_objrs = Nothing
Exit Sub
HELL:
    Set M_DATA = Nothing
    Set m_objrs = Nothing
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
   ListView1.SortKey = ColumnHeader.Index - 1
   ListView1.Sorted = True
End Sub

Private Sub ListView1_DblClick()
If ListView1.ListItems.Count = 0 Then
    Exit Sub
End If
Select Case UCase(MDIForm1.Text3.Text)
Case "CREDIT CARD"
    VIEW_OK = False
    SCREENER_AWAL = False
    SCREENER = False
    VIEW_AVAIL_AWAL = True
    SCREENER_APPROV = False
    REBUT_DATA = False
    FRMCUST_CC.Show vbModal
End Select
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call ListView1_DblClick
End If
End Sub
