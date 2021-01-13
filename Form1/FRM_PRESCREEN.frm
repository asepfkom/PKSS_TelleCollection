VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FRM_PRESCREEN 
   BackColor       =   &H80000004&
   ClientHeight    =   9480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14280
   ForeColor       =   &H00000000&
   Icon            =   "FRM_PRESCREEN.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   632
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   952
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
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
      Height          =   315
      Index           =   0
      Left            =   14370
      TabIndex        =   2
      Top             =   10170
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   10170
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Width           =   15195
      Begin MSComctlLib.ListView ListView1 
         Height          =   10005
         Left            =   30
         TabIndex        =   1
         Top             =   120
         Width           =   15105
         _ExtentX        =   26644
         _ExtentY        =   17648
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "0 data"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   13800
      TabIndex        =   3
      Top             =   10230
      Width           =   450
   End
   Begin VB.Menu mnClaim 
      Caption         =   "Claim"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "FRM_PRESCREEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub HEADER_VIEW_SEARCH()
    ListView1.ColumnHeaders.ADD 1, , "No", 3 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Cust Id", 5 * TXT
    ListView1.ColumnHeaders.ADD 3, , "Priority", 5 * TXT
    ListView1.ColumnHeaders.ADD 4, , "Ref Id", 10 * TXT
    ListView1.ColumnHeaders.ADD 5, , "Ref Name", 10 * TXT
    ListView1.ColumnHeaders.ADD 6, , "Nama Customer", 25 * TXT
    ListView1.ColumnHeaders.ADD 7, , "Tgl Schedule", 10 * TXT
    ListView1.ColumnHeaders.ADD 8, , "Next Action", 17 * TXT
    ListView1.ColumnHeaders.ADD 9, , "Remarks", 17 * TXT
    ListView1.ColumnHeaders.ADD 10, , "SalesCode", 8 * TXT
    ListView1.ColumnHeaders.ADD 11, , "Agent", 8 * TXT
    ListView1.ColumnHeaders.ADD 12, , "DataBase", 10 * TXT
    ListView1.ColumnHeaders.ADD 13, , "LastCall Date", 10 * TXT
    ListView1.ColumnHeaders.ADD 14, , "Sts LastCall", 10 * TXT
    ListView1.ColumnHeaders.ADD 15, , "Code", 5 * TXT
    ListView1.ColumnHeaders.ADD 16, , "Complaint Note", 15 * TXT
    ListView1.ColumnHeaders.ADD 17, , "Check", 10 * TXT
End Sub

Private Sub HEADER_PRESCREEN()
    ListView1.ColumnHeaders.ADD 1, , "No", 3 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Cust Id", 5 * TXT
    ListView1.ColumnHeaders.ADD 3, , "Priority", 5 * TXT
    ListView1.ColumnHeaders.ADD 4, , "Ref Id", 10 * TXT
    ListView1.ColumnHeaders.ADD 5, , "Ref Name", 10 * TXT
    ListView1.ColumnHeaders.ADD 6, , "Nama Customer", 25 * TXT
    ListView1.ColumnHeaders.ADD 7, , "Tgl Schedule", 10 * TXT
    ListView1.ColumnHeaders.ADD 8, , "Next Action", 17 * TXT
    ListView1.ColumnHeaders.ADD 9, , "Remarks", 17 * TXT
    ListView1.ColumnHeaders.ADD 10, , "SalesCode", 8 * TXT
    ListView1.ColumnHeaders.ADD 11, , "Agent", 8 * TXT
    ListView1.ColumnHeaders.ADD 12, , "DataBase", 10 * TXT
    ListView1.ColumnHeaders.ADD 13, , "LastCall Date", 10 * TXT
    ListView1.ColumnHeaders.ADD 14, , "Sts LastCall", 10 * TXT
    ListView1.ColumnHeaders.ADD 15, , "Code", 5 * TXT
    ListView1.ColumnHeaders.ADD 16, , "Complaint Note", 15 * TXT
    ListView1.ColumnHeaders.ADD 17, , "Check", 10 * TXT
End Sub

Private Sub Command1_Click(Index As Integer)
Dim M_DATA As New CLS_FRMSEARCH
Dim m_objrs As ADODB.Recordset
Select Case Index
    Case 0
        Unload Me
End Select
Set M_DATA = Nothing
Set m_objrs = Nothing
End Sub

Private Sub view_search()
Dim listitem As listitem
Call HEADER_VIEW_SEARCH

With FRM_SEARCH
    .Height = 4815
    .Frame1.Visible = True

FRM_SEARCH.ProgressBar1.Max = .m_cari.RecordCount + 1
Label1.Caption = .m_cari.RecordCount & " Data"
While Not .m_cari.EOF
FRM_SEARCH.ProgressBar1.Value = .m_cari.Bookmark
Set listitem = ListView1.ListItems.ADD(, , .m_cari.Bookmark)
    listitem.SubItems(1) = IIf(IsNull(.m_cari("custid")), "", .m_cari("custid"))
    Select Case .m_cari("RECSTATUS")
    Case "1A"
        listitem.SubItems(2) = "Available"
    Case ""
        listitem.SubItems(2) = "Available"
    Case Else
        listitem.SubItems(2) = IIf(IsNull(.m_cari("PRIOR")), "", .m_cari("PRIOR"))
    End Select
    listitem.SubItems(3) = IIf(IsNull(.m_cari("CUSTIDREF")), "", .m_cari("CUSTIDREF"))
    listitem.SubItems(4) = IIf(IsNull(.m_cari("NAMAREF")), "", .m_cari("NAMAREF"))
    listitem.SubItems(5) = IIf(IsNull(.m_cari("NAME")), "", JADI_QUOTE(.m_cari("NAME")))
    listitem.SubItems(6) = IIf(IsNull(.m_cari("NEXTACTDATE")), "", Format(.m_cari("NEXTACTDATE"), "yyyy/mm/dd hh:mm"))
    listitem.SubItems(7) = IIf(IsNull(.m_cari("NEXTACT")), "", .m_cari("NEXTACT"))
    listitem.SubItems(8) = IIf(IsNull(.m_cari("REMARKS")), "", .m_cari("REMARKS"))
    listitem.SubItems(9) = IIf(IsNull(.m_cari("AGENT")), "", .m_cari("AGENT"))
    listitem.SubItems(10) = IIf(IsNull(.m_cari("NamaAGENT")), "", .m_cari("NamaAGENT"))
    listitem.SubItems(11) = IIf(IsNull(.m_cari("RECSOURCEREF")), "", .m_cari("RECSOURCEREF"))
    listitem.SubItems(12) = Format(IIf(IsNull(.m_cari("TGLSTATUS")), "", .m_cari("TGLSTATUS")), "yyyy/mm/dd")
    listitem.SubItems(13) = IIf(IsNull(.m_cari("Kethslkerja")), "", .m_cari("Kethslkerja"))
    listitem.SubItems(14) = IIf(IsNull(.m_cari("KdComplaint")), "", .m_cari("KdComplaint"))
    listitem.SubItems(15) = IIf(IsNull(.m_cari("RemarkComplaint")), "", .m_cari("RemarkComplaint"))
    listitem.SubItems(16) = IIf(IsNull(.m_cari("F_CEK")), "", .m_cari("F_CEK"))
.m_cari.MoveNext
Wend
End With
Unload FRM_SEARCH
End Sub

Private Sub Form_Load()
Dim listitem As listitem
Dim m_objrs As ADODB.Recordset
Dim M_DATA As New CLS_FRMSEARCH
Dim TGL_AKTION As String

Me.Left = 30
Me.Top = 50
Me.Width = 14400
Me.Height = 9885

If search_ok Then
    Call view_search
    Exit Sub
Else
Call HEADER_PRESCREEN


FRM_SCHEDULE.Frame1.Visible = True
FRM_SCHEDULE.Height = 4785
TGL_AKTION = MDIForm1.TDBDate1.Value + 1

If StsMgmSchedule = True Then
   Set m_objrs = M_DATA.QUERY_SEARCH_mgm(M_OBJCONN, "AGENT = '" + FRM_SCHEDULE.Combo1(0).Text + "' AND (NEXTACTDATE BETWEEN '" + Format(FRM_SCHEDULE.TDBDate1(0).Value & " 00:00", "mm/dd/yyyy hh:mm") + "' AND '" + Format(FRM_SCHEDULE.TDBDate1(1).Text & " 23:59", "mm/dd/yyyy hh:mm") + "') ", MDIForm1.Text3.Text)
Else
   Set m_objrs = M_DATA.QUERY_SEARCH(M_OBJCONN, "AGENT = '" + FRM_SCHEDULE.Combo1(0).Text + "' AND (NEXTACTDATE >= '" + Format(FRM_SCHEDULE.TDBDate1(0).Value & " 00:00", "mm/dd/yyyy hh:mm") + "' AND NEXTACTDATE <= '" + Format(FRM_SCHEDULE.TDBDate1(1).Text & " 23:59", "mm/dd/yyyy hh:mm") + "') ", MDIForm1.Text3.Text)
End If
    FRM_SCHEDULE.ProgressBar1.Max = m_objrs.RecordCount + 1
    Label1.Caption = m_objrs.RecordCount & " Data"
    While Not m_objrs.EOF
    FRM_SCHEDULE.ProgressBar1.Value = m_objrs.Bookmark
    Set listitem = ListView1.ListItems.ADD(, , m_objrs.Bookmark)
    listitem.SubItems(1) = IIf(IsNull(m_objrs("custid")), "", JADI_QUOTE(m_objrs("custid")))
    Select Case m_objrs("RECSTATUS")
    Case "1A"
        listitem.SubItems(2) = "Available"
    Case ""
        listitem.SubItems(2) = "Available"
    Case Else
        listitem.SubItems(2) = IIf(IsNull(m_objrs("PRIOR")), "", m_objrs("PRIOR"))
    End Select
    
    If FRM_SCHEDULE.Check1.Value = 1 Then
        listitem.SubItems(3) = "MGM-DATA"
    Else
        listitem.SubItems(3) = IIf(IsNull(m_objrs("CUSTIDREF")), "", m_objrs("CUSTIDREF"))
    End If
    If FRM_SCHEDULE.Check1.Value = 1 Then
        listitem.SubItems(4) = "MGM-DATA"
    Else
        listitem.SubItems(4) = IIf(IsNull(m_objrs("NAMAREF")), "", m_objrs("NAMAREF"))
    End If
    listitem.SubItems(5) = IIf(IsNull(m_objrs("NAME")), "", JADI_QUOTE(m_objrs("NAME")))
    listitem.SubItems(6) = IIf(IsNull(m_objrs("NEXTACTDATE")), "", Format(m_objrs("NEXTACTDATE"), "yyyy/mm/dd hh:mm"))
    listitem.SubItems(7) = IIf(IsNull(m_objrs("NEXTACT")), "", m_objrs("NEXTACT"))
    listitem.SubItems(8) = IIf(IsNull(m_objrs("REMARKS")), "", m_objrs("REMARKS"))
    listitem.SubItems(9) = IIf(IsNull(m_objrs("AGENT")), "", m_objrs("AGENT"))
    listitem.SubItems(10) = IIf(IsNull(m_objrs("NamaAGENT")), "", m_objrs("NamaAGENT"))
    If FRM_SCHEDULE.Check1.Value = 1 Then
        listitem.SubItems(11) = IIf(IsNull(m_objrs("RECSOURCE")), "", m_objrs("RECSOURCE"))
    Else
        listitem.SubItems(11) = IIf(IsNull(m_objrs("RECSOURCEREF")), "", m_objrs("RECSOURCEREF"))
    End If
    listitem.SubItems(12) = Format(IIf(IsNull(m_objrs("TGLSTATUS")), "", m_objrs("TGLSTATUS")), "yyyy/mm/dd")
    listitem.SubItems(13) = IIf(IsNull(m_objrs("Kethslkerja")), "", m_objrs("Kethslkerja"))
    listitem.SubItems(14) = IIf(IsNull(m_objrs("KdComplaint")), "", m_objrs("KdComplaint"))
    listitem.SubItems(15) = IIf(IsNull(m_objrs("RemarkComplaint")), "", m_objrs("RemarkComplaint"))
    listitem.SubItems(16) = IIf(IsNull(m_objrs("F_CEK")), "", m_objrs("F_CEK"))
    
    m_objrs.MoveNext
    Wend
    FRM_SCHEDULE.ProgressBar1.Value = FRM_SCHEDULE.ProgressBar1.Max
    Unload FRM_SCHEDULE
End If

Set m_objrs = Nothing
Set M_DATA = Nothing
End Sub

Private Sub ListView1_DblClick()
Dim m_objrs As ADODB.Recordset
If ListView1.ListItems.count = 0 Then
    Exit Sub
End If
Status_Form = 1
    If UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
        Set m_objrs = New ADODB.Recordset
        m_objrs.CursorLocation = adUseClient
        m_objrs.Open "SELECT USERID FROM USERTBL WHERE SPVCODE ='" + MDIForm1.Text1.Text + "' AND USERID = '" + ListView1.SelectedItem.SubItems(9) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If m_objrs.RecordCount <> 0 Then
        Else
            MsgBox "Data Ini Milik Agent Team Leader Yang Lain", vbCritical + vbOKOnly, "TeleGrandi"
            Set m_objrs = Nothing
            Exit Sub
        End If
        Set m_objrs = Nothing
    Else
        If UCase(MDIForm1.Text2.Text) = "SUPERVISOR" Or UCase(MDIForm1.Text2.Text) = "ADMINISTRATOR" Then
        Else
            If Trim(UCase(MDIForm1.Text1.Text)) = Trim(UCase(ListView1.SelectedItem.SubItems(9))) Then
            Else
                MsgBox "Data Ini Milik Agent Yang Lain", vbCritical + vbOKOnly, "TeleGrandi"
                Set m_objrs = Nothing
                Exit Sub
            End If
        End If
    End If
If StsMgmSchedule = False Then
    FRMCUST_CC.Show vbModal
Else
    Flag_Mgm = True
    FRMCUST_CC_MGM.Show vbModal
End If
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
   ListView1.SortKey = ColumnHeader.Index - 1
   ListView1.Sorted = True
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If ListView1.ListItems.count = 0 Then
    Exit Sub
End If
If KeyAscii = 13 Then
    Call ListView1_DblClick
End If
End Sub

