VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmClaimVerSpv 
   Caption         =   "Verifikasi Claim Supervisor"
   ClientHeight    =   9870
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13455
   Icon            =   "frmClaimVerSpv.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   9870
   ScaleWidth      =   13455
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6225
      Left            =   135
      TabIndex        =   1
      Top             =   1665
      Visible         =   0   'False
      Width           =   13215
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "x"
         Height          =   285
         Left            =   12780
         TabIndex        =   3
         Top             =   135
         Width           =   300
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   5790
         Left            =   45
         TabIndex        =   2
         Top             =   315
         Width           =   12930
         _ExtentX        =   22807
         _ExtentY        =   10213
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   9765
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   13380
      _ExtentX        =   23601
      _ExtentY        =   17224
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu mnfile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mncari 
         Caption         =   "&Cari di data yang telah didistribusi"
      End
      Begin VB.Menu mncari1 
         Caption         =   "&Cari di data yang belum didistribusi"
      End
      Begin VB.Menu MnReject 
         Caption         =   "&Reject"
      End
   End
   Begin VB.Menu mnfile1 
      Caption         =   "File1"
      Visible         =   0   'False
      Begin VB.Menu mnver 
         Caption         =   "&Verifikasi"
      End
   End
End
Attribute VB_Name = "frmClaimVerSpv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "Claim Oleh(AOC)", 15 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Claim Oleh(Nama)", 15 * TXT
    ListView1.ColumnHeaders.ADD 3, , "Tanggal Entry", 10 * TXT
    ListView1.ColumnHeaders.ADD 4, , "Refference No", 10 * TXT
    ListView1.ColumnHeaders.ADD 5, , "Refference Name", 15 * TXT
    ListView1.ColumnHeaders.ADD 6, , "Telp Rumah", 15 * TXT
    ListView1.ColumnHeaders.ADD 7, , "Telp Kantor", 15 * TXT
    ListView1.ColumnHeaders.ADD 8, , "HandPhone", 15 * TXT
    ListView1.ColumnHeaders.ADD 9, , "Leads Name", 15 * TXT
    ListView1.ColumnHeaders.ADD 10, , "Agent(AOC)", 15 * TXT
    ListView1.ColumnHeaders.ADD 11, , "Agent(Nama)", 15 * TXT
    ListView1.ColumnHeaders.ADD 12, , "Status", 15 * TXT
End Sub
               

Private Sub ListView2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

If UCase(MDIForm1.Text2.Text) <> "SUPERVISOR" Then
    Exit Sub
End If
If Button = 2 Then
  Dim hItem As Long, nod As Node, xRect As RECT, xPop As Integer, yPop As Integer
  Set nod = ListView2.HitTest(x, y)

  If ListView2.ListItems.Count <> 0 Then
   TreeView_GetItemRect ListView2.hWnd, hItem, xRect, CTrue
 '  MsgBox xRect.Left & "," & xRect.Top & "," & xRect.Right & "," & xRect.Bottom
   xPop = Frame1.Left + ListView2.SelectedItem.Left + ScaleX(xRect.Left, vbPixels, vbTwips)
   yPop = Frame1.Top + ListView2.Top + ListView2.SelectedItem.Top + ScaleY(xRect.Bottom, vbPixels, vbTwips)
   PopupMenu mnfile1, , xPop, yPop
  End If
End If
End Sub

Private Sub MnReject_Click()
Dim cmdsql As String
Dim reason As String
reason = "Permintaan claim anda di tolak oleh " & MDIForm1.Text7.Text
cmdsql = "Update RequestInbound"
cmdsql = cmdsql + " Set UpdateOleh ='" + MDIForm1.Text7.Text + "', "
cmdsql = cmdsql + " Reason ='" + reason + "' , "
cmdsql = cmdsql + " Status = 2"
cmdsql = cmdsql + " where NOREF ='" + ListView1.SelectedItem.SubItems(3) + "'"
M_OBJCONN.Execute cmdsql
MsgBox "Proses Selesai", vbInformation + vbOKOnly, "Telegrandi"
ListView1.SelectedItem.SubItems(11) = "2"
End Sub

Private Sub Command1_Click()
    ListView2.ListItems.Clear
    Frame1.Visible = False
    ListView1.Enabled = True
End Sub

Private Sub Form_Load()
Dim m_objrs As New ADODB.Recordset
Dim LISTITEM As LISTITEM
Call HEADER_VIEW_SEARCH
Call header
m_objrs.CursorLocation = adUseClient
m_objrs.Open "Select * from RequestInbound where status =0", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not m_objrs.EOF
    Set LISTITEM = ListView1.ListItems.ADD(, , m_objrs("AGENTBaru"))
        LISTITEM.SubItems(1) = IIf(IsNull(m_objrs("NamaAGENTBaru")), "", m_objrs("NamaAGENTBaru"))
        LISTITEM.SubItems(2) = IIf(IsNull(m_objrs("TANGGAL")), "", Format(m_objrs("TANGGAL"), "mm/dd/yyyy"))
        LISTITEM.SubItems(3) = IIf(IsNull(m_objrs("NOREF")), "", m_objrs("NOREF"))
        LISTITEM.SubItems(4) = IIf(IsNull(m_objrs("NAME")), "", m_objrs("NAME"))
        LISTITEM.SubItems(5) = IIf(IsNull(m_objrs("NoTelp")), "", m_objrs("NoTelp"))
        LISTITEM.SubItems(6) = IIf(IsNull(m_objrs("NoTelpKantor")), "", m_objrs("NoTelpKantor"))
        LISTITEM.SubItems(7) = IIf(IsNull(m_objrs("NoHp")), "", m_objrs("NoHp"))
        LISTITEM.SubItems(8) = IIf(IsNull(m_objrs("LEADSNAME")), "", m_objrs("LEADSNAME"))
        LISTITEM.SubItems(9) = IIf(IsNull(m_objrs("AGENTlama")), "", m_objrs("AGENTlama"))
        LISTITEM.SubItems(10) = IIf(IsNull(m_objrs("NamaAGENTlama")), "", m_objrs("NamaAGENTlama"))
        LISTITEM.SubItems(11) = IIf(IsNull(m_objrs("Status")), "", m_objrs("Status"))
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Button = 2 Then
  Dim hItem As Long, nod As Node, xRect As RECT, xPop As Integer, yPop As Integer
  Set nod = ListView1.HitTest(x, y)

  If ListView1.ListItems.Count <> 0 Then
   TreeView_GetItemRect ListView1.hWnd, hItem, xRect, CTrue
 '  MsgBox xRect.Left & "," & xRect.Top & "," & xRect.Right & "," & xRect.Bottom
   xPop = ListView1.SelectedItem.Left + ScaleX(xRect.Left, vbPixels, vbTwips)
   yPop = ListView1.SelectedItem.Top + ScaleY(xRect.Bottom, vbPixels, vbTwips)
   PopupMenu MnFile, , xPop, yPop
  End If
End If
End Sub
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
   ListView1.SortKey = ColumnHeader.Index - 1
   ListView1.Sorted = True
End Sub
Private Sub MnCari_Click()
Dim m_objrs As New ADODB.Recordset
Dim LISTITEM As LISTITEM
ListView1.Enabled = False
m_objrs.CursorLocation = adUseClient
m_objrs.Open "Select * from MGM where nolap = '" + ListView1.SelectedItem.SubItems(3) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If m_objrs.RecordCount = 0 Then
    MsgBox "Data dengan nomor refference" & ListView1.SelectedItem.SubItems(3) & " tidak ditemukan " & vbCr & _
            " silahkan cari di data belum didistribute", vbInformation + vbOKOnly, "Telegrandi"
    Set m_objrs = Nothing
    ListView1.Enabled = True
    Exit Sub
End If
Frame1.Caption = mncari.Caption
ListView2.ListItems.Clear
While Not m_objrs.EOF
    Frame1.Visible = True
    Set LISTITEM = ListView2.ListItems.ADD(, , m_objrs.Bookmark)
            LISTITEM.SubItems(1) = IIf(IsNull(m_objrs("custid")), "", JADI_QUOTE(m_objrs("custid")))
            Select Case m_objrs("RECSTATUS")
            Case "1A"
                LISTITEM.SubItems(2) = "Available"
            Case Else
                LISTITEM.SubItems(2) = IIf(IsNull(m_objrs("PRIOR")), "", m_objrs("PRIOR"))
            End Select
            LISTITEM.SubItems(3) = IIf(IsNull(m_objrs("NEXTACTDATE")), "", Format(m_objrs("NEXTACTDATE"), "yyyy/mm/dd hh:mm"))
            LISTITEM.SubItems(4) = ""
            LISTITEM.SubItems(5) = IIf(IsNull(m_objrs("NAME")), "", JADI_QUOTE(m_objrs("NAME")))
            LISTITEM.SubItems(6) = IIf(IsNull(m_objrs("NEXTACT")), "", m_objrs("NEXTACT"))
            LISTITEM.SubItems(7) = IIf(IsNull(m_objrs("AGENT")), "", m_objrs("AGENT"))
            LISTITEM.SubItems(8) = IIf(IsNull(m_objrs("NamaAGENT")), "", m_objrs("NamaAGENT"))
            LISTITEM.SubItems(9) = IIf(IsNull(m_objrs("RECSOURCE")), "", m_objrs("RECSOURCE"))
            LISTITEM.SubItems(10) = Format(IIf(IsNull(m_objrs("TGLSTATUS")), "", m_objrs("TGLSTATUS")), "yyyy/mm/dd")
            LISTITEM.SubItems(11) = IIf(IsNull(m_objrs("StsLastCall")), "", m_objrs("StsLastCall"))
            LISTITEM.SubItems(12) = IIf(IsNull(m_objrs("KdComplaint")), "", m_objrs("KdComplaint"))
            LISTITEM.SubItems(13) = IIf(IsNull(m_objrs("RemarkComplaint")), "", m_objrs("RemarkComplaint"))
            LISTITEM.SubItems(14) = IIf(IsNull(m_objrs("HOMENO")), "", m_objrs("HOMENO"))
            LISTITEM.SubItems(15) = IIf(IsNull(m_objrs("OFFICENO")), "", m_objrs("OFFICENO"))
            LISTITEM.SubItems(16) = IIf(IsNull(m_objrs("EXTOFFICE")), "", m_objrs("EXTOFFICE"))
            LISTITEM.SubItems(17) = IIf(IsNull(m_objrs("MOBILENO")), "", m_objrs("MOBILENO"))
        m_objrs.MoveNext
Wend
    Set m_objrs = Nothing
End Sub

Private Sub mncari1_Click()
Dim LISTITEM As LISTITEM
Dim m_objrs As New ADODB.Recordset
ListView1.Enabled = False
m_objrs.CursorLocation = adUseClient
m_objrs.Open "Select * from tempCC_CUSTTBL where nolap = '" + ListView1.SelectedItem.SubItems(3) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If m_objrs.RecordCount = 0 Then
    MsgBox "Data dengan nomor refference" & ListView1.SelectedItem.SubItems(3) & " tidak ditemukan " & vbCr & _
            " silahkan cari di data yang telah didistribute", vbInformation + vbOKOnly, "Telegrandi"
    ListView1.Enabled = True
    Set m_objrs = Nothing
    Exit Sub
End If
Frame1.Caption = mncari1.Caption
ListView2.ListItems.Clear
While Not m_objrs.EOF
    Frame1.Visible = True
    Set LISTITEM = ListView2.ListItems.ADD(, , m_objrs.Bookmark)
            LISTITEM.SubItems(1) = IIf(IsNull(m_objrs("custid")), "", JADI_QUOTE(m_objrs("custid")))
            Select Case m_objrs("RECSTATUS")
            Case "1A"
                LISTITEM.SubItems(2) = "Available"
            Case Else
                LISTITEM.SubItems(2) = IIf(IsNull(m_objrs("PRIOR")), "", m_objrs("PRIOR"))
            End Select
            LISTITEM.SubItems(3) = IIf(IsNull(m_objrs("NEXTACTDATE")), "", Format(m_objrs("NEXTACTDATE"), "yyyy/mm/dd hh:mm"))
            LISTITEM.SubItems(4) = "*****"
            LISTITEM.SubItems(5) = IIf(IsNull(m_objrs("NAME")), "", JADI_QUOTE(m_objrs("NAME")))
            LISTITEM.SubItems(6) = IIf(IsNull(m_objrs("NEXTACT")), "", m_objrs("NEXTACT"))
            LISTITEM.SubItems(7) = IIf(IsNull(m_objrs("AGENT")), "", m_objrs("AGENT"))
            LISTITEM.SubItems(8) = IIf(IsNull(m_objrs("NamaAGENT")), "", m_objrs("NamaAGENT"))
            LISTITEM.SubItems(9) = IIf(IsNull(m_objrs("RECSOURCE")), "", m_objrs("RECSOURCE"))
            LISTITEM.SubItems(10) = "*****"
            LISTITEM.SubItems(11) = "*****"
            LISTITEM.SubItems(12) = "*****"
            LISTITEM.SubItems(13) = "*****"
            LISTITEM.SubItems(14) = IIf(IsNull(m_objrs("HOMENO")), "", m_objrs("HOMENO"))
            LISTITEM.SubItems(15) = IIf(IsNull(m_objrs("OFFICENO")), "", m_objrs("OFFICENO"))
            LISTITEM.SubItems(16) = IIf(IsNull(m_objrs("EXTOFFICE")), "", m_objrs("EXTOFFICE"))
            LISTITEM.SubItems(17) = IIf(IsNull(m_objrs("MOBILENO")), "", m_objrs("MOBILENO"))
        m_objrs.MoveNext
Wend
    Set m_objrs = Nothing
End Sub


Private Sub HEADER_VIEW_SEARCH()
    ListView2.ColumnHeaders.ADD 1, , "No", 3 * TXT
    ListView2.ColumnHeaders.ADD 2, , "Cust Id", 5 * TXT
    ListView2.ColumnHeaders.ADD 3, , "Priority", 5 * TXT
    ListView2.ColumnHeaders.ADD 4, , "Tgl Schedule", 10 * TXT
    ListView2.ColumnHeaders.ADD 5, , "Ref Number", 10 * TXT
    ListView2.ColumnHeaders.ADD 6, , "Nama Customer", 25 * TXT
    ListView2.ColumnHeaders.ADD 7, , "Next Action", 17 * TXT
    ListView2.ColumnHeaders.ADD 8, , "SalesCode", 8 * TXT
    ListView2.ColumnHeaders.ADD 9, , "Agent", 8 * TXT
    ListView2.ColumnHeaders.ADD 10, , "DataBase", 10 * TXT
    ListView2.ColumnHeaders.ADD 11, , "LastCall Date", 10 * TXT
    ListView2.ColumnHeaders.ADD 12, , "Sts LastCall", 10 * TXT
    ListView2.ColumnHeaders.ADD 13, , "Code", 5 * TXT
    ListView2.ColumnHeaders.ADD 14, , "Complaint Note", 15 * TXT
    ListView2.ColumnHeaders.ADD 15, , "Telp Rumah", 10 * TXT
    ListView2.ColumnHeaders.ADD 16, , "Telp Kantor", 10 * TXT
    ListView2.ColumnHeaders.ADD 17, , "Ext Telp", 10 * TXT
    ListView2.ColumnHeaders.ADD 18, , "HandPhone", 10 * TXT
End Sub

Private Sub mnver_Click()
Dim cmdsql As String
Dim reason As String
If Frame1.Caption = "&Cari di data yang telah didistribusi" Then

    cmdsql = "Update MGM SET AGENT ='" + ListView1.SelectedItem.Text + "'"
    cmdsql = cmdsql + " where NOLAP ='" + ListView1.SelectedItem.SubItems(3) + "'"
    M_OBJCONN.Execute cmdsql
    
    reason = "Permintaan claim anda di Setujui Oleh " & MDIForm1.Text7.Text
    cmdsql = "Update RequestInbound"
    cmdsql = cmdsql + " Set UpdateOleh ='" + MDIForm1.Text7.Text + "', "
    cmdsql = cmdsql + " Reason ='" + reason + "' , "
    cmdsql = cmdsql + " Status = 1"
    cmdsql = cmdsql + " where NOREF ='" + ListView1.SelectedItem.SubItems(3) + "'"
    M_OBJCONN.Execute cmdsql
Else
    cmdsql = "Insert Into MGM"
    cmdsql = cmdsql + " (CUSTID,"
    cmdsql = cmdsql + " NAME,"
    cmdsql = cmdsql + " HOMENO,"
    cmdsql = cmdsql + " MOBILENO,"
    cmdsql = cmdsql + " OFFICENO,"
    cmdsql = cmdsql + " EXTOFFICE,"
    cmdsql = cmdsql + " NOLAP,"
    cmdsql = cmdsql + " StsLastCall,"
    cmdsql = cmdsql + " Recstatus,"
    cmdsql = cmdsql + " NAMAAGENT,"
    cmdsql = cmdsql + " agent)"
    cmdsql = cmdsql + " Values"
    cmdsql = cmdsql + " ('" + ListView2.SelectedItem.SubItems(1) + "',"
    cmdsql = cmdsql + " '" + ListView2.SelectedItem.SubItems(5) + "',"
    cmdsql = cmdsql + " '" + ListView2.SelectedItem.SubItems(14) + "',"
    cmdsql = cmdsql + " '" + ListView2.SelectedItem.SubItems(17) + "',"
    cmdsql = cmdsql + " '" + ListView2.SelectedItem.SubItems(15) + "',"
    cmdsql = cmdsql + " '" + ListView2.SelectedItem.SubItems(16) + "',"
    cmdsql = cmdsql + " '" + ListView2.SelectedItem.SubItems(1) + "',"
    cmdsql = cmdsql + " '1A',"
    cmdsql = cmdsql + " '1A',"
    cmdsql = cmdsql + " '" + ListView1.SelectedItem.SubItems(1) + "',"
    cmdsql = cmdsql + " '" + ListView1.SelectedItem.Text + "')"
    M_OBJCONN.Execute cmdsql
    M_OBJCONN.Execute "Delete from tempCC_CUSTTBL where NOLAP = '" + ListView2.SelectedItem.SubItems(1) + "'"
    reason = "Permintaan claim anda disetujui Oleh " & MDIForm1.Text7.Text
    cmdsql = "Update RequestInbound"
    cmdsql = cmdsql + " Set UpdateOleh ='" + MDIForm1.Text7.Text + "', "
    cmdsql = cmdsql + " Reason ='" + reason + "' , "
    cmdsql = cmdsql + " Status = 1"
    cmdsql = cmdsql + " where NOREF ='" + ListView1.SelectedItem.SubItems(3) + "'"
    M_OBJCONN.Execute cmdsql
End If
MsgBox "Proses Selesai", vbInformation + vbOKOnly, "Telegrandi"
ListView1.SelectedItem.SubItems(11) = "1"
ListView1.Enabled = True
ListView2.ListItems.Clear
Frame1.Visible = False
End Sub
