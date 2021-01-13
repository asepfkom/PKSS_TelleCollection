VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRM_VER_INBOUND 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Verifikasi Inbound Data Agent"
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13725
   ForeColor       =   &H00000000&
   Icon            =   "FRM_VER_INBOUND.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9645
   ScaleWidth      =   13725
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   11070
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   15090
      Begin VB.CommandButton Command2 
         Caption         =   "&Reject"
         Height          =   375
         Index           =   3
         Left            =   3780
         TabIndex        =   6
         Top             =   255
         Width           =   1650
      End
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "&Tutup"
         Height          =   375
         Index           =   2
         Left            =   5460
         TabIndex        =   5
         Top             =   255
         Width           =   1650
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Verifikasi"
         Height          =   375
         Index           =   1
         Left            =   2085
         TabIndex        =   4
         Top             =   255
         Width           =   1650
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cari Data"
         Height          =   375
         Index           =   0
         Left            =   375
         TabIndex        =   3
         Top             =   255
         Width           =   1650
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   10215
         Left            =   30
         TabIndex        =   1
         Top             =   810
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   18018
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0 = Belum Diproses  1 = Cek Valid   2 = Di Reject / Di Hapus"
         Height          =   360
         Left            =   8745
         TabIndex        =   7
         Top             =   270
         Width           =   4680
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "0 data"
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   12795
         TabIndex        =   2
         Top             =   8235
         Width           =   705
      End
   End
End
Attribute VB_Name = "FRM_VER_INBOUND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub HEADER_PRESCREEN()
    ListView1.ColumnHeaders.ADD 1, , "Customers Id", 15 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Sts Verifikasi", 10 * TXT
    ListView1.ColumnHeaders.ADD 3, , "SalesCode", 10 * TXT
    ListView1.ColumnHeaders.ADD 4, , "Nama Agent", 10 * TXT
    ListView1.ColumnHeaders.ADD 5, , "Customers Name", 20 * TXT
    ListView1.ColumnHeaders.ADD 6, , "Alamat Rumah", 15 * TXT
    ListView1.ColumnHeaders.ADD 7, , "Kantor", 15 * TXT
    ListView1.ColumnHeaders.ADD 8, , "Alamat Kantor", 15 * TXT
    ListView1.ColumnHeaders.ADD 9, , "Home Telp", 10 * TXT
    ListView1.ColumnHeaders.ADD 10, , "Home Telp2", 10 * TXT
    ListView1.ColumnHeaders.ADD 11, , "Office Telp", 10 * TXT
    ListView1.ColumnHeaders.ADD 12, , "Office Telp2", 10 * TXT
    ListView1.ColumnHeaders.ADD 13, , "Fax", 10 * TXT
    ListView1.ColumnHeaders.ADD 14, , "Fax2", 10 * TXT
    ListView1.ColumnHeaders.ADD 15, , "Hp", 10 * TXT
    ListView1.ColumnHeaders.ADD 16, , "Hp2", 10 * TXT
    ListView1.ColumnHeaders.ADD 17, , "Database", 10 * TXT
    ListView1.ColumnHeaders.ADD 18, , "Tgl Entry", 10 * TXT
End Sub


Private Sub Command2_Click(Index As Integer)
Select Case Index
    Case 0
        FRM_VER_SEARCH.Show
    Case 1
        FRM_VER_OK.Label1.Caption = ListView1.SelectedItem.Text
        FRM_VER_OK.Label6.Caption = ListView1.SelectedItem.SubItems(3)
        FRM_VER_OK.Show vbModal
    Case 2
        Unload Me
    Case 3
        FRM_VER_REJECT.Label1.Caption = ListView1.SelectedItem.Text
        FRM_VER_REJECT.Show vbModal
End Select
End Sub

Private Sub Form_Load()
Dim LISTITEM As LISTITEM
Dim m_objrs As ADODB.Recordset
MDIForm1.Text5.Visible = True
MDIForm1.ProgressBar1.Visible = True
Call HEADER_PRESCREEN
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open "Select * from RequestInbound", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
MDIForm1.ProgressBar1.Max = m_objrs.RecordCount + 1
Label1.Caption = m_objrs.RecordCount & " Data"
    While Not m_objrs.EOF
    MDIForm1.ProgressBar1.Value = m_objrs.Bookmark
    Set LISTITEM = ListView1.ListItems.ADD(, , IIf(IsNull(m_objrs("CUSTID")), "", m_objrs("CUSTID")))
        LISTITEM.SubItems(1) = IIf(IsNull(m_objrs("StatusRequest")), 0, m_objrs("StatusRequest"))
        LISTITEM.SubItems(2) = IIf(IsNull(m_objrs("AGENT")), "", m_objrs("AGENT"))
        LISTITEM.SubItems(3) = IIf(IsNull(m_objrs("NamaAGENT")), "", m_objrs("NamaAGENT"))
        LISTITEM.SubItems(4) = IIf(IsNull(m_objrs("NAME")), "", m_objrs("NAME"))
        LISTITEM.SubItems(5) = IIf(IsNull(m_objrs("ADDRNOW")), "", m_objrs("ADDRNOW"))
        LISTITEM.SubItems(6) = IIf(IsNull(m_objrs("NAMAPT")), "", m_objrs("NAMAPT"))
        LISTITEM.SubItems(7) = IIf(IsNull(m_objrs("ADDRPT")), "", m_objrs("ADDRPT"))
        LISTITEM.SubItems(8) = IIf(IsNull(m_objrs("HOMENO")), "", m_objrs("HOMENO"))
        LISTITEM.SubItems(9) = IIf(IsNull(m_objrs("HOMENO2")), "", m_objrs("HOMENO2"))
        LISTITEM.SubItems(10) = IIf(IsNull(m_objrs("OFFICENO")), "", m_objrs("OFFICENO"))
        LISTITEM.SubItems(11) = IIf(IsNull(m_objrs("OFFICENO2")), "", m_objrs("OFFICENO2"))
        LISTITEM.SubItems(12) = IIf(IsNull(m_objrs("FAXNO")), "", m_objrs("FAXNO"))
        LISTITEM.SubItems(13) = IIf(IsNull(m_objrs("FAXNO2")), "", m_objrs("FAXNO2"))
        LISTITEM.SubItems(14) = IIf(IsNull(m_objrs("MOBILENO")), "", m_objrs("MOBILENO"))
        LISTITEM.SubItems(15) = IIf(IsNull(m_objrs("MOBILENO2")), "", m_objrs("MOBILENO2"))
        LISTITEM.SubItems(16) = IIf(IsNull(m_objrs("RECSOURCE")), "", m_objrs("RECSOURCE"))
        LISTITEM.SubItems(17) = IIf(IsNull(m_objrs("TGLSOURCE")), "", Format(m_objrs("TGLSOURCE"), "YYYY/MM/DD"))
    m_objrs.MoveNext
    Wend
Set m_objrs = Nothing
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

