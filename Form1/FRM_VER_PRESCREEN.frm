VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRM_VER_PRESCREEN 
   BackColor       =   &H80000004&
   ClientHeight    =   9930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11385
   ForeColor       =   &H00000000&
   Icon            =   "FRM_VER_PRESCREEN.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   662
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   759
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   11070
      Left            =   15
      TabIndex        =   0
      Top             =   30
      Width           =   15195
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
         Height          =   390
         Index           =   0
         Left            =   14400
         TabIndex        =   2
         Top             =   10620
         Width           =   735
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   10470
         Left            =   30
         TabIndex        =   1
         Top             =   120
         Width           =   15120
         _ExtentX        =   26670
         _ExtentY        =   18468
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "0 data"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   13770
         TabIndex        =   3
         Top             =   10725
         Width           =   450
      End
   End
End
Attribute VB_Name = "FRM_VER_PRESCREEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub HEADER_VIEW_SEARCH()
    ListView1.ColumnHeaders.ADD 1, , "No", 5 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Customers Id", 10 * TXT
    ListView1.ColumnHeaders.ADD 3, , "Customers Name", 20 * TXT
    ListView1.ColumnHeaders.ADD 4, , "Alamat Rumah", 15 * TXT
    ListView1.ColumnHeaders.ADD 5, , "Kantor", 15 * TXT
    ListView1.ColumnHeaders.ADD 6, , "Alamat Kantor", 15 * TXT
    ListView1.ColumnHeaders.ADD 7, , "Home Telp", 10 * TXT
    ListView1.ColumnHeaders.ADD 8, , "Home Telp2", 10 * TXT
    ListView1.ColumnHeaders.ADD 9, , "Office Telp", 10 * TXT
    ListView1.ColumnHeaders.ADD 10, , "Office Telp2", 10 * TXT
    ListView1.ColumnHeaders.ADD 11, , "Fax", 10 * TXT
    ListView1.ColumnHeaders.ADD 12, , "Fax2", 10 * TXT
    ListView1.ColumnHeaders.ADD 13, , "Hp", 10 * TXT
    ListView1.ColumnHeaders.ADD 14, , "Hp2", 10 * TXT
    ListView1.ColumnHeaders.ADD 15, , "Agent", 10 * TXT
    ListView1.ColumnHeaders.ADD 16, , "Database", 10 * TXT
    ListView1.ColumnHeaders.ADD 17, , "Tgl Entry", 10 * TXT
    ListView1.ColumnHeaders.ADD 18, , "Tgl Schedule", 10 * TXT
    ListView1.ColumnHeaders.ADD 19, , "Next Action", 17 * TXT
    ListView1.ColumnHeaders.ADD 20, , "LastCall Date", 10 * TXT
    ListView1.ColumnHeaders.ADD 21, , "Sts LastCall", 10 * TXT
End Sub

Private Sub HEADER_PRESCREEN()
    ListView1.ColumnHeaders.ADD 1, , "No", 5 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Customers Id", 10 * TXT
    ListView1.ColumnHeaders.ADD 3, , "Customers Name", 20 * TXT
    ListView1.ColumnHeaders.ADD 4, , "Alamat Rumah", 15 * TXT
    ListView1.ColumnHeaders.ADD 5, , "Kantor", 15 * TXT
    ListView1.ColumnHeaders.ADD 6, , "Alamat Kantor", 15 * TXT
    ListView1.ColumnHeaders.ADD 7, , "Home Telp", 10 * TXT
    ListView1.ColumnHeaders.ADD 8, , "Home Telp2", 10 * TXT
    ListView1.ColumnHeaders.ADD 9, , "Office Telp", 10 * TXT
    ListView1.ColumnHeaders.ADD 10, , "Office Telp2", 10 * TXT
    ListView1.ColumnHeaders.ADD 11, , "Fax", 10 * TXT
    ListView1.ColumnHeaders.ADD 12, , "Fax2", 10 * TXT
    ListView1.ColumnHeaders.ADD 13, , "Hp", 10 * TXT
    ListView1.ColumnHeaders.ADD 14, , "Hp2", 10 * TXT
    ListView1.ColumnHeaders.ADD 15, , "Agent", 10 * TXT
    ListView1.ColumnHeaders.ADD 16, , "Database", 10 * TXT
    ListView1.ColumnHeaders.ADD 17, , "Tgl Entry", 10 * TXT
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
Dim NAMACUST As String
Dim NAMAAGENT As String
Dim DATASOURCE As String
Dim TGLLAHIR As String
Dim OFFPHONE As String
Dim OFFPHONE2 As String
Dim HOMEPHONE As String
Dim HOMEPHONE2 As String
Dim MOBILEPHONE As String
Dim MOBILEPHONE2 As String
Dim FAXPHONE As String
Dim FAXPHONE2 As String
Dim LISTITEM As LISTITEM
Dim m_objrs As ADODB.Recordset
Dim M_DATA As New CLS_FRMSEARCH
Dim TGL_LHR As String

Call HEADER_VIEW_SEARCH

With FRM_VER_SEARCH
    .Height = 4815
    .Frame1.Visible = True
    If Len(.Text1(2).Text) < 3 Then
            If .Text1(0).Text <> Empty Then
                NAMACUST = "NAME LIKE " + "'%" + UBAH_QUOTE(.Text1(0).Text) + "%'"
            End If
            If .Combo1(0).Text <> Empty Then
                NAMAAGENT = "AGENT = '" + .Combo1(0).Text + "'"
            End If
            If .Combo1(2).Text <> Empty Then
                DATASOURCE = "RECSOURCE = '" + .Combo1(2).Text + "'"
            End If
            If .TDBDate1.ValueIsNull Then
            Else
                TGLLAHIR = "BIRTHD = '" + Format(.TDBDate1.Value, "mm/dd/yyyy") + "'"
            End If
            If Len(.TDBMask1.Value) > 1 Then
                OFFPHONE = "OFFICENO LIKE '%" + .TDBMask1.Value + "%'"
                OFFPHONE2 = "OFFICENO2 LIKE '%" + .TDBMask1.Value + "%'"
                HOMEPHONE = "HOMENO LIKE '%" + .TDBMask1.Value + "%'"
                HOMEPHONE2 = "HOMENO2 LIKE '%" + .TDBMask1.Value + "%'"
                FAXPHONE = "FAXNO LIKE '%" + .TDBMask1.Value + "%'"
                FAXPHONE2 = "FAXNO2 LIKE '%" + .TDBMask1.Value + "%'"
            End If
            If Len(.TDBMask2.Value) > 1 Then
                MOBILEPHONE = "MOBILENO LIKE '%" + .TDBMask2.Value + "%'"
                MOBILEPHONE2 = "MOBILENO2 LIKE '%" + .TDBMask2.Value + "%'"
            End If
            Set m_objrs = M_DATA.QUERY_SEARCH_CONDITION(M_OBJCONN, NAMACUST, NAMAAGENT, DATASOURCE, TGLLAHIR, _
                                                    OFFPHONE, OFFPHONE2, HOMEPHONE, HOMEPHONE2, MOBILEPHONE, _
                                                    MOBILEPHONE2, FAXPHONE, FAXPHONE2, MDIForm1.Text3.Text)
    Else
        Set m_objrs = M_DATA.QUERY_SEARCH(M_OBJCONN, "NOLAP = '" + .Text1(2).Text + "'", MDIForm1.Text3.Text)
    End If
End With
FRM_VER_SEARCH.ProgressBar1.Max = m_objrs.RecordCount + 1
Label1.Caption = m_objrs.RecordCount & " Data"
While Not m_objrs.EOF
FRM_VER_SEARCH.ProgressBar1.Value = m_objrs.Bookmark
Set LISTITEM = ListView1.ListItems.ADD(, , m_objrs.Bookmark)
    LISTITEM.SubItems(1) = IIf(IsNull(m_objrs("custid")), "", JADI_QUOTE(m_objrs("custid")))
    LISTITEM.SubItems(2) = IIf(IsNull(m_objrs("NAME")), "", m_objrs("NAME"))
    LISTITEM.SubItems(3) = IIf(IsNull(m_objrs("ADDRNOW")), "", m_objrs("ADDRNOW"))
    LISTITEM.SubItems(4) = IIf(IsNull(m_objrs("NAMAPT")), "", m_objrs("NAMAPT"))
    LISTITEM.SubItems(5) = IIf(IsNull(m_objrs("ADDRPT")), "", m_objrs("ADDRPT"))
    LISTITEM.SubItems(6) = IIf(IsNull(m_objrs("HOMENO")), "", m_objrs("HOMENO"))
    LISTITEM.SubItems(7) = IIf(IsNull(m_objrs("HOMENO2")), "", m_objrs("HOMENO2"))
    LISTITEM.SubItems(8) = IIf(IsNull(m_objrs("OFFICENO")), "", m_objrs("OFFICENO"))
    LISTITEM.SubItems(9) = IIf(IsNull(m_objrs("OFFICENO2")), "", m_objrs("OFFICENO2"))
    LISTITEM.SubItems(10) = IIf(IsNull(m_objrs("FAXNO")), "", m_objrs("FAXNO"))
    LISTITEM.SubItems(11) = IIf(IsNull(m_objrs("FAXNO2")), "", m_objrs("FAXNO2"))
    LISTITEM.SubItems(12) = IIf(IsNull(m_objrs("MOBILENO")), "", m_objrs("MOBILENO"))
    LISTITEM.SubItems(13) = IIf(IsNull(m_objrs("MOBILENO2")), "", m_objrs("MOBILENO2"))
    LISTITEM.SubItems(14) = IIf(IsNull(m_objrs("AGENT")), "", m_objrs("AGENT"))
    LISTITEM.SubItems(15) = IIf(IsNull(m_objrs("RECSOURCE")), "", m_objrs("RECSOURCE"))
    LISTITEM.SubItems(16) = IIf(IsNull(m_objrs("TGLSOURCE")), "", Format(m_objrs("TGLSOURCE"), "YYYY/MM/DD"))
    LISTITEM.SubItems(17) = IIf(IsNull(m_objrs("NEXTACTDATE")), "", Format(m_objrs("NEXTACTDATE"), "yyyy/mm/dd hh:mm"))
    LISTITEM.SubItems(18) = IIf(IsNull(m_objrs("NEXTACT")), "", m_objrs("NEXTACT"))
    LISTITEM.SubItems(19) = IIf(IsNull(m_objrs("TGLstatus")), "", Format(m_objrs("TGLstatus"), "YYYY/MM/DD"))
    LISTITEM.SubItems(20) = IIf(IsNull(m_objrs("StsLastCall")), "", m_objrs("StsLastCall"))
    
m_objrs.MoveNext
If m_objrs.EOF = False Then
    If m_objrs.Bookmark = 2000 Then
        FRM_VER_SEARCH.ProgressBar1.Value = FRM_VER_SEARCH.ProgressBar1.Max
        Set m_objrs = Nothing
        Unload FRM_VER_SEARCH
        Exit Sub
    End If
End If
Wend
Unload FRM_VER_SEARCH
Set m_objrs = Nothing
Set M_DATA = Nothing
End Sub


Private Sub view_search_UPLOAD()
Dim NAMACUST As String
Dim NAMAAGENT As String
Dim DATASOURCE As String
Dim TGLLAHIR As String
Dim OFFPHONE As String
Dim OFFPHONE2 As String
Dim HOMEPHONE As String
Dim HOMEPHONE2 As String
Dim MOBILEPHONE As String
Dim MOBILEPHONE2 As String
Dim FAXPHONE As String
Dim FAXPHONE2 As String
Dim LISTITEM As LISTITEM
Dim m_objrs As ADODB.Recordset
Dim M_DATA As New CLS_FRMSEARCH
Dim TGL_LHR As String

Call HEADER_VIEW_SEARCH

With FRM_VER_SEARCH
    .Height = 4815
    .Frame1.Visible = True
    If Len(.Text1(2).Text) < 3 Then
            If .Text1(0).Text <> Empty Then
                NAMACUST = "NAME LIKE " + "'%" + UBAH_QUOTE(.Text1(0).Text) + "%'"
            End If
            If .Combo1(0).Text <> Empty Then
                NAMAAGENT = "AGENT = '" + .Combo1(0).Text + "'"
            End If
            If .Combo1(2).Text <> Empty Then
                DATASOURCE = "RECSOURCE = '" + .Combo1(2).Text + "'"
            End If
            If .TDBDate1.ValueIsNull Then
            Else
                TGLLAHIR = "BIRTHD = '" + Format(.TDBDate1.Value, "mm/dd/yyyy") + "'"
            End If
            If Len(.TDBMask1.Value) > 4 Then
                OFFPHONE = "OFFICENO = '" + .TDBMask1.Value + "'"
                OFFPHONE2 = "OFFICENO2 = '" + .TDBMask1.Value + "'"
                HOMEPHONE = "HOMENO = '" + .TDBMask1.Value + "'"
                HOMEPHONE2 = "HOMENO2 = '" + .TDBMask1.Value + "'"
                FAXPHONE = "FAXNO = '" + .TDBMask1.Value + "'"
                FAXPHONE2 = "FAXNO2 = '" + .TDBMask1.Value + "'"
            End If
            If Len(.TDBMask2.Value) > 4 Then
                MOBILEPHONE = "MOBILENO = '" + .TDBMask2.Value + "'"
                MOBILEPHONE2 = "MOBILENO2 = '" + .TDBMask2.Value + "'"
            End If
            Set m_objrs = M_DATA.QUERY_SEARCH_UPLOAD(M_OBJCONN, NAMACUST, NAMAAGENT, DATASOURCE, TGLLAHIR, _
                                                    OFFPHONE, OFFPHONE2, HOMEPHONE, HOMEPHONE2, MOBILEPHONE, _
                                                    MOBILEPHONE2, FAXPHONE, FAXPHONE2, MDIForm1.Text3.Text)
    Else
        
    End If
End With
FRM_VER_SEARCH.ProgressBar1.Max = m_objrs.RecordCount + 1
Label1.Caption = m_objrs.RecordCount & " Data"
While Not m_objrs.EOF
FRM_VER_SEARCH.ProgressBar1.Value = m_objrs.Bookmark
Set LISTITEM = ListView1.ListItems.ADD(, , m_objrs.Bookmark)
    LISTITEM.SubItems(1) = IIf(IsNull(m_objrs("custid")), "", JADI_QUOTE(m_objrs("custid")))
    LISTITEM.SubItems(2) = IIf(IsNull(m_objrs("NAME")), "", m_objrs("NAME"))
    LISTITEM.SubItems(3) = IIf(IsNull(m_objrs("ADDRNOW")), "", m_objrs("ADDRNOW"))
    LISTITEM.SubItems(4) = IIf(IsNull(m_objrs("NAMAPT")), "", m_objrs("NAMAPT"))
    LISTITEM.SubItems(5) = IIf(IsNull(m_objrs("ADDRPT")), "", m_objrs("ADDRPT"))
    LISTITEM.SubItems(6) = IIf(IsNull(m_objrs("HOMENO")), "", m_objrs("HOMENO"))
    LISTITEM.SubItems(7) = IIf(IsNull(m_objrs("HOMENO2")), "", m_objrs("HOMENO2"))
    LISTITEM.SubItems(8) = IIf(IsNull(m_objrs("OFFICENO")), "", m_objrs("OFFICENO"))
    LISTITEM.SubItems(9) = IIf(IsNull(m_objrs("OFFICENO2")), "", m_objrs("OFFICENO2"))
    LISTITEM.SubItems(10) = IIf(IsNull(m_objrs("FAXNO")), "", m_objrs("FAXNO"))
    LISTITEM.SubItems(11) = IIf(IsNull(m_objrs("FAXNO2")), "", m_objrs("FAXNO2"))
    LISTITEM.SubItems(12) = IIf(IsNull(m_objrs("MOBILENO")), "", m_objrs("MOBILENO"))
    LISTITEM.SubItems(13) = IIf(IsNull(m_objrs("MOBILENO2")), "", m_objrs("MOBILENO2"))
    LISTITEM.SubItems(14) = IIf(IsNull(m_objrs("AGENT")), "", m_objrs("AGENT"))
    LISTITEM.SubItems(15) = IIf(IsNull(m_objrs("RECSOURCE")), "", m_objrs("RECSOURCE"))
    LISTITEM.SubItems(16) = IIf(IsNull(m_objrs("TGLSOURCE")), "", Format(m_objrs("TGLSOURCE"), "YYYY/MM/DD"))
m_objrs.MoveNext
If m_objrs.EOF = False Then
    If m_objrs.Bookmark = 2000 Then
        FRM_VER_SEARCH.ProgressBar1.Value = FRM_VER_SEARCH.ProgressBar1.Max
        Set m_objrs = Nothing
        Unload FRM_VER_SEARCH
        Exit Sub
    End If
End If
Wend
Unload FRM_VER_SEARCH
Set m_objrs = Nothing
Set M_DATA = Nothing
End Sub


Private Sub Form_Load()
Dim LISTITEM As LISTITEM
Dim TGL_AKTION As String
If search_ok Then
    Call view_search
    Exit Sub
End If
Call HEADER_PRESCREEN
Call view_search_UPLOAD
End Sub

Private Sub ListView1_DblClick()
Dim m_objrs As ADODB.Recordset
If ListView1.ListItems.Count = 0 Then
    Exit Sub
End If
    If UCase(Me.Caption) = "SEARCH" Then
        If UCase(MDIForm1.Text2.Text) = "AGENT" Then
            If UCase(MDIForm1.Text1.Text) <> UCase(ListView1.SelectedItem.SubItems(6)) Then
                MsgBox "Anda Tidak Berhak Untuk Mengedit Data Ini", vbCritical + vbOKOnly, "TeleGrandi"
                Exit Sub
            End If
        End If
        If UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
            Set m_objrs = New ADODB.Recordset
            m_objrs.CursorLocation = adUseClient
            If UCase(Me.Caption) = "SEARCH" Then
                m_objrs.Open "SELECT USERID FROM USERTBL WHERE SPVCODE ='" + MDIForm1.Text1.Text + "' AND USERID = '" + ListView1.SelectedItem.SubItems(6) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            Else
                m_objrs.Open "SELECT USERID FROM USERTBL WHERE SPVCODE ='" + MDIForm1.Text1.Text + "' AND USERID = '" + ListView1.SelectedItem.SubItems(6) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            End If
            If m_objrs.RecordCount <> 0 Then
            Else
                MsgBox "Data Ini Milik TSE Team Leader Yang Lain", vbCritical + vbOKOnly, "TeleGrandi"
                Set m_objrs = Nothing
                Exit Sub
            End If
            Set m_objrs = Nothing
        End If
    Else
        If UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
            Set m_objrs = New ADODB.Recordset
            m_objrs.CursorLocation = adUseClient
            If UCase(Me.Caption) = "SEARCH" Then
                m_objrs.Open "SELECT USERID FROM USERTBL WHERE SPVCODE ='" + MDIForm1.Text1.Text + "' AND USERID = '" + ListView1.SelectedItem.SubItems(6) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            Else
                m_objrs.Open "SELECT USERID FROM USERTBL WHERE SPVCODE ='" + MDIForm1.Text1.Text + "' AND USERID = '" + ListView1.SelectedItem.SubItems(6) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            End If
            If m_objrs.RecordCount <> 0 Then
            Else
                MsgBox "Data Ini Milik TSE Team Leader Yang Lain", vbCritical + vbOKOnly, "TeleGrandi"
                Set m_objrs = Nothing
                Exit Sub
            End If
            Set m_objrs = Nothing
        End If
    End If
    ADD_CUST = False
    VIEW_OK = False
    SCREENER_AWAL = False
    SCREENER = True
    VIEW_AVAIL_AWAL = False
    SCREENER_APPROV = False
    REBUT_DATA = False
    SCR_SPV_CARI = False
    Select Case UCase(MDIForm1.Text3.Text)
    Case "CREDIT CARD"
    If StsMgmSchedule = False Then
        FRMCUST_CC.Show vbModal
    Else
        FRMCUST_CC_MGM.Show vbModal
    End If
    Case Else
        MsgBox "Anda Tidak Mempunyai Hak Untuk Mengedit Data", vbInformation + vbOKOnly, "Informasi"
        Exit Sub
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
    Call ListView1_DblClick
End If
End Sub

