VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Begin VB.Form Form_addphone_reject 
   Caption         =   "List Add Phone Reject"
   ClientHeight    =   6255
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11235
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   11235
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Request Number"
      TabPicture(0)   =   "Form_addphone_reject.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LstReq"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "TxtReq"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CmdUnCekAll"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CmdCekAll"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CmdApprove"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.CommandButton Command1 
         Caption         =   "&Reject"
         Height          =   555
         Left            =   5040
         TabIndex        =   7
         Top             =   5340
         Width           =   1575
      End
      Begin VB.CommandButton CmdApprove 
         Caption         =   "&Approve"
         Height          =   555
         Left            =   3420
         TabIndex        =   4
         Top             =   5340
         Width           =   1575
      End
      Begin VB.CommandButton CmdCekAll 
         Caption         =   "&Cek All"
         Height          =   555
         Left            =   180
         TabIndex        =   3
         Top             =   5340
         Width           =   1575
      End
      Begin VB.CommandButton CmdUnCekAll 
         Caption         =   "&UnCek All"
         Height          =   555
         Left            =   1740
         TabIndex        =   2
         Top             =   5340
         Width           =   1575
      End
      Begin VB.TextBox TxtReq 
         Enabled         =   0   'False
         Height          =   285
         Left            =   9840
         TabIndex        =   1
         Text            =   "0"
         Top             =   5400
         Width           =   975
      End
      Begin MSComctlLib.ListView LstReq 
         Height          =   4860
         Left            =   180
         TabIndex        =   5
         Top             =   480
         Width           =   10680
         _ExtentX        =   18838
         _ExtentY        =   8573
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label1 
         Caption         =   "Jumlah data:"
         Height          =   255
         Left            =   8820
         TabIndex        =   6
         Top             =   5460
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form_addphone_reject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub HeaderReq()
    LstReq.ColumnHeaders.ADD 1, , "ID", 800
    LstReq.ColumnHeaders.ADD 2, , "Tgl.Request", 2000
    LstReq.ColumnHeaders.ADD 3, , "Custid", 2000
    LstReq.ColumnHeaders.ADD 4, , "Agent", 1500
    LstReq.ColumnHeaders.ADD 5, , "Home 1", 0
    LstReq.ColumnHeaders.ADD 6, , "Home 2", 0
    LstReq.ColumnHeaders.ADD 7, , "Office 1", 0
    LstReq.ColumnHeaders.ADD 8, , "Office 2", 0
    LstReq.ColumnHeaders.ADD 9, , "Mobile 1", 0
    LstReq.ColumnHeaders.ADD 10, , "Mobile 2", 0
    LstReq.ColumnHeaders.ADD 11, , "EcPhone", 0
    
    '@@17042012, Perubahan u/ request number hanya ada nomor dan kategori
    LstReq.ColumnHeaders.ADD 12, , "Request Number", 1500
    LstReq.ColumnHeaders.ADD 13, , "Kategori", 3000
    
    LstReq.ColumnHeaders.ADD 14, , "Keterangan", 3000
    LstReq.ColumnHeaders.ADD 15, , "Jenis", 3000
End Sub

Private Sub CmdApprove_Click()
    Dim m_data      As New CLS_FRMCUST_CC_MGM
    Dim CMDSQL      As String
    Dim M_objrs     As ADODB.Recordset
    Dim W           As Integer
    Dim ListItem    As ListItem
    Dim pesan       As String
    Dim K           As String
    Dim strket_hst  As String
    Dim bAdd_phone  As Boolean
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.ActiveConnection = M_OBJCONN
    M_objrs.CursorLocation = adUseClient
    M_objrs.CursorType = adOpenDynamic
    M_objrs.LockType = adLockOptimistic
    
    If LstReq.ListItems.Count = 0 Then
        MsgBox "Data request tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    K = MsgBox("Anda yakin akan melakukan approve number?", vbYesNo + vbQuestion, "Konfirmasi")
    
    If K = vbNo Then
        Exit Sub
    End If
    
    
    On Error GoTo pesan_error
    
    DoEvents
    bAdd_phone = False
    For W = 1 To LstReq.ListItems.Count
        If LstReq.ListItems(W).Checked = True Then
                M_OBJCONN.Execute "DELETE FROM mandiri.tblrequestadditionalphone_reject WHERE id='" + LstReq.ListItems(W).Text + "'"
                M_OBJCONN.Execute "DELETE FROM mandiri.tbllist_addphone WHERE phone='" & CStr(Trim(LstReq.ListItems(W).SubItems(11))) & "'"
                M_OBJCONN.Execute "INSERT INTO mandiri.tblapprove_reject_addphone(phone_number,approve_by) VALUES('" & CStr(Trim(LstReq.ListItems(W).SubItems(11))) & "','" & MDIForm1.TxtUsername.Text & "')"

                M_OBJCONN.Execute ("INSERT INTO mandiri.tbllist_addphone (custid,phone) VALUES ('" & LstReq.ListItems(W).SubItems(2) & "','" & CStr(Trim(LstReq.ListItems(W).SubItems(11))) & "')")
                
                'If bAdd_phone Then
                pesan = "Request Number di Approve: " & vbCrLf
                pesan = pesan & " Custid : " & LstReq.ListItems(W).SubItems(2) & vbCrLf
                pesan = pesan & " Di approve oleh : " & MDIForm1.TxtUsername.Text
                
                'Update di MGM
                If LstReq.ListItems(W).SubItems(14) = "AddHome1" Then
                    CMDSQL = "update mandiri.mgm set homenoadd1='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(11)) + "',stskathomeadd1='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(12)) + "' where custid='"
                ElseIf LstReq.ListItems(W).SubItems(14) = "AddHome2" Then
                    CMDSQL = "update mandiri.mgm set homenoadd2='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(11)) + "',stskathomeadd2='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(12)) + "' where custid='"
                ElseIf LstReq.ListItems(W).SubItems(14) = "AddOffice1" Then
                    CMDSQL = "update mandiri.mgm set officenoadd1='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(11)) + "',stskatofficeadd1='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(12)) + "' where custid='"
                 ElseIf LstReq.ListItems(W).SubItems(14) = "AddOffice2" Then
                    CMDSQL = "update mandiri.mgm set officenoadd2='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(11)) + "',stskatofficeadd2='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(12)) + "' where custid='"
                 ElseIf LstReq.ListItems(W).SubItems(14) = "AddMobile1" Then
                    CMDSQL = "update mandiri.mgm set mobilenoadd1='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(11)) + "',stskathpadd1='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(12)) + "' where custid='"
                ElseIf LstReq.ListItems(W).SubItems(14) = "AddMobile2" Then
                    CMDSQL = "update mandiri.mgm set mobilenoadd2='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(11)) + "',stskathpadd2='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(12)) + "' where custid='"
                ElseIf LstReq.ListItems(W).SubItems(14) = "AddOther" Then
                    CMDSQL = "update mandiri.mgm set req_nomor_telp='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(11)) + "',status_telp='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(12)) + "' where custid='"
                Else
                    CMDSQL = "update mandiri.mgm set req_nomor_telp='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(11)) + "',status_telp='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(12)) + "' where custid='"
                End If
                
                CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(2)) + "'"
                M_OBJCONN.Execute CMDSQL
                
                'Update Data Ke tabel LOg Telepon
                CMDSQL = "insert into mandiri.tblrequestadditionalphone_log "
                CMDSQL = CMDSQL + "select * from mandiri.tblrequestadditionalphone where id='"
                CMDSQL = CMDSQL + LstReq.ListItems(W).Text + "'"
                M_OBJCONN.Execute CMDSQL
                
                'Update data log, tgl approve dan di approve oleh
                CMDSQL = "update mandiri.tblrequestadditionalphone_log set tglapprove='"
                CMDSQL = CMDSQL + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd hh:mm:ss") + "', approve_by='"
                CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "' where id='"
                CMDSQL = CMDSQL + LstReq.ListItems(W).Text + "'"
                M_OBJCONN.Execute CMDSQL
                
                'Hapus data di tabel tblrequestadditionalphone
                CMDSQL = "delete from mandiri.tblrequestadditionalphone where id='"
                CMDSQL = CMDSQL + LstReq.ListItems(W).Text + "'"
                M_OBJCONN.Execute CMDSQL
                
                '@@12-07-2011 Update status tanda request number
                CMDSQL = "update mandiri.usertbl set f_req_number=null where userid in ("
                CMDSQL = CMDSQL + "select team from mandiri.usertbl where userid='"
                CMDSQL = CMDSQL + LstReq.ListItems(W).SubItems(3) + "') "
                CMDSQL = CMDSQL + " or userid in (select userid  from mandiri.usertbl where "
                CMDSQL = CMDSQL + " usertype='20' or usertype='25' or usertype='11') "
                M_OBJCONN.Execute CMDSQL
                
                'Kasih tau agent
                CMDSQL = "insert into mandiri.msgtbl (recipient,datetime,sender,msg) values ('"
                CMDSQL = CMDSQL + LstReq.ListItems(W).SubItems(3) + "','"
                CMDSQL = CMDSQL + CStr(Format(MDIForm1.TDBDate1.Value, "yyyymmdd")) + "','"
                CMDSQL = CMDSQL + CStr(MDIForm1.TxtUsername.Text) + "','"
                CMDSQL = CMDSQL + pesan + "')"
                M_OBJCONN.Execute CMDSQL
                
                ' Masuk History 21 Juli 2014
                strket_hst = "Approve phone number : " & CStr(LstReq.ListItems(W).SubItems(11)) & "::" & CStr(LstReq.ListItems(W).SubItems(12))
                m_data.ADD_HISTORY LstReq.ListItems(W).SubItems(2), MDIForm1.TDBDate1.Text, Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(CDate(waktu), "hh:nn:ss"), LstReq.ListItems(W).SubItems(3), "COLLECTION", strket_hst, "", "", "", "", "", "", "", "", "", MDIForm1.TxtUsername.Text, "", "0", ""
        End If
    Next W
    
    MsgBox "Nomor Telepon berhasil di approve!", vbOKOnly + vbInformation, "Informasi"
    
    Call Isi_Req
    Exit Sub
pesan_error:
    MsgBox err.Description
End Sub

Private Sub CmdCekAll_Click()
    Dim K As Integer
    
    If LstReq.ListItems.Count = 0 Then
        MsgBox "Data request tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For K = 1 To LstReq.ListItems.Count
        LstReq.ListItems(K).Checked = True
    Next K
End Sub

Private Sub CmdUnCekAll_Click()
    Dim K As Integer
    
    If LstReq.ListItems.Count = 0 Then
        MsgBox "Data request tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For K = 1 To LstReq.ListItems.Count
        LstReq.ListItems(K).Checked = False
    Next K
End Sub

Private Sub Command1_Click()
    Dim pesan As String
    Dim W As Integer
    
    If LstReq.ListItems.Count > 0 Then
        pesan = MsgBox("No telp duplikat yang dipilih akan diapprove??", vbYesNo + vbQuestion, "Confirm")
        
        If pesan = vbYes Then
            For W = 1 To LstReq.ListItems.Count
                If LstReq.ListItems(W).Checked = True Then
                    M_OBJCONN.Execute "DELETE FROM mandiri.tblrequestadditionalphone_reject WHERE id='" + LstReq.ListItems(W).Text + "'"
                    M_OBJCONN.Execute "INSERT INTO mandiri.tblapprove_reject_addphone(phone_number,approve_by) VALUES('" & CStr(Trim(LstReq.ListItems(W).SubItems(11))) & "','" & MDIForm1.TxtUsername.Text & "')"
                End If
            Next W
        End If
        
        MsgBox "Proses Reject berhasil", vbOKOnly, "Info"
        Call Isi_Req
    End If
End Sub

Private Sub Form_Load()
    Call HeaderReq
    Call Isi_Req
End Sub

Private Sub Isi_Req()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim ListItem As ListItem
    
    CMDSQL = "SELECT * FROM mandiri.tblrequestadditionalphone_reject ORDER BY tglreq DESC"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LstReq.ListItems.CLEAR
    
    While Not M_objrs.EOF
        Set ListItem = LstReq.ListItems.ADD(, , M_objrs("id"))
        ListItem.SubItems(1) = IIf(IsNull(M_objrs("tglreq")), "", M_objrs("tglreq"))
        ListItem.SubItems(2) = IIf(IsNull(M_objrs("custid")), "", M_objrs("custid"))
        ListItem.SubItems(3) = IIf(IsNull(M_objrs("agent")), "", M_objrs("agent"))
        ListItem.SubItems(4) = IIf(IsNull(M_objrs("home1")), "", M_objrs("home1"))
        ListItem.SubItems(5) = IIf(IsNull(M_objrs("home2")), "", M_objrs("home2"))
        ListItem.SubItems(6) = IIf(IsNull(M_objrs("office1")), "", M_objrs("office1"))
        ListItem.SubItems(7) = IIf(IsNull(M_objrs("office2")), "", M_objrs("office2"))
        ListItem.SubItems(8) = IIf(IsNull(M_objrs("mobile1")), "", M_objrs("mobile1"))
        ListItem.SubItems(9) = IIf(IsNull(M_objrs("mobile2")), "", M_objrs("mobile2"))
        ListItem.SubItems(10) = IIf(IsNull(M_objrs("ecphone")), "", M_objrs("ecphone"))
        ListItem.SubItems(11) = IIf(IsNull(M_objrs("request_number")), "", M_objrs("request_number"))
        ListItem.SubItems(12) = IIf(IsNull(M_objrs("kategori")), "", M_objrs("kategori"))
        ListItem.SubItems(13) = IIf(IsNull(M_objrs("keterangan")), "", M_objrs("keterangan"))
        ListItem.SubItems(14) = IIf(IsNull(M_objrs("jenis")), "", M_objrs("jenis"))
        M_objrs.MoveNext
    Wend
    
    M_objrs.Close
    Set M_objrs = Nothing
    
    TxtReq.Text = LstReq.ListItems.Count
End Sub

Private Sub LstReq_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim IndexColumnHEader As Integer
    
    LstReq.SortKey = ColumnHeader.Index - 1
    IndexColumnHEader = ColumnHeader.Index - 1
    LstReq.Sorted = True
End Sub
