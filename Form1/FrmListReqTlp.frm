VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmListReqTlp 
   Caption         =   "List Request Number Telephone"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10980
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   10980
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Request Number"
      TabPicture(0)   =   "FrmListReqTlp.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "CmdCekAll"
      Tab(0).Control(1)=   "CmdReject"
      Tab(0).Control(2)=   "TxtReq"
      Tab(0).Control(3)=   "CmdApprove"
      Tab(0).Control(4)=   "LstReq"
      Tab(0).Control(5)=   "Label1"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Log Request Number"
      TabPicture(1)   =   "FrmListReqTlp.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "LstReqLog"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "TxtReqLog"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.CheckBox CmdCekAll 
         Caption         =   "Check All"
         Height          =   375
         Left            =   -74760
         TabIndex        =   9
         Top             =   5400
         Width           =   1215
      End
      Begin VB.CommandButton CmdReject 
         Caption         =   "&Reject"
         Height          =   555
         Left            =   -70020
         TabIndex        =   8
         Top             =   5340
         Width           =   1575
      End
      Begin VB.TextBox TxtReqLog 
         Enabled         =   0   'False
         Height          =   285
         Left            =   9840
         TabIndex        =   7
         Text            =   "0"
         Top             =   5520
         Width           =   975
      End
      Begin VB.TextBox TxtReq 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -65160
         TabIndex        =   5
         Text            =   "0"
         Top             =   5400
         Width           =   975
      End
      Begin VB.CommandButton CmdApprove 
         Caption         =   "&Approve"
         Height          =   555
         Left            =   -71580
         TabIndex        =   3
         Top             =   5340
         Width           =   1575
      End
      Begin MSComctlLib.ListView LstReq 
         Height          =   4860
         Left            =   -74820
         TabIndex        =   1
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
      Begin MSComctlLib.ListView LstReqLog 
         Height          =   4860
         Left            =   180
         TabIndex        =   2
         Top             =   480
         Width           =   10680
         _ExtentX        =   18838
         _ExtentY        =   8573
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
      Begin VB.Label Label2 
         Caption         =   "Jumlah data:"
         Height          =   255
         Left            =   8820
         TabIndex        =   6
         Top             =   5520
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Jumlah data:"
         Height          =   255
         Left            =   -66180
         TabIndex        =   4
         Top             =   5460
         Width           =   975
      End
   End
End
Attribute VB_Name = "FrmListReqTlp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub HeaderReq()
    LstReq.ColumnHeaders.ADD 1, , "ID", 800
    LstReq.ColumnHeaders.ADD 2, , "Tgl.Request", 0
    LstReq.ColumnHeaders.ADD 3, , "Custid", 2000
    LstReq.ColumnHeaders.ADD 4, , "Agent", 1500
    LstReq.ColumnHeaders.ADD 5, , "Home 1", 0
    LstReq.ColumnHeaders.ADD 6, , "Office 1", 0
    LstReq.ColumnHeaders.ADD 7, , "Mobile 1", 0
    
    '@@17042012, Perubahan u/ request number hanya ada nomor dan kategori
    LstReq.ColumnHeaders.ADD 8, , "Request Number", 1500
    LstReq.ColumnHeaders.ADD 9, , "Jenis", 3000
End Sub

Private Sub HeaderLstReq()
    LstReqLog.ColumnHeaders.ADD 1, , "Tgl.Request", 1500
    LstReqLog.ColumnHeaders.ADD 2, , "Custid", 1500
    LstReqLog.ColumnHeaders.ADD 3, , "Agent", 1500
    LstReqLog.ColumnHeaders.ADD 4, , "Home 1", 1500
    LstReqLog.ColumnHeaders.ADD 5, , "Office 1", 1500
    LstReqLog.ColumnHeaders.ADD 6, , "Mobile 1", 1500
    LstReqLog.ColumnHeaders.ADD 7, , "Tgl.Approve", 1500
    LstReqLog.ColumnHeaders.ADD 8, , "Approve By", 1500
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
    
    If LstReq.ListItems.Count = 0 Then
        MsgBox "Data request tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    K = MsgBox("Anda yakin akan melakukan approve number?", vbYesNo + vbQuestion, "Konfirmasi")
    
    If K = vbNo Then
        Exit Sub
    End If
    
    CmdApprove.Enabled = False
    CmdCekAll.Enabled = False
    
    DoEvents
    bAdd_phone = False
    For W = 1 To LstReq.ListItems.Count
        If LstReq.ListItems(W).Checked = True Then
                          
                'If bAdd_phone Then
                pesan = "Request Number di Approve: " & vbCrLf
                pesan = pesan & " Custid : " & LstReq.ListItems(W).SubItems(2) & vbCrLf
                pesan = pesan & " Di approve oleh : " & MDIForm1.TxtUsername.text
                
                'Update di MGM
                If LstReq.ListItems(W).SubItems(8) = "AddHome1" Then
                    CMDSQL = "update mgm set homenoadd1='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(7)) + "' where custid='"
                ElseIf LstReq.ListItems(W).SubItems(8) = "AddOffice1" Then
                    CMDSQL = "update mgm set officenoadd1='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(7)) + "' where custid='"
                ElseIf LstReq.ListItems(W).SubItems(8) = "AddMobile1" Then
                    CMDSQL = "update mgm set mobilenoadd1='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(7)) + "' where custid='"
                ElseIf LstReq.ListItems(W).SubItems(8) = "AddOtherphone" Then
                    CMDSQL = "update mgm set mobilenoadd2='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(7)) + "' where custid='"
                Else
                    CMDSQL = "update mgm set req_nomor_telp='"
                    CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(7)) + "' where custid='"
                End If
                
                CMDSQL = CMDSQL + CStr(LstReq.ListItems(W).SubItems(2)) + "'"
                M_OBJCONN.Execute CMDSQL
                
                'Update Data Ke tabel LOg Telepon
                CMDSQL = "INSERT INTO tblrequestadditionalphone_log "
                CMDSQL = CMDSQL + "select * from tblrequestadditionalphone where id='"
                CMDSQL = CMDSQL + LstReq.ListItems(W).text + "'"
                M_OBJCONN.Execute CMDSQL
                
                'Update data log, tgl approve dan di approve oleh
                CMDSQL = "UPDATE tblrequestadditionalphone_log set tglapprove='"
                CMDSQL = CMDSQL + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd hh:mm:ss") + "', approve_by='"
                CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "' where id='"
                CMDSQL = CMDSQL + LstReq.ListItems(W).text + "'"
                M_OBJCONN.Execute CMDSQL
                
                'Hapus data di tabel tblrequestadditionalphone
                CMDSQL = "DELETE FROM tblrequestadditionalphone where id='"
                CMDSQL = CMDSQL + LstReq.ListItems(W).text + "'"
                M_OBJCONN.Execute CMDSQL
                
                'Kasih tau agent
                CMDSQL = "INSERT INTO msgtbl (recipient,datetime,sender,msg) values ('"
                CMDSQL = CMDSQL + LstReq.ListItems(W).SubItems(3) + "','"
                CMDSQL = CMDSQL + CStr(Format(MDIForm1.TDBDate1.Value, "yyyymmdd")) + "','"
                CMDSQL = CMDSQL + CStr(MDIForm1.TxtUsername.text) + "','"
                CMDSQL = CMDSQL + pesan + "')"
                M_OBJCONN.Execute CMDSQL
                
                ' Masuk History 21 Juli 2014
                strket_hst = "Approve phone number : " & CStr(LstReq.ListItems(W).SubItems(7)) & ""
                'M_DATA.ADD_HISTORY LstReq.ListItems(w).SubItems(2), MDIForm1.TDBDate1.Text, Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(CDate(waktu), "hh:nn:ss"), LstReq.ListItems(w).SubItems(3), "COLLECTION", strket_hst, "", "", "", "", "", "", "", "", "", MDIForm1.Text1.Text, "", "0", ""

        End If
    Next W
    
    MsgBox "Nomor Telepon berhasil di approve!", vbOKOnly + vbInformation, "Informasi"
    
    LstReq.ListItems.CLEAR
    Call Isi_Req

    LstReqLog.ListItems.CLEAR
    Call Isi_Req_log
    
    CmdApprove.Enabled = True
    CmdCekAll.Enabled = True
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

Private Sub CmdReject_Click()
    Dim W As Integer
    Dim CMDSQL As String
    Dim pesan As String
    
    If LstReq.ListItems.Count = 0 Then
        MsgBox "Data request tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    pesan = MsgBox("Yakin data mau dihapus?", vbYesNo + vbQuestion, "Konfirmasi")
    
    If pesan = vbYes Then
        For W = 1 To LstReq.ListItems.Count
            If LstReq.ListItems(W).Checked = True Then
             'Update Data Ke tabel LOg Telepon
                CMDSQL = "INSERT INTO tblrequestadditionalphone_log "
                CMDSQL = CMDSQL + "select * from tblrequestadditionalphone where id='"
                CMDSQL = CMDSQL + LstReq.ListItems(W).text + "'"
                M_OBJCONN.Execute CMDSQL
                
                'Update data log, tgl approve dan di approve oleh
                CMDSQL = "UPDATE tblrequestadditionalphone_log set tglreject='"
                CMDSQL = CMDSQL + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd hh:mm:ss") + "', reject_by='"
                CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "' where id='"
                CMDSQL = CMDSQL + LstReq.ListItems(W).text + "'"
                M_OBJCONN.Execute CMDSQL
                
               CMDSQL = "delete from tblrequestadditionalphone where id='"
               CMDSQL = CMDSQL + LstReq.ListItems(W).text + "'"
               M_OBJCONN.Execute CMDSQL
            End If
        Next W
        
        MsgBox "Data berhasil dihapus!", vbOKOnly + vbInformation, "Informasi"
        LstReq.ListItems.CLEAR
        Call Isi_Req
    End If
End Sub

Private Sub Form_Load()
    Call HeaderLstReq
    Call Isi_Req_log
    
    Call HeaderReq
    Call Isi_Req
End Sub

Private Sub Isi_Req()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim ListItem As ListItem
    
    If UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Then
        CMDSQL = "select * from tblrequestadditionalphone where agent in "
        CMDSQL = CMDSQL + " (select userid from usertbl where spvcode='"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "') "
        CMDSQL = CMDSQL + " order by tglreq desc "
    Else
        CMDSQL = "select * from tblrequestadditionalphone order by tglreq desc"
    End If
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    While Not M_objrs.EOF
        Set ListItem = LstReq.ListItems.ADD(, , M_objrs("id"))
            ListItem.SubItems(1) = IIf(IsNull(M_objrs("tglreq")), "", M_objrs("tglreq"))
            ListItem.SubItems(2) = IIf(IsNull(M_objrs("custid")), "", M_objrs("custid"))
            ListItem.SubItems(3) = IIf(IsNull(M_objrs("agent")), "", M_objrs("agent"))
            ListItem.SubItems(4) = IIf(IsNull(M_objrs("home1")), "", M_objrs("home1"))
            ListItem.SubItems(5) = IIf(IsNull(M_objrs("office1")), "", M_objrs("office1"))
            ListItem.SubItems(6) = IIf(IsNull(M_objrs("mobile1")), "", M_objrs("mobile1"))
            ListItem.SubItems(7) = IIf(IsNull(M_objrs("request_number")), "", M_objrs("request_number"))
            ListItem.SubItems(8) = IIf(IsNull(M_objrs("jenis")), "", M_objrs("jenis"))
        M_objrs.MoveNext
    Wend
    
    Set M_objrs = Nothing
    
    TxtReq.text = LstReq.ListItems.Count
End Sub


Private Sub Isi_Req_log()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim ListItem As ListItem
    
    If UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Then
        CMDSQL = "select * from tblrequestadditionalphone_log where agent in "
        CMDSQL = CMDSQL + " (select userid from usertbl where spvcode='"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "') "
        CMDSQL = CMDSQL + " order by tglreq desc limit 100"
    Else
        CMDSQL = "select * from tblrequestadditionalphone_log order by tglreq desc limit 100"
    End If
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    
    
    While Not M_objrs.EOF
        Set ListItem = LstReqLog.ListItems.ADD(, , M_objrs("tglreq"))
            ListItem.SubItems(1) = IIf(IsNull(M_objrs("custid")), "", M_objrs("custid"))
            ListItem.SubItems(2) = IIf(IsNull(M_objrs("agent")), "", M_objrs("agent"))
            ListItem.SubItems(3) = IIf(IsNull(M_objrs("home1")), "", M_objrs("home1"))
            ListItem.SubItems(4) = IIf(IsNull(M_objrs("office1")), "", M_objrs("office1"))
            ListItem.SubItems(5) = IIf(IsNull(M_objrs("mobile1")), "", M_objrs("mobile1"))
            ListItem.SubItems(6) = IIf(IsNull(M_objrs("tglapprove")), "", M_objrs("tglapprove"))
            ListItem.SubItems(7) = IIf(IsNull(M_objrs("approve_by")), "", M_objrs("approve_by"))
        M_objrs.MoveNext
    Wend
    
    TxtReqLog.text = LstReqLog.ListItems.Count
End Sub



