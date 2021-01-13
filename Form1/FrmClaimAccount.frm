VERSION 5.00
Begin VB.Form FrmClaimAccount 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Claim Account"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4335
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   3180
      TabIndex        =   7
      Top             =   2400
      Width           =   1035
   End
   Begin VB.CommandButton CmdProses 
      Caption         =   "Proses"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   2400
      Width           =   1035
   End
   Begin VB.TextBox TxtAlasanClaim 
      Height          =   1605
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   3015
   End
   Begin VB.TextBox TxtNama 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   420
      Width           =   3015
   End
   Begin VB.TextBox Txtcustid 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   60
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Alasan claim:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Nama:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   420
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Custid:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "FrmClaimAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBatal_Click()
    Unload Me
End Sub

Private Sub cmdproses_Click()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim a, pesan, RemarksClaim As String
    
    If TxtAlasanClaim.Text = "" Then
        MsgBox "Alasan claim tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    a = MsgBox("Apakah anda yakin akan memproses claim account?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbNo Then
        MsgBox "Proses dibatalkan!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    On Error GoTo salah
    TxtAlasanClaim.Enabled = False
    DoEvents
    
    
    'Update ke mgm pindahkan ke account claim
    CMDSQL = "update mgm set agent='CLAIM',user_claim='"
    CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "', waktu_claim=now(),alasan_claim='"
    CMDSQL = CMDSQL + Replace(TxtAlasanClaim.Text, "'", "") + "' "
    CMDSQL = CMDSQL + " where custid='"
    CMDSQL = CMDSQL + CStr(TxtCustid.Text) + "'"
    M_OBJCONN.Execute CMDSQL
    
    'Catet di history nih...
    RemarksClaim = "Agent : " & MDIForm1.TxtUsername.Text & "=> Telah Melakukan Claim pada account ini <="
    RemarksClaim = RemarksClaim & Replace(TxtAlasanClaim.Text, "'", "")
    
    CMDSQL = "insert into mgm_hst (custid,agent,hst,tgl,user_log) values ('"
    CMDSQL = CMDSQL + CStr(TxtCustid.Text) + "','"
    CMDSQL = CMDSQL & FrmCC_Colection.lblaoc.Caption + "','"
    CMDSQL = CMDSQL & RemarksClaim & "',now(),'"
    CMDSQL = CMDSQL & MDIForm1.TxtUsername.Text & "')"
    M_OBJCONN.Execute CMDSQL
    
    ' UPDATED 22 MEI 2013 - IZUDDIN
    CMDSQL = "insert into tbllog_claim_aksesall (custid,agent,agentlama,tgl_claim) values ('"
    CMDSQL = CMDSQL + CStr(TxtCustid.Text) + "','"
    CMDSQL = CMDSQL & MDIForm1.TxtUsername.Text + "','"
    CMDSQL = CMDSQL & FrmCC_Colection.lbl_agentlama.Caption + "', "
    CMDSQL = CMDSQL & "now())"
    M_OBJCONN.Execute CMDSQL
    
    'Kirim pesan ke semua agent yang ada di distribusi
    pesan = "Pesan ini dibuat otomatis oleh system " & vbCrLf
    pesan = pesan & "========================================" & vbCrLf
    pesan = pesan & "Agent : " & MDIForm1.TxtUsername.Text & vbCrLf
    pesan = pesan & "Telah melakukan Claim untuk account : " & vbCrLf & vbCrLf
    pesan = pesan & "Custid :" & TxtCustid.Text & vbCrLf
    pesan = pesan & "Nama :" & txtNama.Text & vbCrLf & vbCrLf
    pesan = pesan & "Alasan untuk claim: " & vbCrLf
    pesan = pesan & Replace(TxtAlasanClaim.Text, "'", "")
    
    '--1. Kirim ke TL dan SPV
    'Cmdsql = "select * from usertbl where usertype in ('11','6','20','25') "
    '@@20022013 Jika agent mengclaim account, pesan ga usah ditampilkan ke spv
    CMDSQL = "select * from usertbl where usertype in ('6','20','25') "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            CMDSQL = "insert into msgtbl "
            CMDSQL = CMDSQL + "( recipient, datetime, sender, sentfrom, msg) values ('"
            CMDSQL = CMDSQL + M_objrs("userid") + "','"
            CMDSQL = CMDSQL + Format(Now(), "yyyymmdd") + "','"
            CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "','"
            CMDSQL = CMDSQL + CStr(MDIForm1.Winsock1.LocalIP) + "','"
            CMDSQL = CMDSQL + pesan + "')"
            M_OBJCONN.Execute CMDSQL
            M_objrs.MoveNext
        Wend
    End If
    Set M_objrs = Nothing
    
    '--2. Kirim juga agent yang ada di tabel distribusi yang ikut mengcollect custid ini
'    Cmdsql = "select agent from tbl_distribusi_account where custid='"
'    Cmdsql = Cmdsql & CStr(TxtCustid.Text) & "'"
    CMDSQL = "SELECT a.*,b.userid as agent FROM tbl_cust_aksesall a,usertbl b WHERE a.kd_profile=b.profile_akses_all "
    CMDSQL = CMDSQL & " AND a.custid='" & CStr(TxtCustid.Text) & "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            CMDSQL = "insert into msgtbl "
            CMDSQL = CMDSQL + "( recipient, datetime, sender, sentfrom, msg) values ('"
            CMDSQL = CMDSQL + M_objrs("agent") + "','"
            CMDSQL = CMDSQL + Format(Now(), "yyyymmdd") + "','"
            CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "','"
            CMDSQL = CMDSQL + CStr(MDIForm1.Winsock1.LocalIP) + "','"
            CMDSQL = CMDSQL + pesan + "')"
            M_OBJCONN.Execute CMDSQL
            M_objrs.MoveNext
        Wend
    End If
    Set M_objrs = Nothing
    
    'Hapus data dari tbl_distribusi_account
'    Cmdsql = "delete from tbl_distribusi_account where custid='"
    CMDSQL = "delete from tbl_cust_aksesall where custid='"
    CMDSQL = CMDSQL & CStr(TxtCustid.Text) & "'"
    M_OBJCONN.Execute CMDSQL
    
    MsgBox "Proses claim sudah dikirim! Jika di ACC sistem akan memberitahukan kepada anda!", vbOKOnly + vbInformation, "Informasi"
    Unload Me
    Exit Sub
salah:
    MsgBox "Mohon maaf ada kesalahan: " & err.Description
    
End Sub
