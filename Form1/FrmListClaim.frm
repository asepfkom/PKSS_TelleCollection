VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmListClaim 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "List Claim Account"
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12810
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   12810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox cek_agent_lama 
      Caption         =   "Cek Agent Lama"
      Height          =   195
      Left            =   8520
      TabIndex        =   29
      Top             =   130
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Detail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   60
      TabIndex        =   17
      Top             =   7140
      Width           =   12675
      Begin MSComDlg.CommonDialog CD 
         Left            =   7560
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Export To Excel"
         Enabled         =   0   'False
         Height          =   435
         Left            =   11040
         TabIndex        =   28
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label LblAlasanClaim 
         Caption         =   "<alasan claim none>"
         Height          =   735
         Left            =   1440
         TabIndex        =   27
         Top             =   1320
         Width           =   11055
      End
      Begin VB.Label Label9 
         Caption         =   "Alasan Claim:"
         Height          =   195
         Left            =   300
         TabIndex        =   26
         Top             =   1320
         Width           =   1035
      End
      Begin VB.Label LblWaktuClaim 
         Caption         =   "<waktu claim none>"
         Height          =   195
         Left            =   4140
         TabIndex        =   25
         Top             =   1020
         Width           =   1635
      End
      Begin VB.Label Label8 
         Caption         =   "Waktu Claim:"
         Height          =   195
         Left            =   3120
         TabIndex        =   24
         Top             =   1020
         Width           =   1035
      End
      Begin VB.Label LblAgentClaim 
         Caption         =   "<agent claim none>"
         Height          =   195
         Left            =   1380
         TabIndex        =   23
         Top             =   1020
         Width           =   1635
      End
      Begin VB.Label Label7 
         Caption         =   "Agent Claim:"
         Height          =   195
         Left            =   300
         TabIndex        =   22
         Top             =   1020
         Width           =   1035
      End
      Begin VB.Label LblNama 
         Caption         =   "<nama none>"
         Height          =   195
         Left            =   960
         TabIndex        =   21
         Top             =   660
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "Nama:"
         Height          =   195
         Left            =   300
         TabIndex        =   20
         Top             =   660
         Width           =   795
      End
      Begin VB.Label lblcustid 
         Caption         =   "<custid none>"
         Height          =   195
         Left            =   960
         TabIndex        =   19
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Custid:"
         Height          =   195
         Left            =   300
         TabIndex        =   18
         Top             =   360
         Width           =   795
      End
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   2220
      TabIndex        =   16
      Top             =   6780
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox TxtJmlh 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   15
      Text            =   "0"
      Top             =   6780
      Width           =   855
   End
   Begin VB.CommandButton CmdPindahkanKe 
      Caption         =   "Pindahkan Ke..."
      Height          =   315
      Left            =   7740
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.ComboBox CmbAgent 
      Height          =   315
      Left            =   7140
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.CommandButton CmdPindahkanKeAgentAwal 
      Caption         =   "Reject Claim"
      Height          =   375
      Left            =   10140
      TabIndex        =   11
      Top             =   60
      Width           =   2235
   End
   Begin VB.CommandButton CmdApproveClaim 
      Caption         =   "Approve claim"
      Height          =   375
      Left            =   6840
      TabIndex        =   9
      Top             =   60
      Width           =   1515
   End
   Begin VB.CommandButton CmdUnCekAll 
      Caption         =   "UnCekAll"
      Height          =   375
      Left            =   5700
      TabIndex        =   8
      Top             =   420
      Width           =   1095
   End
   Begin VB.CommandButton CmdCekAll 
      Caption         =   "Cek All"
      Height          =   375
      Left            =   5700
      TabIndex        =   7
      Top             =   60
      Width           =   1095
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   4380
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton CmdCari 
      Caption         =   "&Cari"
      Height          =   375
      Left            =   4380
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox TxtNama 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   420
      Width           =   2175
   End
   Begin VB.TextBox TxtCustid 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   420
      Width           =   1815
   End
   Begin MSComctlLib.ListView LvClaimAcc 
      Height          =   5775
      Left            =   60
      TabIndex        =   0
      Top             =   840
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   10186
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label4 
      Caption         =   "Jumlah data:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   6780
      Width           =   1395
   End
   Begin VB.Label Label3 
      Caption         =   "Pindahkan ke agent ini =>"
      Height          =   195
      Left            =   6900
      TabIndex        =   10
      Top             =   540
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   5580
      X2              =   5580
      Y1              =   60
      Y2              =   840
   End
   Begin VB.Label Label2 
      Caption         =   "Nama:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2100
      TabIndex        =   3
      Top             =   120
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Custid:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1395
   End
End
Attribute VB_Name = "FrmListClaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rs_temp As ADODB.Recordset

Private Sub HeaderAcc()
    LvClaimAcc.ColumnHeaders.ADD 1, , "Custid", 1500
    LvClaimAcc.ColumnHeaders.ADD 2, , "Nama", 1500
    LvClaimAcc.ColumnHeaders.ADD 3, , "AGENT CLAIM", 1500
    LvClaimAcc.ColumnHeaders.ADD 4, , "WAKTU CLAIM", 1500
    LvClaimAcc.ColumnHeaders.ADD 5, , "ALASAN CLAIM", 3000
    LvClaimAcc.ColumnHeaders.ADD 6, , "Agent Awal", 1500
End Sub


Private Sub CmdApproveClaim_Click()
    Dim CMDSQL As String
    Dim W, K, S As Integer
    Dim a, pesan, RemarksClaim As String
    
    On Error GoTo Salah
    If LvClaimAcc.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    S = 0
    
    For K = 1 To LvClaimAcc.ListItems.Count
        If LvClaimAcc.ListItems(K).Checked = True Then
            S = S + 1
            Exit For
        End If
    Next K
    
    If S = 0 Then
        MsgBox "Anda belum memilih account yang akan di approve Claim nya!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    a = MsgBox("Anda yakin akan mengapprove claim account yang diceklist?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbNo Then
        MsgBox "Proses dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    
    LvClaimAcc.Enabled = False
    CmdCekAll.Enabled = False
    CmdUnCekAll.Enabled = False
    CmdCari.Enabled = False
    
    CmdApproveClaim.Enabled = False
    CmdPindahkanKeAgentAwal.Enabled = False
    CmdPindahkanKe.Enabled = False
    
    PB1.Max = LvClaimAcc.ListItems.Count
    For W = 1 To LvClaimAcc.ListItems.Count
        PB1.Value = W
        If LvClaimAcc.ListItems(W).Checked = True Then
            
            DoEvents
            'Pindahkan ke agent yang meng claim
'            cmdsql = "UPDATE mgm SET agent=user_claim,user_claim=null,"
'            cmdsql = cmdsql + " waktu_claim=null,alasan_claim=null WHERE custid='"
'            cmdsql = cmdsql + CStr(LvClaimAcc.ListItems(W).Text) + "' AND user_claim IS NOT NULL "
'            M_OBJCONN.Execute cmdsql

            ' UNTUK CLAIM 02 APRIL 2014
            CMDSQL = "UPDATE mandiri.mgm SET agent=user_claim,user_claim=null,"
            CMDSQL = CMDSQL + " waktu_claim=null,alasan_claim=null,app_claim=now() WHERE custid='"
            CMDSQL = CMDSQL + CStr(LvClaimAcc.ListItems(W).Text) + "' AND user_claim IS NOT NULL "
            M_OBJCONN.Execute CMDSQL
            ' ------------------------
'
'            cmdsql = "UPDATE mgm SET agent=user_claim WHERE custid='" + CStr(LvClaimAcc.ListItems(W).Text) + "'"
'            M_OBJCONN.Execute cmdsql
'
'            cmdsql = "update mgm set user_claim=null,"
'            cmdsql = cmdsql + " waktu_claim=null,alasan_claim=null,app_claim=now() where custid='"
'            cmdsql = cmdsql + CStr(LvClaimAcc.ListItems(W).Text) + "'"
'            M_OBJCONN.Execute cmdsql
            
            RemarksClaim = "APPROVE CLAIM: "
            RemarksClaim = RemarksClaim & "Oleh => "
            RemarksClaim = RemarksClaim & MDIForm1.TxtUsername.Text & ",Custid =>"
            RemarksClaim = RemarksClaim & CStr(LvClaimAcc.ListItems(W).Text) & ", ke agent => "
            RemarksClaim = RemarksClaim & CStr(LvClaimAcc.ListItems(W).SubItems(2))
            
            'Catet ke mgm hst
            CMDSQL = "insert into mandiri.mgm_hst (custid,agent,hst,tgl,user_log) values ('"
            CMDSQL = CMDSQL + CStr(LvClaimAcc.ListItems(W).Text) + "','"
            CMDSQL = CMDSQL & "CLAIM','"
            CMDSQL = CMDSQL & RemarksClaim & "',now(),'"
            CMDSQL = CMDSQL & MDIForm1.TxtUsername.Text & "')"
            M_OBJCONN.Execute CMDSQL
            
            ' TABEL APPROVE CLAIM ============
            CMDSQL = "INSERT INTO mandiri.tbl_approve_claim(custid,nama,tgl_claim,agent_claim,alasan,agent_asli) "
            CMDSQL = CMDSQL + " VALUES('" & LvClaimAcc.ListItems(W).Text & "','" & LvClaimAcc.ListItems(W).SubItems(1) & "',"
            CMDSQL = CMDSQL + "'" & LvClaimAcc.ListItems(W).SubItems(3) & "','" & LvClaimAcc.ListItems(W).SubItems(2) & "',"
            CMDSQL = CMDSQL + "'" & LvClaimAcc.ListItems(W).SubItems(4) & "','" & LvClaimAcc.ListItems(W).SubItems(5) & "');"
            M_OBJCONN.Execute CMDSQL
            ' ================================
            
            pesan = "Pesan dibuat otomatis oleh sistem " & vbCrLf
            pesan = pesan & "================================== " & vbCrLf
            pesan = pesan & "Account dengan :" & vbCrLf
            pesan = pesan & "Custid: " & CStr(LvClaimAcc.ListItems(W).Text) & vbCrLf
            pesan = pesan & "Nama: " & CStr(LvClaimAcc.ListItems(W).SubItems(1)) & vbCrLf & vbCrLf
            pesan = pesan & "Sekarang account ini telah dipindah ke " & CStr(LvClaimAcc.ListItems(W).SubItems(2)) & vbCrLf
            pesan = pesan & "karena SPV telah mengapprove claim account untuk agent ini!"
            
            'MALIK NIH
            'Kirim pesan ke agent yang meng claim
            CMDSQL = "insert into mandiri.msgtbl "
            CMDSQL = CMDSQL + "( recipient, datetime, sender, sentfrom, msg) values ('"
            CMDSQL = CMDSQL + LvClaimAcc.ListItems(W).SubItems(2) + "','"
            CMDSQL = CMDSQL + Format(Now(), "yyyymmdd") + "','"
            CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "','"
            CMDSQL = CMDSQL + CStr(MDIForm1.Winsock1.LocalIP) + "','"
            CMDSQL = CMDSQL + pesan + "')"
            M_OBJCONN.Execute CMDSQL

            'Kirim pesan ke agent yang lama
            CMDSQL = "insert into mandiri.msgtbl "
            CMDSQL = CMDSQL + "( recipient, datetime, sender, sentfrom, msg) values ('"
            CMDSQL = CMDSQL + LvClaimAcc.ListItems(W).SubItems(5) + "','"
            CMDSQL = CMDSQL + Format(Now(), "yyyymmdd") + "','"
            CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "','"
            CMDSQL = CMDSQL + CStr(MDIForm1.Winsock1.LocalIP) + "','"
            CMDSQL = CMDSQL + pesan + "')"
            M_OBJCONN.Execute CMDSQL

            'Update pesan resetnya,pada account yang meng claim
            CMDSQL = "update mandiri.usertbl set f_pesanresetauto='1' where userid='"
            CMDSQL = CMDSQL + CStr(LvClaimAcc.ListItems(W).SubItems(2)) + "'"
            M_OBJCONN.Execute CMDSQL
            
        End If
    Next W
    
    IsiAccClaim
    
    LvClaimAcc.Enabled = True
    CmdCekAll.Enabled = True
    CmdUnCekAll.Enabled = True
    CmdCari.Enabled = True
    
    CmdApproveClaim.Enabled = True
    CmdPindahkanKeAgentAwal.Enabled = True
    CmdPindahkanKe.Enabled = True
    
    MsgBox "Proses approve berhasil!", vbOKOnly + vbInformation, "Informasi"
    
    Exit Sub
Salah:
    MsgBox "Maaf ada error: " & err.Description, vbOKOnly + vbInformation, "Informasi"
    
    LvClaimAcc.Enabled = True
    CmdCekAll.Enabled = True
    CmdUnCekAll.Enabled = True
    CmdCari.Enabled = True
    
    CmdApproveClaim.Enabled = True
    CmdPindahkanKeAgentAwal.Enabled = True
    CmdPindahkanKe.Enabled = True
    
End Sub

Private Sub CmdCari_Click()
    Call IsiAccClaim
End Sub

Private Sub CmdCekAll_Click()
    Dim K As Integer
    
    If LvClaimAcc.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For K = 1 To LvClaimAcc.ListItems.Count
        LvClaimAcc.ListItems(K).Checked = True
    Next K
End Sub

Private Sub CmdClear_Click()
    TxtCustid.Text = ""
    txtnama.Text = ""
End Sub

Private Sub CmdPindahkanKeAgentAwal_Click()
    Dim CMDSQL As String
    Dim pesan, a As String
    Dim W, K, S As Integer
    Dim RemarksClaim As String
    
    On Error GoTo Salah
    
    If LvClaimAcc.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    S = 0
    
    For W = 1 To LvClaimAcc.ListItems.Count
        If LvClaimAcc.ListItems(W).Checked = True Then
            S = S + 1
            Exit For
        End If
    Next W
    
    If S = 0 Then
        MsgBox "Anda belum menceklist account yang akan dikembalikan ke agent awal!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    a = MsgBox("Anda yakin akan mengembalikan account yang di ceklist?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbNo Then
        MsgBox "Proses dibatalkan!", vbOKOnly + vbExclamation, "Informasi"
        Exit Sub
    End If
    
    LvClaimAcc.Enabled = False
    CmdCekAll.Enabled = False
    CmdUnCekAll.Enabled = False
    CmdCari.Enabled = False
    
    CmdApproveClaim.Enabled = False
    CmdPindahkanKeAgentAwal.Enabled = False
    CmdPindahkanKe.Enabled = False
    
    
    DoEvents
    PB1.Max = LvClaimAcc.ListItems.Count
    For K = 1 To LvClaimAcc.ListItems.Count
        PB1.Value = K
        If LvClaimAcc.ListItems(K).Checked = True Then
            'Update ke agent awal!
            If cek_agent_lama.Value = True Then
                CMDSQL = "update mandiri.mgm set agent=agent_asli,user_claim=null, "
                CMDSQL = CMDSQL + " waktu_claim=null,alasan_claim=null where custid='"
                CMDSQL = CMDSQL & CStr(LvClaimAcc.ListItems(K).Text) + "'"
                M_OBJCONN.Execute CMDSQL
                
                RemarksClaim = "PENGEMBALIAN ACCOUNT ke Agent AWAL : Oleh=> "
                RemarksClaim = RemarksClaim & MDIForm1.TxtUsername.Text & ", Custid=> "
                RemarksClaim = RemarksClaim & CStr(LvClaimAcc.ListItems(W).Text) & ", Ke agent=> "
                RemarksClaim = RemarksClaim & CStr(LvClaimAcc.ListItems(W).SubItems(5))
                
                'Catet ke mgm hst
                CMDSQL = "insert into mandiri.mgm_hst (custid,agent,hst,tgl,user_log) values ('"
                CMDSQL = CMDSQL + CStr(LvClaimAcc.ListItems(W).Text) + "','"
                CMDSQL = CMDSQL & "CLAIM','"
                CMDSQL = CMDSQL & RemarksClaim & "',now(),'"
                CMDSQL = CMDSQL & MDIForm1.TxtUsername.Text & "')"
                M_OBJCONN.Execute CMDSQL
                
                'Update pesan resetnya,pada agent yang asli
                CMDSQL = "update mandiri.usertbl set f_pesanresetauto='1' where userid='"
                CMDSQL = CMDSQL + CStr(LvClaimAcc.ListItems(W).SubItems(5)) + "'"
                M_OBJCONN.Execute CMDSQL
                
                pesan = "Pesan dibuat otomatis oleh system " & vbCrLf
                pesan = pesan & "============================" & vbCrLf
                pesan = pesan & "Account dengan : " & vbCrLf & vbCrLf
                pesan = pesan & "Custid :" & CStr(LvClaimAcc.ListItems(W).Text) & vbCrLf
                pesan = pesan & "Nama : " & CStr(LvClaimAcc.ListItems(W).SubItems(1)) & vbCrLf & vbCrLf
                pesan = pesan & "Telah dikembalikan ke agent : " & LvClaimAcc.ListItems(W).SubItems(5) & vbCrLf
                pesan = pesan & "Oleh: " & MDIForm1.TxtUsername.Text
            Else
                CMDSQL = "update mandiri.mgm set agent='AKSESALL',user_claim=null, "
                CMDSQL = CMDSQL + " waktu_claim=null,alasan_claim=null where custid='"
                CMDSQL = CMDSQL & CStr(LvClaimAcc.ListItems(K).Text) + "'"
                M_OBJCONN.Execute CMDSQL
                
                RemarksClaim = "PENGEMBALIAN ACCOUNT ke Agent AKSESALL : Oleh=> "
                RemarksClaim = RemarksClaim & MDIForm1.TxtUsername.Text & ", Custid=> "
                RemarksClaim = RemarksClaim & CStr(LvClaimAcc.ListItems(K).Text) & ", Ke agent=> AKSESALL"
                
                'Catet ke mgm hst
                CMDSQL = "insert into mandiri.mgm_hst (custid,agent,hst,tgl,user_log) values ('"
                CMDSQL = CMDSQL + CStr(LvClaimAcc.ListItems(K).Text) + "','"
                CMDSQL = CMDSQL & "CLAIM','"
                CMDSQL = CMDSQL & RemarksClaim & "',now(),'"
                CMDSQL = CMDSQL & MDIForm1.TxtUsername.Text & "')"
                M_OBJCONN.Execute CMDSQL
                
                'balikin ke tbl_cust_aksesall
                CMDSQL = "INSERT INTO mandiri.tbl_cust_aksesall"
                CMDSQL = CMDSQL & "(SELECT max(kd_profile), '" & CStr(LvClaimAcc.ListItems(K).Text) & "' FROM tbl_cust_aksesall)"
                M_OBJCONN.Execute CMDSQL
                
                'Update pesan resetnya,pada agent yang asli
                CMDSQL = "update mandiri.usertbl set f_pesanresetauto='1' where userid='"
                CMDSQL = CMDSQL + CStr(LvClaimAcc.ListItems(W).SubItems(5)) + "'"
                M_OBJCONN.Execute CMDSQL
                
                pesan = "Pesan dibuat otomatis oleh system " & vbCrLf
                pesan = pesan & "============================" & vbCrLf
                pesan = pesan & "Account dengan : " & vbCrLf & vbCrLf
                pesan = pesan & "Custid :" & CStr(LvClaimAcc.ListItems(W).Text) & vbCrLf
                pesan = pesan & "Nama : " & CStr(LvClaimAcc.ListItems(W).SubItems(1)) & vbCrLf & vbCrLf
                pesan = pesan & "Telah dikembalikan ke agent : AKSESALL " & vbCrLf
                pesan = pesan & "Oleh: " & MDIForm1.TxtUsername.Text
            End If
            
'            'Catet ke mgm hst
'            cmdsql = "insert into mgm_hst (custid,agent,hst,tgl,user_log) values ('"
'            cmdsql = cmdsql + CStr(LvClaimAcc.ListItems(w).Text) + "','"
'            cmdsql = cmdsql & "CLAIM','"
'            cmdsql = cmdsql & RemarksClaim & "',now(),'"
'            cmdsql = cmdsql & mdiform1.txtusername.text & "')"
'            M_OBJCONN.Execute cmdsql
            
'            'Update pesan resetnya,pada agent yang asli
'            cmdsql = "update usertbl set f_pesanresetauto='1' where userid='"
'            cmdsql = cmdsql + CStr(LvClaimAcc.ListItems(w).SubItems(5)) + "'"
'            M_OBJCONN.Execute cmdsql
            
'            pesan = "Pesan dibuat otomatis oleh system " & vbCrLf
'            pesan = pesan & "============================" & vbCrLf
'            pesan = pesan & "Account dengan : " & vbCrLf & vbCrLf
'            pesan = pesan & "Custid :" & CStr(LvClaimAcc.ListItems(w).Text) & vbCrLf
'            pesan = pesan & "Nama : " & CStr(LvClaimAcc.ListItems(w).SubItems(1)) & vbCrLf & vbCrLf
'            pesan = pesan & "Telah dikembalikan ke agent : " & LvClaimAcc.ListItems(w).SubItems(5) & vbCrLf
'            pesan = pesan & "Oleh: " & mdiform1.txtusername.text
            
            'Kirim pesan ke agent yang lama
            CMDSQL = "insert into mandiri.msgtbl "
            CMDSQL = CMDSQL + "( recipient, datetime, sender, sentfrom, msg) values ('"
            CMDSQL = CMDSQL + LvClaimAcc.ListItems(W).SubItems(5) + "','"
            CMDSQL = CMDSQL + Format(Now(), "yyyymmdd") + "','"
            CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "','"
            CMDSQL = CMDSQL + CStr(MDIForm1.Winsock1.LocalIP) + "','"
            CMDSQL = CMDSQL + pesan + "')"
            M_OBJCONN.Execute CMDSQL
        End If
    Next K
    
    LvClaimAcc.Enabled = True
    CmdCekAll.Enabled = True
    CmdUnCekAll.Enabled = True
    CmdCari.Enabled = True
    
    CmdApproveClaim.Enabled = True
    CmdPindahkanKeAgentAwal.Enabled = True
    CmdPindahkanKe.Enabled = True
    cek_agent_lama.Value = False
    Call IsiAccClaim
    
    MsgBox "Proses Reject Berhasil!", vbOKOnly + vbInformation, "Informasi"
    
    Exit Sub
Salah:
    MsgBox "Maaf ada error: " & err.Description, vbOKOnly + vbExclamation, "Informasi"
End Sub

Private Sub CmdUnCekAll_Click()
    Dim K As Integer
    
    If LvClaimAcc.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For K = 1 To LvClaimAcc.ListItems.Count
        LvClaimAcc.ListItems(K).Checked = False
    Next K
End Sub

Private Sub Command1_Click()
    If rs_temp.state = 1 Then rs_temp.Close
    rs_temp.Open "SELECT * FROM mandiri.temp_exp_claim ;"
    
    CD.Filter = "Excel Files (*.xls) | *.xls"
    CD.ShowSave
    
    ConvertToExcel rs_temp, CD.FileName
End Sub

Private Sub Form_Load()
    Call HeaderAcc
    Call IsiComboAgent
    
    Set rs_temp = New ADODB.Recordset
    rs_temp.ActiveConnection = M_OBJCONN
    rs_temp.CursorLocation = adUseClient
    rs_temp.CursorType = adOpenDynamic
    rs_temp.LockType = adLockOptimistic
    
    cek_agent_lama.Value = False
End Sub

Private Sub IsiAccClaim()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim ListItem As ListItem
    Dim M_WHERE As String
    
    M_WHERE = ""
    
    CMDSQL = "select * from mandiri.mgm  "
    
    Command1.Enabled = False
    
    If TxtCustid.Text <> "" Then
        If M_WHERE = "" Then
            M_WHERE = " where custid like '%" + CStr(TxtCustid.Text) + "%' "
        Else
            M_WHERE = M_WHERE & " and custid like '%" + CStr(TxtCustid.Text) + "%' "
        End If
    End If
    
    If txtnama.Text <> "" Then
        If M_WHERE = "" Then
            M_WHERE = " where name like '%" + CStr(txtnama.Text) + "%' "
        Else
            M_WHERE = M_WHERE & " and name like '%" + CStr(txtnama.Text) + "%' "
        End If
    End If
       
    If M_WHERE <> "" Then
        M_WHERE = M_WHERE & " and agent  in ('CLAIM') order by name asc "
    Else
        M_WHERE = " where agent  in ('CLAIM') order by name asc "
    End If
       
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL & M_WHERE, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvClaimAcc.ListItems.CLEAR
    TxtJmlh.Text = M_objrs.RecordCount
    
    
    If M_objrs.RecordCount = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Peringatan"
        Set M_objrs = Nothing
        Exit Sub
    End If
    
    PB1.Max = M_objrs.RecordCount
    M_OBJCONN.Execute "DELETE FROM mandiri.temp_exp_claim ;"
    While Not M_objrs.EOF
        PB1.Value = M_objrs.Bookmark
        Set ListItem = LvClaimAcc.ListItems.ADD(, , M_objrs("custid"))
            ListItem.SubItems(1) = M_objrs("name")
            ListItem.SubItems(2) = IIf(IsNull(M_objrs("user_claim")), "", M_objrs("user_claim"))
            ListItem.SubItems(3) = IIf(IsNull(M_objrs("waktu_claim")), "", Format(M_objrs("waktu_claim"), "yyyy-mm-dd hh:nn:ss"))
            ListItem.SubItems(4) = IIf(IsNull(M_objrs("alasan_claim")), "", M_objrs("alasan_claim"))
            ListItem.SubItems(5) = IIf(IsNull(M_objrs("agent_asli")), "", M_objrs("agent_asli"))
            'listitem.SubItems(6) = IIf(IsNull(M_Objrs("waktu_akses")), "", Format(M_Objrs("waktu_akses"), "yyyy-mm-dd hh:nn:ss"))
            
'            If UCase(M_Objrs("agent")) = "AKSESALL" Then
'                listitem.ForeColor = vbRed
'                listitem.ListSubItems(1).ForeColor = vbRed
'                listitem.ListSubItems(2).ForeColor = vbRed
'                listitem.ListSubItems(3).ForeColor = vbRed
'                listitem.ListSubItems(4).ForeColor = vbRed
'                listitem.ListSubItems(5).ForeColor = vbRed
'                listitem.ListSubItems(6).ForeColor = vbRed
'            End If
'
'            If UCase(M_Objrs("agent")) = "#KOSONG#" Then
'                listitem.ForeColor = vbBlue
'                listitem.ListSubItems(1).ForeColor = vbBlue
'                listitem.ListSubItems(2).ForeColor = vbBlue
'                listitem.ListSubItems(3).ForeColor = vbBlue
'                listitem.ListSubItems(4).ForeColor = vbBlue
'                listitem.ListSubItems(5).ForeColor = vbBlue
'                listitem.ListSubItems(6).ForeColor = vbBlue
'            End If
            
            M_OBJCONN.Execute "insert into mandiri.temp_exp_claim(custid,name,user_claim,waktu_claim,alasan_claim,agent_asli) " & _
                            " values('" & M_objrs("custid") & "','" & M_objrs("name") & "','" & IIf(IsNull(M_objrs("user_claim")), "", M_objrs("user_claim")) & "','" & IIf(IsNull(M_objrs("waktu_claim")), "", Format(M_objrs("waktu_claim"), "yyyy-mm-dd hh:nn:ss")) & "','" & IIf(IsNull(M_objrs("alasan_claim")), "", M_objrs("alasan_claim")) & "','" & IIf(IsNull(M_objrs("agent_asli")), "", M_objrs("agent_asli")) & "')"
            
        M_objrs.MoveNext
    Wend
    Command1.Enabled = True
    Set M_objrs = Nothing
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Set rs_temp = Nothing
End Sub

Private Sub LvClaimAcc_Click()
    If LvClaimAcc.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    lblcustid.Caption = LvClaimAcc.SelectedItem.Text
    lblnama.Caption = IIf(IsNull(LvClaimAcc.SelectedItem.SubItems(1)), "", LvClaimAcc.SelectedItem.SubItems(1))
    LblAgentClaim.Caption = IIf(IsNull(LvClaimAcc.SelectedItem.SubItems(2)), "", LvClaimAcc.SelectedItem.SubItems(2))
    LblWaktuClaim.Caption = IIf(IsNull(LvClaimAcc.SelectedItem.SubItems(3)), "", LvClaimAcc.SelectedItem.SubItems(3))
    LblAlasanClaim.Caption = IIf(IsNull(LvClaimAcc.SelectedItem.SubItems(4)), "", LvClaimAcc.SelectedItem.SubItems(4))
End Sub

Private Sub IsiComboAgent()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    
    CmbAgent.CLEAR
    CmbAgent.AddItem "ALL"
    
    CMDSQL = "select * from mandiri.usertbl where usertype in ('1','6') and userid "
    CMDSQL = CMDSQL & " not in ('LUNAS','COMPLAIN','COMPLAIN','CLAIM','AKSESALL') and userid is not null order by userid asc "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            CmbAgent.AddItem M_objrs("userid")
            M_objrs.MoveNext
        Wend
    End If
    
    Set M_objrs = Nothing
End Sub

Private Sub LvClaimAcc_DblClick()
    
    If LvClaimAcc.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    VIEW_MGMDATA.txtnocard.Text = LvClaimAcc.SelectedItem.Text
    Me.Hide
    FrmDistribusiAcc.Hide
    VIEW_MGMDATA.Show
End Sub
