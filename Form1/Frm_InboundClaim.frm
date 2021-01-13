VERSION 5.00
Begin VB.Form Frm_InboundClaim 
   Caption         =   "Input Claim"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6090
   Icon            =   "Frm_InboundClaim.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   1935
   ScaleWidth      =   6090
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtHandphone 
      Height          =   345
      Left            =   2520
      TabIndex        =   3
      Top             =   1110
      Width           =   1995
   End
   Begin VB.TextBox TxtNoTelpKantor 
      Height          =   345
      Left            =   2520
      TabIndex        =   2
      Top             =   765
      Width           =   1995
   End
   Begin VB.TextBox txtnotelprumah 
      Height          =   345
      Left            =   2520
      TabIndex        =   1
      Top             =   420
      Width           =   1995
   End
   Begin VB.TextBox txtnama 
      Height          =   360
      Left            =   2520
      TabIndex        =   0
      Top             =   60
      Width           =   3345
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   330
      Left            =   4815
      TabIndex        =   5
      Top             =   1530
      Width           =   900
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save"
      Height          =   330
      Left            =   3855
      TabIndex        =   4
      Top             =   1530
      Width           =   900
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Handphone :"
      Height          =   270
      Index           =   3
      Left            =   0
      TabIndex        =   9
      Top             =   1110
      Width           =   2370
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "No Telp Kantor :"
      Height          =   270
      Index           =   2
      Left            =   0
      TabIndex        =   8
      Top             =   765
      Width           =   2370
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "No Telp Rumah :"
      Height          =   270
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   405
      Width           =   2370
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Nama Pemegang Kartu Kredit :"
      Height          =   270
      Index           =   0
      Left            =   15
      TabIndex        =   6
      Top             =   105
      Width           =   2370
   End
End
Attribute VB_Name = "Frm_InboundClaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdSave_Click()
Dim m_objrs As ADODB.Recordset
Dim cmdsql As String
Dim Lcustid, LName, LHOMENO, LOFFICENO, LMOBILE, LAgent, LNAMAAGENT, LRECSOURCE, LOTHERS, LKethslkerja As String
If Len(txtnama.Text) = 0 Or Len(txtnotelprumah.Text) = 0 Then
    MsgBox "Nama dan no telp rumah harus diisi", vbInformation + vbOKOnly, "Informasi"
    Exit Sub
End If

Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
cmdsql = "Select  * from MGM where LEFT(recsource,3) <>'PRE' AND HOMENO like '%" + txtnotelprumah.Text + "%' "
m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not m_objrs.EOF
    LKethslkerja = m_objrs!KETHSLKERJA
    LName = m_objrs!Name
    If UCase(Trim(LName)) = UCase(Trim(txtnama.Text)) Then
        If UCase(LKethslkerja) = "1A" Then
            m_objrs!NEXTACT = "Data Inbound Call"
            m_objrs!agent = MDIForm1.Text1.Text
            m_objrs!NAMAAGENT = MDIForm1.Text7.Text
            m_objrs.UPDATE
            MsgBox "Sukses... Data Sudah Ditransfer", vbInformation + vbOKOnly, "Informasi"
            Set m_objrs = Nothing
            txtnama.Text = ""
            txtnotelprumah.Text = ""
            TxtNoTelpKantor.Text = ""
            TxtHandphone.Text = ""
            txtnama.SetFocus
            Exit Sub
        Else
            MsgBox "Tidak Sukses... Data Sudah di follow Up oleh " & m_objrs!agent
            Set m_objrs = Nothing
            txtnama.Text = ""
            txtnotelprumah.Text = ""
            TxtNoTelpKantor.Text = ""
            TxtHandphone.Text = ""
            txtnama.SetFocus
            Exit Sub
        End If
    End If
m_objrs.MoveNext
Wend
Set m_objrs = Nothing


Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
cmdsql = "Select  * from tempCC_CUSTTBL where LEFT(recsource,3) <>'PRE' AND HOMENO like '%" + txtnotelprumah.Text + "%'"
m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not m_objrs.EOF
    LKethslkerja = m_objrs!KETHSLKERJA
    lNama = m_objrs!Name
    If UCase(Trim(Replace(lNama, "(PVA)", ""))) = UCase(Trim(txtnama.Text)) Then
        Lcustid = m_objrs!CUSTID
        LName = m_objrs!Name
        LHOMENO = m_objrs!HOMENO
        LOFFICENO = m_objrs!OFFICENO
        LMOBILE = m_objrs!MOBILENO
        LAgent = m_objrs!agent
        LNAMAAGENT = m_objrs!NAMAAGENT
        LRECSOURCE = m_objrs!RECSOURCE
        LOTHERS = m_objrs!OTHERS
        LKethslkerja = m_objrs!KETHSLKERJA
        cmdsql = "Insert Into mgm (CUSTID, NAME, HOMENO, OFFICENO, MOBILENO, AGENT, NAMAAGENT, RECSOURCE, nextact, KETHSLKERJA)"
        cmdsql = cmdsql + " VALUES"
        cmdsql = cmdsql + "('" + Lcustid + "',"
        cmdsql = cmdsql + " '" + LName + "',"
        cmdsql = cmdsql + " '" + LHOMENO + "',"
        cmdsql = cmdsql + " '" + LOFFICENO + "',"
        cmdsql = cmdsql + " '" + LMOBILE + "',"
        cmdsql = cmdsql + " '" + MDIForm1.Text1.Text + "',"
        cmdsql = cmdsql + " '" + MDIForm1.Text7.Text + "',"
        cmdsql = cmdsql + " '" + LRECSOURCE + "',"
        cmdsql = cmdsql + " 'Data Inbound Call',"
        cmdsql = cmdsql + " '" + LKethslkerja + "')"
        M_OBJCONN.Execute cmdsql
        m_objrs.DELETE adAffectCurrent
        MsgBox "Sukses... Data Sudah Ditransfer", vbInformation + vbOKOnly, "Informasi"
        Set m_objrs = Nothing
        txtnama.Text = ""
        txtnotelprumah.Text = ""
        TxtNoTelpKantor.Text = ""
        TxtHandphone.Text = ""
        txtnama.SetFocus
        Exit Sub
    End If
m_objrs.MoveNext
Wend
Set m_objrs = Nothing

        cmdsql = "Insert Into mgm (CUSTID, NAME, HOMENO, OFFICENO, MOBILENO, AGENT, NAMAAGENT, RECSOURCE, nextact, KETHSLKERJA)"
        cmdsql = cmdsql + " VALUES"
        Lcustid = "MGMI-" & CUSTNOMOR(M_OBJCONN, "FRMCUST_CC")
        
        cmdsql = cmdsql + "('" + Lcustid + "',"
        cmdsql = cmdsql + " '" + txtnama.Text + "',"
        cmdsql = cmdsql + " '" + txtnotelprumah.Text + "',"
        cmdsql = cmdsql + " '" + TxtNoTelpKantor.Text + "',"
        cmdsql = cmdsql + " '" + TxtHandphone.Text + "',"
        cmdsql = cmdsql + " '" + MDIForm1.Text1.Text + "',"
        cmdsql = cmdsql + " '" + MDIForm1.Text7.Text + "',"
        cmdsql = cmdsql + " 'MGM_INC',"
        cmdsql = cmdsql + " 'Data Inbound Call',"
        cmdsql = cmdsql + " '1A')"
M_OBJCONN.Execute cmdsql
MsgBox "Sukses..", vbInformation, "Informasi"
txtnama.Text = ""
txtnotelprumah.Text = ""
TxtNoTelpKantor.Text = ""
TxtHandphone.Text = ""
txtnama.SetFocus
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

