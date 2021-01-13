VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Begin VB.Form FrmBlokAgent 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Blok Agent"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9465
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdKeluar 
      Caption         =   "&Keluar"
      Height          =   495
      Left            =   7860
      TabIndex        =   1
      Top             =   6060
      Width           =   1515
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5955
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   10504
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Agent Login"
      TabPicture(0)   =   "FrmBlokAgent.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LvAgentLogin"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CmdCekAll"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CmdUnCekAll"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CmdBlok"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "TxtJmlAgentLogin"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "CmdRefreshAgentLogin"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Agent Blok"
      TabPicture(1)   =   "FrmBlokAgent.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "LvBlok"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "CmdCekAll2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "CmdUncekAll2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "CmdUnBlok"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "TXtJmlhAgentBlok"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "CmdRefreshAgentBlok"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdhistory"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      Begin VB.CommandButton cmdhistory 
         Caption         =   "History"
         Height          =   495
         Left            =   -66960
         TabIndex        =   16
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton CmdRefreshAgentBlok 
         Caption         =   "Refresh"
         Height          =   435
         Left            =   -66960
         TabIndex        =   15
         Top             =   2220
         Width           =   1215
      End
      Begin VB.CommandButton CmdRefreshAgentLogin 
         Caption         =   "Refresh"
         Height          =   435
         Left            =   8040
         TabIndex        =   14
         Top             =   2220
         Width           =   1215
      End
      Begin VB.TextBox TXtJmlhAgentBlok 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   -66840
         TabIndex        =   13
         Text            =   "0"
         Top             =   5040
         Width           =   1035
      End
      Begin VB.TextBox TxtJmlAgentLogin 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   8100
         TabIndex        =   11
         Text            =   "0"
         Top             =   5160
         Width           =   1035
      End
      Begin VB.CommandButton CmdUnBlok 
         Caption         =   "UnBlok"
         Height          =   435
         Left            =   -66960
         TabIndex        =   9
         Top             =   1620
         Width           =   1215
      End
      Begin VB.CommandButton CmdUncekAll2 
         Caption         =   "Uncek All"
         Height          =   435
         Left            =   -66960
         TabIndex        =   8
         Top             =   1020
         Width           =   1215
      End
      Begin VB.CommandButton CmdCekAll2 
         Caption         =   "Cek All"
         Height          =   435
         Left            =   -66960
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton CmdBlok 
         Caption         =   "Blok"
         Height          =   435
         Left            =   8040
         TabIndex        =   5
         Top             =   1620
         Width           =   1215
      End
      Begin VB.CommandButton CmdUnCekAll 
         Caption         =   "Uncek All"
         Height          =   435
         Left            =   8040
         TabIndex        =   4
         Top             =   1020
         Width           =   1215
      End
      Begin VB.CommandButton CmdCekAll 
         Caption         =   "Cek All"
         Height          =   435
         Left            =   8040
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin MSComctlLib.ListView LvAgentLogin 
         Height          =   5175
         Left            =   180
         TabIndex        =   2
         Top             =   600
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   9128
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
      Begin MSComctlLib.ListView LvBlok 
         Height          =   5175
         Left            =   -74880
         TabIndex        =   6
         Top             =   540
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   9128
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
      Begin VB.Label Label2 
         Caption         =   "Jumlah"
         Height          =   195
         Left            =   -66600
         TabIndex        =   12
         Top             =   4740
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Jumlah"
         Height          =   195
         Left            =   8340
         TabIndex        =   10
         Top             =   4860
         Width           =   615
      End
   End
End
Attribute VB_Name = "FrmBlokAgent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBlok_Click()
    Dim W As Integer
    Dim a As String
    Dim cek As Integer
    Dim K As Integer
    Dim CMDSQL As String
    
    If LvAgentLogin.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    a = MsgBox("Anda yakin akan melakukan blok aplikasi TINS pada agent yang diceklist?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbNo Then
        MsgBox "Blok aplikasi TINS dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    cek = 0
    
    For K = 1 To LvAgentLogin.ListItems.Count
        If LvAgentLogin.ListItems(K).Checked = True Then
            cek = cek + 1
        End If
    Next K
    
    If cek = 0 Then
        MsgBox "Anda belum memilih agent yang akan diblok aplikasi TINSnya!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    For W = 1 To LvAgentLogin.ListItems.Count
        If LvAgentLogin.ListItems(W).Checked = True Then
            CMDSQL = "update usertbl set f_blok='1' where userid='"
            CMDSQL = CMDSQL + Trim(LvAgentLogin.ListItems(W).Text) + "'"
            M_OBJCONN.Execute CMDSQL
        End If
    Next W
    
    MsgBox "Agent yang berhasil diblok aplikasi TINSnya sebanyak:" & cek & " agent!", vbOKOnly + vbInformation, "Informasi"
    Call IsiAgentLogin
End Sub

Private Sub CmdCekAll_Click()
    Dim W As Integer
    
    If LvAgentLogin.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LvAgentLogin.ListItems.Count
        LvAgentLogin.ListItems(W).Checked = True
    Next W
End Sub

Private Sub CmdCekAll2_Click()
    Dim K As Integer
    
    If LvBlok.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Informasi"
        Exit Sub
    End If
    
    For K = 1 To LvBlok.ListItems.Count
        LvBlok.ListItems(K).Checked = True
    Next K
End Sub

Private Sub cmdhistory_Click()
    form_history_blok.Show vbModal
End Sub

Private Sub cmdkeluar_Click()
    Unload Me
End Sub

Private Sub HeaderAgentLogin()
    LvAgentLogin.ColumnHeaders.ADD 1, , "Userid", 2500
    LvAgentLogin.ColumnHeaders.ADD 2, , "Nama", 4000
End Sub

Private Sub HeaderAgentBlok()
    LvBlok.ColumnHeaders.ADD 1, , "Userid", 2500
    LvBlok.ColumnHeaders.ADD 2, , "Nama", 4000
    LvBlok.ColumnHeaders.ADD 3, , "Alasan", 5000
End Sub

Private Sub IsiAgentLogin()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim ListItem As ListItem
    
    CMDSQL = "select * from usertbl where usertype in ('1','6') and f_status_login='1' and "
    CMDSQL = CMDSQL + " f_blok is null and userid is not null and agent is not null  "
    CMDSQL = CMDSQL + " order by userid asc "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    TxtJmlAgentLogin.Text = M_objrs.RecordCount
    LvAgentLogin.ListItems.CLEAR
    
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            Set ListItem = LvAgentLogin.ListItems.ADD(, , M_objrs("userid"))
                ListItem.SubItems(1) = M_objrs("agent")
            M_objrs.MoveNext
        Wend
    End If
    Set M_objrs = Nothing
End Sub

Private Sub isiAgentBlok()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim ListItem As ListItem
    
    CMDSQL = "select * from usertbl where "
    CMDSQL = CMDSQL + " f_blok='1' and userid is not null and agent is not null "
    CMDSQL = CMDSQL + " order by userid asc "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    TXtJmlhAgentBlok.Text = M_objrs.RecordCount
    LvBlok.ListItems.CLEAR
    
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            Set ListItem = LvBlok.ListItems.ADD(, , M_objrs("userid"))
                ListItem.SubItems(1) = M_objrs("agent")
                ListItem.SubItems(2) = IIf(IsNull(M_objrs("alasan_blok")), "", M_objrs("alasan_blok"))
                LvBlok.ForeColor = vbRed
                LvBlok.ListItems(1).ForeColor = vbRed
            M_objrs.MoveNext
        Wend
    End If
    Set M_objrs = Nothing
    
End Sub

Private Sub CmdRefreshAgentBlok_Click()
    Call isiAgentBlok
End Sub

Private Sub CmdRefreshAgentLogin_Click()
    Call IsiAgentLogin
End Sub

Private Sub CmdUnBlok_Click()
    Dim W As Integer
    Dim a As String
    Dim cek As Integer
    Dim K As Integer
    Dim CMDSQL As String
    
    If LvBlok.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    a = MsgBox("Anda yakin akan MEMBUKA BLOK aplikasi TINS pada agent yang diceklist?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbNo Then
        MsgBox "Blok aplikasi TINS dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    cek = 0
    
    For K = 1 To LvBlok.ListItems.Count
        If LvBlok.ListItems(K).Checked = True Then
            cek = cek + 1
        End If
    Next K
    
    If cek = 0 Then
        MsgBox "Anda belum memilih agent yang akan dibuka blok aplikasi TINSnya!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    For W = 1 To LvBlok.ListItems.Count
        If LvBlok.ListItems(W).Checked = True Then
            CMDSQL = "INSERT INTO hstopenblock values (now(), '" + Trim(LvBlok.ListItems(W).Text) + "', '" + MDIForm1.TxtUsername.Text + "', '" + LvBlok.ListItems(W).SubItems(2) + "' )"
            M_OBJCONN.Execute CMDSQL
            CMDSQL = "update usertbl set f_blok=null,last_logout='now()' where userid='"
            CMDSQL = CMDSQL + Trim(LvBlok.ListItems(W).Text) + "'"
            M_OBJCONN.Execute CMDSQL
        End If
    Next W
    
    MsgBox "Agent yang berhasil dibuka blok aplikasi TINSnya sebanyak:" & cek & " agent!", vbOKOnly + vbInformation, "Informasi"
    Call isiAgentBlok
End Sub

Private Sub CmdUnCekAll_Click()
    Dim W As Integer
    
    If LvAgentLogin.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LvAgentLogin.ListItems.Count
        LvAgentLogin.ListItems(W).Checked = False
    Next W
End Sub

Private Sub CmdUncekAll2_Click()
        Dim K As Integer
    
    If LvBlok.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Informasi"
        Exit Sub
    End If
    
    For K = 1 To LvBlok.ListItems.Count
        LvBlok.ListItems(K).Checked = False
    Next K
End Sub

Private Sub Form_Load()
    Call HeaderAgentBlok
    Call HeaderAgentLogin
    
    Call isiAgentBlok
    Call IsiAgentLogin
    
End Sub

Private Sub LvAgentLogin_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LvAgentLogin.SortKey = ColumnHeader.Index - 1
    LvAgentLogin.Sorted = True
End Sub

Private Sub LvBlok_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LvBlok.SortKey = ColumnHeader.Index - 1
    LvBlok.Sorted = True
End Sub
