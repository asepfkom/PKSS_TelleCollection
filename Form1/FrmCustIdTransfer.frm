VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmCustIdTransfer 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9120
   Icon            =   "FrmCustIdTransfer.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9120
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Kembalikan Ke Agent Sebelumnya"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   5
      Left            =   7560
      TabIndex        =   14
      Top             =   4680
      Width           =   1485
   End
   Begin VB.CommandButton CmdUnCekAll 
      BackColor       =   &H00C0C0C0&
      Caption         =   "UnCek All"
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
      Left            =   7560
      TabIndex        =   13
      Top             =   2760
      Width           =   1485
   End
   Begin VB.CommandButton CmdCekall 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cek All"
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
      Left            =   7560
      TabIndex        =   12
      Top             =   2340
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Height          =   1005
      Left            =   855
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   6030
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Load All data"
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
      Index           =   4
      Left            =   7560
      TabIndex        =   8
      Top             =   1845
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Remove"
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
      Index           =   3
      Left            =   7560
      TabIndex        =   7
      Top             =   3900
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Clear"
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
      Index           =   2
      Left            =   7560
      TabIndex        =   6
      Top             =   3360
      Width           =   1485
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Index           =   0
      ItemData        =   "FrmCustIdTransfer.frx":000C
      Left            =   810
      List            =   "FrmCustIdTransfer.frx":000E
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   5505
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Index           =   1
      ItemData        =   "FrmCustIdTransfer.frx":0010
      Left            =   2190
      List            =   "FrmCustIdTransfer.frx":0012
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   5520
      Width           =   3000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
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
      Index           =   1
      Left            =   7545
      TabIndex        =   2
      Top             =   540
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Transfer"
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
      Left            =   7545
      TabIndex        =   1
      Top             =   75
      Width           =   1485
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   840
      ItemData        =   "FrmCustIdTransfer.frx":0014
      Left            =   2700
      List            =   "FrmCustIdTransfer.frx":0016
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   4320
      Visible         =   0   'False
      Width           =   1950
   End
   Begin MSComctlLib.ListView LvTransfer 
      Height          =   5460
      Left            =   60
      TabIndex        =   11
      Top             =   0
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   9631
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
   Begin VB.Label Label2 
      Caption         =   "Remark"
      Height          =   240
      Left            =   180
      TabIndex        =   9
      Top             =   6075
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Name :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   60
      TabIndex        =   5
      Top             =   5565
      Width           =   675
   End
End
Attribute VB_Name = "FrmCustIdTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdCekAll_Click()
    Dim W As Integer
    If LvTransfer.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LvTransfer.ListItems.Count
        LvTransfer.ListItems(W).Checked = True
    Next W
End Sub

Private Sub CmdUnCekAll_Click()
    Dim W As Integer
    If LvTransfer.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LvTransfer.ListItems.Count
        LvTransfer.ListItems(W).Checked = False
    Next W
End Sub

Private Sub Combo2_Change(Index As Integer)
    Combo2(0).Locked = True
    Combo2(1).Locked = True
End Sub

Private Sub Combo2_Click(Index As Integer)
    Call Combo2_LostFocus(Index)
    Combo2(0).Locked = True
    Combo2(1).Locked = True
End Sub

Private Sub Combo2_DropDown(Index As Integer)
    Combo2(0).Locked = False
    Combo2(1).Locked = False
End Sub



Private Sub Combo2_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Combo2_LostFocus(Index As Integer)
Dim M_objrs As ADODB.Recordset
On Error GoTo Combo2_LostFocusErr
Select Case Index
    Case 0
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open "Select * from usertbl where USERID ='" + Combo2(0).text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If Not M_objrs.EOF Then
            Combo2(0).text = M_objrs!USERID
            Combo2(1).text = M_objrs!AGENT
        Else
            Combo2(0).text = Empty
            Combo2(1).text = Empty
        End If
    Case 1
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open "Select * from usertbl where AGENT ='" + Combo2(1).text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If Not M_objrs.EOF Then
            Combo2(0).text = M_objrs!USERID
            Combo2(1).text = M_objrs!AGENT
        Else
            Combo2(0).text = Empty
            Combo2(1).text = Empty
        End If
End Select
Set M_objrs = Nothing
Exit Sub
Combo2_LostFocusErr:
    MsgBox err.Description
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim CMDSQL As String
    Dim i As Integer
    Dim M_objrs As ADODB.Recordset
    Dim CmdsqlLog As String
    Dim ListItem As ListItem
    Dim tglsekarang As Date
    Dim agent_old As String
    
    On Error GoTo adderr
    Select Case Index
        Case 0
            If Combo2(0).text = Empty Then
                MsgBox "Agent harus di isi!", vbOKOnly + vbInformation, "Aplikasi"
                Exit Sub
            End If
            
            '@@ 19/08/2011 cek dulu di lvtransfer sudah ada data apa belum
            If LvTransfer.ListItems.Count = 0 Then
                MsgBox "Data belum tersedia! Klik load data!", vbOKOnly + vbInformation, "Informasi"
                Exit Sub
            End If
            
            '@@ 19/08/2011 buat nyatet waktu log yang melakukan transfer
            CMDSQL = "select now()"
            Set M_objrs = New ADODB.Recordset
            M_objrs.CursorLocation = adUseClient
            M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            
            '@@ 19082011 tadinya pake listbox, sekarang diganti pakai listview saja
    '        For i = 0 To List1.ListCount - 1
    '            CMDSQL = "update mgm set agent ='" + Combo2(0).Text + "'"
    '            CMDSQL = CMDSQL + " Where custid ='" + List1.LIST(i) + "'"
    '            'M_OBJCONN.Execute CMDSQL
    '
    '            '@@ 19/08/2011 buat nyatet log yang melakukan transfer
    '            CmdsqlLog = "insert into log_transfer (tgl_transfer,custid,agent_sebelumnya,"
    '            CmdsqlLog = CmdsqlLog + "agent_sekarang,transfer_oleh) values ('"
    '            CmdsqlLog = CmdsqlLog + CStr(Format(m_objrs(0), "yyyy-mm-dd hh:mm:ss")) + "','"
    '
    '         Next i
                
             
            '@@19/08/2011 Proses Transfer Data
            For i = 1 To LvTransfer.ListItems.Count
                If LvTransfer.ListItems(i).Checked = True Then
    '                Cmdsql = "update mgm set agent ='" + Trim(Combo2(0).Text) + "'"
    '                Cmdsql = Cmdsql + " Where custid ='" + Trim(LvTransfer.ListItems(i).Text) + "'"
                    ' Tambah Flag spv_allow=1 buat izinkan user akses data tersebut ( case : 5 hari akun belum dikerjakan )
                    CMDSQL = "UPDATE mgm SET agent ='" + Trim(Combo2(0).text) + "',nama_agent='" + Trim(Combo2(1).text) + "',spv_allow=now()"
                    CMDSQL = CMDSQL + " WHERE custid ='" + Trim(LvTransfer.ListItems(i).text) + "'"
                    M_OBJCONN.Execute CMDSQL
                    
                    '@@ 19/08/2011 buat nyatet log yang melakukan transfer
                    CmdsqlLog = "insert into log_transfer (tgl_transfer,custid,agent_sebelumnya,"
                    CmdsqlLog = CmdsqlLog + "agent_sekarang,transfer_oleh) values ('"
                    CmdsqlLog = CmdsqlLog + CStr(Format(M_objrs(0), "yyyy-mm-dd hh:mm:ss")) + "','"
                    CmdsqlLog = CmdsqlLog + Trim(LvTransfer.ListItems(i).text) + "','"
                    CmdsqlLog = CmdsqlLog + Trim(LvTransfer.ListItems(i).SubItems(1)) + "','"
                    CmdsqlLog = CmdsqlLog + Trim(Combo2(0).text) + "','"
                    CmdsqlLog = CmdsqlLog + Trim(MDIForm1.TxtUsername.text) + "')"
                    M_OBJCONN.Execute CmdsqlLog
                    
                    ' Hapus log 5x Call diblock - Update 2013-04-25 By Izuddin
                    M_OBJCONN.Execute "DELETE FROM user_phone_log WHERE custid='" & Trim(LvTransfer.ListItems(i).text) & "' " & _
                                        " AND agent='" & Trim(Combo2(0).text) & "'"
                End If
            Next i
            
    
    
    '         cmdsql = "Insert Into TblTransferDataRpt (DilakukanOleh, DiterimaOleh, Dari, Jumlah) Values "
    '         cmdsql = cmdsql + "( '" + mdiform1.txtusername.text + "', "
    '         cmdsql = cmdsql + " '" + mdiform1.txtusername.text + "', "
    '         cmdsql = cmdsql + " '" + mdiform1.txtusername.text + "', "
    '         cmdsql = cmdsql + " " + CStr(i) + ") "
             
             MsgBox "Done"
             Set M_objrs = Nothing
        Case 1
            Unload Me
            b_pindah = False
        Case 2
            'List1.CLEAR
            LvTransfer.ListItems.CLEAR
        Case 3
            List1.RemoveItem List1.ListIndex
        Case 4
           '' List1.Clear
            LvTransfer.ListItems.CLEAR
            For i = 1 To VIEW_MGMDATA.LstVwSearchMgm.ListItems.Count
                'List1.AddItem VIEW_MGMDATA.LstVwSearchMgm.ListItems(i).SubItems(1)
                Set ListItem = LvTransfer.ListItems.ADD(, , VIEW_MGMDATA.LstVwSearchMgm.ListItems(i).SubItems(1))
                    ListItem.SubItems(1) = VIEW_MGMDATA.LstVwSearchMgm.ListItems(i).SubItems(11)
                    ListItem.SubItems(2) = VIEW_MGMDATA.LstVwSearchMgm.ListItems(i).SubItems(3)
            Next i
        Case 5
            If MsgBox("Data akan dibalikkan ke agent Sebelumnya??", vbYesNo + vbQuestion, "Konfirmasi") = vbYes Then
                
                CMDSQL = "select now()"
                Set M_objrs = New ADODB.Recordset
                M_objrs.CursorLocation = adUseClient
                M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                tglsekarang = M_objrs(0)
                
                For i = 1 To LvTransfer.ListItems.Count
                    If LvTransfer.ListItems(i).Checked = True Then
                        If M_objrs.State = 1 Then M_objrs.Close
                        M_objrs.Open "SELECT agent_asli FROM mgm WHERE custid ='" + Trim(LvTransfer.ListItems(i).text) + "'"
                        agent_old = IIf(IsNull(M_objrs!agent_asli), "", M_objrs!agent_asli)
                        
                        ' Kembalikan ke Agent Lama user akses data tersebut ( case : 5 hari akun belum dikerjakan )
                        CMDSQL = "UPDATE mgm SET agent=agent_asli,spv_allow=now()"
                        CMDSQL = CMDSQL + " WHERE custid ='" + Trim(LvTransfer.ListItems(i).text) + "' AND agent_asli IS NOT NULL "
                        M_OBJCONN.Execute CMDSQL
                        
                        '@@ 19/08/2011 buat nyatet log yang melakukan transfer
                        CmdsqlLog = "insert into log_transfer (tgl_transfer,custid,agent_sebelumnya,"
                        CmdsqlLog = CmdsqlLog + "agent_sekarang,transfer_oleh) values ('"
                        CmdsqlLog = CmdsqlLog + CStr(Format(tglsekarang, "yyyy-mm-dd hh:mm:ss")) + "','"
                        CmdsqlLog = CmdsqlLog + Trim(LvTransfer.ListItems(i).text) + "','"
                        CmdsqlLog = CmdsqlLog + Trim(LvTransfer.ListItems(i).SubItems(1)) + "','"
                        CmdsqlLog = CmdsqlLog + agent_old + "','"
                        CmdsqlLog = CmdsqlLog + Trim(MDIForm1.TxtUsername.text) + "')"
                        M_OBJCONN.Execute CmdsqlLog
                        
'                        ' Hapus log 5x Call diblock - Update 2013-04-25 By Izuddin
'                        M_OBJCONN.Execute "DELETE FROM user_phone_log WHERE custid='" & Trim(LvTransfer.ListItems(i).Text) & "' " & _
'                                            " AND agent='" & Trim(Combo2(0).Text) & "'"
                    End If
                Next i
                
                Set M_objrs = Nothing
                
                MsgBox "Data berhasil dikembalikan ke Agent Sebelumnya!!", vbOKOnly + vbInformation, "INFO"
            End If
    End Select
    Exit Sub
adderr:
    MsgBox err.Description
End Sub
Private Sub Form_Load()
    Dim m_combo As New ADODB.Recordset
    Set m_combo = New ADODB.Recordset
    m_combo.CursorLocation = adUseClient
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
If UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Or UCase(MDIForm1.txtlevel.text) = "MANAGER" Then
    CMDSQL = "Select * from usertbl where team = '" & MDIForm1.TxtUsername.text & "' order by userid"
ElseIf UCase(Trim(MDIForm1.txtlevel.text)) = "ADMIN" Or UCase(Trim(MDIForm1.txtlevel.text)) = "ADMINISTRATOR" Then
    CMDSQL = "Select * from usertbl order by userid"
ElseIf UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Then
    '@@ 05-09-2011,team leader bisa pindahin ke timnya
    CMDSQL = "Select * from usertbl where USERTYPE =1 AND TEAM in('" + Trim(MDIForm1.TxtUsername.text) + "','KANTOR')  "
    '@@06062012 UNTUK tl cODING LUNAS dihapus
    CMDSQL = CMDSQL + " and userid<>'LUNAS' "
    CMDSQL = CMDSQL + " order by userid "
    'CMDSQL = "Select * from usertbl where  userid in ('LUNAS COMPLETE','LUNAS PENDING')  order by userid "
'    '@@ 23-04-2012, TL hanya bisa pindahin ke coding REVIEW MILIKNYA
'    CMDSQL = "select * from usertbl where usertype='1' and team='"
'    CMDSQL = CMDSQL + Trim(mdiform1.txtusername.text) + "' and userid like 'REVIEW%'"
ElseIf UCase(Trim(MDIForm1.txtlevel.text)) = "AGENT" Then
    CMDSQL = "Select * from usertbl where usertype='1' and userid like 'REVIEW%' and team in ("
    CMDSQL = CMDSQL + "select team from usertbl where userid='"
    CMDSQL = CMDSQL + Trim(MDIForm1.TxtUsername.text) + "')"
    CMDSQL = CMDSQL + " order by userid "
End If
m_combo.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not m_combo.EOF
    Combo2(0).AddItem m_combo!USERID
    Combo2(1).AddItem m_combo!AGENT
    m_combo.MoveNext
Wend
Set m_combo = Nothing
    
    'List1.CLEAR
 '@@ 19082011 diganti menggunakan listview
 Call HeaderListTransfer
    
 Call IsiCustidOtomatis
End Sub

Private Sub Form_Unload(Cancel As Integer)
b_pindah = False
VIEW_MGMDATA.WindowState = 2
End Sub

Private Sub HeaderListTransfer()
    LvTransfer.ColumnHeaders.ADD , , "Custid", 3000
    LvTransfer.ColumnHeaders.ADD , , "Agent", 1500
    LvTransfer.ColumnHeaders.ADD , , "Nama", 3000
End Sub

Private Sub IsiCustidOtomatis()
    Dim ListItem As ListItem
    
    LvTransfer.ListItems.CLEAR
    
    Set ListItem = LvTransfer.ListItems.ADD(, , VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(1))
                   ListItem.SubItems(1) = VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11)
                   ListItem.SubItems(2) = VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(3)
                   LvTransfer.ListItems(1).Checked = True
End Sub

Private Sub LvTransfer_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
      LvTransfer.SortKey = ColumnHeader.Index - 1
      LvTransfer.Sorted = True
End Sub
