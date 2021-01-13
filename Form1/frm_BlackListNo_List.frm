VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_BlackListNo_List 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blacklist No.Telepon"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10080
   Icon            =   "frm_BlackListNo_List.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   10080
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCari 
      Caption         =   "&Cari"
      Height          =   315
      Left            =   4020
      TabIndex        =   8
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox TxtCari 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Top             =   600
      Width           =   2715
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "&Keluar"
      Height          =   495
      Left            =   8580
      TabIndex        =   4
      Top             =   2370
      Width           =   1455
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   8580
      TabIndex        =   3
      Top             =   1770
      Width           =   1455
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   495
      Left            =   8580
      TabIndex        =   2
      Top             =   1170
      Width           =   1455
   End
   Begin VB.CommandButton CmdTambah 
      Caption         =   "&Tambah"
      Height          =   495
      Left            =   8580
      TabIndex        =   1
      Top             =   570
      Width           =   1455
   End
   Begin MSComctlLib.ListView LVBlackList 
      Height          =   4785
      Left            =   0
      TabIndex        =   0
      Top             =   990
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   8440
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Cari No.Telp:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   660
      Width           =   1275
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "BLACKLIST NO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   540
      TabIndex        =   5
      Top             =   60
      Width           =   2325
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   5
      Left            =   30
      Picture         =   "frm_BlackListNo_List.frx":058A
      Stretch         =   -1  'True
      Top             =   30
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   435
      Index           =   8
      Left            =   0
      Picture         =   "frm_BlackListNo_List.frx":1094
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20700
   End
End
Attribute VB_Name = "frm_BlackListNo_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCari_Click()
    Call isi_data
End Sub

Private Sub cmdEdit_Click()
    
    Dim CMDSQL As String
    
    If LVBlackList.ListItems.Count = 0 Then
        Exit Sub
     Else
        With frm_blacklist
            .Caption = "Edit Data No.Telepon BlackList"
            .TxtNoTelp.Text = Trim(LVBlackList.SelectedItem.SubItems(1))
            .TxtKeterangan.Text = LVBlackList.SelectedItem.SubItems(2)
            .txtId.Text = LVBlackList.SelectedItem.SubItems(4)
            .Show vbModal
            If .ok Then
                CMDSQL = "update tblblacklist set no_telp='"
                CMDSQL = CMDSQL + CStr(Trim(.TxtNoTelp.Text)) + "', keterangan='"
                CMDSQL = CMDSQL + IIf(IsNull(.TxtKeterangan.Text), "", Trim(.TxtKeterangan.Text)) + "',tglinput='" & Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & "',userinput='" + MDIForm1.txtnama.Text + "' where id='"
                CMDSQL = CMDSQL + Trim(.txtId.Text) + "'"
                
                M_OBJCONN.Execute CMDSQL
                
                'Update flag di mgm
                Call update_flag_1
                
                LVBlackList.SelectedItem.SubItems(1) = .TxtNoTelp.Text
                LVBlackList.SelectedItem.SubItems(2) = .TxtKeterangan.Text
                LVBlackList.SelectedItem.SubItems(3) = Format(MDIForm1.TDBDate1.Value, "dd/mm/yyyy")
                LVBlackList.SelectedItem.SubItems(5) = MDIForm1.txtnama.Text
            End If
        End With
     End If
End Sub


Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdHapus_Click()
    Dim Cmdsql_Cek As String
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim m_msgbox As String
    
    If LVBlackList.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    m_msgbox = MsgBox("Anda yakin akan menghapus no telepon:" & Trim(LVBlackList.SelectedItem.SubItems(1)), vbYesNo + vbQuestion, "Konfirmasi")
    
    If m_msgbox = vbNo Then
     Exit Sub
    End If
    
    CMDSQL = "delete from tblblacklist where no_telp='"
    CMDSQL = CMDSQL + Trim(LVBlackList.SelectedItem.SubItems(1)) + "'"
    
    M_OBJCONN.Execute CMDSQL
     
    'Update ke flag 0
    Call update_flag_0
    
    LVBlackList.ListItems.Remove LVBlackList.SelectedItem.Index
End Sub

Private Sub cmdkeluar_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdTambah_Click()
Dim noinc As Double
    Dim m_msgbox As Variant
    Dim ListItem As ListItem
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim Cmdsql_Cek As String
    Dim ADD_OK As Boolean
    Dim cmdsql_update
    
    With frm_blacklist
                .Caption = "Tambah Data Black List"
                .Show vbModal
                If .ok Then
                    CMDSQL = "insert into tblblacklist (no_telp,keterangan,tglinput,userinput) values ('"
                    CMDSQL = CMDSQL + Trim(.TxtNoTelp.Text) + "','"
                    CMDSQL = CMDSQL + IIf(IsNull(.TxtKeterangan.Text), "", Trim(.TxtKeterangan.Text)) + "',"
                    CMDSQL = CMDSQL + "'" & Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & "','" + MDIForm1.txtnama.Text + "') "
                    'Cek data no telepon yang sama
                    Set M_objrs = New ADODB.Recordset
                    M_objrs.CursorLocation = adUseClient
                        Cmdsql_Cek = "select * from tblblacklist where no_telp='"
                        Cmdsql_Cek = Cmdsql_Cek + CStr(Trim(.TxtNoTelp.Text)) + "'"
                    M_objrs.Open Cmdsql_Cek, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                    If M_objrs.RecordCount <> 0 Then
                        m_msgbox = MsgBox("No Telepon sudah ada. Data gagal disimpan!", vbOKOnly + vbExclamation, "Peringatan")
                        Exit Sub
                    End If
                    Set M_objrs = Nothing
                    
                    M_OBJCONN.Execute CMDSQL
                    
                    'Update flag ke tabel mgm
                    Call update_flag_1
                    
                    noinc = LVBlackList.ListItems.Count
                    noinc = noinc + 1
                    Set ListItem = LVBlackList.ListItems.ADD(, , CStr(noinc))
                        ListItem.SubItems(1) = .TxtNoTelp.Text
                        ListItem.SubItems(2) = .TxtKeterangan.Text
                        ListItem.SubItems(3) = Format(MDIForm1.TDBDate1.Value, "dd/mm/yyyy")
                        ListItem.SubItems(5) = MDIForm1.txtnama.Text
                        
                        
                End If
    End With
End Sub

Private Sub Form_Load()
    
    Call header_lvblacklist
    Call isi_data
End Sub

Private Sub header_lvblacklist()
    'Membuat Header ListView Program
    LVBlackList.ColumnHeaders.ADD 1, , "No", 10 * 200
    LVBlackList.ColumnHeaders.ADD 2, , "No. Telepon Black List", 10 * 200
    LVBlackList.ColumnHeaders.ADD 3, , "Keterangan", 20 * 100
    LVBlackList.ColumnHeaders.ADD 4, , "Tgl Input", 20 * 100
    LVBlackList.ColumnHeaders.ADD 5, , "Id", 0
    LVBlackList.ColumnHeaders.ADD 6, , "Coding", 20 * 100
End Sub

Private Sub isi_data()
    Dim noinc As Double
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim ListItem As ListItem
    
    CMDSQL = "select * from tblblacklist "
    If txtCari.Text <> Empty Then
        CMDSQL = CMDSQL + " where no_telp like '%"
        CMDSQL = CMDSQL + txtCari.Text + "%' "
    End If
    CMDSQL = CMDSQL + " order by no_telp asc"
    
    LVBlackList.ListItems.CLEAR
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    noinc = 0
    
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    While Not M_objrs.EOF
        noinc = noinc + 1
         Set ListItem = LVBlackList.ListItems.ADD(, , CStr(noinc))
           ListItem.SubItems(1) = IIf(IsNull(M_objrs("no_telp")), "", M_objrs("no_telp"))
           ListItem.SubItems(2) = IIf(IsNull(M_objrs("keterangan")), "", M_objrs("keterangan"))
           ListItem.SubItems(3) = Format(IIf(IsNull(M_objrs("tglinput")), "", M_objrs("tglinput")), "dd/mm/yyyy")
           ListItem.SubItems(4) = IIf(IsNull(M_objrs("id")), "", M_objrs("id"))
           ListItem.SubItems(5) = IIf(IsNull(M_objrs("userinput")), "", M_objrs("userinput"))
         M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
End Sub

Private Sub LVBlackList_DblClick()
    cmdEdit_Click
End Sub

Private Sub update_flag_1()
    Dim cmdsql_homeno As String
    Dim cmdsql_homeno2 As String
    
    Dim cmdsql_mobileno As String
    Dim cmdsql_mobileno2 As String
    
    Dim cmdsql_officeno As String
    Dim cmdsql_officeno2 As String
    
    Dim cmdsql_homenoadd1 As String
    Dim cmdsql_homenoadd2 As String
    
    Dim cmdsql_officenoadd1 As String
    Dim cmdsql_officenoadd2 As String
    
    Dim cmdsql_mobilenoadd1 As String
    Dim cmdsql_mobilenoadd2 As String
    
    Dim cmdsql_ec_telp As String
    
    '@@22062010 Update ke flag di mgm, supaya tanda no merah di agent tidak berat
    'Update flag telepon rumah
    cmdsql_homeno = "update mgm set f_homeno='1' where homeno='"
    cmdsql_homeno = cmdsql_homeno + Trim(frm_blacklist.TxtNoTelp.Text) + "'"
    M_OBJCONN.Execute cmdsql_homeno
    
    cmdsql_homeno2 = "update mgm set f_homeno2='1' where homeno2='"
    cmdsql_homeno2 = cmdsql_homeno2 + Trim(frm_blacklist.TxtNoTelp.Text) + "'"
    M_OBJCONN.Execute cmdsql_homeno2
    
    'Update flag ke telepon hp
    cmdsql_mobileno = "update mgm set f_mobileno='1' where mobileno='"
    cmdsql_mobileno = cmdsql_mobileno + Trim(frm_blacklist.TxtNoTelp.Text) + "'"
    M_OBJCONN.Execute cmdsql_mobileno
    
    cmdsql_mobileno2 = "update mgm set f_mobileno2='1' where mobileno2='"
    cmdsql_mobileno2 = cmdsql_mobileno2 + Trim(frm_blacklist.TxtNoTelp.Text) + "'"
    M_OBJCONN.Execute cmdsql_mobileno2
    
    'Update flag ke telepon office
    cmdsql_officeno = "update mgm set f_officeno='1' where officeno='"
    cmdsql_officeno = cmdsql_officeno + Trim(frm_blacklist.TxtNoTelp.Text) + "'"
    M_OBJCONN.Execute cmdsql_officeno
    
    cmdsql_officeno2 = "update mgm set f_officeno2='1' where officeno2='"
    cmdsql_officeno2 = cmdsql_officeno2 + Trim(frm_blacklist.TxtNoTelp.Text) + "'"
    M_OBJCONN.Execute cmdsql_officeno2
    
    'Update flag ke telepon home add
    cmdsql_homenoadd1 = "update mgm set f_homenoadd1='1' where homenoadd1='"
    cmdsql_homenoadd1 = cmdsql_homenoadd1 + Trim(frm_blacklist.TxtNoTelp.Text) + "'"
    M_OBJCONN.Execute cmdsql_homenoadd1
    
    cmdsql_homenoadd2 = "update mgm set f_homenoadd2='1' where homenoadd2='"
    cmdsql_homenoadd2 = cmdsql_homenoadd2 + Trim(frm_blacklist.TxtNoTelp.Text) + "'"
    M_OBJCONN.Execute cmdsql_homenoadd2
    
    
    'Update flag ke telepon office add
    cmdsql_officenoadd1 = "update mgm set f_officenoadd1='1' where officenoadd1='"
    cmdsql_officenoadd1 = cmdsql_officenoadd1 + Trim(frm_blacklist.TxtNoTelp.Text) + "'"
    M_OBJCONN.Execute cmdsql_officenoadd1
    
    cmdsql_officenoadd2 = "update mgm set f_officenoadd2='1' where officenoadd2='"
    cmdsql_officenoadd2 = cmdsql_officenoadd2 + Trim(frm_blacklist.TxtNoTelp.Text) + "'"
    M_OBJCONN.Execute cmdsql_officenoadd1
    
    'Update flag ke telepon mobileno add
    cmdsql_mobilenoadd1 = "update mgm set f_mobilenoadd1='1' where mobilenoadd1='"
    cmdsql_mobilenoadd1 = cmdsql_mobilenoadd1 + Trim(frm_blacklist.TxtNoTelp.Text) + "'"
    M_OBJCONN.Execute cmdsql_mobilenoadd1
    
    cmdsql_mobilenoadd2 = "update mgm set f_mobilenoadd2='1' where mobilenoadd2='"
    cmdsql_mobilenoadd2 = cmdsql_mobilenoadd2 + Trim(frm_blacklist.TxtNoTelp.Text) + "'"
    M_OBJCONN.Execute cmdsql_mobilenoadd2
    
    'Update flag ke telepon ec_telp
    cmdsql_ec_telp = "update mgm set f_ec_telp='1' where ec_telp='"
    cmdsql_ec_telp = cmdsql_ec_telp + Trim(frm_blacklist.TxtNoTelp.Text) + "'"
    M_OBJCONN.Execute cmdsql_ec_telp
    
End Sub
Private Sub update_flag_0()
    Dim cmdsql_homeno As String
    Dim cmdsql_homeno2 As String
    
    Dim cmdsql_mobileno As String
    Dim cmdsql_mobileno2 As String
    
    Dim cmdsql_officeno As String
    Dim cmdsql_officeno2 As String
    
    Dim cmdsql_homenoadd1 As String
    Dim cmdsql_homenoadd2 As String
    
    Dim cmdsql_officenoadd1 As String
    Dim cmdsql_officenoadd2 As String
    
    Dim cmdsql_mobilenoadd1 As String
    Dim cmdsql_mobilenoadd2 As String
    
    Dim cmdsql_ec_telp As String
    
    '@@22062010 Update ke flag di mgm, supaya tanda no merah di agent tidak berat
    'Update flag telepon rumah
    cmdsql_homeno = "update mgm set f_homeno='0' where homeno='"
    cmdsql_homeno = cmdsql_homeno + Trim(LVBlackList.SelectedItem.SubItems(1)) + "'"
    M_OBJCONN.Execute cmdsql_homeno
    
    cmdsql_homeno2 = "update mgm set f_homeno2='0' where homeno2='"
    cmdsql_homeno2 = cmdsql_homeno2 + Trim(LVBlackList.SelectedItem.SubItems(1)) + "'"
    M_OBJCONN.Execute cmdsql_homeno2
    
    'Update flag ke telepon hp
    cmdsql_mobileno = "update mgm set f_mobileno='0' where mobileno='"
    cmdsql_mobileno = cmdsql_mobileno + Trim(LVBlackList.SelectedItem.SubItems(1)) + "'"
    M_OBJCONN.Execute cmdsql_mobileno
    
    cmdsql_mobileno2 = "update mgm set f_mobileno2='0' where mobileno2='"
    cmdsql_mobileno2 = cmdsql_mobileno2 + Trim(LVBlackList.SelectedItem.SubItems(1)) + "'"
    M_OBJCONN.Execute cmdsql_mobileno2
    
    'Update flag ke telepon office
    cmdsql_officeno = "update mgm set f_officeno='0' where officeno='"
    cmdsql_officeno = cmdsql_officeno + Trim(LVBlackList.SelectedItem.SubItems(1)) + "'"
    M_OBJCONN.Execute cmdsql_officeno
    
    cmdsql_officeno2 = "update mgm set f_officeno2='0' where officeno2='"
    cmdsql_officeno2 = cmdsql_officeno2 + Trim(LVBlackList.SelectedItem.SubItems(1)) + "'"
    M_OBJCONN.Execute cmdsql_officeno2
    
    'Update flag ke telepon home add
    cmdsql_homenoadd1 = "update mgm set f_homenoadd1='0' where homenoadd1='"
    cmdsql_homenoadd1 = cmdsql_homenoadd1 + Trim(LVBlackList.SelectedItem.SubItems(1)) + "'"
    M_OBJCONN.Execute cmdsql_homenoadd1
    
    cmdsql_homenoadd2 = "update mgm set f_homenoadd2='0' where homenoadd2='"
    cmdsql_homenoadd2 = cmdsql_homenoadd2 + Trim(LVBlackList.SelectedItem.SubItems(1)) + "'"
    M_OBJCONN.Execute cmdsql_homenoadd2
    
    
    'Update flag ke telepon office add
    cmdsql_officenoadd1 = "update mgm set f_officenoadd1='0' where officenoadd1='"
    cmdsql_officenoadd1 = cmdsql_officenoadd1 + Trim(LVBlackList.SelectedItem.SubItems(1)) + "'"
    M_OBJCONN.Execute cmdsql_officenoadd1
    
    cmdsql_officenoadd2 = "update mgm set f_officenoadd2='0' where officenoadd2='"
    cmdsql_officenoadd2 = cmdsql_officenoadd2 + Trim(LVBlackList.SelectedItem.SubItems(1)) + "'"
    M_OBJCONN.Execute cmdsql_officenoadd1
    
    'Update flag ke telepon mobileno add
    cmdsql_mobilenoadd1 = "update mgm set f_mobilenoadd1='0' where mobilenoadd1='"
    cmdsql_mobilenoadd1 = cmdsql_mobilenoadd1 + Trim(LVBlackList.SelectedItem.SubItems(1)) + "'"
    M_OBJCONN.Execute cmdsql_mobilenoadd1
    
    cmdsql_mobilenoadd2 = "update mgm set f_mobilenoadd2='0' where mobilenoadd2='"
    cmdsql_mobilenoadd2 = cmdsql_mobilenoadd2 + Trim(LVBlackList.SelectedItem.SubItems(1)) + "'"
    M_OBJCONN.Execute cmdsql_mobilenoadd2
    
    'Update flag ke telepon ec_telp
    cmdsql_ec_telp = "update mgm set f_ec_telp='0' where ec_telp='"
    cmdsql_ec_telp = cmdsql_ec_telp + Trim(LVBlackList.SelectedItem.SubItems(1)) + "'"
    M_OBJCONN.Execute cmdsql_ec_telp

End Sub

