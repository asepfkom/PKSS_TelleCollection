VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAksesLayananTelkom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Akses Layanan Telkom"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4275
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   4275
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCek 
      Caption         =   "CEK"
      Height          =   315
      Left            =   1860
      TabIndex        =   7
      Top             =   4680
      Width           =   1155
   End
   Begin VB.CommandButton CmdUncek 
      Caption         =   "UNCEK"
      Height          =   315
      Left            =   3000
      TabIndex        =   6
      Top             =   4680
      Width           =   1155
   End
   Begin VB.ComboBox CmbPilihGroup 
      Height          =   315
      ItemData        =   "FrmAksesLayananTelkom.frx":0000
      Left            =   180
      List            =   "FrmAksesLayananTelkom.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   4740
      Width           =   1635
   End
   Begin VB.CommandButton CmdProses 
      Caption         =   "&Proses"
      Height          =   315
      Left            =   180
      TabIndex        =   4
      Top             =   5280
      Width           =   1995
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Keluar"
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Top             =   5280
      Width           =   1995
   End
   Begin VB.Frame Frame1 
      Height          =   2355
      Left            =   180
      TabIndex        =   0
      Top             =   5700
      Width           =   3975
      Begin VB.Label Label1 
         Caption         =   "Agent yang DICENTANG adalah, agent yang dapat menggunakan layanan 108 ."
         Height          =   435
         Left            =   180
         TabIndex        =   2
         Top             =   600
         Width           =   3675
      End
      Begin VB.Label Label2 
         Caption         =   "Ageng yang TIDAK DICENTANG adalah agent yang tidak dapat mengakses layanan 108"
         Height          =   495
         Left            =   180
         TabIndex        =   1
         Top             =   1140
         Width           =   3675
      End
   End
   Begin MSComctlLib.ListView LvUser 
      Height          =   4335
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   7646
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "FrmAksesLayananTelkom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub IsiCombo()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    
    CmbPilihGroup.CLEAR
    If UCase(MDIForm1.txtlevel.Text) = "TEAMLEADER" Then
        CmbPilihGroup.AddItem Trim(MDIForm1.TxtUsername.Text)
    Else
        CmbPilihGroup.AddItem "ALL"
        CMDSQL = "select team from usertbl where usertype='6' order by team asc"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_objrs.RecordCount > 0 Then
            While Not M_objrs.EOF
                CmbPilihGroup.AddItem IIf(IsNull(M_objrs("team")), "", M_objrs("team"))
                M_objrs.MoveNext
            Wend
        End If
        
        Set M_objrs = Nothing
    End If
    
End Sub

Private Sub header()
    LvUser.ColumnHeaders.ADD 1, , "Userid", 1500
    LvUser.ColumnHeaders.ADD 2, , "Nama", 5000
    LvUser.ColumnHeaders.ADD 3, , "Team", 4000
End Sub

Private Sub IsiData()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim listitem As listitem
    Dim a As Integer
    
    If UCase(MDIForm1.txtlevel.Text) = "TEAMLEADER" Then
        CMDSQL = " select * from  usertbl where "
        CMDSQL = CMDSQL + " team='"
        CMDSQL = CMDSQL + Trim(MDIForm1.TxtUsername.Text) + "' and usertype='1' order by userid"
        'team,userid asc "
    Else
        CMDSQL = " select * from  usertbl where usertype='1' order by userid"
        'team,userid asc "
    End If
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            Set listitem = LvUser.ListItems.ADD(, , M_objrs("userid"))
                listitem.SubItems(1) = IIf(IsNull(M_objrs("agent")), "", M_objrs("agent"))
                listitem.SubItems(2) = IIf(IsNull(M_objrs("team")), "", M_objrs("team"))
                
                If M_objrs("sts_108") = "1" Then
                    listitem.Checked = True
                End If
            M_objrs.MoveNext
        Wend
        
        
    End If
    
    Set M_objrs = Nothing
End Sub

Private Sub Check1_Click()

End Sub

Private Sub CmdCek_Click()
    Dim W As Integer
    
    
    If LvUser.ListItems.Count = 0 Then
        MsgBox "Tidak ada data user!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If CmbPilihGroup.Text = "" Then
        MsgBox "Pilih kriteria data yang akan diceklist!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    'Jika pilihan= ALL
    If Trim(CmbPilihGroup.Text) = "ALL" Then
        For W = 1 To LvUser.ListItems.Count
            LvUser.ListItems(W).Checked = True
        Next W
    Else
        For W = 1 To LvUser.ListItems.Count
            If Trim(LvUser.ListItems(W).SubItems(2)) = Trim(CmbPilihGroup.Text) Then
                LvUser.ListItems(W).Checked = True
            End If
        Next W
    End If
    
    MsgBox "Data berhasil di ceklist!", vbOKOnly + vbInformation, "Informasi"
End Sub

Private Sub cmdproses_Click()
    Dim K As Integer
    Dim CMDSQL As String
    Dim a As String
    Dim Remarks As String
    
    
    If LvUser.ListItems.Count = 0 Then
        MsgBox "Data user tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    a = MsgBox("Apakah anda yakin akan memproses data ini?", vbYesNo + vbQuestion, "Konfirmasi")
    
    If a = vbNo Then
        MsgBox "Proses dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    
    For K = 1 To LvUser.ListItems.Count
    
        'Jika dicentang maka AGENT dapat mengakses 108
        If LvUser.ListItems(K).Checked = True Then
            CMDSQL = "update usertbl set sts_108='1' where userid='"
            CMDSQL = CMDSQL + Trim(LvUser.ListItems(K).Text) + "'"
            M_OBJCONN.Execute CMDSQL
            
            'Informasikan ke agent melalui pesan
            Remarks = "Informasi : " + vbCrLf
            Remarks = Remarks + "---------------------------------------" + vbCrLf
            Remarks = Remarks + "Anda telah diberi akses dapat menelepon 108" + vbCrLf
            
            
            
            CMDSQL = "insert into msgtbl "
            CMDSQL = CMDSQL + "( recipient, datetime, sender, sentfrom, msg) values ('"
            CMDSQL = CMDSQL + Trim(LvUser.ListItems(K).Text) + "','"
            CMDSQL = CMDSQL + Format(Now(), "yyyymmdd") + "','"
            CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "','"
            CMDSQL = CMDSQL + CStr(MDIForm1.Winsock1.LocalIP) + "','"
            CMDSQL = CMDSQL + Remarks + "')"
            
            M_OBJCONN.Execute CMDSQL
        End If
        
        'Jika tidak dicentang, maka tidak dapat mengakses layanan 108
        If LvUser.ListItems(K).Checked = False Then
            CMDSQL = "update usertbl set sts_108=null where userid='"
            CMDSQL = CMDSQL + Trim(LvUser.ListItems(K).Text) + "'"
            M_OBJCONN.Execute CMDSQL
            
             'Informasikan ke agent melalui pesan
            Remarks = "Informasi : " + vbCrLf
            Remarks = Remarks + "---------------------------------------" + vbCrLf
            Remarks = Remarks + "Hak Akses 108 anda dihentikan!" + vbCrLf
            
            
            
            CMDSQL = "insert into msgtbl "
            CMDSQL = CMDSQL + "( recipient, datetime, sender, sentfrom, msg) values ('"
            CMDSQL = CMDSQL + Trim(LvUser.ListItems(K).Text) + "','"
            CMDSQL = CMDSQL + Format(Now(), "yyyymmdd") + "','"
            CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "','"
            CMDSQL = CMDSQL + CStr(MDIForm1.Winsock1.LocalIP) + "','"
            CMDSQL = CMDSQL + Remarks + "')"
            
            M_OBJCONN.Execute CMDSQL
        End If
        
    Next K
    
    MsgBox "Data berhasil di proses!", vbOKOnly + vbInformation, "Informasi"
    
End Sub

Private Sub CmdUncek_Click()
    Dim W As Integer
    
    
    If LvUser.ListItems.Count = 0 Then
        MsgBox "Tidak ada data user!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If CmbPilihGroup.Text = "" Then
        MsgBox "Pilih kriteria data yang akan diceklist!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    'Jika pilihan= ALL
    If Trim(CmbPilihGroup.Text) = "ALL" Then
        For W = 1 To LvUser.ListItems.Count
            LvUser.ListItems(W).Checked = False
        Next W
    Else
        For W = 1 To LvUser.ListItems.Count
            If Trim(LvUser.ListItems(W).SubItems(2)) = Trim(CmbPilihGroup.Text) Then
                LvUser.ListItems(W).Checked = False
            End If
        Next W
    End If
    
    MsgBox "Data berhasil di uncek!", vbOKOnly + vbInformation, "Informasi"
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call header
    Call IsiCombo
    Call IsiData
End Sub

Private Sub LvUser_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   LvUser.SortKey = ColumnHeader.Index - 1
   IndexColumnHEader = ColumnHeader.Index - 1
   LvUser.Sorted = True
End Sub
