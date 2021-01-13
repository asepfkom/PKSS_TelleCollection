VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmListReportTelepon 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "List Report Telepon"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   12585
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   12585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCekAll 
      Caption         =   "&Cek All"
      Height          =   435
      Left            =   60
      TabIndex        =   5
      Top             =   7740
      Width           =   1155
   End
   Begin VB.CommandButton CmdUncekAll 
      Caption         =   "&UnCekAll"
      Height          =   435
      Left            =   1260
      TabIndex        =   4
      Top             =   7740
      Width           =   1155
   End
   Begin VB.CommandButton CmdFollowUp 
      Caption         =   "&Follow up"
      Height          =   435
      Left            =   3900
      TabIndex        =   3
      Top             =   7740
      Width           =   1155
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus"
      Height          =   435
      Left            =   2580
      TabIndex        =   2
      Top             =   7740
      Width           =   1155
   End
   Begin VB.CommandButton CmdLoadData 
      Caption         =   "&Load data"
      Height          =   435
      Left            =   5040
      TabIndex        =   0
      Top             =   7740
      Width           =   1155
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   315
      Left            =   8820
      TabIndex        =   1
      Top             =   7800
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ListView LvListProblemTelepon 
      Height          =   7620
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12540
      _ExtentX        =   22119
      _ExtentY        =   13441
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
End
Attribute VB_Name = "FrmListReportTelepon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub HeaderList()
    LvListProblemTelepon.ColumnHeaders.ADD 1, , "ID", 900
    LvListProblemTelepon.ColumnHeaders.ADD 2, , "Status", 1200
    LvListProblemTelepon.ColumnHeaders.ADD 3, , "Tgl.Pengajuan", 1500
    LvListProblemTelepon.ColumnHeaders.ADD 4, , "Userid", 1000
    LvListProblemTelepon.ColumnHeaders.ADD 5, , "Nama", 2000
    LvListProblemTelepon.ColumnHeaders.ADD 6, , "Telepon Masalah", 2000
    LvListProblemTelepon.ColumnHeaders.ADD 7, , "Jenis Kerusakan", 5000
    LvListProblemTelepon.ColumnHeaders.ADD 8, , "Keterangan", 4500
    
    '@@18012012 Tambahan
    LvListProblemTelepon.ColumnHeaders.ADD 9, , "Tanggal Solusi", 1500
    LvListProblemTelepon.ColumnHeaders.ADD 10, , "Solusi Oleh", 1500
    LvListProblemTelepon.ColumnHeaders.ADD 11, , "Keterangan", 1500
    LvListProblemTelepon.ColumnHeaders.ADD 12, , "Jenis Telepon", 1500
End Sub


Private Sub CmdCekAll_Click()
    Dim W As Integer
    
    If LvListProblemTelepon.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LvListProblemTelepon.ListItems.Count
        LvListProblemTelepon.ListItems(W).Checked = True
    Next W
End Sub

Private Sub CmdFollowUp_Click()
    If LvListProblemTelepon.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If UCase(LvListProblemTelepon.SelectedItem.SubItems(1)) = "FIXED" Then
        MsgBox "Masalah sudah fix! tidak dapat di edit lagi!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    With FrmFollowUpProblemTelepon
        .txtId.Text = LvListProblemTelepon.SelectedItem.Text
        .TxtTglPengajuan.Text = LvListProblemTelepon.SelectedItem.SubItems(2)
        .txtuserid.Text = LvListProblemTelepon.SelectedItem.SubItems(3)
        .txtnama.Text = LvListProblemTelepon.SelectedItem.SubItems(4)
        .TxtNoTelp.Text = LvListProblemTelepon.SelectedItem.SubItems(5)
        .TxtJenisKerusakan.Text = LvListProblemTelepon.SelectedItem.SubItems(6)
        .txtketerangan.Text = IIf(IsNull(LvListProblemTelepon.SelectedItem.SubItems(7)), "", LvListProblemTelepon.SelectedItem.SubItems(7))
        
        .TxtTglSolusi.Value = IIf(IsNull(LvListProblemTelepon.SelectedItem.SubItems(8)), Format(Now, "dd/mm/yyyy"), Format(LvListProblemTelepon.SelectedItem.SubItems(8), "dd/mm/yyyy"))
        .TxtSolusiOleh.Text = IIf(IsNull(LvListProblemTelepon.SelectedItem.SubItems(9)), "", LvListProblemTelepon.SelectedItem.SubItems(9))
        .TxtKetSolusi.Text = IIf(IsNull(LvListProblemTelepon.SelectedItem.SubItems(10)), "", LvListProblemTelepon.SelectedItem.SubItems(10))
        .CmbStatusSolusi.Text = IIf(UCase(LvListProblemTelepon.SelectedItem.SubItems(1)) = "NOT FOLLOW UP", "Follow Up", LvListProblemTelepon.SelectedItem.SubItems(1))
        .Show vbModal
    End With
        
End Sub

Private Sub cmdHapus_Click()
    Dim CMDSQL As String
    Dim a As String
    Dim W As Integer
    
    If LvListProblemTelepon.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    a = MsgBox("Apakah anda yakin akan menghapus data?", vbYesNo + vbQuestion, "Konfirmasi")
    
    If a = vbNo Then
        Exit Sub
    End If
    
    If a = vbYes Then
        For W = 1 To LvListProblemTelepon.ListItems.Count
            If LvListProblemTelepon.ListItems(W).Checked = True Then
                CMDSQL = "delete from mandiri.tbl_problem_telepon where id='"
                CMDSQL = CMDSQL + CStr(LvListProblemTelepon.ListItems(W).Text) + "'"
                M_OBJCONN.Execute CMDSQL
            End If
        Next W
    End If
    
    MsgBox "Data berhasil dihapus!", vbOKOnly + vbInformation, "Infromasi"
    Call IsiData
End Sub

Private Sub CmdLoadData_Click()
    Call IsiData
End Sub

Private Sub CmdUnCekAll_Click()
    Dim W As Integer
    
    If LvListProblemTelepon.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LvListProblemTelepon.ListItems.Count
        LvListProblemTelepon.ListItems(W).Checked = False
    Next W
End Sub

Private Sub Form_Load()
    Call HeaderList
End Sub

Public Sub IsiData()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim ListItem As ListItem
    Dim K As Integer
    
    CMDSQL = "select * from mandiri.tbl_problem_telepon order by status_solusi,tgl_pengajuan asc"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvListProblemTelepon.ListItems.CLEAR
    
    If M_objrs.RecordCount > 0 Then
        PB1.Max = M_objrs.RecordCount
        While Not M_objrs.EOF
            PB1.Value = M_objrs.Bookmark
            Set ListItem = LvListProblemTelepon.ListItems.ADD(, , M_objrs("id"))
                ListItem.SubItems(1) = IIf(IsNull(M_objrs("status_solusi")), "NOT FOLLOW UP", M_objrs("status_solusi"))
                ListItem.SubItems(2) = Format(M_objrs("tgl_pengajuan"), "yyyy-mm-dd hh:nn:ss")
                ListItem.SubItems(3) = M_objrs("userid")
                ListItem.SubItems(4) = M_objrs("nama")
                ListItem.SubItems(5) = M_objrs("no_telepon")
                ListItem.SubItems(6) = M_objrs("jenis_kerusakan")
                ListItem.SubItems(7) = IIf(IsNull(M_objrs("keterangan")), "", M_objrs("keterangan"))
                
                '@@18012013 Tambahan
                ListItem.SubItems(8) = IIf(IsNull(M_objrs("tgl_solusi")), "", Format(M_objrs("tgl_solusi"), "yyyy-mm-dd"))
                ListItem.SubItems(9) = IIf(IsNull(M_objrs("solusi_by")), "", M_objrs("solusi_by"))
                
                ListItem.SubItems(10) = IIf(IsNull(M_objrs("solusi")), "", M_objrs("solusi"))
                ListItem.SubItems(11) = IIf(IsNull(M_objrs("jenis_telepon")), "", M_objrs("jenis_telepon"))
                
                K = 1
                
                If IsNull(M_objrs("status_solusi")) = True Or M_objrs("status_solusi") = "" Then
                     LvListProblemTelepon.ForeColor = vbRed
                     For K = 1 To 11
                        ListItem.ListSubItems(K).ForeColor = vbRed
                     Next K
                End If
                
                If UCase(M_objrs("status_solusi")) = "FOLLOW UP" Then
                     LvListProblemTelepon.ForeColor = vbYellow
                     For K = 1 To 11
                        ListItem.ListSubItems(K).ForeColor = vbYellow
                     Next K
                End If
                
                If UCase(M_objrs("status_solusi")) = "FIXED" Then
                     LvListProblemTelepon.ForeColor = vbGreen
                     For K = 1 To 11
                        ListItem.ListSubItems(K).ForeColor = vbGreen
                     Next K
                End If
                
            M_objrs.MoveNext
        Wend
    Else
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
    End If
    
    Set M_objrs = Nothing
End Sub


Private Sub LvListProblemTelepon_DblClick()
    CmdFollowUp_Click
End Sub

