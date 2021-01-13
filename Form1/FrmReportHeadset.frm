VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmReportHeadset 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Report Headset"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6210
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdUnCekAll 
      Caption         =   "Uncek..."
      Height          =   315
      Left            =   2460
      TabIndex        =   8
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton CmdCekAll 
      Caption         =   "CekAll"
      Height          =   315
      Left            =   1500
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   435
      Left            =   4680
      TabIndex        =   5
      Top             =   5280
      Width           =   1395
   End
   Begin VB.CommandButton CmdReport 
      Caption         =   "&Report"
      Height          =   435
      Left            =   3300
      TabIndex        =   4
      Top             =   5280
      Width           =   1395
   End
   Begin VB.TextBox TxtKeterangan 
      Height          =   1215
      Left            =   1500
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3900
      Width           =   4575
   End
   Begin MSComctlLib.ListView LvHeadset 
      Height          =   2880
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   5080
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
   Begin VB.Label Label3 
      Caption         =   "Keterangan:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   3900
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Jenis Kerusakan:"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   540
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "Report Headset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6195
   End
End
Attribute VB_Name = "FrmReportHeadset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public JenisKerusakan As String
Private Sub IsiMasalah()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim ListItem As ListItem
    
    CMDSQL = "select * from mandiri.tbl_jenis_masalah where jenis_problem='HEADSET' "
    CMDSQL = CMDSQL + " and status='1' and nama_problem is not null "
    CMDSQL = CMDSQL + " order by nama_problem asc "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvHeadset.ListItems.CLEAR
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            Set ListItem = LvHeadset.ListItems.ADD(, , M_objrs("id"))
                ListItem.SubItems(1) = IIf(IsNull(M_objrs("nama_problem")), "", M_objrs("nama_problem"))
            M_objrs.MoveNext
        Wend
    Else
        MsgBox "Data problem kosong!", vbOKOnly + vbInformation, "Informasi"
        Set M_objrs = Nothing
        Unload Me
    End If
    
    Set M_objrs = Nothing
End Sub

Private Sub HeaderHeadset()
    LvHeadset.ColumnHeaders.ADD 1, , "ID", 1000
    LvHeadset.ColumnHeaders.ADD 2, , "NAMA PROBLEM", 5000
End Sub

Private Sub CmdBatal_Click()
    Unload Me
End Sub

Private Sub CmdCekAll_Click()
    Dim K As Integer
    
    If LvHeadset.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    For K = 1 To LvHeadset.ListItems.Count
        LvHeadset.ListItems(K).Checked = True
    Next K
End Sub

Private Sub CmdReport_Click()
    Dim W As Integer
    Dim CMDSQL As String
    Dim STRSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim Remarks As String
        
    On Error GoTo Salah
    JenisKerusakan = ""
    Remarks = ""
    
    For W = 1 To LvHeadset.ListItems.Count
        If LvHeadset.ListItems(W).Checked = True Then
            If JenisKerusakan = "" Then
                JenisKerusakan = LvHeadset.ListItems(W).SubItems(1)
            Else
                JenisKerusakan = JenisKerusakan & "," & LvHeadset.ListItems(W).SubItems(1)
            End If
        End If
    Next W
    
    If JenisKerusakan = "" Then
        MsgBox "Anda belum memilih jenis kerusakan!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    
    CMDSQL = "insert into tbl_problem_headset (userid,nama,tgl_pengajuan,jenis_kerusakan,keterangan) "
    CMDSQL = CMDSQL + " values ('"
    CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "','"
    CMDSQL = CMDSQL + MDIForm1.txtnama.Text + "',now(),'"
    CMDSQL = CMDSQL + JenisKerusakan + "','"
    CMDSQL = CMDSQL + IIf(IsNull(txtketerangan.Text), "", txtketerangan.Text) + "')"
    M_OBJCONN.Execute CMDSQL
    
    Remarks = "Pesan Create By System: " & Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & vbCrLf
    Remarks = Remarks & "--------------------------------------- " & vbCrLf
    Remarks = Remarks & " AGENT: " & UCase(MDIForm1.TxtUsername.Text) & vbCrLf
    Remarks = Remarks & " NAMA: " & UCase(MDIForm1.txtnama.Text) & vbCrLf & vbCrLf
    Remarks = Remarks & " Telah melakukan reporting masalah headset, sebagai berikut: " & vbCrLf
    Remarks = Remarks & UCase(JenisKerusakan) & vbCrLf & vbCrLf
    Remarks = Remarks & IIf(IsNull(txtketerangan.Text), "", txtketerangan.Text)
    
    
    'Kirim pesan ke TL nya
    If UseridTL <> "" Then
        STRSQL = "select * from usertbl where userid='"
        STRSQL = STRSQL + UseridTL + "' and sts_kirim_pesan_error='1' "
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_objrs.RecordCount > 0 Then
            CMDSQL = "insert into msgtbl "
            CMDSQL = CMDSQL + "( recipient, datetime, sender, sentfrom, msg) values ('"
            CMDSQL = CMDSQL + UseridTL + "','"
            CMDSQL = CMDSQL + Format(Now(), "yyyymmdd") + "','"
            CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "','"
            CMDSQL = CMDSQL + CStr(MDIForm1.Winsock1.LocalIP) + "','"
            CMDSQL = CMDSQL + Remarks + "')"
            M_OBJCONN.Execute CMDSQL
        End If
        Set M_objrs = Nothing
    End If
    
    'Kirim ke usertype lainnya selain TL
    STRSQL = "select * from usertbl where sts_kirim_pesan_error='1' and usertype<>'6' "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            CMDSQL = "insert into msgtbl "
            CMDSQL = CMDSQL + "( recipient, datetime, sender, sentfrom, msg) values ('"
            CMDSQL = CMDSQL + M_objrs("userid") + "','"
            CMDSQL = CMDSQL + Format(Now(), "yyyymmdd") + "','"
            CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "','"
            CMDSQL = CMDSQL + CStr(MDIForm1.Winsock1.LocalIP) + "','"
            CMDSQL = CMDSQL + Remarks + "')"
            M_OBJCONN.Execute CMDSQL
            M_objrs.MoveNext
        Wend
    End If
    
    Set M_objrs = Nothing
   
   MsgBox "Report Headset anda telah terkirim!", vbOKOnly + vbInformation, "Informasi"
   Unload Me
   Exit Sub
Salah:
   MsgBox "Kami mohon maaf, ada error:" & err.Description, vbOKOnly + vbInformation, "Informasi"
    
End Sub

Private Sub CmdUnCekAll_Click()
    Dim K As Integer
    
    If LvHeadset.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    For K = 1 To LvHeadset.ListItems.Count
        LvHeadset.ListItems(K).Checked = False
    Next K
End Sub

Private Sub Form_Load()
    Call HeaderHeadset
    Call IsiMasalah
End Sub
