VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmReportTelepon 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Report Telepon"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6195
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CmbNoTelp 
      Height          =   315
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   3900
      Width           =   3435
   End
   Begin VB.TextBox TxtKeterangan 
      Height          =   1215
      Left            =   1860
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   4260
      Width           =   4215
   End
   Begin VB.CommandButton CmdReport 
      Caption         =   "&Report"
      Height          =   435
      Left            =   3300
      TabIndex        =   3
      Top             =   5700
      Width           =   1395
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   435
      Left            =   4680
      TabIndex        =   2
      Top             =   5700
      Width           =   1395
   End
   Begin VB.CommandButton CmdCekAll 
      Caption         =   "CekAll"
      Height          =   315
      Left            =   1500
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton CmdUnCekAll 
      Caption         =   "Uncek..."
      Height          =   315
      Left            =   2460
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin MSComctlLib.ListView LvTelepon 
      Height          =   2880
      Left            =   0
      TabIndex        =   5
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
   Begin VB.Label LblTelp 
      Height          =   255
      Left            =   60
      TabIndex        =   11
      Top             =   6000
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Label Label4 
      Caption         =   "No.Telepon masalah:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Caption         =   "Report Telepon"
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
      TabIndex        =   8
      Top             =   0
      Width           =   6195
   End
   Begin VB.Label Label2 
      Caption         =   "Jenis Kerusakan:"
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   540
      Width           =   1275
   End
   Begin VB.Label Label3 
      Caption         =   "Keterangan:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   4260
      Width           =   1335
   End
End
Attribute VB_Name = "FrmReportTelepon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public JenisKerusakan As String
Private Sub IsiMasalah()
    Dim cmdsql As String
    Dim M_objrs As ADODB.Recordset
    Dim ListItem As ListItem
    
    cmdsql = "select * from tbl_jenis_masalah where jenis_problem='TELEPON' "
    cmdsql = cmdsql + " and status='1' and nama_problem is not null "
    cmdsql = cmdsql + " order by nama_problem asc "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvTelepon.ListItems.CLEAR
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            Set ListItem = LvTelepon.ListItems.ADD(, , M_objrs("id"))
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

Private Sub HeaderTelepon()
    LvTelepon.ColumnHeaders.ADD 1, , "ID", 1000
    LvTelepon.ColumnHeaders.ADD 2, , "NAMA PROBLEM", 5000
End Sub



Private Sub CmbNoTelp_Click()
    Select Case UCase(CmbNoTelp.Text)
        Case "HOME 1"
            LblTelp.Caption = Trim(FrmCC_Colection.txtHomeNo1.Text)
        Case "HOME 2"
            LblTelp.Caption = Trim(FrmCC_Colection.txtHomeNo2.Text)
        Case "OFFICE 1"
            LblTelp.Caption = Trim(FrmCC_Colection.txtOfficeNo1.Text)
        Case "OFFICE 2"
            LblTelp.Caption = Trim(FrmCC_Colection.txtOfficeNo2.Text)
        Case "MOBILE 1"
            LblTelp.Caption = Trim(FrmCC_Colection.txtMobileNo1.Text)
        Case "MOBILE 2"
            LblTelp.Caption = Trim(FrmCC_Colection.txtMobileNo2.Text)
        Case "EC"
            'LblTelp.Caption = Trim(FrmCC_Colection.txtECno.Text)
        Case "ADD HOME 1"
            LblTelp.Caption = Trim(FrmCC_Colection.txtHomeAdd1.Text)
        Case "ADD HOME 2"
            LblTelp.Caption = Trim(FrmCC_Colection.txtHomeAdd2.Text)
        Case "ADD OFFICE 1"
            LblTelp.Caption = Trim(FrmCC_Colection.txtOfficeAdd1.Text)
        Case "ADD OFFICE 2"
            LblTelp.Caption = Trim(FrmCC_Colection.txtOfficeAdd2.Text)
        Case "ADD MOBILE 1"
            LblTelp.Caption = Trim(FrmCC_Colection.txtMobileAdd1.Text)
        Case "ADD MOBILE 2"
            LblTelp.Caption = Trim(FrmCC_Colection.txtMobileAdd2.Text)
    End Select
End Sub

Private Sub CmdBatal_Click()
    Unload Me
End Sub

Private Sub CmdCekAll_Click()
    Dim K As Integer
    
    If LvTelepon.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    For K = 1 To LvTelepon.ListItems.Count
        LvTelepon.ListItems(K).Checked = True
    Next K
End Sub

Private Sub CmdReport_Click()
    Dim W As Integer
    Dim cmdsql As String
    Dim strsql As String
    Dim M_objrs As ADODB.Recordset
    Dim Remarks As String
        
    On Error GoTo salah
    JenisKerusakan = ""
    Remarks = ""
    
    If LblTelp.Caption = "" Or IsNull(LblTelp.Caption) = True Then
        MsgBox "Anda harus mengisi no telepon yang bermasalah!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LvTelepon.ListItems.Count
        If LvTelepon.ListItems(W).Checked = True Then
            If JenisKerusakan = "" Then
                JenisKerusakan = LvTelepon.ListItems(W).SubItems(1)
            Else
                JenisKerusakan = JenisKerusakan & "," & LvTelepon.ListItems(W).SubItems(1)
            End If
        End If
    Next W
    
    If JenisKerusakan = "" Then
        MsgBox "Anda belum memilih jenis kerusakan!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    
    cmdsql = "insert into tbl_problem_telepon (userid,nama,tgl_pengajuan,jenis_kerusakan,"
    cmdsql = cmdsql + "keterangan,no_telepon,jenis_telepon) "
    cmdsql = cmdsql + " values ('"
    cmdsql = cmdsql + MDIForm1.TxtUsername.Text + "','"
    cmdsql = cmdsql + MDIForm1.txtnama.Text + "',now(),'"
    cmdsql = cmdsql + JenisKerusakan + "','"
    cmdsql = cmdsql + IIf(IsNull(txtketerangan.Text), "", txtketerangan.Text) + "','"
    cmdsql = cmdsql + CStr(LblTelp.Caption) + "','"
    cmdsql = cmdsql + CmbNoTelp.Text + "')"
    
    M_OBJCONN.Execute cmdsql
    
    Remarks = "Pesan Create By System: " & Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & vbCrLf
    Remarks = Remarks & "--------------------------------------- " & vbCrLf
    Remarks = Remarks & " AGENT: " & UCase(MDIForm1.TxtUsername.Text) & vbCrLf
    Remarks = Remarks & " NAMA: " & UCase(MDIForm1.txtnama.Text) & vbCrLf & vbCrLf
    Remarks = Remarks & " Telah melakukan reporting MASALAH TELEPON, sebagai berikut: " & vbCrLf
    Remarks = Remarks & UCase(JenisKerusakan) & vbCrLf & vbCrLf
    Remarks = Remarks & IIf(IsNull(txtketerangan.Text), "", txtketerangan.Text)
    
    
    'Kirim pesan ke TL nya
    If UseridTL <> "" Then
        strsql = "select * from usertbl where userid='"
        strsql = strsql + UseridTL + "' and sts_kirim_pesan_error_telp='1' "
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_objrs.RecordCount > 0 Then
            cmdsql = "insert into msgtbl "
            cmdsql = cmdsql + "( recipient, datetime, sender, sentfrom, msg) values ('"
            cmdsql = cmdsql + UseridTL + "','"
            cmdsql = cmdsql + Format(Now(), "yyyymmdd") + "','"
            cmdsql = cmdsql + MDIForm1.TxtUsername.Text + "','"
            cmdsql = cmdsql + CStr(MDIForm1.Winsock1.LocalIP) + "','"
            cmdsql = cmdsql + Remarks + "')"
            M_OBJCONN.Execute cmdsql
        End If
        Set M_objrs = Nothing
    End If
    
    'Kirim ke usertype lainnya selain TL
    strsql = "select * from usertbl where sts_kirim_pesan_error_telp='1' and usertype<>'6' "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            cmdsql = "insert into msgtbl "
            cmdsql = cmdsql + "( recipient, datetime, sender, sentfrom, msg) values ('"
            cmdsql = cmdsql + M_objrs("userid") + "','"
            cmdsql = cmdsql + Format(Now(), "yyyymmdd") + "','"
            cmdsql = cmdsql + MDIForm1.TxtUsername.Text + "','"
            cmdsql = cmdsql + CStr(MDIForm1.Winsock1.LocalIP) + "','"
            cmdsql = cmdsql + Remarks + "')"
            M_OBJCONN.Execute cmdsql
            M_objrs.MoveNext
        Wend
    End If
    
    Set M_objrs = Nothing
   
   MsgBox "Report Headset anda telah terkirim!", vbOKOnly + vbInformation, "Informasi"
   Unload Me
   Exit Sub
salah:
   MsgBox "Kami mohon maaf, ada error:" & err.Description, vbOKOnly + vbInformation, "Informasi"
    
End Sub

Private Sub CmdUnCekAll_Click()
    Dim K As Integer
    
    If LvTelepon.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    For K = 1 To LvTelepon.ListItems.Count
        LvTelepon.ListItems(K).Checked = False
    Next K
End Sub

Private Sub Form_Load()
    Call HeaderTelepon
    Call IsiMasalah
    
    Call IsiNoTelepon
End Sub

Private Sub IsiNoTelepon()
    CmbNoTelp.CLEAR
    
    CmbNoTelp.AddItem "Home 1"
    CmbNoTelp.AddItem "Home 2"
    CmbNoTelp.AddItem "Office 1"
    CmbNoTelp.AddItem "Office 2"
    CmbNoTelp.AddItem "Mobile 1"
    CmbNoTelp.AddItem "Mobile 2"
    CmbNoTelp.AddItem "EC"
    
    CmbNoTelp.AddItem "Add Home 1"
    CmbNoTelp.AddItem "Add Home 2"
    CmbNoTelp.AddItem "Add Office 1"
    CmbNoTelp.AddItem "Add Office 2"
    CmbNoTelp.AddItem "Add Mobile 1"
    CmbNoTelp.AddItem "Add Mobile 2"
End Sub


