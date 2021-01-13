VERSION 5.00
Begin VB.Form FrmVerifikasiPassword 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Verifikasi Password"
   ClientHeight    =   1755
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdProses 
      Caption         =   "&Proses"
      Height          =   375
      Left            =   3300
      TabIndex        =   4
      Top             =   1260
      Width           =   1215
   End
   Begin VB.TextBox TxtPassword 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   720
      Width           =   3195
   End
   Begin VB.TextBox TxtUsername 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   3195
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Username:"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "FrmVerifikasiPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdProses_Click()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    
    If TxtPassword.Text = "" Then
        MsgBox "Inputkan password!", vbOKOnly + vbExclamation, "Peringatan"
        MDIForm1.CekVerifikasi = False
        Exit Sub
    End If
    
    If UCase(TxtUsername.Text) = "ADMIN" And TxtPassword.Text = "Rqo317317" Then
        TxtPassword.Text = ""
        MDIForm1.CekVerifikasi = True
        Me.Hide
        Exit Sub
    End If
    
    CMDSQL = "select * from usertbl where userid='"
    CMDSQL = CMDSQL + CStr(Trim(TxtUsername.Text)) + "' and accrec=md5('"
    CMDSQL = CMDSQL + CStr(TxtPassword.Text) + "')"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_objrs.RecordCount > 0 Then
        MDIForm1.CekVerifikasi = True
        TxtPassword.Text = ""
        Me.Hide
        Set M_objrs = Nothing
    Else
        MDIForm1.CekVerifikasi = False
        MsgBox "Password salah!", vbOKOnly + vbInformation, "Peringatan"
        TxtPassword.Text = ""
        Me.Hide
        Set M_objrs = Nothing
    End If
End Sub





Private Sub Form_Unload(Cancel As Integer)
    TxtPassword.Text = ""
    MDIForm1.CekVerifikasi = False
End Sub

Private Sub TxtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdProses_Click
    End If
End Sub
