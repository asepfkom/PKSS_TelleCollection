VERSION 5.00
Begin VB.Form FrmGantiPassword 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ganti Password"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3945
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtCoding 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1620
      TabIndex        =   0
      Top             =   120
      Width           =   2235
   End
   Begin VB.TextBox TxtPass 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1620
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   540
      Width           =   2235
   End
   Begin VB.TextBox TxtRePass 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1620
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   900
      Width           =   2235
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   435
      Left            =   780
      TabIndex        =   3
      Top             =   1500
      Width           =   1215
   End
   Begin VB.CommandButton CmdKeluar 
      Caption         =   "&Keluar"
      Height          =   435
      Left            =   1980
      TabIndex        =   4
      Top             =   1500
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Coding:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   1395
   End
   Begin VB.Label Label3 
      Caption         =   "Confirm Password:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   900
      Width           =   1395
   End
End
Attribute VB_Name = "FrmGantiPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdkeluar_Click()
    Dim a As String
    
    a = MsgBox("Anda yakin akan keluar?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbYes Then
        End
    End If
End Sub

Private Sub cmdOK_Click()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    
    If TxtPass.Text = "" Then
        MsgBox "Password tidak boleh kosong!", vbOKOnly + vbExclamation, "Informasi"
        Exit Sub
    End If
    
    'Minimal password 6 karakter
    If Len(TxtPass.Text) < 6 Then
        MsgBox "Minimal karakter password 6 karakter!", vbOKOnly + vbExclamation, "Informasi"
        Exit Sub
    End If
    
    If TxtPass.Text <> TxtRePass.Text Then
        MsgBox "Password dan konfirmasi password tidak cocok!", vbOKOnly + vbExclamation, "Informasi"
        Exit Sub
    End If
    
    'Cek password apakah sama dengan password yang dahulu
    CMDSQL = "select * from mandiri.usertbl where userid='"
    CMDSQL = CMDSQL + Trim(TxtCoding.Text) + "' and accrec=md5('"
    CMDSQL = CMDSQL + Trim(TxtPass.Text) + "')"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
        MsgBox "Password anda sama dengan password yang lalu! Harap ganti!", vbOKOnly + vbExclamation, "Peringatan"
        Set M_objrs = Nothing
        Exit Sub
    End If
    
    Set M_objrs = Nothing
    
    'Update password baru
    CMDSQL = "update mandiri.usertbl set accrec=md5('"
    CMDSQL = CMDSQL + TxtPass.Text + "'),tgl_ubah_pass=date(now()) where userid='"
    CMDSQL = CMDSQL + Trim(TxtCoding.Text) + "'"
    M_OBJCONN.Execute CMDSQL
    
    MsgBox "Password anda berhasil diubah!", vbOKOnly + vbInformation, "Informasi"
    Unload Me
End Sub

