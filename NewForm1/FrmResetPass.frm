VERSION 5.00
Begin VB.Form FrmResetPass 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reset Password"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4770
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   435
      Left            =   3060
      TabIndex        =   9
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox TxtNamaAgent 
      Height          =   285
      Left            =   1860
      TabIndex        =   8
      Top             =   540
      Width           =   2595
   End
   Begin VB.CommandButton CmdReset 
      Caption         =   "&Reset"
      Height          =   435
      Left            =   1620
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox TxtConfirmPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1740
      Width           =   2595
   End
   Begin VB.TextBox TxtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1440
      Width           =   2595
   End
   Begin VB.ComboBox CmbAgent 
      Height          =   315
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "*)Kosongkan nama agent , jika tidak ingin diubah!"
      Height          =   495
      Left            =   1800
      TabIndex        =   11
      Top             =   900
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "*"
      Height          =   255
      Left            =   4560
      TabIndex        =   10
      Top             =   540
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "Nama Agent:"
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   540
      Width           =   1515
   End
   Begin VB.Label Label2 
      Caption         =   "Confirm Password:"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   3
      Top             =   1740
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Agent"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1395
   End
End
Attribute VB_Name = "FrmResetPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim USERID As String
Dim nama As String

Private Sub IsiAgent()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    
    CMDSQL = "select userid,agent from usertbl where userid is not null and agent is not null "
    If UCase(Trim(MDIForm1.txtlevel.Text)) = "TEAMLEADER" Then
        CMDSQL = CMDSQL + " and team='" + Trim(MDIForm1.TxtUsername.Text) + "' "
    End If
    CMDSQL = CMDSQL + " order by userid asc "
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    CmbAgent.CLEAR
    If M_objrs.RecordCount = 0 Then
        MsgBox "Data agent tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Set M_objrs = Nothing
        Unload Me
        Exit Sub
    End If
    
    While Not M_objrs.EOF
        CmbAgent.AddItem Trim(M_objrs("userid"))
        M_objrs.MoveNext
    Wend
    
    Set M_objrs = Nothing
End Sub

Private Sub CmdBatal_Click()
    Unload Me
End Sub

Private Sub CmdReset_Click()
    Dim CMDSQL As String
    Dim Konf As String
    
    Konf = MsgBox("Apakah anda yakin akan mereset password?", vbYesNo + vbQuestion, "Konfirmasi")
    If Konf = vbNo Then
        MsgBox "Reset password dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If CmbAgent.Text = "" Then
        MsgBox "Agent harus diisi!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If TxtConfirmPass.Text = "" Or TxtPass.Text = "" Then
        MsgBox "Password harus diisi!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If TxtConfirmPass.Text <> TxtPass.Text Then
        MsgBox "Konfirmasi password harus cocok dengan password!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    CMDSQL = "update usertbl set accrec=md5('" + TxtPass.Text + "'),tgl_ubah_pass=null "
    If Len(Trim(TxtNamaAgent.Text)) > 0 Then
        CMDSQL = CMDSQL + ",agent='" + UCase(TxtNamaAgent.Text) + "' "
    End If
    CMDSQL = CMDSQL + " where userid='" + CmbAgent.Text + "'"
    M_OBJCONN.Execute CMDSQL
    MsgBox "Password berhasil diubah!", vbOKOnly + vbInformation, "Informasi"
    Unload Me
    Exit Sub
salah:
    MsgBox "Ada error " & err.Description
End Sub

Private Sub Form_Load()
    Call IsiAgent
End Sub
