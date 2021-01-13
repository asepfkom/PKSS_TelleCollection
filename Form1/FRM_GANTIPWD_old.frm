VERSION 5.00
Begin VB.Form frm_gantipas 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Ganti Password"
   ClientHeight    =   2475
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5610
   ForeColor       =   &H00000000&
   Icon            =   "FRM_GANTIPWD.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      ForeColor       =   &H00000000&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   2475
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1305
      Width           =   2250
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3135
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1815
      UseMaskColor    =   -1  'True
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "&Keluar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   4350
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1815
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H000000FF&
      Height          =   315
      Index           =   2
      Left            =   2220
      MaxLength       =   20
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Width           =   2250
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00000000&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2475
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   990
      Width           =   2250
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00000000&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   2475
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   675
      Width           =   2250
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Old Password :"
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   6
      Left            =   480
      TabIndex        =   9
      Top             =   720
      Width           =   1905
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "New Password :"
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   5
      Left            =   480
      TabIndex        =   8
      Top             =   1035
      Width           =   1905
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Confirm New Password :"
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   4
      Left            =   480
      TabIndex        =   7
      Top             =   1365
      Width           =   1905
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "User :"
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   2
      Left            =   735
      TabIndex        =   6
      Top             =   225
      Width           =   1380
   End
End
Attribute VB_Name = "frm_gantipas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CMDSQL2 As String

Private Sub Command1_Click(Index As Integer)
Dim M_objrs As ADODB.Recordset
Dim CMDSQL As String
Dim PASSENCRIPT As String
Dim alphanmr As Boolean
Select Case Index
Case 0
    If Text1(0).Text = Empty Then
        MsgBox "Enter Your Old Password", vbCritical + vbOKOnly, App.Title
        Text1(0).SetFocus
        Exit Sub
    End If
    If Text1(1).Text = Empty Then
        MsgBox "Enter Your New Password", vbCritical + vbOKOnly, App.Title
        Text1(1).SetFocus
        Exit Sub
    End If
    If Text1(0).Text = Text1(1).Text Then
        MsgBox "New Password Must be not the same with old password", vbCritical + vbOKOnly, App.Title
        Text1(1).SetFocus
        Exit Sub
    End If
    If Text1(2).Text = Text1(1).Text Then
        MsgBox "New Password Must be not the same with userid", vbCritical + vbOKOnly, App.Title
        Text1(1).SetFocus
        Exit Sub
    End If
        If Len(Text1(1).Text) < 7 Then
           MsgBox "Minimum lenght Character for Password is 7 Character", vbCritical + vbOKOnly, App.Title
           Text1(1).SetFocus
           Exit Sub
        End If
    If Text1(1).Text <> Text1(3).Text Then
        MsgBox "New password did not match", vbCritical + vbOKOnly, App.Title
        Text1(1).SetFocus
        Exit Sub
    Else
''        alphanmr = cekAlphaNumeric(Text1(1).Text)
''        If alphanmr = False Then
''            MsgBox "Password Must Contain Alpha and Numeric Character", vbCritical + vbOKOnly, App.Title
''            Text1(1).SetFocus
''            Exit Sub
''        Else
''            alphanmr = False
''            alphanmr = cekComplexity(Text1(1).Text)
'            If alphanmr = False Then
'                MsgBox "Password must meet Complexity requirements", vbCritical + vbOKOnly, App.Title
'                Text1(1).SetFocus
'                Exit Sub
'            End If
'        End If
    End If
'    Dim m_cek As ADODB.Recordset
'    Set m_cek = New ADODB.Recordset
'    m_cek.CursorLocation = adUseClient
'    CMDSQL = "Select * from TblHstPassword where UserId ='" + Text1(2).Text + "' and Password like '%" + Encrypt(Len(Text1(2).Text), Text1(1).Text) + "%' and F_VALID = 0 "
'    m_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'    If m_cek.RecordCount <> 0 Then
'        MsgBox "You Already Used this password", vbCritical + vbOKOnly, App.Title
'        Text1(1).SetFocus
'        Set m_cek = Nothing
'        Exit Sub
'    Else
'        Set m_cek = Nothing
'        Dim NEWPASSWORD As String
'        Set m_cek = New ADODB.Recordset
'        m_cek.CursorLocation = adUseClient
'        NEWPASSWORD = Encrypt(Len(Text1(2).Text), Left(Text1(1).Text, 3))
'        NEWPASSWORD = Left(NEWPASSWORD, Len(NEWPASSWORD) - 1)
'        CMDSQL = "Select * from TblHstPassword where UserId ='" + Text1(2).Text + "' and Password like '%" + NEWPASSWORD + "%' and F_VALID = 0 "
'        m_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'        If m_cek.RecordCount <> 0 Then
'            MsgBox "You Already Used this password", vbCritical + vbOKOnly, App.Title
'            Text1(1).SetFocus
'            Set m_cek = Nothing
'            Exit Sub
'        Else
'        End If
'    End If
'    Set m_cek = Nothing
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
'    m_objrs.Open "SELECT * FROM usertbl WHERE USERID ='" + Text1(2).Text + "' ", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
       'M_OBJRS.Open "SELECT * FROM usertbl WHERE USERID ='" + Text1(2).Text + "' AND ACCREC = '" + Encrypt(Len(Text1(2).Text), Text1(0).Text) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    CMDSQL2 = "select * from usertbl where userid='"
    CMDSQL2 = CMDSQL2 + Trim(Text1(2).Text) + "' and accrec=md5('"
    CMDSQL2 = CMDSQL2 + Trim(Text1(0).Text) + "')"
    M_objrs.Open CMDSQL2, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount <> 0 Then
'        'If ADD_HST_PASS = True Then
'            Dim m_pass As ADODB.Recordset
'            Set m_pass = New ADODB.Recordset
'            m_pass.CursorLocation = adUseClient
'            CMDSQL = "Select * from TblHstPassword where UserId = '" + Text1(2).Text + "' AND F_VALID = 0 ORDER BY ID"
'            m_pass.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'                If m_pass.RecordCount = 12 Then
'                    'UPDATE POSISI ID YG PALING KECIL.. ARTINYA ITU PASSWORD YG PALING LAMA TUH.....
'                   m_pass.MoveFirst
'                   m_pass!F_VALID = 1
'                   m_pass.UPDATE
'                   m_pass.Requery
'                Else
'                    'BELUM ADA 12 BIJI PASSWORD .. INSERT AJA LANGSUNG
'                End If
'            m_pass.AddNew
'            m_pass!USERID = Text1(2).Text
'            m_pass!password = Encrypt(Len(Text1(2).Text), Text1(1).Text)
'            m_pass!F_VALID = 0
'            m_pass.UPDATE
'            m_pass.Requery
'            Set m_pass = Nothing
'        'End If
        
        'UBAH DONG PASSWORDNYA
''        Text1(1).Text = "pass123"
''        M_OBJRS!ACCREC = Encrypt(Len(Text1(2).Text), Text1(1).Text)
''        M_OBJRS!PWD = M_OBJRS!ACCREC
'''        m_objrs!TGLGANTIPWD = Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd")
''        M_OBJRS.update

        CMDSQL = "update usertbl set accrec=md5('"
        CMDSQL = CMDSQL + Trim(Text1(1).Text) + "') where userid='"
        CMDSQL = CMDSQL + Trim(Text1(2).Text) + "'"
        M_OBJCONN.Execute CMDSQL
        'insert ke user log
'        cmdsql = "Insert Into TblLogUserAdm ( UserId, Keterangan, UserType) VALUES ( '" + mdiform1.txtusername.text + "','Change Password','" + mdiform1.txtlevel.text + "') "
'        M_OBJCONN.Execute cmdsql
'        cmdsql = "UPDATE usertbl SET ACCREC = '" + Encrypt(Len(Text1(2).Text), Text1(1).Text) + "', PWD ='" + Encrypt(Len(Text1(2).Text), Text1(3).Text) + "', TGLGANTIPWD = '" + Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") + "'  "
'        cmdsql = cmdsql + " WHERE USERID = '" + Text1(2).Text + "' AND ACCREC = '" + Encrypt(Len(Text1(2).Text), Text1(0).Text) + "'"
'        M_OBJCONN.Execute cmdsql
        MsgBox "Password has been change", vbInformation, App.Title
        Unload Me
    Else
        MsgBox "Wrong Password", vbInformation, App.Title
        Text1(0).SetFocus
        Set M_objrs = Nothing
        Exit Sub
    End If
Case 1
    Unload Me
End Select
'Dim M_OBJRS As ADODB.Recordset
'Dim cmdsql As String
'Select Case Index
'Case 0
'    If Text1(0).Text = Empty Then
'        MsgBox "Masukan Password Lama Anda", vbInformation, "Aplikasi"
'        Text1(0).SetFocus
'        Exit Sub
'    End If
'    If Text1(1).Text = Empty Then
'        MsgBox "Masukan Password Baru Anda", vbInformation, "Aplikasi"
'        Text1(1).SetFocus
'        Exit Sub
'    End If
'    Set M_OBJRS = New ADODB.Recordset
'    M_OBJRS.CursorLocation = adUseClient
'    M_OBJRS.Open "SELECT USERID FROM usertbl WHERE USERID ='" + Text1(2).Text + "' AND ACCREC = '" + Text1(0).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'    If M_OBJRS.RecordCount <> 0 Then
'        Set M_OBJRS = Nothing
'        cmdsql = "UPDATE usertbl SET ACCREC = '" + Text1(1).Text + "', PWD ='" + Text1(3).Text + "'"
'        cmdsql = cmdsql + " WHERE USERID = '" + Text1(2).Text + "' AND ACCREC = '" + Text1(0).Text + "'"
'        M_OBJCONN.Execute cmdsql
'        MsgBox "Password Telah Diganti", vbInformation, "Aplikasi"
'        Unload Me
'    Else
'        MsgBox "Password Lama Yang Anda Masukan Salah", vbInformation, "Aplikasi"
'        Text1(0).SetFocus
'        Set M_OBJRS = Nothing
'        Exit Sub
'    End If
'Case 1
'    Unload Me
'End Select
End Sub

Private Sub Form_Load()
If UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
'    Label1(3).Visible = False
 '   Text1(3).Visible = False
End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 0
    Select Case KeyAscii
        Case 13
            Call Command1_Click(0)
        Case 32
            KeyAscii = 0
    End Select
Case 1
    Select Case KeyAscii
        Case 13
            Call Command1_Click(0)
        Case 32
            KeyAscii = 0
    End Select
End Select
End Sub


Private Function cekAlphaNumeric(password As String) As Boolean
Dim a As String
Dim syarat1 As Boolean
Dim syarat2 As Boolean
Dim i As Integer
syarat1 = False
syarat2 = False
    For i = 1 To Len(password)
    If i = 1 Then
        a = Left(password, 1)
    Else
        a = Mid(password, i, 1)
    End If
    Select Case Asc(a)
        Case 48 To 57
          syarat1 = True
        Case Else
            syarat2 = True
    End Select
    Next i
cekAlphaNumeric = syarat1 * syarat2
End Function

Private Function cekComplexity(password As String) As Boolean
Dim a As String
Dim syarat1 As Boolean
Dim syarat2 As Boolean
Dim i As Integer
syarat1 = False
syarat2 = False
    For i = 1 To Len(password)
    If i = 1 Then
        a = Left(password, 1)
    Else
        a = Mid(password, i, 1)
    End If
    Select Case Asc(a)
        Case 65 To 90
          syarat1 = True
        Case 97 To 122
            syarat2 = True
    End Select
    Next i
cekComplexity = syarat1 * syarat2
End Function

