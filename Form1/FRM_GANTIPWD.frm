VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frm_gantipas 
   BorderStyle     =   0  'None
   Caption         =   "Ganti Password"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4470
   ForeColor       =   &H00000000&
   Icon            =   "FRM_GANTIPWD.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      ForeColor       =   &H00000000&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   2025
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1665
      Width           =   2250
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H000000FF&
      Height          =   315
      Index           =   2
      Left            =   2010
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   540
      Width           =   2250
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00000000&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2025
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1350
      Width           =   2250
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00000000&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   2025
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1035
      Width           =   2250
   End
   Begin Threed.SSCommand Command1 
      Height          =   855
      Index           =   0
      Left            =   2610
      TabIndex        =   11
      Top             =   2070
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
      _Version        =   196610
      Font3D          =   5
      ForeColor       =   0
      PictureMaskColor=   15853019
      PictureFrames   =   1
      Picture         =   "FRM_GANTIPWD.frx":0442
      Caption         =   "&Ok"
      Alignment       =   8
   End
   Begin Threed.SSCommand Command1 
      Height          =   855
      Index           =   1
      Left            =   3510
      TabIndex        =   12
      Top             =   2070
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
      _Version        =   196610
      Font3D          =   5
      ForeColor       =   0
      PictureMaskColor=   15853019
      PictureFrames   =   1
      Picture         =   "FRM_GANTIPWD.frx":0937
      Caption         =   "&Batal"
      Alignment       =   8
   End
   Begin VB.Image Image1 
      Height          =   465
      Index           =   0
      Left            =   0
      Picture         =   "FRM_GANTIPWD.frx":0E90
      Stretch         =   -1  'True
      Top             =   0
      Width           =   450
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ganti Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   495
      TabIndex        =   10
      Top             =   45
      Width           =   2145
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ganti Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   315
      Left            =   540
      TabIndex        =   9
      Top             =   60
      Width           =   3885
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
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
      Left            =   450
      TabIndex        =   8
      Top             =   0
      Width           =   2325
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F1E5DB&
      Caption         =   "Old Password :"
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   6
      Left            =   30
      TabIndex        =   7
      Top             =   1080
      Width           =   1905
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F1E5DB&
      Caption         =   "New Password :"
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   5
      Left            =   30
      TabIndex        =   6
      Top             =   1395
      Width           =   1905
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F1E5DB&
      Caption         =   "Confirm New Password :"
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   4
      Left            =   30
      TabIndex        =   5
      Top             =   1725
      Width           =   1905
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F1E5DB&
      Caption         =   "User :"
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   2
      Left            =   15
      TabIndex        =   4
      Top             =   585
      Width           =   1905
   End
   Begin VB.Image Image2 
      Height          =   465
      Left            =   0
      Picture         =   "FRM_GANTIPWD.frx":199A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15000
   End
End
Attribute VB_Name = "frm_gantipas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
Dim M_objrs As ADODB.Recordset
Dim CMDSQL As String
Dim PASSENCRIPT As String
Dim alphanmr As Boolean
Dim sType As String
Dim m_pass As New ADODB.Recordset
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
        If Len(Text1(1).Text) < 8 Then
           MsgBox "Minimum lenght Character for Password is 8 Character", vbCritical + vbOKOnly, App.Title
           Text1(1).SetFocus
           Exit Sub
        End If
    If Text1(1).Text <> Text1(3).Text Then
        MsgBox "New password did not match", vbCritical + vbOKOnly, App.Title
        Text1(1).SetFocus
        Exit Sub
    Else
'        alphanmr = cekAlphaNumeric(Text1(1).Text)
'        If alphanmr = False Then
'            MsgBox "Password Must Contain Alpha and Numeric Character", vbCritical + vbOKOnly, App.Title
'            Text1(1).SetFocus
'            Exit Sub
'        Else
'            alphanmr = False
'            alphanmr = cekComplexity(Text1(1).Text)
'            If alphanmr = False Then
'                MsgBox "Password must meet Complexity requirements", vbCritical + vbOKOnly, App.Title
'                Text1(1).SetFocus
'                Exit Sub
'            End If
'        End If
    End If
    Dim m_cek As ADODB.Recordset
    Set m_cek = New ADODB.Recordset
    m_cek.CursorLocation = adUseClient
    CMDSQL = "Select * from TblHstPassword where UserId ='" + Text1(2).Text + "' and Password like '%" + Encrypt(Len(Text1(2).Text), Text1(1).Text) + "%' and F_VALID = 0 "
    m_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If m_cek.RecordCount <> 0 Then
        MsgBox "You Already Used this password", vbCritical + vbOKOnly, App.Title
        Text1(1).SetFocus
        Set m_cek = Nothing
        Exit Sub
    Else
        Set m_cek = Nothing
        Dim NEWPASSWORD As String
        Set m_cek = New ADODB.Recordset
        m_cek.CursorLocation = adUseClient
        NEWPASSWORD = Encrypt(Len(Text1(2).Text), Text1(1).Text)
        'NEWPASSWORD = Encrypt(Len(Text1(2).Text), Left(Text1(1).Text, 3))
        'NEWPASSWORD = Left(NEWPASSWORD, Len(NEWPASSWORD) - 1)
        CMDSQL = "Select * from TblHstPassword where UserId ='" + Text1(2).Text + "' and Password = '" + NEWPASSWORD + "' and F_VALID = 0 "
        m_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If m_cek.RecordCount <> 0 Then
            MsgBox "You Already Used this password", vbCritical + vbOKOnly, App.Title
            Text1(1).SetFocus
            Set m_cek = Nothing
            Exit Sub
        Else
        End If
    End If
    Set m_cek = Nothing
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
   ' M_OBJRS.Open "SELECT * FROM usertbl WHERE USERID ='" + Text1(2).Text + "' AND ACCREC = '" + Decrypt(Len(Text1(2).Text), Text1(0).Text) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    M_objrs.Open "SELECT * FROM usertbl WHERE USERID ='" + Text1(2).Text + "' AND ACCREC = '" + Encrypt(Len(Text1(2).Text), Text1(0).Text) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
  ' cmdsql = " SELECT * FROM usertbl WHERE USERID ='" + Text1(2).Text + "' AND ACCREC = '" + Decrypt(Len(Text1(2).Text), Text1(0).Text) + "'"
  CMDSQL = "SELECT * FROM usertbl WHERE USERID ='" + Text1(2).Text + "' AND ACCREC = '" + Encrypt(Len(Text1(2).Text), Text1(0).Text) + "'"
    If M_objrs.RecordCount <> 0 Then
'        'If ADD_HST_PASS = True Then
          
            Set m_pass = New ADODB.Recordset
            m_pass.CursorLocation = adUseClient
            CMDSQL = "Select * from TblHstPassword where UserId = '" + Text1(2).Text + "' AND F_VALID = 0 ORDER BY idhstpassword"
            m_pass.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                If m_pass.RecordCount = 18 Then
                    'UPDATE POSISI ID YG PALING KECIL.. ARTINYA ITU PASSWORD YG PALING LAMA TUH.....
                   m_pass.MoveFirst
                   m_pass!F_VALID = 1
                   m_pass.update
                   m_pass.Requery
                Else
                    'BELUM ADA 12 BIJI PASSWORD .. INSERT AJA LANGSUNG
                End If
            m_pass.AddNew
            m_pass!USERID = Text1(2).Text
            m_pass!password = Encrypt(Len(Text1(2).Text), Text1(1).Text)
            m_pass!F_VALID = 0
            m_pass.update
            m_pass.Requery
            Set m_pass = Nothing
'        'End If
        
        'UBAH DONG PASSWORDNYA
        
        'M_OBJRS!ACCREC = Encrypt(Len(Text1(2).Text), Text1(1).Text)
        'M_OBJRS!PWD = M_OBJRS!ACCREC
       
        M_OBJCONN.Execute "UPDATE USERTBL SET ACCREC='" + Encrypt(Len(Text1(2).Text), Text1(1).Text) + "',tgl_ubah_pass='" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd")) + "' WHERE USERID='" + Text1(2).Text + "'"
       
        'insert ke user log
       CMDSQL = "Insert Into TblLogUserAdm (UserId, Keterangan, UserType) VALUES ( '" + Text1(2).Text + "','Change Password','" + CStr(M_objrs("usertype")) + "') "
        M_OBJCONN.Execute CMDSQL
'        cmdsql = "UPDATE usertbl SET ACCREC = '" + Encrypt(Len(Text1(2).Text), Text1(1).Text) + "', PWD ='" + Encrypt(Len(Text1(2).Text), Text1(3).Text) + "', tgl_ubah_pass = '" + Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") + "'  "
'        cmdsql = cmdsql + " WHERE USERID = '" + Text1(2).Text + "' AND ACCREC = '" + Encrypt(Len(Text1(2).Text), Text1(0).Text) + "'"
'        M_OBJCONN.Execute cmdsql
        MsgBox "Password has been change", vbInformation, App.Title
        Unload Me
    Else
        'MsgBox "Wrong Password", vbInformation, App.Title
           
            Set m_pass = New ADODB.Recordset
            m_pass.CursorLocation = adUseClient
            CMDSQL = "Select * from TblHstPassword where UserId = '" + Text1(2).Text + "' AND F_VALID = 0 ORDER BY idhstpassword"
            m_pass.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                If m_pass.RecordCount = 18 Then
                    'UPDATE POSISI ID YG PALING KECIL.. ARTINYA ITU PASSWORD YG PALING LAMA TUH.....
                   m_pass.MoveFirst
                   m_pass!F_VALID = 1
                   m_pass.update
                   m_pass.Requery
                Else
                    'BELUM ADA 12 BIJI PASSWORD .. INSERT AJA LANGSUNG
                End If
            m_pass.AddNew
            m_pass!USERID = Text1(2).Text
            m_pass!password = Encrypt(Len(Text1(2).Text), Text1(1).Text)
            m_pass!F_VALID = 0
            m_pass.update
            m_pass.Requery
            Set m_pass = Nothing
'        'End If
        
        'UBAH DONG PASSWORDNYA
         M_OBJCONN.Execute "UPDATE USERTBL SET ACCREC='" + Encrypt(Len(Text1(2).Text), Text1(1).Text) + "',tgl_ubah_pass='" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd")) + "' WHERE USERID='" + Text1(2).Text + "'"
         
        If UCase(MDIForm1.txtlevel) = "AGENT" Then
            sType = "1"
        ElseIf UCase(MDIForm1.txtlevel) = "SUPERVISOR" Then
            sType = "2"
        ElseIf UCase(MDIForm1.txtlevel) = "MANAGER" Then
            sType = "3"
            Else
            
            sType = "4"
        End If
        
        'insert ke user log
       CMDSQL = "Insert Into TblLogUserAdm (UserId, Keterangan, UserType) VALUES ( '" + Text1(2).Text + "','Change Password','" + sType + "') "
       M_OBJCONN.Execute CMDSQL
'        cmdsql = "UPDATE usertbl SET ACCREC = '" + Encrypt(Len(Text1(2).Text), Text1(1).Text) + "', PWD ='" + Encrypt(Len(Text1(2).Text), Text1(3).Text) + "', tgl_ubah_pass = '" + Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") + "'  "
'        cmdsql = cmdsql + " WHERE USERID = '" + Text1(2).Text + "' AND ACCREC = '" + Encrypt(Len(Text1(2).Text), Text1(0).Text) + "'"
'        M_OBJCONN.Execute cmdsql
        MsgBox "Password has been change", vbInformation, App.Title
        Unload Me
        
        Set M_objrs = Nothing
        'Exit Sub
    End If
Case 1
    Unload Me
End Select
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

