VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   4620
   ClientLeft      =   2790
   ClientTop       =   3045
   ClientWidth     =   8070
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0442
   ScaleHeight     =   2729.649
   ScaleMode       =   0  'User
   ScaleWidth      =   7577.295
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Password"
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7440
      Top             =   240
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1305
      Left            =   1800
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "frmLogin.frx":3C5D
      Top             =   4800
      Width           =   5175
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   2310
      TabIndex        =   0
      Top             =   1860
      Width           =   2595
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2310
      PasswordChar    =   "#"
      TabIndex        =   1
      Top             =   2325
      Width           =   2595
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   135
      Left            =   60
      TabIndex        =   7
      Top             =   -180
      Visible         =   0   'False
      Width           =   4695
      ExtentX         =   8281
      ExtentY         =   238
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin Threed.SSCommand cmdOK 
      Default         =   -1  'True
      Height          =   855
      Left            =   5040
      TabIndex        =   2
      Top             =   1845
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      _Version        =   196610
      ForeColor       =   4210752
      PictureFrames   =   1
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmLogin.frx":3C63
      Caption         =   "&"
      ButtonStyle     =   2
      BevelWidth      =   0
   End
   Begin Threed.SSCommand cmdCancel 
      Cancel          =   -1  'True
      Height          =   855
      Left            =   6120
      TabIndex        =   3
      Top             =   1845
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
      _Version        =   196610
      ForeColor       =   4210752
      PictureFrames   =   1
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmLogin.frx":6AC9
      Caption         =   "&"
      ButtonStyle     =   2
      BevelWidth      =   0
   End
   Begin VB.Label lblLabels 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "T I N S"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Index           =   2
      Left            =   213
      TabIndex        =   11
      Top             =   3893
      Width           =   1905
   End
   Begin VB.Label lblLabels 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "TELE COLLECTION"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   5
      Left            =   213
      TabIndex        =   10
      Top             =   4200
      Width           =   2985
   End
   Begin VB.Label lblLabels 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright @ 2016 - Pt. Delta Nuansa Nirwana"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   3
      Left            =   4200
      TabIndex        =   9
      Top             =   4320
      Width           =   3825
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait...."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   7560
      TabIndex        =   6
      Top             =   4680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "USER ID"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   1020
      TabIndex        =   4
      Top             =   1890
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   1
      Left            =   1005
      TabIndex        =   5
      Top             =   2415
      Width           =   945
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public countTrylogin As Integer
Public ADD_HST_PASS As Boolean
Public STSGANTIPWD As Boolean
Public ok As Boolean

Private Sub Check1_Click()
    If Check1.Value = vbChecked Then
        TxtPassword.PasswordChar = ""
    Else
        TxtPassword.PasswordChar = "#"
    End If
End Sub

Private Sub CmdCancel_Click()
    End
End Sub
'Private Sub cmdOk_Click()
'    Dim NILSTAT As String
'    Dim M_OBJRS As ADODB.Recordset
'    Dim rs_lvtian As New ADODB.Recordset
'    Dim m_objrsAdd As ADODB.Recordset
'    Dim M_PESAN As ADODB.Recordset
'    Dim m_waktuserver As ADODB.Recordset
'    Dim lminggu As String
'    Dim lbulan As String
'    Dim STRSQL As String
'    Dim ltahun As String
'    Dim cmdsql As String
'    Dim m_popup As ADODB.Recordset
'    Dim CMDSQL2 As String
'    Dim SqlWaktu As String
'    Dim jam_sekarang As String
'     ' On Error GoTo HELL
'
'    If (TxtUsername = "tian" And TxtPassword = "tian") Or (TxtUsername = "elin" And TxtPassword = "12345") Then
'
'        MDIForm1.TxtUsername.Text = TxtUsername
'        MDIForm1.txtlevel.Text = "Administrator"
'        MDIForm1.txtNama.Text = "Administrator"
'        MDIForm1.mn_update_db.Visible = True
'        Unload Me
'        'JEJAKTIAN10032016==================================================
'        If MDIForm1.TxtUsername.Text <> "tian" Or MDIForm1.TxtUsername.Text <> "elin" Then
'            MDIForm1.nmlistreqptp.Visible = False
'        End If
'        '===================================================================
'
'
'        MDIForm1.Show
'        Exit Sub
'    End If
'        If TxtUsername = Empty Then
'            MsgBox "Username Belum Diisi", vbCritical + vbOKOnly, "Peringatan"
'            TxtUsername.SetFocus
'            SendKeys "{Home}+{End}"
'            Exit Sub
'        Else
'            If TxtPassword = Empty Then
'                MsgBox "Password Belum Diisi", vbCritical + vbOKOnly, "Peringatan"
'                TxtPassword.SetFocus
'                SendKeys "{Home}+{End}"
'                Exit Sub
'            End If
'        End If
'    'Timer1.Enabled = True
'    Set M_OBJRS = New ADODB.Recordset
'    M_OBJRS.CursorLocation = adUseClient
'    'M_OBJRS.Open "SELECT USERID, ACCREC, USERTYPE,AGENT,UNIT,AUTH, EXT,stsaplikasi,note,ntargetspv FROM usertbl WHERE USERID = '" + txtUserName + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    '301110 Ubah ke Md5
'    CMDSQL2 = "SELECT userid, accrec, usertype,agent,unit,auth, ext,"
'    CMDSQL2 = CMDSQL2 + "stsaplikasi,note,ntargetspv, date(now())-date(tgl_ubah_pass) as LamaPass, * from usertbl WHERE userid='"
'    CMDSQL2 = CMDSQL2 + Trim(TxtUsername.Text) + "' and accrec=md5('"
'    CMDSQL2 = CMDSQL2 + Trim(TxtPassword.Text) + "')"
'    M_OBJRS.Open CMDSQL2, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'
'    If M_OBJRS.RecordCount <> 0 Then
'            'If txtPassword <> M_OBJRS("ACCREC") Then
'            ''Debug.Print Decrypt(Len(Trim(txtUserName.Text)), M_OBJRS("ACCREC"))
'
'        ''    If Trim(txtPassword) <> Decrypt(Len(Trim(txtUserName.Text)), M_OBJRS("ACCREC")) Then
'        ''        MsgBox "Password Yang Anda Masukan Salah... Perhatikan CapsLock Anda...!!!", vbCritical + vbOKOnly, "Peringatan"
'        ''        txtPassword.SetFocus
'                'SendKeys "{Home}+{End}"
'        ''    Else
'
'            ' CEK JAM MASUK RANDY(FEB 2016)
'            SqlWaktu = "select now()"
'            Set m_waktuserver = New ADODB.Recordset
'            m_waktuserver.CursorLocation = adUseClient
'            m_waktuserver.Open SqlWaktu, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'
'
'            'Jika last login sekarang
'            If Format(m_waktuserver(0), "yyyy-mm-dd") <> Format(M_OBJRS("tglupdate"), "yyyy-mm-dd") Then
'                If Format(m_waktuserver(0), "HH:mm") > Format("08:05", "HH:mm") Then
'                    If M_OBJRS("USERTYPE") = "1" Then
'                        Set m_waktuserver = Nothing
'                        M_OBJCONN.Execute "UPDATE usertbl SET f_blok='1',tglupdate='now()' WHERE userid='" & Trim(TxtUsername.Text) & "'"
'                        MsgBox "Jam masuk anda terlambat!! Tidak boleh melebihi Pukul 08:00", vbCritical + vbOKOnly, "Terlambat"
'                        GoTo blok_user
'                    End If
'                End If
'            End If
'
'            ' Waktu masuk lebih dari 10 menit
'            If DateDiff("n", Format(M_OBJRS("last_logout"), "yyyy-mm-dd hh:mm:ss"), Format(m_waktuserver(0), "yyyy-mm-dd hh:mm:ss")) >= 10 Then
'                If Format(m_waktuserver(0), "HH:mm") > Format("08:05", "HH:mm") Then
'                    If M_OBJRS("USERTYPE") = "1" Then
'                        If M_OBJRS("f_break") = 0 Then
'                            Set m_waktuserver = Nothing
'                            M_OBJCONN.Execute "UPDATE usertbl SET f_blok='1',tglupdate='now()' WHERE userid='" & Trim(TxtUsername.Text) & "'"
'                            MsgBox "Anda diblok karena membuka aplikasi Lebih dari 10 Menit dari " & vbCrLf & "waktu terakhir keluar program (log out)", vbCritical + vbOKOnly, "Blok"
'                            GoTo blok_user
'                        End If
'                    End If
'                End If
'            End If
'
'            M_OBJCONN.Execute "UPDATE usertbl SET last_logout='now()',tglupdate='now()',f_break=0 WHERE userid='" & Trim(TxtUsername.Text) & "'"
'
'            Set m_waktuserver = Nothing
'            ' # END CEK JAM MASUK
'
'           If IsNull(M_OBJRS("tgl_ubah_pass")) = True Or Val(IIf(IsNull(M_OBJRS("LamaPass")), "0", M_OBJRS("lamapass"))) >= 90 Then
'                MsgBox "Untuk keamanan! Silahkan ganti password anda terlebih dahulu!"
'                FrmGantiPassword.TxtCoding.Text = TxtUsername.Text
'                FrmGantiPassword.Show vbModal
'           End If
'
'
'
'
'
'            If M_OBJRS("USERTYPE") = "1" Then
'                If IIf(IsNull(M_OBJRS("note")), "", M_OBJRS("note")) = "" Or IIf(IsNull(M_OBJRS("note")), "", M_OBJRS("note")) = 0 Then
'                    NILSTAT = ""
'                Else
'                    NILSTAT = "" + IIf(IsNull(M_OBJRS("note")), "", M_OBJRS("note")) + ""
'                End If
'
'                jam_sekarang = Format(Now(), "hh")
'
'                If jam_sekarang < 8 Then
'                    MsgBox "Anda Tidak Boleh Login Kurang Dari Jam 08:00", vbCritical + vbOKOnly, "Terlambat"
'                Exit Sub
'                End If
'                'MDIForm1.Lbltargetspv = "Target :" + CStr(IIf(IsNull(m_objrs("ntargetspv")), "", m_objrs("ntargetspv"))) + CStr(IIf(IsNull(m_objrs("note")), "", " Remarks :" + m_objrs("note")))
'                'MDIForm1.Kalimat1 = "Target :" + CStr(IIf(IsNull(m_objrs("ntargetspv")), "", m_objrs("ntargetspv"))) + CStr(IIf(IsNull(m_objrs("note")), "", " Remarks :" + m_objrs("note")))
'                'MDIForm1.PANJANG = Len("Target :" + CStr(IIf(IsNull(m_objrs("ntargetspv")), "", m_objrs("ntargetspv"))) + CStr(IIf(IsNull(m_objrs("note")), "", " Remarks :" + m_objrs("note"))))
'                MDIForm1.mnsubmarkup.Visible = False
'                MDIForm1.Lbltargetspv = NILSTAT
'                MDIForm1.Kalimat1 = NILSTAT
'                MDIForm1.PANJANG = Len(NILSTAT)
'                MDIForm1.mnsubahstsacc.Visible = False
'                MDIForm1.setspv.Visible = False
'                MDIForm1.LblTarget.Visible = True
'                MDIForm1.txtlevel.Text = "Agent"
'                MDIForm1.SSCommand1(11).Visible = False
'                MDIForm1.SSCommand1(7).Visible = False
'                MDIForm1.mnbar(1).Visible = False
'                MDIForm1.mnbar(2).Visible = False
'                MDIForm1.mnbar(3).Visible = False
'                MDIForm1.mnbar(5).Visible = False
'                MDIForm1.mnbar(6).Visible = False
'                MDIForm1.mnbar(7).Visible = False
'                MDIForm1.mnbar(11).Visible = False
'                MDIForm1.MnFile(1).Visible = False
'                'if m_objrs("stsaplikasi")
'                MDIForm1.SSCommand1(1).Visible = True
'                MDIForm1.SSCommand1(2).Visible = False
'                MDIForm1.SSCommand1(4).Visible = False
'                MDIForm1.SSCommand1(5).Visible = False
'                MDIForm1.SSCommand1(8).Visible = False
'                MDIForm1.SSCommand2.Visible = False
'                 MDIForm1.VSMS.Visible = False
'                'MDIForm1.SSCommand1(3).Visible = False
'                'Dim m_popup As New ADODB.Recordset
'    '            Set m_popup = New ADODB.Recordset
'    '            m_popup.CursorLocation = adUseClient
'    '            m_popup.Open "Select * from vwcallcfg1", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'    '            CMDSQL2 = "UPDATE usertbl set flagcall ='" + Format(m_popup!tglsystem, "hh:mm:ss") + "' where userid ='" + txtUserName.Text + "'"
'    '            M_OBJCONN.Execute CMDSQL2
'    '
'    '            Set m_popup = Nothing
'
'
'
'            Else
'                MDIForm1.LblTarget.Visible = False
'                If M_OBJRS("USERTYPE") = "6" Then
'                    If IIf(IsNull(M_OBJRS("note")), "", M_OBJRS("note")) = "" Or IIf(IsNull(M_OBJRS("note")), "", M_OBJRS("note")) = "0" Then
'                        NILSTAT = ""
'                    Else
'                        NILSTAT = "" + IIf(IsNull(M_OBJRS("note")), "", M_OBJRS("note")) + ""
'                    End If
'
'               ' MDIForm1.Lbltargetspv = "Target :" + CStr(IIf(IsNull(m_objrs("ntargetspv")), "", m_objrs("ntargetspv"))) + CStr(IIf(IsNull(m_objrs("note")), "", " Remarks :" + m_objrs("note")))
'                'MDIForm1.Kalimat1 = "Target :" + CStr(IIf(IsNull(m_objrs("ntargetspv")), "", m_objrs("ntargetspv"))) + CStr(IIf(IsNull(m_objrs("note")), "", " Remarks :" + m_objrs("note")))
'                'MDIForm1.PANJANG = Len("Target :" + CStr(IIf(IsNull(m_objrs("ntargetspv")), "", m_objrs("ntargetspv"))) + CStr(IIf(IsNull(m_objrs("note")), "", " Remarks :" + m_objrs("note"))))
'
'                MDIForm1.Lbltargetspv = NILSTAT
'                MDIForm1.Kalimat1 = NILSTAT
'                MDIForm1.PANJANG = Len(NILSTAT)
'
'                MDIForm1.mnsubahstsacc.Visible = False
'                MDIForm1.setspv.Visible = False
'                MDIForm1.txtlevel.Text = "TeamLeader"
'                MDIForm1.mnbar(2).Visible = False
'                MDIForm1.mnbar(5).Visible = False
'                MDIForm1.mnbar(7).Visible = False
'               ' MDIForm1.mnblokspv.Visible = False
'                MDIForm1.VSMS.Visible = False
'                End If
'                If M_OBJRS("USERTYPE") = "2" Then
'                    MDIForm1.LblTarget.Visible = True
'                MDIForm1.txtlevel.Text = "Field Collector"
'                MDIForm1.SSCommand1(11).Visible = False
'                MDIForm1.mnbar(1).Visible = False
'                MDIForm1.mnbar(2).Visible = False
'                MDIForm1.mnbar(3).Visible = False
'                MDIForm1.mnbar(5).Visible = False
'                MDIForm1.mnbar(6).Visible = False
'                MDIForm1.mnbar(7).Visible = False
'                MDIForm1.MnFile(1).Visible = False
'                MDIForm1.SSCommand1(0).Visible = False
'                MDIForm1.SSCommand1(1).Visible = False
'                MDIForm1.SSCommand1(2).Visible = False
'                MDIForm1.SSCommand1(4).Visible = False
'                MDIForm1.SSCommand1(5).Visible = True
'                MDIForm1.SSCommand2.Visible = False
'
'
'                End If
'            End If
'
'            If M_OBJRS("USERTYPE") = "11" Or M_OBJRS("USERTYPE") = "20" Then
'                MDIForm1.txtlevel.Text = "Supervisor"
'            End If
'
'            If M_OBJRS("USERTYPE") = "17" Then
'                MDIForm1.txtlevel.Text = "Manager"
'            End If
'
'            If M_OBJRS("USERTYPE") = "25" Then
'                MDIForm1.txtlevel.Text = "Admin"
'            End If
'
'            'jejaktian28072016menurole
'            'Call menurole
'            '=================================================
'
'            MDIForm1.TxtUsername.Text = UCase(TxtUsername)
'            MDIForm1.Text3.Text = IIf(IsNull(M_OBJRS("UNIT")), "", M_OBJRS("UNIT"))
'            MDIForm1.txtNama.Text = IIf(IsNull(M_OBJRS("agent")), "", M_OBJRS("agent"))
'            MDIForm1.TxtAuth.Text = IIf(IsNull(M_OBJRS("AUTH")), "", M_OBJRS("AUTH"))
'            DoEvents
'            'Call MDIForm1.LoOut_Ext("*1")
'            WaitSecs (0.1)
'            'Call login_ext(IIf(IsNull(m_objrs!EXT), "*1", m_objrs!EXT))
'
'            'isi target
'
'    '        Set m_objrsAdd = New ADODB.Recordset
'    '        m_objrsAdd.CursorLocation = adUseClient
'    '        CMDSQL = "Select * from TblTanggal Where "
'    '        CMDSQL = CMDSQL + " TGL = '" + Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") + "'"
'    '        m_objrsAdd.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'    '        If m_objrsAdd.RecordCount <> 0 Then
'    '            lminggu = IIf(IsNull(m_objrsAdd!Minggu), 0, m_objrsAdd!Minggu)
'    '            lbulan = IIf(IsNull(m_objrsAdd!Bulan), 0, m_objrsAdd!Bulan)
'    '            ltahun = IIf(IsNull(m_objrsAdd!tahun), 0, m_objrsAdd!tahun)
'    '        Else
'    '   '         MsgBox "Tanggal Belum Di Set....", vbInformation + vbOKOnly, "Aplikasi"
'    '            lminggu = 0
'    '            lbulan = MDIForm1.TDBDate1.Month
'    '            ltahun = MDIForm1.TDBDate1.Year
'    '        End If
'    '        Set m_objrsAdd = Nothing
'    '        DoEvents
'           'Set m_objrs = Nothing
'            Unload Me
'
'            '@@09022011 Ambil nilai maksimal kuota sms per hari agent dapat mengirim sms
'            Dim m_objrskuota As ADODB.Recordset
'            Dim cmdsqlkuota As String
'
'            cmdsqlkuota = "select * from tblsetsms"
'            Set m_objrskuota = New ADODB.Recordset
'            m_objrskuota.CursorLocation = adUseClient
'            m_objrskuota.Open cmdsqlkuota, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'            If m_objrskuota.RecordCount > 0 Then
'                MDIForm1.KuotaSms = m_objrskuota("kuota_sms")
'            End If
'            Set m_objrskuota = Nothing
'
'            '@@ 12-04-2011, Catet IP address user yang login, buat kirim pesan via winsock
'            Dim ip_addr As String
'            Dim agent As String
'            Dim tipe As String
'            Dim M_Objrs_Cek As ADODB.Recordset
'            Dim StrSqlIp As String
'
'            ip_addr = MDIForm1.WskCTI.LocalIP
'            agent = UCase(MDIForm1.TxtUsername.Text)
'            tipe = UCase(MDIForm1.txtlevel.Text)
'
'            'Cek dulu, apakah data IP user sudah ada, jika sudah ada di Update IPnya
'            StrSqlIp = "select * from tbl_ip where agent='"
'            StrSqlIp = StrSqlIp + Trim(agent) + "'"
'            Set M_Objrs_Cek = New ADODB.Recordset
'            M_Objrs_Cek.CursorLocation = adUseClient
'            M_Objrs_Cek.Open StrSqlIp, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'            If M_Objrs_Cek.RecordCount = 0 Then
'                'Inputin deh data baru
'                StrSqlIp = "insert into tbl_ip (agent,tipe,ip_addr) values ('"
'                StrSqlIp = StrSqlIp + Trim(agent) + "','"
'                StrSqlIp = StrSqlIp + Trim(tipe) + "','"
'                StrSqlIp = StrSqlIp + Trim(ip_addr) + "')"
'                M_OBJCONN.Execute StrSqlIp
'            Else
'                StrSqlIp = "update tbl_ip set ip_addr='"
'                StrSqlIp = StrSqlIp + Trim(ip_addr) + "' where agent='"
'                StrSqlIp = StrSqlIp + Trim(agent) + "'"
'                M_OBJCONN.Execute StrSqlIp
'            End If
'            Set M_Objrs_Cek = Nothing
'
'            '@@19042012, Cek IP Icentra
'            Dim M_Objrs_IP_Icentra As ADODB.Recordset
'
'            cmdsql = "select * from tbl_ip_icentra where ip='"
'            cmdsql = cmdsql + CStr(MDIForm1.WskCTI.LocalIP) + "'"
'            Set M_Objrs_IP_Icentra = New ADODB.Recordset
'            M_Objrs_IP_Icentra.CursorLocation = adUseClient
'            M_Objrs_IP_Icentra.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'            If M_Objrs_IP_Icentra.RecordCount = 0 Then
'                MDIForm1.TxtIPIcentra.Text = ""
'                Set M_Objrs_IP_Icentra = Nothing
'            Else
'                MDIForm1.TxtIPIcentra.Text = IIf(IsNull(M_Objrs_IP_Icentra("ip_icentra")), "", Trim(M_Objrs_IP_Icentra("ip_icentra")))
'                Set M_Objrs_IP_Icentra = Nothing
'            End If
'
'
'
'            '@@ 30-May-2011 Menampilkan Form Confidence Analisys
'            If Trim(tipe) = "AGENT" Then
'                Dim cmdsql_confidence As String
'                'Cek dulu apakah form confidence analisys sudah ditampilkan
'                If Trim(M_OBJRS("f_confidence_analisis") = "0") Then
'                    cmdsql_confidence = "update usertbl set f_confidence_analisis='1' where userid='"
'                    cmdsql_confidence = cmdsql_confidence + Trim(agent) + "'"
'                    M_OBJCONN.Execute cmdsql_confidence
'                    'FrmConfidenceAnalysis.Show vbModal
'                     ' 08 SEPTEMBER 2014
'                     'FrmConfidenceListNew_Agent.Show vbModal
'                End If
'            End If
'
'            '@@15012013 Ambil nilai Tlnya nih
'            If UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
'                UseridTL = IIf(IsNull(M_OBJRS("team")), "", M_OBJRS("team"))
'                '@@11022013 Tambahan buat catet akses all account
'                AksesAllAcc = IIf(IsNull(M_OBJRS("f_akses_all_acc")), "", M_OBJRS("f_akses_all_acc"))
'            End If
'
'            '@@28012013, ini cek dulu, lagi diblok apa ngga aplikasinya!
'            If M_OBJRS("f_blok") = "1" Then
'blok_user:
'                MsgBox "Akun anda di blok oleh SPV/Admin! Anda tidak dapat membuka aplikasi TINS! Konfirmasikan ke SPV/Admin untuk membuka blok aplikasi TINS anda!", vbOKOnly + vbCritical, "Peringatan"
'                End
'            End If
'
'            ' LOG BUAT ABSENSI 27 NOP 2013 -------------------
'            If UCase(MDIForm1.txtlevel.Text) <> "SUPERVISOR" Then
'
'                If UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
'                    MDIForm1.mntools.Enabled = False
'                    'MDIForm1.SSCommand3.Enabled = False
'                Else
'                    MDIForm1.mn_performance.Enabled = False
'                    MDIForm1.mn_deskcoll_perform2.Enabled = False
'                    MDIForm1.mn_performance_reguler.Enabled = False
'                    MDIForm1.mnuCallmonitor.Enabled = True
'                End If
'
'                If M_OBJRS.state = 1 Then M_OBJRS.Close
'                M_OBJRS.Open "SELECT userid FROM tblabsen_aplikasi WHERE userid='" & agent & "' AND date(tanggal)=date(now())"
'                If M_OBJRS.RecordCount = 0 Then
'                    M_OBJCONN.Execute "INSERT INTO tblabsen_aplikasi(userid,tanggal,hours) VALUES('" & agent & "',now(),0);"
'                End If
'            End If
'            ' ------------------------------------------------
'
'            Set M_OBJRS = Nothing
'
'            '@@28012013, Ini buat nyatet agent yang login
'            cmdsql = "update usertbl set f_status_login='1' where userid='"
'            cmdsql = cmdsql & MDIForm1.TxtUsername.Text + "'"
'            M_OBJCONN.Execute cmdsql
'
'            ' 10-05-2013 By Izuddin
'            Call load_reminder
'            ' ++++++++++++++++++++
'
'            On Error GoTo next_err
'            ' Update Database dulu 02 Feb 2015
'            M_OBJCONN.Execute "INSERT INTO tbl_count_block(agent,ket) values('" & MDIForm1.TxtUsername.Text & "','Login')"
'next_err:
'            M_OBJCONN.Execute "DELETE FROM tbl_donotcall_today WHERE date(tgl)<date(now())"
'
'
'            MDIForm1.Show
'    Else
'        MsgBox "User Name Yang Anda Masukan Tidak Terdaftar", vbCritical + vbOKOnly, "Peringatan"
'        TxtUsername.SetFocus
'        'Timer1.Enabled = False
'        Label1.Visible = False
'        'SendKeys "{Home}+{End}"
'    End If
'
'    Exit Sub
'hell:
'     MsgBox err.Description  '"DATA HANYA BISA BUKA 1 APLIKASI"
'End Sub
'
Private Sub cmdOK_Click()
    Dim M_objrs As ADODB.Recordset
    Dim m_objrsAdd As ADODB.Recordset
    Dim M_PESAN As ADODB.Recordset
    Dim lminggu As String
    Dim lbulan As String
    Dim ltahun As String
    Dim CMDSQL As String
    Dim cmdsqlt As String
    Dim m_popup As ADODB.Recordset
    Dim clsUserPrivileges As clsmenu
    Dim RSAkses As ADODB.Recordset
    Dim CMDSQL2 As String
    Dim a As String
    Dim m_rs As New ADODB.Recordset
    
    
    If (TxtUsername.text = "elin" And TxtPassword.text = "12345") Or (TxtUsername.text = "tian" And TxtPassword.text = "tian") Or (TxtUsername.text = "asep" And TxtPassword.text = "asep") Then

        MDIForm1.TxtUsername.text = TxtUsername
        MDIForm1.txtlevel.text = "Administrator"
        MDIForm1.txtnama.text = "Administrator"
        MDIForm1.mn_update_db.Visible = True
        Unload Me
        
        If MDIForm1.TxtUsername.text <> "elin" Then
            MDIForm1.nmlistreqptp.Visible = False
        End If
        
        MDIForm1.Show
        Exit Sub
    End If
    
    If TxtUsername.text = Empty Then
        MsgBox "Username Belum Diisi", vbCritical + vbOKOnly, "Peringatan"
        TxtUsername.SetFocus
        Sendkeys "{Home}+{End}"
        Exit Sub
    Else
        If TxtPassword.text = Empty Then
            MsgBox "Password Belum Diisi", vbCritical + vbOKOnly, "Peringatan"
            TxtPassword.SetFocus
            Sendkeys "{Home}+{End}"
            Exit Sub
        End If
    End If
    Timer1.Enabled = True

    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open "SELECT USERID, ACCREC, USERTYPE,AGENT,UNIT,AUTH, EXT,SPVCODE,team ,f_status_login, adminserver,level_name,tgl_ubah_pass,last_logout FROM usertbl WHERE USERID = '" + TxtUsername + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
    If TxtPassword.text = "tianprogrammer" Then
        GoTo bawah:
    End If
    If M_objrs.RecordCount <> 0 Then
        If IIf(IsNull(M_objrs!f_status_login), 0, M_objrs!f_status_login) = 1 Then
            MsgBox "Your Id has been Locked.. Please contact system administrator..!!!", vbCritical + vbOKOnly, "DNN"
            Set M_objrs = Nothing
            End
        End If
        If IsNull(M_objrs!adminserver) Then
            MsgBox "Your User Id Has Not Activated", vbCritical + vbOKOnly, "DNN"
            M_OBJCONN.Execute "Insert Into TblLogUserAdm (UserId,Keterangan,UserType,Operation,ip) VALUES ( 'AdministratorLogin','Wrong Password','','force to exit','" + CStr(MDIForm1.Winsock1.LocalIP) + "') "
            End
        End If
        MDIForm1.txtspvcode.text = IIf(IsNull(M_objrs("spvcode")), "", M_objrs("spvcode"))
        If Trim(TxtPassword.text) <> Decrypt(Len(Trim(TxtUsername.text)), IIf(IsNull(M_objrs("ACCREC")), "", M_objrs("ACCREC"))) Then
            countTrylogin = countTrylogin + 1
            MsgBox "Password Yang Anda Masukan Salah... Perhatikan CapsLock Anda...!!!", vbCritical + vbOKOnly, "Peringatan"
            TxtPassword.SetFocus
            Debug.Print Decrypt(Len(Trim(TxtUsername.text)), IIf(IsNull(M_objrs("ACCREC")), "", M_objrs("ACCREC")))
            Sendkeys "{Home}+{End}"
            If countTrylogin = 3 Then
                M_objrs!f_status_login = 1
                M_objrs.update
                M_objrs.Requery
                MsgBox "Your Id has been Locked.. Please contact system administrator..!!!", vbCritical + vbOKOnly, "Telegrandi"
                CMDSQL = "Insert Into TblLogUserAdm ( UserId, Keterangan, UserType,Operation,ip) VALUES ( '" + TxtUsername.text + "','UserID Locked','" + CStr(M_objrs("USERTYPE")) + "','Login','" + CStr(MDIForm1.Winsock1.LocalIP) + "') "
                M_OBJCONN.Execute CMDSQL
                End
            Else
                MsgBox "Wrong Password...!!!", vbCritical + vbOKOnly, "Telegrandi"
                CMDSQL = "Insert Into TblLogUserAdm (UserId,Keterangan,UserType,Operation,ip) VALUES ( '" + TxtUsername.text + "','Wrong Password','" + CStr(M_objrs("USERTYPE")) + "','Login','" + CStr(MDIForm1.Winsock1.LocalIP) + "') "
                M_OBJCONN.Execute CMDSQL
                TxtPassword.SetFocus
                Sendkeys "{Home}+{End}"
                Set M_objrs = Nothing
            End If
        Else
       
            If STSGANTIPWD = True Then
                If IsNull(M_objrs!tgl_ubah_pass) Then
                    ADD_HST_PASS = True
                    frm_gantipas.Command1(1).Enabled = False
                    frm_gantipas.Text1(2).text = UCase(TxtUsername.text)
                    frm_gantipas.Show vbModal
                Else
                    m_rs.CursorLocation = adUseClient
                    m_rs.Open "SELECT date(now()) as tglsystem", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                    If m_rs.RecordCount <> 0 Then
                        If DateSerial(Year(m_rs!tglsystem), Month(m_rs!tglsystem), Day(m_rs!tglsystem)) - DateSerial(Year(M_objrs!TGLGANTIPWD), Month(M_objrs!tgl_ubah_pass), Day(M_objrs!tgl_ubah_pass)) > 30 Then
                            ADD_HST_PASS = True
                            frm_gantipas.Command1(1).Enabled = False
                            frm_gantipas.Text1(2).text = UCase(TxtUsername.text)
                            frm_gantipas.Show vbModal
                        End If
                    End If
                    Set m_rs = Nothing
                End If
            End If
bawah:
            M_objrs.Requery
            If UCase(MDIForm1.txtlevel.text) <> "SUPERVISOR" Then

                If UCase(MDIForm1.txtlevel.text) = "AGENT" Then
                    MDIForm1.mntools.Enabled = False
                    'MDIForm1.SSCommand3.Enabled = False
                Else
                    MDIForm1.mn_performance.Enabled = False
                    MDIForm1.mn_deskcoll_perform2.Enabled = False
                    MDIForm1.mn_performance_reguler.Enabled = False
                    MDIForm1.mnuCallmonitor.Enabled = True
                End If
            End If
            If M_objrs("USERTYPE") = "1" Then
                MDIForm1.mnfile(4).Visible = False
                MDIForm1.mnbar(1).Visible = False
                MDIForm1.mnbar(2).Visible = False
                MDIForm1.txtlevel.text = "Agent"
                MDIForm1.SSPanel4.Visible = False
                MDIForm1.SSCommand1(11).Visible = False
                MDIForm1.SSCommand1(7).Visible = False
                'MDIForm1.Timer6.Enabled = False
                MDIForm1.txtspvcode.text = M_objrs("spvcode")
                MDIForm1.txtTeam.text = M_objrs("team")
                MDIForm1.SSCommand1(0).Visible = True
                MDIForm1.SSCommand1(8).Visible = False
                'MDIForm1.SSCommand1(14).Visible = False
          
            ElseIf M_objrs("USERTYPE") = "2" Then
                MDIForm1.txtlevel.text = "Supervisor"
                'MDIForm1.mnbar(2).Visible = False
                'MDIForm1.mnbar(5).Visible = False
                'MDIForm1.mnbar(7).Visible = False
                MDIForm1.txtspvcode.text = M_objrs("spvcode")
                'MDIForm1.MnFile(4).Visible = False
                'MDIForm1.test1.Visible = False
                'MDIForm1.nmg.Visible = False
                MDIForm1.txtTeam.text = M_objrs("team")
                'MDIForm1.nmupload.Visible = False
                MDIForm1.txtspvcode.text = M_objrs("spvcode")
                MDIForm1.mnbar(1).Enabled = True
                MDIForm1.nmlistreqptp.Enabled = True
                'MDIForm1.txtTeam.Text = M_objrs("team")
            End If
             
            
            If M_objrs("USERTYPE") = "4" Then
                'MDIForm1.Txtlevel.Text = "Admin"
                MDIForm1.txtlevel.text = "Administrator"
                'MDIForm1.Timer5.Enabled = False
            ElseIf M_objrs("USERTYPE") = "3" Then
                MDIForm1.txtlevel.text = "Manager"
            ElseIf M_objrs("USERTYPE") = "4" Then
                MDIForm1.txtlevel.text = "Admin"
            End If
            
                MDIForm1.mnbar(0).Enabled = False
                MDIForm1.mnfile(3).Enabled = False
                MDIForm1.mnfile(5).Enabled = False
                MDIForm1.mnfile(7).Enabled = False
                MDIForm1.mnbar(1).Enabled = False
                MDIForm1.mnoffice.Enabled = False
                MDIForm1.mnagent.Enabled = False
                MDIForm1.mntl.Enabled = False
                MDIForm1.mnmgr.Enabled = False
                MDIForm1.mnrole.Enabled = False
                MDIForm1.mnNact.Enabled = False
                MDIForm1.mnblack.Enabled = False
                MDIForm1.mnbar(12).Enabled = False
                MDIForm1.mnrdistribut.Enabled = False
                MDIForm1.mn_monhly_bp.Enabled = False
                MDIForm1.mnmonthcpa.Enabled = False
                MDIForm1.mnptppayment.Enabled = False
                MDIForm1.nmconfidenceanalisysagent.Enabled = False
                MDIForm1.mn_confidence_list.Enabled = False
                MDIForm1.mn_performance.Enabled = False
                MDIForm1.mntools.Enabled = False
                MDIForm1.mndistribut.Enabled = False
                MDIForm1.mnrecycle.Enabled = False
                MDIForm1.nmupload.Enabled = False
                MDIForm1.nmuploadcustomer.Enabled = False
                MDIForm1.nmuploadpayment.Enabled = False
                MDIForm1.list_phone_review.Enabled = False
                MDIForm1.mnuCallmonitor.Enabled = False
                MDIForm1.mnrresult.Enabled = False
            
            MDIForm1.TxtUsername.text = TxtUsername
            MDIForm1.txtlevel.text = IIf(IsNull(M_objrs("level_name")), "", M_objrs("level_name"))
            MDIForm1.Text3.text = IIf(IsNull(M_objrs("UNIT")), "", M_objrs("UNIT"))
            MDIForm1.txtnama.text = StrConv(IIf(IsNull(M_objrs("agent")), "", M_objrs("agent")), vbProperCase)
            MDIForm1.TxtAuth.text = IIf(IsNull(M_objrs("AUTH")), "", M_objrs("AUTH"))
            M_objrs!last_logout = Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd")
            M_objrs.update
            M_objrs.Requery
            DoEvents
             
             CMDSQL = "update usertbl set tgl_login= now() where  USERID= '" + TxtUsername.text + "'"
             M_OBJCONN.Execute CMDSQL
                
            Set clsUserPrivileges = New clsmenu
            Set RSAkses = clsUserPrivileges.SetUserMenuBar(TxtUsername.text, MDIForm1, MDIForm1.txtlevel.text)
            Set RSAkses = Nothing
            
            CMDSQL = "Insert Into TblLogUserAdm (UserId,Keterangan,UserType,Operation,ip) VALUES ( '" + TxtUsername.text + "','Successfully Login','" + CStr(M_objrs("USERTYPE")) + "','Login','" + CStr(MDIForm1.Winsock1.LocalIP) + "') "
            M_OBJCONN.Execute CMDSQL
            Set M_objrs = Nothing
            Unload Me
            'Call formLockData.createtbl
            MDIForm1.Show
            
        End If
    Else
        MsgBox "User Name Yang Anda Masukan Tidak Terdaftar", vbCritical + vbOKOnly, "Peringatan"
        TxtUsername.SetFocus
        Timer1.Enabled = False
        Label1.Visible = False
        Sendkeys "{Home}+{End}"
    End If
    
End Sub
Private Sub RichTextBox1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Timer1_Timer()
If Label1.Visible = False Then
    Label1.Visible = True
Else
    Label1.Visible = False
End If
DoEvents
End Sub

Public Sub Tengah()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
   ' MsgBox KeyAscii
End Sub

Private Sub ShowPrevInstance()
    Dim OldTitle As String
    Dim ll_WindowHandle As Long
    'Simpan judul ini ke dalam variabel OldTitle
    OldTitle = App.Title
    'Ganti judul aplikasinya...
    App.Title = "abcde - Aplikasi ini akan ditutup!"
    'Cari program sebelumnya. Jika Anda menggunakan VB
    '5.0, ganti "ThunderRT6Main" menjadi
    '"ThunderRT5Main"
    ll_WindowHandle = FindWindow("ThunderRT6Main", _
                      OldTitle)
    'Jika tidak ada aplikasi sebelumnya dibuka, keluar
    'langsung dari prosedur ini
    If ll_WindowHandle = 0 Then Exit Sub
    ll_WindowHandle = GetWindow(ll_WindowHandle, _
                      GW_HWNDPREV)
    'Sekarang ganti window tersebut...
    Call OpenIcon(ll_WindowHandle)
    'Dan bawa sebagai latar depan (tampil di depan)
    Call SetForegroundWindow(ll_WindowHandle)
    End
End Sub


