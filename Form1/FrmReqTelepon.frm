VERSION 5.00
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Begin VB.Form FrmReqTelepon 
   Caption         =   "Req.Num.Telp"
   ClientHeight    =   2055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6435
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   6435
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtNoTelp 
      Height          =   315
      Left            =   1680
      TabIndex        =   22
      Top             =   840
      Width           =   1455
   End
   Begin VB.ComboBox CmbRequestDi 
      Height          =   315
      ItemData        =   "FrmReqTelepon.frx":0000
      Left            =   1680
      List            =   "FrmReqTelepon.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   480
      Width           =   3015
   End
   Begin VB.CommandButton CmdKeluar 
      Caption         =   "&Keluar"
      Height          =   375
      Left            =   5220
      TabIndex        =   17
      Top             =   1515
      Width           =   1095
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   4200
      TabIndex        =   16
      Top             =   1515
      Width           =   1035
   End
   Begin TDBMask6Ctl.TDBMask TxtHome1 
      Height          =   315
      Left            =   1620
      TabIndex        =   3
      Top             =   4680
      Visible         =   0   'False
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   556
      Caption         =   "FrmReqTelepon.frx":0045
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmReqTelepon.frx":00B1
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "999999999999999"
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      LookupMode      =   0
      LookupTable     =   ""
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "_______________"
      Value           =   ""
   End
   Begin VB.TextBox TxtCustid 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin TDBMask6Ctl.TDBMask TxtHome2 
      Height          =   315
      Left            =   1620
      TabIndex        =   5
      Top             =   5040
      Visible         =   0   'False
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   556
      Caption         =   "FrmReqTelepon.frx":00F3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmReqTelepon.frx":015F
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "999999999999999"
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      LookupMode      =   0
      LookupTable     =   ""
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "_______________"
      Value           =   ""
   End
   Begin TDBMask6Ctl.TDBMask TxtOffice1 
      Height          =   315
      Left            =   1620
      TabIndex        =   7
      Top             =   5400
      Visible         =   0   'False
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   556
      Caption         =   "FrmReqTelepon.frx":01A1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmReqTelepon.frx":020D
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "999999999999999"
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      LookupMode      =   0
      LookupTable     =   ""
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "_______________"
      Value           =   ""
   End
   Begin TDBMask6Ctl.TDBMask TxtOffice2 
      Height          =   315
      Left            =   1620
      TabIndex        =   8
      Top             =   5760
      Visible         =   0   'False
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   556
      Caption         =   "FrmReqTelepon.frx":024F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmReqTelepon.frx":02BB
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "999999999999999"
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      LookupMode      =   0
      LookupTable     =   ""
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "_______________"
      Value           =   ""
   End
   Begin TDBMask6Ctl.TDBMask TxtMobile1 
      Height          =   315
      Left            =   1620
      TabIndex        =   11
      Top             =   6120
      Visible         =   0   'False
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   556
      Caption         =   "FrmReqTelepon.frx":02FD
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmReqTelepon.frx":0369
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "999999999999999"
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      LookupMode      =   0
      LookupTable     =   ""
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "_______________"
      Value           =   ""
   End
   Begin TDBMask6Ctl.TDBMask TxtMobile2 
      Height          =   315
      Left            =   1620
      TabIndex        =   12
      Top             =   6480
      Visible         =   0   'False
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   556
      Caption         =   "FrmReqTelepon.frx":03AB
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmReqTelepon.frx":0417
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "999999999999999"
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      LookupMode      =   0
      LookupTable     =   ""
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "_______________"
      Value           =   ""
   End
   Begin TDBMask6Ctl.TDBMask TxtEcPhone 
      Height          =   315
      Left            =   1620
      TabIndex        =   15
      Top             =   6840
      Visible         =   0   'False
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   556
      Caption         =   "FrmReqTelepon.frx":0459
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmReqTelepon.frx":04C5
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "999999999999999"
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      LookupMode      =   0
      LookupTable     =   ""
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "_______________"
      Value           =   ""
   End
   Begin TDBMask6Ctl.TDBMask TxtNoTelp1 
      Height          =   315
      Left            =   4680
      TabIndex        =   18
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   556
      Caption         =   "FrmReqTelepon.frx":0507
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmReqTelepon.frx":0573
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "999999999999999"
      HighlightText   =   0
      IMEMode         =   0
      IMEStatus       =   0
      LookupMode      =   0
      LookupTable     =   ""
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   " "
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "               "
      Value           =   ""
   End
   Begin VB.Label Label12 
      Caption         =   "Request di:"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   480
      Width           =   1395
   End
   Begin VB.Label Label11 
      Caption         =   "No.Telepon"
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   900
      Width           =   1395
   End
   Begin VB.Label Label8 
      Caption         =   "EC Phone"
      Height          =   195
      Left            =   60
      TabIndex        =   14
      Top             =   6900
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Label7 
      Caption         =   "Additional Mobile 2"
      Height          =   195
      Left            =   60
      TabIndex        =   13
      Top             =   6540
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Label6 
      Caption         =   "Additional Mobile 1"
      Height          =   195
      Left            =   60
      TabIndex        =   10
      Top             =   6180
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Label5 
      Caption         =   "Additional Officeno 2"
      Height          =   195
      Left            =   60
      TabIndex        =   9
      Top             =   5820
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Additional Officeno 1"
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   5460
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Additional Home 2"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   5100
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Label2 
      Caption         =   "Additional Home 1"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   4740
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Custid:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1395
   End
End
Attribute VB_Name = "FrmReqTelepon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdkeluar_Click()
    Me.Hide
End Sub

Private Sub CmdSimpan_Click()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim strsql As String
    
    Dim home1 As String
    Dim home2 As String
    Dim office1 As String
    Dim office2 As String
    Dim mobile1 As String
    Dim mobile2 As String
    Dim ec As String
        
    '@@17042012, Di Remarks dulu diganti dengan kategori
    If TxtNoTelp.text = "" Then
        MsgBox "Nomor Telepon tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If CmbRequestDi.text = "" Then
        MsgBox "Jenis Request Number tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    'Update Nomor Telepon
    CMDSQL = "insert into tblrequestadditionalphone (custid,"
    CMDSQL = CMDSQL + " request_number,agent,tglreq,jenis) values ('"
    CMDSQL = CMDSQL + TxtCustid.text + "','"
    CMDSQL = CMDSQL + IIf(IsNull(TxtNoTelp.text), "", CStr(TxtNoTelp.text)) + "','"
    CMDSQL = CMDSQL + IIf(IsNull(MDIForm1.TxtUsername.text), "", MDIForm1.TxtUsername.text) + "','"
    CMDSQL = CMDSQL + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd hh:mm:ss")) + "','"
    CMDSQL = CMDSQL + Trim(CmbRequestDi.text) + "')"
    M_OBJCONN.Execute CMDSQL
    
    'jejaktian06042016
    If CmbRequestDi.text = "AddHome1" Then
        FrmCC_Colection.txtadd_phone(0).text = TxtNoTelp.text
        'FrmCC_Colection.txtadd_phone(6).text = Left(TxtNoTelp.text, Len(TxtNoTelp.text) - 3) & "###"
        '====asep==='
        FrmCC_Colection.txtadd_phone(6).text = Left(TxtNoTelp.text, Len(TxtNoTelp.text))
        '=========='
    ElseIf CmbRequestDi.text = "AddOffice1" Then
        FrmCC_Colection.txtadd_phone(1).text = TxtNoTelp.text
        FrmCC_Colection.txtadd_phone(5).text = Left(TxtNoTelp.text, Len(TxtNoTelp.text))
    ElseIf CmbRequestDi.text = "AddMobile1" Then
        FrmCC_Colection.txtadd_phone(2).text = TxtNoTelp.text
        FrmCC_Colection.txtadd_phone(4).text = Left(TxtNoTelp.text, Len(TxtNoTelp.text))
    ElseIf CmbRequestDi.text = "AddOtherphone" Then
        FrmCC_Colection.txtadd_phone(3).text = TxtNoTelp.text
        FrmCC_Colection.txtadd_phone(7).text = Left(TxtNoTelp.text, Len(TxtNoTelp.text))
    End If
    '======================================================
    
'    'Update buat ngasih tanda ke TL/SPV/Admin
'    CMDSQL = "update usertbl set f_req_number='1' where userid in ("
'    CMDSQL = CMDSQL + "select team from usertbl where userid='"
'    CMDSQL = CMDSQL + mdiform1.txtusername.text + "') or userid in (select userid from usertbl where "
'    CMDSQL = CMDSQL + "usertype='20' or usertype='25' or usertype='11') "
'    M_OBJCONN.Execute CMDSQL

    '@@07-08-2012 Kirim Via Form Pesan
    'Kirim Ke Semua TL/SPV
    'CMDSQL = "select userid from usertbl where usertype in ('6','11','25','20') "
    CMDSQL = "select userid from usertbl where userid in ("
    CMDSQL = CMDSQL + "select team from usertbl where userid='"
    CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "') "
    CMDSQL = CMDSQL + " and userid is not null "
    Set M_Objrs_CariTL = New ADODB.Recordset
    M_Objrs_CariTL.CursorLocation = adUseClient
    M_Objrs_CariTL.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_Objrs_CariTL.RecordCount > 0 Then
        Remarks = "Ada Request Number!" & vbCrLf
        Remarks = Remarks + "-------------------------------------------------" & vbCrLf
        Remarks = Remarks + " Custid: " & TxtCustid.text & vbCrLf
        Remarks = Remarks + " Agent:  " & MDIForm1.TxtUsername.text & vbCrLf
        Remarks = Remarks + " Tgl.Request: " & CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd hh:mm:ss"))
        
        While Not M_Objrs_CariTL.EOF
            CMDSQL = "insert into msgtbl "
            CMDSQL = CMDSQL + "( recipient, datetime, sender, sentfrom, msg) values ('"
            CMDSQL = CMDSQL + CStr(Trim(M_Objrs_CariTL("userid"))) + "','"
            CMDSQL = CMDSQL + Format(Now(), "yyyymmdd") + "','"
            CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "','"
            CMDSQL = CMDSQL + CStr(MDIForm1.Winsock1.LocalIP) + "','"
            CMDSQL = CMDSQL + Remarks + "')"
            M_OBJCONN.Execute CMDSQL
            M_Objrs_CariTL.MoveNext
        Wend
    End If
    CMDSQL = "INSERT INTO tblnotif_info "
                CMDSQL = CMDSQL & "( type_notif,notif_from) values ('"
                CMDSQL = CMDSQL & "address','" & Trim$(MDIForm1.TxtUsername.text) & "')"
    M_OBJCONN.Execute CMDSQL
    MsgBox "Request berhasil dikirim!", vbOKOnly + vbInformation, "Informasi"
    Unload Me
End Sub

Private Sub Form_Load()
TxtCustid.text = FrmCC_Colection.lblCustId.text
End Sub
