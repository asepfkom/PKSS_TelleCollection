VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Begin VB.Form FrmFollowUpProblemHeadset 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Follow Up Problem Headset"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5745
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   4320
      TabIndex        =   24
      Top             =   5520
      Width           =   1275
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   3120
      TabIndex        =   23
      Top             =   5520
      Width           =   1275
   End
   Begin VB.TextBox TxtKetSolusi 
      Appearance      =   0  'Flat
      Height          =   645
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   4560
      Width           =   3975
   End
   Begin VB.ComboBox CmbStatusSolusi 
      Height          =   315
      ItemData        =   "FrmFollowUpProblemHeadset.frx":0000
      Left            =   1440
      List            =   "FrmFollowUpProblemHeadset.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   4140
      Width           =   2835
   End
   Begin VB.ComboBox CmbJenisSolusi 
      Height          =   315
      ItemData        =   "FrmFollowUpProblemHeadset.frx":0020
      Left            =   1440
      List            =   "FrmFollowUpProblemHeadset.frx":002D
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   3840
      Width           =   2835
   End
   Begin VB.TextBox TxtSolusiOleh 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   16
      Top             =   3480
      Width           =   2715
   End
   Begin VB.TextBox txtketerangan 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   645
      Left            =   1740
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   2040
      Width           =   3975
   End
   Begin VB.TextBox TxtJenisKerusakan 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   645
      Left            =   1740
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   1380
      Width           =   3975
   End
   Begin VB.TextBox TxtNama 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1740
      TabIndex        =   9
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox TxtUserid 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1740
      TabIndex        =   8
      Top             =   780
      Width           =   1935
   End
   Begin VB.TextBox TxtTglPengajuan 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1740
      TabIndex        =   7
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox TxtID 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1740
      TabIndex        =   6
      Top             =   180
      Width           =   1935
   End
   Begin TDBDate6Ctl.TDBDate TxtTglSolusi 
      Height          =   315
      Left            =   1440
      TabIndex        =   14
      Top             =   3120
      Width           =   1260
      _Version        =   65536
      _ExtentX        =   2222
      _ExtentY        =   556
      Calendar        =   "FrmFollowUpProblemHeadset.frx":0069
      Caption         =   "FrmFollowUpProblemHeadset.frx":0181
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmFollowUpProblemHeadset.frx":01ED
      Keys            =   "FrmFollowUpProblemHeadset.frx":020B
      Spin            =   "FrmFollowUpProblemHeadset.frx":0269
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd/mm/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd/mm/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__/__/____"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   1.12794198814265E-317
      CenturyMode     =   0
   End
   Begin VB.Label Label7 
      Caption         =   "Keterangan:"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   4560
      Width           =   1275
   End
   Begin VB.Label Label6 
      Caption         =   "Status solusi:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   4200
      Width           =   1275
   End
   Begin VB.Label Label5 
      Caption         =   "Jenis solusi:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3840
      Width           =   1275
   End
   Begin VB.Label Label4 
      Caption         =   "Solusi oleh:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3480
      Width           =   1275
   End
   Begin VB.Label Label3 
      Caption         =   "Tanggal solusi:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3120
      Width           =   1275
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   660
      X2              =   5700
      Y1              =   2940
      Y2              =   2940
   End
   Begin VB.Label Label2 
      Caption         =   "Solusi:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   12
      Top             =   2820
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Keterangan:"
      Height          =   255
      Index           =   5
      Left            =   180
      TabIndex        =   5
      Top             =   2100
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Jenis kerusakan:"
      Height          =   255
      Index           =   4
      Left            =   180
      TabIndex        =   4
      Top             =   1380
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Nama:"
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   3
      Top             =   1080
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Userid:"
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   2
      Top             =   780
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Tanggal pengajuan:"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Top             =   480
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "ID data"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   1515
   End
End
Attribute VB_Name = "FrmFollowUpProblemHeadset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBatal_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim VSAVE As Boolean
    Dim CMDSQL As String
    Dim pesan As String
    Dim M_objrs As ADODB.Recordset
    
    On Error GoTo salah
    
    VSAVE = True
    VSAVE = VSAVE And TxtTglSolusi.ValueIsNull = False
    VSAVE = VSAVE And TxtSolusiOleh.Text <> Empty
    VSAVE = VSAVE And CmbJenisSolusi.Text <> Empty
    VSAVE = VSAVE And CmbStatusSolusi.Text <> Empty
    
    If VSAVE Then
        CMDSQL = "update tbl_problem_headset set tgl_solusi='"
        CMDSQL = CMDSQL + Format(TxtTglSolusi.Value, "yyyy-mm-dd") + "',solusi_by='"
        CMDSQL = CMDSQL + TxtSolusiOleh.Text + "',jenis_solusi='"
        CMDSQL = CMDSQL + CmbJenisSolusi.Text + "',status_problem='"
        CMDSQL = CMDSQL + CmbStatusSolusi.Text + "',solusi='"
        CMDSQL = CMDSQL + IIf(IsNull(TxtKetSolusi.Text), "", TxtKetSolusi.Text) + "' where id='"
        CMDSQL = CMDSQL + CStr(txtId.Text) + "'"
        M_OBJCONN.Execute CMDSQL
        
        pesan = "Pesan dibuat otomatis oleh system" & vbCrLf
        pesan = pesan & "-----------------------------------------" & vbCrLf
        pesan = pesan & "Status Request Headset Tanggal: " & TxtTglPengajuan.Text & " ID:" & txtId.Text & vbCrLf
        pesan = pesan & "Request oleh: " & txtUserId.Text & "-" & txtNama.Text & vbCrLf
        pesan = pesan & "Kerusakan: " & vbCrLf & TxtJenisKerusakan.Text & vbCrLf & vbCrLf
        pesan = pesan & "===FOLLOW UP ====" & vbCrLf
        pesan = pesan & "Tanggal: " & Format(TxtTglSolusi.Value, "yyyy-mm-dd") & vbCrLf
        pesan = pesan & "Oleh: " & TxtSolusiOleh.Text & vbCrLf
        pesan = pesan & "Status: " & CmbStatusSolusi.Text & vbCrLf
        pesan = pesan & "Solusi: " & CmbJenisSolusi.Text & vbCrLf
        pesan = pesan & "Keterangan: " & vbCrLf
        pesan = pesan & IIf(IsNull(TxtKetSolusi.Text), "", TxtKetSolusi.Text)
        
        '@@18012013 Kirim Pesan
        'Ke TL nya
        CMDSQL = "select team from usertbl where userid='"
        CMDSQL = CMDSQL + txtUserId.Text + "'"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_objrs.RecordCount > 0 Then
            CMDSQL = "insert into msgtbl "
            CMDSQL = CMDSQL + "( recipient, datetime, sender, sentfrom, msg) values ('"
            CMDSQL = CMDSQL + M_objrs("team") + "','"
            CMDSQL = CMDSQL + Format(Now(), "yyyymmdd") + "','"
            CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "','"
            CMDSQL = CMDSQL + CStr(MDIForm1.Winsock1.LocalIP) + "','"
            CMDSQL = CMDSQL + pesan + "')"
            M_OBJCONN.Execute CMDSQL
        End If
        
        Set M_objrs = Nothing
        
        'Kirim Ke agent nya
        CMDSQL = "insert into msgtbl "
        CMDSQL = CMDSQL + "( recipient, datetime, sender, sentfrom, msg) values ('"
        CMDSQL = CMDSQL + txtUserId.Text + "','"
        CMDSQL = CMDSQL + Format(Now(), "yyyymmdd") + "','"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "','"
        CMDSQL = CMDSQL + CStr(MDIForm1.Winsock1.LocalIP) + "','"
        CMDSQL = CMDSQL + pesan + "')"
        M_OBJCONN.Execute CMDSQL
        
        'Kirim ke admin/manager/supervisor
        CMDSQL = "select * from usertbl where usertype in ('11','20','25') "
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_objrs.RecordCount > 0 Then
            While Not M_objrs.EOF
                CMDSQL = "insert into msgtbl "
                CMDSQL = CMDSQL + "( recipient, datetime, sender, sentfrom, msg) values ('"
                CMDSQL = CMDSQL + M_objrs("userid") + "','"
                CMDSQL = CMDSQL + Format(Now(), "yyyymmdd") + "','"
                CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "','"
                CMDSQL = CMDSQL + CStr(MDIForm1.Winsock1.LocalIP) + "','"
                CMDSQL = CMDSQL + pesan + "')"
                M_OBJCONN.Execute CMDSQL
                M_objrs.MoveNext
            Wend
        End If
        
        
        MsgBox "Data berhasil di update!", vbOKOnly + vbInformation, "Informasi"
        FrmListProblemHeadset.IsiData
        Unload Me
    End If
    Exit Sub
salah:
    MsgBox "Mohon maaf ada error: " & err.Description, vbOKOnly + vbExclamation, "Peringatan"
    
End Sub
