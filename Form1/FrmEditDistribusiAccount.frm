VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Begin VB.Form FrmEditDistribusiAccount 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit Distribusi Account"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3735
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   435
      Left            =   2340
      TabIndex        =   11
      Top             =   1740
      Width           =   1215
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      Height          =   435
      Left            =   1080
      TabIndex        =   10
      Top             =   1740
      Width           =   1215
   End
   Begin VB.TextBox TxtAgent 
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox TxtID 
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   180
      Width           =   1335
   End
   Begin TDBDate6Ctl.TDBDate TxtTglAwal 
      Height          =   315
      Left            =   1200
      TabIndex        =   4
      Top             =   900
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   556
      Calendar        =   "FrmEditDistribusiAccount.frx":0000
      Caption         =   "FrmEditDistribusiAccount.frx":0118
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmEditDistribusiAccount.frx":0184
      Keys            =   "FrmEditDistribusiAccount.frx":01A2
      Spin            =   "FrmEditDistribusiAccount.frx":0200
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   12648384
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "yyyy-mm-dd"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "yyyy-mm-dd"
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
      Text            =   "____-__-__"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   1.12794198814265E-317
      CenturyMode     =   0
   End
   Begin TDBTime6Ctl.TDBTime TxtWaktuAwal 
      Height          =   315
      Left            =   2715
      TabIndex        =   5
      Top             =   900
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   556
      Caption         =   "FrmEditDistribusiAccount.frx":0228
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmEditDistribusiAccount.frx":0294
      Spin            =   "FrmEditDistribusiAccount.frx":02E4
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   12648384
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "hh:nn"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "hh:nn"
      HighlightText   =   0
      Hour12Mode      =   1
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxTime         =   0.99999
      MidnightMode    =   0
      MinTime         =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__:__"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   1.02960316199441E-317
   End
   Begin TDBDate6Ctl.TDBDate TxtTglAkhir 
      Height          =   315
      Left            =   1200
      TabIndex        =   6
      Top             =   1260
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   556
      Calendar        =   "FrmEditDistribusiAccount.frx":030C
      Caption         =   "FrmEditDistribusiAccount.frx":0424
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmEditDistribusiAccount.frx":0490
      Keys            =   "FrmEditDistribusiAccount.frx":04AE
      Spin            =   "FrmEditDistribusiAccount.frx":050C
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   12648384
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "yyyy-mm-dd"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "yyyy-mm-dd"
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
      Text            =   "____-__-__"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   1.12794198814265E-317
      CenturyMode     =   0
   End
   Begin TDBTime6Ctl.TDBTime TxtWaktuAkhir 
      Height          =   315
      Left            =   2715
      TabIndex        =   7
      Top             =   1260
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   556
      Caption         =   "FrmEditDistribusiAccount.frx":0534
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmEditDistribusiAccount.frx":05A0
      Spin            =   "FrmEditDistribusiAccount.frx":05F0
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   12648384
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "hh:nn"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "hh:nn"
      HighlightText   =   0
      Hour12Mode      =   1
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxTime         =   0.99999
      MidnightMode    =   0
      MinTime         =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__:__"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   1.02960316199441E-317
   End
   Begin VB.Label Label11 
      Caption         =   "Waktu Akhir:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label12 
      Caption         =   "Waktu Awal:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Agent:"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "ID:"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   735
   End
End
Attribute VB_Name = "FrmEditDistribusiAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBatal_Click()
    Unload Me
End Sub

Private Sub CmdSimpan_Click()
    Dim CMDSQL As String
    Dim M_ObjrsWaktuServer As ADODB.Recordset
    Dim WaktuServer As String
    Dim Tanggal1 As String
    Dim Tanggal2 As String
    
    'Cek waktu server
    CMDSQL = "select now()"
    Set M_ObjrsWaktuServer = New ADODB.Recordset
    M_ObjrsWaktuServer.CursorLocation = adUseClient
    M_ObjrsWaktuServer.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    WaktuServer = Format(M_ObjrsWaktuServer(0), "m/dd/yyyy hh:nn:ss")
    
    If TxtTglAwal.ValueIsNull = True Or _
       TxtWaktuAwal.ValueIsNull = True Or _
       TxtTglAkhir.ValueIsNull = True Or _
       TxtWaktuAkhir.ValueIsNull = True Then
        MsgBox "Waktu tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    'Cek tanggal awal tidak boleh lebih besar dari tanggal akhir
    Tanggal1 = Format(TxtTglAwal.Value, "m/dd/yyyy") & " " & Format(TxtWaktuAwal.Value, "hh:nn")
    Tanggal2 = Format(TxtTglAkhir.Value, "m/dd/yyyy") & " " & Format(TxtWaktuAkhir.Value, "hh:nn")
     
    'Cek jika waktu akhir server lebih kecil dari waktu server sekarang
    If CDate(Tanggal2) < CDate(WaktuServer) Then
        MsgBox "Waktu akhir tidak boleh lebih kecil dari waktu server! Waktu Server sekarng: " & Format(WaktuServer, "yyyy-mm-dd hh:nn:ss"), vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
     
    If CDate(Tanggal1) > CDate(Tanggal2) Then
        MsgBox "Tanggal awal tidak boleh lebih besar dari tanggal akhir!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    'Update
    CMDSQL = "update mandiri.tbl_distribusi_account set waktu_awal='"
    CMDSQL = CMDSQL & Format(TxtTglAwal.Value, "yyyy-mm-dd") & " " & Format(TxtWaktuAwal.Value, "hh:nn:ss") & "',waktu_akhir='"
    CMDSQL = CMDSQL & Format(TxtTglAkhir.Value, "yyyy-mm-dd") & " " & Format(TxtWaktuAkhir.Value, "hh:nn:ss") & "',log_distribusi='"
    CMDSQL = CMDSQL & MDIForm1.TxtUsername.Text & "',log_tgl_distribusi=now() where id='"
    CMDSQL = CMDSQL & CStr(txtId.Text) & "'"
    M_OBJCONN.Execute CMDSQL
    
    'Update juga status accountnya, biar bisa langsung di refresh
    CMDSQL = "update mandiri.usertbl set f_pesanresetauto='1' where userid='"
    CMDSQL = CMDSQL & TxtAgent.Text & "'"
    M_OBJCONN.Execute CMDSQL
        
    FrmDistribusiAcc.CariAgent
    MsgBox "Data berhasil diupdate!", vbOKOnly + vbInformation, "Informasi"
    Unload Me
End Sub
