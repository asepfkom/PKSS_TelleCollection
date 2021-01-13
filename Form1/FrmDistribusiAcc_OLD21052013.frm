VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmDistribusiAcc 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Manage distribusi account"
   ClientHeight    =   10365
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15660
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10365
   ScaleWidth      =   15660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdFilterExcel 
      Caption         =   "&Filter dari Excel"
      Height          =   315
      Left            =   3780
      TabIndex        =   54
      Top             =   0
      Width           =   3075
   End
   Begin VB.ComboBox CmbStatusCollBersama 
      Height          =   315
      Left            =   9960
      Style           =   2  'Dropdown List
      TabIndex        =   53
      Top             =   360
      Width           =   1635
   End
   Begin VB.ComboBox CmbAgentCollBersama 
      Height          =   315
      Left            =   7980
      Style           =   2  'Dropdown List
      TabIndex        =   51
      Top             =   360
      Width           =   1395
   End
   Begin VB.CommandButton CmdFormClaimAccount 
      BackColor       =   &H0000C0C0&
      Caption         =   "Form Claim Account..."
      Height          =   435
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   6120
      Width           =   2355
   End
   Begin VB.CommandButton CmdUncekallAgent 
      Caption         =   "UnCek All"
      Height          =   375
      Left            =   10740
      TabIndex        =   43
      Top             =   8940
      Width           =   1755
   End
   Begin VB.CommandButton CmdCekAllAgent 
      Caption         =   "Cek All"
      Height          =   375
      Left            =   10740
      TabIndex        =   42
      Top             =   8520
      Width           =   1755
   End
   Begin VB.ComboBox CmbStatusAcc 
      Height          =   315
      Left            =   5220
      Style           =   2  'Dropdown List
      TabIndex        =   41
      Top             =   360
      Width           =   1635
   End
   Begin VB.ComboBox CmbAgent 
      Height          =   315
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   7080
      Width           =   2235
   End
   Begin VB.CommandButton CmdBukaAccount 
      Caption         =   "Buka account terkunci..."
      Height          =   435
      Left            =   10440
      TabIndex        =   36
      Top             =   6120
      Width           =   2355
   End
   Begin VB.CommandButton CmdKembalikanAgent 
      Caption         =   "Kembalikan Ke Agent lama..."
      Height          =   435
      Left            =   8040
      TabIndex        =   37
      Top             =   6120
      Width           =   2355
   End
   Begin VB.ComboBox CmbFilterAcc 
      Height          =   315
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   360
      Width           =   1995
   End
   Begin VB.TextBox TxtJmlhAcc 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   27
      Text            =   "0"
      Top             =   4740
      Width           =   915
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "Clear"
      Height          =   315
      Left            =   14520
      TabIndex        =   25
      Top             =   360
      Width           =   1035
   End
   Begin VB.CommandButton CmdCari 
      Caption         =   "&Cari"
      Height          =   315
      Left            =   14520
      TabIndex        =   24
      Top             =   60
      Width           =   1035
   End
   Begin VB.TextBox TxtCariNama 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   12660
      TabIndex        =   23
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox TxtCariCustid 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   12660
      TabIndex        =   21
      Top             =   60
      Width           =   1815
   End
   Begin VB.CommandButton CmdUncekAll 
      Caption         =   "UnCek all"
      Height          =   315
      Left            =   120
      TabIndex        =   19
      Top             =   360
      Width           =   1035
   End
   Begin VB.CommandButton CmdCekAllAcc 
      Caption         =   "Cek all"
      Height          =   315
      Left            =   120
      TabIndex        =   18
      Top             =   60
      Width           =   1035
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   2580
      TabIndex        =   17
      Top             =   4800
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox TxtJmlhAgent 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   1620
      TabIndex        =   16
      Text            =   "0"
      Top             =   9900
      Width           =   915
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "Edit..."
      Height          =   375
      Left            =   10740
      TabIndex        =   15
      Top             =   7440
      Width           =   1755
   End
   Begin VB.CommandButton CmdHapusAgent 
      Caption         =   "&Hapus Agent"
      Height          =   375
      Left            =   10740
      TabIndex        =   14
      Top             =   7920
      Width           =   1755
   End
   Begin VB.CommandButton CmdProses 
      BackColor       =   &H0000FF00&
      Caption         =   "&Proses..."
      Height          =   435
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      Width           =   1875
   End
   Begin VB.CommandButton CmdLihatListAgent 
      Caption         =   "Lihat list agent..."
      Height          =   435
      Left            =   10740
      TabIndex        =   3
      Top             =   5160
      Width           =   1755
   End
   Begin VB.TextBox TxtAgent 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3600
      TabIndex        =   2
      Top             =   5220
      Width           =   7035
   End
   Begin MSComctlLib.ListView LvAcc 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   15435
      _ExtentX        =   27226
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView LvAgent 
      Height          =   2415
      Left            =   60
      TabIndex        =   6
      Top             =   7440
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin TDBDate6Ctl.TDBDate TxtTglAwal 
      Height          =   315
      Left            =   1260
      TabIndex        =   28
      Top             =   5640
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   556
      Calendar        =   "FrmDistribusiAcc.frx":0000
      Caption         =   "FrmDistribusiAcc.frx":0118
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmDistribusiAcc.frx":0184
      Keys            =   "FrmDistribusiAcc.frx":01A2
      Spin            =   "FrmDistribusiAcc.frx":0200
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
      Left            =   2775
      TabIndex        =   29
      Top             =   5640
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   556
      Caption         =   "FrmDistribusiAcc.frx":0228
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmDistribusiAcc.frx":0294
      Spin            =   "FrmDistribusiAcc.frx":02E4
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
      Left            =   4800
      TabIndex        =   30
      Top             =   5640
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   556
      Calendar        =   "FrmDistribusiAcc.frx":030C
      Caption         =   "FrmDistribusiAcc.frx":0424
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmDistribusiAcc.frx":0490
      Keys            =   "FrmDistribusiAcc.frx":04AE
      Spin            =   "FrmDistribusiAcc.frx":050C
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
      Left            =   6315
      TabIndex        =   31
      Top             =   5640
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   556
      Caption         =   "FrmDistribusiAcc.frx":0534
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmDistribusiAcc.frx":05A0
      Spin            =   "FrmDistribusiAcc.frx":05F0
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
   Begin VB.Label Label20 
      Caption         =   "Status:"
      Height          =   195
      Left            =   9360
      TabIndex        =   52
      Top             =   420
      Width           =   555
   End
   Begin VB.Label Label19 
      Caption         =   "Agent AWAL:"
      Height          =   195
      Left            =   6960
      TabIndex        =   50
      Top             =   420
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Filter Account Sedang di Collect Bersama:"
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
      Left            =   6960
      TabIndex        =   49
      Top             =   60
      Width           =   4215
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   11640
      X2              =   11640
      Y1              =   0
      Y2              =   720
   End
   Begin VB.Label LblWaktuServer 
      BackColor       =   &H000080FF&
      Caption         =   "<Waktu Server>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9540
      TabIndex        =   48
      Top             =   5640
      Width           =   2355
   End
   Begin VB.Label Label18 
      BackColor       =   &H000040C0&
      Caption         =   "Waktu Server Saat ini:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7200
      TabIndex        =   47
      Top             =   5640
      Width           =   2355
   End
   Begin VB.Label Label17 
      Caption         =   "Status Acc:"
      Height          =   195
      Left            =   4380
      TabIndex        =   45
      Top             =   420
      Width           =   1035
   End
   Begin VB.Label Label16 
      Caption         =   "Agent:"
      Height          =   195
      Left            =   1260
      TabIndex        =   44
      Top             =   420
      Width           =   555
   End
   Begin VB.Label Label15 
      Caption         =   "AND"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3900
      TabIndex        =   40
      Top             =   420
      Width           =   435
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   3300
      X2              =   3300
      Y1              =   7080
      Y2              =   7440
   End
   Begin VB.Label Label14 
      Caption         =   "Filter Agent:"
      Height          =   195
      Left            =   60
      TabIndex        =   38
      Top             =   7140
      Width           =   915
   End
   Begin VB.Label Label13 
      Caption         =   "Filter Account:"
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
      Left            =   1260
      TabIndex        =   34
      Top             =   60
      Width           =   2055
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   6900
      X2              =   6900
      Y1              =   0
      Y2              =   720
   End
   Begin VB.Label Label12 
      Caption         =   "Waktu Awal:"
      Height          =   195
      Left            =   180
      TabIndex        =   33
      Top             =   5700
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "Waktu Akhir:"
      Height          =   195
      Left            =   3780
      TabIndex        =   32
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Jumlah Account:"
      Height          =   195
      Left            =   180
      TabIndex        =   26
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Cari Nama:"
      Height          =   195
      Left            =   11700
      TabIndex        =   22
      Top             =   420
      Width           =   1035
   End
   Begin VB.Label Label8 
      Caption         =   "Cari Custid:"
      Height          =   195
      Left            =   11700
      TabIndex        =   20
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label LblStatusAcc 
      Caption         =   "<none>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12480
      TabIndex        =   13
      Top             =   7080
      Width           =   2655
   End
   Begin VB.Label Label7 
      Caption         =   "Status Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10980
      TabIndex        =   12
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label LblNama 
      Caption         =   "<none>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   11
      Top             =   7080
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "Nama:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7500
      TabIndex        =   10
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label LblCustid 
      Caption         =   "<none>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   7080
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Custid terpilih:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3420
      TabIndex        =   8
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Jumlah Data Agent:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   9960
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      Caption         =   $"FrmDistribusiAcc.frx":0618
      Height          =   315
      Left            =   60
      TabIndex        =   5
      Top             =   6720
      Width           =   15375
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   60
      X2              =   15420
      Y1              =   6660
      Y2              =   6660
   End
   Begin VB.Label Label2 
      Caption         =   "Agent yang boleh mengakses account di atas:"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   5220
      Width           =   3735
   End
End
Attribute VB_Name = "FrmDistribusiAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub HeaderAccount()
    LvAcc.ColumnHeaders.ADD 1, , "Custid", 2000
    LvAcc.ColumnHeaders.ADD 2, , "Nama Costumer", 3000
    LvAcc.ColumnHeaders.ADD 3, , "Status Account", 3000
    LvAcc.ColumnHeaders.ADD 4, , "Agent Saat ini", 1500
    LvAcc.ColumnHeaders.ADD 5, , "Agent Terdahulu", 1500
    LvAcc.ColumnHeaders.ADD 6, , "Akses Saat ini", 1500
    LvAcc.ColumnHeaders.ADD 7, , "Waktu Akses Saat ini", 1500
End Sub

Private Sub IsiAccount()
    Dim Cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim Listitem As Listitem
    Dim M_WHERE As String
    
    M_WHERE = ""
    
    Cmdsql = "select * from mgm  "
    
    If TxtCariCustid.Text <> "" Then
        If M_WHERE = "" Then
            M_WHERE = " where custid like '%" + CStr(TxtCariCustid.Text) + "%' "
        Else
            M_WHERE = M_WHERE & " and custid like '%" + CStr(TxtCariCustid.Text) + "%' "
        End If
    End If
    
    If TxtCariNama.Text <> "" Then
        If M_WHERE = "" Then
            M_WHERE = " where name like '%" + CStr(TxtCariNama.Text) + "%' "
        Else
            M_WHERE = M_WHERE & " and name like '%" + CStr(TxtCariNama.Text) + "%' "
        End If
    End If
       
    If M_WHERE <> "" Then
        M_WHERE = M_WHERE & " and agent not in ('COMPLAIN','LUNAS','CLAIM')   "
        M_WHERE = M_WHERE & " and custid not in (select distinct custid from tblsendptp ) "
    Else
        M_WHERE = " where agent not in ('COMPLAIN','LUNAS','CLAIM') "
        M_WHERE = M_WHERE & " and custid not in (select distinct custid from tblsendptp ) "
    End If
       
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open Cmdsql & M_WHERE, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvAcc.ListItems.CLEAR
    TxtJmlhAcc.Text = M_Objrs.RecordCount
    
    
    If M_Objrs.RecordCount = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Peringatan"
        Set M_Objrs = Nothing
        Exit Sub
    End If
    
    PB1.Max = M_Objrs.RecordCount
    While Not M_Objrs.EOF
        PB1.Value = M_Objrs.Bookmark
        Set Listitem = LvAcc.ListItems.ADD(, , M_Objrs("custid"))
            Listitem.SubItems(1) = M_Objrs("name")
            Listitem.SubItems(2) = IIf(IsNull(M_Objrs("f_cek_new")), "", M_Objrs("f_cek_new"))
            Listitem.SubItems(3) = IIf(IsNull(M_Objrs("agent")), "", M_Objrs("agent"))
            Listitem.SubItems(4) = IIf(IsNull(M_Objrs("agent_asli")), "", M_Objrs("agent_asli"))
            Listitem.SubItems(5) = IIf(IsNull(M_Objrs("monitor_akses")), "", M_Objrs("monitor_akses"))
            Listitem.SubItems(6) = IIf(IsNull(M_Objrs("waktu_akses")), "", Format(M_Objrs("waktu_akses"), "yyyy-mm-dd hh:nn:ss"))
            
            If UCase(M_Objrs("agent")) = "AKSESALL" Then
                Listitem.ForeColor = vbRed
                Listitem.ListSubItems(1).ForeColor = vbRed
                Listitem.ListSubItems(2).ForeColor = vbRed
                Listitem.ListSubItems(3).ForeColor = vbRed
                Listitem.ListSubItems(4).ForeColor = vbRed
                Listitem.ListSubItems(5).ForeColor = vbRed
                Listitem.ListSubItems(6).ForeColor = vbRed
            End If
            
            If UCase(M_Objrs("agent")) = "#KOSONG#" Then
                Listitem.ForeColor = vbBlue
                Listitem.ListSubItems(1).ForeColor = vbBlue
                Listitem.ListSubItems(2).ForeColor = vbBlue
                Listitem.ListSubItems(3).ForeColor = vbBlue
                Listitem.ListSubItems(4).ForeColor = vbBlue
                Listitem.ListSubItems(5).ForeColor = vbBlue
                Listitem.ListSubItems(6).ForeColor = vbBlue
            End If
            
        M_Objrs.MoveNext
    Wend
    Set M_Objrs = Nothing
End Sub

Private Sub HeaderAgent()
    LvAgent.ColumnHeaders.ADD 1, , "ID", 1000
    LvAgent.ColumnHeaders.ADD 2, , "AGENT", 2000
    LvAgent.ColumnHeaders.ADD 3, , "CUSTID", 2000
    LvAgent.ColumnHeaders.ADD 4, , "WAKTU AWAL", 3000
    LvAgent.ColumnHeaders.ADD 5, , "WAKTU AKHIR", 3000
    LvAgent.ColumnHeaders.ADD 6, , "LOG DISTRIBUSI", 3000
    LvAgent.ColumnHeaders.ADD 7, , "WAKTU DISTRIBUSI", 3000
End Sub




Private Sub CmbAgent_Click()
    Dim Cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim Listitem As Listitem
    Dim GroupingTL_2 As String
    
    GroupingTL_2 = ""
    
    '@@19022013 Tambahan ini buat grouping TL
    If UCase(Mid(CmbAgent.Text, 1, 2)) = "TL" Then
        GroupingTL_2 = " agent in (select userid from usertbl where spvcode in ("
        GroupingTL_2 = GroupingTL_2 & " select spvcode from usertbl where userid='"
        GroupingTL_2 = GroupingTL_2 & CmbAgent.Text + "')) "
    Else
        GroupingTL_2 = " agent='"
        GroupingTL_2 = GroupingTL_2 + CmbAgent.Text + "' "
    End If
    
    If CmbAgent.Text <> "ALL" Then
        Cmdsql = "select * from tbl_distribusi_account where " & GroupingTL_2
        'Cmdsql = Cmdsql & CmbAgent.Text & "' order by waktu_awal asc "
    Else
        Cmdsql = "select * from tbl_distribusi_account "
        Cmdsql = Cmdsql & " order by agent,waktu_awal asc "
    End If
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    TxtJmlhAgent.Text = M_Objrs.RecordCount
    LvAgent.ListItems.CLEAR
    
    If M_Objrs.RecordCount = 0 Then
        
        Set M_Objrs = Nothing
        Exit Sub
    End If
    
    While Not M_Objrs.EOF
        Set Listitem = LvAgent.ListItems.ADD(, , M_Objrs("id"))
            Listitem.SubItems(1) = M_Objrs("agent")
            Listitem.SubItems(2) = M_Objrs("custid")
            Listitem.SubItems(3) = Format(M_Objrs("waktu_awal"), "yyyy-mm-dd hh:nn:ss")
            Listitem.SubItems(4) = Format(M_Objrs("waktu_akhir"), "yyyy-mm-dd hh:nn:ss")
            Listitem.SubItems(5) = M_Objrs("log_distribusi")
            Listitem.SubItems(6) = Format(M_Objrs("log_tgl_distribusi"), "yyyy-mm-dd hh:nn:ss")
       M_Objrs.MoveNext
    Wend
    
    Set M_Objrs = Nothing
End Sub

Private Sub CmbFilterAcc_Click()
    'Call CariFilter
    'Eventnya diambil berdasarkan
End Sub



Private Sub CmbStatusAcc_Click()
    Dim Cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim Listitem As Listitem
    Dim Bulan, Tahun, Tanggal As String
    Dim M_WHERE As String
    Dim GroupingTL As String
    
    If CmbFilterAcc.Text = "" Then
        MsgBox "Pilih terlebih dahulu agentnya!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If

    M_WHERE = ""
    GroupingTL = ""
    
    Cmdsql = "select * from mgm "
    
    If CmbFilterAcc.Text = "ALL" Then
        
        'Ini jika agent=ALL dan status account=ALL
        If CmbStatusAcc.Text = "ALL" Then
            'CUKUP DEH NGGA USAH PAKE SCRIPT---------
        
        '@@15022013 Ini jika filter agentnya=ALL tetapi status accountnya <> ALL
        ElseIf CmbStatusAcc.Text <> "ALL" Then
            If CmbStatusAcc.Text = "LPD 1" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 31
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                End If
            End If
            
            If CmbStatusAcc.Text = "LPD 2" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 60
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                End If
            End If
            
            If CmbStatusAcc.Text = "LPD 3" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 90
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                End If
            End If
            
            If CmbStatusAcc.Text = "LPD 3<" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 90
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)<'"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)<='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)<'"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)<='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                End If
            'ini berdasarkan status account tertentu
            ElseIf CmbStatusAcc.Text <> "LPD 1" And _
                   CmbStatusAcc.Text <> "LPD 2" And _
                   CmbStatusAcc.Text <> "LPD 3" And _
                   CmbStatusAcc.Text <> "LPD 3<" Then
                   
                If M_WHERE = "" Then
                    M_WHERE = " where f_cek_new like '%"
                    M_WHERE = M_WHERE & CmbStatusAcc.Text & "%' "
                Else
                    M_WHERE = " and f_cek_new like '%"
                    M_WHERE = M_WHERE & CmbStatusAcc.Text & "%' "
                End If
            End If
        End If
        
    ElseIf CmbFilterAcc.Text <> "ALL" Then
    
        '@@19022013 Tambahan ini buat grouping TL
        If UCase(Mid(CmbFilterAcc.Text, 1, 2)) = "TL" Then
            GroupingTL = " agent in (select userid from usertbl where spvcode in ("
            GroupingTL = GroupingTL & " select spvcode from usertbl where userid='"
            GroupingTL = GroupingTL & CmbFilterAcc.Text + "')) "
        Else
            GroupingTL = " agent='"
            GroupingTL = GroupingTL + CmbFilterAcc.Text + "' "
        End If
                
        'Ini jika agent <>ALL dan status account=ALL
        If CmbStatusAcc.Text = "ALL" Then
            If M_WHERE = "" Then
                'M_WHERE = " where agent='" & CmbFilterAcc.Text & "' "
                M_WHERE = " where " & GroupingTL
            Else
                'M_WHERE = " and agent='" & CmbFilterAcc.Text & "' "
                M_WHERE = " and " & GroupingTL
            End If
            
        'Ini jika agent=<>ALL dan status account <> ALL
        ElseIf CmbStatusAcc.Text <> "ALL" Then
            If CmbStatusAcc.Text = "LPD 1" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 31
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                    
                End If
            End If
            
            If CmbStatusAcc.Text = "LPD 2" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 60
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and  " & GroupingTL
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                End If
            End If
            
            If CmbStatusAcc.Text = "LPD 3" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 90
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                    
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                    
                End If
            End If
            
            If CmbStatusAcc.Text = "LPD 3<" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 90
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)<'"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)<='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)<'"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)<='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                End If
            'ini berdasarkan status account tertentu
            ElseIf CmbStatusAcc.Text <> "LPD 1" And _
                   CmbStatusAcc.Text <> "LPD 2" And _
                   CmbStatusAcc.Text <> "LPD 3" And _
                   CmbStatusAcc.Text <> "LPD 3<" Then
                If M_WHERE = "" Then
                    M_WHERE = " where f_cek_new like '%"
'                    M_WHERE = M_WHERE & CmbStatusAcc.Text & "%' and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CmbStatusAcc.Text & "%' and " & GroupingTL
                Else
                    M_WHERE = " and f_cek_new like '%"
'                    M_WHERE = M_WHERE & CmbStatusAcc.Text & "%' and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CmbStatusAcc.Text & "%' and " & GroupingTL
                End If
            End If
        End If
            
    End If
    
    If M_WHERE = "" Then
        M_WHERE = " where agent not in ('LUNAS','COMPLAIN','CLAIM') "
        M_WHERE = M_WHERE & " and custid not in (select distinct custid from tblsendptp ) "
        M_WHERE = M_WHERE & " order by name asc "
    Else
        M_WHERE = M_WHERE & " and agent not in ('LUNAS','COMPLAIN','CLAIM') "
        M_WHERE = M_WHERE & " and custid not in (select distinct custid from tblsendptp ) "
        M_WHERE = M_WHERE & " order by name asc "
    End If
    
    
    
    DoEvents
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open Cmdsql + M_WHERE, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvAcc.ListItems.CLEAR
    TxtJmlhAcc.Text = M_Objrs.RecordCount
    
    
    If M_Objrs.RecordCount = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Peringatan"
        Set M_Objrs = Nothing
        Exit Sub
    End If
    
    PB1.Max = M_Objrs.RecordCount
    While Not M_Objrs.EOF
        PB1.Value = M_Objrs.Bookmark
        Set Listitem = LvAcc.ListItems.ADD(, , M_Objrs("custid"))
            Listitem.SubItems(1) = M_Objrs("name")
            Listitem.SubItems(2) = IIf(IsNull(M_Objrs("f_cek_new")), "", M_Objrs("f_cek_new"))
            Listitem.SubItems(3) = IIf(IsNull(M_Objrs("agent")), "", M_Objrs("agent"))
            Listitem.SubItems(4) = IIf(IsNull(M_Objrs("agent_asli")), "", M_Objrs("agent_asli"))
            Listitem.SubItems(5) = IIf(IsNull(M_Objrs("monitor_akses")), "", M_Objrs("monitor_akses"))
            Listitem.SubItems(6) = IIf(IsNull(M_Objrs("waktu_akses")), "", Format(M_Objrs("waktu_akses"), "yyyy-mm-dd hh:nn:ss"))
            
            If UCase(M_Objrs("agent")) = "AKSESALL" Then
                Listitem.ForeColor = vbRed
                Listitem.ListSubItems(1).ForeColor = vbRed
                Listitem.ListSubItems(2).ForeColor = vbRed
                Listitem.ListSubItems(3).ForeColor = vbRed
                Listitem.ListSubItems(4).ForeColor = vbRed
                Listitem.ListSubItems(5).ForeColor = vbRed
                Listitem.ListSubItems(6).ForeColor = vbRed
            End If
            
            If UCase(M_Objrs("agent")) = "#KOSONG#" Then
                Listitem.ForeColor = vbBlue
                Listitem.ListSubItems(1).ForeColor = vbBlue
                Listitem.ListSubItems(2).ForeColor = vbBlue
                Listitem.ListSubItems(3).ForeColor = vbBlue
                Listitem.ListSubItems(4).ForeColor = vbBlue
                Listitem.ListSubItems(5).ForeColor = vbBlue
                Listitem.ListSubItems(6).ForeColor = vbBlue
            End If
            
        M_Objrs.MoveNext
    Wend
    Set M_Objrs = Nothing
    
End Sub



Private Sub CmbStatusCollBersama_Click()
    Dim Cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim Listitem As Listitem
    Dim Bulan, Tahun, Tanggal As String
    Dim M_WHERE As String
    Dim GroupingTL As String
    
    If CmbStatusCollBersama.Text = "" Then
        MsgBox "Mohon maaf, data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If CmbAgentCollBersama.Text = "" Then
        MsgBox "Pilih terlebih dahulu agent awalnya!", vbOKOnly + vbInformation, "Informasi"
        CmbAgentCollBersama.SetFocus
        Exit Sub
    End If

    M_WHERE = ""
    GroupingTL = ""
    
    Cmdsql = "select * from mgm "
    
    If CmbAgentCollBersama.Text = "ALL" Then
        
        'Ini jika agent=ALL dan status account=ALL
        If CmbStatusCollBersama.Text = "ALL" Then
            'CUKUP DEH NGGA USAH PAKE SCRIPT---------
        
        '@@15022013 Ini jika filter agentnya=ALL tetapi status accountnya <> ALL
        ElseIf CmbStatusCollBersama.Text <> "ALL" Then
            If CmbStatusCollBersama.Text = "LPD 1" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 31
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                End If
            End If
            
            If CmbStatusCollBersama.Text = "LPD 2" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 60
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                End If
            End If
            
            If CmbStatusCollBersama.Text = "LPD 3" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 90
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                End If
            End If
            
            If CmbStatusCollBersama.Text = "LPD 3<" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 90
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)<'"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)<='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)<'"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)<='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                End If
            'ini berdasarkan status account tertentu
            ElseIf CmbStatusCollBersama.Text <> "LPD 1" And _
                   CmbStatusCollBersama.Text <> "LPD 2" And _
                   CmbStatusCollBersama.Text <> "LPD 3" And _
                   CmbStatusCollBersama.Text <> "LPD 3<" Then
                   
                If M_WHERE = "" Then
                    M_WHERE = " where f_cek_new like '%"
                    M_WHERE = M_WHERE & CmbStatusCollBersama.Text & "%' "
                Else
                    M_WHERE = " and f_cek_new like '%"
                    M_WHERE = M_WHERE & CmbStatusCollBersama.Text & "%' "
                End If
            End If
        End If
        
    ElseIf CmbAgentCollBersama.Text <> "ALL" Then
    
        '@@19022013 Tambahan ini buat grouping TL
        If UCase(Mid(CmbAgentCollBersama.Text, 1, 2)) = "TL" Then
            GroupingTL = " agent_asli in (select userid from usertbl where spvcode in ("
            GroupingTL = GroupingTL & " select spvcode from usertbl where userid='"
            GroupingTL = GroupingTL & CmbAgentCollBersama.Text + "')) "
        Else
            GroupingTL = " agent_asli='"
            GroupingTL = GroupingTL + CmbAgentCollBersama.Text + "' "
        End If
                
        'Ini jika agent <>ALL dan status account=ALL
        If CmbStatusCollBersama.Text = "ALL" Then
            If M_WHERE = "" Then
                'M_WHERE = " where agent='" & CmbFilterAcc.Text & "' "
                M_WHERE = " where " & GroupingTL
            Else
                'M_WHERE = " and agent='" & CmbFilterAcc.Text & "' "
                M_WHERE = " and " & GroupingTL
            End If
            
        'Ini jika agent=<>ALL dan status account <> ALL
        ElseIf CmbStatusCollBersama.Text <> "ALL" Then
            If CmbStatusCollBersama.Text = "LPD 1" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 31
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                    
                End If
            End If
            
            If CmbStatusCollBersama.Text = "LPD 2" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 60
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and  " & GroupingTL
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                End If
            End If
            
            If CmbStatusCollBersama.Text = "LPD 3" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 90
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                    
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                    
                End If
            End If
            
            If CmbStatusCollBersama.Text = "LPD 3<" Then
                Tanggal = CDate(Format(MDIForm1.TDBDate1.Value, "m/dd/yyyy")) - 90
                Bulan = Format(Tanggal, "mm")
                Tahun = Format(Tanggal, "yyyy")
                
                If M_WHERE = "" Then
                    M_WHERE = " where custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)<'"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)<='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)<'"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)<='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                End If
            'ini berdasarkan status account tertentu
            ElseIf CmbStatusCollBersama.Text <> "LPD 1" And _
                   CmbStatusCollBersama.Text <> "LPD 2" And _
                   CmbStatusCollBersama.Text <> "LPD 3" And _
                   CmbStatusCollBersama.Text <> "LPD 3<" Then
                If M_WHERE = "" Then
                    M_WHERE = " where f_cek_new like '%"
'                    M_WHERE = M_WHERE & CmbStatusAcc.Text & "%' and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CmbStatusCollBersama.Text & "%' and " & GroupingTL
                Else
                    M_WHERE = " and f_cek_new like '%"
'                    M_WHERE = M_WHERE & CmbStatusAcc.Text & "%' and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CmbStatusCollBersama.Text & "%' and " & GroupingTL
                End If
            End If
        End If
            
    End If
    
    If M_WHERE = "" Then
        M_WHERE = " where agent not in ('LUNAS','COMPLAIN','CLAIM') and agent='AKSESALL' order by name asc "
    Else
        M_WHERE = M_WHERE & " and agent not in ('LUNAS','COMPLAIN','CLAIM') and agent='AKSESALL' order by name asc "
    End If
    
    
    
    DoEvents
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open Cmdsql + M_WHERE, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvAcc.ListItems.CLEAR
    TxtJmlhAcc.Text = M_Objrs.RecordCount
    
    
    If M_Objrs.RecordCount = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Peringatan"
        Set M_Objrs = Nothing
        Exit Sub
    End If
    
    PB1.Max = M_Objrs.RecordCount
    While Not M_Objrs.EOF
        PB1.Value = M_Objrs.Bookmark
        Set Listitem = LvAcc.ListItems.ADD(, , M_Objrs("custid"))
            Listitem.SubItems(1) = M_Objrs("name")
            Listitem.SubItems(2) = IIf(IsNull(M_Objrs("f_cek_new")), "", M_Objrs("f_cek_new"))
            Listitem.SubItems(3) = IIf(IsNull(M_Objrs("agent")), "", M_Objrs("agent"))
            Listitem.SubItems(4) = IIf(IsNull(M_Objrs("agent_asli")), "", M_Objrs("agent_asli"))
            Listitem.SubItems(5) = IIf(IsNull(M_Objrs("monitor_akses")), "", M_Objrs("monitor_akses"))
            Listitem.SubItems(6) = IIf(IsNull(M_Objrs("waktu_akses")), "", Format(M_Objrs("waktu_akses"), "yyyy-mm-dd hh:nn:ss"))
            
            If UCase(M_Objrs("agent")) = "AKSESALL" Then
                Listitem.ForeColor = vbRed
                Listitem.ListSubItems(1).ForeColor = vbRed
                Listitem.ListSubItems(2).ForeColor = vbRed
                Listitem.ListSubItems(3).ForeColor = vbRed
                Listitem.ListSubItems(4).ForeColor = vbRed
                Listitem.ListSubItems(5).ForeColor = vbRed
                Listitem.ListSubItems(6).ForeColor = vbRed
            End If
            
            If UCase(M_Objrs("agent")) = "#KOSONG#" Then
                Listitem.ForeColor = vbBlue
                Listitem.ListSubItems(1).ForeColor = vbBlue
                Listitem.ListSubItems(2).ForeColor = vbBlue
                Listitem.ListSubItems(3).ForeColor = vbBlue
                Listitem.ListSubItems(4).ForeColor = vbBlue
                Listitem.ListSubItems(5).ForeColor = vbBlue
                Listitem.ListSubItems(6).ForeColor = vbBlue
            End If
            
        M_Objrs.MoveNext
    Wend
    Set M_Objrs = Nothing
End Sub

Private Sub CmdBukaAccount_Click()
    Dim Cmdsql As String
    Dim W, K, S As Integer
    Dim a As String
    
    If LvAcc.ListItems.Count = 0 Then
        MsgBox "Data account tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    a = MsgBox("Anda yakin membuka lock account yang terceklist?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbNo Then
        MsgBox "proses dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    S = 0
    For K = 1 To LvAcc.ListItems.Count
        If LvAcc.ListItems(K).Checked = True Then
            S = S + 1
        End If
    Next K
    
    If S = 0 Then
        MsgBox "Anda belum menceklist account yang akan dibuka!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    PB1.Max = LvAcc.ListItems.Count
    For W = 1 To LvAcc.ListItems.Count
        PB1.Value = W
        If LvAcc.ListItems(W).Checked = True Then
            'buka locknya
            Cmdsql = "update mgm set monitor_akses=null,waktu_akses=null where custid='"
            Cmdsql = Cmdsql & CStr(LvAcc.ListItems(W).Text) & "'"
            M_OBJCONN.Execute Cmdsql
        End If
    Next W
    
    Call IsiAccount
    
    MsgBox "Proses berhasil!", vbOKOnly + vbInformation, "Informasi"
    
End Sub

Private Sub CmdCari_Click()
    Call IsiAccount
End Sub

Private Sub CmdCekAllAcc_Click()
    Dim W As Integer
    
    If LvAcc.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LvAcc.ListItems.Count
        LvAcc.ListItems(W).Checked = True
    Next W
End Sub

Private Sub CmdCekAllAgent_Click()
    Dim W As Integer
    
    If LvAgent.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LvAgent.ListItems.Count
        LvAgent.ListItems(W).Checked = True
    Next W
End Sub

Private Sub CmdClear_Click()
    TxtCariCustid.Text = ""
    TxtCariNama.Text = ""
End Sub

Private Sub CmdEdit_Click()
    If LvAcc.ListItems.Count = 0 Then
        MsgBox "Data Account tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If LvAgent.ListItems.Count = 0 Then
        MsgBox "Data Agent tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    FrmEditDistribusiAccount.TxtID.Text = LvAgent.SelectedItem.Text
    FrmEditDistribusiAccount.TxtAgent.Text = LvAgent.SelectedItem.SubItems(1)
    FrmEditDistribusiAccount.Show vbModal
End Sub

Private Sub CmdFilterExcel_Click()
    FrmFilterExcelDistribusiAcc.Show vbModal
End Sub

Private Sub CmdFormClaimAccount_Click()
    FrmListClaim.Show vbModal
End Sub

Private Sub CmdHapusAgent_Click()
    Dim a As String
    Dim Cmdsql As String
    Dim W, i, K As Integer
    Dim M_Objrs As ADODB.Recordset
    
    If LvAgent.ListItems.Count = 0 Then
        MsgBox "Data agent tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    a = MsgBox("Yakin data agent akan dihapus?", vbYesNo + vbInformation, "Informasi")
    If a = vbNo Then
        MsgBox "Proses dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    i = 0
    For W = 1 To LvAgent.ListItems.Count
       If LvAgent.ListItems(W).Checked = True Then
            i = i + 1
       End If
    Next W
    
    
    If i = 0 Then
        MsgBox "Anda belum memilih data agent yang akan dihapus!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
        
    DoEvents
        
    PB1.Max = LvAgent.ListItems.Count
        
    For K = 1 To LvAgent.ListItems.Count
        PB1.Value = K
        If LvAgent.ListItems(K).Checked = True Then
            Cmdsql = "delete from tbl_distribusi_account where id='"
            Cmdsql = Cmdsql + CStr(LvAgent.ListItems(K).Text) + "'"
            M_OBJCONN.Execute Cmdsql
            
            'Update status agentnya nih
            Cmdsql = "update usertbl set f_akses_all_acc=null,f_pesanresetauto='1' "
            Cmdsql = Cmdsql + " where userid='"
            Cmdsql = Cmdsql + CStr(LvAgent.ListItems(K).SubItems(1)) + "'"
            M_OBJCONN.Execute Cmdsql
            
            'Cek apakah custid ini sudah habis agentnya?
            Cmdsql = "select * from tbl_distribusi_account where custid='"
            Cmdsql = Cmdsql & CStr(LvAgent.ListItems(K).SubItems(2)) & "'"
            Set M_Objrs = New ADODB.Recordset
            M_Objrs.CursorLocation = adUseClient
            M_Objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If M_Objrs.RecordCount = 0 Then
                'Update ke agent yang lama
                Cmdsql = "update mgm set agent=agent_asli,agent_asli=null,"
                Cmdsql = Cmdsql + " user_claim=null,waktu_claim=null,alasan_claim=null "
                Cmdsql = Cmdsql + " where custid='"
                Cmdsql = Cmdsql & CStr(LvAgent.ListItems(K).SubItems(2)) + "' and agent_asli is not null "
                M_OBJCONN.Execute Cmdsql
            End If
            Set M_Objrs = Nothing
            
        End If
    Next K
    
    'Cek apakah custid ini sudah habis agentnya?
'    Cmdsql = "select * from tbl_distribusi_account where custid='"
'    Cmdsql = Cmdsql + CStr(lblCustId.Caption) + "'"
'    Set M_Objrs = New ADODB.Recordset
'    M_Objrs.CursorLocation = adUseClient
'    M_Objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
'    If M_Objrs.RecordCount = 0 Then
'        If LvAcc.SelectedItem.SubItems(3) <> "" Then
'            A = MsgBox("Account ini sudah tidak ada yang memiliki. Anda ingin mengembalikannya ke agent terdahulu?", vbYesNo + vbQuestion, "Konfirmasi")
'            If A = vbYes Then
'                'Update ke agent yang lama
'                Cmdsql = "update mgm set agent=agent_asli,agent_asli=null where custid='"
'                Cmdsql = Cmdsql & CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'            Else
'                'Update ke agent yang kosong
'                Cmdsql = "update mgm set agent='#KOSONG#' where custid='"
'                Cmdsql = Cmdsql & CStr(lblCustId.Caption) + "'"
'                M_OBJCONN.Execute Cmdsql
'            End If
'        Else
'            'Update ke agent yang kosong
'            Cmdsql = "update mgm set agent='#KOSONG#' where custid='"
'            Cmdsql = Cmdsql & CStr(lblCustId.Caption) + "'"
'            M_OBJCONN.Execute Cmdsql
'        End If
'    End If
    
    Call CariAgent
    
    MsgBox "Data agent berhasil dihapus!", vbOKOnly + vbInformation, "Informasi"
    
    'Call IsiAccount
End Sub

Private Sub CmdKembalikanAgent_Click()
    Dim Cmdsql As String
    Dim W, K, S As Integer
    Dim a As String
    
    If LvAcc.ListItems.Count = 0 Then
        MsgBox "Data account tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    a = MsgBox("Anda yakin akan mengembalikan account yang diceklist ke agent awal?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbNo Then
        MsgBox "proses dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    S = 0
    For K = 1 To LvAcc.ListItems.Count
        If LvAcc.ListItems(K).Checked = True Then
            S = S + 1
        End If
    Next K
    
    If S = 0 Then
        MsgBox "Anda belum menceklist account yang akan dikembalikan agentnya!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    PB1.Max = LvAcc.ListItems.Count
    For W = 1 To LvAcc.ListItems.Count
        PB1.Value = W
        If LvAcc.ListItems(W).Checked = True Then
            Cmdsql = "update mgm set agent=agent_asli,agent_asli=null where custid='"
            Cmdsql = Cmdsql + CStr(LvAcc.ListItems(W).Text) + "' and agent_asli is not null "
            M_OBJCONN.Execute Cmdsql
            
            'Hapus data di tabel distribusinya
            Cmdsql = "delete from tbl_distribusi_account where custid='"
            Cmdsql = Cmdsql + CStr(LvAcc.ListItems(W).Text) + "'"
            M_OBJCONN.Execute Cmdsql
        End If
    Next W
    
    Call IsiAccount
    
    MsgBox "Account berhasil dikembalikan ke agent awal!", vbOKOnly + vbInformation, "Informasi"
End Sub

Private Sub CmdLihatListAgent_Click()
    FrmListAgent.Show vbModal
End Sub

Private Sub CmdProses_Click()
    Dim Cmdsql, AmbilCustid As String
    Dim M_Objrs As ADODB.Recordset
    Dim W, K, S As Integer
    Dim Tanggal1 As String
    Dim Tanggal2 As String
    Dim Pesan As String
    Dim a As String
    Dim M_ObjrsWaktuServer As ADODB.Recordset
    Dim WaktuServer As String
    
    If LvAcc.ListItems.Count = 0 Then
        MsgBox "Data customer tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If TxtAgent.Text = "" Then
        MsgBox "Agent tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    S = 0
    For W = 1 To LvAcc.ListItems.Count
        If LvAcc.ListItems(W).Checked = True Then
            S = S + 1
        End If
    Next W
    
    If S = 0 Then
        MsgBox "Anda belum memilih data customer!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    a = MsgBox("Anda yakin menandai account dapat di collect bersama?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbNo Then
        MsgBox "Proses dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
        
    MDIForm1.Timer1 = False
    MDIForm1.TimerCTI = False
    MDIForm1.TimerBlink = False
    'Cek waktu server
    Cmdsql = "select now()"
    Set M_ObjrsWaktuServer = New ADODB.Recordset
    M_ObjrsWaktuServer.CursorLocation = adUseClient
    DoEvents
    M_ObjrsWaktuServer.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
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
    
    On Error GoTo salah
    
    'Ambil Data agent
    Cmdsql = "select * from usertbl where userid in ("
    Cmdsql = Cmdsql + TxtAgent.Text + ")"
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount > 0 Then
        DoEvents
        LvAcc.Enabled = False
        CmdCari.Enabled = False
        cmdClear.Enabled = False
        CmdLihatListAgent.Enabled = False
        TxtTglAwal.Enabled = False
        TxtWaktuAwal.Enabled = False
        TxtTglAkhir.Enabled = False
        TxtWaktuAkhir.Enabled = False
        
        CmbAgentCollBersama.Enabled = False
        CmbStatusCollBersama.Enabled = False
        
        PB1.Max = M_Objrs.RecordCount
        While Not M_Objrs.EOF
            PB1.Value = M_Objrs.Bookmark
            DoEvents
            'Update status Agentnya
            Cmdsql = "update usertbl set f_akses_all_acc='1',f_pesanresetauto='1' "
            Cmdsql = Cmdsql + " where userid='"
            Cmdsql = Cmdsql + M_Objrs("userid") + "'"
            M_OBJCONN.Execute Cmdsql
            
            'Kirim pesan ke agent
            Pesan = "Pesan dibuat otomatis oleh system!" & vbCrLf
            Pesan = Pesan & "----------------------------------------------" & vbCrLf
            Pesan = Pesan & "SPV menambahkan account baru untuk anda. " & vbCrLf
            Pesan = Pesan & "Account ini dapat di collect secara bersama-sama oleh anda, " & vbCrLf
            Pesan = Pesan & "mulai dari :" & Format(Tanggal1, "yyyy-mm-dd hh:nn:ss") & " s.d. " & vbCrLf
            Pesan = Pesan & Format(Tanggal2, "yyyy-mm-dd hh:nn:ss") & vbCrLf
            Pesan = Pesan & "Cek account baru anda dengan mengklik ulang tombol search data!"
            
            Cmdsql = "insert into msgtbl "
            Cmdsql = Cmdsql + "( recipient, datetime, sender, sentfrom, msg) values ('"
            Cmdsql = Cmdsql + M_Objrs("userid") + "','"
            Cmdsql = Cmdsql + Format(Now(), "yyyymmdd") + "','"
            Cmdsql = Cmdsql + MDIForm1.Text1.Text + "','"
            Cmdsql = Cmdsql + CStr(MDIForm1.Winsock1.LocalIP) + "','"
            Cmdsql = Cmdsql + Pesan + "')"
            
            M_OBJCONN.Execute Cmdsql
            
            
            For K = 1 To LvAcc.ListItems.Count
                DoEvents
                If LvAcc.ListItems(K).Checked = True Then
                    'Hapus dulu jika ada data sebelumnya
                    Cmdsql = "delete from tbl_distribusi_account where custid='"
                    Cmdsql = Cmdsql & CStr(LvAcc.ListItems(K).Text) & "' and agent='"
                    Cmdsql = Cmdsql & M_Objrs("userid") & "'"
                    M_OBJCONN.Execute Cmdsql

                    'Inputkan data ke database
                    Cmdsql = "insert into tbl_distribusi_account (custid,agent,"
                    Cmdsql = Cmdsql & "waktu_awal,waktu_akhir, log_distribusi,"
                    Cmdsql = Cmdsql & "log_tgl_distribusi) values ('"
                    Cmdsql = Cmdsql & CStr(LvAcc.ListItems(K).Text) + "','"
                    Cmdsql = Cmdsql & M_Objrs("userid") & "','"
                    Cmdsql = Cmdsql & Format(TxtTglAwal.Value, "yyyy-mm-dd") & " " & Format(TxtWaktuAwal.Value) & "','"
                    Cmdsql = Cmdsql & Format(TxtTglAkhir.Value, "yyyy-mm-dd") & " " & Format(TxtWaktuAkhir.Value) & "','"
                    Cmdsql = Cmdsql & MDIForm1.Text1.Text & "',"
                    Cmdsql = Cmdsql & " now())"
                    M_OBJCONN.Execute Cmdsql
    
                End If
                Debug.Print "loop ke" & CStr(K)
            Next K
            Debug.Print "record ke" & M_Objrs.Bookmark
            M_Objrs.MoveNext
        Wend
        
        'Catet agent lama
        PB1.Max = LvAcc.ListItems.Count
        For K = 1 To LvAcc.ListItems.Count
        DoEvents
            PB1.Value = K
            If LvAcc.ListItems(K).Checked = True Then
                '@@12022013 Jika status account sebelumnya #KOSONG# atau AKSESALL, ga usah diupdate
                If UCase(LvAcc.ListItems(K).SubItems(3)) = "#KOSONG#" Then
                    Cmdsql = "update mgm set  agent='AKSESALL' where custid='"
                    Cmdsql = Cmdsql & CStr(LvAcc.ListItems(K).Text) & "'"
                    M_OBJCONN.Execute Cmdsql
                ElseIf UCase(LvAcc.ListItems(K).SubItems(3)) <> "AKSESALL" Then
                    Cmdsql = "update mgm set agent_asli=agent, agent='AKSESALL' where custid='"
                    Cmdsql = Cmdsql & CStr(LvAcc.ListItems(K).Text) & "'"
                    M_OBJCONN.Execute Cmdsql
                    
                    '@@18022013 Ini buat, inputin otomatis buat pemilik agent masuk juga akses all
                    'Hapus dulu jika ada data sebelumnya
                    Cmdsql = "delete from tbl_distribusi_account where custid='"
                    Cmdsql = Cmdsql & CStr(LvAcc.ListItems(K).Text) & "' and agent='"
                    Cmdsql = Cmdsql & CStr(LvAcc.ListItems(K).SubItems(3)) & "'"
                    M_OBJCONN.Execute Cmdsql
                    
                    'Inputkan data ke database
                    Cmdsql = "insert into tbl_distribusi_account (custid,agent,"
                    Cmdsql = Cmdsql & "waktu_awal,waktu_akhir, log_distribusi,"
                    Cmdsql = Cmdsql & "log_tgl_distribusi) values ('"
                    Cmdsql = Cmdsql & CStr(LvAcc.ListItems(K).Text) + "','"
                    Cmdsql = Cmdsql & CStr(LvAcc.ListItems(K).SubItems(3)) & "','"
                    Cmdsql = Cmdsql & Format(TxtTglAwal.Value, "yyyy-mm-dd") & " " & Format(TxtWaktuAwal.Value) & "','"
                    Cmdsql = Cmdsql & Format(TxtTglAkhir.Value, "yyyy-mm-dd") & " " & Format(TxtWaktuAkhir.Value) & "','"
                    Cmdsql = Cmdsql & MDIForm1.Text1.Text & "',"
                    Cmdsql = Cmdsql & " now())"
                    M_OBJCONN.Execute Cmdsql
                                  
                    'Update statusnya
                    Cmdsql = "update usertbl set f_akses_all_acc='1', f_pesanresetauto='1' where "
                    Cmdsql = Cmdsql + " userid='"
                    Cmdsql = Cmdsql + CStr(LvAcc.ListItems(K).SubItems(3)) + "'"
                    M_OBJCONN.Execute Cmdsql
                    
                End If
            End If
            Debug.Print "SET agent ke " & CStr(K)
        Next K
    End If
    
    LvAcc.Enabled = True
    CmdCari.Enabled = True
    cmdClear.Enabled = True
    CmdLihatListAgent.Enabled = True
    TxtTglAwal.Enabled = True
    TxtWaktuAwal.Enabled = True
    TxtTglAkhir.Enabled = True
    TxtWaktuAkhir.Enabled = True
    
    CmbAgentCollBersama.Enabled = True
    CmbStatusCollBersama.Enabled = True
    
    
    'Call IsiAccount
    CmbStatusAcc_Click
    IsiAgentCollectBersama
    IsiStatusCollectBersama
    
    ' Hidupkan TIMER
    MDIForm1.Timer1 = True
    MDIForm1.TimerCTI = True
    MDIForm1.TimerBlink = True
    
    MsgBox "Proses berhasil!", vbOKOnly + vbInformation, "Informasi"
    Exit Sub
salah:
    MsgBox "Mohon maaf ada kesalahan: " & err.Description, "Error"
End Sub

Private Sub CmdUnCekAll_Click()
    Dim W As Integer
    
    If LvAcc.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LvAcc.ListItems.Count
        LvAcc.ListItems(W).Checked = False
    Next W
End Sub

Private Sub CmdUncekallAgent_Click()
        Dim W As Integer
    
    If LvAgent.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LvAgent.ListItems.Count
        LvAgent.ListItems(W).Checked = False
    Next W
End Sub

Private Sub Form_Load()
    Dim Cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    Call HeaderAgent
    Call HeaderAccount
    Call IsiComboFilter
    Call IsiComboAgent
    Call IsiComboStatusAcc
    
    Cmdsql = "select now()"
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_Objrs.RecordCount > 0 Then
        LblWaktuServer.Caption = Format(M_Objrs(0), "yyyy-mm-dd hh:nn")
    End If
    Set M_Objrs = Nothing
    
    Call IsiAgentCollectBersama
    Call IsiStatusCollectBersama
End Sub

Private Sub LvAcc_Click()
    Call CariAgent
End Sub

Public Sub CariAgent()
    Dim Cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim Listitem As Listitem
    
    If LvAcc.ListItems.Count = 0 Then
        MsgBox "Maaf, data tidak tersedia!", vbOKOnly + vbInformation, "Peringatan"
        Exit Sub
    End If
    
    
    lblcustid.Caption = LvAcc.SelectedItem.Text
    LblNama.Caption = LvAcc.SelectedItem.SubItems(1)
    LblStatusAcc.Caption = IIf(IsNull(LvAcc.SelectedItem.SubItems(2)), "-", LvAcc.SelectedItem.SubItems(2))
    
    Cmdsql = "select * from tbl_distribusi_account where custid='"
    Cmdsql = Cmdsql & CStr(lblcustid.Caption) & "' order by agent asc "
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    TxtJmlhAgent.Text = M_Objrs.RecordCount
    LvAgent.ListItems.CLEAR
    
    If M_Objrs.RecordCount = 0 Then
        Set M_Objrs = Nothing
        Exit Sub
    End If
    
    While Not M_Objrs.EOF
        Set Listitem = LvAgent.ListItems.ADD(, , M_Objrs("id"))
            Listitem.SubItems(1) = M_Objrs("agent")
            Listitem.SubItems(2) = M_Objrs("custid")
            Listitem.SubItems(2) = Format(M_Objrs("waktu_awal"), "yyyy-mm-dd hh:nn:ss")
            Listitem.SubItems(3) = Format(M_Objrs("waktu_akhir"), "yyyy-mm-dd hh:nn:ss")
            Listitem.SubItems(4) = M_Objrs("log_distribusi")
            Listitem.SubItems(5) = Format(M_Objrs("log_tgl_distribusi"), "yyyy-mm-dd hh:nn:ss")
       M_Objrs.MoveNext
    Wend
    
    Set M_Objrs = Nothing
End Sub

Private Sub IsiComboFilter()
    Dim Cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    CmbFilterAcc.CLEAR
    
    CmbFilterAcc.AddItem "ALL"
    
    Cmdsql = "select * from usertbl where userid not in ('LUNAS','COMPLAIN','CLAIM') order by userid asc "
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If M_Objrs.RecordCount > 0 Then
        While Not M_Objrs.EOF
            CmbFilterAcc.AddItem M_Objrs("userid")
            M_Objrs.MoveNext
        Wend
    End If
    
    Set M_Objrs = Nothing
End Sub

Private Sub CariFilter()
    Dim Cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim Listitem As Listitem
    Dim M_WHERE As String
    
    M_WHERE = ""
    
    Cmdsql = "select * from mgm  "
    
    If CmbFilterAcc.Text <> "ALL" Then
        If M_WHERE = "" Then
            M_WHERE = " where agent='" + CStr(CmbFilterAcc.Text) + "' "
            M_WHERE = M_WHERE & " and agent not in ('COMPLAIN','LUNAS') "
        Else
            M_WHERE = M_WHERE & " and agent='" + CStr(CmbFilterAcc.Text) + "' "
            M_WHERE = M_WHERE & " and agent not in ('COMPLAIN','LUNAS') "
        End If
    End If
    
    
    If M_WHERE <> "" Then
        M_WHERE = M_WHERE & " and agent not in ('COMPLAIN','LUNAS') "
    Else
        M_WHERE = " where agent not in ('COMPLAIN','LUNAS') "
    End If
    
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open Cmdsql & M_WHERE, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvAcc.ListItems.CLEAR
    TxtJmlhAcc.Text = M_Objrs.RecordCount
    
    
    If M_Objrs.RecordCount = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Peringatan"
        Set M_Objrs = Nothing
        Exit Sub
    End If
    
    PB1.Max = M_Objrs.RecordCount
    While Not M_Objrs.EOF
        PB1.Value = M_Objrs.Bookmark
        Set Listitem = LvAcc.ListItems.ADD(, , M_Objrs("custid"))
            Listitem.SubItems(1) = M_Objrs("name")
            Listitem.SubItems(2) = IIf(IsNull(M_Objrs("f_cek_new")), "", M_Objrs("f_cek_new"))
            Listitem.SubItems(3) = IIf(IsNull(M_Objrs("agent")), "", M_Objrs("agent"))
            Listitem.SubItems(4) = IIf(IsNull(M_Objrs("agent_asli")), "", M_Objrs("agent_asli"))
            Listitem.SubItems(5) = IIf(IsNull(M_Objrs("monitor_akses")), "", M_Objrs("monitor_akses"))
            Listitem.SubItems(6) = IIf(IsNull(M_Objrs("waktu_akses")), "", Format(M_Objrs("waktu_akses"), "yyyy-mm-dd hh:nn:ss"))
            
            If UCase(M_Objrs("agent")) = "AKSESALL" Then
                Listitem.ForeColor = vbRed
                Listitem.ListSubItems(1).ForeColor = vbRed
                Listitem.ListSubItems(2).ForeColor = vbRed
                Listitem.ListSubItems(3).ForeColor = vbRed
                Listitem.ListSubItems(4).ForeColor = vbRed
                Listitem.ListSubItems(5).ForeColor = vbRed
                Listitem.ListSubItems(6).ForeColor = vbRed
            End If
            
            If UCase(M_Objrs("agent")) = "#KOSONG#" Then
                Listitem.ForeColor = vbBlue
                Listitem.ListSubItems(1).ForeColor = vbBlue
                Listitem.ListSubItems(2).ForeColor = vbBlue
                Listitem.ListSubItems(3).ForeColor = vbBlue
                Listitem.ListSubItems(4).ForeColor = vbBlue
                Listitem.ListSubItems(5).ForeColor = vbBlue
                Listitem.ListSubItems(6).ForeColor = vbBlue
            End If
            
        M_Objrs.MoveNext
    Wend
    Set M_Objrs = Nothing
End Sub

Private Sub LvAcc_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LvAcc.SortKey = ColumnHeader.Index - 1
    LvAcc.Sorted = True
End Sub

Private Sub LvAgent_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LvAgent.SortKey = ColumnHeader.Index - 1
    LvAgent.Sorted = True
End Sub

Private Sub LvAgent_DblClick()
    CmdEdit_Click
End Sub

Private Sub IsiComboAgent()
    Dim Cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    CmbAgent.CLEAR
    CmbAgent.AddItem "ALL"
    
    Cmdsql = "select * from usertbl where usertype in ('1','6') and userid "
    Cmdsql = Cmdsql & " not in ('LUNAS','COMPLAIN','COMPLAIN','CLAIM') and userid is not null order by userid asc "
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount > 0 Then
        While Not M_Objrs.EOF
            CmbAgent.AddItem M_Objrs("userid")
            M_Objrs.MoveNext
        Wend
    End If
    
    Set M_Objrs = Nothing
End Sub

'@@14022013 Tambahan filter status account
Private Sub IsiComboStatusAcc()
    Dim Cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    CmbStatusAcc.CLEAR
    CmbStatusAcc.AddItem "ALL"
    CmbStatusAcc.AddItem "LPD 1"
    CmbStatusAcc.AddItem "LPD 2"
    CmbStatusAcc.AddItem "LPD 3"
    CmbStatusAcc.AddItem "LPD 3<"
    
    Cmdsql = "select * from contacteddesc where status='1' and jenis is not null "
    Cmdsql = Cmdsql & " and  jenis<>'CO-' order by jenis asc "
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_Objrs.RecordCount > 0 Then
        While Not M_Objrs.EOF
            CmbStatusAcc.AddItem IIf(IsNull(M_Objrs("jenis")), "", M_Objrs("jenis"))
            M_Objrs.MoveNext
        Wend
    End If
    CmbStatusAcc.AddItem "PTP-"
    Set M_Objrs = Nothing
End Sub

'@@21022013 Tambahan buat program filter account yang bisa di collect bersama
Private Sub IsiAgentCollectBersama()
    Dim Cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    Dim M_Objrs_TL As ADODB.Recordset
    
    CmbAgentCollBersama.CLEAR
    
    Cmdsql = "select distinct agent_asli from mgm "
    Cmdsql = Cmdsql + " where agent_asli is not null and agent='AKSESALL' "
    Cmdsql = Cmdsql + " and agent in (select userid from usertbl where usertype='1') "
    Cmdsql = Cmdsql + " order by agent_asli asc "
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, admcdtext
    If M_Objrs.RecordCount > 0 Then
        CmbAgentCollBersama.AddItem "ALL"
        While Not M_Objrs.EOF
            CmbAgentCollBersama.AddItem M_Objrs("agent_asli")
            M_Objrs.MoveNext
        Wend
    
        'Load TLnya juga buat grouping
        Cmdsql = "select userid from usertbl where usertype='6' order by userid asc "
        Set M_Objrs_TL = New ADODB.Recordset
        M_Objrs_TL.CursorLocation = adUseClient
        M_Objrs_TL.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_Objrs_TL.RecordCount > 0 Then
            While Not M_Objrs_TL.EOF
                CmbAgentCollBersama.AddItem M_Objrs_TL("userid")
                M_Objrs_TL.MoveNext
            Wend
        End If
        Set M_Objrs_TL = Nothing
    End If
    Set M_Objrs = Nothing
    
End Sub


Private Sub IsiStatusCollectBersama()
    Dim Cmdsql As String
    Dim M_Objrs As ADODB.Recordset
    
    CmbStatusCollBersama.CLEAR
    
    Cmdsql = "select distinct f_cek_new from mgm where agent='AKSESALL' "
    Cmdsql = Cmdsql + " and f_cek_new is not null order by f_cek_new asc "
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_Objrs.RecordCount > 0 Then
        CmbStatusCollBersama.AddItem "ALL"
        CmbStatusCollBersama.AddItem "LPD 1"
        CmbStatusCollBersama.AddItem "LPD 2"
        CmbStatusCollBersama.AddItem "LPD 3"
        CmbStatusCollBersama.AddItem "LPD 3<"
        While Not M_Objrs.EOF
            CmbStatusCollBersama.AddItem Trim(UCase(Mid(M_Objrs("f_cek_new"), 1, 3)))
            M_Objrs.MoveNext
        Wend
    End If
    Set M_Objrs = Nothing
End Sub




