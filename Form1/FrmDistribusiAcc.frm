VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
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
   Begin VB.CommandButton cmd_insert_antrian 
      BackColor       =   &H0000FF00&
      Caption         =   "&Proses..."
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   6210
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filter Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   5130
      TabIndex        =   56
      Top             =   330
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CheckBox Check_decease 
         Caption         =   "Include Account Decease [ 835 ]"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   93
         Top             =   7200
         Width           =   3375
      End
      Begin VB.ListBox list_batch 
         Height          =   1035
         ItemData        =   "FrmDistribusiAcc.frx":0000
         Left            =   1080
         List            =   "FrmDistribusiAcc.frx":0007
         MultiSelect     =   2  'Extended
         TabIndex        =   90
         Top             =   360
         Width           =   2175
      End
      Begin VB.Frame Frame3 
         Caption         =   "AMOUNT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   3120
         TabIndex        =   80
         Top             =   3690
         Width           =   2895
         Begin VB.Frame Frame4 
            Enabled         =   0   'False
            Height          =   2295
            Left            =   240
            TabIndex        =   83
            Top             =   600
            Width           =   2415
            Begin Threed.SSOption SSOption1 
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   84
               Top             =   240
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   450
               _Version        =   196610
               Caption         =   "5.000.000 - 10.000.000"
            End
            Begin Threed.SSOption SSOption1 
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   85
               Top             =   600
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   450
               _Version        =   196610
               Caption         =   "10.000.000 - 30.000.000"
            End
            Begin Threed.SSOption SSOption1 
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   86
               Top             =   960
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   450
               _Version        =   196610
               Caption         =   "30.000.000 - 60.000.000"
            End
            Begin Threed.SSOption SSOption1 
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   87
               Top             =   1320
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   450
               _Version        =   196610
               Caption         =   "60.000.000 - 90.000.000"
            End
            Begin Threed.SSOption SSOption1 
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   88
               Top             =   1680
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   450
               _Version        =   196610
               Caption         =   "90.000.000 - 120.000.000"
            End
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Current Balance"
            Height          =   255
            Left            =   120
            TabIndex        =   82
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   79
         Top             =   7080
         Width           =   1335
      End
      Begin VB.ComboBox cb_batch 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   74
         Top             =   360
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Filter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   73
         Top             =   7080
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "WO DATE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   120
         TabIndex        =   57
         Top             =   3690
         Width           =   2895
         Begin VB.CheckBox Check2 
            Caption         =   "WO DATE"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   360
            Width           =   1095
         End
         Begin VB.ListBox List1 
            Enabled         =   0   'False
            Height          =   2400
            ItemData        =   "FrmDistribusiAcc.frx":0017
            Left            =   360
            List            =   "FrmDistribusiAcc.frx":004B
            MultiSelect     =   2  'Extended
            TabIndex        =   58
            Top             =   750
            Width           =   1455
         End
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   1575
         Left            =   120
         TabIndex        =   60
         Top             =   2040
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2778
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "STATUS"
         Begin VB.CheckBox Check1 
            Caption         =   "UN-Uncontacted"
            Height          =   255
            Index           =   7
            Left            =   150
            TabIndex        =   72
            Top             =   2010
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "POP-Progress Of Payment"
            Height          =   255
            Index           =   6
            Left            =   3030
            TabIndex        =   71
            Top             =   510
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "BP-Broken Promise"
            Height          =   255
            Index           =   5
            Left            =   3030
            TabIndex        =   70
            Top             =   240
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "PTP-Promise To Pay"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   69
            Top             =   2070
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "PR-PROSPECT"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   68
            Top             =   720
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "VL-VALID"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   67
            Top             =   240
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "RP-Refuse Payment"
            Height          =   255
            Index           =   1
            Left            =   135
            TabIndex        =   66
            Top             =   480
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "SP-Settled Payment"
            Height          =   255
            Index           =   8
            Left            =   3030
            TabIndex        =   65
            Top             =   750
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "BL-Data Blank"
            Height          =   255
            Index           =   9
            Left            =   150
            TabIndex        =   64
            Top             =   2010
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "OS-On Process"
            Height          =   255
            Index           =   10
            Left            =   3030
            TabIndex        =   63
            Top             =   990
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "ON-On Nego"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   62
            Top             =   990
            Width           =   2535
         End
         Begin VB.CheckBox Check1 
            Caption         =   "SK-SKIP"
            Height          =   255
            Index           =   11
            Left            =   3030
            TabIndex        =   61
            Top             =   1200
            Width           =   2535
         End
      End
      Begin MSComCtl2.DTPicker tgl_lpd 
         Height          =   375
         Left            =   1080
         TabIndex        =   76
         Top             =   1560
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         CustomFormat    =   "MMMM-yyyy"
         Format          =   98959363
         CurrentDate     =   41610
      End
      Begin MSComCtl2.DTPicker tgl_lpd2 
         Height          =   375
         Left            =   3840
         TabIndex        =   81
         Top             =   1560
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         CustomFormat    =   "MMMM-yyyy"
         Format          =   98959363
         CurrentDate     =   41610
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   3360
         TabIndex        =   89
         Top             =   1680
         Width           =   315
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "LPD"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   77
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Batch"
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
         Left            =   240
         TabIndex        =   75
         Top             =   420
         Width           =   555
      End
   End
   Begin VB.CommandButton CmdFilterExcel 
      Caption         =   "&Filter dari Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2700
      TabIndex        =   54
      Top             =   0
      Width           =   1995
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
      Left            =   10680
      TabIndex        =   36
      Top             =   6120
      Width           =   2355
   End
   Begin VB.CommandButton CmdKembalikanAgent 
      Caption         =   "Kembalikan Ke Agent lama..."
      Height          =   435
      Left            =   8280
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
   Begin VB.CommandButton CmdProses 
      BackColor       =   &H0000FF00&
      Caption         =   "&Approve..."
      Height          =   375
      Left            =   5220
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6210
      Width           =   1515
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
      Width           =   15435
      _ExtentX        =   27226
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
      Calendar        =   "FrmDistribusiAcc.frx":00AF
      Caption         =   "FrmDistribusiAcc.frx":01C7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmDistribusiAcc.frx":0233
      Keys            =   "FrmDistribusiAcc.frx":0251
      Spin            =   "FrmDistribusiAcc.frx":02AF
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
      Caption         =   "FrmDistribusiAcc.frx":02D7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmDistribusiAcc.frx":0343
      Spin            =   "FrmDistribusiAcc.frx":0393
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
      Calendar        =   "FrmDistribusiAcc.frx":03BB
      Caption         =   "FrmDistribusiAcc.frx":04D3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmDistribusiAcc.frx":053F
      Keys            =   "FrmDistribusiAcc.frx":055D
      Spin            =   "FrmDistribusiAcc.frx":05BB
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
      Caption         =   "FrmDistribusiAcc.frx":05E3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmDistribusiAcc.frx":064F
      Spin            =   "FrmDistribusiAcc.frx":069F
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
   Begin VB.CommandButton CmdHapusAgent 
      Caption         =   "&Hapus Agent"
      Height          =   375
      Left            =   13500
      TabIndex        =   14
      Top             =   7800
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "Edit..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   13500
      TabIndex        =   15
      Top             =   7320
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CommandButton CmdCekAllAgent 
      Caption         =   "Cek All"
      Height          =   375
      Left            =   13500
      TabIndex        =   42
      Top             =   8400
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CommandButton CmdUncekallAgent 
      Caption         =   "UnCek All"
      Height          =   375
      Left            =   13500
      TabIndex        =   43
      Top             =   8820
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Filter dari Kriteria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4800
      TabIndex        =   78
      Top             =   0
      Width           =   1995
   End
   Begin TDBDate6Ctl.TDBDate TxtTglExpired 
      Height          =   315
      Left            =   2100
      TabIndex        =   92
      Top             =   6000
      Width           =   1530
      _Version        =   65536
      _ExtentX        =   2699
      _ExtentY        =   556
      Calendar        =   "FrmDistribusiAcc.frx":06C7
      Caption         =   "FrmDistribusiAcc.frx":07DF
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmDistribusiAcc.frx":084B
      Keys            =   "FrmDistribusiAcc.frx":0869
      Spin            =   "FrmDistribusiAcc.frx":08C7
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
   Begin VB.Label Label25 
      Caption         =   "AKSESALL MENUNGGU APPROVAL"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   12840
      TabIndex        =   95
      Top             =   5520
      Width           =   2775
   End
   Begin VB.Label Label24 
      Caption         =   "Exp Date :"
      Height          =   225
      Left            =   1260
      TabIndex        =   91
      Top             =   6060
      Width           =   975
   End
   Begin VB.Label lbl_profile 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Profile >>>"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13320
      TabIndex        =   55
      Top             =   5160
      Width           =   2175
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
      Left            =   210
      TabIndex        =   33
      Top             =   5700
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "Waktu Akhir:"
      Height          =   195
      Left            =   3780
      TabIndex        =   32
      Top             =   5700
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
      Caption         =   $"FrmDistribusiAcc.frx":08EF
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

Private Sub new_kdprofile()
    Dim M_objrs As ADODB.Recordset
    Dim index_profile As Integer
    Dim tglprofile As String
    
    ' ------------ KODE PROFILE 21 MEI 2013 ------------------
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open "SELECT now()", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    tglprofile = Format(M_objrs(0), "yyyymmdd")
    
    If M_objrs.state = 1 Then M_objrs.Close
    M_objrs.Open "SELECT * FROM mandiri.tbl_profile_aksesall"
    If M_objrs.RecordCount > 0 Then
        index_profile = M_objrs.RecordCount
        lbl_profile.Caption = tglprofile & Right("0000" & index_profile + 1, 4)
    Else
        lbl_profile.Caption = tglprofile & "0001"
    End If
    ' ---------------------------------------------------------
    
    Set M_objrs = Nothing
End Sub

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
    Dim cmdsql As String
    Dim M_objrs As ADODB.Recordset
    Dim ListItem As ListItem
    Dim M_WHERE As String
    
    M_WHERE = ""
    
    cmdsql = "select * from mandiri.mgm  "
    
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
        M_WHERE = M_WHERE & " and custid not in (select distinct custid from mandiri.tblsendptp ) "
    Else
        M_WHERE = " where agent not in ('COMPLAIN','LUNAS','CLAIM') "
        M_WHERE = M_WHERE & " and custid not in (select distinct custid from mandiri.tblsendptp ) "
    End If
       
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open cmdsql & M_WHERE, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvAcc.ListItems.CLEAR
    TxtJmlhAcc.Text = M_objrs.RecordCount
    
    
    If M_objrs.RecordCount = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Peringatan"
        Set M_objrs = Nothing
        Exit Sub
    End If
    
    PB1.Max = M_objrs.RecordCount
    While Not M_objrs.EOF
        PB1.Value = M_objrs.Bookmark
        Set ListItem = LvAcc.ListItems.ADD(, , M_objrs("custid"))
            ListItem.SubItems(1) = M_objrs("name")
            ListItem.SubItems(2) = IIf(IsNull(M_objrs("f_cek_new")), "", M_objrs("f_cek_new"))
            ListItem.SubItems(3) = IIf(IsNull(M_objrs("agent")), "", M_objrs("agent"))
            ListItem.SubItems(4) = IIf(IsNull(M_objrs("agent_asli")), "", M_objrs("agent_asli"))
            ListItem.SubItems(5) = IIf(IsNull(M_objrs("monitor_akses")), "", M_objrs("monitor_akses"))
            ListItem.SubItems(6) = IIf(IsNull(M_objrs("waktu_akses")), "", Format(M_objrs("waktu_akses"), "yyyy-mm-dd hh:nn:ss"))
            
            If UCase(M_objrs("agent")) = "AKSESALL" Then
                ListItem.ForeColor = vbRed
                ListItem.ListSubItems(1).ForeColor = vbRed
                ListItem.ListSubItems(2).ForeColor = vbRed
                ListItem.ListSubItems(3).ForeColor = vbRed
                ListItem.ListSubItems(4).ForeColor = vbRed
                ListItem.ListSubItems(5).ForeColor = vbRed
                ListItem.ListSubItems(6).ForeColor = vbRed
            End If
            
            If UCase(M_objrs("agent")) = "#KOSONG#" Then
                ListItem.ForeColor = vbBlue
                ListItem.ListSubItems(1).ForeColor = vbBlue
                ListItem.ListSubItems(2).ForeColor = vbBlue
                ListItem.ListSubItems(3).ForeColor = vbBlue
                ListItem.ListSubItems(4).ForeColor = vbBlue
                ListItem.ListSubItems(5).ForeColor = vbBlue
                ListItem.ListSubItems(6).ForeColor = vbBlue
            End If
            
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
End Sub

Private Sub HeaderAgent()
    LVAgent.ColumnHeaders.ADD 1, , "ID", 500
    LVAgent.ColumnHeaders.ADD 2, , "AGENT", 1000
    LVAgent.ColumnHeaders.ADD 3, , "CUSTID", 2000
    LVAgent.ColumnHeaders.ADD 4, , "WAKTU AWAL", 2000
    LVAgent.ColumnHeaders.ADD 5, , "WAKTU AKHIR", 2000
    LVAgent.ColumnHeaders.ADD 6, , "LOG DISTRIBUSI", 1500
    LVAgent.ColumnHeaders.ADD 7, , "WAKTU DISTRIBUSI", 2000
    LVAgent.ColumnHeaders.ADD 8, , "KODE PROFILE", 2000
End Sub




Private Sub Check2_Click()
    If Check2.Value = 1 Then
        List1.Enabled = True
    Else
        List1.Enabled = False
    End If
End Sub

Private Sub Check3_Click()
    If Check3.Value = 1 Then
        Frame4.Enabled = True
    Else
        Frame4.Enabled = False
    End If
End Sub

Private Sub CmbAgent_Click()
    Dim cmdsql As String
    Dim M_objrs As ADODB.Recordset
    Dim ListItem As ListItem
    Dim GroupingTL_2 As String
    
    GroupingTL_2 = ""
    
    '@@19022013 Tambahan ini buat grouping TL
    If UCase(Mid(CmbAgent.Text, 1, 2)) = "TL" Then
        GroupingTL_2 = " agent in (select userid from mandiri.usertbl where spvcode in ("
        GroupingTL_2 = GroupingTL_2 & " select spvcode from mandiri.usertbl where userid='"
        GroupingTL_2 = GroupingTL_2 & CmbAgent.Text + "')) "
    Else
        GroupingTL_2 = " agent='"
        GroupingTL_2 = GroupingTL_2 + CmbAgent.Text + "' "
    End If
    
    If CmbAgent.Text <> "ALL" Then
        'Cmdsql = "select * from tbl_distribusi_account where " & GroupingTL_2
        'Cmdsql = Cmdsql & CmbAgent.Text & "' order by waktu_awal asc "
    Else
        cmdsql = "select * from mandiri.tbl_distribusi_account "
        cmdsql = cmdsql & " order by agent,waktu_awal asc "
    End If
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    TxtJmlhAgent.Text = M_objrs.RecordCount
    LVAgent.ListItems.CLEAR
    
    If M_objrs.RecordCount = 0 Then
        
        Set M_objrs = Nothing
        Exit Sub
    End If
    
    While Not M_objrs.EOF
        Set ListItem = LVAgent.ListItems.ADD(, , M_objrs("id"))
            ListItem.SubItems(1) = M_objrs("agent")
            ListItem.SubItems(2) = M_objrs("custid")
            ListItem.SubItems(3) = Format(M_objrs("waktu_awal"), "yyyy-mm-dd hh:nn:ss")
            ListItem.SubItems(4) = Format(M_objrs("waktu_akhir"), "yyyy-mm-dd hh:nn:ss")
            ListItem.SubItems(5) = M_objrs("log_distribusi")
            ListItem.SubItems(6) = Format(M_objrs("log_tgl_distribusi"), "yyyy-mm-dd hh:nn:ss")
            ListItem.SubItems(7) = Format(M_objrs("log_tgl_distribusi"), "yyyy-mm-dd hh:nn:ss")
       M_objrs.MoveNext
    Wend
    
    Set M_objrs = Nothing
End Sub

Private Sub CmbFilterAcc_Click()
    'Call CariFilter
    'Eventnya diambil berdasarkan
End Sub



Private Sub CmbStatusAcc_Click()
    Dim cmdsql As String
    Dim M_objrs As ADODB.Recordset
    Dim ListItem As ListItem
    Dim Bulan, Tahun, Tanggal As String
    Dim M_WHERE As String
    Dim GroupingTL As String
    
    If CmbFilterAcc.Text = "" Then
        'MsgBox "Pilih terlebih dahulu agentnya!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If

    M_WHERE = ""
    GroupingTL = ""
    
    cmdsql = "select * from mandiri.mgm "
    
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
                    M_WHERE = " where custid in (select distinct custid from mandiri.tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from mandiri.tbllunas "
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
                    M_WHERE = " where custid in (select distinct custid from mandiri.tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from mandiri.tbllunas "
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
                    M_WHERE = " where custid in (select distinct custid from mandiri.tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from mandiri.tbllunas "
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
                    M_WHERE = " where custid in (select distinct custid from mandiri.tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)<'"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)<='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from mandiri.tbllunas "
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
            GroupingTL = " agent in (select userid from mandiri.usertbl where spvcode in ("
            GroupingTL = GroupingTL & " select spvcode from mandiri.usertbl where userid='"
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
                    M_WHERE = " where custid in (select distinct custid from mandiri.tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from mandiri.tbllunas "
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
                    M_WHERE = " where custid in (select distinct custid from mandiri.tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and  " & GroupingTL
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from mandiri.tbllunas "
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
                    M_WHERE = " where custid in (select distinct custid from mandiri.tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                    
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from mandiri.tbllunas "
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
                    M_WHERE = " where custid in (select distinct custid from mandiri.tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)<'"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)<='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from mandiri.tbllunas "
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
        M_WHERE = M_WHERE & " and custid not in (select distinct custid from mandiri.tblsendptp ) "
        M_WHERE = M_WHERE & " order by name asc "
    Else
        M_WHERE = M_WHERE & " and agent not in ('LUNAS','COMPLAIN','CLAIM') "
        M_WHERE = M_WHERE & " and custid not in (select distinct custid from mandiri.tblsendptp ) "
        M_WHERE = M_WHERE & " order by name asc "
    End If
    
    
    
    DoEvents
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open cmdsql + M_WHERE, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvAcc.ListItems.CLEAR
    TxtJmlhAcc.Text = M_objrs.RecordCount
    
    
    If M_objrs.RecordCount = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Peringatan"
        Set M_objrs = Nothing
        Exit Sub
    End If
    
    PB1.Max = M_objrs.RecordCount
    While Not M_objrs.EOF
        PB1.Value = M_objrs.Bookmark
        Set ListItem = LvAcc.ListItems.ADD(, , M_objrs("custid"))
            ListItem.SubItems(1) = M_objrs("name")
            ListItem.SubItems(2) = IIf(IsNull(M_objrs("f_cek_new")), "", M_objrs("f_cek_new"))
            ListItem.SubItems(3) = IIf(IsNull(M_objrs("agent")), "", M_objrs("agent"))
            ListItem.SubItems(4) = IIf(IsNull(M_objrs("agent_asli")), "", M_objrs("agent_asli"))
            ListItem.SubItems(5) = IIf(IsNull(M_objrs("monitor_akses")), "", M_objrs("monitor_akses"))
            ListItem.SubItems(6) = IIf(IsNull(M_objrs("waktu_akses")), "", Format(M_objrs("waktu_akses"), "yyyy-mm-dd hh:nn:ss"))
            
            If UCase(M_objrs("agent")) = "AKSESALL" Then
                ListItem.ForeColor = vbRed
                ListItem.ListSubItems(1).ForeColor = vbRed
                ListItem.ListSubItems(2).ForeColor = vbRed
                ListItem.ListSubItems(3).ForeColor = vbRed
                ListItem.ListSubItems(4).ForeColor = vbRed
                ListItem.ListSubItems(5).ForeColor = vbRed
                ListItem.ListSubItems(6).ForeColor = vbRed
            End If
            
            If UCase(M_objrs("agent")) = "#KOSONG#" Then
                ListItem.ForeColor = vbBlue
                ListItem.ListSubItems(1).ForeColor = vbBlue
                ListItem.ListSubItems(2).ForeColor = vbBlue
                ListItem.ListSubItems(3).ForeColor = vbBlue
                ListItem.ListSubItems(4).ForeColor = vbBlue
                ListItem.ListSubItems(5).ForeColor = vbBlue
                ListItem.ListSubItems(6).ForeColor = vbBlue
            End If
            
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
    
End Sub



Private Sub CmbStatusCollBersama_Click()
    Dim cmdsql As String
    Dim M_objrs As ADODB.Recordset
    Dim ListItem As ListItem
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
    
    cmdsql = "select * from mandiri.mgm "
    
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
                    M_WHERE = " where custid in (select distinct custid from mandiri.tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from mandiri.tbllunas "
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
                    M_WHERE = " where custid in (select distinct custid from mandiri.tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from mandiri.tbllunas "
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
                    M_WHERE = " where custid in (select distinct custid from mandiri.tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from mandiri.tbllunas "
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
                    M_WHERE = " where custid in (select distinct custid from mandiri.tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)<'"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)<='"
                    M_WHERE = M_WHERE & CStr(Tahun) & "') "
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from mandiri.tbllunas "
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
            GroupingTL = " agent_asli in (select userid from mandiri.usertbl where spvcode in ("
            GroupingTL = GroupingTL & " select spvcode from mandiri.usertbl where userid='"
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
                    M_WHERE = " where custid in (select distinct custid from mandiri.tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from mandiri.tbllunas "
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
                    M_WHERE = " where custid in (select distinct custid from mandiri.tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and  " & GroupingTL
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from mandiri.tbllunas "
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
                    M_WHERE = " where custid in (select distinct custid from mandiri.tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)='"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                    
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from mandiri.tbllunas "
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
                    M_WHERE = " where custid in (select distinct custid from mandiri.tbllunas "
                    M_WHERE = M_WHERE + " where date_part('month',paydate)<'"
                    M_WHERE = M_WHERE & CStr(Bulan) & "' and date_part('year',paydate)<='"
'                    M_WHERE = M_WHERE & CStr(Tahun) & "') and agent='"
'                    M_WHERE = M_WHERE & CmbFilterAcc.Text & "' "
                    M_WHERE = M_WHERE & CStr(Tahun) & "') and " & GroupingTL
                Else
                    M_WHERE = M_WHERE & " and custid in (select distinct custid from mandiri.tbllunas "
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
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open cmdsql + M_WHERE, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvAcc.ListItems.CLEAR
    TxtJmlhAcc.Text = M_objrs.RecordCount
    
    
    If M_objrs.RecordCount = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Peringatan"
        Set M_objrs = Nothing
        Exit Sub
    End If
    
    PB1.Max = M_objrs.RecordCount
    While Not M_objrs.EOF
        PB1.Value = M_objrs.Bookmark
        Set ListItem = LvAcc.ListItems.ADD(, , M_objrs("custid"))
            ListItem.SubItems(1) = M_objrs("name")
            ListItem.SubItems(2) = IIf(IsNull(M_objrs("f_cek_new")), "", M_objrs("f_cek_new"))
            ListItem.SubItems(3) = IIf(IsNull(M_objrs("agent")), "", M_objrs("agent"))
            ListItem.SubItems(4) = IIf(IsNull(M_objrs("agent_asli")), "", M_objrs("agent_asli"))
            ListItem.SubItems(5) = IIf(IsNull(M_objrs("monitor_akses")), "", M_objrs("monitor_akses"))
            ListItem.SubItems(6) = IIf(IsNull(M_objrs("waktu_akses")), "", Format(M_objrs("waktu_akses"), "yyyy-mm-dd hh:nn:ss"))
            
            If UCase(M_objrs("agent")) = "AKSESALL" Then
                ListItem.ForeColor = vbRed
                ListItem.ListSubItems(1).ForeColor = vbRed
                ListItem.ListSubItems(2).ForeColor = vbRed
                ListItem.ListSubItems(3).ForeColor = vbRed
                ListItem.ListSubItems(4).ForeColor = vbRed
                ListItem.ListSubItems(5).ForeColor = vbRed
                ListItem.ListSubItems(6).ForeColor = vbRed
            End If
            
            If UCase(M_objrs("agent")) = "#KOSONG#" Then
                ListItem.ForeColor = vbBlue
                ListItem.ListSubItems(1).ForeColor = vbBlue
                ListItem.ListSubItems(2).ForeColor = vbBlue
                ListItem.ListSubItems(3).ForeColor = vbBlue
                ListItem.ListSubItems(4).ForeColor = vbBlue
                ListItem.ListSubItems(5).ForeColor = vbBlue
                ListItem.ListSubItems(6).ForeColor = vbBlue
            End If
            
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
End Sub

Private Sub cmd_insert_antrian_Click()
    Dim iQuery As String
    Dim CustId, nama_customer, status_account, agent_saat_ini, agent_terdahulu, kode_profile As String
    Dim AGENT, tanggal_awal, waktu_awal, tanggal_akhir, waktu_akhir As String
    Dim i As Integer
    Dim S As Long
    
    If LvAcc.ListItems.Count = 0 Then
        MsgBox "Data Customer Tidak Tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If TxtAgent.Text = "" Then
        MsgBox "Agent Tidak Boleh Kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If TxtTglAwal.ValueIsNull Then
        MsgBox "Tanggal Mulai Aksesall Tidak Boleh Kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If TxtWaktuAwal.ValueIsNull Then
        MsgBox "Waktu Mulai Aksesall Tidak Boleh Kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If TxtTglAkhir.ValueIsNull Then
        MsgBox "Tanggal Selesai Aksesall Tidak Boleh Kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If TxtWaktuAkhir.ValueIsNull Then
        MsgBox "Waktu Selesai Aksesall Tidak Boleh Kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    M_OBJCONN.Execute "DELETE FROM mandiri.temp_approval_aksesall"
    
    PB1.Max = LvAcc.ListItems.Count
    For S = 1 To LvAcc.ListItems.Count
        PB1.Value = S
        Me.MousePointer = vbHourglass
        Me.Enabled = False
        CustId = LvAcc.ListItems(S).Text
        nama_customer = LvAcc.ListItems(S).ListSubItems(1)
        status_account = LvAcc.ListItems(S).ListSubItems(2)
        agent_saat_ini = LvAcc.ListItems(S).ListSubItems(3)
        agent_terdahulu = LvAcc.ListItems(S).ListSubItems(4)
        kode_profile = lbl_profile.Caption
        AGENT = Replace(TxtAgent.Text, "'", "''")
        tanggal_awal = Format(TxtTglAwal.Value, "dd/mm/yyyy")
        waktu_awal = Format(TxtWaktuAwal.Value, "hh:nn")
        tanggal_akhir = Format(TxtTglAkhir.Value, "dd/mm/yyyy")
        waktu_akhir = Format(TxtWaktuAkhir.Value, "hh:nn")

        iQuery = "INSERT INTO mandiri.temp_approval_aksesall"
        iQuery = iQuery + " VALUES ('" & CustId & "', '" & nama_customer & "', '" & status_account & "', "
        iQuery = iQuery + " '" & agent_saat_ini & "', '" & agent_terdahulu & "',  '" & kode_profile & "',  "
        iQuery = iQuery + " '" & AGENT & "', '" & tanggal_awal & "', '" & waktu_awal & "', '" & tanggal_akhir & "', "
        iQuery = iQuery + " '" & waktu_akhir & "')"
        
        M_OBJCONN.Execute iQuery
    Next S
     
    MsgBox "Aksesall Berhasil Diproses, Menunggu Approve Dari SPV Atau Manager!", vbOKOnly + vbInformation, "Informasi"
        Me.MousePointer = vbNormal
        Me.Enabled = True
    Unload Me

End Sub

Private Sub CmdBukaAccount_Click()
    Dim cmdsql As String
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
            cmdsql = "update mandiri.mgm set monitor_akses=null,waktu_akses=null where custid='"
            cmdsql = cmdsql & CStr(LvAcc.ListItems(W).Text) & "'"
            M_OBJCONN.Execute cmdsql
        End If
    Next W
    
    Call IsiAccount
    
    MsgBox "Proses berhasil!", vbOKOnly + vbInformation, "Informasi"
    
End Sub

Private Sub CmdCari_Click()
    Call IsiAccount
End Sub

Private Sub CmdCekAllAcc_Click()
    Dim W As Long
    
    If LvAcc.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LvAcc.ListItems.Count
        LvAcc.ListItems(W).Checked = True
    Next W
    TxtJmlhAcc.Text = LvAcc.ListItems.Count
End Sub

Private Sub CmdCekAllAgent_Click()
    Dim W As Integer
    
    If LVAgent.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LVAgent.ListItems.Count
        LVAgent.ListItems(W).Checked = True
    Next W
End Sub

Private Sub CmdClear_Click()
    TxtCariCustid.Text = ""
    TxtCariNama.Text = ""
End Sub

Private Sub cmdEdit_Click()
    If LvAcc.ListItems.Count = 0 Then
        MsgBox "Data Account tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If LVAgent.ListItems.Count = 0 Then
        MsgBox "Data Agent tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    FrmEditDistribusiAccount.txtId.Text = LVAgent.SelectedItem.Text
    FrmEditDistribusiAccount.TxtAgent.Text = LVAgent.SelectedItem.SubItems(1)
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
    Dim cmdsql As String
    Dim W, i, K As Integer
    Dim M_objrs As ADODB.Recordset
    
    If LVAgent.ListItems.Count = 0 Then
        MsgBox "Data agent tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    a = MsgBox("Yakin data agent akan dihapus?", vbYesNo + vbInformation, "Informasi")
    If a = vbNo Then
        MsgBox "Proses dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    i = 0
    For W = 1 To LVAgent.ListItems.Count
       If LVAgent.ListItems(W).Checked = True Then
            i = i + 1
       End If
    Next W
    
    
    If i = 0 Then
        MsgBox "Anda belum memilih data agent yang akan dihapus!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
        
    DoEvents
        
    PB1.Max = LVAgent.ListItems.Count
        
    For K = 1 To LVAgent.ListItems.Count
        PB1.Value = K
        If LVAgent.ListItems(K).Checked = True Then
            cmdsql = "DELETE FROM mandiri.tbl_cust_not_aksesall WHERE kd_profile='" & LVAgent.ListItems(K).SubItems(7) & "' " & _
                    "AND custid='" & LVAgent.ListItems(K).SubItems(2) & "' AND agent='" & LVAgent.ListItems(K).SubItems(1) & "'"
            M_OBJCONN.Execute cmdsql
            
            cmdsql = "INSERT INTO mandiri.tbl_cust_not_aksesall(kd_profile,custid,agent) " & _
                    "VALUES('" & LVAgent.ListItems(K).SubItems(7) & "','" & LVAgent.ListItems(K).SubItems(2) & "' & LvAgent.ListItems(K).SubItems(1) & " ')"
'            Cmdsql = "delete from tbl_distribusi_account where id='"
'            Cmdsql = Cmdsql + CStr(LvAgent.ListItems(K).Text) + "'"
            M_OBJCONN.Execute cmdsql
'
'            'Update status agentnya nih
'            Cmdsql = "update usertbl set f_akses_all_acc=null,f_pesanresetauto='1' "
'            Cmdsql = Cmdsql + " where userid='"
'            Cmdsql = Cmdsql + CStr(LvAgent.ListItems(K).SubItems(1)) + "'"
'            M_OBJCONN.Execute Cmdsql
'
'            'Cek apakah custid ini sudah habis agentnya?
'            Cmdsql = "select * from tbl_distribusi_account where custid='"
'            Cmdsql = Cmdsql & CStr(LvAgent.ListItems(K).SubItems(2)) & "'"
'            Set M_Objrs = New ADODB.Recordset
'            M_Objrs.CursorLocation = adUseClient
'            M_Objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'            If M_Objrs.RecordCount = 0 Then
'                'Update ke agent yang lama
'                Cmdsql = "update mgm set agent=agent_asli,agent_asli=null,"
'                Cmdsql = Cmdsql + " user_claim=null,waktu_claim=null,alasan_claim=null "
'                Cmdsql = Cmdsql + " where custid='"
'                Cmdsql = Cmdsql & CStr(LvAgent.ListItems(K).SubItems(2)) + "' and agent_asli is not null "
'                M_OBJCONN.Execute Cmdsql
'            End If
'            Set M_Objrs = Nothing
            
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
    Dim cmdsql As String
    Dim W, K, S As Integer
    Dim a As String
    Dim M_objrs As ADODB.Recordset
    
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
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    
    PB1.Max = LvAcc.ListItems.Count
    For W = 1 To LvAcc.ListItems.Count
        PB1.Value = W
        If LvAcc.ListItems(W).Checked = True Then
            If M_objrs.state = 1 Then M_objrs.Close
            cmdsql = "SELECT * FROM mandiri.tbllog_claim_aksesall WHERE custid='" & CStr(LvAcc.ListItems(W).Text) & "'"
            M_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If M_objrs.RecordCount = 0 Then
                ' 19 AGUSTUS 2014 agent_asli=null dihilangkan
                cmdsql = "update mandiri.mgm set agent=agent_asli where custid='"
                cmdsql = cmdsql + CStr(LvAcc.ListItems(W).Text) + "' and agent_asli is not null "
                M_OBJCONN.Execute cmdsql
            End If
            'Hapus data di tabel distribusinya
'            Cmdsql = "delete from tbl_distribusi_account where custid='"
'            Cmdsql = Cmdsql + CStr(LvAcc.ListItems(W).Text) + "'"
            cmdsql = "DELETE FROM mandiri.tbl_cust_aksesall WHERE custid='" & CStr(LvAcc.ListItems(W).Text) & "'"
            M_OBJCONN.Execute cmdsql
        End If
    Next W
    
    Set M_objrs = Nothing
    
    Call IsiAccount
    
    MsgBox "Account berhasil dikembalikan ke agent awal!", vbOKOnly + vbInformation, "Informasi"
End Sub

Private Sub CmdLihatListAgent_Click()
    FrmListAgent.Show vbModal
End Sub

Private Sub cmdProses_Click()
    Dim cmdsql, AmbilCustid As String
    Dim M_objrs As ADODB.Recordset
    Dim W, K, S As Long
    Dim Tanggal1 As String
    Dim Tanggal2 As String
    Dim pesan As String
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
        
    'MDIForm1.Timer1 = False
    MDIForm1.TimerCTI = False
    'MDIForm1.TimerBlink = False
    'Cek waktu server
    cmdsql = "select now()"
    Set M_ObjrsWaktuServer = New ADODB.Recordset
    M_ObjrsWaktuServer.CursorLocation = adUseClient
    DoEvents
    M_ObjrsWaktuServer.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    WaktuServer = Format(M_ObjrsWaktuServer(0), "m/dd/yyyy hh:nn:ss")
    
    

    If TxtTglAwal.ValueIsNull = True Or _
       TxtWaktuAwal.ValueIsNull = True Or _
       TxtTglAkhir.ValueIsNull = True Or _
       TxtWaktuAkhir.ValueIsNull = True Or _
       TxtTglExpired.ValueIsNull = True Then
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
    
    'On Error GoTo SALAH
    
    'Ambil Data agent
    cmdsql = "select * from mandiri.usertbl where userid in ("
    cmdsql = cmdsql + TxtAgent.Text + ")"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

    If M_objrs.RecordCount > 0 Then
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

        While Not M_objrs.EOF
            DoEvents
            'Update status Agentnya
            cmdsql = "update mandiri.usertbl set f_akses_all_acc='1',f_pesanresetauto='1',profile_akses_all='" & lbl_profile.Caption & "' "
            cmdsql = cmdsql + " where userid='"
            cmdsql = cmdsql + M_objrs("userid") + "'"
            M_OBJCONN.Execute cmdsql

            'Kirim pesan ke agent
            pesan = "Pesan dibuat otomatis oleh system!" & vbCrLf
            pesan = pesan & "----------------------------------------------" & vbCrLf
            pesan = pesan & "SPV menambahkan account baru untuk anda. " & vbCrLf
            pesan = pesan & "Account ini dapat di collect secara bersama-sama oleh anda, " & vbCrLf
            pesan = pesan & "mulai dari :" & Format(Tanggal1, "yyyy-mm-dd hh:nn:ss") & " s.d. " & vbCrLf
            pesan = pesan & Format(Tanggal2, "yyyy-mm-dd hh:nn:ss") & vbCrLf
            pesan = pesan & "Cek account baru anda dengan mengklik ulang tombol search data!"

            cmdsql = "insert into mandiri.msgtbl "
            cmdsql = cmdsql + "( recipient, datetime, sender, sentfrom, msg) values ('"
            cmdsql = cmdsql + M_objrs("userid") + "','"
            cmdsql = cmdsql + Format(Now(), "yyyymmdd") + "','"
            cmdsql = cmdsql + MDIForm1.txtusername.Text + "','"
            cmdsql = cmdsql + CStr(MDIForm1.Winsock1.LocalIP) + "','"
            cmdsql = cmdsql + pesan + "')"

            M_OBJCONN.Execute cmdsql
            M_objrs.MoveNext
        Wend
    End If
    Set M_objrs = Nothing
    
    ' Hapus Header
    cmdsql = "delete from mandiri.tbl_profile_aksesall where kd_profile='" & lbl_profile.Caption & "'"
    M_OBJCONN.Execute cmdsql
    
    ' Hapus detail
    cmdsql = "delete from mandiri.tbl_cust_aksesall where kd_profile='" & lbl_profile.Caption & "'"
    M_OBJCONN.Execute cmdsql

    ' Insert Header
    cmdsql = "INSERT INTO mandiri.tbl_profile_aksesall(kd_profile,waktu_awal,waktu_akhir,log_distribusi,log_tgl_distribusi) VALUES"
    cmdsql = cmdsql & "('" & lbl_profile.Caption & "', '"
    cmdsql = cmdsql & Format(TxtTglAwal.Value, "yyyy-mm-dd") & " " & Format(TxtWaktuAwal.Value) & "','"
    cmdsql = cmdsql & Format(TxtTglAkhir.Value, "yyyy-mm-dd") & " " & Format(TxtWaktuAkhir.Value) & "','"
    cmdsql = cmdsql & MDIForm1.txtusername.Text & "',"
    cmdsql = cmdsql & " now())"
    M_OBJCONN.Execute cmdsql
    
    PB1.Max = LvAcc.ListItems.Count
    For K = 1 To LvAcc.ListItems.Count
        PB1.Value = K
        DoEvents
        If LvAcc.ListItems(K).Checked = True Then
            ' DELETE DATA CUSTOMER DI PROFILE SEBELUMNYA KLO ADA
            cmdsql = "DELETE FROM mandiri.tbl_cust_aksesall WHERE custid='" & CStr(LvAcc.ListItems(K).Text) & "'"
            M_OBJCONN.Execute cmdsql
            
            ' UPDATE MGM TGL EXPIRED CLAIM
            cmdsql = "update mandiri.mgm set tgl_exp_claim = '" & Format(TxtTglExpired.Value, "yyyy-mm-dd") & "' where custid = '" & CStr(LvAcc.ListItems(K).Text) & "'"
            M_OBJCONN.Execute cmdsql
            
            'Inputkan data ke Detail
            cmdsql = "INSERT INTO mandiri.tbl_cust_aksesall values('" & lbl_profile.Caption & "','" & CStr(LvAcc.ListItems(K).Text) & "')"
            M_OBJCONN.Execute cmdsql
            
            If UCase(LvAcc.ListItems(K).SubItems(3)) = "#KOSONG#" Then
                cmdsql = "UPDATE mandiri.mgm SET agent='AKSESALL' WHERE custid='"
                cmdsql = cmdsql & CStr(LvAcc.ListItems(K).Text) & "'"
                M_OBJCONN.Execute cmdsql
            ElseIf UCase(LvAcc.ListItems(K).SubItems(3)) <> "AKSESALL" Then
                cmdsql = "UPDATE mandiri.mgm SET agent_asli=agent, agent='AKSESALL' WHERE custid='"
                cmdsql = cmdsql & CStr(LvAcc.ListItems(K).Text) & "'"
                M_OBJCONN.Execute cmdsql
            End If
            
            ' ====== UPDATE IZUDDIN 08 OKTOBER 2013 =======
            cmdsql = "SELECT custid,agent_asli, agent FROM mandiri.mgm WHERE custid='" & CStr(LvAcc.ListItems(K).Text) & "'"
            
            Set M_objrs = New ADODB.Recordset
            M_objrs.CursorLocation = adUseClient
            M_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            
            If M_objrs.RecordCount > 0 Then
                M_OBJCONN.Execute "INSERT INTO mandiri.tbl_hst_aksesall(custid,agent_asli) values('" & M_objrs!CustId & "','" & M_objrs!agent_asli & "')"
            End If
            
            Set M_objrs = Nothing
            ' =============================================
        End If
        Debug.Print "loop ke" & CStr(K)
    Next K
    
    'Ambil Data agent
'    Cmdsql = "select * from usertbl where userid in ("
'    Cmdsql = Cmdsql + TxtAgent.Text + ")"
'    Set M_Objrs = New ADODB.Recordset
'    M_Objrs.CursorLocation = adUseClient
'    M_Objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'    If M_Objrs.RecordCount > 0 Then
'        DoEvents
'        LvAcc.Enabled = False
'        CmdCari.Enabled = False
'        cmdClear.Enabled = False
'        CmdLihatListAgent.Enabled = False
'        TxtTglAwal.Enabled = False
'        TxtWaktuAwal.Enabled = False
'        TxtTglAkhir.Enabled = False
'        TxtWaktuAkhir.Enabled = False
'
'        CmbAgentCollBersama.Enabled = False
'        CmbStatusCollBersama.Enabled = False
'
'        PB1.Max = M_Objrs.RecordCount
'        While Not M_Objrs.EOF
'            PB1.Value = M_Objrs.Bookmark
'            DoEvents
'            'Update status Agentnya
'            Cmdsql = "update usertbl set f_akses_all_acc='1',f_pesanresetauto='1' "
'            Cmdsql = Cmdsql + " where userid='"
'            Cmdsql = Cmdsql + M_Objrs("userid") + "'"
'            M_OBJCONN.Execute Cmdsql
'
'            'Kirim pesan ke agent
'            Pesan = "Pesan dibuat otomatis oleh system!" & vbCrLf
'            Pesan = Pesan & "----------------------------------------------" & vbCrLf
'            Pesan = Pesan & "SPV menambahkan account baru untuk anda. " & vbCrLf
'            Pesan = Pesan & "Account ini dapat di collect secara bersama-sama oleh anda, " & vbCrLf
'            Pesan = Pesan & "mulai dari :" & Format(Tanggal1, "yyyy-mm-dd hh:nn:ss") & " s.d. " & vbCrLf
'            Pesan = Pesan & Format(Tanggal2, "yyyy-mm-dd hh:nn:ss") & vbCrLf
'            Pesan = Pesan & "Cek account baru anda dengan mengklik ulang tombol search data!"
'
'            Cmdsql = "insert into msgtbl "
'            Cmdsql = Cmdsql + "( recipient, datetime, sender, sentfrom, msg) values ('"
'            Cmdsql = Cmdsql + M_Objrs("userid") + "','"
'            Cmdsql = Cmdsql + Format(Now(), "yyyymmdd") + "','"
'            Cmdsql = Cmdsql + mdiform1.txtusername.text + "','"
'            Cmdsql = Cmdsql + CStr(MDIForm1.Winsock1.LocalIP) + "','"
'            Cmdsql = Cmdsql + Pesan + "')"
'
'            M_OBJCONN.Execute Cmdsql
'
'
'            For K = 1 To LvAcc.ListItems.Count
'                DoEvents
'                If LvAcc.ListItems(K).Checked = True Then
'                    'Hapus dulu jika ada data sebelumnya
'                    Cmdsql = "delete from tbl_distribusi_account where custid='"
'                    Cmdsql = Cmdsql & CStr(LvAcc.ListItems(K).Text) & "' and agent='"
'                    Cmdsql = Cmdsql & M_Objrs("userid") & "'"
'                    M_OBJCONN.Execute Cmdsql
'
'                    'Inputkan data ke database
'                    Cmdsql = "insert into tbl_distribusi_account (custid,agent,"
'                    Cmdsql = Cmdsql & "waktu_awal,waktu_akhir, log_distribusi,"
'                    Cmdsql = Cmdsql & "log_tgl_distribusi) values ('"
'                    Cmdsql = Cmdsql & CStr(LvAcc.ListItems(K).Text) + "','"
'                    Cmdsql = Cmdsql & M_Objrs("userid") & "','"
'                    Cmdsql = Cmdsql & Format(TxtTglAwal.Value, "yyyy-mm-dd") & " " & Format(TxtWaktuAwal.Value) & "','"
'                    Cmdsql = Cmdsql & Format(TxtTglAkhir.Value, "yyyy-mm-dd") & " " & Format(TxtWaktuAkhir.Value) & "','"
'                    Cmdsql = Cmdsql & mdiform1.txtusername.text & "',"
'                    Cmdsql = Cmdsql & " now())"
'                    M_OBJCONN.Execute Cmdsql
'
'                End If
'                Debug.Print "loop ke" & CStr(K)
'            Next K
'            Debug.Print "record ke" & M_Objrs.Bookmark
'            M_Objrs.MoveNext
'        Wend
'
'        'Catet agent lama
'        PB1.Max = LvAcc.ListItems.Count
'        For K = 1 To LvAcc.ListItems.Count
'        DoEvents
'            PB1.Value = K
'            If LvAcc.ListItems(K).Checked = True Then
'                '@@12022013 Jika status account sebelumnya #KOSONG# atau AKSESALL, ga usah diupdate
'                If UCase(LvAcc.ListItems(K).SubItems(3)) = "#KOSONG#" Then
'                    Cmdsql = "update mgm set  agent='AKSESALL' where custid='"
'                    Cmdsql = Cmdsql & CStr(LvAcc.ListItems(K).Text) & "'"
'                    M_OBJCONN.Execute Cmdsql
'                ElseIf UCase(LvAcc.ListItems(K).SubItems(3)) <> "AKSESALL" Then
'                    Cmdsql = "update mgm set agent_asli=agent, agent='AKSESALL' where custid='"
'                    Cmdsql = Cmdsql & CStr(LvAcc.ListItems(K).Text) & "'"
'                    M_OBJCONN.Execute Cmdsql
'
'                    '@@18022013 Ini buat, inputin otomatis buat pemilik agent masuk juga akses all
'                    'Hapus dulu jika ada data sebelumnya
'                    Cmdsql = "delete from tbl_distribusi_account where custid='"
'                    Cmdsql = Cmdsql & CStr(LvAcc.ListItems(K).Text) & "' and agent='"
'                    Cmdsql = Cmdsql & CStr(LvAcc.ListItems(K).SubItems(3)) & "'"
'                    M_OBJCONN.Execute Cmdsql
'
'                    'Inputkan data ke database
'                    Cmdsql = "insert into tbl_distribusi_account (custid,agent,"
'                    Cmdsql = Cmdsql & "waktu_awal,waktu_akhir, log_distribusi,"
'                    Cmdsql = Cmdsql & "log_tgl_distribusi) values ('"
'                    Cmdsql = Cmdsql & CStr(LvAcc.ListItems(K).Text) + "','"
'                    Cmdsql = Cmdsql & CStr(LvAcc.ListItems(K).SubItems(3)) & "','"
'                    Cmdsql = Cmdsql & Format(TxtTglAwal.Value, "yyyy-mm-dd") & " " & Format(TxtWaktuAwal.Value) & "','"
'                    Cmdsql = Cmdsql & Format(TxtTglAkhir.Value, "yyyy-mm-dd") & " " & Format(TxtWaktuAkhir.Value) & "','"
'                    Cmdsql = Cmdsql & mdiform1.txtusername.text & "',"
'                    Cmdsql = Cmdsql & " now())"
'                    M_OBJCONN.Execute Cmdsql
'
'                    'Update statusnya
'                    Cmdsql = "update usertbl set f_akses_all_acc='1', f_pesanresetauto='1' where "
'                    Cmdsql = Cmdsql + " userid='"
'                    Cmdsql = Cmdsql + CStr(LvAcc.ListItems(K).SubItems(3)) + "'"
'                    M_OBJCONN.Execute Cmdsql
'
'                End If
'            End If
'            Debug.Print "SET agent ke " & CStr(K)
'        Next K
'    End If
    
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
    
    Call new_kdprofile
    
    ' Hidupkan TIMER
    'MDIForm1.Timer1 = True
    MDIForm1.TimerCTI = True
    'MDIForm1.TimerBlink = True
    
    MsgBox "Proses berhasil!", vbOKOnly + vbInformation, "Informasi"
    M_OBJCONN.Execute "DELETE FROM mandiri.temp_approval_aksesall"
    Label25.Visible = False
    Exit Sub
'SALAH:
'    MsgBox "Mohon maaf ada kesalahan: " & err.Description, "Error"
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
    TxtJmlhAcc.Text = 0
End Sub

Private Sub CmdUncekallAgent_Click()
        Dim W As Integer
    
    If LVAgent.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LVAgent.ListItems.Count
        LVAgent.ListItems(W).Checked = False
    Next W
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Command1_Click()
    Dim M_objrs As ADODB.Recordset
    Dim cmdsql As String
    Dim sqlfilter As String
    Dim pil_check As Boolean
    Dim pil_batch As Boolean
    Dim list_sts As String
    Dim list_btch As String
    Dim sts_arr() As String
    Dim sts_WO As String
    
    pil_check = False
    
    sqlfilter = ""
    
'    If cb_batch.Text <> "" Then
'        sqlfilter = " AND trim(RECSOURCE)='" & Trim(cb_batch.Text) & "' "
'    End If

    For i = 0 To list_batch.ListCount - 1
        If list_batch.Selected(i) = True Then
            pil_batch = True
            list_btch = list_btch & "'" & list_batch.list(i) & "',"
        End If
    Next i
    
    If pil_batch = True Then
        list_btch = Mid(list_btch, 1, Len(list_btch) - 1)
        sqlfilter = " AND trim(RECSOURCE) in (" & list_btch & ") "
    End If
    
    If IsDate(tgl_lpd.Value) And IsDate(tgl_lpd2.Value) Then
        If tgl_lpd.Value < tgl_lpd2.Value Then
            sqlfilter = sqlfilter & " AND ( CASE WHEN pay_dt_update IS NOT NULL THEN " & _
                        "date(pay_dt_update) between '" & Format(tgl_lpd.Value, "yyyy-mm-01") & "' AND '" & Format(DateAdd("d", -1, DateAdd("m", 1, Format(tgl_lpd2.Value, "yyyy-mm-01"))), "yyyy-mm-dd") & "' " & _
                        "ELSE date(pay_dt) between '" & Format(tgl_lpd.Value, "yyyy-mm-01") & "' AND '" & Format(DateAdd("d", -1, DateAdd("m", 1, Format(tgl_lpd2.Value, "yyyy-mm-01"))), "yyyy-mm-dd") & "' END ) "
        Else
            MsgBox "LPD 1 harus lebih kecil dari LPD 2", vbCritical + vbInformation, "INFO"
            Exit Sub
        End If
    End If
    
    list_sts = ""
    For i = 0 To Check1.UBound
        If Check1(i).Value = 1 Then
            pil_check = True
            sts_arr = Split(Check1(i).Caption, "-")
            list_sts = list_sts & "'" & sts_arr(0) & "-',"
        End If
    Next i
    
    If pil_check = True Then
        list_sts = Mid(list_sts, 1, Len(list_sts) - 1)
        sqlfilter = sqlfilter & " AND substring(f_cek_new,1,3) in (" & list_sts & ") "
    End If
    
    sts_WO = ""
    If Check2.Value = 1 Then
        For i = List1.ListCount - 1 To 0 Step -1
            If List1.Selected(i) = True Then
                sts_WO = sts_WO & List1.list(i) & ","
            End If
        Next i
        
        If Trim(sts_WO) <> "" Then
            sts_WO = Mid(sts_WO, 1, Len(sts_WO) - 1)
            sqlfilter = sqlfilter & " AND (date_part('year',b_d) in (" & sts_WO & ")) "
        End If
    End If
    
    If Check3.Value = 1 Then
        For i = 0 To SSOption1.UBound
            If SSOption1(i).Value = True Then
                sts_arr = Split(SSOption1(i).Caption, "-")
                sqlfilter = sqlfilter & " AND (curbal >=" & Replace(Trim(sts_arr(0)), ".", "") & " AND curbal <= " & Replace(Trim(sts_arr(1)), ".", "") & " ) "
            End If
        Next i
    End If
    
    ' 20 AGUSTUS 2014 - Review tidak di akses all
    cmdsql = "SELECT * FROM mandiri.mgm WHERE custid is not null " & sqlfilter
    cmdsql = cmdsql & " AND agent NOT IN ('LUNAS','COMPLAIN','CLAIM','AKSESALL','REVIEW','REVIEW1','REVIEW2','REVIEW3','REVIEW4','REVIEW5','REVIEW6','REVIEW7','REVIEW8','REVIEW9','REVIEW10') AND coalesce(agent,'')<>'' "
    cmdsql = cmdsql & " AND custid NOT IN (select distinct custid from mandiri.tblsendptp ) "
    ' TAMBAHAN AGAR CLASS 835 TIDAK KENA AKSES ALL
    If Check_decease.Value = 0 Then
        cmdsql = cmdsql & " AND coalesce(cust_class,'')<>'835' "
    End If
    ' -------------------------------------------
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    LvAcc.ListItems.CLEAR
    If M_objrs.RecordCount > 0 Then
        With Me
            .PB1.Max = M_objrs.RecordCount
            While Not M_objrs.EOF
                .PB1.Value = M_objrs.Bookmark
                Set ListItem = .LvAcc.ListItems.ADD(, , M_objrs("custid"))
                ListItem.SubItems(1) = M_objrs("name")
                ListItem.SubItems(2) = IIf(IsNull(M_objrs("f_cek_new")), "", M_objrs("f_cek_new"))
                ListItem.SubItems(3) = IIf(IsNull(M_objrs("agent")), "", M_objrs("agent"))
                ListItem.SubItems(4) = IIf(IsNull(M_objrs("agent_asli")), "", M_objrs("agent_asli"))
                ListItem.SubItems(5) = IIf(IsNull(M_objrs("monitor_akses")), "", M_objrs("monitor_akses"))
                ListItem.SubItems(6) = IIf(IsNull(M_objrs("waktu_akses")), "", Format(M_objrs("waktu_akses"), "yyyy-mm-dd hh:nn:ss"))
            
                If UCase(M_objrs("agent")) = "AKSESALL" Then
                    ListItem.ForeColor = vbRed
                    ListItem.ListSubItems(1).ForeColor = vbRed
                    ListItem.ListSubItems(2).ForeColor = vbRed
                    ListItem.ListSubItems(3).ForeColor = vbRed
                    ListItem.ListSubItems(4).ForeColor = vbRed
                    ListItem.ListSubItems(5).ForeColor = vbRed
                    ListItem.ListSubItems(6).ForeColor = vbRed
                End If
            
                If UCase(M_objrs("agent")) = "#KOSONG#" Then
                    ListItem.ForeColor = vbBlue
                    ListItem.ListSubItems(1).ForeColor = vbBlue
                    ListItem.ListSubItems(2).ForeColor = vbBlue
                    ListItem.ListSubItems(3).ForeColor = vbBlue
                    ListItem.ListSubItems(4).ForeColor = vbBlue
                    ListItem.ListSubItems(5).ForeColor = vbBlue
                    ListItem.ListSubItems(6).ForeColor = vbBlue
                End If
                M_objrs.MoveNext
            Wend
        End With
        MsgBox "Data berhasil di load!", vbOKOnly + vbInformation, "Informasi"
    Else
        MsgBox "Data tidak ditemukan !", vbOKOnly + vbInformation, "Info"
    End If
    Set M_objrs = Nothing
End Sub

Private Sub Command2_Click()
    Frame1.Visible = True
End Sub

Private Sub Command3_Click()
    Frame1.Visible = False
End Sub

Private Sub NgisiDataAksesallPending()
    Dim sQuery As String
    Dim Randy_RS As ADODB.Recordset
    
    sQuery = "SELECT * FROM mandiri.temp_approval_aksesall"
    Set Randy_RS = New ADODB.Recordset
    Randy_RS.CursorLocation = adUseClient
    Randy_RS.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If Randy_RS.RecordCount > 0 Then
        TxtAgent.Text = IIf(IsNull(Randy_RS("agent")), "", Randy_RS("agent"))
        TxtTglAwal.Text = Format(IIf(IsNull(Randy_RS("tanggal_awal")), "", Randy_RS("tanggal_awal")), "YYYY-MM-DD")
        TxtWaktuAwal.Text = IIf(IsNull(Randy_RS("waktu_awal")), "", Randy_RS("waktu_awal"))
        TxtTglAkhir.Text = Format(IIf(IsNull(Randy_RS("tanggal_akhir")), "", Randy_RS("tanggal_akhir")), "YYYY-MM-DD")
        TxtWaktuAkhir.Text = IIf(IsNull(Randy_RS("waktu_akhir")), "", Randy_RS("waktu_akhir"))
    End If
    
     With FrmDistribusiAcc
        .LvAcc.ListItems.CLEAR
        .PB1.Max = Randy_RS.RecordCount
        While Not Randy_RS.EOF
            .PB1.Value = Randy_RS.Bookmark
            Set ListItem = .LvAcc.ListItems.ADD(, , Randy_RS("custid"))
            ListItem.SubItems(1) = Randy_RS("nama_customer")
            ListItem.SubItems(2) = IIf(IsNull(Randy_RS("status_account")), "", Randy_RS("status_account"))
            ListItem.SubItems(3) = IIf(IsNull(Randy_RS("agent_saat_ini")), "", Randy_RS("agent_saat_ini"))
            ListItem.SubItems(4) = IIf(IsNull(Randy_RS("agent_terdahulu")), "", Randy_RS("agent_terdahulu"))
                        
            If UCase(Randy_RS("agent")) = "AKSESALL" Then
                ListItem.ForeColor = vbRed
                ListItem.ListSubItems(1).ForeColor = vbRed
                ListItem.ListSubItems(2).ForeColor = vbRed
                ListItem.ListSubItems(3).ForeColor = vbRed
                ListItem.ListSubItems(4).ForeColor = vbRed
                ListItem.ListSubItems(5).ForeColor = vbRed
                ListItem.ListSubItems(6).ForeColor = vbRed
            End If
        
            If UCase(Randy_RS("agent")) = "#KOSONG#" Then
                ListItem.ForeColor = vbBlue
                ListItem.ListSubItems(1).ForeColor = vbBlue
                ListItem.ListSubItems(2).ForeColor = vbBlue
                ListItem.ListSubItems(3).ForeColor = vbBlue
                ListItem.ListSubItems(4).ForeColor = vbBlue
                ListItem.ListSubItems(5).ForeColor = vbBlue
                ListItem.ListSubItems(6).ForeColor = vbBlue
            End If
            
            Randy_RS.MoveNext
        Wend
    End With
End Sub

Private Sub Form_Load()
    Dim tglprofile As String
    Dim cmdsql, ran As String
    Dim M_objrs As ADODB.Recordset
    
    
    Call HeaderAgent
    Call HeaderAccount
    Call IsiComboFilter
    Call IsiComboAgent
    Call IsiComboStatusAcc
    
    cmdsql = "select now()"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_objrs.RecordCount > 0 Then
        LblWaktuServer.Caption = Format(M_objrs(0), "yyyy-mm-dd hh:nn")
        tglprofile = Format(M_objrs(0), "yyyymmdd")
    End If
    
    Call new_kdprofile
    Call isi_batch
    
    Set M_objrs = Nothing
    
    Call IsiAgentCollectBersama
    Call IsiStatusCollectBersama
    
    If CekAksesallPending = False Then
        cmdproses.Enabled = False
        Label25.Visible = False
    Else
        Label25.Visible = True
        
        ran = MsgBox("Ada Pendingan Aksesall, Apakah Mau Di-Approve?", vbYesNo + vbQuestion, "Konfirmasi")
        If ran = vbNo Then
            cmdproses.Enabled = False
            cmd_insert_antrian.Enabled = True
        Else
            cmdproses.Enabled = True
            cmd_insert_antrian.Enabled = False
            Call NgisiDataAksesallPending
        End If
    End If
    
    
    
End Sub

Private Function CekAksesallPending() As Boolean
    Dim sQuery As String
    Dim Randy_RS As ADODB.Recordset
    
    sQuery = "SELECT custid FROM mandiri.temp_approval_aksesall limit 1"
    Set Randy_RS = New ADODB.Recordset
    Randy_RS.CursorLocation = adUseClient
    Randy_RS.Open sQuery, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If Randy_RS.RecordCount > 0 Then
        CekAksesallPending = True
    Else
        CekAksesallPending = False
    End If
    
    Set Randy_RS = Nothing
End Function

Private Sub isi_batch()
    Dim M_objrs As ADODB.Recordset
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open "SELECT distinct RECSOURCE FROM mandiri.mgm WHERE RECSOURCE is not null ORDER BY RECSOURCE", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    list_batch.CLEAR
    If M_objrs.RecordCount > 0 Then
        'cb_batch.CLEAR
        'cb_batch.AddItem ""
        Do Until M_objrs.EOF
            'cb_batch.AddItem cnull(M_Objrs!RECSOURCE)
            list_batch.AddItem cnull(M_objrs!RECSOURCE)
            M_objrs.MoveNext
        Loop
    End If
    
    Set M_objrs = Nothing
End Sub

Private Sub LvAcc_Click()
    Call CariAgent
End Sub

Public Sub CariAgent()
    Dim cmdsql As String
    Dim M_objrs As ADODB.Recordset
    Dim m_objrs2 As ADODB.Recordset
    Dim ListItem As ListItem
    Dim z As Integer
    
    If LvAcc.ListItems.Count = 0 Then
        MsgBox "Maaf, data tidak tersedia!", vbOKOnly + vbInformation, "Peringatan"
        Exit Sub
    End If
    
    
    lblcustid.Caption = LvAcc.SelectedItem.Text
    LblNama.Caption = LvAcc.SelectedItem.SubItems(1)
    LblStatusAcc.Caption = IIf(IsNull(LvAcc.SelectedItem.SubItems(2)), "-", LvAcc.SelectedItem.SubItems(2))
    
'    Cmdsql = "select * from tbl_distribusi_account where custid='"
'    Cmdsql = Cmdsql & CStr(lblcustid.Caption) & "' order by agent asc "
'    Set M_Objrs = New ADODB.Recordset
'    M_Objrs.CursorLocation = adUseClient
'    M_Objrs.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    ' UPDATE 22 MEI 2013 BY IZUDDIN
    cmdsql = "select x.userid as agent,a.kd_profile,b.custid,a.waktu_awal,a.waktu_akhir,a.log_distribusi,a.log_tgl_distribusi from mandiri.usertbl x,mandiri.tbl_profile_aksesall a, mandiri.tbl_cust_aksesall b WHERE a.kd_profile=b.kd_profile AND a.kd_profile=x.profile_akses_all AND b.custid='"
    cmdsql = cmdsql & CStr(lblcustid.Caption) & "' ORDER BY userid"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    TxtJmlhAgent.Text = M_objrs.RecordCount
    LVAgent.ListItems.CLEAR
    
    If M_objrs.RecordCount = 0 Then
        Set M_objrs = Nothing
        Exit Sub
    End If
    
'    Set M_Objrs2 = New ADODB.Recordset
'    M_Objrs2.CursorLocation = adUseClient

    z = 0
    While Not M_objrs.EOF
'        If M_Objrs2.state = 1 Then M_Objrs2.Close
'        Cmdsql = "SELECT agent FROM tbl_cust_not_aksesall WHERE custid='" & M_Objrs("custid") & "' AND agent='" & M_Objrs("agent") & "' AND kd_profile='" & M_Objrs("kd_profile") & "'"
'        M_Objrs2.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        z = z + 1
        Set ListItem = LVAgent.ListItems.ADD(, , z)
            ListItem.SubItems(1) = M_objrs("agent")
            ListItem.SubItems(2) = M_objrs("custid")
            ListItem.SubItems(3) = Format(M_objrs("waktu_awal"), "yyyy-mm-dd hh:nn:ss")
            ListItem.SubItems(4) = Format(M_objrs("waktu_akhir"), "yyyy-mm-dd hh:nn:ss")
            ListItem.SubItems(5) = M_objrs("log_distribusi")
            ListItem.SubItems(6) = Format(M_objrs("log_tgl_distribusi"), "yyyy-mm-dd hh:nn:ss")
            ListItem.SubItems(7) = M_objrs("kd_profile")
       M_objrs.MoveNext
    Wend
    
    Set M_objrs = Nothing
End Sub

Private Sub IsiComboFilter()
    Dim cmdsql As String
    Dim M_objrs As ADODB.Recordset
    
    CmbFilterAcc.CLEAR
    
    CmbFilterAcc.AddItem "ALL"
    
    cmdsql = "select * from mandiri.usertbl where userid not in ('LUNAS','COMPLAIN','CLAIM') order by userid asc "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            CmbFilterAcc.AddItem M_objrs("userid")
            M_objrs.MoveNext
        Wend
    End If
    
    Set M_objrs = Nothing
End Sub

Private Sub CariFilter()
    Dim cmdsql As String
    Dim M_objrs As ADODB.Recordset
    Dim ListItem As ListItem
    Dim M_WHERE As String
    
    M_WHERE = ""
    
    cmdsql = "select * from mandiri.mgm  "
    
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
    
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open cmdsql & M_WHERE, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvAcc.ListItems.CLEAR
    TxtJmlhAcc.Text = M_objrs.RecordCount
    
    
    If M_objrs.RecordCount = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbExclamation, "Peringatan"
        Set M_objrs = Nothing
        Exit Sub
    End If
    
    PB1.Max = M_objrs.RecordCount
    While Not M_objrs.EOF
        PB1.Value = M_objrs.Bookmark
        Set ListItem = LvAcc.ListItems.ADD(, , M_objrs("custid"))
            ListItem.SubItems(1) = M_objrs("name")
            ListItem.SubItems(2) = IIf(IsNull(M_objrs("f_cek_new")), "", M_objrs("f_cek_new"))
            ListItem.SubItems(3) = IIf(IsNull(M_objrs("agent")), "", M_objrs("agent"))
            ListItem.SubItems(4) = IIf(IsNull(M_objrs("agent_asli")), "", M_objrs("agent_asli"))
            ListItem.SubItems(5) = IIf(IsNull(M_objrs("monitor_akses")), "", M_objrs("monitor_akses"))
            ListItem.SubItems(6) = IIf(IsNull(M_objrs("waktu_akses")), "", Format(M_objrs("waktu_akses"), "yyyy-mm-dd hh:nn:ss"))
            
            If UCase(M_objrs("agent")) = "AKSESALL" Then
                ListItem.ForeColor = vbRed
                ListItem.ListSubItems(1).ForeColor = vbRed
                ListItem.ListSubItems(2).ForeColor = vbRed
                ListItem.ListSubItems(3).ForeColor = vbRed
                ListItem.ListSubItems(4).ForeColor = vbRed
                ListItem.ListSubItems(5).ForeColor = vbRed
                ListItem.ListSubItems(6).ForeColor = vbRed
            End If
            
            If UCase(M_objrs("agent")) = "#KOSONG#" Then
                ListItem.ForeColor = vbBlue
                ListItem.ListSubItems(1).ForeColor = vbBlue
                ListItem.ListSubItems(2).ForeColor = vbBlue
                ListItem.ListSubItems(3).ForeColor = vbBlue
                ListItem.ListSubItems(4).ForeColor = vbBlue
                ListItem.ListSubItems(5).ForeColor = vbBlue
                ListItem.ListSubItems(6).ForeColor = vbBlue
            End If
            
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
End Sub

Private Sub LvAcc_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LvAcc.SortKey = ColumnHeader.Index - 1
    LvAcc.Sorted = True
End Sub

Private Sub LvAgent_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LVAgent.SortKey = ColumnHeader.Index - 1
    LVAgent.Sorted = True
End Sub

Private Sub LvAgent_DblClick()
    cmdEdit_Click
End Sub

Private Sub IsiComboAgent()
    Dim cmdsql As String
    Dim M_objrs As ADODB.Recordset
    
    CmbAgent.CLEAR
    CmbAgent.AddItem "ALL"
    
    cmdsql = "select * from mandiri.usertbl where usertype in ('1','6') and userid "
    cmdsql = cmdsql & " not in ('LUNAS','COMPLAIN','COMPLAIN','CLAIM') and userid is not null order by userid asc "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            CmbAgent.AddItem M_objrs("userid")
            M_objrs.MoveNext
        Wend
    End If
    
    Set M_objrs = Nothing
End Sub

'@@14022013 Tambahan filter status account
Private Sub IsiComboStatusAcc()
    Dim cmdsql As String
    Dim M_objrs As ADODB.Recordset
    
    CmbStatusAcc.CLEAR
    CmbStatusAcc.AddItem "ALL"
    CmbStatusAcc.AddItem "LPD 1"
    CmbStatusAcc.AddItem "LPD 2"
    CmbStatusAcc.AddItem "LPD 3"
    CmbStatusAcc.AddItem "LPD 3<"
    
    cmdsql = "select * from mandiri.contacteddesc where status='1' and jenis is not null "
    cmdsql = cmdsql & " and  jenis<>'CO-' order by jenis asc "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_objrs.RecordCount > 0 Then
        While Not M_objrs.EOF
            CmbStatusAcc.AddItem IIf(IsNull(M_objrs("jenis")), "", M_objrs("jenis"))
            M_objrs.MoveNext
        Wend
    End If
    CmbStatusAcc.AddItem "PTP-"
    Set M_objrs = Nothing
End Sub

'@@21022013 Tambahan buat program filter account yang bisa di collect bersama
Private Sub IsiAgentCollectBersama()
    Dim cmdsql As String
    Dim M_objrs As ADODB.Recordset
    Dim M_Objrs_TL As ADODB.Recordset
    
    CmbAgentCollBersama.CLEAR
    
    cmdsql = "select distinct agent_asli from mandiri.mgm "
    cmdsql = cmdsql + " where agent_asli is not null and agent='AKSESALL' "
    cmdsql = cmdsql + " and agent in (select userid from mandiri.usertbl where usertype='1') "
    cmdsql = cmdsql + " order by agent_asli asc "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, admcdtext
    If M_objrs.RecordCount > 0 Then
        CmbAgentCollBersama.AddItem "ALL"
        While Not M_objrs.EOF
            CmbAgentCollBersama.AddItem M_objrs("agent_asli")
            M_objrs.MoveNext
        Wend
    
        'Load TLnya juga buat grouping
        cmdsql = "select userid from mandiri.usertbl where usertype='6' order by userid asc "
        Set M_Objrs_TL = New ADODB.Recordset
        M_Objrs_TL.CursorLocation = adUseClient
        M_Objrs_TL.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_Objrs_TL.RecordCount > 0 Then
            While Not M_Objrs_TL.EOF
                CmbAgentCollBersama.AddItem M_Objrs_TL("userid")
                M_Objrs_TL.MoveNext
            Wend
        End If
        Set M_Objrs_TL = Nothing
    End If
    Set M_objrs = Nothing
    
End Sub


Private Sub IsiStatusCollectBersama()
    Dim cmdsql As String
    Dim M_objrs As ADODB.Recordset
    
    CmbStatusCollBersama.CLEAR
    
    cmdsql = "select distinct f_cek_new from mandiri.mgm where agent='AKSESALL' "
    cmdsql = cmdsql + " and f_cek_new is not null order by f_cek_new asc "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_objrs.RecordCount > 0 Then
        CmbStatusCollBersama.AddItem "ALL"
        CmbStatusCollBersama.AddItem "LPD 1"
        CmbStatusCollBersama.AddItem "LPD 2"
        CmbStatusCollBersama.AddItem "LPD 3"
        CmbStatusCollBersama.AddItem "LPD 3<"
        While Not M_objrs.EOF
            CmbStatusCollBersama.AddItem Trim(UCase(Mid(M_objrs("f_cek_new"), 1, 3)))
            M_objrs.MoveNext
        Wend
    End If
    Set M_objrs = Nothing
End Sub



