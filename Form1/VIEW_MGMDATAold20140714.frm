VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form VIEW_MGMDATA 
   BackColor       =   &H80000004&
   ClientHeight    =   10275
   ClientLeft      =   330
   ClientTop       =   2025
   ClientWidth     =   17820
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   Icon            =   "VIEW_MGMDATAold.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10275
   ScaleWidth      =   17820
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00B1FDD5&
      ForeColor       =   &H00000000&
      Height          =   7860
      Left            =   0
      TabIndex        =   50
      Top             =   2100
      Width           =   19185
      Begin VB.TextBox TxtJmlVolMgm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   12090
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   7095
         Width           =   3045
      End
      Begin VB.TextBox TxtJmlDtMgm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   15435
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   7080
         Width           =   3045
      End
      Begin MSComctlLib.ListView LstVwSearchMgm 
         Height          =   6900
         Left            =   0
         TabIndex        =   53
         Top             =   120
         Width           =   18960
         _ExtentX        =   33443
         _ExtentY        =   12171
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3555
      Left            =   555
      TabIndex        =   43
      Top             =   5310
      Visible         =   0   'False
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   6271
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Card Holder Data"
      TabPicture(0)   =   "VIEW_MGMDATAold.frx":1CFA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LblTarget(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Check1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Referall Data"
      TabPicture(1)   =   "VIEW_MGMDATAold.frx":1D16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LblTarget(1)"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).ControlCount=   2
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MGM Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   3090
         MaskColor       =   &H000000FF&
         TabIndex        =   49
         Top             =   1950
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   3225
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H80000004&
         Height          =   8685
         Left            =   -74940
         TabIndex        =   44
         Top             =   345
         Width           =   15075
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   300
            Left            =   11910
            Locked          =   -1  'True
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   7980
            Width           =   3045
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   7845
            Left            =   0
            TabIndex        =   46
            Top             =   120
            Width           =   15015
            _ExtentX        =   26485
            _ExtentY        =   13838
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   12582912
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            MousePointer    =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Label LblTarget 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Index           =   0
         Left            =   3315
         TabIndex        =   48
         Top             =   180
         Visible         =   0   'False
         Width           =   4605
      End
      Begin VB.Label LblTarget 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Index           =   1
         Left            =   -71820
         TabIndex        =   47
         Top             =   -15
         Visible         =   0   'False
         Width           =   9465
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   390
      Left            =   3675
      TabIndex        =   55
      Top             =   -15
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   688
      _Version        =   196610
      Font3D          =   4
      ForeColor       =   12582912
      BackColor       =   10147522
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Customer"
      BevelWidth      =   2
      BorderWidth     =   1
      BevelOuter      =   1
      BevelInner      =   2
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   2130
      Left            =   -15
      TabIndex        =   21
      Top             =   0
      Width           =   19125
      _ExtentX        =   33734
      _ExtentY        =   3757
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   706
      BackColor       =   11664853
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Find"
      TabPicture(0)   =   "VIEW_MGMDATAold.frx":1D32
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Schedulle"
      TabPicture(1)   =   "VIEW_MGMDATAold.frx":41F0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H00B1FDD5&
         ForeColor       =   &H80000008&
         Height          =   1425
         Left            =   -74985
         TabIndex        =   36
         Top             =   600
         Width           =   18930
         Begin VB.CheckBox Check2 
            Caption         =   "MGM Data"
            BeginProperty DataFormat 
               Type            =   4
               Format          =   "M/d/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   8
            EndProperty
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   4065
            TabIndex        =   16
            Top             =   1470
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   5
            Left            =   3360
            TabIndex        =   15
            Top             =   150
            Width           =   3480
         End
         Begin Threed.SSFrame SSFrame1 
            Height          =   510
            Left            =   1500
            TabIndex        =   37
            Top             =   540
            Width           =   3390
            _ExtentX        =   5980
            _ExtentY        =   900
            _Version        =   196610
            BackColor       =   -2147483644
            BackStyle       =   1
            Begin TDBDate6Ctl.TDBDate TDBDate1 
               Height          =   315
               Index           =   0
               Left            =   75
               TabIndex        =   17
               Top             =   60
               Width           =   1125
               _Version        =   65536
               _ExtentX        =   1984
               _ExtentY        =   556
               Calendar        =   "VIEW_MGMDATAold.frx":67BD
               Caption         =   "VIEW_MGMDATAold.frx":68D5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "VIEW_MGMDATAold.frx":6941
               Keys            =   "VIEW_MGMDATAold.frx":695F
               Spin            =   "VIEW_MGMDATAold.frx":69BD
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd-mmm-yyyy"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               FirstMonth      =   4
               ForeColor       =   -2147483640
               Format          =   "dd-mmm-yyyy"
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
               Text            =   "__-___-____"
               ValidateMode    =   0
               ValueVT         =   1
               Value           =   37609
               CenturyMode     =   0
            End
            Begin TDBDate6Ctl.TDBDate TDBDate1 
               Height          =   315
               Index           =   1
               Left            =   1830
               TabIndex        =   18
               Top             =   90
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               Calendar        =   "VIEW_MGMDATAold.frx":69E5
               Caption         =   "VIEW_MGMDATAold.frx":6AFD
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "VIEW_MGMDATAold.frx":6B69
               Keys            =   "VIEW_MGMDATAold.frx":6B87
               Spin            =   "VIEW_MGMDATAold.frx":6BE5
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd-mmm-yyyy"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               FirstMonth      =   4
               ForeColor       =   -2147483640
               Format          =   "dd-mmm-yyyy"
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
               Text            =   "__-___-____"
               ValidateMode    =   0
               ValueVT         =   1
               Value           =   37609
               CenturyMode     =   0
            End
            Begin VB.Label Label1 
               BackColor       =   &H80000004&
               BackStyle       =   0  'Transparent
               Caption         =   "S/d"
               ForeColor       =   &H00000000&
               Height          =   300
               Index           =   7
               Left            =   1365
               TabIndex        =   38
               Top             =   150
               Width           =   285
            End
         End
         Begin Threed.SSCommand CmdScheduleoK 
            Height          =   690
            Left            =   5160
            TabIndex        =   19
            Top             =   540
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   1217
            _Version        =   196610
            Font3D          =   5
            MousePointer    =   16
            BackColor       =   -2147483644
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Search"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand CmdToday 
            Height          =   480
            Left            =   120
            TabIndex        =   41
            Top             =   1920
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   847
            _Version        =   196610
            Font3D          =   5
            MousePointer    =   16
            BackColor       =   -2147483644
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Today"
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand CmdMissed 
            Height          =   450
            Left            =   150
            TabIndex        =   42
            Top             =   2475
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   794
            _Version        =   196610
            Font3D          =   5
            MousePointer    =   16
            BackColor       =   -2147483644
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Miss"
            ButtonStyle     =   2
         End
         Begin VB.ComboBox Combo1 
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   4
            Left            =   1485
            TabIndex        =   39
            Top             =   165
            Width           =   1905
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Tgl Schedule :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Index           =   12
            Left            =   240
            TabIndex        =   54
            Top             =   600
            Width           =   1170
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Agent :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   13
            Left            =   180
            TabIndex        =   40
            Top             =   195
            Width           =   1170
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00B1FDD5&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1725
         Left            =   0
         TabIndex        =   22
         Top             =   480
         Width           =   19035
         Begin VB.CommandButton Cmd_listrequestdecease 
            BackColor       =   &H0080FFFF&
            Caption         =   "&List Request Acc Decease"
            Height          =   375
            Left            =   5040
            Style           =   1  'Graphical
            TabIndex        =   77
            Top             =   1200
            Width           =   2175
         End
         Begin VB.CommandButton cmd_claimback_acc 
            Caption         =   "Batal claim acc"
            Height          =   375
            Left            =   16440
            TabIndex        =   75
            Top             =   1200
            Width           =   2415
         End
         Begin VB.CommandButton cmd_review 
            Caption         =   "List REVIEW"
            Height          =   375
            Left            =   16440
            TabIndex        =   74
            Top             =   800
            Width           =   2415
         End
         Begin VB.CommandButton CmdListHotProsPect 
            Caption         =   "&List Hot Prospect"
            Height          =   375
            Left            =   7380
            TabIndex        =   73
            Top             =   1200
            Width           =   1575
         End
         Begin VB.CommandButton CmdReleaseUnRealesePTP 
            Caption         =   "Realese/ UnRealese PTP"
            Height          =   615
            Left            =   17460
            TabIndex        =   72
            Top             =   180
            Width           =   1455
         End
         Begin VB.CommandButton CmdSearchPTP 
            Caption         =   "Search Tgl.Tagih"
            Height          =   615
            Left            =   16440
            TabIndex        =   71
            Top             =   180
            Width           =   975
         End
         Begin VB.TextBox txtcurbalance 
            Height          =   285
            Left            =   14130
            Locked          =   -1  'True
            TabIndex        =   69
            Top             =   1140
            Width           =   2085
         End
         Begin VB.TextBox txtamount 
            Height          =   315
            Left            =   14130
            Locked          =   -1  'True
            TabIndex        =   67
            Top             =   450
            Width           =   2085
         End
         Begin VB.Timer Timer1 
            Left            =   12600
            Top             =   960
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00B1FDD5&
            Caption         =   "Status CPA"
            Height          =   600
            Left            =   9135
            TabIndex        =   63
            Top             =   855
            Width           =   2445
            Begin VB.ComboBox Combo3 
               Height          =   315
               ItemData        =   "VIEW_MGMDATAold.frx":6C0D
               Left            =   45
               List            =   "VIEW_MGMDATAold.frx":6C17
               Style           =   2  'Dropdown List
               TabIndex        =   64
               Top             =   180
               Width           =   2310
            End
         End
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "VIEW_MGMDATAold.frx":6C26
            Left            =   4995
            List            =   "VIEW_MGMDATAold.frx":6C28
            TabIndex        =   56
            Top             =   735
            Width           =   3975
         End
         Begin VB.ComboBox cmbStsLastCall 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   0
            Left            =   4200
            TabIndex        =   7
            Top             =   1620
            Visible         =   0   'False
            Width           =   3180
         End
         Begin VB.ComboBox CmbStatusCek 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   8745
            TabIndex        =   14
            Top             =   2265
            Width           =   1800
         End
         Begin VB.CheckBox CekDtDistribute 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Searching Data Belum Distribute"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   150
            MaskColor       =   &H000000FF&
            TabIndex        =   20
            Top             =   -345
            Visible         =   0   'False
            Width           =   3225
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   2
            Left            =   1200
            MaxLength       =   20
            TabIndex        =   0
            Top             =   120
            Width           =   1965
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Index           =   3
            Left            =   11460
            TabIndex        =   13
            Top             =   435
            Width           =   2580
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Index           =   2
            Left            =   9630
            TabIndex        =   12
            Top             =   435
            Width           =   1800
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Index           =   1
            Left            =   6195
            TabIndex        =   6
            Top             =   435
            Width           =   2790
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Index           =   0
            Left            =   4995
            TabIndex        =   5
            Top             =   435
            Width           =   1200
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   0
            Left            =   4995
            TabIndex        =   4
            Top             =   120
            Width           =   2025
         End
         Begin TDBDate6Ctl.TDBDate TdDob 
            Height          =   315
            Left            =   1200
            TabIndex        =   1
            Top             =   1650
            Visible         =   0   'False
            Width           =   1170
            _Version        =   65536
            _ExtentX        =   2064
            _ExtentY        =   556
            Calendar        =   "VIEW_MGMDATAold.frx":6C2A
            Caption         =   "VIEW_MGMDATAold.frx":6D42
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "VIEW_MGMDATAold.frx":6DAE
            Keys            =   "VIEW_MGMDATAold.frx":6DCC
            Spin            =   "VIEW_MGMDATAold.frx":6E2A
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   16777215
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "dd-mmm-yyyy"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            FirstMonth      =   4
            ForeColor       =   0
            Format          =   "dd-mmm-yyyy"
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
            Text            =   "__-___-____"
            ValidateMode    =   0
            ValueVT         =   1
            Value           =   37475
            CenturyMode     =   0
         End
         Begin TDBMask6Ctl.TDBMask TDBMask1 
            Height          =   315
            Left            =   1200
            TabIndex        =   2
            Top             =   435
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   556
            Caption         =   "VIEW_MGMDATAold.frx":6E52
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "VIEW_MGMDATAold.frx":6EBE
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   16777215
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   0
            Format          =   "999-999999999999"
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
            Text            =   "___-____________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask TDBMask2 
            Height          =   315
            Left            =   1320
            TabIndex        =   3
            Top             =   420
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   556
            Caption         =   "VIEW_MGMDATAold.frx":6F00
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "VIEW_MGMDATAold.frx":6F6C
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   16777215
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   0
            Format          =   "9999-999999999999999999"
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
            Text            =   "____-__________________"
            Value           =   ""
         End
         Begin TDBTime6Ctl.TDBTime DTimeLastCall 
            Height          =   300
            Index           =   0
            Left            =   6300
            TabIndex        =   9
            Top             =   780
            Visible         =   0   'False
            Width           =   825
            _Version        =   65536
            _ExtentX        =   1455
            _ExtentY        =   529
            Caption         =   "VIEW_MGMDATAold.frx":6FAE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "VIEW_MGMDATAold.frx":701A
            Spin            =   "VIEW_MGMDATAold.frx":706A
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   16777215
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
            Text            =   "00:00"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   6.13425925925926E-04
         End
         Begin TDBDate6Ctl.TDBDate DtLastCall 
            Height          =   315
            Index           =   0
            Left            =   5520
            TabIndex        =   8
            Top             =   780
            Visible         =   0   'False
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   556
            Calendar        =   "VIEW_MGMDATAold.frx":7092
            Caption         =   "VIEW_MGMDATAold.frx":71AA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "VIEW_MGMDATAold.frx":7216
            Keys            =   "VIEW_MGMDATAold.frx":7234
            Spin            =   "VIEW_MGMDATAold.frx":7292
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   16777215
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "dd-mmm-yyyy"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            FirstMonth      =   4
            ForeColor       =   0
            Format          =   "dd-mmm-yyyy"
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
            Text            =   "__-___-____"
            ValidateMode    =   0
            ValueVT         =   1
            Value           =   37468
            CenturyMode     =   0
         End
         Begin TDBDate6Ctl.TDBDate DtLastCall 
            Height          =   315
            Index           =   1
            Left            =   7020
            TabIndex        =   10
            Top             =   720
            Visible         =   0   'False
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   556
            Calendar        =   "VIEW_MGMDATAold.frx":72BA
            Caption         =   "VIEW_MGMDATAold.frx":73D2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "VIEW_MGMDATAold.frx":743E
            Keys            =   "VIEW_MGMDATAold.frx":745C
            Spin            =   "VIEW_MGMDATAold.frx":74BA
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   16777215
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "dd-mmm-yyyy"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            FirstMonth      =   4
            ForeColor       =   0
            Format          =   "dd-mmm-yyyy"
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
            Text            =   "__-___-____"
            ValidateMode    =   0
            ValueVT         =   1
            Value           =   37468
            CenturyMode     =   0
         End
         Begin TDBTime6Ctl.TDBTime DTimeLastCall 
            Height          =   300
            Index           =   1
            Left            =   7800
            TabIndex        =   11
            Top             =   780
            Visible         =   0   'False
            Width           =   825
            _Version        =   65536
            _ExtentX        =   1455
            _ExtentY        =   529
            Caption         =   "VIEW_MGMDATAold.frx":74E2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "VIEW_MGMDATAold.frx":754E
            Spin            =   "VIEW_MGMDATAold.frx":759E
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   16777215
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
            Text            =   "20:53"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   0.870289351851852
         End
         Begin VB.Frame Frame3 
            Caption         =   "Proses....!!"
            Height          =   615
            Left            =   9045
            TabIndex        =   23
            Top             =   1740
            Visible         =   0   'False
            Width           =   2025
            Begin MSComctlLib.ProgressBar ProgressBar1 
               Height          =   270
               Left            =   15
               TabIndex        =   24
               Top             =   165
               Width           =   1980
               _ExtentX        =   3493
               _ExtentY        =   476
               _Version        =   393216
               Appearance      =   0
            End
         End
         Begin Threed.SSCommand Command1 
            Height          =   690
            Index           =   0
            Left            =   1140
            TabIndex        =   59
            Top             =   780
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   1217
            _Version        =   196610
            Font3D          =   5
            MousePointer    =   16
            BackColor       =   -2147483644
            PictureFrames   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "VIEW_MGMDATAold.frx":75C6
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand Command1 
            Height          =   690
            Index           =   2
            Left            =   1980
            TabIndex        =   60
            Top             =   780
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   1217
            _Version        =   196610
            Font3D          =   5
            MousePointer    =   16
            BackColor       =   -2147483644
            PictureFrames   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "VIEW_MGMDATAold.frx":7A8D
            ButtonStyle     =   2
         End
         Begin Threed.SSCommand Command1 
            Height          =   690
            Index           =   1
            Left            =   2820
            TabIndex        =   70
            Top             =   780
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   1217
            _Version        =   196610
            Font3D          =   5
            MousePointer    =   16
            BackColor       =   -2147483644
            PictureFrames   =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "VIEW_MGMDATAold.frx":7F07
            AutoSize        =   1
            ButtonStyle     =   2
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Curr Balance"
            Height          =   225
            Left            =   14160
            TabIndex        =   68
            Top             =   810
            Width           =   1935
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            Height          =   225
            Left            =   14160
            TabIndex        =   66
            Top             =   150
            Width           =   735
         End
         Begin VB.Label LBLCOUNT 
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   12300
            TabIndex        =   65
            Top             =   990
            Width           =   1815
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00B1FDD5&
            Caption         =   "C&lose"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   3000
            TabIndex        =   62
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00B1FDD5&
            Caption         =   "&Clear"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   2100
            TabIndex        =   61
            Top             =   1455
            Width           =   495
         End
         Begin VB.Label Label2 
            BackColor       =   &H00B1FDD5&
            Caption         =   "&Search"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   1230
            TabIndex        =   58
            Top             =   1455
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Status Account"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   14
            Left            =   7440
            TabIndex        =   57
            Top             =   1740
            Width           =   1365
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Status Telp."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   8
            Left            =   3630
            TabIndex        =   35
            Top             =   810
            Width           =   1365
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tgl Terakhir Telp."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   9
            Left            =   5040
            TabIndex        =   34
            Top             =   840
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Status Check"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   10
            Left            =   7320
            TabIndex        =   33
            Top             =   2310
            Width           =   1365
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   11
            Left            =   6720
            TabIndex        =   32
            Top             =   780
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "# Kartu :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   5
            Left            =   0
            TabIndex        =   31
            Top             =   135
            Width           =   1125
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "HP :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   6
            Left            =   2310
            TabIndex        =   30
            Top             =   210
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Telp Rumah :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   4
            Left            =   60
            TabIndex        =   29
            Top             =   440
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Batch"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   9030
            TabIndex        =   28
            Top             =   435
            Width           =   570
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tgl Lahir :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   2
            Left            =   60
            TabIndex        =   27
            Top             =   1650
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TeleCollection"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   1
            Left            =   3630
            TabIndex        =   26
            Top             =   495
            Width           =   1365
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nama"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   0
            Left            =   3540
            TabIndex        =   25
            Top             =   150
            Width           =   1395
         End
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Time "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   17280
         TabIndex        =   76
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Menu MnFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnclaim 
         Caption         =   "Claim"
      End
   End
End
Attribute VB_Name = "VIEW_MGMDATA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_cari2 As ADODB.Recordset
Dim m_cari As ADODB.Recordset
    Dim mrs_cek As ADODB.Recordset
Dim StsVl  As String
Dim StsPR As String
Dim StsSK As String
Dim StsON As String
Dim StsOS As String
Dim StsPTP As String
Dim StsBP As String
Dim StsPOP As String
Dim StsSP As String
Dim StsUC As String
Dim StsAll As String
Dim Stsblank As String
Dim Stsf_fresh As String
Dim StsRP As String
Dim StsWO_Date As String
Dim StsWO_2009 As String
Dim StsWO_2008 As String
Dim StsWO_2007 As String
Dim StsWO_2006 As String
Dim StsWO_2005 As String
Dim StsWO_2004 As String
Dim StsWO_2003 As String
Dim StsWO_2002 As String
Dim StsWO_2001 As String
Dim StsWO_2000 As String
Dim StsWO_1999 As String
Dim StsWO_2010 As String
Dim cmdsql As String
Dim Bloked As String
Dim LUserType As String
Dim F_CEK As String
Dim WO_DATE As String
Dim f_Pending As String
Dim datajml As Integer
'@@ 14072010 Blok entry list
Dim BlokedEntry As String


Private Sub HEADER_VIEW_Refferall()
    ListView1.ColumnHeaders.ADD 1, , "No", 3 * TXT
    ListView1.ColumnHeaders.ADD 2, , "Cust Id", 5 * TXT
    ListView1.ColumnHeaders.ADD 3, , "Priority", 1
    ListView1.ColumnHeaders.ADD 4, , "Ref Id", 10 * TXT
    ListView1.ColumnHeaders.ADD 5, , "Ref Name", 10 * TXT
    ListView1.ColumnHeaders.ADD 6, , "Nama Customer", 25 * TXT
    'ListView1.ColumnHeaders.ADD 7, , "Batch Expire Date", 25 * TXT
    ListView1.ColumnHeaders.ADD 7, , "Tgl Schedule", 10 * TXT
    ListView1.ColumnHeaders.ADD 8, , "Next Action", 12 * TXT
    ListView1.ColumnHeaders.ADD 9, , "Remarks", 17 * TXT
    ListView1.ColumnHeaders.ADD 10, , "Sts LastCall", 17 * TXT
    ListView1.ColumnHeaders.ADD 11, , "SalesCode", 8 * TXT
    ListView1.ColumnHeaders.ADD 12, , "Agent", 8 * TXT
    ListView1.ColumnHeaders.ADD 13, , "DataBase", 10 * TXT
    ListView1.ColumnHeaders.ADD 14, , "LastCall Date", 10 * TXT
    ListView1.ColumnHeaders.ADD 15, , "Code", 5 * TXT
    ListView1.ColumnHeaders.ADD 16, , "Complaint Note", 15 * TXT
    ListView1.ColumnHeaders.ADD 17, , "Check", 10 * TXT
    ListView1.ColumnHeaders.ADD 18, , "ID", 10 * TXT
End Sub
Private Sub isi_dataClaimKeGrid(gCUSTID As String, gNama As String, gnextact As String, gremarks As String, gagent As String, gnamaagent As String, grecsource As String)
    ' insert ke grid list view
Dim ListItem As ListItem
Set ListItem = LstVwSearchMgm.ListItems.ADD(, , "9999")
    ListItem.SubItems(1) = gCUSTID
    ListItem.SubItems(2) = ""
    ListItem.SubItems(3) = gNama
    ListItem.SubItems(4) = Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd hh:nn")
    ListItem.SubItems(5) = gnextact
    ListItem.SubItems(6) = gremarks
    ListItem.SubItems(7) = gagent
    ListItem.SubItems(8) = gnamaagent
    ListItem.SubItems(9) = grecsource
    ListItem.SubItems(10) = ""
    ListItem.SubItems(11) = "1A"
    ListItem.SubItems(12) = ""
    ListItem.SubItems(13) = ""
    ListItem.SubItems(14) = ""
End Sub

Private Sub cmd_claimback_acc_Click()
    Form_return_claim.Show 1
End Sub

Private Sub Cmd_listrequestdecease_Click()
    Form_listreq_decease.Show 1
End Sub

Private Sub cmd_review_Click()
    FrmCustIdReview.Show
End Sub

Private Sub CmdListHotProsPect_Click()
    FrmListHotProspect.Show vbModal
End Sub

Private Sub CmdReleaseUnRealesePTP_Click()
    If UCase(MDIForm1.Text2.Text) = "AGENT" Then
        MsgBox "Anda tidak mendapatkan akses!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    FrmReleaseUnRealesePTP.Show vbModal
End Sub

Private Sub CmdScheduleoK_Click()
If Combo1(4).Text = Empty Then
    MsgBox "Agent Harus Diisi", vbCritical + vbOKOnly, "Informasi"
    Exit Sub
End If
If TDBDate1(0).ValueIsNull Or TDBDate1(1).ValueIsNull Then
    MsgBox "Tanggal Tidak Boleh Kosong", vbInformation + vbOKOnly, "Informasi"
    Exit Sub
End If
If TDBDate1(0).Value > TDBDate1(1).Value Then
    MsgBox "Tanggal Periode Awal harus Lebih Kecil Dari Tanggal Periode Akhir", vbInformation + vbOKOnly, "Informasi"
    Exit Sub
End If
Call cari_Schedule
End Sub

Private Sub cari_Schedule()
Dim M_DATA As New CLS_FRMSEARCH
Dim ListItem As ListItem
Dim VOLUMEAMOUNT As Long
'@@ 19 Juli 2010 tambahan u/ blok data
Dim Blokedsearch As String
Dim BlokedEntrysearch As String
Dim Strsql, StrsqlBlok, strinject As String
Dim M_Objrs As ADODB.Recordset
Dim blokeddatamarkup As String
Dim STSLOCKTL As String
Dim STSfromaccount As String
Dim NMAGETPREV As String
If Check2.Value = 1 Then

    LstVwSearchMgm.ListItems.CLEAR
    SSTab1.Tab = 0
    ' searching schedule mgm
  Call CEK_STATUS_F_CEK
  
  '--------- @@Start 19 Juli 2010 tambahan bloked
   Strsql = "select * from usertbl where userid='"
   Strsql = Strsql + Trim(MDIForm1.Text1.Text) + "'"
   Set M_Objrs = New ADODB.Recordset
   M_Objrs.CursorLocation = adUseClient
   M_Objrs.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
   If M_Objrs.RecordCount <> 0 Then
   
     STSLOCKTL = CStr(Trim(IIf(IsNull(M_Objrs!lockdarispvbuattl), "", M_Objrs!lockdarispvbuattl)))
                
                STSfromaccount = CStr(Trim(IIf(IsNull(M_Objrs!fromaccount), "", M_Objrs!fromaccount)))
                
        If UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
                        NMAGETPREV = STSLOCKTL
         End If
         
        
        If M_Objrs("usertype") = "1" Then
            strinject = IIf(IsNull(M_Objrs!lockdarispv), "", M_Objrs!lockdarispv)
           If strinject = "" Then
              Blokedsearch = ""
           Else
             Blokedsearch = IIf(IsNull(M_Objrs!lockdarispv), "", Replace(M_Objrs!lockdarispv, "@", "'"))
           End If
           BlokedEntrysearch = ""
           BlokedEntrysearch = IIf(IsNull(M_Objrs!lock_entry_lpd), "", M_Objrs!lock_entry_lpd)
           blokeddatamarkup = IIf(IsNull(M_Objrs!lockmarkup), "", M_Objrs!lockmarkup)
           If blokeddatamarkup <> "" Then
                    F_CEK = ""
                    Blokedsearch = ""

                     BlokedEntrysearch = ""
                End If
        End If
   End If
   
   If StsWO_Date = "1" Then
            If LUserType = "1" Then
                WO_DATE = "substring(B_D,1,4) in ('" + StsWO_2009 + "','" + StsWO_2008 + "','" + StsWO_2007 + "','" + StsWO_2006 + "','" + StsWO_2005 + "', "
                WO_DATE = WO_DATE + "'" + StsWO_2004 + "', '" + StsWO_2003 + "', '" + StsWO_2002 + "', '" + StsWO_2001 + "','" + StsWO_2000 + "','" + StsWO_1999 + "','" + StsWO_2010 + "')"
            End If
      End If
      
  
' If STSLOCKTL <> Empty Then
'        If Left(Combo1(0).Text, 5) = "LUNAS" Then
'                If STSfromaccount = "LUNAS PENDING" Then
'                    STSLOCKTL = STSLOCKTL
'                ElseIf STSfromaccount = "LUNAS COMPLETE" Then
'                      STSLOCKTL = STSLOCKTL
'                Else
'                     STSLOCKTL = ""
'                End If
'
'        Else
'                STSLOCKTL = ""
'        End If
'        End If
        
  '--------- @@End 19 Juli 2010 Tambahan bloked
  
   Set m_cari = M_DATA.QUERY_SEARCH_mgm(M_OBJCONN, "AGENT = '" + Combo1(4).Text + "' AND (NEXTACTDATE BETWEEN '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " 00:00" + "' AND '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " 23:59" + "') ", MDIForm1.Text3.Text, F_CEK, f_Pending, Blokedsearch, BlokedEntrysearch, blokeddatamarkup, WO_DATE, NMAGETPREV)
    ProgressBar1.Max = m_cari.RecordCount + 1
    If Check2.Value = 1 Then
        TxtJmlDtMgm.Text = m_cari.RecordCount & " Data"
    Else
        Text2.Text = m_cari.RecordCount & " Data"
    End If
    
    While Not m_cari.EOF
    ProgressBar1.Value = m_cari.Bookmark
    
    Set ListItem = LstVwSearchMgm.ListItems.ADD(, , m_cari.Bookmark)
    
        ListItem.SubItems(1) = IIf(IsNull(m_cari("CUSTID")), "", m_cari("CUSTID"))
        ListItem.SubItems(2) = IIf(IsNull(m_cari("PRIOR")), "", m_cari("PRIOR"))
        ListItem.SubItems(3) = IIf(IsNull(m_cari("NAME")), "", m_cari("NAME"))
        'listitem.SubItems(4) = IIf(IsNull(m_cari("TGLSOURCE")), "", DateAdd("d", 90, Format(m_cari("TGLSOURCE"), "dd-mm-yyyy")))
        ListItem.SubItems(4) = IIf(IsNull(m_cari("recsource")), "", m_cari("recsource"))
        ListItem.SubItems(5) = IIf(IsNull(m_cari("NEXTACTDATE")), "", Format(m_cari("NEXTACTDATE"), "yyyy/mm/dd hh:nn"))
        ListItem.SubItems(6) = IIf(IsNull(m_cari("NEXTACT")), "", m_cari("NEXTACT"))
        ListItem.SubItems(7) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS")) & "-" & IIf(IsNull(m_cari("f_pending")), "", m_cari("f_pending"))
        ListItem.SubItems(8) = IIf(IsNull(m_cari("KETHSLKERJA_NEW")), "", m_cari("KETHSLKERJA_NEW"))
        ListItem.SubItems(9) = IIf(IsNull(m_cari("StatusCall")), "", m_cari("StatusCall"))
        ListItem.SubItems(11) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
        'listitem.SubItems(10) = IIf(IsNull(m_cari("NamaAGENT")), "", m_cari("NamaAGENT"))
        ListItem.SubItems(12) = Format(IIf(IsNull(m_cari("Principal")), 0, m_cari("Principal")), "##,###")
        ListItem.SubItems(13) = Format(IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo")), "##,###")

       VOLUMEAMOUNT = VOLUMEAMOUNT + IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo"))
'        'LISTITEM.SubItems(15) = IIf(IsNull(m_cari("[NO]")), "", m_cari("[NO]"))
'
'        '@@16032011 Tambahan DOB dan No KTP
'        listitem.SubItems(26) = IIf(IsNull(m_cari("dob")), "", Format(m_cari("dob"), "yyyy-mm-dd"))
'        listitem.SubItems(27) = IIf(IsNull(m_cari("ktpno")), "", m_cari("ktpno"))

         ListItem.SubItems(14) = Format(IIf(IsNull(m_cari("OpenDate")), "", m_cari("OpenDate")), "yyyy/mm/dd")
        ListItem.SubItems(15) = Format(IIf(IsNull(m_cari("B_D")), 0, m_cari("B_D")))
        ListItem.SubItems(16) = Format(IIf(IsNull(m_cari("Pay_Dt")), 0, m_cari("Pay_Dt")), "yyyy/mm/dd")
         ListItem.SubItems(17) = Format(IIf(IsNull(m_cari("lastpay")), 0, m_cari("lastpay")), "##,###")
        
        ListItem.SubItems(18) = IIf(IsNull(m_cari("TGLSTATUS")), "", Format(m_cari("TGLSTATUS"), "YYYY/MM/DD"))
        ListItem.SubItems(19) = IIf(IsNull(m_cari("TGLCALL")), "", Format(m_cari("TGLCALL"), "YYYY/MM/DD"))
        ListItem.SubItems(20) = IIf(IsNull(m_cari("Kethslkerja")), "", m_cari("Kethslkerja"))
        ListItem.SubItems(21) = Format(IIf(IsNull(m_cari("TGLINCOMING")), "", m_cari("TGLINCOMING")), "YYYY/MM/DD")
        ListItem.SubItems(23) = IIf(IsNull(m_cari("resultcpa")), "", m_cari("resultcpa"))
        ListItem.SubItems(24) = IIf(IsNull(m_cari("tglinsertfrmcpa")), "", m_cari("tglinsertfrmcpa"))
        ListItem.SubItems(25) = Format(IIf(IsNull(m_cari("curbal")), "", m_cari("curbal")), "##,###")
        'TOTCURBALANCE = TOTCURBALANCE + IIf(IsNull(m_cari("curbal")), 0, m_cari("curbal"))
       
        '@@16032011 Tambahan DOB dan No KTP
        ListItem.SubItems(26) = IIf(IsNull(m_cari("dob")), "", Format(m_cari("dob"), "yyyy-mm-dd"))
        ListItem.SubItems(27) = IIf(IsNull(m_cari("ktpno")), "", m_cari("ktpno"))


'        Set listitem = LstVwSearchMgm.ListItems.ADD(, , m_cari.Bookmark)
'
'        If MDIForm1.Text2 = "TeamLeader" Then
'            If IIf(IsNull(m_cari("stscpa")), "0", m_cari("stscpa")) = 1 Then
'                listitem.ForeColor = vbRed
'            End If
'
'            If IIf(IsNull(m_cari("intapprovel")), "0", m_cari("intapprovel")) = 1 Then
'              listitem.ForeColor = vbBlue
'            End If
'
'        End If
'
'        If UCase(MDIForm1.Text7) = "JOKO" Or UCase(MDIForm1.Text7) = "WULANDARI" Or UCase(MDIForm1.Text7) = "ANDRI" Then
'            If IIf(IsNull(m_cari("intverify")), "0", m_cari("intverify")) = 1 Then
'                listitem.ForeColor = vbYellow
'            End If
'
'            If IIf(IsNull(m_cari("intapprovel")), "0", m_cari("intapprovel")) = 1 Then
'              listitem.ForeColor = vbGreen
'            End If
'        End If
'
'
'        'statusprior = IIf(IsNull(m_cari("StatusPrior")), "", m_cari("StatusPrior"))
'        listitem.SubItems(1) = IIf(IsNull(m_cari("CUSTID")), "", m_cari("CUSTID"))
'        listitem.SubItems(2) = IIf(IsNull(m_cari("PRIOR")), "", m_cari("PRIOR"))
'        listitem.SubItems(3) = IIf(IsNull(m_cari("NAME")), "", m_cari("NAME"))
'        listitem.SubItems(4) = IIf(IsNull(m_cari("RECSOURCE")), "", m_cari("RECSOURCE"))
'        listitem.SubItems(5) = IIf(IsNull(m_cari("NEXTACTDATE")), "", Format(m_cari("NEXTACTDATE"), "dd/mm/yyyy hh:nn"))
'        listitem.SubItems(6) = IIf(IsNull(m_cari("NEXTACT")), "", m_cari("NEXTACT"))
'        listitem.SubItems(7) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS"))
'        listitem.SubItems(8) = CStr(IIf(IsNull(m_cari("kethslkerja_new")), "", m_cari("kethslkerja_new")))
'        listitem.SubItems(9) = IIf(IsNull(m_cari("StatusCall")), "", m_cari("StatusCall"))
'        listitem.SubItems(11) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
'
'
'        If UCase(MDIForm1.Text2) <> "SUPERVISOR" Then
'                If Format(IIf(IsNull(m_cari("flaglead")), 0, m_cari("flaglead")), "##,###") = 1 Then
'                       listitem.SubItems(12) = ""
'                Else
'                    listitem.SubItems(12) = Format(IIf(IsNull(m_cari("Principal")), 0, m_cari("Principal")), "##,###")
'                End If
'        Else
'             listitem.SubItems(12) = Format(IIf(IsNull(m_cari("Principal")), 0, m_cari("Principal")), "##,###")
'        End If
'
'        listitem.SubItems(13) = Format(IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo")), "##,###")
'        'TOTAMOUNT = TOTAMOUNT + IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo"))
'
'
'        listitem.SubItems(14) = Format(IIf(IsNull(m_cari("OpenDate")), "", m_cari("OpenDate")), "yyyy/mm/dd")
'        listitem.SubItems(15) = Format(IIf(IsNull(m_cari("B_D")), 0, m_cari("B_D")))
'        listitem.SubItems(16) = Format(IIf(IsNull(m_cari("Pay_Dt")), 0, m_cari("Pay_Dt")), "yyyy/mm/dd")
'        listitem.SubItems(17) = IIf(IsNull(m_cari("RECSOURCE")), "", m_cari("RECSOURCE"))
'        listitem.SubItems(18) = IIf(IsNull(m_cari("TGLSTATUS")), "", Format(m_cari("TGLSTATUS"), "YYYY/MM/DD"))
'        listitem.SubItems(19) = IIf(IsNull(m_cari("TGLCALL")), "", Format(m_cari("TGLCALL"), "YYYY/MM/DD"))
'        listitem.SubItems(20) = IIf(IsNull(m_cari("Kethslkerja")), "", m_cari("Kethslkerja"))
'        listitem.SubItems(21) = Format(IIf(IsNull(m_cari("TGLINCOMING")), "", m_cari("TGLINCOMING")), "YYYY/MM/DD")
'        listitem.SubItems(23) = IIf(IsNull(m_cari("resultcpa")), "", m_cari("resultcpa"))
'        listitem.SubItems(24) = IIf(IsNull(m_cari("tglinsertfrmcpa")), "", m_cari("tglinsertfrmcpa"))
'        listitem.SubItems(25) = Format(IIf(IsNull(m_cari("curbal")), "", m_cari("curbal")), "##,###")
'        'TOTCURBALANCE = TOTCURBALANCE + IIf(IsNull(m_cari("curbal")), 0, m_cari("curbal"))
'
'        '@@16032011 Tambahan DOB dan No KTP
'        listitem.SubItems(26) = IIf(IsNull(m_cari("dob")), "", Format(m_cari("dob"), "yyyy-mm-dd"))
'        listitem.SubItems(27) = IIf(IsNull(m_cari("ktpno")), "", m_cari("ktpno"))

        m_cari.MoveNext
    Wend
        If LstVwSearchMgm.ListItems.Count = 0 Then
            TxtJmlDtMgm.Text = "Tidak Ada Data"
            TxtJmlVolMgm.Text = "0"
        Else
            TxtJmlDtMgm.Text = "Total " + CStr(m_cari.RecordCount) + " Records"
            'TxtJmlVolMgm.Text = "Total " + CStr(m_cari.RecordCount)
            TxtJmlVolMgm.Text = Format(VOLUMEAMOUNT, "##,###")
        End If
        
Else
    ' searching schedule leads
    Set m_cari = M_DATA.QUERY_SEARCH(M_OBJCONN, "AGENT = '" + Combo1(4).Text + "' AND (NEXTACTDATE BETWEEN '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " 00:00" + "' AND '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " 23:59" + "') ", MDIForm1.Text3.Text)
        ListView1.ListItems.CLEAR
        SSTab1.Tab = 1
        ' searching schedule mgm
        ProgressBar1.Max = m_cari.RecordCount + 1
        Text2.Text = m_cari.RecordCount & " Data"
        While Not m_cari.EOF
        ProgressBar1.Value = m_cari.Bookmark
        Set ListItem = ListView1.ListItems.ADD(, , m_cari.Bookmark)
        ListItem.SubItems(1) = IIf(IsNull(m_cari("custid")), "", JADI_QUOTE(m_cari("custid")))
        Select Case m_cari("RECSTATUS")
        Case "1A"
            ListItem.SubItems(2) = "Available"
        Case ""
            ListItem.SubItems(2) = "Available"
        Case Else
            ListItem.SubItems(2) = IIf(IsNull(m_cari("PRIOR")), "", m_cari("PRIOR"))
        End Select
        ListItem.SubItems(3) = IIf(IsNull(m_cari("CUSTIDREF")), "", m_cari("CUSTIDREF"))
        ListItem.SubItems(4) = IIf(IsNull(m_cari("NAMAREF")), "", m_cari("NAMAREF"))
        ListItem.SubItems(5) = IIf(IsNull(m_cari("NAME")), "", JADI_QUOTE(m_cari("NAME")))
        ListItem.SubItems(6) = IIf(IsNull(m_cari("NEXTACTDATE")), "", Format(m_cari("NEXTACTDATE"), "yyyy/mm/dd hh:mm"))
        ListItem.SubItems(7) = IIf(IsNull(m_cari("NEXTACT")), "", m_cari("NEXTACT"))
        ListItem.SubItems(8) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS"))
        ListItem.SubItems(9) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
        ListItem.SubItems(10) = IIf(IsNull(m_cari("NamaAGENT")), "", m_cari("NamaAGENT"))
        ListItem.SubItems(11) = IIf(IsNull(m_cari("RECSOURCEREF")), "", m_cari("RECSOURCEREF"))
        ListItem.SubItems(12) = Format(IIf(IsNull(m_cari("TGLSTATUS")), "", m_cari("TGLSTATUS")), "yyyy/mm/dd")
        ListItem.SubItems(13) = IIf(IsNull(m_cari("Kethslkerja")), "", m_cari("Kethslkerja"))
        ListItem.SubItems(14) = IIf(IsNull(m_cari("KdComplaint")), "", m_cari("KdComplaint"))
        ListItem.SubItems(15) = IIf(IsNull(m_cari("RemarkComplaint")), "", m_cari("RemarkComplaint"))
        ListItem.SubItems(16) = IIf(IsNull(m_cari("KETHSLKERJA_NEW")), "", m_cari("KETHSLKERJA_NEW"))
        m_cari.MoveNext
    Wend
End If
Set M_DATA = Nothing

End Sub

Private Sub CmdMissed_Click()
'Dim M_DATA As New CLS_FRMSEARCH
'Dim listitem As listitem
'Dim VOLUMEAMOUNT As Double
'If Check2.Value = 1 Then
'    Call CEK_STATUS_F_CEK
'    Set m_cari = M_DATA.QUERY_SEARCH_mgm(M_OBJCONN, "AGENT = '" + Combo1(4).Text + "'  AND (NEXTACTDATE < '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " 00:00" + "') ", MDIForm1.Text3.Text, F_CEK, f_Pending, "", "", "")
'       LstVwSearchMgm.ListItems.CLEAR
'        SSTab1.Tab = 0
'        ' searching schedule mgm
'       ProgressBar1.Max = m_cari.RecordCount + 1
''        If Check2.Value = 1 Then
''            TxtJmlDtmgm.Text = m_cari.RecordCount & " Data"
''        Else
''            Text2.Text = m_cari.RecordCount & " Data"
''        End If
'        While Not m_cari.EOF
'            ProgressBar1.Value = m_cari.Bookmark
'
'    Set listitem = LstVwSearchMgm.ListItems.ADD(, , m_cari.Bookmark)
'        listitem.SubItems(1) = IIf(IsNull(m_cari("CUSTID")), "", m_cari("CUSTID"))
'        listitem.SubItems(3) = IIf(IsNull(m_cari("NAME")), "", m_cari("NAME"))
'        listitem.SubItems(4) = IIf(IsNull(m_cari("NEXTACTDATE")), "", Format(m_cari("NEXTACTDATE"), "yyyy/mm/dd hh:nn"))
'        listitem.SubItems(5) = IIf(IsNull(m_cari("NEXTACT")), "", m_cari("NEXTACT"))
'        listitem.SubItems(6) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS"))
'        'listitem.SubItems(7) = IIf(IsNull(m_cari("F_cek")), "", m_cari("F_cek"))
'        listitem.SubItems(8) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
'        'listitem.SubItems(9) = IIf(IsNull(m_cari("NamaAGENT")), "", m_cari("NamaAGENT"))
'        listitem.SubItems(10) = Format(IIf(IsNull(m_cari("Principal")), 0, m_cari("Principal")), "##,###")
'        listitem.SubItems(11) = Format(IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo")), "##,###")
'        listitem.SubItems(12) = Format(IIf(IsNull(m_cari("OpenDate")), "", m_cari("OpenDate")), "yyyy/mm/dd")
'        listitem.SubItems(13) = Format(IIf(IsNull(m_cari("TtlPTP")), 0, m_cari("TtlPTP")), "##,###")
'        listitem.SubItems(14) = IIf(IsNull(m_cari("RECSOURCE")), "", m_cari("RECSOURCE"))
'        listitem.SubItems(15) = IIf(IsNull(m_cari("TGLSTATUS")), "", Format(m_cari("TGLSTATUS"), "YYYY/MM/DD"))
'        listitem.SubItems(16) = IIf(IsNull(m_cari("Kethslkerja")), "", m_cari("Kethslkerja"))
'        listitem.SubItems(17) = Format(IIf(IsNull(m_cari("TGLINCOMING")), "", m_cari("TGLINCOMING")), "YYYY/MM/DD")
'        VOLUMEAMOUNT = VOLUMEAMOUNT + IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo"))
'        m_cari.MoveNext
'    Wend
'        If LstVwSearchMgm.ListItems.Count = 0 Then
'            TxtJmlDtMgm.Text = "Tidak Ada Data"
'            TxtJmlVolMgm.Text = "0"
'        Else
'            TxtJmlDtMgm.Text = "Total " + CStr(m_cari.RecordCount) + " Records"
'            TxtJmlVolMgm.Text = "Total " + CStr(m_cari.RecordCount)
'        End If
'
'Else
'    Set m_cari = M_DATA.QUERY_SEARCH(M_OBJCONN, "AGENT = '" + Combo1(4).Text + "'  AND (NEXTACTDATE < '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " 00:00" + "') ", MDIForm1.Text3.Text)
'        ListView1.ListItems.CLEAR
'        SSTab1.Tab = 1
'        ' searching schedule mgm
'        ProgressBar1.Max = m_cari.RecordCount + 1
'        Text2.Text = m_cari.RecordCount & " Data"
'        While Not m_cari.EOF
'        ProgressBar1.Value = m_cari.Bookmark
'        Set listitem = ListView1.ListItems.ADD(, , m_cari.Bookmark)
'        listitem.SubItems(1) = IIf(IsNull(m_cari("custid")), "", JADI_QUOTE(m_cari("custid")))
'        Select Case m_cari("RECSTATUS")
'        Case "1A"
'            listitem.SubItems(2) = "Available"
'        Case ""
'            listitem.SubItems(2) = "Available"
'        Case Else
'            listitem.SubItems(2) = IIf(IsNull(m_cari("PRIOR")), "", m_cari("PRIOR"))
'        End Select
'        listitem.SubItems(3) = IIf(IsNull(m_cari("CUSTIDREF")), "", m_cari("CUSTIDREF"))
'        listitem.SubItems(4) = IIf(IsNull(m_cari("NAMAREF")), "", m_cari("NAMAREF"))
'        listitem.SubItems(5) = IIf(IsNull(m_cari("NAME")), "", JADI_QUOTE(m_cari("NAME")))
'        listitem.SubItems(6) = IIf(IsNull(m_cari("NEXTACTDATE")), "", Format(m_cari("NEXTACTDATE"), "yyyy/mm/dd hh:mm"))
'        listitem.SubItems(7) = IIf(IsNull(m_cari("NEXTACT")), "", m_cari("NEXTACT"))
'        listitem.SubItems(8) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS"))
'        listitem.SubItems(9) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
'        listitem.SubItems(10) = IIf(IsNull(m_cari("NamaAGENT")), "", m_cari("NamaAGENT"))
'        listitem.SubItems(11) = IIf(IsNull(m_cari("RECSOURCEREF")), "", m_cari("RECSOURCEREF"))
'        listitem.SubItems(12) = Format(IIf(IsNull(m_cari("TGLSTATUS")), "", m_cari("TGLSTATUS")), "yyyy/mm/dd")
'        listitem.SubItems(13) = IIf(IsNull(m_cari("Kethslkerja")), "", m_cari("Kethslkerja"))
'        listitem.SubItems(14) = IIf(IsNull(m_cari("KdComplaint")), "", m_cari("KdComplaint"))
'        listitem.SubItems(15) = IIf(IsNull(m_cari("RemarkComplaint")), "", m_cari("RemarkComplaint"))
'        listitem.SubItems(16) = IIf(IsNull(m_cari("F_CEK")), "", m_cari("F_CEK"))
'        m_cari.MoveNext
'    Wend
'End If
End Sub
Private Sub CEK_STATUS_F_CEK()
Dim cmdsql As String
Dim M_Objrs As New ADODB.Recordset

F_CEK = Empty
Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        cmdsql = "SELECT * FROM usertbl WHERE USERID = '" + MDIForm1.Text1.Text + "'"
         M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        
        
        While Not M_Objrs.EOF
                StsVl = CStr(Trim(IIf(IsNull(M_Objrs!f_VL), "", M_Objrs!f_VL)))
                StsPR = CStr(Trim(IIf(IsNull(M_Objrs!f_PR), "", M_Objrs!f_PR)))
                StsPTP = CStr(Trim(IIf(IsNull(M_Objrs!f_ptp), "", M_Objrs!f_ptp)))
                StsBP = CStr(Trim(IIf(IsNull(M_Objrs!f_bp), Empty, M_Objrs!f_bp)))
                StsPOP = CStr(Trim(IIf(IsNull(M_Objrs!f_pop), "", M_Objrs!f_pop)))
                StsSP = CStr(Trim(IIf(IsNull(M_Objrs!f_sp), "", M_Objrs!f_sp)))
                StsRP = CStr(Trim(IIf(IsNull(M_Objrs!f_rp), "", M_Objrs!f_rp)))
                StsUC = CStr(Trim(IIf(IsNull(M_Objrs!F_UC), "", M_Objrs!F_UC)))
                StsSK = CStr(Trim(IIf(IsNull(M_Objrs!f_SK), "", M_Objrs!f_SK)))
                StsON = CStr(Trim(IIf(IsNull(M_Objrs!f_ON), "", M_Objrs!f_ON)))
                StsOS = CStr(Trim(IIf(IsNull(M_Objrs!f_OS), "", M_Objrs!f_OS)))
                LUserType = CStr(Trim(IIf(IsNull(M_Objrs!usertype), "", M_Objrs!usertype)))
                Bloked = Replace(IIf(IsNull(M_Objrs!lockdarispv), "", M_Objrs!lockdarispv), "@", "'")
                Stsblank = CStr(Trim(IIf(IsNull(M_Objrs!F_blank), "", M_Objrs!F_blank)))
                StsWO_Date = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_DATE), "", M_Objrs!f_WO_DATE)))
                StsWO_2009 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2009), "", M_Objrs!f_WO_2009)))
                StsWO_2010 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2010), "", M_Objrs!f_WO_2010)))
                StsWO_2008 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2008), "", M_Objrs!f_WO_2008)))
                StsWO_2007 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2007), "", M_Objrs!f_WO_2007)))
                StsWO_2006 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2006), "", M_Objrs!f_WO_2006)))
                StsWO_2005 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2005), "", M_Objrs!f_WO_2005)))
                StsWO_2004 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2004), "", M_Objrs!f_WO_2004)))
                StsWO_2003 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2003), "", M_Objrs!f_WO_2003)))
                StsWO_2002 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2002), "", M_Objrs!f_WO_2002)))
                StsWO_2001 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2001), "", M_Objrs!f_WO_2001)))
                StsWO_2000 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2000), "", M_Objrs!f_WO_2000)))
                StsWO_1999 = CStr(Trim(IIf(IsNull(M_Objrs!F_WO_1999), "", M_Objrs!F_WO_1999)))
                M_Objrs.MoveNext
            Wend
            Set M_Objrs = Nothing
             StsAll = StsVl + StsPR + StsPTP + StsBP + StsPOP + StsSP + StsRP + StsUC + StsON + StsSK + StsOS
            
            
        If StsAll <> "" Then
            If LUserType = "1" Then
                    If StsUC = "UC" Then
                       If Bloked <> "" Then
                       
                            F_CEK = " + Bloked + "
                       Else
                            F_CEK = "(substring(F_CEK_NEW,1,3)IN( '" + StsVl + "','" + StsPR + "','" + StsPTP + "','" + StsBP + "','" + StsPOP + "','" + StsSP + "','" + StsRP + "','" + StsSK + "','" + StsON + "','" + StsOS + "') or F_CEK_NEW IS NULL) "
                       End If
                       
                    Else
                     If Bloked <> "" Then
                            F_CEK = "(" + Bloked + " )"
                    Else
                         F_CEK = "(substring(F_CEK_NEW,1,3)IN( '" + StsVl + "','" + StsPR + "','" + StsPTP + "','" + StsBP + "','" + StsPOP + "','" + StsSP + "','" + StsRP + "','" + StsSK + "','" + StsON + "','" + StsOS + "') or F_CEK_NEW IS NULL) "
                    End If
                    
                    End If
                        
                End If
                Else
                
                 If Bloked <> "" Then
                            F_CEK = "(" + Bloked + " )"
                 End If
     End If
     
  
      
'     If StsAll <> Empty Then
'            If LUserType = "1" Then
'                    If StsUC <> Empty Then
'                        If Bloked <> "" Then
'                            F_CEK = "(" + Bloked + " )"
'                        Else
'                          '  F_CEK = " substring(F_CEK_NEW,1,3)  IN('" + StsVl + "','" + StsPR + "','" + StsPTP + "','" + StsBP + "','" + StsPOP + "','" + StsSP + "','" + StsSK + "', '" + StsON + "','" + StsOS + "','') "
'                            F_CEK = "(substring(F_CEK_NEW,1,3)IN( '" + StsVl + "','" + StsPR + "','" + StsPTP + "','" + StsBP + "','" + StsPOP + "','" + StsSP + "','" + StsRP + "','" + StsSK + "','" + StsON + "','" + StsOS + "') or F_CEK_NEW IS NULL) "
'                        End If
'
'             '           F_CEK = "(substring(F_CEK_NEW,1,3)IN('NK-','MV-','WN-','NBP','" + StsNa + "','" + StsOP + "','" + StsPTP + "','" + StsBP + "','" + StsPOP + "','" + StsSP + "','" + StsRP + "') or F_CEK_NEW IS NULL)"
'              '          Else
'                        'F_CEK = "(substring(F_CEK_NEW,1,3)IN( '" + StsVl + "','" + StsPR + "','" + StsPTP + "','" + StsBP + "','" + StsPOP + "','" + StsSP + "','" + StsRP + "','" + StsSK + "','" + StsON + "','" + StsOS + "') or F_CEK_NEW IS NULL) "
'                    End If
'
'            End If
'     End If

End Sub

Private Sub CmdSearchPTP_Click()
    '@@ 24 Januari 2012, Buat Search PTP
    Dim Cmdsql_user As String
    Dim M_Objrs As ADODB.Recordset
    Dim M_WHERE As String
    Dim Status_PTP As String
    Dim ListItem As ListItem
    
    Dim totamount As Double
    Dim TOTCURBALANCE As Double
    Dim VOLUMEAMOUNT As Double
    
    M_WHERE = ""
    
    If UCase(Trim(MDIForm1.Text2.Text)) = "AGENT" Then
        'Cek dulu Apakah agent tersebut Dapat Melihat Status All PTP atau Hanya Sebagian PTP
        Cmdsql_user = "select f_status_ptp from usertbl where userid='"
        Cmdsql_user = Cmdsql_user + Trim(MDIForm1.Text1.Text) + "'"
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        M_Objrs.Open Cmdsql_user, M_OBJCONN, adOpenDynamic, adLockOptimistic
        
        'Jika Data ditemukan
        If M_Objrs.RecordCount > 0 Then
            Status_PTP = IIf(IsNull(M_Objrs("f_status_ptp")), "", M_Objrs("f_status_ptp"))
        End If
        Set M_Objrs = Nothing
        
        'set kriteria SQL PTP
        M_WHERE = " where agent='" + Trim(MDIForm1.Text1.Text) + "'  "
        If Status_PTP = "" Then
            'M_WHERE = M_WHERE + " and custid in (select custid from reportptp where promisedate between "
            'M_WHERE = M_WHERE + "date(now()) and date(now())+3 ) "
            '@@ 03-04-2012, Diubah berdasarkan tanggal tagih
            M_WHERE = M_WHERE + " and date(tgl_tagih) between date(now()) and date(now())+3 "
        End If
        
        M_WHERE = M_WHERE + " and substring(f_cek_new,1,3)='PTP' "
    End If
    
    If UCase(Trim(MDIForm1.Text2.Text)) = "TEAMLEADER" Then
        M_WHERE = " where agent in (select userid from usertbl where team='"
        M_WHERE = M_WHERE + MDIForm1.Text1.Text + "') and substring(f_cek_new,1,3)='PTP' "
        M_WHERE = M_WHERE + " and date(tgl_tagih) is not null "
    ElseIf UCase(Trim(MDIForm1.Text2.Text)) = "ADMIN" Or UCase(Trim(MDIForm1.Text2.Text)) = "ADMINISTRATOR" Or UCase(Trim(MDIForm1.Text2.Text)) = "SUPERVISOR" Then
        M_WHERE = " where substring(f_cek_new,1,3)='PTP' "
        M_WHERE = M_WHERE + " and date(tgl_tagih) is not null "
    End If
    
    cmdsql = " select * from mgm  " + M_WHERE
    
    Set M_Objrs = New ADODB.Recordset
    M_Objrs.CursorLocation = adUseClient
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs.RecordCount = 0 Then
     MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
     Set M_Objrs = Nothing
     Exit Sub
    End If
    
    LstVwSearchMgm.ListItems.CLEAR
    
    While Not M_Objrs.EOF
        Set ListItem = LstVwSearchMgm.ListItems.ADD(, , M_Objrs.Bookmark)
        'statusprior = IIf(IsNull(M_Objrs("StatusPrior")), "", M_Objrs("StatusPrior"))
        ListItem.SubItems(1) = IIf(IsNull(M_Objrs("CUSTID")), "", M_Objrs("CUSTID"))
        ListItem.SubItems(2) = IIf(IsNull(M_Objrs("PRIOR")), "", M_Objrs("PRIOR"))
        ListItem.SubItems(3) = IIf(IsNull(M_Objrs("NAME")), "", M_Objrs("NAME"))
        ListItem.SubItems(4) = IIf(IsNull(M_Objrs("RECSOURCE")), "", M_Objrs("RECSOURCE"))
        ListItem.SubItems(5) = IIf(IsNull(M_Objrs("NEXTACTDATE")), "", Format(M_Objrs("NEXTACTDATE"), "dd/mm/yyyy hh:nn"))
        ListItem.SubItems(6) = IIf(IsNull(M_Objrs("NEXTACT")), "", M_Objrs("NEXTACT"))
        ListItem.SubItems(7) = IIf(IsNull(M_Objrs("REMARKS")), "", M_Objrs("REMARKS"))
        ListItem.SubItems(8) = CStr(IIf(IsNull(M_Objrs("kethslkerja_new")), "", M_Objrs("kethslkerja_new")) & " ")  'sPending)
        ListItem.SubItems(9) = IIf(IsNull(M_Objrs("StatusCall")), "", M_Objrs("StatusCall"))
        ListItem.SubItems(11) = IIf(IsNull(M_Objrs("AGENT")), "", M_Objrs("AGENT"))
        
        
        If UCase(MDIForm1.Text2) <> "SUPERVISOR" Then
            If Format(IIf(IsNull(M_Objrs("flaglead")), 0, M_Objrs("flaglead")), "##,###") = 1 Then
                ListItem.SubItems(12) = ""
            Else
                ListItem.SubItems(12) = Format(IIf(IsNull(M_Objrs("Principal")), 0, M_Objrs("Principal")), "##,###")
            End If
        Else
             ListItem.SubItems(12) = Format(IIf(IsNull(M_Objrs("Principal")), 0, M_Objrs("Principal")), "##,###")
        End If
        
        ListItem.SubItems(13) = Format(IIf(IsNull(M_Objrs("AmountWo")), 0, M_Objrs("AmountWo")), "##,###")
        totamount = totamount + IIf(IsNull(M_Objrs("AmountWo")), 0, M_Objrs("AmountWo"))
        
        
        ListItem.SubItems(14) = Format(IIf(IsNull(M_Objrs("OpenDate")), "", M_Objrs("OpenDate")), "yyyy/mm/dd")
        ListItem.SubItems(15) = Format(IIf(IsNull(M_Objrs("B_D")), 0, M_Objrs("B_D")))
        ListItem.SubItems(16) = Format(IIf(IsNull(M_Objrs("Pay_Dt")), 0, M_Objrs("Pay_Dt")), "yyyy/mm/dd")
        
        ListItem.SubItems(17) = Format(IIf(IsNull(M_Objrs("lastpay")), 0, M_Objrs("lastpay")), "##,###")
        
        ListItem.SubItems(18) = IIf(IsNull(M_Objrs("TGLSTATUS")), "", Format(M_Objrs("TGLSTATUS"), "YYYY/MM/DD"))
        ListItem.SubItems(19) = IIf(IsNull(M_Objrs("TGLCALL")), "", Format(M_Objrs("TGLCALL"), "YYYY/MM/DD"))
        ListItem.SubItems(20) = IIf(IsNull(M_Objrs("Kethslkerja")), "", M_Objrs("Kethslkerja"))
        ListItem.SubItems(21) = Format(IIf(IsNull(M_Objrs("TGLINCOMING")), "", M_Objrs("TGLINCOMING")), "YYYY/MM/DD")
        ListItem.SubItems(23) = IIf(IsNull(M_Objrs("resultcpa")), "", M_Objrs("resultcpa"))
        ListItem.SubItems(24) = IIf(IsNull(M_Objrs("tglinsertfrmcpa")), "", M_Objrs("tglinsertfrmcpa"))
        ListItem.SubItems(25) = Format(IIf(IsNull(M_Objrs("curbal")), "", M_Objrs("curbal")), "##,###")
        TOTCURBALANCE = TOTCURBALANCE + IIf(IsNull(M_Objrs("curbal")), 0, M_Objrs("curbal"))
       
        ListItem.SubItems(26) = IIf(IsNull(M_Objrs("dob")), "", Format(M_Objrs("dob"), "yyyy-mm-dd"))
        ListItem.SubItems(27) = IIf(IsNull(M_Objrs("ktpno")), "", M_Objrs("ktpno"))
        
            
SorryLompat:
        
        VOLUMEAMOUNT = VOLUMEAMOUNT + IIf(IsNull(M_Objrs("AmountWo")), 0, M_Objrs("AmountWo"))
        
        M_Objrs.MoveNext
    Wend
    
    txtamount.Text = Format(totamount, "##,###")
    txtcurbalance.Text = Format(TOTCURBALANCE, "##,###")
    
    If LstVwSearchMgm.ListItems.Count = 0 Then
        TxtJmlDtMgm.Text = "Tidak Ada Data"
        TxtJmlVolMgm.Text = "0"
    Else
        TxtJmlDtMgm.Text = "Total " + CStr(M_Objrs.RecordCount) + " Records"
        TxtJmlVolMgm.Text = Format(VOLUMEAMOUNT, "##,###")
    End If
    
    Set M_Objrs = Nothing
End Sub

Private Sub Combo1_Click(Index As Integer)



Dim M_DATA As New CLS_FRMSEARCH
Dim M_Objrs As ADODB.Recordset
Select Case Index
Case 0
        If Combo1(0).Text = "LUNAS" Then
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
        Exit Sub
        End If
    Set M_Objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "USERID = '" + Combo1(Index).Text + "'")
    If M_Objrs.RecordCount <> 0 Then
        Combo1(0).Text = M_Objrs("USERID")
        Combo1(1).Text = M_Objrs("AGENT")
    Else
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
    End If
    
   

Case 1
    Set M_Objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "AGENT = '" + Combo1(Index).Text + "'")
    If M_Objrs.RecordCount <> 0 Then
        Combo1(0).Text = M_Objrs("USERID")
        Combo1(1).Text = M_Objrs("AGENT")
    Else
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
    End If
Case 2
Set M_Objrs = M_DATA.QUERY_DATASOURCE(M_OBJCONN, "KODEDS = '" + Combo1(Index).Text + "'")
    If M_Objrs.RecordCount <> 0 Then
        Combo1(2).Text = M_Objrs("KODEDS")
        Combo1(3).Text = M_Objrs("KETERANGAN")
    Else
        Combo1(2).Text = Empty
        Combo1(3).Text = Empty
    End If
Case 3
Set M_Objrs = M_DATA.QUERY_DATASOURCE(M_OBJCONN, "KETERANGAN = '" + Combo1(Index).Text + "'")
    If M_Objrs.RecordCount <> 0 Then
        Combo1(2).Text = M_Objrs("KODEDS")
        Combo1(3).Text = M_Objrs("KETERANGAN")
    Else
        Combo1(2).Text = Empty
        Combo1(3).Text = Empty
    End If
Case 4
    Set M_Objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "USERID = '" + Combo1(Index).Text + "'")
    If M_Objrs.RecordCount <> 0 Then
        Combo1(4).Text = M_Objrs("USERID")
        Combo1(5).Text = M_Objrs("AGENT")
    Else
        Combo1(4).Text = Empty
        Combo1(5).Text = Empty
    End If
Case 5
    Set M_Objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "AGENT = '" + Combo1(Index).Text + "'")
    If M_Objrs.RecordCount <> 0 Then
        Combo1(4).Text = M_Objrs("USERID")
        Combo1(5).Text = M_Objrs("AGENT")
    Else
        Combo1(4).Text = Empty
        Combo1(5).Text = Empty
    End If
End Select
Set M_DATA = Nothing
Set M_Objrs = Nothing
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
Dim sSearchText As String
Dim lReturn As Long
Select Case Index
Case 0, 1, 2, 3
If KeyAscii = 13 Then
   Combo1_Click (Index)
   KeyAscii = 0
Else
   sSearchText = Left$(Combo1(Index).Text, Combo1(Index).SelStart) & Chr$(KeyAscii)
   lReturn = SendMessage(Combo1(Index).hwnd, CB_FINDSTRING, -1, ByVal sSearchText)
   If lReturn <> CB_ERR Then
      mbIgnoreListClick = True
      Combo1(Index).ListIndex = lReturn
      mbIgnoreListClick = False
      Combo1(Index).Text = Combo1(Index).list(lReturn)
      Combo1(Index).SelStart = Len(sSearchText)
      Combo1(Index).SelLength = Len(Combo1(Index).Text)
      KeyAscii = 0
   End If
End If
End Select
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim M_DATA As New CLS_FRMSEARCH
Dim M_Objrs As ADODB.Recordset
Select Case Index
Case 0
    Set M_Objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "USERID = '" + Combo1(Index).Text + "'")
    If M_Objrs.RecordCount <> 0 Then
        Combo1(0).Text = M_Objrs("USERID")
        Combo1(1).Text = M_Objrs("AGENT")
    Else
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
    End If
Case 1
    Set M_Objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "AGENT = '" + Combo1(Index).Text + "'")
    If M_Objrs.RecordCount <> 0 Then
        Combo1(0).Text = M_Objrs("USERID")
        Combo1(1).Text = M_Objrs("AGENT")
    Else
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
    End If
Case 2
Set M_Objrs = M_DATA.QUERY_DATASOURCE(M_OBJCONN, "KODEDS = '" + Combo1(Index).Text + "'")
    If M_Objrs.RecordCount <> 0 Then
        Combo1(2).Text = M_Objrs("KODEDS")
        Combo1(3).Text = M_Objrs("KETERANGAN")
    Else
        Combo1(2).Text = Empty
        Combo1(3).Text = Empty
    End If
Case 3
Set M_Objrs = M_DATA.QUERY_DATASOURCE(M_OBJCONN, "KETERANGAN = '" + Combo1(Index).Text + "'")
    If M_Objrs.RecordCount <> 0 Then
        Combo1(2).Text = M_Objrs("KODEDS")
        Combo1(3).Text = M_Objrs("KETERANGAN")
    Else
        Combo1(2).Text = Empty
        Combo1(3).Text = Empty
    End If
End Select
Set M_DATA = Nothing
Set M_Objrs = Nothing
End Sub

Private Sub Form_Activate()

'    Dim m_objrsPesanReset As ADODB.Recordset
'    Dim m_objrsPesanLock As ADODB.Recordset
'    Dim CMDSQL As String
'
'    'Cek pesan reset
'    CMDSQL = "select f_pesanresetauto from usertbl where userid='"
'    CMDSQL = CMDSQL + Trim(MDIForm1.Text1.Text) + "'"
'    Set m_objrsPesanReset = New ADODB.Recordset
'    m_objrsPesanReset.CursorLocation = adUseClient
'    m_objrsPesanReset.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'    If m_objrsPesanReset("f_pesanresetauto") = "1" Then
'        MsgBox "Reset Data! Ini adalah lock data automatic, data anda akan segera diperbaharui!", vbOKOnly + vbInformation, "Informasi"
'        CMDSQL = "update usertbl set f_pesanresetauto=null where userid='"
'        CMDSQL = CMDSQL + Trim(MDIForm1.Text1.Text) + "'"
'        M_OBJCONN.Execute CMDSQL
'    End If
'
'    Set m_objrsPesanReset = Nothing
'
'    'Cek pesan Lock
'    CMDSQL = "select f_pesanlockauto from usertbl where userid='"
'    CMDSQL = CMDSQL + Trim(MDIForm1.Text1.Text) + "'"
'    Set m_objrsPesanLock = New ADODB.Recordset
'    m_objrsPesanLock.CursorLocation = adUseClient
'    m_objrsPesanLock.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'    If m_objrsPesanLock("f_pesanlockauto") = "1" Then
'        MsgBox "Lock Data! Ini adalah lock data automatic, data anda akan segera diperbaharui!", vbOKOnly + vbInformation, "Informasi"
'        CMDSQL = "update usertbl set f_pesanlockauto=null where userid='"
'        CMDSQL = CMDSQL + Trim(MDIForm1.Text1.Text) + "'"
'        M_OBJCONN.Execute CMDSQL
'    End If
'
'    Set m_objrsPesanLock = Nothing
    Call tampil_waktu
End Sub

Public Sub tampil_waktu()
    Dim m_objwaktu As ADODB.Recordset
    
    Set m_objwaktu = New ADODB.Recordset
    m_objwaktu.CursorLocation = adUseClient
    m_objwaktu.Open "SELECT now() as waktu", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    Label5.Caption = IIf(IsNull(m_objwaktu!waktu), "", Format(m_objwaktu!waktu, "hh : mm "))
    Set m_objwaktu = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Exit Sub
End Sub

Private Sub Form_Load()
Dim M_Objrs As ADODB.Recordset
Dim M_DATA As New CLS_FRMSEARCH

Call HEADER_VIEW_mgm
Call HEADER_VIEW_Refferall

Combo2.AddItem "OS"
Combo2.AddItem "ON"
Combo2.AddItem "PR"
Combo2.AddItem "BP"
Combo2.AddItem "RP"
Combo2.AddItem "SK"
Combo2.AddItem "POP"
Combo2.AddItem "SP"
Combo2.AddItem "VL"
Combo2.AddItem "PTP"
Combo2.AddItem "HOT PROSPECT"

If UCase(MDIForm1.Text2.Text) = "AGENT" Then
    CekDtDistribute.Visible = False
    Combo1(4).Text = MDIForm1.Text1.Text
    Combo1(5).Text = MDIForm1.Text7.Text
    Combo1(4).Visible = False
    Combo1(5).Visible = False
    Label1(13).Visible = False
    Combo1(0).Text = MDIForm1.Text1.Text
    Combo1(1).Text = MDIForm1.Text7.Text
    cmd_review.Visible = False
    cmd_claimback_acc.Visible = False
    Cmd_listrequestdecease.Visible = False
Else
    CekDtDistribute.Visible = True
End If

    If UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
        cmd_review.Visible = False
        cmd_claimback_acc.Visible = False
    End If

    CmbStatusCek.AddItem "ACCEPT"
    CmbStatusCek.AddItem "RETURN"
    CmbStatusCek.AddItem "NOT CHECK"
    SSTab1.Tab = 0
    SSTab2.Tab = 0
    
Me.Width = 11880
Me.Height = 9945
'Me.Height = 6105
SSTab1.TabVisible(1) = False
Me.Top = 500
Me.Left = 1000


cmbStsLastCall(0).AddItem "NOT CONTACTED"
cmbStsLastCall(0).AddItem "NOT AVAILABLE"
cmbStsLastCall(0).AddItem "STILL THINKING"
cmbStsLastCall(0).AddItem "DISAGREE"
cmbStsLastCall(0).AddItem "SENDING APPLICATION"
cmbStsLastCall(0).AddItem "PICKUP"
cmbStsLastCall(0).AddItem "INCOMPLETE DOCUMENT"
cmbStsLastCall(0).AddItem "INCOMING"
cmbStsLastCall(0).AddItem "AVAILABLE"

DTimeLastCall(0).Value = "00:00"
DTimeLastCall(1).Value = "23:59"
StsmgmSchedule = False


Set M_Objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "")
    While Not M_Objrs.EOF
        Combo1(0).AddItem M_Objrs("USERID")
        Combo1(1).AddItem M_Objrs("AGENT")
        Combo1(4).AddItem M_Objrs("USERID")
        Combo1(5).AddItem M_Objrs("AGENT")
        M_Objrs.MoveNext
    Wend
 Set M_Objrs = Nothing
  
 
 Set M_Objrs = M_DATA.QUERY_DATASOURCE(M_OBJCONN, "")
    While Not M_Objrs.EOF
        Combo1(2).AddItem M_Objrs("KODEDS")
        Combo1(3).AddItem M_Objrs("KETERANGAN")
        M_Objrs.MoveNext
    Wend
If UCase(MDIForm1.Text3.Text) = "ADMIN" Then
    Label1(5).Visible = True
    Text1(2).Visible = True
End If
M_Objrs.Close
Set M_Objrs = Nothing
Set M_DATA = Nothing
End Sub
Private Sub show_Search_mgmData()
Dim harga As Double
Dim ListItem As ListItem
Dim Lcustid1 As String
Dim Lcustid2 As String
Dim LCall As String
Dim I As Integer
Dim cmdsql As String
Dim sPending As String
Dim M_Objrs As ADODB.Recordset
Dim VOLUMEAMOUNT As Double
Dim statusprior As String
Dim exp%
Dim totamount As Double
Dim TOTCURBALANCE As Double

Dim expired_claim As Integer
Dim sts_data_acc() As String
Dim rs_cek_data As ADODB.Recordset

Dim tgl_janji As Date
Dim tgl_bayar As Date

I = 1
'On Error GoTo HELL
        
    datajml = m_cari.RecordCount
    LstVwSearchMgm.ListItems.CLEAR
    Me.MousePointer = vbHourglass
    ProgressBar1.Max = m_cari.RecordCount + 1
    
    '@@19-11-10 ///////// Ini tambahan buat mencatat custid per session yang dicolect agent pd blok data
        'cek status id lock session di tabel usertbl
        Dim CekSession As String
        Dim M_Objrs_Session As ADODB.Recordset
        Dim NilaiSession As String
        
        Dim CmdWaktuServer As String
        Dim m_ObjrsWktServer As ADODB.Recordset
        Dim WaktuServer As Date
        
        Dim SimpanDtSession As String
        
        'Ambil Id Session Start
        CekSession = "select f_idsessstart from usertbl where userid='"
        CekSession = CekSession + Trim(MDIForm1.Text1.Text) + "'"
        Set M_Objrs_Session = New ADODB.Recordset
        M_Objrs_Session.CursorLocation = adUseClient
        M_Objrs_Session.Open CekSession, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_Objrs_Session.RecordCount <> 0 Then
         NilaiSession = IIf(IsNull(M_Objrs_Session("f_idsessstart")), "", M_Objrs_Session("f_idsessstart"))
        Else
         NilaiSession = ""
        End If
        Set M_Objrs_Session = Nothing
        
        'Ambil Waktu Server Terkini
        CmdWaktuServer = "select now() "
        Set m_ObjrsWktServer = New ADODB.Recordset
        m_ObjrsWktServer.CursorLocation = adUseClient
        m_ObjrsWktServer.Open CmdWaktuServer, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        WaktuServer = Format(m_ObjrsWktServer(0), "yyyy-mm-dd hh:mm:ss")
        Set m_ObjrsWktServer = Nothing
    
    '@@20-11-10 ///////// Ini tambahan buat mencatat custid per session yang dicolect agent pd blok data
    
    
    While Not m_cari.EOF
        
        '@@20-11-10 ///Ini buat mencatat custid dan status awal yang dikerjain agent per session ////
            If NilaiSession <> Empty Then
                '@@ 04-10-2011, Log session log data dinonaktifkan terlebih dahulu, karena lama
'                SimpanDtSession = "insert into tblperformpersessionlock(idlock,"
'                SimpanDtSession = SimpanDtSession + "startlock,agent,custid,name,"
'                SimpanDtSession = SimpanDtSession + "f_ceklalu,tgl_f_ceklalu) values ('"
'                SimpanDtSession = SimpanDtSession + Trim(NilaiSession) + "','"
'                SimpanDtSession = SimpanDtSession + CStr(Format(WaktuServer, "yyyy-mm-dd hh:mm:ss")) + "','"
'                SimpanDtSession = SimpanDtSession + Trim(m_cari!agent) + "','"
'                SimpanDtSession = SimpanDtSession + Trim(m_cari!CustId) + "','"
'                SimpanDtSession = SimpanDtSession + Trim(m_cari!Name) + "','"
'                SimpanDtSession = SimpanDtSession + IIf(IsNull(m_cari!f_cek_new), "", Trim(m_cari!f_cek_new)) + "',"
'                SimpanDtSession = SimpanDtSession + IIf(IsNull(m_cari!tglcall), "null", "'" + CStr(Format(m_cari!tglcall, "yyyy-mm-dd hh:mm:ss")) + "'") + ")"
'                M_OBJCONN.Execute SimpanDtSession
            End If
        '@@20-11-10 ///Ini buat mencatat custid dan status awal yang dikerjain agent per session ////
    
        ProgressBar1.Value = m_cari.Bookmark
        Lcustid1 = CStr(IIf(IsNull(m_cari!CustId), "", m_cari!CustId))
        sPending = CStr(Trim(IIf(IsNull(m_cari!f_Pending), "", m_cari!f_Pending)))
        
        ' 12 SEP 2013 - CEK CLAIM ACC -> JIKA LEWAT 3HR TDK PTP SET KE AGENT LAMA, JIKA STATUS PTP DAN LEBIH DARI 3HR TDK BAYAR SET KE AGENT LAMA
        If UCase(MDIForm1.Text2) = "AGENT" Then
            ' CEK DATA APPROVAL CLAIM
            If IIf(IsNull(m_cari!app_claim), "", Trim(m_cari!app_claim)) <> "" Then
                If IIf(IsNull(m_cari!f_cek_new), "", m_cari!f_cek_new) <> "" Then
                    expired_claim = DateDiff("d", Format(m_cari!app_claim, "yyyy-mm-dd"), Format(WaktuServer, "yyyy-mm-dd"))
                    sts_data_acc = Split(m_cari!f_cek_new, "-")

                   'JIKA LEWAT 15HR TDK PTP SET KE AGENT LAMA REQ JOKO 16 JUNI 2014
                    If expired_claim >= 15 And sts_data_acc(0) <> "PTP" Then

                        If IIf(IsNull(m_cari!agent_asli), "", m_cari!agent_asli) <> "" Then
                            ' Kalau Account Broken Promise, ON Nego, Prospect 04 Juni 2014
                            If (sts_data_acc(0) = "BP") Or (sts_data_acc(0) = "ON") Or (sts_data_acc(0) = "PR") Or (sts_data_acc(0) = "") Or (sts_data_acc(0) = "OS") Or (sts_data_acc(0) = "VL") Then
                                ' INSERT KE LOG SET AGENT LAMA CLAIM
                                M_OBJCONN.Execute "INSERT INTO log_claim_back(custid,agent_claim,agent_asli,reason,tgl_claim) VALUES('" & Lcustid1 & "','" & MDIForm1.Text1.Text & "','" & m_cari!agent_asli & "','Status belum sampai PTP status terakhir " & sts_data_acc(0) & "','" & Format(m_cari!app_claim, "yyyy-mm-dd") & "')"
    
                                M_OBJCONN.Execute "INSERT INTO log_claim_back_hst(custid,agent_claim,agent_asli,reason,tgl_claim) VALUES('" & Lcustid1 & "','" & MDIForm1.Text1.Text & "','" & m_cari!agent_asli & "','Return To Agent Asli','" & Format(m_cari!app_claim, "yyyy-mm-dd") & "')"
    
                                M_OBJCONN.Execute "UPDATE mgm SET agent=agent_asli,app_claim=null WHERE custid='" & Lcustid1 & "'"
                                GoTo SorryLompat
                            End If
                        End If

                    ElseIf sts_data_acc(0) = "PTP" Then
                       'JIKA STATUS PTP DAN LEBIH DARI 3HR TDK BAYAR
                        If IIf(IsNull(m_cari!agent_asli), "", m_cari!agent_asli) <> "" Then
                            ' CEK TGL PTP DAN PEMBAYARAN
                            Set rs_cek_data = New ADODB.Recordset
                            rs_cek_data.ActiveConnection = M_OBJCONN
                            rs_cek_data.CursorType = adOpenDynamic
                            rs_cek_data.LockType = adLockOptimistic
                            rs_cek_data.CursorLocation = adUseClient
                            ' TGL PTP
                            rs_cek_data.Open "SELECT max(promisedate) as tglpromise FROM tblnegoptp WHERE custid='" & Lcustid1 & "'"

                            If rs_cek_data.RecordCount > 0 Then
                                tgl_janji = Format(rs_cek_data!tglpromise, "yyyy-mm-dd")

                                'CEK PEMBAYARAN
                                If rs_cek_data.state = 1 Then rs_cek_data.Close
                                rs_cek_data.Open "SELECT date(paydate) as tgl_bayar FROM tbllunas WHERE custid='" & Lcustid1 & "' AND date(paydate) > '" & Format(tgl_janji, "yyyy-mm-dd") & "'"

                                ' Jika tidak ada pembayaran 3 hari expired claim
                                If rs_cek_data.RecordCount = 0 Then
                                    '--- Tgl bayar
                                    expired_claim = DateDiff("d", tgl_janji, Format(WaktuServer, "yyyy-mm-dd"))

                                    If expired_claim > 3 Then
                                        ' INSERT KE LOG SET AGENT LAMA CLAIM
                                        M_OBJCONN.Execute "INSERT INTO log_claim_back(custid,agent_claim,agent_asli,reason,tgl_janji) VALUES('" & Lcustid1 & "','" & MDIForm1.Text1.Text & "','" & m_cari!agent_asli & "','Status PTP tetapi belum ada pembayaran','" & Format(tgl_janji, "yyyy-mm-dd") & "')"

                                        M_OBJCONN.Execute "INSERT INTO log_claim_back_hst(custid,agent_claim,agent_asli,reason,tgl_janji) VALUES('" & Lcustid1 & "','" & MDIForm1.Text1.Text & "','" & m_cari!agent_asli & "','Return To Agent Asli','" & Format(tgl_janji, "yyyy-mm-dd") & "')"

                                        M_OBJCONN.Execute "UPDATE mgm SET agent=agent_asli,app_claim=null WHERE custid='" & Lcustid1 & "'"

                                        Set rs_cek_data = Nothing

                                        GoTo SorryLompat
                                    End If
                                Else
                                    tgl_bayar = rs_cek_data!tgl_bayar
                                    Set rs_cek_data = Nothing
                                End If
                            End If
                        End If

                    End If

                End If
            End If
        End If
'        END CEK CLAIM ACC -------------------------------------------------------------------------------------
        
        Set ListItem = LstVwSearchMgm.ListItems.ADD(, , m_cari.Bookmark)
        
        If MDIForm1.Text2 = "TeamLeader" Then
            If IIf(IsNull(m_cari("stscpa")), "0", m_cari("stscpa")) = 1 Then
                ListItem.ForeColor = vbRed
            End If
            
            If IIf(IsNull(m_cari("intapprovel")), "0", m_cari("intapprovel")) = 1 Then
              ListItem.ForeColor = vbBlue
            End If

        End If
        
        If UCase(MDIForm1.Text7) = "JOKO" Or UCase(MDIForm1.Text7) = "WULANDARI" Or UCase(MDIForm1.Text7) = "ANDRI" Then
            If IIf(IsNull(m_cari("intverify")), "0", m_cari("intverify")) = 1 Then
                ListItem.ForeColor = vbYellow
            End If
            
            If IIf(IsNull(m_cari("intapprovel")), "0", m_cari("intapprovel")) = 1 Then
              ListItem.ForeColor = vbGreen
            End If
        End If
        
        
        statusprior = IIf(IsNull(m_cari("StatusPrior")), "", m_cari("StatusPrior"))
        ListItem.SubItems(1) = IIf(IsNull(m_cari("CUSTID")), "", m_cari("CUSTID"))
        ListItem.SubItems(2) = IIf(IsNull(m_cari("PRIOR")), "", m_cari("PRIOR"))
        ListItem.SubItems(3) = IIf(IsNull(m_cari("NAME")), "", m_cari("NAME"))
        ListItem.SubItems(4) = IIf(IsNull(m_cari("RECSOURCE")), "", m_cari("RECSOURCE"))
        ListItem.SubItems(5) = IIf(IsNull(m_cari("NEXTACTDATE")), "", Format(m_cari("NEXTACTDATE"), "dd/mm/yyyy hh:nn"))
        ListItem.SubItems(6) = IIf(IsNull(m_cari("NEXTACT")), "", m_cari("NEXTACT"))
        ListItem.SubItems(7) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS"))
        ListItem.SubItems(8) = CStr(IIf(IsNull(m_cari("kethslkerja_new")), "", m_cari("kethslkerja_new")) & " " & sPending)
        ListItem.SubItems(9) = IIf(IsNull(m_cari("StatusCall")), "", m_cari("StatusCall"))
        ListItem.SubItems(11) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
        
       
       '  If Format(IIf(IsNull(m_cari("flaglead")), 0, m_cari("flaglead")), "##,###") = 1 Then
        '     harga = IIf(IsNull(m_cari("Principal")), 0, m_cari("Principal"))
         '    harga = harga + (harga * 26.05) / 100
          '   listitem.SubItems(12) = Format(harga, "##,###")
        'Else
        
        
        If UCase(MDIForm1.Text2) <> "SUPERVISOR" Then
                If Format(IIf(IsNull(m_cari("flaglead")), 0, m_cari("flaglead")), "##,###") = 1 Then
                       ListItem.SubItems(12) = ""
                Else
                    ListItem.SubItems(12) = Format(IIf(IsNull(m_cari("Principal")), 0, m_cari("Principal")), "##,###")
                End If
        Else
             ListItem.SubItems(12) = Format(IIf(IsNull(m_cari("Principal")), 0, m_cari("Principal")), "##,###")
        End If
        
        
        
                
       ' End If
        
        
        
        
        
        'If Format(IIf(IsNull(m_cari("flaglead")), 0, m_cari("flaglead")), "##,###") = 1 Then
         '    harga = IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo"))
          '   harga = harga + (harga * 18.26) / 100
           '  listitem.SubItems(13) = Format(harga, "##,###")
        'Else
        ListItem.SubItems(13) = Format(IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo")), "##,###")
        totamount = totamount + IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo"))
        'End If
        
        ListItem.SubItems(14) = Format(IIf(IsNull(m_cari("OpenDate")), "", m_cari("OpenDate")), "yyyy/mm/dd")
        ListItem.SubItems(15) = Format(IIf(IsNull(m_cari("B_D")), 0, m_cari("B_D")))
        ListItem.SubItems(16) = Format(IIf(IsNull(m_cari("Pay_Dt")), 0, m_cari("Pay_Dt")), "yyyy/mm/dd")
        
        '@@ 13-09-2011, Subitems 17 diganti LPA
        'listitem.SubItems(17) = IIf(IsNull(m_cari("RECSOURCE")), "", m_cari("RECSOURCE"))
        ListItem.SubItems(17) = Format(IIf(IsNull(m_cari("lastpay")), 0, m_cari("lastpay")), "##,###")
        
        ListItem.SubItems(18) = IIf(IsNull(m_cari("TGLSTATUS")), "", Format(m_cari("TGLSTATUS"), "YYYY/MM/DD"))
        ListItem.SubItems(19) = IIf(IsNull(m_cari("TGLCALL")), "", Format(m_cari("TGLCALL"), "YYYY/MM/DD"))
        ListItem.SubItems(20) = IIf(IsNull(m_cari("Kethslkerja")), "", m_cari("Kethslkerja"))
        ListItem.SubItems(21) = Format(IIf(IsNull(m_cari("TGLINCOMING")), "", m_cari("TGLINCOMING")), "YYYY/MM/DD")
        ListItem.SubItems(23) = IIf(IsNull(m_cari("resultcpa")), "", m_cari("resultcpa"))
        ListItem.SubItems(24) = IIf(IsNull(m_cari("tglinsertfrmcpa")), "", m_cari("tglinsertfrmcpa"))
        ListItem.SubItems(25) = Format(IIf(IsNull(m_cari("curbal")), "", m_cari("curbal")), "##,###")
        TOTCURBALANCE = TOTCURBALANCE + IIf(IsNull(m_cari("curbal")), 0, m_cari("curbal"))
       
        '@@16032011 Tambahan DOB dan No KTP
        ListItem.SubItems(26) = IIf(IsNull(m_cari("dob")), "", Format(m_cari("dob"), "yyyy-mm-dd"))
        ListItem.SubItems(27) = IIf(IsNull(m_cari("ktpno")), "", m_cari("ktpno"))
        
        'listitem.SubItems(21) = IIf(IsNull(m_cari("id")), "", m_cari("id"))
'         While Not m_cari2.EOF
'                Lcustid2 = CStr(IIf(IsNull(m_cari2!CustId), "", m_cari2!CustId))
'                LCall = CInt(IIf(IsNull(m_cari2!Call), 0, m_cari2!Call))
'                If Lcustid2 = Lcustid1 Then
'                   listitem.SubItems(8) = LCall
'                        Else
'                            GoTo SorryLompat
'                End If
'                     m_cari2.MoveNext
'         Wend
            
SorryLompat:
        'listitem.SubItems(19) = IIf(IsNull(m_cari("F_CEK")), "", m_cari("F_CEK"))
        VOLUMEAMOUNT = VOLUMEAMOUNT + IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo"))
        'LISTITEM.SubItems(15) = IIf(IsNull(m_cari("[NO]")), "", m_cari("[NO]"))
'            Set m_objrs = New ADODB.Recordset
'                    m_objrs.CursorLocation = adUseClient
'                    CMDSQL = "SELECT count(custid) as callInit from mgm_hst where custid ='" + LstVwSearchmgm.SelectedItem.SubItems(1) + "'  group by custid "
'                    m_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
'                    While Not m_objrs.EOF
'                    'Set listitem = LstVwSearchmgm.ListItems.ADD(, , m_cari.Bookmark)
'                        listitem.SubItems(18) = IIf(IsNull(m_objrs("callinit")), "", m_objrs("callinit"))
'                        m_objrs.MoveNext
'                        Wend
'                        'm_objrs.Close
'                        Set m_objrs = Nothing
        m_cari.MoveNext
        
    Wend
    
    '@@ 20-11-10 update f_idsesstart ke null jika nilaisession ada (tidak null)
      If NilaiSession <> "" Or IsNull(NilaiSession) = False Or NilaiSession <> Empty Then
        Dim UpdateSess As String
        
        UpdateSess = "update usertbl set f_idsessstart=null,f_pesanlockauto=null where userid='"
        UpdateSess = UpdateSess + Trim(MDIForm1.Text1.Text) + "'"
        M_OBJCONN.Execute UpdateSess
        NilaiSession = ""
      End If
    '@@ 20-11-10 update f_idsesstart ke null jika nilaisession ada (tidak null)
  
    
   txtamount.Text = Format(totamount, "##,###")
   txtcurbalance.Text = Format(TOTCURBALANCE, "##,###")
        
        If LstVwSearchMgm.ListItems.Count = 0 Then
            TxtJmlDtMgm.Text = "Tidak Ada Data"
         TxtJmlVolMgm.Text = "0"
        Else
            TxtJmlDtMgm.Text = "Total " + CStr(m_cari.RecordCount) + " Records"
            TxtJmlVolMgm.Text = "Total " + CStr(Format(VOLUMEAMOUNT, "##,###"))
        End If
LstVwSearchMgm.SortKey = 2
LstVwSearchMgm.Sorted = True
ProgressBar1.Value = 0
ProgressBar1.Visible = False
MousePointer = vbNormal
Set m_cari = Nothing
Set m_cari2 = Nothing
Exit Sub
hell:
    Me.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub HEADER_VIEW_mgm()
    LstVwSearchMgm.ColumnHeaders.ADD 1, , "No", 3 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 2, , "No CC", 15 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 3, , "Priority", 0
    LstVwSearchMgm.ColumnHeaders.ADD 4, , "Nama Customer", 15 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 5, , "Batch", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 6, , "Tgl FollowUp", 13 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 7, , "Visit", 0
    LstVwSearchMgm.ColumnHeaders.ADD 8, , "History Call", 15 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 9, , "Status Account", 13 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 10, , "Status Telp", 0
    LstVwSearchMgm.ColumnHeaders.ADD 11, , "Call Initial", 0
    LstVwSearchMgm.ColumnHeaders.ADD 12, , "Coding", 7 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 13, , "Principle", 0 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 14, , "WO Amount", 13 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 15, , "Open Date", 13 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 16, , "WO Date", 13 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 17, , "LPD", 13 * TXT
    
    'LstVwSearchMgm.ColumnHeaders.ADD 18, , "DataBase", 0
    '@@ 13-09-2011, Nomor 18 diganti dengan LPA
    LstVwSearchMgm.ColumnHeaders.ADD 18, , "LPA", 13 * TXT
    
    LstVwSearchMgm.ColumnHeaders.ADD 19, , "Tgl Status", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 20, , "Tgl Call", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 21, , "Sts Account", 0
    LstVwSearchMgm.ColumnHeaders.ADD 22, , "PTP Date", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 23, , "id", 0 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 24, , "STS", 0 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 25, , "Tanggal status CPA", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 26, , "Current Balance", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 27, , "DOB", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 28, , "KTP", 10 * TXT
End Sub

Sub WaitSecs(Seconds As Single)
 Dim a As Long
 Seconds = Seconds + Timer
 While Seconds > Timer
  a = DoEvents
 Wend
End Sub

Private Sub Command1_Click(Index As Integer)
Dim NAMACUST As String
Dim NamaAgent As String
Dim DATASOURCE As String
Dim TGLLAHIR As String
Dim OFFPHONE As String
Dim OFFPHONE2 As String
Dim HOMEPHONE As String
Dim HOMEPHONE2 As String
Dim MOBILEPHONE As String
Dim MOBILEPHONE2 As String
Dim FAXPHONE As String
Dim Lcustid As String
Dim FAXPHONE2 As String
Dim KETHSLKERJA As String
Dim lLastCallDate As String
Dim lStatusCek As String
Dim sPending As String
Dim FCEKSTATUS As String
Dim strverify As String
Dim strapprovel As String
Dim M_DATA As New CLS_FRMSEARCH
Dim M_Objrs As New ADODB.Recordset
Dim PANJANG As Integer
Dim nmagentprev As String
Dim strReject As String
Dim strSukses As String
Dim strapprovelyet As String
Dim strinject As String
Dim blokeddatamarkup As String
Dim BlokedPTPNoPayment As String
Dim STSLOCKTL As String
Dim STSfromaccount As String
'@@ 15 Agustus 2011, Bloked Payment request gaby
Dim BlokedPayment As String

Select Case Index
    Case 0
    Combo1(0).Enabled = True
    Combo1(1).Enabled = True
    F_CEK = Empty
    WO_DATE = Empty
    
    
    
    If UCase(Left(MDIForm1.Text2.Text, 5)) = "ADMIN" Or UCase(Left(MDIForm1.Text2.Text, 5)) = "SUPER" Or UCase(Left(MDIForm1.Text2.Text, 5)) = "TEAML" Then
    Else
    Call CEK_STATUS_F_CEK
    End If
'    Call UPDATE_BP
        M_Objrs.CursorLocation = adUseClient
        cmdsql = "SELECT *  FROM usertbl WHERE USERID = '" + MDIForm1.Text1.Text + "'"
        M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic

        If Not M_Objrs.EOF Then
           strinject = IIf(IsNull(M_Objrs!lockdarispv), "", M_Objrs!lockdarispv)
           If strinject = "" Then
                Bloked = ""
           Else
                Bloked = IIf(IsNull(M_Objrs!lockdarispv), "", Replace(M_Objrs!lockdarispv, "@", "'"))
           End If
           '@@140710 Bloked Entry data
           BlokedEntry = IIf(IsNull(M_Objrs!lock_entry_lpd), "", M_Objrs!lock_entry_lpd)
           blokeddatamarkup = IIf(IsNull(M_Objrs!lockmarkup), "", M_Objrs!lockmarkup)
           
           '@@15 Agustus 2011 Bloked Data Payment request gaby
           BlokedPayment = IIf(IsNull(M_Objrs!lockpayment), "", M_Objrs!lockpayment)
           
           '@@ 21 April 2014 Bloked Data PTP-NoPayment Request Joko
           BlokedPTPNoPayment = IIf(IsNull(M_Objrs!lock_ptp_payment), "", M_Objrs!lock_ptp_payment)
        End If
        
        While Not M_Objrs.EOF
                StsVl = CStr(Trim(IIf(IsNull(M_Objrs!f_VL), "", M_Objrs!f_VL)))
                StsON = CStr(Trim(IIf(IsNull(M_Objrs!f_ON), "", M_Objrs!f_ON)))
                StsOS = CStr(Trim(IIf(IsNull(M_Objrs!f_OS), "", M_Objrs!f_OS)))
                StsSK = CStr(Trim(IIf(IsNull(M_Objrs!f_SK), "", M_Objrs!f_SK)))
                StsPR = CStr(Trim(IIf(IsNull(M_Objrs!f_PR), "", M_Objrs!f_PR)))
                StsPTP = CStr(Trim(IIf(IsNull(M_Objrs!f_ptp), "", M_Objrs!f_ptp)))
                StsBP = CStr(Trim(IIf(IsNull(M_Objrs!f_bp), "", M_Objrs!f_bp)))
                StsPOP = CStr(Trim(IIf(IsNull(M_Objrs!f_pop), "", M_Objrs!f_pop)))
                StsSP = CStr(Trim(IIf(IsNull(M_Objrs!f_sp), "", M_Objrs!f_sp)))
                StsUC = CStr(Trim(IIf(IsNull(M_Objrs!F_UC), "", M_Objrs!F_UC)))
                Stsblank = CStr(Trim(IIf(IsNull(M_Objrs!F_blank), "", M_Objrs!F_blank)))
                StsWO_Date = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_DATE), "", M_Objrs!f_WO_DATE)))
                StsWO_2009 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2009), "", M_Objrs!f_WO_2009)))
                StsWO_2010 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2010), "", M_Objrs!f_WO_2010)))
                StsWO_2008 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2008), "", M_Objrs!f_WO_2008)))
                StsWO_2007 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2007), "", M_Objrs!f_WO_2007)))
                StsWO_2006 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2006), "", M_Objrs!f_WO_2006)))
                StsWO_2005 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2005), "", M_Objrs!f_WO_2005)))
                StsWO_2004 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2004), "", M_Objrs!f_WO_2004)))
                StsWO_2003 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2003), "", M_Objrs!f_WO_2003)))
                StsWO_2002 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2002), "", M_Objrs!f_WO_2002)))
                StsWO_2001 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2001), "", M_Objrs!f_WO_2001)))
                StsWO_2000 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2000), "", M_Objrs!f_WO_2000)))
                StsWO_1999 = CStr(Trim(IIf(IsNull(M_Objrs!F_WO_1999), "", M_Objrs!F_WO_1999)))
                LUserType = CStr(Trim(IIf(IsNull(M_Objrs!usertype), "", M_Objrs!usertype)))
                STSLOCKTL = Replace(CStr(Trim(IIf(IsNull(M_Objrs!lockdarispvbuattl), "", M_Objrs!lockdarispvbuattl))), "@", "'")
                STSfromaccount = CStr(Trim(IIf(IsNull(M_Objrs!fromaccount), "", M_Objrs!fromaccount)))
                
                M_Objrs.MoveNext
        Wend
        
        
       
If STSLOCKTL <> Empty Then Combo1(0).Text = "": Combo1(0).Enabled = False: Combo1(1).Enabled = False: GoTo CUY
        
        Set M_Objrs = Nothing
        StsAll = StsVl + StsPR + StsBP + StsPOP + StsUC + StsON + StsSK + StsOS

     If StsAll <> "" Then
            If LUserType = "1" Then
                    If StsUC = "UC" Then
                            If Bloked <> "" Then
                            
                                 F_CEK = "(" + Bloked + " )"
                             Else
                                 F_CEK = " substring(F_CEK_NEW,1,3)  IN('" + StsVl + "','" + StsPR + "','" + StsPTP + "','" + StsBP + "','" + StsPOP + "','" + StsSP + "','" + StsSK + "', '" + StsON + "','" + StsOS + "','') "
                            End If
                       
                    Else
                             If Bloked <> "" Then
                                    F_CEK = "(" + Bloked + " )"
                                  '  F_CEK = "(substring(F_CEK,1,3)  IN( '" + StsNa + "','" + StsOP + "','" + StsPTP + "','" + StsBP + "','" + StsPOP + "','" + StsSP + "','') OR " + bloked + " ) "
                            
                            Else
                                 '   F_CEK = "substring(F_CEK,1,3)  IN( '" + StsNa + "','" + StsOP + "','" + StsPTP + "','" + StsBP + "','" + StsPOP + "','" + StsSP + "','') "
                                  F_CEK = " substring(F_CEK_NEW,1,3)  IN('" + StsVl + "','" + StsPR + "','" + StsPTP + "','" + StsBP + "','" + StsPOP + "','" + StsSP + "','" + StsSK + "', '" + StsON + "','" + StsOS + "','') "
                            End If
                    
                    End If
                        
            End If
    Else
                
'                 If Bloked <> "" Then
'                            F_CEK = "(" + Bloked + " )"
'                          '  F_CEK = "(substring(F_CEK,1,3)  IN( '" + StsNa + "','" + StsOP + "','" + StsPTP + "','" + StsBP + "','" + StsPOP + "','" + StsSP + "','') OR " + bloked + " ) "
'
'                 End If
                 
                
     End If
        
      If StsWO_Date = "1" Then
            If LUserType = "1" Then
                WO_DATE = "substring(B_D,1,4) in ('" + StsWO_2009 + "','" + StsWO_2008 + "','" + StsWO_2007 + "','" + StsWO_2006 + "','" + StsWO_2005 + "', "
                WO_DATE = WO_DATE + "'" + StsWO_2004 + "', '" + StsWO_2003 + "', '" + StsWO_2002 + "', '" + StsWO_2001 + "','" + StsWO_2000 + "','" + StsWO_1999 + "','" + StsWO_2010 + "')"
            End If
      End If
      
      
      If UCase(MDIForm1.Text2) = "SUPERVISOR" Then
      GoTo CUY
      End If
      
      
      
      
      
        If Trim(Text1(0).Text) = Empty And Trim(Combo1(0).Text) = Empty And Combo1(2).Text = Empty And Len(TDBMask2.Value) < 1 And Len(TDBMask1.Value) < 1 And TdDob.ValueIsNull And Len(Text1(2).Text) < 3 _
        And cmbStsLastCall(0).Text = "" And CmbStatusCek.Text = "" And DtLastCall(0).ValueIsNull And CekDtDistribute.Value = 0 And Combo3.Text = "" Then
            
            
            MsgBox "Masukan Kriteria Customer Yang Akan Dicari...!!!", vbCritical + vbOKOnly, "Peringatan"
            Text1(0).SetFocus
            Set M_DATA = Nothing
            Exit Sub
            
            
        Else
CUY:
         LstVwSearchMgm.ListItems.CLEAR
         Frame3.Visible = True
          If UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
                        nmagentprev = STSLOCKTL
         End If

         If CekDtDistribute.Value = 1 Then
            NamaAgent = "AGENT is null"
         Else
                If Text1(2).Text <> Empty Then ' cek no custid
                        Lcustid = "CUSTID LIKE " + "'%" + UBAH_QUOTE(Trim(Text1(2).Text)) + "%'"
                End If
                
                If Text1(0).Text <> Empty Then ' cek nama customer
                    NAMACUST = "name LIKE " + "'%" + UBAH_QUOTE(Trim(Text1(0).Text)) + "%'"
                End If
                
                If Combo1(0).Text <> Empty Then
                    If UCase(Left(MDIForm1.Text1.Text, 2)) = "TL" And Combo1(0).Text = Trim(MDIForm1.Text1.Text) Then
                        NamaAgent = "AGENT in(select userid from usertbl where team='" + Trim(Combo1(0).Text) + "')"
                    ElseIf UCase(Left(MDIForm1.Text1.Text, 4)) = "JOKO  " And Trim(Combo1(0).Text) = Trim(MDIForm1.Text1.Text) Then
                           NamaAgent = "AGENT in(select userid from usertbl where spvcode='" + Trim(Combo1(0).Text) + "')"
                    Else
                        NamaAgent = "AGENT = '" + Trim(Combo1(0).Text) + "'"
                    End If
                End If
                
                
                If Combo1(2).Text <> Empty Then
                    DATASOURCE = "RECSOURCE = '" + Trim(Combo1(2).Text) + "'"
                End If
                
                If TdDob.ValueIsNull Then
                Else
                    TGLLAHIR = "DOB = '" + Format(TdDob.Text, "yyyy/mm/dd") + "'"
                End If
                If Len(TDBMask1.Value) > 1 Then
                    OFFPHONE = "OFFICENO Like '%" + TDBMask1.Value + "%'"
                    OFFPHONE2 = "OFFICENO2 Like '%" + TDBMask1.Value + "%'"
                    HOMEPHONE = "HOMENO Like '%" + TDBMask1.Value + "%'"
                    HOMEPHONE2 = "HOMENO2 Like '%" + TDBMask1.Value + "%'"
                    FAXPHONE = "FAXNO Like '%" + TDBMask1.Value + "%'"
                    FAXPHONE2 = "FAXNO2 Like '%" + TDBMask1.Value + "%'"
                End If
                
                If Len(TDBMask2.Value) > 1 Then
                    MOBILEPHONE = "MOBILENO like '%" + TDBMask2.Value + "%'"
                    MOBILEPHONE2 = "MOBILENO2 like '%" + TDBMask2.Value + "%'"
                End If
                
                
                If Left(Combo3.Text, 3) = "ALL" Then
                    strverify = " stscpa=1"
                End If
                
                
                
                
                If cmbStsLastCall(0).Text <> "" Then
                    Select Case UCase(Trim(cmbStsLastCall(0).Text))
                        Case "NOT CONTACTED"
                            KETHSLKERJA = "KETHSLKERJA IN('WN','WN','NP','BT')"
                        Case "NOT AVAILABLE"
                            KETHSLKERJA = "KETHSLKERJA = 'NA'"
                        Case "STILL THINKING"
                            KETHSLKERJA = "KETHSLKERJA= 'ST'"
                        Case "DISAGREE"
                            KETHSLKERJA = "LEFT(KETHSLKERJA,1)= 'D'"
                        Case "SENDING APPLICATION"
                            KETHSLKERJA = "KETHSLKERJA= 'SA'"
                        Case "PICKUP"
                            KETHSLKERJA = "KETHSLKERJA= 'PU'"
                        Case "INCOMPLETE DOCUMENT"
                            KETHSLKERJA = "KETHSLKERJA= 'ID'"
                        Case "INCOMING"
                            KETHSLKERJA = "KETHSLKERJA= 'I'"
                        Case "AVAILABLE"
                            KETHSLKERJA = "KETHSLKERJA= '1A'"
                    End Select
                End If
                If Len(Combo2) > 0 Then
                    '@@ 09-04-2012, diganti nih jika berdasarkan hot prospect
                    If Combo2.Text = "HOT PROSPECT" Then
                        FCEKSTATUS = " status_htc='1' "
                    Else
                        FCEKSTATUS = " F_CEK_NEW LIKE '" + Combo2.Text + "%'"
                    End If
                    
                End If
                If CmbStatusCek.Text <> "" Then
                    Select Case UCase(Trim(CmbStatusCek.Text))
                        Case "ACCEPT"
                            lStatusCek = "UPPER(F_cek_new) ='ACCEPT'"
                        Case "RETURN"
                            lStatusCek = "UPPER(F_cek_new) = 'RETURN'"
                        Case "NOT CHECK"
                            lStatusCek = "(f_cek_NEW is null or f_cek = 'NOT ACCEPT' OR F_CEK_NEW ='') AND KETHSLKERJA ='I'"
                    End Select
                End If
                If DtLastCall(0).ValueIsNull Then
                Else
                    lLastCallDate = "TGLSTATUS BETWEEN '" + Format(DtLastCall(0).Value, "MM/DD/YYYY") & " " & CStr(DTimeLastCall(0).Value) + "' AND '" + Format(DtLastCall(1).Value, "MM/DD/YYYY") & " " & CStr(DTimeLastCall(1).Value) + "'"
                End If
        End If
        End If
          
                'Unload FRM_SEARCH
                If Check1.Value = 0 Then
                    Set m_cari = M_DATA.QUERY_SEARCH_CONDITION(M_OBJCONN, NAMACUST, NamaAgent, DATASOURCE, TGLLAHIR, _
                                                            OFFPHONE, OFFPHONE2, HOMEPHONE, HOMEPHONE2, MOBILEPHONE, _
                                                            MOBILEPHONE2, FAXPHONE, FAXPHONE2, MDIForm1.Text3.Text, Lcustid, F_CEK, lLastCallDate, lStatusCek)
                Else
                If blokeddatamarkup <> "" Then
                   ' F_CEK = ""
                   ' WO_DATE = ""
                   ' BlokedEntry = ""
                   ' Bloked = ""
                End If
                
'                Set m_cari = M_DATA.QUERY_SEARCH_CONDITION_mgm(M_OBJCONN, NAMACUST, NAMAAGENT, DATASOURCE, TGLLAHIR, _
'                                                             OFFPHONE, OFFPHONE2, HOMEPHONE, HOMEPHONE2, MOBILEPHONE, _
'                                                            MOBILEPHONE2, FAXPHONE, FAXPHONE2, MDIForm1.Text3.Text, Lcustid, F_CEK, lLastCallDate, lStatusCek, sPending, FCEKSTATUS, WO_DATE, strverify, strapprovel, strapprovelyet, strReject, strSukses, Bloked, BlokedEntry, blokeddatamarkup, nmagentprev, "")
'
                ' TAMBAHAN BlokedPTPNoPayment 21 APRIL 2014
                Set m_cari = M_DATA.QUERY_SEARCH_CONDITION_mgm(M_OBJCONN, NAMACUST, NamaAgent, DATASOURCE, TGLLAHIR, _
                                                             OFFPHONE, OFFPHONE2, HOMEPHONE, HOMEPHONE2, MOBILEPHONE, _
                                                            MOBILEPHONE2, FAXPHONE, FAXPHONE2, MDIForm1.Text3.Text, Lcustid, Bloked, lLastCallDate, lStatusCek, sPending, FCEKSTATUS, WO_DATE, strverify, strapprovel, strapprovelyet, strReject, strSukses, Bloked, BlokedEntry, blokeddatamarkup, nmagentprev, "", BlokedPayment, BlokedPTPNoPayment)

                End If
            LBLCOUNT.Caption = "Jumlah Record :" + CStr(m_cari.RecordCount)
            If m_cari.RecordCount = 0 Then
                MsgBox "Data Anda Tidak Ditemukan! Sistem akan coba mencari data yang dapat di collect bersama!", vbInformation + vbOKOnly, "Aplikasi"
                
                'Jika data tidak ditemukan maka jika idsessstart maka di null-in lagi idessstarnya
                If Text1(2).Text = Empty _
                  And TDBMask1.Value = Empty _
                  And Text1(0).Text = Empty _
                  And Combo2.Text = Empty _
                  And IsNull(DtLastCall(0).Value) _
                  And IsNull(DtLastCall(1).Value) _
                  And Combo1(2).Text = Empty _
                  And Combo1(3).Text = Empty _
                  And Combo3.Text = Empty _
                  And Combo1(0).Text = Trim(MDIForm1.Text1.Text) Then
                    
                    Dim m_objrsCekSess As ADODB.Recordset
                    Dim CmdCekSess As String
                    Dim idsess As String
                    
                    
                    
                    CmdCekSess = "select f_idsessstart from usertbl where userid='"
                    CmdCekSess = CmdCekSess + Trim(MDIForm1.Text1.Text) + "'"
                    Set m_objrsCekSess = New ADODB.Recordset
                    m_objrsCekSess.CursorLocation = adUseClient
                    m_objrsCekSess.Open CmdCekSess, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                    'idsess = IIf(IsNull(CStr(m_objrsCekSess("f_idsessstart"))), "", CStr(m_objrsCekSess("f_idsessstart")))
                    'idsess = CStr(m_objrsCekSess("f_idsessstart"))
                    If IsNull(m_objrsCekSess("f_idsessstart")) Then
                        idsess = ""
                    Else
                        idsess = CStr(m_objrsCekSess("f_idsessstart"))
                    End If
                    Set m_objrsCekSess = Nothing
                    
                    If idsess <> "" Or idsess <> Empty Then
                        'Kasih informasi ke agent kenapa datanya kosong
                        Dim CmdSql_Info As String
                        Dim M_ObjrsInfo As ADODB.Recordset
                        Dim Pesan As String
                        
                        
                        CmdSql_Info = "select * from tbltemplockacc_current where id='"
                        CmdSql_Info = CmdSql_Info + Trim(idsess) + "'"
                        Set M_ObjrsInfo = New ADODB.Recordset
                        M_ObjrsInfo.CursorLocation = adUseClient
                        M_ObjrsInfo.Open CmdSql_Info, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                        
                        Pesan = "Kondisi Lock data anda saat ini adalah:" + Chr(13)
                        Pesan = Pesan + "Start Lock: " + CStr(M_ObjrsInfo("start_lock")) + Chr(13)
                        Pesan = Pesan + "End Lock: " + CStr(M_ObjrsInfo("end_lock")) + Chr(13)
                        Pesan = Pesan + "Account yang di lock: " + M_ObjrsInfo("account_lock") + Chr(13)
                        Pesan = Pesan + "Di lock oleh: " + M_ObjrsInfo("lock_by") + Chr(13)
                        Pesan = Pesan + "Status yang di lock: " + M_ObjrsInfo("status_lock") + Chr(13)
                        Pesan = Pesan + Chr(13)
                        Pesan = Pesan + "System tidak dapat menemukan data sesuai lock anda, " + Chr(13)
                        Pesan = Pesan + "anda dapat menghubungi TL/SPV/Orang yang melock data anda(lihat lock by) " + Chr(13)
                        Pesan = Pesan + "untuk merelease data anda! Terima kasih."
                        
                        MsgBox Pesan, vbOKOnly + vbInformation, "Informasi"
                        Set M_ObjrsInfo = Nothing
                        
                        CmdCekSess = "update usertbl set f_idsessstart=null,f_idsessend=null where userid='"
                        CmdCekSess = CmdCekSess + Trim(MDIForm1.Text1.Text) + "'"
                        M_OBJCONN.Execute CmdCekSess
                        
                        '@@20022013 tambahan jika data tidak di temukan
                    End If
                    
                End If
                  
                Set M_DATA = Nothing
                
                Call CariDataAll
                Exit Sub
            Else
               
                search_ok = True
                If Check1.Value = 1 Then
                    'kalau found refferall data
                    'Unload FRM_PRESCREEN
                    'FRM_PRESCREEN.Caption = "Search Non mgm Data"
                    'FRM_PRESCREEN.Show
                    SSTab1.Tab = 0
'                    Call show_UCDATA
                    ' Untuk ngecek apakah data sudah 5 hari belum dikerjakan juga // By Izuddin
                    Call cek_lama_account
                    ' -------------------------------------------------------------------------
                    Call show_Search_mgmData
                    
                Else
                    ' kalau found mgm data
                    SSTab1.Tab = 1
'Call show_Search_Refferal
                    
'                    Unload VIEW_mgmDATA
'                    If MDIForm1.Text2.Text = "Agent" Then
'                        VIEW_mgmDATA.Caption = "Search mgm Data"
'                    Else
'                        VIEW_mgmDATA.Caption = "Search mgm Data  .... Tekan Huruf ""P"" untuk Mengalihkan Data"
'                    End If
'
'
'                    VIEW_mgmDATA.Show
                End If
            'FRM_PRESCREEN.Show vbModal
'                Unload Me
                    
            '@@12022013, buat akses akun yang bareng2 khusus TL dan agent
            If UCase(MDIForm1.Text2.Text) = "AGENT" Or UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
                Call CariDataAll
            End If
        End If
    Case 1
        MDIForm1.LstGrade.ListItems.CLEAR
        Unload Me
        
    Case 2
        Text1(2).Text = Empty
        Text1(0).Text = Empty
        TdDob.Text = Empty
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
        Combo1(2).Text = Empty
        Combo1(3).Text = Empty
        TDBMask1.Text = Empty
        TDBMask2.Text = Empty
        cmbStsLastCall(0).Text = Empty
        DtLastCall(0).Value = Empty
        DtLastCall(1).Value = Empty
        CmbStatusCek.Text = Empty
        
End Select
Set M_DATA = Nothing

' Frame3.Visible = False
End Sub
Sub cek_lama_account()
    Dim rs_cek As ADODB.Recordset
    Dim tglserver As Date
    Dim interval As Integer
    Dim interval2 As Integer
    Dim cek_available As Integer
    Dim TL_Review As String
    
    If UCase(Left(MDIForm1.Text2.Text, 5)) = "AGENT" Then
        cek_available = 0
        Set rs_cek = New ADODB.Recordset
        rs_cek.CursorLocation = adUseClient
        rs_cek.ActiveConnection = M_OBJCONN
        rs_cek.CursorType = adOpenDynamic
        rs_cek.LockType = adLockOptimistic
        
        If rs_cek.state = 1 Then rs_cek.Close
        rs_cek.Open "SELECT userid FROM usertbl WHERE lower(userid) LIKE 'review%' AND team IN (SELECT team FROM usertbl WHERE userid='" & Trim(MDIForm1.Text1.Text) & "')"
        TL_Review = IIf(IsNull(rs_cek!Userid), "", rs_cek!Userid)
        
        If rs_cek.state = 1 Then rs_cek.Close
        rs_cek.Open "SELECT now() as tgl_server"
        tglserver = Format(rs_cek!tgl_server, "yyyy-mm-dd")
        
        If rs_cek.state = 1 Then rs_cek.Close
        rs_cek.Open "SELECT id,tglsource,tglcall,spv_allow FROM mgm WHERE tglcall is null AND spv_allow is null AND agent='" & Trim(MDIForm1.Text1.Text) & "'"
        If rs_cek.RecordCount > 0 Then
            Do Until rs_cek.EOF
                interval = DateDiff("d", Format(rs_cek!TGLSOURCE, "yyyy-mm-dd"), Format(tglserver, "yyyy-mm-dd"))
                ' Jika kelewat 5 hari dari tgl upload
                If interval > 5 Then
                    cek_available = cek_available + 1
                    M_OBJCONN.Execute "UPDATE mgm SET agent_asli=agent WHERE id=" & rs_cek!ID & ""
                    M_OBJCONN.Execute "UPDATE mgm SET agent='" & TL_Review & "' WHERE id=" & rs_cek!ID & ""
                End If
                rs_cek.MoveNext
            Loop
            
            If cek_available > 0 Then
                'M_OBJCONN.Execute "UPDATE usertbl SET f_blok='1' WHERE userid='" & Trim(MDIForm1.Text1.Text) & "'"
                MsgBox cek_available & " Data(s) Masuk ke coding TL REVIEW karena Data lebih dari 5 hari belum dikerjakan", vbCritical + vbOKOnly, "Akun data 5 hari"
                'End
            End If
        End If
    End If
End Sub

Sub UPDATE_BP()
'Dim TGLSYS As ADODB.Recordset
'Dim SPTP As ADODB.Recordset
'Dim CMDSQL As String
'Dim TGLKOMP As Date
'Set TGLSYS = New ADODB.Recordset
'TGLSYS.CursorLocation = adUseClient
'TGLSYS.Open "SELECT TGLSYSTEM FrOM VWCALLCFG1", M_OBJCONN, adOpenDynamic, adLockOptimistic
'TGLKOMP = TGLSYS!TGLSYSTEM
'Set SPTP = New ADODB.Recordset
'CMDSQL = "UPDATE mgm SET KETHSLKErJA='BP-BROKEN PROMISE',F_CEK='BP-',KETHSLKErJADESC='' "
'CMDSQL = CMDSQL + "WHERE CUSTID IN (SELECT CUSTID FrOM mgm WHErE DATEDIFF(day,TdbdatePTP,TGLKOMP)>5)"

End Sub

Private Sub show_UCDATA()
Dim sdata As ADODB.Recordset
Dim ListItem As ListItem
Dim I%
Set sdata = New ADODB.Recordset
sdata.CursorLocation = adUseClient
cmdsql = "select CUSTID,KETHSLKERJA,AGENT FROM mgm WHERE left(KETHSLKERJA,2) in ('WN','NK','MV') AND AGENT='" & VIEW_MGMDATA.Combo1(0).Text & "'"
sdata.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
I = 1
MDIForm1.LstGrade.ListItems.CLEAR
Do While Not sdata.EOF
    Set ListItem = MDIForm1.LstGrade.ListItems.ADD(, , I)
        ListItem.SubItems(1) = IIf(IsNull(sdata!CustId), "", sdata!CustId)
        ListItem.SubItems(2) = IIf(IsNull(sdata!CustId), "", sdata!KETHSLKERJA)
        ListItem.SubItems(3) = IIf(IsNull(sdata!agent), "", sdata!agent)
        sdata.MoveNext
        I = I + 1
Loop
End Sub
'Private Sub show_Search_Refferal()
'Dim listitem As listitem
'ProgressBar1.Max = m_cari.RecordCount + 1
'Text2.Text = m_cari.RecordCount & " Data"
'ListView1.ListItems.Clear
'
'While Not m_cari.EOF
'ProgressBar1.Value = m_cari.Bookmark
'Set listitem = ListView1.ListItems.ADD(, , m_cari.Bookmark)
'    listitem.SubItems(1) = IIf(IsNull(m_cari("custid")), "", m_cari("custid"))
'    Select Case m_cari("RECSTATUS")
'    Case "1A"
'        listitem.SubItems(2) = "Available"
'    Case ""
'        listitem.SubItems(2) = "Available"
'    Case Else
'        listitem.SubItems(2) = IIf(IsNull(m_cari("PRIOR")), "", m_cari("PRIOR"))
'    End Select
'    listitem.SubItems(3) = IIf(IsNull(m_cari("CUSTIDREF")), "", m_cari("CUSTIDREF"))
'    listitem.SubItems(4) = IIf(IsNull(m_cari("NAMAREF")), "", m_cari("NAMAREF"))
'    listitem.SubItems(5) = IIf(IsNull(m_cari("NAME")), "", m_cari("NAME"))
'    listitem.SubItems(6) = IIf(IsNull(m_cari("NEXTACTDATE")), "", Format(m_cari("NEXTACTDATE"), "yyyy/mm/dd hh:mm"))
'    listitem.SubItems(7) = IIf(IsNull(m_cari("NEXTACT")), "", m_cari("NEXTACT"))
'    listitem.SubItems(8) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS"))
'    listitem.SubItems(9) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
'    listitem.SubItems(10) = IIf(IsNull(m_cari("NamaAGENT")), "", m_cari("NamaAGENT"))
'    listitem.SubItems(11) = IIf(IsNull(m_cari("RECSOURCEREF")), "", m_cari("RECSOURCEREF"))
'    listitem.SubItems(12) = Format(IIf(IsNull(m_cari("TGLSTATUS")), "", m_cari("TGLSTATUS")), "yyyy/mm/dd")
'    listitem.SubItems(13) = IIf(IsNull(m_cari("Kethslkerja")), "", m_cari("Kethslkerja"))
'    listitem.SubItems(14) = IIf(IsNull(m_cari("KdComplaint")), "", m_cari("KdComplaint"))
'    listitem.SubItems(15) = IIf(IsNull(m_cari("RemarkComplaint")), "", m_cari("RemarkComplaint"))
'    listitem.SubItems(16) = IIf(IsNull(m_cari("F_CEK")), "", m_cari("F_CEK"))
'    listitem.SubItems(17) = IIf(IsNull(m_cari("Nomor")), "", m_cari("Nomor"))
'    m_cari.MoveNext
'Wend
'Set m_cari = Nothing
'End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
MDIForm1.m_targetview = False

End Sub



Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
   ListView1.SortKey = ColumnHeader.Index - 1
   ListView1.Sorted = True
End Sub

Private Sub ListView1_DblClick()
Dim M_Objrs As ADODB.Recordset
If ListView1.ListItems.Count = 0 Then
    Exit Sub
End If
Status_Form = 1
    If UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        M_Objrs.Open "SELECT USERID FROM usertbl WHERE SPVCODE ='" + MDIForm1.Text1.Text + "' AND USERID = '" + ListView1.SelectedItem.SubItems(9) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_Objrs.RecordCount <> 0 Then
        Else
            MsgBox "Data Ini Milik Agent Team Leader Yang Lain", vbCritical + vbOKOnly, "Aplikasi"
            Set M_Objrs = Nothing
            Exit Sub
        End If
        Set M_Objrs = Nothing
    Else
        If UCase(MDIForm1.Text2.Text) = "SUPERVISOR" Or UCase(MDIForm1.Text2.Text) = "ADMINISTRATOR" Then
        Else
            If Trim(UCase(MDIForm1.Text1.Text)) = Trim(UCase(ListView1.SelectedItem.SubItems(9))) Then
            Else
                MsgBox "Data Ini Milik Agent Yang Lain", vbCritical + vbOKOnly, "Aplikasi"
                Set M_Objrs = Nothing
                Exit Sub
            End If
        End If
    End If
    FrmCC_Colection.Show vbModal
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
Dim M_Objrs As ADODB.Recordset
Dim m_msgbox As Variant
If ListView1.ListItems.Count = 0 Then
    Exit Sub
End If
If KeyAscii = 13 Then
    Call ListView1_DblClick
End If
If UCase(MDIForm1.Text2.Text) <> "AGENT" Then
    If KeyAscii = 100 Or KeyAscii = 68 Then
        If UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
            Set M_Objrs = New ADODB.Recordset
            M_Objrs.CursorLocation = adUseClient
            M_Objrs.Open "SELECT USERID FROM usertbl WHERE SPVCODE ='" + MDIForm1.Text1.Text + "' AND USERID = '" + ListView1.SelectedItem.SubItems(9) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If M_Objrs.RecordCount <> 0 Then
                m_msgbox = MsgBox("Yakin Akan Di Hapus??...", vbExclamation + vbYesNo, "Informasi")
                If m_msgbox = vbYes Then
                    M_OBJCONN.Execute "Delete From cc_custtbl where Nomor = " + ListView1.SelectedItem.SubItems(17) + " And RECSOURCE <> 'infomedia1' and left(recsource,3)<>'inf'"
                    ListView1.ListItems.Remove ListView1.SelectedItem.Index
                End If
            Else
                MsgBox "Data Ini Milik Agent Team Leader Yang Lain", vbCritical + vbOKOnly, "Aplikasi"
                Set M_Objrs = Nothing
                Exit Sub
            End If
            Set M_Objrs = Nothing
        Else
            m_msgbox = MsgBox("Yakin Akan Di Hapus??...", vbExclamation + vbYesNo, "Informasi")
            If m_msgbox = vbYes Then
                'M_OBJCONN.Execute "Delete From cc_custtbl where custid ='" + ListView1.SelectedItem.SubItems(1) + "'"
                M_OBJCONN.Execute "Delete From cc_custtbl where Nomor = " + ListView1.SelectedItem.SubItems(17) + " And RECSOURCE <> 'infomedia1'"
                ListView1.ListItems.Remove ListView1.SelectedItem.Index
            End If
        End If
    End If
End If
End Sub

Private Sub LstVwSearchmgm_Click()
If b_pindah = True Then
    FrmCustIdTransfer.List1.AddItem LstVwSearchMgm.SelectedItem.SubItems(1)
End If
End Sub

Private Sub LstVwSearchmgm_ColumnClick(ByVal ColumnHeader As ColumnHeader)
   LstVwSearchMgm.SortKey = ColumnHeader.Index - 1
   LstVwSearchMgm.Sorted = True
End Sub
Private Sub LstVwSearchmgm_DblClick()
    Dim Strsql  As String
    Dim MOBJRSKISRUT As New ADODB.Recordset
    'On Error GoTo ke
    
    Dim M_Objrs As ADODB.Recordset
    If LstVwSearchMgm.ListItems.Count = 0 Then
        Exit Sub
    End If
    glexp = LstVwSearchMgm.SelectedItem.SubItems(4)
    Status_Form = 2
    If LstVwSearchMgm.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    '@@12022013 ini jika statusnya AKSESALL
    If Trim(UCase(LstVwSearchMgm.SelectedItem.SubItems(11))) = "AKSESALL" Then
        If UCase(MDIForm1.Text2.Text) = "AGENT" Or _
           UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
            Dim CmdsqlCek As String
            Dim M_ObjrsCekAkses As ADODB.Recordset
            
'            CmdsqlCek = "select * from tbl_distribusi_account where custid='"
'            CmdsqlCek = CmdsqlCek & CStr(LstVwSearchMgm.SelectedItem.SubItems(1)) & "' and agent='"
'            CmdsqlCek = CmdsqlCek & MDIForm1.Text1.Text & "' and waktu_akhir > now() "
            CmdsqlCek = "select * from tbl_profile_aksesall a,tbl_cust_aksesall b WHERE a.kd_profile=b.kd_profile AND b.custid='"
            CmdsqlCek = CmdsqlCek & CStr(LstVwSearchMgm.SelectedItem.SubItems(1)) & "' AND a.waktu_akhir > now() "
            Set M_ObjrsCekAkses = New ADODB.Recordset
            M_ObjrsCekAkses.CursorLocation = adUseClient
            M_ObjrsCekAkses.Open CmdsqlCek, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            
            If M_ObjrsCekAkses.RecordCount > 0 Then
                'Ini jika waktunya masih berlaku
                Set M_ObjrsCekAkses = Nothing
                
                'Cek dulu apakah lagi diakses oleh agent yang lain
                Dim M_Objrs_Akses_Acc As ADODB.Recordset
                cmdsql = "select monitor_akses,waktu_akses from mgm where custid='"
                cmdsql = cmdsql & CStr(LstVwSearchMgm.SelectedItem.SubItems(1)) & "'"
                Set M_Objrs_Akses_Acc = New ADODB.Recordset
                M_Objrs_Akses_Acc.CursorLocation = adUseClient
                M_Objrs_Akses_Acc.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                If M_Objrs_Akses_Acc.RecordCount > 0 Then
                    If IsNull(M_Objrs_Akses_Acc("monitor_akses")) = False Or _
                       M_Objrs_Akses_Acc("monitor_akses") <> "" Then
                        MsgBox "Mohon maaf! Account ini sedang diakses oleh agent lain. " & M_Objrs_Akses_Acc("monitor_akses") & " " & M_Objrs_Akses_Acc("waktu_akses") & ". Cobalah akses di lain waktu, atau hubungi SPV untuk membuka Account ini!", vbOKOnly + vbInformation, "Informasi"
                        Set M_Objrs_Akses_Acc = Nothing
                        Exit Sub
                    End If
                End If
                
                Set M_Objrs_Akses_Acc = Nothing
                
                '@@13022013 Tandai dulu deh biar ga diakses oleh yang lain
                cmdsql = "update mgm set monitor_akses='AKSES OLEH "
                cmdsql = cmdsql & MDIForm1.Text1.Text & "',waktu_akses=now() where custid='"
                cmdsql = cmdsql & CStr(LstVwSearchMgm.SelectedItem.SubItems(1)) & "'"
                M_OBJCONN.Execute cmdsql
                GoTo KE
            Else
                'Ini jika waktunya sudah tidak berlaku
                Set M_ObjrsCekAkses = Nothing
               
                ' UPDATE 03 JUNI 2014 BY IZUDDIN
                cmdsql = "update mgm set agent=agent_asli WHERE monitor_akses is null" & _
                        " AND agent='AKSESALL' AND custid in(SELECT custid FROM tbl_cust_aksesall a,tbl_profile_aksesall b WHERE " & _
                        " a.kd_profile=b.kd_profile AND b.waktu_akhir < now()) AND agent_asli is not null"
                M_OBJCONN.Execute cmdsql
               
                ' UPDATE 02 JULI 2013 BY IZUDDIN
                cmdsql = "update mgm set agent_asli=null WHERE monitor_akses is null" & _
                        " AND agent='AKSESALL' AND custid in(SELECT custid FROM tbl_cust_aksesall a,tbl_profile_aksesall b WHERE " & _
                        " a.kd_profile=b.kd_profile AND b.waktu_akhir < now()) AND agent_asli is not null"
                M_OBJCONN.Execute cmdsql
                
                cmdsql = "DELETE FROM tbl_cust_aksesall "
                cmdsql = cmdsql & " WHERE kd_profile in (SELECT kd_profile FROM tbl_profile_aksesall WHERE waktu_akhir < now()) "
                M_OBJCONN.Execute cmdsql
                
                AksesAllAcc = ""
                
                MsgBox "Mohon maaf! Waktu akses untuk account ini bagi anda sudah habis! Data anda akan diperbaharui!", vbOKOnly + vbInformation, "Informasi"
                Command1_Click (0)
                Exit Sub
            End If
        End If
    End If

    If UCase(MDIForm1.Text2.Text) = "AGENT" Then
        If UCase(MDIForm1.Text1.Text) <> Trim(UCase(LstVwSearchMgm.SelectedItem.SubItems(11))) Then
            'Dim Cmdsql As String
            Dim M_Objrs_Cek As ADODB.Recordset
            Dim Vcek As Boolean
            
          
            
            '@@16032011 Tambahan jika CH tersebut memiliki data Visa, tapi punya agent lain tetep bisa dibuka sama agent tsb
            'Cek dulu punya data no.ktp apa ngga
            If LstVwSearchMgm.SelectedItem.SubItems(27) <> "" Then
                cmdsql = "select custid,agent from mgm where (name='"
                cmdsql = cmdsql + Trim(LstVwSearchMgm.SelectedItem.SubItems(3)) + "' and dob='"
                cmdsql = cmdsql + Format(LstVwSearchMgm.SelectedItem.SubItems(26), "yyyy-mm-dd") + "' or ktpno='"
                cmdsql = cmdsql + Trim(LstVwSearchMgm.SelectedItem.SubItems(27)) + "') and custid<>'"
                cmdsql = cmdsql + Trim(LstVwSearchMgm.SelectedItem.SubItems(1)) + "'"
            Else
                cmdsql = "select custid,agent from mgm where name='"
                cmdsql = cmdsql + Trim(LstVwSearchMgm.SelectedItem.SubItems(3)) + "' and dob='"
                cmdsql = cmdsql + Format(LstVwSearchMgm.SelectedItem.SubItems(26), "yyyy-mm-dd") + "' and custid <>'"
                cmdsql = cmdsql + Trim(LstVwSearchMgm.SelectedItem.SubItems(1)) + "'"
            End If
            Set M_Objrs_Cek = New ADODB.Recordset
            M_Objrs_Cek.CursorLocation = adUseClient
            M_Objrs_Cek.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            Vcek = False
            While Not M_Objrs_Cek.EOF
                If UCase(MDIForm1.Text1.Text) = UCase(Trim(M_Objrs_Cek("agent"))) Then
                    Vcek = True
                End If
                M_Objrs_Cek.MoveNext
            Wend
            Set M_Objrs_Cek = Nothing
        
            '@@02082012 Cek Coding nih......
            cmdsql = "select * from "
            cmdsql = cmdsql + "(select spvcode from usertbl where userid='"
            cmdsql = cmdsql + CStr(Trim(MDIForm1.Text1.Text))
            cmdsql = cmdsql + "') as a, "
            cmdsql = cmdsql + " (select spvcode as spvcode_new,sts_akses_agent as sts_akses_agent_new "
            cmdsql = cmdsql + " from usertbl where userid='"
            cmdsql = cmdsql + CStr(Trim(LstVwSearchMgm.SelectedItem.SubItems(11)))
            cmdsql = cmdsql + "') as b "
            cmdsql = cmdsql + " where a.SPVCODE = b.spvcode_new "
            Set M_Objrs_Cek = New ADODB.Recordset
            M_Objrs_Cek.CursorLocation = adUseClient
            M_Objrs_Cek.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If M_Objrs_Cek.RecordCount > 0 Then
                If IsNull(M_Objrs_Cek("sts_akses_agent_new")) = True Then
                    Vcek = False
                ElseIf CStr(Trim(M_Objrs_Cek("sts_akses_agent_new"))) = "1" Then
                    Vcek = True
                End If
            End If
            Set M_Objrs_Cek = Nothing
        
            If Vcek = False Then
                MsgBox "Anda Tidak Berhak Untuk Mengedit Data Ini", vbCritical + vbOKOnly, "Aplikasi"
                Exit Sub
            Else
                GoTo lanjut
            End If
        End If
    End If
lanjut:
    If UCase(MDIForm1.Text2.Text) = "SUPERVISOR" Then
      'If UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
        Dim PO_AGENT As String
        If VIEW_MGMDATA.Combo1(0).Text = "PULLOUT" Then
            Set M_Objrs = New ADODB.Recordset
            M_Objrs.CursorLocation = adUseClient
            cmdsql = "SELECT PO_Agent FROM mgm where CUSTID='" & LstVwSearchMgm.SelectedItem.SubItems(11) & "'"
            M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If Not M_Objrs.EOF Then
            PO_AGENT = M_Objrs!PO_AGENT
            
            End If
            Set M_Objrs = Nothing
        Else
            PO_AGENT = VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11)
        End If
    Else
        
'        If UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
'
'            STRSQL = "SELECT * FROM USERTBL WHERE  USERID IN (SELECT  agentprev FROM MGM WHERE CUSTID ='" + Trim(LstVwSearchMgm.SelectedItem.SubItems(1)) + "') AND TEAM ='" + MDIForm1.Text1.Text + "'"
'            Set MOBJRSKISRUT = New ADODB.Recordset
'            MOBJRSKISRUT.CursorLocation = adUseClient
'            MOBJRSKISRUT.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
'            If MOBJRSKISRUT.RecordCount = 0 Then
'                MsgBox "Data Ini Milik Agent Team Leader Yang Lain", vbCritical + vbOKOnly, "Aplikasi"
'                Exit Sub
'            End If
'
'        End If
  '      Set MOBJRSKISRUT = Nothing
        Set M_Objrs = New ADODB.Recordset
        M_Objrs.CursorLocation = adUseClient
        If UCase(Trim(MDIForm1.Text2)) = "SUPERVISOR" Or UCase(Trim(MDIForm1.Text2)) = "ADMINISTRATOR" Then
            cmdsql = "SELECT USERID FROM usertbl WHERE  USERID = '" + Trim(LstVwSearchMgm.SelectedItem.SubItems(11)) + "'"
                ElseIf UCase(MDIForm1.Text2) = "TEAMLEADER" Then
                    cmdsql = "SELECT USERID FROM usertbl WHERE TEAM ='" + MDIForm1.Text1.Text + "' AND USERID = '" + Trim(LstVwSearchMgm.SelectedItem.SubItems(11)) + "'"
                ElseIf UCase(MDIForm1.Text2) = "ADMIN" Then
                cmdsql = "SELECT USERID FROM usertbl WHERE  USERID = '" + Trim(LstVwSearchMgm.SelectedItem.SubItems(11)) + "'"
        End If
        '@@ 19 Juli 2010 .. Ini pengalihan error buka data oleh agent
        On Error GoTo SALAH
        M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
selanjutnya:
        If M_Objrs.RecordCount <> 0 Then
        Else
        Strsql = "SELECT * FROM USERTBL WHERE  USERID IN (SELECT  agentprev FROM MGM WHERE CUSTID ='" + Trim(LstVwSearchMgm.SelectedItem.SubItems(1)) + "') AND TEAM ='" + MDIForm1.Text1.Text + "'"
            Set MOBJRSKISRUT = New ADODB.Recordset
            MOBJRSKISRUT.CursorLocation = adUseClient
            MOBJRSKISRUT.Open Strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
            If MOBJRSKISRUT.RecordCount > 0 Then
                GoTo KE
            End If
            
                MsgBox "Data Ini Milik Agent Team Leader Yang Lain", vbCritical + vbOKOnly, "Aplikasi"
            Set MOBJRSKISRUT = Nothing
            Set M_Objrs = Nothing
               Exit Sub
        End If
        Set M_Objrs = Nothing
    End If
KE:
    Me.MousePointer = vbHourglass
    Flag_mgm = False
    'Matikan main timer activity By Izuddin 16042013
    main_timer_activity = 0
    MDIForm1.Timer7.Enabled = False
    ' ##############################################
'    FrmCC_Colection.Show vbModal
    'SET WAKTU LOGOUT
    M_OBJCONN.Execute "UPDATE usertbl SET last_logout='now()' WHERE userid='" + Trim(MDIForm1.Text1.Text) + "'"
    '----------------
    FrmCC_Colection.Show vbModal
    
    If LstVwSearchMgm.ListItems.Count <> 0 Then
    strStatusCpa = LstVwSearchMgm.SelectedItem.SubItems(23)
    End If
    
    Me.MousePointer = vbNormal
    Exit Sub
'ke:
 '   strsql = "update usertbl set stsaplikasi=0  where userid ='" + MDIForm1.Text1.Text + "'"
  '  M_OBJCONN.Execute (strsql)
   ' MsgBox Err.Description
SALAH:
    cmdsql = "select * from mgm where "
    cmdsql = cmdsql + "agent = '" + MDIForm1.Text1.Text + "' and custid='"
    cmdsql = cmdsql + Trim(LstVwSearchMgm.SelectedItem.SubItems(1)) + "'"
    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    GoTo selanjutnya
    Exit Sub
End Sub

Private Sub LstVwSearchmgm_KeyPress(KeyAscii As Integer)
Dim M_Objrs As ADODB.Recordset
If KeyAscii = 13 Then
    Call LstVwSearchmgm_DblClick
    Exit Sub
End If
If UCase(MDIForm1.Text2.Text) <> "AGENT1" Then
    If KeyAscii = 112 Or KeyAscii = 80 Then
        If UCase(MDIForm1.Text2.Text) = "SUPERVISOR" Then
        'If UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
            Set M_Objrs = New ADODB.Recordset
            M_Objrs.CursorLocation = adUseClient
            M_Objrs.Open "SELECT USERID FROM usertbl WHERE TEAM ='" + MDIForm1.Text1.Text + "' AND USERID = '" + LstVwSearchMgm.SelectedItem.SubItems(11) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If M_Objrs.RecordCount <> 0 Then
                With View_AlihData
                    .Show vbModal
                    If .ok Then
                        LstVwSearchMgm.ListItems.Remove LstVwSearchMgm.SelectedItem.Index
                    End If
                End With
                Unload View_AlihData
            Else
                MsgBox "Data Ini Milik Agent Team Leader Yang Lain", vbCritical + vbOKOnly, "Aplikasi"
                Set M_Objrs = Nothing
                Exit Sub
            End If
            Set M_Objrs = Nothing
        Else
            With View_AlihData
                .Show vbModal
                If .ok Then
                    LstVwSearchMgm.ListItems.Remove LstVwSearchMgm.SelectedItem.Index
                End If
            End With
            Unload View_AlihData
        End If
    End If
    If KeyAscii = 73 Or KeyAscii = 105 Then
'        b_pindah = True
'        FrmCustIdTransfer.Show
        If UCase(Left(MDIForm1.Text2.Text, 5)) = "ADMIN" Then
            b_pindah = True
            FrmCustIdTransfer.Show
        ElseIf UCase(MDIForm1.Text2.Text) = "SUPERVISOR" Or UCase(MDIForm1.Text2.Text) = "MANAGER" Then
                    cmdsql = "SELECT USERID FROM usertbl WHERE USERID = '" + LstVwSearchMgm.SelectedItem.SubItems(11) + "'"
        
        '@@18092012 Teamleader Tidak diperbolehkan melakukan transfer
        ElseIf UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Or UCase(MDIForm1.Text2.Text) = "MANAGER" Then
                     '@@16-06-2010 Team Leader tidak boleh melakukan transfer ch kecuali pullout dan lunas
                     'If Combo1(0).Text <> "PULLOUT" Or Combo1(0).Text <> "LUNAS" Then
                        'MsgBox "Anda tidak dapat melakukan transfer data! Hubungi AM!", vbOKOnly + vbInformation, "Informasi"
                        'Exit Sub
                      'End If
                     MsgBox "Mohon maaf, pemindahan account data saat ini tidak diperbolehkan!", vbOKOnly + vbExclamation, "Peringatan"
                     
                     Exit Sub
                     cmdsql = "SELECT USERID FROM usertbl WHERE TEAM ='" + MDIForm1.Text1.Text + "' AND USERID = '" + LstVwSearchMgm.SelectedItem.SubItems(11) + "'"
        ElseIf UCase(MDIForm1.Text2.Text) = "AGENT" Then
                cmdsql = "SELECT USERID FROM usertbl where   userid='REVIEW'"
        End If
        
                    Set M_Objrs = New ADODB.Recordset
                    M_Objrs.CursorLocation = adUseClient
                    M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                    If M_Objrs.RecordCount <> 0 Then
                        b_pindah = True
                        FrmCustIdTransfer.Show
                    Else
                        'MsgBox "Data Ini Milik Agent Team Leader Yang Lain", vbCritical + vbOKOnly, "Aplikasi"
                        Set M_Objrs = Nothing
                        Exit Sub
                    End If
                    Set M_Objrs = Nothing
            
         
    End If
    If KeyAscii = 88 Or KeyAscii = 120 Then
        Dim n%
        If UCase(Left(MDIForm1.Text2.Text, 5)) = "ADMIN" Then
                n = 1
                Do While n <= LstVwSearchMgm.ListItems.Count
                    If LstVwSearchMgm.ListItems(n).Checked = True Then
                        Frmlock.List1.AddItem LstVwSearchMgm.ListItems(n).SubItems(1)
        '                Set ls2 = Frmlock.LSTACCESS.ListItems.ADD(, , LstVwSearchmgm.ListItems(N))
        '                ls2.SubItems(1) = LstVwSearchmgm.SelectedItem
                    End If
                    n = n + 1
                Loop
                Frmlock.Show
        End If
        
        If UCase(MDIForm1.Text2.Text) = "SUPERVISOR" And Combo1(2).Text = "" Then
            Set M_Objrs = New ADODB.Recordset
            M_Objrs.CursorLocation = adUseClient
            M_Objrs.Open "SELECT USERID FROM usertbl WHERE TEAM ='" + MDIForm1.Text1.Text + "' AND USERID = '" + LstVwSearchMgm.SelectedItem.SubItems(11) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If M_Objrs.RecordCount <> 0 Then
                'MsgBox "jkasja"
                n = 1
                Do While n <= LstVwSearchMgm.ListItems.Count
                    If LstVwSearchMgm.ListItems(n).Checked = True Then
                        Frmlock.List1.AddItem LstVwSearchMgm.ListItems(n).SubItems(1)
        '                Set ls2 = Frmlock.LSTACCESS.ListItems.ADD(, , LstVwSearchmgm.ListItems(N))
        '                ls2.SubItems(1) = LstVwSearchmgm.SelectedItem
                    End If
                    n = n + 1
                Loop
                Frmlock.Show
            Else
                MsgBox "Data Ini Milik Agent Team Leader Yang Lain", vbCritical + vbOKOnly, "Aplikasi"
                Set M_Objrs = Nothing
                Exit Sub
            End If
            Set M_Objrs = Nothing
        End If
    End If
End If
End Sub

'INBOUND LEADSE
'Private Sub cmbRecsource_LostFocus()
'Dim m_obj As New ADODB.Recordset
'm_obj.CursorLocation = adUseClient
'm_obj.Open "Select * from DATASOURCETBL WHERE KODEDS = '" + cmbRecsource.Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'If m_obj.RecordCount <> 0 Then
'    cmbRecsource.Text = m_obj!KODEDS
'Else
'    cmbRecsource.Text = ""
'End If
'Set m_obj = Nothing
'End Sub

'Private Sub CmdClaimLeads_Click(Index As Integer)
'Select Case Index
'Case 0
'    If Len(TxtNamaLeads.Text) < 2 Then
'        MsgBox "Nama harus diisi", vbInformation + vbOKOnly, "Aplikasi"
'        Exit Sub
'    End If
'    If Len(TxtTelpRumah.Text) < 2 And Len(TxtTelpKantor.Text) < 2 And Len(TxtHandPhone.Text) < 2 Then
'        MsgBox "Minimal salah satu dari telp harus diisi", vbInformation + vbOKOnly, "Aplikasi"
'        Exit Sub
'    End If
'    'CmdSave.Enabled = False
'    Call cari_duplicate_Leads
'Case 1
'    TdbDOB.Value = Empty
'    TxtNama.Text = Empty
'    TxtTelpRumah.Text = Empty
'    TxtTelpKantor.Text = Empty
'    TxtHandPhone.Text = Empty
'End Select
'End Sub
'

'Private Sub cari_duplicate_Leads()
'    Dim CMDSQL As String
'
'    Dim kriteria1 As String
'    Dim kriteria2 As String
'    Dim CUSTID1 As String
'    ' kriteria pertama
'    'nama ama notelp
'    If Len(TxtNamaLeads.Text) > 2 And Len(TxtTelpRumah.Text) > 2 Then
'        kriteria2 = Left(TxtTelpRumah.Text, 5)
'        CMDSQL = "Select * from cc_custtbl where name like '%" + TxtNamaLeads.Text + "%' "
'        CMDSQL = CMDSQL + " and (HOMENO like '%" + kriteria2 + "%' or HOMENO2 like '%" + kriteria2 + "%' or mobileno like '%" + kriteria2 + "%' or mobileno2 like '%" + kriteria2 + "%' or officeno like '%" + kriteria2 + "%' or officeno2 like '%" + kriteria2 + "%') "
'
'    Set mrs_cek = New ADODB.Recordset
'        mrs_cek.CursorLocation = adUseClient
'
'        mrs_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'        If mrs_cek.RecordCount <> 0 Then
'            CUSTID1 = Empty
'            While Not mrs_cek.EOF
'                CUSTID1 = "REFI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC"))
'                CMDSQL = "Insert into TBL_DUPLIKASI (CUSTID, CUSTIDBARU, NAMABARU, HOMENOBARU, OFFICENOBARU, MOBILENOBARU, AGENTBARU, "
'                If TdbDOB.ValueIsNull = False Then
'                    CMDSQL = CMDSQL + "DOB,"
'                End If
'                CMDSQL = CMDSQL + " RECSOURCEBARU) values "
'                CMDSQL = CMDSQL + "('" + mrs_cek!CUSTID + "',"
'                CMDSQL = CMDSQL + "'" + CUSTID1 + "',"
'                CMDSQL = CMDSQL + "'" + TxtNamaLeads.Text + "',"
'                CMDSQL = CMDSQL + "'" + TxtTelpRumah.Text + "',"
'                CMDSQL = CMDSQL + "'" + TxtTelpKantor.Text + "',"
'                CMDSQL = CMDSQL + "'" + TxtHandPhoneLeads.Text + "',"
'                CMDSQL = CMDSQL + "'" + MDIForm1.Text1.Text + "',"
'                If TdbDOB.ValueIsNull = False Then
'                    CMDSQL = CMDSQL + "'" + Format(TdbDOB.Value, "yyyy/mm/dd") + "',"
'                End If
'                CMDSQL = CMDSQL + "'" + cmbRecsource.Text + "')"
'                M_OBJCONN.Execute CMDSQL
'                mrs_cek.MoveNext
'            Wend
'
'            ' tampilin yang duplicate deh...
'                Call show_Leads_Duplicate
'                MsgBox " Nama dan Telp Rumah Ada yg sama", vbInformation + vbOKOnly, "Aplikasi"
'            Set mrs_cek = Nothing
'            Exit Sub
'
'        End If
'        Set mrs_cek = Nothing
'    End If
'    If Len(TxtNamaLeads.Text) > 2 And Len(TxtTelpKantor.Text) > 2 Then
'        kriteria2 = Left(TxtTelpKantor.Text, 5)
'        CMDSQL = "Select * from cc_custtbl where name like '%" + TxtNamaLeads.Text + "%' "
'        CMDSQL = CMDSQL + " and (HOMENO like '%" + kriteria2 + "%' or HOMENO2 like '%" + kriteria2 + "%' or mobileno like '%" + kriteria2 + "%' or mobileno2 like '%" + kriteria2 + "%' or officeno like '%" + kriteria2 + "%' or officeno2 like '%" + kriteria2 + "%') "
'    Set mrs_cek = New ADODB.Recordset
'    mrs_cek.CursorLocation = adUseClient
'
'        mrs_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'        If mrs_cek.RecordCount <> 0 Then
'            CUSTID1 = Empty
'            While Not mrs_cek.EOF
'                CUSTID1 = "REFI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC"))
'
'                CMDSQL = "Insert into TBL_DUPLIKASI (CUSTID, CUSTIDBARU, NAMABARU, HOMENOBARU, OFFICENOBARU, MOBILENOBARU, AGENTBARU, "
'                If TdbDOB.ValueIsNull = False Then
'                    CMDSQL = CMDSQL + "DOB,"
'                End If
'                CMDSQL = CMDSQL + " RECSOURCEBARU) values "
'                CMDSQL = CMDSQL + "('" + mrs_cek!CUSTID + "',"
'                CMDSQL = CMDSQL + "'" + CUSTID1 + "',"
'                CMDSQL = CMDSQL + "'" + TxtNamaLeads.Text + "',"
'                CMDSQL = CMDSQL + "'" + TxtTelpRumah.Text + "',"
'                CMDSQL = CMDSQL + "'" + TxtTelpKantor.Text + "',"
'                CMDSQL = CMDSQL + "'" + TxtHandPhoneLeads.Text + "',"
'                CMDSQL = CMDSQL + "'" + MDIForm1.Text1.Text + "',"
'                If TdbDOB.ValueIsNull = False Then
'                    CMDSQL = CMDSQL + "'" + Format(TdbDOB.Value, "yyyy/mm/dd") + "',"
'                End If
'                CMDSQL = CMDSQL + "'" + cmbRecsource.Text + "')"
'                M_OBJCONN.Execute CMDSQL
'                mrs_cek.MoveNext
'            Wend
'
'            ' show data
'            Call show_Leads_Duplicate
'            MsgBox " Nama dan Telp Kantor Ada yg sama", vbInformation + vbOKOnly, "Aplikasi"
'            Set mrs_cek = Nothing
'            Exit Sub
'        End If
'        Set mrs_cek = Nothing
'
'    End If
'    If Len(TxtNamaLeads.Text) > 2 And Len(TxtHandPhone.Text) > 2 Then
'        kriteria2 = Left(TxtHandPhone.Text, 8)
'        CMDSQL = "Select * from cc_custtbl where name like '%" + TxtNamaLeads.Text + "%' "
'        CMDSQL = CMDSQL + " and (HOMENO like '%" + kriteria2 + "%' or HOMENO2 like '%" + kriteria2 + "%' or mobileno like '%" + kriteria2 + "%' or mobileno2 like '%" + kriteria2 + "%' or officeno like '%" + kriteria2 + "%' or officeno2 like '%" + kriteria2 + "%') "
'    Set mrs_cek = New ADODB.Recordset
'    mrs_cek.CursorLocation = adUseClient
'
'        mrs_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'        If mrs_cek.RecordCount <> 0 Then
'
'
'            CUSTID1 = Empty
'            While Not mrs_cek.EOF
'                CUSTID1 = "REFI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC"))
'
'                CMDSQL = "Insert into TBL_DUPLIKASI (CUSTID, CUSTIDBARU, NAMABARU, HOMENOBARU, OFFICENOBARU, MOBILENOBARU, AGENTBARU, "
'                If TdbDOB.ValueIsNull = False Then
'                    CMDSQL = CMDSQL + "DOB,"
'                End If
'                CMDSQL = CMDSQL + " RECSOURCEBARU) values "
'                CMDSQL = CMDSQL + "('" + mrs_cek!CUSTID + "',"
'                CMDSQL = CMDSQL + "'" + CUSTID1 + "',"
'                CMDSQL = CMDSQL + "'" + TxtNamaLeads.Text + "',"
'                CMDSQL = CMDSQL + "'" + TxtTelpRumah.Text + "',"
'                CMDSQL = CMDSQL + "'" + TxtTelpKantor.Text + "',"
'                CMDSQL = CMDSQL + "'" + TxtHandPhoneLeads.Text + "',"
'                CMDSQL = CMDSQL + "'" + MDIForm1.Text1.Text + "',"
'                If TdbDOB.ValueIsNull = False Then
'                    CMDSQL = CMDSQL + "'" + Format(TdbDOB.Value, "yyyy/mm/dd") + "',"
'                End If
'                CMDSQL = CMDSQL + "'" + cmbRecsource.Text + "')"
'
'                M_OBJCONN.Execute CMDSQL
'                mrs_cek.MoveNext
'            Wend
'
'            ' show data
'            Call show_Leads_Duplicate
'            MsgBox "Nama dan Handphone Ada yg sama", vbInformation + vbOKOnly, "Aplikasi"
'            Set mrs_cek = Nothing
'            Exit Sub
'        End If
'        Set mrs_cek = Nothing
'
'    End If
'    If Len(TxtNamaLeads.Text) > 2 And TdbDOB.ValueIsNull = False Then
'        kriteria2 = Format(TdbDOB.Value, "yyyy/mm/dd")
'        CMDSQL = "Select * from cc_custtbl where name like '%" + TxtNamaLeads.Text + "%' "
'        CMDSQL = CMDSQL + " and birthd = '" + Format(TdbDOB.Value, "yyyy/mm/dd") + "'"
'        Set mrs_cek = New ADODB.Recordset
'            mrs_cek.CursorLocation = adUseClient
'
'        mrs_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'        If mrs_cek.RecordCount <> 0 Then
'
'
'            CUSTID1 = Empty
'            While Not mrs_cek.EOF
'                CUSTID1 = "REFI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC"))
'
'                CMDSQL = "Insert into TBL_DUPLIKASI (CUSTID, CUSTIDBARU, NAMABARU, HOMENOBARU, OFFICENOBARU, MOBILENOBARU, AGENTBARU, "
'                If TdbDOB.ValueIsNull = False Then
'                    CMDSQL = CMDSQL + "DOB,"
'                End If
'                CMDSQL = CMDSQL + " RECSOURCEBARU) values "
'                CMDSQL = CMDSQL + "('" + mrs_cek!CUSTID + "',"
'                CMDSQL = CMDSQL + "'" + CUSTID1 + "',"
'                CMDSQL = CMDSQL + "'" + TxtNamaLeads.Text + "',"
'                CMDSQL = CMDSQL + "'" + TxtTelpRumah.Text + "',"
'                CMDSQL = CMDSQL + "'" + TxtTelpKantor.Text + "',"
'                CMDSQL = CMDSQL + "'" + TxtHandPhoneLeads.Text + "',"
'                CMDSQL = CMDSQL + "'" + MDIForm1.Text1.Text + "',"
'                If TdbDOB.ValueIsNull = False Then
'                    CMDSQL = CMDSQL + "'" + Format(TdbDOB.Value, "yyyy/mm/dd") + "',"
'                End If
'                CMDSQL = CMDSQL + "'" + cmbRecsource.Text + "')"
'
'                M_OBJCONN.Execute CMDSQL
'                mrs_cek.MoveNext
'            Wend
'
'            ' show data
'            Call show_Leads_Duplicate
'            MsgBox "Nama dan DOB Ada yg sama", vbInformation + vbOKOnly, "Aplikasi"
'            Set mrs_cek = Nothing
'            Exit Sub
'        End If
'        Set mrs_cek = Nothing
'    End If
''        custid1 = "REFI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC"))
''
''        cmdsql = "Insert into CC_CUSTTBL(CUSTID, NAME, HOMENO, MOBILENO, OFFICENO, AGENT, RECSOURCE,CustIdRef,"
''        If TdbDOB.ValueIsNull = False Then
''            cmdsql = cmdsql + "BIRTHD,"
''        End If
''        cmdsql = cmdsql + " RecSourceRef) values"
''        cmdsql = cmdsql + "('" + custid1 + "',"
''        cmdsql = cmdsql + "'" + txtnamaleads.Text + "',"
''        cmdsql = cmdsql + "'" + TxtTelpRumah.Text + "',"
''        cmdsql = cmdsql + "'" + TxtHandphoneleads.Text + "',"
''        cmdsql = cmdsql + "'" + TxtTelpKantor.Text + "',"
''        cmdsql = cmdsql + "'" + MDIForm1.Text1.Text + "',"
''        cmdsql = cmdsql + "'" + cmbRecsource.Text + "',"
''        cmdsql = cmdsql + "'" + TxtIdReff.Text + "',"
''        If TdbDOB.ValueIsNull = False Then
''            cmdsql = cmdsql + "'" + Format(TdbDOB.Value, "yyyy/mm/dd") + "',"
''        End If
''        cmdsql = cmdsql + "'" + cmbRecsource.Text + "')"
''        M_OBJCONN.Execute cmdsql
'
'
'        ' munculin tuh form buat entry reff
'        With FrmEntryReff
'            .TxtIdReff.Text = "Inbound Leads"
'            .TxtNama.Text = TxtNamaLeads.Text
'            .TxtTelpRumah.Text = TxtTelpRumah.Text
'            .TxtTelpKantor.Text = TxtTelpKantor.Text
'            .TxtHandPhone.Text = TxtHandPhoneLeads.Text
'            .TdbDOB.Value = TdbDOB.Value
'            .TxtIdReff.Enabled = False
'             .Show vbModal
'             If .okReff = True Then
'                MsgBox "Data sudah tersimpan", vbInformation + vbOKOnly, "Aplikasi"
'             Else
'                MsgBox "Cancel", vbInformation + vbOKOnly, "Aplikasi"
'             End If
'        End With
'            Unload FrmEntryReff
'End Sub

'Private Sub show_Leads_Duplicate()
'Dim listitem As listitem
'ListView1.ListItems.Clear
'SSTab1.Tab = 1
'mrs_cek.MoveFirst
'While Not mrs_cek.EOF
'    Set listitem = ListView1.ListItems.ADD(, , mrs_cek.Bookmark)
'        listitem.SubItems(1) = IIf(IsNull(mrs_cek("custid")), "", mrs_cek("custid"))
'        Select Case mrs_cek("RECSTATUS")
'        Case "1A"
'            listitem.SubItems(2) = "Available"
'        Case ""
'            listitem.SubItems(2) = "Available"
'        Case Else
'            listitem.SubItems(2) = IIf(IsNull(mrs_cek("PRIOR")), "", mrs_cek("PRIOR"))
'        End Select
'        listitem.SubItems(3) = IIf(IsNull(mrs_cek("CUSTIDREF")), "", mrs_cek("CUSTIDREF"))
'        listitem.SubItems(4) = IIf(IsNull(mrs_cek("NAMAREF")), "", mrs_cek("NAMAREF"))
'        listitem.SubItems(5) = IIf(IsNull(mrs_cek("NAME")), "", mrs_cek("NAME"))
'        listitem.SubItems(6) = IIf(IsNull(mrs_cek("NEXTACTDATE")), "", Format(mrs_cek("NEXTACTDATE"), "yyyy/mm/dd hh:mm"))
'        listitem.SubItems(7) = IIf(IsNull(mrs_cek("NEXTACT")), "", mrs_cek("NEXTACT"))
'        listitem.SubItems(8) = IIf(IsNull(mrs_cek("REMARKS")), "", mrs_cek("REMARKS"))
'        listitem.SubItems(9) = IIf(IsNull(mrs_cek("AGENT")), "", mrs_cek("AGENT"))
'        listitem.SubItems(10) = IIf(IsNull(mrs_cek("NamaAGENT")), "", mrs_cek("NamaAGENT"))
'        listitem.SubItems(11) = IIf(IsNull(mrs_cek("RECSOURCEREF")), "", mrs_cek("RECSOURCEREF"))
'        listitem.SubItems(12) = Format(IIf(IsNull(mrs_cek("TGLSTATUS")), "", mrs_cek("TGLSTATUS")), "yyyy/mm/dd")
'        listitem.SubItems(13) = IIf(IsNull(mrs_cek("Kethslkerja")), "", mrs_cek("Kethslkerja"))
'        listitem.SubItems(14) = IIf(IsNull(mrs_cek("KdComplaint")), "", mrs_cek("KdComplaint"))
'        listitem.SubItems(15) = IIf(IsNull(mrs_cek("RemarkComplaint")), "", mrs_cek("RemarkComplaint"))
'        listitem.SubItems(16) = IIf(IsNull(mrs_cek("F_CEK")), "", mrs_cek("F_CEK"))
'        listitem.SubItems(17) = IIf(IsNull(mrs_cek("Nomor")), "", mrs_cek("Nomor"))
'        mrs_cek.MoveNext
'Wend
'Set mrs_cek = Nothing
'End Sub


'Private Sub CmdClaim_Click(Index As Integer)
'
'Select Case Index
'    Case 0
'        If Len(TxtNama.Text) < 2 Then
'            MsgBox "Nama harus diisi", vbInformation + vbOKOnly, "Aplikasi"
'            Exit Sub
'        End If
'        If Len(txtnotelprumah.Text) < 2 And Len(TxtNoTelpKantor.Text) < 2 And Len(TxtHandPhone.Text) < 2 Then
'            MsgBox "Minimal salah satu dari telp harus diisi", vbInformation + vbOKOnly, "Aplikasi"
'            Exit Sub
'        End If
'        'CmdSave.Enabled = False
'        Call cari_duplicate_CH
'    Case 1
'        TxtNama.Text = ""
'        txtnotelprumah.Text = ""
'        TxtNoTelpKantor.Text = ""
'        TxtHandPhone.Text = ""
'        TdbDOBCH.Value = Empty
'        cmbRecsourcech.Text = ""
'End Select




'Dim m_objrs As ADODB.Recordset
'Dim cmdsql As String
'Dim Lcustid, LName, LHOMENO, LOFFICENO, LMOBILE, LAgent, LNAMAAGENT, LRECSOURCE, LOTHERS, LKethslkerja As String
'Select Case Index
'Case 0
'    If Len(TxtNama.Text) = 0 Or Len(txtnotelprumah.Text) = 0 Then
'        MsgBox "Nama atau no telp rumah harus diisi", vbInformation + vbOKOnly, "Informasi"
'        Exit Sub
'    End If
'
'    Set m_objrs = New ADODB.Recordset
'    m_objrs.CursorLocation = adUseClient
'    cmdsql = "Select  * from mgm where LEFT(recsource,3) <>'PRE' AND HOMENO like '%" + txtnotelprumah.Text + "%' "
'    m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'    While Not m_objrs.EOF
'        LKethslkerja = m_objrs!KETHSLKERJA
'        LName = m_objrs!Name
'        If UCase(Trim(LName)) = UCase(Trim(TxtNama.Text)) Then
'            If UCase(LKethslkerja) = "1A" Then
'                m_objrs!NEXTACT = "Data Inbound Call"
'                m_objrs!agent = MDIForm1.Text1.Text
'                m_objrs!NAMAAGENT = MDIForm1.Text7.Text
'                m_objrs.UPDATE
'                MsgBox "Sukses... Data Sudah Ditransfer", vbInformation + vbOKOnly, "Informasi"
'
'                Call isi_dataClaimKeGrid(m_objrs!CUSTID, m_objrs!Name, "Data Inbound Call", "Data Inbound Call", MDIForm1.Text1.Text, MDIForm1.Text7.Text, m_objrs!RECSOURCE)
'
'                Set m_objrs = Nothing
'                TxtNama.Text = ""
'                txtnotelprumah.Text = ""
'                TxtNoTelpKantor.Text = ""
'                TxtHandPhone.Text = ""
'                TxtNama.SetFocus
'                Exit Sub
'            Else
'                MsgBox "Tidak Sukses... Data Sudah di follow Up oleh " & m_objrs!agent
'                Set m_objrs = Nothing
'                TxtNama.Text = ""
'                txtnotelprumah.Text = ""
'                TxtNoTelpKantor.Text = ""
'                TxtHandPhone.Text = ""
'                TxtNama.SetFocus
'                Exit Sub
'            End If
'        End If
'    m_objrs.MoveNext
'    Wend
'    Set m_objrs = Nothing
'
'
'    Set m_objrs = New ADODB.Recordset
'    m_objrs.CursorLocation = adUseClient
'    cmdsql = "Select  * from tempCC_CUSTTBL where LEFT(recsource,3) <>'PRE' AND HOMENO like '%" + txtnotelprumah.Text + "%'"
'    m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'    While Not m_objrs.EOF
'        LKethslkerja = m_objrs!KETHSLKERJA
'        LName = m_objrs!Name
'        If UCase(Trim(Replace(LName, "(PVA)", ""))) = UCase(Trim(TxtNama.Text)) Then
'            Lcustid = m_objrs!CUSTID
'            LName = m_objrs!Name
'            LHOMENO = m_objrs!HOMENO
'            LOFFICENO = m_objrs!OFFICENO
'            LMOBILE = m_objrs!MOBILENO
'            LAgent = m_objrs!agent
'            LNAMAAGENT = m_objrs!NAMAAGENT
'            LRECSOURCE = m_objrs!RECSOURCE
'            LOTHERS = m_objrs!OTHERS
'            LKethslkerja = m_objrs!KETHSLKERJA
'            cmdsql = "Insert Into mgm (CUSTID, NAME, TGLDISTRIBUSI, HOMENO, OFFICENO, MOBILENO, AGENT, NAMAAGENT, RECSOURCE, nextact, KETHSLKERJA)"
'            cmdsql = cmdsql + " VALUES"
'            cmdsql = cmdsql + "('" + Lcustid + "',"
'            cmdsql = cmdsql + " '" + LName + "',"
'            cmdsql = cmdsql + " '" + Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd hh:nn") + "',"
'            cmdsql = cmdsql + " '" + LHOMENO + "',"
'            cmdsql = cmdsql + " '" + LOFFICENO + "',"
'            cmdsql = cmdsql + " '" + LMOBILE + "',"
'            cmdsql = cmdsql + " '" + MDIForm1.Text1.Text + "',"
'            cmdsql = cmdsql + " '" + MDIForm1.Text7.Text + "',"
'            cmdsql = cmdsql + " '" + LRECSOURCE + "',"
'            cmdsql = cmdsql + " 'Data Inbound Call',"
'            cmdsql = cmdsql + " '" + LKethslkerja + "')"
'            M_OBJCONN.Execute cmdsql
'            Call isi_dataClaimKeGrid(CStr(Lcustid), CStr(LName), "Data Inbound Call", "Data Inbound Call", MDIForm1.Text1.Text, MDIForm1.Text7.Text, CStr(LRECSOURCE))
'            MsgBox "Sukses... Data Sudah Ditransfer", vbInformation + vbOKOnly, "Informasi"
'            M_OBJCONN.Execute "Delete from TempCC_Custtbl where custid ='" + Lcustid + "'"
'            Set m_objrs = Nothing
'            TxtNama.Text = ""
'            txtnotelprumah.Text = ""
'            TxtNoTelpKantor.Text = ""
'            TxtHandPhone.Text = ""
'            TxtNama.SetFocus
'            Exit Sub
'        End If
'    m_objrs.MoveNext
'    Wend
'    Set m_objrs = Nothing
'
'            cmdsql = "Insert Into mgm (CUSTID, NAME, TGLDISTRIBUSI, HOMENO, OFFICENO, MOBILENO, AGENT, NAMAAGENT, RECSOURCE, nextact, KETHSLKERJA)"
'            cmdsql = cmdsql + " VALUES"
'            Lcustid = "mgmI-" & CUSTNOMOR(M_OBJCONN, "FRMCUST_CC")
'
'            cmdsql = cmdsql + "('" + Lcustid + "',"
'            cmdsql = cmdsql + " '" + TxtNama.Text + "',"
'            cmdsql = cmdsql + " '" + Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd hh:nn") + "',"
'            cmdsql = cmdsql + " '" + txtnotelprumah.Text + "',"
'            cmdsql = cmdsql + " '" + TxtNoTelpKantor.Text + "',"
'            cmdsql = cmdsql + " '" + TxtHandPhone.Text + "',"
'            cmdsql = cmdsql + " '" + MDIForm1.Text1.Text + "',"
'            cmdsql = cmdsql + " '" + MDIForm1.Text7.Text + "',"
'            cmdsql = cmdsql + " 'mgm_INC',"
'            cmdsql = cmdsql + " 'Data Inbound Call',"
'            cmdsql = cmdsql + " '1A')"
'    M_OBJCONN.Execute cmdsql
'    MsgBox "Sukses..", vbInformation, "Informasi"
'
'    Call isi_dataClaimKeGrid(CStr(Lcustid), TxtNama.Text, "Data Inbound Call", "Data Inbound Call", MDIForm1.Text1.Text, MDIForm1.Text7.Text, "mgm_INC")
'
'    TxtNama.Text = ""
'    txtnotelprumah.Text = ""
'    TxtNoTelpKantor.Text = ""
'    TxtHandPhone.Text = ""
'    TxtNama.SetFocus
'Case 1
'    TxtNama.Text = ""
'    txtnotelprumah.Text = ""
'    TxtNoTelpKantor.Text = ""
'    TxtHandPhone.Text = ""
'End Select
'End Sub


'Private Sub cari_duplicate_CH()
'    Dim CMDSQL As String
'
'    Dim kriteria1 As String
'    Dim kriteria2 As String
'    Dim CUSTID1 As String
'    ' kriteria pertama
'    'nama ama notelp
'    If Len(TxtNama.Text) > 2 And Len(txtnotelprumah.Text) > 2 Then
'        kriteria2 = Left(txtnotelprumah.Text, 5)
'        CMDSQL = "Select * from mgm where name like '%" + TxtNama.Text + "%' "
'        CMDSQL = CMDSQL + " and (HOMENO like '%" + kriteria2 + "%' or HOMENO2 like '%" + kriteria2 + "%' or mobileno like '%" + kriteria2 + "%' or mobileno2 like '%" + kriteria2 + "%' or officeno like '%" + kriteria2 + "%' or officeno2 like '%" + kriteria2 + "%') "
'
'    Set mrs_cek = New ADODB.Recordset
'        mrs_cek.CursorLocation = adUseClient
'
'        mrs_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'        If mrs_cek.RecordCount <> 0 Then
'            CUSTID1 = Empty
'            While Not mrs_cek.EOF
'                CUSTID1 = "mgmI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC"))
'                CMDSQL = "Insert into TBL_DUPLIKASICH (CUSTID, CUSTIDBARU, NAMABARU, HOMENOBARU, OFFICENOBARU, MOBILENOBARU, AGENTBARU, "
'                If TdbDOB.ValueIsNull = False Then
'                    CMDSQL = CMDSQL + "DOB,"
'                End If
'                CMDSQL = CMDSQL + " RECSOURCEBARU) values "
'                CMDSQL = CMDSQL + "('" + mrs_cek!CustId + "',"
'                CMDSQL = CMDSQL + "'" + CUSTID1 + "',"
'                CMDSQL = CMDSQL + "'" + TxtNama.Text + "',"
'                CMDSQL = CMDSQL + "'" + txtnotelprumah.Text + "',"
'                CMDSQL = CMDSQL + "'" + TxtNoTelpKantor.Text + "',"
'                CMDSQL = CMDSQL + "'" + TxtHandPhone.Text + "',"
'                CMDSQL = CMDSQL + "'" + MDIForm1.Text1.Text + "',"
'                If TdbDOB.ValueIsNull = False Then
'                    CMDSQL = CMDSQL + "'" + Format(TdbDOBCH.Value, "yyyy/mm/dd") + "',"
'                End If
'                CMDSQL = CMDSQL + "'" + cmbRecsourcech.Text + "')"
'                M_OBJCONN.Execute CMDSQL
'                mrs_cek.MoveNext
'            Wend
'
'            ' tampilin yang duplicate deh...
'                Call show_Ch_Duplicate
'                MsgBox " Nama dan Telp Rumah Ada yg sama", vbInformation + vbOKOnly, "Aplikasi"
'            Set mrs_cek = Nothing
'            Exit Sub
'
'        End If
'        Set mrs_cek = Nothing
'    End If
'    If Len(TxtNama.Text) > 2 And Len(TxtNoTelpKantor.Text) > 2 Then
'        kriteria2 = Left(TxtNoTelpKantor.Text, 5)
'        CMDSQL = "Select * from mgm where name like '%" + TxtNama.Text + "%' "
'        CMDSQL = CMDSQL + " and (HOMENO like '%" + kriteria2 + "%' or HOMENO2 like '%" + kriteria2 + "%' or mobileno like '%" + kriteria2 + "%' or mobileno2 like '%" + kriteria2 + "%' or officeno like '%" + kriteria2 + "%' or officeno2 like '%" + kriteria2 + "%') "
'    Set mrs_cek = New ADODB.Recordset
'    mrs_cek.CursorLocation = adUseClient
'
'        mrs_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'        If mrs_cek.RecordCount <> 0 Then
'            CUSTID1 = Empty
'            While Not mrs_cek.EOF
'                CUSTID1 = "mgmI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC"))
'
'                CMDSQL = "Insert into TBL_DUPLIKASICH (CUSTID, CUSTIDBARU, NAMABARU, HOMENOBARU, OFFICENOBARU, MOBILENOBARU, AGENTBARU, "
'                If TdbDOBCH.ValueIsNull = False Then
'                    CMDSQL = CMDSQL + "DOB,"
'                End If
'                CMDSQL = CMDSQL + " RECSOURCEBARU) values "
'                CMDSQL = CMDSQL + "('" + mrs_cek!CustId + "',"
'                CMDSQL = CMDSQL + "'" + CUSTID1 + "',"
'                CMDSQL = CMDSQL + "'" + TxtNama.Text + "',"
'                CMDSQL = CMDSQL + "'" + txtnotelprumah.Text + "',"
'                CMDSQL = CMDSQL + "'" + TxtNoTelpKantor.Text + "',"
'                CMDSQL = CMDSQL + "'" + TxtHandPhone.Text + "',"
'                CMDSQL = CMDSQL + "'" + MDIForm1.Text1.Text + "',"
'                If TdbDOBCH.ValueIsNull = False Then
'                    CMDSQL = CMDSQL + "'" + Format(TdbDOBCH.Value, "yyyy/mm/dd") + "',"
'                End If
'                CMDSQL = CMDSQL + "'" + cmbRecsourcech.Text + "')"
'                M_OBJCONN.Execute CMDSQL
'                mrs_cek.MoveNext
'            Wend
'
'            ' show data
'            Call show_Ch_Duplicate
'            MsgBox " Nama dan Telp Kantor Ada yg sama", vbInformation + vbOKOnly, "Aplikasi"
'            Set mrs_cek = Nothing
'            Exit Sub
'        End If
'        Set mrs_cek = Nothing
'
'    End If
'    If Len(TxtNama.Text) > 2 And Len(TxtHandPhone.Text) > 2 Then
'        kriteria2 = Left(TxtHandPhone.Text, 8)
'        CMDSQL = "Select * from mgm where name like '%" + TxtNama.Text + "%' "
'        CMDSQL = CMDSQL + " and (HOMENO like '%" + kriteria2 + "%' or HOMENO2 like '%" + kriteria2 + "%' or mobileno like '%" + kriteria2 + "%' or mobileno2 like '%" + kriteria2 + "%' or officeno like '%" + kriteria2 + "%' or officeno2 like '%" + kriteria2 + "%') "
'    Set mrs_cek = New ADODB.Recordset
'    mrs_cek.CursorLocation = adUseClient
'
'        mrs_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'        If mrs_cek.RecordCount <> 0 Then
'
'
'            CUSTID1 = Empty
'            While Not mrs_cek.EOF
'                CUSTID1 = "mgmI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC"))
'
'                CMDSQL = "Insert into TBL_DUPLIKASICH (CUSTID, CUSTIDBARU, NAMABARU, HOMENOBARU, OFFICENOBARU, MOBILENOBARU, AGENTBARU, "
'                If TdbDOBCH.ValueIsNull = False Then
'                    CMDSQL = CMDSQL + "DOB,"
'                End If
'                CMDSQL = CMDSQL + " RECSOURCEBARU) values "
'                CMDSQL = CMDSQL + "('" + mrs_cek!CustId + "',"
'                CMDSQL = CMDSQL + "'" + CUSTID1 + "',"
'                CMDSQL = CMDSQL + "'" + TxtNama.Text + "',"
'                CMDSQL = CMDSQL + "'" + txtnotelprumah.Text + "',"
'                CMDSQL = CMDSQL + "'" + TxtNoTelpKantor.Text + "',"
'                CMDSQL = CMDSQL + "'" + TxtHandPhone.Text + "',"
'                CMDSQL = CMDSQL + "'" + MDIForm1.Text1.Text + "',"
'                If TdbDOBCH.ValueIsNull = False Then
'                    CMDSQL = CMDSQL + "'" + Format(TdbDOBCH.Value, "yyyy/mm/dd") + "',"
'                End If
'                CMDSQL = CMDSQL + "'" + cmbRecsourcech.Text + "')"
'
'                M_OBJCONN.Execute CMDSQL
'                mrs_cek.MoveNext
'            Wend
'
'            ' show data
'            Call show_Ch_Duplicate
'            MsgBox "Nama dan Handphone Ada yg sama", vbInformation + vbOKOnly, "Aplikasi"
'            Set mrs_cek = Nothing
'            Exit Sub
'        End If
'        Set mrs_cek = Nothing
'
'    End If
'    If Len(TxtNama.Text) > 2 And TdbDOBCH.ValueIsNull = False Then
'        kriteria2 = Format(TdbDOBCH.Value, "yyyy/mm/dd")
'        CMDSQL = "Select * from mgm where name like '%" + TxtNama.Text + "%' "
'        CMDSQL = CMDSQL + " and birthd = '" + Format(TdbDOBCH.Value, "yyyy/mm/dd") + "'"
'        Set mrs_cek = New ADODB.Recordset
'            mrs_cek.CursorLocation = adUseClient
'
'        mrs_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'        If mrs_cek.RecordCount <> 0 Then
'
'
'            CUSTID1 = Empty
'            While Not mrs_cek.EOF
'                CUSTID1 = "mgmI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC"))
'
'                CMDSQL = "Insert into TBL_DUPLIKASICH (CUSTID, CUSTIDBARU, NAMABARU, HOMENOBARU, OFFICENOBARU, MOBILENOBARU, AGENTBARU, "
'                If TdbDOBCH.ValueIsNull = False Then
'                    CMDSQL = CMDSQL + "DOB,"
'                End If
'                CMDSQL = CMDSQL + " RECSOURCEBARU) values "
'                CMDSQL = CMDSQL + "('" + mrs_cek!CustId + "',"
'                CMDSQL = CMDSQL + "'" + CUSTID1 + "',"
'                CMDSQL = CMDSQL + "'" + TxtNama.Text + "',"
'                CMDSQL = CMDSQL + "'" + txtnotelprumah.Text + "',"
'                CMDSQL = CMDSQL + "'" + TxtNoTelpKantor.Text + "',"
'                CMDSQL = CMDSQL + "'" + TxtHandPhone.Text + "',"
'                CMDSQL = CMDSQL + "'" + MDIForm1.Text1.Text + "',"
'                If TdbDOBCH.ValueIsNull = False Then
'                    CMDSQL = CMDSQL + "'" + Format(TdbDOBCH.Value, "yyyy/mm/dd") + "',"
'                End If
'                CMDSQL = CMDSQL + "'" + cmbRecsourcech.Text + "')"
'
'                M_OBJCONN.Execute CMDSQL
'                mrs_cek.MoveNext
'            Wend
'
'            ' show data
'            Call show_Ch_Duplicate
'            MsgBox "Nama dan DOB Ada yg sama", vbInformation + vbOKOnly, "Aplikasi"
'            Set mrs_cek = Nothing
'            Exit Sub
'        End If
'        Set mrs_cek = Nothing
'    End If
''        custid1 = "mgmI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC"))
''
''        cmdsql = "Insert into mgm(CUSTID, NAME, HOMENO, MOBILENO, OFFICENO, AGENT, RECSOURCE,CustIdRef,"
''        If TdbDobCh.ValueIsNull = False Then
''            cmdsql = cmdsql + "BIRTHD,"
''        End If
''        cmdsql = cmdsql + " RecSourceRef) values"
''        cmdsql = cmdsql + "('" + custid1 + "',"
''        cmdsql = cmdsql + "'" + TxtNama.Text + "',"
''        cmdsql = cmdsql + "'" + txtnotelprumah.Text + "',"
''        cmdsql = cmdsql + "'" + TxtHandphone.Text + "',"
''        cmdsql = cmdsql + "'" + TxtNoTelpKantor.Text + "',"
''        cmdsql = cmdsql + "'" + MDIForm1.Text1.Text + "',"
''        cmdsql = cmdsql + "'" + CmbRecsourceCh.Text + "',"
''        cmdsql = cmdsql + "'" + TxtIdReff.Text + "',"
''        If TdbDobCh.ValueIsNull = False Then
''            cmdsql = cmdsql + "'" + Format(TdbDobCh.Value, "yyyy/mm/dd") + "',"
''        End If
''        cmdsql = cmdsql + "'" + CmbRecsourceCh.Text + "')"
''        M_OBJCONN.Execute cmdsql
'
'
'        ' munculin tuh form buat entry mgm
'        With FrmEntryCH
'            .TxtIdReff.Text = "Inbound mgm"
'            .TxtNama.Text = TxtNama.Text
'            .TxtTelpRumah.Text = txtnotelprumah.Text
'            .TxtTelpKantor.Text = TxtNoTelpKantor.Text
'            .TxtHandPhone.Text = TxtHandPhone.Text
'            .TdbDOBCH.Value = TdbDOBCH.Value
'            .TxtIdReff.Enabled = False
'             .Show vbModal
'             If .okReff = True Then
'                MsgBox "Data sudah tersimpan", vbInformation + vbOKOnly, "Aplikasi"
'             Else
'                MsgBox "Cancel", vbInformation + vbOKOnly, "Aplikasi"
'             End If
'        End With
'            Unload FrmEntryCH
'
'
'End Sub

'Private Sub show_Ch_Duplicate()
'Dim listitem As listitem
'SSTab1.Tab = 0
'
'On Error GoTo HELL
'    mrs_cek.MoveFirst
'    LstVwSearchmgm.ListItems.Clear
'
'    While Not mrs_cek.EOF
'        Set listitem = LstVwSearchmgm.ListItems.ADD(, , mrs_cek.Bookmark)
'        listitem.SubItems(1) = IIf(IsNull(mrs_cek("CUSTID")), "", mrs_cek("CUSTID"))
'        listitem.SubItems(2) = IIf(IsNull(mrs_cek("PRIOR")), "", mrs_cek("PRIOR"))
'        listitem.SubItems(3) = IIf(IsNull(mrs_cek("NAME")), "", mrs_cek("NAME"))
'        listitem.SubItems(4) = IIf(IsNull(mrs_cek("NEXTACTDATE")), "", Format(mrs_cek("NEXTACTDATE"), "yyyy/mm/dd hh:nn"))
'        listitem.SubItems(5) = IIf(IsNull(mrs_cek("NEXTACT")), "", mrs_cek("NEXTACT"))
'        listitem.SubItems(6) = IIf(IsNull(mrs_cek("REMARKS")), "", mrs_cek("REMARKS"))
'        listitem.SubItems(7) = IIf(IsNull(mrs_cek("AGENT")), "", mrs_cek("AGENT"))
'        listitem.SubItems(8) = IIf(IsNull(mrs_cek("NamaAGENT")), "", mrs_cek("NamaAGENT"))
'        listitem.SubItems(9) = IIf(IsNull(mrs_cek("RECSOURCE")), "", mrs_cek("RECSOURCE"))
'        listitem.SubItems(10) = IIf(IsNull(mrs_cek("TGLSTATUS")), "", Format(mrs_cek("TGLSTATUS"), "DD/MM/YYYY"))
'        listitem.SubItems(11) = IIf(IsNull(mrs_cek("Kethslkerja")), "", mrs_cek("Kethslkerja"))
'        listitem.SubItems(12) = IIf(IsNull(mrs_cek("KdComplaint")), "", mrs_cek("KdComplaint"))
'        listitem.SubItems(13) = IIf(IsNull(mrs_cek("RemarkComplaint")), "", mrs_cek("RemarkComplaint"))
'        listitem.SubItems(14) = IIf(IsNull(mrs_cek("F_CEK")), "", mrs_cek("F_CEK"))
'        'LISTITEM.SubItems(15) = IIf(IsNull(m_cari("[NO]")), "", m_cari("[NO]"))
'        mrs_cek.MoveNext
'    Wend
'        If LstVwSearchmgm.ListItems.count = 0 Then
'            TxtJmlDtmgm.Text = "Tidak Ada Data"
'        Else
'            TxtJmlDtmgm.Text = "Total " + CStr(mrs_cek.RecordCount) + " Records"
'        End If
'LstVwSearchmgm.SortKey = 2
'LstVwSearchmgm.Sorted = True
'Exit Sub
'HELL:
'    Me.MousePointer = vbNormal
'    MsgBox Err.Description
'  ''  Resume
'End Sub
'
Private Sub LstVwSearchmgm_KeyUp(KeyCode As Integer, Shift As Integer)

End Sub
Private Sub Timer1_Timer()


'strsql = "select * from usertbl where tgl update between "
End Sub

Public Sub renderdonk()
Dim NAMACUST As String
Dim NamaAgent As String
Dim DATASOURCE As String
Dim TGLLAHIR As String
Dim nmagentprev As String
Dim OFFPHONE As String
Dim OFFPHONE2 As String
Dim HOMEPHONE As String
Dim HOMEPHONE2 As String
Dim MOBILEPHONE As String
Dim MOBILEPHONE2 As String
Dim FAXPHONE As String
Dim Lcustid As String
Dim FAXPHONE2 As String
Dim KETHSLKERJA As String
Dim lLastCallDate As String
Dim lStatusCek As String
Dim sPending As String
Dim FCEKSTATUS As String
Dim strverify As String
Dim strapprovel As String
Dim M_DATA As New CLS_FRMSEARCH
Dim M_Objrs As New ADODB.Recordset
Dim PANJANG As Integer
Dim strReject As String
Dim strSukses As String
Dim strapprovelyet As String
Dim strinject As String
Dim strmarkup As String
Dim BlokedEntry As String
Dim STSLOCKTL As String
Dim STSfromaccount As String
    F_CEK = Empty
    WO_DATE = Empty
    If UCase(Left(MDIForm1.Text2.Text, 5)) = "ADMIN" Or UCase(Left(MDIForm1.Text2.Text, 5)) = "SUPER" Or UCase(Left(MDIForm1.Text2.Text, 5)) = "TEAML" Then
    Else
    Call CEK_STATUS_F_CEK
    End If
'    Call UPDATE_BP
        M_Objrs.CursorLocation = adUseClient
        cmdsql = "SELECT * FROM usertbl WHERE USERID = '" + MDIForm1.Text1.Text + "'"
        M_Objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic

        If Not M_Objrs.EOF Then
           strinject = IIf(IsNull(M_Objrs!lockdarispv), "", M_Objrs!lockdarispv)
           If strinject = "" Then
              Bloked = ""
           Else
            Bloked = IIf(IsNull(M_Objrs!lockdarispv), "", Replace(M_Objrs!lockdarispv, "@", "'"))
            BlokedEntry = IIf(IsNull(M_Objrs!lock_entry_lpd), "", M_Objrs!lock_entry_lpd)
           End If
         strmarkup = IIf(IsNull(M_Objrs!lockmarkup), "", M_Objrs!lockmarkup)
        End If
        
        While Not M_Objrs.EOF
                 StsVl = CStr(Trim(IIf(IsNull(M_Objrs!f_VL), "", M_Objrs!f_VL)))
                StsON = CStr(Trim(IIf(IsNull(M_Objrs!f_ON), "", M_Objrs!f_ON)))
                StsOS = CStr(Trim(IIf(IsNull(M_Objrs!f_OS), "", M_Objrs!f_OS)))
                StsSK = CStr(Trim(IIf(IsNull(M_Objrs!f_SK), "", M_Objrs!f_SK)))
                StsPR = CStr(Trim(IIf(IsNull(M_Objrs!f_PR), "", M_Objrs!f_PR)))
                StsPTP = CStr(Trim(IIf(IsNull(M_Objrs!f_ptp), "", M_Objrs!f_ptp)))
                StsBP = CStr(Trim(IIf(IsNull(M_Objrs!f_bp), "", M_Objrs!f_bp)))
                StsPOP = CStr(Trim(IIf(IsNull(M_Objrs!f_pop), "", M_Objrs!f_pop)))
                StsSP = CStr(Trim(IIf(IsNull(M_Objrs!f_sp), "", M_Objrs!f_sp)))
                StsUC = CStr(Trim(IIf(IsNull(M_Objrs!F_UC), "", M_Objrs!F_UC)))
                StsWO_Date = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_DATE), "", M_Objrs!f_WO_DATE)))
                StsWO_2009 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2009), "", M_Objrs!f_WO_2009)))
                StsWO_2008 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2008), "", M_Objrs!f_WO_2008)))
                StsWO_2007 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2007), "", M_Objrs!f_WO_2007)))
                StsWO_2006 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2006), "", M_Objrs!f_WO_2006)))
                StsWO_2005 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2005), "", M_Objrs!f_WO_2005)))
                StsWO_2004 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2004), "", M_Objrs!f_WO_2004)))
                StsWO_2003 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2003), "", M_Objrs!f_WO_2003)))
                StsWO_2002 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2002), "", M_Objrs!f_WO_2002)))
                StsWO_2001 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2001), "", M_Objrs!f_WO_2001)))
                StsWO_2000 = CStr(Trim(IIf(IsNull(M_Objrs!f_WO_2000), "", M_Objrs!f_WO_2000)))
                StsWO_1999 = CStr(Trim(IIf(IsNull(M_Objrs!F_WO_1999), "", M_Objrs!F_WO_1999)))
                LUserType = CStr(Trim(IIf(IsNull(M_Objrs!usertype), "", M_Objrs!usertype)))
                STSLOCKTL = CStr(Trim(IIf(IsNull(M_Objrs!lockdarispvbuattl), "", M_Objrs!lockdarispvbuattl)))
                STSfromaccount = CStr(Trim(IIf(IsNull(M_Objrs!fromaccount), "", M_Objrs!fromaccount)))
                
                M_Objrs.MoveNext
        Wend
        Set M_Objrs = Nothing
        StsAll = StsVl + StsPR + StsBP + StsPOP + StsUC + StsSK + StsON + StsOS
       
       If STSLOCKTL <> Empty Then
        If Left(Combo1(0).Text, 5) = "LUNAS" Then
                If STSfromaccount = "LUNAS PENDING" Then
                    STSLOCKTL = STSLOCKTL
                ElseIf STSfromaccount = "LUNAS COMPLETE" Then
                      STSLOCKTL = STSLOCKTL
                Else
                     STSLOCKTL = ""
                End If
                
        Else
                STSLOCKTL = ""
        End If
        End If
        
     If StsAll <> "" Then
            If LUserType = "1" Then
                    If StsUC = "UC" Then
                     '       F_CEK = "substring(F_CEK_NEW,1,3)  IN('NK-','MV-','WN-','" + StsNa + "','" + StsOP + "','" + StsPTP + "','" + StsBP + "','" + StsPOP + "','" + StsSP + "','')"
                     '   Else
                            F_CEK = "substring(F_CEK_NEW,1,3)  IN( '" + StsVl + "','" + StsPR + "','" + StsPTP + "','" + StsBP + "','" + StsPOP + "','" + StsSP + "','" + StsSK + "','" + StsON + "','" + StsOS + "','')"
                        End If
                        
                    End If
     End If
        
      If StsWO_Date = "1" Then
            If LUserType = "1" Then
                WO_DATE = "substring(B_D,1,4) in ('" + StsWO_2009 + "','" + StsWO_2008 + "','" + StsWO_2007 + "','" + StsWO_2006 + "','" + StsWO_2005 + "', "
                WO_DATE = WO_DATE + "'" + StsWO_2004 + "', '" + StsWO_2003 + "', '" + StsWO_2002 + "', '" + StsWO_2001 + "','" + StsWO_2000 + "','" + StsWO_1999 + "')"
            End If
      End If
        If Trim(Text1(0).Text) = Empty And Trim(Combo1(0).Text) = Empty And Combo1(2).Text = Empty And Len(TDBMask2.Value) < 1 And Len(TDBMask1.Value) < 1 And TdDob.ValueIsNull And Len(Text1(2).Text) < 3 _
        And cmbStsLastCall(0).Text = "" And CmbStatusCek.Text = "" And DtLastCall(0).ValueIsNull And CekDtDistribute.Value = 0 And Combo3.Text = "" Then
            MsgBox "Masukan Kriteria Customer Yang Akan Dicari...!!!", vbCritical + vbOKOnly, "Peringatan"
            Text1(0).SetFocus
            Set M_DATA = Nothing
            Exit Sub
        Else
        
         LstVwSearchMgm.ListItems.CLEAR
         Frame3.Visible = True
         If CekDtDistribute.Value = 1 Then
            NamaAgent = "AGENT is null"
         Else
            If Text1(2).Text <> Empty Then
                Lcustid = "CUSTID LIKE " + "'%" + UBAH_QUOTE(Text1(2).Text) + "%'"
            Else
                If Text1(0).Text <> Empty Then
                    NAMACUST = "name LIKE " + "'%" + UBAH_QUOTE(Text1(0).Text) + "%'"
                End If
                If Combo1(0).Text <> Empty Then
                    If UCase(Left(MDIForm1.Text1.Text, 2)) = "TL" And Combo1(0).Text = Trim(MDIForm1.Text1.Text) Then
                        NamaAgent = "AGENT in(select userid from usertbl where team='" + Trim(Combo1(0).Text) + "')"
                    ElseIf UCase(Left(MDIForm1.Text1.Text, 4)) = "JOKO  " And Trim(Combo1(0).Text) = Trim(MDIForm1.Text1.Text) Then
                           NamaAgent = "AGENT in(select userid from usertbl where spvcode='" + Trim(Combo1(0).Text) + "')"
                    Else
                        NamaAgent = "AGENT = '" + Trim(Combo1(0).Text) + "'"
                    End If
                
                
                    If UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
                        nmagentprev = "agentprev IN (SELECT USERID FROM USERTBL WHERE TEAM='" + MDIForm1.Text1.Text + "' )"
                ElseIf UCase(MDIForm1.Text2.Text) = "AGENT" Then
                        nmagentprev = "agentprev ='" + MDIForm1.Text1.Text + "' "
                Else
                        nmagentprev = ""
                End If
              End If
                If Combo1(2).Text <> Empty Then
                    DATASOURCE = "RECSOURCE = '" + Trim(Combo1(2).Text) + "'"
                End If
                If TdDob.ValueIsNull Then
                Else
                    TGLLAHIR = "DOB = '" + Format(TdDob.Text, "yyyy/mm/dd") + "'"
                End If
                If Len(TDBMask1.Value) > 1 Then
                    OFFPHONE = "OFFICENO Like '%" + TDBMask1.Value + "%'"
                    OFFPHONE2 = "OFFICENO2 Like '%" + TDBMask1.Value + "%'"
                    HOMEPHONE = "HOMENO Like '%" + TDBMask1.Value + "%'"
                    HOMEPHONE2 = "HOMENO2 Like '%" + TDBMask1.Value + "%'"
                    FAXPHONE = "FAXNO Like '%" + TDBMask1.Value + "%'"
                    FAXPHONE2 = "FAXNO2 Like '%" + TDBMask1.Value + "%'"
                End If
                
                If Len(TDBMask2.Value) > 1 Then
                    MOBILEPHONE = "MOBILENO like '%" + TDBMask2.Value + "%'"
                    MOBILEPHONE2 = "MOBILENO2 like '%" + TDBMask2.Value + "%'"
                End If
                
                
                If Left(Combo3.Text, 1) = 6 Then
                    strverify = "intverify=0 and  stscpa=1 and (resultcpa is null or resultcpa='')"
                End If
                
                If Left(Combo3.Text, 1) = 1 Then
                  strapprovel = " (intapprovel=0 or intapprovel is null )  and intverify=1  and (resultcpa is null or resultcpa='')  "
                End If
                
                If Left(Combo3.Text, 1) = 4 Then
                  strapprovelyet = " (intapprovel=0 or intapprovel is null )  and (intverify=0 or intverify isnull) and stscpa=1 and (resultcpa is null or resultcpa='')  "
                End If
                
                If Left(Combo3.Text, 1) = 2 Then
                  strReject = "   resultcpa ='GAGAL'  "
                End If
                
                If Left(Combo3.Text, 1) = 3 Then
                  strSukses = "   resultcpa ='SUKSES'  "
                End If
                
                
                
                If cmbStsLastCall(0).Text <> "" Then
                    Select Case UCase(Trim(cmbStsLastCall(0).Text))
                        Case "NOT CONTACTED"
                            KETHSLKERJA = "KETHSLKERJA IN('WN','WN','NP','BT')"
                        Case "NOT AVAILABLE"
                            KETHSLKERJA = "KETHSLKERJA = 'NA'"
                        Case "STILL THINKING"
                            KETHSLKERJA = "KETHSLKERJA= 'ST'"
                        Case "DISAGREE"
                            KETHSLKERJA = "LEFT(KETHSLKERJA,1)= 'D'"
                        Case "SENDING APPLICATION"
                            KETHSLKERJA = "KETHSLKERJA= 'SA'"
                        Case "PICKUP"
                            KETHSLKERJA = "KETHSLKERJA= 'PU'"
                        Case "INCOMPLETE DOCUMENT"
                            KETHSLKERJA = "KETHSLKERJA= 'ID'"
                        Case "INCOMING"
                            KETHSLKERJA = "KETHSLKERJA= 'I'"
                        Case "AVAILABLE"
                            KETHSLKERJA = "KETHSLKERJA= '1A'"
                    End Select
                End If
                If Len(Combo2) > 0 Then
                    FCEKSTATUS = " F_CEK_NEW LIKE '" + Combo2.Text + "%'"
                End If
                If CmbStatusCek.Text <> "" Then
                    Select Case UCase(Trim(CmbStatusCek.Text))
                        Case "ACCEPT"
                            lStatusCek = "UPPER(F_cek_NEW) ='ACCEPT'"
                        Case "RETURN"
                            lStatusCek = "UPPER(F_cek_NEW) = 'RETURN'"
                        Case "NOT CHECK"
                            lStatusCek = "(f_cek_NEW is null or f_cek_NEW = 'NOT ACCEPT' OR F_CEK ='') AND KETHSLKERJA ='I'"
                    End Select
                End If
                If DtLastCall(0).ValueIsNull Then
                Else
                    lLastCallDate = "TGLSTATUS BETWEEN '" + Format(DtLastCall(0).Value, "MM/DD/YYYY") & " " & CStr(DTimeLastCall(0).Value) + "' AND '" + Format(DtLastCall(1).Value, "MM/DD/YYYY") & " " & CStr(DTimeLastCall(1).Value) + "'"
                End If
        End If
        End If
          
                'Unload FRM_SEARCH
                If Check1.Value = 0 Then
                    Set m_cari = M_DATA.QUERY_SEARCH_CONDITION(M_OBJCONN, NAMACUST, NamaAgent, DATASOURCE, TGLLAHIR, _
                                                            OFFPHONE, OFFPHONE2, HOMEPHONE, HOMEPHONE2, MOBILEPHONE, _
                                                            MOBILEPHONE2, FAXPHONE, FAXPHONE2, MDIForm1.Text3.Text, Lcustid, F_CEK, lLastCallDate, lStatusCek)
                Else
                   If strmarkup <> "" Then
                    F_CEK = ""
                    WO_DATE = ""
                    BlokedEntry = ""
                    Bloked = ""
                End If
                    Set m_cari = M_DATA.QUERY_SEARCH_CONDITION_mgm(M_OBJCONN, NAMACUST, NamaAgent, DATASOURCE, TGLLAHIR, _
                                                             OFFPHONE, OFFPHONE2, HOMEPHONE, HOMEPHONE2, MOBILEPHONE, _
                                                            MOBILEPHONE2, FAXPHONE, FAXPHONE2, MDIForm1.Text3.Text, Lcustid, F_CEK, lLastCallDate, lStatusCek, sPending, FCEKSTATUS, WO_DATE, strverify, strapprovel, strapprovelyet, strReject, strSukses, Bloked, BlokedEntry, strmarkup, nmagentprev, STSLOCKTL, "", "")
                End If
        
            If m_cari.RecordCount = 0 Then
                MsgBox "Data Tidak Ditemukan", vbInformation + vbOKOnly, "Aplikasi"
                Set M_DATA = Nothing
                Exit Sub
            Else
               
                search_ok = True
                If Check1.Value = 1 Then
                    'kalau found refferall data
                    'Unload FRM_PRESCREEN
                    'FRM_PRESCREEN.Caption = "Search Non mgm Data"
                    'FRM_PRESCREEN.Show
                    SSTab1.Tab = 0
'                    Call show_UCDATA
                    Call show_Search_mgmData
                    
                Else
                    ' kalau found mgm data
                    SSTab1.Tab = 1
'Call show_Search_Refferal
                    
'                    Unload VIEW_mgmDATA
'                    If MDIForm1.Text2.Text = "Agent" Then
'                        VIEW_mgmDATA.Caption = "Search mgm Data"
'                    Else
'                        VIEW_mgmDATA.Caption = "Search mgm Data  .... Tekan Huruf ""P"" untuk Mengalihkan Data"
'                    End If
'
'
'                    VIEW_mgmDATA.Show
                End If
            'FRM_PRESCREEN.Show vbModal
'                Unload Me
            End If
        End If

End Sub

'=================================@@11022013 Buat Nambahin script Cari Data ALL ================
Private Sub CariDataAll()
    Dim harga As Double
    Dim ListItem As ListItem
    Dim Lcustid1 As String
    Dim Lcustid2 As String
    Dim LCall As String
    Dim I, K As Integer
    Dim cmdsql As String
    Dim sPending As String
    Dim M_Objrs As ADODB.Recordset
    Dim VOLUMEAMOUNT As Double
    Dim statusprior As String
    Dim exp%
    Dim totamount As Double
    Dim TOTCURBALANCE As Double
    Dim kdprofile_aksesall As String
    
    Dim M_Objrs_Cek As ADODB.Recordset
    'Dim CMDSQL As String
    
    '@@19022013
    Dim M_ObjrsCekStatus As ADODB.Recordset
    
    'Cek statusnya langsung dari usertbl aja
    ' UPDATE 21 MEI 2013 IZUDDIN
    If UCase(MDIForm1.Text2.Text) = "AGENT" Or _
        UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
        cmdsql = "select f_akses_all_acc,profile_akses_all from usertbl where userid='"
        cmdsql = cmdsql + MDIForm1.Text1.Text + "'"
        Set M_ObjrsCekStatus = New ADODB.Recordset
        M_ObjrsCekStatus.CursorLocation = adUseClient
        M_ObjrsCekStatus.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_ObjrsCekStatus.RecordCount > 0 Then
            AksesAllAcc = IIf(IsNull(M_ObjrsCekStatus("f_akses_all_acc")), "", M_ObjrsCekStatus("f_akses_all_acc"))
            kdprofile_aksesall = IIf(IsNull(M_ObjrsCekStatus("profile_akses_all")), "", M_ObjrsCekStatus("profile_akses_all"))
        End If
        Set M_ObjrsCekStatus = Nothing
    End If
    
    If AksesAllAcc = "1" Then
        cmdsql = "SELECT * FROM mgm WHERE custid in "
        cmdsql = cmdsql + " (select b.custid from tbl_profile_aksesall a, tbl_cust_aksesall b  "
        cmdsql = cmdsql + " where a.kd_profile=b.kd_profile AND a.kd_profile='"
        cmdsql = cmdsql + kdprofile_aksesall + "' "
        cmdsql = cmdsql + " AND a.waktu_akhir > now() AND  "
        cmdsql = cmdsql + " a.waktu_awal <= now()) AND agent='AKSESALL'"
    Else
        ' Balikkin ke agent sebelumnya 03 Juni 2014
        M_OBJCONN.Execute "UPDATE mgm SET agent=agent_asli WHERE agent='AKSESALL' AND agent_asli IS NOT NULL AND custid not in (SELECT custid FROM tbl_cust_aksesall);"
        Exit Sub
    End If
    
    I = 1
    
    Set M_Objrs_Cek = New ADODB.Recordset
    M_Objrs_Cek.CursorLocation = adUseClient
    M_Objrs_Cek.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
        
        
        
    datajml = M_Objrs_Cek.RecordCount
    
    Me.MousePointer = vbHourglass
    
    ProgressBar1.Max = M_Objrs_Cek.RecordCount + 1
    
    
    
    If M_Objrs_Cek.RecordCount > 0 Then
        ' SET FOCUS ACCOUNT 13 MEI 2014 --------------
        LstVwSearchMgm.ListItems.CLEAR
        ' --------------------------------------
        MsgBox "INFO: Anda memiliki tambahan account sebanyak: " & M_Objrs_Cek.RecordCount & "  yang dapat di collect bersama. Lihat baris paling bawah dan bertanda merah!", vbOKOnly + vbInformation, "Informasi"
        While Not M_Objrs_Cek.EOF
            ProgressBar1.Value = M_Objrs_Cek.Bookmark
            Lcustid1 = CStr(IIf(IsNull(M_Objrs_Cek!CustId), "", M_Objrs_Cek!CustId))
            sPending = CStr(Trim(IIf(IsNull(M_Objrs_Cek!f_Pending), "", M_Objrs_Cek!f_Pending)))
            
            Set ListItem = LstVwSearchMgm.ListItems.ADD(, , M_Objrs_Cek.Bookmark)
            
            If MDIForm1.Text2 = "TeamLeader" Then
                If IIf(IsNull(M_Objrs_Cek("stscpa")), "0", M_Objrs_Cek("stscpa")) = 1 Then
                    ListItem.ForeColor = vbRed
                End If
                
                If IIf(IsNull(M_Objrs_Cek("intapprovel")), "0", M_Objrs_Cek("intapprovel")) = 1 Then
                  ListItem.ForeColor = vbBlue
                End If
            End If
            
            
            If UCase(MDIForm1.Text7) = "JOKO" Or UCase(MDIForm1.Text7) = "WULANDARI" Or UCase(MDIForm1.Text7) = "ANDRI" Then
                If IIf(IsNull(M_Objrs_Cek("intverify")), "0", M_Objrs_Cek("intverify")) = 1 Then
                    ListItem.ForeColor = vbYellow
                End If
                
                If IIf(IsNull(M_Objrs_Cek("intapprovel")), "0", M_Objrs_Cek("intapprovel")) = 1 Then
                  ListItem.ForeColor = vbGreen
                End If
            End If
            
            
            statusprior = IIf(IsNull(M_Objrs_Cek("StatusPrior")), "", M_Objrs_Cek("StatusPrior"))
            ListItem.SubItems(1) = IIf(IsNull(M_Objrs_Cek("CUSTID")), "", M_Objrs_Cek("CUSTID"))
            ListItem.SubItems(2) = IIf(IsNull(M_Objrs_Cek("PRIOR")), "", M_Objrs_Cek("PRIOR"))
            ListItem.SubItems(3) = IIf(IsNull(M_Objrs_Cek("NAME")), "", M_Objrs_Cek("NAME"))
            ListItem.SubItems(4) = IIf(IsNull(M_Objrs_Cek("RECSOURCE")), "", M_Objrs_Cek("RECSOURCE"))
            ListItem.SubItems(5) = IIf(IsNull(M_Objrs_Cek("NEXTACTDATE")), "", Format(M_Objrs_Cek("NEXTACTDATE"), "dd/mm/yyyy hh:nn"))
            ListItem.SubItems(6) = IIf(IsNull(M_Objrs_Cek("NEXTACT")), "", M_Objrs_Cek("NEXTACT"))
            ListItem.SubItems(7) = IIf(IsNull(M_Objrs_Cek("REMARKS")), "", M_Objrs_Cek("REMARKS"))
            ListItem.SubItems(8) = CStr(IIf(IsNull(M_Objrs_Cek("kethslkerja_new")), "", M_Objrs_Cek("kethslkerja_new")) & " " & sPending)
            ListItem.SubItems(9) = IIf(IsNull(M_Objrs_Cek("StatusCall")), "", M_Objrs_Cek("StatusCall"))
            ListItem.SubItems(11) = IIf(IsNull(M_Objrs_Cek("AGENT")), "", M_Objrs_Cek("AGENT"))
            
            
            
            If UCase(MDIForm1.Text2) <> "SUPERVISOR" Then
                If Format(IIf(IsNull(M_Objrs_Cek("flaglead")), 0, M_Objrs_Cek("flaglead")), "##,###") = 1 Then
                    ListItem.SubItems(12) = ""
                Else
                    ListItem.SubItems(12) = Format(IIf(IsNull(M_Objrs_Cek("Principal")), 0, M_Objrs_Cek("Principal")), "##,###")
                End If
            Else
                 ListItem.SubItems(12) = Format(IIf(IsNull(M_Objrs_Cek("Principal")), 0, M_Objrs_Cek("Principal")), "##,###")
            End If
            
            
            ListItem.SubItems(13) = Format(IIf(IsNull(M_Objrs_Cek("AmountWo")), 0, M_Objrs_Cek("AmountWo")), "##,###")
            totamount = totamount + IIf(IsNull(M_Objrs_Cek("AmountWo")), 0, M_Objrs_Cek("AmountWo"))
            
            
            ListItem.SubItems(14) = Format(IIf(IsNull(M_Objrs_Cek("OpenDate")), "", M_Objrs_Cek("OpenDate")), "yyyy/mm/dd")
            ListItem.SubItems(15) = Format(IIf(IsNull(M_Objrs_Cek("B_D")), 0, M_Objrs_Cek("B_D")))
            ListItem.SubItems(16) = Format(IIf(IsNull(M_Objrs_Cek("Pay_Dt")), 0, M_Objrs_Cek("Pay_Dt")), "yyyy/mm/dd")
            
            ListItem.SubItems(17) = Format(IIf(IsNull(M_Objrs_Cek("lastpay")), 0, M_Objrs_Cek("lastpay")), "##,###")
            
            ListItem.SubItems(18) = IIf(IsNull(M_Objrs_Cek("TGLSTATUS")), "", Format(M_Objrs_Cek("TGLSTATUS"), "YYYY/MM/DD"))
            ListItem.SubItems(19) = IIf(IsNull(M_Objrs_Cek("TGLCALL")), "", Format(M_Objrs_Cek("TGLCALL"), "YYYY/MM/DD"))
            ListItem.SubItems(20) = IIf(IsNull(M_Objrs_Cek("Kethslkerja")), "", M_Objrs_Cek("Kethslkerja"))
            ListItem.SubItems(21) = Format(IIf(IsNull(M_Objrs_Cek("TGLINCOMING")), "", M_Objrs_Cek("TGLINCOMING")), "YYYY/MM/DD")
            ListItem.SubItems(23) = IIf(IsNull(M_Objrs_Cek("resultcpa")), "", M_Objrs_Cek("resultcpa"))
            ListItem.SubItems(24) = IIf(IsNull(M_Objrs_Cek("tglinsertfrmcpa")), "", M_Objrs_Cek("tglinsertfrmcpa"))
            ListItem.SubItems(25) = Format(IIf(IsNull(M_Objrs_Cek("curbal")), "", M_Objrs_Cek("curbal")), "##,###")
            TOTCURBALANCE = TOTCURBALANCE + IIf(IsNull(M_Objrs_Cek("curbal")), 0, M_Objrs_Cek("curbal"))
           
            '@@16032011 Tambahan DOB dan No KTP
            ListItem.SubItems(26) = IIf(IsNull(M_Objrs_Cek("dob")), "", Format(M_Objrs_Cek("dob"), "yyyy-mm-dd"))
            ListItem.SubItems(27) = IIf(IsNull(M_Objrs_Cek("ktpno")), "", M_Objrs_Cek("ktpno"))
            
                
SorryLompat:
            
            VOLUMEAMOUNT = VOLUMEAMOUNT + IIf(IsNull(M_Objrs_Cek("AmountWo")), 0, M_Objrs_Cek("AmountWo"))
            
            ListItem.ForeColor = vbRed
            For K = 1 To 26
                ListItem.ListSubItems(K).ForeColor = vbRed
            Next K
            M_Objrs_Cek.MoveNext
        Wend
    Else
        '@@18022013 delete datanya di tbl_distribusi_account
'        Cmdsql = "delete from tbl_distribusi_account where agent='"
'        Cmdsql = Cmdsql + MDIForm1.Text1.Text + "' and waktu_akhir < now()"
'        cmdsql = "update mgm set agent=agent_asli,agent_asli=null WHERE monitor_akses is null" & _
'                " AND agent='AKSESALL' AND custid in(SELECT custid FROM tbl_cust_aksesall a,tbl_profile_aksesall b WHERE " & _
'                " a.kd_profile=b.kd_profile AND b.waktu_akhir < now()) AND agent_asli is not null"
        ' UPDATE 30 OKT 2013 - BY IZUDDIN
        cmdsql = "UPDATE mgm SET agent=agent_asli,agent_asli=null WHERE " & _
                " agent='AKSESALL' AND custid in(SELECT custid FROM tbl_cust_aksesall a,tbl_profile_aksesall b WHERE " & _
                " a.kd_profile=b.kd_profile AND b.waktu_akhir < now()) AND agent_asli is not null"

        M_OBJCONN.Execute cmdsql
        
'        cmdsql = "UPDATE mgm SET agent_asli=null WHERE " & _
'                " agent='AKSESALL' AND custid in(SELECT custid FROM tbl_cust_aksesall a,tbl_profile_aksesall b WHERE " & _
'                " a.kd_profile=b.kd_profile AND b.waktu_akhir < now()) AND agent_asli is not null"
'
'        M_OBJCONN.Execute cmdsql
        
        cmdsql = "DELETE FROM tbl_cust_aksesall "
        cmdsql = cmdsql & " WHERE kd_profile in (SELECT kd_profile FROM tbl_profile_aksesall WHERE waktu_akhir < now()) "
        M_OBJCONN.Execute cmdsql
        
        'Cek dulu nih, kalo di tbl_distribusi_account=0 update aja f_akses_all_acc=null
'        Cmdsql = "select * from tbl_distribusi_account where agent='"
'        Cmdsql = Cmdsql + MDIForm1.Text1.Text + "'"
'        Set M_ObjrsCekStatus = New ADODB.Recordset
'        M_ObjrsCekStatus.CursorLocation = adUseClient
'        M_ObjrsCekStatus.Open Cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'        If M_ObjrsCekStatus.RecordCount = 0 Then
            '03062013 - UPDATE BY IZUDDIN
'            If M_Objrs_Cek.state = 1 Then M_Objrs_Cek.Close
'            M_Objrs_Cek.Open "SELECT * FROM tbl_cust_aksesall WHERE kd_profile='" & kdprofile_aksesall & "'"
'            If Not M_Objrs_Cek.RecordCount > 0 Then
                cmdsql = "UPDATE usertbl SET profile_akses_all=null,f_akses_all_acc=null,f_pesanresetauto=null WHERE profile_akses_all in (SELECT kd_profile FROM tbl_profile_aksesall WHERE waktu_akhir < now());"
                'cmdsql = cmdsql + MDIForm1.Text1.Text + "'"
                M_OBJCONN.Execute cmdsql
'            End If
            AksesAllAcc = ""
'        End If
'        Set M_ObjrsCekStatus = Nothing
    End If
        
    MousePointer = vbNormal
    Set M_Objrs_Cek = Nothing
End Sub
