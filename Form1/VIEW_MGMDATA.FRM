VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form VIEW_MGMDATA 
   BackColor       =   &H80000004&
   ClientHeight    =   9825
   ClientLeft      =   255
   ClientTop       =   1425
   ClientWidth     =   16500
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   Icon            =   "VIEW_MGMDATA.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9825
   ScaleWidth      =   16500
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      ForeColor       =   &H00000000&
      Height          =   8055
      Left            =   -15
      TabIndex        =   52
      Top             =   1740
      Width           =   16485
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
         Left            =   10020
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   7740
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
         Left            =   13365
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   7725
         Width           =   3045
      End
      Begin MSComctlLib.ListView LstVwSearchMgm 
         Height          =   7590
         Left            =   0
         TabIndex        =   55
         Top             =   105
         Width           =   16455
         _ExtentX        =   29025
         _ExtentY        =   13388
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   3555
      Left            =   555
      TabIndex        =   45
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
      TabPicture(0)   =   "VIEW_MGMDATA.frx":1CFA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LblTarget(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Check1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Referall Data"
      TabPicture(1)   =   "VIEW_MGMDATA.frx":1D16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "LblTarget(1)"
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
         TabIndex        =   51
         Top             =   1950
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   3225
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H80000004&
         Height          =   8685
         Left            =   -74940
         TabIndex        =   46
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
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   7980
            Width           =   3045
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   7845
            Left            =   0
            TabIndex        =   48
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
         TabIndex        =   50
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
         TabIndex        =   49
         Top             =   -15
         Visible         =   0   'False
         Width           =   9465
      End
   End
   Begin Threed.SSCommand Command1 
      Cancel          =   -1  'True
      Height          =   345
      Index           =   1
      Left            =   15585
      TabIndex        =   56
      Top             =   -15
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   609
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
      Caption         =   "&Close"
      ButtonStyle     =   2
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   390
      Left            =   3675
      TabIndex        =   58
      Top             =   -15
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   688
      _Version        =   196610
      Font3D          =   4
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      Height          =   1770
      Left            =   -15
      TabIndex        =   23
      Top             =   -15
      Width           =   16470
      _ExtentX        =   29051
      _ExtentY        =   3122
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Find"
      TabPicture(0)   =   "VIEW_MGMDATA.frx":1D32
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Schedulle"
      TabPicture(1)   =   "VIEW_MGMDATA.frx":1D4E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Height          =   1425
         Left            =   -74985
         TabIndex        =   38
         Top             =   285
         Width           =   16410
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
            TabIndex        =   18
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
            TabIndex        =   17
            Top             =   150
            Width           =   3480
         End
         Begin Threed.SSFrame SSFrame1 
            Height          =   510
            Left            =   9675
            TabIndex        =   39
            Top             =   120
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
               TabIndex        =   19
               Top             =   60
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               Calendar        =   "VIEW_MGMDATA.frx":1D6A
               Caption         =   "VIEW_MGMDATA.frx":1E82
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "VIEW_MGMDATA.frx":1EEE
               Keys            =   "VIEW_MGMDATA.frx":1F0C
               Spin            =   "VIEW_MGMDATA.frx":1F6A
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
               TabIndex        =   20
               Top             =   90
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               Calendar        =   "VIEW_MGMDATA.frx":1F92
               Caption         =   "VIEW_MGMDATA.frx":20AA
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "VIEW_MGMDATA.frx":2116
               Keys            =   "VIEW_MGMDATA.frx":2134
               Spin            =   "VIEW_MGMDATA.frx":2192
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
               Left            =   1485
               TabIndex        =   40
               Top             =   90
               Visible         =   0   'False
               Width           =   285
            End
         End
         Begin Threed.SSCommand CmdScheduleoK 
            Height          =   390
            Left            =   13650
            TabIndex        =   21
            Top             =   135
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   688
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
            TabIndex        =   43
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
            TabIndex        =   44
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
            TabIndex        =   41
            Top             =   165
            Visible         =   0   'False
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   12
            Left            =   8415
            TabIndex        =   57
            Top             =   210
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   13
            Left            =   180
            TabIndex        =   42
            Top             =   195
            Width           =   1170
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000004&
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
         Height          =   1455
         Left            =   45
         TabIndex        =   24
         Top             =   285
         Width           =   16395
         Begin VB.ComboBox cmbStsLastCall 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   0
            Left            =   4155
            TabIndex        =   7
            Top             =   735
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
            TabIndex        =   22
            Top             =   -345
            Visible         =   0   'False
            Width           =   3225
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00FFFFFF&
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
            Index           =   3
            Left            =   9180
            TabIndex        =   13
            Top             =   105
            Width           =   2580
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00FFFFFF&
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
            Index           =   2
            Left            =   7350
            TabIndex        =   12
            Top             =   105
            Width           =   1800
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00FFFFFF&
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
            Index           =   1
            Left            =   5115
            TabIndex        =   6
            Top             =   435
            Width           =   2310
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00FFFFFF&
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
            Index           =   0
            Left            =   4155
            TabIndex        =   5
            Top             =   435
            Width           =   960
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00FFFFFF&
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
            Index           =   0
            Left            =   4155
            TabIndex        =   4
            Top             =   120
            Width           =   2025
         End
         Begin TDBDate6Ctl.TDBDate TdDob 
            Height          =   315
            Left            =   1200
            TabIndex        =   1
            Top             =   1440
            Visible         =   0   'False
            Width           =   1170
            _Version        =   65536
            _ExtentX        =   2064
            _ExtentY        =   556
            Calendar        =   "VIEW_MGMDATA.frx":21BA
            Caption         =   "VIEW_MGMDATA.frx":22D2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "VIEW_MGMDATA.frx":233E
            Keys            =   "VIEW_MGMDATA.frx":235C
            Spin            =   "VIEW_MGMDATA.frx":23BA
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
            Top             =   750
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   556
            Caption         =   "VIEW_MGMDATA.frx":23E2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "VIEW_MGMDATA.frx":244E
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
            Left            =   1200
            TabIndex        =   3
            Top             =   1065
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   556
            Caption         =   "VIEW_MGMDATA.frx":2490
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "VIEW_MGMDATA.frx":24FC
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
         Begin Threed.SSCommand Command1 
            Height          =   360
            Index           =   0
            Left            =   8925
            TabIndex        =   15
            Top             =   1005
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   635
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
         Begin Threed.SSCommand Command1 
            Height          =   360
            Index           =   2
            Left            =   9870
            TabIndex        =   16
            Top             =   1005
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   635
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
            Caption         =   "&Clear"
            ButtonStyle     =   2
         End
         Begin TDBTime6Ctl.TDBTime DTimeLastCall 
            Height          =   300
            Index           =   0
            Left            =   5430
            TabIndex        =   9
            Top             =   1050
            Width           =   825
            _Version        =   65536
            _ExtentX        =   1455
            _ExtentY        =   529
            Caption         =   "VIEW_MGMDATA.frx":253E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "VIEW_MGMDATA.frx":25AA
            Spin            =   "VIEW_MGMDATA.frx":25FA
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
            Left            =   4155
            TabIndex        =   8
            Top             =   1050
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   556
            Calendar        =   "VIEW_MGMDATA.frx":2622
            Caption         =   "VIEW_MGMDATA.frx":273A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "VIEW_MGMDATA.frx":27A6
            Keys            =   "VIEW_MGMDATA.frx":27C4
            Spin            =   "VIEW_MGMDATA.frx":2822
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
            Left            =   6555
            TabIndex        =   10
            Top             =   1050
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   556
            Calendar        =   "VIEW_MGMDATA.frx":284A
            Caption         =   "VIEW_MGMDATA.frx":2962
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "VIEW_MGMDATA.frx":29CE
            Keys            =   "VIEW_MGMDATA.frx":29EC
            Spin            =   "VIEW_MGMDATA.frx":2A4A
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
            Left            =   7830
            TabIndex        =   11
            Top             =   1050
            Width           =   825
            _Version        =   65536
            _ExtentX        =   1455
            _ExtentY        =   529
            Caption         =   "VIEW_MGMDATA.frx":2A72
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "VIEW_MGMDATA.frx":2ADE
            Spin            =   "VIEW_MGMDATA.frx":2B2E
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
            Height          =   495
            Left            =   9045
            TabIndex        =   25
            Top             =   1500
            Visible         =   0   'False
            Width           =   2025
            Begin MSComctlLib.ProgressBar ProgressBar1 
               Height          =   270
               Left            =   15
               TabIndex        =   26
               Top             =   165
               Width           =   1980
               _ExtentX        =   3493
               _ExtentY        =   476
               _Version        =   393216
               Appearance      =   0
            End
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Status Telp."
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
            Height          =   225
            Index           =   8
            Left            =   2790
            TabIndex        =   37
            Top             =   810
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tgl Terakhir Telp."
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
            Height          =   210
            Index           =   9
            Left            =   2775
            TabIndex        =   36
            Top             =   1095
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
            TabIndex        =   35
            Top             =   2310
            Width           =   1365
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "To"
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
            Height          =   225
            Index           =   11
            Left            =   6285
            TabIndex        =   34
            Top             =   1080
            Width           =   225
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "# Kartu :"
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
            Index           =   5
            Left            =   0
            TabIndex        =   33
            Top             =   135
            Width           =   1125
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "HP :"
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
            Height          =   225
            Index           =   6
            Left            =   75
            TabIndex        =   32
            Top             =   1140
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Telp Rumah :"
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
            Index           =   4
            Left            =   60
            TabIndex        =   31
            Top             =   795
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Batch"
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
            Height          =   255
            Index           =   3
            Left            =   6750
            TabIndex        =   30
            Top             =   150
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
            Left            =   75
            TabIndex        =   29
            Top             =   1440
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TeleCollection"
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
            Height          =   210
            Index           =   1
            Left            =   2790
            TabIndex        =   28
            Top             =   495
            Width           =   1365
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nama"
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
            Index           =   0
            Left            =   2700
            TabIndex        =   27
            Top             =   150
            Width           =   1395
         End
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
Dim StsNa  As String
Dim StsOP As String
Dim StsPTP As String
Dim StsBP As String
Dim StsPOP As String
Dim StsSP As String
Dim StsUC As String
Dim StsAll As String
Dim StsRP As String
Dim cmdsql As String
Dim LUserType As String
Dim F_CEK As String
Dim f_Pending As String
Dim datajml As Integer

Private Sub HEADER_VIEW_Refferall()
    listview1.ColumnHeaders.ADD 1, , "No", 3 * TXT
    listview1.ColumnHeaders.ADD 2, , "Cust Id", 5 * TXT
    listview1.ColumnHeaders.ADD 3, , "Priority", 1
    listview1.ColumnHeaders.ADD 4, , "Ref Id", 10 * TXT
    listview1.ColumnHeaders.ADD 5, , "Ref Name", 10 * TXT
    listview1.ColumnHeaders.ADD 6, , "Nama Customer", 25 * TXT
    'ListView1.ColumnHeaders.ADD 7, , "Batch Expire Date", 25 * TXT
    listview1.ColumnHeaders.ADD 7, , "Tgl Schedule", 10 * TXT
    listview1.ColumnHeaders.ADD 8, , "Next Action", 12 * TXT
    listview1.ColumnHeaders.ADD 9, , "Remarks", 17 * TXT
    listview1.ColumnHeaders.ADD 10, , "Sts LastCall", 17 * TXT
    listview1.ColumnHeaders.ADD 11, , "SalesCode", 8 * TXT
    listview1.ColumnHeaders.ADD 12, , "Agent", 8 * TXT
    listview1.ColumnHeaders.ADD 13, , "DataBase", 10 * TXT
    listview1.ColumnHeaders.ADD 14, , "LastCall Date", 10 * TXT
    listview1.ColumnHeaders.ADD 15, , "Code", 5 * TXT
    listview1.ColumnHeaders.ADD 16, , "Complaint Note", 15 * TXT
    listview1.ColumnHeaders.ADD 17, , "Check", 10 * TXT
    listview1.ColumnHeaders.ADD 18, , "ID", 10 * TXT
End Sub
Private Sub isi_dataClaimKeGrid(gCUSTID As String, gNama As String, gnextact As String, gremarks As String, gagent As String, gnamaagent As String, grecsource As String)
    ' insert ke grid list view
Dim listitem As listitem
Set listitem = LstVwSearchMgm.ListItems.ADD(, , "9999")
    listitem.SubItems(1) = gCUSTID
    listitem.SubItems(2) = ""
    listitem.SubItems(3) = gNama
    listitem.SubItems(4) = Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd hh:nn")
    listitem.SubItems(5) = gnextact
    listitem.SubItems(6) = gremarks
    listitem.SubItems(7) = gagent
    listitem.SubItems(8) = gnamaagent
    listitem.SubItems(9) = grecsource
    listitem.SubItems(10) = ""
    listitem.SubItems(11) = "1A"
    listitem.SubItems(12) = ""
    listitem.SubItems(13) = ""
    listitem.SubItems(14) = ""
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
Dim listitem As listitem
Dim VOLUMEAMOUNT As Boolean


If Check2.Value = 1 Then

    LstVwSearchMgm.ListItems.Clear
    SSTab1.Tab = 0
    ' searching schedule mgm
  Call CEK_STATUS_F_CEK
  
   Set m_cari = M_DATA.QUERY_SEARCH_mgm(M_OBJCONN, "AGENT = '" + Combo1(4).Text + "' AND (NEXTACTDATE BETWEEN '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " 00:00" + "' AND '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " 23:59" + "') ", MDIForm1.Text3.Text, F_CEK, f_Pending)
    ProgressBar1.Max = m_cari.RecordCount + 1
    If Check2.Value = 1 Then
        TxtJmlDtMgm.Text = m_cari.RecordCount & " Data"
    Else
        Text2.Text = m_cari.RecordCount & " Data"
    End If
    
    While Not m_cari.EOF
    ProgressBar1.Value = m_cari.Bookmark
    
    Set listitem = LstVwSearchMgm.ListItems.ADD(, , m_cari.Bookmark)
        listitem.SubItems(1) = IIf(IsNull(m_cari("CUSTID")), "", m_cari("CUSTID"))
        listitem.SubItems(2) = IIf(IsNull(m_cari("PRIOR")), "", m_cari("PRIOR"))
        listitem.SubItems(3) = IIf(IsNull(m_cari("NAME")), "", m_cari("NAME"))
        listitem.SubItems(4) = IIf(IsNull(m_cari("TGLSOURCE")), "", DateAdd("d", 90, Format(m_cari("TGLSOURCE"), "dd-mm-yyyy")))
        listitem.SubItems(5) = IIf(IsNull(m_cari("NEXTACTDATE")), "", Format(m_cari("NEXTACTDATE"), "yyyy/mm/dd hh:nn"))
        listitem.SubItems(6) = IIf(IsNull(m_cari("NEXTACT")), "", m_cari("NEXTACT"))
        listitem.SubItems(7) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS")) & "-" & IIf(IsNull(m_cari("f_pending")), "", m_cari("f_pending"))
        listitem.SubItems(8) = IIf(IsNull(m_cari("F_cek")), "", m_cari("F_cek"))
        listitem.SubItems(9) = IIf(IsNull(m_cari("StatusCall")), "", m_cari("StatusCall"))
        listitem.SubItems(11) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
        'listitem.SubItems(10) = IIf(IsNull(m_cari("NamaAGENT")), "", m_cari("NamaAGENT"))
        listitem.SubItems(12) = Format(IIf(IsNull(m_cari("Principal")), 0, m_cari("Principal")), "##,###")
        listitem.SubItems(13) = Format(IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo")), "##,###")
        listitem.SubItems(14) = Format(IIf(IsNull(m_cari("OpenDate")), "", m_cari("OpenDate")), "yyyy/mm/dd")
        listitem.SubItems(15) = Format(IIf(IsNull(m_cari("TtlPTP")), 0, m_cari("TtlPTP")), "##,###")
        listitem.SubItems(16) = IIf(IsNull(m_cari("RECSOURCE")), "", m_cari("RECSOURCE"))
        listitem.SubItems(17) = IIf(IsNull(m_cari("TGLSTATUS")), "", Format(m_cari("TGLSTATUS"), "YYYY/MM/DD"))
        listitem.SubItems(18) = IIf(IsNull(m_cari("TGLcall")), "", Format(m_cari("TGLcall"), "YYYY/MM/DD"))
        listitem.SubItems(19) = IIf(IsNull(m_cari("Kethslkerja")), "", m_cari("Kethslkerja"))
        listitem.SubItems(20) = Format(IIf(IsNull(m_cari("TGLINCOMING")), "", m_cari("TGLINCOMING")), "YYYY/MM/DD")
        'listitem.SubItems(18) = IIf(IsNull(m_cari("RemarkComplaint")), "", m_cari("RemarkComplaint"))
        VOLUMEAMOUNT = VOLUMEAMOUNT + IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo"))
        'LISTITEM.SubItems(15) = IIf(IsNull(m_cari("[NO]")), "", m_cari("[NO]"))
        m_cari.MoveNext
    Wend
        If LstVwSearchMgm.ListItems.Count = 0 Then
            TxtJmlDtMgm.Text = "Tidak Ada Data"
            TxtJmlVolMgm.Text = "0"
        Else
            TxtJmlDtMgm.Text = "Total " + CStr(m_cari.RecordCount) + " Records"
            TxtJmlVolMgm.Text = "Total " + CStr(m_cari.RecordCount)
        End If
        
Else
    ' searching schedule leads
    Set m_cari = M_DATA.QUERY_SEARCH(M_OBJCONN, "AGENT = '" + Combo1(4).Text + "' AND (NEXTACTDATE BETWEEN '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " 00:00" + "' AND '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " 23:59" + "') ", MDIForm1.Text3.Text)
        listview1.ListItems.Clear
        SSTab1.Tab = 1
        ' searching schedule mgm
        ProgressBar1.Max = m_cari.RecordCount + 1
        Text2.Text = m_cari.RecordCount & " Data"
        While Not m_cari.EOF
        ProgressBar1.Value = m_cari.Bookmark
        Set listitem = listview1.ListItems.ADD(, , m_cari.Bookmark)
        listitem.SubItems(1) = IIf(IsNull(m_cari("custid")), "", JADI_QUOTE(m_cari("custid")))
        Select Case m_cari("RECSTATUS")
        Case "1A"
            listitem.SubItems(2) = "Available"
        Case ""
            listitem.SubItems(2) = "Available"
        Case Else
            listitem.SubItems(2) = IIf(IsNull(m_cari("PRIOR")), "", m_cari("PRIOR"))
        End Select
        listitem.SubItems(3) = IIf(IsNull(m_cari("CUSTIDREF")), "", m_cari("CUSTIDREF"))
        listitem.SubItems(4) = IIf(IsNull(m_cari("NAMAREF")), "", m_cari("NAMAREF"))
        listitem.SubItems(5) = IIf(IsNull(m_cari("NAME")), "", JADI_QUOTE(m_cari("NAME")))
        listitem.SubItems(6) = IIf(IsNull(m_cari("NEXTACTDATE")), "", Format(m_cari("NEXTACTDATE"), "yyyy/mm/dd hh:mm"))
        listitem.SubItems(7) = IIf(IsNull(m_cari("NEXTACT")), "", m_cari("NEXTACT"))
        listitem.SubItems(8) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS"))
        listitem.SubItems(9) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
        listitem.SubItems(10) = IIf(IsNull(m_cari("NamaAGENT")), "", m_cari("NamaAGENT"))
        listitem.SubItems(11) = IIf(IsNull(m_cari("RECSOURCEREF")), "", m_cari("RECSOURCEREF"))
        listitem.SubItems(12) = Format(IIf(IsNull(m_cari("TGLSTATUS")), "", m_cari("TGLSTATUS")), "yyyy/mm/dd")
        listitem.SubItems(13) = IIf(IsNull(m_cari("Kethslkerja")), "", m_cari("Kethslkerja"))
        listitem.SubItems(14) = IIf(IsNull(m_cari("KdComplaint")), "", m_cari("KdComplaint"))
        listitem.SubItems(15) = IIf(IsNull(m_cari("RemarkComplaint")), "", m_cari("RemarkComplaint"))
        listitem.SubItems(16) = IIf(IsNull(m_cari("F_CEK")), "", m_cari("F_CEK"))
        m_cari.MoveNext
    Wend
End If
Set M_DATA = Nothing

End Sub

Private Sub CmdMissed_Click()
Dim M_DATA As New CLS_FRMSEARCH
Dim listitem As listitem
Dim VOLUMEAMOUNT As Double
If Check2.Value = 1 Then
    Call CEK_STATUS_F_CEK
    Set m_cari = M_DATA.QUERY_SEARCH_mgm(M_OBJCONN, "AGENT = '" + Combo1(4).Text + "'  AND (NEXTACTDATE < '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " 00:00" + "') ", MDIForm1.Text3.Text, F_CEK, f_Pending)
       LstVwSearchMgm.ListItems.Clear
        SSTab1.Tab = 0
        ' searching schedule mgm
       ProgressBar1.Max = m_cari.RecordCount + 1
'        If Check2.Value = 1 Then
'            TxtJmlDtMgm.Text = m_cari.RecordCount & " Data"
'        Else
'            Text2.Text = m_cari.RecordCount & " Data"
'        End If
        While Not m_cari.EOF
            ProgressBar1.Value = m_cari.Bookmark
            
    Set listitem = LstVwSearchMgm.ListItems.ADD(, , m_cari.Bookmark)
        listitem.SubItems(1) = IIf(IsNull(m_cari("CUSTID")), "", m_cari("CUSTID"))
        listitem.SubItems(3) = IIf(IsNull(m_cari("NAME")), "", m_cari("NAME"))
        listitem.SubItems(4) = IIf(IsNull(m_cari("NEXTACTDATE")), "", Format(m_cari("NEXTACTDATE"), "yyyy/mm/dd hh:nn"))
        listitem.SubItems(5) = IIf(IsNull(m_cari("NEXTACT")), "", m_cari("NEXTACT"))
        listitem.SubItems(6) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS"))
        'listitem.SubItems(7) = IIf(IsNull(m_cari("F_cek")), "", m_cari("F_cek"))
        listitem.SubItems(8) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
        'listitem.SubItems(9) = IIf(IsNull(m_cari("NamaAGENT")), "", m_cari("NamaAGENT"))
        listitem.SubItems(10) = Format(IIf(IsNull(m_cari("Principal")), 0, m_cari("Principal")), "##,###")
        listitem.SubItems(11) = Format(IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo")), "##,###")
        listitem.SubItems(12) = Format(IIf(IsNull(m_cari("OpenDate")), "", m_cari("OpenDate")), "yyyy/mm/dd")
        listitem.SubItems(13) = Format(IIf(IsNull(m_cari("TtlPTP")), 0, m_cari("TtlPTP")), "##,###")
        listitem.SubItems(14) = IIf(IsNull(m_cari("RECSOURCE")), "", m_cari("RECSOURCE"))
        listitem.SubItems(15) = IIf(IsNull(m_cari("TGLSTATUS")), "", Format(m_cari("TGLSTATUS"), "YYYY/MM/DD"))
        listitem.SubItems(16) = IIf(IsNull(m_cari("Kethslkerja")), "", m_cari("Kethslkerja"))
        listitem.SubItems(17) = Format(IIf(IsNull(m_cari("TGLINCOMING")), "", m_cari("TGLINCOMING")), "YYYY/MM/DD")
        VOLUMEAMOUNT = VOLUMEAMOUNT + IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo"))
        m_cari.MoveNext
    Wend
        If LstVwSearchMgm.ListItems.Count = 0 Then
            TxtJmlDtMgm.Text = "Tidak Ada Data"
            TxtJmlVolMgm.Text = "0"
        Else
            TxtJmlDtMgm.Text = "Total " + CStr(m_cari.RecordCount) + " Records"
            TxtJmlVolMgm.Text = "Total " + CStr(m_cari.RecordCount)
        End If
            
Else
    Set m_cari = M_DATA.QUERY_SEARCH(M_OBJCONN, "AGENT = '" + Combo1(4).Text + "'  AND (NEXTACTDATE < '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " 00:00" + "') ", MDIForm1.Text3.Text)
        listview1.ListItems.Clear
        SSTab1.Tab = 1
        ' searching schedule mgm
        ProgressBar1.Max = m_cari.RecordCount + 1
        Text2.Text = m_cari.RecordCount & " Data"
        While Not m_cari.EOF
        ProgressBar1.Value = m_cari.Bookmark
        Set listitem = listview1.ListItems.ADD(, , m_cari.Bookmark)
        listitem.SubItems(1) = IIf(IsNull(m_cari("custid")), "", JADI_QUOTE(m_cari("custid")))
        Select Case m_cari("RECSTATUS")
        Case "1A"
            listitem.SubItems(2) = "Available"
        Case ""
            listitem.SubItems(2) = "Available"
        Case Else
            listitem.SubItems(2) = IIf(IsNull(m_cari("PRIOR")), "", m_cari("PRIOR"))
        End Select
        listitem.SubItems(3) = IIf(IsNull(m_cari("CUSTIDREF")), "", m_cari("CUSTIDREF"))
        listitem.SubItems(4) = IIf(IsNull(m_cari("NAMAREF")), "", m_cari("NAMAREF"))
        listitem.SubItems(5) = IIf(IsNull(m_cari("NAME")), "", JADI_QUOTE(m_cari("NAME")))
        listitem.SubItems(6) = IIf(IsNull(m_cari("NEXTACTDATE")), "", Format(m_cari("NEXTACTDATE"), "yyyy/mm/dd hh:mm"))
        listitem.SubItems(7) = IIf(IsNull(m_cari("NEXTACT")), "", m_cari("NEXTACT"))
        listitem.SubItems(8) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS"))
        listitem.SubItems(9) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
        listitem.SubItems(10) = IIf(IsNull(m_cari("NamaAGENT")), "", m_cari("NamaAGENT"))
        listitem.SubItems(11) = IIf(IsNull(m_cari("RECSOURCEREF")), "", m_cari("RECSOURCEREF"))
        listitem.SubItems(12) = Format(IIf(IsNull(m_cari("TGLSTATUS")), "", m_cari("TGLSTATUS")), "yyyy/mm/dd")
        listitem.SubItems(13) = IIf(IsNull(m_cari("Kethslkerja")), "", m_cari("Kethslkerja"))
        listitem.SubItems(14) = IIf(IsNull(m_cari("KdComplaint")), "", m_cari("KdComplaint"))
        listitem.SubItems(15) = IIf(IsNull(m_cari("RemarkComplaint")), "", m_cari("RemarkComplaint"))
        listitem.SubItems(16) = IIf(IsNull(m_cari("F_CEK")), "", m_cari("F_CEK"))
        m_cari.MoveNext
    Wend
End If
End Sub
Private Sub CEK_STATUS_F_CEK()
Dim cmdsql As String
Dim M_OBJRS As New ADODB.Recordset

F_CEK = Empty
Set M_OBJRS = New ADODB.Recordset
        M_OBJRS.CursorLocation = adUseClient
        cmdsql = "SELECT F_NA, F_OP, F_PTP, F_BP, F_POP, F_SP, F_UC,F_RP,usertype FROM usertbl WHERE USERID = '" + MDIForm1.Text1.Text + "'"
         M_OBJRS.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        
        
        While Not M_OBJRS.EOF
                StsNa = CStr(Trim(IIf(IsNull(M_OBJRS!f_na), "", M_OBJRS!f_na)))
                StsOP = CStr(Trim(IIf(IsNull(M_OBJRS!f_op), "", M_OBJRS!f_op)))
                StsPTP = CStr(Trim(IIf(IsNull(M_OBJRS!f_ptp), "", M_OBJRS!f_ptp)))
                StsBP = CStr(Trim(IIf(IsNull(M_OBJRS!f_bp), Empty, M_OBJRS!f_bp)))
                StsPOP = CStr(Trim(IIf(IsNull(M_OBJRS!f_pop), "", M_OBJRS!f_pop)))
                StsSP = CStr(Trim(IIf(IsNull(M_OBJRS!f_sp), "", M_OBJRS!f_sp)))
                StsRP = CStr(Trim(IIf(IsNull(M_OBJRS!f_rp), "", M_OBJRS!f_rp)))
                StsUC = CStr(Trim(IIf(IsNull(M_OBJRS!f_uc), "", M_OBJRS!f_uc)))
                LUserType = CStr(Trim(IIf(IsNull(M_OBJRS!usertype), "", M_OBJRS!usertype)))
                M_OBJRS.MoveNext
            Wend
            Set M_OBJRS = Nothing
             StsAll = StsNa + StsOP + StsPTP + StsBP + StsPOP + StsSP + StsRP + StsUC
            
     If StsAll <> Empty Then
            If LUserType = "1" Then
                    If StsUC <> Empty Then
                        F_CEK = "(left(F_CEK,3)IN('NK-','MV-','WN-','NBP','" + StsNa + "','" + StsOP + "','" + StsPTP + "','" + StsBP + "','" + StsPOP + "','" + StsSP + "','" + StsRP + "') or F_CEK IS NULL)"
                        Else
                        F_CEK = "(left(F_CEK,3)IN( '" + StsNa + "','" + StsOP + "','" + StsPTP + "','" + StsBP + "','" + StsPOP + "','" + StsSP + "','" + StsRP + "') or F_CEK IS NULL) "
                    End If
            End If
     End If

End Sub
Private Sub CmdToday_Click()
Dim M_DATA As New CLS_FRMSEARCH
Dim listitem As listitem
Dim cmdsql As String

Dim VOLUMEAMOUNT As Double
If Check2.Value = 1 Then
    Call CEK_STATUS_F_CEK
    Set m_cari = M_DATA.QUERY_SEARCH_mgm(M_OBJCONN, "AGENT = '" + Combo1(4).Text + "'  AND (NEXTACTDATE BETWEEN '" + Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") & " 00:00" + "' AND '" + Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") & " 23:59" + "') ", MDIForm1.Text3.Text, F_CEK, f_Pending)
       LstVwSearchMgm.ListItems.Clear
        SSTab1.Tab = 0
        ' searching schedule mgm
       ProgressBar1.Max = m_cari.RecordCount + 1
'        If Check2.Value = 1 Then
'            TxtJmlDtMgm.Text = m_cari.RecordCount & " Data"
'        Else
'            Text2.Text = m_cari.RecordCount & " Data"
'        End If
        While Not m_cari.EOF
            ProgressBar1.Value = m_cari.Bookmark
            
    Set listitem = LstVwSearchMgm.ListItems.ADD(, , m_cari.Bookmark)
       listitem.SubItems(1) = IIf(IsNull(m_cari("CUSTID")), "", m_cari("CUSTID"))
        listitem.SubItems(2) = IIf(IsNull(m_cari("PRIOR")), "", m_cari("PRIOR"))
        listitem.SubItems(3) = IIf(IsNull(m_cari("NAME")), "", m_cari("NAME"))
        listitem.SubItems(4) = IIf(IsNull(m_cari("NEXTACTDATE")), "", Format(m_cari("NEXTACTDATE"), "yyyy/mm/dd hh:nn"))
        listitem.SubItems(5) = IIf(IsNull(m_cari("NEXTACT")), "", m_cari("NEXTACT"))
        listitem.SubItems(6) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS"))
        listitem.SubItems(7) = IIf(IsNull(m_cari("F_cek")), "", m_cari("F_cek"))
        listitem.SubItems(9) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
        listitem.SubItems(10) = Format(IIf(IsNull(m_cari("Principal")), 0, m_cari("Principal")), "##,###")
        listitem.SubItems(11) = Format(IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo")), "##,###")
        listitem.SubItems(12) = Format(IIf(IsNull(m_cari("OpenDate")), "", m_cari("OpenDate")), "yyyy/mm/dd")
        listitem.SubItems(13) = Format(IIf(IsNull(m_cari("TtlPTP")), 0, m_cari("TtlPTP")), "##,###")
        listitem.SubItems(14) = IIf(IsNull(m_cari("RECSOURCE")), "", m_cari("RECSOURCE"))
        listitem.SubItems(15) = IIf(IsNull(m_cari("TGLSTATUS")), "", Format(m_cari("TGLSTATUS"), "YYYY/MM/DD"))
        listitem.SubItems(16) = IIf(IsNull(m_cari("Kethslkerja")), "", m_cari("Kethslkerja"))
        listitem.SubItems(17) = Format(IIf(IsNull(m_cari("TGLINCOMING")), "", m_cari("TGLINCOMING")), "YYYY/MM/DD")
        VOLUMEAMOUNT = VOLUMEAMOUNT + IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo"))
        m_cari.MoveNext
    Wend
        If LstVwSearchMgm.ListItems.Count = 0 Then
            TxtJmlDtMgm.Text = "Tidak Ada Data"
            TxtJmlVolMgm.Text = "0"
        Else
            TxtJmlDtMgm.Text = "Total " + CStr(m_cari.RecordCount) + " Records"
            TxtJmlVolMgm.Text = "Total " + CStr(m_cari.RecordCount)
        End If
            
Else
    Set m_cari = M_DATA.QUERY_SEARCH(M_OBJCONN, "AGENT = '" + Combo1(4).Text + "'  AND (NEXTACTDATE BETWEEN '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " 00:00" + "' AND '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " 23:59" + "') ", MDIForm1.Text3.Text)
        listview1.ListItems.Clear
        SSTab1.Tab = 1
        ' searching schedule mgm
        ProgressBar1.Max = m_cari.RecordCount + 1
        Text2.Text = m_cari.RecordCount & " Data"
        While Not m_cari.EOF
        ProgressBar1.Value = m_cari.Bookmark
        Set listitem = listview1.ListItems.ADD(, , m_cari.Bookmark)
        listitem.SubItems(1) = IIf(IsNull(m_cari("custid")), "", JADI_QUOTE(m_cari("custid")))
        Select Case m_cari("RECSTATUS")
        Case "1A"
            listitem.SubItems(2) = "Available"
        Case ""
            listitem.SubItems(2) = "Available"
        Case Else
            listitem.SubItems(2) = IIf(IsNull(m_cari("PRIOR")), "", m_cari("PRIOR"))
        End Select
        listitem.SubItems(3) = IIf(IsNull(m_cari("CUSTIDREF")), "", m_cari("CUSTIDREF"))
        listitem.SubItems(4) = IIf(IsNull(m_cari("NAMAREF")), "", m_cari("NAMAREF"))
        listitem.SubItems(5) = IIf(IsNull(m_cari("NAME")), "", JADI_QUOTE(m_cari("NAME")))
        listitem.SubItems(6) = IIf(IsNull(m_cari("NEXTACTDATE")), "", Format(m_cari("NEXTACTDATE"), "yyyy/mm/dd hh:mm"))
        listitem.SubItems(7) = IIf(IsNull(m_cari("NEXTACT")), "", m_cari("NEXTACT"))
        listitem.SubItems(8) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS"))
        listitem.SubItems(9) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
        listitem.SubItems(10) = IIf(IsNull(m_cari("NamaAGENT")), "", m_cari("NamaAGENT"))
        listitem.SubItems(11) = IIf(IsNull(m_cari("RECSOURCEREF")), "", m_cari("RECSOURCEREF"))
        listitem.SubItems(12) = Format(IIf(IsNull(m_cari("TGLSTATUS")), "", m_cari("TGLSTATUS")), "yyyy/mm/dd")
        listitem.SubItems(13) = IIf(IsNull(m_cari("Kethslkerja")), "", m_cari("Kethslkerja"))
        listitem.SubItems(14) = IIf(IsNull(m_cari("KdComplaint")), "", m_cari("KdComplaint"))
        listitem.SubItems(15) = IIf(IsNull(m_cari("RemarkComplaint")), "", m_cari("RemarkComplaint"))
        listitem.SubItems(16) = IIf(IsNull(m_cari("F_CEK")), "", m_cari("F_CEK"))
        m_cari.MoveNext
    Wend
    
End If
End Sub

Private Sub Combo1_Click(Index As Integer)
Dim M_DATA As New CLS_FRMSEARCH
Dim M_OBJRS As ADODB.Recordset
Select Case Index
Case 0
    Set M_OBJRS = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "USERID = '" + Combo1(Index).Text + "'")
    If M_OBJRS.RecordCount <> 0 Then
        Combo1(0).Text = M_OBJRS("USERID")
        Combo1(1).Text = M_OBJRS("AGENT")
    Else
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
    End If
Case 1
    Set M_OBJRS = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "AGENT = '" + Combo1(Index).Text + "'")
    If M_OBJRS.RecordCount <> 0 Then
        Combo1(0).Text = M_OBJRS("USERID")
        Combo1(1).Text = M_OBJRS("AGENT")
    Else
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
    End If
Case 2
Set M_OBJRS = M_DATA.QUERY_DATASOURCE(M_OBJCONN, "KODEDS = '" + Combo1(Index).Text + "'")
    If M_OBJRS.RecordCount <> 0 Then
        Combo1(2).Text = M_OBJRS("KODEDS")
        Combo1(3).Text = M_OBJRS("KETERANGAN")
    Else
        Combo1(2).Text = Empty
        Combo1(3).Text = Empty
    End If
Case 3
Set M_OBJRS = M_DATA.QUERY_DATASOURCE(M_OBJCONN, "KETERANGAN = '" + Combo1(Index).Text + "'")
    If M_OBJRS.RecordCount <> 0 Then
        Combo1(2).Text = M_OBJRS("KODEDS")
        Combo1(3).Text = M_OBJRS("KETERANGAN")
    Else
        Combo1(2).Text = Empty
        Combo1(3).Text = Empty
    End If
Case 4
    Set M_OBJRS = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "USERID = '" + Combo1(Index).Text + "'")
    If M_OBJRS.RecordCount <> 0 Then
        Combo1(4).Text = M_OBJRS("USERID")
        Combo1(5).Text = M_OBJRS("AGENT")
    Else
        Combo1(4).Text = Empty
        Combo1(5).Text = Empty
    End If
Case 5
    Set M_OBJRS = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "AGENT = '" + Combo1(Index).Text + "'")
    If M_OBJRS.RecordCount <> 0 Then
        Combo1(4).Text = M_OBJRS("USERID")
        Combo1(5).Text = M_OBJRS("AGENT")
    Else
        Combo1(4).Text = Empty
        Combo1(5).Text = Empty
    End If
End Select
Set M_DATA = Nothing
Set M_OBJRS = Nothing
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
   lReturn = SendMessage(Combo1(Index).hWnd, CB_FINDSTRING, -1, ByVal sSearchText)
   If lReturn <> CB_ERR Then
      mbIgnoreListClick = True
      Combo1(Index).ListIndex = lReturn
      mbIgnoreListClick = False
      Combo1(Index).Text = Combo1(Index).LIST(lReturn)
      Combo1(Index).SelStart = Len(sSearchText)
      Combo1(Index).SelLength = Len(Combo1(Index).Text)
      KeyAscii = 0
   End If
End If
End Select
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim M_DATA As New CLS_FRMSEARCH
Dim M_OBJRS As ADODB.Recordset
Select Case Index
Case 0
    Set M_OBJRS = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "USERID = '" + Combo1(Index).Text + "'")
    If M_OBJRS.RecordCount <> 0 Then
        Combo1(0).Text = M_OBJRS("USERID")
        Combo1(1).Text = M_OBJRS("AGENT")
    Else
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
    End If
Case 1
    Set M_OBJRS = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "AGENT = '" + Combo1(Index).Text + "'")
    If M_OBJRS.RecordCount <> 0 Then
        Combo1(0).Text = M_OBJRS("USERID")
        Combo1(1).Text = M_OBJRS("AGENT")
    Else
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
    End If
Case 2
Set M_OBJRS = M_DATA.QUERY_DATASOURCE(M_OBJCONN, "KODEDS = '" + Combo1(Index).Text + "'")
    If M_OBJRS.RecordCount <> 0 Then
        Combo1(2).Text = M_OBJRS("KODEDS")
        Combo1(3).Text = M_OBJRS("KETERANGAN")
    Else
        Combo1(2).Text = Empty
        Combo1(3).Text = Empty
    End If
Case 3
Set M_OBJRS = M_DATA.QUERY_DATASOURCE(M_OBJCONN, "KETERANGAN = '" + Combo1(Index).Text + "'")
    If M_OBJRS.RecordCount <> 0 Then
        Combo1(2).Text = M_OBJRS("KODEDS")
        Combo1(3).Text = M_OBJRS("KETERANGAN")
    Else
        Combo1(2).Text = Empty
        Combo1(3).Text = Empty
    End If
End Select
Set M_DATA = Nothing
Set M_OBJRS = Nothing
End Sub

Private Sub Form_Load()
Dim M_OBJRS As ADODB.Recordset
Dim M_DATA As New CLS_FRMSEARCH

Call HEADER_VIEW_MGM
Call HEADER_VIEW_Refferall


If UCase(MDIForm1.Text2.Text) = "AGENT" Then
    CekDtDistribute.Visible = False
    Combo1(4).Text = MDIForm1.Text1.Text
    Combo1(5).Text = MDIForm1.Text7.Text
    Combo1(4).Visible = False
    Combo1(5).Visible = False
    label1(13).Visible = False
    Combo1(0).Text = MDIForm1.Text1.Text
    Combo1(1).Text = MDIForm1.Text7.Text
Else
    CekDtDistribute.Visible = True
End If
    CmbStatusCek.AddItem "ACCEPT"
    CmbStatusCek.AddItem "RETURN"
    CmbStatusCek.AddItem "NOT CHECK"
    SSTab1.Tab = 0
    Sstab2.Tab = 0
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
StsMgmSchedule = False


Set M_OBJRS = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "")
    While Not M_OBJRS.EOF
        Combo1(0).AddItem M_OBJRS("USERID")
        Combo1(1).AddItem M_OBJRS("AGENT")
        Combo1(4).AddItem M_OBJRS("USERID")
        Combo1(5).AddItem M_OBJRS("AGENT")
        M_OBJRS.MoveNext
    Wend
 Set M_OBJRS = Nothing
  
 
 Set M_OBJRS = M_DATA.QUERY_DATASOURCE(M_OBJCONN, "")
    While Not M_OBJRS.EOF
        Combo1(2).AddItem M_OBJRS("KODEDS")
        Combo1(3).AddItem M_OBJRS("KETERANGAN")
        M_OBJRS.MoveNext
    Wend
If UCase(MDIForm1.Text3.Text) = "ADMIN" Then
    label1(5).Visible = True
    Text1(2).Visible = True
End If
M_OBJRS.Close
Set M_OBJRS = Nothing
Set M_DATA = Nothing
End Sub
Private Sub show_Search_mgmData()
Dim listitem As listitem
Dim Lcustid1 As String
Dim Lcustid2 As String
Dim LCall As String
Dim i As Integer
Dim cmdsql As String
Dim sPending As String
Dim M_OBJRS As ADODB.Recordset
Dim VOLUMEAMOUNT As Double
Dim statusprior As String
Dim exp%
i = 1
On Error GoTo HELL
        
    datajml = m_cari.RecordCount
    LstVwSearchMgm.ListItems.Clear
    Me.MousePointer = vbHourglass
    ProgressBar1.Max = m_cari.RecordCount + 1
    While Not m_cari.EOF
    ProgressBar1.Value = m_cari.Bookmark
        Lcustid1 = CStr(IIf(IsNull(m_cari!CustId), "", m_cari!CustId))
        sPending = CStr(Trim(IIf(IsNull(m_cari!f_Pending), "", m_cari!f_Pending)))
        
        Set listitem = LstVwSearchMgm.ListItems.ADD(, , m_cari.Bookmark)
        statusprior = IIf(IsNull(m_cari("StatusPrior")), "", m_cari("StatusPrior"))
        listitem.SubItems(1) = IIf(IsNull(m_cari("CUSTID")), "", m_cari("CUSTID"))
        listitem.SubItems(2) = IIf(IsNull(m_cari("PRIOR")), "", m_cari("PRIOR"))
        listitem.SubItems(3) = IIf(IsNull(m_cari("NAME")), "", m_cari("NAME"))
        listitem.SubItems(4) = IIf(IsNull(m_cari("RECSOURCE")), "", m_cari("RECSOURCE"))
        listitem.SubItems(5) = IIf(IsNull(m_cari("NEXTACTDATE")), "", Format(m_cari("NEXTACTDATE"), "dd/mm/yyyy hh:nn"))
        listitem.SubItems(6) = IIf(IsNull(m_cari("NEXTACT")), "", m_cari("NEXTACT"))
        listitem.SubItems(7) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS"))
        listitem.SubItems(8) = CStr(IIf(IsNull(m_cari("F_cek")), "", m_cari("F_cek")) & " " & sPending)
        listitem.SubItems(9) = IIf(IsNull(m_cari("StatusCall")), "", m_cari("StatusCall"))
        listitem.SubItems(11) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
        listitem.SubItems(12) = Format(IIf(IsNull(m_cari("Principal")), 0, m_cari("Principal")), "##,###")
        listitem.SubItems(13) = Format(IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo")), "##,###")
        listitem.SubItems(14) = Format(IIf(IsNull(m_cari("OpenDate")), "", m_cari("OpenDate")), "yyyy/mm/dd")
        listitem.SubItems(15) = Format(IIf(IsNull(m_cari("TtlPTP")), 0, m_cari("TtlPTP")), "##,###")
        listitem.SubItems(16) = IIf(IsNull(m_cari("RECSOURCE")), "", m_cari("RECSOURCE"))
        listitem.SubItems(17) = IIf(IsNull(m_cari("TGLSTATUS")), "", Format(m_cari("TGLSTATUS"), "YYYY/MM/DD"))
        listitem.SubItems(18) = IIf(IsNull(m_cari("TGLCALL")), "", Format(m_cari("TGLCALL"), "YYYY/MM/DD"))
        listitem.SubItems(19) = IIf(IsNull(m_cari("Kethslkerja")), "", m_cari("Kethslkerja"))
        listitem.SubItems(20) = Format(IIf(IsNull(m_cari("TGLINCOMING")), "", m_cari("TGLINCOMING")), "YYYY/MM/DD")
               
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
'                    CMDSQL = "SELECT count(custid) as callInit from mgm_hst where custid ='" + LstVwSearchMgm.SelectedItem.SubItems(1) + "'  group by custid "
'                    m_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
'                    While Not m_objrs.EOF
'                    'Set listitem = LstVwSearchMgm.ListItems.ADD(, , m_cari.Bookmark)
'                        listitem.SubItems(18) = IIf(IsNull(m_objrs("callinit")), "", m_objrs("callinit"))
'                        m_objrs.MoveNext
'                        Wend
'                        'm_objrs.Close
'                        Set m_objrs = Nothing
        m_cari.MoveNext
    Wend
   
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
HELL:
    Me.MousePointer = vbNormal
    MsgBox Err.Description
End Sub

Private Sub HEADER_VIEW_MGM()
    LstVwSearchMgm.ColumnHeaders.ADD 1, , "No", 3 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 2, , "Cust No", 5 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 3, , "Priority", 0
    LstVwSearchMgm.ColumnHeaders.ADD 4, , "Nama Customer", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 5, , "Batch", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 6, , "Schedule", 5 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 7, , "Next Action", 5 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 8, , "History", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 9, , "Status Account", 17 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 10, , "Status Telp", 17 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 11, , "Call Initial", 5 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 12, , "TeleColl", 5 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 13, , "Principle", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 14, , "Total Amount", 10 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 15, , "Join Date", 5 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 16, , "PTP Amount", 5 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 17, , "DataBase", 5 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 18, , "LastStatus Date", 5 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 19, , "LastCall Date", 5 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 20, , "Sts Account", 5 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 21, , "PTP Date", 5 * TXT
    LstVwSearchMgm.ColumnHeaders.ADD 22, , "Complaint Note", 0
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
Dim NAMAAGENT As String
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


Dim M_DATA As New CLS_FRMSEARCH
Dim M_OBJRS As New ADODB.Recordset
Dim PANJANG As Integer

Select Case Index
    Case 0
    F_CEK = Empty
    If UCase(Left(MDIForm1.Text2.Text, 5)) = "ADMIN" Or UCase(Left(MDIForm1.Text2.Text, 5)) = "SUPER" Or UCase(Left(MDIForm1.Text2.Text, 5)) = "TEAML" Then
    Else
    Call CEK_STATUS_F_CEK
    End If
'    Call UPDATE_BP
        M_OBJRS.CursorLocation = adUseClient
        cmdsql = "SELECT F_NA, F_OP, F_PTP, F_BP, F_POP, F_SP, F_UC,usertype FROM usertbl WHERE USERID = '" + MDIForm1.Text1.Text + "'"
        M_OBJRS.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic

        While Not M_OBJRS.EOF
                StsNa = CStr(Trim(IIf(IsNull(M_OBJRS!f_na), "", M_OBJRS!f_na)))
                StsOP = CStr(Trim(IIf(IsNull(M_OBJRS!f_op), "", M_OBJRS!f_op)))
                StsPTP = CStr(Trim(IIf(IsNull(M_OBJRS!f_ptp), "", M_OBJRS!f_ptp)))
                StsBP = CStr(Trim(IIf(IsNull(M_OBJRS!f_bp), "", M_OBJRS!f_bp)))
                StsPOP = CStr(Trim(IIf(IsNull(M_OBJRS!f_pop), "", M_OBJRS!f_pop)))
                StsSP = CStr(Trim(IIf(IsNull(M_OBJRS!f_sp), "", M_OBJRS!f_sp)))
                StsUC = CStr(Trim(IIf(IsNull(M_OBJRS!f_uc), "", M_OBJRS!f_uc)))
                LUserType = CStr(Trim(IIf(IsNull(M_OBJRS!usertype), "", M_OBJRS!usertype)))
                M_OBJRS.MoveNext
        Wend
        Set M_OBJRS = Nothing
        StsAll = StsNa + StsOP + StsBP + StsPOP + StsUC

     If StsAll <> "" Then
            If LUserType = "1" Then
                    If StsUC = "UC" Then
                        F_CEK = "left(F_CEK,3)  IN('NK-','MV-','WN-','" + StsNa + "','" + StsOP + "','" + StsPTP + "','" + StsBP + "','" + StsPOP + "','" + StsSP + "')"
                        Else
                        F_CEK = "left(F_CEK,3)  IN( '" + StsNa + "','" + StsOP + "','" + StsPTP + "','" + StsBP + "','" + StsPOP + "','" + StsSP + "')"
                    End If
            End If
     End If

        If Trim(Text1(0).Text) = Empty And Trim(Combo1(0).Text) = Empty And Combo1(2).Text = Empty And Len(TDBMask2.Value) < 1 And Len(TDBMask1.Value) < 1 And TdDob.ValueIsNull And Len(Text1(2).Text) < 3 _
        And cmbStsLastCall(0).Text = "" And CmbStatusCek.Text = "" And DtLastCall(0).ValueIsNull And CekDtDistribute.Value = 0 Then
            MsgBox "Masukan Kriteria Customer Yang Akan Dicari...!!!", vbCritical + vbOKOnly, "Peringatan"
            Text1(0).SetFocus
            Set M_DATA = Nothing
            Exit Sub
        Else
        
         LstVwSearchMgm.ListItems.Clear
         Frame3.Visible = True
         If CekDtDistribute.Value = 1 Then
            NAMAAGENT = "AGENT is null"
         Else
            If Text1(2).Text <> Empty Then
                Lcustid = "CUSTID LIKE " + "'%" + UBAH_QUOTE(Text1(2).Text) + "%'"
            Else
                If Text1(0).Text <> Empty Then
                    NAMACUST = "NAME LIKE " + "'%" + UBAH_QUOTE(Text1(0).Text) + "%'"
                End If
                If Combo1(0).Text <> Empty Then
'                    If UCase(Left(MDIForm1.Text1.Text, 2)) = "TL" Then
                        NAMAAGENT = "AGENT = '" + Trim(Combo1(0).Text) + "'"
'                    Else
'                        NAMAAGENT = "AGENT = '" + Trim(Combo1(0).Text) + "'"
'                    End If
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
            
                If CmbStatusCek.Text <> "" Then
                    Select Case UCase(Trim(CmbStatusCek.Text))
                        Case "ACCEPT"
                            lStatusCek = "UPPER(F_cek) ='ACCEPT'"
                        Case "RETURN"
                            lStatusCek = "UPPER(F_cek) = 'RETURN'"
                        Case "NOT CHECK"
                            lStatusCek = "(f_cek is null or f_cek = 'NOT ACCEPT' OR F_CEK ='') AND KETHSLKERJA ='I'"
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
                    Set m_cari = M_DATA.QUERY_SEARCH_CONDITION(M_OBJCONN, NAMACUST, NAMAAGENT, DATASOURCE, TGLLAHIR, _
                                                            OFFPHONE, OFFPHONE2, HOMEPHONE, HOMEPHONE2, MOBILEPHONE, _
                                                            MOBILEPHONE2, FAXPHONE, FAXPHONE2, MDIForm1.Text3.Text, Lcustid, F_CEK, lLastCallDate, lStatusCek)
                Else
                    Set m_cari = M_DATA.QUERY_SEARCH_CONDITION_mgm(M_OBJCONN, NAMACUST, NAMAAGENT, DATASOURCE, TGLLAHIR, _
                                                             OFFPHONE, OFFPHONE2, HOMEPHONE, HOMEPHONE2, MOBILEPHONE, _
                                                            MOBILEPHONE2, FAXPHONE, FAXPHONE2, MDIForm1.Text3.Text, Lcustid, F_CEK, lLastCallDate, lStatusCek, sPending)
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
                    'FRM_PRESCREEN.Caption = "Search Non Mgm Data"
                    'FRM_PRESCREEN.Show
                    SSTab1.Tab = 0
'                    Call show_UCDATA
                    Call show_Search_mgmData
                    
                Else
                    ' kalau found mgm data
                    SSTab1.Tab = 1
                    'Call show_Search_Refferal
                    
'                    Unload VIEW_MGMDATA
'                    If MDIForm1.Text2.Text = "Agent" Then
'                        VIEW_MGMDATA.Caption = "Search MGM Data"
'                    Else
'                        VIEW_MGMDATA.Caption = "Search MGM Data  .... Tekan Huruf ""P"" untuk Mengalihkan Data"
'                    End If
'
'
'                    VIEW_MGMDATA.Show
                End If
            'FRM_PRESCREEN.Show vbModal
'                Unload Me
            End If
        End If
    Case 1
        MDIForm1.LstGrade.ListItems.Clear
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
'CMDSQL = "UPDATE MGM SET KETHSLKErJA='BP-BROKEN PROMISE',F_CEK='BP-',KETHSLKErJADESC='' "
'CMDSQL = CMDSQL + "WHERE CUSTID IN (SELECT CUSTID FrOM MGM WHErE DATEDIFF(day,TdbdatePTP,TGLKOMP)>5)"

End Sub

Private Sub show_UCDATA()
Dim sdata As ADODB.Recordset
Dim listitem As listitem
Dim i%
Set sdata = New ADODB.Recordset
sdata.CursorLocation = adUseClient
cmdsql = "select CUSTID,KETHSLKERJA,AGENT FROM MGM WHERE left(KETHSLKERJA,2) in ('WN','NK','MV') AND AGENT='" & VIEW_MGMDATA.Combo1(0).Text & "'"
sdata.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
i = 1
MDIForm1.LstGrade.ListItems.Clear
Do While Not sdata.EOF
    Set listitem = MDIForm1.LstGrade.ListItems.ADD(, , i)
        listitem.SubItems(1) = IIf(IsNull(sdata!CustId), "", sdata!CustId)
        listitem.SubItems(2) = IIf(IsNull(sdata!CustId), "", sdata!KETHSLKERJA)
        listitem.SubItems(3) = IIf(IsNull(sdata!agent), "", sdata!agent)
        sdata.MoveNext
        i = i + 1
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
   listview1.SortKey = ColumnHeader.Index - 1
   listview1.Sorted = True
End Sub

Private Sub ListView1_DblClick()
Dim M_OBJRS As ADODB.Recordset
If listview1.ListItems.Count = 0 Then
    Exit Sub
End If
Status_Form = 1
    If UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
        Set M_OBJRS = New ADODB.Recordset
        M_OBJRS.CursorLocation = adUseClient
        M_OBJRS.Open "SELECT USERID FROM USERTBL WHERE SPVCODE ='" + MDIForm1.Text1.Text + "' AND USERID = '" + listview1.SelectedItem.SubItems(9) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_OBJRS.RecordCount <> 0 Then
        Else
            MsgBox "Data Ini Milik Agent Team Leader Yang Lain", vbCritical + vbOKOnly, "Aplikasi"
            Set M_OBJRS = Nothing
            Exit Sub
        End If
        Set M_OBJRS = Nothing
    Else
        If UCase(MDIForm1.Text2.Text) = "SUPERVISOR" Or UCase(MDIForm1.Text2.Text) = "ADMINISTRATOR" Then
        Else
            If Trim(UCase(MDIForm1.Text1.Text)) = Trim(UCase(listview1.SelectedItem.SubItems(9))) Then
            Else
                MsgBox "Data Ini Milik Agent Yang Lain", vbCritical + vbOKOnly, "Aplikasi"
                Set M_OBJRS = Nothing
                Exit Sub
            End If
        End If
    End If
    frmCC_Colection.Show vbModal
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
Dim M_OBJRS As ADODB.Recordset
Dim m_msgbox As Variant
If listview1.ListItems.Count = 0 Then
    Exit Sub
End If
If KeyAscii = 13 Then
    Call ListView1_DblClick
End If
If UCase(MDIForm1.Text2.Text) <> "AGENT" Then
    If KeyAscii = 100 Or KeyAscii = 68 Then
        If UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
            Set M_OBJRS = New ADODB.Recordset
            M_OBJRS.CursorLocation = adUseClient
            M_OBJRS.Open "SELECT USERID FROM USERTBL WHERE SPVCODE ='" + MDIForm1.Text1.Text + "' AND USERID = '" + listview1.SelectedItem.SubItems(9) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If M_OBJRS.RecordCount <> 0 Then
                m_msgbox = MsgBox("Yakin Akan Di Hapus??...", vbExclamation + vbYesNo, "Informasi")
                If m_msgbox = vbYes Then
                    M_OBJCONN.Execute "Delete From cc_custtbl where Nomor = " + listview1.SelectedItem.SubItems(17) + " And RECSOURCE <> 'infomedia1' and left(recsource,3)<>'inf'"
                    listview1.ListItems.Remove listview1.SelectedItem.Index
                End If
            Else
                MsgBox "Data Ini Milik Agent Team Leader Yang Lain", vbCritical + vbOKOnly, "Aplikasi"
                Set M_OBJRS = Nothing
                Exit Sub
            End If
            Set M_OBJRS = Nothing
        Else
            m_msgbox = MsgBox("Yakin Akan Di Hapus??...", vbExclamation + vbYesNo, "Informasi")
            If m_msgbox = vbYes Then
                'M_OBJCONN.Execute "Delete From cc_custtbl where custid ='" + ListView1.SelectedItem.SubItems(1) + "'"
                M_OBJCONN.Execute "Delete From cc_custtbl where Nomor = " + listview1.SelectedItem.SubItems(17) + " And RECSOURCE <> 'infomedia1'"
                listview1.ListItems.Remove listview1.SelectedItem.Index
            End If
        End If
    End If
End If
End Sub

Private Sub LstVwSearchMgm_Click()
If b_pindah = True Then
    FrmCustIdTransfer.List1.AddItem LstVwSearchMgm.SelectedItem.SubItems(1)
End If
End Sub

Private Sub LstVwSearchMgm_ColumnClick(ByVal ColumnHeader As ColumnHeader)
   LstVwSearchMgm.SortKey = ColumnHeader.Index - 1
   LstVwSearchMgm.Sorted = True
End Sub

Private Sub LstVwSearchMgm_DblClick()

If LstVwSearchMgm.ListItems.Count = 0 Then
    Exit Sub
End If

Status_Form = 2
Dim M_OBJRS As ADODB.Recordset
If LstVwSearchMgm.ListItems.Count = 0 Then
    Exit Sub
End If
    If UCase(MDIForm1.Text2.Text) = "AGENT" Then
        If UCase(MDIForm1.Text1.Text) <> Trim(UCase(LstVwSearchMgm.SelectedItem.SubItems(11))) Then
            MsgBox "Anda Tidak Berhak Untuk Mengedit Data Ini", vbCritical + vbOKOnly, "Aplikasi"
            Exit Sub
        End If
    End If
    If UCase(MDIForm1.Text2.Text) = "SUPERVISOR" Then
    'If UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
        Dim PO_AGENT As String
        If VIEW_MGMDATA.Combo1(0).Text = "PULLOUT" Then
            Set M_OBJRS = New ADODB.Recordset
            M_OBJRS.CursorLocation = adUseClient
            cmdsql = "SELECT PO_Agent FROM mgm where CUSTID='" & LstVwSearchMgm.SelectedItem.SubItems(11) & "'"
            M_OBJRS.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            PO_AGENT = M_OBJRS!PO_AGENT
            Set M_OBJRS = Nothing
        Else
            PO_AGENT = VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11)
        End If
    
        Set M_OBJRS = New ADODB.Recordset
        M_OBJRS.CursorLocation = adUseClient
        cmdsql = "SELECT USERID FROM USERTBL WHERE TEAM ='" + MDIForm1.Text1.Text + "' AND USERID = '" + PO_AGENT + "'"
        M_OBJRS.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_OBJRS.RecordCount <> 0 Then
        Else
            MsgBox "Data Ini Milik Agent Team Leader Yang Lain", vbCritical + vbOKOnly, "Aplikasi"
            Set M_OBJRS = Nothing
            Exit Sub
        End If
        Set M_OBJRS = Nothing
    End If
    Me.MousePointer = vbHourglass
    Flag_Mgm = False
        frmCC_Colection.Show
    Me.MousePointer = vbNormal
End Sub

Private Sub LstVwSearchMgm_KeyPress(KeyAscii As Integer)
Dim M_OBJRS As ADODB.Recordset
If KeyAscii = 13 Then
    Call LstVwSearchMgm_DblClick
    Exit Sub
End If
If UCase(MDIForm1.Text2.Text) <> "AGENT" Then
    If KeyAscii = 112 Or KeyAscii = 80 Then
        If UCase(MDIForm1.Text2.Text) = "SUPERVISOR" Then
        'If UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
            Set M_OBJRS = New ADODB.Recordset
            M_OBJRS.CursorLocation = adUseClient
            M_OBJRS.Open "SELECT USERID FROM USERTBL WHERE TEAM ='" + MDIForm1.Text1.Text + "' AND USERID = '" + LstVwSearchMgm.SelectedItem.SubItems(11) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If M_OBJRS.RecordCount <> 0 Then
                With View_AlihData
                    .Show vbModal
                    If .ok Then
                        LstVwSearchMgm.ListItems.Remove LstVwSearchMgm.SelectedItem.Index
                    End If
                End With
                Unload View_AlihData
            Else
                MsgBox "Data Ini Milik Agent Team Leader Yang Lain", vbCritical + vbOKOnly, "Aplikasi"
                Set M_OBJRS = Nothing
                Exit Sub
            End If
            Set M_OBJRS = Nothing
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
        'If UCase(MDIForm1.Text2.Text) = "SUPERVISOR" Or UCase(Left(MDIForm1.Text2.Text, 5)) = "ADMIN" Then
                b_pindah = True
                FrmCustIdTransfer.Show
            Else
                If UCase(MDIForm1.Text2.Text) = "SUPERVISOR" Then
                'If UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
                    Set M_OBJRS = New ADODB.Recordset
                    M_OBJRS.CursorLocation = adUseClient
                    M_OBJRS.Open "SELECT USERID FROM USERTBL WHERE TEAM ='" + MDIForm1.Text1.Text + "' AND USERID = '" + LstVwSearchMgm.SelectedItem.SubItems(11) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                    If M_OBJRS.RecordCount <> 0 Then
                        b_pindah = True
                        FrmCustIdTransfer.Show
                    Else
                        MsgBox "Data Ini Milik Agent Team Leader Yang Lain", vbCritical + vbOKOnly, "Aplikasi"
                        Set M_OBJRS = Nothing
                        Exit Sub
                    End If
                    Set M_OBJRS = Nothing
                End If
         End If
    End If
    If KeyAscii = 88 Or KeyAscii = 120 Then
        Dim n%
        If UCase(Left(MDIForm1.Text2.Text, 5)) = "ADMIN" Then
                n = 1
                Do While n <= LstVwSearchMgm.ListItems.Count
                    If LstVwSearchMgm.ListItems(n).Checked = True Then
                        Frmlock.List1.AddItem LstVwSearchMgm.ListItems(n).SubItems(1)
        '                Set ls2 = Frmlock.LSTACCESS.ListItems.ADD(, , LstVwSearchMgm.ListItems(N))
        '                ls2.SubItems(1) = LstVwSearchMgm.SelectedItem
                    End If
                    n = n + 1
                Loop
                Frmlock.Show
        End If
        
        If UCase(MDIForm1.Text2.Text) = "SUPERVISOR" And Combo1(2).Text = "" Then
            Set M_OBJRS = New ADODB.Recordset
            M_OBJRS.CursorLocation = adUseClient
            M_OBJRS.Open "SELECT USERID FROM USERTBL WHERE TEAM ='" + MDIForm1.Text1.Text + "' AND USERID = '" + LstVwSearchMgm.SelectedItem.SubItems(11) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If M_OBJRS.RecordCount <> 0 Then
                'MsgBox "jkasja"
                n = 1
                Do While n <= LstVwSearchMgm.ListItems.Count
                    If LstVwSearchMgm.ListItems(n).Checked = True Then
                        Frmlock.List1.AddItem LstVwSearchMgm.ListItems(n).SubItems(1)
        '                Set ls2 = Frmlock.LSTACCESS.ListItems.ADD(, , LstVwSearchMgm.ListItems(N))
        '                ls2.SubItems(1) = LstVwSearchMgm.SelectedItem
                    End If
                    n = n + 1
                Loop
                Frmlock.Show
            Else
                MsgBox "Data Ini Milik Agent Team Leader Yang Lain", vbCritical + vbOKOnly, "Aplikasi"
                Set M_OBJRS = Nothing
                Exit Sub
            End If
            Set M_OBJRS = Nothing
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
'    cmdsql = "Select  * from MGM where LEFT(recsource,3) <>'PRE' AND HOMENO like '%" + txtnotelprumah.Text + "%' "
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
'            Lcustid = "MGMI-" & CUSTNOMOR(M_OBJCONN, "FRMCUST_CC")
'
'            cmdsql = cmdsql + "('" + Lcustid + "',"
'            cmdsql = cmdsql + " '" + TxtNama.Text + "',"
'            cmdsql = cmdsql + " '" + Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd hh:nn") + "',"
'            cmdsql = cmdsql + " '" + txtnotelprumah.Text + "',"
'            cmdsql = cmdsql + " '" + TxtNoTelpKantor.Text + "',"
'            cmdsql = cmdsql + " '" + TxtHandPhone.Text + "',"
'            cmdsql = cmdsql + " '" + MDIForm1.Text1.Text + "',"
'            cmdsql = cmdsql + " '" + MDIForm1.Text7.Text + "',"
'            cmdsql = cmdsql + " 'MGM_INC',"
'            cmdsql = cmdsql + " 'Data Inbound Call',"
'            cmdsql = cmdsql + " '1A')"
'    M_OBJCONN.Execute cmdsql
'    MsgBox "Sukses..", vbInformation, "Informasi"
'
'    Call isi_dataClaimKeGrid(CStr(Lcustid), TxtNama.Text, "Data Inbound Call", "Data Inbound Call", MDIForm1.Text1.Text, MDIForm1.Text7.Text, "MGM_INC")
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
'        CMDSQL = "Select * from MGM where name like '%" + TxtNama.Text + "%' "
'        CMDSQL = CMDSQL + " and (HOMENO like '%" + kriteria2 + "%' or HOMENO2 like '%" + kriteria2 + "%' or mobileno like '%" + kriteria2 + "%' or mobileno2 like '%" + kriteria2 + "%' or officeno like '%" + kriteria2 + "%' or officeno2 like '%" + kriteria2 + "%') "
'
'    Set mrs_cek = New ADODB.Recordset
'        mrs_cek.CursorLocation = adUseClient
'
'        mrs_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'        If mrs_cek.RecordCount <> 0 Then
'            CUSTID1 = Empty
'            While Not mrs_cek.EOF
'                CUSTID1 = "MGMI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC"))
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
'        CMDSQL = "Select * from MGM where name like '%" + TxtNama.Text + "%' "
'        CMDSQL = CMDSQL + " and (HOMENO like '%" + kriteria2 + "%' or HOMENO2 like '%" + kriteria2 + "%' or mobileno like '%" + kriteria2 + "%' or mobileno2 like '%" + kriteria2 + "%' or officeno like '%" + kriteria2 + "%' or officeno2 like '%" + kriteria2 + "%') "
'    Set mrs_cek = New ADODB.Recordset
'    mrs_cek.CursorLocation = adUseClient
'
'        mrs_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'        If mrs_cek.RecordCount <> 0 Then
'            CUSTID1 = Empty
'            While Not mrs_cek.EOF
'                CUSTID1 = "MGMI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC"))
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
'        CMDSQL = "Select * from MGM where name like '%" + TxtNama.Text + "%' "
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
'                CUSTID1 = "MGMI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC"))
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
'        CMDSQL = "Select * from MGM where name like '%" + TxtNama.Text + "%' "
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
'                CUSTID1 = "MGMI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC"))
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
''        custid1 = "MGMI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC"))
''
''        cmdsql = "Insert into MGM(CUSTID, NAME, HOMENO, MOBILENO, OFFICENO, AGENT, RECSOURCE,CustIdRef,"
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
'        ' munculin tuh form buat entry MGM
'        With FrmEntryCH
'            .TxtIdReff.Text = "Inbound MGM"
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
'    LstVwSearchMgm.ListItems.Clear
'
'    While Not mrs_cek.EOF
'        Set listitem = LstVwSearchMgm.ListItems.ADD(, , mrs_cek.Bookmark)
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
'        If LstVwSearchMgm.ListItems.count = 0 Then
'            TxtJmlDtMgm.Text = "Tidak Ada Data"
'        Else
'            TxtJmlDtMgm.Text = "Total " + CStr(mrs_cek.RecordCount) + " Records"
'        End If
'LstVwSearchMgm.SortKey = 2
'LstVwSearchMgm.Sorted = True
'Exit Sub
'HELL:
'    Me.MousePointer = vbNormal
'    MsgBox Err.Description
'  ''  Resume
'End Sub
'
Private Sub LstVwSearchMgm_KeyUp(KeyCode As Integer, Shift As Integer)

End Sub

