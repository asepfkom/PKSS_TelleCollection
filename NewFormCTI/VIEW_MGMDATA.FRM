VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Begin VB.Form VIEW_MGMDATAold 
   BackColor       =   &H80000004&
   Caption         =   "Data"
   ClientHeight    =   9435
   ClientLeft      =   255
   ClientTop       =   1815
   ClientWidth     =   11760
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   Icon            =   "VIEW_MGMDATA.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9435
   ScaleWidth      =   11760
   WindowState     =   2  'Maximized
   Begin Threed.SSCommand Command1 
      Cancel          =   -1  'True
      Height          =   345
      Index           =   1
      Left            =   10830
      TabIndex        =   24
      Top             =   2250
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
   Begin TabDlg.SSTab SSTab2 
      Height          =   2295
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   4048
      _Version        =   393216
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
      TabCaption(0)   =   "Searching..."
      TabPicture(0)   =   "VIEW_MGMDATA.frx":1CFA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Schedulle"
      TabPicture(1)   =   "VIEW_MGMDATA.frx":1D16
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFC0C0&
         Height          =   1725
         Left            =   -74925
         TabIndex        =   39
         Top             =   375
         Width           =   11565
         Begin VB.CheckBox Check2 
            BackColor       =   &H00FFC0C0&
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
            Left            =   4095
            TabIndex        =   18
            Top             =   555
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
            Left            =   5445
            TabIndex        =   17
            Top             =   210
            Width           =   3480
         End
         Begin Threed.SSFrame SSFrame1 
            Height          =   510
            Left            =   2970
            TabIndex        =   40
            Top             =   825
            Width           =   3390
            _ExtentX        =   5980
            _ExtentY        =   900
            _Version        =   196610
            BackColor       =   -2147483644
            BackStyle       =   1
            Begin TDBDate6Ctl.TDBDate TDBDate1 
               Height          =   315
               Index           =   0
               Left            =   90
               TabIndex        =   19
               Top             =   75
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               Calendar        =   "VIEW_MGMDATA.frx":1D32
               Caption         =   "VIEW_MGMDATA.frx":1E4A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "VIEW_MGMDATA.frx":1EB6
               Keys            =   "VIEW_MGMDATA.frx":1ED4
               Spin            =   "VIEW_MGMDATA.frx":1F32
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   8454016
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
               Calendar        =   "VIEW_MGMDATA.frx":1F5A
               Caption         =   "VIEW_MGMDATA.frx":2072
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "VIEW_MGMDATA.frx":20DE
               Keys            =   "VIEW_MGMDATA.frx":20FC
               Spin            =   "VIEW_MGMDATA.frx":215A
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   8454016
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
               TabIndex        =   41
               Top             =   90
               Visible         =   0   'False
               Width           =   285
            End
         End
         Begin Threed.SSCommand CmdScheduleoK 
            Height          =   390
            Left            =   6645
            TabIndex        =   21
            Top             =   1230
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
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   4
            Left            =   3570
            TabIndex        =   42
            Top             =   225
            Visible         =   0   'False
            Width           =   1905
         End
         Begin Threed.SSCommand CmdToday 
            Height          =   480
            Left            =   120
            TabIndex        =   44
            Top             =   225
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
            TabIndex        =   45
            Top             =   780
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
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Name :"
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
            Index           =   13
            Left            =   2175
            TabIndex        =   43
            Top             =   255
            Width           =   1170
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
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
         Height          =   1770
         Left            =   90
         TabIndex        =   25
         Top             =   390
         Width           =   11550
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   4170
            TabIndex        =   58
            Top             =   1020
            Width           =   3255
         End
         Begin VB.ComboBox cmbStsLastCall 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   0
            Left            =   2970
            TabIndex        =   7
            Top             =   1620
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.ComboBox CmbStatusCek 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   8850
            TabIndex        =   14
            Top             =   630
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
            Left            =   930
            MaskColor       =   &H000000FF&
            TabIndex        =   22
            Top             =   60
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
            Top             =   345
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
            Left            =   8460
            TabIndex        =   13
            Top             =   315
            Width           =   2940
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
            Left            =   6900
            TabIndex        =   12
            Top             =   315
            Width           =   1560
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
            Top             =   660
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
            Top             =   660
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
            Top             =   345
            Width           =   2025
         End
         Begin TDBDate6Ctl.TDBDate TdDob 
            Height          =   315
            Left            =   1200
            TabIndex        =   1
            Top             =   660
            Width           =   1170
            _Version        =   65536
            _ExtentX        =   2064
            _ExtentY        =   556
            Calendar        =   "VIEW_MGMDATA.frx":2182
            Caption         =   "VIEW_MGMDATA.frx":229A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "VIEW_MGMDATA.frx":2306
            Keys            =   "VIEW_MGMDATA.frx":2324
            Spin            =   "VIEW_MGMDATA.frx":2382
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
            Top             =   975
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   556
            Caption         =   "VIEW_MGMDATA.frx":23AA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "VIEW_MGMDATA.frx":2416
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
            Top             =   1380
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   556
            Caption         =   "VIEW_MGMDATA.frx":2458
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "VIEW_MGMDATA.frx":24C4
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
            Left            =   10770
            TabIndex        =   15
            Top             =   840
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
            Left            =   10755
            TabIndex        =   16
            Top             =   1260
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
            Top             =   1365
            Width           =   825
            _Version        =   65536
            _ExtentX        =   1455
            _ExtentY        =   529
            Caption         =   "VIEW_MGMDATA.frx":2506
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "VIEW_MGMDATA.frx":2572
            Spin            =   "VIEW_MGMDATA.frx":25C2
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
            Top             =   1365
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   556
            Calendar        =   "VIEW_MGMDATA.frx":25EA
            Caption         =   "VIEW_MGMDATA.frx":2702
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "VIEW_MGMDATA.frx":276E
            Keys            =   "VIEW_MGMDATA.frx":278C
            Spin            =   "VIEW_MGMDATA.frx":27EA
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
            Top             =   1365
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   556
            Calendar        =   "VIEW_MGMDATA.frx":2812
            Caption         =   "VIEW_MGMDATA.frx":292A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "VIEW_MGMDATA.frx":2996
            Keys            =   "VIEW_MGMDATA.frx":29B4
            Spin            =   "VIEW_MGMDATA.frx":2A12
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
            Top             =   1365
            Width           =   825
            _Version        =   65536
            _ExtentX        =   1455
            _ExtentY        =   529
            Caption         =   "VIEW_MGMDATA.frx":2A3A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "VIEW_MGMDATA.frx":2AA6
            Spin            =   "VIEW_MGMDATA.frx":2AF6
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
            BackColor       =   &H00FFC0C0&
            Caption         =   "Proses....!!"
            Height          =   510
            Left            =   8700
            TabIndex        =   26
            Top             =   1260
            Visible         =   0   'False
            Width           =   1995
            Begin MSComctlLib.ProgressBar ProgressBar1 
               Height          =   270
               Left            =   15
               TabIndex        =   27
               Top             =   165
               Width           =   2025
               _ExtentX        =   3572
               _ExtentY        =   476
               _Version        =   393216
               Appearance      =   0
            End
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
            Index           =   12
            Left            =   2790
            TabIndex        =   57
            Top             =   1050
            Width           =   1365
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Status Last Call"
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
            Index           =   8
            Left            =   2580
            TabIndex        =   38
            Top             =   1740
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Last Call Date"
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
            Index           =   9
            Left            =   2775
            TabIndex        =   37
            Top             =   1425
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
            Left            =   7380
            TabIndex        =   36
            Top             =   690
            Width           =   1365
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "To"
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
            Index           =   11
            Left            =   6285
            TabIndex        =   35
            Top             =   1395
            Width           =   225
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ref No. "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   5
            Left            =   0
            TabIndex        =   34
            Top             =   360
            Width           =   1125
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cellular No."
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
            Index           =   6
            Left            =   75
            TabIndex        =   33
            Top             =   1455
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Home No."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   75
            TabIndex        =   32
            Top             =   1035
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Batch "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   6180
            TabIndex        =   31
            Top             =   360
            Width           =   690
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "DOB"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   60
            TabIndex        =   30
            Top             =   705
            Width           =   1050
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Name"
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
            Index           =   1
            Left            =   2790
            TabIndex        =   29
            Top             =   720
            Width           =   1365
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Customer "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   2700
            TabIndex        =   28
            Top             =   375
            Width           =   1485
         End
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   -15
      TabIndex        =   46
      Top             =   2280
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   12515
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Card Holder Data"
      TabPicture(0)   =   "VIEW_MGMDATA.frx":2B1E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LblTarget(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Check1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Referall Data"
      TabPicture(1)   =   "VIEW_MGMDATA.frx":2B3A
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
         Left            =   2895
         MaskColor       =   &H000000FF&
         TabIndex        =   55
         Top             =   45
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   3225
      End
      Begin VB.Frame Frame1 
         ForeColor       =   &H00000000&
         Height          =   6720
         Left            =   60
         TabIndex        =   50
         Top             =   315
         Width           =   11595
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
            Height          =   300
            Left            =   5040
            Locked          =   -1  'True
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   6315
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
            Height          =   300
            Left            =   8385
            Locked          =   -1  'True
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   6315
            Width           =   3045
         End
         Begin MSComctlLib.ListView LstVwSearchMgm 
            Height          =   6135
            Left            =   0
            TabIndex        =   52
            Top             =   120
            Width           =   11520
            _ExtentX        =   20320
            _ExtentY        =   10821
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
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
            Picture         =   "VIEW_MGMDATA.frx":2B56
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H80000004&
         Height          =   8685
         Left            =   -74940
         TabIndex        =   47
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
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   7980
            Width           =   3045
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   7845
            Left            =   0
            TabIndex        =   49
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
         TabIndex        =   54
         Top             =   -15
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
         TabIndex        =   53
         Top             =   -15
         Width           =   9465
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
Attribute VB_Name = "VIEW_MGMDATAold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim m_cari2 As ADODB.Recordset
'Dim m_cari As ADODB.Recordset
'    Dim mrs_cek As ADODB.Recordset
'Dim StsNa  As String
'Dim StsOP As String
'Dim StsPTP As String
'Dim StsBP As String
'Dim StsPOP As String
'Dim StsSP As String
'Dim StsUC As String
'Dim StsAll As String
'Dim StsRP As String
'Dim cmdsql As String
'Dim LUserType As String
'Dim F_CEK As String
'Dim f_Pending As String
'Dim WO_DATE As String
'Dim datajml As Integer
'Private Sub HEADER_VIEW_Refferall()
'    ListView1.ColumnHeaders.ADD 1, , "No", 3 * TXT
'    ListView1.ColumnHeaders.ADD 2, , "Cust Id", 5 * TXT
'    ListView1.ColumnHeaders.ADD 3, , "Priority", 5 * TXT
'    ListView1.ColumnHeaders.ADD 4, , "Ref Id", 10 * TXT
'    ListView1.ColumnHeaders.ADD 5, , "Ref Name", 10 * TXT
'    ListView1.ColumnHeaders.ADD 6, , "Nama Customer", 25 * TXT
'    'ListView1.ColumnHeaders.ADD 7, , "Batch Expire Date", 25 * TXT
'    ListView1.ColumnHeaders.ADD 7, , "Tgl Schedule", 10 * TXT
'    ListView1.ColumnHeaders.ADD 8, , "Next Action", 12 * TXT
'    ListView1.ColumnHeaders.ADD 9, , "Remarks", 17 * TXT
'    ListView1.ColumnHeaders.ADD 10, , "Sts LastCall", 17 * TXT
'    ListView1.ColumnHeaders.ADD 11, , "SalesCode", 8 * TXT
'    ListView1.ColumnHeaders.ADD 12, , "Agent", 8 * TXT
'    ListView1.ColumnHeaders.ADD 13, , "DataBase", 10 * TXT
'    ListView1.ColumnHeaders.ADD 14, , "LastCall Date", 10 * TXT
'    ListView1.ColumnHeaders.ADD 15, , "Code", 5 * TXT
'    ListView1.ColumnHeaders.ADD 16, , "Complaint Note", 15 * TXT
'    ListView1.ColumnHeaders.ADD 17, , "Check", 10 * TXT
'    ListView1.ColumnHeaders.ADD 18, , "ID", 10 * TXT
'End Sub
'Private Sub isi_dataClaimKeGrid(gCUSTID As String, gNama As String, gnextact As String, gremarks As String, gagent As String, gnamaagent As String, grecsource As String)
'    ' insert ke grid list view
'Dim listitem As listitem
'Set listitem = LstVwSearchMgm.ListItems.ADD(, , "9999")
'    listitem.SubItems(1) = gCUSTID
'    listitem.SubItems(2) = ""
'    listitem.SubItems(3) = gNama
'    listitem.SubItems(4) = Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd hh:nn")
'    listitem.SubItems(5) = gnextact
'    listitem.SubItems(6) = gremarks
'    listitem.SubItems(7) = gagent
'    listitem.SubItems(8) = gnamaagent
'    listitem.SubItems(9) = grecsource
'    listitem.SubItems(10) = ""
'    listitem.SubItems(11) = "1A"
'    listitem.SubItems(12) = ""
'    listitem.SubItems(13) = ""
'    listitem.SubItems(14) = ""
'End Sub
'
'Private Sub CmdScheduleoK_Click()
'If Combo1(4).Text = Empty Then
'    MsgBox "Agent Harus Diisi", vbCritical + vbOKOnly, "Informasi"
'    Exit Sub
'End If
'If TDBDate1(0).ValueIsNull Or TDBDate1(1).ValueIsNull Then
'    MsgBox "Tanggal Tidak Boleh Kosong", vbInformation + vbOKOnly, "Informasi"
'    Exit Sub
'End If
'If TDBDate1(0).Value > TDBDate1(1).Value Then
'    MsgBox "Tanggal Periode Awal harus Lebih Kecil Dari Tanggal Periode Akhir", vbInformation + vbOKOnly, "Informasi"
'    Exit Sub
'End If
'Call cari_Schedule
'End Sub
'
'Private Sub cari_Schedule()
'Dim M_DATA As New CLS_FRMSEARCH
'Dim listitem As listitem
'Dim VOLUMEAMOUNT As Boolean
'Dim sql As String
'Dim rscari As New ADODB.Recordset
'Dim Lcustid1 As String
'If Check2.Value = 1 Then
'
'    LstVwSearchMgm.ListItems.CLEAR
'    SSTab1.Tab = 0
'    ' searching schedule mgm
'  Call CEK_STATUS_F_CEK
'
'   Set m_cari = M_DATA.QUERY_SEARCH_mgm(M_OBJCONN, "AGENT = '" + Combo1(4).Text + "' AND (NEXTACTDATE BETWEEN '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " 00:00" + "' AND '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " 23:59" + "') ", MDIForm1.Text3.Text, F_CEK, f_Pending, "", "", "")
'    ProgressBar1.Max = m_cari.RecordCount + 1
'    If Check2.Value = 1 Then
'        TxtJmlDtMgm.Text = m_cari.RecordCount & " Data"
'    Else
'        Text2.Text = m_cari.RecordCount & " Data"
'    End If
'
'    While Not m_cari.EOF
'    ProgressBar1.Value = m_cari.Bookmark
'     Lcustid1 = CStr(IIf(IsNull(m_cari!CustId), "", m_cari!CustId))
'    Set listitem = LstVwSearchMgm.ListItems.ADD(, , m_cari.Bookmark)
''        listitem.SubItems(1) = IIf(IsNull(m_cari("CUSTID")), "", m_cari("CUSTID"))
''        listitem.SubItems(2) = IIf(IsNull(m_cari("PRIOR")), "", m_cari("PRIOR"))
''        listitem.SubItems(3) = IIf(IsNull(m_cari("NAME")), "", m_cari("NAME"))
''        listitem.SubItems(4) = IIf(IsNull(m_cari("NEXTACTDATE")), "", Format(m_cari("NEXTACTDATE"), "yyyy/mm/dd hh:nn"))
''        listitem.SubItems(5) = IIf(IsNull(m_cari("NEXTACT")), "", m_cari("NEXTACT"))
''        listitem.SubItems(6) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS"))
''        listitem.SubItems(7) = IIf(IsNull(m_cari("F_cek")), "", m_cari("F_cek"))
''        listitem.SubItems(8) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
''        listitem.SubItems(9) = IIf(IsNull(m_cari("NamaAGENT")), "", m_cari("NamaAGENT"))
''        listitem.SubItems(10) = Format(IIf(IsNull(m_cari("Principal")), 0, m_cari("Principal")), "##,###")
''        listitem.SubItems(11) = Format(IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo")), "##,###")
''        listitem.SubItems(12) = IIf(IsNull(m_cari("RECSOURCE")), "", m_cari("RECSOURCE"))
''        listitem.SubItems(13) = IIf(IsNull(m_cari("TGLSTATUS")), "", Format(m_cari("TGLSTATUS"), "YYYY/MM/DD"))
''        listitem.SubItems(14) = IIf(IsNull(m_cari("Kethslkerja")), "", m_cari("Kethslkerja"))
''        listitem.SubItems(15) = Format(IIf(IsNull(m_cari("TGLINCOMING")), "", m_cari("TGLINCOMING")), "YYYY/MM/DD")
''        listitem.SubItems(16) = IIf(IsNull(m_cari("RemarkComplaint")), "", m_cari("RemarkComplaint"))
''        'listitem.SubItems(1) = IIf(IsNull(m_cari("F_CEK")), "", m_cari("F_CEK"))
'
'      '  statusprior = IIf(IsNull(m_cari("StatusPrior")), "", m_cari("StatusPrior"))
'        listitem.SubItems(1) = IIf(IsNull(m_cari("CUSTID")), "", m_cari("CUSTID"))
'        listitem.SubItems(2) = IIf(IsNull(m_cari("NAME")), "", m_cari("NAME"))
'        listitem.SubItems(3) = CStr(IIf(IsNull(m_cari("F_cek")), "", m_cari("F_cek")))
'        sql = "select paydate from tbllunas where custid ='" + Lcustid1 + "' order by paydate desc"
'        Set rscari = New ADODB.Recordset
'        rscari.CursorLocation = adUseClient
'        rscari.Open sql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'        If rscari.BOF And rscari.EOF Then
'            listitem.SubItems(4) = IIf(IsNull(m_cari("date_into_clas")), "", Format(DateAdd("d", 60, m_cari("date_into_clas")), "dd-mm-yyyy"))
'        Else
'            rscari.MoveFirst
'            listitem.SubItems(4) = IIf(IsNull(rscari("paydate")), "", Format(DateAdd("d", 60, rscari("paydate")), "dd-mm-yyyy"))
'        End If
'
'        listitem.SubItems(5) = IIf(IsNull(m_cari("AmountWo")), "", m_cari("AmountWo"))
'        listitem.SubItems(6) = IIf(IsNull(m_cari("Aging")), "", m_cari("Aging"))
'        listitem.SubItems(7) = IIf(IsNull(m_cari("Willing_Ness")), "", m_cari("Willing_Ness"))
'        listitem.SubItems(8) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS")) & "-" & IIf(IsNull(m_cari("f_pending")), "", m_cari("f_pending"))
'        listitem.SubItems(9) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
'        listitem.SubItems(10) = IIf(IsNull(m_cari("StatusCall")), "", m_cari("StatusCall"))
'        listitem.SubItems(11) = IIf(IsNull(m_cari("recsource")), "", m_cari("recsource"))
'        listitem.SubItems(12) = IIf(IsNull(m_cari("TGLCALL")), "", m_cari("TGLCALL"))
''        listitem.SubItems(1) = IIf(IsNull(m_cari("CUSTID")), "", m_cari("CUSTID"))
''        listitem.SubItems(2) = IIf(IsNull(m_cari("PRIOR")), "", m_cari("PRIOR"))
''        listitem.SubItems(3) = IIf(IsNull(m_cari("NAME")), "", m_cari("NAME"))
''        listitem.SubItems(4) = IIf(IsNull(m_cari("TGLSOURCE")), "", DateAdd("d", 90, Format(m_cari("TGLSOURCE"), "dd-mm-yyyy")))
''        listitem.SubItems(5) = IIf(IsNull(m_cari("NEXTACTDATE")), "", Format(m_cari("NEXTACTDATE"), "yyyy/mm/dd hh:nn"))
''        listitem.SubItems(6) = IIf(IsNull(m_cari("NEXTACT")), "", m_cari("NEXTACT"))
''        listitem.SubItems(7) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS")) & "-" & IIf(IsNull(m_cari("f_pending")), "", m_cari("f_pending"))
''        listitem.SubItems(8) = IIf(IsNull(m_cari("F_cek")), "", m_cari("F_cek"))
'
''        listitem.SubItems(11) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
''        'listitem.SubItems(10) = IIf(IsNull(m_cari("NamaAGENT")), "", m_cari("NamaAGENT"))
''        listitem.SubItems(12) = Format(IIf(IsNull(m_cari("Principal")), 0, m_cari("Principal")), "##,###")
''        listitem.SubItems(13) = Format(IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo")), "##,###")
''        listitem.SubItems(14) = Format(IIf(IsNull(m_cari("OpenDate")), "", m_cari("OpenDate")), "yyyy/mm/dd")
''        listitem.SubItems(15) = Format(IIf(IsNull(m_cari("TtlPTP")), 0, m_cari("TtlPTP")), "##,###")
''        listitem.SubItems(16) = IIf(IsNull(m_cari("RECSOURCE")), "", m_cari("RECSOURCE"))
''        listitem.SubItems(17) = IIf(IsNull(m_cari("TGLSTATUS")), "", Format(m_cari("TGLSTATUS"), "YYYY/MM/DD"))
''        listitem.SubItems(18) = IIf(IsNull(m_cari("TGLcall")), "", Format(m_cari("TGLcall"), "YYYY/MM/DD"))
''        listitem.SubItems(19) = IIf(IsNull(m_cari("Kethslkerja")), "", m_cari("Kethslkerja"))
''        listitem.SubItems(20) = Format(IIf(IsNull(m_cari("TGLINCOMING")), "", m_cari("TGLINCOMING")), "YYYY/MM/DD")
''        'listitem.SubItems(18) = IIf(IsNull(m_cari("RemarkComplaint")), "", m_cari("RemarkComplaint"))
''        VOLUMEAMOUNT = VOLUMEAMOUNT + IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo"))
'        'LISTITEM.SubItems(15) = IIf(IsNull(m_cari("[NO]")), "", m_cari("[NO]"))
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
'    ' searching schedule leads
'    Set m_cari = M_DATA.QUERY_SEARCH(M_OBJCONN, "AGENT = '" + Combo1(4).Text + "' AND (NEXTACTDATE BETWEEN '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " 00:00" + "' AND '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " 23:59" + "') ", MDIForm1.Text3.Text)
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
'Set M_DATA = Nothing
'
'End Sub
'
'Private Sub CmdMissed_Click()
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
'        listitem.SubItems(2) = IIf(IsNull(m_cari("PRIOR")), "", m_cari("PRIOR"))
'        listitem.SubItems(3) = IIf(IsNull(m_cari("NAME")), "", m_cari("NAME"))
'        listitem.SubItems(4) = IIf(IsNull(m_cari("NEXTACTDATE")), "", Format(m_cari("NEXTACTDATE"), "yyyy/mm/dd hh:nn"))
'        listitem.SubItems(5) = IIf(IsNull(m_cari("NEXTACT")), "", m_cari("NEXTACT"))
'        listitem.SubItems(6) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS"))
'        listitem.SubItems(7) = IIf(IsNull(m_cari("F_cek")), "", m_cari("F_cek"))
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
'        'listitem.SubItems(18) = IIf(IsNull(m_cari("RemarkComplaint")), "", m_cari("RemarkComplaint"))
''        listitem.SubItems(1) = IIf(IsNull(m_cari("CUSTID")), "", m_cari("CUSTID"))
''        listitem.SubItems(2) = IIf(IsNull(m_cari("PRIOR")), "", m_cari("PRIOR"))
''        listitem.SubItems(3) = IIf(IsNull(m_cari("NAME")), "", m_cari("NAME"))
''        listitem.SubItems(4) = IIf(IsNull(m_cari("NEXTACTDATE")), "", Format(m_cari("NEXTACTDATE"), "yyyy/mm/dd hh:nn"))
''        listitem.SubItems(5) = IIf(IsNull(m_cari("NEXTACT")), "", m_cari("NEXTACT"))
''        listitem.SubItems(6) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS"))
''        listitem.SubItems(7) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
''        listitem.SubItems(8) = IIf(IsNull(m_cari("NamaAGENT")), "", m_cari("NamaAGENT"))
''        listitem.SubItems(9) = Format(IIf(IsNull(m_cari("Principal")), 0, m_cari("Principal")), "##,###")
''        listitem.SubItems(10) = Format(IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo")), "##,###")
''        listitem.SubItems(11) = IIf(IsNull(m_cari("RECSOURCE")), "", m_cari("RECSOURCE"))
''        listitem.SubItems(12) = IIf(IsNull(m_cari("TGLSTATUS")), "", Format(m_cari("TGLSTATUS"), "YYYY/MM/DD"))
''        listitem.SubItems(13) = IIf(IsNull(m_cari("Kethslkerja")), "", m_cari("Kethslkerja"))
''        listitem.SubItems(14) = Format(IIf(IsNull(m_cari("TGLINCOMING")), "", m_cari("TGLINCOMING")), "YYYY/MM/DD")
''        listitem.SubItems(15) = IIf(IsNull(m_cari("RemarkComplaint")), "", m_cari("RemarkComplaint"))
''        listitem.SubItems(16) = IIf(IsNull(m_cari("F_CEK")), "", m_cari("F_CEK"))
'        VOLUMEAMOUNT = VOLUMEAMOUNT + IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo"))
'        'LISTITEM.SubItems(15) = IIf(IsNull(m_cari("[NO]")), "", m_cari("[NO]"))
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
'End Sub
'Private Sub CEK_STATUS_F_CEK()
'Dim cmdsql As String
'Dim m_objrs As New ADODB.Recordset
'
'F_CEK = Empty
'Set m_objrs = New ADODB.Recordset
'        m_objrs.CursorLocation = adUseClient
'        cmdsql = "SELECT F_NA, F_OP, F_PTP, F_BP, F_POP, F_SP, F_UC,F_RP,usertype FROM usertbl WHERE USERID = '" + mdiform1.txtusername.text + "'"
'         m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'
'
'        While Not m_objrs.EOF
'                StsNa = CStr(Trim(IIf(IsNull(m_objrs!f_na), "", m_objrs!f_na)))
'                StsOP = CStr(Trim(IIf(IsNull(m_objrs!f_op), "", m_objrs!f_op)))
'                StsPTP = CStr(Trim(IIf(IsNull(m_objrs!f_ptp), "", m_objrs!f_ptp)))
'                StsBP = CStr(Trim(IIf(IsNull(m_objrs!f_bp), Empty, m_objrs!f_bp)))
'                StsPOP = CStr(Trim(IIf(IsNull(m_objrs!f_pop), "", m_objrs!f_pop)))
'                StsSP = CStr(Trim(IIf(IsNull(m_objrs!f_sp), "", m_objrs!f_sp)))
'                StsRP = CStr(Trim(IIf(IsNull(m_objrs!f_rp), "", m_objrs!f_rp)))
'                StsUC = CStr(Trim(IIf(IsNull(m_objrs!F_UC), "", m_objrs!F_UC)))
'                LUserType = CStr(Trim(IIf(IsNull(m_objrs!usertype), "", m_objrs!usertype)))
'                m_objrs.MoveNext
'            Wend
'            Set m_objrs = Nothing
'             StsAll = StsNa + StsOP + StsPTP + StsBP + StsPOP + StsSP + StsRP + StsUC
'
'     If StsAll <> Empty Then
'            If LUserType = "1" Then
'                    If StsUC <> Empty Then
'                        F_CEK = "(left(F_CEK,3)IN('NK-','MV-','WN-','NBP','" + StsNa + "','" + StsOP + "','" + StsPTP + "','" + StsBP + "','" + StsPOP + "','" + StsSP + "','" + StsRP + "') or F_CEK IS NULL)"
'                        Else
'                        F_CEK = "(left(F_CEK,3)IN( '" + StsNa + "','" + StsOP + "','" + StsPTP + "','" + StsBP + "','" + StsPOP + "','" + StsSP + "','" + StsRP + "') or F_CEK IS NULL) "
'                    End If
'            End If
'     End If
'
'End Sub
'Private Sub CmdToday_Click()
'Dim M_DATA As New CLS_FRMSEARCH
'Dim listitem As listitem
'Dim cmdsql As String
'
'Dim VOLUMEAMOUNT As Double
'If Check2.Value = 1 Then
'    Call CEK_STATUS_F_CEK
'    Set m_cari = M_DATA.QUERY_SEARCH_mgm(M_OBJCONN, "AGENT = '" + Combo1(4).Text + "'  AND (NEXTACTDATE BETWEEN '" + Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") & " 00:00" + "' AND '" + Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") & " 23:59" + "') ", MDIForm1.Text3.Text, F_CEK, f_Pending, "", "", "")
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
'       listitem.SubItems(1) = IIf(IsNull(m_cari("CUSTID")), "", m_cari("CUSTID"))
'        listitem.SubItems(2) = IIf(IsNull(m_cari("PRIOR")), "", m_cari("PRIOR"))
'        listitem.SubItems(3) = IIf(IsNull(m_cari("NAME")), "", m_cari("NAME"))
'        listitem.SubItems(4) = IIf(IsNull(m_cari("NEXTACTDATE")), "", Format(m_cari("NEXTACTDATE"), "yyyy/mm/dd hh:nn"))
'        listitem.SubItems(5) = IIf(IsNull(m_cari("NEXTACT")), "", m_cari("NEXTACT"))
'        listitem.SubItems(6) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS"))
'        listitem.SubItems(7) = IIf(IsNull(m_cari("F_cek")), "", m_cari("F_cek"))
'        listitem.SubItems(9) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
'        'listitem.SubItems(9) = IIf(IsNull(m_cari("NamaAGENT")), "", m_cari("NamaAGENT"))
'        listitem.SubItems(10) = Format(IIf(IsNull(m_cari("Principal")), 0, m_cari("Principal")), "##,###")
'        listitem.SubItems(11) = Format(IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo")), "##,###")
'        listitem.SubItems(12) = Format(IIf(IsNull(m_cari("OpenDate")), "", m_cari("OpenDate")), "yyyy/mm/dd")
'        listitem.SubItems(13) = Format(IIf(IsNull(m_cari("TtlPTP")), 0, m_cari("TtlPTP")), "##,###")
'        listitem.SubItems(14) = IIf(IsNull(m_cari("RECSOURCE")), "", m_cari("RECSOURCE"))
'        listitem.SubItems(15) = IIf(IsNull(m_cari("TGLSTATUS")), "", Format(m_cari("TGLSTATUS"), "YYYY/MM/DD"))
'        listitem.SubItems(16) = IIf(IsNull(m_cari("Kethslkerja")), "", m_cari("Kethslkerja"))
'        listitem.SubItems(17) = Format(IIf(IsNull(m_cari("TGLINCOMING")), "", m_cari("TGLINCOMING")), "YYYY/MM/DD")
'        'listitem.SubItems(18) = IIf(IsNull(m_cari("RemarkComplaint")), "", m_cari("RemarkComplaint"))
''        listitem.SubItems(1) = IIf(IsNull(m_cari("CUSTID")), "", m_cari("CUSTID"))
''        listitem.SubItems(2) = IIf(IsNull(m_cari("PRIOR")), "", m_cari("PRIOR"))
''        listitem.SubItems(3) = IIf(IsNull(m_cari("NAME")), "", m_cari("NAME"))
''        listitem.SubItems(4) = IIf(IsNull(m_cari("NEXTACTDATE")), "", Format(m_cari("NEXTACTDATE"), "yyyy/mm/dd hh:nn"))
''        listitem.SubItems(5) = IIf(IsNull(m_cari("NEXTACT")), "", m_cari("NEXTACT"))
''        listitem.SubItems(6) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS"))
''        listitem.SubItems(7) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
''        listitem.SubItems(8) = IIf(IsNull(m_cari("NamaAGENT")), "", m_cari("NamaAGENT"))
''        listitem.SubItems(9) = Format(IIf(IsNull(m_cari("Principal")), 0, m_cari("Principal")), "##,###")
''        listitem.SubItems(10) = Format(IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo")), "##,###")
''        listitem.SubItems(11) = IIf(IsNull(m_cari("RECSOURCE")), "", m_cari("RECSOURCE"))
''        listitem.SubItems(12) = IIf(IsNull(m_cari("TGLSTATUS")), "", Format(m_cari("TGLSTATUS"), "YYYY/MM/DD"))
''        listitem.SubItems(13) = IIf(IsNull(m_cari("Kethslkerja")), "", m_cari("Kethslkerja"))
''        listitem.SubItems(14) = Format(IIf(IsNull(m_cari("TGLINCOMING")), "", m_cari("TGLINCOMING")), "YYYY/MM/DD")
''        listitem.SubItems(15) = IIf(IsNull(m_cari("RemarkComplaint")), "", m_cari("RemarkComplaint"))
''        listitem.SubItems(16) = IIf(IsNull(m_cari("F_CEK")), "", m_cari("F_CEK"))
'        VOLUMEAMOUNT = VOLUMEAMOUNT + IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo"))
'        'LISTITEM.SubItems(15) = IIf(IsNull(m_cari("[NO]")), "", m_cari("[NO]"))
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
'    Set m_cari = M_DATA.QUERY_SEARCH(M_OBJCONN, "AGENT = '" + Combo1(4).Text + "'  AND (NEXTACTDATE BETWEEN '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " 00:00" + "' AND '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " 23:59" + "') ", MDIForm1.Text3.Text)
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
'
'End If
'End Sub
'
'Private Sub Combo1_Click(Index As Integer)
'Dim M_DATA As New CLS_FRMSEARCH
'Dim m_objrs As ADODB.Recordset
'Select Case Index
'Case 0
'    Set m_objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "USERID = '" + Combo1(Index).Text + "'")
'    If m_objrs.RecordCount <> 0 Then
'        Combo1(0).Text = m_objrs("USERID")
'        Combo1(1).Text = m_objrs("AGENT")
'    Else
'        Combo1(0).Text = Empty
'        Combo1(1).Text = Empty
'    End If
'Case 1
'    Set m_objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "AGENT = '" + Combo1(Index).Text + "'")
'    If m_objrs.RecordCount <> 0 Then
'        Combo1(0).Text = m_objrs("USERID")
'        Combo1(1).Text = m_objrs("AGENT")
'    Else
'        Combo1(0).Text = Empty
'        Combo1(1).Text = Empty
'    End If
'Case 2
'Set m_objrs = M_DATA.QUERY_DATASOURCE(M_OBJCONN, "KODEDS = '" + Combo1(Index).Text + "'")
'    If m_objrs.RecordCount <> 0 Then
'        Combo1(2).Text = m_objrs("KODEDS")
'        Combo1(3).Text = m_objrs("KETERANGAN")
'    Else
'        Combo1(2).Text = Empty
'        Combo1(3).Text = Empty
'    End If
'Case 3
'Set m_objrs = M_DATA.QUERY_DATASOURCE(M_OBJCONN, "KETERANGAN = '" + Combo1(Index).Text + "'")
'    If m_objrs.RecordCount <> 0 Then
'        Combo1(2).Text = m_objrs("KODEDS")
'        Combo1(3).Text = m_objrs("KETERANGAN")
'    Else
'        Combo1(2).Text = Empty
'        Combo1(3).Text = Empty
'    End If
'Case 4
'    Set m_objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "USERID = '" + Combo1(Index).Text + "'")
'    If m_objrs.RecordCount <> 0 Then
'        Combo1(4).Text = m_objrs("USERID")
'        Combo1(5).Text = m_objrs("AGENT")
'    Else
'        Combo1(4).Text = Empty
'        Combo1(5).Text = Empty
'    End If
'Case 5
'    Set m_objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "AGENT = '" + Combo1(Index).Text + "'")
'    If m_objrs.RecordCount <> 0 Then
'        Combo1(4).Text = m_objrs("USERID")
'        Combo1(5).Text = m_objrs("AGENT")
'    Else
'        Combo1(4).Text = Empty
'        Combo1(5).Text = Empty
'    End If
'End Select
'Set M_DATA = Nothing
'Set m_objrs = Nothing
'End Sub
'
'Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
'Dim sSearchText As String
'Dim lReturn As Long
'Select Case Index
'Case 0, 1, 2, 3
'If KeyAscii = 13 Then
'   Combo1_Click (Index)
'   KeyAscii = 0
'Else
'   sSearchText = Left$(Combo1(Index).Text, Combo1(Index).SelStart) & Chr$(KeyAscii)
'   lReturn = SendMessage(Combo1(Index).hwnd, CB_FINDSTRING, -1, ByVal sSearchText)
'   If lReturn <> CB_ERR Then
'      mbIgnoreListClick = True
'      Combo1(Index).ListIndex = lReturn
'      mbIgnoreListClick = False
'      Combo1(Index).Text = Combo1(Index).list(lReturn)
'      Combo1(Index).SelStart = Len(sSearchText)
'      Combo1(Index).SelLength = Len(Combo1(Index).Text)
'      KeyAscii = 0
'   End If
'End If
'End Select
'End Sub
'
'Private Sub Combo1_LostFocus(Index As Integer)
'Dim M_DATA As New CLS_FRMSEARCH
'Dim m_objrs As ADODB.Recordset
'Select Case Index
'Case 0
'    Set m_objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "USERID = '" + Combo1(Index).Text + "'")
'    If m_objrs.RecordCount <> 0 Then
'        Combo1(0).Text = m_objrs("USERID")
'        Combo1(1).Text = m_objrs("AGENT")
'    Else
'        Combo1(0).Text = Empty
'        Combo1(1).Text = Empty
'    End If
'Case 1
'    Set m_objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "AGENT = '" + Combo1(Index).Text + "'")
'    If m_objrs.RecordCount <> 0 Then
'        Combo1(0).Text = m_objrs("USERID")
'        Combo1(1).Text = m_objrs("AGENT")
'    Else
'        Combo1(0).Text = Empty
'        Combo1(1).Text = Empty
'    End If
'Case 2
'Set m_objrs = M_DATA.QUERY_DATASOURCE(M_OBJCONN, "KODEDS = '" + Combo1(Index).Text + "'")
'    If m_objrs.RecordCount <> 0 Then
'        Combo1(2).Text = m_objrs("KODEDS")
'        Combo1(3).Text = m_objrs("KETERANGAN")
'    Else
'        Combo1(2).Text = Empty
'        Combo1(3).Text = Empty
'    End If
'Case 3
'Set m_objrs = M_DATA.QUERY_DATASOURCE(M_OBJCONN, "KETERANGAN = '" + Combo1(Index).Text + "'")
'    If m_objrs.RecordCount <> 0 Then
'        Combo1(2).Text = m_objrs("KODEDS")
'        Combo1(3).Text = m_objrs("KETERANGAN")
'    Else
'        Combo1(2).Text = Empty
'        Combo1(3).Text = Empty
'    End If
'End Select
'Set M_DATA = Nothing
'Set m_objrs = Nothing
'End Sub
'
'Private Sub Form_Load()
'Dim m_objrs As ADODB.Recordset
'Dim M_DATA As New CLS_FRMSEARCH
'
'Call HEADER_VIEW_mgm
'Call HEADER_VIEW_Refferall
'
'
'If UCase(mdiform1.txtlevel.text) = "AGENT" Then
'    CekDtDistribute.Visible = False
'    Combo1(4).Text = mdiform1.txtusername.text
'    Combo1(5).Text = mdiform1.txtnama.text
'    Combo1(4).Visible = False
'    Combo1(5).Visible = False
'    Label1(13).Visible = False
'    Combo1(0).Text = mdiform1.txtusername.text
'    Combo1(1).Text = mdiform1.txtnama.text
'Else
'    CekDtDistribute.Visible = True
'End If
'    CmbStatusCek.AddItem "ACCEPT"
'    CmbStatusCek.AddItem "RETURN"
'    CmbStatusCek.AddItem "NOT CHECK"
'    SSTab1.Tab = 0
'    SSTab2.Tab = 0
'Me.Width = 11880
'Me.Height = 9945
''Me.Height = 6105
'SSTab1.TabVisible(1) = False
'Me.Top = 500
'Me.Left = 1000
'
'
'cmbStsLastCall(0).AddItem "NOT CONTACTED"
'cmbStsLastCall(0).AddItem "NOT AVAILABLE"
'cmbStsLastCall(0).AddItem "STILL THINKING"
'cmbStsLastCall(0).AddItem "DISAGREE"
'cmbStsLastCall(0).AddItem "SENDING APPLICATION"
'cmbStsLastCall(0).AddItem "PICKUP"
'cmbStsLastCall(0).AddItem "INCOMPLETE DOCUMENT"
'cmbStsLastCall(0).AddItem "INCOMING"
'cmbStsLastCall(0).AddItem "AVAILABLE"
'
'DTimeLastCall(0).Value = "00:00"
'DTimeLastCall(1).Value = "23:59"
'StsmgmSchedule = False
'
'
'Set m_objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "")
'    While Not m_objrs.EOF
'        Combo1(0).AddItem m_objrs("USERID")
'        Combo1(1).AddItem m_objrs("AGENT")
'        Combo1(4).AddItem m_objrs("USERID")
'        Combo1(5).AddItem m_objrs("AGENT")
'        m_objrs.MoveNext
'    Wend
' Set m_objrs = Nothing
'
'Combo2.AddItem "OP"
'Combo2.AddItem "PTP"
'Combo2.AddItem "MV"
'Combo2.AddItem "BP"
'Combo2.AddItem "RP"
'Combo2.AddItem "NK"
'Combo2.AddItem "POP"
'Combo2.AddItem "NBP"
'Combo2.AddItem "SP"
'Combo2.AddItem "WN"
'Combo2.AddItem "NA"
'
'
'
' Set m_objrs = M_DATA.QUERY_DATASOURCE(M_OBJCONN, "")
'    While Not m_objrs.EOF
'        Combo1(2).AddItem m_objrs("KODEDS")
'        Combo1(3).AddItem m_objrs("KETERANGAN")
'        m_objrs.MoveNext
'    Wend
'If UCase(MDIForm1.Text3.Text) = "ADMIN" Then
'    Label1(5).Visible = True
'    Text1(2).Visible = True
'End If
'm_objrs.Close
'Set m_objrs = Nothing
'Set M_DATA = Nothing
'End Sub
'Private Sub show_Search_mgmData()
'Dim listitem As listitem
'Dim Lcustid1 As String
'Dim Lcustid2 As String
'Dim LCall As String
'Dim i As Integer
'Dim cmdsql As String
'Dim sPending As String
'Dim m_objrs As ADODB.Recordset
'Dim VOLUMEAMOUNT As Double
'Dim rscari As New ADODB.Recordset
'Dim sql As String
'
'Dim statusprior As String
'Dim exp%
'i = 1
'On Error GoTo HELL
'
'    datajml = m_cari.RecordCount
'    LstVwSearchMgm.ListItems.CLEAR
'    Me.MousePointer = vbHourglass
'    ProgressBar1.Max = m_cari.RecordCount + 1
'    While Not m_cari.EOF
'    ProgressBar1.Value = m_cari.Bookmark
'        Lcustid1 = CStr(IIf(IsNull(m_cari!CustId), "", m_cari!CustId))
'        sPending = CStr(Trim(IIf(IsNull(m_cari!f_Pending), "", m_cari!f_Pending)))
'        'If sPending = "" Then sPending = ""
'
'        Set listitem = LstVwSearchMgm.ListItems.ADD(, , m_cari.Bookmark)
'        statusprior = IIf(IsNull(m_cari("StatusPrior")), "", m_cari("StatusPrior"))
'        listitem.SubItems(1) = IIf(IsNull(m_cari("CUSTID")), "", m_cari("CUSTID"))
'        listitem.SubItems(2) = IIf(IsNull(m_cari("NAME")), "", m_cari("NAME"))
'        listitem.SubItems(3) = CStr(IIf(IsNull(m_cari("F_cek")), "", m_cari("F_cek")) & " " & sPending)
'        sql = "select paydate from tbllunas where custid ='" + Lcustid1 + "' order by paydate desc"
'        Set rscari = New ADODB.Recordset
'        rscari.CursorLocation = adUseClient
'        rscari.Open sql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'        If rscari.BOF And rscari.EOF Then
'            listitem.SubItems(4) = IIf(IsNull(m_cari("date_into_clas")), "", Format(DateAdd("d", 60, m_cari("date_into_clas")), "dd-mm-yyyy"))
'        Else
'            rscari.MoveFirst
'            listitem.SubItems(4) = IIf(IsNull(rscari("paydate")), "", Format(DateAdd("d", 60, rscari("paydate")), "dd-mm-yyyy"))
'        End If
'
'
'        listitem.SubItems(5) = IIf(IsNull(m_cari("AmountWo")), "", Format(m_cari("AmountWo"), "###,##"))
'        listitem.SubItems(6) = IIf(IsNull(m_cari("Aging")), "", m_cari("Aging"))
'        listitem.SubItems(7) = IIf(IsNull(m_cari("Willing_Ness")), "", m_cari("Willing_Ness"))
'        listitem.SubItems(8) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS"))
'        listitem.SubItems(9) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
'        listitem.SubItems(10) = IIf(IsNull(m_cari("StatusCall")), "", m_cari("StatusCall"))
'        listitem.SubItems(11) = IIf(IsNull(m_cari("recsource")), "", m_cari("recsource"))
'        listitem.SubItems(12) = IIf(IsNull(m_cari("TGLCALL")), "", m_cari("TGLCALL"))
'SorryLompat:
'        VOLUMEAMOUNT = VOLUMEAMOUNT + IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo"))
'        m_cari.MoveNext
'    Wend
'
'        If LstVwSearchMgm.ListItems.Count = 0 Then
'            TxtJmlDtMgm.Text = "Tidak Ada Data"
'            TxtJmlVolMgm.Text = "0"
'        Else
'            TxtJmlDtMgm.Text = "Total " + CStr(m_cari.RecordCount) + " Records"
'            TxtJmlVolMgm.Text = "Total " + CStr(Format(VOLUMEAMOUNT, "##,###"))
'        End If
'LstVwSearchMgm.SortKey = 2
'LstVwSearchMgm.Sorted = True
'ProgressBar1.Value = 0
'ProgressBar1.Visible = False
'MousePointer = vbNormal
'Set m_cari = Nothing
'Set m_cari2 = Nothing
'
'Exit Sub
'HELL:
'    Me.MousePointer = vbNormal
'    MsgBox Err.Description
'  ''  Resume
'End Sub
'
'Private Sub HEADER_VIEW_mgm()
'    LstVwSearchMgm.ColumnHeaders.ADD 1, , "No", 3 * TXT
'    LstVwSearchMgm.ColumnHeaders.ADD 2, , "Cust Id", 12 * TXT
'    LstVwSearchMgm.ColumnHeaders.ADD 3, , "Cust Name", 15 * TXT
'    LstVwSearchMgm.ColumnHeaders.ADD 4, , "Sts Account", 10 * TXT
'
'    LstVwSearchMgm.ColumnHeaders.ADD 5, , "Expire Date", 10 * TXT
'    LstVwSearchMgm.ColumnHeaders.ADD 6, , "AmountWo", 8 * TXT
'    LstVwSearchMgm.ColumnHeaders.ADD 7, , "Aging", 5 * TXT
'    LstVwSearchMgm.ColumnHeaders.ADD 8, , "WillingNess", 5 * TXT
'    LstVwSearchMgm.ColumnHeaders.ADD 9, , "Remarks", 35 * TXT
'    LstVwSearchMgm.ColumnHeaders.ADD 10, , "SalesCode", 8 * TXT
'    LstVwSearchMgm.ColumnHeaders.ADD 11, , "Last Call", 15 * TXT
'    LstVwSearchMgm.ColumnHeaders.ADD 12, , "Batch", 15 * TXT
'    LstVwSearchMgm.ColumnHeaders.ADD 13, , "Last Call Date", 15 * TXT
'
''    LstVwSearchmgm.ColumnHeaders.ADD 3, , "Priority", 5 * TXT
''    LstVwSearchmgm.ColumnHeaders.ADD 4, , "Nama Customer", 10 * TXT
''    LstVwSearchmgm.ColumnHeaders.ADD 5, , "Batch", 10 * TXT
''    LstVwSearchmgm.ColumnHeaders.ADD 6, , "Tgl Schedule", 5 * TXT
''    LstVwSearchmgm.ColumnHeaders.ADD 7, , "Next Action", 10 * TXT
''    LstVwSearchmgm.ColumnHeaders.ADD 8, , "Remarks", 17 * TXT
''    LstVwSearchmgm.ColumnHeaders.ADD 9, , "Sts Account", 17 * TXT
''    LstVwSearchmgm.ColumnHeaders.ADD 10, , "Sts LastCall", 17 * TXT
''    LstVwSearchmgm.ColumnHeaders.ADD 11, , "Call Initial", 5 * TXT
''    LstVwSearchmgm.ColumnHeaders.ADD 12, , "SalesCode", 5 * TXT
''    'LstVwSearchmgm.ColumnHeaders.ADD 10, , "Agent", 1 * 1
''
''    LstVwSearchmgm.ColumnHeaders.ADD 13, , "Principle", 10 * TXT
''    LstVwSearchmgm.ColumnHeaders.ADD 14, , "Total Amount", 10 * TXT
''
''    LstVwSearchmgm.ColumnHeaders.ADD 15, , "Join Date", 5 * TXT
''    LstVwSearchmgm.ColumnHeaders.ADD 16, , "PTP Amount", 5 * TXT
''
''    LstVwSearchmgm.ColumnHeaders.ADD 17, , "DataBase", 5 * TXT
''    LstVwSearchmgm.ColumnHeaders.ADD 18, , "LastStatus Date", 5 * TXT
''    LstVwSearchmgm.ColumnHeaders.ADD 19, , "LastCall Date", 5 * TXT
''    LstVwSearchmgm.ColumnHeaders.ADD 20, , "Sts Account", 5 * TXT
''
''    LstVwSearchmgm.ColumnHeaders.ADD 21, , "PTP Date", 5 * TXT
''    LstVwSearchmgm.ColumnHeaders.ADD 22, , "Complaint Note", 5 * TXT
''    'LstVwSearchmgm.ColumnHeaders.ADD 23, , "PULLOUT", 5 * TXT
'
'    'LstVwSearchmgm.ColumnHeaders.ADD 16, , "ID", 10 * TXT
'End Sub
'  Sub WaitSecs(Seconds As Single)
' Dim a As Long
' Seconds = Seconds + Timer
' While Seconds > Timer
'  a = DoEvents
' Wend
'End Sub
'Private Sub Command1_Click(Index As Integer)
'Dim NAMACUST As String
'Dim NAMAAGENT As String
'Dim DATASOURCE As String
'Dim TGLLAHIR As String
'Dim OFFPHONE As String
'Dim OFFPHONE2 As String
'Dim HOMEPHONE As String
'Dim HOMEPHONE2 As String
'Dim MOBILEPHONE As String
'Dim MOBILEPHONE2 As String
'Dim FAXPHONE As String
'Dim Lcustid As String
'Dim FAXPHONE2 As String
'Dim KETHSLKERJA As String
'Dim lLastCallDate As String
'Dim lStatusCek As String
'Dim sPending As String
'Dim FCEKSTATUS As String
'
'
'
'Dim M_DATA As New CLS_FRMSEARCH
'Dim m_objrs As New ADODB.Recordset
'Dim PANJANG As Integer
'
'Select Case Index
'    Case 0
'    F_CEK = Empty
'    If UCase(Left(mdiform1.txtlevel.text, 5)) = "ADMIN" Or UCase(Left(mdiform1.txtlevel.text, 5)) = "SUPER" Or UCase(Left(mdiform1.txtlevel.text, 5)) = "TEAML" Then
'    Else
'    Call CEK_STATUS_F_CEK
'    End If
'    'Call UPDATE_BP
'    '         m_objrs.CursorLocation = adUseClient
''        CMDSQL = "SELECT F_NA, F_OP, F_PTP, F_BP, F_POP, F_SP, F_UC,usertype FROM usertbl WHERE USERID = '" + mdiform1.txtusername.text + "'"
''        m_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
''
''
''        While Not m_objrs.EOF
''                StsNa = CStr(Trim(IIf(IsNull(m_objrs!f_na), "", m_objrs!f_na)))
''                StsOP = CStr(Trim(IIf(IsNull(m_objrs!f_op), "", m_objrs!f_op)))
''                StsPTP = CStr(Trim(IIf(IsNull(m_objrs!f_ptp), "", m_objrs!f_ptp)))
''                StsBP = CStr(Trim(IIf(IsNull(m_objrs!f_bp), "", m_objrs!f_bp)))
''                StsPOP = CStr(Trim(IIf(IsNull(m_objrs!f_pop), "", m_objrs!f_pop)))
''                StsSP = CStr(Trim(IIf(IsNull(m_objrs!f_sp), "", m_objrs!f_sp)))
''                StsUC = CStr(Trim(IIf(IsNull(m_objrs!f_uc), "", m_objrs!f_uc)))
''                LUserType = CStr(Trim(IIf(IsNull(m_objrs!usertype), "", m_objrs!usertype)))
''                m_objrs.MoveNext
''            Wend
''            Set m_objrs = Nothing
''            StsAll = StsNa + StsOP + StsBP + StsPOP + StsUC
''
''     If StsAll <> "" Then
''            If LUserType = "1" Then
''                    If StsUC = "UC" Then
''                        F_CEK = "left(F_CEK,3)  IN('NK-','MV-','WN-','" + StsNa + "','" + StsOP + "','" + StsPTP + "','" + StsBP + "','" + StsPOP + "','" + StsSP + "')"
''                        Else
''                        F_CEK = "left(F_CEK,3)  IN( '" + StsNa + "','" + StsOP + "','" + StsPTP + "','" + StsBP + "','" + StsPOP + "','" + StsSP + "')"
''                    End If
''            End If
''     End If
'
'        If Trim(Text1(0).Text) = Empty And Trim(Combo1(0).Text) = Empty And Combo1(2).Text = Empty And Len(TDBMask2.Value) < 1 And Len(TDBMask1.Value) < 1 And TdDob.ValueIsNull And Len(Text1(2).Text) < 3 _
'        And cmbStsLastCall(0).Text = "" And CmbStatusCek.Text = "" And DtLastCall(0).ValueIsNull And CekDtDistribute.Value = 0 Then
'            MsgBox "Masukan Kriteria Customer Yang Akan Dicari...!!!", vbCritical + vbOKOnly, "Peringatan"
'            Text1(0).SetFocus
'            Set M_DATA = Nothing
'            Exit Sub
'        Else
'
'         LstVwSearchMgm.ListItems.CLEAR
'         Frame3.Visible = True
'         If CekDtDistribute.Value = 1 Then
'            NAMAAGENT = "AGENT is null"
'         Else
'            If Text1(2).Text <> Empty Then
'                Lcustid = "CUSTID LIKE " + "'%" + UBAH_QUOTE(Text1(2).Text) + "%'"
'            Else
'                If Text1(0).Text <> Empty Then
'                    NAMACUST = "NAME LIKE " + "'%" + UBAH_QUOTE(Text1(0).Text) + "%'"
'                End If
'                If Combo1(0).Text <> Empty Then
''                    If UCase(Left(mdiform1.txtusername.text, 2)) = "TL" Then
'                        NAMAAGENT = "AGENT = '" + Trim(Combo1(0).Text) + "'"
''                    Else
''                        NAMAAGENT = "AGENT = '" + Trim(Combo1(0).Text) + "'"
''                    End If
'                End If
'                If Combo1(2).Text <> Empty Then
'                    DATASOURCE = "RECSOURCE = '" + Trim(Combo1(2).Text) + "'"
'                End If
'                If TdDob.ValueIsNull Then
'                Else
'                    TGLLAHIR = "DOB = '" + Format(TdDob.Text, "yyyy/mm/dd") + "'"
'                End If
'                If Len(TDBMask1.Value) > 1 Then
'                    OFFPHONE = "OFFICENO Like '%" + TDBMask1.Value + "%'"
'                    OFFPHONE2 = "OFFICENO2 Like '%" + TDBMask1.Value + "%'"
'                    HOMEPHONE = "HOMENO Like '%" + TDBMask1.Value + "%'"
'                    HOMEPHONE2 = "HOMENO2 Like '%" + TDBMask1.Value + "%'"
'                    FAXPHONE = "FAXNO Like '%" + TDBMask1.Value + "%'"
'                    FAXPHONE2 = "FAXNO2 Like '%" + TDBMask1.Value + "%'"
'                End If
'
'                If Len(Combo2) > 0 Then
'                FCEKSTATUS = " F_CEK LIKE '" + Combo2.Text + "%'"
'                End If
'
'                If Len(TDBMask2.Value) > 1 Then
'                    MOBILEPHONE = "MOBILENO like '%" + TDBMask2.Value + "%'"
'                    MOBILEPHONE2 = "MOBILENO2 like '%" + TDBMask2.Value + "%'"
'                End If
'
'                If cmbStsLastCall(0).Text <> "" Then
'                    Select Case UCase(Trim(cmbStsLastCall(0).Text))
'                        Case "NOT CONTACTED"
'                            KETHSLKERJA = "KETHSLKERJA IN('WN','WN','NP','BT')"
'                        Case "NOT AVAILABLE"
'                            KETHSLKERJA = "KETHSLKERJA = 'NA'"
'                        Case "STILL THINKING"
'                            KETHSLKERJA = "KETHSLKERJA= 'ST'"
'                        Case "DISAGREE"
'                            KETHSLKERJA = "LEFT(KETHSLKERJA,1)= 'D'"
'                        Case "SENDING APPLICATION"
'                            KETHSLKERJA = "KETHSLKERJA= 'SA'"
'                            'KETHSLKERJA = "KETHSLKERJA= 'TBO'"
'                        Case "PICKUP"
'                            KETHSLKERJA = "KETHSLKERJA= 'PU'"
'                            'KETHSLKERJA = "KETHSLKERJA= 'TBO'"
'                        Case "INCOMPLETE DOCUMENT"
'                            KETHSLKERJA = "KETHSLKERJA= 'ID'"
'                            'KETHSLKERJA = "KETHSLKERJA= 'TBO'"
'                        Case "INCOMING"
'                            KETHSLKERJA = "KETHSLKERJA= 'I'"
'                        Case "AVAILABLE"
'                            KETHSLKERJA = "KETHSLKERJA= '1A'"
'                    End Select
'                End If
'
'                If CmbStatusCek.Text <> "" Then
'                    Select Case UCase(Trim(CmbStatusCek.Text))
'                        Case "ACCEPT"
'                            lStatusCek = "UPPER(F_cek) ='ACCEPT'"
'                        Case "RETURN"
'                            lStatusCek = "UPPER(F_cek) = 'RETURN'"
'                        Case "NOT CHECK"
'                            lStatusCek = "(f_cek is null or f_cek = 'NOT ACCEPT' OR F_CEK ='') AND KETHSLKERJA ='I'"
'                    End Select
'                End If
'
'                If DtLastCall(0).ValueIsNull Then
'                Else
'                    lLastCallDate = "TGLSTATUS BETWEEN '" + Format(DtLastCall(0).Value, "MM/DD/YYYY") & " " & CStr(DTimeLastCall(0).Value) + "' AND '" + Format(DtLastCall(1).Value, "MM/DD/YYYY") & " " & CStr(DTimeLastCall(1).Value) + "'"
'                End If
'        End If
'        End If
'
'                'Unload FRM_SEARCH
'                If Check1.Value = 0 Then
'                    Set m_cari = M_DATA.QUERY_SEARCH_CONDITION(M_OBJCONN, NAMACUST, NAMAAGENT, DATASOURCE, TGLLAHIR, _
'                                                            OFFPHONE, OFFPHONE2, HOMEPHONE, HOMEPHONE2, MOBILEPHONE, _
'                                                            MOBILEPHONE2, FAXPHONE, FAXPHONE2, MDIForm1.Text3.Text, Lcustid, F_CEK, lLastCallDate, lStatusCek)
'                Else
'                    Set m_cari = M_DATA.QUERY_SEARCH_CONDITION_mgm(M_OBJCONN, NAMACUST, NAMAAGENT, DATASOURCE, TGLLAHIR, _
'                                                             OFFPHONE, OFFPHONE2, HOMEPHONE, HOMEPHONE2, MOBILEPHONE, _
'                                                            MOBILEPHONE2, FAXPHONE, FAXPHONE2, MDIForm1.Text3.Text, Lcustid, F_CEK, lLastCallDate, lStatusCek, sPending, FCEKSTATUS, WO_DATE, "", "", "", "", "", "", "", "")
'
'                   '  Set m_cari2 = M_DATA.QUERY_SEARCH_CALLINITIATED(M_OBJCONN, NAMACUST, NAMAAGENT, DATASOURCE, TGLLAHIR, _
'                                                             OFFPHONE, OFFPHONE2, HOMEPHONE, HOMEPHONE2, MOBILEPHONE, _
'                                                            MOBILEPHONE2, FAXPHONE, FAXPHONE2, MDIForm1.Text3.Text, Lcustid, F_CEK, lLastCallDate, lStatusCek)
'
'                End If
'
'            If m_cari.RecordCount = 0 Then
'                MsgBox "Data Tidak Ditemukan", vbInformation + vbOKOnly, "TINS"
'                Set M_DATA = Nothing
'                Exit Sub
'            Else
'
'                search_ok = True
'                If Check1.Value = 1 Then
'                    'kalau found refferall data
'                    'Unload FRM_PRESCREEN
'                    'FRM_PRESCREEN.Caption = "Search Non mgm Data"
'                    'FRM_PRESCREEN.Show
'                    SSTab1.Tab = 0
'                    Call show_UCDATA
'                    Call show_Search_mgmData
'
'                Else
'                    ' kalau found mgm data
'                    SSTab1.Tab = 1
'                    'Call show_Search_Refferal
'
''                    Unload VIEW_mgmDATA
''                    If mdiform1.txtlevel.text = "Agent" Then
''                        VIEW_mgmDATA.Caption = "Search mgm Data"
''                    Else
''                        VIEW_mgmDATA.Caption = "Search mgm Data  .... Tekan Huruf ""P"" untuk Mengalihkan Data"
''                    End If
''
''
''                    VIEW_mgmDATA.Show
'                End If
'            'FRM_PRESCREEN.Show vbModal
''                Unload Me
'            End If
'        End If
'    Case 1
'        MDIForm1.LstGrade.ListItems.CLEAR
'        Unload Me
'
'    Case 2
'        Text1(2).Text = Empty
'        Text1(0).Text = Empty
'        TdDob.Text = Empty
'        Combo1(0).Text = Empty
'        Combo1(1).Text = Empty
'        Combo1(2).Text = Empty
'        Combo1(3).Text = Empty
'        TDBMask1.Text = Empty
'        TDBMask2.Text = Empty
'        cmbStsLastCall(0).Text = Empty
'        DtLastCall(0).Value = Empty
'        DtLastCall(1).Value = Empty
'        CmbStatusCek.Text = Empty
'
'End Select
'Set M_DATA = Nothing
'
'' Frame3.Visible = False
'End Sub
'
'Sub UPDATE_BP()
''Dim TGLSYS As ADODB.Recordset
''Dim SPTP As ADODB.Recordset
''Dim CMDSQL As String
''Dim TGLKOMP As Date
''Set TGLSYS = New ADODB.Recordset
''TGLSYS.CursorLocation = adUseClient
''TGLSYS.Open "SELECT TGLSYSTEM FrOM VWCALLCFG1", M_OBJCONN, adOpenDynamic, adLockOptimistic
''TGLKOMP = TGLSYS!TGLSYSTEM
''Set SPTP = New ADODB.Recordset
''CMDSQL = "UPDATE mgm SET KETHSLKErJA='BP-BROKEN PROMISE',F_CEK='BP-',KETHSLKErJADESC='' "
''CMDSQL = CMDSQL + "WHERE CUSTID IN (SELECT CUSTID FrOM mgm WHErE DATEDIFF(day,TdbdatePTP,TGLKOMP)>5)"
'
'End Sub
'
'Private Sub show_UCDATA()
'Dim sdata As ADODB.Recordset
'Dim listitem As listitem
'Dim i%
'Set sdata = New ADODB.Recordset
'sdata.CursorLocation = adUseClient
'cmdsql = "select CUSTID,KETHSLKERJA,AGENT FROM mgm WHERE left(KETHSLKERJA,2) in ('WN','NK','MV') AND AGENT='" & VIEW_MGMDATA.Combo1(0).Text & "'"
'sdata.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'i = 1
'MDIForm1.LstGrade.ListItems.CLEAR
'Do While Not sdata.EOF
'    Set listitem = MDIForm1.LstGrade.ListItems.ADD(, , i)
'        listitem.SubItems(1) = IIf(IsNull(sdata!CustId), "", sdata!CustId)
'        listitem.SubItems(2) = IIf(IsNull(sdata!CustId), "", sdata!KETHSLKERJA)
'        listitem.SubItems(3) = IIf(IsNull(sdata!agent), "", sdata!agent)
'        sdata.MoveNext
'        i = i + 1
'Loop
'End Sub
''Private Sub show_Search_Refferal()
''Dim listitem As listitem
''ProgressBar1.Max = m_cari.RecordCount + 1
''Text2.Text = m_cari.RecordCount & " Data"
''ListView1.ListItems.Clear
''
''While Not m_cari.EOF
''ProgressBar1.Value = m_cari.Bookmark
''Set listitem = ListView1.ListItems.ADD(, , m_cari.Bookmark)
''    listitem.SubItems(1) = IIf(IsNull(m_cari("custid")), "", m_cari("custid"))
''    Select Case m_cari("RECSTATUS")
''    Case "1A"
''        listitem.SubItems(2) = "Available"
''    Case ""
''        listitem.SubItems(2) = "Available"
''    Case Else
''        listitem.SubItems(2) = IIf(IsNull(m_cari("PRIOR")), "", m_cari("PRIOR"))
''    End Select
''    listitem.SubItems(3) = IIf(IsNull(m_cari("CUSTIDREF")), "", m_cari("CUSTIDREF"))
''    listitem.SubItems(4) = IIf(IsNull(m_cari("NAMAREF")), "", m_cari("NAMAREF"))
''    listitem.SubItems(5) = IIf(IsNull(m_cari("NAME")), "", m_cari("NAME"))
''    listitem.SubItems(6) = IIf(IsNull(m_cari("NEXTACTDATE")), "", Format(m_cari("NEXTACTDATE"), "yyyy/mm/dd hh:mm"))
''    listitem.SubItems(7) = IIf(IsNull(m_cari("NEXTACT")), "", m_cari("NEXTACT"))
''    listitem.SubItems(8) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS"))
''    listitem.SubItems(9) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
''    listitem.SubItems(10) = IIf(IsNull(m_cari("NamaAGENT")), "", m_cari("NamaAGENT"))
''    listitem.SubItems(11) = IIf(IsNull(m_cari("RECSOURCEREF")), "", m_cari("RECSOURCEREF"))
''    listitem.SubItems(12) = Format(IIf(IsNull(m_cari("TGLSTATUS")), "", m_cari("TGLSTATUS")), "yyyy/mm/dd")
''    listitem.SubItems(13) = IIf(IsNull(m_cari("Kethslkerja")), "", m_cari("Kethslkerja"))
''    listitem.SubItems(14) = IIf(IsNull(m_cari("KdComplaint")), "", m_cari("KdComplaint"))
''    listitem.SubItems(15) = IIf(IsNull(m_cari("RemarkComplaint")), "", m_cari("RemarkComplaint"))
''    listitem.SubItems(16) = IIf(IsNull(m_cari("F_CEK")), "", m_cari("F_CEK"))
''    listitem.SubItems(17) = IIf(IsNull(m_cari("Nomor")), "", m_cari("Nomor"))
''    m_cari.MoveNext
''Wend
''Set m_cari = Nothing
''End Sub
'
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'MDIForm1.m_targetview = False
'
'End Sub
'
'
'Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
'   ListView1.SortKey = ColumnHeader.Index - 1
'   ListView1.Sorted = True
'End Sub
'
'Private Sub ListView1_DblClick()
'Dim m_objrs As ADODB.Recordset
'If ListView1.ListItems.Count = 0 Then
'    Exit Sub
'End If
'Status_Form = 1
'    If UCase(mdiform1.txtlevel.text) = "TEAMLEADER" Then
'        Set m_objrs = New ADODB.Recordset
'        m_objrs.CursorLocation = adUseClient
'        m_objrs.Open "SELECT USERID FROM usertbl WHERE SPVCODE ='" + mdiform1.txtusername.text + "' AND USERID = '" + ListView1.SelectedItem.SubItems(9) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'        If m_objrs.RecordCount <> 0 Then
'        Else
'            MsgBox "Data Ini Milik Agent Team Leader Yang Lain", vbCritical + vbOKOnly, "TINS"
'            Set m_objrs = Nothing
'            Exit Sub
'        End If
'        Set m_objrs = Nothing
'    Else
'        If UCase(mdiform1.txtlevel.text) = "SUPERVISOR" Or UCase(mdiform1.txtlevel.text) = "ADMINISTRATOR" Then
'        Else
'            If Trim(UCase(mdiform1.txtusername.text)) = Trim(UCase(ListView1.SelectedItem.SubItems(9))) Then
'            Else
'                MsgBox "Data Ini Milik Agent Yang Lain", vbCritical + vbOKOnly, "TINS"
'                Set m_objrs = Nothing
'                Exit Sub
'            End If
'        End If
'    End If
'    FrmCC_Colection.Show vbModal
'End Sub
'
'Private Sub ListView1_KeyPress(KeyAscii As Integer)
'Dim m_objrs As ADODB.Recordset
'Dim m_msgbox As Variant
'If ListView1.ListItems.Count = 0 Then
'    Exit Sub
'End If
'If KeyAscii = 13 Then
'    Call ListView1_DblClick
'End If
'If UCase(mdiform1.txtlevel.text) <> "AGENT" Then
'    If KeyAscii = 100 Or KeyAscii = 68 Then
'        If UCase(mdiform1.txtlevel.text) = "TEAMLEADER" Then
'            Set m_objrs = New ADODB.Recordset
'            m_objrs.CursorLocation = adUseClient
'            m_objrs.Open "SELECT USERID FROM usertbl WHERE SPVCODE ='" + mdiform1.txtusername.text + "' AND USERID = '" + ListView1.SelectedItem.SubItems(9) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'            If m_objrs.RecordCount <> 0 Then
'                m_msgbox = MsgBox("Yakin Akan Di Hapus??...", vbExclamation + vbYesNo, "Informasi")
'                If m_msgbox = vbYes Then
'                    M_OBJCONN.Execute "Delete From cc_custtbl where Nomor = " + ListView1.SelectedItem.SubItems(17) + " And RECSOURCE <> 'infomedia1' and left(recsource,3)<>'inf'"
'                    ListView1.ListItems.Remove ListView1.SelectedItem.Index
'                End If
'            Else
'                MsgBox "Data Ini Milik Agent Team Leader Yang Lain", vbCritical + vbOKOnly, "TINS"
'                Set m_objrs = Nothing
'                Exit Sub
'            End If
'            Set m_objrs = Nothing
'        Else
'            m_msgbox = MsgBox("Yakin Akan Di Hapus??...", vbExclamation + vbYesNo, "Informasi")
'            If m_msgbox = vbYes Then
'                'M_OBJCONN.Execute "Delete From cc_custtbl where custid ='" + ListView1.SelectedItem.SubItems(1) + "'"
'                M_OBJCONN.Execute "Delete From cc_custtbl where Nomor = " + ListView1.SelectedItem.SubItems(17) + " And RECSOURCE <> 'infomedia1'"
'                ListView1.ListItems.Remove ListView1.SelectedItem.Index
'            End If
'        End If
'    End If
'End If
'End Sub
'
'Private Sub LstVwSearchmgm_Click()
'If b_pindah = True Then
'    FrmCustIdTransfer.List1.AddItem LstVwSearchMgm.SelectedItem.SubItems(1)
'End If
'End Sub
'
'Private Sub LstVwSearchmgm_ColumnClick(ByVal ColumnHeader As ColumnHeader)
'   LstVwSearchMgm.SortKey = ColumnHeader.Index - 1
'   LstVwSearchMgm.Sorted = True
'End Sub
'
'Private Sub LstVwSearchmgm_DblClick()
'Dim m_objrs As ADODB.Recordset
'Dim PO_AGENT As String
'
'If LstVwSearchMgm.ListItems.Count = 0 Then
'    Exit Sub
'End If
'glexp = LstVwSearchMgm.SelectedItem.SubItems(4)
'Status_Form = 2
'    If LstVwSearchMgm.ListItems.Count = 0 Then
'        Exit Sub
'    End If
'
'    If UCase(mdiform1.txtlevel.text) = "AGENT" Then
'        If UCase(mdiform1.txtusername.text) <> Trim(UCase(LstVwSearchMgm.SelectedItem.SubItems(9))) Then
'            MsgBox "Anda Tidak Berhak Untuk Mengedit Data Ini", vbCritical + vbOKOnly, "TINS"
'            Exit Sub
'        End If
'    End If
'
'    If UCase(mdiform1.txtlevel.text) = "SUPERVISOR" Then
'        If VIEW_MGMDATA.Combo1(0).Text = "PULLOUT" Then
'            Set m_objrs = New ADODB.Recordset
'            m_objrs.CursorLocation = adUseClient
'            cmdsql = "SELECT PO_Agent FROM mgm where CUSTID='" & LstVwSearchMgm.SelectedItem.SubItems(9) & "'"
'            m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'            PO_AGENT = m_objrs!PO_AGENT
'            Set m_objrs = Nothing
'        Else
'            PO_AGENT = VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(9)
'        End If
'
'        Set m_objrs = New ADODB.Recordset
'        m_objrs.CursorLocation = adUseClient
'        cmdsql = "SELECT USERID FROM usertbl WHERE TEAM ='" + mdiform1.txtusername.text + "' AND USERID = '" + PO_AGENT + "'"
'        m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'        If m_objrs.RecordCount <> 0 Then
'        Else
'            MsgBox "Data Ini Milik Agent Team Leader Yang Lain", vbCritical + vbOKOnly, "TINS"
'            Set m_objrs = Nothing
'            Exit Sub
'        End If
'        Set m_objrs = Nothing
'    End If
'    Me.MousePointer = vbHourglass
'    Flag_mgm = False
'    FrmCC_Colection.Show
'    Me.MousePointer = vbNormal
'End Sub
'
'Private Sub LstVwSearchmgm_KeyPress(KeyAscii As Integer)
'Dim m_objrs As ADODB.Recordset
'If KeyAscii = 13 Then
'    Call LstVwSearchmgm_DblClick
'    Exit Sub
'End If
'If UCase(mdiform1.txtlevel.text) <> "AGENT" Then
'    If KeyAscii = 112 Or KeyAscii = 80 Then
'        If UCase(mdiform1.txtlevel.text) = "SUPERVISOR" Then
'        'If UCase(mdiform1.txtlevel.text) = "TEAMLEADER" Then
'            Set m_objrs = New ADODB.Recordset
'            m_objrs.CursorLocation = adUseClient
'            m_objrs.Open "SELECT USERID FROM usertbl WHERE TEAM ='" + mdiform1.txtusername.text + "' AND USERID = '" + LstVwSearchMgm.SelectedItem.SubItems(11) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'            If m_objrs.RecordCount <> 0 Then
'                With View_AlihData
'                    .Show vbModal
'                    If .ok Then
'                        LstVwSearchMgm.ListItems.Remove LstVwSearchMgm.SelectedItem.Index
'                    End If
'                End With
'                Unload View_AlihData
'            Else
'                MsgBox "Data Ini Milik Agent Team Leader Yang Lain", vbCritical + vbOKOnly, "TINS"
'                Set m_objrs = Nothing
'                Exit Sub
'            End If
'            Set m_objrs = Nothing
'        Else
'            With View_AlihData
'                .Show vbModal
'                If .ok Then
'                    LstVwSearchMgm.ListItems.Remove LstVwSearchMgm.SelectedItem.Index
'                End If
'            End With
'            Unload View_AlihData
'        End If
'    End If
'    If KeyAscii = 73 Or KeyAscii = 105 Then
''        b_pindah = True
''        FrmCustIdTransfer.Show
'        If UCase(Left(mdiform1.txtlevel.text, 5)) = "ADMIN" Then
'        'If UCase(mdiform1.txtlevel.text) = "SUPERVISOR" Or UCase(Left(mdiform1.txtlevel.text, 5)) = "ADMIN" Then
'                b_pindah = True
'                FrmCustIdTransfer.Show
'            Else
'                If UCase(mdiform1.txtlevel.text) = "SUPERVISOR" Then
'                'If UCase(mdiform1.txtlevel.text) = "TEAMLEADER" Then
'                    Set m_objrs = New ADODB.Recordset
'                    m_objrs.CursorLocation = adUseClient
'                    m_objrs.Open "SELECT USERID FROM usertbl WHERE TEAM ='" + mdiform1.txtusername.text + "' AND USERID = '" + LstVwSearchMgm.SelectedItem.SubItems(9) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'                    If m_objrs.RecordCount <> 0 Then
'                        b_pindah = True
'                        FrmCustIdTransfer.Show
'                    Else
'                        MsgBox "Data Ini Milik Agent Team Leader Yang Lain", vbCritical + vbOKOnly, "TINS"
'                        Set m_objrs = Nothing
'                        Exit Sub
'                    End If
'                    Set m_objrs = Nothing
'                End If
'         End If
'    End If
'    If KeyAscii = 88 Or KeyAscii = 120 Then
'        Dim n%
'        If UCase(Left(mdiform1.txtlevel.text, 5)) = "ADMIN" Then
'                n = 1
'                Do While n <= LstVwSearchMgm.ListItems.Count
'                    If LstVwSearchMgm.ListItems(n).Checked = True Then
'                        Frmlock.List1.AddItem LstVwSearchMgm.ListItems(n).SubItems(1)
'        '                Set ls2 = Frmlock.LSTACCESS.ListItems.ADD(, , LstVwSearchmgm.ListItems(N))
'        '                ls2.SubItems(1) = LstVwSearchmgm.SelectedItem
'                    End If
'                    n = n + 1
'                Loop
'                Frmlock.Show
'        End If
'
'        If UCase(mdiform1.txtlevel.text) = "SUPERVISOR" And Combo1(2).Text = "" Then
'            Set m_objrs = New ADODB.Recordset
'            m_objrs.CursorLocation = adUseClient
'            m_objrs.Open "SELECT USERID FROM usertbl WHERE TEAM ='" + mdiform1.txtusername.text + "' AND USERID = '" + LstVwSearchMgm.SelectedItem.SubItems(11) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'            If m_objrs.RecordCount <> 0 Then
'                'MsgBox "jkasja"
'                n = 1
'                Do While n <= LstVwSearchMgm.ListItems.Count
'                    If LstVwSearchMgm.ListItems(n).Checked = True Then
'                        Frmlock.List1.AddItem LstVwSearchMgm.ListItems(n).SubItems(1)
'        '                Set ls2 = Frmlock.LSTACCESS.ListItems.ADD(, , LstVwSearchmgm.ListItems(N))
'        '                ls2.SubItems(1) = LstVwSearchmgm.SelectedItem
'                    End If
'                    n = n + 1
'                Loop
'                Frmlock.Show
'            Else
'                MsgBox "Data Ini Milik Agent Team Leader Yang Lain", vbCritical + vbOKOnly, "TINS"
'                Set m_objrs = Nothing
'                Exit Sub
'            End If
'            Set m_objrs = Nothing
'        End If
'    End If
'End If
'End Sub
'
''INBOUND LEADSE
''Private Sub cmbRecsource_LostFocus()
''Dim m_obj As New ADODB.Recordset
''m_obj.CursorLocation = adUseClient
''m_obj.Open "Select * from DATASOURCETBL WHERE KODEDS = '" + cmbRecsource.Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
''If m_obj.RecordCount <> 0 Then
''    cmbRecsource.Text = m_obj!KODEDS
''Else
''    cmbRecsource.Text = ""
''End If
''Set m_obj = Nothing
''End Sub
'
''Private Sub CmdClaimLeads_Click(Index As Integer)
''Select Case Index
''Case 0
''    If Len(TxtNamaLeads.Text) < 2 Then
''        MsgBox "Nama harus diisi", vbInformation + vbOKOnly, "TINS"
''        Exit Sub
''    End If
''    If Len(TxtTelpRumah.Text) < 2 And Len(TxtTelpKantor.Text) < 2 And Len(TxtHandPhone.Text) < 2 Then
''        MsgBox "Minimal salah satu dari telp harus diisi", vbInformation + vbOKOnly, "TINS"
''        Exit Sub
''    End If
''    'CmdSave.Enabled = False
''    Call cari_duplicate_Leads
''Case 1
''    TdbDOB.Value = Empty
''    TxtNama.Text = Empty
''    TxtTelpRumah.Text = Empty
''    TxtTelpKantor.Text = Empty
''    TxtHandPhone.Text = Empty
''End Select
''End Sub
''
'
''Private Sub cari_duplicate_Leads()
''    Dim CMDSQL As String
''
''    Dim kriteria1 As String
''    Dim kriteria2 As String
''    Dim CUSTID1 As String
''    ' kriteria pertama
''    'nama ama notelp
''    If Len(TxtNamaLeads.Text) > 2 And Len(TxtTelpRumah.Text) > 2 Then
''        kriteria2 = Left(TxtTelpRumah.Text, 5)
''        CMDSQL = "Select * from cc_custtbl where name like '%" + TxtNamaLeads.Text + "%' "
''        CMDSQL = CMDSQL + " and (HOMENO like '%" + kriteria2 + "%' or HOMENO2 like '%" + kriteria2 + "%' or mobileno like '%" + kriteria2 + "%' or mobileno2 like '%" + kriteria2 + "%' or officeno like '%" + kriteria2 + "%' or officeno2 like '%" + kriteria2 + "%') "
''
''    Set mrs_cek = New ADODB.Recordset
''        mrs_cek.CursorLocation = adUseClient
''
''        mrs_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
''        If mrs_cek.RecordCount <> 0 Then
''            CUSTID1 = Empty
''            While Not mrs_cek.EOF
''                CUSTID1 = "REFI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC"))
''                CMDSQL = "Insert into TBL_DUPLIKASI (CUSTID, CUSTIDBARU, NAMABARU, HOMENOBARU, OFFICENOBARU, MOBILENOBARU, AGENTBARU, "
''                If TdbDOB.ValueIsNull = False Then
''                    CMDSQL = CMDSQL + "DOB,"
''                End If
''                CMDSQL = CMDSQL + " RECSOURCEBARU) values "
''                CMDSQL = CMDSQL + "('" + mrs_cek!CUSTID + "',"
''                CMDSQL = CMDSQL + "'" + CUSTID1 + "',"
''                CMDSQL = CMDSQL + "'" + TxtNamaLeads.Text + "',"
''                CMDSQL = CMDSQL + "'" + TxtTelpRumah.Text + "',"
''                CMDSQL = CMDSQL + "'" + TxtTelpKantor.Text + "',"
''                CMDSQL = CMDSQL + "'" + TxtHandPhoneLeads.Text + "',"
''                CMDSQL = CMDSQL + "'" + mdiform1.txtusername.text + "',"
''                If TdbDOB.ValueIsNull = False Then
''                    CMDSQL = CMDSQL + "'" + Format(TdbDOB.Value, "yyyy/mm/dd") + "',"
''                End If
''                CMDSQL = CMDSQL + "'" + cmbRecsource.Text + "')"
''                M_OBJCONN.Execute CMDSQL
''                mrs_cek.MoveNext
''            Wend
''
''            ' tampilin yang duplicate deh...
''                Call show_Leads_Duplicate
''                MsgBox " Nama dan Telp Rumah Ada yg sama", vbInformation + vbOKOnly, "TINS"
''            Set mrs_cek = Nothing
''            Exit Sub
''
''        End If
''        Set mrs_cek = Nothing
''    End If
''    If Len(TxtNamaLeads.Text) > 2 And Len(TxtTelpKantor.Text) > 2 Then
''        kriteria2 = Left(TxtTelpKantor.Text, 5)
''        CMDSQL = "Select * from cc_custtbl where name like '%" + TxtNamaLeads.Text + "%' "
''        CMDSQL = CMDSQL + " and (HOMENO like '%" + kriteria2 + "%' or HOMENO2 like '%" + kriteria2 + "%' or mobileno like '%" + kriteria2 + "%' or mobileno2 like '%" + kriteria2 + "%' or officeno like '%" + kriteria2 + "%' or officeno2 like '%" + kriteria2 + "%') "
''    Set mrs_cek = New ADODB.Recordset
''    mrs_cek.CursorLocation = adUseClient
''
''        mrs_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
''        If mrs_cek.RecordCount <> 0 Then
''            CUSTID1 = Empty
''            While Not mrs_cek.EOF
''                CUSTID1 = "REFI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC"))
''
''                CMDSQL = "Insert into TBL_DUPLIKASI (CUSTID, CUSTIDBARU, NAMABARU, HOMENOBARU, OFFICENOBARU, MOBILENOBARU, AGENTBARU, "
''                If TdbDOB.ValueIsNull = False Then
''                    CMDSQL = CMDSQL + "DOB,"
''                End If
''                CMDSQL = CMDSQL + " RECSOURCEBARU) values "
''                CMDSQL = CMDSQL + "('" + mrs_cek!CUSTID + "',"
''                CMDSQL = CMDSQL + "'" + CUSTID1 + "',"
''                CMDSQL = CMDSQL + "'" + TxtNamaLeads.Text + "',"
''                CMDSQL = CMDSQL + "'" + TxtTelpRumah.Text + "',"
''                CMDSQL = CMDSQL + "'" + TxtTelpKantor.Text + "',"
''                CMDSQL = CMDSQL + "'" + TxtHandPhoneLeads.Text + "',"
''                CMDSQL = CMDSQL + "'" + mdiform1.txtusername.text + "',"
''                If TdbDOB.ValueIsNull = False Then
''                    CMDSQL = CMDSQL + "'" + Format(TdbDOB.Value, "yyyy/mm/dd") + "',"
''                End If
''                CMDSQL = CMDSQL + "'" + cmbRecsource.Text + "')"
''                M_OBJCONN.Execute CMDSQL
''                mrs_cek.MoveNext
''            Wend
''
''            ' show data
''            Call show_Leads_Duplicate
''            MsgBox " Nama dan Telp Kantor Ada yg sama", vbInformation + vbOKOnly, "TINS"
''            Set mrs_cek = Nothing
''            Exit Sub
''        End If
''        Set mrs_cek = Nothing
''
''    End If
''    If Len(TxtNamaLeads.Text) > 2 And Len(TxtHandPhone.Text) > 2 Then
''        kriteria2 = Left(TxtHandPhone.Text, 8)
''        CMDSQL = "Select * from cc_custtbl where name like '%" + TxtNamaLeads.Text + "%' "
''        CMDSQL = CMDSQL + " and (HOMENO like '%" + kriteria2 + "%' or HOMENO2 like '%" + kriteria2 + "%' or mobileno like '%" + kriteria2 + "%' or mobileno2 like '%" + kriteria2 + "%' or officeno like '%" + kriteria2 + "%' or officeno2 like '%" + kriteria2 + "%') "
''    Set mrs_cek = New ADODB.Recordset
''    mrs_cek.CursorLocation = adUseClient
''
''        mrs_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
''        If mrs_cek.RecordCount <> 0 Then
''
''
''            CUSTID1 = Empty
''            While Not mrs_cek.EOF
''                CUSTID1 = "REFI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC"))
''
''                CMDSQL = "Insert into TBL_DUPLIKASI (CUSTID, CUSTIDBARU, NAMABARU, HOMENOBARU, OFFICENOBARU, MOBILENOBARU, AGENTBARU, "
''                If TdbDOB.ValueIsNull = False Then
''                    CMDSQL = CMDSQL + "DOB,"
''                End If
''                CMDSQL = CMDSQL + " RECSOURCEBARU) values "
''                CMDSQL = CMDSQL + "('" + mrs_cek!CUSTID + "',"
''                CMDSQL = CMDSQL + "'" + CUSTID1 + "',"
''                CMDSQL = CMDSQL + "'" + TxtNamaLeads.Text + "',"
''                CMDSQL = CMDSQL + "'" + TxtTelpRumah.Text + "',"
''                CMDSQL = CMDSQL + "'" + TxtTelpKantor.Text + "',"
''                CMDSQL = CMDSQL + "'" + TxtHandPhoneLeads.Text + "',"
''                CMDSQL = CMDSQL + "'" + mdiform1.txtusername.text + "',"
''                If TdbDOB.ValueIsNull = False Then
''                    CMDSQL = CMDSQL + "'" + Format(TdbDOB.Value, "yyyy/mm/dd") + "',"
''                End If
''                CMDSQL = CMDSQL + "'" + cmbRecsource.Text + "')"
''
''                M_OBJCONN.Execute CMDSQL
''                mrs_cek.MoveNext
''            Wend
''
''            ' show data
''            Call show_Leads_Duplicate
''            MsgBox "Nama dan Handphone Ada yg sama", vbInformation + vbOKOnly, "TINS"
''            Set mrs_cek = Nothing
''            Exit Sub
''        End If
''        Set mrs_cek = Nothing
''
''    End If
''    If Len(TxtNamaLeads.Text) > 2 And TdbDOB.ValueIsNull = False Then
''        kriteria2 = Format(TdbDOB.Value, "yyyy/mm/dd")
''        CMDSQL = "Select * from cc_custtbl where name like '%" + TxtNamaLeads.Text + "%' "
''        CMDSQL = CMDSQL + " and birthd = '" + Format(TdbDOB.Value, "yyyy/mm/dd") + "'"
''        Set mrs_cek = New ADODB.Recordset
''            mrs_cek.CursorLocation = adUseClient
''
''        mrs_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
''        If mrs_cek.RecordCount <> 0 Then
''
''
''            CUSTID1 = Empty
''            While Not mrs_cek.EOF
''                CUSTID1 = "REFI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC"))
''
''                CMDSQL = "Insert into TBL_DUPLIKASI (CUSTID, CUSTIDBARU, NAMABARU, HOMENOBARU, OFFICENOBARU, MOBILENOBARU, AGENTBARU, "
''                If TdbDOB.ValueIsNull = False Then
''                    CMDSQL = CMDSQL + "DOB,"
''                End If
''                CMDSQL = CMDSQL + " RECSOURCEBARU) values "
''                CMDSQL = CMDSQL + "('" + mrs_cek!CUSTID + "',"
''                CMDSQL = CMDSQL + "'" + CUSTID1 + "',"
''                CMDSQL = CMDSQL + "'" + TxtNamaLeads.Text + "',"
''                CMDSQL = CMDSQL + "'" + TxtTelpRumah.Text + "',"
''                CMDSQL = CMDSQL + "'" + TxtTelpKantor.Text + "',"
''                CMDSQL = CMDSQL + "'" + TxtHandPhoneLeads.Text + "',"
''                CMDSQL = CMDSQL + "'" + mdiform1.txtusername.text + "',"
''                If TdbDOB.ValueIsNull = False Then
''                    CMDSQL = CMDSQL + "'" + Format(TdbDOB.Value, "yyyy/mm/dd") + "',"
''                End If
''                CMDSQL = CMDSQL + "'" + cmbRecsource.Text + "')"
''
''                M_OBJCONN.Execute CMDSQL
''                mrs_cek.MoveNext
''            Wend
''
''            ' show data
''            Call show_Leads_Duplicate
''            MsgBox "Nama dan DOB Ada yg sama", vbInformation + vbOKOnly, "TINS"
''            Set mrs_cek = Nothing
''            Exit Sub
''        End If
''        Set mrs_cek = Nothing
''    End If
'''        custid1 = "REFI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC"))
'''
'''        cmdsql = "Insert into CC_CUSTTBL(CUSTID, NAME, HOMENO, MOBILENO, OFFICENO, AGENT, RECSOURCE,CustIdRef,"
'''        If TdbDOB.ValueIsNull = False Then
'''            cmdsql = cmdsql + "BIRTHD,"
'''        End If
'''        cmdsql = cmdsql + " RecSourceRef) values"
'''        cmdsql = cmdsql + "('" + custid1 + "',"
'''        cmdsql = cmdsql + "'" + txtnamaleads.Text + "',"
'''        cmdsql = cmdsql + "'" + TxtTelpRumah.Text + "',"
'''        cmdsql = cmdsql + "'" + TxtHandphoneleads.Text + "',"
'''        cmdsql = cmdsql + "'" + TxtTelpKantor.Text + "',"
'''        cmdsql = cmdsql + "'" + mdiform1.txtusername.text + "',"
'''        cmdsql = cmdsql + "'" + cmbRecsource.Text + "',"
'''        cmdsql = cmdsql + "'" + TxtIdReff.Text + "',"
'''        If TdbDOB.ValueIsNull = False Then
'''            cmdsql = cmdsql + "'" + Format(TdbDOB.Value, "yyyy/mm/dd") + "',"
'''        End If
'''        cmdsql = cmdsql + "'" + cmbRecsource.Text + "')"
'''        M_OBJCONN.Execute cmdsql
''
''
''        ' munculin tuh form buat entry reff
''        With FrmEntryReff
''            .TxtIdReff.Text = "Inbound Leads"
''            .TxtNama.Text = TxtNamaLeads.Text
''            .TxtTelpRumah.Text = TxtTelpRumah.Text
''            .TxtTelpKantor.Text = TxtTelpKantor.Text
''            .TxtHandPhone.Text = TxtHandPhoneLeads.Text
''            .TdbDOB.Value = TdbDOB.Value
''            .TxtIdReff.Enabled = False
''             .Show vbModal
''             If .okReff = True Then
''                MsgBox "Data sudah tersimpan", vbInformation + vbOKOnly, "TINS"
''             Else
''                MsgBox "Cancel", vbInformation + vbOKOnly, "TINS"
''             End If
''        End With
''            Unload FrmEntryReff
''End Sub
'
''Private Sub show_Leads_Duplicate()
''Dim listitem As listitem
''ListView1.ListItems.Clear
''SSTab1.Tab = 1
''mrs_cek.MoveFirst
''While Not mrs_cek.EOF
''    Set listitem = ListView1.ListItems.ADD(, , mrs_cek.Bookmark)
''        listitem.SubItems(1) = IIf(IsNull(mrs_cek("custid")), "", mrs_cek("custid"))
''        Select Case mrs_cek("RECSTATUS")
''        Case "1A"
''            listitem.SubItems(2) = "Available"
''        Case ""
''            listitem.SubItems(2) = "Available"
''        Case Else
''            listitem.SubItems(2) = IIf(IsNull(mrs_cek("PRIOR")), "", mrs_cek("PRIOR"))
''        End Select
''        listitem.SubItems(3) = IIf(IsNull(mrs_cek("CUSTIDREF")), "", mrs_cek("CUSTIDREF"))
''        listitem.SubItems(4) = IIf(IsNull(mrs_cek("NAMAREF")), "", mrs_cek("NAMAREF"))
''        listitem.SubItems(5) = IIf(IsNull(mrs_cek("NAME")), "", mrs_cek("NAME"))
''        listitem.SubItems(6) = IIf(IsNull(mrs_cek("NEXTACTDATE")), "", Format(mrs_cek("NEXTACTDATE"), "yyyy/mm/dd hh:mm"))
''        listitem.SubItems(7) = IIf(IsNull(mrs_cek("NEXTACT")), "", mrs_cek("NEXTACT"))
''        listitem.SubItems(8) = IIf(IsNull(mrs_cek("REMARKS")), "", mrs_cek("REMARKS"))
''        listitem.SubItems(9) = IIf(IsNull(mrs_cek("AGENT")), "", mrs_cek("AGENT"))
''        listitem.SubItems(10) = IIf(IsNull(mrs_cek("NamaAGENT")), "", mrs_cek("NamaAGENT"))
''        listitem.SubItems(11) = IIf(IsNull(mrs_cek("RECSOURCEREF")), "", mrs_cek("RECSOURCEREF"))
''        listitem.SubItems(12) = Format(IIf(IsNull(mrs_cek("TGLSTATUS")), "", mrs_cek("TGLSTATUS")), "yyyy/mm/dd")
''        listitem.SubItems(13) = IIf(IsNull(mrs_cek("Kethslkerja")), "", mrs_cek("Kethslkerja"))
''        listitem.SubItems(14) = IIf(IsNull(mrs_cek("KdComplaint")), "", mrs_cek("KdComplaint"))
''        listitem.SubItems(15) = IIf(IsNull(mrs_cek("RemarkComplaint")), "", mrs_cek("RemarkComplaint"))
''        listitem.SubItems(16) = IIf(IsNull(mrs_cek("F_CEK")), "", mrs_cek("F_CEK"))
''        listitem.SubItems(17) = IIf(IsNull(mrs_cek("Nomor")), "", mrs_cek("Nomor"))
''        mrs_cek.MoveNext
''Wend
''Set mrs_cek = Nothing
''End Sub
'
'
''Private Sub CmdClaim_Click(Index As Integer)
''
''Select Case Index
''    Case 0
''        If Len(TxtNama.Text) < 2 Then
''            MsgBox "Nama harus diisi", vbInformation + vbOKOnly, "TINS"
''            Exit Sub
''        End If
''        If Len(txtnotelprumah.Text) < 2 And Len(TxtNoTelpKantor.Text) < 2 And Len(TxtHandPhone.Text) < 2 Then
''            MsgBox "Minimal salah satu dari telp harus diisi", vbInformation + vbOKOnly, "TINS"
''            Exit Sub
''        End If
''        'CmdSave.Enabled = False
''        Call cari_duplicate_CH
''    Case 1
''        TxtNama.Text = ""
''        txtnotelprumah.Text = ""
''        TxtNoTelpKantor.Text = ""
''        TxtHandPhone.Text = ""
''        TdbDOBCH.Value = Empty
''        cmbRecsourcech.Text = ""
''End Select
'
'
'
'
''Dim m_objrs As ADODB.Recordset
''Dim cmdsql As String
''Dim Lcustid, LName, LHOMENO, LOFFICENO, LMOBILE, LAgent, LNAMAAGENT, LRECSOURCE, LOTHERS, LKethslkerja As String
''Select Case Index
''Case 0
''    If Len(TxtNama.Text) = 0 Or Len(txtnotelprumah.Text) = 0 Then
''        MsgBox "Nama atau no telp rumah harus diisi", vbInformation + vbOKOnly, "Informasi"
''        Exit Sub
''    End If
''
''    Set m_objrs = New ADODB.Recordset
''    m_objrs.CursorLocation = adUseClient
''    cmdsql = "Select  * from mgm where LEFT(recsource,3) <>'PRE' AND HOMENO like '%" + txtnotelprumah.Text + "%' "
''    m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
''    While Not m_objrs.EOF
''        LKethslkerja = m_objrs!KETHSLKERJA
''        LName = m_objrs!Name
''        If UCase(Trim(LName)) = UCase(Trim(TxtNama.Text)) Then
''            If UCase(LKethslkerja) = "1A" Then
''                m_objrs!NEXTACT = "Data Inbound Call"
''                m_objrs!agent = mdiform1.txtusername.text
''                m_objrs!NAMAAGENT = mdiform1.txtnama.text
''                m_objrs.UPDATE
''                MsgBox "Sukses... Data Sudah Ditransfer", vbInformation + vbOKOnly, "Informasi"
''
''                Call isi_dataClaimKeGrid(m_objrs!CUSTID, m_objrs!Name, "Data Inbound Call", "Data Inbound Call", mdiform1.txtusername.text, mdiform1.txtnama.text, m_objrs!RECSOURCE)
''
''                Set m_objrs = Nothing
''                TxtNama.Text = ""
''                txtnotelprumah.Text = ""
''                TxtNoTelpKantor.Text = ""
''                TxtHandPhone.Text = ""
''                TxtNama.SetFocus
''                Exit Sub
''            Else
''                MsgBox "Tidak Sukses... Data Sudah di follow Up oleh " & m_objrs!agent
''                Set m_objrs = Nothing
''                TxtNama.Text = ""
''                txtnotelprumah.Text = ""
''                TxtNoTelpKantor.Text = ""
''                TxtHandPhone.Text = ""
''                TxtNama.SetFocus
''                Exit Sub
''            End If
''        End If
''    m_objrs.MoveNext
''    Wend
''    Set m_objrs = Nothing
''
''
''    Set m_objrs = New ADODB.Recordset
''    m_objrs.CursorLocation = adUseClient
''    cmdsql = "Select  * from tempCC_CUSTTBL where LEFT(recsource,3) <>'PRE' AND HOMENO like '%" + txtnotelprumah.Text + "%'"
''    m_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
''    While Not m_objrs.EOF
''        LKethslkerja = m_objrs!KETHSLKERJA
''        LName = m_objrs!Name
''        If UCase(Trim(Replace(LName, "(PVA)", ""))) = UCase(Trim(TxtNama.Text)) Then
''            Lcustid = m_objrs!CUSTID
''            LName = m_objrs!Name
''            LHOMENO = m_objrs!HOMENO
''            LOFFICENO = m_objrs!OFFICENO
''            LMOBILE = m_objrs!MOBILENO
''            LAgent = m_objrs!agent
''            LNAMAAGENT = m_objrs!NAMAAGENT
''            LRECSOURCE = m_objrs!RECSOURCE
''            LOTHERS = m_objrs!OTHERS
''            LKethslkerja = m_objrs!KETHSLKERJA
''            cmdsql = "Insert Into mgm (CUSTID, NAME, TGLDISTRIBUSI, HOMENO, OFFICENO, MOBILENO, AGENT, NAMAAGENT, RECSOURCE, nextact, KETHSLKERJA)"
''            cmdsql = cmdsql + " VALUES"
''            cmdsql = cmdsql + "('" + Lcustid + "',"
''            cmdsql = cmdsql + " '" + LName + "',"
''            cmdsql = cmdsql + " '" + Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd hh:nn") + "',"
''            cmdsql = cmdsql + " '" + LHOMENO + "',"
''            cmdsql = cmdsql + " '" + LOFFICENO + "',"
''            cmdsql = cmdsql + " '" + LMOBILE + "',"
''            cmdsql = cmdsql + " '" + mdiform1.txtusername.text + "',"
''            cmdsql = cmdsql + " '" + mdiform1.txtnama.text + "',"
''            cmdsql = cmdsql + " '" + LRECSOURCE + "',"
''            cmdsql = cmdsql + " 'Data Inbound Call',"
''            cmdsql = cmdsql + " '" + LKethslkerja + "')"
''            M_OBJCONN.Execute cmdsql
''            Call isi_dataClaimKeGrid(CStr(Lcustid), CStr(LName), "Data Inbound Call", "Data Inbound Call", mdiform1.txtusername.text, mdiform1.txtnama.text, CStr(LRECSOURCE))
''            MsgBox "Sukses... Data Sudah Ditransfer", vbInformation + vbOKOnly, "Informasi"
''            M_OBJCONN.Execute "Delete from TempCC_Custtbl where custid ='" + Lcustid + "'"
''            Set m_objrs = Nothing
''            TxtNama.Text = ""
''            txtnotelprumah.Text = ""
''            TxtNoTelpKantor.Text = ""
''            TxtHandPhone.Text = ""
''            TxtNama.SetFocus
''            Exit Sub
''        End If
''    m_objrs.MoveNext
''    Wend
''    Set m_objrs = Nothing
''
''            cmdsql = "Insert Into mgm (CUSTID, NAME, TGLDISTRIBUSI, HOMENO, OFFICENO, MOBILENO, AGENT, NAMAAGENT, RECSOURCE, nextact, KETHSLKERJA)"
''            cmdsql = cmdsql + " VALUES"
''            Lcustid = "mgmI-" & CUSTNOMOR(M_OBJCONN, "FRMCUST_CC")
''
''            cmdsql = cmdsql + "('" + Lcustid + "',"
''            cmdsql = cmdsql + " '" + TxtNama.Text + "',"
''            cmdsql = cmdsql + " '" + Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd hh:nn") + "',"
''            cmdsql = cmdsql + " '" + txtnotelprumah.Text + "',"
''            cmdsql = cmdsql + " '" + TxtNoTelpKantor.Text + "',"
''            cmdsql = cmdsql + " '" + TxtHandPhone.Text + "',"
''            cmdsql = cmdsql + " '" + mdiform1.txtusername.text + "',"
''            cmdsql = cmdsql + " '" + mdiform1.txtnama.text + "',"
''            cmdsql = cmdsql + " 'mgm_INC',"
''            cmdsql = cmdsql + " 'Data Inbound Call',"
''            cmdsql = cmdsql + " '1A')"
''    M_OBJCONN.Execute cmdsql
''    MsgBox "Sukses..", vbInformation, "Informasi"
''
''    Call isi_dataClaimKeGrid(CStr(Lcustid), TxtNama.Text, "Data Inbound Call", "Data Inbound Call", mdiform1.txtusername.text, mdiform1.txtnama.text, "mgm_INC")
''
''    TxtNama.Text = ""
''    txtnotelprumah.Text = ""
''    TxtNoTelpKantor.Text = ""
''    TxtHandPhone.Text = ""
''    TxtNama.SetFocus
''Case 1
''    TxtNama.Text = ""
''    txtnotelprumah.Text = ""
''    TxtNoTelpKantor.Text = ""
''    TxtHandPhone.Text = ""
''End Select
''End Sub
'
'
''Private Sub cari_duplicate_CH()
''    Dim CMDSQL As String
''
''    Dim kriteria1 As String
''    Dim kriteria2 As String
''    Dim CUSTID1 As String
''    ' kriteria pertama
''    'nama ama notelp
''    If Len(TxtNama.Text) > 2 And Len(txtnotelprumah.Text) > 2 Then
''        kriteria2 = Left(txtnotelprumah.Text, 5)
''        CMDSQL = "Select * from mgm where name like '%" + TxtNama.Text + "%' "
''        CMDSQL = CMDSQL + " and (HOMENO like '%" + kriteria2 + "%' or HOMENO2 like '%" + kriteria2 + "%' or mobileno like '%" + kriteria2 + "%' or mobileno2 like '%" + kriteria2 + "%' or officeno like '%" + kriteria2 + "%' or officeno2 like '%" + kriteria2 + "%') "
''
''    Set mrs_cek = New ADODB.Recordset
''        mrs_cek.CursorLocation = adUseClient
''
''        mrs_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
''        If mrs_cek.RecordCount <> 0 Then
''            CUSTID1 = Empty
''            While Not mrs_cek.EOF
''                CUSTID1 = "mgmI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC"))
''                CMDSQL = "Insert into TBL_DUPLIKASICH (CUSTID, CUSTIDBARU, NAMABARU, HOMENOBARU, OFFICENOBARU, MOBILENOBARU, AGENTBARU, "
''                If TdbDOB.ValueIsNull = False Then
''                    CMDSQL = CMDSQL + "DOB,"
''                End If
''                CMDSQL = CMDSQL + " RECSOURCEBARU) values "
''                CMDSQL = CMDSQL + "('" + mrs_cek!CustId + "',"
''                CMDSQL = CMDSQL + "'" + CUSTID1 + "',"
''                CMDSQL = CMDSQL + "'" + TxtNama.Text + "',"
''                CMDSQL = CMDSQL + "'" + txtnotelprumah.Text + "',"
''                CMDSQL = CMDSQL + "'" + TxtNoTelpKantor.Text + "',"
''                CMDSQL = CMDSQL + "'" + TxtHandPhone.Text + "',"
''                CMDSQL = CMDSQL + "'" + mdiform1.txtusername.text + "',"
''                If TdbDOB.ValueIsNull = False Then
''                    CMDSQL = CMDSQL + "'" + Format(TdbDOBCH.Value, "yyyy/mm/dd") + "',"
''                End If
''                CMDSQL = CMDSQL + "'" + cmbRecsourcech.Text + "')"
''                M_OBJCONN.Execute CMDSQL
''                mrs_cek.MoveNext
''            Wend
''
''            ' tampilin yang duplicate deh...
''                Call show_Ch_Duplicate
''                MsgBox " Nama dan Telp Rumah Ada yg sama", vbInformation + vbOKOnly, "TINS"
''            Set mrs_cek = Nothing
''            Exit Sub
''
''        End If
''        Set mrs_cek = Nothing
''    End If
''    If Len(TxtNama.Text) > 2 And Len(TxtNoTelpKantor.Text) > 2 Then
''        kriteria2 = Left(TxtNoTelpKantor.Text, 5)
''        CMDSQL = "Select * from mgm where name like '%" + TxtNama.Text + "%' "
''        CMDSQL = CMDSQL + " and (HOMENO like '%" + kriteria2 + "%' or HOMENO2 like '%" + kriteria2 + "%' or mobileno like '%" + kriteria2 + "%' or mobileno2 like '%" + kriteria2 + "%' or officeno like '%" + kriteria2 + "%' or officeno2 like '%" + kriteria2 + "%') "
''    Set mrs_cek = New ADODB.Recordset
''    mrs_cek.CursorLocation = adUseClient
''
''        mrs_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
''        If mrs_cek.RecordCount <> 0 Then
''            CUSTID1 = Empty
''            While Not mrs_cek.EOF
''                CUSTID1 = "mgmI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC"))
''
''                CMDSQL = "Insert into TBL_DUPLIKASICH (CUSTID, CUSTIDBARU, NAMABARU, HOMENOBARU, OFFICENOBARU, MOBILENOBARU, AGENTBARU, "
''                If TdbDOBCH.ValueIsNull = False Then
''                    CMDSQL = CMDSQL + "DOB,"
''                End If
''                CMDSQL = CMDSQL + " RECSOURCEBARU) values "
''                CMDSQL = CMDSQL + "('" + mrs_cek!CustId + "',"
''                CMDSQL = CMDSQL + "'" + CUSTID1 + "',"
''                CMDSQL = CMDSQL + "'" + TxtNama.Text + "',"
''                CMDSQL = CMDSQL + "'" + txtnotelprumah.Text + "',"
''                CMDSQL = CMDSQL + "'" + TxtNoTelpKantor.Text + "',"
''                CMDSQL = CMDSQL + "'" + TxtHandPhone.Text + "',"
''                CMDSQL = CMDSQL + "'" + mdiform1.txtusername.text + "',"
''                If TdbDOBCH.ValueIsNull = False Then
''                    CMDSQL = CMDSQL + "'" + Format(TdbDOBCH.Value, "yyyy/mm/dd") + "',"
''                End If
''                CMDSQL = CMDSQL + "'" + cmbRecsourcech.Text + "')"
''                M_OBJCONN.Execute CMDSQL
''                mrs_cek.MoveNext
''            Wend
''
''            ' show data
''            Call show_Ch_Duplicate
''            MsgBox " Nama dan Telp Kantor Ada yg sama", vbInformation + vbOKOnly, "TINS"
''            Set mrs_cek = Nothing
''            Exit Sub
''        End If
''        Set mrs_cek = Nothing
''
''    End If
''    If Len(TxtNama.Text) > 2 And Len(TxtHandPhone.Text) > 2 Then
''        kriteria2 = Left(TxtHandPhone.Text, 8)
''        CMDSQL = "Select * from mgm where name like '%" + TxtNama.Text + "%' "
''        CMDSQL = CMDSQL + " and (HOMENO like '%" + kriteria2 + "%' or HOMENO2 like '%" + kriteria2 + "%' or mobileno like '%" + kriteria2 + "%' or mobileno2 like '%" + kriteria2 + "%' or officeno like '%" + kriteria2 + "%' or officeno2 like '%" + kriteria2 + "%') "
''    Set mrs_cek = New ADODB.Recordset
''    mrs_cek.CursorLocation = adUseClient
''
''        mrs_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
''        If mrs_cek.RecordCount <> 0 Then
''
''
''            CUSTID1 = Empty
''            While Not mrs_cek.EOF
''                CUSTID1 = "mgmI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC"))
''
''                CMDSQL = "Insert into TBL_DUPLIKASICH (CUSTID, CUSTIDBARU, NAMABARU, HOMENOBARU, OFFICENOBARU, MOBILENOBARU, AGENTBARU, "
''                If TdbDOBCH.ValueIsNull = False Then
''                    CMDSQL = CMDSQL + "DOB,"
''                End If
''                CMDSQL = CMDSQL + " RECSOURCEBARU) values "
''                CMDSQL = CMDSQL + "('" + mrs_cek!CustId + "',"
''                CMDSQL = CMDSQL + "'" + CUSTID1 + "',"
''                CMDSQL = CMDSQL + "'" + TxtNama.Text + "',"
''                CMDSQL = CMDSQL + "'" + txtnotelprumah.Text + "',"
''                CMDSQL = CMDSQL + "'" + TxtNoTelpKantor.Text + "',"
''                CMDSQL = CMDSQL + "'" + TxtHandPhone.Text + "',"
''                CMDSQL = CMDSQL + "'" + mdiform1.txtusername.text + "',"
''                If TdbDOBCH.ValueIsNull = False Then
''                    CMDSQL = CMDSQL + "'" + Format(TdbDOBCH.Value, "yyyy/mm/dd") + "',"
''                End If
''                CMDSQL = CMDSQL + "'" + cmbRecsourcech.Text + "')"
''
''                M_OBJCONN.Execute CMDSQL
''                mrs_cek.MoveNext
''            Wend
''
''            ' show data
''            Call show_Ch_Duplicate
''            MsgBox "Nama dan Handphone Ada yg sama", vbInformation + vbOKOnly, "TINS"
''            Set mrs_cek = Nothing
''            Exit Sub
''        End If
''        Set mrs_cek = Nothing
''
''    End If
''    If Len(TxtNama.Text) > 2 And TdbDOBCH.ValueIsNull = False Then
''        kriteria2 = Format(TdbDOBCH.Value, "yyyy/mm/dd")
''        CMDSQL = "Select * from mgm where name like '%" + TxtNama.Text + "%' "
''        CMDSQL = CMDSQL + " and birthd = '" + Format(TdbDOBCH.Value, "yyyy/mm/dd") + "'"
''        Set mrs_cek = New ADODB.Recordset
''            mrs_cek.CursorLocation = adUseClient
''
''        mrs_cek.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
''        If mrs_cek.RecordCount <> 0 Then
''
''
''            CUSTID1 = Empty
''            While Not mrs_cek.EOF
''                CUSTID1 = "mgmI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC"))
''
''                CMDSQL = "Insert into TBL_DUPLIKASICH (CUSTID, CUSTIDBARU, NAMABARU, HOMENOBARU, OFFICENOBARU, MOBILENOBARU, AGENTBARU, "
''                If TdbDOBCH.ValueIsNull = False Then
''                    CMDSQL = CMDSQL + "DOB,"
''                End If
''                CMDSQL = CMDSQL + " RECSOURCEBARU) values "
''                CMDSQL = CMDSQL + "('" + mrs_cek!CustId + "',"
''                CMDSQL = CMDSQL + "'" + CUSTID1 + "',"
''                CMDSQL = CMDSQL + "'" + TxtNama.Text + "',"
''                CMDSQL = CMDSQL + "'" + txtnotelprumah.Text + "',"
''                CMDSQL = CMDSQL + "'" + TxtNoTelpKantor.Text + "',"
''                CMDSQL = CMDSQL + "'" + TxtHandPhone.Text + "',"
''                CMDSQL = CMDSQL + "'" + mdiform1.txtusername.text + "',"
''                If TdbDOBCH.ValueIsNull = False Then
''                    CMDSQL = CMDSQL + "'" + Format(TdbDOBCH.Value, "yyyy/mm/dd") + "',"
''                End If
''                CMDSQL = CMDSQL + "'" + cmbRecsourcech.Text + "')"
''
''                M_OBJCONN.Execute CMDSQL
''                mrs_cek.MoveNext
''            Wend
''
''            ' show data
''            Call show_Ch_Duplicate
''            MsgBox "Nama dan DOB Ada yg sama", vbInformation + vbOKOnly, "TINS"
''            Set mrs_cek = Nothing
''            Exit Sub
''        End If
''        Set mrs_cek = Nothing
''    End If
'''        custid1 = "mgmI-" & CUSTNOMOR(M_OBJCONN, UCase("FRMCUST_CC"))
'''
'''        cmdsql = "Insert into mgm(CUSTID, NAME, HOMENO, MOBILENO, OFFICENO, AGENT, RECSOURCE,CustIdRef,"
'''        If TdbDobCh.ValueIsNull = False Then
'''            cmdsql = cmdsql + "BIRTHD,"
'''        End If
'''        cmdsql = cmdsql + " RecSourceRef) values"
'''        cmdsql = cmdsql + "('" + custid1 + "',"
'''        cmdsql = cmdsql + "'" + TxtNama.Text + "',"
'''        cmdsql = cmdsql + "'" + txtnotelprumah.Text + "',"
'''        cmdsql = cmdsql + "'" + TxtHandphone.Text + "',"
'''        cmdsql = cmdsql + "'" + TxtNoTelpKantor.Text + "',"
'''        cmdsql = cmdsql + "'" + mdiform1.txtusername.text + "',"
'''        cmdsql = cmdsql + "'" + CmbRecsourceCh.Text + "',"
'''        cmdsql = cmdsql + "'" + TxtIdReff.Text + "',"
'''        If TdbDobCh.ValueIsNull = False Then
'''            cmdsql = cmdsql + "'" + Format(TdbDobCh.Value, "yyyy/mm/dd") + "',"
'''        End If
'''        cmdsql = cmdsql + "'" + CmbRecsourceCh.Text + "')"
'''        M_OBJCONN.Execute cmdsql
''
''
''        ' munculin tuh form buat entry mgm
''        With FrmEntryCH
''            .TxtIdReff.Text = "Inbound mgm"
''            .TxtNama.Text = TxtNama.Text
''            .TxtTelpRumah.Text = txtnotelprumah.Text
''            .TxtTelpKantor.Text = TxtNoTelpKantor.Text
''            .TxtHandPhone.Text = TxtHandPhone.Text
''            .TdbDOBCH.Value = TdbDOBCH.Value
''            .TxtIdReff.Enabled = False
''             .Show vbModal
''             If .okReff = True Then
''                MsgBox "Data sudah tersimpan", vbInformation + vbOKOnly, "TINS"
''             Else
''                MsgBox "Cancel", vbInformation + vbOKOnly, "TINS"
''             End If
''        End With
''            Unload FrmEntryCH
''
''
''End Sub
'
''Private Sub show_Ch_Duplicate()
''Dim listitem As listitem
''SSTab1.Tab = 0
''
''On Error GoTo HELL
''    mrs_cek.MoveFirst
''    LstVwSearchmgm.ListItems.Clear
''
''    While Not mrs_cek.EOF
''        Set listitem = LstVwSearchmgm.ListItems.ADD(, , mrs_cek.Bookmark)
''        listitem.SubItems(1) = IIf(IsNull(mrs_cek("CUSTID")), "", mrs_cek("CUSTID"))
''        listitem.SubItems(2) = IIf(IsNull(mrs_cek("PRIOR")), "", mrs_cek("PRIOR"))
''        listitem.SubItems(3) = IIf(IsNull(mrs_cek("NAME")), "", mrs_cek("NAME"))
''        listitem.SubItems(4) = IIf(IsNull(mrs_cek("NEXTACTDATE")), "", Format(mrs_cek("NEXTACTDATE"), "yyyy/mm/dd hh:nn"))
''        listitem.SubItems(5) = IIf(IsNull(mrs_cek("NEXTACT")), "", mrs_cek("NEXTACT"))
''        listitem.SubItems(6) = IIf(IsNull(mrs_cek("REMARKS")), "", mrs_cek("REMARKS"))
''        listitem.SubItems(7) = IIf(IsNull(mrs_cek("AGENT")), "", mrs_cek("AGENT"))
''        listitem.SubItems(8) = IIf(IsNull(mrs_cek("NamaAGENT")), "", mrs_cek("NamaAGENT"))
''        listitem.SubItems(9) = IIf(IsNull(mrs_cek("RECSOURCE")), "", mrs_cek("RECSOURCE"))
''        listitem.SubItems(10) = IIf(IsNull(mrs_cek("TGLSTATUS")), "", Format(mrs_cek("TGLSTATUS"), "DD/MM/YYYY"))
''        listitem.SubItems(11) = IIf(IsNull(mrs_cek("Kethslkerja")), "", mrs_cek("Kethslkerja"))
''        listitem.SubItems(12) = IIf(IsNull(mrs_cek("KdComplaint")), "", mrs_cek("KdComplaint"))
''        listitem.SubItems(13) = IIf(IsNull(mrs_cek("RemarkComplaint")), "", mrs_cek("RemarkComplaint"))
''        listitem.SubItems(14) = IIf(IsNull(mrs_cek("F_CEK")), "", mrs_cek("F_CEK"))
''        'LISTITEM.SubItems(15) = IIf(IsNull(m_cari("[NO]")), "", m_cari("[NO]"))
''        mrs_cek.MoveNext
''    Wend
''        If LstVwSearchmgm.ListItems.count = 0 Then
''            TxtJmlDtmgm.Text = "Tidak Ada Data"
''        Else
''            TxtJmlDtmgm.Text = "Total " + CStr(mrs_cek.RecordCount) + " Records"
''        End If
''LstVwSearchmgm.SortKey = 2
''LstVwSearchmgm.Sorted = True
''Exit Sub
''HELL:
''    Me.MousePointer = vbNormal
''    MsgBox Err.Description
''  ''  Resume
''End Sub
''
'Private Sub LstVwSearchmgm_KeyUp(KeyCode As Integer, Shift As Integer)
'
'End Sub
'
