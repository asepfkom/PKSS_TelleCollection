VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Begin VB.Form FrmSearching 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Searching Data"
   ClientHeight    =   8850
   ClientLeft      =   -540
   ClientTop       =   1410
   ClientWidth     =   11940
   Icon            =   "FrmSearching.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   11940
   Begin Threed.SSCommand Command1 
      Cancel          =   -1  'True
      Height          =   345
      Index           =   1
      Left            =   10845
      TabIndex        =   0
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
      Height          =   2175
      Left            =   15
      TabIndex        =   1
      Top             =   0
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   3836
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
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
      TabPicture(0)   =   "FrmSearching.frx":1272
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Schedulle"
      TabPicture(1)   =   "FrmSearching.frx":128E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
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
         Height          =   1680
         Left            =   -74910
         TabIndex        =   14
         Top             =   390
         Width           =   11550
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Proses....!!"
            Height          =   510
            Left            =   8715
            TabIndex        =   33
            Top             =   990
            Visible         =   0   'False
            Width           =   1995
            Begin MSComctlLib.ProgressBar ProgressBar1 
               Height          =   270
               Left            =   15
               TabIndex        =   34
               Top             =   165
               Width           =   2025
               _ExtentX        =   3572
               _ExtentY        =   476
               _Version        =   393216
               Appearance      =   0
            End
         End
         Begin VB.TextBox Text1 
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
            TabIndex        =   23
            Top             =   345
            Width           =   2025
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
            Index           =   0
            Left            =   4155
            TabIndex        =   22
            Top             =   660
            Visible         =   0   'False
            Width           =   960
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
            Index           =   1
            Left            =   5115
            TabIndex        =   21
            Top             =   660
            Visible         =   0   'False
            Width           =   2310
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
            Index           =   2
            Left            =   6900
            TabIndex        =   20
            Top             =   315
            Visible         =   0   'False
            Width           =   1560
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
            Index           =   3
            Left            =   8460
            TabIndex        =   19
            Top             =   315
            Visible         =   0   'False
            Width           =   2940
         End
         Begin VB.TextBox Text1 
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
            TabIndex        =   18
            Top             =   345
            Width           =   1965
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
            TabIndex        =   17
            Top             =   30
            Visible         =   0   'False
            Width           =   3225
         End
         Begin VB.ComboBox CmbStatusCek 
            Height          =   315
            Left            =   8745
            TabIndex        =   16
            Top             =   630
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.ComboBox cmbStsLastCall 
            Height          =   315
            Index           =   0
            Left            =   4155
            TabIndex        =   15
            Top             =   960
            Visible         =   0   'False
            Width           =   3180
         End
         Begin TDBDate6Ctl.TDBDate TdDob 
            Height          =   315
            Left            =   1200
            TabIndex        =   24
            Top             =   660
            Visible         =   0   'False
            Width           =   1170
            _Version        =   65536
            _ExtentX        =   2064
            _ExtentY        =   556
            Calendar        =   "FrmSearching.frx":12AA
            Caption         =   "FrmSearching.frx":13C2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FrmSearching.frx":142E
            Keys            =   "FrmSearching.frx":144C
            Spin            =   "FrmSearching.frx":14AA
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
            TabIndex        =   25
            Top             =   975
            Visible         =   0   'False
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   556
            Caption         =   "FrmSearching.frx":14D2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "FrmSearching.frx":153E
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
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
            TabIndex        =   26
            Top             =   1290
            Visible         =   0   'False
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   556
            Caption         =   "FrmSearching.frx":1580
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "FrmSearching.frx":15EC
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
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
            TabIndex        =   27
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
            TabIndex        =   28
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
            TabIndex        =   29
            Top             =   1275
            Visible         =   0   'False
            Width           =   825
            _Version        =   65536
            _ExtentX        =   1455
            _ExtentY        =   529
            Caption         =   "FrmSearching.frx":162E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "FrmSearching.frx":169A
            Spin            =   "FrmSearching.frx":16EA
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
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
            TabIndex        =   30
            Top             =   1275
            Visible         =   0   'False
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   556
            Calendar        =   "FrmSearching.frx":1712
            Caption         =   "FrmSearching.frx":182A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FrmSearching.frx":1896
            Keys            =   "FrmSearching.frx":18B4
            Spin            =   "FrmSearching.frx":1912
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
            TabIndex        =   31
            Top             =   1275
            Visible         =   0   'False
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2249
            _ExtentY        =   556
            Calendar        =   "FrmSearching.frx":193A
            Caption         =   "FrmSearching.frx":1A52
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FrmSearching.frx":1ABE
            Keys            =   "FrmSearching.frx":1ADC
            Spin            =   "FrmSearching.frx":1B3A
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
            TabIndex        =   32
            Top             =   1275
            Visible         =   0   'False
            Width           =   825
            _Version        =   65536
            _ExtentX        =   1455
            _ExtentY        =   529
            Caption         =   "FrmSearching.frx":1B62
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "FrmSearching.frx":1BCE
            Spin            =   "FrmSearching.frx":1C1E
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
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
            TabIndex        =   45
            Top             =   375
            Width           =   1485
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
            TabIndex        =   44
            Top             =   720
            Visible         =   0   'False
            Width           =   1365
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
            TabIndex        =   43
            Top             =   705
            Visible         =   0   'False
            Width           =   1050
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
            TabIndex        =   42
            Top             =   360
            Visible         =   0   'False
            Width           =   690
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
            TabIndex        =   41
            Top             =   1035
            Visible         =   0   'False
            Width           =   1005
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
            TabIndex        =   40
            Top             =   1365
            Visible         =   0   'False
            Width           =   1005
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
            TabIndex        =   39
            Top             =   360
            Width           =   1125
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
            TabIndex        =   38
            Top             =   1305
            Visible         =   0   'False
            Width           =   225
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
            TabIndex        =   37
            Top             =   675
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
            TabIndex        =   36
            Top             =   1335
            Visible         =   0   'False
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
            Left            =   2790
            TabIndex        =   35
            Top             =   1035
            Visible         =   0   'False
            Width           =   1365
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFC0C0&
         Height          =   1725
         Left            =   75
         TabIndex        =   2
         Top             =   375
         Width           =   11565
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   4
            Left            =   3570
            TabIndex        =   10
            Top             =   225
            Visible         =   0   'False
            Width           =   1905
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
            TabIndex        =   4
            Top             =   210
            Width           =   3480
         End
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
            TabIndex        =   3
            Top             =   555
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1245
         End
         Begin Threed.SSFrame SSFrame1 
            Height          =   510
            Left            =   2970
            TabIndex        =   5
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
               TabIndex        =   6
               Top             =   75
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               Calendar        =   "FrmSearching.frx":1C46
               Caption         =   "FrmSearching.frx":1D5E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FrmSearching.frx":1DCA
               Keys            =   "FrmSearching.frx":1DE8
               Spin            =   "FrmSearching.frx":1E46
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
               TabIndex        =   7
               Top             =   90
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               Calendar        =   "FrmSearching.frx":1E6E
               Caption         =   "FrmSearching.frx":1F86
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FrmSearching.frx":1FF2
               Keys            =   "FrmSearching.frx":2010
               Spin            =   "FrmSearching.frx":206E
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
               TabIndex        =   8
               Top             =   90
               Visible         =   0   'False
               Width           =   285
            End
         End
         Begin Threed.SSCommand CmdScheduleoK 
            Height          =   390
            Left            =   6645
            TabIndex        =   9
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
         Begin Threed.SSCommand CmdToday 
            Height          =   480
            Left            =   120
            TabIndex        =   11
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
            TabIndex        =   12
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
            TabIndex        =   13
            Top             =   255
            Width           =   1170
         End
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   0
      TabIndex        =   46
      Top             =   2250
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   12515
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Debitur Info"
      TabPicture(0)   =   "FrmSearching.frx":2096
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "LblTarget(0)"
      Tab(0).Control(1)=   "Check1"
      Tab(0).Control(2)=   "Frame1"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Address"
      TabPicture(1)   =   "FrmSearching.frx":20B2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "LblTarget(1)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Kredit"
      TabPicture(2)   =   "FrmSearching.frx":20CE
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ListView2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Phone"
      TabPicture(3)   =   "FrmSearching.frx":20EA
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "ListView3"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Jobs"
      TabPicture(4)   =   "FrmSearching.frx":2106
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "ListView4"
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame4 
         BackColor       =   &H80000004&
         Height          =   8685
         Left            =   -74940
         TabIndex        =   52
         Top             =   645
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
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   7980
            Width           =   3045
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   7845
            Left            =   -30
            TabIndex        =   54
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
      Begin VB.Frame Frame1 
         ForeColor       =   &H00000000&
         Height          =   6720
         Left            =   -74940
         TabIndex        =   48
         Top             =   615
         Width           =   11595
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
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   6315
            Width           =   3045
         End
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
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   6315
            Width           =   3045
         End
         Begin MSComctlLib.ListView LstVwSearchMgm 
            Height          =   6135
            Left            =   0
            TabIndex        =   51
            Top             =   120
            Width           =   11520
            _ExtentX        =   20320
            _ExtentY        =   10821
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   0
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
         Left            =   -72105
         MaskColor       =   &H000000FF&
         TabIndex        =   47
         Top             =   345
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   3225
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   6885
         Left            =   -74880
         TabIndex        =   57
         Top             =   600
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   12144
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
      Begin MSComctlLib.ListView ListView3 
         Height          =   6885
         Left            =   240
         TabIndex        =   58
         Top             =   600
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   12144
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
      Begin MSComctlLib.ListView ListView4 
         Height          =   6885
         Left            =   -74880
         TabIndex        =   59
         Top             =   600
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   12144
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
         TabIndex        =   56
         Top             =   285
         Width           =   9465
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
         Left            =   -71685
         TabIndex        =   55
         Top             =   285
         Width           =   4605
      End
   End
End
Attribute VB_Name = "FrmSearching"
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
'Dim CMDSQL As String
'Dim LUserType As String
'Dim F_CEK As String
'Dim f_Pending As String
'Private Sub HEADER_VIEW_Refferall()
'    ListView1.ColumnHeaders.ADD 1, , "No", 3 * TXT
'    ListView1.ColumnHeaders.ADD 2, , "Reff_Num", 5 * TXT
'    ListView1.ColumnHeaders.ADD 3, , "DebiturInfo_Id", 5 * TXT
'    ListView1.ColumnHeaders.ADD 4, , "Debitur_Name", 10 * TXT
'    ListView1.ColumnHeaders.ADD 5, , "Born_Place", 10 * TXT
'    ListView1.ColumnHeaders.ADD 6, , "Born_Place2", 25 * TXT
'    ListView1.ColumnHeaders.ADD 7, , "Born_Date", 10 * TXT
'    ListView1.ColumnHeaders.ADD 8, , "Din", 12 * TXT
'    ListView1.ColumnHeaders.ADD 9, , "NPWP", 17 * TXT
'    ListView1.ColumnHeaders.ADD 10, , "KTP_Num", 17 * TXT
'    ListView1.ColumnHeaders.ADD 11, , "KTP_Num2", 8 * TXT
'    ListView1.ColumnHeaders.ADD 12, , "Pasport_Num", 8 * TXT
'    ListView1.ColumnHeaders.ADD 13, , "Pasport_Num2", 10 * TXT
'    ListView1.ColumnHeaders.ADD 14, , "Matching_Result", 10 * TXT
''    ListView1.ColumnHeaders.ADD 15, , "Code", 5 * TXT
''    ListView1.ColumnHeaders.ADD 16, , "Complaint Note", 15 * TXT
''    ListView1.ColumnHeaders.ADD 17, , "Check", 10 * TXT
''    ListView1.ColumnHeaders.ADD 18, , "ID", 10 * TXT
'End Sub
'Private Sub isi_dataClaimKeGrid(gCUSTID As String, gNama As String, gnextact As String, gremarks As String, gagent As String, gnamaagent As String, grecsource As String)
'    ' insert ke grid list view
'Dim listitem As listitem
'Set listitem = LstVwSearchmgm.ListItems.ADD(, , "9999")
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
'
'
'If Check2.Value = 1 Then
'
'    LstVwSearchmgm.ListItems.Clear
'    SSTab1.Tab = 0
'    ' searching schedule mgm
'  Call CEK_STATUS_F_CEK
'
'   Set m_cari = M_DATA.QUERY_SEARCH_mgm(M_OBJCONN, "AGENT = '" + Combo1(4).Text + "' AND (NEXTACTDATE BETWEEN '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " 00:00" + "' AND '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " 23:59" + "') ", MDIForm1.Text3.Text, F_CEK, f_Pending)
'    ProgressBar1.Max = m_cari.RecordCount + 1
'    If Check2.Value = 1 Then
'        TxtJmlDtmgm.Text = m_cari.RecordCount & " Data"
'    Else
'        Text2.Text = m_cari.RecordCount & " Data"
'    End If
'
'    While Not m_cari.EOF
'    ProgressBar1.Value = m_cari.Bookmark
'
'    Set listitem = LstVwSearchmgm.ListItems.ADD(, , m_cari.Bookmark)
'        listitem.SubItems(1) = IIf(IsNull(m_cari("CUSTID")), "", m_cari("CUSTID"))
'        listitem.SubItems(2) = IIf(IsNull(m_cari("PRIOR")), "", m_cari("PRIOR"))
'        listitem.SubItems(3) = IIf(IsNull(m_cari("NAME")), "", m_cari("NAME"))
'        listitem.SubItems(4) = IIf(IsNull(m_cari("NEXTACTDATE")), "", Format(m_cari("NEXTACTDATE"), "yyyy/mm/dd hh:nn"))
'        listitem.SubItems(5) = IIf(IsNull(m_cari("NEXTACT")), "", m_cari("NEXTACT"))
'        listitem.SubItems(6) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS")) & "-" & IIf(IsNull(m_cari("f_pending")), "", m_cari("f_pending"))
'        listitem.SubItems(7) = IIf(IsNull(m_cari("F_cek")), "", m_cari("F_cek"))
'        listitem.SubItems(9) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
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
'        If LstVwSearchmgm.ListItems.Count = 0 Then
'            TxtJmlDtmgm.Text = "Tidak Ada Data"
'            TxtJmlVolmgm.Text = "0"
'        Else
'            TxtJmlDtmgm.Text = "Total " + CStr(m_cari.RecordCount) + " Records"
'            TxtJmlVolmgm.Text = "Total " + CStr(m_cari.RecordCount)
'        End If
'
'Else
'    ' searching schedule leads
'    Set m_cari = M_DATA.QUERY_SEARCH(M_OBJCONN, "AGENT = '" + Combo1(4).Text + "' AND (NEXTACTDATE BETWEEN '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " 00:00" + "' AND '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " 23:59" + "') ", MDIForm1.Text3.Text)
'        ListView1.ListItems.Clear
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
'    Set m_cari = M_DATA.QUERY_SEARCH_mgm(M_OBJCONN, "AGENT = '" + Combo1(4).Text + "'  AND (NEXTACTDATE < '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " 00:00" + "') ", MDIForm1.Text3.Text, F_CEK, f_Pending)
'       LstVwSearchmgm.ListItems.Clear
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
'    Set listitem = LstVwSearchmgm.ListItems.ADD(, , m_cari.Bookmark)
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
'        If LstVwSearchmgm.ListItems.Count = 0 Then
'            TxtJmlDtmgm.Text = "Tidak Ada Data"
'            TxtJmlVolmgm.Text = "0"
'        Else
'            TxtJmlDtmgm.Text = "Total " + CStr(m_cari.RecordCount) + " Records"
'            TxtJmlVolmgm.Text = "Total " + CStr(m_cari.RecordCount)
'        End If
'
'Else
'    Set m_cari = M_DATA.QUERY_SEARCH(M_OBJCONN, "AGENT = '" + Combo1(4).Text + "'  AND (NEXTACTDATE < '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " 00:00" + "') ", MDIForm1.Text3.Text)
'        ListView1.ListItems.Clear
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
'Dim CMDSQL As String
'Dim m_objrs As New ADODB.Recordset
'
'
'Set m_objrs = New ADODB.Recordset
'        m_objrs.CursorLocation = adUseClient
'        CMDSQL = "SELECT F_NA, F_OP, F_PTP, F_BP, F_POP, F_SP, F_UC,F_RP,usertype FROM usertbl WHERE USERID = '" + mdiform1.txtusername.text + "'"
'         m_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
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
'                StsUC = CStr(Trim(IIf(IsNull(m_objrs!f_uc), "", m_objrs!f_uc)))
'                LUserType = CStr(Trim(IIf(IsNull(m_objrs!usertype), "", m_objrs!usertype)))
'                m_objrs.MoveNext
'            Wend
'            Set m_objrs = Nothing
'             StsAll = StsNa + StsOP + StsPTP + StsBP + StsPOP + StsSP + StsRP + StsUC
'
'     If StsAll <> Empty Then
'            If LUserType = "1" Then
'                    If StsUC <> Empty Then
'                        F_CEK = "(left(F_CEK,3)IN('NK-','MV-','WN-','" + StsNa + "','" + StsOP + "','" + StsPTP + "','" + StsBP + "','" + StsPOP + "','" + StsSP + "','" + StsRP + "') or F_CEK IS NULL)"
''
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
'Dim CMDSQL As String
'
'Dim VOLUMEAMOUNT As Double
'If Check2.Value = 1 Then
'    Call CEK_STATUS_F_CEK
'    Set m_cari = M_DATA.QUERY_SEARCH_mgm(M_OBJCONN, "AGENT = '" + Combo1(4).Text + "'  AND (NEXTACTDATE BETWEEN '" + Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") & " 00:00" + "' AND '" + Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") & " 23:59" + "') ", MDIForm1.Text3.Text, F_CEK, f_Pending)
'       LstVwSearchmgm.ListItems.Clear
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
'    Set listitem = LstVwSearchmgm.ListItems.ADD(, , m_cari.Bookmark)
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
'        If LstVwSearchmgm.ListItems.Count = 0 Then
'            TxtJmlDtmgm.Text = "Tidak Ada Data"
'            TxtJmlVolmgm.Text = "0"
'        Else
'            TxtJmlDtmgm.Text = "Total " + CStr(m_cari.RecordCount) + " Records"
'            TxtJmlVolmgm.Text = "Total " + CStr(m_cari.RecordCount)
'        End If
'
'Else
'    Set m_cari = M_DATA.QUERY_SEARCH(M_OBJCONN, "AGENT = '" + Combo1(4).Text + "'  AND (NEXTACTDATE BETWEEN '" + Format(TDBDate1(0).Value, "yyyy/mm/dd") & " 00:00" + "' AND '" + Format(TDBDate1(1).Value, "yyyy/mm/dd") & " 23:59" + "') ", MDIForm1.Text3.Text)
'        ListView1.ListItems.Clear
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
'   lReturn = SendMessage(Combo1(Index).hWnd, CB_FINDSTRING, -1, ByVal sSearchText)
'   If lReturn <> CB_ERR Then
'      mbIgnoreListClick = True
'      Combo1(Index).ListIndex = lReturn
'      mbIgnoreListClick = False
'      Combo1(Index).Text = Combo1(Index).LIST(lReturn)
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
''SSTab1.TabVisible(1) = False
''SSTab2.TabVisible(1) = False
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
'Dim CMDSQL As String
'Dim sPending As String
'Dim m_objrs As ADODB.Recordset
'Dim VOLUMEAMOUNT As Double
'i = 1
'On Error GoTo HELL
'
'
'    LstVwSearchmgm.ListItems.Clear
'    Me.MousePointer = vbHourglass
'    ProgressBar1.Max = m_cari.RecordCount + 1
'    While Not m_cari.EOF
'    ProgressBar1.Value = m_cari.Bookmark
'        Lcustid1 = CStr(IIf(IsNull(m_cari!CustId), "", m_cari!CustId))
'        sPending = CStr(Trim(IIf(IsNull(m_cari!f_Pending), "", m_cari!f_Pending)))
'        If sPending = "OK" Then sPending = ""
'
'        Set listitem = LstVwSearchmgm.ListItems.ADD(, , m_cari.Bookmark)
'        listitem.SubItems(1) = IIf(IsNull(m_cari("CUSTID")), "", m_cari("CUSTID"))
'        listitem.SubItems(2) = IIf(IsNull(m_cari("PRIOR")), "", m_cari("PRIOR"))
'        listitem.SubItems(3) = IIf(IsNull(m_cari("NAME")), "", m_cari("NAME"))
'        listitem.SubItems(4) = IIf(IsNull(m_cari("NEXTACTDATE")), "", Format(m_cari("NEXTACTDATE"), "yyyy/mm/dd hh:nn"))
'        listitem.SubItems(5) = IIf(IsNull(m_cari("NEXTACT")), "", m_cari("NEXTACT"))
'        listitem.SubItems(6) = IIf(IsNull(m_cari("REMARKS")), "", m_cari("REMARKS"))
'        listitem.SubItems(7) = CStr(IIf(IsNull(m_cari("F_cek")), "", m_cari("F_cek")) & " " & sPending)
''       listitem.SubItems(8) = IIf(IsNull(m_cari("NamaAGENT")), "", m_cari("NamaAGENT"))
'        listitem.SubItems(9) = IIf(IsNull(m_cari("AGENT")), "", m_cari("AGENT"))
'        listitem.SubItems(10) = Format(IIf(IsNull(m_cari("Principal")), 0, m_cari("Principal")), "##,###")
'        listitem.SubItems(11) = Format(IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo")), "##,###")
'        listitem.SubItems(12) = Format(IIf(IsNull(m_cari("OpenDate")), "", m_cari("OpenDate")), "yyyy/mm/dd")
'        listitem.SubItems(13) = Format(IIf(IsNull(m_cari("TtlPTP")), 0, m_cari("TtlPTP")), "##,###")
'        listitem.SubItems(14) = IIf(IsNull(m_cari("RECSOURCE")), "", m_cari("RECSOURCE"))
'        listitem.SubItems(15) = IIf(IsNull(m_cari("TGLSTATUS")), "", Format(m_cari("TGLSTATUS"), "YYYY/MM/DD"))
'        listitem.SubItems(16) = IIf(IsNull(m_cari("Kethslkerja")), "", m_cari("Kethslkerja"))
'        listitem.SubItems(17) = Format(IIf(IsNull(m_cari("TGLINCOMING")), "", m_cari("TGLINCOMING")), "YYYY/MM/DD")
'
''         While Not m_cari2.EOF
''                Lcustid2 = CStr(IIf(IsNull(m_cari2!CustId), "", m_cari2!CustId))
''                LCall = CInt(IIf(IsNull(m_cari2!Call), 0, m_cari2!Call))
''                If Lcustid2 = Lcustid1 Then
''                   listitem.SubItems(8) = LCall
''                        Else
''                            GoTo SorryLompat
''                End If
''                     m_cari2.MoveNext
''         Wend
'
'SorryLompat:
'        'listitem.SubItems(19) = IIf(IsNull(m_cari("F_CEK")), "", m_cari("F_CEK"))
'        VOLUMEAMOUNT = VOLUMEAMOUNT + IIf(IsNull(m_cari("AmountWo")), 0, m_cari("AmountWo"))
'        'LISTITEM.SubItems(15) = IIf(IsNull(m_cari("[NO]")), "", m_cari("[NO]"))
''            Set m_objrs = New ADODB.Recordset
''                    m_objrs.CursorLocation = adUseClient
''                    CMDSQL = "SELECT count(custid) as callInit from mgm_hst where custid ='" + LstVwSearchmgm.SelectedItem.SubItems(1) + "'  group by custid "
''                    m_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
''                    While Not m_objrs.EOF
''                    'Set listitem = LstVwSearchmgm.ListItems.ADD(, , m_cari.Bookmark)
''                        listitem.SubItems(18) = IIf(IsNull(m_objrs("callinit")), "", m_objrs("callinit"))
''                        m_objrs.MoveNext
''                        Wend
''                        'm_objrs.Close
''                        Set m_objrs = Nothing
'        m_cari.MoveNext
'    Wend
'
'        If LstVwSearchmgm.ListItems.Count = 0 Then
'            TxtJmlDtmgm.Text = "Tidak Ada Data"
'            TxtJmlVolmgm.Text = "0"
'        Else
'            TxtJmlDtmgm.Text = "Total " + CStr(m_cari.RecordCount) + " Records"
'            TxtJmlVolmgm.Text = "Total " + CStr(Format(VOLUMEAMOUNT, "##,###"))
'        End If
'LstVwSearchmgm.SortKey = 2
'LstVwSearchmgm.Sorted = True
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
'    LstVwSearchmgm.ColumnHeaders.ADD 1, , "No", 3 * TXT
'    LstVwSearchmgm.ColumnHeaders.ADD 2, , "Reff_Num", 5 * TXT
'    LstVwSearchmgm.ColumnHeaders.ADD 3, , "No_Laporan", 5 * TXT
'    LstVwSearchmgm.ColumnHeaders.ADD 4, , "Tgl_Laporan", 10 * TXT
'    LstVwSearchmgm.ColumnHeaders.ADD 5, , "Idi_User", 5 * TXT
'    LstVwSearchmgm.ColumnHeaders.ADD 6, , "Debitur_Name", 10 * TXT
'    LstVwSearchmgm.ColumnHeaders.ADD 7, , "Born_Place", 17 * TXT
'    LstVwSearchmgm.ColumnHeaders.ADD 8, , "Born_Place2", 17 * TXT
'    LstVwSearchmgm.ColumnHeaders.ADD 9, , "Born_Date", 5 * TXT
'    LstVwSearchmgm.ColumnHeaders.ADD 10, , "Din", 5 * TXT
'    LstVwSearchmgm.ColumnHeaders.ADD 11, , "NPWP", 10 * TXT
'    LstVwSearchmgm.ColumnHeaders.ADD 12, , "NPWP2", 10 * TXT
'    LstVwSearchmgm.ColumnHeaders.ADD 13, , "KTP_NUM", 5 * TXT
'    LstVwSearchmgm.ColumnHeaders.ADD 14, , "KTP_NUM2", 5 * TXT
'    LstVwSearchmgm.ColumnHeaders.ADD 15, , "Passport_Num", 5 * TXT
'    LstVwSearchmgm.ColumnHeaders.ADD 16, , "Passport_Num2", 5 * TXT
'    LstVwSearchmgm.ColumnHeaders.ADD 17, , "HTML_SOURCE", 5 * TXT
'End Sub
'  Sub WaitSecs(Seconds As Single)
' Dim a As Long
' Seconds = Seconds + Timer
' While Seconds > Timer
'  a = DoEvents
' Wend
'End Sub
'Private Sub Command1_Click(Index As Integer)
'Dim Reff_Num As String
'Dim No_Laporan As String
'Dim Tgl_Laporan As String
'Dim Idi_User As String
'Dim Debitur_Name As String
'Dim Born_Place As String
'Dim Born_Place2 As String
'Dim Born_Date As String
'Dim Din As String
'Dim NPWP As String
'Dim NPWP2 As String
'Dim KTP_NUM As String
'Dim KTP_NUM2 As String
'Dim Passport_Num As String
'Dim Passport_Num2 As String
'Dim HTML_SOURCE As String
'
'
'Dim M_DATA As New CLS_FRMSEARCH
'Dim m_objrs As New ADODB.Recordset
'Dim PANJANG As Integer
'
'Select Case Index
'    Case 0
'    Call CEK_STATUS_F_CEK
'
'        If Trim(Text1(0).Text) = Empty And Trim(Combo1(0).Text) = Empty And Combo1(2).Text = Empty And Len(TDBMask2.Value) < 1 And Len(TDBMask1.Value) < 1 And TdDob.ValueIsNull And Len(Text1(2).Text) < 3 _
'        And cmbStsLastCall(0).Text = "" And CmbStatusCek.Text = "" And DtLastCall(0).ValueIsNull And CekDtDistribute.Value = 0 Then
'            MsgBox "Masukan Kriteria Customer Yang Akan Dicari...!!!", vbCritical + vbOKOnly, "Peringatan"
'            Text1(0).SetFocus
'            Set M_DATA = Nothing
'            Exit Sub
'        Else
'
'         LstVwSearchmgm.ListItems.Clear
'         Frame3.Visible = True
''         If CekDtDistribute.Value = 1 Then
''            NAMAAGENT = "AGENT is null"
''         Else
'            If Text1(2).Text <> Empty Then
'                Reff_Num = "Reff_Num LIKE " + "'%" + UBAH_QUOTE(Text1(2).Text) + "%'"
'            Else
'                If Text1(0).Text <> Empty Then
'                    Debitur_Name = "Debitur_Name LIKE " + "'%" + UBAH_QUOTE(Text1(0).Text) + "%'"
'                End If
''                If Combo1(0).Text <> Empty Then
''                    NAMAAGENT = "AGENT = '" + Trim(Combo1(0).Text) + "'"
''                End If
''                If Combo1(2).Text <> Empty Then
''                    DATASOURCE = "RECSOURCE = '" + Trim(Combo1(2).Text) + "'"
''                End If
''                If TdDob.ValueIsNull Then
''                Else
''                    TGLLAHIR = "DOB = '" + Format(TdDob.Text, "yyyy/mm/dd") + "'"
''                End If
''                If Len(TDBMask1.Value) > 1 Then
''                    OFFPHONE = "OFFICENO Like '%" + TDBMask1.Value + "%'"
''                    OFFPHONE2 = "OFFICENO2 Like '%" + TDBMask1.Value + "%'"
''                    HOMEPHONE = "HOMENO Like '%" + TDBMask1.Value + "%'"
''                    HOMEPHONE2 = "HOMENO2 Like '%" + TDBMask1.Value + "%'"
''                    FAXPHONE = "FAXNO Like '%" + TDBMask1.Value + "%'"
''                    FAXPHONE2 = "FAXNO2 Like '%" + TDBMask1.Value + "%'"
''                End If
''                If Len(TDBMask2.Value) > 1 Then
''                    MOBILEPHONE = "MOBILENO like '%" + TDBMask2.Value + "%'"
''                    MOBILEPHONE2 = "MOBILENO2 like '%" + TDBMask2.Value + "%'"
''                End If
''
''                If cmbStsLastCall(0).Text <> "" Then
''                    Select Case UCase(Trim(cmbStsLastCall(0).Text))
''                        Case "NOT CONTACTED"
''                            KETHSLKERJA = "KETHSLKERJA IN('WN','WN','NP','BT')"
''                        Case "NOT AVAILABLE"
''                            KETHSLKERJA = "KETHSLKERJA = 'NA'"
''                        Case "STILL THINKING"
''                            KETHSLKERJA = "KETHSLKERJA= 'ST'"
''                        Case "DISAGREE"
''                            KETHSLKERJA = "LEFT(KETHSLKERJA,1)= 'D'"
''                        Case "SENDING APPLICATION"
''                            KETHSLKERJA = "KETHSLKERJA= 'SA'"
''                            'KETHSLKERJA = "KETHSLKERJA= 'TBO'"
''                        Case "PICKUP"
''                            KETHSLKERJA = "KETHSLKERJA= 'PU'"
''                            'KETHSLKERJA = "KETHSLKERJA= 'TBO'"
''                        Case "INCOMPLETE DOCUMENT"
''                            KETHSLKERJA = "KETHSLKERJA= 'ID'"
''                            'KETHSLKERJA = "KETHSLKERJA= 'TBO'"
''                        Case "INCOMING"
''                            KETHSLKERJA = "KETHSLKERJA= 'I'"
''                        Case "AVAILABLE"
''                            KETHSLKERJA = "KETHSLKERJA= '1A'"
''                    End Select
''                End If
''
''                If CmbStatusCek.Text <> "" Then
''                    Select Case UCase(Trim(CmbStatusCek.Text))
''                        Case "ACCEPT"
''                            lStatusCek = "UPPER(F_cek) ='ACCEPT'"
''                        Case "RETURN"
''                            lStatusCek = "UPPER(F_cek) = 'RETURN'"
''                        Case "NOT CHECK"
''                            lStatusCek = "(f_cek is null or f_cek = 'NOT ACCEPT' OR F_CEK ='') AND KETHSLKERJA ='I'"
''                    End Select
''                End If
''
''                If DtLastCall(0).ValueIsNull Then
''                Else
''                    lLastCallDate = "TGLSTATUS BETWEEN '" + Format(DtLastCall(0).Value, "MM/DD/YYYY") & " " & CStr(DTimeLastCall(0).Value) + "' AND '" + Format(DtLastCall(1).Value, "MM/DD/YYYY") & " " & CStr(DTimeLastCall(1).Value) + "'"
''                End If
''        End If
''        End If
'
'                'Unload FRM_SEARCH
'                If Check1.Value = 0 Then
'                    Set m_cari = M_DATA.QUERY_SEARCH_CONDITION(M_OBJCONN, NAMACUST, NAMAAGENT, DATASOURCE, TGLLAHIR, _
'                                                            OFFPHONE, OFFPHONE2, HOMEPHONE, HOMEPHONE2, MOBILEPHONE, _
'                                                            MOBILEPHONE2, FAXPHONE, FAXPHONE2, MDIForm1.Text3.Text, Lcustid, F_CEK, lLastCallDate, lStatusCek)
'                Else
'                    Set m_cari = M_DATA.QUERY_SEARCH_CONDITION_mgm2(M_OBJCONN, Reff_Num, No_Laporan, Tgl_Laporan, Idi_User, _
'                                                             Debitur_Name, Born_Place, Born_Place2, Born_Date, Din, _
'                                                            NPWP, NPWP2, KTP_NUM, KTP_NUM2, Passport_Num, Passport_Num2, HTML_SOURCE)
'
'
'                End If
'
'            If m_cari.RecordCount = 0 Then
'                MsgBox "Data Tidak Ditemukan", vbInformation + vbOKOnly, "Aplikasi"
'                Set M_DATA = Nothing
'                Exit Sub
'            Else
'
'                search_ok = True
'                If Check1.Value = 1 Then
'                    SSTab1.Tab = 0
'                    Call show_Search_mgmData
'
'                Else
'                    SSTab1.Tab = 1
'                End If
'            End If
'        End If
'    Case 1
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
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'MDIForm1.m_targetview = False
'End Sub
'
'Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
'   ListView1.SortKey = ColumnHeader.Index - 1
'   ListView1.Sorted = True
'End Sub
'
'Private Sub LstVwSearchmgm_ColumnClick(ByVal ColumnHeader As ColumnHeader)
'   LstVwSearchmgm.SortKey = ColumnHeader.Index - 1
'   LstVwSearchmgm.Sorted = True
'End Sub
'
'Private Sub LstVwSearchmgm_DblClick()
'
'
'
'
'
'
'
Private Sub Command1_Click(Index As Integer)

End Sub
