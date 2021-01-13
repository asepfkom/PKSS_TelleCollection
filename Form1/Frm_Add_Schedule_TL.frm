VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frm_add_schedule_tl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Lock Data TL"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8970
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer_stopwatch 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7770
      Top             =   1800
   End
   Begin VB.CheckBox CheckEntry 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tampilkan data entry:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   210
      TabIndex        =   0
      Top             =   5250
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.Frame FrameEntry 
      BackColor       =   &H00E0E0E0&
      Height          =   2025
      Left            =   0
      TabIndex        =   1
      Top             =   5250
      Visible         =   0   'False
      Width           =   8895
      Begin VB.OptionButton OptNewEntry 
         BackColor       =   &H00E0E0E0&
         Caption         =   "New entry "
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
         Left            =   5355
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.ComboBox CmbNewEntry 
         Height          =   315
         ItemData        =   "Frm_Add_Schedule_TL.frx":0000
         Left            =   3660
         List            =   "Frm_Add_Schedule_TL.frx":0007
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   450
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.OptionButton OptReguler 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reguler "
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
         Left            =   5355
         TabIndex        =   9
         Top             =   840
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.ComboBox CmbReguler 
         Height          =   315
         ItemData        =   "Frm_Add_Schedule_TL.frx":0010
         Left            =   3660
         List            =   "Frm_Add_Schedule_TL.frx":0017
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.ComboBox CmbSwap 
         Height          =   315
         ItemData        =   "Frm_Add_Schedule_TL.frx":0020
         Left            =   3660
         List            =   "Frm_Add_Schedule_TL.frx":0027
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1260
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.OptionButton OptSwap 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Swap"
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
         Left            =   5355
         TabIndex        =   6
         Top             =   1260
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox chkcurrent 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Current Entry"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   600
         TabIndex        =   5
         Top             =   1530
         Width           =   1755
      End
      Begin VB.CheckBox chkswap 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Swap"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   600
         TabIndex        =   4
         Top             =   1140
         Width           =   1755
      End
      Begin VB.CheckBox chkreguler 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reguler"
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
         Left            =   600
         TabIndex        =   3
         Top             =   690
         Width           =   1485
      End
      Begin VB.CheckBox chknewentry 
         BackColor       =   &H00E0E0E0&
         Caption         =   "New Entry"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   600
         TabIndex        =   2
         Top             =   300
         Width           =   1725
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bulan"
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
         Left            =   4560
         TabIndex        =   14
         Top             =   510
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bulan"
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
         Left            =   4560
         TabIndex        =   13
         Top             =   900
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bulan"
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
         Left            =   4560
         TabIndex        =   12
         Top             =   1320
         Visible         =   0   'False
         Width           =   585
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   1095
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   1931
      _Version        =   196610
      BackColor       =   14737632
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1800
         TabIndex        =   19
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   3120
         TabIndex        =   18
         Top             =   600
         Width           =   2655
      End
      Begin VB.CheckBox CHKLUNASPENDING 
         BackColor       =   &H00E0E0E0&
         Caption         =   "LUNAS PENDING"
         Height          =   465
         Left            =   5970
         TabIndex        =   17
         Top             =   60
         Width           =   1725
      End
      Begin VB.CheckBox CHKACCOUNT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "LUNAS COMPLETE"
         Height          =   285
         Left            =   5970
         TabIndex        =   16
         Top             =   570
         Width           =   1845
      End
      Begin Threed.SSOption SSOption1 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   196610
         BackColor       =   14737632
         Caption         =   "All TeleCollection"
      End
      Begin Threed.SSOption SSOption1 
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   21
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   196610
         BackColor       =   14737632
         Caption         =   "Pilih TeleCollection"
      End
      Begin Threed.SSOption SSOption1 
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   22
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   196610
         BackColor       =   14737632
         Caption         =   "Pilih SPV Name"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "TeleCollection Name :"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   1575
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   3825
      Left            =   0
      TabIndex        =   24
      Top             =   1260
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   6747
      _Version        =   196610
      BackColor       =   14737632
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "S P - Settled Payment"
         Height          =   255
         Index           =   5
         Left            =   3000
         TabIndex        =   47
         Top             =   810
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "RP - Refuse Payment"
         Height          =   255
         Index           =   7
         Left            =   135
         TabIndex        =   46
         Top             =   480
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "VL-Valid"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   45
         Top             =   120
         Width           =   1245
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "OS - On Process"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   44
         Top             =   840
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "P T P - Promise To Pay"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   43
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "B P - Broken Promise"
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   42
         Top             =   120
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "P O P - Progress Of Payment"
         Height          =   255
         Index           =   4
         Left            =   3000
         TabIndex        =   41
         Top             =   480
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Blank Data"
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   40
         Top             =   1170
         Width           =   2535
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   2325
         Left            =   90
         TabIndex        =   28
         Top             =   1410
         Width           =   8685
         Begin VB.CommandButton cmd 
            Caption         =   "<<"
            Height          =   375
            Index           =   3
            Left            =   3960
            TabIndex        =   35
            Top             =   1860
            Width           =   675
         End
         Begin VB.CommandButton cmd 
            Caption         =   ">>"
            Height          =   375
            Index           =   2
            Left            =   3960
            TabIndex        =   34
            Top             =   1470
            Width           =   675
         End
         Begin VB.CommandButton cmd 
            Caption         =   "<"
            Height          =   375
            Index           =   1
            Left            =   3960
            TabIndex        =   33
            Top             =   1080
            Width           =   675
         End
         Begin VB.CommandButton cmd 
            Caption         =   ">"
            Height          =   375
            Index           =   0
            Left            =   3960
            TabIndex        =   32
            Top             =   690
            Width           =   675
         End
         Begin VB.CheckBox chksingle 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Single"
            Height          =   345
            Left            =   1620
            TabIndex        =   31
            Top             =   240
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.CheckBox chkmultiple 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Multiple"
            Height          =   345
            Left            =   2430
            TabIndex        =   30
            Top             =   270
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Hapus"
            Height          =   315
            Left            =   3360
            TabIndex        =   29
            Top             =   720
            Width           =   585
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1575
            Left            =   90
            TabIndex        =   36
            Top             =   690
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   2778
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            TextBackground  =   -1  'True
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
         Begin MSComctlLib.ListView ListView2 
            Height          =   1575
            Left            =   4710
            TabIndex        =   37
            Top             =   660
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   2778
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            TextBackground  =   -1  'True
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
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Source Mark Up"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   330
            Width           =   1185
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Destination Lead Markup"
            Height          =   255
            Left            =   4800
            TabIndex        =   38
            Top             =   390
            Width           =   2685
         End
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "PR-Prospect"
         Height          =   255
         Index           =   9
         Left            =   5700
         TabIndex        =   27
         Top             =   150
         Width           =   1245
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ON-On Nego"
         Height          =   255
         Index           =   10
         Left            =   5700
         TabIndex        =   26
         Top             =   450
         Width           =   1245
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "SK-Skip"
         Height          =   255
         Index           =   11
         Left            =   5670
         TabIndex        =   25
         Top             =   780
         Width           =   1245
      End
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Index           =   0
      Left            =   45
      TabIndex        =   48
      Top             =   8610
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   196610
      MousePointer    =   16
      Caption         =   "&Add Schedule"
      ButtonStyle     =   3
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Index           =   1
      Left            =   4020
      TabIndex        =   49
      Top             =   8640
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   196610
      MousePointer    =   16
      Caption         =   "&Release"
      ButtonStyle     =   3
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Index           =   2
      Left            =   2040
      TabIndex        =   50
      Top             =   8580
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   196610
      MousePointer    =   16
      Caption         =   "E&xit"
      ButtonStyle     =   3
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Index           =   3
      Left            =   5100
      TabIndex        =   51
      Top             =   8640
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   196610
      MousePointer    =   16
      Caption         =   "&SHUT"
      ButtonStyle     =   3
   End
   Begin TDBDate6Ctl.TDBDate StartDate 
      Height          =   315
      Left            =   1515
      TabIndex        =   52
      Top             =   7425
      Width           =   1560
      _Version        =   65536
      _ExtentX        =   2752
      _ExtentY        =   556
      Calendar        =   "Frm_Add_Schedule_TL.frx":0030
      Caption         =   "Frm_Add_Schedule_TL.frx":0148
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "Frm_Add_Schedule_TL.frx":01B4
      Keys            =   "Frm_Add_Schedule_TL.frx":01D2
      Spin            =   "Frm_Add_Schedule_TL.frx":0230
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
   Begin TDBTime6Ctl.TDBTime StartTime 
      Height          =   315
      Left            =   3090
      TabIndex        =   53
      Top             =   7425
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
      _ExtentY        =   556
      Caption         =   "Frm_Add_Schedule_TL.frx":0258
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "Frm_Add_Schedule_TL.frx":02C4
      Spin            =   "Frm_Add_Schedule_TL.frx":0314
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
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
      Text            =   "__:__"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   1.02960316199441E-317
   End
   Begin TDBDate6Ctl.TDBDate EndDate 
      Height          =   315
      Left            =   1515
      TabIndex        =   54
      Top             =   7845
      Width           =   1560
      _Version        =   65536
      _ExtentX        =   2752
      _ExtentY        =   556
      Calendar        =   "Frm_Add_Schedule_TL.frx":033C
      Caption         =   "Frm_Add_Schedule_TL.frx":0454
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "Frm_Add_Schedule_TL.frx":04C0
      Keys            =   "Frm_Add_Schedule_TL.frx":04DE
      Spin            =   "Frm_Add_Schedule_TL.frx":053C
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
   Begin TDBTime6Ctl.TDBTime EndTime 
      Height          =   315
      Left            =   3090
      TabIndex        =   55
      Top             =   7845
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
      _ExtentY        =   556
      Caption         =   "Frm_Add_Schedule_TL.frx":0564
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "Frm_Add_Schedule_TL.frx":05D0
      Spin            =   "Frm_Add_Schedule_TL.frx":0620
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
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
      Text            =   "__:__"
      ValidateMode    =   0
      ValueVT         =   1
      Value           =   1.02960316199441E-317
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Waktu Server Sekarang:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   4470
      TabIndex        =   59
      Top             =   7410
      Width           =   4380
   End
   Begin VB.Label LblWaktu 
      BackStyle       =   0  'Transparent
      Caption         =   "Label Waktu"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   4470
      TabIndex        =   58
      Top             =   7920
      Width           =   2700
   End
   Begin VB.Label Label7 
      Caption         =   "Start Lock:"
      Height          =   225
      Left            =   150
      TabIndex        =   57
      Top             =   7425
      Width           =   1170
   End
   Begin VB.Label Label8 
      Caption         =   "End Lock:"
      Height          =   225
      Left            =   150
      TabIndex        =   56
      Top             =   7845
      Width           =   1170
   End
End
Attribute VB_Name = "frm_add_schedule_tl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TotalTenthDetik, TotalDetik, TenthDetik, Detik, Menit, JAM As Integer
Dim Jam1 As String

Dim cmdsql As String
Dim StsVl As String
Dim StsOS As String
Dim StsON As String
Dim StsSK As String
Dim StsPR As String
Dim StsPTP As String
Dim StsBP As String
Dim StsPOP As String
Dim StsSP As String
Dim StsRP As String
Dim StsOP As String
Dim StsFresh As String
Dim Stsblank As String
Dim Stsuncontact As String
Dim spv As Boolean
'@@ 140710 Tambahan buat blok entry yang diambil dari field entry_date dan pay_dt di mgm
Dim BlokEntry As String
'@@ 061110 Blok data automatic dengan waktu
Dim StringBlokTimer As String
Dim StatusLocked As String
'@@ 18-11-10 Perbaikan dari blok data entry
Dim StsNewEntry As String
Dim StsReguler As String
Dim StsSwap As String
Dim StsCurrent As String
Dim CekValidLock As Boolean

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0
         KeyAscii = 0
        Case 1
         KeyAscii = 0
    End Select
End Sub

Private Sub Form_Activate()
    'SSOption1(2).Value = True
    'SSOption1(2).Enabled = False
    'Combo1(0).Text = Replace(Trim(mdiform1.txtusername.text), "TL", "SPV")
End Sub

Private Sub header()
    ListView1.ColumnHeaders.ADD 1, , "Batch", 10 * 120
    ListView2.ColumnHeaders.ADD 1, , "Batch", 20 * 120
End Sub
Private Sub CheckEntry_Click()
    If FrameEntry.Enabled = False Then
        FrameEntry.Enabled = True
        OptNewEntry.Value = True
    Else
        FrameEntry.Enabled = False
        OptNewEntry.Value = False
        OptReguler.Value = False
        OptSwap.Value = False
    End If
End Sub


Private Sub chkmultiple_Click()
    If chkmultiple.Value = vbChecked Then
        chksingle.Value = vbUnchecked
    End If
End Sub

Private Sub chksingle_Click()
    If chksingle.Value = vbChecked Then
        chkmultiple.Value = vbUnchecked
    End If
End Sub

Private Sub cmd_Click(Index As Integer)
Select Case Index
Case 1
    If ListView2.ListItems.Count <> 0 Then
            Set lList = ListView1.ListItems.ADD(, , ListView2.SelectedItem.Text)
            ListView2.ListItems.Remove ListView2.SelectedItem.Index
    End If
Case 3
    For i = 1 To ListView2.ListItems.Count
                Set lList = ListView1.ListItems.ADD(, , ListView2.SelectedItem.Text)
                ListView2.ListItems.Remove ListView2.SelectedItem.Index
    Next
Case 0
    If ListView1.ListItems.Count <> 0 Then
        Set lList = ListView2.ListItems.ADD(, , ListView1.SelectedItem.Text)
        ListView1.ListItems.Remove ListView1.SelectedItem.Index
    End If

Case 2
    For i = 1 To ListView1.ListItems.Count
            Set lList = ListView2.ListItems.ADD(, , ListView1.SelectedItem.Text)
                   
            ListView1.ListItems.Remove ListView1.SelectedItem.Index
    Next
End Select
End Sub

Private Sub Combo1_Click(Index As Integer)
Dim M_DATA As New CLS_FRMSEARCH
Dim M_objrs As ADODB.Recordset
Select Case Index
Case 0
    If spv = False Then
        Set M_objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "USERID = '" + Combo1(Index).Text + "'")
        If M_objrs.RecordCount <> 0 Then
            Combo1(0).Text = M_objrs("USERID")
            Combo1(1).Text = M_objrs("AGENT")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    Else
        Set M_objrs = New ADODB.Recordset
        M_objrs.Open "select * from SPVTBL where SPVCODE='" + Combo1(0) + "'", M_OBJCONN, adOpenDynamic, adLockBatchOptimistic
            While Not M_objrs.EOF
                Combo1(0).Text = M_objrs("SPVCODE")
                Combo1(1).Text = M_objrs("SPVNAME")
                M_objrs.MoveNext
            Wend
        Set M_objrs = Nothing
        spv = True
    End If
Case 1
    If spv = False Then
        Set M_objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "AGENT = '" + Combo1(Index).Text + "'")
        If M_objrs.RecordCount <> 0 Then
            Combo1(0).Text = M_objrs("USERID")
            Combo1(1).Text = M_objrs("AGENT")
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    Else
        Set M_objrs = New ADODB.Recordset
        M_objrs.Open "select * from SPVTBL where SPVNAME='" + Combo1(1) + "'", M_OBJCONN, adOpenDynamic, adLockBatchOptimistic
            While Not M_objrs.EOF
                Combo1(0).Text = M_objrs("SPVCODE")
                Combo1(1).Text = M_objrs("SPVNAME")
                M_objrs.MoveNext
            Wend
        Set M_objrs = Nothing
        spv = True
    End If
    
 End Select
 Set M_DATA = Nothing
 Set M_objrs = Nothing
End Sub

Private Sub Command1_Click()
    '@@ 19-11-10 tambahan kasih konfirmasi dan cek dulu datanya
    Dim a As String
    Dim SqlCek As String
    Dim M_objrs As ADODB.Recordset
    
    a = MsgBox("Yakin data akan dihapus?", vbYesNo + vbQuestion, "Konfirmasi")
    
    If a = vbYes Then
        'Cek data dulu nih,apakah masih diblok atau belum
        '@@== 06122010 salah logik diganti :)
'        SqlCek = "select * from mgm where exclude='"
'        SqlCek = SqlCek + Trim(ListView1.SelectedItem.Text) + "'"

        SqlCek = "select lockmarkup from usertbl where lockmarkup like '%"
        SqlCek = SqlCek + Trim(ListView1.SelectedItem.Text) + "%'"
        
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open SqlCek, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_objrs.RecordCount <> 0 Then
            MsgBox "Data tidak dapat dihapus! Karena data masih dalam proses blok data!", vbOKOnly + vbExclamation, "Peringatan"
            Set M_objrs = Nothing
            Exit Sub
        Else
            M_OBJCONN.Execute "UPDATE MGM SET EXCLUDE =NULL WHERE EXCLUDE='" + Trim(ListView1.SelectedItem.Text) + "'"
            getMarkup
            MsgBox "Data berhasil dihapus!", vbOKOnly + vbInformation, "Informasi"
        End If
    End If
    Set M_objrs = Nothing
End Sub

Private Sub Form_Load()
    Dim M_objrs As ADODB.Recordset
    Dim M_DATA As New CLS_FRMSEARCH
    Dim m_waktuserver As ADODB.Recordset
    Dim SqlWaktu As String
    
    
    Set M_objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "")
        While Not M_objrs.EOF
            Combo1(0).AddItem M_objrs("USERID")
            Combo1(1).AddItem M_objrs("AGENT")
            M_objrs.MoveNext
        Wend
    Set M_objrs = Nothing
    
    'Jika yang login tl maka blok data ke semua agent di nonaktifkan
    If Left(Trim(MDIForm1.txtusername.Text), 2) = "TL" Then
        SSOption1(0).Enabled = False
        SSOption1(1).Enabled = True
        SSOption1(1).Value = True
    Else
        SSOption1(0).Value = True
    End If
        
        
    spv = False
    header
    CmbNewEntry.Text = "< 2"
    CmbReguler.Text = "< 2"
    CmbSwap.Text = "> 2"
    getMarkup
    
    'Ambil waktu server
    SqlWaktu = "select now()"
    Set m_waktuserver = New ADODB.Recordset
    m_waktuserver.CursorLocation = adUseClient
    m_waktuserver.Open SqlWaktu, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        JAM = Val(Format(m_waktuserver(0), "hh"))
        Menit = Val(Format(m_waktuserver(0), "nn"))
        Detik = Val(Format(m_waktuserver(0), "ss"))
    Set m_waktuserver = Nothing
    LblWaktu.Caption = JAM & ":" & Menit
   ' Timer_stopwatch.Enabled = True
End Sub


Private Sub SSCommand1_Click(Index As Integer)
Dim M_objrs As New ADODB.Recordset
Dim sStrsql As String
Dim MWHERE As String
Select Case Index
Case 0
    '@@ 061110 - awal blok dengan timer -
        StringBlokTimer = "AWAL | "
    '@@ 061110 - akhir blok dengan timer -
        StatusLocked = ""

    
If CHKLUNASPENDING.Value = vbChecked And CHKACCOUNT.Value = vbChecked Then
    If Combo1(1).Text <> Empty And SSOption1(2).Value = True Then
        sStrsql = " agent in (@LUNAS PENDING@,@LUNAS COMPLETE@) AND AGENTPREV IN (SELECT USERID FROM USERTBL WHERE SPVCODE LIKE @%SPV%@ ) "
       'mwhere = " WHERE SPVCODE LIKE '%SPV%'"
        sStrsql = sStrsql + " AND AGENTPREV IN (SELECT USERID FROM USERTBL WHERE SPVCODE=@" + Combo1(0).Text + "@)"
        MWHERE = "Where spvcode='" + Combo1(0).Text + "'"
        STRSQL = "UPDATE usertbl SET dilockoleh='"
        STRSQL = STRSQL + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
        STRSQL = STRSQL + "lockdarispvbuattl ='" + sStrsql + "'" + MWHERE
        'M_OBJCONN.Execute STRSQL
        
        '@@ 061110 - awal blok dengan timer -
          StringBlokTimer = StringBlokTimer + STRSQL + " | "
          StatusLocked = StatusLocked + " LUNAS PENDING- LUNAS COMPLETE"
        '@@ 061110 - akhir blok dengan timer -

        
    ElseIf Combo1(1).Text = Empty And SSOption1(2).Value = True Then
            Set M_objrs = New ADODB.Recordset
            M_objrs.CursorLocation = adUseClient
            M_objrs.Open "SELECT * FROM USERTBL WHERE USERTYPE='6'", M_OBJCONN, adOpenDynamic, adLockOptimistic
            While Not M_objrs.EOF
                sStrsql = " agent in (@LUNAS PENDING@,@LUNAS COMPLETE@) "
                sStrsql = sStrsql + " AND AGENTPREV IN (SELECT USERID FROM USERTBL WHERE SPVCODE=@" + M_objrs("SPVCODE") + "@)"
                MWHERE = "Where spvcode='" + M_objrs("SPVCODE") + "'"
                STRSQL = "UPDATE usertbl SET dilockoleh='"
                STRSQL = STRSQL + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                STRSQL = STRSQL + "lockdarispvbuattl ='" + sStrsql + "'" + MWHERE
                'M_OBJCONN.Execute STRSQL
                
                '@@ 061110 - awal blok dengan timer -
                StringBlokTimer = StringBlokTimer + STRSQL + " | "
                '@@ 061110 - akhir blok dengan timer -
                
                M_objrs.MoveNext
            Wend
            '@@ 061110
            StatusLocked = StatusLocked + " LUNAS PENDING- LUNAS COMPLETE"
            '@@ 061110
            Set M_objrs = Nothing
    Else
    Exit Sub
    End If
    
    'MsgBox "Data Berhasil di Blok", vbOKOnly + vbInformation, "Pesan"
    Exit Sub
  ElseIf CHKLUNASPENDING.Value = vbChecked Then
        If Combo1(1).Text <> Empty And SSOption1(2).Value = True Then
          sStrsql = " agent in (@LUNAS PENDING@) AND AGENTPREV IN (SELECT USERID FROM USERTBL WHERE SPVCODE LIKE @%SPV%@ ) "
            MWHERE = " WHERE SPVCODE LIKE '%SPV%'"
            sStrsql = sStrsql + " AND AGENTPREV IN (SELECT USERID FROM USERTBL WHERE SPVCODE=@" + Combo1(0).Text + "@)"
            MWHERE = "Where spvcode='" + Combo1(0).Text + "'"
             STRSQL = "UPDATE usertbl SET dilockoleh='"
             STRSQL = STRSQL + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
             STRSQL = STRSQL + " lockdarispvbuattl ='" + sStrsql + "'" + MWHERE
            'M_OBJCONN.Execute STRSQL
            
         '@@ 061110 - awal blok dengan timer -
          StringBlokTimer = StringBlokTimer + STRSQL + " | "
         '@@ 061110 - akhir blok dengan timer -
         
         '@@ 061110
            StatusLocked = StatusLocked + " LUNAS PENDING- "
          '@@ 061110
            
        ElseIf Combo1(1).Text = Empty And SSOption1(2).Value = True Then
            Set M_objrs = New ADODB.Recordset
            M_objrs.CursorLocation = adUseClient
            M_objrs.Open "SELECT * FROM USERTBL WHERE USERTYPE='6'", M_OBJCONN, adOpenDynamic, adLockOptimistic
            While Not M_objrs.EOF
                sStrsql = " agent in (@LUNAS PENDING@) "
                sStrsql = sStrsql + " AND AGENTPREV IN (SELECT USERID FROM USERTBL WHERE SPVCODE=@" + M_objrs("SPVCODE") + "@)"
                MWHERE = "Where spvcode='" + M_objrs("SPVCODE") + "'"
                STRSQL = "UPDATE usertbl SET dilockoleh='"
                STRSQL = STRSQL + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                STRSQL = STRSQL + " lockdarispvbuattl ='" + sStrsql + "'" + MWHERE
                'M_OBJCONN.Execute STRSQL
                
                '@@ 061110 - awal blok dengan timer -
                StringBlokTimer = StringBlokTimer + STRSQL + " | "
                '@@ 061110 - akhir blok dengan timer -
                
                M_objrs.MoveNext
            Wend
            '@@ 061110
            StatusLocked = StatusLocked + " LUNAS PENDING-"
            '@@ 061110
            Set M_objrs = Nothing
    Else
        Exit Sub
        
        End If
        
        'MsgBox "Data Berhasil di Blok", vbOKOnly + vbInformation, "Pesan"
        Exit Sub
ElseIf CHKACCOUNT.Value = vbChecked Then
           sStrsql = " agent in (@LUNAS COMPLETE@) AND AGENTPREV IN (SELECT USERID FROM USERTBL WHERE SPVCODE LIKE @%SPV%@ ) "
            MWHERE = " WHERE SPVCODE LIKE '%SPV%'"
        If Combo1(1).Text <> Empty And SSOption1(2).Value = True Then
            sStrsql = sStrsql + " AND AGENTPREV IN (SELECT USERID FROM USERTBL WHERE SPVCODE=@" + Combo1(0).Text + "@)"
            MWHERE = "Where spvcode='" + Combo1(0).Text + "'"
            STRSQL = "UPDATE usertbl SET dilockoleh='"
            STRSQL = STRSQL + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
            STRSQL = STRSQL + "lockdarispvbuattl ='" + sStrsql + "'" + MWHERE
            'M_OBJCONN.Execute STRSQL
            
            '@@ 061110 - awal blok dengan timer -
            StringBlokTimer = StringBlokTimer + STRSQL + " | "
            '@@ 061110 - akhir blok dengan timer -
            
            '@@ 061110
            StatusLocked = StatusLocked + " LUNAS COMPLETE -"
            '@@ 061110
            
        ElseIf Combo1(1).Text = Empty And SSOption1(2).Value = True Then
            Set M_objrs = New ADODB.Recordset
            M_objrs.CursorLocation = adUseClient
            M_objrs.Open "SELECT * FROM USERTBL WHERE USERTYPE='6'", M_OBJCONN, adOpenDynamic, adLockOptimistic
            While Not M_objrs.EOF
                sStrsql = " agent in (@LUNAS COMPLETE@) "
                sStrsql = sStrsql + " AND AGENTPREV IN (SELECT USERID FROM USERTBL WHERE SPVCODE=@" + M_objrs("SPVCODE") + "@)"
                MWHERE = "Where spvcode='" + M_objrs("SPVCODE") + "'"
                STRSQL = "UPDATE usertbl SET dilockoleh='"
                STRSQL = STRSQL + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                STRSQL = STRSQL + "lockdarispvbuattl ='" + sStrsql + "'" + MWHERE
                'M_OBJCONN.Execute STRSQL
                
                '@@ 061110 - awal blok dengan timer -
                StringBlokTimer = StringBlokTimer + STRSQL + " | "
                '@@ 061110 - akhir blok dengan timer -
                
                M_objrs.MoveNext
            Wend
            Set M_objrs = Nothing
            
            '@@ 061110
            StatusLocked = StatusLocked + " LUNAS COMPLETE -"
            '@@ 061110
            
    Else
        Exit Sub
        
        End If
        
            'MsgBox "Data Berhasil di Blok", vbOKOnly + vbInformation, "Pesan"
        Exit Sub
End If


        
        If SSOption1(0).Value = False And SSOption1(1).Value = False And SSOption1(2).Value = False Then
            MsgBox "Select DCR Name To Proccess OR All"
         Else
                If SSOption1(0).Value Then
                    Call ceksts
                    STRSQL = "UPDATE usertbl SET F_NA=NULL,F_PR=NULL,F_VL=NULL,F_ON=NULL,F_OS=NULL,F_SK=NULL, F_OP=NULL, F_PTP=NULL, F_BP=NULL, F_POP=NULL, F_SP=NULL, F_UC=NULL, F_RP=NULL "
                    STRSQL = STRSQL + ", F_WO_DATE=NULL, F_WO_2009=NULL, F_WO_2008=NULL, F_WO_2007=NULL, F_WO_2006=NULL, F_WO_2005=NULL "
                    STRSQL = STRSQL + ", F_WO_2004=NULL, F_WO_2003=NULL, F_WO_2002=NULL, F_WO_2001=NULL, F_WO_2000=NULL, F_WO_1999=NULL,lock_entry_lpd=NULL, lockmarkup=NULL,lockdarispv=NULL Where usertype='1'"
                    M_OBJCONN.Execute (STRSQL)
                    
                    
                    STRSQL = "UPDATE usertbl SET f_flagrender=1, lockdarispv ='"
                    STRSQL = STRSQL + getblock + "',lock_entry_lpd='"
                    STRSQL = STRSQL + GetBlockEntry + "',dilockoleh='"  '@@ 18-11-10 awalnya BlokEntry
                    STRSQL = STRSQL + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "' "
                    STRSQL = STRSQL + " Where usertype='1'"
                    'M_OBJCONN.Execute (STRSQL)
                    
                    '@@ 061110 - awal blok dengan timer -
                    StringBlokTimer = StringBlokTimer + STRSQL + " | "
                    '@@ 061110 - akhir blok dengan timer -
                    
                    
                        If ListView2.ListItems.Count <> 0 Then
                            HLSMARKUP = Replace(GETSELECTMARKUP, "'", "@")
                            sStrsql = " UPDATE USERTBL SET dilockoleh='"
                            sStrsql = sStrsql + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                            sStrsql = sStrsql + "lockmarkup='" + HLSMARKUP + "' FROM ( "
                            sStrsql = sStrsql + " select distinct(agent) AS AGENT FROM MGM where exclude in(" + GETSELECTMARKUP + ")) AS C "
                            sStrsql = sStrsql + " Where usertbl.userid = C.agent and usertbl.usertype='1' "
                            'M_OBJCONN.Execute (sStrsql)
                            
                            '@@ 061110 - awal blok dengan timer -
                            StringBlokTimer = StringBlokTimer + sStrsql + " | "
                            '@@ 061110 - akhir blok dengan timer -
                            '@@ 061110
                            StatusLocked = StatusLocked + HLSMARKUP + "-"
                            '@@ 061110
                        End If
                    'MsgBox "Proccess to All DCR Name Done.....!"
                  End If
        
                If SSOption1(1).Value Then
                    If Combo1(0).Text = "" Then
                        MsgBox "Select DCR Name To Proccess..!"
                        Combo1(0).SetFocus
                    Else
                        Call ceksts
                        STRSQL = "UPDATE usertbl SET F_NA=NULL,F_PR=NULL,F_VL=NULL,F_ON=NULL,F_OS=NULL,F_SK=NULL, F_OP=NULL, F_PTP=NULL, F_BP=NULL, F_POP=NULL, F_SP=NULL, F_UC=NULL, F_RP=NULL "
                        STRSQL = STRSQL + ", F_WO_DATE=NULL, F_WO_2009=NULL, F_WO_2008=NULL, F_WO_2007=NULL, F_WO_2006=NULL, F_WO_2005=NULL "
                        STRSQL = STRSQL + ", F_WO_2004=NULL, F_WO_2003=NULL, F_WO_2002=NULL, F_WO_2001=NULL, F_WO_2000=NULL, F_WO_1999=NULL,lock_entry_lpd=NULL, lockmarkup=NULL,lockdarispv=NULL Where userid='" + Combo1(0).Text + "'"
                        M_OBJCONN.Execute (STRSQL)
                        
                        STRSQL = "UPDATE usertbl SET  f_flagrender=1,lockdarispv ='" + getblock + "',lock_entry_lpd='"
                        STRSQL = STRSQL + GetBlockEntry + "',dilockoleh='"  '@@ 18-11-10 awalnya BlokEntry
                        STRSQL = STRSQL + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "' "
                        STRSQL = STRSQL + "Where userid='" + Trim(Combo1(0).Text) + "'"
                        'M_OBJCONN.Execute STRSQL
                        
                        '@@ 061110 - awal blok dengan timer -
                        StringBlokTimer = StringBlokTimer + STRSQL + " | "
                        '@@ 061110 - akhir blok dengan timer -
                        
                        If ListView2.ListItems.Count <> 0 Then
                            HLSMARKUP = Replace(GETSELECTMARKUP, "'", "@")
                            sStrsql = " UPDATE USERTBL SET dilockoleh='"
                            sStrsql = sStrsql + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                            sStrsql = sStrsql + "lockmarkup='" + HLSMARKUP + "' FROM ( "
                            sStrsql = sStrsql + " select distinct(agent) AS AGENT FROM MGM where exclude in(" + GETSELECTMARKUP + ") AND AGENT='" + Combo1(0).Text + "') AS C "
                            sStrsql = sStrsql + " Where usertbl.userid = C.agent and usertbl.usertype='1'"
                            'M_OBJCONN.Execute (sStrsql)
                            
                            '@@ 061110 - awal blok dengan timer -
                            StringBlokTimer = StringBlokTimer + sStrsql + " | "
                            '@@ 061110 - akhir blok dengan timer -
                            
                            '@@ 061110
                            StatusLocked = StatusLocked + HLSMARKUP + "-"
                            '@@ 061110
                            
                        End If
                        
                    
                        'MsgBox "Proccess To  " + Combo1(0).Text + "  " + Combo1(1).Text + " Done.....!"
                    End If
                Else
                    If SSOption1(2).Value = True Then
                        If Combo1(0).Text = "" Then
                            MsgBox "Select SPV Name To Proccess..!"
                        Else
                            Call ceksts
                        STRSQL = "UPDATE usertbl SET F_NA=NULL,F_PR=NULL,F_VL=NULL,F_ON=NULL,F_OS=NULL,F_SK=NULL,F_OP=NULL, F_PTP=NULL, F_BP=NULL, F_POP=NULL, F_SP=NULL, F_UC=NULL, F_RP=NULL "
                        STRSQL = STRSQL + ", F_WO_DATE=NULL, F_WO_2009=NULL, F_WO_2008=NULL, F_WO_2007=NULL, F_WO_2006=NULL, F_WO_2005=NULL "
                        STRSQL = STRSQL + ", F_WO_2004=NULL, F_WO_2003=NULL, F_WO_2002=NULL, F_WO_2001=NULL, F_WO_2000=NULL, F_WO_1999=NULL,lockmarkup=NULL,lockdarispv=null,lock_entry_lpd=null Where spvcode='" + Combo1(0).Text + "'"
                        M_OBJCONN.Execute (STRSQL)
                            
                       'If CHKLUNASPENDING.Value = vbChecked Then
                        '    STRSQL = "UPDATE usertbl SET f_flagrender=1,lockdarispvbuattl ='" + getblock + "',lock_entry_lpd='"
                         '   STRSQL = STRSQL + BlokEntry + "',fromaccount ='" + cboaccount.Text + "' Where spvcode='" + Combo1(0).Text + "'"
                          '  M_OBJCONN.Execute STRSQL
                       'End If
                       
                       'If CHKACCOUNT.Value = vbChecked Then
                            STRSQL = "UPDATE usertbl SET f_flagrender=1,lockdarispv ='" + getblock + "',lock_entry_lpd='"
                            STRSQL = STRSQL + GetBlockEntry + "',dilockoleh='"  '@@ 18-11-10 awalnya BlokEntry
                            STRSQL = STRSQL + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "' "
                            STRSQL = STRSQL + "Where spvcode='" + Trim(Combo1(0).Text) + "' and usertype='1'"
                            'M_OBJCONN.Execute STRSQL
                       ' End If
                       
                            '@@ 061110 - awal blok dengan timer -
                            StringBlokTimer = StringBlokTimer + STRSQL + " | "
                            '@@ 061110 - akhir blok dengan timer -
                        
                            
                        If ListView2.ListItems.Count <> 0 Then
                            HLSMARKUP = Replace(GETSELECTMARKUP, "'", "@")
                            sStrsql = " UPDATE USERTBL SET dilockoleh='"
                            sStrsql = sStrsql + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                            sStrsql = sStrsql + "lockmarkup='" + HLSMARKUP + "' FROM ( "
                            sStrsql = sStrsql + " select distinct(agent) AS AGENT FROM MGM where exclude in(" + GETSELECTMARKUP + ") AND AGENT IN (SELECT USERID FROM USERTBL WHERE SPVCODE='" + Combo1(0).Text + "')) AS C "
                            sStrsql = sStrsql + " Where usertbl.userid = C.agent and usertbl.usertype='1' "
                            'M_OBJCONN.Execute (sStrsql)
                            
                            '@@ 061110 - awal blok dengan timer -
                            StringBlokTimer = StringBlokTimer + sStrsql + " | "
                            '@@ 061110 - akhir blok dengan timer -
                            
                            '@@ 061110
                            StatusLocked = StatusLocked + HLSMARKUP + "-"
                            '@@ 061110
                        End If
                        
                            
                            'MsgBox "Proccess To  " + Combo1(0).Text + "  " + Combo1(1).Text + " Done.....!"
                        End If
                    End If
             End If
        End If
        
        'Cek validitas status yang di lock
        Call CekValidNullLock
        If CekValidLock = False Then
            Exit Sub
        End If
        
        
        StsVl = ""
        StsPR = ""
        StsOS = ""
        StsON = ""
        StsSK = ""
        StsOP = ""
       StsPTP = ""
       StsBP = ""
       StsPOP = ""
       StsSP = ""
       StsUC = ""
       StsRP = ""
       StsWO_Date = ""
       StsWO_2009 = ""
       StsWO_2008 = ""
       StsWO_2007 = ""
       StsWO_2006 = ""
       StsWO_2005 = ""
       StsWO_2004 = ""
       StsWO_2003 = ""
       StsWO_2002 = ""
       StsWO_2001 = ""
       StsWO_2000 = ""
       StsWO_1999 = ""
       STRSQL = ""
       '@@ 18-11-10 perbaikan blok data entry
       StsNewEntry = ""
       StsReguler = ""
       StsCurrent = ""
       StsSwap = ""
       
            '@@ 061110 - awal blok dengan timer -
           
            
            StringBlokTimer = Replace(StringBlokTimer + " AKHIR ", "'", "$")
            
            Dim CmdTimer As String
            Dim AccLock As String
            
            If SSOption1(0).Value = True Then
                AccLock = "ALL"
            Else
               AccLock = Trim(Combo1(0).Text)
            End If
            
            Dim WaktuAwal As Date
            Dim WaktuAkhir As Date
            Dim WaktuServer As Date
            Dim m_ObjrsWktServer As ADODB.Recordset
            
            Set m_ObjrsWktServer = New ADODB.Recordset
            m_ObjrsWktServer.CursorLocation = adUseClient
            m_ObjrsWktServer.Open "select now() as waktu ", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            WaktuServer = Format(m_ObjrsWktServer(0), "mm-dd-yyyy hh:mm")
            Set m_ObjrsWktServer = Nothing
            
            
            WaktuAwal = Format(StartDate.Value, "mm-dd-yyyy") & " " & StartTime.Value
            WaktuAkhir = Format(EndDate.Value, "mm-dd-yyyy") & " " & EndTime.Value
            
            '@@ 19-11-10 buat validasi pembuatan schedule
           
            
            'cek dulu tanggalnya udah diisi apa belum??
            If IsNull(StartDate.Value) Or IsNull(EndDate.Value) Or IsNull(StartTime.Value) Or IsNull(EndTime.Value) Then
                MsgBox "Start date dan End date, tidak boleh kosong!", vbOKOnly + vbExclamation, "Peringatan"
                StatusLocked = ""
                Exit Sub
            End If
            
            'Start Lock tidak boleh lebih kecil dari waktu server
            If WaktuAwal < WaktuServer Then
                MsgBox "Start lock tidak boleh lebih kecil dari waktu server!", vbOKOnly + vbExclamation, "Peringatan"
                Exit Sub
            End If
            
            'Waktu Awal tidak boleh sama dengan waktu akhir
            If WaktuAwal = WaktuAkhir Then
                MsgBox "Start Lock tidak boleh sama dengan End Lock!", vbOKOnly + vbExclamation, "Peringatan"
                Exit Sub
            End If
            
            'Cek waktu awal tidak boleh lebih besar dari wktu akhir
            If WaktuAwal > WaktuAkhir Then
                MsgBox "Waktu awal tidak boleh lebih besar dari waktu akhir!", vbOKOnly + vbExclamation
                StatusLocked = ""
                Exit Sub
            End If
            
            'Cek apakah dalam tanggal waktu yang sama, ada lock account yang sudah dibuat dlm antrian
            Dim CekSch As String
            Dim M_Objrs_CekSch As ADODB.Recordset
           
           'cek apakah dia sebagai teamleader or spv or admin
'           If UCase(Trim(mdiform1.txtlevel.text)) = "TEAMLEADER" Then
'                 'Ini jika yang login TL
'                 CekSch = "select * from tbltemplockacc where date(start_lock)='"
'                 CekSch = CekSch + Format(WaktuAwal, "yyyy-mm-dd") + "' and f_locked isnull and (lock_by='"
'                 CekSch = CekSch + Trim(mdiform1.txtusername.text) + "' or lock_by in ('SEPTIAN','JOKO','WULAN','ADMINISTRATOR','ADMIN','SUPERVISOR'))"
'                 CekSch = CekSch + " and account_lock in ('" + Trim(AccLock) + "','ALL')"
'           Else
'                'Ini jika yang login SPV or Admin
'                If Trim(AccLock) = "ALL" Then
'                    CekSch = "select * from tbltemplockacc where date(start_lock)='"
'                    CekSch = CekSch + Format(WaktuAwal, "yyyy-mm-dd") + "' and f_locked isnull"
'                Else
'                    CekSch = "select * from tbltemplockacc where date(start_lock)='"
'                    CekSch = CekSch + Format(WaktuAwal, "yyyy-mm-dd") + "' and account_lock='"
'                    CekSch = CekSch + Trim(AccLock) + "' and f_locked isnull"
'                End If
'            End If
'
            
            '@@diubah logikanya 14-12-2010
''
            'Ini jika akan me-lock data semua
            If (Trim(AccLock) = "ALL") Then
                CekSch = "select * from tbltemplockacc where f_locked isnull"
            End If
            'Ini jika akan me-lock data spv
            If Left(Trim(AccLock), 3) = "SPV" Then
                CekSch = "select * from tbltemplockacc where f_locked isnull and (account_lock='ALL' or "
                CekSch = CekSch + "account_lock in (select userid from usertbl where spvcode='"
                CekSch = CekSch + Trim(AccLock) + "') or account_lock='"
                CekSch = CekSch + Trim(AccLock) + "')"
            End If
            If Trim(AccLock) <> "ALL" And Left(Trim(AccLock), 3) <> "SPV" Then
                'Ini jika akan me-lock data per agent
                CekSch = "select * from tbltemplockacc where f_locked isnull and (account_lock='ALL' or "
                CekSch = CekSch + "account_lock in (select spvcode from usertbl where userid='"
                CekSch = CekSch + Trim(AccLock) + "') or account_lock='"
                CekSch = CekSch + Trim(AccLock) + "')"
            End If
            
    
            
            Set M_Objrs_CekSch = New ADODB.Recordset
            M_Objrs_CekSch.CursorLocation = adUseClient
            M_Objrs_CekSch.Open CekSch, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If M_Objrs_CekSch.RecordCount <> 0 Then
                While Not M_Objrs_CekSch.EOF
                    
'                    If WaktuAwal <= CDate(Format(M_Objrs_CekSch("end_lock"), "mm-dd-yyyy hh:mm")) _
'                       Or WaktuAwal = CDate(Format(M_Objrs_CekSch("start_lock"), "mm-dd-yyyy hh:mm")) _
'                    Then
                        
                     'Ubah @@ Logika 15-12-2010
                     If (Format(WaktuAwal, "mm-dd-yyyy hh:mm") >= CDate(Format(M_Objrs_CekSch("start_lock"), "mm-dd-yyyy hh:mm")) _
                        And Format(WaktuAwal, "mm-dd-yyyy hh:mm") <= CDate(Format(M_Objrs_CekSch("end_lock"), "mm-dd-yyyy hh:mm"))) _
                        Or _
                        (Format(WaktuAwal, "mm-dd-yyyy hh:mm") <= CDate(Format(M_Objrs_CekSch("start_lock"), "mm-dd-yyyy hh:mm")) _
                        And Format(WaktuAkhir, "mm-dd-yyyy hh:mm") >= CDate(Format(M_Objrs_CekSch("start_lock"), "mm-dd-yyyy hh:mm"))) _
                      Then
                        'cek dulu deh, data lock dalam antrian di lock oleh siapa?
                        'jika di lock sama spv, tl ga bisa release
                        '@@ 22-11-10 ini di nonaktifkan dulu jadi team leader juga bisa release punya spv, ruquest mba wulan
                        If UCase(Trim(MDIForm1.txtlevel.Text)) = "TEAMLEADER" Then
                            If Trim(M_Objrs_CekSch("lock_by")) = "WULAN" _
                               Or UCase(Trim(M_Objrs_CekSch("lock_by"))) = "SEPTIAN" _
                               Or UCase(Trim(M_Objrs_CekSch("lock_by"))) = "JOKO" _
                               Or UCase(Trim(M_Objrs_CekSch("lock_by"))) = "ADMIN" _
                               Or UCase(Trim(M_Objrs_CekSch("lock_by"))) = "ADMINISTRATOR" _
                               Or UCase(Trim(M_Objrs_CekSch("lock_by"))) = "SUPERVISOR" Then

                                pesan = "Anda tidak dapat membuat lock data terhadap agent ini!" + Chr(13)
                                pesan = pesan + " Karena agent ini telah di lock oleh SPV pada waktu yang bersamaan!" + Chr(13)
                                pesan = pesan + " Silahkan hubungi Supervisor anda!"

                                MsgBox pesan, vbOKOnly + vbExclamation, "Peringatan"
                                Exit Sub
                            End If
                        End If
                        '@@ 22-11-10 ini di nonaktifkan dulu jadi team leader juga bisa release punya spv, ruquest mba wulan

                        
                        'Jika yang lock pak septian maka wulan dan joko diblok
                        If UCase(MDIForm1.txtusername.Text) = "WULAN" Or UCase(MDIForm1.txtusername.Text) = "JOKO" Or UCase(MDIForm1.txtlevel.Text) = "TEAMLEADER" Then
                            If UCase(Trim(M_Objrs_CekSch("lock_by"))) = "SEPTIAN" Then
                                pesan = "Anda tidak dapat membuat lock data terhadap agent ini!" + Chr(13)
                                pesan = pesan + " Karena agent ini telah di lock oleh Pak Septian pada waktu yang bersamaan!" + Chr(13)
                                pesan = pesan + " Silahkan hubungi Pak Septian!"
                                MsgBox pesan, vbOKOnly + vbInformation, "Informasi"
                                Exit Sub
                            End If
                        End If
                        
                        'Di lock dulu nih datanya biar ga bisa runing di agent
                        LockStatus = "Update tbltemplockacc set f_locked='5' where id='"
                        LockStatus = LockStatus + Trim(M_Objrs_CekSch("id")) + "'"
                        M_OBJCONN.Execute LockStatus
                        
                        'Kasih tau pesan ke user
                        pesan = "PERHATIAN! :" + Chr(13)
                        pesan = pesan + "Ada lock data dalam antrian yang sudah dibuat " + Chr(13)
                        pesan = pesan + "dengan AGENT yang di lock = AGENT lock data yang akan dibuat, dan " + Chr(13)
                        pesan = pesan + "WAKTU AWAL lock data yang dibuat lebih kecil dari  " + Chr(13)
                        pesan = pesan + "WAKTU AKHIR lock data dalam antrian yang telah dibuat. " + Chr(13)
                        pesan = pesan + "Hal ini dapat menimpa lock data yang telah dibuat. " + Chr(13)
                        
                        MsgBox pesan, vbOKOnly + vbCritical, "Peringatan"
                        
                        pesan = "Anda dapat menghapus lock data ini:" + Chr(13)
                        pesan = pesan + "Start Lock: " + CStr(M_Objrs_CekSch("start_lock")) + Chr(13)
                        pesan = pesan + "End Lock: " + CStr(M_Objrs_CekSch("end_lock")) + Chr(13)
                        pesan = pesan + "Account yang di lock: " + M_Objrs_CekSch("account_lock") + Chr(13)
                        pesan = pesan + "Di lock oleh: " + M_Objrs_CekSch("lock_by") + Chr(13)
                        pesan = pesan + "Status yang di lock: " + M_Objrs_CekSch("status_lock") + Chr(13)
                        pesan = pesan + Chr(13)
                        pesan = pesan + "Jika anda ingin menyimpan schedule yang sedang dibuat, " + Chr(13)
                        pesan = pesan + "anda harus menghapus data di atas!" + Chr(13)
                        pesan = pesan + "Yakin lock data di atas akan dihapus?"
                        
                        a = MsgBox(pesan, vbYesNo + vbQuestion, "Konfirmasi")
                        If a = vbYes Then
                            Delsch = "delete from tbltemplockacc where id='"
                            Delsch = Delsch + Trim(M_Objrs_CekSch("id")) + "'"
                            M_OBJCONN.Execute Delsch
                            MsgBox "System akan kembali mengecek data! Jika ada data yang sama, system akan kembali mengkonfirmasi anda!", vbOKOnly + vbInformation, "Informasi"
                        Else
                            MsgBox "Lock data gagal ditambahkan!", vbOKOnly + vbExclamation, "Peringatan"
                            'Buka lock data dalam antrian
                            LockStatus = "update tbltemplockacc set f_locked=null where id='"
                            LockStatus = LockStatus + Trim(M_Objrs_CekSch("id")) + "'"
                            M_OBJCONN.Execute LockStatus
                            Set M_Objrs_CekSch = Nothing
                            Exit Sub
                        End If
                        
                    End If
                    M_Objrs_CekSch.MoveNext
                Wend
            End If
            Set M_Objrs_CekSch = Nothing
            
            
            
            'Cek apakah dalam tanggal waktu yang sama, ada lock data yang sedang RUNNING??
            Dim CekSchCurrent As String
            Dim M_Objrs_CekSchCurrent As ADODB.Recordset
            
'            'cek apakah dia sebagai teamleader or spv or admin
'            If UCase(Trim(mdiform1.txtlevel.text)) = "TEAMLEADER" Then
'                 'Ini jika yang login TL
'                 CekSchCurrent = "select * from tbltemplockacc_current where date(start_lock)='"
'                 CekSchCurrent = CekSchCurrent + Format(WaktuAwal, "yyyy-mm-dd") + "' and  (lock_by='"
'                 CekSchCurrent = CekSchCurrent + Trim(mdiform1.txtusername.text) + "' or lock_by in ('SEPTIAN','JOKO','WULAN','ADMIN','ADMINISTRATOR','SUPERVISOR')) "
'                 CekSchCurrent = CekSchCurrent + " and account_lock in ('" + Trim(AccLock) + "','ALL')"
'            Else
'                'Ini jika yang login SPV atau Admin
'                If Trim(AccLock) = "ALL" Then
'                    CekSchCurrent = "select * from tbltemplockacc_current where date(start_lock)='"
'                    CekSchCurrent = CekSchCurrent + Format(WaktuAwal, "yyyy-mm-dd") + "' "
'                Else
'                    CekSchCurrent = "select * from tbltemplockacc_current where date(start_lock)='"
'                    CekSchCurrent = CekSchCurrent + Format(WaktuAwal, "yyyy-mm-dd") + "' and account_lock='"
'                    CekSchCurrent = CekSchCurrent + Trim(AccLock) + "'"
'                End If
'            End If
                
''            '@@diubah logikanya 14-12-2010
''            CekSchCurrent = "select * from tbltemplockacc_current where f_locked isnull and account_lock in ('ALL','"
''            CekSchCurrent = CekSchCurrent + Trim(AccLock) + "','"
''            CekSchCurrent = CekSchCurrent + Trim(Replace(mdiform1.txtusername.text, "TL", "SPV")) + "')"
                
            'Ini untuk mengecek semua data
            If (Trim(AccLock) = "ALL") Then
                CekSchCurrent = "select * from tbltemplockacc_current where f_locked isnull"
            End If
            'Ini untuk mengecek data SPV
            If Left(Trim(AccLock), 3) = "SPV" Then
                CekSchCurrent = "select * from tbltemplockacc_current where f_locked isnull and (account_lock='ALL' or "
                CekSchCurrent = CekSchCurrent + "account_lock in (select userid from usertbl where spvcode='"
                CekSchCurrent = CekSchCurrent + Trim(AccLock) + "') or account_lock='"
                CekSchCurrent = CekSchCurrent + Trim(AccLock) + "')"
            End If
            If Trim(AccLock) <> "ALL" And Left(Trim(AccLock), 3) <> "SPV" Then
                'Ini untuk mengecek data per agent
                CekSchCurrent = "select * from tbltemplockacc_current where f_locked isnull and (account_lock='ALL' or "
                CekSchCurrent = CekSchCurrent + "account_lock in (select spvcode from usertbl where userid='"
                CekSchCurrent = CekSchCurrent + Trim(AccLock) + "') or account_lock='"
                CekSchCurrent = CekSchCurrent + Trim(AccLock) + "')"
            End If
            
            
            Set M_Objrs_CekSchCurrent = New ADODB.Recordset
            M_Objrs_CekSchCurrent.CursorLocation = adUseClient
            M_Objrs_CekSchCurrent.Open CekSchCurrent, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If M_Objrs_CekSchCurrent.RecordCount <> 0 Then
                While Not M_Objrs_CekSchCurrent.EOF
                    
'                    If WaktuAwal <= CDate(Format(M_Objrs_CekSchCurrent("end_lock"), "mm-dd-yyyy hh:mm")) _
'                       Or WaktuAwal = CDate(Format(M_Objrs_CekSchCurrent("start_lock"), "mm-dd-yyyy hh:mm")) _
'                    Then
                        
                     'Ubah @@ Logika 15-12-2010
                     If (Format(WaktuAwal, "mm-dd-yyyy hh:mm") >= CDate(Format(M_Objrs_CekSch("start_lock"), "mm-dd-yyyy hh:mm")) _
                        And Format(WaktuAwal, "mm-dd-yyyy hh:mm") <= CDate(Format(M_Objrs_CekSch("end_lock"), "mm-dd-yyyy hh:mm"))) _
                        Or _
                        (Format(WaktuAwal, "mm-dd-yyyy hh:mm") <= CDate(Format(M_Objrs_CekSch("start_lock"), "mm-dd-yyyy hh:mm")) _
                        And Format(WaktuAkhir, "mm-dd-yyyy hh:mm") >= CDate(Format(M_Objrs_CekSch("start_lock"), "mm-dd-yyyy hh:mm"))) _
                      Then
                        
                        'cek dulu deh, data lock dalam antrian di lock oleh siapa?
                        'jika di lock sama spv, tl ga bisa release
                        '@@22-11-10 INI DINONAKTIFKAN TERLEBIH DAHULU, REQUEST MBA WULAN
                        If UCase(Trim(MDIForm1.txtlevel.Text)) = "TEAMLEADER" Then
                            If UCase(Trim(M_Objrs_CekSchCurrent("lock_by"))) = "WULAN" _
                               Or UCase(Trim(M_Objrs_CekSchCurrent("lock_by"))) = "SEPTIAN" _
                               Or UCase(Trim(M_Objrs_CekSchCurrent("lock_by"))) = "JOKO" _
                               Or UCase(Trim(M_Objrs_CekSchCurrent("lock_by"))) = "ADMIN" _
                               Or UCase(Trim(M_Objrs_CekSchCurrent("lock_by"))) = "ADMINISTRATOR" _
                               Or UCase(Trim(M_Objrs_CekSchCurrent("lock_by"))) = "SUPERVISOR" Then

                                pesan = "Anda tidak dapat membuat lock data terhadap agent ini!" + Chr(13)
                                pesan = pesan + " Karena waktu dari lock data dan agent yang di lock = waktu dari lock yang sedang berjalan!" + Chr(13)
                                pesan = pesan + " Silahkan hubungi Supervisor anda!"

                                MsgBox pesan, vbOKOnly + vbExclamation, "Peringatan"
                                Exit Sub
                            End If
                        End If
                        '@@22-11-10 INI DINONAKTIFKAN TERLEBIH DAHULU, REQUEST MBA WULAN
                        
                        
                        'Jika yang lock pak septian maka wulan dan joko diblok
                        If UCase(Trim(MDIForm1.txtusername.Text)) = "WULAN" Or UCase(Trim(MDIForm1.txtusername.Text)) = "JOKO" Or UCase(Trim(MDIForm1.txtlevel.Text)) = "TEAMLEADER" Then
                            If UCase(Trim(M_Objrs_CekSchCurrent("lock_by"))) = "SEPTIAN" Then
                                pesan = "Anda tidak dapat membuat lock data terhadap agent ini!" + Chr(13)
                                pesan = pesan + " Karena agent ini telah di lock oleh Pak Septian pada waktu yang bersamaan!" + Chr(13)
                                pesan = pesan + " Silahkan hubungi Pak Septian!"
                                MsgBox pesan, vbInformation + vbOKOnly, "Informasi"
                                Exit Sub
                            End If
                        End If
                        
                        
                        'Di lock dulu nih datanya biar ga bisa runing di agent
                        'LockStatus = "Update tbltemplockacc_current set f_locked='6' where id='"
                        'LockStatus = LockStatus + Trim(M_Objrs_CekSchCurrent("id")) + "'"
                        'M_OBJCONN.Execute LockStatus
                        
                        'Kasih tau pesan ke user
                        pesan = "PERHATIAN! :" + Chr(13)
                        pesan = pesan + "Ada lock data yang sedang RUNNING " + Chr(13)
                        pesan = pesan + "dengan AGENT yang di lock = AGENT lock data yang akan dibuat, dan " + Chr(13)
                        pesan = pesan + "WAKTU AWAL lock data yang dibuat lebih kecil dari  " + Chr(13)
                        pesan = pesan + "WAKTU AKHIR lock data yang sedang RUNNING. " + Chr(13)
                        pesan = pesan + "Hal ini dapat menimpa lock data yang telah dibuat. Harap RELEASE Terlebih dahulu lock data tersebut!" + Chr(13)
                       
                        
                        MsgBox pesan, vbOKOnly + vbCritical, "Peringatan"
                        
                       
                       
                        
                        pesan = "Anda dapat me-RELEASE lock data ini:" + Chr(13)
                        pesan = pesan + "Start Lock: " + CStr(M_Objrs_CekSchCurrent("start_lock")) + Chr(13)
                        pesan = pesan + "End Lock: " + CStr(M_Objrs_CekSchCurrent("end_lock")) + Chr(13)
                        pesan = pesan + "Account yang di lock: " + M_Objrs_CekSchCurrent("account_lock") + Chr(13)
                        pesan = pesan + "Di lock oleh: " + M_Objrs_CekSchCurrent("lock_by") + Chr(13)
                        pesan = pesan + "Status yang di lock: " + M_Objrs_CekSchCurrent("status_lock") + Chr(13)
                        pesan = pesan + Chr(13)
                        pesan = pesan + "Jika anda ingin menyimpan schedule yang sedang dibuat, " + Chr(13)
                        pesan = pesan + "anda harus me-RELEASE data di atas!" + Chr(13)
                        pesan = pesan + "Yakin lock data di atas akan di RELEASE?"
                        
                        a = MsgBox(pesan, vbYesNo + vbQuestion, "Konfirmasi")
                        If a = vbYes Then
                        
                             'Clear lock data yang sedang berjalan sesuai dengan agent yang di lock
                            cmdsqlserver = "update usertbl set dilockoleh='Clear by:" + Trim(MDIForm1.txtlevel.Text) + "',"
                            cmdsqlserver = cmdsqlserver + "lockdarispv=null,lock_entry_lpd=null,fromaccount=null,"
                            cmdsqlserver = cmdsqlserver + "lockmarkup=null,lockdarispvbuattl=null"
                            'Buat ambil kondisi agent yang sedang di lock
                            If Trim(M_Objrs_CekSchCurrent("account_lock")) = "ALL" Then
                                cmdsqlserver = cmdsqlserver + " where usertype='1' "
                            ElseIf Left(Trim(M_Objrs_CekSchCurrent("account_lock")), 3) = "SPV" Then
                                cmdsqlserver = cmdsqlserver + " where spvcode='"
                                cmdsqlserver = cmdsqlserver + Trim(M_Objrs_CekSchCurrent("account_lock")) + "'"
                            Else
                                cmdsqlserver = cmdsqlserver + " where userid='"
                                cmdsqlserver = cmdsqlserver + Trim(M_Objrs_CekSchCurrent("account_lock")) + "'"
                            End If
                            M_OBJCONN.Execute cmdsqlserver
                        
                        
                            'Update status pesan ke nilai 1,untuk menampilkan pesan ke agent
                            cmdsqlserver = "update usertbl set f_pesanresetauto='1', f_pesanlockauto=null,f_idsessstart=null,f_idsessend=null "
                            'Buat mengupdate pesan kondisi agent yang di lock
                            If Trim(M_Objrs_CekSchCurrent("account_lock")) = "ALL" Then
                                cmdsqlserver = cmdsqlserver + " where usertype='1' "
                            ElseIf Left(Trim(M_Objrs_CekSchCurrent("account_lock")), 3) = "SPV" Then
                                cmdsqlserver = cmdsqlserver + " where spvcode='"
                                cmdsqlserver = cmdsqlserver + Trim(M_Objrs_CekSchCurrent("account_lock")) + "'"
                            Else
                                cmdsqlserver = cmdsqlserver + " where userid='"
                                cmdsqlserver = cmdsqlserver + Trim(M_Objrs_CekSchCurrent("account_lock")) + "'"
                            End If
                            M_OBJCONN.Execute cmdsqlserver
                            
                            'Clossing Session
                            Dim UpdateDtCloseSession As String
                            UpdateDtCloseSession = "update tblperformpersessionlock set f_ceksekrg=a.f_cek_akhir ,"
                            UpdateDtCloseSession = UpdateDtCloseSession + " tgl_f_ceksekrg=a.tgl_akhir,endlock='" + CStr(Format(WaktuServer, "yyyy-mm-dd hh:mm:ss")) + "' from "
                            UpdateDtCloseSession = UpdateDtCloseSession + " (select mgm.custid as custid_mgm,mgm.agent as agent_mgm,"
                            UpdateDtCloseSession = UpdateDtCloseSession + " tblperformpersessionlock.custid as custid_lock,"
                            UpdateDtCloseSession = UpdateDtCloseSession + " tblperformpersessionlock.agent as agent_lock,"
                            UpdateDtCloseSession = UpdateDtCloseSession + " tblperformpersessionlock.idlock as id_lock,"
                            UpdateDtCloseSession = UpdateDtCloseSession + " mgm.f_cek_new as f_cek_akhir, mgm.tglcall as tgl_akhir"
                            UpdateDtCloseSession = UpdateDtCloseSession + " from tblperformpersessionlock inner join mgm "
                            UpdateDtCloseSession = UpdateDtCloseSession + " on mgm.custid=tblperformpersessionlock.custid "
                            UpdateDtCloseSession = UpdateDtCloseSession + " and mgm.agent=tblperformpersessionlock.agent) as a "
                            UpdateDtCloseSession = UpdateDtCloseSession + " where tblperformpersessionlock.custid=a.custid_mgm and tblperformpersessionlock.agent=a.agent_mgm and a.id_lock='"
                            UpdateDtCloseSession = UpdateDtCloseSession + Trim(M_Objrs_CekSchCurrent("id")) + "'"
                            M_OBJCONN.Execute UpdateDtCloseSession
                            'Akhir dari closing session
                            
                            'Pindahkan data lock account current ke tabel data log tbltemplockacc_log
                            cmdsqlserver = "insert into tbltemplockacc_log select * from tbltemplockacc_current where "
                            cmdsqlserver = cmdsqlserver + " id='"
                            cmdsqlserver = cmdsqlserver + Trim(M_Objrs_CekSchCurrent("id")) + "'"
                            M_OBJCONN.Execute cmdsqlserver
                        
                        
                            Delsch = "delete from tbltemplockacc_current where id='"
                            Delsch = Delsch + Trim(M_Objrs_CekSchCurrent("id")) + "'"
                            M_OBJCONN.Execute Delsch
                            MsgBox "System akan kembali mengecek data! Jika ada data yang sama, system akan kembali mengkonfirmasi anda!", vbOKOnly + vbInformation, "Informasi"
                        Else
                            MsgBox "Lock data gagal ditambahkan!", vbOKOnly + vbExclamation, "Peringatan"
                            'Buka lock data dalam antrian
                            'LockStatus = "update tbltemplockacc_current set f_locked=null where id='"
                            'LockStatus = LockStatus + Trim(M_Objrs_CekSchCurrent("id")) + "'"
                            'M_OBJCONN.Execute LockStatus
                            Set M_Objrs_CekSchCurrent = Nothing
                            Exit Sub
                        End If
                        
                    End If
                    M_Objrs_CekSchCurrent.MoveNext
                Wend
            End If
            Set M_Objrs_CekSchCurrent = Nothing
            
            
            CmdTimer = "insert into tbltemplockacc (date_lock,start_lock,end_lock,"
            CmdTimer = CmdTimer + "account_lock,lock_by,status_lock,script_lock) values ('"
            CmdTimer = CmdTimer + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") + "','"
            CmdTimer = CmdTimer + Format(StartDate.Value, "yyyy-mm-dd") + " " + Format(StartTime.Value, "hh:mm:00") + "','"
            CmdTimer = CmdTimer + Format(EndDate.Value, "yyyy-mm-dd") + " " + Format(EndTime.Value, "hh:mm:00") + "','"
            CmdTimer = CmdTimer + Trim(AccLock) + "','"
            CmdTimer = CmdTimer + Trim(MDIForm1.txtusername.Text) + "','"
            CmdTimer = CmdTimer + Trim(StatusLocked) + "','"
            CmdTimer = CmdTimer + Trim(StringBlokTimer) + "')"
            
            Dim m_objrsTimer As ADODB.Recordset
            Set m_objrsTimer = New ADODB.Recordset
            m_objrsTimer.CursorLocation = adUseClient
            m_objrsTimer.Open CmdTimer, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            Set m_objrsTimer = Nothing
            
            MsgBox "Lock data berhasil ditambahkan dalam schedulle!", vbOKOnly + vbInformation, "Informasi"
            Unload Me
            '@@ 061110 - akhir blok dengan timer -
 
Case 1
     If SSOption1(2).Value = True Then
               
            If CHKLUNASPENDING.Value = vbChecked Or CHKACCOUNT.Value = vbChecked Then
                    
                   Set M_objrs = New ADODB.Recordset
            M_objrs.CursorLocation = adUseClient
            
            If Combo1(0).Text = "" Then
                    M_objrs.Open "SELECT * FROM USERTBL WHERE usertype='6'", M_OBJCONN, adOpenDynamic, adLockOptimistic
            Else
                    M_objrs.Open "SELECT * FROM USERTBL WHERE SPVCODE='" + Combo1(0).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
            End If
            
            While Not M_objrs.EOF
                STRSQL = "UPDATE usertbl SET dilockoleh='Clear by:"
                STRSQL = STRSQL + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                STRSQL = STRSQL + " lockdarispvbuattl=NULL WHERE SPVCODE='" + M_objrs("SPVCODE") + "'"
                M_OBJCONN.Execute STRSQL
                M_objrs.MoveNext
            Wend
            Set M_objrs = Nothing
            MsgBox "Data telah direlease"
           Exit Sub
           End If
            
                If Combo1(0).Text = "" Then
                MsgBox "CLIK DULU COMBO SPV", vbInformation + vbOKOnly, "PESAN"
                Exit Sub
               End If
                    STRSQL = "UPDATE usertbl SET dilockoleh='Clear by:"
                    STRSQL = STRSQL + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                    STRSQL = STRSQL + " lockdarispv=NULL,lock_entry_lpd=NULL,fromaccount=NULL,lockmarkup=NULL,lockdarispvbuattl=NULL WHERE SPVCODE='" + Combo1(0).Text + "'"
                
     Else
            If SSOption1(1).Value = True Then
                If Combo1(0).Text = "" Then
                    MsgBox "CLIK DULU COMBO NYA", vbInformation + vbOKOnly, "PESAN"
                Exit Sub
                End If
                STRSQL = "UPDATE usertbl SET dilockoleh='Clear by:"
                STRSQL = STRSQL + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "',"
                STRSQL = STRSQL + " lockdarispv=NULL,lock_entry_lpd=NULL,lockmarkup=NULL ,fromaccount=NULL,lockdarispvbuattl=NULL WHERE userid='" + Combo1(0).Text + "'"
            Else
                    STRSQL = "UPDATE usertbl SET lockdarispv=NULL,lock_entry_lpd=NULL,lockmarkup=NULL,dilockoleh='Clear by:"
                    STRSQL = STRSQL + MDIForm1.txtusername.Text + "-" + Format(Now, "yyyy-mm-dd") + "'"
            End If
            
    End If
    If STRSQL <> "" Then
        M_OBJCONN.Execute STRSQL
        MsgBox "Reset Done.....!"
        End If
        
Case 2
        Unload Me

Case 3
        STRSQL1 = "UPDATE  tblshut SET nshut=1 "
        M_OBJCONN.Execute STRSQL1
End Select

End Sub
Sub ceksts()
If Check1(0).Value Then
    StsVl = "VL-"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsVl
    
End If

If Check1(1).Value Then
    StsOP = "OS-"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsOP
End If

If Check1(2).Value Then
    StsPTP = "PTP"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsPTP + "-"
End If

If Check1(3).Value Then
    StsBP = "BP-"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsBP
End If

If Check1(4).Value Then
    StsPOP = "POP"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsPOP
End If
 
If Check1(5).Value Then
    StsSP = "SP-"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsSP
End If

If Check1(7).Value Then
    StsRP = "RP-"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsRP
End If

If Check1(6).Value Then
    Stsblank = "anto"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + "BlankData -"
End If

If Check1(9).Value Then
    StsPR = "PR-"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsPR
End If

If Check1(10).Value Then
    StsON = "ON-"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsON
End If

If Check1(11).Value Then
    StsSK = "SK-"
    '@@ 061110 Mencatat Status yang di lock u/ lock timer
    StatusLocked = StatusLocked + StsSK
End If

If chknewentry.Value Then
    StatusLocked = StatusLocked + "NewEntry -"
    
    StsNewEntry = " ("
    StsNewEntry = StsNewEntry + "date_part(''month'',entry_date) between (date_part(''month'',now())- 2) "
    StsNewEntry = StsNewEntry + " and date_part(''month'',now())-1 "
    StsNewEntry = StsNewEntry + " and date_part(''year'',entry_date)=date_part(''year'',now()) "
    StsNewEntry = StsNewEntry + " )"

End If

If chkreguler.Value Then
    StatusLocked = StatusLocked + "Reguler -"
    
    StsReguler = " ("
    StsReguler = StsReguler + "date_part(''month'',pay_dt_update) between (date_part(''month'',now())- 2) "
    StsReguler = StsReguler + " and date_part(''month'',now())-1 "
    StsReguler = StsReguler + " and date_part(''year'',pay_dt_update)=date_part(''year'',now())"
    StsReguler = StsReguler + " )"
    
End If


If chkswap.Value Then
    StatusLocked = StatusLocked + "Swap -"
    
    StsSwap = " ("
    StsSwap = StsSwap + "( (date_part(''month'',pay_dt_update) < (date_part(''month'',now())- 2) "
    StsSwap = StsSwap + " and date_part(''year'',pay_dt_update) <= date_part(''year'',now())) "
    StsSwap = StsSwap + " or pay_dt_update isnull ) "
    StsSwap = StsSwap + " and "
    StsSwap = StsSwap + " date_part(''month'',entry_date) < (date_part(''month'',now())-2) "
    StsSwap = StsSwap + " and date_part(''year'',entry_date) <= date_part(''year'',now()) "
    StsSwap = StsSwap + " )"
End If


If chkcurrent.Value Then
    StatusLocked = StatusLocked + "Current -"
    
    StsCurrent = " ("
    StsCurrent = StsCurrent + " date_part(''month'',tglsource)=date_part(''month'',now()) "
    StsCurrent = StsCurrent + " and date_part(''year'',tglsource)=date_part(''year'',now()) "
    StsCurrent = StsCurrent + " )"
End If


'--- @@ 18-11-10 blok dulu deh buat diperbaiki skripnya u/ yg blok entry---------------
'BlokEntry = ""
'bCheckNewentry = False
'bCheckReguler = False
'bCheckSwap = False
'bCheckCurrent = False
'
'If chknewentry.Value = vbChecked Then
'    bCheckNewentry = True
'    '@@ 061110 Mencatat Status yang di lock u/ lock timer
'    StatusLocked = StatusLocked + "NewEntry -"
'End If
'
'
'If chkreguler.Value = vbChecked Then
'   bCheckReguler = True
'   '@@ 061110 Mencatat Status yang di lock u/ lock timer
'    StatusLocked = StatusLocked + "Reguler -"
'End If
'
'If chkswap.Value = vbChecked Then
'   bCheckSwap = True
'   '@@ 061110 Mencatat Status yang di lock u/ lock timer
'    StatusLocked = StatusLocked + "SWAP -"
'End If
'
'
'If chkcurrent.Value = vbChecked Then
'   bCheckCurrent = True
'   '@@ 061110 Mencatat Status yang di lock u/ lock timer
'    StatusLocked = StatusLocked + "Current -"
'End If
'
'
'
'
'
'
'If bCheckSwap = True And bCheckNewentry = True And bCheckReguler = True And bCheckCurrent = True Then
'    BlokEntry = " (date_part(''month'',entry_date) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
'    BlokEntry = BlokEntry + " and date_part(''year'',entry_date)=date_part(''year'',now) or "
'    BlokEntry = BlokEntry + " (date_part(''month'',pay_dt_update) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
'    BlokEntry = BlokEntry + " and date_part(''year'',pay_dt_update)=date_part(''year'',now)) or "
'    BlokEntry = BlokEntry + " (((date_part(''month'',pay_dt_update)< (date_part(''month'',now)- 2 ) "
'    BlokEntry = BlokEntry + " and date_part(''year'',pay_dt_update)<=date_part(''year'',now)) or "
'    BlokEntry = BlokEntry + " pay_dt_update isnull) and "
'    BlokEntry = BlokEntry + " date_part(''month'',entry_date) < (date_part(''month'',now)-2) "
'    BlokEntry = BlokEntry + " and date_part(''year'',entry_date) <= date_part(''year'',now)) or "
'    BlokEntry = BlokEntry + " (date_part(''month'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "mm") + " and date_part(''year'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "yyyy") + ")"
'    BlokEntry = BlokEntry + " )"
'    Exit Sub
'ElseIf bCheckNewentry = True And bCheckReguler = True And bCheckCurrent = True Then
'    BlokEntry = " (date_part(''month'',entry_date) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
'    BlokEntry = BlokEntry + " and date_part(''year'',entry_date)=date_part(''year'',now) or "
'    BlokEntry = BlokEntry + " (date_part(''month'',pay_dt_update) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
'    BlokEntry = BlokEntry + " and date_part(''year'',pay_dt_update)=date_part(''year'',now)) or "
'    BlokEntry = BlokEntry + " (date_part(''month'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "mm") + " and date_part(''year'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "yyyy") + ")"
'    BlokEntry = BlokEntry + " )"
'    Exit Sub
'ElseIf bCheckNewentry = True And bCheckSwap = True And bCheckCurrent = True Then
'    BlokEntry = " (date_part(''month'',entry_date) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
'    BlokEntry = BlokEntry + " and date_part(''year'',entry_date)=date_part(''year'',now) or "
'    BlokEntry = BlokEntry + " (((date_part(''month'',pay_dt_update)< (date_part(''month'',now)- 2 ) "
'    BlokEntry = BlokEntry + "and date_part(''year'',pay_dt_update)<=date_part(''year'',now)) or "
'    BlokEntry = BlokEntry + "pay_dt_update isnull) and "
'    BlokEntry = BlokEntry + "date_part(''month'',entry_date) < (date_part(''month'',now)-2) "
'    BlokEntry = BlokEntry + " and date_part(''year'',entry_date) <= date_part(''year'',now)) or "
'    BlokEntry = BlokEntry + " (date_part(''month'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "mm") + " and date_part(''year'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "yyyy") + ")"
'    BlokEntry = BlokEntry + " )"
'    Exit Sub
'ElseIf bCheckReguler = True And bCheckSwap = True And bCheckCurrent = True Then
'   BlokEntry = " (date_part(''month'',pay_dt_update) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
'   BlokEntry = BlokEntry + "and date_part(''year'',pay_dt_update)=date_part(''year'',now) or "
'   BlokEntry = BlokEntry + " (((date_part(''month'',pay_dt_update)< (date_part(''month'',now)- 2 ) "
'   BlokEntry = BlokEntry + "and date_part(''year'',pay_dt_update)<=date_part(''year'',now)) or "
'   BlokEntry = BlokEntry + "pay_dt_update isnull) and "
'   BlokEntry = BlokEntry + "date_part(''month'',entry_date) < (date_part(''month'',now)-2) "
'   BlokEntry = BlokEntry + " and date_part(''year'',entry_date) <= date_part(''year'',now)) or "
'   BlokEntry = BlokEntry + " (date_part(''month'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "mm") + " and date_part(''year'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "yyyy") + ")"
'   BlokEntry = BlokEntry + " )"
'   Exit Sub
'End If
'
'
'
'
'If bCheckNewentry = True Then
'    BlokEntry = " date_part(''month'',entry_date) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
'    BlokEntry = BlokEntry + "and date_part(''year'',entry_date)=date_part(''year'',now)"
'    Exit Sub
'End If
'
'
'If bCheckReguler = True Then
'    BlokEntry = " date_part(''month'',pay_dt_update) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
'    BlokEntry = BlokEntry + "and date_part(''year'',pay_dt_update)=date_part(''year'',now)"
'    Exit Sub
'End If
'
'
'If bCheckSwap = True Then
'    BlokEntry = " ((date_part(''month'',pay_dt_update)< (date_part(''month'',now)- 2 ) "
'    BlokEntry = BlokEntry + "and date_part(''year'',pay_dt_update)<=date_part(''year'',now)) or "
'    BlokEntry = BlokEntry + "pay_dt_update isnull) and "
'    BlokEntry = BlokEntry + "date_part(''month'',entry_date) < (date_part(''month'',now)-2) "
'    BlokEntry = BlokEntry + " and date_part(''year'',entry_date) <= date_part(''year'',now)"
'    Exit Sub
'End If
'
'If bCheckCurrent = True Then
'   BlokEntry = " (date_part(''month'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "mm") + " and date_part(''year'',tglsource)=" + Format(MDIForm1.TDBDate1.Value, "yyyy") + ")"
'    Exit Sub
'End If
'--- @@ 18-11-10 blok dulu deh buat diperbaiki skripnya u/ yg blok entry---------------


'@@ 140710 Tambahan buat blok entry

'If OptNewEntry.Value = True Then
'    BlokEntry = " datediff(''month'',entry_date,now) "
'    BlokEntry = BlokEntry + CmbNewEntry.Text
'End If
'
'If OptReguler.Value = True Then
'    BlokEntry = " datediff(''month'',pay_dt,now) "
'    BlokEntry = BlokEntry + CmbReguler.Text
'End If
'
'If OptSwap.Value = True Then
'    BlokEntry = " datediff(''month'',pay_dt,now) "
'    BlokEntry = BlokEntry + CmbSwap.Text
'End If

'@@ 150710 Ubah blok entry
'If OptNewEntry.Value = True Then
'    BlokEntry = " date_part(''month'',entry_date) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
'    BlokEntry = BlokEntry + "and date_part(''year'',entry_date)=date_part(''year'',now)"
'End If
'
'If OptReguler.Value = True Then
'    BlokEntry = " date_part(''month'',pay_dt_update) between (date_part(''month'',now)- 2 ) and date_part(''month'',now)-1 "
'    BlokEntry = BlokEntry + "and date_part(''year'',pay_dt_update)=date_part(''year'',now)"
'End If
'
'If OptSwap.Value = True Then
'    BlokEntry = " ((date_part(''month'',pay_dt_update)< (date_part(''month'',now)- 2 ) "
'    BlokEntry = BlokEntry + "and date_part(''year'',pay_dt_update)<=date_part(''year'',now)) or "
'    BlokEntry = BlokEntry + "pay_dt_update isnull) and "
'    BlokEntry = BlokEntry + "date_part(''month'',entry_date) < (date_part(''month'',now)-2) "
'    BlokEntry = BlokEntry + " and date_part(''year'',entry_date) <= date_part(''year'',now)"
'End If

End Sub
Public Function getblock() As String


                    STRINGBLOK = ""
                    
                    If StsVl <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                            STRINGBLOK = " substring(F_cek_new,1,3) in (@" + StsVl + "@"
                        Else
                            STRINGBLOK = STRINGBLOK + ",@" + StsVl + "@"
                        End If
                    End If
                    
                    If StsPR <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsPR + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsPR + "@"
                        End If
                    End If
                    
                    If StsPTP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsPTP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsPTP + "@"
                        End If
                    End If
                    
                    If StsPOP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsPOP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsPOP + "@"
                        End If
                    End If
                    
                    If StsBP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsBP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsBP + "@"
                        End If
                    End If
                    
                    If StsSP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)  in (@" + StsSP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsSP + "@"
                        End If
                    End If
                    
                    If StsRP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsRP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsRP + "@"
                        End If
                    End If
                    
                    If StsSK <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsSK + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsSK + "@"
                        End If
                    End If
                    
                     If StsON <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsON + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsON + "@"
                        End If
                    End If
                    
                     If StsOP <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_new,1,3)   in (@" + StsOP + "@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@" + StsOP + "@"
                        End If
                    End If
                    
                    
                     If Stsblank <> "" Then
                        If Len(STRINGBLOK) = 0 Then
                                STRINGBLOK = " substring(F_cek_NEW,1,3)   in (@@"
                        Else
                                STRINGBLOK = STRINGBLOK + ",@@"
                        End If
                    End If
                    
                    
                    
                
                    If Len(STRINGBLOK) > 0 Then
                            STRINGBLOK = STRINGBLOK + ")"
                    End If
                    getblock = STRINGBLOK
End Function
Private Sub SSOption1_Click(Index As Integer, Value As Integer)
Dim M_objrs As ADODB.Recordset
Dim cmdsqluser As String

Select Case Index
Case 0
        Combo1(0).Enabled = False
        Combo1(1).Enabled = False
Case 1
        Combo1(0).Enabled = True
        Combo1(1).Enabled = True
        Combo1(0).CLEAR
        Combo1(1).CLEAR
        
        '@@221010
'        Dim M_DATA As New CLS_FRMSEARCH
'        Set M_OBJRS = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "")
'            While Not M_OBJRS.EOF
'                Combo1(0).AddItem M_OBJRS("USERID")
'                Combo1(1).AddItem M_OBJRS("AGENT")
'                M_OBJRS.MoveNext
'            Wend
'        Set M_OBJRS = Nothing

       '@@ 11-11-10 tambahan kode jika tl yang menggunakan
        If Left(Trim(MDIForm1.txtusername.Text), 2) = "TL" Then
            cmdsqluser = "select * from usertbl where usertype='1' and spvcode='"
            cmdsqluser = cmdsqluser + Replace(Trim(MDIForm1.txtusername.Text), "TL", "SPV") + "'"
            cmdsqluser = cmdsqluser + " order by userid asc"
        Else
           cmdsqluser = "select * from usertbl where usertype='1' order by userid asc"
        End If
       
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open cmdsqluser, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        While Not M_objrs.EOF
            Combo1(0).AddItem M_objrs("userid")
            Combo1(1).AddItem M_objrs("agent")
            M_objrs.MoveNext
        Wend
        
        'SSOption1(0).Value = True
        spv = False
Case 2
        Combo1(0).Enabled = True
        Combo1(1).Enabled = True
        Combo1(0).CLEAR
        Combo1(1).CLEAR
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        If UCase(MDIForm1.txtlevel.Text) = "SUPERVISOR" Then
            M_objrs.Open "select * from SPVTBL ", M_OBJCONN, adOpenDynamic, adLockOptimistic
        ElseIf UCase(MDIForm1.txtlevel.Text) = "ADMINISTRATOR" Or UCase(MDIForm1.txtlevel.Text) = "ADMIN" Then
            M_objrs.Open "select * from SPVTBL", M_OBJCONN, adOpenDynamic, adLockOptimistic
        ElseIf UCase(MDIForm1.txtlevel.Text) = "TEAMLEADER" Then
        M_objrs.Open "select * from SPVTBL where team='" + MDIForm1.txtusername.Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
        End If
            While Not M_objrs.EOF
                Combo1(0).AddItem M_objrs("SPVCODE")
                Combo1(1).AddItem M_objrs("SPVNAME")
                M_objrs.MoveNext
            Wend
        Set M_objrs = Nothing
        spv = True
'        SSOption1(0).Value = True
'        SSOption1(1).Value = True
        
End Select

End Sub
Public Sub getMarkup()
Dim list As listitem
Dim RSNEW As New ADODB.Recordset
Set Rs = New ADODB.Recordset
RSNEW.CursorLocation = adUseClient
RSNEW.Open "select distinct(exclude) from mgm WHERE (exclude <>'')", M_OBJCONN, adOpenDynamic, adLockOptimistic
ListView1.ListItems.CLEAR
While Not RSNEW.EOF
Set list = ListView1.ListItems.ADD(, , IIf(IsNull(RSNEW!exclude), "", RSNEW!exclude))
    RSNEW.MoveNext
Wend

End Sub

Public Function GETSELECTMARKUP() As String
Dim J As Integer
Dim TMPSELECTMARKUP As String
GETSELECTMARKUP = ""
For J = 1 To ListView2.ListItems.Count
        If J = 1 Then
            TMPSELECTMARKUP = TMPSELECTMARKUP + Chr(39) + ListView2.ListItems(J).Text + Chr(39)
        Else
            TMPSELECTMARKUP = TMPSELECTMARKUP + "," + Chr(39) + ListView2.ListItems(J).Text + Chr(39)
        End If
    
        
Next J
GETSELECTMARKUP = TMPSELECTMARKUP
End Function

'@@ 18-11-10 ini perbaikan script blok entry data
Public Function GetBlockEntry() As String
    Dim StringBlokEntry As String


                    StringBlokEntry = ""
                    
                    If StsNewEntry <> "" Then
                        If Len(StringBlokEntry) = 0 Then
                            StringBlokEntry = StsNewEntry
                        Else
                            StringBlokEntry = StringBlokEntry + " or " + StsNewEntry
                        End If
                    End If
                    
                    If StsReguler <> "" Then
                        If Len(StringBlokEntry) = 0 Then
                                StringBlokEntry = StsReguler
                        Else
                                StringBlokEntry = StringBlokEntry + " or " + StsReguler
                        End If
                    End If
                    
                    If StsSwap <> "" Then
                        If Len(StringBlokEntry) = 0 Then
                                StringBlokEntry = StsSwap
                        Else
                                StringBlokEntry = StringBlokEntry + " or " + StsSwap
                        End If
                    End If
                    
                    If StsCurrent <> "" Then
                        If Len(StringBlokEntry) = 0 Then
                                StringBlokEntry = StsCurrent
                        Else
                                StringBlokEntry = StringBlokEntry + " or " + StsCurrent
                        End If
                    End If
                    
                    'Ini buat ngasih kurung buka dan kurung tutup blok entry
                    If StringBlokEntry <> "" Then
                        StringBlokEntry = "( " + StringBlokEntry
                        StringBlokEntry = StringBlokEntry + " )"
                    End If
                    
                    GetBlockEntry = StringBlokEntry
End Function

Private Sub CekValidNullLock()
     If (IsNull(StsVl) Or StsVl = "") _
        And (IsNull(StsOS) Or StsOS = "") _
        And (IsNull(StsON) Or StsON = "") _
        And (IsNull(StsSK) Or StsSK = "") _
        And (IsNull(StsPR) Or StsPR = "") _
        And (IsNull(StsPTP) Or StsPTP = "") _
        And (IsNull(StsBP) Or StsBP = "") _
        And (IsNull(StsPOP) Or StsPOP = "") _
        And (IsNull(StsSP) Or StsSP = "") _
        And (IsNull(StsRP) Or StsRP = "") _
        And (IsNull(StsOP) Or StsOP = "") _
        And (IsNull(Stsblank) Or Stsblank = "") _
        And (IsNull(Stsuncontact) Or Stsuncontact = "") _
        And (IsNull(StsNewEntry) Or StsNewEntry = "") _
        And (IsNull(StsReguler) Or StsReguler = "") _
        And (IsNull(StsSwap) Or StsSwap = "") _
        And (IsNull(StsCurrent) Or StsCurrent = "") _
        And (ListView2.ListItems.Count = 0) Then
        
        MsgBox "Anda belum memilih, status data yang akan di lock!", vbOKOnly + vbInformation, "Informasi"
        CekValidLock = False
        Exit Sub
     End If
     CekValidLock = True
End Sub

Private Sub Timer_stopwatch_Timer()
    
    'Tambah dengan satu untuk total sepersepuluh detik.
    'Kita mengeset interval Timer menjadi 10, jadi
    'setiap sepersepuluh detik prosedur ini akan
    'dieksekusi
    TotalTenthDetik = TotalTenthDetik + 1
    'Jika TotalTenthSeconds = 10,
    'set kembali menjadi 0.
    TenthDetik = TotalTenthDetik Mod 10
    '10 kali sepersepuluh detik sama dengan 1 detik.
    'int - akan mengembalikan bilangan integer (bulat)
    'dari pecahan 'Contoh: Int(0.9) = 0 menghasilkan 0
    TotalDetik = Int(TotalTenthDetik / 10)
    'Jika variabel Seconds = 60, set kembali menjadi 0
    Detik = TotalDetik Mod 60
    If Len(Detik) = 1 Then
       Detik = "0" & Detik  'Agar selalu dalam dua
                            'digit
    End If
    Menit = Int(TotalDetik / 60) Mod 60
    If Len(Menit) = 1 Then
       Menit = "0" & Menit    'Agar selalu dalam dua
                          'digit
    End If
    JAM = Int(TotalDetik / 3600)
    If JAM < 9 Then
       Jam1 = "0" & JAM       'Agar selalu dalam dua'digit
    End If
    'Tampilkan hasilnya di Lblwaktu (update terus Lblwaktu)
    LblWaktu.Caption = Jam1 & ":" & Menit & ":" & Detik & ":" & TenthDetik & ""
End Sub
