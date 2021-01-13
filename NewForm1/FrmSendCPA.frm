VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FrmSendCPA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List Send CPA"
   ClientHeight    =   9660
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11940
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9660
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   9555
      Left            =   60
      TabIndex        =   0
      Tag             =   "0"
      Top             =   60
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   16854
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "List Sending CPA"
      TabPicture(0)   =   "FrmSendCPA.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CmdApprove"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "LstCpa"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdcpa(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "SSPanel1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdcpa(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdcpa(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Detail CPA"
      TabPicture(1)   =   "FrmSendCPA.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Frame13"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame13 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame13"
         ForeColor       =   &H80000008&
         Height          =   9015
         Left            =   -75120
         TabIndex        =   53
         Top             =   390
         Width           =   6510
         Begin VB.TextBox label8 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   3780
            Locked          =   -1  'True
            TabIndex        =   78
            Top             =   7335
            Width           =   1995
         End
         Begin VB.TextBox label5 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   3780
            Locked          =   -1  'True
            TabIndex        =   77
            Top             =   6570
            Width           =   1995
         End
         Begin VB.TextBox txtperiodpay 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1530
            Locked          =   -1  'True
            TabIndex        =   76
            Top             =   7650
            Width           =   2175
         End
         Begin VB.Frame Frame15 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   2
            Left            =   270
            TabIndex        =   74
            Top             =   5805
            Width           =   6090
            Begin VB.Label Label38 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Payment Arrangement"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   2
               Left            =   510
               TabIndex        =   75
               Top             =   90
               Width           =   3255
            End
            Begin VB.Image Image1 
               Height          =   375
               Index           =   2
               Left            =   75
               Picture         =   "FrmSendCPA.frx":0038
               Stretch         =   -1  'True
               Top             =   40
               Width           =   375
            End
         End
         Begin VB.TextBox txtagency 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1530
            Locked          =   -1  'True
            TabIndex        =   73
            Top             =   5400
            Width           =   2220
         End
         Begin VB.TextBox txtplace 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1530
            Locked          =   -1  'True
            TabIndex        =   72
            Top             =   5040
            Width           =   2220
         End
         Begin VB.TextBox txtcollect 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1530
            Locked          =   -1  'True
            TabIndex        =   71
            Top             =   4680
            Width           =   2220
         End
         Begin VB.ComboBox cbosts 
            Height          =   315
            ItemData        =   "FrmSendCPA.frx":18D2
            Left            =   1530
            List            =   "FrmSendCPA.frx":18E2
            TabIndex        =   70
            Top             =   4320
            Width           =   2220
         End
         Begin VB.TextBox txtcycle 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1530
            Locked          =   -1  'True
            TabIndex        =   69
            Top             =   3960
            Width           =   1995
         End
         Begin VB.TextBox txtname 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1530
            Locked          =   -1  'True
            TabIndex        =   68
            Top             =   3285
            Width           =   4875
         End
         Begin VB.TextBox txtcardno 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1500
            Locked          =   -1  'True
            TabIndex        =   67
            Top             =   2970
            Width           =   1995
         End
         Begin VB.Frame Frame15 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   1
            Left            =   270
            TabIndex        =   65
            Top             =   2430
            Width           =   6180
            Begin VB.Image Image1 
               Height          =   375
               Index           =   1
               Left            =   75
               Picture         =   "FrmSendCPA.frx":18F6
               Stretch         =   -1  'True
               Top             =   40
               Width           =   375
            End
            Begin VB.Label Label38 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Account Overview"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   510
               TabIndex        =   66
               Top             =   90
               Width           =   3255
            End
         End
         Begin VB.TextBox txtarrangement 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1485
            Locked          =   -1  'True
            TabIndex        =   64
            Top             =   1980
            Width           =   1455
         End
         Begin VB.TextBox txtproduct 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1485
            Locked          =   -1  'True
            TabIndex        =   63
            Top             =   1620
            Width           =   1455
         End
         Begin VB.TextBox txtreff 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1485
            MaxLength       =   1
            TabIndex        =   62
            Top             =   1260
            Width           =   1455
         End
         Begin VB.TextBox txtregion 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1485
            Locked          =   -1  'True
            TabIndex        =   61
            Top             =   585
            Width           =   1440
         End
         Begin VB.Frame Frame15 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   0
            Left            =   225
            TabIndex        =   59
            Top             =   90
            Width           =   6180
            Begin VB.Label Label38 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Request Info"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   510
               TabIndex        =   60
               Top             =   45
               Width           =   1455
            End
            Begin VB.Image Image1 
               Height          =   375
               Index           =   0
               Left            =   75
               Picture         =   "FrmSendCPA.frx":3190
               Stretch         =   -1  'True
               Top             =   40
               Width           =   375
            End
         End
         Begin VB.TextBox TxtLPDPayment 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   3840
            Locked          =   -1  'True
            TabIndex        =   58
            Top             =   4680
            Width           =   1995
         End
         Begin VB.TextBox TxtCustidMMU 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3600
            TabIndex        =   57
            Top             =   3000
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox TxtIdCpa 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   3000
            TabIndex        =   56
            Top             =   600
            Width           =   675
         End
         Begin VB.CommandButton CmdSendApproval 
            Caption         =   "&Send Approval "
            Height          =   435
            Left            =   3720
            TabIndex        =   55
            Top             =   8520
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.ComboBox CmbSendApproval 
            Height          =   315
            ItemData        =   "FrmSendCPA.frx":4A2A
            Left            =   1740
            List            =   "FrmSendCPA.frx":4A3A
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   8580
            Visible         =   0   'False
            Width           =   1875
         End
         Begin TDBDate6Ctl.TDBDate dtpropsal 
            Height          =   255
            Left            =   1485
            TabIndex        =   79
            Top             =   945
            Width           =   1410
            _Version        =   65536
            _ExtentX        =   2487
            _ExtentY        =   450
            Calendar        =   "FrmSendCPA.frx":4A68
            Caption         =   "FrmSendCPA.frx":4B80
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FrmSendCPA.frx":4BEC
            Keys            =   "FrmSendCPA.frx":4C0A
            Spin            =   "FrmSendCPA.frx":4C68
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   16777215
            BorderStyle     =   0
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
            Value           =   3.54028054673894E-316
            CenturyMode     =   0
         End
         Begin TDBDate6Ctl.TDBDate dtcardopen 
            Height          =   255
            Left            =   1530
            TabIndex        =   80
            Top             =   3645
            Width           =   2250
            _Version        =   65536
            _ExtentX        =   3969
            _ExtentY        =   450
            Calendar        =   "FrmSendCPA.frx":4C90
            Caption         =   "FrmSendCPA.frx":4DA8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FrmSendCPA.frx":4E14
            Keys            =   "FrmSendCPA.frx":4E32
            Spin            =   "FrmSendCPA.frx":4E90
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   16744576
            BorderStyle     =   0
            BtnPositioning  =   0
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "dd/mm/yyyy"
            EditMode        =   0
            Enabled         =   0
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   1
            Value           =   3.54028054673894E-316
            CenturyMode     =   0
         End
         Begin TDBNumber6Ctl.TDBNumber lblLastPay 
            Height          =   255
            Left            =   1530
            TabIndex        =   81
            Top             =   6660
            Width           =   2160
            _Version        =   65536
            _ExtentX        =   3810
            _ExtentY        =   450
            Calculator      =   "FrmSendCPA.frx":4EB8
            Caption         =   "FrmSendCPA.frx":4ED8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FrmSendCPA.frx":4F44
            Keys            =   "FrmSendCPA.frx":4F62
            Spin            =   "FrmSendCPA.frx":4FAC
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   16777215
            BorderStyle     =   0
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,###,##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###,###,###,##0"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   99999999999999
            MinValue        =   -99999999999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   5
            Value           =   0
            MaxValueVT      =   1701642245
            MinValueVT      =   3801093
         End
         Begin TDBNumber6Ctl.TDBNumber txtdownpayment 
            Height          =   255
            Left            =   1530
            TabIndex        =   82
            Top             =   7020
            Width           =   2160
            _Version        =   65536
            _ExtentX        =   3810
            _ExtentY        =   450
            Calculator      =   "FrmSendCPA.frx":4FD4
            Caption         =   "FrmSendCPA.frx":4FF4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FrmSendCPA.frx":5060
            Keys            =   "FrmSendCPA.frx":507E
            Spin            =   "FrmSendCPA.frx":50C8
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   16777215
            BorderStyle     =   0
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,###,##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###,###,###,##0"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   99999999999999
            MinValue        =   -99999999999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   1245189
            Value           =   0
            MaxValueVT      =   1701642245
            MinValueVT      =   3801093
         End
         Begin TDBNumber6Ctl.TDBNumber txtfuture 
            Height          =   255
            Left            =   1530
            TabIndex        =   83
            Top             =   7335
            Width           =   2160
            _Version        =   65536
            _ExtentX        =   3810
            _ExtentY        =   450
            Calculator      =   "FrmSendCPA.frx":50F0
            Caption         =   "FrmSendCPA.frx":5110
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FrmSendCPA.frx":517C
            Keys            =   "FrmSendCPA.frx":519A
            Spin            =   "FrmSendCPA.frx":51E4
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   16744576
            BorderStyle     =   0
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,###,##0"
            EditMode        =   0
            Enabled         =   0
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###,###,###,##0"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   99999999999999
            MinValue        =   -99999999999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   5
            Value           =   0
            MaxValueVT      =   1701642245
            MinValueVT      =   3801093
         End
         Begin TDBNumber6Ctl.TDBNumber txtprincipal 
            Height          =   255
            Left            =   1530
            TabIndex        =   84
            Top             =   8055
            Width           =   2160
            _Version        =   65536
            _ExtentX        =   3810
            _ExtentY        =   450
            Calculator      =   "FrmSendCPA.frx":520C
            Caption         =   "FrmSendCPA.frx":522C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FrmSendCPA.frx":5298
            Keys            =   "FrmSendCPA.frx":52B6
            Spin            =   "FrmSendCPA.frx":5300
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   16777215
            BorderStyle     =   0
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,###,##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###,###,###,##0"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   99999999999999
            MinValue        =   -99999999999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   5
            Value           =   0
            MaxValueVT      =   1701642245
            MinValueVT      =   3801093
         End
         Begin TDBNumber6Ctl.TDBNumber txtbalance 
            Height          =   255
            Left            =   1530
            TabIndex        =   85
            Top             =   6300
            Width           =   1800
            _Version        =   65536
            _ExtentX        =   3175
            _ExtentY        =   450
            Calculator      =   "FrmSendCPA.frx":5328
            Caption         =   "FrmSendCPA.frx":5348
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FrmSendCPA.frx":53B4
            Keys            =   "FrmSendCPA.frx":53D2
            Spin            =   "FrmSendCPA.frx":541C
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   16777215
            BorderStyle     =   0
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,###,##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###,###,###,##0"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   99999999999999
            MinValue        =   -99999999999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   5
            Value           =   0
            MaxValueVT      =   1701642245
            MinValueVT      =   3801093
         End
         Begin TDBDate6Ctl.TDBDate dwo 
            Height          =   285
            Left            =   4620
            TabIndex        =   86
            Top             =   3660
            Width           =   1620
            _Version        =   65536
            _ExtentX        =   2857
            _ExtentY        =   503
            Calendar        =   "FrmSendCPA.frx":5444
            Caption         =   "FrmSendCPA.frx":555C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FrmSendCPA.frx":55C8
            Keys            =   "FrmSendCPA.frx":55E6
            Spin            =   "FrmSendCPA.frx":5644
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   16777215
            BorderStyle     =   0
            BtnPositioning  =   0
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "dd/mm/yyyy"
            EditMode        =   0
            Enabled         =   0
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
            Value           =   3.54028054673894E-316
            CenturyMode     =   0
         End
         Begin TDBNumber6Ctl.TDBNumber tdbisnstallment 
            Height          =   255
            Left            =   3840
            TabIndex        =   87
            Top             =   7980
            Width           =   2160
            _Version        =   65536
            _ExtentX        =   3810
            _ExtentY        =   450
            Calculator      =   "FrmSendCPA.frx":566C
            Caption         =   "FrmSendCPA.frx":568C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FrmSendCPA.frx":56F8
            Keys            =   "FrmSendCPA.frx":5716
            Spin            =   "FrmSendCPA.frx":5760
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   16777215
            BorderStyle     =   0
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,###,##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###,###,###,##0"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   99999999999999
            MinValue        =   -99999999999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   5
            Value           =   0
            MaxValueVT      =   1701642245
            MinValueVT      =   3801093
         End
         Begin TDBNumber6Ctl.TDBNumber TxtLPAPayment 
            Height          =   255
            Left            =   3840
            TabIndex        =   88
            Top             =   5280
            Width           =   2160
            _Version        =   65536
            _ExtentX        =   3810
            _ExtentY        =   450
            Calculator      =   "FrmSendCPA.frx":5788
            Caption         =   "FrmSendCPA.frx":57A8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FrmSendCPA.frx":5814
            Keys            =   "FrmSendCPA.frx":5832
            Spin            =   "FrmSendCPA.frx":587C
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   16744576
            BorderStyle     =   0
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,###,##0"
            EditMode        =   0
            Enabled         =   0
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###,###,###,##0"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   99999999999999
            MinValue        =   -99999999999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   5
            Value           =   0
            MaxValueVT      =   1701642245
            MinValueVT      =   3801093
         End
         Begin VB.Label Label7 
            Caption         =   "Principal di database"
            Height          =   285
            Left            =   3780
            TabIndex        =   115
            Top             =   6975
            Width           =   1590
         End
         Begin VB.Label Label6 
            Caption         =   "Balance di database"
            Height          =   285
            Left            =   3735
            TabIndex        =   114
            Top             =   6300
            Width           =   1455
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FF8080&
            Caption         =   ")*  D=SETTLEMENT R=RESCHEDULE X=PAID OFF"
            ForeColor       =   &H00C0FFC0&
            Height          =   990
            Left            =   3060
            TabIndex        =   113
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Principal"
            Height          =   285
            Index           =   19
            Left            =   315
            TabIndex        =   112
            Top             =   8100
            Width           =   1230
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Payment period month"
            Height          =   465
            Index           =   17
            Left            =   315
            TabIndex        =   111
            Top             =   7695
            Width           =   1230
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Future Payment"
            Height          =   195
            Index           =   16
            Left            =   315
            TabIndex        =   110
            Top             =   7380
            Width           =   1230
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "down Payment"
            Height          =   195
            Index           =   15
            Left            =   315
            TabIndex        =   109
            Top             =   7020
            Width           =   1230
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Payment"
            Height          =   330
            Index           =   14
            Left            =   315
            TabIndex        =   108
            Top             =   6705
            Width           =   1230
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Balance"
            Height          =   285
            Index           =   13
            Left            =   360
            TabIndex        =   107
            Top             =   6345
            Width           =   1230
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Agency name"
            Height          =   240
            Index           =   12
            Left            =   405
            TabIndex        =   106
            Top             =   5445
            Width           =   1230
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "placement"
            Height          =   240
            Index           =   11
            Left            =   405
            TabIndex        =   105
            Top             =   5085
            Width           =   1230
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "collector"
            Height          =   240
            Index           =   10
            Left            =   360
            TabIndex        =   104
            Top             =   4770
            Width           =   1230
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "status card"
            Height          =   240
            Index           =   9
            Left            =   360
            TabIndex        =   103
            Top             =   4365
            Width           =   1230
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Cycle Dlq"
            Height          =   240
            Index           =   8
            Left            =   360
            TabIndex        =   102
            Top             =   4005
            Width           =   1230
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Card Open"
            Height          =   240
            Index           =   7
            Left            =   360
            TabIndex        =   101
            Top             =   3645
            Width           =   1230
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "cust name"
            Height          =   240
            Index           =   6
            Left            =   360
            TabIndex        =   100
            Top             =   3330
            Width           =   1230
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Card no"
            Height          =   240
            Index           =   5
            Left            =   360
            TabIndex        =   99
            Top             =   3015
            Width           =   1230
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Arrangement"
            Height          =   240
            Index           =   4
            Left            =   360
            TabIndex        =   98
            Top             =   2025
            Width           =   1230
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Product"
            Height          =   240
            Index           =   3
            Left            =   360
            TabIndex        =   97
            Top             =   1665
            Width           =   1230
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Reffno"
            Height          =   240
            Index           =   2
            Left            =   360
            TabIndex        =   96
            Top             =   1305
            Width           =   1230
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Proposal Date"
            Height          =   240
            Index           =   1
            Left            =   315
            TabIndex        =   95
            Top             =   990
            Width           =   1140
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Region"
            Height          =   240
            Index           =   18
            Left            =   315
            TabIndex        =   94
            Top             =   630
            Width           =   1230
         End
         Begin VB.Label Label11 
            BackColor       =   &H00FF8080&
            Caption         =   "Wo date"
            Height          =   285
            Left            =   3870
            TabIndex        =   93
            Top             =   3660
            Width           =   795
         End
         Begin VB.Label Label12 
            Caption         =   "Installment Period"
            Height          =   375
            Left            =   3810
            TabIndex        =   92
            Top             =   7620
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "LPD from Payment:"
            Height          =   240
            Index           =   20
            Left            =   3900
            TabIndex        =   91
            Top             =   4380
            Width           =   1830
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "LPA from Payment:"
            Height          =   240
            Index           =   22
            Left            =   3840
            TabIndex        =   90
            Top             =   5040
            Width           =   1890
         End
         Begin VB.Label Label14 
            Caption         =   "Send Approval to:"
            Height          =   315
            Left            =   300
            TabIndex        =   89
            Top             =   8580
            Visible         =   0   'False
            Width           =   1755
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame13"
         ForeColor       =   &H80000008&
         Height          =   9045
         Left            =   -68490
         TabIndex        =   10
         Top             =   360
         Width           =   5265
         Begin VB.Frame Frame15 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   5
            Left            =   -45
            TabIndex        =   30
            Top             =   90
            Width           =   4785
            Begin VB.Image Image1 
               Height          =   375
               Index           =   5
               Left            =   75
               Picture         =   "FrmSendCPA.frx":58A4
               Stretch         =   -1  'True
               Top             =   40
               Width           =   375
            End
            Begin VB.Label Label38 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Calculation"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   5
               Left            =   510
               TabIndex        =   31
               Top             =   45
               Width           =   1455
            End
         End
         Begin VB.TextBox txtfrombalancepersen 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1845
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   1260
            Width           =   2085
         End
         Begin VB.TextBox txtpersenprincipal 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1845
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   1665
            Width           =   2085
         End
         Begin VB.Frame Frame15 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   4
            Left            =   0
            TabIndex        =   26
            Top             =   2430
            Width           =   4650
            Begin VB.Label Label38 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Background"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   4
               Left            =   510
               TabIndex        =   27
               Top             =   90
               Width           =   3255
            End
            Begin VB.Image Image1 
               Height          =   375
               Index           =   4
               Left            =   75
               Picture         =   "FrmSendCPA.frx":713E
               Stretch         =   -1  'True
               Top             =   40
               Width           =   375
            End
         End
         Begin VB.TextBox txtoccupation 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1530
            TabIndex        =   25
            Top             =   2970
            Width           =   1995
         End
         Begin VB.TextBox txtreason 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1530
            TabIndex        =   24
            Top             =   3285
            Width           =   3120
         End
         Begin VB.TextBox txtpaymenthandle 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1530
            TabIndex        =   23
            Top             =   4005
            Width           =   2220
         End
         Begin VB.TextBox txtnodlq 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1530
            TabIndex        =   22
            Top             =   3645
            Width           =   2220
         End
         Begin VB.TextBox txtjust 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   780
            Left            =   1530
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   21
            Top             =   4365
            Width           =   3165
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "DOC"
            Height          =   2535
            Left            =   60
            TabIndex        =   12
            Top             =   5520
            Width           =   4665
            Begin VB.CheckBox chkfaxed 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Faxed"
               Height          =   285
               Left            =   180
               TabIndex        =   20
               Top             =   180
               Width           =   1005
            End
            Begin VB.CheckBox chkwentalk 
               BackColor       =   &H00FFC0C0&
               Caption         =   "When Talking Sulun"
               Height          =   285
               Left            =   180
               TabIndex        =   19
               Top             =   420
               Width           =   1905
            End
            Begin VB.CheckBox chkKTP 
               BackColor       =   &H00FFC0C0&
               Caption         =   "KTP"
               Enabled         =   0   'False
               Height          =   165
               Left            =   360
               TabIndex        =   18
               Top             =   720
               Width           =   765
            End
            Begin VB.CheckBox chkpp 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Surper"
               Enabled         =   0   'False
               Height          =   165
               Left            =   1140
               TabIndex        =   17
               Top             =   720
               Width           =   825
            End
            Begin VB.CheckBox chkbillings 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Billings"
               Enabled         =   0   'False
               Height          =   405
               Left            =   360
               TabIndex        =   16
               Top             =   900
               Width           =   825
            End
            Begin VB.CheckBox Check1 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Others"
               Enabled         =   0   'False
               Height          =   225
               Left            =   360
               TabIndex        =   15
               Top             =   1260
               Width           =   795
            End
            Begin VB.TextBox txtothers 
               BackColor       =   &H00E0E0E0&
               Height          =   615
               Left            =   1260
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   14
               Top             =   1260
               Width           =   3225
            End
            Begin VB.CommandButton CmdJadwalPembayaran 
               Caption         =   "&List jadwal pembayaran"
               Enabled         =   0   'False
               Height          =   495
               Left            =   2460
               TabIndex        =   13
               Top             =   1920
               Visible         =   0   'False
               Width           =   1995
            End
         End
         Begin VB.ComboBox CmbApprove 
            Height          =   315
            ItemData        =   "FrmSendCPA.frx":89D8
            Left            =   120
            List            =   "FrmSendCPA.frx":89DA
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   8400
            Width           =   1875
         End
         Begin TDBNumber6Ctl.TDBNumber txtcharge 
            Height          =   255
            Left            =   1845
            TabIndex        =   32
            Top             =   585
            Width           =   2160
            _Version        =   65536
            _ExtentX        =   3810
            _ExtentY        =   450
            Calculator      =   "FrmSendCPA.frx":89DC
            Caption         =   "FrmSendCPA.frx":89FC
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FrmSendCPA.frx":8A68
            Keys            =   "FrmSendCPA.frx":8A86
            Spin            =   "FrmSendCPA.frx":8AD0
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   16744576
            BorderStyle     =   0
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,###,##0"
            EditMode        =   0
            Enabled         =   0
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###,###,###,##0"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   99999999999999
            MinValue        =   -99999999999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   5
            Value           =   0
            MaxValueVT      =   1701642245
            MinValueVT      =   3801093
         End
         Begin TDBNumber6Ctl.TDBNumber txtdiscount 
            Height          =   255
            Left            =   1845
            TabIndex        =   33
            Top             =   945
            Width           =   2160
            _Version        =   65536
            _ExtentX        =   3810
            _ExtentY        =   450
            Calculator      =   "FrmSendCPA.frx":8AF8
            Caption         =   "FrmSendCPA.frx":8B18
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FrmSendCPA.frx":8B84
            Keys            =   "FrmSendCPA.frx":8BA2
            Spin            =   "FrmSendCPA.frx":8BEC
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   16744576
            BorderStyle     =   0
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,###,##0"
            EditMode        =   0
            Enabled         =   0
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###,###,###,##0"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   99999999999999
            MinValue        =   -99999999999999
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   5
            Value           =   0
            MaxValueVT      =   1701642245
            MinValueVT      =   3801093
         End
         Begin Threed.SSCommand SSCommand1 
            Height          =   660
            Index           =   0
            Left            =   2940
            TabIndex        =   34
            Tag             =   "0"
            Top             =   8040
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   1164
            _Version        =   196610
            Font3D          =   2
            MousePointer    =   16
            ForeColor       =   8388608
            PictureMaskColor=   -2147483644
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
            Picture         =   "FrmSendCPA.frx":8C14
            AutoSize        =   1
            Alignment       =   8
            PictureAlignment=   1
         End
         Begin Threed.SSCommand SSCommand1 
            Cancel          =   -1  'True
            Height          =   660
            Index           =   1
            Left            =   4590
            TabIndex        =   35
            Top             =   8040
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   1164
            _Version        =   196610
            Font3D          =   2
            MousePointer    =   16
            ForeColor       =   12582912
            PictureMaskColor=   -2147483644
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
            Picture         =   "FrmSendCPA.frx":9147
            AutoSize        =   1
            Alignment       =   4
            PictureAlignment=   1
         End
         Begin TDBDate6Ctl.TDBDate dtpelunasan 
            Height          =   255
            Left            =   1485
            TabIndex        =   36
            Top             =   5220
            Width           =   1620
            _Version        =   65536
            _ExtentX        =   2857
            _ExtentY        =   450
            Calendar        =   "FrmSendCPA.frx":97AC
            Caption         =   "FrmSendCPA.frx":98C4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FrmSendCPA.frx":9930
            Keys            =   "FrmSendCPA.frx":994E
            Spin            =   "FrmSendCPA.frx":99AC
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   16777215
            BorderStyle     =   0
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
            Value           =   3.54028054673894E-316
            CenturyMode     =   0
         End
         Begin Threed.SSCommand SSCommand1 
            Height          =   660
            Index           =   2
            Left            =   3780
            TabIndex        =   37
            Tag             =   "0"
            Top             =   8010
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   1164
            _Version        =   196610
            Font3D          =   2
            MousePointer    =   16
            ForeColor       =   8388608
            PictureMaskColor=   -2147483644
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
            Picture         =   "FrmSendCPA.frx":99D4
            AutoSize        =   1
            Alignment       =   8
            PictureAlignment=   1
         End
         Begin Crystal.CrystalReport RPT 
            Left            =   4470
            Top             =   1140
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin Threed.SSCommand CmdApprove2 
            Height          =   660
            Left            =   2040
            TabIndex        =   38
            Top             =   8220
            Visible         =   0   'False
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   1164
            _Version        =   196610
            BackColor       =   12640511
            Caption         =   "&Approve"
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Exit"
            Height          =   285
            Left            =   4815
            TabIndex        =   52
            Top             =   8775
            Width           =   825
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal Lunas"
            Height          =   240
            Index           =   21
            Left            =   90
            TabIndex        =   51
            Top             =   5265
            Width           =   1635
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Save"
            Height          =   285
            Left            =   3090
            TabIndex        =   50
            Top             =   8745
            Width           =   510
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Charge"
            Height          =   240
            Index           =   39
            Left            =   0
            TabIndex        =   49
            Top             =   630
            Width           =   1230
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Discount Amount"
            Height          =   240
            Index           =   38
            Left            =   0
            TabIndex        =   48
            Top             =   990
            Width           =   1140
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "From o/s balance %"
            Height          =   330
            Index           =   37
            Left            =   45
            TabIndex        =   47
            Top             =   1305
            Width           =   2220
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "principal (%) from "
            Height          =   240
            Index           =   36
            Left            =   45
            TabIndex        =   46
            Top             =   1665
            Width           =   1230
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "occupation"
            Height          =   240
            Index           =   34
            Left            =   0
            TabIndex        =   45
            Top             =   3015
            Width           =   1230
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "reason"
            Height          =   240
            Index           =   33
            Left            =   0
            TabIndex        =   44
            Top             =   3330
            Width           =   1230
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "payment handle by"
            Height          =   240
            Index           =   30
            Left            =   0
            TabIndex        =   43
            Top             =   4050
            Width           =   1230
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Justification"
            Height          =   240
            Index           =   29
            Left            =   0
            TabIndex        =   42
            Top             =   4365
            Width           =   1230
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "no of DlQ"
            Height          =   240
            Index           =   28
            Left            =   45
            TabIndex        =   41
            Top             =   3690
            Width           =   1230
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Preview"
            Height          =   285
            Left            =   3810
            TabIndex        =   40
            Top             =   8760
            Width           =   690
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Approve By:"
            Height          =   240
            Index           =   23
            Left            =   180
            TabIndex        =   39
            Top             =   8100
            Width           =   1230
         End
      End
      Begin Threed.SSCommand cmdcpa 
         Height          =   780
         Index           =   0
         Left            =   10740
         TabIndex        =   2
         Top             =   690
         Visible         =   0   'False
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1376
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FrmSendCPA.frx":9F07
         AutoSize        =   1
         Alignment       =   8
      End
      Begin Threed.SSCommand cmdcpa 
         Height          =   780
         Index           =   1
         Left            =   10740
         TabIndex        =   3
         Top             =   1830
         Visible         =   0   'False
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1376
         _Version        =   196610
         PictureFrames   =   1
         Picture         =   "FrmSendCPA.frx":A490
         AutoSize        =   1
         Alignment       =   8
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   390
         Left            =   0
         TabIndex        =   4
         Top             =   600
         Width           =   10410
         _ExtentX        =   18362
         _ExtentY        =   688
         _Version        =   196610
         Font3D          =   4
         ForeColor       =   12582912
         BackColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "List Send CPA"
         BevelWidth      =   2
         BorderWidth     =   1
         BevelOuter      =   1
         BevelInner      =   2
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSCommand cmdcpa 
         Height          =   780
         Index           =   3
         Left            =   10800
         TabIndex        =   5
         Top             =   4050
         Visible         =   0   'False
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1376
         _Version        =   196610
         Font3D          =   2
         MousePointer    =   16
         ForeColor       =   12582912
         PictureMaskColor=   -2147483644
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
         Picture         =   "FrmSendCPA.frx":AA19
         AutoSize        =   1
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin MSComctlLib.ListView LstCpa 
         Height          =   7620
         Left            =   60
         TabIndex        =   6
         Top             =   990
         Width           =   10380
         _ExtentX        =   18309
         _ExtentY        =   13441
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
      Begin Threed.SSCommand CmdApprove 
         Height          =   780
         Left            =   9240
         TabIndex        =   1
         Top             =   3120
         Visible         =   0   'False
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1376
         _Version        =   196610
         BackColor       =   12640511
         Caption         =   "&Approve"
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         Height          =   240
         Left            =   10680
         TabIndex        =   9
         Top             =   4875
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Remove"
         Height          =   240
         Left            =   10620
         TabIndex        =   8
         Top             =   2670
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ADD"
         Height          =   240
         Index           =   0
         Left            =   10620
         TabIndex        =   7
         Top             =   1530
         Visible         =   0   'False
         Width           =   1140
      End
   End
End
Attribute VB_Name = "FrmSendCPA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strtglapprrove  As String
Dim LpdPayment As String
Dim StatusChekcBox As String
Public IdCPA As String

Private Sub Check1_Click()
If Check1.Value = vbChecked Then
    txtothers.Enabled = True
    txtothers.BackColor = vbWhite
Else
    txtothers.Enabled = False
    txtothers.BackColor = &HC0C0C0
End If

End Sub

Private Sub chkfaxed_Click()
    '@@ 08092011
    If chkfaxed.Value = vbChecked Or chkwentalk.Value = vbChecked Then
        chkKTP.Enabled = True
        chkpp.Enabled = True
        chkbillings.Enabled = True
        Check1.Enabled = True
    End If
    If chkfaxed.Value = vbUnchecked And chkwentalk.Value = vbUnchecked Then
        chkKTP.Enabled = False
        chkpp.Enabled = False
        chkbillings.Enabled = False
        Check1.Enabled = False
        
        chkKTP.Value = vbUnchecked
        chkpp.Value = vbUnchecked
        chkbillings.Value = vbUnchecked
        Check1.Value = vbUnchecked
    End If
End Sub

Private Sub chkwentalk_Click()
    '@@ 08092011
    If chkfaxed.Value = vbChecked Or chkwentalk.Value = vbChecked Then
        chkKTP.Enabled = True
        chkpp.Enabled = True
        chkbillings.Enabled = True
        Check1.Enabled = True
    End If
    If chkfaxed.Value = vbUnchecked And chkwentalk.Value = vbUnchecked Then
        chkKTP.Enabled = False
        chkpp.Enabled = False
        chkbillings.Enabled = False
        Check1.Enabled = False
        
        chkKTP.Value = vbUnchecked
        chkpp.Value = vbUnchecked
        chkbillings.Value = vbUnchecked
        Check1.Value = vbUnchecked
    End If
End Sub



Private Sub CmdApprove_Click()
    Dim CMDSQL As String
    Dim Remarks As String
    Dim M_objrs As ADODB.Recordset
    Dim waktu As String
    Dim M_Objrs_Data As ADODB.Recordset
    
    If CmbApprove.Text = "" Then
        MsgBox "Combo approve by tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    
    If LstCpa.ListItems.Count = 0 Then
        MsgBox "Tidak ada data yang akan di Approve!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If LstCpa.SelectedItem.SubItems(32) = "1" Then
        MsgBox "Data sudah ditandatangan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    '@@21-09-2011
    'Ambil Tanggal Dari Server
    CMDSQL = "select now() as waktu "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    waktu = CStr(Format(M_objrs(0), "yyyy-mm-dd"))
    Set M_objrs = Nothing
    
    
    CMDSQL = "update mandiri.tblcpa set sts_approve='1', logapprove_by='"
    CMDSQL = CMDSQL + Trim(MDIForm1.TxtUsername.Text) + "', tglapprove='"
    CMDSQL = CMDSQL + waktu + "',approve_by='"
    CMDSQL = CMDSQL + CmbApprove.Text + "' "
    CMDSQL = CMDSQL + " where nid='"
    CMDSQL = CMDSQL + Trim(LstCpa.SelectedItem.Text) + "'"
    
    M_OBJCONN.Execute CMDSQL
    'Remarks = "App By:" + mdiform1.txtusername.text + "-"
    Remarks = "App By:" + CmbApprove.Text + "-"
    Remarks = Remarks + "Ttl.Pymt: Rp." + CStr(lblLastPay.Text) + " -"
    Remarks = Remarks + "Instl: " + CStr(tdbisnstallment.Text) + " -"
    Remarks = Remarks + "From Bal.: Rp." + CStr(txtbalance.Text) + " -"
    Remarks = Remarks + "From Prin.: Rp." + CStr(txtprincipal.Text) + " -"
    Remarks = Remarks + "%Balance: " + txtfrombalancepersen.Text + "% -"
    Remarks = Remarks + "%Principal: " + txtpersenprincipal.Text + "% "
    
    '--- Ambil Data dari Mgm
    STRSQL = "Select * from mandiri.mgm,mandiri.usertbl where mgm.custid='"
    STRSQL = STRSQL + LstCpa.SelectedItem.SubItems(1) + "' and mgm.agent=usertbl.userid"
    Set M_Objrs_Data = New ADODB.Recordset
    M_Objrs_Data.CursorLocation = adUseClient
    M_Objrs_Data.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    
    
    'With FrmCC_Colection
        CMDSQL = "insert into mandiri.mgm_hst (custid, agent, products, "
        CMDSQL = CMDSQL + "hst,user_log) values ('"
        CMDSQL = CMDSQL + M_Objrs_Data("custid") + "','"
        CMDSQL = CMDSQL + M_Objrs_Data("agent") + "','"
        CMDSQL = CMDSQL + "Collection" + "','"
        CMDSQL = CMDSQL + Remarks + "','"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "')"
        
        M_OBJCONN.Execute CMDSQL
    'End With
    
    LstCpa.SelectedItem.ForeColor = vbRed
    
    'With FrmCC_Colection
        'Kirim pesan ke agent, account yang di approve
        Remarks = "INFO CPA APPROVE!" + vbCrLf
        Remarks = Remarks + "Custid :" + M_Objrs_Data("custid") + vbCrLf
        Remarks = Remarks + "---------------------------------------" + vbCrLf
        'Remarks = Remarks + "App By:" + mdiform1.txtusername.text + "-"
        Remarks = Remarks + "App By:" + CmbApprove.Text + "-"
        Remarks = Remarks + "Ttl.Pymt: Rp." + CStr(lblLastPay.Text) + " -"
        Remarks = Remarks + "Instl: " + CStr(tdbisnstallment.Text) + " -"
        Remarks = Remarks + "From Bal.: Rp." + CStr(txtbalance.Text) + " -"
        Remarks = Remarks + "From Prin.: Rp." + CStr(txtprincipal.Text) + " -"
        Remarks = Remarks + "%Balance: " + txtfrombalancepersen.Text + "% -"
        Remarks = Remarks + "%Principal: " + txtpersenprincipal.Text + "% "
        
        CMDSQL = "insert into mandiri.msgtbl "
        CMDSQL = CMDSQL + "( recipient, datetime, sender, sentfrom, msg) values ('"
        CMDSQL = CMDSQL + M_Objrs_Data("userid") + "','"
        CMDSQL = CMDSQL + Format(Now(), "yyyymmdd") + "','"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "','"
        CMDSQL = CMDSQL + CStr(MDIForm1.Winsock1.LocalIP) + "','"
        CMDSQL = CMDSQL + Remarks + "')"
        
        M_OBJCONN.Execute CMDSQL
        
    'End With
    
    'Kirim Data Ke TLnya
    Remarks = "INFO CPA APPROVE!" + vbCrLf
    Remarks = Remarks + "Custid :" + M_Objrs_Data("custid") + vbCrLf
    Remarks = Remarks + "Agent :" + M_Objrs_Data("userid") + vbCrLf
    Remarks = Remarks + "---------------------------------------" + vbCrLf
    Remarks = Remarks + "App By:" + CmbApprove.Text + "-"
    Remarks = Remarks + "Ttl.Pymt: Rp." + CStr(lblLastPay.Text) + " -"
    Remarks = Remarks + "Instl: " + CStr(tdbisnstallment.Text) + " -"
    Remarks = Remarks + "From Bal.: Rp." + CStr(txtbalance.Text) + " -"
    Remarks = Remarks + "From Prin.: Rp." + CStr(txtprincipal.Text) + " -"
    Remarks = Remarks + "%Balance: " + txtfrombalancepersen.Text + "% -"
    Remarks = Remarks + "%Principal: " + txtpersenprincipal.Text + "% "
    
    CMDSQL = "insert into mandiri.msgtbl "
    CMDSQL = CMDSQL + "( recipient, datetime, sender, sentfrom, msg) values ('"
    CMDSQL = CMDSQL + M_Objrs_Data("team") + "','"
    CMDSQL = CMDSQL + Format(Now(), "yyyymmdd") + "','"
    CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "','"
    CMDSQL = CMDSQL + CStr(MDIForm1.Winsock1.LocalIP) + "','"
    CMDSQL = CMDSQL + Remarks + "')"
        
    M_OBJCONN.Execute CMDSQL
    
    CMDSQL = "delete from mandiri.tblsendcpa where nid='"
    CMDSQL = CMDSQL + Trim(TxtIdCpa.Text) + "'"
    M_OBJCONN.Execute CMDSQL
    
    MsgBox "Approve berhasil!", vbOKOnly + vbInformation, "Informasi"
End Sub

Private Sub CmdApprove2_Click()
    CmdApprove_Click
End Sub

Private Sub cmdcpa_Click(Index As Integer)
Dim rsfound As New ADODB.Recordset
    Select Case Index
    Case 0
           SSTab1.Tab = 1
           SSTab1.tag = 0
           
           Frame1.Enabled = True
            SSCommand1(0).tag = 1
            Label5.Text = IIf(FrmCC_Colection.lblAmount.ValueIsNull, "0", FrmCC_Colection.lblAmount)
            Label8.Text = IIf(FrmCC_Colection.lblPromPA.ValueIsNull, "0", FrmCC_Colection.lblPromPA)
            txtregion.Text = FrmCC_Colection.lblregion
            txtcardno.Text = FrmCC_Colection.lblcustid
            dwo.Value = FrmCC_Colection.lblBD.Value
            txtname.Text = FrmCC_Colection.lblnama.Caption
            txtproduct.Text = "CARD"
            dtcardopen.Value = FrmCC_Colection.lblOpenDate.Value
            txtplace.Text = "CardHolder"
            txtcollect.Text = VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11)
            Call Cari_LPD_LPA_Payment
    Case 1
         If LstCpa.ListItems.Count <> 0 Then
            If MsgBox("Yakin Akan dihapus...!!!!", vbQuestion + vbYesNo, "Peringatan") = vbYes Then
                       
                      STRSQL = "delete from mandiri.tblcpa where nid='" + LstCpa.SelectedItem.Text + "'"
                      M_OBJCONN.Execute (STRSQL)
                       
                    
                       STRSQL = "select * from mandiri.tblcpa where vcustid ='" + LstCpa.SelectedItem.SubItems(1) + "' order by dtglinsert asc  "
                       Set rsfound = New ADODB.Recordset
                       rsfound.CursorLocation = adUseClient
                       rsfound.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
                       If rsfound.RecordCount = 0 Then
                            STRSQL = "update  mandiri.mgm set stscpa=0, tglinsertfrmcpa =null,tglupdatefromcpa=null where custid='" + FrmCC_Colection.lblcustid.Caption + "'"
                            M_OBJCONN.Execute (STRSQL)
                       Else
                            rsfound.MoveLast
                            STRSQL = "update  mandiri.mgm set stscpa=0, tglinsertfrmcpa ='" + CStr(Format(rsfound("dtglinsert"), "yyyy-mm-dd hh:mm:ss")) + "',tglupdatefromcpa='" + CStr(Format(rsfound("dtgllastupdate"), "yyyy-mm-dd hh:mm:ss")) + "' where custid='" + FrmCC_Colection.lblcustid.Caption + "'"
                            M_OBJCONN.Execute (STRSQL)
                       End If
                       
                       LstCpa.ListItems.Remove LstCpa.SelectedItem.Index
                    MsgBox "Data Telah Di hapus"
            End If
         End If
     Case 3
     Unload Me
     
     
    End Select

End Sub

Private Sub CmdJadwalPembayaran_Click()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim NamaTL As String
    
    'Cari nama TL
    CMDSQL = "select * from mandiri.usertbl where userid='"
    CMDSQL = CMDSQL + Trim(FrmCC_Colection.lblaoc.Caption) + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
        NamaTL = M_objrs("team")
    End If
    
    Set M_objrs = Nothing
    
    IdCPA = TxtIdCpa.Text
    
    With FrmJadwalPembayaranCpa
        .TxtIdCpa.Text = TxtIdCpa.Text
        .TxtCustid.Text = txtcardno.Text
        .TxtAgent.Text = FrmCC_Colection.lblaoc.Caption
        .TxtInstallment.Value = IIf(IsNull(tdbisnstallment.Value), "0", tdbisnstallment.Value)
        .txtnama.Text = txtname.Text
        .txtPayment.Value = IIf(IsNull(lblLastPay.Value), "0", lblLastPay.Value)
        .TxtTL.Text = IIf(IsNull(NamaTL), "", NamaTL)
        .TxtAlamat.Text = IIf(IsNull(FrmCC_Colection.lblAddr.Text), "", FrmCC_Colection.lblAddr.Text)
        .txtbalance.Value = txtbalance.Value
        .TxtFromOs.Text = IIf(IsNull(txtfrombalancepersen.Text), "", txtfrombalancepersen.Text)
        
        'Cari Nomor Telepon
        CMDSQL = "select * from mandiri.mgm where custid='"
        CMDSQL = CMDSQL + Trim(txtcardno.Text) + "'"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_objrs.RecordCount > 0 Then
            With FrmJadwalPembayaranCpa
                .TxtNoTelp.CLEAR
                .TxtNoTelp.AddItem IIf(IsNull(M_objrs("mobileno")), "", M_objrs("mobileno"))
                .TxtNoTelp.AddItem IIf(IsNull(M_objrs("mobileno2")), "", M_objrs("mobileno2"))
                .TxtNoTelp.AddItem IIf(IsNull(M_objrs("mobilenoadd1")), "", M_objrs("mobilenoadd1"))
                .TxtNoTelp.AddItem IIf(IsNull(M_objrs("mobilenoadd2")), "", M_objrs("mobilenoadd2"))
                .TxtNoTelp.AddItem IIf(IsNull(M_objrs("homeno")), "", M_objrs("homeno"))
                .TxtNoTelp.AddItem IIf(IsNull(M_objrs("homeno2")), "", M_objrs("homeno2"))
                .TxtNoTelp.AddItem IIf(IsNull(M_objrs("homenoadd1")), "", M_objrs("homenoadd1"))
                .TxtNoTelp.AddItem IIf(IsNull(M_objrs("homenoadd2")), "", M_objrs("homenoadd2"))
                .TxtNoTelp.AddItem IIf(IsNull(M_objrs("officeno")), "", M_objrs("officeno"))
                .TxtNoTelp.AddItem IIf(IsNull(M_objrs("officenoadd1")), "", M_objrs("officenoadd1"))
                .TxtNoTelp.AddItem IIf(IsNull(M_objrs("officenoadd2")), "", M_objrs("officenoadd2"))
                .TxtNoTelp.AddItem IIf(IsNull(M_objrs("ec_telp")), "", M_objrs("ec_telp"))
                .TxtNoTelp.AddItem IIf(IsNull(M_objrs("telp_additional")), "", M_objrs("telp_additional"))
            End With
        End If
        
        Set M_objrs = Nothing
        
        
        .Show vbModal
    End With
End Sub




Private Sub CmdSendApproval_Click()
    Dim Amount As Double
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    
    Amount = txtbalance.Value - lblLastPay.Value
    
    If TxtIdCpa.Text = "" Then
         MsgBox "Simpan terlebih dahulu data CPA yang anda buat!", vbOKOnly + vbExclamation, "Peringatan"
         Exit Sub
    End If

    If Amount > 5000000 Then
        MsgBox "Amount tidak boleh kurang dari 5.000.000!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    If CmbSendApproval.Text = "" Then
        MsgBox "Anda belum menentukan kepada siapa send approval ditujukan!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    'Cek data, apakah data sebelumnya sudah di send??
    CMDSQL = "select * from mandiri.tblcpa where status_send='1' and nid='"
    CMDSQL = CMDSQL + Trim(TxtIdCpa.Text) + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
        MsgBox "Data sebelumnya sudah di send!", vbOKOnly + vbInformation, "Informasi"
        Set M_objrs = Nothing
        Exit Sub
    End If
    Set M_objrs = Nothing
    
    'Cek Data, Apakah sebelumnya data sudah di approve??
    CMDSQL = "select * from mandiri.tblcpa where nid='"
    CMDSQL = CMDSQL + Trim(TxtIdCpa.Text) + "' and sts_approve='1'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_objrs.RecordCount > 0 Then
        MsgBox "Data sudah di approve! oleh: " & M_objrs("approve_by") & "!", vbOKOnly + vbInformation, "Informasi"
        Set M_objrs = Nothing
        Exit Sub
    End If
    Set M_objrs = Nothing
    
    '-----Proses Send CPA
    CMDSQL = "update mandiri.tblcpa set tgl_send=now(), status_send='1', send_to='"
    CMDSQL = CMDSQL + Trim(CmbSendApproval.Text) + "' where nid='"
    CMDSQL = CMDSQL + Trim(TxtIdCpa.Text) + "'"
    M_OBJCONN.Execute CMDSQL
    
    
    CMDSQL = "insert into mandiri.tblsendcpa select * from mandiri.tblcpa "
    M_OBJCONN.Execute CMDSQL
    
    MsgBox "Data CPA berhasil dikirim ke: " + CmbSendApproval.Text & " untuk di approve!", vbOKOnly + vbInformation, "Informasi"
    
    
End Sub

Private Sub Form_Load()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    
    TxtCustidMMU.Text = ""
    
    'SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    Frame1.Enabled = False
    createHeader
    showlist
    IsiNamaApprove
    SSTab1.Tab = 0
    
    If UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
        cmdcpa(1).Enabled = False
        cmdcpa(0).Enabled = False
    End If
        
    If UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
        SSCommand1(0).Visible = False
        SSCommand1(1).Visible = True
        Label2.Visible = False
        Label3.Visible = True
        CmdApprove.Visible = False
        CmdApprove2.Visible = False
    End If

    If UCase(MDIForm1.txtlevel.Text) = "ADMIN" Or _
        UCase(MDIForm1.txtlevel.Text) = "ADMINISTRATOR" Or _
        UCase(MDIForm1.txtlevel.Text) = "SUPERVISOR" Then
            CmdApprove.Visible = True
            CmdApprove2.Visible = True
    End If
    cbosts.ListIndex = 1

 
    
End Sub

Public Sub createHeader()
    With LstCpa
        .ColumnHeaders.ADD 1, , "ID", 1000
        .ColumnHeaders.ADD 2, , "custid", 1000
        .ColumnHeaders.ADD 3, , "cust name", 2000
        .ColumnHeaders.ADD 4, , "Proposal Date", 1200
        .ColumnHeaders.ADD 5, , "reff no", 1200
        .ColumnHeaders.ADD 6, , "Product", 1300
        .ColumnHeaders.ADD 7, , "Arrangement", 1500
        .ColumnHeaders.ADD 8, , "card status", 1000
        .ColumnHeaders.ADD 9, , "Total Payment", 1500
        .ColumnHeaders.ADD 10, , "Down Payment", 1500
        .ColumnHeaders.ADD 11, , "future Pay", 1500
        .ColumnHeaders.ADD 12, , "Charges", 1500
        .ColumnHeaders.ADD 13, , "discount amount", 1
        .ColumnHeaders.ADD 14, , " O/S balance (%)", 1
        .ColumnHeaders.ADD 15, , " Principal (%)", 1
        .ColumnHeaders.ADD 16, , " verify", 1000
        .ColumnHeaders.ADD 17, , " Approvel ", 1000
        .ColumnHeaders.ADD 18, , " Tanggal Pelunasan ", 1200
        .ColumnHeaders.ADD 19, , "Justification ", 1
        .ColumnHeaders.ADD 20, , "Balance ", 1500
        .ColumnHeaders.ADD 21, , "Principal", 1500
        .ColumnHeaders.ADD 22, , "Tanggal lunas", 1500
        .ColumnHeaders.ADD 23, , "Tanggal Update", 1500
        .ColumnHeaders.ADD 24, , "Occupation", 1500
        .ColumnHeaders.ADD 25, , "Reason", 1
        .ColumnHeaders.ADD 26, , "DLQ", 1
        .ColumnHeaders.ADD 27, , "Payment Handle", 1
        .ColumnHeaders.ADD 28, , "Justification", 1
        .ColumnHeaders.ADD 29, , "Verify", 1
        .ColumnHeaders.ADD 30, , "Approvel", 1
        .ColumnHeaders.ADD 31, , "Tanggal Insert", 1500
        .ColumnHeaders.ADD 32, , "nperiod", 1
        .ColumnHeaders.ADD 33, , "Status Approve", 500
        .ColumnHeaders.ADD 34, , "LPD From Payment", 1500
        .ColumnHeaders.ADD 35, , "LPA From Payment", 1500
    End With

End Sub
Public Sub showlist()
       Dim M_objrs  As ADODB.Recordset
       Dim CMDSQL As String
       
       STRSQL = "select * from mandiri.tblsendcpa where send_to='"
       STRSQL = STRSQL + MDIForm1.TxtUsername.Text + "' order by tgl_send desc"
       
       Set rsTemporary = New ADODB.Recordset
       rsTemporary.CursorLocation = adUseClient
       rsTemporary.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
       LstCpa.ListItems.CLEAR
       While Not rsTemporary.EOF
            
                
                
            Set iListitem = LstCpa.ListItems.ADD(, , rsTemporary("nid"))
                iListitem.SubItems(1) = IIf(IsNull(rsTemporary("vcustid")), "", rsTemporary("vcustid"))
                iListitem.SubItems(2) = IIf(IsNull(rsTemporary("vcustname")), "", rsTemporary("vcustname"))
                iListitem.SubItems(3) = IIf(IsNull(rsTemporary("dpropsal")), "", rsTemporary("dpropsal"))
                iListitem.SubItems(4) = IIf(IsNull(rsTemporary("vreffno")), "", rsTemporary("vreffno"))
                iListitem.SubItems(5) = IIf(IsNull(rsTemporary("vproduct")), "", rsTemporary("vproduct"))
                iListitem.SubItems(6) = IIf(IsNull(rsTemporary("varragement")), "", rsTemporary("varragement"))
                iListitem.SubItems(7) = IIf(IsNull(rsTemporary("vcardsts")), "", rsTemporary("vcardsts"))
                iListitem.SubItems(8) = Format(IIf(IsNull(rsTemporary("nttlpayment")), 0, rsTemporary("nttlpayment")), "##,###")
                iListitem.SubItems(9) = Format(IIf(IsNull(rsTemporary("ndownpay")), 0, rsTemporary("ndownpay")), "##,###")
                iListitem.SubItems(10) = Format(IIf(IsNull(rsTemporary("nfuturepay")), 0, rsTemporary("nfuturepay")), "##,###")
                iListitem.SubItems(11) = Format(IIf(IsNull(rsTemporary("ncharge")), 0, rsTemporary("ncharge")), "##,###")
                iListitem.SubItems(12) = Format(IIf(IsNull(rsTemporary("ndiscountamt")), 0, rsTemporary("ndiscountamt")), "##,###")
                iListitem.SubItems(13) = IIf(IsNull(rsTemporary("vosbalance")), "", rsTemporary("vosbalance"))
                iListitem.SubItems(14) = IIf(IsNull(rsTemporary("vosprincipal")), "", rsTemporary("vosprincipal"))
                iListitem.SubItems(15) = IIf(IsNull(rsTemporary("vverify")), "", rsTemporary("vverify"))
                iListitem.SubItems(16) = IIf(IsNull(rsTemporary("votority")), "", rsTemporary("votority"))
                iListitem.SubItems(17) = IIf(IsNull(rsTemporary("dtglpelunasan")), "", Format(rsTemporary("dtglpelunasan"), "dd/mm/yyyy"))
               iListitem.SubItems(18) = IIf(IsNull(rsTemporary("vjust")), "", rsTemporary("vjust"))
               iListitem.SubItems(19) = Format(IIf(IsNull(rsTemporary("nbalance")), 0, rsTemporary("nbalance")), "##,###")
               iListitem.SubItems(20) = Format(IIf(IsNull(rsTemporary("nprincipal")), 0, rsTemporary("nprincipal")), "##,###")
               iListitem.SubItems(21) = IIf(IsNull(rsTemporary("dtglpelunasan")), "", Format(rsTemporary("dtglpelunasan"), "dd/mm/yyyy"))
                iListitem.SubItems(22) = IIf(IsNull(rsTemporary("dtgllastupdate")), "", Format(rsTemporary("dtgllastupdate"), "dd/mm/yyyy"))
                iListitem.SubItems(23) = IIf(IsNull(rsTemporary("voccupation")), "", rsTemporary("voccupation"))
                iListitem.SubItems(24) = IIf(IsNull(rsTemporary("vreason")), "", rsTemporary("vreason"))
                iListitem.SubItems(25) = IIf(IsNull(rsTemporary("vnodlq")), "", rsTemporary("vnodlq"))
                iListitem.SubItems(26) = IIf(IsNull(rsTemporary("vpaymenthandle")), "", rsTemporary("vpaymenthandle"))
                 iListitem.SubItems(27) = IIf(IsNull(rsTemporary("vjust")), "", rsTemporary("vjust"))
                iListitem.SubItems(28) = IIf(IsNull(rsTemporary("intverify")), "0", rsTemporary("intverify"))
                iListitem.SubItems(29) = IIf(IsNull(rsTemporary("intapprovel")), "0", rsTemporary("intapprovel"))
                iListitem.SubItems(30) = IIf(IsNull(rsTemporary("dtglinsert")), "", Format(rsTemporary("dtglinsert"), "dd/mm/yyyy"))
                 iListitem.SubItems(31) = IIf(IsNull(rsTemporary("nperiod")), 0, rsTemporary("nperiod"))
                strtglapprrove = IIf(IsNull(rsTemporary("tglapprove")), "0", Format(rsTemporary("tglapprove"), "dd/mm/yyyy"))
                 iListitem.SubItems(32) = IIf(IsNull(rsTemporary("sts_approve")), "", rsTemporary("sts_approve"))
                 
                 iListitem.SubItems(33) = IIf(IsNull(rsTemporary("lpd_from_payment")), "", Format(rsTemporary("lpd_from_payment"), "yyyy-mm-dd"))
                 iListitem.SubItems(34) = IIf(IsNull(rsTemporary("lpa_from_payment")), "0", rsTemporary("lpa_from_payment"))
                
                '@@ 16-03-2011, Jika sudah ditanda tangan akan berwarna merah
                If rsTemporary("sts_approve") = "1" Then
                    LstCpa.SelectedItem.ForeColor = vbRed
                Else
                    LstCpa.SelectedItem.ForeColor = vbBlack
                End If
            rsTemporary.MoveNext
       Wend
       
       
       Set rsTemporary = Nothing
       Set iListitem = Nothing
       
       Dim i As Integer
       For i = 1 To LstCpa.ListItems.Count
            If LstCpa.ListItems(i).SubItems(32) = "1" Then
                LstCpa.ListItems(i).ForeColor = vbRed
            End If
       Next i
       
'       '@@ 15-09-2011, Buat ambil data custidmmu untuk pil
'       If StatusCPA = "CPA Form 2" Then
'        'Jika Form CPA di load dari FrmCC2_Collection
'        CMDSQL = "select * from mgm where custid='" + frmCC_Colection2.lblCustId.Caption + "'"
'       Else
'        'Jika form CPA di load dari FrmCC_Collection
'        CMDSQL = "select * from mgm where custid='" + FrmCC_Colection.lblCustId.Caption + "'"
'       End If
'
'       Set M_OBJRS = New ADODB.Recordset
'       M_OBJRS.CursorLocation = adUseClient
'       M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'       If M_OBJRS("acc_type") = "PIL" Then
'         TxtCustidMMU.Text = IIf(IsNull(M_OBJRS("custidmmu")), "", M_OBJRS("custidmmu"))
'       End If
'
'       Set M_OBJRS = Nothing
       
End Sub
Private Sub lstCpa_DblClick()
    Dim RSNEW As New ADODB.Recordset
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    
 If LstCpa.ListItems.Count <> 0 Then
 Set RSNEW = New ADODB.Recordset
 RSNEW.CursorLocation = adUseClient
 stringSql = "select * from mandiri.tblcpa where nid =" + CStr(Val(LstCpa.SelectedItem.Text)) + ""
   
 RSNEW.Open stringSql, M_OBJCONN, adOpenDynamic, adLockOptimistic
 
 If Not RSNEW.EOF Then
    If IIf(IsNull(RSNEW!chkfaxed), "0", RSNEW!chkfaxed) = "1" Then
        chkfaxed.Value = vbChecked
    Else
       chkfaxed.Value = vbUnchecked
    End If
    
    If IIf(IsNull(RSNEW!chkwentalking), "0", RSNEW!chkwentalking) = "1" Then
        chkwentalk.Value = vbChecked
    Else
        chkwentalk.Value = vbUnchecked
    End If
    
    If IIf(IsNull(RSNEW!chkKTP), "0", RSNEW!chkKTP) = "1" Then
        chkKTP.Value = vbChecked
    Else
        chkKTP.Value = vbUnchecked
    End If
    
    If IIf(IsNull(RSNEW!chksup), "0", RSNEW!chksup) = "1" Then
        chkpp.Value = vbChecked
    Else
        chkpp.Value = vbUnchecked
    End If
    
    If IIf(IsNull(RSNEW!chkbillings), "0", RSNEW!chkbillings) = "1" Then
        chkbillings.Value = vbChecked
    Else
        chkbillings.Value = vbUnchecked
    End If
    
    If IIf(IsNull(RSNEW!chkothers), "0", RSNEW!chkothers) = "1" Then
        Check1.Value = vbChecked
    Else
        Check1.Value = vbUnchecked
    End If
    
 txtothers.Text = IIf(IsNull(RSNEW!ketother), "", RSNEW!ketother)
    
 End If
 
   SSTab1.tag = 1
 With FrmSendCPA
          SSTab1.Tab = 1
                    
          CMDSQL = "select * from mandiri.mgm where custid='"
          CMDSQL = CMDSQL + Trim(LstCpa.SelectedItem.SubItems(1)) + "'"
          Set M_objrs = New ADODB.Recordset
          M_objrs.CursorLocation = adUseClient
          M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
          
          
            .Caption = "Edit"
            Frame1.Enabled = True
            .SSCommand1(0).tag = 2
            .txtregion.Text = IIf(IsNull(M_objrs("region")), "", M_objrs("region")) 'FrmCC_Colection.lblregion
            .txtcardno.Text = IIf(IsNull(M_objrs("custid")), "", M_objrs("custid")) 'FrmCC_Colection.lblCustId.Caption
            .txtname.Text = IIf(IsNull(M_objrs("name")), "", M_objrs("name")) 'FrmCC_Colection.lblNama.Caption
            .txtproduct.Text = "CARD"
            .dtcardopen.Value = IIf(IsNull(M_objrs("OpenDate")), "", Format(M_objrs("OpenDate"), "dd/mm/yyyy")) 'FrmCC_Colection.lblOpenDate.Value
            .lblLastPay.Value = IIf(LstCpa.SelectedItem.SubItems(8) = "", "0", LstCpa.SelectedItem.SubItems(8))
            .txtdownpayment.Value = IIf(LstCpa.SelectedItem.SubItems(9) = "", "0", LstCpa.SelectedItem.SubItems(9))
            .txtplace.Text = "CardHolder"
            .dwo.Value = IIf(IsNull(M_objrs("b_d")), "", Format(M_objrs("b_d"), "dd/mm/yyyy")) 'FrmCC_Colection.lblBD.Value
            .Label5.Text = IIf(FrmCC_Colection.lblAmount.ValueIsNull, "0", FrmCC_Colection.lblAmount)
            .Label8.Text = IIf(FrmCC_Colection.lblPromPA.ValueIsNull, "0", FrmCC_Colection.lblPromPA)
            .txtreff = LstCpa.SelectedItem.SubItems(4)
            .txtcharge = IIf(LstCpa.SelectedItem.SubItems(10) = "", "0", LstCpa.SelectedItem.SubItems(10))
            .txtprincipal.Value = IIf(LstCpa.SelectedItem.SubItems(20) = "", "0", LstCpa.SelectedItem.SubItems(20))
            .txtcollect.Text = IIf(IsNull(M_objrs("agent")), "", M_objrs("agent")) 'VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11)
            .cbosts.Text = IIf(LstCpa.SelectedItem.SubItems(7) = "", "WO", LstCpa.SelectedItem.SubItems(7))
            .txtbalance.Value = IIf(LstCpa.SelectedItem.SubItems(19) = "", "0", LstCpa.SelectedItem.SubItems(19))
            .txtarrangement.Text = LstCpa.SelectedItem.SubItems(6)
            .txtfrombalancepersen.Text = LstCpa.SelectedItem.SubItems(13)
            .txtpersenprincipal.Text = LstCpa.SelectedItem.SubItems(14)
            .dtpropsal.Value = Format(LstCpa.SelectedItem.SubItems(3), "dd/mm/yyyy")
            .dtpelunasan = Format(LstCpa.SelectedItem.SubItems(21), "dd/mm/yyyy")
            .txtoccupation.Text = LstCpa.SelectedItem.SubItems(23)
            .txtreason.Text = LstCpa.SelectedItem.SubItems(24)
            .txtnodlq.Text = LstCpa.SelectedItem.SubItems(25)
            .txtpaymenthandle.Text = LstCpa.SelectedItem.SubItems(26)
            .txtjust.Text = LstCpa.SelectedItem.SubItems(27)
            .tdbisnstallment.Value = IIf(LstCpa.SelectedItem.SubItems(31) = "", "0", LstCpa.SelectedItem.SubItems(31))
            '@@ 11-10-2011, Tambahan ID CPA
            TxtIdCpa.Text = LstCpa.SelectedItem.Text
            CmdJadwalPembayaran.Enabled = True
    
        End With
        Call Cari_LPD_LPA_Payment
    End If
        
    CMDSQL = "select * from mandiri.mgm where custid='" + Trim(LstCpa.SelectedItem.SubItems(1)) + "'"
       
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

    If M_objrs("acc_type") = "PIL" Then
      TxtCustidMMU.Text = IIf(IsNull(M_objrs("custidmmu")), "", M_objrs("custidmmu"))
    End If

    Set M_objrs = Nothing
End Sub
Private Sub SSCommand1_Click(Index As Integer)
Dim STRSQL As String
Dim rsTemp1 As New ADODB.Recordset
Dim rsTemporary As New ADODB.Recordset
Dim rsfound As New ADODB.Recordset
Dim strFaxed As String
Dim strOthers As String
Dim strwentalk As String
Dim strKTP As String
Dim strSup As String
Dim strBilling As String

Select Case Index
        Case 0
            Select Case SSCommand1(0).tag
            Case "1"
                 strFaxed = ""
                 strOthers = ""
                 strwentalk = ""
                 strKTP = ""
                 strSup = ""
                 strBilling = ""

                   If txtcardno.Text = "" Then
                        MsgBox "Anda belum klik tombol [ADD/Edit klik Di grid]"
                        Exit Sub
                   End If
                   
                    STRSQL = "select max(date(dtglinsert))  as tgl from mandiri.tblcpa where vcustid='" + txtcardno.Text + "' group by vcustid "
                    Set rsfound = New ADODB.Recordset
                    rsfound.CursorLocation = adUseClient
                    rsfound.Open STRSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
                    If Not rsfound.EOF Then
                    tglinsert = Format(IIf(IsNull(rsfound("tgl")), "", rsfound("tgl")), "dd/mm/yyyy")
                    End If
                    
'                    If Format(MDIForm1.TDBDate1.Value, "dd/mm/yyyy") = tglinsert Then
'                            MsgBox "Anda Sudah Pernah Create CPa Sebelum nya mohon hapus dulu"
'                            Debug.Print Format(MDIForm1.TDBDate1.Value, "dd/mm/yyyy")
'                            Exit Sub
'                    End If
                    Set rsfound = Nothing
                    
                    StatusChekcBox = ""
                    
                    If chkfaxed.Value = vbChecked Then
                        strFaxed = "1"
                    Else
                        strFaxed = "0"
                    End If
                    
                    If chkwentalk.Value = vbChecked Then
                        strwentalk = "1"
                    Else
                        strwentalk = "0"
                    End If
                    
                    If chkKTP.Value = vbChecked Then
                        
                        '@@ 08092011, Buat nyatet Option Yang dipilih
                        If StatusChekcBox = "" Then
                            StatusChekcBox = "KTP "
                        Else
                            StatusChekcBox = StatusChekcBox + ",KTP "
                        End If
                        
                        strKTP = "1"
                    Else
                        strKTP = "0"
                    End If
                   
'                   If chkKTP.Value = vbChecked Then
'                        strKTP = "1"
'                    Else
'                        strKTP = "0"
'                    End If
                    
                    If chkpp.Value = vbChecked Then
                        '@@ 08092011, Buat nyatet Option Yang dipilih
                        If StatusChekcBox = "" Then
                            StatusChekcBox = "Surper "
                        Else
                            StatusChekcBox = StatusChekcBox + ",Surper "
                        End If
                        strSup = "1"
                    Else
                        strSup = "0"
                    End If
                    
                    If chkbillings.Value = vbChecked Then
                        '@@ 08092011, Buat nyatet Option Yang dipilih
                        If StatusChekcBox = "" Then
                            StatusChekcBox = "Billing "
                        Else
                            StatusChekcBox = StatusChekcBox + ",Billing"
                        End If
                        strBilling = "1"
                    Else
                        strBilling = "0"
                    End If
                    
                    If Check1.Value = vbChecked Then
                        '@@ 08092011, Buat nyatet Option Yang dipilih
                        If StatusChekcBox = "" Then
                            StatusChekcBox = "Other "
                        Else
                            StatusChekcBox = StatusChekcBox + ",Other"
                        End If
                        strOthers = "1"
                    Else
                        strOthers = "0"
                    End If
                    
                    If StatusChekcBox = "" Then
                        MsgBox "Anda belum memilih salah satu/beberapa dokumen seperti KTP,Surper,Billing atau Other! Data gagal disimpan!", vbOKOnly + vbCritical, "Peringatan"
                        Exit Sub
                    End If
                    
                    Call Cari_LPD_LPA_Payment
                    
                    STRSQL = "insert into mandiri.tblCpa(vcustid,vregion,dpropsal,vreffno,vproduct,varragement,vcardsts,nttlpayment,ndownpay,nfuturepay,ncharge,"
                    STRSQL = STRSQL + " ndiscountamt,vosbalance,vosprincipal ,dtglinsert,"
                    STRSQL = STRSQL + " dtgllastupdate,dtglpelunasan,vjust,vcustname,"
                    STRSQL = STRSQL + " voccupation,vreason,vnodlq,vpaymenthandle,"
                    STRSQL = STRSQL + " agency,nbalance,nprincipal,nperiod,chkfaxed,"
                    STRSQL = STRSQL + " chkwentalking,chkktp,chksup,chkbillings,chkothers,"
                    STRSQL = STRSQL + " ketother,lpd_from_payment,lpa_from_payment,vcustid_mmu) values ( "
                    STRSQL = STRSQL + "'" + FrmCC_Colection.lblcustid.Caption + "','" + txtregion.Text + "',"
                    STRSQL = STRSQL + IIf(dtpropsal.ValueIsNull, "null", "'" + Format(dtpropsal.Value, "yyyy-mm-dd") + "'") + ", '" + txtreff.Text + "','" + txtproduct.Text + "' ,"
                    STRSQL = STRSQL + "'" + txtarrangement.Text + "','" + cbosts.Text + "'," + CStr(lblLastPay.Value) + "," + CStr(txtdownpayment.Value) + ","
                    STRSQL = STRSQL + "" + CStr(Val(txtfuture.Value)) + "," + CStr((txtcharge.Value)) + "," + CStr(txtdiscount.Value) + ",'" + txtfrombalancepersen.Text + "','" + txtpersenprincipal.Text + "',"
                    STRSQL = STRSQL + "'" + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") + " " + Format(Now, "hh:mm:ss") + "','" + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") + " " + Format(Now, "hh:mm:ss") + "',"
                    STRSQL = STRSQL + IIf(dtpelunasan.ValueIsNull, "null", " '" + Format(dtpelunasan.Value, "yyyy-mm-dd") + "'") + ",'" + txtjust.Text + "','" + FrmCC_Colection.lblnama.Caption + "',"
                    STRSQL = STRSQL + "'" + txtoccupation.Text + "', '" + txtreason.Text + "','" + txtnodlq.Text + "','" + txtpaymenthandle.Text + "',"
                    STRSQL = STRSQL + "'" + txtagency.Text + "',"
                    STRSQL = STRSQL + "" + CStr(Val(txtbalance.Value)) + "," + CStr((txtprincipal.Value)) + ","
                    STRSQL = STRSQL + "" + CStr(tdbisnstallment.Value) + ",'" + strFaxed + "','" + strwentalk + "','" + strKTP + "','" + strSup + "','" + strBilling + "','" + strOthers + "','" + txtothers.Text + "',"
                    STRSQL = STRSQL + LpdPayment + ",'"
                    STRSQL = STRSQL + CStr(TxtLPAPayment.Value) + "','"
                    STRSQL = STRSQL + IIf(IsNull(TxtCustidMMU.Text), "", Trim(TxtCustidMMU.Text)) + "')"
                    M_OBJCONN.Execute (STRSQL)
                    STRSQL = "update mandiri.mgm set stscpa=1, tglinsertfrmcpa ='" + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") + " " + Format(Now, "hh:mm:ss") + "' ,tglupdatefromcpa='" + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") + " " + Format(Now, "hh:mm:ss") + "'"
                    
                    If UCase(MDIForm1.txtlevel.Text) = "SUPERVISOR" Or UCase(MDIForm1.txtlevel.Text) = "ADMIN" Or UCase(MDIForm1.txtlevel.Text) = "ADMINISTRATOR" Then
                        STRSQL = STRSQL + " ,vnameapprovel ='" + MDIForm1.TxtUsername.Text + "' "
                    End If
                    STRSQL = STRSQL + " where custid ='" + FrmCC_Colection.lblcustid.Caption + "'"
                    M_OBJCONN.Execute (STRSQL)
                    
                     '@@ 07092011, jika fax atau when talking sulun di ceklist
                   If chkfaxed.Value = vbChecked Or chkwentalk.Value = vbChecked Then
                    Dim CMDSQL As String
                    Dim Remarks  As String
                    
                    With FrmCC_Colection
                        If chkfaxed.Value = vbChecked Then
                            Remarks = "CH/Payment Handle akan Fax. dokumen ke Rit Team "
                            Remarks = Remarks + "(" + StatusChekcBox + ")"
                            
                            CMDSQL = "insert into mandiri.mgm_hst (custid, agent, products, "
                            CMDSQL = CMDSQL + "hst,user_log) values ('"
                            CMDSQL = CMDSQL + .lblcustid.Caption + "','"
                            CMDSQL = CMDSQL + .lblaoc.Caption + "','"
                            CMDSQL = CMDSQL + "Collection" + "','"
                            CMDSQL = CMDSQL + Remarks + "','"
                            CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "')"
                            
                            M_OBJCONN.Execute CMDSQL
                        End If
                        
                        If chkwentalk.Value = vbChecked Then
                            Remarks = "CH/Payment Handle akan membawa dokumen sesuai perjanjian ke cabang HSBC"
                            Remarks = Remarks + " (" + StatusChekcBox + ")"
                            CMDSQL = "insert into mandiri.mgm_hst (custid, agent, products, "
                            CMDSQL = CMDSQL + "hst,user_log) values ('"
                            CMDSQL = CMDSQL + .lblcustid.Caption + "','"
                            CMDSQL = CMDSQL + .lblaoc.Caption + "','"
                            CMDSQL = CMDSQL + "Collection" + "','"
                            CMDSQL = CMDSQL + Remarks + "','"
                            CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "')"
                            
                            M_OBJCONN.Execute CMDSQL
                        End If
                    End With
                     
                    
                   End If
                    
                    MsgBox "data sudah disimpan", vbExclamation + vbOKOnly, "Pesan"
                    StatusChekcBox = ""
                   SSTab1.tag = 0
                                       
                  
                   
                                       
               'SSCommand1_Click (1)
               CLEAR
            Case "2"
                    If txtcardno.Text = "" Then
                        MsgBox "Anda belum klik tombol [ADD/Edit klik Di grid]"
                        Exit Sub
                    End If
                   
                   strFaxed = ""
                 strOthers = ""
                 strwentalk = ""
                 strKTP = ""
                 strSup = ""
                 strBilling = ""
                 
                    StatusChekcBox = ""
                    
                    If chkfaxed.Value = vbChecked Then
                        strFaxed = "1"
                    Else
                        strFaxed = "0"
                    End If
                    
                    If chkwentalk.Value = vbChecked Then
                        strwentalk = "1"
                    Else
                        strwentalk = "0"
                    End If
                    
                    If chkKTP.Value = vbChecked Then
                         '@@ 08092011, Buat nyatet Option Yang dipilih
                        If StatusChekcBox = "" Then
                            StatusChekcBox = "KTP "
                        Else
                            StatusChekcBox = StatusChekcBox + ",KTP "
                        End If
                        strKTP = "1"
                    Else
                        strKTP = "0"
                    End If
                   
'                   If chkKTP.Value = vbChecked Then
'
'                        strKTP = "1"
'                    Else
'                        strKTP = "0"
'                    End If
                    
                    If chkpp.Value = vbChecked Then
                         '@@ 08092011, Buat nyatet Option Yang dipilih
                        If StatusChekcBox = "" Then
                            StatusChekcBox = "Surper "
                        Else
                            StatusChekcBox = StatusChekcBox + ",Surper "
                        End If
                        strSup = "1"
                    Else
                        strSup = "0"
                    End If
                    
                    If chkbillings.Value = vbChecked Then
                        
                         '@@ 08092011, Buat nyatet Option Yang dipilih
                        If StatusChekcBox = "" Then
                            StatusChekcBox = "Billing "
                        Else
                            StatusChekcBox = StatusChekcBox + ",Billing "
                        End If
                        
                        strBilling = "1"
                    Else
                        strBilling = "0"
                    End If
                    
                    If Check1.Value = vbChecked Then
                        
                         '@@ 08092011, Buat nyatet Option Yang dipilih
                        If StatusChekcBox = "" Then
                            StatusChekcBox = "Other "
                        Else
                            StatusChekcBox = StatusChekcBox + ",Other "
                        End If
                        
                        strOthers = "1"
                    Else
                        strOthers = "0"
                    End If
                    
                    If StatusChekcBox = "" Then
                        MsgBox "Anda belum memilih salah satu/beberapa dokumen seperti KTP,Surper,Billing atau Other! Data gagal disimpan!", vbOKOnly + vbCritical, "Peringatan"
                        Exit Sub
                    End If
                    
                    Call Cari_LPD_LPA_Payment
                    
                    STRSQL = "update mandiri.tblcpa set  dtgllastupdate= '" + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") + " " + Format(Now, "hh:mm:ss") + "' ,nttlpayment='" + CStr(lblLastPay.Value) + "',ndownpay='" + CStr(txtdownpayment.Value) + "',"
                    STRSQL = STRSQL + "vregion='" + txtregion.Text + "',dpropsal=" + IIf(dtpropsal.ValueIsNull, "null", " '" + Format(dtpropsal.Value, "yyyy-mm-dd") + "'") + ",vreffno='" + txtreff.Text + "',vproduct ='" + txtproduct.Text + "',"
                    STRSQL = STRSQL + "varragement='" + txtarrangement.Text + "', vcardsts='" + cbosts.Text + "',nfuturepay =" + CStr(Val(txtfuture.Value)) + ",ncharge='" + CStr(Val(txtcharge.Value)) + "',"
                    STRSQL = STRSQL + " ndiscountamt=" + CStr(txtdiscount.Value) + ",vosbalance='" + txtfrombalancepersen.Text + "',vosprincipal='" + txtpersenprincipal.Text + "',"
                    STRSQL = STRSQL + "dtglpelunasan=" + IIf(dtpelunasan.ValueIsNull, "null", " '" + Format(dtpelunasan.Value, "yyyy-mm-dd") + "'") + ",vjust='" + txtjust.Text + "', agency='" + txtagency.Text + "',"
                    STRSQL = STRSQL + "voccupation='" + txtoccupation.Text + "',vreason='" + txtreason.Text + "',vnodlq='" + txtnodlq.Text + "',vpaymenthandle='" + txtpaymenthandle.Text + "',"
                    STRSQL = STRSQL + "nperiod=" + CStr(tdbisnstallment.Value) + ", nbalance=" + CStr(txtbalance.Value) + ",nprincipal=" + CStr(txtprincipal.Value) + ",chkfaxed= '" + strFaxed + "',chkwentalking= '" + strwentalk + "',chkktp= '" + strKTP + "',chksup= '" + strSup + "',chkbillings= '" + strBilling + "',chkothers= '" + strOthers + "',ketother='" + txtothers.Text + "',lpd_from_payment="
                    STRSQL = STRSQL + LpdPayment + ",lpa_from_payment='"
                    STRSQL = STRSQL + CStr(TxtLPAPayment.Value) + "',vcustid_mmu='"
                    STRSQL = STRSQL + IIf(IsNull(TxtCustidMMU.Text), "", Trim(TxtCustidMMU.Text)) + "' "
                    STRSQL = STRSQL + " where nid='" + LstCpa.SelectedItem.Text + "'"
                    M_OBJCONN.Execute (STRSQL)
                    
                    
                      '@@ 07092011, jika fax atau when talking sulun di ceklist
                   If chkfaxed.Value = vbChecked Or chkwentalk.Value = vbChecked Then
                    'Dim cmdsql As String
                    'Dim remarks  As String
                    
                    'With FrmCC_Colection
                        If chkfaxed.Value = vbChecked Then
                            Remarks = "CH/Payment Handle akan Fax. dokumen ke Rit Team "
                            Remarks = Remarks + "(" + StatusChekcBox + ")"
                            
                            CMDSQL = "insert into mandiri.mgm_hst (custid, agent, products, "
                            CMDSQL = CMDSQL + "hst,user_log) values ('"
                            CMDSQL = CMDSQL + LstCpa.SelectedItem.SubItems(1) + "','"
                            CMDSQL = CMDSQL + txtcollect.Text + "','"
                            CMDSQL = CMDSQL + "Collection" + "','"
                            CMDSQL = CMDSQL + Remarks + "','"
                            CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "')"
                            
                            M_OBJCONN.Execute CMDSQL
                        End If
                        
                        If chkwentalk.Value = vbChecked Then
                            Remarks = "CH/Payment Handle akan membawa dokumen sesuai perjanjian ke cabang HSBC"
                            Remarks = Remarks + " (" + StatusChekcBox + ")"
                            
                            CMDSQL = "insert into mandiri.mgm_hst (custid, agent, products, "
                            CMDSQL = CMDSQL + "hst,user_log) values ('"
                            CMDSQL = CMDSQL + LstCpa.SelectedItem.SubItems(1) + "','"
                            CMDSQL = CMDSQL + txtcollect.Text + "','"
                            CMDSQL = CMDSQL + "Collection" + "','"
                            CMDSQL = CMDSQL + Remarks + "','"
                            CMDSQL = CMDSQL + MDIForm1.TxtUsername.Text + "')"
                            
                            M_OBJCONN.Execute CMDSQL
                        End If
                    'End With
                     
                    
                   End If
                    
                    StatusChekcBox = ""
                    
                    MsgBox "data telah di update", vbInformation + vbOKOnly, "Pesan"
                    STRSQL = "update mandiri.mgm set stscpa=1, tglupdatefromcpa='" + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") + " " + Format(Now, "hh:mm:ss") + "' where custid ='" + LstCpa.SelectedItem.SubItems(1) + "'"
                    CLEAR
                    SSTab1.tag = 0
                    
        End Select
      Case 1
        showlist
        Unload Me
      Case 2
     'RPT.Reset
            SSCommand1_Click (0)
           M_RPTCONN.Execute "delete from tblreportcpa "
           STRSQL = "select * from tblreportcpa"
           Set rsTemp1 = New ADODB.Recordset
           rsTemp1.CursorLocation = adUseClient
           rsTemp1.Open STRSQL, M_RPTCONN, adOpenDynamic, adLockOptimistic
           'cmdsql = "SELECT  * FROM ( "
           'cmdsql = cmdsql + " SELECT * FROM TBLCPA WHERE DTGLINSERT  IN (SELECT MAX(DTGLINSERT) FROM TBLCPA  GROUP BY VCUSTID)) AS A"
           'cmdsql = cmdsql + " Inner Join (SELECT * FROM   MGM) AS B  ON A.VCUSTID=B.CUSTID WHERE VCUSTID='" + txtcardno.Text + "'"
           
           
           
          CMDSQL = "  SELECT * FROM ( "
          CMDSQL = CMDSQL + " SELECT * FROM (SELECT *  FROM mandiri.USERTBL) AS B"
          CMDSQL = CMDSQL + " Right JOIN  ( "
          CMDSQL = CMDSQL + " SELECT * FROM ( "
          CMDSQL = CMDSQL + " SELECT  * FROM (  SELECT * FROM mandiri.TBLCPA WHERE VCUSTID='" + LstCpa.SelectedItem.SubItems(1) + "'   ) AS A Inner Join "
          CMDSQL = CMDSQL + "  (SELECT * FROM   mandiri.MGM) AS B  ON A.VCUSTID=B.CUSTID  ) as c)  AS BRU ON BRU.AGENT=B.USERID) AS TBLBARU"
          CMDSQL = CMDSQL + " Left Join ( "
          CMDSQL = CMDSQL + "   select * from ( "
          CMDSQL = CMDSQL + " SELECT custid as cust_no,PAYDATE AS lpd,payment as lpa FROM mandiri.TBLLUNAS  WHERE ID IN (SELECT MAX(ID) FROM mandiri.tbllunas GROUP BY CUSTID))  as tblbaru1 WHERE cust_no='" + LstCpa.SelectedItem.SubItems(1) + "' ) as bru on tblbaru.custid=bru.cust_no "


          '  CMDSQL = " SELECT * FROM (SELECT *  FROM USERTBL) AS B"
          ' CMDSQL = CMDSQL + " Right JOIN  ("
          ' CMDSQL = CMDSQL + " SELECT * FROM ("
          ' CMDSQL = CMDSQL + " SELECT  * FROM (  SELECT * FROM TBLCPA WHERE VCUSTID='" + FrmCC_Colection.lblCustId.Caption + "' ) AS A Inner Join"
          ' CMDSQL = CMDSQL + " (SELECT * FROM   MGM) AS B  ON A.VCUSTID=B.CUSTID  ))  AS BRU ON BRU.AGENT=B.USERID"
           
          ' cmdsql = "SELECT  * FROM ( "
          ' cmdsql = cmdsql + " SELECT * FROM TBLCPA WHERE VCUSTID='" + FrmCC_Colection.lblCustId.Caption + "' ) AS A"
          ' cmdsql = cmdsql + " Inner Join (SELECT * FROM   MGM) AS B  ON A.VCUSTID=B.CUSTID "


           Set rsTemporary = New ADODB.Recordset
           rsTemporary.CursorLocation = adUseClient
          
           rsTemporary.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
           
           While Not rsTemporary.EOF

            rsTemp1.AddNew
            rsTemp1("dtglinsert") = IIf(IsNull(rsTemporary("dtglinsert")), "", rsTemporary("dtglinsert"))
            rsTemp1("vregion") = IIf(IsNull(rsTemporary("vregion")), "", rsTemporary("vregion"))
            rsTemp1("dproposal") = IIf(IsNull(rsTemporary("dpropsal")), Null, rsTemporary("dpropsal"))
            rsTemp1("vreffno") = IIf(IsNull(rsTemporary("vreffno")), "", rsTemporary("vreffno"))
            rsTemp1("product") = IIf(IsNull(rsTemporary("vproduct")), "", rsTemporary("vproduct"))
            rsTemp1("arrangement") = IIf(IsNull(rsTemporary("varragement")), "", rsTemporary("varragement"))
            
            '@@ 15-09-2011, jika custid mmu ada isinya, maka cardno diisi sesuai dengan custidmmu
            If rsTemporary("vcustid_mmu") = "" Then
                rsTemp1("cardno") = IIf(IsNull(rsTemporary("vcustid")), "", rsTemporary("vcustid"))
            Else
                rsTemp1("cardno") = IIf(IsNull(rsTemporary("vcustid_mmu")), "", rsTemporary("vcustid_mmu"))
            End If
            
            rsTemp1("custname") = IIf(IsNull(rsTemporary("name")), "", rsTemporary("name"))
            rsTemp1("cardopen") = IIf(IsNull(rsTemporary("opendate")), Null, rsTemporary("opendate"))
            rsTemp1("agent") = IIf(IsNull(rsTemporary("agent")), "", rsTemporary("agent"))
            rsTemp1("outbalance") = IIf(IsNull(rsTemporary("nbalance")), 0, rsTemporary("nbalance"))
            rsTemp1("ttlpayment") = IIf(IsNull(rsTemporary("nttlpayment")), 0, rsTemporary("nttlpayment"))
            rsTemp1("downpayment") = IIf(IsNull(rsTemporary("ndownpay")), 0, rsTemporary("ndownpay"))
            rsTemp1("futurepayment") = IIf(IsNull(rsTemporary("nfuturepay")), "", rsTemporary("nfuturepay"))
            rsTemp1("nprincipal") = IIf(IsNull(rsTemporary("nprincipal")), 0, rsTemporary("nprincipal"))
            rsTemp1("ncharge") = IIf(IsNull(rsTemporary("ncharge")), 0, rsTemporary("ncharge"))
            rsTemp1("ndiskon") = IIf(IsNull(rsTemporary("ndiscountamt")), 0, rsTemporary("ndiscountamt"))
            rsTemp1("osfrombalance") = IIf(IsNull(rsTemporary("vosbalance")), "", rsTemporary("vosbalance"))
            rsTemp1("osfromprincipal") = IIf(IsNull(rsTemporary("vosprincipal")), "", rsTemporary("vosprincipal"))
            rsTemp1("custid") = IIf(IsNull(rsTemporary("vcustid")), "", rsTemporary("vcustid"))
            rsTemp1("approve") = IIf(IsNull(rsTemporary("vnameapprovel")), "", rsTemporary("vnameapprovel"))
            rsTemp1("vjust") = IIf(IsNull(rsTemporary("vjust")), "", rsTemporary("vjust"))
            rsTemp1("agency") = IIf(IsNull(rsTemporary("agency")), "", rsTemporary("agency"))
            rsTemp1("vnameverify") = IIf(IsNull(rsTemporary("vnameverify")), "", rsTemporary("vnameverify"))
            rsTemp1("vreason") = IIf(IsNull(rsTemporary("vreason")), "", rsTemporary("vreason"))
            rsTemp1("vdlq") = IIf(IsNull(rsTemporary("vnodlq")), "", rsTemporary("vnodlq"))
            rsTemp1("vpaymenthandle") = IIf(IsNull(rsTemporary("vpaymenthandle")), "", rsTemporary("vpaymenthandle"))
            rsTemp1("voccupation") = IIf(IsNull(rsTemporary("voccupation")), "", rsTemporary("voccupation"))
            rsTemp1("nperiod") = IIf(IsNull(rsTemporary("nperiod")), 0, rsTemporary("nperiod"))
            rsTemp1("dtglapprove") = IIf(IsNull(rsTemporary("tglapprove")), Null, rsTemporary("tglapprove"))
            rsTemp1("userid") = IIf(IsNull(rsTemporary("userid")), "", rsTemporary("userid"))
            rsTemp1("team") = IIf(IsNull(rsTemporary("team")), "", rsTemporary("team"))
            rsTemp1("chkfaxed") = IIf(IsNull(rsTemporary("chkfaxed")), "", rsTemporary("chkfaxed"))
            rsTemp1("chkwentalking") = IIf(IsNull(rsTemporary("chkwentalking")), "", rsTemporary("chkwentalking"))
            rsTemp1("chkktp") = IIf(IsNull(rsTemporary("chkktp")), "", rsTemporary("chkktp"))
            rsTemp1("chksup") = IIf(IsNull(rsTemporary("chksup")), "", rsTemporary("chksup"))
            rsTemp1("chkbillings") = IIf(IsNull(rsTemporary("chkbillings")), "", rsTemporary("chkbillings"))
            rsTemp1("chkothers") = IIf(IsNull(rsTemporary("chkothers")), "", rsTemporary("chkothers"))
            rsTemp1("ketother") = IIf(IsNull(rsTemporary("ketother")), "", rsTemporary("ketother"))
            rsTemp1("ed") = IIf(IsNull(rsTemporary("tglsource")), Null, rsTemporary("tglsource"))
            rsTemp1("wo_date") = IIf(IsNull(rsTemporary("b_d")), Null, rsTemporary("b_d"))
            rsTemp1("odt") = IIf(IsNull(rsTemporary("OpenDate")), Null, rsTemporary("OpenDate"))
            If IIf(IsNull(rsTemporary("lpd")), "", rsTemporary("lpd")) = "" Then
                  rsTemp1("lpd") = IIf(IsNull(rsTemporary("pay_dt")), Null, rsTemporary("pay_dt"))
            Else
                  '@@27-07-2011 Dinonaktifkan, LPD diambil dari ldp MGM bukan dari tbllunas
                  'rsTemp1("lpd") = IIf(IsNull(rsTemporary("lpd")), Null, rsTemporary("lpd"))
            End If
            
            If IIf(IsNull(rsTemporary("lpa")), "", rsTemporary("lpa")) = "" Then
                  rsTemp1("lpa") = IIf(IsNull(rsTemporary("LastPay")), 0, rsTemporary("LastPay"))
            Else
                  '@@27-07-2011 Dinonaktifkan, LPA diambil dari lpa MGM bukan dari tbllunas
                  'rsTemp1("lpa") = IIf(IsNull(rsTemporary("lpa")), 0, rsTemporary("lpa"))
            End If
            
            rsTemp1("lpd_from_payment") = IIf(IsNull(rsTemporary("lpd_from_payment")), Null, Format(rsTemporary("lpd_from_payment"), "yyyy-mm-dd"))
            rsTemp1("lpa_from_payment") = IIf(IsNull(rsTemporary("lpa_from_payment")), 0, rsTemporary("lpa_from_payment"))

            
           
            rsTemp1.update
           
                    rsTemporary.MoveNext
           Wend
           
          
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptCpaRincian.rpt"
            WaitSecs (2)
            Call SHOW_PRN
            Set rsTemp1 = Nothing
            Set rsTemporary = Nothing
      
End Select

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.tag = 1 Then
Exit Sub
End If
If SSTab1.Tab = 0 Then
    showlist
End If

End Sub
Private Sub txtbalance_Change()
    txtcharge.Value = txtbalance.Value - txtprincipal.Value
    txtdiscount.Value = txtbalance.Value - lblLastPay.Value
    If txtbalance.Value <> 0 Then
        If lblLastPay.Value < txtbalance.Value Then
            txtfrombalancepersen.Text = "-" + CStr(Round((txtdiscount.Value / txtbalance.Value) * 100, 2))
        Else
            txtfrombalancepersen.Text = Round((txtdiscount.Value / txtbalance.Value) * 100, 2)
        End If
    End If

End Sub

Private Sub txtdiscount_Change()
If txtbalance.Value <> 0 Then
    If lblLastPay.Value < txtbalance.Value Then
        txtfrombalancepersen.Text = "-" + CStr(Round((txtdiscount.Value / txtbalance.Value) * 100, 2))
    Else
        txtfrombalancepersen.Text = Round((txtdiscount.Value / txtbalance.Value) * 100, 2)
    End If
End If

End Sub

Private Sub txtdownpayment_Change()
    txtfuture.Value = lblLastPay.Value - txtdownpayment.Value
End Sub

Private Sub txtprincipal_Change()
    txtcharge.Value = txtbalance.Value - txtprincipal.Value
    If txtprincipal.Value <> 0 Then
      If lblLastPay.Value < txtprincipal.Value Then
         txtpersenprincipal.Text = "-" + CStr(Round(((lblLastPay.Value - txtprincipal.Value) / txtprincipal.Value) * 100, 2))
      Else
         txtpersenprincipal.Text = Round(((lblLastPay.Value - txtprincipal.Value) / txtprincipal.Value) * 100, 2)
    End If

End If
End Sub

Private Sub txtreff_Change()
    Select Case UCase(txtreff.Text)
            Case "D"
                 txtarrangement.Text = "SETTLEMENT"
            Case "R"
                txtarrangement.Text = "RESCHEDULE"
            Case "X"
                txtarrangement.Text = "PAID-OFF"
    End Select


End Sub
Private Sub lblLastPay_Change()
  txtdiscount.Value = txtbalance.Value - lblLastPay.Value
If txtprincipal.Value <> 0 Then
If lblLastPay.Value < txtprincipal.Value Then
        txtpersenprincipal.Text = "-" + CStr(Round(((lblLastPay.Value - txtprincipal.Value) / txtprincipal.Value) * 100, 2))
    Else
        txtpersenprincipal.Text = Round(((lblLastPay.Value - txtprincipal.Value) / txtprincipal.Value) * 100, 2)
    End If

End If

End Sub
Private Sub SHOW_PRN()
    RPT.RetrieveDataFiles
    RPT.WindowLeft = 0
    RPT.WindowTop = 0
    RPT.WindowState = crptMaximized
    RPT.WindowShowPrintBtn = True
    RPT.WindowShowRefreshBtn = True
    RPT.WindowShowSearchBtn = True
    RPT.WindowShowPrintSetupBtn = True
    RPT.WindowControls = True
    RPT.PrintReport
    'RPT.Action = 1
    'RPT.Reset
End Sub


Public Sub CLEAR()
txtregion.Text = ""
dtpropsal.Text = ""
txtproduct.Text = ""
txtcardno.Text = ""
txtname.Text = ""
dtcardopen.Text = ""
dwo.Text = ""
txtcycle.Text = ""
txtcollect.Text = ""
txtplace.Text = ""
txtagency.Text = ""
txtbalance.Value = 0
lblLastPay.Value = 0
txtdownpayment.Value = 0
txtfuture.Value = 0
txtperiodpay.Text = ""
txtprincipal.Value = 0
tdbisnstallment.Value = 0
Label5.Text = ""
Label8.Text = ""
txtcharge.Value = 0
txtdiscount.Value = 0
txtfrombalancepersen.Text = ""
txtpersenprincipal.Text = ""
txtoccupation.Text = ""
txtreason.Text = ""
txtnodlq.Text = ""
txtpaymenthandle.Text = ""
txtjust.Text = ""
dtpelunasan.Text = ""
txtreff.Text = ""
chkfaxed.Value = vbUnchecked
chkwentalk.Value = vbUnchecked
chkbillings.Value = vbUnchecked
chkKTP.Value = vbUnchecked
chkpp.Value = vbUnchecked
Check1.Value = vbUnchecked


txtothers.Text = ""
End Sub

'@@ 16-03-2011, Ini buat nyari LPD dan LPA terakhir dari tabel lunas
Private Sub Cari_LPD_LPA_Payment()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    
    CMDSQL = "select paydate,payment from mandiri.tbllunas where custid='"
    CMDSQL = CMDSQL + Trim(LstCpa.SelectedItem.SubItems(1)) + "' order by paydate desc limit 1 "
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_objrs.RecordCount > 0 Then
            TxtLPDPayment.Text = IIf(IsNull(M_objrs("paydate")), "", Format(M_objrs("paydate"), "yyyy-mm-dd"))
            TxtLPAPayment.Value = IIf(IsNull(M_objrs("payment")), "0", M_objrs("payment"))
            LpdPayment = "'" + TxtLPDPayment.Text + "'"
        Else
            LpdPayment = "null"
            TxtLPDPayment = ""
            TxtLPAPayment.Value = "0"
        End If
    Set M_objrs = Nothing
End Sub

'@@ 08-12-2011, List nama yang bisa approve
Private Sub IsiNamaApprove()
    CmbApprove.CLEAR
    CmbApprove.AddItem MDIForm1.TxtUsername.Text
End Sub
