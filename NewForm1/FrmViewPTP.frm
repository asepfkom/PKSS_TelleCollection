VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmViewPTP 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form View CPA"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11910
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame13 
      Appearance      =   0  'Flat
      BackColor       =   &H00B1FDD5&
      BorderStyle     =   0  'None
      Caption         =   "Frame13"
      ForeColor       =   &H80000008&
      Height          =   8355
      Left            =   0
      TabIndex        =   23
      Top             =   450
      Width           =   6510
      Begin VB.ComboBox CmbJenisPTP 
         BackColor       =   &H0000FFFF&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmViewPTP.frx":0000
         Left            =   3780
         List            =   "FrmViewPTP.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   100
         Top             =   600
         Width           =   1995
      End
      Begin VB.TextBox TxtDob 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4620
         TabIndex        =   90
         Top             =   4020
         Width           =   1575
      End
      Begin VB.TextBox label8 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3780
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   7335
         Width           =   1995
      End
      Begin VB.TextBox label5 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3780
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   6570
         Width           =   1995
      End
      Begin VB.TextBox txtperiodpay 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   7650
         Width           =   2175
      End
      Begin VB.Frame Frame15 
         Appearance      =   0  'Flat
         BackColor       =   &H002F735C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   2
         Left            =   270
         TabIndex        =   42
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
            TabIndex        =   43
            Top             =   90
            Width           =   3255
         End
         Begin VB.Image Image1 
            Height          =   375
            Index           =   2
            Left            =   75
            Picture         =   "FrmViewPTP.frx":002D
            Stretch         =   -1  'True
            Top             =   40
            Width           =   375
         End
      End
      Begin VB.TextBox txtagency 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   5400
         Width           =   2220
      End
      Begin VB.TextBox txtplace 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   5040
         Width           =   2220
      End
      Begin VB.TextBox txtcollect 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   4680
         Width           =   2220
      End
      Begin VB.ComboBox cbosts 
         Height          =   315
         ItemData        =   "FrmViewPTP.frx":18C7
         Left            =   1530
         List            =   "FrmViewPTP.frx":18D7
         TabIndex        =   38
         Top             =   4320
         Width           =   2220
      End
      Begin VB.TextBox txtcycle 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   3960
         Width           =   1995
      End
      Begin VB.TextBox txtname 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   3285
         Width           =   4875
      End
      Begin VB.TextBox txtcardno 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   2970
         Width           =   1995
      End
      Begin VB.Frame Frame15 
         Appearance      =   0  'Flat
         BackColor       =   &H002F735C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   1
         Left            =   270
         TabIndex        =   33
         Top             =   2430
         Width           =   6180
         Begin VB.Image Image1 
            Height          =   375
            Index           =   1
            Left            =   75
            Picture         =   "FrmViewPTP.frx":18EB
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
            TabIndex        =   34
            Top             =   90
            Width           =   3255
         End
      End
      Begin VB.TextBox txtarrangement 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1485
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   1980
         Width           =   1455
      End
      Begin VB.TextBox txtproduct 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1485
         Locked          =   -1  'True
         TabIndex        =   31
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
         TabIndex        =   30
         Top             =   1260
         Width           =   1455
      End
      Begin VB.TextBox txtregion 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1485
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   585
         Width           =   1440
      End
      Begin VB.Frame Frame15 
         Appearance      =   0  'Flat
         BackColor       =   &H002F735C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   0
         Left            =   225
         TabIndex        =   27
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
            TabIndex        =   28
            Top             =   45
            Width           =   1455
         End
         Begin VB.Image Image1 
            Height          =   375
            Index           =   0
            Left            =   75
            Picture         =   "FrmViewPTP.frx":3185
            Stretch         =   -1  'True
            Top             =   40
            Width           =   375
         End
      End
      Begin VB.TextBox TxtLPDPayment 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   4680
         Width           =   1995
      End
      Begin VB.TextBox TxtCustidMMU 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         TabIndex        =   25
         Top             =   3000
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox TxtIdCpa 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   3000
         TabIndex        =   24
         Top             =   600
         Width           =   675
      End
      Begin TDBDate6Ctl.TDBDate dtpropsal 
         Height          =   255
         Left            =   1485
         TabIndex        =   47
         Top             =   945
         Width           =   1410
         _Version        =   65536
         _ExtentX        =   2487
         _ExtentY        =   450
         Calendar        =   "FrmViewPTP.frx":4A1F
         Caption         =   "FrmViewPTP.frx":4B37
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmViewPTP.frx":4BA3
         Keys            =   "FrmViewPTP.frx":4BC1
         Spin            =   "FrmViewPTP.frx":4C1F
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
      Begin TDBDate6Ctl.TDBDate dtcardopen 
         Height          =   255
         Left            =   1530
         TabIndex        =   48
         Top             =   3645
         Width           =   2250
         _Version        =   65536
         _ExtentX        =   3969
         _ExtentY        =   450
         Calendar        =   "FrmViewPTP.frx":4C47
         Caption         =   "FrmViewPTP.frx":4D5F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmViewPTP.frx":4DCB
         Keys            =   "FrmViewPTP.frx":4DE9
         Spin            =   "FrmViewPTP.frx":4E47
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   10147522
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
         TabIndex        =   49
         Top             =   6660
         Width           =   2160
         _Version        =   65536
         _ExtentX        =   3810
         _ExtentY        =   450
         Calculator      =   "FrmViewPTP.frx":4E6F
         Caption         =   "FrmViewPTP.frx":4E8F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmViewPTP.frx":4EFB
         Keys            =   "FrmViewPTP.frx":4F19
         Spin            =   "FrmViewPTP.frx":4F63
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
         TabIndex        =   50
         Top             =   7020
         Width           =   2160
         _Version        =   65536
         _ExtentX        =   3810
         _ExtentY        =   450
         Calculator      =   "FrmViewPTP.frx":4F8B
         Caption         =   "FrmViewPTP.frx":4FAB
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmViewPTP.frx":5017
         Keys            =   "FrmViewPTP.frx":5035
         Spin            =   "FrmViewPTP.frx":507F
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
      Begin TDBNumber6Ctl.TDBNumber txtfuture 
         Height          =   255
         Left            =   1530
         TabIndex        =   51
         Top             =   7335
         Width           =   2160
         _Version        =   65536
         _ExtentX        =   3810
         _ExtentY        =   450
         Calculator      =   "FrmViewPTP.frx":50A7
         Caption         =   "FrmViewPTP.frx":50C7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmViewPTP.frx":5133
         Keys            =   "FrmViewPTP.frx":5151
         Spin            =   "FrmViewPTP.frx":519B
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   10147522
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
         TabIndex        =   52
         Top             =   8040
         Width           =   2160
         _Version        =   65536
         _ExtentX        =   3810
         _ExtentY        =   450
         Calculator      =   "FrmViewPTP.frx":51C3
         Caption         =   "FrmViewPTP.frx":51E3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmViewPTP.frx":524F
         Keys            =   "FrmViewPTP.frx":526D
         Spin            =   "FrmViewPTP.frx":52B7
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
         TabIndex        =   53
         Top             =   6300
         Width           =   1800
         _Version        =   65536
         _ExtentX        =   3175
         _ExtentY        =   450
         Calculator      =   "FrmViewPTP.frx":52DF
         Caption         =   "FrmViewPTP.frx":52FF
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmViewPTP.frx":536B
         Keys            =   "FrmViewPTP.frx":5389
         Spin            =   "FrmViewPTP.frx":53D3
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
         Height          =   255
         Left            =   4620
         TabIndex        =   54
         Top             =   3600
         Width           =   1620
         _Version        =   65536
         _ExtentX        =   2857
         _ExtentY        =   450
         Calendar        =   "FrmViewPTP.frx":53FB
         Caption         =   "FrmViewPTP.frx":5513
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmViewPTP.frx":557F
         Keys            =   "FrmViewPTP.frx":559D
         Spin            =   "FrmViewPTP.frx":55FB
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
         TabIndex        =   55
         Top             =   7980
         Width           =   2160
         _Version        =   65536
         _ExtentX        =   3810
         _ExtentY        =   450
         Calculator      =   "FrmViewPTP.frx":5623
         Caption         =   "FrmViewPTP.frx":5643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmViewPTP.frx":56AF
         Keys            =   "FrmViewPTP.frx":56CD
         Spin            =   "FrmViewPTP.frx":5717
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   16777215
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999
         MinValue        =   -999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   1
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin TDBNumber6Ctl.TDBNumber TxtLPAPayment 
         Height          =   255
         Left            =   3840
         TabIndex        =   56
         Top             =   5280
         Width           =   2160
         _Version        =   65536
         _ExtentX        =   3810
         _ExtentY        =   450
         Calculator      =   "FrmViewPTP.frx":573F
         Caption         =   "FrmViewPTP.frx":575F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmViewPTP.frx":57CB
         Keys            =   "FrmViewPTP.frx":57E9
         Spin            =   "FrmViewPTP.frx":5833
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   10147522
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
      Begin VB.Label Label11 
         BackColor       =   &H00B1FDD5&
         Caption         =   "DOB:"
         Height          =   285
         Index           =   1
         Left            =   3840
         TabIndex        =   89
         Top             =   4080
         Width           =   795
      End
      Begin VB.Label Label7 
         Caption         =   "Principal di database"
         Height          =   285
         Index           =   0
         Left            =   3780
         TabIndex        =   82
         Top             =   6975
         Width           =   1590
      End
      Begin VB.Label Label6 
         Caption         =   "Balance di database"
         Height          =   285
         Left            =   3735
         TabIndex        =   81
         Top             =   6300
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00B1FDD5&
         Caption         =   ")*  D=SETTLEMENT R=RESCHEDULE X=PAID OFF"
         ForeColor       =   &H000000FF&
         Height          =   990
         Left            =   3060
         TabIndex        =   80
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Principal"
         Height          =   285
         Index           =   19
         Left            =   315
         TabIndex        =   79
         Top             =   8100
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment period month"
         Height          =   465
         Index           =   17
         Left            =   315
         TabIndex        =   78
         Top             =   7695
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Future Payment"
         Height          =   195
         Index           =   16
         Left            =   315
         TabIndex        =   77
         Top             =   7380
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Down Payment"
         Height          =   195
         Index           =   15
         Left            =   315
         TabIndex        =   76
         Top             =   7020
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Payment"
         Height          =   330
         Index           =   14
         Left            =   315
         TabIndex        =   75
         Top             =   6705
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         Height          =   285
         Index           =   13
         Left            =   360
         TabIndex        =   74
         Top             =   6345
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Agency name"
         Height          =   240
         Index           =   12
         Left            =   405
         TabIndex        =   73
         Top             =   5445
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "placement"
         Height          =   240
         Index           =   11
         Left            =   405
         TabIndex        =   72
         Top             =   5085
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "collector"
         Height          =   240
         Index           =   10
         Left            =   360
         TabIndex        =   71
         Top             =   4770
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "status card"
         Height          =   240
         Index           =   9
         Left            =   360
         TabIndex        =   70
         Top             =   4365
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cycle Dlq"
         Height          =   240
         Index           =   8
         Left            =   360
         TabIndex        =   69
         Top             =   4005
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Card Open"
         Height          =   240
         Index           =   7
         Left            =   360
         TabIndex        =   68
         Top             =   3645
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "cust name"
         Height          =   240
         Index           =   6
         Left            =   360
         TabIndex        =   67
         Top             =   3330
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Card no"
         Height          =   240
         Index           =   5
         Left            =   360
         TabIndex        =   66
         Top             =   3015
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Arrangement"
         Height          =   240
         Index           =   4
         Left            =   360
         TabIndex        =   65
         Top             =   2025
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Product"
         Height          =   240
         Index           =   3
         Left            =   360
         TabIndex        =   64
         Top             =   1665
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Reffno"
         Height          =   240
         Index           =   2
         Left            =   360
         TabIndex        =   63
         Top             =   1305
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Proposal Date"
         Height          =   240
         Index           =   1
         Left            =   315
         TabIndex        =   62
         Top             =   990
         Width           =   1140
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Region"
         Height          =   240
         Index           =   18
         Left            =   315
         TabIndex        =   61
         Top             =   630
         Width           =   1230
      End
      Begin VB.Label Label11 
         BackColor       =   &H00B1FDD5&
         Caption         =   "Wo date"
         Height          =   285
         Index           =   0
         Left            =   3840
         TabIndex        =   60
         Top             =   3630
         Width           =   795
      End
      Begin VB.Label Label12 
         Caption         =   "Installment Period"
         Height          =   375
         Left            =   3810
         TabIndex        =   59
         Top             =   7620
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "LPD from Payment:"
         Height          =   240
         Index           =   20
         Left            =   3900
         TabIndex        =   58
         Top             =   4380
         Width           =   1830
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "LPA from Payment:"
         Height          =   240
         Index           =   22
         Left            =   3840
         TabIndex        =   57
         Top             =   5040
         Width           =   1890
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00B1FDD5&
      BorderStyle     =   0  'None
      Caption         =   "Frame13"
      ForeColor       =   &H80000008&
      Height          =   8325
      Left            =   6540
      TabIndex        =   1
      Top             =   450
      Width           =   5385
      Begin VB.ComboBox CmbOccupation 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "FrmViewPTP.frx":585B
         Left            =   1500
         List            =   "FrmViewPTP.frx":5868
         TabIndex        =   96
         Top             =   3480
         Width           =   3015
      End
      Begin VB.ComboBox CmbReason 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "FrmViewPTP.frx":5878
         Left            =   1500
         List            =   "FrmViewPTP.frx":588E
         TabIndex        =   95
         Top             =   3840
         Width           =   3015
      End
      Begin VB.ComboBox CmbPaymentHandle 
         Height          =   315
         ItemData        =   "FrmViewPTP.frx":58D0
         Left            =   1500
         List            =   "FrmViewPTP.frx":58E0
         TabIndex        =   94
         Top             =   3060
         Width           =   3015
      End
      Begin VB.CommandButton CmdKeluar 
         Caption         =   "&Keluar"
         Height          =   495
         Left            =   4020
         TabIndex        =   83
         Top             =   7560
         Width           =   1335
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "&Save"
         Height          =   495
         Left            =   2700
         TabIndex        =   85
         Top             =   7560
         Width           =   1335
      End
      Begin VB.CommandButton CmdCetak 
         Caption         =   "&Print"
         Height          =   495
         Left            =   1380
         TabIndex        =   92
         Top             =   7560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton CmdGetJustification 
         Caption         =   "&Get Justification from remarks..."
         Height          =   315
         Left            =   1800
         TabIndex        =   88
         Top             =   4920
         Width           =   2895
      End
      Begin VB.CommandButton CmdApprove 
         Caption         =   "&Approve"
         Height          =   495
         Left            =   60
         TabIndex        =   84
         Top             =   7560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Frame Frame15 
         Appearance      =   0  'Flat
         BackColor       =   &H002F735C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   5
         Left            =   -45
         TabIndex        =   14
         Top             =   90
         Width           =   4785
         Begin VB.Image Image1 
            Height          =   375
            Index           =   5
            Left            =   75
            Picture         =   "FrmViewPTP.frx":58FF
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
            TabIndex        =   15
            Top             =   45
            Width           =   1455
         End
      End
      Begin VB.TextBox txtfrombalancepersen 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1845
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1260
         Width           =   2085
      End
      Begin VB.TextBox txtpersenprincipal 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1845
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1665
         Width           =   2085
      End
      Begin VB.Frame Frame15 
         Appearance      =   0  'Flat
         BackColor       =   &H002F735C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   4
         Left            =   0
         TabIndex        =   10
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
            TabIndex        =   11
            Top             =   90
            Width           =   3255
         End
         Begin VB.Image Image1 
            Height          =   375
            Index           =   4
            Left            =   75
            Picture         =   "FrmViewPTP.frx":7199
            Stretch         =   -1  'True
            Top             =   40
            Width           =   375
         End
      End
      Begin VB.TextBox txtjust 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   1530
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   4185
         Width           =   3165
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00B1FDD5&
         Caption         =   "DOC"
         Height          =   2355
         Left            =   150
         TabIndex        =   2
         Top             =   5160
         Width           =   4665
         Begin VB.TextBox txtothers 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Left            =   660
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   91
            Top             =   1500
            Width           =   3225
         End
         Begin VB.CheckBox chkfaxed 
            BackColor       =   &H00B1FDD5&
            Caption         =   "Faxed"
            Height          =   285
            Left            =   180
            TabIndex        =   8
            Top             =   180
            Width           =   1005
         End
         Begin VB.CheckBox chkwentalk 
            BackColor       =   &H00B1FDD5&
            Caption         =   "When Talking Surlun"
            Height          =   285
            Left            =   180
            TabIndex        =   7
            Top             =   420
            Width           =   1905
         End
         Begin VB.CheckBox chkKTP 
            BackColor       =   &H00B1FDD5&
            Caption         =   "KTP"
            Enabled         =   0   'False
            Height          =   165
            Left            =   360
            TabIndex        =   6
            Top             =   720
            Width           =   765
         End
         Begin VB.CheckBox chkpp 
            BackColor       =   &H00B1FDD5&
            Caption         =   "Surper"
            Enabled         =   0   'False
            Height          =   165
            Left            =   1140
            TabIndex        =   5
            Top             =   720
            Width           =   825
         End
         Begin VB.CheckBox chkbillings 
            BackColor       =   &H00B1FDD5&
            Caption         =   "Billings"
            Enabled         =   0   'False
            Height          =   285
            Left            =   360
            TabIndex        =   4
            Top             =   900
            Width           =   825
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00B1FDD5&
            Caption         =   "Others"
            Enabled         =   0   'False
            Height          =   225
            Left            =   360
            TabIndex        =   3
            Top             =   1260
            Width           =   795
         End
      End
      Begin TDBNumber6Ctl.TDBNumber txtcharge 
         Height          =   255
         Left            =   1845
         TabIndex        =   16
         Top             =   585
         Width           =   2160
         _Version        =   65536
         _ExtentX        =   3810
         _ExtentY        =   450
         Calculator      =   "FrmViewPTP.frx":8A33
         Caption         =   "FrmViewPTP.frx":8A53
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmViewPTP.frx":8ABF
         Keys            =   "FrmViewPTP.frx":8ADD
         Spin            =   "FrmViewPTP.frx":8B27
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   10147522
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
         TabIndex        =   17
         Top             =   945
         Width           =   2160
         _Version        =   65536
         _ExtentX        =   3810
         _ExtentY        =   450
         Calculator      =   "FrmViewPTP.frx":8B4F
         Caption         =   "FrmViewPTP.frx":8B6F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmViewPTP.frx":8BDB
         Keys            =   "FrmViewPTP.frx":8BF9
         Spin            =   "FrmViewPTP.frx":8C43
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   10147522
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
      Begin Crystal.CrystalReport RPT 
         Left            =   4470
         Top             =   1140
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin TDBNumber6Ctl.TDBNumber TxtPaymentMonthSebenarnya 
         Height          =   255
         Left            =   2520
         TabIndex        =   86
         Top             =   2100
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   450
         Calculator      =   "FrmViewPTP.frx":8C6B
         Caption         =   "FrmViewPTP.frx":8C8B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmViewPTP.frx":8CF7
         Keys            =   "FrmViewPTP.frx":8D15
         Spin            =   "FrmViewPTP.frx":8D5F
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   65535
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "#########0;;Null"
         EditMode        =   0
         Enabled         =   0
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "#########0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999999
         MinValue        =   -99999999
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   -1
         ValueVT         =   2097153
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin VB.Label LblJmlVjust 
         Alignment       =   2  'Center
         Caption         =   "0"
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
         Left            =   120
         TabIndex        =   99
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H00B1FDD5&
         Caption         =   "Occupation:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   98
         Top             =   3480
         Width           =   1155
      End
      Begin VB.Label Label7 
         BackColor       =   &H00B1FDD5&
         Caption         =   "Reason:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   97
         Top             =   3840
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Handle:"
         Height          =   240
         Index           =   23
         Left            =   60
         TabIndex        =   93
         Top             =   3060
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment/Month By System:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   21
         Left            =   120
         TabIndex        =   87
         Top             =   2100
         Width           =   2370
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Charge"
         Height          =   240
         Index           =   39
         Left            =   0
         TabIndex        =   22
         Top             =   630
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Amount"
         Height          =   240
         Index           =   38
         Left            =   0
         TabIndex        =   21
         Top             =   990
         Width           =   1140
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "From o/s balance %"
         Height          =   330
         Index           =   37
         Left            =   45
         TabIndex        =   20
         Top             =   1305
         Width           =   2220
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "principal (%) from "
         Height          =   240
         Index           =   36
         Left            =   45
         TabIndex        =   19
         Top             =   1665
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Justification"
         Height          =   240
         Index           =   29
         Left            =   120
         TabIndex        =   18
         Top             =   4185
         Width           =   1230
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "Review CPA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11955
   End
End
Attribute VB_Name = "FrmViewPTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StatusPTP As String
Dim PaymentTenor As Double

Private Sub CmbOccupation_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub CmbPaymentHandle_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub



Private Sub CmbReason_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub CmdApprove_Click()
    If FrmListRequestPTP.CmbTampilkan.text = "PTP DISC." And _
       FrmListRequestPTP.LvPTP.SelectedItem.SubItems(4) = "Belum di Approve" Then
        MsgBox "PTP yang diajukan PTP Discon! Setelah perview ini tampil, Harap Print dan ajukan ke SPV anda terlebih dahulu!", vbOKOnly + vbInformation, "Informasi"
        
        'Jika yang login Administrator,Admin atau Supervisor
        If UCase(MDIForm1.txtlevel.text) = "ADMINISTRATOR" Or _
           UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Or _
           UCase(MDIForm1.txtlevel.text) = "ADMIN" Then
            MsgBox "Maaf, untuk approve PTP Discount harus melalui Form List Request PTP di depan, karena Approve PTP Discon harus menentukan SPV/Manager yang approve PTP!", vbOKOnly + vbInformation, "Informasi"
            Exit Sub
        End If
        
        Call BikinCPASementara
        Call PreviewCPA
    Else
        DoEvents
        Call BikinCPA
        DoEvents
        Call BikinPTP
        DoEvents
        Call CatetLogApprove
        DoEvents
        Call BikinStatusPTP
        DoEvents
        Call HapusData
        DoEvents
        Call KirimPesan
        MsgBox "Pembuatan CPA dan PTP berhasil!", vbOKOnly + vbInformation, "Informasi"
    End If
    
    
    FrmListRequestPTP.IsiLog
    Me.Hide
End Sub

Private Sub CmdCetak_Click()
    Call BikinCPASementara
    Call PreviewCPA
End Sub

Private Sub CmdGetJustification_Click()
    FrmGetJustification_SendPTP.Show vbModal
End Sub

Private Sub cmdkeluar_Click()
    Unload Me
End Sub



Private Sub CmdUpdate_Click()
    Dim CMDSQL As String
    Dim W As String
    Dim Occupation() As String
    Dim Reason() As String
    
    If CmbOccupation.text = "" Or _
       IsNull(CmbOccupation.text) = True Or _
       CmbOccupation.text = Empty Then
       MsgBox "Occupation tidak boleh kosong!", vbOKOnly + vbExclamation, "Peringatan"
       Exit Sub
    End If
    
    If CmbReason.text = "" Or _
       IsNull(CmbReason.text) = True Or _
       CmbReason.text = Empty Then
       MsgBox "Reason tidak boleh kosong!", vbOKOnly + vbExclamation, "Peringatan"
       Exit Sub
    End If
    
    If CmbPaymentHandle.text = "" Or _
       IsNull(CmbPaymentHandle.text) = True Or _
       CmbPaymentHandle.text = Empty Then
       MsgBox "Payment Handle tidak boleh kosong!", vbOKOnly + vbExclamation, "Peringatan"
       Exit Sub
    End If
    
    W = MsgBox("Apakah anda yakin data akan diupdate?", vbYesNo + vbQuestion, "Konfirmasi")
    If W = vbNo Then
        MsgBox "Update dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If lblLastPay.Value = Empty Or lblLastPay.Value = 0 Then
        MsgBox "Total Payment tidak boleh kosong!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    
    If txtbalance.Value = Empty Or txtbalance.Value = 0 Then
        MsgBox "Balance tidak boleh kosong!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    If tdbisnstallment.Value = Empty Then
        MsgBox "Tenor tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    
    If lblLastPay.Value > txtbalance.Value Then
        MsgBox "Payment tidak boleh lebih besar dari balance!", vbOKOnly + vbExclamation, "Peringatan"
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
    
    'Occupation()=split(CmbOccupation.Text,"-")
    
         
    '@@22062012 Jika Payment=Balance di Database maka otomatis jadi PTP No Discount
    If lblLastPay.Value = txtbalance.Value Then
        CmbJenisPTP.text = "PTP No Discount"
    End If
    
    If lblLastPay.Value < txtbalance.Value Then
        CmbJenisPTP.text = "PTP Discount"
    End If
    
    CMDSQL = "update tblsendptp set total_amount_deal='"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(lblLastPay.Value), "0", lblLastPay.Value)) + "',tenor='"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(tdbisnstallment.Value), "1", tdbisnstallment.Value)) + "',balance='"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(txtbalance.Value), "0", txtbalance.Value)) + "',principal='"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(txtprincipal.Value), "0", txtprincipal.Value)) + "',nttlpayment='"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(lblLastPay.Value), "0", lblLastPay.Value)) + "',ndownpay='"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(txtdownpayment.Value), "0", txtdownpayment.Value)) + "',nfuturepay='"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(txtfuture.Value), "0", txtfuture.Value)) + "',ncharge='"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(txtcharge.Value), "0", txtcharge.Value)) + "',ndiscountamt='"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(txtdiscount.Value), "0", txtdiscount.Value)) + "',vosbalance='"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(txtfrombalancepersen.text), "", txtfrombalancepersen.text)) + "',vosprincipal='"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(txtpersenprincipal.text), "", txtpersenprincipal.text)) + "',vjust='"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(txtjust.text), "", txtjust.text)) + "',nbalance='"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(txtbalance.Value), "0", txtbalance.Value)) + "',nperiod='"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(tdbisnstallment.Value), "1", tdbisnstallment.Value)) + "',chkfaxed='"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(strFaxed), "0", strFaxed)) + "',nprincipal='"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(txtprincipal.Value), "0", txtprincipal.Value)) + "',chkwentalking='"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(strwentalk), "0", strwentalk)) + "',chkktp='"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(strKTP), "0", strKTP)) + "',chksup='"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(strSup), "0", strSup)) + "',chkbillings='"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(strBilling), "0", strBilling)) + "',chkothers='"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(strOthers), "0", strOthers)) + "',payment_after_tenor='"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(TxtPaymentMonthSebenarnya.Value), "0", TxtPaymentMonthSebenarnya.Value)) + "', ket_other='"
    CMDSQL = CMDSQL + IIf(IsNull(txtothers.text), "", txtothers.text) + "',payment_handle='"
    CMDSQL = CMDSQL + IIf(IsNull(CmbPaymentHandle.text), "", CmbPaymentHandle.text) + "', occupation='"
    
    CMDSQL = CMDSQL + IIf(IsNull(CmbOccupation.text), "", CmbOccupation.text) + "',reason='"
    CMDSQL = CMDSQL + IIf(IsNull(CmbReason.text), "", CmbReason.text) + "', jenis_ptp='"
    CMDSQL = CMDSQL + IIf(IsNull(CmbJenisPTP.text), "", CmbJenisPTP.text) + "' "
    
    CMDSQL = CMDSQL + " where id='"
    CMDSQL = CMDSQL + CStr(TxtIdCpa.text) + "'"
    
    On Error GoTo Salah
    M_OBJCONN.Execute CMDSQL
    MsgBox "Data berhasil diupdate!", vbOKOnly + vbInformation, "Informasi"
    FrmListRequestPTP.IsiLog
    Unload Me
    Exit Sub
Salah:
    MsgBox "Ada error: " & err.Description
End Sub



Private Sub Form_Load()
    If UCase(MDIForm1.txtlevel.text) <> "TEAMLEADER" Then
        CmdCetak.Visible = True
    Else
        CmdCetak.Visible = False
    End If
End Sub

Private Sub tdbisnstallment_Change()
    Call PaymentAfterTenor
End Sub

Private Sub txtbalance_Change()
    txtcharge.Value = txtbalance.Value - txtprincipal.Value
    txtdiscount.Value = txtbalance.Value - lblLastPay.Value
    If txtbalance.Value <> 0 Then
         txtfrombalancepersen.text = Round(((lblLastPay.Value / txtbalance.Value) - 1) * 100, 2)
    End If
    
    '@@ 12Juni2012, Jika Balance=0 maka persentase balance =0
    If txtbalance.Value = 0 Then
        txtfrombalancepersen.text = 0
    End If
     
    '@@22062012 Jika Payment=Balance di Database maka otomatis jadi PTP No Discount
    If lblLastPay.Value = txtbalance.Value Then
        CmbJenisPTP.text = "PTP No Discount"
    End If
    
    If lblLastPay.Value < txtbalance.Value Then
        CmbJenisPTP.text = "PTP Discount"
    End If
    
End Sub


Private Sub lblLastPay_Change()
    txtdiscount.Value = txtbalance.Value - lblLastPay.Value
    '@@ 23022012, rumusnya diubah nih
    If txtbalance.Value <> 0 Then
        txtfrombalancepersen.text = Round(((lblLastPay.Value / txtbalance.Value) - 1) * 100, 2)
    End If
    If txtprincipal.Value <> 0 Then
        txtpersenprincipal.text = Round(((lblLastPay.Value / txtprincipal.Value) - 1) * 100, 2)
    End If
    
    '@@ 12Juni2012, Jika Balance=0 maka persentase balance=0. Jika Principal=0 maka persentase principal=0
    If txtbalance.Value = 0 Then
        txtfrombalancepersen.text = "0"
    End If
    If txtprincipal.Value = 0 Then
        txtpersenprincipal.text = "0"
    End If
    
    '@@19062012 Tambahan View PTP bisa diedit
    Call PaymentAfterTenor
    
      '@@22062012 Jika Payment=Balance di Database maka otomatis jadi PTP No Discount
    If lblLastPay.Value = txtbalance.Value Then
        CmbJenisPTP.text = "PTP No Discount"
    End If
    
    If lblLastPay.Value < txtbalance.Value Then
        CmbJenisPTP.text = "PTP Discount"
    End If
    
'    If lblLastPay.Value > txtbalance.Value Then
'        MsgBox "Total Payment tidak boleh lebih besar dari balance!", vbOKOnly + vbInformation, "Informasi"
'        lblLastPay.Value = 1
'    End If
End Sub

Private Sub txtdownpayment_Change()
    txtfuture.Value = lblLastPay.Value - txtdownpayment.Value
    Call PaymentAfterTenor
End Sub


Private Sub txtjust_Change()
    LblJmlVjust.Caption = Len(txtjust.text)
    If Val(Len(txtjust.text)) >= 250 Then
        MsgBox "Maksimal Justifikasi hanya 250 Karakter!", vbOKOnly + vbInformation, "Informasi"
    End If
End Sub

Private Sub txtprincipal_Change()
    txtcharge.Value = txtbalance.Value - txtprincipal.Value
    If txtprincipal.Value <> 0 Then
        txtpersenprincipal.text = Round(((lblLastPay.Value / txtprincipal.Value) - 1) * 100, 2)
    End If
    
    '@@12 Juni 2012 Jika principal=0 maka persentase principal =0
    If txtprincipal.Value = 0 Then
        txtpersenprincipal.text = "0"
    End If
End Sub

'================================== Approve CPA Dan PTP ============================================
Private Sub BikinCPA()
    Dim CMDSQL As String
    Dim Remarks As String
    
    With FrmListRequestPTP.LvPTP.SelectedItem
        Call Cari_LPD_LPA_Payment
        
        CMDSQL = "insert into tblcpa (dpropsal,vcustid,vproduct,nttlpayment,ndownpay,"
        CMDSQL = CMDSQL + "ncharge,ndiscountamt,vosbalance,vosprincipal,dtglinsert,vcustname,vjust,"
        CMDSQL = CMDSQL + "nbalance,nprincipal,nperiod,chkfaxed,chkwentalking,chkktp,"
        CMDSQL = CMDSQL + "chksup,chkbillings,chkothers,lpd_from_payment,lpa_from_payment,"
        CMDSQL = CMDSQL + "f_system,dob,status_ptp,"
        CMDSQL = CMDSQL + "ketother "
        
        '@@19062012 Jika Status PTP DISCON Catat Approvenya
        If Trim(UCase(.SubItems(1))) = "PTP DISCOUNT" Then
            CMDSQL = CMDSQL + ",tglapprove,sts_approve,approve_by,logapprove_by "
        End If
    
        'Catet Juga yang PTP No Discon 20062012
        If Trim(UCase(.SubItems(1))) = "PTP NO DISCOUNT" Then
            CMDSQL = CMDSQL + ",tglapprove,sts_approve,approve_by,logapprove_by "
        End If
                
        '@@16-07-2012 Tambahan Payment Handle
        CMDSQL = CMDSQL + ",vpaymenthandle,voccupation,vreason "
                
        CMDSQL = CMDSQL + ") values ("
        CMDSQL = CMDSQL + "now(),'"
        CMDSQL = CMDSQL + CStr(.SubItems(2)) + "','CARD','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(15)), "0", Replace(.SubItems(15), ",", ""))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(16)), "0", Replace(.SubItems(16), ",", ""))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(17)), "0", Replace(.SubItems(17), ",", ""))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(18)), "0", Replace(.SubItems(18), ",", ""))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(19)), "", .SubItems(19))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(20)), "", .SubItems(20))) + "',"
        CMDSQL = CMDSQL + "now(),'"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(3)), "", .SubItems(3))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(21)), "", .SubItems(21))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(12)), "0", Replace(.SubItems(12), ",", ""))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(11)), "0", Replace(.SubItems(11), ",", ""))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(8)), "", .SubItems(8))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(22)), "", .SubItems(22))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(23)), "", .SubItems(23))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(24)), "", .SubItems(24))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(25)), "", .SubItems(25))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(26)), "", .SubItems(26))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(27)), "", .SubItems(27))) + "',"
        CMDSQL = CMDSQL + IIf(TxtLPDPayment.text = "", "null", "'" + TxtLPDPayment.text + "'") + ",'"
        CMDSQL = CMDSQL + CStr(TxtLPAPayment.Value) + "','1'"
        
        '@@20062012 Tambahkan DOB dan Status PTP
        CMDSQL = CMDSQL + IIf(.SubItems(29) = "", ",null", ",'" + .SubItems(29) + "'")
        CMDSQL = CMDSQL + ",'" + .SubItems(1) + "','"
        '@@21062012 Tambahin Buat Keterangan Other
        CMDSQL = CMDSQL + IIf(IsNull(.SubItems(30)), "", .SubItems(30)) + "' "
        
        '@@19062012 Buat nyatet approvenya
         If Trim(UCase(.SubItems(1))) = "PTP DISCOUNT" Then
            CMDSQL = CMDSQL + ",now(),'1','"
            CMDSQL = CMDSQL + Trim(FrmListRequestPTP.CmbApprove.text) + "','"
            CMDSQL = CMDSQL + MDIForm1.txtusername.text + "'"
         End If
         
         'Buat nyatet yang jenisnya PTP NO Discount.
         If Trim(UCase(.SubItems(1))) = "PTP NO DISCOUNT" Then
            CMDSQL = CMDSQL + ",now(),'1','"
            CMDSQL = CMDSQL + MDIForm1.txtusername.text + "','"
            CMDSQL = CMDSQL + MDIForm1.txtusername.text + "'"
         End If
        
        
        CMDSQL = CMDSQL + ",'"
        CMDSQL = CMDSQL + IIf(IsNull(.SubItems(31)), "", .SubItems(31)) + "','"
        CMDSQL = CMDSQL + IIf(IsNull(.SubItems(32)), "", .SubItems(32)) + "','"
        CMDSQL = CMDSQL + IIf(IsNull(.SubItems(33)), "", .SubItems(33)) + "')"
        M_OBJCONN.Execute CMDSQL
        
       '@@ 11092012 Buat Nyatet Approve
            Remarks = "PTPNoDisc-App By:" + MDIForm1.txtusername.text + "-"
            Remarks = Remarks + "Ttl.Pymt: Rp." + CStr(IIf(IsNull(.SubItems(7)), "", .SubItems(7))) + " -"
            Remarks = Remarks + "Instl: " + CStr(IIf(IsNull(.SubItems(8)), "", .SubItems(8))) + " -"
            Remarks = Remarks + "From Bal.: Rp." + CStr(IIf(IsNull(.SubItems(12)), "", .SubItems(12))) + " -"
            Remarks = Remarks + "From Prin.: Rp." + CStr(IIf(IsNull(.SubItems(14)), "", .SubItems(14))) + " -"
            Remarks = Remarks + "%Balance: " + CStr(IIf(IsNull(.SubItems(19)), "", .SubItems(19))) + "% -"
            Remarks = Remarks + "%Principal: " + CStr(IIf(IsNull(.SubItems(20)), "", .SubItems(20))) + "% #USER LOG:" + MDIForm1.txtusername.text
            
            CMDSQL = "insert into mgm_hst (custid, agent, products, "
            CMDSQL = CMDSQL + "hst,user_log) values ('"
            CMDSQL = CMDSQL + CStr(.SubItems(2)) + "','"
            CMDSQL = CMDSQL + .SubItems(28) + "','"
            CMDSQL = CMDSQL + "Collection" + "','"
            CMDSQL = CMDSQL + Remarks + "','"
            CMDSQL = CMDSQL + MDIForm1.txtusername.text + "')"
            M_OBJCONN.Execute CMDSQL
    End With
 End Sub

Private Sub Cari_LPD_LPA_Payment()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    
    StatusPTP = ""
    TxtLPDPayment.text = ""
    TxtLPAPayment.Value = "0"
    
    With FrmListRequestPTP.LvPTP.SelectedItem
        CMDSQL = "select paydate,payment from tbllunas where custid='"
        CMDSQL = CMDSQL + Trim(.SubItems(2)) + "' order by paydate desc limit 1 "
        
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If M_objrs.RecordCount > 0 Then
                StatusPTP = "PTP-POP"
                TxtLPDPayment.text = IIf(IsNull(M_objrs("paydate")), "", Format(M_objrs("paydate"), "yyyy-mm-dd"))
                TxtLPAPayment.Value = IIf(IsNull(M_objrs("payment")), "0", M_objrs("payment"))
                LpdPayment = "'" + TxtLPDPayment.text + "'"
            Else
                StatusPTP = "PTP-NEW"
                'LpdPayment = "null"
                TxtLPDPayment.text = ""
                TxtLPAPayment.Value = "0"
            End If
        Set M_objrs = Nothing
    End With
End Sub

Private Sub CatetLogApprove()
    Dim CMDSQL As String
    
    CMDSQL = "insert into tblsendptp_log_approve "
    CMDSQL = CMDSQL + "select * from tblsendptp where id='"
    CMDSQL = CMDSQL + CStr(FrmListRequestPTP.LvPTP.SelectedItem.text) + "'"
    M_OBJCONN.Execute CMDSQL
End Sub

Private Sub BikinStatusPTP()
    Dim CMDSQL As String
    
    With FrmListRequestPTP.LvPTP.SelectedItem
        'Jika StatusPTP=PTP NEW
        If StatusPTP = "PTP-NEW" Then
            CMDSQL = "update mgm set dateptpnew='"
            CMDSQL = CMDSQL + .SubItems(6) + "',tgl_tagih='"
            CMDSQL = CMDSQL + .SubItems(10) + "', amountnew='"
            CMDSQL = CMDSQL + CStr(Replace(.SubItems(15), ",", "")) + "',tglallptp='"
            CMDSQL = CMDSQL + .SubItems(6) + "',f_cek_new='PTP-NE',"
            CMDSQL = CMDSQL + "tglincoming=now(),ttlptp='"
            CMDSQL = CMDSQL + CStr(Replace(.SubItems(15), ",", "")) + "',"
            CMDSQL = CMDSQL + "kethslkerja_new='PTP-NEW',kethslkerjadesc_new='PTP-NEW',ptpvia='"
            CMDSQL = CMDSQL + CStr(.SubItems(9)) + "',ptpdesc='PTP-NEW', dateptp='"
            CMDSQL = CMDSQL + .SubItems(6) + "',tglptpnew=now(),tenor='"
            CMDSQL = CMDSQL + CStr(.SubItems(8)) + "' "
            CMDSQL = CMDSQL + "where custid='"
            CMDSQL = CMDSQL + CStr(.SubItems(2)) + "'"
            M_OBJCONN.Execute CMDSQL
        End If
        
        If StatusPTP = "PTP-POP" Then
            CMDSQL = "update mgm set dateptp='"
            CMDSQL = CMDSQL + .SubItems(6) + "',tgl_tagih='"
            CMDSQL = CMDSQL + .SubItems(10) + "',tglallptp='"
            CMDSQL = CMDSQL + .SubItems(6) + "',f_cek_new='PTP-PO',"
            CMDSQL = CMDSQL + "tglincoming=now(),ttlptp='"
            CMDSQL = CMDSQL + CStr(Replace(.SubItems(15), ",", "")) + "',"
            CMDSQL = CMDSQL + "kethslkerja_new='PTP-POP',kethslkerjadesc_new='PTP-POP',ptpvia='"
            CMDSQL = CMDSQL + CStr(.SubItems(9)) + "',ptpdesc='PTP-POP',amountptp='"
            CMDSQL = CMDSQL + CStr(Replace(.SubItems(15), ",", "")) + "',tenor='"
            CMDSQL = CMDSQL + CStr(.SubItems(8)) + "' "
            CMDSQL = CMDSQL + "where custid='"
            CMDSQL = CMDSQL + CStr(.SubItems(2)) + "'"
            M_OBJCONN.Execute CMDSQL
        End If
    End With
End Sub

Private Sub KirimPesan()
    Dim CMDSQL As String
    Dim Remarks As String
    Dim M_objrs As ADODB.Recordset
    
    With FrmListRequestPTP.LvPTP.SelectedItem
        Remarks = "Pembuatan PTP untuk custid: " & .SubItems(2) & " telah di approve!"
        CMDSQL = "insert into msgtbl "
        CMDSQL = CMDSQL + "( recipient, datetime, sender, sentfrom, msg) values ('"
        CMDSQL = CMDSQL + .SubItems(28) + "','"
        CMDSQL = CMDSQL + Format(Now(), "yyyymmdd") + "','"
        CMDSQL = CMDSQL + MDIForm1.txtusername.text + "','"
        CMDSQL = CMDSQL + CStr(MDIForm1.Winsock1.LocalIP) + "','"
        CMDSQL = CMDSQL + Remarks + "')"
        M_OBJCONN.Execute CMDSQL
        
        
        '@@19072012 Kirim Pesan Buat Ke TL
        'Cari Nama TLNYA
        CMDSQL = "select team from usertbl where userid='"
        CMDSQL = CMDSQL + CStr(Trim(.SubItems(28))) + "' "
        CMDSQL = CMDSQL + " and team is not null "
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_objrs.RecordCount > 0 Then
            CMDSQL = "insert into msgtbl "
            CMDSQL = CMDSQL + "(recipient, datetime, sender, sentfrom, msg) values ('"
            CMDSQL = CMDSQL + CStr(Trim(M_objrs("team"))) + "','"
            CMDSQL = CMDSQL + Format(Now(), "yyyymmdd") + "','"
            CMDSQL = CMDSQL + MDIForm1.txtusername.text + "','"
            CMDSQL = CMDSQL + CStr(MDIForm1.Winsock1.LocalIP) + "','"
            CMDSQL = CMDSQL + Remarks + "')"
            M_OBJCONN.Execute CMDSQL
        End If
        
        Set M_objrs = Nothing
        
    End With
End Sub


Private Sub BikinPTP()
    Dim CMDSQL As String
    Dim i As Integer
    Dim M_Objrs_Cek_Tgl As ADODB.Recordset
    Dim jumlah_tenor As Integer
        
        
    With FrmListRequestPTP.LvPTP.SelectedItem
    jumlah_tenor = Val(.SubItems(8))
    bcekptp = True

'UNTUK PTP SEKALI BAYAR
    If jumlah_tenor = 1 Then
        '@@14-04-2012 Cek Data
        'PROSES INPUT DATA KE TBLNEGOPTP
        CMDSQL = "select * from tblnegoptp where custid='"
        CMDSQL = CMDSQL + CStr(.SubItems(2)) + "' and date(promisedate)='"
        CMDSQL = CMDSQL + CStr(.SubItems(6)) + "'"
        Set M_Objrs_Cek_Tgl = New ADODB.Recordset
        M_Objrs_Cek_Tgl.CursorLocation = adUseClient
        M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            
        If M_Objrs_Cek_Tgl.RecordCount > 0 Then
            'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
            While Not M_Objrs_Cek_Tgl.EOF
                CMDSQL = "delete from tblnegoptp where id='"
                CMDSQL = CMDSQL + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                M_OBJCONN.Execute CMDSQL
                M_Objrs_Cek_Tgl.MoveNext
            Wend
        End If
        Set M_Objrs_Cek_Tgl = Nothing
                  
        jatuhtempo = .SubItems(6)
        CMDSQL = "INSERT INTO TblNegoPTP "
        CMDSQL = CMDSQL + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type,Tenor,down_payment) "
        CMDSQL = CMDSQL + "VALUES "
        CMDSQL = CMDSQL + "('" + CStr(.SubItems(2)) + "', "
        CMDSQL = CMDSQL + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
        CMDSQL = CMDSQL + "" + CStr(Replace(.SubItems(13), ",", "")) + " , "
        CMDSQL = CMDSQL + "now(), "
        CMDSQL = CMDSQL + "'PAID-OFF'),"
        CMDSQL = CMDSQL + "'1'),"
        CMDSQL = CMDSQL + " '" + CStr(Replace(.SubItems(13), ",", "")) + ")"
        M_OBJCONN.Execute CMDSQL
        
            
            
        '@@14-04-2012 Cek Data
        'PROSES INPUT DATA KE TBLNEGOPTP_LOG
        CMDSQL = "select * from tblnegoptp_log where custid='"
        CMDSQL = CMDSQL + CStr(.SubItems(2)) + "' and date(promisedate)='"
        CMDSQL = CMDSQL + CStr(.SubItems(6)) + "'"
        Set M_Objrs_Cek_Tgl = New ADODB.Recordset
        M_Objrs_Cek_Tgl.CursorLocation = adUseClient
        M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_Objrs_Cek_Tgl.RecordCount > 0 Then
            'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
            While Not M_Objrs_Cek_Tgl.EOF
                CMDSQL = "delete from tblnegoptp_log where id='"
                CMDSQL = CMDSQL + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                M_OBJCONN.Execute CMDSQL
                M_Objrs_Cek_Tgl.MoveNext
            Wend
        End If
        Set M_Objrs_Cek_Tgl = Nothing
    
    
        ' isi ke tbl log_ptp
        CMDSQL = "INSERT INTO tblnegoptp_log "
        CMDSQL = CMDSQL + "(custid,PromiseDate, Promisepay,tglInput,agent,stsacc) "
        CMDSQL = CMDSQL + "VALUES "
        CMDSQL = CMDSQL + "('" + CStr(.SubItems(2)) + "', "
        CMDSQL = CMDSQL + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
        CMDSQL = CMDSQL + "" + CStr(Replace(.SubItems(13), ",", "")) + " , "
        CMDSQL = CMDSQL + "now(), "
        CMDSQL = CMDSQL + "'" + CStr(.SubItems(28)) + "','P')"
        M_OBJCONN.Execute CMDSQL
'================================================================================================================
    Else
        'Untuk Tenor yang lebih dari 1
               
        'Hapus Reserved Data
        CMDSQL = "delete from tblreserve where custid='"
        CMDSQL = CMDSQL + CStr(.SubItems(2)) + "'"
        M_OBJCONN.Execute CMDSQL
                
        jatuhtempo = CStr(.SubItems(6))
    
        '@@14-04-2012 Cek Data
        CMDSQL = "select * from tblnegoptp where custid='"
        CMDSQL = CMDSQL + CStr(.SubItems(2)) + "' and date(promisedate)='"
        CMDSQL = CMDSQL + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "'"
        Set M_Objrs_Cek_Tgl = New ADODB.Recordset
        M_Objrs_Cek_Tgl.CursorLocation = adUseClient
        M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_Objrs_Cek_Tgl.RecordCount > 0 Then
            'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
            While Not M_Objrs_Cek_Tgl.EOF
                CMDSQL = "delete from tblnegoptp where id='"
                CMDSQL = CMDSQL + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                M_OBJCONN.Execute CMDSQL
                M_Objrs_Cek_Tgl.MoveNext
            Wend
        End If
        Set M_Objrs_Cek_Tgl = Nothing
            
        CMDSQL = "INSERT INTO TblNegoPTP "
        CMDSQL = CMDSQL + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
        CMDSQL = CMDSQL + "VALUES "
        CMDSQL = CMDSQL + "('" + CStr(.SubItems(2)) + "', "
        CMDSQL = CMDSQL + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
        CMDSQL = CMDSQL + "" + CStr(Replace(.SubItems(13), ",", "")) + " , "
        CMDSQL = CMDSQL + "now(), "
        CMDSQL = CMDSQL + "'IPO')"
        M_OBJCONN.Execute CMDSQL
        
        '@@14-04-2012 Cek Data
        CMDSQL = "select * from tblnegoptp_log where custid='"
        CMDSQL = CMDSQL + CStr(.SubItems(2)) + "' and date(promisedate)='"
        CMDSQL = CMDSQL + CStr(.SubItems(6)) + "'"
        Set M_Objrs_Cek_Tgl = New ADODB.Recordset
        M_Objrs_Cek_Tgl.CursorLocation = adUseClient
        M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_Objrs_Cek_Tgl.RecordCount > 0 Then
            'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
            While Not M_Objrs_Cek_Tgl.EOF
                CMDSQL = "delete from tblnegoptp_log where id='"
                CMDSQL = CMDSQL + CStr(cnull(M_Objrs_Cek_Tgl("id"))) + "'"
                M_OBJCONN.Execute CMDSQL
                M_Objrs_Cek_Tgl.MoveNext
            Wend
        End If
        Set M_Objrs_Cek_Tgl = Nothing
            
            
        'isi ke tbl log_ptp
        CMDSQL = "INSERT INTO tblnegoptp_log "
        CMDSQL = CMDSQL + "(custid,PromiseDate, Promisepay,tglInput,agent,stsacc) "
        CMDSQL = CMDSQL + "VALUES "
        CMDSQL = CMDSQL + "('" + CStr(.SubItems(2)) + "', "
        CMDSQL = CMDSQL + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
        CMDSQL = CMDSQL + "" + CStr(Replace(.SubItems(13), ",", "")) + " , "
        CMDSQL = CMDSQL + "now(), "
        CMDSQL = CMDSQL + "'" + CStr(.SubItems(28)) + "','P')"
        M_OBJCONN.Execute CMDSQL
                
'            Set listitem = .LstPayment.ListItems.ADD(, , "")
'                listitem.SubItems(1) = ""
'                listitem.SubItems(2) = Format(.TDBDate3.Value, "dd/mm/yyyy")
'                listitem.SubItems(3) = CStr(txtPembayaranAwal.Value)
'                listitem.SubItems(4) = "IPO"
'                listitem.SubItems(5) = MDIForm1.TDBDate1.Value
                
        n = 0
        
        
        Call HitungInstallmentPtp
        
        For i = 1 To (jumlah_tenor - 1)
            n = n + 1
            'JMLPAY = ((.TxtPayment - txtPembayaranAwal.Value) - PaymentTenor) / (.txttenor.Value - 1)
            JmlPay = PaymentTenor
            Vrdate = DateAdd("m", n, Format(.SubItems(6), "yyyy-mm-dd"))
                    
            '@@14-04-2012 Cek Data
            CMDSQL = "select * from tblreserve where custid='"
            CMDSQL = CMDSQL + CStr(.SubItems(2)) + "' and date(promisedate)='"
            CMDSQL = CMDSQL + CStr(Format(Vrdate, "yyyy-mm-dd")) + "'"
            Set M_Objrs_Cek_Tgl = New ADODB.Recordset
            M_Objrs_Cek_Tgl.CursorLocation = adUseClient
            M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
            If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                While Not M_Objrs_Cek_Tgl.EOF
                    CMDSQL = "delete from tblreserve where id='"
                    CMDSQL = CMDSQL + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                    M_OBJCONN.Execute CMDSQL
                    M_Objrs_Cek_Tgl.MoveNext
                Wend
            End If
            Set M_Objrs_Cek_Tgl = Nothing
                    
            CMDSQL = "INSERT INTO tblreserve "
            CMDSQL = CMDSQL + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
            CMDSQL = CMDSQL + "VALUES "
            CMDSQL = CMDSQL + "('" + CStr(.SubItems(2)) + "', "
            CMDSQL = CMDSQL + "'" + CStr(Format(Vrdate, "yyyy-mm-dd")) + "', "
            CMDSQL = CMDSQL + "" + CStr(JmlPay) + " , "
            CMDSQL = CMDSQL + "now(), "
            CMDSQL = CMDSQL + "'IPO')"
            M_OBJCONN.Execute CMDSQL
        
            '@@14-04-2012 Cek Data
            CMDSQL = "select * from tblnegoptp_log where custid='"
            CMDSQL = CMDSQL + CStr(.SubItems(2)) + "' and date(promisedate)='"
            CMDSQL = CMDSQL + CStr(Format(Vrdate, "yyyy-mm-dd")) + "'"
            Set M_Objrs_Cek_Tgl = New ADODB.Recordset
            M_Objrs_Cek_Tgl.CursorLocation = adUseClient
            M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            
            If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                While Not M_Objrs_Cek_Tgl.EOF
                    CMDSQL = "delete from tblnegoptp_log where id='"
                    CMDSQL = CMDSQL + CStr(cnull(M_Objrs_Cek_Tgl("id"))) + "'"
                    M_OBJCONN.Execute CMDSQL
                    M_Objrs_Cek_Tgl.MoveNext
                Wend
            End If
            Set M_Objrs_Cek_Tgl = Nothing
        
            CMDSQL = "INSERT INTO TblNegoptp_log "
            CMDSQL = CMDSQL + "(custid,PromiseDate, Promisepay,tglinput,agent,stsacc) "
            CMDSQL = CMDSQL + "VALUES "
            CMDSQL = CMDSQL + "('" + CStr(.SubItems(2)) + "', "
            CMDSQL = CMDSQL + "'" + CStr(Format(Vrdate, "yyyy-mm-dd")) + "', "
            CMDSQL = CMDSQL + "" + CStr(JmlPay) + " , "
            CMDSQL = CMDSQL + "now(), "
            CMDSQL = CMDSQL + "'" + CStr(.SubItems(28)) + "','R')"
            M_OBJCONN.Execute CMDSQL
            
            'INSERT KE TABEL PTP-REGULER(Randy07-04-2015)
            CMDSQL = "select * from tblnegoptp_reguler where custid='"
            CMDSQL = CMDSQL + CStr(.SubItems(2)) + "' and date(promisedate)='"
            CMDSQL = CMDSQL + CStr(Format(Vrdate, "yyyy-mm-dd")) + "'"
            Set M_Objrs_Cek_Tgl = New ADODB.Recordset
            M_Objrs_Cek_Tgl.CursorLocation = adUseClient
            M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            
            If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                While Not M_Objrs_Cek_Tgl.EOF
                    CMDSQL = "delete from tblnegoptp_reguler where id='"
                    CMDSQL = CMDSQL + CStr(cnull(M_Objrs_Cek_Tgl("id"))) + "'"
                    M_OBJCONN.Execute CMDSQL
                    M_Objrs_Cek_Tgl.MoveNext
                Wend
            End If
            Set M_Objrs_Cek_Tgl = Nothing
        
            CMDSQL = "INSERT INTO tblnegoptp_reguler"
            CMDSQL = CMDSQL + "(custid, balance, PromiseDate, Promisepay, inputdate, type, tenor, down_payment, agent) "
            CMDSQL = CMDSQL + "VALUES "
            CMDSQL = CMDSQL + "('" + CStr(.SubItems(2)) + "', "
            CMDSQL = CMDSQL + " '" + CStr(.SubItems(7)) + "',"
            CMDSQL = CMDSQL + "'" + CStr(Format(Vrdate, "yyyy-mm-dd")) + "', "
            CMDSQL = CMDSQL + "" + CStr(JmlPay) + " , "
            CMDSQL = CMDSQL + "now(), "
            CMDSQL = CMDSQL + "'Reguler',"
            CMDSQL = CMDSQL + " '" + CStr(.SubItems(8)) + "',"
            CMDSQL = CMDSQL + " '" + CStr(.SubItems(16)) + "',"
            CMDSQL = CMDSQL + "'" + CStr(.SubItems(28)) + "' " 'agent
            M_OBJCONN.Execute CMDSQL
            Next i
       End If
    
    PaymentTenor = 0
    End With
End Sub

Private Sub HapusData()
    Dim CMDSQL As String
    
    CMDSQL = "delete from tblsendptp where id='"
    CMDSQL = CMDSQL + CStr(FrmListRequestPTP.LvPTP.SelectedItem.text) + "'"
    M_OBJCONN.Execute CMDSQL
End Sub

'@@22-09-2011 Hitung InstallmentPtp
Private Sub HitungInstallmentPtp()
    Dim installment As Double
    
    With FrmListRequestPTP.LvPTP.SelectedItem
        If Val(.SubItems(8)) = 0 Or Val(.SubItems(8)) = 1 Then
            installment = Val(Replace(.SubItems(15), ",", "")) / 1
        Else
            installment = (Val(Replace(.SubItems(15), ",", "")) - Val(Replace(.SubItems(13), ",", ""))) / (Val(.SubItems(8)) - 1)
        End If
        PaymentTenor = installment
    End With
End Sub


Private Sub PreviewCPA()
    Dim CMDSQL As String
    Dim rsTemp1 As ADODB.Recordset
    Dim rsTemporary As ADODB.Recordset
    
    M_RPTCONN.Execute "delete from tblreportcpa "
    strsql = "select * from tblreportcpa"
    Set rsTemp1 = New ADODB.Recordset
    rsTemp1.CursorLocation = adUseClient
    rsTemp1.Open strsql, M_RPTCONN, adOpenDynamic, adLockOptimistic
                
    CMDSQL = "  SELECT * FROM ( "
    CMDSQL = CMDSQL + " SELECT * FROM (SELECT *  FROM USERTBL) AS B"
    CMDSQL = CMDSQL + " Right JOIN  ( "
    CMDSQL = CMDSQL + " SELECT * FROM ( "
    CMDSQL = CMDSQL + " SELECT  * FROM (  SELECT * FROM TBLCPA_sementara WHERE VCUSTID='" + txtcardno.text + "' "
    'Cmdsql = Cmdsql + " and nid='"
    'Cmdsql = Cmdsql + CStr(TxtIdCpa.Text) + "'"
    CMDSQL = CMDSQL + " order by nid desc limit 1 "
    CMDSQL = CMDSQL + " ) AS A Inner Join "
    CMDSQL = CMDSQL + "  (SELECT * FROM   MGM) AS B  ON A.VCUSTID=B.CUSTID  ) as c)  AS BRU "
    CMDSQL = CMDSQL + " ON BRU.AGENT=B.USERID) AS TBLBARU"
    CMDSQL = CMDSQL + " Left Join ( "
    CMDSQL = CMDSQL + "   select * from ( "
    CMDSQL = CMDSQL + " SELECT custid as cust_no,PAYDATE AS lpd,payment as lpa FROM "
    CMDSQL = CMDSQL + " TBLLUNAS  WHERE ID IN (SELECT MAX(ID) FROM tbllunas GROUP BY CUSTID))  as tblbaru1 "
    CMDSQL = CMDSQL + " WHERE cust_no='" + txtcardno.text + "' ) as bru on tblbaru.custid=bru.cust_no "

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
            If rsTemporary("vcustid_mmu") = "" Or IsNull(rsTemporary("vcustid_mmu")) = True Then
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
            rsTemp1("futurepayment") = IIf(IsNull(rsTemporary("nfuturepay")), 0, rsTemporary("nfuturepay"))
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
            
            
            '@@24022012, Mengambil data LPD dan LPA dari MGM
            rsTemp1("lpd") = IIf(IsNull(rsTemporary("pay_dt")), Null, Format(rsTemporary("pay_dt"), "yyyy-mm-dd"))
            rsTemp1("lpa") = IIf(IsNull(rsTemporary("lastpay")), 0, rsTemporary("lastpay"))
            
            
            rsTemp1("lpd_from_payment") = IIf(IsNull(rsTemporary("lpd_from_payment")), Null, Format(rsTemporary("lpd_from_payment"), "yyyy-mm-dd"))
            rsTemp1("lpa_from_payment") = IIf(IsNull(rsTemporary("lpa_from_payment")), 0, rsTemporary("lpa_from_payment"))
            
'            '@@23022012, Tambahan DOB
            If TxtDob.text <> "" Then
                rsTemp1("dob") = TxtDob.text
            End If
            
           
                rsTemp1.update
                rsTemporary.MoveNext
           Wend
           
            ' Tandain klo udah di print // 13 Oktober 2014
            M_OBJCONN.Execute "UPDATE tblsendptp SET s_print=1 WHERE id=" & Val(Me.TxtIdCpa.text)
            ' --------------------------------------------
            
            RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptCpaRincian.rpt"
            WaitSecs (2)
            Call SHOW_PRN
            Set rsTemp1 = Nothing
            Set rsTemporary = Nothing
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

Private Sub BikinCPASementara()
    Dim CMDSQL As String
    
    With FrmListRequestPTP.LvPTP.SelectedItem
        Call Cari_LPD_LPA_Payment
        
        CMDSQL = "insert into tblcpa_sementara (dpropsal,vcustid,vproduct,nttlpayment,ndownpay,"
        CMDSQL = CMDSQL + "ncharge,ndiscountamt,vosbalance,vosprincipal,dtglinsert,vcustname,vjust,"
        CMDSQL = CMDSQL + "nbalance,nprincipal,nperiod,chkfaxed,chkwentalking,chkktp,"
        CMDSQL = CMDSQL + "chksup,chkbillings,chkothers,lpd_from_payment,"
        CMDSQL = CMDSQL + "lpa_from_payment,f_system,ketother,"
        CMDSQL = CMDSQL + "vpaymenthandle,voccupation,vreason"
        CMDSQL = CMDSQL + ") values ("
        CMDSQL = CMDSQL + "now(),'"
        CMDSQL = CMDSQL + CStr(.SubItems(2)) + "','"
        CMDSQL = CMDSQL + IIf(IsNull(txtproduct.text), "", txtproduct.text) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(15)), "0", Replace(.SubItems(15), ",", ""))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(16)), "0", Replace(.SubItems(16), ",", ""))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(17)), "0", Replace(.SubItems(17), ",", ""))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(18)), "0", Replace(.SubItems(18), ",", ""))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(19)), "", .SubItems(19))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(20)), "", .SubItems(20))) + "',"
        CMDSQL = CMDSQL + "now(),'"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(3)), "", .SubItems(3))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(21)), "", .SubItems(21))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(12)), "0", Replace(.SubItems(12), ",", ""))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(11)), "0", Replace(.SubItems(11), ",", ""))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(8)), "", .SubItems(8))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(22)), "", .SubItems(22))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(23)), "", .SubItems(23))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(24)), "", .SubItems(24))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(25)), "", .SubItems(25))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(26)), "", .SubItems(26))) + "','"
        CMDSQL = CMDSQL + CStr(IIf(IsNull(.SubItems(27)), "", .SubItems(27))) + "',"
        CMDSQL = CMDSQL + IIf(TxtLPDPayment.text = "", "null", "'" + TxtLPDPayment.text + "'") + ",'"
        CMDSQL = CMDSQL + CStr(TxtLPAPayment.Value) + "','1','"
        CMDSQL = CMDSQL + IIf(IsNull(txtothers.text), "", txtothers.text) + "','"
        CMDSQL = CMDSQL + IIf(IsNull(CmbPaymentHandle.text), "", CmbPaymentHandle.text) + "','"
        CMDSQL = CMDSQL + IIf(IsNull(CmbOccupation.text), "", CmbOccupation.text) + "','"
        CMDSQL = CMDSQL + IIf(IsNull(CmbReason.text), "", CmbReason.text) + "')"
        M_OBJCONN.Execute CMDSQL
    End With
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

Private Sub Check1_Click()
    If Check1.Value = vbChecked Then
        txtothers.Enabled = True
        txtothers.BackColor = vbWhite
    Else
        txtothers.Enabled = False
        txtothers.BackColor = &HC0C0C0
    End If
End Sub

'-------------------------------------- Tambahan Rumus --------------------------
Private Sub CariTenor()
    Dim Payment As Double
    Dim DownPayment As Double
    Dim Tenor As Double
    Dim PaymentAfterTenor As Double
    
    Payment = lblLastPay
    DownPayment = txtdownpayment.Value
'    PaymentAfterTenor = TxtPayAfterTenor.Value
'
'    On Error Resume Next
'    Tenor = ((Payment - DownPayment) / PaymentAfterTenor) + 1
'    txttenor.Value = Ceiling(Tenor)
End Sub

Private Sub PaymentAfterTenor()
    Dim PayAfterTenor As Double
    
    PayAfterTenor = 0
    If (tdbisnstallment.Value - 1) = 0 Then
        PayAfterTenor = 0
    Else
        PayAfterTenor = (lblLastPay.Value - txtdownpayment.Value) / (tdbisnstallment - 1)
    End If
    On Error Resume Next
    'TxtPayAfterTenor.Value = PayAfterTenor
    TxtPaymentMonthSebenarnya.Value = Ceiling(PayAfterTenor)
End Sub

Private Function Ceiling(number As Double) As Long
    Ceiling = -Int(-number)
End Function

