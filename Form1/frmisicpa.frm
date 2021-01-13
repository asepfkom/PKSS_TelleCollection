VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmisicpa 
   BackColor       =   &H00B1FDD5&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   9375
   ClientLeft      =   6615
   ClientTop       =   960
   ClientWidth     =   14295
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   14295
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame13 
      Appearance      =   0  'Flat
      BackColor       =   &H00B1FDD5&
      BorderStyle     =   0  'None
      Caption         =   "Frame13"
      ForeColor       =   &H80000008&
      Height          =   9135
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   6510
      Begin VB.Frame Frame15 
         Appearance      =   0  'Flat
         BackColor       =   &H002F735C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   0
         Left            =   225
         TabIndex        =   51
         Top             =   90
         Width           =   6180
         Begin VB.Image Image1 
            Height          =   375
            Index           =   0
            Left            =   75
            Picture         =   "frmisicpa.frx":0000
            Stretch         =   -1  'True
            Top             =   40
            Width           =   375
         End
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
            TabIndex        =   52
            Top             =   45
            Width           =   1455
         End
      End
      Begin VB.TextBox txtregion 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1485
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   585
         Width           =   1995
      End
      Begin VB.TextBox txtreff 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1485
         MaxLength       =   1
         TabIndex        =   49
         Top             =   1260
         Width           =   1995
      End
      Begin VB.TextBox txtproduct 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1485
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   1620
         Width           =   1995
      End
      Begin VB.TextBox txtarrangement 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1485
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   1980
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
         TabIndex        =   45
         Top             =   2430
         Width           =   6090
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
            TabIndex        =   46
            Top             =   90
            Width           =   3255
         End
         Begin VB.Image Image1 
            Height          =   375
            Index           =   1
            Left            =   75
            Picture         =   "frmisicpa.frx":189A
            Stretch         =   -1  'True
            Top             =   40
            Width           =   375
         End
      End
      Begin VB.TextBox txtcardno 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   2970
         Width           =   1995
      End
      Begin VB.TextBox txtname 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   3285
         Width           =   4875
      End
      Begin VB.TextBox txtcycle 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   3960
         Width           =   1995
      End
      Begin VB.ComboBox cbosts 
         Height          =   315
         ItemData        =   "frmisicpa.frx":3134
         Left            =   1530
         List            =   "frmisicpa.frx":3144
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   4320
         Width           =   2220
      End
      Begin VB.TextBox txtcollect 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   4680
         Width           =   2220
      End
      Begin VB.TextBox txtplace 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   5040
         Width           =   2220
      End
      Begin VB.TextBox txtagency 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   5400
         Width           =   2220
      End
      Begin VB.Frame Frame15 
         Appearance      =   0  'Flat
         BackColor       =   &H002F735C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   2
         Left            =   270
         TabIndex        =   36
         Top             =   5805
         Width           =   6090
         Begin VB.Image Image1 
            Height          =   375
            Index           =   2
            Left            =   75
            Picture         =   "frmisicpa.frx":315A
            Stretch         =   -1  'True
            Top             =   40
            Width           =   375
         End
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
            TabIndex        =   37
            Top             =   90
            Width           =   3255
         End
      End
      Begin VB.TextBox txtperiodpay 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   7650
         Width           =   2175
      End
      Begin VB.TextBox label5 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3780
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   6570
         Width           =   1995
      End
      Begin VB.TextBox label8 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3780
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   7335
         Width           =   1995
      End
      Begin TDBDate6Ctl.TDBDate dtpropsal 
         Height          =   255
         Left            =   1485
         TabIndex        =   53
         Top             =   945
         Width           =   2250
         _Version        =   65536
         _ExtentX        =   3969
         _ExtentY        =   450
         Calendar        =   "frmisicpa.frx":49F4
         Caption         =   "frmisicpa.frx":4B0C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmisicpa.frx":4B78
         Keys            =   "frmisicpa.frx":4B96
         Spin            =   "frmisicpa.frx":4BF4
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
         TabIndex        =   54
         Top             =   3645
         Width           =   2250
         _Version        =   65536
         _ExtentX        =   3969
         _ExtentY        =   450
         Calendar        =   "frmisicpa.frx":4C1C
         Caption         =   "frmisicpa.frx":4D34
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmisicpa.frx":4DA0
         Keys            =   "frmisicpa.frx":4DBE
         Spin            =   "frmisicpa.frx":4E1C
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
         TabIndex        =   55
         Top             =   6660
         Width           =   2160
         _Version        =   65536
         _ExtentX        =   3810
         _ExtentY        =   450
         Calculator      =   "frmisicpa.frx":4E44
         Caption         =   "frmisicpa.frx":4E64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmisicpa.frx":4ED0
         Keys            =   "frmisicpa.frx":4EEE
         Spin            =   "frmisicpa.frx":4F38
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
      Begin TDBNumber6Ctl.TDBNumber txtdownpayment 
         Height          =   255
         Left            =   1530
         TabIndex        =   56
         Top             =   7020
         Width           =   2160
         _Version        =   65536
         _ExtentX        =   3810
         _ExtentY        =   450
         Calculator      =   "frmisicpa.frx":4F60
         Caption         =   "frmisicpa.frx":4F80
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmisicpa.frx":4FEC
         Keys            =   "frmisicpa.frx":500A
         Spin            =   "frmisicpa.frx":5054
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
         TabIndex        =   57
         Top             =   7335
         Width           =   2160
         _Version        =   65536
         _ExtentX        =   3810
         _ExtentY        =   450
         Calculator      =   "frmisicpa.frx":507C
         Caption         =   "frmisicpa.frx":509C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmisicpa.frx":5108
         Keys            =   "frmisicpa.frx":5126
         Spin            =   "frmisicpa.frx":5170
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin TDBNumber6Ctl.TDBNumber txtprincipal 
         Height          =   255
         Left            =   1530
         TabIndex        =   58
         Top             =   8055
         Width           =   2160
         _Version        =   65536
         _ExtentX        =   3810
         _ExtentY        =   450
         Calculator      =   "frmisicpa.frx":5198
         Caption         =   "frmisicpa.frx":51B8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmisicpa.frx":5224
         Keys            =   "frmisicpa.frx":5242
         Spin            =   "frmisicpa.frx":528C
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
      Begin TDBNumber6Ctl.TDBNumber txtbalance 
         Height          =   255
         Left            =   1530
         TabIndex        =   59
         Top             =   6300
         Width           =   2160
         _Version        =   65536
         _ExtentX        =   3810
         _ExtentY        =   450
         Calculator      =   "frmisicpa.frx":52B4
         Caption         =   "frmisicpa.frx":52D4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmisicpa.frx":5340
         Keys            =   "frmisicpa.frx":535E
         Spin            =   "frmisicpa.frx":53A8
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Region"
         Height          =   240
         Index           =   0
         Left            =   315
         TabIndex        =   81
         Top             =   630
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Proposal Date"
         Height          =   240
         Index           =   1
         Left            =   315
         TabIndex        =   80
         Top             =   990
         Width           =   1140
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Reffno"
         Height          =   240
         Index           =   2
         Left            =   360
         TabIndex        =   79
         Top             =   1305
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Product"
         Height          =   240
         Index           =   3
         Left            =   360
         TabIndex        =   78
         Top             =   1665
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Arrangement"
         Height          =   240
         Index           =   4
         Left            =   360
         TabIndex        =   77
         Top             =   2025
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Card no"
         Height          =   240
         Index           =   5
         Left            =   360
         TabIndex        =   76
         Top             =   3015
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "cust name"
         Height          =   240
         Index           =   6
         Left            =   360
         TabIndex        =   75
         Top             =   3330
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Card Open"
         Height          =   240
         Index           =   7
         Left            =   360
         TabIndex        =   74
         Top             =   3645
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cycle Dlq"
         Height          =   240
         Index           =   8
         Left            =   360
         TabIndex        =   73
         Top             =   4005
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "status card"
         Height          =   240
         Index           =   9
         Left            =   360
         TabIndex        =   72
         Top             =   4365
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
         Caption         =   "placement"
         Height          =   240
         Index           =   11
         Left            =   405
         TabIndex        =   70
         Top             =   5085
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Agency name"
         Height          =   240
         Index           =   12
         Left            =   405
         TabIndex        =   69
         Top             =   5445
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         Height          =   285
         Index           =   13
         Left            =   360
         TabIndex        =   68
         Top             =   6345
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Payment"
         Height          =   330
         Index           =   14
         Left            =   315
         TabIndex        =   67
         Top             =   6705
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "down Payment"
         Height          =   195
         Index           =   15
         Left            =   315
         TabIndex        =   66
         Top             =   7020
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Future Payment"
         Height          =   195
         Index           =   16
         Left            =   315
         TabIndex        =   65
         Top             =   7380
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment period month"
         Height          =   465
         Index           =   17
         Left            =   315
         TabIndex        =   64
         Top             =   7695
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Principal"
         Height          =   285
         Index           =   19
         Left            =   315
         TabIndex        =   63
         Top             =   8100
         Width           =   1230
      End
      Begin VB.Label Label4 
         BackColor       =   &H00B1FDD5&
         Caption         =   ")*  D=SETTLEMENT R=RESCHEDULE X=PAID OFF"
         ForeColor       =   &H000000FF&
         Height          =   510
         Left            =   3600
         TabIndex        =   62
         Top             =   1215
         Width           =   2895
      End
      Begin VB.Label Label6 
         Caption         =   "Balance di database"
         Height          =   285
         Left            =   3735
         TabIndex        =   61
         Top             =   6300
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Principal di database"
         Height          =   285
         Left            =   3780
         TabIndex        =   60
         Top             =   6975
         Width           =   1590
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00B1FDD5&
      BorderStyle     =   0  'None
      Caption         =   "Frame13"
      ForeColor       =   &H80000008&
      Height          =   9135
      Left            =   6930
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.CheckBox chkcek 
         BackColor       =   &H00B1FDD5&
         Caption         =   "Reject"
         Height          =   285
         Index           =   2
         Left            =   2610
         TabIndex        =   31
         Top             =   5220
         Width           =   1995
      End
      Begin VB.CheckBox chkcek 
         BackColor       =   &H00B1FDD5&
         Caption         =   "Approval"
         Height          =   285
         Index           =   1
         Left            =   1500
         TabIndex        =   30
         Top             =   5220
         Width           =   2040
      End
      Begin VB.CheckBox chkcek 
         BackColor       =   &H00B1FDD5&
         Caption         =   "Verify"
         Height          =   285
         Index           =   0
         Left            =   1530
         TabIndex        =   29
         Top             =   5220
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Frame Frame15 
         Appearance      =   0  'Flat
         BackColor       =   &H002F735C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Index           =   5
         Left            =   -45
         TabIndex        =   10
         Top             =   90
         Width           =   4785
         Begin VB.Image Image1 
            Height          =   375
            Index           =   5
            Left            =   75
            Picture         =   "frmisicpa.frx":53D0
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
            TabIndex        =   11
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   6
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
            TabIndex        =   7
            Top             =   90
            Width           =   3255
         End
         Begin VB.Image Image1 
            Height          =   375
            Index           =   4
            Left            =   75
            Picture         =   "frmisicpa.frx":6C6A
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
         TabIndex        =   5
         Top             =   2970
         Width           =   1995
      End
      Begin VB.TextBox txtreason 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1530
         TabIndex        =   4
         Top             =   3285
         Width           =   3120
      End
      Begin VB.TextBox txtpaymenthandle 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1530
         TabIndex        =   3
         Top             =   4005
         Width           =   2220
      End
      Begin VB.TextBox txtnodlq 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1530
         TabIndex        =   2
         Top             =   3645
         Width           =   2220
      End
      Begin VB.TextBox txtjust 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   780
         Left            =   1530
         TabIndex        =   1
         Top             =   4365
         Width           =   3165
      End
      Begin TDBNumber6Ctl.TDBNumber txtcharge 
         Height          =   255
         Left            =   1845
         TabIndex        =   21
         Top             =   585
         Width           =   2160
         _Version        =   65536
         _ExtentX        =   3810
         _ExtentY        =   450
         Calculator      =   "frmisicpa.frx":8504
         Caption         =   "frmisicpa.frx":8524
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmisicpa.frx":8590
         Keys            =   "frmisicpa.frx":85AE
         Spin            =   "frmisicpa.frx":85F8
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin TDBNumber6Ctl.TDBNumber txtdiscount 
         Height          =   255
         Left            =   1845
         TabIndex        =   22
         Top             =   945
         Width           =   2160
         _Version        =   65536
         _ExtentX        =   3810
         _ExtentY        =   450
         Calculator      =   "frmisicpa.frx":8620
         Caption         =   "frmisicpa.frx":8640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmisicpa.frx":86AC
         Keys            =   "frmisicpa.frx":86CA
         Spin            =   "frmisicpa.frx":8714
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
         ValueVT         =   1245189
         Value           =   0
         MaxValueVT      =   1701642245
         MinValueVT      =   3801093
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   660
         Index           =   0
         Left            =   2790
         TabIndex        =   23
         Tag             =   "0"
         Top             =   6660
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
         Picture         =   "frmisicpa.frx":873C
         AutoSize        =   1
         Alignment       =   8
         PictureAlignment=   1
      End
      Begin Threed.SSCommand SSCommand1 
         Cancel          =   -1  'True
         Height          =   660
         Index           =   1
         Left            =   3600
         TabIndex        =   25
         Top             =   6660
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
         Picture         =   "frmisicpa.frx":8C6F
         AutoSize        =   1
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin TDBDate6Ctl.TDBDate dtpelunasan 
         Height          =   255
         Left            =   1485
         TabIndex        =   26
         Top             =   5580
         Width           =   1620
         _Version        =   65536
         _ExtentX        =   2857
         _ExtentY        =   450
         Calendar        =   "frmisicpa.frx":92D4
         Caption         =   "frmisicpa.frx":93EC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmisicpa.frx":9458
         Keys            =   "frmisicpa.frx":9476
         Spin            =   "frmisicpa.frx":94D4
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
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         Height          =   285
         Left            =   3825
         TabIndex        =   28
         Top             =   7335
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Lunas"
         Height          =   240
         Index           =   21
         Left            =   90
         TabIndex        =   27
         Top             =   5625
         Width           =   1635
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Save"
         Height          =   285
         Left            =   2880
         TabIndex        =   24
         Top             =   7335
         Width           =   510
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Charge"
         Height          =   240
         Index           =   39
         Left            =   0
         TabIndex        =   20
         Top             =   630
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Amount"
         Height          =   240
         Index           =   38
         Left            =   0
         TabIndex        =   19
         Top             =   990
         Width           =   1140
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "From o/s balance %"
         Height          =   330
         Index           =   37
         Left            =   45
         TabIndex        =   18
         Top             =   1305
         Width           =   2220
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "principal (%) from "
         Height          =   240
         Index           =   36
         Left            =   45
         TabIndex        =   17
         Top             =   1665
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "occupation"
         Height          =   240
         Index           =   34
         Left            =   0
         TabIndex        =   16
         Top             =   3015
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "reason"
         Height          =   240
         Index           =   33
         Left            =   0
         TabIndex        =   15
         Top             =   3330
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "payment handle by"
         Height          =   240
         Index           =   30
         Left            =   0
         TabIndex        =   14
         Top             =   4050
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Justification"
         Height          =   240
         Index           =   29
         Left            =   0
         TabIndex        =   13
         Top             =   4365
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "no of DlQ"
         Height          =   240
         Index           =   28
         Left            =   45
         TabIndex        =   12
         Top             =   3690
         Width           =   1230
      End
   End
End
Attribute VB_Name = "frmisicpa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Cboaoc_Click(Index As Integer)
'  Select Case Index
'    Case 1
'        Cboaoc(0).ListIndex = Cboaoc(Index).ListIndex
'    Case 0
'        Cboaoc(1).ListIndex = Cboaoc(Index).ListIndex
'End Select
'End Sub

Private Sub chkcek_Click(Index As Integer)
Select Case Index
Case 2
If chkcek(2).Value = 1 Then
         
        chkcek(1).Value = 0
Else
  chkcek(2).Value = 0
End If
Case 1
If chkcek(1).Value Then
        
        chkcek(2).Value = 0
        Else
        chkcek(1).Value = 0
End If
End Select
End Sub

Private Sub Form_Load()
    chkcek(0).Visible = False
    chkcek(1).Visible = False
    chkcek(2).Visible = False
If UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
    chkcek(0).Visible = False
    chkcek(1).Visible = False
    SSCommand1(0).Visible = False
    SSCommand1(1).Visible = True
    Label2.Visible = False
    Label3.Visible = True
ElseIf UCase(MDIForm1.txtlevel.Text) = "TEAMLEADER" Then
    chkcek(0).Visible = False
Else
    chkcek(0).Visible = True
    chkcek(1).Visible = True
    chkcek(2).Visible = True
End If
cbosts.ListIndex = 1

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmCPA.showlist
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

Private Sub SSCommand1_Click(Index As Integer)
Select Case Index
        Case 0
            Select Case SSCommand1(0).tag
            Case "1"
                    chkcek(0).Value = 1
                    STRSQL = "insert into mandiri.tblCpa(vcustid,vregion,dpropsal,vreffno,vproduct,varragement,vcardsts,nttlpayment,ndownpay,nfuturepay,ncharge,"
                    STRSQL = STRSQL + " ndiscountamt,vosbalance,vosprincipal ,dtglinsert,dtgllastupdate,dtglpelunasan,vjust,vcustname,voccupation,vreason,vnodlq,vpaymenthandle,intverify,intapprovel,agency,nbalance,nprincipal) values ( "
                    STRSQL = STRSQL + "'" + FrmCC_Colection.lblcustid.Caption + "','" + txtregion.Text + "',"
                    STRSQL = STRSQL + IIf(dtpropsal.ValueIsNull, "null", "'" + Format(dtpropsal.Value, "yyyy-mm-dd") + "'") + ", '" + txtreff.Text + "','" + txtproduct.Text + "' ,"
                    STRSQL = STRSQL + "'" + txtarrangement.Text + "','" + cbosts.Text + "'," + CStr(lblLastPay.Value) + "," + CStr(txtdownpayment.Value) + ","
                    STRSQL = STRSQL + "" + CStr(Val(txtfuture.Value)) + "," + CStr((txtcharge.Value)) + "," + CStr(txtdiscount.Value) + ",'" + txtfrombalancepersen.Text + "','" + txtpersenprincipal.Text + "',"
                    STRSQL = STRSQL + "'" + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") + " " + Format(Now, "hh:mm:ss") + "','" + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") + " " + Format(Now, "hh:mm:ss") + "',"
                    STRSQL = STRSQL + IIf(dtpelunasan.ValueIsNull, "null", " '" + Format(dtpelunasan.Value, "yyyy-mm-dd") + "'") + ",'" + txtjust.Text + "','" + FrmCC_Colection.lblnama.Caption + "',"
                    STRSQL = STRSQL + "'" + txtoccupation.Text + "', '" + txtreason.Text + "','" + txtnodlq.Text + "','" + txtpaymenthandle.Text + "',"
                    STRSQL = STRSQL + CStr(IIf(chkcek(0).Value, 1, 0)) + "," + CStr(IIf(chkcek(1).Value, 1, 0)) + ",'" + txtagency.Text + "',"
                    STRSQL = STRSQL + "" + CStr(Val(txtbalance.Value)) + "," + CStr((txtprincipal.Value)) + ")"
                    M_OBJCONN.Execute (STRSQL)
                    If chkcek(1).Value = 1 Then
                        MsgBox "CPA approved ", vbInformation + vbOKOnly, "Pesan"
                    ElseIf chkcek(2).Value = 1 Then
                        MsgBox "CPA Rejected ", vbInformation + vbOKOnly, "Pesan"
                    Else
                        MsgBox "data telah di simpan", vbInformation + vbOKOnly, "Pesan"
                    End If
                    
                    
                    STRSQL = "update mandiri.mgm set intverify=0,vnameverify='',stscpa=1,intapprovel=0 where custid ='" + FrmCC_Colection.lblcustid.Caption + "'"
                    M_OBJCONN.Execute (STRSQL)
                                
                    If chkcek(0).Value = 1 Then
                        STRSQL = "update mandiri.mgm set stscpa=0,intverify=1,vnameverify='" + MDIForm1.TxtUsername.Text + "' where custid ='" + FrmCC_Colection.lblcustid.Caption + "'"
                        M_OBJCONN.Execute (STRSQL)
                    End If
                    
                    If chkcek(1).Value = 1 Then
                        STRSQL = "update mandiri.mgm set intapprovel=1,intverify=1 ,tglstscpa= '" + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") + " " + Format(Now, "hh:mm:ss") + "',  vnameapprovel='" + MDIForm1.TxtUsername.Text + "' where custid ='" + FrmCC_Colection.lblcustid.Caption + "'"
                        M_OBJCONN.Execute (STRSQL)
                    ElseIf chkcek(1).Value = 0 And chkcek(2).Value = 1 Then
                        STRSQL = "update mandiri.mgm set intapprovel=0 ,vnameapprovel='',tglstscpa= '" + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") + " " + Format(Now, "hh:mm:ss") + "', resultcpa='GAGAL' where custid ='" + FrmCC_Colection.lblcustid.Caption + "'"
                        M_OBJCONN.Execute (STRSQL)
                    End If
                    SSCommand1_Click (1)
            Case "2"
                    STRSQL = "update mandiri.tblcpa set  dtgllastupdate= '" + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") + " " + Format(Now, "hh:mm:ss") + "' ,nttlpayment='" + CStr(lblLastPay.Value) + "',ndownpay='" + CStr(txtdownpayment.Value) + "',"
                    STRSQL = STRSQL + "vregion='" + txtregion.Text + "',dpropsal=" + IIf(dtpropsal.ValueIsNull, "null", " '" + Format(dtpropsal.Value, "yyyy-mm-dd") + "'") + ",vreffno='" + txtreff.Text + "',vproduct ='" + txtproduct.Text + "',"
                    STRSQL = STRSQL + "varragement='" + txtarrangement.Text + "', vcardsts='" + cbosts.Text + "',nfuturepay =" + CStr(Val(txtfuture.Value)) + ",ncharge='" + CStr(Val(txtcharge.Value)) + "',"
                    STRSQL = STRSQL + " ndiscountamt=" + CStr(txtdiscount.Value) + ",vosbalance='" + txtfrombalancepersen.Text + "',vosprincipal='" + txtpersenprincipal.Text + "',"
                    STRSQL = STRSQL + "dtglpelunasan=" + IIf(dtpelunasan.ValueIsNull, "null", " '" + Format(dtpelunasan.Value, "yyyy-mm-dd") + "'") + ",vjust='" + txtjust.Text + "', agency='" + txtagency.Text + "',"
                    STRSQL = STRSQL + "voccupation='" + txtoccupation.Text + "',vreason='" + txtreason.Text + "',vnodlq='" + txtnodlq.Text + "',vpaymenthandle='" + txtpaymenthandle.Text + "',intverify=" + IIf(chkcek(0).Value, "1", "0") + ",intapprovel=" + IIf(chkcek(1).Value, "1", "0") + ","
                    STRSQL = STRSQL + "nbalance=" + CStr(txtbalance.Value) + ",nprincipal=" + CStr(txtprincipal.Value) + ""
                    STRSQL = STRSQL + " where nid='" + frmCPA.LstCpa.SelectedItem.Text + "'"
                    M_OBJCONN.Execute (STRSQL)
                    If UCase(MDIForm1.txtlevel.Text) = "AGENT" Then Exit Sub
                    If chkcek(0).Value = 1 Then
                        STRSQL = "update mandiri.mgm set intverify=1,vnameverify='" + MDIForm1.TxtUsername.Text + "' where custid ='" + FrmCC_Colection.lblcustid.Caption + "' and  (vnameverify is null  or vnameverify='')"
                        M_OBJCONN.Execute (STRSQL)
                    ElseIf chkcek(0).Value = 0 Then
                        STRSQL = "update mandiri.mgm set intverify=0,vnameverify='' where custid ='" + FrmCC_Colection.lblcustid.Caption + "'"
                        M_OBJCONN.Execute (STRSQL)
                    End If
                    
                    If chkcek(1).Value = 1 Then
                        MsgBox "CPA approved ", vbInformation + vbOKOnly, "Pesan"
                    ElseIf chkcek(2).Value = 1 Then
                        MsgBox "CPA Rejected ", vbInformation + vbOKOnly, "Pesan"
                    Else
                    
                        MsgBox "data telah di update", vbInformation + vbOKOnly, "Pesan"
                    End If
                    
                    If UCase(MDIForm1.txtlevel.Text) = "TEAMLEADER" Then Exit Sub
                    If chkcek(1).Value = 1 Then
                        STRSQL = "update mandiri.mgm set intapprovel=1,intverify=1,resultcpa='SUKSES', tglstscpa= '" + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") + " " + Format(Now, "hh:mm:ss") + "',vnameapprovel='" + MDIForm1.TxtUsername.Text + "'   where custid ='" + FrmCC_Colection.lblcustid.Caption + "' and (vnameverify<>'' or vnameverify is not null)  "
                        M_OBJCONN.Execute (STRSQL)
                    ElseIf chkcek(1).Value = 0 And chkcek(2).Value = 1 Then
                        STRSQL = "update mandiri.mgm set intapprovel=0 ,vnameapprovel='',tglstscpa= '" + Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") + " " + Format(Now, "hh:mm:ss") + "', resultcpa='GAGAL' where custid ='" + FrmCC_Colection.lblcustid.Caption + "'"
                        M_OBJCONN.Execute (STRSQL)
                    Else
                        STRSQL = "update mandiri.mgm set intapprovel=0 ,vnameapprovel='' where custid ='" + FrmCC_Colection.lblcustid.Caption + "'"
                        M_OBJCONN.Execute (STRSQL)
                    End If
                    
        End Select
      Case 1
        frmCPA.showlist
        Unload Me
End Select

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

