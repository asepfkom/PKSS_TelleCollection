VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmCC_ColectionCARD 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10605
   ClientLeft      =   210
   ClientTop       =   60
   ClientWidth     =   18330
   ControlBox      =   0   'False
   DrawMode        =   14  'Copy Pen
   Icon            =   "frmCC_Colection_RITCARD.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10605
   ScaleWidth      =   18330
   Begin VB.Frame Frame11 
      Caption         =   "Other Card"
      Height          =   1800
      Left            =   165
      TabIndex        =   306
      Top             =   7695
      Width           =   7635
      Begin MSComctlLib.ListView LstDoubleId 
         Height          =   1410
         Left            =   75
         TabIndex        =   307
         Top             =   240
         Width           =   7500
         _ExtentX        =   13229
         _ExtentY        =   2487
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Remarks   (mm-dd-yyyy)"
      Height          =   1725
      Left            =   60
      TabIndex        =   303
      Top             =   3930
      Width           =   13995
      Begin MSComctlLib.ListView listview1 
         Height          =   1455
         Index           =   1
         Left            =   75
         TabIndex        =   304
         Top             =   180
         Width           =   13845
         _ExtentX        =   24421
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   3885
      Left            =   105
      TabIndex        =   182
      Top             =   -15
      Width           =   18180
      _ExtentX        =   32068
      _ExtentY        =   6853
      _Version        =   196610
      Font3D          =   1
      ForeColor       =   12583104
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Data Nasabah"
      Alignment       =   2
      ShadowStyle     =   1
      Begin VB.Frame Frame1 
         Height          =   3810
         Left            =   120
         TabIndex        =   183
         Top             =   255
         Width           =   18150
         Begin VB.TextBox txtECAdd 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Left            =   15165
            TabIndex        =   184
            Top             =   2520
            Width           =   2880
         End
         Begin RichTextLib.RichTextBox lblOfficeAddr 
            Height          =   675
            Left            =   825
            TabIndex        =   185
            Top             =   2235
            Width           =   3270
            _ExtentX        =   5768
            _ExtentY        =   1191
            _Version        =   393217
            BackColor       =   16777215
            BorderStyle     =   0
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"frmCC_Colection_RITCARD.frx":000C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin TDBDate6Ctl.TDBDate lblDate 
            Height          =   285
            Left            =   1020
            TabIndex        =   186
            Top             =   1200
            Visible         =   0   'False
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   503
            Calendar        =   "frmCC_Colection_RITCARD.frx":0091
            Caption         =   "frmCC_Colection_RITCARD.frx":01A9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD.frx":0215
            Keys            =   "frmCC_Colection_RITCARD.frx":0233
            Spin            =   "frmCC_Colection_RITCARD.frx":0291
            AlignHorizontal =   2
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   1
            Value           =   3.54031216694028E-316
            CenturyMode     =   0
         End
         Begin RichTextLib.RichTextBox lblAddr 
            Height          =   690
            Left            =   810
            TabIndex        =   187
            Top             =   1515
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   1217
            _Version        =   393217
            BackColor       =   16777215
            BorderStyle     =   0
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"frmCC_Colection_RITCARD.frx":02B9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin TDBDate6Ctl.TDBDate lblOpenDate 
            Height          =   315
            Left            =   5145
            TabIndex        =   188
            Top             =   1230
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            Calendar        =   "frmCC_Colection_RITCARD.frx":033E
            Caption         =   "frmCC_Colection_RITCARD.frx":0456
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD.frx":04C2
            Keys            =   "frmCC_Colection_RITCARD.frx":04E0
            Spin            =   "frmCC_Colection_RITCARD.frx":053E
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
         Begin TDBDate6Ctl.TDBDate lblBD 
            Height          =   315
            Left            =   5145
            TabIndex        =   189
            Top             =   1590
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            Calendar        =   "frmCC_Colection_RITCARD.frx":0566
            Caption         =   "frmCC_Colection_RITCARD.frx":067E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD.frx":06EA
            Keys            =   "frmCC_Colection_RITCARD.frx":0708
            Spin            =   "frmCC_Colection_RITCARD.frx":0766
            AlignHorizontal =   2
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   1
            Value           =   1.07202956713409E-317
            CenturyMode     =   0
         End
         Begin TDBNumber6Ctl.TDBNumber lblLimit 
            Height          =   315
            Left            =   5145
            TabIndex        =   190
            Top             =   870
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            Calculator      =   "frmCC_Colection_RITCARD.frx":078E
            Caption         =   "frmCC_Colection_RITCARD.frx":07AE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD.frx":081A
            Keys            =   "frmCC_Colection_RITCARD.frx":0838
            Spin            =   "frmCC_Colection_RITCARD.frx":0882
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   1245189
            Value           =   0
            MaxValueVT      =   1701642245
            MinValueVT      =   3801093
         End
         Begin TDBNumber6Ctl.TDBNumber lblAmount 
            Height          =   315
            Left            =   5145
            TabIndex        =   191
            Top             =   150
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            Calculator      =   "frmCC_Colection_RITCARD.frx":08AA
            Caption         =   "frmCC_Colection_RITCARD.frx":08CA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD.frx":0936
            Keys            =   "frmCC_Colection_RITCARD.frx":0954
            Spin            =   "frmCC_Colection_RITCARD.frx":099E
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   1245189
            Value           =   0
            MaxValueVT      =   1701642245
            MinValueVT      =   3801093
         End
         Begin TDBNumber6Ctl.TDBNumber lblLastPay 
            Height          =   270
            Left            =   5145
            TabIndex        =   192
            Top             =   2310
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   476
            Calculator      =   "frmCC_Colection_RITCARD.frx":09C6
            Caption         =   "frmCC_Colection_RITCARD.frx":09E6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD.frx":0A52
            Keys            =   "frmCC_Colection_RITCARD.frx":0A70
            Spin            =   "frmCC_Colection_RITCARD.frx":0ABA
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   1245189
            Value           =   0
            MaxValueVT      =   1701642245
            MinValueVT      =   3801093
         End
         Begin TDBDate6Ctl.TDBDate lblPayDt 
            Height          =   315
            Left            =   5145
            TabIndex        =   193
            Top             =   1950
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            Calendar        =   "frmCC_Colection_RITCARD.frx":0AE2
            Caption         =   "frmCC_Colection_RITCARD.frx":0BFA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD.frx":0C66
            Keys            =   "frmCC_Colection_RITCARD.frx":0C84
            Spin            =   "frmCC_Colection_RITCARD.frx":0CE2
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   1
            Value           =   3.54027066542603E-316
            CenturyMode     =   0
         End
         Begin TDBNumber6Ctl.TDBNumber Woafter 
            Height          =   315
            Left            =   5145
            TabIndex        =   194
            Top             =   2625
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   556
            Calculator      =   "frmCC_Colection_RITCARD.frx":0D0A
            Caption         =   "frmCC_Colection_RITCARD.frx":0D2A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD.frx":0D96
            Keys            =   "frmCC_Colection_RITCARD.frx":0DB4
            Spin            =   "frmCC_Colection_RITCARD.frx":0DFE
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
            ReadOnly        =   -1
            Separator       =   ","
            ShowContextMenu =   -1
            ValueVT         =   1245189
            Value           =   0
            MaxValueVT      =   1701642245
            MinValueVT      =   3801093
         End
         Begin TDBNumber6Ctl.TDBNumber txtPrinciple_A 
            Height          =   300
            Left            =   6435
            TabIndex        =   195
            Top             =   510
            Visible         =   0   'False
            Width           =   180
            _Version        =   65536
            _ExtentX        =   317
            _ExtentY        =   529
            Calculator      =   "frmCC_Colection_RITCARD.frx":0E26
            Caption         =   "frmCC_Colection_RITCARD.frx":0E46
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD.frx":0EB2
            Keys            =   "frmCC_Colection_RITCARD.frx":0ED0
            Spin            =   "frmCC_Colection_RITCARD.frx":0F1A
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
            ForeColor       =   0
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
         Begin TDBMask6Ctl.TDBMask txtHomeNo2 
            Height          =   300
            Left            =   8145
            TabIndex        =   196
            Top             =   555
            Visible         =   0   'False
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   529
            Caption         =   "frmCC_Colection_RITCARD.frx":0F42
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":0FAE
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   16777215
            BorderStyle     =   0
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "____________________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask AHome2 
            Height          =   285
            Left            =   7455
            TabIndex        =   197
            Top             =   570
            Width           =   555
            _Version        =   65536
            _ExtentX        =   979
            _ExtentY        =   503
            Caption         =   "frmCC_Colection_RITCARD.frx":0FF0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":105C
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   16777215
            BorderStyle     =   0
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "____"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtHomeNo1 
            Height          =   300
            Left            =   8145
            TabIndex        =   198
            Top             =   180
            Visible         =   0   'False
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   529
            Caption         =   "frmCC_Colection_RITCARD.frx":109E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":110A
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   16777215
            BorderStyle     =   0
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "____________________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask AHome1 
            Height          =   300
            Left            =   7440
            TabIndex        =   199
            Top             =   195
            Width           =   540
            _Version        =   65536
            _ExtentX        =   952
            _ExtentY        =   529
            Caption         =   "frmCC_Colection_RITCARD.frx":114C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":11B8
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   16777215
            BorderStyle     =   0
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "____"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtOfficeNo1 
            Height          =   315
            Left            =   8145
            TabIndex        =   200
            Top             =   930
            Visible         =   0   'False
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   556
            Caption         =   "frmCC_Colection_RITCARD.frx":11FA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":1266
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   16777215
            BorderStyle     =   0
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   0
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "____________________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask AOffice1 
            Height          =   285
            Left            =   7440
            TabIndex        =   201
            Top             =   915
            Width           =   585
            _Version        =   65536
            _ExtentX        =   1032
            _ExtentY        =   503
            Caption         =   "frmCC_Colection_RITCARD.frx":12A8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":1314
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   16777215
            BorderStyle     =   0
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "____"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtOfficeNo2 
            Height          =   270
            Left            =   8145
            TabIndex        =   202
            Top             =   1275
            Visible         =   0   'False
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   476
            Caption         =   "frmCC_Colection_RITCARD.frx":1356
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":13C2
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   16777215
            BorderStyle     =   0
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "____________________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask AOffice2 
            Height          =   270
            Left            =   7440
            TabIndex        =   203
            Top             =   1305
            Width           =   585
            _Version        =   65536
            _ExtentX        =   1032
            _ExtentY        =   476
            Caption         =   "frmCC_Colection_RITCARD.frx":1404
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":1470
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   16777215
            BorderStyle     =   0
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "____"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtMobileNo1 
            Height          =   285
            Left            =   8145
            TabIndex        =   204
            Top             =   1650
            Visible         =   0   'False
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   503
            Caption         =   "frmCC_Colection_RITCARD.frx":14B2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":151E
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   16777215
            BorderStyle     =   0
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "____________________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtMobileNo2 
            Height          =   285
            Left            =   8145
            TabIndex        =   205
            Top             =   1995
            Visible         =   0   'False
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   503
            Caption         =   "frmCC_Colection_RITCARD.frx":1560
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":15CC
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   16777215
            BorderStyle     =   0
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "____________________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtHomeNo1A 
            Height          =   300
            Left            =   8145
            TabIndex        =   206
            Top             =   180
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   529
            Caption         =   "frmCC_Colection_RITCARD.frx":160E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":167A
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   16777215
            BorderStyle     =   0
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "____________________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtHomeNo2A 
            Height          =   300
            Left            =   8145
            TabIndex        =   207
            Top             =   555
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   529
            Caption         =   "frmCC_Colection_RITCARD.frx":16BC
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":1728
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   16777215
            BorderStyle     =   0
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "____________________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtOfficeNo1A 
            Height          =   315
            Left            =   8145
            TabIndex        =   208
            Top             =   915
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   556
            Caption         =   "frmCC_Colection_RITCARD.frx":176A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":17D6
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   16777215
            BorderStyle     =   0
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   0
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "____________________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtOfficeNo2A 
            Height          =   315
            Left            =   8130
            TabIndex        =   209
            Top             =   1275
            Width           =   1725
            _Version        =   65536
            _ExtentX        =   3043
            _ExtentY        =   556
            Caption         =   "frmCC_Colection_RITCARD.frx":1818
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":1884
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   16777215
            BorderStyle     =   0
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "____________________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtMobileNo1A 
            Height          =   285
            Left            =   8145
            TabIndex        =   210
            Top             =   1650
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   503
            Caption         =   "frmCC_Colection_RITCARD.frx":18C6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":1932
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   16777215
            BorderStyle     =   0
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "____________________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtMobileNo2A 
            Height          =   285
            Left            =   8145
            TabIndex        =   211
            Top             =   1995
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   503
            Caption         =   "frmCC_Colection_RITCARD.frx":1974
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":19E0
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   16777215
            BorderStyle     =   0
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "____________________"
            Value           =   ""
         End
         Begin RichTextLib.RichTextBox TxtEC 
            Height          =   315
            Left            =   15165
            TabIndex        =   212
            Top             =   1830
            Width           =   2760
            _ExtentX        =   4868
            _ExtentY        =   556
            _Version        =   393217
            BackColor       =   16777215
            Enabled         =   -1  'True
            Appearance      =   0
            TextRTF         =   $"frmCC_Colection_RITCARD.frx":1A22
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin TDBMask6Ctl.TDBMask txtECno 
            Height          =   330
            Left            =   15180
            TabIndex        =   213
            Top             =   2160
            Visible         =   0   'False
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   582
            Caption         =   "frmCC_Colection_RITCARD.frx":1AA7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":1B13
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
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
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__________________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtECnoA 
            Height          =   330
            Left            =   15165
            TabIndex        =   214
            Top             =   2160
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   582
            Caption         =   "frmCC_Colection_RITCARD.frx":1B55
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":1BC1
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
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
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__________________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtFaxAdd1 
            Height          =   330
            Left            =   12240
            TabIndex        =   215
            Top             =   1785
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   582
            Caption         =   "frmCC_Colection_RITCARD.frx":1C03
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":1C6F
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
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
            Format          =   "&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "____________________________________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtFaxAdd2 
            Height          =   330
            Left            =   12240
            TabIndex        =   216
            Top             =   2145
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   582
            Caption         =   "frmCC_Colection_RITCARD.frx":1CB1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":1D1D
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
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
            Format          =   "&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__________________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask AFaxAdd 
            Height          =   330
            Index           =   4
            Left            =   11430
            TabIndex        =   217
            Top             =   1785
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   582
            Caption         =   "frmCC_Colection_RITCARD.frx":1D5F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":1DCB
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   1
            AutoConvert     =   1
            BackColor       =   16777215
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   0
            Format          =   "[9999]"
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
            ReadOnly        =   -1
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "[____]"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask AFaxAdd 
            Height          =   330
            Index           =   5
            Left            =   11430
            TabIndex        =   218
            Top             =   2145
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   582
            Caption         =   "frmCC_Colection_RITCARD.frx":1E0D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":1E79
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   1
            AutoConvert     =   1
            BackColor       =   16777215
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   0
            Format          =   "[9999]"
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
            ReadOnly        =   -1
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "[____]"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtMobileAdd1 
            Height          =   330
            Left            =   11445
            TabIndex        =   219
            Top             =   2490
            Visible         =   0   'False
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   582
            Caption         =   "frmCC_Colection_RITCARD.frx":1EBB
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":1F27
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
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
            Format          =   "&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__________________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtMobileAdd2 
            Height          =   330
            Left            =   11445
            TabIndex        =   220
            Top             =   2835
            Visible         =   0   'False
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   582
            Caption         =   "frmCC_Colection_RITCARD.frx":1F69
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":1FD5
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
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
            Format          =   "&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__________________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtOfficeAdd1 
            Height          =   330
            Left            =   12255
            TabIndex        =   221
            Top             =   1065
            Visible         =   0   'False
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   582
            Caption         =   "frmCC_Colection_RITCARD.frx":2017
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":2083
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
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
            Format          =   "&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__________________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtOfficeAdd2 
            Height          =   330
            Left            =   12255
            TabIndex        =   222
            Top             =   1425
            Visible         =   0   'False
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   582
            Caption         =   "frmCC_Colection_RITCARD.frx":20C5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":2131
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
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
            Format          =   "&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__________________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask AOfficeAdd 
            Height          =   330
            Index           =   2
            Left            =   11430
            TabIndex        =   223
            Top             =   1065
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   582
            Caption         =   "frmCC_Colection_RITCARD.frx":2173
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":21DF
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   1
            AutoConvert     =   1
            BackColor       =   16777215
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   0
            Format          =   "[9999]"
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
            ReadOnly        =   -1
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "[____]"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask AOfficeAdd 
            Height          =   330
            Index           =   3
            Left            =   11430
            TabIndex        =   224
            Top             =   1425
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   582
            Caption         =   "frmCC_Colection_RITCARD.frx":2221
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":228D
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   1
            AutoConvert     =   1
            BackColor       =   16777215
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   0
            Format          =   "[9999]"
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
            ReadOnly        =   -1
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "[____]"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtHomeAdd1 
            Height          =   330
            Left            =   12255
            TabIndex        =   225
            Top             =   375
            Visible         =   0   'False
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   582
            Caption         =   "frmCC_Colection_RITCARD.frx":22CF
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":233B
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
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
            Format          =   "&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__________________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtHomeAdd2 
            Height          =   330
            Left            =   12255
            TabIndex        =   226
            Top             =   720
            Visible         =   0   'False
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   582
            Caption         =   "frmCC_Colection_RITCARD.frx":237D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":23E9
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
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
            Format          =   "&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__________________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask AHomeAdd1 
            Height          =   300
            Index           =   0
            Left            =   11430
            TabIndex        =   227
            Top             =   390
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   529
            Caption         =   "frmCC_Colection_RITCARD.frx":242B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":2497
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   1
            AutoConvert     =   1
            BackColor       =   16777215
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   0
            Format          =   "[9999]"
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
            ReadOnly        =   -1
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "[____]"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask AHomeAdd2 
            Height          =   330
            Index           =   1
            Left            =   11430
            TabIndex        =   228
            Top             =   705
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   582
            Caption         =   "frmCC_Colection_RITCARD.frx":24D9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":2545
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   1
            AutoConvert     =   1
            BackColor       =   16777215
            BorderStyle     =   1
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   0
            Format          =   "[9999]"
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
            ReadOnly        =   -1
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "[____]"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtHomeAdd1A 
            Height          =   330
            Left            =   12255
            TabIndex        =   229
            Top             =   375
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   582
            Caption         =   "frmCC_Colection_RITCARD.frx":2587
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":25F3
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
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
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__________________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtHomeAdd2A 
            Height          =   330
            Left            =   12255
            TabIndex        =   230
            Top             =   720
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   582
            Caption         =   "frmCC_Colection_RITCARD.frx":2635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":26A1
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
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
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__________________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtOfficeAdd1A 
            Height          =   330
            Left            =   12255
            TabIndex        =   231
            Top             =   1065
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   582
            Caption         =   "frmCC_Colection_RITCARD.frx":26E3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":274F
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
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
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__________________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtOfficeAdd2A 
            Height          =   330
            Left            =   12255
            TabIndex        =   232
            Top             =   1425
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   582
            Caption         =   "frmCC_Colection_RITCARD.frx":2791
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":27FD
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
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
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__________________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtMobileAdd1A 
            Height          =   330
            Left            =   11430
            TabIndex        =   233
            Top             =   2490
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   582
            Caption         =   "frmCC_Colection_RITCARD.frx":283F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":28AB
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
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
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__________________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtMobileAdd2A 
            Height          =   330
            Left            =   11430
            TabIndex        =   234
            Top             =   2835
            Width           =   1815
            _Version        =   65536
            _ExtentX        =   3201
            _ExtentY        =   582
            Caption         =   "frmCC_Colection_RITCARD.frx":28ED
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":2959
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
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
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__________________"
            Value           =   ""
         End
         Begin RichTextLib.RichTextBox AddrNow 
            Height          =   1215
            Left            =   15195
            TabIndex        =   235
            Top             =   315
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   2143
            _Version        =   393217
            BackColor       =   16777215
            Enabled         =   -1  'True
            Appearance      =   0
            TextRTF         =   $"frmCC_Colection_RITCARD.frx":299B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSFrame SSFrame2 
            Height          =   3675
            Index           =   0
            Left            =   6690
            TabIndex        =   289
            Top             =   0
            Width           =   45
            _ExtentX        =   79
            _ExtentY        =   6482
            _Version        =   196610
            Font3D          =   1
            Caption         =   "SSFrame2"
            ShadowStyle     =   1
         End
         Begin Threed.SSFrame SSFrame2 
            Height          =   3615
            Index           =   1
            Left            =   10545
            TabIndex        =   290
            Top             =   0
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   6376
            _Version        =   196610
            Font3D          =   1
            Caption         =   "SSFrame2"
            ShadowStyle     =   1
         End
         Begin TDBMask6Ctl.TDBMask TxtExt1 
            Height          =   315
            Left            =   9900
            TabIndex        =   295
            Top             =   930
            Width           =   615
            _Version        =   65536
            _ExtentX        =   1085
            _ExtentY        =   556
            Caption         =   "frmCC_Colection_RITCARD.frx":2A20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":2A8C
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   16777215
            BorderStyle     =   0
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   0
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "____________________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask TxtExt2 
            Height          =   315
            Left            =   9900
            TabIndex        =   296
            Top             =   1275
            Width           =   645
            _Version        =   65536
            _ExtentX        =   1138
            _ExtentY        =   556
            Caption         =   "frmCC_Colection_RITCARD.frx":2ACE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":2B3A
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   16777215
            BorderStyle     =   0
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   0
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "____________________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask TxtExt3 
            Height          =   330
            Left            =   14100
            TabIndex        =   297
            Top             =   1065
            Width           =   675
            _Version        =   65536
            _ExtentX        =   1191
            _ExtentY        =   582
            Caption         =   "frmCC_Colection_RITCARD.frx":2B7C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":2BE8
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   16777215
            BorderStyle     =   0
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   0
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "____________________"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask TxtExt4 
            Height          =   330
            Left            =   14100
            TabIndex        =   298
            Top             =   1410
            Width           =   675
            _Version        =   65536
            _ExtentX        =   1191
            _ExtentY        =   582
            Caption         =   "frmCC_Colection_RITCARD.frx":2C2A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_RITCARD.frx":2C96
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            AllowSpace      =   -1
            AutoConvert     =   -1
            BackColor       =   16777215
            BorderStyle     =   0
            ClipMode        =   0
            CursorPosition  =   -1
            DataProperty    =   0
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "&&&&&&&&&&&&&&&&&&&&"
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
            ReadOnly        =   -1
            ShowContextMenu =   0
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "____________________"
            Value           =   ""
         End
         Begin TDBNumber6Ctl.TDBNumber LblPrompA 
            Height          =   300
            Left            =   5145
            TabIndex        =   310
            Top             =   495
            Visible         =   0   'False
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   529
            Calculator      =   "frmCC_Colection_RITCARD.frx":2CD8
            Caption         =   "frmCC_Colection_RITCARD.frx":2CF8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD.frx":2D64
            Keys            =   "frmCC_Colection_RITCARD.frx":2D82
            Spin            =   "frmCC_Colection_RITCARD.frx":2DCC
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
            ForeColor       =   0
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
         Begin VB.Label lblaoc 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
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
            Height          =   300
            Left            =   8220
            TabIndex        =   309
            Top             =   2700
            Width           =   1500
         End
         Begin VB.Label Label32 
            Caption         =   "Coding :"
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
            Height          =   240
            Left            =   7095
            TabIndex        =   308
            Top             =   2730
            Width           =   975
         End
         Begin VB.Label lblRecsource 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2640
            TabIndex        =   302
            Top             =   3330
            Width           =   180
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Batch"
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
            Height          =   195
            Index           =   80
            Left            =   2040
            TabIndex        =   301
            Top             =   3330
            Width           =   480
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "No CC"
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
            Height          =   195
            Index           =   65
            Left            =   150
            TabIndex        =   300
            Top             =   210
            Width           =   465
         End
         Begin VB.Label lblCustId 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   810
            TabIndex        =   299
            Top             =   240
            Width           =   180
         End
         Begin VB.Label lblregion 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   840
            TabIndex        =   288
            Top             =   2955
            Width           =   2655
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Region"
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
            Height          =   195
            Left            =   150
            TabIndex        =   287
            Top             =   3045
            Width           =   585
         End
         Begin VB.Label Label36 
            Caption         =   "Priority"
            Height          =   195
            Left            =   3195
            TabIndex        =   286
            Top             =   1875
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.Label LblRiskLevel 
            AutoSize        =   -1  'True
            BackColor       =   &H00E8BE91&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3210
            TabIndex        =   285
            Top             =   1815
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label CustId 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Risk Level"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   2550
            TabIndex        =   284
            Top             =   1725
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblPriority 
            AutoSize        =   -1  'True
            BackColor       =   &H00E8BE91&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3300
            TabIndex        =   283
            Top             =   1740
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label LblDOB 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   795
            TabIndex        =   282
            Top             =   1215
            Width           =   1470
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ZipCode"
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
            Height          =   195
            Left            =   75
            TabIndex        =   281
            Top             =   3330
            Width           =   675
         End
         Begin VB.Label lblZIP 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   825
            TabIndex        =   280
            Top             =   3360
            Width           =   120
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   72
            Left            =   75
            TabIndex        =   279
            Top             =   315
            Width           =   60
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "Office Add"
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
            Height          =   555
            Left            =   150
            TabIndex        =   278
            Top             =   2235
            Width           =   525
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
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
            Height          =   195
            Left            =   150
            TabIndex        =   277
            Top             =   1515
            Width           =   690
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   150
            TabIndex        =   276
            Top             =   1230
            Width           =   345
         End
         Begin VB.Label lblID 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   810
            TabIndex        =   275
            Top             =   870
            Width           =   120
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ID No"
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
            Height          =   195
            Left            =   150
            TabIndex        =   274
            Top             =   855
            Width           =   450
         End
         Begin VB.Label lblNama 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   810
            TabIndex        =   273
            Top             =   495
            Width           =   120
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
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
            Height          =   195
            Left            =   150
            TabIndex        =   272
            Top             =   495
            Width           =   480
         End
         Begin VB.Label lblCardNo 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3540
            TabIndex        =   271
            Top             =   3555
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.Label CustId 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "# Card"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   2985
            TabIndex        =   270
            Top             =   3525
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Wo After Pay"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   390
            Index           =   11
            Left            =   4185
            TabIndex        =   269
            Top             =   2535
            Width           =   660
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblwilling 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "-----------------"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5145
            TabIndex        =   268
            Top             =   2985
            Width           =   1515
         End
         Begin VB.Label Label11 
            BackColor       =   &H0000FF00&
            BackStyle       =   0  'Transparent
            Caption         =   "Willingness"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   10
            Left            =   4185
            TabIndex        =   267
            Top             =   3030
            Width           =   900
         End
         Begin VB.Label lblaging 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "                         "
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5145
            TabIndex        =   266
            Top             =   3345
            Width           =   1515
         End
         Begin VB.Label Label11 
            BackColor       =   &H0000FF00&
            BackStyle       =   0  'Transparent
            Caption         =   "Aging"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   9
            Left            =   4200
            TabIndex        =   265
            Top             =   3360
            Width           =   690
         End
         Begin VB.Label Label11 
            BackColor       =   &H0000FF00&
            BackStyle       =   0  'Transparent
            Caption         =   "WO_Date"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   4185
            TabIndex        =   264
            Top             =   1605
            Width           =   840
         End
         Begin VB.Label Label11 
            BackColor       =   &H0000FF00&
            BackStyle       =   0  'Transparent
            Caption         =   "Limit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   4185
            TabIndex        =   263
            Top             =   915
            Width           =   405
         End
         Begin VB.Label Label11 
            BackColor       =   &H0000FF00&
            BackStyle       =   0  'Transparent
            Caption         =   "Balance"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   4185
            TabIndex        =   262
            Top             =   180
            Width           =   675
         End
         Begin VB.Label Label11 
            BackColor       =   &H0000FF00&
            BackStyle       =   0  'Transparent
            Caption         =   "LPA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   4
            Left            =   4185
            TabIndex        =   261
            Top             =   2355
            Width           =   375
         End
         Begin VB.Label Label11 
            BackColor       =   &H0000FF00&
            BackStyle       =   0  'Transparent
            Caption         =   "LPD"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Index           =   2
            Left            =   4185
            TabIndex        =   260
            Top             =   1995
            Width           =   345
         End
         Begin VB.Label Label18 
            BackColor       =   &H0000FF00&
            BackStyle       =   0  'Transparent
            Caption         =   "Open Date"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   4185
            TabIndex        =   259
            Top             =   1230
            Width           =   795
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FF00&
            BackStyle       =   0  'Transparent
            Caption         =   "#Card"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2520
            TabIndex        =   258
            Top             =   165
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Label lblNoCard 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "-------------------"
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   3075
            TabIndex        =   257
            Top             =   210
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label11 
            Caption         =   "Principle Afterpay"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   8
            Left            =   4185
            TabIndex        =   256
            Top             =   480
            Visible         =   0   'False
            Width           =   660
            WordWrap        =   -1  'True
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "HP II"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   6780
            TabIndex        =   255
            Top             =   2070
            Width           =   360
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "HP I"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   6780
            TabIndex        =   254
            Top             =   1665
            Width           =   300
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Rumah II"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   6780
            TabIndex        =   253
            Top             =   615
            Width           =   660
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Rumah I"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   6780
            TabIndex        =   252
            Top             =   210
            Width           =   600
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Kantor I"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   6780
            TabIndex        =   251
            Top             =   885
            Width           =   585
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Kantor II"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   6765
            TabIndex        =   250
            Top             =   1320
            Width           =   645
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            Caption         =   " Address"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   14265
            TabIndex        =   249
            Top             =   2490
            Width           =   855
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            Caption         =   "Telp "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   14640
            TabIndex        =   248
            Top             =   2205
            Width           =   540
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   14610
            TabIndex        =   247
            Top             =   1845
            Width           =   540
         End
         Begin VB.Label label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Emergency Contact"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   46
            Left            =   15195
            TabIndex        =   246
            Top             =   1590
            Width           =   1890
         End
         Begin VB.Label label1 
            BackColor       =   &H00E8BE91&
            BackStyle       =   0  'Transparent
            Caption         =   "Fax I"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   10635
            TabIndex        =   245
            Top             =   1875
            Width           =   765
         End
         Begin VB.Label label1 
            BackColor       =   &H00E8BE91&
            BackStyle       =   0  'Transparent
            Caption         =   "Fax II"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   11
            Left            =   10635
            TabIndex        =   244
            Top             =   2220
            Width           =   765
         End
         Begin VB.Label label1 
            BackColor       =   &H00E8BE91&
            BackStyle       =   0  'Transparent
            Caption         =   "HP II"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   13
            Left            =   10635
            TabIndex        =   243
            Top             =   2910
            Width           =   765
         End
         Begin VB.Label label1 
            BackColor       =   &H00E8BE91&
            BackStyle       =   0  'Transparent
            Caption         =   "HP I"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   14
            Left            =   10635
            TabIndex        =   242
            Top             =   2520
            Width           =   765
         End
         Begin VB.Label label1 
            BackColor       =   &H00E8BE91&
            BackStyle       =   0  'Transparent
            Caption         =   "Kantor II"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   16
            Left            =   10635
            TabIndex        =   241
            Top             =   1485
            Width           =   765
         End
         Begin VB.Label label1 
            BackColor       =   &H00E8BE91&
            BackStyle       =   0  'Transparent
            Caption         =   "Kantor I"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   17
            Left            =   10635
            TabIndex        =   240
            Top             =   1125
            Width           =   765
         End
         Begin VB.Label label1 
            BackColor       =   &H00E8BE91&
            BackStyle       =   0  'Transparent
            Caption         =   "Rumah II"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   19
            Left            =   10635
            TabIndex        =   239
            Top             =   765
            Width           =   765
         End
         Begin VB.Label label1 
            BackColor       =   &H00E8BE91&
            BackStyle       =   0  'Transparent
            Caption         =   "Rumah I"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   20
            Left            =   10635
            TabIndex        =   238
            Top             =   450
            Width           =   765
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Telp Tambahan"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   22
            Left            =   10680
            TabIndex        =   237
            Top             =   135
            Width           =   1500
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "Add  Adress:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   14370
            TabIndex        =   236
            Top             =   345
            Width           =   795
         End
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   3705
      Index           =   3
      Left            =   18135
      TabIndex        =   294
      Top             =   195
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   6535
      _Version        =   196610
      Font3D          =   1
      Caption         =   "SSFrame2"
      ShadowStyle     =   1
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   3705
      Index           =   2
      Left            =   30
      TabIndex        =   293
      Top             =   195
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   6535
      _Version        =   196610
      Font3D          =   1
      Caption         =   "SSFrame2"
      ShadowStyle     =   1
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   75
      Index           =   1
      Left            =   15
      TabIndex        =   292
      Top             =   3825
      Width           =   18180
      _ExtentX        =   32068
      _ExtentY        =   132
      _Version        =   196610
      ShadowStyle     =   1
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   75
      Index           =   0
      Left            =   30
      TabIndex        =   291
      Top             =   300
      Width           =   18180
      _ExtentX        =   32068
      _ExtentY        =   132
      _Version        =   196610
      ShadowStyle     =   1
   End
   Begin VB.Frame Frame7 
      Height          =   1755
      Left            =   13875
      TabIndex        =   170
      Top             =   3915
      Width           =   4290
      Begin VB.ComboBox CmbPhone 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmCC_Colection_RITCARD.frx":2DF4
         Left            =   270
         List            =   "frmCC_Colection_RITCARD.frx":2DF6
         TabIndex        =   175
         Top             =   435
         Width           =   2010
      End
      Begin VB.TextBox txtPhone 
         Height          =   285
         Left            =   765
         Locked          =   -1  'True
         TabIndex        =   172
         Top             =   1860
         Visible         =   0   'False
         Width           =   2640
      End
      Begin VB.TextBox txtPhoneA 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   171
         Top             =   1935
         Width           =   2625
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   375
         Index           =   0
         Left            =   2685
         TabIndex        =   176
         Top             =   315
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         _Version        =   196610
         Font3D          =   4
         MousePointer    =   16
         ForeColor       =   12582912
         PictureMaskColor=   -2147483644
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Call"
         AutoSize        =   2
         ButtonStyle     =   2
         PictureAlignment=   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   375
         Index           =   2
         Left            =   2685
         TabIndex        =   177
         Top             =   1155
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   661
         _Version        =   196610
         Font3D          =   2
         MousePointer    =   16
         ForeColor       =   8388608
         PictureMaskColor=   -2147483644
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Save"
         AutoSize        =   2
         Alignment       =   4
         ButtonStyle     =   2
         PictureAlignment=   1
      End
      Begin Threed.SSCommand SSCommand1 
         Cancel          =   -1  'True
         Height          =   375
         Index           =   3
         Left            =   3450
         TabIndex        =   178
         Top             =   1155
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   661
         _Version        =   196610
         Font3D          =   2
         MousePointer    =   16
         ForeColor       =   12582912
         PictureMaskColor=   -2147483644
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Exit"
         AutoSize        =   2
         Alignment       =   4
         ButtonStyle     =   2
         PictureAlignment=   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   375
         Index           =   1
         Left            =   3345
         TabIndex        =   179
         Top             =   300
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   661
         _Version        =   196610
         Font3D          =   4
         MousePointer    =   16
         ForeColor       =   12582912
         PictureMaskColor=   -2147483644
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&Hangup"
         AutoSize        =   2
         ButtonStyle     =   2
         PictureAlignment=   1
      End
      Begin VB.Label lbltype 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   165
         TabIndex        =   181
         Top             =   1155
         Width           =   2250
      End
      Begin VB.Label label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Pilih No Telp :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   255
         TabIndex        =   180
         Top             =   225
         Width           =   1185
      End
      Begin VB.Label LBLEXP 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   2205
         TabIndex        =   174
         Top             =   795
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label lblstatus 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3360
         TabIndex        =   173
         Top             =   210
         Width           =   60
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "Result Telp."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4890
      Left            =   7755
      TabIndex        =   101
      Top             =   5700
      Width           =   10440
      Begin VB.CheckBox C_SKIP 
         Caption         =   "Skip"
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
         Left            =   5715
         TabIndex        =   126
         Top             =   1260
         Width           =   705
      End
      Begin VB.CheckBox C_Contacted 
         Caption         =   "Contacted"
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
         Left            =   5760
         TabIndex        =   136
         Top             =   330
         Width           =   1215
      End
      Begin VB.Frame Frame5 
         Caption         =   "Reserved PTP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1890
         Left            =   150
         TabIndex        =   150
         Top             =   2895
         Width           =   4965
         Begin MSComctlLib.ListView LstPayment 
            Height          =   1575
            Left            =   150
            TabIndex        =   151
            Top             =   225
            Width           =   3660
            _ExtentX        =   6456
            _ExtentY        =   2778
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16761024
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin Threed.SSCommand SSCommand2 
            Height          =   375
            Index           =   0
            Left            =   3930
            TabIndex        =   152
            Top             =   315
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   661
            _Version        =   196610
            Caption         =   "Tambah"
         End
         Begin Threed.SSCommand SSCommand2 
            Height          =   375
            Index           =   2
            Left            =   3930
            TabIndex        =   153
            Top             =   675
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   661
            _Version        =   196610
            Caption         =   "Hapus"
         End
         Begin Threed.SSCommand SSCommand2 
            Height          =   375
            Index           =   1
            Left            =   3915
            TabIndex        =   154
            Top             =   315
            Visible         =   0   'False
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   661
            _Version        =   196610
            Caption         =   "&Ubah"
         End
      End
      Begin VB.ComboBox Cmbwith 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         ItemData        =   "frmCC_Colection_RITCARD.frx":2DF8
         Left            =   6450
         List            =   "frmCC_Colection_RITCARD.frx":2E05
         TabIndex        =   149
         Top             =   2910
         Width           =   2715
      End
      Begin VB.Frame FrmContacted 
         Height          =   900
         Left            =   5700
         TabIndex        =   137
         Top             =   315
         Width           =   4590
         Begin VB.ComboBox cmbDescCon 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   660
            TabIndex        =   139
            Top             =   540
            Width           =   3825
         End
         Begin VB.ComboBox cmbContacted 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            ItemData        =   "frmCC_Colection_RITCARD.frx":2E23
            Left            =   660
            List            =   "frmCC_Colection_RITCARD.frx":2E25
            TabIndex        =   138
            Top             =   240
            Width           =   3195
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Cont."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   40
            Left            =   165
            TabIndex        =   141
            Top             =   240
            Width           =   405
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Desc."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   38
            Left            =   150
            TabIndex        =   140
            Top             =   495
            Width           =   405
         End
      End
      Begin VB.ComboBox cbolastcall 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6450
         TabIndex        =   135
         Top             =   3210
         Width           =   2715
      End
      Begin VB.Frame FrmSKIP 
         Height          =   885
         Left            =   5700
         TabIndex        =   130
         Top             =   1230
         Width           =   4575
         Begin VB.ComboBox cboskip 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            ItemData        =   "frmCC_Colection_RITCARD.frx":2E27
            Left            =   570
            List            =   "frmCC_Colection_RITCARD.frx":2E29
            TabIndex        =   132
            Top             =   255
            Width           =   3285
         End
         Begin VB.ComboBox cbodescskip 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   570
            TabIndex        =   131
            Top             =   540
            Width           =   3885
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Desc."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   33
            Left            =   165
            TabIndex        =   134
            Top             =   510
            Width           =   405
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Skip."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   32
            Left            =   225
            TabIndex        =   133
            Top             =   300
            Width           =   345
         End
      End
      Begin VB.Frame frmpopsp 
         Height          =   495
         Left            =   5700
         TabIndex        =   127
         Top             =   2160
         Width           =   4560
         Begin VB.ComboBox cboPOPSP 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   570
            TabIndex        =   128
            Top             =   150
            Width           =   3915
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   39
            Left            =   255
            TabIndex        =   129
            Top             =   150
            Width           =   225
         End
      End
      Begin VB.CheckBox C_PTP 
         Caption         =   "PTP"
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
         Left            =   255
         TabIndex        =   103
         Top             =   1425
         Width           =   750
      End
      Begin VB.CheckBox C_VALID 
         Caption         =   "UnContacted"
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
         Left            =   255
         TabIndex        =   102
         Top             =   330
         Width           =   1560
      End
      Begin VB.Frame FrMValid 
         Height          =   945
         Left            =   165
         TabIndex        =   107
         Top             =   300
         Width           =   4980
         Begin VB.ComboBox cbovalid 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            ItemData        =   "frmCC_Colection_RITCARD.frx":2E2B
            Left            =   1005
            List            =   "frmCC_Colection_RITCARD.frx":2E2D
            TabIndex        =   109
            Top             =   285
            Width           =   2955
         End
         Begin VB.ComboBox cbodescvalid 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1005
            TabIndex        =   108
            Top             =   585
            Width           =   3465
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Description:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   25
            Left            =   90
            TabIndex        =   111
            Top             =   615
            Width           =   855
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Valid :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   28
            Left            =   525
            TabIndex        =   110
            Top             =   330
            Width           =   435
         End
      End
      Begin VB.Frame frmPTP 
         Height          =   1440
         Left            =   150
         TabIndex        =   104
         Top             =   1455
         Width           =   4995
         Begin VB.CheckBox C_Payment 
            Enabled         =   0   'False
            Height          =   255
            Left            =   3690
            TabIndex        =   116
            Top             =   150
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.ComboBox cmbDiscount 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            ItemData        =   "frmCC_Colection_RITCARD.frx":2E2F
            Left            =   3420
            List            =   "frmCC_Colection_RITCARD.frx":2E31
            TabIndex        =   115
            Text            =   "0"
            Top             =   450
            Width           =   975
         End
         Begin VB.ComboBox CmbBaseOn 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            ItemData        =   "frmCC_Colection_RITCARD.frx":2E33
            Left            =   1095
            List            =   "frmCC_Colection_RITCARD.frx":2E35
            TabIndex        =   114
            Top             =   435
            Width           =   1425
         End
         Begin VB.ComboBox cboPTP 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1095
            TabIndex        =   105
            Top             =   165
            Width           =   2415
         End
         Begin TDBNumber6Ctl.TDBNumber txttenor 
            Height          =   300
            Left            =   3405
            TabIndex        =   112
            Top             =   1095
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   529
            Calculator      =   "frmCC_Colection_RITCARD.frx":2E37
            Caption         =   "frmCC_Colection_RITCARD.frx":2E57
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD.frx":2EC3
            Keys            =   "frmCC_Colection_RITCARD.frx":2EE1
            Spin            =   "frmCC_Colection_RITCARD.frx":2F2B
            AlignHorizontal =   2
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "####0;;Null"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "####0"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   99999
            MinValue        =   -99999
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
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin TDBDate6Ctl.TDBDate TDBDate3 
            Height          =   345
            Left            =   3420
            TabIndex        =   113
            Top             =   750
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   609
            Calendar        =   "frmCC_Colection_RITCARD.frx":2F53
            Caption         =   "frmCC_Colection_RITCARD.frx":306B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD.frx":30D7
            Keys            =   "frmCC_Colection_RITCARD.frx":30F5
            Spin            =   "frmCC_Colection_RITCARD.frx":3153
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
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
            ShowContextMenu =   1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   6815745
            Value           =   39876
            CenturyMode     =   0
         End
         Begin TDBNumber6Ctl.TDBNumber txtPayment 
            Height          =   345
            Left            =   1095
            TabIndex        =   117
            Top             =   720
            Width           =   1410
            _Version        =   65536
            _ExtentX        =   2487
            _ExtentY        =   609
            Calculator      =   "frmCC_Colection_RITCARD.frx":317B
            Caption         =   "frmCC_Colection_RITCARD.frx":319B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD.frx":3207
            Keys            =   "frmCC_Colection_RITCARD.frx":3225
            Spin            =   "frmCC_Colection_RITCARD.frx":326F
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   16777215
            BorderStyle     =   1
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
         Begin TDBNumber6Ctl.TDBNumber Tdabamoint 
            Height          =   345
            Left            =   1095
            TabIndex        =   118
            Top             =   1065
            Width           =   1410
            _Version        =   65536
            _ExtentX        =   2487
            _ExtentY        =   609
            Calculator      =   "frmCC_Colection_RITCARD.frx":3297
            Caption         =   "frmCC_Colection_RITCARD.frx":32B7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD.frx":3323
            Keys            =   "frmCC_Colection_RITCARD.frx":3341
            Spin            =   "frmCC_Colection_RITCARD.frx":338B
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   16777215
            BorderStyle     =   1
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
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Payment"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   79
            Left            =   3945
            TabIndex        =   125
            Top             =   150
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Date PTP:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   2535
            TabIndex        =   124
            Top             =   795
            Width           =   855
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Disc:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   75
            Left            =   2955
            TabIndex        =   123
            Top             =   480
            Width           =   435
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "AmountPTP:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   77
            Left            =   30
            TabIndex        =   122
            Top             =   765
            Width           =   1050
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Base On :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   30
            TabIndex        =   121
            Top             =   465
            Width           =   1050
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Installment:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   42
            Left            =   30
            TabIndex        =   120
            Top             =   1110
            Width           =   1050
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tenor:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   44
            Left            =   2805
            TabIndex        =   119
            Top             =   1095
            Width           =   570
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            Caption         =   "PTP:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   37
            Left            =   30
            TabIndex        =   106
            Top             =   210
            Width           =   1050
         End
      End
      Begin TDBDate6Ctl.TDBDate cmbDateSch 
         Height          =   315
         Left            =   6420
         TabIndex        =   157
         Top             =   3585
         Width           =   1740
         _Version        =   65536
         _ExtentX        =   3069
         _ExtentY        =   556
         Calendar        =   "frmCC_Colection_RITCARD.frx":33B3
         Caption         =   "frmCC_Colection_RITCARD.frx":34CB
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_RITCARD.frx":3537
         Keys            =   "frmCC_Colection_RITCARD.frx":3555
         Spin            =   "frmCC_Colection_RITCARD.frx":35B3
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
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
      Begin TDBTime6Ctl.TDBTime cmbTimeSch 
         Height          =   315
         Left            =   8190
         TabIndex        =   158
         Top             =   3570
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   556
         Caption         =   "frmCC_Colection_RITCARD.frx":35DB
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection_RITCARD.frx":3647
         Spin            =   "frmCC_Colection_RITCARD.frx":3697
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
         Text            =   "__:__"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   1.02960316199441E-317
      End
      Begin RichTextLib.RichTextBox txtRemarks 
         Height          =   840
         Left            =   6390
         TabIndex        =   159
         Top             =   3945
         Width           =   4020
         _ExtentX        =   7091
         _ExtentY        =   1482
         _Version        =   393217
         BackColor       =   16777215
         Enabled         =   -1  'True
         TextRTF         =   $"frmCC_Colection_RITCARD.frx":36BF
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox cmbNextAct 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8565
         TabIndex        =   155
         Top             =   4260
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox cmbPrior 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         ItemData        =   "frmCC_Colection_RITCARD.frx":373A
         Left            =   8520
         List            =   "frmCC_Colection_RITCARD.frx":3747
         Style           =   2  'Dropdown List
         TabIndex        =   156
         Top             =   4230
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label39 
         Caption         =   "Tgl FollowUp."
         Height          =   255
         Left            =   5355
         TabIndex        =   169
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label31 
         Caption         =   "Remarks:"
         Height          =   135
         Index           =   1
         Left            =   5385
         TabIndex        =   161
         Top             =   4035
         Width           =   780
      End
      Begin VB.Label Label34 
         Caption         =   "Berbicara Dgn:"
         Height          =   255
         Left            =   5355
         TabIndex        =   160
         Top             =   2955
         Width           =   1065
      End
      Begin VB.Label Label31 
         Caption         =   "Status Telp:"
         Height          =   255
         Index           =   0
         Left            =   5355
         TabIndex        =   142
         Top             =   3225
         Width           =   885
      End
   End
   Begin VB.Frame Frame9 
      Height          =   4005
      Left            =   75
      TabIndex        =   78
      Top             =   5640
      Width           =   7635
      Begin VB.OptionButton Option8 
         Caption         =   "Tambah"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   345
         TabIndex        =   81
         Top             =   2070
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Batal"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   1395
         TabIndex        =   80
         Top             =   2085
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Frame FrmPayment 
         Caption         =   "Detail Payment"
         Height          =   1860
         Left            =   60
         TabIndex        =   79
         Top             =   120
         Width           =   7605
         Begin VB.CommandButton CmdDeletePelunasan 
            Caption         =   "Hapus"
            Height          =   420
            Left            =   4935
            TabIndex        =   168
            Top             =   1245
            Visible         =   0   'False
            Width           =   795
         End
         Begin TDBNumber6Ctl.TDBNumber txtSisaHutang 
            Height          =   345
            Left            =   5715
            TabIndex        =   162
            Top             =   840
            Width           =   1830
            _Version        =   65536
            _ExtentX        =   3228
            _ExtentY        =   609
            Calculator      =   "frmCC_Colection_RITCARD.frx":375F
            Caption         =   "frmCC_Colection_RITCARD.frx":377F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD.frx":37EB
            Keys            =   "frmCC_Colection_RITCARD.frx":3809
            Spin            =   "frmCC_Colection_RITCARD.frx":3853
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483624
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,###,##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   0
            Format          =   "###,###,###,##0"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999999999999
            MinValue        =   -999999999999
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
            MaxValueVT      =   6750213
            MinValueVT      =   3538949
         End
         Begin TDBNumber6Ctl.TDBNumber TxtAfterPay 
            Height          =   330
            Left            =   5715
            TabIndex        =   163
            Top             =   495
            Width           =   1830
            _Version        =   65536
            _ExtentX        =   3228
            _ExtentY        =   582
            Calculator      =   "frmCC_Colection_RITCARD.frx":387B
            Caption         =   "frmCC_Colection_RITCARD.frx":389B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD.frx":3907
            Keys            =   "frmCC_Colection_RITCARD.frx":3925
            Spin            =   "frmCC_Colection_RITCARD.frx":396F
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483624
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,###,##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   0
            Format          =   "###,###,###,##0"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999999999999
            MinValue        =   -999999999999
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
            MaxValueVT      =   6750213
            MinValueVT      =   3538949
         End
         Begin TDBNumber6Ctl.TDBNumber TxtPayment2 
            Height          =   315
            Left            =   5715
            TabIndex        =   164
            Top             =   165
            Width           =   1830
            _Version        =   65536
            _ExtentX        =   3228
            _ExtentY        =   556
            Calculator      =   "frmCC_Colection_RITCARD.frx":3997
            Caption         =   "frmCC_Colection_RITCARD.frx":39B7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD.frx":3A23
            Keys            =   "frmCC_Colection_RITCARD.frx":3A41
            Spin            =   "frmCC_Colection_RITCARD.frx":3A8B
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483624
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,###,##0"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   0
            Format          =   "###,###,###,##0"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999999999
            MinValue        =   -999999999
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
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin MSComctlLib.ListView listview1 
            Height          =   1620
            Index           =   0
            Left            =   45
            TabIndex        =   305
            Top             =   180
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   2858
            View            =   3
            LabelEdit       =   1
            SortOrder       =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Jml PTP:"
            Height          =   240
            Index           =   0
            Left            =   4830
            TabIndex        =   167
            Top             =   210
            Width           =   885
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Jml Dibayar:"
            Height          =   255
            Left            =   4800
            TabIndex        =   166
            Top             =   525
            Width           =   885
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Sisa:"
            Height          =   225
            Left            =   4785
            TabIndex        =   165
            Top             =   855
            Width           =   885
         End
      End
      Begin VB.Frame Frame8 
         ForeColor       =   &H000000FF&
         Height          =   1725
         Left            =   60
         TabIndex        =   82
         Top             =   2145
         Visible         =   0   'False
         Width           =   7560
         Begin VB.TextBox TxtName 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Left            =   915
            Locked          =   -1  'True
            TabIndex        =   88
            Top             =   540
            Width           =   3135
         End
         Begin VB.TextBox TxtCustid 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Left            =   1740
            Locked          =   -1  'True
            TabIndex        =   87
            Top             =   3375
            Width           =   1935
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
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
            Left            =   915
            TabIndex        =   86
            Top             =   225
            Width           =   1815
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Alamat Billing"
            Height          =   195
            Index           =   0
            Left            =   4125
            TabIndex        =   85
            Top             =   855
            Width           =   1440
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Rumah"
            Height          =   195
            Index           =   1
            Left            =   5565
            TabIndex        =   84
            Top             =   855
            Width           =   840
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Kantor"
            Height          =   195
            Index           =   2
            Left            =   6525
            TabIndex        =   83
            Top             =   840
            Width           =   840
         End
         Begin TDBNumber6Ctl.TDBNumber TDBNumber1 
            Height          =   315
            Left            =   915
            TabIndex        =   89
            Top             =   870
            Width           =   615
            _Version        =   65536
            _ExtentX        =   1085
            _ExtentY        =   556
            Calculator      =   "frmCC_Colection_RITCARD.frx":3AB3
            Caption         =   "frmCC_Colection_RITCARD.frx":3AD3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD.frx":3B3F
            Keys            =   "frmCC_Colection_RITCARD.frx":3B5D
            Spin            =   "frmCC_Colection_RITCARD.frx":3BA7
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   16777215
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "####0;;Null"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "####0"
            HighlightText   =   0
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   99999
            MinValue        =   -99999
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
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin RichTextLib.RichTextBox TXtDetails 
            Height          =   570
            Left            =   4080
            TabIndex        =   90
            Top             =   225
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   1005
            _Version        =   393217
            BackColor       =   16777215
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmCC_Colection_RITCARD.frx":3BCF
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin TDBDate6Ctl.TDBDate TDBDate2 
            Height          =   315
            Left            =   915
            TabIndex        =   91
            Top             =   1200
            Visible         =   0   'False
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   556
            Calendar        =   "frmCC_Colection_RITCARD.frx":3C54
            Caption         =   "frmCC_Colection_RITCARD.frx":3D6C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD.frx":3DD8
            Keys            =   "frmCC_Colection_RITCARD.frx":3DF6
            Spin            =   "frmCC_Colection_RITCARD.frx":3E54
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   16777215
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            CursorPosition  =   0
            DataProperty    =   0
            DisplayFormat   =   "mm/dd/yyyy"
            EditMode        =   0
            Enabled         =   -1
            ErrorBeep       =   0
            FirstMonth      =   4
            ForeColor       =   -2147483640
            Format          =   "mm/dd/yyyy"
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
            ValueVT         =   2010382337
            Value           =   2.12482692446619E-314
            CenturyMode     =   0
         End
         Begin TDBDate6Ctl.TDBDate TDBDate1 
            Height          =   315
            Left            =   1590
            TabIndex        =   92
            Top             =   870
            Visible         =   0   'False
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   556
            Calendar        =   "frmCC_Colection_RITCARD.frx":3E7C
            Caption         =   "frmCC_Colection_RITCARD.frx":3F94
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD.frx":4000
            Keys            =   "frmCC_Colection_RITCARD.frx":401E
            Spin            =   "frmCC_Colection_RITCARD.frx":407C
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   16777215
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
            ValueVT         =   2010382337
            Value           =   2.12482692446619E-314
            CenturyMode     =   0
         End
         Begin RichTextLib.RichTextBox TxtAddress 
            Height          =   540
            Left            =   4065
            TabIndex        =   93
            Top             =   1065
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   953
            _Version        =   393217
            BackColor       =   16777215
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmCC_Colection_RITCARD.frx":40A4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Nomor"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   30
            TabIndex        =   100
            Top             =   240
            Width           =   810
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            Caption         =   "Note:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2925
            TabIndex        =   99
            Top             =   195
            Width           =   1095
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "Visit Date"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   30
            TabIndex        =   98
            Top             =   1245
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            Caption         =   "Visit Ke"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   30
            TabIndex        =   97
            Top             =   930
            Width           =   810
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   30
            TabIndex        =   96
            Top             =   540
            Width           =   810
         End
         Begin VB.Label label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Custid"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   2
            Left            =   420
            TabIndex        =   95
            Top             =   3375
            Width           =   1095
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "Visit Ke:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   3390
            TabIndex        =   94
            Top             =   915
            Width           =   615
         End
      End
   End
   Begin VB.Frame FrmPayment1 
      Height          =   1365
      Left            =   5520
      TabIndex        =   143
      Top             =   7935
      Width           =   2085
      Begin VB.CheckBox Check3 
         Caption         =   "Regular to paid Off"
         Height          =   195
         Left            =   75
         TabIndex        =   146
         Top             =   285
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Iregular to Paid Off"
         Height          =   195
         Left            =   60
         TabIndex        =   145
         Top             =   360
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Regular Payment"
         Height          =   195
         Left            =   75
         TabIndex        =   144
         Top             =   870
         Visible         =   0   'False
         Width           =   435
      End
      Begin TDBDate6Ctl.TDBDate TdbPTP 
         Height          =   255
         Left            =   60
         TabIndex        =   147
         Top             =   585
         Visible         =   0   'False
         Width           =   1440
         _Version        =   65536
         _ExtentX        =   2540
         _ExtentY        =   450
         Calendar        =   "frmCC_Colection_RITCARD.frx":4129
         Caption         =   "frmCC_Colection_RITCARD.frx":4241
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_RITCARD.frx":42AD
         Keys            =   "frmCC_Colection_RITCARD.frx":42CB
         Spin            =   "frmCC_Colection_RITCARD.frx":4329
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
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
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   3.54027066542603E-316
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate TdbDatePTP 
         Height          =   225
         Left            =   60
         TabIndex        =   148
         Top             =   1065
         Visible         =   0   'False
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   397
         Calendar        =   "frmCC_Colection_RITCARD.frx":4351
         Caption         =   "frmCC_Colection_RITCARD.frx":4469
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_RITCARD.frx":44D5
         Keys            =   "frmCC_Colection_RITCARD.frx":44F3
         Spin            =   "frmCC_Colection_RITCARD.frx":4551
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
         Value           =   3.54027066542603E-316
         CenturyMode     =   0
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1695
      Left            =   3870
      TabIndex        =   0
      Top             =   6105
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   2990
      _Version        =   393216
      Tabs            =   6
      Tab             =   4
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Personal Data"
      TabPicture(0)   =   "frmCC_Colection_RITCARD.frx":4579
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Option3"
      Tab(0).Control(1)=   "Option4"
      Tab(0).Control(2)=   "Option1"
      Tab(0).Control(3)=   "Option2"
      Tab(0).Control(4)=   "Option6"
      Tab(0).Control(5)=   "Option5"
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(7)=   "Frame3"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Additional Fields"
      TabPicture(1)   =   "frmCC_Colection_RITCARD.frx":4595
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "History"
      TabPicture(2)   =   "frmCC_Colection_RITCARD.frx":45B1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Results"
      TabPicture(3)   =   "frmCC_Colection_RITCARD.frx":45CD
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label33"
      Tab(3).Control(1)=   "FrmUnContacted"
      Tab(3).Control(2)=   "txtResult"
      Tab(3).Control(3)=   "txtResultDesc"
      Tab(3).Control(4)=   "txtDiscount"
      Tab(3).Control(5)=   "FrmLunas"
      Tab(3).Control(6)=   "C_NotContacted"
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "Detail Payment"
      TabPicture(4)   =   "frmCC_Colection_RITCARD.frx":45E9
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Request Visit"
      TabPicture(5)   =   "frmCC_Colection_RITCARD.frx":4605
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "LstVisit"
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame4 
         Caption         =   "Emergency Contact"
         Height          =   2475
         Left            =   -72105
         TabIndex        =   70
         Top             =   825
         Width           =   4575
      End
      Begin VB.CheckBox C_NotContacted 
         BackColor       =   &H00C5974B&
         Height          =   270
         Left            =   -74430
         TabIndex        =   68
         Top             =   7950
         Width           =   375
      End
      Begin VB.Frame FrmLunas 
         Height          =   1215
         Left            =   -74640
         TabIndex        =   57
         Top             =   8520
         Visible         =   0   'False
         Width           =   4335
         Begin RichTextLib.RichTextBox TxtFieldName 
            Height          =   375
            Left            =   1560
            TabIndex        =   64
            Top             =   1200
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393217
            Enabled         =   -1  'True
            TextRTF         =   $"frmCC_Colection_RITCARD.frx":4621
         End
         Begin TDBNumber6Ctl.TDBNumber TDBTot_payment 
            Height          =   375
            Left            =   1560
            TabIndex        =   63
            Top             =   720
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   661
            Calculator      =   "frmCC_Colection_RITCARD.frx":46A3
            Caption         =   "frmCC_Colection_RITCARD.frx":46C3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD.frx":472F
            Keys            =   "frmCC_Colection_RITCARD.frx":474D
            Spin            =   "frmCC_Colection_RITCARD.frx":4797
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   -2147483643
            BorderStyle     =   1
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
            MaxValue        =   99999999999
            MinValue        =   -99999999999
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
            MaxValueVT      =   6750213
            MinValueVT      =   3538949
         End
         Begin VB.CheckBox C_lunas 
            BackColor       =   &H00C5974B&
            Caption         =   "Lunas"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   390
            TabIndex        =   58
            Top             =   900
            Width           =   1455
         End
         Begin TDBDate6Ctl.TDBDate TdbLunas 
            Height          =   285
            Left            =   1560
            TabIndex        =   59
            Top             =   360
            Width           =   1350
            _Version        =   65536
            _ExtentX        =   2381
            _ExtentY        =   503
            Calendar        =   "frmCC_Colection_RITCARD.frx":47BF
            Caption         =   "frmCC_Colection_RITCARD.frx":48D7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD.frx":4943
            Keys            =   "frmCC_Colection_RITCARD.frx":4961
            Spin            =   "frmCC_Colection_RITCARD.frx":49BF
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
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
            Value           =   3.54027066542603E-316
            CenturyMode     =   0
         End
         Begin VB.Label LblLunas 
            Caption         =   "Label19"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1620
            TabIndex        =   66
            Top             =   660
            Width           =   4215
         End
         Begin VB.Label Label17 
            Caption         =   "Label17"
            Height          =   375
            Left            =   1320
            TabIndex        =   65
            Top             =   0
            Width           =   135
         End
         Begin VB.Label Label9 
            Caption         =   "Field Name"
            Height          =   255
            Left            =   240
            TabIndex        =   62
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Total Payment"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   61
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Date of Payment"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   60
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Height          =   555
         Left            =   -66585
         TabIndex        =   45
         Top             =   1095
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000013&
         Height          =   3255
         Left            =   -71385
         TabIndex        =   35
         Top             =   330
         Width           =   5970
         Begin VB.Frame Frame6 
            Height          =   615
            Left            =   1275
            TabIndex        =   75
            Top             =   1455
            Visible         =   0   'False
            Width           =   3045
            Begin TDBNumber6Ctl.TDBNumber txtAmountwo_A 
               Height          =   315
               Left            =   1200
               TabIndex        =   76
               Top             =   720
               Width           =   1485
               _Version        =   65536
               _ExtentX        =   2619
               _ExtentY        =   564
               Calculator      =   "frmCC_Colection_RITCARD.frx":49E7
               Caption         =   "frmCC_Colection_RITCARD.frx":4A07
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_Colection_RITCARD.frx":4A73
               Keys            =   "frmCC_Colection_RITCARD.frx":4A91
               Spin            =   "frmCC_Colection_RITCARD.frx":4ADB
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   65280
               BorderStyle     =   0
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,###,##0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   16711680
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
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               Caption         =   "AmountWo Afterpay"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   7
               Left            =   120
               TabIndex        =   77
               Top             =   600
               Width           =   930
               WordWrap        =   -1  'True
            End
         End
         Begin TDBDate6Ctl.TDBDate lblLastBill 
            Height          =   300
            Left            =   3150
            TabIndex        =   36
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   529
            Calendar        =   "frmCC_Colection_RITCARD.frx":4B03
            Caption         =   "frmCC_Colection_RITCARD.frx":4C1B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD.frx":4C87
            Keys            =   "frmCC_Colection_RITCARD.frx":4CA5
            Spin            =   "frmCC_Colection_RITCARD.frx":4D03
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   15253137
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   1
            Value           =   3.54028845178928E-316
            CenturyMode     =   0
         End
         Begin TDBDate6Ctl.TDBDate lblLcAtm 
            Height          =   285
            Left            =   1785
            TabIndex        =   37
            Top             =   360
            Visible         =   0   'False
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            Calendar        =   "frmCC_Colection_RITCARD.frx":4D2B
            Caption         =   "frmCC_Colection_RITCARD.frx":4E43
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD.frx":4EAF
            Keys            =   "frmCC_Colection_RITCARD.frx":4ECD
            Spin            =   "frmCC_Colection_RITCARD.frx":4F2B
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   15253137
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
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "__/__/____"
            ValidateMode    =   0
            ValueVT         =   1
            Value           =   3.54025880785053E-316
            CenturyMode     =   0
         End
         Begin TDBNumber6Ctl.TDBNumber lblPromPA1 
            Height          =   300
            Left            =   4290
            TabIndex        =   51
            Top             =   210
            Width           =   1125
            _Version        =   65536
            _ExtentX        =   1984
            _ExtentY        =   529
            Calculator      =   "frmCC_Colection_RITCARD.frx":4F53
            Caption         =   "frmCC_Colection_RITCARD.frx":4F73
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD.frx":4FDF
            Keys            =   "frmCC_Colection_RITCARD.frx":4FFD
            Spin            =   "frmCC_Colection_RITCARD.frx":5047
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   15253137
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
         Begin TDBNumber6Ctl.TDBNumber lblTtlPay 
            Height          =   315
            Left            =   4020
            TabIndex        =   73
            Top             =   2190
            Visible         =   0   'False
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   556
            Calculator      =   "frmCC_Colection_RITCARD.frx":506F
            Caption         =   "frmCC_Colection_RITCARD.frx":508F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_RITCARD.frx":50FB
            Keys            =   "frmCC_Colection_RITCARD.frx":5119
            Spin            =   "frmCC_Colection_RITCARD.frx":5163
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   15253137
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
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Ttl Pay"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   5
            Left            =   5280
            TabIndex        =   74
            Top             =   2550
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.Label LblFees 
            AutoSize        =   -1  'True
            BackColor       =   &H00E8BE91&
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2730
            TabIndex        =   55
            Top             =   2730
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Label LblInterest 
            AutoSize        =   -1  'True
            BackColor       =   &H00E8BE91&
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4200
            TabIndex        =   54
            Top             =   2250
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FF00&
            BackStyle       =   0  'Transparent
            Caption         =   "Fees"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   2160
            TabIndex        =   53
            Top             =   2700
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FF00&
            BackStyle       =   0  'Transparent
            Caption         =   "Interest"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   5970
            TabIndex        =   52
            Top             =   2460
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.Label lblBrokenPromised 
            AutoSize        =   -1  'True
            BackColor       =   &H00E8BE91&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4170
            TabIndex        =   44
            Top             =   2610
            Visible         =   0   'False
            Width           =   105
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FF00&
            BackStyle       =   0  'Transparent
            Caption         =   "Broken Promise"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   390
            Left            =   1830
            TabIndex        =   43
            Top             =   2760
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FF00&
            BackStyle       =   0  'Transparent
            Caption         =   "Lc atmp"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   435
            Index           =   0
            Left            =   2430
            TabIndex        =   42
            Top             =   2760
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FF00&
            BackStyle       =   0  'Transparent
            Caption         =   "Last Bill"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   360
            Left            =   4620
            TabIndex        =   41
            Top             =   2520
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FF00&
            BackStyle       =   0  'Transparent
            Caption         =   "Principle"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   4320
            TabIndex        =   40
            Top             =   2790
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label lblNoPay 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4680
            TabIndex        =   39
            Top             =   2820
            Visible         =   0   'False
            Width           =   75
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FF00&
            BackStyle       =   0  'Transparent
            Caption         =   "No Pay"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   2880
            TabIndex        =   38
            Top             =   2640
            Visible         =   0   'False
            Width           =   600
         End
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00E8BE91&
         Height          =   195
         Left            =   -64260
         TabIndex        =   33
         Top             =   4440
         Width           =   225
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00E8BE91&
         Height          =   195
         Left            =   -64290
         TabIndex        =   31
         Top             =   4065
         Width           =   210
      End
      Begin VB.TextBox txtDiscount 
         Height          =   285
         Left            =   -70380
         TabIndex        =   7
         Top             =   7770
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtResultDesc 
         Height          =   285
         Left            =   -69540
         TabIndex        =   6
         Top             =   7830
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtResult 
         Height          =   285
         Left            =   -67560
         TabIndex        =   5
         Top             =   7620
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E8BE91&
         Height          =   195
         Left            =   -71100
         TabIndex        =   4
         Top             =   4380
         Width           =   225
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E8BE91&
         Height          =   195
         Left            =   -71130
         TabIndex        =   3
         Top             =   4035
         Width           =   240
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E8BE91&
         Height          =   195
         Left            =   -67500
         TabIndex        =   2
         Top             =   4065
         Width           =   210
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E8BE91&
         Height          =   195
         Left            =   -67515
         TabIndex        =   1
         Top             =   4425
         Width           =   255
      End
      Begin MSComctlLib.ListView listview1 
         Height          =   5400
         Index           =   3
         Left            =   -74850
         TabIndex        =   8
         Top             =   375
         Width           =   10275
         _ExtentX        =   18124
         _ExtentY        =   9525
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDropMode     =   1
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16436909
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
         NumItems        =   0
      End
      Begin VB.Frame FrmUnContacted 
         Height          =   1095
         Left            =   -74430
         TabIndex        =   46
         Top             =   8640
         Width           =   4620
         Begin VB.CheckBox chkAppv 
            BackColor       =   &H00C5974B&
            Caption         =   "NO"
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   1
            Left            =   3600
            TabIndex        =   69
            Top             =   360
            Width           =   975
         End
         Begin VB.CheckBox chkAppv 
            BackColor       =   &H00C5974B&
            Caption         =   "YES"
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   0
            Left            =   3600
            TabIndex        =   67
            Top             =   120
            Width           =   975
         End
         Begin VB.ComboBox cmbUncontacted 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmCC_Colection_RITCARD.frx":518B
            Left            =   1250
            List            =   "frmCC_Colection_RITCARD.frx":518D
            TabIndex        =   48
            Top             =   320
            Width           =   2340
         End
         Begin VB.ComboBox cmbDescUn 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmCC_Colection_RITCARD.frx":518F
            Left            =   1245
            List            =   "frmCC_Colection_RITCARD.frx":5191
            TabIndex        =   47
            Top             =   630
            Width           =   3285
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C5974B&
            Caption         =   "Uncontacted"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   66
            Left            =   480
            TabIndex        =   56
            Top             =   0
            Width           =   1170
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Uncontacted"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   34
            Left            =   150
            TabIndex        =   50
            Top             =   360
            Width           =   1050
         End
         Begin VB.Label label1 
            AutoSize        =   -1  'True
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   35
            Left            =   150
            TabIndex        =   49
            Top             =   720
            Width           =   960
         End
      End
      Begin MSComctlLib.ListView LstVisit 
         Height          =   1215
         Left            =   -74760
         TabIndex        =   71
         Top             =   2880
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   2143
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   8454016
         BorderStyle     =   1
         Appearance      =   0
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
      Begin VB.Label Label33 
         Caption         =   "PTP Warna merah sudah ada payment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   -74790
         TabIndex        =   72
         Top             =   7710
         Width           =   4695
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Mobile Phone II"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   30
         Left            =   -67650
         TabIndex        =   34
         Top             =   4395
         Width           =   1335
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Mobile Phone I"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   31
         Left            =   -67650
         TabIndex        =   32
         Top             =   4035
         Width           =   1260
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFC0C0&
         X1              =   0
         X2              =   9000
         Y1              =   -3960
         Y2              =   -3960
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Fax I"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   63
         Left            =   -74850
         TabIndex        =   30
         Top             =   2790
         Width           =   435
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Fax II"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   61
         Left            =   -74850
         TabIndex        =   29
         Top             =   3150
         Width           =   510
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C5974B&
         Caption         =   "Fax Phone Additional"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   71
         Left            =   -74895
         TabIndex        =   28
         Top             =   2535
         Width           =   1785
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Mobile Phone II"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   49
         Left            =   -74910
         TabIndex        =   27
         Top             =   4110
         Width           =   1335
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Mobile Phone I"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   50
         Left            =   -74910
         TabIndex        =   26
         Top             =   3750
         Width           =   1260
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C5974B&
         Caption         =   "Mobile Phone Additional"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   70
         Left            =   -74910
         TabIndex        =   25
         Top             =   3510
         Width           =   2025
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Office Phone II"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   54
         Left            =   -74835
         TabIndex        =   24
         Top             =   2190
         Width           =   1290
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Office Phone I"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   53
         Left            =   -74835
         TabIndex        =   23
         Top             =   1830
         Width           =   1215
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C5974B&
         Caption         =   "Office Phone Additional"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   69
         Left            =   -74895
         TabIndex        =   22
         Top             =   1560
         Width           =   1980
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Home Phone II"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   57
         Left            =   -74820
         TabIndex        =   21
         Top             =   1185
         Width           =   1290
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E8BE91&
         Caption         =   "Home Phone I"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   58
         Left            =   -74820
         TabIndex        =   20
         Top             =   825
         Width           =   1215
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C5974B&
         Caption         =   "Home Phone Additional"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   68
         Left            =   -74850
         TabIndex        =   19
         Top             =   540
         Width           =   1980
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C5974B&
         Caption         =   "Next Action "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   74
         Left            =   -74805
         TabIndex        =   18
         Top             =   4395
         Width           =   1035
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Priority"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   45
         Left            =   -74235
         TabIndex        =   17
         Top             =   5355
         Width           =   615
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Schedule"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   43
         Left            =   -74385
         TabIndex        =   16
         Top             =   4995
         Width           =   780
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Next Action"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   41
         Left            =   -74580
         TabIndex        =   15
         Top             =   4635
         Width           =   975
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   48
         Left            =   -70200
         TabIndex        =   14
         Top             =   4365
         Width           =   765
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Home Phone II"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   23
         Left            =   -74730
         TabIndex        =   13
         Top             =   4320
         Width           =   1290
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Home Phone I"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   21
         Left            =   -74730
         TabIndex        =   12
         Top             =   4005
         Width           =   1215
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Off Phone I"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   27
         Left            =   -70830
         TabIndex        =   11
         Top             =   4065
         Width           =   975
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Off Phone II"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   26
         Left            =   -70830
         TabIndex        =   10
         Top             =   4365
         Width           =   1050
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C5974B&
         Caption         =   "Data Phone Customer"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   73
         Left            =   -74730
         TabIndex        =   9
         Top             =   3735
         Width           =   1890
      End
   End
End
Attribute VB_Name = "FrmCC_ColectionCARD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_cust As ADODB.Recordset
Dim M_update As ADODB.Recordset
Dim M_OBJRS As ADODB.Recordset
Dim stscall As Boolean
Dim TYPETELP As String
Dim kontak As Boolean
Dim spend As Boolean
Dim adaSCH As Boolean
Dim adaREG As Boolean
Dim adaPO As Boolean
Dim vrcek As String
Dim vrdateptp As String
Dim vramount As String
Dim vrtdbdateptp As String
Dim vrbaseon As String
Dim vrdiskon As String
Dim vrtenor As String
Dim vrttlptp As String

Private Sub C_Contacted_Click()
If C_Contacted.Value Then
        C_VALID.Value = False
        C_SKIP.Value = False
        C_Payment.Value = False
        C_PTP.Value = False
      '  C_POPSP.Value = False
        FrmContacted.Enabled = True
        cboPOPSP.Text = ""
   Else
        cmbContacted.Text = ""
        cmbDescCon.Text = ""
        FrmContacted.Enabled = False
        If cboPOPSP.Text = "" Then
            C_Payment.Value = False
        End If
        CmbBaseOn.Text = ""
        cmbDiscount.Text = 0
        TdbPTP.Value = ""
        txtPayment.Value = 0
End If
End Sub

Private Sub C_NotContacted_Click()
   If C_NotContacted.Value Then
      FrmUnContacted.Enabled = True
      C_Contacted.Value = False
      C_Payment.Value = False
   Else
      FrmUnContacted.Enabled = False
      cmbDescUn.Text = ""
      cmbUncontacted = ""
   End If
End Sub

Private Sub C_Payment_Click()
   If C_Payment.Value Then
     ' Frame54.Enabled = True
   Else
     ' Frame54.Enabled = False
     If cboPOPSP.Text <> "" Then
     Exit Sub
     End If
     
      cmbDiscount.Text = ""
   End If
End Sub
Private Sub C_PTP_Click()
If C_PTP.Value Then
        C_VALID.Value = False
        C_SKIP.Value = False
        C_Contacted.Value = False
        frmPTP.Enabled = True
        FrmPayment.Enabled = True
        cboPOPSP.Tag = 0
        cboPOPSP.Text = ""
        C_Payment.Value = 1
   Else
   
        'C_Payment.Value = 0
       ' CmbBaseOn.Text = ""
       ' cmbDiscount.Text = 0
        'txtPayment.Value = 0
'        TxtPtpAddr.Text = ""
 '       TxtPhonePTP.Text = ""
        FrmPayment.Enabled = False
        cboPTP.Text = ""
        frmPTP.Enabled = False
        TdbPTP.Value = ""
        CmbBaseOn.Text = ""
        cmbDiscount.Text = 0
        TdbPTP.Value = ""
        txtPayment.Value = 0
        'C_Payment = False
End If

End Sub

Private Sub C_SKIP_Click()
If C_SKIP.Value Then
        C_VALID.Value = False
        C_Contacted.Value = False
        C_Payment.Value = False
        C_PTP.Value = False
     
        FrmSKIP.Enabled = True
   Else
        cboskip.Text = ""
        cbodescskip.Text = ""
        FrmSKIP.Enabled = False
End If

End Sub

Private Sub C_VALID_Click()
If C_VALID.Value Then
        C_Contacted.Value = False
        C_SKIP.Value = False
        C_Payment.Value = False
        C_PTP.Value = False
        
        FrMValid.Enabled = True
   Else
        cbovalid.Text = ""
        cbodescvalid.Text = ""
        FrMValid.Enabled = False
End If

End Sub
Private Sub cbolastcall_GotFocus()
cbolastcall.Clear
Dim M_OBJRS As ADODB.Recordset
Set M_OBJRS = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient
If Left(cmbContacted.Text, 2) = "OP" Then
CMDSQL = " Select * from ContactedDesc where kdnoprodPresented <> 'PTP-PROMISE TO PAY' "
ElseIf Left(cboPTP.Text, 3) = "PTP" Then
CMDSQL = " Select * from ContactedDesc where kdnoprodPresented <> 'OP-ON PROGRESS' "
Else
CMDSQL = " Select * from ContactedDesc"
End If
M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
While Not M_OBJRS.EOF
    cbolastcall.AddItem M_OBJRS("KdNoProdPresented")
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing

Set M_OBJRS = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient
M_OBJRS.Open "Select * from UnContactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
While Not M_OBJRS.EOF
    cbolastcall.AddItem M_OBJRS("KdNoProdPresented")
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
End Sub

Private Sub cbolastcall_KeyDown(KeyCode As Integer, Shift As Integer)

cbolastcall.Text = ""
Exit Sub
End Sub

Private Sub cboPOPSP_Click()
Dim M_COL1 As New ADODB.Recordset
If Left(cboPOPSP.Text, 2) = "SP" Then
    C_Contacted.Value = 0
    C_SKIP.Value = 0
    C_PTP.Value = 0
    C_VALID.Value = 0
    CmbBaseOn.Text = ""
    cmbDiscount.Text = ""
    txtPayment.Value = 0
    Tdabamoint.Value = 0
    TDBDate3.Value = ""
    txttenor.Value = 0
    cmbDescCon.Enabled = False
    C_Payment.Value = 1
    FrmPayment.Enabled = True
            Set M_COL1 = New ADODB.Recordset
            CMDSQL = "SELECT tglincoming,tdbdatePTP,ttlPTP,discpersen,cmbbaseon,dateptp,tenor,amountptp from mgm where custid ='" & VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(1) & "'"
            M_COL1.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
'            CmbBaseOn.Text = "PRINCIPLE"
            txtPayment.Value = CStr(IIf(IsNull(M_COL1!ttlptp), "", M_COL1!ttlptp))
            CmbBaseOn.Text = CStr(IIf(IsNull(M_COL1!CmbBaseOn), "", M_COL1!CmbBaseOn))
            TdbPTP.Value = CStr(IIf(IsNull(M_COL1!TdbDatePTP), "", M_COL1!TdbDatePTP))
            cmbDiscount.Text = CStr(IIf(IsNull(M_COL1!discpersen), "", M_COL1!discpersen))
            TDBDate3.Value = CStr(IIf(IsNull(M_COL1!dateptp), "", M_COL1!dateptp))
            txttenor.Value = CStr(IIf(IsNull(M_COL1!Tenor), 0, M_COL1!Tenor))
            Tdabamoint.Value = CStr(IIf(IsNull(M_COL1!amountptp), 0, M_COL1!amountptp))
End If

'C_Payment.Value = 0



'txtPayment.Value = 0

End Sub

Private Sub cboPOPSP_KeyDown(KeyCode As Integer, Shift As Integer)

cboPOPSP.Text = ""
End Sub


Private Sub cboskip_Click()
cbodescskip.Clear
If Left(cboskip.Text, 2) <> "MV" Then
   Set M_OBJRS = New ADODB.Recordset
   M_OBJRS.CursorLocation = adUseClient
   M_OBJRS.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
         For i = 0 To 3
           cbodescskip.AddItem M_OBJRS("Description")
           M_OBJRS.MoveNext
         Next i
   Set M_OBJRS = Nothing
   C_Payment.Value = 0
Else
   Set M_OBJRS = New ADODB.Recordset
   M_OBJRS.CursorLocation = adUseClient
      M_OBJRS.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
       While Not M_OBJRS.EOF
           cbodescskip.AddItem M_OBJRS("Description")
           M_OBJRS.MoveNext
       Wend
   Set M_OBJRS = Nothing
   C_Payment.Value = 0
End If

End Sub

Private Sub cbovalid_Click()
Dim i As Integer
cbodescvalid.Clear
If Left(cbovalid.Text, 2) = "NA" Then
        cbodescvalid.Enabled = True
'        CmbBaseOn.Text = ""
'        txtPayment.Text = 0
'        cmbDiscount.Text = ""
'        TdbPTP.Text = ""
'        TdbDatePTP.Text = ""
        Set M_OBJRS = New ADODB.Recordset
        M_OBJRS.CursorLocation = adUseClient
          M_OBJRS.Open "Select * from DescContacted WHERE id >= 13 ", M_OBJCONN, adOpenDynamic, adLockOptimistic
        While Not M_OBJRS.EOF
            cbodescvalid.AddItem M_OBJRS("Description")
            M_OBJRS.MoveNext
        Wend
        C_Payment.Value = 0
        Set M_OBJRS = Nothing
'        FrmPayment.Enabled = False
Else
        Set M_OBJRS = New ADODB.Recordset
        M_OBJRS.CursorLocation = adUseClient
          M_OBJRS.Open "Select * from DescunContacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
        While Not M_OBJRS.EOF
            cbodescvalid.AddItem M_OBJRS("Description")
            M_OBJRS.MoveNext
        Wend
        C_Payment.Value = 0
        Set M_OBJRS = Nothing
End If

End Sub

Private Sub cbovalid_KeyDown(KeyCode As Integer, Shift As Integer)

cbovalid.Text = ""
Exit Sub
End Sub

Private Sub Check1_Click()
regnego = False
Check2.Value = 0
Check3.Value = 0
If CmbBaseOn.Text = "PRINCIPLE" Then
    MsgBox "Regular payment only TOTAL AMOUNT"
    CmbBaseOn.SetFocus
    Exit Sub
Else
'Call CEKPTP
'If adaSCH Then
'    MsgBox "Hapus dulu PTP yang ada atau selesaikan paymennya!"
'    Exit Sub
'Else
    Call ISIJMLPAYMENT
    If Check1.Value = 1 Then
        frmregpayment.Show
    End If
End If
End Sub

Sub CEKPTP()
Dim RS As New ADODB.Recordset
RS.CursorLocation = adUseClient
RS.Open "select TYPE from TBLNEGOPTP where custid='" & lblCustId.Caption & "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
If RS.BOF And RS.EOF Then
Else
    While Not RS.EOF
        If RS!Type = "SCH" Then
            adaSCH = True
        ElseIf RS!Type = "REG" Then
            adaREG = True
        ElseIf RS!Type = "PO" Then
            adaPO = True
        End If
        RS.MoveNext
    Wend
End If
Set RS = Nothing
End Sub


Private Sub Check2_Click()
Check1.Value = 0
Check3.Value = 0
If Check2.Value = 1 Then
'    If CmbBaseOn.Text = "PRINCIPLE" Then
'        MsgBox "Regular payment only TOTAL AMOUNT"
'        CmbBaseOn.SetFocus
'        Exit Sub
'    Else
'        Call CEKPTP
'        If adaREG Then
'            MsgBox "Hapus dulu PTP yang ada atau selesaikan paymennya!"
'            Exit Sub
'        Else
            'Call ISIJMLPAYMENT
            regnego = True
            FrmNegoPTP.Show
'        End If
End If
'End If
End Sub

Private Sub Check3_Click()
regnego = False
Check1.Value = 0
Check2.Value = 0

'Call CEKPTP
'If adaPO Then
'    MsgBox "Hapus dulu PTP yang ada atau selesaikan paymennya!"
'    Exit Sub
'Else
    Call ISIJMLPAYMENT
    If Check3.Value = 1 Then
        Frmpaidoff.Show
    End If
'End If
End Sub

Private Sub chkAppv_Click(Index As Integer)
Select Case Index
Case 0:
    chkAppv(1).Value = 0
Case 1:
    chkAppv(0).Value = 0
End Select
End Sub

Private Sub CmbBaseOn_Click()
If CmbBaseOn.Text = "PRINCIPLE" Then
CmbBaseOn.Text = ""
End If
    Call cmbDiscount_Click
End Sub

Private Sub CmbBaseOn_LostFocus()
    'Call cmbDiscount_Click
End Sub

Private Sub cmbContacted_Click()
'DESCRIPTION CONTACTED
Dim i As Integer
cmbDescCon.Clear

'If Left(vrcek, 2) = "BP" And Left(cmbContacted.Text, 3) = "POP" Then
'    cmbContacted.Text = ""
'End If

If Left(cmbContacted.Text, 2) = "RP" Then
    cmbDescCon.Enabled = True
    CmbBaseOn.Text = ""
    txtPayment.Text = 0
    cmbDiscount.Text = ""
    TdbPTP.Text = ""
    TdbDatePTP.Text = ""
   Set M_OBJRS = New ADODB.Recordset
   M_OBJRS.CursorLocation = adUseClient
     M_OBJRS.Open "Select * from DescContacted where id <= 12", M_OBJCONN, adOpenDynamic, adLockOptimistic
    While Not M_OBJRS.EOF
        cmbDescCon.AddItem M_OBJRS("Description")
        M_OBJRS.MoveNext
    Wend
    Set M_OBJRS = Nothing
    C_Payment.Value = 0
    FrmPayment.Enabled = False
    Else
'    If Left(cmbContacted.Text, 2) = "NA" Then
'        cmbDescCon.Enabled = True
'        CmbBaseOn.Text = ""
'        txtPayment.Text = 0
'        cmbDiscount.Text = ""
'        TdbPTP.Text = ""
'        TdbDatePTP.Text = ""
'        Set M_OBJRS = New ADODB.Recordset
'          M_OBJRS.Open "Select * from DescContacted WHERE id >= 13 ", M_OBJCONN, adOpenDynamic, adLockOptimistic
'        While Not M_OBJRS.EOF
'            cmbDescCon.AddItem M_OBJRS("Description")
'            M_OBJRS.MoveNext
'        Wend
'        C_Payment.Value = 0
'        FrmPayment.Enabled = False
        
'    Else
         If Left(cmbContacted.Text, 2) = "PT" Then
            cmbDescCon.Enabled = False
            C_Payment.Value = 1
            FrmPayment.Enabled = True
            CmbBaseOn.Text = "PRINCIPLE"
    Else
        If Left(cmbContacted.Text, 2) = "BP" Then
            cmbDescCon.Enabled = False
            CmbBaseOn.Text = ""
            txtPayment.Text = 0
            cmbDiscount.Text = ""
            TdbPTP.Text = ""
            TdbDatePTP.Text = ""
            C_Payment.Value = 0
            FrmPayment.Enabled = False
    Else
    If Left(cmbContacted.Text, 2) = "OP" Then
            cmbDescCon.Enabled = False
'            CmbBaseOn.Text = ""
'            txtPayment.Text = 0
'            cmbDiscount.Text = ""
'            TdbPTP.Text = ""
'            TdbDatePTP.Text = ""
          '  C_Payment.Value = 1
             'C_Payment.Value = False
            FrmPayment.Enabled = True
      Else
      
    If Left(cmbContacted.Text, 2) = "PO" Or Left(cmbContacted.Text, 2) = "SP" Then
            cmbDescCon.Enabled = False
            C_Payment.Value = 1
            FrmPayment.Enabled = True
        Set m_cust = New ADODB.Recordset
        m_cust.CursorLocation = adUseClient
        CMDSQL = "SELECT tglincoming,tdbdatePTP,ttlPTP,discpersen,cmbbaseon,dateptp,tenor, amountptp from mgm where custid ='" & VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(1) & "'"
        m_cust.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
'            CmbBaseOn.Text = "PRINCIPLE"
            txtPayment.Value = CStr(IIf(IsNull(m_cust!ttlptp), "", m_cust!ttlptp))
           CmbBaseOn.Text = CStr(IIf(IsNull(m_cust!CmbBaseOn), "", m_cust!CmbBaseOn))
            TdbPTP.Value = CStr(IIf(IsNull(m_cust!TdbDatePTP), "", m_cust!TdbDatePTP))
            cmbDiscount.Text = CStr(IIf(IsNull(m_cust!discpersen), "", m_cust!discpersen))
            TDBDate3.Value = CStr(IIf(IsNull(m_cust!dateptp), "", m_cust!dateptp))
            txttenor.Value = CStr(IIf(IsNull(m_cust!Tenor), "0", m_cust!Tenor))
            Tdabamoint.Value = CStr(IIf(IsNull(m_cust!amountptp), 0, m_cust!amountptp))
            
      Set m_cust = Nothing
    End If
End If
End If
End If
End If
'End If

Set M_OBJRS = Nothing
End Sub

Private Sub cmbContacted_KeyDown(KeyCode As Integer, Shift As Integer)

cmbContacted.Text = ""
Exit Sub
End Sub

Private Sub cmbDescCon_GotFocus()
'DESCRIPTION CONTACTED
Dim i As Integer
cmbDescCon.Clear
If Left(cmbContacted.Text, 2) = "RP" Then
    cmbDescCon.Enabled = True
    CmbBaseOn.Text = ""
    txtPayment.Text = 0
    cmbDiscount.Text = ""
    TdbPTP.Text = ""
    TdbDatePTP.Text = ""
   Set M_OBJRS = New ADODB.Recordset
   M_OBJRS.CursorLocation = adUseClient
     M_OBJRS.Open "Select * from DescContacted where id <= 12", M_OBJCONN, adOpenDynamic, adLockOptimistic
    While Not M_OBJRS.EOF
        cmbDescCon.AddItem M_OBJRS("Description")
        M_OBJRS.MoveNext
    Wend
    C_Payment.Value = 0
    FrmPayment.Enabled = False
    Set M_OBJRS = Nothing
    Else
'    If Left(cmbContacted.Text, 2) = "NA" Then
'        cmbDescCon.Enabled = True
'        CmbBaseOn.Text = ""
'        txtPayment.Text = 0
'        cmbDiscount.Text = ""
'        TdbPTP.Text = ""
'        TdbDatePTP.Text = ""
'        Set M_OBJRS = New ADODB.Recordset
'          M_OBJRS.Open "Select * from DescContacted WHERE id >= 13 ", M_OBJCONN, adOpenDynamic, adLockOptimistic
'        While Not M_OBJRS.EOF
'            cmbDescCon.AddItem M_OBJRS("Description")
'            M_OBJRS.MoveNext
'        Wend
'        C_Payment.Value = 0
'        FrmPayment.Enabled = False
        
'    Else
         If Left(cmbContacted.Text, 2) = "PT" Then
            cmbDescCon.Enabled = False
            C_Payment.Value = 1
            FrmPayment.Enabled = True
            CmbBaseOn.Text = "PRINCIPLE"
    Else
        If Left(cmbContacted.Text, 2) = "BP" Then
            cmbDescCon.Enabled = False
            CmbBaseOn.Text = ""
            txtPayment.Text = 0
            cmbDiscount.Text = ""
            TdbPTP.Text = ""
            TdbDatePTP.Text = ""
            C_Payment.Value = 0
            FrmPayment.Enabled = False
    Else
    If Left(cmbContacted.Text, 2) = "OP" Then
            cmbDescCon.Enabled = False
            CmbBaseOn.Text = ""
            txtPayment.Text = 0
            cmbDiscount.Text = ""
            TdbPTP.Text = ""
            TdbDatePTP.Text = ""
            C_Payment.Value = 0
            FrmPayment.Enabled = False
      Else
      
    If Left(cmbContacted.Text, 2) = "PO" Or Left(cmbContacted.Text, 2) = "SP" Then
            cmbDescCon.Enabled = False
            C_Payment.Value = 1
            FrmPayment.Enabled = True
Set m_cust = New ADODB.Recordset
m_cust.CursorLocation = adUseClient
CMDSQL = "SELECT tglincoming,tdbdatePTP,ttlPTP,discpersen,cmbbaseon from mgm where custid ='" & VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(1) & "'"
    m_cust.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
'            CmbBaseOn.Text = "PRINCIPLE"
            txtPayment.Value = CStr(IIf(IsNull(m_cust!ttlptp), "", m_cust!ttlptp))
            CmbBaseOn.Text = CStr(IIf(IsNull(m_cust!CmbBaseOn), "", m_cust!CmbBaseOn))
            TdbPTP.Value = CStr(IIf(IsNull(m_cust!TdbDatePTP), "", m_cust!TdbDatePTP))
            cmbDiscount.Text = CStr(IIf(IsNull(m_cust!discpersen), "", m_cust!discpersen))
            
      Set m_cust = Nothing
    End If
End If
End If
End If
End If
'End If

Set M_OBJRS = Nothing
End Sub

Private Sub cmbDescCon_KeyDown(KeyCode As Integer, Shift As Integer)

cmbDescCon.Text = ""
Exit Sub
End Sub

Private Sub cmbDescUn_GotFocus()
Dim i As Integer
cmbDescUn.Clear
If Left(cmbUncontacted.Text, 2) = "NA" Then
        cmbDescUn.Enabled = True
'        CmbBaseOn.Text = ""
'        txtPayment.Text = 0
'        cmbDiscount.Text = ""
'        TdbPTP.Text = ""
'        TdbDatePTP.Text = ""
        Set M_OBJRS = New ADODB.Recordset
        M_OBJRS.CursorLocation = adUseClient
          M_OBJRS.Open "Select * from DescContacted WHERE id >= 13 ", M_OBJCONN, adOpenDynamic, adLockOptimistic
        While Not M_OBJRS.EOF
            cmbDescUn.AddItem M_OBJRS("Description")
            M_OBJRS.MoveNext
        Wend
        C_Payment.Value = 0
        Set M_OBJRS = Nothing
'        FrmPayment.Enabled = False
Else
If Left(cmbUncontacted.Text, 2) <> "MV" Then
   Set M_OBJRS = New ADODB.Recordset
   M_OBJRS.CursorLocation = adUseClient
   M_OBJRS.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
         For i = 0 To 3
           cmbDescUn.AddItem M_OBJRS("Description")
           M_OBJRS.MoveNext
         Next i
   Set M_OBJRS = Nothing
   C_Payment.Value = 0
Else
   Set M_OBJRS = New ADODB.Recordset
   M_OBJRS.CursorLocation = adUseClient
'   If kontak = True Then
'        m_objrs.Open "Select * from DescUncontacted where ", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    Else
      M_OBJRS.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    End If
       While Not M_OBJRS.EOF
           cmbDescUn.AddItem M_OBJRS("Description")
           M_OBJRS.MoveNext
       Wend
   Set M_OBJRS = Nothing
   C_Payment.Value = 0
End If
End If
End Sub

Private Sub cmbDescUn_KeyDown(KeyCode As Integer, Shift As Integer)

cmbDescUn.Text = ""
Exit Sub
End Sub

Private Sub cmbDiscount_Change()
Call ISIJMLPAYMENT
End Sub

Private Sub cmbDiscount_Click()
Call ISIJMLPAYMENT
'Check1.Enabled = True
'Check2.Enabled = True
'Check3.Enabled = True
'If Left(cmbContacted.Text, 2) = "OP" Then
'    Check1.Enabled = False
'    Check3.Enabled = False
'End If
End Sub

Sub ISIJMLPAYMENT()
Dim M_OBJRS As New ADODB.Recordset
'If cmbDiscount.Text = "" Then
'    Exit Sub
'End If

M_OBJRS.CursorLocation = adUseClient
M_OBJRS.Open "Select * from tbldiscount where Description = '" + cmbDiscount.Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If M_OBJRS.RecordCount <> 0 Then
    TdbDatePTP.Value = MDIForm1.TDBDate1.Value + IIf(IsNull(M_OBJRS!hari), 7, M_OBJRS!hari)
Else
    TdbDatePTP.Value = MDIForm1.TDBDate1.Value + 7
End If
If cmbDiscount.Text = "0" Or cmbDiscount.Text = "" Then
    If CmbBaseOn.Text = "PRINCIPLE" Then
        txtPayment.Value = lblPromPA.Value
    Else
    
         txtPayment.Value = lblAmount.Value
         Exit Sub
         
'         If CmbBaseOn.Text = "TOTAL AMOUNT" Then
'            If lblAmount.Value = 0 Or lblAmount.ValueIsNull Or cmbDiscount = "" Then
'                txtPayment.Value = 0
'            Else
'                txtDiscount.Text = CStr((cmbDiscount.Text) / 100)
'                txtPayment.Value = lblAmount.Value - (CCur(txtDiscount.Text) * lblAmount.Value)
'            End If
'        End If
    End If
End If

        If CmbBaseOn.Text = "TOTAL AMOUNT" Then
            If lblAmount.Value = 0 Or lblAmount.ValueIsNull Or cmbDiscount = "" Then
                txtPayment.Value = 0
            Else
                txtDiscount.Text = CStr((cmbDiscount.Text) / 100)
                txtPayment.Value = lblAmount.Value - (CCur(txtDiscount.Text) * lblAmount.Value)
                End If

                
            End If
       ' End If

'    If CmbBaseOn.Text = "PRINCIPLE" Then
'        If lblPromPA.Value = 0 Or lblPromPA.ValueIsNull Then
'            txtPayment.Value = 0
'        Else
'            txtDiscount.Text = CStr((cmbDiscount.Text) / 100)
'            txtPayment.Value = lblPromPA.Value - (CCur(txtDiscount.Text) * lblPromPA.Value)
'        End If
'    Else
'        If lblAmount.Value = 0 Or lblAmount.ValueIsNull Then
'            txtPayment.Value = 0
'        Else
'            txtDiscount.Text = CStr((cmbDiscount.Text) / 100)
'            txtPayment.Value = lblAmount.Value - (CCur(txtDiscount.Text) * lblAmount.Value)
'        End If
'    End If
'End If
'End If

End Sub

Private Sub cmbDiscount_LostFocus()
'Dim M_OBJRS As New ADODB.Recordset
'If cmbDiscount.Text = "" Then
'    Exit Sub
'End If
'
'M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from tbldiscount where Description = '" + cmbDiscount.Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'If M_OBJRS.RecordCount <> 0 Then
'    TdbDatePTP.Value = MDIForm1.TDBDate1.Value + IIf(IsNull(M_OBJRS!hari), 7, M_OBJRS!hari)
'Else
'    TdbDatePTP.Value = MDIForm1.TDBDate1.Value + 7
'End If
'If cmbDiscount.Text = "0" Then
'Else
'
'    If CmbBaseOn.Text = "PRINCIPLE" Then
'        If lblPromPA.Value = 0 Then
'            txtPayment.Value = 0
'        Else
'            txtDiscount.Text = CStr((cmbDiscount.Text) / 100)
'            txtPayment.Value = lblPromPA.Value - (CCur(txtDiscount.Text) * lblPromPA.Value)
'        End If
'    Else
'        If lblAmount.Value = 0 Then
'            txtPayment.Value = 0
'        Else
'            txtDiscount.Text = CStr((cmbDiscount.Text) / 100)
'            txtPayment.Value = lblAmount.Value - (CCur(txtDiscount.Text) * lblAmount.Value)
'        End If
'    End If
'End If
End Sub

Private Sub cmbNextAct_KeyDown(KeyCode As Integer, Shift As Integer)
cmbNextAct.Text = ""
Exit Sub
End Sub

Private Sub CmbPhone_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmbUncontacted_Click()
'DESCRIPTION UNCONTACTED
Dim i As Integer
cmbDescUn.Clear
If Left(cmbUncontacted.Text, 2) = "NA" Then
        cmbDescUn.Enabled = True
'        CmbBaseOn.Text = ""
'        txtPayment.Text = 0
'        cmbDiscount.Text = ""
'        TdbPTP.Text = ""
'        TdbDatePTP.Text = ""
        Set M_OBJRS = New ADODB.Recordset
        M_OBJRS.CursorLocation = adUseClient
          M_OBJRS.Open "Select * from DescContacted WHERE id >= 13 ", M_OBJCONN, adOpenDynamic, adLockOptimistic
        While Not M_OBJRS.EOF
            cmbDescUn.AddItem M_OBJRS("Description")
            M_OBJRS.MoveNext
        Wend
        C_Payment.Value = 0
        Set M_OBJRS = Nothing
'        FrmPayment.Enabled = False
Else
If Left(cmbUncontacted.Text, 2) <> "MV" Then
   Set M_OBJRS = New ADODB.Recordset
   M_OBJRS.CursorLocation = adUseClient
   M_OBJRS.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
         For i = 0 To 3
           cmbDescUn.AddItem M_OBJRS("Description")
           M_OBJRS.MoveNext
         Next i
   Set M_OBJRS = Nothing
   C_Payment.Value = 0
Else
   Set M_OBJRS = New ADODB.Recordset
   M_OBJRS.CursorLocation = adUseClient
'   If kontak = True Then
'        m_objrs.Open "Select * from DescUncontacted where ", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    Else
      M_OBJRS.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    End If
       While Not M_OBJRS.EOF
           cmbDescUn.AddItem M_OBJRS("Description")
           M_OBJRS.MoveNext
       Wend
   Set M_OBJRS = Nothing
   C_Payment.Value = 0
End If
End If
' Set M_OBJRS = New ADODB.Recordset
'   If kontak = False Then
'          M_OBJRS.Open "Select * from UncontactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
'       While Not M_OBJRS.EOF
'           cmbDescUn.AddItem M_OBJRS("NMnoProdpresented")
'           M_OBJRS.MoveNext
'       Wend
'        Set M_OBJRS = Nothing
'   End If
'   C_Payment.Value = 0
'End If

End Sub

Private Sub headerDatePayment()
LstPayment.ColumnHeaders.ADD 1, , "", 0 * TXT
LstPayment.ColumnHeaders.ADD 2, , "ID", 2 * TXT
LstPayment.ColumnHeaders.ADD 3, , "DATE PROMISE", 15 * TXT
LstPayment.ColumnHeaders.ADD 4, , "PAYMENT", 30 * TXT
LstPayment.ColumnHeaders.ADD 5, , "TYPE", 30 * TXT
LstPayment.ColumnHeaders.ADD 6, , "INPUT DATE", 15 * TXT
End Sub
Private Sub headerCustid_Double()
    LstDoubleId.ColumnHeaders.ADD 1, , "Id", 10 * TXT
    LstDoubleId.ColumnHeaders.ADD 2, , "Nama", 15 * TXT
    LstDoubleId.ColumnHeaders.ADD 3, , "DescColl", 10 * TXT
    LstDoubleId.ColumnHeaders.ADD 4, , "AmountWo", 10 * TXT
    LstDoubleId.ColumnHeaders.ADD 5, , "Principle", 20 * TXT
End Sub
Private Sub cmbUncontacted_KeyDown(KeyCode As Integer, Shift As Integer)
cmbUncontacted.Text = ""
Exit Sub
End Sub
Private Sub Cmbwith_KeyDown(KeyCode As Integer, Shift As Integer)
Cmbwith.Text = ""
Exit Sub
End Sub
Private Sub CmdDeletePelunasan_Click()
Dim m_msgbox As Variant
If listview1(0).ListItems.Count = 0 Then
    Exit Sub
End If
m_msgbox = MsgBox("Yakin Akan Dihapus...!!! ", vbCritical + vbOKCancel, "Peringatan")
If m_msgbox = vbOK Then
    M_OBJCONN.Execute "Delete from tbllunas where id = " + listview1(0).SelectedItem.SubItems(4) + ""
    listview1(0).ListItems.Remove listview1(0).SelectedItem.Index
    MsgBox "Done"
    Call isi_datapayment
End If
End Sub

Private Sub Form_Load()

frmCC_Colection.Left = 10
frmCC_Colection.Top = 20

'cek list pelunasan
Dim i, iIndex As Integer
Dim sKata, cCombo As String


'------->>>  setting No Visit  <<<---------------

Text1.Text = Format(Now, "yymmddhhmmss")
TDBDate1.Value = Now
'If UCase(Left(MDIForm1.Text2.Text, 5)) = "ADMIN" Or UCase(Left(MDIForm1.Text2.Text, 5)) = "SUPER" Then
If UCase(Left(MDIForm1.Text2.Text, 5)) = "ADMIN" Then
    txtHomeNo1.Visible = True
    txtHomeNo1A.Visible = False
    txtHomeNo2.Visible = True
    txtHomeNo2A.Visible = False
    txtOfficeNo1.Visible = True
    txtOfficeNo1A.Visible = False
    txtOfficeNo2.Visible = True
    txtOfficeNo2A.Visible = False
    txtMobileNo1.Visible = True
    txtMobileNo1A.Visible = False
    txtMobileNo2.Visible = True
    txtMobileNo2A.Visible = False
    txtPhone.Visible = True
    txtPhoneA.Visible = False
    txtHomeAdd1.Visible = True
    txtHomeAdd1A.Visible = False
    txtHomeAdd2.Visible = True
    txtHomeAdd2A.Visible = False
    txtOfficeAdd1.Visible = True
    txtOfficeAdd1A.Visible = False
    txtOfficeAdd2.Visible = True
    txtOfficeAdd2A.Visible = False
    txtMobileAdd1.Visible = True
    txtMobileAdd1A.Visible = False
    txtMobileAdd2.Visible = True
    txtMobileAdd2A.Visible = False
    txtECno.Visible = True
    txtECnoA.Visible = False
End If

If UCase(MDIForm1.Text2.Text) = "AGENT" Then
        C_lunas.Enabled = False
        TdbLunas.Enabled = False
        chkAppv(0).Enabled = False
        chkAppv(1).Enabled = False
        TDBTot_payment.Enabled = False
        TxtFieldName.Enabled = False
        CmdDeletePelunasan.Enabled = False
         ' Tampilkan PRincipal
        lblPromPA.Visible = True
        Label11(8).Visible = True
       
ElseIf UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Then
        txtHomeAdd1.ReadOnly = False
        txtHomeAdd2.ReadOnly = False
        txtOfficeAdd1.ReadOnly = False
        txtOfficeAdd2.ReadOnly = False
        txtMobileAdd1.ReadOnly = False
        txtMobileAdd2.ReadOnly = False
         ' Tampilkan PRincipal
        lblPromPA.Visible = True
        Label11(8).Visible = True
       
Else ' utk SPV tampilkan no telp
        txtHomeAdd1.ReadOnly = False
        txtHomeAdd2.ReadOnly = False
        txtOfficeAdd1.ReadOnly = False
        txtOfficeAdd2.ReadOnly = False
        txtMobileAdd1.ReadOnly = False
        txtMobileAdd2.ReadOnly = False
        
        
        txtHomeNo1.Visible = True
        txtHomeNo1A.Visible = False
        txtHomeNo2.Visible = True
        txtHomeNo2A.Visible = False
        
        txtOfficeNo1.Visible = True
        txtOfficeNo1A.Visible = False
        
        txtOfficeNo2.Visible = True
        txtOfficeNo2A.Visible = False
        
        txtMobileNo1.Visible = True
        txtMobileNo1A.Visible = False
        txtMobileNo2.Visible = True
        txtMobileNo2A.Visible = False
        
        txtECno.Visible = True
        txtECnoA.Visible = False
        
        txtHomeAdd1.Visible = True
        txtHomeAdd1A.Visible = False
        txtHomeAdd2.Visible = True
        txtHomeAdd2A.Visible = False
        
        txtOfficeAdd1.Visible = True
        txtOfficeAdd1A.Visible = False
        txtOfficeAdd2.Visible = True
        txtOfficeAdd2A.Visible = False
        
        txtMobileAdd1.Visible = True
        txtMobileAdd1A.Visible = False
        txtMobileAdd2.Visible = True
        txtMobileAdd2A.Visible = False
        ' Tampilkan PRincipal
        lblPromPA.Visible = True
        Label11(8).Visible = True
        
End If
 
   FrmContacted.Enabled = False
   FrmUnContacted.Enabled = False
   FrmPayment.Enabled = False
   
    Call headerDatePayment
    Call headerCustid_Double
    Call HEADER_HISTORY
    Call HEADER_HISTORY_PAID
    Call HEADER_RequestVisit
    Call show_cust
    Call VisitNo
    'Call isi_lastcall
    
    If UCase(MDIForm1.Text2.Text) = "TEAMLEADER" Or UCase(MDIForm1.Text2.Text) = "SUPERVISOR" Or UCase(MDIForm1.Text2.Text) = "ADMINISTRATOR" Then
        Call aktifphone
    End If
    
    If UCase(MDIForm1.Text2.Text) = "AGENT" Then
        Call aktifphoneAGENT
    End If
        
  '  SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
SSTab1.Tab = 0
cmbDateSch.Value = Now
cmbDateSch.Value = ""
'CONTACTED
CmbBaseOn.AddItem "PRINCIPLE"
CmbBaseOn.AddItem "TOTAL AMOUNT"


Set M_OBJRS = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient
M_OBJRS.Open "Select * from tblvalid", M_OBJCONN, adOpenDynamic, adLockOptimistic
    While Not M_OBJRS.EOF
        cbovalid.AddItem M_OBJRS!KdNoProdPresented
        M_OBJRS.MoveNext
    Wend
    Set M_OBJRS = Nothing
    
    Set M_OBJRS = New ADODB.Recordset
    M_OBJRS.CursorLocation = adUseClient
M_OBJRS.Open "Select * from tblptp", M_OBJCONN, adOpenDynamic, adLockOptimistic
    While Not M_OBJRS.EOF
        cboPTP.AddItem M_OBJRS!KdNoProdPresented
        M_OBJRS.MoveNext
    Wend
    Set M_OBJRS = Nothing
    
    Set M_OBJRS = New ADODB.Recordset
    M_OBJRS.CursorLocation = adUseClient
M_OBJRS.Open "Select * from tblskip", M_OBJCONN, adOpenDynamic, adLockOptimistic
    While Not M_OBJRS.EOF
        cboskip.AddItem M_OBJRS!KdNoProdPresented
        M_OBJRS.MoveNext
    Wend
    Set M_OBJRS = Nothing
    
    
    Set M_OBJRS = New ADODB.Recordset
    M_OBJRS.CursorLocation = adUseClient
    M_OBJRS.Open "Select * from popspdesc ", M_OBJCONN, adOpenDynamic, adLockOptimistic
    While Not M_OBJRS.EOF
        cboPOPSP.AddItem M_OBJRS!KdNoProdPresented
        M_OBJRS.MoveNext
    Wend
    Set M_OBJRS = Nothing
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
Set M_OBJRS = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from ContactedDesc where KdNoProdPresented not like 'ptp%'", M_OBJCONN, adOpenDynamic, adLockOptimistic

M_OBJRS.Open "Select * from contacteddesc where KdNoProdPresented not like 'ptp%'", M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    While Not M_OBJRS.EOF
    '----tambahan 05 Maret 2007----'
         scombo = M_OBJRS("KdNoProdPresented")
            sKata = cmbContacted.Text
            ' initialisasi index
            If scombo = "BP-BROKEN PROMISE" Or scombo = "PTP-PROMISE TO PAY" Or scombo = "RP-REFUSE PAYMENT" Then
                  iIndex = 1
            ElseIf scombo = "POP-PROGRESS OF PAYMENT" Then
                  iIndex = 2
            ElseIf scombo = "SP-SETTLE PAYMENT" Then
                  iIndex = 3
            Else
                  iIndex = 4
            End If

            ' saring tampilan
            If iIndex = 1 Then
               If iIndex = 4 Or sKata = "POP-PROGRESS OF PAYMENT" Or sKata = "SP-SETTLED PAYMENT" Then
                  'lewat boo
               Else
                    If scombo = "BP-BROKEN PROMISE" And UCase(MDIForm1.Text2.Text) = "AGENT" Then
                    Else
                        cmbContacted.AddItem scombo
                    End If
               End If
            ElseIf iIndex = 2 Then
               If iIndex = 1 Or iIndex = 4 Or Left(sKata, 2) = "SP" Then
                  'lewat boo
               Else
                  cmbContacted.AddItem scombo
               End If
            ElseIf iIndex = 3 Then
                If UCase(MDIForm1.Text2.Text) = "AGENT" Then
                Else
                  cmbContacted.AddItem scombo
                End If
            Else
                  If sKata = "BP-BROKEN PROMISE" Or sKata = "PTP-PROMISE TO PAY" Or sKata = "POP-PROGRESS OF PAYMENT" Or sKata = "SP-SETTLED PAYMENT" Then
                     'lewat boo
                  Else
                     cmbContacted.AddItem scombo
                  End If
            End If
            M_OBJRS.MoveNext
    Wend
Set M_OBJRS = Nothing

If Left(cmbContacted.Text, 2) = "SP" Then
    'C_Contacted.Enabled = False
    'cmbContacted.Enabled = False
    C_NotContacted.Enabled = False
End If

If Left(cmbContacted.Text, 3) = "POP" Then
    'C_Contacted.Enabled = False
    'cmbContacted.Enabled = False
    C_NotContacted.Enabled = False
End If

'UNCONTACTED
Set M_OBJRS = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient
'If kontak = True Then
'    M_OBJRS.Open "Select * from UnContactedDesc where KdNoProdPresented IN ('NBP-NO BODY PICKUP','NA-NOT AVAILABLE')", M_OBJCONN, adOpenDynamic, adLockOptimistic
'Else
'    M_OBJRS.Open "Select * from UnContactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
'End If
If kontak = True Then
    M_OBJRS.Open "Select * from uncontacteddesc where kdnoprodpresented IN ('NBP-NO BODY PICKUP','NA-NOT AVAILABLE')", M_OBJCONN, adOpenDynamic, adLockOptimistic
ElseIf Left(VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(8), 2) = "NA" Then
    M_OBJRS.Open "Select * from uncontacteddesc  where kdnoprodpresented  IN ('NBP-NO BODY PICKUP','NA-NOT AVAILABLE')", M_OBJCONN, adOpenDynamic, adLockOptimistic
Else
    M_OBJRS.Open "Select * from uncontacteddesc ", M_OBJCONN, adOpenDynamic, adLockOptimistic
End If
    While Not M_OBJRS.EOF
        cmbUncontacted.AddItem M_OBJRS("KdNoProdPresented")
        'cmbDescUn.AddItem M_OBJRS("nmNoProdPresented")
        M_OBJRS.MoveNext
    Wend
Set M_OBJRS = Nothing

'Set M_OBJRS = New ADODB.Recordset
'If kontak = True Then
'    C_NotContacted.Enabled = False
'Else
'    M_OBJRS.Open "Select * from UnContactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    While Not M_OBJRS.EOF
'        cmbUncontacted.AddItem M_OBJRS("KdNoProdPresented")
'        'cmbDescUn.AddItem M_OBJRS("nmNoProdPresented")
'        M_OBJRS.MoveNext
'    Wend
'End If
'Set M_OBJRS = Nothing




'DISCOUNT

'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.Open "Select * from tblDiscount", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    While Not M_OBJRS.EOF
'        cmbDiscount.AddItem M_OBJRS("Description")
'        M_OBJRS.MoveNext
'    Wend
'Set M_OBJRS = Nothing

'NEXT ACTION
Set M_OBJRS = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient
M_OBJRS.Open "Select * from stsnextact", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not M_OBJRS.EOF
    cmbNextAct.AddItem M_OBJRS("NmStsNextAct")
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
'untuk 108
Set M_OBJRS = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient
M_OBJRS.Open "Select * from tbllayanantelkom", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not M_OBJRS.EOF
    CmbPhone.AddItem IIf(IsNull(M_OBJRS("Nolayanan")), "", M_OBJRS("Nolayanan"))
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing

End Sub

Sub isi_lastcall()
cbolastcall.Clear
Dim M_OBJRS As ADODB.Recordset
Set M_OBJRS = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient
M_OBJRS.Open "Select * from ContactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
While Not M_OBJRS.EOF
    cbolastcall.AddItem M_OBJRS("KdNoProdPresented")
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing

Set M_OBJRS = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient
M_OBJRS.Open "Select * from UnContactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
While Not M_OBJRS.EOF
    cbolastcall.AddItem M_OBJRS("KdNoProdPresented")
    M_OBJRS.MoveNext
Wend
Set M_OBJRS = Nothing
End Sub

Private Sub aktifphone()
AHomeAdd1(0).ReadOnly = False
AHomeAdd2(1).ReadOnly = False
txtHomeAdd1.ReadOnly = False
txtHomeAdd1A.ReadOnly = False
txtHomeAdd2.ReadOnly = False
txtHomeAdd2A.ReadOnly = False
AOfficeAdd(2).ReadOnly = False
AOfficeAdd(3).ReadOnly = False
txtOfficeAdd1.ReadOnly = False
txtOfficeAdd1A.ReadOnly = False
txtOfficeAdd2.ReadOnly = False
txtOfficeAdd2A.ReadOnly = False
AFaxAdd(4).ReadOnly = False
AFaxAdd(5).ReadOnly = False
txtFaxAdd1.ReadOnly = False
txtFaxAdd2.ReadOnly = False
txtMobileAdd1.ReadOnly = False
txtMobileAdd1A.ReadOnly = False
txtMobileAdd2.ReadOnly = False
txtMobileAdd2A.ReadOnly = False
txtECno.ReadOnly = False
txtECnoA.ReadOnly = False
End Sub

Private Sub aktifphoneAGENT()
If txtHomeAdd1.Value = "" Then
    txtHomeAdd1.ReadOnly = False
    AHomeAdd1(0).ReadOnly = False
End If
If txtHomeAdd1A.Value = "" Then
    txtHomeAdd1A.ReadOnly = False
    AHomeAdd1(0).ReadOnly = False
End If
If txtHomeAdd2.Value = "" Then
    txtHomeAdd2.ReadOnly = False
    AHomeAdd2(1).ReadOnly = False
End If
If txtHomeAdd2A.Value = "" Then
    txtHomeAdd2A.ReadOnly = False
    AHomeAdd2(1).ReadOnly = False
End If
If txtOfficeAdd1.Value = "" Then
    txtOfficeAdd1.ReadOnly = False
    AOfficeAdd(2).ReadOnly = False
End If
If txtOfficeAdd1A.Value = "" Then
    txtOfficeAdd1A.ReadOnly = False
    AOfficeAdd(2).ReadOnly = False
End If
If txtOfficeAdd2.Value = "" Then
    txtOfficeAdd2.ReadOnly = False
    AOfficeAdd(3).ReadOnly = False
End If
If txtOfficeAdd2A.Value = "" Then
    txtOfficeAdd2A.ReadOnly = False
    AOfficeAdd(3).ReadOnly = False
End If
If txtMobileAdd1.Value = "" Then
    txtMobileAdd1.ReadOnly = False
End If
If txtMobileAdd1A.Value = "" Then
    txtMobileAdd1A.ReadOnly = False
End If
If txtMobileAdd2.Value = "" Then
    txtMobileAdd2.ReadOnly = False
End If
If txtMobileAdd2A.Value = "" Then
    txtMobileAdd2A.ReadOnly = False
End If
If txtECno.Value = "" Then
    txtECno.ReadOnly = False
End If
If txtECnoA.Value = "" Then
    txtECnoA.ReadOnly = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim n%
For n = 1 To LstPayment.ListItems.Count
        If LstPayment.ListItems(n).SubItems(4) = "UNSCH" And regnego = True Then
            regnego = True
        End If
Next n

If regnego = False Or LstPayment.ListItems.Count = 0 Then
    kontak = False
    shedulePTP_Show = False
    regnego = False
    ' 'M_OBJCONN.Close
    M_OBJCONN.Close
    Set M_OBJCONN = Nothing
    M_OBJCONN.Open CMDSQLOPEN
    VIEW_MGMDATA.WindowState = 2
Else
        MsgBox "Lakukan PTP yang benar,Jumlah PTP harus >= Deal Payment " & txtPayment.Text & " , Atau data simpan dulu!!!"
        Cancel = 1
        Exit Sub
End If
End Sub
Private Sub ListView1_Click(Index As Integer)
Dim KET As String
Select Case Index
Case 0

Case 1
If listview1(1).ListItems.Count = 0 Then
Exit Sub
Else
   KET = TXtDetails.Text
      If Len(TXtDetails) = 0 Then
         TXtDetails.Text = " - " + listview1(1).SelectedItem.SubItems(1)
      Else
         TXtDetails.Text = KET + " - " + listview1(1).SelectedItem.SubItems(1)
      End If
End If
End Select
End Sub

Private Sub LstPayment_DblClick()
If LstPayment.ListItems.Count = 0 Then
Exit Sub
Else
Call SSCommand2_Click(1)
End If
End Sub
Private Sub LstVisit_DblClick()
 If LstVisit.ListItems.Count > 0 Then
            
        
           With FRM_UpdateVisit
                .Text1.Text = LstVisit.SelectedItem.SubItems(2)
                .Show vbModal
                

'                    M_DATA.UPDATE_NegoPTP M_OBJCONN, .TxtCustid.Text, .TDBDate1.Value, CStr(.TDBNumber1.Value), LstPayment.SelectedItem.SubItems(1)
'
'                    On Error GoTo add_error
'                    If M_DATA.ADD_OK Then
'                        'LstPayment.SelectedItem.SubItems(1) = ""
'                        LstPayment.SelectedItem.SubItems(2) = .TDBDate1.Value
'                        LstPayment.SelectedItem.SubItems(3) = .TDBNumber1.Value
'
'
'                    On Error GoTo 0
'                    End If
'                End If
               End With
Else
Exit Sub
End If

End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
TYPETELP = ""
   txtPhone.Text = GetNumber(CStr(AHome1.Value & txtHomeNo1.Value))
   If txtHomeNo1.Value <> "" Then
        txtPhoneA.Text = CStr(AHome1.Value & txtHomeNo1A.Value)
    Else
        txtPhoneA.Text = ""
    End If
   Option2.Value = False
   Option3.Value = False
   Option4.Value = False
   Option5.Value = False
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
TYPETELP = ""
   txtPhone.Text = GetNumber(CStr(AHome2.Value & txtHomeNo2.Value))
   If txtHomeNo2.Value <> "" Then
        txtPhoneA.Text = CStr(AHome2.Value & txtHomeNo2A.Value)
    Else
        txtPhoneA.Text = ""
    End If
   Option1.Value = False
   Option3.Value = False
   Option4.Value = False
   Option5.Value = False
End If
End Sub

Private Sub Option3_Click()
   If Option3.Value = True Then
   TYPETELP = ""
   txtPhone.Text = GetNumber(CStr(AOffice2.Value & txtOfficeNo2.Value))
   If txtOfficeNo2.Value <> "" Then
        txtPhoneA.Text = CStr(AOffice2.Value & txtOfficeNo2A.Value)
    Else
        txtPhoneA.Text = ""
   End If
   Option2.Value = False
   Option4.Value = False
   Option1.Value = False
   Option5.Value = False
   End If
End Sub

Private Sub Option4_Click()
   If Option4.Value = True Then
   TYPETELP = ""
   txtPhone.Text = GetNumber(CStr(AOffice1.Value & txtOfficeNo1.Value))
   If txtOfficeNo1.Value <> "" Then
        txtPhoneA.Text = CStr(AOffice1.Value & txtOfficeNo1A.Value)
    Else
        txtPhoneA.Text = ""
   End If
   Option2.Value = False
   Option3.Value = False
   Option1.Value = False
   Option5.Value = False
End If
End Sub

Private Sub Option5_Click()
 If Option5.Value = True Then
 TYPETELP = ""
   txtPhone.Text = GetNumber(CStr(txtMobileNo2.Value))
    If txtMobileNo2.Value <> "" Then
        txtPhoneA.Text = CStr(txtMobileNo2A.Value)
    Else
        txtPhoneA.Text = ""
   End If
   Option2.Value = False
   Option3.Value = False
   Option1.Value = False
   Option4.Value = False
   Option6.Value = False
   End If
End Sub

Private Sub Option6_Click()
 If Option6.Value = True Then
 TYPETELP = ""
   txtPhone.Text = GetNumber(CStr(txtMobileNo1.Value))
   If txtMobileNo1.Value <> "" Then
        txtPhoneA.Text = CStr(txtMobileNo1A.Value)
    Else
        txtPhoneA.Text = ""
   End If
   Option2.Value = False
   Option3.Value = False
   Option1.Value = False
   Option4.Value = False
   Option5.Value = False
   End If
End Sub

Private Sub Option7_Click(Index As Integer)
Select Case Index
Case 0
TxtAddress.Text = AddrNow.Text
Case 1
TxtAddress.Text = lblAddr.Text
Case 2
TxtAddress.Text = lblOfficeAddr.Text
End Select

End Sub

Private Sub Option8_Click(Index As Integer)
Select Case Index
Case 0
Frame8.Enabled = True
VisitYES
Case 1
VisitNo
Frame8.Enabled = False
End Select
End Sub

Private Sub SSCommand1_Click(Index As Integer)
Dim n As Integer
Select Case Index
  Case 0
'  If Len(CmbPhone.Text) < 2 Then
'    MsgBox "Pilihan No Telephone harus diisi"
'    CmbPhone.SetFocus
'    Exit Sub
'  End If
        Select Case CmbPhone
            Case "Hp"
                txtPhone.Text = txtMobileNo1.Value
                telpno = txtPhone.Text
            Case "Hp2"
                txtPhone.Text = txtMobileNo2.Value
                telpno = txtPhone.Text
            Case "HomePhone"
                If AHome1.Value = "021" Or AHome1.Value = "" Then
                    txtPhone.Text = Trim(txtHomeNo1.Value)
                Else
                    txtPhone.Text = Trim(AHome1.Value) & txtHomeNo1.Value
                End If
                telpno = txtPhone.Text
            Case "HomePhone2"
                If AHome1.Value = "021" Or AHome1.Value = "" Then
                    txtPhone.Text = txtHomeNo2.Value
                Else
                    txtPhone.Text = Trim(AHome1.Value) & Trim(txtHomeNo2.Value)
                End If
                telpno = txtPhone.Text
            Case "OfficePhone"
                If AOffice1.Value = "021" Or AOffice1.Value = "" Then
                    txtPhone.Text = txtOfficeNo1.Value
                Else
                    txtPhone.Text = AOffice1.Value & txtOfficeNo1.Value
                End If
                telpno = txtPhone.Text
            Case "OfficePhone2"
                If AOffice2.Value = "021" Or AOffice2.Value = "" Then
                    txtPhone.Text = txtOfficeNo2.Value
                Else
                    txtPhone.Text = AOffice1.Value & txtOfficeNo2.Value
                End If
                telpno = txtPhone.Text
            Case "EconPhone"
                If txtECno.ReadOnly = False And UCase(MDIForm1.Text2.Text) = "AGENT" Then
                    MsgBox "Simpan Data terlebih dahulu"
                    Exit Sub
                End If
                If Left(txtECno.Text, 3) = "021" Then
                 txtPhone.Text = Trim(Mid(txtECno.Value, 4, 16))
                 Else
                 txtPhone.Text = Trim(txtECno.Value)
                End If
                txtPhone.Text = txtECno.Value
                telpno = txtPhone.Text
            Case "AddHome1"
                If txtHomeAdd1A.ReadOnly = False And UCase(MDIForm1.Text2.Text) = "AGENT" Then
                    MsgBox "Simpan Data terlebih dahulu"
                    Exit Sub
                End If
                If AHomeAdd1(0).Value = "021" Or AHomeAdd1(0).Value = "" Then
                    txtPhone.Text = txtHomeAdd1.Value
                Else
                    txtPhone.Text = AHomeAdd1(0).Value & txtHomeAdd1.Value
                End If
                    telpno = txtPhone.Text
            Case "AddHome2"
                If txtHomeAdd2A.ReadOnly = False And UCase(MDIForm1.Text2.Text) = "AGENT" Then
                    MsgBox "Simpan Data terlebih dahulu"
                    Exit Sub
                End If
                If AHomeAdd2(1).Value = "021" Or AHomeAdd2(1).Value = "" Then
                    txtPhone.Text = txtHomeAdd2.Value
                Else
                    txtPhone.Text = AHomeAdd2(1).Value & txtHomeAdd2.Value
                End If
                telpno = txtPhone.Text
            Case "AddOffice1"
                If txtOfficeAdd1A.ReadOnly = False And UCase(MDIForm1.Text2.Text) = "AGENT" Then
                    MsgBox "Simpan Data terlebih dahulu"
                    Exit Sub
                End If
                If AOfficeAdd(2).Value = "021" Or AOfficeAdd(2).Value = "" Then
                    txtPhone.Text = txtOfficeAdd1.Value
                Else
                    txtPhone.Text = AOfficeAdd(2).Value & txtOfficeAdd1.Value
                End If
                telpno = txtPhone.Text
            Case "AddOffice2"
                If txtOfficeAdd2A.ReadOnly = False And UCase(MDIForm1.Text2.Text) = "AGENT" Then
                    MsgBox "Simpan Data terlebih dahulu"
                    Exit Sub
                End If
                If AOfficeAdd(3).Value = "021" Or AOfficeAdd(3).Value = "" Then
                    txtPhone.Text = Trim(txtOfficeAdd2.Value)
                Else
                    txtPhone.Text = Trim(AOfficeAdd(3).Value) & Trim(txtOfficeAdd2.Value)
                End If
                telpno = txtPhone.Text
            Case "AddMobile1"
                If txtMobileAdd1A.ReadOnly = False And UCase(MDIForm1.Text2.Text) = "AGENT" Then
                    MsgBox "Simpan Data terlebih dahulu"
                    Exit Sub
                End If
                txtPhone.Text = Trim(txtMobileAdd1.Value)
                telpno = txtPhone.Text
            Case "AddMobile2"
                If txtMobileAdd2A.ReadOnly = False And UCase(MDIForm1.Text2.Text) = "AGENT" Then
                    MsgBox "Simpan Data terlebih dahulu"
                    Exit Sub
                End If
                txtPhone.Text = Trim(txtMobileAdd2.Value)
                telpno = txtPhone.Text
            Case Else
                txtPhone.Text = Replace(CmbPhone.Text, " ", "")
        End Select
    MDIForm1.ActionCTI ("DIAL|49682" & GetNumber(CStr(Replace(txtPhone.Text, " ", ""))) & "|" & Trim(frmCC_Colection.lblCustId.Caption) & "|" & Trim(frmCC_Colection.lblCustId.Caption))
    CMDSQL = "Insert Into tblphonemonitorhst(UserId, CustId, NamaCh,StartDate, TelpNo, Recsource) Values ('" + MDIForm1.Text1.Text + "' , '" + frmCC_Colection.lblCustId.Caption + "','" + frmCC_Colection.lblNama.Caption + "', '" + Format(CStr(MDIForm1.TDBDate1.Value), "mm/dd/yyyy") & " " & Format(Now, "hh:nn") + "' , '" + Replace(txtPhone.Text, " ", "") + "' ,'" + frmCC_Colection.lblRecsource.Caption + "')"
    M_OBJCONN.Execute CMDSQL
    MDIForm1.CmbNo.Text = ""
    stscall = True
    TYPETELP = ""
   Case 2
        V_SAVE = CEK_DATA_VALID
        If V_SAVE = False Then
            Exit Sub
        Else
        End If
        If ADD_CUST Then
        Else
            Call CEK_UPDATE_PELANGGAN
            stscall = False
            Call isi_datapayment
        End If
   Case 3
    kontak = False
        For n = 1 To LstPayment.ListItems.Count
            If LstPayment.ListItems(n).SubItems(4) = "UNSCH" And regnego = True Then
                regnego = True
            End If
        Next n
        If regnego = True And LstPayment.ListItems.Count <> 0 Then
            MsgBox "Lakukan PTP yang benar, Jumlah PTP harus >= Deal Payment " & txtPayment.Text & " ,Atau data simpan dulu!!!"
            Exit Sub
        End If
        Unload Me
    Case 1
        MDIForm1.ActionCTI ("HANGUP")
End Select
End Sub

Public Sub Show_NEGOPTP()
Dim ShowList As New ADODB.Recordset
Dim listitem As listitem
Dim CMDSQL As String
Dim TOTPTP As Currency
Dim ssql As String
ssql = "SELECT CUSTID,sum(PAYMENT) as Jum FROM tbllunas WHERE custid = '" + lblCustId.Caption + "' GROUP BY CUSTID"
ShowList.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic
If ShowList.BOF And ShowList.EOF Then
    TOTPTP = 0
Else
    TOTPTP = IIf(IsNull(ShowList!jum), 0, ShowList!jum)
End If
'If ShowList.BOF And ShowList.EOF Then
'    'CMDSQL = "SELECT * FROM TBLNEGOPTP WHERE custid = '" + lblCustId.Caption + "'"
'    'AND CUSTID NOT IN (SELECT CUSTID FROM tbllunas)"
'    CMDSQL = "SELECT DISTINCT TBLNEGOPTP.PROMISEDATE,TBLNEGOPTP.ID,TBLNEGOPTP.PROMISEPAY,TBLNEGOPTP.TYPE FROM TBLNEGOPTP,tbllunas WHERE "
'    CMDSQL = CMDSQL + "tbllunas.CUSTID<>TBLNEGOPTP.CUSTID AND TBLNEGOPTP.CUSTID='" + lblCustId.Caption + "' order by TBLNEGOPTP.promisedate desc"
'Else
'    CMDSQL = "SELECT distinct TBLNEGOPTP.PROMISEDATE,TBLNEGOPTP.PROMISEPAY,TBLNEGOPTP.ID,TBLNEGOPTP.TYPE "
'    CMDSQL = CMDSQL + "FROM VWLISTPTP,TBLNEGOPTP WHERE TBLNEGOPTP.CUSTID=VWLISTPTP.CUSTID AND "
'    CMDSQL = CMDSQL + "VWLISTPTP.PAYDATE<TBLNEGOPTP.PROMISEDATE AND TBLNEGOPTP.CUSTID='" + lblCustId.Caption + "' order by TBLNEGOPTP.promisedate desc"
'End If
CMDSQL = "SELECT * FROM tblnegoptp where custid = '" + lblCustId.Caption + "' order by promisedate"

Set ShowList = New ADODB.Recordset
ShowList.CursorLocation = adUseClient
ShowList.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic

LstPayment.ListItems.Clear
Dim n As Currency
While Not ShowList.EOF
    Set listitem = LstPayment.ListItems.ADD(, , "")
        listitem.SubItems(1) = CStr(IIf(IsNull(ShowList!ID), "", (ShowList!ID)))
        listitem.SubItems(2) = CStr(IIf(IsNull(ShowList!PromiseDate), "", Format(ShowList!PromiseDate, "dd/mm/yyyy")))
        listitem.SubItems(3) = CStr(IIf(IsNull(ShowList!PromisePay), "", (ShowList!PromisePay)))
        n = n + Val(listitem.SubItems(3))
        If n <= TOTPTP Then
            listitem.ListSubItems(1).ForeColor = vbRed
            listitem.ListSubItems(2).ForeColor = vbRed
            listitem.ListSubItems(3).ForeColor = vbRed
        End If
        
        listitem.SubItems(4) = IIf(IsNull(ShowList!Type), "", ShowList!Type)
        listitem.SubItems(5) = CStr(IIf(IsNull(ShowList!inputdate), "", Format(ShowList!inputdate, "dd/mm/yyyy")))
     ShowList.MoveNext
Wend

Set ShowList = Nothing
End Sub
Public Sub show_cust()
Dim listitem As listitem
Dim M_DATA As New CLS_FRMCUST_CC
Dim m_cust1 As ADODB.Recordset
Dim m_cust2 As ADODB.Recordset
Dim CMDSQL As String
Dim CMDSQL2 As String
Dim sPending As String
'On Error GoTo HELL:
'CMDSQL = "SELECT mgm.*, mgm_DETAIL.* FROM mgm INNER JOIN "
'CMDSQL = CMDSQL + "mgm_DETAIL ON mgm.CUSTID = dbo.mgm_DETAIL.CUSTID"

CMDSQL = "select * from mgm"
'CMDSQL2 = "select * from mgm_detail"

Set m_cust = New ADODB.Recordset
'Set m_cust2 = New ADODB.Recordset
m_cust.CursorLocation = adUseClient
'm_cust2.CursorLocation = adUseClient
If shedulePTP_Show = True Then
    CMDSQL = CMDSQL + " where custid ='" & MDIForm1.LstGrade.SelectedItem.SubItems(1) & "'"
    m_cust.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
   
Else
    CMDSQL = CMDSQL + " where custid ='" & VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(1) & "'"
    m_cust.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
    'CMDSQL2 = CMDSQL2 + " where custid ='" & VIEW_mgmDATA.LstVwSearchmgm.SelectedItem.SubItems(1) & "'"
    'm_cust2.Open CMDSQL2, M_OBJCONN, adOpenDynamic, adLockOptimistic
    'm_cust.Open "Select * from mgm where custid='" & VIEW_mgmDATA.LstVwSearchmgm.SelectedItem.SubItems(1) & "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
End If

'tampilkan data tabel mgm
If Not m_cust.EOF Then

    lblstatus.Caption = IIf(IsNull(m_cust("statusprior")), "", "Status : " & m_cust("statusprior"))
    lblCustId.Caption = IIf(IsNull(m_cust("CUSTID")), "", m_cust("CUSTID"))
    
    'sql = "delete  from tblnegoptp where custid in (select custid from tbllunas where custid ='" + IIf(IsNull(m_cust("CUSTID")), "", m_cust("CUSTID")) + "')"
    TxtCustid.Text = IIf(IsNull(m_cust("CUSTID")), "", m_cust("CUSTID"))
    TxtName.Text = IIf(IsNull(m_cust("NAME")), "", m_cust("NAME"))
    lblaoc.Caption = IIf(IsNull(m_cust("agent")), "", m_cust("Agent"))
    LblInterest.Caption = Format(IIf(IsNull(m_cust("INTEREST")), "0", m_cust("INTEREST")), "##,###")
    LblFees.Caption = Format(IIf(IsNull(m_cust("FEES")), "0", m_cust("FEES")), "##,###")
    lblregion.Caption = IIf(IsNull(m_cust("region")), "", m_cust("region"))
    lblaging.Caption = IIf(IsNull(m_cust("Aging")), "            ", m_cust("Aging"))
    lblwilling.Caption = IIf(IsNull(m_cust("Willing_Ness")), "              ", m_cust("Willing_Ness"))
    lblRecsource.Caption = IIf(IsNull(m_cust("RECSOURCE")), "", m_cust("RECSOURCE"))
    LBLEXP.Caption = IIf(IsNull(m_cust("date_into_clas")), "", "Expire date " & Format(DateAdd("d", 60, m_cust("date_into_clas")), "dd-mm-yyyy"))
     LblRiskLevel.Caption = IIf(IsNull(m_cust("RiskLevel")), "", m_cust("RiskLevel"))
    lblPriority.Caption = IIf(IsNull(m_cust("Priority")), "", m_cust("Priority"))
    lblNama.Caption = IIf(IsNull(m_cust("NAME")), "", m_cust("NAME"))
    lblCardNo.Caption = IIf(IsNull(m_cust("NoCard")), "", m_cust("NoCard"))
    lblID.Caption = IIf(IsNull(m_cust("ktpno")), "", m_cust("ktpno"))
    'lblDate.Value = IIf(IsNull(m_cust("BIRTHD")), "", Format(m_cust("BIRTHD"), "dd-mmm-yyyy"))
    LblDOB.Caption = IIf(IsNull(m_cust("DOB")), "", Left(m_cust("DOB"), 10))
    lblAddr.Text = IIf(IsNull(m_cust("ADDRNOW")), "", m_cust("ADDRNOW"))
'
  vrcek = IIf(IsNull(m_cust("f_cek")), "", m_cust("f_cek"))
    If Left(vrcek, 2) = "BP" Then
    cboPOPSP.Enabled = False
'        FrmContacted.Enabled = False
'        C_Contacted.Enabled = False
'        cmbContacted.Enabled = False
'        cmbDescCon.Enabled = False
     End If
    
    lblOfficeAddr.Text = IIf(IsNull(m_cust("ADDRPT")), "", m_cust("ADDRPT"))
    lblZIP.Caption = IIf(IsNull(m_cust("ZIPNOW")), "", m_cust("ZIPNOW"))
    lblNoCard.Caption = IIf(IsNull(m_cust("NoCard")), "", m_cust("NoCard"))
    lblNoPay.Caption = IIf(IsNull(m_cust("NoPay")), "", m_cust("NoPay"))
    lblPromPA.Value = IIf(IsNull(m_cust("Principal")), "", m_cust("Principal"))
    lblOpenDate.Value = IIf(IsNull(m_cust("OpenDate")), "", m_cust("OpenDate"))
    lblLastBill.Value = IIf(IsNull(m_cust("LastBill")), "", m_cust("LastBill"))
    lblLcAtm.Value = IIf(IsNull(m_cust("LcATMP")), "", m_cust("LcATMP"))
    txttenor.Value = IIf(IsNull(m_cust("tenor")), 0, m_cust("tenor"))
    vrtenor = IIf(IsNull(m_cust("tenor")), 0, m_cust("tenor"))
    lblBrokenPromised.Caption = IIf(IsNull(m_cust("BrokenPromise")), "", m_cust("BrokenPromise"))
    lblBD.Value = IIf(IsNull(m_cust("B_D")), "", m_cust("B_D"))
    lblLimit.Value = IIf(IsNull(m_cust("Limit")), "", m_cust("Limit"))
   
    If listview1(0).ListItems.Count = 0 Then
    lblPayDt.Value = IIf(IsNull(m_cust("Pay_Dt")), "", m_cust("Pay_Dt"))
    End If
    
    
    If listview1(0).ListItems.Count = 0 Then
    lblLastPay.Value = IIf(IsNull(m_cust("LastPay")), "", m_cust("LastPay"))
    End If
    
    lblTtlPay.Value = IIf(IsNull(m_cust("TtlPay")), "", m_cust("TtlPay"))
    lblAmount.Value = IIf(IsNull(m_cust("AmountWo")), "", Format(m_cust("AmountWo"), "##.##0"))
    
    
    txtHomeNo1.Value = IIf(IsNull(m_cust("HOMENO")), "", m_cust("HOMENO"))
    AHome1.Value = IIf(IsNull(m_cust("AHOMENO")), "", m_cust("AHOMENO"))
    
    Cmbwith.Text = IIf(IsNull(m_cust("contacwith")), "", m_cust("contacwith"))
    
    If IsNull(m_cust("HOMENO")) = False And m_cust("HOMENO") <> "" Then
        'txtHomeNo1A.Value = Left(m_cust("HOMENO"), Len(m_cust("HOMENO")) - 3) & "XXX"
        txtHomeNo1A.Value = Left(m_cust("HOMENO"), 4) & "BBB" & Mid(m_cust("HOMENO"), 8, 15)
        CmbPhone.AddItem "HomePhone"
    End If
    
    AHome2.Value = IIf(IsNull(m_cust("AHOMENO2")), "", m_cust("AHOMENO2"))
    txtHomeNo2.Value = IIf(IsNull(m_cust("HOMENO2")), "", m_cust("HOMENO2"))
    If IsNull(m_cust("HOMENO2")) = False And m_cust("HOMENO2") <> "" Then
        'txtHomeNo2A.Value = Left(m_cust("HOMENO2"), Len(m_cust("HOMENO2")) - 3) & "XXX"
        txtHomeNo2A.Value = Left(m_cust("HOMENO2"), 4) & "BBB" & Mid(m_cust("HOMENO2"), 8, 15)
        CmbPhone.AddItem "HomePhone2"
    End If
    AOffice1.Value = IIf(IsNull(m_cust("AOFFICENO")), "", m_cust("AOFFICENO"))
    txtOfficeNo1.Value = IIf(IsNull(m_cust("OFFICENO")), "", m_cust("OFFICENO"))
    If IsNull(m_cust("OFFICENO")) = False And m_cust("OFFICENO") <> "" Then
        'txtOfficeNo1A.Value = Left(m_cust("OFFICENO"), Len(m_cust("OFFICENO")) - 3) & "XXX"
        txtOfficeNo1A.Value = Left(m_cust("OFFICENO"), 4) & "BBB" & Mid(m_cust("OFFICENO"), 8, 15)
        CmbPhone.AddItem "OfficePhone"
    End If
    AOffice2.Value = IIf(IsNull(m_cust("AOFFICENO2")), "", m_cust("AOFFICENO2"))
    txtOfficeNo2.Value = IIf(IsNull(m_cust("OFFICENO2")), "", m_cust("OFFICENO2"))
    If IsNull(m_cust("OFFICENO2")) = False And m_cust("OFFICENO2") <> "" Then
        'txtOfficeNo2A.Value = Left(m_cust("OFFICENO2"), Len(m_cust("OFFICENO2")) - 3) & "XXX"
        txtOfficeNo2A.Value = Left(m_cust("OFFICENO2"), 4) & "BBB" & Mid(m_cust("OFFICENO2"), 8, 15)
        CmbPhone.AddItem "OfficePhone2"
    End If
    txtMobileNo1.Value = IIf(IsNull(m_cust("MOBILENO")), "", m_cust("MOBILENO"))
    If IsNull(m_cust("MOBILENO")) = False And m_cust("MOBILENO") <> "" Then
        'txtMobileNo1A.Value = Left(m_cust("MOBILENO"), Len(m_cust("MOBILENO")) - 3) & "XXX"
        txtMobileNo1A.Value = Left(m_cust("MOBILENO"), 4) & "BBB" & Mid(m_cust("MOBILENO"), 8, 15)
        CmbPhone.AddItem "Hp"
    End If
    txtMobileNo2.Value = IIf(IsNull(m_cust("MOBILENO2")), "", m_cust("MOBILENO2"))
    If IsNull(m_cust("MOBILENO2")) = False And m_cust("MOBILENO2") <> "" Then
        'txtMobileNo2A.Value = Left(m_cust("MOBILENO2"), Len(m_cust("MOBILENO2")) - 3) & "XXX"
        txtMobileNo2A.Value = Left(m_cust("MOBILENO2"), 4) & "BBB" & Mid(m_cust("MOBILENO2"), 8, 15)
        CmbPhone.AddItem "Hp2"
    End If
    AHomeAdd1(0).Value = IIf(IsNull(m_cust("AHOMENOADD1")), "", m_cust("AHOMENOADD1"))
    AHomeAdd2(1).Value = IIf(IsNull(m_cust("AHOMENOADD2")), "", m_cust("AHOMENOADD2"))
    AOfficeAdd(2).Value = IIf(IsNull(m_cust("AOFFICENOADD1")), "", m_cust("AOFFICENOADD1"))
    AOfficeAdd(3).Value = IIf(IsNull(m_cust("AOFFICENOADD2")), "", m_cust("AOFFICENOADD2"))
    AFaxAdd(4).Value = IIf(IsNull(m_cust("AFAXNOADD1")), "", m_cust("AFAXNOADD1"))
    AFaxAdd(5).Value = IIf(IsNull(m_cust("AFAXNOADD2")), "", m_cust("AFAXNOADD2"))
    txtHomeAdd1.Value = IIf(IsNull(m_cust("HOMENOADD1")), "", m_cust("HOMENOADD1"))
    If IsNull(m_cust("HOMENOADD1")) = False And m_cust("HOMENOADD1") <> "" Then
        txtHomeAdd1A.Value = Left(m_cust("HOMENOADD1"), 4) & "BBB" & Mid(m_cust("HOMENOADD1"), 8, 15)
        CmbPhone.AddItem "AddHome1"
    Else
        txtHomeAdd1.Visible = True
        txtHomeAdd1A.Visible = False
    End If
    txtHomeAdd2.Value = IIf(IsNull(m_cust("HOMENOADD2")), "", m_cust("HOMENOADD2"))
    If IsNull(m_cust("HOMENOADD2")) = False And m_cust("HOMENOADD2") <> "" Then
        txtHomeAdd2A.Value = Left(m_cust("HOMENOADD2"), 4) & "BBB" & Mid(m_cust("HOMENOADD2"), 8, 15)
        CmbPhone.AddItem "AddHome2"
    Else
        txtHomeAdd2A.Visible = False
        txtHomeAdd2.Visible = True
    End If
    txtOfficeAdd1.Value = IIf(IsNull(m_cust("OFFICENOADD1")), "", m_cust("OFFICENOADD1"))
    If IsNull(m_cust("OFFICENOADD1")) = False And m_cust("OFFICENOADD1") <> "" Then
        txtOfficeAdd1A.Value = Left(m_cust("OFFICENOADD1"), 4) & "BBB" & Mid(m_cust("OFFICENOADD1"), 8, 15)
        CmbPhone.AddItem "AddOffice1"
    Else
        txtOfficeAdd1A.Visible = False
        txtOfficeAdd1.Visible = True
    End If
    txtOfficeAdd2.Value = IIf(IsNull(m_cust("OFFICENOADD2")), "", m_cust("OFFICENOADD2"))
    If IsNull(m_cust("OFFICENOADD2")) = False And m_cust("OFFICENOADD2") <> "" Then
        txtOfficeAdd2A.Value = Left(m_cust("OFFICENOADD2"), 4) & "BBB" & Mid(m_cust("OFFICENOADD2"), 8, 15)
        CmbPhone.AddItem "AddOffice2"
    Else
        txtOfficeAdd2.Visible = True
        txtOfficeAdd2A.Visible = False
    End If
    txtMobileAdd1.Value = IIf(IsNull(m_cust("MOBILENOADD1")), "", m_cust("MOBILENOADD1"))
    If IsNull(m_cust("MOBILENOADD1")) = False And m_cust("MOBILENOADD1") <> "" Then
        txtMobileAdd1A.Value = Left(m_cust("MOBILENOADD1"), 4) & "BBB" & Mid(m_cust("MOBILENOADD1"), 8, 15)
        CmbPhone.AddItem "AddMobile1"
    Else
        txtMobileAdd1.Visible = True
        txtMobileAdd1A.Visible = False
    End If
    txtMobileAdd2.Value = IIf(IsNull(m_cust("MOBILENOADD2")), "", m_cust("MOBILENOADD2"))
    If IsNull(m_cust("MOBILENOADD2")) = False And m_cust("MOBILENOADD2") <> "" Then
        txtMobileAdd2A.Value = Left(m_cust("MOBILENOADD2"), 4) & "BBB" & Mid(m_cust("MOBILENOADD2"), 8, 15)
        CmbPhone.AddItem "AddMobile2"
    Else
        txtMobileAdd2.Visible = True
        txtMobileAdd2A.Visible = False
    End If
    txtFaxAdd1.Value = IIf(IsNull(m_cust("FAXNOADD1")), "", m_cust("FAXNOADD1"))
    txtFaxAdd2.Value = IIf(IsNull(m_cust("FAXNOADD2")), "", m_cust("FAXNOADD2"))
    AddrNow.Text = IIf(IsNull(m_cust("TxtPtpAddr")), "", m_cust("TxtPtpAddr"))
    LblLunas.Caption = IIf(IsNull(m_cust!tgllunas), "", "TELAH LUNAS")
    TxtEC.Text = IIf(IsNull(m_cust!ec_name), "", m_cust!ec_name)
    txtECno.Value = IIf(IsNull(m_cust!ec_telp), "", m_cust!ec_telp)
    If IsNull(m_cust("ec_telp")) = False And m_cust("ec_telp") <> "" Then
        txtECnoA.Value = Left(m_cust("ec_telp"), 4) & "BBB" & Mid(m_cust("ec_telp"), 8, 15)
        CmbPhone.AddItem "EconPhone"
    Else
        txtECnoA.Visible = False
        txtECno.Visible = True
    End If
    txtECAdd.Text = IIf(IsNull(m_cust!ECAddr), "", m_cust!ECAddr)
    cbolastcall.Text = IIf(IsNull(m_cust!statuscall), "", m_cust!statuscall)
'    If cbolastcall.Text = "" Then
'        Call isi_lastcall
'    End If
' cari extension
    If InStr(1, txtOfficeNo1.Value, "X", vbTextCompare) > 0 Then
        TxtExt1.Text = Right(txtOfficeNo1.Value, Len(txtOfficeNo1.Value) - InStr(1, txtOfficeNo1.Value, "X", vbTextCompare))
    End If
    If InStr(1, txtOfficeNo2.Value, "X", vbTextCompare) > 0 Then
        TxtExt2.Text = Right(txtOfficeNo2.Value, Len(txtOfficeNo2.Value) - InStr(1, txtOfficeNo2.Value, "X", vbTextCompare))
    End If
    If InStr(1, txtOfficeAdd1.Value, "X", vbTextCompare) > 0 Then
        TxtExt3.Text = Right(txtOfficeAdd1.Value, Len(txtOfficeAdd1.Value) - InStr(1, txtOfficeAdd1.Value, "X", vbTextCompare))
    End If
    If InStr(1, txtOfficeAdd2.Value, "X", vbTextCompare) > 0 Then
        TxtExt4.Text = Right(txtOfficeAdd2.Value, Len(txtOfficeAdd2.Value) - InStr(1, txtOfficeAdd2.Value, "X", vbTextCompare))
    End If
    If UCase(MDIForm1.Text2.Text) = "AGENT" Then
        If Len(txtECno.Value) > 2 Then
            txtECno.ReadOnly = True
        End If
        If Len(txtHomeAdd1.Value) > 2 Then
            txtHomeAdd1.ReadOnly = True
        End If
        If Len(txtHomeAdd2.Value) > 2 Then
            txtHomeAdd2.ReadOnly = True
        End If
        If Len(txtOfficeAdd1.Value) > 2 Then
            txtOfficeAdd1.ReadOnly = True
        End If
        If Len(txtOfficeAdd2.Value) > 2 Then
            txtOfficeAdd2.ReadOnly = True
        End If
        If Len(txtMobileAdd1.Value) > 2 Then
            txtMobileAdd1.ReadOnly = True
        End If
        If Len(txtMobileAdd2.Value) > 2 Then
            txtMobileAdd2.ReadOnly = True
        End If
        If Len(txtECno.Value) > 2 Then
            txtECno.ReadOnly = True
        End If
    End If
    cmbNextAct.Text = IIf(IsNull(m_cust("NEXTACT")), "", m_cust("NEXTACT"))
    
    sPending = CStr(Trim(IIf(IsNull(m_cust!f_Pending), "", m_cust!f_Pending)))
     If sPending = "Pending" Then
         chkAppv(0).Value = 0
    End If
    
    Select Case m_cust!RECSTATUS
        Case "V"
            C_VALID.Value = 1
            cbovalid.Text = IIf(IsNull(m_cust("KETHSLKERJA")), "", m_cust("KETHSLKERJA"))
            cbodescvalid.Text = IIf(IsNull(m_cust("KETHSLKERJADESC")), "", m_cust("KETHSLKERJADESC"))
        Case "N"
            C_NotContacted.Value = 1
            cmbUncontacted.Text = IIf(IsNull(m_cust("KETHSLKERJA")), "", m_cust("KETHSLKERJA"))
            cmbDescUn.Text = IIf(IsNull(m_cust("KETHSLKERJADESC")), "", m_cust("KETHSLKERJADESC"))
        Case "C"
            C_Contacted.Value = 1
            kontak = True
            If MDIForm1.Text2 = "Agent" Then
                If Left(vrcek, 3) = "POP" Then
                    C_SKIP.Enabled = False
                    C_VALID.Enabled = False
                    cboPOPSP.Enabled = False
                    FrmPayment.Enabled = True
                    C_Payment.Value = 1
                End If
            End If
            cmbContacted.Text = IIf(IsNull(m_cust("KETHSLKERJA")), "", m_cust("KETHSLKERJA"))
      Case "P"
            C_PTP.Value = 1
            cboPTP.Text = IIf(IsNull(m_cust("KETHSLKERJA")), "", m_cust("KETHSLKERJA"))
            'cmbDescCon.Text = IIf(IsNull(m_cust("KETHSLKERJADESC")), "", m_cust("KETHSLKERJADESC"))
            If MDIForm1.Text2 = "Agent" Then
                C_VALID.Enabled = False
                C_Contacted.Enabled = False
                FrMValid.Enabled = False
                C_SKIP.Enabled = False
                FrmSKIP.Enabled = False
            End If
         Case "S"
            C_SKIP.Value = 1
            cboskip.Text = IIf(IsNull(m_cust("KETHSLKERJA")), "", m_cust("KETHSLKERJA"))
            cbodescskip.Text = IIf(IsNull(m_cust("KETHSLKERJADESC")), "", m_cust("KETHSLKERJADESC"))
         Case "O"
            'C_POPSP.Value = 1
            cboPOPSP.Text = IIf(IsNull(m_cust("KETHSLKERJA")), "", m_cust("KETHSLKERJA"))
            'cmbDescCon.Text = IIf(IsNull(m_cust("KETHSLKERJADESC")), "", m_cust("KETHSLKERJADESC"))      cmbDescCon.Text = IIf(IsNull(m_cust("KETHSLKERJADESC")), "", m_cust("KETHSLKERJADESC"))
     End Select
     
    If MDIForm1.Text2 = "Agent" Then
        If IIf(IsNull(m_cust!RECSTATUS), "", m_cust!RECSTATUS) <> "O" Then
            frmpopsp.Enabled = False
           cboPOPSP.Enabled = False
        End If
    End If
        If IIf(IsNull(m_cust!F_CEK), "", Left(m_cust!F_CEK, 3)) = "PTP" Or Left(m_cust!F_CEK, 3) = "POP" Or Left(m_cust!F_CEK, 3) = "SP-" Or Left(m_cust!F_CEK, 3) = "PRE" Then
            C_Payment.Value = 1
            TdbPTP.Value = IIf(IsNull(m_cust!TdbDatePTP), "", m_cust!TdbDatePTP)
            vrtdbdateptp = IIf(IsNull(m_cust!TdbDatePTP), "", m_cust!TdbDatePTP)
            vrdateptp = IIf(IsNull(m_cust!TdbDatePTP), "", m_cust!TdbDatePTP)
            TDBDate3.Value = IIf(IsNull(m_cust!dateptp), "", m_cust!dateptp)
            txtPayment.Value = IIf(IsNull(m_cust!ttlptp), 0, m_cust!ttlptp)
            vrttlptp = IIf(IsNull(m_cust!ttlptp), 0, m_cust!ttlptp)
            Tdabamoint.Value = IIf(IsNull(m_cust!amountptp), 0, m_cust!amountptp)
            vramount = IIf(IsNull(m_cust!amountptp), 0, m_cust!amountptp)
            TxtPayment2.Value = IIf(IsNull(m_cust!ttlptp), 0, m_cust!ttlptp) 'tampilkan di detail payment
            cmbDiscount.Text = IIf(IsNull(m_cust!discpersen), 0, m_cust!discpersen)
            vrdiskon = IIf(IsNull(m_cust!discpersen), 0, m_cust!discpersen)
            CmbBaseOn.Text = IIf(IsNull(m_cust!CmbBaseOn), "", m_cust!CmbBaseOn)
            vrbaseon = IIf(IsNull(m_cust!CmbBaseOn), "", m_cust!CmbBaseOn)
            'TdbDatePTP.Value = IIf(IsNull(m_cust!TGLINCOMING), "", m_cust!TGLINCOMING)
        Else
        End If
End If
Call Custid_Double
'Set m_cust1 = M_DATA.QUERY_HIST_CUST(M_OBJCONN, "CUSTID = '" + lblCustId.Caption + "'", MDIForm1.Text2.Text)
Set m_cust1 = M_DATA.QUERY_HIST_CUST(M_OBJCONN, "CUSTID = '" + lblCustId.Caption + "'")
While Not m_cust1.EOF
    'Set listitem = ListView1(1).ListItems.ADD(, , Left(m_cust1("TGL"), 4) & "/" & Mid(m_cust1("TGL"), 5, 2) & "/" & IIf(IsNull(m_cust1("TGL")), "", Mid(m_cust1("TGL"), 7, 2)) & " " & IIf(IsNull(m_cust1("TGL")), "", Mid(m_cust1("TGL"), 9, 2)) & ":" & Right(m_cust1("TGL"), 2))
     Set listitem = listview1(1).ListItems.ADD(, , Format(IIf(IsNull(m_cust1("TGL")), "", m_cust1!TGL), "mm-dd-yyyy hh:mm:ss"))
        listitem.SubItems(1) = IIf(IsNull(m_cust1("HST")), "", m_cust1("HST"))
        listitem.SubItems(2) = IIf(IsNull(m_cust1("AGENT")), "", m_cust1("AGENT"))
        listitem.SubItems(3) = IIf(IsNull(m_cust1("KodeDs")), "", m_cust1("KodeDs"))
        'listitem.SubItems(4) = IIf(IsNull(m_cust1("f_cek")), "", m_cust1("f_cek"))
m_cust1.MoveNext
Wend

Call isi_datapayment
Call Show_NEGOPTP
Call Show_Visit
Set M_OBJRS = New ADODB.Recordset
M_OBJRS.CursorLocation = adUseClient
M_OBJRS.Open "Select custid, sum(payment) as jml from tbllunas where custid = '" + lblCustId.Caption + "' GROUP BY CUSTID", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not M_OBJRS.EOF
        TxtAfterPay.Value = IIf(IsNull(M_OBJRS("jml")), 0, M_OBJRS("jml"))
        M_OBJRS.MoveNext
Wend
 
 'hitung sisa hutang
 txtSisaHutang.Value = Val(TxtPayment2.Value) - Val(TxtAfterPay.Value)
 
 '---------->> hitung PRINCIPLE & AMOUNTWO  after pay  <<-----------------
 If TxtAfterPay.Value = 0 Then
    txtPrinciple_A.Value = 0
    txtAmountwo_A.Value = 0
    Else
    If lblPromPA.ValueIsNull Or lblAmount.ValueIsNull Then
    Exit Sub
    End If
  txtPrinciple_A.Value = Val(lblPromPA.Value) - Val(TxtAfterPay.Value)
  txtAmountwo_A.Value = Val(lblAmount.Value) - Val(TxtAfterPay.Value)
 End If
 
    If lblAmount.ValueIsNull Then
           Woafter.Value = 0
       Else
           Woafter.Value = lblAmount - TxtAfterPay.Value
    End If
  
    If listview1(0).ListItems.Count <> 0 Then
          lblPayDt.Value = listview1(0).ListItems(listview1(0).ListItems.Count).Text
          lblLastPay.Value = listview1(0).ListItems(listview1(0).ListItems.Count).SubItems(1)
          LBLEXP.Caption = "Expire Date " + glexp
    End If
 
 
    Set m_cust = Nothing
    Set M_OBJRS = Nothing

Exit Sub
'HELL:
   'MsgBox Err.Description
' Resume
 Set M_OBJRS = Nothing
Set m_cust = Nothing
End Sub


Private Sub isi_datapayment()
Dim m_cust2 As New ADODB.Recordset
Dim NilaiAfterPay As Currency
Dim M_DATA As New CLS_FRMCUST_CC
Set m_cust2 = M_DATA.QUERY_HIST_PAID(M_OBJCONN, "a.custid = '" + lblCustId.Caption + "' ")
listview1(0).ListItems.Clear
While Not m_cust2.EOF
    Set listitem = listview1(0).ListItems.ADD(, , IIf(IsNull(m_cust2("Paydate")), "", m_cust2("Paydate")))
        listitem.SubItems(1) = IIf(IsNull(m_cust2("payment")), "0", Format(m_cust2("Payment"), "##,###"))
        listitem.SubItems(2) = IIf(IsNull(m_cust2("AGENT")), "", m_cust2("AGENT"))
        listitem.SubItems(3) = IIf(IsNull(m_cust2("FieldName")), "", m_cust2("FieldName"))
        listitem.SubItems(4) = IIf(IsNull(m_cust2("Id")), "0", m_cust2("Id"))
        NilaiAfterPay = NilaiAfterPay + IIf(IsNull(m_cust2("payment")), "0", m_cust2("Payment"))
    m_cust2.MoveNext
Wend
Set m_cust2 = Nothing
TxtAfterPay.Value = NilaiAfterPay
txtSisaHutang.Value = Format(TxtPayment2.Value - TxtAfterPay.Value, "##,###")
End Sub
Private Sub Show_Visit()
Dim m_cust2 As New ADODB.Recordset
Dim m_Visit As New ClsVisit
Dim Jml As String
Dim CMDSQL As String
Set m_cust2 = New ADODB.Recordset
CMDSQL = "SELECT requestdate,visitdate,detailsR,detailsV,visitke,VisitNo,id,F_CEK FROM tblvisit where custid='" + lblCustId.Caption + "'"
m_cust2.CursorLocation = adUseClient
m_cust2.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
'Set m_cust2 = m_Visit.SELECT_RequestVisit(M_OBJCONN, lblCustId.Caption)
LstVisit.ListItems.Clear
While Not m_cust2.EOF
    Set listitem = LstVisit.ListItems.ADD(, , IIf(IsNull(m_cust2!RequestDate), "", m_cust2!RequestDate))
        listitem.SubItems(1) = IIf(IsNull(m_cust2!VisitDate), "", m_cust2!VisitDate)
        listitem.SubItems(2) = Trim(IIf(IsNull(m_cust2!VisitNo), "", m_cust2!VisitNo))
        listitem.SubItems(3) = IIf(IsNull(m_cust2!DetailsR), "", m_cust2!DetailsR)
        listitem.SubItems(4) = IIf(IsNull(m_cust2!DetailsV), "", m_cust2!DetailsV)
        listitem.SubItems(5) = IIf(IsNull(m_cust2!VisitKe), "0", m_cust2!VisitKe)
        listitem.SubItems(6) = IIf(IsNull(m_cust2!ID), "0", m_cust2!ID)
        listitem.SubItems(7) = IIf(IsNull(m_cust2!F_CEK), "0", m_cust2!F_CEK)
        m_cust2.MoveNext
Wend
Jml = m_cust2.RecordCount + 1
TDBNumber1.Value = Jml
'Select Case Jml
'Case "0"
'Combo1.Text = "I"
'Case "1"
'Combo1.Text = "II"
'Case "2"
'Combo1.Text = "III"
'Case "3"
'Combo1.Text = "IV"
'Case "4"
'Combo1.Text = "V"
'Case "5"
'Combo1.Text = "VI"
'End Select
Set m_cust2 = Nothing

End Sub
Private Sub CEK_UPDATE_PELANGGAN()

Dim M_DATA As New CLS_FRMCUST_CC_MGM
Dim m_Visit As New ClsVisit
Dim pStatusHstLstCall As String
Dim statusptp As String
'On Error GoTo editErr

'M_OBJCONN.BeginTrans
Set M_update = New ADODB.Recordset
   M_update.CursorLocation = adUseServer
   
   M_update.Open "Select * from mgm where custid='" & lblCustId.Caption & "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
        'ADDITIONAL PHONE
        
        M_update("AHOMENOADD1") = AHomeAdd1(0).Value
        M_update("AHOMENOADD2") = AHomeAdd2(1).Value
        M_update("AOFFICENOADD1") = AOfficeAdd(2).Value
        M_update("AOFFICENOADD2") = AOfficeAdd(3).Value
        M_update("AFAXNOADD1") = AFaxAdd(4).Value
        M_update("AFAXNOADD2") = AFaxAdd(5).Value
        If UCase(Left(MDIForm1.Text2.Text, 5)) = "ADMIN" Then
            M_update("HOMENOADD1") = txtHomeAdd1.Value
            M_update("HOMENOADD2") = txtHomeAdd2.Value
            M_update("OFFICENOADD1") = txtOfficeAdd1.Value
            M_update("OFFICENOADD2") = txtOfficeAdd2.Value
            M_update("MOBILENOADD1") = txtMobileAdd1.Value
            M_update("MOBILENOADD2") = txtMobileAdd2.Value
            M_update("FAXNOADD1") = txtFaxAdd1.Value
            M_update("FAXNOADD2") = txtFaxAdd2.Value
            M_update!TxtPtpAddr = AddrNow.Text
            M_update!ec_name = TxtEC.Text
            M_update!ec_telp = txtECno.Value
        Else
            If txtHomeAdd1A.Value = "" And txtHomeAdd1A.Visible = True Then
                M_update("HOMENOADD1") = txtHomeAdd1A.Value
            ElseIf txtHomeAdd1.Value <> "" And txtHomeAdd1.Visible = True Then
                M_update("HOMENOADD1") = txtHomeAdd1.Value
            End If
            
            If txtHomeAdd2A.Value = "" And txtHomeAdd2A.Visible = True Then
                M_update("HOMENOADD2") = txtHomeAdd2A.Value
            ElseIf txtHomeAdd2.Value <> "" And txtHomeAdd2.Visible = True Then
                M_update("HOMENOADD2") = txtHomeAdd2.Value
            End If
            
            If txtOfficeAdd1A.Value = "" And txtOfficeAdd1A.Visible = True Then
                M_update("OFFICENOADD1") = txtOfficeAdd1A.Value
            ElseIf txtOfficeAdd1.Value <> "" And txtOfficeAdd1.Visible = True Then
                M_update("OFFICENOADD1") = txtOfficeAdd1.Value
            End If
            
            If txtOfficeAdd2A.Value = "" And txtOfficeAdd2A.Visible = True Then
                M_update("OFFICENOADD2") = txtOfficeAdd2A.Value
            ElseIf txtOfficeAdd2.Value <> "" And txtOfficeAdd2.Visible = True Then
                M_update("OFFICENOADD2") = txtOfficeAdd2.Value
            End If
            
            If txtMobileAdd1A.Value = "" And txtMobileAdd1A.Visible = True Then
                M_update("MOBILENOADD1") = txtMobileAdd1A.Value
            ElseIf txtMobileAdd1.Value <> "" And txtMobileAdd1.Visible = True Then
                M_update("MOBILENOADD1") = txtMobileAdd1.Value
            End If
            
            If txtMobileAdd2A.Value = "" And txtMobileAdd2A.Visible = True Then
                M_update("MOBILENOADD2") = txtMobileAdd2A.Value
            ElseIf txtMobileAdd2.Value <> "" And txtMobileAdd2.Visible = True Then
                M_update("MOBILENOADD2") = txtMobileAdd2.Value
            End If
            
        
            M_update("FAXNOADD1") = txtFaxAdd1.Value
            M_update("FAXNOADD2") = txtFaxAdd2.Value
            M_update!TxtPtpAddr = AddrNow.Text
            M_update!ec_name = TxtEC.Text
            M_update!ECAddr = txtECAdd.Text
            M_update!contacwith = Cmbwith.Text
            
                        If txtECnoA.Value = "" And txtECnoA.Visible = True Then
                M_update("ec_telp") = txtECnoA.Value
            ElseIf txtECno.Value <> "" And txtECno.Visible = True Then
                M_update!ec_telp = txtECno.Value
            End If
        End If
        
        If UCase(MDIForm1.Text2.Text) = "AGENT" Then
            If Len(txtECno.Value) > 2 Then
                txtECno.ReadOnly = True
            End If
            If Len(txtHomeAdd1.Value) > 2 Then
                txtHomeAdd1.ReadOnly = True
            End If
            If Len(txtHomeAdd2.Value) > 2 Then
                txtHomeAdd2.ReadOnly = True
            End If
            If Len(txtOfficeAdd1.Value) > 2 Then
                txtOfficeAdd1.ReadOnly = True
            End If
            If Len(txtOfficeAdd2.Value) > 2 Then
                txtOfficeAdd2.ReadOnly = True
            End If
            If Len(txtMobileAdd1.Value) > 2 Then
                txtMobileAdd1.ReadOnly = True
            End If
            If Len(txtMobileAdd2.Value) > 2 Then
                txtMobileAdd2.ReadOnly = True
            End If
        End If
        
'    m_update!f_payment = "PAYMENT"
'    End If
    
     
'        m_update("PRIOR") = cmbPrior.Text
'        m_update("ADDRPT") = lblOfficeAddr.Text
'        m_update("AHOMENO") = AHome1.Value
'        m_update("AHOMENO2") = AHome2.Value
'        m_update("AOFFICENO") = AOffice1.Value
'        m_update("AOFFICENO2") = AOffice2.Value
        M_update("TGLCALL") = Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") & " " & Format(Now, "hh:nn")
'        If Len(IIf(IsNull(m_update!HOMENO), "", m_update!HOMENO)) > 2 Then
'            txtHomeNo1.ReadOnly = True
'        End If
'        m_update("HOMENO2") = txtHomeNo2.Value
'        If Len(IIf(IsNull(m_update!HOMENO2), "", m_update!HOMENO2)) > 2 Then
'            txtHomeNo2.ReadOnly = True
'        End If
'        m_update("MOBILENO") = txtMobileNo1.Value
'        If Len(IIf(IsNull(m_update!MOBILENO), "", m_update!MOBILENO)) > 2 Then
'            txtMobileNo1.ReadOnly = True
'        End If
'        m_update("MOBILENO2") = txtMobileNo2.Value
'        If Len(IIf(IsNull(m_update!MOBILENO2), "", m_update!MOBILENO2)) > 2 Then
'            txtMobileNo2.ReadOnly = True
'        End If
        
'        m_update("OFFICENO") = txtOfficeNo1.Value
'        If Len(IIf(IsNull(m_update!OFFICENO), "", m_update!OFFICENO)) > 2 Then
'            txtOfficeNo1.ReadOnly = True
'        End If
'        m_update("OFFICENO2") = txtOfficeNo2.Value
'        If Len(IIf(IsNull(m_update!OFFICENO2), "", m_update!OFFICENO2)) > 2 Then
'            txtOfficeNo2.ReadOnly = True
            
'         If Len(IIf(IsNull(m_update!HOMENO), "", m_update!HOMENO)) > 2 Then
'            txtHomeNo1.ReadOnly = True
'        End If
'        End If
        'sebelum f_cek diubah statusnya
        statusptp = IIf(IsNull(M_update!F_CEK), "", M_update!F_CEK)
'        If chkAppv(0).Value Then
'            m_update("F_Pending") = "OK"
'        End If


         If C_VALID.Value Then
                M_update("RECSTATUS") = "V"
               pStatusLstCall = cbovalid.Text
               txtResult.Text = pStatusLstCall
               pStatusLstCalldesc = cbodescvalid.Text
               txtResultDesc.Text = pStatusLstCalldesc
                 If Left(cbovalid.Text, 3) = "NBP" Then
                    M_update!F_CEK = "NBP"
                 ElseIf Left(cbovalid.Text, 2) = "NA" Then
                    M_update!F_CEK = Left(cbovalid.Text, 3) & Left(cbodescvalid.Text, 1)
                End If
            Else

        If C_Contacted.Value Then
            M_update("RECSTATUS") = "C"
               pStatusLstCall = cmbContacted.Text
               txtResult.Text = pStatusLstCall
               pStatusLstCalldesc = cmbDescCon.Text
               txtResultDesc.Text = pStatusLstCalldesc
               M_update!F_CEK = Left(cmbContacted.Text, 3) & Left(cmbDescCon.Text, 1)
         Else
                If C_PTP.Value Then
                        pStatusLstCall = cboPTP.Text
                        txtResult.Text = pStatusLstCall
                        'pStatusLstCalldesc = cbodesc.Text
                        txtResultDesc.Text = pStatusLstCalldesc
                        M_update("RECSTATUS") = "P"
                        M_update!F_CEK = Left(cboPTP.Text, 6)
                 Else
                        If C_SKIP.Value Then
                            pStatusLstCall = cboskip.Text
                            txtResult.Text = pStatusLstCall
                            pStatusLstCalldesc = cbodescskip.Text
                            txtResultDesc.Text = pStatusLstCalldesc
                            M_update("RECSTATUS") = "S"
                            M_update!F_CEK = Left(cboskip.Text, 3) & Left(cbodescskip.Text, 2)
                        Else
                                If cboPOPSP.Text <> "" Then
                                    pStatusLstCall = cboPOPSP.Text
                                    txtResult.Text = pStatusLstCall
                                    'pStatusLstCalldesc = cbodescskip.Text
                                    txtResultDesc.Text = pStatusLstCalldesc
                                    M_update("RECSTATUS") = "O"
                                    M_update!F_CEK = Left(cboPOPSP.Text, 3)
                                Else
                                    M_update!F_CEK = ""
                                End If
                          End If
                   End If
                 End If
         End If
        If C_Payment.Value Then
            If statusptp <> Empty Then
                If statusptp = M_update!F_CEK Then
                Else
                    M_update!TGLINCOMING = Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") & " " & Format(Now, "hh:nn")
                End If
            End If
            M_update!ttlptp = txtPayment.Value
            
           ' If txtPayment.ValueIsNull Then
            '    M_update!ttlptp = 0
            'Else
                
             '   If C_PTP.Value = 1 Then
               '     M_update!ttlptp = txtPayment.Value
              '  Else
                '    If vrttlptp <> "" Then
                      '  M_update!ttlptp = vrttlptp
                    'End If
                'End If
            'End If
            
            
            'If Tdabamoint.ValueIsNull Then
             '    M_update!AmountPtp = 0
            'Else
        
             '   If C_PTP.Value = 1 Then
                   M_update!amountptp = Tdabamoint.Value
               ' Else
                '    If vramount <> "" Then
                 '       M_update!AmountPtp = vramount
                  '  End If
                'End If
                
            'End If
            
            'M_update!AmountPtp = Tdabamoint.Value
            'If C_PTP.Value = 1 Then
               M_update!discpersen = cmbDiscount.Text
            'Else
              '  If vrdiskon = "" Then
               ' M_update!discpersen = 0
             '   Else
                
               ' M_update!discpersen = vrdiskon
              '  End If
                
            'End If
            
'            If C_PTP.Value = 1 Then
'                M_update!CmbBaseOn = CmbBaseOn.Text
'            Else
'                    M_update!CmbBaseOn = vrbaseon
'            End If
            
            
            'If txttenor.ValueIsNull Then
            'M_update!tenor = 0
            'Else
            
             'If C_PTP.Value = 1 Then
                   M_update!Tenor = txttenor.Value
              '  Else
               '     If vrtenor <> "" Then
                '        M_update!tenor = vrtenor
                 '   End If
                'End If
           ' End If
            
           ' M_update!tenor = txttenor.Value
           
            
           ' M_update!TdbDatePTP = Format(TdbPTP.Value, "yyyy/mm/dd")
          ' If TDBDate3.ValueIsNull Then
           '    M_update!DatePTP = Null
           'Else
            '    If C_PTP.Value = 1 Then
                    M_update!dateptp = Format(TDBDate3.Value, "yyyy/mm/dd")
             '   Else
              '      If vrdateptp <> "" Then
               '         M_update!DatePTP = vrdateptp
                '    End If
               ' End If
           'End If
            
            'm_update!TxtPtpAddr = TxtPtpAddr.Text
           ' m_update!TxtPhonePTP = TxtPhonePTP.Text
        
        Else
            'm_update!TGLINCOMING = Null
            M_update!ttlptp = 0
            M_update!discpersen = 0
        End If
        
'        If C_lunas.Value Then
'            m_update!TglLunas = Format(TdbLunas.Value, "yyyy/mm/dd")
'            m_update!TotLunas = TDBTot_payment.Value
'            m_update!fieldName = TxtFieldName.Text
'        Else
'            m_update!TglLunas = Null
'            m_update!TotLunas = 0
'            m_update!fieldName = Null
'
'        End If
        
        If Trim(UCase(IIf(IsNull(M_update("KETHSLKERJA")), "", M_update("KETHSLKERJA")))) = Trim(UCase(pStatusLstCall)) Then
            TGLSTATUS = IIf(IsNull(M_update("TGLSTATUS")), "", Format(M_update("TGLSTATUS"), "yyyy/mm/dd"))
        Else
            M_update("KETHSLKERJA") = pStatusLstCall
            M_update("TGLSTATUS") = Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") & " " & Format(Now, "hh:nn")
            TGLSTATUS = Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd")
        End If
        pStatusHstLstCall = IIf(IsNull(M_update("KETHSLKERJA")), "", M_update("KETHSLKERJA"))
        
        M_update("KETHSLKERJADESC") = txtResultDesc.Text
        M_update("PRIOR") = cmbPrior.Text
        M_update("NEXTACT") = cmbNextAct.Text
        M_update("REMARKS") = txtRemarks.Text
        M_update!NEXTACTDATE = Format(cmbDateSch.Value, "yyyy/mm/dd") & " " & Format(Now, "hh:nn")
        M_update("Statuscall") = cbolastcall.Text
    M_update.UPDATE

'M_DATA.UPDATE_CUSTOMER_BARU M_OBJCONN, KETHSLKERJA, STATUS_FIELD_LAMA, M_CALL, M_STATUS, DOK1
If C_NotContacted.Value = 1 Then
    If txtRemarks.Text <> Empty Or cmbNextAct.Text <> Empty Then
        M_DATA.ADD_HISTORY lblCustId.Caption, MDIForm1.TDBDate1.Text, Time, VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11), "COLLECTION", txtRemarks.Text, txtResult.Text, cmbNextAct.Text, "", CStr(IIf(IsNull(M_update!F_CEK), "", M_update!F_CEK)), cbolastcall.Text
    End If
ElseIf C_Contacted.Value = 1 Then
If txtRemarks.Text <> Empty Or cmbNextAct.Text <> Empty Then
       M_DATA.ADD_HISTORY lblCustId.Caption, MDIForm1.TDBDate1.Text, Time, VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11), "COLLECTION", txtRemarks.Text, txtResult.Text, cmbNextAct.Text, "", CStr(IIf(IsNull(M_update!F_CEK), "", M_update!F_CEK)), cbolastcall.Text
End If
ElseIf C_VALID.Value = 1 Then
    If txtRemarks.Text <> Empty Or cmbNextAct.Text <> Empty Then
            M_DATA.ADD_HISTORY lblCustId.Caption, MDIForm1.TDBDate1.Text, Time, VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11), "COLLECTION", txtRemarks.Text, txtResult.Text, cmbNextAct.Text, "", CStr(IIf(IsNull(M_update!F_CEK), "", M_update!F_CEK)), cbolastcall.Text
    End If
ElseIf C_SKIP.Value = 1 Then
    If txtRemarks.Text <> Empty Or cmbNextAct.Text <> Empty Then
            M_DATA.ADD_HISTORY lblCustId.Caption, MDIForm1.TDBDate1.Text, Time, VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11), "COLLECTION", txtRemarks.Text, txtResult.Text, cmbNextAct.Text, "", CStr(IIf(IsNull(M_update!F_CEK), "", M_update!F_CEK)), cbolastcall.Text
    End If
ElseIf C_PTP.Value = 1 Then
    If txtRemarks.Text <> Empty Or cmbNextAct.Text <> Empty Then
            M_DATA.ADD_HISTORY lblCustId.Caption, MDIForm1.TDBDate1.Text, Time, VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11), "COLLECTION", txtRemarks.Text, txtResult.Text, cmbNextAct.Text, "", CStr(IIf(IsNull(M_update!F_CEK), "", M_update!F_CEK)), cbolastcall.Text
    End If
ElseIf cboPOPSP.Text <> "" Then
    If txtRemarks.Text <> Empty Or cmbNextAct.Text <> Empty Then
            M_DATA.ADD_HISTORY lblCustId.Caption, MDIForm1.TDBDate1.Text, Time, VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11), "COLLECTION", txtRemarks.Text, txtResult.Text, cmbNextAct.Text, "", CStr(IIf(IsNull(M_update!F_CEK), "", M_update!F_CEK)), cbolastcall.Text
    End If
End If

    If Len(TDBTot_payment) > 2 Then
    M_DATA.ADD_tbllunas M_OBJCONN, lblCustId.Caption, Format(TdbLunas.Value, "yyyy/mm/dd"), CCur(TDBTot_payment.Value), VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11), TxtFieldName.Text, ""
    Else
    On Error Resume Next
    End If
    '------------>> simpan ke table Visit <<--------------------
   If Option8(0).Value Then
   m_Visit.ADD_RequestVisit M_OBJCONN, lblCustId.Caption, M_update!F_CEK, Text1.Text, TDBDate1.Value, TXtDetails.Text, TDBNumber1.Value, TxtAddress.Text, Trim(UCase(VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11)))
   
   Else
    On Error Resume Next
   End If

'M_OBJCONN.CommitTrans
MsgBox "Data Sudah Tersimpan", vbInformation + vbOKOnly, "Sukses"
kontak = False
Set M_update = Nothing

If shedulePTP_Show = True Then
  '  MDIForm1.LstGrade.ListItems.Remove MDIForm1.LstGrade.SelectedItem.Index
Else
    VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(7) = pStatusLstCall
    VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(8) = txtRemarks.Text
    VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(10) = cbolastcall.Text
'    VIEW_mgmDATA.LstVwSearchmgm.SelectedItem.SubItems(17) = TGLSTATUS
'    VIEW_mgmDATA.LstVwSearchmgm.SelectedItem.SubItems(18) = Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd")
'    VIEW_mgmDATA.LstVwSearchmgm.SelectedItem.SubItems(19) = pStatusHstLstCall
End If
pStatusLstCall = ""
pStatusHstLstCall = ""
txtRemarks.Text = Empty
'cmbNextAct.Text = Empty
'Unload Me
Set M_DATA = Nothing
Exit Sub
'editErr:
'    M_OBJCONN.RollbackTrans
 '   MsgBox Err.Description
  Resume
End Sub
Private Sub HEADER_HISTORY()
    listview1(1).ColumnHeaders.ADD 1, , "Tanggal(mm-dd-yyyy)", 10 * TXT
    listview1(1).ColumnHeaders.ADD 2, , "History", 70 * TXT
    listview1(1).ColumnHeaders.ADD 3, , "Coding", 10 * TXT
    listview1(1).ColumnHeaders.ADD 4, , "Sts Call", 10 * TXT
    listview1(1).ColumnHeaders.ADD 5, , "Sts Call1", 10 * TXT
End Sub
Private Sub HEADER_RequestVisit()
    LstVisit.ColumnHeaders.ADD 1, , "RequestDate", 10 * TXT
    LstVisit.ColumnHeaders.ADD 2, , "VisitDate", 10 * TXT
    LstVisit.ColumnHeaders.ADD 3, , "VisitNo", 10 * TXT
    LstVisit.ColumnHeaders.ADD 4, , "Details", 20 * TXT
    LstVisit.ColumnHeaders.ADD 5, , "Hasil Visit", 20 * TXT
    LstVisit.ColumnHeaders.ADD 6, , "VisitKe", 2 * TXT
    LstVisit.ColumnHeaders.ADD 7, , "ID", 1 * TXT
    LstVisit.ColumnHeaders.ADD 8, , "Status", 1 * TXT
    End Sub
Private Sub HEADER_HISTORY_PAID()
    listview1(0).ColumnHeaders.ADD 1, , "PayDate", 15 * TXT
    listview1(0).ColumnHeaders.ADD 2, , "Payment", 30 * TXT
    listview1(0).ColumnHeaders.ADD 3, , "Agent", 10 * TXT
    listview1(0).ColumnHeaders.ADD 4, , "FieldName", 30 * TXT
    listview1(0).ColumnHeaders.ADD 5, , "Id", 30 * TXT
End Sub
Private Function CEK_DATA_VALID() As Boolean
Dim m_msgbox As Variant
If TDBTot_payment > 2 Then
CEK_DATA_VALID = True
Exit Function
Else

'If MDIForm1.Text2.Text = "TeamLeader" Or MDIForm1.Text2.Text = "Administrator" And (chkAppv(0).Value = 1 Or chkAppv(1).Value = 1) Then
If (chkAppv(0).Value = 1 Or chkAppv(1).Value = 1) Then
        Call UpdateAppv
        'VIEW_mgmDATA.LstVwSearchmgm.SelectedItem.SubItems(8) = VIEW_mgmDATA.LstVwSearchmgm.SelectedItem.SubItems(8) & "Pending"
        Exit Function
'Else
'   CEK_DATA_VALID = False
'End If
Else
    If Left(cmbContacted, 3) = "PTP" And LstPayment.ListItems.Count = 0 Then
            MsgBox "PTP harus buat Nego PTP di tabel yang hijau !!!", vbInformation + vbOKOnly, "TINS"
            CEK_DATA_VALID = False
            Exit Function
    End If
    If cbolastcall.Text = "" Then
            MsgBox "Status Last Call harus diisi", vbInformation + vbOKOnly, "TINS"
            CEK_DATA_VALID = False
            Exit Function
    End If
    If Cmbwith.Text = "" Then
            MsgBox "Contacted With harus diisi", vbInformation + vbOKOnly, "TINS"
            CEK_DATA_VALID = False
            Exit Function
    End If
    
    If Left(cmbContacted.Text, 2) = "RP" Or Left(cmbContacted.Text, 2) = "NA" Then
        If cmbDescCon.Text = "" Then
            CEK_DATA_VALID = False
            MsgBox "Description Contacted Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
            SSTab1.Tab = 3
            Exit Function
        End If
      End If
      If C_Contacted.Value = 1 Then
      If cmbContacted.Text = Empty Then
      CEK_DATA_VALID = False
            MsgBox "Contacted Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
        SSTab1.Tab = 3
        Exit Function
      End If
      End If
     If C_Payment.Value = 1 Then
            If TDBDate3.ValueIsNull Then
             CEK_DATA_VALID = False
             MsgBox "Tanggal PTP Harus Diisi", vbCritical + vbOKOnly, "Peringatan"
            Exit Function
            End If
     End If
            

'      If TdbDatePTP.Text = "__/__/____" Then
'      CEK_DATA_VALID = False
'      MsgBox "Tanggal PTP Harus Diisi", vbCritical + vbOKOnly, "Peringatan"
'      SSTab1.Tab = 3
'      'TdbDatePTP.SetFocus
'      Exit Function
'      End If
      
      
'    If (CmbContacted.Text) = "" And C_NotContacted.Value = 0 Then
'            CEK_DATA_VALID = False
'            MsgBox "Contacted Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
'            SSTab1.Tab = 3
'            Exit Function
'      End If
      
    If Left(cmbUncontacted.Text, 2) <> "" Then
        If cmbDescUn.Text = "" Then
            CEK_DATA_VALID = False
            MsgBox "Description UnContacted Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
            SSTab1.Tab = 3
            Exit Function
       End If
    End If
    
    If cbovalid.Text <> "" Then
        If cbodescvalid.Text = "" Then
            CEK_DATA_VALID = False
            MsgBox "Description Valid Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
            SSTab1.Tab = 3
            Exit Function
        End If
     End If
     
    If cboskip.Text <> "" Then
        If cbodescskip.Text = "" Then
            CEK_DATA_VALID = False
            MsgBox "Description SKIP Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
            SSTab1.Tab = 3
            Exit Function
        End If
     End If
     
     If C_SKIP.Value = 1 Then
     If cboskip.Text = Empty Then
      CEK_DATA_VALID = False
      MsgBox "Description Skip Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
        Exit Function
        SSTab1.Tab = 3
     End If
     End If
     
     If C_VALID.Value = 1 Then
     If cbovalid.Text = Empty Then
      CEK_DATA_VALID = False
      MsgBox "Description Valid Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
        Exit Function
        SSTab1.Tab = 3
     End If
     End If
     
     
     If C_PTP.Value = 1 Then
        If cboPTP.Text = Empty Then
            CEK_DATA_VALID = False
            MsgBox "Description PTP Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
            Exit Function
            SSTab1.Tab = 3
     End If
     End If

     
 
      
     
     
        
         If C_NotContacted.Value = 1 Then
        If cmbUncontacted.Text = Empty Then
            CEK_DATA_VALID = False
            MsgBox "Not Contacted Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
            SSTab1.Tab = 3
            Exit Function
        Else
                  If cmbDescUn.Text = Empty Then
                     MsgBox "Not Contacted Description harus diisi", vbCritical + vbOKOnly, "Peringatan"
                     Exit Function
                  End If
                  If txtRemarks.Text = "" And cmbNextAct.Text = "" Then
                       CEK_DATA_VALID = False
                        MsgBox "Remarks Atau Next Action Harus Diisi...!!!", vbCritical + vbOKOnly, "Peringatan"
                        SSTab1.Tab = 3
                        Exit Function
                  End If
        End If
     End If
     
   If C_Contacted.Value = 0 And C_VALID.Value = 0 And C_PTP.Value = 0 And C_SKIP.Value = 0 And cboPOPSP.Text = "" Then
     CEK_DATA_VALID = False
     MsgBox "Status Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
     SSTab1.Tab = 3
     Exit Function
  End If
 
 
  
    If ADD_CUST = True Then
    Else
    If C_Contacted.Value = 1 Or C_VALID.Value = 1 Or C_PTP.Value = 1 Or C_SKIP.Value = 1 Or cboPOPSP.Text <> "" Then
            If cmbDateSch.ValueIsNull = True Or cmbTimeSch.ValueIsNull = True Then
                CEK_DATA_VALID = False
                MsgBox "Tanggal Schedule Harus Di isi", vbCritical + vbOKOnly, "Peringatan"
                SSTab1.Tab = 3
                Exit Function
            End If
            If txtRemarks.Text = "" And cmbNextAct.Text = "" Then
                CEK_DATA_VALID = False
                MsgBox "Remarks Atau Next Action Harus Diisi...!!!", vbCritical + vbOKOnly, "Peringatan"
                SSTab1.Tab = 3
                Exit Function
            End If
'            If C_Contacted.Value = 1 Then
'                'txtRemarks.Text = cmbContacted & " -" & cmbDescCon & " - " & txtRemarks.Text
'                If cmbDescCon.Text = "" Then
'                    txtRemarks.Text = cmbContacted & " - " & "Contac with " & Cmbwith.Text & " - " & cbolastcall.Text & " - " & txtRemarks.Text
'                Else
'                    txtRemarks.Text = cmbContacted & " - " & "Contac with " & Cmbwith.Text & " - " & cmbDescCon & " - " & cbolastcall.Text & " - " & txtRemarks.Text
'                End If
'            End If

    End If
        If stscall = True Then
            If C_NotContacted.Value = 0 And C_Contacted.Value = 0 And cboPOPSP.Text = "" And C_PTP.Value = 0 And C_SKIP.Value = 0 And C_VALID.Value = 0 Then
                        CEK_DATA_VALID = False
                        MsgBox "Status Harus Diisi", vbCritical + vbOKOnly, "Peringatan"
                        SSTab1.Tab = 3
                        Exit Function
            End If
        End If
'            If C_NotContacted.Value = 1 Then
'                'txtRemarks.Text = cmbUncontacted & " -" & cmbDescUn & " - " & txtRemarks.Text
'                txtRemarks.Text = cmbUncontacted & " - " & cmbDescUn & " - " & cbolastcall.Text & " - " & txtRemarks.Text
'            End If
    End If
    If C_Payment.Value = 1 Then
    
        If CmbBaseOn.Text = "" Then
            MsgBox "Base On harus diisi", vbInformation + vbOKOnly, "TINS"
            CEK_DATA_VALID = False
            Exit Function
        End If
        If cmbDiscount.Text = "" Then
            MsgBox "Diskon harus diisi", vbInformation + vbOKOnly, "TINS"
            CEK_DATA_VALID = False
            Exit Function
        End If
        End If
End If
End If
'cek valid uncontacted pending

If C_Contacted.Value = 1 Then
    'txtRemarks.Text = cmbContacted & " -" & cmbDescCon & " - " & txtRemarks.Text
    If cmbDescCon.Text = "" Then
        txtRemarks.Text = cmbContacted & " - " & "Contac with " & Cmbwith.Text & " - " & cbolastcall.Text & " - " & txtRemarks.Text
    Else
        txtRemarks.Text = cmbContacted & " - " & "Contac with " & Cmbwith.Text & " - " & cmbDescCon & " - " & cbolastcall.Text & " - " & txtRemarks.Text
    End If
End If

If C_NotContacted.Value = 1 Then
    'txtRemarks.Text = cmbUncontacted & " -" & cmbDescUn & " - " & txtRemarks.Text
    txtRemarks.Text = cmbUncontacted & " - " & cmbDescUn & " - " & cbolastcall.Text & " - " & txtRemarks.Text
End If

If C_VALID.Value = 1 Then
                If cbodescvalid.Text = "" Then
                    txtRemarks.Text = cbovalid & " - " & cbolastcall.Text & " - " & txtRemarks.Text
                Else
                    txtRemarks.Text = cbovalid & " - " & cbodescvalid & " - " & cbolastcall.Text & " - " & txtRemarks.Text
                End If
            End If
If C_PTP.Value = 1 Then
        txtRemarks.Text = cboPTP & " - " & cbolastcall.Text & " - " & txtRemarks.Text
End If
If C_SKIP.Value = 1 Then
    If cbodescskip.Text = "" Then
        txtRemarks.Text = cboskip & " - " & cbolastcall.Text & " - " & txtRemarks.Text
    Else
        txtRemarks.Text = cboskip & " - " & cbodescskip & " - " & cbolastcall.Text & " - " & txtRemarks.Text
    End If
End If

If regnego = True Then
    Dim n%
    Dim jum As Currency
    For n = 1 To frmCC_Colection.LstPayment.ListItems.Count
        jum = jum + frmCC_Colection.LstPayment.ListItems(n).SubItems(3)
    Next n
    If jum < frmCC_Colection.txtPayment.Value Then
        MsgBox "Jumlah PTP Belum sama dengan Jumlah Deal Payment"
        CEK_DATA_VALID = False
        txtRemarks.Text = ""
        Exit Function
    End If
End If
regnego = False
CEK_DATA_VALID = True
End Function
Public Sub Custid_Double()
Dim listitem As listitem
Dim test As String
Set m_cust = New ADODB.Recordset
m_cust.CursorLocation = adUseClient
test = Format(LblDOB.Caption, "yyyy/mm/dd")
m_cust.Open "Select a.custid, a.name,a.agent, a.amountwo,a.principal from mgm a where (a.name='" + Trim(TxtName.Text) + "' and dob='" + test + "' or ktpno='" & Trim(lblID.Caption) & "') and a.custid <> '" + Trim(lblCustId.Caption) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
While Not m_cust.EOF
    Set listitem = LstDoubleId.ListItems.ADD(, , IIf(IsNull(m_cust("CUSTID")), "", m_cust("CUSTID")))
        listitem.SubItems(1) = IIf(IsNull(m_cust("NAME")), "", m_cust("NAME"))
        listitem.SubItems(2) = IIf(IsNull(m_cust("AGENT")), "", m_cust("AGENT")) '
        listitem.SubItems(3) = Format(IIf(IsNull(m_cust("AMOUNTWO")), "0", m_cust("AMOUNTWO")), "##,###")
        listitem.SubItems(4) = Format(IIf(IsNull(m_cust("PRINCIPAL")), "0", m_cust("PRINCIPAL")), "##,###")
    m_cust.MoveNext
Wend
Set m_cust = Nothing
End Sub

Private Sub SSCommand2_Click(Index As Integer)
Dim m_msgbox As Variant
Dim STATUS As String
Dim gaji As Currency
Dim gaji1 As String
Dim listitem As listitem
Dim M_DATA As New ClsNegoPTP
Dim JMLPAY As Double
Dim i As Integer
Dim n As Integer
Dim VRDATE As String
Dim jatuhtempo As String
Select Case Index
    Case 0
        If TDBDate3.ValueIsNull Or Tdabamoint.ValueIsNull Or txttenor.ValueIsNull Then
        MsgBox "Pengisian Data Belum Lengkap (installment,tenor,dateptp) "
        Exit Sub
        End If
        jatuhtempo = Format(TDBDate3.Value, "yyyy-mm-dd")
        CMDSQL = "INSERT INTO TblNegoPTP "
            CMDSQL = CMDSQL + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
            CMDSQL = CMDSQL + "VALUES "
            CMDSQL = CMDSQL + "('" + lblCustId + "', "
            CMDSQL = CMDSQL + "'" + jatuhtempo + "', "
            CMDSQL = CMDSQL + "" + CStr(Tdabamoint.Value) + " , "
            CMDSQL = CMDSQL + "'" + CStr(MDIForm1.TDBDate1.Value) + "', "
            CMDSQL = CMDSQL + "'IPO')"
            M_OBJCONN.Execute CMDSQL
            Set listitem = LstPayment.ListItems.ADD(, , "")
            listitem.SubItems(1) = ""
            listitem.SubItems(2) = Format(TDBDate3.Value, "dd/mm/yyyy")
            listitem.SubItems(3) = CStr(Tdabamoint.Value)
            listitem.SubItems(4) = "IPO"
            listitem.SubItems(5) = MDIForm1.TDBDate1.Value

    n = 0
    For i = 1 To Val(txttenor - 1)
            n = n + 1
            JMLPAY = (txtPayment - Tdabamoint) / (txttenor.Value - 1)
            'VRDATE = Format(DateAdd("m", n, TDBDate3.Value), "mm/dd/yyyy")
            VRDATE = DateAdd("m", n, Format(TDBDate3.Value, "yyyy-mm-dd"))
            CMDSQL = "INSERT INTO TblNegoPTP "
            CMDSQL = CMDSQL + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
            CMDSQL = CMDSQL + "VALUES "
            CMDSQL = CMDSQL + "('" + lblCustId + "', "
            CMDSQL = CMDSQL + "'" + CStr(VRDATE) + "', "
            CMDSQL = CMDSQL + "" + CStr(JMLPAY) + " , "
            CMDSQL = CMDSQL + "'" + CStr(MDIForm1.TDBDate1.Value) + "', "
            CMDSQL = CMDSQL + "'IPO')"
            M_OBJCONN.Execute CMDSQL
        Set listitem = LstPayment.ListItems.ADD(, , "")
            listitem.SubItems(1) = ""
                               'listitem.SubItems(2) = .TDBDate1.Value
            listitem.SubItems(2) = Format(VRDATE, "dd/mm/yyyy")
            listitem.SubItems(3) = JMLPAY
            listitem.SubItems(4) = "IPO"
            listitem.SubItems(5) = MDIForm1.TDBDate1.Value
    Next i
   
         '   regnego = True
          '  FrmNegoPTP.Show
            
'        With FrmNegoPTP
'                .Caption = "Tambah Data"
'                .Show vbModal
'                If .ok Then
'                 M_DATA.ADD_NegoPTP M_OBJCONN, .TxtCustid.Text, .TDBDate1.Value, CStr(.TDBNumber1.Value), MDIForm1.TDBDate1.Value, jenis
'                    On Error GoTo add_error
'                    If M_DATA.ADD_OK Then
'                        Set listitem = LstPayment.ListItems.ADD(, , "")
'                            listitem.SubItems(1) = ""
'                            listitem.SubItems(2) = .TDBDate1.Value
'                            listitem.SubItems(3) = .TDBNumber1.Value
'                      On Error GoTo 0
'                    End If
'                End If
'                Unload FrmNegoPTP
'            End With
'        Exit Sub
     
    
    Case 1
         If LstPayment.ListItems.Count = 0 Then
            Exit Sub
        End If
           With FrmNegoPTP
                .Caption = "Ubah Data"

                .TDBDate1.Value = LstPayment.SelectedItem.SubItems(2)
                .TDBNumber1.Value = LstPayment.SelectedItem.SubItems(3)
                .Show vbModal
                If .ok Then

                    M_DATA.UPDATE_NegoPTP M_OBJCONN, .TxtCustid.Text, .TDBDate1.Value, CStr(.TDBNumber1.Value), LstPayment.SelectedItem.SubItems(1)

                    On Error GoTo add_error
                    If M_DATA.ADD_OK Then
                        'LstPayment.SelectedItem.SubItems(1) = ""
                        LstPayment.SelectedItem.SubItems(2) = .TDBDate1.Value
                        LstPayment.SelectedItem.SubItems(3) = .TDBNumber1.Value
                        
                        
                    On Error GoTo 0
                    End If
                End If
                Unload FrmNegoPTP
            End With
        Exit Sub
    Case 2
      
         Frmdelete.Show vbModal
'        If LstPayment.ListItems.Count = 0 Then
'            Exit Sub
'        End If
'        m_msgbox = MsgBox("Yakin Akan Dihapus...!!! ", vbCritical + vbOKCancel, "Peringatan")
'        If m_msgbox = 1 Then
'            M_DATA.DELETE_Nego_PTP M_OBJCONN, LstPayment.SelectedItem.SubItems(1)
'            If M_DATA.ADD_OK Then
'                LstPayment.ListItems.Remove LstPayment.SelectedItem.Index
'            End If
'        End If
'        Exit Sub
    
    
End Select
add_error:
End Sub
Private Sub VisitYES()
Text1.BackColor = &HFF00&
TxtCustid.BackColor = &H80000005
TxtName.BackColor = &H80000005
TDBNumber1.BackColor = &H80000005
TXtDetails.BackColor = &H80000005
'LstVisit.BackColor = &HFF00&
TxtAddress.BackColor = &H80000005
TxtAddress.Enabled = True
TXtDetails.Enabled = True
Option7(0).Enabled = True
Option7(1).Enabled = True
Option7(2).Enabled = True
End Sub
Private Sub VisitNo()
Text1.BackColor = &H8000000F
TxtCustid.BackColor = &H8000000F
TxtName.BackColor = &H8000000F
TDBNumber1.BackColor = &H8000000F
TXtDetails.BackColor = &H8000000F
TxtAddress.BackColor = &H8000000F
'LstVisit.BackColor = &H8000000F
Option8(1).Value = True
Option7(0).Enabled = False
Option7(1).Enabled = False
Option7(2).Enabled = False

TxtAddress.Enabled = False
TXtDetails.Enabled = False
End Sub
Private Sub TdbPTP_Change()
TdbPTP.Value = TDBDate1.Value
End Sub

Private Sub txtECno_Click()
TYPETELP = "Emergency Contact"
txtPhone.Text = txtECno.Value
txtPhoneA.Text = txtECnoA.Value
CmbPhone.Text = "EconPhone"
End Sub


Private Sub txtECnoA_Change()
'txtECno.Text = txtECnoA.Text
End Sub

Private Sub txtECnoA_Click()
TYPETELP = "Emergency Contact"
txtPhone.Text = txtECno.Value
txtPhoneA.Text = txtECnoA.Value
CmbPhone.Text = "EconPhone"
End Sub

Private Sub txtFaxAdd1_KeyDown(KeyCode As Integer, Shift As Integer)
MsgBox "Anda tidak boleh mengisi di fax, kecuali SPV!"
End Sub

Private Sub txtFaxAdd2_KeyDown(KeyCode As Integer, Shift As Integer)
MsgBox "Anda tidak boleh mengisi di fax, kecuali SPV!"
End Sub

Private Sub txtHomeAdd1_Click()
TYPETELP = "HOME1"
    If Trim(AHomeAdd1(0).Value) = "021" Or AHomeAdd1(0).Value = "" Then
        txtPhone.Text = txtHomeAdd1.Value
        txtPhoneA.Text = txtHomeAdd1.Value
    Else
        txtPhone.Text = Trim(AHomeAdd1(0).Value) & Trim(txtHomeAdd1.Value)
        txtPhoneA.Text = Trim(AHomeAdd1(0).Value) & Trim(txtHomeAdd1.Value)
    End If
    CmbPhone.Text = "AddHome1"
End Sub
Private Sub txtHomeAdd1A_Click()
TYPETELP = "HOME1"
    If Trim(AHomeAdd1(0).Value) = "021" Or AHomeAdd1(0).Value = "" Then
        txtPhone.Text = txtHomeAdd1.Value
        txtPhoneA.Text = txtHomeAdd1A.Value
        
    Else
        txtPhone.Text = Trim(AHomeAdd1(0).Value) & Trim(txtHomeAdd1.Value)
        txtPhoneA.Text = Trim(AHomeAdd1(0).Value) & Trim(txtHomeAdd1A.Value)
    End If
    CmbPhone.Text = "AddHome1"
End Sub
Private Sub txtHomeAdd2_Click()
TYPETELP = "HOME2"
If Trim(AHomeAdd2(1).Value) = "021" Or AHomeAdd2(1).Value = "" Then
    txtPhone.Text = txtHomeAdd2.Value
    txtPhoneA.Text = txtHomeAdd2.Value
Else
    txtPhone.Text = Trim(AHomeAdd2(1).Value) & Trim(txtHomeAdd2.Value)
    txtPhoneA.Text = Trim(AHomeAdd2(1).Value) & Trim(txtHomeAdd2.Value)
End If
CmbPhone.Text = "AddHome2"
End Sub
Private Sub txtHomeAdd2A_Change()
'txtHomeAdd2.Text = txtHomeAdd2A.Text
End Sub
Private Sub txtHomeAdd2A_Click()
TYPETELP = "HOME2"
If Trim(AHomeAdd2(1).Value) = "021" Or AHomeAdd2(1).Value = "" Then
    txtPhone.Text = txtHomeAdd2.Value
    txtPhoneA.Text = txtHomeAdd2A.Value
Else
    txtPhone.Text = Trim(AHomeAdd2(1).Value) & Trim(txtHomeAdd2.Value)
    txtPhoneA.Text = Trim(AHomeAdd2(1).Value) & Trim(txtHomeAdd2A.Value)
End If
CmbPhone.Text = "AddHome2"
End Sub

Private Sub txtHomeNo1_Click()
    If Len(txtHomeNo1.Text) > 3 Then
    CmbPhone.Text = "HomePhone"
    Else
    CmbPhone.Text = ""
    End If
End Sub

Private Sub txtHomeNo1A_Click()
If Len(txtHomeNo1A.Text) > 3 Then
    CmbPhone.Text = "HomePhone"
    Else
    CmbPhone.Text = ""
    End If
End Sub

Private Sub txtHomeNo2_Click()
    If Len(txtHomeNo2.Text) > 3 Then
    CmbPhone.Text = "HomePhone2"
    Else
    CmbPhone.Text = ""
    End If
End Sub

Private Sub txtHomeNo2A_Click()
  If Len(txtHomeNo2A.Text) > 3 Then
    CmbPhone.Text = "HomePhone2"
    Else
    CmbPhone.Text = ""
    End If
End Sub

Private Sub txtMobileAdd1A_Click()
TYPETELP = "MOBILE1"
    txtPhone.Text = txtMobileAdd1.Value
    txtPhoneA.Text = txtMobileAdd1A.Value
    
    CmbPhone.Text = "AddMobile1"
End Sub

Private Sub txtMobileAdd2A_Change()
'    txtMobileAdd2.Text = txtMobileAdd2A.Text
End Sub
Private Sub txtMobileAdd2A_Click()
TYPETELP = "MOBILE2"
    txtPhone.Text = txtMobileAdd2.Value
    txtPhoneA.Text = txtMobileAdd2A.Value
    If Len(txtMobileAdd2A.Text) > 3 Then
    CmbPhone.Text = "AddMobile2"
    Else
    CmbPhone.Text = ""
    End If
End Sub
Private Sub txtMobileNo1_Click()
If Len(txtMobileNo1.Text) > 3 Then
CmbPhone.Text = "Hp"
Else
CmbPhone.Text = ""
End If
End Sub

Private Sub txtMobileNo1A_Click()
If Len(txtMobileNo1A.Text) > 3 Then
CmbPhone.Text = "Hp"
Else
CmbPhone.Text = ""
End If
End Sub
Private Sub txtMobileNo2_Click()
If Len(txtMobileNo2.Text) > 3 Then
CmbPhone.Text = "Hp2"
Else
CmbPhone.Text = ""
End If
End Sub
Private Sub txtMobileNo2A_Click()
If Len(txtMobileNo2A.Text) > 3 Then
CmbPhone.Text = "Hp2"
Else
CmbPhone.Text = ""
End If
End Sub

Private Sub txtOfficeAdd1_Click()
TYPETELP = "OFFICE1"
If Trim(AOfficeAdd(2).Value) = "021" Or AOfficeAdd(2).Value = "" Then
    txtPhone.Text = txtOfficeAdd1.Value
    txtPhoneA.Text = txtOfficeAdd1.Value
Else
    txtPhone.Text = Trim(AOfficeAdd(2).Value) & Trim(txtOfficeAdd1.Value)
    txtPhoneA.Text = Trim(AOfficeAdd(2).Value) & Trim(txtOfficeAdd1.Value)
End If
CmbPhone.Text = "AddOffice1"
End Sub

Private Sub txtOfficeAdd1A_Change()
'    txtOfficeAdd1.Text = txtOfficeAdd1A.Text
End Sub

Private Sub txtOfficeAdd1A_Click()
TYPETELP = "OFFICE1"
If Trim(AOfficeAdd(2).Value) = "021" Or AOfficeAdd(2).Value = "" Then
    txtPhone.Text = txtOfficeAdd1.Value
    txtPhoneA.Text = txtOfficeAdd1A.Value
Else
    txtPhone.Text = Trim(AOfficeAdd(2).Value) & Trim(txtOfficeAdd1.Value)
    txtPhoneA.Text = Trim(AOfficeAdd(2).Value) & Trim(txtOfficeAdd1A.Value)
End If
CmbPhone.Text = "AddOffice1"
End Sub

Private Sub txtOfficeAdd2_Click()
TYPETELP = "OFFICE2"
If Trim(AOfficeAdd(3).Value) = "021" Or AOfficeAdd(3).Value = "" Then
    txtPhone.Text = txtOfficeAdd2.Value
    txtPhoneA.Text = txtOfficeAdd2.Value
Else
    txtPhone.Text = Trim(AOfficeAdd(3).Value) & Trim(txtOfficeAdd2.Value)
    txtPhoneA.Text = Trim(AOfficeAdd(3).Value) & Trim(txtOfficeAdd2.Value)
End If
CmbPhone.Text = "AddOffice2"
End Sub

Private Sub txtMobileAdd1_Click()
TYPETELP = "MOBILE1"
    txtPhone.Text = txtMobileAdd1.Value
    txtPhoneA.Text = txtMobileAdd1.Value
If Len(txtMobileAdd1.Text) > 3 Then
    CmbPhone.Text = "AddMobile1"
    Else
    CmbPhone.Text = ""
End If
End Sub

Private Sub txtMobileAdd2_Click()
TYPETELP = "MOBILE2"
    txtPhone.Text = txtMobileAdd2.Value
    txtPhoneA.Text = txtMobileAdd2.Value

If Len(txtMobileAdd2.Text) > 3 Then
    CmbPhone.Text = "AddMobile2"
    Else
    CmbPhone.Text = ""
End If
    
End Sub
Public Sub UpdateAppv()
If chkAppv(0).Value Then
    x = MsgBox("Pindahkan data ke Agent DA ?", vbYesNo + vbExclamation, "Info !")
    If x = vbYes Then
        CMDSQL = "update mgm set F_pending='Pending',Agent='DA',PO_Agent='" & lblaoc.Caption & "' where custid='" + lblCustId.Caption + "'"
        M_OBJCONN.Execute CMDSQL
        spend = True
        MsgBox "Data berhasil dipindah ke agent DA", vbInformation
        VIEW_MGMDATA.LstVwSearchMgm.ListItems.Clear
        MDIForm1.LstGrade.ListItems.Clear
    End If
Else
    If chkAppv(1).Value Then
        Dim spo As ADODB.Recordset
        Set spo = New ADODB.Recordset
        spo.CursorLocation = adUseClient
        spo.Open "select PO_Agent from mgm where custid='" + lblCustId.Caption + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
        
        If spo!PO_AGENT <> "" And IsNull(spo!PO_AGENT) = False Then
            CMDSQL = "update mgm set F_pending='',AGENT=PO_Agent where custid='" + lblCustId.Caption + "'"
            M_OBJCONN.Execute CMDSQL
            CMDSQL = "update mgm set PO_Agent='' where custid='" + lblCustId.Caption + "'"
            M_OBJCONN.Execute CMDSQL
            MsgBox "Data berhasil dikembalikan", vbInformation
            VIEW_MGMDATA.LstVwSearchMgm.ListItems.Clear
            MDIForm1.LstGrade.ListItems.Clear
        Else
            MsgBox "Silahkan Pilih Status !," & vbCrLf & "untuk menyimpan hilangkan ceklist NO !", vbInformation
            Exit Sub
        End If
    End If
End If
End Sub

Private Sub txtOfficeAdd2A_Change()
'    txtOfficeAdd2.Text = txtOfficeAdd2A.Text
End Sub

Private Sub txtOfficeAdd2A_Click()
TYPETELP = "OFFICE2"
If Trim(AOfficeAdd(3).Value) = "021" Or AOfficeAdd(3).Value = "" Then
    txtPhone.Text = txtOfficeAdd2.Value
    txtPhoneA.Text = txtOfficeAdd2A.Value
Else
    txtPhone.Text = Trim(AOfficeAdd(3).Value) & Trim(txtOfficeAdd2.Value)
    txtPhoneA.Text = Trim(AOfficeAdd(3).Value) & Trim(txtOfficeAdd2A.Value)
End If

CmbPhone.Text = "AddOffice2"
End Sub
Private Sub txtOfficeNo1_Click()
If Len(txtOfficeNo1.Text) > 3 Then
CmbPhone.Text = "OfficePhone"
Else
CmbPhone.Text = ""
End If
End Sub
Private Sub txtOfficeNo1A_Click()
If Len(txtOfficeNo1A.Text) > 3 Then
CmbPhone.Text = "OfficePhone"
Else
CmbPhone.Text = ""
End If

End Sub
Private Sub txtOfficeNo2_Click()
If Len(txtOfficeNo2.Text) > 3 Then
CmbPhone.Text = "OfficePhone2"
Else
CmbPhone.Text = ""
End If

End Sub
Private Sub txtOfficeNo2A_Click()
If Len(txtOfficeNo2A.Text) > 3 Then
CmbPhone.Text = "OfficePhone2"
Else
CmbPhone.Text = ""
End If

End Sub
