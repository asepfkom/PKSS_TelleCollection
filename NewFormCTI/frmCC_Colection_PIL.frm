VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCC_ColectionRitpil 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10695
   ClientLeft      =   -1350
   ClientTop       =   210
   ClientWidth     =   18960
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   14  'Copy Pen
   Icon            =   "frmCC_Colection_PIL.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10695
   ScaleWidth      =   18960
   Begin VB.Frame Frame17 
      BackColor       =   &H00B1FDD5&
      BorderStyle     =   0  'None
      Height          =   8895
      Left            =   6720
      TabIndex        =   185
      Top             =   0
      Width           =   12255
      Begin VB.Frame Frame15 
         Appearance      =   0  'Flat
         BackColor       =   &H002F735C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   188
         Top             =   4605
         Width           =   2895
         Begin VB.Image Image1 
            Height          =   375
            Index           =   5
            Left            =   75
            Picture         =   "frmCC_Colection_PIL.frx":000C
            Stretch         =   -1  'True
            Top             =   60
            Width           =   375
         End
         Begin VB.Label Label38 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Additional Info"
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
            Left            =   480
            TabIndex        =   189
            Top             =   0
            Width           =   1575
         End
      End
      Begin VB.Frame Frame15 
         Appearance      =   0  'Flat
         BackColor       =   &H002F735C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   291
         Top             =   7320
         Width           =   2895
         Begin VB.Label Label38 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Emergency Contact"
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
            Height          =   375
            Index           =   3
            Left            =   480
            TabIndex        =   292
            Top             =   0
            Width           =   2055
         End
         Begin VB.Image Image1 
            Height          =   375
            Index           =   2
            Left            =   60
            Picture         =   "frmCC_Colection_PIL.frx":052B
            Stretch         =   -1  'True
            Top             =   60
            Width           =   375
         End
      End
      Begin VB.Frame Frame12 
         Appearance      =   0  'Flat
         BackColor       =   &H00B8E2D4&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1410
         Left            =   5040
         TabIndex        =   280
         Top             =   4920
         Width           =   7080
         Begin VB.CommandButton CmdDeletePelunasan 
            BackColor       =   &H009AD6C2&
            Caption         =   "&Hapus"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6240
            Style           =   1  'Graphical
            TabIndex        =   281
            Top             =   960
            Width           =   705
         End
         Begin TDBNumber6Ctl.TDBNumber txtSisaHutang 
            Height          =   255
            Left            =   5400
            TabIndex        =   282
            Top             =   660
            Width           =   1530
            _Version        =   65536
            _ExtentX        =   2699
            _ExtentY        =   450
            Calculator      =   "frmCC_Colection_PIL.frx":09BC
            Caption         =   "frmCC_Colection_PIL.frx":09DC
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_PIL.frx":0A48
            Keys            =   "frmCC_Colection_PIL.frx":0A66
            Spin            =   "frmCC_Colection_PIL.frx":0AB0
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
            Height          =   255
            Left            =   5400
            TabIndex        =   283
            Top             =   380
            Width           =   1530
            _Version        =   65536
            _ExtentX        =   2699
            _ExtentY        =   450
            Calculator      =   "frmCC_Colection_PIL.frx":0AD8
            Caption         =   "frmCC_Colection_PIL.frx":0AF8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_PIL.frx":0B64
            Keys            =   "frmCC_Colection_PIL.frx":0B82
            Spin            =   "frmCC_Colection_PIL.frx":0BCC
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
            Height          =   255
            Left            =   5400
            TabIndex        =   284
            Top             =   105
            Width           =   1530
            _Version        =   65536
            _ExtentX        =   2699
            _ExtentY        =   450
            Calculator      =   "frmCC_Colection_PIL.frx":0BF4
            Caption         =   "frmCC_Colection_PIL.frx":0C14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_PIL.frx":0C80
            Keys            =   "frmCC_Colection_PIL.frx":0C9E
            Spin            =   "frmCC_Colection_PIL.frx":0CE8
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
            Height          =   1155
            Index           =   0
            Left            =   0
            TabIndex        =   285
            Top             =   0
            Width           =   4065
            _ExtentX        =   7170
            _ExtentY        =   2037
            View            =   3
            LabelEdit       =   1
            SortOrder       =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   0
            BackColor       =   10147522
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
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
            BackColor       =   &H009AD6C2&
            Caption         =   "PTP:"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   4425
            TabIndex        =   288
            Top             =   105
            Width           =   960
         End
         Begin VB.Label Label13 
            BackColor       =   &H009AD6C2&
            Caption         =   "Payment:"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4425
            TabIndex        =   287
            Top             =   380
            Width           =   960
         End
         Begin VB.Label Label15 
            BackColor       =   &H009AD6C2&
            Caption         =   "Sisa Hutang:"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4425
            TabIndex        =   286
            Top             =   660
            Width           =   960
         End
      End
      Begin VB.Frame Frame15 
         Appearance      =   0  'Flat
         BackColor       =   &H002F735C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   7
         Left            =   5040
         TabIndex        =   273
         Top             =   4560
         Width           =   2895
         Begin VB.Label Label38 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Detail Payment"
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
            Height          =   375
            Index           =   7
            Left            =   480
            TabIndex        =   274
            Top             =   0
            Width           =   1695
         End
         Begin VB.Image Image1 
            Height          =   375
            Index           =   4
            Left            =   75
            Picture         =   "frmCC_Colection_PIL.frx":0D10
            Stretch         =   -1  'True
            Top             =   60
            Width           =   375
         End
      End
      Begin VB.Frame Frame15 
         Appearance      =   0  'Flat
         BackColor       =   &H002F735C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   6
         Left            =   5040
         TabIndex        =   271
         Top             =   7080
         Width           =   2895
         Begin VB.Image Image1 
            Height          =   375
            Index           =   6
            Left            =   75
            Picture         =   "frmCC_Colection_PIL.frx":11C1
            Stretch         =   -1  'True
            Top             =   0
            Width           =   375
         End
         Begin VB.Label Label38 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Script"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   6
            Left            =   600
            TabIndex        =   272
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.Frame Frame15 
         Appearance      =   0  'Flat
         BackColor       =   &H002F735C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   186
         Top             =   0
         Width           =   2895
         Begin VB.Image Image1 
            Height          =   375
            Index           =   3
            Left            =   75
            Picture         =   "frmCC_Colection_PIL.frx":1602
            Stretch         =   -1  'True
            Top             =   60
            Width           =   375
         End
         Begin VB.Label Label38 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Call Actvity"
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
            Left            =   480
            TabIndex        =   187
            Top             =   0
            Width           =   1455
         End
      End
      Begin TDBMask6Ctl.TDBMask txtFaxAdd1 
         Height          =   255
         Left            =   1980
         TabIndex        =   241
         Top             =   6120
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   450
         Caption         =   "frmCC_Colection_PIL.frx":1B4A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection_PIL.frx":1BB6
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
         PromptChar      =   " "
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "                                    "
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask txtFaxAdd2 
         Height          =   255
         Left            =   1980
         TabIndex        =   242
         Top             =   6405
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   450
         Caption         =   "frmCC_Colection_PIL.frx":1BF8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection_PIL.frx":1C64
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
         PromptChar      =   " "
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "                  "
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask AFaxAdd 
         Height          =   255
         Index           =   4
         Left            =   1425
         TabIndex        =   243
         Top             =   6120
         Width           =   540
         _Version        =   65536
         _ExtentX        =   952
         _ExtentY        =   450
         Caption         =   "frmCC_Colection_PIL.frx":1CA6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection_PIL.frx":1D12
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   1
         AutoConvert     =   1
         BackColor       =   -2147483643
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
         PromptChar      =   " "
         ReadOnly        =   -1
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "[    ]"
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask AFaxAdd 
         Height          =   255
         Index           =   5
         Left            =   1425
         TabIndex        =   244
         Top             =   6405
         Width           =   540
         _Version        =   65536
         _ExtentX        =   952
         _ExtentY        =   450
         Caption         =   "frmCC_Colection_PIL.frx":1D54
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection_PIL.frx":1DC0
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   1
         AutoConvert     =   1
         BackColor       =   -2147483643
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
         PromptChar      =   " "
         ReadOnly        =   -1
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "[    ]"
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask txtMobileAdd1 
         Height          =   255
         Left            =   1980
         TabIndex        =   245
         Top             =   6690
         Visible         =   0   'False
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   450
         Caption         =   "frmCC_Colection_PIL.frx":1E02
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection_PIL.frx":1E6E
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
         PromptChar      =   " "
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "                  "
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask txtMobileAdd2 
         Height          =   255
         Left            =   1980
         TabIndex        =   246
         Top             =   6975
         Visible         =   0   'False
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   450
         Caption         =   "frmCC_Colection_PIL.frx":1EB0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection_PIL.frx":1F1C
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
         PromptChar      =   " "
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "                  "
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask txtOfficeAdd1 
         Height          =   255
         Left            =   1980
         TabIndex        =   247
         Top             =   5550
         Visible         =   0   'False
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   450
         Caption         =   "frmCC_Colection_PIL.frx":1F5E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection_PIL.frx":1FCA
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
         PromptChar      =   " "
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "                  "
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask txtOfficeAdd2 
         Height          =   255
         Left            =   1980
         TabIndex        =   248
         Top             =   5835
         Visible         =   0   'False
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   450
         Caption         =   "frmCC_Colection_PIL.frx":200C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection_PIL.frx":2078
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
         PromptChar      =   " "
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "                  "
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask AOfficeAdd 
         Height          =   255
         Index           =   2
         Left            =   1425
         TabIndex        =   249
         Top             =   5550
         Width           =   540
         _Version        =   65536
         _ExtentX        =   952
         _ExtentY        =   450
         Caption         =   "frmCC_Colection_PIL.frx":20BA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection_PIL.frx":2126
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   1
         AutoConvert     =   1
         BackColor       =   -2147483643
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
         PromptChar      =   " "
         ReadOnly        =   -1
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "[    ]"
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask AOfficeAdd 
         Height          =   255
         Index           =   3
         Left            =   1425
         TabIndex        =   250
         Top             =   5835
         Width           =   540
         _Version        =   65536
         _ExtentX        =   952
         _ExtentY        =   450
         Caption         =   "frmCC_Colection_PIL.frx":2168
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection_PIL.frx":21D4
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   1
         AutoConvert     =   1
         BackColor       =   -2147483643
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
         PromptChar      =   " "
         ReadOnly        =   -1
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "[    ]"
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask txtHomeAdd1 
         Height          =   255
         Left            =   1980
         TabIndex        =   251
         Top             =   5010
         Visible         =   0   'False
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   450
         Caption         =   "frmCC_Colection_PIL.frx":2216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection_PIL.frx":2282
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
         PromptChar      =   " "
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "                  "
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask txtHomeAdd2 
         Height          =   255
         Left            =   1980
         TabIndex        =   252
         Top             =   5280
         Visible         =   0   'False
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   450
         Caption         =   "frmCC_Colection_PIL.frx":22C4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection_PIL.frx":2330
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
         PromptChar      =   " "
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "                  "
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask AHomeAdd1 
         Height          =   255
         Index           =   0
         Left            =   1425
         TabIndex        =   253
         Top             =   5010
         Width           =   540
         _Version        =   65536
         _ExtentX        =   952
         _ExtentY        =   450
         Caption         =   "frmCC_Colection_PIL.frx":2372
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection_PIL.frx":23DE
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   1
         AutoConvert     =   1
         BackColor       =   -2147483643
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
         PromptChar      =   " "
         ReadOnly        =   -1
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "[    ]"
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask AHomeAdd2 
         Height          =   255
         Index           =   1
         Left            =   1425
         TabIndex        =   254
         Top             =   5280
         Width           =   540
         _Version        =   65536
         _ExtentX        =   952
         _ExtentY        =   450
         Caption         =   "frmCC_Colection_PIL.frx":2420
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection_PIL.frx":248C
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         AllowSpace      =   1
         AutoConvert     =   1
         BackColor       =   -2147483643
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
         PromptChar      =   " "
         ReadOnly        =   -1
         ShowContextMenu =   1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "[    ]"
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask txtHomeAdd1A 
         Height          =   255
         Left            =   1980
         TabIndex        =   255
         Top             =   5010
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   450
         Caption         =   "frmCC_Colection_PIL.frx":24CE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection_PIL.frx":253A
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
         PromptChar      =   " "
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "                  "
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask txtHomeAdd2A 
         Height          =   255
         Left            =   1980
         TabIndex        =   256
         Top             =   5280
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   450
         Caption         =   "frmCC_Colection_PIL.frx":257C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection_PIL.frx":25E8
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
         PromptChar      =   " "
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "                  "
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask txtOfficeAdd1A 
         Height          =   255
         Left            =   1980
         TabIndex        =   257
         Top             =   5550
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   450
         Caption         =   "frmCC_Colection_PIL.frx":262A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection_PIL.frx":2696
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
         PromptChar      =   " "
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "                  "
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask txtOfficeAdd2A 
         Height          =   255
         Left            =   1980
         TabIndex        =   258
         Top             =   5835
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   450
         Caption         =   "frmCC_Colection_PIL.frx":26D8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection_PIL.frx":2744
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
         PromptChar      =   " "
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "                  "
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask txtMobileAdd1A 
         Height          =   255
         Left            =   1980
         TabIndex        =   259
         Top             =   6690
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   450
         Caption         =   "frmCC_Colection_PIL.frx":2786
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection_PIL.frx":27F2
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
         PromptChar      =   " "
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "                  "
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask txtMobileAdd2A 
         Height          =   255
         Left            =   1980
         TabIndex        =   260
         Top             =   6975
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   450
         Caption         =   "frmCC_Colection_PIL.frx":2834
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection_PIL.frx":28A0
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
         PromptChar      =   " "
         ReadOnly        =   -1
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "                  "
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask TxtExt3 
         Height          =   255
         Left            =   3810
         TabIndex        =   261
         Top             =   5550
         Width           =   630
         _Version        =   65536
         _ExtentX        =   1111
         _ExtentY        =   450
         Caption         =   "frmCC_Colection_PIL.frx":28E2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection_PIL.frx":294E
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
         PromptChar      =   " "
         ReadOnly        =   -1
         ShowContextMenu =   0
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "                    "
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask TxtExt4 
         Height          =   255
         Left            =   3810
         TabIndex        =   262
         Top             =   5835
         Width           =   630
         _Version        =   65536
         _ExtentX        =   1111
         _ExtentY        =   450
         Caption         =   "frmCC_Colection_PIL.frx":2990
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection_PIL.frx":29FC
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
         PromptChar      =   " "
         ReadOnly        =   -1
         ShowContextMenu =   0
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "                    "
         Value           =   ""
      End
      Begin RichTextLib.RichTextBox AddrNow 
         Height          =   510
         Left            =   825
         TabIndex        =   293
         Top             =   8280
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   900
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"frmCC_Colection_PIL.frx":2A3E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox TxtEC 
         Height          =   255
         Left            =   825
         TabIndex        =   294
         Top             =   7725
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   450
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"frmCC_Colection_PIL.frx":2ABF
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin TDBMask6Ctl.TDBMask txtECno 
         Height          =   255
         Left            =   825
         TabIndex        =   295
         Top             =   7995
         Visible         =   0   'False
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   450
         Caption         =   "frmCC_Colection_PIL.frx":2B40
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection_PIL.frx":2BAC
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
         PromptChar      =   " "
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "                  "
         Value           =   ""
      End
      Begin TDBMask6Ctl.TDBMask txtECnoA 
         Height          =   255
         Left            =   825
         TabIndex        =   299
         Top             =   7995
         Width           =   1815
         _Version        =   65536
         _ExtentX        =   3201
         _ExtentY        =   450
         Caption         =   "frmCC_Colection_PIL.frx":2BEE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Keys            =   "frmCC_Colection_PIL.frx":2C5A
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
         PromptChar      =   " "
         ReadOnly        =   0
         ShowContextMenu =   -1
         ShowLiterals    =   0
         TabAction       =   0
         Text            =   "                  "
         Value           =   ""
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00B8E2D4&
         BorderStyle     =   0  'None
         Height          =   1260
         Left            =   9840
         TabIndex        =   275
         Top             =   2400
         Width           =   2145
         Begin Threed.SSCommand SSCommand1 
            Height          =   900
            Index           =   2
            Left            =   120
            TabIndex        =   276
            Top             =   120
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   1588
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
            Picture         =   "frmCC_Colection_PIL.frx":2C9C
            AutoSize        =   2
            Alignment       =   4
            ButtonStyle     =   2
            PictureAlignment=   1
         End
         Begin Threed.SSCommand SSCommand1 
            Cancel          =   -1  'True
            Height          =   900
            Index           =   3
            Left            =   1095
            TabIndex        =   277
            Top             =   105
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   1588
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
            Picture         =   "frmCC_Colection_PIL.frx":31CF
            AutoSize        =   2
            Alignment       =   4
            ButtonStyle     =   2
            PictureAlignment=   1
         End
         Begin VB.Label Label42 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H009AD6C2&
            Caption         =   "Save"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   120
            TabIndex        =   279
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label42 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H009AD6C2&
            Caption         =   "Exit"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   1080
            TabIndex        =   278
            Top             =   960
            Width           =   900
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00B8E2D4&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4320
         Left            =   120
         TabIndex        =   190
         Top             =   240
         Width           =   11985
         Begin VB.CheckBox C_PTP 
            BackColor       =   &H00B8E2D4&
            Caption         =   "PTP"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   255
            TabIndex        =   214
            Top             =   1035
            Width           =   750
         End
         Begin VB.CheckBox C_VALID 
            BackColor       =   &H00B8E2D4&
            Caption         =   "Validity??"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   255
            TabIndex        =   213
            Top             =   135
            Width           =   1170
         End
         Begin VB.CheckBox C_POPSP 
            BackColor       =   &H009AD6C2&
            Caption         =   "POP - SP"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5640
            TabIndex        =   201
            Top             =   2025
            Width           =   1155
         End
         Begin VB.CheckBox C_SKIP 
            BackColor       =   &H009AD6C2&
            Caption         =   "Skip"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5640
            TabIndex        =   200
            Top             =   1035
            Width           =   720
         End
         Begin VB.CheckBox C_Contacted 
            BackColor       =   &H009AD6C2&
            Caption         =   "Contacted"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5640
            TabIndex        =   199
            Top             =   90
            Width           =   1215
         End
         Begin VB.ComboBox cbolastcall 
            Height          =   315
            Left            =   6345
            TabIndex        =   198
            Top             =   2700
            Width           =   3015
         End
         Begin VB.ComboBox cmbNextAct 
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
            Left            =   1650
            TabIndex        =   197
            Top             =   4215
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.ComboBox cmbPrior 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            ItemData        =   "frmCC_Colection_PIL.frx":3834
            Left            =   3315
            List            =   "frmCC_Colection_PIL.frx":3841
            Style           =   2  'Dropdown List
            TabIndex        =   196
            Top             =   4215
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00B8E2D4&
            Caption         =   "ReservedPTP"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1485
            Left            =   270
            TabIndex        =   191
            Top             =   2685
            Width           =   4605
            Begin MSComctlLib.ListView LstPayment 
               Height          =   1215
               Left            =   60
               TabIndex        =   192
               Top             =   210
               Width           =   3675
               _ExtentX        =   6482
               _ExtentY        =   2143
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   10147522
               BorderStyle     =   1
               Appearance      =   0
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
               Height          =   615
               Index           =   0
               Left            =   3720
               TabIndex        =   193
               Top             =   240
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   1085
               _Version        =   196610
               Font3D          =   1
               PictureFrames   =   1
               Picture         =   "frmCC_Colection_PIL.frx":3859
               Caption         =   "&Tambah"
               Alignment       =   8
            End
            Begin Threed.SSCommand SSCommand2 
               Height          =   615
               Index           =   2
               Left            =   3765
               TabIndex        =   194
               Top             =   810
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   1085
               _Version        =   196610
               Font3D          =   1
               PictureFrames   =   1
               Picture         =   "frmCC_Colection_PIL.frx":3DE2
               Caption         =   "Hapus"
               Alignment       =   8
            End
            Begin Threed.SSCommand SSCommand2 
               Height          =   615
               Index           =   1
               Left            =   3765
               TabIndex        =   195
               Top             =   210
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   1085
               _Version        =   196610
               Font3D          =   1
               PictureFrames   =   1
               Picture         =   "frmCC_Colection_PIL.frx":4377
               Caption         =   "&Ubah"
               Alignment       =   8
            End
         End
         Begin RichTextLib.RichTextBox txtRemarks 
            Height          =   840
            Left            =   6330
            TabIndex        =   234
            Top             =   3420
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   1482
            _Version        =   393217
            Appearance      =   0
            TextRTF         =   $"frmCC_Colection_PIL.frx":4900
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin TDBDate6Ctl.TDBDate cmbDateSch 
            Height          =   315
            Left            =   6330
            TabIndex        =   235
            Top             =   3030
            Width           =   1890
            _Version        =   65536
            _ExtentX        =   3334
            _ExtentY        =   556
            Calendar        =   "frmCC_Colection_PIL.frx":497C
            Caption         =   "frmCC_Colection_PIL.frx":4A94
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_PIL.frx":4B00
            Keys            =   "frmCC_Colection_PIL.frx":4B1E
            Spin            =   "frmCC_Colection_PIL.frx":4B7C
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
            Left            =   8265
            TabIndex        =   236
            Top             =   3030
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   556
            Caption         =   "frmCC_Colection_PIL.frx":4BA4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_PIL.frx":4C10
            Spin            =   "frmCC_Colection_PIL.frx":4C60
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
         Begin VB.Frame FrmContacted 
            BackColor       =   &H00B8E2D4&
            Height          =   855
            Left            =   5520
            TabIndex        =   202
            Top             =   135
            Width           =   4185
            Begin VB.ComboBox cmbContacted 
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
               ItemData        =   "frmCC_Colection_PIL.frx":4C88
               Left            =   780
               List            =   "frmCC_Colection_PIL.frx":4C8A
               TabIndex        =   204
               Top             =   180
               Width           =   3285
            End
            Begin VB.ComboBox cmbDescCon 
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
               Left            =   780
               TabIndex        =   203
               Top             =   495
               Width           =   3285
            End
            Begin VB.Label label1 
               BackColor       =   &H009AD6C2&
               Caption         =   "Ket:"
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   38
               Left            =   90
               TabIndex        =   206
               Top             =   540
               Width           =   720
            End
            Begin VB.Label label1 
               BackColor       =   &H009AD6C2&
               Caption         =   "Kode:"
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   40
               Left            =   90
               TabIndex        =   205
               Top             =   210
               Width           =   720
            End
         End
         Begin VB.Frame FrMValid 
            BackColor       =   &H00B8E2D4&
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   270
            TabIndex        =   215
            Top             =   180
            Width           =   4530
            Begin VB.ComboBox cbovalid 
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
               ItemData        =   "frmCC_Colection_PIL.frx":4C8C
               Left            =   645
               List            =   "frmCC_Colection_PIL.frx":4C8E
               TabIndex        =   217
               Top             =   195
               Width           =   3765
            End
            Begin VB.ComboBox cbodescvalid 
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
               Left            =   645
               TabIndex        =   216
               Top             =   510
               Width           =   3765
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               BackColor       =   &H009AD6C2&
               Caption         =   "Desc."
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   25
               Left            =   90
               TabIndex        =   219
               Top             =   540
               Width           =   450
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               BackColor       =   &H009AD6C2&
               Caption         =   "Kode"
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   28
               Left            =   90
               TabIndex        =   218
               Top             =   270
               Width           =   420
            End
         End
         Begin VB.Frame frmPTP 
            BackColor       =   &H00B8E2D4&
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   270
            TabIndex        =   220
            Top             =   1065
            Width           =   4560
            Begin VB.ComboBox cboPTP 
               Height          =   315
               Left            =   690
               TabIndex        =   221
               Top             =   150
               Width           =   3735
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               BackColor       =   &H009AD6C2&
               Caption         =   "Kode"
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   37
               Left            =   105
               TabIndex        =   222
               Top             =   210
               Width           =   420
            End
         End
         Begin VB.Frame FrmSKIP 
            BackColor       =   &H00B8E2D4&
            Height          =   885
            Left            =   5535
            TabIndex        =   207
            Top             =   1065
            Width           =   4155
            Begin VB.ComboBox cboskip 
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
               ItemData        =   "frmCC_Colection_PIL.frx":4C90
               Left            =   780
               List            =   "frmCC_Colection_PIL.frx":4C92
               TabIndex        =   209
               Top             =   195
               Width           =   3225
            End
            Begin VB.ComboBox cbodescskip 
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
               Left            =   780
               TabIndex        =   208
               Top             =   510
               Width           =   3225
            End
            Begin VB.Label label1 
               BackColor       =   &H009AD6C2&
               Caption         =   "Kode:"
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   12
               Left            =   80
               TabIndex        =   313
               Top             =   240
               Width           =   720
            End
            Begin VB.Label label1 
               BackColor       =   &H009AD6C2&
               Caption         =   "Ket:"
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   10
               Left            =   80
               TabIndex        =   312
               Top             =   570
               Width           =   720
            End
         End
         Begin VB.Frame FrmPayment 
            BackColor       =   &H00B8E2D4&
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   270
            TabIndex        =   223
            Top             =   1485
            Width           =   4560
            Begin VB.ComboBox CmbBaseOn 
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
               Left            =   915
               TabIndex        =   226
               Top             =   120
               Width           =   1725
            End
            Begin VB.ComboBox cmbDiscount 
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
               Left            =   915
               TabIndex        =   225
               Top             =   435
               Width           =   975
            End
            Begin VB.CheckBox C_Payment 
               BackColor       =   &H009AD6C2&
               Caption         =   "Payment"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3120
               TabIndex        =   224
               Top             =   180
               Width           =   990
            End
            Begin TDBNumber6Ctl.TDBNumber txtPayment 
               Height          =   345
               Left            =   915
               TabIndex        =   227
               Top             =   750
               Width           =   1965
               _Version        =   65536
               _ExtentX        =   3466
               _ExtentY        =   609
               Calculator      =   "frmCC_Colection_PIL.frx":4C94
               Caption         =   "frmCC_Colection_PIL.frx":4CB4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_Colection_PIL.frx":4D20
               Keys            =   "frmCC_Colection_PIL.frx":4D3E
               Spin            =   "frmCC_Colection_PIL.frx":4D88
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
            Begin TDBDate6Ctl.TDBDate TdbPTP 
               Height          =   285
               Left            =   2790
               TabIndex        =   228
               Top             =   435
               Width           =   1350
               _Version        =   65536
               _ExtentX        =   2381
               _ExtentY        =   503
               Calendar        =   "frmCC_Colection_PIL.frx":4DB0
               Caption         =   "frmCC_Colection_PIL.frx":4EC8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_Colection_PIL.frx":4F34
               Keys            =   "frmCC_Colection_PIL.frx":4F52
               Spin            =   "frmCC_Colection_PIL.frx":4FB0
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
            Begin TDBDate6Ctl.TDBDate TdbDatePTP 
               Height          =   285
               Left            =   2940
               TabIndex        =   229
               Top             =   765
               Width           =   1080
               _Version        =   65536
               _ExtentX        =   1905
               _ExtentY        =   503
               Calendar        =   "frmCC_Colection_PIL.frx":4FD8
               Caption         =   "frmCC_Colection_PIL.frx":50F0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_Colection_PIL.frx":515C
               Keys            =   "frmCC_Colection_PIL.frx":517A
               Spin            =   "frmCC_Colection_PIL.frx":51D8
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
            Begin VB.Label label1 
               BackColor       =   &H009AD6C2&
               Caption         =   "Base On:"
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   90
               TabIndex        =   233
               Top             =   135
               Width           =   825
            End
            Begin VB.Label label1 
               BackColor       =   &H009AD6C2&
               Caption         =   "Jumlah:"
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   77
               Left            =   90
               TabIndex        =   232
               Top             =   795
               Width           =   825
            End
            Begin VB.Label label1 
               BackColor       =   &H009AD6C2&
               Caption         =   "Disc."
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   75
               Left            =   90
               TabIndex        =   231
               Top             =   450
               Width           =   825
            End
            Begin VB.Label label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H009AD6C2&
               Caption         =   "Date PTP"
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   2040
               TabIndex        =   230
               Top             =   480
               Width           =   690
            End
         End
         Begin VB.Frame frmpopsp 
            BackColor       =   &H00B8E2D4&
            Height          =   555
            Left            =   5535
            TabIndex        =   210
            Top             =   2040
            Width           =   4185
            Begin VB.ComboBox cboPOPSP 
               Height          =   315
               Left            =   840
               TabIndex        =   211
               Top             =   180
               Width           =   3255
            End
            Begin VB.Label label1 
               BackColor       =   &H009AD6C2&
               Caption         =   "POP-SP"
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   39
               Left            =   90
               TabIndex        =   212
               Top             =   210
               Width           =   765
            End
         End
         Begin VB.Label Label38 
            Caption         =   "Ket. FollowUp:"
            Height          =   255
            Index           =   1
            Left            =   3390
            TabIndex        =   240
            Top             =   4410
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.Label Label39 
            BackColor       =   &H009AD6C2&
            Caption         =   "Tgl FollowUp:"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5460
            TabIndex        =   239
            Top             =   3060
            Width           =   855
         End
         Begin VB.Label Label31 
            BackColor       =   &H009AD6C2&
            Caption         =   "Note:"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Index           =   1
            Left            =   5460
            TabIndex        =   238
            Top             =   3420
            Width           =   855
         End
         Begin VB.Label Label31 
            BackColor       =   &H009AD6C2&
            Caption         =   "Status Telp:"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   5460
            TabIndex        =   237
            Top             =   2745
            Width           =   855
         End
      End
      Begin MSComctlLib.ListView Lstscript 
         Height          =   1335
         Left            =   5040
         TabIndex        =   309
         Top             =   7440
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   2355
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   10147522
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label23 
         BackColor       =   &H009AD6C2&
         Caption         =   "Telp "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   298
         Top             =   7995
         Width           =   720
      End
      Begin VB.Label Label21 
         BackColor       =   &H009AD6C2&
         Caption         =   "Nama"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   297
         Top             =   7725
         Width           =   720
      End
      Begin VB.Label Label19 
         BackColor       =   &H009AD6C2&
         Caption         =   "Add  Addr:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   150
         TabIndex        =   296
         Top             =   8280
         Width           =   690
      End
      Begin VB.Label label1 
         BackColor       =   &H009AD6C2&
         Caption         =   "Fax I"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   59
         Left            =   240
         TabIndex        =   270
         Top             =   6120
         Width           =   1170
      End
      Begin VB.Label label1 
         BackColor       =   &H009AD6C2&
         Caption         =   "Fax II"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   56
         Left            =   240
         TabIndex        =   269
         Top             =   6405
         Width           =   1170
      End
      Begin VB.Label label1 
         BackColor       =   &H009AD6C2&
         Caption         =   "HP II"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   268
         Top             =   6975
         Width           =   1170
      End
      Begin VB.Label label1 
         BackColor       =   &H009AD6C2&
         Caption         =   "HP I"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   267
         Top             =   6690
         Width           =   1170
      End
      Begin VB.Label label1 
         BackColor       =   &H009AD6C2&
         Caption         =   "Kantor II"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   266
         Top             =   5835
         Width           =   1170
      End
      Begin VB.Label label1 
         BackColor       =   &H009AD6C2&
         Caption         =   "Kantor I"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   265
         Top             =   5550
         Width           =   1170
      End
      Begin VB.Label label1 
         BackColor       =   &H009AD6C2&
         Caption         =   "Rumah II"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   264
         Top             =   5280
         Width           =   1170
      End
      Begin VB.Label label1 
         BackColor       =   &H009AD6C2&
         Caption         =   "Rumah I"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   263
         Top             =   5010
         Width           =   1170
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00B1FDD5&
      BorderStyle     =   0  'None
      Height          =   8895
      Index           =   1
      Left            =   0
      TabIndex        =   93
      Top             =   0
      Width           =   6615
      Begin VB.Frame Frame15 
         Appearance      =   0  'Flat
         BackColor       =   &H002F735C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   158
         Top             =   5880
         Width           =   2895
         Begin VB.Label Label38 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Phone Information"
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
            Height          =   375
            Index           =   2
            Left            =   480
            TabIndex        =   159
            Top             =   0
            Width           =   1815
         End
         Begin VB.Image Image1 
            Height          =   375
            Index           =   1
            Left            =   60
            Picture         =   "frmCC_Colection_PIL.frx":5200
            Stretch         =   -1  'True
            Top             =   60
            Width           =   375
         End
      End
      Begin VB.Frame Frame15 
         Appearance      =   0  'Flat
         BackColor       =   &H002F735C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   156
         Top             =   0
         Width           =   2895
         Begin VB.Label Label38 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Personal Data"
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
            Left            =   1320
            TabIndex        =   157
            Top             =   0
            Width           =   1455
         End
         Begin VB.Image Image1 
            Height          =   255
            Index           =   0
            Left            =   75
            Picture         =   "frmCC_Colection_PIL.frx":6A9A
            Stretch         =   -1  'True
            Top             =   60
            Width           =   375
         End
      End
      Begin RichTextLib.RichTextBox LblCHAditionalAddr 
         Height          =   855
         Left            =   855
         TabIndex        =   94
         Top             =   3200
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   1508
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmCC_Colection_PIL.frx":8334
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin TDBDate6Ctl.TDBDate lblDate 
         Height          =   285
         Left            =   2535
         TabIndex        =   95
         Top             =   1140
         Visible         =   0   'False
         Width           =   960
         _Version        =   65536
         _ExtentX        =   1693
         _ExtentY        =   503
         Calendar        =   "frmCC_Colection_PIL.frx":83B5
         Caption         =   "frmCC_Colection_PIL.frx":84CD
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_PIL.frx":8539
         Keys            =   "frmCC_Colection_PIL.frx":8557
         Spin            =   "frmCC_Colection_PIL.frx":85B5
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
         Value           =   3.54031216694028E-316
         CenturyMode     =   0
      End
      Begin RichTextLib.RichTextBox lblOfficeAddr 
         Height          =   645
         Left            =   855
         TabIndex        =   96
         Top             =   2500
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   1138
         _Version        =   393217
         BackColor       =   16777215
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmCC_Colection_PIL.frx":85DD
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin TDBDate6Ctl.TDBDate lblOpenDate 
         Height          =   255
         Left            =   5310
         TabIndex        =   116
         Top             =   1215
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   450
         Calendar        =   "frmCC_Colection_PIL.frx":865E
         Caption         =   "frmCC_Colection_PIL.frx":8776
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_PIL.frx":87E2
         Keys            =   "frmCC_Colection_PIL.frx":8800
         Spin            =   "frmCC_Colection_PIL.frx":885E
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
      Begin TDBDate6Ctl.TDBDate lblLastBill 
         Height          =   255
         Left            =   5310
         TabIndex        =   117
         Top             =   1500
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   450
         Calendar        =   "frmCC_Colection_PIL.frx":8886
         Caption         =   "frmCC_Colection_PIL.frx":899E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_PIL.frx":8A0A
         Keys            =   "frmCC_Colection_PIL.frx":8A28
         Spin            =   "frmCC_Colection_PIL.frx":8A86
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
         Value           =   3.54028845178928E-316
         CenturyMode     =   0
      End
      Begin TDBNumber6Ctl.TDBNumber lblPromPA 
         Height          =   255
         Left            =   5310
         TabIndex        =   118
         Top             =   360
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   450
         Calculator      =   "frmCC_Colection_PIL.frx":8AAE
         Caption         =   "frmCC_Colection_PIL.frx":8ACE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_PIL.frx":8B3A
         Keys            =   "frmCC_Colection_PIL.frx":8B58
         Spin            =   "frmCC_Colection_PIL.frx":8BA2
         AlignHorizontal =   1
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
      Begin TDBNumber6Ctl.TDBNumber LblInstallment 
         Height          =   255
         Left            =   5310
         TabIndex        =   119
         Top             =   2880
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   450
         Calculator      =   "frmCC_Colection_PIL.frx":8BCA
         Caption         =   "frmCC_Colection_PIL.frx":8BEA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_PIL.frx":8C56
         Keys            =   "frmCC_Colection_PIL.frx":8C74
         Spin            =   "frmCC_Colection_PIL.frx":8CBE
         AlignHorizontal =   1
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
      Begin TDBDate6Ctl.TDBDate lblBD 
         Height          =   255
         Left            =   5310
         TabIndex        =   140
         Top             =   3435
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   450
         Calendar        =   "frmCC_Colection_PIL.frx":8CE6
         Caption         =   "frmCC_Colection_PIL.frx":8DFE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_PIL.frx":8E6A
         Keys            =   "frmCC_Colection_PIL.frx":8E88
         Spin            =   "frmCC_Colection_PIL.frx":8EE6
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
         Value           =   1.07202956713409E-317
         CenturyMode     =   0
      End
      Begin TDBNumber6Ctl.TDBNumber lblLimit 
         Height          =   255
         Left            =   5310
         TabIndex        =   141
         Top             =   3705
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   450
         Calculator      =   "frmCC_Colection_PIL.frx":8F0E
         Caption         =   "frmCC_Colection_PIL.frx":8F2E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_PIL.frx":8F9A
         Keys            =   "frmCC_Colection_PIL.frx":8FB8
         Spin            =   "frmCC_Colection_PIL.frx":9002
         AlignHorizontal =   1
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
      Begin TDBNumber6Ctl.TDBNumber lblAmount 
         Height          =   255
         Left            =   5310
         TabIndex        =   142
         Top             =   4800
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   450
         Calculator      =   "frmCC_Colection_PIL.frx":902A
         Caption         =   "frmCC_Colection_PIL.frx":904A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_PIL.frx":90B6
         Keys            =   "frmCC_Colection_PIL.frx":90D4
         Spin            =   "frmCC_Colection_PIL.frx":911E
         AlignHorizontal =   1
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
      Begin TDBNumber6Ctl.TDBNumber lblTtlPay 
         Height          =   255
         Left            =   5310
         TabIndex        =   143
         Top             =   4530
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   450
         Calculator      =   "frmCC_Colection_PIL.frx":9146
         Caption         =   "frmCC_Colection_PIL.frx":9166
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_PIL.frx":91D2
         Keys            =   "frmCC_Colection_PIL.frx":91F0
         Spin            =   "frmCC_Colection_PIL.frx":923A
         AlignHorizontal =   1
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
      Begin TDBNumber6Ctl.TDBNumber lblLastPay 
         Height          =   255
         Left            =   5310
         TabIndex        =   144
         Top             =   4260
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   450
         Calculator      =   "frmCC_Colection_PIL.frx":9262
         Caption         =   "frmCC_Colection_PIL.frx":9282
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_PIL.frx":92EE
         Keys            =   "frmCC_Colection_PIL.frx":930C
         Spin            =   "frmCC_Colection_PIL.frx":9356
         AlignHorizontal =   1
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
      Begin TDBDate6Ctl.TDBDate lblPayDt 
         Height          =   255
         Left            =   5310
         TabIndex        =   145
         Top             =   3975
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   450
         Calendar        =   "frmCC_Colection_PIL.frx":937E
         Caption         =   "frmCC_Colection_PIL.frx":9496
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_PIL.frx":9502
         Keys            =   "frmCC_Colection_PIL.frx":9520
         Spin            =   "frmCC_Colection_PIL.frx":957E
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
      Begin TDBNumber6Ctl.TDBNumber txtAmountwo_A 
         Height          =   255
         Left            =   5310
         TabIndex        =   146
         Top             =   5355
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   450
         Calculator      =   "frmCC_Colection_PIL.frx":95A6
         Caption         =   "frmCC_Colection_PIL.frx":95C6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_PIL.frx":9632
         Keys            =   "frmCC_Colection_PIL.frx":9650
         Spin            =   "frmCC_Colection_PIL.frx":969A
         AlignHorizontal =   1
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
      Begin TDBNumber6Ctl.TDBNumber txtPrinciple_A 
         Height          =   255
         Left            =   5310
         TabIndex        =   147
         Top             =   5085
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   450
         Calculator      =   "frmCC_Colection_PIL.frx":96C2
         Caption         =   "frmCC_Colection_PIL.frx":96E2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_PIL.frx":974E
         Keys            =   "frmCC_Colection_PIL.frx":976C
         Spin            =   "frmCC_Colection_PIL.frx":97B6
         AlignHorizontal =   1
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
      Begin VB.Frame Frame16 
         BackColor       =   &H00B8E2D4&
         BorderStyle     =   0  'None
         Height          =   1935
         Left            =   120
         TabIndex        =   160
         Top             =   6240
         Width           =   6135
         Begin VB.ComboBox CmbPhone 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            ItemData        =   "frmCC_Colection_PIL.frx":97DE
            Left            =   3840
            List            =   "frmCC_Colection_PIL.frx":97E0
            TabIndex        =   178
            Top             =   465
            Width           =   2070
         End
         Begin TDBMask6Ctl.TDBMask AHome2 
            Height          =   255
            Left            =   1245
            TabIndex        =   161
            Top             =   405
            Width           =   405
            _Version        =   65536
            _ExtentX        =   714
            _ExtentY        =   450
            Caption         =   "frmCC_Colection_PIL.frx":97E2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_PIL.frx":984E
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
            Format          =   "[&&&&]"
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
            PromptChar      =   " "
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "[    ]"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask AHome1 
            Height          =   255
            Left            =   1245
            TabIndex        =   162
            Top             =   120
            Width           =   405
            _Version        =   65536
            _ExtentX        =   714
            _ExtentY        =   450
            Caption         =   "frmCC_Colection_PIL.frx":9890
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_PIL.frx":98FC
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
            Format          =   "[&&&&]"
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
            PromptChar      =   " "
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "[    ]"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask AOffice1 
            Height          =   255
            Left            =   1245
            TabIndex        =   163
            Top             =   675
            Width           =   405
            _Version        =   65536
            _ExtentX        =   714
            _ExtentY        =   450
            Caption         =   "frmCC_Colection_PIL.frx":993E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_PIL.frx":99AA
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
            Format          =   "[&&&&]"
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
            PromptChar      =   " "
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "[    ]"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask AOffice2 
            Height          =   255
            Left            =   1245
            TabIndex        =   164
            Top             =   960
            Width           =   405
            _Version        =   65536
            _ExtentX        =   714
            _ExtentY        =   450
            Caption         =   "frmCC_Colection_PIL.frx":99EC
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_PIL.frx":9A58
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
            Format          =   "[&&&&]"
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
            PromptChar      =   " "
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "[    ]"
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtHomeNo1A 
            Height          =   255
            Left            =   1710
            TabIndex        =   165
            Top             =   120
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   450
            Caption         =   "frmCC_Colection_PIL.frx":9A9A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_PIL.frx":9B06
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
            PromptChar      =   " "
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "                    "
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtHomeNo2A 
            Height          =   255
            Left            =   1710
            TabIndex        =   166
            Top             =   405
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   450
            Caption         =   "frmCC_Colection_PIL.frx":9B48
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_PIL.frx":9BB4
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
            PromptChar      =   " "
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "                    "
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtOfficeNo1A 
            Height          =   255
            Left            =   1710
            TabIndex        =   167
            Top             =   675
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   450
            Caption         =   "frmCC_Colection_PIL.frx":9BF6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_PIL.frx":9C62
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
            PromptChar      =   " "
            ReadOnly        =   -1
            ShowContextMenu =   0
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "                    "
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtOfficeNo2A 
            Height          =   255
            Left            =   1710
            TabIndex        =   168
            Top             =   960
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   450
            Caption         =   "frmCC_Colection_PIL.frx":9CA4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_PIL.frx":9D10
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
            PromptChar      =   " "
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "                    "
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtMobileNo1A 
            Height          =   255
            Left            =   1215
            TabIndex        =   169
            Top             =   1245
            Width           =   1515
            _Version        =   65536
            _ExtentX        =   2672
            _ExtentY        =   450
            Caption         =   "frmCC_Colection_PIL.frx":9D52
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_PIL.frx":9DBE
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
            PromptChar      =   " "
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "                    "
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask TxtExt1 
            Height          =   255
            Left            =   3150
            TabIndex        =   170
            Top             =   675
            Width           =   585
            _Version        =   65536
            _ExtentX        =   1032
            _ExtentY        =   450
            Caption         =   "frmCC_Colection_PIL.frx":9E00
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_PIL.frx":9E6C
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
            PromptChar      =   " "
            ReadOnly        =   -1
            ShowContextMenu =   0
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "                    "
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask TxtExt2 
            Height          =   255
            Left            =   3150
            TabIndex        =   171
            Top             =   960
            Width           =   585
            _Version        =   65536
            _ExtentX        =   1032
            _ExtentY        =   450
            Caption         =   "frmCC_Colection_PIL.frx":9EAE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_PIL.frx":9F1A
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
            PromptChar      =   " "
            ReadOnly        =   -1
            ShowContextMenu =   0
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "                    "
            Value           =   ""
         End
         Begin Threed.SSCommand SSCommand1 
            Height          =   780
            Index           =   0
            Left            =   3945
            TabIndex        =   179
            Top             =   885
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   1376
            _Version        =   196610
            Font3D          =   4
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
            Picture         =   "frmCC_Colection_PIL.frx":9F5C
            AutoSize        =   1
            ButtonStyle     =   2
            PictureAlignment=   1
         End
         Begin Threed.SSCommand SSCommand1 
            Height          =   780
            Index           =   1
            Left            =   4905
            TabIndex        =   180
            Top             =   885
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   1376
            _Version        =   196610
            Font3D          =   4
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
            Picture         =   "frmCC_Colection_PIL.frx":A41C
            AutoSize        =   1
            ButtonStyle     =   2
            PictureAlignment=   1
         End
         Begin TDBMask6Ctl.TDBMask txtHomeNo2 
            Height          =   255
            Left            =   1710
            TabIndex        =   314
            Top             =   405
            Visible         =   0   'False
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   450
            Caption         =   "frmCC_Colection_PIL.frx":A938
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_PIL.frx":A9A4
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
            PromptChar      =   " "
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "                    "
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtHomeNo1 
            Height          =   255
            Left            =   1710
            TabIndex        =   315
            Top             =   120
            Visible         =   0   'False
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   450
            Caption         =   "frmCC_Colection_PIL.frx":A9E6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_PIL.frx":AA52
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
            PromptChar      =   " "
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "                    "
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtOfficeNo1 
            Height          =   255
            Left            =   1710
            TabIndex        =   316
            Top             =   675
            Visible         =   0   'False
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   450
            Caption         =   "frmCC_Colection_PIL.frx":AA94
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_PIL.frx":AB00
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
            PromptChar      =   " "
            ReadOnly        =   -1
            ShowContextMenu =   0
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "                    "
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtOfficeNo2 
            Height          =   255
            Left            =   1710
            TabIndex        =   317
            Top             =   960
            Visible         =   0   'False
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2355
            _ExtentY        =   450
            Caption         =   "frmCC_Colection_PIL.frx":AB42
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_PIL.frx":ABAE
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
            PromptChar      =   " "
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "                    "
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtMobileNo1 
            Height          =   255
            Left            =   1215
            TabIndex        =   318
            Top             =   1245
            Visible         =   0   'False
            Width           =   1875
            _Version        =   65536
            _ExtentX        =   3307
            _ExtentY        =   450
            Caption         =   "frmCC_Colection_PIL.frx":ABF0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_PIL.frx":AC5C
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
            PromptChar      =   " "
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "                    "
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtMobileNo2 
            Height          =   255
            Left            =   1215
            TabIndex        =   319
            Top             =   1530
            Visible         =   0   'False
            Width           =   1875
            _Version        =   65536
            _ExtentX        =   3307
            _ExtentY        =   450
            Caption         =   "frmCC_Colection_PIL.frx":AC9E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_PIL.frx":AD0A
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
            PromptChar      =   " "
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "                    "
            Value           =   ""
         End
         Begin TDBMask6Ctl.TDBMask txtMobileNo2A 
            Height          =   255
            Left            =   1215
            TabIndex        =   320
            Top             =   1530
            Width           =   1875
            _Version        =   65536
            _ExtentX        =   3307
            _ExtentY        =   450
            Caption         =   "frmCC_Colection_PIL.frx":AD4C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_Colection_PIL.frx":ADB8
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
            PromptChar      =   " "
            ReadOnly        =   -1
            ShowContextMenu =   -1
            ShowLiterals    =   0
            TabAction       =   0
            Text            =   "                    "
            Value           =   ""
         End
         Begin VB.Label Label42 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H009AD6C2&
            Caption         =   "Call"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   3960
            TabIndex        =   184
            Top             =   1680
            Width           =   765
         End
         Begin VB.Label Label42 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H009AD6C2&
            Caption         =   "Hang Up"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   4920
            TabIndex        =   183
            Top             =   1680
            Width           =   780
         End
         Begin VB.Label lblstatus 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   182
            Top             =   2145
            Width           =   1485
         End
         Begin VB.Label label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H009AD6C2&
            Caption         =   "Telephone Tujuan"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   3870
            TabIndex        =   181
            Top             =   135
            Width           =   1980
         End
         Begin VB.Label label1 
            BackColor       =   &H009AD6C2&
            Caption         =   "HP II"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   55
            Left            =   45
            TabIndex        =   177
            Top             =   1520
            Width           =   1185
         End
         Begin VB.Label label1 
            BackColor       =   &H009AD6C2&
            Caption         =   "HP I"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   52
            Left            =   45
            TabIndex        =   176
            Top             =   1250
            Width           =   1185
         End
         Begin VB.Label label1 
            BackColor       =   &H009AD6C2&
            Caption         =   "Telp Rumah II"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   51
            Left            =   45
            TabIndex        =   175
            Top             =   400
            Width           =   1185
         End
         Begin VB.Label label1 
            BackColor       =   &H009AD6C2&
            Caption         =   "Telp Rumah I"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   47
            Left            =   45
            TabIndex        =   174
            Top             =   120
            Width           =   1185
         End
         Begin VB.Label label1 
            BackColor       =   &H009AD6C2&
            Caption         =   "Telp Kantor I"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   46
            Left            =   45
            TabIndex        =   173
            Top             =   675
            Width           =   1185
         End
         Begin VB.Label label1 
            BackColor       =   &H009AD6C2&
            Caption         =   "Telp Kantor II"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   44
            Left            =   45
            TabIndex        =   172
            Top             =   960
            Width           =   1185
         End
      End
      Begin RichTextLib.RichTextBox lblAddr 
         Height          =   690
         Left            =   855
         TabIndex        =   302
         Top             =   1780
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   1217
         _Version        =   393217
         BackColor       =   16777215
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmCC_Colection_PIL.frx":ADFA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin TDBDate6Ctl.TDBDate lblLcAtm 
         Height          =   255
         Left            =   5310
         TabIndex        =   303
         Top             =   1785
         Width           =   1230
         _Version        =   65536
         _ExtentX        =   2170
         _ExtentY        =   450
         Calendar        =   "frmCC_Colection_PIL.frx":AE7B
         Caption         =   "frmCC_Colection_PIL.frx":AF93
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_Colection_PIL.frx":AFFF
         Keys            =   "frmCC_Colection_PIL.frx":B01D
         Spin            =   "frmCC_Colection_PIL.frx":B07B
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
         Value           =   3.54025880785053E-316
         CenturyMode     =   0
      End
      Begin VB.Label lbltype 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1590
         TabIndex        =   308
         Top             =   5295
         Width           =   1755
      End
      Begin VB.Label LBLEXP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   307
         Top             =   5295
         Width           =   1335
      End
      Begin VB.Label label1 
         BackColor       =   &H009AD6C2&
         Caption         =   "Recsource:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   80
         Left            =   465
         TabIndex        =   306
         Top             =   4920
         Width           =   1065
      End
      Begin VB.Label lblRecsource 
         BackColor       =   &H00FFFFFF&
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1605
         TabIndex        =   305
         Top             =   4935
         Width           =   1740
      End
      Begin VB.Label Label11 
         BackColor       =   &H009AD6C2&
         Caption         =   "Lc atmp"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3975
         TabIndex        =   304
         Top             =   1785
         Width           =   1320
      End
      Begin VB.Label lblCustId 
         BackColor       =   &H00FFFFFF&
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   855
         TabIndex        =   301
         Top             =   360
         Width           =   2835
      End
      Begin VB.Label label1 
         BackColor       =   &H009AD6C2&
         Caption         =   "No Card"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   65
         Left            =   120
         TabIndex        =   300
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label11 
         BackColor       =   &H009AD6C2&
         Caption         =   "Wo A.P"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   3960
         TabIndex        =   155
         Top             =   5355
         Width           =   1320
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label11 
         BackColor       =   &H009AD6C2&
         Caption         =   "Princ A. P:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   3960
         TabIndex        =   154
         Top             =   5085
         Width           =   1320
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label11 
         BackColor       =   &H009AD6C2&
         Caption         =   "WO Date:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   153
         Top             =   3435
         Width           =   1320
      End
      Begin VB.Label Label11 
         BackColor       =   &H009AD6C2&
         Caption         =   "Limit:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   152
         Top             =   3705
         Width           =   1320
      End
      Begin VB.Label Label11 
         BackColor       =   &H009AD6C2&
         Caption         =   "Amount wo:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   3960
         TabIndex        =   151
         Top             =   4800
         Width           =   1320
      End
      Begin VB.Label Label11 
         BackColor       =   &H009AD6C2&
         Caption         =   "Ttl Pay:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   3960
         TabIndex        =   150
         Top             =   4530
         Width           =   1320
      End
      Begin VB.Label Label11 
         BackColor       =   &H009AD6C2&
         Caption         =   "Last Pay:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   3960
         TabIndex        =   149
         Top             =   4260
         Width           =   1320
      End
      Begin VB.Label Label11 
         BackColor       =   &H009AD6C2&
         Caption         =   "LPD :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3960
         TabIndex        =   148
         Top             =   3975
         Width           =   1320
      End
      Begin VB.Label Label34 
         BackColor       =   &H009AD6C2&
         Caption         =   "Range"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3975
         TabIndex        =   139
         Top             =   3150
         Width           =   1320
      End
      Begin VB.Label lblrange 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5310
         TabIndex        =   138
         Top             =   3150
         Width           =   1230
      End
      Begin VB.Label Label33 
         BackColor       =   &H009AD6C2&
         Caption         =   "Last Installment Paid:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3975
         TabIndex        =   137
         Top             =   2595
         Width           =   1320
      End
      Begin VB.Label lblLIP 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5310
         TabIndex        =   136
         Top             =   2595
         Width           =   1230
      End
      Begin VB.Label lblNoCard 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "-------------------"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5310
         TabIndex        =   135
         Top             =   0
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H009AD6C2&
         Caption         =   "#Card"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3975
         TabIndex        =   134
         Top             =   0
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label Label14 
         BackColor       =   &H009AD6C2&
         Caption         =   "#Pay"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3975
         TabIndex        =   133
         Top             =   165
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label lblNoPay 
         BackColor       =   &H00FFFFFF&
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
         Left            =   5310
         TabIndex        =   132
         Top             =   165
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label Label16 
         BackColor       =   &H009AD6C2&
         Caption         =   "Principle"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3975
         TabIndex        =   131
         Top             =   360
         Width           =   1320
      End
      Begin VB.Label Label18 
         BackColor       =   &H009AD6C2&
         Caption         =   "Open Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3975
         TabIndex        =   130
         Top             =   1215
         Width           =   1320
      End
      Begin VB.Label Label20 
         BackColor       =   &H009AD6C2&
         Caption         =   "Last Bill"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3975
         TabIndex        =   129
         Top             =   1500
         Width           =   1320
      End
      Begin VB.Label Label25 
         BackColor       =   &H009AD6C2&
         Caption         =   "B P"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3975
         TabIndex        =   128
         Top             =   2055
         Width           =   1320
      End
      Begin VB.Label lblBrokenPromised 
         BackColor       =   &H00FFFFFF&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5310
         TabIndex        =   127
         Top             =   2055
         Width           =   1230
      End
      Begin VB.Label Label3 
         BackColor       =   &H009AD6C2&
         Caption         =   "Interest"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3975
         TabIndex        =   126
         Top             =   645
         Width           =   1320
      End
      Begin VB.Label Label4 
         BackColor       =   &H009AD6C2&
         Caption         =   "Fees"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3975
         TabIndex        =   125
         Top             =   930
         Width           =   1320
      End
      Begin VB.Label LblInterest 
         BackColor       =   &H00FFFFFF&
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5310
         TabIndex        =   124
         Top             =   645
         Width           =   1230
      End
      Begin VB.Label LblFees 
         BackColor       =   &H00FFFFFF&
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5310
         TabIndex        =   123
         Top             =   930
         Width           =   1230
      End
      Begin VB.Label Label36 
         BackColor       =   &H009AD6C2&
         Caption         =   "Tgl BD-OP"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3975
         TabIndex        =   122
         Top             =   2325
         Width           =   1320
      End
      Begin VB.Label lbllama 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5310
         TabIndex        =   121
         Top             =   2325
         Width           =   1230
      End
      Begin VB.Label Label33 
         BackColor       =   &H009AD6C2&
         Caption         =   "Installment "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3975
         TabIndex        =   120
         Top             =   2880
         Width           =   1320
      End
      Begin VB.Label Label32 
         BackColor       =   &H009AD6C2&
         Caption         =   "Last Month End Dlg"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1620
         TabIndex        =   115
         Top             =   4095
         Width           =   1755
      End
      Begin VB.Label lblLMED 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3405
         TabIndex        =   114
         Top             =   4095
         Width           =   195
      End
      Begin VB.Label CustId 
         BackColor       =   &H009AD6C2&
         Caption         =   "No Kartu:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   113
         Top             =   4440
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblCardNo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
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
         Left            =   855
         TabIndex        =   112
         Top             =   4440
         Visible         =   0   'False
         Width           =   2835
      End
      Begin VB.Label Label2 
         BackColor       =   &H009AD6C2&
         Caption         =   "Nama:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   111
         Top             =   640
         Width           =   720
      End
      Begin VB.Label lblNama 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   855
         TabIndex        =   110
         Top             =   645
         Width           =   2820
      End
      Begin VB.Label Label5 
         BackColor       =   &H009AD6C2&
         Caption         =   "KTP:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   109
         Top             =   930
         Width           =   720
      End
      Begin VB.Label lblID 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   855
         TabIndex        =   108
         Top             =   930
         Width           =   2820
      End
      Begin VB.Label Label6 
         BackColor       =   &H009AD6C2&
         Caption         =   "Tgl Lahir:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   107
         Top             =   1220
         Width           =   720
      End
      Begin VB.Label Label8 
         BackColor       =   &H009AD6C2&
         Caption         =   "Alamat Rumah:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         TabIndex        =   106
         Top             =   1780
         Width           =   720
      End
      Begin VB.Label Label27 
         BackColor       =   &H009AD6C2&
         Caption         =   "Alamat Kantor:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   120
         TabIndex        =   105
         Top             =   2500
         Width           =   720
      End
      Begin VB.Label lblZIP 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   855
         TabIndex        =   104
         Top             =   4095
         Width           =   660
      End
      Begin VB.Label Label22 
         BackColor       =   &H009AD6C2&
         Caption         =   "Zip:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   103
         Top             =   4100
         Width           =   720
      End
      Begin VB.Label LblDOB 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   855
         TabIndex        =   102
         Top             =   1215
         Width           =   1620
      End
      Begin VB.Label Label35 
         BackColor       =   &H009AD6C2&
         Caption         =   "Priority:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   101
         Top             =   1500
         Width           =   720
      End
      Begin VB.Label LblRiskLevel 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2970
         TabIndex        =   100
         Top             =   1500
         Width           =   780
      End
      Begin VB.Label CustId 
         AutoSize        =   -1  'True
         BackColor       =   &H009AD6C2&
         Caption         =   "Risk Level"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   2085
         TabIndex        =   99
         Top             =   1500
         Width           =   810
      End
      Begin VB.Label lblPriority 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   855
         TabIndex        =   98
         Top             =   1500
         Width           =   1140
      End
      Begin VB.Label Label37 
         BackColor       =   &H009AD6C2&
         Caption         =   "Alamat Add :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   97
         Top             =   3200
         Width           =   720
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5850
      Left            =   2520
      TabIndex        =   1
      Top             =   11160
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   10319
      _Version        =   393216
      Tabs            =   6
      Tab             =   5
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   16711680
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
      TabPicture(0)   =   "frmCC_Colection_PIL.frx":B0A3
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame9"
      Tab(0).Control(1)=   "Frame6"
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(4)=   "Option5"
      Tab(0).Control(5)=   "Option6"
      Tab(0).Control(6)=   "Option2"
      Tab(0).Control(7)=   "Option1"
      Tab(0).Control(8)=   "Option4"
      Tab(0).Control(9)=   "Option3"
      Tab(0).Control(10)=   "LstDoubleId"
      Tab(0).Control(11)=   "label1(9)"
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Additional Fields"
      TabPicture(1)   =   "frmCC_Colection_PIL.frx":B0BF
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "History"
      TabPicture(2)   =   "frmCC_Colection_PIL.frx":B0DB
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Results"
      TabPicture(3)   =   "frmCC_Colection_PIL.frx":B0F7
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FrmUnContacted"
      Tab(3).Control(1)=   "txtResult"
      Tab(3).Control(2)=   "txtResultDesc"
      Tab(3).Control(3)=   "txtDiscount"
      Tab(3).Control(4)=   "FrmLunas"
      Tab(3).Control(5)=   "C_NotContacted"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Detail Payment"
      TabPicture(4)   =   "frmCC_Colection_PIL.frx":B113
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Request Visit"
      TabPicture(5)   =   "frmCC_Colection_PIL.frx":B12F
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "LstVisit"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Frame888"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).ControlCount=   2
      Begin VB.Frame Frame888 
         ForeColor       =   &H000000FF&
         Height          =   3255
         Left            =   390
         TabIndex        =   65
         Top             =   900
         Width           =   9855
      End
      Begin VB.Frame Frame4 
         Caption         =   "Emergency Contact"
         Height          =   1395
         Left            =   -67950
         TabIndex        =   63
         Top             =   1260
         Width           =   4215
      End
      Begin VB.Frame Frame9 
         Height          =   615
         Left            =   -74880
         TabIndex        =   62
         Top             =   3810
         Width           =   6375
      End
      Begin VB.Frame Frame6 
         Height          =   1215
         Left            =   -67050
         TabIndex        =   58
         Top             =   2760
         Width           =   2775
      End
      Begin VB.CheckBox C_NotContacted 
         BackColor       =   &H00C5974B&
         Height          =   270
         Left            =   -70320
         TabIndex        =   56
         Top             =   6120
         Width           =   375
      End
      Begin VB.Frame FrmLunas 
         Height          =   1695
         Left            =   -74880
         TabIndex        =   45
         Top             =   5880
         Visible         =   0   'False
         Width           =   4335
         Begin RichTextLib.RichTextBox TxtFieldName 
            Height          =   375
            Left            =   1560
            TabIndex        =   52
            Top             =   1200
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393217
            TextRTF         =   $"frmCC_Colection_PIL.frx":B14B
         End
         Begin TDBNumber6Ctl.TDBNumber TDBTot_payment 
            Height          =   375
            Left            =   1560
            TabIndex        =   51
            Top             =   720
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   661
            Calculator      =   "frmCC_Colection_PIL.frx":B1CD
            Caption         =   "frmCC_Colection_PIL.frx":B1ED
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_PIL.frx":B259
            Keys            =   "frmCC_Colection_PIL.frx":B277
            Spin            =   "frmCC_Colection_PIL.frx":B2C1
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
            Left            =   0
            TabIndex        =   46
            Top             =   0
            Width           =   1455
         End
         Begin TDBDate6Ctl.TDBDate TdbLunas 
            Height          =   285
            Left            =   1560
            TabIndex        =   47
            Top             =   360
            Width           =   1350
            _Version        =   65536
            _ExtentX        =   2381
            _ExtentY        =   503
            Calendar        =   "frmCC_Colection_PIL.frx":B2E9
            Caption         =   "frmCC_Colection_PIL.frx":B401
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_PIL.frx":B46D
            Keys            =   "frmCC_Colection_PIL.frx":B48B
            Spin            =   "frmCC_Colection_PIL.frx":B4E9
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
            Left            =   1560
            TabIndex        =   54
            Top             =   0
            Width           =   4215
         End
         Begin VB.Label Label17 
            Caption         =   "Label17"
            Height          =   375
            Left            =   1320
            TabIndex        =   53
            Top             =   0
            Width           =   135
         End
         Begin VB.Label Label9 
            Caption         =   "Field Name"
            Height          =   255
            Left            =   240
            TabIndex        =   50
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Total Payment"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   49
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Date of Payment"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   48
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2300
         Left            =   -67020
         TabIndex        =   37
         Top             =   480
         Width           =   2745
      End
      Begin VB.Frame Frame2 
         Height          =   3375
         Left            =   -70020
         TabIndex        =   36
         Top             =   450
         Width           =   2955
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00E8BE91&
         Height          =   195
         Left            =   -64740
         TabIndex        =   34
         Top             =   5160
         Width           =   225
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00E8BE91&
         Height          =   195
         Left            =   -64740
         TabIndex        =   32
         Top             =   4785
         Width           =   210
      End
      Begin VB.TextBox txtDiscount 
         Height          =   285
         Left            =   -74760
         TabIndex        =   8
         Top             =   5880
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtResultDesc 
         Height          =   285
         Left            =   -74880
         TabIndex        =   7
         Top             =   6120
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtResult 
         Height          =   285
         Left            =   -74880
         TabIndex        =   6
         Top             =   6480
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E8BE91&
         Height          =   195
         Left            =   -71205
         TabIndex        =   5
         Top             =   5100
         Width           =   225
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E8BE91&
         Height          =   195
         Left            =   -71205
         TabIndex        =   4
         Top             =   4755
         Width           =   240
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E8BE91&
         Height          =   195
         Left            =   -67860
         TabIndex        =   3
         Top             =   4785
         Width           =   210
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E8BE91&
         Height          =   195
         Left            =   -67875
         TabIndex        =   2
         Top             =   5145
         Width           =   255
      End
      Begin MSComctlLib.ListView listview1 
         Height          =   5400
         Index           =   3
         Left            =   -74850
         TabIndex        =   9
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
      Begin MSComctlLib.ListView LstDoubleId 
         Height          =   975
         Left            =   -74910
         TabIndex        =   43
         Top             =   5760
         Width           =   10245
         _ExtentX        =   18071
         _ExtentY        =   1720
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         NumItems        =   0
      End
      Begin VB.Frame FrmUnContacted 
         Height          =   1095
         Left            =   -70320
         TabIndex        =   38
         Top             =   6120
         Width           =   4620
         Begin VB.CheckBox chkAppv 
            BackColor       =   &H00C5974B&
            Caption         =   "NO"
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   1
            Left            =   3600
            TabIndex        =   57
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
            TabIndex        =   55
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
            ItemData        =   "frmCC_Colection_PIL.frx":B511
            Left            =   1250
            List            =   "frmCC_Colection_PIL.frx":B513
            TabIndex        =   40
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
            ItemData        =   "frmCC_Colection_PIL.frx":B515
            Left            =   1245
            List            =   "frmCC_Colection_PIL.frx":B517
            TabIndex        =   39
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
            Left            =   360
            TabIndex        =   44
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
            TabIndex        =   42
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
            TabIndex        =   41
            Top             =   720
            Width           =   960
         End
      End
      Begin MSComctlLib.ListView LstVisit 
         Height          =   1215
         Left            =   375
         TabIndex        =   59
         Top             =   4335
         Width           =   9885
         _ExtentX        =   17436
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
         Index           =   9
         Left            =   -74760
         TabIndex        =   60
         Top             =   4440
         Width           =   1890
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
         TabIndex        =   35
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
         TabIndex        =   33
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   3735
         Width           =   1890
      End
   End
   Begin VB.Frame Frame10 
      Height          =   495
      Left            =   7140
      TabIndex        =   64
      Top             =   5910
      Width           =   450
      Begin VB.OptionButton Option8 
         Caption         =   "Batal"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   6735
         TabIndex        =   67
         Top             =   5265
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Tambah"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   5685
         TabIndex        =   66
         Top             =   5265
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Frame Frame8 
         ForeColor       =   &H000000FF&
         Height          =   570
         Left            =   5685
         TabIndex        =   68
         Top             =   5175
         Visible         =   0   'False
         Width           =   960
         Begin VB.OptionButton Option7 
            Caption         =   "Kantor"
            Height          =   195
            Index           =   2
            Left            =   6525
            TabIndex        =   74
            Top             =   840
            Width           =   840
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Rumah"
            Height          =   195
            Index           =   1
            Left            =   5565
            TabIndex        =   73
            Top             =   855
            Width           =   840
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Alamat Billing"
            Height          =   195
            Index           =   0
            Left            =   4125
            TabIndex        =   72
            Top             =   855
            Width           =   1440
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
            TabIndex        =   71
            Top             =   225
            Width           =   1815
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
            TabIndex        =   70
            Top             =   3375
            Width           =   1935
         End
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
            TabIndex        =   69
            Top             =   540
            Width           =   3135
         End
         Begin TDBNumber6Ctl.TDBNumber TDBNumber1 
            Height          =   315
            Left            =   915
            TabIndex        =   75
            Top             =   870
            Width           =   615
            _Version        =   65536
            _ExtentX        =   1085
            _ExtentY        =   556
            Calculator      =   "frmCC_Colection_PIL.frx":B519
            Caption         =   "frmCC_Colection_PIL.frx":B539
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_PIL.frx":B5A5
            Keys            =   "frmCC_Colection_PIL.frx":B5C3
            Spin            =   "frmCC_Colection_PIL.frx":B60D
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
            TabIndex        =   76
            Top             =   225
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   1005
            _Version        =   393217
            BackColor       =   16777215
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmCC_Colection_PIL.frx":B635
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
            TabIndex        =   77
            Top             =   1200
            Visible         =   0   'False
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   556
            Calendar        =   "frmCC_Colection_PIL.frx":B6BA
            Caption         =   "frmCC_Colection_PIL.frx":B7D2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_PIL.frx":B83E
            Keys            =   "frmCC_Colection_PIL.frx":B85C
            Spin            =   "frmCC_Colection_PIL.frx":B8BA
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
            TabIndex        =   78
            Top             =   870
            Visible         =   0   'False
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   556
            Calendar        =   "frmCC_Colection_PIL.frx":B8E2
            Caption         =   "frmCC_Colection_PIL.frx":B9FA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_Colection_PIL.frx":BA66
            Keys            =   "frmCC_Colection_PIL.frx":BA84
            Spin            =   "frmCC_Colection_PIL.frx":BAE2
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
            TabIndex        =   79
            Top             =   1065
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   953
            _Version        =   393217
            BackColor       =   16777215
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmCC_Colection_PIL.frx":BB0A
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
            TabIndex        =   86
            Top             =   915
            Width           =   615
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
            TabIndex        =   85
            Top             =   3375
            Width           =   1095
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
            TabIndex        =   84
            Top             =   540
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
            TabIndex        =   83
            Top             =   930
            Width           =   810
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
            TabIndex        =   82
            Top             =   1245
            Visible         =   0   'False
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
            Left            =   3390
            TabIndex        =   81
            Top             =   180
            Width           =   660
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
            TabIndex        =   80
            Top             =   240
            Width           =   810
         End
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   525
      Left            =   15735
      TabIndex        =   87
      Top             =   1005
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   926
      _Version        =   196610
      Font3D          =   1
      ForeColor       =   16711935
      BackColor       =   11664853
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      ShadowStyle     =   1
      Begin VB.Frame Frame1 
         Height          =   4230
         Left            =   1365
         TabIndex        =   88
         Top             =   90
         Width           =   180
         Begin VB.Frame Frame14 
            Height          =   3855
            Index           =   1
            Left            =   7800
            TabIndex        =   92
            Top             =   0
            Width           =   15
         End
         Begin VB.Label lblregion 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   135
            Left            =   2655
            TabIndex        =   91
            Top             =   225
            Width           =   1215
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
            TabIndex        =   90
            Top             =   450
            Width           =   60
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
            Index           =   60
            Left            =   15510
            TabIndex        =   89
            Top             =   120
            Width           =   1500
         End
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00B1FDD5&
      BorderStyle     =   0  'None
      Height          =   1800
      Left            =   0
      TabIndex        =   289
      Top             =   9000
      Width           =   18960
      Begin VB.Frame Frame15 
         Appearance      =   0  'Flat
         BackColor       =   &H002F735C&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   310
         Top             =   0
         Width           =   2895
         Begin VB.Label Label38 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "History"
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
            Height          =   375
            Index           =   8
            Left            =   480
            TabIndex        =   311
            Top             =   0
            Width           =   1335
         End
         Begin VB.Image Image1 
            Height          =   375
            Index           =   7
            Left            =   60
            Picture         =   "frmCC_Colection_PIL.frx":BB8F
            Stretch         =   -1  'True
            Top             =   60
            Width           =   375
         End
      End
      Begin MSComctlLib.ListView listview1 
         Height          =   1290
         Index           =   1
         Left            =   135
         TabIndex        =   290
         Top             =   330
         Width           =   18645
         _ExtentX        =   32888
         _ExtentY        =   2275
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   10147522
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
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
   Begin VB.TextBox txtPhoneA 
      Height          =   285
      Left            =   780
      Locked          =   -1  'True
      TabIndex        =   61
      Top             =   10050
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.TextBox txtPhone 
      Height          =   285
      Left            =   1650
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   10065
      Visible         =   0   'False
      Width           =   2640
   End
End
Attribute VB_Name = "frmCC_ColectionRitpil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim m_cust As ADODB.Recordset
'Dim M_update As ADODB.Recordset
'Dim M_OBJRS As ADODB.Recordset
'Dim stscall As Boolean
'Dim TYPETELP As String
'Dim kontak As Boolean
'Dim spend As Boolean
'Dim M_Col As ADODB.Recordset
'Dim qt As String
'Dim statusptp2 As String
'
'Private Sub C_Contacted_Click()
''   If C_Contacted.Value Then
''        C_NotContacted.Value = False
''        C_Payment.Value = False
''        FrmContacted.Enabled = True
''   Else
''        cmbContacted.Text = ""
''        cmbDescCon.Text = ""
''        CmbBaseOn.Text = ""
''        cmbDiscount.Text = 0
''        TdbPTP.Value = ""
''        txtPayment.Value = 0
'''        TxtPtpAddr.Text = ""
'' '       TxtPhonePTP.Text = ""
''        FrmContacted.Enabled = False
''        FrmPayment.Enabled = False
''   End If
''   If C_Contacted.Value = False Then
''   C_Payment.Value = False
''   End If
'
'If C_Contacted.Value Then
'        C_VALID.Value = False
'        C_SKIP.Value = False
'        C_Payment.Value = False
'        C_PTP.Value = False
'        C_POPSP.Value = False
'        FrmContacted.Enabled = True
'   Else
'        cmbContacted.Text = ""
'        cmbDescCon.Text = ""
'        FrmContacted.Enabled = False
'End If
'End Sub
'
'Private Sub C_NotContacted_Click()
'   If C_NotContacted.Value Then
'      FrmUnContacted.Enabled = True
'      C_Contacted.Value = False
'      C_Payment.Value = False
'   Else
'      FrmUnContacted.Enabled = False
'      cmbDescUn.Text = ""
'      cmbUncontacted = ""
'   End If
'End Sub
'
'Private Sub C_Payment_Click()
'   If C_Payment.Value Then
'     ' Frame54.Enabled = True
'   Else
'     ' Frame54.Enabled = False
'      cmbDiscount.Text = ""
'   End If
'End Sub
'
'
'
'
'
'Private Sub C_POPSP_Click()
'If C_POPSP.Value Then
'        C_VALID.Value = False
'        C_SKIP.Value = False
'        'C_Payment.Value = 1
'        C_PTP.Value = False
'        C_Contacted.Value = False
'        frmpopsp.Enabled = True
'        FrmPayment.Enabled = True
'   Else
'        cboPOPSP.Text = ""
'        frmpopsp.Enabled = False
'        CmbBaseOn.Text = ""
'        cmbDiscount.Text = 0
'        TdbPTP.Value = ""
'        txtPayment.Value = 0
''        TxtPtpAddr.Text = ""
' '       TxtPhonePTP.Text = ""
'        FrmPayment.Enabled = False
'        'C_Payment = False
'End If
'End Sub
'
'Private Sub C_PTP_Click()
'If C_PTP.Value Then
'        C_VALID.Value = False
'        C_SKIP.Value = False
'        C_Payment.Value = 1
'        C_POPSP.Value = False
'        C_Contacted.Value = False
'        frmPTP.Enabled = True
'        FrmPayment.Enabled = True
'
'
'
'   Else
'        CmbBaseOn.Text = ""
'        cmbDiscount.Text = 0
'        TdbPTP.Value = ""
'        txtPayment.Value = 0
''        TxtPtpAddr.Text = ""
' '       TxtPhonePTP.Text = ""
'        FrmPayment.Enabled = False
'        cboPTP.Text = ""
'        frmPTP.Enabled = False
'        'C_Payment = False
'End If
'End Sub
'
'Private Sub C_SKIP_Click()
'If C_SKIP.Value Then
'        C_VALID.Value = False
'        C_Contacted.Value = False
'        C_Payment.Value = False
'        C_PTP.Value = False
'        C_POPSP.Value = False
'        FrmSKIP.Enabled = True
'   Else
'        cboskip.Text = ""
'        cbodescskip.Text = ""
'        FrmSKIP.Enabled = False
'End If
'End Sub
'
'Private Sub C_VALID_Click()
'If C_VALID.Value Then
'        C_Contacted.Value = False
'        C_SKIP.Value = False
'        C_Payment.Value = False
'        C_PTP.Value = False
'        C_POPSP.Value = False
'        FrMValid.Enabled = True
'   Else
'        cbovalid.Text = ""
'        cbodescvalid.Text = ""
'        FrMValid.Enabled = False
'End If
'End Sub
'
'Private Sub cbodescskip_KeyDown(KeyCode As Integer, Shift As Integer)
''MsgBox "Jangan di ketik, tapi di pilih, Bisa ga sih kamu...  !!!"
'cbodescskip.Text = ""
'Exit Sub
'End Sub
'
'Private Sub cbodescvalid_KeyDown(KeyCode As Integer, Shift As Integer)
''MsgBox "Jangan di ketik, tapi di pilih, Bisa ga sih kamu...  !!!"
'cbodescvalid.Text = ""
'Exit Sub
'End Sub
'
'Private Sub cbolastcall_GotFocus()
'cbolastcall.CLEAR
'Dim M_OBJRS As ADODB.Recordset
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from ContactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'While Not M_OBJRS.EOF
'    cbolastcall.AddItem M_OBJRS("KdNoProdPresented")
'    M_OBJRS.MoveNext
'Wend
'Set M_OBJRS = Nothing
'
''Set m_objrs = New ADODB.Recordset
''m_objrs.CursorLocation = adUseClient
''m_objrs.Open "Select * from UnContactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
''While Not m_objrs.EOF
''    cbolastcall.AddItem m_objrs("KdNoProdPresented")
''    m_objrs.MoveNext
''Wend
''Set m_objrs = Nothing
'
''LIST VALID
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from tblvalid", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    While Not M_OBJRS.EOF
'        cbolastcall.AddItem M_OBJRS!KdNoProdPresented
'        M_OBJRS.MoveNext
'    Wend
'Set M_OBJRS = Nothing
''LIST Contacted
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from ContactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    While Not M_OBJRS.EOF
'        cbolastcall.AddItem M_OBJRS!KdNoProdPresented
'        M_OBJRS.MoveNext
'    Wend
'Set M_OBJRS = Nothing
''LIST PTP
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from tblPTP", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    While Not M_OBJRS.EOF
'        cbolastcall.AddItem M_OBJRS!KdNoProdPresented
'        M_OBJRS.MoveNext
'    Wend
' Set M_OBJRS = Nothing
''LIST SKIP
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from tblSKIP", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    While Not M_OBJRS.EOF
'        cbolastcall.AddItem M_OBJRS!KdNoProdPresented
'        M_OBJRS.MoveNext
'    Wend
'Set M_OBJRS = Nothing
''LIST POPSP
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from POPSPDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    While Not M_OBJRS.EOF
'        cbolastcall.AddItem M_OBJRS!KdNoProdPresented
'        M_OBJRS.MoveNext
'    Wend
'Set M_OBJRS = Nothing
'End Sub
'
'Private Sub cbolastcall_KeyDown(KeyCode As Integer, Shift As Integer)
''MsgBox "Jangan di ketik, tapi di pilih, Bisa ga sih kamu...  !!!"
'cbolastcall.Text = ""
'Exit Sub
'End Sub
'
'Private Sub cboPOPSP_Click()
'If Left(cmbContacted.Text, 2) = "PO" Or Left(cmbContacted.Text, 2) = "SP" Then
'            cmbDescCon.Enabled = False
'            C_Payment.Value = 1
'            FrmPayment.Enabled = True
'            Set M_COL1 = New ADODB.Recordset
'            CMDSQL = "SELECT tglincoming,tdbdatePTP,ttlPTP,discpersen,cmbbaseon from mgm where custid ='" & VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(1) & "'"
'            M_COL1.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
''            CmbBaseOn.Text = "PRINCIPLE"
'            txtPayment.Value = CStr(IIf(IsNull(M_COL1!ttlptp), "", M_COL1!ttlptp))
'            CmbBaseOn.Text = CStr(IIf(IsNull(M_COL1!CmbBaseOn), "", M_COL1!CmbBaseOn))
'            TdbPTP.Value = CStr(IIf(IsNull(M_COL1!TdbDatePTP), "", M_COL1!TdbDatePTP))
'            cmbDiscount.Text = CStr(IIf(IsNull(M_COL1!discpersen), "", M_COL1!discpersen))
'      'Set m_cust = Nothing
'    End If
'End Sub
'
'Private Sub cboPOPSP_KeyDown(KeyCode As Integer, Shift As Integer)
''MsgBox "Jangan di ketik, tapi di pilih, Bisa ga sih kamu...  !!!"
'cboPOPSP.Text = ""
'Exit Sub
'End Sub
'
'Private Sub cboPTP_Click()
'  If qt = "O" Then
'        If mdiform1.txtlevel.text = "Agent" And (cboPTP.Text = "PTP-NEW" Or cboPTP.Text = "PTP-PAIDOFF") Then
'            cboPTP.Text = ""
'            End If
'        End If
'
'    If mdiform1.txtlevel.text = "Agent" And Left(statusptp2, 3) = "OP-" And (cboPTP.Text = "PTP-POP") Then
'    cboPTP.Text = ""
'    End If
'
'If Left(cmbContacted.Text, 2) = "PT" Then
'    C_Payment.Value = 1
'    FrmPayment.Enabled = True
'    CmbBaseOn.Text = "PRINCIPLE"
'End If
'End Sub
'Private Sub cboPTP_KeyDown(KeyCode As Integer, Shift As Integer)
''MsgBox "Jangan di ketik, tapi di pilih, Bisa ga sih kamu...  !!!"
'cboPTP.Text = ""
'Exit Sub
'End Sub
'
'Private Sub cboskip_Click()
'cbodescskip.CLEAR
'If Left(cboskip.Text, 2) <> "MV" Then
'   Set M_OBJRS = New ADODB.Recordset
'   M_OBJRS.CursorLocation = adUseClient
'   M_OBJRS.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
'         For i = 0 To 3
'           cbodescskip.AddItem M_OBJRS("Description")
'           M_OBJRS.MoveNext
'         Next i
'   Set M_OBJRS = Nothing
'   C_Payment.Value = 0
'Else
'   Set M_OBJRS = New ADODB.Recordset
'   M_OBJRS.CursorLocation = adUseClient
'      M_OBJRS.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
'       While Not M_OBJRS.EOF
'           cbodescskip.AddItem M_OBJRS("Description")
'           M_OBJRS.MoveNext
'       Wend
'   Set M_OBJRS = Nothing
'   C_Payment.Value = 0
'End If
'End Sub
'
'Private Sub cboskip_KeyDown(KeyCode As Integer, Shift As Integer)
''MsgBox "Jangan di ketik, tapi di pilih, Bisa ga sih kamu...  !!!"
'cboskip.Text = ""
'Exit Sub
'End Sub
'
'Private Sub cbovalid_Click()
'Dim i As Integer
'cbodescvalid.CLEAR
'If Left(cbovalid.Text, 2) = "NA" Then
'        cbodescvalid.Enabled = True
''        CmbBaseOn.Text = ""
''        txtPayment.Text = 0
''        cmbDiscount.Text = ""
''        TdbPTP.Text = ""
''        TdbDatePTP.Text = ""
'        Set M_OBJRS = New ADODB.Recordset
'        M_OBJRS.CursorLocation = adUseClient
'          M_OBJRS.Open "Select * from DescContacted WHERE id >= 13 ", M_OBJCONN, adOpenDynamic, adLockOptimistic
'        While Not M_OBJRS.EOF
'            cbodescvalid.AddItem M_OBJRS("Description")
'            M_OBJRS.MoveNext
'        Wend
'        C_Payment.Value = 0
'        Set M_OBJRS = Nothing
''        FrmPayment.Enabled = False
'Else
'        Set M_OBJRS = New ADODB.Recordset
'        M_OBJRS.CursorLocation = adUseClient
'          M_OBJRS.Open "Select * from DescunContacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
'        While Not M_OBJRS.EOF
'            cbodescvalid.AddItem M_OBJRS("Description")
'            M_OBJRS.MoveNext
'        Wend
'        C_Payment.Value = 0
'        Set M_OBJRS = Nothing
'End If
'End Sub
'
'Private Sub Check1_Click()
'If C_Contacted.Value Then
'        C_VALID.Value = False
'        C_SKIP.Value = False
'        C_Payment.Value = False
'        C_PTP.Value = False
'        FrmContacted.Enabled = True
'   Else
'        cmbContacted.Text = ""
'        cmbDescCon.Text = ""
'        FrmContacted.Enabled = False
'End If
'End Sub
'
'Private Sub cbovalid_KeyDown(KeyCode As Integer, Shift As Integer)
''MsgBox "Jangan di ketik, tapi di pilih, Bisa ga sih kamu...  !!!"
'cbovalid.Text = ""
'Exit Sub
'End Sub
'
'Private Sub chkAppv_Click(Index As Integer)
'Select Case Index
'Case 0:
'    chkAppv(1).Value = 0
'Case 1:
'    chkAppv(0).Value = 0
'End Select
'End Sub
'
'Private Sub CmbBaseOn_Click()
'If CmbBaseOn.Text = "PRINCIPLE" Then CmbBaseOn.Text = ""
'    Call cmbDiscount_Click
'End Sub
'
'Private Sub CmbBaseOn_LostFocus()
'    Call cmbDiscount_Click
'End Sub
'
'Private Sub cmbContacted_Click()
''DESCRIPTION CONTACTED
'Dim i As Integer
'Dim M_COL1 As ADODB.Recordset
'cmbDescCon.CLEAR
'If Left(cmbContacted.Text, 2) = "RP" Then
'    cmbDescCon.Enabled = True
'    CmbBaseOn.Text = ""
'    txtPayment.Text = 0
'    cmbDiscount.Text = ""
'    TdbPTP.Text = ""
'    TdbDatePTP.Text = ""
'   Set M_OBJRS = New ADODB.Recordset
'   M_OBJRS.CursorLocation = adUseClient
'     M_OBJRS.Open "Select * from DescContacted where id <= 12", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    While Not M_OBJRS.EOF
'        cmbDescCon.AddItem M_OBJRS("Description")
'        M_OBJRS.MoveNext
'    Wend
'    Set M_OBJRS = Nothing
'    C_Payment.Value = 0
'    FrmPayment.Enabled = False
'    Else
''    If Left(cmbContacted.Text, 2) = "NA" Then
''        cmbDescCon.Enabled = True
''        CmbBaseOn.Text = ""
''        txtPayment.Text = 0
''        cmbDiscount.Text = ""
''        TdbPTP.Text = ""
''        TdbDatePTP.Text = ""
''        Set M_OBJRS = New ADODB.Recordset
''          M_OBJRS.Open "Select * from DescContacted WHERE id >= 13 ", M_OBJCONN, adOpenDynamic, adLockOptimistic
''        While Not M_OBJRS.EOF
''            cmbDescCon.AddItem M_OBJRS("Description")
''            M_OBJRS.MoveNext
''        Wend
''        C_Payment.Value = 0
''        FrmPayment.Enabled = False
'
''    Else
'         If Left(cmbContacted.Text, 2) = "PT" Then
'            cmbDescCon.Enabled = False
'            C_Payment.Value = 1
'            FrmPayment.Enabled = True
'            CmbBaseOn.Text = "PRINCIPLE"
'    Else
'        If Left(cmbContacted.Text, 2) = "BP" Then
'            cmbDescCon.Enabled = False
'            CmbBaseOn.Text = ""
'            txtPayment.Text = 0
'            cmbDiscount.Text = ""
'            TdbPTP.Text = ""
'            TdbDatePTP.Text = ""
'            C_Payment.Value = 0
'            FrmPayment.Enabled = False
'    Else
'    If Left(cmbContacted.Text, 2) = "OP" Then
'            cmbDescCon.Enabled = False
'            CmbBaseOn.Text = ""
'            txtPayment.Text = 0
'            cmbDiscount.Text = ""
'            TdbPTP.Text = ""
'            TdbDatePTP.Text = ""
'            C_Payment.Value = 0
'            FrmPayment.Enabled = False
'      Else
'
'    If Left(cmbContacted.Text, 2) = "PO" Or Left(cmbContacted.Text, 2) = "SP" Then
'            cmbDescCon.Enabled = False
'            C_Payment.Value = 1
'            FrmPayment.Enabled = True
'Set M_COL1 = New ADODB.Recordset
'CMDSQL = "SELECT tglincoming,tdbdatePTP,ttlPTP,discpersen,cmbbaseon from mgm where custid ='" & VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(1) & "'"
'    M_COL1.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
''            CmbBaseOn.Text = "PRINCIPLE"
'            txtPayment.Value = CStr(IIf(IsNull(M_COL1!ttlptp), "", M_COL1!ttlptp))
'            CmbBaseOn.Text = CStr(IIf(IsNull(M_COL1!CmbBaseOn), "", M_COL1!CmbBaseOn))
'            TdbPTP.Value = CStr(IIf(IsNull(M_COL1!TdbDatePTP), "", M_COL1!TdbDatePTP))
'            cmbDiscount.Text = CStr(IIf(IsNull(M_COL1!discpersen), "", M_COL1!discpersen))
'
'      'Set m_cust = Nothing
'    End If
'End If
'End If
'End If
'End If
''End If
'
'Set M_OBJRS = Nothing
'End Sub
'
'Private Sub cmbContacted_KeyDown(KeyCode As Integer, Shift As Integer)
''MsgBox "Jangan di ketik, tapi di pilih, Bisa ga sih kamu...  !!!"
'cmbContacted.Text = ""
'Exit Sub
'End Sub
'
'Private Sub cmbDescCon_GotFocus()
''DESCRIPTION CONTACTED
'Dim i As Integer
'cmbDescCon.CLEAR
'If Left(cmbContacted.Text, 2) = "RP" Then
'    cmbDescCon.Enabled = True
'    CmbBaseOn.Text = ""
'    txtPayment.Text = 0
'    cmbDiscount.Text = ""
'    TdbPTP.Text = ""
'    TdbDatePTP.Text = ""
'   Set M_OBJRS = New ADODB.Recordset
'   M_OBJRS.CursorLocation = adUseClient
'     M_OBJRS.Open "Select * from DescContacted where id <= 12", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    While Not M_OBJRS.EOF
'        cmbDescCon.AddItem M_OBJRS("Description")
'        M_OBJRS.MoveNext
'    Wend
'    C_Payment.Value = 0
'    FrmPayment.Enabled = False
'    Set M_OBJRS = Nothing
'    Else
''    If Left(cmbContacted.Text, 2) = "NA" Then
''        cmbDescCon.Enabled = True
''        CmbBaseOn.Text = ""
''        txtPayment.Text = 0
''        cmbDiscount.Text = ""
''        TdbPTP.Text = ""
''        TdbDatePTP.Text = ""
''        Set M_OBJRS = New ADODB.Recordset
''          M_OBJRS.Open "Select * from DescContacted WHERE id >= 13 ", M_OBJCONN, adOpenDynamic, adLockOptimistic
''        While Not M_OBJRS.EOF
''            cmbDescCon.AddItem M_OBJRS("Description")
''            M_OBJRS.MoveNext
''        Wend
''        C_Payment.Value = 0
''        FrmPayment.Enabled = False
'
''    Else
'         If Left(cmbContacted.Text, 2) = "PT" Then
'            cmbDescCon.Enabled = False
'            C_Payment.Value = 1
'            FrmPayment.Enabled = True
'            CmbBaseOn.Text = "PRINCIPLE"
'    Else
'        If Left(cmbContacted.Text, 2) = "BP" Then
'            cmbDescCon.Enabled = False
'            CmbBaseOn.Text = ""
'            txtPayment.Text = 0
'            cmbDiscount.Text = ""
'            TdbPTP.Text = ""
'            TdbDatePTP.Text = ""
'            C_Payment.Value = 0
'            FrmPayment.Enabled = False
'    Else
'    If Left(cmbContacted.Text, 2) = "OP" Then
'            cmbDescCon.Enabled = False
'            CmbBaseOn.Text = ""
'            txtPayment.Text = 0
'            cmbDiscount.Text = ""
'            TdbPTP.Text = ""
'            TdbDatePTP.Text = ""
'            C_Payment.Value = 0
'            FrmPayment.Enabled = False
'      Else
'
'    If Left(cmbContacted.Text, 2) = "PO" Or Left(cmbContacted.Text, 2) = "SP" Then
'            cmbDescCon.Enabled = False
'            C_Payment.Value = 1
'            FrmPayment.Enabled = True
'Set m_cust = New ADODB.Recordset
'
'CMDSQL = "SELECT tglincoming,tdbdatePTP,ttlPTP,discpersen,cmbbaseon from mgm where custid ='" & VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(1) & "'"
'    m_cust.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
''            CmbBaseOn.Text = "PRINCIPLE"
'            txtPayment.Value = CStr(IIf(IsNull(m_cust!ttlptp), "", m_cust!ttlptp))
'            CmbBaseOn.Text = CStr(IIf(IsNull(m_cust!CmbBaseOn), "", m_cust!CmbBaseOn))
'            TdbPTP.Value = CStr(IIf(IsNull(m_cust!TdbDatePTP), "", m_cust!TdbDatePTP))
'            cmbDiscount.Text = CStr(IIf(IsNull(m_cust!discpersen), "", m_cust!discpersen))
'
'      Set m_cust = Nothing
'    End If
'End If
'End If
'End If
'End If
''End If
'
'Set M_OBJRS = Nothing
'End Sub
'
'Private Sub cmbDescCon_KeyDown(KeyCode As Integer, Shift As Integer)
''MsgBox "Jangan di ketik, tapi di pilih, Bisa ga sih kamu...  !!!"
'cmbDescCon.Text = ""
'Exit Sub
'End Sub
'
'Private Sub cmbDescUn_GotFocus()
'Dim i As Integer
'cmbDescUn.CLEAR
'If Left(cmbUncontacted.Text, 2) = "NA" Then
'        cmbDescUn.Enabled = True
''        CmbBaseOn.Text = ""
''        txtPayment.Text = 0
''        cmbDiscount.Text = ""
''        TdbPTP.Text = ""
''        TdbDatePTP.Text = ""
'        Set M_OBJRS = New ADODB.Recordset
'        M_OBJRS.CursorLocation = adUseClient
'          M_OBJRS.Open "Select * from DescContacted WHERE id >= 13 ", M_OBJCONN, adOpenDynamic, adLockOptimistic
'        While Not M_OBJRS.EOF
'            cmbDescUn.AddItem M_OBJRS("Description")
'            M_OBJRS.MoveNext
'        Wend
'        C_Payment.Value = 0
''        FrmPayment.Enabled = False
'        Set M_OBJRS = Nothing
'Else
'If Left(cmbUncontacted.Text, 2) <> "MV" Then
'   Set M_OBJRS = New ADODB.Recordset
'   M_OBJRS.CursorLocation = adUseClient
'   M_OBJRS.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
'         For i = 0 To 3
'           cmbDescUn.AddItem M_OBJRS("Description")
'           M_OBJRS.MoveNext
'         Next i
'   Set M_OBJRS = Nothing
'   C_Payment.Value = 0
'Else
'   Set M_OBJRS = New ADODB.Recordset
'   M_OBJRS.CursorLocation = adUseClient
''   If kontak = True Then
''        m_objrs.Open "Select * from DescUncontacted where ", M_OBJCONN, adOpenDynamic, adLockOptimistic
''    Else
'      M_OBJRS.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
''    End If
'       While Not M_OBJRS.EOF
'           cmbDescUn.AddItem M_OBJRS("Description")
'           M_OBJRS.MoveNext
'       Wend
'   Set M_OBJRS = Nothing
'   C_Payment.Value = 0
'End If
'End If
'End Sub
'
'Private Sub cmbDescUn_KeyDown(KeyCode As Integer, Shift As Integer)
''MsgBox "Jangan di ketik, tapi di pilih, Bisa ga sih kamu...  !!!"
'cmbDescUn.Text = ""
'Exit Sub
'End Sub
'
'Private Sub cmbDiscount_Click()
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
'End Sub
'
'
'
'Private Sub cmbDiscount_LostFocus()
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
'End Sub
'
'
'
'
'Private Sub cmbNextAct_KeyDown(KeyCode As Integer, Shift As Integer)
''MsgBox "Jangan di ketik, tapi di pilih, Bisa ga sih kamu...  !!!"
'cmbNextAct.Text = ""
'Exit Sub
'End Sub
'
'
'
'Private Sub cmbUncontacted_Click()
''DESCRIPTION UNCONTACTED
'Dim i As Integer
'cmbDescUn.CLEAR
'If Left(cmbUncontacted.Text, 2) = "NA" Then
'        cmbDescUn.Enabled = True
''        CmbBaseOn.Text = ""
''        txtPayment.Text = 0
''        cmbDiscount.Text = ""
''        TdbPTP.Text = ""
''        TdbDatePTP.Text = ""
'        Set M_OBJRS = New ADODB.Recordset
'        M_OBJRS.CursorLocation = adUseClient
'          M_OBJRS.Open "Select * from DescContacted WHERE id >= 13 ", M_OBJCONN, adOpenDynamic, adLockOptimistic
'        While Not M_OBJRS.EOF
'            cmbDescUn.AddItem M_OBJRS("Description")
'            M_OBJRS.MoveNext
'        Wend
'        C_Payment.Value = 0
'        Set M_OBJRS = Nothing
''        FrmPayment.Enabled = False
'Else
'If Left(cmbUncontacted.Text, 2) <> "MV" Then
'   Set M_OBJRS = New ADODB.Recordset
'   M_OBJRS.CursorLocation = adUseClient
'   M_OBJRS.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
'         For i = 0 To 3
'           cmbDescUn.AddItem M_OBJRS("Description")
'           M_OBJRS.MoveNext
'         Next i
'   Set M_OBJRS = Nothing
'   C_Payment.Value = 0
'Else
'   Set M_OBJRS = New ADODB.Recordset
'   M_OBJRS.CursorLocation = adUseClient
''   If kontak = True Then
''        m_objrs.Open "Select * from DescUncontacted where ", M_OBJCONN, adOpenDynamic, adLockOptimistic
''    Else
'      M_OBJRS.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
''    End If
'       While Not M_OBJRS.EOF
'           cmbDescUn.AddItem M_OBJRS("Description")
'           M_OBJRS.MoveNext
'       Wend
'   Set M_OBJRS = Nothing
'   C_Payment.Value = 0
'End If
'End If
'' Set M_OBJRS = New ADODB.Recordset
''   If kontak = False Then
''          M_OBJRS.Open "Select * from UncontactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
''       While Not M_OBJRS.EOF
''           cmbDescUn.AddItem M_OBJRS("NMnoProdpresented")
''           M_OBJRS.MoveNext
''       Wend
''        Set M_OBJRS = Nothing
''   End If
''   C_Payment.Value = 0
''End If
'
'End Sub
'
'Private Sub headerDatePayment()
'    LstPayment.ColumnHeaders.ADD 1, , "", 0 * TXT
'    LstPayment.ColumnHeaders.ADD 2, , "Id", 2 * TXT
'    LstPayment.ColumnHeaders.ADD 3, , "Tgl Janji", 15 * TXT
'    LstPayment.ColumnHeaders.ADD 4, , "Jumlah", 30 * TXT
'    LstPayment.ColumnHeaders.ADD 5, , "Jenis", 30 * TXT
'    LstPayment.ColumnHeaders.ADD 6, , "Tgl Entry", 15 * TXT
'End Sub
'
'Private Sub headerCustid_Double()
'    LstDoubleId.ColumnHeaders.ADD 1, , "CUSTID", 15 * TXT
'    LstDoubleId.ColumnHeaders.ADD 2, , "NAME", 30 * TXT
'    LstDoubleId.ColumnHeaders.ADD 3, , "Agent", 10 * TXT
'    LstDoubleId.ColumnHeaders.ADD 4, , "AMOUNTWO", 10 * TXT
'    LstDoubleId.ColumnHeaders.ADD 5, , "PRICIPLE", 20 * TXT
'End Sub
'
'
'Private Sub cmbUncontacted_KeyDown(KeyCode As Integer, Shift As Integer)
''MsgBox "Jangan di ketik, tapi di pilih, Bisa ga sih kamu...  !!!"
'cmbUncontacted.Text = ""
'Exit Sub
'End Sub
'
'Private Sub CmdDeletePelunasan_Click()
'Dim m_msgbox As Variant
'If listview1(0).ListItems.Count = 0 Then
'    Exit Sub
'End If
'm_msgbox = MsgBox("Yakin Akan Dihapus...!!! ", vbCritical + vbOKCancel, "Peringatan")
'If m_msgbox = vbOK Then
'    M_OBJCONN.Execute "Delete from tbllunas where id = " + listview1(0).SelectedItem.SubItems(4) + ""
'    listview1(0).ListItems.Remove listview1(0).SelectedItem.Index
'    MsgBox "Done"
'    Call isi_datapayment
'End If
'End Sub
'
'
'Private Sub Form_Load()
'qt = ""
''cek list pelunasan
'Dim i, iIndex As Integer
'Dim sKata, cCombo As String
'
'
''------->>>  setting No Visit  <<<---------------
'
'Text1.Text = Format(Now, "yymmddhhmmss")
'TDBDate1.Value = Now
'If UCase(Left(mdiform1.txtlevel.text, 5)) = "ADMIN" Or UCase(Left(mdiform1.txtlevel.text, 5)) = "SUPER" Or UCase(mdiform1.txtlevel.text) = "TEAMLEADER" Then
''If UCase(Left(mdiform1.txtlevel.text, 5)) = "ADMIN" Then
'    txtHomeNo1.Visible = True
'    txtHomeNo1A.Visible = False
'    txtHomeNo2.Visible = True
'    txtHomeNo2A.Visible = False
'    txtOfficeNo1.Visible = True
'    txtOfficeNo1A.Visible = False
'    txtOfficeNo2.Visible = True
'    txtOfficeNo2A.Visible = False
'    txtMobileNo1.Visible = True
'    txtMobileNo1A.Visible = False
'    txtMobileNo2.Visible = True
'    txtMobileNo2A.Visible = False
'    txtPhone.Visible = True
'    txtPhoneA.Visible = False
'    txtHomeAdd1.Visible = True
'    txtHomeAdd1A.Visible = False
'    txtHomeAdd2.Visible = True
'    txtHomeAdd2A.Visible = False
'    txtOfficeAdd1.Visible = True
'    txtOfficeAdd1A.Visible = False
'    txtOfficeAdd2.Visible = True
'    txtOfficeAdd2A.Visible = False
'    txtMobileAdd1.Visible = True
'    txtMobileAdd1A.Visible = False
'    txtMobileAdd2.Visible = True
'    txtMobileAdd2A.Visible = False
'    txtECno.Visible = True
'    txtECnoA.Visible = False
'End If
'
'If UCase(mdiform1.txtlevel.text) = "AGENT" Then
'        C_lunas.Enabled = False
'        TdbLunas.Enabled = False
'        chkAppv(0).Enabled = False
'        chkAppv(1).Enabled = False
'        TDBTot_payment.Enabled = False
'        TxtFieldName.Enabled = False
'        CmdDeletePelunasan.Enabled = False
'Else
'        txtHomeAdd1.ReadOnly = False
'        txtHomeAdd2.ReadOnly = False
'        txtOfficeAdd1.ReadOnly = False
'        txtOfficeAdd2.ReadOnly = False
'        txtMobileAdd1.ReadOnly = False
'        txtMobileAdd2.ReadOnly = False
'
'End If
'
'   FrmContacted.Enabled = False
'   'FrmUnContacted.Enabled = False
'   FrmPayment.Enabled = False
'   FrMValid.Enabled = False
'   FrmSKIP.Enabled = False
'   frmPTP.Enabled = False
'   frmpopsp.Enabled = False
'   'cboPTP.Enabled = False
'
'    Call headerDatePayment
'    Call headerCustid_Double
'    Call HEADER_HISTORY
'    Call HEADER_HISTORY_PAID
'    Call HEADER_RequestVisit
'    Call show_cust
'    Call Custid_Double
'    Call VisitNo
'    'Call isi_lastcall
'
'    If UCase(mdiform1.txtlevel.text) = "TEAMLEADER" Or UCase(mdiform1.txtlevel.text) = "SUPERVISOR" Or UCase(mdiform1.txtlevel.text) = "ADMINISTRATOR" Then
'        Call aktifphone
'    End If
' '   SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
'SSTab1.Tab = 0
'cmbDateSch.Value = Now
'cmbDateSch.Value = ""
''CONTACTED
'CmbBaseOn.AddItem "PRINCIPLE"
'CmbBaseOn.AddItem "TOTAL AMOUNT"
''SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
'
''LIST VALID
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from tblvalid", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    While Not M_OBJRS.EOF
'        cbovalid.AddItem M_OBJRS!KdNoProdPresented
'        M_OBJRS.MoveNext
'    Wend
'Set M_OBJRS = Nothing
''LIST Contacted
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from ContactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    While Not M_OBJRS.EOF
'        cmbContacted.AddItem M_OBJRS!KdNoProdPresented
'        M_OBJRS.MoveNext
'    Wend
'Set M_OBJRS = Nothing
''LIST PTP
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from tblPTP", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    While Not M_OBJRS.EOF
'        cboPTP.AddItem M_OBJRS!KdNoProdPresented
'        M_OBJRS.MoveNext
'    Wend
'Set M_OBJRS = Nothing
''LIST SKIP
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from tblSKIP", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    While Not M_OBJRS.EOF
'        cboskip.AddItem M_OBJRS!KdNoProdPresented
'        M_OBJRS.MoveNext
'    Wend
' Set M_OBJRS = Nothing
''LIST POPSP
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'If UCase(Left(cboPOPSP.Text, 2)) = "SP" Then
'    M_OBJRS.Open "Select * from POPSPDesc WHERE KdNoProdPresented LIKE '%SP%'", M_OBJCONN, adOpenDynamic, adLockOptimistic
'ElseIf UCase(mdiform1.txtlevel.text) = "AGENT" Then
'    M_OBJRS.Open "Select * from POPSPDesc WHERE substring(KdNoProdPresented,1,2) not in ('SP','PR')", M_OBJCONN, adOpenDynamic, adLockOptimistic
'Else
'    M_OBJRS.Open "Select * from POPSPDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
'End If
'    While Not M_OBJRS.EOF
'        cboPOPSP.AddItem M_OBJRS!KdNoProdPresented
'        M_OBJRS.MoveNext
'    Wend
'Set M_OBJRS = Nothing
'If C_POPSP.Value Then
'    C_VALID.Enabled = False
'    'C_Contacted.Enabled = False
'    'C_PTP.Enabled = False
'    C_SKIP.Enabled = False
'End If
'
''If UCase(mdiform1.txtlevel.text) = "AGENT" Then
''    C_POPSP.Enabled = False
''    frmpopsp.Enabled = False
''End If
''---REMARKS BY RIF
''Set m_objrs = New ADODB.Recordset
''m_objrs.Open "Select * from ContactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
''
''    While Not m_objrs.EOF
''    '----tambahan 05 Maret 2007----'
''         scombo = m_objrs("KdNoProdPresented")
''            sKata = cmbContacted.Text
''            ' initialisasi index
''            If scombo = "BP-BROKEN PROMISE" Or scombo = "PTP-PROMISE TO PAY" Or scombo = "RP-REFUSE PAYMENT" Then
''                  iIndex = 1
''            ElseIf scombo = "POP-PROGRESS OF PAYMENT" Then
''                  iIndex = 2
''            ElseIf scombo = "SP-SETTLE PAYMENT" Then
''                  iIndex = 3
''            Else
''                  iIndex = 4
''            End If
''
''            ' saring tampilan
''            If iIndex = 1 Then
''               If iIndex = 4 Or sKata = "POP-PROGRESS OF PAYMENT" Then
''                  'lewat boo
''               Else
''                  cmbContacted.AddItem scombo
''               End If
''            ElseIf iIndex = 2 Then
''               If (iIndex = 1 Or iIndex = 4) Then
''                  'lewat boo
''               Else 'If UCase(mdiform1.txtlevel.text) <> "AGENT" Then
''                  cmbContacted.AddItem scombo
''               End If
''            ElseIf iIndex = 3 Then
''                If UCase(mdiform1.txtlevel.text) = "AGENT" Then
''                Else
''                  cmbContacted.AddItem scombo
''                End If
''            Else
''                  If sKata = "BP-BROKEN PROMISE" Or sKata = "PTP-PROMISE TO PAY" Or sKata = "POP-PROGRESS OF PAYMENT" Then
''                  'If sKata = "BP-BROKEN PROMISE" Or sKata = "PTP-PROMISE TO PAY" Then
''                     'lewat boo
''                  Else
''                     cmbContacted.AddItem scombo
''                  End If
''            End If
''            m_objrs.MoveNext
''    Wend
''Set m_objrs = Nothing
''
''
'''UNCONTACTED
''If Left(cmbContacted.Text, 2) = "SP" Then
''    cmbContacted.Clear
''    cmbContacted.Text = "SP-SETTLE PAYMENT"
''    C_Contacted.Enabled = False
''    C_NotContacted.Enabled = False
''End If
''
''Set m_objrs = New ADODB.Recordset
''m_objrs.CursorLocation = adUseClient
''If kontak = True Then
''    m_objrs.Open "Select * from UnContactedDesc where KdNoProdPresented IN ('NBP-NO BODY PICKUP','NA-NOT AVAILABLE')", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
''ElseIf Left(VIEW_mgmDATA.LstVwSearchmgm.SelectedItem.SubItems(8), 2) = "NA" Then
''    m_objrs.Open "Select * from UnContactedDesc where KdNoProdPresented IN ('NBP-NO BODY PICKUP','NA-NOT AVAILABLE')", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
''Else
''    m_objrs.Open "Select * from UnContactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
''End If
''    While Not m_objrs.EOF
''        cmbUncontacted.AddItem m_objrs("KdNoProdPresented")
''        'cmbDescUn.AddItem M_OBJRS("nmNoProdPresented")
''        m_objrs.MoveNext
''    Wend
''Set m_objrs = Nothing
''
''If cmbContacted.Text = "POP-PROGRESS OF PAYMENT" Then
''    C_NotContacted.Enabled = False
''End If
''
''
''
'''Set M_OBJRS = New ADODB.Recordset
'''If kontak = True Then
'''    C_NotContacted.Enabled = False
'''Else
'''    M_OBJRS.Open "Select * from UnContactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
'''    While Not M_OBJRS.EOF
'''        cmbUncontacted.AddItem M_OBJRS("KdNoProdPresented")
'''        'cmbDescUn.AddItem M_OBJRS("nmNoProdPresented")
'''        M_OBJRS.MoveNext
'''    Wend
'''End If
'''Set M_OBJRS = Nothing
'
'
'
'
''DISCOUNT
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from tblDiscount", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'    While Not M_OBJRS.EOF
'        cmbDiscount.AddItem M_OBJRS("Description")
'        M_OBJRS.MoveNext
'    Wend
'Set M_OBJRS = Nothing
'
''NEXT ACTION
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from StsNextAct", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'While Not M_OBJRS.EOF
'    cmbNextAct.AddItem M_OBJRS("NmStsNextAct")
'    M_OBJRS.MoveNext
'Wend
'Set M_OBJRS = Nothing
''untuk 108
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from tbllayanantelkom", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'While Not M_OBJRS.EOF
'    CmbPhone.AddItem IIf(IsNull(M_OBJRS("Nolayanan")), "", M_OBJRS("Nolayanan"))
'    M_OBJRS.MoveNext
'Wend
'Set M_OBJRS = Nothing
'End Sub
'
'Sub isi_lastcall()
'cbolastcall.CLEAR
'Dim M_OBJRS As ADODB.Recordset
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from ContactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
'While Not M_OBJRS.EOF
'    cbolastcall.AddItem M_OBJRS("KdNoProdPresented")
'    M_OBJRS.MoveNext
'Wend
'Set M_OBJRS = Nothing
'
''Set m_objrs = New ADODB.Recordset
''m_objrs.Open "Select * from UnContactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
''While Not m_objrs.EOF
''    cbolastcall.AddItem m_objrs("KdNoProdPresented")
''    m_objrs.MoveNext
''Wend
''Set m_objrs = Nothing
'
''LIST VALID
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from tblvalid", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    While Not M_OBJRS.EOF
'        cbolastcall.AddItem M_OBJRS!KdNoProdPresented
'        M_OBJRS.MoveNext
'    Wend
'Set M_OBJRS = Nothing
''LIST Contacted
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from ContactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    While Not M_OBJRS.EOF
'        cbolastcall.AddItem M_OBJRS!KdNoProdPresented
'        M_OBJRS.MoveNext
'    Wend
'    Set M_OBJRS = Nothing
''LIST PTP
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from tblPTP", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    While Not M_OBJRS.EOF
'        cbolastcall.AddItem M_OBJRS!KdNoProdPresented
'        M_OBJRS.MoveNext
'    Wend
' Set M_OBJRS = Nothing
''LIST SKIP
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from tblSKIP", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    While Not M_OBJRS.EOF
'        cbolastcall.AddItem M_OBJRS!KdNoProdPresented
'        M_OBJRS.MoveNext
'    Wend
' Set M_OBJRS = Nothing
''LIST POPSP
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from POPSPDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    While Not M_OBJRS.EOF
'        cbolastcall.AddItem M_OBJRS!KdNoProdPresented
'        M_OBJRS.MoveNext
'    Wend
'Set M_OBJRS = Nothing
'End Sub
'
'Private Sub aktifphone()
'AHomeAdd1(0).ReadOnly = False
'AHomeAdd2(1).ReadOnly = False
'txtHomeAdd1.ReadOnly = False
'txtHomeAdd1A.ReadOnly = False
'txtHomeAdd2.ReadOnly = False
'txtHomeAdd2A.ReadOnly = False
'AOfficeAdd(2).ReadOnly = False
'AOfficeAdd(3).ReadOnly = False
'txtOfficeAdd1.ReadOnly = False
'txtOfficeAdd1A.ReadOnly = False
'txtOfficeAdd2.ReadOnly = False
'txtOfficeAdd2A.ReadOnly = False
'AFaxAdd(4).ReadOnly = False
'AFaxAdd(5).ReadOnly = False
'txtFaxAdd1.ReadOnly = False
'txtFaxAdd2.ReadOnly = False
'txtMobileAdd1.ReadOnly = False
'txtMobileAdd1A.ReadOnly = False
'txtMobileAdd2.ReadOnly = False
'txtMobileAdd2A.ReadOnly = False
'End Sub
'
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Set M_Col = Nothing
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'kontak = False
'shedulePTP_Show = False
'' 'M_OBJCONN.Close
'M_OBJCONN.Close
'Set M_OBJCONN = Nothing
'M_OBJCONN.Open CMDSQLOPEN
'VIEW_MGMDATA.WindowState = 2
'End Sub
'
'
'
'
'Private Sub ListView1_Click(Index As Integer)
'Dim KET As String
'Select Case Index
'Case 0
'
'Case 1
'If listview1(1).ListItems.Count = 0 Then
'Exit Sub
'Else
'   KET = TXtDetails.Text
'      If Len(TXtDetails) = 0 Then
'         TXtDetails.Text = " - " + listview1(1).SelectedItem.SubItems(1)
'      Else
'         TXtDetails.Text = KET + " - " + listview1(1).SelectedItem.SubItems(1)
'      End If
'End If
'End Select
'End Sub
'
'Private Sub LstPayment_DblClick()
'If LstPayment.ListItems.Count = 0 Then
'Exit Sub
'Else
'Call SSCommand2_Click(1)
'End If
'End Sub
'
'
'
'Private Sub LstVisit_DblClick()
' If LstVisit.ListItems.Count > 0 Then
'
'
'           With FRM_UpdateVisit
'                .Text1.Text = LstVisit.SelectedItem.SubItems(2)
'                .Show vbModal
'
'
''                    M_DATA.UPDATE_NegoPTP M_OBJCONN, .TxtCustid.Text, .TDBDate1.Value, CStr(.TDBNumber1.Value), LstPayment.SelectedItem.SubItems(1)
''
''                    On Error GoTo add_error
''                    If M_DATA.ADD_OK Then
''                        'LstPayment.SelectedItem.SubItems(1) = ""
''                        LstPayment.SelectedItem.SubItems(2) = .TDBDate1.Value
''                        LstPayment.SelectedItem.SubItems(3) = .TDBNumber1.Value
''
''
''                    On Error GoTo 0
''                    End If
''                End If
'               End With
'Else
'Exit Sub
'End If
'
'End Sub
'
'Private Sub Option1_Click()
'If Option1.Value = True Then
'TYPETELP = ""
'   txtPhone.Text = GetNumber(CStr(AHome1.Value & txtHomeNo1.Value))
'   If txtHomeNo1.Value <> "" Then
'        txtPhoneA.Text = CStr(AHome1.Value & txtHomeNo1A.Value)
'    Else
'        txtPhoneA.Text = ""
'    End If
'   Option2.Value = False
'   Option3.Value = False
'   Option4.Value = False
'   Option5.Value = False
'End If
'End Sub
'
'Private Sub Option2_Click()
'If Option2.Value = True Then
'TYPETELP = ""
'   txtPhone.Text = GetNumber(CStr(AHome2.Value & txtHomeNo2.Value))
'   If txtHomeNo2.Value <> "" Then
'        txtPhoneA.Text = CStr(AHome2.Value & txtHomeNo2A.Value)
'    Else
'        txtPhoneA.Text = ""
'    End If
'   Option1.Value = False
'   Option3.Value = False
'   Option4.Value = False
'   Option5.Value = False
'End If
'End Sub
'
'Private Sub Option3_Click()
'   If Option3.Value = True Then
'   TYPETELP = ""
'   txtPhone.Text = GetNumber(CStr(AOffice2.Value & txtOfficeNo2.Value))
'   If txtOfficeNo2.Value <> "" Then
'        txtPhoneA.Text = CStr(AOffice2.Value & txtOfficeNo2A.Value)
'    Else
'        txtPhoneA.Text = ""
'   End If
'   Option2.Value = False
'   Option4.Value = False
'   Option1.Value = False
'   Option5.Value = False
'   End If
'End Sub
'Private Sub Option4_Click()
'   If Option4.Value = True Then
'   TYPETELP = ""
'   txtPhone.Text = GetNumber(CStr(AOffice1.Value & txtOfficeNo1.Value))
'   If txtOfficeNo1.Value <> "" Then
'        txtPhoneA.Text = CStr(AOffice1.Value & txtOfficeNo1A.Value)
'    Else
'        txtPhoneA.Text = ""
'   End If
'   Option2.Value = False
'   Option3.Value = False
'   Option1.Value = False
'   Option5.Value = False
'End If
'End Sub
'Private Sub Option5_Click()
' If Option5.Value = True Then
' TYPETELP = ""
'   txtPhone.Text = GetNumber(CStr(txtMobileNo2.Value))
'    If txtMobileNo2.Value <> "" Then
'        txtPhoneA.Text = CStr(txtMobileNo2A.Value)
'    Else
'        txtPhoneA.Text = ""
'   End If
'   Option2.Value = False
'   Option3.Value = False
'   Option1.Value = False
'   Option4.Value = False
'   Option6.Value = False
'   End If
'End Sub
'
'Private Sub Option6_Click()
' If Option6.Value = True Then
' TYPETELP = ""
'   txtPhone.Text = GetNumber(CStr(txtMobileNo1.Value))
'   If txtMobileNo1.Value <> "" Then
'        txtPhoneA.Text = CStr(txtMobileNo1A.Value)
'    Else
'        txtPhoneA.Text = ""
'   End If
'   Option2.Value = False
'   Option3.Value = False
'   Option1.Value = False
'   Option4.Value = False
'   Option5.Value = False
'   End If
'End Sub
'
'Private Sub Option7_Click(Index As Integer)
'Select Case Index
'Case 0
'TxtAddress.Text = AddrNow.Text
'Case 1
'TxtAddress.Text = lblAddr.Text
'Case 2
'TxtAddress.Text = lblOfficeAddr.Text
'End Select
'
'End Sub
'
'Private Sub Option8_Click(Index As Integer)
'Select Case Index
'Case 0
'Frame8.Enabled = True
'VisitYES
'Case 1
'VisitNo
'Frame8.Enabled = False
'End Select
'End Sub
'
'Private Sub SSCommand1_Click(Index As Integer)
'Select Case Index
'  Case 0
'  If Len(CmbPhone.Text) < 2 Then
'    MsgBox "Pilihan No Telephone harus diisi"
'    CmbPhone.SetFocus
'    Exit Sub
'  End If
'
'    If Len(CmbPhone.Text) > 1 Then
'        idcust = lblCustId.Caption
'
'        Select Case CmbPhone
'            Case "Hp"
'                txtPhone.Text = txtMobileNo1.Value
'                telpno = txtPhone.Text
'            Case "Hp2"
'                txtPhone.Text = txtMobileNo2.Value
'                telpno = txtPhone.Text
'            Case "HomePhone"
'                If AHome1.Value = "021" Or AHome1.Value = "" Then
'                    txtPhone.Text = txtHomeNo1.Value
'                Else
'                    txtPhone.Text = AHome1.Value & txtHomeNo1.Value
'                End If
'                telpno = txtPhone.Text
'            Case "HomePhone2"
'                If AHome1.Value = "021" Or AHome1.Value = "" Then
'                    txtPhone.Text = txtHomeNo2.Value
'                Else
'                    txtPhone.Text = AHome1.Value & txtHomeNo2.Value
'                End If
'                telpno = txtPhone.Text
'            Case "OfficePhone"
'                If AOffice1.Value = "021" Or AOffice1.Value = "" Then
'                    txtPhone.Text = txtOfficeNo1.Value
'                Else
'                    txtPhone.Text = AOffice1.Value & txtOfficeNo1.Value
'                End If
'                telpno = txtPhone.Text
'            Case "OfficePhone2"
'                If AOffice2.Value = "021" Or AOffice2.Value = "" Then
'                    txtPhone.Text = txtOfficeNo2.Value
'                Else
'                    txtPhone.Text = AOffice1.Value & txtOfficeNo2.Value
'                End If
'                telpno = txtPhone.Text
'            Case "EconPhone"
'                If txtECno.ReadOnly = False And UCase(mdiform1.txtlevel.text) = "AGENT" Then
'                    MsgBox "Simpan Data terlebih dahulu"
'                    Exit Sub
'                End If
'                txtPhone.Text = txtECno.Value
'                telpno = txtPhone.Text
'            Case "AddHome1"
'                If txtHomeAdd1A.ReadOnly = False And UCase(mdiform1.txtlevel.text) = "AGENT" Then
'                    MsgBox "Simpan Data terlebih dahulu"
'                    Exit Sub
'                End If
'                If AHomeAdd1(0).Value = "021" Or AHomeAdd1(0).Value = "" Then
'                    txtPhone.Text = txtHomeAdd1.Value
'                Else
'                    txtPhone.Text = AHomeAdd1(0).Value & txtHomeAdd1.Value
'                End If
'                telpno = txtPhone.Text
'            Case "AddHome2"
'                If txtHomeAdd2A.ReadOnly = False And UCase(mdiform1.txtlevel.text) = "AGENT" Then
'                    MsgBox "Simpan Data terlebih dahulu"
'                    Exit Sub
'                End If
'                If AHomeAdd2(1).Value = "021" Or AHomeAdd2(1).Value = "" Then
'                    txtPhone.Text = txtHomeAdd2.Value
'                Else
'                    txtPhone.Text = AHomeAdd2(1).Value & txtHomeAdd2.Value
'                End If
'                telpno = txtPhone.Text
'            Case "AddOffice1"
'                If txtOfficeAdd1A.ReadOnly = False And UCase(mdiform1.txtlevel.text) = "AGENT" Then
'                    MsgBox "Simpan Data terlebih dahulu"
'                    Exit Sub
'                End If
'                If AOfficeAdd(2).Value = "021" Or AOfficeAdd(2).Value = "" Then
'                    txtPhone.Text = txtOfficeAdd1.Value
'                Else
'                    txtPhone.Text = AOfficeAdd(2).Value & txtOfficeAdd1.Value
'                End If
'                telpno = txtPhone.Text
'            Case "AddOffice2"
'                If txtOfficeAdd2A.ReadOnly = False And UCase(mdiform1.txtlevel.text) = "AGENT" Then
'                    MsgBox "Simpan Data terlebih dahulu"
'                    Exit Sub
'                End If
'                If AOfficeAdd(3).Value = "021" Or AOfficeAdd(3).Value = "" Then
'                    txtPhone.Text = txtOfficeAdd2.Value
'                Else
'                    txtPhone.Text = AOfficeAdd(3).Value & txtOfficeAdd2.Value
'                End If
'                telpno = txtPhone.Text
'            Case "AddMobile1"
'                If txtMobileAdd1A.ReadOnly = False And UCase(mdiform1.txtlevel.text) = "AGENT" Then
'                    MsgBox "Simpan Data terlebih dahulu"
'                    Exit Sub
'                End If
'                txtPhone.Text = txtMobileAdd1.Value
'                telpno = txtPhone.Text
'            Case "AddMobile2"
'                If txtMobileAdd2A.ReadOnly = False And UCase(mdiform1.txtlevel.text) = "AGENT" Then
'                    MsgBox "Simpan Data terlebih dahulu"
'                    Exit Sub
'                End If
'                txtPhone.Text = txtMobileAdd2.Value
'                telpno = txtPhone.Text
'            Case Else
'                txtPhone.Text = Trim(CmbPhone.Text)
'        End Select
'    'untuk ritpil
'    MDIForm1.ActionCTI ("DIAL|49682" & GetNumber(CStr(Replace(txtPhone.Text, " ", ""))) & "|" & Trim(FrmCC_Colection.lblCustId.Caption) & "|" & Trim(FrmCC_Colection.lblCustId.Caption))
'    'untuk awarness
'    ' MDIForm1.ActionCTI ("DIAL|54869" & GetNumber(CStr(Replace(txtPhone.Text, " ", ""))) & "|" & Trim(frmCC_Colection.lblCustId.Caption) & "|" & Trim(frmCC_Colection.lblCustId.Caption))
'        'CMDSQL = "Insert Into TblPhoneMonitorHst(UserId, CustId, NamaCh,StartDate, TelpNo, Recsource) Values ('" + mdiform1.txtusername.text + "' , '" + frmCC_Colection.lblCustId.Caption + "','" + frmCC_Colection.lblNama.Caption + "', '" + Format(CStr(MDIForm1.TDBDate1.Value), "mm/dd/yyyy") & " " & Format(Now, "hh:nn") + "' , '" + Replace(txtPhone.Text, " ", "") + "' ,'" + frmCC_Colection.lblRecsource.Caption + "')"
'        'M_OBJCONN.Execute CMDSQL
'        MDIForm1.CmbNo.Text = ""
'        stscall = True
'End If
'TYPETELP = ""
'   Case 2
'        V_SAVE = CEK_DATA_VALID
'        If V_SAVE = False Then
'            Exit Sub
'        Else
'        End If
'        If ADD_CUST Then
'            'Call CEK_ADD_PELANGGAN
'        Else
'            Call CEK_UPDATE_PELANGGAN
'            stscall = False
'            Call isi_datapayment
'        End If
'   Case 3
'    kontak = False
'        Unload Me
'    Case 1
'        MDIForm1.ActionCTI ("HANGUP")
'End Select
'End Sub
'
'Public Sub Show_NEGOPTP()
'Dim showlist As New ADODB.Recordset
'Dim CMDSQL As String
'
''CMDSQL = "SELECT * FROM tblnegoPTP where custid = '" + lblCustId.Caption + "' AND TGLSOURCEorder by promisedate desc"
'CMDSQL = "SELECT * FROM ReportPTP where custid = '" + lblCustId.Caption + "'"
'Set showlist = New ADODB.Recordset
'showlist.CursorLocation = adUseClient
'showlist.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
'
'
'LstPayment.ListItems.CLEAR
'While Not showlist.EOF
'    Set listitem = LstPayment.ListItems.ADD(, , "")
'        'listitem.SubItems(1) = ""
'        listitem.SubItems(1) = CStr(IIf(IsNull(showlist!ID), "", (showlist!ID)))
'        listitem.SubItems(2) = CStr(IIf(IsNull(showlist!PromiseDate), "", Format(showlist!PromiseDate, "dd/mm/yyyy")))
'        listitem.SubItems(3) = CStr(IIf(IsNull(showlist!PromisePay), "", (showlist!PromisePay)))
'        n = n + Val(listitem.SubItems(3))
'        If n <= TOTPTP Then
'            listitem.ListSubItems(1).ForeColor = vbRed
'            listitem.ListSubItems(2).ForeColor = vbRed
'            listitem.ListSubItems(3).ForeColor = vbRed
'        End If
'        listitem.SubItems(4) = IIf(IsNull(showlist!Type), "", showlist!Type)
'        listitem.SubItems(5) = CStr(IIf(IsNull(showlist!inputdate), "", Format(showlist!inputdate, "dd/mm/yyyy")))
'     showlist.MoveNext
''    Set listitem = LstPayment.ListItems.ADD(, , "")
''        listitem.SubItems(1) = CStr(IIf(IsNull(ShowList!ID), "", (ShowList!ID)))
''        listitem.SubItems(2) = CStr(IIf(IsNull(ShowList!PromiseDate), "", (ShowList!PromiseDate)))
''        listitem.SubItems(3) = CStr(IIf(IsNull(ShowList!PromisePay), "", (ShowList!PromisePay)))
'Wend
'Set showlist = Nothing
'End Sub
'Public Sub show_cust()
'Dim listitem As listitem
'Dim M_DATA As New CLS_FRMCUST_CC
'Dim m_cust1 As ADODB.Recordset
'Dim m_cust2 As ADODB.Recordset
'Dim CMDSQL As String
'Dim CMDSQL2 As String
'Dim sPending As String
'Dim batchtype As String
'On Error GoTo HELL:
'
''CMDSQL = "SELECT mgm.*, mgm_DETAIL.* FROM mgm INNER JOIN "
''CMDSQL = CMDSQL + "mgm_DETAIL ON mgm.CUSTID = dbo.mgm_DETAIL.CUSTID"
'
'CMDSQL = "select * from mgm"
''CMDSQL2 = "select * from mgm_detail"
'
'Set M_Col = New ADODB.Recordset
''Set m_col2 = New ADODB.Recordset
'M_Col.CursorLocation = adUseClient
''m_col2.CursorLocation = adUseClient
'If shedulePTP_Show = True Then
'    CMDSQL = CMDSQL + " where custid ='" & MDIForm1.LstGrade.SelectedItem.SubItems(1) & "'"
'    M_Col.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
'
'Else
'    CMDSQL = CMDSQL + " where custid ='" & VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(1) & "'"
'    M_Col.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
'    'CMDSQL2 = CMDSQL2 + " where custid ='" & VIEW_mgmDATA.LstVwSearchmgm.SelectedItem.SubItems(1) & "'"
'    'm_col2.Open CMDSQL2, M_OBJCONN, adOpenDynamic, adLockOptimistic
'    'm_col.Open "Select * from mgm where custid='" & VIEW_mgmDATA.LstVwSearchmgm.SelectedItem.SubItems(1) & "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
'End If
'
''tampilkan data tabel mgm
'If Not M_Col.EOF Then
'    lblstatus.Caption = IIf(IsNull(M_Col("statusprior")), "", "Status : " & M_Col("statusprior"))
'    lblCustId.Caption = IIf(IsNull(M_Col("CUSTID")), "", M_Col("CUSTID"))
'    TxtCustid.Text = IIf(IsNull(M_Col("CUSTID")), "", M_Col("CUSTID"))
'    TxtName.Text = IIf(IsNull(M_Col("NAME")), "", M_Col("NAME"))
'    LblInterest.Caption = Format(IIf(IsNull(M_Col("INTEREST")), "0", M_Col("INTEREST")), "##,###")
'    LblFees.Caption = Format(IIf(IsNull(M_Col("FEES")), "0", M_Col("FEES")), "##,###")
'
'    lblRecsource.Caption = IIf(IsNull(M_Col("RECSOURCE")), "", M_Col("RECSOURCE"))
'    LBLEXP.Caption = IIf(IsNull(M_Col("RECSOURCE")), "", "Expire date " & Format(DateAdd("d", 90, Format(M_Col("TGLSOURCE"), "MM-DD-yyyy")), "dd-mm-yyyy"))
'     LblRiskLevel.Caption = IIf(IsNull(M_Col("RiskLevel")), "", M_Col("RiskLevel"))
'    lblPriority.Caption = IIf(IsNull(M_Col("Priority")), "", M_Col("Priority"))
'    lblNama.Caption = IIf(IsNull(M_Col("NAME")), "", M_Col("NAME"))
'    lblCardNo.Caption = IIf(IsNull(M_Col("NoCard")), "", M_Col("NoCard"))
'    lblID.Caption = IIf(IsNull(M_Col("ktpno")), "", M_Col("ktpno"))
'    'lblDate.Value = IIf(IsNull(m_col("BIRTHD")), "", Format(m_col("BIRTHD"), "dd-mmm-yyyy"))
'    statusptp2 = IIf(IsNull(M_Col!F_CEK), "", M_Col!F_CEK)
'
'
''    If statusptp2 = "" Then
''        sqlbck = "insert into tbllunashst (custid,paydate,payment,agent,fieldname,datafrom,sts,id) select custid,paydate,payment,agent,fieldname,datafrom,sts,id from tbllunas where custid ='" + IIf(IsNull(M_Col("CUSTID")), "", M_Col("CUSTID")) + "'"
''        M_OBJCONN.Execute (sqlbck)
''        sqldel = "delete from tbllunas where custid='" + IIf(IsNull(M_Col("CUSTID")), "", M_Col("CUSTID")) + "'"
''        M_OBJCONN.Execute (sqldel)
''        sqlbcknego = "insert into tblnegoptphst (id,custid,promisedate,promisepay,inputdate) select id,custid,promisedate,promisepay,inputdate from tblnegoptp where custid ='" + IIf(IsNull(M_Col("CUSTID")), "", M_Col("CUSTID")) + "'"
''        M_OBJCONN.Execute (sqlbcknego)
''        sqldelnegoptp = "delete from tblnegoptp where custid='" + IIf(IsNull(M_Col("CUSTID")), "", M_Col("CUSTID")) + "'"
''        M_OBJCONN.Execute (sqldelnegoptp)
''        sqlremarkhstupdate = "update mgm_hst set bckhst=hst where custid='" + IIf(IsNull(M_Col("CUSTID")), "", M_Col("CUSTID")) + "'"
''        M_OBJCONN.Execute (sqlremarkhstupdate)
''        sqlremarkhst = "update mgm_hst set hst='' where custid='" + IIf(IsNull(M_Col("CUSTID")), "", M_Col("CUSTID")) + "'"
''        M_OBJCONN.Execute (sqlremarkhst)
''    End If
'
'    LblDOB.Caption = IIf(IsNull(M_Col("DOB")), "", M_Col("DOB"))
'    lblAddr.Text = IIf(IsNull(M_Col("ADDRNOW")), "", M_Col("ADDRNOW"))
'    lblOfficeAddr.Text = IIf(IsNull(M_Col("ADDRPT")), "", M_Col("ADDRPT"))
'    LblCHAditionalAddr.Text = IIf(IsNull(M_Col("addressadd")), "", M_Col("addressadd"))
'    lblZIP.Caption = IIf(IsNull(M_Col("ZIPNOW")), "", M_Col("ZIPNOW"))
'    lblNoCard.Caption = IIf(IsNull(M_Col("NoCard")), "", M_Col("NoCard"))
'    lblNoPay.Caption = IIf(IsNull(M_Col("NoPay")), "", M_Col("NoPay"))
'    lblPromPA.Value = IIf(IsNull(M_Col("Principal")), "", M_Col("Principal"))
'    lblOpenDate.Value = IIf(IsNull(M_Col("OpenDate")), "", M_Col("OpenDate"))
'     If (lblBD.ValueIsNull) Or (lblOpenDate.ValueIsNull) Then
'        Else
'        hsl = DateDiff("m", M_Col("OpenDate"), M_Col("b_D"))
'        lbllama.Caption = IIf(IsNull(CStr(hsl)), "", CStr(hsl)) + "  Bulan "
'    End If
'    lblLastBill.Value = IIf(IsNull(M_Col("LastBill")), "", M_Col("LastBill"))
'    lblLcAtm.Value = IIf(IsNull(M_Col("LcATMP")), "", M_Col("LcATMP"))
'    lblBrokenPromised.Caption = IIf(IsNull(M_Col("BrokenPromise")), "", M_Col("BrokenPromise"))
'    lblBD.Value = CStr(Format(IIf(IsNull(M_Col("B_D")), "", M_Col("B_D")), "yyyy/mm/dd"))
'    lblLimit.Value = IIf(IsNull(M_Col("Limit")), "", M_Col("Limit"))
'    lblPayDt.Value = IIf(IsNull(M_Col("Pay_Dt")), "", M_Col("Pay_Dt"))
'    lblLastPay.Value = IIf(IsNull(M_Col("LastPay")), "", M_Col("LastPay"))
'    lblTtlPay.Value = IIf(IsNull(M_Col("TtlPay")), "", M_Col("TtlPay"))
'    lblAmount.Value = IIf(IsNull(M_Col("AmountWo")), "", Format(M_Col("AmountWo"), "##.##0"))
'    lblregion.Caption = IIf(IsNull(M_Col("region")), "", M_Col("region"))
'    lblLMED.Caption = IIf(IsNull(M_Col("LMED")), "", M_Col("LMED"))
'    lblLIP.Caption = IIf(IsNull(M_Col("LIP")), "", Format(M_Col("LIP"), "dd/mm/yyyy"))
'    LblInstallment.Value = IIf(IsNull(M_Col("Installment")), "", Format(M_Col("Installment"), "##.##0"))
'    lblrange.Caption = IIf(IsNull(M_Col("RANGE")), "", M_Col("RANGE"))
'    AHome1.Value = IIf(IsNull(M_Col("AHOMENO")), "", M_Col("AHOMENO"))
'    txtHomeNo1.Value = IIf(IsNull(M_Col("HOMENO")), "", M_Col("HOMENO"))
'    If IsNull(M_Col("HOMENO")) = False And M_Col("HOMENO") <> "" Then
'        'txtHomeNo1A.Value = Left(m_col("HOMENO"), Len(m_col("HOMENO")) - 3) & "XXX"
'        txtHomeNo1A.Value = Left(M_Col("HOMENO"), 4) & "BBB" & Mid(M_Col("HOMENO"), 8, 15)
'        CmbPhone.AddItem "HomePhone"
'    End If
'    AHome2.Value = IIf(IsNull(M_Col("AHOMENO2")), "", M_Col("AHOMENO2"))
'    txtHomeNo2.Value = IIf(IsNull(M_Col("HOMENO2")), "", M_Col("HOMENO2"))
'    If IsNull(M_Col("HOMENO2")) = False And M_Col("HOMENO2") <> "" Then
'        'txtHomeNo2A.Value = Left(m_col("HOMENO2"), Len(m_col("HOMENO2")) - 3) & "XXX"
'        txtHomeNo2A.Value = Left(M_Col("HOMENO2"), 4) & "BBB" & Mid(M_Col("HOMENO2"), 8, 15)
'        CmbPhone.AddItem "HomePhone2"
'    End If
'    AOffice1.Value = IIf(IsNull(M_Col("AOFFICENO")), "", M_Col("AOFFICENO"))
'    txtOfficeNo1.Value = IIf(IsNull(M_Col("OFFICENO")), "", M_Col("OFFICENO"))
'    If IsNull(M_Col("OFFICENO")) = False And M_Col("OFFICENO") <> "" Then
'        'txtOfficeNo1A.Value = Left(m_col("OFFICENO"), Len(m_col("OFFICENO")) - 3) & "XXX"
'        txtOfficeNo1A.Value = Left(M_Col("OFFICENO"), 4) & "BBB" & Mid(M_Col("OFFICENO"), 8, 15)
'        CmbPhone.AddItem "OfficePhone"
'    End If
'
'    AOffice2.Value = IIf(IsNull(M_Col("AOFFICENO2")), "", M_Col("AOFFICENO2"))
'    txtOfficeNo2.Value = IIf(IsNull(M_Col("OFFICENO2")), "", M_Col("OFFICENO2"))
'    If IsNull(M_Col("OFFICENO2")) = False And M_Col("OFFICENO2") <> "" Then
'        'txtOfficeNo2A.Value = Left(m_col("OFFICENO2"), Len(m_col("OFFICENO2")) - 3) & "XXX"
'        txtOfficeNo2A.Value = Left(M_Col("OFFICENO2"), 4) & "BBB" & Mid(M_Col("OFFICENO2"), 8, 15)
'        CmbPhone.AddItem "OfficePhone2"
'    End If
'    txtMobileNo1.Value = IIf(IsNull(M_Col("MOBILENO")), "", M_Col("MOBILENO"))
'    If IsNull(M_Col("MOBILENO")) = False And M_Col("MOBILENO") <> "" Then
'        'txtMobileNo1A.Value = Left(m_col("MOBILENO"), Len(m_col("MOBILENO")) - 3) & "XXX"
'        txtMobileNo1A.Value = Left(M_Col("MOBILENO"), 4) & "BBB" & Mid(M_Col("MOBILENO"), 8, 15)
'        CmbPhone.AddItem "Hp"
'    End If
'    txtMobileNo2.Value = IIf(IsNull(M_Col("MOBILENO2")), "", M_Col("MOBILENO2"))
'    If IsNull(M_Col("MOBILENO2")) = False And M_Col("MOBILENO2") <> "" Then
'        'txtMobileNo2A.Value = Left(m_col("MOBILENO2"), Len(m_col("MOBILENO2")) - 3) & "XXX"
'        txtMobileNo2A.Value = Left(M_Col("MOBILENO2"), 4) & "BBB" & Mid(M_Col("MOBILENO2"), 8, 15)
'        CmbPhone.AddItem "Hp2"
'    End If
'    AHomeAdd1(0).Value = IIf(IsNull(M_Col("AHOMENOADD1")), "", M_Col("AHOMENOADD1"))
'    AHomeAdd2(1).Value = IIf(IsNull(M_Col("AHOMENOADD2")), "", M_Col("AHOMENOADD2"))
'    AOfficeAdd(2).Value = IIf(IsNull(M_Col("AOFFICENOADD1")), "", M_Col("AOFFICENOADD1"))
'    AOfficeAdd(3).Value = IIf(IsNull(M_Col("AOFFICENOADD2")), "", M_Col("AOFFICENOADD2"))
'    AFaxAdd(4).Value = IIf(IsNull(M_Col("AFAXNOADD1")), "", M_Col("AFAXNOADD1"))
'    AFaxAdd(5).Value = IIf(IsNull(M_Col("AFAXNOADD2")), "", M_Col("AFAXNOADD2"))
'    txtHomeAdd1.Value = IIf(IsNull(M_Col("HOMENOADD1")), "", M_Col("HOMENOADD1"))
'    If IsNull(M_Col("HOMENOADD1")) = False And M_Col("HOMENOADD1") <> "" Then
'        txtHomeAdd1A.Value = Left(M_Col("HOMENOADD1"), 4) & "BBB" & Mid(M_Col("HOMENOADD1"), 8, 15)
'        CmbPhone.AddItem "AddHome1"
'    Else
'        txtHomeAdd1.Visible = True
'        txtHomeAdd1A.Visible = False
'    End If
'    txtHomeAdd2.Value = IIf(IsNull(M_Col("HOMENOADD2")), "", M_Col("HOMENOADD2"))
'    If IsNull(M_Col("HOMENOADD2")) = False And M_Col("HOMENOADD2") <> "" Then
'        txtHomeAdd2A.Value = Left(M_Col("HOMENOADD2"), 4) & "BBB" & Mid(M_Col("HOMENOADD2"), 8, 15)
'        CmbPhone.AddItem "AddHome2"
'    Else
'        txtHomeAdd2A.Visible = False
'        txtHomeAdd2.Visible = True
'    End If
'    txtOfficeAdd1.Value = IIf(IsNull(M_Col("OFFICENOADD1")), "", M_Col("OFFICENOADD1"))
'    If IsNull(M_Col("OFFICENOADD1")) = False And M_Col("OFFICENOADD1") <> "" Then
'        txtOfficeAdd1A.Value = Left(M_Col("OFFICENOADD1"), 4) & "BBB" & Mid(M_Col("OFFICENOADD1"), 8, 15)
'        CmbPhone.AddItem "AddOffice1"
'    Else
'        txtOfficeAdd1A.Visible = False
'        txtOfficeAdd1.Visible = True
'    End If
'    txtOfficeAdd2.Value = IIf(IsNull(M_Col("OFFICENOADD2")), "", M_Col("OFFICENOADD2"))
'    If IsNull(M_Col("OFFICENOADD2")) = False And M_Col("OFFICENOADD2") <> "" Then
'        txtOfficeAdd2A.Value = Left(M_Col("OFFICENOADD2"), 4) & "BBB" & Mid(M_Col("OFFICENOADD2"), 8, 15)
'        CmbPhone.AddItem "AddOffice2"
'    Else
'        txtOfficeAdd2.Visible = True
'        txtOfficeAdd2A.Visible = False
'    End If
'    txtMobileAdd1.Value = IIf(IsNull(M_Col("MOBILENOADD1")), "", M_Col("MOBILENOADD1"))
'    If IsNull(M_Col("MOBILENOADD1")) = False And M_Col("MOBILENOADD1") <> "" Then
'        txtMobileAdd1A.Value = Left(M_Col("MOBILENOADD1"), 4) & "BBB" & Mid(M_Col("MOBILENOADD1"), 8, 15)
'        CmbPhone.AddItem "AddMobile1"
'    Else
'        txtMobileAdd1.Visible = True
'        txtMobileAdd1A.Visible = False
'    End If
'    txtMobileAdd2.Value = IIf(IsNull(M_Col("MOBILENOADD2")), "", M_Col("MOBILENOADD2"))
'    If IsNull(M_Col("MOBILENOADD2")) = False And M_Col("MOBILENOADD2") <> "" Then
'        txtMobileAdd2A.Value = Left(M_Col("MOBILENOADD2"), 4) & "BBB" & Mid(M_Col("MOBILENOADD2"), 8, 15)
'        CmbPhone.AddItem "AddMobile2"
'    Else
'        txtMobileAdd2.Visible = True
'        txtMobileAdd2A.Visible = False
'    End If
'    txtFaxAdd1.Value = IIf(IsNull(M_Col("FAXNOADD1")), "", M_Col("FAXNOADD1"))
'    txtFaxAdd2.Value = IIf(IsNull(M_Col("FAXNOADD2")), "", M_Col("FAXNOADD2"))
'    AddrNow.Text = IIf(IsNull(M_Col("TxtPtpAddr")), "", M_Col("TxtPtpAddr"))
'    LblLunas.Caption = IIf(IsNull(M_Col!tgllunas), "", "TELAH LUNAS")
'    TxtEC.Text = IIf(IsNull(M_Col!ec_name), "", M_Col!ec_name)
'    txtECno.Value = IIf(IsNull(M_Col!ec_telp), "", M_Col!ec_telp)
'    If IsNull(M_Col("ec_telp")) = False And M_Col("ec_telp") <> "" Then
'        txtECnoA.Value = Left(M_Col("ec_telp"), 4) & "BBB" & Mid(M_Col("ec_telp"), 8, 15)
'        CmbPhone.AddItem "EconPhone"
'    Else
'        txtECnoA.Visible = False
'        txtECno.Visible = True
'    End If
'    cbolastcall.Text = IIf(IsNull(M_Col!statuscall), "", M_Col!statuscall)
''    If cbolastcall.Text = "" Then
''        Call isi_lastcall
''    End If
'
'' cari extension
'    If InStr(1, txtOfficeNo1.Value, "X", vbTextCompare) > 0 Then
'        TxtExt1.Text = Right(txtOfficeNo1.Value, Len(txtOfficeNo1.Value) - InStr(1, txtOfficeNo1.Value, "X", vbTextCompare))
'    End If
'    If InStr(1, txtOfficeNo2.Value, "X", vbTextCompare) > 0 Then
'        TxtExt2.Text = Right(txtOfficeNo2.Value, Len(txtOfficeNo2.Value) - InStr(1, txtOfficeNo2.Value, "X", vbTextCompare))
'    End If
'    If InStr(1, txtOfficeAdd1.Value, "X", vbTextCompare) > 0 Then
'        TxtExt3.Text = Right(txtOfficeAdd1.Value, Len(txtOfficeAdd1.Value) - InStr(1, txtOfficeAdd1.Value, "X", vbTextCompare))
'    End If
'    If InStr(1, txtOfficeAdd2.Value, "X", vbTextCompare) > 0 Then
'        TxtExt4.Text = Right(txtOfficeAdd2.Value, Len(txtOfficeAdd2.Value) - InStr(1, txtOfficeAdd2.Value, "X", vbTextCompare))
'    End If
'
'    If UCase(mdiform1.txtlevel.text) = "AGENT" Then
'        If Len(txtECno.Value) > 2 Then
'            txtECno.ReadOnly = True
'        End If
'        If Len(txtHomeAdd1.Value) > 2 Then
'            txtHomeAdd1.ReadOnly = True
'        End If
'        If Len(txtHomeAdd2.Value) > 2 Then
'            txtHomeAdd2.ReadOnly = True
'        End If
'        If Len(txtOfficeAdd1.Value) > 2 Then
'            txtOfficeAdd1.ReadOnly = True
'        End If
'        If Len(txtOfficeAdd2.Value) > 2 Then
'            txtOfficeAdd2.ReadOnly = True
'        End If
'        If Len(txtMobileAdd1.Value) > 2 Then
'            txtMobileAdd1.ReadOnly = True
'        End If
'        If Len(txtMobileAdd2.Value) > 2 Then
'            txtMobileAdd2.ReadOnly = True
'        End If
'        If Len(txtECno.Value) > 2 Then
'            txtECno.ReadOnly = True
'        End If
'    End If
'    cmbNextAct.Text = IIf(IsNull(M_Col("NEXTACT")), "", M_Col("NEXTACT"))
'
'    sPending = CStr(Trim(IIf(IsNull(M_Col!f_Pending), "", M_Col!f_Pending)))
'     If sPending = "Pending" Then
'         chkAppv(0).Value = 0
'    End If
'
'
''---REMARKS BY RIF
''    Select Case M_Col!RECSTATUS
''        Case "N"
''            C_NotContacted.Value = 1
''            cmbUncontacted.Text = IIf(IsNull(M_Col("KETHSLKERJA")), "", M_Col("KETHSLKERJA"))
''            cmbDescUn.Text = IIf(IsNull(M_Col("KETHSLKERJADESC")), "", M_Col("KETHSLKERJADESC"))
''        Case "C"
''            C_Contacted.Value = 1
''            kontak = True
''            cmbContacted.Text = IIf(IsNull(M_Col("KETHSLKERJA")), "", M_Col("KETHSLKERJA"))
''            cmbDescCon.Text = IIf(IsNull(M_Col("KETHSLKERJADESC")), "", M_Col("KETHSLKERJADESC"))
''     End Select
'        If IIf(IsNull(M_Col!F_CEK), "", Left(M_Col!F_CEK, 3)) = "PTP" Or M_Col!F_CEK = "POP" Or M_Col!F_CEK = "SP-" Then
'            C_Payment.Value = 1
'            TdbPTP.Value = IIf(IsNull(M_Col!TdbDatePTP), "", M_Col!TdbDatePTP)
'            txtPayment.Value = IIf(IsNull(M_Col!ttlptp), 0, M_Col!ttlptp)
'            TxtPayment2.Value = IIf(IsNull(M_Col!ttlptp), 0, M_Col!ttlptp) 'tampilkan di detail payment
'             cmbDiscount.Text = IIf(IsNull(M_Col!discpersen), 0, M_Col!discpersen)
'            CmbBaseOn.Text = IIf(IsNull(M_Col!CmbBaseOn), "", M_Col!CmbBaseOn)
'            'TdbDatePTP.Value = IIf(IsNull(m_col!TGLINCOMING), "", m_col!TGLINCOMING)
'
'        Else
'
'        End If
'    Select Case M_Col!RECSTATUS
'        Case "V"
'            C_VALID.Value = 1
'            cbovalid.Text = IIf(IsNull(M_Col("KETHSLKERJA")), "", M_Col("KETHSLKERJA"))
'            cbodescvalid.Text = IIf(IsNull(M_Col("KETHSLKERJADESC")), "", M_Col("KETHSLKERJADESC"))
'        Case "C"
'            C_Contacted.Value = 1
'            kontak = True
'            cmbContacted.Text = IIf(IsNull(M_Col("KETHSLKERJA")), "", M_Col("KETHSLKERJA"))
'            cmbDescCon.Text = IIf(IsNull(M_Col("KETHSLKERJADESC")), "", M_Col("KETHSLKERJADESC"))
'            If Left(statusptp2, 2) = "OP" Then
'                C_SKIP.Enabled = False
'                FrmSKIP.Enabled = False
'                C_VALID.Enabled = False
'                FrMValid.Enabled = False
'            End If
'         Case "P"
'            C_PTP.Value = 1
'            If mdiform1.txtlevel.text = "Agent" Then
'            C_VALID.Enabled = False
'            C_Contacted.Enabled = False
'            FrMValid.Enabled = False
'            C_SKIP.Enabled = False
'            FrmSKIP.Enabled = False
'            End If
'
'            cboPTP.Text = IIf(IsNull(M_Col("KETHSLKERJA")), "", M_Col("KETHSLKERJA"))
'            'cmbDescCon.Text = IIf(IsNull(M_Col("KETHSLKERJADESC")), "", M_Col("KETHSLKERJADESC"))
'         Case "S"
'            C_SKIP.Value = 1
'            cboskip.Text = IIf(IsNull(M_Col("KETHSLKERJA")), "", M_Col("KETHSLKERJA"))
'            cbodescskip.Text = IIf(IsNull(M_Col("KETHSLKERJADESC")), "", M_Col("KETHSLKERJADESC"))
'         Case "O"
'            C_POPSP.Value = 1
'             If mdiform1.txtlevel.text = "Agent" Then
'               'qt = M_Col!RECSTATUS
'               If IIf(IsNull(M_Col("KETHSLKERJA")), "", Left(M_Col("KETHSLKERJA"), 3)) = "POP" Then
'                    qt = "O"
'                    FrmSKIP.Enabled = False
'                    FrmContacted.Enabled = False
'                    FrMValid.Enabled = False
'                    C_SKIP.Enabled = False
'                    C_Contacted.Enabled = False
'                    C_VALID.Enabled = False
'
'               Else
'                    qt = ""
'               End If
'
'                C_VALID.Enabled = False
'                C_Contacted.Enabled = False
'                FrMValid.Enabled = False
'                C_SKIP.Enabled = False
'                FrmSKIP.Enabled = False
'            End If
'
'            cboPOPSP.Text = IIf(IsNull(M_Col("KETHSLKERJA")), "", M_Col("KETHSLKERJA"))
'            'cmbDescCon.Text = IIf(IsNull(M_Col("KETHSLKERJADESC")), "", M_Col("KETHSLKERJADESC"))
'     End Select
'
'    If mdiform1.txtlevel.text = "Agent" Then
'        If IIf(IsNull(M_Col!RECSTATUS), "", M_Col!RECSTATUS) <> "O" Then
'            frmpopsp.Enabled = False
'            C_POPSP.Enabled = False
'        End If
'    End If
'
'    batchtype = IIf(IsNull(M_Col("batchtype")), "", M_Col("batchtype"))
'    Select Case batchtype
'    Case "A"
'        lbltype.Caption = "AWARENESS"
'    Case "R"
'        lbltype.Caption = "RECOVERY"
'    Case "P"
'        lbltype.Caption = "PIL"
'    End Select
'End If
'
''Set m_cust1 = M_DATA.QUERY_HIST_CUST(M_OBJCONN, "CUSTID = '" + lblCustId.Caption + "'", mdiform1.txtlevel.text)
'Set m_cust1 = M_DATA.QUERY_HIST_CUST(M_OBJCONN, "CUSTID = '" + lblCustId.Caption + "'")
'While Not m_cust1.EOF
'    'Set listitem = ListView1(1).ListItems.ADD(, , Left(m_cust1("TGL"), 4) & "/" & Mid(m_cust1("TGL"), 5, 2) & "/" & IIf(IsNull(m_cust1("TGL")), "", Mid(m_cust1("TGL"), 7, 2)) & " " & IIf(IsNull(m_cust1("TGL")), "", Mid(m_cust1("TGL"), 9, 2)) & ":" & Right(m_cust1("TGL"), 2))
'     Set listitem = listview1(1).ListItems.ADD(, , IIf(IsNull(m_cust1("TGL")), "", m_cust1!TGL))
'        listitem.SubItems(1) = IIf(IsNull(m_cust1("HST")), "", m_cust1("HST"))
'        listitem.SubItems(2) = IIf(IsNull(m_cust1("AGENT")), "", m_cust1("AGENT"))
'        listitem.SubItems(3) = IIf(IsNull(m_cust1("KodeDs")), "", m_cust1("KodeDs"))
'        listitem.SubItems(4) = IIf(IsNull(m_cust1("f_cek")), "", m_cust1("f_cek"))
'm_cust1.MoveNext
'Wend
'If statusptp2 <> "" Then
'Call isi_datapayment
'Call Show_NEGOPTP
'End If
'Call Show_Visit
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
''M_OBJRS.Open "Select custid, sum(payment) as jml from tbllunas where custid = '" + lblCustId.Caption + "' GROUP BY CUSTID", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'CMDSQL = " Select a.custid, sum(a.payment) as jml from tbllunas a inner join mgm b  on a.custid=b.custid "
'CMDSQL = CMDSQL + " where  a.custid = '" + lblCustId.Caption + "'  and date(a.Paydate)+1  > b.tglsource  group by a.custid"
'M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
'While Not M_OBJRS.EOF
'        TxtAfterPay.Value = IIf(IsNull(M_OBJRS("jml")), 0, M_OBJRS("jml"))
'        M_OBJRS.MoveNext
'Wend
' 'hitung sisa hutang
' txtSisaHutang.Value = Val(TxtPayment2.Value) - Val(TxtAfterPay.Value)
'
' '---------->> hitung PRINCIPLE & AMOUNTWO  after pay  <<-----------------
' If TxtAfterPay.Value = 0 Then
'    txtPrinciple_A.Value = 0
'    txtAmountwo_A.Value = 0
'    Else
'  txtPrinciple_A.Value = Val(lblPromPA.Value) - Val(TxtAfterPay.Value)
'  txtAmountwo_A.Value = Val(lblAmount.Value) - Val(TxtAfterPay.Value)
' End If
' If UCase(mdiform1.txtlevel.text) = "AGENT" Then
' 'UNTUK RITPIL, principal di hidden
'    If UCase(Left(mdiform1.txtusername.text, 2)) = "PR" Or UCase(Left(mdiform1.txtusername.text, 1)) = "R" Then
'    lblPromPA.Visible = False
'    Label16.Visible = False
'    txtPrinciple_A.Visible = False
'    Label11(8).Visible = False
'    End If
' End If
'
'
'
''Set m_cust = Nothing
' Set M_OBJRS = Nothing
'
'Exit Sub
'HELL:
'   MsgBox Err.Description
' Resume
' Set M_OBJRS = Nothing
''Set m_cust = Nothing
'
'
'End Sub
'
'Private Sub isi_datapayment()
'Dim m_cust2 As New ADODB.Recordset
'Dim NilaiAfterPay As Currency
'Dim M_DATA As New CLS_FRMCUST_CC
'Set m_cust2 = M_DATA.QUERY_HIST_PAID(M_OBJCONN, "a.custid = '" + lblCustId.Caption + "' ")
'listview1(0).ListItems.CLEAR
'While Not m_cust2.EOF
'    Set listitem = listview1(0).ListItems.ADD(, , IIf(IsNull(m_cust2("Paydate")), "", m_cust2("Paydate")))
'        listitem.SubItems(1) = IIf(IsNull(m_cust2("payment")), "0", Format(m_cust2("Payment"), "##,###"))
'        listitem.SubItems(2) = IIf(IsNull(m_cust2("AGENT")), "", m_cust2("AGENT"))
'        listitem.SubItems(3) = IIf(IsNull(m_cust2("FieldName")), "", m_cust2("FieldName"))
'        listitem.SubItems(4) = IIf(IsNull(m_cust2("Id")), "0", m_cust2("Id"))
'        NilaiAfterPay = NilaiAfterPay + IIf(IsNull(m_cust2("payment")), "0", m_cust2("Payment"))
'    m_cust2.MoveNext
'Wend
'Set m_cust2 = Nothing
'TxtAfterPay.Value = NilaiAfterPay
'txtSisaHutang.Value = Format(TxtPayment2.Value - TxtAfterPay.Value, "##,###")
'End Sub
'Private Sub Show_Visit()
'Dim m_cust2 As New ADODB.Recordset
'Dim m_Visit As New ClsVisit
'Dim Jml As String
'Dim CMDSQL As String
'Set m_cust2 = New ADODB.Recordset
'CMDSQL = "SELECT requestdate,visitdate,detailsR,detailsV,visitke,VisitNo,id,F_CEK FROM tblVisit where custid='" + lblCustId.Caption + "'"
'm_cust2.CursorLocation = adUseClient
'm_cust2.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
''Set m_cust2 = m_Visit.SELECT_RequestVisit(M_OBJCONN, lblCustId.Caption)
'LstVisit.ListItems.CLEAR
'While Not m_cust2.EOF
'    Set listitem = LstVisit.ListItems.ADD(, , IIf(IsNull(m_cust2!RequestDate), "", m_cust2!RequestDate))
'        listitem.SubItems(1) = IIf(IsNull(m_cust2!VisitDate), "", m_cust2!VisitDate)
'        listitem.SubItems(2) = Trim(IIf(IsNull(m_cust2!VisitNo), "", m_cust2!VisitNo))
'        listitem.SubItems(3) = IIf(IsNull(m_cust2!DetailsR), "", m_cust2!DetailsR)
'        listitem.SubItems(4) = IIf(IsNull(m_cust2!DetailsV), "", m_cust2!DetailsV)
'        listitem.SubItems(5) = IIf(IsNull(m_cust2!VisitKe), "0", m_cust2!VisitKe)
'        listitem.SubItems(6) = IIf(IsNull(m_cust2!ID), "0", m_cust2!ID)
'        listitem.SubItems(7) = IIf(IsNull(m_cust2!F_CEK), "0", m_cust2!F_CEK)
'        m_cust2.MoveNext
'Wend
'Jml = m_cust2.RecordCount + 1
'TDBNumber1.Value = Jml
''Select Case Jml
''Case "0"
''Combo1.Text = "I"
''Case "1"
''Combo1.Text = "II"
''Case "2"
''Combo1.Text = "III"
''Case "3"
''Combo1.Text = "IV"
''Case "4"
''Combo1.Text = "V"
''Case "5"
''Combo1.Text = "VI"
''End Select
'Set m_cust2 = Nothing
'
'End Sub
'
'
'Private Sub CEK_UPDATE_PELANGGAN()
'Dim M_DATA As New CLS_FRMCUST_CC_MGM
'Dim m_Visit As New ClsVisit
'Dim pStatusHstLstCall As String
'Dim statusptp As String
'Dim TGLCALL, TGLSTATUS As Date
'On Error GoTo editErr
'
''       M_OBJCONN.BeginTrans
''Set M_Col = New ADODB.Recordset
''M_Col.CursorLocation = adUseClient
''   M_Col.Open "Select * from mgm where custid='" & lblCustId.Caption & "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
'        'ADDITIONAL PHONE
'
'        M_Col("AHOMENOADD1") = AHomeAdd1(0).Value
'        M_Col("AHOMENOADD2") = AHomeAdd2(1).Value
'        M_Col("AOFFICENOADD1") = AOfficeAdd(2).Value
'        M_Col("AOFFICENOADD2") = AOfficeAdd(3).Value
'        M_Col("AFAXNOADD1") = AFaxAdd(4).Value
'        M_Col("AFAXNOADD2") = AFaxAdd(5).Value
'        If txtHomeAdd1A.Value = "" And txtHomeAdd1A.Visible = True Then
'            M_Col("HOMENOADD1") = txtHomeAdd1A.Value
'        ElseIf txtHomeAdd1.Value <> "" And txtHomeAdd1.Visible = True Then
'            M_Col("HOMENOADD1") = txtHomeAdd1.Value
'        End If
'
'        If txtHomeAdd2A.Value = "" And txtHomeAdd2A.Visible = True Then
'            M_Col("HOMENOADD2") = txtHomeAdd2A.Value
'        ElseIf txtHomeAdd2.Value <> "" And txtHomeAdd2.Visible = True Then
'            M_Col("HOMENOADD2") = txtHomeAdd2.Value
'        End If
'
'        If txtOfficeAdd1A.Value = "" And txtOfficeAdd1A.Visible = True Then
'            M_Col("OFFICENOADD1") = txtOfficeAdd1A.Value
'        ElseIf txtOfficeAdd1.Value <> "" And txtOfficeAdd1.Visible = True Then
'            M_Col("OFFICENOADD1") = txtOfficeAdd1.Value
'        End If
'
'        If txtOfficeAdd2A.Value = "" And txtOfficeAdd2A.Visible = True Then
'            M_Col("OFFICENOADD2") = txtOfficeAdd2A.Value
'        ElseIf txtOfficeAdd2.Value <> "" And txtOfficeAdd2.Visible = True Then
'            M_Col("OFFICENOADD2") = txtOfficeAdd2.Value
'        End If
'
'        If txtMobileAdd1A.Value = "" And txtMobileAdd1A.Visible = True Then
'            M_Col("MOBILENOADD1") = txtMobileAdd1A.Value
'        ElseIf txtMobileAdd1.Value <> "" And txtMobileAdd1.Visible = True Then
'            M_Col("MOBILENOADD1") = txtMobileAdd1.Value
'        End If
'
'        If txtMobileAdd2A.Value = "" And txtMobileAdd2A.Visible = True Then
'            M_Col("MOBILENOADD2") = txtMobileAdd2A.Value
'        ElseIf txtMobileAdd2.Value <> "" And txtMobileAdd2.Visible = True Then
'            M_Col("MOBILENOADD2") = txtMobileAdd2.Value
'        End If
'
'        M_Col("FAXNOADD1") = txtFaxAdd1.Value
'        M_Col("FAXNOADD2") = txtFaxAdd2.Value
'        M_Col!TxtPtpAddr = AddrNow.Text
'        M_Col!ec_name = TxtEC.Text
'
'        If txtECnoA.Value = "" And txtECnoA.Visible = True Then
'            M_Col("ec_telp") = txtECnoA.Value
'        ElseIf txtECno.Value <> "" And txtECno.Visible = True Then
'            M_Col!ec_telp = txtECno.Value
'        End If
'
'        If UCase(mdiform1.txtlevel.text) = "AGENT" Then
'            If Len(txtECno.Value) > 2 Then
'                txtECno.ReadOnly = True
'            End If
'            If Len(txtHomeAdd1.Value) > 2 Then
'                txtHomeAdd1.ReadOnly = True
'            End If
'            If Len(txtHomeAdd2.Value) > 2 Then
'                txtHomeAdd2.ReadOnly = True
'            End If
'            If Len(txtOfficeAdd1.Value) > 2 Then
'                txtOfficeAdd1.ReadOnly = True
'            End If
'            If Len(txtOfficeAdd2.Value) > 2 Then
'                txtOfficeAdd2.ReadOnly = True
'            End If
'            If Len(txtMobileAdd1.Value) > 2 Then
'                txtMobileAdd1.ReadOnly = True
'            End If
'            If Len(txtMobileAdd2.Value) > 2 Then
'                txtMobileAdd2.ReadOnly = True
'            End If
'        End If
'
''    m_col!f_payment = "PAYMENT"
''    End If
'
'
''        m_col("PRIOR") = cmbPrior.Text
''        m_col("ADDRPT") = lblOfficeAddr.Text
''        m_col("AHOMENO") = AHome1.Value
''        m_col("AHOMENO2") = AHome2.Value
''        m_col("AOFFICENO") = AOffice1.Value
''        m_col("AOFFICENO2") = AOffice2.Value
'        M_Col("TGLCALL") = Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") & " " & Format(Now, "hh:nn")
''        If Len(IIf(IsNull(m_col!HOMENO), "", m_col!HOMENO)) > 2 Then
''            txtHomeNo1.ReadOnly = True
''        End If
''        m_col("HOMENO2") = txtHomeNo2.Value
''        If Len(IIf(IsNull(m_col!HOMENO2), "", m_col!HOMENO2)) > 2 Then
''            txtHomeNo2.ReadOnly = True
''        End If
''        m_col("MOBILENO") = txtMobileNo1.Value
''        If Len(IIf(IsNull(m_col!MOBILENO), "", m_col!MOBILENO)) > 2 Then
''            txtMobileNo1.ReadOnly = True
''        End If
''        m_col("MOBILENO2") = txtMobileNo2.Value
''        If Len(IIf(IsNull(m_col!MOBILENO2), "", m_col!MOBILENO2)) > 2 Then
''            txtMobileNo2.ReadOnly = True
''        End If
'
''        m_col("OFFICENO") = txtOfficeNo1.Value
''        If Len(IIf(IsNull(m_col!OFFICENO), "", m_col!OFFICENO)) > 2 Then
''            txtOfficeNo1.ReadOnly = True
''        End If
''        m_col("OFFICENO2") = txtOfficeNo2.Value
''        If Len(IIf(IsNull(m_col!OFFICENO2), "", m_col!OFFICENO2)) > 2 Then
''            txtOfficeNo2.ReadOnly = True
'
''         If Len(IIf(IsNull(m_col!HOMENO), "", m_col!HOMENO)) > 2 Then
''            txtHomeNo1.ReadOnly = True
''        End If
''        End If
'        'sebelum f_cek diubah statusnya
'        statusptp = IIf(IsNull(M_Col!F_CEK), "", M_Col!F_CEK)
''        If chkAppv(0).Value Then
''            m_col("F_Pending") = "OK"
''        End If
'
''---REMARK BY RIF
''        If C_Contacted.Value Then
''            M_Col("RECSTATUS") = "C"
''               pStatusLstCall = cmbContacted.Text
''               txtResult.Text = pStatusLstCall
''               pStatusLstCalldesc = cmbDescCon.Text
''               txtResultDesc.Text = pStatusLstCalldesc
''               M_Col!F_CEK = Left(cmbContacted.Text, 3) & Left(cmbDescCon.Text, 1)
''            Else
''                If C_NotContacted.Value Then
''                    M_Col("RECSTATUS") = "N"
''                    pStatusLstCall = cmbUncontacted.Text
''                    txtResult.Text = pStatusLstCall
''                    pStatusLstCalldesc = cmbDescUn.Text
''                    txtResultDesc.Text = pStatusLstCalldesc
''                    M_Col!f_Pending = "Pending"
''                    If Left(cmbUncontacted.Text, 3) = "NBP" Then
''                    M_Col!F_CEK = "NBP"
''                    ElseIf Left(cmbUncontacted.Text, 2) = "NA" Then
''                    M_Col!F_CEK = Left(cmbUncontacted.Text, 3) & Left(cmbDescUn.Text, 1)
''                    Else
''                    M_Col!F_CEK = Left(cmbUncontacted.Text, 3) & Left(cmbDescUn.Text, 2)
''                End If
''                Else
''                    M_Col!F_CEK = ""
''                End If
''        End If
'
'            If C_VALID.Value Then
'                M_Col("RECSTATUS") = "V"
'               pStatusLstCall = cbovalid.Text
'               txtResult.Text = pStatusLstCall
'               pStatusLstCalldesc = cbodescvalid.Text
'               txtResultDesc.Text = pStatusLstCalldesc
'                 If Left(cbovalid.Text, 3) = "NBP" Then
'                    M_Col!F_CEK = "NBP"
'                 ElseIf Left(cbovalid.Text, 2) = "NA" Then
'                    M_Col!F_CEK = Left(cbovalid.Text, 3) & Left(cbodescvalid.Text, 1)
'                End If
'            Else
'                If C_Contacted.Value Then
'                    M_Col("RECSTATUS") = "C"
'                    pStatusLstCall = cmbContacted.Text
'                    txtResult.Text = pStatusLstCall
'                    pStatusLstCalldesc = cmbDescCon.Text
'                    txtResultDesc.Text = pStatusLstCalldesc
'
''                    txtResult.Text = pStatusLstCall
''                    pStatusLstCalldesc = cmbDescUn.Text
''                    txtResultDesc.Text = pStatusLstCalldesc
''                    M_Col!f_Pending = "Pending"
''                    If Left(cmbUncontacted.Text, 3) = "NBP" Then
''                    M_Col!F_CEK = "NBP"
''                    ElseIf Left(cmbUncontacted.Text, 2) = "NA" Then
''                    M_Col!F_CEK = Left(cmbUncontacted.Text, 3) & Left(cmbDescUn.Text, 1)
''                    Else
''                    M_Col!F_CEK = Left(cmbUncontacted.Text, 3) & Left(cmbDescUn.Text, 2)
''                    End If
'                    M_Col!F_CEK = Left(cmbContacted.Text, 3) & Left(cmbDescCon.Text, 1)
'                Else
'                    If C_PTP.Value Then
'                        pStatusLstCall = cboPTP.Text
'                        txtResult.Text = pStatusLstCall
'                        'pStatusLstCalldesc = cbodesc.Text
'                        txtResultDesc.Text = pStatusLstCalldesc
'                        M_Col("RECSTATUS") = "P"
'                        M_Col!F_CEK = Left(cboPTP.Text, 6)
'                    Else
'                        If C_SKIP.Value Then
'                            pStatusLstCall = cboskip.Text
'                            txtResult.Text = pStatusLstCall
'                            pStatusLstCalldesc = cbodescskip.Text
'                            txtResultDesc.Text = pStatusLstCalldesc
'                            M_Col("RECSTATUS") = "S"
'                            M_Col!F_CEK = Left(cboskip.Text, 3) & Left(cbodescskip.Text, 2)
'                        Else
'                             If C_POPSP.Value Then
'                                pStatusLstCall = cboPOPSP.Text
'                                txtResult.Text = pStatusLstCall
'                                'pStatusLstCalldesc = cbodescskip.Text
'                                txtResultDesc.Text = pStatusLstCalldesc
'                                M_Col("RECSTATUS") = "O"
'                                M_Col!F_CEK = Left(cboPOPSP.Text, 3)
'                             Else
'                                M_Col!F_CEK = ""
'                             End If
'                        End If
'                    End If
'                End If
'            End If
'
'        If C_Payment.Value Then
'            If statusptp <> Empty Then
'                If statusptp = M_Col!F_CEK Then
'
'                Else
'                    M_Col!TGLINCOMING = Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") & " " & Format(Now, "hh:nn")
'                End If
'            End If
'            M_Col!ttlptp = txtPayment.Value
'            M_Col!discpersen = cmbDiscount.Text
'            M_Col!CmbBaseOn = CmbBaseOn.Text
'            M_Col!TdbDatePTP = Format(TdbPTP.Value, "yyyy/mm/dd")
'            'm_col!TxtPtpAddr = TxtPtpAddr.Text
'           ' m_col!TxtPhonePTP = TxtPhonePTP.Text
'        Else
'            'm_col!TGLINCOMING = Null
'            M_Col!ttlptp = 0
'            M_Col!discpersen = 0
'        End If
'
''        If C_lunas.Value Then
''            m_col!TglLunas = Format(TdbLunas.Value, "yyyy/mm/dd")
''            m_col!TotLunas = TDBTot_payment.Value
''            m_col!fieldName = TxtFieldName.Text
''        Else
''            m_col!TglLunas = Null
''            m_col!TotLunas = 0
''            m_col!fieldName = Null
''
''        End If
'
'        If Trim(UCase(IIf(IsNull(M_Col("KETHSLKERJA")), "", M_Col("KETHSLKERJA")))) = Trim(UCase(pStatusLstCall)) Then
'            TGLSTATUS = IIf(IsNull(M_Col("TGLSTATUS")), "", Format(M_Col("TGLSTATUS"), "yyyy/mm/dd"))
'        Else
'            M_Col("KETHSLKERJA") = pStatusLstCall
'            M_Col("TGLSTATUS") = Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") & " " & Format(Now, "hh:nn")
'            TGLSTATUS = Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd")
'        End If
'        pStatusHstLstCall = IIf(IsNull(M_Col("KETHSLKERJA")), "", M_Col("KETHSLKERJA"))
'
'        M_Col("KETHSLKERJADESC") = txtResultDesc.Text
'        M_Col("PRIOR") = cmbPrior.Text
'        M_Col("NEXTACT") = cmbNextAct.Text
'        M_Col("REMARKS") = txtRemarks.Text
'        M_Col!NEXTACTDATE = Format(cmbDateSch.Value, "yyyy/mm/dd") & " " & Format(Now, "hh:nn")
'        M_Col("Statuscall") = cbolastcall.Text
'    M_Col.update
'
''M_DATA.UPDATE_CUSTOMER_BARU M_OBJCONN, KETHSLKERJA, STATUS_FIELD_LAMA, M_CALL, M_STATUS, DOK1
'If C_NotContacted.Value = 1 Then
'    If txtRemarks.Text <> Empty Or cmbNextAct.Text <> Empty Then
'        M_DATA.ADD_HISTORY lblCustId.Caption, MDIForm1.TDBDate1.Text, Time, VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11), "COLLECTION", txtRemarks.Text, txtResult.Text, cmbNextAct.Text, "", CStr(IIf(IsNull(M_Col!F_CEK), "", M_Col!F_CEK)), cbolastcall.Text
'    End If
'ElseIf C_Contacted.Value = 1 Then
'    If txtRemarks.Text <> Empty Or cmbNextAct.Text <> Empty Then
'            M_DATA.ADD_HISTORY lblCustId.Caption, MDIForm1.TDBDate1.Text, Time, VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11), "COLLECTION", txtRemarks.Text, txtResult.Text, cmbNextAct.Text, "", CStr(IIf(IsNull(M_Col!F_CEK), "", M_Col!F_CEK)), cbolastcall.Text
'    End If
'ElseIf C_VALID.Value = 1 Then
'    If txtRemarks.Text <> Empty Or cmbNextAct.Text <> Empty Then
'            M_DATA.ADD_HISTORY lblCustId.Caption, MDIForm1.TDBDate1.Text, Time, VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11), "COLLECTION", txtRemarks.Text, txtResult.Text, cmbNextAct.Text, "", CStr(IIf(IsNull(M_Col!F_CEK), "", M_Col!F_CEK)), cbolastcall.Text
'    End If
'ElseIf C_SKIP.Value = 1 Then
'    If txtRemarks.Text <> Empty Or cmbNextAct.Text <> Empty Then
'            M_DATA.ADD_HISTORY lblCustId.Caption, MDIForm1.TDBDate1.Text, Time, VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11), "COLLECTION", txtRemarks.Text, txtResult.Text, cmbNextAct.Text, "", CStr(IIf(IsNull(M_Col!F_CEK), "", M_Col!F_CEK)), cbolastcall.Text
'    End If
'ElseIf C_PTP.Value = 1 Then
'    If txtRemarks.Text <> Empty Or cmbNextAct.Text <> Empty Then
'            M_DATA.ADD_HISTORY lblCustId.Caption, MDIForm1.TDBDate1.Text, Time, VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11), "COLLECTION", txtRemarks.Text, txtResult.Text, cmbNextAct.Text, "", CStr(IIf(IsNull(M_Col!F_CEK), "", M_Col!F_CEK)), cbolastcall.Text
'    End If
'ElseIf C_POPSP.Value = 1 Then
'    If txtRemarks.Text <> Empty Or cmbNextAct.Text <> Empty Then
'            M_DATA.ADD_HISTORY lblCustId.Caption, MDIForm1.TDBDate1.Text, Time, VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11), "COLLECTION", txtRemarks.Text, txtResult.Text, cmbNextAct.Text, "", CStr(IIf(IsNull(M_Col!F_CEK), "", M_Col!F_CEK)), cbolastcall.Text
'    End If
'End If
'    If Len(TDBTot_payment) > 2 Then
'    M_DATA.ADD_tbllunas M_OBJCONN, lblCustId.Caption, Format(TdbLunas.Value, "yyyy/mm/dd"), CCur(TDBTot_payment.Value), VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11), TxtFieldName.Text, ""
'    Else
'    On Error Resume Next
'    End If
'    '------------>> simpan ke table Visit <<--------------------
'   If Option8(0).Value Then
'    m_Visit.ADD_RequestVisit M_OBJCONN, lblCustId.Caption, M_Col!F_CEK, Text1.Text, TDBDate1.Value, TXtDetails.Text, TDBNumber1.Value, TxtAddress.Text, Trim(UCase(VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11)))
'
'   Else
'    On Error Resume Next
'   End If
''End If
''M_OBJCONN.CommitTrans
'MsgBox "Data Sudah Tersimpan", vbInformation + vbOKOnly, "Sukses"
'kontak = False
'
'If shedulePTP_Show = True Then
'  '  MDIForm1.LstGrade.ListItems.Remove MDIForm1.LstGrade.SelectedItem.Index
'Else
'    VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(5) = Format(cmbDateSch.Value, "dd/mm/yyyy") & " " & Format(Now, "hh:nn")
'    VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(6) = cmbNextAct.Text
'    VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(7) = txtRemarks.Text
'    VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(8) = pStatusLstCall
'    VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(9) = cbolastcall.Text
'    'VIEW_mgmDATA.LstVwSearchmgm.SelectedItem.SubItems(17) = Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd")
'    VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(17) = TGLSTATUS
'    VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(18) = Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd")
'    VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(19) = pStatusHstLstCall
'End If
'pStatusLstCall = ""
'pStatusHstLstCall = ""
'txtRemarks.Text = Empty
''cmbNextAct.Text = Empty
''Unload Me
'Set M_DATA = Nothing
'Exit Sub
'editErr:
''    M_OBJCONN.RollbackTrans
'    MsgBox Err.Description
'  Resume
'End Sub
'Private Sub HEADER_HISTORY()
'    listview1(1).ColumnHeaders.ADD 1, , "Tanggal Jam", 26 * TXT
'    listview1(1).ColumnHeaders.ADD 2, , "History", 30 * TXT
'    listview1(1).ColumnHeaders.ADD 3, , "Agent", 10 * TXT
'    listview1(1).ColumnHeaders.ADD 4, , "Sts Call", 10 * TXT
'    listview1(1).ColumnHeaders.ADD 5, , "Sts Call1", 20 * TXT
'End Sub
'Private Sub HEADER_RequestVisit()
'    LstVisit.ColumnHeaders.ADD 1, , "RequestDate", 10 * TXT
'    LstVisit.ColumnHeaders.ADD 2, , "VisitDate", 10 * TXT
'    LstVisit.ColumnHeaders.ADD 3, , "VisitNo", 10 * TXT
'    LstVisit.ColumnHeaders.ADD 4, , "Details", 20 * TXT
'    LstVisit.ColumnHeaders.ADD 5, , "Hasil Visit", 20 * TXT
'    LstVisit.ColumnHeaders.ADD 6, , "VisitKe", 2 * TXT
'    LstVisit.ColumnHeaders.ADD 7, , "ID", 1 * TXT
'    LstVisit.ColumnHeaders.ADD 8, , "Status", 1 * TXT
'    End Sub
'Private Sub HEADER_HISTORY_PAID()
'    listview1(0).ColumnHeaders.ADD 1, , "PayDate", 15 * TXT
'    listview1(0).ColumnHeaders.ADD 2, , "Payment", 30 * TXT
'    listview1(0).ColumnHeaders.ADD 3, , "Agent", 10 * TXT
'    listview1(0).ColumnHeaders.ADD 4, , "FieldName", 30 * TXT
'    listview1(0).ColumnHeaders.ADD 5, , "Id", 30 * TXT
'End Sub
'Private Function CEK_DATA_VALID() As Boolean
'Dim m_msgbox As Variant
'If TDBTot_payment > 2 Then
'CEK_DATA_VALID = True
'Exit Function
'Else
''If mdiform1.txtlevel.text = "TeamLeader" Or mdiform1.txtlevel.text = "Administrator" And (chkAppv(0).Value = 1 Or chkAppv(1).Value = 1) Then
'If (chkAppv(0).Value = 1 Or chkAppv(1).Value = 1) Then
'        Call UpdateAppv
'        'VIEW_mgmDATA.LstVwSearchmgm.SelectedItem.SubItems(8) = VIEW_mgmDATA.LstVwSearchmgm.SelectedItem.SubItems(8) & "Pending"
'        Exit Function
''Else
''   CEK_DATA_VALID = False
''End If
'Else
'    If cbolastcall.Text = "" Then
'            MsgBox "Status Last Call harus diisi", vbInformation + vbOKOnly, "Telegrandi"
'            CEK_DATA_VALID = False
'            Exit Function
'    End If
'    If LstPayment.ListItems.Count = 0 And Left(cboPTP.Text, 3) = "PTP" Then
'        MsgBox "Status PTP harus mengisi Tabel PTP yang Hijau !"
'        CEK_DATA_VALID = False
'        Exit Function
'    End If
'
'    If Left(cmbContacted.Text, 2) = "RP" Or Left(cmbContacted.Text, 2) = "NA" Then
'        If cmbDescCon.Text = "" Then
'            CEK_DATA_VALID = False
'            MsgBox "Description Contacted Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
'            SSTab1.Tab = 3
'            Exit Function
'        End If
'      End If
'
'      If cbovalid.Text <> "" Then
'        If cbodescvalid.Text = "" Then
'            CEK_DATA_VALID = False
'            MsgBox "Description Valid Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
'            SSTab1.Tab = 3
'            Exit Function
'        End If
'     End If
'
'    If cboskip.Text <> "" Then
'        If cbodescskip.Text = "" Then
'            CEK_DATA_VALID = False
'            MsgBox "Description SKIP Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
'            SSTab1.Tab = 3
'            Exit Function
'        End If
'     End If
'
'
'    If C_SKIP.Value = 1 Then
'     If cboskip.Text = Empty Then
'      CEK_DATA_VALID = False
'      MsgBox "Description Skip Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
'        Exit Function
'        SSTab1.Tab = 3
'     End If
'     End If
'
'
'    If C_POPSP.Value = 1 Then
'     If cboPOPSP.Text = Empty Then
'      CEK_DATA_VALID = False
'      MsgBox "Description POP Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
'        Exit Function
'        SSTab1.Tab = 3
'     End If
'     End If
'
'
'     If C_VALID.Value = 1 Then
'     If cbovalid.Text = Empty Then
'      CEK_DATA_VALID = False
'      MsgBox "Description Valid Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
'        Exit Function
'        SSTab1.Tab = 3
'     End If
'     End If
'
'
'     If C_PTP.Value = 1 Then
'        If cboPTP.Text = Empty Then
'            CEK_DATA_VALID = False
'            MsgBox "Description PTP Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
'            Exit Function
'            SSTab1.Tab = 3
'     End If
'     End If
'
'      If C_Contacted.Value = 1 Then
'      If cmbContacted.Text = Empty Then
'      CEK_DATA_VALID = False
'            MsgBox "Contacted Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
'        SSTab1.Tab = 3
'        Exit Function
'      End If
'      End If
''      If C_Payment.Value = 1 Then
''      If TdbDatePTP.Text = "__/__/____" Then
''      CEK_DATA_VALID = False
''      MsgBox "Tanggal PTP Harus Diisi", vbCritical + vbOKOnly, "Peringatan"
''      SSTab1.Tab = 3
''      'TdbDatePTP.SetFocus
''      Exit Function
''      End If
'
'
''    If (CmbContacted.Text) = "" And C_NotContacted.Value = 0 Then
''            CEK_DATA_VALID = False
''            MsgBox "Contacted Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
''            SSTab1.Tab = 3
''            Exit Function
''      End If
'
'    If Left(cmbUncontacted.Text, 2) <> "" Then
'        If cmbDescUn.Text = "" Then
'            CEK_DATA_VALID = False
'            MsgBox "Description UnContacted Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
'            SSTab1.Tab = 3
'            Exit Function
'       End If
'
'
'
'    End If
'      If C_NotContacted.Value = 1 Then
'        If cmbUncontacted.Text = Empty Then
'            CEK_DATA_VALID = False
'            MsgBox "Not Contacted Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
'            SSTab1.Tab = 3
'            Exit Function
'        Else
'                  If cmbDescUn.Text = Empty Then
'                     MsgBox "Not Contacted Description harus diisi", vbCritical + vbOKOnly, "Peringatan"
'                     Exit Function
'                  End If
'                  If txtRemarks.Text = "" And cmbNextAct.Text = "" Then
'                       CEK_DATA_VALID = False
'                        MsgBox "Remarks Atau Next Action Harus Diisi...!!!", vbCritical + vbOKOnly, "Peringatan"
'                        SSTab1.Tab = 3
'                        Exit Function
'                  End If
'        End If
'     End If
'
'  If C_Contacted.Value = 0 And C_VALID.Value = 0 And C_PTP.Value = 0 And C_SKIP.Value = 0 And C_POPSP.Value = 0 Then
'     CEK_DATA_VALID = False
'     MsgBox "Status Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
'     SSTab1.Tab = 3
'     Exit Function
'  End If
'
'    If ADD_CUST = True Then
'    Else
'        If C_Contacted.Value = 1 Or C_VALID.Value = 1 Or C_PTP.Value = 1 Or C_SKIP.Value = 1 Or C_POPSP.Value = 1 Then
'            If cmbDateSch.ValueIsNull = True Or cmbTimeSch.ValueIsNull = True Then
'                CEK_DATA_VALID = False
'                MsgBox "Tanggal Schedule Harus Di isi", vbCritical + vbOKOnly, "Peringatan"
'                SSTab1.Tab = 3
'                Exit Function
'            End If
'            If txtRemarks.Text = "" And cmbNextAct.Text = "" Then
'                CEK_DATA_VALID = False
'                MsgBox "Remarks Atau Next Action Harus Diisi...!!!", vbCritical + vbOKOnly, "Peringatan"
'                SSTab1.Tab = 3
'                Exit Function
'            End If
'            If C_Contacted.Value = 1 Then
'                If cmbDescCon.Text = "" Then
'                    txtRemarks.Text = cmbContacted & " - " & cbolastcall.Text & " - " & txtRemarks.Text
'                Else
'                    txtRemarks.Text = cmbContacted & " - " & cmbDescCon & " - " & cbolastcall.Text & " - " & txtRemarks.Text
'                End If
'            End If
'            If C_VALID.Value = 1 Then
'                If cbodescvalid.Text = "" Then
'                    txtRemarks.Text = cbovalid & " - " & cbolastcall.Text & " - " & txtRemarks.Text
'                Else
'                    txtRemarks.Text = cbovalid & " - " & cbodescvalid & " - " & cbolastcall.Text & " - " & txtRemarks.Text
'                End If
'            End If
'            If C_PTP.Value = 1 Then
'                    txtRemarks.Text = cboPTP & " - " & cbolastcall.Text & " - " & txtRemarks.Text
'            End If
'            If C_SKIP.Value = 1 Then
'                If cbodescskip.Text = "" Then
'                    txtRemarks.Text = cboskip & " - " & cbolastcall.Text & " - " & txtRemarks.Text
'                Else
'                    txtRemarks.Text = cboskip & " - " & cbodescskip & " - " & cbolastcall.Text & " - " & txtRemarks.Text
'                End If
'            End If
'            If C_POPSP.Value = 1 Then
'                    txtRemarks.Text = cboPOPSP & " - " & cbolastcall.Text & " - " & txtRemarks.Text
'            End If
'        End If
''        If stscall = True Then
''            If C_NotContacted.Value = 0 And C_Contacted.Value = 0 Then
''                        CEK_DATA_VALID = False
''                        MsgBox "Status Harus Diisi", vbCritical + vbOKOnly, "Peringatan"
''                        SSTab1.Tab = 3
''                        Exit Function
''            End If
''        End If
''            If C_NotContacted.Value = 1 Then
''                txtRemarks.Text = cmbUncontacted & " - " & cmbDescUn & " - " & cbolastcall.Text & " - " & txtRemarks.Text
''            End If
''    End If
'    If C_Payment.Value = 1 Then
'        If CmbBaseOn.Text = "" Then
'            MsgBox "Base On harus diisi", vbInformation + vbOKOnly, "Telegrandi"
'            CEK_DATA_VALID = False
'            Exit Function
'        End If
'        If cmbDiscount.Text = "" Then
'            MsgBox "Diskon harus diisi", vbInformation + vbOKOnly, "Telegrandi"
'            CEK_DATA_VALID = False
'            Exit Function
'        End If
'        If TdbPTP.ValueIsNull Then
'            MsgBox "Tanggal PTP harus diisi", vbInformation + vbOKOnly, "Telegrandi"
'            CEK_DATA_VALID = False
'            Exit Function
'        End If
'    End If
'End If
'End If
'End If
''cek valid uncontacted pending
'
'CEK_DATA_VALID = True
'End Function
'
'Public Sub Custid_Double()
'Dim listitem As listitem
'Dim M_COL1 As ADODB.Recordset
'Set M_COL1 = New ADODB.Recordset
'M_COL1.CursorLocation = adUseClient
'M_COL1.Open "Select * from mgm where KTPNO='" & lblID.Caption & "' and CUSTID <> '" + lblCustId.Caption + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'While Not M_COL1.EOF
'    Set listitem = LstDoubleId.ListItems.ADD(, , IIf(IsNull(M_COL1("CUSTID")), "", M_COL1("CUSTID")))
'        listitem.SubItems(1) = IIf(IsNull(M_COL1("NAME")), "", M_COL1("NAME"))
'        listitem.SubItems(2) = IIf(IsNull(M_COL1("AGENT")), "", M_COL1("AGENT")) '
'        listitem.SubItems(3) = Format(IIf(IsNull(M_COL1("AMOUNTWO")), "0", M_COL1("AMOUNTWO")), "##,###")
'        listitem.SubItems(4) = Format(IIf(IsNull(M_COL1("PRINCIPAL")), "0", M_COL1("PRINCIPAL")), "##,###")
'    M_COL1.MoveNext
'Wend
''Set m_col = Nothing
'End Sub
'
'Private Sub SSCommand2_Click(Index As Integer)
'Dim m_msgbox As Variant
'Dim STATUS As String
'Dim rscek As New ADODB.Recordset
'Dim gaji As Currency
'Dim gaji1 As String
'Dim listitem As listitem
'Dim M_DATA As New ClsNegoPTP
'
'Select Case Index
'    Case 0
'           ' If LstPayment.ListItems.Item(0).SubItems(1) Then
'
'            'End If
'
'            If Left(statusptp2, 3) = "POP" And cboPTP = "" Then
'             MsgBox "Anda Harus pilih jenis PTP Jika insert di nego payment"
'            Exit Sub
'            End If
'
'            If Left(cmbContacted, 3) = "OP-" Or C_Contacted.Value = 1 Then
'            Exit Sub
'            End If
'
'        With FrmNegoPTP
'                .Caption = "Tambah Data"
'                .Show vbModal
'                If .ok Then
'                 M_DATA.ADD_NegoPTP M_OBJCONN, .TxtCustid.Text, IIf(IsNull(.TDBDate1.Value), Null, Format(.TDBDate1.Value, "yyyy/mm/dd")), CStr(.TDBNumber1.Value), CStr(MDIForm1.TDBDate1.Value), ""
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
'
'
'    Case 1
'         If LstPayment.ListItems.Count = 0 Then
'            Exit Sub
'        End If
'           With FrmNegoPTP
'                .Caption = "Ubah Data"
'
'                .TDBDate1.Value = Format(LstPayment.SelectedItem.SubItems(2), "dd/mm/yyyy")
'                .TDBNumber1.Value = LstPayment.SelectedItem.SubItems(3)
'                .Show vbModal
'                If .ok Then
'
'                    M_DATA.UPDATE_NegoPTP M_OBJCONN, .TxtCustid.Text, CStr(.TDBDate1.Value), CStr(.TDBNumber1.Value), LstPayment.SelectedItem.SubItems(1)
'
'                    On Error GoTo add_error
'                    If M_DATA.ADD_OK Then
'                        'LstPayment.SelectedItem.SubItems(1) = ""
'                        LstPayment.SelectedItem.SubItems(2) = .TDBDate1.Value
'                        LstPayment.SelectedItem.SubItems(3) = .TDBNumber1.Value
'
'                    On Error GoTo 0
'                    End If
'                End If
'                Unload FrmNegoPTP
'            End With
'        Exit Sub
'    Case 2
'            If mdiform1.txtlevel.text <> "Agent" Then
'
'            If LstPayment.ListItems.Count = 0 Then
'                Exit Sub
'            End If
'            m_msgbox = MsgBox("Yakin Akan Dihapus...!!! ", vbCritical + vbOKCancel, "Peringatan")
'            If m_msgbox = 1 Then
'               M_DATA.DELETE_Nego_PTP M_OBJCONN, LstPayment.SelectedItem.SubItems(1)
'                If M_DATA.ADD_OK Then
'                    LstPayment.ListItems.Remove LstPayment.SelectedItem.Index
'                End If
'            End If
'        End If
'        Exit Sub
'
'
'End Select
'add_error:
'End Sub
'Private Sub VisitYES()
'Text1.BackColor = &HFF00&
'TxtCustid.BackColor = &H80000005
'TxtName.BackColor = &H80000005
'TDBNumber1.BackColor = &H80000005
'TXtDetails.BackColor = &H80000005
''LstVisit.BackColor = &HFF00&
'TxtAddress.BackColor = &H80000005
'TxtAddress.Enabled = True
'TXtDetails.Enabled = True
'Option7(0).Enabled = True
'Option7(1).Enabled = True
'Option7(2).Enabled = True
'
'
'End Sub
'Private Sub VisitNo()
'Text1.BackColor = &H8000000F
'TxtCustid.BackColor = &H8000000F
'TxtName.BackColor = &H8000000F
'TDBNumber1.BackColor = &H8000000F
'TXtDetails.BackColor = &H8000000F
'TxtAddress.BackColor = &H8000000F
''LstVisit.BackColor = &H8000000F
'Option8(1).Value = True
'Option7(0).Enabled = False
'Option7(1).Enabled = False
'Option7(2).Enabled = False
'
'TxtAddress.Enabled = False
'TXtDetails.Enabled = False
'End Sub
'
'Private Sub txtECno_Click()
'TYPETELP = "Emergency Contact"
'txtPhone.Text = txtECno.Value
'txtPhoneA.Text = txtECnoA.Value
'
'End Sub
'
'
'Private Sub txtECnoA_Change()
''txtECno.Text = txtECnoA.Text
'End Sub
'
'Private Sub txtECnoA_Click()
'TYPETELP = "Emergency Contact"
'txtPhone.Text = txtECno.Value
'txtPhoneA.Text = txtECnoA.Value
'End Sub
'
'Private Sub txtECnoA_LostFocus()
'If Len(txtECno.Text) > 3 Then
'    CmbPhone.AddItem "EconPhone"
'End If
'End Sub
'
'Private Sub txtHomeAdd1_Click()
'TYPETELP = "HOME1"
'    If Trim(AHomeAdd1(0).Value) = "021" Or AHomeAdd1(0).Value = "" Then
'        txtPhone.Text = txtHomeAdd1.Value
'        txtPhoneA.Text = txtHomeAdd1.Value
'    Else
'        txtPhone.Text = Trim(AHomeAdd1(0).Value) & Trim(txtHomeAdd1.Value)
'        txtPhoneA.Text = Trim(AHomeAdd1(0).Value) & Trim(txtHomeAdd1.Value)
'    End If
'End Sub
'
'Private Sub txtHomeAdd1A_Change()
''    txtHomeAdd1.Text = txtHomeAdd1A.Text
'End Sub
'
'Private Sub txtHomeAdd1A_Click()
'TYPETELP = "HOME1"
'    If Trim(AHomeAdd1(0).Value) = "021" Or AHomeAdd1(0).Value = "" Then
'        txtPhone.Text = txtHomeAdd1.Value
'        txtPhoneA.Text = txtHomeAdd1A.Value
'    Else
'        txtPhone.Text = Trim(AHomeAdd1(0).Value) & Trim(txtHomeAdd1.Value)
'        txtPhoneA.Text = Trim(AHomeAdd1(0).Value) & Trim(txtHomeAdd1A.Value)
'    End If
'End Sub
'
'Private Sub txtHomeAdd1A_LostFocus()
'    If txtHomeAdd1A.Value <> "" Then
'        CmbPhone.AddItem "AddHome1"
'    End If
'End Sub
'
'Private Sub txtHomeAdd2_Click()
'TYPETELP = "HOME2"
'If Trim(AHomeAdd2(1).Value) = "021" Or AHomeAdd2(1).Value = "" Then
'    txtPhone.Text = txtHomeAdd2.Value
'    txtPhoneA.Text = txtHomeAdd2.Value
'Else
'    txtPhone.Text = Trim(AHomeAdd2(1).Value) & Trim(txtHomeAdd2.Value)
'    txtPhoneA.Text = Trim(AHomeAdd2(1).Value) & Trim(txtHomeAdd2.Value)
'End If
'End Sub
'
'Private Sub txtHomeAdd2A_Change()
''txtHomeAdd2.Text = txtHomeAdd2A.Text
'End Sub
'
'Private Sub txtHomeAdd2A_Click()
'TYPETELP = "HOME2"
'If Trim(AHomeAdd2(1).Value) = "021" Or AHomeAdd2(1).Value = "" Then
'    txtPhone.Text = txtHomeAdd2.Value
'    txtPhoneA.Text = txtHomeAdd2A.Value
'Else
'    txtPhone.Text = Trim(AHomeAdd2(1).Value) & Trim(txtHomeAdd2.Value)
'    txtPhoneA.Text = Trim(AHomeAdd2(1).Value) & Trim(txtHomeAdd2A.Value)
'End If
'End Sub
'
'
'Private Sub txtHomeAdd2A_LostFocus()
'    If txtHomeAdd2A.Value <> "" Then
'        CmbPhone.AddItem "AddHome2"
'    End If
'End Sub
'
'Private Sub txtMobileAdd1A_Change()
''txtMobileAdd1.Text = txtMobileAdd1A.Text
'End Sub
'
'Private Sub txtMobileAdd1A_Click()
'TYPETELP = "MOBILE1"
'    txtPhone.Text = txtMobileAdd1.Value
'    txtPhoneA.Text = txtMobileAdd1A.Value
'End Sub
'
'Private Sub txtMobileAdd1A_LostFocus()
'If txtMobileAdd1A.Value <> "" Then
'    CmbPhone.AddItem "AddMobile1"
'End If
'End Sub
'
'Private Sub txtMobileAdd2A_Change()
''    txtMobileAdd2.Text = txtMobileAdd2A.Text
'End Sub
'
'Private Sub txtMobileAdd2A_Click()
'TYPETELP = "MOBILE2"
'    txtPhone.Text = txtMobileAdd2.Value
'    txtPhoneA.Text = txtMobileAdd2A.Value
'End Sub
'
'Private Sub txtMobileAdd2A_LostFocus()
'    If txtMobileAdd2A.Value <> "" Then
'        CmbPhone.AddItem "AddMobile2"
'    End If
'End Sub
'
'Private Sub txtMobileNo1A_Change()
''
'End Sub
'
'Private Sub txtOfficeAdd1_Click()
'TYPETELP = "OFFICE1"
'If Trim(AOfficeAdd(2).Value) = "021" Or AOfficeAdd(2).Value = "" Then
'    txtPhone.Text = txtOfficeAdd1.Value
'    txtPhoneA.Text = txtOfficeAdd1.Value
'Else
'    txtPhone.Text = Trim(AOfficeAdd(2).Value) & Trim(txtOfficeAdd1.Value)
'    txtPhoneA.Text = Trim(AOfficeAdd(2).Value) & Trim(txtOfficeAdd1.Value)
'End If
'End Sub
'
'Private Sub txtOfficeAdd1A_Change()
''txtOfficeAdd1.Text = txtOfficeAdd1A.Text
'End Sub
'
'Private Sub txtOfficeAdd1A_Click()
'TYPETELP = "OFFICE1"
'If Trim(AOfficeAdd(2).Value) = "021" Or AOfficeAdd(2).Value = "" Then
'    txtPhone.Text = txtOfficeAdd1.Value
'    txtPhoneA.Text = txtOfficeAdd1A.Value
'Else
'    txtPhone.Text = Trim(AOfficeAdd(2).Value) & Trim(txtOfficeAdd1.Value)
'    txtPhoneA.Text = Trim(AOfficeAdd(2).Value) & Trim(txtOfficeAdd1A.Value)
'End If
'End Sub
'
'Private Sub txtOfficeAdd1A_LostFocus()
'    If txtOfficeAdd1A.Value <> "" Then
'        CmbPhone.AddItem "AddOffice1"
'    End If
'End Sub
'
'Private Sub txtOfficeAdd2_Click()
'TYPETELP = "OFFICE2"
'If Trim(AOfficeAdd(3).Value) = "021" Or AOfficeAdd(3).Value = "" Then
'    txtPhone.Text = txtOfficeAdd2.Value
'    txtPhoneA.Text = txtOfficeAdd2.Value
'Else
'    txtPhone.Text = Trim(AOfficeAdd(3).Value) & Trim(txtOfficeAdd2.Value)
'    txtPhoneA.Text = Trim(AOfficeAdd(3).Value) & Trim(txtOfficeAdd2.Value)
'End If
'
'End Sub
'
'Private Sub txtMobileAdd1_Click()
'TYPETELP = "MOBILE1"
'    txtPhone.Text = txtMobileAdd1.Value
'    txtPhoneA.Text = txtMobileAdd1.Value
'End Sub
'
'Private Sub txtMobileAdd2_Click()
'TYPETELP = "MOBILE2"
'    txtPhone.Text = txtMobileAdd2.Value
'    txtPhoneA.Text = txtMobileAdd2.Value
'End Sub
'Public Sub UpdateAppv()
'If chkAppv(0).Value Then
'    x = MsgBox("Pindahkan data ke Agent DA ?", vbYesNo + vbExclamation, "Info !")
'    If x = vbYes Then
'        CMDSQL = "update mgm set F_pending='Pending',Agent='DA',PO_Agent='" & VIEW_MGMDATA.Combo1(0).Text & "' where custid='" + lblCustId.Caption + "'"
'        M_OBJCONN.Execute CMDSQL
'        spend = True
'        MsgBox "Data berhasil dipindah ke agent DA", vbInformation
'        VIEW_MGMDATA.LstVwSearchMgm.ListItems.CLEAR
'        MDIForm1.LstGrade.ListItems.CLEAR
'    End If
'Else
'    If chkAppv(1).Value Then
'        Dim spo As ADODB.Recordset
'        Set spo = New ADODB.Recordset
'        spo.CursorLocation = adUseClient
'        spo.Open "select PO_Agent from mgm where custid='" + lblCustId.Caption + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
'
'        If spo!PO_AGENT <> "" And IsNull(spo!PO_AGENT) = False Then
'            CMDSQL = "update mgm set F_pending='',AGENT=PO_Agent where custid='" + lblCustId.Caption + "'"
'            M_OBJCONN.Execute CMDSQL
'            CMDSQL = "update mgm set PO_Agent='' where custid='" + lblCustId.Caption + "'"
'            M_OBJCONN.Execute CMDSQL
'            MsgBox "Data berhasil dikembalikan", vbInformation
'            VIEW_MGMDATA.LstVwSearchMgm.ListItems.CLEAR
'            MDIForm1.LstGrade.ListItems.CLEAR
'        Else
'            MsgBox "Silahkan Pilih Status !," & vbCrLf & "untuk menyimpan hilangkan ceklist NO !", vbInformation
'            Exit Sub
'        End If
'    End If
'End If
'End Sub
'
'Private Sub txtOfficeAdd2A_Change()
''txtOfficeAdd2.Text = txtOfficeAdd2A.Text
'End Sub
'
'Private Sub txtOfficeAdd2A_Click()
'TYPETELP = "OFFICE2"
'If Trim(AOfficeAdd(3).Value) = "021" Or AOfficeAdd(3).Value = "" Then
'    txtPhone.Text = txtOfficeAdd2.Value
'    txtPhoneA.Text = txtOfficeAdd2A.Value
'Else
'    txtPhone.Text = Trim(AOfficeAdd(3).Value) & Trim(txtOfficeAdd2.Value)
'    txtPhoneA.Text = Trim(AOfficeAdd(3).Value) & Trim(txtOfficeAdd2A.Value)
'End If
'End Sub
'
'Private Sub txtOfficeAdd2A_LostFocus()
'    If txtOfficeAdd2A.Value <> "" Then
'        CmbPhone.AddItem "AddOffice2"
'    End If
'End Sub
'
'
