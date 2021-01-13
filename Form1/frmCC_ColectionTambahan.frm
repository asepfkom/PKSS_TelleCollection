VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCC_Colection2 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   10890
   ClientLeft      =   -75
   ClientTop       =   165
   ClientWidth     =   19080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmCC_ColectionTambahan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10890
   ScaleWidth      =   19080
   ShowInTaskbar   =   0   'False
   Begin VB.Timer TimerBlink 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   14280
      Top             =   5265
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   10905
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   19260
      _ExtentX        =   33973
      _ExtentY        =   19235
      _Version        =   196610
      Font3D          =   1
      ForeColor       =   12583104
      BackColor       =   16777215
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
      Begin VB.Frame Frame10 
         Appearance      =   0  'Flat
         BackColor       =   &H00B8E2D4&
         ForeColor       =   &H80000008&
         Height          =   4755
         Left            =   6900
         TabIndex        =   21
         Top             =   6180
         Width           =   12225
         Begin VB.Frame Frame15 
            Appearance      =   0  'Flat
            BackColor       =   &H002F735C&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   7
            Left            =   60
            TabIndex        =   23
            Top             =   120
            Width           =   2895
            Begin VB.Label Label38 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "History"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   7
               Left            =   540
               TabIndex        =   24
               Top             =   60
               Width           =   1335
            End
            Begin VB.Image Image1 
               Height          =   375
               Index           =   7
               Left            =   60
               Picture         =   "frmCC_ColectionTambahan.frx":000C
               Stretch         =   -1  'True
               Top             =   30
               Width           =   375
            End
         End
         Begin VB.Timer TimerBlinkSms 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   4440
            Top             =   1080
         End
         Begin VB.Timer TimerCekMapping 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   3420
            Top             =   840
         End
         Begin VB.CommandButton CmdHapusRemarks 
            Caption         =   "&Hapus Remarks"
            Height          =   435
            Left            =   3300
            TabIndex        =   22
            Top             =   120
            Visible         =   0   'False
            Width           =   1755
         End
         Begin MSComctlLib.ListView listview1 
            Height          =   4020
            Index           =   1
            Left            =   60
            TabIndex        =   25
            Top             =   660
            Width           =   12135
            _ExtentX        =   21405
            _ExtentY        =   7091
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
      Begin VB.Frame Frame19 
         BackColor       =   &H00B8E2D4&
         Height          =   2205
         Left            =   60
         TabIndex        =   1
         Top             =   8760
         Width           =   6555
         Begin VB.ComboBox cbolastcall 
            Appearance      =   0  'Flat
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
            ItemData        =   "frmCC_ColectionTambahan.frx":0480
            Left            =   4440
            List            =   "frmCC_ColectionTambahan.frx":0490
            TabIndex        =   5
            Top             =   540
            Width           =   2055
         End
         Begin VB.ComboBox cboaccount 
            Height          =   315
            Left            =   1320
            TabIndex        =   4
            Top             =   180
            Width           =   1905
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmCC_ColectionTambahan.frx":04AF
            Left            =   4440
            List            =   "frmCC_ColectionTambahan.frx":04B9
            TabIndex        =   3
            Top             =   180
            Width           =   2055
         End
         Begin VB.TextBox txtremarks 
            Height          =   1335
            Left            =   60
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   720
            Width           =   3135
         End
         Begin TDBDate6Ctl.TDBDate cmbDateSch 
            Height          =   315
            Left            =   4425
            TabIndex        =   6
            Top             =   900
            Width           =   1260
            _Version        =   65536
            _ExtentX        =   2222
            _ExtentY        =   556
            Calendar        =   "frmCC_ColectionTambahan.frx":04D1
            Caption         =   "frmCC_ColectionTambahan.frx":05E9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_ColectionTambahan.frx":0655
            Keys            =   "frmCC_ColectionTambahan.frx":0673
            Spin            =   "frmCC_ColectionTambahan.frx":06D1
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
         Begin TDBTime6Ctl.TDBTime cmbTimeSch 
            Height          =   315
            Left            =   5700
            TabIndex        =   7
            Top             =   900
            Width           =   795
            _Version        =   65536
            _ExtentX        =   1402
            _ExtentY        =   556
            Caption         =   "frmCC_ColectionTambahan.frx":06F9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "frmCC_ColectionTambahan.frx":0765
            Spin            =   "frmCC_ColectionTambahan.frx":07B5
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
         Begin Threed.SSCommand SSCommand1 
            Height          =   600
            Index           =   2
            Left            =   5040
            TabIndex        =   8
            Top             =   1320
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   1058
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
            Picture         =   "frmCC_ColectionTambahan.frx":07DD
            AutoSize        =   1
            Alignment       =   8
            PictureAlignment=   1
         End
         Begin Threed.SSCommand SSCommand1 
            Cancel          =   -1  'True
            Height          =   600
            Index           =   3
            Left            =   5760
            TabIndex        =   9
            Top             =   1320
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   1058
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
            Picture         =   "frmCC_ColectionTambahan.frx":0D10
            AutoSize        =   1
            Alignment       =   4
            PictureAlignment=   1
         End
         Begin Threed.SSCommand SSCommand1 
            Height          =   600
            Index           =   4
            Left            =   3600
            TabIndex        =   10
            Top             =   1320
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   1058
            _Version        =   196610
            Font3D          =   2
            MousePointer    =   16
            ForeColor       =   8388608
            PictureMaskColor=   -2147483644
            PictureFrames   =   1
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Picture         =   "frmCC_ColectionTambahan.frx":1375
            AutoSize        =   1
            Alignment       =   8
            PictureAlignment=   1
         End
         Begin Threed.SSCommand CmdUnlock 
            Height          =   600
            Left            =   4320
            TabIndex        =   11
            Top             =   1320
            Visible         =   0   'False
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   1058
            _Version        =   196610
            PictureFrames   =   1
            Enabled         =   0   'False
            Picture         =   "frmCC_ColectionTambahan.frx":D3C7
            AutoSize        =   1
            Alignment       =   8
         End
         Begin VB.Label Label31 
            BackColor       =   &H009AD6C2&
            Caption         =   "Remarks:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   20
            Top             =   480
            Width           =   1035
         End
         Begin VB.Label Label39 
            BackColor       =   &H009AD6C2&
            Caption         =   "Tgl Follow up"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3210
            TabIndex        =   19
            Top             =   900
            Width           =   1245
         End
         Begin VB.Label Label43 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H009AD6C2&
            Caption         =   "CPA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   3615
            TabIndex        =   18
            Top             =   1920
            Width           =   615
         End
         Begin VB.Label Label44 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H009AD6C2&
            Caption         =   "Exit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5790
            TabIndex        =   17
            Top             =   1920
            Width           =   585
         End
         Begin VB.Label Label43 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H009AD6C2&
            Caption         =   "Save"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   5040
            TabIndex        =   16
            Top             =   1920
            Width           =   675
         End
         Begin VB.Label Label31 
            BackColor       =   &H009AD6C2&
            Caption         =   "Contact with"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   3210
            TabIndex        =   15
            Top             =   570
            Width           =   1245
         End
         Begin VB.Label label1 
            BackColor       =   &H009AD6C2&
            Caption         =   "Status Call"
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
            Index           =   10
            Left            =   3210
            TabIndex        =   14
            Top             =   180
            Width           =   1215
         End
         Begin VB.Label label1 
            BackColor       =   &H009AD6C2&
            Caption         =   "Select Status"
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
            Height          =   315
            Index           =   12
            Left            =   60
            TabIndex        =   13
            Top             =   180
            Width           =   1305
         End
         Begin VB.Label Label43 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H009AD6C2&
            Caption         =   "Unlock"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   225
            Index           =   1
            Left            =   4350
            TabIndex        =   12
            Top             =   1920
            Visible         =   0   'False
            Width           =   615
         End
      End
      Begin VB.Frame Frame11 
         Appearance      =   0  'Flat
         BackColor       =   &H00B1FDD5&
         BorderStyle     =   0  'None
         Caption         =   "Frame11"
         ForeColor       =   &H80000008&
         Height          =   10875
         Left            =   6900
         TabIndex        =   33
         Top             =   0
         Width           =   12495
         Begin VB.Frame Frame7 
            Appearance      =   0  'Flat
            BackColor       =   &H00B8E2D4&
            ForeColor       =   &H80000008&
            Height          =   1725
            Left            =   6060
            TabIndex        =   127
            Top             =   8160
            Visible         =   0   'False
            Width           =   5805
            Begin VB.Timer Timer_cek_inbox 
               Enabled         =   0   'False
               Interval        =   30000
               Left            =   4020
               Top             =   420
            End
            Begin VB.TextBox Text4 
               Height          =   285
               Left            =   4200
               TabIndex        =   132
               Text            =   "0"
               Top             =   960
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.OptionButton Option9 
               BackColor       =   &H00B8E2D4&
               Caption         =   "Inbox"
               Height          =   255
               Left            =   4710
               TabIndex        =   131
               Top             =   120
               Value           =   -1  'True
               Width           =   735
            End
            Begin VB.OptionButton Option10 
               BackColor       =   &H00B8E2D4&
               Caption         =   "Send"
               Height          =   255
               Left            =   4710
               TabIndex        =   130
               Top             =   360
               Width           =   735
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Left            =   3360
               TabIndex        =   129
               Text            =   "0"
               Top             =   960
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.TextBox Text3 
               Height          =   285
               Left            =   3720
               TabIndex        =   128
               Text            =   "0"
               Top             =   960
               Visible         =   0   'False
               Width           =   375
            End
         End
         Begin VB.Frame Frame12 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   3270
            Left            =   60
            TabIndex        =   52
            Top             =   510
            Width           =   12060
            Begin VB.Frame Frame16 
               Appearance      =   0  'Flat
               BackColor       =   &H00B8E2D4&
               ForeColor       =   &H80000008&
               Height          =   3285
               Left            =   0
               TabIndex        =   90
               Top             =   -90
               Width           =   3615
               Begin VB.ComboBox CmbPhone 
                  Appearance      =   0  'Flat
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
                  Height          =   345
                  ItemData        =   "frmCC_ColectionTambahan.frx":D95C
                  Left            =   1170
                  List            =   "frmCC_ColectionTambahan.frx":D963
                  Locked          =   -1  'True
                  TabIndex        =   91
                  Text            =   "CmbPhone"
                  Top             =   210
                  Width           =   1680
               End
               Begin TDBMask6Ctl.TDBMask txtHomeNo2 
                  Height          =   255
                  Left            =   1500
                  TabIndex        =   92
                  Top             =   945
                  Visible         =   0   'False
                  Width           =   1335
                  _Version        =   65536
                  _ExtentX        =   2355
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":D96C
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":D9D8
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
               Begin TDBMask6Ctl.TDBMask txtOfficeNo2 
                  Height          =   255
                  Left            =   1500
                  TabIndex        =   93
                  Top             =   1605
                  Visible         =   0   'False
                  Width           =   1335
                  _Version        =   65536
                  _ExtentX        =   2355
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":DA1A
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":DA86
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
                  Left            =   915
                  TabIndex        =   94
                  Top             =   1965
                  Visible         =   0   'False
                  Width           =   1695
                  _Version        =   65536
                  _ExtentX        =   2990
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":DAC8
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":DB34
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
                  Left            =   900
                  TabIndex        =   95
                  Top             =   2295
                  Visible         =   0   'False
                  Width           =   1695
                  _Version        =   65536
                  _ExtentX        =   2990
                  _ExtentY        =   441
                  Caption         =   "frmCC_ColectionTambahan.frx":DB76
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":DBE2
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
                  Left            =   1500
                  TabIndex        =   96
                  Top             =   945
                  Width           =   1335
                  _Version        =   65536
                  _ExtentX        =   2355
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":DC24
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Lucida Console"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":DC90
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
               Begin TDBMask6Ctl.TDBMask txtOfficeNo2A 
                  Height          =   255
                  Left            =   1500
                  TabIndex        =   97
                  Top             =   1605
                  Width           =   1335
                  _Version        =   65536
                  _ExtentX        =   2355
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":DCD2
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":DD3E
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
                  Left            =   930
                  TabIndex        =   98
                  Top             =   1965
                  Width           =   1695
                  _Version        =   65536
                  _ExtentX        =   2990
                  _ExtentY        =   441
                  Caption         =   "frmCC_ColectionTambahan.frx":DD80
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Lucida Console"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":DDEC
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
                  Left            =   915
                  TabIndex        =   99
                  Top             =   2295
                  Width           =   1695
                  _Version        =   65536
                  _ExtentX        =   2990
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":DE2E
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Lucida Console"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":DE9A
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
                  Left            =   2895
                  TabIndex        =   100
                  Top             =   1245
                  Width           =   645
                  _Version        =   65536
                  _ExtentX        =   1138
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":DEDC
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":DF48
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
                  Left            =   2895
                  TabIndex        =   101
                  Top             =   1605
                  Width           =   645
                  _Version        =   65536
                  _ExtentX        =   1138
                  _ExtentY        =   441
                  Caption         =   "frmCC_ColectionTambahan.frx":DF8A
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":DFF6
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
               Begin TDBMask6Ctl.TDBMask txtHomeNo1A 
                  Height          =   255
                  Left            =   1500
                  TabIndex        =   102
                  Top             =   630
                  Width           =   1335
                  _Version        =   65536
                  _ExtentX        =   2355
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":E038
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":E0A4
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
                  Left            =   1500
                  TabIndex        =   103
                  Top             =   1275
                  Visible         =   0   'False
                  Width           =   1335
                  _Version        =   65536
                  _ExtentX        =   2355
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":E0E6
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":E152
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
               Begin TDBMask6Ctl.TDBMask txtOfficeNo1A 
                  Height          =   255
                  Left            =   1500
                  TabIndex        =   104
                  Top             =   1275
                  Width           =   1335
                  _Version        =   65536
                  _ExtentX        =   2355
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":E194
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Lucida Console"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":E200
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
               Begin TDBMask6Ctl.TDBMask AHome2 
                  Height          =   255
                  Left            =   915
                  TabIndex        =   105
                  Top             =   945
                  Width           =   540
                  _Version        =   65536
                  _ExtentX        =   952
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":E242
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":E2AE
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
                  PromptChar      =   " "
                  ReadOnly        =   -1
                  ShowContextMenu =   -1
                  ShowLiterals    =   0
                  TabAction       =   0
                  Text            =   "    "
                  Value           =   ""
               End
               Begin TDBMask6Ctl.TDBMask AOffice1 
                  Height          =   255
                  Left            =   915
                  TabIndex        =   106
                  Top             =   1275
                  Width           =   540
                  _Version        =   65536
                  _ExtentX        =   952
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":E2F0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":E35C
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
                  PromptChar      =   " "
                  ReadOnly        =   -1
                  ShowContextMenu =   -1
                  ShowLiterals    =   0
                  TabAction       =   0
                  Text            =   "    "
                  Value           =   ""
               End
               Begin TDBMask6Ctl.TDBMask AOffice2 
                  Height          =   255
                  Left            =   915
                  TabIndex        =   107
                  Top             =   1605
                  Width           =   540
                  _Version        =   65536
                  _ExtentX        =   952
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":E39E
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":E40A
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
                  PromptChar      =   " "
                  ReadOnly        =   -1
                  ShowContextMenu =   -1
                  ShowLiterals    =   0
                  TabAction       =   0
                  Text            =   "    "
                  Value           =   ""
               End
               Begin TDBMask6Ctl.TDBMask txtHomeNo1 
                  Height          =   255
                  Left            =   1500
                  TabIndex        =   108
                  Top             =   630
                  Visible         =   0   'False
                  Width           =   1335
                  _Version        =   65536
                  _ExtentX        =   2355
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":E44C
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Trebuchet MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":E4B8
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
               Begin TDBMask6Ctl.TDBMask AHome1 
                  Height          =   255
                  Left            =   915
                  TabIndex        =   109
                  Top             =   615
                  Width           =   540
                  _Version        =   65536
                  _ExtentX        =   952
                  _ExtentY        =   441
                  Caption         =   "frmCC_ColectionTambahan.frx":E4FA
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":E566
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
                  PromptChar      =   " "
                  ReadOnly        =   -1
                  ShowContextMenu =   -1
                  ShowLiterals    =   0
                  TabAction       =   0
                  Text            =   "    "
                  Value           =   ""
               End
               Begin TDBMask6Ctl.TDBMask tdbhptrace 
                  Height          =   255
                  Left            =   915
                  TabIndex        =   110
                  Top             =   1965
                  Visible         =   0   'False
                  Width           =   1695
                  _Version        =   65536
                  _ExtentX        =   2990
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":E5A8
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":E614
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
               Begin TDBMask6Ctl.TDBMask tdbtelptrace 
                  Height          =   255
                  Left            =   900
                  TabIndex        =   111
                  Top             =   2295
                  Visible         =   0   'False
                  Width           =   1695
                  _Version        =   65536
                  _ExtentX        =   2990
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":E656
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":E6C2
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
               Begin VB.Label label1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H009AD6C2&
                  Caption         =   "HP II"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   3
                  Left            =   120
                  TabIndex        =   122
                  Top             =   2295
                  Width           =   735
               End
               Begin VB.Label label1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H009AD6C2&
                  Caption         =   "HP I"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   4
                  Left            =   120
                  TabIndex        =   121
                  Top             =   1935
                  Width           =   735
               End
               Begin VB.Label label1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H009AD6C2&
                  Caption         =   "Rumah II"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   5
                  Left            =   120
                  TabIndex        =   120
                  Top             =   945
                  Width           =   735
               End
               Begin VB.Label label1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H009AD6C2&
                  Caption         =   "Rumah I"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   6
                  Left            =   120
                  TabIndex        =   119
                  Top             =   615
                  Width           =   735
               End
               Begin VB.Label label1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H009AD6C2&
                  Caption         =   "Kantor I"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   7
                  Left            =   120
                  TabIndex        =   118
                  Top             =   1275
                  Width           =   735
               End
               Begin VB.Label label1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H009AD6C2&
                  Caption         =   "Kantor II"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   8
                  Left            =   120
                  TabIndex        =   117
                  Top             =   1605
                  Width           =   735
               End
               Begin VB.Label label1 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H009AD6C2&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "No Tujuan :"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   285
                  Index           =   9
                  Left            =   120
                  TabIndex        =   116
                  Top             =   240
                  Width           =   1005
               End
               Begin VB.Label label1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H009AD6C2&
                  Caption         =   "HP Trace"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   11
                  Left            =   960
                  TabIndex        =   115
                  Top             =   2010
                  Visible         =   0   'False
                  Width           =   915
               End
               Begin VB.Label label1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H009AD6C2&
                  Caption         =   "Telp Trace"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   15
                  Left            =   960
                  TabIndex        =   114
                  Top             =   2310
                  Visible         =   0   'False
                  Width           =   915
               End
               Begin VB.Label Label22 
                  Appearance      =   0  'Flat
                  BackColor       =   &H009AD6C2&
                  Caption         =   "Mother Name"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   113
                  Top             =   2640
                  Width           =   735
                  WordWrap        =   -1  'True
               End
               Begin VB.Label LblMother 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "-"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   915
                  TabIndex        =   112
                  Top             =   2640
                  Width           =   1695
               End
            End
            Begin VB.Frame Frame17 
               Appearance      =   0  'Flat
               BackColor       =   &H00B8E2D4&
               ForeColor       =   &H80000008&
               Height          =   3285
               Left            =   3660
               TabIndex        =   63
               Top             =   -90
               Width           =   4035
               Begin TDBMask6Ctl.TDBMask txtOfficeAdd1 
                  Height          =   255
                  Left            =   1380
                  TabIndex        =   64
                  Top             =   720
                  Visible         =   0   'False
                  Width           =   1815
                  _Version        =   65536
                  _ExtentX        =   3201
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":E704
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":E770
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
                  ReadOnly        =   1
                  ShowContextMenu =   1
                  ShowLiterals    =   0
                  TabAction       =   0
                  Text            =   "                  "
                  Value           =   ""
               End
               Begin TDBMask6Ctl.TDBMask txtOfficeAdd2 
                  Height          =   255
                  Left            =   1380
                  TabIndex        =   65
                  Top             =   990
                  Visible         =   0   'False
                  Width           =   1815
                  _Version        =   65536
                  _ExtentX        =   3201
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":E7B2
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":E81E
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
                  ReadOnly        =   1
                  ShowContextMenu =   1
                  ShowLiterals    =   0
                  TabAction       =   0
                  Text            =   "                  "
                  Value           =   ""
               End
               Begin TDBMask6Ctl.TDBMask AOfficeAdd 
                  Height          =   255
                  Index           =   2
                  Left            =   915
                  TabIndex        =   66
                  Top             =   720
                  Width           =   405
                  _Version        =   65536
                  _ExtentX        =   714
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":E860
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":E8CC
                  AlignHorizontal =   0
                  AlignVertical   =   0
                  Appearance      =   0
                  AllowSpace      =   1
                  AutoConvert     =   1
                  BackColor       =   16777215
                  BorderStyle     =   0
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
                  Left            =   915
                  TabIndex        =   67
                  Top             =   1020
                  Width           =   405
                  _Version        =   65536
                  _ExtentX        =   714
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":E90E
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":E97A
                  AlignHorizontal =   0
                  AlignVertical   =   0
                  Appearance      =   0
                  AllowSpace      =   1
                  AutoConvert     =   1
                  BackColor       =   16777215
                  BorderStyle     =   0
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
               Begin TDBMask6Ctl.TDBMask AHomeAdd1 
                  Height          =   250
                  Index           =   0
                  Left            =   915
                  TabIndex        =   68
                  Top             =   135
                  Width           =   405
                  _Version        =   65536
                  _ExtentX        =   714
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":E9BC
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":EA28
                  AlignHorizontal =   0
                  AlignVertical   =   0
                  Appearance      =   0
                  AllowSpace      =   1
                  AutoConvert     =   1
                  BackColor       =   16777215
                  BorderStyle     =   0
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
                  Left            =   915
                  TabIndex        =   69
                  Top             =   420
                  Width           =   405
                  _Version        =   65536
                  _ExtentX        =   714
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":EA6A
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":EAD6
                  AlignHorizontal =   0
                  AlignVertical   =   0
                  Appearance      =   0
                  AllowSpace      =   1
                  AutoConvert     =   1
                  BackColor       =   16777215
                  BorderStyle     =   0
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
               Begin TDBMask6Ctl.TDBMask txtOfficeAdd1A 
                  Height          =   255
                  Left            =   1380
                  TabIndex        =   70
                  Top             =   720
                  Width           =   1815
                  _Version        =   65536
                  _ExtentX        =   3201
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":EB18
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":EB84
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
                  ReadOnly        =   1
                  ShowContextMenu =   1
                  ShowLiterals    =   0
                  TabAction       =   0
                  Text            =   "                  "
                  Value           =   ""
               End
               Begin TDBMask6Ctl.TDBMask txtOfficeAdd2A 
                  Height          =   255
                  Left            =   1380
                  TabIndex        =   71
                  Top             =   990
                  Width           =   1815
                  _Version        =   65536
                  _ExtentX        =   3201
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":EBC6
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":EC32
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
                  ReadOnly        =   1
                  ShowContextMenu =   1
                  ShowLiterals    =   0
                  TabAction       =   0
                  Text            =   "                  "
                  Value           =   ""
               End
               Begin TDBMask6Ctl.TDBMask TxtExt3 
                  Height          =   255
                  Left            =   3285
                  TabIndex        =   72
                  Top             =   720
                  Width           =   675
                  _Version        =   65536
                  _ExtentX        =   1191
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":EC74
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":ECE0
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
                  Left            =   3285
                  TabIndex        =   73
                  Top             =   1020
                  Width           =   675
                  _Version        =   65536
                  _ExtentX        =   1191
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":ED22
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":ED8E
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
               Begin TDBMask6Ctl.TDBMask txtMobileAdd1 
                  Height          =   255
                  Left            =   900
                  TabIndex        =   74
                  Top             =   1350
                  Visible         =   0   'False
                  Width           =   1815
                  _Version        =   65536
                  _ExtentX        =   3201
                  _ExtentY        =   441
                  Caption         =   "frmCC_ColectionTambahan.frx":EDD0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":EE3C
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
                  Left            =   900
                  TabIndex        =   75
                  Top             =   1650
                  Visible         =   0   'False
                  Width           =   1815
                  _Version        =   65536
                  _ExtentX        =   3201
                  _ExtentY        =   441
                  Caption         =   "frmCC_ColectionTambahan.frx":EE7E
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":EEEA
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
                  Left            =   900
                  TabIndex        =   76
                  Top             =   1350
                  Width           =   1815
                  _Version        =   65536
                  _ExtentX        =   3201
                  _ExtentY        =   441
                  Caption         =   "frmCC_ColectionTambahan.frx":EF2C
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":EF98
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
                  Left            =   900
                  TabIndex        =   77
                  Top             =   1650
                  Width           =   1815
                  _Version        =   65536
                  _ExtentX        =   3201
                  _ExtentY        =   441
                  Caption         =   "frmCC_ColectionTambahan.frx":EFDA
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":F046
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
               Begin RichTextLib.RichTextBox AddrNow 
                  Height          =   1155
                  Left            =   900
                  TabIndex        =   78
                  Top             =   1950
                  Width           =   3075
                  _ExtentX        =   5424
                  _ExtentY        =   2037
                  _Version        =   393217
                  BackColor       =   16777215
                  ScrollBars      =   2
                  Appearance      =   0
                  TextRTF         =   $"frmCC_ColectionTambahan.frx":F088
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
               Begin TDBMask6Ctl.TDBMask txtHomeAdd1 
                  Height          =   255
                  Left            =   1380
                  TabIndex        =   79
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1815
                  _Version        =   65536
                  _ExtentX        =   3201
                  _ExtentY        =   441
                  Caption         =   "frmCC_ColectionTambahan.frx":F109
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":F175
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
                  ReadOnly        =   1
                  ShowContextMenu =   1
                  ShowLiterals    =   0
                  TabAction       =   0
                  Text            =   "                  "
                  Value           =   ""
               End
               Begin TDBMask6Ctl.TDBMask txtHomeAdd2 
                  Height          =   255
                  Left            =   1380
                  TabIndex        =   80
                  Top             =   420
                  Visible         =   0   'False
                  Width           =   1815
                  _Version        =   65536
                  _ExtentX        =   3201
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":F1B7
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":F223
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
                  ReadOnly        =   1
                  ShowContextMenu =   1
                  ShowLiterals    =   0
                  TabAction       =   0
                  Text            =   "                  "
                  Value           =   ""
               End
               Begin TDBMask6Ctl.TDBMask txtHomeAdd1A 
                  Height          =   250
                  Left            =   1380
                  TabIndex        =   81
                  Top             =   120
                  Width           =   1815
                  _Version        =   65536
                  _ExtentX        =   3201
                  _ExtentY        =   441
                  Caption         =   "frmCC_ColectionTambahan.frx":F265
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":F2D1
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
                  ReadOnly        =   1
                  ShowContextMenu =   1
                  ShowLiterals    =   0
                  TabAction       =   0
                  Text            =   "                  "
                  Value           =   ""
               End
               Begin TDBMask6Ctl.TDBMask txtHomeAdd2A 
                  Height          =   255
                  Left            =   1380
                  TabIndex        =   82
                  Top             =   420
                  Width           =   1815
                  _Version        =   65536
                  _ExtentX        =   3201
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":F313
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":F37F
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
                  ReadOnly        =   1
                  ShowContextMenu =   1
                  ShowLiterals    =   0
                  TabAction       =   0
                  Text            =   "                  "
                  Value           =   ""
               End
               Begin VB.Label label1 
                  BackColor       =   &H009AD6C2&
                  Caption         =   "Kantor II"
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
                  Index           =   16
                  Left            =   120
                  TabIndex        =   89
                  Top             =   1020
                  Width           =   765
               End
               Begin VB.Label label1 
                  BackColor       =   &H009AD6C2&
                  Caption         =   "Kantor I"
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
                  Index           =   17
                  Left            =   120
                  TabIndex        =   88
                  Top             =   720
                  Width           =   765
               End
               Begin VB.Label label1 
                  BackColor       =   &H009AD6C2&
                  Caption         =   "Rumah II"
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
                  Height          =   250
                  Index           =   19
                  Left            =   120
                  TabIndex        =   87
                  Top             =   420
                  Width           =   765
               End
               Begin VB.Label label1 
                  BackColor       =   &H009AD6C2&
                  Caption         =   "Rumah I"
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
                  Height          =   250
                  Index           =   20
                  Left            =   120
                  TabIndex        =   86
                  Top             =   120
                  Width           =   765
               End
               Begin VB.Label label1 
                  BackColor       =   &H009AD6C2&
                  Caption         =   "HP II"
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
                  Index           =   13
                  Left            =   120
                  TabIndex        =   85
                  Top             =   1650
                  Width           =   765
               End
               Begin VB.Label label1 
                  BackColor       =   &H009AD6C2&
                  Caption         =   "HP I"
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
                  Index           =   14
                  Left            =   120
                  TabIndex        =   84
                  Top             =   1350
                  Width           =   765
               End
               Begin VB.Label Label19 
                  BackColor       =   &H009AD6C2&
                  Caption         =   "Add  Adress:"
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
                  Left            =   120
                  TabIndex        =   83
                  Top             =   1950
                  Width           =   795
               End
            End
            Begin VB.Frame Frame20 
               Appearance      =   0  'Flat
               BackColor       =   &H00B8E2D4&
               ForeColor       =   &H80000008&
               Height          =   3285
               Left            =   7710
               TabIndex        =   53
               Top             =   -90
               Width           =   4215
               Begin VB.TextBox txtECAdd 
                  Appearance      =   0  'Flat
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
                  Height          =   765
                  Left            =   735
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   55
                  Top             =   720
                  Width           =   3270
               End
               Begin VB.TextBox txtremarkstrace 
                  Appearance      =   0  'Flat
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
                  Height          =   1365
                  Left            =   30
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   54
                  Top             =   1860
                  Width           =   4110
               End
               Begin TDBMask6Ctl.TDBMask txtECnoA 
                  Height          =   255
                  Left            =   720
                  TabIndex        =   56
                  Top             =   150
                  Width           =   1815
                  _Version        =   65536
                  _ExtentX        =   3201
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":F3C1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Trebuchet MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":F42D
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
               Begin RichTextLib.RichTextBox TxtEC 
                  Height          =   255
                  Left            =   720
                  TabIndex        =   57
                  Top             =   450
                  Width           =   3210
                  _ExtentX        =   5662
                  _ExtentY        =   450
                  _Version        =   393217
                  BackColor       =   16777215
                  Appearance      =   0
                  TextRTF         =   $"frmCC_ColectionTambahan.frx":F46F
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
                  Left            =   720
                  TabIndex        =   58
                  Top             =   150
                  Visible         =   0   'False
                  Width           =   1815
                  _Version        =   65536
                  _ExtentX        =   3201
                  _ExtentY        =   450
                  Caption         =   "frmCC_ColectionTambahan.frx":F4F0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Trebuchet MS"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Keys            =   "frmCC_ColectionTambahan.frx":F55C
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
               Begin VB.Label Label21 
                  BackColor       =   &H009AD6C2&
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
                  Height          =   255
                  Left            =   30
                  TabIndex        =   62
                  Top             =   420
                  Width           =   660
               End
               Begin VB.Label Label23 
                  BackColor       =   &H009AD6C2&
                  Caption         =   "Telp "
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
                  Left            =   30
                  TabIndex        =   61
                  Top             =   150
                  Width           =   1815
               End
               Begin VB.Label Label34 
                  Alignment       =   2  'Center
                  BackColor       =   &H009AD6C2&
                  Caption         =   "Add.Info From Tracer"
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
                  Left            =   30
                  TabIndex        =   60
                  Top             =   1560
                  Width           =   4125
               End
               Begin VB.Label Label35 
                  BackColor       =   &H009AD6C2&
                  Caption         =   "Addr"
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
                  Height          =   765
                  Left            =   30
                  TabIndex        =   59
                  Top             =   720
                  Width           =   705
               End
            End
         End
         Begin VB.Frame Frame15 
            Appearance      =   0  'Flat
            BackColor       =   &H002F735C&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   1
            Left            =   60
            TabIndex        =   50
            Top             =   60
            Width           =   2895
            Begin VB.Label Label38 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Phone Information"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   540
               TabIndex        =   51
               Top             =   105
               Width           =   1815
            End
            Begin VB.Image Image1 
               Height          =   375
               Index           =   1
               Left            =   60
               Picture         =   "frmCC_ColectionTambahan.frx":F59E
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
            Height          =   450
            Index           =   5
            Left            =   3720
            TabIndex        =   48
            Top             =   60
            Visible         =   0   'False
            Width           =   2895
            Begin VB.Label Label38 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Additional Info"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   540
               TabIndex        =   49
               Top             =   105
               Width           =   1575
            End
            Begin VB.Image Image1 
               Height          =   375
               Index           =   5
               Left            =   75
               Picture         =   "frmCC_ColectionTambahan.frx":10E38
               Stretch         =   -1  'True
               Top             =   60
               Width           =   375
            End
         End
         Begin VB.Frame FrmPayment 
            Appearance      =   0  'Flat
            BackColor       =   &H00B8E2D4&
            ForeColor       =   &H80000008&
            Height          =   1770
            Left            =   6000
            TabIndex        =   39
            Top             =   4260
            Width           =   6195
            Begin VB.CommandButton CmdDeletePelunasan 
               BackColor       =   &H00B8E2D4&
               Caption         =   "Hapus"
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   420
               Left            =   3795
               MaskColor       =   &H00000000&
               Style           =   1  'Graphical
               TabIndex        =   40
               Top             =   1050
               Visible         =   0   'False
               Width           =   795
            End
            Begin TDBNumber6Ctl.TDBNumber txtSisaHutang 
               Height          =   255
               Left            =   4845
               TabIndex        =   41
               Top             =   750
               Width           =   1230
               _Version        =   65536
               _ExtentX        =   2170
               _ExtentY        =   450
               Calculator      =   "frmCC_ColectionTambahan.frx":11357
               Caption         =   "frmCC_ColectionTambahan.frx":11377
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_ColectionTambahan.frx":113E3
               Keys            =   "frmCC_ColectionTambahan.frx":11401
               Spin            =   "frmCC_ColectionTambahan.frx":1144B
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
               Left            =   4845
               TabIndex        =   42
               Top             =   480
               Width           =   1230
               _Version        =   65536
               _ExtentX        =   2170
               _ExtentY        =   450
               Calculator      =   "frmCC_ColectionTambahan.frx":11473
               Caption         =   "frmCC_ColectionTambahan.frx":11493
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_ColectionTambahan.frx":114FF
               Keys            =   "frmCC_ColectionTambahan.frx":1151D
               Spin            =   "frmCC_ColectionTambahan.frx":11567
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
               Left            =   4845
               TabIndex        =   43
               Top             =   195
               Width           =   1230
               _Version        =   65536
               _ExtentX        =   2170
               _ExtentY        =   450
               Calculator      =   "frmCC_ColectionTambahan.frx":1158F
               Caption         =   "frmCC_ColectionTambahan.frx":115AF
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_ColectionTambahan.frx":1161B
               Keys            =   "frmCC_ColectionTambahan.frx":11639
               Spin            =   "frmCC_ColectionTambahan.frx":11683
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
               MaxValue        =   99999999999
               MinValue        =   -99999999999
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
               Height          =   1530
               Index           =   0
               Left            =   45
               TabIndex        =   44
               Top             =   180
               Width           =   3675
               _ExtentX        =   6482
               _ExtentY        =   2699
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
            Begin VB.Label Label10 
               BackColor       =   &H009AD6C2&
               Caption         =   "Jml PTP:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   3795
               TabIndex        =   47
               Top             =   195
               Width           =   1005
            End
            Begin VB.Label Label13 
               BackColor       =   &H009AD6C2&
               Caption         =   "Jml Dibayar:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3795
               TabIndex        =   46
               Top             =   480
               Width           =   1005
            End
            Begin VB.Label Label15 
               BackColor       =   &H009AD6C2&
               Caption         =   "Sisa:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   3795
               TabIndex        =   45
               Top             =   750
               Width           =   1005
            End
         End
         Begin VB.Frame Frame15 
            Appearance      =   0  'Flat
            BackColor       =   &H002F735C&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   2
            Left            =   7770
            TabIndex        =   37
            Top             =   60
            Visible         =   0   'False
            Width           =   2895
            Begin VB.Image Image1 
               Height          =   375
               Index           =   2
               Left            =   90
               Picture         =   "frmCC_ColectionTambahan.frx":116AB
               Stretch         =   -1  'True
               Top             =   60
               Width           =   375
            End
            Begin VB.Label Label38 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Emergency Contact"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   510
               TabIndex        =   38
               Top             =   120
               Width           =   2175
            End
         End
         Begin VB.Frame Frame15 
            Appearance      =   0  'Flat
            BackColor       =   &H002F735C&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   4
            Left            =   6000
            TabIndex        =   35
            Top             =   3900
            Width           =   2895
            Begin VB.Label Label38 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Detail Payment"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   180
               TabIndex        =   36
               Top             =   105
               Width           =   1575
            End
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Inbox/Outbox"
            Enabled         =   0   'False
            Height          =   720
            Left            =   3840
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   4200
            Width           =   1665
         End
         Begin Threed.SSCommand SSCommand1 
            Height          =   720
            Index           =   0
            Left            =   840
            TabIndex        =   123
            Top             =   4200
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   1270
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
            Picture         =   "frmCC_ColectionTambahan.frx":12F45
            AutoSize        =   1
            ButtonStyle     =   2
            PictureAlignment=   1
         End
         Begin Threed.SSCommand SSCommand1 
            Height          =   720
            Index           =   1
            Left            =   1860
            TabIndex        =   124
            Top             =   4200
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   1270
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
            Picture         =   "frmCC_ColectionTambahan.frx":13405
            AutoSize        =   1
            ButtonStyle     =   2
            PictureAlignment=   1
         End
         Begin Threed.SSCommand SSCommand1 
            Height          =   720
            Index           =   7
            Left            =   5880
            TabIndex        =   125
            Top             =   5040
            Visible         =   0   'False
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   1270
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
            Picture         =   "frmCC_ColectionTambahan.frx":13921
            AutoSize        =   1
            ButtonStyle     =   2
            PictureAlignment=   1
         End
         Begin Threed.SSCommand SSCommand1 
            Height          =   720
            Index           =   5
            Left            =   2880
            TabIndex        =   126
            Top             =   4200
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   1270
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
            Picture         =   "frmCC_ColectionTambahan.frx":13E3D
            AutoSize        =   1
            Alignment       =   8
            PictureAlignment=   1
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
            Left            =   7980
            TabIndex        =   139
            Top             =   7080
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label Label42 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H009AD6C2&
            Caption         =   "Hang Up"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   1860
            TabIndex        =   138
            Top             =   4920
            Width           =   900
         End
         Begin VB.Label Label42 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H009AD6C2&
            Caption         =   "Call"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   840
            TabIndex        =   137
            Top             =   4920
            Width           =   900
         End
         Begin VB.Label Label42 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H009AD6C2&
            Caption         =   "OST"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   5880
            TabIndex        =   136
            Top             =   5760
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label Label42 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H009AD6C2&
            Caption         =   "Offers"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   2880
            TabIndex        =   135
            Top             =   4920
            Width           =   900
         End
         Begin VB.Label LabelSms 
            Alignment       =   2  'Center
            BackColor       =   &H0080FF80&
            Caption         =   "Label SMS"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3840
            TabIndex        =   134
            Top             =   4920
            Width           =   1665
         End
         Begin VB.Label Label48 
            BackStyle       =   0  'Transparent
            Caption         =   "RITCARD"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   300
            TabIndex        =   133
            Top             =   5340
            Width           =   1695
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H0000FF00&
            BackStyle       =   1  'Opaque
            FillColor       =   &H0000FF00&
            Height          =   75
            Left            =   300
            Top             =   5340
            Width           =   5655
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   75
            Left            =   1680
            Top             =   5400
            Width           =   4155
         End
      End
      Begin VB.Frame Frame1 
         Height          =   930
         Left            =   9690
         TabIndex        =   26
         Top             =   9210
         Width           =   2775
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
            TabIndex        =   32
            Top             =   135
            Width           =   1500
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
            TabIndex        =   31
            Top             =   1590
            Width           =   1890
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
            Left            =   1905
            TabIndex        =   30
            Top             =   285
            Visible         =   0   'False
            Width           =   510
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
            Left            =   2460
            TabIndex        =   29
            Top             =   315
            Visible         =   0   'False
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
            TabIndex        =   28
            Top             =   315
            Width           =   60
         End
         Begin VB.Label LblStatus 
            Caption         =   "Label42"
            Height          =   255
            Left            =   600
            TabIndex        =   27
            Top             =   360
            Width           =   255
         End
      End
      Begin VB.Frame Frame13 
         Appearance      =   0  'Flat
         BackColor       =   &H00B1FDD5&
         BorderStyle     =   0  'None
         Caption         =   "Frame13"
         ForeColor       =   &H80000008&
         Height          =   10995
         Left            =   -90
         TabIndex        =   140
         Top             =   30
         Width           =   6825
         Begin VB.CheckBox C_PTP 
            BackColor       =   &H00B8E2D4&
            Caption         =   "PTP"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Left            =   300
            TabIndex        =   153
            Top             =   5760
            Width           =   750
         End
         Begin VB.Frame Frame14 
            Appearance      =   0  'Flat
            BackColor       =   &H00B1FDD5&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3795
            Left            =   240
            TabIndex        =   181
            Top             =   390
            Width           =   6465
            Begin VB.TextBox Text6 
               Height          =   285
               Left            =   4860
               TabIndex        =   182
               Top             =   3960
               Visible         =   0   'False
               Width           =   585
            End
            Begin RichTextLib.RichTextBox lblOfficeAddr 
               Height          =   675
               Left            =   780
               TabIndex        =   183
               Top             =   2160
               Width           =   3000
               _ExtentX        =   5292
               _ExtentY        =   1191
               _Version        =   393217
               BackColor       =   16777215
               BorderStyle     =   0
               ReadOnly        =   -1  'True
               ScrollBars      =   2
               Appearance      =   0
               TextRTF         =   $"frmCC_ColectionTambahan.frx":149D9
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin TDBDate6Ctl.TDBDate lblDate 
               Height          =   285
               Left            =   2265
               TabIndex        =   184
               Top             =   1095
               Visible         =   0   'False
               Width           =   1545
               _Version        =   65536
               _ExtentX        =   2725
               _ExtentY        =   503
               Calendar        =   "frmCC_ColectionTambahan.frx":14A50
               Caption         =   "frmCC_ColectionTambahan.frx":14B68
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_ColectionTambahan.frx":14BD4
               Keys            =   "frmCC_ColectionTambahan.frx":14BF2
               Spin            =   "frmCC_ColectionTambahan.frx":14C50
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
               Left            =   780
               TabIndex        =   185
               Top             =   1425
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   1217
               _Version        =   393217
               BackColor       =   16777215
               BorderStyle     =   0
               ReadOnly        =   -1  'True
               ScrollBars      =   2
               Appearance      =   0
               TextRTF         =   $"frmCC_ColectionTambahan.frx":14C78
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin TDBDate6Ctl.TDBDate lblOpenDate 
               Height          =   255
               Left            =   4860
               TabIndex        =   186
               Top             =   1170
               Width           =   1545
               _Version        =   65536
               _ExtentX        =   2725
               _ExtentY        =   450
               Calendar        =   "frmCC_ColectionTambahan.frx":14CEF
               Caption         =   "frmCC_ColectionTambahan.frx":14E07
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_ColectionTambahan.frx":14E73
               Keys            =   "frmCC_ColectionTambahan.frx":14E91
               Spin            =   "frmCC_ColectionTambahan.frx":14EEF
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
               Height          =   255
               Left            =   4860
               TabIndex        =   187
               Top             =   1455
               Width           =   1545
               _Version        =   65536
               _ExtentX        =   2725
               _ExtentY        =   450
               Calendar        =   "frmCC_ColectionTambahan.frx":14F17
               Caption         =   "frmCC_ColectionTambahan.frx":1502F
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_ColectionTambahan.frx":1509B
               Keys            =   "frmCC_ColectionTambahan.frx":150B9
               Spin            =   "frmCC_ColectionTambahan.frx":15117
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
               Left            =   4860
               TabIndex        =   188
               Top             =   840
               Width           =   1545
               _Version        =   65536
               _ExtentX        =   2725
               _ExtentY        =   450
               Calculator      =   "frmCC_ColectionTambahan.frx":1513F
               Caption         =   "frmCC_ColectionTambahan.frx":1515F
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_ColectionTambahan.frx":151CB
               Keys            =   "frmCC_ColectionTambahan.frx":151E9
               Spin            =   "frmCC_ColectionTambahan.frx":15233
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1701642245
               MinValueVT      =   3801093
            End
            Begin TDBNumber6Ctl.TDBNumber lblAmount 
               Height          =   255
               Left            =   4860
               TabIndex        =   189
               Top             =   210
               Width           =   1545
               _Version        =   65536
               _ExtentX        =   2725
               _ExtentY        =   450
               Calculator      =   "frmCC_ColectionTambahan.frx":1525B
               Caption         =   "frmCC_ColectionTambahan.frx":1527B
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_ColectionTambahan.frx":152E7
               Keys            =   "frmCC_ColectionTambahan.frx":15305
               Spin            =   "frmCC_ColectionTambahan.frx":1534F
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
            Begin TDBNumber6Ctl.TDBNumber lblLastPay 
               Height          =   255
               Left            =   4860
               TabIndex        =   190
               Top             =   2040
               Width           =   1545
               _Version        =   65536
               _ExtentX        =   2725
               _ExtentY        =   450
               Calculator      =   "frmCC_ColectionTambahan.frx":15377
               Caption         =   "frmCC_ColectionTambahan.frx":15397
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_ColectionTambahan.frx":15403
               Keys            =   "frmCC_ColectionTambahan.frx":15421
               Spin            =   "frmCC_ColectionTambahan.frx":1546B
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1701642245
               MinValueVT      =   3801093
            End
            Begin TDBDate6Ctl.TDBDate lblPayDt 
               Height          =   255
               Left            =   4860
               TabIndex        =   191
               Top             =   1755
               Width           =   1545
               _Version        =   65536
               _ExtentX        =   2725
               _ExtentY        =   450
               Calendar        =   "frmCC_ColectionTambahan.frx":15493
               Caption         =   "frmCC_ColectionTambahan.frx":155AB
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_ColectionTambahan.frx":15617
               Keys            =   "frmCC_ColectionTambahan.frx":15635
               Spin            =   "frmCC_ColectionTambahan.frx":15693
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
               Height          =   255
               Left            =   6330
               TabIndex        =   192
               Top             =   2160
               Visible         =   0   'False
               Width           =   525
               _Version        =   65536
               _ExtentX        =   926
               _ExtentY        =   450
               Calculator      =   "frmCC_ColectionTambahan.frx":156BB
               Caption         =   "frmCC_ColectionTambahan.frx":156DB
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_ColectionTambahan.frx":15747
               Keys            =   "frmCC_ColectionTambahan.frx":15765
               Spin            =   "frmCC_ColectionTambahan.frx":157AF
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1701642245
               MinValueVT      =   3801093
            End
            Begin TDBNumber6Ctl.TDBNumber txtPrinciple_A 
               Height          =   300
               Left            =   6000
               TabIndex        =   193
               Top             =   555
               Visible         =   0   'False
               Width           =   180
               _Version        =   65536
               _ExtentX        =   317
               _ExtentY        =   529
               Calculator      =   "frmCC_ColectionTambahan.frx":157D7
               Caption         =   "frmCC_ColectionTambahan.frx":157F7
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_ColectionTambahan.frx":15863
               Keys            =   "frmCC_ColectionTambahan.frx":15881
               Spin            =   "frmCC_ColectionTambahan.frx":158CB
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   0
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1701642245
               MinValueVT      =   3801093
            End
            Begin TDBNumber6Ctl.TDBNumber LblPrompA 
               Height          =   255
               Left            =   4860
               TabIndex        =   194
               Top             =   510
               Visible         =   0   'False
               Width           =   1545
               _Version        =   65536
               _ExtentX        =   2725
               _ExtentY        =   450
               Calculator      =   "frmCC_ColectionTambahan.frx":158F3
               Caption         =   "frmCC_ColectionTambahan.frx":15913
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_ColectionTambahan.frx":1597F
               Keys            =   "frmCC_ColectionTambahan.frx":1599D
               Spin            =   "frmCC_ColectionTambahan.frx":159E7
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1701642245
               MinValueVT      =   3801093
            End
            Begin TDBNumber6Ctl.TDBNumber tdbmaxad 
               Height          =   255
               Left            =   4860
               TabIndex        =   195
               Top             =   3270
               Visible         =   0   'False
               Width           =   1545
               _Version        =   65536
               _ExtentX        =   2725
               _ExtentY        =   450
               Calculator      =   "frmCC_ColectionTambahan.frx":15A0F
               Caption         =   "frmCC_ColectionTambahan.frx":15A2F
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_ColectionTambahan.frx":15A9B
               Keys            =   "frmCC_ColectionTambahan.frx":15AB9
               Spin            =   "frmCC_ColectionTambahan.frx":15B03
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1701642245
               MinValueVT      =   3801093
            End
            Begin TDBNumber6Ctl.TDBNumber tdbminad 
               Height          =   255
               Left            =   4860
               TabIndex        =   196
               Top             =   3600
               Visible         =   0   'False
               Width           =   1545
               _Version        =   65536
               _ExtentX        =   2725
               _ExtentY        =   450
               Calculator      =   "frmCC_ColectionTambahan.frx":15B2B
               Caption         =   "frmCC_ColectionTambahan.frx":15B4B
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_ColectionTambahan.frx":15BB7
               Keys            =   "frmCC_ColectionTambahan.frx":15BD5
               Spin            =   "frmCC_ColectionTambahan.frx":15C1F
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1701642245
               MinValueVT      =   3801093
            End
            Begin TDBNumber6Ctl.TDBNumber Tdbbalance 
               Height          =   255
               Left            =   4860
               TabIndex        =   197
               Top             =   2670
               Visible         =   0   'False
               Width           =   1545
               _Version        =   65536
               _ExtentX        =   2725
               _ExtentY        =   450
               Calculator      =   "frmCC_ColectionTambahan.frx":15C47
               Caption         =   "frmCC_ColectionTambahan.frx":15C67
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_ColectionTambahan.frx":15CD3
               Keys            =   "frmCC_ColectionTambahan.frx":15CF1
               Spin            =   "frmCC_ColectionTambahan.frx":15D3B
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
            Begin TDBNumber6Ctl.TDBNumber tdbprincipal 
               Height          =   255
               Left            =   4860
               TabIndex        =   198
               Top             =   2970
               Visible         =   0   'False
               Width           =   1545
               _Version        =   65536
               _ExtentX        =   2725
               _ExtentY        =   450
               Calculator      =   "frmCC_ColectionTambahan.frx":15D63
               Caption         =   "frmCC_ColectionTambahan.frx":15D83
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_ColectionTambahan.frx":15DEF
               Keys            =   "frmCC_ColectionTambahan.frx":15E0D
               Spin            =   "frmCC_ColectionTambahan.frx":15E57
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1701642245
               MinValueVT      =   3801093
            End
            Begin TDBNumber6Ctl.TDBNumber TDB_cur_bal 
               Height          =   255
               Left            =   4860
               TabIndex        =   199
               Top             =   2370
               Width           =   1545
               _Version        =   65536
               _ExtentX        =   2725
               _ExtentY        =   450
               Calculator      =   "frmCC_ColectionTambahan.frx":15E7F
               Caption         =   "frmCC_ColectionTambahan.frx":15E9F
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_ColectionTambahan.frx":15F0B
               Keys            =   "frmCC_ColectionTambahan.frx":15F29
               Spin            =   "frmCC_ColectionTambahan.frx":15F73
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1701642245
               MinValueVT      =   3801093
            End
            Begin TDBNumber6Ctl.TDBNumber LblMinPayment 
               Height          =   375
               Left            =   2340
               TabIndex        =   200
               Top             =   3480
               Width           =   1380
               _Version        =   65536
               _ExtentX        =   2434
               _ExtentY        =   661
               Calculator      =   "frmCC_ColectionTambahan.frx":15F9B
               Caption         =   "frmCC_ColectionTambahan.frx":15FBB
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_ColectionTambahan.frx":16027
               Keys            =   "frmCC_ColectionTambahan.frx":16045
               Spin            =   "frmCC_ColectionTambahan.frx":1608F
               AlignHorizontal =   2
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   0
               BorderStyle     =   0
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "###,###,###,##0"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   65280
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
            Begin VB.Label label1 
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               Caption         =   "No CC"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   65
               Left            =   0
               TabIndex        =   240
               Top             =   210
               Width           =   720
            End
            Begin VB.Label lblCustId 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "--"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   780
               TabIndex        =   239
               Top             =   225
               Width           =   3030
            End
            Begin VB.Label lblregion 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   780
               TabIndex        =   238
               Top             =   2880
               Width           =   1140
            End
            Begin VB.Label Label37 
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               Caption         =   "Region"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   0
               TabIndex        =   237
               Top             =   2880
               Width           =   720
            End
            Begin VB.Label LblDOB 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   780
               TabIndex        =   236
               Top             =   1110
               Width           =   1380
            End
            Begin VB.Label Label22 
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               Caption         =   "ZipCode"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   1980
               TabIndex        =   235
               Top             =   2880
               Width           =   720
            End
            Begin VB.Label lblZIP 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   2760
               TabIndex        =   234
               Top             =   2880
               Width           =   1020
            End
            Begin VB.Label Label27 
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               Caption         =   "Office Add"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   675
               Left            =   0
               TabIndex        =   233
               Top             =   2160
               Width           =   720
            End
            Begin VB.Label Label8 
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               Caption         =   "Address"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   690
               Left            =   0
               TabIndex        =   232
               Top             =   1420
               Width           =   720
            End
            Begin VB.Label Label6 
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               Caption         =   "DOB"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   0
               TabIndex        =   231
               Top             =   1140
               Width           =   720
            End
            Begin VB.Label lblID 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   780
               TabIndex        =   230
               Top             =   810
               Width           =   3030
            End
            Begin VB.Label Label5 
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               Caption         =   "ID No"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   0
               TabIndex        =   229
               Top             =   840
               Width           =   720
            End
            Begin VB.Label lblNama 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   780
               TabIndex        =   228
               Top             =   525
               Width           =   3030
            End
            Begin VB.Label Label2 
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   0
               TabIndex        =   227
               Top             =   525
               Width           =   720
            End
            Begin VB.Label Label36 
               Caption         =   "Priority"
               Height          =   195
               Left            =   5040
               TabIndex        =   226
               Top             =   3630
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
               Left            =   5670
               TabIndex        =   225
               Top             =   4050
               Visible         =   0   'False
               Width           =   435
            End
            Begin VB.Label CustId 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Risk Level"
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
               Index           =   2
               Left            =   4875
               TabIndex        =   224
               Top             =   3930
               Visible         =   0   'False
               Width           =   735
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
               Left            =   4905
               TabIndex        =   223
               Top             =   4215
               Visible         =   0   'False
               Width           =   285
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
               Left            =   1875
               TabIndex        =   222
               Top             =   165
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               Caption         =   "Curr Bal"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   11
               Left            =   3975
               TabIndex        =   221
               Top             =   2385
               Width           =   840
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblwilling 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "-------------"
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
               Height          =   270
               Left            =   3000
               TabIndex        =   220
               Top             =   3960
               Visible         =   0   'False
               Width           =   810
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Willingness"
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   10
               Left            =   5520
               TabIndex        =   219
               Top             =   3990
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.Label lblaging 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "                         "
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
               Height          =   270
               Left            =   5460
               TabIndex        =   218
               Top             =   3900
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Aging"
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   9
               Left            =   5280
               TabIndex        =   217
               Top             =   4110
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               Caption         =   "WO_Date"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   3975
               TabIndex        =   216
               Top             =   1455
               Width           =   840
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               Caption         =   "Limit"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   250
               Index           =   3
               Left            =   3980
               TabIndex        =   215
               Top             =   840
               Width           =   840
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               Caption         =   "Balance"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   250
               Index           =   6
               Left            =   3980
               TabIndex        =   214
               Top             =   225
               Width           =   840
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               Caption         =   "LPA"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   4
               Left            =   3975
               TabIndex        =   213
               Top             =   2025
               Width           =   840
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               Caption         =   "LPD"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   3975
               TabIndex        =   212
               Top             =   1720
               Width           =   840
            End
            Begin VB.Label Label18 
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               Caption         =   "Open Date"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   3975
               TabIndex        =   211
               Top             =   1140
               Width           =   840
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               Caption         =   "Princ A.P"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   250
               Index           =   8
               Left            =   3980
               TabIndex        =   210
               Top             =   520
               Visible         =   0   'False
               Width           =   840
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               Caption         =   "Max A.d"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   12
               Left            =   3960
               TabIndex        =   209
               Top             =   3270
               Visible         =   0   'False
               Width           =   840
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               Caption         =   "Min A.d"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   13
               Left            =   3960
               TabIndex        =   208
               Top             =   3600
               Visible         =   0   'False
               Width           =   840
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               Caption         =   "Balance"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   14
               Left            =   3960
               TabIndex        =   207
               Top             =   2670
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               Caption         =   "Princ A.P"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   15
               Left            =   3960
               TabIndex        =   206
               Top             =   2970
               Visible         =   0   'False
               Width           =   840
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label46 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               Caption         =   "MAP"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   60
               TabIndex        =   205
               Top             =   3240
               Width           =   960
            End
            Begin VB.Label Label47 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               Caption         =   "CYCLE"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   1260
               TabIndex        =   204
               Top             =   3240
               Width           =   960
            End
            Begin VB.Label LblMap 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0080FFFF&
               Height          =   375
               Left            =   60
               TabIndex        =   203
               Top             =   3480
               Width           =   960
            End
            Begin VB.Label LblCycle 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   375
               Left            =   1260
               TabIndex        =   202
               Top             =   3480
               Width           =   960
            End
            Begin VB.Label Label45 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H009AD6C2&
               Caption         =   "MIN.PAYMENT"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   2340
               TabIndex        =   201
               Top             =   3240
               Width           =   1380
            End
         End
         Begin VB.Frame Frame15 
            Appearance      =   0  'Flat
            BackColor       =   &H002F735C&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   0
            Left            =   240
            TabIndex        =   177
            Top             =   15
            Width           =   2895
            Begin VB.CommandButton Command1 
               Caption         =   "Command1"
               Height          =   255
               Left            =   2640
               TabIndex        =   178
               Tag             =   "0"
               Top             =   180
               Visible         =   0   'False
               Width           =   135
            End
            Begin VB.Label Label38 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Personal Data*"
               BeginProperty Font 
                  Name            =   "Arial"
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
               Left            =   480
               TabIndex        =   179
               Top             =   105
               Width           =   1755
            End
            Begin VB.Image Image1 
               Height          =   375
               Index           =   0
               Left            =   75
               Picture         =   "frmCC_ColectionTambahan.frx":160B7
               Stretch         =   -1  'True
               Tag             =   "0"
               Top             =   30
               Width           =   375
            End
         End
         Begin VB.TextBox TXTRUMUS 
            Height          =   315
            Left            =   300
            TabIndex        =   176
            Top             =   4740
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Frame frmPTP 
            Appearance      =   0  'Flat
            BackColor       =   &H00B8E2D4&
            Caption         =   "s"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   1500
            Left            =   180
            TabIndex        =   154
            Top             =   5760
            Width           =   6465
            Begin VB.CheckBox Chktenor 
               BackColor       =   &H00B8E2D4&
               Height          =   240
               Left            =   60
               TabIndex        =   159
               Top             =   1140
               Width           =   195
            End
            Begin VB.ComboBox cboPTP 
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
               ItemData        =   "frmCC_ColectionTambahan.frx":17951
               Left            =   1095
               List            =   "frmCC_ColectionTambahan.frx":17953
               TabIndex        =   158
               Top             =   165
               Width           =   2415
            End
            Begin VB.CheckBox C_Payment 
               Enabled         =   0   'False
               Height          =   255
               Left            =   3540
               TabIndex        =   157
               Top             =   180
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
               ItemData        =   "frmCC_ColectionTambahan.frx":17955
               Left            =   4995
               List            =   "frmCC_ColectionTambahan.frx":17957
               TabIndex        =   156
               Text            =   "0"
               Top             =   1200
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.ComboBox CmbBaseOn 
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
               ItemData        =   "frmCC_ColectionTambahan.frx":17959
               Left            =   2670
               List            =   "frmCC_ColectionTambahan.frx":1795B
               TabIndex        =   155
               Top             =   1140
               Visible         =   0   'False
               Width           =   1425
            End
            Begin TDBNumber6Ctl.TDBNumber txttenor 
               Height          =   255
               Left            =   1080
               TabIndex        =   160
               Top             =   1140
               Width           =   495
               _Version        =   65536
               _ExtentX        =   873
               _ExtentY        =   441
               Calculator      =   "frmCC_ColectionTambahan.frx":1795D
               Caption         =   "frmCC_ColectionTambahan.frx":1797D
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_ColectionTambahan.frx":179E9
               Keys            =   "frmCC_ColectionTambahan.frx":17A07
               Spin            =   "frmCC_ColectionTambahan.frx":17A51
               AlignHorizontal =   2
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   16384
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   "."
               DisplayFormat   =   "####0;;Null"
               EditMode        =   0
               Enabled         =   0
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   5
               MinValueVT      =   5
            End
            Begin TDBDate6Ctl.TDBDate TDBDate3 
               Height          =   285
               Left            =   3840
               TabIndex        =   161
               Top             =   840
               Width           =   1485
               _Version        =   65536
               _ExtentX        =   2619
               _ExtentY        =   494
               Calendar        =   "frmCC_ColectionTambahan.frx":17A79
               Caption         =   "frmCC_ColectionTambahan.frx":17B91
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_ColectionTambahan.frx":17BFD
               Keys            =   "frmCC_ColectionTambahan.frx":17C1B
               Spin            =   "frmCC_ColectionTambahan.frx":17C79
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
               Height          =   255
               Left            =   1095
               TabIndex        =   162
               Top             =   540
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   450
               Calculator      =   "frmCC_ColectionTambahan.frx":17CA1
               Caption         =   "frmCC_ColectionTambahan.frx":17CC1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_ColectionTambahan.frx":17D2D
               Keys            =   "frmCC_ColectionTambahan.frx":17D4B
               Spin            =   "frmCC_ColectionTambahan.frx":17D95
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1701642245
               MinValueVT      =   3801093
            End
            Begin TDBNumber6Ctl.TDBNumber Tdabamoint 
               Height          =   255
               Left            =   1095
               TabIndex        =   163
               Top             =   810
               Width           =   1410
               _Version        =   65536
               _ExtentX        =   2487
               _ExtentY        =   450
               Calculator      =   "frmCC_ColectionTambahan.frx":17DBD
               Caption         =   "frmCC_ColectionTambahan.frx":17DDD
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_ColectionTambahan.frx":17E49
               Keys            =   "frmCC_ColectionTambahan.frx":17E67
               Spin            =   "frmCC_ColectionTambahan.frx":17EB1
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
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
               ValueVT         =   5
               Value           =   0
               MaxValueVT      =   1701642245
               MinValueVT      =   3801093
            End
            Begin Threed.SSCommand SSCommand2 
               Height          =   615
               Index           =   0
               Left            =   5580
               TabIndex        =   164
               Top             =   360
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   1085
               _Version        =   196610
               PictureFrames   =   1
               Picture         =   "frmCC_ColectionTambahan.frx":17ED9
               AutoSize        =   1
               Alignment       =   8
            End
            Begin TDBDate6Ctl.TDBDate tdbptpnew 
               Height          =   285
               Left            =   3840
               TabIndex        =   165
               Top             =   540
               Width           =   1485
               _Version        =   65536
               _ExtentX        =   2619
               _ExtentY        =   494
               Calendar        =   "frmCC_ColectionTambahan.frx":18462
               Caption         =   "frmCC_ColectionTambahan.frx":1857A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_ColectionTambahan.frx":185E6
               Keys            =   "frmCC_ColectionTambahan.frx":18604
               Spin            =   "frmCC_ColectionTambahan.frx":18662
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
               ShowContextMenu =   1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "__/__/____"
               ValidateMode    =   0
               ValueVT         =   6815745
               Value           =   39876
               CenturyMode     =   0
            End
            Begin VB.Label lbltambahedit 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H009AD6C2&
               Caption         =   "Add"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   5685
               TabIndex        =   175
               Top             =   960
               Width           =   345
            End
            Begin VB.Label label1 
               BackColor       =   &H009AD6C2&
               Caption         =   "PTP"
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
               Index           =   37
               Left            =   45
               TabIndex        =   174
               Top             =   285
               Width           =   1005
            End
            Begin VB.Label label1 
               BackColor       =   &H009AD6C2&
               Caption         =   "Tenor"
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
               Index           =   44
               Left            =   240
               TabIndex        =   173
               Top             =   1140
               Width           =   870
            End
            Begin VB.Label label1 
               BackColor       =   &H009AD6C2&
               Caption         =   "Installment"
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
               Index           =   42
               Left            =   45
               TabIndex        =   172
               Top             =   840
               Width           =   1005
            End
            Begin VB.Label label1 
               BackColor       =   &H009AD6C2&
               Caption         =   "AmountPTP"
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
               Index           =   77
               Left            =   45
               TabIndex        =   171
               Top             =   540
               Width           =   1005
            End
            Begin VB.Label label1 
               BackColor       =   &H009AD6C2&
               Caption         =   "Date PTP"
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
               Index           =   0
               Left            =   2565
               TabIndex        =   170
               Top             =   840
               Width           =   1245
            End
            Begin VB.Label label1 
               AutoSize        =   -1  'True
               BackColor       =   &H009AD6C2&
               Caption         =   "Payment"
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
               Height          =   240
               Index           =   79
               Left            =   3825
               TabIndex        =   169
               Top             =   150
               Visible         =   0   'False
               Width           =   690
            End
            Begin VB.Label label1 
               BackColor       =   &H009AD6C2&
               Caption         =   "Date PTPNew"
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
               Index           =   18
               Left            =   2580
               TabIndex        =   168
               Top             =   540
               Width           =   1245
            End
            Begin VB.Label label1 
               BackColor       =   &H009AD6C2&
               Caption         =   "Disc"
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
               Index           =   75
               Left            =   4140
               TabIndex        =   167
               Top             =   1200
               Visible         =   0   'False
               Width           =   870
            End
            Begin VB.Label label1 
               BackColor       =   &H009AD6C2&
               Caption         =   "Base On"
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
               Index           =   1
               Left            =   1620
               TabIndex        =   166
               Top             =   1170
               Visible         =   0   'False
               Width           =   1005
            End
         End
         Begin VB.Frame Frame15 
            Appearance      =   0  'Flat
            BackColor       =   &H002F735C&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   3
            Left            =   180
            TabIndex        =   151
            Top             =   5280
            Width           =   2895
            Begin VB.Label Label38 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Call Actvity"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   3
               Left            =   480
               TabIndex        =   152
               Top             =   60
               Width           =   1455
            End
            Begin VB.Image Image1 
               Height          =   375
               Index           =   3
               Left            =   75
               Picture         =   "frmCC_ColectionTambahan.frx":1868A
               Stretch         =   -1  'True
               Top             =   30
               Width           =   375
            End
         End
         Begin VB.Frame Frame5 
            Appearance      =   0  'Flat
            BackColor       =   &H00B8E2D4&
            Caption         =   "PTP Jatuh Tempo"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1365
            Left            =   180
            TabIndex        =   146
            Top             =   7320
            Width           =   3465
            Begin MSComctlLib.ListView LstPayment 
               Height          =   1005
               Left            =   120
               TabIndex        =   147
               Top             =   240
               Width           =   2625
               _ExtentX        =   4630
               _ExtentY        =   1773
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
            Begin Threed.SSCommand SSCommand2 
               Height          =   615
               Index           =   2
               Left            =   2790
               TabIndex        =   148
               Top             =   270
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   1085
               _Version        =   196610
               PictureFrames   =   1
               Picture         =   "frmCC_ColectionTambahan.frx":18BD2
               AutoSize        =   1
               Alignment       =   8
            End
            Begin Threed.SSCommand SSCommand2 
               Height          =   735
               Index           =   1
               Left            =   3690
               TabIndex        =   149
               Top             =   1710
               Visible         =   0   'False
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   1296
               _Version        =   196610
               PictureFrames   =   1
               Picture         =   "frmCC_ColectionTambahan.frx":19167
               Caption         =   "&Ubah"
               Alignment       =   8
            End
            Begin VB.Label lblhapus 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00B8E2D4&
               Caption         =   "Del"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   2985
               TabIndex        =   150
               Top             =   855
               Width           =   285
            End
         End
         Begin VB.Frame Frame18 
            Appearance      =   0  'Flat
            BackColor       =   &H00B8E2D4&
            Caption         =   "Reserve PTP"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1365
            Left            =   3600
            TabIndex        =   142
            Top             =   7320
            Width           =   3090
            Begin MSComctlLib.ListView LstReserve 
               Height          =   1035
               Left            =   75
               TabIndex        =   143
               Top             =   225
               Width           =   2325
               _ExtentX        =   4101
               _ExtentY        =   1826
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
            Begin Threed.SSCommand SSCommand2 
               Height          =   615
               Index           =   3
               Left            =   2430
               TabIndex        =   144
               Top             =   210
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   1085
               _Version        =   196610
               PictureFrames   =   1
               Picture         =   "frmCC_ColectionTambahan.frx":196F0
               AutoSize        =   1
               Alignment       =   8
            End
            Begin VB.Label Label41 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00B8E2D4&
               Caption         =   "Del"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   2565
               TabIndex        =   145
               Top             =   810
               Width           =   285
            End
         End
         Begin VB.CommandButton CmdDataMapping 
            Caption         =   "&Data Mapping..."
            Height          =   435
            Left            =   4980
            Style           =   1  'Graphical
            TabIndex        =   141
            Top             =   5280
            Width           =   1635
         End
         Begin MSComctlLib.ListView LstDoubleId 
            Height          =   870
            Left            =   180
            TabIndex        =   180
            Top             =   4365
            Width           =   6480
            _ExtentX        =   11430
            _ExtentY        =   1535
            View            =   3
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
         Begin VB.TextBox txthasil 
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H000080FF&
            Height          =   375
            Left            =   3960
            TabIndex        =   241
            Top             =   3840
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.Label Label40 
            BackColor       =   &H009AD6C2&
            Caption         =   "Other Card"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            TabIndex        =   246
            Top             =   4170
            Width           =   975
         End
         Begin VB.Label Label32 
            BackColor       =   &H009AD6C2&
            Caption         =   "Coding "
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
            Height          =   240
            Left            =   3180
            TabIndex        =   245
            Top             =   0
            Width           =   735
         End
         Begin VB.Label lblaoc 
            Alignment       =   2  'Center
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
            Height          =   240
            Left            =   3900
            TabIndex        =   244
            Top             =   15
            Width           =   750
         End
         Begin VB.Label label1 
            Appearance      =   0  'Flat
            BackColor       =   &H009AD6C2&
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
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   80
            Left            =   4740
            TabIndex        =   243
            Tag             =   "0"
            Top             =   0
            Width           =   660
         End
         Begin VB.Label lblRecsource 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "--"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5400
            TabIndex        =   242
            Top             =   0
            Width           =   1290
         End
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1695
      Left            =   150
      TabIndex        =   253
      Top             =   6600
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
      TabPicture(0)   =   "frmCC_ColectionTambahan.frx":19C85
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Additional Fields"
      TabPicture(1)   =   "frmCC_ColectionTambahan.frx":19CA1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "History"
      TabPicture(2)   =   "frmCC_ColectionTambahan.frx":19CBD
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Results"
      TabPicture(3)   =   "frmCC_ColectionTambahan.frx":19CD9
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "C_NotContacted"
      Tab(3).Control(1)=   "FrmLunas"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Detail Payment"
      TabPicture(4)   =   "frmCC_ColectionTambahan.frx":19CF5
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Request Visit"
      TabPicture(5)   =   "frmCC_ColectionTambahan.frx":19D11
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      Begin VB.Frame FrmUnContacted 
         Height          =   1095
         Left            =   -74430
         TabIndex        =   297
         Top             =   8640
         Width           =   4620
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
            ItemData        =   "frmCC_ColectionTambahan.frx":19D2D
            Left            =   1245
            List            =   "frmCC_ColectionTambahan.frx":19D2F
            TabIndex        =   301
            Top             =   630
            Width           =   3285
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
            ItemData        =   "frmCC_ColectionTambahan.frx":19D31
            Left            =   1250
            List            =   "frmCC_ColectionTambahan.frx":19D33
            TabIndex        =   300
            Top             =   320
            Width           =   2340
         End
         Begin VB.CheckBox chkAppv 
            BackColor       =   &H00C5974B&
            Caption         =   "YES"
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   0
            Left            =   3600
            TabIndex        =   299
            Top             =   120
            Width           =   975
         End
         Begin VB.CheckBox chkAppv 
            BackColor       =   &H00C5974B&
            Caption         =   "NO"
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   1
            Left            =   3600
            TabIndex        =   298
            Top             =   360
            Width           =   975
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
            TabIndex        =   304
            Top             =   720
            Width           =   960
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
            TabIndex        =   303
            Top             =   360
            Width           =   1050
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
            TabIndex        =   302
            Top             =   0
            Width           =   1170
         End
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E8BE91&
         Height          =   195
         Left            =   -67515
         TabIndex        =   295
         Top             =   4425
         Width           =   255
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E8BE91&
         Height          =   195
         Left            =   -67500
         TabIndex        =   294
         Top             =   4065
         Width           =   210
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E8BE91&
         Height          =   195
         Left            =   -71130
         TabIndex        =   293
         Top             =   4035
         Width           =   240
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E8BE91&
         Height          =   195
         Left            =   -71100
         TabIndex        =   292
         Top             =   4380
         Width           =   225
      End
      Begin VB.TextBox txtResult 
         Height          =   285
         Left            =   -67560
         TabIndex        =   291
         Top             =   7620
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtResultDesc 
         Height          =   285
         Left            =   -69540
         TabIndex        =   290
         Top             =   7830
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtDiscount 
         Height          =   285
         Left            =   -70380
         TabIndex        =   289
         Top             =   7770
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00E8BE91&
         Height          =   195
         Left            =   -64290
         TabIndex        =   288
         Top             =   4065
         Width           =   210
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00E8BE91&
         Height          =   195
         Left            =   -64260
         TabIndex        =   287
         Top             =   4440
         Width           =   225
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000013&
         Height          =   3255
         Left            =   -71385
         TabIndex        =   267
         Top             =   330
         Width           =   5970
         Begin VB.Frame Frame6 
            Height          =   615
            Left            =   1275
            TabIndex        =   268
            Top             =   1455
            Visible         =   0   'False
            Width           =   3045
            Begin TDBNumber6Ctl.TDBNumber txtAmountwo_A 
               Height          =   315
               Left            =   1200
               TabIndex        =   269
               Top             =   720
               Width           =   1485
               _Version        =   65536
               _ExtentX        =   2619
               _ExtentY        =   564
               Calculator      =   "frmCC_ColectionTambahan.frx":19D35
               Caption         =   "frmCC_ColectionTambahan.frx":19D55
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmCC_ColectionTambahan.frx":19DC1
               Keys            =   "frmCC_ColectionTambahan.frx":19DDF
               Spin            =   "frmCC_ColectionTambahan.frx":19E29
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
               TabIndex        =   270
               Top             =   600
               Width           =   930
               WordWrap        =   -1  'True
            End
         End
         Begin TDBDate6Ctl.TDBDate lblLastBill 
            Height          =   300
            Left            =   3150
            TabIndex        =   271
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   529
            Calendar        =   "frmCC_ColectionTambahan.frx":19E51
            Caption         =   "frmCC_ColectionTambahan.frx":19F69
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_ColectionTambahan.frx":19FD5
            Keys            =   "frmCC_ColectionTambahan.frx":19FF3
            Spin            =   "frmCC_ColectionTambahan.frx":1A051
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
            TabIndex        =   272
            Top             =   360
            Visible         =   0   'False
            Width           =   1215
            _Version        =   65536
            _ExtentX        =   2143
            _ExtentY        =   503
            Calendar        =   "frmCC_ColectionTambahan.frx":1A079
            Caption         =   "frmCC_ColectionTambahan.frx":1A191
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_ColectionTambahan.frx":1A1FD
            Keys            =   "frmCC_ColectionTambahan.frx":1A21B
            Spin            =   "frmCC_ColectionTambahan.frx":1A279
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
            TabIndex        =   273
            Top             =   210
            Width           =   1125
            _Version        =   65536
            _ExtentX        =   1984
            _ExtentY        =   529
            Calculator      =   "frmCC_ColectionTambahan.frx":1A2A1
            Caption         =   "frmCC_ColectionTambahan.frx":1A2C1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_ColectionTambahan.frx":1A32D
            Keys            =   "frmCC_ColectionTambahan.frx":1A34B
            Spin            =   "frmCC_ColectionTambahan.frx":1A395
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
            TabIndex        =   274
            Top             =   2190
            Visible         =   0   'False
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   556
            Calculator      =   "frmCC_ColectionTambahan.frx":1A3BD
            Caption         =   "frmCC_ColectionTambahan.frx":1A3DD
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_ColectionTambahan.frx":1A449
            Keys            =   "frmCC_ColectionTambahan.frx":1A467
            Spin            =   "frmCC_ColectionTambahan.frx":1A4B1
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
            TabIndex        =   286
            Top             =   2640
            Visible         =   0   'False
            Width           =   600
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
            TabIndex        =   285
            Top             =   2820
            Visible         =   0   'False
            Width           =   75
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
            TabIndex        =   284
            Top             =   2790
            Visible         =   0   'False
            Width           =   840
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
            TabIndex        =   283
            Top             =   2520
            Visible         =   0   'False
            Width           =   645
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
            TabIndex        =   282
            Top             =   2760
            Visible         =   0   'False
            Width           =   450
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
            TabIndex        =   281
            Top             =   2760
            Visible         =   0   'False
            Width           =   735
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
            TabIndex        =   280
            Top             =   2610
            Visible         =   0   'False
            Width           =   105
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
            TabIndex        =   279
            Top             =   2460
            Visible         =   0   'False
            Width           =   690
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
            TabIndex        =   278
            Top             =   2700
            Visible         =   0   'False
            Width           =   720
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
            TabIndex        =   277
            Top             =   2250
            Visible         =   0   'False
            Width           =   1110
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
            TabIndex        =   276
            Top             =   2730
            Visible         =   0   'False
            Width           =   1200
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
            TabIndex        =   275
            Top             =   2550
            Visible         =   0   'False
            Width           =   555
         End
      End
      Begin VB.Frame Frame3 
         Height          =   555
         Left            =   -66585
         TabIndex        =   266
         Top             =   1095
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.Frame FrmLunas 
         Height          =   1215
         Left            =   -74640
         TabIndex        =   256
         Top             =   8520
         Visible         =   0   'False
         Width           =   4335
         Begin VB.CheckBox C_lunas 
            BackColor       =   &H00C5974B&
            Caption         =   "Lunas"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   390
            TabIndex        =   259
            Top             =   900
            Width           =   1455
         End
         Begin RichTextLib.RichTextBox TxtFieldName 
            Height          =   375
            Left            =   1560
            TabIndex        =   257
            Top             =   1200
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393217
            Enabled         =   -1  'True
            TextRTF         =   $"frmCC_ColectionTambahan.frx":1A4D9
         End
         Begin TDBNumber6Ctl.TDBNumber TDBTot_payment 
            Height          =   375
            Left            =   1560
            TabIndex        =   258
            Top             =   720
            Width           =   1935
            _Version        =   65536
            _ExtentX        =   3413
            _ExtentY        =   661
            Calculator      =   "frmCC_ColectionTambahan.frx":1A55B
            Caption         =   "frmCC_ColectionTambahan.frx":1A57B
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_ColectionTambahan.frx":1A5E7
            Keys            =   "frmCC_ColectionTambahan.frx":1A605
            Spin            =   "frmCC_ColectionTambahan.frx":1A64F
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
         Begin TDBDate6Ctl.TDBDate TdbLunas 
            Height          =   285
            Left            =   1560
            TabIndex        =   260
            Top             =   360
            Width           =   1350
            _Version        =   65536
            _ExtentX        =   2381
            _ExtentY        =   503
            Calendar        =   "frmCC_ColectionTambahan.frx":1A677
            Caption         =   "frmCC_ColectionTambahan.frx":1A78F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_ColectionTambahan.frx":1A7FB
            Keys            =   "frmCC_ColectionTambahan.frx":1A819
            Spin            =   "frmCC_ColectionTambahan.frx":1A877
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
         Begin VB.Label Label7 
            Caption         =   "Date of Payment"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   265
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Total Payment"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   264
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label9 
            Caption         =   "Field Name"
            Height          =   255
            Left            =   240
            TabIndex        =   263
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label17 
            Caption         =   "Label17"
            Height          =   375
            Left            =   1320
            TabIndex        =   262
            Top             =   0
            Width           =   135
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
            TabIndex        =   261
            Top             =   660
            Width           =   4215
         End
      End
      Begin VB.CheckBox C_NotContacted 
         BackColor       =   &H00C5974B&
         Height          =   270
         Left            =   -74430
         TabIndex        =   255
         Top             =   7950
         Width           =   375
      End
      Begin VB.Frame Frame4 
         Caption         =   "Emergency Contact"
         Height          =   2475
         Left            =   -72105
         TabIndex        =   254
         Top             =   825
         Width           =   4575
      End
      Begin MSComctlLib.ListView listview1 
         Height          =   5400
         Index           =   3
         Left            =   -74850
         TabIndex        =   296
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
      Begin MSComctlLib.ListView LstVisit 
         Height          =   1215
         Left            =   -74760
         TabIndex        =   305
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
         TabIndex        =   330
         Top             =   3735
         Width           =   1890
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
         TabIndex        =   329
         Top             =   4365
         Width           =   1050
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
         TabIndex        =   328
         Top             =   4065
         Width           =   975
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
         TabIndex        =   327
         Top             =   4005
         Width           =   1215
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
         TabIndex        =   326
         Top             =   4320
         Width           =   1290
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
         TabIndex        =   325
         Top             =   4365
         Width           =   765
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
         TabIndex        =   324
         Top             =   4635
         Width           =   975
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
         TabIndex        =   323
         Top             =   4995
         Width           =   780
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
         TabIndex        =   322
         Top             =   5355
         Width           =   615
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
         TabIndex        =   321
         Top             =   4395
         Width           =   1035
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
         TabIndex        =   320
         Top             =   540
         Width           =   1980
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
         TabIndex        =   319
         Top             =   825
         Width           =   1215
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
         TabIndex        =   318
         Top             =   1185
         Width           =   1290
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
         TabIndex        =   317
         Top             =   1560
         Width           =   1980
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
         TabIndex        =   316
         Top             =   1830
         Width           =   1215
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
         TabIndex        =   315
         Top             =   2190
         Width           =   1290
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
         TabIndex        =   314
         Top             =   3510
         Width           =   2025
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
         TabIndex        =   313
         Top             =   3750
         Width           =   1260
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
         TabIndex        =   312
         Top             =   4110
         Width           =   1335
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
         TabIndex        =   311
         Top             =   2535
         Width           =   1785
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
         TabIndex        =   310
         Top             =   3150
         Width           =   510
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
         TabIndex        =   309
         Top             =   2790
         Width           =   435
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
         TabIndex        =   308
         Top             =   4035
         Width           =   1260
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
         TabIndex        =   307
         Top             =   4395
         Width           =   1335
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
         TabIndex        =   306
         Top             =   7710
         Width           =   4695
      End
   End
   Begin TDBNumber6Ctl.TDBNumber TDBNumber2 
      Height          =   255
      Left            =   0
      TabIndex        =   355
      Top             =   15
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   441
      Calculator      =   "frmCC_ColectionTambahan.frx":1A89F
      Caption         =   "frmCC_ColectionTambahan.frx":1A8BF
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmCC_ColectionTambahan.frx":1A92B
      Keys            =   "frmCC_ColectionTambahan.frx":1A949
      Spin            =   "frmCC_ColectionTambahan.frx":1A993
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
   Begin VB.Frame Frame9 
      Height          =   3405
      Left            =   75
      TabIndex        =   331
      Top             =   6495
      Visible         =   0   'False
      Width           =   1755
      Begin VB.Frame Frame8 
         ForeColor       =   &H000000FF&
         Height          =   1725
         Left            =   60
         TabIndex        =   334
         Top             =   2145
         Visible         =   0   'False
         Width           =   7560
         Begin VB.OptionButton Option7 
            Caption         =   "Kantor"
            Height          =   195
            Index           =   2
            Left            =   6525
            TabIndex        =   340
            Top             =   840
            Width           =   840
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Rumah"
            Height          =   195
            Index           =   1
            Left            =   5565
            TabIndex        =   339
            Top             =   855
            Width           =   840
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Alamat Billing"
            Height          =   195
            Index           =   0
            Left            =   4125
            TabIndex        =   338
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
            TabIndex        =   337
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
            TabIndex        =   336
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
            TabIndex        =   335
            Top             =   540
            Width           =   3135
         End
         Begin TDBNumber6Ctl.TDBNumber TDBNumber1 
            Height          =   315
            Left            =   915
            TabIndex        =   341
            Top             =   870
            Width           =   615
            _Version        =   65536
            _ExtentX        =   1085
            _ExtentY        =   556
            Calculator      =   "frmCC_ColectionTambahan.frx":1A9BB
            Caption         =   "frmCC_ColectionTambahan.frx":1A9DB
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_ColectionTambahan.frx":1AA47
            Keys            =   "frmCC_ColectionTambahan.frx":1AA65
            Spin            =   "frmCC_ColectionTambahan.frx":1AAAF
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
            TabIndex        =   342
            Top             =   225
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   1005
            _Version        =   393217
            BackColor       =   16777215
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmCC_ColectionTambahan.frx":1AAD7
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
            TabIndex        =   343
            Top             =   1200
            Visible         =   0   'False
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   556
            Calendar        =   "frmCC_ColectionTambahan.frx":1AB5C
            Caption         =   "frmCC_ColectionTambahan.frx":1AC74
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_ColectionTambahan.frx":1ACE0
            Keys            =   "frmCC_ColectionTambahan.frx":1ACFE
            Spin            =   "frmCC_ColectionTambahan.frx":1AD5C
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
            TabIndex        =   344
            Top             =   870
            Visible         =   0   'False
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   556
            Calendar        =   "frmCC_ColectionTambahan.frx":1AD84
            Caption         =   "frmCC_ColectionTambahan.frx":1AE9C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmCC_ColectionTambahan.frx":1AF08
            Keys            =   "frmCC_ColectionTambahan.frx":1AF26
            Spin            =   "frmCC_ColectionTambahan.frx":1AF84
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
            TabIndex        =   345
            Top             =   1065
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   953
            _Version        =   393217
            BackColor       =   16777215
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmCC_ColectionTambahan.frx":1AFAC
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
            TabIndex        =   352
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
            TabIndex        =   351
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
            TabIndex        =   350
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
            TabIndex        =   349
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
            TabIndex        =   348
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
            Left            =   2925
            TabIndex        =   347
            Top             =   195
            Width           =   1095
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
            TabIndex        =   346
            Top             =   240
            Width           =   810
         End
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Batal"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   1395
         TabIndex        =   333
         Top             =   2085
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Tambah"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   345
         TabIndex        =   332
         Top             =   2070
         Visible         =   0   'False
         Width           =   885
      End
   End
   Begin VB.Frame FrmPayment1 
      Height          =   1365
      Left            =   1920
      TabIndex        =   247
      Top             =   8310
      Width           =   2085
      Begin VB.CheckBox Check1 
         Caption         =   "Regular Payment"
         Height          =   195
         Left            =   75
         TabIndex        =   250
         Top             =   870
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Iregular to Paid Off"
         Height          =   195
         Left            =   60
         TabIndex        =   249
         Top             =   360
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Regular to paid Off"
         Height          =   195
         Left            =   75
         TabIndex        =   248
         Top             =   285
         Visible         =   0   'False
         Width           =   1695
      End
      Begin TDBDate6Ctl.TDBDate TdbPTP 
         Height          =   255
         Left            =   60
         TabIndex        =   251
         Top             =   585
         Visible         =   0   'False
         Width           =   1440
         _Version        =   65536
         _ExtentX        =   2540
         _ExtentY        =   450
         Calendar        =   "frmCC_ColectionTambahan.frx":1B031
         Caption         =   "frmCC_ColectionTambahan.frx":1B149
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_ColectionTambahan.frx":1B1B5
         Keys            =   "frmCC_ColectionTambahan.frx":1B1D3
         Spin            =   "frmCC_ColectionTambahan.frx":1B231
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
         TabIndex        =   252
         Top             =   1065
         Visible         =   0   'False
         Width           =   285
         _Version        =   65536
         _ExtentX        =   503
         _ExtentY        =   397
         Calendar        =   "frmCC_ColectionTambahan.frx":1B259
         Caption         =   "frmCC_ColectionTambahan.frx":1B371
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmCC_ColectionTambahan.frx":1B3DD
         Keys            =   "frmCC_ColectionTambahan.frx":1B3FB
         Spin            =   "frmCC_ColectionTambahan.frx":1B459
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
   Begin VB.TextBox txtPhoneA 
      Height          =   285
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   354
      Top             =   7695
      Width           =   1905
   End
   Begin VB.TextBox txtPhone 
      Height          =   285
      Left            =   4710
      Locked          =   -1  'True
      TabIndex        =   353
      Top             =   7710
      Visible         =   0   'False
      Width           =   1920
   End
End
Attribute VB_Name = "frmCC_Colection2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_cust As ADODB.Recordset
Dim M_update As ADODB.Recordset
Dim M_objrs As ADODB.Recordset
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
Dim TglPTPNew As String
Dim vrnewdate As String
Dim KelapKelip As Integer
Public CPAString  As String
Dim StatusAccount As String

Private Sub C_Contacted_Click()
If C_Contacted.Value Then
        C_VALID.Value = False
        C_SKIP.Value = False
        C_Payment.Value = False
        C_PTP.Value = False
      '  C_POPSP.Value = False
        FrmContacted.Enabled = True
      '  cboPOPSP.Text = ""
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
     'If cboPOPSP.Text <> "" Then
     'Exit Sub
     'End If
     
      'cmbDiscount.Text = ""
   End If
End Sub
Private Sub C_PTP_Click()
If C_PTP.Value Then
       If Left(cboaccount.Text, 3) <> "ON-" Then
         cboaccount.Text = ""
       End If
       
        bcekptp = False
 '       C_VALID.Value = False
'        C_SKIP.Value = False
'        C_Contacted.Value = False
        frmPTP.Enabled = True
        FrmPayment.Enabled = True
        'cboPOPSP.Tag = 0
        Label43(2).Visible = True
       ' cboPOPSP.Text = ""
        C_Payment.Value = 1
        If UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
            SSCommand1(4).Visible = False
            Label43(2).Visible = False
            Else
            SSCommand1(4).Visible = True
            Label43(2).Visible = True
        End If
        
   Else
       bcekptp = False
       Label43(2).Visible = False
    
        'C_Payment.Value = 0
       ' CmbBaseOn.Text = ""
       ' cmbDiscount.Text = 0
        'txtPayment.Value = 0
'        TxtPtpAddr.Text = ""
 '       TxtPhonePTP.Text = ""
      '  FrmPayment.Enabled = False
        cboPTP.Text = ""
                SSCommand1(4).Visible = False
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

Private Sub cbodescskip_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
KeyAscii = 0
End If

End Sub

Private Sub cbodescvalid_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
KeyAscii = 0
End If

End Sub

Private Sub cboaccount_Click()
Dim M_COL1 As New ADODB.Recordset
If Left(cboaccount, 3) <> "ON-" Then
  C_Payment.Value = vbUnchecked
        C_PTP.Value = vbUnchecked
End If


If UCase(Left(cboaccount.Text, 2)) = "SP" Then
    C_PTP.Value = 0
    CmbBaseOn.Text = ""
    cmbDiscount.Text = ""
    txtPayment.Value = 0
    Tdabamoint.Value = 0
    TDBDate3.Value = ""
    txttenor.Value = 0
    C_Payment.Value = 1
    FrmPayment.Enabled = True
            Set M_COL1 = New ADODB.Recordset
            cmdsql = "SELECT tglincoming,tdbdatePTP,ttlPTP,discpersen,cmbbaseon,dateptp,tenor,amountptp from mgm where custid ='" & VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(1) & "'"
            M_COL1.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'            CmbBaseOn.Text = "PRINCIPLE"
            txtPayment.Value = CStr(IIf(IsNull(M_COL1!ttlptp), "", M_COL1!ttlptp))
            CmbBaseOn.Text = CStr(IIf(IsNull(M_COL1!CmbBaseOn), "", M_COL1!CmbBaseOn))
            TdbPTP.Value = CStr(IIf(IsNull(M_COL1!TdbDatePTP), "", M_COL1!TdbDatePTP))
            cmbDiscount.Text = CStr(IIf(IsNull(M_COL1!discpersen), "", M_COL1!discpersen))
            TDBDate3.Value = CStr(IIf(IsNull(M_COL1!dateptp), "", M_COL1!dateptp))
            txttenor.Value = CStr(IIf(IsNull(M_COL1!Tenor), 0, M_COL1!Tenor))
            Tdabamoint.Value = CStr(IIf(IsNull(M_COL1!amountptp), 0, M_COL1!amountptp))
End If


End Sub

Private Sub cboaccount_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub

Private Sub cbolastcall_GotFocus()
'cbolastcall.CLEAR
'Dim M_OBJRS As ADODB.Recordset
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'If UCase(mdiform1.txtlevel.text) = "AGENT" Then
'    If Left(cmbContacted.Text, 2) = "OP" Then
'    CMDSQL = " Select * from ContactedDesc where kdnoprodPresented not in('SP-SETTLE PAYMENT','PTP-PROMISE TO PAY') "
'    ElseIf Left(cboPTP.Text, 3) = "PTP" Then
'    CMDSQL = " Select * from ContactedDesc where kdnoprodPresented not in('OP-ON PROGRESS','SP-SETTLE PAYMENT') "
'    Else
'    CMDSQL = " Select * from ContactedDesc where kdnoprodPresented not in('SP-SETTLE PAYMENT')"
'    End If
' Else
'    If Left(cmbContacted.Text, 2) = "OP" Then
'    CMDSQL = " Select * from ContactedDesc where kdnoprodPresented <> 'PTP-PROMISE TO PAY' "
'    ElseIf Left(cboPTP.Text, 3) = "PTP" Then
'    CMDSQL = " Select * from ContactedDesc where kdnoprodPresented <> 'OP-ON PROGRESS' "
'    Else
'    CMDSQL = " Select * from ContactedDesc"
'    End If
' End If
'M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
'While Not M_OBJRS.EOF
'    cbolastcall.AddItem M_OBJRS("KdNoProdPresented")
'    M_OBJRS.MoveNext
'Wend
'Set M_OBJRS = Nothing
'
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from UnContactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
'While Not M_OBJRS.EOF
'    cbolastcall.AddItem M_OBJRS("KdNoProdPresented")
'    M_OBJRS.MoveNext
'Wend
'Set M_OBJRS = Nothing
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
            cmdsql = "SELECT tglincoming,tdbdatePTP,ttlPTP,discpersen,cmbbaseon,dateptp,tenor,amountptp from mgm where custid ='" & VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(1) & "'"
            M_COL1.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
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
cbodescskip.CLEAR
If Left(cboskip.Text, 2) <> "MV" Then
   Set M_objrs = New ADODB.Recordset
   M_objrs.CursorLocation = adUseClient
   M_objrs.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
         For i = 0 To 3
           cbodescskip.AddItem M_objrs("Description")
           M_objrs.MoveNext
         Next i
   Set M_objrs = Nothing
   C_Payment.Value = 0
Else
   Set M_objrs = New ADODB.Recordset
   M_objrs.CursorLocation = adUseClient
      M_objrs.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
       While Not M_objrs.EOF
           cbodescskip.AddItem M_objrs("Description")
           M_objrs.MoveNext
       Wend
   Set M_objrs = Nothing
   C_Payment.Value = 0
End If

End Sub

Private Sub cbovalid_Click()
Dim i As Integer
cbodescvalid.CLEAR
If Left(cbovalid.Text, 2) = "NA" Then
        cbodescvalid.Enabled = True
'        CmbBaseOn.Text = ""
'        txtPayment.Text = 0
'        cmbDiscount.Text = ""
'        TdbPTP.Text = ""
'        TdbDatePTP.Text = ""
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
          M_objrs.Open "Select * from DescContacted WHERE id >= 13 ", M_OBJCONN, adOpenDynamic, adLockOptimistic
        While Not M_objrs.EOF
            cbodescvalid.AddItem M_objrs("Description")
            M_objrs.MoveNext
        Wend
        C_Payment.Value = 0
        Set M_objrs = Nothing
'        FrmPayment.Enabled = False
Else
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
          M_objrs.Open "Select * from DescunContacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
        While Not M_objrs.EOF
            cbodescvalid.AddItem M_objrs("Description")
            M_objrs.MoveNext
        Wend
        C_Payment.Value = 0
        Set M_objrs = Nothing
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
Dim Rs As New ADODB.Recordset
Rs.CursorLocation = adUseClient
Rs.Open "select TYPE from TBLNEGOPTP where custid='" & lblcustid.Caption & "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
If Rs.BOF And Rs.EOF Then
Else
    While Not Rs.EOF
        If Rs!Type = "SCH" Then
            adaSCH = True
        ElseIf Rs!Type = "REG" Then
            adaREG = True
        ElseIf Rs!Type = "PO" Then
            adaPO = True
        End If
        Rs.MoveNext
    Wend
End If
Set Rs = Nothing
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

Private Sub Chktenor_Click()
If Chktenor.Value = 1 Then
    txttenor.Enabled = True
    txttenor.BackColor = vbWhite
Else
    txttenor.Enabled = False
    txttenor.BackColor = &H4000&
    Chktenor.Value = 0
End If


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
cmbDescCon.CLEAR

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
   Set M_objrs = New ADODB.Recordset
   M_objrs.CursorLocation = adUseClient
     M_objrs.Open "Select * from DescContacted where id <= 12", M_OBJCONN, adOpenDynamic, adLockOptimistic
    While Not M_objrs.EOF
        cmbDescCon.AddItem M_objrs("Description")
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
    C_Payment.Value = 0
   ' FrmPayment.Enabled = False
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
           ' FrmPayment.Enabled = False
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
        cmdsql = "SELECT tglincoming,tdbdatePTP,ttlPTP,discpersen,cmbbaseon,dateptp,tenor, amountptp from mgm where custid ='" & VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(1) & "'"
        m_cust.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
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

Set M_objrs = Nothing
End Sub

Private Sub cmbContacted_KeyDown(KeyCode As Integer, Shift As Integer)

cmbContacted.Text = ""
Exit Sub
End Sub

Private Sub cmbDescCon_GotFocus()
'DESCRIPTION CONTACTED
Dim i As Integer
cmbDescCon.CLEAR
If Left(cmbContacted.Text, 2) = "RP" Then
    cmbDescCon.Enabled = True
    CmbBaseOn.Text = ""
    txtPayment.Text = 0
    cmbDiscount.Text = ""
    TdbPTP.Text = ""
    TdbDatePTP.Text = ""
   Set M_objrs = New ADODB.Recordset
   M_objrs.CursorLocation = adUseClient
     M_objrs.Open "Select * from DescContacted where id <= 12", M_OBJCONN, adOpenDynamic, adLockOptimistic
    While Not M_objrs.EOF
        cmbDescCon.AddItem M_objrs("Description")
        M_objrs.MoveNext
    Wend
    C_Payment.Value = 0
   ' FrmPayment.Enabled = False
    Set M_objrs = Nothing
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
'            FrmPayment.Enabled = False
    Else
    If Left(cmbContacted.Text, 2) = "OP" Then
            cmbDescCon.Enabled = False
            CmbBaseOn.Text = ""
            txtPayment.Text = 0
            cmbDiscount.Text = ""
            TdbPTP.Text = ""
            TdbDatePTP.Text = ""
            C_Payment.Value = 0
           ' FrmPayment.Enabled = False
      Else
      
    If Left(cmbContacted.Text, 2) = "PO" Or Left(cmbContacted.Text, 2) = "SP" Then
            cmbDescCon.Enabled = False
            C_Payment.Value = 1
            FrmPayment.Enabled = True
Set m_cust = New ADODB.Recordset
m_cust.CursorLocation = adUseClient
cmdsql = "SELECT tglincoming,tdbdatePTP,ttlPTP,discpersen,cmbbaseon from mgm where custid ='" & VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(1) & "'"
    m_cust.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
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

Set M_objrs = Nothing
End Sub

Private Sub cmbDescCon_KeyDown(KeyCode As Integer, Shift As Integer)

cmbDescCon.Text = ""
Exit Sub
End Sub

Private Sub cmbDescCon_KeyPress(KeyAscii As Integer)
If Not KeyAscii = 13 Then
KeyAscii = 0
End If

End Sub

Private Sub cmbDescUn_GotFocus()
Dim i As Integer
cmbDescUn.CLEAR
If Left(cmbUncontacted.Text, 2) = "NA" Then
        cmbDescUn.Enabled = True
'        CmbBaseOn.Text = ""
'        txtPayment.Text = 0
'        cmbDiscount.Text = ""
'        TdbPTP.Text = ""
'        TdbDatePTP.Text = ""
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
          M_objrs.Open "Select * from DescContacted WHERE id >= 13 ", M_OBJCONN, adOpenDynamic, adLockOptimistic
        While Not M_objrs.EOF
            cmbDescUn.AddItem M_objrs("Description")
            M_objrs.MoveNext
        Wend
        C_Payment.Value = 0
        Set M_objrs = Nothing
'        FrmPayment.Enabled = False
Else
If Left(cmbUncontacted.Text, 2) <> "MV" Then
   Set M_objrs = New ADODB.Recordset
   M_objrs.CursorLocation = adUseClient
   M_objrs.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
         For i = 0 To 3
           cmbDescUn.AddItem M_objrs("Description")
           M_objrs.MoveNext
         Next i
   Set M_objrs = Nothing
   C_Payment.Value = 0
Else
   Set M_objrs = New ADODB.Recordset
   M_objrs.CursorLocation = adUseClient
'   If kontak = True Then
'        m_objrs.Open "Select * from DescUncontacted where ", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    Else
      M_objrs.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    End If
       While Not M_objrs.EOF
           cmbDescUn.AddItem M_objrs("Description")
           M_objrs.MoveNext
       Wend
   Set M_objrs = Nothing
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
Dim M_objrs As New ADODB.Recordset
'If cmbDiscount.Text = "" Then
'    Exit Sub
'End If

M_objrs.CursorLocation = adUseClient
M_objrs.Open "Select * from tbldiscount where Description = '" + cmbDiscount.Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If M_objrs.RecordCount <> 0 Then
    TdbDatePTP.Value = MDIForm1.TDBDate1.Value + IIf(IsNull(M_objrs!hari), 7, M_objrs!hari)
Else
    TdbDatePTP.Value = MDIForm1.TDBDate1.Value + 7
End If
If cmbDiscount.Text = "0" Or cmbDiscount.Text = "" Then
    If CmbBaseOn.Text = "PRINCIPLE" Then
        txtPayment.Value = LblPrompA.Value
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
                txtdiscount.Text = CStr((cmbDiscount.Text) / 100)
                txtPayment.Value = lblAmount.Value - (CCur(txtdiscount.Text) * lblAmount.Value)
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

Private Sub CmbPhone_Click()
    CmbPhone.Locked = True
    If CmbPhone.Text = "Add" Then
        Frm_Tambah_Telp.Show vbModal
    End If
End Sub

Private Sub CmbPhone_DropDown()
    CmbPhone.Locked = False
End Sub

Private Sub CmbPhone_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmbUncontacted_Click()
'DESCRIPTION UNCONTACTED
Dim i As Integer
cmbDescUn.CLEAR
If Left(cmbUncontacted.Text, 2) = "NA" Then
        cmbDescUn.Enabled = True
'        CmbBaseOn.Text = ""
'        txtPayment.Text = 0
'        cmbDiscount.Text = ""
'        TdbPTP.Text = ""
'        TdbDatePTP.Text = ""
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
          M_objrs.Open "Select * from DescContacted WHERE id >= 13 ", M_OBJCONN, adOpenDynamic, adLockOptimistic
        While Not M_objrs.EOF
            cmbDescUn.AddItem M_objrs("Description")
            M_objrs.MoveNext
        Wend
        C_Payment.Value = 0
        Set M_objrs = Nothing
'        FrmPayment.Enabled = False
Else
If Left(cmbUncontacted.Text, 2) <> "MV" Then
   Set M_objrs = New ADODB.Recordset
   M_objrs.CursorLocation = adUseClient
   M_objrs.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
         For i = 0 To 3
           cmbDescUn.AddItem M_objrs("Description")
           M_objrs.MoveNext
         Next i
   Set M_objrs = Nothing
   C_Payment.Value = 0
Else
   Set M_objrs = New ADODB.Recordset
   M_objrs.CursorLocation = adUseClient
'   If kontak = True Then
'        m_objrs.Open "Select * from DescUncontacted where ", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    Else
      M_objrs.Open "Select * from DescUncontacted", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    End If
       While Not M_objrs.EOF
           cmbDescUn.AddItem M_objrs("Description")
           M_objrs.MoveNext
       Wend
   Set M_objrs = Nothing
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
LstPayment.ColumnHeaders.ADD 2, , "ID", 1
LstPayment.ColumnHeaders.ADD 3, , "DATE", 1100
LstPayment.ColumnHeaders.ADD 4, , "PAYMENT", 30 * TXT
LstPayment.ColumnHeaders.ADD 5, , "TYPE", 30 * TXT
LstPayment.ColumnHeaders.ADD 6, , "INPUT DATE", 15 * TXT

LstReserve.ColumnHeaders.ADD 1, , "", 0 * TXT
LstReserve.ColumnHeaders.ADD 2, , "ID", 1
LstReserve.ColumnHeaders.ADD 3, , "DATE", 1100
LstReserve.ColumnHeaders.ADD 4, , "PAYMENT", 30 * TXT
LstReserve.ColumnHeaders.ADD 5, , "TYPE", 30 * TXT
LstReserve.ColumnHeaders.ADD 6, , "INPUT DATE", 15 * TXT

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

Private Sub CmdDataMapping_Click()
    FrmDataMapping.Show vbModal
End Sub

Private Sub CmdDeletePelunasan_Click()
Dim m_msgbox As Variant
If ListView1(0).ListItems.Count = 0 Then
    Exit Sub
End If
m_msgbox = MsgBox("Yakin Akan Dihapus...!!! ", vbCritical + vbOKCancel, "Peringatan")
If m_msgbox = vbOK Then
    M_OBJCONN.Execute "Delete from tbllunas where id = " + ListView1(0).SelectedItem.SubItems(4) + ""
    ListView1(0).ListItems.Remove ListView1(0).SelectedItem.Index
    MsgBox "Done"
    Call isi_datapayment
End If
End Sub

Private Sub CmdHapusRemarks_Click()
    Dim cmdsql As String
    Dim a As String
    
    If ListView1(1).ListItems.Count = 0 Then
        MsgBox "Tidak ada data remarks yang akan dihapus!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    a = MsgBox("Yakin data: " & ListView1(1).SelectedItem.SubItems(1) & " akan dihapus?", vbYesNo + vbQuestion, "Konfirmasi")
    
    If a = vbNo Then
        Exit Sub
    End If
    
    cmdsql = "delete from mgm_hst where id='"
    cmdsql = cmdsql + Trim(ListView1(1).SelectedItem.SubItems(7)) + "'"
    
    M_OBJCONN.Execute cmdsql
    
    ListView1(1).ListItems.Remove ListView1(1).SelectedItem.Index
End Sub

Private Sub CmdUnlock_Click()
    '@@ 01/02/2011 UnLock Data Oleh agent
    Dim a As String
    Dim ID As String
    Dim M_objrs As ADODB.Recordset
    Dim m_objrs_cekid As ADODB.Recordset
    Dim cmdsql As String
    Dim UpdateDtCloseSession As String
    Dim m_objrs_ambilTL As ADODB.Recordset
    Dim cmdsql_ambilTL As String
    
    Dim pesan As String
    Dim TglLock As String
    Dim StartLock As String
    Dim EndLock As String
    Dim AccLock As String
    Dim Status_lock As String
    
    'Cek dulu apakah yang login agent?
    If UCase(Trim(MDIForm1.txtlevel.Text)) <> "AGENT" Then
        MsgBox "Unlock data ini hanya untuk AGENT!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    a = MsgBox("Anda yakin akan melakukan Unlock Data?", vbYesNo + vbQuestion, "Konfirmasi")
    If a = vbNo Then
        Exit Sub
    End If
        
    'Cek apakah ada data yang sedang di lock?
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    cmdsql = "select * from usertbl where userid='"
    cmdsql = cmdsql + Trim(MDIForm1.txtusername.Text) + "'"
    M_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs("lockdarispv") = "" And M_objrs("lock_entry_lpd") = "" And M_objrs("lockmarkup") = "" Then
        MsgBox "Tidak ada data yang akan di unlock!", vbOKOnly + vbInformation, "Informasi"
        Set M_objrs = Nothing
        Exit Sub
    End If
    Set M_objrs = Nothing
    
    'Cari id data yang sedang di lock
    cmdsql = "select *,now() as tanggal_sekarang from tbltemplockacc_current where id in "
    cmdsql = cmdsql + "(select max(idlock) as idlock from tblperformpersessionlock where agent='"
    cmdsql = cmdsql + Trim(MDIForm1.txtusername.Text) + "')"
    
    Set m_objrs_cekid = New ADODB.Recordset
    m_objrs_cekid.CursorLocation = adUseClient
    m_objrs_cekid.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    ID = Trim(m_objrs_cekid("id"))
    TglLock = Format(m_objrs_cekid("date_lock"), "yyyy-mm-dd hh:mm:ss")
    StartLock = Format(m_objrs_cekid("start_lock"), "yyyy-mm-dd hh:mm:ss")
    EndLock = Format(m_objrs_cekid("end_lock"), "yyyy-mm-dd hh:mm:ss")
    AccLock = Trim(IIf(IsNull(m_objrs_cekid("account_lock")), "", m_objrs_cekid("account_lock")))
    Status_lock = Trim(m_objrs_cekid("status_lock"))
    
    
    'Catat ke dalam log
    cmdsql = "insert into log_unlock_agent (script_lock,date_lock,"
    cmdsql = cmdsql + "start_lock,end_lock,account_lock,lock_by,f_locked,tgl_unlock,agent_unlock,status_lock,id) values ('"
    cmdsql = cmdsql + Trim(IIf(IsNull(m_objrs_cekid("script_lock")), "", m_objrs_cekid("script_lock"))) + "','"
    cmdsql = cmdsql + Format(m_objrs_cekid("date_lock"), "yyyy-mm-dd hh:mm:ss") + "','"
    cmdsql = cmdsql + Format(m_objrs_cekid("start_lock"), "yyyy-mm-dd hh:mm:ss") + "','"
    cmdsql = cmdsql + Format(m_objrs_cekid("end_lock"), "yyyy-mm-dd hh:mm:ss") + "','"
    cmdsql = cmdsql + Trim(IIf(IsNull(m_objrs_cekid("account_lock")), "", m_objrs_cekid("account_lock"))) + "','"
    cmdsql = cmdsql + Trim(IIf(IsNull(m_objrs_cekid("lock_by")), "", m_objrs_cekid("lock_by"))) + "','"
    cmdsql = cmdsql + Trim(IIf(IsNull(m_objrs_cekid("f_locked")), "", m_objrs_cekid("f_locked"))) + "','"
    cmdsql = cmdsql + Format(m_objrs_cekid("tanggal_sekarang"), "yyyy-mm-dd hh:mm:ss") + "','"
    cmdsql = cmdsql + Trim(MDIForm1.txtusername.Text) + "','"
    cmdsql = cmdsql + Trim(m_objrs_cekid("status_lock")) + "','"
    cmdsql = cmdsql + Trim(ID) + "')"
    
    M_OBJCONN.Execute cmdsql
    
    'Bikin pesan ke TL,jika lock datanya sudah di unlock oleh agent
    pesan = vbCrLf + "INFORMASI OLEH SISTEM : " + vbCrLf
    pesan = pesan + "Agent: " + MDIForm1.txtusername.Text + vbCrLf
    pesan = pesan + "Melakukan Unlock data untuk accountnya sendiri." + vbCrLf
    pesan = pesan + "Berikut informasi lock data yang di unlock:" + vbCrLf
    pesan = pesan + "------------------------------------------------" + vbCrLf
    pesan = pesan + "Tgl.Lock data :" + StartLock + vbCrLf
    pesan = pesan + "Start.Lock data:" + EndLock + vbCrLf
    pesan = pesan + "Account yang di lock:" + AccLock + vbCrLf
    pesan = pesan + "Status yang di lock:" + Status_lock + vbCrLf
    pesan = pesan + "------------------------------------------------" + vbCrLf
    pesan = pesan + "Terima Kasih" + vbCrLf
    pesan = pesan + "Message Created automatic by system"
    
    MsgBox "Silahkan tunggu sebentar! Setelah menekan tombol OK ini, sistem akan melakukan unlock data. Harap Tunggu hingga muncul pesan Unlock data berhasil!", vbOKOnly + vbInformation, "Informasi"
    
    'Pindahkan data ke tabel tblperformpersessionlock
    DoEvents
    UpdateDtCloseSession = "update tblperformpersessionlock set f_ceksekrg=a.f_cek_akhir ,"
    UpdateDtCloseSession = UpdateDtCloseSession + " tgl_f_ceksekrg=a.tgl_akhir,endlock='" + CStr(Format(m_objrs_cekid("tanggal_sekarang"), "yyyy-mm-dd hh:mm:ss")) + "' from "
    UpdateDtCloseSession = UpdateDtCloseSession + " (select mgm.custid as custid_mgm,mgm.agent as agent_mgm,"
    UpdateDtCloseSession = UpdateDtCloseSession + " tblperformpersessionlock.custid as custid_lock,"
    UpdateDtCloseSession = UpdateDtCloseSession + " tblperformpersessionlock.agent as agent_lock,"
    UpdateDtCloseSession = UpdateDtCloseSession + " tblperformpersessionlock.idlock as id_lock,"
    UpdateDtCloseSession = UpdateDtCloseSession + " mgm.f_cek_new as f_cek_akhir, mgm.tglcall as tgl_akhir"
    UpdateDtCloseSession = UpdateDtCloseSession + " from tblperformpersessionlock inner join mgm "
    UpdateDtCloseSession = UpdateDtCloseSession + " on mgm.custid=tblperformpersessionlock.custid "
    UpdateDtCloseSession = UpdateDtCloseSession + " and mgm.agent=tblperformpersessionlock.agent) as a "
    UpdateDtCloseSession = UpdateDtCloseSession + " where tblperformpersessionlock.custid=a.custid_mgm and tblperformpersessionlock.agent=a.agent_mgm and a.id_lock='"
    UpdateDtCloseSession = UpdateDtCloseSession + Trim(ID) + "' and tblperformpersessionlock.agent='"
    UpdateDtCloseSession = UpdateDtCloseSession + Trim(MDIForm1.txtusername.Text) + "'"
    M_OBJCONN.Execute UpdateDtCloseSession
                    
    Set m_objrs_cekid = Nothing
    
    cmdsqlserver = "update usertbl set dilockoleh='Release by:" + Trim(MDIForm1.txtlevel.Text) + "',"
    cmdsqlserver = cmdsqlserver + "lockdarispv=null,lock_entry_lpd=null,fromaccount=null,"
    cmdsqlserver = cmdsqlserver + "lockmarkup=null,lockdarispvbuattl=null where userid='"
    cmdsqlserver = cmdsqlserver + Trim(MDIForm1.txtusername.Text) + "'"
    M_OBJCONN.Execute cmdsqlserver
    
    'Berikan pesan ke TL-nya
    cmdsql_ambilTL = "select * from usertbl where userid='"
    cmdsql_ambilTL = cmdsql_ambilTL + Trim(MDIForm1.txtusername.Text) + "'"
    Set m_objrs_ambilTL = New ADODB.Recordset
    m_objrs_ambilTL.CursorLocation = adUseClient
    m_objrs_ambilTL.Open cmdsql_ambilTL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    cmdsql = "insert into msgtbl  (recipient, datetime, sender, sentfrom, msg) VALUES ('"
    cmdsql = cmdsql + Trim(m_objrs_ambilTL("team")) + "','"
    cmdsql = cmdsql + CStr(Format(Now, "yyyymmdd")) + "','"
    cmdsql = cmdsql + Trim(MDIForm1.txtusername.Text) + "','"
    cmdsql = cmdsql + CStr(MDIForm1.Winsock1.LocalIP) + "','"
    cmdsql = cmdsql + Trim(pesan) + "')"
    M_OBJCONN.Execute cmdsql
    
    Set m_objrs_ambilTL = Nothing
    
    MsgBox "Data anda berhasil di Unlock!", vbOKOnly + vbInformation, "Informasi"
    VIEW_MGMDATA.ListView1.ListItems.CLEAR
End Sub

Private Sub Command1_Click()
     If Command1.tag = 0 Then
        TdbBalance.Visible = True
        tdbprincipal.Visible = True
        Label11(14).Visible = True
        Label11(15).Visible = True
        Command1.tag = 1
        LblPrompA.Visible = True
        Label11(8).Visible = True
        Else
        TdbBalance.Visible = False
        tdbprincipal.Visible = False
        Label11(14).Visible = False
        Label11(15).Visible = False
        Label11(8).Visible = False
        Command1.tag = 0
        LblPrompA.Visible = False
        End If
        
End Sub

Private Sub Command2_Click()
'Load FrmSendSMS
'FrmSendSMS.Show vbModal
'@@ 09022011, diubah formnya
FrmInboXSms.Show vbModal
End Sub

Private Sub Form_Load()
If UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
    SSCommand1(4).Visible = False
    Command1.Visible = False
    

ElseIf UCase(MDIForm1.txtlevel.Text) = "SUPERVISOR" Or UCase(MDIForm1.txtlevel.Text) = "ADMIN" Or UCase(MDIForm1.txtlevel.Text) = "ADMINISTRATOR" Then
        SSCommand1(4).Visible = True
        Command1.Visible = False
        CmdHapusRemarks.Visible = True
End If



FrmCC_Colection.Left = 10
FrmCC_Colection.Top = 20

'cek list pelunasan
Dim i, iIndex As Integer
Dim sKata, cCombo As String


'------->>>  setting No Visit  <<<---------------

Text1.Text = Format(Now, "yymmddhhmmss")
TDBDate1.Value = Now
'If UCase(Left(mdiform1.txtlevel.text, 5)) = "ADMIN" Or UCase(Left(mdiform1.txtlevel.text, 5)) = "SUPER" Then
If UCase(Left(MDIForm1.txtlevel.Text, 5)) = "ADMIN" Then
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
    TdbBalance.Visible = False
        tdbprincipal.Visible = False
        Label11(14).Visible = False
        Label11(15).Visible = False
        
        'aktifkan recsource @@ 160610
        Label1(80).Visible = True
        lblRecsource.Visible = True
End If

If UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
        C_lunas.Enabled = False
        TdbLunas.Enabled = False
        chkAppv(0).Enabled = False
        chkAppv(1).Enabled = False
        TDBTot_payment.Enabled = False
        TxtFieldName.Enabled = False
        CmdDeletePelunasan.Enabled = False
         ' Tampilkan PRincipal
        SSCommand2(3).Enabled = False
        SSCommand2(2).Enabled = False
        lblhapus.Enabled = False
        Label41.Enabled = False
        LblPrompA.Visible = True
        Label11(8).Visible = True
        TdbBalance.Visible = False
        tdbprincipal.Visible = False
        Label11(14).Visible = False
        Label11(15).Visible = False
ElseIf UCase(MDIForm1.txtlevel.Text) = "TEAMLEADER" Then
        txtHomeAdd1.ReadOnly = False
        txtHomeAdd2.ReadOnly = False
        txtOfficeAdd1.ReadOnly = False
        txtOfficeAdd2.ReadOnly = False
        txtMobileAdd1.ReadOnly = False
        txtMobileAdd2.ReadOnly = False
        SSCommand2(3).Enabled = False
        SSCommand2(2).Enabled = True
        lblhapus.Enabled = False
        Label41.Enabled = False
        Command1.Visible = False
         ' Tampilkan PRincipal
        LblPrompA.Visible = True
        Label11(8).Visible = True
        TdbBalance.Visible = False
        tdbprincipal.Visible = False
        Label11(14).Visible = False
        Label11(15).Visible = False
       
Else ' utk SPV tampilkan no telp
        txtHomeAdd1.ReadOnly = False
        txtHomeAdd2.ReadOnly = False
        txtOfficeAdd1.ReadOnly = False
        txtOfficeAdd2.ReadOnly = False
        txtMobileAdd1.ReadOnly = False
        txtMobileAdd2.ReadOnly = False
        SSCommand2(3).Enabled = True
        SSCommand2(2).Enabled = True
        lblhapus.Enabled = True
        Label41.Enabled = True
        
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
        LblPrompA.Visible = True
        Label11(8).Visible = True
        'aktifkan recsource @@ 160610
        Label1(80).Visible = True
        lblRecsource.Visible = True
        
End If
 
 '  FrmContacted.Enabled = False
   FrmUnContacted.Enabled = False
   'FrmPayment.Enabled = False
   
    Call headerDatePayment
    Call headerCustid_Double
    Call HEADER_HISTORY
    Call HEADER_HISTORY_PAID
    Call HEADER_RequestVisit
    'Call HEADER_SendSMS
   
    Call show_cust
    Call VisitNo
'    Call isi_lastcall
    
    If UCase(MDIForm1.txtlevel.Text) = "TEAMLEADER" Or UCase(MDIForm1.txtlevel.Text) = "SUPERVISOR" Or UCase(MDIForm1.txtlevel.Text) = "ADMINISTRATOR" Then
        Call aktifphone
    End If
    
    If UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
        Call aktifphoneAGENT
    End If
    
    '@@14022011
    Call CekSms
    
    '@@ 08032011 Cek Data Mapping
    Call CekDataMapping
        
  '  SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
SSTab1.Tab = 0
cmbDateSch.Value = Now
cmbDateSch.Value = ""
'CONTACTED
CmbBaseOn.AddItem "PRINCIPLE"
CmbBaseOn.AddItem "TOTAL AMOUNT"


'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from tblvalid", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    While Not M_OBJRS.EOF
'        cbovalid.AddItem M_OBJRS!KdNoProdPresented
'        M_OBJRS.MoveNext
'    Wend
'    Set M_OBJRS = Nothing
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
M_objrs.Open "Select * from tblptp where KdNoProdPresented not like 'PTP-PAID%' ", M_OBJCONN, adOpenDynamic, adLockOptimistic
    While Not M_objrs.EOF
        cboPTP.AddItem M_objrs!KdNoProdPresented
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
    
'    Set M_OBJRS = New ADODB.Recordset
'    M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from tblskip", M_OBJCONN, adOpenDynamic, adLockOptimistic
'    While Not M_OBJRS.EOF
'        cboskip.AddItem M_OBJRS!KdNoProdPresented
'        M_OBJRS.MoveNext
'    Wend
'    Set M_OBJRS = Nothing

    
    
    
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open "Select * from popspdesc ", M_OBJCONN, adOpenDynamic, adLockOptimistic
    While Not M_objrs.EOF
        cboPOPSP.AddItem M_objrs!KdNoProdPresented
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
'SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
''M_OBJRS.Open "Select * from ContactedDesc where KdNoProdPresented not like 'ptp%'", M_OBJCONN, adOpenDynamic, adLockOptimistic
'
'If UCase(mdiform1.txtlevel.text) = "AGENT" Then
'M_OBJRS.Open "Select * from contacteddesc where KdNoProdPresented not like 'ptp%' and KdNoProdPresented <>'SP-SETTLE PAYMENT' ", M_OBJCONN, adOpenDynamic, adLockOptimistic
'Else
'M_OBJRS.Open "Select * from contacteddesc where KdNoProdPresented not like 'ptp%'", M_OBJCONN, adOpenDynamic, adLockOptimistic
'End If
'    While Not M_OBJRS.EOF
'    '----tambahan 05 Maret 2007----'
'         scombo = M_OBJRS("KdNoProdPresented")
'            sKata = cmbContacted.Text
'            ' initialisasi index
'            If scombo = "BP-BROKEN PROMISE" Or scombo = "PTP-PROMISE TO PAY" Or scombo = "RP-REFUSE PAYMENT" Then
'                  iIndex = 1
'            ElseIf scombo = "POP-PROGRESS OF PAYMENT" Then
'                  iIndex = 2
'            ElseIf scombo = "SP-SETTLE PAYMENT" Then
'                  iIndex = 3
'            Else
'                  iIndex = 4
'            End If
'
'            ' saring tampilan
'            If iIndex = 1 Then
'               If iIndex = 4 Or sKata = "POP-PROGRESS OF PAYMENT" Or sKata = "SP-SETTLED PAYMENT" Then
'                  'lewat boo
'               Else
'                    If scombo = "BP-BROKEN PROMISE" And UCase(mdiform1.txtlevel.text) = "AGENT" Then
'                    Else
'                        cmbContacted.AddItem scombo
'                    End If
'               End If
'            ElseIf iIndex = 2 Then
'               If iIndex = 1 Or iIndex = 4 Or Left(sKata, 2) = "SP" Then
'                  'lewat boo
'               Else
'                  cmbContacted.AddItem scombo
'               End If
'            ElseIf iIndex = 3 Then
'                If UCase(mdiform1.txtlevel.text) = "AGENT" Then
'                Else
'                  cmbContacted.AddItem scombo
'                End If
'            Else
'                  If sKata = "BP-BROKEN PROMISE" Or sKata = "PTP-PROMISE TO PAY" Or sKata = "POP-PROGRESS OF PAYMENT" Or sKata = "SP-SETTLED PAYMENT" Then
'                     'lewat boo
'                  Else
'                     cmbContacted.AddItem scombo
'                  End If
'            End If
'            M_OBJRS.MoveNext
'    Wend
'
'
'Set M_OBJRS = Nothing

'If Left(cmbContacted.Text, 2) = "SP" Then
'    'C_Contacted.Enabled = False
'    'cmbContacted.Enabled = False
'    C_NotContacted.Enabled = False
'End If

'If Left(cmbContacted.Text, 3) = "POP" Then
'    'C_Contacted.Enabled = False
'    'cmbContacted.Enabled = False
'    C_NotContacted.Enabled = False
'End If

'UNCONTACTED
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
''If kontak = True Then
''    M_OBJRS.Open "Select * from UnContactedDesc where KdNoProdPresented IN ('NBP-NO BODY PICKUP','NA-NOT AVAILABLE')", M_OBJCONN, adOpenDynamic, adLockOptimistic
''Else
''    M_OBJRS.Open "Select * from UnContactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
''End If
'If kontak = True Then
'    M_OBJRS.Open "Select * from uncontacteddesc where kdnoprodpresented IN ('NBP-NO BODY PICKUP','NA-NOT AVAILABLE')", M_OBJCONN, adOpenDynamic, adLockOptimistic
'ElseIf Left(VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(8), 2) = "NA" Then
'    M_OBJRS.Open "Select * from uncontacteddesc  where kdnoprodpresented  IN ('NBP-NO BODY PICKUP','NA-NOT AVAILABLE')", M_OBJCONN, adOpenDynamic, adLockOptimistic
'Else
'    M_OBJRS.Open "Select * from uncontacteddesc ", M_OBJCONN, adOpenDynamic, adLockOptimistic
'End If
'    While Not M_OBJRS.EOF
'        cmbUncontacted.AddItem M_OBJRS("KdNoProdPresented")
'        'cmbDescUn.AddItem M_OBJRS("nmNoProdPresented")
'        M_OBJRS.MoveNext
'    Wend
'Set M_OBJRS = Nothing

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
'Set M_OBJRS = New ADODB.Recordset
'M_OBJRS.CursorLocation = adUseClient
'M_OBJRS.Open "Select * from stsnextact", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'While Not M_OBJRS.EOF
'    cmbNextAct.AddItem M_OBJRS("NmStsNextAct")
'    M_OBJRS.MoveNext
'Wend
'Set M_OBJRS = Nothing
'untuk 108

'@@ 24 May 2012 Akses 108, untuk agent tertentu saja
Dim M_objrs_108 As ADODB.Recordset
cmdsql = "select sts_108 from usertbl where userid='"
cmdsql = cmdsql + CStr(MDIForm1.txtusername.Text) + "' and sts_108='1'"
Set M_objrs_108 = New ADODB.Recordset
M_objrs_108.CursorLocation = adUseClient
M_objrs_108.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
If M_objrs_108.RecordCount > 0 Then
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open "Select * from tbllayanantelkom", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    While Not M_objrs.EOF
        CmbPhone.AddItem IIf(IsNull(M_objrs("Nolayanan")), "", M_objrs("Nolayanan"))
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
End If
Set M_objrs_108 = Nothing

'@@25052012 Jika yang login Admin,Superviso
If UCase(MDIForm1.txtlevel.Text) = "SUPERVISOR" Or _
   UCase(MDIForm1.txtlevel.Text) = "ADMIN" Or _
   UCase(MDIForm1.txtlevel.Text) = "ADMINISTRATOR" Then
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open "Select * from tbllayanantelkom", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    While Not M_objrs.EOF
        CmbPhone.AddItem IIf(IsNull(M_objrs("Nolayanan")), "", M_objrs("Nolayanan"))
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
End If

'sembunyiin principle kecuali SPV
If UCase(MDIForm1.txtlevel.Text) <> "SUPERVISOR" Then
    LblPrompA.Visible = False
    Label11(8).Visible = False
Else
    LblPrompA.Visible = True
    Label11(8).Visible = True
End If

'@@ 23-11-10 ini tambahan buat sembunyikan/tampilkan tombol ost jika ada data
'Dim M_OBJRS_ost As New ADODB.Recordset
'Set M_OBJRS_ost = New ADODB.Recordset
'M_OBJRS_ost.CursorLocation = adUseClient
'M_OBJRS_ost.Open "SELECT * FROM opening_screen where name like '%" + Trim(FrmCC_Colection.lblNama.Caption) + "%'", M_OBJCONN, adOpenDynamic, adLockOptimistic
'If M_OBJRS_ost.RecordCount <> 0 Then
'    SSCommand1(7).Visible = True
'Else
'    SSCommand1(7).Visible = True
'End If
'Set M_OBJRS_ost = Nothing

'@@ 22-03-2011 Tambahan, jika yang login agent maka tidak dapat ptp
If UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
    C_PTP.Enabled = False
    frmPTP.Enabled = False
    Frame5.Enabled = False
    Frame18.Enabled = False
    '@@ 15-06-2011 Tambahan, jika yang login agent tidka dapat menulis remarks
    'Frame19.Enabled = False
End If

End Sub

Sub isi_lastcall()
cbolastcall.CLEAR
Dim M_objrs As ADODB.Recordset
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient

If MDIForm1.txtlevel.Text = "AGENT" Then
    M_objrs.Open "Select * from ContactedDesc where kdnoprodpresented <> 'SP-SETTLE PAYMENT' ", M_OBJCONN, adOpenDynamic, adLockOptimistic
    Else
    M_objrs.Open "Select * from ContactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
End If
While Not M_objrs.EOF
    cbolastcall.AddItem M_objrs("KdNoProdPresented")
    M_objrs.MoveNext
Wend
Set M_objrs = Nothing

Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open "Select * from UnContactedDesc", M_OBJCONN, adOpenDynamic, adLockOptimistic
While Not M_objrs.EOF
    cbolastcall.AddItem M_objrs("KdNoProdPresented")
    M_objrs.MoveNext
Wend
Set M_objrs = Nothing
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

Private Sub Image1_Click(Index As Integer)
    Select Case Index
       Case 0
'          If Image1(0).Tag = 0 Then
'            Tdbbalance.Visible = True
'            tdbprincipal.Visible = True
'            Label11(14).Visible = True
'            Label11(15).Visible = True
'            Image1(0).Tag = 1
'            LblPrompA.Visible = True
'            Label11(8).Visible = True
'        Else
'            Tdbbalance.Visible = False
'            tdbprincipal.Visible = False
'            Label11(14).Visible = False
'            Label11(15).Visible = False
'            Label11(8).Visible = False
'            Image1(0).Tag = 0
'            LblPrompA.Visible = False
'        End If

    End Select
End Sub

Private Sub Label1_Click(Index As Integer)
  Dim ami As Integer
  
  Select Case Index
        Case 80
  'If Label1(80).Tag = 0 Then
     If UCase(MDIForm1.txtlevel.Text) = "SUPERVISOR" Or UCase(MDIForm1.txtlevel.Text) = "ADMIN" Or UCase(MDIForm1.txtlevel.Text) = "ADMINISTRATOR" Then
            TdbBalance.Visible = True
            tdbprincipal.Visible = True
            Label11(14).Visible = True
            Label11(15).Visible = True
            Label1(80).tag = 1
            LblPrompA.Visible = True
            Label11(8).Visible = True
            For ami = 1 To LstDoubleId.ListItems.Count
                LstDoubleId.ListItems(ami).SubItems(4) = ENCRIPY(True, LstDoubleId.ListItems(ami).SubItems(4))
            Next ami
        Else
            TdbBalance.Visible = False
            tdbprincipal.Visible = False
            Label11(14).Visible = False
            Label11(15).Visible = False
            Label11(8).Visible = False
            Label1(80).tag = 0
            LblPrompA.Visible = False
             For ami = 1 To LstDoubleId.ListItems.Count
                LstDoubleId.ListItems(ami).SubItems(4) = ENCRIPY(False, LstDoubleId.ListItems(ami).SubItems(4))
            Next ami
        End If
End Select

End Sub


Private Sub ListView1_Click(Index As Integer)
Dim KET As String
Select Case Index
Case 0

Case 1
If ListView1(1).ListItems.Count = 0 Then
Exit Sub
Else
   KET = TXtDetails.Text
      If Len(TXtDetails) = 0 Then
         TXtDetails.Text = " - " + ListView1(1).SelectedItem.SubItems(1)
      Else
         TXtDetails.Text = KET + " - " + ListView1(1).SelectedItem.SubItems(1)
      End If
End If
End Select
End Sub



Private Sub LstDoubleId_DblClick()
    If LstDoubleId.ListItems.Count = 0 Then
        Exit Sub
    End If
    frmCC_Colection2.Hide
    FrmCC_Colection.Show vbModal
End Sub

Private Sub LstPayment_DblClick()
If LstPayment.ListItems.Count = 0 Then
Exit Sub
Else
Call SSCommand2_Click(1)
End If
End Sub
Private Sub Lstscript_DblClick()
  If Lstscript.ListItems.Count > 0 Then
  StartMeUp (Lstscript.SelectedItem.SubItems(2))
  'MsgBox (LstScript.SelectedItem.SubItems(2))
   End If
End Sub
'@@ 11-03-2011 Di remarks, udah tidak diapakai
'Private Sub LstSMS_DblClick()
'If LstSMS.ListItems.Count > 0 Then
'
'no_telp = LstSMS.SelectedItem.Text
'isi_Pesan = LstSMS.SelectedItem.SubItems(3)
'
'MsgBox "No Telepon : " & no_telp & vbCrLf & "Isi Pesan : " & Trim(isi_Pesan)
'
'    Else
'    Exit Sub
' End If
'End Sub

'@@ 11-03-2011 Di remarks, udah tidak diapakai

'Private Sub LstSMS2_DblClick()
'If LstSMS2.ListItems.Count > 0 Then
'
'no_telp = LstSMS2.SelectedItem.Text
'isi_Pesan = LstSMS2.SelectedItem.SubItems(2)
'
'MsgBox "No Telepon : " & no_telp & vbCrLf & "Isi Pesan : " & Trim(isi_Pesan)
'
'    Else
'    Exit Sub
' End If
'End Sub

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

Private Sub Option9_Click()
If Option9.Value = True Then
'@@ 11-03-2011 Di remarks, udah tidak diapakai
'LstSMS.Visible = True
'LstSMS2.Visible = False
End If
End Sub

Private Sub Option10_Click()
If Option10.Value = True Then
'@@ 11-03-2011 Di remarks, udah tidak diapakai
'LstSMS.Visible = False
'LstSMS2.Visible = True
End If

End Sub

Private Sub SSCommand1_Click(Index As Integer)
Dim rsshut As New ADODB.Recordset
'On Error GoTo ke

Dim n As Integer
Select Case Index
  Case 7
  frmdetailskip.Show 1
  Case 5
    'FRMSCRIPT.Show 1
     Call OfferingDiscGuide
  Case 0
'  If Len(CmbPhone.Text) < 2 Then
'    MsgBox "Pilihan No Telephone harus diisi"
'    CmbPhone.SetFocus
'    Exit Sub
'  End If
        
        '@@220610 --- Agar agent tidak dapat mengisi no.telepon di combo phone
'        If IsNumeric(CmbPhone.Text) = True Then
'            If CmbPhone.Text <> "108" Then
'                CmbPhone.Text = ""
'                MsgBox "Pilih no telepon!", vbOKOnly + vbCritical, "Peringatan"
'                Exit Sub
'            End If
'        End If
        
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
                If txtECno.ReadOnly = False And UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
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
                If txtHomeAdd1A.ReadOnly = False And UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
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
                If txtHomeAdd2A.ReadOnly = False And UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
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
                If txtOfficeAdd1A.ReadOnly = False And UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
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
                If txtOfficeAdd2A.ReadOnly = False And UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
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
                If txtMobileAdd1A.ReadOnly = False And UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
                    MsgBox "Simpan Data terlebih dahulu"
                    Exit Sub
                End If
                txtPhone.Text = Trim(txtMobileAdd1.Value)
                telpno = txtPhone.Text
            Case "AddMobile2"
                If txtMobileAdd2A.ReadOnly = False And UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
                    MsgBox "Simpan Data terlebih dahulu"
                    Exit Sub
                End If
                txtPhone.Text = Trim(txtMobileAdd2.Value)
                telpno = txtPhone.Text
            Case Else
                Dim M_Objrs_Cek As ADODB.Recordset
                
                '@@09092011 Cek dulu apakah user telepon ada di tbllayanan telkom
                 txtPhone.Text = Replace(CmbPhone.Text, " ", "")
                cmdsql = "select * from tbllayanantelkom where nolayanan='"
                cmdsql = cmdsql + Trim(txtPhone.Text) + "'"
                Set M_Objrs_Cek = New ADODB.Recordset
                M_Objrs_Cek.CursorLocation = adUseClient
                M_Objrs_Cek.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek.RecordCount = 0 Then
                    MsgBox "Maaf, anda tidak dapat menelepon nomor yang tidak terdapat dalam database!", vbOKOnly + vbCritical, "Peringatan"
                    Set M_Objrs_Cek = Nothing
                    Exit Sub
                End If
        End Select
    
    '@@31-05-2012 Jika Status Account=PO dan CO maka tidak dapat di call
    If StatusAccount = "PO-" Or StatusAccount = "CO-" Then
        MsgBox "Mohon maaf! Status Account PAID OFF atau COMPLAIN tidak dapat di call!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    
    'Cek no telepon yang apakah masuk daftar blacklist. Jika masuk maka keluar sub!
    cmdsql = "select no_telp from tblblacklist where no_telp='"
    cmdsql = cmdsql + Replace(Trim(txtPhone.Text), " ", "") + "'"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
        M_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_objrs.RecordCount <> 0 Then
            MsgBox "No.Telepon yang anda hubungi masuk dalam daftar blacklist!. Silahkan hubungi TL  anda!.", vbOKOnly + vbExclamation, "Peringatan"
            Exit Sub
        End If
    Set M_objrs = Nothing
    
    '@@ 07-05-2012, Cek Apakah termasuk Unvalid number?
    cmdsql = "select no_telp from tblunvalid_number where no_telp='"
    cmdsql = cmdsql + Replace(Trim(txtPhone.Text), " ", "") + "' and custid='"
    cmdsql = cmdsql + CStr(lblcustid.Caption) + "'"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
        M_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_objrs.RecordCount <> 0 Then
            MsgBox "No.Telepon yang anda hubungi masuk dalam daftar Unvalid number!. Silahkan hubungi TL  anda!.", vbOKOnly + vbExclamation, "Peringatan"
            Exit Sub
        End If
    Set M_objrs = Nothing
    
    ' Matikan monitoring activity
'    i_monitoring_activity = 0
'    MDIForm1.Timer2.Enabled = False
    ' #####
    
    'jejaktian14042016==============================================
     If Obelisk = False Then
        'UNTUK ORANGE CLIENT
        MDIForm1.ActionCTI ("DIAL|020892" & GetNumber(CStr(Replace(txtPhone.Text, " ", ""))) & "|" & Trim(FrmCC_Colection.lblcustid.Caption) & "|" & Trim(FrmCC_Colection.lblcustid.Caption))
    Else
        'UNTUK OBELISK
        'MDIForm1.ActionCTI ("DIAL|" & GetNumber(CStr(Replace(txtPhone.Text, " ", ""))) & "|" & Trim(FrmCC_Colection.lblCustId.Caption) & "|" & Trim(FrmCC_Colection.lblCustId.Caption)) & "-" & mdiform1.txtusername.text
        MDIForm1.ActionCTI ("DIAL|" & GetNumber(CStr(Replace(txtPhone.Text, " ", ""))) & "|" & Trim(FrmCC_Colection.lblcustid.Caption))
    End If
    '============================================
    'Asli
    'MDIForm1.ActionCTI ("DIAL|020892" & GetNumber(CStr(Replace(txtPhone.Text, " ", ""))) & "|" & Trim(FrmCC_Colection.lblcustid.Caption) & "|" & Trim(FrmCC_Colection.lblcustid.Caption))
    '===========================================
    cmdsql = "Insert Into tblphonemonitorhst(UserId, CustId, NamaCh,StartDate, TelpNo, Recsource) Values ('" + MDIForm1.txtusername.Text + "' , '" + FrmCC_Colection.lblcustid.Caption + "','" + FrmCC_Colection.lblnama.Caption + "', '" + Format(CStr(MDIForm1.TDBDate1.Value), "yyyy-mm-dd") & " " & Format(Now, "hh:nn") + "' , '" + Replace(txtPhone.Text, " ", "") + "' ,'" + FrmCC_Colection.lblRecsource.Caption + "')"
    M_OBJCONN.Execute cmdsql
    
    Call OfferingDiscGuide
    
    'MDIForm1.CmbNo.Text = ""
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
     If bRenderrecord = True Then
          '  VIEW_MGMDATA.renderdonk
      End If
      bRenderrecord = False
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
     strsql = "select * from tblshut where nshut=1"
     Set rsshut = New ADODB.Recordset
     rsshut.CursorLocation = adUseClient
     rsshut.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
      If Not rsshut.EOF Then
         strsql = "UPDATE  tblshut SET nshut=0"
        M_OBJCONN.Execute (strsql)
        End
        Exit Sub
      End If
      Set rsshut = Nothing
      
'      '@@ Awal 061110 cek lock account sesuai settingan timer
'        Dim m_objrsTemp As ADODB.Recordset
'        Dim m_objrsWaktuServer As ADODB.Recordset
'        Dim m_objrsCurrent As ADODB.Recordset
'
'
'        Dim cmdsqlserver As String
'        Dim WaktuServer As Date
'        Dim WaktuAkhirCurrent As Date
'
'        'ambil waktu server
'        cmdsqlserver = "select now() as WaktuServer "
'        Set m_objrsWaktuServer = New ADODB.Recordset
'        m_objrsWaktuServer.CursorLocation = adUseClient
'        m_objrsWaktuServer.Open cmdsqlserver, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'        WaktuServer = Format(m_objrsWaktuServer(0), "mm-dd-yyyy hh:mm")
'        Set m_objrsWaktuServer = Nothing
'
'        'Cek lock account yang sedang berjalan
'        cmdsqlserver = "select * from tbltemplockacc_current "
'        Set m_objrsCurrent = New ADODB.Recordset
'        m_objrsCurrent.CursorLocation = adUseClient
'        m_objrsCurrent.Open cmdsqlserver, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'        If m_objrsCurrent.RecordCount <> 0 Then
'            WaktuAkhirCurrent = Format(m_objrsCurrent("end_lock"), "mm-dd-yyyy hh:mm")
'        Else
'            GoTo lockdata
'        End If
'
'        While Not m_objrsCurrent.EOF
'
'            WaktuAkhirCurrent = Format(m_objrsCurrent("end_lock"), "mm-dd-yyyy hh:mm")
'
'            If WaktuAkhirCurrent <= WaktuServer Then
'                'Cek dulu apakah ada user yang sedang mereset data
'                If Trim(m_objrsCurrent("f_locked")) = "2" Then
'                    GoTo KeluarLockAutoTL
'                End If
'
'                'update dulu status lock yang sedang berakhir, supaya agent lain ga ikut ngereset
'                cmdsqlserver = "update tbltemplockacc_current set f_locked='2' where id='"
'                cmdsqlserver = cmdsqlserver + Trim(m_objrsCurrent("id")) + "'"
'                M_OBJCONN.Execute cmdsqlserver
'
'                'Clear lock data yang sedang berjalan sesuai dengan agent yang di lock
'                cmdsqlserver = "update usertbl set dilockoleh='ClearByAutomatic',"
'                cmdsqlserver = cmdsqlserver + "lockdarispv=null,lock_entry_lpd=null,fromaccount=null,"
'                cmdsqlserver = cmdsqlserver + "lockmarkup=null,lockdarispvbuattl=null"
'                'Buat ambil kondisi agent yang sedang di lock
'                If Trim(m_objrsCurrent("account_lock")) = "ALL" Then
'                    cmdsqlserver = cmdsqlserver + " where usertype='1' "
'                ElseIf Left(Trim(m_objrsCurrent("account_lock")), 3) = "SPV" Then
'                    cmdsqlserver = cmdsqlserver + " where spvcode='"
'                    cmdsqlserver = cmdsqlserver + Trim(m_objrsCurrent("account_lock")) + "'"
'                Else
'                    cmdsqlserver = cmdsqlserver + " where userid='"
'                    cmdsqlserver = cmdsqlserver + Trim(m_objrsCurrent("account_lock")) + "'"
'                End If
'                M_OBJCONN.Execute cmdsqlserver
'
'                'Update status pesan ke nilai 1,untuk menampilkan pesan ke agent
'                cmdsqlserver = "update usertbl set f_pesanresetauto='1'"
'
'                'Buat mengupdate pesan kondisi agent yang di lock
'                If Trim(m_objrsCurrent("account_lock")) = "ALL" Then
'                    cmdsqlserver = cmdsqlserver + " where usertype='1'  "
'                ElseIf Left(Trim(m_objrsCurrent("account_lock")), 3) = "SPV" Then
'                    cmdsqlserver = cmdsqlserver + " where spvcode='"
'                    cmdsqlserver = cmdsqlserver + Trim(m_objrsCurrent("account_lock")) + "'"
'                Else
'                    cmdsqlserver = cmdsqlserver + " where userid='"
'                    cmdsqlserver = cmdsqlserver + Trim(m_objrsCurrent("account_lock")) + "'"
'                End If
'                M_OBJCONN.Execute cmdsqlserver
'
'                'Pindahkan data lock account current ke tabel data log tbltemplockacc_log
'                cmdsqlserver = "insert into tbltemplockacc_log select * from tbltemplockacc_current "
'                cmdsqlserver = cmdsqlserver + " where id='"
'                cmdsqlserver = cmdsqlserver + Trim(m_objrsCurrent("id")) + "'"
'                M_OBJCONN.Execute cmdsqlserver
'
'                'Hapus data di tabel locktemp current
'                cmdsqlserver = "delete from tbltemplockacc_current where id='"
'                cmdsqlserver = cmdsqlserver + Trim(m_objrsCurrent("id")) + "'"
'                M_OBJCONN.Execute cmdsqlserver
'
'             End If
'KeluarLockAutoTL:
'                m_objrsCurrent.MoveNext
'            Wend
'            Set m_objrsCurrent = Nothing
'
'
'
'
'        '=======
'lockdata:
'        'Setelah cek waktu lock yang habis, sekarang cek lock yg masih dalam antrian
'        cmdsqlserver = "select * from tbltemplockacc where f_locked isnull order by start_lock asc "
'        Set m_objrsTemp = New ADODB.Recordset
'        m_objrsTemp.CursorLocation = adUseClient
'        m_objrsTemp.Open cmdsqlserver, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'            'Cek ada ga data lock dalam antrian
'            If m_objrsTemp.RecordCount <> 0 Then
'                Dim WaktuAwal As Date
'                Dim WaktuAkhir As Date
'
'                While Not m_objrsTemp.EOF
'
'                    WaktuAwal = Format(m_objrsTemp("start_lock"), "mm-dd-yyyy hh:mm")
'                    WaktuAkhir = Format(m_objrsTemp("end_lock"), "mm-dd-yyyy hh:mm")
'
'                    If (WaktuAwal <= WaktuServer) And (WaktuAkhir > WaktuServer) Then
'                        'Cek apakah datanya sedang di lock sama agent lain?
'                        If Trim(m_objrsTemp("f_locked")) = "1" Then
'                            GoTo KeluarLockAutoTLLock
'                        End If
'
'                        'update status  f_lockednya jadi 1, supaya ga di log sama agent lain
'                        cmdsqlserver = "update tbltemplockacc set f_locked='1' where id='"
'                        cmdsqlserver = cmdsqlserver + Trim(m_objrsTemp("id")) + "'"
'                        M_OBJCONN.Execute cmdsqlserver
'
'                        'LAKUKAN LOCK DATA
'                        Dim i As Integer
'                        a = Split(m_objrsTemp("script_lock"), "|")
'
'                        For i = LBound(a) + 1 To UBound(a) - 1
'                            cmdsqlserver = Replace(a(i), "$", "'")
'                            M_OBJCONN.Execute cmdsqlserver
'                        Next i
'
'                        'Pindahin dulu data di tabel current ke tabel log, terus data di tabel current dihapus
''                        cmdsqlserver = "insert into tbltemplockacc_current "
''                        cmdsqlserver = cmdsqlserver + " select * from tbltemplockacc_log"
''                        M_OBJCONN.Execute cmdsqlserver --- Remarks dulu 10-11-10
'
''                        cmdsqlserver = "delete from tbltemplockacc_current"
''                        M_OBJCONN.Execute cmdsqlserver --- Remarks dulu 10-11-10
'
'                        'Pindahin data dari tabel temp lock ke tabel current log
'                        cmdsqlserver = "insert into tbltemplockacc_current "
'                        cmdsqlserver = cmdsqlserver + "select * from tbltemplockacc where id='"
'                        cmdsqlserver = cmdsqlserver + Trim(m_objrsTemp("id")) + "'"
'                        M_OBJCONN.Execute cmdsqlserver
'
'
'
'                       'Update status pesan ke nilai 1,untuk menampilkan pesan ke agent
'                        cmdsqlserver = "update usertbl set f_pesanlockauto='1',f_idsessstart='"
'                        cmdsqlserver = cmdsqlserver + Trim(m_objrsTemp("id")) + "',f_idsessend='"
'                        cmdsqlserver = cmdsqlserver + Trim(m_objrsTemp("id")) + "' "
'                        'Buat mengupdate pesan kondisi agent yang di lock
'                        If Trim(m_objrsTemp("account_lock")) = "ALL" Then
'                            cmdsqlserver = cmdsqlserver + " where usertype='1' "
'                        ElseIf Left(Trim(m_objrsTemp("account_lock")), 3) = "SPV" Then
'                            cmdsqlserver = cmdsqlserver + " where spvcode='"
'                            cmdsqlserver = cmdsqlserver + Trim(m_objrsTemp("account_lock")) + "'"
'                        Else
'                            cmdsqlserver = cmdsqlserver + " where userid='"
'                            cmdsqlserver = cmdsqlserver + Trim(m_objrsTemp("account_lock")) + "'"
'                        End If
'                        M_OBJCONN.Execute cmdsqlserver
'
'                        'Hapus data di templock
'                        cmdsqlserver = "delete from tbltemplockacc where id='"
'                        cmdsqlserver = cmdsqlserver + Trim(m_objrsTemp("id")) + "'"
'                        M_OBJCONN.Execute cmdsqlserver
'
'
'                    End If
'
'KeluarLockAutoTLLock:
'                    m_objrsTemp.MoveNext
'               Wend
'
'            End If
'
'        Set m_objrsTemp = Nothing
        Call PesanLockAuto
        
        
      '@@ Akhir 061110 cek lock account sesuai settingan timer
        ' Matikan monitoring activity
'        i_monitoring_activity = 0
'        MDIForm1.Timer2.Enabled = False
        ' #####

        Unload Me
        Exit Sub
'KeluarLockAuto:
        'Unload Me
    Case 1
        ' hidupkan monitoring activity
'        i_monitoring_activity = 0
'        MDIForm1.Timer2.Enabled = True
        ' #####
        MDIForm1.ActionCTI ("HANGUP")
    Case 4
        StatusCPA = "CPA Form 2"
        frmcpanew.Show 1
   

        
End Select
Exit Sub
'ke:
strsql = "update usertbl set stsaplikasi=0  where userid ='" + MDIForm1.txtusername.Text + "'"
M_OBJCONN.Execute (strsql)
MsgBox err.Description
 Exit Sub
 
End Sub

Public Sub Show_NEGOPTP()
Dim showlist As New ADODB.Recordset
Dim listitem As listitem
Dim cmdsql As String
Dim TOTPTP As Currency
Dim ssql As String
ssql = "SELECT CUSTID,sum(PAYMENT) as Jum FROM tbllunas WHERE custid = '" + lblcustid.Caption + "' GROUP BY CUSTID"
showlist.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic
If showlist.BOF And showlist.EOF Then
    TOTPTP = 0
Else
    TOTPTP = IIf(IsNull(showlist!jum), 0, showlist!jum)
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
cmdsql = "SELECT * FROM tblnegoptp where custid = '" + lblcustid.Caption + "' order by promisedate"

Set showlist = New ADODB.Recordset
showlist.CursorLocation = adUseClient
showlist.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic

LstPayment.ListItems.CLEAR
Dim n As Currency
While Not showlist.EOF
    Set listitem = LstPayment.ListItems.ADD(, , "")
        listitem.SubItems(1) = CStr(IIf(IsNull(showlist!ID), "", (showlist!ID)))
        listitem.SubItems(2) = CStr(IIf(IsNull(showlist!PromiseDate), "", Format(showlist!PromiseDate, "dd/mm/yyyy")))
        listitem.SubItems(3) = CStr(IIf(IsNull(showlist!PromisePay), "", (showlist!PromisePay)))
        n = n + Val(listitem.SubItems(3))
        If n <= TOTPTP Then
            listitem.ListSubItems(1).ForeColor = vbRed
            listitem.ListSubItems(2).ForeColor = vbRed
            listitem.ListSubItems(3).ForeColor = vbRed
        End If
        
        listitem.SubItems(4) = IIf(IsNull(showlist!Type), "", showlist!Type)
        listitem.SubItems(5) = CStr(IIf(IsNull(showlist!inputdate), "", Format(showlist!inputdate, "dd/mm/yyyy")))
     showlist.MoveNext
Wend

Set showlist = Nothing
End Sub
Public Sub show_cust()
Dim listitem As listitem
Dim m_data As New CLS_FRMCUST_CC
Dim m_cust1 As ADODB.Recordset
Dim m_cust2 As ADODB.Recordset
Dim cmdsql As String
Dim CMDSQL2 As String
Dim sPending As String
Dim CEKREC As New ADODB.Recordset
'On Error GoTo HELL:
'CMDSQL = "SELECT mgm.*, mgm_DETAIL.* FROM mgm INNER JOIN "
'CMDSQL = CMDSQL + "mgm_DETAIL ON mgm.CUSTID = dbo.mgm_DETAIL.CUSTID"

cmdsql = "select * from mgm"
'CMDSQL2 = "select * from mgm_detail"

Set m_cust = New ADODB.Recordset
'Set m_cust2 = New ADODB.Recordset
m_cust.CursorLocation = adUseClient
'm_cust2.CursorLocation = adUseClient
If shedulePTP_Show = True Then
    cmdsql = cmdsql + " where custid ='" & MDIForm1.LstGrade.SelectedItem.SubItems(1) & "'"
    m_cust.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
   
Else
    cmdsql = cmdsql + " where custid ='" & FrmCC_Colection.LstDoubleId.SelectedItem.Text & "'"
    m_cust.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    'CMDSQL2 = CMDSQL2 + " where custid ='" & VIEW_mgmDATA.LstVwSearchmgm.SelectedItem.SubItems(1) & "'"
    'm_cust2.Open CMDSQL2, M_OBJCONN, adOpenDynamic, adLockOptimistic
    'm_cust.Open "Select * from mgm where custid='" & VIEW_mgmDATA.LstVwSearchmgm.SelectedItem.SubItems(1) & "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
End If

'tampilkan data tabel mgm
If Not m_cust.EOF Then
    
    '@@31052012 Buat Menyimpan Status Account
    StatusAccount = ""
    StatusAccount = IIf(IsNull(m_cust("f_cek_new")), "", m_cust("f_cek_new"))
    
    '@@220610 - Memberikan tanda merah pada no telepon yang di blacklist
    If m_cust("f_homeno") = 1 Then
        txtHomeNo1.ForeColor = vbRed
        txtHomeNo1A.ForeColor = vbRed
    End If
    If m_cust("f_homeno2") = 1 Then
        txtHomeNo2.ForeColor = vbRed
        txtHomeNo2A.ForeColor = vbRed
    End If
    
    If m_cust("f_officeno") = 1 Then
        txtOfficeNo1.ForeColor = vbRed
        txtOfficeNo1A.ForeColor = vbRed
    End If
    If m_cust("f_officeno2") = 1 Then
        txtOfficeNo2.ForeColor = vbRed
        txtOfficeNo2A.ForeColor = vbRed
    End If
    
    If m_cust("f_mobileno") = 1 Then
        txtMobileNo1.ForeColor = vbRed
        txtMobileNo1A.ForeColor = vbRed
    End If
    If m_cust("f_mobileno2") = 1 Then
        txtMobileNo2.ForeColor = vbRed
        txtMobileNo2A.ForeColor = vbRed
    End If
    
    If m_cust("f_homenoadd1") = 1 Then
        txtHomeAdd1.ForeColor = vbRed
        txtHomeAdd1A.ForeColor = vbRed
    End If
    If m_cust("f_homenoadd2") = 1 Then
        txtHomeAdd2.ForeColor = vbRed
        txtHomeAdd2A.ForeColor = vbRed
    End If

    If m_cust("f_officenoadd1") = 1 Then
         txtOfficeAdd1.ForeColor = vbRed
         txtOfficeAdd1A.ForeColor = vbRed
    End If
    If m_cust("f_officenoadd2") = 1 Then
        txtOfficeAdd2.ForeColor = vbRed
        txtOfficeAdd2A.ForeColor = vbRed
    End If
    
    If m_cust("f_mobilenoadd1") = 1 Then
         txtMobileAdd1.ForeColor = vbRed
         txtMobileAdd1A.ForeColor = vbRed
    End If
    If m_cust("f_mobilenoadd2") = 1 Then
        txtMobileAdd2.ForeColor = vbRed
        txtMobileAdd2A.ForeColor = vbRed
    End If
    
    If m_cust("f_ec_telp") = 1 Then
         txtECno.ForeColor = vbRed
         txtECnoA.ForeColor = vbRed
    End If
    
  '@@ 08-06-2011 SEMUA TELEPON DIBUKA, STATUS APAPUN
' '@@ 11-04-2011 , Sementara untuk custid yang diberikan
'    If m_cust("status_additional") = "1" Then
'        Frame15(5).Visible = True
'        Frame17.Visible = True
'
'        Frame15(2).Visible = True
'        Frame20.Visible = True
'    End If
'
'    '@@ 02-05-2011, untuk memunculkan additional info dan EC disesuaikan dengan status
'    'Status ON-, VL-, PR- munculkan additional info
'    '@@ 26 May 2011 BP- dan PTP- ditampilkan
'    Dim CekStatus As String
'    CekStatus = IIf(IsNull(m_cust("f_cek_new")), "", m_cust("f_cek_new"))
'    If Trim(CekStatus) = "ON-" Or Trim(CekStatus) = "VL-" Or Trim(CekStatus) = "PR-" Or Trim(CekStatus) = "BP-" Or Mid(Trim(CekStatus), 1, 3) = "PTP" Then
'        Frame15(5).Visible = True
'        Frame17.Visible = True
'
'        Frame15(2).Visible = True
'        Frame20.Visible = True
'    End If
'
'    'Jika status OS maka yang ditampilkan EC saja
'    If Trim(CekStatus) = "OS-" Then
'        Frame15(5).Visible = False
'        Frame17.Visible = False
'
'        Frame15(2).Visible = True
'        Frame20.Visible = True
'    End If
'
'    'Jika status account masih kosong, maka tampilkan EC
'    '@@ 11-May-2011
'    If CekStatus = "" Then
'        Frame15(5).Visible = False
'        Frame17.Visible = False
'
'        Frame15(2).Visible = True
'        Frame20.Visible = True
'    End If
    
    
    '@@ 02022011
    LblMinPayment.Value = IIf(IsNull(m_cust("minpayment")), "0", m_cust("minpayment"))

    lblstatus.Caption = IIf(IsNull(m_cust("statusprior")), "", "Status : " & m_cust("statusprior"))
    lblcustid.Caption = IIf(IsNull(m_cust("CUSTID")), "", m_cust("CUSTID"))
    LblMother.Caption = IIf(IsNull(m_cust("mother")), "", m_cust("mother"))
    'sql = "delete  from tblnegoptp where custid in (select custid from tbllunas where custid ='" + IIf(IsNull(m_cust("CUSTID")), "", m_cust("CUSTID")) + "')"
    TxtCustid.Text = IIf(IsNull(m_cust("CUSTID")), "", m_cust("CUSTID"))
    txtname.Text = IIf(IsNull(m_cust("NAME")), "", m_cust("NAME"))
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
    lblnama.Caption = IIf(IsNull(m_cust("NAME")), "", m_cust("NAME"))
    lblCardNo.Caption = IIf(IsNull(m_cust("NoCard")), "", m_cust("NoCard"))
    lblID.Caption = IIf(IsNull(m_cust("ktpno")), "", m_cust("ktpno"))
    'lblDate.Value = IIf(IsNull(m_cust("BIRTHD")), "", Format(m_cust("BIRTHD"), "dd-mmm-yyyy"))
    LblDOB.Caption = IIf(IsNull(m_cust("DOB")), "", Left(m_cust("DOB"), 10))
    lblAddr.Text = IIf(IsNull(m_cust("ADDRNOW")), "", m_cust("ADDRNOW"))
    TDB_cur_bal = IIf(IsNull(m_cust("CURBAL")), "", m_cust("CURBAL"))
    TXTRUMUS.Text = IIf(IsNull(m_cust("typerumus")), "", m_cust("typerumus"))
    Combo1.Text = IIf(IsNull(m_cust("stscallcust")), "", m_cust("stscallcust"))
    'tdbmaxad.Value = Format(IIf(IsNull(m_cust("maxad")), "0", m_cust("maxad")), "##,###")
    'tdbminad.Value = Format(IIf(IsNull(m_cust("minad")), "0", m_cust("minad")), "##,###")
     
    '@@ Tambahan 2 field (map dan cycle)
    LblMap = IIf(IsNull(m_cust("map")), "0", m_cust("map"))
    LblCycle = IIf(IsNull(m_cust("cycle")), "0", m_cust("cycle"))

   Set CEKREC = New ADODB.Recordset
    CEKREC.CursorLocation = adUseClient
    CEKREC.Open "select * from opening_screen where custid='" + lblcustid.Caption + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
    If CEKREC.RecordCount > 0 Then
        SSCommand1(7).BackColor = vbRed
        TimerBlink.Enabled = True
    Else
        TimerBlink.Enabled = False
    End If
    
     If InStr(1, VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(3), "DE") > 0 Then
        txthasil.Visible = True
     Else
        txthasil.Visible = False
     End If
     
     Text6.Text = IIf(IsNull(m_cust("disapp")), "0", m_cust("disapp"))
     tdbhptrace.Value = IIf(IsNull(m_cust("hp1trace")), "", m_cust("hp1trace"))
     tdbtelptrace.Value = IIf(IsNull(m_cust("tlp1trace")), "", m_cust("tlp1trace"))
     txtremarkstrace.Text = IIf(IsNull(m_cust("addrtrace")), "", m_cust("addrtrace"))
     
     bcekptp = False
    vrcek = IIf(IsNull(m_cust("f_cek_new")), "", m_cust("f_cek_new"))
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
     
    '@@ 04-03-2011 Ubah status jika TL/SPV/Admin yang buka dapat membuka semua status
    If UCase(Trim(MDIForm1.txtlevel.Text)) = "ADMINISTRATOR" Or _
       UCase(Trim(MDIForm1.txtlevel.Text)) = "ADMIN" Or _
       UCase(Trim(MDIForm1.txtlevel.Text)) = "SUPERVISOR" Or _
       UCase(Trim(MDIForm1.txtlevel.Text)) = "TEAMLEADER" Then
       
        If vrcek <> "BP-" Or Mid(vrcek, 1, 3) = "PTP" Or Mid(vrcek, 1, 3) = "POP" Then
            strsql = "Select * from contacteddesc WHERE status=1"
        ElseIf vrcek = "BP-" Then
                 strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('BP-') AND status=1"
        ElseIf Mid(vrcek, 1, 3) = "PTP" Then
                 strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('BP-') AND status=1"
        ElseIf Mid(vrcek, 1, 3) = "POP" Then
                 strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('POP') AND status=1"
        End If
        
    Else
    '@@ 04-03-2011 Nah ini jika yang login Agent
        If vrcek = "" Then
            strsql = "Select * from contacteddesc WHERE status=1"
        Else
            If vrcek = "VL-" Then
                strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('VL-','PR-','ON-','PO-','CO-') and status=1"
            ElseIf vrcek = "OS-" Then
                 strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('VL-','PR-','ON-','SK-','PO-','CO-') AND status=1"
            ElseIf vrcek = "PR-" Then
                 strsql = "Select * from contacteddesc WHERE  substring(KdNoProdPresented,1,3) in('PR-','ON-','PO-','CO-') AND status=1"
            ElseIf vrcek = "ON-" Then
                 strsql = "Select * from contacteddesc WHERE  substring(KdNoProdPresented,1,3) in('ON-','PO-','CO-') AND status=1"
            ElseIf vrcek = "SK-" Then
                 strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('VL-','PR-','SK-','PO-','CO-') AND status=1"
            ElseIf vrcek = "SP-" Then
                 strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('SP-','PO-','CO-') AND status=1"
            ElseIf vrcek = "BP-" Then
                 strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('BP-') AND status=1"
            ElseIf Mid(vrcek, 1, 3) = "PTP" Then
                 strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('BP-') AND status=1"
            ElseIf Mid(vrcek, 1, 3) = "POP" Then
                 strsql = "Select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('POP') AND status=1"
            '@@31052012Tambahan JIKA PAID OFF (PO-) DAN COMPLAIN (CO-)
            ElseIf Mid(vrcek, 1, 3) = "PO-" Then
                strsql = "select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('PO-') AND status=1"
            ElseIf Mid(vrcek, 1, 3) = "CO-" Then
                strsql = "select * from contacteddesc WHERE substring(KdNoProdPresented,1,3) in('CO-') AND status=1"
            Else
                strsql = " Select * from contacteddesc WHERE status=1 "
            End If
            
        End If
    End If
    'STRSQL = " Select * from contacteddesc WHERE status=1 "
    cboaccount.CLEAR
    M_objrs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    While Not M_objrs.EOF
        cboaccount.AddItem M_objrs!KdNoProdPresented
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
    
   If Left(IIf(IsNull(m_cust("f_cek_new")), "", m_cust("f_cek_new")), 3) <> "PTP" Then
    'cboaccount.Text = IIf(IsNull(m_cust("f_cek_new")), "", m_cust("f_cek_new"))
    cboaccount.Text = IIf(IsNull(m_cust("kethslkerja_new")), "", m_cust("kethslkerja_new"))
   ElseIf Left(IIf(IsNull(m_cust("f_cek_new")), "", m_cust("f_cek_new")), 3) = "PTP" Then
     cboPTP.Text = IIf(IsNull(m_cust("kethslkerja_new")), "", m_cust("kethslkerja_new"))
     cboaccount = IIf(IsNull(m_cust("ptpdesc")), "", m_cust("ptpdesc"))
   End If
  
  
   
   If Left(IIf(IsNull(m_cust("f_cek_new")), "", m_cust("f_cek_new")), 3) = "PTP" Then
        C_PTP.Value = vbChecked
   End If
   
   
  TglPTPNew = IIf(IsNull(m_cust("tglptpnew")), "", m_cust("tglptpnew"))
  If TglPTPNew <> "" Then
        tdbptpnew.Value = Format(TglPTPNew, "dd/mm/yyyy")
  End If
  
If Left(vrcek, 3) = "PTP" Then
        SSCommand1(4).Visible = True
        Label43(2).Visible = True
Else
        SSCommand1(4).Visible = False
        Label43(2).Visible = False
End If

    If Left(vrcek, 2) = "BP" Then
  '  cboPOPSP.Enabled = False
'        FrmContacted.Enabled = False
'        C_Contacted.Enabled = False
'        cmbContacted.Enabled = False
'        cmbDescCon.Enabled = False
     End If
    
    lblOfficeAddr.Text = IIf(IsNull(m_cust("ADDRPT")), "", m_cust("ADDRPT"))
    lblZIP.Caption = IIf(IsNull(m_cust("ZIPNOW")), "", m_cust("ZIPNOW"))
    lblNoCard.Caption = IIf(IsNull(m_cust("NoCard")), "", m_cust("NoCard"))
    lblNoPay.Caption = IIf(IsNull(m_cust("NoPay")), "", m_cust("NoPay"))
      
       
        
        
        
    'Else
    
     LblPrompA.Value = IIf(IsNull(m_cust("Principal")), "", m_cust("Principal"))
     
        
   If UCase(MDIForm1.txtlevel.Text) <> "SUPERVISOR" Then
        If IIf(IsNull(m_cust("flaglead")), 0, m_cust("flaglead")) = 1 Then
            LblPrompA.Visible = False
            Label11(8).Visible = False
        Else
            LblPrompA.Visible = True
            Label11(8).Visible = True
       End If
    Else
          LblPrompA.Visible = False
          Label11(8).Visible = False
    End If
    
   ' End If
    
    
    tdbprincipal.Value = IIf(IsNull(m_cust("Principal")), "", m_cust("Principal"))
    lblOpenDate.Value = IIf(IsNull(m_cust("OpenDate")), "", m_cust("OpenDate"))
    lblLastBill.Value = IIf(IsNull(m_cust("LastBill")), "", m_cust("LastBill"))
    lblLcAtm.Value = IIf(IsNull(m_cust("LcATMP")), "", m_cust("LcATMP"))
    txttenor.Value = IIf(IsNull(m_cust("tenor")), 0, m_cust("tenor"))
    vrtenor = IIf(IsNull(m_cust("tenor")), 0, m_cust("tenor"))
    lblBrokenPromised.Caption = IIf(IsNull(m_cust("BrokenPromise")), "", m_cust("BrokenPromise"))
    lblBD.Value = IIf(IsNull(m_cust("B_D")), "", m_cust("B_D"))
    lblLimit.Value = IIf(IsNull(m_cust("Limit")), "", m_cust("Limit"))
    vramount = IIf(IsNull(m_cust!amountptp), 0, m_cust!amountptp)
    vrcekamont = IIf(IsNull(m_cust!amountptp), 0, m_cust!amountptp)
    If ListView1(0).ListItems.Count = 0 Then
    lblPayDt.Value = IIf(IsNull(m_cust("Pay_Dt")), "", m_cust("Pay_Dt"))
    End If
    
    
    If ListView1(0).ListItems.Count = 0 Then
    lblLastPay.Value = IIf(IsNull(m_cust("LastPay")), "", m_cust("LastPay"))
    End If
    
    lblTtlPay.Value = IIf(IsNull(m_cust("TtlPay")), "", m_cust("TtlPay"))
    'If IIf(IsNull(m_cust("flaglead")), 0, m_cust("flaglead")) = 1 Then
     '   lblAmount.Value = IIf(IsNull(m_cust("AmountWo")), "", Format(m_cust("AmountWo"), "##.##0"))
     '   If lblAmount.ValueIsNull Then
      '      lblAmount.Value = 0
      '  Else
       '     lblAmount.Value = lblAmount.Value + (lblAmount.Value * 18.26) / 100
       ' End If
        
    'Else
    
    
    lblAmount.Value = IIf(IsNull(m_cust("AmountWo")), "", Format(m_cust("AmountWo"), "##.##0"))
    
    'End If
    
    If lblAmount.ValueIsNull Then
    
            tdbmaxad.Value = 0
        Else
            tdbmaxad.Value = lblAmount.Value - (lblAmount.Value * 24) / 100
        End If
    
    
     If lblAmount.ValueIsNull Then
            tdbminad.Value = tdbminad.Value - (lblAmount.Value * 35) / 100
        Else
            tdbminad.Value = lblAmount.Value - (lblAmount.Value * 31) / 100
        End If
        
    TdbBalance.Value = IIf(IsNull(m_cust("AmountWo")), "", Format(m_cust("AmountWo"), "##.##0"))
    
    txtHomeNo1.Value = IIf(IsNull(m_cust("HOMENO")), "", m_cust("HOMENO"))
    AHome1.Value = IIf(IsNull(m_cust("AHOMENO")), "", m_cust("AHOMENO"))
    
    
    
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
   
    txtHomeAdd1.Value = IIf(IsNull(m_cust("HOMENOADD1")), "", m_cust("HOMENOADD1"))
    If IsNull(m_cust("HOMENOADD1")) = False And m_cust("HOMENOADD1") <> "" Then
        txtHomeAdd1A.Value = Left(m_cust("HOMENOADD1"), 4) & "BBB" & Mid(m_cust("HOMENOADD1"), 8, 15)
'        '@@ 11-04-2011 Di hide dulu, kecuali data tertentu
'        If m_cust("status_additional") = "1" Then
'            CmbPhone.AddItem "AddHome1"
'        End If
'        '@@02-05-2011, Jika status account ON-, VL-, dan PR- maka Additional&EC di tampilkan
'        '@@26 May 2011 BP- dan PTP- ditampilkan juga
'        If Trim(CekStatus) = "ON-" Or Trim(CekStatus) = "PR-" Or Trim(CekStatus) = "VL-" Or Trim(CekStatus) = "BP-" Or Mid(Trim(CekStatus), 1, 3) = "PTP" Then
'            CmbPhone.AddItem "AddHome1"
'        End If
        '@@ 08-06-2011, TELEPON DIBUKA SEMUA, STATUS APAPUN
        CmbPhone.AddItem "AddHome1"
    Else
        txtHomeAdd1.Visible = True
        txtHomeAdd1A.Visible = False
    End If
    txtHomeAdd2.Value = IIf(IsNull(m_cust("HOMENOADD2")), "", m_cust("HOMENOADD2"))
    If IsNull(m_cust("HOMENOADD2")) = False And m_cust("HOMENOADD2") <> "" Then
        txtHomeAdd2A.Value = Left(m_cust("HOMENOADD2"), 4) & "BBB" & Mid(m_cust("HOMENOADD2"), 8, 15)
'        '@@ 11-04-2011 Di hide dulu, kecuali data tertentu
'        If m_cust("status_additional") = "1" Then
'            CmbPhone.AddItem "AddHome2"
'        End If
'        '@@02-05-2011, Jika status account ON-, VL-, dan PR- maka Additional&EC di tampilkan
'        '@@26 May 2011, BP- dan PTP- juga ditampilkan
'        If Trim(CekStatus) = "ON-" Or Trim(CekStatus) = "PR-" Or Trim(CekStatus) = "VL-" Or Trim(CekStatus) = "BP-" Or Mid(Trim(CekStatus), 1, 3) = "PTP" Then
'            CmbPhone.AddItem "AddHome2"
'        End If
        '@@ 08-06-2011, Telepon dibuka semua, status apapun
        CmbPhone.AddItem "AddHome2"
    Else
        txtHomeAdd2A.Visible = False
        txtHomeAdd2.Visible = True
    End If
    txtOfficeAdd1.Value = IIf(IsNull(m_cust("OFFICENOADD1")), "", m_cust("OFFICENOADD1"))
    If IsNull(m_cust("OFFICENOADD1")) = False And m_cust("OFFICENOADD1") <> "" Then
        txtOfficeAdd1A.Value = Left(m_cust("OFFICENOADD1"), 4) & "BBB" & Mid(m_cust("OFFICENOADD1"), 8, 15)
'        '@@ 11-04-2011 Di hide dulu, kecuali data tertentu
'        If m_cust("status_additional") = "1" Then
'            CmbPhone.AddItem "AddOffice1"
'        End If
'        '@@02-05-2011, Jika status account ON-, VL-, dan PR- maka Additional&EC di tampilkan
'        '@@26 May 2011, BP- dan PTP- ditampilkan juga
'        If Trim(CekStatus) = "ON-" Or Trim(CekStatus) = "PR-" Or Trim(CekStatus) = "VL-" Or Trim(CekStatus) = "BP-" Or Mid(Trim(CekStatus), 1, 3) = "PTP" Then
'            CmbPhone.AddItem "AddOffice1"
'        End If
        '@@08-06-2011, Semua telepon dibuka, status apapun
        CmbPhone.AddItem "AddOffice1"
    Else
        txtOfficeAdd1A.Visible = False
        txtOfficeAdd1.Visible = True
    End If
    txtOfficeAdd2.Value = IIf(IsNull(m_cust("OFFICENOADD2")), "", m_cust("OFFICENOADD2"))
    If IsNull(m_cust("OFFICENOADD2")) = False And m_cust("OFFICENOADD2") <> "" Then
        
        anto = Trim(Left(m_cust("OFFICENOADD2"), 4) + " " + Mid(m_cust("OFFICENOADD2"), 8, 15))
        If Len(anto) = 0 Then
        txtOfficeAdd2A.Value = ""
        
        Else
        
        txtOfficeAdd2A.Value = Left(m_cust("OFFICENOADD2"), 4) & "BBB" & Mid(m_cust("OFFICENOADD2"), 8, 15)
        
        End If
        
'        '@@ 11-04-2011 Di hide dulu, kecuali data tertentu
'        If m_cust("status_additional") = "1" Then
'            CmbPhone.AddItem "AddOffice2"
'        End If
'        '@@02-05-2011, Jika status account ON-, VL-, dan PR- maka Additional&EC di tampilkan
'        '@@26 May 2011 BP- dan PTP- juga ditampilkan
'        If Trim(CekStatus) = "ON-" Or Trim(CekStatus) = "PR-" Or Trim(CekStatus) = "VL-" Or Trim(CekStatus) = "BP-" Or Mid(Trim(CekStatus), 1, 3) = "PTP" Then
'            CmbPhone.AddItem "AddOffice2"
'        End If
        '@@08-06-2011,SEMUA TELEPON DIBUKA,STATUS APAPUN
        CmbPhone.AddItem "AddOffice2"
    Else
        txtOfficeAdd2.Visible = True
        txtOfficeAdd2A.Visible = False
    End If
    txtMobileAdd1.Value = IIf(IsNull(m_cust("MOBILENOADD1")), "", m_cust("MOBILENOADD1"))
    If IsNull(m_cust("MOBILENOADD1")) = False And m_cust("MOBILENOADD1") <> "" Then
        txtMobileAdd1A.Value = Left(m_cust("MOBILENOADD1"), 4) & "BBB" & Mid(m_cust("MOBILENOADD1"), 8, 15)
'        '@@ 11-04-2011 Di hide dulu, kecuali data tertentu
'        If m_cust("status_additional") = "1" Then
'            CmbPhone.AddItem "AddMobile1"
'        End If
'        '@@02-05-2011, Jika status account ON-, VL-, dan PR- maka Additional&EC di tampilkan
'        '@@26 May 2011, BP- dan PTP- juga ditampilkan
'        If Trim(CekStatus) = "ON-" Or Trim(CekStatus) = "PR-" Or Trim(CekStatus) = "VL-" Or Trim(CekStatus) = "BP-" Or Mid(Trim(CekStatus), 1, 3) = "PTP" Then
'            CmbPhone.AddItem "AddMobile1"
'        End If
        '@@ 08-06-2011 Semua telepon dibuka,status apapun
        CmbPhone.AddItem "AddMobile1"
    Else
        txtMobileAdd1.Visible = True
        txtMobileAdd1A.Visible = False
    End If
    txtMobileAdd2.Value = IIf(IsNull(m_cust("MOBILENOADD2")), "", m_cust("MOBILENOADD2"))
    If IsNull(m_cust("MOBILENOADD2")) = False And m_cust("MOBILENOADD2") <> "" Then
        txtMobileAdd2A.Value = Left(m_cust("MOBILENOADD2"), 4) & "BBB" & Mid(m_cust("MOBILENOADD2"), 8, 15)
'        '@@ 11-04-2011 Di hide dulu, kecuali data tertentu
'        If m_cust("status_additional") = "1" Then
'            CmbPhone.AddItem "AddMobile2"
'        End If
'        '@@02-05-2011, Jika status account ON-, VL-, dan PR- maka Additional&EC di tampilkan
'        '@@26 May 2011, BP- dan PTP- juga ditampilkan
'        If Trim(CekStatus) = "ON-" Or Trim(CekStatus) = "PR-" Or Trim(CekStatus) = "VL-" Or Trim(CekStatus) = "BP-" Or Mid(Trim(CekStatus), 1, 3) = "PTP" Then
'            CmbPhone.AddItem "AddMobile2"
'        End If
        '@@ 08-06-2011 Semua telepon dibuka, status apapun
        CmbPhone.AddItem "AddMobile2"
    Else
        txtMobileAdd2.Visible = True
        txtMobileAdd2A.Visible = False
    End If
   
    AddrNow.Text = IIf(IsNull(m_cust("TxtPtpAddr")), "", m_cust("TxtPtpAddr"))
    LblLunas.Caption = IIf(IsNull(m_cust!tgllunas), "", "TELAH LUNAS")
    TxtEC.Text = IIf(IsNull(m_cust!ec_name), "", m_cust!ec_name)
    txtECno.Value = IIf(IsNull(m_cust!ec_telp), "", m_cust!ec_telp)
    If IsNull(m_cust("ec_telp")) = False And m_cust("ec_telp") <> "" Then
        txtECnoA.Value = Left(m_cust("ec_telp"), 4) & "BBB" & Mid(m_cust("ec_telp"), 8, 15)
'        '@@ 11-04-2011 Di hide dulu, kecuali data tertentu
'        If m_cust("status_additional") = "1" Then
'            CmbPhone.AddItem "EconPhone"
'        End If
'          '@@02-05-2011, Jika status account ON-, VL-, dan PR- maka Additional&EC di tampilkan
'          '@@26 May 2011 , BP- dan PTP- juga ditampilkan
'        If Trim(CekStatus) = "ON-" Or Trim(CekStatus) = "PR-" Or Trim(CekStatus) = "VL-" Or Trim(CekStatus) = "OS-" Or CekStatus = "" Or Trim(CekStatus) = "BP-" Or Mid(Trim(CekStatus), 1, 3) = "PTP" Then
'            CmbPhone.AddItem "EconPhone"
'        End If
        '@@ 08-06-2011 Semua Telepon dibuka,status apapun
        CmbPhone.AddItem "EconPhone"
    Else
        txtECnoA.Visible = False
        txtECno.Visible = True
    End If
    txtECAdd.Text = IIf(IsNull(m_cust!ECAddr), "", m_cust!ECAddr)
    cbolastcall.Text = IIf(IsNull(m_cust!statuscall), "", m_cust!statuscall)
    cbolastcall.Text = IIf(IsNull(m_cust!stscallwith), "", m_cust!stscallwith)
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
    If UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
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
   
    
    sPending = CStr(Trim(IIf(IsNull(m_cust!f_Pending), "", m_cust!f_Pending)))
     If sPending = "Pending" Then
         chkAppv(0).Value = 0
    End If
    
'    Select Case m_cust!RECSTATUS
'        Case "V"
'            C_VALID.Value = 1
'            cbovalid.Text = IIf(IsNull(m_cust("KETHSLKERJA")), "", m_cust("KETHSLKERJA"))
'            cbodescvalid.Text = IIf(IsNull(m_cust("KETHSLKERJADESC")), "", m_cust("KETHSLKERJADESC"))
'        Case "N"
'            C_NotContacted.Value = 1
'            cmbUncontacted.Text = IIf(IsNull(m_cust("KETHSLKERJA")), "", m_cust("KETHSLKERJA"))
'            cmbDescUn.Text = IIf(IsNull(m_cust("KETHSLKERJADESC")), "", m_cust("KETHSLKERJADESC"))
'        Case "C"
'            C_Contacted.Value = 1
'            kontak = True
'            If mdiform1.txtlevel.text = "Agent" Then
'                If Left(vrcek, 3) = "POP" Then
'                    C_SKIP.Enabled = False
'                    C_VALID.Enabled = False
'                    cboPOPSP.Enabled = False
'                    FrmPayment.Enabled = True
'                    C_Payment.Value = 1
'                End If
'            End If
'            cmbContacted.Text = IIf(IsNull(m_cust("KETHSLKERJA")), "", m_cust("KETHSLKERJA"))
'      Case "P"
'            C_PTP.Value = 1
'            cboPTP.Text = IIf(IsNull(m_cust("KETHSLKERJA")), "", m_cust("KETHSLKERJA"))
'            'cmbDescCon.Text = IIf(IsNull(m_cust("KETHSLKERJADESC")), "", m_cust("KETHSLKERJADESC"))
'            If mdiform1.txtlevel.text = "Agent" Then
'                C_VALID.Enabled = False
'                C_Contacted.Enabled = False
'                FrMValid.Enabled = False
'                C_SKIP.Enabled = False
'                FrmSKIP.Enabled = False
'            End If
'         Case "S"
'            C_SKIP.Value = 1
'            cboskip.Text = IIf(IsNull(m_cust("KETHSLKERJA")), "", m_cust("KETHSLKERJA"))
'            cbodescskip.Text = IIf(IsNull(m_cust("KETHSLKERJADESC")), "", m_cust("KETHSLKERJADESC"))
'         Case "O"
'            'C_POPSP.Value = 1
'            cboPOPSP.Text = IIf(IsNull(m_cust("KETHSLKERJA")), "", m_cust("KETHSLKERJA"))
'            'cmbDescCon.Text = IIf(IsNull(m_cust("KETHSLKERJADESC")), "", m_cust("KETHSLKERJADESC"))      cmbDescCon.Text = IIf(IsNull(m_cust("KETHSLKERJADESC")), "", m_cust("KETHSLKERJADESC"))
'     End Select
     
    If MDIForm1.txtlevel.Text = "Agent" Then
'        If IIf(IsNull(m_cust!RECSTATUS), "", m_cust!RECSTATUS) <> "O" Then
'            frmpopsp.Enabled = False
'           cboPOPSP.Enabled = False
'        End If
    End If
        If IIf(IsNull(m_cust!f_cek_new), "", Left(m_cust!f_cek_new, 3)) = "PTP" Or Left(m_cust!f_cek_new, 3) = "POP" Or Left(m_cust!f_cek_new, 3) = "SP-" Or Left(m_cust!f_cek_new, 3) = "PRE" Then
            C_Payment.Value = 1
            TdbPTP.Value = IIf(IsNull(m_cust!TdbDatePTP), "", m_cust!TdbDatePTP)
            vrtdbdateptp = IIf(IsNull(m_cust!TdbDatePTP), "", m_cust!TdbDatePTP)
            vrdateptp = IIf(IsNull(m_cust!TdbDatePTP), "", m_cust!TdbDatePTP)
            TDBDate3.Value = IIf(IsNull(m_cust!dateptp), "", m_cust!dateptp)
            vrnewdate = IIf(IsNull(m_cust!dateptp), "", Format(m_cust!dateptp, "dd/mm/yyyy"))
            txtPayment.Value = IIf(IsNull(m_cust!ttlptp), 0, m_cust!ttlptp)
            vrttlptp = IIf(IsNull(m_cust!ttlptp), 0, m_cust!ttlptp)
            Tdabamoint.Value = IIf(IsNull(m_cust!amountptp), 0, m_cust!amountptp)
            vramount = IIf(IsNull(m_cust!amountptp), 0, m_cust!amountptp)
            txtpayment2.Value = IIf(IsNull(m_cust!ttlptp), 0, m_cust!ttlptp) 'tampilkan di detail payment
            cmbDiscount.Text = IIf(IsNull(m_cust!discpersen), 0, m_cust!discpersen)
            vrdiskon = IIf(IsNull(m_cust!discpersen), 0, m_cust!discpersen)
            CmbBaseOn.Text = IIf(IsNull(m_cust!CmbBaseOn), "", m_cust!CmbBaseOn)
            vrbaseon = IIf(IsNull(m_cust!CmbBaseOn), "", m_cust!CmbBaseOn)
            'TdbDatePTP.Value = IIf(IsNull(m_cust!TGLINCOMING), "", m_cust!TGLINCOMING)
        Else
        End If
End If
Call Custid_Double
'Set m_cust1 = M_DATA.QUERY_HIST_CUST(M_OBJCONN, "CUSTID = '" + lblCustId.Caption + "'", mdiform1.txtlevel.text)
Set m_cust1 = m_data.QUERY_HIST_CUST(M_OBJCONN, "CUSTID = '" + lblcustid.Caption + "'")
While Not m_cust1.EOF
    'Set listitem = ListView1(1).ListItems.ADD(, , Left(m_cust1("TGL"), 4) & "/" & Mid(m_cust1("TGL"), 5, 2) & "/" & IIf(IsNull(m_cust1("TGL")), "", Mid(m_cust1("TGL"), 7, 2)) & " " & IIf(IsNull(m_cust1("TGL")), "", Mid(m_cust1("TGL"), 9, 2)) & ":" & Right(m_cust1("TGL"), 2))
     Set listitem = ListView1(1).ListItems.ADD(, , Format(IIf(IsNull(m_cust1("TGL")), "", m_cust1!TGL), "mm-dd-yyyy hh:mm:ss"))
        listitem.SubItems(1) = IIf(IsNull(m_cust1("HST")), "", m_cust1("HST"))
        listitem.SubItems(2) = IIf(IsNull(m_cust1("user_log")), "", m_cust1("user_log"))
        listitem.SubItems(3) = IIf(IsNull(m_cust1("AGENT")), "", m_cust1("AGENT"))
        listitem.SubItems(4) = IIf(IsNull(m_cust1("KodeDs")), "", m_cust1("KodeDs"))
        listitem.SubItems(5) = IIf(IsNull(m_cust1("statuscall")), "", m_cust1("statuscall"))
        listitem.SubItems(6) = IIf(IsNull(m_cust1("ststelpwith")), "", m_cust1("ststelpwith"))
        listitem.SubItems(7) = IIf(IsNull(m_cust1("id")), "", m_cust1("id"))
        'listitem.SubItems(4) = IIf(IsNull(m_cust1("f_cek")), "", m_cust1("f_cek"))
m_cust1.MoveNext
Wend

Call isi_datapayment
Call Show_NEGOPTP
Call Show_Reserve
Call Show_Visit
Call Isi_listScript
Call Isi_SendSMS

Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open "Select custid, sum(payment) as jml from tbllunas where custid = '" + lblcustid.Caption + "' GROUP BY CUSTID", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not M_objrs.EOF
        TxtAfterPay.Value = IIf(IsNull(M_objrs("jml")), 0, M_objrs("jml"))
        M_objrs.MoveNext
Wend
 
 'hitung sisa hutang
 txtSisaHutang.Value = Val(txtpayment2.Value) - Val(TxtAfterPay.Value)
 
 '---------->> hitung PRINCIPLE & AMOUNTWO  after pay  <<-----------------
 If TxtAfterPay.Value = 0 Then
    txtPrinciple_A.Value = 0
    txtAmountwo_A.Value = 0
    Else
    If LblPrompA.ValueIsNull Or lblAmount.ValueIsNull Then
    Exit Sub
    End If
  txtPrinciple_A.Value = Val(LblPrompA.Value) - Val(TxtAfterPay.Value)
  txtAmountwo_A.Value = Val(lblAmount.Value) - Val(TxtAfterPay.Value)
 End If
 
    If lblAmount.ValueIsNull Then
           Woafter.Value = 0
       Else
           Woafter.Value = lblAmount - TxtAfterPay.Value
    End If
  
    If ListView1(0).ListItems.Count <> 0 Then
          'lblPayDt.Value = ListView1(0).ListItems(ListView1(0).ListItems.Count).Text
          'lblLastPay.Value = ListView1(0).ListItems(ListView1(0).ListItems.Count).SubItems(1)
          LBLEXP.Caption = "Expire Date " + glexp
    End If
 
 
    Set m_cust = Nothing
    Set M_objrs = Nothing

Exit Sub
'HELL:
   'MsgBox Err.Description
' Resume
 Set M_objrs = Nothing
Set m_cust = Nothing
End Sub

Function ReplaceFirstInstance(SourceString, _
Searchstring, Replacestring)
  'Static StartLoc
  If StartLoc = 0 Then StartLoc = 1
  FoundLoc = InStr(StartLoc, SourceString, Searchstring) '*
  If FoundLoc <> 0 And FoundLoc < 2 Then
     ReplaceFirstInstance = Left(SourceString, FoundLoc - 1) & Replacestring & Right(SourceString, Len(SourceString) - (FoundLoc - 1) - Len(Searchstring))
     StartLoc = FoundLoc + Len(Replacestring)
  ElseIf FoundLoc > 1 Then
  
      ReplaceFirstInstance = Replacestring & "21" & SourceString

  Else
     StartLoc = 1

    ReplaceFirstInstance = SourceString
  End If
End Function

Function FindReplace(SourceString, Searchstring, Replacestring) As String
  tmpString1 = SourceString
 
      tmpString2 = tmpString1
      tmpString1 = ReplaceFirstInstance(tmpString1, _
                   Searchstring, Replacestring)
      
      FindReplace = tmpString1
End Function
Private Sub Isi_SendSMS()
'@@ 11-03-2011 di remarks, cznya udah tidak diapke
'Dim satu As String
'Dim dua As String
'Dim tiga As String
'Dim empat As String
'
'
'Dim RSsms_i As ADODB.Recordset
'Set RSsms_i = New ADODB.Recordset
'
'
'satu = FindReplace(TxtMobileno1.Text, "0", "+62")
'dua = FindReplace(TxtMobileno2.Text, "0", "+62")
'tiga = FindReplace(TxtMobileAdd1.Text, "0", "+62")
'empat = FindReplace(TxtMobileAdd2, "0", "+62")
'
'cmdsql_inbox = "Select receivingdatetime, sendernumber, textdecoded from inbox where (sendernumber='" + Trim$(satu) + "' or sendernumber='" + Trim$(dua) + "' or sendernumber='" + Trim$(tiga) + "' or sendernumber='" + Trim$(empat) + "') and processed='FALSE' "
'RSsms_i.Open cmdsql_inbox, M_OBJCONN1, adOpenDynamic, adLockOptimistic
'While Not RSsms_i.EOF
's = Format(RSsms_i!receivingdatetime, "yyyy-mm-dd hh:mm:ss")
't = Trim(RSsms_i!sendernumber)
'u = Replace(RSsms_i!textdecoded, "'", " ")
'
''u1 = Replace(KATAUBAH, "- -", "-")
'v = FindReplace(t, "+62", "0")
'
'
'
'            CMDSQL = "INSERT INTO receive_sms (tgl_terima, notelp, pesan) VALUES ('" & s & "',"
'            CMDSQL = CMDSQL + " '" + v + "',"
'            CMDSQL = CMDSQL + " '" + u + "')"
'            M_OBJCONN.Execute CMDSQL
'
'            cmdsql_update = "update inbox set processed='TRUE'  where (sendernumber='" + Trim$(satu) + "' or sendernumber='" + Trim$(dua) + "' or sendernumber='" + Trim$(tiga) + "' or sendernumber='" + Trim$(empat) + "')"
'            M_OBJCONN1.Execute cmdsql_update
'
'
'RSsms_i.MoveNext
'Wend
'
''=======================================
'Dim RSsms As ADODB.Recordset
'Set RSsms = New ADODB.Recordset
'Dim lst As listitem
'RSsms.CursorLocation = adUseClient
'If Left(TxtMobileno1, 1) <> "0" And TxtMobileno1 <> "" Then
'satua = "021" & TxtMobileno1
'Else
'satua = TxtMobileno1
'End If
'
'If Left(TxtMobileno2, 1) <> "0" And TxtMobileno2 <> "" Then
'duaa = "021" & TxtMobileno2
'Else
'duaa = TxtMobileno2
'End If
'
'If Left(TxtMobileAdd1, 1) <> "0" And TxtMobileAdd1 <> "" Then
'tigaa = "021" & TxtMobileAdd1
'Else
'tigaa = TxtMobileAdd1
'End If
'
'If Left(TxtMobileAdd2, 1) <> "0" And TxtMobileAdd2 <> "" Then
'empata = "021" & TxtMobileAdd2
'Else
'empata = TxtMobileAdd2
'End If
'
'
'CMDSQL = "Select a.*, b.custid from receive_sms a, mgm b where (a.notelp='" + satua + "' or a.notelp='" + duaa + "' or a.notelp='" + tigaa + "' or a.notelp='" + empata + "') and b.custid='" + lblCustId + "'"
'RSsms.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
'While Not RSsms.EOF
'    Set lst = LstSMS.ListItems.ADD(, , IIf(IsNull(RSsms("notelp")), "", RSsms("notelp")))
'         lst.SubItems(1) = lblNama
'         lst.SubItems(2) = IIf(IsNull(RSsms("custid")), "", RSsms("custid"))
'         lst.SubItems(3) = IIf(IsNull(RSsms("pesan")), "", RSsms("pesan"))
'         lst.SubItems(4) = IIf(IsNull(RSsms("tgl_terima")), "", RSsms("tgl_terima"))
'
'RSsms.MoveNext
'Wend
'Set RSsms = Nothing
'Text3 = LstSMS.ListItems.Count
'
''--------------------------------
'If Text4.Text <> "0" Then
'If Int(Text3) > Int(Text2) Then
'
'Dim RSsms_cek As ADODB.Recordset
'Set RSsms_cek = New ADODB.Recordset
'
'RSsms_cek.CursorLocation = adUseClient
'cmdsql_cek = "select * from receive_sms order by tgl_terima desc limit 1"
'RSsms_cek.Open cmdsql_cek, M_OBJCONN, adOpenDynamic, adLockOptimistic
'While Not RSsms_cek.EOF
'MsgBox "Anda mendapatkan satu SMS baru" & vbCrLf & "No Telepon : " & RSsms_cek("notelp") & vbCrLf & "Isi Pesan : " & Trim(RSsms_cek("pesan"))
'RSsms_cek.MoveNext
'Wend
'Set RSsms_cek = Nothing
'End If
'End If
'
'Text4.Text = "1"

End Sub
Private Sub Isi_SendSMS2()

Dim RSsms2 As ADODB.Recordset
'@@ 11-03-2011 Di remarks, udah tidak diapakai

'Set RSsms2 = New ADODB.Recordset
'Dim Lst2 As listitem
'RSsms2.CursorLocation = adUseClient
'CMDSQL = "Select * from sentitems where destinationnumber='" + TxtMobileno1 + "' or destinationnumber='" + TxtMobileno2 + "' or destinationnumber='" + TxtMobileAdd1 + "' or destinationnumber='" + TxtMobileAdd2 + "'"
'RSsms2.Open CMDSQL, M_OBJCONN1, adOpenDynamic, adLockOptimistic
'While Not RSsms2.EOF
'    Set Lst2 = LstSMS2.ListItems.ADD(, , IIf(IsNull(RSsms2("destinationnumber")), "", RSsms2("destinationnumber")))
'         Lst2.SubItems(1) = lblNama
'         Lst2.SubItems(2) = IIf(IsNull(RSsms2("textdecoded")), "", RSsms2("textdecoded"))
'         Lst2.SubItems(3) = IIf(IsNull(RSsms2("sendingdatetime")), "", RSsms2("sendingdatetime"))
'         Lst2.SubItems(4) = lblCustId
'         'Lst.SubItems(5) = IIf(IsNull(RSsms2("receivingdatetime")), "", RSsms2("receivingdatetime"))
''
'RSsms2.MoveNext
'Wend
'Set RSsms2 = Nothing
End Sub

Private Sub Isi_listScript()
'Mengisi Data di List LstScript
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open "select * from tblinformationlokasi", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not M_objrs.EOF
  Set listitem = Lstscript.ListItems.ADD(, , M_objrs.Bookmark)
      listitem.SubItems(1) = M_objrs("description")
      listitem.SubItems(2) = M_objrs("direktori")
  M_objrs.MoveNext
Wend
Set M_objrs = Nothing
End Sub

Private Sub isi_datapayment()
Dim m_cust2 As New ADODB.Recordset
Dim NilaiAfterPay As Currency
Dim m_data As New CLS_FRMCUST_CC
Set m_cust2 = m_data.QUERY_HIST_PAID(M_OBJCONN, "a.custid = '" + lblcustid.Caption + "' ")
ListView1(0).ListItems.CLEAR
While Not m_cust2.EOF
    Set listitem = ListView1(0).ListItems.ADD(, , IIf(IsNull(m_cust2("Paydate")), "", m_cust2("Paydate")))
        listitem.SubItems(1) = IIf(IsNull(m_cust2("payment")), "0", Format(m_cust2("Payment"), "##,###"))
        listitem.SubItems(2) = IIf(IsNull(m_cust2("AGENT")), "", m_cust2("AGENT"))
        listitem.SubItems(3) = IIf(IsNull(m_cust2("FieldName")), "", m_cust2("FieldName"))
        listitem.SubItems(4) = IIf(IsNull(m_cust2("Id")), "0", m_cust2("Id"))
        NilaiAfterPay = NilaiAfterPay + IIf(IsNull(m_cust2("payment")), "0", m_cust2("Payment"))
    m_cust2.MoveNext
Wend
Set m_cust2 = Nothing
TxtAfterPay.Value = NilaiAfterPay
txtSisaHutang.Value = Format(txtpayment2.Value - TxtAfterPay.Value, "##,###")
End Sub
Private Sub Show_Visit()
Dim m_cust2 As New ADODB.Recordset
Dim m_Visit As New ClsVisit
Dim jml As String
Dim cmdsql As String
Set m_cust2 = New ADODB.Recordset
cmdsql = "SELECT requestdate,visitdate,detailsR,detailsV,visitke,VisitNo,id,F_CEK_new FROM tblvisit where custid='" + lblcustid.Caption + "'"
m_cust2.CursorLocation = adUseClient
m_cust2.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'Set m_cust2 = m_Visit.SELECT_RequestVisit(M_OBJCONN, lblCustId.Caption)
LstVisit.ListItems.CLEAR
While Not m_cust2.EOF
    Set listitem = LstVisit.ListItems.ADD(, , IIf(IsNull(m_cust2!RequestDate), "", m_cust2!RequestDate))
        listitem.SubItems(1) = IIf(IsNull(m_cust2!VisitDate), "", m_cust2!VisitDate)
        listitem.SubItems(2) = Trim(IIf(IsNull(m_cust2!VisitNo), "", m_cust2!VisitNo))
        listitem.SubItems(3) = IIf(IsNull(m_cust2!DetailsR), "", m_cust2!DetailsR)
        listitem.SubItems(4) = IIf(IsNull(m_cust2!DetailsV), "", m_cust2!DetailsV)
        listitem.SubItems(5) = IIf(IsNull(m_cust2!VisitKe), "0", m_cust2!VisitKe)
        listitem.SubItems(6) = IIf(IsNull(m_cust2!ID), "0", m_cust2!ID)
        listitem.SubItems(7) = IIf(IsNull(m_cust2!f_cek_new), "0", m_cust2!f_cek_new)
        m_cust2.MoveNext
Wend
jml = m_cust2.RecordCount + 1
TDBNumber1.Value = jml
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

    Dim m_data As New CLS_FRMCUST_CC_MGM
    Dim m_Visit As New ClsVisit
    Dim pStatusHstLstCall As String
    Dim StatusPTP As String

    Dim M_objrs As ADODB.Recordset
    Dim cmdsql_waktu As String
    Dim waktu As String

    cmdsql_waktu = "select now() as waktu"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open cmdsql_waktu, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    waktu = CDate(Format(M_objrs("waktu"), "hh:nn:ss"))
    Set M_objrs = Nothing


    Set M_update = New ADODB.Recordset
    M_update.CursorLocation = adUseServer
    M_update.Open "Select * from mgm where custid='" & lblcustid.Caption & "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
        
    M_update("AHOMENOADD1") = AHomeAdd1(0).Value
    M_update("AHOMENOADD2") = AHomeAdd2(1).Value
    M_update("AOFFICENOADD1") = AOfficeAdd(2).Value
    M_update("AOFFICENOADD2") = AOfficeAdd(3).Value
    M_update!maxad = tdbmaxad.Value
    M_update!minad = tdbminad.Value
    vrcekamont = Tdabamoint.Value
    If UCase(Left(MDIForm1.txtlevel.Text, 5)) = "ADMIN" Or _
        UCase(MDIForm1.txtlevel.Text) = "SUPERVISOR" Or _
       UCase(MDIForm1.txtlevel.Text) = "TEAMLEADER" Then
        M_update("HOMENOADD1") = txtHomeAdd1.Value
        M_update("HOMENOADD2") = txtHomeAdd2.Value
        M_update("OFFICENOADD1") = txtOfficeAdd1.Value
        M_update("OFFICENOADD2") = txtOfficeAdd2.Value
        M_update("MOBILENOADD1") = txtMobileAdd1.Value
        M_update("MOBILENOADD2") = txtMobileAdd2.Value
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
        ElseIf txtHomeAdd2.Value = "" And txtHomeAdd2.Visible = True Then
            M_update("HOMENOADD2") = txtHomeAdd2.Value
        End If
                
        If txtOfficeAdd1A.Value = "" And txtOfficeAdd1A.Visible = True Then
            M_update("OFFICENOADD1") = txtOfficeAdd1A.Value
        ElseIf txtOfficeAdd1.Value <> "" And txtOfficeAdd1.Visible = True Then
            M_update("OFFICENOADD1") = txtOfficeAdd1.Value
        ElseIf txtOfficeAdd1.Value = "" And txtOfficeAdd1.Visible = True Then
            M_update("OFFICENOADD1") = txtOfficeAdd1.Value
        End If
                
        If txtOfficeAdd2A.Value = "" And txtOfficeAdd2A.Visible = True Then
            M_update("OFFICENOADD2") = txtOfficeAdd2A.Value
        ElseIf txtOfficeAdd2.Value <> "" And txtOfficeAdd2.Visible = True Then
            M_update("OFFICENOADD2") = txtOfficeAdd2.Value
        ElseIf txtOfficeAdd2.Value = "" And txtOfficeAdd2.Visible = True Then
            M_update("OFFICENOADD2") = txtOfficeAdd2.Value
        End If
            
        If txtMobileAdd1A.Value = "" And txtMobileAdd1A.Visible = True Then
            M_update("MOBILENOADD1") = txtMobileAdd1A.Value
        ElseIf txtMobileAdd1.Value <> "" And txtMobileAdd1.Visible = True Then
            M_update("MOBILENOADD1") = txtMobileAdd1.Value
        ElseIf txtMobileAdd1.Value = "" And txtMobileAdd1.Visible = True Then
            M_update("MOBILENOADD1") = txtMobileAdd1.Value
        End If
            
        If txtMobileAdd2A.Value = "" And txtMobileAdd2A.Visible = True Then
            M_update("MOBILENOADD2") = txtMobileAdd2A.Value
        ElseIf txtMobileAdd2.Value <> "" And txtMobileAdd2.Visible = True Then
            M_update("MOBILENOADD2") = txtMobileAdd2.Value
        ElseIf txtMobileAdd2.Value = "" And txtMobileAdd2.Visible = True Then
            M_update("MOBILENOADD2") = txtMobileAdd2.Value
        End If
            
        M_update!TxtPtpAddr = AddrNow.Text
        M_update!ec_name = TxtEC.Text
        M_update!ECAddr = txtECAdd.Text
                 
        If txtECnoA.Value = "" And txtECnoA.Visible = True Then
            M_update("ec_telp") = txtECnoA.Value
        ElseIf txtECno.Value <> "" And txtECno.Visible = True Then
            M_update!ec_telp = txtECno.Value
        End If
    End If
        
    If UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
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
    
    '@@121110 Tambahan nih buat nyatet history perubahan status account
    If (IsNull(M_update!tglcall)) = True Then
        tglcalllalu = ""
    Else
        tglcalllalu = CStr(M_update("tglcall"))
    End If
        
    M_update("TGLCALL") = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & waktu
    'sebelum f_cek diubah statusnya
    StatusPTP = IIf(IsNull(M_update!f_cek_new), "", M_update!f_cek_new)

    Dim StatusAccCurrent As String  '@@ 121110 tambahan nih buat nyatet history f_cek_new
        
    If C_PTP.Value = vbChecked Then
        GoTo keptp
    End If
        
    If cboaccount.Text <> "" Then
        pStatusLstCall = cboaccount.Text
        M_update!f_cek_new = Left(cboaccount.Text, 3)
        txtResult.Text = pStatusLstCall
        '@@121110 tambahan buat nyatet history f_cek_new
        StatusAccCurrent = Left(cboaccount.Text, 3)
    Else
keptp:
        If C_PTP.Value Then
            M_update!ptpdesc = cboaccount.Text
            
            '//////////////////////// Awal Logika PTP 1 ////////////////////////////////////////////
            If vrcek = "BP-" And Len(TglPTPNew) > 0 And UCase(cboPTP.Text) = "PTP-NEW" Then
                M_update!TglPTPNew = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(Now, "hh:nn")
                                        
                    If TDBDate1.ValueIsNull Then
                        M_update!dateptpnew = Null
                    Else
                        M_update!dateptpnew = Format(TDBDate3.Value, "yyyy-mm-dd")
                    End If
                                        
                    If Tdabamoint.ValueIsNull Then
                        M_update!AmountNew = 0
                    Else
                        M_update!AmountNew = Tdabamoint.Value
                    End If
                                                         
            Else
                If cboPTP.Text = "PTP-NEW" Then
                    If vrcek <> "PTP-NE" Then
                    
                        If UCase(cboPTP.Text) = "PTP-NEW" And ListView1(0).ListItems.Count = 0 Then
                            M_update!TglPTPNew = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(Now, "hh:nn")
                            If TDBDate1.ValueIsNull Then
                                M_update!dateptpnew = Null
                            Else
                                M_update!dateptpnew = Format(TDBDate3.Value, "yyyy-mm-dd")
                            End If
                                        
                            If Tdabamoint.ValueIsNull Then
                                M_update!AmountNew = 0
                            Else
                                M_update!AmountNew = Tdabamoint.Value
                            End If
                        End If
                                                    
                    End If
                End If
            End If
            '//////////////////////// Akhir Logika PTP 1 ////////////////////////////////////////////
            
            '//////////////////////// Awal Logika PTP 2 ////////////////////////////////////////////
            If vrcek = "BP-" And Len(TglPTPNew) > 0 And Left(UCase(cboPTP.Text), 3) = "PTP" Then
                M_update!tglallptp = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(Now, "hh:nn")
            Else
                If Left(cboPTP.Text, 3) = "PTP" Then
                    If Left(vrcek, 6) <> Left(cboPTP.Text, 6) Then
                        M_update!tglallptp = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(Now, "hh:nn")
                    ElseIf vrnewdate <> TDBDate3.Text Then
                        M_update!tglallptp = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(Now, "hh:nn")
                    End If
                End If
            End If
            '//////////////////////// Akhir Logika PTP 2 ////////////////////////////////////////////
    
            pStatusLstCall = cboPTP.Text
            txtResult.Text = pStatusLstCall
            txtResultDesc.Text = pStatusLstCalldesc
            M_update("RECSTATUS") = "P"
            M_update!f_cek_new = Left(cboPTP.Text, 6)
                                
            '@@121110 tambahan buat nyatet history f_cek_new
            StatusAccCurrent = Left(cboPTP.Text, 6)
            
        Else
        End If
    End If
        
    If C_Payment.Value Then
        If StatusPTP <> Empty Then
            If StatusPTP = M_update!f_cek_new Then
            Else
                M_update!TGLINCOMING = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(Now, "hh:nn")
            End If
        End If
        M_update!ttlptp = txtPayment.Value
        M_update!amountptp = Tdabamoint.Value
        M_update!discpersen = cmbDiscount.Text
        M_update!Tenor = txttenor.Value
        M_update!dateptp = Format(TDBDate3.Value, "yyyy/mm/dd")
    Else
        M_update!ttlptp = 0
        M_update!discpersen = 0
    End If
               
    If Trim(UCase(IIf(IsNull(M_update("kethslkerja_new")), "", M_update("kethslkerja_new")))) = Trim(UCase(pStatusLstCall)) Then
        TGLSTATUS = IIf(IsNull(M_update("TGLSTATUS")), "", Format(M_update("TGLSTATUS"), "yyyy/mm/dd"))
    Else
        M_update("kethslkerja_new") = pStatusLstCall
        M_update("TGLSTATUS") = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(Now, "hh:nn")
        TGLSTATUS = Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")
    End If
        M_update!stscallwith = cbolastcall.Text
        M_update("kethslkerja_new") = pStatusLstCall
        pStatusHstLstCall = IIf(IsNull(M_update("kethslkerja_new")), "", M_update("kethslkerja_new"))
        M_update("kethslkerjadesc_new") = cboaccount.Text
        M_update("REMARKS") = Replace(TxtRemarks.Text, "'", "`")
    If Not (cmbDateSch.ValueIsNull) Then
        M_update!NEXTACTDATE = Format(cmbDateSch.Value, "yyyy/mm/dd") & " " & Format(Now, "hh:nn")
    End If
        
    M_update("Statuscall") = cbolastcall.Text
    M_update("stscallcust") = Combo1.Text
        
    '@@ 12-11-10 ini nambahin history perubahan status f_cek_new
    'If statusptp <> "" Or IsNull(statusptp) = False Then
'            Dim HISTORYFCEK As String
'            'HISTORYFCEK = IIf(IsNull(M_update("f_cekhst")), "AWAL", M_update("f_cekhst")) + " > " + statusptp + " [" + CStr(tglcalllalu) + "] " + " > " + StatusAccCurrent + " [" + CStr(M_update("tglcall")) + "] "
'            HISTORYFCEK = IIf(IsNull(M_update("f_cekhst")), "AWAL", M_update("f_cekhst")) + " > " + statusptp + " | " + CStr(tglcalllalu) + " "
'            M_update("f_cekhst") = HISTORYFCEK
    'End If
    M_update.update
    
    If C_PTP.Value = vbChecked Then
        GoTo BRO
    End If

    If cboaccount.Text <> "" Then
        If TxtRemarks.Text <> Empty Then
            'M_DATA.ADD_HISTORY lblCustId.Caption, MDIForm1.TDBDate1.Text, Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(CDate(waktu), "hh:nn:ss"), Trim(lblaoc.Caption), "COLLECTION", txtremarks.Text, txtResult.Text, "", "", CStr(Left(IIf(IsNull(M_update!f_cek_new), "", M_update!f_cek_new), 3)), Combo1.Text, Combo1.Text, CStr(Left(IIf(IsNull(M_update!f_cek_new), "", M_update!f_cek_new), 3)), Left(cboaccount.Text, 3), cbolastcall.Text, mdiform1.txtusername.text
        End If
    End If
    
BRO:
    If C_PTP.Value = 1 Then
        If TxtRemarks.Text <> Empty Then
            'M_DATA.ADD_HISTORY lblCustId.Caption, MDIForm1.TDBDate1.Text, Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd") & " " & Format(CDate(waktu), "hh:nn:ss"), Trim(lblaoc.Caption), "COLLECTION", txtremarks.Text, txtResult.Text, "", "", CStr(Left(IIf(IsNull(M_update!f_cek_new), "", M_update!f_cek_new), 3)), Combo1.Text, Combo1.Text, CStr(Left(IIf(IsNull(M_update!f_cek_new), "", M_update!f_cek_new), 3)), Left(cboPTP.Text, 5), cbolastcall.Text, mdiform1.txtusername.text
        End If
    End If

    If Len(TDBTot_payment) > 2 Then
        m_data.ADD_tbllunas M_OBJCONN, lblcustid.Caption, Format(TdbLunas.Value, "yyyy/mm/dd"), CCur(TDBTot_payment.Value), VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11), TxtFieldName.Text, ""
    Else
        On Error Resume Next
    End If
    '------------>> simpan ke table Visit <<--------------------
    If Option8(0).Value Then
        m_Visit.ADD_RequestVisit M_OBJCONN, lblcustid.Caption, M_update!f_cek_new, Text1.Text, Format(TDBDate1.Value, "yyyy-mm-dd"), TXtDetails.Text, TDBNumber1.Value, TxtAddress.Text, Trim(UCase(VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11)))
    Else
        On Error Resume Next
    End If

    MsgBox "Data Sudah Tersimpan", vbInformation + vbOKOnly, "Sukses"
    
    kontak = False
    Set M_update = Nothing

    If shedulePTP_Show = True Then
    Else
        VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(7) = TxtRemarks.Text
        VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(8) = pStatusLstCall
        If cboaccount <> "" Then
            VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(10) = Left(cboaccount, 3)
        ElseIf cboPTP <> "" Then
            VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(10) = Left(cboPTP, 6)
        End If
    End If
    pStatusLstCall = ""
    pStatusHstLstCall = ""
    TxtRemarks.Text = Empty


    Set m_data = Nothing
    Exit Sub
    Resume
End Sub

'@@ 11-03-2011 Di remarks, udah tidak diapakai
'Private Sub HEADER_SendSMS()
'LstSMS.ColumnHeaders.ADD 1, , "No Telp", 5 * TXT
'LstSMS.ColumnHeaders.ADD 2, , "Nama", 5 * TXT
'LstSMS.ColumnHeaders.ADD 3, , "Custid", 15 * TXT
'LstSMS.ColumnHeaders.ADD 4, , "Pesan", 5 * TXT
'LstSMS.ColumnHeaders.ADD 5, , "Tanggal Terima", 5 * TXT
'
'LstSMS2.ColumnHeaders.ADD 1, , "Sender", 5 * TXT
'LstSMS2.ColumnHeaders.ADD 2, , "Nama", 5 * TXT
'LstSMS2.ColumnHeaders.ADD 3, , "Pesan", 15 * TXT
'LstSMS2.ColumnHeaders.ADD 4, , "Jam", 5 * TXT
'LstSMS2.ColumnHeaders.ADD 5, , "Custid", 5 * TXT
'End Sub


Private Sub HEADER_HISTORY()
    ListView1(1).ColumnHeaders.ADD 1, , "Tanggal(mm-dd-yyyy)", 10 * TXT
    ListView1(1).ColumnHeaders.ADD 2, , "History", 70 * TXT
    ListView1(1).ColumnHeaders.ADD 3, , "User Log", 10 * TXT
    ListView1(1).ColumnHeaders.ADD 4, , "Handle By", 10 * TXT
    ListView1(1).ColumnHeaders.ADD 5, , "Sts Account", 10 * TXT
    ListView1(1).ColumnHeaders.ADD 6, , "Sts Call", 10 * TXT
    ListView1(1).ColumnHeaders.ADD 7, , "Sts Telp With", 25 * TXT
    ListView1(1).ColumnHeaders.ADD 8, , "Id", 25 * TXT
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
    ListView1(0).ColumnHeaders.ADD 1, , "PayDate", 15 * TXT
    ListView1(0).ColumnHeaders.ADD 2, , "Payment", 30 * TXT
    ListView1(0).ColumnHeaders.ADD 3, , "Agent", 10 * TXT
    ListView1(0).ColumnHeaders.ADD 4, , "FieldName", 30 * TXT
    ListView1(0).ColumnHeaders.ADD 5, , "Id", 30 * TXT
End Sub
Private Function CEK_DATA_VALID() As Boolean
Dim m_msgbox As Variant
If TDBTot_payment > 2 Then
CEK_DATA_VALID = True
Exit Function
Else

'If mdiform1.txtlevel.text = "TeamLeader" Or mdiform1.txtlevel.text = "Administrator" And (chkAppv(0).Value = 1 Or chkAppv(1).Value = 1) Then
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
    
    
    
    If Combo1.Text = "" Then
            MsgBox "Status Call harus diisi", vbInformation + vbOKOnly, "TINS"
            Combo1.SetFocus
            CEK_DATA_VALID = False
            Exit Function
    End If
    
    
    If cboaccount.Text = "" And C_PTP.Value = vbUnchecked Then
            MsgBox "Status Account harus diisi", vbInformation + vbOKOnly, "TINS"
            CEK_DATA_VALID = False
         Exit Function
    End If
    
    
    If cbolastcall.Text = "" Then
            MsgBox "Status Telepon With harus diisi", vbInformation + vbOKOnly, "TINS"
            cbolastcall.SetFocus
            CEK_DATA_VALID = False
            Exit Function
    End If
    
    
     If C_PTP.Value = vbChecked Then
            If Val(vrcekamont) <> Tdabamoint.Value And bcekptp = False Then
            MsgBox "anda harus klik tambah di Call Activity untuk Negotiation", vbInformation + vbOKOnly, "TINS"
            CEK_DATA_VALID = False
            Exit Function
    End If
    End If
    
    

     If C_Payment.Value = 1 Then
      CmbBaseOn.Text = "TOTAL AMOUNT"
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
      
'    If Left(cmbUncontacted.Text, 2) <> "" Then
'        If cmbDescUn.Text = "" Then
'            CEK_DATA_VALID = False
'            MsgBox "Description UnContacted Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
'            SSTab1.Tab = 3
'            Exit Function
'       End If
'    End If
    
 
     If C_PTP.Value = 1 Then
        If cboPTP.Text = Empty Then
            CEK_DATA_VALID = False
            MsgBox "Description PTP Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
            Exit Function
            SSTab1.Tab = 3
     End If
     End If

       
     If TxtRemarks.Text = "" Then
       CEK_DATA_VALID = False
        MsgBox "Remarks Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
       Exit Function
     End If
     
   
'   If cboaccount.Text <> "" Or C_PTP.Value = 1 Then
'   If cmbDateSch.ValueIsNull = True Or cmbTimeSch.ValueIsNull = True Then
'                CEK_DATA_VALID = False
'                MsgBox "Tanggal Schedule Harus Di isi", vbCritical + vbOKOnly, "Peringatan"
'                SSTab1.Tab = 3
'                Exit Function
'            End If
'   End If
   
 
 
  
    If ADD_CUST = True Then
    Else
'    If C_Contacted.Value = 1 Or C_VALID.Value = 1 Or C_PTP.Value = 1 Or C_SKIP.Value = 1 Or cboPOPSP.Text <> "" Then
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
            If cboaccount.Text <> "" Then
''                'txtRemarks.Text = cmbContacted & " -" & cmbDescCon & " - " & txtRemarks.Text
''                If cmbDescCon.Text = "" Then
''                    txtRemarks.Text = cmbContacted & " - " & "Contac with " & Cmbwith.Text & " - " & cbolastcall.Text & " - " & txtRemarks.Text
''                Else
                    TxtRemarks.Text = Combo1.Text & " - " & cbolastcall.Text & " - " & TxtRemarks.Text
''                End If
             ElseIf cboPTP.Text <> "" Then
                 TxtRemarks.Text = Combo1.Text & " - " & cbolastcall.Text & " - " & " - " & TxtRemarks.Text
          End If
'
'    End If
        If stscall = True Then
            If C_PTP.Value = vbUnchecked And cboaccount.Text = "" Then
                        CEK_DATA_VALID = False
                        MsgBox "Status Account Harus Diisi", vbCritical + vbOKOnly, "Peringatan"
                        SSTab1.Tab = 3
                        Exit Function
            End If
        End If
'            If C_NotContacted.Value = 1 Then
'                'txtRemarks.Text = cmbUncontacted & " -" & cmbDescUn & " - " & txtRemarks.Text
'                txtRemarks.Text = cmbUncontacted & " - " & cmbDescUn & " - " & cbolastcall.Text & " - " & txtRemarks.Text
'            End If
    End If
'    If C_Payment.Value = 1 Then
'        If CmbBaseOn.Text = "" Then
'                If Left(cmbContacted.Text, 3) <> "POP" Then
'                    MsgBox "Base On harus diisi", vbInformation + vbOKOnly, "TINS"
'                    CEK_DATA_VALID = False
'                    Exit Function
'                End If
'        End If
        
        If cmbDiscount.Text = "" Then
            MsgBox "Diskon harus diisi", vbInformation + vbOKOnly, "TINS"
            CEK_DATA_VALID = False
            Exit Function
        End If
        End If

End If
'cek valid uncontacted pending
'If C_Contacted.Value = 1 Then
'    Cmbwith.Text = Replace(Cmbwith, "Customer", "Ch")
'    Cmbwith.Text = Replace(Cmbwith, "Spouse", "SPS")
'    txtRemarks.Text = UBAH_STRIP(Left(cmbContacted.Text, 3) & " -" & cmbDescCon & " - " & txtRemarks.Text)
'    If cmbDescCon.Text = "" Then
'        txtRemarks.Text = UBAH_STRIP(Left(cmbContacted.Text, 3) & " - " & "" & Cmbwith.Text & " - " & Left(cbolastcall.Text, 3) & " - " & txtRemarks.Text)
'    Else
'        txtRemarks.Text = UBAH_STRIP(Left(cmbContacted.Text, 3) & " - " & "" & Cmbwith.Text & " - " & cmbDescCon & " - " & Left(cbolastcall.Text, 3) & " - " & txtRemarks.Text)
'    End If
'End If

'If C_NotContacted.Value = 1 Then
'    'txtRemarks.Text = cmbUncontacted & " -" & cmbDescUn & " - " & txtRemarks.Text
'    txtremarks.Text =
'
'End If

'If C_VALID.Value = 1 Then
'                If cbodescvalid.Text = "" Then
'                    txtRemarks.Text = UBAH_STRIP(Left(cbovalid, 3) & " - " & Left(cbolastcall.Text, 3) & " - " & txtRemarks.Text)
'                Else
'                    txtRemarks.Text = UBAH_STRIP(Left(cbovalid, 3) & " - " & cbodescvalid & " - " & Left(cbolastcall.Text, 3) & " - " & txtRemarks.Text)
'                End If
'            End If
If C_PTP.Value = 1 Then
        TxtRemarks.Text = TxtRemarks.Text
End If
'If C_SKIP.Value = 1 Then
'    If cbodescskip.Text = "" Then
'        txtRemarks.Text = UBAH_STRIP(Left(cboskip, 3) & " - " & Left(cbolastcall.Text, 3) & " - " & txtRemarks.Text)
'    Else
'        txtRemarks.Text = UBAH_STRIP(Left(cboskip, 3) & " - " & cbodescskip & " - " & Left(cbolastcall.Text, 3) & " - " & txtRemarks.Text)
'    End If
'End If

If regnego = True Then
    Dim n%
    Dim jum As Currency
    For n = 1 To FrmCC_Colection.LstPayment.ListItems.Count
        jum = jum + FrmCC_Colection.LstPayment.ListItems(n).SubItems(3)
    Next n
    If jum < FrmCC_Colection.txtPayment.Value Then
        MsgBox "Jumlah PTP Belum sama dengan Jumlah Deal Payment"
        CEK_DATA_VALID = False
        TxtRemarks.Text = ""
        Exit Function
    End If
End If
regnego = False
CEK_DATA_VALID = True
End Function
Public Sub Custid_Double()
Dim listitem As listitem
Dim test As String
Dim cmdsql As String



Set m_cust = New ADODB.Recordset
m_cust.CursorLocation = adUseClient
test = Format(LblDOB.Caption, "yyyy/mm/dd")

'@@ 26-11-10 Ubah logik double custid, harus cek ktpnya dulu
If Trim(lblID.Caption) <> "" Then
    If LblDOB.Caption <> "" Then
        cmdsql = "Select a.custid, a.name,a.agent, a.amountwo,"
        cmdsql = cmdsql + "a.principal,a.flaglead from mgm a where (a.name='"
        cmdsql = cmdsql + Trim(txtname.Text) + "' and dob='"
        cmdsql = cmdsql + test + "' or ktpno='"
        cmdsql = cmdsql + Trim(lblID.Caption) + "')  and a.custid <> '"
        cmdsql = cmdsql + Trim(lblcustid.Caption) + "'"
    Else
        cmdsql = "Select a.custid, a.name,a.agent, a.amountwo,"
        cmdsql = cmdsql + "a.principal,a.flaglead from mgm a where (a.name='"
        cmdsql = cmdsql + Trim(txtname.Text) + "' "
        cmdsql = cmdsql + " or ktpno='"
        cmdsql = cmdsql + Trim(lblID.Caption) + "')  and a.custid <> '"
        cmdsql = cmdsql + Trim(lblcustid.Caption) + "'"
    End If
Else
    cmdsql = "Select a.custid, a.name,a.agent, a.amountwo,"
    cmdsql = cmdsql + "a.principal,a.flaglead from mgm a where a.name='"
    cmdsql = cmdsql + Trim(txtname.Text) + "' and dob='"
    cmdsql = cmdsql + test + "'"
    cmdsql = cmdsql + " and a.custid <> '"
    cmdsql = cmdsql + Trim(lblcustid.Caption) + "'"
End If


'm_cust.Open "Select a.custid, a.name,a.agent, a.amountwo,a.principal,a.flaglead from mgm a where (a.name='" + Trim(txtname.Text) + "' and dob='" + test + "' or ktpno='" & Trim(lblID.Caption) & "') and a.custid <> '" + Trim(lblCustId.Caption) + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
m_cust.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic

While Not m_cust.EOF
    Set listitem = LstDoubleId.ListItems.ADD(, , IIf(IsNull(m_cust("CUSTID")), "", m_cust("CUSTID")))
        listitem.SubItems(1) = IIf(IsNull(m_cust("NAME")), "", m_cust("NAME"))
        listitem.SubItems(2) = IIf(IsNull(m_cust("AGENT")), "", m_cust("AGENT")) '
      '  If Format(IIf(IsNull(m_cust("flaglead")), 0, m_cust("flaglead")), "##,###") = 1 Then
         '    harga = IIf(IsNull(m_cust("AmountWo")), 0, m_cust("AmountWo"))
           '  harga = harga + (harga * 18.26) / 100
          '   listitem.SubItems(3) = Format(harga, "##,###")
        'Else
            listitem.SubItems(3) = Format(IIf(IsNull(m_cust("AmountWo")), 0, m_cust("AmountWo")), "##,###")
        'End If
        
        
       ' If Format(IIf(IsNull(m_cust("flaglead")), 0, m_cust("flaglead")), "##,###") = 1 Then
        '     harga = IIf(IsNull(m_cust("principal")), 0, m_cust("principal"))
         '    harga = harga + (harga * 26.05) / 100
          '   listitem.SubItems(4) = Format(harga, "##,###")
        'Else
        
        
     If UCase(MDIForm1.txtlevel.Text) <> "SUPERVISOR" Then
        If IIf(IsNull(m_cust("flaglead")), 0, m_cust("flaglead")) = 1 Then
            listitem.SubItems(4) = ""
        Else
           listitem.SubItems(4) = ENCRIPY(False, CStr(Format(IIf(IsNull(m_cust("principal")), 0, m_cust("principal")), "##,###")))
        End If
    Else
            listitem.SubItems(4) = ENCRIPY(False, CStr(Format(IIf(IsNull(m_cust("principal")), 0, m_cust("principal")), "##,###")))
    End If
      
        
   
      
     
       
       ' End If
        
    
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
Dim m_data As New ClsNegoPTP
Dim JmlPay As Double
Dim i As Integer
Dim n As Integer
Dim Vrdate As String
Dim jatuhtempo As String
Select Case Index
    Case 0
        If TDBDate3.ValueIsNull Or Tdabamoint.ValueIsNull Or txttenor.ValueIsNull Then
        MsgBox "Pengisian Data Belum Lengkap (installment,tenor,dateptp) "
        Exit Sub
        End If
        bcekptp = True
           If Chktenor.Value = 0 Then
            jatuhtempo = Format(TDBDate3.Value, "yyyy-mm-dd")
            cmdsql = "INSERT INTO TblNegoPTP "
            cmdsql = cmdsql + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
            cmdsql = cmdsql + "VALUES "
            cmdsql = cmdsql + "('" + lblcustid + "', "
            cmdsql = cmdsql + "'" + Format(jatuhtempo, "yyyy-mm-dd") + "', "
            cmdsql = cmdsql + "" + CStr(Tdabamoint.Value) + " , "
            cmdsql = cmdsql + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
            cmdsql = cmdsql + "'IPO')"
            M_OBJCONN.Execute cmdsql
            ' isi ke tbl log_ptp
            cmdsql = "INSERT INTO tblnegoptp_log "
            cmdsql = cmdsql + "(custid,PromiseDate, Promisepay,tglInput,agent,stsacc) "
            cmdsql = cmdsql + "VALUES "
            cmdsql = cmdsql + "('" + lblcustid + "', "
            cmdsql = cmdsql + "'" + Format(jatuhtempo, "yyyy-mm-dd") + "', "
            cmdsql = cmdsql + "" + CStr(Tdabamoint.Value) + " , "
            cmdsql = cmdsql + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
            cmdsql = cmdsql + "'" + lblaoc.Caption + "','P')"
            M_OBJCONN.Execute cmdsql
            
            Set listitem = LstPayment.ListItems.ADD(, , "")
            listitem.SubItems(1) = ""
            listitem.SubItems(2) = Format(TDBDate3.Value, "dd/mm/yyyy")
            listitem.SubItems(3) = CStr(Tdabamoint.Value)
            listitem.SubItems(4) = "IPO"
            listitem.SubItems(5) = MDIForm1.TDBDate1.Value
            
            Else
            
            
            jatuhtempo = Format(TDBDate3.Value, "yyyy-mm-dd")
            cmdsql = "INSERT INTO TblNegoPTP "
            cmdsql = cmdsql + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
            cmdsql = cmdsql + "VALUES "
            cmdsql = cmdsql + "('" + lblcustid + "', "
            cmdsql = cmdsql + "'" + Format(jatuhtempo, "yyyy-mm-dd") + "', "
            cmdsql = cmdsql + "" + CStr(Tdabamoint.Value) + " , "
            cmdsql = cmdsql + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
            cmdsql = cmdsql + "'IPO')"
            M_OBJCONN.Execute cmdsql
            ' isi ke tbl log_ptp
            
            
            
            cmdsql = "INSERT INTO tblnegoptp_log "
            cmdsql = cmdsql + "(custid,PromiseDate, Promisepay,tglInput,agent,stsacc) "
            cmdsql = cmdsql + "VALUES "
            cmdsql = cmdsql + "('" + lblcustid + "', "
            cmdsql = cmdsql + "'" + Format(jatuhtempo, "yyyy-mm-dd") + "', "
            cmdsql = cmdsql + "" + CStr(Tdabamoint.Value) + " , "
            cmdsql = cmdsql + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
            cmdsql = cmdsql + "'" + lblaoc.Caption + "','P')"
            M_OBJCONN.Execute cmdsql
            
            Set listitem = LstPayment.ListItems.ADD(, , "")
            listitem.SubItems(1) = ""
            listitem.SubItems(2) = Format(TDBDate3.Value, "dd/mm/yyyy")
            listitem.SubItems(3) = CStr(Tdabamoint.Value)
            listitem.SubItems(4) = "IPO"
            listitem.SubItems(5) = MDIForm1.TDBDate1.Value
            
    

    n = 0
    For i = 1 To Val(txttenor - 1)
            n = n + 1
            JmlPay = (txtPayment - Tdabamoint) / (txttenor.Value - 1)
            'VRDATE = Format(DateAdd("m", n, TDBDate3.Value), "mm/dd/yyyy")
            Vrdate = DateAdd("m", n, Format(TDBDate3.Value, "yyyy-mm-dd"))
            cmdsql = "INSERT INTO tblreserve "
            cmdsql = cmdsql + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
            cmdsql = cmdsql + "VALUES "
            cmdsql = cmdsql + "('" + lblcustid + "', "
            cmdsql = cmdsql + "'" + CStr(Format(Vrdate, "yyyy-mm-dd")) + "', "
            cmdsql = cmdsql + "" + CStr(JmlPay) + " , "
            cmdsql = cmdsql + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
            cmdsql = cmdsql + "'IPO')"
            M_OBJCONN.Execute cmdsql
            
            cmdsql = "INSERT INTO TblNegoptp_log "
            cmdsql = cmdsql + "(custid,PromiseDate, Promisepay,tglinput,agent,stsacc) "
            cmdsql = cmdsql + "VALUES "
            cmdsql = cmdsql + "('" + lblcustid + "', "
            cmdsql = cmdsql + "'" + CStr(Format(Vrdate, "yyyy-mm-dd")) + "', "
            cmdsql = cmdsql + "" + CStr(JmlPay) + " , "
            cmdsql = cmdsql + "'" + CStr(Format(MDIForm1.TDBDate1.Value, "yyyy-mm-dd")) + "', "
            cmdsql = cmdsql + "'" + lblaoc.Caption + "','R')"
            M_OBJCONN.Execute cmdsql

        Set listitem = LstReserve.ListItems.ADD(, , "")
            listitem.SubItems(1) = ""
                               'listitem.SubItems(2) = .TDBDate1.Value
            listitem.SubItems(2) = Format(Vrdate, "dd/mm/yyyy")
            listitem.SubItems(3) = JmlPay
            listitem.SubItems(4) = "IPO"
            listitem.SubItems(5) = MDIForm1.TDBDate1.Value
    Next i
   End If
   
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

                    m_data.UPDATE_NegoPTP M_OBJCONN, .TxtCustid.Text, Format(.TDBDate1.Value, "yyyy-mm-dd"), CStr(.TDBNumber1.Value), LstPayment.SelectedItem.SubItems(1)

                    On Error GoTo add_error
                    If m_data.ADD_OK Then
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
    
    Case 3
           frmdeletereserve.Show vbModal
End Select
add_error:
End Sub
Private Sub VisitYES()
Text1.BackColor = &HFF00&
TxtCustid.BackColor = &H80000005
txtname.BackColor = &H80000005
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
txtname.BackColor = &H8000000F
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

Private Sub Tdabamoint_Change()
bcekptp = False
End Sub

Private Sub TdbPTP_Change()
TdbPTP.Value = TDBDate1.Value
End Sub

Private Sub Timer1_Timer()

End Sub

'Private Sub Timer_cek_inbox_Timer()
''@@ 11-03-2011 Di remarks, udah tidak diapakai
''    Text2 = LstSMS.ListItems.Count
''
''    LstSMS.ListItems.CLEAR
''    LstSMS2.ListItems.CLEAR
''    Isi_SendSMS
''    Isi_SendSMS2
'End Sub

Private Sub blink(Seconds As Single)
 Dim a As Long
 Seconds = Seconds + Timer
 While Seconds > Timer
  a = DoEvents
 Wend
End Sub

Private Sub TimerBlink_Timer()
   
               If SSCommand1(7).BackColor = vbRed Then
                 SSCommand1(7).BackColor = vbGreen
                 KelapKelip = KelapKelip + 1
               Else
                 SSCommand1(7).BackColor = vbRed
                 KelapKelip = KelapKelip + 1
               End If
           
           If KelapKelip = 7 Then
            KelapKelip = 0
            WaitSecs (3)
            'TimerBlink.Enabled = False
           End If
    
End Sub

Private Sub TimerBlinkSms_Timer()
    If LabelSms.ForeColor = vbBlack Then
        LabelSms.ForeColor = vbRed
        Command2.BackColor = vbRed
        KelapKelip = KelapKelip + 1
    Else
        LabelSms.ForeColor = vbBlack
        Command2.BackColor = vbYellow
        KelapKelip = KelapKelip + 1
    End If
           
    If KelapKelip = 7 Then
            KelapKelip = 0
            WaitSecs (3)
            'TimerBlink.Enabled = False
    End If
End Sub










Private Sub TimerCekMapping_Timer()
     If CmdDataMapping.BackColor = vbGreen Then
        CmdDataMapping.BackColor = vbRed
        KelapKelip = KelapKelip + 1
    Else
        CmdDataMapping.BackColor = vbYellow
        KelapKelip = KelapKelip + 1
    End If
           
    If KelapKelip = 7 Then
            KelapKelip = 0
            WaitSecs (3)
            'TimerBlink.Enabled = False
    End If
End Sub

'Private Sub TimerCekSms_Timer()
'
'    On Error Resume Next
'    Dim M_OBJRS As New ADODB.Recordset
'    Dim cmdsql34 As String
'    Dim TELPo As String
'    Dim codea As String
'    Dim m_objrscek As ADODB.Recordset
'
'    If Left(mdiform1.txtusername.text, 1) = "D" Or Text1 = "JOKO" Or Text1 = "SPV1" Or Left(mdiform1.txtusername.text, 1) = "T" Then
'        Select Case Text1.Text
'            Case "TL1"
'                codea = "ACC1"
'            Case "TL2"
'                codea = "ACC2"
'            Case "TL3"
'                codea = "ACC3"
'            Case "TL4"
'                codea = "ACC4"
'            Case "TL5"
'                codea = "ACC5"
'            Case "TL6"
'                codea = "ACC6"
'            Case "TL7"
'                codea = "ACC7"
'            Case "TL8"
'                codea = "ACC8"
'            Case "TL9"
'                codea = "ACC9"
'            Case "TL10"
'                codea = "ACC10"
'            Case Else
'                codea = mdiform1.txtusername.text
'        End Select
'
'        TELPo = "Select count(*) as banyak from inbox where sendernumber in ('a',"
'
'        Set M_OBJRS = New ADODB.Recordset
'        M_OBJRS.CursorLocation = adUseClient
'        cmdsql34 = "select mobileno,mobileno2,mobilenoadd1,mobilenoadd2 from mgm where agent = '" + codea + "'"
'        M_OBJRS.Open cmdsql34, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'        If M_OBJRS.RecordCount = 0 Then
'            Timer6.Interval = 60000
'            Exit Sub
'        End If
'
'        While Not M_OBJRS.EOF
'
'            If Len(M_OBJRS("mobileno")) <> 0 Then
'                satu = FindReplace(M_OBJRS("mobileno"), "0", "+62")
'                TELPo = TELPo + "'" + satu + "',"
'            Else
'                TELPo = TELPo
'            End If
'
'            If Len(M_OBJRS("mobileno2")) <> 0 Then
'                dua = FindReplace(M_OBJRS("mobileno2"), "0", "+62")
'                TELPo = TELPo + "'" + dua + "',"
'            Else
'                TELPo = TELPo
'            End If
'
'            If Len(M_OBJRS("mobilenoadd1")) <> 0 Then
'                tiga = FindReplace(M_OBJRS("mobilenoadd1"), "0", "+62")
'                TELPo = TELPo + "'" + tiga + "',"
'            Else
'                TELPo = TELPo
'            End If
'
'            If Len(M_OBJRS("mobilenoadd2")) <> 0 Then
'                empat = FindReplace(M_OBJRS("mobilenoadd2"), "0", "+62")
'                TELPo = TELPo + "'" + empat + "',"
'            Else
'                TELPo = TELPo
'            End If
'
'            M_OBJRS.MoveNext
'        Wend
'        Set M_OBJRS = Nothing
'
'
'        TELPo = Left(TELPo, Len(TELPo) - 1)
'        Dim TELPo1
'
'
'        TELPo1 = TELPo + ") and processed='f'"
'        'TELPo2 = TELPo + ") and processed='t'"
'
'        Set m_objrscek = New ADODB.Recordset
'        m_objrscek.CursorLocation = adUseClient
'        m_objrscek.Open TELPo1, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
''        While Not M_OBJRS.EOF
''            'LblJmlSmsBaru.Caption = M_OBJRS("banyak")
''            LabelSms.Caption = "ADA SMS BARU!" '& LblJmlSmsBaru.Caption & " SMS"
''            M_OBJRS.MoveNext
''        Wend
'
''        'JIKA ADA SMS BARU MASUK
''        If Trim(LabelSms.Caption) = "SMS BARU 0 SMS" Then
''            'MsgBox "Tidak ada sms baru!"
''            TimerBlink.Enabled = False
''            LabelSms.ForeColor = vbBlack
''        Else
''            If Trim(LabelSms.Caption) <> "" Then
''                TimerBlink.Enabled = True
''                MsgBox "Ada SMS BARU MASUK! Silahkan cek!", vbOKOnly + vbInformation, "Informasi"
''            End If
''        End If
'         If m_objrscek(0) > 0 Then
'            TimerBlinkSms.Enabled = True
'            LabelSms.Caption = "Ada SMS Baru!"
'         Else
'            LabelSms.Caption = "Tidak ada SMS baru!"
'            LabelSms.ForeColor = vbBlack
'            Command2.BackColor = vbGreen
'            TimerBlinkSms.Enabled = False
'         End If
'
'        Set m_objrscek = Nothing
'End If
'        Timer6.Interval = 60000
'End Sub



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
Private Sub txtECnoA_DblClick()
txthasil.Text = txtECno.Text
End Sub

Private Sub txthasil_KeyPress(KeyAscii As Integer)
KeyAscii = 0
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

Private Sub txtHomeAdd1A_DblClick()
txthasil.Text = txtHomeAdd1.Text

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

Private Sub txtHomeAdd2A_DblClick()
txthasil.Text = txtHomeAdd2.Text
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
Private Sub txtHomeNo1A_DblClick()
txthasil.Text = txtHomeNo1.Text
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
Private Sub txtHomeNo2A_DblClick()
txthasil.Text = txtHomeNo2.Text
End Sub

Private Sub txtMobileAdd1A_Click()
TYPETELP = "MOBILE1"
    txtPhone.Text = txtMobileAdd1.Value
    txtPhoneA.Text = txtMobileAdd1A.Value
    
    CmbPhone.Text = "AddMobile1"
End Sub

Private Sub txtMobileAdd1A_DblClick()
txthasil.Text = txtMobileAdd1.Text
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

Private Sub txtMobileAdd2A_DblClick()
txthasil.Text = txtMobileAdd2.Text
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

Private Sub txtMobileNo1A_DblClick()
txthasil.Text = txtMobileNo1.Text
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
Private Sub txtMobileNo2A_DblClick()
    txthasil.Text = txtMobileNo2.Text
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
Private Sub txtOfficeAdd1A_DblClick()
    txthasil.Text = txtOfficeAdd1.Text
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
        cmdsql = "update mgm set F_pending='Pending',Agent='DA',PO_Agent='" & lblaoc.Caption & "' where custid='" + lblcustid.Caption + "'"
        M_OBJCONN.Execute cmdsql
        spend = True
        MsgBox "Data berhasil dipindah ke agent DA", vbInformation
        VIEW_MGMDATA.LstVwSearchMgm.ListItems.CLEAR
        MDIForm1.LstGrade.ListItems.CLEAR
    End If
Else
    If chkAppv(1).Value Then
        Dim spo As ADODB.Recordset
        Set spo = New ADODB.Recordset
        spo.CursorLocation = adUseClient
        spo.Open "select PO_Agent from mgm where custid='" + lblcustid.Caption + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic
        
        If spo!PO_AGENT <> "" And IsNull(spo!PO_AGENT) = False Then
            cmdsql = "update mgm set F_pending='',AGENT=PO_Agent where custid='" + lblcustid.Caption + "'"
            M_OBJCONN.Execute cmdsql
            cmdsql = "update mgm set PO_Agent='' where custid='" + lblcustid.Caption + "'"
            M_OBJCONN.Execute cmdsql
            MsgBox "Data berhasil dikembalikan", vbInformation
            VIEW_MGMDATA.LstVwSearchMgm.ListItems.CLEAR
            MDIForm1.LstGrade.ListItems.CLEAR
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

Private Sub txtOfficeAdd2A_DblClick()
txthasil.Text = txtOfficeAdd2.Text
End Sub

Private Sub txtOfficeNo1_Click()
If Len(txtOfficeNo1.Text) > 3 Then
CmbPhone.Text = "OfficePhone"
Else
CmbPhone.Text = ""
End If
End Sub
Private Sub txtOfficeNo1A_DblClick()
 txthasil.Text = txtOfficeNo1.Text
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
Public Sub Show_Reserve()
Dim showlist As New ADODB.Recordset
Dim listitem As listitem
Dim cmdsql As String
Dim TOTPTP As Currency
Dim ssql As String
ssql = "SELECT CUSTID,sum(PAYMENT) as Jum FROM tbllunas WHERE custid = '" + lblcustid.Caption + "' GROUP BY CUSTID"
showlist.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic
If showlist.BOF And showlist.EOF Then
    TOTPTP = 0
Else
    TOTPTP = IIf(IsNull(showlist!jum), 0, showlist!jum)
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
If MDIForm1.txtlevel.Text = "SUPERVISOR" Then
    cmdsql = "SELECT * FROM tblreserve where custid = '" + lblcustid.Caption + "' order by promisedate"
Else
    cmdsql = "SELECT * FROM tblreserve where custid = '" + lblcustid.Caption + "' and stsmove=0 order by promisedate"
End If

Set showlist = New ADODB.Recordset
showlist.CursorLocation = adUseClient
showlist.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic

LstReserve.ListItems.CLEAR
Dim n As Currency
While Not showlist.EOF
    Set listitem = LstReserve.ListItems.ADD(, , "")
        listitem.SubItems(1) = CStr(IIf(IsNull(showlist!ID), "", (showlist!ID)))
        listitem.SubItems(2) = CStr(IIf(IsNull(showlist!PromiseDate), "", Format(showlist!PromiseDate, "dd/mm/yyyy")))
        listitem.SubItems(3) = CStr(IIf(IsNull(showlist!PromisePay), "", (Round(showlist!PromisePay, 1))))
        n = n + Val(listitem.SubItems(3))
        If n <= TOTPTP Then
            listitem.ListSubItems(1).ForeColor = vbRed
            listitem.ListSubItems(2).ForeColor = vbRed
            listitem.ListSubItems(3).ForeColor = vbRed
        End If
        
        listitem.SubItems(4) = IIf(IsNull(showlist!Type), "", showlist!Type)
        listitem.SubItems(5) = CStr(IIf(IsNull(showlist!inputdate), "", Format(showlist!inputdate, "dd/mm/yyyy")))
     showlist.MoveNext
Wend

Set showlist = Nothing
End Sub

Private Sub txtOfficeNo2A_DblClick()
txthasil.Text = txtOfficeNo2.Text
End Sub

Public Sub PesanLockAuto()
    Dim m_objrsPesanReset As ADODB.Recordset
    Dim m_objrsPesanLock As ADODB.Recordset
    Dim M_ObjWktServer As ADODB.Recordset
    Dim WaktuServer As Date
    Dim cmdsql As String
    
    'Ambil Waktu Server Sekarang
    Set M_ObjWktServer = New ADODB.Recordset
    M_ObjWktServer.CursorLocation = adUseClient
    M_ObjWktServer.Open "Select now() as WktSrv ", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    WaktuServer = Format(M_ObjWktServer(0), "yyyy-mm-dd hh:mm")
    Set M_ObjWktServer = Nothing
    
    'Cek pesan reset
    cmdsql = "select f_pesanresetauto,f_idsessend from usertbl where userid='"
    cmdsql = cmdsql + Trim(MDIForm1.txtusername.Text) + "'"
    Set m_objrsPesanReset = New ADODB.Recordset
    m_objrsPesanReset.CursorLocation = adUseClient
    m_objrsPesanReset.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    
    If m_objrsPesanReset.RecordCount <> 0 Then
        If m_objrsPesanReset("f_pesanresetauto") = "1" Then
            MsgBox "Reset Data! Ini adalah lock data automatic, data anda akan segera diperbaharui!", vbOKOnly + vbInformation, "Informasi"
           
            VIEW_MGMDATA.LstVwSearchMgm.ListItems.CLEAR
            '@@20-11-10 akhiri session dengan mencatat hasil akhir perubahan status data yang dikerjain agent
                If m_objrsPesanReset("f_idsessend") <> "" Or IsNull(m_objrsPesanReset("f_idsessend")) = False Or m_objrsPesanReset("f_idsessend") <> Empty Then
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
                    UpdateDtCloseSession = UpdateDtCloseSession + Trim(m_objrsPesanReset("f_idsessend")) + "' and tblperformpersessionlock.agent='"
                    UpdateDtCloseSession = UpdateDtCloseSession + Trim(MDIForm1.txtusername.Text) + "'"
                    M_OBJCONN.Execute UpdateDtCloseSession
                    'bikin null lagi nilai f_idsessend
                    UpdateDtCloseSession = "update usertbl set f_idsessend=null where userid='"
                    UpdateDtCloseSession = UpdateDtCloseSession + Trim(MDIForm1.txtusername.Text) + "'"
                    M_OBJCONN.Execute UpdateDtCloseSession
                End If
            '@@20-11-10 akhiri session dengan mencatat hasil akhir perubahan status data yang dikerjain agent
             
            cmdsql = "update usertbl set f_pesanresetauto=null where userid='"
            cmdsql = cmdsql + Trim(MDIForm1.txtusername.Text) + "'"
            M_OBJCONN.Execute cmdsql
        End If
    End If
    
    Set m_objrsPesanReset = Nothing
    
    'Cek pesan Lock
    cmdsql = "select f_pesanlockauto from usertbl where userid='"
    cmdsql = cmdsql + Trim(MDIForm1.txtusername.Text) + "'"
    Set m_objrsPesanLock = New ADODB.Recordset
    m_objrsPesanLock.CursorLocation = adUseClient
    m_objrsPesanLock.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If m_objrsPesanLock.RecordCount <> 0 Then
        If m_objrsPesanLock("f_pesanlockauto") = "1" Then
            MsgBox "Lock Data! Ini adalah lock data automatic, data anda akan segera diperbaharui!", vbOKOnly + vbInformation, "Informasi"
            cmdsql = "update usertbl set f_pesanlockauto=null where userid='"
            cmdsql = cmdsql + Trim(MDIForm1.txtusername.Text) + "'"
            M_OBJCONN.Execute cmdsql
            VIEW_MGMDATA.LstVwSearchMgm.ListItems.CLEAR
        End If
     End If
    
    Set m_objrsPesanLock = Nothing
End Sub

'@@ 14022011
Private Sub CekSms()
    Dim cmdsql As String
    Dim M_objrs As ADODB.Recordset
    
    '@@ 14/02/2010,, Cek smsnya melalui field blink di usertbl aja, jadinya lebih ringan
    If UCase(Trim(MDIForm1.txtlevel.Text)) = "AGENT" Then
        cmdsql = "select status_sms from usertbl where userid='"
        cmdsql = cmdsql + Trim(MDIForm1.txtusername.Text) + "'"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If M_objrs("status_sms") <> "" Then
            TimerBlinkSms.Enabled = True
            LabelSms.Caption = "Ada SMS Baru!"
        Else
            LabelSms.Caption = "Tidak ada SMS baru!"
            LabelSms.ForeColor = vbBlack
            Command2.BackColor = vbGreen
            TimerBlinkSms.Enabled = False
        End If
        
        Set M_objrs = Nothing
    End If
End Sub



'@@ 08-03-2011 Cek data mapping
Private Sub CekDataMapping()
    Dim cmdsql As String
    Dim M_objrs As ADODB.Recordset
    
    
    cmdsql = "select * from mgm_mapping_pil where custidcard='"
    cmdsql = cmdsql + Trim(lblcustid.Caption) + "' or ktpno='"
    '@@ 25-07-2011 , Tambahan cari juga berdasarkan Nomor KTP
    cmdsql = cmdsql + Trim(lblID.Caption) + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
    
    
    If M_objrs.RecordCount = 0 Then
        CmdDataMapping.BackColor = vbGreen
        TimerCekMapping.Enabled = False
    Else
        TimerCekMapping.Enabled = True
    End If
        
    Set M_objrs = Nothing
End Sub


'@@ 06-May 2011 Tambahan Offering Discon Guide
Private Sub OfferingDiscGuide()
    '@@06 May 2011 Tambahan Offering
        Dim K As Integer
        Dim W As String
        Dim l As Integer
        Dim diskon As Integer
        
        Dim M_objrs As ADODB.Recordset
        Dim m_objrs_waktu As ADODB.Recordset
        Dim cmdsql As String
              
        
        'Cek dulu ada pembayaran apa ngga di tabel lunas
        cmdsql = "select * from tbllunas where custid='"
        cmdsql = cmdsql + Trim(lblcustid.Caption) + "'"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        
        'Ambil waktu sekarang
        cmdsql = "select now() as waktu "
        Set m_objrs_waktu = New ADODB.Recordset
        m_objrs_waktu.CursorLocation = adUseClient
        m_objrs_waktu.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        '@@ 08-06-2011, Jika lblpaydt=kosong on error goto salah
        On Error GoTo salah
        l = DateDiff("M", Format(lblPayDt.Value, "mm/dd/yyyy"), Format(CDate(m_objrs_waktu("waktu")), "mm/dd/yyyy"))
        
        '@@ 09-05-2011 Jika tidak ada nopay atau lpd > 4 bulan dari current date maka
        'tampilkan offering
        
        
        If M_objrs.RecordCount = 0 Or _
            l > 4 Then
            On Error GoTo salah
            K = DateDiff("M", Format(lblOpenDate.Value, "mm/dd/yyyy"), Format(lblBD.Value, "mm/dd/yyyy"))
            If K < 12 Then
                W = "Penawaran Diskon Maximal 60%"
                diskon = 60
            ElseIf K >= 12 And K <= 17 Then
                W = "Penawaran Diskon Maximal 50%"
                diskon = 50
            ElseIf K >= 18 And K <= 36 Then
                W = "Penawaran Diskon Maximal 40%"
                diskon = 40
            ElseIf K > 37 Then
                W = "Cicilan panjang " & " dan diskon 30%"
                diskon = 30
            End If
        
            'MsgBox "Pemandu Offering: " & w, vbOKOnly + vbInformation, "Offering Disc.Guide..."
            'With FrmOfferingGuide
            With FRMSCRIPT
                .LblTextGuide.Caption = "Pemandu Offering: " & W
                .TdbBalance.Value = lblAmount.Value
                .TdbMaxDisc.Value = diskon
                .Show vbModal
            End With
        End If
        
        Set M_objrs = Nothing
        Set m_objrs_waktu = Nothing
        Exit Sub
salah:
    Set M_objrs = Nothing
    Set m_objrs_waktu = Nothing
End Sub


'@@ 09092011, Skrip Ofering yang awalnya di FormOfferingGuide, Sekarang Dipindah ke FormScript
Private Sub OfferingDiscGuideNew()
    '@@06 May 2011 Tambahan Offering
        Dim K As Integer
        Dim W As String
        Dim l As Integer
        Dim diskon As Integer
        
        Dim M_objrs As ADODB.Recordset
        Dim m_objrs_waktu As ADODB.Recordset
        Dim cmdsql As String
              
        
        'Cek dulu ada pembayaran apa ngga di tabel lunas
        cmdsql = "select * from tbllunas where custid='"
        cmdsql = cmdsql + Trim(lblcustid.Caption) + "'"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        
        'Ambil waktu sekarang
        cmdsql = "select now() as waktu "
        Set m_objrs_waktu = New ADODB.Recordset
        m_objrs_waktu.CursorLocation = adUseClient
        m_objrs_waktu.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        '@@ 08-06-2011, Jika lblpaydt=kosong on error goto salah
        On Error GoTo salah
        l = DateDiff("M", Format(lblPayDt.Value, "mm/dd/yyyy"), Format(CDate(m_objrs_waktu("waktu")), "mm/dd/yyyy"))
        
        '@@ 09-05-2011 Jika tidak ada nopay atau lpd > 4 bulan dari current date maka
        'tampilkan offering
        
        
        If M_objrs.RecordCount = 0 Or _
            l > 4 Then
            On Error GoTo salah
            K = DateDiff("M", Format(lblOpenDate.Value, "mm/dd/yyyy"), Format(lblBD.Value, "mm/dd/yyyy"))
            If K < 12 Then
                W = "Penawaran Diskon Maximal 60%"
                diskon = 60
            ElseIf K >= 12 And K <= 17 Then
                W = "Penawaran Diskon Maximal 50%"
                diskon = 50
            ElseIf K >= 18 And K <= 36 Then
                W = "Penawaran Diskon Maximal 40%"
                diskon = 40
            ElseIf K > 37 Then
                W = "Cicilan panjang " & " dan diskon 30%"
                diskon = 30
            End If
        
            'MsgBox "Pemandu Offering: " & w, vbOKOnly + vbInformation, "Offering Disc.Guide..."
            With FRMSCRIPT
                .LblTextGuide.Caption = "Pemandu Offering: " & W
                .TdbBalance.Value = lblAmount.Value
                .TdbMaxDisc.Value = diskon
                '.Show vbModal
            End With
        End If
        
        Set M_objrs = Nothing
        Set m_objrs_waktu = Nothing
        Exit Sub
salah:
    Set M_objrs = Nothing
    Set m_objrs_waktu = Nothing
End Sub

