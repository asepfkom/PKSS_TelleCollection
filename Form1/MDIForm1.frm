VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "TINS PKSS  vr.04/11/2019"
   ClientHeight    =   8910
   ClientLeft      =   195
   ClientTop       =   705
   ClientWidth     =   12810
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   NegotiateToolbars=   0   'False
   Picture         =   "MDIForm1.frx":08CA
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer100 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   1440
   End
   Begin Threed.SSPanel SSPanel4 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   74
      Top             =   0
      Width           =   12810
      _ExtentX        =   22595
      _ExtentY        =   635
      _Version        =   196610
      BackColor       =   -2147483635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Dashboard"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.Frame Frame3 
         BackColor       =   &H8000000B&
         Caption         =   "Dashboard"
         Height          =   4800
         Left            =   120
         TabIndex        =   75
         Top             =   480
         Width           =   18615
         Begin VB.Frame Frame4 
            Caption         =   "Search"
            Height          =   4455
            Left            =   14880
            TabIndex        =   77
            Top             =   240
            Width           =   3255
            Begin VB.CommandButton Command5 
               Caption         =   "Touch per Custid per Agent"
               Height          =   435
               Left            =   1680
               TabIndex        =   85
               Top             =   2880
               Width           =   1455
            End
            Begin VB.CommandButton Command7 
               Caption         =   "Export"
               Height          =   375
               Left            =   1680
               TabIndex        =   84
               Top             =   3840
               Width           =   1455
            End
            Begin VB.CommandButton Command6 
               Caption         =   "Search"
               Height          =   375
               Left            =   1680
               TabIndex        =   83
               Top             =   3360
               Width           =   1455
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               ItemData        =   "MDIForm1.frx":3468
               Left            =   120
               List            =   "MDIForm1.frx":3484
               TabIndex        =   82
               Top             =   1560
               Width           =   3015
            End
            Begin TDBDate6Ctl.TDBDate TDBDate3 
               Height          =   285
               Left            =   120
               TabIndex        =   79
               Top             =   720
               Width           =   1470
               _Version        =   65536
               _ExtentX        =   2593
               _ExtentY        =   503
               Calendar        =   "MDIForm1.frx":34C5
               Caption         =   "MDIForm1.frx":35DD
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "MDIForm1.frx":3649
               Keys            =   "MDIForm1.frx":3667
               Spin            =   "MDIForm1.frx":36C5
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   12648447
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
            Begin TDBDate6Ctl.TDBDate TDBDate4 
               Height          =   285
               Left            =   1680
               TabIndex        =   80
               Top             =   720
               Width           =   1470
               _Version        =   65536
               _ExtentX        =   2593
               _ExtentY        =   503
               Calendar        =   "MDIForm1.frx":36ED
               Caption         =   "MDIForm1.frx":3805
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "MDIForm1.frx":3871
               Keys            =   "MDIForm1.frx":388F
               Spin            =   "MDIForm1.frx":38ED
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   12648447
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
            Begin VB.Label Label8 
               Caption         =   "Client"
               Height          =   255
               Left            =   120
               TabIndex        =   81
               Top             =   1200
               Width           =   855
            End
            Begin VB.Label Label7 
               Caption         =   "Date"
               Height          =   255
               Left            =   120
               TabIndex        =   78
               Top             =   360
               Width           =   615
            End
         End
         Begin MSComctlLib.ListView LvAgent 
            Height          =   4455
            Left            =   120
            TabIndex        =   76
            Top             =   240
            Width           =   14595
            _ExtentX        =   25744
            _ExtentY        =   7858
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin MSComDlg.CommonDialog CD_save 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Timer Timer6 
      Interval        =   8000
      Left            =   5520
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   6000
      Left            =   5040
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6480
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer TimerCTI 
      Interval        =   300
      Left            =   6000
      Top             =   0
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   15
      Left            =   0
      TabIndex        =   2
      Top             =   8145
      Visible         =   0   'False
      Width           =   12810
      _ExtentX        =   22595
      _ExtentY        =   26
      _Version        =   196610
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox TxtLamaFollowup 
         Height          =   285
         Left            =   1050
         TabIndex        =   17
         Top             =   945
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.TextBox TxtJamSelesaiTelp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   7605
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "23:59:59"
         Top             =   120
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.TextBox TxtJamMulaiTelp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5715
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "00:00:00"
         Top             =   120
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.TextBox TxtModemAcod 
         Height          =   285
         Left            =   420
         TabIndex        =   10
         Text            =   "Text8"
         Top             =   1275
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.TextBox TxtAuthPrefix 
         Height          =   285
         Left            =   1770
         TabIndex        =   9
         Text            =   "Text8"
         Top             =   2715
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.TextBox TxtAuth 
         Height          =   285
         Left            =   4110
         TabIndex        =   8
         Top             =   90
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2700
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   75
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox TxtCommPort 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2250
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   75
         Width           =   390
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   12435
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   75
         Width           =   2685
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Serial Communication"
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
         Height          =   225
         Left            =   285
         TabIndex        =   5
         Top             =   90
         Width           =   1920
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   1  'Align Top
      Height          =   15
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   12810
      _ExtentX        =   22595
      _ExtentY        =   26
      _Version        =   196610
      BevelWidth      =   2
      BorderWidth     =   2
      BevelInner      =   1
      AutoSize        =   3
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox LblTargetMaster 
         Height          =   285
         Left            =   9165
         TabIndex        =   15
         Top             =   495
         Visible         =   0   'False
         Width           =   2175
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   3
         Left            =   2430
         TabIndex        =   0
         ToolTipText     =   "Pesan"
         Top             =   120
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   1085
         _Version        =   196610
         Font3D          =   1
         MousePointer    =   16
         ForeColor       =   12582912
         BackColor       =   12632256
         PictureFrames   =   1
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MDIForm1.frx":3915
         Caption         =   "Msg"
         AutoSize        =   1
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   6
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   9
         Left            =   3240
         TabIndex        =   7
         ToolTipText     =   "Program Report"
         Top             =   135
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   1085
         _Version        =   196610
         Font3D          =   1
         MousePointer    =   16
         ForeColor       =   12582912
         BackColor       =   12632256
         PictureFrames   =   1
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MDIForm1.frx":3D67
         Caption         =   "Report"
         AutoSize        =   1
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   6
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   1000
         Index           =   1
         Left            =   10
         TabIndex        =   11
         ToolTipText     =   "Mgm Program Report"
         Top             =   500
         Visible         =   0   'False
         Width           =   1000
         _ExtentX        =   1746
         _ExtentY        =   1773
         _Version        =   196610
         Font3D          =   1
         MousePointer    =   16
         ForeColor       =   12582912
         BackColor       =   12632256
         PictureFrames   =   1
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MDIForm1.frx":41B9
         Caption         =   "Daily Report"
         AutoSize        =   1
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   6
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   2
         Left            =   5760
         TabIndex        =   16
         ToolTipText     =   "Mgm Program Report"
         Top             =   105
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1085
         _Version        =   196610
         Font3D          =   1
         MousePointer    =   16
         ForeColor       =   12582912
         BackColor       =   12632256
         PictureFrames   =   1
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MDIForm1.frx":460B
         Caption         =   "Preembos Report"
         AutoSize        =   1
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   6
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   4
         Left            =   560
         TabIndex        =   19
         ToolTipText     =   "Mgm Program Report"
         Top             =   120
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1085
         _Version        =   196610
         Font3D          =   1
         MousePointer    =   16
         ForeColor       =   12582912
         BackColor       =   12632256
         PictureFrames   =   1
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MDIForm1.frx":4A5D
         Caption         =   "Access Data"
         AutoSize        =   1
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   6
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   5
         Left            =   7440
         TabIndex        =   20
         ToolTipText     =   "Tambah Data Baru"
         Top             =   120
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1085
         _Version        =   196610
         Font3D          =   1
         MousePointer    =   16
         ForeColor       =   12582912
         BackColor       =   12632256
         PictureFrames   =   1
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MDIForm1.frx":4EAF
         Caption         =   "&Searching Visit"
         AutoSize        =   1
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   6
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   6
         Left            =   1200
         TabIndex        =   21
         ToolTipText     =   "Pesan"
         Top             =   120
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   1085
         _Version        =   196610
         Font3D          =   1
         MousePointer    =   16
         ForeColor       =   12582912
         BackColor       =   12632256
         PictureFrames   =   1
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MDIForm1.frx":5301
         Caption         =   "New Search"
         AutoSize        =   1
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   6
         BevelWidth      =   1
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   7
         Left            =   5760
         TabIndex        =   22
         ToolTipText     =   "Tambah Data Baru"
         Top             =   135
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1085
         _Version        =   196610
         Font3D          =   1
         MousePointer    =   16
         ForeColor       =   12582912
         BackColor       =   12632256
         PictureFrames   =   1
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MDIForm1.frx":6153
         Caption         =   "Un Lock Data"
         AutoSize        =   1
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   6
         BevelWidth      =   1
      End
      Begin VB.Label Label4 
         Caption         =   "JANGAN TUTUP PROGRAM INI, UNTUK OTOMATIC BP/POP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   0
         TabIndex        =   23
         Top             =   120
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.Label LblTarget 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   2100
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   5790
      End
   End
   Begin MSCommLib.MSComm MsComLogin 
      Left            =   7440
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSWinsockLib.Winsock WskCTI 
      Left            =   6960
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   18000
   End
   Begin Threed.SSPanel SSPanel3 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   18
      Top             =   8160
      Width           =   12810
      _ExtentX        =   22595
      _ExtentY        =   1323
      _Version        =   196610
      PictureMaskColor=   16448250
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   8040
         TabIndex        =   73
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txt_unique_id 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9885
         TabIndex        =   69
         Top             =   630
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.TextBox txtChannel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   10125
         TabIndex        =   68
         Top             =   630
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Caption         =   "360"
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   9960
         TabIndex        =   54
         Top             =   0
         Width           =   8535
         Begin VB.TextBox TxtStatus 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   4320
            Locked          =   -1  'True
            TabIndex        =   59
            Text            =   $"MDIForm1.frx":65A5
            Top             =   380
            Width           =   2775
         End
         Begin VB.TextBox txtnama 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00202020&
            Height          =   330
            Left            =   960
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   58
            TabStop         =   0   'False
            Text            =   "USER NAME"
            Top             =   360
            Width           =   2295
         End
         Begin VB.TextBox txtlevel 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00202020&
            Height          =   330
            Left            =   4320
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   56
            TabStop         =   0   'False
            Text            =   "USER LEVEL"
            Top             =   60
            Width           =   2655
         End
         Begin VB.TextBox txtusername 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00202020&
            Height          =   330
            Left            =   960
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   55
            TabStop         =   0   'False
            Text            =   "USER ID"
            Top             =   50
            Width           =   2295
         End
         Begin TDBDate6Ctl.TDBDate TDBDate1 
            Height          =   315
            Left            =   7320
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   240
            Width           =   1230
            _Version        =   65536
            _ExtentX        =   2170
            _ExtentY        =   556
            Calendar        =   "MDIForm1.frx":65B6
            Caption         =   "MDIForm1.frx":66CE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "MDIForm1.frx":673A
            Keys            =   "MDIForm1.frx":6758
            Spin            =   "MDIForm1.frx":67B6
            AlignHorizontal =   0
            AlignVertical   =   2
            Appearance      =   0
            BackColor       =   -2147483633
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
            ForeColor       =   2105376
            Format          =   "dd/mm/yyyy"
            HighlightText   =   1
            IMEMode         =   3
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxDate         =   2958465
            MinDate         =   -657434
            MousePointer    =   1
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
            Value           =   37475
            CenturyMode     =   0
         End
         Begin VB.Label Label5 
            Caption         =   "STAT    :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   3600
            TabIndex        =   64
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "LEVEL   :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   3600
            TabIndex        =   63
            Top             =   50
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "NAME     :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   62
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "USER ID  :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   61
            Top             =   50
            Width           =   855
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000016&
            Index           =   2
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   720
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000016&
            Index           =   1
            X1              =   7080
            X2              =   7080
            Y1              =   0
            Y2              =   720
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000016&
            Index           =   0
            X1              =   3480
            X2              =   3480
            Y1              =   0
            Y2              =   720
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Sampah"
         Height          =   1095
         Left            =   14880
         TabIndex        =   28
         Top             =   720
         Visible         =   0   'False
         Width           =   2535
         Begin VB.CommandButton cmd_break 
            BackColor       =   &H008080FF&
            Caption         =   "Break Time !!"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   750
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox TxtOnline 
            Height          =   915
            Left            =   720
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   45
            Top             =   480
            Visible         =   0   'False
            Width           =   4035
         End
         Begin VB.CommandButton Command2 
            Caption         =   "<< &Show Media"
            Height          =   375
            Left            =   3240
            TabIndex        =   44
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox TxtIPIcentra 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   3300
            TabIndex        =   43
            Top             =   1200
            Width           =   1335
         End
         Begin VB.TextBox txtspvcode 
            Height          =   285
            Left            =   4995
            TabIndex        =   42
            Top             =   990
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.TextBox TxtWaktuRefresh 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5100
            TabIndex        =   41
            Text            =   "00:00:30"
            Top             =   795
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton CmdInbox 
            Caption         =   "Command1"
            Height          =   315
            Left            =   360
            TabIndex        =   40
            Top             =   915
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox RichTextBox1 
            Height          =   285
            Left            =   2010
            TabIndex        =   39
            Text            =   "Text4"
            Top             =   1035
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.CommandButton apaja 
            Caption         =   "Command3"
            Enabled         =   0   'False
            Height          =   195
            Left            =   3480
            TabIndex        =   38
            Top             =   675
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   5760
            TabIndex        =   37
            Text            =   "Text4"
            Top             =   915
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
            Height          =   375
            Left            =   5880
            TabIndex        =   36
            Top             =   2355
            Width           =   375
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Command4"
            Enabled         =   0   'False
            Height          =   375
            Left            =   4320
            TabIndex        =   35
            Top             =   555
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtTeam 
            Height          =   285
            Left            =   2715
            TabIndex        =   34
            Top             =   480
            Visible         =   0   'False
            Width           =   615
         End
         Begin Threed.SSFrame SSFrame1 
            Height          =   3450
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   2640
            _ExtentX        =   4657
            _ExtentY        =   6085
            _Version        =   196610
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin Threed.SSCommand SSCommand2 
               Height          =   375
               Left            =   120
               TabIndex        =   31
               Top             =   2880
               Visible         =   0   'False
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   661
               _Version        =   196610
               Caption         =   "&REFRESH"
            End
            Begin MSComctlLib.ListView LstGrade 
               Height          =   3345
               Left            =   45
               TabIndex        =   32
               Top             =   60
               Width           =   2580
               _ExtentX        =   4551
               _ExtentY        =   5900
               View            =   3
               Sorted          =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   16761087
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
            Begin VB.Label Label3 
               Caption         =   "Label3"
               Height          =   375
               Left            =   120
               TabIndex        =   33
               Top             =   3000
               Width           =   2295
            End
         End
         Begin MSComctlLib.ListView LstInformation 
            Height          =   135
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   238
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.Label LblBersihkan 
            BackColor       =   &H80000008&
            Height          =   195
            Left            =   240
            TabIndex        =   50
            Top             =   360
            Width           =   195
         End
         Begin VB.Shape ShapeTanda 
            FillStyle       =   0  'Solid
            Height          =   495
            Left            =   360
            Shape           =   3  'Circle
            Top             =   720
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label_OL_count 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2880
            TabIndex        =   49
            Top             =   840
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label LblWaktu 
            BackStyle       =   0  'Transparent
            Caption         =   "Label Waktu"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Visible         =   0   'False
            Width           =   2700
         End
         Begin VB.Label lbl_timer_activity 
            BackStyle       =   0  'Transparent
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4740
            TabIndex        =   46
            Top             =   600
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Shape ShapeReq 
            FillStyle       =   0  'Solid
            Height          =   495
            Left            =   600
            Shape           =   3  'Circle
            Top             =   660
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label2 
            Caption         =   "-"
            Height          =   465
            Left            =   120
            TabIndex        =   29
            Top             =   600
            Visible         =   0   'False
            Width           =   2175
         End
      End
      Begin VB.Label lblapp 
         Caption         =   "NEW APPROVAL**"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   5775
         TabIndex        =   72
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblsms_unread 
         Caption         =   "NEW MESSAGE"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   5790
         TabIndex        =   71
         Top             =   60
         Visible         =   0   'False
         Width           =   1815
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   12
         Left            =   3360
         TabIndex        =   70
         ToolTipText     =   "Follow Up"
         Top             =   60
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         _Version        =   196610
         ForeColor       =   4210752
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "SMS"
         ButtonStyle     =   2
         PictureAlignment=   1
         BevelWidth      =   0
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   8
         Left            =   2265
         TabIndex        =   67
         ToolTipText     =   "Follow Up"
         Top             =   60
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         _Version        =   196610
         ForeColor       =   4210752
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Reminder"
         ButtonStyle     =   2
         PictureAlignment=   1
         BevelWidth      =   0
      End
      Begin VB.Label Label6 
         Caption         =   "0"
         Height          =   255
         Left            =   6360
         TabIndex        =   66
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   65
         ToolTipText     =   "Follow Up"
         Top             =   60
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         _Version        =   196610
         ForeColor       =   4210752
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Follow Up"
         ButtonStyle     =   2
         PictureAlignment=   1
         BevelWidth      =   0
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   17
         Left            =   4470
         TabIndex        =   60
         Top             =   60
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         _Version        =   196610
         ForeColor       =   4210752
         PictureFrames   =   1
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "MDIForm1.frx":67DE
         Caption         =   "&"
         ButtonStyle     =   2
         BevelWidth      =   0
      End
      Begin Threed.SSCommand SSCommand1 
         Default         =   -1  'True
         Height          =   255
         Index           =   11
         Left            =   6045
         TabIndex        =   53
         Top             =   750
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   196610
         ForeColor       =   4210752
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Report"
         ButtonStyle     =   2
         BevelWidth      =   0
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   615
         Index           =   10
         Left            =   1200
         TabIndex        =   52
         ToolTipText     =   "Broadcast"
         Top             =   60
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1085
         _Version        =   196610
         ForeColor       =   4210752
         Windowless      =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "BroadCast"
         ButtonStyle     =   2
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin VB.Label LblJmlSmsBaru 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   195
         Left            =   19080
         TabIndex        =   27
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label9 
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "aaaa"
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
         Left            =   4980
         TabIndex        =   26
         Top             =   480
         Visible         =   0   'False
         Width           =   4515
      End
      Begin VB.Label Label10 
         BackColor       =   &H008080FF&
         BackStyle       =   0  'Transparent
         Height          =   285
         Left            =   12885
         TabIndex        =   25
         Top             =   150
         Width           =   4605
      End
      Begin VB.Label Lbltargetspv 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   8040
         TabIndex        =   24
         Top             =   570
         Width           =   11265
      End
   End
   Begin VB.Menu mnbar 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu MnFile 
         Caption         =   "&Set Password"
         Index           =   3
      End
      Begin VB.Menu MnFile 
         Caption         =   "-"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu MnFile 
         Caption         =   "&Change Password"
         Index           =   5
      End
      Begin VB.Menu MnFile 
         Caption         =   "-"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu MnFile 
         Caption         =   "E&xit"
         Index           =   7
      End
   End
   Begin VB.Menu mnbar 
      Caption         =   "&Master"
      Index           =   1
      Begin VB.Menu mnoffice 
         Caption         =   "&Data Officer"
         Begin VB.Menu mnagent 
            Caption         =   "Agent"
         End
         Begin VB.Menu mntl 
            Caption         =   "SPV"
         End
         Begin VB.Menu mnmgr 
            Caption         =   "Manager"
            Visible         =   0   'False
         End
         Begin VB.Menu mnadmin 
            Caption         =   "Admin"
         End
         Begin VB.Menu nmgdata 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu nmrestoredeleteacc 
            Caption         =   "Restore and delete account"
            Visible         =   0   'False
         End
         Begin VB.Menu nmuploadtempdata 
            Caption         =   "Upload Temporary Data"
            Visible         =   0   'False
         End
         Begin VB.Menu nmg31 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu nmbackup 
            Caption         =   "Backup data tabel backup"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnsubmarkup 
         Caption         =   "Upload For Lock Account"
         Visible         =   0   'False
      End
      Begin VB.Menu MnTarget 
         Caption         =   "&Target"
         Visible         =   0   'False
      End
      Begin VB.Menu MnDsr 
         Caption         =   "Daily Submission Report"
         Visible         =   0   'False
      End
      Begin VB.Menu MnWpi 
         Caption         =   "&Weekly Performance Indicator"
         Visible         =   0   'False
      End
      Begin VB.Menu MnTglSeThn 
         Caption         =   "Set Tang&gal Setahun"
         Visible         =   0   'False
      End
      Begin VB.Menu test1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MnAccDecSub 
         Caption         =   "Accept Decline Submission"
         Visible         =   0   'False
      End
      Begin VB.Menu mnrole 
         Caption         =   "Menu Role"
      End
      Begin VB.Menu mnNact 
         Caption         =   "&Status Call"
      End
      Begin VB.Menu MnCCode 
         Caption         =   "Com&plaint Code"
         Visible         =   0   'False
      End
      Begin VB.Menu mndata 
         Caption         =   "&Data Quality"
         Visible         =   0   'False
      End
      Begin VB.Menu mnreason 
         Caption         =   "&Uncontacted Status Call "
         Visible         =   0   'False
      End
      Begin VB.Menu mncontacted 
         Caption         =   "&Contacted Status Call"
         Visible         =   0   'False
      End
      Begin VB.Menu mndata2 
         Caption         =   "&Campaign"
         Visible         =   0   'False
      End
      Begin VB.Menu mnProduct 
         Caption         =   "&Product"
         Visible         =   0   'False
      End
      Begin VB.Menu mnPr 
         Caption         =   "Set Product &Knowledge"
         Visible         =   0   'False
      End
      Begin VB.Menu MnBb 
         Caption         =   "Bulletin &Board"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuProdInfo 
         Caption         =   "Product Info"
         Visible         =   0   'False
      End
      Begin VB.Menu sbgaris 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnblack 
         Caption         =   "Black List No telpon"
         Visible         =   0   'False
      End
      Begin VB.Menu subupdate 
         Caption         =   "Update status"
         Visible         =   0   'False
      End
      Begin VB.Menu MnBlokData 
         Caption         =   "Blok Data"
         Visible         =   0   'False
         Begin VB.Menu mnblokspv 
            Caption         =   "&Schedule Blok Data"
         End
      End
      Begin VB.Menu nmSchLocktl 
         Caption         =   "Schedule Blok Data"
         Visible         =   0   'False
      End
      Begin VB.Menu setspv 
         Caption         =   "Set Target From SPV"
         Visible         =   0   'False
      End
      Begin VB.Menu mnsubahstsacc 
         Caption         =   "Ubah Status Account"
         Visible         =   0   'False
      End
      Begin VB.Menu nmformceksts 
         Caption         =   "Cek Account Status Progress"
         Visible         =   0   'False
      End
      Begin VB.Menu nmlistreqform 
         Caption         =   "Viewer List Request Form"
         Visible         =   0   'False
      End
      Begin VB.Menu nmlstreqnumber 
         Caption         =   "Approval Request Additional Phone"
         Visible         =   0   'False
      End
      Begin VB.Menu nmmenuformlistconfidence 
         Caption         =   "Form List Confidence"
         Visible         =   0   'False
      End
      Begin VB.Menu mnbalance 
         Caption         =   "Payment Pattern"
         Visible         =   0   'False
      End
      Begin VB.Menu separ 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnrptsms 
         Caption         =   "Report SMS NEW"
         Visible         =   0   'False
      End
      Begin VB.Menu VSMS 
         Caption         =   "Verify SMS"
         Visible         =   0   'False
      End
      Begin VB.Menu smsblast 
         Caption         =   "Blast SMS Text"
         Visible         =   0   'False
      End
      Begin VB.Menu nmlistsmsscript 
         Caption         =   "List sms script"
         Visible         =   0   'False
      End
      Begin VB.Menu nmapprovreject1 
         Caption         =   "Approved and rejected"
         Visible         =   0   'False
      End
      Begin VB.Menu nmblastsmsexcel 
         Caption         =   "Send SMS Blast Via Excel"
         Visible         =   0   'False
      End
      Begin VB.Menu nmg10 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu MNUOFFER 
         Caption         =   "Form Offering"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuploadskip 
         Caption         =   "Upload Skip Tracer"
         Visible         =   0   'False
      End
      Begin VB.Menu nmg 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu nmReportCall 
         Caption         =   "Report Call"
         Visible         =   0   'False
         Begin VB.Menu nmRptCallServer4 
            Caption         =   "Report Call Server 4"
         End
         Begin VB.Menu nmReportCallServer5 
            Caption         =   "Report Call Server 5"
         End
      End
      Begin VB.Menu nmg3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu nmlistsendcpa 
         Caption         =   "List Send CPA"
         Visible         =   0   'False
      End
      Begin VB.Menu nmuploadcpaptp 
         Caption         =   "Upload CPA dan PTP"
         Visible         =   0   'False
      End
      Begin VB.Menu nmListUnValidNumber 
         Caption         =   "List Unvalid Number"
         Visible         =   0   'False
      End
      Begin VB.Menu nmAksesLayanaTelkom 
         Caption         =   "Akses Layanan Telkom"
         Visible         =   0   'False
      End
      Begin VB.Menu nmlistreqptp 
         Caption         =   "&List Request PTP"
      End
      Begin VB.Menu nmresetpass 
         Caption         =   "&Reset Password"
         Visible         =   0   'False
      End
      Begin VB.Menu nmReportProblemHeadset 
         Caption         =   "List Report Problem Headset"
         Visible         =   0   'False
      End
      Begin VB.Menu nmListReportProblemTelepon 
         Caption         =   "List Report Problem Telepon"
         Visible         =   0   'False
      End
      Begin VB.Menu nmblokaplikasitins 
         Caption         =   "Blok Aplikasi TINS"
         Visible         =   0   'False
      End
      Begin VB.Menu nmManageDistribusiAccount 
         Caption         =   "Manage Distribusi Account"
         Visible         =   0   'False
      End
      Begin VB.Menu mnListAccountLunas 
         Caption         =   "List Account Lunas"
         Visible         =   0   'False
      End
      Begin VB.Menu mn_list_complaint 
         Caption         =   "List Data Complaint"
         Visible         =   0   'False
      End
      Begin VB.Menu mn_list_sid 
         Caption         =   "List SID"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnbar 
      Caption         =   "&Distribusi Data"
      Index           =   2
      Visible         =   0   'False
      Begin VB.Menu mnup 
         Caption         =   "Distribusi &Data Upload"
      End
      Begin VB.Menu mnhslupload 
         Caption         =   "&Hasil Upload"
      End
      Begin VB.Menu mndist 
         Caption         =   "Hasil &Distribusi"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnbar 
      Caption         =   "&Pesan"
      Index           =   3
      Visible         =   0   'False
      Begin VB.Menu mnsend 
         Caption         =   "&Kirim"
      End
      Begin VB.Menu mnbaca 
         Caption         =   "&Baca"
      End
   End
   Begin VB.Menu mnbar 
      Caption         =   "&About"
      Index           =   4
      Visible         =   0   'False
      Begin VB.Menu mnProductKnowLedge 
         Caption         =   "Product &Knowledge"
         Visible         =   0   'False
      End
      Begin VB.Menu mnabout 
         Caption         =   "&Tentang Kami"
      End
   End
   Begin VB.Menu mnbar 
      Caption         =   "Distribusi S&TP"
      Index           =   5
      Visible         =   0   'False
      Begin VB.Menu msdisstp 
         Caption         =   "Distribusi Data STP"
      End
   End
   Begin VB.Menu mnbar 
      Caption         =   "Distribusi &SPV"
      Index           =   6
      Visible         =   0   'False
      Begin VB.Menu UPTMGM 
         Caption         =   "Distribusi Data MGM"
      End
      Begin VB.Menu UPTSTP 
         Caption         =   "Distribusi STP"
      End
   End
   Begin VB.Menu mnbar 
      Caption         =   "Distr&ibusi Data Tarik"
      Index           =   7
      Visible         =   0   'False
      Begin VB.Menu TarikMGM 
         Caption         =   "Data CH (Player)"
      End
      Begin VB.Menu TarikLeads 
         Caption         =   "Data Leads"
      End
      Begin VB.Menu TarikStp 
         Caption         =   "Data STP"
      End
   End
   Begin VB.Menu mnbar 
      Caption         =   "&Duplikasi"
      Index           =   8
      Visible         =   0   'False
      Begin VB.Menu MNDUPLIKASI 
         Caption         =   "Duplikasi Leads"
      End
      Begin VB.Menu MNDUPLIKASICH 
         Caption         =   "Duplikasi CH"
      End
   End
   Begin VB.Menu mnbar 
      Caption         =   "Pending &Duplikasi"
      Index           =   9
      Visible         =   0   'False
      Begin VB.Menu MnPendingCh 
         Caption         =   "Pending CH"
      End
      Begin VB.Menu MnPendingLeads 
         Caption         =   "Pending Leads"
      End
   End
   Begin VB.Menu mnbar 
      Caption         =   "Kurir"
      Index           =   10
      Visible         =   0   'False
      Begin VB.Menu mnkrmaplikasi 
         Caption         =   "Kirim Aplikasi"
      End
   End
   Begin VB.Menu mnbar 
      Caption         =   "Report"
      Index           =   11
      Visible         =   0   'False
      Begin VB.Menu MnReportTracking 
         Caption         =   "Report Tracking"
      End
      Begin VB.Menu MnVisit 
         Caption         =   "Visit Status"
      End
      Begin VB.Menu nmReportSms 
         Caption         =   "Report SMS"
      End
   End
   Begin VB.Menu mnbar 
      Caption         =   "Report"
      Index           =   12
      Begin VB.Menu mnrboard 
         Caption         =   "Dashboard Agent"
         Visible         =   0   'False
      End
      Begin VB.Menu mnrdistribut 
         Caption         =   "Distribute Report"
      End
      Begin VB.Menu mnrmis 
         Caption         =   "Report MIS"
         Visible         =   0   'False
      End
      Begin VB.Menu mnract 
         Caption         =   "Report Call Activiry"
         Visible         =   0   'False
      End
      Begin VB.Menu mnrdetail 
         Caption         =   "Report Call Detail"
         Visible         =   0   'False
      End
      Begin VB.Menu mnlast 
         Caption         =   "Report Last Statuscall"
         Visible         =   0   'False
      End
      Begin VB.Menu mnrresult 
         Caption         =   "Result Report"
      End
      Begin VB.Menu mnroutrpt 
         Caption         =   "OutgoingCall Report"
         Visible         =   0   'False
      End
      Begin VB.Menu mnrsummery 
         Caption         =   "Report Summery Status Call"
         Visible         =   0   'False
      End
      Begin VB.Menu mnnsms 
         Caption         =   "Report SMS"
         Visible         =   0   'False
      End
      Begin VB.Menu mn_monhly_bp 
         Caption         =   "Monthly BP (Broken Promise)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnmonthcpa 
         Caption         =   "Monthly CPA"
         Visible         =   0   'False
      End
      Begin VB.Menu mnptppayment 
         Caption         =   "Monthly PTP - Payment"
         Visible         =   0   'False
      End
      Begin VB.Menu nmconfidenceanalisysagent 
         Caption         =   "Confidence Analisys Agent"
         Visible         =   0   'False
      End
      Begin VB.Menu mn_confidence_list 
         Caption         =   "Confidence List"
         Visible         =   0   'False
      End
      Begin VB.Menu mn_performance 
         Caption         =   "DeskColl Performance"
         Visible         =   0   'False
      End
      Begin VB.Menu mn_deskcoll_perform2 
         Caption         =   "Average Performance"
         Visible         =   0   'False
      End
      Begin VB.Menu rrld 
         Caption         =   "Report Remarks Last Day"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu nmenu 
      Caption         =   "Menu"
      Visible         =   0   'False
   End
   Begin VB.Menu mntools 
      Caption         =   "&Tools"
      Begin VB.Menu tigaA 
         Caption         =   "Approval Add Number"
      End
      Begin VB.Menu nmapprovreject 
         Caption         =   "Approved and rejected sms"
         Visible         =   0   'False
      End
      Begin VB.Menu mntquery 
         Caption         =   "Query Analizer"
         Visible         =   0   'False
      End
      Begin VB.Menu mndistribut 
         Caption         =   "&Distribusi Data"
      End
      Begin VB.Menu mndpc 
         Caption         =   "&Distribusi Per Custid"
      End
      Begin VB.Menu mnrecycle 
         Caption         =   "&Recycle Data"
      End
      Begin VB.Menu nmupload 
         Caption         =   "&Upload Data"
         Begin VB.Menu nmuploadcustomer 
            Caption         =   "Upload Data Customer"
         End
         Begin VB.Menu nmuploadpayment 
            Caption         =   "Upload Payment"
            Visible         =   0   'False
         End
         Begin VB.Menu cuda 
            Caption         =   "Upload Data All"
            Visible         =   0   'False
         End
         Begin VB.Menu nmuploadaddress 
            Caption         =   "Upload Address"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuuploadcpa 
            Caption         =   "Upload CPA"
            Visible         =   0   'False
         End
         Begin VB.Menu nmswapdata 
            Caption         =   "Swap data"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnPO 
         Caption         =   "&PullOut Data"
         Visible         =   0   'False
      End
      Begin VB.Menu mnpd 
         Caption         =   "Parkir Data"
      End
      Begin VB.Menu mnaddclient 
         Caption         =   "&Add Client"
         Visible         =   0   'False
      End
      Begin VB.Menu mntd 
         Caption         =   "&Tarik Data"
      End
      Begin VB.Menu mnqcd 
         Caption         =   "Query Cek Data"
      End
      Begin VB.Menu nmgupload 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu list_phone_review 
         Caption         =   "&List Phone Review"
         Visible         =   0   'False
      End
      Begin VB.Menu mn_aoc 
         Caption         =   "&AOC"
         Visible         =   0   'False
      End
      Begin VB.Menu transfer_data 
         Caption         =   "Transfer Data"
         Visible         =   0   'False
      End
      Begin VB.Menu add_special_history 
         Caption         =   "Add Special History"
         Visible         =   0   'False
      End
      Begin VB.Menu upload_fresh_wo 
         Caption         =   "&Upload Data Fresh WO"
         Visible         =   0   'False
      End
      Begin VB.Menu mn_report_temp 
         Caption         =   "Report Temp Agent"
         Visible         =   0   'False
      End
      Begin VB.Menu mndrm 
         Caption         =   "Delete & Restore Marks"
         Index           =   55
         Visible         =   0   'False
      End
      Begin VB.Menu mn_performance_reguler 
         Caption         =   "DeskColl Performance Reguler"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCallmonitor 
         Caption         =   "Call Monitor"
         Visible         =   0   'False
      End
      Begin VB.Menu mn_copyfile 
         Caption         =   "Copy File CPA dan Dokumen Pendukung"
         Visible         =   0   'False
      End
      Begin VB.Menu mnLDS 
         Caption         =   "Lock Data System"
         Visible         =   0   'False
      End
      Begin VB.Menu mn_option_hide 
         Caption         =   "Filter Hide System"
         Visible         =   0   'False
      End
      Begin VB.Menu mnMPD 
         Caption         =   "Monitoring Perpindahan Data"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnMenuRole 
         Caption         =   "Menu Role"
      End
   End
   Begin VB.Menu mn_update_db 
      Caption         =   "Update DB"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim m_TelpTglAwal As String
Dim m_TelpTglAkhir As String
Dim m_TelpAgent As String
Dim f_InsertTelp As Boolean
Dim M_LOGINRS As ADODB.Recordset
Public TXT As Integer
Public m_TelpUserId As String
Public m_TelpNoTelp As String
Public m_faxName As String
Public B_FAX As Boolean
Public m_targetview As Boolean
Public ParameterCTI  As String
Public m_data As New CLS_FRMCUST_CC_MGM
Dim COUNTER As Integer
Public Kalimat1 As String
Public PANJANG As Double
Dim satu As String
Dim dua As String
Dim tiga As String
Dim empat As String
Dim KelapKelip As Integer
Public KuotaSms As Integer
Dim STATUS As String
Dim z As Integer
'@@ 13-04-2011 Tambahan buat jumlah maksimal request koneksi
Dim JmlKoneksiReq As Integer
Dim MaxKoneksiReq As Integer

'@@=== 15-12-2010 buat lock data , (di running cuma di tl)
Dim TotalTenthDetik, TotalDetik, TenthDetik, Detik, Menit, JAM As Integer
Dim jam1 As String
'@@=== 15-12-2010 buat lock data , (di running cuma di tl)


'===============@@ 6-12-2010 buat tooltip kalo ada sms yang masuk ===========================
Private Declare Function CreateWindowEx Lib "user32" _
Alias "CreateWindowExA" (ByVal dwExStyle As Long, _
ByVal lpClassName As String, ByVal lpWindowName As _
String, ByVal dwStyle As Long, ByVal X As Long, _
ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight _
As Long, ByVal hWndParent As Long, ByVal hMenu _
As Long, ByVal hInstance As Long, lpParam As Any) _
As Long

Private Declare Function SetWindowPos Lib "user32" _
(ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function SendMessage Lib "user32" _
Alias "SendMessageA" (ByVal hwnd As Long, _
ByVal wMsg As Long, ByVal wParam As Long, _
lParam As Any) As Long

Private Declare Function GetClientRect Lib "user32" _
(ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function DestroyWindow Lib "user32" _
(ByVal hwnd As Long) As Long

'UDT (User Defined Type) RECT.
'Digunakan untuk pengaturan batas dari jendela tooltip.
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

'UDT TOOLINFO.
'Digunakan untuk menentukan semua tanda yang diperlukan
'untuk membuat jendela tooltip.
Private Type TOOLINFO
  cbSize As Long
  uFlags As Long
  hwnd As Long
  uid As Long
  RECT As RECT
  hinst As Long
  lpszText As String
  lParam As Long
End Type

'Sebuah konstanta yang digunakan untuk menghubungkan
'ke fungsi API yang bernama: CreateWindowEx.
'Hal ini untuk menandakan nilai default yang digunakan.
Private Const CW_USEDEFAULT = &H80000000

'Konstanta untuk fungsi API bernama: SetWindowPosition.
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOSIZE = &H1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1
Private Const HWND_BOTTOM = 1

'Konstanta untuk menentukan gaya dari jendela tooltip.
Private Const WS_POPUP = &H80000000
Private Const WS_EX_TOPMOST = &H8&

'Konstanta yang digunakan dengan fungsi API SendMessage
'untuk mendefinisikan pesan private.
Private Const WM_USER = &H400

'Messages yang digunakan untuk menentukan durasi waktu 'dari tooltips. Tidak digunakan di sini.
Private Const TTDT_AUTOMATIC = 0
Private Const TTDT_AUTOPOP = 2
Private Const TTDT_INITIAL = 3
Private Const TTDT_RESHOW = 1

'Semua "penanda" untuk jendela tooltip.
Private Const TTF_ABSOLUTE = &H80
Private Const TTF_CENTERTIP = &H2
Private Const TTF_DI_SETITEM = &H8000
Private Const TTF_IDISHWND = &H1
Private Const TTF_RTLREADING = &H4
Private Const TTF_SUBCLASS = &H10
Private Const TTF_TRACK = &H20
Private Const TTF_TRANSPARENT = &H100

'Semua pesan yang tersedia untuk tooltip Windows.
Private Const TTM_ACTIVATE = (WM_USER + 1)
Private Const TTM_ADDTOOLA = (WM_USER + 4)
Private Const TTM_ADDTOOLW = (WM_USER + 50)
Private Const TTM_ADJUSTRECT = (WM_USER + 31)
Private Const TTM_DELTOOLA = (WM_USER + 5)
Private Const TTM_DELTOOLW = (WM_USER + 51)
Private Const TTM_ENUMTOOLSA = (WM_USER + 14)
Private Const TTM_ENUMTOOLSW = (WM_USER + 58)
Private Const TTM_GETBUBBLESIZE = (WM_USER + 30)
Private Const TTM_GETCURRENTTOOLA = (WM_USER + 15)
Private Const TTM_GETCURRENTTOOLW = (WM_USER + 59)
Private Const TTM_GETDELAYTIME = (WM_USER + 21)
Private Const TTM_GETMARGIN = (WM_USER + 27)
Private Const TTM_GETMAXTIPWIDTH = (WM_USER + 25)
Private Const TTM_GETTEXTA = (WM_USER + 11)
Private Const TTM_GETTEXTW = (WM_USER + 56)
Private Const TTM_GETTIPBKCOLOR = (WM_USER + 22)
Private Const TTM_GETTIPTEXTCOLOR = (WM_USER + 23)
Private Const TTM_GETTOOLCOUNT = (WM_USER + 13)
Private Const TTM_GETTOOLINFOA = (WM_USER + 8)
Private Const TTM_GETTOOLINFOW = (WM_USER + 53)
Private Const TTM_HITTESTA = (WM_USER + 10)
Private Const TTM_HITTESTW = (WM_USER + 55)
Private Const TTM_NEWTOOLRECTA = (WM_USER + 6)
Private Const TTM_NEWTOOLRECTW = (WM_USER + 52)
Private Const TTM_POP = (WM_USER + 28)
Private Const TTM_RELAYEVENT = (WM_USER + 7)
Private Const TTM_SETDELAYTIME = (WM_USER + 3)
Private Const TTM_SETMARGIN = (WM_USER + 26)
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Private Const TTM_SETTITLEA = (WM_USER + 32)
Private Const TTM_SETTITLEW = (WM_USER + 33)
Private Const TTM_SETTOOLINFOA = (WM_USER + 9)
Private Const TTM_SETTOOLINFOW = (WM_USER + 54)
Private Const TTM_TRACKACTIVATE = (WM_USER + 17)
Private Const TTM_TRACKPOSITION = (WM_USER + 18)
Private Const TTM_UPDATE = (WM_USER + 29)
Private Const TTM_UPDATETIPTEXTA = (WM_USER + 12)
Private Const TTM_UPDATETIPTEXTW = (WM_USER + 57)
Private Const TTM_WINDOWFROMPOINT = (WM_USER + 16)

'Konstanta untuk menentukan gaya dari jendela tooltip.
'Selalu tip, walalupun jika jendela utama tidak aktif.
Private Const TTS_ALWAYSTIP = &H1
'Menggunakan gaya balon tooltip.
Private Const TTS_BALLOON = &H40
'Win98 and up - jangan gunakan sliding tooltips.
Private Const TTS_NOANIMATE = &H10
'Win2K and up - jangan hilangkan tooltips.
Private Const TTS_NOFADE = &H20
'Mencegah Windows dari penghapusan karakter ampersand 'apapun di dalam string tooltip. Tanpa penanda ini, 'Windows otomatis akan menghapus karakter ampersand 'dari string tersebut. Hal ini dilakukan untuk 'mengizinkan string yang sama dapat digunakan
'sebagai teks dari tooltip, dan sebagai tulisan dari 'sebuah control.
Private Const TTS_NOPREFIX = &H2

'Class untuk dua tooltip yang berbeda.
Private Const TOOLTIPS_CLASS = "tooltips_class"
Private Const TOOLTIPS_CLASSA = "tooltips_class32"

'Sebuah variabel bertipe Long untuk menyimpan hwnd '(window handle) dari jendela tooltip yang dibuat di 'contoh ini.Hal ini akan menjadi sebuah array bertipe 'Long jika kita membuat tooltip Windows untuk banyak 'control atau banyak jendela.
Dim hwndTT As Long
'===============@@ 6-12-2010 buat tooltip kalo ada sms yang masuk ===========================

'@@ 13-07-2012, Buat Verifikasi form yang hanya boleh dibuka di admin
Public CekVerifikasi As Boolean
 
Private Const GWL_STYLE = (-16)
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZE = &H20000
 
Private Const WS_THICKFRAME = &H40000
 
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
        (ByVal hwnd As Long, ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
 
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
        (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Sub add_special_history_Click()
    form_add_history.Show vbModal
End Sub

Private Sub apaja_Click()
    Form3.Show
End Sub

Private Sub cmd_break_Click()
    If MsgBox("Waktu istirahat akan diset sekarang juga??", vbYesNo + vbQuestion, "Confirm") = vbYes Then
        M_OBJCONN.Execute "UPDATE usertbl SET f_break=1"
        MsgBox "Waktu istirahat telah diset", vbOKOnly + vbInformation, "INFO"
    Else
        MsgBox "Waktu istirahat dibatalkan", vbOKOnly + vbInformation, "INFO"
    End If
End Sub

Private Sub cmdenabledptp_Click()
    Dim sql As String
    
    If cmdenabledptp.Left = 480 Then
       cmdenabledptp.Left = 0
       sql = "UPDATE enabledptp SET enabled = 0"
    Else
       cmdenabledptp.Left = 480
       sql = "UPDATE enabledptp SET enabled = 1"
    End If
    M_OBJCONN.Execute sql
    
    'Call enabledptp
End Sub

Private Sub enabledptp()
    Dim sql As String
    Dim M_objrs As ADODB.Recordset
    
    sql = "SELECT * FROM enabledptp"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    
    If M_objrs(0) = 0 Then
       cmdenabledptp.Left = 0
       Option1.Value = 1
       Option2.Value = 0
    Else
       cmdenabledptp.Left = 480
       Option1.Value = 0
       Option2.Value = 1
    End If
    'M_OBJCONN.Execute sql
End Sub

Private Sub Command1_Click()
    'Frm_DailyAplikasi.Show
    SSPanel_browse.Visible = False
End Sub

Private Sub CmdAddLeads_Click()
'        With FrmEntryReff
'            .TxtIdReff.Text = "Inbound Leads"
'            .TxtIdReff.Enabled = False
'             .Show vbModal
'             If .okReff Then
'             Else
'             End If
'        End With
End Sub

Private Sub CmdInbox_Click()
    FrmInboXSms.Caption = "SMS 1"
    FrmInboXSms.Show vbModal
End Sub

Private Sub Command3_Click()
    FormupdateDB.Show vbModal
End Sub

Private Sub Command4_Click()
    Form4.Show
End Sub

Private Sub Command5_Click()
    formattemptpercustidperagent.Show 1
End Sub

Private Sub Command6_Click()
    Call dashboard
End Sub

Private Sub Command7_Click()
Dim objExcel As New Excel.Application
Dim objExcelSheet As Excel.Worksheet
Dim col, row As Integer
Dim a As String
If LVAgent.ListItems.Count > 0 Then
    objExcel.Workbooks.ADD
    Set objExcelSheet = objExcel.Worksheets.ADD
 

    For col = 1 To LVAgent.ColumnHeaders.Count
        objExcelSheet.Cells(1, col).Value = LVAgent.ColumnHeaders(col)
    Next
 
    For row = 2 To LVAgent.ListItems.Count + 1
        For col = 1 To LVAgent.ColumnHeaders.Count
        If col = 1 Then
                objExcelSheet.Cells(row, col).Value = LVAgent.ListItems(row - 1).text
        Else
            '" 'cararandy 29032016 "
            Dim hasil1 As String
                hasil1 = "'" + LVAgent.ListItems(row - 1).SubItems(col - 1)
                objExcelSheet.Cells(row, col).Value = hasil1
            End If
        Next
    Next
 
    objExcelSheet.Columns.AutoFit
    Cd_save.ShowOpen
    a = Cd_save.FileName
 
    'objExcelSheet.SaveAs a & ".xlsx"
    objExcelSheet.SaveAs a & ".xls"
    MsgBox "Export Completed", vbInformation, Me.Caption
 
    'objExcel.Workbooks.Open a & ".xlsx"
    objExcel.Workbooks.Open a & ".xls"
    objExcel.Visible = True
Else
    MsgBox "No data to export", vbInformation, Me.Caption
End If
End Sub

Private Sub cuda_Click()
    form_upload_all.Show
End Sub

Private Sub Label10_Click()
    If UCase(Text1) <> "ADMIN" Then
        'Load frm_showsms
        'frm_showsms.Show vbModal
        FrmInboXSms.Caption = "SMS"
        FrmInboXSms.Show vbModal
    End If
End Sub

Private Sub Label12_Click()
    FormupdateDB.Show
End Sub

Private Sub Label9_Click()
    If UCase(Text1) <> "ADMIN" Then
        'Load frm_showsms
        'frm_showsms.Show vbModal
        FrmInboXSms.Caption = "SMS"
        FrmInboXSms.Show vbModal
    End If
End Sub

Private Sub LblBersihkan_Click()
    Dim a As String
    a = InputBox("P?", "P")
    If a = "DNN#123" Then
        FrmBersihkanNegoPTP.Show vbModal
    End If
End Sub

Private Sub LblJmlSmsBaru_Change()
    Label9 = "SMS BARU " & LblJmlSmsBaru.Caption & " SMS"
End Sub

Private Sub list_phone_review_Click()
    If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Or _
       UCase(Trim(MDIForm1.txtlevel.text)) = "AGENT" Then
        MsgBox "Mohon maaf, anda tidak memiliki akses!", vbOKOnly + vbCritical, "Informasi"
        Exit Sub
    End If
    
    If UCase(Trim(MDIForm1.txtlevel.text)) = "SUPERVISOR" Or _
       UCase(Trim(MDIForm1.txtlevel.text)) = "ADMINISTRATOR" Or UCase(Trim(MDIForm1.txtlevel.text)) = "MANAGER" Then
            Form_List_Phone_Review.Show vbModal
    End If
End Sub

Private Sub LstGrade_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    MDIForm1.LstGrade.SortKey = ColumnHeader.Index - 1
    MDIForm1.LstGrade.Sorted = True
End Sub

Private Sub LstGrade_DblClick()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    If LstGrade.ListItems.Count = 0 Then
    Else
        shedulePTP_Show = True
        If UCase(MDIForm1.txtlevel.text) = "AGENT" Then

            If UCase(MDIForm1.TxtUsername.text) <> Trim(UCase(MDIForm1.LstGrade.SelectedItem.SubItems(3))) Then

                MsgBox "Anda Tidak Berhak Untuk Mengedit Data Ini", vbCritical + vbOKOnly, "Aplikasi"
                Exit Sub
            End If
        End If
        If UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Then
        'If UCase(mdiform1.txtlevel.text) = "TEAMLEADER" Then
        Dim PO_AGENT As String
        If VIEW_MGMDATA.cmb_kdagent.text = "PULLOUT" Then
            Set M_objrs = New ADODB.Recordset
            M_objrs.CursorLocation = adUseClient
            CMDSQL = "SELECT PO_Agent FROM mgm where CUSTID='" & MDIForm1.LstGrade.SelectedItem.SubItems(1) & "'"
            M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            PO_AGENT = M_objrs!PO_AGENT
            Set M_objrs = Nothing
        Else
            PO_AGENT = MDIForm1.LstGrade.SelectedItem.SubItems(3)
        End If

        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        CMDSQL = "SELECT USERID FROM usertbl WHERE TEAM ='" + MDIForm1.TxtUsername.text + "' AND USERID = '" + PO_AGENT + "'"
        M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_objrs.RecordCount <> 0 Then

        Else
            MsgBox "Data Ini Milik Agent Team Leader Yang Lain", vbCritical + vbOKOnly, "Aplikasi"
            Set M_objrs = Nothing
            Exit Sub
        End If
        Set M_objrs = Nothing
    End If
    Me.MousePointer = vbHourglass
    Flag_mgm = False
    FrmCC_Colection.Show
    Me.MousePointer = vbNormal
    'frmCC_Colection.Show
End If
End Sub

Public Sub LoOut_Ext(number$)
    Dim cancelflag As Boolean
    Dim DialString$, FromModem$, dummy
    DialString$ = "ATDT" + number$ + ";" + vbCr
    On Error Resume Next
    If MSComm1.PortOpen Then
    Else
        If MDIForm1.TxtCommPort.text = Empty Then
            MsgBox "Tidak Ada Variable buat Comport", vbInformation + vbOKOnly
            Exit Sub
        End If
        MSComm1.CommPort = MDIForm1.TxtCommPort.text
        MSComm1.Settings = "9600,N,8,1"
        MSComm1.PortOpen = True
    End If
    Me.MousePointer = 11
    If err Then
        MsgBox err.Description, vbCritical + vbOKOnly, "Aplikasi"
        MSComm1.PortOpen = False
        cancelflag = True
        Me.MousePointer = 0
        Exit Sub
    End If
    MSComm1.InBufferCount = 0
    MSComm1.Output = DialString$
    Me.MousePointer = 0
    Do
        dummy = DoEvents()
        If MSComm1.InBufferCount Then
            FromModem$ = FromModem$ + MSComm1.Input
            If InStr(FromModem$, "OK") Then
          '      Beep
                WaitSecs (0.1)
                cancelflag = True
                Exit Do
            End If
            If InStr(FromModem$, "NO DIALTONE") Then
          '      Beep
          '      Beep
                MsgBox err.Description, vbInformation + vbOKOnly, "Aplikasi"
                cancelflag = True
                Exit Do
            End If
        End If
        If cancelflag Then
            cancelflag = False
            Me.MousePointer = 0
            Exit Do
        End If
    Loop
    If MSComm1.PortOpen = True And cancelflag = True Then
        MSComm1.Output = "ATH" + vbCr
        MSComm1.PortOpen = False
    End If
    Me.MousePointer = 0
End Sub

Private Sub MDIForm_Activate()
    JmlKoneksiReq = 0
    MaxKoneksiReq = 200
        
    'Call enabledptp
    
    TDBDate3.Value = Now
    TDBDate4.Value = Now
    
    If UCase(txtlevel.text) = "ADMINISTRATOR" Or UCase(txtlevel.text) = "ADMIN" Then
        mn_update_db.Visible = True
    Else
        mn_update_db.Visible = False
    End If
    
    If UCase(Trim(MDIForm1.txtlevel.text)) = "SUPERVISOR" Or _
       UCase(Trim(MDIForm1.txtlevel.text)) = "ADMINISTRATOR" Or _
       UCase(Trim(MDIForm1.txtlevel.text)) = "ADMIN" Then
        'cmd_break.Visible = True
    End If
    
    If UCase(Trim(MDIForm1.txtlevel.text)) <> "SUPERVISOR" And UCase(Trim(MDIForm1.txtlevel.text)) <> "MANAGER" Then
        'Disable Close Button
        DisableCloseBtn Me
        '=======================
        Dim hMenu   As Long
        Dim lStyle  As Long
    
        'disable MAXIMIZE button
        lStyle = GetWindowLong(Me.hwnd, GWL_STYLE)
        lStyle = lStyle And Not WS_MINIMIZE
        lStyle = lStyle And Not WS_MAXIMIZEBOX
        Call SetWindowLong(Me.hwnd, GWL_STYLE, lStyle)
    End If
    
    'jejaktian12042016
    'SSPanel5.Width = 100
    '================================
   'FrmConfidenceAnalysis.Show vbModal
   
    STATUS = "Connected"
    'Call insertlogcti(STATUS)
    Frame2.Left = Screen.Width - Frame2.Width
    
    WaitSecs 1
    If Label6.Caption = 0 Then
        Call tampil_reminder
    Else
        MDIForm1.Timer100.Enabled = True
    End If
End Sub
Private Sub tampil_reminder()
 Dim CMDSQL As String
 Dim M_objrs As ADODB.Recordset
    CMDSQL = "select * from tblnegoptp_log where date(promisedate)=date(now()) + 1 AND AGENT='" + MDIForm1.TxtUsername.text + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
    If M_objrs.RecordCount <> 0 Then
        If MDIForm1.txtlevel.text = "Admin" Or MDIForm1.txtlevel.text = "Supervisor" Then
            Form_reminder.Show
        End If
        Form_reminder.Show
        SetWindowPos Form_reminder.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    End If
    
    
Set M_objrs = Nothing
End Sub
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim strsql As String

    strsql = "UPDATE usertbl SET stsaplikasi=0 WHERE userid ='" + MDIForm1.TxtUsername.text + "'"

    M_OBJCONN.Execute (strsql)
    
    '@@ 13-04-2011 Hapus data ip
    strsql = "delete from tbl_ip where agent='"

    strsql = strsql + Trim(MDIForm1.TxtUsername.text) + "'"

    M_OBJCONN.Execute strsql
    
   ' Winsock2.Close
    
    '@@28012013 ini buat update status loginnya
    strsql = "UPDATE usertbl SET f_status_login=null,last_logout='now()' WHERE userid='"

    strsql = strsql + Trim(MDIForm1.TxtUsername.text) + "' AND usertype = '1'"

    M_OBJCONN.Execute strsql
    
    Call set_count_ol("log out")
    
    End
End Sub



Private Sub MDIForm_Unload(Cancel As Integer)
    strsql = "UPDATE usertbl SET  f_status_login=null,last_logout='now()' WHERE userid='"

    strsql = strsql + Trim(MDIForm1.TxtUsername.text) + "'"

    M_OBJCONN.Execute strsql
    
    Cancel = 1
End Sub

Private Sub mn_aoc_Click()
    FormAOC.Show vbModal
End Sub

Private Sub mn_confidence_list_Click()
    FrmConfidenceList.Show 1
End Sub

Private Sub mn_copyfile_Click()
    Form_CopyFIleCPA.Show vbModal
End Sub

Private Sub mn_deskcoll_perform2_Click()
    Form_deskcoll_performance2.Show 1
End Sub

Private Sub mn_list_complaint_Click()
    Frm_list_complaint.Show 1
End Sub

Private Sub mn_list_sid_Click()
    CekVerifikasi = False
    
    If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    
    'Jika Yang loginnya SPV/ADMIN--> Tanya Ulang Passwordnya untuk melakukan Swap Data
    If UCase(Trim(MDIForm1.txtlevel.text)) = "SUPERVISOR" Or _
        UCase(Trim(MDIForm1.txtlevel.text)) = "ADMINISTRATOR" Or _
        UCase(Trim(MDIForm1.txtlevel.text)) = "MANAGER" Or _
        UCase(Trim(MDIForm1.txtlevel.text)) = "ADMIN" Then
        

        FrmVerifikasiPassword.TxtUsername.text = UCase(MDIForm1.TxtUsername.text)

        FrmVerifikasiPassword.Show vbModal
        
        If CekVerifikasi = True Then
            FrmSID.Show vbModal
        Else
            MsgBox "Mohon maaf, password yang anda inputkan salah! Akses ditolak!", vbOKOnly + vbCritical, "Peringatan"
        End If
    End If
End Sub

Private Sub mn_monhly_bp_Click()
    Form_monthly_BP.Show 1
End Sub

Private Sub mn_option_hide_Click()
    Form_filter_hide.Show 1
End Sub

Private Sub mn_performance_Click()
    Form_deskcoll_performance.Show 1
End Sub

Private Sub mn_performance_reguler_Click()
    Form_deskcoll_performance_reguler.Show 1
End Sub

Private Sub mn_report_temp_Click()
    If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Or UCase(Trim(MDIForm1.txtlevel.text)) = "AGENT" Or UCase(Trim(MDIForm1.txtlevel.text)) = "ADMIN" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
        Form_report_temp.Show vbModal
End Sub

Private Sub mn_update_db_Click()
    On Error Resume Next
    'M_OBJCONN.Execute "ALTER TABLE mgm ALTER COLUMN product_desc TYPE varchar(100);"
    
    'M_OBJCONN.Execute "ALTER TABLE mgm ALTER COLUMN afaxno TYPE varchar(100);"
    
    M_OBJCONN.Execute "CREATE TABLE tbl_notif_sms( " & _
                    "id serial primary key, " & _
                    "id_sms bigint, " & _
                    "received_sms_date timestamp(0) without time zone," & _
                    "sender_number varchar(30)," & _
                    "text_sms text," & _
                    "agent varchar(30)," & _
                    "log_date timestamp(0) without time zone DEFAULT now(),custid varchar(50));"

    M_OBJCONN.Execute "CREATE TABLE tbl_notif_sms_log( " & _
                    "id serial primary key, " & _
                    "id_sms bigint, " & _
                    "received_sms_date timestamp(0) without time zone," & _
                    "sender_number varchar(30)," & _
                    "text_sms text," & _
                    "agent varchar(30)," & _
                    "log_date timestamp(0) without time zone DEFAULT now(),custid varchar(50));"
                    
    'M_OBJCONN1.Execute "ALTER TABLE inbox ADD f_notif smallint default 0;"
    
    M_OBJCONN.Execute "CREATE TABLE tblnotif_info(" & _
                    "id serial primary key," & _
                    "log_date timestamp(0) without time zone default now()," & _
                    "type_notif varchar(50)," & _
                    "notif_from VarChar(30));"
                   
    MsgBox "Update Database Successfully !!!"
End Sub



Private Sub mnaddclient_Click()
    If MDIForm1.txtlevel.text = "Administrator" Or MDIForm1.txtlevel.text = "Admin" Then
        form_add_new_client.Show 1
    End If
End Sub

Private Sub mnadmin_Click()
If UCase(MDIForm1.txtlevel.text) = "ADMIN" Or UCase(MDIForm1.txtlevel.text) = "ADMINISTRATOR" Then
    Form_Setup_User.sKdlevel = "5"
    Form_Setup_User.sUsertype = "5"
    Form_Setup_User.Show vbModal
End If
End Sub

Private Sub mnbalance_Click()
    FrmBalance.Show vbModal
End Sub

Private Sub MnBb_Click()
'    FRM_Bulletin_LIST.Show vbModal
End Sub

Private Sub mnblack_Click()
    If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    
    If UCase(Trim(MDIForm1.txtlevel.text)) = "SUPERVISOR" Or _
       UCase(Trim(MDIForm1.txtlevel.text)) = "ADMINISTRATOR" Or _
       UCase(Trim(MDIForm1.txtlevel.text)) = "ADMIN" Or _
       UCase(Trim(MDIForm1.txtlevel.text)) = "MANAGER" Then

        FrmVerifikasiPassword.TxtUsername.text = UCase(MDIForm1.TxtUsername.text)

        FrmVerifikasiPassword.Show vbModal
        
        If CekVerifikasi = True Then
            frm_BlackListNo_List.Show vbModal
        Else
            MsgBox "Mohon maaf, password yang anda inputkan salah! Akses ditolak!", vbOKOnly + vbCritical, "Peringatan"
        End If
    End If
    
    'frm_BlackListNo_List.Show 1
End Sub

Private Sub mnblokspv_Click()
    'frmlockaccountfromspv.Show 1
    frm_map_lock_acc.Show 1
End Sub

Private Sub MnCCode_Click()
    FRM_Complaint_LIST.Show vbModal
End Sub

Private Sub mncontacted_Click()
    FRM_ContactedDesc_LIST.Show vbModal
End Sub

Private Sub mndata_Click()
  '  FRM_DataQuality_LIST.Show vbModal
End Sub


Private Sub mndistribut_Click()
If MDIForm1.txtlevel = "Supervisor" Then
    Form_distributeteam.Show
    'Form_distribusiteam.ZOrder vbBringToFront
ElseIf MDIForm1.txtlevel.text = "Administrator" Or MDIForm1.txtlevel.text = "Admin" Then
    Form_distribute.Show
    Form_distribute.ZOrder vbBringToFront
End If
End Sub

Private Sub mndpc_Click()
    frmdistributecustid.Show
End Sub

Private Sub mndrm_Click(Index As Integer)
    FrmRestoreRemarks.Show vbModal
End Sub

Private Sub MnDuplikasi_Click()
   ' FrmDuplikasi.Show
End Sub

Private Sub MNDUPLIKASICH_Click()
   ' FrmDuplikasiCh.Show
End Sub

Private Sub mnkrmaplikasi_Click()
   ' FrmKurirSpvApp.Show vbModal
End Sub

Private Sub mnlast_Click()
    Form_rpt_last_status.Show
End Sub

Private Sub mnLDS_Click()
    formLockData.Show
End Sub

Private Sub mnListAccountLunas_Click()
       CekVerifikasi = False
    
    If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    
    'Jika Yang loginnya SPV/ADMIN--> Tanya Ulang Passwordnya untuk melakukan Swap Data
    If UCase(Trim(MDIForm1.txtlevel.text)) = "MANAGER" Or UCase(Trim(MDIForm1.txtlevel.text)) = "ADMIN" Then

        FrmVerifikasiPassword.TxtUsername.text = UCase(MDIForm1.TxtUsername.text)

        FrmVerifikasiPassword.Show vbModal
        
        If CekVerifikasi = True Then
            FrmAccLunas.Show vbModal
        Else
            MsgBox "Mohon maaf, password yang anda inputkan salah! Akses ditolak!", vbOKOnly + vbCritical, "Peringatan"
            Exit Sub
        End If
    End If
    
    FrmAccLunas.Show vbModal
End Sub

Private Sub mnMenuRole_Click()
    Form_menu_Role.Show vbModal
End Sub

Private Sub mnmgr_Click()
    Form_Setup_User.sKdlevel = "3"
    Form_Setup_User.sUsertype = "3"
    Form_Setup_User.Show vbModal
End Sub

Private Sub mnmonthcpa_Click()
    Form_List_CPA.Show vbModal
End Sub

Private Sub mnMPD_Click()
    formmonitoringperpindahandata.Show
End Sub

Private Sub mnNact_Click()
If UCase(MDIForm1.txtlevel.text) = "ADMIN" Or UCase(MDIForm1.txtlevel.text) = "ADMINISTRATOR" Then
    Form_statuscall.Show vbModal
End If
End Sub

Private Sub mnnsms_Click()
    FormReport_sms.Show
End Sub

Private Sub mnpd_Click()
    form_tarikdata.Show
End Sub

Private Sub mnPO_Click()
    frmdeletedata.Show
End Sub

Private Sub mnptppayment_Click()
    Form_ptp_payment.Show vbModal
End Sub

Private Sub mnqcd_Click()

'If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Then
'    MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
'    Exit Sub
'End If
frm_cekdata.Show

'If UCase(MDIForm1.txtlevel.text) = "ADMIN" Or UCase(MDIForm1.txtlevel.text) = "ADMINISTRATOR" Then
'    frm_cekdata.Show
'End If

End Sub

Private Sub mnract_Click()
    Form_Report_activity.Show
End Sub

Private Sub mnrboard_Click()
    FormReport_dashboard.Show
End Sub

Private Sub mnrdetail_Click()
    Form_Report_call_detail.Show
End Sub

Private Sub mnrdistribut_Click()
    FormReport_distribusi.Show
End Sub

Private Sub mnrecycle_Click()
    Form_recycle.Show
End Sub

Private Sub MnReportTracking_Click()
 FrmMgmReport.Show
End Sub

Private Sub mnrmis_Click()
    'Form_Report_dika.Show
    Form_Report_MIS.Show
End Sub

Private Sub mnrole_Click()
If UCase(MDIForm1.txtlevel.text) = "ADMIN" Or UCase(MDIForm1.txtlevel.text) = "ADMINISTRATOR" Then
    Form_menu_Role.Show
End If
End Sub

Private Sub mnroutrpt_Click()
    Form_rpt_reason.Show
End Sub

Private Sub mnrptsms_Click()
    If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    frm_report_sms1.Show 1
End Sub

Private Sub mnrresult_Click()
    'Form_Report_Submit.Show
    Form_rpt_reason_detail.Show
End Sub

Private Sub mnrsummery_Click()
    frm_reportsummery.Show
End Sub

Private Sub mnsubahstsacc_Click()
    If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    frnubahstsaccount.Show 1
End Sub

Private Sub mnsubmarkup_Click()
'    If UCase(Trim(mdiform1.txtlevel.text)) = "TEAMLEADER" Then
'        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
'        Exit Sub
'    End If
    
    'Jika Yang loginnya SPV/ADMIN--> Tanya Ulang Passwordnya untuk melakukan Swap Data
    If UCase(Trim(MDIForm1.txtlevel.text)) = "SUPERVISOR" Or _
       UCase(Trim(MDIForm1.txtlevel.text)) = "ADMINISTRATOR" Or _
       UCase(Trim(MDIForm1.txtlevel.text)) = "ADMIN" Or _
       UCase(Trim(MDIForm1.txtlevel.text)) = "MANAGER" Or _
       UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Then

        FrmVerifikasiPassword.TxtUsername.text = UCase(MDIForm1.TxtUsername.text)

        FrmVerifikasiPassword.Show vbModal
        
        If CekVerifikasi = True Then
            FRMMARKUP.Show vbModal
        Else
            MsgBox "Mohon maaf, password yang anda inputkan salah! Akses ditolak!", vbOKOnly + vbCritical, "Peringatan"
        End If
    End If
    
    
    'FRMMARKUP.Show 1
End Sub

Private Sub mntd_Click()
    Form_Tarik_Data.Show
End Sub

Private Sub mntl_Click()
If UCase(MDIForm1.txtlevel.text) = "ADMIN" Or UCase(MDIForm1.txtlevel.text) = "ADMINISTRATOR" Then
    Dim m_msgbox As String
    If (MDIForm1.txtlevel.text = "TeamLeader") Then
        m_msgbox = MsgBox("Anda tidak memiliki hak akses!", vbOKOnly + vbExclamation, "Peringatan")
        Exit Sub
    End If
'    frm_list_team_leader.Show vbModal
    Form_Setup_User.sKdlevel = "2"
    Form_Setup_User.sUsertype = "2"
    Form_Setup_User.Show vbModal
End If
End Sub

Private Sub mntquery_Click()
    Form_QueryAnlyzer.Show
End Sub

Private Sub mnuCallmonitor_Click()
    Form_call_mon.Show
End Sub

Private Sub MNUOFFER_Click()
    If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    frmheaderoffeer.Show 1
End Sub

Private Sub mnuploadskip_Click()
    If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    frmuploadskiptarcer.Show 1
End Sub
Private Sub mnuProdInfo_Click()
    FrmProductList.Show
End Sub

Private Sub mnuuploadcpa_Click()
    If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Or UCase(Trim(MDIForm1.txtlevel.text)) = "ADMIN" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    
    Form_upload_CPA.Show vbModal
End Sub

Private Sub MnVisit_Click()
    FrmVisit.Show
End Sub

Private Sub MSComm1_OnComm()

End Sub

Private Sub nmAksesLayanaTelkom_Click()
       If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    FrmAksesLayananTelkom.Show vbModal
End Sub

Private Sub nmapprovreject_Click()
    If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    'frm_approved_rejected.Show vbModal
    M_OBJCONN.Execute "DELETE FROM tblnotif_info WHERE type_notif='sms' "
    Frm_verify.Show vbModal
End Sub

Private Sub nmbackup_Click()
     If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Or UCase(Trim(MDIForm1.txtlevel.text)) = "ADMIN" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    FrmBackupDbToExcel.Show vbModal
End Sub

Private Sub nmblastsmsexcel_Click()
    If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    FrmSmsBlastExcel.Show
End Sub

Private Sub nmblokaplikasitins_Click()
    
    If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Or _
       UCase(Trim(MDIForm1.txtlevel.text)) = "AGENT" Then
        MsgBox "Mohon maaf, anda tidak memiliki akses!", vbOKOnly + vbCritical, "Informasi"
        Exit Sub
    End If
    
    'Jika Yang loginnya SPV/ADMIN--> Tanya Ulang Passwordnya untuk melakukan Swap Data
    If UCase(Trim(MDIForm1.txtlevel.text)) = "SUPERVISOR" Or _
       UCase(Trim(MDIForm1.txtlevel.text)) = "ADMINISTRATOR" Or UCase(Trim(MDIForm1.txtlevel.text)) = "MANAGER" Then
       'UCase(Trim(mdiform1.txtlevel.text)) = "ADMIN"
        ' REQUEST JOKO TGL. 30 SEP 2013 - Tanpa Verifikasi
        'FrmVerifikasiPassword.txtusername.text.Text = UCase(mdiform1.txtusername.text.text)
        'FrmVerifikasiPassword.Show vbModal
        
'        If CekVerifikasi = True Then
            FrmBlokAgent.Show vbModal
'        Else
'            MsgBox "Mohon maaf, password yang anda inputkan salah! Akses ditolak!", vbOKOnly + vbCritical, "Peringatan"
'        End If
    End If
End Sub

Private Sub nmconfidenceanalisysagent_Click()
    FrmConfidenceListNew_Agent.Show vbModal
End Sub

Private Sub nmenu_Click()
    FrmListReqTlp.Show vbModal
End Sub

Private Sub nmformceksts_Click()
    Frm_Cek_status_acc.Show 1
End Sub

Private Sub nmListReportProblemTelepon_Click()
    Dim a As String
    a = InputBox("Password?", "@@@@@")
    If a = "Dnn#12345" Then
        FrmListReportTelepon.Show vbModal
    Else
        MsgBox "Akses di tolak!", vbOKOnly + vbExclamation, "Peringatan"
    End If
End Sub

Private Sub nmlistreqform_Click()
    FrmListRequest.Show vbModal
End Sub

Private Sub nmlistreqptp_Click()
If UCase(MDIForm1.txtlevel.text) = "ADMIN" Or UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Or UCase(MDIForm1.txtlevel.text) = "ADMINISTRATOR" Then
    FrmListRequestPTP.Show vbModal
End If
End Sub

Private Sub nmlistsendcpa_Click()
    If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    FrmSendCPA.Show vbModal
End Sub

Private Sub nmlistsmsscript_Click()
    If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    Frm_List_SMS_Script.Show vbModal
End Sub

Private Sub nmListUnValidNumber_Click()
    FrmListUnValidNumber.Show vbModal
    
End Sub

Private Sub nmlstreqnumber_Click()
    FrmListReqTlp.Show vbModal
End Sub

Private Sub nmManageDistribusiAccount_Click()
    CekVerifikasi = False
    
    If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Or UCase(Trim(MDIForm1.txtlevel.text)) = "ADMIN" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    
    'Jika Yang loginnya SPV/ADMIN--> Tanya Ulang Passwordnya untuk melakukan Swap Data
    If UCase(Trim(MDIForm1.txtlevel.text)) = "SUPERVISOR" Or _
    UCase(Trim(MDIForm1.txtlevel.text)) = "ADMINISTRATOR" Or _
           UCase(Trim(MDIForm1.txtlevel.text)) = "MANAGER" Then
        'FrmVerifikasiPassword.txtusername.text.Text = UCase(mdiform1.txtusername.text.text)
        'FrmVerifikasiPassword.Show vbModal
        
        'If CekVerifikasi = True Then
            FrmDistribusiAcc.Show vbModal
        'Else
        '    MsgBox "Mohon maaf, password yang anda inputkan salah! Akses ditolak!", vbOKOnly + vbCritical, "Peringatan"
        'End If
    End If
    
End Sub

Private Sub nmmenuformlistconfidence_Click()
    FrmConfidenceListNew.Show vbModal
End Sub

Private Sub nmreportcall_Click()
    'rptCallTracking.Show vbModal
End Sub

Private Sub nmReportCallServer5_Click()
    RptCallTrackingServer5.Show vbModal
End Sub

Private Sub nmReportProblemHeadset_Click()
    Dim a As String
    a = InputBox("Password?", "@@@@@")
    If a = "Dnn#12345" Then
        FrmListProblemHeadset.Show vbModal
    Else
        MsgBox "Akses ditolak!", vbOKOnly + vbCritical, "Peringatan"
    End If
End Sub

Private Sub nmReportSms_Click()
    frm_report_sms.Show vbModal
End Sub

Private Sub nmresetpass_Click()
    FrmResetPass.Show vbModal
End Sub

Private Sub nmrestoredeleteacc_Click()
     If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Or UCase(Trim(MDIForm1.txtlevel.text)) = "ADMIN" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    FrmRestoreDelete.Show vbModal
End Sub

Private Sub nmRptCallServer4_Click()
    rptCallTrackingServer4.Show vbModal
End Sub

Private Sub nmSchLocktl_Click()
'DIBUKA LAGI BY REQUEST DODDY 5-6-2015
'    If UCase(Trim(mdiform1.txtlevel.text)) = "TEAMLEADER" Then
'        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
'        Exit Sub
'    End If
'
    'Jika Yang loginnya SPV/ADMIN--> Tanya Ulang Passwordnya untuk melakukan Swap Data
    If UCase(Trim(MDIForm1.txtlevel.text)) = "SUPERVISOR" Or _
       UCase(Trim(MDIForm1.txtlevel.text)) = "ADMINISTRATOR" Or _
       UCase(Trim(MDIForm1.txtlevel.text)) = "ADMIN" Or _
       UCase(Trim(MDIForm1.txtlevel.text)) = "MANAGER" Or _
       UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Then
        'FrmVerifikasiPassword.txtusername.text.Text = UCase(mdiform1.txtusername.text.text)
        'FrmVerifikasiPassword.Show vbModal
        
'        If CekVerifikasi = True Then
            frm_list_schedule_tl.Show vbModal
'        Else
'            MsgBox "Mohon maaf, password yang anda inputkan salah! Akses ditolak!", vbOKOnly + vbCritical, "Peringatan"
'        End If
    End If
    
    'frm_list_schedule_tl.Show 1
End Sub



Private Sub nmswapdata_Click()
    CekVerifikasi = False
    
    If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Or UCase(Trim(MDIForm1.txtlevel.text)) = "ADMIN" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    
    'Jika Yang loginnya SPV/ADMIN--> Tanya Ulang Passwordnya untuk melakukan Swap Data
    If UCase(Trim(MDIForm1.txtlevel.text)) = "SUPERVISOR" Or _
    UCase(Trim(MDIForm1.txtlevel.text)) = "ADMINISTRATOR" Or _
           UCase(Trim(MDIForm1.txtlevel.text)) = "MANAGER" Then

        FrmVerifikasiPassword.TxtUsername.text = UCase(MDIForm1.TxtUsername.text)

        FrmVerifikasiPassword.Show vbModal
        
        If CekVerifikasi = True Then
            Form_swap.Show vbModal
        Else
            MsgBox "Mohon maaf, password yang anda inputkan salah! Akses ditolak!", vbOKOnly + vbCritical, "Peringatan"
        End If
    End If
End Sub

Private Sub nmuploadaddress_Click()
    Frm_Upload_address.Show
End Sub

Private Sub nmuploadcpaptp_Click()
    FrmUploadCPAPTP.Show vbModal
End Sub

Private Sub nmuploadcustomer_Click()
    If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Or UCase(Trim(MDIForm1.txtlevel.text)) = "SUPERVISOR" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    
    form_upload_customer.Show vbModal
    'Frm_Upload_Data.Show vbModal
End Sub

Private Sub nmuploadpayment_Click()
    If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Or UCase(Trim(MDIForm1.txtlevel.text)) = "SUPERVISOR" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    Frm_Upload_Payment.Show vbModal
End Sub

Private Sub nmuploadtempdata_Click()
    If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Or UCase(Trim(MDIForm1.txtlevel.text)) = "SUPERVISOR" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    FrmUploadTempData.Show vbModal
End Sub

Private Sub Option1_Click()
    Option2.Value = 0
    sql = "UPDATE enabledptp SET enabled = 0"
    M_OBJCONN.Execute sql
End Sub

Private Sub Option2_Click()
    Option1.Value = 0
    sql = "UPDATE enabledptp SET enabled = 1"
    M_OBJCONN.Execute sql
End Sub

Private Sub rapcpa_Click()

End Sub

Private Sub rrld_Click()
    If UCase(txtlevel.text) = "ADMIN" Then
        formreportremarkslastday.Show 1
    Else
        MsgBox "Anda Tidak Memiliki Akses untuk Report Ini"
    End If
End Sub

Private Sub setspv_Click()
    If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    frmsettarget.Show 1
End Sub

Private Sub smsblast_Click()
    If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    Me.MousePointer = vbHourglass
    load frm_sms_blast
    DoEvents
    frm_sms_blast.Show vbModal
    DoEvents
End Sub

Private Sub SSCommand1_Click(Index As Integer)
Dim M_objrs As ADODB.Recordset
Dim CMDSQL As String
Dim m_objrscekmonitoring As ADODB.Recordset

Select Case Index
    Case 8
        'Label6.Caption = 0
        'Call tampil_reminder
        Form_reminder.Show
    Case 0
        m_targetview = True
        If MDIForm1.txtlevel.text = "Agent" Then
            VIEW_MGMDATA.LblTarget(0).Caption = LblTarget.Caption
            VIEW_MGMDATA.LblTarget(1).Caption = LblTarget.Caption
            VIEW_MGMDATA.LstVwSearchMgm.Checkboxes = False
        End If
        
        Dim ds As New ADODB.Recordset
        
        ds.CursorLocation = adUseClient

        ds.Open "select lockdarispv, F_LOCK, f_akses_all_acc FrOM usertbl WHERE USERID='" & MDIForm1.TxtUsername.text & "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

        
        If ds.BOF And ds.EOF Then
        Else
            If ds!F_LOCK = "Y" Then
                F_LOCK = True
            Else
                F_LOCK = False
            End If
        End If
        
        If ds.EOF = False Then
            cek_aksesall = cnull(ds("f_akses_all_acc"))
        Else
            cek_aksesall = 0
        End If
        
        If cek_aksesall = "1" Then
            VIEW_MGMDATA.CmdSearchPTP.Enabled = False
        End If
        
        If cek_aksesall = "0" Then
            VIEW_MGMDATA.CmdSearchPTP.Enabled = True
        End If
        
        ' CEK LOCK ACC
        If ds.EOF = False Then
            If ds("lockdarispv") <> "" Then
                VIEW_MGMDATA.CmdSearchPTP.Enabled = False
            End If
        End If
                    
        If UCase(Trim(MDIForm1.txtlevel.text)) = "AGENT" Or UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Then
            '@@ 10-03-2011 Tambahan buat nambahin monitoring headset
            'cek dulu apakah status monitoring aktif

'            CMDSQL = "select * from manajemen_site  where status='1'"
'            Set M_objrs = New ADODB.Recordset
'            M_objrs.CursorLocation = adUseClient
'            M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'                If M_objrs.RecordCount > 0 Then
'                    CMDSQL = "select monitoring_headset from usertbl where userid='"
'                    CMDSQL = CMDSQL + Trim(MDIForm1.TxtUsername.Text) + "'"
'
'                    Set m_objrscekmonitoring = New ADODB.Recordset
'                    m_objrscekmonitoring.CursorLocation = adUseClient
'                    m_objrscekmonitoring.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'                    If Trim(m_objrscekmonitoring("monitoring_headset")) = "1" Then
'                        FrmMonitoringHeadset.Show vbModal
'                        Set m_objrscekmonitoring = Nothing
'                    Else
'                        VIEW_MGMDATA.Show
'                        Set M_objrs = Nothing
'                        Set m_objrscekmonitoring = Nothing
'                    End If
'                Else
'                    VIEW_MGMDATA.Show
'                    Set M_objrs = Nothing
'                End If
        Else
            VIEW_MGMDATA.Show
        End If
        'FRM_SEARCH.Show
    Case 1
        'FrmmgmReportKeDua.Show
   ' Case 2
    '    FrmPreembosReport.Show
    Case 10
        FRMSENDMSG.Show vbModal
        'FRMSENDMSG.Show
    Case 4
        
    Case 5
        FrmVisit.Show
    Case 6
'        FrmSearching.Show
        FrmCari.Show
    Case 7
        fmunlock.Show
    'Case 8
    '    FrmAccessData.Show
    Case 9
       ' FrmmgmReport.Show
    Case 11
        FrmMgmReport.Show ' UTK RITCARD
    Case 12
        FrmInboXSms.Show vbModal
    Case 17
        Call MnFile_Click(7)
    'FrmMgmReport_AWARNESS.Show ' utk RITPIL dan AwarNESS
End Select
End Sub

Private Sub MDIForm_Load()
    addphone = False
    lg_call = False
    
    Dim m_data      As New CLS_LOGIN
    Dim M_LOGINRS   As ADODB.Recordset
    Dim m_port      As String
    
    waktu_start = waktu_server_sekarang
    COUNTER = 0
    count_timer_detik = 0
    
    ' MONITORING ACTIVITY BY IZUDDIN 16 04 2013
    i_monitoring_activity = 0
    lbl_timer_activity = 0
    open_sms = False
    ' #########################################
    
    'Timer2.Enabled = False

    On Error GoTo hell
    'Winsock2.Listen
                
    bRenderrecord = False
    
    Call PromiseToPay
    Call HeaderInformation
    Call LstDataInformation
    
    'SSTab1.TabVisible(1) = False
    'Call tglhost
    Set M_LOGINRS = New ADODB.Recordset
    M_LOGINRS.CursorLocation = adUseClient
    M_LOGINRS.Open "SELECT * FROM vwcallcfg1", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_LOGINRS.RecordCount <> 0 Then
        Text6.text = IIf(IsNull(M_LOGINRS("DELAY_TONE")), "0", M_LOGINRS("DELAY_TONE"))
        TxtAuthPrefix.text = IIf(IsNull(M_LOGINRS("AUTHPREFIX")), "", M_LOGINRS("AUTHPREFIX"))
        TxtModemAcod.text = IIf(IsNull(M_LOGINRS("MODEMACOD")), "", M_LOGINRS("MODEMACOD"))
        TxtCommPort.text = IIf(IsNull(M_LOGINRS("COMMPORT")), "", M_LOGINRS("COMMPORT"))
        TDBDate1.Value = IIf(IsNull(M_LOGINRS("TglSystem")), "", M_LOGINRS("TglSystem"))
        TxtJamMulaiTelp.text = IIf(IsNull(M_LOGINRS("JAMMULAITELP")), "", M_LOGINRS("JAMMULAITELP"))
        TxtJamSelesaiTelp.text = IIf(IsNull(M_LOGINRS("JAMSELESAITELP")), "", M_LOGINRS("JAMSELESAITELP"))
        TxtLamaFollowup.text = IIf(IsNull(M_LOGINRS("LAMAFOLLOWUP")), "99", M_LOGINRS("LAMAFOLLOWUP"))
    Else
        TDBDate1.Value = Now
        TxtLamaFollowup.text = "99"
    End If

    m_port = BUKA_FILE_KONEKSI("comport.txt")
    If m_port <> "" Then
        TxtCommPort.text = m_port
    End If
    M_LOGINRS.Close
    Set M_LOGINRS = Nothing
    
    Exit Sub
    
hell:
    MsgBox err.Description + "only one aplication is open", vbCritical + vbOKOnly, "Warning"
    End
    
End Sub

Private Sub messageapptransfer()
    cm_connect.Enabled = False
    cm_disconnect.Enabled = False
    cm_send.Enabled = False
End Sub

Private Sub mnabout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnagent_Click()
    
If UCase(MDIForm1.txtlevel.text) = "ADMIN" Or UCase(MDIForm1.txtlevel.text) = "ADMINISTRATOR" Then
    Form_Setup_User.sKdlevel = "1"
    Form_Setup_User.sUsertype = "1"
    Form_Setup_User.Show vbModal
End If
End Sub

Private Sub mnbaca_Click()
    FRMBACAMSG.Show vbModal
End Sub

Private Sub mndata2_Click()
    FRM_DATASOURCE_LIST.Show vbModal
End Sub

Private Sub MnFile_Click(Index As Integer)
    Dim strsql As String
    Select Case Index
        Case 0
            Unload MDIForm1
            frmLogin.Show vbModal
        Case 1
            FRM_SET_PWD.Show
        Case 3
            frmsetpassword.Show vbModal
        Case 5
            frm_gantipas.Text1(2).text = UCase(MDIForm1.TxtUsername.text)

            frm_gantipas.Show vbModal
        Case 6
        Case 7
            strsql = "update usertbl set stsaplikasi=0  where userid ='" + MDIForm1.TxtUsername.text + "'"

            M_OBJCONN.Execute (strsql)
            '@@ 13-04-2011 Hapus data ip
            strsql = "delete from tbl_ip where agent='"

            strsql = strsql + Trim(MDIForm1.TxtUsername.text) + "'"

            M_OBJCONN.Execute strsql
            Unload Me
        End
    End Select
End Sub

Private Sub mnhslupload_Click()
    FRM_HASILUPLOAD.Show vbModal
End Sub

Private Sub mnproduct_Click()
    FRM_PRODUCT_LIST.Show vbModal
End Sub

Private Sub mnreason_Click()
    FRM_CLOSSING_LIST.Show vbModal
End Sub
Private Sub mnsend_Click()
    FRMSENDMSG.Show vbModal
End Sub

Private Sub mnspv_Click()
    FRM_SPV_LIST.Show vbModal
End Sub

Private Sub mnup_Click()
    FRM_SETUSER.Show vbModal
End Sub

Private Sub SSCommand2_Click()
    Dim M_objrs As New ADODB.Recordset
    Dim CMDSQL As String
    Dim ListItem As ListItem
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    CMDSQL = "SELECT custid,f_cek,agent FROM mgm where f_pending='pending' "
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic
    If M_objrs.RecordCount <> 0 Then
        LstGrade.ListItems.CLEAR
    End If
    While Not M_objrs.EOF
        Set ListItem = LstGrade.ListItems.ADD(, , M_objrs.Bookmark)
        ListItem.SubItems(1) = IIf(IsNull(M_objrs("CUSTID")), "", M_objrs("CUSTID"))
        ListItem.SubItems(2) = IIf(IsNull(M_objrs("F_CEK")), "", M_objrs("F_CEK"))
        ListItem.SubItems(3) = IIf(IsNull(M_objrs!AGENT), "", M_objrs!AGENT)
        M_objrs.MoveNext
    Wend
    If M_objrs.RecordCount = 0 Then
    LstGrade.ListItems.CLEAR
    End If
    Set M_objrs = Nothing
End Sub

Private Sub SSPanel4_Click()
    
    If SSPanel4.Height = 360 Then
        Call supervisorole
        Call dashboard
        SSPanel4.Height = 5505
        SSPanel4.BackColor = &H8000000D
    Else
        SSPanel4.Height = 360
        SSPanel4.BackColor = &H8000000D
    End If
End Sub

Private Sub supervisorole()
    Dim M_objrs As ADODB.Recordset
    Dim M_objrsc As ADODB.Recordset
    If UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Then
        q = "select distinct recsource from mgm where agent in (select userid from usertbl where team = '" & MDIForm1.TxtUsername.text & "' or userid = '" & MDIForm1.TxtUsername.text & "' )  "
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        Dim zz As String
        Dim sss As String
        
        qs = "select * from tbl_list_client_indium order by 1"
        Set M_objrsc = New ADODB.Recordset
        M_objrsc.CursorLocation = adUseClient
        M_objrsc.Open qs, M_OBJCONN, adOpenDynamic, adLockOptimistic
        
        sss = ""
        zz = ""
        Combo1.CLEAR
        While Not M_objrs.EOF
        'If M_objrs.RecordCount > 0 Then
                For i = 1 To M_objrsc.RecordCount
                    a = M_objrsc!client
                    If M_objrs!recsource Like "*" & a & "*" Then
                        If M_objrsc!client = "PLUS" Then
                            If sss Like "*PLUS*" Then
                            Else
                                Combo1.AddItem "RUPIAH PLUS"
                                sss = sss & " PLUS "
                            End If
                        ElseIf M_objrsc!client = "EXPRES" Then
                            If sss Like "*EXPRES*" Then
                            Else
                                Combo1.AddItem "UANGEXPRESS"
                                sss = sss & " EXPRES "
                            End If
                        ElseIf M_objrsc!client = "GLOBAL" Then
                            If sss Like "*GLOBAL*" Then
                            Else
                                Combo1.AddItem "GLOBALINDO"
                                sss = sss & " GLOBAL "
                            End If
                        Else
                            'If zz Like "*" & aa(i - 1) & "*" Then
                            'Else
                            If sss Like "*" & M_objrsc!client & "*" Then
                            Else
                                Combo1.AddItem M_objrsc!client
                                zz = zz & " " & M_objrsc!client
                                sss = sss & " " & M_objrsc!client & " "
                            End If
                        End If
                    End If
                    M_objrsc.MoveNext
                Next i
            M_objrsc.MoveFirst
            M_objrs.MoveNext
        'End If
        Wend
        
    End If

End Sub

Private Sub dashboard()
   
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    CMDSQL = " select tblstatuscall_kdstscall as stts from tblstatuscall where tblstatuscall_kdstatus = '1' order by tblstatuscall_keterangan"
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LVAgent.ColumnHeaders.CLEAR
    LVAgent.ColumnHeaders.ADD 1, , "No", 10 * 120
    LVAgent.ColumnHeaders.ADD 2, , "Agent", 10 * 120
    z = 3
    While Not M_objrs.EOF
        LVAgent.ColumnHeaders.ADD z, , "" & M_objrs!stts & "", 7 * 120
        M_objrs.MoveNext
        z = z + 1
    Wend
    LVAgent.ColumnHeaders.ADD z, , "TOTAL", 10 * 120
    
    'isi
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    CMDSQL = " select tblstatuscall_keterangan as stts from tblstatuscall where tblstatuscall_kdstatus = '1' order by 1"
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    a = ""
    B = ""
    c = ""
    
    While Not M_objrs.EOF
        a = a + " ,case when kodeds = '" & "" & M_objrs!stts & "" & "' then 1 else 0 end as """ & "" & M_objrs!stts & """"
        B = B + """" & "" & M_objrs!stts & """+"
        c = c + " ,sum( """ & "" & M_objrs!stts & """" & " ) as """ & "" & M_objrs!stts & """"
        M_objrs.MoveNext
    Wend
        B = Left(B, Len(B) - 1)
        c = c
    
    q = " select agent" & "" & c & ", sum(total) as total from ("
    q = q + "select *," & "" & B & " as Total from ("
    q = q & "select agent " & "" & a & ""
    q = q & "from (select a.agent, a.custid, a.kodeds, b.recsource from mgm_hst a inner join mgm b on a.custid = b.custid "
    q = q & " where tgl between '" & Format(TDBDate3.Value, "yyyy-mm-dd") & " 00:00:00' and '" & Format(TDBDate4.Value, "yyyy-mm-dd") & " 23:59:59' "
    
    If UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Then
        q = q & " and b.recsource in (select distinct recsource from mgm where agent in (select userid from usertbl where team = '" & MDIForm1.TxtUsername.text & "' or userid = '" & MDIForm1.TxtUsername.text & "' )) "
    End If
    
    If Combo1.text = "RUPIAH PLUS" Then
        q = q & " and recsource ilike '%PLUS%') hst "
    ElseIf Combo1.text = "UANGEXPRESS" Then
        q = q & " and recsource ilike '%EXPRESS%') hst "
    ElseIf Combo1.text = "GLOBALINDO" Then
        q = q & " and recsource ilike '%GLOBAL%') hst "
    Else
        q = q & " and recsource ilike '%" & Combo1.text & "%') hst "
    End If
    
    q = q & " ) abc "
    q = q & " ) a group by agent "
    Set M_objrs = New ADODB.Recordset
    
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open q, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    LVAgent.ListItems.CLEAR
    While Not M_objrs.EOF
        Set ListItem = LVAgent.ListItems.ADD(, , M_objrs.Bookmark)
        For i = 1 To z - 1
            ListItem.SubItems(i) = IIf(IsNull(M_objrs(i - 1)), "", M_objrs(i - 1))
        Next i
        M_objrs.MoveNext
    Wend

End Sub

'Private Sub SSCommand3_Click()
'    Form_manual_dial.Show
'End Sub

'Private Sub SSPanel5_Click()
'    If SSPanel5.Width = 1125 Then
'        SSPanel5.Width = 100
'    Else
'        SSPanel5.Width = 1125
'    End If
'End Sub

Private Sub subupdate_Click()
    If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    frmupdate.Show 1
End Sub

Private Sub tigaA_Click()
    M_OBJCONN.Execute "DELETE FROM tblnotif_info WHERE type_notif='address' "
    FrmListReqTlp.Show
End Sub

Private Sub Timer1_Timer()
    'Dim ConMSG As New ADODB.Connection
    Dim cmdsqlnew As String
    Dim cmdsql3 As String
    Dim M_objrs As New ADODB.Recordset
    Dim CMDSQL As String

    'On Error GoTo SALAH

    'ConMSG.Open CMDSQLOPEN
    M_objrs.CursorLocation = adUseClient
    cmdsql3 = "select sender, recipient, datetime, msg, t from msgtbl where recipient ='" + Trim(MDIForm1.TxtUsername.text) + "' and sts ='0'"
    M_objrs.Open cmdsql3, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

    If M_objrs.RecordCount = 0 Then
        Set M_objrs = Nothing
        Exit Sub
    End If

    While Not M_objrs.EOF
        FRMTERIMAMSG.RichTextBox1.SelColor = &HC00000
        FRMTERIMAMSG.Text1.text = IIf(IsNull(M_objrs!Sender), "", M_objrs!Sender)
        FRMTERIMAMSG.RichTextBox1.text = FRMTERIMAMSG.RichTextBox1.text + "Dari :" + IIf(IsNull(M_objrs!Sender), "", M_objrs!Sender) + vbCrLf
        FRMTERIMAMSG.RichTextBox1.text = FRMTERIMAMSG.RichTextBox1.text + "Kepada :" + IIf(IsNull(M_objrs!RECIPIENT), "", M_objrs!RECIPIENT) + vbCrLf
        FRMTERIMAMSG.RichTextBox1.text = FRMTERIMAMSG.RichTextBox1.text + "Tanggal :" + IIf(IsNull(M_objrs!DateTime), "", M_objrs!DateTime) + vbCrLf
        FRMTERIMAMSG.RichTextBox1.text = FRMTERIMAMSG.RichTextBox1.text + "Isi Pesan :" + vbCrLf
        FRMTERIMAMSG.RichTextBox1.text = FRMTERIMAMSG.RichTextBox1.text + IIf(IsNull(M_objrs!msg), "", M_objrs!msg)
        FRMTERIMAMSG.RichTextBox1.text = FRMTERIMAMSG.RichTextBox1.text + " " + vbCrLf
        FRMTERIMAMSG.RichTextBox1.text = FRMTERIMAMSG.RichTextBox1.text & vbCrLf
        M_objrs.MoveNext
    Wend
    If M_objrs.RecordCount <> 0 Then
        On Error GoTo Salah
        'FRMTERIMAMSG.Show vbModal
        FRMTERIMAMSG.Show vbModal
        cmdsql3 = "UPDATE msgtbl SET STS ='1' WHERE RECIPIENT ='" + MDIForm1.TxtUsername.text + "'"
        M_OBJCONN.Execute cmdsql3
    End If
    Set M_objrs = Nothing
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    cmdsqlnew = "select * from usertbl where userid='" + MDIForm1.TxtUsername.text + "' and f_flagrender=1"
    M_objrs.Open cmdsqlnew, M_OBJCONN, adOpenDynamic, adLockOptimistic

    If M_objrs.RecordCount <> 0 Then
        bRenderrecord = True
        cmdsql3 = "UPDATE usertbl SET f_flagrender =0 where userid='" + MDIForm1.TxtUsername.text + "'"
        M_OBJCONN.Execute cmdsql3
    End If

    'ConMSG.Close
    'Set ConMSG = Nothing
    Exit Sub
Salah:
    FRMTERIMAMSG.Hide
    MsgBox "Ada error : " & err.Description
'
'
'    'jejaktian===========================================================================
'    ConMSG.Open CMDSQLOPEN
'    M_objrs.CursorLocation = adUseClient
'    cmdsql3 = "select pemohon from tblpermohonantransferdata where penggaprove = ''"
'    M_objrs.Open cmdsql3, ConMSG, adOpenDynamic, adLockOptimistic, adCmdText
'
'    If M_objrs.RecordCount = 0 Then
'        Set M_objrs = Nothing
'        Exit Sub
'    End If
'
'    While Not M_objrs.EOF
'        Formapptransferdata.lblpemohon.Caption = IIf(IsNull(M_objrs!pemohon), "", M_objrs!pemohon)
'        M_objrs.MoveNext
'    Wend
'    If M_objrs.RecordCount <> 0 Then
'        On Error GoTo SALAH
'        'FRMTERIMAMSG.Show vbModal
'        Formapptransferdata.Show vbModal
'        cmdsql3 = "UPDATE tblpermohonantransferdata SET penggaprove = '" + MDIForm1.txtusername.Text + "' WHERE pemohon is null"
'        ConMSG.Execute cmdsql3
'    End If

'    ConMSG.Close
'    Set ConMSG = Nothing
'    Exit Sub
'SALAH:
'    Formapptransferdata.Hide
'    MsgBox "Ada error : " & err.Description
End Sub

'Private Sub Timer2_Timer()
'    If UCase(mdiform1.txtlevel.text) <> "AGENT" Then
'       i_monitoring_activity = 0
'       Timer2.Enabled = False
'       Exit Sub
'    End If
'
'    FrmCC_Colection.Label12.Caption = i_monitoring_activity
'    i_monitoring_activity = i_monitoring_activity + 1
'
'    If i_monitoring_activity > 180 Then
'        If UCase(mdiform1.txtlevel.text) = "AGENT" Then
'            M_OBJCONN.Execute "UPDATE usertbl SET f_blok='1' WHERE userid='" & Trim(mdiform1.txtusername.text.text) & "'"
'            MsgBox "Akun anda di blok, karena tidak melakukan aktivitas selama lebih dari 3 menit. oleh SPV/Admin! Anda tidak dapat membuka aplikasi TINS! Konfirmasikan ke SPV/Admin untuk membuka blok aplikasi TINS anda!", vbOKOnly + vbCritical, "Peringatan"
'            End
'        End If
'    End If
'End Sub

'Private Sub Timer2_Timer()
'
'    'WskCTI_DataArrival (FEEDBACKprogressing)
'
'    Dim dura_blok As Integer
'
'    If UCase(MDIForm1.txtlevel.Text) <> "AGENT" Then
'       i_monitoring_activity = 0
'       Timer2.Enabled = False
'       Exit Sub
'    End If
'
''    TxtStatus.Text = "FEEDBACKhangup"
''
'    If TxtStatus.Text Like "*FEEDBACKhangup*" Then
'        i_monitoring_activity = 0
'        TxtStatus.Text = ""
'        Call logwktcti("RESET WAKTU BY TIAN")
'    End If
'
'    'If Not TxtStatus.Text Like "*FEEDBACKhangup*" Then
'        waktu_selesai_ngitung = waktu_server_sekarang
'    'End If
'    'TxtStatus.Text = "FEEDBACKbusy"
'
'    'Call WskCTI_DataArrival(FEEDBACKbusy)
'
''    If TxtStatus.Text = "FEEDBACKhangup" Then
''        FrmCC_Colection.Label12.Caption = 0
''        TxtStatus.Text = ""
''    End If
'
'    'dura_blok = DateDiff("s", waktu_mulai_ngitung, waktu_selesai_ngitung)
'
'    i_monitoring_activity = i_monitoring_activity + 1
'    'FrmCC_Colection.Label12.Caption = dura_blok
'    FrmCC_Colection.Label12.Caption = i_monitoring_activity
'
'    'If TxtStatus.Text <> "FEEDBACKbusy" Then
'        If FrmCC_Colection.Label12.Caption > 180 Then
'            If UCase(MDIForm1.txtlevel.Text) = "AGENT" Then
'                M_OBJCONN.Execute "UPDATE usertbl SET f_blok='1' WHERE userid='" & Trim(MDIForm1.TxtUsername.Text) & "'"
'                MsgBox "Akun anda di blok, karena tidak melakukan aktivitas selama lebih dari 3 menit. oleh SPV/Admin! Anda tidak dapat membuka aplikasi TINS! Konfirmasikan ke SPV/Admin untuk membuka blok aplikasi TINS anda!", vbOKOnly + vbCritical, "Peringatan"
'                Call logwktcti("terblok timer2")
'                Call set_count_ol
'                End
'            End If
'        End If
'    'End If
'End Sub
'
'Sub KEDAPKEDIP()
'    If Label4.Visible = True Then
'        Label4.Visible = False
'    ElseIf Label4.Visible = False Then
'        Label4.Visible = True
'    End If
'End Sub
'
'Private Sub Timer3_Timer()
'    Call KEDAPKEDIP
'    'Label3.Caption = Now()
'    Dim M_DATA As New CLS_FRMCUST_CC_MGM
'    Dim cmdsql As String
'    Dim n As String
'    Dim tglnow As Date
'    'tglnow=format(
'    'Label4.Caption = Now()
'    n = Format(Now(), "hh:mm:ss")
'
'    If n = "22:00:00" Then
'        Label3.Caption = "haiii"
'
'        'Otomatis BP
'        cmdsql = "update mgm SET LASTSTATUS=KETHSLKERJA,KETHSLKERJA='BP-BROKEN PROMISE',F_CEK='BP-',REMARKS = 'BP-BROKEN PROMISE-Auto',RECSTATUS='C',OTO='Y',TGLSTATUS='" & Format(Now, "yyyy/mm/dd") & "'"
'        cmdsql = cmdsql + "where custid in (select custid from vwptp1 "
'        cmdsql = cmdsql + "where datediff(day,promisedate,getdate())>7 and custid not in ( "
'        cmdsql = cmdsql + "select distinct custid from tbllunas)) And F_CEK like '%PTP%'"
'        M_OBJCONN.Execute cmdsql
'
'        'Otomatis POP
'        cmdsql = "update mgm SET"
'        cmdsql = cmdsql + " LASTSTATUS=KETHSLKErJA,KETHSLKErJA='POP-PROGRESS OF PAYMENT',F_CEK='POP',rEMArKS = 'POP-PROGRESS OF PAYMENT-Auto',RECSTATUS='C',OTO='Y',TGLSTATUS='" & Format(Now, "yyyy/mm/dd") & "'"
'        cmdsql = cmdsql + " where custid in ("
'        cmdsql = cmdsql + " select distinct custid from tbllunas)"
'        cmdsql = cmdsql + " And F_CEK<>'POP' AND F_CEK='PTP'"
'        M_OBJCONN.Execute cmdsql
'
'        '
'        cmdsql = "SELECT CUSTID,AGENT,REMARKS,NEXTACT,F_CEK,Statuscall,MOBILENO from mgm where OTO='Y'"
'        Dim ds As ADODB.Recordset
'        Set ds = New ADODB.Recordset
'        ds.CursorLocation = adUseClient
'        ds.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'        If ds.EOF And ds.BOF Then
'        Else
'            Do While Not ds.EOF
'                M_DATA.ADD_HISTORY_OTO M_OBJCONN, ds!CustId, Format(Now, "yyyy/mm/dd hh:mm:dd"), Time, ds!agent, "COLLECTION", IIf(IsNull(ds!Remarks), "", ds!Remarks), "", IIf(IsNull(ds!NEXTACT), "", ds!NEXTACT), "", IIf(IsNull(ds!F_CEK), "", ds!F_CEK), IIf(IsNull(ds!statuscall), "", ds!statuscall), IIf(IsNull(ds!MOBILENO), "", ds!MOBILENO)
'                ds.MoveNext
'            Loop
'        End If
'
'        cmdsql = "update mgm SET OTO=''"
'        M_OBJCONN.Execute cmdsql
'    End If
'End Sub
'
'Private Sub Timer4_Timer()
'    Dim tglserver As String
'    Dim TGLCLICK As String
'    Dim listitem As listitem
'    Dim ConnPTP As New ADODB.Connection
'    Dim M_objrs As New ADODB.Recordset
'    Dim cmdsql3 As String
'
'    If shedulePTP = True Then
'        'ngak ada kegiatan
'    Else
'    ConnPTP.Open CMDSQLOPEN
'    Set M_objrs = New ADODB.Recordset
'    M_objrs.CursorLocation = adUseClient
'    cmdsql3 = "select custid,name,tdbDatePTP from mgm where TdbDatePTP = '" + Format((Now + 7), "yyyy/mm/dd") + "' and agent ='" + MDIForm1.TxtUsername.Text + "'"
'    'cmdsql3 = "select * from mgm where TGLINCOMING = '" + Format((MDIForm1.TDBDate1.Value + 7), "yyyy/mm/dd") + "'"
'    M_objrs.Open cmdsql3, ConnPTP, adOpenDynamic, adLockOptimistic, adCmdText
'    If M_objrs.RecordCount <> 0 Then
'        LstGrade.ListItems.CLEAR
'    End If
'    While Not M_objrs.EOF
'        Set listitem = LstGrade.ListItems.ADD(, , M_objrs.Bookmark)
'        listitem.SubItems(1) = IIf(IsNull(M_objrs("CUSTID")), "", M_objrs("CUSTID"))
'        listitem.SubItems(2) = IIf(IsNull(M_objrs("NAME")), "", M_objrs("NAME"))
'        listitem.SubItems(3) = Format(IIf(IsNull(M_objrs("TdbDatePTP")), "", M_objrs("TdbDatePTP")), "yyyy/mm/dd hh:nn")
'        M_objrs.MoveNext
'    Wend
'    Set M_objrs = Nothing
'    If M_objrs.RecordCount <> 0 Then
'        MsgBox "You Got Schedule PTP to Follow Up", vbInformation + vbOKOnly, "Aplikasi"
'        shedulePTP = True
'    End If
'
''    Set m_objrs = New ADODB.Recordset
''    m_objrs.CursorLocation = adUseClient
''    cmdsql3 = "select custid, name, TdbDatePTP from mgm where TdbDatePTP = '" + Format((MDIForm1.TDBDate1.Value + 1), "yyyy/mm/dd") + "' and agent ='" + mdiform1.txtusername.text.text + "'"
''    'cmdsql3 = "select * from mgm where TGLINCOMING = '" + Format((MDIForm1.TDBDate1.Value + 1), "yyyy/mm/dd") + "'"
''    m_objrs.Open cmdsql3, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
''    If m_objrs.RecordCount <> 0 Then
''    '    LstGrade.ListItems.Clear
''    End If
''    While Not m_objrs.EOF
''        Set listitem = LstGrade.ListItems.ADD(, , m_objrs.Bookmark)
''        listitem.SubItems(1) = IIf(IsNull(m_objrs("CUSTID")), "", m_objrs("CUSTID"))
''        listitem.SubItems(2) = IIf(IsNull(m_objrs("NAME")), "", m_objrs("NAME"))
''        listitem.SubItems(3) = Format(IIf(IsNull(m_objrs("TdbDatePTP")), "", m_objrs("TdbDatePTP")), "yyyy/mm/dd hh:nn")
''        m_objrs.MoveNext
''    Wend
''    If m_objrs.RecordCount <> 0 Then
''        MsgBox "You Got Schedule PTP to Follow Up", vbInformation + vbOKOnly, "Aplikasi"
''        shedulePTP = True
''    End If
''End If
''Set m_objrs = Nothing
'
'    Set M_objrs = New ADODB.Recordset
'    M_objrs.CursorLocation = adUseClient
'    cmdsql3 = "select CUSTID, NAME, NEXTACTDATE from mgm where NEXTACTDATE BETWEEN '" + Format((Now), "yyyy/mm/dd") & " 00:00" + "' and '" + Format((Now), "yyyy/mm/dd") & " 23:59" + "' and agent ='" + MDIForm1.TxtUsername.Text + "'"
'    'cmdsql3 = "select * from mgm where NEXTACTDATE BETWEEN '" + Format((MDIForm1.TDBDate1.Value), "yyyy/mm/dd") & " 00:00" + "' and '" + Format((MDIForm1.TDBDate1.Value), "yyyy/mm/dd") & " 23:59" + "'"
'    M_objrs.Open cmdsql3, ConnPTP, adOpenDynamic, adLockOptimistic, adCmdText
'    If M_objrs.RecordCount <> 0 Then
'        LstGrade.ListItems.CLEAR
'    End If
'    While Not M_objrs.EOF
'        Set listitem = LstGrade.ListItems.ADD(, , M_objrs.Bookmark)
'        listitem.SubItems(1) = IIf(IsNull(M_objrs("CUSTID")), "", M_objrs("CUSTID"))
'        listitem.SubItems(2) = IIf(IsNull(M_objrs("NAME")), "", M_objrs("NAME"))
'        listitem.SubItems(3) = Format(IIf(IsNull(M_objrs("NEXTACTDATE")), "", M_objrs("NEXTACTDATE")), "yyyy/mm/dd hh:nn")
'        M_objrs.MoveNext
'    Wend
'    Set M_objrs = Nothing
'    ConnPTP.Close
'    Set ConnPTP = Nothing
'    End If
'End Sub

Public Sub ActionCTI(nilai As String)
    'ParameterCTI = Nilai & " /r/n"
    ParameterCTI = nilai
    TimerCTI.Enabled = True
End Sub


Private Sub Timer2_Timer()
End Sub

Private Sub Timer100_Timer()
    'belum pernah dijalankan
    On Error GoTo hell
    If UCase(MDIForm1.txtlevel.text) = "AGENT" Then
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        CMDSQL1 = "select spvcode from usertbl where userid = '" & MDIForm1.TxtUsername.text & "'"
        M_objrs.Open CMDSQL1, M_OBJCONN, adOpenDynamic, adLockOptimistic
        spvnye = LCase(M_objrs!SPVCODE)
        Set M_objrs = Nothing

        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        CMDSQL1 = "select * from tbl_lock_log_" & spvnye & " where agent = '" & MDIForm1.TxtUsername.text & "' and jam_awal < now() and jam_akhir > now() and coalesce(stsopenagent,0) <> 1"
        M_objrs.Open CMDSQL1, M_OBJCONN, adOpenDynamic, adLockOptimistic

        If M_objrs.RecordCount > 0 Then
            qupd = "Update tbl_lock_log_" & spvnye & " set stsopenagent = 1 where agent = '" & MDIForm1.TxtUsername.text & "' and jam_awal < now() and jam_akhir > now();"
            M_OBJCONN.Execute qupd
            MsgBox "Sedang dilakukan Lock Data"
            On Error GoTo z
atas:
                VIEW_MGMDATA.SetFocus
                VIEW_MGMDATA.txtnocard.text = ""
                VIEW_MGMDATA.Text1(0).text = ""
                VIEW_MGMDATA.Combo2.text = ""
                VIEW_MGMDATA.Combo1(2).text = ""
                VIEW_MGMDATA.txtregion.text = ""
                VIEW_MGMDATA.txtamount.text = ""
                VIEW_MGMDATA.txtcurbalance.text = ""
                VIEW_MGMDATA.Command1(0).SetFocus
                Sendkeys "{ENTER}"
                GoTo bawah:
z:
            SSCommand1_Click (0)
            GoTo atas:
bawah:
        End If

        Set M_objrs = Nothing

    End If
hell:
End Sub

Private Sub Timer6_Timer()
    Dim M_objrs     As ADODB.Recordset
    Dim CMDSQL      As String
    Dim query       As String
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    CMDSQL = "SELECT * FROM tbl_notif_sms WHERE agent='" & TxtUsername.text & "'"
    
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_objrs.RecordCount <> 0 Then
        ssql = " UPDATE inbox SET `SenderNumber` = replace(`SenderNumber`,'+62','0')"
        M_OBJCONN1.Execute (ssql)
        
        frm_unread_msg.Show 1
    End If

    Set M_objrs = Nothing
    'Call sms_muncul
    If MDIForm1.txtlevel.text = "Agent" Then
        
        '---------reminder----------------------------------------------------
        Dim str_detik As String
        Dim str_group_time As Integer
        Dim LocTextFile, CallerIDfromTextFile As String
        Dim str_time, isi As String
        Dim arr_reminder() As String
        Dim reminder_custid As String
        Dim reminder_jam As String
        Dim reminder_custname As String
                
                On Error GoTo err
                
                LocTextFile = "C:\reminder.txt"
            
                SqlWaktu = "select now()"
                Set m_waktuserver = New ADODB.Recordset
                m_waktuserver.CursorLocation = adUseClient
                m_waktuserver.Open SqlWaktu, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                str_time = Format(m_waktuserver(0), "HH:MM")
                
                Set m_waktuserver = Nothing
                
                Open LocTextFile For Input As #1    'Buka file text
                Do Until EOF(1)
                    Line Input #1, CallerIDfromTextFile      'Baca Baris Pertama
                    isi = Replace(CallerIDfromTextFile, """", "")
                    arr_reminder = Split(isi, "|")
                    reminder_custid = arr_reminder(0)
                    reminder_custname = arr_reminder(1)
                    reminder_jam = arr_reminder(2)
                    If reminder_jam = str_time Then
                        With frm_reminder
                            .txtcustid.text = reminder_custid
                            .txtnama.text = reminder_custname
                            .Label1(6).Caption = Format(Now, "DD/MM/YYYY") & " - " & reminder_jam
                            .ZOrder 0
                            .Show vbModal
                        End With
                    End If
                Loop
err:
                Close #1 'Tutup File file text
    End If
    
    If MDIForm1.txtlevel.text = "Admin" Or MDIForm1.txtlevel.text = "Supervisor" Then
        query = "Select type_notif,count(type_notif) as jml_notif from tblnotif_info GROUP BY 1"
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open query, M_OBJCONN, adOpenDynamic, adLockOptimistic
        lblsms_unread.Visible = False
        lblapp.Visible = False
        If M_objrs.RecordCount <> 0 Then
            While Not M_objrs.EOF
                If M_objrs!type_notif = "sms" Then
                    lblsms_unread.Caption = "NEW MESSAGE (" & M_objrs!jml_notif & ")"
                    lblsms_unread.Visible = True
                Else
                    lblsms_unread.Caption = "NEW APPROVAL NUMBER (" & M_objrs!jml_notif & ")"
                    lblapp.Visible = True
                End If
                M_objrs.MoveNext
            Wend
        End If
    End If
End Sub

Private Sub TimerCTI_Timer()
    Select Case WskCTI.State
        Case 9, 8, 1
            Debug.Print WskCTI.State
            WskCTI.Close
            WskCTI.RemoteHost = "127.0.0.1"
            'buat connect ke chromium
            'WskCTI.RemotePort = 2121
            'buat connect ke cti
            WskCTI.RemotePort = 18000
            WskCTI.Connect
        Case 6
            Debug.Print WskCTI.State
        Case 7
            Debug.Print WskCTI.State
            If Len(ParameterCTI) > 2 Then
                WskCTI.SendData ParameterCTI + vbCrLf
                'MsgBox ParameterCTI
            End If
            Debug.Print ParameterCTI
            TimerCTI.Enabled = False
        Case 0
            Debug.Print WskCTI.State
            WskCTI.RemoteHost = "127.0.0.1"
            'buat connect ke chromium
            'WskCTI.RemotePort = 2121
            'buat connect ke cti
            WskCTI.RemotePort = 18000
            WskCTI.Connect
        Case Else
    End Select
End Sub

Private Sub TimerRequest_Timer()
'    If UCase(MDIForm1.txtlevel.Text) = "TEAMLEADER" Or _
'       UCase(MDIForm1.txtlevel.Text) = "SUPERVISOR" Or _
'       UCase(MDIForm1.txtlevel.Text) = "ADMIN" Or _
'       UCase(MDIForm1.txtlevel.Text) = "ADMINISTRATOR" Then
'       Call CekReqNumber
'    Else
'        TimerRequest.Enabled = False
'    End If
End Sub

Private Sub CekReqNumber()
'@@11092012 - DinonAktifkan
'    Dim CMDSQL As String
'    Dim M_OBJRS As ADODB.Recordset
'
'    'cek status f_req_number
'    CMDSQL = "select * from usertbl where userid='"
'    CMDSQL = CMDSQL + mdiform1.txtusername.text.text + "'"
'    Set M_OBJRS = New ADODB.Recordset
'    M_OBJRS.CursorLocation = adUseClient
'    M_OBJRS.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'    If M_OBJRS.RecordCount > 0 Then
'        If M_OBJRS("f_req_number") = "1" Then
'            On Error GoTo salah
'            FrmPemberitahuan.Show vbModal
'        End If
'    End If
'
'    Set M_OBJRS = Nothing
'    Exit Sub
'salah:
'    FrmPemberitahuan.Hide
'    MsgBox "Ada error :" & Err.Description
End Sub


Private Sub TimerTandaReq_Timer()
'    If ShapeReq.FillColor = vbBlack Then
'        ShapeReq.FillColor = vbRed
'        KelapKelip = KelapKelip + 1
'    Else
'        ShapeReq.FillColor = vbBlack
'        KelapKelip = KelapKelip + 1
'    End If
'
'    If KelapKelip = 7 Then
'        KelapKelip = 0
'        WaitSecs (3)
'        ShapeReq.FillColor = vbBlack
'        TimerTandaReq.Enabled = False
'    End If
End Sub

Private Sub transfer_data_Click()
    formtransferdata.Show vbModal
End Sub

Private Sub upload_fresh_wo_Click()
    Form_upload_fresh_wo.Show vbModal
End Sub

Private Sub VSMS_Click()
    If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbCritical, "Peringatan"
        Exit Sub
    End If
    load Frm_verify
    Frm_verify.Show vbModal
End Sub

Private Sub Winsock2_Error(ByVal number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

End Sub

Private Sub WskCTI_DataArrival(ByVal bytesTotal As Long)
    Dim StrMsgDrCti As String
    Dim lStrMsgDrCti As String
    Dim arr_str() As String
    Dim str_unique() As String
    'get data balik dari cti
    'StrMsgDrCti = "FEEDBACKhangup"
    WskCTI.GetData StrMsgDrCti, vbString
    Debug.Print StrMsgDrCti
    If Len(StrMsgDrCti) > 1 Then
        'FrmSoftPhone.Caption = StrMsgDrCti
        lStrMsgDrCti = Left(StrMsgDrCti, Len(StrMsgDrCti) - 2)
        Txtstatus.text = StrMsgDrCti
        If InStr(lStrMsgDrCti, "FEEDBACKinitiated") Then
            arr_str = Split(lStrMsgDrCti, "|")
            str_unique = Split(arr_str(2), vbCrLf)
            txt_unique_id.text = str_unique(0)
            txtChannel.text = arr_str(1)
        ElseIf InStr(lStrMsgDrCti, "FEEDBACKhangup") Then '<-- jika tidak get wktu hangup, ganti dgn yg dibawah
            arr_str = Split(lStrMsgDrCti, "|")
            ' -- OTOMATIS HANGUP JIKA BELUM DI HANGUP
            If MDIForm1.txtlevel.text = "Agent" Then
                If FrmCC_Colection.stshangup.text = 0 Then
                    FrmCC_Colection.hangup_event
                End If
            End If
            'txtChannel.Text = arr_str(1)
        End If
        
        If InStr(lStrMsgDrCti, "connected..") Then
            Obelisk = True
        End If
    
        If InStr(lStrMsgDrCti, "free") Then
                    'hang up
            '            Call savecall
            '            FBILL.Timer6.Enabled = False
            '            Unload FBILL
             '           Unload FrmAnswerCall
        End If
    End If
End Sub

Private Sub WskCTI_SendComplete()
    Debug.Print "OK send"
End Sub

Private Sub WskCTI_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    Debug.Print "Sending" & CStr(bytesSent) & "/" & CStr(bytesRemaining)
End Sub

'softphone control
Private Sub CmdACW_Click()
    MDIForm1.ActionCTI ("ACW")
End Sub

Private Sub CmdAUX_Click()
    MDIForm1.ActionCTI ("AUX")
End Sub

Private Sub CmdCall_Click()


    If CmbNo.text = "108" Then
        MDIForm1.ActionCTI ("DIAL|268" & GetNumber(108))
    Else
        If Len(CmbNo.text) >= 6 And Len(CmbNo.text) <= 7 Then
            If Left(CmbNo.text, 2) <> "08" And Right(CmbNo.text, 3) = "108" Then
                MDIForm1.ActionCTI ("DIAL|268" & GetNumber(CmbNo.text))
            End If
        End If
    End If
    
    If CmbNo.text = "109" Then
        MDIForm1.ActionCTI ("DIAL|268" & GetNumber(109))
    Else
        If Len(CmbNo.text) >= 6 And Len(CmbNo.text) <= 7 Then
            If Left(CmbNo.text, 2) <> "08" And Right(CmbNo.text, 3) = "109" Then
                MDIForm1.ActionCTI ("DIAL|268" & GetNumber(CmbNo.text))
            End If
        End If
    End If
    
    If CmbNo.text = "147" Then
        MDIForm1.ActionCTI ("DIAL|268" & GetNumber(147))
    Else
        If Len(CmbNo.text) >= 6 And Len(CmbNo.text) <= 7 Then
            If Left(CmbNo.text, 2) <> "08" And Right(CmbNo.text, 3) = "147" Then
                MDIForm1.ActionCTI ("DIAL|268" & GetNumber(CmbNo.text))
            End If
        End If
    End If
    
    If CmbNo.text = "109" Or CmbNo.text = "108" Or CmbNo.text = "147" Then
    Else
        MDIForm1.ActionCTI ("DIAL|" + CmbNo)
    End If
End Sub

Private Sub CmdConference_Click()
    MsgBox "Conference"
End Sub

Private Sub CmdHangUp_Click()
    MDIForm1.ActionCTI ("HANGUP")
End Sub

Private Sub CmdLogin_Click()
    MsgBox "Login"
End Sub

Private Sub Cmdready_Click()
    MDIForm1.ActionCTI ("READY")
End Sub

Private Sub CmdOutbound_Click()
    MDIForm1.ActionCTI ("NOTREADY")
End Sub

Private Sub CmdTransfer_Click()
    If CmbNo.text = "" Then
    Else
        MDIForm1.ActionCTI ("TRANSFER" + CmbNo)
    End If
End Sub

Private Sub CmdBintang_Click()
    CmbNo.text = CmbNo.text + "*"
End Sub

Private Sub CmdCancel_Click()
    CmbNo.text = ""
End Sub

Private Sub CmdNo_Click(Index As Integer)
    CmbNo.text = CmbNo.text + CStr(Index)
End Sub

Private Sub CmdPager_Click()
    CmbNo.text = CmbNo.text + "#"
End Sub


Private Sub Cmddtmf_Click()
    If CmbNo.text = "" Then
    Else
        MDIForm1.ActionCTI ("DTMF" + CmbNo)
    End If
End Sub

Private Sub PromiseToPay()
    LstGrade.ColumnHeaders.ADD 1, , "No", 3 * 120
    LstGrade.ColumnHeaders.ADD 2, , "Cust ID", 10 * 120
    LstGrade.ColumnHeaders.ADD 3, , "Status", 10 * 120
    LstGrade.ColumnHeaders.ADD 4, , "Agent", 10 * 120
End Sub

Private Sub HeaderInformation()
    LstInformation.ColumnHeaders.ADD 1, , "Description", 20 * 120
    LstInformation.ColumnHeaders.ADD 2, , "No", 1
    LstInformation.ColumnHeaders.ADD 3, , "Lokasi", 1
End Sub

Private Sub CmdAccept_Click()
    MDIForm1.ActionCTI ("DTMF")
End Sub

Private Sub CmdOutgoing_Click()
    MDIForm1.ActionCTI ("OUTGOING")
End Sub

Private Sub LstInformation_DblClick()
    If LstInformation.ListItems.Count = 0 Then
        Exit Sub
    End If
    If StartMeUp(LstInformation.SelectedItem.SubItems(2)) <= 32 Then
       MsgBox "File Tidak Ditemukan", vbOKOnly + vbCritical, "Pemberitahuan"
       
    Else
    SSTab1.Tab = 0
    End If
End Sub

Private Sub LstDataInformation()
    Dim ListItem As ListItem
    Dim ssql As String
    Set M_LOGINRS = New ADODB.Recordset
    M_LOGINRS.CursorLocation = adUseClient

    ssql = "SELECT ExpiryDate, Description, id, Direktori FROM tblinformationlokasi " & _
           "ORDER BY Description"
    M_LOGINRS.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText

    While Not M_LOGINRS.EOF
    If Format(M_LOGINRS!ExpiryDate, "yyyy/mm/dd") > Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") Then
        Set ListItem = MDIForm1.LstInformation.ListItems.ADD(, , IIf(IsNull(M_LOGINRS("Description")), "", M_LOGINRS("Description")))
            ListItem.SubItems(1) = IIf(IsNull(M_LOGINRS("id")), "", M_LOGINRS("id"))
            ListItem.SubItems(2) = IIf(IsNull(M_LOGINRS("Direktori")), "", M_LOGINRS("Direktori"))
    End If
        M_LOGINRS.MoveNext
    Wend

    Set M_LOGINRS = Nothing
End Sub

Function ReplaceFirstInstance(SourceString, _
    Searchstring, Replacestring)
    Dim StartLoc
    Dim FoundLoc
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
    Dim tmpString1
    Dim tmpString2
    tmpString1 = SourceString
 
    tmpString2 = tmpString1
    tmpString1 = ReplaceFirstInstance(tmpString1, _
                 Searchstring, Replacestring)
    
    FindReplace = tmpString1
End Function

'@@ 15-12-2010 buat timer lock data
Private Sub Timer_stopwatch_Timer()
'
'    'Tambah dengan satu untuk total sepersepuluh detik.
'    'Kita mengeset interval Timer menjadi 10, jadi
'    'setiap sepersepuluh detik prosedur ini akan
'    'dieksekusi
'    TotalTenthDetik = TotalTenthDetik + 1
'    'Jika TotalTenthSeconds = 10,
'    'set kembali menjadi 0.
'    TenthDetik = TotalTenthDetik Mod 10
'    '10 kali sepersepuluh detik sama dengan 1 detik.
'    'int - akan mengembalikan bilangan integer (bulat)
'    'dari pecahan 'Contoh: Int(0.9) = 0 menghasilkan 0
'    TotalDetik = Int(TotalTenthDetik / 10)
'    'Jika variabel Seconds = 60, set kembali menjadi 0
'    Detik = TotalDetik Mod 60
'    If Len(Detik) = 1 Then
'       Detik = "0" & Detik  'Agar selalu dalam dua
'                            'digit
'    End If
'    Menit = Int(TotalDetik / 60) Mod 60
'    If Len(Menit) = 1 Then
'       Menit = "0" & Menit    'Agar selalu dalam dua
'                          'digit
'    End If
'    JAM = Int(TotalDetik / 3600)
'    If JAM < 9 Then
'       Jam1 = "0" & JAM       'Agar selalu dalam dua'digit
'    End If
'    'Tampilkan hasilnya di Lblwaktu (update terus Lblwaktu)
'    LblWaktu.Caption = Jam1 & ":" & Menit & ":" & Detik & ":" & TenthDetik & ""
'
'    If LblWaktu.Caption = TxtWaktuRefresh.Text & ":0" Then
'        DoEvents
'
'        If Format(Now, "hh:mm:ss") > CDate(#9:00:00 PM#) Then
'            GoTo selanjutnya
'        End If
'
'        If UCase(txtlevel.Text) = "TEAMLEADER" Then
'             'MsgBox "Ok"
'             Call LockDataAuto
'        End If
'
'selanjutnya:
'        'Memulai atau menghentikan timer kembali
'         'Timer_stopwatch.Enabled = Not Timer1.Enabled
'         'Inisialisasi total sepersepuluh detik
'         TotalTenthDetik = -1
'         'Aktifkan timer
'         Timer_stopwatch.Enabled = True
'
'    End If


End Sub

Private Sub LockDataAuto()
        '@@ Awal 061110 cek lock account sesuai settingan timer
        Dim m_objrsTemp As ADODB.Recordset
        Dim M_ObjrsWaktuServer As ADODB.Recordset
        Dim m_objrsCurrent As ADODB.Recordset
        
        
        Dim cmdsqlserver As String
        Dim WaktuServer As Date
        Dim WaktuAkhirCurrent As Date
        
        'ambil waktu server
        cmdsqlserver = "select now() as WaktuServer "
        Set M_ObjrsWaktuServer = New ADODB.Recordset
        M_ObjrsWaktuServer.CursorLocation = adUseClient
    
        M_ObjrsWaktuServer.Open cmdsqlserver, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        WaktuServer = Format(M_ObjrsWaktuServer(0), "mm-dd-yyyy hh:mm")
        Set M_ObjrsWaktuServer = Nothing
        
        'Cek lock account yang sedang berjalan
        cmdsqlserver = "select * from tbltemplockacc_current "
        Set m_objrsCurrent = New ADODB.Recordset
        m_objrsCurrent.CursorLocation = adUseClient
        m_objrsCurrent.Open cmdsqlserver, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        If m_objrsCurrent.RecordCount <> 0 Then
            WaktuAkhirCurrent = Format(m_objrsCurrent("end_lock"), "mm-dd-yyyy hh:mm")
        Else
            GoTo lockdata
        End If
        
        While Not m_objrsCurrent.EOF
            
            WaktuAkhirCurrent = Format(m_objrsCurrent("end_lock"), "mm-dd-yyyy hh:mm")
        
            If WaktuAkhirCurrent <= WaktuServer Then
                'Cek dulu apakah ada user yang sedang mereset data
                If Trim(m_objrsCurrent("f_locked")) = "2" Then
                    GoTo KeluarLockAutoTL
                End If
                
                'update dulu status lock yang sedang berakhir, supaya agent lain ga ikut ngereset
                cmdsqlserver = "update tbltemplockacc_current set f_locked='2' where id='"
                cmdsqlserver = cmdsqlserver + Trim(m_objrsCurrent("id")) + "'"
                M_OBJCONN.Execute cmdsqlserver
            
                'Clear lock data yang sedang berjalan sesuai dengan agent yang di lock
                cmdsqlserver = "update usertbl set dilockoleh='ClearByAutomatic',"
                cmdsqlserver = cmdsqlserver + "lockdarispv=null,lock_entry_lpd=null,fromaccount=null,"
                cmdsqlserver = cmdsqlserver + "lockmarkup=null,lockdarispvbuattl=null,f_pesanlockauto=null,f_idsessstart=null,f_pesanresetauto='1',f_idsessend='"
                cmdsqlserver = cmdsqlserver + Trim(m_objrsCurrent("id")) + "' "
                'Buat ambil kondisi agent yang sedang di lock
                If Trim(m_objrsCurrent("account_lock")) = "ALL" Then
                    cmdsqlserver = cmdsqlserver + " where usertype='1' "
                ElseIf Left(Trim(m_objrsCurrent("account_lock")), 3) = "SPV" Then
                    cmdsqlserver = cmdsqlserver + " where spvcode='"
                    cmdsqlserver = cmdsqlserver + Trim(m_objrsCurrent("account_lock")) + "' "
                Else
                    cmdsqlserver = cmdsqlserver + " where userid='"
                    cmdsqlserver = cmdsqlserver + Trim(m_objrsCurrent("account_lock")) + "' "
                End If
'                cmdsqlserver = cmdsqlserver + " and f_idsessstart='"
'                cmdsqlserver = cmdsqlserver + Trim(m_objrsCurrent("id")) + "' "
                M_OBJCONN.Execute cmdsqlserver
                
                
                'Pindahkan data lock account current ke tabel data log tbltemplockacc_log
                cmdsqlserver = "insert into tbltemplockacc_log select * from tbltemplockacc_current "
                cmdsqlserver = cmdsqlserver + " where id='"
                cmdsqlserver = cmdsqlserver + Trim(m_objrsCurrent("id")) + "'"
                M_OBJCONN.Execute cmdsqlserver
                
                'Hapus data di tabel locktemp current
                cmdsqlserver = "delete from tbltemplockacc_current where id='"
                cmdsqlserver = cmdsqlserver + Trim(m_objrsCurrent("id")) + "'"
                M_OBJCONN.Execute cmdsqlserver
                
             End If
KeluarLockAutoTL:
                m_objrsCurrent.MoveNext
            Wend
            Set m_objrsCurrent = Nothing

            
       
        
        '=======
lockdata:
        'Setelah cek waktu lock yang habis, sekarang cek lock yg masih dalam antrian
        cmdsqlserver = "select * from tbltemplockacc where f_locked isnull order by start_lock asc "
        Set m_objrsTemp = New ADODB.Recordset
        m_objrsTemp.CursorLocation = adUseClient
        m_objrsTemp.Open cmdsqlserver, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            
            'Cek ada ga data lock dalam antrian
            If m_objrsTemp.RecordCount <> 0 Then
                Dim WaktuAwal As Date
                Dim WaktuAkhir As Date
                
                While Not m_objrsTemp.EOF
                
                    WaktuAwal = Format(m_objrsTemp("start_lock"), "mm-dd-yyyy hh:mm")
                    WaktuAkhir = Format(m_objrsTemp("end_lock"), "mm-dd-yyyy hh:mm")
                    
                    If (WaktuAwal <= WaktuServer) And (WaktuAkhir > WaktuServer) Then
                        'Cek apakah datanya sedang di lock sama agent lain?
                        If Trim(m_objrsTemp("f_locked")) = "1" Then
                            GoTo KeluarLockAutoTLLock
                        End If
                        
                        'update status  f_lockednya jadi 1, supaya ga di log sama agent lain
                        cmdsqlserver = "update tbltemplockacc set f_locked='1' where id='"
                        cmdsqlserver = cmdsqlserver + Trim(m_objrsTemp("id")) + "'"
                        M_OBJCONN.Execute cmdsqlserver
                        
                        'LAKUKAN LOCK DATA
                        Dim i As Integer
                       
                        a = Split(m_objrsTemp("script_lock"), "|")
                        
                        For i = LBound(a) + 1 To UBound(a) - 1
                            cmdsqlserver = Replace(a(i), "$", "'")
                            M_OBJCONN.Execute cmdsqlserver
                        Next i
                        
                        'Pindahin dulu data di tabel current ke tabel log, terus data di tabel current dihapus
'                        cmdsqlserver = "insert into tbltemplockacc_current "
'                        cmdsqlserver = cmdsqlserver + " select * from tbltemplockacc_log"
'                        M_OBJCONN.Execute cmdsqlserver --- Remarks dulu 10-11-10
                        
'                        cmdsqlserver = "delete from tbltemplockacc_current"
'                        M_OBJCONN.Execute cmdsqlserver --- Remarks dulu 10-11-10
                        
                        'Pindahin data dari tabel temp lock ke tabel current log
                        cmdsqlserver = "insert into tbltemplockacc_current "
                        cmdsqlserver = cmdsqlserver + "select * from tbltemplockacc where id='"
                        cmdsqlserver = cmdsqlserver + Trim(m_objrsTemp("id")) + "'"
                        M_OBJCONN.Execute cmdsqlserver
                        
                        
                        
                       'Update status pesan ke nilai 1,untuk menampilkan pesan ke agent
                        cmdsqlserver = "update usertbl set f_pesanlockauto='1',f_idsessstart='"
                        cmdsqlserver = cmdsqlserver + Trim(m_objrsTemp("id")) + "' "
                        'Buat mengupdate pesan kondisi agent yang di lock
                        If Trim(m_objrsTemp("account_lock")) = "ALL" Then
                            cmdsqlserver = cmdsqlserver + " where usertype='1' "
                        ElseIf Left(Trim(m_objrsTemp("account_lock")), 3) = "SPV" Then
                            cmdsqlserver = cmdsqlserver + " where spvcode='"
                            cmdsqlserver = cmdsqlserver + Trim(m_objrsTemp("account_lock")) + "'"
                        Else
                            cmdsqlserver = cmdsqlserver + " where userid='"
                            cmdsqlserver = cmdsqlserver + Trim(m_objrsTemp("account_lock")) + "'"
                        End If
                        M_OBJCONN.Execute cmdsqlserver
                        
                        'Hapus data di templock
                        cmdsqlserver = "delete from tbltemplockacc where id='"
                        cmdsqlserver = cmdsqlserver + Trim(m_objrsTemp("id")) + "'"
                        M_OBJCONN.Execute cmdsqlserver
                        
                        
                    End If
                   
KeluarLockAutoTLLock:
                    m_objrsTemp.MoveNext
               Wend

            End If
        
        Set m_objrsTemp = Nothing
      
      '@@ Akhir 061110 cek lock account sesuai settingan timer
End Sub

'@@ 14022011 ini buat cek sms
Private Sub CekSms()
    On Error Resume Next
    'Dim ConnPTP As New ADODB.Connection
    Dim M_objrs As New ADODB.Recordset
    Dim cmdsql34 As String
    Dim TELPo As String
    Dim codea As String
    
    If Left(Text1, 1) = "D" Or Text1 = "JOKO" Or Text1 = "SPV1" Or Left(Text1, 1) = "T" Then

        Select Case TxtUsername.text

            Case "TL1"
                codea = "ACC1"
            Case "TL2"
                codea = "ACC2"
            Case "TL3"
                codea = "ACC3"
            Case "TL4"
                codea = "ACC4"
            Case "TL5"
                codea = "ACC5"
            Case "TL6"
                codea = "ACC6"
            Case "TL7"
                codea = "ACC7"
            Case "TL8"
                codea = "ACC8"
            Case "TL9"
                codea = "ACC9"
            Case "TL10"
                codea = "ACC10"
            Case Else

                codea = TxtUsername.text

        End Select
    
        TELPo = "Select count(*) as banyak from inbox where sendernumber in ('a',"
    
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        cmdsql34 = "select mobileno,mobileno2,mobilenoadd1,mobilenoadd2 from mgm where agent = '" + codea + "'"
        M_objrs.Open cmdsql34, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
        If M_objrs.RecordCount = 0 Then
            Timer6.interval = 60000
            Exit Sub
        End If
    
        While Not M_objrs.EOF
            If Len(M_objrs("mobileno")) <> 0 Then
                satu = FindReplace(M_objrs("mobileno"), "0", "+62")
                TELPo = TELPo + "'" + satu + "',"
            Else
                TELPo = TELPo
            End If
    
            If Len(M_objrs("mobileno2")) <> 0 Then
                dua = FindReplace(M_objrs("mobileno2"), "0", "+62")
                TELPo = TELPo + "'" + dua + "',"
            Else
                TELPo = TELPo
            End If
    
            If Len(M_objrs("mobilenoadd1")) <> 0 Then
                tiga = FindReplace(M_objrs("mobilenoadd1"), "0", "+62")
                TELPo = TELPo + "'" + tiga + "',"
            Else
                TELPo = TELPo
            End If
            
            If Len(M_objrs("mobilenoadd2")) <> 0 Then
                empat = FindReplace(M_objrs("mobilenoadd2"), "0", "+62")
                TELPo = TELPo + "'" + empat + "',"
            Else
                TELPo = TELPo
            End If
        
            M_objrs.MoveNext
        Wend
        Set M_objrs = Nothing
    
        TELPo = Left(TELPo, Len(TELPo) - 1)
        Dim TELPo1
        Dim TELPo2
    
        TELPo1 = TELPo + ") and processed='f'"
        TELPo2 = TELPo + ") and processed='t'"
    
        M_objrs.Open TELPo1, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
        While Not M_objrs.EOF
            LblJmlSmsBaru.Caption = M_objrs("banyak")
            Label9 = "SMS BARU " & LblJmlSmsBaru.Caption & " SMS"
            M_objrs.MoveNext
        Wend
    
        'JIKA ADA SMS BARU MASUK
        If Trim(Label9.Caption) = "SMS BARU 0 SMS" Then
            'MsgBox "Tidak ada sms baru!"
            TimerBlink.Enabled = False
            Label9.ForeColor = vbBlack
        Else
            If Trim(Label9.Caption) <> "" Then
                TimerBlink.Enabled = True
                MsgBox "Ada SMS BARU MASUK! Silahkan cek!", vbOKOnly + vbInformation, "Informasi"
            End If
        End If
    
        Set M_objrs = Nothing
    
    '-----------------------------
        M_objrs.Open TELPo2, M_OBJCONN1, adOpenDynamic, adLockOptimistic, adCmdText
        While Not M_objrs.EOF
            Label10 = "SMS LAMA " & M_objrs("banyak") & " SMS"
            M_objrs.MoveNext
        Wend
        Set M_objrs = Nothing
    End If
    
    'MsgBox TELPo
     ' M_OBJCONN.Close
      'Set M_OBJCONN = Nothing
    Timer6.interval = 60000
End Sub

'@@06-04-2011, Tambahan jika sudah jam 11 Siang maka TL diingatkan untuk segera menarik report Contacto
'Private Sub TimerTanda_Timer()
'    If ShapeTanda.FillColor = vbBlack Then
'        ShapeTanda.FillColor = vbRed
'        KelapKelip = KelapKelip + 1
'    Else
'        ShapeTanda.FillColor = vbBlack
'        KelapKelip = KelapKelip + 1
'    End If
'
'    If KelapKelip = 7 Then
'        KelapKelip = 0
'        WaitSecs (3)
'        'TimerBlink.Enabled = False
'    End If
'End Sub

Private Sub TimerWaktu_Timer()
    '@@06-04-2011 Jika yang login Agent, matikan timer waktu
'    If UCase(Trim(MDIForm1.txtlevel.Text)) = "AGENT" Then
'        ShapeTanda.Visible = False
'        TimerWaktu.Enabled = False
'    End If
'    LblWaktu.Caption = Format(Now, "hh:mm:ss")
'    If UCase(Trim(mdiform1.txtlevel.text)) = "TEAMLEADER" Or _
'       UCase(Trim(mdiform1.txtlevel.text)) = "ADMIN" Or _
'       UCase(Trim(mdiform1.txtlevel.text)) = "ADMINISTRATOR" Then
'            If LblWaktu.Caption = "11:00:00" ThenXSELLbank
'                TimerTanda.Enabled = True
'                MsgBox "Sudahkah anda menarik report Productivity? Tekan Ok untuk melihatnya!", vbOKOnly + vbInformation, "Informasi"
'                Call IsiAgentContactto
'                Call IsiContactto
'                Call IsiContacttoJmlAcc
'                WaitSecs (2)
'                FrmMgmReport.RPT.Reset
'                FrmMgmReport.RPT.Formulas(1) = "@User = totext('" + CStr(mdiform1.txtusername.text.text) + "')"
'                FrmMgmReport.RPT.Formulas(2) = "@TglShow = totext('" + CStr(Format(Now, "dd-mm-yyyy") & " " & Format(Now, "hh:mm:ss")) + "')"
'                FrmMgmReport.RPT.Formulas(3) = "@TglShow1 = totext('" + CStr(Format(Now, "dd-mm-yyyy") & " " & Format(Now, "hh:mm:ss")) + "')"
'                FrmMgmReport.RPT.ReportFileName = "D:\COLLECTION_RITCARD\Report\RptContactto.rpt"
'                MDIForm1.TimerTanda.Enabled = False
'                MDIForm1.ShapeTanda.FillColor = vbBlack
'                Call SHOW_PRN
'            End If
'    End If
End Sub


'@@ 17-03-2011 Report Contactto
Private Sub IsiAgentContactto()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    
    
        CMDSQL = "select distinct u.spvcode as spv ,m.agent as agent"
        CMDSQL = CMDSQL + " from mgm as m, usertbl as u where "
        CMDSQL = CMDSQL + " m.agent=u.userid and u.userid in (select userid from usertbl "
        CMDSQL = CMDSQL + " where spvcode between '"
        CMDSQL = CMDSQL + "SPV1" + "' and '"
        CMDSQL = CMDSQL + "SPV9" + "' and usertype='1') and date(m.tglcall) between '"
        CMDSQL = CMDSQL + Format(Now, "yyyy-mm-dd") + "' and '"
        CMDSQL = CMDSQL + Format(Now, "yyyy-mm-dd") + "' "
        CMDSQL = CMDSQL + " group by m.agent,u.spvcode "
        CMDSQL = CMDSQL + "order by u.spvcode,m.agent asc"
    
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    M_RPTCONN.Execute "delete from TblRptContactto"
    If M_objrs.RecordCount > 0 Then
        'ProgressBar1.Max = M_OBJRS.RecordCount
        While Not M_objrs.EOF
            'ProgressBar1.Value = M_OBJRS.Bookmark
            CMDSQL = "insert into TblRptContactto (spvcode,agent) values ('"
            CMDSQL = CMDSQL + Trim(M_objrs("spv")) + "','"
            CMDSQL = CMDSQL + Trim(M_objrs("agent")) + "')"
             M_RPTCONN.Execute CMDSQL
            M_objrs.MoveNext
        Wend
    End If
    Set M_objrs = Nothing
End Sub

Private Sub IsiContactto()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    
    
'        CMDSQL = "select u.spvcode as spv ,m.agent as agent,m.stscallwith as status,count(m.stscallwith) as jumlah "
'        CMDSQL = CMDSQL + " from mgm as m, usertbl as u where "
'        CMDSQL = CMDSQL + " m.agent=u.userid and u.userid in (select userid from usertbl "
'        CMDSQL = CMDSQL + " where spvcode between '"
'        CMDSQL = CMDSQL + "SPV1" + "' and '"
'        CMDSQL = CMDSQL + "SPV9" + "' and usertype='1') and date(m.tglcall) between '"
'        CMDSQL = CMDSQL + Format(Now, "yyyy-mm-dd") + "' and '"
'        CMDSQL = CMDSQL + Format(Now, "yyyy-mm-dd") + "' "
'        CMDSQL = CMDSQL + " group by m.agent,m.stscallwith,u.spvcode "
'        CMDSQL = CMDSQL + "order by m.agent,m.stscallwith,u.spvcode asc"

        '@@01 Juni 2011 diubah querynya
        CMDSQL = "select u.spvcode as spv ,m.agent as agent,m.ststelpwith as status,count(m.ststelpwith) as jumlah "
        CMDSQL = CMDSQL + " from mgm_hst as m, usertbl as u where "
        CMDSQL = CMDSQL + " m.agent=u.userid and u.userid in (select userid from usertbl "
        CMDSQL = CMDSQL + " where spvcode between '"
        CMDSQL = CMDSQL + "SPV1" + "' and '"
        CMDSQL = CMDSQL + "SPV9" + "' and usertype='1') and date(m.tgl) between '"
'        CMDSQL = CMDSQL + Trim(Format(TDBDate1(0).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(0).Value) + "' and '"
'        CMDSQL = CMDSQL + Trim(Format(TDBDate1(1).Value, "yyyy-mm-dd")) & " " & Trim(DTimeLastCall(1).Value) + "' and "
        CMDSQL = CMDSQL + Format(Now, "yyyy-mm-dd") + "' and '"
        CMDSQL = CMDSQL + Format(Now, "yyyy-mm-dd") + "' and "
        CMDSQL = CMDSQL + " m.ststelpwith in ('OTHER','CH','SPOUSE','PARENT')"
        CMDSQL = CMDSQL + " group by m.agent,m.ststelpwith,u.spvcode "
        CMDSQL = CMDSQL + "order by m.agent,m.ststelpwith,u.spvcode asc"

    
    
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
        'ProgressBar1.Max = M_OBJRS.RecordCount
        While Not M_objrs.EOF
'             If Trim(m_objrs("status")) <> "SPOUSE" Or _
'                Trim(m_objrs("status")) <> "CONTACTED-CH" Or _
'                Trim(m_objrs("status")) <> "OTHER" Or _
'                Trim(m_objrs("status")) <> "PARENT" Or _
'                Trim(m_objrs("status")) <> "CH" Or _
'                Trim(m_objrs("status")) = "" Then
'                m_objrs.MoveNext
'             End If
            On Error Resume Next
            'ProgressBar1.Value = M_OBJRS.Bookmark
             CMDSQL = "update tblrptcontactto set ["
             CMDSQL = CMDSQL + Trim(Replace(M_objrs("status"), "/", "")) + "]='"
             CMDSQL = CMDSQL + CStr(M_objrs("jumlah")) + "' where spvcode='"
             CMDSQL = CMDSQL + Trim(M_objrs("spv")) + "' and agent='"
             CMDSQL = CMDSQL + Trim(M_objrs("agent")) + "'"
             M_RPTCONN.Execute CMDSQL
            M_objrs.MoveNext
        Wend
    End If
    Set M_objrs = Nothing
End Sub
'@@01 Juni 2011
Private Sub IsiContacttoJmlAcc()
    Dim M_objrs As ADODB.Recordset
    Dim m_objrs_rpt As ADODB.Recordset
    Dim CMDSQL As String
    
    
    'Ambil Data agentnya
    CMDSQL = "select spvcode,agent from tblrptcontactto order by spvcode,agent"
    Set m_objrs_rpt = New ADODB.Recordset
    m_objrs_rpt.CursorLocation = adUseClient
    m_objrs_rpt.Open CMDSQL, M_RPTCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If m_objrs_rpt.RecordCount > 0 Then
        
        While Not m_objrs_rpt.EOF
            
            CMDSQL = "select distinct custid from mgm_hst where agent='"
            CMDSQL = CMDSQL + Trim(m_objrs_rpt("agent")) + "' and date(tgl) between '"
            CMDSQL = CMDSQL + Format(Now, "yyyy-mm-dd") + "' and '"
            CMDSQL = CMDSQL + Format(Now, "yyyy-mm-dd") + "' "
            
            
            Set M_objrs = New ADODB.Recordset
            M_objrs.CursorLocation = adUseClient
            M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            
            'update data ke access
            CMDSQL = "update tblrptcontactto set jml_acc='"
            CMDSQL = CMDSQL + CStr(M_objrs.RecordCount) + "' where agent='"
            CMDSQL = CMDSQL + Trim(m_objrs_rpt("agent")) + "' and spvcode='"
            CMDSQL = CMDSQL + Trim(m_objrs_rpt("spvcode")) + "'"
            M_RPTCONN.Execute CMDSQL
            
            Set M_objrs = Nothing
            
            m_objrs_rpt.MoveNext
        Wend
    End If
    
    Set m_objrs_rpt = Nothing
End Sub

Private Sub SHOW_PRN()
    FrmMgmReport.RPT.RetrieveDataFiles
    FrmMgmReport.RPT.WindowLeft = 0
    FrmMgmReport.RPT.WindowTop = 0
    FrmMgmReport.RPT.WindowState = crptMaximized
    FrmMgmReport.RPT.WindowShowPrintBtn = True
    FrmMgmReport.RPT.WindowShowRefreshBtn = True
    FrmMgmReport.RPT.WindowShowSearchBtn = True
    FrmMgmReport.RPT.WindowShowPrintSetupBtn = True
    FrmMgmReport.RPT.WindowControls = True
    FrmMgmReport.RPT.PrintReport
    'RPT.Action = 1
    'RPT.Reset
End Sub

Private Sub CloseWskReq(ByVal Index As Integer)
'    WskRequest(Index).Close
'    Unload WskRequest(Index)
'
'    JmlKoneksiReq = JmlKoneksiReq - 1
End Sub

Private Sub WskRequest_ConnectionRequest(Index As Integer, ByVal requestID As Long)
'    JmlKoneksiReq = JmlKoneksiReq + 1
'    Load WskRequest(JmlKoneksiReq)
'    WskRequest(JmlKoneksiReq).Close
'
'    If JmlKoneksiReq <= MaxKoneksiReq Then
'        WskRequest(JmlKoneksiReq).Accept requestID
'    Else
'        CloseWskReq JmlKoneksiReq
'    End If
End Sub

Private Sub WskRequest_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'    Dim st As String
'
'    WskRequest(Index).GetData st
'    TxtOnline.Text = st & vbCrLf & TxtOnline.Text
'    TimerTandaReq.Enabled = True
End Sub
Private Sub sms_muncul()
    
    Dim satu As String
    Dim dua As String
    Dim tiga As String
    Dim empat As String
    Dim M_objrs As New ADODB.Recordset
    Dim cmdsql34 As String
    Dim TELPo As String
    Dim lst As ListItem
    Dim JmlBelumBaca As Integer
    Dim JmlSudahBaca As Integer

    'On Error Resume Next

    TELPo = "Select `ReceivingDateTime`, `SenderNumber`, `TextDecoded`,`ID`,`Processed` FROM inbox WHERE `SenderNumber` in ('a',"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    
    'Jika yang login Agent
    If UCase(Trim(MDIForm1.txtlevel.text)) = "AGENT" Then
        cmdsql34 = "select mobileno,mobileno2,mobilenoadd1,mobilenoadd2 from mgm where agent = '" + Trim(MDIForm1.TxtUsername.text) + "'"
        
    End If
    'Jika yang login TL
     If UCase(Trim(MDIForm1.txtlevel.text)) = "TEAMLEADER" Then
        MsgBox "Silahkan tunggu! Program akan mencari inbox dari semua team anda!", vbOKOnly + vbInformation, "Informasi"
        'cmdsql34 = "SELECT contact1,contact2,mobileno FROM tbl_address WHERE custid in (SELECT custno FROM mgm WHERE agent in ("
        cmdsql34 = "select mobileno,mobileno2,mobilenoadd1,mobilenoadd2 from mgm where agent IN ("
        cmdsql34 = cmdsql34 + "select userid from usertbl where team='"
        cmdsql34 = cmdsql34 + Trim(MDIForm1.TxtUsername.text) + "')) "
    End If
    'Jika yang login admin
    If UCase(Trim(MDIForm1.txtlevel.text)) = "ADMIN" Or UCase(Trim(MDIForm1.txtlevel.text)) = "ADMINISTRATOR" Or UCase(Trim(MDIForm1.txtlevel.text)) = "SUPERVISOR" Then
        'MsgBox "Silahkan tunggu! Program akan mencari inbox dari semua AGENT!", vbOKOnly + vbInformation, "Informasi"
        cmdsql34 = "select mobileno,mobileno2,mobilenoadd1,mobilenoadd2 from mgm "
    End If
    
    M_objrs.Open cmdsql34, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_objrs.EOF = False Then
        If M_objrs.RecordCount <> 0 Then
            'PB1.Max = M_objrs.RecordCount
        End If
    End If
    
    While Not M_objrs.EOF
        'PB1.Value = M_objrs.Bookmark
        
        If M_objrs("mobileno2") <> "" Then
            TELPo = TELPo + "'" + Replace(Trim(M_objrs("mobileno2")), " ", "") + "',"
        End If
        If M_objrs("mobilenoadd1") <> "" Then
            TELPo = TELPo + "'" + Replace(Trim(M_objrs("mobilenoadd1")), " ", "") + "',"
        End If
        If M_objrs("mobileno") <> "" Then
            TELPo = TELPo + "'" + Replace(Trim(M_objrs("mobileno")), " ", "") + "',"
        End If
        If M_objrs("mobilenoadd2") <> "" Then
            TELPo = TELPo + "'" + Replace(Trim(M_objrs("mobilenoadd2")), " ", "") + "',"
        End If
    
        M_objrs.MoveNext
    Wend
    
    Set M_objrs = Nothing
    
    TELPo = Left(TELPo, Len(TELPo) - 1)
    Dim TELPo1
    Dim TELPo2
    
    TELPo1 = TELPo + ") and `Processed`='false' order by `ReceivingDateTime` desc " 'Ini yang belum pernah di baca
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open TELPo1, M_OBJCONN1, adOpenDynamic, adLockOptimistic
    
    'Ini buat data inbox yang belum dibaca
    JmlBelumBaca = M_objrs.RecordCount
    If M_objrs.RecordCount <> 0 Then
        'PB1.Max = JmlBelumBaca
    Else
        Dim Update_Status As String
        'MsgBox "Tidak ada sms baru!", vbOKOnly + vbInformation, "Informasi"
        'Update status sms di usertbl jadi null, supaya ga blink
        Update_Status = "update usertbl set status_sms=null where userid='"
        Update_Status = Update_Status + Trim(MDIForm1.TxtUsername.text) + "'"
        M_OBJCONN.Execute Update_Status
        'MDIForm1.TimerBlink.Enabled = False
        MDIForm1.Label9.ForeColor = vbBlack
    End If
    While Not M_objrs.EOF
        'PB1.Value = M_objrs.Bookmark
        
        S = Format(M_objrs!receivingdatetime, "yyyy-mm-dd hh:mm:ss")
        t = Trim(M_objrs!sendernumber)
        u = M_objrs!textdecoded
        v = FindReplace(t, "+62", "0")
    
        If (Left(v, 3) = "021") Then
            v = Mid(v, 4, 20)
        End If
    
        Dim showlist As New ADODB.Recordset
        Dim TOTPTP As Currency
        Dim ssql As String
        
        If showlist.State = 1 Then showlist.Close
        ssql = "SELECT custid, name FROM mgm WHERE mobileno='" & v & "'  or mobileno2='" & v & "'  or mobilenoadd1='" & v & "'  or mobilenoadd2='" & v & "'"
        showlist.Open ssql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        frm_unread_msg.LvInboxOutbox.ListItems.CLEAR

        
        If showlist.EOF = False Then
            isicustid = showlist!CustId
            isiname = showlist!Name
            Set showlist = Nothing
        End If
        
        Set lst = frm_unread_msg.LvInboxOutbox.ListItems.ADD(, , Trim(isicustid)) 'custid
            lst.SubItems(1) = Trim(isiname)  'Isi nama
            lst.SubItems(2) = Trim(v) 'Telepon
            lst.SubItems(3) = Trim(S) 'Receivingdatetime
            lst.SubItems(4) = Trim(IIf(IsNull(M_objrs("TextDecoded")), "", M_objrs("TextDecoded"))) 'Textsms
            lst.SubItems(5) = M_objrs("id")
            lst.SubItems(6) = M_objrs("Processed")
            lst.Bold = True
            'frm_unread_msg.LvInboxOutbox.SelectedItem.ForeColor = vbRed
            lst.ListSubItems.ADD.ForeColor = vbRed
            lst.ListSubItems(1).ForeColor = vbRed
            lst.ListSubItems(2).ForeColor = vbRed
            lst.ListSubItems(3).ForeColor = vbRed
            lst.ListSubItems(4).ForeColor = vbRed
            lst.ListSubItems(5).ForeColor = vbRed
            lst.ListSubItems(6).ForeColor = vbRed
            M_objrs.MoveNext
        frm_unread_msg.Show 1
    Wend
    
    Set M_objrs = Nothing
End Sub
