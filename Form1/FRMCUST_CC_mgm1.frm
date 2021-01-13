VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FRMCUST_CC_mgm1 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   10500
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   Icon            =   "FRMCUST_CC_mgm1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      Height          =   405
      Left            =   4500
      TabIndex        =   94
      Top             =   0
      Visible         =   0   'False
      Width           =   2100
   End
   Begin Threed.SSCommand SSCommand1 
      Cancel          =   -1  'True
      Height          =   420
      Index           =   3
      Left            =   9180
      TabIndex        =   3
      Top             =   30
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   741
      _Version        =   196610
      Font3D          =   2
      MousePointer    =   16
      ForeColor       =   128
      BackColor       =   -2147483644
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
      Picture         =   "FRMCUST_CC_mgm1.frx":0442
      Caption         =   "&Keluar"
      AutoSize        =   2
      Alignment       =   4
      ButtonStyle     =   2
      PictureAlignment=   1
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   420
      Index           =   2
      Left            =   8100
      TabIndex        =   2
      Top             =   30
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   741
      _Version        =   196610
      Font3D          =   2
      MousePointer    =   16
      ForeColor       =   128
      BackColor       =   -2147483644
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
      Picture         =   "FRMCUST_CC_mgm1.frx":059C
      Caption         =   "&Simpan"
      AutoSize        =   2
      Alignment       =   4
      ButtonStyle     =   2
      PictureAlignment=   1
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   420
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   30
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   741
      _Version        =   196610
      Font3D          =   2
      MousePointer    =   16
      ForeColor       =   128
      BackColor       =   -2147483644
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
      Picture         =   "FRMCUST_CC_mgm1.frx":08BE
      Caption         =   "&Call"
      AutoSize        =   2
      Alignment       =   4
      ButtonStyle     =   2
      PictureAlignment=   1
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   420
      Index           =   1
      Left            =   1395
      TabIndex        =   1
      Top             =   30
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   741
      _Version        =   196610
      Font3D          =   2
      MousePointer    =   16
      ForeColor       =   128
      BackColor       =   -2147483644
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
      Picture         =   "FRMCUST_CC_mgm1.frx":1132
      Caption         =   "&MGM"
      AutoSize        =   2
      Alignment       =   4
      ButtonStyle     =   2
      PictureAlignment=   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   150
      TabIndex        =   52
      Top             =   435
      Width           =   10200
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
         Left            =   915
         TabIndex        =   4
         Top             =   135
         Width           =   1020
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         Left            =   6870
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   150
         Width           =   3270
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
         Left            =   1965
         MaxLength       =   200
         TabIndex        =   5
         Top             =   135
         Width           =   3900
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000004&
         Caption         =   "Id :"
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
         Index           =   1
         Left            =   6525
         TabIndex        =   54
         Top             =   210
         Width           =   1080
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000004&
         Caption         =   "Name :"
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
         Index           =   0
         Left            =   195
         TabIndex        =   53
         Top             =   195
         Width           =   960
      End
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
      Left            =   3015
      TabIndex        =   8
      Top             =   975
      Width           =   3825
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
      Left            =   1500
      TabIndex        =   7
      Top             =   975
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Index           =   3
      Left            =   1500
      Locked          =   -1  'True
      MaxLength       =   200
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   975
      Visible         =   0   'False
      Width           =   4410
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   45
      TabIndex        =   10
      Top             =   1380
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   -2147483644
      ForeColor       =   4194368
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Data Pri&badi"
      TabPicture(0)   =   "FRMCUST_CC_mgm1.frx":15FA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Data Pe&kerjaan"
      TabPicture(1)   =   "FRMCUST_CC_mgm1.frx":1616
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Ca&tatan Data"
      TabPicture(2)   =   "FRMCUST_CC_mgm1.frx":1632
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&Catatan Kerja"
      TabPicture(3)   =   "FRMCUST_CC_mgm1.frx":164E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame9"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "&Hasil Kerja"
      TabPicture(4)   =   "FRMCUST_CC_mgm1.frx":166A
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Frame10"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame6 
         BackColor       =   &H80000004&
         Height          =   4680
         Left            =   -74910
         TabIndex        =   92
         Top             =   450
         Width           =   10275
         Begin VB.TextBox Text3 
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
            Height          =   4455
            Left            =   30
            Locked          =   -1  'True
            TabIndex        =   93
            Top             =   165
            Width           =   10215
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H80000004&
         Height          =   4590
         Left            =   -74940
         TabIndex        =   90
         Top             =   450
         Width           =   10320
         Begin MSComctlLib.ListView ListView1 
            Height          =   4410
            Index           =   1
            Left            =   30
            TabIndex        =   39
            Top             =   135
            Width           =   10260
            _ExtentX        =   18098
            _ExtentY        =   7779
            View            =   3
            LabelEdit       =   1
            SortOrder       =   -1  'True
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
      Begin VB.Frame Frame10 
         BackColor       =   &H80000004&
         Height          =   5100
         Left            =   90
         TabIndex        =   80
         Top             =   375
         Width           =   10305
         Begin VB.CheckBox Check2 
            BackColor       =   &H80000004&
            Caption         =   "Tidak Terhubungi"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Index           =   1
            Left            =   5625
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   40
            Top             =   135
            Width           =   1800
         End
         Begin VB.Frame Frame22 
            BackColor       =   &H80000004&
            Height          =   1260
            Left            =   60
            TabIndex        =   86
            Top             =   3765
            Width           =   4710
            Begin VB.TextBox Text6 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               BorderStyle     =   0  'None
               ForeColor       =   &H00800080&
               Height          =   225
               Left            =   390
               TabIndex        =   96
               Text            =   "Priority"
               Top             =   870
               Width           =   600
            End
            Begin VB.ComboBox Combo5 
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
               Left            =   1065
               TabIndex        =   95
               Top             =   825
               Width           =   1350
            End
            Begin VB.TextBox Text2 
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
               Height          =   330
               Left            =   1065
               TabIndex        =   44
               Top             =   165
               Width           =   3585
            End
            Begin TDBTime6Ctl.TDBTime TDBTime1 
               Height          =   315
               Left            =   2535
               TabIndex        =   46
               Top             =   495
               Width           =   900
               _Version        =   65536
               _ExtentX        =   1587
               _ExtentY        =   556
               Caption         =   "FRMCUST_CC_mgm1.frx":1686
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_mgm1.frx":16F2
               Spin            =   "FRMCUST_CC_mgm1.frx":1742
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
               ForeColor       =   0
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
               ShowContextMenu =   1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "__:__"
               ValidateMode    =   0
               ValueVT         =   1
               Value           =   0.429861111111111
            End
            Begin TDBDate6Ctl.TDBDate TDBDate1 
               Height          =   315
               Index           =   3
               Left            =   1065
               TabIndex        =   45
               Top             =   495
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               Calendar        =   "FRMCUST_CC_mgm1.frx":176A
               Caption         =   "FRMCUST_CC_mgm1.frx":1882
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FRMCUST_CC_mgm1.frx":18EE
               Keys            =   "FRMCUST_CC_mgm1.frx":190C
               Spin            =   "FRMCUST_CC_mgm1.frx":196A
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
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000004&
               Caption         =   "Next Action "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800080&
               Height          =   195
               Index           =   3
               Left            =   75
               TabIndex        =   88
               Top             =   240
               Width           =   960
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000004&
               Caption         =   "Schedule"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800080&
               Height          =   180
               Index           =   5
               Left            =   165
               TabIndex        =   87
               Top             =   570
               Width           =   825
            End
         End
         Begin VB.Frame Frame25 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   615
            Left            =   5430
            TabIndex        =   81
            Top             =   165
            Width           =   4800
            Begin VB.ComboBox Combo3 
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
               Index           =   1
               Left            =   540
               TabIndex        =   50
               Top             =   225
               Width           =   4155
            End
            Begin VB.ComboBox Combo3 
               Height          =   315
               Index           =   0
               Left            =   540
               TabIndex        =   82
               TabStop         =   0   'False
               Top             =   225
               Visible         =   0   'False
               Width           =   4155
            End
            Begin VB.Label Label7 
               BackColor       =   &H80000004&
               Caption         =   "Ket:"
               ForeColor       =   &H00800080&
               Height          =   225
               Left            =   180
               TabIndex        =   83
               Top             =   285
               Width           =   315
            End
         End
         Begin RichTextLib.RichTextBox RichTextBox1 
            Height          =   735
            Index           =   3
            Left            =   4800
            TabIndex        =   47
            Top             =   4035
            Width           =   5460
            _ExtentX        =   9631
            _ExtentY        =   1296
            _Version        =   393217
            TextRTF         =   $"FRMCUST_CC_mgm1.frx":1992
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
         Begin VB.Frame Frame12 
            BackColor       =   &H80000004&
            Height          =   3630
            Left            =   75
            TabIndex        =   84
            Top             =   150
            Width           =   5340
            Begin VB.CheckBox Check2 
               BackColor       =   &H80000004&
               Caption         =   "Terhubungi"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   240
               Index           =   0
               Left            =   120
               MouseIcon       =   "FRMCUST_CC_mgm1.frx":1A0D
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   41
               Top             =   -30
               Width           =   1350
            End
            Begin VB.Frame Frame19 
               BackColor       =   &H80000004&
               Enabled         =   0   'False
               Height          =   3465
               Index           =   0
               Left            =   30
               TabIndex        =   85
               Top             =   120
               Width           =   5280
               Begin VB.OptionButton Option2 
                  BackColor       =   &H80000004&
                  Caption         =   "Ada Referensi"
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
                  Height          =   225
                  Index           =   0
                  Left            =   60
                  TabIndex        =   43
                  Top             =   495
                  Width           =   1725
               End
               Begin VB.OptionButton Option2 
                  BackColor       =   &H80000004&
                  Caption         =   "Tidak Ada Referensi"
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
                  Height          =   195
                  Index           =   1
                  Left            =   60
                  TabIndex        =   42
                  Top             =   195
                  Width           =   2250
               End
               Begin VB.Frame Frame11 
                  BackColor       =   &H80000004&
                  Enabled         =   0   'False
                  Height          =   2880
                  Left            =   60
                  TabIndex        =   97
                  Top             =   555
                  Width           =   5025
                  Begin VB.CommandButton Command1 
                     Caption         =   "&Delete"
                     Height          =   300
                     Index           =   2
                     Left            =   2295
                     TabIndex        =   101
                     Top             =   225
                     Width           =   810
                  End
                  Begin VB.CommandButton Command1 
                     Caption         =   "&Update"
                     Height          =   300
                     Index           =   1
                     Left            =   1455
                     TabIndex        =   100
                     Top             =   225
                     Width           =   810
                  End
                  Begin VB.CommandButton Command1 
                     Caption         =   "&Add"
                     Height          =   300
                     Index           =   0
                     Left            =   615
                     TabIndex        =   99
                     Top             =   225
                     Width           =   810
                  End
                  Begin MSComctlLib.ListView ListView1 
                     Height          =   2325
                     Index           =   0
                     Left            =   30
                     TabIndex        =   98
                     Top             =   510
                     Width           =   4950
                     _ExtentX        =   8731
                     _ExtentY        =   4101
                     View            =   3
                     LabelEdit       =   1
                     SortOrder       =   -1  'True
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
            End
         End
         Begin VB.Label Label4 
            BackColor       =   &H80000004&
            Caption         =   "Catatan :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   270
            Index           =   3
            Left            =   4800
            TabIndex        =   89
            Top             =   3840
            Width           =   825
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H80000004&
         Height          =   5130
         Left            =   -74955
         TabIndex        =   66
         Top             =   360
         Width           =   10350
         Begin VB.OptionButton Option4 
            BackColor       =   &H80000004&
            Caption         =   "Data Wiraswasta/Professional/Lainnya"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   1
            Left            =   5340
            TabIndex        =   25
            Top             =   555
            Width           =   3840
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H80000004&
            Caption         =   " Data Karyawan"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   0
            Left            =   375
            TabIndex        =   22
            Top             =   555
            Width           =   1755
         End
         Begin VB.Frame Frame18 
            BackColor       =   &H80000004&
            Height          =   1695
            Left            =   3075
            TabIndex        =   71
            Top             =   2415
            Width           =   4275
            Begin VB.TextBox Text1 
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
               Index           =   21
               Left            =   3030
               MaxLength       =   6
               TabIndex        =   31
               Top             =   195
               Width           =   765
            End
            Begin VB.TextBox Text1 
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
               Index           =   32
               Left            =   3030
               MaxLength       =   6
               TabIndex        =   34
               Top             =   510
               Width           =   765
            End
            Begin TDBMask6Ctl.TDBMask TDBMask1 
               Height          =   360
               Index           =   4
               Left            =   1725
               TabIndex        =   30
               Top             =   165
               Width           =   990
               _Version        =   65536
               _ExtentX        =   1746
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC_mgm1.frx":1E4F
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_mgm1.frx":1EBB
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
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
               Format          =   "999-9999999999999999"
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
               ShowContextMenu =   1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "___-________________"
               Value           =   ""
            End
            Begin TDBMask6Ctl.TDBMask TDBMask1 
               Height          =   360
               Index           =   5
               Left            =   1725
               TabIndex        =   33
               Top             =   540
               Width           =   990
               _Version        =   65536
               _ExtentX        =   1746
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC_mgm1.frx":1EFD
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_mgm1.frx":1F69
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
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
               Format          =   "999-9999999999999999"
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
               ShowContextMenu =   1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "___-________________"
               Value           =   ""
            End
            Begin TDBMask6Ctl.TDBMask TDBMask1 
               Height          =   360
               Index           =   6
               Left            =   1725
               TabIndex        =   36
               Top             =   915
               Width           =   990
               _Version        =   65536
               _ExtentX        =   1746
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC_mgm1.frx":1FAB
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_mgm1.frx":2017
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
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
               Format          =   "999-99999999999999999"
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
               ShowContextMenu =   1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "___-_________________"
               Value           =   ""
            End
            Begin TDBMask6Ctl.TDBMask TDBMask1 
               Height          =   360
               Index           =   7
               Left            =   1725
               TabIndex        =   38
               Top             =   1275
               Width           =   990
               _Version        =   65536
               _ExtentX        =   1746
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC_mgm1.frx":2059
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_mgm1.frx":20C5
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
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
               Format          =   "999-9999999999999999"
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
               ShowContextMenu =   1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "___-________________"
               Value           =   ""
            End
            Begin TDBMask6Ctl.TDBMask TDBMask1 
               Height          =   360
               Index           =   10
               Left            =   1125
               TabIndex        =   29
               Top             =   165
               Width           =   585
               _Version        =   65536
               _ExtentX        =   1032
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC_mgm1.frx":2107
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_mgm1.frx":2173
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
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
               PromptChar      =   "_"
               ReadOnly        =   0
               ShowContextMenu =   1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "[____]"
               Value           =   ""
            End
            Begin TDBMask6Ctl.TDBMask TDBMask1 
               Height          =   360
               Index           =   11
               Left            =   1125
               TabIndex        =   32
               Top             =   540
               Width           =   585
               _Version        =   65536
               _ExtentX        =   1032
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC_mgm1.frx":21B5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_mgm1.frx":2221
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
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
               PromptChar      =   "_"
               ReadOnly        =   0
               ShowContextMenu =   1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "[____]"
               Value           =   ""
            End
            Begin TDBMask6Ctl.TDBMask TDBMask1 
               Height          =   360
               Index           =   12
               Left            =   1125
               TabIndex        =   35
               Top             =   915
               Width           =   585
               _Version        =   65536
               _ExtentX        =   1032
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC_mgm1.frx":2263
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_mgm1.frx":22CF
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
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
               PromptChar      =   "_"
               ReadOnly        =   0
               ShowContextMenu =   1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "[____]"
               Value           =   ""
            End
            Begin TDBMask6Ctl.TDBMask TDBMask1 
               Height          =   360
               Index           =   13
               Left            =   1125
               TabIndex        =   37
               Top             =   1275
               Width           =   585
               _Version        =   65536
               _ExtentX        =   1032
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC_mgm1.frx":2311
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_mgm1.frx":237D
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
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
               PromptChar      =   "_"
               ReadOnly        =   0
               ShowContextMenu =   1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "[____]"
               Value           =   ""
            End
            Begin VB.Label Label6 
               BackColor       =   &H80000004&
               Caption         =   "Ext."
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800080&
               Height          =   270
               Index           =   16
               Left            =   2730
               TabIndex        =   74
               Top             =   255
               Width           =   270
            End
            Begin VB.Label Label6 
               BackColor       =   &H80000004&
               Caption         =   "No. Telp"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800080&
               Height          =   270
               Index           =   10
               Left            =   330
               TabIndex        =   73
               Top             =   255
               Width           =   810
            End
            Begin VB.Label Label6 
               BackColor       =   &H80000004&
               Caption         =   "No. Fax"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800080&
               Height          =   270
               Index           =   13
               Left            =   345
               TabIndex        =   72
               Top             =   930
               Width           =   810
            End
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H80000004&
            Caption         =   "     "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1740
            Left            =   5265
            TabIndex        =   67
            Top             =   540
            Width           =   4950
            Begin VB.TextBox Text1 
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
               Index           =   34
               Left            =   1410
               MaxLength       =   30
               TabIndex        =   28
               Top             =   1200
               Width           =   3435
            End
            Begin VB.TextBox Text1 
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
               Index           =   18
               Left            =   1410
               MaxLength       =   30
               TabIndex        =   26
               Top             =   315
               Width           =   3435
            End
            Begin RichTextLib.RichTextBox RichTextBox1 
               Height          =   615
               Index           =   4
               Left            =   1395
               TabIndex        =   27
               Top             =   615
               Width           =   3465
               _ExtentX        =   6112
               _ExtentY        =   1085
               _Version        =   393217
               TextRTF         =   $"FRMCUST_CC_mgm1.frx":23BF
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
            Begin VB.Label Label3 
               BackColor       =   &H80000004&
               Caption         =   "Jenis Usaha"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800080&
               Height          =   225
               Index           =   0
               Left            =   180
               TabIndex        =   70
               Top             =   1260
               Width           =   1155
            End
            Begin VB.Label Label4 
               BackColor       =   &H80000004&
               Caption         =   "Alamat "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800080&
               Height          =   255
               Index           =   4
               Left            =   195
               TabIndex        =   69
               Top             =   690
               Width           =   1080
            End
            Begin VB.Label Label3 
               BackColor       =   &H80000004&
               Caption         =   "Nama Usaha"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800080&
               Height          =   225
               Index           =   14
               Left            =   180
               TabIndex        =   68
               Top             =   360
               Width           =   1020
            End
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H80000004&
            Caption         =   "   "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1755
            Left            =   210
            TabIndex        =   75
            Top             =   540
            Width           =   4995
            Begin VB.TextBox Text1 
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
               Index           =   20
               Left            =   1470
               MaxLength       =   30
               TabIndex        =   23
               Top             =   240
               Width           =   3435
            End
            Begin VB.ComboBox Combo1 
               BackColor       =   &H00C0FFFF&
               ForeColor       =   &H00C00000&
               Height          =   315
               Index           =   16
               Left            =   1470
               TabIndex        =   48
               Top             =   1770
               Visible         =   0   'False
               Width           =   1800
            End
            Begin VB.TextBox Text1 
               BackColor       =   &H00C0FFFF&
               ForeColor       =   &H00C00000&
               Height          =   315
               Index           =   30
               Left            =   1470
               MaxLength       =   30
               TabIndex        =   49
               Top             =   2085
               Visible         =   0   'False
               Width           =   3300
            End
            Begin RichTextLib.RichTextBox RichTextBox1 
               Height          =   615
               Index           =   2
               Left            =   1455
               TabIndex        =   24
               Top             =   540
               Width           =   3465
               _ExtentX        =   6112
               _ExtentY        =   1085
               _Version        =   393217
               TextRTF         =   $"FRMCUST_CC_mgm1.frx":243A
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
            Begin VB.Label Label3 
               BackColor       =   &H80000004&
               Caption         =   "Nama Perusahaan"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800080&
               Height          =   225
               Index           =   16
               Left            =   105
               TabIndex        =   79
               Top             =   285
               Width           =   1320
            End
            Begin VB.Label Label4 
               BackColor       =   &H80000004&
               Caption         =   "Alamat Perusahaan"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800080&
               Height          =   375
               Index           =   2
               Left            =   105
               TabIndex        =   78
               Top             =   570
               Width           =   1365
            End
            Begin VB.Label Label3 
               Caption         =   "Gaji / Bulan"
               ForeColor       =   &H00800080&
               Height          =   225
               Index           =   20
               Left            =   105
               TabIndex        =   77
               Top             =   1845
               Visible         =   0   'False
               Width           =   1305
            End
            Begin VB.Label Label3 
               Caption         =   "Nama Atasan "
               ForeColor       =   &H00800080&
               Height          =   225
               Index           =   18
               Left            =   120
               TabIndex        =   76
               Top             =   2130
               Visible         =   0   'False
               Width           =   1095
            End
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000004&
         Height          =   5130
         Left            =   -74955
         TabIndex        =   55
         Top             =   360
         Width           =   10350
         Begin VB.Frame Frame3 
            BackColor       =   &H80000004&
            Height          =   4230
            Index           =   0
            Left            =   30
            TabIndex        =   59
            Top             =   120
            Width           =   5880
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800080&
               Height          =   315
               Index           =   2
               Left            =   3165
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   15
               TabStop         =   0   'False
               Text            =   "0"
               Top             =   1005
               Width           =   375
            End
            Begin VB.TextBox Text1 
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
               Index           =   6
               Left            =   1785
               MaxLength       =   5
               TabIndex        =   12
               Top             =   675
               Width           =   945
            End
            Begin VB.TextBox Text1 
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
               Index           =   7
               Left            =   3255
               MaxLength       =   20
               TabIndex        =   13
               Top             =   675
               Width           =   2520
            End
            Begin TDBDate6Ctl.TDBDate TDBDate1 
               Height          =   315
               Index           =   0
               Left            =   1770
               TabIndex        =   14
               Top             =   1005
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   556
               Calendar        =   "FRMCUST_CC_mgm1.frx":24B5
               Caption         =   "FRMCUST_CC_mgm1.frx":25CD
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FRMCUST_CC_mgm1.frx":2639
               Keys            =   "FRMCUST_CC_mgm1.frx":2657
               Spin            =   "FRMCUST_CC_mgm1.frx":26B5
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
            Begin RichTextLib.RichTextBox RichTextBox1 
               Height          =   555
               Index           =   0
               Left            =   1770
               TabIndex        =   11
               Top             =   135
               Width           =   4020
               _ExtentX        =   7091
               _ExtentY        =   979
               _Version        =   393217
               TextRTF         =   $"FRMCUST_CC_mgm1.frx":26DD
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
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   315
               Left            =   1770
               TabIndex        =   60
               Top             =   1005
               Visible         =   0   'False
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
               _Version        =   393216
               Format          =   56688640
               CurrentDate     =   37459
            End
            Begin VB.Label Label3 
               BackColor       =   &H80000004&
               Caption         =   "Tahun"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800080&
               Height          =   225
               Index           =   6
               Left            =   3585
               TabIndex        =   65
               Top             =   1050
               Width           =   555
            End
            Begin VB.Label Label3 
               BackColor       =   &H80000004&
               Caption         =   "Tanggal Lahir"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800080&
               Height          =   210
               Index           =   1
               Left            =   135
               TabIndex        =   64
               Top             =   1065
               Width           =   1350
            End
            Begin VB.Label Label6 
               BackColor       =   &H80000004&
               Caption         =   "Kota"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800080&
               Height          =   225
               Index           =   0
               Left            =   2790
               TabIndex        =   63
               Top             =   720
               Width           =   405
            End
            Begin VB.Label Label4 
               BackColor       =   &H80000004&
               Caption         =   "Alamat Rumah Sekarang"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800080&
               Height          =   435
               Index           =   0
               Left            =   165
               TabIndex        =   62
               Top             =   225
               Width           =   1365
            End
            Begin VB.Label Label5 
               BackColor       =   &H80000004&
               Caption         =   "Kode Pos"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800080&
               Height          =   255
               Index           =   0
               Left            =   165
               TabIndex        =   61
               Top             =   720
               Width           =   1275
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H80000004&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   4230
            Index           =   1
            Left            =   5925
            TabIndex        =   56
            Top             =   120
            Width           =   4170
            Begin TDBMask6Ctl.TDBMask TDBMask1 
               Height          =   360
               Index           =   0
               Left            =   2130
               TabIndex        =   17
               Top             =   150
               Width           =   1110
               _Version        =   65536
               _ExtentX        =   1958
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC_mgm1.frx":2758
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_mgm1.frx":27C4
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
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
               Format          =   "999-99999999999999999"
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
               ShowContextMenu =   1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "___-_________________"
               Value           =   ""
            End
            Begin TDBMask6Ctl.TDBMask TDBMask1 
               Height          =   360
               Index           =   1
               Left            =   2130
               TabIndex        =   19
               Top             =   510
               Width           =   1110
               _Version        =   65536
               _ExtentX        =   1958
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC_mgm1.frx":2806
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_mgm1.frx":2872
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
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
               Format          =   "999-99999999999999999"
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
               ShowContextMenu =   1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "___-_________________"
               Value           =   ""
            End
            Begin TDBMask6Ctl.TDBMask TDBMask1 
               Height          =   360
               Index           =   2
               Left            =   1515
               TabIndex        =   20
               Top             =   870
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC_mgm1.frx":28B4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_mgm1.frx":2920
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
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
               Format          =   "9999-99999999999999999"
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
               ShowContextMenu =   1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "____-_________________"
               Value           =   ""
            End
            Begin TDBMask6Ctl.TDBMask TDBMask1 
               Height          =   360
               Index           =   3
               Left            =   1515
               TabIndex        =   21
               Top             =   1245
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC_mgm1.frx":2962
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_mgm1.frx":29CE
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
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
               Format          =   "9999-99999999999999999"
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
               ShowContextMenu =   1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "____-_________________"
               Value           =   ""
            End
            Begin TDBMask6Ctl.TDBMask TDBMask1 
               Height          =   360
               Index           =   8
               Left            =   1515
               TabIndex        =   16
               Top             =   150
               Width           =   585
               _Version        =   65536
               _ExtentX        =   1032
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC_mgm1.frx":2A10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_mgm1.frx":2A7C
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
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
               PromptChar      =   "_"
               ReadOnly        =   0
               ShowContextMenu =   1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "[____]"
               Value           =   ""
            End
            Begin TDBMask6Ctl.TDBMask TDBMask1 
               Height          =   360
               Index           =   9
               Left            =   1515
               TabIndex        =   18
               Top             =   510
               Width           =   585
               _Version        =   65536
               _ExtentX        =   1032
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC_mgm1.frx":2ABE
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_mgm1.frx":2B2A
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
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
               PromptChar      =   "_"
               ReadOnly        =   0
               ShowContextMenu =   1
               ShowLiterals    =   0
               TabAction       =   0
               Text            =   "[____]"
               Value           =   ""
            End
            Begin VB.Label Label6 
               BackColor       =   &H80000004&
               Caption         =   "No. Mobile"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800080&
               Height          =   270
               Index           =   3
               Left            =   600
               TabIndex        =   58
               Top             =   900
               Width           =   810
            End
            Begin VB.Label Label6 
               BackColor       =   &H80000004&
               Caption         =   "No. Telp"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800080&
               Height          =   270
               Index           =   2
               Left            =   600
               TabIndex        =   57
               Top             =   225
               Width           =   810
            End
         End
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Credit Card"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   390
      Index           =   5
      Left            =   4965
      TabIndex        =   91
      Top             =   60
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000004&
      Caption         =   "Sumber Info :"
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
      Height          =   270
      Left            =   300
      TabIndex        =   51
      Top             =   1020
      Width           =   1260
   End
End
Attribute VB_Name = "FRMCUST_CC_mgm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public DOK1 As String
Public STATUS_CUSTOMER As String
Public KETHSLKERJA As String
Public STATUS_FIELD_LAMA As String


Private Sub Check2_Click(Index As Integer)
Select Case Index
    Case 0
        If Check2(Index).Value Then
            Check2(1).Enabled = False
            Check2(1).Value = 0
            Combo3(1).Text = Empty
            Frame12.Enabled = True
            Frame19(0).Enabled = True
        Else
            Frame19(0).Enabled = False
            Option2(0).Value = False
            Option2(1).Value = False
'            Combo2(0).Text = Empty
'            Check9(0).Value = 0
'            Check9(1).Value = 0
'            Check9(2).Value = 0
            Check2(1).Enabled = True
'            TDBNumber1(0).Value = 0
'            TDBNumber1(3).Value = 0
'           Check2(4).Value = 0
'            Check2(4).Visible = False
'            Frame11.Visible = False
'            Combo2(0).Text = Empty
'            Combo2(2).Text = Empty
        End If
    Case 1
        If Check2(Index).Value Then
            Check2(0).Enabled = False
            Check2(0).Value = 0
            Frame25.Enabled = True
            Option2(0).Value = False
            Option2(1).Value = False
'            Combo2(0).Text = Empty
'            Check9(0).Value = 0
'            Check9(1).Value = 0
'            Check9(2).Value = 0
        Else
            Frame25.Enabled = False
            Check2(0).Enabled = True
            Frame12.Enabled = True
            Combo3(1).Text = Empty
            Combo3(0).Text = Empty
        End If
    Case 2
        If Check2(Index).Value Then
'            Frame20(0).Enabled = True
        Else
'            Frame20(0).Enabled = False
'            Option3(0).Value = False
'            Option3(1).Value = False
'            Check9(3).Value = 0
'            Check9(4).Value = 0
'            Check9(5).Value = 0
'            Check9(6).Value = 0
'            Check9(7).Value = 0
'            Check9(8).Value = 0
'            Check9(9).Value = 0
        End If
    Case 3
        If Check2(Index).Value Then
'            Frame20(0).Enabled = False
'            Frame23.Enabled = True
'            Check9(2).Enabled = False
'            Option3(0).Enabled = False
'            Option3(1).Enabled = False
        Else
'            Frame20(0).Enabled = True
'            Frame23.Enabled = False
'            Check9(2).Enabled = True
'            Option5(0).Value = False
'            Option5(1).Value = False
'            Option5(2).Value = False
'            Option6(0).Value = False
'            Option6(1).Value = False
'            Combo2(1).Text = Empty
'            Option3(0).Enabled = True
'            Option3(1).Enabled = True
'            TDBNumber1(2).Value = 0
'            TDBNumber1(1).Value = 0
'            TDBDate1(2).Text = Empty
'            Combo2(3).Text = Empty
        End If
    Case 4
        If Check2(Index).Value Then
'            TDBNumber1(3).Enabled = True
'            Frame11.Enabled = True
        Else
'            TDBNumber1(3).Value = 0
'           Option8(0).Value = False
'            Option8(1).Value = False
            TDBDate1(1).Text = Empty
'            TDBNumber1(3).Enabled = False
'            Frame11.Enabled = False
        End If
End Select
End Sub

Private Sub Combo1_Click(Index As Integer)
Dim M_DATA As New CLS_FRMCUST_CC_MGM
Dim M_OBJRS As ADODB.Recordset
Select Case Index
    Case 0
        Set M_OBJRS = M_DATA.QUERY_COMBO_DATASOURCE(M_OBJCONN, "KODEDS = '" + Combo1(Index).Text + "'")
        If M_OBJRS.RecordCount <> 0 Then
            Combo1(0).Text = M_OBJRS("KODEDS")
            Combo1(1).Text = IIf(IsNull(M_OBJRS("KETERANGAN")), "", M_OBJRS("KETERANGAN"))
            Text1(3).Text = IIf(IsNull(M_OBJRS("KETERANGAN")), "", M_OBJRS("KETERANGAN"))
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
            Text1(3).Text = Empty
        End If
    Case 1
        Set M_OBJRS = M_DATA.QUERY_COMBO_DATASOURCE(M_OBJCONN, "KETERANGAN = '" + Combo1(Index).Text + "'")
        If M_OBJRS.RecordCount <> 0 Then
            Combo1(0).Text = M_OBJRS("KODEDS")
            Combo1(1).Text = IIf(IsNull(M_OBJRS("KETERANGAN")), "", M_OBJRS("KETERANGAN"))
            Text1(3).Text = IIf(IsNull(M_OBJRS("KETERANGAN")), "", M_OBJRS("KETERANGAN"))
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
            Text1(3).Text = Empty
        End If
    Case 2
        Select Case UCase(Combo1(Index).Text)
            Case "BAPAK"
            Case "IBU"
            Case Else
                Combo1(Index).Text = Empty
        End Select
    Case 7
        Select Case UCase(Combo1(Index).Text)
            Case "MILIK SENDIRI"
            Case "MILIK PERUSAHAAN"
            Case "KELUARGA"
            Case "ANGSURAN"
            Case "KOS"
            Case "LAINNYA"
            Case Else
                Combo1(Index).Text = Empty
        End Select
End Select
Set M_OBJRS = Nothing
Set M_DATA = Nothing
End Sub

Private Sub Combo1_GotFocus(Index As Integer)
SendKeys "{Home}+{End}"
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim M_DATA As New CLS_FRMCUST_CC_MGM
Dim M_OBJRS As ADODB.Recordset
Select Case Index
    Case 0
        Set M_OBJRS = M_DATA.QUERY_COMBO_DATASOURCE(M_OBJCONN, "KODEDS = '" + Combo1(Index).Text + "'")
        If M_OBJRS.RecordCount <> 0 Then
            Combo1(0).Text = M_OBJRS("KODEDS")
            Combo1(1).Text = IIf(IsNull(M_OBJRS("KETERANGAN")), "", M_OBJRS("KETERANGAN"))
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    Case 1
        Set M_OBJRS = M_DATA.QUERY_COMBO_DATASOURCE(M_OBJCONN, "KETERANGAN = '" + Combo1(Index).Text + "'")
        If M_OBJRS.RecordCount <> 0 Then
            Combo1(0).Text = M_OBJRS("KODEDS")
            Combo1(1).Text = IIf(IsNull(M_OBJRS("KETERANGAN")), "", M_OBJRS("KETERANGAN"))
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    Case 2
        Select Case UCase(Combo1(Index).Text)
            Case "BAPAK"
            Case "IBU"
            Case Else
                Combo1(Index).Text = Empty
        End Select
    Case 7
        Select Case UCase(Combo1(Index).Text)
            Case "MILIK SENDIRI"
            Case "MILIK PERUSAHAAN"
            Case "KELUARGA"
            Case "ANGSURAN"
            Case "KOS"
            Case "LAINNYA"
            Case Else
                Combo1(Index).Text = Empty
        End Select
End Select
Set M_OBJRS = Nothing
Set M_DATA = Nothing
End Sub



Private Sub Combo3_LostFocus(Index As Integer)
Dim M_DATA As New CLS_FRMCUST_CC_MGM
Dim M_OBJRS As ADODB.Recordset
Select Case Index
    Case 0
        Set M_OBJRS = M_DATA.QUERY_COMBO_CLOSSING(M_OBJCONN, "KDCLS = '" + Combo3(Index).Text + "'")
        If M_OBJRS.RecordCount <> 0 Then
            Combo3(0).Text = M_OBJRS("KDCLS")
            Combo3(1).Text = IIf(IsNull(M_OBJRS("KETCLS")), "", M_OBJRS("KETCLS"))
        Else
            Combo3(0).Text = Empty
            Combo3(1).Text = Empty
        End If
    Case 1
        Set M_OBJRS = M_DATA.QUERY_COMBO_CLOSSING(M_OBJCONN, "KETCLS = '" + Combo3(Index).Text + "'")
        If M_OBJRS.RecordCount <> 0 Then
            Combo3(0).Text = M_OBJRS("KDCLS")
            Combo3(1).Text = IIf(IsNull(M_OBJRS("KETCLS")), "", M_OBJRS("KETCLS"))
        Else
            Combo3(0).Text = Empty
            Combo3(1).Text = Empty
        End If
End Select
Set M_OBJRS = Nothing
Set M_DATA = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
Dim CMDSQL As String
Dim LISTITEM As LISTITEM
Dim l_mgm_id As String
Dim LMSGBOX As Variant
Select Case Index
Case 0
    With frmcust_cc_ref
        .ADD = True
        .Show vbModal
        If .ok Then
            l_mgm_id = "REF" + CUSTNOMOR(M_OBJCONN, UCase(Me.Name))
            
            CMDSQL = "Insert into MGM_REF"
            CMDSQL = CMDSQL + " ( CUSTID,MGM_ID,NAMA,OFFICENO,EXTOFFICENO,OFFICENO2,EXTOFFICENO2,HOMENO,HOMENO2,HUBUNGAN)"
            CMDSQL = CMDSQL + " VALUES"
            CMDSQL = CMDSQL + " ('" + Text1(1).Text + "',"
            CMDSQL = CMDSQL + " '" + l_mgm_id + "',"
            CMDSQL = CMDSQL + " '" + .Text1.Text + "',"
            CMDSQL = CMDSQL + " '" + .Text2.Text + "',"
            CMDSQL = CMDSQL + " '" + .Text3.Text + "',"
            CMDSQL = CMDSQL + " '" + .Text4.Text + "',"
            CMDSQL = CMDSQL + " '" + .Text5.Text + "',"
            CMDSQL = CMDSQL + " '" + .Text6.Text + "',"
            CMDSQL = CMDSQL + " '" + .Text7.Text + "',"
            CMDSQL = CMDSQL + " '" + .Combo1.Text + "')"
            M_OBJCONN.Execute CMDSQL
            Call show_gridref
        End If
    End With
    Unload frmcust_cc_ref
    
Case 1
    If ListView1(0).ListItems.Count = 0 Then
        Exit Sub
    End If
    
    With frmcust_cc_ref
        .ADD = False
        .Show vbModal
        CMDSQL = "Update MGM_REF SET "
        CMDSQL = CMDSQL + " NAMA ='" + .Text1.Text + "',"
        CMDSQL = CMDSQL + " OFFICENO ='" + .Text2.Text + "',"
        CMDSQL = CMDSQL + " EXTOFFICENO ='" + .Text3.Text + "',"
        CMDSQL = CMDSQL + " OFFICENO2 ='" + .Text4.Text + "',"
        CMDSQL = CMDSQL + " EXTOFFICENO2 ='" + .Text5.Text + "',"
        CMDSQL = CMDSQL + " HOMENO ='" + .Text6.Text + "',"
        CMDSQL = CMDSQL + " HOMENO2 ='" + .Text7.Text + "',"
        CMDSQL = CMDSQL + " HUBUNGAN ='" + .Combo1.Text + "' "
        CMDSQL = CMDSQL + " WHERE MGM_ID ='" + ListView1(0).SelectedItem.SubItems(3) + "'"
        M_OBJCONN.Execute CMDSQL
        Call show_gridref
    End With
    Unload frmcust_cc_ref
Case 2
    If ListView1(0).ListItems.Count = 0 Then
        Exit Sub
    End If
    
    LMSGBOX = MsgBox("Yakin Akan Dihapus??", vbOKCancel + vbExclamation, "Telegrandi")
    If LMSGBOX = vbOK Then
        M_OBJCONN.Execute "Delete from MGM_REF where mgm_id ='" + ListView1(0).SelectedItem.SubItems(3) + "' "
        ListView1(0).ListItems.Remove ListView1(0).SelectedItem.Index
    End If
End Select
Exit Sub

erradd:
    MsgBox Err.Number + Err.Description
End Sub

Private Sub Form_Terminate()
    HAK_TeamLeader = False
    SCREENER_APPROV = False
    ID_CUST = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    HAK_TeamLeader = False
    SCREENER_APPROV = False
    ID_CUST = ""
End Sub

Private Sub ListView1_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Select Case Index
Case 1
    ListView1(Index).SortKey = ColumnHeader.Index - 1
   ListView1(Index).Sorted = True
End Select
End Sub

Private Sub Option2_Click(Index As Integer)
Select Case Index
    Case 0
        Frame11.Enabled = True
    Case 1
        Frame11.Enabled = False
End Select
End Sub

Private Sub Option4_Click(Index As Integer)
Select Case Index
    Case 0
        Frame7.Enabled = True
        Frame8.Enabled = False
        Text1(18).Text = Empty
        Text1(34).Text = Empty
        RichTextBox1(4).Text = Empty
    Case 1
        Frame7.Enabled = False
        Frame8.Enabled = True
        Text1(20).Text = Empty
        Text1(30).Text = Empty
        Combo1(16).Text = Empty
        RichTextBox1(2).Text = Empty
End Select
End Sub


Private Function CEK_DATA_VALID() As Boolean
Dim M_MSGBOX As Variant
    If Len(TDBMask1(0).Value) < 3 And Len(TDBMask1(1).Value) < 3 And Len(TDBMask1(2).Value) < 3 And Len(TDBMask1(3).Value) < 3 And Len(TDBMask1(4).Value) < 3 And Len(TDBMask1(5).Value) < 3 And Len(TDBMask1(6).Value) < 3 And Len(TDBMask1(7).Value) < 3 Then
        CEK_DATA_VALID = False
        MsgBox "Minimal Satu Nomor Telpon Harus Diisi", vbCritical + vbOKOnly, "Peringatan"
        SSTab1.Tab = 0
        TDBMask1(0).SetFocus
        Exit Function
    End If
    If Text1(3).Text = Empty Then
        CEK_DATA_VALID = False
        MsgBox "Sumber Info Harus Diisi", vbCritical + vbOKOnly, "Peringatan"
'        Combo1(0).SetFocus
        Exit Function
    End If
    If Text1(0).Text = Empty Then
        CEK_DATA_VALID = False
        MsgBox "Nama Harus Diisi", vbCritical + vbOKOnly, "Peringatan"
        Text1(0).SetFocus
        Exit Function
    End If
    If Combo1(2).Text = Empty Then
        CEK_DATA_VALID = False
        MsgBox "Title Harus Diisi", vbCritical + vbOKOnly, "Peringatan"
        Combo1(2).SetFocus
        Exit Function
    End If
    If Option2(0).Value Then
    End If
    If Check2(1).Value = 1 Then
        If Combo3(1).Text = Empty Then
            CEK_DATA_VALID = False
            MsgBox "Clossing Reason Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
            SSTab1.Tab = 4
'            Combo3(1).SetFocus
            Exit Function
        End If
    End If
    If Check2(1).Value = 1 Then
        RichTextBox1(3).Text = Combo3(1).Text
    Else
        If RichTextBox1(3).Text = Empty And Text2.Text = Empty Then
            CEK_DATA_VALID = False
            MsgBox "Catatan(Perubahan Pada Data Customer Ini) Atau Next Action Harus Diisi...!!!", vbCritical + vbOKOnly, "Peringatan"
            SSTab1.Tab = 4
            RichTextBox1(3).SetFocus
            Exit Function
        End If
    End If
CEK_DATA_VALID = True
End Function

Private Sub RichTextBox1_LostFocus(Index As Integer)
Select Case Index
Case 3
    RichTextBox1(3).Text = UCase(RichTextBox1(3).Text)
End Select
End Sub

Private Sub ssCommand1_Click(Index As Integer)
Dim M_MSGBOX As Variant
Dim V_SAVE As Boolean
V_SAVE = True
Select Case Index
    Case 0
        ID_CUST = Text1(1).Text
        frmtelp.Text11.Text = Text1(0).Text
        frmtelp.Text10.Text = Combo1(0).Text
        REMO = True
        frmtelp_mgm.Show vbModal
    Case 1
        MsgBox "PEMBERI REFERENSI", vbOKOnly
    Case 2
        V_SAVE = CEK_DATA_VALID
        If V_SAVE = False Then
            Exit Sub
        Else
        End If
        Call CEK_UPDATE_PELANGGAN
    Case 3
        HAK_TeamLeader = False
        Unload Me
End Select
End Sub

Private Sub CEK_UPDATE_PELANGGAN()
Dim M_DATA As New CLS_FRMCUST_CC_MGM
Dim M_DATA1 As New CLS_CARI_HISTORY_CC_mgm
Dim M_OBJRS As ADODB.Recordset
Dim CMDSQL As String
Dim M_MSGBOX As Variant
Dim M_CALL As String
Dim M_STATUS As String

M_CALL = "1"

Call M_DATA1.CARI_STATUS_CUSTOMER(MDIForm1.Text3.Text)


SCREENER_APPROV = False
If HAK_TeamLeader Then
    M_MSGBOX = MsgBox("Apakah Anda Team Leader", vbInformation + vbYesNo, "Konfirmasi")
    If M_MSGBOX = vbYes Then
        FRMPASWORD.Show vbModal
        If HAK_TeamLeader = False Then
            M_DATA.UPDATE_CUSTOMER_BARU M_OBJCONN, KETHSLKERJA, STATUS_FIELD_LAMA, M_CALL, M_STATUS, DOK1
            If Check2(1).Value = 1 Then
                If RichTextBox1(3).Text <> Empty Or Text2.Text <> Empty Then
                    M_DATA.ADD_HISTORY M_OBJCONN, Text1(1).Text, MDIForm1.TDBDate1.Text, Time, MDIForm1.Text1.Text, "MGM", Text2.Text & " " & RichTextBox1(3).Text, Combo1(0).Text
                End If
            End If
'            M_DATA.ADD_HISTORY M_OBJCONN, Text1(1).Text, MDIForm1.TDBDate1.Text, Time, MDIForm1.Text1.Text, Combo2(0).Text, Text2.Text & " " & RichTextBox1(3).Text
        MsgBox "Data Sudah Tersimpan", vbInformation + vbOKOnly, "Sukses"
        RichTextBox1(3).Text = Empty
        Text2.Text = Empty
'        Unload Me
            Exit Sub
        End If
        Exit Sub
    Else
        MsgBox "Hubungi Team Leader Anda..!!!", vbInformation + vbOKOnly, "Informasi"
        FRM_DATASAMA_CC.Show vbModal
        Exit Sub
    End If
End If

If UCase(MDIForm1.Text2.Text) = "AGENT" Then
    Call cek_update_telp_sama
End If

If HAK_TeamLeader = True Then
    Exit Sub
End If

If UCase(MDIForm1.Text3.Text) = "ADMIN" Then
    M_STATUS = 1
    M_CALL = 0
End If

M_DATA.UPDATE_CUSTOMER_BARU M_OBJCONN, KETHSLKERJA, STATUS_FIELD_LAMA, M_CALL, M_STATUS, DOK1
If Check2(1).Value = 1 Then
    If RichTextBox1(3).Text <> Empty Or Text2.Text <> Empty Then
        M_DATA.ADD_HISTORY M_OBJCONN, Text1(1).Text, MDIForm1.TDBDate1.Text, Time, MDIForm1.Text1.Text, "MGM", Text2.Text & " " & RichTextBox1(3).Text, Combo1(0).Text
    End If
Else
        M_DATA.ADD_HISTORY M_OBJCONN, Text1(1).Text, MDIForm1.TDBDate1.Text, Time, MDIForm1.Text1.Text, "MGM", Text2.Text & " " & RichTextBox1(3).Text, Combo1(0).Text
End If
MsgBox "Data Sudah Tersimpan", vbInformation + vbOKOnly, "Sukses"
RichTextBox1(3).Text = Empty
Text2.Text = Empty
Set M_DATA = Nothing
Set M_DATA1 = Nothing
End Sub
            
Private Sub TDBDate1_Click(Index As Integer)
Dim tahun As Integer
Dim tahunlhr As Integer
Select Case Index
Case 0
    tahun = Year(Date)
    If TDBDate1(0).ValueIsNull Then
        Text1(2).Text = "0"
    Else
        tahunlhr = Year(TDBDate1(0).Value)
        Text1(2).Text = CStr(tahun - tahunlhr)
    End If
End Select
End Sub

Private Sub TDBDate1_CloseUp(Index As Integer, Cancel As Boolean, Value As Date, Escape As Boolean)
Dim tahun As Integer
Dim tahunlhr As Integer
Select Case Index
Case 0
    tahun = Year(Date)
    If TDBDate1(0).ValueIsNull Then
        Text1(2).Text = "0"
    Else
        tahunlhr = Year(TDBDate1(0).Value)
        Text1(2).Text = CStr(tahun - tahunlhr)
    End If
End Select
End Sub

Private Sub Form_Load()
Dim LISTITEM As LISTITEM
Dim M_OBJRS As ADODB.Recordset
Dim M_DATA As New CLS_FRMCUST_CC_MGM
    CALL_OK = False
    ADD_CUST_REM = False
    ID_CUST = ""
    Call ChangeTab(SSTab1)
    TDBDate1(3).Value = MDIForm1.TDBDate1.Value
    TDBTime1.Value = Format(Time, "hh:mm")
    Option4(0).Value = True
    Call Isi_Combo
    Call HEADER_HISTORY
    Call ISI_COMBO_DATASOURCE
    Call ISI_COMBO_PRODUCT_CLOSING
    If ADD_CUST Then
        ADD_CUST_REM = True
    Else
        VIEW_ADD = False
        Call VIEW_DATA_CUST
'            Check2(3).Visible = False
'            Frame23.Visible = False
'            Check2(3).Enabled = False
'            Frame23.Enabled = False
        If UCase(MDIForm1.Text2.Text) = "AGENT" Then
            Text1(0).Locked = True
            Combo1(0).Visible = False
            Combo1(1).Visible = False
            Text1(3).Visible = True
            If Len(TDBMask1(0).Value) > 5 Then
                TDBMask1(0).ReadOnly = True
            End If
            If Len(TDBMask1(1).Value) > 5 Then
                TDBMask1(1).ReadOnly = True
            End If
            If Len(TDBMask1(2).Value) > 5 Then
                TDBMask1(2).ReadOnly = True
            End If
            If Len(TDBMask1(3).Value) > 5 Then
                TDBMask1(3).ReadOnly = True
            End If
            If Len(TDBMask1(4).Value) > 5 Then
                TDBMask1(4).ReadOnly = True
            End If
            If Len(TDBMask1(5).Value) > 5 Then
                TDBMask1(5).ReadOnly = True
            End If
            If Len(TDBMask1(6).Value) > 5 Then
                TDBMask1(6).ReadOnly = True
            End If
            If Len(TDBMask1(7).Value) > 5 Then
                TDBMask1(7).ReadOnly = True
            End If
        End If
    End If

Set M_OBJRS = Nothing
Set M_DATA = Nothing
SSTab1.Tab = 0
End Sub

Private Sub HEADER_HISTORY()
    ListView1(1).ColumnHeaders.ADD 1, , "Tanggal Jam", 15 * TXT
'    ListView1(1).ColumnHeaders.ADD 2, , "Jam", 8 * TXT
    ListView1(1).ColumnHeaders.ADD 2, , "History", 50 * TXT
    ListView1(1).ColumnHeaders.ADD 3, , "User Update", 20 * TXT
    
    ListView1(0).ColumnHeaders.ADD 1, , "Nama", 15 * TXT
    ListView1(0).ColumnHeaders.ADD 2, , "Telp Kantor", 10 * TXT
    ListView1(0).ColumnHeaders.ADD 3, , "Hubungan", 20 * TXT
    ListView1(0).ColumnHeaders.ADD 4, , "Id", 1 * TXT

End Sub

Private Sub VIEW_DATA_CUST()
Dim M_OBJRS As ADODB.Recordset
Dim m_objrs1 As ADODB.Recordset
Dim M_DATA As New CLS_FRMCUST_CC_MGM
Dim M_DATA1 As New CLS_CARI_HISTORY_CC_mgm
Dim LISTITEM As LISTITEM
Dim M_CAT As String
Dim M_QUALIFIED As String
Dim m_balance As String
Dim CMDSQL As String

Set M_OBJRS = M_DATA.QUERY_CUST(M_OBJCONN, "CUSTID = '" + VIEW_MGMDATA.ListView1.SelectedItem.SubItems(1) + "'")
        If M_OBJRS.RecordCount <> 0 Then
            Text1(1).Text = IIf(IsNull(M_OBJRS("CUSTID")), "", M_OBJRS("CUSTID"))
            ID_CUST = IIf(IsNull(M_OBJRS("CUSTID")), "", M_OBJRS("CUSTID"))
            M_DATA1.VIEW_HISTORY_KERJA IIf(IsNull(M_OBJRS("KETHSLKERJA")), "", M_OBJRS("KETHSLKERJA")), " ", MDIForm1.Text3.Text
            Text1(0).Text = IIf(IsNull(M_OBJRS("NAME")), "", M_OBJRS("NAME"))
            Combo1(2).Text = IIf(IsNull(M_OBJRS("TITLE")), "", M_OBJRS("TITLE"))
            Call Combo1_LostFocus(2)
            Combo3(0).Text = IIf(IsNull(M_OBJRS("KD_CLS")), "", M_OBJRS("KD_CLS"))
            Call Combo3_LostFocus(0)
            TDBDate1(0).Value = IIf(IsNull(M_OBJRS("BIRTHD")), "", Format(M_OBJRS("BIRTHD"), "dd-mmm-yyyy"))
            Call TDBDate1_Click(0)
            RichTextBox1(0).Text = IIf(IsNull(M_OBJRS("ADDRNOW")), "", M_OBJRS("ADDRNOW"))
            Text1(6).Text = IIf(IsNull(M_OBJRS("ZIPNOW")), "", M_OBJRS("ZIPNOW"))
            Text1(7).Text = IIf(IsNull(M_OBJRS("CITYNOW")), "", M_OBJRS("CITYNOW"))
            TDBMask1(0).Value = IIf(IsNull(M_OBJRS("HOMENO")), "", M_OBJRS("HOMENO"))
            TDBMask1(1).Value = IIf(IsNull(M_OBJRS("HOMENO2")), "", M_OBJRS("HOMENO2"))
            TDBMask1(2).Value = IIf(IsNull(M_OBJRS("MOBILENO")), "", M_OBJRS("MOBILENO"))
            TDBMask1(3).Value = IIf(IsNull(M_OBJRS("MOBILENO2")), "", M_OBJRS("MOBILENO2"))
            TDBMask1(4).Value = IIf(IsNull(M_OBJRS("OFFICENO")), "", M_OBJRS("OFFICENO"))
            Text1(21).Text = IIf(IsNull(M_OBJRS("EXTOFFICE")), "", M_OBJRS("EXTOFFICE"))
            TDBMask1(5).Value = IIf(IsNull(M_OBJRS("OFFICENO2")), "", M_OBJRS("OFFICENO2"))
            Text1(32).Text = IIf(IsNull(M_OBJRS("EXTOFFICE2")), "", M_OBJRS("EXTOFFICE2"))
            TDBMask1(6).Value = IIf(IsNull(M_OBJRS("FAXNO")), "", M_OBJRS("FAXNO"))
            TDBMask1(7).Value = IIf(IsNull(M_OBJRS("FAXNO2")), "", M_OBJRS("FAXNO2"))
            M_CAT = IIf(IsNull(M_OBJRS("CAT")), "0", M_OBJRS("CAT"))
            Combo5.Text = IIf(IsNull(M_OBJRS("PRIOR")), "", M_OBJRS("PRIOR"))
            If M_CAT = "0" Then
                Option4(0).Value = True
                Text1(20).Text = IIf(IsNull(M_OBJRS("NAMAPT")), "", M_OBJRS("NAMAPT"))
                RichTextBox1(2).Text = IIf(IsNull(M_OBJRS("ADDRPT")), "", M_OBJRS("ADDRPT"))
            Else
                Option4(1).Value = True
                Text1(34).Text = IIf(IsNull(M_OBJRS("JENISUSAHA")), "", M_OBJRS("JENISUSAHA"))
                Text1(18).Text = IIf(IsNull(M_OBJRS("NAMAPT")), "", M_OBJRS("NAMAPT"))
                RichTextBox1(4).Text = IIf(IsNull(M_OBJRS("ADDRPT")), "", M_OBJRS("ADDRPT"))
            End If
            Combo1(0).Text = IIf(IsNull(M_OBJRS("RECSOURCE")), "", M_OBJRS("RECSOURCE"))
            Call Combo1_Click(0)
            Text3.Text = IIf(IsNull(M_OBJRS("OTHERS")), "", M_OBJRS("OTHERS"))
            TDBMask1(8).Value = IIf(IsNull(M_OBJRS("AHOMENO")), "", M_OBJRS("AHOMENO"))
            TDBMask1(9).Value = IIf(IsNull(M_OBJRS("AHOMENO2")), "", M_OBJRS("AHOMENO2"))
            TDBMask1(10).Value = IIf(IsNull(M_OBJRS("AOFFICENO")), "", M_OBJRS("AOFFICENO"))
            TDBMask1(11).Value = IIf(IsNull(M_OBJRS("AOFFICENO2")), "", M_OBJRS("AOFFICENO2"))
            TDBMask1(12).Value = IIf(IsNull(M_OBJRS("AFAXNO")), "", M_OBJRS("AFAXNO"))
            TDBMask1(13).Value = IIf(IsNull(M_OBJRS("AFAXNO2")), "", M_OBJRS("AFAXNO2"))
        End If
Set M_OBJRS = Nothing

Set m_objrs1 = New ADODB.Recordset
If SCR_SPV_CARI = True Then
    Set m_objrs1 = M_DATA.QUERY_HIST_CUST(M_OBJCONN, "CUSTID = '" + Text1(1).Text + "'")
Else
    If SCREENER_APPROV = True Then
        Set m_objrs1 = M_DATA.QUERY_HIST_CUST(M_OBJCONN, "CUSTID = '" + Text1(1).Text + "'")
    Else
        Set m_objrs1 = M_DATA.QUERY_HIST_CUST(M_OBJCONN, "CUSTID = '" + Text1(1).Text + "'")
    End If
End If

While Not m_objrs1.EOF
    Set LISTITEM = ListView1(1).ListItems.ADD(, , Left(m_objrs1("DATETIME"), 4) & "/" & Mid(m_objrs1("DATETIME"), 5, 2) & "/" & IIf(IsNull(m_objrs1("DATETIME")), "", Mid(m_objrs1("DATETIME"), 7, 2)) & " " & IIf(IsNull(m_objrs1("DATETIME")), "", Mid(m_objrs1("DATETIME"), 9, 2)) & ":" & Right(m_objrs1("DATETIME"), 2))
        LISTITEM.SubItems(1) = IIf(IsNull(m_objrs1("HST")), "", m_objrs1("HST"))
        LISTITEM.SubItems(2) = IIf(IsNull(m_objrs1("AGENT")), "", m_objrs1("AGENT"))
m_objrs1.MoveNext
Wend
Set m_objrs1 = Nothing
Set M_DATA = Nothing
Call show_gridref
End Sub

Private Sub show_gridref()
Dim m_objrs1 As New ADODB.Recordset
Dim LISTITEM As LISTITEM
ListView1(0).ListItems.Clear
m_objrs1.CursorLocation = adUseClient
m_objrs1.Open "Select * from MGM_REF WHERE CUSTID ='" + Text1(1).Text + "' order by mgm_id", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not m_objrs1.EOF
    Set LISTITEM = ListView1(0).ListItems.ADD(, , IIf(IsNull(m_objrs1("NAMA")), "", m_objrs1("NAMA")))
        LISTITEM.SubItems(1) = IIf(IsNull(m_objrs1("OFFICENO")), "", m_objrs1("OFFICENO"))
        LISTITEM.SubItems(2) = IIf(IsNull(m_objrs1("HUBUNGAN")), "", m_objrs1("HUBUNGAN"))
        LISTITEM.SubItems(3) = IIf(IsNull(m_objrs1("mgm_id")), "", m_objrs1("mgm_id"))
m_objrs1.MoveNext
Wend
Set m_objrs1 = Nothing

End Sub

Private Sub ISI_COMBO_DATASOURCE()
Dim M_OBJRS As ADODB.Recordset
Dim M_DATA As New CLS_FRMCUST_CC_MGM
    Set M_OBJRS = M_DATA.QUERY_COMBO_DATASOURCE_ISI(M_OBJCONN, "")
    While Not M_OBJRS.EOF
        Combo1(0).AddItem M_OBJRS("KODEDS")
        Combo1(0).DataField = M_OBJRS("KODEDS")
        Combo1(1).AddItem IIf(IsNull(M_OBJRS("KETERANGAN")), "", M_OBJRS("KETERANGAN"))
        Combo1(1).DataField = M_OBJRS("KETERANGAN")
        M_OBJRS.MoveNext
    Wend
Set M_OBJRS = Nothing
Set M_DATA = Nothing
End Sub

Private Sub ISI_COMBO_PRODUCT_CLOSING()
Dim M_OBJRS As ADODB.Recordset
Dim M_DATA As New CLS_FRMCUST_CC_MGM
    Set M_OBJRS = M_DATA.QUERY_COMBO_CLOSSING(M_OBJCONN, "")
    While Not M_OBJRS.EOF
        Combo3(0).AddItem M_OBJRS("KDCLS")
        Combo3(0).DataField = M_OBJRS("KDCLS")
        Combo3(1).AddItem M_OBJRS("KETCLS")
        Combo3(1).DataField = M_OBJRS("KETCLS")
        M_OBJRS.MoveNext
    Wend
    Set M_OBJRS = Nothing
Set M_DATA = Nothing
End Sub

Private Sub Isi_Combo()
' ISI COMBO TITLE
        Combo1(2).AddItem "Bapak"
        Combo1(2).AddItem "Ibu"
        Combo5.AddItem "Low"
        Combo5.AddItem "Normal"
        Combo5.AddItem "High"
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Call ChangeTab(SSTab1)
End Sub

Private Sub ChangeTab(SSTab As SSTab)
    Dim ctrl As Control, TabIndex As Long
    TabIndex = 99999          ' A very high value.
    On Error Resume Next
    For Each ctrl In SSTab.Parent.Controls
        If ctrl.Container Is SSTab Then
            If ctrl.Left < -10000 Then
                ctrl.Enabled = False
            Else
                ctrl.Enabled = True
                If ctrl.TabIndex >= TabIndex Then
                Else
                    TabIndex = ctrl.TabIndex
                    ctrl.SetFocus
                End If
            End If
        End If
    Next
End Sub

Private Sub TDBDate1_LostFocus(Index As Integer)
Dim tahun As Integer
Dim tahunlhr As Integer
Select Case Index
Case 0
    tahun = Year(Date)
    If TDBDate1(0).ValueIsNull Then
        Text1(2).Text = "0"
    Else
        tahunlhr = Year(TDBDate1(0).Value)
        Text1(2).Text = CStr(tahun - tahunlhr)
    End If
Case 3
If TDBDate1(Index).ReadOnly = True Then
Else
    If TDBDate1(Index).ValueIsNull Then
        TDBDate1(Index).Value = Format(Date, "dd-mmm-yyyy")
        TDBDate1(Index).Value = TDBDate1(Index).Value + 1
    End If
End If
End Select
End Sub

Private Sub TDBMask1_GotFocus(Index As Integer)
Select Case Index
Case 0 To 13
    SendKeys "{Home}+{End}"
End Select
End Sub

Private Sub TDBTime1_LostFocus()
    If TDBTime1.ValueIsNull Then
        TDBTime1.Value = Format(Time, "hh:mm")
    End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Select Case Index
Case 21, 32
    SendKeys "{Home}+{End}"
End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
    Case 2, 6, 8, 10, 11, 12, 13, 14, 15, 16, 17, 22, 23, 24, 25, 26, 27, 28, 29, 21, 32
        Select Case KeyAscii
            Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 8
                Exit Sub
            Case Else
                KeyAscii = 0
        End Select
End Select
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim M_OBJRS As ADODB.Recordset
Select Case Index
    Case 0
        Text1(Index).Text = UCase(Text1(Index).Text)
    Case 4
        If Len(Text1(Index).Text) > 2 Then
            Set M_OBJRS = New ADODB.Recordset
            M_OBJRS.CursorLocation = adUseClient
            M_OBJRS.Open "SELECT NOLAP FROM CC_CUSTTBL WHERE NOLAP = '" + Text1(Index).Text + "' AND  CUSTID <> '" + Text1(1).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If M_OBJRS.RecordCount <> 0 Then
                MsgBox "Nomor Yang Anda Masukan Telah Ada", vbInformation, "TeleGrandi"
                Text1(Index).Text = Empty
            End If
        End If
End Select
Set M_OBJRS = Nothing
End Sub

Private Sub Text2_LostFocus()
    Text2.Text = UCase(Text2.Text)
End Sub


Private Sub cek_update_telp_sama()
Dim CMDSQL As String
Dim M_DATA As New CLS_FRMCUST_CC_MGM
Dim M_OBJRS As ADODB.Recordset
Dim M_MSGBOX As Variant
        If TDBMask1(0).ReadOnly = False Then
            If Len(TDBMask1(0).Value) > 4 Then
                CMDSQL = "(HOMENO = '" + TDBMask1(0).Value + "'"
                CMDSQL = CMDSQL + " OR HOMENO2 = '" + TDBMask1(0).Value + "'"
                CMDSQL = CMDSQL + " OR MOBILENO = '" + TDBMask1(0).Value + "'"
                CMDSQL = CMDSQL + " OR MOBILENO2 = '" + TDBMask1(0).Value + "'"
                CMDSQL = CMDSQL + " OR FAXNO = '" + TDBMask1(0).Value + "'"
                CMDSQL = CMDSQL + " OR FAXNO2 = '" + TDBMask1(0).Value + "'"
                CMDSQL = CMDSQL + " OR OFFICENO = '" + TDBMask1(0).Value + "'"
                CMDSQL = CMDSQL + " OR OFFICENO2 = '" + TDBMask1(0).Value + "'"
            End If
        End If
        If TDBMask1(1).ReadOnly = False Then
            If Len(TDBMask1(1).Value) > 4 Then
                If CMDSQL = Empty Then
                    CMDSQL = CMDSQL + " (HOMENO = '" + TDBMask1(1).Value + "'"
                Else
                    CMDSQL = CMDSQL + " OR HOMENO = '" + TDBMask1(1).Value + "'"
                End If
                    CMDSQL = CMDSQL + " OR HOMENO2 = '" + TDBMask1(1).Value + "'"
                    CMDSQL = CMDSQL + " OR MOBILENO = '" + TDBMask1(1).Value + "'"
                    CMDSQL = CMDSQL + " OR MOBILENO2 = '" + TDBMask1(1).Value + "'"
                    CMDSQL = CMDSQL + " OR FAXNO = '" + TDBMask1(1).Value + "'"
                    CMDSQL = CMDSQL + " OR FAXNO2 = '" + TDBMask1(1).Value + "'"
                    CMDSQL = CMDSQL + " OR OFFICENO = '" + TDBMask1(1).Value + "'"
                    CMDSQL = CMDSQL + " OR OFFICENO2 = '" + TDBMask1(1).Value + "'"
            End If
        End If
        If TDBMask1(2).ReadOnly = False Then
            If Len(TDBMask1(2).Value) > 4 Then
                If CMDSQL = Empty Then
                    CMDSQL = CMDSQL + " (HOMENO = '" + TDBMask1(2).Value + "'"
                Else
                    CMDSQL = CMDSQL + " OR HOMENO = '" + TDBMask1(2).Value + "'"
                End If
                    CMDSQL = CMDSQL + " OR HOMENO2 = '" + TDBMask1(2).Value + "'"
                    CMDSQL = CMDSQL + " OR MOBILENO = '" + TDBMask1(2).Value + "'"
                    CMDSQL = CMDSQL + " OR MOBILENO2 = '" + TDBMask1(2).Value + "'"
                    CMDSQL = CMDSQL + " OR FAXNO = '" + TDBMask1(2).Value + "'"
                    CMDSQL = CMDSQL + " OR FAXNO2 = '" + TDBMask1(2).Value + "'"
                    CMDSQL = CMDSQL + " OR OFFICENO = '" + TDBMask1(2).Value + "'"
                    CMDSQL = CMDSQL + " OR OFFICENO2 = '" + TDBMask1(2).Value + "'"
            End If
        End If
        If TDBMask1(3).ReadOnly = False Then
            If Len(TDBMask1(3).Value) > 4 Then
                If CMDSQL = Empty Then
                    CMDSQL = CMDSQL + " (HOMENO = '" + TDBMask1(3).Value + "'"
                Else
                    CMDSQL = CMDSQL + " OR HOMENO = '" + TDBMask1(3).Value + "'"
                End If
                    CMDSQL = CMDSQL + " OR HOMENO2 = '" + TDBMask1(3).Value + "'"
                    CMDSQL = CMDSQL + " OR MOBILENO = '" + TDBMask1(3).Value + "'"
                    CMDSQL = CMDSQL + " OR MOBILENO2 = '" + TDBMask1(3).Value + "'"
                    CMDSQL = CMDSQL + " OR FAXNO = '" + TDBMask1(3).Value + "'"
                    CMDSQL = CMDSQL + " OR FAXNO2 = '" + TDBMask1(3).Value + "'"
                    CMDSQL = CMDSQL + " OR OFFICENO = '" + TDBMask1(3).Value + "'"
                    CMDSQL = CMDSQL + " OR OFFICENO2 = '" + TDBMask1(3).Value + "'"
            End If
        End If
        If TDBMask1(4).ReadOnly = False Then
            If Len(TDBMask1(4).Value) > 4 Then
                If CMDSQL = Empty Then
                    CMDSQL = CMDSQL + " (HOMENO = '" + TDBMask1(4).Value + "'"
                Else
                    CMDSQL = CMDSQL + " OR HOMENO = '" + TDBMask1(4).Value + "'"
                End If
                    CMDSQL = CMDSQL + " OR HOMENO = '" + TDBMask1(4).Value + "'"
                    CMDSQL = CMDSQL + " OR HOMENO2 = '" + TDBMask1(4).Value + "'"
                    CMDSQL = CMDSQL + " OR MOBILENO = '" + TDBMask1(4).Value + "'"
                    CMDSQL = CMDSQL + " OR MOBILENO2 = '" + TDBMask1(4).Value + "'"
                    CMDSQL = CMDSQL + " OR FAXNO = '" + TDBMask1(4).Value + "'"
                    CMDSQL = CMDSQL + " OR FAXNO2 = '" + TDBMask1(4).Value + "'"
                    CMDSQL = CMDSQL + " OR OFFICENO = '" + TDBMask1(4).Value + "'"
                    CMDSQL = CMDSQL + " OR OFFICENO2 = '" + TDBMask1(4).Value + "'"
            End If
        End If
        If TDBMask1(5).ReadOnly = False Then
            If Len(TDBMask1(5).Value) > 4 Then
                If CMDSQL = Empty Then
                    CMDSQL = CMDSQL + " (HOMENO = '" + TDBMask1(5).Value + "'"
                Else
                    CMDSQL = CMDSQL + " OR HOMENO = '" + TDBMask1(5).Value + "'"
                End If
                    CMDSQL = CMDSQL + " OR HOMENO2 = '" + TDBMask1(5).Value + "'"
                    CMDSQL = CMDSQL + " OR MOBILENO = '" + TDBMask1(5).Value + "'"
                    CMDSQL = CMDSQL + " OR MOBILENO2 = '" + TDBMask1(5).Value + "'"
                    CMDSQL = CMDSQL + " OR FAXNO = '" + TDBMask1(5).Value + "'"
                    CMDSQL = CMDSQL + " OR FAXNO2 = '" + TDBMask1(5).Value + "'"
                    CMDSQL = CMDSQL + " OR OFFICENO = '" + TDBMask1(5).Value + "'"
                    CMDSQL = CMDSQL + " OR OFFICENO2 = '" + TDBMask1(5).Value + "'"
        End If
        End If
        If TDBMask1(6).ReadOnly = False Then
            If Len(TDBMask1(6).Value) > 4 Then
                If CMDSQL = Empty Then
                    CMDSQL = CMDSQL + " (HOMENO = '" + TDBMask1(6).Value + "'"
                Else
                    CMDSQL = CMDSQL + " OR HOMENO = '" + TDBMask1(6).Value + "'"
                End If
                    CMDSQL = CMDSQL + " OR HOMENO2 = '" + TDBMask1(6).Value + "'"
                    CMDSQL = CMDSQL + " OR MOBILENO = '" + TDBMask1(6).Value + "'"
                    CMDSQL = CMDSQL + " OR MOBILENO2 = '" + TDBMask1(6).Value + "'"
                    CMDSQL = CMDSQL + " OR FAXNO = '" + TDBMask1(6).Value + "'"
                    CMDSQL = CMDSQL + " OR FAXNO2 = '" + TDBMask1(6).Value + "'"
                    CMDSQL = CMDSQL + " OR OFFICENO = '" + TDBMask1(6).Value + "'"
                    CMDSQL = CMDSQL + " OR OFFICENO2 = '" + TDBMask1(6).Value + "'"
        End If
        End If
        If TDBMask1(7).ReadOnly = False Then
            If Len(TDBMask1(7).Value) > 4 Then
                If CMDSQL = Empty Then
                    CMDSQL = CMDSQL + " (HOMENO = '" + TDBMask1(7).Value + "'"
                Else
                    CMDSQL = CMDSQL + " or HOMENO = '" + TDBMask1(7).Value + "'"
                End If
                    CMDSQL = CMDSQL + " OR HOMENO2 = '" + TDBMask1(7).Value + "'"
                    CMDSQL = CMDSQL + " OR MOBILENO = '" + TDBMask1(7).Value + "'"
                    CMDSQL = CMDSQL + " OR MOBILENO2 = '" + TDBMask1(7).Value + "'"
                    CMDSQL = CMDSQL + " OR FAXNO = '" + TDBMask1(7).Value + "'"
                    CMDSQL = CMDSQL + " OR FAXNO2 = '" + TDBMask1(7).Value + "'"
                    CMDSQL = CMDSQL + " OR OFFICENO = '" + TDBMask1(7).Value + "'"
                    CMDSQL = CMDSQL + " OR OFFICENO2 = '" + TDBMask1(7).Value + "'"
        End If
        End If
        If Len(CMDSQL) <> 0 Then
            CMDSQL = CMDSQL + ") AND CUSTID <> '" + Text1(1).Text + "'"
            Set M_OBJRS = M_DATA.QUERY_CEK_ADDCUST(M_OBJCONN, CMDSQL)
            If M_OBJRS.RecordCount >= 3 Or M_OBJRS.RecordCount < 1 Then
                HAK_TeamLeader = False
                Set M_OBJRS = Nothing
                Exit Sub
            Else
                HAK_TeamLeader = True
                update_TELP_SAMA = True
                MsgBox "Update Gagal... No Telepon Ada Yang Sama... Hubungi Team Leader Untuk Menyimpan Data ", vbCritical + vbOKOnly, "Peringatan"
                Set M_OBJRS = Nothing
                FRM_DATASAMA_CC.Show vbModal
            Exit Sub
            End If
        End If
Set M_OBJRS = Nothing
End Sub
