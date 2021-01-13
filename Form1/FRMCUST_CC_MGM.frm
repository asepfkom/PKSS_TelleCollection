VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FRMCUST_CC_MGM 
   Caption         =   "MgM Data"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   10065
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   Icon            =   "FRMCUST_CC_MGM.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5805
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSCommand SSCommand1 
      Cancel          =   -1  'True
      Height          =   420
      Index           =   3
      Left            =   8715
      TabIndex        =   3
      Top             =   15
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   741
      _Version        =   196610
      Font3D          =   2
      MousePointer    =   16
      ForeColor       =   128
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
      Picture         =   "FRMCUST_CC_MGM.frx":0442
      Caption         =   "&Exit"
      AutoSize        =   2
      Alignment       =   4
      ButtonStyle     =   2
      PictureAlignment=   1
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   420
      Index           =   2
      Left            =   7605
      TabIndex        =   2
      Top             =   15
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   741
      _Version        =   196610
      Font3D          =   2
      MousePointer    =   16
      ForeColor       =   128
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
      Picture         =   "FRMCUST_CC_MGM.frx":059C
      Caption         =   "&Save"
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
      Top             =   15
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   741
      _Version        =   196610
      Font3D          =   2
      MousePointer    =   16
      ForeColor       =   128
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
      Picture         =   "FRMCUST_CC_MGM.frx":08BE
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
      Top             =   15
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   741
      _Version        =   196610
      Font3D          =   2
      MousePointer    =   16
      ForeColor       =   128
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
      Picture         =   "FRMCUST_CC_MGM.frx":1132
      Caption         =   "&Add Ref"
      AutoSize        =   2
      Alignment       =   4
      ButtonStyle     =   2
      PictureAlignment=   1
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   150
      TabIndex        =   10
      Top             =   450
      Width           =   9735
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Left            =   6840
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   135
         Width           =   2820
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
         Left            =   840
         MaxLength       =   200
         TabIndex        =   4
         Top             =   150
         Width           =   3900
      End
      Begin VB.Label Label1 
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
         Left            =   6390
         TabIndex        =   12
         Top             =   195
         Width           =   345
      End
      Begin VB.Label Label1 
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
         TabIndex        =   11
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
      Left            =   3150
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   4005
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
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   960
      Width           =   1665
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   885
      Visible         =   0   'False
      Width           =   5670
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   60
      TabIndex        =   13
      Top             =   1350
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
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
      TabCaption(0)   =   "Personal Data"
      TabPicture(0)   =   "FRMCUST_CC_MGM.frx":15FA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "History"
      TabPicture(1)   =   "FRMCUST_CC_MGM.frx":1616
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame9"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Results"
      TabPicture(2)   =   "FRMCUST_CC_MGM.frx":1632
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame10"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Leads"
      TabPicture(3)   =   "FRMCUST_CC_MGM.frx":164E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame6 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   107
         Top             =   360
         Width           =   9735
         Begin MSComctlLib.ListView ListView1 
            Height          =   3645
            Index           =   0
            Left            =   120
            TabIndex        =   108
            Top             =   240
            Width           =   9555
            _ExtentX        =   16854
            _ExtentY        =   6429
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
         Height          =   4035
         Left            =   375
         TabIndex        =   58
         Top             =   345
         Width           =   10200
         Begin VB.CheckBox Check2 
            Caption         =   "Disagree"
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
            Index           =   3
            Left            =   4680
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   103
            Top             =   1080
            Width           =   1590
         End
         Begin VB.Frame frame5 
            Height          =   615
            Left            =   4560
            TabIndex        =   102
            Top             =   1080
            Width           =   4215
            Begin VB.ComboBox CmbDisagree 
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
               Left            =   600
               Sorted          =   -1  'True
               TabIndex        =   106
               Top             =   240
               Width           =   3630
            End
            Begin VB.ComboBox CmbDisagree 
               Height          =   315
               Index           =   0
               Left            =   720
               TabIndex        =   104
               TabStop         =   0   'False
               Top             =   240
               Visible         =   0   'False
               Width           =   1260
            End
            Begin VB.Label Label7 
               Caption         =   "Desc:"
               ForeColor       =   &H00800080&
               Height          =   225
               Index           =   1
               Left            =   120
               TabIndex        =   105
               Top             =   300
               Width           =   435
            End
         End
         Begin VB.Frame Frame13 
            Caption         =   "Frame11"
            Height          =   960
            Left            =   300
            TabIndex        =   83
            Top             =   1635
            Width           =   8490
            Begin VB.ComboBox Combo6 
               Height          =   315
               Left            =   630
               TabIndex        =   84
               Top             =   330
               Width           =   885
            End
            Begin RichTextLib.RichTextBox RichTextBox1 
               Height          =   675
               Index           =   1
               Left            =   2145
               TabIndex        =   85
               Top             =   75
               Width           =   6225
               _ExtentX        =   10980
               _ExtentY        =   1191
               _Version        =   393217
               TextRTF         =   $"FRMCUST_CC_MGM.frx":166A
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
            Begin VB.Label Label4 
               Caption         =   "Note :"
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
               Index           =   6
               Left            =   1560
               TabIndex        =   88
               Top             =   375
               Width           =   600
            End
            Begin VB.Label Label4 
               Caption         =   "Code :"
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
               Index           =   5
               Left            =   60
               TabIndex        =   87
               Top             =   375
               Width           =   600
            End
            Begin VB.Label Label4 
               Caption         =   "Complaint :"
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
               Index           =   1
               Left            =   90
               TabIndex        =   86
               Top             =   0
               Width           =   1035
            End
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Contacted"
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
            Left            =   375
            MouseIcon       =   "FRMCUST_CC_MGM.frx":16E5
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   82
            Top             =   450
            Width           =   1260
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Refferal Recieved"
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
            Height          =   405
            Index           =   2
            Left            =   405
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   81
            Top             =   1125
            Width           =   2115
         End
         Begin VB.Frame Frame11 
            Caption         =   "Frame11"
            Height          =   1260
            Left            =   3780
            TabIndex        =   78
            Top             =   2610
            Width           =   5025
            Begin RichTextLib.RichTextBox RichTextBox1 
               Height          =   975
               Index           =   3
               Left            =   45
               TabIndex        =   79
               Top             =   240
               Width           =   4875
               _ExtentX        =   8599
               _ExtentY        =   1720
               _Version        =   393217
               TextRTF         =   $"FRMCUST_CC_MGM.frx":1B27
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
            Begin VB.Label Label4 
               Caption         =   "Remarks :"
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
               Left            =   90
               TabIndex        =   80
               Top             =   0
               Width           =   825
            End
         End
         Begin VB.Frame Frame12 
            Height          =   615
            Left            =   225
            TabIndex        =   74
            Top             =   480
            Width           =   4305
            Begin VB.ComboBox CmbContacted 
               Height          =   315
               Index           =   1
               Left            =   1635
               Sorted          =   -1  'True
               TabIndex        =   76
               Top             =   225
               Width           =   3450
            End
            Begin VB.ComboBox CmbContacted 
               Height          =   315
               Index           =   0
               Left            =   795
               TabIndex        =   75
               Text            =   "Combo2"
               Top             =   225
               Width           =   1110
            End
            Begin VB.Label Label9 
               Caption         =   "Desc:"
               ForeColor       =   &H00800080&
               Height          =   225
               Index           =   0
               Left            =   285
               TabIndex        =   77
               Top             =   300
               Width           =   435
            End
         End
         Begin VB.Frame Frame22 
            Height          =   1245
            Left            =   225
            TabIndex        =   66
            Top             =   2625
            Width           =   3525
            Begin VB.ComboBox Combo7 
               Height          =   315
               Left            =   1065
               TabIndex        =   69
               Top             =   165
               Width           =   2400
            End
            Begin VB.TextBox Text6 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               ForeColor       =   &H00800080&
               Height          =   225
               Left            =   390
               TabIndex        =   68
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
               Sorted          =   -1  'True
               TabIndex        =   67
               Top             =   825
               Width           =   1350
            End
            Begin TDBTime6Ctl.TDBTime TDBTime1 
               Height          =   315
               Left            =   2550
               TabIndex        =   70
               Top             =   495
               Width           =   900
               _Version        =   65536
               _ExtentX        =   1587
               _ExtentY        =   556
               Caption         =   "FRMCUST_CC_MGM.frx":1BA2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_MGM.frx":1C0E
               Spin            =   "FRMCUST_CC_MGM.frx":1C5E
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
               Value           =   0.984467592592593
            End
            Begin TDBDate6Ctl.TDBDate TDBDate1 
               Height          =   315
               Index           =   3
               Left            =   1065
               TabIndex        =   71
               Top             =   495
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               Calendar        =   "FRMCUST_CC_MGM.frx":1C86
               Caption         =   "FRMCUST_CC_MGM.frx":1D9E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FRMCUST_CC_MGM.frx":1E0A
               Keys            =   "FRMCUST_CC_MGM.frx":1E28
               Spin            =   "FRMCUST_CC_MGM.frx":1E86
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
               TabIndex        =   73
               Top             =   240
               Width           =   960
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
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
               TabIndex        =   72
               Top             =   570
               Width           =   825
            End
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Not Contacted"
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
            Left            =   4740
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   65
            Top             =   450
            Width           =   1590
         End
         Begin VB.Frame Frame25 
            Enabled         =   0   'False
            Height          =   615
            Left            =   4545
            TabIndex        =   61
            Top             =   480
            Width           =   4245
            Begin VB.ComboBox CmbNotContacted 
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
               Left            =   1050
               Sorted          =   -1  'True
               TabIndex        =   63
               Top             =   225
               Width           =   3630
            End
            Begin VB.ComboBox CmbNotContacted 
               Height          =   315
               Index           =   0
               Left            =   570
               TabIndex        =   62
               TabStop         =   0   'False
               Top             =   225
               Visible         =   0   'False
               Width           =   1260
            End
            Begin VB.Label Label7 
               Caption         =   "Desc:"
               ForeColor       =   &H00800080&
               Height          =   225
               Index           =   0
               Left            =   90
               TabIndex        =   64
               Top             =   285
               Width           =   435
            End
         End
         Begin VB.ComboBox CmbDtQuality 
            Height          =   315
            Index           =   1
            Left            =   3015
            Sorted          =   -1  'True
            TabIndex        =   60
            Top             =   120
            Width           =   4005
         End
         Begin VB.ComboBox CmbDtQuality 
            Height          =   315
            Index           =   0
            Left            =   3585
            TabIndex        =   59
            Text            =   "Combo4"
            Top             =   120
            Visible         =   0   'False
            Width           =   1950
         End
         Begin TDBDate6Ctl.TDBDate TDBDate1 
            Height          =   315
            Index           =   1
            Left            =   3795
            TabIndex        =   89
            Top             =   1770
            Visible         =   0   'False
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   556
            Calendar        =   "FRMCUST_CC_MGM.frx":1EAE
            Caption         =   "FRMCUST_CC_MGM.frx":1FC6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FRMCUST_CC_MGM.frx":2032
            Keys            =   "FRMCUST_CC_MGM.frx":2050
            Spin            =   "FRMCUST_CC_MGM.frx":20AE
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
         Begin VB.Label Label9 
            Caption         =   "Data Quality:"
            ForeColor       =   &H00800080&
            Height          =   225
            Index           =   1
            Left            =   1935
            TabIndex        =   90
            Top             =   180
            Width           =   1035
         End
      End
      Begin VB.Frame Frame9 
         Height          =   4020
         Left            =   -74940
         TabIndex        =   56
         Top             =   315
         Width           =   9840
         Begin MSComctlLib.ListView ListView1 
            Height          =   3765
            Index           =   1
            Left            =   15
            TabIndex        =   57
            Top             =   210
            Width           =   9795
            _ExtentX        =   17277
            _ExtentY        =   6641
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
      Begin VB.Frame Frame2 
         Height          =   4020
         Left            =   -74940
         TabIndex        =   14
         Top             =   360
         Width           =   9825
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   1410
            TabIndex        =   93
            Top             =   1935
            Width           =   7440
         End
         Begin VB.Frame Frame7 
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
            Left            =   60
            TabIndex        =   47
            Top             =   2205
            Width           =   5220
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
               Left            =   1140
               MaxLength       =   30
               TabIndex        =   50
               Top             =   240
               Width           =   3990
            End
            Begin VB.ComboBox Combo1 
               BackColor       =   &H00C0FFFF&
               ForeColor       =   &H00C00000&
               Height          =   315
               Index           =   16
               Left            =   1470
               TabIndex        =   49
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
               TabIndex        =   48
               Top             =   2085
               Visible         =   0   'False
               Width           =   3300
            End
            Begin RichTextLib.RichTextBox RichTextBox1 
               Height          =   615
               Index           =   2
               Left            =   1125
               TabIndex        =   51
               Top             =   540
               Width           =   4020
               _ExtentX        =   7091
               _ExtentY        =   1085
               _Version        =   393217
               TextRTF         =   $"FRMCUST_CC_MGM.frx":20D6
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
               Alignment       =   1  'Right Justify
               Caption         =   "Company"
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
               Left            =   75
               TabIndex        =   55
               Top             =   270
               Width           =   990
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "Address"
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
               Left            =   30
               TabIndex        =   54
               Top             =   570
               Width           =   1035
            End
            Begin VB.Label Label3 
               Caption         =   "Gaji / Bulan"
               ForeColor       =   &H00800080&
               Height          =   225
               Index           =   20
               Left            =   105
               TabIndex        =   53
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
               TabIndex        =   52
               Top             =   2130
               Visible         =   0   'False
               Width           =   1095
            End
         End
         Begin VB.Frame Frame18 
            Height          =   1770
            Left            =   5325
            TabIndex        =   33
            Top             =   2190
            Width           =   4455
            Begin VB.CheckBox CekTelp 
               Caption         =   "Invalid"
               Height          =   270
               Index           =   7
               Left            =   2700
               TabIndex        =   101
               Top             =   1305
               Width           =   855
            End
            Begin VB.CheckBox CekTelp 
               Caption         =   "Invalid"
               Height          =   270
               Index           =   6
               Left            =   2700
               TabIndex        =   100
               Top             =   945
               Width           =   855
            End
            Begin VB.CheckBox CekTelp 
               Caption         =   "Invalid"
               Height          =   270
               Index           =   4
               Left            =   3660
               TabIndex        =   99
               Top             =   195
               Width           =   765
            End
            Begin VB.CheckBox CekTelp 
               Caption         =   "Invalid"
               Height          =   270
               Index           =   5
               Left            =   3660
               TabIndex        =   98
               Top             =   510
               Width           =   765
            End
            Begin VB.TextBox TxtExt 
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
               Index           =   0
               Left            =   2940
               MaxLength       =   6
               TabIndex        =   35
               Top             =   165
               Width           =   735
            End
            Begin VB.TextBox TxtExt 
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
               Left            =   2940
               MaxLength       =   6
               TabIndex        =   34
               Top             =   510
               Width           =   735
            End
            Begin TDBMask6Ctl.TDBMask TDBMaskOffice 
               Height          =   360
               Index           =   0
               Left            =   1455
               TabIndex        =   36
               Top             =   165
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC_MGM.frx":2151
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_MGM.frx":21BD
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
               Format          =   "&&&-&&&&&&&&&&&&&&&&&"
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
            Begin TDBMask6Ctl.TDBMask TDBMaskOffice 
               Height          =   360
               Index           =   1
               Left            =   1455
               TabIndex        =   37
               Top             =   540
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC_MGM.frx":21FF
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_MGM.frx":226B
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
               Format          =   "&&&-&&&&&&&&&&&&&&&&&"
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
            Begin TDBMask6Ctl.TDBMask TDBMaskFax 
               Height          =   360
               Index           =   0
               Left            =   1455
               TabIndex        =   38
               Top             =   915
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC_MGM.frx":22AD
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_MGM.frx":2319
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
               Format          =   "&&&-&&&&&&&&&&&&&&&&&"
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
            Begin TDBMask6Ctl.TDBMask TDBMaskFax 
               Height          =   360
               Index           =   1
               Left            =   1455
               TabIndex        =   39
               Top             =   1275
               Width           =   1185
               _Version        =   65536
               _ExtentX        =   2090
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC_MGM.frx":235B
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_MGM.frx":23C7
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
               Format          =   "&&&-&&&&&&&&&&&&&&&&&"
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
            Begin TDBMask6Ctl.TDBMask TDBMaskAOffice 
               Height          =   360
               Index           =   0
               Left            =   855
               TabIndex        =   40
               Top             =   165
               Width           =   585
               _Version        =   65536
               _ExtentX        =   1032
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC_MGM.frx":2409
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_MGM.frx":2475
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
            Begin TDBMask6Ctl.TDBMask TDBMaskAOffice 
               Height          =   360
               Index           =   1
               Left            =   855
               TabIndex        =   41
               Top             =   540
               Width           =   585
               _Version        =   65536
               _ExtentX        =   1032
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC_MGM.frx":24B7
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_MGM.frx":2523
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
            Begin TDBMask6Ctl.TDBMask TDBMaskAFax 
               Height          =   360
               Index           =   0
               Left            =   855
               TabIndex        =   42
               Top             =   915
               Width           =   585
               _Version        =   65536
               _ExtentX        =   1032
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC_MGM.frx":2565
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_MGM.frx":25D1
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
            Begin TDBMask6Ctl.TDBMask TDBMaskAFax 
               Height          =   360
               Index           =   1
               Left            =   855
               TabIndex        =   43
               Top             =   1275
               Width           =   585
               _Version        =   65536
               _ExtentX        =   1032
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC_MGM.frx":2613
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_MGM.frx":267F
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
               Left            =   2670
               TabIndex        =   46
               Top             =   240
               Width           =   240
            End
            Begin VB.Label Label6 
               Caption         =   "Office Phone No."
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
               Height          =   390
               Index           =   10
               Left            =   105
               TabIndex        =   45
               Top             =   240
               Width           =   915
            End
            Begin VB.Label Label6 
               Caption         =   "Fax No."
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
               Left            =   105
               TabIndex        =   44
               Top             =   930
               Width           =   585
            End
         End
         Begin VB.Frame Frame4 
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
            Height          =   1725
            Index           =   1
            Left            =   5325
            TabIndex        =   24
            Top             =   120
            Width           =   4440
            Begin VB.CheckBox CekTelp 
               Caption         =   "Invalid"
               Height          =   270
               Index           =   3
               Left            =   2670
               TabIndex        =   97
               Top             =   1260
               Width           =   855
            End
            Begin VB.CheckBox CekTelp 
               Caption         =   "Invalid"
               Height          =   270
               Index           =   2
               Left            =   2685
               TabIndex        =   96
               Top             =   900
               Width           =   855
            End
            Begin VB.CheckBox CekTelp 
               Caption         =   "Invalid"
               Height          =   270
               Index           =   1
               Left            =   3300
               TabIndex        =   95
               Top             =   540
               Width           =   855
            End
            Begin VB.CheckBox CekTelp 
               Caption         =   "Invalid"
               Height          =   270
               Index           =   0
               Left            =   3300
               TabIndex        =   94
               Top             =   180
               Width           =   855
            End
            Begin TDBMask6Ctl.TDBMask TDBMaskHome 
               Height          =   360
               Index           =   0
               Left            =   1695
               TabIndex        =   25
               Top             =   150
               Width           =   1545
               _Version        =   65536
               _ExtentX        =   2725
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC_MGM.frx":26C1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_MGM.frx":272D
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
               Format          =   "&&&-&&&&&&&&&&&&&&&&&"
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
            Begin TDBMask6Ctl.TDBMask TDBMaskHome 
               Height          =   360
               Index           =   1
               Left            =   1695
               TabIndex        =   26
               Top             =   510
               Width           =   1545
               _Version        =   65536
               _ExtentX        =   2725
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC_MGM.frx":276F
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_MGM.frx":27DB
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
               Format          =   "&&&-&&&&&&&&&&&&&&&&&"
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
            Begin TDBMask6Ctl.TDBMask TDBMaskMobile 
               Height          =   360
               Index           =   0
               Left            =   1080
               TabIndex        =   27
               Top             =   870
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC_MGM.frx":281D
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_MGM.frx":2889
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
               Format          =   "&&&&-&&&&&&&&&&&&&&&&&"
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
            Begin TDBMask6Ctl.TDBMask TDBMaskMobile 
               Height          =   360
               Index           =   1
               Left            =   1080
               TabIndex        =   28
               Top             =   1245
               Width           =   1575
               _Version        =   65536
               _ExtentX        =   2778
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC_MGM.frx":28CB
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_MGM.frx":2937
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
               Format          =   "&&&&-&&&&&&&&&&&&&&&&&"
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
            Begin TDBMask6Ctl.TDBMask TDBMaskAHome 
               Height          =   360
               Index           =   0
               Left            =   1080
               TabIndex        =   29
               Top             =   150
               Width           =   585
               _Version        =   65536
               _ExtentX        =   1032
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC_MGM.frx":2979
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_MGM.frx":29E5
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
            Begin TDBMask6Ctl.TDBMask TDBMaskAHome 
               Height          =   360
               Index           =   1
               Left            =   1080
               TabIndex        =   30
               Top             =   510
               Width           =   585
               _Version        =   65536
               _ExtentX        =   1032
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC_MGM.frx":2A27
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC_MGM.frx":2A93
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
               Caption         =   "Home Phone No :"
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
               Height          =   480
               Index           =   2
               Left            =   75
               TabIndex        =   32
               Top             =   225
               Width           =   885
            End
            Begin VB.Label Label6 
               Caption         =   "Mobile No :"
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
               Left            =   90
               TabIndex        =   31
               Top             =   900
               Width           =   795
            End
         End
         Begin VB.Frame Frame3 
            Height          =   1725
            Index           =   0
            Left            =   60
            TabIndex        =   15
            Top             =   120
            Width           =   5220
            Begin VB.TextBox TxtCity 
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
               Left            =   2610
               MaxLength       =   20
               TabIndex        =   92
               Top             =   675
               Width           =   2520
            End
            Begin VB.TextBox TxtZip 
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
               Left            =   1140
               MaxLength       =   5
               TabIndex        =   17
               Top             =   675
               Width           =   945
            End
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
               Left            =   2520
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   16
               TabStop         =   0   'False
               Text            =   "0"
               Top             =   1005
               Width           =   375
            End
            Begin TDBDate6Ctl.TDBDate TDBDate1 
               Height          =   315
               Index           =   0
               Left            =   1140
               TabIndex        =   18
               Top             =   1005
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   556
               Calendar        =   "FRMCUST_CC_MGM.frx":2AD5
               Caption         =   "FRMCUST_CC_MGM.frx":2BED
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FRMCUST_CC_MGM.frx":2C59
               Keys            =   "FRMCUST_CC_MGM.frx":2C77
               Spin            =   "FRMCUST_CC_MGM.frx":2CD5
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
            Begin RichTextLib.RichTextBox TxtAlamat 
               Height          =   555
               Left            =   1125
               TabIndex        =   91
               Top             =   135
               Width           =   4020
               _ExtentX        =   7091
               _ExtentY        =   979
               _Version        =   393217
               TextRTF         =   $"FRMCUST_CC_MGM.frx":2CFD
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
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "ZIP"
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
               TabIndex        =   23
               Top             =   720
               Width           =   915
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "Address"
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
               Left            =   195
               TabIndex        =   22
               Top             =   240
               Width           =   900
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               Caption         =   "City"
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
               Left            =   2145
               TabIndex        =   21
               Top             =   720
               Width           =   405
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "Date Of Birth"
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
               Left            =   120
               TabIndex        =   20
               Top             =   1065
               Width           =   975
            End
            Begin VB.Label Label3 
               Caption         =   "Years Old"
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
               Left            =   2940
               TabIndex        =   19
               Top             =   1050
               Width           =   780
            End
         End
      End
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   420
      Index           =   4
      Left            =   135
      TabIndex        =   109
      Top             =   15
      Visible         =   0   'False
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   741
      _Version        =   196610
      Font3D          =   2
      MousePointer    =   16
      ForeColor       =   128
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
      Picture         =   "FRMCUST_CC_MGM.frx":2D78
      Caption         =   "&Close"
      AutoSize        =   2
      Alignment       =   4
      ButtonStyle     =   2
      PictureAlignment=   1
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   420
      Index           =   5
      Left            =   60
      TabIndex        =   110
      Top             =   15
      Visible         =   0   'False
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   741
      _Version        =   196610
      Font3D          =   2
      MousePointer    =   16
      ForeColor       =   128
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
      Picture         =   "FRMCUST_CC_MGM.frx":309A
      Caption         =   "&Close"
      AutoSize        =   2
      Alignment       =   4
      ButtonStyle     =   2
      PictureAlignment=   1
   End
   Begin VB.TextBox TxtReasonClosing 
      Height          =   375
      Left            =   2700
      TabIndex        =   111
      Top             =   30
      Visible         =   0   'False
      Width           =   4290
   End
   Begin VB.Label Label2 
      Caption         =   "Data Code :"
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
      TabIndex        =   9
      Top             =   990
      Width           =   1260
   End
End
Attribute VB_Name = "FRMCUST_CC_MGM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_customer_mgm As ADODB.Recordset
Dim pStatusLstCall As String
Dim M_followup As Boolean
Public closeOk  As Boolean

'tambah coding 14 juli 2006
'Private Sub HEADER_VIEW_Refferall()
'    ListView1(0).ColumnHeaders.ADD 1, , "No", 3 * TXT
'    ListView1(0).ColumnHeaders.ADD 2, , "Cust Id", 5 * TXT
'    ListView1(0).ColumnHeaders.ADD 3, , "Nama Customer", 10 * TXT
'    ListView1(0).ColumnHeaders.ADD 4, , "Tgl Schedule", 5 * TXT
'    ListView1(0).ColumnHeaders.ADD 5, , "Next Action", 12 * TXT
'    ListView1(0).ColumnHeaders.ADD 6, , "Remarks", 17 * TXT
'    ListView1(0).ColumnHeaders.ADD 7, , "SalesCode", 8 * TXT
'    ListView1(0).ColumnHeaders.ADD 8, , "LastCall Date", 10 * TXT
'    ListView1(0).ColumnHeaders.ADD 9, , "Sts LastCall", 10 * TXT
'End Sub
'
'Private Sub show_Reff()
'Dim m_reff As New ADODB.Recordset
'Dim listitem As listitem
'm_reff.CursorLocation = adUseClient
'm_reff.Open "Select * from cc_custtbl where CustIdRef ='" + Text1(1).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'While Not m_reff.EOF
'    Set listitem = ListView1(0).ListItems.ADD(, , CStr(m_reff.Bookmark))
'        listitem.SubItems(1) = IIf(IsNull(m_reff("CUSTID")), "", m_reff("CUSTID"))
'        listitem.SubItems(2) = IIf(IsNull(m_reff("NAME")), "", m_reff("NAME"))
'        listitem.SubItems(3) = Format(IIf(IsNull(m_reff("NEXTACTDATE")), "", m_reff("NEXTACTDATE")), "yyyy/mm/dd hh:mm:ss")
'        listitem.SubItems(4) = IIf(IsNull(m_reff("NEXTACT")), "", m_reff("NEXTACT"))
'        listitem.SubItems(5) = IIf(IsNull(m_reff("REMARKS")), "", m_reff("REMARKS"))
'        listitem.SubItems(6) = IIf(IsNull(m_reff("AGENT")), "", m_reff("AGENT"))
'        listitem.SubItems(7) = IIf(IsNull(m_reff("TGLSTATUS")), "", m_reff("TGLSTATUS"))
'        listitem.SubItems(8) = IIf(IsNull(m_reff("KETHSLKERJA")), "", m_reff("KETHSLKERJA"))
'    m_reff.MoveNext
'Wend
'Set m_reff = Nothing
'End Sub
'
''end coding 14 juli 2006
'
'Private Sub Check2_Click(Index As Integer)
'Select Case Index
'    Case 0
'        If Check2(Index).Value Then
'            Frame25.Enabled = False
'            Frame12.Enabled = True
'            Check2(1).Value = 0
'            Check2(2).Value = 0
'            Check2(3).Value = 0
'            Frame5.Enabled = False
'            TDBDate1(1).Text = Empty
'            TDBDate1(1).Enabled = False
'        Else
'            TDBDate1(1).Enabled = True
'            cmbContacted(0).Text = ""
'            cmbContacted(1).Text = ""
'        End If
'    Case 1
'        If Check2(Index).Value Then
'            Check2(0).Value = 0
'            Check2(2).Value = 0
'            Check2(3).Value = 0
'            Frame5.Enabled = False
'            Frame25.Enabled = True
'            Frame12.Enabled = False
'            TDBDate1(1).Text = Empty
'            TDBDate1(1).Enabled = False
'        Else
'            Frame25.Enabled = False
'            TDBDate1(1).Enabled = True
'            CmbNotContacted(0).Text = ""
'            CmbNotContacted(1).Text = ""
'        End If
'    Case 2
'        If Check2(Index).Value Then
'            Check2(0).Value = 0
'            Check2(1).Value = 0
'            Check2(3).Value = 0
'            Frame5.Enabled = False
'            Frame25.Enabled = True
'            Frame12.Enabled = False
'        Else
'            Frame25.Enabled = False
'            TDBDate1(1).Value = Empty
'        End If
'    Case 3
'        If Check2(Index).Value Then
'            Frame25.Enabled = False
'            Frame12.Enabled = False
'            Frame5.Enabled = True
'            Check2(1).Value = 0
'            Check2(2).Value = 0
'            Check2(0).Value = 0
'            TDBDate1(1).Text = Empty
'            TDBDate1(1).Enabled = False
'        Else
'            TDBDate1(1).Enabled = True
'            CmbDisagree(0).Text = ""
'            CmbDisagree(1).Text = ""
'            Frame5.Enabled = False
'        End If
'End Select
'End Sub
'
'Private Sub cmbContacted_Click(Index As Integer)
'    CmbContacted_LostFocus (Index)
'End Sub
'
'Private Sub CmbContacted_LostFocus(Index As Integer)
'Dim m_objrs As ADODB.Recordset
'Select Case Index
'Case 0
'    Set m_objrs = New ADODB.Recordset
'    m_objrs.CursorLocation = adUseClient
'    m_objrs.Open "Select * from ContactedDesc WHERE KdNoProdPresented ='" + cmbContacted(Index).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'    If m_objrs.RecordCount <> 0 Then
'        cmbContacted(0).Text = IIf(IsNull(m_objrs!KdNoProdPresented), "", m_objrs!KdNoProdPresented)
'        cmbContacted(1).Text = IIf(IsNull(m_objrs!NmNoProdPresented), "", m_objrs!NmNoProdPresented)
'        Select Case cmbContacted(0).Text
'            Case "SA"
'                Combo7.Text = "Call Again"
'            Case "D"
'                Combo7.Text = "Do Not Call"
'            Case "ST"
'                Combo7.Text = "Call Again"
'            Case "PU"
'                Combo7.Text = "Call Again"
'            Case "ID"
'                Combo7.Text = "Call Again"
'        End Select
'    Else
'        cmbContacted(0).Text = Empty
'        cmbContacted(1).Text = Empty
'    End If
'    Set m_objrs = Nothing
'Case 1
'    Set m_objrs = New ADODB.Recordset
'    m_objrs.CursorLocation = adUseClient
'    m_objrs.Open "Select * from ContactedDesc WHERE NmNoProdPresented ='" + cmbContacted(Index).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'    If m_objrs.RecordCount <> 0 Then
'        cmbContacted(0).Text = IIf(IsNull(m_objrs!KdNoProdPresented), "", m_objrs!KdNoProdPresented)
'        cmbContacted(1).Text = IIf(IsNull(m_objrs!NmNoProdPresented), "", m_objrs!NmNoProdPresented)
'        Select Case cmbContacted(0).Text
'            Case "SA"
'                Combo7.Text = "Call Again"
'            Case "D"
'                Combo7.Text = "Do Not Call"
'            Case "ST"
'                Combo7.Text = "Call Again"
'            Case "PU"
'                Combo7.Text = "Call Again"
'            Case "ID"
'                Combo7.Text = "Call Again"
'        End Select
'    Else
'        cmbContacted(0).Text = Empty
'        cmbContacted(1).Text = Empty
'    End If
'    Set m_objrs = Nothing
'End Select
'End Sub
'Private Sub CmbDisagree_Click(Index As Integer)
'    Call CmbDisagree_LostFocus(Index)
'End Sub
'
'Private Sub CmbDisagree_LostFocus(Index As Integer)
'Dim m_objrs As ADODB.Recordset
'Select Case Index
'Case 0
'    Set m_objrs = New ADODB.Recordset
'    m_objrs.CursorLocation = adUseClient
'    m_objrs.Open "Select * from TblDisAgreeMGM WHERE KdDisagree ='" + CmbDisagree(Index).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'    If m_objrs.RecordCount <> 0 Then
'        CmbDisagree(0).Text = IIf(IsNull(m_objrs!kddisagree), "", m_objrs!kddisagree)
'        CmbDisagree(1).Text = IIf(IsNull(m_objrs!ketdisagree), "", m_objrs!ketdisagree)
'        'Combo7.Text = "Do Not Call"
'    Else
'        CmbDisagree(0).Text = Empty
'        CmbDisagree(1).Text = Empty
'    End If
'    Set m_objrs = Nothing
'Case 1
'    Set m_objrs = New ADODB.Recordset
'    m_objrs.CursorLocation = adUseClient
'    m_objrs.Open "Select * from TblDisAgreeMGM WHERE KetDisagree='" + CmbDisagree(Index).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'    If m_objrs.RecordCount <> 0 Then
'        CmbDisagree(0).Text = IIf(IsNull(m_objrs!kddisagree), "", m_objrs!kddisagree)
'        CmbDisagree(1).Text = IIf(IsNull(m_objrs!ketdisagree), "", m_objrs!ketdisagree)
'        'Combo7.Text = "Do Not Call"
'    Else
'        CmbDisagree(0).Text = Empty
'        CmbDisagree(1).Text = Empty
'    End If
'    Set m_objrs = Nothing
'End Select
'End Sub
'
'Private Sub CmbDtQuality_Click(Index As Integer)
'    CmbDtQuality_LostFocus (Index)
'End Sub
'
'Private Sub CmbDtQuality_LostFocus(Index As Integer)
'Dim m_objrs As ADODB.Recordset
'Select Case Index
'Case 0
'    Set m_objrs = New ADODB.Recordset
'    m_objrs.CursorLocation = adUseClient
'    m_objrs.Open "Select * from DataQuality WHERE KdDataQuality ='" + CmbDtQuality(Index).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'    If m_objrs.RecordCount <> 0 Then
'        CmbDtQuality(0).Text = IIf(IsNull(m_objrs!KdDataQuality), "", m_objrs!KdDataQuality)
'        CmbDtQuality(1).Text = IIf(IsNull(m_objrs!NmDataQuality), "", m_objrs!NmDataQuality)
'    Else
'        CmbDtQuality(0).Text = Empty
'        CmbDtQuality(1).Text = Empty
'    End If
'    Set m_objrs = Nothing
'Case 1
'    Set m_objrs = New ADODB.Recordset
'    m_objrs.CursorLocation = adUseClient
'    m_objrs.Open "Select * from DataQuality WHERE NmDataQuality='" + CmbDtQuality(Index).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'    If m_objrs.RecordCount <> 0 Then
'        CmbDtQuality(0).Text = IIf(IsNull(m_objrs!KdDataQuality), "", m_objrs!KdDataQuality)
'        CmbDtQuality(1).Text = IIf(IsNull(m_objrs!NmDataQuality), "", m_objrs!NmDataQuality)
'    Else
'        CmbDtQuality(0).Text = Empty
'        CmbDtQuality(1).Text = Empty
'    End If
'    Set m_objrs = Nothing
'End Select
'End Sub
'
'Private Sub CmbNotContacted_Click(Index As Integer)
'    CmbNotContacted_LostFocus (Index)
'End Sub
'
'Private Sub CmbNotContacted_LostFocus(Index As Integer)
'Dim m_data As New CLS_FRMCUST_CC_MGM
'Dim m_objrs As ADODB.Recordset
'Select Case Index
'    Case 0
'        Set m_objrs = m_data.QUERY_COMBO_CLOSSING(M_OBJCONN, "KDCLS = '" + CmbNotContacted(Index).Text + "'")
'        If m_objrs.RecordCount <> 0 Then
'            CmbNotContacted(0).Text = m_objrs("KDCLS")
'            CmbNotContacted(1).Text = IIf(IsNull(m_objrs("KETCLS")), "", m_objrs("KETCLS"))
'            Select Case CmbNotContacted(0).Text
'                Case ""
'                    Combo7.Text = ""
'                Case "WN"
'                    Combo7.Text = "Do Not Call"
'                Case Else
'                    Combo7.Text = "Call Again"
'            End Select
'        Else
'            CmbNotContacted(0).Text = Empty
'            CmbNotContacted(1).Text = Empty
'        End If
'    Case 1
'        Set m_objrs = m_data.QUERY_COMBO_CLOSSING(M_OBJCONN, "KETCLS = '" + CmbNotContacted(Index).Text + "'")
'        If m_objrs.RecordCount <> 0 Then
'            CmbNotContacted(0).Text = m_objrs("KDCLS")
'            CmbNotContacted(1).Text = IIf(IsNull(m_objrs("KETCLS")), "", m_objrs("KETCLS"))
'            Select Case CmbNotContacted(0).Text
'                Case ""
'                    Combo7.Text = ""
'                Case "WN"
'                    Combo7.Text = "Do Not Call"
'                Case "UD"
'                    Combo7.Text = "Do Not Call"
'                Case Else
'                    Combo7.Text = "Call Again"
'            End Select
'        Else
'            CmbNotContacted(0).Text = Empty
'            CmbNotContacted(1).Text = Empty
'        End If
'End Select
'Set m_objrs = Nothing
'Set m_data = Nothing
'End Sub
'
'
'Private Sub Combo1_Click(Index As Integer)
'Dim m_data As New CLS_FRMCUST_CC_MGM
'Dim m_objrs As ADODB.Recordset
'Select Case Index
'    Case 0
'        Set m_objrs = m_data.QUERY_COMBO_DATASOURCE(M_OBJCONN, "KODEDS = '" + Combo1(Index).Text + "'")
'        If m_objrs.RecordCount <> 0 Then
'            Combo1(0).Text = m_objrs("KODEDS")
'            Combo1(1).Text = IIf(IsNull(m_objrs("KETERANGAN")), "", m_objrs("KETERANGAN"))
'            Text1(3).Text = IIf(IsNull(m_objrs("KETERANGAN")), "", m_objrs("KETERANGAN"))
'            ket = Text1(3).Text
'        Else
'            Combo1(0).Text = Empty
'            Combo1(1).Text = Empty
'            Text1(3).Text = Empty
'        End If
'    Case 1
'        Set m_objrs = m_data.QUERY_COMBO_DATASOURCE(M_OBJCONN, "KETERANGAN = '" + Combo1(Index).Text + "'")
'        If m_objrs.RecordCount <> 0 Then
'            Combo1(0).Text = m_objrs("KODEDS")
'            Combo1(1).Text = IIf(IsNull(m_objrs("KETERANGAN")), "", m_objrs("KETERANGAN"))
'            Text1(3).Text = IIf(IsNull(m_objrs("KETERANGAN")), "", m_objrs("KETERANGAN"))
'        Else
'            Combo1(0).Text = Empty
'            Combo1(1).Text = Empty
'            Text1(3).Text = Empty
'        End If
'End Select
'Set m_objrs = Nothing
'Set m_data = Nothing
'End Sub
'
'Private Sub Combo1_LostFocus(Index As Integer)
'Dim m_data As New CLS_FRMCUST_CC_MGM
'Dim m_objrs As ADODB.Recordset
'Select Case Index
'    Case 0
'        Set m_objrs = m_data.QUERY_COMBO_DATASOURCE(M_OBJCONN, "KODEDS = '" + Combo1(Index).Text + "'")
'        If m_objrs.RecordCount <> 0 Then
'            Combo1(0).Text = m_objrs("KODEDS")
'            Combo1(1).Text = IIf(IsNull(m_objrs("KETERANGAN")), "", m_objrs("KETERANGAN"))
'        Else
'            Combo1(0).Text = Empty
'            Combo1(1).Text = Empty
'        End If
'    Case 1
'        Set m_objrs = m_data.QUERY_COMBO_DATASOURCE(M_OBJCONN, "KETERANGAN = '" + Combo1(Index).Text + "'")
'        If m_objrs.RecordCount <> 0 Then
'            Combo1(0).Text = m_objrs("KODEDS")
'            Combo1(1).Text = IIf(IsNull(m_objrs("KETERANGAN")), "", m_objrs("KETERANGAN"))
'        Else
'            Combo1(0).Text = Empty
'            Combo1(1).Text = Empty
'        End If
'End Select
'Set m_objrs = Nothing
'Set m_data = Nothing
'End Sub
'
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    Set m_customer_mgm = Nothing
'    Status_Form = 0
'End Sub
'
'Private Sub ListView1_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'
'Select Case Index
'Case 1, 0
'    ListView1(Index).SortKey = ColumnHeader.Index - 1
'    ListView1(Index).Sorted = True
'End Select
'End Sub
'
'Private Function CEK_DATA_VALID() As Boolean
'Dim m_msgbox As Variant
'    If TDBDate1(3).ValueIsNull = True Or TDBTime1.ValueIsNull = True Then
'        CEK_DATA_VALID = False
'        MsgBox "Tanggal Schedule Harus Di isi", vbCritical + vbOKOnly, "Peringatan"
'        SSTab1.Tab = 2
'        Exit Function
'    End If
'
'    If Check2(2).Value = 0 And Check2(1).Value = 0 And Check2(0).Value = 0 And Check2(3).Value = 0 Then
'        CEK_DATA_VALID = False
'        MsgBox "Status Harus Di isi", vbCritical + vbOKOnly, "Peringatan"
'        SSTab1.Tab = 2
'        Exit Function
'    End If
'
'    If Text1(0).Text = Empty Then
'        CEK_DATA_VALID = False
'        MsgBox "Nama Harus Diisi", vbCritical + vbOKOnly, "Peringatan"
'        Text1(0).SetFocus
'        Exit Function
'    End If
'    If Check2(1).Value = 1 Then
'        If CmbNotContacted(1).Text = Empty Then
'            CEK_DATA_VALID = False
'            MsgBox "Not Contacted Desc Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
'            SSTab1.Tab = 2
'            Exit Function
'        End If
'    End If
'    If Check2(3).Value = 1 Then
'        If CmbDisagree(1).Text = Empty Then
'            CEK_DATA_VALID = False
'            MsgBox "Disagree Desc Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
'            SSTab1.Tab = 2
'            Exit Function
'        End If
'    End If
'    If Check2(0).Value = 1 Then
'        If cmbContacted(1).Text = Empty Then
'            CEK_DATA_VALID = False
'            MsgBox "Contacted Desc Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
'            SSTab1.Tab = 2
'            Exit Function
'        End If
'    End If
'    If Check2(1).Value = 1 Then
'        RichTextBox1(3).Text = RichTextBox1(3).Text & " " & CmbNotContacted(1).Text
'    Else
'        If RichTextBox1(3).Text = Empty And Combo7.Text = Empty Then
'            CEK_DATA_VALID = False
'            MsgBox "Catatan(Perubahan Pada Data Customer Ini) Atau Next Action Harus Diisi...!!!", vbCritical + vbOKOnly, "Peringatan"
'            SSTab1.Tab = 2
'            Combo7.SetFocus
'            Exit Function
'        End If
'    End If
'CEK_DATA_VALID = True
'End Function
'
'Private Sub ListView1_DblClick(Index As Integer)
'    reff_View = True
'    FRMCUST_CC.Show vbModal
'    reff_View = False
'End Sub
'
'Private Sub RichTextBox1_LostFocus(Index As Integer)
'Select Case Index
'Case 3
'    RichTextBox1(3).Text = UCase(RichTextBox1(3).Text)
'End Select
'End Sub
'
'Private Sub SSCommand1_Click(Index As Integer)
'Dim m_msgbox As Variant
'Dim V_SAVE As Boolean
'Dim CMDSQL2 As String
'Dim m_popup As New ADODB.Recordset
'
'V_SAVE = True
'Select Case Index
'    Case 0
'        If ADD_CUST = True Then
'        Else
'            With frmtelp_mgm
'                .TELPRUMAH1 = IIf(IsNull(m_customer_mgm("HOMENO")), "", m_customer_mgm("HOMENO"))
'                .TELPRUMAH2 = IIf(IsNull(m_customer_mgm("HOMENO2")), "", m_customer_mgm("HOMENO2"))
'                .HANDPHONE1 = IIf(IsNull(m_customer_mgm("MOBILENO")), "", m_customer_mgm("MOBILENO"))
'                .HANDPHONE2 = IIf(IsNull(m_customer_mgm("MOBILENO2")), "", m_customer_mgm("MOBILENO2"))
'                .TELPKANTOR1 = IIf(IsNull(m_customer_mgm("OFFICENO")), "", m_customer_mgm("OFFICENO"))
'                .TELPKANTOR2 = IIf(IsNull(m_customer_mgm("OFFICENO2")), "", m_customer_mgm("OFFICENO2"))
'                If Left(.TELPRUMAH1, 3) = "021" Then
'                    .TELPRUMAH1 = ILANGIN_AREA(.TELPRUMAH1)
'                End If
'                If Left(.TELPKANTOR1, 3) = "021" Then
'                    .TELPKANTOR1 = ILANGIN_AREA(.TELPKANTOR1)
'                End If
'                .Show vbModal
'                M_followup = True
'            Set m_popup = New ADODB.Recordset
'            m_popup.CursorLocation = adUseClient
'            m_popup.Open "Select * from vwcallcfg1", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'            CMDSQL2 = "UPDATE usertbl set flagcall ='" + Format(m_popup!tglsystem, "hh:mm:dd") + "' where userid ='" + MDIForm1.Text1.Text + "'"
'            M_OBJCONN.Execute CMDSQL2
'            Set m_popup = Nothing
'            End With
'
'        End If
'
'    Case 1
'        ADD_CUST = True
'        AddMgm = True
'        Check2(2).Value = 1
''        FRMCUST_CC.TxtMgmName.Text = Text1(0).Text
''        FRMCUST_CC.Show vbModal
'
''tambah coding 14 juli 2006
'        With FrmEntryReffSearch
'            .TxtRecsourceRef.Text = Combo1(0).Text
'            .TxtIdReff.Text = Text1(1).Text
'            .TxtNamaReff.Text = Text1(0).Text
'            .TxtIdReff.Enabled = False
'             .Show vbModal
'             If FrmEntryReff.okReff Then
'                Dim listitem As listitem
'                VIEW_MGMDATA.SSTab1.Tab = 1
'                    Set listitem = VIEW_MGMDATA.ListView1.ListItems.ADD(, , "99999")
'                        listitem.SubItems(1) = FrmEntryReff.IdCusti
'                        listitem.SubItems(2) = ""
'                        listitem.SubItems(3) = FrmEntryReff.TxtIdReff.Text
'                        listitem.SubItems(4) = FrmEntryReff.TxtNamaReff.Text
'                        listitem.SubItems(5) = FrmEntryReff.TxtNama.Text
'                        listitem.SubItems(6) = ""
'                        listitem.SubItems(7) = ""
'                        listitem.SubItems(8) = ""
'                        listitem.SubItems(9) = MDIForm1.Text1.Text
'                        listitem.SubItems(10) = MDIForm1.Text7.Text
'                        listitem.SubItems(11) = FrmEntryReff.cmbRecsource.Text
'                        listitem.SubItems(12) = ""
'                        listitem.SubItems(13) = ""
'                        listitem.SubItems(14) = ""
'                        listitem.SubItems(15) = ""
'                        listitem.SubItems(16) = ""
'                        listitem.SubItems(17) = ""
'             End If
'             Unload FrmEntryReff
'             Unload FrmEntryReffSearch
'        End With
'        ADD_CUST = False
'        AddMgm = False
'    Case 2
'        V_SAVE = CEK_DATA_VALID
'        If V_SAVE = False Then
'            Exit Sub
'        Else
'        End If
'        If ADD_CUST Then
'        Else
'            Call CEK_UPDATE_PELANGGAN
'        End If
'    Case 3
'        If M_followup = True Then
'            If UCase(Trim(MDIForm1.Text2.Text)) = "AGENT" Then
'            If ADD_CUST Then
'                Unload Me
'            Else
'                MsgBox "Simpan data terlebih dahulu", vbInformation + vbOKOnly, "Telegrandi"
'                Exit Sub
''                V_SAVE = cekSaveOnExit
''                If V_SAVE = False Then
''                Else
''                    Call Save_OnExits
''                    Unload Me
''                End If
'            End If
'            Else
'                Unload Me
'            End If
'        Else
'            Unload Me
'        End If
'    Case 4
'
'           'close data org
''        If (Now - IIf(IsNull(m_customer_mgm!TGLSTATUS), Now, m_customer_mgm!TGLSTATUS)) < CCur(MDIForm1.TxtLamaFollowup.Text) Then
''            MsgBox "Data Masih Bisa DiFollow Up Oleh Agent Yang Lama", vbInformation + vbOKOnly, adCmdText
''            closeOk = False
''            Exit Sub
''        Else
'            If Len(TxtReasonClosing.Text) < 5 Then
'                MsgBox "Reason Harus Di isi", vbInformation + vbOKOnly, "Telegrandi"
'                TxtReasonClosing.SetFocus
'                Exit Sub
'            End If
'            m_msgbox = MsgBox("Close data ini ??..", vbYesNo, "Telegrandi")
'            If m_msgbox = vbNo Then
'                closeOk = False
'                Exit Sub
'            End If
'
'             ' kirim ke pending duplikasi dulu
'             Dim CHBARUid As String
'            Set m_popup = New ADODB.Recordset
'            m_popup.CursorLocation = adUseClient
'            CMDSQL2 = "Select * from TBL_DUPLIKASICH WHERE custid ='" + Text1(1).Text + "' AND STS=0"
'            m_popup.Open CMDSQL2, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'            If m_popup.RecordCount <> 0 Then
'                CHBARUid = IIf(IsNull(m_popup!CUSTIDBARU), "", m_popup!CUSTIDBARU)
'                m_popup!sts = 1
'                m_popup.UPDATE
'                m_popup.Requery
'                CMDSQL2 = " insert into TblPendingDuplikasiCh (CustIdLama ,ReasonClosing,CustIdBaru)"
'                CMDSQL2 = CMDSQL2 + " values ('" + Text1(1).Text + "', "
'                CMDSQL2 = CMDSQL2 + " '" + TxtReasonClosing.Text + "', "
'                CMDSQL2 = CMDSQL2 + " '" + CHBARUid + "')"
'                M_OBJCONN.Execute CMDSQL2
'            End If
'            closeOk = True
'            Me.Hide
''        End If
'    Case 5
'            m_msgbox = MsgBox("Close data ini ??..", vbYesNo, "Telegrandi")
'            If m_msgbox = vbNo Then
'                closeOk = False
'                Exit Sub
'            End If
'            'pindahin data ke tbl close
'            CMDSQL2 = "Insert into TBL_CLOSEDATAMGM select * from MGM where custid ='" + Text1(1).Text + "'"
'            M_OBJCONN.Execute CMDSQL2
'            WaitSecs (0.25)
'            'CMDSQL2 = "Delete from cc_custtbl where custid ='" + Text1(1).Text + "'"
'            'pakai update aja..
'            Dim AGENTBARU As String
'            Set m_popup = New ADODB.Recordset
'            m_popup.CursorLocation = adUseClient
'            m_popup.Open "Select * from TblPendingDuplikasiCh WHERE custidlama ='" + Text1(1).Text + "' and sts =0", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'            If m_popup.RecordCount <> 0 Then
'                AGENTBARU = FrmDuplikasiPendingCh.LstExLeads.SelectedItem.SubItems(13)
'                'IIf(IsNull(m_popup!AGENTBARU), "", m_popup!AGENTBARU)
'                    m_popup!sts = 1
'                    m_popup.UPDATE
'                    m_popup.Requery
'                CMDSQL2 = "Update MGM set agent ='" + AGENTBARU + "', NAMAAGENT ='' where custid ='" + Text1(1).Text + "'"
'                M_OBJCONN.Execute CMDSQL2
'            End If
'            Set m_popup = Nothing
'            closeOk = True
'            Me.Hide
'End Select
'End Sub
'
'Private Function cekSaveOnExit() As Boolean
'    If Check2(2).Value = 0 And Check2(1).Value = 0 And Check2(0).Value = 0 And Check2(3).Value = 0 Then
'        MsgBox "Status Harus Di isi", vbCritical + vbOKOnly, "Peringatan"
'        SSTab1.Tab = 2
'        cekSaveOnExit = False
'        Exit Function
'    End If
'    If Check2(1).Value = 1 Then
'        If CmbNotContacted(1).Text = Empty Then
'            MsgBox "Not Contacted Desc Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
'            SSTab1.Tab = 2
'            cekSaveOnExit = False
'            Exit Function
'        End If
'    End If
'    If Check2(3).Value = 1 Then
'        If CmbDisagree(1).Text = Empty Then
'            MsgBox "Disagree Desc Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
'            SSTab1.Tab = 2
'            cekSaveOnExit = False
'            Exit Function
'        End If
'    End If
'
'    If Check2(0).Value = 1 Then
'        If cmbContacted(1).Text = Empty Then
'            MsgBox "Contacted Desc Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
'            SSTab1.Tab = 2
'            cekSaveOnExit = False
'            Exit Function
'        End If
'    End If
'cekSaveOnExit = True
'End Function
'
'
'
'Private Sub Save_OnExits()
'On Error GoTo saveexitErr
'    m_customer_mgm("TGLCALL") = Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") & " " & Format(Now, "hh:nn")
'    If Check2(0).Value Then
'        m_customer_mgm("RECSTATUS") = "C"
'        pStatusLstCall = cmbContacted(0).Text
'    End If
'    If Check2(1).Value Then
'        m_customer_mgm("RECSTATUS") = "N"
'        pStatusLstCall = CmbNotContacted(0).Text
'    End If
'    If Check2(3).Value Then
'        m_customer_mgm("RECSTATUS") = "DA"
'        pStatusLstCall = CmbDisagree(0).Text
'    End If
'    If Check2(2).Value Then
'        m_customer_mgm("RECSTATUS") = "RR"
'        pStatusLstCall = "RR"
'        m_customer_mgm("INCOMING") = "RR"
'        m_customer_mgm("TGLINCOMING") = Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") & " " & Format(Now, "hh:nn")
'    End If
'    If Trim(UCase(IIf(IsNull(m_customer_mgm("KETHSLKERJA")), "", m_customer_mgm("KETHSLKERJA")))) = Trim(UCase(pStatusLstCall)) Then
'    Else
'        m_customer_mgm("KETHSLKERJA") = pStatusLstCall
'        m_customer_mgm("TGLSTATUS") = Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") & " " & Format(Now, "hh:nn")
'    End If
'm_customer_mgm.UPDATE
'Exit Sub
'saveexitErr:
'    MsgBox Err.Description
'End Sub
'
'Private Sub CEK_UPDATE_PELANGGAN()
'Dim m_data As New CLS_FRMCUST_CC_MGM
'Dim pStatusHstLstCall As String
'On Error GoTo editErr
'    M_OBJCONN.BeginTrans
'
'        m_customer_mgm("CUSTID") = Text1(1).Text
'        m_customer_mgm("NAME") = Text1(0).Text
'        m_customer_mgm("BIRTHD") = TDBDate1(0).Value
'    '    m_customer_mgm("RECSOURCE") = Combo1(0).Text
'        m_customer_mgm("ADDRNOW") = TxtAlamat.Text
'        m_customer_mgm("ZIPNOW") = TxtZip.Text
'        m_customer_mgm("CITYNOW") = TxtCity.Text
'
'        m_customer_mgm("HOMENO") = TDBMaskHome(0).Value
'        If Len(IIf(IsNull(m_customer_mgm!HOMENO), "", m_customer_mgm!HOMENO)) > 2 Then
'            TDBMaskHome(0).ReadOnly = True
'        End If
'
'        m_customer_mgm("HOMENO2") = TDBMaskHome(1).Value
'        If Len(IIf(IsNull(m_customer_mgm!HOMENO2), "", m_customer_mgm!HOMENO2)) > 2 Then
'            TDBMaskHome(1).ReadOnly = True
'        End If
'
'        m_customer_mgm("MOBILENO") = TDBMaskMobile(0).Value
'        If Len(IIf(IsNull(m_customer_mgm!MOBILENO), "", m_customer_mgm!MOBILENO)) > 2 Then
'            TDBMaskMobile(0).ReadOnly = True
'        End If
'
'        m_customer_mgm("MOBILENO2") = TDBMaskMobile(1).Value
'        If Len(IIf(IsNull(m_customer_mgm!MOBILENO2), "", m_customer_mgm!MOBILENO2)) > 2 Then
'            TDBMaskMobile(1).ReadOnly = True
'        End If
'
'        m_customer_mgm("OFFICENO") = TDBMaskOffice(0).Value
'        If Len(IIf(IsNull(m_customer_mgm!OFFICENO), "", m_customer_mgm!OFFICENO)) > 2 Then
'            TDBMaskOffice(0).ReadOnly = True
'        End If
'        m_customer_mgm("EXTOFFICE") = TxtExt(0).Text
'
'        m_customer_mgm("OFFICENO2") = TDBMaskOffice(1).Value
'        If Len(IIf(IsNull(m_customer_mgm!OFFICENO2), "", m_customer_mgm!OFFICENO2)) > 2 Then
'            TDBMaskOffice(1).ReadOnly = True
'        End If
'        m_customer_mgm("EXTOFFICE2") = TxtExt(1).Text
'        m_customer_mgm("FAXNO") = TDBMaskFax(0).Value
'        m_customer_mgm("FAXNO2") = TDBMaskFax(1).Value
'        m_customer_mgm("PRIOR") = Combo5.Text
'        m_customer_mgm("NAMAPT") = Text1(20).Text
'        m_customer_mgm("ADDRPT") = RichTextBox1(2).Text
'        m_customer_mgm("AHOMENO") = TDBMaskAHome(0).Value
'        m_customer_mgm("AHOMENO2") = TDBMaskAHome(1).Value
'        m_customer_mgm("AOFFICENO") = TDBMaskAOffice(0).Value
'        m_customer_mgm("AOFFICENO2") = TDBMaskAOffice(1).Value
'        m_customer_mgm("AFAXNO") = TDBMaskAFax(0).Value
'        m_customer_mgm("AFAXNO2") = TDBMaskAFax(1).Value
'        m_customer_mgm("TGLCALL") = Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") & " " & Format(Now, "hh:nn")
'        m_customer_mgm("F_TELP1") = CekTelp(0).Value
'        m_customer_mgm("F_TELP2") = CekTelp(1).Value
'        m_customer_mgm("F_TELP3") = CekTelp(2).Value
'        m_customer_mgm("F_TELP4") = CekTelp(3).Value
'        m_customer_mgm("F_TELP5") = CekTelp(4).Value
'        m_customer_mgm("F_TELP6") = CekTelp(5).Value
'        m_customer_mgm("F_TELP7") = CekTelp(6).Value
'        m_customer_mgm("F_TELP8") = CekTelp(7).Value
'
'        If Check2(0).Value Then
'            m_customer_mgm("RECSTATUS") = "C"
'            pStatusLstCall = cmbContacted(0).Text
'        End If
'        If Check2(1).Value Then
'            m_customer_mgm("RECSTATUS") = "N"
'            pStatusLstCall = CmbNotContacted(0).Text
'        End If
'        If Check2(3).Value Then
'            m_customer_mgm("RECSTATUS") = "DA"
'            pStatusLstCall = CmbDisagree(0).Text
'        End If
'        If Check2(2).Value Then
'            m_customer_mgm("RECSTATUS") = "RR"
'            pStatusLstCall = "RR"
'            m_customer_mgm("INCOMING") = "RR"
'            m_customer_mgm("TGLINCOMING") = Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") & " " & Format(Now, "hh:nn")
'        End If
'        If Trim(UCase(IIf(IsNull(m_customer_mgm("KETHSLKERJA")), "", m_customer_mgm("KETHSLKERJA")))) = Trim(UCase(pStatusLstCall)) Then
'        Else
'            m_customer_mgm("KETHSLKERJA") = pStatusLstCall
'            m_customer_mgm("TGLSTATUS") = Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd") & " " & Format(Now, "hh:nn")
'        End If
'        pStatusHstLstCall = m_customer_mgm("KETHSLKERJA")
'        m_customer_mgm("KD_CLS") = CmbDtQuality(0).Text
'        m_customer_mgm("PRIOR") = Combo5.Text
'        m_customer_mgm("NEXTACT") = Combo7.Text
'        m_customer_mgm("REMARKS") = RichTextBox1(3).Text
'        m_customer_mgm!NEXTACTDATE = Format(TDBDate1(3).Value, "yyyy/mm/dd") & " " & Format(Now, "hh:nn")
'
'    m_customer_mgm.UPDATE
'
''M_DATA.UPDATE_CUSTOMER_BARU M_OBJCONN, KETHSLKERJA, STATUS_FIELD_LAMA, M_CALL, M_STATUS, DOK1
'If Check2(1).Value = 1 Then
'    If RichTextBox1(3).Text <> Empty Or Combo7.Text <> Empty Then
'        m_data.ADD_HISTORY M_OBJCONN, Text1(1).Text, MDIForm1.TDBDate1.Text, Time, MDIForm1.Text1.Text, "MGM", Combo7.Text & " " & RichTextBox1().Text
'    End If
'Else
'    m_data.ADD_HISTORY M_OBJCONN, Text1(1).Text, MDIForm1.TDBDate1.Text, Time, MDIForm1.Text1.Text, "MGM", Combo7.Text & " " & RichTextBox1(3).Text, Combo1(0).Text
'End If
'M_OBJCONN.CommitTrans
'MsgBox "Data Sudah Tersimpan", vbInformation + vbOKOnly, "Sukses"
''If Status_Form = 1 Then
''    VIEW_MGMDATA.ListView1.SelectedItem.SubItems(8) = Combo7.Text
''    VIEW_MGMDATA.ListView1.SelectedItem.SubItems(9) = RichTextBox1(3).Text
''Else
' '  If Status_Form = 2 Then
'        VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(4) = Format(TDBDate1(3).Value, "yyyy/mm/dd") & " " & Format(Now, "hh:nn")
'        VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(5) = Combo7.Text
'        VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(6) = RichTextBox1(3).Text
'        VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(10) = Format(MDIForm1.TDBDate1.Value, "yyyy/mm/dd")
'        VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(11) = pStatusHstLstCall
' '  End If
''End If
'M_followup = False
'pStatusLstCall = ""
'pStatusHstLstCall = ""
'RichTextBox1(3).Text = Empty
'Combo7.Text = Empty
'Set m_data = Nothing
'Exit Sub
'editErr:
'    M_OBJCONN.RollbackTrans
'    MsgBox Err.Description
'End Sub
'
'Private Sub TDBDate1_Click(Index As Integer)
'Dim tahun As Integer
'Dim tahunlhr As Integer
'Select Case Index
'Case 0
'    tahun = Year(Date)
'    If TDBDate1(0).ValueIsNull Then
'        Text1(2).Text = "0"
'    Else
'        tahunlhr = Year(TDBDate1(0).Value)
'        Text1(2).Text = CStr(tahun - tahunlhr)
'    End If
'End Select
'End Sub
'
'Private Sub Form_Load()
'Dim listitem As listitem
'    Me.MousePointer = vbHourglass
'    Call ChangeTab(SSTab1)
'    TDBDate1(3).Value = ""
'    TDBTime1.Value = ""
'    Call HEADER_HISTORY
'    Call ISI_COMBO_DATASOURCE
'    Call ISI_COMBO_PRODUCT_CLOSING
'    Call HEADER_VIEW_Refferall
'    Combo5.AddItem "High"
'    Combo5.AddItem "Normal"
'    Combo5.AddItem "Low"
'    Combo5.Text = "Normal"
'    Combo1(0).Text = "MGM-REF"
'    Combo1(1).Text = "MGM-REF"
'    M_followup = False
'    If reff_Duplikasi = True Then
'        SSCommand1(4).Visible = True
'        TxtReasonClosing.Visible = True
'        SSCommand1(1).Visible = False
'        SSCommand1(0).Visible = False
'        SSCommand1(2).Visible = False
'        ListView1(0).Enabled = False
'    Else
'        If reff_Duplikasi1 = True Then
'            SSCommand1(5).Visible = True
'            TxtReasonClosing.Visible = True
'            SSCommand1(4).Visible = False
'            SSCommand1(1).Visible = False
'            SSCommand1(0).Visible = False
'            SSCommand1(2).Visible = False
'            ListView1(0).Enabled = False
'        Else
'            SSCommand1(4).Visible = False
'            SSCommand1(5).Visible = False
'            SSCommand1(1).Visible = True
'            SSCommand1(0).Visible = True
'            SSCommand1(2).Visible = True
'        End If
'    End If
'
'    If ADD_CUST Then
'    Else
'   '     SSCommand1(2).Enabled = False
'        Call VIEW_DATA_CUST
'        If UCase(MDIForm1.Text2.Text) = "AGENT" Then
'            Text1(0).Enabled = False
'            Combo1(0).Visible = False
'            Combo1(1).Visible = False
'            Text1(3).Visible = True
'            If Len(TDBMaskHome(0).Value) > 5 Then
'                TDBMaskHome(0).Enabled = False
'            End If
'            If Len(TDBMaskHome(1).Value) > 5 Then
'                TDBMaskHome(1).Enabled = False
'            End If
'            If Len(TDBMaskMobile(0).Value) > 5 Then
'                TDBMaskMobile(0).Enabled = False
'            End If
'            If Len(TDBMaskMobile(1).Value) > 5 Then
'                TDBMaskMobile(1).Enabled = False
'            End If
'            If Len(TDBMaskOffice(0).Value) > 5 Then
'                TDBMaskOffice(0).Enabled = False
'            End If
'            If Len(TDBMaskOffice(1).Value) > 5 Then
'                TDBMaskOffice(1).Enabled = False
'            End If
'            If Len(TDBMaskFax(0).Value) > 5 Then
'                TDBMaskFax(0).Enabled = False
'            End If
'            If Len(TDBMaskFax(1).Value) > 5 Then
'                TDBMaskFax(1).Enabled = False
'            End If
'        End If
'    End If
'SSTab1.Tab = 0
'Me.MousePointer = vbNormal
'End Sub
'
'Private Sub HEADER_HISTORY()
'    ListView1(1).ColumnHeaders.ADD 1, , "Tanggal Jam", 15 * TXT
''    ListView1(1).ColumnHeaders.ADD 2, , "Jam", 8 * TXT
'    ListView1(1).ColumnHeaders.ADD 2, , "History", 30 * TXT
'    ListView1(1).ColumnHeaders.ADD 3, , "Agent", 10 * TXT
'    ListView1(1).ColumnHeaders.ADD 4, , "Sts Call", 10 * TXT
'    ListView1(1).ColumnHeaders.ADD 5, , "Complaint Note", 20 * TXT
'End Sub
'
'
'Private Function ILANGIN_AREA(TELP As String) As String
'    ILANGIN_AREA = Replace(TELP, "021", "")
'End Function
'
'Private Sub VIEW_DATA_CUST()
'Dim m_objrs1 As ADODB.Recordset
'Dim m_data As New CLS_FRMCUST_CC_MGM
'Dim listitem As listitem
'If TodayList = True Then
'    Set m_customer_mgm = m_data.QUERY_CUST(M_OBJCONN, "CUSTID = '" + FrmTodayList.LstVwSearchMgm.SelectedItem.SubItems(1) + "'")
'Else
'    If reff_Duplikasi = True Then
'        Set m_customer_mgm = m_data.QUERY_CUST(M_OBJCONN, "CUSTID = '" + FrmDuplikasiCh.LstExLeads.SelectedItem.SubItems(1) + "'")
'        SSCommand1(2).Visible = False
'        SSCommand1(0).Visible = False
'        SSCommand1(4).Visible = True
'    Else
'        If reff_Duplikasi1 = True Then
'                Set m_customer_mgm = m_data.QUERY_CUST(M_OBJCONN, "CUSTID = '" + FrmDuplikasiPendingCh.LstExLeads.SelectedItem.SubItems(1) + "'")
'                SSCommand1(2).Visible = False
'                SSCommand1(0).Visible = False
'                SSCommand1(5).Visible = True
'        Else
'            If Flag_Mgm = True Then
'                Set m_customer_mgm = m_data.QUERY_CUST(M_OBJCONN, "CUSTID = '" + VIEW_MGMDATA.ListView1.SelectedItem.SubItems(1) + "'")
'            Else
'                Set m_customer_mgm = m_data.QUERY_CUST(M_OBJCONN, "CUSTID = '" + VIEW_MGMDATA.LstVwSearchMgm.SelectedItem.SubItems(1) + "'")
'            End If
'        End If
'    End If
'End If
'    If m_customer_mgm.RecordCount <> 0 Then
'        If IIf(IsNull(m_customer_mgm("F_TELP1")), 0, m_customer_mgm("F_TELP1")) = 1 Then
'            CekTelp(0).Value = 1
'        End If
'        If IIf(IsNull(m_customer_mgm("F_TELP2")), 0, m_customer_mgm("F_TELP2")) = 1 Then
'            CekTelp(1).Value = 1
'        End If
'        If IIf(IsNull(m_customer_mgm("F_TELP3")), 0, m_customer_mgm("F_TELP3")) = 1 Then
'            CekTelp(2).Value = 1
'        End If
'        If IIf(IsNull(m_customer_mgm("F_TELP4")), 0, m_customer_mgm("F_TELP4")) = 1 Then
'            CekTelp(3).Value = 1
'        End If
'        If IIf(IsNull(m_customer_mgm("F_TELP5")), 0, m_customer_mgm("F_TELP5")) = 1 Then
'            CekTelp(4).Value = 1
'        End If
'        If IIf(IsNull(m_customer_mgm("F_TELP6")), 0, m_customer_mgm("F_TELP6")) = 1 Then
'            CekTelp(5).Value = 1
'        End If
'        If IIf(IsNull(m_customer_mgm("F_TELP7")), 0, m_customer_mgm("F_TELP7")) = 1 Then
'            CekTelp(6).Value = 1
'        End If
'        If IIf(IsNull(m_customer_mgm("F_TELP8")), 0, m_customer_mgm("F_TELP8")) = 1 Then
'            CekTelp(7).Value = 1
'        End If
'
'        Text1(1).Text = IIf(IsNull(m_customer_mgm("CUSTID")), "", m_customer_mgm("CUSTID"))
'        Text1(0).Text = IIf(IsNull(m_customer_mgm("NAME")), "", m_customer_mgm("NAME"))
'        TDBDate1(0).Value = IIf(IsNull(m_customer_mgm("BIRTHD")), "", Format(m_customer_mgm("BIRTHD"), "dd-mmm-yyyy"))
'        Call TDBDate1_Click(0)
'        Combo1(0).Text = IIf(IsNull(m_customer_mgm("RECSOURCE")), "", m_customer_mgm("RECSOURCE"))
'        Call Combo1_Click(0)
'        Text1(3).Text = IIf(IsNull(m_customer_mgm("RECSOURCE")), "", m_customer_mgm("RECSOURCE")) + "     [ " + ket + " ]"
'        TxtAlamat.Text = IIf(IsNull(m_customer_mgm("ADDRNOW")), "", m_customer_mgm("ADDRNOW"))
'        TxtZip.Text = IIf(IsNull(m_customer_mgm("ZIPNOW")), "", m_customer_mgm("ZIPNOW"))
'        TxtCity.Text = IIf(IsNull(m_customer_mgm("CITYNOW")), "", m_customer_mgm("CITYNOW"))
'        TDBMaskHome(0).Value = ILANGIN_AREA(IIf(IsNull(m_customer_mgm("HOMENO")), "", m_customer_mgm("HOMENO")))
'        If Left(TDBMaskHome(0).Text, 3) = "021" Then
'            TDBMaskHome(0).Value = ILANGIN_AREA(CStr(TDBMaskHome(0).Value))
'        End If
'        TDBMaskHome(1).Value = IIf(IsNull(m_customer_mgm("HOMENO2")), "", m_customer_mgm("HOMENO2"))
'        TDBMaskMobile(0).Value = IIf(IsNull(m_customer_mgm("MOBILENO")), "", m_customer_mgm("MOBILENO"))
'        TDBMaskMobile(1).Value = IIf(IsNull(m_customer_mgm("MOBILENO2")), "", m_customer_mgm("MOBILENO2"))
'        TDBMaskOffice(0).Value = ILANGIN_AREA(IIf(IsNull(m_customer_mgm("OFFICENO")), "", m_customer_mgm("OFFICENO")))
'        If Left(TDBMaskOffice(0).Text, 3) = "021" Then
'            TDBMaskOffice(0).Value = ILANGIN_AREA(CStr(TDBMaskOffice(0).Value))
'        End If
'        TxtExt(0).Text = IIf(IsNull(m_customer_mgm("EXTOFFICE")), "", m_customer_mgm("EXTOFFICE"))
'        TDBMaskOffice(1).Value = IIf(IsNull(m_customer_mgm("OFFICENO2")), "", m_customer_mgm("OFFICENO2"))
'        TxtExt(1).Text = IIf(IsNull(m_customer_mgm("EXTOFFICE2")), "", m_customer_mgm("EXTOFFICE2"))
'
'        TDBMaskFax(0).Value = IIf(IsNull(m_customer_mgm("FAXNO")), "", m_customer_mgm("FAXNO"))
'        TDBMaskFax(1).Value = IIf(IsNull(m_customer_mgm("FAXNO2")), "", m_customer_mgm("FAXNO2"))
'        Combo5.Text = IIf(IsNull(m_customer_mgm("PRIOR")), "", m_customer_mgm("PRIOR"))
'        If Combo5.Text = Empty Then
'            Combo5.Text = "Normal"
'        End If
'        Text1(20).Text = IIf(IsNull(m_customer_mgm("NAMAPT")), "", m_customer_mgm("NAMAPT"))
'        RichTextBox1(2).Text = IIf(IsNull(m_customer_mgm("ADDRPT")), "", m_customer_mgm("ADDRPT"))
'        TDBMaskAHome(0).Value = IIf(IsNull(m_customer_mgm("AHOMENO")), "", m_customer_mgm("AHOMENO"))
'        TDBMaskAHome(1).Value = IIf(IsNull(m_customer_mgm("AHOMENO2")), "", m_customer_mgm("AHOMENO2"))
'        TDBMaskAOffice(0).Value = IIf(IsNull(m_customer_mgm("AOFFICENO")), "", m_customer_mgm("AOFFICENO"))
'        TDBMaskAOffice(1).Value = IIf(IsNull(m_customer_mgm("AOFFICENO2")), "", m_customer_mgm("AOFFICENO2"))
'        TDBMaskAFax(0).Value = IIf(IsNull(m_customer_mgm("AFAXNO")), "", m_customer_mgm("AFAXNO"))
'        TDBMaskAFax(1).Value = IIf(IsNull(m_customer_mgm("AFAXNO2")), "", m_customer_mgm("AFAXNO2"))
'        Text2.Text = IIf(IsNull(m_customer_mgm("others")), "", m_customer_mgm("others"))
'        Select Case m_customer_mgm!RECSTATUS
'        Case "N"
'            Check2(1).Value = 1
'            CmbNotContacted(0).Text = IIf(IsNull(m_customer_mgm("KETHSLKERJA")), "", m_customer_mgm("KETHSLKERJA"))
'            Call CmbNotContacted_LostFocus(0)
'        Case "C"
'            Check2(0).Value = 1
'            cmbContacted(0).Text = IIf(IsNull(m_customer_mgm("KETHSLKERJA")), "", m_customer_mgm("KETHSLKERJA"))
'            Call CmbContacted_LostFocus(0)
'        Case "DA"
'            Check2(3).Value = 1
'            CmbDisagree(0).Text = IIf(IsNull(m_customer_mgm("KETHSLKERJA")), "", m_customer_mgm("KETHSLKERJA"))
'            Call CmbDisagree_LostFocus(0)
'        Case "RR"
'            Check2(2).Value = 1
'            If UCase(Trim(MDIForm1.Text2.Text)) = "AGENT" Then
'                Check2(0).Enabled = False
'                Check2(1).Enabled = False
'                Check2(2).Enabled = False
'            End If
'            TDBDate1(1).Value = IIf(IsNull(m_customer_mgm("TGLSTATUS")), "", Format(m_customer_mgm("TGLSTATUS"), "dd-mmm-yyyy"))
'        End Select
'        CmbDtQuality(0).Text = IIf(IsNull(m_customer_mgm("KD_CLS")), "", m_customer_mgm("KD_CLS"))
'        Call CmbDtQuality_LostFocus(0)
'        pStatusLstCall = IIf(IsNull(m_customer_mgm("KETHSLKERJA")), "", m_customer_mgm("KETHSLKERJA"))
'    ' isi history
'            Set m_objrs1 = m_data.QUERY_HIST_CUST(M_OBJCONN, "CUSTID = '" + Text1(1).Text + "'")
'            While Not m_objrs1.EOF
'                Set listitem = ListView1(1).ListItems.ADD(, , Left(m_objrs1("DATETIME"), 4) & "/" & Mid(m_objrs1("DATETIME"), 5, 2) & "/" & IIf(IsNull(m_objrs1("DATETIME")), "", Mid(m_objrs1("DATETIME"), 7, 2)) & " " & IIf(IsNull(m_objrs1("DATETIME")), "", Mid(m_objrs1("DATETIME"), 9, 2)) & ":" & Right(m_objrs1("DATETIME"), 2))
'                    listitem.SubItems(1) = IIf(IsNull(m_objrs1("HST")), "", m_objrs1("HST"))
'                    listitem.SubItems(2) = IIf(IsNull(m_objrs1("AGENT")), "", m_objrs1("AGENT"))
'                    listitem.SubItems(3) = IIf(IsNull(m_objrs1("KdComplaint")), "", m_objrs1("KdComplaint"))
'                    listitem.SubItems(4) = IIf(IsNull(m_objrs1("RemarkComplaint")), "", m_objrs1("RemarkComplaint"))
'            m_objrs1.MoveNext
'            Wend
'            Set m_objrs1 = Nothing
'    'isi refferall
'        Call show_Reff
'
'    End If
'Set m_data = Nothing
'ket = ""
'End Sub
'
'
'Private Sub ISI_COMBO_DATASOURCE()
'Dim m_objrs As ADODB.Recordset
'Dim m_data As New CLS_FRMCUST_CC_MGM
'    Set m_objrs = m_data.QUERY_COMBO_DATASOURCE_ISI(M_OBJCONN, "")
'    While Not m_objrs.EOF
'        Combo1(0).AddItem m_objrs("KODEDS")
'        Combo1(0).DataField = m_objrs("KODEDS")
'        Combo1(1).AddItem IIf(IsNull(m_objrs("KETERANGAN")), "", m_objrs("KETERANGAN"))
'        Combo1(1).DataField = m_objrs("KETERANGAN")
'        m_objrs.MoveNext
'    Wend
'Set m_objrs = Nothing
'Set m_objrs = New ADODB.Recordset
'm_objrs.CursorLocation = adUseClient
'm_objrs.Open "Select * from ComplaintCode", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'While Not m_objrs.EOF
'    Combo6.AddItem m_objrs("KdComplaint")
'    m_objrs.MoveNext
'Wend
'Set m_objrs = Nothing
'
'Set m_objrs = New ADODB.Recordset
'm_objrs.CursorLocation = adUseClient
'm_objrs.Open "Select * from StsNextAct", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'While Not m_objrs.EOF
'    Combo7.AddItem m_objrs("NmStsNextAct")
'    m_objrs.MoveNext
'Wend
'Set m_objrs = Nothing
'
'Set m_data = Nothing
'End Sub
'
'Private Sub ISI_COMBO_PRODUCT_CLOSING()
'Dim m_objrs As ADODB.Recordset
'Dim m_data As New CLS_FRMCUST_CC_MGM
'    Set m_objrs = m_data.QUERY_COMBO_CLOSSING(M_OBJCONN, " UPPER(JENIS) <> 'LEADS'")
'    While Not m_objrs.EOF
'        CmbNotContacted(0).AddItem m_objrs("KDCLS")
'        CmbNotContacted(0).DataField = m_objrs("KDCLS")
'        CmbNotContacted(1).AddItem m_objrs("KETCLS")
'        CmbNotContacted(1).DataField = m_objrs("KETCLS")
'        m_objrs.MoveNext
'    Wend
'    Set m_objrs = Nothing
'Set m_data = Nothing
'
'Set m_objrs = New ADODB.Recordset
'm_objrs.CursorLocation = adUseClient
'm_objrs.Open "Select * from ContactedDesc WHERE UPPER(JENIS)<> 'LEADS'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'While Not m_objrs.EOF
'    cmbContacted(0).AddItem IIf(IsNull(m_objrs!KdNoProdPresented), "", m_objrs!KdNoProdPresented)
'    cmbContacted(1).AddItem IIf(IsNull(m_objrs!NmNoProdPresented), "", m_objrs!NmNoProdPresented)
'    m_objrs.MoveNext
'Wend
'Set m_objrs = Nothing
'
'Set m_objrs = New ADODB.Recordset
'm_objrs.CursorLocation = adUseClient
'm_objrs.Open "Select * from TblDisagreeMGM WHERE UPPER(JENIS)<> 'LEADS'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'While Not m_objrs.EOF
'    CmbDisagree(0).AddItem IIf(IsNull(m_objrs!kddisagree), "", m_objrs!kddisagree)
'    CmbDisagree(1).AddItem IIf(IsNull(m_objrs!ketdisagree), "", m_objrs!ketdisagree)
'    m_objrs.MoveNext
'Wend
'Set m_objrs = Nothing
'
'
'Set m_objrs = New ADODB.Recordset
'm_objrs.CursorLocation = adUseClient
'm_objrs.Open "Select * from DataQuality", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'While Not m_objrs.EOF
'    CmbDtQuality(0).AddItem IIf(IsNull(m_objrs!KdDataQuality), "", m_objrs!KdDataQuality)
'    CmbDtQuality(1).AddItem IIf(IsNull(m_objrs!NmDataQuality), "", m_objrs!NmDataQuality)
'    m_objrs.MoveNext
'Wend
'Set m_objrs = Nothing
'Set m_data = Nothing
'End Sub
'
'Private Sub SSTab1_Click(PreviousTab As Integer)
'    Call ChangeTab(SSTab1)
'End Sub
'
'Private Sub ChangeTab(SSTab As SSTab)
'    Dim ctrl As Control, TabIndex As Long
'    TabIndex = 99999          ' A very high value.
'    On Error Resume Next
'    For Each ctrl In SSTab.Parent.Controls
'        If ctrl.Container Is SSTab Then
'            If ctrl.Left < -10000 Then
'                ctrl.Enabled = False
'            Else
'                ctrl.Enabled = True
'                If ctrl.TabIndex >= TabIndex Then
'                Else
'                    TabIndex = ctrl.TabIndex
'                    ctrl.SetFocus
'                End If
'            End If
'        End If
'    Next
'End Sub
'
'Private Sub TDBTime1_LostFocus()
'    If TDBTime1.ValueIsNull Then
'        TDBTime1.Value = Format(Time, "hh:mm")
'    End If
'End Sub
'
'Private Sub Text1_LostFocus(Index As Integer)
'Select Case Index
'    Case 0
'        Text1(Index).Text = UCase(Text1(Index).Text)
'End Select
'End Sub
'
'
Private Sub SSCommand1_Click(Index As Integer)

End Sub
