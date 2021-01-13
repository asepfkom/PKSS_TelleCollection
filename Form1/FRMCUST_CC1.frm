VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FRMCUST_CC1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   10500
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   Icon            =   "FRMCUST_CC1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text9 
      Height          =   420
      Left            =   3405
      TabIndex        =   114
      Text            =   "Text9"
      Top             =   855
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   112
      Top             =   915
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.TextBox Text7 
      Height          =   405
      Left            =   7455
      TabIndex        =   111
      Top             =   900
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1650
      Locked          =   -1  'True
      TabIndex        =   110
      Top             =   45
      Width           =   6675
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
      Picture         =   "FRMCUST_CC1.frx":0442
      Caption         =   "&Exit"
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
      Picture         =   "FRMCUST_CC1.frx":059C
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
      Top             =   30
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
      Picture         =   "FRMCUST_CC1.frx":08BE
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
      Picture         =   "FRMCUST_CC1.frx":1132
      Caption         =   "&MGM"
      AutoSize        =   2
      Alignment       =   4
      ButtonStyle     =   2
      PictureAlignment=   1
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   150
      TabIndex        =   48
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
         Sorted          =   -1  'True
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
         TabIndex        =   50
         Top             =   210
         Width           =   1080
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
         TabIndex        =   49
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
      Sorted          =   -1  'True
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
      Sorted          =   -1  'True
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
      Height          =   4635
      Left            =   60
      TabIndex        =   10
      Top             =   1350
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   8176
      _Version        =   393216
      Tabs            =   5
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
      TabCaption(0)   =   "Personal Data"
      TabPicture(0)   =   "FRMCUST_CC1.frx":15FA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Working Desc"
      TabPicture(1)   =   "FRMCUST_CC1.frx":1616
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Note"
      TabPicture(2)   =   "FRMCUST_CC1.frx":1632
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&History"
      TabPicture(3)   =   "FRMCUST_CC1.frx":164E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame9"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "&Result"
      TabPicture(4)   =   "FRMCUST_CC1.frx":166A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame10"
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame10 
         Height          =   4155
         Left            =   -74910
         TabIndex        =   76
         Top             =   375
         Width           =   10305
         Begin VB.CommandButton Command1 
            Caption         =   "Appl &Form"
            Height          =   375
            Left            =   6690
            TabIndex        =   113
            Top             =   1125
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Frame Frame13 
            Caption         =   "Frame11"
            Height          =   1095
            Left            =   135
            TabIndex        =   103
            Top             =   1680
            Width           =   10125
            Begin VB.ComboBox Combo6 
               Height          =   315
               Left            =   645
               TabIndex        =   107
               Top             =   225
               Width           =   885
            End
            Begin RichTextLib.RichTextBox RichTextBox1 
               Height          =   795
               Index           =   1
               Left            =   2115
               TabIndex        =   104
               Top             =   240
               Width           =   7950
               _ExtentX        =   14023
               _ExtentY        =   1402
               _Version        =   393217
               TextRTF         =   $"FRMCUST_CC1.frx":1686
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
               TabIndex        =   108
               Top             =   255
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
               TabIndex        =   106
               Top             =   255
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
               TabIndex        =   105
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
            Left            =   210
            MouseIcon       =   "FRMCUST_CC1.frx":1701
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   102
            Top             =   450
            Width           =   1260
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Incoming"
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
            Index           =   2
            Left            =   4035
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   100
            Top             =   1155
            Width           =   1095
         End
         Begin VB.Frame Frame11 
            Caption         =   "Frame11"
            Height          =   1260
            Left            =   4770
            TabIndex        =   97
            Top             =   2775
            Width           =   5490
            Begin RichTextLib.RichTextBox RichTextBox1 
               Height          =   975
               Index           =   3
               Left            =   45
               TabIndex        =   98
               Top             =   240
               Width           =   5415
               _ExtentX        =   9551
               _ExtentY        =   1720
               _Version        =   393217
               TextRTF         =   $"FRMCUST_CC1.frx":1B43
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
               TabIndex        =   99
               Top             =   0
               Width           =   825
            End
         End
         Begin VB.Frame Frame12 
            Height          =   615
            Left            =   60
            TabIndex        =   80
            Top             =   480
            Width           =   5340
            Begin VB.ComboBox Combo2 
               Height          =   315
               Index           =   1
               Left            =   765
               Sorted          =   -1  'True
               TabIndex        =   92
               Text            =   "Combo2"
               Top             =   240
               Width           =   4170
            End
            Begin VB.ComboBox Combo2 
               Height          =   315
               Index           =   0
               Left            =   945
               TabIndex        =   91
               Text            =   "Combo2"
               Top             =   240
               Width           =   1110
            End
            Begin VB.Label Label9 
               Caption         =   "Desc:"
               ForeColor       =   &H00800080&
               Height          =   225
               Index           =   0
               Left            =   285
               TabIndex        =   93
               Top             =   300
               Width           =   435
            End
         End
         Begin VB.Frame Frame22 
            Height          =   1260
            Left            =   60
            TabIndex        =   81
            Top             =   2775
            Width           =   4710
            Begin VB.ComboBox Combo7 
               Height          =   315
               Left            =   1065
               TabIndex        =   109
               Top             =   165
               Width           =   3600
            End
            Begin VB.TextBox Text6 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               ForeColor       =   &H00800080&
               Height          =   225
               Left            =   390
               TabIndex        =   90
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
               TabIndex        =   89
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
               TabIndex        =   41
               Top             =   165
               Width           =   3585
            End
            Begin TDBTime6Ctl.TDBTime TDBTime1 
               Height          =   315
               Left            =   2535
               TabIndex        =   43
               Top             =   495
               Width           =   900
               _Version        =   65536
               _ExtentX        =   1587
               _ExtentY        =   556
               Caption         =   "FRMCUST_CC1.frx":1BBE
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC1.frx":1C2A
               Spin            =   "FRMCUST_CC1.frx":1C7A
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
               Value           =   0.52662037037037
            End
            Begin TDBDate6Ctl.TDBDate TDBDate1 
               Height          =   315
               Index           =   3
               Left            =   1065
               TabIndex        =   42
               Top             =   495
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               Calendar        =   "FRMCUST_CC1.frx":1CA2
               Caption         =   "FRMCUST_CC1.frx":1DBA
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FRMCUST_CC1.frx":1E26
               Keys            =   "FRMCUST_CC1.frx":1E44
               Spin            =   "FRMCUST_CC1.frx":1EA2
               AlignHorizontal =   0
               AlignVertical   =   0
               Appearance      =   1
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               CursorPosition  =   0
               DataProperty    =   0
               DisplayFormat   =   "dd-mm-yyyy"
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
               TabIndex        =   83
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
               TabIndex        =   82
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
            Left            =   5625
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   40
            Top             =   450
            Width           =   1590
         End
         Begin VB.Frame Frame25 
            Enabled         =   0   'False
            Height          =   615
            Left            =   5430
            TabIndex        =   77
            Top             =   480
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
               Left            =   585
               Sorted          =   -1  'True
               TabIndex        =   46
               Top             =   225
               Width           =   4155
            End
            Begin VB.ComboBox Combo3 
               Height          =   315
               Index           =   0
               Left            =   570
               TabIndex        =   78
               TabStop         =   0   'False
               Top             =   225
               Visible         =   0   'False
               Width           =   4155
            End
            Begin VB.Label Label7 
               Caption         =   "Desc:"
               ForeColor       =   &H00800080&
               Height          =   225
               Left            =   90
               TabIndex        =   79
               Top             =   285
               Width           =   435
            End
         End
         Begin VB.ComboBox Combo4 
            BackColor       =   &H8000000E&
            Height          =   315
            Index           =   1
            Left            =   3585
            Sorted          =   -1  'True
            TabIndex        =   95
            Text            =   "Combo4"
            Top             =   120
            Width           =   4005
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            Index           =   0
            Left            =   3585
            TabIndex        =   94
            Text            =   "Combo4"
            Top             =   120
            Width           =   1950
         End
         Begin TDBDate6Ctl.TDBDate TDBDate1 
            Height          =   315
            Index           =   1
            Left            =   5250
            TabIndex        =   101
            Top             =   1125
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   556
            Calendar        =   "FRMCUST_CC1.frx":1ECA
            Caption         =   "FRMCUST_CC1.frx":1FE2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FRMCUST_CC1.frx":204E
            Keys            =   "FRMCUST_CC1.frx":206C
            Spin            =   "FRMCUST_CC1.frx":20CA
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
            Left            =   2580
            TabIndex        =   96
            Top             =   135
            Width           =   1035
         End
      End
      Begin VB.Frame Frame6 
         Height          =   4110
         Left            =   -74910
         TabIndex        =   86
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
            Height          =   3900
            Left            =   30
            Locked          =   -1  'True
            TabIndex        =   87
            Top             =   165
            Width           =   10215
         End
      End
      Begin VB.Frame Frame9 
         Height          =   4110
         Left            =   -74940
         TabIndex        =   84
         Top             =   450
         Width           =   10320
         Begin MSComctlLib.ListView ListView1 
            Height          =   3915
            Index           =   1
            Left            =   30
            TabIndex        =   39
            Top             =   135
            Width           =   10260
            _ExtentX        =   18098
            _ExtentY        =   6906
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
      Begin VB.Frame Frame5 
         Height          =   4185
         Left            =   -74955
         TabIndex        =   62
         Top             =   360
         Width           =   10350
         Begin VB.OptionButton Option4 
            Caption         =   "Entrepreneur Data"
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
            Width           =   2130
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Official Data"
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
            Height          =   1695
            Left            =   3075
            TabIndex        =   67
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
               Caption         =   "FRMCUST_CC1.frx":20F2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC1.frx":215E
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
               Caption         =   "FRMCUST_CC1.frx":21A0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC1.frx":220C
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
               Caption         =   "FRMCUST_CC1.frx":224E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC1.frx":22BA
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
               Caption         =   "FRMCUST_CC1.frx":22FC
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC1.frx":2368
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
               Caption         =   "FRMCUST_CC1.frx":23AA
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC1.frx":2416
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
               Caption         =   "FRMCUST_CC1.frx":2458
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC1.frx":24C4
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
               Caption         =   "FRMCUST_CC1.frx":2506
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC1.frx":2572
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
               Caption         =   "FRMCUST_CC1.frx":25B4
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC1.frx":2620
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
               Left            =   2730
               TabIndex        =   70
               Top             =   255
               Width           =   270
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
               TabIndex        =   69
               Top             =   240
               Width           =   975
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
               Left            =   480
               TabIndex        =   68
               Top             =   930
               Width           =   585
            End
         End
         Begin VB.Frame Frame8 
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
            TabIndex        =   63
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
               TextRTF         =   $"FRMCUST_CC1.frx":2662
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
               TabIndex        =   66
               Top             =   1260
               Width           =   1155
            End
            Begin VB.Label Label4 
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
               Height          =   255
               Index           =   4
               Left            =   195
               TabIndex        =   65
               Top             =   690
               Width           =   1080
            End
            Begin VB.Label Label3 
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
               TabIndex        =   64
               Top             =   360
               Width           =   1020
            End
         End
         Begin VB.Frame Frame7 
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
            TabIndex        =   71
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
               TabIndex        =   44
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
               TabIndex        =   45
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
               TextRTF         =   $"FRMCUST_CC1.frx":26DD
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
               Caption         =   "Company Name"
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
               TabIndex        =   75
               Top             =   285
               Width           =   1320
            End
            Begin VB.Label Label4 
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
               Left            =   105
               TabIndex        =   74
               Top             =   570
               Width           =   1365
            End
            Begin VB.Label Label3 
               Caption         =   "Gaji / Bulan"
               ForeColor       =   &H00800080&
               Height          =   225
               Index           =   20
               Left            =   105
               TabIndex        =   73
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
               TabIndex        =   72
               Top             =   2130
               Visible         =   0   'False
               Width           =   1095
            End
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4155
         Left            =   45
         TabIndex        =   51
         Top             =   360
         Width           =   10350
         Begin VB.Frame Frame3 
            Height          =   3975
            Index           =   0
            Left            =   90
            TabIndex        =   55
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
               Left            =   2850
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
               Left            =   1470
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
               Left            =   2940
               MaxLength       =   20
               TabIndex        =   13
               Top             =   675
               Width           =   2520
            End
            Begin TDBDate6Ctl.TDBDate TDBDate1 
               Height          =   315
               Index           =   0
               Left            =   1455
               TabIndex        =   14
               Top             =   1005
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   556
               Calendar        =   "FRMCUST_CC1.frx":2758
               Caption         =   "FRMCUST_CC1.frx":2870
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FRMCUST_CC1.frx":28DC
               Keys            =   "FRMCUST_CC1.frx":28FA
               Spin            =   "FRMCUST_CC1.frx":2958
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
               Left            =   1455
               TabIndex        =   11
               Top             =   135
               Width           =   4020
               _ExtentX        =   7091
               _ExtentY        =   979
               _Version        =   393217
               TextRTF         =   $"FRMCUST_CC1.frx":2980
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
               TabIndex        =   56
               Top             =   1005
               Visible         =   0   'False
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   556
               _Version        =   393216
               Format          =   43974656
               CurrentDate     =   37459
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
               Left            =   3285
               TabIndex        =   61
               Top             =   1050
               Width           =   795
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "Date Of Birth :"
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
               Left            =   180
               TabIndex        =   60
               Top             =   1065
               Width           =   1200
            End
            Begin VB.Label Label6 
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
               Left            =   2475
               TabIndex        =   59
               Top             =   720
               Width           =   405
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "Address :"
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
               Left            =   180
               TabIndex        =   58
               Top             =   225
               Width           =   1215
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "ZIP :"
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
               Left            =   255
               TabIndex        =   57
               Top             =   720
               Width           =   1125
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
            Height          =   3975
            Index           =   1
            Left            =   6075
            TabIndex        =   52
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
               Caption         =   "FRMCUST_CC1.frx":29FB
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC1.frx":2A67
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
               Caption         =   "FRMCUST_CC1.frx":2AA9
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC1.frx":2B15
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
               Caption         =   "FRMCUST_CC1.frx":2B57
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC1.frx":2BC3
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
               Format          =   "&&&&-&&&&&&&&&&&&&&&&&&"
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
               Text            =   "____-__________________"
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
               Caption         =   "FRMCUST_CC1.frx":2C05
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC1.frx":2C71
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
               Format          =   "&&&&-&&&&&&&&&&&&&&&&&&"
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
               Text            =   "____-__________________"
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
               Caption         =   "FRMCUST_CC1.frx":2CB3
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC1.frx":2D1F
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
               Caption         =   "FRMCUST_CC1.frx":2D61
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC1.frx":2DCD
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
               Left            =   645
               TabIndex        =   54
               Top             =   900
               Width           =   810
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
               Left            =   210
               TabIndex        =   53
               Top             =   225
               Width           =   1245
            End
         End
      End
   End
   Begin VB.TextBox Text5 
      Height          =   405
      Left            =   1740
      TabIndex        =   88
      Top             =   30
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Credit Card"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   390
      Index           =   5
      Left            =   3945
      TabIndex        =   85
      Top             =   45
      Visible         =   0   'False
      Width           =   1695
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
      TabIndex        =   47
      Top             =   1020
      Width           =   1260
   End
End
Attribute VB_Name = "FRMCUST_CC1"
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
            Frame25.Enabled = False
            Frame12.Enabled = True
            Check2(2).Enabled = False
            TDBDate1(1).Text = Empty
            TDBDate1(1).Enabled = False
        Else
            Check2(1).Enabled = True
            Check2(1).Value = 0
            Combo2(0).Text = Empty
            Combo2(1).Text = Empty
            Check2(2).Enabled = True
            TDBDate1(1).Enabled = True
        End If
    Case 1
        If Check2(Index).Value Then
            Check2(0).Enabled = False
            Check2(0).Value = 0
            Frame25.Enabled = True
            Frame12.Enabled = False
            Check2(2).Enabled = False
            TDBDate1(1).Text = Empty
            TDBDate1(1).Enabled = False
        Else
            Check2(0).Enabled = True
            Check2(0).Value = 0
            Frame25.Enabled = False
            Combo1(0).Text = Empty
            Combo3(1).Text = Empty
            Check2(2).Enabled = True
            TDBDate1(1).Enabled = True
        End If
    Case 2
        If Check2(Index).Value Then
            Check2(0).Enabled = False
            Check2(0).Value = 0
            Check2(1).Enabled = False
            Check2(1).Value = 0
            Frame25.Enabled = True
            Frame12.Enabled = False
        Else
            Check2(0).Enabled = True
            Check2(0).Value = 0
            Check2(1).Enabled = True
            Check2(1).Value = 0
            Frame25.Enabled = False
            Combo1(0).Text = Empty
            Combo3(1).Text = Empty
            TDBDate1(1).Value = Empty
        End If
End Select
End Sub

Private Sub Combo1_Click(Index As Integer)
Dim M_DATA As New CLS_FRMCUST_CC
Dim m_objrs As ADODB.Recordset
Select Case Index
    Case 0
        Set m_objrs = M_DATA.QUERY_COMBO_DATASOURCE(M_OBJCONN, "KODEDS = '" + Combo1(Index).Text + "'")
        If m_objrs.RecordCount <> 0 Then
            Combo1(0).Text = m_objrs("KODEDS")
            Combo1(1).Text = IIf(IsNull(m_objrs("KETERANGAN")), "", m_objrs("KETERANGAN"))
            Text1(3).Text = IIf(IsNull(m_objrs("KETERANGAN")), "", m_objrs("KETERANGAN"))
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
            Text1(3).Text = Empty
        End If
    Case 1
        Set m_objrs = M_DATA.QUERY_COMBO_DATASOURCE(M_OBJCONN, "KETERANGAN = '" + Combo1(Index).Text + "'")
        If m_objrs.RecordCount <> 0 Then
            Combo1(0).Text = m_objrs("KODEDS")
            Combo1(1).Text = IIf(IsNull(m_objrs("KETERANGAN")), "", m_objrs("KETERANGAN"))
            Text1(3).Text = IIf(IsNull(m_objrs("KETERANGAN")), "", m_objrs("KETERANGAN"))
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
Set m_objrs = Nothing
Set M_DATA = Nothing
End Sub

Private Sub Combo1_GotFocus(Index As Integer)
SendKeys "{Home}+{End}"
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim M_DATA As New CLS_FRMCUST_CC
Dim m_objrs As ADODB.Recordset
Select Case Index
    Case 0
        Set m_objrs = M_DATA.QUERY_COMBO_DATASOURCE(M_OBJCONN, "KODEDS = '" + Combo1(Index).Text + "'")
        If m_objrs.RecordCount <> 0 Then
            Combo1(0).Text = m_objrs("KODEDS")
            Combo1(1).Text = IIf(IsNull(m_objrs("KETERANGAN")), "", m_objrs("KETERANGAN"))
        Else
            Combo1(0).Text = Empty
            Combo1(1).Text = Empty
        End If
    Case 1
        Set m_objrs = M_DATA.QUERY_COMBO_DATASOURCE(M_OBJCONN, "KETERANGAN = '" + Combo1(Index).Text + "'")
        If m_objrs.RecordCount <> 0 Then
            Combo1(0).Text = m_objrs("KODEDS")
            Combo1(1).Text = IIf(IsNull(m_objrs("KETERANGAN")), "", m_objrs("KETERANGAN"))
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
Set m_objrs = Nothing
Set M_DATA = Nothing
End Sub

Private Sub Combo2_Click(Index As Integer)
Call Combo2_LostFocus(Index)
End Sub

Private Sub Combo2_LostFocus(Index As Integer)
Dim m_objrs As ADODB.Recordset
Select Case Index
Case 0
    Set m_objrs = New ADODB.Recordset
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open "Select * from aiu_NoProdPresented WHERE KdNoProdPresented ='" + Combo2(Index).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If m_objrs.RecordCount <> 0 Then
        Combo2(0).Text = IIf(IsNull(m_objrs!KdNoProdPresented), "", m_objrs!KdNoProdPresented)
        Combo2(1).Text = IIf(IsNull(m_objrs!NmNoProdPresented), "", m_objrs!NmNoProdPresented)
    Else
        Combo2(0).Text = Empty
        Combo2(1).Text = Empty
    End If
    Set m_objrs = Nothing
Case 1
    Set m_objrs = New ADODB.Recordset
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open "Select * from aiu_NoProdPresented WHERE NmNoProdPresented ='" + Combo2(Index).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If m_objrs.RecordCount <> 0 Then
        Combo2(0).Text = IIf(IsNull(m_objrs!KdNoProdPresented), "", m_objrs!KdNoProdPresented)
        Combo2(1).Text = IIf(IsNull(m_objrs!NmNoProdPresented), "", m_objrs!NmNoProdPresented)
    Else
        Combo2(0).Text = Empty
        Combo2(1).Text = Empty
    End If
    Set m_objrs = Nothing
End Select
End Sub

Private Sub Combo3_Click(Index As Integer)
Call Combo3_LostFocus(Index)
End Sub

Private Sub Combo3_LostFocus(Index As Integer)
Dim M_DATA As New CLS_FRMCUST_CC
Dim m_objrs As ADODB.Recordset
Select Case Index
    Case 0
        Set m_objrs = M_DATA.QUERY_COMBO_CLOSSING(M_OBJCONN, "KDCLS = '" + Combo3(Index).Text + "'")
        If m_objrs.RecordCount <> 0 Then
            Combo3(0).Text = m_objrs("KDCLS")
            Combo3(1).Text = IIf(IsNull(m_objrs("KETCLS")), "", m_objrs("KETCLS"))
        Else
            Combo3(0).Text = Empty
            Combo3(1).Text = Empty
        End If
    Case 1
        Set m_objrs = M_DATA.QUERY_COMBO_CLOSSING(M_OBJCONN, "KETCLS = '" + Combo3(Index).Text + "'")
        If m_objrs.RecordCount <> 0 Then
            Combo3(0).Text = m_objrs("KDCLS")
            Combo3(1).Text = IIf(IsNull(m_objrs("KETCLS")), "", m_objrs("KETCLS"))
        Else
            Combo3(0).Text = Empty
            Combo3(1).Text = Empty
        End If
End Select
Set m_objrs = Nothing
Set M_DATA = Nothing
End Sub

Private Sub Combo4_Click(Index As Integer)
    Call Combo4_LostFocus(Index)
End Sub

Private Sub Combo4_LostFocus(Index As Integer)
Dim m_objrs As ADODB.Recordset
Select Case Index
Case 0
    Set m_objrs = New ADODB.Recordset
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open "Select * from Aiu_DataQuality WHERE KdDataQuality ='" + Combo4(Index).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If m_objrs.RecordCount <> 0 Then
        Combo4(0).Text = IIf(IsNull(m_objrs!KdDataQuality), "", m_objrs!KdDataQuality)
        Combo4(1).Text = IIf(IsNull(m_objrs!NmDataQuality), "", m_objrs!NmDataQuality)
    Else
        Combo4(0).Text = Empty
        Combo4(1).Text = Empty
    End If
    Set m_objrs = Nothing
Case 1
    Set m_objrs = New ADODB.Recordset
    m_objrs.CursorLocation = adUseClient
    m_objrs.Open "Select * from Aiu_DataQuality WHERE NmDataQuality='" + Combo4(Index).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If m_objrs.RecordCount <> 0 Then
        Combo4(0).Text = IIf(IsNull(m_objrs!KdDataQuality), "", m_objrs!KdDataQuality)
        Combo4(1).Text = IIf(IsNull(m_objrs!NmDataQuality), "", m_objrs!NmDataQuality)
    Else
        Combo4(0).Text = Empty
        Combo4(1).Text = Empty
    End If
    Set m_objrs = Nothing
End Select
End Sub

Private Sub Combo7_Click()
    Text2.Text = Combo7.Text
End Sub

Private Sub Combo7_LostFocus()
    Text2.Text = Combo7.Text
End Sub

Private Sub Command1_Click()
    If Text1(1).Text = Empty Then
        MsgBox "Save data terlebih dahulu", vbOKOnly + vbCritical, "Telegrandi"
        Exit Sub
    End If
End Sub

Private Sub ListView1_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Select Case Index
Case 1
    ListView1(Index).SortKey = ColumnHeader.Index - 1
   ListView1(Index).Sorted = True
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


Private Function CEK_DATA_VALID_mgm() As Boolean
Dim M_MSGBOX As Variant
    If Text1(0).Text = Empty Then
        CEK_DATA_VALID_mgm = False
        MsgBox "Nama Harus Diisi", vbCritical + vbOKOnly, "Peringatan"
        Text1(0).SetFocus
        Exit Function
    End If
    If Combo1(0).Text = Empty Then
        CEK_DATA_VALID_mgm = False
        MsgBox "Sumber Info Harus Diisi", vbCritical + vbOKOnly, "Peringatan"
'        Combo1(0).SetFocus
        Exit Function
    End If
    If Combo1(2).Text = Empty Then
        CEK_DATA_VALID_mgm = False
        MsgBox "Title Harus Diisi", vbCritical + vbOKOnly, "Peringatan"
        Combo1(2).SetFocus
        Exit Function
    End If
    If Len(TDBMask1(0).Value) < 3 And Len(TDBMask1(1).Value) < 3 And Len(TDBMask1(2).Value) < 3 And Len(TDBMask1(3).Value) < 3 And Len(TDBMask1(4).Value) < 3 And Len(TDBMask1(5).Value) < 3 And Len(TDBMask1(6).Value) < 3 And Len(TDBMask1(7).Value) < 3 Then
        CEK_DATA_VALID_mgm = False
        MsgBox "Minimal Satu Nomor Telpon Harus Diisi", vbCritical + vbOKOnly, "Peringatan"
        SSTab1.Tab = 0
        TDBMask1(0).SetFocus
        Exit Function
    End If
CEK_DATA_VALID_mgm = True
End Function


Private Function CEK_DATA_VALID() As Boolean
Dim M_MSGBOX As Variant
    If Text1(0).Text = Empty Then
        CEK_DATA_VALID = False
        MsgBox "Nama Harus Diisi", vbCritical + vbOKOnly, "Peringatan"
        Text1(0).SetFocus
        Exit Function
    End If
    If Text1(3).Text = Empty Then
        CEK_DATA_VALID = False
'        MsgBox "Sumber Info Harus Diisi", vbCritical + vbOKOnly, "Peringatan"
'        Combo1(0).SetFocus
'        Exit Function
    End If
    If Combo1(2).Text = Empty Then
        CEK_DATA_VALID = False
        MsgBox "Title Harus Diisi", vbCritical + vbOKOnly, "Peringatan"
        Combo1(2).SetFocus
        Exit Function
    End If
    If Len(TDBMask1(0).Value) < 3 And Len(TDBMask1(1).Value) < 3 And Len(TDBMask1(2).Value) < 3 And Len(TDBMask1(3).Value) < 3 And Len(TDBMask1(4).Value) < 3 And Len(TDBMask1(5).Value) < 3 And Len(TDBMask1(6).Value) < 3 And Len(TDBMask1(7).Value) < 3 Then
        CEK_DATA_VALID = False
        MsgBox "Minimal Satu Nomor Telpon Harus Diisi", vbCritical + vbOKOnly, "Peringatan"
        SSTab1.Tab = 0
        TDBMask1(0).SetFocus
        Exit Function
    End If
    If Check2(1).Value = 1 Then
        If Combo3(1).Text = Empty Then
            CEK_DATA_VALID = False
            MsgBox "Clossing Reason Harus diIsi", vbCritical + vbOKOnly, "Peringatan"
            SSTab1.Tab = 4
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
            Combo7.SetFocus
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
'        ID_CUST = Text1(1).Text
'        frmtelp.Text11.Text = Text1(0).Text
'        frmtelp.Text10.Text = Combo1(0).Text
'        REMO = True
        frmtelp.Show vbModal
    Case 1
        MsgBox "PEMBERI REFERENSI", vbOKOnly
    Case 2
        If AddMgm = True Then
            V_SAVE = CEK_DATA_VALID_mgm
        Else
            V_SAVE = CEK_DATA_VALID
        End If
        If V_SAVE = False Then
            Exit Sub
        Else
        End If
        If ADD_CUST Then
            ADD_CUST_REM = True
            Combo1(0).Text = "MGM-REF"
            Combo1(1).Text = "MGM-REF"
            Call CEK_ADD_PELANGGAN
        Else
            Combo1(0).Text = "MGM-REF"
            Combo1(1).Text = "MGM-REF"
            Call CEK_UPDATE_PELANGGAN
        End If
    Case 3
        HAK_TeamLeader = False
        Unload Me
End Select
End Sub

Private Function CEK_ADD_PELANGGAN()
Dim M_DATA As New CLS_FRMCUST_CC
Dim m_objrs As ADODB.Recordset
Dim cmdsql As String
Dim M_MSGBOX As Variant

If AddMgm = True Then
' PERIKSA APAKAH INI TERMASUK EVENT INSERT DATA REFERENSI ATAU BUKAN
    Text1(1).Text = "REFI-" & CUSTNOMOR(M_OBJCONN, UCase(Me.Name))
    M_DATA.ADD_CUSTOMER_BARU
    M_DATA.ADD_HISTORY M_OBJCONN, Text1(1).Text, MDIForm1.TDBDate1.Text, Time, MDIForm1.Text1.Text, Combo2(0).Text, "REFERENSI DARI DATA MGM " & Text2.Text & " " & RichTextBox1(3).Text, Combo1(0).Text, Combo6.Text, RichTextBox1(1).Text
        MsgBox "Data Sudah Tersimpan", vbInformation + vbOKOnly, "Sukses"
        ADD_CUST = False
        VIEW_ADD = True
        RichTextBox1(3).Text = Empty
        Text2.Text = Empty
Else

'Text1(1).Text = "CC-I-" & CUSTNOMOR(M_OBJCONN, UCase(Me.Name))

'M_DATA.ADD_CUSTOMER_BARU
'M_DATA.ADD_RequestInbound

'M_DATA.ADD_HISTORY M_OBJCONN, Text1(1).Text, MDIForm1.TDBDate1.Text, Time, MDIForm1.Text1.Text, Combo2(0).Text, Text2.Text & " " & RichTextBox1(3).Text, Combo1(0).Text, Combo6.Text, RichTextBox1(1).Text
'        MsgBox "Data Sudah Tersimpan", vbInformation + vbOKOnly, "Sukses"
'    ADD_CUST = False
'    VIEW_ADD = True
'    RichTextBox1(3).Text = Empty
'    Text2.Text = Empty
End If
'Unload Me
Set M_DATA = Nothing
End Function


Private Sub CEK_UPDATE_PELANGGAN()
Dim M_DATA As New CLS_FRMCUST_CC
Dim m_objrs As ADODB.Recordset
Dim cmdsql As String
Dim M_MSGBOX As Variant
Dim M_CALL As String
Dim M_STATUS As String

M_CALL = "1"
If AddMgm = False Then
' CEK KALAU TRANSAKSI INI ADALAH EVENT UNTUK UPDATE DATA DARI REFERENSI
    SCREENER_APPROV = False
    If HAK_TeamLeader Then
        M_MSGBOX = MsgBox("Apakah Anda Team Leader", vbInformation + vbYesNo, "Konfirmasi")
        If M_MSGBOX = vbYes Then
            FRMPASWORD.Show vbModal
            If HAK_TeamLeader = False Then
                M_DATA.UPDATE_CUSTOMER_BARU M_OBJCONN, KETHSLKERJA, STATUS_FIELD_LAMA, M_CALL, M_STATUS, DOK1
                If Check2(1).Value = 1 Then
                    If RichTextBox1(3).Text <> Empty Or Text2.Text <> Empty Then
                        M_DATA.ADD_HISTORY M_OBJCONN, Text1(1).Text, MDIForm1.TDBDate1.Text, Time, MDIForm1.Text1.Text, "MGM", Text2.Text & " " & RichTextBox1(3).Text, Combo1(0).Text, Combo6.Text, RichTextBox1(1).Text
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
End If

If UCase(MDIForm1.Text2.Text) = "AGENT" Then
 '   Call cek_update_telp_sama
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
        M_DATA.ADD_HISTORY M_OBJCONN, Text1(1).Text, MDIForm1.TDBDate1.Text, Time, MDIForm1.Text1.Text, "MGM", Text2.Text & " " & RichTextBox1(3).Text, Combo1(0).Text, Combo6.Text, RichTextBox1(1).Text
    End If
Else
        M_DATA.ADD_HISTORY M_OBJCONN, Text1(1).Text, MDIForm1.TDBDate1.Text, Time, MDIForm1.Text1.Text, "MGM", Text2.Text & " " & RichTextBox1(3).Text, Combo1(0).Text, Combo6.Text, RichTextBox1(1).Text
End If
MsgBox "Data Sudah Tersimpan", vbInformation + vbOKOnly, "Sukses"
RichTextBox1(3).Text = Empty
Text2.Text = Empty

Set M_DATA = Nothing
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
Dim m_objrs As ADODB.Recordset
Dim M_DATA As New CLS_FRMCUST_CC
    Combo2(0).Text = Empty
    Combo2(1).Text = Empty
    Combo4(0).Text = Empty
    Combo4(1).Text = Empty
    Me.Top = 3000
    Me.Left = 2700
    Call ChangeTab(SSTab1)
    TDBDate1(3).Value = MDIForm1.TDBDate1.Value
    TDBTime1.Value = Format(Time, "hh:mm")
    Option4(0).Value = True
    Call Isi_Combo
    Call HEADER_HISTORY
    Call ISI_COMBO_DATASOURCE
    Call ISI_COMBO_PRODUCT_CLOSING
    If ADD_CUST Then
        Combo1(0).Text = "MGM-REF"
        Combo1(1).Text = "MGM-REF"
        Combo1(0).Enabled = False
        Combo1(1).Enabled = False
        ADD_CUST_REM = True
    Else
        VIEW_ADD = False
        SSCommand1(2).Enabled = False
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

Set m_objrs = Nothing
Set M_DATA = Nothing
SSTab1.Tab = 0
End Sub

Private Sub HEADER_HISTORY()
    ListView1(1).ColumnHeaders.ADD 1, , "Tanggal Jam", 15 * TXT
'    ListView1(1).ColumnHeaders.ADD 2, , "Jam", 8 * TXT
    ListView1(1).ColumnHeaders.ADD 2, , "History", 30 * TXT
    ListView1(1).ColumnHeaders.ADD 3, , "Agent", 10 * TXT
    ListView1(1).ColumnHeaders.ADD 4, , "Code", 10 * TXT
    ListView1(1).ColumnHeaders.ADD 5, , "Complaint Note", 20 * TXT
End Sub

Private Sub VIEW_DATA_CUST()
Dim m_objrs As ADODB.Recordset
Dim m_objrs1 As ADODB.Recordset
Dim M_DATA As New CLS_FRMCUST_CC
Dim LISTITEM As LISTITEM
Dim M_CAT As String
Dim M_QUALIFIED As String
Dim m_balance As String
If VIEW_AVAIL_AWAL Then
    Set m_objrs = M_DATA.QUERY_CUST(M_OBJCONN, "CUSTID = '" + VIEWCUSTAVAIL_AGENT.ListView1.SelectedItem.SubItems(1) + "'")
End If
If VIEW_OK Then
    Set m_objrs = M_DATA.QUERY_CUST(M_OBJCONN, "CUSTID = '" + VIEWCUSTAVAIL.ListView1.SelectedItem.SubItems(1) + "'")
End If
If SCREENER_AWAL = True Then
    Set m_objrs = M_DATA.QUERY_CUST(M_OBJCONN, "CUSTID = '" + FRM_PRESCREEN_AWAL.ListView1.SelectedItem.Text + "'")
End If
If SCREENER = True Then
    Set m_objrs = M_DATA.QUERY_CUST(M_OBJCONN, "CUSTID = '" + FRM_PRESCREEN.ListView1.SelectedItem.SubItems(1) + "'")
End If
If VIEW_ADD = True Then
    Exit Sub
End If
    If m_objrs.RecordCount <> 0 Then
        Text1(1).Text = IIf(IsNull(m_objrs("CUSTID")), "", m_objrs("CUSTID"))
        ID_CUST = IIf(IsNull(m_objrs("CUSTID")), "", m_objrs("CUSTID"))
        Text7.Text = IIf(IsNull(m_objrs("NamaAgent")), "", m_objrs("NamaAgent"))
        Text8.Text = IIf(IsNull(m_objrs("RecSourceRef")), "", m_objrs("RecSourceRef"))
        
        Text1(0).Text = IIf(IsNull(m_objrs("NAME")), "", m_objrs("NAME"))
        Text4.Text = IIf(IsNull(m_objrs("CustIdRef")), "", m_objrs("CustIdRef"))
        Text9.Text = IIf(IsNull(m_objrs("NAMAREF")), "", m_objrs("NAMAREF"))
        Combo1(2).Text = IIf(IsNull(m_objrs("TITLE")), "", m_objrs("TITLE"))
        Call Combo1_LostFocus(2)
        
        TDBDate1(0).Value = IIf(IsNull(m_objrs("BIRTHD")), "", Format(m_objrs("BIRTHD"), "dd-mmm-yyyy"))
        Call TDBDate1_Click(0)
        RichTextBox1(0).Text = IIf(IsNull(m_objrs("ADDRNOW")), "", m_objrs("ADDRNOW"))
        Text1(6).Text = IIf(IsNull(m_objrs("ZIPNOW")), "", m_objrs("ZIPNOW"))
        Text1(7).Text = IIf(IsNull(m_objrs("CITYNOW")), "", m_objrs("CITYNOW"))
        TDBMask1(0).Value = ILANGIN_AREA(IIf(IsNull(m_objrs("HOMENO")), "", m_objrs("HOMENO")))
        TDBMask1(1).Value = IIf(IsNull(m_objrs("HOMENO2")), "", m_objrs("HOMENO2"))
        TDBMask1(2).Value = IIf(IsNull(m_objrs("MOBILENO")), "", m_objrs("MOBILENO"))
        TDBMask1(3).Value = IIf(IsNull(m_objrs("MOBILENO2")), "", m_objrs("MOBILENO2"))
        TDBMask1(4).Value = ILANGIN_AREA(IIf(IsNull(m_objrs("OFFICENO")), "", m_objrs("OFFICENO")))
        Text1(21).Text = IIf(IsNull(m_objrs("EXTOFFICE")), "", m_objrs("EXTOFFICE"))
        TDBMask1(5).Value = IIf(IsNull(m_objrs("OFFICENO2")), "", m_objrs("OFFICENO2"))
        Text1(32).Text = IIf(IsNull(m_objrs("EXTOFFICE2")), "", m_objrs("EXTOFFICE2"))
        TDBMask1(6).Value = IIf(IsNull(m_objrs("FAXNO")), "", m_objrs("FAXNO"))
        TDBMask1(7).Value = IIf(IsNull(m_objrs("FAXNO2")), "", m_objrs("FAXNO2"))
        M_CAT = IIf(IsNull(m_objrs("CAT")), "0", m_objrs("CAT"))
        Combo5.Text = IIf(IsNull(m_objrs("PRIOR")), "", m_objrs("PRIOR"))
        If M_CAT = "0" Then
            Option4(0).Value = True
            Text1(20).Text = IIf(IsNull(m_objrs("NAMAPT")), "", m_objrs("NAMAPT"))
            RichTextBox1(2).Text = IIf(IsNull(m_objrs("ADDRPT")), "", m_objrs("ADDRPT"))
        Else
            Option4(1).Value = True
            Text1(34).Text = IIf(IsNull(m_objrs("JENISUSAHA")), "", m_objrs("JENISUSAHA"))
            Text1(18).Text = IIf(IsNull(m_objrs("NAMAPT")), "", m_objrs("NAMAPT"))
            RichTextBox1(4).Text = IIf(IsNull(m_objrs("ADDRPT")), "", m_objrs("ADDRPT"))
        End If
        Combo1(0).Text = IIf(IsNull(m_objrs("RECSOURCE")), "", m_objrs("RECSOURCE"))
        Call Combo1_Click(0)
        Text3.Text = IIf(IsNull(m_objrs("OTHERS")), "", m_objrs("OTHERS"))
        TDBMask1(8).Value = IIf(IsNull(m_objrs("AHOMENO")), "", m_objrs("AHOMENO"))
        TDBMask1(9).Value = IIf(IsNull(m_objrs("AHOMENO2")), "", m_objrs("AHOMENO2"))
        TDBMask1(10).Value = IIf(IsNull(m_objrs("AOFFICENO")), "", m_objrs("AOFFICENO"))
        TDBMask1(11).Value = IIf(IsNull(m_objrs("AOFFICENO2")), "", m_objrs("AOFFICENO2"))
        TDBMask1(12).Value = IIf(IsNull(m_objrs("AFAXNO")), "", m_objrs("AFAXNO"))
        TDBMask1(13).Value = IIf(IsNull(m_objrs("AFAXNO2")), "", m_objrs("AFAXNO2"))
    
        Select Case m_objrs!RECSTATUS
        Case "N"
            Check2(1).Value = 1
            Combo3(0).Text = IIf(IsNull(m_objrs("KETHSLKERJA")), "", m_objrs("KETHSLKERJA"))
            Call Combo3_LostFocus(0)
        Case "C"
            Check2(0).Value = 1
            Combo2(0).Text = IIf(IsNull(m_objrs("KETHSLKERJA")), "", m_objrs("KETHSLKERJA"))
            Call Combo2_LostFocus(0)
        Case "I"
            Check2(2).Value = 1
            TDBDate1(1).Value = IIf(IsNull(m_objrs("TGLSTATUS")), "", Format(m_objrs("TGLSTATUS"), "dd-mmm-yyyy"))
        End Select
        Combo4(0).Text = IIf(IsNull(m_objrs("KD_CLS")), "", m_objrs("KD_CLS"))
        Call Combo4_LostFocus(0)
    End If
Set m_objrs = Nothing
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
        LISTITEM.SubItems(3) = IIf(IsNull(m_objrs1("KdComplaint")), "", m_objrs1("KdComplaint"))
        LISTITEM.SubItems(4) = IIf(IsNull(m_objrs1("RemarkComplaint")), "", m_objrs1("RemarkComplaint"))
m_objrs1.MoveNext
Wend
Set m_objrs1 = Nothing
Set M_DATA = Nothing

End Sub


Private Sub ISI_COMBO_DATASOURCE()
Dim m_objrs As ADODB.Recordset
Dim M_DATA As New CLS_FRMCUST_CC
    Set m_objrs = M_DATA.QUERY_COMBO_DATASOURCE_ISI(M_OBJCONN, "")
    While Not m_objrs.EOF
        Combo1(0).AddItem m_objrs("KODEDS")
        Combo1(0).DataField = m_objrs("KODEDS")
        Combo1(1).AddItem IIf(IsNull(m_objrs("KETERANGAN")), "", m_objrs("KETERANGAN"))
        Combo1(1).DataField = m_objrs("KETERANGAN")
        m_objrs.MoveNext
    Wend
Set m_objrs = Nothing
Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open "Select * from ComplaintCode", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not m_objrs.EOF
    Combo6.AddItem m_objrs("KdComplaint")
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing

Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open "Select * from StsNextAct", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not m_objrs.EOF
    Combo7.AddItem m_objrs("NmStsNextAct")
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing

Set M_DATA = Nothing
End Sub

Private Sub ISI_COMBO_PRODUCT_CLOSING()
Dim m_objrs As ADODB.Recordset
Dim M_DATA As New CLS_FRMCUST_CC
    Set m_objrs = M_DATA.QUERY_COMBO_CLOSSING(M_OBJCONN, " UPPER(JENIS)<>'MGM'")
    While Not m_objrs.EOF
        Combo3(0).AddItem m_objrs("KDCLS")
        Combo3(0).DataField = m_objrs("KDCLS")
        Combo3(1).AddItem m_objrs("KETCLS")
        Combo3(1).DataField = m_objrs("KETCLS")
        m_objrs.MoveNext
    Wend
    Set m_objrs = Nothing
Set M_DATA = Nothing

Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open "Select * from Aiu_NoProdPresented WHERE UPPER(Jenis) <> 'MGM' ", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not m_objrs.EOF
    Combo2(0).AddItem IIf(IsNull(m_objrs!KdNoProdPresented), "", m_objrs!KdNoProdPresented)
    Combo2(1).AddItem IIf(IsNull(m_objrs!NmNoProdPresented), "", m_objrs!NmNoProdPresented)
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing

Set m_objrs = New ADODB.Recordset
m_objrs.CursorLocation = adUseClient
m_objrs.Open "Select * from Aiu_DataQuality", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
While Not m_objrs.EOF
    Combo4(0).AddItem IIf(IsNull(m_objrs!KdDataQuality), "", m_objrs!KdDataQuality)
    Combo4(1).AddItem IIf(IsNull(m_objrs!NmDataQuality), "", m_objrs!NmDataQuality)
    m_objrs.MoveNext
Wend
Set m_objrs = Nothing

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
Dim m_objrs As ADODB.Recordset
Select Case Index
    Case 0
        Text1(Index).Text = UCase(Text1(Index).Text)
    Case 4
        If Len(Text1(Index).Text) > 2 Then
            Set m_objrs = New ADODB.Recordset
            m_objrs.CursorLocation = adUseClient
            m_objrs.Open "SELECT NOLAP FROM CC_CUSTTBL WHERE NOLAP = '" + Text1(Index).Text + "' AND  CUSTID <> '" + Text1(1).Text + "'", M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            If m_objrs.RecordCount <> 0 Then
                MsgBox "Nomor Yang Anda Masukan Telah Ada", vbInformation, "TeleGrandi"
                Text1(Index).Text = Empty
            End If
        End If
End Select
Set m_objrs = Nothing
End Sub

Private Sub Text2_LostFocus()
    Text2.Text = UCase(Text2.Text)
End Sub


Private Sub cek_update_telp_sama()
Dim cmdsql As String
Dim M_DATA As New CLS_FRMCUST_CC
Dim m_objrs As ADODB.Recordset
Dim M_MSGBOX As Variant
        If TDBMask1(0).ReadOnly = False Then
            If Len(TDBMask1(0).Value) > 4 Then
                cmdsql = "(HOMENO = '" + TDBMask1(0).Value + "'"
                cmdsql = cmdsql + " OR HOMENO2 = '" + TDBMask1(0).Value + "'"
                cmdsql = cmdsql + " OR MOBILENO = '" + TDBMask1(0).Value + "'"
                cmdsql = cmdsql + " OR MOBILENO2 = '" + TDBMask1(0).Value + "'"
                cmdsql = cmdsql + " OR FAXNO = '" + TDBMask1(0).Value + "'"
                cmdsql = cmdsql + " OR FAXNO2 = '" + TDBMask1(0).Value + "'"
                cmdsql = cmdsql + " OR OFFICENO = '" + TDBMask1(0).Value + "'"
                cmdsql = cmdsql + " OR OFFICENO2 = '" + TDBMask1(0).Value + "'"
            End If
        End If
        If TDBMask1(1).ReadOnly = False Then
            If Len(TDBMask1(1).Value) > 4 Then
                If cmdsql = Empty Then
                    cmdsql = cmdsql + " (HOMENO = '" + TDBMask1(1).Value + "'"
                Else
                    cmdsql = cmdsql + " OR HOMENO = '" + TDBMask1(1).Value + "'"
                End If
                    cmdsql = cmdsql + " OR HOMENO2 = '" + TDBMask1(1).Value + "'"
                    cmdsql = cmdsql + " OR MOBILENO = '" + TDBMask1(1).Value + "'"
                    cmdsql = cmdsql + " OR MOBILENO2 = '" + TDBMask1(1).Value + "'"
                    cmdsql = cmdsql + " OR FAXNO = '" + TDBMask1(1).Value + "'"
                    cmdsql = cmdsql + " OR FAXNO2 = '" + TDBMask1(1).Value + "'"
                    cmdsql = cmdsql + " OR OFFICENO = '" + TDBMask1(1).Value + "'"
                    cmdsql = cmdsql + " OR OFFICENO2 = '" + TDBMask1(1).Value + "'"
            End If
        End If
        If TDBMask1(2).ReadOnly = False Then
            If Len(TDBMask1(2).Value) > 4 Then
                If cmdsql = Empty Then
                    cmdsql = cmdsql + " (HOMENO = '" + TDBMask1(2).Value + "'"
                Else
                    cmdsql = cmdsql + " OR HOMENO = '" + TDBMask1(2).Value + "'"
                End If
                    cmdsql = cmdsql + " OR HOMENO2 = '" + TDBMask1(2).Value + "'"
                    cmdsql = cmdsql + " OR MOBILENO = '" + TDBMask1(2).Value + "'"
                    cmdsql = cmdsql + " OR MOBILENO2 = '" + TDBMask1(2).Value + "'"
                    cmdsql = cmdsql + " OR FAXNO = '" + TDBMask1(2).Value + "'"
                    cmdsql = cmdsql + " OR FAXNO2 = '" + TDBMask1(2).Value + "'"
                    cmdsql = cmdsql + " OR OFFICENO = '" + TDBMask1(2).Value + "'"
                    cmdsql = cmdsql + " OR OFFICENO2 = '" + TDBMask1(2).Value + "'"
            End If
        End If
        If TDBMask1(3).ReadOnly = False Then
            If Len(TDBMask1(3).Value) > 4 Then
                If cmdsql = Empty Then
                    cmdsql = cmdsql + " (HOMENO = '" + TDBMask1(3).Value + "'"
                Else
                    cmdsql = cmdsql + " OR HOMENO = '" + TDBMask1(3).Value + "'"
                End If
                    cmdsql = cmdsql + " OR HOMENO2 = '" + TDBMask1(3).Value + "'"
                    cmdsql = cmdsql + " OR MOBILENO = '" + TDBMask1(3).Value + "'"
                    cmdsql = cmdsql + " OR MOBILENO2 = '" + TDBMask1(3).Value + "'"
                    cmdsql = cmdsql + " OR FAXNO = '" + TDBMask1(3).Value + "'"
                    cmdsql = cmdsql + " OR FAXNO2 = '" + TDBMask1(3).Value + "'"
                    cmdsql = cmdsql + " OR OFFICENO = '" + TDBMask1(3).Value + "'"
                    cmdsql = cmdsql + " OR OFFICENO2 = '" + TDBMask1(3).Value + "'"
            End If
        End If
        If TDBMask1(4).ReadOnly = False Then
            If Len(TDBMask1(4).Value) > 4 Then
                If cmdsql = Empty Then
                    cmdsql = cmdsql + " (HOMENO = '" + TDBMask1(4).Value + "'"
                Else
                    cmdsql = cmdsql + " OR HOMENO = '" + TDBMask1(4).Value + "'"
                End If
                    cmdsql = cmdsql + " OR HOMENO = '" + TDBMask1(4).Value + "'"
                    cmdsql = cmdsql + " OR HOMENO2 = '" + TDBMask1(4).Value + "'"
                    cmdsql = cmdsql + " OR MOBILENO = '" + TDBMask1(4).Value + "'"
                    cmdsql = cmdsql + " OR MOBILENO2 = '" + TDBMask1(4).Value + "'"
                    cmdsql = cmdsql + " OR FAXNO = '" + TDBMask1(4).Value + "'"
                    cmdsql = cmdsql + " OR FAXNO2 = '" + TDBMask1(4).Value + "'"
                    cmdsql = cmdsql + " OR OFFICENO = '" + TDBMask1(4).Value + "'"
                    cmdsql = cmdsql + " OR OFFICENO2 = '" + TDBMask1(4).Value + "'"
            End If
        End If
        If TDBMask1(5).ReadOnly = False Then
            If Len(TDBMask1(5).Value) > 4 Then
                If cmdsql = Empty Then
                    cmdsql = cmdsql + " (HOMENO = '" + TDBMask1(5).Value + "'"
                Else
                    cmdsql = cmdsql + " OR HOMENO = '" + TDBMask1(5).Value + "'"
                End If
                    cmdsql = cmdsql + " OR HOMENO2 = '" + TDBMask1(5).Value + "'"
                    cmdsql = cmdsql + " OR MOBILENO = '" + TDBMask1(5).Value + "'"
                    cmdsql = cmdsql + " OR MOBILENO2 = '" + TDBMask1(5).Value + "'"
                    cmdsql = cmdsql + " OR FAXNO = '" + TDBMask1(5).Value + "'"
                    cmdsql = cmdsql + " OR FAXNO2 = '" + TDBMask1(5).Value + "'"
                    cmdsql = cmdsql + " OR OFFICENO = '" + TDBMask1(5).Value + "'"
                    cmdsql = cmdsql + " OR OFFICENO2 = '" + TDBMask1(5).Value + "'"
        End If
        End If
        If TDBMask1(6).ReadOnly = False Then
            If Len(TDBMask1(6).Value) > 4 Then
                If cmdsql = Empty Then
                    cmdsql = cmdsql + " (HOMENO = '" + TDBMask1(6).Value + "'"
                Else
                    cmdsql = cmdsql + " OR HOMENO = '" + TDBMask1(6).Value + "'"
                End If
                    cmdsql = cmdsql + " OR HOMENO2 = '" + TDBMask1(6).Value + "'"
                    cmdsql = cmdsql + " OR MOBILENO = '" + TDBMask1(6).Value + "'"
                    cmdsql = cmdsql + " OR MOBILENO2 = '" + TDBMask1(6).Value + "'"
                    cmdsql = cmdsql + " OR FAXNO = '" + TDBMask1(6).Value + "'"
                    cmdsql = cmdsql + " OR FAXNO2 = '" + TDBMask1(6).Value + "'"
                    cmdsql = cmdsql + " OR OFFICENO = '" + TDBMask1(6).Value + "'"
                    cmdsql = cmdsql + " OR OFFICENO2 = '" + TDBMask1(6).Value + "'"
        End If
        End If
        If TDBMask1(7).ReadOnly = False Then
            If Len(TDBMask1(7).Value) > 4 Then
                If cmdsql = Empty Then
                    cmdsql = cmdsql + " (HOMENO = '" + TDBMask1(7).Value + "'"
                Else
                    cmdsql = cmdsql + " or HOMENO = '" + TDBMask1(7).Value + "'"
                End If
                    cmdsql = cmdsql + " OR HOMENO2 = '" + TDBMask1(7).Value + "'"
                    cmdsql = cmdsql + " OR MOBILENO = '" + TDBMask1(7).Value + "'"
                    cmdsql = cmdsql + " OR MOBILENO2 = '" + TDBMask1(7).Value + "'"
                    cmdsql = cmdsql + " OR FAXNO = '" + TDBMask1(7).Value + "'"
                    cmdsql = cmdsql + " OR FAXNO2 = '" + TDBMask1(7).Value + "'"
                    cmdsql = cmdsql + " OR OFFICENO = '" + TDBMask1(7).Value + "'"
                    cmdsql = cmdsql + " OR OFFICENO2 = '" + TDBMask1(7).Value + "'"
        End If
        End If
        If Len(cmdsql) <> 0 Then
            cmdsql = cmdsql + ") AND CUSTID <> '" + Text1(1).Text + "'"
            Set m_objrs = M_DATA.QUERY_CEK_ADDCUST(M_OBJCONN, cmdsql)
            If m_objrs.RecordCount >= 3 Or m_objrs.RecordCount < 1 Then
                HAK_TeamLeader = False
                Set m_objrs = Nothing
                Exit Sub
            Else
                HAK_TeamLeader = True
                update_TELP_SAMA = True
                MsgBox "Update Gagal... No Telepon Ada Yang Sama... Hubungi Team Leader Untuk Menyimpan Data ", vbCritical + vbOKOnly, "Peringatan"
                Set m_objrs = Nothing
                FRM_DATASAMA_CC.Show vbModal
            Exit Sub
            End If
        End If
Set m_objrs = Nothing
End Sub


Private Function ILANGIN_AREA(TELP As String) As String
    ILANGIN_AREA = Replace(TELP, "021", "")
End Function

