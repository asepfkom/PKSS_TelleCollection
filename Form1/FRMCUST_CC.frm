VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FRMCUST_CC 
   Caption         =   "Referall Data"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   10065
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   Icon            =   "FRMCUST_CC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtReasonClosing 
      Height          =   375
      Left            =   1725
      TabIndex        =   134
      Top             =   45
      Visible         =   0   'False
      Width           =   5550
   End
   Begin VB.TextBox TxtMgmName 
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
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   91
      Top             =   30
      Width           =   4035
   End
   Begin Threed.SSCommand SSCommand1 
      Cancel          =   -1  'True
      Height          =   420
      Index           =   3
      Left            =   8730
      TabIndex        =   2
      Top             =   60
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
      Picture         =   "FRMCUST_CC.frx":0442
      Caption         =   "&Exit"
      AutoSize        =   2
      Alignment       =   4
      ButtonStyle     =   2
      PictureAlignment=   1
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   420
      Index           =   2
      Left            =   7560
      TabIndex        =   1
      Top             =   120
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
      Picture         =   "FRMCUST_CC.frx":059C
      Caption         =   "&Save"
      AutoSize        =   2
      Alignment       =   4
      ButtonStyle     =   2
      PictureAlignment=   1
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   420
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   0
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
      Picture         =   "FRMCUST_CC.frx":08BE
      Caption         =   "&Call"
      AutoSize        =   2
      Alignment       =   4
      ButtonStyle     =   2
      PictureAlignment=   1
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   150
      TabIndex        =   49
      Top             =   435
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
         Left            =   6960
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   120
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
         Left            =   960
         MaxLength       =   200
         TabIndex        =   3
         Top             =   120
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
         TabIndex        =   51
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
         TabIndex        =   50
         Top             =   195
         Width           =   960
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4590
      Left            =   120
      TabIndex        =   35
      Top             =   960
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   8096
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
      TabPicture(0)   =   "FRMCUST_CC.frx":1132
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "History"
      TabPicture(1)   =   "FRMCUST_CC.frx":114E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame9"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Results"
      TabPicture(2)   =   "FRMCUST_CC.frx":116A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame10"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Submission"
      TabPicture(3)   =   "FRMCUST_CC.frx":1186
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "TDBDate1(2)"
      Tab(3).Control(1)=   "frmSubmit"
      Tab(3).Control(2)=   "Check2(2)"
      Tab(3).Control(3)=   "Command2"
      Tab(3).Control(4)=   "FrmCek"
      Tab(3).ControlCount=   5
      Begin VB.Frame FrmCek 
         Caption         =   "Check"
         Height          =   675
         Left            =   -74475
         TabIndex        =   129
         Top             =   2430
         Width           =   1080
         Begin VB.CheckBox ChkCek 
            Caption         =   "Accept"
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   131
            Top             =   180
            Width           =   915
         End
         Begin VB.CheckBox ChkCek 
            Caption         =   "Return"
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   130
            Top             =   405
            Width           =   915
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Update Tgl"
         Height          =   300
         Left            =   -71820
         TabIndex        =   127
         Top             =   2595
         Width           =   915
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Submission"
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
         Height          =   285
         Index           =   2
         Left            =   -74160
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   98
         Top             =   540
         Width           =   1260
      End
      Begin VB.Frame frmSubmit 
         Height          =   1620
         Left            =   -74280
         TabIndex        =   99
         Top             =   600
         Width           =   8295
         Begin VB.CheckBox CekBt 
            Caption         =   "Yes"
            ForeColor       =   &H00000080&
            Height          =   270
            Left            =   5520
            TabIndex        =   110
            Top             =   795
            Width           =   645
         End
         Begin VB.CheckBox CekCarded 
            Caption         =   "Carded"
            ForeColor       =   &H00000080&
            Height          =   270
            Index           =   0
            Left            =   5520
            TabIndex        =   109
            Top             =   555
            Width           =   840
         End
         Begin VB.CheckBox CekCarded 
            Caption         =   "Uncarded"
            ForeColor       =   &H00000080&
            Height          =   270
            Index           =   1
            Left            =   6405
            TabIndex        =   108
            Top             =   570
            Width           =   1035
         End
         Begin VB.CheckBox cekSegmented 
            Caption         =   "Gold"
            ForeColor       =   &H00000080&
            Height          =   270
            Index           =   0
            Left            =   5520
            TabIndex        =   107
            Top             =   1050
            Width           =   720
         End
         Begin VB.CheckBox cekSegmented 
            Caption         =   "Clasik"
            ForeColor       =   &H00000080&
            Height          =   270
            Index           =   1
            Left            =   6390
            TabIndex        =   106
            Top             =   1035
            Width           =   900
         End
         Begin VB.TextBox TxtDob 
            Height          =   315
            Index           =   0
            Left            =   1110
            MaxLength       =   2
            TabIndex        =   105
            Top             =   540
            Width           =   420
         End
         Begin VB.TextBox TxtDob 
            Height          =   315
            Index           =   1
            Left            =   1560
            MaxLength       =   2
            TabIndex        =   104
            Top             =   540
            Width           =   420
         End
         Begin VB.TextBox TxtDob 
            Height          =   315
            Index           =   2
            Left            =   2010
            MaxLength       =   2
            TabIndex        =   103
            Top             =   540
            Width           =   420
         End
         Begin VB.TextBox TxtSubmission 
            Height          =   315
            Index           =   0
            Left            =   1110
            TabIndex        =   102
            Top             =   225
            Width           =   3180
         End
         Begin VB.TextBox TxtSubmission 
            Height          =   315
            Index           =   1
            Left            =   1110
            TabIndex        =   101
            Top             =   1215
            Visible         =   0   'False
            Width           =   3180
         End
         Begin VB.CheckBox CekCreditShield 
            Caption         =   "Yes"
            ForeColor       =   &H00000080&
            Height          =   270
            Left            =   7335
            TabIndex        =   100
            Top             =   810
            Width           =   675
         End
         Begin TDBMask6Ctl.TDBMask TDBMaskHomeSub 
            Height          =   360
            Index           =   1
            Left            =   1725
            TabIndex        =   111
            Top             =   855
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
            _ExtentY        =   635
            Caption         =   "FRMCUST_CC.frx":11A2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "FRMCUST_CC.frx":120E
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
         Begin TDBMask6Ctl.TDBMask TDBMaskHomeSub 
            Height          =   360
            Index           =   0
            Left            =   1110
            TabIndex        =   112
            Top             =   855
            Width           =   585
            _Version        =   65536
            _ExtentX        =   1032
            _ExtentY        =   635
            Caption         =   "FRMCUST_CC.frx":1250
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "FRMCUST_CC.frx":12BC
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
         Begin TDBMask6Ctl.TDBMask TDBMaskOfficeSub 
            Height          =   360
            Index           =   1
            Left            =   6105
            TabIndex        =   113
            Top             =   210
            Visible         =   0   'False
            Width           =   1545
            _Version        =   65536
            _ExtentX        =   2725
            _ExtentY        =   635
            Caption         =   "FRMCUST_CC.frx":12FE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "FRMCUST_CC.frx":136A
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
         Begin TDBMask6Ctl.TDBMask TDBMaskOfficeSub 
            Height          =   360
            Index           =   0
            Left            =   5490
            TabIndex        =   114
            Top             =   210
            Visible         =   0   'False
            Width           =   585
            _Version        =   65536
            _ExtentX        =   1032
            _ExtentY        =   635
            Caption         =   "FRMCUST_CC.frx":13AC
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "FRMCUST_CC.frx":1418
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
         Begin TDBMask6Ctl.TDBMask TDBMaskHpSub 
            Height          =   360
            Left            =   3120
            TabIndex        =   115
            Top             =   2040
            Visible         =   0   'False
            Width           =   1965
            _Version        =   65536
            _ExtentX        =   3466
            _ExtentY        =   635
            Caption         =   "FRMCUST_CC.frx":145A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Keys            =   "FRMCUST_CC.frx":14C6
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
            Index           =   0
            Left            =   45
            TabIndex        =   126
            Top             =   585
            Width           =   1065
         End
         Begin VB.Label Label6 
            Caption         =   "Home Phone :"
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
            Left            =   105
            TabIndex        =   125
            Top             =   915
            Width           =   1050
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Carded :"
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
            Index           =   2
            Left            =   4740
            TabIndex        =   124
            Top             =   555
            Width           =   660
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "BT :"
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
            Index           =   3
            Left            =   4740
            TabIndex        =   123
            Top             =   810
            Width           =   660
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Jenis Kartu :"
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
            Index           =   4
            Left            =   4410
            TabIndex        =   122
            Top             =   1080
            Width           =   990
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "dd/mm/yy"
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
            Index           =   7
            Left            =   2490
            TabIndex        =   121
            Top             =   585
            Width           =   750
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Name :"
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
            Index           =   8
            Left            =   375
            TabIndex        =   120
            Top             =   285
            Width           =   735
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Id # :"
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
            Index           =   9
            Left            =   210
            TabIndex        =   119
            Top             =   1245
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.Label Label6 
            Caption         =   "Office Phone :"
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
            Index           =   4
            Left            =   4455
            TabIndex        =   118
            Top             =   270
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.Label Label6 
            Caption         =   "Hp. "
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
            Index           =   5
            Left            =   3330
            TabIndex        =   117
            Top             =   1215
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Credit Shield :"
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
            Index           =   10
            Left            =   6315
            TabIndex        =   116
            Top             =   825
            Width           =   1005
         End
      End
      Begin VB.Frame Frame10 
         Height          =   4095
         Left            =   240
         TabIndex        =   73
         Top             =   360
         Width           =   9720
         Begin VB.ComboBox cmbPayNo 
            Height          =   315
            Left            =   1080
            TabIndex        =   160
            Top             =   960
            Width           =   2295
         End
         Begin VB.Frame Frame8 
            Height          =   615
            Left            =   3480
            TabIndex        =   152
            Top             =   240
            Width           =   3255
            Begin VB.ComboBox cmbContacted 
               Height          =   315
               Index           =   4
               Left            =   960
               TabIndex        =   154
               Top             =   240
               Width           =   2175
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
               Left            =   0
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   153
               Top             =   0
               Width           =   1500
            End
            Begin VB.Label Label4 
               Caption         =   "Desc :"
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
               Index           =   11
               Left            =   360
               TabIndex        =   155
               Top             =   240
               Width           =   555
            End
         End
         Begin VB.Frame Frame5 
            Height          =   615
            Left            =   120
            TabIndex        =   149
            Top             =   240
            Width           =   3255
            Begin VB.ComboBox cmbContacted 
               Height          =   315
               Index           =   2
               Left            =   1080
               TabIndex        =   157
               Top             =   240
               Width           =   2055
            End
            Begin VB.ComboBox cmbContacted 
               Height          =   315
               Index           =   3
               Left            =   1080
               TabIndex        =   156
               Top             =   240
               Width           =   1095
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
               Left            =   0
               MouseIcon       =   "FRMCUST_CC.frx":1508
               MousePointer    =   14  'Arrow and Question
               TabIndex        =   150
               Top             =   0
               Width           =   1260
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "Desc :"
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
               Index           =   7
               Left            =   240
               TabIndex        =   151
               Top             =   240
               Width           =   555
            End
         End
         Begin VB.Frame Frame6 
            Height          =   1335
            Left            =   120
            TabIndex        =   143
            Top             =   1320
            Width           =   4455
            Begin VB.ComboBox cmbDiscount 
               Height          =   315
               Left            =   1080
               TabIndex        =   158
               Top             =   240
               Width           =   855
            End
            Begin VB.TextBox txtPayment 
               Height          =   285
               Left            =   1080
               TabIndex        =   144
               Top             =   580
               Width           =   3015
            End
            Begin TDBDate6Ctl.TDBDate TDBDate1 
               Height          =   315
               Index           =   4
               Left            =   1080
               TabIndex        =   145
               Top             =   870
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               Calendar        =   "FRMCUST_CC.frx":194A
               Caption         =   "FRMCUST_CC.frx":1A62
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FRMCUST_CC.frx":1ACE
               Keys            =   "FRMCUST_CC.frx":1AEC
               Spin            =   "FRMCUST_CC.frx":1B4A
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
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "Tgl Bayar :"
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
               Index           =   10
               Left            =   120
               TabIndex        =   148
               Top             =   900
               Width           =   795
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "Payment :"
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
               Index           =   9
               Left            =   120
               TabIndex        =   147
               Top             =   600
               Width           =   795
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "Discount :"
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
               Index           =   8
               Left            =   120
               TabIndex        =   146
               Top             =   240
               Width           =   795
            End
         End
         Begin VB.ComboBox CmbDisagree 
            Height          =   315
            Index           =   0
            Left            =   8520
            TabIndex        =   142
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.ComboBox cmbContacted 
            Height          =   315
            Index           =   0
            Left            =   6480
            TabIndex        =   141
            Text            =   "Combo2"
            Top             =   3720
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.ComboBox cmbContacted 
            Height          =   315
            Index           =   1
            Left            =   5730
            Sorted          =   -1  'True
            TabIndex        =   137
            Top             =   3270
            Visible         =   0   'False
            Width           =   3450
         End
         Begin VB.TextBox TxtSendPOD 
            Height          =   300
            Left            =   5790
            TabIndex        =   136
            Top             =   3570
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.TextBox TxtPuPOD 
            Height          =   300
            Left            =   5775
            TabIndex        =   135
            Top             =   3975
            Visible         =   0   'False
            Width           =   2070
         End
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
            Left            =   11280
            Sorted          =   -1  'True
            TabIndex        =   95
            Top             =   2160
            Visible         =   0   'False
            Width           =   1380
         End
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
            Left            =   8280
            Sorted          =   -1  'True
            TabIndex        =   93
            Top             =   2160
            Width           =   3225
         End
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
            Left            =   7200
            MousePointer    =   14  'Arrow and Question
            TabIndex        =   92
            Top             =   240
            Width           =   1110
         End
         Begin VB.Frame Frame11 
            Caption         =   "Frame11"
            Height          =   1245
            Left            =   120
            TabIndex        =   81
            Top             =   2760
            Width           =   5145
            Begin RichTextLib.RichTextBox RichTextBox1 
               Height          =   855
               Index           =   3
               Left            =   120
               TabIndex        =   41
               Top             =   240
               Width           =   4875
               _ExtentX        =   8599
               _ExtentY        =   1508
               _Version        =   393217
               TextRTF         =   $"FRMCUST_CC.frx":1B72
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
               Left            =   120
               TabIndex        =   82
               Top             =   -15
               Width           =   825
            End
         End
         Begin VB.Frame Frame22 
            Height          =   1365
            Left            =   4800
            TabIndex        =   77
            Top             =   1320
            Width           =   3645
            Begin VB.ComboBox Combo7 
               Height          =   315
               Left            =   1065
               TabIndex        =   37
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
               TabIndex        =   78
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
               TabIndex        =   40
               Top             =   825
               Width           =   1350
            End
            Begin TDBTime6Ctl.TDBTime TDBTime1 
               Height          =   315
               Left            =   2535
               TabIndex        =   39
               Top             =   495
               Width           =   900
               _Version        =   65536
               _ExtentX        =   1587
               _ExtentY        =   556
               Caption         =   "FRMCUST_CC.frx":1BED
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC.frx":1C59
               Spin            =   "FRMCUST_CC.frx":1CA9
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
               Left            =   1080
               TabIndex        =   38
               Top             =   495
               Width           =   1395
               _Version        =   65536
               _ExtentX        =   2461
               _ExtentY        =   556
               Calendar        =   "FRMCUST_CC.frx":1CD1
               Caption         =   "FRMCUST_CC.frx":1DE9
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FRMCUST_CC.frx":1E55
               Keys            =   "FRMCUST_CC.frx":1E73
               Spin            =   "FRMCUST_CC.frx":1ED1
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
               TabIndex        =   80
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
               TabIndex        =   79
               Top             =   570
               Width           =   825
            End
         End
         Begin TDBDate6Ctl.TDBDate TDBDate1 
            Height          =   315
            Index           =   1
            Left            =   4695
            TabIndex        =   89
            Top             =   3615
            Visible         =   0   'False
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   556
            Calendar        =   "FRMCUST_CC.frx":1EF9
            Caption         =   "FRMCUST_CC.frx":2011
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FRMCUST_CC.frx":207D
            Keys            =   "FRMCUST_CC.frx":209B
            Spin            =   "FRMCUST_CC.frx":20F9
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
         Begin VB.Frame Frame13 
            Caption         =   "Frame11"
            Height          =   1080
            Left            =   8880
            TabIndex        =   83
            Top             =   1200
            Visible         =   0   'False
            Width           =   8490
            Begin VB.ComboBox Combo6 
               Height          =   315
               Left            =   630
               TabIndex        =   84
               Top             =   210
               Width           =   885
            End
            Begin RichTextLib.RichTextBox RichTextBox1 
               Height          =   795
               Index           =   1
               Left            =   2145
               TabIndex        =   85
               Top             =   195
               Width           =   6225
               _ExtentX        =   10980
               _ExtentY        =   1402
               _Version        =   393217
               TextRTF         =   $"FRMCUST_CC.frx":2121
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
               TabIndex        =   87
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
               TabIndex        =   86
               Top             =   0
               Width           =   1035
            End
         End
         Begin VB.ComboBox CmbDtQuality 
            Height          =   315
            Index           =   0
            Left            =   3225
            TabIndex        =   74
            Top             =   -405
            Visible         =   0   'False
            Width           =   1950
         End
         Begin VB.ComboBox CmbDtQuality 
            Height          =   315
            Index           =   1
            Left            =   3660
            Sorted          =   -1  'True
            TabIndex        =   44
            Top             =   -405
            Visible         =   0   'False
            Width           =   4005
         End
         Begin VB.Frame Frame25 
            Enabled         =   0   'False
            Height          =   615
            Left            =   9810
            TabIndex        =   75
            Top             =   2700
            Width           =   4470
            Begin VB.Label Label7 
               Caption         =   "Desc:"
               ForeColor       =   &H00800080&
               Height          =   225
               Left            =   90
               TabIndex        =   76
               Top             =   285
               Width           =   435
            End
         End
         Begin VB.ComboBox CmbNotContacted 
            Height          =   315
            Index           =   0
            Left            =   9120
            TabIndex        =   96
            TabStop         =   0   'False
            Top             =   2160
            Visible         =   0   'False
            Width           =   1260
         End
         Begin RichTextLib.RichTextBox RichTextBox1 
            Height          =   615
            Index           =   0
            Left            =   6240
            TabIndex        =   132
            Top             =   3240
            Visible         =   0   'False
            Width           =   4020
            _ExtentX        =   7091
            _ExtentY        =   1085
            _Version        =   393217
            TextRTF         =   $"FRMCUST_CC.frx":219C
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
            Alignment       =   1  'Right Justify
            Caption         =   "Choose :"
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
            Index           =   12
            Left            =   120
            TabIndex        =   159
            Top             =   960
            Width           =   795
         End
         Begin VB.Label Label9 
            Caption         =   "Desc:"
            ForeColor       =   &H00800080&
            Height          =   225
            Index           =   0
            Left            =   8520
            TabIndex        =   140
            Top             =   1320
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label Label10 
            Caption         =   "Send POD :"
            ForeColor       =   &H00800080&
            Height          =   390
            Left            =   8520
            TabIndex        =   139
            Top             =   1530
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label11 
            Caption         =   "PU POD :"
            ForeColor       =   &H00800080&
            Height          =   390
            Left            =   8535
            TabIndex        =   138
            Top             =   1935
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Alamat Kirim :"
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
            Index           =   4
            Left            =   8115
            TabIndex        =   133
            Top             =   2385
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label Label9 
            Caption         =   "Data Quality:"
            ForeColor       =   &H00800080&
            Height          =   225
            Index           =   1
            Left            =   1995
            TabIndex        =   90
            Top             =   -405
            Visible         =   0   'False
            Width           =   1035
         End
      End
      Begin VB.Frame Frame9 
         Height          =   4305
         Left            =   -74940
         TabIndex        =   72
         Top             =   315
         Width           =   9840
         Begin MSComctlLib.ListView ListView1 
            Height          =   4110
            Index           =   1
            Left            =   15
            TabIndex        =   36
            Top             =   165
            Width           =   9795
            _ExtentX        =   17277
            _ExtentY        =   7250
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
         Height          =   4155
         Left            =   -74940
         TabIndex        =   52
         Top             =   405
         Width           =   9825
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
            TabIndex        =   65
            Top             =   2250
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
               TabIndex        =   42
               Top             =   240
               Width           =   3990
            End
            Begin VB.ComboBox Combo1 
               BackColor       =   &H00C0FFFF&
               ForeColor       =   &H00C00000&
               Height          =   315
               Index           =   16
               Left            =   1470
               TabIndex        =   67
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
               TabIndex        =   66
               Top             =   2085
               Visible         =   0   'False
               Width           =   3300
            End
            Begin RichTextLib.RichTextBox RichTextBox1 
               Height          =   615
               Index           =   2
               Left            =   1125
               TabIndex        =   43
               Top             =   540
               Width           =   4020
               _ExtentX        =   7091
               _ExtentY        =   1085
               _Version        =   393217
               TextRTF         =   $"FRMCUST_CC.frx":2217
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
               TabIndex        =   71
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
               TabIndex        =   70
               Top             =   570
               Width           =   1035
            End
            Begin VB.Label Label3 
               Caption         =   "Gaji / Bulan"
               ForeColor       =   &H00800080&
               Height          =   225
               Index           =   20
               Left            =   105
               TabIndex        =   69
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
               TabIndex        =   68
               Top             =   2130
               Visible         =   0   'False
               Width           =   1095
            End
         End
         Begin VB.Frame Frame18 
            Height          =   1770
            Left            =   5340
            TabIndex        =   61
            Top             =   2235
            Width           =   4440
            Begin VB.CheckBox CekTelp 
               Caption         =   "Invalid"
               Height          =   270
               Index           =   7
               Left            =   2640
               TabIndex        =   34
               Top             =   1305
               Width           =   855
            End
            Begin VB.CheckBox CekTelp 
               Caption         =   "Invalid"
               Height          =   270
               Index           =   6
               Left            =   2640
               TabIndex        =   31
               Top             =   945
               Width           =   855
            End
            Begin VB.CheckBox CekTelp 
               Caption         =   "Invalid"
               Height          =   270
               Index           =   5
               Left            =   3630
               TabIndex        =   28
               Top             =   540
               Width           =   780
            End
            Begin VB.CheckBox CekTelp 
               Caption         =   "Invalid"
               Height          =   270
               Index           =   4
               Left            =   3630
               TabIndex        =   24
               Top             =   210
               Width           =   780
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
               Left            =   2910
               MaxLength       =   6
               TabIndex        =   23
               Top             =   195
               Width           =   675
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
               Left            =   2910
               MaxLength       =   6
               TabIndex        =   27
               Top             =   540
               Width           =   720
            End
            Begin TDBMask6Ctl.TDBMask TDBMaskOffice 
               Height          =   360
               Index           =   0
               Left            =   1410
               TabIndex        =   22
               Top             =   165
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC.frx":2292
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC.frx":22FE
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
               Left            =   1410
               TabIndex        =   26
               Top             =   540
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC.frx":2340
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC.frx":23AC
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
               Left            =   1410
               TabIndex        =   30
               Top             =   915
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC.frx":23EE
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC.frx":245A
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
               Left            =   1410
               TabIndex        =   33
               Top             =   1275
               Width           =   1200
               _Version        =   65536
               _ExtentX        =   2117
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC.frx":249C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC.frx":2508
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
               Left            =   810
               TabIndex        =   21
               Top             =   165
               Width           =   585
               _Version        =   65536
               _ExtentX        =   1032
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC.frx":254A
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC.frx":25B6
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
               Left            =   810
               TabIndex        =   25
               Top             =   540
               Width           =   585
               _Version        =   65536
               _ExtentX        =   1032
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC.frx":25F8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC.frx":2664
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
               Left            =   810
               TabIndex        =   29
               Top             =   915
               Width           =   585
               _Version        =   65536
               _ExtentX        =   1032
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC.frx":26A6
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC.frx":2712
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
               Left            =   810
               TabIndex        =   32
               Top             =   1275
               Width           =   585
               _Version        =   65536
               _ExtentX        =   1032
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC.frx":2754
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC.frx":27C0
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
               Left            =   2655
               TabIndex        =   64
               Top             =   255
               Width           =   225
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
               Left            =   60
               TabIndex        =   63
               Top             =   240
               Width           =   750
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
               Left            =   75
               TabIndex        =   62
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
            Height          =   1800
            Index           =   1
            Left            =   5340
            TabIndex        =   58
            Top             =   180
            Width           =   4425
            Begin VB.CheckBox CekTelp 
               Caption         =   "Invalid"
               Height          =   270
               Index           =   3
               Left            =   2745
               TabIndex        =   20
               Top             =   1380
               Width           =   855
            End
            Begin VB.CheckBox CekTelp 
               Caption         =   "Invalid"
               Height          =   270
               Index           =   2
               Left            =   2730
               TabIndex        =   18
               Top             =   1020
               Width           =   855
            End
            Begin VB.CheckBox CekTelp 
               Caption         =   "Invalid"
               Height          =   270
               Index           =   1
               Left            =   2925
               TabIndex        =   16
               Top             =   630
               Width           =   855
            End
            Begin VB.CheckBox CekTelp 
               Caption         =   "Invalid"
               Height          =   270
               Index           =   0
               Left            =   2925
               TabIndex        =   13
               Top             =   285
               Width           =   855
            End
            Begin TDBMask6Ctl.TDBMask TDBMaskHome 
               Height          =   360
               Index           =   0
               Left            =   1470
               TabIndex        =   12
               Top             =   225
               Width           =   1425
               _Version        =   65536
               _ExtentX        =   2514
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC.frx":2802
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC.frx":286E
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
               Left            =   1470
               TabIndex        =   15
               Top             =   585
               Width           =   1425
               _Version        =   65536
               _ExtentX        =   2514
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC.frx":28B0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC.frx":291C
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
               Left            =   855
               TabIndex        =   17
               Top             =   945
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC.frx":295E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC.frx":29CA
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
               Left            =   855
               TabIndex        =   19
               Top             =   1320
               Width           =   1860
               _Version        =   65536
               _ExtentX        =   3281
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC.frx":2A0C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC.frx":2A78
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
               Left            =   855
               TabIndex        =   11
               Top             =   225
               Width           =   585
               _Version        =   65536
               _ExtentX        =   1032
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC.frx":2ABA
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC.frx":2B26
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
               Left            =   855
               TabIndex        =   14
               Top             =   585
               Width           =   585
               _Version        =   65536
               _ExtentX        =   1032
               _ExtentY        =   635
               Caption         =   "FRMCUST_CC.frx":2B68
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Keys            =   "FRMCUST_CC.frx":2BD4
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
               Left            =   60
               TabIndex        =   60
               Top             =   300
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
               Left            =   75
               TabIndex        =   59
               Top             =   975
               Width           =   795
            End
         End
         Begin VB.Frame Frame3 
            Height          =   1800
            Index           =   0
            Left            =   60
            TabIndex        =   53
            Top             =   180
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
               TabIndex        =   7
               Top             =   750
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
               TabIndex        =   6
               Top             =   750
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
               TabIndex        =   9
               TabStop         =   0   'False
               Text            =   "0"
               Top             =   1080
               Width           =   375
            End
            Begin TDBDate6Ctl.TDBDate TDBDate1 
               Height          =   315
               Index           =   0
               Left            =   1140
               TabIndex        =   8
               Top             =   1080
               Width           =   1335
               _Version        =   65536
               _ExtentX        =   2355
               _ExtentY        =   556
               Calendar        =   "FRMCUST_CC.frx":2C16
               Caption         =   "FRMCUST_CC.frx":2D2E
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FRMCUST_CC.frx":2D9A
               Keys            =   "FRMCUST_CC.frx":2DB8
               Spin            =   "FRMCUST_CC.frx":2E16
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
               TabIndex        =   5
               Top             =   210
               Width           =   4020
               _ExtentX        =   7091
               _ExtentY        =   979
               _Version        =   393217
               TextRTF         =   $"FRMCUST_CC.frx":2E3E
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
               TabIndex        =   57
               Top             =   795
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
               TabIndex        =   56
               Top             =   315
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
               TabIndex        =   55
               Top             =   795
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
               TabIndex        =   54
               Top             =   1140
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
               TabIndex        =   10
               Top             =   1125
               Width           =   780
            End
         End
      End
      Begin TDBDate6Ctl.TDBDate TDBDate1 
         Height          =   300
         Index           =   2
         Left            =   -73290
         TabIndex        =   128
         Top             =   2580
         Width           =   1350
         _Version        =   65536
         _ExtentX        =   2381
         _ExtentY        =   529
         Calendar        =   "FRMCUST_CC.frx":2EB9
         Caption         =   "FRMCUST_CC.frx":2FD1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FRMCUST_CC.frx":303D
         Keys            =   "FRMCUST_CC.frx":305B
         Spin            =   "FRMCUST_CC.frx":30B9
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
         Format          =   "dd-mm-yyyy"
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
         Text            =   "__-__-____"
         ValidateMode    =   0
         ValueVT         =   1
         Value           =   37468
         CenturyMode     =   0
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
      Index           =   0
      Left            =   1500
      Sorted          =   -1  'True
      TabIndex        =   45
      Top             =   975
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   2970
      Sorted          =   -1  'True
      TabIndex        =   46
      Top             =   975
      Visible         =   0   'False
      Width           =   3495
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
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   975
      Visible         =   0   'False
      Width           =   4410
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   420
      Index           =   1
      Left            =   195
      TabIndex        =   94
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
      Picture         =   "FRMCUST_CC.frx":30E1
      Caption         =   "&Close"
      AutoSize        =   2
      Alignment       =   4
      ButtonStyle     =   2
      PictureAlignment=   1
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   420
      Index           =   4
      Left            =   75
      TabIndex        =   97
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
      Picture         =   "FRMCUST_CC.frx":3403
      Caption         =   "&Close"
      AutoSize        =   2
      Alignment       =   4
      ButtonStyle     =   2
      PictureAlignment=   1
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
      TabIndex        =   48
      Top             =   1020
      Visible         =   0   'False
      Width           =   1260
   End
End
Attribute VB_Name = "FRMCUST_CC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_referall_mgm As ADODB.Recordset
Dim pStatusLstCall As String
Dim M_followup As Boolean
Dim pStatusHstLstCall As String
Public closeOk As Boolean

