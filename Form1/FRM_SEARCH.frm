VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form FRM_SEARCH 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7455
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "FRM_SEARCH.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FRM_SEARCH.frx":0442
   ScaleHeight     =   5775
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbStsLastCall 
      Height          =   315
      Index           =   0
      Left            =   1710
      TabIndex        =   11
      Top             =   3180
      Width           =   3495
   End
   Begin VB.ComboBox CmbStatusCek 
      Height          =   315
      Left            =   1710
      TabIndex        =   16
      Top             =   3810
      Width           =   2160
   End
   Begin VB.CheckBox CekDtDistribute 
      BackColor       =   &H000000FF&
      Caption         =   "Check2"
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   4395
      MaskColor       =   &H000000FF&
      TabIndex        =   1
      Top             =   705
      Width           =   225
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   690
      Width           =   225
   End
   Begin VB.Frame Frame1 
      Caption         =   "Proses....!!"
      Height          =   675
      Left            =   15
      TabIndex        =   28
      Top             =   4710
      Visible         =   0   'False
      Width           =   3990
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   390
         Left            =   60
         TabIndex        =   29
         Top             =   210
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   1
      End
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
      Index           =   2
      Left            =   1710
      MaxLength       =   20
      TabIndex        =   2
      Top             =   975
      Width           =   2325
   End
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   315
      Left            =   1710
      TabIndex        =   4
      Top             =   1605
      Width           =   1530
      _Version        =   65536
      _ExtentX        =   2699
      _ExtentY        =   556
      Calendar        =   "FRM_SEARCH.frx":2634B
      Caption         =   "FRM_SEARCH.frx":26463
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FRM_SEARCH.frx":264CF
      Keys            =   "FRM_SEARCH.frx":264ED
      Spin            =   "FRM_SEARCH.frx":2654B
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
      Value           =   37475
      CenturyMode     =   0
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
      Index           =   3
      Left            =   3615
      TabIndex        =   8
      Top             =   2235
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
      Index           =   2
      Left            =   1710
      TabIndex        =   7
      Top             =   2235
      Width           =   1905
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
      Left            =   3615
      TabIndex        =   6
      Top             =   1920
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
      Left            =   1710
      TabIndex        =   5
      Top             =   1920
      Width           =   1905
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
      Left            =   1710
      TabIndex        =   3
      Top             =   1290
      Width           =   4590
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   1164
      _Version        =   196610
      Font3D          =   5
      ForeColor       =   4194368
      BackColor       =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Search"
      BevelWidth      =   2
      BorderWidth     =   2
      BevelInner      =   1
      FloodColor      =   4194368
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin TDBMask6Ctl.TDBMask TDBMask1 
      Height          =   315
      Left            =   1710
      TabIndex        =   9
      Top             =   2550
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Caption         =   "FRM_SEARCH.frx":26573
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FRM_SEARCH.frx":265DF
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   0
      Format          =   "999-99999"
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
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "___-_____"
      Value           =   ""
   End
   Begin TDBMask6Ctl.TDBMask TDBMask2 
      Height          =   315
      Left            =   1710
      TabIndex        =   10
      Top             =   2865
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Caption         =   "FRM_SEARCH.frx":26621
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FRM_SEARCH.frx":2668D
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   -2147483643
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   0
      Format          =   "9999-99999999"
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
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "____-________"
      Value           =   ""
   End
   Begin Threed.SSCommand Command1 
      Height          =   360
      Index           =   0
      Left            =   330
      TabIndex        =   17
      Top             =   4260
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   635
      _Version        =   196610
      Font3D          =   5
      MousePointer    =   16
      BackColor       =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Search"
      ButtonStyle     =   2
   End
   Begin Threed.SSCommand Command1 
      Height          =   360
      Index           =   2
      Left            =   1140
      TabIndex        =   18
      Top             =   4260
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   635
      _Version        =   196610
      Font3D          =   5
      MousePointer    =   16
      BackColor       =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Clear"
      ButtonStyle     =   2
   End
   Begin Threed.SSCommand Command1 
      Cancel          =   -1  'True
      Height          =   360
      Index           =   1
      Left            =   1965
      TabIndex        =   19
      Top             =   4260
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   635
      _Version        =   196610
      Font3D          =   5
      MousePointer    =   16
      BackColor       =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Close"
      ButtonStyle     =   2
   End
   Begin TDBTime6Ctl.TDBTime DTimeLastCall 
      Height          =   300
      Index           =   0
      Left            =   3105
      TabIndex        =   13
      Top             =   3495
      Width           =   825
      _Version        =   65536
      _ExtentX        =   1455
      _ExtentY        =   529
      Caption         =   "FRM_SEARCH.frx":266CF
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FRM_SEARCH.frx":2673B
      Spin            =   "FRM_SEARCH.frx":2678B
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
      Value           =   0.870289351851852
   End
   Begin TDBDate6Ctl.TDBDate DtLastCall 
      Height          =   315
      Index           =   0
      Left            =   1710
      TabIndex        =   12
      Top             =   3495
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Calendar        =   "FRM_SEARCH.frx":267B3
      Caption         =   "FRM_SEARCH.frx":268CB
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FRM_SEARCH.frx":26937
      Keys            =   "FRM_SEARCH.frx":26955
      Spin            =   "FRM_SEARCH.frx":269B3
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
   Begin TDBDate6Ctl.TDBDate DtLastCall 
      Height          =   315
      Index           =   1
      Left            =   4230
      TabIndex        =   14
      Top             =   3495
      Width           =   1395
      _Version        =   65536
      _ExtentX        =   2461
      _ExtentY        =   556
      Calendar        =   "FRM_SEARCH.frx":269DB
      Caption         =   "FRM_SEARCH.frx":26AF3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FRM_SEARCH.frx":26B5F
      Keys            =   "FRM_SEARCH.frx":26B7D
      Spin            =   "FRM_SEARCH.frx":26BDB
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
   Begin TDBTime6Ctl.TDBTime DTimeLastCall 
      Height          =   300
      Index           =   1
      Left            =   5625
      TabIndex        =   15
      Top             =   3495
      Width           =   825
      _Version        =   65536
      _ExtentX        =   1455
      _ExtentY        =   529
      Caption         =   "FRM_SEARCH.frx":26C03
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FRM_SEARCH.frx":26C6F
      Spin            =   "FRM_SEARCH.frx":26CBF
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
      Text            =   ""
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   0.870289351851852
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Status Last Call"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   8
      Left            =   210
      TabIndex        =   35
      Top             =   3255
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Last Call Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   9
      Left            =   210
      TabIndex        =   34
      Top             =   3555
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Status Check"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   10
      Left            =   210
      TabIndex        =   33
      Top             =   3855
      Width           =   1365
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   11
      Left            =   3960
      TabIndex        =   32
      Top             =   3525
      Width           =   225
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Searching Data Belum Distribute"
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
      Height          =   255
      Index           =   12
      Left            =   4680
      TabIndex        =   31
      Top             =   705
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MGM Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   390
      TabIndex        =   30
      Top             =   705
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ref No. "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   150
      TabIndex        =   27
      Top             =   1005
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cellular No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   6
      Left            =   225
      TabIndex        =   25
      Top             =   2940
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Home Phone No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   225
      TabIndex        =   24
      Top             =   2610
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Batch Database"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   225
      TabIndex        =   23
      Top             =   2280
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Height          =   300
      Index           =   2
      Left            =   210
      TabIndex        =   22
      Top             =   1650
      Width           =   1410
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   225
      TabIndex        =   21
      Top             =   1980
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   135
      TabIndex        =   20
      Top             =   1320
      Width           =   1485
   End
End
Attribute VB_Name = "FRM_SEARCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_cari As ADODB.Recordset

Private Sub Combo1_Click(Index As Integer)
Dim M_DATA As New CLS_FRMSEARCH
Dim m_objrs As ADODB.Recordset
Select Case Index
Case 0
    Set m_objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "USERID = '" + Combo1(Index).Text + "'")
    If m_objrs.RecordCount <> 0 Then
        Combo1(0).Text = m_objrs("USERID")
        Combo1(1).Text = m_objrs("AGENT")
    Else
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
    End If
Case 1
    Set m_objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "AGENT = '" + Combo1(Index).Text + "'")
    If m_objrs.RecordCount <> 0 Then
        Combo1(0).Text = m_objrs("USERID")
        Combo1(1).Text = m_objrs("AGENT")
    Else
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
    End If
Case 2
Set m_objrs = M_DATA.QUERY_DATASOURCE(M_OBJCONN, "KODEDS = '" + Combo1(Index).Text + "'")
    If m_objrs.RecordCount <> 0 Then
        Combo1(2).Text = m_objrs("KODEDS")
        Combo1(3).Text = m_objrs("KETERANGAN")
    Else
        Combo1(2).Text = Empty
        Combo1(3).Text = Empty
    End If
Case 3
Set m_objrs = M_DATA.QUERY_DATASOURCE(M_OBJCONN, "KETERANGAN = '" + Combo1(Index).Text + "'")
    If m_objrs.RecordCount <> 0 Then
        Combo1(2).Text = m_objrs("KODEDS")
        Combo1(3).Text = m_objrs("KETERANGAN")
    Else
        Combo1(2).Text = Empty
        Combo1(3).Text = Empty
    End If
End Select
Set M_DATA = Nothing
Set m_objrs = Nothing
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
Dim sSearchText As String
Dim lReturn As Long
Select Case Index
Case 0, 1, 2, 3
If KeyAscii = 13 Then
   Combo1_Click (Index)
   KeyAscii = 0
Else
   sSearchText = Left$(Combo1(Index).Text, Combo1(Index).SelStart) & Chr$(KeyAscii)
   lReturn = SendMessage(Combo1(Index).hWnd, CB_FINDSTRING, -1, ByVal sSearchText)
   If lReturn <> CB_ERR Then
      mbIgnoreListClick = True
      Combo1(Index).ListIndex = lReturn
      mbIgnoreListClick = False
      Combo1(Index).Text = Combo1(Index).List(lReturn)
      Combo1(Index).SelStart = Len(sSearchText)
      Combo1(Index).SelLength = Len(Combo1(Index).Text)
      KeyAscii = 0
   End If
End If
End Select
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim M_DATA As New CLS_FRMSEARCH
Dim m_objrs As ADODB.Recordset
Select Case Index
Case 0
    Set m_objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "USERID = '" + Combo1(Index).Text + "'")
    If m_objrs.RecordCount <> 0 Then
        Combo1(0).Text = m_objrs("USERID")
        Combo1(1).Text = m_objrs("AGENT")
    Else
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
    End If
Case 1
    Set m_objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "AGENT = '" + Combo1(Index).Text + "'")
    If m_objrs.RecordCount <> 0 Then
        Combo1(0).Text = m_objrs("USERID")
        Combo1(1).Text = m_objrs("AGENT")
    Else
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
    End If
Case 2
Set m_objrs = M_DATA.QUERY_DATASOURCE(M_OBJCONN, "KODEDS = '" + Combo1(Index).Text + "'")
    If m_objrs.RecordCount <> 0 Then
        Combo1(2).Text = m_objrs("KODEDS")
        Combo1(3).Text = m_objrs("KETERANGAN")
    Else
        Combo1(2).Text = Empty
        Combo1(3).Text = Empty
    End If
Case 3
Set m_objrs = M_DATA.QUERY_DATASOURCE(M_OBJCONN, "KETERANGAN = '" + Combo1(Index).Text + "'")
    If m_objrs.RecordCount <> 0 Then
        Combo1(2).Text = m_objrs("KODEDS")
        Combo1(3).Text = m_objrs("KETERANGAN")
    Else
        Combo1(2).Text = Empty
        Combo1(3).Text = Empty
    End If
End Select
Set M_DATA = Nothing
Set m_objrs = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
Dim NAMACUST As String
Dim NAMAAGENT As String
Dim DATASOURCE As String
Dim TGLLAHIR As String
Dim OFFPHONE As String
Dim OFFPHONE2 As String
Dim HOMEPHONE As String
Dim HOMEPHONE2 As String
Dim MOBILEPHONE As String
Dim MOBILEPHONE2 As String
Dim FAXPHONE As String
Dim Lcustid As String
Dim FAXPHONE2 As String
Dim KETHSLKERJA As String
Dim lLastCallDate As String
Dim lStatusCek As String

Dim M_DATA As New CLS_FRMSEARCH
Dim PANJANG As Integer
Select Case Index
    Case 0
        If Text1(0).Text = Empty And Combo1(0).Text = Empty And Combo1(2).Text = Empty And Len(TDBMask2.Value) < 1 And Len(TDBMask1.Value) < 1 And TDBDate1.ValueIsNull And Len(Text1(2).Text) < 3 _
        And cmbStsLastCall(0).Text = "" And CmbStatusCek.Text = "" And DtLastCall(0).ValueIsNull Then
            MsgBox "Masukan Kriteria Customer Yang Akan Dicari...!!!", vbCritical + vbOKOnly, "Peringatan"
            Text1(0).SetFocus
            Set M_DATA = Nothing
            Exit Sub
        Else
            If Text1(2).Text <> Empty Then
                Lcustid = "CUSTID LIKE " + "'%" + UBAH_QUOTE(Text1(2).Text) + "%'"
            Else
                If Text1(0).Text <> Empty Then
                    NAMACUST = "NAME LIKE " + "'%" + UBAH_QUOTE(Text1(0).Text) + "%'"
                End If
                If Combo1(0).Text <> Empty Then
                    NAMAAGENT = "AGENT = '" + Combo1(0).Text + "'"
                End If
                If Combo1(2).Text <> Empty Then
                    DATASOURCE = "RECSOURCE = '" + Combo1(2).Text + "'"
                End If
                If TDBDate1.ValueIsNull Then
                Else
                    TGLLAHIR = "BIRTHD = '" + Format(TDBDate1.Text, "mm/dd/yyyy") + "'"
                End If
                If Len(TDBMask1.Value) > 1 Then
                    OFFPHONE = "OFFICENO Like '%" + TDBMask1.Value + "%'"
                    OFFPHONE2 = "OFFICENO2 Like '%" + TDBMask1.Value + "%'"
                    HOMEPHONE = "HOMENO Like '%" + TDBMask1.Value + "%'"
                    HOMEPHONE2 = "HOMENO2 Like '%" + TDBMask1.Value + "%'"
                    FAXPHONE = "FAXNO Like '%" + TDBMask1.Value + "%'"
                    FAXPHONE2 = "FAXNO2 Like '%" + TDBMask1.Value + "%'"
                End If
                If Len(TDBMask2.Value) > 1 Then
                    MOBILEPHONE = "MOBILENO like '%" + TDBMask2.Value + "%'"
                    MOBILEPHONE2 = "MOBILENO2 like '%" + TDBMask2.Value + "%'"
                End If
                
                If cmbStsLastCall(0).Text <> "" Then
                    Select Case UCase(Trim(cmbStsLastCall(0).Text))
                        Case "NOT CONTACTED"
                            KETHSLKERJA = "KETHSLKERJA IN('WN','WN','NP','BT')"
                        Case "NOT AVAILABLE"
                            KETHSLKERJA = "KETHSLKERJA = 'NA'"
                        Case "STILL THINKING"
                            KETHSLKERJA = "KETHSLKERJA= 'ST'"
                        Case "DISAGREE"
                            KETHSLKERJA = "LEFT(KETHSLKERJA,1)= 'D'"
                        Case "SENDING APPLICATION"
                            KETHSLKERJA = "KETHSLKERJA= 'SA'"
                            'KETHSLKERJA = "KETHSLKERJA= 'TBO'"
                        Case "PICKUP"
                            KETHSLKERJA = "KETHSLKERJA= 'PU'"
                            'KETHSLKERJA = "KETHSLKERJA= 'TBO'"
                        Case "INCOMPLETE DOCUMENT"
                            KETHSLKERJA = "KETHSLKERJA= 'ID'"
                            'KETHSLKERJA = "KETHSLKERJA= 'TBO'"
                        Case "INCOMING"
                            KETHSLKERJA = "KETHSLKERJA= 'I'"
                    End Select
                End If
            
                If CmbStatusCek.Text <> "" Then
                    Select Case UCase(Trim(CmbStatusCek.Text))
                        Case "ACCEPT"
                            lStatusCek = "UPPER(F_cek) ='ACCEPT'"
                        Case "RETURN"
                            lStatusCek = "UPPER(F_cek) = 'RETURN'"
                        Case "NOT CHECK"
                            lStatusCek = "(f_cek is null or   f_cek = 'NOT ACCEPT' OR F_CEK ='') AND KETHSLKERJA ='I'"
                    End Select
                End If
                
                If DtLastCall(0).ValueIsNull Then
                Else
                    lLastCallDate = "TGLSTATUS BETWEEN '" + Format(DtLastCall(0).Value, "MM/DD/YYYY") & " " & CStr(DTimeLastCall(0).Value) + "' AND '" + Format(DtLastCall(1).Value, "MM/DD/YYYY") & " " & CStr(DTimeLastCall(1).Value) + "'"
                End If
        
        End If
                        
                'Unload FRM_SEARCH
                If Check1.Value = 0 Then
                    Set m_cari = M_DATA.QUERY_SEARCH_CONDITION(M_OBJCONN, NAMACUST, NAMAAGENT, DATASOURCE, TGLLAHIR, _
                                                            OFFPHONE, OFFPHONE2, HOMEPHONE, HOMEPHONE2, MOBILEPHONE, _
                                                            MOBILEPHONE2, FAXPHONE, FAXPHONE2, MDIForm1.Text3.Text, Lcustid, KETHSLKERJA, lLastCallDate, lStatusCek)
                Else
                    Set m_cari = M_DATA.QUERY_SEARCH_CONDITION_mgm(M_OBJCONN, NAMACUST, NAMAAGENT, DATASOURCE, TGLLAHIR, _
                                                            OFFPHONE, OFFPHONE2, HOMEPHONE, HOMEPHONE2, MOBILEPHONE, _
                                                            MOBILEPHONE2, FAXPHONE, FAXPHONE2, MDIForm1.Text3.Text, Lcustid, KETHSLKERJA, lLastCallDate, lStatusCek)
                End If
        
            If m_cari.RecordCount = 0 Then
                MsgBox "Data Tidak Ditemukan", vbInformation + vbOKOnly, "TeleGrandi"
                Set M_DATA = Nothing
                Exit Sub
            Else
                search_ok = True
                If Check1.Value = 0 Then
                    Unload FRM_PRESCREEN
                    FRM_PRESCREEN.Caption = "Search Non Mgm Data"
                    FRM_PRESCREEN.Show
                Else
                    Unload VIEW_MGMDATA
                    If MDIForm1.Text2.Text = "Agent" Then
                        VIEW_MGMDATA.Caption = "Search MGM Data"
                    Else
                        VIEW_MGMDATA.Caption = "Search MGM Data  .... Tekan Huruf ""P"" untuk Mengalihkan Data"
                    End If
                    
                    
                    VIEW_MGMDATA.Show
                End If
            'FRM_PRESCREEN.Show vbModal
                Unload Me
            End If
        End If
    Case 1
        Unload Me
    Case 2
        Text1(2).Text = Empty
        Text1(0).Text = Empty
        TDBDate1.Text = Empty
        Combo1(0).Text = Empty
        Combo1(1).Text = Empty
        Combo1(2).Text = Empty
        Combo1(3).Text = Empty
        TDBMask1.Text = Empty
        TDBMask2.Text = Empty
End Select
Set M_DATA = Nothing
End Sub

Private Sub Form_Load()
Dim m_objrs As ADODB.Recordset
Dim M_DATA As New CLS_FRMSEARCH

If UCase(MDIForm1.Text2.Text) = "AGENT" Then
    CekDtDistribute.Visible = False
    Label1(12).Visible = False
End If
    CmbStatusCek.AddItem "ACCEPT"
    CmbStatusCek.AddItem "RETURN"
    CmbStatusCek.AddItem "NOT CHECK"
    

cmbStsLastCall(0).AddItem "NOT CONTACTED"
cmbStsLastCall(0).AddItem "NOT AVAILABLE"
cmbStsLastCall(0).AddItem "STILL THINKING"
cmbStsLastCall(0).AddItem "DISAGREE"
cmbStsLastCall(0).AddItem "SENDING APPLICATION"
cmbStsLastCall(0).AddItem "PICKUP"
cmbStsLastCall(0).AddItem "INCOMPLETE DOCUMENT"
cmbStsLastCall(0).AddItem "INCOMING"

Me.Top = 1500
Me.Left = 4000
DTimeLastCall(0).Value = "00:00"
DTimeLastCall(1).Value = "23:59"
StsMgmSchedule = False

Set m_objrs = M_DATA.QUERY_AGENT_JADWAL(M_OBJCONN, "")
    While Not m_objrs.EOF
        Combo1(0).AddItem m_objrs("USERID")
        Combo1(0).DataField = m_objrs("USERID")
        Combo1(1).AddItem m_objrs("AGENT")
        Combo1(1).DataField = m_objrs("AGENT")
        m_objrs.MoveNext
    Wend
Set m_objrs = Nothing
Set m_objrs = M_DATA.QUERY_DATASOURCE(M_OBJCONN, "")
    While Not m_objrs.EOF
        Combo1(2).AddItem m_objrs("KODEDS")
        Combo1(2).DataField = m_objrs("KODEDS")
        Combo1(3).AddItem m_objrs("KETERANGAN")
        Combo1(3).DataField = m_objrs("KETERANGAN")
        m_objrs.MoveNext
    Wend
If UCase(MDIForm1.Text3.Text) = "ADMIN" Then
    Label1(5).Visible = True
    Text1(2).Visible = True
End If
m_objrs.Close
Set m_objrs = Nothing
Set M_DATA = Nothing
End Sub



Private Sub Form_Unload(Cancel As Integer)
Set m_cari = Nothing
Frame1.Visible = False
ProgressBar1.Value = 0
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Command1_Click(0)
End If

Select Case Index
Case 1
    Select Case KeyAscii
        Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 8
            Exit Sub
        Case Else
            KeyAscii = 0
    End Select
End Select
End Sub

