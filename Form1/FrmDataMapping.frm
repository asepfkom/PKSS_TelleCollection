VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{66A90C01-346D-11D2-9BC0-00A024695830}#1.0#0"; "timask6.ocx"
Begin VB.Form FrmDataMapping 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Data Mapping PIL"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   19005
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   19005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame17 
      BackColor       =   &H00000000&
      Height          =   1215
      Left            =   10950
      TabIndex        =   145
      Top             =   5130
      Width           =   7845
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H0030FCED&
         Height          =   885
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   146
         Top             =   210
         Width           =   7620
      End
   End
   Begin VB.TextBox TxtNoPayment 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   8850
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3945
      Width           =   795
   End
   Begin VB.TextBox txtBucket 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   8850
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   4245
      Width           =   795
   End
   Begin VB.TextBox TxtTipeAccount 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0080FF80&
      Height          =   225
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "--"
      Top             =   6000
      Width           =   765
   End
   Begin VB.TextBox txtTglSource 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   8850
      TabIndex        =   7
      Top             =   3630
      Width           =   1695
   End
   Begin VB.TextBox StsAcc 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0080FF80&
      Height          =   225
      Left            =   13005
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   5460
      Width           =   1005
   End
   Begin VB.TextBox lblAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0080FF80&
      Height          =   675
      Left            =   1110
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2040
      Width           =   3330
   End
   Begin VB.TextBox lblOfficeAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0080FF80&
      Height          =   675
      Left            =   1110
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2760
      Width           =   3330
   End
   Begin VB.TextBox LblCHAditionalAddr 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   1110
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3450
      Width           =   3330
   End
   Begin VB.TextBox AddrNow 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0030FCED&
      Height          =   675
      Left            =   11940
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   3000
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.TextBox TxtEC 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0030FCED&
      Height          =   345
      Left            =   11940
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "--"
      Top             =   3810
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.TextBox TxtMap 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Enabled         =   0   'False
      Height          =   285
      Left            =   8850
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   4530
      Width           =   795
   End
   Begin TDBDate6Ctl.TDBDate lblDate 
      Height          =   285
      Left            =   2730
      TabIndex        =   11
      Top             =   1380
      Visible         =   0   'False
      Width           =   960
      _Version        =   65536
      _ExtentX        =   1693
      _ExtentY        =   503
      Calendar        =   "FrmDataMapping.frx":0000
      Caption         =   "FrmDataMapping.frx":0118
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmDataMapping.frx":0184
      Keys            =   "FrmDataMapping.frx":01A2
      Spin            =   "FrmDataMapping.frx":0200
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   0
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
      ForeColor       =   8454016
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
   Begin TDBNumber6Ctl.TDBNumber lblPromPA 
      Height          =   300
      Left            =   6090
      TabIndex        =   12
      Top             =   1200
      Width           =   1530
      _Version        =   65536
      _ExtentX        =   2699
      _ExtentY        =   529
      Calculator      =   "FrmDataMapping.frx":0228
      Caption         =   "FrmDataMapping.frx":0248
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmDataMapping.frx":02B4
      Keys            =   "FrmDataMapping.frx":02D2
      Spin            =   "FrmDataMapping.frx":031C
      AlignHorizontal =   1
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
      ForeColor       =   8454016
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
   Begin TDBDate6Ctl.TDBDate lblOpenDate 
      Height          =   300
      Left            =   6090
      TabIndex        =   13
      Top             =   2055
      Width           =   1530
      _Version        =   65536
      _ExtentX        =   2699
      _ExtentY        =   529
      Calendar        =   "FrmDataMapping.frx":0344
      Caption         =   "FrmDataMapping.frx":045C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmDataMapping.frx":04C8
      Keys            =   "FrmDataMapping.frx":04E6
      Spin            =   "FrmDataMapping.frx":0544
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   0
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
      ForeColor       =   8454016
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
      Height          =   300
      Left            =   6090
      TabIndex        =   14
      Top             =   2370
      Width           =   1530
      _Version        =   65536
      _ExtentX        =   2699
      _ExtentY        =   529
      Calendar        =   "FrmDataMapping.frx":056C
      Caption         =   "FrmDataMapping.frx":0684
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmDataMapping.frx":06F0
      Keys            =   "FrmDataMapping.frx":070E
      Spin            =   "FrmDataMapping.frx":076C
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   0
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
      ForeColor       =   8454016
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
      Left            =   6090
      TabIndex        =   15
      Top             =   2700
      Width           =   1530
      _Version        =   65536
      _ExtentX        =   2699
      _ExtentY        =   503
      Calendar        =   "FrmDataMapping.frx":0794
      Caption         =   "FrmDataMapping.frx":08AC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmDataMapping.frx":0918
      Keys            =   "FrmDataMapping.frx":0936
      Spin            =   "FrmDataMapping.frx":0994
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   0
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
      ForeColor       =   8454016
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
   Begin TDBNumber6Ctl.TDBNumber LblInstallment 
      Height          =   285
      Left            =   6090
      TabIndex        =   16
      Top             =   3930
      Width           =   1545
      _Version        =   65536
      _ExtentX        =   2725
      _ExtentY        =   503
      Calculator      =   "FrmDataMapping.frx":09BC
      Caption         =   "FrmDataMapping.frx":09DC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmDataMapping.frx":0A48
      Keys            =   "FrmDataMapping.frx":0A66
      Spin            =   "FrmDataMapping.frx":0AB0
      AlignHorizontal =   1
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
      ForeColor       =   8454016
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
   Begin TDBMask6Ctl.TDBMask AHome1 
      Height          =   300
      Left            =   11895
      TabIndex        =   17
      Top             =   945
      Width           =   525
      _Version        =   65536
      _ExtentX        =   926
      _ExtentY        =   529
      Caption         =   "FrmDataMapping.frx":0AD8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmDataMapping.frx":0B44
      AlignHorizontal =   0
      AlignVertical   =   1
      Appearance      =   0
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   0
      BorderStyle     =   0
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   3210477
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
      PromptChar      =   "_"
      ReadOnly        =   -1
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "[____]"
      Value           =   ""
   End
   Begin TDBMask6Ctl.TDBMask AHome2 
      Height          =   300
      Left            =   11910
      TabIndex        =   18
      Top             =   1290
      Width           =   525
      _Version        =   65536
      _ExtentX        =   926
      _ExtentY        =   529
      Caption         =   "FrmDataMapping.frx":0B86
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmDataMapping.frx":0BF2
      AlignHorizontal =   0
      AlignVertical   =   1
      Appearance      =   0
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   0
      BorderStyle     =   0
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   3210477
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
      PromptChar      =   "_"
      ReadOnly        =   -1
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "[____]"
      Value           =   ""
   End
   Begin TDBMask6Ctl.TDBMask AOffice1 
      Height          =   300
      Left            =   11925
      TabIndex        =   19
      Top             =   1620
      Width           =   525
      _Version        =   65536
      _ExtentX        =   926
      _ExtentY        =   529
      Caption         =   "FrmDataMapping.frx":0C34
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmDataMapping.frx":0CA0
      AlignHorizontal =   0
      AlignVertical   =   1
      Appearance      =   0
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   0
      BorderStyle     =   0
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   3210477
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
      PromptChar      =   "_"
      ReadOnly        =   -1
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "[____]"
      Value           =   ""
   End
   Begin TDBMask6Ctl.TDBMask TxtExt1 
      Height          =   315
      Left            =   13635
      TabIndex        =   20
      Top             =   1605
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   556
      Caption         =   "FrmDataMapping.frx":0CE2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmDataMapping.frx":0D4E
      AlignHorizontal =   0
      AlignVertical   =   1
      Appearance      =   0
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   0
      BorderStyle     =   0
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   3210477
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
      Left            =   13635
      TabIndex        =   21
      Top             =   1965
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   556
      Caption         =   "FrmDataMapping.frx":0D90
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmDataMapping.frx":0DFC
      AlignHorizontal =   0
      AlignVertical   =   1
      Appearance      =   0
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   0
      BorderStyle     =   0
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   3210477
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
   Begin TDBMask6Ctl.TDBMask txtECno 
      Height          =   330
      Left            =   11940
      TabIndex        =   22
      Top             =   4140
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   582
      Caption         =   "FrmDataMapping.frx":0E3E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmDataMapping.frx":0EAA
      AlignHorizontal =   0
      AlignVertical   =   1
      Appearance      =   0
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   0
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   3210477
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
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "__________________"
      Value           =   ""
   End
   Begin TDBMask6Ctl.TDBMask txtFaxAdd1 
      Height          =   330
      Left            =   16230
      TabIndex        =   23
      Top             =   2310
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   582
      Caption         =   "FrmDataMapping.frx":0EEC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmDataMapping.frx":0F58
      AlignHorizontal =   0
      AlignVertical   =   1
      Appearance      =   0
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   0
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   3210477
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
      Left            =   16230
      TabIndex        =   24
      Top             =   2640
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   582
      Caption         =   "FrmDataMapping.frx":0F9A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmDataMapping.frx":1006
      AlignHorizontal =   0
      AlignVertical   =   1
      Appearance      =   0
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   0
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   3210477
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
      Left            =   15675
      TabIndex        =   25
      Top             =   2310
      Width           =   540
      _Version        =   65536
      _ExtentX        =   952
      _ExtentY        =   582
      Caption         =   "FrmDataMapping.frx":1048
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmDataMapping.frx":10B4
      AlignHorizontal =   0
      AlignVertical   =   1
      Appearance      =   0
      AllowSpace      =   1
      AutoConvert     =   1
      BackColor       =   0
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   3210477
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
      Left            =   15675
      TabIndex        =   26
      Top             =   2655
      Width           =   540
      _Version        =   65536
      _ExtentX        =   952
      _ExtentY        =   582
      Caption         =   "FrmDataMapping.frx":10F6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmDataMapping.frx":1162
      AlignHorizontal =   0
      AlignVertical   =   1
      Appearance      =   0
      AllowSpace      =   1
      AutoConvert     =   1
      BackColor       =   0
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   3210477
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
      Index           =   2
      Left            =   15675
      TabIndex        =   27
      Top             =   1620
      Width           =   540
      _Version        =   65536
      _ExtentX        =   952
      _ExtentY        =   582
      Caption         =   "FrmDataMapping.frx":11A4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmDataMapping.frx":1210
      AlignHorizontal =   0
      AlignVertical   =   1
      Appearance      =   0
      AllowSpace      =   1
      AutoConvert     =   1
      BackColor       =   0
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   3210477
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
      Left            =   15675
      TabIndex        =   28
      Top             =   1965
      Width           =   540
      _Version        =   65536
      _ExtentX        =   952
      _ExtentY        =   582
      Caption         =   "FrmDataMapping.frx":1252
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmDataMapping.frx":12BE
      AlignHorizontal =   0
      AlignVertical   =   1
      Appearance      =   0
      AllowSpace      =   1
      AutoConvert     =   1
      BackColor       =   0
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   3210477
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
   Begin TDBMask6Ctl.TDBMask AHomeAdd1 
      Height          =   330
      Index           =   0
      Left            =   15675
      TabIndex        =   29
      Top             =   930
      Width           =   540
      _Version        =   65536
      _ExtentX        =   952
      _ExtentY        =   582
      Caption         =   "FrmDataMapping.frx":1300
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmDataMapping.frx":136C
      AlignHorizontal =   0
      AlignVertical   =   1
      Appearance      =   0
      AllowSpace      =   1
      AutoConvert     =   1
      BackColor       =   0
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   3210477
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
      Left            =   15675
      TabIndex        =   30
      Top             =   1275
      Width           =   540
      _Version        =   65536
      _ExtentX        =   952
      _ExtentY        =   582
      Caption         =   "FrmDataMapping.frx":13AE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmDataMapping.frx":141A
      AlignHorizontal =   0
      AlignVertical   =   1
      Appearance      =   0
      AllowSpace      =   1
      AutoConvert     =   1
      BackColor       =   0
      BorderStyle     =   1
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   3210477
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
   Begin TDBMask6Ctl.TDBMask TxtExt3 
      Height          =   330
      Left            =   18090
      TabIndex        =   31
      Top             =   1620
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   582
      Caption         =   "FrmDataMapping.frx":145C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmDataMapping.frx":14C8
      AlignHorizontal =   0
      AlignVertical   =   1
      Appearance      =   0
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   0
      BorderStyle     =   0
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   3210477
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
      Left            =   18090
      TabIndex        =   32
      Top             =   1965
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   582
      Caption         =   "FrmDataMapping.frx":150A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmDataMapping.frx":1576
      AlignHorizontal =   0
      AlignVertical   =   1
      Appearance      =   0
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   0
      BorderStyle     =   0
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   3210477
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
   Begin TDBNumber6Ctl.TDBNumber lblLimit 
      Height          =   240
      Left            =   8850
      TabIndex        =   33
      Top             =   1185
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   423
      Calculator      =   "FrmDataMapping.frx":15B8
      Caption         =   "FrmDataMapping.frx":15D8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmDataMapping.frx":1644
      Keys            =   "FrmDataMapping.frx":1662
      Spin            =   "FrmDataMapping.frx":16AC
      AlignHorizontal =   1
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
      ForeColor       =   8454016
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
      Left            =   8850
      TabIndex        =   34
      Top             =   2295
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   450
      Calculator      =   "FrmDataMapping.frx":16D4
      Caption         =   "FrmDataMapping.frx":16F4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmDataMapping.frx":1760
      Keys            =   "FrmDataMapping.frx":177E
      Spin            =   "FrmDataMapping.frx":17C8
      AlignHorizontal =   1
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
      ForeColor       =   8454016
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
   Begin TDBNumber6Ctl.TDBNumber lblTtlPay 
      Height          =   255
      Left            =   8850
      TabIndex        =   35
      Top             =   2010
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   450
      Calculator      =   "FrmDataMapping.frx":17F0
      Caption         =   "FrmDataMapping.frx":1810
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmDataMapping.frx":187C
      Keys            =   "FrmDataMapping.frx":189A
      Spin            =   "FrmDataMapping.frx":18E4
      AlignHorizontal =   1
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
      ForeColor       =   8454016
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
   Begin TDBNumber6Ctl.TDBNumber lblLastPay 
      Height          =   240
      Left            =   8850
      TabIndex        =   36
      Top             =   1740
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   423
      Calculator      =   "FrmDataMapping.frx":190C
      Caption         =   "FrmDataMapping.frx":192C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmDataMapping.frx":1998
      Keys            =   "FrmDataMapping.frx":19B6
      Spin            =   "FrmDataMapping.frx":1A00
      AlignHorizontal =   1
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
      ForeColor       =   8454016
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
      Height          =   240
      Left            =   8850
      TabIndex        =   37
      Top             =   1455
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   423
      Calendar        =   "FrmDataMapping.frx":1A28
      Caption         =   "FrmDataMapping.frx":1B40
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmDataMapping.frx":1BAC
      Keys            =   "FrmDataMapping.frx":1BCA
      Spin            =   "FrmDataMapping.frx":1C28
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   0
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
      ForeColor       =   8454016
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
      Height          =   225
      Left            =   8850
      TabIndex        =   38
      Top             =   3105
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   397
      Calculator      =   "FrmDataMapping.frx":1C50
      Caption         =   "FrmDataMapping.frx":1C70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmDataMapping.frx":1CDC
      Keys            =   "FrmDataMapping.frx":1CFA
      Spin            =   "FrmDataMapping.frx":1D44
      AlignHorizontal =   1
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
      ForeColor       =   8454016
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
      Height          =   225
      Left            =   16380
      TabIndex        =   39
      Top             =   4065
      Visible         =   0   'False
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   397
      Calculator      =   "FrmDataMapping.frx":1D6C
      Caption         =   "FrmDataMapping.frx":1D8C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmDataMapping.frx":1DF8
      Keys            =   "FrmDataMapping.frx":1E16
      Spin            =   "FrmDataMapping.frx":1E60
      AlignHorizontal =   1
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
      ForeColor       =   8454016
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
   Begin TDBDate6Ctl.TDBDate lblBD 
      Height          =   255
      Left            =   8850
      TabIndex        =   40
      Top             =   900
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   450
      Calendar        =   "FrmDataMapping.frx":1E88
      Caption         =   "FrmDataMapping.frx":1FA0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmDataMapping.frx":200C
      Keys            =   "FrmDataMapping.frx":202A
      Spin            =   "FrmDataMapping.frx":2088
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   0
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
      ForeColor       =   8454016
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
   Begin TDBMask6Ctl.TDBMask AOffice2 
      Height          =   300
      Left            =   11925
      TabIndex        =   41
      Top             =   1965
      Width           =   525
      _Version        =   65536
      _ExtentX        =   926
      _ExtentY        =   529
      Caption         =   "FrmDataMapping.frx":20B0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "FrmDataMapping.frx":211C
      AlignHorizontal =   0
      AlignVertical   =   1
      Appearance      =   0
      AllowSpace      =   -1
      AutoConvert     =   -1
      BackColor       =   0
      BorderStyle     =   0
      ClipMode        =   0
      CursorPosition  =   -1
      DataProperty    =   0
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   3210477
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
      PromptChar      =   "_"
      ReadOnly        =   -1
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "[____]"
      Value           =   ""
   End
   Begin TDBNumber6Ctl.TDBNumber TxtAfterPay 
      Height          =   315
      Left            =   9030
      TabIndex        =   42
      Top             =   2655
      Width           =   1530
      _Version        =   65536
      _ExtentX        =   2699
      _ExtentY        =   556
      Calculator      =   "FrmDataMapping.frx":215E
      Caption         =   "FrmDataMapping.frx":217E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmDataMapping.frx":21EA
      Keys            =   "FrmDataMapping.frx":2208
      Spin            =   "FrmDataMapping.frx":2252
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   0
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   8454016
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
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   6750213
      MinValueVT      =   3538949
   End
   Begin TDBNumber6Ctl.TDBNumber LblPrincipleafterpay 
      Height          =   225
      Left            =   6360
      TabIndex        =   43
      Top             =   4590
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   397
      Calculator      =   "FrmDataMapping.frx":227A
      Caption         =   "FrmDataMapping.frx":229A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmDataMapping.frx":2306
      Keys            =   "FrmDataMapping.frx":2324
      Spin            =   "FrmDataMapping.frx":236E
      AlignHorizontal =   1
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
      ForeColor       =   8454016
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
   Begin TDBDate6Ctl.TDBDate LblLastCall 
      Height          =   300
      Left            =   4140
      TabIndex        =   138
      Top             =   5820
      Width           =   1170
      _Version        =   65536
      _ExtentX        =   2064
      _ExtentY        =   529
      Calendar        =   "FrmDataMapping.frx":2396
      Caption         =   "FrmDataMapping.frx":24AE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmDataMapping.frx":251A
      Keys            =   "FrmDataMapping.frx":2538
      Spin            =   "FrmDataMapping.frx":2596
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   0
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
      ForeColor       =   8454016
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
   Begin TDBNumber6Ctl.TDBNumber TxtLastPayment 
      Height          =   225
      Left            =   9180
      TabIndex        =   161
      Top             =   5580
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   397
      Calculator      =   "FrmDataMapping.frx":25BE
      Caption         =   "FrmDataMapping.frx":25DE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmDataMapping.frx":264A
      Keys            =   "FrmDataMapping.frx":2668
      Spin            =   "FrmDataMapping.frx":26B2
      AlignHorizontal =   1
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
      ForeColor       =   8454016
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
   Begin TDBDate6Ctl.TDBDate TxtLastPaydate 
      Height          =   285
      Left            =   9300
      TabIndex        =   163
      Top             =   6000
      Width           =   1050
      _Version        =   65536
      _ExtentX        =   1852
      _ExtentY        =   503
      Calendar        =   "FrmDataMapping.frx":26DA
      Caption         =   "FrmDataMapping.frx":27F2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmDataMapping.frx":285E
      Keys            =   "FrmDataMapping.frx":287C
      Spin            =   "FrmDataMapping.frx":28DA
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   0
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
      ForeColor       =   8454016
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
   Begin VB.Label label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Paydate:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   15
      Left            =   7965
      TabIndex        =   162
      Top             =   6060
      Width           =   1065
   End
   Begin VB.Label label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Payment:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   13
      Left            =   7920
      TabIndex        =   160
      Top             =   5520
      Width           =   1110
   End
   Begin VB.Label TxtMobileAdd2 
      BackColor       =   &H00000000&
      Caption         =   "-"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   15720
      TabIndex        =   159
      Top             =   3420
      Width           =   1455
   End
   Begin VB.Label TxtMobileAdd1 
      BackColor       =   &H00000000&
      Caption         =   "-"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   15720
      TabIndex        =   158
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label TxtOfficeAdd2 
      BackColor       =   &H00000000&
      Caption         =   "-"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   16260
      TabIndex        =   157
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label TxtOfficeAdd1 
      BackColor       =   &H00000000&
      Caption         =   "-"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   16260
      TabIndex        =   156
      Top             =   1740
      Width           =   1455
   End
   Begin VB.Label TxtHomeAdd2 
      BackColor       =   &H00000000&
      Caption         =   "-"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   16260
      TabIndex        =   155
      Top             =   1380
      Width           =   1455
   End
   Begin VB.Label TxtHomeAdd1 
      BackColor       =   &H00000000&
      Caption         =   "-"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   16260
      TabIndex        =   154
      Top             =   1020
      Width           =   1455
   End
   Begin VB.Label TxtMobileno2 
      BackColor       =   &H00000000&
      Caption         =   "-"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   11940
      TabIndex        =   153
      Top             =   2700
      Width           =   1455
   End
   Begin VB.Label TxtMobileno1 
      BackColor       =   &H00000000&
      Caption         =   "-"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   11940
      TabIndex        =   152
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label TxtOfficeno2 
      BackColor       =   &H00000000&
      Caption         =   "-"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   12480
      TabIndex        =   151
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label TxtOfficeno1 
      BackColor       =   &H00000000&
      Caption         =   "-"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   12480
      TabIndex        =   150
      Top             =   1740
      Width           =   1455
   End
   Begin VB.Label TxtHomeno2 
      BackColor       =   &H00000000&
      Caption         =   "-"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   12480
      TabIndex        =   149
      Top             =   1380
      Width           =   1455
   End
   Begin VB.Label txtHomeNo1 
      BackColor       =   &H00000000&
      Caption         =   "-"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   12480
      TabIndex        =   148
      Top             =   1020
      Width           =   1455
   End
   Begin VB.Label label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   18
      Left            =   10920
      TabIndex        =   147
      Top             =   4920
      Width           =   750
   End
   Begin VB.Label LblContactto 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
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
      ForeColor       =   &H0080FF80&
      Height          =   195
      Left            =   6660
      TabIndex        =   144
      Top             =   5880
      Width           =   180
   End
   Begin VB.Label label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact to:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   11
      Left            =   5640
      TabIndex        =   143
      Top             =   5820
      Width           =   900
   End
   Begin VB.Label LblPlace 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
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
      ForeColor       =   &H0080FF80&
      Height          =   195
      Left            =   6660
      TabIndex        =   142
      Top             =   5580
      Width           =   180
   End
   Begin VB.Label label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Place:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   9
      Left            =   6060
      TabIndex        =   141
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label LblActivity 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
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
      ForeColor       =   &H0080FF80&
      Height          =   195
      Left            =   4140
      TabIndex        =   140
      Top             =   6120
      Width           =   180
   End
   Begin VB.Label label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Activity:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   2
      Left            =   3360
      TabIndex        =   139
      Top             =   6060
      Width           =   690
   End
   Begin VB.Label label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tgl. Last Call:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   1
      Left            =   3000
      TabIndex        =   137
      Top             =   5820
      Width           =   1050
   End
   Begin VB.Label LblFcek 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
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
      ForeColor       =   &H0080FF80&
      Height          =   195
      Left            =   4140
      TabIndex        =   136
      Top             =   5520
      Width           =   180
   End
   Begin VB.Label label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status Account:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   0
      Left            =   2760
      TabIndex        =   135
      Top             =   5520
      Width           =   1290
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Phone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   15300
      TabIndex        =   79
      Top             =   1020
      Width           =   3075
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   5
      Left            =   13260
      TabIndex        =   134
      Top             =   480
      Width           =   2595
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   4
      Left            =   10800
      Picture         =   "FrmDataMapping.frx":2902
      Stretch         =   -1  'True
      Top             =   420
      Width           =   8070
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Nasabah"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   8520
      TabIndex        =   133
      Top             =   30
      Width           =   2325
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   600
      TabIndex        =   132
      Top             =   450
      Width           =   2325
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   0
      Left            =   60
      Picture         =   "FrmDataMapping.frx":383C
      Stretch         =   -1  'True
      Top             =   450
      Width           =   420
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000A&
      Height          =   4050
      Index           =   1
      Left            =   0
      Top             =   840
      Width           =   10620
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   120
      TabIndex        =   131
      Top             =   1695
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   120
      TabIndex        =   130
      Top             =   1380
      Width           =   825
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   120
      TabIndex        =   129
      Top             =   1125
      Width           =   825
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   120
      TabIndex        =   128
      Top             =   870
      Width           =   960
   End
   Begin VB.Label lblPriority 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   210
      Left            =   1110
      TabIndex        =   127
      Top             =   1710
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label LblDOB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   210
      Left            =   1110
      TabIndex        =   126
      Top             =   1380
      Width           =   75
   End
   Begin VB.Label lblID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   210
      Left            =   1110
      TabIndex        =   125
      Top             =   1140
      Width           =   75
   End
   Begin VB.Label lblNama 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   210
      Left            =   1140
      TabIndex        =   124
      Top             =   915
      Width           =   75
   End
   Begin VB.Label lblCardNo 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   210
      Left            =   6090
      TabIndex        =   123
      Top             =   4290
      Width           =   75
   End
   Begin VB.Label CustId 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   2
      Left            =   2580
      TabIndex        =   122
      Top             =   1710
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label LblRiskLevel 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   210
      Left            =   3465
      TabIndex        =   121
      Top             =   1710
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   435
      Left            =   120
      TabIndex        =   120
      Top             =   3420
      Width           =   810
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   555
      Left            =   120
      TabIndex        =   119
      Top             =   2715
      Width           =   825
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   450
      Left            =   135
      TabIndex        =   118
      Top             =   2055
      Width           =   825
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   120
      TabIndex        =   117
      Top             =   3990
      Width           =   825
   End
   Begin VB.Label lblZIP 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "- "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   210
      Left            =   1110
      TabIndex        =   116
      Top             =   4020
      Width           =   135
   End
   Begin VB.Label lblLMED 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   210
      Left            =   4305
      TabIndex        =   115
      Top             =   4050
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label32 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Last Month End Dlg:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   2265
      TabIndex        =   114
      Top             =   4050
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.Label LblInterest 
      BackColor       =   &H00000000&
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   240
      Left            =   6090
      TabIndex        =   113
      Top             =   1515
      Width           =   1530
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Interest:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   4635
      TabIndex        =   112
      Top             =   1515
      Width           =   690
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Principle:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   4635
      TabIndex        =   111
      Top             =   1200
      Width           =   780
   End
   Begin VB.Label lblNoPay 
      BackColor       =   &H00000000&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   6090
      TabIndex        =   110
      Top             =   900
      Width           =   1530
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#Pay:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   4635
      TabIndex        =   109
      Top             =   900
      Width           =   435
   End
   Begin VB.Label LblFees 
      BackColor       =   &H00000000&
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   240
      Left            =   6060
      TabIndex        =   108
      Top             =   1770
      Width           =   1530
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fees:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   4635
      TabIndex        =   107
      Top             =   1770
      Width           =   435
   End
   Begin VB.Label lblBrokenPromised 
      BackColor       =   &H00000000&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   6090
      TabIndex        =   106
      Top             =   3030
      Width           =   1530
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Broken Promise:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   4635
      TabIndex        =   105
      Top             =   3030
      Width           =   1335
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lc atmp:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   0
      Left            =   4635
      TabIndex        =   104
      Top             =   2700
      Width           =   690
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Bill:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   4635
      TabIndex        =   103
      Top             =   2370
      Width           =   660
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Open Date:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   4635
      TabIndex        =   102
      Top             =   2055
      Width           =   900
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "Installment :"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   1
      Left            =   4635
      TabIndex        =   101
      Top             =   3930
      Width           =   1185
   End
   Begin VB.Label lbllama 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   210
      Left            =   6090
      TabIndex        =   100
      Top             =   3330
      Width           =   150
   End
   Begin VB.Label Label36 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tgl BD-OP:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   4635
      TabIndex        =   99
      Top             =   3330
      Width           =   825
   End
   Begin VB.Label lblLIP 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   210
      Left            =   6090
      TabIndex        =   98
      Top             =   3615
      Width           =   75
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   0
      Left            =   4635
      TabIndex        =   97
      Top             =   3615
      Width           =   1545
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "Range:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   0
      Left            =   4635
      TabIndex        =   96
      Top             =   4260
      Width           =   660
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Index           =   2
      Left            =   7740
      TabIndex        =   95
      Top             =   1455
      Width           =   930
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   300
      Index           =   4
      Left            =   7740
      TabIndex        =   94
      Top             =   1740
      Width           =   930
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Index           =   5
      Left            =   7740
      TabIndex        =   93
      Top             =   2010
      Width           =   930
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   3
      Left            =   7740
      TabIndex        =   92
      Top             =   1185
      Width           =   930
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   1
      Left            =   7740
      TabIndex        =   91
      Top             =   900
      Width           =   930
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Principle Afterpay:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   480
      Index           =   8
      Left            =   15300
      TabIndex        =   90
      Top             =   4065
      Visible         =   0   'False
      Width           =   1155
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000A&
      Height          =   3990
      Index           =   2
      Left            =   10830
      Top             =   870
      Width           =   3930
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   47
      Left            =   10905
      TabIndex        =   89
      Top             =   945
      Width           =   630
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   51
      Left            =   10905
      TabIndex        =   88
      Top             =   1290
      Width           =   675
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   46
      Left            =   10905
      TabIndex        =   87
      Top             =   1620
      Width           =   630
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   44
      Left            =   10905
      TabIndex        =   86
      Top             =   1965
      Width           =   675
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   52
      Left            =   10905
      TabIndex        =   85
      Top             =   2295
      Width           =   300
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   55
      Left            =   10905
      TabIndex        =   84
      Top             =   2625
      Width           =   345
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   480
      Left            =   10920
      TabIndex        =   83
      Top             =   2940
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   10905
      TabIndex        =   82
      Top             =   3885
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   330
      Left            =   10905
      TabIndex        =   81
      Top             =   4215
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Data Econ"
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
      Index           =   10
      Left            =   10905
      TabIndex        =   80
      Top             =   3600
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label label1 
      BackColor       =   &H00E8BE91&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   3
      Left            =   14940
      TabIndex        =   78
      Top             =   930
      Width           =   1170
   End
   Begin VB.Label label1 
      BackColor       =   &H00E8BE91&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   4
      Left            =   14940
      TabIndex        =   77
      Top             =   1275
      Width           =   1170
   End
   Begin VB.Label label1 
      BackColor       =   &H00E8BE91&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   5
      Left            =   14940
      TabIndex        =   76
      Top             =   1635
      Width           =   1170
   End
   Begin VB.Label label1 
      BackColor       =   &H00E8BE91&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   6
      Left            =   14940
      TabIndex        =   75
      Top             =   1965
      Width           =   1170
   End
   Begin VB.Label label1 
      BackColor       =   &H00E8BE91&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   7
      Left            =   14940
      TabIndex        =   74
      Top             =   3030
      Width           =   1170
   End
   Begin VB.Label label1 
      BackColor       =   &H00E8BE91&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   8
      Left            =   14940
      TabIndex        =   73
      Top             =   3405
      Width           =   1170
   End
   Begin VB.Label label1 
      BackColor       =   &H00E8BE91&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   56
      Left            =   14940
      TabIndex        =   72
      Top             =   2655
      Width           =   1170
   End
   Begin VB.Label label1 
      BackColor       =   &H00E8BE91&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Index           =   59
      Left            =   14940
      TabIndex        =   71
      Top             =   2310
      Width           =   1170
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment History"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   570
      TabIndex        =   70
      Top             =   8505
      Width           =   3075
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "Other and Call"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   6
      Left            =   660
      TabIndex        =   69
      Top             =   4980
      Width           =   3075
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   5
      Left            =   180
      Picture         =   "FrmDataMapping.frx":4346
      Stretch         =   -1  'True
      Top             =   4980
      Width           =   420
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000A&
      Height          =   930
      Index           =   6
      Left            =   60
      Top             =   5430
      Width           =   10500
   End
   Begin VB.Label lblRecsource 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
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
      ForeColor       =   &H0080FF80&
      Height          =   195
      Left            =   1320
      TabIndex        =   68
      Top             =   5730
      Width           =   180
   End
   Begin VB.Label label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   80
      Left            =   240
      TabIndex        =   67
      Top             =   5730
      Width           =   945
   End
   Begin VB.Label label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   65
      Left            =   435
      TabIndex        =   66
      Top             =   5490
      Width           =   750
   End
   Begin VB.Label lblCustId 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
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
      ForeColor       =   &H0080FF80&
      Height          =   195
      Left            =   1320
      TabIndex        =   65
      Top             =   5520
      Width           =   180
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000A&
      X1              =   7680
      X2              =   7680
      Y1              =   840
      Y2              =   4890
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000A&
      X1              =   4560
      X2              =   4560
      Y1              =   870
      Y2              =   4890
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "Tgl FollowUp:"
      Height          =   255
      Left            =   8610
      TabIndex        =   64
      Top             =   10200
      Width           =   975
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Status Telp:"
      Height          =   255
      Index           =   2
      Left            =   9000
      TabIndex        =   63
      Top             =   10230
      Width           =   885
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Note:"
      Height          =   255
      Index           =   1
      Left            =   8760
      TabIndex        =   62
      Top             =   10845
      Width           =   450
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Zip Office:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   210
      Index           =   9
      Left            =   120
      TabIndex        =   61
      Top             =   4620
      Width           =   870
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblzipoffice 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "- "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   210
      Left            =   1110
      TabIndex        =   60
      Top             =   4620
      Width           =   135
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Tgl.Source:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   10
      Left            =   7740
      TabIndex        =   59
      Top             =   3600
      Width           =   1140
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "No Payment:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   11
      Left            =   7740
      TabIndex        =   58
      Top             =   3945
      Width           =   1320
      WordWrap        =   -1  'True
   End
   Begin VB.Label LBLBUCKET 
      BackStyle       =   0  'Transparent
      Caption         =   "Bucket:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   7740
      TabIndex        =   57
      Top             =   4245
      Width           =   960
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Days"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   13
      Left            =   9735
      TabIndex        =   56
      Top             =   3945
      Width           =   570
      WordWrap        =   -1  'True
   End
   Begin VB.Label label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type Acc:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   12
      Left            =   315
      TabIndex        =   55
      Top             =   5970
      Width           =   855
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Region:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   210
      Index           =   12
      Left            =   120
      TabIndex        =   54
      Top             =   4320
      Width           =   930
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblRegion2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "- "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   210
      Left            =   1110
      TabIndex        =   53
      Top             =   4320
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000A&
      Height          =   3990
      Index           =   3
      Left            =   14850
      Top             =   870
      Visible         =   0   'False
      Width           =   4050
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   6
      Left            =   7740
      TabIndex        =   52
      Top             =   2295
      Width           =   1050
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount After Pay:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   480
      Index           =   7
      Left            =   7740
      TabIndex        =   51
      Top             =   3105
      Width           =   930
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Paid:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   450
      Left            =   7740
      TabIndex        =   50
      Top             =   2655
      Width           =   960
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "Principle After Pay:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   2
      Left            =   4620
      TabIndex        =   49
      Top             =   4590
      Width           =   1605
   End
   Begin VB.Label Label41 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Visa:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   1980
      TabIndex        =   48
      Top             =   4320
      Width           =   420
   End
   Begin VB.Label LblVisa 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   210
      Left            =   2460
      TabIndex        =   47
      Top             =   4320
      Width           =   75
   End
   Begin VB.Label Label42 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Master:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   1980
      TabIndex        =   46
      Top             =   4620
      Width           =   600
   End
   Begin VB.Label LblMaster 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   210
      Left            =   2700
      TabIndex        =   45
      Top             =   4620
      Width           =   75
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Map:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   7740
      TabIndex        =   44
      Top             =   4530
      Width           =   960
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   0
      Left            =   0
      Picture         =   "FrmDataMapping.frx":4E50
      Stretch         =   -1  'True
      Top             =   390
      Width           =   10620
   End
   Begin VB.Image Image2 
      Height          =   405
      Index           =   8
      Left            =   0
      Picture         =   "FrmDataMapping.frx":5D8A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20700
   End
   Begin VB.Image Image2 
      Height          =   465
      Index           =   5
      Left            =   60
      Picture         =   "FrmDataMapping.frx":6CC4
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   10560
   End
End
Attribute VB_Name = "FrmDataMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub IsiData()
    Dim M_COL As ADODB.Recordset
    Dim CMDSQL As String
        
    CMDSQL = "select * from mgm_mapping_pil where custidcard='"
    CMDSQL = CMDSQL + Trim(FrmCC_Colection.lblcustid.Caption) + "'"

    Set M_COL = New ADODB.Recordset
    M_COL.CursorLocation = adUseClient
    M_COL.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If Not M_COL.EOF Then
        
        '@@02-05-2011 Tambahan field last payment dan last paydate
        TxtLastPayment.Value = IIf(IsNull(M_COL("last_payment")), "0", M_COL("last_payment"))
        TxtLastPaydate.Value = IIf(IsNull(M_COL("last_paydate")), "", Format(M_COL("last_paydate"), "dd/mm/yyyy"))
        
        LblPrincipleafterpay.Value = IIf(IsNull(M_COL("principle_after_pay")), "0", M_COL("principle_after_pay"))
            
        LblVisa.Caption = IIf(IsNull(M_COL("visa")), "-", M_COL("visa"))
        LblMaster.Caption = IIf(IsNull(M_COL("master")), "-", M_COL("master"))
        
        
        TxtMap.Text = IIf(IsNull(M_COL("map")), "0", M_COL("map"))
        
        lblcustid.Caption = IIf(IsNull(M_COL("CUSTID")), "", M_COL("CUSTID"))
        LblInterest.Caption = Format(IIf(IsNull(M_COL("INTEREST")), "0", M_COL("INTEREST")), "##,###")
        LblFees.Caption = Format(IIf(IsNull(M_COL("FEES")), "0", M_COL("FEES")), "##,###")
        lblzipoffice.Caption = IIf(IsNull(M_COL("zipoffice")), "", M_COL("zipoffice"))
        lblRecsource.Caption = IIf(IsNull(M_COL("RECSOURCE")), "", M_COL("RECSOURCE"))
        LblRiskLevel.Caption = IIf(IsNull(M_COL("RiskLevel")), "", M_COL("RiskLevel"))
        lblPriority.Caption = IIf(IsNull(M_COL("Priority")), "", M_COL("Priority"))
        lblnama.Caption = IIf(IsNull(M_COL("NAME")), "", M_COL("NAME"))
        lblCardNo.Caption = IIf(IsNull(M_COL("NoCard")), "", M_COL("NoCard"))
        lblID.Caption = IIf(IsNull(M_COL("ktpno")), "", M_COL("ktpno"))
        statusptp2 = IIf(IsNull(M_COL!F_CEK), "", M_COL!F_CEK)
        txtTglSource.Text = IIf(IsNull(M_COL!TGLSOURCE), "", M_COL!TGLSOURCE)
        StsAcc.Text = IIf(IsNull(M_COL!stspurge), "", M_COL!stspurge)
        
        If Trim(StsAcc.Text) = "PURGE" Then
            StsAcc.ForeColor = vbRed
        Else
            StsAcc.ForeColor = vbGreen
        End If
        
        TxtTipeAccount.Text = IIf(IsNull(M_COL!Type), "", M_COL!Type)
        
        
        LblDOB.Caption = IIf(IsNull(M_COL("DOB")), "", M_COL("DOB"))
        lblAddr.Text = IIf(IsNull(M_COL("ADDRNOW")), "", M_COL("ADDRNOW"))
        lblOfficeAddr.Text = IIf(IsNull(M_COL("ADDRPT")), "", M_COL("ADDRPT"))
        LblCHAditionalAddr.Text = IIf(IsNull(M_COL("addressadd")), "", M_COL("addressadd"))
        lblZIP.Caption = IIf(IsNull(M_COL("ZIPNOW")), "", M_COL("ZIPNOW"))
        lblNoPay.Caption = IIf(IsNull(M_COL("NoPay")), "", M_COL("NoPay"))
        LblPrompA.Value = IIf(IsNull(M_COL("Principal")), "", M_COL("Principal"))
        lblOpenDate.Value = IIf(IsNull(M_COL("OpenDate")), "", M_COL("OpenDate"))
        If (lblBD.ValueIsNull) Or (lblOpenDate.ValueIsNull) Then
        Else
            hsl = DateDiff("m", M_COL("OpenDate"), M_COL("b_D"))
            lbllama.Caption = IIf(IsNull(CStr(hsl)), "", CStr(hsl)) + "  Bulan "
        End If
        lblLastBill.Value = IIf(IsNull(M_COL("LastBill")), "", M_COL("LastBill"))
        lblLcAtm.Value = IIf(IsNull(M_COL("LcATMP")), "", M_COL("LcATMP"))
        lblBrokenPromised.Caption = IIf(IsNull(M_COL("BrokenPromise")), "", M_COL("BrokenPromise"))
        lblBD.Value = CStr(Format(IIf(IsNull(M_COL("B_D")), "", M_COL("B_D")), "yyyy/mm/dd"))
        lblLimit.Value = IIf(IsNull(M_COL("Limit")), "", M_COL("Limit"))
        lblPayDt.Value = IIf(IsNull(M_COL("Pay_Dt")), "", M_COL("Pay_Dt"))
        lblLastPay.Value = IIf(IsNull(M_COL("LastPay")), "", M_COL("LastPay"))
        lblTtlPay.Value = IIf(IsNull(M_COL("TtlPay")), "", M_COL("TtlPay"))
        lblAmount.Value = IIf(IsNull(M_COL("AmountWo")), "", Format(M_COL("AmountWo"), "##.##0"))
        lblRegion2.Caption = IIf(IsNull(M_COL("region")), "", M_COL("region"))
        lblLMED.Caption = IIf(IsNull(M_COL("LMED")), "", M_COL("LMED"))
        lblLIP.Caption = IIf(IsNull(M_COL("LIP")), "", Format(M_COL("LIP"), "dd/mm/yyyy"))
        LblInstallment.Value = IIf(IsNull(M_COL("Installment")), "", Format(M_COL("Installment"), "##.##0"))
        
        AHome1.Value = IIf(IsNull(M_COL("AHOMENO")), "", M_COL("AHOMENO"))
        txtHomeNo1.Caption = IIf(IsNull(M_COL("HOMENO")), "", M_COL("HOMENO"))
        
        AHome2.Value = IIf(IsNull(M_COL("AHOMENO2")), "", M_COL("AHOMENO2"))
        txtHomeNo2.Caption = IIf(IsNull(M_COL("HOMENO2")), "", M_COL("HOMENO2"))
        
        AOffice1.Value = IIf(IsNull(M_COL("AOFFICENO")), "", M_COL("AOFFICENO"))
        txtOfficeNo1.Caption = IIf(IsNull(M_COL("OFFICENO")), "", M_COL("OFFICENO"))
        
        AOffice2.Value = IIf(IsNull(M_COL("AOFFICENO2")), "", M_COL("AOFFICENO2"))
        txtOfficeNo2.Caption = IIf(IsNull(M_COL("OFFICENO2")), "", M_COL("OFFICENO2"))
        
        txtMobileNo1.Caption = IIf(IsNull(M_COL("MOBILENO")), "", M_COL("MOBILENO"))
        
        txtMobileNo2.Caption = IIf(IsNull(M_COL("MOBILENO2")), "", M_COL("MOBILENO2"))
        
   
        AHomeAdd1(0).Value = IIf(IsNull(M_COL("AHOMENOADD1")), "", M_COL("AHOMENOADD1"))
        AHomeAdd2(1).Value = IIf(IsNull(M_COL("AHOMENOADD2")), "", M_COL("AHOMENOADD2"))
        AOfficeAdd(2).Value = IIf(IsNull(M_COL("AOFFICENOADD1")), "", M_COL("AOFFICENOADD1"))
        AOfficeAdd(3).Value = IIf(IsNull(M_COL("AOFFICENOADD2")), "", M_COL("AOFFICENOADD2"))
        AFaxAdd(4).Value = IIf(IsNull(M_COL("AFAXNOADD1")), "", M_COL("AFAXNOADD1"))
        AFaxAdd(5).Value = IIf(IsNull(M_COL("AFAXNOADD2")), "", M_COL("AFAXNOADD2"))
        
        txtHomeAdd1.Caption = IIf(IsNull(M_COL("HOMENOADD1")), "", M_COL("HOMENOADD1"))
        
        txtHomeAdd2.Caption = IIf(IsNull(M_COL("HOMENOADD2")), "", M_COL("HOMENOADD2"))
        
        txtOfficeAdd1.Caption = IIf(IsNull(M_COL("OFFICENOADD1")), "", M_COL("OFFICENOADD1"))
    
        txtOfficeAdd2.Caption = IIf(IsNull(M_COL("OFFICENOADD2")), "", M_COL("OFFICENOADD2"))
        
        txtMobileAdd1.Caption = IIf(IsNull(M_COL("MOBILENOADD1")), "", M_COL("MOBILENOADD1"))
        
        txtMobileAdd2.Caption = IIf(IsNull(M_COL("MOBILENOADD2")), "", M_COL("MOBILENOADD2"))
        
        txtFaxAdd1.Value = IIf(IsNull(M_COL("FAXNOADD1")), "", M_COL("FAXNOADD1"))
        txtFaxAdd2.Value = IIf(IsNull(M_COL("FAXNOADD2")), "", M_COL("FAXNOADD2"))
        AddrNow.Text = IIf(IsNull(M_COL("TxtPtpAddr")), "", M_COL("TxtPtpAddr"))
        TxtEC.Text = IIf(IsNull(M_COL!ec_name), "", M_COL!ec_name)
        txtECno.Value = IIf(IsNull(M_COL!ec_telp), "", M_COL!ec_telp)
       
        'cbolastcall.Text = IIf(IsNull(M_COL!statuscall), "", M_COL!statuscall)
    
    
        'cari extension
'        If InStr(1, txtOfficeNo1.Value, "X", vbTextCompare) > 0 Then
'            TxtExt1.Text = Right(txtOfficeNo1.Value, Len(txtOfficeNo1.Value) - InStr(1, txtOfficeNo1.Value, "X", vbTextCompare))
'        End If
'        If InStr(1, txtOfficeNo2.Value, "X", vbTextCompare) > 0 Then
'            TxtExt2.Text = Right(txtOfficeNo2.Value, Len(txtOfficeNo2.Value) - InStr(1, txtOfficeNo2.Value, "X", vbTextCompare))
'        End If
'        If InStr(1, txtOfficeAdd1.Value, "X", vbTextCompare) > 0 Then
'            TxtExt3.Text = Right(txtOfficeAdd1.Value, Len(txtOfficeAdd1.Value) - InStr(1, txtOfficeAdd1.Value, "X", vbTextCompare))
'        End If
'        If InStr(1, txtOfficeAdd2.Value, "X", vbTextCompare) > 0 Then
'            TxtExt4.Text = Right(txtOfficeAdd2.Value, Len(txtOfficeAdd2.Value) - InStr(1, txtOfficeAdd2.Value, "X", vbTextCompare))
'        End If
        
        'If UCase(mdiform1.txtlevel.text) = "AGENT" Then
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
'            If Len(txtECno.Value) > 2 Then
'                txtECno.ReadOnly = True
'            End If
'        'End If
        LblLastCall.Value = IIf(IsNull(M_COL("tglcall")), "", Format(M_COL("tglcall"), "dd/mm/yyyy"))
        LblFcek.Caption = IIf(IsNull(M_COL("f_cek")), "", Left(M_COL!F_CEK, 2))
        
        LblActivity.Caption = IIf(IsNull(M_COL!ACTIVITY), "", M_COL!ACTIVITY)
        LblPlace.Caption = IIf(IsNull(M_COL!place), "", M_COL!place)
        LblContactto.Caption = IIf(IsNull(M_COL!contactto), "", M_COL!contactto)
        TxtRemarks.Text = IIf(IsNull(M_COL!Remarks), "", M_COL!Remarks)
        
   End If
   Set M_COL = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.Hide
    End If
    
End Sub

Private Sub Form_Load()
    Call IsiData
End Sub
