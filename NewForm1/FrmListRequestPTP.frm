VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmListRequestPTP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List Request PTP"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   13245
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   13245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6795
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13155
      _ExtentX        =   23204
      _ExtentY        =   11986
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   8421504
      TabCaption(0)   =   "List Request PTP"
      TabPicture(0)   =   "FrmListRequestPTP.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "CD_save"
      Tab(0).Control(2)=   "PanelExport"
      Tab(0).Control(3)=   "CmdExport"
      Tab(0).Control(4)=   "CmbApprove"
      Tab(0).Control(5)=   "CmbTampilkan"
      Tab(0).Control(6)=   "CmdCari"
      Tab(0).Control(7)=   "TxtNama"
      Tab(0).Control(8)=   "TxtCustid"
      Tab(0).Control(9)=   "PB1"
      Tab(0).Control(10)=   "TxtLPDPayment"
      Tab(0).Control(11)=   "CmdRefresh"
      Tab(0).Control(12)=   "CmdReject"
      Tab(0).Control(13)=   "CmdApprove"
      Tab(0).Control(14)=   "CmdUnCekAll"
      Tab(0).Control(15)=   "CmdCekall"
      Tab(0).Control(16)=   "TxtJml"
      Tab(0).Control(17)=   "LvPTP"
      Tab(0).Control(18)=   "TxtLPAPayment"
      Tab(0).Control(19)=   "Label2(1)"
      Tab(0).Control(20)=   "Shape1"
      Tab(0).Control(21)=   "Label5"
      Tab(0).Control(22)=   "Label4"
      Tab(0).Control(23)=   "Line1"
      Tab(0).Control(24)=   "Label3"
      Tab(0).Control(25)=   "Label2(0)"
      Tab(0).Control(26)=   "Label1"
      Tab(0).ControlCount=   27
      TabCaption(1)   =   "SEND PTP REJECTED"
      TabPicture(1)   =   "FrmListRequestPTP.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label7"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Line2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label8"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label9"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "LvPTPRejected"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "TxtJmlDataRejected"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "CmbJenisRejected"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "CmdCariRejected"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "TxtNamaRejected"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "TxtCustidRejected"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "CmbKembalikan"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "SEND PTP APPROVED"
      TabPicture(2)   =   "FrmListRequestPTP.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdsudahemail"
      Tab(2).Control(1)=   "cmrefresh"
      Tab(2).Control(2)=   "cmdsearch"
      Tab(2).Control(3)=   "txtsearch"
      Tab(2).Control(4)=   "cbsearch"
      Tab(2).Control(5)=   "TxtJmlApproved"
      Tab(2).Control(6)=   "LvPTPApproved"
      Tab(2).Control(7)=   "date1"
      Tab(2).Control(8)=   "date2"
      Tab(2).Control(9)=   "Shape2"
      Tab(2).Control(10)=   "Label2(2)"
      Tab(2).Control(11)=   "Label17"
      Tab(2).Control(12)=   "Label16"
      Tab(2).Control(13)=   "Label10"
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "Approve By Pak Hamanto"
      TabPicture(3)   =   "FrmListRequestPTP.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "CmdUnCekAllHamanto"
      Tab(3).Control(1)=   "CmdCekAllHamanto"
      Tab(3).Control(2)=   "CmdApproveHamanto"
      Tab(3).Control(3)=   "TxtJmlhAppHamanto"
      Tab(3).Control(4)=   "TxtCustidHamanto"
      Tab(3).Control(5)=   "TxtCariNamaHamanto"
      Tab(3).Control(6)=   "CmdCariAppHamanto"
      Tab(3).Control(7)=   "LvHamanto"
      Tab(3).Control(8)=   "PB2"
      Tab(3).Control(9)=   "TxtTglApprove"
      Tab(3).Control(10)=   "Label14"
      Tab(3).Control(11)=   "Label13"
      Tab(3).Control(12)=   "Label12"
      Tab(3).Control(13)=   "Label11"
      Tab(3).ControlCount=   14
      Begin VB.Frame Frame1 
         Caption         =   "SAMPAH"
         Height          =   1695
         Left            =   -72240
         TabIndex        =   68
         Top             =   1920
         Visible         =   0   'False
         Width           =   1815
         Begin VB.CommandButton cmd_SID 
            Caption         =   "Add To SID"
            Height          =   375
            Left            =   600
            TabIndex        =   71
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton CmdApproveByPTP 
            Caption         =   "&Approve PTP DISC. By SPV"
            Height          =   795
            Left            =   480
            TabIndex        =   70
            Top             =   900
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CommandButton CmdApproveVP 
            Caption         =   "To Be Approve By Pak Hamanto"
            Height          =   615
            Left            =   480
            TabIndex        =   69
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdsudahemail 
         Caption         =   "Already email"
         Height          =   375
         Left            =   -63120
         TabIndex        =   66
         Top             =   5880
         Width           =   1095
      End
      Begin VB.CommandButton cmrefresh 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   -70200
         TabIndex        =   65
         Top             =   6360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdsearch 
         Caption         =   "Search"
         Height          =   375
         Left            =   -71280
         TabIndex        =   64
         Top             =   6360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtsearch 
         Height          =   285
         Left            =   -69480
         TabIndex        =   59
         Top             =   5940
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox cbsearch 
         Height          =   315
         ItemData        =   "FrmListRequestPTP.frx":0070
         Left            =   -71280
         List            =   "FrmListRequestPTP.frx":0080
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   5940
         Width           =   1695
      End
      Begin MSComDlg.CommonDialog CD_save 
         Left            =   -74520
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Threed.SSPanel PanelExport 
         Height          =   1335
         Left            =   -67320
         TabIndex        =   50
         Top             =   1680
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2355
         _Version        =   196610
         ActiveColors    =   -1  'True
         BevelInner      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox Txtlocation 
            Height          =   285
            Left            =   1800
            TabIndex        =   55
            Top             =   1080
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Exit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2280
            TabIndex        =   54
            Top             =   0
            Width           =   615
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Export"
            Height          =   375
            Left            =   1800
            TabIndex        =   53
            Top             =   480
            Width           =   735
         End
         Begin TDBDate6Ctl.TDBDate TdbDateExport 
            Height          =   285
            Left            =   240
            TabIndex        =   52
            Top             =   600
            Width           =   1485
            _Version        =   65536
            _ExtentX        =   2619
            _ExtentY        =   494
            Calendar        =   "FrmListRequestPTP.frx":00C3
            Caption         =   "FrmListRequestPTP.frx":01DB
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FrmListRequestPTP.frx":0247
            Keys            =   "FrmListRequestPTP.frx":0265
            Spin            =   "FrmListRequestPTP.frx":02C3
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
            BackColor       =   12648384
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
         Begin VB.Label Label15 
            Caption         =   "Tanggal Approve"
            Height          =   255
            Left            =   240
            TabIndex        =   51
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.CommandButton CmdExport 
         Caption         =   "Export to Excel"
         Height          =   375
         Left            =   -63720
         TabIndex        =   49
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton CmdUnCekAllHamanto 
         Caption         =   "&UnCek All"
         Height          =   315
         Left            =   -67080
         TabIndex        =   47
         Top             =   420
         Width           =   1395
      End
      Begin VB.CommandButton CmdCekAllHamanto 
         Caption         =   "&Cek All"
         Height          =   315
         Left            =   -68460
         TabIndex        =   48
         Top             =   420
         Width           =   1395
      End
      Begin VB.CommandButton CmdApproveHamanto 
         Caption         =   "&Approve"
         Height          =   435
         Left            =   -63720
         TabIndex        =   44
         Top             =   840
         Width           =   1755
      End
      Begin VB.TextBox TxtJmlhAppHamanto 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73860
         TabIndex        =   42
         Text            =   "0"
         Top             =   6000
         Width           =   1215
      End
      Begin VB.TextBox TxtCustidHamanto 
         Height          =   285
         Left            =   -74160
         TabIndex        =   37
         Top             =   480
         Width           =   2115
      End
      Begin VB.TextBox TxtCariNamaHamanto 
         Height          =   285
         Left            =   -71280
         TabIndex        =   36
         Top             =   480
         Width           =   1635
      End
      Begin VB.CommandButton CmdCariAppHamanto 
         Caption         =   "&Cari"
         Height          =   315
         Left            =   -69600
         TabIndex        =   35
         Top             =   420
         Width           =   915
      End
      Begin VB.TextBox TxtJmlApproved 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73860
         TabIndex        =   33
         Text            =   "0"
         Top             =   5940
         Width           =   1215
      End
      Begin VB.CommandButton CmbKembalikan 
         Caption         =   "&Kembalikan ke list request PTP"
         Height          =   435
         Left            =   9960
         TabIndex        =   31
         Top             =   720
         Width           =   2955
      End
      Begin VB.TextBox TxtCustidRejected 
         Height          =   285
         Left            =   780
         TabIndex        =   27
         Top             =   900
         Width           =   2115
      End
      Begin VB.TextBox TxtNamaRejected 
         Height          =   285
         Left            =   3660
         TabIndex        =   26
         Top             =   900
         Width           =   1635
      End
      Begin VB.CommandButton CmdCariRejected 
         Caption         =   "&Cari"
         Height          =   315
         Left            =   5340
         TabIndex        =   25
         Top             =   840
         Width           =   915
      End
      Begin VB.ComboBox CmbJenisRejected 
         Height          =   315
         ItemData        =   "FrmListRequestPTP.frx":02EB
         Left            =   7920
         List            =   "FrmListRequestPTP.frx":02F5
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox TxtJmlDataRejected 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   1140
         TabIndex        =   22
         Text            =   "0"
         Top             =   5940
         Width           =   1215
      End
      Begin VB.ComboBox CmbApprove 
         Height          =   315
         ItemData        =   "FrmListRequestPTP.frx":0312
         Left            =   -63720
         List            =   "FrmListRequestPTP.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   4440
         Width           =   1755
      End
      Begin VB.ComboBox CmbTampilkan 
         Height          =   315
         ItemData        =   "FrmListRequestPTP.frx":0316
         Left            =   -67200
         List            =   "FrmListRequestPTP.frx":0320
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   780
         Width           =   3315
      End
      Begin VB.CommandButton CmdCari 
         Caption         =   "&Cari"
         Height          =   315
         Left            =   -69660
         TabIndex        =   16
         Top             =   780
         Width           =   915
      End
      Begin VB.TextBox TxtNama 
         Height          =   285
         Left            =   -71340
         TabIndex        =   15
         Top             =   840
         Width           =   1635
      End
      Begin VB.TextBox TxtCustid 
         Height          =   285
         Left            =   -74220
         TabIndex        =   13
         Top             =   840
         Width           =   2115
      End
      Begin MSComctlLib.ProgressBar PB1 
         Height          =   315
         Left            =   -72480
         TabIndex        =   11
         Top             =   5940
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.TextBox TxtLPDPayment 
         Appearance      =   0  'Flat
         BackColor       =   &H009AD6C2&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   -65880
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   5160
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   -63720
         TabIndex        =   8
         Top             =   1200
         Width           =   1635
      End
      Begin VB.CommandButton CmdReject 
         Caption         =   "Reject"
         Height          =   375
         Left            =   -63720
         TabIndex        =   7
         Top             =   5400
         Width           =   1635
      End
      Begin VB.CommandButton CmdApprove 
         Caption         =   "&Approve"
         Height          =   375
         Left            =   -63720
         TabIndex        =   6
         Top             =   4920
         Width           =   1635
      End
      Begin VB.CommandButton CmdUnCekAll 
         Caption         =   "&UnCek All"
         Height          =   375
         Left            =   -63720
         TabIndex        =   5
         Top             =   2040
         Width           =   1635
      End
      Begin VB.CommandButton CmdCekall 
         Caption         =   "&Cek All"
         Height          =   375
         Left            =   -63720
         TabIndex        =   4
         Top             =   1680
         Width           =   1635
      End
      Begin VB.TextBox TxtJml 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   -73860
         TabIndex        =   3
         Text            =   "0"
         Top             =   5940
         Width           =   1215
      End
      Begin MSComctlLib.ListView LvPTP 
         Height          =   4620
         Left            =   -74940
         TabIndex        =   1
         ToolTipText     =   "Double click untuk melihat detail CPA"
         Top             =   1200
         Width           =   11100
         _ExtentX        =   19579
         _ExtentY        =   8149
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
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
      Begin TDBNumber6Ctl.TDBNumber TxtLPAPayment 
         Height          =   255
         Left            =   -67260
         TabIndex        =   10
         Top             =   4380
         Visible         =   0   'False
         Width           =   1380
         _Version        =   65536
         _ExtentX        =   2434
         _ExtentY        =   450
         Calculator      =   "FrmListRequestPTP.frx":033D
         Caption         =   "FrmListRequestPTP.frx":035D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmListRequestPTP.frx":03C9
         Keys            =   "FrmListRequestPTP.frx":03E7
         Spin            =   "FrmListRequestPTP.frx":0431
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   10147522
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0"
         EditMode        =   0
         Enabled         =   0
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
      Begin MSComctlLib.ListView LvPTPRejected 
         Height          =   4620
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "Double click untuk melihat detail CPA"
         Top             =   1320
         Width           =   12840
         _ExtentX        =   22648
         _ExtentY        =   8149
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   8438015
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
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
      Begin MSComctlLib.ListView LvPTPApproved 
         Height          =   5160
         Left            =   -74940
         TabIndex        =   32
         ToolTipText     =   "Double click untuk melihat detail CPA"
         Top             =   720
         Width           =   12960
         _ExtentX        =   22860
         _ExtentY        =   9102
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   8454016
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
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
      Begin MSComctlLib.ListView LvHamanto 
         Height          =   5040
         Left            =   -74880
         TabIndex        =   38
         ToolTipText     =   "Double click untuk melihat detail CPA"
         Top             =   840
         Width           =   11100
         _ExtentX        =   19579
         _ExtentY        =   8890
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   1
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
      Begin MSComctlLib.ProgressBar PB2 
         Height          =   315
         Left            =   -72480
         TabIndex        =   41
         Top             =   6000
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin TDBDate6Ctl.TDBDate TxtTglApprove 
         Height          =   285
         Left            =   -63600
         TabIndex        =   45
         Top             =   1560
         Width           =   1485
         _Version        =   65536
         _ExtentX        =   2619
         _ExtentY        =   494
         Calendar        =   "FrmListRequestPTP.frx":0459
         Caption         =   "FrmListRequestPTP.frx":0571
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmListRequestPTP.frx":05DD
         Keys            =   "FrmListRequestPTP.frx":05FB
         Spin            =   "FrmListRequestPTP.frx":0659
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   12648384
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
      Begin TDBDate6Ctl.TDBDate date1 
         Height          =   285
         Left            =   -69480
         TabIndex        =   61
         Top             =   5940
         Visible         =   0   'False
         Width           =   1485
         _Version        =   65536
         _ExtentX        =   2619
         _ExtentY        =   503
         Calendar        =   "FrmListRequestPTP.frx":0681
         Caption         =   "FrmListRequestPTP.frx":0799
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmListRequestPTP.frx":0805
         Keys            =   "FrmListRequestPTP.frx":0823
         Spin            =   "FrmListRequestPTP.frx":0881
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   12648384
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
      Begin TDBDate6Ctl.TDBDate date2 
         Height          =   285
         Left            =   -67560
         TabIndex        =   63
         Top             =   5940
         Visible         =   0   'False
         Width           =   1485
         _Version        =   65536
         _ExtentX        =   2619
         _ExtentY        =   494
         Calendar        =   "FrmListRequestPTP.frx":08A9
         Caption         =   "FrmListRequestPTP.frx":09C1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FrmListRequestPTP.frx":0A2D
         Keys            =   "FrmListRequestPTP.frx":0A4B
         Spin            =   "FrmListRequestPTP.frx":0AA9
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   12648384
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
      Begin VB.Shape Shape2 
         BackColor       =   &H000000FF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   -69000
         Top             =   6360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Belum di Email"
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
         Index           =   2
         Left            =   -68640
         TabIndex        =   67
         Top             =   6360
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label17 
         Caption         =   "To"
         Height          =   255
         Left            =   -67920
         TabIndex        =   60
         Top             =   5940
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label16 
         Caption         =   "Search By :"
         Height          =   255
         Left            =   -72240
         TabIndex        =   58
         Top             =   5940
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Belum di print"
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
         Index           =   1
         Left            =   -63360
         TabIndex        =   56
         Top             =   6000
         Width           =   1275
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000000FF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   -63720
         Top             =   6000
         Width           =   255
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Tanggal Approve:"
         Height          =   195
         Left            =   -63720
         TabIndex        =   46
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label13 
         Caption         =   "Jumlah Data:"
         Height          =   315
         Left            =   -74880
         TabIndex        =   43
         Top             =   6000
         Width           =   1035
      End
      Begin VB.Label Label12 
         Caption         =   "Custid:"
         Height          =   195
         Left            =   -74820
         TabIndex        =   40
         Top             =   540
         Width           =   675
      End
      Begin VB.Label Label11 
         Caption         =   "Nama:"
         Height          =   195
         Left            =   -71940
         TabIndex        =   39
         Top             =   540
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Jumlah Data:"
         Height          =   315
         Left            =   -74880
         TabIndex        =   34
         Top             =   5940
         Width           =   1035
      End
      Begin VB.Label Label9 
         Caption         =   "Custid:"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label8 
         Caption         =   "Nama:"
         Height          =   195
         Left            =   3000
         TabIndex        =   29
         Top             =   960
         Width           =   615
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   6420
         X2              =   6420
         Y1              =   720
         Y2              =   1260
      End
      Begin VB.Label Label7 
         Caption         =   "Tampilkan hanya:"
         Height          =   195
         Left            =   6540
         TabIndex        =   28
         Top             =   900
         Width           =   1275
      End
      Begin VB.Label Label6 
         Caption         =   "Jumlah Data:"
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   5940
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Approve By:"
         Height          =   195
         Left            =   -63720
         TabIndex        =   19
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Tampilkan hanya:"
         Height          =   195
         Left            =   -68580
         TabIndex        =   17
         Top             =   840
         Width           =   1275
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   -68640
         X2              =   -68640
         Y1              =   660
         Y2              =   1200
      End
      Begin VB.Label Label3 
         Caption         =   "Nama:"
         Height          =   195
         Left            =   -72000
         TabIndex        =   14
         Top             =   900
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Custid:"
         Height          =   195
         Index           =   0
         Left            =   -74880
         TabIndex        =   12
         Top             =   900
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Jumlah Data:"
         Height          =   315
         Left            =   -74880
         TabIndex        =   2
         Top             =   5940
         Width           =   1035
      End
   End
   Begin TDBDate6Ctl.TDBDate TDBDate3 
      Height          =   285
      Left            =   11400
      TabIndex        =   62
      Top             =   0
      Width           =   1485
      _Version        =   65536
      _ExtentX        =   2619
      _ExtentY        =   494
      Calendar        =   "FrmListRequestPTP.frx":0AD1
      Caption         =   "FrmListRequestPTP.frx":0BE9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmListRequestPTP.frx":0C55
      Keys            =   "FrmListRequestPTP.frx":0C73
      Spin            =   "FrmListRequestPTP.frx":0CD1
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   12648384
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
   Begin VB.Menu OH 
      Caption         =   ""
   End
End
Attribute VB_Name = "FrmListRequestPTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StatusPTP As String
Dim PaymentTenor As Double


Private Sub HeaderLog()
    LvPTP.ColumnHeaders.CLEAR
    With LvPTP.ColumnHeaders
        .ADD 1, , "ID", 500
        .ADD 2, , "Jenis PTP", 1000
        .ADD 3, , "Custid", 2000
        .ADD 4, , "Nama CH", 3000
        .ADD 5, , "Status", 2000
        .ADD 6, , "Tanggal Approve", 2000
        .ADD 7, , "Tgl.Payment Effective", 2500
        .ADD 8, , "Total Amount", 1000
        .ADD 9, , "Tenor", 700
        .ADD 10, , "Pembayaran Via", 2000
        .ADD 11, , "Tgl.Tagih", 1500
        .ADD 12, , "Principal", 1000
        .ADD 13, , "Balance", 1000
        .ADD 14, , "Pembayaran Awal", 2000
        .ADD 15, , "Principal", 2000
        .ADD 16, , "Total Payment", 2000
        .ADD 17, , "Down Payment", 2000
        .ADD 18, , "Charge", 2000
        .ADD 19, , "Discount", 2000
        .ADD 20, , "From o/s balance %", 2000
        .ADD 21, , "Principal %", 2000
        .ADD 22, , "Justtification", 2000
        .ADD 23, , "Fax", 800
        .ADD 24, , "When Talking Surlun", 800
        .ADD 25, , "KTP", 800
        .ADD 26, , "Surper", 800
        .ADD 27, , "Billing", 800
        .ADD 28, , "Other", 800
        .ADD 29, , "Agent", 800
        .ADD 30, , "DOB", 1000
        .ADD 31, , "Ket.Other", 1000
        
        '@@ 16-07-2012 Tambahan Payment Handle
        .ADD 32, , "Payment Handle", 2000
        
        '@@17-07-2012 Tambahan Occupation dan Reason
        .ADD 33, , "Occupation", 2000
        .ADD 34, , "Reason", 2000
    End With
End Sub


Private Sub HeaderLogRejected()
    LvPTPRejected.ColumnHeaders.CLEAR
    With LvPTPRejected.ColumnHeaders
        .ADD 1, , "ID", 500
        .ADD 2, , "Jenis PTP", 1000
        .ADD 3, , "Custid", 2000
        .ADD 4, , "Nama CH", 3000
        .ADD 5, , "Status", 2000
        .ADD 6, , "Tanggal Approve", 2000
        .ADD 7, , "Tgl.Payment Effective", 2500
        .ADD 8, , "Total Amount", 1000
        .ADD 9, , "Tenor", 700
        .ADD 10, , "Pembayaran Via", 2000
        .ADD 11, , "Tgl.Tagih", 1500
        .ADD 12, , "Principal", 1000
        .ADD 13, , "Balance", 1000
        .ADD 14, , "Pembayaran Awal", 2000
        .ADD 15, , "Principal", 2000
        .ADD 16, , "Total Payment", 2000
        .ADD 17, , "Down Payment", 2000
        .ADD 18, , "Charge", 2000
        .ADD 19, , "Discount", 2000
        .ADD 20, , "From o/s balance %", 2000
        .ADD 21, , "Principal %", 2000
        .ADD 22, , "Justtification", 2000
        .ADD 23, , "Fax", 800
        .ADD 24, , "When Talking Surlun", 800
        .ADD 25, , "KTP", 800
        .ADD 26, , "Surper", 800
        .ADD 27, , "Billing", 800
        .ADD 28, , "Other", 800
        .ADD 29, , "Agent", 800
        .ADD 30, , "DOB", 1000
        .ADD 31, , "Ket.Other", 1000
    End With
End Sub

Private Sub HeaderLogApproved()
'    LvPTPApproved.ColumnHeaders.CLEAR
'    With LvPTPApproved.ColumnHeaders
'        .ADD 1, , "ID", 500
'        .ADD 2, , "Jenis PTP", 1000
'        .ADD 3, , "Custid", 2000
'        .ADD 4, , "Nama CH", 3000
'        .ADD 5, , "Status", 2000
'        .ADD 6, , "Tanggal Approve", 2000
'        .ADD 7, , "Tgl.Payment Effective", 2500
'        .ADD 8, , "Total Amount", 1000
'        .ADD 9, , "Tenor", 700
'        .ADD 10, , "Pembayaran Via", 2000
'        .ADD 11, , "Tgl.Tagih", 1500
'        .ADD 12, , "Principal", 1000
'        .ADD 13, , "Balance", 1000
'        .ADD 14, , "Pembayaran Awal", 2000
'        .ADD 15, , "Principal", 2000
'        .ADD 16, , "Total Payment", 2000
'        .ADD 17, , "Down Payment", 2000
'        .ADD 18, , "Charge", 2000
'        .ADD 19, , "Discount", 2000
'        .ADD 20, , "From o/s balance %", 2000
'        .ADD 21, , "Principal %", 2000
'        .ADD 22, , "Justtification", 2000
'        .ADD 23, , "Fax", 800
'        .ADD 24, , "When Talking Surlun", 800
'        .ADD 25, , "KTP", 800
'        .ADD 26, , "Surper", 800
'        .ADD 27, , "Billing", 800
'        .ADD 28, , "Other", 800
'        .ADD 29, , "Agent", 800
'        .ADD 30, , "DOB", 1000
'        .ADD 31, , "Ket.Other", 1000
'        .ADD 32, , "Level Approval", 2000
'    End With

'jejaktian
    LvPTPApproved.ColumnHeaders.CLEAR
    With LvPTPApproved.ColumnHeaders
        .ADD 1, , "ID", 500
        .ADD 2, , "Jenis PTP", 1000
        .ADD 3, , "Custid", 2000
        .ADD 4, , "Nama CH", 3000
        .ADD 5, , "Status", 2000
        .ADD 6, , "Tanggal Approve", 2000
        .ADD 7, , "Level Approval", 2000
        .ADD 8, , "Tanggal Request by Email", 2000
    End With

End Sub


Public Sub IsiLog()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim ListItem As ListItem
    
    Call HeaderLog
    
    CMDSQL = "select * from tblsendptp where id is not null "
    If UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Then
        CMDSQL = CMDSQL + " and agent in "
        CMDSQL = CMDSQL + "(select userid from usertbl where team='"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "' and usertype='1') "
        
        'CMDSQL = CMDSQL + " and jenis_ptp='PTP No Discount' "
    End If
    If txtnama.text <> Empty Then
        CMDSQL = CMDSQL + " and vcustname like '%"
        CMDSQL = CMDSQL + txtnama.text + "%' "
    End If
    If txtcustid.text <> Empty Then
        CMDSQL = CMDSQL + " and custid like '%"
        CMDSQL = CMDSQL + txtcustid.text + "%' "
    End If
    
    If CmbTampilkan.text = "PTP NO DISC." Then
       CMDSQL = CMDSQL + " and jenis_ptp='PTP No Discount' "
       CMDSQL = CMDSQL + " and status='0' "
    End If
    If CmbTampilkan.text = "PTP DISC." Then
        CMDSQL = CMDSQL + " and jenis_ptp='PTP Discount' "
        CMDSQL = CMDSQL + " and status='0'"
    End If
    If CmbTampilkan.text = "PTP DISC. APPROVED" Then
        CMDSQL = CMDSQL + " and jenis_ptp='PTP Discount' "
        CMDSQL = CMDSQL + " and status='1' "
    End If
    
    
   '@@221012 Tambahan buat approve by VP
    CMDSQL = CMDSQL + " and  sts_app_vp is null "
    
    CMDSQL = CMDSQL + " order by tgldata desc"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvPTP.ListItems.CLEAR
    txtjml.text = M_objrs.RecordCount
    
    If M_objrs.RecordCount > 0 Then
        Dim STATUS As String
        Dim DOB As String
        While Not M_objrs.EOF
            'On Error Resume Next
            Set ListItem = LvPTP.ListItems.ADD(, , M_objrs("id"))
                ListItem.SubItems(1) = IIf(IsNull(M_objrs("jenis_ptp")), "", M_objrs("jenis_ptp"))
                ListItem.SubItems(2) = IIf(IsNull(M_objrs("custid")), "", M_objrs("custid"))
                ListItem.SubItems(3) = IIf(IsNull(M_objrs("vcustname")), "", M_objrs("vcustname"))
                
                If M_objrs("status") = "0" Then
                    STATUS = "Belum di Approve"
                End If
                If M_objrs("status") = "1" Then
                    STATUS = "Approve"
                End If
                If M_objrs("status") = "2" Then
                    STATUS = "Rejected"
                End If
                
                ListItem.SubItems(4) = STATUS
                ListItem.SubItems(5) = IIf(IsNull(M_objrs("tgl_approve")), "", Format(M_objrs("tgl_approve"), "yyyy-mm-dd"))
                ListItem.SubItems(6) = IIf(IsNull(M_objrs("date_payment_effective")), "", Format(M_objrs("date_payment_effective"), "yyyy-mm-dd"))
                ListItem.SubItems(7) = IIf(IsNull(M_objrs("total_amount_deal")), "0", M_objrs("total_amount_deal"))
                ListItem.SubItems(8) = IIf(IsNull(M_objrs("tenor")), "1", M_objrs("tenor"))
                ListItem.SubItems(9) = IIf(IsNull(M_objrs("pembayaran_via")), "", M_objrs("pembayaran_via"))
                ListItem.SubItems(10) = IIf(IsNull(M_objrs("tgl_tagih")), "", Format(M_objrs("tgl_tagih"), "yyyy-mm-dd"))
                ListItem.SubItems(11) = IIf(IsNull(M_objrs("principal")), "0", M_objrs("principal"))
                ListItem.SubItems(12) = IIf(IsNull(M_objrs("balance")), "0", M_objrs("balance"))
                ListItem.SubItems(13) = IIf(IsNull(M_objrs("pembayaran_awal")), "0", M_objrs("pembayaran_awal"))
                ListItem.SubItems(14) = IIf(IsNull(M_objrs("principal")), "0", M_objrs("principal"))
                ListItem.SubItems(15) = IIf(IsNull(M_objrs("nttlpayment")), "0", M_objrs("nttlpayment"))
                ListItem.SubItems(16) = IIf(IsNull(M_objrs("ndownpay")), "0", M_objrs("ndownpay"))
                ListItem.SubItems(17) = IIf(IsNull(M_objrs("ncharge")), "0", M_objrs("ncharge"))
                ListItem.SubItems(18) = IIf(IsNull(M_objrs("ndiscountamt")), "0", M_objrs("ndiscountamt"))
                ListItem.SubItems(19) = IIf(IsNull(M_objrs("vosbalance")), "", M_objrs("vosbalance"))
                ListItem.SubItems(20) = IIf(IsNull(M_objrs("vosprincipal")), "", M_objrs("vosprincipal"))
                ListItem.SubItems(21) = IIf(IsNull(M_objrs("vjust")), "", M_objrs("vjust"))
                ListItem.SubItems(22) = IIf(IsNull(M_objrs("chkfaxed")), "", M_objrs("chkfaxed"))
                ListItem.SubItems(23) = IIf(IsNull(M_objrs("chkwentalking")), "", M_objrs("chkwentalking"))
                ListItem.SubItems(24) = IIf(IsNull(M_objrs("chkktp")), "", M_objrs("chkktp"))
                ListItem.SubItems(25) = IIf(IsNull(M_objrs("chksup")), "", M_objrs("chksup"))
                ListItem.SubItems(26) = IIf(IsNull(M_objrs("chkbillings")), "", M_objrs("chkbillings"))
                ListItem.SubItems(27) = IIf(IsNull(M_objrs("chkothers")), "", M_objrs("chkothers"))
                ListItem.SubItems(28) = IIf(IsNull(M_objrs("agent")), "", M_objrs("agent"))
                  
                If IsNull(M_objrs("dob")) = True Or M_objrs("dob") = "" Or M_objrs("dob") = Empty Then
                    DOB = ""
                Else
                    DOB = Format(M_objrs("dob"), "yyyy-mm-dd")
                End If
                 
                ListItem.SubItems(29) = DOB
                ListItem.SubItems(30) = IIf(IsNull(M_objrs("ket_other")), "", M_objrs("ket_other"))
                ListItem.SubItems(31) = IIf(IsNull(M_objrs("payment_handle")), "", M_objrs("payment_handle"))
                
                ListItem.SubItems(32) = IIf(IsNull(M_objrs("occupation")), "", M_objrs("occupation"))
                ListItem.SubItems(33) = IIf(IsNull(M_objrs("reason")), "", M_objrs("reason"))
                'listItem.SubItems(34) = IIf(IsNull(M_Objrs("tgl_send_email")), "", M_Objrs("tgl_send_email"))

                ' Tandain klo ini belum di print 13 Okt 2014
                If IIf(IsNull(M_objrs("s_print")), 0, M_objrs("s_print")) = 0 Then
                    For K = 1 To 7
                        ListItem.ListSubItems(K).ForeColor = vbRed
                    Next K
                End If
                ' ------------------------------------------
            M_objrs.MoveNext
        Wend
    End If
    Set M_objrs = Nothing
End Sub


Public Sub IsiLogRejected()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim ListItem As ListItem
    
    CMDSQL = "select * from tblsendptp_log_reject where id is not null "
    If UCase(MDIForm1.txtlevel.text) = "TEAMLEADER" Then
        CMDSQL = CMDSQL + " and agent in "
        CMDSQL = CMDSQL + "(select userid from usertbl where team='"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "' and usertype='1' and aktif='0') "
        
        'CMDSQL = CMDSQL + " and jenis_ptp='PTP No Discount' "
    End If
    If TxtNamaRejected.text <> Empty Then
        CMDSQL = CMDSQL + " and vcustname like '%"
        CMDSQL = CMDSQL + TxtNamaRejected.text + "%' "
    End If
    If TxtCustidRejected.text <> Empty Then
        CMDSQL = CMDSQL + " and custid like '%"
        CMDSQL = CMDSQL + TxtCustidRejected.text + "%' "
    End If
    
    If CmbJenisRejected.text = "PTP NO DISC." Then
       CMDSQL = CMDSQL + " and jenis_ptp='PTP No Discount' "
       CMDSQL = CMDSQL + " and status='0' "
    End If
    If CmbJenisRejected.text = "PTP DISC." Then
        CMDSQL = CMDSQL + " and jenis_ptp='PTP Discount' "
        CMDSQL = CMDSQL + " and status='0'"
    End If
    If CmbJenisRejected.text = "PTP DISC. APPROVED" Then
        CMDSQL = CMDSQL + " and jenis_ptp='PTP Discount' "
        CMDSQL = CMDSQL + " and status='1' "
    End If
    
    
    
    CMDSQL = CMDSQL + " order by tgldata desc limit 300 "
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvPTPRejected.ListItems.CLEAR
    TxtJmlDataRejected.text = M_objrs.RecordCount
    
    If M_objrs.RecordCount > 0 Then
        Dim STATUS As String
        Dim DOB As String
        While Not M_objrs.EOF
            'On Error Resume Next
            Set ListItem = LvPTPRejected.ListItems.ADD(, , M_objrs("id"))
                ListItem.SubItems(1) = IIf(IsNull(M_objrs("jenis_ptp")), "", M_objrs("jenis_ptp"))
                ListItem.SubItems(2) = IIf(IsNull(M_objrs("custid")), "", M_objrs("custid"))
                ListItem.SubItems(3) = IIf(IsNull(M_objrs("vcustname")), "", M_objrs("vcustname"))
                
                If M_objrs("status") = "0" Then
                    STATUS = "Belum di Approve"
                End If
                If M_objrs("status") = "1" Then
                    STATUS = "Approve"
                End If
                If M_objrs("status") = "2" Then
                    STATUS = "Rejected"
                End If
                
                ListItem.SubItems(4) = STATUS
                ListItem.SubItems(5) = IIf(IsNull(M_objrs("tgl_approve")), "", Format(M_objrs("tgl_approve"), "yyyy-mm-dd"))
                ListItem.SubItems(6) = IIf(IsNull(M_objrs("date_payment_effective")), "", Format(M_objrs("date_payment_effective"), "yyyy-mm-dd"))
                ListItem.SubItems(7) = IIf(IsNull(M_objrs("total_amount_deal")), "0", M_objrs("total_amount_deal"))
                ListItem.SubItems(8) = IIf(IsNull(M_objrs("tenor")), "1", M_objrs("tenor"))
                ListItem.SubItems(9) = IIf(IsNull(M_objrs("pembayaran_via")), "", M_objrs("pembayaran_via"))
                ListItem.SubItems(10) = IIf(IsNull(M_objrs("tgl_tagih")), "", Format(M_objrs("tgl_tagih"), "yyyy-mm-dd"))
                ListItem.SubItems(11) = IIf(IsNull(M_objrs("principal")), "0", M_objrs("principal"))
                ListItem.SubItems(12) = IIf(IsNull(M_objrs("balance")), "0", M_objrs("balance"))
                ListItem.SubItems(13) = IIf(IsNull(M_objrs("pembayaran_awal")), "0", M_objrs("pembayaran_awal"))
                ListItem.SubItems(14) = IIf(IsNull(M_objrs("principal")), "0", M_objrs("principal"))
                ListItem.SubItems(15) = IIf(IsNull(M_objrs("nttlpayment")), "0", M_objrs("nttlpayment"))
                ListItem.SubItems(16) = IIf(IsNull(M_objrs("ndownpay")), "0", M_objrs("ndownpay"))
                ListItem.SubItems(17) = IIf(IsNull(M_objrs("ncharge")), "0", M_objrs("ncharge"))
                ListItem.SubItems(18) = IIf(IsNull(M_objrs("ndiscountamt")), "0", M_objrs("ndiscountamt"))
                ListItem.SubItems(19) = IIf(IsNull(M_objrs("vosbalance")), "", M_objrs("vosbalance"))
                ListItem.SubItems(20) = IIf(IsNull(M_objrs("vosprincipal")), "", M_objrs("vosprincipal"))
                ListItem.SubItems(21) = IIf(IsNull(M_objrs("vjust")), "", M_objrs("vjust"))
                ListItem.SubItems(22) = IIf(IsNull(M_objrs("chkfaxed")), "", M_objrs("chkfaxed"))
                ListItem.SubItems(23) = IIf(IsNull(M_objrs("chkwentalking")), "", M_objrs("chkwentalking"))
                ListItem.SubItems(24) = IIf(IsNull(M_objrs("chkktp")), "", M_objrs("chkktp"))
                ListItem.SubItems(25) = IIf(IsNull(M_objrs("chksup")), "", M_objrs("chksup"))
                ListItem.SubItems(26) = IIf(IsNull(M_objrs("chkbillings")), "", M_objrs("chkbillings"))
                ListItem.SubItems(27) = IIf(IsNull(M_objrs("chkothers")), "", M_objrs("chkothers"))
                ListItem.SubItems(28) = IIf(IsNull(M_objrs("agent")), "", M_objrs("agent"))
                  
                If IsNull(M_objrs("dob")) = True Or M_objrs("dob") = "" Or M_objrs("dob") = Empty Then
                    DOB = ""
                Else
                    DOB = Format(M_objrs("dob"), "yyyy-mm-dd")
                End If
                 
                ListItem.SubItems(29) = DOB
                ListItem.SubItems(30) = IIf(IsNull(M_objrs("ket_other")), "", M_objrs("ket_other"))
                'listItem.SubItems(31) = IIf(IsNull(M_Objrs("tgl_send_email")), "", M_Objrs("tgl_send_email"))
 
            M_objrs.MoveNext
        Wend
    End If
    Set M_objrs = Nothing
End Sub


Public Sub IsiLogApproved()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    Dim ListItem As ListItem
    
    CMDSQL = "select * from tblsendptp_log_approve where id is not null "
    If UCase(MDIForm1.txtlevel.text) = "TEAMLEADER" Then
        CMDSQL = CMDSQL + " and agent in "
        CMDSQL = CMDSQL + "(select userid from usertbl where team='"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "' and usertype='1' and aktif='0') "
        
        'CMDSQL = CMDSQL + " and jenis_ptp='PTP No Discount' "
    End If
    
    'jejaktian16032016
    '========================================
    CMDSQL = CMDSQL + " order by tgldata desc limit 300 "
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvPTPApproved.ListItems.CLEAR
    TxtJmlApproved.text = M_objrs.RecordCount
    
    If M_objrs.RecordCount > 0 Then
        Dim STATUS As String
        Dim DOB As String
        Dim discount As Double
        Dim S As String
        
        While Not M_objrs.EOF
            'On Error Resume Next
            Set ListItem = LvPTPApproved.ListItems.ADD(, , M_objrs("id"))
                ListItem.SubItems(1) = IIf(IsNull(M_objrs("jenis_ptp")), "", M_objrs("jenis_ptp"))
                ListItem.SubItems(2) = IIf(IsNull(M_objrs("custid")), "", M_objrs("custid"))
                ListItem.SubItems(3) = IIf(IsNull(M_objrs("vcustname")), "", M_objrs("vcustname"))
                
                If M_objrs("status") = "0" Then
                    STATUS = "Belum di Approve"
                End If
                If M_objrs("status") = "1" Then
                    STATUS = "Approve"
                End If
                If M_objrs("status") = "2" Then
                    STATUS = "Rejected"
                End If
            
         discount = IIf(IsNull(M_objrs("ndiscountamt")), "0", M_objrs("ndiscountamt"))
            If discount > 0 And discount <= 2000000 Then
                S = "Coll SPV"
            ElseIf discount <= 10000000 Then
                S = "Coll Band 6"
            ElseIf discount <= 20000000 Then
                S = "Coll Band 5"
            ElseIf discount <= 30000000 Then
                S = "Coll Band 4"
            ElseIf discount <= 50000000 Then
                S = "Head of Coll"
            ElseIf discount <= 100000000 Then
                S = "Head of CCC"
            ElseIf discount <= 2300000000# Then
                S = "Head of CRM"
            End If
                      
                ListItem.SubItems(4) = STATUS
                ListItem.SubItems(5) = IIf(IsNull(M_objrs("tgl_approve")), "", Format(M_objrs("tgl_approve"), "yyyy-mm-dd"))
                'jejaktianremark16032016
'                listItem.SubItems(6) = IIf(IsNull(M_Objrs("date_payment_effective")), "", Format(M_Objrs("date_payment_effective"), "yyyy-mm-dd"))
'                listItem.SubItems(7) = IIf(IsNull(M_Objrs("total_amount_deal")), "0", M_Objrs("total_amount_deal"))
'                listItem.SubItems(8) = IIf(IsNull(M_Objrs("tenor")), "1", M_Objrs("tenor"))
'                listItem.SubItems(9) = IIf(IsNull(M_Objrs("pembayaran_via")), "", M_Objrs("pembayaran_via"))
'                listItem.SubItems(10) = IIf(IsNull(M_Objrs("tgl_tagih")), "", Format(M_Objrs("tgl_tagih"), "yyyy-mm-dd"))
'                listItem.SubItems(11) = IIf(IsNull(M_Objrs("principal")), "0", M_Objrs("principal"))
'                listItem.SubItems(12) = IIf(IsNull(M_Objrs("balance")), "0", M_Objrs("balance"))
'                listItem.SubItems(13) = IIf(IsNull(M_Objrs("pembayaran_awal")), "0", M_Objrs("pembayaran_awal"))
'                listItem.SubItems(14) = IIf(IsNull(M_Objrs("principal")), "0", M_Objrs("principal"))
'                listItem.SubItems(15) = IIf(IsNull(M_Objrs("nttlpayment")), "0", M_Objrs("nttlpayment"))
'                listItem.SubItems(16) = IIf(IsNull(M_Objrs("ndownpay")), "0", M_Objrs("ndownpay"))
'                listItem.SubItems(17) = IIf(IsNull(M_Objrs("ncharge")), "0", M_Objrs("ncharge"))
'                listItem.SubItems(6) = discount

'                listItem.SubItems(19) = IIf(IsNull(M_Objrs("vosbalance")), "", M_Objrs("vosbalance"))
'                listItem.SubItems(20) = IIf(IsNull(M_Objrs("vosprincipal")), "", M_Objrs("vosprincipal"))
'                listItem.SubItems(21) = IIf(IsNull(M_Objrs("vjust")), "", M_Objrs("vjust"))
'                listItem.SubItems(22) = IIf(IsNull(M_Objrs("chkfaxed")), "", M_Objrs("chkfaxed"))
'                listItem.SubItems(23) = IIf(IsNull(M_Objrs("chkwentalking")), "", M_Objrs("chkwentalking"))
'                listItem.SubItems(24) = IIf(IsNull(M_Objrs("chkktp")), "", M_Objrs("chkktp"))
'                listItem.SubItems(25) = IIf(IsNull(M_Objrs("chksup")), "", M_Objrs("chksup"))
'                listItem.SubItems(26) = IIf(IsNull(M_Objrs("chkbillings")), "", M_Objrs("chkbillings"))
'                listItem.SubItems(27) = IIf(IsNull(M_Objrs("chkothers")), "", M_Objrs("chkothers"))
'                listItem.SubItems(28) = IIf(IsNull(M_Objrs("agent")), "", M_Objrs("agent"))
'
'                If IsNull(M_Objrs("dob")) = True Or M_Objrs("dob") = "" Or M_Objrs("dob") = Empty Then
'                    DOB = ""
'                Else
'                    DOB = Format(M_Objrs("dob"), "yyyy-mm-dd")
'                End If
'
'                listItem.SubItems(29) = DOB
'                listItem.SubItems(30) = IIf(IsNull(M_Objrs("ket_other")), "", M_Objrs("ket_other"))
                '==============================================
                ListItem.SubItems(6) = S
                ListItem.SubItems(7) = IIf(IsNull(M_objrs("tgl_send_email")), "", Format(M_objrs("tgl_send_email"), "yyyy-mm-dd"))
                
            
                ' Tandain klo ini belum di send_email 'jejaktian 14032016
                If IIf(IsNull(M_objrs("tgl_send_email")), 0, M_objrs("tgl_send_email")) = 0 Then
                    For K = 1 To 7
                        ListItem.ListSubItems(K).ForeColor = vbRed
                        ListItem.ListSubItems(K).Bold = True
                    Next K
                End If
                ' ------------------------------------------
            
            M_objrs.MoveNext
        Wend
    End If
    Set M_objrs = Nothing
End Sub


Private Sub cbsearch_Click()
Dim S As String
Dim M_objrs As ADODB.Recordset
Dim ListItem As ListItem
Dim sql As String

    If cbsearch.text = "Tanggal Request by Email" Then
        S = "Tanggal Request by Email"
        txtsearch.Visible = False
        txtsearch.text = ""
        date1.Visible = True
        date2.Visible = True
        Label17.Visible = True
        cmdsearch.Visible = True
        cmrefresh.Visible = True
        date1.text = ""
        date2.text = ""
    ElseIf cbsearch.text = "Cust ID" Then
        S = "CustId"
        txtsearch.Visible = True
        cmdsearch.Visible = True
        cmrefresh.Visible = True
        txtsearch.text = ""
        date1.Visible = False
        date2.Visible = False
        Label17.Visible = False
    ElseIf cbsearch.text = "Nama Cust" Then
        S = "vcustname"
        txtsearch.Visible = True
        cmdsearch.Visible = True
        cmrefresh.Visible = True
        txtsearch.text = ""
        date1.Visible = False
        date2.Visible = False
        Label17.Visible = False
    ElseIf cbsearch.text = "Tanggal Approve" Then
        S = "tgl_approve"
        txtsearch.Visible = False
        txtsearch.text = ""
        date1.Visible = True
        date2.Visible = True
        Label17.Visible = True
        cmdsearch.Visible = True
        cmrefresh.Visible = True
        date1.text = ""
        date2.text = ""
    End If
       
'sql = "select * from tblsendptp_log_approve where id is not null "
'sql = sql + "and '" & s & "' =  '" & txtsearch.Text & "' or '" & s & "' between '" & Format(date1.Value, "yyyy/mm/dd") & "' and '" & Format(date2.Value, "yyyy/mm/dd") & "' order by tgldata desc limit 300 "
'M_OBJCONN.Execute sql
'
'LvPTPApproved.ListItems.CLEAR
'    TxtJmlApproved.Text = M_Objrs.RecordCount
'
'    If M_Objrs.RecordCount > 0 Then
'        Dim STATUS As String
'        Dim DOB As String
'        While Not M_Objrs.EOF
'            'On Error Resume Next
'            Set listItem = LvPTPApproved.ListItems.ADD(, , M_Objrs("id"))
'                listItem.SubItems(1) = IIf(IsNull(M_Objrs("jenis_ptp")), "", M_Objrs("jenis_ptp"))
'                listItem.SubItems(2) = IIf(IsNull(M_Objrs("custid")), "", M_Objrs("custid"))
'                listItem.SubItems(3) = IIf(IsNull(M_Objrs("vcustname")), "", M_Objrs("vcustname"))
'
'                If M_Objrs("status") = "0" Then
'                    STATUS = "Belum di Approve"
'                End If
'                If M_Objrs("status") = "1" Then
'                    STATUS = "Approve"
'                End If
'                If M_Objrs("status") = "2" Then
'                    STATUS = "Rejected"
'                End If
'
'                listItem.SubItems(4) = STATUS
'                listItem.SubItems(5) = IIf(IsNull(M_Objrs("tgl_approve")), "", Format(M_Objrs("tgl_approve"), "yyyy-mm-dd"))
'                listItem.SubItems(6) = IIf(IsNull(M_Objrs("date_payment_effective")), "", Format(M_Objrs("date_payment_effective"), "yyyy-mm-dd"))
'                listItem.SubItems(7) = IIf(IsNull(M_Objrs("total_amount_deal")), "0", M_Objrs("total_amount_deal"))
'                listItem.SubItems(8) = IIf(IsNull(M_Objrs("tenor")), "1", M_Objrs("tenor"))
'                listItem.SubItems(9) = IIf(IsNull(M_Objrs("pembayaran_via")), "", M_Objrs("pembayaran_via"))
'                listItem.SubItems(10) = IIf(IsNull(M_Objrs("tgl_tagih")), "", Format(M_Objrs("tgl_tagih"), "yyyy-mm-dd"))
'                listItem.SubItems(11) = IIf(IsNull(M_Objrs("principal")), "0", M_Objrs("principal"))
'                listItem.SubItems(12) = IIf(IsNull(M_Objrs("balance")), "0", M_Objrs("balance"))
'                listItem.SubItems(13) = IIf(IsNull(M_Objrs("pembayaran_awal")), "0", M_Objrs("pembayaran_awal"))
'                listItem.SubItems(14) = IIf(IsNull(M_Objrs("principal")), "0", M_Objrs("principal"))
'                listItem.SubItems(15) = IIf(IsNull(M_Objrs("nttlpayment")), "0", M_Objrs("nttlpayment"))
'                listItem.SubItems(16) = IIf(IsNull(M_Objrs("ndownpay")), "0", M_Objrs("ndownpay"))
'                listItem.SubItems(17) = IIf(IsNull(M_Objrs("ncharge")), "0", M_Objrs("ncharge"))
'                listItem.SubItems(18) = IIf(IsNull(M_Objrs("ndiscountamt")), "0", M_Objrs("ndiscountamt"))
'                listItem.SubItems(19) = IIf(IsNull(M_Objrs("vosbalance")), "", M_Objrs("vosbalance"))
'                listItem.SubItems(20) = IIf(IsNull(M_Objrs("vosprincipal")), "", M_Objrs("vosprincipal"))
'                listItem.SubItems(21) = IIf(IsNull(M_Objrs("vjust")), "", M_Objrs("vjust"))
'                listItem.SubItems(22) = IIf(IsNull(M_Objrs("chkfaxed")), "", M_Objrs("chkfaxed"))
'                listItem.SubItems(23) = IIf(IsNull(M_Objrs("chkwentalking")), "", M_Objrs("chkwentalking"))
'                listItem.SubItems(24) = IIf(IsNull(M_Objrs("chkktp")), "", M_Objrs("chkktp"))
'                listItem.SubItems(25) = IIf(IsNull(M_Objrs("chksup")), "", M_Objrs("chksup"))
'                listItem.SubItems(26) = IIf(IsNull(M_Objrs("chkbillings")), "", M_Objrs("chkbillings"))
'                listItem.SubItems(27) = IIf(IsNull(M_Objrs("chkothers")), "", M_Objrs("chkothers"))
'                listItem.SubItems(28) = IIf(IsNull(M_Objrs("agent")), "", M_Objrs("agent"))
'
'                If IsNull(M_Objrs("dob")) = True Or M_Objrs("dob") = "" Or M_Objrs("dob") = Empty Then
'                    DOB = ""
'                Else
'                    DOB = Format(M_Objrs("dob"), "yyyy-mm-dd")
'                End If
'
'                listItem.SubItems(29) = DOB
'                listItem.SubItems(30) = IIf(IsNull(M_Objrs("ket_other")), "", M_Objrs("ket_other"))
'            M_Objrs.MoveNext
'        Wend
'    End If
'    Set M_Objrs = Nothing

End Sub

Private Sub CmbJenisRejected_Click()
    Call HeaderLogRejected
    Call IsiLogRejected
End Sub

Private Sub CmbKembalikan_Click()
    Dim CMDSQL As String
    Dim K As String
    Dim W As Integer
    
    If LvPTPRejected.ListItems.Count = 0 Then
        MsgBox "Data List PTP Rejected tidak tersedia!", vbOKOnly + vbExclamation, "Peringatan"
        Exit Sub
    End If
    
    K = MsgBox("Apakah anda yakin akan mengembalikan PTP Rejected ke List PTP Request?", vbYesNo + vbQuestion, "Konfirmasi")
    
    If K = vbNo Then
        MsgBox "Pengembalian data dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    For W = 1 To LvPTPRejected.ListItems.Count
        If LvPTPRejected.ListItems(W).Checked = True Then
            CMDSQL = "insert into tblsendptp "
            CMDSQL = CMDSQL + " select * from tblsendptp_log_reject where id='"
            CMDSQL = CMDSQL + CStr(LvPTPRejected.ListItems(W).text) + "'"
            M_OBJCONN.Execute CMDSQL
            
            CMDSQL = "delete from tblsendptp_log_reject where id='"
            CMDSQL = CMDSQL + CStr(LvPTPRejected.ListItems(W).text) + "'"
            M_OBJCONN.Execute CMDSQL
        End If
    Next W
    
    MsgBox "Data PTP Rejected berhasil dikembalikan ke list request PTP!", vbOKOnly + vbInformation, "Informasi"
    
    Call IsiLogRejected
End Sub

Private Sub CmbTampilkan_Click()
'    If CmbTampilkan.Text = "PTP DISC." Then
'        LvPTP.CheckBoxes = False
'    End If
'    If CmbTampilkan.Text = "PTP NO DISC." Then
'        LvPTP.CheckBoxes = True
'    End If

    If CmbTampilkan.text = "PTP NO DISC." Then
        'CmdApproveByPTP.Visible = False
        CmdApprove.Visible = True
    End If
    If CmbTampilkan.text = "PTP DISC." Then
        If UCase(MDIForm1.txtlevel.text) = "TEAMLEADER" Then
            'CmdApproveByPTP.Visible = False
            CmbApprove.Visible = False
        Else
             'CmdApproveByPTP.Visible = True
             'CmdApprove.Visible = False
             CmbApprove.Visible = True
        End If
        
    End If
    If CmbTampilkan.text = "PTP DISC. APPROVED" Then
        'CmdApproveByPTP.Visible = False
    End If
    
    Call HeaderLog
    Call IsiLog
End Sub

Private Sub cmd_SID_Click()
    Dim K As Integer
    Dim W As String
    Dim r As Integer
    
    If LvPTP.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Infromasi"
        Exit Sub
    End If
    
    'UPDATED BY RANDY BUAT SID -- REQ BY : JOKO
    For K = 1 To LvPTP.ListItems.Count
        PB1.Value = K
        If LvPTP.ListItems(K).Checked = True Then
            FrmSID.list1.AddItem LvPTP.ListItems(K).ListSubItems(2)
            LvPTP.ListItems(K).Checked = False
        End If
    Next K
  
    FrmSID.Show vbModal
End Sub




Private Sub CmdApprove_Click()
    Dim K As Integer
    Dim W As String
    
    If LvPTP.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Infromasi"
        Exit Sub
    End If
    
    If UCase(MDIForm1.txtlevel.text) = "TEAMLEADER" Or UCase(MDIForm1.txtlevel.text) = "MANAGER" Then
        If CmbTampilkan.text = "PTP DISC." And _
           LvPTP.SelectedItem.SubItems(4) = "Belum di Approve" Then
            MsgBox "PTP yang akan anda approve adalah PTP Discon! Untuk Meng-approve-nya harus melalui persetujuan SPV, double click data yang akan di approve kemudian Print dan ajukan ke SPV!", vbOKOnly + vbInformation, "Informasi"
            Exit Sub
        End If
     End If
    
    If CmbTampilkan.text = "PTP DISC." And _
       CmbApprove.text = "" Then
        MsgBox "Approve By, tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    
    W = MsgBox("Anda yakin akan melakukan Approve untuk membuat PTP dan CPA?", vbYesNo + vbQuestion, "Konfirmasi")
    If W = vbNo Then
        MsgBox "Pembuatan PTP dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    PB1.Max = LvPTP.ListItems.Count
    
    'UPDATED BY RANDY BUAT NYATET TGL_APPROVE KARENA SEBELUMNYA KOSONG -- REQ BY : NYOTO
    For K = 1 To LvPTP.ListItems.Count
        PB1.Value = K
        If LvPTP.ListItems(K).Checked = True Then
            CMDSQL = "update tblsendptp set tgl_approve=now() where id='"
            CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).text) + "'"
            M_OBJCONN.Execute CMDSQL
        End If
    Next K
    
    CmdApprove.Enabled = False
    For K = 1 To LvPTP.ListItems.Count
        PB1.Value = K
        If LvPTP.ListItems(K).Checked = True Then
            Call BikinCPA(K)
            DoEvents
            Call BikinPTP(K)
            DoEvents
            Call CatetLogApprove(K)
            DoEvents
            Call BikinStatusPTP(K)
            DoEvents
            Call HapusData(K)
            DoEvents
            Call KirimPesan(K)
        End If
    Next K
    Call IsiLog
    MsgBox "Approve PTP berhasil dibuat!", vbOKOnly + vbInformation, "Informasi"
    CmdApprove.Enabled = True
End Sub

Private Sub CmdApproveByPTP_Click()
    Dim CMDSQL, W As String
    Dim K As Integer
    
    If UCase(MDIForm1.txtlevel.text) = "TEAMLEADER" Or UCase(MDIForm1.txtlevel.text) = "MANAGER" Then
        MsgBox "Approve PTP Discon Hanya Boleh dilakukan oleh SPV!", vbOKOnly + vbInformation, "Informasi"
    End If
       
    
    W = MsgBox("Anda yakin akan melakukan Approve untuk membuat PTP data yang dicentang?", vbYesNo + vbQuestion, "Konfirmasi")
    If W = vbNo Then
        MsgBox "Pembuatan PTP dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If

    PB1.Max = LvPTP.ListItems.Count
    
    For K = 1 To LvPTP.ListItems.Count
        PB1.Value = K
        If LvPTP.ListItems(K).Checked = True Then
            CMDSQL = "update tblsendptp set status='1',tgl_approve=now() where id='"
            CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).text) + "'"
            M_OBJCONN.Execute CMDSQL
        End If
    Next K
    Call IsiLog
    MsgBox "Approve PTP berhasil dibuat!", vbOKOnly + vbInformation, "Informasi"
End Sub

Private Sub CmdApproveHamanto_Click()
    Dim W As String
    Dim K As Integer
    
    If LvHamanto.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Infromasi"
        Exit Sub
    End If
    
    If UCase(MDIForm1.txtlevel.text) = "TEAMLEADER" Then
        MsgBox "Anda tidak memiliki akses!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    If TxtTglApprove.ValueIsNull = True Then
        MsgBox "Tanggal Approve tidak boleh kosong!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    W = MsgBox("Anda yakin akan melakukan Approve untuk membuat PTP dan CPA?", vbYesNo + vbQuestion, "Konfirmasi")
    If W = vbNo Then
        MsgBox "Pembuatan PTP dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    PB2.Max = LvHamanto.ListItems.Count
    
    CmdApproveHamanto.Enabled = False
    TxtTglApprove.Enabled = False
    For K = 1 To LvHamanto.ListItems.Count
        PB1.Value = K
        If LvHamanto.ListItems(K).Checked = True Then
            Call BikinCPA_Hamanto(K)
            DoEvents
            Call BikinPTP_Hamanto(K)
            DoEvents
            Call CatetLogApprove_Hamanto(K)
            DoEvents
            Call BikinStatusPTP_Hamanto(K)
            DoEvents
            Call HapusData_Hamanto(K)
            DoEvents
            Call KirimPesan_Hamanto(K)
        End If
    Next K
    Call IsiLog
    MsgBox "Approve PTP berhasil dibuat!", vbOKOnly + vbInformation, "Informasi"
    CmdApproveHamanto.Enabled = True
    TxtTglApprove.Enabled = True
    CmdCariAppHamanto_Click
End Sub

Private Sub CmdApproveVP_Click()
    Dim K As Integer
    
    If LvPTP.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Infromasi"
        Exit Sub
    End If
    
    If UCase(MDIForm1.txtlevel.text) = "TEAMLEADER" Then
        If CmbTampilkan.text = "PTP DISC." And _
           LvPTP.SelectedItem.SubItems(4) = "Belum di Approve" Then
            MsgBox "PTP yang akan anda approve adalah PTP Discon! Untuk Meng-approve-nya harus melalui persetujuan SPV, double click data yang akan di approve kemudian Print dan ajukan ke SPV!", vbOKOnly + vbInformation, "Informasi"
            Exit Sub
        End If
    End If
    
    If CmbTampilkan.text <> "PTP DISC." Then
        MsgBox "Approve By VP hanya untuk PTP Disc.!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    W = MsgBox("Anda yakin akan melakukan Approve untuk membuat PTP dan CPA?", vbYesNo + vbQuestion, "Konfirmasi")
    If W = vbNo Then
        MsgBox "Pembuatan PTP dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    PB1.Max = LvPTP.ListItems.Count
    
    CmdApproveVP.Enabled = False
    For K = 1 To LvPTP.ListItems.Count
        PB1.Value = K
        If LvPTP.ListItems(K).Checked = True Then
            Call BikinCPA_AppVP(K)
            DoEvents
            'Call BikinPTP(K)
            'DoEvents
            'Call CatetLogApprove(K)
            'DoEvents
            'Call BikinStatusPTP(K)
            'DoEvents
            'Call HapusData(K)
            'DoEvents
            Call KirimPesan_AppVP(K)
        End If
    Next K
    Call IsiLog
    MsgBox "Pengajuan CPA  berhasil dibuat!", vbOKOnly + vbInformation, "Informasi"
    CmdApproveVP.Enabled = True
End Sub

Private Sub CmdCari_Click()
    Call IsiLog
End Sub

Private Sub CmdCariAppHamanto_Click()
    Dim CMDSQL As String
    Dim M_WHERE As String
    Dim M_objrs As ADODB.Recordset
    Dim ListItem As ListItem
  
    
    
    M_WHERE = ""
    
    If TxtCustidHamanto.text <> "" Then
        If M_WHERE = "" Then
            M_WHERE = " where custid like '%" + CStr(TxtCustidHamanto.text) + "%' "
        Else
            M_WHERE = M_WHERE + " and custid like '%" + CStr(TxtCustidHamanto.text) + "%' "
        End If
    End If
    
    If TxtCariNamaHamanto.text <> "" Then
        If M_WHERE = "" Then
            M_WHERE = " where vcustname like '%" + CStr(TxtCariNamaHamanto.text) + "%' "
        Else
            M_WHERE = M_WHERE + " and vcustname like '%" + CStr(TxtCariNamaHamanto.text) + "%' "
        End If
    End If
    
    If M_WHERE = "" Then
        M_WHERE = " where sts_app_vp='1' "
    Else
        M_WHERE = M_WHERE + " and sts_app_vp='1' "
    End If
    
    CMDSQL = "select * from tblsendptp " + M_WHERE
        
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LvHamanto.ListItems.CLEAR
    TxtJmlhAppHamanto.text = M_objrs.RecordCount
    
    If M_objrs.RecordCount > 0 Then
        Dim STATUS As String
        Dim DOB As String
        While Not M_objrs.EOF
            'On Error Resume Next
            Set ListItem = LvHamanto.ListItems.ADD(, , M_objrs("id"))
                ListItem.SubItems(1) = IIf(IsNull(M_objrs("jenis_ptp")), "", M_objrs("jenis_ptp"))
                ListItem.SubItems(2) = IIf(IsNull(M_objrs("custid")), "", M_objrs("custid"))
                ListItem.SubItems(3) = IIf(IsNull(M_objrs("vcustname")), "", M_objrs("vcustname"))
                
                If M_objrs("status") = "0" Then
                    STATUS = "Belum di Approve"
                End If
                If M_objrs("status") = "1" Then
                    STATUS = "Approve"
                End If
                If M_objrs("status") = "2" Then
                    STATUS = "Rejected"
                End If
                
                ListItem.SubItems(4) = STATUS
                ListItem.SubItems(5) = IIf(IsNull(M_objrs("tgl_approve")), "", Format(M_objrs("tgl_approve"), "yyyy-mm-dd"))
                ListItem.SubItems(6) = IIf(IsNull(M_objrs("date_payment_effective")), "", Format(M_objrs("date_payment_effective"), "yyyy-mm-dd"))
                ListItem.SubItems(7) = IIf(IsNull(M_objrs("total_amount_deal")), "0", M_objrs("total_amount_deal"))
                ListItem.SubItems(8) = IIf(IsNull(M_objrs("tenor")), "1", M_objrs("tenor"))
                ListItem.SubItems(9) = IIf(IsNull(M_objrs("pembayaran_via")), "", M_objrs("pembayaran_via"))
                ListItem.SubItems(10) = IIf(IsNull(M_objrs("tgl_tagih")), "", Format(M_objrs("tgl_tagih"), "yyyy-mm-dd"))
                ListItem.SubItems(11) = IIf(IsNull(M_objrs("principal")), "0", M_objrs("principal"))
                ListItem.SubItems(12) = IIf(IsNull(M_objrs("balance")), "0", M_objrs("balance"))
                ListItem.SubItems(13) = IIf(IsNull(M_objrs("pembayaran_awal")), "0", M_objrs("pembayaran_awal"))
                ListItem.SubItems(14) = IIf(IsNull(M_objrs("principal")), "0", M_objrs("principal"))
                ListItem.SubItems(15) = IIf(IsNull(M_objrs("nttlpayment")), "0", M_objrs("nttlpayment"))
                ListItem.SubItems(16) = IIf(IsNull(M_objrs("ndownpay")), "0", M_objrs("ndownpay"))
                ListItem.SubItems(17) = IIf(IsNull(M_objrs("ncharge")), "0", M_objrs("ncharge"))
                ListItem.SubItems(18) = IIf(IsNull(M_objrs("ndiscountamt")), "0", M_objrs("ndiscountamt"))
                ListItem.SubItems(19) = IIf(IsNull(M_objrs("vosbalance")), "", M_objrs("vosbalance"))
                ListItem.SubItems(20) = IIf(IsNull(M_objrs("vosprincipal")), "", M_objrs("vosprincipal"))
                ListItem.SubItems(21) = IIf(IsNull(M_objrs("vjust")), "", M_objrs("vjust"))
                ListItem.SubItems(22) = IIf(IsNull(M_objrs("chkfaxed")), "", M_objrs("chkfaxed"))
                ListItem.SubItems(23) = IIf(IsNull(M_objrs("chkwentalking")), "", M_objrs("chkwentalking"))
                ListItem.SubItems(24) = IIf(IsNull(M_objrs("chkktp")), "", M_objrs("chkktp"))
                ListItem.SubItems(25) = IIf(IsNull(M_objrs("chksup")), "", M_objrs("chksup"))
                ListItem.SubItems(26) = IIf(IsNull(M_objrs("chkbillings")), "", M_objrs("chkbillings"))
                ListItem.SubItems(27) = IIf(IsNull(M_objrs("chkothers")), "", M_objrs("chkothers"))
                ListItem.SubItems(28) = IIf(IsNull(M_objrs("agent")), "", M_objrs("agent"))
                  
                If IsNull(M_objrs("dob")) = True Or M_objrs("dob") = "" Or M_objrs("dob") = Empty Then
                    DOB = ""
                Else
                    DOB = Format(M_objrs("dob"), "yyyy-mm-dd")
                End If
                 
                ListItem.SubItems(29) = DOB
                ListItem.SubItems(30) = IIf(IsNull(M_objrs("ket_other")), "", M_objrs("ket_other"))
                ListItem.SubItems(31) = IIf(IsNull(M_objrs("payment_handle")), "", M_objrs("payment_handle"))
                
                ListItem.SubItems(32) = IIf(IsNull(M_objrs("occupation")), "", M_objrs("occupation"))
                ListItem.SubItems(33) = IIf(IsNull(M_objrs("reason")), "", M_objrs("reason"))
            M_objrs.MoveNext
        Wend
    End If
    Set M_objrs = Nothing

End Sub

Private Sub CmdCariRejected_Click()
    Call IsiLogRejected
End Sub

Private Sub CmdCekAll_Click()
    Dim K As Integer
    
    If LvPTP.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    PB1.Max = LvPTP.ListItems.Count
    For K = 1 To LvPTP.ListItems.Count
        PB1.Value = K
        LvPTP.ListItems(K).Checked = True
    Next K
End Sub

Private Sub CmdCekAllHamanto_Click()
    Dim K As Integer
    
    If LvHamanto.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    PB2.Max = LvHamanto.ListItems.Count
    For K = 1 To LvHamanto.ListItems.Count
        PB1.Value = K
        LvHamanto.ListItems(K).Checked = True
    Next K
End Sub

Private Sub CmdExport_Click()
    'PanelExport.Visible = True
    'Call Export_To_Excel
    Dim xx As Integer
    Dim ceklst As Boolean
    
    ceklst = False
    For xx = 1 To LvPTP.ListItems.Count
        If LvPTP.ListItems(xx).Checked = True Then
            ceklst = True
            Exit For
        End If
    Next xx
    
    If ceklst Then
        If CmbApprove.text <> "" Then
            Call My_Export_Excel
        Else
            MsgBox "Anda belum memilih data akan di 'Approve By : '", vbCritical + vbOKOnly, "Info"
        End If
    Else
        MsgBox "Anda belum memilih data!", vbOKOnly + vbCritical, "INFO"
    End If
End Sub

Private Sub CmdRefresh_Click()
    Call IsiLog
End Sub

Private Sub CmdReject_Click()
    Dim K As Integer
    Dim W As String
    
    If LvPTP.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Infromasi"
        Exit Sub
    End If
    
    W = MsgBox("Anda yakin akan melakukan Menghapus Request PTP?", vbYesNo + vbQuestion, "Konfirmasi")
    If W = vbNo Then
        MsgBox "Penghapusan PTP dibatalkan!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    PB1.Max = LvPTP.ListItems.Count
    
    For K = 1 To LvPTP.ListItems.Count
        PB1.Value = K
        If LvPTP.ListItems(K).Checked = True Then
            Call CatetLogReject(K)
            Call HapusData(K)
            Call KirimPesanGagal(K)
        End If
    Next K
    
    Call IsiLog
    MsgBox "Data Berhasil dihapus!", vbOKOnly + vbInformation, "Informasi"
End Sub

Private Sub cmdsearch_Click()
Dim S As String
Dim sql As String

    If cbsearch.text = "Tanggal Request by Email" Then
        S = "tgl_send_email"
    ElseIf cbsearch.text = "Cust ID" Then
        S = "CustId"
    ElseIf cbsearch.text = "Nama Cust" Then
        S = "vcustname"
    ElseIf cbsearch.text = "Tanggal Approve" Then
        S = "tgl_approve"
    End If
       
sql = "select * from tblsendptp_log_approve where id is not null "
If cbsearch.text = "Jenis PTP" Or cbsearch.text = "Cust ID" Or cbsearch.text = "Nama Cust" Then
sql = sql + "and " & S & " =  '" & txtsearch.text & "'"
ElseIf cbsearch.text = "Tanggal Approve" Or cbsearch.text = "Tanggal Request by Email" Then
sql = sql + "and " & S & " between '" & Format(date1.Value, "yyyy/mm/dd") & "' and '" & Format(date2.Value, "yyyy/mm/dd") & "' order by tgldata desc limit 300 "
End If

If IsNull(date1.Value) Or IsNull(date2.Value) Then
    MsgBox "Tanggal Wajib Diisi", vbOKOnly + vbInformation, "Informasi"
    Exit Sub
Else
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open sql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
End If

    LvPTPApproved.ListItems.CLEAR
    TxtJmlApproved.text = M_objrs.RecordCount

    If M_objrs.RecordCount > 0 Then
        Dim STATUS As String
        Dim DOB As String
         While Not M_objrs.EOF
            'On Error Resume Next
            Set ListItem = LvPTPApproved.ListItems.ADD(, , M_objrs("id"))
                ListItem.SubItems(1) = IIf(IsNull(M_objrs("jenis_ptp")), "", M_objrs("jenis_ptp"))
                ListItem.SubItems(2) = IIf(IsNull(M_objrs("custid")), "", M_objrs("custid"))
                ListItem.SubItems(3) = IIf(IsNull(M_objrs("vcustname")), "", M_objrs("vcustname"))
                
                If M_objrs("status") = "0" Then
                    STATUS = "Belum di Approve"
                End If
                If M_objrs("status") = "1" Then
                    STATUS = "Approve"
                End If
                If M_objrs("status") = "2" Then
                    STATUS = "Rejected"
                End If
            
         discount = IIf(IsNull(M_objrs("ndiscountamt")), "0", M_objrs("ndiscountamt"))
            
            
            If discount > 0 And discount <= 2000000 Then
                S = "Coll SPV"
            ElseIf discount <= 10000000 Then
                S = "Coll Band 6"
            ElseIf discount <= 20000000 Then
                S = "Coll Band 5"
            ElseIf discount <= 30000000 Then
                S = "Coll Band 4"
            ElseIf discount <= 50000000 Then
                S = "Head of Coll"
            ElseIf discount <= 100000000 Then
                S = "Head of CCC"
            ElseIf discount <= 2300000000# Then
                S = "Head of CRM"
            End If
                      
                ListItem.SubItems(4) = STATUS
                ListItem.SubItems(5) = IIf(IsNull(M_objrs("tgl_approve")), "", Format(M_objrs("tgl_approve"), "yyyy-mm-dd"))
                ListItem.SubItems(6) = S
                ListItem.SubItems(7) = IIf(IsNull(M_objrs("tgl_send_email")), "", Format(M_objrs("tgl_send_email"), "yyyy-mm-dd"))
                            
                ' Tandain klo ini belum di send_email 'jejaktian 14032016
                If IIf(IsNull(M_objrs("tgl_send_email")), 0, M_objrs("tgl_send_email")) = 0 Then
                    For K = 1 To 7
                        ListItem.ListSubItems(K).ForeColor = vbRed
                    Next K
                End If
                ' ------------------------------------------
            
            M_objrs.MoveNext
        Wend

    End If
    Set M_objrs = Nothing
    
End Sub

Private Sub cmdsudahemail_Click()
    Dim S As String
    
    For K = 1 To LvPTPApproved.ListItems.Count
        If LvPTPApproved.ListItems(K).Checked = True Then
            S = "update tblsendptp_log_approve set tgl_send_email=now() where id='"
            S = S + CStr(LvPTPApproved.ListItems(K).text) + "'"
            M_OBJCONN.Execute S
        End If
    Next K
    
    Call IsiLogApproved
    
End Sub

Private Sub CmdUnCekAll_Click()
    Dim K As Integer
    
    If LvPTP.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    PB1.Max = LvPTP.ListItems.Count
    For K = 1 To LvPTP.ListItems.Count
        PB1.Value = K
        LvPTP.ListItems(K).Checked = False
    Next K
End Sub



Private Sub CmdUnCekAllHamanto_Click()
    Dim K As Integer
    
    If LvHamanto.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    PB2.Max = LvHamanto.ListItems.Count
    For K = 1 To LvHamanto.ListItems.Count
        PB1.Value = K
        LvHamanto.ListItems(K).Checked = False
    Next K
End Sub
Private Function Export_To_Excel()
    Dim strsql          As String
    Dim rs              As ADODB.Recordset
    Dim ExlObj          As Excel.Application
    Dim objBook         As Excel.Workbook
    Dim objSheet        As Excel.Worksheet
    Dim ListCustId      As String
    
    On Error GoTo adderr
    
    For K = 1 To LvPTP.ListItems.Count
        If LvPTP.ListItems(K).Checked = True Then
            ListCustId = ListCustId & ",'" & LvPTP.ListItems(K).SubItems(2) & "'"
        End If
    Next K
    
    ListCustId = Mid(ListCustId, 2)
    
    strsql = "select custid,vcustname,'" & CmbApprove.text & "' as Approved FROM tblsendptp WHERE custid in (" & ListCustId & ")"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic

Form_Save:
    CD_save.ShowSave
    txtlocation.text = CD_save.FileName
    
    'Cek apakah user menekan tombol cancel pada dialog save
    If txtlocation.text = Empty Then
        'Tanyakan ke user.. apakah benar2 akan membatalkan proses download???
        m_msgbox = MsgBox("Anda ingin Download dibatalkan?", vbYesNo + vbQuestion, "Konfirmasi")
        'Jika user benar-benar akan membatalkan proses download, keluar dari fungsi ini!
        If m_msgbox = vbYes Then
              MsgBox "Download dibatalkan!", vbOKOnly + vbInformation, "Informasi"
              CmdExport.Enabled = True
              Exit Function
        End If
        If m_msgbox = vbNo Then '-> jika user tidak membatalkan proses download
          GoTo Form_Save        '-> maka goto form_save
        End If
    End If
 
    If rs.RecordCount > 0 Then
       PB1.Max = rs.RecordCount
    End If

 
    Set ExlObj = CreateObject("Excel.Application")
    Set objBook = ExlObj.Workbooks.ADD
    Set objSheet = objBook.ActiveSheet
        
    
    On Error GoTo Salah
    'Proses pengsisian nama field ke excel
    objSheet.Cells(1, 1).Value = "List CPA Approve"
    objSheet.Cells(1, 1).Font.Name = "Verdana"
    objSheet.Cells(1, 1).Font.Bold = True:
    objSheet.Cells(2, 1).Value = "Tanggal : " + Format(Now, "dd-mm-yyyy")
    objSheet.Cells(2, 1).Font.Name = "Verdana"
    objSheet.Cells(2, 1).Font.Bold = True:
    objSheet.Cells(4, 1).Value = "NO"
    objSheet.Cells(4, 2).Value = "CARD NUMBER"
    objSheet.Cells(4, 3).Value = "CH NAME"
    objSheet.Cells(4, 4).Value = "APPROVED"
    objSheet.Cells(4, 5).Value = "ADMIN CREATED"
    objSheet.Cells(4, 6).Value = "RECEIVED BY" 'Dikosongkan
    objSheet.Cells(4, 7).Value = "BAN 6"
    objSheet.Cells(4, 8).Value = "BAN 5"
    objSheet.Cells(4, 9).Value = "BAN 4"
    objSheet.Cells(4, 10).Value = "BAN 3"
    objSheet.Cells(4, 11).Value = "BAN 2"
    objSheet.Cells(4, 12).Value = "BAN 1"
    objSheet.Cells(4, 13).Value = "ADMIN RECEIVED"
    objSheet.Cells(4, 14).Value = "SENT BY"
'    objSheet.Cells(4, 4).Value = "DOB":      objSheet.Cells(4, 5).Value = "STATUS PTP"
'    objSheet.Cells(4, 6).Value = "TOTAL PAYMENT":      objSheet.Cells(4, 7).Value = "DOWN PAYMENT"
'    objSheet.Cells(4, 8).Value = "LPD FROM PAYMENT":      objSheet.Cells(4, 9).Value = "LPA FROM PAYMENT"
                
'select dpropsal,vcustid,vcustname,dob,status_ptp,nttlpayment,ndownpay,lpd_from_payment,lpa_from_payment
    
    objSheet.Range("A5").CopyFromRecordset rs '-> Proses pengisian data dimulai dari Cell A2
    objBook.SaveAs txtlocation.text, xlWorkbookNormal
    ExlObj.Quit
    Set ExlObj = Nothing: Set objBook = Nothing: Set objSheet = Nothing
    

    MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
    PB1.Value = 0
    Command1.Enabled = True
    
    
    
    Set rs = Nothing

    StartMeUp (txtlocation.text)

    txtlocation.text = ""
Salah:
    Exit Function
    
        
adderr:
    If err.number = -2147217900 Then
    On Error Resume Next
    Resume
    End If
    MsgBox err.Description



End Function

Private Sub cmrefresh_Click()
Unload Me
FrmListRequestPTP.Show vbModal
End Sub

Private Sub Command1_Click()
    Export_To_Excel
End Sub

Private Sub Command2_Click()
PanelExport.Visible = False
End Sub

Private Sub Form_Load()
    CmbTampilkan.text = "PTP NO DISC."
    CmbJenisRejected.text = "PTP NO DISC."
    SSTab1.TabVisible(3) = False
    Call HeaderLog
    Call IsiLog
    
    Call HeaderLogRejected
    Call IsiLogRejected
    
    Call HeaderLogApproved
    Call IsiLogApproved
    
    '@@221012 Bikin header log pak hamanto
    Call HeaderAppHamanto
        
    CmbApprove.AddItem MDIForm1.txtnama.text
    
    PanelExport.Visible = False
    'To Be Approved By Pak Hamanto
End Sub

Private Sub BikinCPA(K As Integer)
    Dim CMDSQL As String
    Dim Remarks As String
    Dim M_Objrs_Cek_Type As ADODB.Recordset
    Dim TypeAcc As String
    
    TypeAcc = ""
    
    '@@13022013 Cek type account dulu nih .. pil/card
    CMDSQL = "select acc_type from mgm where custid='"
    CMDSQL = CMDSQL & CStr(LvPTP.ListItems(K).SubItems(2)) & "'"
    Set M_Objrs_Cek_Type = New ADODB.Recordset
    M_Objrs_Cek_Type.CursorLocation = adUseClient
    M_Objrs_Cek_Type.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_Objrs_Cek_Type.RecordCount > 0 Then
        TypeAcc = IIf(IsNull(M_Objrs_Cek_Type("acc_type")), "", M_Objrs_Cek_Type("acc_type"))
    End If
    
    Set M_Objrs_Cek_Type = Nothing
    
    Call Cari_LPD_LPA_Payment(K)
    
    CMDSQL = "insert into tblcpa (dpropsal,vcustid,vproduct,nttlpayment,ndownpay,"
    CMDSQL = CMDSQL + "ncharge,ndiscountamt,vosbalance,vosprincipal,dtglinsert,vcustname,vjust,"
    CMDSQL = CMDSQL + "nbalance,nprincipal,nperiod,chkfaxed,chkwentalking,chkktp,"
    CMDSQL = CMDSQL + "chksup,chkbillings,chkothers,lpd_from_payment,lpa_from_payment,"
    CMDSQL = CMDSQL + "f_system,dob,status_ptp,ketother "
    
    '@@19062012 Jika Status PTP DISCON Catat Approvenya
    If Trim(UCase(LvPTP.ListItems(K).SubItems(1))) = "PTP DISCOUNT" Then
        CMDSQL = CMDSQL + ",tglapprove,sts_approve,approve_by,logapprove_by "
    End If
    
    'Catet Juga yang PTP No Discon 20062012
    If Trim(UCase(LvPTP.ListItems(K).SubItems(1))) = "PTP NO DISCOUNT" Then
        CMDSQL = CMDSQL + ",tglapprove,sts_approve,approve_by,logapprove_by "
    End If
    
    '@@16-07-2012 Buat Catet Payment Handle
    CMDSQL = CMDSQL + " ,vpaymenthandle,voccupation,vreason "
    
    CMDSQL = CMDSQL + ") values ("
    CMDSQL = CMDSQL + "now(),'"
    'Cmdsql = Cmdsql + CStr(LvPTP.ListItems(K).SubItems(2)) + "','CARD','"
    CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).SubItems(2)) + "','"
    CMDSQL = CMDSQL + TypeAcc + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(15)), "0", Replace(LvPTP.ListItems(K).SubItems(15), ",", ""))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(16)), "0", Replace(LvPTP.ListItems(K).SubItems(16), ",", ""))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(17)), "0", Replace(LvPTP.ListItems(K).SubItems(17), ",", ""))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(18)), "0", Replace(LvPTP.ListItems(K).SubItems(18), ",", ""))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(19)), "", LvPTP.ListItems(K).SubItems(19))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(20)), "", LvPTP.ListItems(K).SubItems(20))) + "',"
    CMDSQL = CMDSQL + "now(),'"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(3)), "", LvPTP.ListItems(K).SubItems(3))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(21)), "", LvPTP.ListItems(K).SubItems(21))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(12)), "0", Replace(LvPTP.ListItems(K).SubItems(12), ",", ""))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(11)), "0", Replace(LvPTP.ListItems(K).SubItems(11), ",", ""))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(8)), "", LvPTP.ListItems(K).SubItems(8))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(22)), "", LvPTP.ListItems(K).SubItems(22))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(23)), "", LvPTP.ListItems(K).SubItems(23))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(24)), "", LvPTP.ListItems(K).SubItems(24))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(25)), "", LvPTP.ListItems(K).SubItems(25))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(26)), "", LvPTP.ListItems(K).SubItems(26))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(27)), "", LvPTP.ListItems(K).SubItems(27))) + "',"
    CMDSQL = CMDSQL + IIf(TxtLPDPayment.text = "", "null", "'" + TxtLPDPayment.text + "'") + ",'"
    CMDSQL = CMDSQL + CStr(TxtLPAPayment.Value) + "','1',"
    '@@20062012 Tambahkan DOB dan Status PTP
    CMDSQL = CMDSQL + IIf(LvPTP.ListItems(K).SubItems(29) = "", "null", "'" + LvPTP.ListItems(K).SubItems(29) + "'")
    CMDSQL = CMDSQL + ",'" + LvPTP.ListItems(K).SubItems(1) + "',' "
    '@@21062012 Tambahkan Keterangan Other
    CMDSQL = CMDSQL + IIf(IsNull(LvPTP.ListItems(K).SubItems(30)), "", LvPTP.ListItems(K).SubItems(30)) + "' "
    
    '@@19062012 Buat nyatet approvenya
     If Trim(UCase(LvPTP.ListItems(K).SubItems(1))) = "PTP DISCOUNT" Then
        CMDSQL = CMDSQL + ",now(),'1','"
        CMDSQL = CMDSQL + Trim(CmbApprove.text) + "','"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "'"
     End If
     
     'Buat nyatet yang jenisnya PTP NO Discount.
     If Trim(UCase(LvPTP.ListItems(K).SubItems(1))) = "PTP NO DISCOUNT" Then
        CMDSQL = CMDSQL + ",now(),'1','"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "','"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "'"
     End If
    
    CMDSQL = CMDSQL + ",'"
    CMDSQL = CMDSQL + IIf(IsNull(LvPTP.ListItems(K).SubItems(31)), "", LvPTP.ListItems(K).SubItems(31)) + "','"
    
    CMDSQL = CMDSQL + IIf(IsNull(LvPTP.ListItems(K).SubItems(32)), "", LvPTP.ListItems(K).SubItems(32)) + "','"
    CMDSQL = CMDSQL + IIf(IsNull(LvPTP.ListItems(K).SubItems(33)), "", LvPTP.ListItems(K).SubItems(33)) + "')"
    DoEvents
    M_OBJCONN.Execute CMDSQL
    
    '@@19062012 Bikin Remarks untuk CPA
     '@@11092012 Tulis Remarks baik untuk yang ptp discon/no discon
     If Trim(UCase(LvPTP.ListItems(K).SubItems(1))) = "PTP DISCOUNT" Then
        Remarks = "PtpDisc-"
     Else
        Remarks = "PTPNoDisc-"
     End If
        Remarks = Remarks + "App By:" + CmbApprove.text + "-"
        Remarks = Remarks + "Ttl.Pymt: Rp." + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(7)), "", LvPTP.ListItems(K).SubItems(7))) + " -"
        Remarks = Remarks + "Instl: " + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(8)), "", LvPTP.ListItems(K).SubItems(8))) + " -"
        Remarks = Remarks + "From Bal.: Rp." + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(12)), "", LvPTP.ListItems(K).SubItems(12))) + " -"
        Remarks = Remarks + "From Prin.: Rp." + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(14)), "", LvPTP.ListItems(K).SubItems(14))) + " -"
        Remarks = Remarks + "%Balance: " + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(19)), "", LvPTP.ListItems(K).SubItems(19))) + "% -"
        Remarks = Remarks + "%Principal: " + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(20)), "", LvPTP.ListItems(K).SubItems(20))) + "% #USER LOG:" + MDIForm1.TxtUsername.text
        
        CMDSQL = "insert into mgm_hst (custid, agent, products, "
        CMDSQL = CMDSQL + "hst,user_log) values ('"
        CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).SubItems(2)) + "','"
        CMDSQL = CMDSQL + LvPTP.ListItems(K).SubItems(28) + "','"
        CMDSQL = CMDSQL + "Collection" + "','"
        CMDSQL = CMDSQL + Remarks + "','"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "')"
        M_OBJCONN.Execute CMDSQL
    
    
    '@@25072012,Update yang approve dan tanggal proposalnya di tabel tblsendptp jika PTP discount
    If Trim(UCase(LvPTP.ListItems(K).SubItems(1))) = "PTP DISCOUNT" Then
        CMDSQL = "update tblsendptp set tgl_proposal=now(), approve_by='"
        CMDSQL = CMDSQL + CStr(Trim(CmbApprove.text)) + "', log_approve='"
        CMDSQL = CMDSQL + CStr(Trim(MDIForm1.TxtUsername.text)) + "' where id='"
        CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).text) + "'"
        M_OBJCONN.Execute CMDSQL
    End If
    
    If Trim(UCase(LvPTP.ListItems(K).SubItems(1))) = "PTP NO DISCOUNT" Then
        CMDSQL = "update tblsendptp set tgl_proposal=now(), approve_by='"
        CMDSQL = CMDSQL + CStr(Trim(MDIForm1.TxtUsername.text)) + "', log_approve='"
        CMDSQL = CMDSQL + CStr(Trim(MDIForm1.TxtUsername.text)) + "' where id='"
        CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).text) + "'"
        M_OBJCONN.Execute CMDSQL
    End If
End Sub

'@@ 16-03-2011, Ini buat nyari LPD dan LPA terakhir dari tabel lunas
Private Sub Cari_LPD_LPA_Payment(K As Integer)
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    
    StatusPTP = ""
    TxtLPDPayment.text = ""
    TxtLPAPayment.Value = "0"
    
    CMDSQL = "select paydate,payment from tbllunas where custid='"
    CMDSQL = CMDSQL + Trim(LvPTP.ListItems(K).SubItems(2)) + "' order by paydate desc limit 1 "
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_objrs.RecordCount > 0 Then
            StatusPTP = "PTP-POP"
            TxtLPDPayment.text = IIf(IsNull(M_objrs("paydate")), "", Format(M_objrs("paydate"), "yyyy-mm-dd"))
            TxtLPAPayment.Value = IIf(IsNull(M_objrs("payment")), "0", M_objrs("payment"))
            LpdPayment = "'" + TxtLPDPayment.text + "'"
        Else
            StatusPTP = "PTP-NEW"
            'LpdPayment = "null"
            TxtLPDPayment.text = ""
            TxtLPAPayment.Value = "0"
        End If
    Set M_objrs = Nothing
End Sub

Private Sub BikinPTP(K As Integer)
    Dim CMDSQL As String
    Dim i As Integer
    Dim M_Objrs_Cek_Tgl As ADODB.Recordset
    Dim jumlah_tenor As Integer
    
    'Tambahan Randy 7April2015 Untuk ambil jumlah tenor sebagai validasi
    jumlah_tenor = Val(LvPTP.ListItems(K).SubItems(8))
    
    bcekptp = True
    
        'Jika Tenor=1
        If jumlah_tenor = 1 Then
                  
                '@@14-04-2012 Cek Data
                CMDSQL = "select * from tblnegoptp where custid='"
                CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).SubItems(2)) + "' and date(promisedate)='"
                CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).SubItems(6)) + "'"
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        CMDSQL = "delete from tblnegoptp where id='"
                        CMDSQL = CMDSQL + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.Execute CMDSQL
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
                  
            jatuhtempo = LvPTP.ListItems(K).SubItems(6)
            CMDSQL = "INSERT INTO TblNegoPTP "
            CMDSQL = CMDSQL + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
            CMDSQL = CMDSQL + "VALUES "
            CMDSQL = CMDSQL + "('" + CStr(LvPTP.ListItems(K).SubItems(2)) + "', "
            CMDSQL = CMDSQL + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
            CMDSQL = CMDSQL + "" + CStr(Replace(LvPTP.ListItems(K).SubItems(13), ",", "")) + " , "
            CMDSQL = CMDSQL + "now(), "
            CMDSQL = CMDSQL + "'IPO')"
            M_OBJCONN.Execute CMDSQL
            
            
'-------------------------------- 02-07-2012 Negoptp Log ga usah di cek deh lama---------
'                '@@14-04-2012 Cek Data
'                CMDSQL = "select * from tblnegoptp_log where custid='"
'                CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).SubItems(2)) + "' and date(promisedate)='"
'                CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).SubItems(6)) + "'"
'                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
'                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
'                M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
'                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
'                    While Not M_Objrs_Cek_Tgl.EOF
'                        CMDSQL = "delete from tblnegoptp_log where id='"
'                        CMDSQL = CMDSQL + CStr(M_Objrs_Cek_Tgl("id")) + "'"
'                        M_OBJCONN.Execute CMDSQL
'                        M_Objrs_Cek_Tgl.MoveNext
'                    Wend
'                End If
'                Set M_Objrs_Cek_Tgl = Nothing
'-------------------------------- 02-07-2012 Negoptp Log ga usah di cek deh lama---------
            
            
            ' isi ke tbl log_ptp
            CMDSQL = "INSERT INTO tblnegoptp_log "
            CMDSQL = CMDSQL + "(custid,PromiseDate, Promisepay,tglInput,agent,stsacc) "
            CMDSQL = CMDSQL + "VALUES "
            CMDSQL = CMDSQL + "('" + CStr(LvPTP.ListItems(K).SubItems(2)) + "', "
            CMDSQL = CMDSQL + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
            CMDSQL = CMDSQL + "" + CStr(Replace(LvPTP.ListItems(K).SubItems(13), ",", "")) + " , "
            CMDSQL = CMDSQL + "now(), "
            CMDSQL = CMDSQL + "'" + CStr(LvPTP.ListItems(K).SubItems(28)) + "','P')"
            M_OBJCONN.Execute CMDSQL
                
        Else
            'Untuk Tenor yang lebih dari 1
                        
                'Hapus Reserved Data
                CMDSQL = "delete from tblreserve where custid='"
                CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).SubItems(2)) + "'"
                M_OBJCONN.Execute CMDSQL
                        
                jatuhtempo = CStr(LvPTP.ListItems(K).SubItems(6))
            
                '@@14-04-2012 Cek Data
                CMDSQL = "select * from tblnegoptp where custid='"
                CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).SubItems(2)) + "' and date(promisedate)='"
                CMDSQL = CMDSQL + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "'"
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        CMDSQL = "delete from tblnegoptp where id='"
                        CMDSQL = CMDSQL + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.Execute CMDSQL
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
            
            CMDSQL = "INSERT INTO TblNegoPTP "
            CMDSQL = CMDSQL + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
            CMDSQL = CMDSQL + "VALUES "
            CMDSQL = CMDSQL + "('" + CStr(LvPTP.ListItems(K).SubItems(2)) + "', "
            CMDSQL = CMDSQL + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
            CMDSQL = CMDSQL + "" + CStr(Replace(LvPTP.ListItems(K).SubItems(13), ",", "")) + " , "
            CMDSQL = CMDSQL + "now(), "
            CMDSQL = CMDSQL + "'IPO')"
            M_OBJCONN.Execute CMDSQL
            
            
'-------------------------------- 02-07-2012 Negoptp Log ga usah di cek deh lama---------
'                '@@14-04-2012 Cek Data
'                CMDSQL = "select * from tblnegoptp_log where custid='"
'                CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).SubItems(2)) + "' and date(promisedate)='"
'                CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).SubItems(6)) + "'"
'                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
'                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
'                M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
'
'                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
'                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
'                    While Not M_Objrs_Cek_Tgl.EOF
'                        CMDSQL = "delete from tblnegoptp_log where id='"
'                        CMDSQL = CMDSQL + CStr(M_Objrs_Cek_Tgl("id")) + "'"
'                        M_OBJCONN.Execute CMDSQL
'                        M_Objrs_Cek_Tgl.MoveNext
'                    Wend
'                End If
'                Set M_Objrs_Cek_Tgl = Nothing
'-------------------------------- 02-07-2012 Negoptp Log ga usah di cek deh lama---------
            
            
            'isi ke tbl log_ptp
            CMDSQL = "INSERT INTO tblnegoptp_log "
            CMDSQL = CMDSQL + "(custid,PromiseDate, Promisepay,tglInput,agent,stsacc) "
            CMDSQL = CMDSQL + "VALUES "
            CMDSQL = CMDSQL + "('" + CStr(LvPTP.ListItems(K).SubItems(2)) + "', "
            CMDSQL = CMDSQL + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
            CMDSQL = CMDSQL + "" + CStr(Replace(LvPTP.ListItems(K).SubItems(13), ",", "")) + " , "
            CMDSQL = CMDSQL + "now(), "
            CMDSQL = CMDSQL + "'" + CStr(LvPTP.ListItems(K).SubItems(28)) + "','P')"
            M_OBJCONN.Execute CMDSQL
                
                
            n = 0
            
            Call HitungInstallmentPtp(K)
            
            For i = 1 To (Val(LvPTP.ListItems(K).SubItems(8)))
                n = n + 1
                'JMLPAY = ((.TxtPayment - txtPembayaranAwal.Value) - PaymentTenor) / (.txttenor.Value - 1)
                JmlPay = PaymentTenor
                Vrdate = DateAdd("m", n, Format(LvPTP.ListItems(K).SubItems(6), "yyyy-mm-dd"))
                    
                '@@14-04-2012 Cek Data
                CMDSQL = "select * from tblreserve where custid='"
                CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).SubItems(2)) + "' and date(promisedate)='"
                CMDSQL = CMDSQL + CStr(Format(Vrdate, "yyyy-mm-dd")) + "'"
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        CMDSQL = "delete from tblreserve where id='"
                        CMDSQL = CMDSQL + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.Execute CMDSQL
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
                    
                    CMDSQL = "INSERT INTO tblreserve "
                    CMDSQL = CMDSQL + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
                    CMDSQL = CMDSQL + "VALUES "
                    CMDSQL = CMDSQL + "('" + CStr(LvPTP.ListItems(K).SubItems(2)) + "', "
                    CMDSQL = CMDSQL + "'" + CStr(Format(Vrdate, "yyyy-mm-dd")) + "', "
                    CMDSQL = CMDSQL + "" + CStr(JmlPay) + " , "
                    CMDSQL = CMDSQL + "now(), "
                    CMDSQL = CMDSQL + "'IPO')"
                    M_OBJCONN.Execute CMDSQL
  
                    CMDSQL = "INSERT INTO TblNegoptp_log "
                    CMDSQL = CMDSQL + "(custid,PromiseDate, Promisepay,tglinput,agent,stsacc) "
                    CMDSQL = CMDSQL + "VALUES "
                    CMDSQL = CMDSQL + "('" + CStr(LvPTP.ListItems(K).SubItems(2)) + "', "
                    CMDSQL = CMDSQL + "'" + CStr(Format(Vrdate, "yyyy-mm-dd")) + "', "
                    CMDSQL = CMDSQL + "" + CStr(JmlPay) + " , "
                    CMDSQL = CMDSQL + "now(), "
                    CMDSQL = CMDSQL + "'" + CStr(LvPTP.ListItems(K).SubItems(28)) + "','R')"
                    M_OBJCONN.Execute CMDSQL

            'INSERT KE TABEL PTP-REGULER(Randy07-04-2015)
            CMDSQL = "select * from tblnegoptp_reguler where custid='"
            CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).SubItems(2)) + "' and date(promisedate)='"
            CMDSQL = CMDSQL + CStr(Format(Vrdate, "yyyy-mm-dd")) + "'"
            Set M_Objrs_Cek_Tgl = New ADODB.Recordset
            M_Objrs_Cek_Tgl.CursorLocation = adUseClient
            M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
            
            If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                While Not M_Objrs_Cek_Tgl.EOF
                    CMDSQL = "delete from tblnegoptp_reguler where id='"
                    CMDSQL = CMDSQL + CStr(cnull(M_Objrs_Cek_Tgl("id"))) + "'"
                    M_OBJCONN.Execute CMDSQL
                    M_Objrs_Cek_Tgl.MoveNext
                Wend
            End If
            Set M_Objrs_Cek_Tgl = Nothing
        
            CMDSQL = "INSERT INTO tblnegoptp_reguler"
            CMDSQL = CMDSQL + "(custid, balance, PromiseDate, Promisepay, inputdate, type, tenor, down_payment, agent, keterangan_ptp) "
            CMDSQL = CMDSQL + "VALUES "
            CMDSQL = CMDSQL + "('" + CStr(LvPTP.ListItems(K).SubItems(2)) + "', "
            CMDSQL = CMDSQL + " '" + CStr(LvPTP.ListItems(K).SubItems(7)) + "',"
            CMDSQL = CMDSQL + "'" + CStr(Format(Vrdate, "yyyy-mm-dd")) + "', "
            CMDSQL = CMDSQL + "" + CStr(JmlPay) + " , "
            CMDSQL = CMDSQL + "now(), "
            CMDSQL = CMDSQL + "'Reguler',"
            CMDSQL = CMDSQL + " '" + CStr(LvPTP.ListItems(K).SubItems(8)) + "',"
            CMDSQL = CMDSQL + " '" + CStr(LvPTP.ListItems(K).SubItems(16)) + "',"
            CMDSQL = CMDSQL + " '" + CStr(LvPTP.ListItems(K).SubItems(28)) + "', "
            CMDSQL = CMDSQL + "'PTP-NEW')"
            M_OBJCONN.Execute CMDSQL
       Next i
       End If
    
    
    PaymentTenor = 0
    
    'MsgBox "PTP berhasil ditambahkan!", vbOKOnly + vbInformation, "Informasi"
End Sub


'@@22-09-2011 Hitung InstallmentPtp
Private Sub HitungInstallmentPtp(K As Integer)
    Dim installment As Double
    
        If Val(LvPTP.ListItems(K).SubItems(8)) = 0 Or Val(LvPTP.ListItems(K).SubItems(8)) = 1 Then
            installment = Val(Replace(LvPTP.ListItems(K).SubItems(15), ",", "")) / 1
        Else
            installment = (Val(Replace(LvPTP.ListItems(K).SubItems(15), ",", "")) - Val(Replace(LvPTP.ListItems(K).SubItems(13), ",", ""))) / (Val(LvPTP.ListItems(K).SubItems(8)) - 1)
        End If
        PaymentTenor = Ceiling(installment)
End Sub

Private Sub CatetLogApprove(K As Integer)
    Dim CMDSQL As String
    CMDSQL = "update tblsendptp set status = 1 where id = '" + CStr(LvPTP.ListItems(K).text) + "'"
    DoEvents
    M_OBJCONN.Execute CMDSQL
    CMDSQL = "insert into tblsendptp_log_approve "
    CMDSQL = CMDSQL + "select * from tblsendptp where id='"
    CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).text) + "'"
    DoEvents
    M_OBJCONN.Execute CMDSQL
End Sub

Private Sub CatetLogReject(K As Integer)
    Dim CMDSQL As String
    
    '@@25072012 Catet nih siapa yang melakukan reject
    CMDSQL = "update tblsendptp set tgl_proposal=now(),log_approve='"
    CMDSQL = CMDSQL + CStr(Trim(MDIForm1.TxtUsername.text)) + "' where id='"
    CMDSQL = CMDSQL + CStr(Trim(LvPTP.ListItems(K).text)) + "'"
    M_OBJCONN.Execute CMDSQL
    
    CMDSQL = "insert into tblsendptp_log_reject "
    CMDSQL = CMDSQL + "select * from tblsendptp where id='"
    CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).text) + "'"
    M_OBJCONN.Execute CMDSQL
End Sub

Private Sub BikinStatusPTP(K As Integer)
    Dim CMDSQL As String
    Dim Cmdsql_Cek As String
    Dim StatusRemarks As String
    Dim M_Objrs_Cek As ADODB.Recordset
    Dim AmountNew As Double
    
    AmountNew = 0
    
    Cmdsql_Cek = "select * from tblnegoptp where custid='"
    Cmdsql_Cek = Cmdsql_Cek + CStr(LvPTP.ListItems(K).SubItems(2)) + "' order by id desc limit 1"
    Set M_Objrs_Cek = New ADODB.Recordset
    M_Objrs_Cek.CursorLocation = adUseClient
    M_Objrs_Cek.Open Cmdsql_Cek, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs_Cek.RecordCount > 0 Then
        AmountNew = Val(IIf(IsNull(M_Objrs_Cek("promisepay")), "0", M_Objrs_Cek("promisepay")))
    Else
       AmountNew = 0
    End If
    
    'Jika StatusPTP=PTP NEW
    If StatusPTP = "PTP-NEW" Then
        Dim M_Objrs_Cek_Status As ADODB.Recordset
        Dim Cmdsql_Cek_status As String
        Dim TglPTPNew As String
        
        'Cari apakah sebelumnya status data=ptp new, jika iya maka tglptpnew tidak usah diupdate
        'Tapi jika status sebelumnya bukan ptp new maka update tglptpnew=now
        Cmdsql_Cek_status = "select * from mgm where custid='"
        Cmdsql_Cek_status = Cmdsql_Cek_status + CStr(LvPTP.ListItems(K).SubItems(2)) + "'"
        Set M_Objrs_Cek_Status = New ADODB.Recordset
        M_Objrs_Cek_Status.CursorLocation = adUseClient
        M_Objrs_Cek_Status.Open Cmdsql_Cek_status, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_Objrs_Cek_Status.RecordCount > 0 Then
            If M_Objrs_Cek_Status("tglptpnew") = "" Or IsNull(M_Objrs_Cek_Status("tglptpnew")) = True _
               Or M_Objrs_Cek_Status("tglptpnew") = Empty Then
                TglPTPNew = "now()"
             Else
                TglPTPNew = "'" + CStr(Format(M_Objrs_Cek_Status("tglptpnew"), "yyyy-mm-dd")) + "'"
             End If
        End If
        
        Set M_Objrs_Cek_Status = Nothing
    
        CMDSQL = "update mgm set dateptpnew='"
        CMDSQL = CMDSQL + LvPTP.ListItems(K).SubItems(6) + "',tgl_tagih='"
        CMDSQL = CMDSQL + LvPTP.ListItems(K).SubItems(10) + "', amountnew='"
        'CMDSQL = CMDSQL + CStr(Replace(LvPTP.ListItems(K).SubItems(15), ",", "")) + "',tglallptp='"
        'CMDSQL = CMDSQL + CStr(Replace(LvPTP.ListItems(K).SubItems(13), ",", "")) + "',tglallptp='"
        
        '@@20062012, amountnew ambil dari negoptp terakhir aja deh....
        CMDSQL = CMDSQL + CStr(AmountNew) + "',tglallptp='"
        
        CMDSQL = CMDSQL + LvPTP.ListItems(K).SubItems(6) + "',f_cek_new='PTP-NE',"
        CMDSQL = CMDSQL + "tglincoming=now(),ttlptp='"
        CMDSQL = CMDSQL + CStr(Replace(LvPTP.ListItems(K).SubItems(15), ",", "")) + "',"
        CMDSQL = CMDSQL + "kethslkerja_new='PTP-NEW',kethslkerjadesc_new='PTP-NEW',ptpvia='"
        CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).SubItems(9)) + "',ptpdesc='PTP-NEW', dateptp='"
        CMDSQL = CMDSQL + LvPTP.ListItems(K).SubItems(6) + "',tglptpnew=" + TglPTPNew
        CMDSQL = CMDSQL + ",tenor='"
        CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).SubItems(8)) + "' "
        CMDSQL = CMDSQL + "where custid='"
        CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).SubItems(2)) + "'"
        DoEvents
        M_OBJCONN.Execute CMDSQL
        
    End If
    
    If StatusPTP = "PTP-POP" Then
        CMDSQL = "update mgm set dateptp='"
        CMDSQL = CMDSQL + LvPTP.ListItems(K).SubItems(6) + "',tgl_tagih='"
        CMDSQL = CMDSQL + LvPTP.ListItems(K).SubItems(10) + "',tglallptp='"
        CMDSQL = CMDSQL + LvPTP.ListItems(K).SubItems(6) + "',f_cek_new='PTP-PO',"
        CMDSQL = CMDSQL + "tglincoming=now(),ttlptp='"
        CMDSQL = CMDSQL + CStr(Replace(LvPTP.ListItems(K).SubItems(15), ",", "")) + "',"
        CMDSQL = CMDSQL + "kethslkerja_new='PTP-POP',kethslkerjadesc_new='PTP-POP',ptpvia='"
        CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).SubItems(9)) + "',ptpdesc='PTP-POP',amountptp='"
        CMDSQL = CMDSQL + CStr(Replace(LvPTP.ListItems(K).SubItems(15), ",", "")) + "',tenor='"
        CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).SubItems(8)) + "' "
        CMDSQL = CMDSQL + "where custid='"
        CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).SubItems(2)) + "'"
        M_OBJCONN.Execute CMDSQL
    End If
    
     '@@19062012 Bikin Remark Status PTP
        StatusRemarks = "PTP Approve by: " & MDIForm1.TxtUsername.text & "/"
        StatusRemarks = StatusRemarks & "Jenis PTP:" & StatusPTP & "/"
        StatusRemarks = StatusRemarks & "Amount PTP:"
        StatusRemarks = StatusRemarks & CStr(Replace(LvPTP.ListItems(K).SubItems(15), ",", "")) & "/"
        StatusRemarks = StatusRemarks & "PTP Via:" & ""
        StatusRemarks = StatusRemarks & CStr(Replace(LvPTP.ListItems(K).SubItems(9), ",", "")) & "/"
        StatusRemarks = StatusRemarks & "Date PTP:" & Format(LvPTP.ListItems(K).SubItems(6), "yyyy-mm-dd")
        
        CMDSQL = "insert into mgm_hst(custid,agent,hst,f_cek_new,user_log,kodeds,lastcall) values ('"
        CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).SubItems(2)) + "','"
        CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).SubItems(28)) + "','"
        CMDSQL = CMDSQL + StatusRemarks + "','"
        CMDSQL = CMDSQL + StatusPTP + "','"
        CMDSQL = CMDSQL + CStr(MDIForm1.TxtUsername.text) + "','"
        CMDSQL = CMDSQL + StatusPTP + "','PTP')"
        
        M_OBJCONN.Execute CMDSQL
        
   Set M_Objrs_Cek = Nothing
        
End Sub

Private Sub HapusData(K As Integer)
    Dim CMDSQL As String
    
    CMDSQL = "delete from tblsendptp where id='"
    CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).text) + "'"
    M_OBJCONN.Execute CMDSQL
End Sub

Private Sub KirimPesan(K As Integer)
    Dim CMDSQL As String
    Dim Remarks As String
    Dim M_objrs As ADODB.Recordset
    
    Remarks = "Pembuatan PTP untuk custid: " & LvPTP.ListItems(K).SubItems(2) & " telah di approve!"
    
    CMDSQL = "insert into msgtbl "
    CMDSQL = CMDSQL + "( recipient, datetime, sender, sentfrom, msg) values ('"
    CMDSQL = CMDSQL + LvPTP.ListItems(K).SubItems(28) + "','"
    CMDSQL = CMDSQL + Format(Now(), "yyyymmdd") + "','"
    CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "','"
    CMDSQL = CMDSQL + CStr(MDIForm1.Winsock1.LocalIP) + "','"
    CMDSQL = CMDSQL + Remarks + "')"
    M_OBJCONN.Execute CMDSQL
        
        
    '@@19072012 Kirim Pesan Buat Ke TL
    'Cari Nama TLNYA
    CMDSQL = "select team from usertbl where userid='"
    CMDSQL = CMDSQL + CStr(Trim(LvPTP.ListItems(K).SubItems(28))) + "' "
    CMDSQL = CMDSQL + " and team is not null "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
        CMDSQL = "insert into msgtbl "
        CMDSQL = CMDSQL + "(recipient, datetime, sender, sentfrom, msg) values ('"
        CMDSQL = CMDSQL + CStr(Trim(M_objrs("team"))) + "','"
        CMDSQL = CMDSQL + Format(Now(), "yyyymmdd") + "','"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "','"
        CMDSQL = CMDSQL + CStr(MDIForm1.Winsock1.LocalIP) + "','"
        CMDSQL = CMDSQL + Remarks + "')"
        M_OBJCONN.Execute CMDSQL
    End If
    
    Set M_objrs = Nothing
End Sub

Private Sub KirimPesanGagal(K As Integer)
    Dim CMDSQL As String
    Dim Remarks As String
    
    Remarks = "Pembuatan PTP untuk custid: " & LvPTP.ListItems(K).SubItems(2) & " telah di reject!"
    
    CMDSQL = "insert into msgtbl "
    CMDSQL = CMDSQL + "( recipient, datetime, sender, sentfrom, msg) values ('"
    CMDSQL = CMDSQL + LvPTP.ListItems(K).SubItems(28) + "','"
    CMDSQL = CMDSQL + Format(Now(), "yyyymmdd") + "','"
    CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "','"
    CMDSQL = CMDSQL + CStr(MDIForm1.Winsock1.LocalIP) + "','"
    CMDSQL = CMDSQL + Remarks + "')"
    
    M_OBJCONN.Execute CMDSQL
End Sub





Private Sub LvPTP_DblClick()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    Dim TypeAcc As String
    
    TypeAcc = ""
    
    If LvPTP.ListItems.Count = 0 Then
        MsgBox "Data tidak tersedia!", vbOKOnly + vbInformation, "Informasi"
        Exit Sub
    End If
    
    CMDSQL = "select * from mgm where custid='"
    CMDSQL = CMDSQL + CStr(LvPTP.SelectedItem.SubItems(2)) + "'"
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
        FrmViewPTP.dtcardopen.Value = IIf(IsNull(M_objrs("opendate")), "", Format(M_objrs("opendate"), "dd/mm/yyyy"))
        FrmViewPTP.dwo.Value = IIf(IsNull(M_objrs("b_d")), "", Format(M_objrs("b_d"), "dd/mm/yyyy"))
        FrmViewPTP.txtregion.text = IIf(IsNull(M_objrs("region")), "", M_objrs("region"))
    End If
    
    TypeAcc = IIf(IsNull(M_objrs("acc_type")), "", M_objrs("acc_type"))
    
    Set M_objrs = Nothing
    
    Call Cari_LPD_LPA_Payment_2
    
    With LvPTP.SelectedItem
    
        If UCase(MDIForm1.txtlevel.text) = "TEAMLEADER" Then
            If CmbTampilkan.text = "PTP DISC." And _
               LvPTP.SelectedItem.SubItems(4) = "Belum di Approve" Then
                FrmViewPTP.CmdApprove.Caption = "Cetak"
            Else
                FrmViewPTP.CmdApprove.Caption = "Approve"
            End If
        End If
           
        FrmViewPTP.CmbJenisPTP.text = Trim(IIf(IsNull(.SubItems(1)), "", .SubItems(1)))
        
        FrmViewPTP.CmbPaymentHandle.text = IIf(IsNull(.SubItems(31)), "", .SubItems(31))
        FrmViewPTP.CmbOccupation.text = IIf(IsNull(.SubItems(32)), "", .SubItems(32))
        FrmViewPTP.CmbReason.text = IIf(IsNull(.SubItems(33)), "", .SubItems(33))
        
        FrmViewPTP.txtothers.text = IIf(IsNull(.SubItems(30)), "", .SubItems(30))
        
        FrmViewPTP.txtproduct.text = TypeAcc
        FrmViewPTP.dtpropsal.Value = Now()
        
        FrmViewPTP.TxtIdCpa.text = IIf(IsNull(.text), "", .text)
        FrmViewPTP.txtcardno.text = IIf(IsNull(.SubItems(2)), "", .SubItems(2))
        FrmViewPTP.txtname.text = IIf(IsNull(.SubItems(3)), "", .SubItems(3))
        FrmViewPTP.lblLastPay.Value = IIf(IsNull(.SubItems(7)), "0", Replace(.SubItems(7), ",", ""))
        FrmViewPTP.tdbisnstallment.Value = IIf(IsNull(.SubItems(8)), "1", .SubItems(8))
        
        FrmViewPTP.txtprincipal.Value = IIf(IsNull(.SubItems(11)), "0", Replace(.SubItems(11), ",", ""))
        FrmViewPTP.Label8.text = IIf(IsNull(.SubItems(11)), "0", Replace(.SubItems(11), ",", ""))
        
        FrmViewPTP.txtbalance.Value = IIf(IsNull(.SubItems(12)), "0", Replace(.SubItems(12), ",", ""))
        FrmViewPTP.Label5.text = IIf(IsNull(.SubItems(12)), "0", Replace(.SubItems(12), ",", ""))
        
        FrmViewPTP.txtcharge.Value = IIf(IsNull(.SubItems(17)), "0", Replace(.SubItems(17), ",", ""))
        FrmViewPTP.txtjust.text = IIf(IsNull(.SubItems(21)), "", .SubItems(21))
        FrmViewPTP.txtdownpayment.Value = IIf(IsNull(.SubItems(16)), "0", Replace(.SubItems(16), ",", ""))
        
        If .SubItems(22) = "1" Then
            FrmViewPTP.chkfaxed.Value = 1
        End If
        
        If .SubItems(23) = "1" Then
            FrmViewPTP.chkwentalk.Value = 1
        End If
        
        If .SubItems(24) = "1" Then
            FrmViewPTP.chkKTP.Value = 1
        End If
        
        If .SubItems(25) = "1" Then
            FrmViewPTP.chkpp.Value = 1
        End If
        
        If .SubItems(26) = "1" Then
            FrmViewPTP.chkbillings.Value = 1
        End If
        
        If .SubItems(27) = "1" Then
            FrmViewPTP.Check1.Value = 1
        End If
        
        FrmViewPTP.txtcollect.text = .SubItems(28)
        '@@20062012 Tambahan DOb
        'asep FrmViewPTP.TxtDob.text = IIf(.SubItems(29) = "", "", .SubItems(29))
        'FrmViewPTP.txtproduct.Text = "CARD"
        FrmViewPTP.txtplace.text = "CardHolder"
        FrmViewPTP.Show vbModal
    End With
End Sub

Private Sub Cari_LPD_LPA_Payment_2()
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    
    CMDSQL = "select paydate,payment from tbllunas where custid='"
    CMDSQL = CMDSQL + Trim(LvPTP.SelectedItem.SubItems(2)) + "' order by paydate desc limit 1 "
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        With FrmViewPTP
            If M_objrs.RecordCount > 0 Then
                .TxtLPDPayment.text = IIf(IsNull(M_objrs("paydate")), "", Format(M_objrs("paydate"), "yyyy-mm-dd"))
                .TxtLPAPayment.Value = IIf(IsNull(M_objrs("payment")), "0", M_objrs("payment"))
                LpdPayment = "'" + TxtLPDPayment.text + "'"
            Else
                LpdPayment = "null"
                TxtLPDPayment = ""
                .TxtLPAPayment.Value = "0"
            End If
        End With
    Set M_objrs = Nothing
End Sub

Private Function Ceiling(number As Double) As Long
    Ceiling = -Int(-number)
End Function


Private Sub TxtCustid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdCari_Click
    End If
End Sub

Private Sub TxtNama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdCari_Click
    End If
End Sub
'----------------------------221012 App By VP ----------------------------------------------------------
Private Sub BikinCPA_AppVP(K As Integer)
    Dim CMDSQL As String
    Dim Remarks As String
    Dim M_Objrs_Cek_Type As ADODB.Recordset
    Dim TypeAcc As String
    
    TypeAcc = ""

    '@@13022013 Cek type account dulu nih .. pil/card
    CMDSQL = "select acc_type from mgm where custid='"
    CMDSQL = CMDSQL & CStr(LvPTP.ListItems(K).SubItems(2)) & "'"
    Set M_Objrs_Cek_Type = New ADODB.Recordset
    M_Objrs_Cek_Type.CursorLocation = adUseClient
    M_Objrs_Cek_Type.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_Objrs_Cek_Type.RecordCount > 0 Then
        TypeAcc = IIf(IsNull(M_Objrs_Cek_Type("acc_type")), "", M_Objrs_Cek_Type("acc_type"))
    End If
    
    Set M_Objrs_Cek_Type = Nothing
    
    
    
    Call Cari_LPD_LPA_Payment(K)
    
    CMDSQL = "insert into tblcpa (dpropsal,vcustid,vproduct,nttlpayment,ndownpay,"
    CMDSQL = CMDSQL + "ncharge,ndiscountamt,vosbalance,vosprincipal,dtglinsert,vcustname,vjust,"
    CMDSQL = CMDSQL + "nbalance,nprincipal,nperiod,chkfaxed,chkwentalking,chkktp,"
    CMDSQL = CMDSQL + "chksup,chkbillings,chkothers,lpd_from_payment,lpa_from_payment,"
    CMDSQL = CMDSQL + "f_system,dob,status_ptp,ketother "
    
    '@@19062012 Jika Status PTP DISCON Catat Approvenya
    If Trim(UCase(LvPTP.ListItems(K).SubItems(1))) = "PTP DISCOUNT" Then
        CMDSQL = CMDSQL + ",tglapprove,sts_approve,approve_by,logapprove_by "
    End If
    
    'Catet Juga yang PTP No Discon 20062012
    If Trim(UCase(LvPTP.ListItems(K).SubItems(1))) = "PTP NO DISCOUNT" Then
        CMDSQL = CMDSQL + ",tglapprove,sts_approve,approve_by,logapprove_by "
    End If
    
    '@@16-07-2012 Buat Catet Payment Handle
    CMDSQL = CMDSQL + " ,vpaymenthandle,voccupation,vreason "
    
    CMDSQL = CMDSQL + ") values ("
    CMDSQL = CMDSQL + "now(),'"
    CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).SubItems(2)) + "','"
    CMDSQL = CMDSQL + TypeAcc + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(15)), "0", Replace(LvPTP.ListItems(K).SubItems(15), ",", ""))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(16)), "0", Replace(LvPTP.ListItems(K).SubItems(16), ",", ""))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(17)), "0", Replace(LvPTP.ListItems(K).SubItems(17), ",", ""))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(18)), "0", Replace(LvPTP.ListItems(K).SubItems(18), ",", ""))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(19)), "", LvPTP.ListItems(K).SubItems(19))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(20)), "", LvPTP.ListItems(K).SubItems(20))) + "',"
    CMDSQL = CMDSQL + "now(),'"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(3)), "", LvPTP.ListItems(K).SubItems(3))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(21)), "", LvPTP.ListItems(K).SubItems(21))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(12)), "0", Replace(LvPTP.ListItems(K).SubItems(12), ",", ""))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(11)), "0", Replace(LvPTP.ListItems(K).SubItems(11), ",", ""))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(8)), "", LvPTP.ListItems(K).SubItems(8))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(22)), "", LvPTP.ListItems(K).SubItems(22))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(23)), "", LvPTP.ListItems(K).SubItems(23))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(24)), "", LvPTP.ListItems(K).SubItems(24))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(25)), "", LvPTP.ListItems(K).SubItems(25))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(26)), "", LvPTP.ListItems(K).SubItems(26))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(27)), "", LvPTP.ListItems(K).SubItems(27))) + "',"
    CMDSQL = CMDSQL + IIf(TxtLPDPayment.text = "", "null", "'" + TxtLPDPayment.text + "'") + ",'"
    CMDSQL = CMDSQL + CStr(TxtLPAPayment.Value) + "','1',"
    '@@20062012 Tambahkan DOB dan Status PTP
    CMDSQL = CMDSQL + IIf(LvPTP.ListItems(K).SubItems(29) = "", "null", "'" + LvPTP.ListItems(K).SubItems(29) + "'")
    CMDSQL = CMDSQL + ",'" + LvPTP.ListItems(K).SubItems(1) + "',' "
    '@@21062012 Tambahkan Keterangan Other
    CMDSQL = CMDSQL + IIf(IsNull(LvPTP.ListItems(K).SubItems(30)), "", LvPTP.ListItems(K).SubItems(30)) + "' "
    
    '@@19062012 Buat nyatet approvenya
     If Trim(UCase(LvPTP.ListItems(K).SubItems(1))) = "PTP DISCOUNT" Then
        CMDSQL = CMDSQL + ",now(),'1','"
        CMDSQL = CMDSQL + Trim(CmbApprove.text) + "','"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "'"
     End If
     
     'Buat nyatet yang jenisnya PTP NO Discount.
     If Trim(UCase(LvPTP.ListItems(K).SubItems(1))) = "PTP NO DISCOUNT" Then
        CMDSQL = CMDSQL + ",now(),'1','"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "','"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "'"
     End If
    
    CMDSQL = CMDSQL + ",'"
    CMDSQL = CMDSQL + IIf(IsNull(LvPTP.ListItems(K).SubItems(31)), "", LvPTP.ListItems(K).SubItems(31)) + "','"
    
    CMDSQL = CMDSQL + IIf(IsNull(LvPTP.ListItems(K).SubItems(32)), "", LvPTP.ListItems(K).SubItems(32)) + "','"
    CMDSQL = CMDSQL + IIf(IsNull(LvPTP.ListItems(K).SubItems(33)), "", LvPTP.ListItems(K).SubItems(33)) + "')"
    DoEvents
    M_OBJCONN.Execute CMDSQL
    
    '@@19062012 Bikin Remarks untuk CPA
     '@@11092012 Tulis Remarks baik untuk yang ptp discon/no discon
     If Trim(UCase(LvPTP.ListItems(K).SubItems(1))) = "PTP DISCOUNT" Then
        Remarks = "PtpDisc-"
     Else
        Remarks = "PTPNoDisc-"
     End If
        Remarks = Remarks + "CPA Pengajuan Ke:" + "Pak Hamanto " + "-"
        Remarks = Remarks + "Ttl.Pymt: Rp." + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(7)), "", LvPTP.ListItems(K).SubItems(7))) + " -"
        Remarks = Remarks + "Instl: " + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(8)), "", LvPTP.ListItems(K).SubItems(8))) + " -"
        Remarks = Remarks + "From Bal.: Rp." + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(12)), "", LvPTP.ListItems(K).SubItems(12))) + " -"
        Remarks = Remarks + "From Prin.: Rp." + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(14)), "", LvPTP.ListItems(K).SubItems(14))) + " -"
        Remarks = Remarks + "%Balance: " + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(19)), "", LvPTP.ListItems(K).SubItems(19))) + "% -"
        Remarks = Remarks + "%Principal: " + CStr(IIf(IsNull(LvPTP.ListItems(K).SubItems(20)), "", LvPTP.ListItems(K).SubItems(20))) + "% #USER LOG:" + MDIForm1.TxtUsername.text
        
        CMDSQL = "insert into mgm_hst (custid, agent, products, "
        CMDSQL = CMDSQL + "hst,user_log) values ('"
        CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).SubItems(2)) + "','"
        CMDSQL = CMDSQL + LvPTP.ListItems(K).SubItems(28) + "','"
        CMDSQL = CMDSQL + "Collection" + "','"
        CMDSQL = CMDSQL + Remarks + "','"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "')"
        M_OBJCONN.Execute CMDSQL
    
    
    '@@25072012,Update yang approve dan tanggal proposalnya di tabel tblsendptp jika PTP discount
    If Trim(UCase(LvPTP.ListItems(K).SubItems(1))) = "PTP DISCOUNT" Then
        CMDSQL = "update tblsendptp set tgl_proposal=now(), approve_by='To Be Approved By Pak Hamanto',"
        CMDSQL = CMDSQL + " log_approve='"
        'CMDSQL = CMDSQL + CStr(Trim(CmbApprove.Text)) + "', log_approve='"
        CMDSQL = CMDSQL + CStr(Trim(MDIForm1.TxtUsername.text)) + "', sts_app_vp='1' "
        CMDSQL = CMDSQL + " where id='"
        CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).text) + "'"
        M_OBJCONN.Execute CMDSQL
    End If
    
    If Trim(UCase(LvPTP.ListItems(K).SubItems(1))) = "PTP NO DISCOUNT" Then
        CMDSQL = "update tblsendptp set tgl_proposal=now(), approve_by='"
        CMDSQL = CMDSQL + CStr(Trim(MDIForm1.TxtUsername.text)) + "', log_approve='"
        CMDSQL = CMDSQL + CStr(Trim(MDIForm1.TxtUsername.text)) + "', sts_app_vp='1' "
        CMDSQL = CMDSQL + " where id='"
        CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).text) + "'"
        M_OBJCONN.Execute CMDSQL
    End If
End Sub

Private Sub KirimPesan_AppVP(K As Integer)
    Dim CMDSQL As String
    Dim Remarks As String
    Dim M_objrs As ADODB.Recordset
    
    Remarks = "Pembuatan PTP untuk custid: " & LvPTP.ListItems(K).SubItems(2) & " sedang dalam proses pengajuan ke Pak Hamanto!"
    
    CMDSQL = "insert into msgtbl "
    CMDSQL = CMDSQL + "( recipient, datetime, sender, sentfrom, msg) values ('"
    CMDSQL = CMDSQL + LvPTP.ListItems(K).SubItems(28) + "','"
    CMDSQL = CMDSQL + Format(Now(), "yyyymmdd") + "','"
    CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "','"
    CMDSQL = CMDSQL + CStr(MDIForm1.Winsock1.LocalIP) + "','"
    CMDSQL = CMDSQL + Remarks + "')"
    M_OBJCONN.Execute CMDSQL
        
        
    '@@19072012 Kirim Pesan Buat Ke TL
    'Cari Nama TLNYA
    CMDSQL = "select team from usertbl where userid='"
    CMDSQL = CMDSQL + CStr(Trim(LvPTP.ListItems(K).SubItems(28))) + "' "
    CMDSQL = CMDSQL + " and team is not null "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
        CMDSQL = "insert into msgtbl "
        CMDSQL = CMDSQL + "(recipient, datetime, sender, sentfrom, msg) values ('"
        CMDSQL = CMDSQL + CStr(Trim(M_objrs("team"))) + "','"
        CMDSQL = CMDSQL + Format(Now(), "yyyymmdd") + "','"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "','"
        CMDSQL = CMDSQL + CStr(MDIForm1.Winsock1.LocalIP) + "','"
        CMDSQL = CMDSQL + Remarks + "')"
        M_OBJCONN.Execute CMDSQL
    End If
    
    Set M_objrs = Nothing
End Sub

'@@221012 Buat Approve Hamanto ---------------------------------------------------------------------------
Private Sub HeaderAppHamanto()
    LvPTP.ColumnHeaders.CLEAR
    With LvHamanto.ColumnHeaders
        .ADD 1, , "ID", 500
        .ADD 2, , "Jenis PTP", 1000
        .ADD 3, , "Custid", 2000
        .ADD 4, , "Nama CH", 3000
        .ADD 5, , "Status", 2000
        .ADD 6, , "Tanggal Approve", 2000
        .ADD 7, , "Tgl.Payment Effective", 2500
        .ADD 8, , "Total Amount", 1000
        .ADD 9, , "Tenor", 700
        .ADD 10, , "Pembayaran Via", 2000
        .ADD 11, , "Tgl.Tagih", 1500
        .ADD 12, , "Principal", 1000
        .ADD 13, , "Balance", 1000
        .ADD 14, , "Pembayaran Awal", 2000
        .ADD 15, , "Principal", 2000
        .ADD 16, , "Total Payment", 2000
        .ADD 17, , "Down Payment", 2000
        .ADD 18, , "Charge", 2000
        .ADD 19, , "Discount", 2000
        .ADD 20, , "From o/s balance %", 2000
        .ADD 21, , "Principal %", 2000
        .ADD 22, , "Justtification", 2000
        .ADD 23, , "Fax", 800
        .ADD 24, , "When Talking Surlun", 800
        .ADD 25, , "KTP", 800
        .ADD 26, , "Surper", 800
        .ADD 27, , "Billing", 800
        .ADD 28, , "Other", 800
        .ADD 29, , "Agent", 800
        .ADD 30, , "DOB", 1000
        .ADD 31, , "Ket.Other", 1000
        
        '@@ 16-07-2012 Tambahan Payment Handle
        .ADD 32, , "Payment Handle", 2000
        
        '@@17-07-2012 Tambahan Occupation dan Reason
        .ADD 33, , "Occupation", 2000
        .ADD 34, , "Reason", 2000
    End With
End Sub

Private Sub BikinCPA_Hamanto(K As Integer)
    Dim CMDSQL As String
    Dim Remarks As String
    Dim M_Objrs_Cek_Type As ADODB.Recordset
    Dim TypeAcc As String
    
    TypeAcc = ""

    '@@13022013 Cek type account dulu nih .. pil/card
    CMDSQL = "select acc_type from mgm where custid='"
    CMDSQL = CMDSQL & CStr(LvPTP.ListItems(K).SubItems(2)) & "'"
    Set M_Objrs_Cek_Type = New ADODB.Recordset
    M_Objrs_Cek_Type.CursorLocation = adUseClient
    M_Objrs_Cek_Type.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    If M_Objrs_Cek_Type.RecordCount > 0 Then
        TypeAcc = IIf(IsNull(M_Objrs_Cek_Type("acc_type")), "", M_Objrs_Cek_Type("acc_type"))
    End If
    
    Set M_Objrs_Cek_Type = Nothing
    
    
    Call Cari_LPD_LPA_Payment_Hamanto(K)
    
    
    CMDSQL = "insert into tblcpa (dpropsal,vcustid,vproduct,nttlpayment,ndownpay,"
    CMDSQL = CMDSQL + "ncharge,ndiscountamt,vosbalance,vosprincipal,dtglinsert,vcustname,vjust,"
    CMDSQL = CMDSQL + "nbalance,nprincipal,nperiod,chkfaxed,chkwentalking,chkktp,"
    CMDSQL = CMDSQL + "chksup,chkbillings,chkothers,lpd_from_payment,lpa_from_payment,"
    CMDSQL = CMDSQL + "f_system,dob,status_ptp,ketother "
    
    '@@19062012 Jika Status PTP DISCON Catat Approvenya
    If Trim(UCase(LvHamanto.ListItems(K).SubItems(1))) = "PTP DISCOUNT" Then
        CMDSQL = CMDSQL + ",tglapprove,sts_approve,approve_by,logapprove_by "
    End If
    
    'Catet Juga yang PTP No Discon 20062012
    If Trim(UCase(LvHamanto.ListItems(K).SubItems(1))) = "PTP NO DISCOUNT" Then
        CMDSQL = CMDSQL + ",tglapprove,sts_approve,approve_by,logapprove_by "
    End If
    
    '@@16-07-2012 Buat Catet Payment Handle
    CMDSQL = CMDSQL + " ,vpaymenthandle,voccupation,vreason "
    
    CMDSQL = CMDSQL + ") values ('"
    'Cmdsql = Cmdsql + "now(),'"
    CMDSQL = CMDSQL + Format(TxtTglApprove.Value, "yyyy-mm-dd") + "','"
    
    CMDSQL = CMDSQL + CStr(LvHamanto.ListItems(K).SubItems(2)) + "','"
    CMDSQL = CMDSQL + TypeAcc + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(15)), "0", Replace(LvHamanto.ListItems(K).SubItems(15), ",", ""))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(16)), "0", Replace(LvHamanto.ListItems(K).SubItems(16), ",", ""))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(17)), "0", Replace(LvHamanto.ListItems(K).SubItems(17), ",", ""))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(18)), "0", Replace(LvHamanto.ListItems(K).SubItems(18), ",", ""))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(19)), "", LvHamanto.ListItems(K).SubItems(19))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(20)), "", LvHamanto.ListItems(K).SubItems(20))) + "',"
    CMDSQL = CMDSQL + "now(),'"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(3)), "", LvHamanto.ListItems(K).SubItems(3))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(21)), "", LvHamanto.ListItems(K).SubItems(21))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(12)), "0", Replace(LvHamanto.ListItems(K).SubItems(12), ",", ""))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(11)), "0", Replace(LvHamanto.ListItems(K).SubItems(11), ",", ""))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(8)), "", LvHamanto.ListItems(K).SubItems(8))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(22)), "", LvHamanto.ListItems(K).SubItems(22))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(23)), "", LvHamanto.ListItems(K).SubItems(23))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(24)), "", LvHamanto.ListItems(K).SubItems(24))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(25)), "", LvHamanto.ListItems(K).SubItems(25))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(26)), "", LvHamanto.ListItems(K).SubItems(26))) + "','"
    CMDSQL = CMDSQL + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(27)), "", LvHamanto.ListItems(K).SubItems(27))) + "',"
    CMDSQL = CMDSQL + IIf(TxtLPDPayment.text = "", "null", "'" + TxtLPDPayment.text + "'") + ",'"
    CMDSQL = CMDSQL + CStr(TxtLPAPayment.Value) + "','1',"
    '@@20062012 Tambahkan DOB dan Status PTP
    CMDSQL = CMDSQL + IIf(LvHamanto.ListItems(K).SubItems(29) = "", "null", "'" + LvHamanto.ListItems(K).SubItems(29) + "'")
    CMDSQL = CMDSQL + ",'" + LvHamanto.ListItems(K).SubItems(1) + "',' "
    '@@21062012 Tambahkan Keterangan Other
    CMDSQL = CMDSQL + IIf(IsNull(LvHamanto.ListItems(K).SubItems(30)), "", LvHamanto.ListItems(K).SubItems(30)) + "' "
    
    '@@19062012 Buat nyatet approvenya
     If Trim(UCase(LvHamanto.ListItems(K).SubItems(1))) = "PTP DISCOUNT" Then
        CMDSQL = CMDSQL + ",'" + Format(TxtTglApprove.Value, "yyyy-mm-dd") + "','1','"
        'Cmdsql = Cmdsql + ",now(),'1','"
        'CMDSQL = CMDSQL + Trim(CmbApprove.Text) + "','"
        CMDSQL = CMDSQL + "Hamanto" + "','"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "'"
     End If
     
     'Buat nyatet yang jenisnya PTP NO Discount.
     If Trim(UCase(LvHamanto.ListItems(K).SubItems(1))) = "PTP NO DISCOUNT" Then
        CMDSQL = CMDSQL + ",'" + Format(TxtTglApprove.Value, "yyyy-mm-dd") + "','1','"
        'Cmdsql = Cmdsql + ",now(),'1','"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "','"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "'"
     End If
    
    CMDSQL = CMDSQL + ",'"
    CMDSQL = CMDSQL + IIf(IsNull(LvHamanto.ListItems(K).SubItems(31)), "", LvHamanto.ListItems(K).SubItems(31)) + "','"
    
    CMDSQL = CMDSQL + IIf(IsNull(LvHamanto.ListItems(K).SubItems(32)), "", LvHamanto.ListItems(K).SubItems(32)) + "','"
    CMDSQL = CMDSQL + IIf(IsNull(LvHamanto.ListItems(K).SubItems(33)), "", LvHamanto.ListItems(K).SubItems(33)) + "')"
    DoEvents
    M_OBJCONN.Execute CMDSQL
    
    '@@19062012 Bikin Remarks untuk CPA
     '@@11092012 Tulis Remarks baik untuk yang ptp discon/no discon
     If Trim(UCase(LvHamanto.ListItems(K).SubItems(1))) = "PTP DISCOUNT" Then
        Remarks = "PtpDisc-"
     Else
        Remarks = "PTPNoDisc-"
     End If
        Remarks = Remarks + "App By:" + "Pak Hamanto" + "-"
        Remarks = Remarks + "Ttl.Pymt: Rp." + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(7)), "", LvHamanto.ListItems(K).SubItems(7))) + " -"
        Remarks = Remarks + "Instl: " + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(8)), "", LvHamanto.ListItems(K).SubItems(8))) + " -"
        Remarks = Remarks + "From Bal.: Rp." + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(12)), "", LvHamanto.ListItems(K).SubItems(12))) + " -"
        Remarks = Remarks + "From Prin.: Rp." + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(14)), "", LvHamanto.ListItems(K).SubItems(14))) + " -"
        Remarks = Remarks + "%Balance: " + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(19)), "", LvHamanto.ListItems(K).SubItems(19))) + "% -"
        Remarks = Remarks + "%Principal: " + CStr(IIf(IsNull(LvHamanto.ListItems(K).SubItems(20)), "", LvHamanto.ListItems(K).SubItems(20))) + "% #USER LOG:" + MDIForm1.TxtUsername.text
        
        CMDSQL = "insert into mgm_hst (custid, agent, products, "
        CMDSQL = CMDSQL + "hst,user_log) values ('"
        CMDSQL = CMDSQL + CStr(LvHamanto.ListItems(K).SubItems(2)) + "','"
        CMDSQL = CMDSQL + LvHamanto.ListItems(K).SubItems(28) + "','"
        CMDSQL = CMDSQL + "Collection" + "','"
        CMDSQL = CMDSQL + Remarks + "','"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "')"
        M_OBJCONN.Execute CMDSQL
    
    
    '@@25072012,Update yang approve dan tanggal proposalnya di tabel tblsendptp jika PTP discount
    If Trim(UCase(LvHamanto.ListItems(K).SubItems(1))) = "PTP DISCOUNT" Then
        'Cmdsql = "update tblsendptp set tgl_proposal=now(), approve_by='"
        CMDSQL = "update tblsendptp set approve_by='Hamanto', log_approve='"
        'CMDSQL = CMDSQL + CStr(Trim(CmbApprove.Text)) + "', log_approve='"
        CMDSQL = CMDSQL + CStr(Trim(MDIForm1.TxtUsername.text)) + "',tgl_approve_vp='"
        CMDSQL = CMDSQL + Format(TxtTglApprove.Value, "yyyy-mm-dd") + "',tgl_proposal='"
        CMDSQL = CMDSQL + Format(TxtTglApprove.Value, "yyyy-mm-dd") + "' "
        CMDSQL = CMDSQL + " where id='"
        CMDSQL = CMDSQL + CStr(LvHamanto.ListItems(K).text) + "'"
        M_OBJCONN.Execute CMDSQL
    End If
    
    If Trim(UCase(LvHamanto.ListItems(K).SubItems(1))) = "PTP NO DISCOUNT" Then
        'Cmdsql = "update tblsendptp set tgl_proposal=now(), approve_by='"
        CMDSQL = "update tblsendptp set approve_by='"
        CMDSQL = CMDSQL + CStr(Trim(MDIForm1.TxtUsername.text)) + "', log_approve='"
        CMDSQL = CMDSQL + CStr(Trim(MDIForm1.TxtUsername.text)) + "',tgl_approve_vp='"
        CMDSQL = CMDSQL + Format(TxtTglApprove.Value, "yyyy-mm-dd") + "',tgl_proposal=now() "
        CMDSQL = CMDSQL + " where id='"
        CMDSQL = CMDSQL + CStr(LvPTP.ListItems(K).text) + "' "
        M_OBJCONN.Execute CMDSQL
    End If
End Sub

Private Sub Cari_LPD_LPA_Payment_Hamanto(K As Integer)
    Dim CMDSQL As String
    Dim M_objrs As ADODB.Recordset
    
    StatusPTP = ""
    TxtLPDPayment.text = ""
    TxtLPAPayment.Value = "0"
    
    CMDSQL = "select paydate,payment from tbllunas where custid='"
    CMDSQL = CMDSQL + Trim(LvHamanto.ListItems(K).SubItems(2)) + "' order by paydate desc limit 1 "
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_objrs.RecordCount > 0 Then
            StatusPTP = "PTP-POP"
            TxtLPDPayment.text = IIf(IsNull(M_objrs("paydate")), "", Format(M_objrs("paydate"), "yyyy-mm-dd"))
            TxtLPAPayment.Value = IIf(IsNull(M_objrs("payment")), "0", M_objrs("payment"))
            LpdPayment = "'" + TxtLPDPayment.text + "'"
        Else
            StatusPTP = "PTP-NEW"
            'LpdPayment = "null"
            TxtLPDPayment.text = ""
            TxtLPAPayment.Value = "0"
        End If
    Set M_objrs = Nothing
End Sub


Private Sub BikinPTP_Hamanto(K As Integer)
    Dim CMDSQL As String
    Dim i As Integer
    Dim M_Objrs_Cek_Tgl As ADODB.Recordset
    
    
    bcekptp = True
    
        'Jika Tenor=1
        If Val(LvHamanto.ListItems(K).SubItems(8)) = 1 Then
                  
                '@@14-04-2012 Cek Data
                CMDSQL = "select * from tblnegoptp where custid='"
                CMDSQL = CMDSQL + CStr(LvHamanto.ListItems(K).SubItems(2)) + "' and date(promisedate)='"
                CMDSQL = CMDSQL + CStr(LvHamanto.ListItems(K).SubItems(6)) + "'"
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        CMDSQL = "delete from tblnegoptp where id='"
                        CMDSQL = CMDSQL + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.Execute CMDSQL
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
                  
            jatuhtempo = LvHamanto.ListItems(K).SubItems(6)
            CMDSQL = "INSERT INTO TblNegoPTP "
            CMDSQL = CMDSQL + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
            CMDSQL = CMDSQL + "VALUES "
            CMDSQL = CMDSQL + "('" + CStr(LvHamanto.ListItems(K).SubItems(2)) + "', "
            CMDSQL = CMDSQL + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
            CMDSQL = CMDSQL + "" + CStr(Replace(LvHamanto.ListItems(K).SubItems(13), ",", "")) + " , "
            CMDSQL = CMDSQL + "now(), "
            CMDSQL = CMDSQL + "'IPO')"
            M_OBJCONN.Execute CMDSQL
            
            
            ' isi ke tbl log_ptp
            CMDSQL = "INSERT INTO tblnegoptp_log "
            CMDSQL = CMDSQL + "(custid,PromiseDate, Promisepay,tglInput,agent,stsacc) "
            CMDSQL = CMDSQL + "VALUES "
            CMDSQL = CMDSQL + "('" + CStr(LvHamanto.ListItems(K).SubItems(2)) + "', "
            CMDSQL = CMDSQL + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
            CMDSQL = CMDSQL + "" + CStr(Replace(LvHamanto.ListItems(K).SubItems(13), ",", "")) + " , "
            CMDSQL = CMDSQL + "now(), "
            CMDSQL = CMDSQL + "'" + CStr(LvHamanto.ListItems(K).SubItems(28)) + "','P')"
            M_OBJCONN.Execute CMDSQL
                
        Else
            'Untuk Tenor yang lebih dari 1
                        
                'Hapus Reserved Data
                CMDSQL = "delete from tblreserve where custid='"
                CMDSQL = CMDSQL + CStr(LvHamanto.ListItems(K).SubItems(2)) + "'"
                M_OBJCONN.Execute CMDSQL
                        
                jatuhtempo = CStr(LvHamanto.ListItems(K).SubItems(6))
            
                '@@14-04-2012 Cek Data
                CMDSQL = "select * from tblnegoptp where custid='"
                CMDSQL = CMDSQL + CStr(LvHamanto.ListItems(K).SubItems(2)) + "' and date(promisedate)='"
                CMDSQL = CMDSQL + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "'"
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        CMDSQL = "delete from tblnegoptp where id='"
                        CMDSQL = CMDSQL + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.Execute CMDSQL
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
            
            CMDSQL = "INSERT INTO TblNegoPTP "
            CMDSQL = CMDSQL + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
            CMDSQL = CMDSQL + "VALUES "
            CMDSQL = CMDSQL + "('" + CStr(LvHamanto.ListItems(K).SubItems(2)) + "', "
            CMDSQL = CMDSQL + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
            CMDSQL = CMDSQL + "" + CStr(Replace(LvHamanto.ListItems(K).SubItems(13), ",", "")) + " , "
            CMDSQL = CMDSQL + "now(), "
            CMDSQL = CMDSQL + "'IPO')"
            M_OBJCONN.Execute CMDSQL
            
            
            'isi ke tbl log_ptp
            CMDSQL = "INSERT INTO tblnegoptp_log "
            CMDSQL = CMDSQL + "(custid,PromiseDate, Promisepay,tglInput,agent,stsacc) "
            CMDSQL = CMDSQL + "VALUES "
            CMDSQL = CMDSQL + "('" + CStr(LvHamanto.ListItems(K).SubItems(2)) + "', "
            CMDSQL = CMDSQL + "'" + CStr(Format(jatuhtempo, "yyyy-mm-dd")) + "', "
            CMDSQL = CMDSQL + "" + CStr(Replace(LvHamanto.ListItems(K).SubItems(13), ",", "")) + " , "
            CMDSQL = CMDSQL + "now(), "
            CMDSQL = CMDSQL + "'" + CStr(LvHamanto.ListItems(K).SubItems(28)) + "','P')"
            M_OBJCONN.Execute CMDSQL
                
                
            n = 0
            
            Call HitungInstallmentPtp_Hamanto(K)
            
            For i = 1 To (Val(LvHamanto.ListItems(K).SubItems(8)) - 1)
                    n = n + 1
                    'JMLPAY = ((.TxtPayment - txtPembayaranAwal.Value) - PaymentTenor) / (.txttenor.Value - 1)
                    JmlPay = PaymentTenor
                    Vrdate = DateAdd("m", n, Format(LvHamanto.ListItems(K).SubItems(6), "yyyy-mm-dd"))
                    
                '@@14-04-2012 Cek Data
                CMDSQL = "select * from tblreserve where custid='"
                CMDSQL = CMDSQL + CStr(LvHamanto.ListItems(K).SubItems(2)) + "' and date(promisedate)='"
                CMDSQL = CMDSQL + CStr(Format(Vrdate, "yyyy-mm-dd")) + "'"
                Set M_Objrs_Cek_Tgl = New ADODB.Recordset
                M_Objrs_Cek_Tgl.CursorLocation = adUseClient
                M_Objrs_Cek_Tgl.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
                
                If M_Objrs_Cek_Tgl.RecordCount > 0 Then
                    'Jika ada data di tanggal yang akan diinputkan, hapus dulu deh tanggal yg lama
                    While Not M_Objrs_Cek_Tgl.EOF
                        CMDSQL = "delete from tblreserve where id='"
                        CMDSQL = CMDSQL + CStr(M_Objrs_Cek_Tgl("id")) + "'"
                        M_OBJCONN.Execute CMDSQL
                        M_Objrs_Cek_Tgl.MoveNext
                    Wend
                End If
                Set M_Objrs_Cek_Tgl = Nothing
                    
                    CMDSQL = "INSERT INTO tblreserve "
                    CMDSQL = CMDSQL + "(CUSTID,PromiseDate, Promisepay,Inputdate,Type) "
                    CMDSQL = CMDSQL + "VALUES "
                    CMDSQL = CMDSQL + "('" + CStr(LvHamanto.ListItems(K).SubItems(2)) + "', "
                    CMDSQL = CMDSQL + "'" + CStr(Format(Vrdate, "yyyy-mm-dd")) + "', "
                    CMDSQL = CMDSQL + "" + CStr(JmlPay) + " , "
                    CMDSQL = CMDSQL + "now(), "
                    CMDSQL = CMDSQL + "'IPO')"
                    M_OBJCONN.Execute CMDSQL
                    

                    
                    CMDSQL = "INSERT INTO TblNegoptp_log "
                    CMDSQL = CMDSQL + "(custid,PromiseDate, Promisepay,tglinput,agent,stsacc) "
                    CMDSQL = CMDSQL + "VALUES "
                    CMDSQL = CMDSQL + "('" + CStr(LvHamanto.ListItems(K).SubItems(2)) + "', "
                    CMDSQL = CMDSQL + "'" + CStr(Format(Vrdate, "yyyy-mm-dd")) + "', "
                    CMDSQL = CMDSQL + "" + CStr(JmlPay) + " , "
                    CMDSQL = CMDSQL + "now(), "
                    CMDSQL = CMDSQL + "'" + CStr(LvHamanto.ListItems(K).SubItems(28)) + "','R')"
                    M_OBJCONN.Execute CMDSQL
        

            Next i
       End If
    PaymentTenor = 0
End Sub

Private Sub HitungInstallmentPtp_Hamanto(K As Integer)
    Dim installment As Double
    
        If Val(LvHamanto.ListItems(K).SubItems(8)) = 0 Or Val(LvHamanto.ListItems(K).SubItems(8)) = 1 Then
            installment = Val(Replace(LvHamanto.ListItems(K).SubItems(15), ",", "")) / 1
        Else
            installment = (Val(Replace(LvHamanto.ListItems(K).SubItems(15), ",", "")) - Val(Replace(LvHamanto.ListItems(K).SubItems(13), ",", ""))) / (Val(LvHamanto.ListItems(K).SubItems(8)) - 1)
        End If
        PaymentTenor = Ceiling(installment)
End Sub


Private Sub CatetLogApprove_Hamanto(K As Integer)
    Dim CMDSQL As String
        
    CMDSQL = "insert into tblsendptp_log_approve "
    CMDSQL = CMDSQL + "select * from tblsendptp where id='"
    CMDSQL = CMDSQL + CStr(LvHamanto.ListItems(K).text) + "'"
    DoEvents
    M_OBJCONN.Execute CMDSQL
End Sub

Private Sub BikinStatusPTP_Hamanto(K As Integer)
    Dim CMDSQL As String
    Dim Cmdsql_Cek As String
    Dim StatusRemarks As String
    Dim M_Objrs_Cek As ADODB.Recordset
    Dim AmountNew As Double
    
    AmountNew = 0
    
    Cmdsql_Cek = "select * from tblnegoptp where custid='"
    Cmdsql_Cek = Cmdsql_Cek + CStr(LvHamanto.ListItems(K).SubItems(2)) + "' order by id desc limit 1"
    Set M_Objrs_Cek = New ADODB.Recordset
    M_Objrs_Cek.CursorLocation = adUseClient
    M_Objrs_Cek.Open Cmdsql_Cek, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_Objrs_Cek.RecordCount > 0 Then
        AmountNew = Val(IIf(IsNull(M_Objrs_Cek("promisepay")), "0", M_Objrs_Cek("promisepay")))
    Else
       AmountNew = 0
    End If
    
    'Jika StatusPTP=PTP NEW
    If StatusPTP = "PTP-NEW" Then
        Dim M_Objrs_Cek_Status As ADODB.Recordset
        Dim Cmdsql_Cek_status As String
        Dim TglPTPNew As String
        
        'Cari apakah sebelumnya status data=ptp new, jika iya maka tglptpnew tidak usah diupdate
        'Tapi jika status sebelumnya bukan ptp new maka update tglptpnew=now
        Cmdsql_Cek_status = "select * from mgm where custid='"
        Cmdsql_Cek_status = Cmdsql_Cek_status + CStr(LvHamanto.ListItems(K).SubItems(2)) + "'"
        Set M_Objrs_Cek_Status = New ADODB.Recordset
        M_Objrs_Cek_Status.CursorLocation = adUseClient
        M_Objrs_Cek_Status.Open Cmdsql_Cek_status, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        If M_Objrs_Cek_Status.RecordCount > 0 Then
            If M_Objrs_Cek_Status("tglptpnew") = "" Or IsNull(M_Objrs_Cek_Status("tglptpnew")) = True _
               Or M_Objrs_Cek_Status("tglptpnew") = Empty Then
                TglPTPNew = "now()"
             Else
                TglPTPNew = "'" + CStr(Format(M_Objrs_Cek_Status("tglptpnew"), "yyyy-mm-dd")) + "'"
             End If
        End If
        
        Set M_Objrs_Cek_Status = Nothing
    
        CMDSQL = "update mgm set dateptpnew='"
        CMDSQL = CMDSQL + LvHamanto.ListItems(K).SubItems(6) + "',tgl_tagih='"
        CMDSQL = CMDSQL + LvHamanto.ListItems(K).SubItems(10) + "', amountnew='"
        'CMDSQL = CMDSQL + CStr(Replace(LvPTP.ListItems(K).SubItems(15), ",", "")) + "',tglallptp='"
        'CMDSQL = CMDSQL + CStr(Replace(LvPTP.ListItems(K).SubItems(13), ",", "")) + "',tglallptp='"
        
        '@@20062012, amountnew ambil dari negoptp terakhir aja deh....
        CMDSQL = CMDSQL + CStr(AmountNew) + "',tglallptp='"
        
        CMDSQL = CMDSQL + LvHamanto.ListItems(K).SubItems(6) + "',f_cek_new='PTP-NE',"
        CMDSQL = CMDSQL + "tglincoming=now(),ttlptp='"
        CMDSQL = CMDSQL + CStr(Replace(LvHamanto.ListItems(K).SubItems(15), ",", "")) + "',"
        CMDSQL = CMDSQL + "kethslkerja_new='PTP-NEW',kethslkerjadesc_new='PTP-NEW',ptpvia='"
        CMDSQL = CMDSQL + CStr(LvHamanto.ListItems(K).SubItems(9)) + "',ptpdesc='PTP-NEW', dateptp='"
        CMDSQL = CMDSQL + LvHamanto.ListItems(K).SubItems(6) + "',tglptpnew=" + TglPTPNew
        CMDSQL = CMDSQL + ",tenor='"
        CMDSQL = CMDSQL + CStr(LvHamanto.ListItems(K).SubItems(8)) + "' "
        CMDSQL = CMDSQL + "where custid='"
        CMDSQL = CMDSQL + CStr(LvHamanto.ListItems(K).SubItems(2)) + "'"
        DoEvents
        M_OBJCONN.Execute CMDSQL
        
    End If
    
    If StatusPTP = "PTP-POP" Then
        CMDSQL = "update mgm set dateptp='"
        CMDSQL = CMDSQL + LvHamanto.ListItems(K).SubItems(6) + "',tgl_tagih='"
        CMDSQL = CMDSQL + LvHamanto.ListItems(K).SubItems(10) + "',tglallptp='"
        CMDSQL = CMDSQL + LvHamanto.ListItems(K).SubItems(6) + "',f_cek_new='PTP-PO',"
        CMDSQL = CMDSQL + "tglincoming=now(),ttlptp='"
        CMDSQL = CMDSQL + CStr(Replace(LvHamanto.ListItems(K).SubItems(15), ",", "")) + "',"
        CMDSQL = CMDSQL + "kethslkerja_new='PTP-POP',kethslkerjadesc_new='PTP-POP',ptpvia='"
        CMDSQL = CMDSQL + CStr(LvHamanto.ListItems(K).SubItems(9)) + "',ptpdesc='PTP-POP',amountptp='"
        CMDSQL = CMDSQL + CStr(Replace(LvHamanto.ListItems(K).SubItems(15), ",", "")) + "',tenor='"
        CMDSQL = CMDSQL + CStr(LvHamanto.ListItems(K).SubItems(8)) + "' "
        CMDSQL = CMDSQL + "where custid='"
        CMDSQL = CMDSQL + CStr(LvHamanto.ListItems(K).SubItems(2)) + "'"
        M_OBJCONN.Execute CMDSQL
    End If
    
     '@@19062012 Bikin Remark Status PTP
        StatusRemarks = "PTP Approve by: " & MDIForm1.TxtUsername.text & "/"
        StatusRemarks = StatusRemarks & "Jenis PTP:" & StatusPTP & "/"
        StatusRemarks = StatusRemarks & "Amount PTP:"
        StatusRemarks = StatusRemarks & CStr(Replace(LvHamanto.ListItems(K).SubItems(15), ",", "")) & "/"
        StatusRemarks = StatusRemarks & "PTP Via:" & ""
        StatusRemarks = StatusRemarks & CStr(Replace(LvHamanto.ListItems(K).SubItems(9), ",", "")) & "/"
        StatusRemarks = StatusRemarks & "Date PTP:" & Format(LvHamanto.ListItems(K).SubItems(6), "yyyy-mm-dd")
        
        CMDSQL = "insert into mgm_hst(custid,agent,hst,f_cek_new,user_log) values ('"
        CMDSQL = CMDSQL + CStr(LvHamanto.ListItems(K).SubItems(2)) + "','"
        CMDSQL = CMDSQL + CStr(LvHamanto.ListItems(K).SubItems(28)) + "','"
        CMDSQL = CMDSQL + StatusRemarks + "','"
        CMDSQL = CMDSQL + StatusPTP + "','"
        CMDSQL = CMDSQL + CStr(MDIForm1.TxtUsername.text) + "')"
        M_OBJCONN.Execute CMDSQL
        
   Set M_Objrs_Cek = Nothing
        
End Sub

Private Sub HapusData_Hamanto(K As Integer)
    Dim CMDSQL As String
    
    CMDSQL = "delete from tblsendptp where id='"
    CMDSQL = CMDSQL + CStr(LvHamanto.ListItems(K).text) + "'"
    M_OBJCONN.Execute CMDSQL
End Sub

Private Sub KirimPesan_Hamanto(K As Integer)
    Dim CMDSQL As String
    Dim Remarks As String
    Dim M_objrs As ADODB.Recordset
    
    Remarks = "Pembuatan PTP untuk custid: " & LvHamanto.ListItems(K).SubItems(2) & " telah di approve!"
    
    CMDSQL = "insert into msgtbl "
    CMDSQL = CMDSQL + "( recipient, datetime, sender, sentfrom, msg) values ('"
    CMDSQL = CMDSQL + LvHamanto.ListItems(K).SubItems(28) + "','"
    CMDSQL = CMDSQL + Format(Now(), "yyyymmdd") + "','"
    CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "','"
    CMDSQL = CMDSQL + CStr(MDIForm1.Winsock1.LocalIP) + "','"
    CMDSQL = CMDSQL + Remarks + "')"
    M_OBJCONN.Execute CMDSQL
        
        
    '@@19072012 Kirim Pesan Buat Ke TL
    'Cari Nama TLNYA
    CMDSQL = "select team from usertbl where userid='"
    CMDSQL = CMDSQL + CStr(Trim(LvHamanto.ListItems(K).SubItems(28))) + "' "
    CMDSQL = CMDSQL + " and team is not null "
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    If M_objrs.RecordCount > 0 Then
        CMDSQL = "insert into msgtbl "
        CMDSQL = CMDSQL + "(recipient, datetime, sender, sentfrom, msg) values ('"
        CMDSQL = CMDSQL + CStr(Trim(M_objrs("team"))) + "','"
        CMDSQL = CMDSQL + Format(Now(), "yyyymmdd") + "','"
        CMDSQL = CMDSQL + MDIForm1.TxtUsername.text + "','"
        CMDSQL = CMDSQL + CStr(MDIForm1.Winsock1.LocalIP) + "','"
        CMDSQL = CMDSQL + Remarks + "')"
        M_OBJCONN.Execute CMDSQL
    End If
    
    Set M_objrs = Nothing
End Sub

Private Sub My_Export_Excel()
    Dim a           As Long
    Dim B           As Long
    Dim ExlObj      As Excel.Application
    Dim ListCustId  As String
    Dim rs          As ADODB.Recordset
    Dim iRow        As Integer
    Dim i           As Integer
    
    For K = 1 To LvPTP.ListItems.Count
        If LvPTP.ListItems(K).Checked = True Then
            ListCustId = ListCustId & ",'" & LvPTP.ListItems(K).SubItems(2) & "'"
        End If
    Next K
    
    ListCustId = Mid(ListCustId, 2)
    
    'Strsql = "select custid,vcustname,'" & CmbApprove.Text & "' as Approved,* FROM tblsendptp WHERE custid in (" & listcustid & ")"
'    Strsql = "SELECT * FROM ("
'    Strsql = Strsql + " SELECT custid,vcustname,'" & CmbApprove.Text & "' as Approved,* "
'    Strsql = Strsql + " FROM tblsendptp WHERE custid in (" & listcustid & ")) As a"
'    Strsql = Strsql + " LEFT JOIN (SELECT custid, OpenDate, b_d, FROM mgm WHERE custid in (" & listcustid & ")) As b"
'    Strsql = Strsql + " on a.custid = b.custid"
    
    strsql = "SELECT * FROM ("
    strsql = strsql + " SELECT '" & CmbApprove.text & "' as Approved,* "
    strsql = strsql + " FROM tblsendptp WHERE custid in (" & ListCustId & ")) As a"
    strsql = strsql + " LEFT JOIN (SELECT custid, OpenDate, b_d,LastPay,Pay_Dt FROM mgm WHERE custid in (" & ListCustId & ")) As b"
    strsql = strsql + " on a.custid = b.custid"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic

    Set ExlObj = CreateObject("excel.application")
    ExlObj.Workbooks.ADD
    ExlObj.Visible = True
    
    ExlObj.Range("A1:N1").MergeCells = True
    ExlObj.Range("A2:N2").MergeCells = True
    ExlObj.Range("A4:N4").Font.Bold = True
    
    
    With ExlObj.ActiveSheet
        .Cells(1, 1).Value = "List CPA Approve"
        .Cells(1, 1).Font.Name = "Verdana"
        .Cells(1, 1).Font.Bold = True
        .Cells(2, 1).Value = "Tanggal : " + Format(Now, "dd-mm-yyyy")
        .Cells(2, 1).Font.Name = "Verdana"
        .Cells(2, 1).Font.Bold = True
        .Cells(4, 1).Value = "NO"
        .Cells(4, 2).Value = "CARD NUMBER"
        .Cells(4, 3).Value = "CH NAME"
        .Cells(4, 4).Value = "APPROVED"
        .Cells(4, 5).Value = "ADMIN CREATED"
        .Cells(4, 6).Value = "RECEIVED BY" 'Dikosongkan
        .Cells(4, 7).Value = "ID"""
        .Cells(4, 8).Value = "Jenis PTP"
        .Cells(4, 9).Value = "Custid"
        .Cells(4, 10).Value = "Nama CH"
        .Cells(4, 11).Value = "Status"
        .Cells(4, 12).Value = "Tanggal Approve"
        .Cells(4, 13).Value = "Tgl.Payment Effective"
        .Cells(4, 14).Value = "Total Amount"
        .Cells(4, 15).Value = "Tenor"
        .Cells(4, 16).Value = "Pembayaran Via"
        .Cells(4, 17).Value = "Tgl.Tagih"
        .Cells(4, 18).Value = "Principal"
        .Cells(4, 19).Value = "Balance"
        .Cells(4, 20).Value = "Pembayaran Awal"
        .Cells(4, 21).Value = "Principal"
        .Cells(4, 22).Value = "Total Payment"
        .Cells(4, 23).Value = "Down Payment"
        .Cells(4, 24).Value = "Charge"
        .Cells(4, 25).Value = "Discount"
        .Cells(4, 26).Value = "From o/s balance %"
        .Cells(4, 27).Value = "Principal %"
        .Cells(4, 28).Value = "Justtification"
        .Cells(4, 29).Value = "Fax"
        .Cells(4, 30).Value = "When Talking Surlun"
        .Cells(4, 31).Value = "KTP"
        .Cells(4, 32).Value = "Surper"
        .Cells(4, 33).Value = "Billing"
        .Cells(4, 34).Value = "Other"
        .Cells(4, 35).Value = "Agent"
        .Cells(4, 36).Value = "DOB"
        .Cells(4, 37).Value = "Ket.Other"
        .Cells(4, 38).Value = "Open Date"
        .Cells(4, 39).Value = "WO Date"
        .Cells(4, 40).Value = "LPD"
        .Cells(4, 41).Value = "LPA"
        
        iRow = 4
        If rs.RecordCount > 0 Then
            PB1.Max = rs.RecordCount
            i = 0
            Do Until rs.EOF
                i = i + 1
                iRow = iRow + 1
                PB1.Value = rs.Bookmark
                .Cells(iRow, 1).Value = i
                .Cells(iRow, 2).Value = IIf(IsNull(rs!CustId), "", rs!CustId)
                .Cells(iRow, 3).Value = IIf(IsNull(rs!vcustname), "", rs!vcustname)
                .Cells(iRow, 4).Value = IIf(IsNull(rs!approved), "", rs!approved)
                .Cells(iRow, 5).Value = MDIForm1.TxtUsername.text
                .Cells(iRow, 6).Value = "" 'Dikosongkan
                .Cells(iRow, 7).Value = ""
                .Cells(iRow, 8).Value = IIf(IsNull(rs!jenis_ptp), "", rs!jenis_ptp)
                .Cells(iRow, 9).Value = IIf(IsNull(rs!CustId), "", rs!CustId)
                .Cells(iRow, 10).Value = ""
                .Cells(iRow, 11).Value = IIf(IsNull(rs!STATUS), "", rs!STATUS)
                .Cells(iRow, 12).Value = IIf(IsNull(rs("tgl_approve")), "", Format(rs("tgl_approve"), "yyyy-mm-dd"))
                .Cells(iRow, 13).Value = IIf(IsNull(rs("date_payment_effective")), "", Format(rs("date_payment_effective"), "yyyy-mm-dd"))
                .Cells(iRow, 14).Value = IIf(IsNull(rs!total_amount_deal), "", rs!total_amount_deal)
                .Cells(iRow, 15).Value = IIf(IsNull(rs!Tenor), "", rs!Tenor)
                .Cells(iRow, 16).Value = IIf(IsNull(rs!pembayaran_via), "", rs!pembayaran_via)
                .Cells(iRow, 17).Value = IIf(IsNull(rs("tgl_tagih")), "", Format(rs("tgl_tagih"), "yyyy-mm-dd"))
                .Cells(iRow, 18).Value = IIf(IsNull(rs!Principal), "", rs!Principal)
                .Cells(iRow, 19).Value = IIf(IsNull(rs!balance), "", rs!balance)
                .Cells(iRow, 20).Value = IIf(IsNull(rs!Pembayaran_awal), "", rs!Pembayaran_awal)
                .Cells(iRow, 21).Value = IIf(IsNull(rs!Principal), "", rs!Principal)
                .Cells(iRow, 22).Value = IIf(IsNull(rs!nttlpayment), "", rs!nttlpayment)
                .Cells(iRow, 23).Value = IIf(IsNull(rs!ndownpay), "", rs!ndownpay)
                .Cells(iRow, 24).Value = IIf(IsNull(rs!ncharge), "", rs!ncharge)
                .Cells(iRow, 25).Value = IIf(IsNull(rs!ndiscountamt), "", rs!ndiscountamt)
                .Cells(iRow, 26).Value = IIf(IsNull(rs!vosbalance), "", rs!vosbalance)
                .Cells(iRow, 27).Value = IIf(IsNull(rs!vosprincipal), "", rs!vosprincipal)
                .Cells(iRow, 28).Value = IIf(IsNull(rs!vjust), "", rs!vjust)
                .Cells(iRow, 29).Value = IIf(IsNull(rs!chkfaxed), "", rs!chkfaxed)
                .Cells(iRow, 30).Value = IIf(IsNull(rs!chkwentalking), "", rs!chkwentalking)
                .Cells(iRow, 31).Value = IIf(IsNull(rs!chkKTP), "", rs!chkKTP)
                .Cells(iRow, 32).Value = IIf(IsNull(rs!chksup), "", rs!chksup)
                .Cells(iRow, 33).Value = IIf(IsNull(rs!chkbillings), "", rs!chkbillings)
                .Cells(iRow, 34).Value = IIf(IsNull(rs!chkothers), "", rs!chkothers)
                .Cells(iRow, 35).Value = IIf(IsNull(rs!AGENT), "", rs!AGENT)
                .Cells(iRow, 36).Value = IIf(IsNull(rs("DOB")), "", Format(rs("DOB"), "yyyy-mm-dd"))
                .Cells(iRow, 37).Value = IIf(IsNull(rs!ket_other), "", rs!ket_other)
'                .Cells(iRow, 38).Value = IIf(IsNull(RS!OpenDate), "", RS!OpenDate)
'                .Cells(iRow, 39).Value = IIf(IsNull(RS!b_d), "", RS!b_d)
                .Cells(iRow, 38).Value = cnull(rs!opendate)
                .Cells(iRow, 39).Value = IIf(IsNull(rs!b_d), "", rs!b_d)
                .Cells(iRow, 40).Value = IIf(IsNull(rs("Pay_Dt")), "", Format(rs("Pay_Dt"), "dd-mm-yyyy"))
                .Cells(iRow, 41).Value = IIf(IsNull(rs!lastpay), "", rs!lastpay)


                rs.MoveNext
            Loop
        End If
    
        'OTOMATISASI CELL
        For iColom = 1 To 14
            ExlObj.Cells(4, iColom).EntireColumn.AutoFit
        Next
        
        MsgBox "Data berhasil di download!", vbOKOnly + vbInformation, "Informasi"
        PB1.Value = 0
        Command1.Enabled = True
    
        Set ExlObj = Nothing
        Set rs = Nothing

        'StartMeUp (Txtlocation.Text)
        'FILL COLOR CELL
        'ExlObj.Range(.Cells(NoUrut, 1), .Cells(NoUrut, 7)).Interior.Color = RGB(6, 207, 250)
    End With
End Sub
