VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form_recycle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recycle Data"
   ClientHeight    =   9690
   ClientLeft      =   45
   ClientTop       =   840
   ClientWidth     =   14580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9690
   ScaleWidth      =   14580
   Begin TabDlg.SSTab SSTab1 
      Height          =   8790
      Left            =   0
      TabIndex        =   1
      Top             =   855
      Width           =   14460
      _ExtentX        =   25506
      _ExtentY        =   15505
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Criteria Recycle"
      TabPicture(0)   =   "Form_recycle.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image3(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "SSTab2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Histori"
      TabPicture(1)   =   "Form_recycle.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Image3(0)"
      Tab(1).Control(1)=   "Label1(4)"
      Tab(1).Control(2)=   "ListView1(1)"
      Tab(1).Control(3)=   "txtjmlrow(1)"
      Tab(1).Control(4)=   "CmdSearchBaru(1)"
      Tab(1).ControlCount=   5
      Begin TabDlg.SSTab SSTab2 
         Height          =   4605
         Left            =   90
         TabIndex        =   33
         Top             =   3615
         Width           =   14100
         _ExtentX        =   24871
         _ExtentY        =   8123
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "User                 "
         TabPicture(0)   =   "Form_recycle.frx":0038
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame4"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Detail Customer"
         TabPicture(1)   =   "Form_recycle.frx":0054
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "DataGrid1"
         Tab(1).ControlCount=   1
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   4065
            Left            =   -74910
            TabIndex        =   38
            Top             =   405
            Width           =   13920
            _ExtentX        =   24553
            _ExtentY        =   7170
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Height          =   4155
            Left            =   45
            TabIndex        =   34
            Top             =   360
            Width           =   14010
            Begin VB.CommandButton cmdkeluar 
               BackColor       =   &H00F1E5DB&
               Caption         =   "&Keluar"
               Height          =   495
               Left            =   12450
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   3600
               Width           =   1575
            End
            Begin VB.CommandButton cmdProses 
               BackColor       =   &H00F1E5DB&
               Caption         =   "&PROSES"
               Enabled         =   0   'False
               Height          =   495
               Left            =   10890
               Style           =   1  'Graphical
               TabIndex        =   35
               Top             =   3600
               Width           =   1575
            End
            Begin MSComctlLib.ListView LVRecycle 
               Height          =   3420
               Left            =   90
               TabIndex        =   37
               Top             =   180
               Width           =   13920
               _ExtentX        =   24553
               _ExtentY        =   6033
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   12582912
               BackColor       =   16777215
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
      Begin VB.CommandButton CmdSearchBaru 
         Height          =   360
         Index           =   1
         Left            =   -74910
         Picture         =   "Form_recycle.frx":0070
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   405
         Width           =   1515
      End
      Begin VB.TextBox txtjmlrow 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   1
         Left            =   -62490
         MaxLength       =   20
         TabIndex        =   25
         Top             =   8370
         Width           =   1785
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Informasi data Periode"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Left            =   8280
         TabIndex        =   14
         Top             =   1290
         Width           =   5910
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "New Data"
            Height          =   375
            Left            =   2610
            TabIndex        =   39
            Top             =   1050
            Visible         =   0   'False
            Width           =   1725
         End
         Begin VB.CheckBox cek_otomatis 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Isi jumlah recycle otomatis"
            Height          =   255
            Left            =   150
            TabIndex        =   23
            Top             =   1110
            Width           =   3375
         End
         Begin VB.TextBox txtSudahDistribusi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H0000FF00&
            Height          =   285
            Left            =   4530
            Locked          =   -1  'True
            TabIndex        =   18
            Text            =   "0"
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtSisaCampaign 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   2070
            Locked          =   -1  'True
            TabIndex        =   17
            Text            =   "0"
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtjmlcampaign 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   16
            Text            =   "0"
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtJmlAgent 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4530
            Locked          =   -1  'True
            TabIndex        =   15
            Text            =   "0"
            Top             =   750
            Width           =   975
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Akan di Recycle :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3060
            TabIndex        =   22
            Top             =   360
            Width           =   1515
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Sisa data :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   210
            TabIndex        =   21
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Total data :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   240
            TabIndex        =   20
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Jumlah Agent:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3090
            TabIndex        =   19
            Top             =   690
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   3075
         Left            =   90
         TabIndex        =   5
         Top             =   405
         Width           =   8160
         Begin VB.TextBox txtnokartu 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1530
            TabIndex        =   32
            Top             =   1080
            Width           =   6390
         End
         Begin VB.TextBox txtnama 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1530
            TabIndex        =   30
            Top             =   675
            Width           =   6390
         End
         Begin VB.CommandButton cmdloadcampaign 
            BackColor       =   &H00F1E5DB&
            Caption         =   "&Load Data"
            Height          =   360
            Left            =   3870
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   270
            Width           =   1575
         End
         Begin VB.ComboBox cbotelesales 
            Height          =   315
            Left            =   1530
            TabIndex        =   13
            Top             =   2250
            Width           =   6480
         End
         Begin VB.ComboBox cbosupervisor 
            Height          =   315
            Left            =   1530
            TabIndex        =   11
            Top             =   1890
            Width           =   6480
         End
         Begin VB.ComboBox cmbcampaigncode 
            Height          =   315
            Left            =   1530
            TabIndex        =   7
            Top             =   1485
            Width           =   6480
         End
         Begin VB.TextBox txtno_case 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1530
            TabIndex        =   6
            Top             =   270
            Width           =   2205
         End
         Begin TDBDate6Ctl.TDBDate TdTglCall1 
            Height          =   285
            Left            =   1500
            TabIndex        =   41
            Top             =   2610
            Width           =   1620
            _Version        =   65536
            _ExtentX        =   2857
            _ExtentY        =   503
            Calendar        =   "Form_recycle.frx":065E
            Caption         =   "Form_recycle.frx":0776
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Form_recycle.frx":07E2
            Keys            =   "Form_recycle.frx":0800
            Spin            =   "Form_recycle.frx":085E
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
         Begin TDBDate6Ctl.TDBDate TdTglCall2 
            Height          =   285
            Left            =   3585
            TabIndex        =   42
            Top             =   2610
            Width           =   1395
            _Version        =   65536
            _ExtentX        =   2461
            _ExtentY        =   503
            Calendar        =   "Form_recycle.frx":0886
            Caption         =   "Form_recycle.frx":099E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Form_recycle.frx":0A0A
            Keys            =   "Form_recycle.frx":0A28
            Spin            =   "Form_recycle.frx":0A86
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
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "To "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   8
            Left            =   2985
            TabIndex        =   43
            Top             =   2610
            Width           =   825
         End
         Begin VB.Label Label1 
            BackColor       =   &H00F1E5DB&
            BackStyle       =   0  'Transparent
            Caption         =   "Tgl Call"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   180
            TabIndex        =   40
            Top             =   2640
            Width           =   870
         End
         Begin VB.Label Label1 
            BackColor       =   &H00F1E5DB&
            BackStyle       =   0  'Transparent
            Caption         =   "Region"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   135
            TabIndex        =   31
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label1 
            BackColor       =   &H00F1E5DB&
            BackStyle       =   0  'Transparent
            Caption         =   "Nama"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   135
            TabIndex        =   29
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label1 
            BackColor       =   &H00F1E5DB&
            BackStyle       =   0  'Transparent
            Caption         =   "Agent"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   135
            TabIndex        =   12
            Top             =   2295
            Width           =   870
         End
         Begin VB.Label Label1 
            BackColor       =   &H00F1E5DB&
            BackStyle       =   0  'Transparent
            Caption         =   "Supervisor"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   135
            TabIndex        =   10
            Top             =   1890
            Width           =   870
         End
         Begin VB.Label Label1 
            BackColor       =   &H00F1E5DB&
            BackStyle       =   0  'Transparent
            Caption         =   "Customer ID"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   135
            TabIndex        =   9
            Top             =   315
            Width           =   1455
         End
         Begin VB.Label Label1 
            BackColor       =   &H00F1E5DB&
            BackStyle       =   0  'Transparent
            Caption         =   "Campaign  Code"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   135
            TabIndex        =   8
            Top             =   1530
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Status call yang akan di Recycle"
         Height          =   840
         Left            =   8280
         TabIndex        =   2
         Top             =   405
         Width           =   5910
         Begin VB.ComboBox CmbStatusCall 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "Form_recycle.frx":0AAE
            Left            =   1320
            List            =   "Form_recycle.frx":0AB0
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   345
            Width           =   4230
         End
         Begin VB.Label Label3 
            BackColor       =   &H00F1E5DB&
            BackStyle       =   0  'Transparent
            Caption         =   "Status call :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   135
            TabIndex        =   4
            Top             =   345
            Width           =   1455
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   7440
         Index           =   1
         Left            =   -74865
         TabIndex        =   27
         Top             =   855
         Width           =   14145
         _ExtentX        =   24950
         _ExtentY        =   13123
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
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   4
         Left            =   -63300
         TabIndex        =   28
         Top             =   8415
         Width           =   810
      End
      Begin VB.Image Image3 
         Height          =   18630
         Index           =   0
         Left            =   -74955
         Picture         =   "Form_recycle.frx":0AB2
         Top             =   315
         Width           =   26295
      End
      Begin VB.Image Image3 
         Height          =   18630
         Index           =   2
         Left            =   60
         Picture         =   "Form_recycle.frx":80BC
         Top             =   330
         Width           =   26295
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Recycle Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Index           =   0
      Left            =   615
      TabIndex        =   0
      Top             =   300
      Width           =   3585
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   0
      Left            =   165
      Picture         =   "Form_recycle.frx":F6C6
      Stretch         =   -1  'True
      Top             =   240
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   825
      Left            =   -2070
      Picture         =   "Form_recycle.frx":101D0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   19980
   End
End
Attribute VB_Name = "Form_recycle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub header_distribusi_Recycle()
    LVRecycle.ColumnHeaders.ADD 1, , "No", 5 * TXT
    LVRecycle.ColumnHeaders.ADD 2, , "Kode Spv/Agent", 20 * TXT
    LVRecycle.ColumnHeaders.ADD 3, , "Nama", 31 * TXT
    LVRecycle.ColumnHeaders.ADD 4, , "Jumlah data", 15 * TXT
    LVRecycle.ColumnHeaders.ADD 5, , "Jumlah yang di Recycle", 20 * TXT
    
    ListView1(1).ColumnHeaders.ADD 1, , "No", 5 * TXT
    ListView1(1).ColumnHeaders.ADD 2, , "Nama", 20 * TXT
    ListView1(1).ColumnHeaders.ADD 3, , "  tgl", 31 * TXT
    ListView1(1).ColumnHeaders.ADD 4, , "Jumlah data", 15 * TXT
End Sub

Private Sub cbosupervisor_DropDown()
    load_spv
End Sub

Private Sub cbosupervisor_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cbotelesales_DropDown()
load_telesales
End Sub

Private Sub cek_otomatis_Click()
  Dim m_msgbox As String
  Dim i As Integer
  
  'Mengisi jumlah yang di Recycle sesuai dengan jumlah data yang dimiliki agent
    If cek_otomatis.Value = vbChecked Then
        m_msgbox = MsgBox("Semua data pada semua agent sesuai kriteria status call akan di Recycle! Periksa kembali data per agent yang akan di recycle sebelum menekan tombol proses!", vbOKOnly + vbInformation, "Informasi")
        txtSudahDistribusi.text = 0
        For i = 1 To Val(LVRecycle.ListItems.Count)
            LVRecycle.ListItems.Item(i).SubItems(4) = LVRecycle.ListItems.Item(i).SubItems(3)
            'Jumlah yang di recycle
            txtSudahDistribusi.text = Val(txtSudahDistribusi.text) + Val(LVRecycle.ListItems.Item(i).SubItems(3))
            'Jumlah sisa data
            txtSisaCampaign.text = Val(txtjmlcampaign.text) - Val(txtSudahDistribusi.text)
        Next i
    End If
    
    'Tidak jadi mengeset otomatis
    If cek_otomatis.Value = vbUnchecked Then
        
        For i = 1 To Val(LVRecycle.ListItems.Count)
            LVRecycle.ListItems.Item(i).SubItems(4) = 0
            txtSudahDistribusi.text = 0
            txtSisaCampaign.text = txtjmlcampaign.text
        Next i
    End If

End Sub

Private Sub cmbcampaigncode_DropDown()
    load_campaign
End Sub

Private Sub cmbcampaigncode_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub CmbStatusCall_DropDown()
    load_statuscall
End Sub

Private Sub cmdkeluar_Click()
    Unload Me
End Sub
Private Sub cmdloadcampaign_Click()
Dim list As ListItem
Dim mobjrs2 As New ADODB.Recordset
mwhere = ""
If txtno_case.text <> "" Then
    mwhere = " and custid = '" + txtno_case.text + "'"
End If

If txtnama.text <> "" Then
    mwhere = mwhere + " and name like '%" + txtnama.text + "%'"
End If

If txtnokartu.text <> "" Then
    mwhere = mwhere + " and region= '" + txtnokartu.text + "'"
End If

If Not (TdTglCall1.ValueIsNull) And Not (TdTglCall2.ValueIsNull) Then
    mwhere = mwhere + " and date(tglCALL) between '"
    mwhere = mwhere + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' and '"
    mwhere = mwhere + Format(TdTglCall2.Value, "yyyy-mm-dd") + "'"
End If

If cmbcampaigncode.text <> "" Then
    mwhere = mwhere + " and recsource = '" + cmbcampaigncode.text + "'"
End If

If MDIForm1.txtlevel = "Agent" Then
    mwhere = mwhere + " and agent = '" + MDIForm1.TxtUsername.text + "'"
ElseIf MDIForm1.txtlevel = "Supervisor" Then
    mwhere = mwhere + " and agent in (select  userid  from usertbl where spvcode= '" + MDIForm1.TxtUsername.text + "' and  AKTIF ='1') "
End If

If cbosupervisor.text <> "" Then
intvrl = InStr(1, cbosupervisor, "-", vbTextCompare)
    If intvrl <> 0 Then
       ArrayString = Split(cbosupervisor.text, "-", 2, vbTextCompare)
       getUserid = ArrayString(0)
       getUser_name = ArrayString(1)
    End If
    
    If MDIForm1.txtlevel = "Supervisor" Then
        mwhere = mwhere + " and  agent in (select  userid  from usertbl where spvcode= '" + getUserid + "' and  AKTIF ='1') "
    ElseIf MDIForm1.txtlevel = "Agent" Then
        mwhere = mwhere + " and  agent in (select  userid  from usertbl where spvcode= '" + getUserid + "' and  AKTIF ='1') "
    Else
        mwhere = mwhere + " and  agent in (select  userid  from usertbl where (spvcode= '" + getUserid + "' or userid='" + getUserid + "'  ) and  AKTIF ='1') "
    End If
End If

If cbotelesales.text <> "" Then
    intvrl = InStr(1, cbotelesales, "-", vbTextCompare)
    If intvrl <> 0 Then
       ArrayString = Split(cbotelesales.text, "-", 2, vbTextCompare)
       getUserid = ArrayString(0)
       getUser_name = ArrayString(1)
    End If
    mwhere = mwhere + " and agent = '" + getUserid + "' "
End If

If CmbStatusCall.text <> "" Then
    If CmbStatusCall.text = "New Data" Then
        mwhere = mwhere + " and coalesce(statuscall,'') = '' AND coalesce(F_CEK_NEW,'')='' "
    Else
        mwhere = mwhere + " and statuscall = '" + CmbStatusCall.text + "' "
    End If
End If

sStrsql = "select agent,nama_agent,count(custid) as jmllead  from mgm where name  <>'' AND (statuscall<>'PTP' OR COALESCE(STATUSCALL,'')='') " + mwhere + " group by agent,nama_agent "
sstrSql2 = "select *  from mgm where (statuscall<>'PTP' OR COALESCE(STATUSCALL,'')='')  " + mwhere + "  "

Set mobjrs2 = New ADODB.Recordset
mobjrs2.CursorLocation = adUseClient
mobjrs2.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
no = 0
LVRecycle.ListItems.CLEAR
txtjmlcampaign.text = 0
While Not mobjrs2.EOF
     no = no + 1
     Set list = LVRecycle.ListItems.ADD(, , no)
     list.SubItems(1) = IIf(IsNull(mobjrs2!AGENT), "", mobjrs2!AGENT)
     list.SubItems(2) = IIf(IsNull(mobjrs2!nama_agent), "", mobjrs2!nama_agent)
     list.SubItems(3) = IIf(IsNull(mobjrs2!jmllead), "", mobjrs2!jmllead)
     txtjmlcampaign.text = Val(txtjmlcampaign.text) + Val(IIf(IsNull(mobjrs2!jmllead), "", mobjrs2!jmllead))
     mobjrs2.MoveNext
Wend
Set mobjrs2 = Nothing

Set mobjrs2 = New ADODB.Recordset
mobjrs2.CursorLocation = adUseClient
'mobjrs2.Open sstrSql2, M_OBJCONN, adOpenKeyset, adLockReadOnly
Set DataGrid1.DATASOURCE = mobjrs2
Set mobjrs2 = Nothing

cmdProses.Enabled = True
End Sub

Private Sub CmdProses_Click()
    
Dim list As ListItem
Dim mobjrs2 As New ADODB.Recordset
Dim rs As New ADODB.Recordset
mwhere = ""
sstrSql2 = ""
strsql = ""


If txtno_case.text <> "" Then
    mwhere = " and CUSTID = '" + txtno_case.text + "'"
End If

If txtnama.text <> "" Then
    mwhere = mwhere + " and   name like '%" + txtnama.text + "%'"
End If


If txtnokartu.text <> "" Then
    mwhere = mwhere + " and region ='" + txtnokartu.text + "'"
End If

If Not (TdTglCall1.ValueIsNull) And Not (TdTglCall2.ValueIsNull) Then
    mwhere = mwhere + " and date(tglcall) between '"
    mwhere = mwhere + Format(TdTglCall1.Value, "yyyy-mm-dd") + "' and '"
    mwhere = mwhere + Format(TdTglCall2.Value, "yyyy-mm-dd") + "'"
End If

If cmbcampaigncode.text <> "" Then
    mwhere = mwhere + " and recsource = '" + cmbcampaigncode.text + "'"
End If

If MDIForm1.txtlevel = "Agent" Then
    mwhere = mwhere + " and agent = '" + MDIForm1.TxtUsername.text + "'"
ElseIf MDIForm1.txtlevel = "Supervisor" Then
    mwhere = mwhere + " and agent in (select  userid  from usertbl where spvcode= '" + MDIForm1.TxtUsername.text + "' and  AKTIF ='1') "
Else
    mwhere = mwhere + " and agent in (select  userid  from usertbl where AKTIF ='1') "
End If

If cbosupervisor.text <> "" Then
    intvrl = InStr(1, cbosupervisor, "-", vbTextCompare)
    If intvrl <> 0 Then
       ArrayString = Split(cbosupervisor.text, "-", 2, vbTextCompare)
       getUserid = ArrayString(0)
       getUser_name = ArrayString(1)
    End If
    
    If MDIForm1.txtlevel = "Supervisor" Then
           mwhere = mwhere + " and  agent in (select  userid  from usertbl where spvcode= '" + getUserid + "' and  AKTIF ='1') "
    ElseIf MDIForm1.txtlevel = "Agent" Then
         mwhere = mwhere + " and  agent in (select  userid  from usertbl where spvcode= '" + getUserid + "' and  AKTIF ='1') "
    Else
         mwhere = mwhere + " and  agent in (select  userid  from usertbl where (spvcode= '" + getUserid + "' or userid='" + getUserid + "'  ) and  AKTIF ='1') "
    End If
End If

If cbotelesales.text <> "" Then
    intvrl = InStr(1, cbotelesales, "-", vbTextCompare)
    If intvrl <> 0 Then
       ArrayString = Split(cbotelesales.text, "-", 2, vbTextCompare)
       getUserid = ArrayString(0)
       getUser_name = ArrayString(1)
    End If
    mwhere = mwhere + " and agent = '" + getUserid + "' "
End If

If CmbStatusCall.text <> "" Then
    If CmbStatusCall.text = "New Data" Then
        mwhere = mwhere + " and coalesce(statuscall,'') = '' AND coalesce(F_CEK_NEW,'')='' "
    Else
        mwhere = mwhere + " and statuscall = '" + CmbStatusCall.text + "' "
    End If
End If

If Check1.Value = vbChecked Then
    mwhere = mwhere + " and  coalesce(statuscall,'')='' "
End If

waktu = FungsiWaktuServer
sstrSql2 = "select id  from mgm where name<>''  and  (STATUSCALL <> 'PTP' OR COALESCE(STATUSCALL,'')='')  " + mwhere + "  "
      
    For i = 1 To LVRecycle.ListItems.Count
         If Val(LVRecycle.ListItems(i).SubItems(4)) <> 0 Then
            If MDIForm1.txtlevel = "Agent " Then
               strsql = "insert into  tblrecyle_hst (recycle_id,tblrecyle_no_case,nama ,remarks,tglentry,agent,nmagent,statuscall ,campaign_code,userinput )"
                strsql = strsql + " select id,CUSTID,name,remarks,'" + waktu + "',agent,nama_agent,statuscall,recsource,'" + MDIForm1.txtnama.text + "'"
                strsql = strsql + " from mgm where ID in (" + sstrSql2 + " )"
                strsql = strsql + " and agent='" + LVRecycle.ListItems(i).SubItems(1) + "' limit " + CStr(Val(LVRecycle.ListItems(i).SubItems(4))) + ""
                M_OBJCONN.Execute (strsql)
                
                strsql = "update  mgm set statusccall='',retur=NULL ,remarks =null,agent='" + MDIForm1.TxtUsername + "',nama_agent='" + MDIForm1.txtnama.text + "' , bucket ='R' where custid in (" + sstrSql2 + " )"
                strsql = strsql + " and agent='" + LVRecycle.ListItems(i).SubItems(1) + "' limit " + Val(LVRecycle.ListItems(i).SubItems(4)) + ""
                'M_OBJCONN.Execute (STRSQL)
            ElseIf MDIForm1.txtlevel = "Supervisor" Then
                            
                strsql = "insert into  tblrecyle_hst (recycle_id,tblrecyle_no_case,nama ,remarks,tglentry, agent,nmagent,statuscall ,userinput )"
                strsql = strsql + " select id,CUSTID,name,remarks,'" + waktu + "',agent,nama_agent,statuscall,'" + MDIForm1.txtnama.text + "'"
                strsql = strsql + " from mgm where ID in (" + sstrSql2 + " )"
                strsql = strsql + " and agent='" + LVRecycle.ListItems(i).SubItems(1) + "' AND (STATUSCALL <> 'PTP' OR COALESCE(STATUSCALL,'')='') limit " + CStr(Val(LVRecycle.ListItems(i).SubItems(4))) + ""
                M_OBJCONN.Execute (strsql)
               
                
                strsql = "update  mgm set agent='" + MDIForm1.TxtUsername.text + "',nama_agent='" + MDIForm1.txtnama.text + "', statuscall='',tglcall=null,remarks=null,retur=NULL ,  bucket ='R' where id in (" + sstrSql2 + " "
                strsql = strsql + " and agent='" + LVRecycle.ListItems(i).SubItems(1) + "'  limit " + CStr(Val(LVRecycle.ListItems(i).SubItems(4))) + ") AND (STATUSCALL <> 'PTP' OR COALESCE(STATUSCALL,'')='') "
                M_OBJCONN.Execute (strsql)
                
            Else
                    strsql = "insert into tblrecyle_hst (recycle_id,tblrecyle_no_case,nama ,remarks,tglentry, agent,nmagent,statuscall ,userinput )"
                    strsql = strsql + " select id,CUSTID,name,remarks,'" + waktu + "',agent,nama_agent,statuscall,'" + MDIForm1.txtnama.text + "'"
                    strsql = strsql + " from mgm where id in (" + sstrSql2 + " )"
                    strsql = strsql + " and agent='" + LVRecycle.ListItems(i).SubItems(1) + "' limit " + CStr(Val(LVRecycle.ListItems(i).SubItems(4))) + ""
                    M_OBJCONN.Execute (strsql)
               
                    strsql = "update  mgm set agent=null,nama_agent=null,bucket ='R' ,statuscall='',tglcall =null, remarks =null,retur=NULL where id in (" + sstrSql2 + " "
                    strsql = strsql + " and agent='" + LVRecycle.ListItems(i).SubItems(1) + "' limit " + CStr(Val(LVRecycle.ListItems(i).SubItems(4))) + ")"
                    M_OBJCONN.Execute (strsql)
                                    
            End If
        
         End If
    Next i
    txtSudahDistribusi.text = "0"
    If LVRecycle.ListItems.Count = 0 Then
        MsgBox "Mohon masukan jumlah data yang akan di Recycle", vbInformation + vbOKOnly
        Exit Sub
    End If
    MsgBox "Data telah direcycle", vbInformation + vbOKOnly
End Sub
Public Function FungsiWaktuServer()
 'Fungsi Untuk mengambil waktu dan tanggal di server database
 Dim CMDSQL As String
 Dim M_objrs As ADODB.Recordset
 
 CMDSQL = "select now() as waktu"
 
 Set M_objrs = New ADODB.Recordset
 M_objrs.CursorLocation = adUseClient
 
 M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
 
 WaktuServer = Format(M_objrs(0), "yyyy-mm-dd hh:mm:00")
 FungsiWaktuServer = WaktuServer
 Set M_objrs = Nothing
End Function
Private Sub CmdSearchBaru_Click(Index As Integer)
Dim mobjrs As New ADODB.Recordset
Dim list As ListItem
Select Case Index
    Case 1
    Set mobjrs = New ADODB.Recordset
    mobjrs.CursorLocation = adUseClient
    strsql = " select tglentry,userinput ,count(tblrecyle_no_case) as jml from tblrecyle_hst group by tglentry,userinput "
    mobjrs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    no = 0
    ListView1(1).ListItems.CLEAR
    While Not mobjrs.EOF
        no = no + 1
        Set list = ListView1(1).ListItems.ADD(, , no)
        list.SubItems(1) = IIf(IsNull(mobjrs!userinput), "", mobjrs!userinput)
        list.SubItems(2) = Format(IIf(IsNull(mobjrs!tglentry), "", mobjrs!tglentry), "dd-mm-yyy hh:mm:nn")
        list.SubItems(3) = IIf(IsNull(mobjrs!jml), "0", mobjrs!jml)
        mobjrs.MoveNext
    Wend
End Select
End Sub

Private Sub Form_Load()
    header_distribusi_Recycle
End Sub
Public Sub load_campaign()

    If MDIForm1.txtlevel.text = "Agent" Then
        sStrsql = "select * from  datasourcetbl where "
        sStrsql = sStrsql + " kodeds in (select distinct recsource from mgm "
        sStrsql = sStrsql + " where agent ='" + MDIForm1.TxtUsername.text + "') and status ='1'"
    ElseIf MDIForm1.txtlevel.text = "Supervisor" Then
        sStrsql = "select * from  datasourcetbl where "
        sStrsql = sStrsql + " kodeds in (select distinct recsource from mgm "
        sStrsql = sStrsql + " where agent in "
        sStrsql = sStrsql + " (select userid  from usertbl where  spvcode='" + MDIForm1.TxtUsername.text + "' or userid = '" + MDIForm1.TxtUsername.text + "')) and   status ='1'"
    Else
        sStrsql = "select * from  datasourcetbl  where  status ='1'"
    End If
    
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        cmbcampaigncode.CLEAR
        While Not M_objrs.EOF
            cmbcampaigncode.AddItem IIf(IsNull(M_objrs!KODEDS), "", M_objrs!KODEDS)
            M_objrs.MoveNext
        Wend
        
    Set M_objrs = Nothing
    
End Sub
Public Sub load_spv()
    If MDIForm1.txtlevel.text = "Agent" Then
        sStrsql = " select userid , agent  from usertbl where  userid in  (select distinct spvcode  from usertbl where  spvcode= '" + MDIForm1.TxtUsername.text + "') and AKTIF ='1'"
    ElseIf MDIForm1.txtlevel.text = "Supervisor" Then
        sStrsql = " select userid , agent  from usertbl where  userid = '" + MDIForm1.TxtUsername.text + "' and AKTIF ='1'"
        Else
        sStrsql = "select userid , agent  from usertbl  where  AKTIF ='1' and  level_name ='Supervisor'"
    End If
    
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        cbosupervisor.CLEAR
        While Not M_objrs.EOF
                cbosupervisor.AddItem IIf(IsNull(M_objrs!USERID), "", M_objrs!USERID) & "-" & IIf(IsNull(M_objrs!AGENT), "", M_objrs!AGENT)
                M_objrs.MoveNext
        Wend
        
    Set M_objrs = Nothing
End Sub
Public Sub load_telesales()
Dim getUserid As String
Dim getUser_name As String
    If MDIForm1.txtlevel.text = "Agent" Then
        sStrsql = " select  userid , agent  from usertbl where  userid = '" + MDIForm1.TxtUsername.text + "' and AKTIF ='1'"
    ElseIf MDIForm1.txtlevel.text = "Supervisor" Then
        sStrsql = " select  userid , agent  from usertbl where spvcode = '" + MDIForm1.TxtUsername.text + "' and  AKTIF ='1'"
    Else
        intvrl = InStr(1, cbosupervisor, "-", vbTextCompare)
           If intvrl <> 0 Then
              ArrayString = Split(cbosupervisor.text, "-", 2, vbTextCompare)
              getUserid = ArrayString(0)
              getUser_name = ArrayString(1)
           End If
    mwhere = ""
    If cbosupervisor.text <> "" Then
         mwhere = " AND spvcode ='" + getUserid + "' "
    End If
        sStrsql = "select  userid , agent  from usertbl  where  AKTIF ='1' and  level_name ='Agent'" + mwhere
    End If
    
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
       cbotelesales.CLEAR
        While Not M_objrs.EOF
                cbotelesales.AddItem IIf(IsNull(M_objrs!USERID), "", M_objrs!USERID) & "-" & IIf(IsNull(M_objrs!AGENT), "", M_objrs!AGENT)
                M_objrs.MoveNext
        Wend
    Set M_objrs = Nothing
End Sub
Public Sub load_statuscall()
    sStrsql = " select tblstatuscall_kdstscall, tblstatuscall_keterangan  from tblstatuscall where tblstatuscall_keterangan<> 'PTP' "
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        CmbStatusCall.CLEAR
        
        While Not M_objrs.EOF
            CmbStatusCall.AddItem IIf(IsNull(M_objrs!tblstatuscall_keterangan), "", M_objrs!tblstatuscall_keterangan)
            M_objrs.MoveNext
        Wend
    Set M_objrs = Nothing
End Sub
Public Sub load_statuscall2()
    sStrsql = " select tblstatuscall_id,tblstatuscall_kdstscall,tblstatuscall_keterangan  from tblstatuscall where tblstatuscall_keterangan<> 'PTP'  order by tblstatuscall_keterangan asc"
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        cbodestination_statuscall.CLEAR
        kdstatuscall.CLEAR
        While Not M_objrs.EOF
            cbodestination_statuscall.AddItem IIf(IsNull(M_objrs!tblstatuscall_keterangan), "", M_objrs!tblstatuscall_keterangan)
            kdstatuscall.AddItem IIf(IsNull(M_objrs!tblstatuscall_id), "", M_objrs!tblstatuscall_id)
            M_objrs.MoveNext
        Wend
    Set M_objrs = Nothing
End Sub

Private Sub LVRecycle_DblClick()
If LVRecycle.ListItems.Count = 0 Then
       MsgBox "Data agent tidak ada!", vbOKOnly + vbInformation, "Informasi"
       Exit Sub
End If
    
    setJmlDistribusi = InputBox("Inputkan jumlah data recycle untuk:" & LVRecycle.SelectedItem.SubItems(1) & "-" & LVRecycle.SelectedItem.SubItems(2), "Recycle Data")
    If setJmlDistribusi = "" Then setJmlDistribusi = 0
        If Val(LVRecycle.SelectedItem.SubItems(3)) < setJmlDistribusi Then
            m_msgbox = MsgBox("Data melebihi jumlah data yang dimiliki agent:" & LVRecycle.SelectedItem.SubItems(2), vbOKOnly + vbInformation, "Informasi")
            Exit Sub
        End If
    
    LVRecycle.SelectedItem.SubItems(4) = setJmlDistribusi
    jmlDtSudahDistribusi = 0
    For i = 1 To LVRecycle.ListItems.Count
        jmlDtSudahDistribusi = jmlDtSudahDistribusi + Val(LVRecycle.ListItems.Item(i).SubItems(4))
    Next i
    
    txtSudahDistribusi.text = jmlDtSudahDistribusi
    txtSisaCampaign.text = Val(txtjmlcampaign.text) - jmlDtSudahDistribusi


    
End Sub

