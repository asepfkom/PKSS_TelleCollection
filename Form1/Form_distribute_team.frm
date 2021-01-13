VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form_distributeteam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Distribute Data"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   11325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdkeluar 
      BackColor       =   &H00F1E5DB&
      Caption         =   "&Keluar"
      Height          =   495
      Left            =   9645
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   3630
      Width           =   1575
   End
   Begin VB.CommandButton cmdProses 
      BackColor       =   &H00F1E5DB&
      Caption         =   "&PROSES"
      Enabled         =   0   'False
      Height          =   495
      Left            =   9645
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   3135
      Width           =   1575
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10365
      Left            =   -120
      TabIndex        =   1
      Top             =   810
      Width           =   18480
      _ExtentX        =   32597
      _ExtentY        =   18283
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Criteria Distribute"
      TabPicture(0)   =   "Form_distribute_team.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image3(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label8"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label11"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label20"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label10"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label26"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label27"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label9"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "ProgressBar1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmbcampaigncode"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cbolimit"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cbooperand"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cbolimit1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Frame2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdloadcampaign"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cbostatus"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Command1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Frame4"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Frame5"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Frame6"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cbopendate1"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cbopendate2"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtJumlah(0)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cbofieldfilter"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Check2"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cboarea"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Check3"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Check4"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Frame3"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Frame9"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Command4"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "cmbregion"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).ControlCount=   35
      TabCaption(1)   =   "History Distribute"
      TabPicture(1)   =   "Form_distribute_team.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Image3(2)"
      Tab(1).Control(1)=   "Label19"
      Tab(1).Control(2)=   "Label21"
      Tab(1).Control(3)=   "Label25"
      Tab(1).Control(4)=   "Label28"
      Tab(1).Control(5)=   "ListView6"
      Tab(1).Control(6)=   "ListView5"
      Tab(1).Control(7)=   "ListView4"
      Tab(1).Control(8)=   "CMBHISTORY"
      Tab(1).Control(9)=   "Command3"
      Tab(1).Control(10)=   "Combo2"
      Tab(1).Control(11)=   "Text5"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Bucket Monitoring"
      TabPicture(2)   =   "Form_distribute_team.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame7"
      Tab(2).Control(1)=   "Frame8"
      Tab(2).ControlCount=   2
      Begin VB.ComboBox cmbregion 
         Height          =   315
         ItemData        =   "Form_distribute_team.frx":0054
         Left            =   1470
         List            =   "Form_distribute_team.frx":0056
         TabIndex        =   92
         Top             =   705
         Width           =   2265
      End
      Begin VB.CommandButton Command4 
         Caption         =   "TARIK NEW DATA"
         Height          =   375
         Left            =   14760
         TabIndex        =   91
         Top             =   840
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Agent"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5715
         Left            =   135
         TabIndex        =   84
         Top             =   2250
         Width           =   9345
         Begin VB.CheckBox CheckAll_Agent 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check All"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   5295
            Width           =   1110
         End
         Begin VB.TextBox txtJmlAgent 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8205
            Locked          =   -1  'True
            TabIndex        =   85
            Text            =   "0"
            Top             =   5295
            Width           =   975
         End
         Begin MSComctlLib.ListView LVAgent 
            Height          =   4920
            Left            =   90
            TabIndex        =   86
            Top             =   300
            Width           =   9105
            _ExtentX        =   16060
            _ExtentY        =   8678
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   12582912
            BackColor       =   16777215
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
         Begin MSComctlLib.ProgressBar ProgressBar3 
            Height          =   255
            Left            =   1380
            TabIndex        =   87
            Top             =   5295
            Width           =   4545
            _ExtentX        =   8017
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
         End
         Begin VB.Label Label30 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Jumlah User :"
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
            Left            =   6165
            TabIndex        =   88
            Top             =   5295
            Width           =   2055
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Admin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Left            =   -9210
         TabIndex        =   76
         Top             =   2190
         Visible         =   0   'False
         Width           =   9375
         Begin VB.CheckBox CheckAll_MGR 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check All"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   89
            Top             =   2040
            Width           =   1455
         End
         Begin VB.CommandButton Cmd_Refersh 
            BackColor       =   &H00F1E5DB&
            Caption         =   "REFRESH"
            Height          =   255
            Left            =   3960
            Style           =   1  'Graphical
            TabIndex        =   83
            Top             =   2040
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtJmlManager 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8280
            Locked          =   -1  'True
            TabIndex        =   77
            Text            =   "0"
            Top             =   2040
            Width           =   975
         End
         Begin MSComctlLib.ListView LVMgr 
            Height          =   1635
            Left            =   90
            TabIndex        =   78
            Top             =   300
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   2884
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   12582912
            BackColor       =   16777215
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
         Begin MSComctlLib.ProgressBar ProgressBar2 
            Height          =   255
            Left            =   2760
            TabIndex        =   79
            Top             =   4440
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
         End
         Begin VB.Label Label29 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Jumlah User :"
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
            Left            =   6240
            TabIndex        =   80
            Top             =   2040
            Width           =   2055
         End
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Not Active"
         Height          =   375
         Left            =   2640
         TabIndex        =   75
         Top             =   975
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Active"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         TabIndex        =   74
         Top             =   975
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -67560
         TabIndex        =   72
         Top             =   9405
         Width           =   855
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   -73545
         TabIndex        =   69
         Top             =   480
         Width           =   5040
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00F1E5DB&
         Caption         =   "&Load Data"
         Height          =   375
         Left            =   -68280
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   480
         Width           =   1440
      End
      Begin VB.ComboBox cboarea 
         Height          =   315
         ItemData        =   "Form_distribute_team.frx":0058
         Left            =   19410
         List            =   "Form_distribute_team.frx":0065
         TabIndex        =   66
         Top             =   1140
         Visible         =   0   'False
         Width           =   4185
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Show Jml Assign"
         Height          =   345
         Left            =   4440
         TabIndex        =   65
         Top             =   975
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Frame Frame8 
         Caption         =   "Data"
         Height          =   6435
         Left            =   -74790
         TabIndex        =   58
         Top             =   1290
         Width           =   9825
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8730
            Locked          =   -1  'True
            TabIndex        =   60
            Top             =   6030
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8730
            Locked          =   -1  'True
            TabIndex        =   59
            Top             =   2280
            Width           =   855
         End
         Begin MSComctlLib.ListView ListView7 
            Height          =   2055
            Left            =   60
            TabIndex        =   61
            Top             =   180
            Width           =   9585
            _ExtentX        =   16907
            _ExtentY        =   3625
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
         Begin MSComctlLib.ListView ListView8 
            Height          =   3435
            Left            =   90
            TabIndex        =   62
            Top             =   2580
            Width           =   9525
            _ExtentX        =   16801
            _ExtentY        =   6059
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
         Begin VB.Label Label24 
            Caption         =   "Total all agent Bucket:"
            Height          =   315
            Left            =   7080
            TabIndex        =   64
            Top             =   6060
            Width           =   2055
         End
         Begin VB.Label Label23 
            Caption         =   "Total all TL Bucket:"
            Height          =   315
            Left            =   7290
            TabIndex        =   63
            Top             =   2310
            Width           =   1665
         End
      End
      Begin VB.Frame Frame7 
         Height          =   855
         Left            =   -74790
         TabIndex        =   52
         Top             =   390
         Width           =   9795
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1500
            TabIndex        =   56
            Top             =   180
            Width           =   2445
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00F1E5DB&
            Caption         =   "&Load Data"
            Height          =   375
            Left            =   4080
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   150
            Width           =   1755
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Similiar Search (use % char)"
            Height          =   225
            Left            =   120
            TabIndex        =   54
            Top             =   600
            Visible         =   0   'False
            Width           =   2265
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            Height          =   315
            Left            =   7140
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   210
            Width           =   2475
         End
         Begin VB.Label Label22 
            BackColor       =   &H00F1E5DB&
            BackStyle       =   0  'Transparent
            Caption         =   "Campaign Code :"
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
            Left            =   90
            TabIndex        =   57
            Top             =   210
            Width           =   1455
         End
      End
      Begin VB.ComboBox cbofieldfilter 
         Height          =   315
         Left            =   20820
         TabIndex        =   50
         Top             =   390
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtJumlah 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   21720
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   1470
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.ComboBox cbopendate2 
         Height          =   315
         Left            =   19620
         TabIndex        =   46
         Top             =   1500
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.ComboBox cbopendate1 
         Height          =   315
         Left            =   19800
         TabIndex        =   45
         Top             =   1500
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.ComboBox CMBHISTORY 
         Height          =   315
         Left            =   -62280
         TabIndex        =   40
         Top             =   495
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Data yang sudah di assign ke agent "
         Height          =   3705
         Left            =   20160
         TabIndex        =   35
         Top             =   3840
         Visible         =   0   'False
         Width           =   9180
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   1395
            TabIndex        =   36
            Top             =   3180
            Width           =   1125
         End
         Begin MSComctlLib.ListView ListView3 
            Height          =   2865
            Left            =   90
            TabIndex        =   37
            Top             =   210
            Width           =   9000
            _ExtentX        =   15875
            _ExtentY        =   5054
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
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Already Send"
            Height          =   345
            Left            =   195
            TabIndex        =   38
            Top             =   3240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Status Data Di TL"
         Height          =   2265
         Left            =   19560
         TabIndex        =   28
         Top             =   2880
         Visible         =   0   'False
         Width           =   9180
         Begin VB.TextBox txtalreadytl 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   4410
            TabIndex        =   30
            Top             =   1830
            Width           =   1125
         End
         Begin VB.TextBox txtavalabletl 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2010
            TabIndex        =   29
            Top             =   1860
            Width           =   1125
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   1515
            Left            =   90
            TabIndex        =   31
            Top             =   210
            Width           =   9000
            _ExtentX        =   15875
            _ExtentY        =   2672
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
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Already Send"
            Height          =   345
            Left            =   3210
            TabIndex        =   34
            Top             =   1890
            Width           =   1095
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Available"
            Height          =   345
            Left            =   1320
            TabIndex        =   33
            Top             =   1890
            Width           =   1095
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Lead :"
            Height          =   255
            Left            =   150
            TabIndex        =   32
            Top             =   1890
            Width           =   1245
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Lead By Campaign"
         Height          =   2265
         Left            =   19200
         TabIndex        =   21
         Top             =   1440
         Visible         =   0   'False
         Width           =   9180
         Begin VB.TextBox txtavailable 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2010
            TabIndex        =   23
            Top             =   1860
            Width           =   1125
         End
         Begin VB.TextBox txtalreadyassign 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   4410
            TabIndex        =   22
            Top             =   1830
            Width           =   1125
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1605
            Left            =   90
            TabIndex        =   24
            Top             =   210
            Width           =   8955
            _ExtentX        =   15796
            _ExtentY        =   2831
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
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Lead :"
            Height          =   255
            Left            =   150
            TabIndex        =   27
            Top             =   1890
            Width           =   1245
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Available"
            Height          =   345
            Left            =   1215
            TabIndex        =   26
            Top             =   1890
            Width           =   1095
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Already Send"
            Height          =   345
            Left            =   3210
            TabIndex        =   25
            Top             =   1890
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00F1E5DB&
         Caption         =   "&View All"
         Height          =   375
         Left            =   14040
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2400
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.ComboBox cbostatus 
         Height          =   315
         ItemData        =   "Form_distribute_team.frx":0082
         Left            =   1470
         List            =   "Form_distribute_team.frx":008C
         TabIndex        =   17
         Top             =   375
         Width           =   2265
      End
      Begin VB.CommandButton cmdloadcampaign 
         BackColor       =   &H00F1E5DB&
         Caption         =   "&Load Data"
         Height          =   705
         Left            =   9630
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   405
         Width           =   1485
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Informasi data Periode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   120
         TabIndex        =   9
         Top             =   1380
         Width           =   11355
         Begin VB.TextBox txtjmlcampaign 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   "0"
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtSisaCampaign 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   8325
            Locked          =   -1  'True
            TabIndex        =   11
            Text            =   "0"
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtSudahDistribusi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H0000FF00&
            Height          =   285
            Left            =   5190
            Locked          =   -1  'True
            TabIndex        =   10
            Text            =   "0"
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total Lead :"
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
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Sisa Leads :"
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
            Left            =   6645
            TabIndex        =   14
            Top             =   360
            Width           =   1665
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Sudah didistribusi :"
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
            Left            =   3510
            TabIndex        =   13
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.ComboBox cbolimit1 
         Height          =   315
         Left            =   23790
         MousePointer    =   1  'Arrow
         TabIndex        =   8
         Top             =   750
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox cbooperand 
         Height          =   315
         ItemData        =   "Form_distribute_team.frx":009E
         Left            =   22590
         List            =   "Form_distribute_team.frx":00B4
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   1950
      End
      Begin VB.ComboBox cbolimit 
         Height          =   315
         Left            =   16680
         TabIndex        =   5
         Top             =   1920
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.ComboBox cmbcampaigncode 
         Height          =   315
         Left            =   1470
         TabIndex        =   3
         Top             =   45
         Width           =   7665
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   18600
         TabIndex        =   20
         Top             =   600
         Visible         =   0   'False
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   8370
         Left            =   -74940
         TabIndex        =   39
         Top             =   870
         Width           =   13290
         _ExtentX        =   23442
         _ExtentY        =   14764
         View            =   3
         SortOrder       =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
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
      Begin MSComctlLib.ListView ListView5 
         Height          =   4020
         Left            =   -63795
         TabIndex        =   42
         Top             =   1035
         Visible         =   0   'False
         Width           =   6930
         _ExtentX        =   12224
         _ExtentY        =   7091
         View            =   3
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
      Begin MSComctlLib.ListView ListView6 
         Height          =   3375
         Left            =   -63840
         TabIndex        =   43
         Top             =   5580
         Visible         =   0   'False
         Width           =   7020
         _ExtentX        =   12383
         _ExtentY        =   5953
         View            =   3
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
      Begin VB.Label Label9 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   885
         TabIndex        =   93
         Top             =   690
         Width           =   645
      End
      Begin VB.Label Label27 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Status User"
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
         Left            =   435
         TabIndex        =   73
         Top             =   1050
         Width           =   1050
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   255
         Left            =   -68160
         TabIndex        =   71
         Top             =   9405
         Width           =   615
      End
      Begin VB.Label Label25 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Campaign Code"
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
         Left            =   -74880
         TabIndex        =   70
         Top             =   525
         Width           =   1455
      End
      Begin VB.Label Label26 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Area"
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
         Left            =   18600
         TabIndex        =   67
         Top             =   1200
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label10 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Field Filter "
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
         Left            =   16680
         TabIndex        =   51
         Top             =   1320
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label2 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Trans :"
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
         Left            =   15120
         TabIndex        =   48
         Top             =   1530
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Between"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   18480
         TabIndex        =   47
         Top             =   1530
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL PER AGENT"
         Height          =   255
         Left            =   -61710
         TabIndex        =   44
         Top             =   5220
         Width           =   3375
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Pilih Tanggal  "
         Height          =   255
         Left            =   -63840
         TabIndex        =   41
         Top             =   540
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label11 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         Left            =   885
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label8 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Operator "
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
         Left            =   15120
         TabIndex        =   6
         Top             =   1440
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label5 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Limit Amount "
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
         Left            =   24030
         TabIndex        =   4
         Top             =   960
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Campaign Code"
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
         Left            =   135
         TabIndex        =   2
         Top             =   90
         Width           =   1455
      End
      Begin VB.Image Image3 
         Height          =   18630
         Index           =   0
         Left            =   -7680
         Picture         =   "Form_distribute_team.frx":00D1
         Top             =   -240
         Width           =   26295
      End
      Begin VB.Image Image3 
         Height          =   18630
         Index           =   2
         Left            =   -74955
         Picture         =   "Form_distribute_team.frx":76DB
         Top             =   315
         Visible         =   0   'False
         Width           =   26295
      End
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   1
      Left            =   150
      Picture         =   "Form_distribute_team.frx":ECE5
      Stretch         =   -1  'True
      Top             =   240
      Width           =   420
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Distribute Data"
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
      Index           =   1
      Left            =   645
      TabIndex        =   0
      Top             =   300
      Width           =   3585
   End
   Begin VB.Image Image2 
      Height          =   825
      Left            =   -2280
      Picture         =   "Form_distribute_team.frx":F7EF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20700
   End
End
Attribute VB_Name = "Form_distributeteam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ttlsudah_distribute  As Double
Dim sQerySelect As String
Dim sGetSPV As String
Dim sUserid, sLevel As String
Public Function GETSPV(ByVal KetLevel As String) As Variant
    Dim row As Double
    row = 1
    strsql = ""
    If KetLevel = "1" Then 'AGENT
        row = 1
        For i = 1 To LvAgent.ListItems.Count
           If LvAgent.ListItems(i).Checked = True Then
                If row = 1 Then
                    strsql = "'" + LvAgent.ListItems(i).SubItems(1) + "'"
                Else
                    strsql = strsql + ",'" + LvAgent.ListItems(i).SubItems(1) + "'"
                End If
                row = row + 1
          End If
        Next i
    End If
    GETSPV = strsql
End Function

Private Sub cbooperand_Click()
If cbooperand.text = "between" Then
    cbolimit1.Visible = True
Else
    cbolimit1.Visible = False
End If

End Sub

Private Sub cbooperand_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub

Private Sub cbopocket_DropDown()
Select Case MDIForm1.txtlevel
    Case "Supervisor"
      
       cbopocket.CLEAR
       cbopocket.AddItem sUserid & "!" & MDIForm1.txtnama
    Case "Manager"
    sStrsql = "select * from  usertbl where kdlevel<>'1'"
       Set M_objrs = New ADODB.Recordset
           M_objrs.CursorLocation = adUseClient
           M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
       cbopocket.CLEAR
       cbopocket.AddItem "New Bucket!"
     cbopocket.AddItem sUserid & "!" & MDIForm1.txtnama
       While Not M_objrs.EOF
            cbopocket.AddItem IIf(IsNull(M_objrs!USERID), "", M_objrs!USERID) & "!" & IIf(IsNull(M_objrs!AGENT), "", M_objrs!AGENT)
            M_objrs.MoveNext
       Wend
       
    Case "Administrator"
       sStrsql = "select * from  usertbl where kdlevel<>'1'"
       Set M_objrs = New ADODB.Recordset
           M_objrs.CursorLocation = adUseClient
           M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
       cbopocket.CLEAR
       cbopocket.AddItem "New Bucket!"
       cbopocket.AddItem sUserid & "!" & MDIForm1.txtnama
       
       
       While Not M_objrs.EOF
            cbopocket.AddItem IIf(IsNull(M_objrs!USERID), "", M_objrs!USERID) & "!" & IIf(IsNull(M_objrs!AGENT), "", M_objrs!AGENT)
            M_objrs.MoveNext
       Wend
       
End Select

End Sub

Private Sub cbopocket_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Check3_Click()
If Check3.Value = vbChecked Then
Check4.Value = vbUnchecked
End If

End Sub

Private Sub Check4_Click()
If Check4.Value = vbChecked Then
Check3.Value = vbUnchecked
End If
End Sub

Private Sub CheckAll_Agent_Click()
    If CheckAll_Agent.Value = 1 Then
        If LvAgent.ListItems.Count <> 0 Then
            For i = 1 To LvAgent.ListItems.Count
                LvAgent.ListItems(i).Checked = True
            Next i
        End If
    ElseIf CheckAll_Agent.Value = 0 Then
        If LvAgent.ListItems.Count <> 0 Then
            For i = 1 To LvAgent.ListItems.Count
                LvAgent.ListItems(i).Checked = False
            Next i
        End If
    End If
End Sub

Private Sub cmbcampaigncode_Click()
    Dim M_objrs As ADODB.Recordset
    Dim M_OBJRS_history As ADODB.Recordset
    Dim M_OBJRS_history3 As ADODB.Recordset
    Dim M_OBJRS_tglhistory As ADODB.Recordset
    Dim CMDSQL As String
    Dim cmdsql_history As String
    Dim cmdsql_history3 As String
    Dim cmdsql_tglhistory As String
    Dim ListItem As ListItem
    Dim getCampaign_code As String
    Dim getCampaign_name As String

    
     intvrl = InStr(1, cmbcampaigncode.text, "!", vbTextCompare)
               If intvrl <> 0 Then
                  ArrayString = Split(cmbcampaigncode.text, "!", 2, vbTextCompare)
                  getCampaign_code = ArrayString(0)
                  getCampaign_name = ArrayString(1)
               End If
               getCampaign_code = cmbcampaigncode.text 'HENDRI CODE
    ListView4.ListItems.CLEAR
    ListView5.ListItems.CLEAR
    ListView6.ListItems.CLEAR

End Sub
Private Sub cmbcampaigncode_DropDown()
'sstrsql = "select * from tbldatasource where   tbldatasource_kdstatus ='1' order by   tbldatasource_tglentry,  tbldatasource_keterangan "
'Set M_objrs = New ADODB.Recordset
'    M_objrs.CursorLocation = adUseClient
'    M_objrs.Open sstrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'    cmbcampaigncode.Clear
'    While Not M_objrs.EOF
'        'cmbcampaigncode.AddItem IIf(IsNull(M_OBJRS!tbldatasource_campaign_code), "", M_OBJRS!tbldatasource_campaign_code) & "!" & IIf(IsNull(M_OBJRS!tbldatasource_keterangan), "", M_OBJRS!tbldatasource_keterangan)
'        cmbcampaigncode.AddItem IIf(IsNull(M_objrs!tbldatasource_campaign_code), "", M_objrs!tbldatasource_campaign_code)
'        M_objrs.MoveNext
'    Wend
'Set M_objrs = Nothing
End Sub
Private Sub cmbcampaigncode_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub CMBHISTORY_Click()
    Dim M_OBJRS_history2 As ADODB.Recordset
    Dim cmdsql_history2 As String
    Dim ListItem As ListItem
    
    ListView5.ListItems.CLEAR
       intvrl = InStr(1, cmbcampaigncode.text, "!", vbTextCompare)
               If intvrl <> 0 Then
                  ArrayString = Split(cmbcampaigncode.text, "!", 2, vbTextCompare)
                  getCampaign_code = ArrayString(0)
                  getCampaign_name = ArrayString(1)
               End If
               getCampaign_code = cmbcampaigncode.text 'HENDRI CODE
    
        cmdsql_history2 = "select USERID, SUM(JMLDATA) AS TOTAL from tbllogdistribusi where campaign_code='" + getCampaign_code + "' and sendby = '" + sUserid + "' AND DATE(TGL) = '" + CMBHISTORY.text + "' GROUP BY USERID ORDER BY USERID"
        Set M_OBJRS_history2 = New ADODB.Recordset
        M_OBJRS_history2.CursorLocation = adUseClient
        M_OBJRS_history2.Open cmdsql_history2, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        
    While Not M_OBJRS_history2.EOF
        
        Set ListItem = ListView5.ListItems.ADD(, , IIf(IsNull(M_OBJRS_history2("USERID")), "", M_OBJRS_history2("USERID")))
            ListItem.SubItems(1) = IIf(IsNull(M_OBJRS_history2("TOTAL")), "", M_OBJRS_history2("TOTAL"))
            M_OBJRS_history2.MoveNext
    Wend
    Set M_OBJRS_history2 = Nothing

End Sub

Private Sub Cmd_Refersh_Click()
    KeluarListView ("3")
End Sub
Private Function KeluarListView(ByVal KetLevel As String)
    'MALIK
    
    Dim rsTBLUSER As ADODB.Recordset
    Dim mwhere   As String
    Dim getUserid  As String
    Dim getUsername As String
    Dim getCampaign_code As String
    Dim getCampaign_name   As String
    Dim no As Integer

    If KetLevel = "1" Then
        
            
                   
        intvrl = InStr(1, cmbcampaigncode.text, "!", vbTextCompare)
        If intvrl <> 0 Then
            ArrayString = Split(cmbcampaigncode.text, "!", 2, vbTextCompare)
            getCampaign_code = ArrayString(0)
            getCampaign_name = ArrayString(1)
        End If
        'sGetSPV = GETSPV("2")
        getCampaign_code = cmbcampaigncode.text
  
        Select Case MDIForm1.txtlevel
            Case "Supervisor"
                mwhere = ""
                If cbofieldfilter.text <> "" Then
                    If cbolimit.text <> "" And cbooperand.text = "" Then
                        mwhere = mwhere + " and  " + cbofieldfilter + " = " + CStr(cbolimit) + ""
                    Else
                        If cbolimit.text <> "" Then
                            If cbooperand.text <> "" Then
                                If cbooperand.text = "between" Then
                                    mwhere = mwhere + " and  " + cbofieldfilter + " between " + CStr(cbolimit) + " and " + CStr(cbolimit1.text) + ""
                                Else
                                    mwhere = mwhere + " and " + cbofieldfilter + " " + cbooperand.text + " " + CStr(cbolimit)
                                End If
                            End If
                        End If
                    End If
                End If
                If cbostatus.text <> "" Then
                    If cbostatus.text = "Recycle" Then
                        mwhere = mwhere + " and  bucket ='R'"
                    End If
                    If cbostatus.text = "New" Then
                        mwhere = mwhere + " and  (bucket is null or bucket='N')"
                    End If
                End If
                
                If cmbcampaigncode.text <> "" Then
                    mwhere = mwhere + " and  recsource ='" + getCampaign_code + "'"
                End If

                
                If Check2.Value = vbUnchecked Then
                    If Check3.Value = vbChecked Then
                        sStrsql = " select a.userid,a.agent, a.level_name from (SELECT * FROM usertbl WHERE (KDLEVEL='2') AND spvcode='" + MDIForm1.txtusername.text + "' " & mwHere2 & " ) a " ' and date(tgl_login) =date(now()) ) a "
                    ElseIf Check4.Value = vbChecked Then
                        sStrsql = " select a.userid,a.agent, a.level_name from (SELECT * FROM usertbl WHERE (KDLEVEL='2') AND spvcode='" + MDIForm1.txtusername.text + "'" & mwHere2 & " ) a " ' and (date(tgl_login) <date(now()) or tgl_login is null)) a "
                    Else
                        sStrsql = " select a.userid,a.agent, a.level_name from (SELECT * FROM usertbl WHERE (KDLEVEL='2') AND spvcode='" + MDIForm1.txtusername.text + "'" & mwHere2 & ") a "
                    End If
                Else
                    If Check3.Value = vbChecked Then
                        sStrsql = " select a.userid,a.agent,b.jml,a.level_name from (SELECT * FROM usertbl WHERE (KDLEVEL='1') AND aktif='1' AND spvcode='" + MDIForm1.txtusername.text + "' " & mwHere2 & " order by userid) a " 'and date(tgl_login) =date(now())
                        sStrsql = sStrsql + " Left Join "
                        sStrsql = sStrsql + "  ( "
                        sStrsql = sStrsql + " SELECT mgm.AGENT,COUNT(mgm.id)  AS JML FROM MGM,usertbl  "
                        sStrsql = sStrsql + "  where MGM.AGENT=userid " + mwhere + "  GROUP BY mgm.AGENT) b on a.userid=b.agent"
                    ElseIf Check4.Value = vbChecked Then
                        sStrsql = " select a.userid,a.agent,b.jml,a.level_name from (SELECT * FROM usertbl WHERE (KDLEVEL='1')  AND spvcode='" + MDIForm1.txtusername.text + "' " & mwHere2 & "  ) a " 'and  (date(tgl_login) <date(now()) or tgl_login is null)
                        sStrsql = sStrsql + " Left Join "
                        sStrsql = sStrsql + "  ( "
                        sStrsql = sStrsql + " SELECT mgm.AGENT,COUNT(mgm.id)  AS JML FROM MGM,usertbl  "
                        sStrsql = sStrsql + "  where MGM.AGENT=userid " + mwhere + "  GROUP BY mgm.AGENT) b on a.userid=b.agent"
                    Else
                        sStrsql = " select a.userid,a.agent,b.jml,a.level_name from (SELECT * FROM usertbl WHERE (KDLEVEL='1') AND aktif='1'  AND spvcode='" + MDIForm1.txtusername.text + "' " & mwHere2 & " order by userid) a "
                        sStrsql = sStrsql + " Left Join "
                        sStrsql = sStrsql + "  ( "
                        sStrsql = sStrsql + " SELECT mgm.AGENT,COUNT(mgm.id)  AS JML FROM MGM,usertbl  "
                        sStrsql = sStrsql + "  where MGM.AGENT=userid " + mwhere + "  GROUP BY mgm.AGENT) b on a.userid=b.agent"
                    End If
                End If
        End Select
        Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
        LvAgent.ListItems.CLEAR
        While Not M_objrs.EOF
            'Menginputkan data ke listview
            no = no + 1
            Set list = LvAgent.ListItems.ADD(, , no)
            list.SubItems(1) = IIf(IsNull(M_objrs!USERID), "", M_objrs!USERID)
            list.SubItems(2) = IIf(IsNull(M_objrs!AGENT), "", M_objrs!AGENT)
            list.SubItems(3) = IIf(IsNull(M_objrs!level_name), "", M_objrs!level_name)
            If Check2.Value = vbUnchecked Then
                list.SubItems(4) = 0
            Else
                list.SubItems(4) = IIf(IsNull(M_objrs!jml), 0, M_objrs!jml)
            End If
            M_objrs.MoveNext
        Wend
        Warna_Row_Listview Form_distributeteam, LvAgent, &HFFFFC0, vbWhite
        txtJmlAgent.text = M_objrs.RecordCount '-> isi jumlah spv ke txtjmlagent dan txtsisacampaign
        
        cmdProses.Enabled = True
        Set M_objrs = Nothing
    End If
    
    
End Function

Private Sub cmdkeluar_Click()
Unload Me
End Sub

Private Sub cmdloadcampaign_Click()
    Dim rs As New ADODB.Recordset
    Dim mobjrec As New ADODB.Recordset
    Dim total_lead As Double
    Dim mwhere As String
    LvAgent.ListItems.CLEAR
    If cmbcampaigncode.text = "" Then
         MsgBox "Campaign code harus diisi", vbInformation + vbOKOnly, "Pesan"
         Exit Sub
    End If
    
    If cbostatus.text <> "" Then
         If cbostatus.text = "Recycle" Then
            mwhere = mwhere + " and  bucket ='R'"
         End If
         
         If cbostatus.text = "New" Then
            mwhere = mwhere + " and  (bucket is null or bucket='N')"
        End If
    End If
    
    If cmbregion.text <> "" Then
            mwhere = mwhere + " and  region ='" + cmbregion.text + "'"
    End If
    
       '------------------------------------------------------------------------------------------------------------------------------------------
        strsql = "SELECT id FROM mgm where recsource ='" + cmbcampaigncode.text + "' and agent ='" & sUserid & "' " + mwhere + ""
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        ttlbelum_distribute = rs.RecordCount
        rs.Close
        Set rs = Nothing
       '------------------------------------------------------------------------------------------------------------------------------------------
        strsql = "select id  from mgm where recsource ='" + cmbcampaigncode.text + "' and agent in (" & _
                 "select userid from usertbl Where(spvcode ='" & sUserid & "')) " + mwhere + ""
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        txtSudahDistribusi.text = rs.RecordCount
        ttlsudah_distribute = rs.RecordCount
        rs.Close
        Set rs = Nothing
    total_lead = Val(ttlbelum_distribute) + Val(ttlsudah_distribute)
    txtjmlcampaign.text = total_lead
        KeluarListView ("1")
    txtSisaCampaign = Val(txtjmlcampaign.text) - Val(txtSudahDistribusi.text)
End Sub

Private Sub CmdProses_Click()
    Prosesdistribusi
End Sub

Private Sub Combo1_DropDown()
sStrsql = "select * from datasourcetbl where  status ='1' order by tglentry,  campaign_ket "
Set M_objrs = New ADODB.Recordset
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    Combo1.CLEAR
    While Not M_objrs.EOF
        Combo1.AddItem IIf(IsNull(M_objrs!KODEDS), "", M_objrs!KODEDS) & "!" & IIf(IsNull(M_objrs!campaign_ket), "", M_objrs!campaign_ket)
        M_objrs.MoveNext
    Wend
Set M_objrs = Nothing

End Sub
Private Sub header_distribusi_Spv()
    'elin
    
    LvAgent.ColumnHeaders.ADD 1, , "No", 5 * TXT
    LvAgent.ColumnHeaders.ADD 2, , "Kode ", 8 * TXT
    LvAgent.ColumnHeaders.ADD 3, , "Nama", 28 * TXT
    LvAgent.ColumnHeaders.ADD 4, , "Level", 7 * TXT
    LvAgent.ColumnHeaders.ADD 5, , "Jumlah Total", 7 * TXT
    LvAgent.ColumnHeaders.ADD 6, , "Jumlah awal", 7 * TXT
   
    
    ListView1.ColumnHeaders.ADD 1, , "NO", 5 * TXT
    ListView1.ColumnHeaders.ADD 2, , "BATCH", 15 * TXT
    ListView1.ColumnHeaders.ADD 3, , "ALL DATA", 10 * TXT
    ListView1.ColumnHeaders.ADD 4, , "AVAILABLE", 10 * TXT
    ListView1.ColumnHeaders.ADD 5, , "ALREADY ASSIGN", 10 * TXT
    
    
    
End Sub

Private Sub Combo2_DropDown()
sStrsql = "select * from tbldatasource where   tbldatasource_kdstatus ='1' order by   tbldatasource_tglentry,  tbldatasource_keterangan "
Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
Combo2.CLEAR
Combo2.AddItem "ALL"
    While Not M_objrs.EOF
        'cmbcampaigncode.AddItem IIf(IsNull(M_OBJRS!tbldatasource_campaign_code), "", M_OBJRS!tbldatasource_campaign_code) & "!" & IIf(IsNull(M_OBJRS!tbldatasource_keterangan), "", M_OBJRS!tbldatasource_keterangan)
        Combo2.AddItem IIf(IsNull(M_objrs!tbldatasource_campaign_code), "", M_objrs!tbldatasource_campaign_code)
        M_objrs.MoveNext
    Wend
Set M_objrs = Nothing


End Sub

Private Sub cmbregion_DropDown()
    Call load_region
End Sub
Public Sub load_region()
    sStrsql = " select distinct region  from mgm  "
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        cmbregion.CLEAR
        While Not M_objrs.EOF
                cmbregion.AddItem IIf(IsNull(M_objrs!region), "", M_objrs!region)
                M_objrs.MoveNext
        Wend
    Set M_objrs = Nothing
End Sub
Private Sub Command1_Click()
summeryCampaign
summerybyTL
summerybyAGENT

End Sub

Private Sub Command2_Click()
Command2.Enabled = False
    loadbucketTL '<--Load Bucket TeamLeader
Command2.Enabled = True
If Combo1.text <> Empty Then
    Text4.text = Combo1.text
End If

End Sub

Private Sub Command3_Click()
Dim M_OBJRS_history As New ADODB.Recordset
    If UCase(MDIForm1.txtlevel) = "ADMINISTRATOR" Or UCase(MDIForm1.txtlevel) = "MANAGER" Then
        If Combo2.text = "ALL" Then
            cmdsql_history = "select USERID,NAMA,CAMPAIGN_CODE,JMLDATA,SENDBY,TGL from tbllogdistribusi ORDER BY CAMPAIGN_CODE,TGL DESC"
        Else
            cmdsql_history = "select USERID,NAMA,CAMPAIGN_CODE,JMLDATA,SENDBY,TGL from tbllogdistribusi where campaign_code='" + Combo2.text + "' ORDER BY CAMPAIGN_CODE,TGL DESC"
        End If
    Else
        cmdsql_history = "select USERID,NAMA,CAMPAIGN_CODE,JMLDATA,SENDBY,TGL from tbllogdistribusi where campaign_code='" + Combo2.text + "' and sendby = '" + sUserid + "' ORDER BY TGL DESC"
    End If
     Set M_OBJRS_history = New ADODB.Recordset
     M_OBJRS_history.CursorLocation = adUseClient
     M_OBJRS_history.Open cmdsql_history, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
     ListView4.ListItems.CLEAR
     'Text5.Text = M_OBJRS_history.RecordCount
     
     While Not M_OBJRS_history.EOF
            Set ListItem = ListView4.ListItems.ADD(, , IIf(IsNull(M_OBJRS_history("CAMPAIGN_CODE")), "", M_OBJRS_history("CAMPAIGN_CODE")))
            ListItem.SubItems(1) = IIf(IsNull(M_OBJRS_history("USERID")), "", M_OBJRS_history("USERID"))
            ListItem.SubItems(2) = IIf(IsNull(M_OBJRS_history("NAMA")), "", M_OBJRS_history("NAMA"))
            ListItem.SubItems(3) = IIf(IsNull(M_OBJRS_history("JMLDATA")), "", M_OBJRS_history("JMLDATA"))
            ListItem.SubItems(4) = IIf(IsNull(M_OBJRS_history("SENDBY")), "", M_OBJRS_history("SENDBY"))
            ListItem.SubItems(5) = IIf(IsNull(M_OBJRS_history("TGL")), "", Format(M_OBJRS_history("TGL"), "dd-mmm-yyyy"))
            M_OBJRS_history.MoveNext
    Wend
    
End Sub

Private Sub Command4_Click()
    Form_recycleNewData.Show 1
End Sub

Private Sub Form_Load()
    If MDIForm1.txtlevel.text <> "Branch Manager" Then
        sUserid = MDIForm1.txtusername.text
        sLevel = MDIForm1.txtlevel.text
    Else
        sUserid = Form_DistributeSub.UserManagerSub
        sLevel = "Manager"
    End If
    header_distribusi_Spv
    header
    SSTab1.TabVisible(2) = False
    
'pemisahaan
If UCase(MDIForm1.txtlevel.text) = "SUPERVISOR" Then
    sStrsql = "select * from datasourcetbl where status ='1'  and (tglexpire > DATE(NOW()) OR tglexpire IS NULL  ) and kodeds in (select distinct recsource from mgm where agent in (select userid from usertbl where team = '" & MDIForm1.txtusername.text & "' or userid = '" & MDIForm1.txtusername.text & "' )) order by   kodeds"
ElseIf UCase(MDIForm1.txtlevel.text) = "ADMIN" Or UCase(MDIForm1.txtlevel.text) = "ADMINISTRATOR" Then
    sStrsql = "select * from datasourcetbl where status ='1'  and (tglexpire > DATE(NOW()) OR tglexpire IS NULL  ) order by   kodeds"
End If

Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    cmbcampaigncode.CLEAR
    While Not M_objrs.EOF
        'cmbcampaigncode.AddItem IIf(IsNull(M_OBJRS!tbldatasource_campaign_code), "", M_OBJRS!tbldatasource_campaign_code) & "!" & IIf(IsNull(M_OBJRS!tbldatasource_keterangan), "", M_OBJRS!tbldatasource_keterangan)
        cmbcampaigncode.AddItem IIf(IsNull(M_objrs!KODEDS), "", M_objrs!KODEDS)
        M_objrs.MoveNext
    Wend
Set M_objrs = Nothing

End Sub




Private Sub header()
    ListView4.ColumnHeaders.ADD 1, , "Campaign", 20 * TXT
    ListView4.ColumnHeaders.ADD 2, , "Agent", 20 * TXT
    ListView4.ColumnHeaders.ADD 3, , "Nama", 20 * TXT
    ListView4.ColumnHeaders.ADD 4, , "Jumlah Data", 12 * TXT
    ListView4.ColumnHeaders.ADD 5, , "Send By", 20 * TXT
    ListView4.ColumnHeaders.ADD 6, , "Tanggal", 20 * TXT
    
    ListView5.ColumnHeaders.ADD 1, , "AGENT", 20 * TXT
    ListView5.ColumnHeaders.ADD 2, , "TOTAL", 20 * TXT
    
    ListView6.ColumnHeaders.ADD 1, , "AGENT", 20 * TXT
    ListView6.ColumnHeaders.ADD 2, , "TOTAL", 20 * TXT
    
    
    
    ListView1.ColumnHeaders.ADD 1, , "NO", 5 * TXT
    ListView1.ColumnHeaders.ADD 2, , "BATCH", 15 * TXT
    ListView1.ColumnHeaders.ADD 3, , "ALL DATA", 10 * TXT
    ListView1.ColumnHeaders.ADD 4, , "AVAILABLE", 10 * TXT
    ListView1.ColumnHeaders.ADD 5, , "ALREADY ASSIGN", 10 * TXT
    
    ListView2.ColumnHeaders.ADD 1, , "NO", 5 * TXT
    ListView2.ColumnHeaders.ADD 2, , "Userid", 15 * TXT
    ListView2.ColumnHeaders.ADD 3, , "BATCH", 15 * TXT
    ListView2.ColumnHeaders.ADD 4, , "ALL DATA", 10 * TXT
    ListView2.ColumnHeaders.ADD 5, , "AVAILABLE", 10 * TXT
    ListView2.ColumnHeaders.ADD 6, , "ALREADY ASSIGN", 10 * TXT
    
    ListView3.ColumnHeaders.ADD 1, , "NO", 5 * TXT
    ListView3.ColumnHeaders.ADD 2, , "Userid", 15 * TXT
    ListView3.ColumnHeaders.ADD 3, , "BATCH", 15 * TXT
    ListView3.ColumnHeaders.ADD 4, , "ALREADY ASSIGN", 10 * TXT
    
    ListView7.ColumnHeaders.ADD 1, , "NO", 5 * TXT
    ListView7.ColumnHeaders.ADD 2, , "TL", 10 * TXT
    ListView7.ColumnHeaders.ADD 3, , "TeamLeader", 20 * TXT
    ListView7.ColumnHeaders.ADD 4, , "Total", 10 * TXT
    
    ListView8.ColumnHeaders.ADD 1, , "NO", 5 * TXT
    ListView8.ColumnHeaders.ADD 2, , "Agent", 10 * TXT
    ListView8.ColumnHeaders.ADD 3, , "Userid", 20 * TXT
    ListView8.ColumnHeaders.ADD 4, , "Total", 10 * TXT
    
End Sub
Public Sub isijumlcahcampaign()
   Dim mwhere   As String
   Dim getUserid  As String
   Dim getUsername As String
   Dim getCampaign_code As String
   Dim getCampaign_name   As String
   Dim m_objrs2  As New ADODB.Recordset
               
               intvrl = InStr(1, cmbcampaigncode.text, "!", vbTextCompare)
               If intvrl <> 0 Then
                  ArrayString = Split(cmbcampaigncode.text, "!", 2, vbTextCompare)
                  getCampaign_code = ArrayString(0)
                  getCampaign_name = ArrayString(1)
               End If
               
               getCampaign_code = cmbcampaigncode.text
sStrsql = ""
Select Case MDIForm1.txtlevel
        Case "Administrator", "Manager"
          
            mwhere = ""
            sStrsql = " select id from mgm   "
            mwhere = " where recsource='" + getCampaign_code + "' and (agent is null or agent='')"
            
            
            
            
            If cbofieldfilter.text <> "" Then
                    
                    If cbolimit.text <> "" And cbooperand.text = "" Then
                                If Len(mwhere) = 0 Then
                                
                                    mwhere = " where " + cbofieldfilter + "=" + CStr(cbolimit) + ""
                                Else
                            
                                 mwhere = mwhere + " and  " + cbofieldfilter + " = " + CStr(cbolimit) + ""
                                End If
                    
                        Else
                                If cbolimit.text <> "" Then
                                   If cbooperand.text <> "" Then
                                        If cbooperand.text = "between" Then
                                                If Len(mwhere) = 0 Then
                                                    mwhere = " where " + cbofieldfilter + " between " + CStr(cbolimit) + " and " + CStr(cbolimit1.text) + ""
                                                Else
                                                  mwhere = mwhere + " and  " + cbofieldfilter + " between " + CStr(cbolimit) + " and " + CStr(cbolimit1.text) + ""
                                                End If
                                        Else
                                                If Len(mwhere) = 0 Then
                                                    mwhere = " where " + cbofieldfilter + " " + cbooperand.text + " " + CStr(cbolimit)
                                                Else
                                                  mwhere = mwhere + " and " + cbofieldfilter + " " + cbooperand.text + " " + CStr(cbolimit)
                                                End If
                                        End If
                                   End If
                                   
                                End If
                    End If
            End If
            
            
            If cmbcampaigncode.text <> "" Then
                  If Len(mwhere) = 0 Then
                            mwhere = "  recsource ='" + getCampaign_code + "'"
                        Else
                            mwhere = mwhere + " and  recsource ='" + getCampaign_code + "'"
                        End If
                        
            End If
            
            Set M_objrs = New ADODB.Recordset
                M_objrs.CursorLocation = adUseClient
                M_objrs.Open sStrsql + mwhere, M_OBJCONN, adOpenDynamic, adLockOptimistic
                
                   ' txtjmlcampaign.Text = M_objrs.RecordCount
                
            Set M_objrs = Nothing
        
        
   End Select
End Sub
Public Sub isilimit()
   Dim mwhere   As String
   Dim getUserid  As String
   Dim getUsername As String
   Dim getCampaign_code As String
   Dim getCampaign_name   As String
      
               
               intvrl = InStr(1, cmbcampaigncode.text, "!", vbTextCompare)
               If intvrl <> 0 Then
                  ArrayString = Split(cmbcampaigncode.text, "!", 2, vbTextCompare)
                  getCampaign_code = ArrayString(0)
                  getCampaign_name = ArrayString(1)
               End If
                getCampaign_code = cmbcampaigncode.text

Select Case MDIForm1.txtlevel
        Case "Administrator", "Manager"
            mwhere = ""
            If cbofieldfilter <> "" Then
                sQlnew = " select distinct( " + cbofieldfilter + ") as jml from mgm   "
            Else
                Exit Sub
            End If
            
            If cbopocket <> "" Then
                If Len(mwhere) = 0 Then
                    If cbopocket.text = "New Bucket!" Then
                         mwhere = " where  agent is null"
                    Else
                         mwhere = " where  agent='" + getUserid + "'"
                    End If
                Else
                    If cbopocket.text = "New Bucket!" Then
                         mwhere = mwhere + " and  agent is null"
                    Else
                         mwhere = mwhere + " and  agent='" + getUserid + "'"
                    End If
                End If
            End If
                      
            If cbofieldfilter.text <> "" Then
                    
                    If cbolimit.text <> "" And cbooperand.text = "" Then
                                If Len(mwhere) = 0 Then
                                
                                    mwhere = " where " + cbofieldfilter + "=" + CStr(cbolimit) + ""
                                Else
                            
                                 mwhere = mwhere + " and  " + cbofieldfilter + " = " + CStr(cbolimit) + ""
                                End If
                    
                    Else
                                If cbolimit.text <> "" Then
                                   If cbooperand.text <> "" Then
                                        If cbooperand.text = "between" Then
                                                If Len(mwhere) = 0 Then
                                                    mwhere = " where " + cbofieldfilter + " between " + CStr(cbolimit) + " and " + CStr(cbolimit1.text) + ""
                                                Else
                                                  mwhere = mwhere + " and  " + cbofieldfilter + " between " + CStr(cbolimit) + " and " + CStr(cbolimit1.text) + ""
                                                End If
                                        Else
                                                If Len(mwhere) = 0 Then
                                                    mwhere = " where " + cbofieldfilter + " " + cbooperand.text + " " + CStr(cbolimit)
                                                Else
                                                  mwhere = mwhere + " and " + cbofieldfilter + " " + cbooperand.text + " " + CStr(cbolimit)
                                                End If
                                        End If
                                   End If
                                   
                                End If
                    End If
            End If
            If cbostatus.text <> "" Then
                 If cbostatus.text = "Recycle" Then
                        If Len(mwhere) = 0 Then
                            mwhere = " where bucket ='R'"
                        Else
                            mwhere = mwhere + " and  bucket ='R'"
                        End If
                 End If
                 
                 If cbostatus.text = "New" Then
                        If Len(mwhere) = 0 Then
                            mwhere = " where (bucket is null or bucket='N')"
                        Else
                            mwhere = mwhere + " and  (bucket is null or bucket='N')"
                        End If
                End If
            End If
            
             
            If cmbcampaigncode.text <> "" Then
                  If Len(mwhere) = 0 Then
                            mwhere = " WHERE recsource ='" + getCampaign_code + "'"
                        Else
                            mwhere = mwhere + " and  recsource ='" + getCampaign_code + "'"
                        End If
                        
            End If
             CBOLIMT = cbolimit
             CBOLIMT1 = cbolimit1.text
            If sQlnew <> "" Then
            Set M_objrs = New ADODB.Recordset
                M_objrs.CursorLocation = adUseClient
                M_objrs.Open sQlnew + mwhere, M_OBJCONN, adOpenDynamic, adLockOptimistic
                 cbolimit.CLEAR
                 cbolimit1.CLEAR
                 While Not M_objrs.EOF
                    If IIf(IsNull(M_objrs!jml), "", M_objrs!jml) <> "" Then
                     cbolimit.AddItem M_objrs!jml
                      cbolimit1.AddItem M_objrs!jml
                    End If
                    
                     M_objrs.MoveNext
                 Wend
                
            Set M_objrs = Nothing
            End If
           cbolimit.text = CBOLIMT
           cbolimit1.text = CBOLIMT1
 
   End Select


End Sub
Private Sub ListView7_DblClick()
If ListView7.ListItems.Count <> 0 Then
    Command2.Enabled = False
    loadbucketAgent ListView7.SelectedItem.SubItems(1), Text4.text
    Command2.Enabled = True
End If

End Sub




Private Sub LvAgent_DblClick()
    Dim setJmlDistribusi As Double
    Dim jmlDtSudahDistribusi As Double
    Dim ListItem As ListItem
    Dim m_msgbox As String

    On Error Resume Next

    setJmlDistribusi = InputBox("Inputkan jumlah data distribusi untuk:" & LvAgent.SelectedItem.SubItems(1) & "-" & LvAgent.SelectedItem.SubItems(2), "Distribusi Data")



    If Val(txtSisaCampaign.text) < setJmlDistribusi Then
        m_msgbox = MsgBox("Data melebihi jumlah sisa campaign!", vbOKOnly + vbInformation, "Informasi")
        Exit Sub
    End If


    LvAgent.SelectedItem.SubItems(5) = setJmlDistribusi

    jmlDtSudahDistribusi = 0
    For i = 1 To Val(txtJmlAgent.text)
        jmlDtSudahDistribusi = jmlDtSudahDistribusi + Val(LvAgent.ListItems.Item(i).SubItems(5))
    Next i
    txtSudahDistribusi.text = Val(ttlsudah_distribute) + Val(jmlDtSudahDistribusi)
    txtSisaCampaign.text = Val(txtjmlcampaign.text) - Val(txtSudahDistribusi.text)

End Sub

Private Sub LVMgr_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    KeluarListView ("2")
    'KeluarListView ("1")
    
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
Private Sub CreateInsert_Waterfall_Hst(sQueryid As String, AGENT As String, nmAgent As String, Level As String)
    Dim sBulan      As String
    Dim sYear       As String
    Dim sTableName  As String
    
    'sBulan = Format(FungsiWaktuServer, "mmmm")
    'sYear = Format(FungsiWaktuServer, "yyyy")
    sTableName = "tbl_mgm_hst_" & sBulan & "_" & sYear
    On Error GoTo InsertTable
    sQueryCreate = " create table  " & sTableName & "( " & _
                   " id serial, " & _
                   " id_cust integer, " & _
                   " statuscall character varying (100), " & _
                   " statuscall_old character varying (100), " & _
                   " agent character varying (100), " & _
                   " nmagent character varying (100), " & _
                   " kdspv character varying (100), " & _
                   " tglinput timestamp with time zone , " & _
                   " campaign_code character varying (100), " & _
                   " tglinput timestamp with time zone default now() " & _
                   " ) "
    M_OBJCONN.Execute sQueryCreate
'    sQueryInsert = "INSERT INTO tbl_create_waterfall_hst (table_name,bulan_create) values ('" & sTableName & "','" & sBulan & "')"
'    M_OBJCONN.Execute sQueryInsert
InsertTable:
    Level = UCase(Level)
    If Level = "AGENT" Then
    
        sQuerySelect = " select A.id as id_cust,'New Data'::text,f_cek_new,'" & AGENT & "'::text,'" & nmAgent & "'::text,SPVCODE,now(),recsource from mgm a,usertbl b where  " & _
                       "  A.id in ( " & sQueryid & "  "
    Else
        sQuerySelect = " select id as id_cust,'New Data'::text,f_cek_new,'" & AGENT & "'::text,'" & nmAgent & "'::text,'" & AGENT & "'::text,now(),recsource from mgm where  " & _
                       " id in ( " & sQueryid & "   "
    End If
                   
    sQueryInsert = " INSERT INTO " & sTableName & "(" & _
                   " id_cust,statuscall,statuscall_old,agent,nmagent,kdspv,tglinput,campaign_code " & _
                   " ) " & sQuerySelect
    M_OBJCONN.Execute sQueryInsert
End Sub
Public Sub Prosesdistribusi()
    Dim mwhere, mwhere1  As String
    Dim getUserid  As String
    Dim getUsername As String
    Dim getCampaign_code As String
    Dim getCampaign_name   As String
    '-----------------------------------------------------------------------------
    'Pecah cmbcampaigncode.text
    '-----------------------------------------------------------------------------
    intvrl = InStr(1, cmbcampaigncode.text, "!", vbTextCompare)
    If intvrl <> 0 Then
       ArrayString = Split(cmbcampaigncode.text, "!", 2, vbTextCompare)
       getCampaign_code = ArrayString(0)
       getCampaign_name = ArrayString(1)
    End If
    getCampaign_code = cmbcampaigncode.text
    
    If cbostatus.text <> "" Then
         If cbostatus.text = "Recycle" Then
            mwhere = mwhere + " and  bucket ='R'"
         End If
         
         If cbostatus.text = "New" Then
            mwhere = mwhere + " and  (bucket is null or bucket='N')"
        End If
    End If
    
    If cmbregion.text <> "" Then
            mwhere = mwhere + " and  region ='" + cmbregion.text + "'"
    End If
    
    '-----------------------------------------------------------------------------
    sStrsql = "SELECT ID FROM mgm where recsource ='" + cmbcampaigncode.text + "' and agent ='" & sUserid & "' " + mwhere + ""
    
    For i = 1 To LvAgent.ListItems.Count
        If Val(LvAgent.ListItems(i).SubItems(5)) <> 0 Then
            'mwhere1 = " AND SPVCODE=(SELECT spvcode from usertbl where userid='" & LVAgent.ListItems(i).SubItems(1) & "')"
            Call CreateInsert_Waterfall_Hst(sStrsql & mwhere1 & " order by ID)GROUP BY ID_CUST,f_cek_new,SPVCODE,recsource   limit " & LvAgent.ListItems(i).SubItems(5), LvAgent.ListItems(i).SubItems(1), LvAgent.ListItems(i).SubItems(2), "Agent")
            
            STRUPDATE = " update mgm set agent='" + LvAgent.ListItems(i).SubItems(1) + "',nama_agent='" + LvAgent.ListItems(i).SubItems(2) + "', tgldistribusi =date(now()) where id in( "
            STRUPDATE = STRUPDATE + "SELECT id_cust FROM tbl_mgm_hst__ where CAMPAIGN_CODE='" + cmbcampaigncode.text + "' AND agent='" + LvAgent.ListItems(i).SubItems(1) + "'   order by tglinput desc,ID limit " + LvAgent.ListItems(i).SubItems(5) + " )  "
            M_OBJCONN.Execute (STRUPDATE)
        
            'cmdsql_update = "insert into tbllogdistribusi(userid,nama,campaign_code,jmldata,sendby) values "
            'cmdsql_update = cmdsql_update + "('" + LVAgent.ListItems.Item(i).SubItems(1) + "','" + LVAgent.ListItems.Item(i).SubItems(2) + "','" + getCampaign_code + "'," + CStr(Val(LVAgent.ListItems(i).SubItems(5))) + ",'" + sUserid + "')"
            cmdsql_update = "insert into tbllogdistribusi(userid,nama,custid,name_ch,statuscall,region,tagihan,campaign_code,sendby,statuscall_old) "
            cmdsql_update = cmdsql_update + "select b.agent,nmagent,id_cust,name,b.statuscall,region,curbal,campaign_code,'" + MDIForm1.txtusername.text + "',statuscall_old from mgm a left join tbl_mgm_hst__ b on (a.id=b.id_cust) where b.agent='" + LvAgent.ListItems(i).SubItems(1) + "' AND campaign_code ='" + cmbcampaigncode.text + "' order by tglinput desc limit " + LvAgent.ListItems(i).SubItems(5) + " "
            M_OBJCONN.Execute (cmdsql_update)
            LvAgent.ListItems(i).SubItems(4) = Val(LvAgent.ListItems(i).SubItems(4)) + LvAgent.ListItems(i).SubItems(5)
            LvAgent.ListItems(i).SubItems(5) = ""
        End If
    Next i
    ListView4.ListItems.CLEAR
    ListView5.ListItems.CLEAR
    ListView6.ListItems.CLEAR
     
    
    Set M_OBJRS_tglhistory = Nothing
    Set M_objrs = Nothing
    Set M_OBJRS_history = Nothing
    Set M_OBJRS_history3 = Nothing
    'PB.Value = 0
    cmdloadcampaign_Click
    cmdProses.Enabled = True
    m_msgbox = MsgBox("Proses distribusi berhasil!", vbOKOnly + vbInformation, "Informasi")
End Sub
Public Sub summeryCampaign()
Dim TOTALSPACE As Double
Dim TOTALALREADY As Double
Dim ListItem  As ListItem
Dim strsql As String
Dim MOBJ As New ADODB.Recordset
Set MOBJ = New ADODB.Recordset
MOBJ.CursorLocation = adUseClient
strsql = strsql + " select * from ("
strsql = strsql + " SELECT alldata.CAMPAIGN_CODE as batch,JML_DATA as total_lead, AVAILABLE_SPACE as space_lead FROM ("
strsql = strsql + " select recsource ,COUNT(id) AS JML_DATA from mgm "
strsql = strsql + "  GROUP by recsource) AS ALLDATA LEFT JOIN "
strsql = strsql + " ( "
strsql = strsql + " select recsource ,COUNT(id) AS AVAILABLE_SPACE from mgm  WHERE (AGENT IS NULL OR AGENT='' ) "
strsql = strsql + " GROUP by recsource) AS TBLSPACE  ON ALLDATA.CAMPAIGN_CODE=TBLSPACE.CAMPAIGN_CODE ) as tblsatu left join "

strsql = strsql + " (select recsource ,COUNT(id) AS ALREADY_ASSIGN from mgm  WHERE (AGENT IS NOT NULL and AGENT<>'')"
strsql = strsql + " GROUP by recsource) AS TBLASSIGN  ON TBLSatu.batch=TBLASSIGN.recsource ORDER BY BATCH "



MOBJ.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
ListView1.ListItems.CLEAR
TOTALALREADY = 0
TOTALSPACE = 0
While Not MOBJ.EOF
Set ListItem = ListView1.ListItems.ADD(, , MOBJ.Bookmark)
      ListItem.SubItems(1) = IIf(IsNull(MOBJ!batch), "", MOBJ!batch)
      ListItem.SubItems(2) = IIf(IsNull(MOBJ!total_lead), "", MOBJ!total_lead)
      ListItem.SubItems(3) = IIf(IsNull(MOBJ!space_lead), "", MOBJ!space_lead)
      NILSPACE = IIf(IsNull(MOBJ!space_lead), 0, MOBJ!space_lead)
      TOTALSPACE = TOTALSPACE + Val(NILSPACE)
      ListItem.SubItems(4) = IIf(IsNull(MOBJ!ALREADY_ASSIGN), "", MOBJ!ALREADY_ASSIGN)
      NILALREADY = IIf(IsNull(MOBJ!ALREADY_ASSIGN), 0, MOBJ!ALREADY_ASSIGN)
      TOTALALREADY = TOTALALREADY + Val(NILALREADY)
      MOBJ.MoveNext
Wend
 Warna_Row_Listview Form_distribute, ListView1, &HFFFFC0, vbWhite
txtavailable.text = TOTALSPACE
txtalreadyassign = TOTALALREADY
End Sub
Public Sub summerybyTL()
Dim TOTALSPACE As Double
Dim TOTALALREADY As Double
Dim ListItem  As ListItem
Dim strsql As String
Dim MOBJ1 As New ADODB.Recordset
Dim MOBJ As New ADODB.Recordset
Set MOBJ = New ADODB.Recordset
MOBJ.CursorLocation = adUseClient
'STRSQL = " SELECT * FROM("
'STRSQL = STRSQL + " SELECT CAMPAIGN_CODE,TEAM,COUNT(NO_CASE) AS JML FROM ("
'STRSQL = STRSQL + " SELECT * FROM MGM WHERE AGENT IN (SELECT tbluser_userid FROM tbluser WHERE TEAM IN (SELECT DISTINCT(TEAM) FROM USERTBL WHERE USERTYPE='6'))) TBLMGM ,tbluser"
'STRSQL = STRSQL + " WHERE TBLMGM.AGENT=tbluser.tbluser_userid GROUP BY CAMPAIGN_CODE,TEAM) as ggg ORDER BY TEAM, CAMPAIGN_CODE"


strsql = " SELECT TEAM,CAMPAIGN_CODE,SUM(JML) FROM ("

strsql = strsql + " SELECT usertbl.spvcode as team, count(id) as jml,recsource FROM MGM ,usertbl where mgm.agent=usertbl.userid  and usertbl.kdlevel=2 group by spvcode  ,recsource"
strsql = strsql + " union all("
strsql = strsql + " select agent as team, count(id) as jml,recsource from mgm where  agent in (select usertbl.userid from usertbl WHERE  kdlevel='2') group by agent,recsource )"
strsql = strsql + " ) A GROUP BY TEAM,CAMPAIGN_CODE "


ListView2.ListItems.CLEAR
MOBJ.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'ProgressBar1.Max = MOBJ.RecordCount
    
     NILALREADY = 0
     TOTALSPACE = 0
     txtalreadytl.text = ""
     
While Not MOBJ.EOF
TOTALALREADY = 0
NILALREADY = 0
 'ProgressBar1.Value = MOBJ.Bookmark
 DoEvents
Set ListItem = ListView2.ListItems.ADD(, , MOBJ.Bookmark)
      ListItem.SubItems(1) = IIf(IsNull(MOBJ!TEAM), "", MOBJ!TEAM)
      sTEAM = IIf(IsNull(MOBJ!TEAM), "", MOBJ!TEAM)
      scampaign = IIf(IsNull(MOBJ!campaign_code), "", MOBJ!campaign_code)
      ListItem.SubItems(2) = IIf(IsNull(MOBJ!campaign_code), "", MOBJ!campaign_code)
      
      Set MOBJ1 = New ADODB.Recordset
      MOBJ1.CursorLocation = adUseClient
      strsql = "select count(*) as jml from mgm where campaign_code ='" + scampaign + "'"
      MOBJ1.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
      ListItem.SubItems(3) = IIf(IsNull(MOBJ1!jml), "0", MOBJ1!jml)
      NILALREADY = IIf(IsNull(MOBJ1!jml), "0", MOBJ1!jml)
      TOTALALREADY = TOTALALREADY + Val(NILALREADY)
      Set MOBJ1 = Nothing
      
      
      Set MOBJ1 = New ADODB.Recordset
      MOBJ1.CursorLocation = adUseClient
      strsql = "select count(*) as jml from mgm where agent in (select userid from usertbl where spvcode= '" + sTEAM + "' and kdlevel =2) and recsource ='" + scampaign + "'"
      MOBJ1.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
      ListItem.SubItems(5) = IIf(IsNull(MOBJ1!jml), "0", MOBJ1!jml)
      txtalreadytl.text = Val(txtalreadytl.text) + MOBJ1!jml
      NILALREADY = IIf(IsNull(MOBJ1!jml), "0", MOBJ1!jml)
      TOTALALREADY = TOTALALREADY + Val(NILALREADY)
      Set MOBJ1 = Nothing
      
      Set MOBJ1 = New ADODB.Recordset
      MOBJ1.CursorLocation = adUseClient
      strsql = "select count(*) as jml from mgm where agent in (select userid from usertbl where userid= '" + sTEAM + "' and  kdlevel =2) and recsource ='" + scampaign + "'"
      MOBJ1.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
      ListItem.SubItems(4) = IIf(IsNull(MOBJ1!jml), "0", MOBJ1!jml)
      txtavalabletl.text = Val(txtavalabletl.text) + MOBJ1!jml
      NILSPACE = IIf(IsNull(MOBJ1!jml), "0", MOBJ1!jml)
      TOTALSPACE = TOTALSPACE + Val(NILSPACE)
      Set MOBJ1 = Nothing
      MOBJ.MoveNext
      
Wend

 Warna_Row_Listview Form_distribute, ListView2, &HFFFFC0, vbWhite
txtavalabletl.text = TOTALSPACE
'txtalreadytl.Text = TOTALALREADY

End Sub
Public Sub summerybyAGENT()
Dim TOTALSPACE As Double
Dim TOTALALREADY As Double
Dim ListItem  As ListItem
Dim strsql As String
Dim MOBJ1 As New ADODB.Recordset
Dim MOBJ As New ADODB.Recordset
Set MOBJ = New ADODB.Recordset
MOBJ.CursorLocation = adUseClient
strsql = " SELECT * FROM( SELECT CAMPAIGN_CODE,TBLMGM.AGENT,COUNT(NO_CASE) AS JML FROM"
strsql = strsql + " ( SELECT recsource,AGENT,id FROM MGM WHERE AGENT IN (SELECT userid FROM usertbl WHERE userid IN (SELECT DISTINCT(userid)"
strsql = strsql + " FROM usertbl WHERE kdlevel ='2'))) TBLMGM ,usertbl WHERE TBLMGM.AGENT=usertbl.userid GROUP BY CAMPAIGN_CODE,TBLMGM.AGENT)  ORDER BY AGENT, CAMPAIGN_CODE"
 
ListView3.ListItems.CLEAR
MOBJ.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
Text2.text = ""

While Not MOBJ.EOF
Set ListItem = ListView3.ListItems.ADD(, , MOBJ.Bookmark)
      ListItem.SubItems(1) = IIf(IsNull(MOBJ!AGENT), "", MOBJ!AGENT)
      ListItem.SubItems(2) = IIf(IsNull(MOBJ!campaign_code), "", MOBJ!campaign_code)
      ListItem.SubItems(3) = IIf(IsNull(MOBJ!jml), "", MOBJ!jml)
      Text2.text = Val(Text2.text) + Val(IIf(IsNull(MOBJ!jml), "", MOBJ!jml))
      MOBJ.MoveNext
Wend
Warna_Row_Listview Form_distribute, ListView3, &HFFFFC0, vbWhite
'txtavalabletl.Text = TOTALSPACE
'txtalreadytl.Text = TOTALALREADY

End Sub
Public Sub isicombo_opendate()
 Dim M_objrs As New ADODB.Recordset
     ListView5.ListItems.CLEAR
       intvrl = InStr(1, cmbcampaigncode.text, "!", vbTextCompare)
               If intvrl <> 0 Then
                  ArrayString = Split(cmbcampaigncode.text, "!", 2, vbTextCompare)
                  getCampaign_code = ArrayString(0)
                  getCampaign_name = ArrayString(1)
               End If
               getCampaign_code = cmbcampaigncode.text 'HENDRI CODE
               
 Set M_objrs = New ADODB.Recordset
 M_objrs.CursorLocation = adUseClient
 strsql = "select DISTINCT(tgl_trans) from mgm where recsource='" + getCampaign_code + "' ORDER BY tgl_trans ASC"
 M_objrs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
 
 Tgl1 = cbopendate1.text
 Tgl2 = cbopendate2.text
 cbopendate1.CLEAR
 cbopendate2.CLEAR
 While Not M_objrs.EOF
        cbopendate1.AddItem Format(IIf(IsNull(M_objrs!tgl_trans), "", M_objrs!tgl_trans), "yyyy-mm-dd")
        cbopendate2.AddItem Format(IIf(IsNull(M_objrs!tgl_trans), "", M_objrs!tgl_trans), "yyyy-mm-dd")
        M_objrs.MoveNext
 Wend
 cbopendate1.text = Tgl1
 cbopendate2.text = Tgl2
End Sub
Public Sub loadbucketTL()
Dim M_objrs As ADODB.Recordset
Dim strsql As String
Dim ListItem As ListItem
   
If Combo1.text = Empty Then
    m_msgbox = MsgBox("Textbox campaign code tidak boleh kosong!", vbOKOnly + vbExclamation, "Informasi")
    Exit Sub
End If
    
 intvrl = InStr(1, Combo1.text, "!", vbTextCompare)
               If intvrl <> 0 Then
                  ArrayString = Split(Combo1.text, "!", 2, vbTextCompare)
                  getCampaign_code = ArrayString(0)
                  getCampaign_name = ArrayString(1)
               End If
               getCampaign_code = cmbcampaigncode.text 'HENDRI CODE
               
If Check1.Value = 1 Then '<-- kalo di cek berarti pakai like
    sUseSimiliar = " LIKE '%" + Combo1.text + "%'"
Else
    sUseSimiliar = " = '" + Combo1.text + "'"
End If
    
strsql = " SELECT * FROM ("
strsql = strsql + " SELECT U.userid,U.agent,COALESCE(ACU.JML,0) AS JML FROM (SELECT userid,agent FROM usertbl WHERE kdlevel = 2 AND aktif= 1) U"
strsql = strsql + " LEFT JOIN (SELECT AGENT,COUNT(AGENT) AS JML FROM MGM"
strsql = strsql + " WHERE recsource " + sUseSimiliar + " AND AGENT IN (SELECT userid FROM usertbl WHERE kdlevel  = 2 AND aktif = 1)"
strsql = strsql + " GROUP BY AGENT ORDER BY JML"
strsql = strsql + " ) AS ACU ON(ACU.AGENT=U.userid)) AS GG ORDER BY JML DESC,userid"
   
Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    ListView7.ListItems.CLEAR
    hjml = 0
    no = 0
    While Not M_objrs.EOF
        'Menginputkan data ke listview
        no = no + 1
        Set list = ListView7.ListItems.ADD(, , no)
        list.SubItems(1) = IIf(IsNull(M_objrs!USERID), "", M_objrs!USERID)
        list.SubItems(2) = IIf(IsNull(M_objrs!AGENT), "", M_objrs!AGENT)
        list.SubItems(3) = IIf(IsNull(M_objrs!jml), 0, M_objrs!jml)
        If list.SubItems(1) <> "AM1" Then
            hjml = hjml + list.SubItems(3)
        End If
        M_objrs.MoveNext
    Wend
    Warna_Row_Listview Form_distribute, ListView7, &HFFFFC0, vbWhite
    Text1.text = hjml '-> jumlah all
    Set M_objrs = Nothing
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 2 Then


        Combo1.text = cmbcampaigncode.text
        additemcombo1
    End If
End Sub
Public Sub additemcombo1()
    Dim M_objrs As ADODB.Recordset
    Dim CMDSQL As String
    
    'Mengisi data ke combo campaigncode
    CMDSQL = "select kodeds from datasourcetbl order by tglentry DESC"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    
    M_objrs.Open CMDSQL, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    Combo1.CLEAR
    'Combo1.AddItem "<Select More Campaign>"
    While Not M_objrs.EOF
        Combo1.AddItem M_objrs("kodeds")
        M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
    
End Sub
Public Sub loadbucketAgent(sTl As String, sCamapign As String)
Dim M_objrs As ADODB.Recordset
Dim strsql As String
Dim ListItem As ListItem
   
If sCamapign = Empty Or sTl = Empty Then
    m_msgbox = MsgBox("Parameter uncomplete!", vbOKOnly + vbExclamation, "Informasi")
    Exit Sub
End If
    
 intvrl = InStr(1, sCamapign, "!", vbTextCompare)
               If intvrl <> 0 Then
                  ArrayString = Split(sCamapign, "!", 2, vbTextCompare)
                  getCampaign_code = ArrayString(0)
                  getCampaign_name = ArrayString(1)
               End If
                getCampaign_code = cmbcampaigncode.text 'HENDRI CODE
    
If Check1.Value = 1 Then '<-- kalo di cek berarti pakai like
    sUseSimiliar = " LIKE '%" + sCamapign + "%'"
Else
    sUseSimiliar = " = '" + sCamapign + "'"
End If

strsql = " SELECT * FROM ("
strsql = strsql + " SELECT U.userid,U.agent,COALESCE(ACU.JML,0) AS JML FROM (SELECT userid,agent FROM usertbl WHERE spvcode = '" + sTl + "' AND  aktif ='1' AND kdlevel='2') U"
strsql = strsql + " LEFT JOIN ("
strsql = strsql + " SELECT AGENT,COUNT(AGENT) AS JML FROM MGM"
strsql = strsql + " WHERE recsource " + sUseSimiliar + " AND AGENT IN (SELECT userid FROM usertbl WHERE spvcode = '" + sTl + "' AND aktif ='1'  AND kdlevel='2' )"
strsql = strsql + " GROUP BY AGENT "
strsql = strsql + " ORDER BY JML"
strsql = strsql + " ) AS ACU ON(ACU.AGENT=U.userid)"
strsql = strsql + " ) AS GG"
strsql = strsql + " ORDER BY userid = '" + sTl + "' DESC,JML DESC"

Set M_objrs = New ADODB.Recordset
M_objrs.CursorLocation = adUseClient
M_objrs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    ListView8.ListItems.CLEAR
    hjmls = 0
    no = 0
    While Not M_objrs.EOF
        'Menginputkan data ke listview
        no = no + 1
        Set list = ListView8.ListItems.ADD(, , no)
        list.SubItems(1) = IIf(IsNull(M_objrs!USERID), "", M_objrs!USERID)
        list.SubItems(2) = IIf(IsNull(M_objrs!AGENT), "", M_objrs!AGENT)
        list.SubItems(3) = IIf(IsNull(M_objrs!jml), 0, M_objrs!jml)
        If no > 1 Then
            hjmls = hjmls + list.SubItems(3)
        End If
        M_objrs.MoveNext
    Wend
    Warna_Row_Listview Form_distribute, ListView8, &HFFFFC0, vbWhite
    Text3.text = hjmls '-> jumlah all
    Set M_objrs = Nothing
End Sub
Private Sub Text4_Click()
Combo1.text = Text4.text
End Sub




