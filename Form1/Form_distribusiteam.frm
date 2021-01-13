VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form_distribusiteam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Distribute to agent"
   ClientHeight    =   10140
   ClientLeft      =   645
   ClientTop       =   -300
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10140
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   9285
      Left            =   -15
      TabIndex        =   0
      Top             =   810
      Width           =   18240
      _ExtentX        =   32173
      _ExtentY        =   16378
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Criteria Distribute"
      TabPicture(0)   =   "Form_distribusiteam.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image3(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label10"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label20"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label11"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label8"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label26"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label29"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ProgressBar1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cbofieldfilter"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtJumlah(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cbopendate2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cbopendate1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Command1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cbostatus"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdloadcampaign"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Frame1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Frame2"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cbolimit1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cbooperand"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cbolimit"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmbcampaigncode"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "SSTab2"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Check3"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cboarea"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Combo2"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Check4"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Check5"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Command5"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).ControlCount=   31
      TabCaption(1)   =   "History Distribute"
      TabPicture(1)   =   "Form_distribusiteam.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label21"
      Tab(1).Control(1)=   "Label19"
      Tab(1).Control(2)=   "Image3(2)"
      Tab(1).Control(3)=   "Label27"
      Tab(1).Control(4)=   "Label28"
      Tab(1).Control(5)=   "ListView6"
      Tab(1).Control(6)=   "ListView5"
      Tab(1).Control(7)=   "ListView4"
      Tab(1).Control(8)=   "CMBHISTORY"
      Tab(1).Control(9)=   "Combo3"
      Tab(1).Control(10)=   "Command4"
      Tab(1).Control(11)=   "Text5"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Bucket Monitoring"
      TabPicture(2)   =   "Form_distribusiteam.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame8"
      Tab(2).Control(1)=   "Frame7"
      Tab(2).ControlCount=   2
      Begin VB.CommandButton Command5 
         Caption         =   "TARIK NEW DATA"
         Height          =   375
         Left            =   1080
         TabIndex        =   92
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Active"
         Height          =   375
         Left            =   1395
         TabIndex        =   90
         Top             =   1200
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Not Active"
         Height          =   375
         Left            =   2475
         TabIndex        =   89
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -67440
         TabIndex        =   88
         Top             =   8880
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00F1E5DB&
         Caption         =   "&Load Data"
         Height          =   375
         Left            =   -68280
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   360
         Width           =   1440
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   -73545
         TabIndex        =   84
         Top             =   360
         Width           =   5040
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   5880
         TabIndex        =   83
         Top             =   1680
         Visible         =   0   'False
         Width           =   4470
      End
      Begin VB.ComboBox cboarea 
         Height          =   315
         ItemData        =   "Form_distribusiteam.frx":0054
         Left            =   4650
         List            =   "Form_distribusiteam.frx":0061
         TabIndex        =   82
         Top             =   1500
         Visible         =   0   'False
         Width           =   4185
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Jml Assign"
         Height          =   345
         Left            =   4080
         TabIndex        =   80
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   8835
         Left            =   8850
         TabIndex        =   54
         Top             =   840
         Visible         =   0   'False
         Width           =   9435
         _ExtentX        =   16642
         _ExtentY        =   15584
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Leads By Campaign"
         TabPicture(0)   =   "Form_distribusiteam.frx":007E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame4"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame5"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame6"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Tracking Distribusi"
         TabPicture(1)   =   "Form_distribusiteam.frx":009A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Command3"
         Tab(1).Control(1)=   "Check2"
         Tab(1).Control(2)=   "TDBDate1"
         Tab(1).Control(3)=   "TDBDate2"
         Tab(1).Control(4)=   "ListView9"
         Tab(1).Control(5)=   "Label25"
         Tab(1).Control(6)=   "Label4(3)"
         Tab(1).ControlCount=   7
         Begin VB.CommandButton Command3 
            Caption         =   "Go"
            Height          =   345
            Left            =   -69990
            TabIndex        =   74
            Top             =   450
            Width           =   1095
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Include data Recycle"
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   -74790
            TabIndex        =   73
            Top             =   840
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   2025
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Data yang sudah di assign ke agent "
            Height          =   3705
            Left            =   150
            TabIndex        =   69
            Top             =   4770
            Width           =   9180
            Begin VB.TextBox Text2 
               Appearance      =   0  'Flat
               Height          =   345
               Left            =   1395
               TabIndex        =   70
               Top             =   3180
               Width           =   1125
            End
            Begin MSComctlLib.ListView ListView3 
               Height          =   2865
               Left            =   90
               TabIndex        =   71
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
               TabIndex        =   72
               Top             =   3240
               Width           =   1095
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Status Data Di TL"
            Height          =   2265
            Left            =   120
            TabIndex        =   62
            Top             =   2520
            Width           =   9180
            Begin VB.TextBox txtalreadytl 
               Appearance      =   0  'Flat
               Height          =   345
               Left            =   4410
               TabIndex        =   64
               Top             =   1830
               Width           =   1125
            End
            Begin VB.TextBox txtavalabletl 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2010
               TabIndex        =   63
               Top             =   1860
               Width           =   1125
            End
            Begin MSComctlLib.ListView ListView2 
               Height          =   1515
               Left            =   90
               TabIndex        =   65
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
               TabIndex        =   68
               Top             =   1890
               Width           =   1095
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "Available"
               Height          =   345
               Left            =   1320
               TabIndex        =   67
               Top             =   1890
               Width           =   1095
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Jumlah Lead :"
               Height          =   255
               Left            =   150
               TabIndex        =   66
               Top             =   1890
               Width           =   1245
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Lead By Campaign"
            Height          =   2265
            Left            =   120
            TabIndex        =   55
            Top             =   270
            Width           =   9180
            Begin VB.TextBox txtavailable 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2010
               TabIndex        =   57
               Top             =   1860
               Width           =   1125
            End
            Begin VB.TextBox txtalreadyassign 
               Appearance      =   0  'Flat
               Height          =   345
               Left            =   4410
               TabIndex        =   56
               Top             =   1830
               Width           =   1125
            End
            Begin MSComctlLib.ListView ListView1 
               Height          =   1605
               Left            =   90
               TabIndex        =   58
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
               TabIndex        =   61
               Top             =   1890
               Width           =   1245
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Available"
               Height          =   345
               Left            =   1215
               TabIndex        =   60
               Top             =   1890
               Width           =   1095
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Already Send"
               Height          =   345
               Left            =   3210
               TabIndex        =   59
               Top             =   1890
               Width           =   1095
            End
         End
         Begin TDBDate6Ctl.TDBDate TDBDate1 
            Height          =   315
            Left            =   -73830
            TabIndex        =   75
            Top             =   450
            Width           =   1725
            _Version        =   65536
            _ExtentX        =   3043
            _ExtentY        =   556
            Calendar        =   "Form_distribusiteam.frx":00B6
            Caption         =   "Form_distribusiteam.frx":01CE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Form_distribusiteam.frx":023A
            Keys            =   "Form_distribusiteam.frx":0258
            Spin            =   "Form_distribusiteam.frx":02B6
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
         Begin TDBDate6Ctl.TDBDate TDBDate2 
            Height          =   315
            Left            =   -71760
            TabIndex        =   76
            Top             =   450
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   556
            Calendar        =   "Form_distribusiteam.frx":02DE
            Caption         =   "Form_distribusiteam.frx":03F6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "Form_distribusiteam.frx":0462
            Keys            =   "Form_distribusiteam.frx":0480
            Spin            =   "Form_distribusiteam.frx":04DE
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
         Begin MSComctlLib.ListView ListView9 
            Height          =   7155
            Left            =   -74820
            TabIndex        =   77
            Top             =   1110
            Width           =   8865
            _ExtentX        =   15637
            _ExtentY        =   12621
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
         Begin VB.Label Label25 
            BackColor       =   &H00E87211&
            BackStyle       =   0  'Transparent
            Caption         =   "Tgl Distribusi :"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   -74850
            TabIndex        =   79
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "to"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Index           =   3
            Left            =   -72390
            TabIndex        =   78
            Top             =   495
            Width           =   885
         End
      End
      Begin VB.ComboBox cmbcampaigncode 
         Height          =   315
         Left            =   1455
         TabIndex        =   39
         Top             =   405
         Width           =   5760
      End
      Begin VB.ComboBox cbolimit 
         Height          =   315
         Left            =   5040
         TabIndex        =   38
         Top             =   1080
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ComboBox cbooperand 
         Height          =   315
         ItemData        =   "Form_distribusiteam.frx":0506
         Left            =   4590
         List            =   "Form_distribusiteam.frx":051C
         TabIndex        =   37
         Top             =   1110
         Visible         =   0   'False
         Width           =   1950
      End
      Begin VB.ComboBox cbolimit1 
         Height          =   315
         Left            =   6525
         MousePointer    =   1  'Arrow
         TabIndex        =   36
         Top             =   1110
         Visible         =   0   'False
         Width           =   2295
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
         Height          =   855
         Left            =   60
         TabIndex        =   29
         Top             =   2160
         Width           =   8670
         Begin VB.TextBox txtSudahDistribusi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H0000FF00&
            Height          =   285
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   32
            Text            =   "0"
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtSisaCampaign 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   7530
            Locked          =   -1  'True
            TabIndex        =   31
            Text            =   "0"
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtjmlcampaign 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   30
            Text            =   "0"
            Top             =   360
            Width           =   975
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
            Left            =   3120
            TabIndex        =   35
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Sisa Lead :"
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
            Left            =   5880
            TabIndex        =   34
            Top             =   360
            Width           =   1665
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
            Index           =   0
            Left            =   240
            TabIndex        =   33
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Distribusikan data ke officer :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6240
         Left            =   90
         TabIndex        =   22
         Top             =   2970
         Width           =   8670
         Begin VB.TextBox txtJmlAgent 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2190
            Locked          =   -1  'True
            TabIndex        =   25
            Text            =   "0"
            Top             =   5805
            Width           =   975
         End
         Begin VB.CommandButton cmdkeluar 
            BackColor       =   &H00F1E5DB&
            Caption         =   "&Keluar"
            Height          =   495
            Left            =   6945
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   5625
            Width           =   1575
         End
         Begin VB.CommandButton cmdProses 
            BackColor       =   &H00F1E5DB&
            Caption         =   "&PROSES"
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   5640
            Width           =   1575
         End
         Begin MSComctlLib.ListView LVSpv 
            Height          =   5250
            Left            =   90
            TabIndex        =   26
            Top             =   300
            Width           =   8505
            _ExtentX        =   15002
            _ExtentY        =   9260
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
         Begin MSComctlLib.ProgressBar PB 
            Height          =   255
            Left            =   3465
            TabIndex        =   27
            Top             =   5850
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
         End
         Begin VB.Label Label9 
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
            Left            =   165
            TabIndex        =   28
            Top             =   5805
            Width           =   2055
         End
      End
      Begin VB.CommandButton cmdloadcampaign 
         BackColor       =   &H00F1E5DB&
         Caption         =   "&Load Data"
         Height          =   285
         Left            =   7320
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   360
         Width           =   1440
      End
      Begin VB.ComboBox cbostatus 
         Height          =   315
         ItemData        =   "Form_distribusiteam.frx":0539
         Left            =   1470
         List            =   "Form_distribusiteam.frx":0543
         TabIndex        =   20
         Top             =   750
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00F1E5DB&
         Caption         =   "&View All"
         Height          =   375
         Left            =   8955
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.ComboBox CMBHISTORY 
         Height          =   315
         Left            =   -62280
         TabIndex        =   18
         Top             =   495
         Width           =   2295
      End
      Begin VB.ComboBox cbopendate1 
         Height          =   315
         Left            =   3960
         TabIndex        =   17
         Top             =   1440
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.ComboBox cbopendate2 
         Height          =   315
         Left            =   4650
         TabIndex        =   16
         Top             =   1860
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.TextBox txtJumlah 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   6990
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1920
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.ComboBox cbofieldfilter 
         Height          =   315
         Left            =   5850
         TabIndex        =   14
         Top             =   750
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Frame Frame7 
         Height          =   855
         Left            =   -74790
         TabIndex        =   8
         Top             =   390
         Width           =   9795
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF80&
            Height          =   315
            Left            =   7140
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   210
            Width           =   2475
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Similiar Search (use % char)"
            Height          =   225
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Visible         =   0   'False
            Width           =   2265
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00F1E5DB&
            Caption         =   "&Load Data"
            Height          =   375
            Left            =   4080
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   150
            Width           =   1755
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1500
            TabIndex        =   9
            Top             =   180
            Width           =   2445
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
            TabIndex        =   13
            Top             =   210
            Width           =   1455
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Data"
         Height          =   6435
         Left            =   -74790
         TabIndex        =   1
         Top             =   1290
         Width           =   9825
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8730
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   2280
            Width           =   855
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8730
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   6030
            Width           =   855
         End
         Begin MSComctlLib.ListView ListView7 
            Height          =   2055
            Left            =   60
            TabIndex        =   4
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
            TabIndex        =   5
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
         Begin VB.Label Label23 
            Caption         =   "Total all TL Bucket:"
            Height          =   315
            Left            =   7290
            TabIndex        =   7
            Top             =   2310
            Width           =   1665
         End
         Begin VB.Label Label24 
            Caption         =   "Total all agent Bucket:"
            Height          =   315
            Left            =   7080
            TabIndex        =   6
            Top             =   6060
            Width           =   2055
         End
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   10755
         TabIndex        =   40
         Top             =   405
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   7755
         Left            =   -74880
         TabIndex        =   41
         Top             =   840
         Width           =   8490
         _ExtentX        =   14975
         _ExtentY        =   13679
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
         NumItems        =   0
      End
      Begin MSComctlLib.ListView ListView5 
         Height          =   4020
         Left            =   -63795
         TabIndex        =   42
         Top             =   1035
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
      Begin VB.Label Label29 
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
         Left            =   240
         TabIndex        =   91
         Top             =   1320
         Width           =   1050
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   255
         Left            =   -68040
         TabIndex        =   87
         Top             =   8880
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label27 
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
         TabIndex        =   85
         Top             =   405
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
         Left            =   3810
         TabIndex        =   81
         Top             =   1560
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   18630
         Index           =   2
         Left            =   -74955
         Picture         =   "Form_distribusiteam.frx":0555
         Top             =   315
         Width           =   26295
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
         TabIndex        =   52
         Top             =   450
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H00F1E5DB&
         BackStyle       =   0  'Transparent
         Caption         =   "Saving Type"
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
         Left            =   6600
         TabIndex        =   51
         Top             =   1320
         Visible         =   0   'False
         Width           =   1275
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
         Left            =   3780
         TabIndex        =   50
         Top             =   1155
         Visible         =   0   'False
         Width           =   1275
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
         Left            =   600
         TabIndex        =   49
         Top             =   810
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Pilih Tanggal  "
         Height          =   255
         Left            =   -63840
         TabIndex        =   48
         Top             =   540
         Width           =   2295
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL PER AGENT"
         Height          =   255
         Left            =   -61710
         TabIndex        =   47
         Top             =   5220
         Width           =   3375
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
         Left            =   3690
         TabIndex        =   46
         Top             =   1890
         Visible         =   0   'False
         Width           =   975
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
         Left            =   5400
         TabIndex        =   45
         Top             =   1560
         Visible         =   0   'False
         Width           =   1095
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
         Left            =   5250
         TabIndex        =   44
         Top             =   780
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Image Image3 
         Height          =   18630
         Index           =   0
         Left            =   45
         Picture         =   "Form_distribusiteam.frx":7B5F
         Top             =   315
         Width           =   26295
      End
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
      Left            =   630
      TabIndex        =   53
      Top             =   300
      Width           =   3585
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   1
      Left            =   135
      Picture         =   "Form_distribusiteam.frx":F169
      Stretch         =   -1  'True
      Top             =   240
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   825
      Left            =   -1950
      Picture         =   "Form_distribusiteam.frx":FC73
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20160
   End
End
Attribute VB_Name = "Form_distribusiteam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ttl_sudahdistribusi As Double
Private Sub cbofieldfilter_DropDown()
sStrsql = " SELECT column_name as nama_kolom  From information_schema.Columns WHERE table_name='mgm' and data_type in ('numeric') ORDER BY ordinal_position "
Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    cbofieldfilter.CLEAR
    While Not M_objrs.EOF
        cbofieldfilter.AddItem IIf(IsNull(M_objrs!nama_kolom), "", M_objrs!nama_kolom)
        M_objrs.MoveNext
    Wend
  Set M_objrs = Nothing
End Sub

Private Sub cbofieldfilter_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub cbooperand_Click()
If cbooperand.Text = "between" Then
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
       cbopocket.AddItem MDIForm1.TxtUsername.Text & "!" & MDIForm1.txtnama
    Case "Assisten Manager"
       sStrsql = "select * from  tbluser where tbluser_kdlevel<>1 AND  tbluser_amcode  ='" + MDIForm1.TxtUsername.Text + "'"
       Set M_objrs = New ADODB.Recordset
           M_objrs.CursorLocation = adUseClient
           M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
       cbopocket.CLEAR
      cbopocket.AddItem MDIForm1.TxtUsername.Text & "!" & MDIForm1.txtnama
       While Not M_objrs.EOF
            cbopocket.AddItem IIf(IsNull(M_objrs!tbluser_userid), "", M_objrs!tbluser_userid) & "!" & IIf(IsNull(M_objrs!tbluser_name), "", M_objrs!tbluser_name)
            M_objrs.MoveNext
       Wend
    Case "Manager"
    sStrsql = "select * from  tbluser where tbluser_kdlevel<>'1'"
       Set M_objrs = New ADODB.Recordset
           M_objrs.CursorLocation = adUseClient
           M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
       cbopocket.CLEAR
       cbopocket.AddItem "New Bucket!"
     cbopocket.AddItem MDIForm1.TxtUsername.Text & "!" & MDIForm1.txtnama
       While Not M_objrs.EOF
            cbopocket.AddItem IIf(IsNull(M_objrs!tbluser_userid), "", M_objrs!tbluser_userid) & "!" & IIf(IsNull(M_objrs!tbluser_name), "", M_objrs!tbluser_name)
            M_objrs.MoveNext
       Wend
       
    Case "Administrator"
       sStrsql = "select * from  tbluser where tbluser_kdlevel<>'1'"
       Set M_objrs = New ADODB.Recordset
           M_objrs.CursorLocation = adUseClient
           M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
       cbopocket.CLEAR
       cbopocket.AddItem "New Bucket!"
       cbopocket.AddItem MDIForm1.TxtUsername.Text & "!" & MDIForm1.txtnama
       
       
       While Not M_objrs.EOF
            cbopocket.AddItem IIf(IsNull(M_objrs!tbluser_userid), "", M_objrs!tbluser_userid) & "!" & IIf(IsNull(M_objrs!tbluser_name), "", M_objrs!tbluser_name)
            M_objrs.MoveNext
       Wend
       
End Select

End Sub

Private Sub cbopocket_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Check4_Click()
If Check4.Value = vbChecked Then
Check5.Value = vbUnchecked
End If
End Sub

Private Sub Check5_Click()
If Check5.Value = vbChecked Then
Check4.Value = vbUnchecked
End If

End Sub

Private Sub cmbcampaigncode_Click()
    Dim M_objrs As ADODB.Recordset
    Dim M_OBJRS_history As ADODB.Recordset
    Dim M_OBJRS_history3 As ADODB.Recordset
    Dim M_OBJRS_tglhistory As ADODB.Recordset
    Dim cmdsql As String
    Dim cmdsql_history As String
    Dim cmdsql_history3 As String
    Dim cmdsql_tglhistory As String
    Dim ListItem As ListItem
    Dim getCampaign_code As String
    Dim getCampaign_name As String

    
     intvrl = InStr(1, cmbcampaigncode.Text, "!", vbTextCompare)
               If intvrl <> 0 Then
                  ArrayString = Split(cmbcampaigncode.Text, "!", 2, vbTextCompare)
                  getCampaign_code = ArrayString(0)
                  getCampaign_name = ArrayString(1)
               End If
               getCampaign_code = cmbcampaigncode.Text 'HENDRI CODE
               
    ListView4.ListItems.CLEAR
    ListView5.ListItems.CLEAR
    ListView6.ListItems.CLEAR
    
    
    
    ' ISI HISTORY DISTRIBUSI
    
        cmdsql_history = "select USERID,NAMA,CAMPAIGN_CODE,JMLDATA,SENDBY,TGL from tbllogdistribusi where campaign_code='" + getCampaign_code + "' and sendby = '" + MDIForm1.TxtUsername.Text + "' ORDER BY TGL DESC"
        Set M_OBJRS_history = New ADODB.Recordset
        M_OBJRS_history.CursorLocation = adUseClient
        M_OBJRS_history.Open cmdsql_history, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        
        While Not M_OBJRS_history.EOF
            Set ListItem = ListView4.ListItems.ADD(, , IIf(IsNull(M_OBJRS_history("USERID")), "", M_OBJRS_history("USERID")))
            ListItem.SubItems(1) = IIf(IsNull(M_OBJRS_history("NAMA")), "", M_OBJRS_history("NAMA"))
            ListItem.SubItems(2) = IIf(IsNull(M_OBJRS_history("CAMPAIGN_CODE")), "", M_OBJRS_history("CAMPAIGN_CODE"))
            ListItem.SubItems(3) = IIf(IsNull(M_OBJRS_history("JMLDATA")), "", M_OBJRS_history("JMLDATA"))
            ListItem.SubItems(4) = IIf(IsNull(M_OBJRS_history("SENDBY")), "", M_OBJRS_history("SENDBY"))
            ListItem.SubItems(5) = IIf(IsNull(M_OBJRS_history("TGL")), "", Format(M_OBJRS_history("TGL"), "dd-mmm-yyyy"))
            M_OBJRS_history.MoveNext
        Wend

        'isi kombo tanggal history
        cmdsql_tglhistory = "SELECT * FROM (select DATE(TGL) AS TGL from tbllogdistribusi where campaign_code='" + getCampaign_code + "' and sendby = '" + MDIForm1.TxtUsername.Text + "' GROUP BY TGL) AS TBLBARU GROUP BY TGL  ORDER BY TGL DESC"
        Set M_OBJRS_tglhistory = New ADODB.Recordset
        M_OBJRS_tglhistory.CursorLocation = adUseClient
        M_OBJRS_tglhistory.Open cmdsql_tglhistory, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        CMBHISTORY.CLEAR
        While Not M_OBJRS_tglhistory.EOF
            CMBHISTORY.AddItem Format(M_OBJRS_tglhistory("TGL"), "YYYY-MM-DD")
            M_OBJRS_tglhistory.MoveNext
        Wend
    'ISI KOMBO TOTAL HISTORY
    
        ListView6.ListItems.CLEAR
        cmdsql_history3 = "select USERID, SUM(JMLDATA) AS TOTAL from tbllogdistribusi where campaign_code='" + getCampaign_code + "' and sendby = '" + MDIForm1.TxtUsername.Text + "' GROUP BY USERID ORDER BY USERID"
        Set M_OBJRS_history3 = New ADODB.Recordset
        M_OBJRS_history3.CursorLocation = adUseClient
        M_OBJRS_history3.Open cmdsql_history3, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
        While Not M_OBJRS_history3.EOF
            Set ListItem = ListView6.ListItems.ADD(, , IIf(IsNull(M_OBJRS_history3("USERID")), "", M_OBJRS_history3("USERID")))
            ListItem.SubItems(1) = IIf(IsNull(M_OBJRS_history3("TOTAL")), "", M_OBJRS_history3("TOTAL"))
            M_OBJRS_history3.MoveNext
        Wend
    
    Set M_OBJRS_tglhistory = Nothing
    Set M_objrs = Nothing
    Set M_OBJRS_history = Nothing
    Set M_OBJRS_history3 = Nothing

End Sub

Private Sub cmbcampaigncode_DropDown()
'sstrsql = "select * from tbldatasource where   tbldatasource_kdstatus ='1' and tbldatasource_campaign_code in (select campaign_code from mgm where agent in (select tbluser_userid from tbluser where (tbluser_groupspvcode ='" + MDIForm1.txtUserName.Text + "' or tbluser_userid='" + MDIForm1.txtUserName.Text + "' )) ) order by   tbldatasource_tglentry,  tbldatasource_keterangan "
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
       intvrl = InStr(1, cmbcampaigncode.Text, "!", vbTextCompare)
               If intvrl <> 0 Then
                  ArrayString = Split(cmbcampaigncode.Text, "!", 2, vbTextCompare)
                  getCampaign_code = ArrayString(0)
                  getCampaign_name = ArrayString(1)
               End If
               getCampaign_code = cmbcampaigncode.Text 'HENDRI CODE
    
        cmdsql_history2 = "select USERID, SUM(JMLDATA) AS TOTAL from tbllogdistribusi where campaign_code='" + getCampaign_code + "' and sendby = '" + MDIForm1.TxtUsername.Text + "' AND DATE(TGL) = '" + CMBHISTORY.Text + "' GROUP BY USERID ORDER BY USERID"
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

Private Sub cmdkeluar_Click()
Unload Me
End Sub
Private Sub cekField()
Dim M_objrs As New ADODB.Recordset
sStrsql = "SELECT KETERANGAN FROM TBL_SETTING WHERE USERID = '" + MDIForm1.TxtUsername.Text + "' LIMIT 1"
Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    If M_objrs.RecordCount <> 0 Then
       While Not M_objrs.EOF
           cbofieldfilter.Text = IIf(IsNull(M_objrs!keterangan), "", M_objrs!keterangan)
           M_objrs.MoveNext
       Wend
    End If
Set M_objrs = Nothing
End Sub
Private Sub addField()
Dim M_objrs As New ADODB.Recordset
If cbofieldfilter.Text <> Empty Then
    sStrsql = "SELECT KETERANGAN FROM TBL_SETTING WHERE USERID = '" + MDIForm1.TxtUsername.Text + "' LIMIT 1"
    Set M_objrs = New ADODB.Recordset
        M_objrs.CursorLocation = adUseClient
        M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
        If M_objrs.RecordCount = 0 Then
            strsql = "INSERT INTO TBL_SETTING (KETERANGAN,TGL_INSERT,USERID) VALUES ('" + cbofieldfilter.Text + "',now(),'" + MDIForm1.TxtUsername.Text + "')"
            M_OBJCONN.Execute (strsql)
        Else
            strsql = "UPDATE TBL_SETTING SET KETERANGAN = '" + cbofieldfilter.Text + "',TIME_LASTUPDATE=now() WHERE USERID = '" + MDIForm1.TxtUsername.Text + "'"
            M_OBJCONN.Execute (strsql)
        End If
    Set M_objrs = Nothing
End If
End Sub
Private Sub cmdloadcampaign_Click()
    Dim mobjrec As New ADODB.Recordset
    If cmbcampaigncode.Text = "" Then
         MsgBox "Campaign code harus diisi", vbInformation + vbOKOnly, "Pesan"
         Exit Sub
    End If
    
    intvrl = InStr(1, cmbcampaigncode.Text, "!", vbTextCompare)
    If intvrl <> 0 Then
        ArrayString = Split(cmbcampaigncode.Text, "!", 2, vbTextCompare)
        getCampaign_code = ArrayString(0)
        getCampaign_name = ArrayString(1)
    End If
    getCampaign_code = cmbcampaigncode.Text 'hendri code
    strsql = "select id_CUST from mgm where campaign_code ='" + cmbcampaigncode.Text + "' AND  agent ='" + MDIForm1.TxtUsername.Text + "' and statuscall='New Data' and flag_recycle=0"
    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseClient
    Rs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    
    ttl_belumdistribusi = Rs.RecordCount
    Rs.Close
    Set Rs = Nothing
    strsql = "select id_CUST  from mgm where campaign_code ='" + cmbcampaigncode.Text + "' and agent in (select tbluser_userid  from tbluser where  tbluser_groupspvcode ='" + MDIForm1.TxtUsername.Text + "' and TBLUSER_KDLEVEL='1')"
    Set Rs = New ADODB.Recordset
    Rs.CursorLocation = adUseClient
    Rs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    ttl_sudahdistribusi = Rs.RecordCount
    txtSudahDistribusi.Text = Rs.RecordCount
    
    Rs.Close
    Set Rs = Nothing
    
    
    txtjmlcampaign.Text = Val(ttl_belumdistribusi) + Val(ttl_sudahdistribusi)
    isidetailuser
    txtSisaCampaign = Val(txtjmlcampaign.Text) - Val(txtSudahDistribusi.Text)
End Sub

Private Sub cmdProses_Click()
cmdProses.Enabled = False
Prosesdistribusi
cmdloadcampaign_Click
cmdProses.Enabled = True
End Sub

Private Sub Combo1_DropDown()
sStrsql = "select * from tbldatasource where   tbldatasource_kdstatus ='1' order by   tbldatasource_tglentry,  tbldatasource_keterangan "
Set M_objrs = New ADODB.Recordset
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    Combo1.CLEAR
    While Not M_objrs.EOF
        Combo1.AddItem IIf(IsNull(M_objrs!tbldatasource_campaign_code), "", M_objrs!tbldatasource_campaign_code) & "!" & IIf(IsNull(M_objrs!tbldatasource_keterangan), "", M_objrs!tbldatasource_keterangan)
        M_objrs.MoveNext
    Wend
Set M_objrs = Nothing

End Sub
Private Sub header_distribusi_Spv()
    
    LVSpv.ColumnHeaders.ADD 1, , "No", 5 * TXT
    LVSpv.ColumnHeaders.ADD 2, , "Kode", 8 * TXT
    LVSpv.ColumnHeaders.ADD 3, , "Nama", 31 * TXT
    LVSpv.ColumnHeaders.ADD 4, , "Level", 15 * TXT
    
    LVSpv.ColumnHeaders.ADD 5, , "Jumlah Total", 7 * TXT
    LVSpv.ColumnHeaders.ADD 6, , "Jumlah Awal", 15 * TXT
    
    ListView1.ColumnHeaders.ADD 1, , "NO", 5 * TXT
    ListView1.ColumnHeaders.ADD 2, , "BATCH", 15 * TXT
    ListView1.ColumnHeaders.ADD 3, , "ALL DATA", 10 * TXT
    ListView1.ColumnHeaders.ADD 4, , "AVAILABLE", 10 * TXT
    ListView1.ColumnHeaders.ADD 5, , "ALREADY ASSIGN", 10 * TXT
    
    
    
End Sub

Private Sub Combo2_DropDown()
Dim MOBJ As New ADODB.Recordset
Set MOBJ = New ADODB.Recordset
    MOBJ.CursorLocation = adUseClient
    strsql = "select distinct(saving_type) as saving  from mgm "
    MOBJ.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
   Combo2.CLEAR
   
   While Not MOBJ.EOF
        Combo2.AddItem IIf(IsNull(MOBJ!saving), "", MOBJ!saving)
        MOBJ.MoveNext
   Wend
End Sub

Private Sub Combo3_DropDown()
sStrsql = "select * from tbldatasource where   tbldatasource_kdstatus ='1'  order by   tbldatasource_tglentry,  tbldatasource_keterangan "
Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
Combo3.CLEAR
    While Not M_objrs.EOF
        'cmbcampaigncode.AddItem IIf(IsNull(M_OBJRS!tbldatasource_campaign_code), "", M_OBJRS!tbldatasource_campaign_code) & "!" & IIf(IsNull(M_OBJRS!tbldatasource_keterangan), "", M_OBJRS!tbldatasource_keterangan)
        Combo3.AddItem IIf(IsNull(M_objrs!tbldatasource_campaign_code), "", M_objrs!tbldatasource_campaign_code)
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
If Combo1.Text <> Empty Then
    Text4.Text = Combo1.Text
End If

End Sub

Private Sub Command3_Click()
trackdistribusi

End Sub

Private Sub Command4_Click()
Dim M_OBJRS_history As New ADODB.Recordset
     cmdsql_history = "select USERID,NAMA,CAMPAIGN_CODE,JMLDATA,SENDBY,TGL from tbllogdistribusi where campaign_code='" + Combo2.Text + "' and sendby = '" + MDIForm1.TxtUsername.Text + "' ORDER BY TGL DESC"
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

Private Sub Command5_Click()
    Form_recycleNewData.Show 1
End Sub

Private Sub Form_Load()
    header_distribusi_Spv
    HEADER
    HEADER_TRACKDIS
    cekField
    SSTab1.TabVisible(1) = True
    SSTab1.TabVisible(2) = False
    sStrsql = "select * from tbldatasource where   tbldatasource_kdstatus ='1'  and (tbldatasource_tglexpired   > DATE(NOW()) OR tbldatasource_tglexpired IS NULL  ) order by   tbldatasource_campaign_code"
Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    DoEvents
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
    cmbcampaigncode.CLEAR
    While Not M_objrs.EOF
    DoEvents
        'cmbcampaigncode.AddItem IIf(IsNull(M_OBJRS!tbldatasource_campaign_code), "", M_OBJRS!tbldatasource_campaign_code) & "!" & IIf(IsNull(M_OBJRS!tbldatasource_keterangan), "", M_OBJRS!tbldatasource_keterangan)
        cmbcampaigncode.AddItem IIf(IsNull(M_objrs!tbldatasource_campaign_code), "", M_objrs!tbldatasource_campaign_code)
        M_objrs.MoveNext
    Wend
    M_objrs.Close
    
Set M_objrs = Nothing


End Sub



Private Sub HEADER()
    ListView4.ColumnHeaders.ADD 1, , "Batch", 20 * TXT
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
               
               intvrl = InStr(1, cmbcampaigncode.Text, "!", vbTextCompare)
               If intvrl <> 0 Then
                  ArrayString = Split(cmbcampaigncode.Text, "!", 2, vbTextCompare)
                  getCampaign_code = ArrayString(0)
                  getCampaign_name = ArrayString(1)
               End If
               getCampaign_code = cmbcampaigncode.Text 'hendri code
               
sStrsql = ""
Select Case MDIForm1.txtlevel
        Case "Supervisor"
            mwhere = ""
            sStrsql = " select no_case from mgm   "
            'mwhere = " where campaign_code='" + getCampaign_code + "' and agent='" + MDIForm1.txtUserName.Text + "'"
            mwhere = " where campaign_code='" + getCampaign_code + "' and ( agent in (SELECT tbluser_userid FROM TBLUSER WHERE tbluser_groupspvcode ='" + MDIForm1.TxtUsername.Text + "') OR agent='" + MDIForm1.TxtUsername.Text + "') "
            
            
            
            
            If cbofieldfilter.Text <> "" Then
                    
                    If cbolimit.Text <> "" And cbooperand.Text = "" Then
                                If Len(mwhere) = 0 Then
                                
                                    mwhere = " where " + cbofieldfilter + " = " + CStr(cbolimit) + ""
                                Else
                            
                                 mwhere = mwhere + " and  " + cbofieldfilter + " = " + CStr(cbolimit) + ""
                                End If
                    
                        Else
                                If cbolimit.Text <> "" Then
                                   If cbooperand.Text <> "" Then
                                        If cbooperand.Text = "between" Then
                                                If Len(mwhere) = 0 Then
                                                    mwhere = " where " + cbofieldfilter + " between " + CStr(cbolimit) + " and " + CStr(cbolimit1.Text) + ""
                                                Else
                                                  mwhere = mwhere + " and  " + cbofieldfilter + " between " + CStr(cbolimit) + " and " + CStr(cbolimit1.Text) + ""
                                                End If
                                        Else
                                                If Len(mwhere) = 0 Then
                                                    mwhere = " where " + cbofieldfilter + " " + cbooperand.Text + " " + CStr(cbolimit)
                                                Else
                                                  mwhere = mwhere + " and " + cbofieldfilter + " " + cbooperand.Text + " " + CStr(cbolimit)
                                                End If
                                        End If
                                   End If
                                   
                                End If
                    End If
            End If
            
            If cbopendate1.Text <> "" And cbopendate2.Text <> "" Then
                     If Len(mwhere) = 0 Then
                    mwhere = " WHERE date(tgl_trans) between '" + cbopendate1.Text + "'  and '" + cbopendate2.Text + "'"
                    Else
                   mwhere = mwhere + " AND date(tgl_trans) between '" + cbopendate1.Text + "'  and '" + cbopendate2.Text + "'"
                    End If
             End If
             
             
           
             If cboarea.Text <> "" Then
                   If cboarea.Text = "Jakarta" Then
                      If Len(mwhere) = 0 Then
                      
                            mwhere = " where  ( city_1  ilike 'JAKARTA%' OR  CITY_1 ILIKE 'BODETABEK' )"
                      Else
                        mwhere = mwhere + " AND ( city_1  ilike 'JAKARTA%' OR  CITY_1 ILIKE 'BODETABEK' )"
                      End If
                      
                   ElseIf cboarea.Text = "Luar Jakarta" Then
                        If Len(mwhere) = 0 Then
                            mwhere = " where  city_1 not ilike 'JAKARTA%' AND CITY_1 NOT ILIKE 'BODETABEK'"
                        Else
                            mwhere = mwhere + " AND  city_1 not ilike 'JAKARTA%' AND CITY_1 NOT ILIKE 'BODETABEK'"
                        End If
                      
                   End If
             End If
             
             
             
             
'            If cbostatus.Text <> "" Then
'                 If cbostatus.Text = "Recycle" Then
'                        If Len(mwhere) = 0 Then
'                            mwhere = " where bucket ='R'"
'                        Else
'                            mwhere = mwhere + " and  bucket ='R'"
'                        End If
'                 End If
'
'                 If cbostatus.Text = "New" Then
'                        If Len(mwhere) = 0 Then
'                            mwhere = " where (bucket is null or bucket='N')"
'                        Else
'                            mwhere = mwhere + " and  (bucket is null or bucket='N')"
'                        End If
'                End If
'            End If
            
            
                
            
            
            
            If cmbcampaigncode.Text <> "" Then
                  If Len(mwhere) = 0 Then
                            mwhere = "  campaign_code ='" + getCampaign_code + "'"
                        Else
                            mwhere = mwhere + " and  campaign_code ='" + getCampaign_code + "'"
                        End If
                        
            End If
            
            
             If Combo2.Text <> "" Then
                  If Len(mwhere) = 0 Then
                            mwhere = "  where saving_type ='" + Combo2.Text + "'"
                  Else
                            mwhere = mwhere + " and  saving_type ='" + Combo2.Text + "'"
                  End If
            End If
             
'            If Len(mwhere) = 0 Then
'                    mwhere = " where sts_move = 0"
'            Else
'                            mwhere = mwhere + " and  sts_move = 0"
'            End If
'
            
            Set M_objrs = New ADODB.Recordset
                M_objrs.CursorLocation = adUseClient
                M_objrs.Open sStrsql + mwhere, M_OBJCONN, adOpenDynamic, adLockOptimistic
              
                    txtjmlcampaign.Text = M_objrs.RecordCount
              
            Set M_objrs = Nothing
        
        
   End Select
End Sub
Public Sub isilimit()
   Dim mwhere   As String
   Dim getUserid  As String
   Dim getUsername As String
   Dim getCampaign_code As String
   Dim getCampaign_name   As String
      
               
               intvrl = InStr(1, cmbcampaigncode.Text, "!", vbTextCompare)
               If intvrl <> 0 Then
                  ArrayString = Split(cmbcampaigncode.Text, "!", 2, vbTextCompare)
                  getCampaign_code = ArrayString(0)
                  getCampaign_name = ArrayString(1)
               End If
               getCampaign_code = cmbcampaigncode.Text 'hendri code

Select Case MDIForm1.txtlevel
        Case "Supervisor"
            mwhere = ""
            If cbofieldfilter <> "" Then
                sQlnew = " select distinct( " + cbofieldfilter + ") as jml from mgm   "
            Else
                Exit Sub
            End If
                    
             mwhere = "where agent ='" + MDIForm1.TxtUsername.Text + "'"
            
            If cbofieldfilter.Text <> "" Then
                    
                    If cbolimit.Text <> "" And cbooperand.Text = "" Then
                                If Len(mwhere) = 0 Then
                                
                                    mwhere = " where " + cbofieldfilter + "=" + CStr(cbolimit) + ""
                                Else
                            
                                 mwhere = mwhere + " and  " + cbofieldfilter + " = " + CStr(cbolimit) + ""
                                End If
                    
                        Else
                                If cbolimit.Text <> "" Then
                                   If cbooperand.Text <> "" Then
                                        If cbooperand.Text = "between" Then
                                                If Len(mwhere) = 0 Then
                                                    mwhere = " where " + cbofieldfilter + " between " + CStr(cbolimit) + " and " + CStr(cbolimit1.Text) + ""
                                                Else
                                                  mwhere = mwhere + " and  " + cbofieldfilter + " between " + CStr(cbolimit) + " and " + CStr(cbolimit1.Text) + ""
                                                End If
                                        Else
                                                If Len(mwhere) = 0 Then
                                                    mwhere = " where " + cbofieldfilter + " " + cbooperand.Text + " " + CStr(cbolimit)
                                                Else
                                                  mwhere = mwhere + " and " + cbofieldfilter + " " + cbooperand.Text + " " + CStr(cbolimit)
                                                End If
                                        End If
                                   End If
                                   
                                End If
                    End If
            End If
            If cbostatus.Text <> "" Then
                 If cbostatus.Text = "Recycle" Then
                        If Len(mwhere) = 0 Then
                            mwhere = " where bucket ='R'"
                        Else
                            mwhere = mwhere + " and  bucket ='R'"
                        End If
                 End If
                 
                 If cbostatus.Text = "New" Then
                        If Len(mwhere) = 0 Then
                            mwhere = " where (bucket is null or bucket='N')"
                        Else
                            mwhere = mwhere + " and  (bucket is null or bucket='N')"
                        End If
                End If
            End If
            
            
            
             If cboarea.Text <> "" Then
                   If cboarea.Text = "Jakarta" Then
                      If Len(mwhere) = 0 Then
                      
                            mwhere = " where  ( city_1  ilike 'JAKARTA%' OR  CITY_1 ILIKE 'BODETABEK' )"
                      Else
                        mwhere = mwhere + " AND ( city_1  ilike 'JAKARTA%' OR  CITY_1 ILIKE 'BODETABEK' )"
                      End If
                      
                   ElseIf cboarea.Text = "Luar Jakarta" Then
                        If Len(mwhere) = 0 Then
                            mwhere = " where  city_1 not ilike 'JAKARTA%' AND CITY_1 NOT ILIKE 'BODETABEK'"
                        Else
                            mwhere = mwhere + " AND  city_1 not ilike 'JAKARTA%' AND CITY_1 NOT ILIKE 'BODETABEK'"
                        End If
                      
                   End If
             End If
             
             
             
             If cbopendate1.Text <> "" And cbopendate2.Text <> "" Then
                     If Len(mwhere) = 0 Then
                    mwhere = " WHERE date(tgl_trans) between '" + cbopendate1.Text + "'  and '" + cbopendate2.Text + "'"
                    Else
                   mwhere = mwhere + " AND date(tgl_trans) between '" + cbopendate1.Text + "'  and '" + cbopendate2.Text + "'"
                    End If
             End If
             
             
             
             
            If cmbcampaigncode.Text <> "" Then
                  If Len(mwhere) = 0 Then
                            mwhere = " WHERE campaign_code ='" + getCampaign_code + "'"
                        Else
                            mwhere = mwhere + " and  campaign_code ='" + getCampaign_code + "'"
                        End If
                        
            End If
            
            
            If Combo2.Text <> "" Then
                  If Len(mwhere) = 0 Then
                            mwhere = "  where saving_type ='" + Combo2.Text + "'"
                  Else
                            mwhere = mwhere + " and  saving_type ='" + Combo2.Text + "'"
                  End If
            End If
            
            If Len(mwhere) = 0 Then
                    mwhere = " where sts_move = 0"
            Else
                            mwhere = mwhere + " and  sts_move = 0"
            End If
            
            
             
            
             CBOLIMT = cbolimit
             CBOLIMT1 = cbolimit1.Text
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
            cbolimit.Text = CBOLIMT
              cbolimit1.Text = CBOLIMT1
        
   End Select


End Sub
Public Sub isidetailuser()
  
    Dim mwhere   As String
    Dim getUserid  As String
    Dim getUsername As String
    Dim getCampaign_code As String
    Dim getCampaign_name   As String
     
               
    intvrl = InStr(1, cmbcampaigncode.Text, "!", vbTextCompare)
    If intvrl <> 0 Then
        ArrayString = Split(cmbcampaigncode.Text, "!", 2, vbTextCompare)
        getCampaign_code = ArrayString(0)
        getCampaign_name = ArrayString(1)
    End If
    getCampaign_code = cmbcampaigncode.Text 'hendri code
                
    Select Case MDIForm1.txtlevel
        Case "Supervisor"
            mwhere = ""
            If cbofieldfilter.Text <> "" Then
                If cbolimit.Text <> "" And cbooperand.Text = "" Then
                    mwhere = mwhere + " and  " + cbofieldfilter + " = " + CStr(cbolimit) + ""
                Else
                    If cbolimit.Text <> "" Then
                        If cbooperand.Text <> "" Then
                            If cbooperand.Text = "between" Then
                                mwhere = mwhere + " and  " + cbofieldfilter + " between " + CStr(cbolimit) + " and " + CStr(cbolimit1.Text) + ""
                            Else
                                mwhere = mwhere + " and " + cbofieldfilter + " " + cbooperand.Text + " " + CStr(cbolimit)
                            End If
                        End If
            
                    End If
                End If
            End If
            If cbostatus.Text <> "" Then
                If cbostatus.Text = "Recycle" Then
                    mwhere = mwhere + " and  bucket ='R'"
                End If
                If cbostatus.Text = "New" Then
                    mwhere = mwhere + " and  (bucket is null or bucket='N')"
                End If
            End If
            
            If cmbcampaigncode.Text <> "" Then
                mwhere = mwhere + " and  campaign_code ='" + getCampaign_code + "'"
            End If
            
            If cbopendate1.Text <> "" And cbopendate2.Text <> "" Then
                mwhere = mwhere + " AND date(tgl_trans) between '" + cbopendate1.Text + "'  and '" + cbopendate2.Text + "'"
            End If
            If cboarea.Text <> "" Then
                If cboarea.Text = "Jakarta" Then
                    mwhere = mwhere + " AND ( city_1  ilike 'JAKARTA%' OR  CITY_1 ILIKE 'BODETABEK' )"
                ElseIf cboarea.Text = "Luar Jakarta" Then
                    mwhere = mwhere + " AND  city_1 not ilike 'JAKARTA%' AND CITY_1 NOT ILIKE 'BODETABEK'"
                End If
            End If
        
    
            If Combo2.Text <> "" Then
                mwhere = mwhere + " and  saving_type ='" + Combo2.Text + "'"
            End If
    
            If Len(mwhere) = 0 Then
                mwhere = mwhere + " and  sts_move = 0"
            End If
            If Check3.Value = vbUnchecked Then
                If Check5.Value = vbChecked Then
                sStrsql = " select a.tbluser_userid,a.tbluser_name, tbluser_ketlevel  from (SELECT * FROM TBLUSER WHERE (TBLUSER_KDLEVEL='1') AND TBLUSER_KDSTATUS = '1' AND tbluser_groupspvcode ='" + MDIForm1.TxtUsername.Text + "' and date(tgl_login) =date(now())order by tbluser_userid) a "
                ElseIf Check4.Value = vbChecked Then
                
                sStrsql = " select a.tbluser_userid,a.tbluser_name, tbluser_ketlevel  from (SELECT * FROM TBLUSER WHERE (TBLUSER_KDLEVEL='1') AND TBLUSER_KDSTATUS = '1' AND tbluser_groupspvcode ='" + MDIForm1.TxtUsername.Text + "' and (date(tgl_login) <date(now()) or tgl_login is null)) a "
                Else
                sStrsql = " select a.tbluser_userid,a.tbluser_name, tbluser_ketlevel  from (SELECT * FROM TBLUSER WHERE (TBLUSER_KDLEVEL='1') AND TBLUSER_KDSTATUS = '1' AND tbluser_groupspvcode ='" + MDIForm1.TxtUsername.Text + "') a "
                End If
            Else
                If Check3.Value = vbChecked Then
                    If Check5.Value = vbChecked Then
                        sStrsql = " select a.tbluser_userid,a.tbluser_name,a.tbluser_ketlevel ,b.jml from (SELECT * FROM TBLUSER WHERE (TBLUSER_KDLEVEL='1') AND TBLUSER_KDSTATUS = '1' AND tbluser_groupspvcode ='" + MDIForm1.TxtUsername.Text + "'    and date(tgl_login) =date(now()) order by tbluser_userid) a "
                        sStrsql = sStrsql + " Left Join "
                        sStrsql = sStrsql + "  ( "
                        sStrsql = sStrsql + " SELECT AGENT,count(id_cust)  AS JML FROM MGM,TBLUSER  "
                        sStrsql = sStrsql + "  where MGM.AGENT=TBLUSER_USERID " + mwhere + "  GROUP BY AGENT) b on a.tbluser_userid=b.agent "
                    ElseIf Check4.Value = vbChecked Then
                        sStrsql = " select a.tbluser_userid,a.tbluser_name,a.tbluser_ketlevel ,b.jml from (SELECT * FROM TBLUSER WHERE (TBLUSER_KDLEVEL='1') AND TBLUSER_KDSTATUS = '1' AND tbluser_groupspvcode ='" + MDIForm1.TxtUsername.Text + "'  and  (date(tgl_login) <date(now()) or tgl_login is null) ) a "
                        sStrsql = sStrsql + " Left Join "
                        sStrsql = sStrsql + "  ( "
                        sStrsql = sStrsql + " SELECT AGENT,count(id_cust)  AS JML FROM MGM,TBLUSER  "
                        sStrsql = sStrsql + "  where MGM.AGENT=TBLUSER_USERID " + mwhere + "  GROUP BY AGENT) b on a.tbluser_userid=b.agent "
                    Else
                        sStrsql = " select a.tbluser_userid,a.tbluser_name,a.tbluser_ketlevel ,b.jml from (SELECT * FROM TBLUSER WHERE (TBLUSER_KDLEVEL='1') AND TBLUSER_KDSTATUS = '1' AND tbluser_groupspvcode ='" + MDIForm1.TxtUsername.Text + "') a "
                        sStrsql = sStrsql + " Left Join "
                        sStrsql = sStrsql + "  ( "
                        sStrsql = sStrsql + " SELECT AGENT,count(id_cust)  AS JML FROM MGM,TBLUSER  "
                        sStrsql = sStrsql + "  where MGM.AGENT=TBLUSER_USERID " + mwhere + "  GROUP BY AGENT) b on a.tbluser_userid=b.agent "
                    End If
                End If
    
            End If
    
    
    End Select
    
    
    
    'Koneksi untuk mengambil data Supervisor
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    M_objrs.Open sStrsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    LVSpv.ListItems.CLEAR
    While Not M_objrs.EOF
        'Menginputkan data ke listview
        no = no + 1
        Set list = LVSpv.ListItems.ADD(, , no)
        list.SubItems(1) = IIf(IsNull(M_objrs!tbluser_userid), "", M_objrs!tbluser_userid)
        list.SubItems(2) = IIf(IsNull(M_objrs!tbluser_name), "", M_objrs!tbluser_name)
        list.SubItems(3) = IIf(IsNull(M_objrs!tbluser_ketlevel), "", M_objrs!tbluser_ketlevel)
        If Check3.Value = vbUnchecked Then
            list.SubItems(4) = 0
        Else
            list.SubItems(4) = IIf(IsNull(M_objrs!jml), 0, M_objrs!jml)
        End If
        M_objrs.MoveNext
    Wend
    Warna_Row_Listview Form_distribusiteam, LVSpv, &HFFFFC0, vbWhite
    txtJmlAgent.Text = M_objrs.RecordCount '-> isi jumlah spv ke txtjmlagent dan txtsisacampaign
    txtSisaCampaign.Text = txtjmlcampaign.Text
    cmdProses.Enabled = True
    
    Set M_objrs = Nothing
End Sub

Private Sub ListView7_DblClick()
If ListView7.ListItems.Count <> 0 Then
    Command2.Enabled = False
    loadbucketAgent ListView7.SelectedItem.SubItems(1), Text4.Text
    Command2.Enabled = True
End If

End Sub

Private Sub LVSpv_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    If ColumnHeader.Index = 1 Then
'        If LVSpv.SortOrder = 0 Then
'
'        Call SortColumn(LVSpv, ColumnHeader.Index, sortDescending, sortNumeric)
'
'   Else
'        Call SortColumn(LVSpv, ColumnHeader.Index, sortAscending, sortNumeric)
'   End If
'
'  Else
'   ' LVSearchMgm.SortKey = ColumnHeader.Index
'   If LVSpv.SortOrder = 0 Then
'         Call SortColumn(LVSpv, ColumnHeader.Index, sortDescending, sortAlpha)
'   Else
'        Call SortColumn(LVSpv, ColumnHeader.Index, sortAscending, sortAlpha)
'   End If
'   End If


End Sub

Private Sub LVSpv_DblClick()
    Dim setJmlDistribusi As Double
    Dim jmlDtSudahDistribusi As Double
    Dim ListItem As ListItem
    Dim m_msgbox As String
    
    On Error Resume Next
    
    'Cek jumlah data di listview
    If LVSpv.ListItems.Count = 0 Then
       MsgBox "Data agent tidak ada!", vbOKOnly + vbInformation, "Informasi"
       Exit Sub
    End If
    
    setJmlDistribusi = InputBox("Inputkan jumlah data distribusi untuk:" & LVSpv.SelectedItem.SubItems(1) & "-" & LVSpv.SelectedItem.SubItems(2), "Distribusi Data")
    
    
    
    If Val(txtSisaCampaign.Text) < setJmlDistribusi Then
        m_msgbox = MsgBox("Data melebihi jumlah sisa campaign!", vbOKOnly + vbInformation, "Informasi")
        Exit Sub
    End If
     
    
    LVSpv.SelectedItem.SubItems(5) = setJmlDistribusi
    
    jmlDtSudahDistribusi = 0
    For i = 1 To Val(txtJmlAgent.Text)
    
        jmlDtSudahDistribusi = jmlDtSudahDistribusi + Val(LVSpv.ListItems.Item(i).SubItems(5))
     
    Next i

   txtSudahDistribusi.Text = Val(ttl_sudahdistribusi) + Val(jmlDtSudahDistribusi)
 
    txtSisaCampaign.Text = Val(txtjmlcampaign.Text) - Val(txtSudahDistribusi.Text)
    
   ' txtSudahDistribusi.Text = Val(txtSudahDistribusi.Text) + Val(jmlDtSudahDistribusi)
    'txtSisaCampaign.Text = Val(txtjmlcampaign.Text) - jmlDtSudahDistribusi
    'If txtSudahDistribusi.Text = "" Then
    'Else
    'txtSisaCampaign.Text = Val(txtjmlcampaign.Text) - CCur(txtSudahDistribusi.Text)
    'End If
End Sub
Private Sub CreateInsert_Waterfall_Hst(sQueryid As String, AGENT As String, nmAgent As String)
    Dim sBulan      As String
    Dim sYear       As String
    Dim sTableName  As String
    
    sBulan = Format(FungsiWaktuServer, "mmmm")
    sYear = Format(FungsiWaktuServer, "yyyy")
    sTableName = "tbl_mgm_hst_" & sBulan & "_" & sYear
    On Error GoTo InsertTable
    sQueryCreate = " create table  " & sTableName & "( " & _
                   " id serial, " & _
                   " id_cust integer, " & _
                   " statuscall character varying (100), " & _
                   " reasoncall character varying (100), " & _
                   " agent character varying (100), " & _
                   " nmagent character varying (100), " & _
                   " kdspv character varying (100), " & _
                   " nmspv character varying (100), " & _
                   " tglcall timestamp with time zone , " & _
                   " campaign_code character varying (100), " & _
                   " campaign_name  character varying (100), " & _
                   " tglinput timestamp with time zone default now() " & _
                   " ) "
    M_OBJCONN.Execute sQueryCreate
    sQueryInsert = "INSERT INTO tbl_create_waterfall_hst (table_name,bulan_create) values ('" & sTableName & "','" & sBulan & "')"
    M_OBJCONN.Execute sQueryInsert
InsertTable:
    sQuerySelect = " select id_CUST as id_cust,'New Data'::text,'New Data'::text,'" & AGENT & "'::text,'" & nmAgent & "'::text,tbluser_groupspvcode,tbluser_ketgroupspv,now(),campaign_code,campaign_name from mgm a,tbluser b where a.agent=b.tbluser_userid " & _
                   " and id_CUST in ( " & sQueryid & " ) and statuscall <> 'Agree'  "
                   
    sQueryInsert = " INSERT INTO " & sTableName & "(" & _
                   " id_cust,statuscall,reasoncall,agent,nmagent,kdspv,nmspv,tglcall,campaign_code,campaign_name " & _
                   " ) " & sQuerySelect
    M_OBJCONN.Execute sQueryInsert
End Sub
Public Sub Prosesdistribusi()
    Dim mwhere   As String
    Dim getUserid  As String
    Dim getUsername As String
    Dim getCampaign_code As String
    Dim getCampaign_name   As String
        
               
    intvrl = InStr(1, cmbcampaigncode.Text, "!", vbTextCompare)
    If intvrl <> 0 Then
       ArrayString = Split(cmbcampaigncode.Text, "!", 2, vbTextCompare)
       getCampaign_code = ArrayString(0)
       getCampaign_name = ArrayString(1)
    End If
    getCampaign_code = cmbcampaigncode.Text 'HENDRI CODE
               
               
    sStrsql = ""
    Select Case MDIForm1.txtlevel
        Case "Supervisor"
            mwhere = ""
            sStrsql = ""
            sStrsql = " select id_CUST as jml from mgm   "
            mwhere = " where campaign_code='" + getCampaign_code + "' and (agent='" + MDIForm1.TxtUsername.Text + "') and statuscall='New Data'"
            If cbopendate1.Text <> "" And cbopendate2.Text <> "" Then
                cmdsql = cmdsql + " AND date(tgl_trans) between '" + cbopendate1.Text + "'  and '" + cbopendate2.Text + "'"
            End If
            If cbofieldfilter.Text <> "" Then
                If cbolimit.Text <> "" And cbooperand.Text = "" Then
                    If Len(mwhere) = 0 Then
                        mwhere = " where " + cbofieldfilter + "=" + CStr(cbolimit) + ""
                    Else
                        mwhere = mwhere + " and  " + cbofieldfilter + " = " + CStr(cbolimit) + ""
                    End If
                Else
                    If cbolimit.Text <> "" Then
                        If cbooperand.Text <> "" Then
                            If cbooperand.Text = "between" Then
                                If Len(mwhere) = 0 Then
                                    mwhere = " where " + cbofieldfilter + " between " + CStr(cbolimit) + " and " + CStr(cbolimit1.Text) + ""
                                Else
                                    mwhere = mwhere + " and  " + cbofieldfilter + " between " + CStr(cbolimit) + " and " + CStr(cbolimit1.Text) + ""
                                End If
                            Else
                                If Len(mwhere) = 0 Then
                                    mwhere = " where " + cbofieldfilter + " " + cbooperand.Text + " " + CStr(cbolimit)
                                Else
                                    mwhere = mwhere + " and " + cbofieldfilter + " " + cbooperand.Text + " " + CStr(cbolimit)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            
            'If cbostatus.Text <> "" Then
            'If cbostatus.Text = "Recycle" Then
            'If Len(mwhere) = 0 Then
            'mwhere = " where bucket ='R'"
            'Else
            'mwhere = mwhere + " and  bucket ='R'"
            'End If
            'End If
            '
            'If cbostatus.Text = "New" Then
            'If Len(mwhere) = 0 Then
            'mwhere = " where (bucket is null or bucket='N')"
            'Else
            'mwhere = mwhere + " and  (bucket is null or bucket='N')"
            'End If
            'End If
            'End If
            'If cboarea.Text <> "" Then
            'If cboarea.Text = "Jakarta" Then
            'If Len(mwhere) = 0 Then
            '
            'mwhere = " where  ( city_1  ilike 'JAKARTA%' OR  CITY_1 ILIKE 'BODETABEK' )"
            'Else
            'mwhere = mwhere + " AND ( city_1  ilike 'JAKARTA%' OR  CITY_1 ILIKE 'BODETABEK' )"
            'End If
            '
            'ElseIf cboarea.Text = "Luar Jakarta" Then
            'If Len(mwhere) = 0 Then
            'mwhere = " where  city_1 not ilike 'JAKARTA%' AND CITY_1 NOT ILIKE 'BODETABEK'"
            'Else
            'mwhere = mwhere + " AND  city_1 not ilike 'JAKARTA%' AND CITY_1 NOT ILIKE 'BODETABEK'"
            'End If
            '
            'End If
            'End If
            '
            '
            'If cbopendate1.Text <> "" And cbopendate2.Text <> "" Then
            'If Len(mwhere) = 0 Then
            'mwhere = " WHERE date(tgl_trans) between '" + cbopendate1.Text + "'  and '" + cbopendate2.Text + "'"
            'Else
            'mwhere = mwhere + " AND date(tgl_trans) between '" + cbopendate1.Text + "'  and '" + cbopendate2.Text + "'"
            'End If
            'End If
             
             
                
            If cmbcampaigncode.Text <> "" Then
                If Len(mwhere) = 0 Then
                    mwhere = " where campaign_code ='" + getCampaign_code + "'"
                Else
                    mwhere = mwhere + " and  campaign_code ='" + getCampaign_code + "'"
                End If
            End If
            
            'If Combo2.Text <> "" Then
            'If Len(mwhere) = 0 Then
            'mwhere = " where saving_type ='" + Combo2.Text + "'"
            'Else
            'mwhere = mwhere + " and  saving_type ='" + Combo2.Text + "'"
            'End If
            'End If
            '
            'If Len(mwhere) = 0 Then
            'mwhere = " where sts_move = 0"
            'Else
            'mwhere = mwhere + " and  sts_move = 0"
            'End If
            '
            '
            'If Len(mwhere) = 0 Then
            'mwhere = "  where   status_open=0"
            'Else
            'mwhere = mwhere + " and  status_open=0"
            'End If
            '
                        
                        
            
            Set M_objrs = New ADODB.Recordset
                M_objrs.CursorLocation = adUseClient
                M_objrs.Open sStrsql + mwhere, M_OBJCONN, adOpenDynamic, adLockOptimistic
                If Not M_objrs.EOF Then
                    'txtjmlcampaign.Text = IIf(IsNull(M_objrs!JML), "0", M_objrs!JML)
                End If
            Set M_objrs = Nothing
   End Select


    For i = 1 To LVSpv.ListItems.Count
                If Val(LVSpv.ListItems(i).SubItems(5)) <> 0 Then
                    STRUPDATE = "insert into mgm_hst(tglcall,statuscall,reasoncall,agent,campaign_code,campaign_name,agentnama,id_cust) select now(),'New Data'::character varying(30),'New Data'::character varying(30),'" + LVSpv.ListItems(i).SubItems(1) + "',campaign_code,campaign_name,'" + LVSpv.ListItems(i).SubItems(2) + "',id_CUST from mgm where id_cust in( "
                    STRUPDATE = STRUPDATE + sStrsql + mwhere + " order by id_cust  limit " + LVSpv.ListItems(i).SubItems(5) + " ) and statuscall <> 'Agree'   "
                    M_OBJCONN.Execute (STRUPDATE)
                    
                    Call CreateInsert_Waterfall_Hst(sStrsql & mwhere & " order by id_cust  limit " & LVSpv.ListItems(i).SubItems(5), LVSpv.ListItems(i).SubItems(1), LVSpv.ListItems(i).SubItems(2))
                    
                    STRUPDATE = " update mgm set agent='" + LVSpv.ListItems(i).SubItems(1) + "', nmagent= '" + LVSpv.ListItems(i).SubItems(2) + "',flag_recycle= 0,tgldistribusi =date(now()) where id_cust in( "
                    STRUPDATE = STRUPDATE + sStrsql + mwhere + " order by id_cust  limit " + LVSpv.ListItems(i).SubItems(5) + " ) and statuscall <> 'Agree'   "
                    M_OBJCONN.Execute (STRUPDATE)
                    If cbostatus.Text <> "Recycle" Then
                        cmdsql_update = "insert into tbllogdistribusi(userid,nama,campaign_code,jmldata,sendby) values "
                        cmdsql_update = cmdsql_update + "('" + LVSpv.ListItems.Item(i).SubItems(1) + "','" + LVSpv.ListItems.Item(i).SubItems(2) + "','" + getCampaign_code + "'," + CStr(Val(LVSpv.ListItems(i).SubItems(5))) + ",'" + MDIForm1.TxtUsername.Text + "')"
                        M_OBJCONN.Execute (cmdsql_update)
                        LVSpv.ListItems(i).SubItems(4) = Val(LVSpv.ListItems(i).SubItems(4)) + LVSpv.ListItems(i).SubItems(5)
                        LVSpv.ListItems(i).SubItems(5) = ""
                    End If
                    
        End If
    Next i
     
   
    
    Set M_OBJRS_tglhistory = Nothing
    Set M_objrs = Nothing
    Set M_OBJRS_history = Nothing
    Set M_OBJRS_history3 = Nothing
        
     
    
    m_msgbox = MsgBox("Proses distribusi berhasil!", vbOKOnly + vbInformation, "Informasi")
    PB.Value = 0
    cmdProses.Enabled = True
    'txtSisaCampaign.Text = Val(txtjmlcampaign.Text) - Val(txtSudahDistribusi.Text)
      
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
strsql = strsql + " select CAMPAIGN_CODE ,COUNT(NO_CASE) AS JML_DATA from mgm "
strsql = strsql + "  GROUP by campaign_code) AS ALLDATA LEFT JOIN "
strsql = strsql + " ( "
strsql = strsql + " select CAMPAIGN_CODE ,COUNT(NO_CASE) AS AVAILABLE_SPACE from mgm  WHERE (AGENT IS NULL OR AGENT='' ) "
strsql = strsql + " GROUP by campaign_code) AS TBLSPACE  ON ALLDATA.CAMPAIGN_CODE=TBLSPACE.CAMPAIGN_CODE ) as tblsatu left join "

strsql = strsql + " (select CAMPAIGN_CODE ,COUNT(NO_CASE) AS ALREADY_ASSIGN from mgm  WHERE (AGENT IS NOT NULL and AGENT<>'')"
strsql = strsql + " GROUP by campaign_code) AS TBLASSIGN  ON TBLSatu.batch=TBLASSIGN.CAMPAIGN_CODE ORDER BY BATCH "



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
txtavailable.Text = TOTALSPACE
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
strsql = strsql + " SELECT tbluser.tbluser_groupspvcode as team, count(no_case) as jml,campaign_code FROM MGM ,tbluser where mgm.agent=tbluser.tbluser_userid AND  tbluser.tbluser_groupspvcode ='" + MDIForm1.TxtUsername.Text + "' group by   tbluser_groupspvcode  ,campaign_code"
strsql = strsql + " union all("
strsql = strsql + " select agent as team, count(no_case) as jml,campaign_code from mgm where  agent in (select tbluser.tbluser_userid from tbluser WHERE  tbluser_kdlevel='2') AND AGENT ='" + MDIForm1.TxtUsername.Text + "'  group by agent,campaign_code )"
strsql = strsql + " ) A GROUP BY TEAM,CAMPAIGN_CODE "

ListView2.ListItems.CLEAR
MOBJ.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
'ProgressBar1.Max = MOBJ.RecordCount
    
     NILALREADY = 0
     TOTALSPACE = 0
     txtalreadytl.Text = ""
     
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
      strsql = "select count(*) as jml from mgm where agent in (select tbluser_userid from tbluser where   tbluser_groupspvcode= '" + sTEAM + "' and  tbluser_kdlevel =1) and campaign_code ='" + scampaign + "'"
      MOBJ1.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
      ListItem.SubItems(5) = IIf(IsNull(MOBJ1!jml), "0", MOBJ1!jml)
      txtalreadytl.Text = Val(txtalreadytl.Text) + MOBJ1!jml
      NILALREADY = IIf(IsNull(MOBJ1!jml), "0", MOBJ1!jml)
      TOTALALREADY = TOTALALREADY + Val(NILALREADY)
      Set MOBJ1 = Nothing
      
      Set MOBJ1 = New ADODB.Recordset
      MOBJ1.CursorLocation = adUseClient
      strsql = "select count(*) as jml from mgm where agent in (select tbluser_userid from tbluser where tbluser_userid= '" + sTEAM + "' and  tbluser_kdlevel =3) and campaign_code ='" + scampaign + "'"
      MOBJ1.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
      ListItem.SubItems(4) = IIf(IsNull(MOBJ1!jml), "0", MOBJ1!jml)
      txtavalabletl.Text = Val(txtavalabletl.Text) + MOBJ1!jml
      NILSPACE = IIf(IsNull(MOBJ1!jml), "0", MOBJ1!jml)
      TOTALSPACE = TOTALSPACE + Val(NILSPACE)
      Set MOBJ1 = Nothing
      MOBJ.MoveNext
      
Wend

 Warna_Row_Listview Form_distribute, ListView2, &HFFFFC0, vbWhite
txtavalabletl.Text = TOTALSPACE
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
strsql = strsql + " ( SELECT CAMPAIGN_CODE,AGENT,NO_CASE FROM MGM WHERE AGENT IN (SELECT tbluser_userid FROM tbluser WHERE tbluser_userid IN (SELECT DISTINCT(tbluser_userid)"
strsql = strsql + " FROM tbluser WHERE  tbluser_kdlevel ='1' AND   tbluser_groupspvcode  ='" + MDIForm1.TxtUsername.Text + "'))) TBLMGM ,tbluser WHERE TBLMGM.AGENT=tbluser.tbluser_userid GROUP BY CAMPAIGN_CODE,TBLMGM.AGENT)  ORDER BY AGENT, CAMPAIGN_CODE"
 
ListView3.ListItems.CLEAR
MOBJ.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
Text2.Text = ""

While Not MOBJ.EOF
Set ListItem = ListView3.ListItems.ADD(, , MOBJ.Bookmark)
      ListItem.SubItems(1) = IIf(IsNull(MOBJ!AGENT), "", MOBJ!AGENT)
      ListItem.SubItems(2) = IIf(IsNull(MOBJ!campaign_code), "", MOBJ!campaign_code)
      ListItem.SubItems(3) = IIf(IsNull(MOBJ!jml), "", MOBJ!jml)
      Text2.Text = Val(Text2.Text) + Val(IIf(IsNull(MOBJ!jml), "", MOBJ!jml))
      MOBJ.MoveNext
Wend
Warna_Row_Listview Form_distribute, ListView3, &HFFFFC0, vbWhite
'txtavalabletl.Text = TOTALSPACE
'txtalreadytl.Text = TOTALALREADY

End Sub
Public Sub isicombo_opendate()
 Dim M_objrs As New ADODB.Recordset
     ListView5.ListItems.CLEAR
       intvrl = InStr(1, cmbcampaigncode.Text, "!", vbTextCompare)
               If intvrl <> 0 Then
                  ArrayString = Split(cmbcampaigncode.Text, "!", 2, vbTextCompare)
                  getCampaign_code = ArrayString(0)
                  getCampaign_name = ArrayString(1)
               End If
               getCampaign_code = cmbcampaigncode.Text 'HENDRI CODE
               
 Set M_objrs = New ADODB.Recordset
 M_objrs.CursorLocation = adUseClient
 strsql = "select DISTINCT(tgl_trans) from mgm where campaign_code='" + getCampaign_code + "' ORDER BY tgl_trans ASC"
 M_objrs.Open strsql, M_OBJCONN, adOpenDynamic, adLockOptimistic
 
 Tgl1 = cbopendate1.Text
 Tgl2 = cbopendate2.Text
 cbopendate1.CLEAR
 cbopendate2.CLEAR
 While Not M_objrs.EOF
        cbopendate1.AddItem Format(IIf(IsNull(M_objrs!tgl_trans), "", M_objrs!tgl_trans), "yyyy-mm-dd")
        cbopendate2.AddItem Format(IIf(IsNull(M_objrs!tgl_trans), "", M_objrs!tgl_trans), "yyyy-mm-dd")
        M_objrs.MoveNext
 Wend
 cbopendate1.Text = Tgl1
 cbopendate2.Text = Tgl2
End Sub
Public Sub loadbucketTL()
Dim M_objrs As ADODB.Recordset
Dim strsql As String
Dim ListItem As ListItem
   
If Combo1.Text = Empty Then
    m_msgbox = MsgBox("Textbox campaign code tidak boleh kosong!", vbOKOnly + vbExclamation, "Informasi")
    Exit Sub
End If
    
 intvrl = InStr(1, Combo1.Text, "!", vbTextCompare)
               If intvrl <> 0 Then
                  ArrayString = Split(Combo1.Text, "!", 2, vbTextCompare)
                  getCampaign_code = ArrayString(0)
                  getCampaign_name = ArrayString(1)
               End If
               getCampaign_code = cmbcampaigncode.Text 'HENDRI CODE
               
If Check1.Value = 1 Then '<-- kalo di cek berarti pakai like
    sUseSimiliar = " LIKE '%" + Combo1.Text + "%'"
Else
    sUseSimiliar = " = '" + Combo1.Text + "'"
End If
    
strsql = " SELECT * FROM ("
strsql = strsql + " SELECT U.tbluser_userid,U.tbluser_name,COALESCE(ACU.JML,0) AS JML FROM (SELECT tbluser_userid,tbluser_name FROM tbluser WHERE tbluser_kdlevel = 2 AND tbluser_kdstatus= 1) U"
strsql = strsql + " LEFT JOIN (SELECT AGENT,COUNT(AGENT) AS JML FROM MGM"
strsql = strsql + " WHERE CAMPAIGN_CODE " + sUseSimiliar + " AND AGENT IN (SELECT tbluser_userid FROM tbluser WHERE tbluser_kdlevel  = 2 AND tbluser_kdstatus = 1)"
strsql = strsql + " GROUP BY AGENT ORDER BY JML"
strsql = strsql + " ) AS ACU ON(ACU.AGENT=U.tbluser_userid)) AS GG ORDER BY JML DESC,tbluser_userid"
   
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
        list.SubItems(1) = IIf(IsNull(M_objrs!tbluser_userid), "", M_objrs!tbluser_userid)
        list.SubItems(2) = IIf(IsNull(M_objrs!tbluser_name), "", M_objrs!tbluser_name)
        list.SubItems(3) = IIf(IsNull(M_objrs!jml), 0, M_objrs!jml)
        If list.SubItems(1) <> "AM1" Then
            hjml = hjml + list.SubItems(3)
        End If
        M_objrs.MoveNext
    Wend
    Warna_Row_Listview Form_distribute, ListView7, &HFFFFC0, vbWhite
    Text1.Text = hjml '-> jumlah all
    Set M_objrs = Nothing
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 2 Then


        Combo1.Text = cmbcampaigncode.Text
        additemcombo1
    End If
End Sub
Public Sub additemcombo1()
    Dim M_objrs As ADODB.Recordset
    Dim cmdsql As String
    
    'Mengisi data ke combo campaigncode
    cmdsql = "select tbldatasource_campaign_code from tbldatasource order by   tbldatasource_tglentry DESC"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    
    M_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    
    Combo1.CLEAR
    'Combo1.AddItem "<Select More Campaign>"
    While Not M_objrs.EOF
        Combo1.AddItem M_objrs("tbldatasource_campaign_code")
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
                getCampaign_code = cmbcampaigncode.Text 'HENDRI CODE
    
If Check1.Value = 1 Then '<-- kalo di cek berarti pakai like
    sUseSimiliar = " LIKE '%" + sCamapign + "%'"
Else
    sUseSimiliar = " = '" + sCamapign + "'"
End If

strsql = " SELECT * FROM ("
strsql = strsql + " SELECT U.tbluser_userid,U.tbluser_name,COALESCE(ACU.JML,0) AS JML FROM (SELECT tbluser_userid,tbluser_name FROM tbluser WHERE tbluser_groupspvcode = '" + sTl + "' AND  tbluser_kdstatus ='1' AND tbluser_kdlevel='1') U"
strsql = strsql + " LEFT JOIN ("
strsql = strsql + " SELECT AGENT,COUNT(AGENT) AS JML FROM MGM"
strsql = strsql + " WHERE CAMPAIGN_CODE " + sUseSimiliar + " AND AGENT IN (SELECT tbluser_userid FROM tbluser WHERE tbluser_groupspvcode = '" + sTl + "' AND tbluser_kdstatus ='1'  AND tbluser_kdlevel='1' )"
strsql = strsql + " GROUP BY AGENT "
strsql = strsql + " ORDER BY JML"
strsql = strsql + " ) AS ACU ON(ACU.AGENT=U.tbluser_userid)"
strsql = strsql + " ) AS GG"
strsql = strsql + " ORDER BY tbluser_userid = '" + sTl + "' DESC,JML DESC"

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
        list.SubItems(1) = IIf(IsNull(M_objrs!tbluser_userid), "", M_objrs!tbluser_userid)
        list.SubItems(2) = IIf(IsNull(M_objrs!tbluser_name), "", M_objrs!tbluser_name)
        list.SubItems(3) = IIf(IsNull(M_objrs!jml), 0, M_objrs!jml)
        If no > 1 Then
            hjmls = hjmls + list.SubItems(3)
        End If
        M_objrs.MoveNext
    Wend
    Warna_Row_Listview Form_distribute, ListView8, &HFFFFC0, vbWhite
    Text3.Text = hjmls '-> jumlah all
    Set M_objrs = Nothing
End Sub
Private Sub Text4_Click()
Combo1.Text = Text4.Text
End Sub
Public Sub trackdistribusi()
    Dim M_objrs As ADODB.Recordset
    Dim cmdsql As String
    
    
    If UCase(MDIForm1.txtlevel.Text) <> "SUPERVISOR" Then
        Exit Sub
    End If
    
    sWhere = " WHERE SENDBY = '" + MDIForm1.TxtUsername.Text + "'"
    
    
    If TDBDate1.ValueIsNull = False And TDBDate2.ValueIsNull = False Then
        sWhere = sWhere + "AND DATE(TGL) BETWEEN '" + Format(TDBDate1.Value, "YYYY-MM-DD") + "' AND '" + Format(TDBDate2.Value, "YYYY-MM-DD") + "' "
    End If
    
    
    
    cmdsql = "SELECT COUNT(USERID) AS TIMES,SUM(JMLDATA) AS JUMLAH,USERID,NAMA,textcat_all(CAMPAIGN_CODE || ', ') AS CAMPAIGN FROM ("
    cmdsql = cmdsql + " SELECT * FROM tbllogdistribusi " + sWhere
    cmdsql = cmdsql + " ) GROUP BY USERID,NAMA"
    cmdsql = cmdsql + " ORDER BY JUMLAH DESC"
    
    Set M_objrs = New ADODB.Recordset
    M_objrs.CursorLocation = adUseClient
    
    M_objrs.Open cmdsql, M_OBJCONN, adOpenDynamic, adLockOptimistic, adCmdText
    ListView9.ListItems.CLEAR
    While Not M_objrs.EOF
    Set ListItem = ListView9.ListItems.ADD(, , M_objrs.Bookmark)
      ListItem.SubItems(1) = IIf(IsNull(M_objrs!TIMES), "", M_objrs!TIMES)
      ListItem.SubItems(2) = IIf(IsNull(M_objrs!JUMLAH), "", M_objrs!JUMLAH)
      ListItem.SubItems(3) = IIf(IsNull(M_objrs!USERID), "", M_objrs!USERID)
      ListItem.SubItems(4) = IIf(IsNull(M_objrs!nama), "", M_objrs!nama)
      ListItem.SubItems(5) = IIf(IsNull(M_objrs!Campaign), "", M_objrs!Campaign)
            M_objrs.MoveNext
    Wend
    Set M_objrs = Nothing
End Sub

Private Sub HEADER_TRACKDIS()
   
    ListView9.ColumnHeaders.ADD 1, , "No", 5 * TXT
    ListView9.ColumnHeaders.ADD 2, , "Times", 5 * TXT
    ListView9.ColumnHeaders.ADD 3, , "Jumlah", 5 * TXT
    ListView9.ColumnHeaders.ADD 4, , "Userid", 10 * TXT
    ListView9.ColumnHeaders.ADD 5, , "Nama", 15 * TXT
    ListView9.ColumnHeaders.ADD 6, , "Campaign", 25 * TXT
End Sub

